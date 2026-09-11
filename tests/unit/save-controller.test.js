import { test, mock } from 'node:test';
import assert from 'node:assert/strict';
import { SaveController } from '../../site/js/save/save-controller.js';
import { ConflictError, MarkerCollisionError } from '../../site/js/workbook/ledger-workbook.js';
import { ExcelApiError } from '../../site/js/workbook/excel-client.js';

const deferred = () => { let resolve, reject; const p = new Promise((res, rej) => { resolve = res; reject = rej; }); return { p, resolve, reject }; };

/**
 * perform / inspect / resume are recorded. `store` simulates the workbook: rows keyed by marker, each
 * { draft, complete }. inspect is READ-ONLY (never touches the store); resume finishes an incomplete row.
 */
function harness({ performImpl, inspectImpl, resumeImpl } = {}) {
  const calls = { perform: [], inspect: [], resume: [] };
  const states = [];
  const store = new Map();
  let markers = 0;
  const ctl = new SaveController({
    perform: async (draft, ctx) => {
      calls.perform.push({ draft, marker: ctx.marker });
      if (performImpl) return performImpl(draft, ctx, calls.perform.length, store);
      store.set(ctx.marker, { draft, complete: true });
      return { id: 100 + store.size, draft };
    },
    inspect: async (payload, ctx) => {
      calls.inspect.push({ payload, marker: ctx.marker });
      if (inspectImpl) return inspectImpl(payload, ctx, store);
      const row = store.get(ctx.marker);
      if (!row) return { outcome: 'missing' };
      if (JSON.stringify(row.draft) !== JSON.stringify(payload)) return { outcome: 'conflict', error: new MarkerCollisionError(ctx.marker, row) };
      if (!row.complete) return { outcome: 'incomplete', message: 'landed but not finished' };
      return { outcome: 'landed', result: { id: 100 + [...store.keys()].indexOf(ctx.marker) + 1, draft: row.draft, alreadyExisted: true } };
    },
    resume: async (payload, ctx) => {
      calls.resume.push({ payload, marker: ctx.marker });
      if (resumeImpl) return resumeImpl(payload, ctx, store);
      const row = store.get(ctx.marker);
      if (!row) return null;
      row.complete = true;
      return { id: 100 + [...store.keys()].indexOf(ctx.marker) + 1, draft: row.draft, alreadyExisted: true };
    },
    onChange: c => states.push(c.state),
    makeMarker: () => `m${++markers}`,
  });
  return { ctl, calls, states, store };
}
const lost = () => new ExcelApiError({ message: 'lost', status: 0, ambiguous: true });

test('a normal save goes unsaved -> saving -> saved, passes draft and marker, and clears the pending operation', async () => {
  const { ctl, calls, states } = harness();
  assert.equal(ctl.state, 'idle');
  ctl.edit({ description: 'a' });
  assert.equal(ctl.state, 'unsaved');
  const r = await ctl.submit();
  assert.equal(r.ok, true);
  assert.equal(ctl.state, 'saved');
  assert.deepEqual(states, ['unsaved', 'saving', 'saved']);
  assert.equal(calls.perform.length, 1);
  assert.equal(calls.inspect.length + calls.resume.length, 0, 'nothing to check on a clean first attempt');
  assert.equal(calls.perform[0].marker, 'm1');
  assert.equal(ctl.pending, null);
});

test('a second submit while saving is ignored: perform runs once', async () => {
  const gate = deferred();
  const { ctl, calls } = harness({ performImpl: async () => gate.p });
  ctl.edit({ description: 'a' });
  const p1 = ctl.submit();
  const r2 = await ctl.submit();
  assert.equal(r2.ignored, true);
  gate.resolve({ id: 1 });
  await p1;
  assert.equal(calls.perform.length, 1);
  assert.equal(ctl.state, 'saved');
});

test('an edit made while a save is uploading is kept and the state ends as unsaved; the next save is a new operation', async () => {
  const gate = deferred();
  const { ctl, calls } = harness({ performImpl: async (d, ctx, n, store) => { const v = await gate.p; store.set(ctx.marker, { draft: d, complete: true }); return v; } });
  ctl.edit({ description: 'first' });
  const p = ctl.submit();
  ctl.edit({ description: 'second' });
  assert.equal(ctl.hasUnsavedEdits, true);
  gate.resolve({ id: 1 });
  const r = await p;
  assert.equal(r.ok, true);
  assert.equal(ctl.state, 'unsaved');
  assert.deepEqual(ctl.draft, { description: 'second' });
  assert.equal(ctl.pending, null);
  await ctl.submit();
  assert.equal(calls.perform.length, 2);
  assert.notEqual(calls.perform[1].marker, calls.perform[0].marker);
  assert.equal(ctl.state, 'saved');
});

// ---------------------------------------------------------------------------------------------- R3 / S3
test('R3: after a lost response, pressing Save again (same content) resumes under the same marker and does not write twice', async () => {
  let n = 0;
  const { ctl, calls, store } = harness({ performImpl: async (d, ctx, k, st) => { n++; st.set(ctx.marker, { draft: d, complete: n > 1 }); if (n === 1) throw lost(); return { id: 200 }; } });
  ctl.edit({ description: 'a' });
  const r1 = await ctl.submit();
  assert.equal(r1.ok, false);
  assert.equal(ctl.state, 'failed');
  assert.equal(ctl.needsVerification, true);
  assert.equal(ctl.pending.marker, 'm1');
  ctl.edit({ description: 'a' });
  const r2 = await ctl.submit();
  assert.equal(r2.ok, true);
  assert.equal(r2.resolvedPending, true);
  assert.equal(ctl.state, 'saved');
  assert.equal(calls.perform.length, 1, 'no second write of the entry');
  assert.deepEqual(calls.resume.map(c => c.marker), ['m1']);
  assert.equal(store.size, 1);
});

test('R3: after a lost response that did NOT land, Save writes the current content under the SAME marker', async () => {
  let n = 0;
  const { ctl, calls } = harness({ performImpl: async (d, ctx, k, st) => { n++; if (n === 1) throw lost(); st.set(ctx.marker, { draft: d, complete: true }); return { id: 201 }; } });
  ctl.edit({ description: 'a', amount: 1 });
  await ctl.submit();
  ctl.edit({ description: 'a', amount: 2 });
  const r = await ctl.submit();
  assert.equal(r.ok, true);
  assert.equal(calls.resume.length, 1);
  assert.equal(calls.perform.length, 2);
  assert.equal(calls.perform[1].marker, 'm1');
  assert.deepEqual(calls.perform[1].draft, { description: 'a', amount: 2 });
  assert.equal(ctl.state, 'saved');
});

test('S3: check() is read-only: it never calls perform or resume, whatever the outcome', async () => {
  let n = 0;
  const { ctl, calls, store } = harness({ performImpl: async (d, ctx, k, st) => { n++; st.set(ctx.marker, { draft: d, complete: false }); if (n === 1) throw lost(); return { id: 1 }; } });
  ctl.edit({ description: 'a' });
  await ctl.submit();
  const c1 = await ctl.check();                       // landed but incomplete
  assert.equal(c1.outcome, 'incomplete');
  assert.equal(ctl.state, 'failed');
  assert.equal(ctl.needsVerification, false, 'certain: it is there');
  assert.ok(ctl.incomplete, 'the controller exposes the incomplete detail');
  assert.ok(ctl.pending, 'the operation stays pending until a Save finishes it');
  store.get('m1').complete = true;                    // finished elsewhere
  const c2 = await ctl.check();
  assert.equal(c2.outcome, 'landed');
  assert.equal(ctl.state, 'saved');
  store.clear();
  ctl.edit({ description: 'b' });
  await assert.rejects(async () => { throw lost(); });
  ctl.pending = { marker: 'm9', payload: { description: 'b' } };
  const c3 = await ctl.check();
  assert.equal(c3.outcome, 'missing');
  assert.equal(ctl.state, 'failed');
  assert.equal(calls.perform.length, 1);
  assert.equal(calls.resume.length, 0, 'check never resumes');
});

test('S3: Save after an incomplete check finishes the operation through resume, once', async () => {
  let n = 0;
  const { ctl, calls, store } = harness({ performImpl: async (d, ctx, k, st) => { n++; st.set(ctx.marker, { draft: d, complete: false }); throw lost(); } });
  ctl.edit({ description: 'a' });
  await ctl.submit();
  await ctl.check();
  assert.equal(ctl.state, 'failed');
  const r = await ctl.submit();
  assert.equal(r.ok, true);
  assert.equal(r.resolvedPending, true);
  assert.equal(ctl.state, 'saved');
  assert.equal(calls.resume.length, 1);
  assert.equal(calls.perform.length, 1);
  assert.equal(store.get('m1').complete, true);
});

// ---------------------------------------------------------------------------------------------- S2
test('S2: a restored pending operation and a NEWER draft: check finds the old payload landed -> the new draft stays unsaved and visible', async () => {
  const { ctl, calls, store } = harness();
  store.set('op-old', { draft: { description: 'Earlier saved version' }, complete: true });
  ctl.edit({ description: 'Later unsaved version' });
  ctl.restorePending({ marker: 'op-old', payload: { description: 'Earlier saved version' } });
  assert.equal(ctl.state, 'failed');
  assert.deepEqual(ctl.draft, { description: 'Later unsaved version' }, 'restorePending does not replace the draft');
  const c = await ctl.check();
  assert.equal(c.outcome, 'landed');
  assert.equal(c.currentDraftSaved, false);
  assert.equal(ctl.state, 'unsaved', 'the newer text is NOT reported as saved');
  assert.equal(ctl.hasUnsavedEdits, true);
  assert.deepEqual(ctl.draft, { description: 'Later unsaved version' });
  assert.equal(ctl.pending, null);
  assert.equal(calls.perform.length, 0);
  const r = await ctl.submit();                       // the newer draft becomes its own operation
  assert.equal(r.ok, true);
  assert.equal(calls.perform[0].marker, 'm1');
  assert.deepEqual(calls.perform[0].draft, { description: 'Later unsaved version' });
  assert.equal(store.size, 2);
});

test('S2: the same payload restored after a reload is reported saved once it is found, with no write', async () => {
  const { ctl, calls, store } = harness();
  store.set('op-old', { draft: { description: 'from before the reload' }, complete: true });
  ctl.edit({ description: 'from before the reload' });
  ctl.restorePending({ marker: 'op-old', payload: { description: 'from before the reload' } });
  assert.equal(ctl.hasUnsavedEdits, true, 'an unresolved attempt counts as unsaved work');
  const c = await ctl.check();
  assert.equal(c.outcome, 'landed');
  assert.equal(ctl.state, 'saved');
  assert.equal(calls.perform.length + calls.resume.length, 0);
});

test('S2: a landed operation with edits typed during the save keeps the newer edits unsaved even if the generation happens to match', async () => {
  let n = 0;
  const { ctl } = harness({ performImpl: async (d, ctx, k, st) => { n++; st.set(ctx.marker, { draft: d, complete: true }); if (n === 1) throw lost(); return { id: 202 }; } });
  ctl.edit({ description: 'a' });
  await ctl.submit();
  ctl.edit({ description: 'b' });
  const r = await ctl.submit();                       // resume finds 'a' landed
  assert.equal(r.ok, true);
  assert.equal(r.resolvedPending, true);
  assert.deepEqual(r.result.draft, { description: 'a' });
  assert.equal(ctl.state, 'unsaved', 'content "b" still needs its own save');
  assert.equal(ctl.pending, null);
});

// ---------------------------------------------------------------------------------------------- other outcomes
test('a definite failure keeps the pending operation; the next Save resumes (nothing there) and reuses the marker', async () => {
  let n = 0;
  const { ctl, calls } = harness({ performImpl: async (d, ctx, k, st) => { n++; if (n === 1) throw new ExcelApiError({ message: 'busy', status: 429, retryAfter: 2 }); st.set(ctx.marker, { draft: d, complete: true }); return { id: 1 }; } });
  ctl.edit({ a: 1 });
  await ctl.submit();
  assert.equal(ctl.state, 'failed');
  assert.equal(ctl.needsVerification, false);
  await ctl.retry();
  assert.equal(calls.resume.length, 1);
  assert.equal(calls.perform[1].marker, 'm1');
  assert.equal(ctl.state, 'saved');
});

test('a conflict becomes the conflict state with the details attached and no pending operation', async () => {
  const { ctl } = harness({ performImpl: async () => { throw new ConflictError('changed', { expected: { a: 1 }, current: { a: 2 } }); } });
  ctl.edit({ a: 1 });
  const r = await ctl.submit();
  assert.equal(r.ok, false);
  assert.equal(ctl.state, 'conflict');
  assert.equal(ctl.conflict.reason, 'changed');
  assert.equal(ctl.pending, null);
});

test('a marker collision (on resume or on check) is a conflict that drops the pending operation so the next Save gets a new marker', async () => {
  const { ctl, calls, store } = harness({ performImpl: async () => { throw lost(); } });
  ctl.edit({ description: 'mine' });
  await ctl.submit();
  store.set('m1', { draft: { description: 'someone else' }, complete: true });
  const c = await ctl.check();
  assert.equal(c.outcome, 'conflict');
  assert.equal(ctl.state, 'conflict');
  assert.equal(ctl.pending, null);
  assert.equal(calls.perform.length, 1);
  const { ctl: ctl2, calls: calls2, store: store2 } = harness({ performImpl: async (d, ctx, k, st) => { if (k === 1) throw lost(); st.set(ctx.marker, { draft: d, complete: true }); return { id: 1 }; }, resumeImpl: async (p, ctx) => { throw new MarkerCollisionError(ctx.marker, {}); } });
  ctl2.edit({ description: 'mine' });
  await ctl2.submit();
  const r = await ctl2.submit();
  assert.equal(r.ok, false);
  assert.equal(ctl2.state, 'conflict');
  assert.equal(ctl2.pending, null);
  const r2 = await ctl2.submit();
  assert.equal(r2.ok, true);
  assert.equal(calls2.perform[1].marker, 'm2', 'a fresh marker');
  assert.equal(store2.size, 1);
});

test('reset returns to idle and forgets the pending operation only when told to', async () => {
  const { ctl } = harness({ performImpl: async () => { throw lost(); } });
  ctl.edit({ a: 1 });
  await ctl.submit();
  assert.ok(ctl.pending);
  ctl.reset();
  assert.equal(ctl.state, 'failed', 'a pending operation is not forgotten by an ordinary reset');
  assert.ok(ctl.pending);
  ctl.reset({ discardPending: true });
  assert.equal(ctl.state, 'idle');
  assert.equal(ctl.pending, null);
  assert.equal(ctl.draft, null);
});

// ---------------------------------------------------------------------------------------------- T3: operation ids
test('T3: the default operation id is ISO 8601 with nine fraction digits and differs even when the clock does not move', () => {
  mock.timers.enable({ apis: ['Date'], now: new Date('2026-09-07T12:00:00.000Z') });
  try {
    const ctl = new SaveController({ perform: async () => ({}) });
    const a = ctl.makeMarker(), b = ctl.makeMarker();
    assert.match(a, /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{9}Z$/);
    assert.ok(a.startsWith('2026-09-07T12:00:00.000'), 'the clock part is the frozen time');
    assert.notEqual(a, b, 'two operations in the same millisecond get different ids');
    assert.ok(!Number.isNaN(new Date(a).getTime()));
  } finally { mock.timers.reset(); }
});

test('T3: one pending attempt keeps its marker across retries, and each distinct new operation gets a fresh one', async () => {
  const markers = [];
  let fail = true;
  const ctl = new SaveController({ perform: async (d, { marker }) => { markers.push(marker); if (fail) { fail = false; throw new ExcelApiError({ message: 'lost', status: 0, ambiguous: true }); } return { id: 1 }; } });
  ctl.edit({ a: 1 });
  await ctl.submit();
  await ctl.submit();
  assert.equal(markers.length, 2);
  assert.equal(markers[0], markers[1], 'same attempt, same id');
  ctl.edit({ a: 1 });
  await ctl.submit();
  assert.equal(markers.length, 3);
  assert.notEqual(markers[2], markers[0], 'a new operation with identical content has its own id');
});
