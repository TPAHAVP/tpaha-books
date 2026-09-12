import { test } from 'node:test';
import assert from 'node:assert/strict';
import { incidentKey, IncidentStore } from '../../site/js/save/incidents.js';
import { createTabGuard } from '../../site/js/save/tab-guard.js';
import { reportFormattingKey, ReportFormattingStore, ReportFormattingStorageError } from '../../site/js/save/report-formatting.js';

class MemStorage {
  constructor() { this.m = new Map(); this.failWrites = false; }
  get length() { return this.m.size; }
  key(i) { return [...this.m.keys()][i] ?? null; }
  getItem(k) { return this.m.has(k) ? this.m.get(k) : null; }
  setItem(k, v) { if (this.failWrites) throw new Error('quota'); this.m.set(k, String(v)); }
  removeItem(k) { this.m.delete(k); }
}

test('incident keys separate tenant, account and workbook; records round-trip; junk is dropped', () => {
  const a = incidentKey({ tenantId: 't1', homeAccountId: 'u1', workbookId: 'D!I' });
  const b = incidentKey({ tenantId: 't1', homeAccountId: 'u2', workbookId: 'D!I' });
  const c = incidentKey({ tenantId: 't1', homeAccountId: 'u1', workbookId: 'D!J' });
  assert.notEqual(a, b); assert.notEqual(a, c);
  assert.throws(() => incidentKey({ tenantId: '', homeAccountId: 'u', workbookId: 'w' }));
  const s = new MemStorage();
  const store = new IncidentStore(s, a);
  assert.equal(store.read(), null);
  const rec = { kind: 'wrong-row-deleted', message: 'x', details: { victim: { id: 8 } }, at: '2026-09-07T00:00:00Z', resolved: false };
  assert.equal(store.write(rec), true);
  assert.deepEqual(store.read(), rec);
  s.setItem(a, '{not json');
  assert.equal(store.read(), null);
  s.setItem(a, JSON.stringify({ nokind: true }));
  assert.equal(store.read(), null);
  s.failWrites = true;
  assert.equal(store.write(rec), false, 'a storage failure is reported, not thrown');
  store.clear();
});

// ---- tab guard with a fake broadcast channel -------------------------------------------------
function fakeChannels() {
  const members = [];
  return name => {
    const ch = { name, onmessage: null, closed: false, postMessage(data) { for (const m of members) if (m !== ch && !m.closed && m.onmessage) m.onmessage({ data }); }, close() { ch.closed = true; } };
    members.push(ch);
    return ch;
  };
}
const immediate = fn => setImmediate(fn);

test('the first tab is primary; a tab opened later is secondary; a tab without BroadcastChannel is primary', async () => {
  const factory = fakeChannels();
  const first = createTabGuard({ id: 'a', channelFactory: factory, timeoutMs: 0, setTimeoutImpl: immediate });
  const r1 = await first.start();
  assert.deepEqual(r1, { primary: true, others: 0 });
  const second = createTabGuard({ id: 'b', channelFactory: factory, timeoutMs: 0, setTimeoutImpl: immediate });
  const r2 = await second.start();
  assert.equal(r2.primary, false);
  assert.equal(r2.others, 1);
  assert.equal(first.primary, true, 'the first tab keeps writing');
  const none = createTabGuard({ id: 'c', channelFactory: () => null });
  const r3 = await none.start();
  assert.equal(r3.primary, true);
  assert.equal(r3.unsupported, true);
});

test('two tabs starting at the same moment: exactly one becomes primary (lower id wins)', async () => {
  const factory = fakeChannels();
  const a = createTabGuard({ id: 'aaa', channelFactory: factory, timeoutMs: 0, setTimeoutImpl: immediate });
  const b = createTabGuard({ id: 'bbb', channelFactory: factory, timeoutMs: 0, setTimeoutImpl: immediate });
  const [ra, rb] = await Promise.all([a.start(), b.start()]);
  assert.equal(ra.primary, true);
  assert.equal(rb.primary, false);
});

// ---- unfinished monthly report formatting ----------------------------------------------------
const fmtStore = () => new ReportFormattingStore(new MemStorage(), reportFormattingKey({ tenantId: 't', homeAccountId: 'u', workbookId: 'w' }));

test('W2b: reserving a month that is already outstanding adds nothing to release, so a conflict cannot erase the older reminder', () => {
  const s = fmtStore();
  s.add([9], 'Transaction saved. Excel report formatting still needs updating.', 'formatting');
  const { rec, added } = s.reserve([9], 'A change was interrupted…', 'interrupted');
  assert.deepEqual(added, [], 'September was already outstanding: this operation added nothing');
  assert.deepEqual(rec.months, [9]);
  assert.equal(rec.message, 'Transaction saved. Excel report formatting still needs updating.', 'the older, more precise reason is kept');
  assert.equal(rec.blockedBy, 'formatting');
  s.remove(added);                                      // the operation conflicted and released what it added
  assert.deepEqual(s.read().months, [9], 'the older reminder survives');
  assert.equal(s.read().message, 'Transaction saved. Excel report formatting still needs updating.');
});

test('W2b: a reservation releases only the months it introduced', () => {
  const s = fmtStore();
  s.add([9], 'older', 'formatting');
  const { added } = s.reserve([9, 10], 'newer', 'interrupted');
  assert.deepEqual(added, [10]);
  assert.deepEqual(s.read().months, [9, 10]);
  s.remove(added);
  assert.deepEqual(s.read().months, [9], 'only the month this operation brought in is dropped');
  assert.equal(s.read().message, 'older');
});

test('W2b: with nothing outstanding, a reservation owns what it reserves and uses its own wording', () => {
  const s = fmtStore();
  const { rec, added } = s.reserve([3], 'interrupted wording', 'interrupted');
  assert.deepEqual(added, [3]);
  assert.equal(rec.message, 'interrupted wording');
  assert.equal(rec.blockedBy, 'interrupted');
  s.remove(added);
  assert.equal(s.read(), null, 'nothing left');
});

test('W2b: only verified completion clears an outstanding month', () => {
  const s = fmtStore();
  s.add([9], 'pending', 'formatting');
  s.reserve([9], 'x', 'interrupted');
  assert.deepEqual(s.read().months, [9]);
  assert.equal(s.remove([9]), null, 'a completed, verified format clears it');
  assert.equal(s.read(), null);
});

test('a storage refusal is reported, and still tells the caller what it would have added', () => {
  const mem = new MemStorage();
  const s = new ReportFormattingStore(mem, reportFormattingKey({ tenantId: 't', homeAccountId: 'u', workbookId: 'w' }));
  s.add([9], 'pending', 'formatting');
  mem.failWrites = true;
  try { s.reserve([9, 11], 'x', 'interrupted'); assert.fail('should have thrown'); }
  catch (e) {
    assert.ok(e instanceof ReportFormattingStorageError);
    assert.deepEqual(e.added, [11]);
    assert.deepEqual(e.fallback.months, [9, 11]);
    assert.equal(e.fallback.message, 'pending', 'the existing reason is still the one to show');
  }
});
