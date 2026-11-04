// Session errors from Microsoft Graph, after the live refusal of 2026-11-03 ("The target session is invalid.").
// Microsoft's paused-session envelope nests two innerError levels; the client must classify by the documented
// second-level code, recover automatically only where Microsoft says it may (invalidSessionReCreatable), and
// stop reusing a session, and sending anything, after a session error Microsoft does not call recreatable.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { MockWorkbook, PAUSED_SESSION } from '../../site/js/workbook/mock-excel.js';
import { ExcelClient, ExcelApiError, classifyError } from '../../site/js/workbook/excel-client.js';
import { LedgerWorkbook } from '../../site/js/workbook/ledger-workbook.js';
import { SaveController } from '../../site/js/save/save-controller.js';
import { SAMPLE_WORKBOOK as fixture } from '../../site/js/workbook/sample-workbook.js';

const makeClient = (mock, extra = {}) => new ExcelClient({ getToken: async () => 'tok', driveId: 'D1', itemId: 'I1', fetchImpl: (u, o) => mock.fetch(u, o), sleep: async () => {}, ...extra });
const nonBlank = rows => rows.filter(r => r.some(v => v !== '' && v !== null));
const sessionsOpened = mock => mock.log.filter(e => /createSession/.test(e.url)).length;
const appends = mock => mock.log.filter(e => e.method === 'POST' && /LOG_Table\/rows\/add/.test(e.url)).length;
const writes = (mock, from = 0) => mock.log.slice(from).filter(e => e.method !== 'GET' && !/Session/.test(e.url));
const PAUSED_BODY = {
  error: {
    code: 'InvalidSession', message: 'The target session is invalid.',
    innerError: {
      code: 'invalidSessionReCreatable', message: 'The session specified in the request does not exist or is invalid due to a transient error.',
      innerError: { code: 'InvalidOrTimedOutSession', message: "We noticed that you haven't been interacting with this workbook, so we paused your session." },
    },
  },
};

// ---------------------------------------------------------------------------------------------- classification
test('classifyError: the paused-session envelope (three levels) is a recreatable session error, classified by the second-level code', () => {
  const c = classifyError(400, PAUSED_BODY, new Headers());
  assert.equal(c.sessionInvalid, true);
  assert.equal(c.recreatable, true);
  assert.equal(c.retryableRead, true);
  assert.equal(c.innerCode, 'invalidSessionReCreatable', 'the second-level code, as Microsoft documents');
  assert.deepEqual(c.innerCodes, ['invalidSessionReCreatable', 'InvalidOrTimedOutSession'], 'the whole chain, for the log');
});

test('classifyError: a top-level InvalidSession with an inner code Microsoft does not call recreatable is a session error that must not be recreated', () => {
  const c = classifyError(400, { error: { code: 'InvalidSession', message: 'The target session is invalid.', innerError: { code: 'invalidSessionNotFound' } } }, new Headers());
  assert.equal(c.sessionInvalid, true);
  assert.equal(c.recreatable, false);
  assert.equal(c.retryableRead, false);
});

// ---------------------------------------------------------------------------------------------- client: recreatable
test('a paused session on a read is recreated and the read retried once (Microsoft: the client may recreate and resume)', async () => {
  const mock = new MockWorkbook(fixture);
  const client = makeClient(mock);
  await client.openSession();
  const old = client.sessionId;
  mock.pauseSession();
  const range = await client.getRange('ENTRY', 'B23');
  assert.deepEqual(range.values, [[5000]]);
  assert.notEqual(client.sessionId, old);
  assert.equal(sessionsOpened(mock), 2);
  assert.equal(client.blocked, null);
});

test('a paused session on a write: one attempt, no replay, the dead session is dropped, and the next read opens a new session', async () => {
  const mock = new MockWorkbook(fixture);
  const client = makeClient(mock);
  await client.openSession();
  mock.pauseSession();
  const before = nonBlank(mock.table('LOG_Table').rows).length;
  await assert.rejects(() => client.addTableRows('LOG_Table', [[30, 'm', 46271, 'Deposit', 'Membership', 1, 'x', '', '']]),
    e => e instanceof ExcelApiError && e.sessionInvalid === true && e.ambiguous === false && e.status === 400);
  assert.equal(appends(mock), 1, 'exactly one attempt');
  assert.equal(nonBlank(mock.table('LOG_Table').rows).length, before, 'nothing was written');
  assert.equal(client.sessionId, null, 'the paused session is not reused');
  assert.equal(client.blocked, null, 'a recreatable error does not block the client');
  const range = await client.getRange('ENTRY', 'B23');
  assert.deepEqual(range.values, [[5000]]);
  assert.equal(sessionsOpened(mock), 2, 'the read opened a new session; the write was not replayed');
  assert.equal(appends(mock), 1);
});

// ---------------------------------------------------------------------------------------------- client: not recreatable
test('a session error Microsoft does not call recreatable blocks the client: nothing is sent, in any form, until the block is cleared deliberately', async () => {
  const mock = new MockWorkbook(fixture);
  const log = [];
  const client = makeClient(mock, { log: e => log.push(e) });
  await client.openSession();
  mock.breakSession('invalidSessionNotFound');
  await assert.rejects(() => client.getRange('ENTRY', 'B23'), e => e instanceof ExcelApiError && e.sessionInvalid === true && e.blocked === true && e.innerCode === 'invalidSessionNotFound');
  assert.ok(client.blocked && /invalidSessionNotFound/.test(client.blocked.innerCode), 'the client records why it is blocked');
  assert.equal(client.sessionId, null);
  const sent = mock.log.length;
  const refused = e => e instanceof ExcelApiError && e.blocked === true && /invalidSessionNotFound/.test(e.message);
  await assert.rejects(() => client.getRange('ENTRY', 'B23'), refused);
  await assert.rejects(() => client.addTableRows('LOG_Table', [[30, 'm', 46271, 'Deposit', 'Membership', 1, 'x', '', '']]), refused);
  await assert.rejects(() => client.patchRange('ENTRY', 'B23', { values: [[1]] }), refused);
  await assert.rejects(() => client.batch([{ method: 'GET', path: '/workbook/worksheets/ENTRY/range(address=\'B23\')' }]), refused);
  await assert.rejects(() => client.openSession(), refused);
  await assert.rejects(() => client.calculate(), refused);
  assert.equal(mock.log.length, sent, 'no request of any kind reached the service while blocked');
  assert.equal(sessionsOpened(mock), 1, 'no session was recreated');
  client.clearBlock();
  assert.equal(client.blocked, null);
  const range = await client.getRange('ENTRY', 'B23');
  assert.deepEqual(range.values, [[5000]]);
  assert.equal(sessionsOpened(mock), 2, 'after the deliberate clear a new session is opened');
});

test('a non-recreatable session error inside a batch blocks the client the same way', async () => {
  const mock = new MockWorkbook(fixture);
  const client = makeClient(mock);
  await client.openSession();
  mock.breakSession('invalidSessionAuthentication');
  const results = await client.batch([{ method: 'GET', path: '/workbook/worksheets/ENTRY/range(address=\'B23\')' }]);
  assert.ok(results[0].error && results[0].error.sessionInvalid);
  assert.ok(client.blocked && client.blocked.innerCode === 'invalidSessionAuthentication');
  const sent = mock.log.length;
  await assert.rejects(() => client.getRange('ENTRY', 'B23'), e => e.blocked === true);
  assert.equal(mock.log.length, sent);
});

// ---------------------------------------------------------------------------------------------- client: the log
test('the request log never carries a session id, and a refusal records the whole inner-code chain', async () => {
  const mock = new MockWorkbook(fixture);
  const log = [];
  const client = makeClient(mock, { log: e => log.push(e) });
  await client.openSession();
  const opened = log.find(e => e.event === 'session');
  assert.ok(opened, 'opening a session is logged');
  assert.equal(opened.id, undefined, 'without the session id');
  assert.ok(!JSON.stringify(log).includes(client.sessionId), 'the session id appears nowhere in the log');
  mock.pauseSession();
  await assert.rejects(() => client.addTableRows('LOG_Table', [[30, 'm', 46271, 'Deposit', 'Membership', 1, 'x', '', '']]));
  const refusal = log.find(e => e.status === 400);
  assert.equal(refusal.code, 'InvalidSession');
  assert.equal(refusal.innerCode, 'invalidSessionReCreatable');
  assert.deepEqual(refusal.innerCodes, ['invalidSessionReCreatable', 'InvalidOrTimedOutSession']);
  assert.ok(!JSON.stringify(log).includes('"tok"'), 'no token in the log');
});

// ---------------------------------------------------------------------------------------------- the correction, stage by stage
const setup = () => { const mock = new MockWorkbook(fixture); const client = makeClient(mock); return { mock, client, wb: new LedgerWorkbook(client) }; };
const controllerFor = (wb, ref) => new SaveController({ perform: d => wb.updateTransaction(ref, d), inspect: d => wb.inspectCorrection(ref, d), resume: d => wb.completeCorrection(ref, d) });

test('a session paused before a correction is recovered on the correction\'s initial read; the correction then succeeds, written once', async () => {
  const { mock, wb } = setup();
  const ref = (await wb.load()).transactions.find(t => t.id === 2);
  mock.pauseSession();
  const done = await wb.updateTransaction(ref, { amount: 90 });
  assert.equal(done.amount, 90);
  const rows = nonBlank(mock.table('LOG_Table').rows).filter(r => r[0] === 2);
  assert.equal(rows.length, 1, 'one copy');
  assert.equal(rows[0][5], 90);
  assert.equal(appends(mock), 1, 'one append');
  assert.equal(sessionsOpened(mock), 2, 'the read recreated the session once');
  assert.equal(wb.halted, null);
});

test('a paused session refusing the correction append itself: nothing is written and nothing replayed; Check reads on a new session and finds nothing; Save then writes once', async () => {
  const { mock, wb } = setup();
  const ref = (await wb.load()).transactions.find(t => t.id === 2);
  const ctl = controllerFor(wb, ref);
  mock.failNext({ match: /LOG_Table\/rows\/add/, ...PAUSED_SESSION });
  ctl.edit({ amount: '90' });
  const r1 = await ctl.submit();
  assert.equal(r1.ok, false);
  assert.ok(r1.error instanceof ExcelApiError && r1.error.sessionInvalid && !r1.error.ambiguous, String(r1.error));
  assert.equal(ctl.state, 'failed');
  assert.ok(ctl.pending, 'the attempt stays pending');
  assert.equal(ctl.needsVerification, false, 'a refusal is definite, not a lost answer');
  assert.equal(appends(mock), 1, 'the refused append was sent once and never replayed');
  assert.equal(nonBlank(mock.table('LOG_Table').rows).filter(r => r[0] === 2).length, 1);
  assert.equal(nonBlank(mock.table('LOG_Table').rows).find(r => r[0] === 2)[5], ref.amount, 'unchanged');
  const start = mock.log.length;
  const c = await ctl.check();
  assert.equal(c.outcome, 'missing');
  assert.equal(writes(mock, start).length, 0, 'the check wrote nothing');
  const r2 = await ctl.submit();
  assert.equal(r2.ok, true);
  assert.equal(ctl.state, 'saved');
  const rows = nonBlank(mock.table('LOG_Table').rows).filter(r => r[0] === 2);
  assert.equal(rows.length, 1);
  assert.equal(rows[0][5], 90);
  assert.equal(appends(mock), 2, 'the one refused attempt plus the one that saved');
});

test('a blocking session error after the corrected copy was appended: the attempt stays pending with its evidence, the blocked client neither checks nor writes, and a fresh client finishes it without a duplicate', async () => {
  const { mock, wb, client } = setup();
  const ref = (await wb.load()).transactions.find(t => t.id === 5);
  const ctl = controllerFor(wb, ref);
  let appended = false;
  mock.beforeRespond = async e => {
    if (e.method === 'POST' && /LOG_Table\/rows\/add/.test(e.url)) appended = true;
    else if (appended && /LOG_Table\/dataBodyRange/.test(e.url)) { mock.beforeRespond = null; mock.breakSession('invalidSessionUnexpected'); }
  };
  ctl.edit({ amount: '380' });
  const r1 = await ctl.submit();
  assert.equal(r1.ok, false);
  assert.ok(r1.error instanceof ExcelApiError && r1.error.blocked, String(r1.error));
  assert.ok(client.blocked);
  assert.equal(nonBlank(mock.table('LOG_Table').rows).filter(r => r[0] === 5).length, 2, 'the corrected copy landed; the old copy remains');
  assert.ok(ctl.pending && ctl.pending.payload.amount === '380', 'the pending operation keeps its payload and marker');
  const marker = ctl.pending.marker;
  const sent = mock.log.length;
  const c = await ctl.check();
  assert.ok(c.error && c.error.blocked, 'the blocked client cannot even read');
  assert.ok(ctl.checkFailed, 'the controller says the check failed');
  assert.ok(ctl.pending && ctl.pending.marker === marker, 'the pending operation survives a failed check');
  assert.equal(mock.log.length, sent, 'nothing was sent while blocked');
  // A reload constructs a new client. The pending operation is restored and finished without a second append.
  const fresh = new LedgerWorkbook(makeClient(mock));
  const ctl2 = controllerFor(fresh, ref);
  ctl2.restorePending(ctl.pending);
  const c2 = await ctl2.check();
  assert.equal(c2.outcome, 'incomplete');
  assert.ok(ctl2.incomplete.missingSteps.includes('removal of the old copy'));
  const r2 = await ctl2.submit();
  assert.equal(r2.ok, true);
  assert.equal(r2.resolvedPending, true);
  const rows = nonBlank(mock.table('LOG_Table').rows).filter(r => r[0] === 5);
  assert.equal(rows.length, 1, 'no duplicate');
  assert.equal(rows[0][5], 380);
  assert.equal(appends(mock), 1, 'exactly one append in the whole sequence');
  assert.equal(fresh.halted, null);
});

// ---------------------------------------------------------------------------------------------- the controller's wording
test('a check that cannot read keeps the pending attempt and says the check failed, not that the answer was lost; a later check that answers clears it', async () => {
  const store = new Map();
  let performs = 0, inspects = 0;
  const ctl = new SaveController({
    perform: async (draft, ctx) => { performs += 1; if (performs === 1) throw new ExcelApiError({ message: 'The target session is invalid.', status: 400, sessionInvalid: true }); store.set(ctx.marker, draft); return { id: 1 }; },
    inspect: async (payload, ctx) => { inspects += 1; if (inspects === 1) throw new ExcelApiError({ message: 'The target session is invalid.', status: 400, sessionInvalid: true }); return store.has(ctx.marker) ? { outcome: 'landed', result: { id: 1 } } : { outcome: 'missing' }; },
    makeMarker: () => 'm1',
  });
  ctl.edit({ amount: '2' });
  const r1 = await ctl.submit();
  assert.equal(r1.ok, false);
  assert.equal(ctl.needsVerification, false);
  assert.equal(ctl.checkFailed, null);
  const c1 = await ctl.check();
  assert.ok(c1.error);
  assert.equal(ctl.state, 'failed');
  assert.equal(ctl.checkFailed, c1.error, 'the failed check is reported as such');
  assert.ok(ctl.pending && ctl.pending.marker === 'm1', 'the attempt remains pending');
  assert.equal(ctl.needsVerification, true, 'and unconfirmed');
  const c2 = await ctl.check();
  assert.equal(c2.outcome, 'missing');
  assert.equal(ctl.checkFailed, null);
  assert.equal(ctl.needsVerification, false);
  const r2 = await ctl.submit();
  assert.equal(r2.ok, true);
  assert.equal(performs, 2);
  assert.equal(ctl.state, 'saved');
});
