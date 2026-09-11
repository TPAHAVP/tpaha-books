import { test } from 'node:test';
import assert from 'node:assert/strict';
import { MockWorkbook } from '../../site/js/workbook/mock-excel.js';
import { ExcelClient, ExcelApiError } from '../../site/js/workbook/excel-client.js';
import {
  LedgerWorkbook, ConflictError, VerificationError, UnresolvedOperationError, MarkerCollisionError,
  rowFingerprint, rowToTxn, makeMarker, identityOf, expectedSortedRows, sortedTableMatches,
} from '../../site/js/workbook/ledger-workbook.js';
import { SAMPLE_WORKBOOK as fixture } from '../../site/js/workbook/sample-workbook.js';
import { SaveController } from '../../site/js/save/save-controller.js';

const MONTHS_RE = /worksheets\/(January|February|March|April|May|June|July|August|September|October|November|December|Annual)/;

function setup() {
  const mock = new MockWorkbook(fixture);
  const client = new ExcelClient({ getToken: async () => 't', driveId: 'D1', itemId: 'I1', fetchImpl: (u, o) => mock.fetch(u, o), sleep: async () => {} });
  const wb = new LedgerWorkbook(client);
  return { mock, client, wb };
}
const memberOn = mock => new LedgerWorkbook(new ExcelClient({ getToken: async () => 't', driveId: 'D1', itemId: 'I1', fetchImpl: (u, o) => mock.fetch(u, o), sleep: async () => {} }));
const nonBlank = rows => rows.filter(r => r.some(v => v !== '' && v !== null));
const entry = (over = {}) => ({ date: '2026-09-06', type: 'Deposit', category: 'Membership', amount: '25.50', description: 'Test add', chequeNum: '', notes: '', ...over });
const noWritesToFormulaSheets = mock => mock.log.filter(e => e.method !== 'GET' && MONTHS_RE.test(e.url));
const ids = rows => nonBlank(rows).map(r => r[0]);
const identities = rows => nonBlank(rows).map(r => `${r[0]}|${r[1]}`);
const bodyJson = async mock => (await mock.fetch('https://graph.microsoft.com/v1.0/drives/D1/items/I1/workbook/tables/LOG_Table/dataBodyRange', {})).json();

// ---------------------------------------------------------------------------------------------- basics
test('load reads year, prior-year balance, categories and 12 transactions with ISO dates and fingerprints', async () => {
  const { wb } = setup();
  const snap = await wb.load();
  assert.equal(snap.year, 2026);
  assert.equal(snap.priorYearBalance, 5000);
  assert.equal(snap.categories.length, 8);
  assert.equal(snap.transactions.length, 12);
  assert.equal(snap.transactions[0].date, '2026-01-05');
  assert.equal(snap.transactions[0].rowIndex, 1, 'blank first body row skipped');
  assert.equal(snap.transactions[0].id, 1);
  assert.ok(snap.transactions.every(t => typeof t.fingerprint === 'string' && t.fingerprint.length > 0));
  assert.equal(snap.nextId, 14);
  assert.equal(wb.halted, null);
});

test('rowToTxn, rowFingerprint, identityOf and makeMarker', () => {
  const t = rowToTxn([30, 'm', 46271, 'Deposit', 'Membership', 25.5, 'x', null, undefined], 7);
  assert.equal(t.date, '2026-09-06');
  assert.equal(t.chequeNum, '');
  assert.equal(identityOf(t), '30|m');
  assert.equal(rowFingerprint([30, 'm', 46271, 'Deposit', 'Membership', 25.5, 'x', null, undefined]), rowFingerprint([30, 'm', 46271, 'Deposit', 'Membership', 25.5, 'x', '', '']));
  const a = makeMarker(), b = makeMarker();
  assert.match(a, /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{9}Z$/, 'ISO 8601 UTC with nine fraction digits');
  assert.ok(!Number.isNaN(new Date(a).getTime()), 'still parses as a date (Office Scripts using new Date() keep working)');
  assert.notEqual(a, b, 'two markers made at the same instant differ');
});

// ---------------------------------------------------------------------------------------------- add
test('addTransaction appends with the next id, sets formats, verifies by read-back, rebuilds LOG_Sorted, touches no formula sheet', async () => {
  const { mock, wb } = setup();
  await wb.load();
  const marker = makeMarker();
  const res = await wb.addTransaction(entry(), { marker, user: 'tester' });
  assert.equal(res.alreadyExisted, false);
  assert.equal(res.renumberedFrom, null);
  assert.equal(res.txn.id, 14);
  assert.equal(res.txn.amount, 25.5);
  assert.equal(nonBlank(mock.table('LOG_Table').rows).length, 13);
  const body = await bodyJson(mock);
  assert.equal(body.numberFormat[13][2], 'yyyy-mm-dd');
  assert.equal(body.numberFormat[13][5], '$#,##0.00');
  const sorted = mock.table('LOG_Sorted_Table').rows;
  assert.equal(sorted.length, 13);
  const serials = sorted.map(r => r[2]);
  assert.deepEqual(serials, [...serials].sort((a, b) => a - b));
  assert.equal(sorted.at(-1)[1], marker);
  const sep = await wb.readMonth(9);
  assert.equal(sep.subtotal.byCategory.Membership, 25.5);
  assert.deepEqual(noWritesToFormulaSheets(mock), []);
  assert.equal(mock.log.filter(e => e.method === 'PATCH' && /rows\/itemAt/.test(e.url)).length, 1, 'formats set exactly once, never retried');
});

test('adding twice with the same marker and the same content is idempotent', async () => {
  const { mock, wb } = setup();
  await wb.load();
  const marker = makeMarker();
  const a = await wb.addTransaction(entry(), { marker });
  const b = await wb.addTransaction(entry(), { marker });
  assert.equal(b.alreadyExisted, true);
  assert.equal(b.txn.id, a.txn.id);
  assert.equal(nonBlank(mock.table('LOG_Table').rows).length, 13);
});

test('R4: the same marker with DIFFERENT content is a collision, never silently accepted as saved', async () => {
  const { mock, wb } = setup();
  await wb.load();
  const marker = '2026-09-06T12:00:00.000Z#samesame';
  await wb.addTransaction(entry(), { marker });
  await assert.rejects(() => wb.addTransaction(entry({ description: 'Different member entry', amount: 99 }), { marker }), e => e instanceof MarkerCollisionError);
  assert.equal(nonBlank(mock.table('LOG_Table').rows).length, 13, 'the second entry was not written under a stolen identity');
  assert.equal(mock.table('LOG_Table').rows.some(r => r[6] === 'Different member entry'), false);
});


test('R4/S4: two rows carrying the same marker is an unresolved operation; load() does not lift the pause, a verified explicit repair does', async () => {
  const { mock, wb } = setup();
  await wb.load();
  const marker = makeMarker();
  await wb.addTransaction(entry(), { marker });
  mock.table('LOG_Table').append([99, marker, 46271, 'Deposit', 'Membership', 25.5, 'Test add', '', '']);   // a late duplicate landing
  await assert.rejects(() => wb.addTransaction(entry(), { marker }), e => e instanceof UnresolvedOperationError && e.kind === 'duplicate-marker');
  assert.ok(wb.halted, 'writes are paused');
  await assert.rejects(() => wb.addTransaction(entry({ description: 'another' }), { marker: makeMarker() }), e => e instanceof UnresolvedOperationError && e.kind === 'halted');
  const snap = await wb.load();
  assert.ok(wb.halted, 'S4: reloading does not lift the pause');
  assert.equal(snap.incident.kind, 'duplicate-marker');
  let v = await wb.verifyResolution();
  assert.equal(v.resolved, false);
  assert.ok(v.checks.some(c => c.ok === false));
  const dup = snap.transactions.find(t => t.id === 99);
  const r = await wb.removeCopy(identityOf(dup), dup.fingerprint, { confirm: true });   // explicit, member-confirmed
  assert.equal(r.removed, true);
  v = await wb.verifyResolution();
  assert.equal(v.resolved, true);
  assert.equal(wb.halted, null);
  assert.equal(wb.incident.resolution, 'verified');
  const ok = await wb.addTransaction(entry({ description: 'another' }), { marker: makeMarker() });
  assert.equal(ok.txn.id, 15);
});

test('an ambiguous failure after the row was written: completeAppend finds it and finishes; a retry with the same marker does not duplicate', async () => {
  const { mock, wb } = setup();
  await wb.load();
  const marker = makeMarker();
  mock.failNext({ match: /rows\/add/, networkError: true, afterApply: true });
  await assert.rejects(() => wb.addTransaction(entry(), { marker }), e => e instanceof ExcelApiError && e.ambiguous === true);
  const done = await wb.completeAppend(marker, entry());
  assert.ok(done && done.alreadyExisted === true && done.txn.id === 14);
  assert.equal(await wb.completeAppend(makeMarker(), entry()), null, 'unknown marker -> nothing there');
  const again = await wb.addTransaction(entry(), { marker });
  assert.equal(again.alreadyExisted, true);
  assert.equal(nonBlank(mock.table('LOG_Table').rows).length, 13);
  assert.equal(mock.table('LOG_Sorted_Table').rows.length, 13);
  assert.equal((await bodyJson(mock)).numberFormat[13][2], 'yyyy-mm-dd');
});

test('R5: a lost response on the format-only PATCH is not retried; the add still succeeds with a warning', async () => {
  const { mock, wb } = setup();
  await wb.load();
  mock.failNext({ match: /rows\/itemAt/, networkError: true });   // fails before applying
  const res = await wb.addTransaction(entry(), { marker: makeMarker() });
  assert.equal(res.txn.id, 14);
  assert.equal(mock.log.filter(e => e.method === 'PATCH' && /rows\/itemAt/.test(e.url)).length, 1, 'no automatic second PATCH');
  assert.ok(res.warnings.some(w => /format/i.test(w)));
});

test('two members adding at once: the colliding id is repaired, both rows survive, uniqueness is proven at the end', async () => {
  const { mock, wb } = setup();
  await wb.load();
  let raced = false;
  mock.beforeRespond = async e => {
    if (!raced && e.method === 'POST' && /rows\/add/.test(e.url)) { raced = true; mock.table('LOG_Table').append([14, 'other-member', 46271, 'Deposit', 'Grub Box', 5, 'Other member', '', '']); }
  };
  const res = await wb.addTransaction(entry(), { marker: makeMarker() });
  assert.equal(res.renumberedFrom, 14);
  assert.equal(res.txn.id, 15);
  const all = ids(mock.table('LOG_Table').rows);
  assert.equal(new Set(all).size, all.length);
  assert.equal(mock.table('LOG_Sorted_Table').rows.length, 14);
});

test('R4: if the number cannot be made unique after the repair rounds, the result is unresolved, not saved', async () => {
  const { mock, wb } = setup();
  await wb.load();
  // Another writer keeps taking every number we pick.
  mock.beforeRespond = async e => {
    if (e.method === 'PATCH' && /rows\/itemAt/.test(e.url) && e.body && /"values"/.test(e.body)) {
      const id = JSON.parse(e.body).values[0][0];
      mock.table('LOG_Table').append([id, `thief-${id}`, 46271, 'Deposit', 'Grub Box', 1, 'thief', '', '']);
    }
    if (e.method === 'POST' && /rows\/add/.test(e.url) && !/thief/.test(e.body || '') && !e.__seen) { e.__seen = true; mock.table('LOG_Table').append([14, 'thief-14', 46271, 'Deposit', 'Grub Box', 1, 'thief', '', '']); }
  };
  await assert.rejects(() => wb.addTransaction(entry(), { marker: makeMarker() }), e => e instanceof UnresolvedOperationError && e.kind === 'duplicate-number');
  assert.ok(mock.table('LOG_Table').rows.some(r => r[6] === 'Test add'), 'the entry itself is in the workbook');
  assert.ok(wb.halted);
});

// ---------------------------------------------------------------------------------------------- delete
test('deleteTransaction refuses a record someone changed, then works with a fresh reference', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const target = snap.transactions.find(t => t.id === 13);
  mock.table('LOG_Table').setCell(target.rowIndex, 6, 'Edited directly in Excel');
  await assert.rejects(() => wb.deleteTransaction(target), e => e instanceof ConflictError && e.reason === 'changed' && e.current.description === 'Edited directly in Excel');
  assert.equal(nonBlank(mock.table('LOG_Table').rows).length, 12);
  const fresh = (await wb.load()).transactions.find(t => t.id === 13);
  const res = await wb.deleteTransaction(fresh);
  assert.equal(res.deleted, true);
  const rows = nonBlank(mock.table('LOG_Table').rows);
  assert.equal(rows.length, 11);
  assert.ok(!rows.some(r => r[0] === 13));
  assert.equal(mock.table('LOG_Sorted_Table').rows.length, 11);
  const deletes = mock.log.filter(e => e.method === 'DELETE' && /LOG_Table\/rows/.test(e.url));
  assert.equal(deletes.length, 1, 'one DELETE, sent exactly once');
  assert.equal(mock.log.filter(e => /\$batch$/.test(e.url)).length, 0, 'no pre-decided read+delete batch');
  const i = mock.log.indexOf(deletes[0]);
  assert.ok(mock.log[i - 1].method === 'GET' && /LOG_Table\/dataBodyRange/.test(mock.log[i - 1].url), 'the request just before the DELETE is a fresh read of the table, evaluated before sending');
});

test('deleting a record that no longer exists is a missing-record conflict', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const target = snap.transactions.find(t => t.id === 13);
  mock.table('LOG_Table').remove(target.rowIndex);
  await assert.rejects(() => wb.deleteTransaction(target), e => e instanceof ConflictError && e.reason === 'missing');
});

test('R2: deleting #13 while another member legitimately edits #2 leaves exactly one #2 with the new text', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const target = snap.transactions.find(t => t.id === 13);
  const other = snap.transactions.find(t => t.id === 2);
  let once = false;
  mock.beforeRespond = async e => { if (!once && e.method === 'DELETE') { once = true; mock.table('LOG_Table').setCell(other.rowIndex, 6, 'Legitimate other edit'); } };
  const res = await wb.deleteTransaction(target);
  assert.equal(res.deleted, true);
  const rows = mock.table('LOG_Table').rows;
  assert.equal(rows.filter(r => r[0] === 2).length, 1);
  assert.equal(rows.find(r => r[0] === 2)[6], 'Legitimate other edit');
  assert.ok(!rows.some(r => r[0] === 13));
  assert.equal(nonBlank(rows).length, 11);
});

test('R2: deleting #13 while another member appends keeps the appended row', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const target = snap.transactions.find(t => t.id === 13);
  let once = false;
  mock.beforeRespond = async e => { if (!once && e.method === 'DELETE') { once = true; mock.table('LOG_Table').append([14, 'other', 46271, 'Deposit', 'Grub Box', 5, 'Other member', '', '']); } };
  const res = await wb.deleteTransaction(target);
  assert.equal(res.deleted, true);
  assert.ok(mock.table('LOG_Table').rows.some(r => r[1] === 'other'));
  assert.equal(nonBlank(mock.table('LOG_Table').rows).length, 12);
});


test('S1: another member deletes #2 just before our check: the fresh read is evaluated first, the delete re-aims at #7, every other transaction survives', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const target = snap.transactions.find(t => t.id === 7);
  const other = snap.transactions.find(t => t.id === 2);
  const before = identities(mock.table('LOG_Table').rows);
  let gets = 0, done = false;
  mock.beforeRespond = async e => {
    if (e.method === 'GET' && /LOG_Table\/dataBodyRange/.test(e.url)) { gets++; if (gets === 2 && !done) { done = true; mock.table('LOG_Table').remove(other.rowIndex); } }   // between our locate-read and the pre-delete check
  };
  const res = await wb.deleteTransaction(target);
  assert.equal(res.deleted, true);
  assert.ok(res.notices.some(n => /removed: #2/.test(n)), 'the other member\'s delete is reported as a notice');
  const after = identities(mock.table('LOG_Table').rows);
  const lost = before.filter(x => !after.includes(x));
  assert.deepEqual(lost.sort(), [`2|${other.timestamp}`, `7|${target.timestamp}`].sort(), 'exactly #2 (by the other member) and #7 (by us) are gone');
  assert.equal(mock.log.filter(e => e.method === 'DELETE' && /LOG_Table\/rows/.test(e.url)).length, 1);
  assert.equal(wb.halted, null);
});


test('S1 (residual race): a shift between the check and the DELETE is detected by identity, named, paused and never repaired by guessing; a verified explicit restore resolves it', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const target = snap.transactions.find(t => t.id === 7);
  const victim = mock.table('LOG_Table').rows[target.rowIndex + 1];
  let once = false;
  mock.beforeRespond = async e => { if (!once && e.method === 'DELETE') { once = true; mock.table('LOG_Table').remove(0); } };   // blank row removed by someone at the last instant
  await assert.rejects(() => wb.deleteTransaction(target), e => e instanceof UnresolvedOperationError && e.kind === 'wrong-row-deleted' && e.details.victim.id === victim[0]);
  const rows = mock.table('LOG_Table').rows;
  assert.ok(rows.some(r => r[0] === 7), 'the target is untouched');
  assert.equal(rows.filter(r => r[0] === victim[0]).length, 0, 'the named row is gone and was NOT recreated by guessing');
  assert.equal(nonBlank(rows).length, 11);
  assert.ok(wb.halted);
  await wb.load();
  assert.ok(wb.halted, 'S4: a reload does not lift the pause');
  let v = await wb.verifyResolution();
  assert.equal(v.resolved, false);
  assert.ok(v.checks.some(c => c.name === `Transaction #${victim[0]} is present again` && c.ok === false));
  await assert.rejects(() => wb.acknowledgeIncident({ confirm: true }), /cannot be acknowledged/);
  const back = await wb.reappendRow(wb.incident.details.victim.values, { confirm: true });   // explicit, member-confirmed
  assert.equal(back.id, victim[0]);
  v = await wb.verifyResolution();
  assert.equal(v.resolved, true);
  assert.equal(wb.halted, null);
  assert.equal(nonBlank(mock.table('LOG_Table').rows).length, 12);
});

// ---------------------------------------------------------------------------------------------- correct
test('updateTransaction keeps id and timestamp, applies the change, verifies, leaves one copy, rebuilds LOG_Sorted, never writes a formula sheet', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const ref = snap.transactions.find(t => t.id === 2);
  const t2 = await wb.updateTransaction(ref, { description: 'Corrected description', amount: '91' });
  assert.equal(t2.id, 2);
  assert.equal(t2.timestamp, ref.timestamp);
  assert.equal(t2.amount, 91);
  const rows = mock.table('LOG_Table').rows;
  assert.equal(rows.filter(r => r[0] === 2).length, 1);
  assert.equal(rows.find(r => r[0] === 2)[6], 'Corrected description');
  assert.equal(nonBlank(rows).length, 12);
  assert.ok(mock.table('LOG_Sorted_Table').rows.some(r => r[0] === 2 && r[6] === 'Corrected description' && r[5] === 91));
  await assert.rejects(() => wb.updateTransaction(ref, { amount: '92' }), e => e instanceof ConflictError && e.reason === 'changed');
  assert.deepEqual(noWritesToFormulaSheets(mock), []);
  assert.equal(mock.log.filter(e => e.method === 'PATCH' && /LOG_Table\/rows\/itemAt/.test(e.url) && /"values"/.test(e.body || '')).length, 0, 'no in-place value overwrite of a row by index');
});


test('S1 (residual race) during a correction: the misdirected delete is detected by identity and named, the leftover old copy is reported, nothing is guessed; explicit repairs plus verified resolution', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const ref = snap.transactions.find(t => t.id === 2);
  const before = identities(mock.table('LOG_Table').rows);
  let once = false;
  mock.beforeRespond = async e => { if (!once && e.method === 'DELETE') { once = true; mock.table('LOG_Table').remove(0); } };
  let err;
  try { await wb.updateTransaction(ref, { description: 'Changed' }); } catch (e) { err = e; }
  const rows = mock.table('LOG_Table').rows;
  assert.ok(err instanceof UnresolvedOperationError, 'the shift is reported, not papered over');
  assert.equal(err.kind, 'wrong-row-deleted');
  assert.equal(err.details.duplicateCopy, true, 'the report says the old copy of #2 is still there');
  assert.equal(rows.filter(r => r[0] === 2 && r[6] === 'Changed').length, 1, 'the corrected copy exists exactly once');
  const after = identities(rows);
  const lost = before.filter(x => !after.includes(x));
  assert.equal(lost.length, 1, 'exactly the one row the misdirected delete hit is missing, and it is named');
  assert.equal(`${err.details.victim.id}|${err.details.victim.timestamp}`, lost[0]);
  assert.ok(wb.halted);
  await wb.load();
  assert.ok(wb.halted, 'S4: a reload does not lift the pause');
  assert.equal((await wb.verifyResolution()).resolved, false);
  await wb.reappendRow(err.details.victim.values, { confirm: true });                       // explicit repair 1
  assert.equal((await wb.verifyResolution()).resolved, false, 'the old copy of #2 is still there');
  const fixed = await wb.removeCopy(err.details.identity, err.details.oldFingerprint, { confirm: true });   // explicit repair 2
  assert.equal(fixed.removed, true);
  assert.equal((await wb.verifyResolution()).resolved, true);
  assert.equal(wb.halted, null);
  assert.deepEqual(identities(mock.table('LOG_Table').rows).sort(), before.sort(), 'every identity is back exactly once');
  assert.equal(mock.table('LOG_Table').rows.find(r => r[0] === 2)[6], 'Changed');
});

test('R5: a lost response on the correction append is verified by reading, not by sending the append again', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const ref = snap.transactions.find(t => t.id === 2);
  mock.failNext({ match: /LOG_Table\/rows\/add/, networkError: true, afterApply: true });
  const t2 = await wb.updateTransaction(ref, { amount: 90 });
  assert.equal(t2.amount, 90);
  assert.equal(mock.log.filter(e => e.method === 'POST' && /LOG_Table\/rows\/add/.test(e.url)).length, 1, 'exactly one append request');
  assert.equal(mock.table('LOG_Table').rows.filter(r => r[0] === 2).length, 1);
});


test('R5: a lost response on the cleanup DELETE is verified by reading; a delete that never applied leaves a duplicate copy that is reported, paused, and removed only explicitly', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const ref = snap.transactions.find(t => t.id === 2);
  mock.failNext({ match: /LOG_Table\/rows\/\d+$/, networkError: true, afterApply: true });
  const t2 = await wb.updateTransaction(ref, { amount: 90 });
  assert.equal(t2.amount, 90);
  assert.equal(mock.table('LOG_Table').rows.filter(r => r[0] === 2).length, 1, 'the delete had applied; the read proved it');
  assert.equal(mock.log.filter(e => e.method === 'DELETE').length, 1, 'never re-sent');
  assert.ok(t2.notices.some(n => /answer to the delete was lost/.test(n)));
  // A DELETE that fails BEFORE applying: the old copy remains -> duplicate copy, paused, explicit repair, verified resolution.
  const snap2 = await wb.load();
  const ref3 = snap2.transactions.find(t => t.id === 3);
  mock.failNext({ match: /LOG_Table\/rows\/\d+$/, networkError: true });
  let err;
  try { await wb.updateTransaction(ref3, { amount: 46 }); } catch (e) { err = e; }
  assert.ok(err instanceof UnresolvedOperationError && err.kind === 'duplicate-copy', String(err));
  assert.equal(mock.table('LOG_Table').rows.filter(r => r[0] === 3).length, 2);
  assert.equal(mock.log.filter(e => e.method === 'DELETE').length, 2, 'the failed delete was not retried');
  await wb.load();
  assert.ok(wb.halted, 'a reload keeps the pause');
  assert.equal((await wb.verifyResolution()).resolved, false);
  const fixed = await wb.removeCopy(err.details.identity, err.details.oldFingerprint, { confirm: true });
  assert.equal(fixed.removed, true);
  assert.equal(mock.table('LOG_Table').rows.filter(r => r[0] === 3).length, 1);
  assert.equal(mock.table('LOG_Table').rows.find(r => r[0] === 3)[5], 46);
  assert.equal((await wb.verifyResolution()).resolved, true);
  assert.equal(wb.halted, null);
});

// ---------------------------------------------------------------------------------------------- helpers
test('rebuildSorted repairs extra, missing and misordered rows, reports consistency against a fresh LOG read, and is a no-op when correct', async () => {
  const { mock, wb } = setup();
  await wb.load();
  const sorted = mock.table('LOG_Sorted_Table');
  sorted.append([99, 'junk', 46000, 'Deposit', 'Membership', 1, 'junk', '', '']);
  let r = await wb.rebuildSorted();
  assert.equal(r.changed, true);
  assert.equal(r.consistent, true);
  assert.equal(sorted.rows.length, 12);
  sorted.remove(3);
  r = await wb.rebuildSorted();
  assert.equal(sorted.rows.length, 12);
  const a = sorted.rows[0], b = sorted.rows[1];
  for (let c = 0; c < 9; c++) { sorted.setCell(0, c, b[c]); sorted.setCell(1, c, a[c]); }
  r = await wb.rebuildSorted();
  const serials = sorted.rows.map(x => x[2]);
  assert.deepEqual(serials, [...serials].sort((x, y) => x - y));
  r = await wb.rebuildSorted();
  assert.equal(r.changed, false);
  assert.equal(mock.log.filter(e => e.method === 'PATCH' && /LOG_Sorted/.test(e.url)).length > 0, true);
});

test('rebuildSorted retries against a LOG that changed underneath it and reports when still inconsistent', async () => {
  const { mock, wb } = setup();
  await wb.load();
  let n = 0;
  mock.beforeRespond = async e => { if (e.method === 'PATCH' && /LOG_Sorted/.test(e.url) && n < 1) { n++; mock.table('LOG_Table').append([14, 'late', 46271, 'Deposit', 'Grub Box', 5, 'Late add', '', '']); } };
  mock.table('LOG_Sorted_Table').append([99, 'junk', 46000, 'Deposit', 'Membership', 1, 'junk', '', '']);
  const r = await wb.rebuildSorted();
  assert.equal(r.consistent, true, 'a second round caught the late append');
  assert.equal(mock.table('LOG_Sorted_Table').rows.length, 13);
});

test('setPriorYearBalance verifies the current value before writing', async () => {
  const { mock, wb } = setup();
  await wb.load();
  assert.equal(await wb.setPriorYearBalance(100, 5000), 100);
  assert.equal(mock.getCell('ENTRY', 'B23'), 100);
  await assert.rejects(() => wb.setPriorYearBalance(200, 5000), e => e instanceof ConflictError);
  assert.equal(mock.getCell('ENTRY', 'B23'), 100);
});

test('readMonth and readAnnual return the workbook figures', async () => {
  const { wb } = setup();
  await wb.load();
  const jul = await wb.readMonth(7);
  assert.equal(jul.rows.length, 3);
  assert.equal(jul.subtotal.amount, 860);
  assert.equal(jul.recon.closing, 5321.75);
  const annual = await wb.readAnnual();
  assert.equal(annual.rows.find(r => r[0] === 'Net')[13], 340.5);
});

test('selfTest adds one labelled row, removes it, and proves the compared areas are identical', async () => {
  const { mock, wb } = setup();
  await assert.rejects(() => wb.selfTest({ confirmTestCopy: false }), /test copy/);
  const before = mock.snapshot();
  const report = await wb.selfTest({ confirmTestCopy: true });
  assert.equal(report.ok, true, JSON.stringify(report.steps));
  assert.equal(report.restored, true);
  assert.ok(report.compared.includes('LOG values and number formats'));
  assert.ok(report.compared.some(x => /formulas/.test(x)));
  assert.deepEqual(mock.snapshot(), before);
  assert.deepEqual(noWritesToFormulaSheets(mock), []);
});

// ---------------------------------------------------------------------------------------------- resuming a correction
test('a correction whose read-back failed is finished by completeCorrection without a second append', async () => {
  const { mock, wb } = setup();
  const ref = (await wb.load()).transactions.find(t => t.id === 2);
  mock.beforeRespond = async e => {
    if (e.method === 'POST' && /LOG_Table\/rows\/add/.test(e.url)) {
      mock.beforeRespond = null;
      mock.failNext({ match: 'LOG_Table/dataBodyRange', networkError: true });   // the read-back fails,
      mock.failNext({ match: 'LOG_Table/dataBodyRange', networkError: true });   // and so does the one read retry
    }
  };
  await assert.rejects(() => wb.updateTransaction(ref, { description: 'Corrected once' }), e => e instanceof ExcelApiError && !e.ambiguous);
  assert.equal(nonBlank(mock.table('LOG_Table').rows).filter(r => r[0] === 2).length, 2, 'the corrected copy landed and the old copy remains');
  assert.equal(await wb.completeCorrection(ref, { description: 'Something else' }), null, 'a different correction is reported as not landed');
  const done = await wb.completeCorrection(ref, { description: 'Corrected once' });
  assert.equal(done.description, 'Corrected once');
  assert.deepEqual(done.notices, []);
  const rows = nonBlank(mock.table('LOG_Table').rows).filter(r => r[0] === 2);
  assert.equal(rows.length, 1);
  assert.equal(rows[0][6], 'Corrected once');
  assert.equal(mock.log.filter(e => e.method === 'POST' && /LOG_Table\/rows\/add/.test(e.url)).length, 1, 'exactly one append');
  assert.equal(wb.halted, null);
});


test('S3: with the SaveController, Check after a failed correction only reads and reports "incomplete"; Save then finishes it without a second append', async () => {
  const { mock, wb } = setup();
  const ref = (await wb.load()).transactions.find(t => t.id === 5);
  const ctl = new SaveController({ perform: d => wb.updateTransaction(ref, d), inspect: d => wb.inspectCorrection(ref, d), resume: d => wb.completeCorrection(ref, d) });
  mock.beforeRespond = async e => {
    if (e.method === 'POST' && /LOG_Table\/rows\/add/.test(e.url)) {
      mock.beforeRespond = null;
      mock.failNext({ match: 'LOG_Table/dataBodyRange', networkError: true });
      mock.failNext({ match: 'LOG_Table/dataBodyRange', networkError: true });
    }
  };
  ctl.edit({ amount: '380' });
  const r1 = await ctl.submit();
  assert.equal(r1.ok, false);
  assert.equal(ctl.state, 'failed');
  const start = mock.log.length;
  const c = await ctl.check();
  assert.equal(c.outcome, 'incomplete');
  assert.ok(ctl.incomplete.missingSteps.includes('removal of the old copy'));
  assert.deepEqual(mock.log.slice(start).filter(e => e.method !== 'GET').map(e => e.method), [], 'S3: checking issues no writes');
  assert.equal(ctl.state, 'failed');
  const r2 = await ctl.submit();
  assert.equal(r2.ok, true);
  assert.equal(r2.resolvedPending, true);
  assert.equal(ctl.state, 'saved');
  const rows = nonBlank(mock.table('LOG_Table').rows).filter(r => r[0] === 5);
  assert.equal(rows.length, 1);
  assert.equal(rows[0][5], 380);
  assert.equal(mock.log.filter(e => e.method === 'POST' && /LOG_Table\/rows\/add/.test(e.url)).length, 1);
});

// ---------------------------------------------------------------------------------------------- S1: the fresh read is evaluated before any positional request
test('S1: a correction whose old copy moved just before the check: the fresh read is evaluated first, the delete re-aims, every transaction survives', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const ref = snap.transactions.find(t => t.id === 2);
  const before = identities(mock.table('LOG_Table').rows);
  let done = false;
  mock.beforeRespond = async e => {
    // the first table read after the format PATCH is the pre-delete check: the blank row disappears just before it
    if (!done && e.method === 'GET' && /LOG_Table\/dataBodyRange/.test(e.url) && mock.log.some(x => x.method === 'PATCH' && /rows\/itemAt/.test(x.url))) { done = true; mock.table('LOG_Table').remove(0); }
  };
  const t2 = await wb.updateTransaction(ref, { description: 'Changed' });
  assert.equal(t2.description, 'Changed');
  const rows = mock.table('LOG_Table').rows;
  assert.equal(rows.filter(r => r[0] === 2).length, 1);
  assert.deepEqual(identities(rows).sort(), before.sort(), 'no transaction lost');
  assert.equal(mock.log.filter(e => e.method === 'DELETE' && /LOG_Table\/rows/.test(e.url)).length, 1);
  assert.equal(wb.halted, null);
});

test('S1: renumbering re-reads before the positional PATCH: when our new row moved, nothing at its old index is renumbered', async () => {
  const { mock, wb } = setup();
  await wb.load();
  let raced = false, shifted = false;
  mock.beforeRespond = async e => {
    if (!raced && e.method === 'POST' && /rows\/add/.test(e.url)) { raced = true; mock.table('LOG_Table').append([14, 'other-member', 46271, 'Deposit', 'Grub Box', 5, 'Other member', '', '']); }
    // the first table read after the format PATCH is the pre-renumber check: someone removes the blank row just before it
    if (raced && !shifted && e.method === 'GET' && /LOG_Table\/dataBodyRange/.test(e.url) && mock.log.some(x => x.method === 'PATCH' && /numberFormat/.test(x.body || ''))) { shifted = true; mock.table('LOG_Table').remove(0); }
  };
  const res = await wb.addTransaction(entry(), { marker: makeMarker() });
  assert.equal(res.renumberedFrom, 14);
  assert.equal(res.txn.id, 15);
  const rows = nonBlank(mock.table('LOG_Table').rows);
  assert.equal(rows.find(r => r[1] === 'other-member')[0], 14, 'the other member\'s row kept its number');
  assert.deepEqual(rows.filter(r => r[0] <= 13).map(r => r[0]).sort((a, b) => a - b), [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13], 'no historical row renumbered');
  assert.equal(mock.log.filter(e => e.method === 'PATCH' && /rows\/itemAt/.test(e.url) && /"values"/.test(e.body || '')).length, 1, 'one renumber PATCH, on the re-located row');
  assert.equal(wb.halted, null);
});

// ---------------------------------------------------------------------------------------------- S3: read-only inspection
test('S3: inspectAppend and inspectCorrection issue only GETs and report missing / incomplete / landed / conflict', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const marker = makeMarker();
  const e = entry({ description: 'Inspect me' });
  const writes = from => mock.log.slice(from).filter(x => x.method !== 'GET' && !/Session/.test(x.url)).map(x => `${x.method} ${x.url}`);
  let start = mock.log.length;
  assert.deepEqual(await wb.inspectAppend(marker, e), { outcome: 'missing' });
  assert.deepEqual(writes(start), []);
  mock.failNext({ match: /LOG_Table\/rows\/add/, networkError: true, afterApply: true });
  await assert.rejects(() => wb.addTransaction(e, { marker }), err => err instanceof ExcelApiError && err.ambiguous);
  start = mock.log.length;
  const inc = await wb.inspectAppend(marker, e);
  assert.equal(inc.outcome, 'incomplete');
  assert.equal(inc.txn.id, 14);
  assert.ok(inc.missingSteps.some(x => /format/.test(x)) && inc.missingSteps.some(x => /sorted/.test(x)), String(inc.missingSteps));
  assert.deepEqual(writes(start), [], 'inspection wrote nothing');
  const col = await wb.inspectAppend(marker, entry({ description: 'Different content' }));
  assert.equal(col.outcome, 'conflict');
  assert.ok(col.error instanceof MarkerCollisionError);
  const done = await wb.completeAppend(marker, e);     // the explicit finish (a Save)
  assert.equal(done.alreadyExisted, true);
  start = mock.log.length;
  const landed = await wb.inspectAppend(marker, e);
  assert.equal(landed.outcome, 'landed');
  assert.equal(landed.result.txn.id, 14);
  assert.deepEqual(writes(start), []);
  // corrections
  const ref = snap.transactions.find(t => t.id === 2);
  start = mock.log.length;
  assert.deepEqual(await wb.inspectCorrection(ref, { amount: 77 }), { outcome: 'missing' });
  mock.table('LOG_Table').append([...ref.values.slice(0, 5), 77, ...ref.values.slice(6)]);   // the corrected copy landed; the old copy is still there
  const ic = await wb.inspectCorrection(ref, { amount: 77 });
  assert.equal(ic.outcome, 'incomplete');
  assert.ok(ic.missingSteps.includes('removal of the old copy'));
  assert.deepEqual(writes(start), []);
  const fin = await wb.completeCorrection(ref, { amount: 77 });
  assert.equal(fin.amount, 77);
  start = mock.log.length;
  assert.equal((await wb.inspectCorrection(ref, { amount: 77 })).outcome, 'landed');
  assert.deepEqual(writes(start), []);
  const gone = snap.transactions.find(t => t.id === 13);
  mock.table('LOG_Table').remove(mock.table('LOG_Table').rows.findIndex(r => r[0] === 13));
  const cf = await wb.inspectCorrection(gone, { amount: 1 });
  assert.equal(cf.outcome, 'conflict');
  assert.ok(cf.error instanceof ConflictError && cf.error.reason === 'missing');
});

// ---------------------------------------------------------------------------------------------- S4: incidents survive reload; resolution is verified
test('S4: load() pauses on duplicate copies already in the workbook (integrity); an incident restored after a reload re-arms the pause; identical copies are removed one at a time', async () => {
  const { mock, wb } = setup();
  const snap0 = await wb.load();
  const row = snap0.transactions.find(t => t.id === 5);
  mock.table('LOG_Table').append(row.values);        // a copy pasted in Excel, or left by another device's interrupted correction
  const snap = await wb.load();
  assert.ok(wb.halted);
  assert.equal(snap.incident.kind, 'integrity');
  assert.equal(snap.incident.details.copies.length, 2);
  await assert.rejects(() => wb.addTransaction(entry(), { marker: makeMarker() }), e => e.kind === 'halted');
  const record = JSON.parse(JSON.stringify(wb.incident));
  const wb2 = memberOn(mock);                          // the page after a reload
  assert.equal(wb2.halted, null);
  wb2.restoreIncident(record);
  assert.ok(wb2.halted);
  await assert.rejects(() => wb2.deleteTransaction(row), e => e.kind === 'halted');
  assert.equal((await wb2.verifyResolution()).resolved, false);
  const refused = await wb2.removeCopy(identityOf(row), row.fingerprint, { confirm: true });
  assert.equal(refused.removed, false);
  assert.equal(refused.reason, 'multiple', 'two identical copies: removeCopy does not guess without being told');
  const one = await wb2.removeCopy(identityOf(row), row.fingerprint, { confirm: true, allowIdentical: true });
  assert.equal(one.removed, true);
  assert.equal(mock.table('LOG_Table').rows.filter(r => r[0] === 5).length, 1);
  assert.equal((await wb2.verifyResolution()).resolved, true);
  assert.equal(wb2.halted, null);
  assert.ok(wb.halted, 'the other page still holds its record until it verifies too');
  assert.equal((await wb.verifyResolution()).resolved, true);
  assert.equal(wb.halted, null);
});

test('S4: acknowledgement is refused while a check fails and accepted only when the checks are inconclusive', async () => {
  const { mock, wb } = setup();
  await wb.load();
  wb.restoreIncident({ kind: 'wrong-row-deleted', message: 'x', details: { victim: null }, at: '2026-09-07T00:00:00Z', resolved: false });
  let v = await wb.verifyResolution();
  assert.equal(v.resolved, false);
  assert.equal(v.inconclusive, true);
  await assert.rejects(() => wb.acknowledgeIncident({}), /confirmation/);
  const ack = await wb.acknowledgeIncident({ confirm: true, by: 'tester' });
  assert.equal(ack.resolution, 'acknowledged');
  assert.equal(wb.halted, null);
  const snap = await wb.load();
  const row = snap.transactions.find(t => t.id === 4);
  mock.table('LOG_Table').append(row.values);
  await wb.load();
  assert.equal(wb.halted.kind, 'integrity');
  await assert.rejects(() => wb.acknowledgeIncident({ confirm: true }), /cannot be acknowledged/);
  assert.ok(wb.halted);
});

// ---------------------------------------------------------------------------------------------- T1: one complete operation at a time per page
test('T1: two deletes confirmed from one page at once are serialised: the second reads fresh positions; only the two intended rows disappear', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const before = identities(mock.table('LOG_Table').rows);
  const results = await Promise.allSettled([wb.deleteTransaction(snap.transactions.find(t => t.id === 2)), wb.deleteTransaction(snap.transactions.find(t => t.id === 3))]);
  assert.deepEqual(results.map(r => r.status), ['fulfilled', 'fulfilled']);
  const after = identities(mock.table('LOG_Table').rows);
  const lost = before.filter(x => !after.includes(x)).map(x => Number(x.split('|')[0])).sort((a, b) => a - b);
  assert.deepEqual(lost, [2, 3], 'exactly the two selected rows are gone');
  assert.ok(after.some(x => x.startsWith('4|')), '#4 survives with its content');
  assert.equal(wb.halted, null, 'no incident: nothing hit a wrong row');
  assert.equal(mock.log.filter(e => e.method === 'DELETE' && /LOG_Table\/rows/.test(e.url)).length, 2);
});

test('T1: mixed overlapping operations (add, correction, delete) from one page complete in order with fresh reads; every unrelated row survives', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const ref5 = snap.transactions.find(t => t.id === 5), ref13 = snap.transactions.find(t => t.id === 13);
  const before = identities(mock.table('LOG_Table').rows);
  assert.equal(wb.busy, false);
  const ops = [wb.addTransaction(entry({ description: 'Added while others run' }), { marker: makeMarker() }), wb.updateTransaction(ref5, { amount: '380' }), wb.deleteTransaction(ref13)];
  assert.equal(wb.busy, true, 'busy while operations are queued or running');
  const [added, corrected, deleted] = await Promise.all(ops);
  assert.equal(wb.busy, false);
  assert.equal(added.txn.id, 14);
  assert.equal(corrected.amount, 380);
  assert.equal(deleted.deleted, true);
  const rows = nonBlank(mock.table('LOG_Table').rows);
  assert.equal(rows.length, 12);
  assert.equal(rows.filter(r => r[0] === 5).length, 1);
  assert.equal(rows.find(r => r[0] === 5)[5], 380);
  assert.ok(!rows.some(r => r[0] === 13));
  assert.ok(rows.some(r => r[6] === 'Added while others run'));
  const after = identities(rows);
  assert.deepEqual(before.filter(x => !after.includes(x)), [`13|${ref13.timestamp}`], 'only the deleted row is missing');
  assert.equal(wb.halted, null);
  assert.equal((await wb.rebuildSorted()).consistent, true);
});

test('T1: a queued operation re-checks the pause when it starts, and the connection test refuses to overlap itself', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const target = snap.transactions.find(t => t.id === 7);
  let once = false;
  mock.beforeRespond = async e => { if (!once && e.method === 'DELETE') { once = true; mock.table('LOG_Table').remove(0); } };
  const results = await Promise.allSettled([wb.deleteTransaction(target), wb.addTransaction(entry(), { marker: makeMarker() })]);
  assert.equal(results[0].status, 'rejected');
  assert.equal(results[0].reason.kind, 'wrong-row-deleted');
  assert.equal(results[1].status, 'rejected');
  assert.equal(results[1].reason.kind, 'halted', 'the add queued behind the delete saw the new pause and did not write');
  assert.ok(!mock.table('LOG_Table').rows.some(r => r[6] === 'Test add'));
  const { wb: wb2 } = setup();
  const both = await Promise.all([wb2.selfTest({ confirmTestCopy: true }), wb2.selfTest({ confirmTestCopy: true })]);
  assert.equal(both.filter(r => r.ok).length, 1);
  assert.match(both.find(r => !r.ok).error, /already running/);
});

// --------------------------------------------------------------- compatibility with the workbook's Office Scripts
// Faithful ports of the shared helpers in SubmitEntry / DeleteTransaction (source read 2026-09-10). Those scripts
// rewrite the WHOLE of LOG_Table and LOG_Sorted_Table on every save and delete, and refuse to run at all when
// validateLog or verifySorted fails, so anything this app leaves behind must satisfy both.
const scriptTransactions = rows => rows.filter(r => r.some(v => v !== ''));
const scriptOrdered = rows => scriptTransactions(rows).slice().sort((a, b) => a[2] - b[2] || a[0] - b[0]);
const scriptVerifySorted = (log, sorted) => JSON.stringify(scriptOrdered(log)) === JSON.stringify(scriptTransactions(sorted));
function scriptValidateEntry(date, type, category, amount, categories, year) {
  if (typeof date !== 'number' || !Number.isInteger(date) || new Date((date - 25569) * 86400000).getUTCFullYear() !== year) throw new Error(`script rejects the date ${date}`);
  if (type !== 'Deposit' && type !== 'Withdrawal') throw new Error(`script rejects the type ${type}`);
  if (!categories.some(r => r[1] === category && r[2] === type)) throw new Error(`script rejects the category ${category}`);
  if (typeof amount !== 'number' || !Number.isFinite(amount) || amount <= 0 || Math.abs(amount * 100 - Math.round(amount * 100)) > 0.000001) throw new Error(`script rejects the amount ${amount}`);
}
function scriptValidateLog(rows, categories, year) {
  const seen = [];
  for (const r of scriptTransactions(rows)) {
    if (typeof r[0] !== 'number' || !Number.isSafeInteger(r[0]) || r[0] < 1 || seen.includes(r[0])) throw new Error(`script rejects the transaction number ${r[0]}`);
    seen.push(r[0]);
    scriptValidateEntry(r[2], r[3], r[4], r[5], categories, year);
  }
}

test('Office Scripts: LOG_Sorted keeps the order the scripts require (date, then transaction number), including after a correction moves a row to the end of LOG', async () => {
  const { mock, wb } = setup();
  await wb.load();
  await wb.addTransaction(entry({ date: '2026-05-04', description: 'Same day A' }), { marker: makeMarker() });
  await wb.addTransaction(entry({ date: '2026-05-04', description: 'Same day B' }), { marker: makeMarker() });
  assert.ok(scriptVerifySorted(mock.table('LOG_Table').rows, mock.table('LOG_Sorted_Table').rows), 'after two adds on the same date');
  const ref = (await wb.load()).transactions.find(t => t.description === 'Same day A');
  await wb.updateTransaction(ref, { amount: '77' });
  const log = mock.table('LOG_Table').rows, sorted = mock.table('LOG_Sorted_Table').rows;
  const rows = nonBlank(log);
  assert.equal(rows.at(-1)[6], 'Same day A', 'the corrected row is now last in LOG, so table order and number order differ');
  assert.ok(rows.at(-1)[0] < rows.at(-2)[0], 'and it carries the LOWER number of the two');
  assert.ok(scriptVerifySorted(log, sorted), 'the scripts accept LOG_Sorted; ordering it by table position would make them refuse to run');
  assert.deepEqual(scriptTransactions(sorted), scriptOrdered(log));
});

test('Office Scripts: every row this app writes passes the scripts\' own validateLog, after an add, a correction and a delete', async () => {
  const { mock, wb } = setup();
  const snap = await wb.load();
  const categories = mock.table('LISTS_Categories').rows;
  assert.doesNotThrow(() => scriptValidateLog(mock.table('LOG_Table').rows, categories, snap.year), 'the starting workbook itself passes');
  await wb.addTransaction(entry({ date: '2026-07-15', amount: '12.34', description: 'Written by the web app', notes: 'a note the scripts never write' }), { marker: makeMarker() });
  await wb.updateTransaction((await wb.load()).transactions.find(t => t.id === 5), { amount: '380.50' });
  await wb.deleteTransaction((await wb.load()).transactions.find(t => t.id === 13));
  assert.doesNotThrow(() => scriptValidateLog(mock.table('LOG_Table').rows, categories, snap.year));
  assert.ok(scriptVerifySorted(mock.table('LOG_Table').rows, mock.table('LOG_Sorted_Table').rows));
});

// NOTE: this exercises the MOCK's round trip, not Excel's. It shows the app's own logic copes with a whole-table
// rewrite; it says nothing about what the real service does to a nine-digit timestamp. That is the separately
// authorised live check in docs/checkpoint-2-runbook.md.
test('Office Scripts (mock only): a nine-digit operation id survives a simulated whole-table rewrite and the app still finds its row', async () => {
  const { mock, wb } = setup();
  await wb.load();
  const marker = makeMarker();
  const added = await wb.addTransaction(entry({ description: 'Row the scripts will rewrite' }), { marker });
  // What SubmitEntry / DeleteTransaction do to every row: read the whole body, then write it back.
  const log = mock.table('LOG_Table'), sorted = mock.table('LOG_Sorted_Table');
  const roundTrip = (table, rows) => rows.forEach((r, i) => r.forEach((v, c) => table.setCell(i, c, v)));
  roundTrip(log, log.rows.map(r => [...r]));
  roundTrip(sorted, sorted.rows.map(r => [...r]));
  const back = await wb.findAppended(marker);
  assert.ok(back, 'the row is still identifiable by its operation id after the rewrite');
  assert.equal(back.id, added.txn.id);
  assert.equal(back.timestamp, marker);
  assert.match(String(back.timestamp), /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{9}Z$/);
});

test('V1: one shared rule orders LOG_Sorted, so the rebuild and the read-only Diagnostics check cannot disagree', async () => {
  const { mock, wb } = setup();
  await wb.load();
  await wb.addTransaction(entry({ date: '2026-05-04', description: 'Same day A' }), { marker: makeMarker() });
  await wb.addTransaction(entry({ date: '2026-05-04', description: 'Same day B' }), { marker: makeMarker() });
  await wb.updateTransaction((await wb.load()).transactions.find(t => t.description === 'Same day A'), { amount: '77' });
  const log = mock.table('LOG_Table').rows, sorted = mock.table('LOG_Sorted_Table').rows;
  assert.equal(sortedTableMatches(log, sorted).ok, true, 'the check accepts exactly what the rebuild wrote');
  assert.ok(scriptVerifySorted(log, sorted), 'and the workbook scripts accept it too');
  const sameDay = expectedSortedRows(log).filter(r => String(r[6]).startsWith('Same day'));
  assert.deepEqual(sameDay.map(r => r[0]), [14, 15], 'the lower number comes first on a shared date');
  assert.equal(sameDay[0][6], 'Same day A');
  const dateOnly = nonBlank(log).slice().sort((a, b) => a[2] - b[2]);   // the old rule: date only, ties in table order
  assert.equal(sortedTableMatches(log, dateOnly).ok, false, 'the check rejects date-only order');
  assert.equal(scriptVerifySorted(log, dateOnly), false, 'as would the workbook scripts');
  assert.equal(sortedTableMatches(log, []).ok, false);
  assert.deepEqual(expectedSortedRows([]), []);
  assert.deepEqual(expectedSortedRows(null), []);
});

test('V1: a row with no transaction number sorts last within its date and keeps table order', () => {
  const day = 46376, other = 46377;
  const rows = [
    ['', '', '', '', '', '', '', '', ''],
    [7, 'ts-7', day, 'Deposit', 'Membership', 1, 'has a number', '', ''],
    ['', 'ts-blank-id', day, 'Deposit', 'Membership', 1, 'no number', '', ''],
    [2, 'ts-2', other, 'Deposit', 'Membership', 1, 'later date', '', ''],
  ];
  const want = expectedSortedRows(rows);
  assert.equal(want.length, 3, 'the blank row is dropped');
  assert.deepEqual(want.map(r => r[6]), ['has a number', 'no number', 'later date']);
});
