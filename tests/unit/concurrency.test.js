// Two independent clients/adapters (two members) against one shared mock workbook, plus direct "Excel" edits.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { MockWorkbook } from '../../site/js/workbook/mock-excel.js';
import { ExcelClient } from '../../site/js/workbook/excel-client.js';
import { LedgerWorkbook, UnresolvedOperationError, ConflictError, makeMarker } from '../../site/js/workbook/ledger-workbook.js';
import { SAMPLE_WORKBOOK as fixture } from '../../site/js/workbook/sample-workbook.js';

function member(mock) {
  const client = new ExcelClient({ getToken: async () => 't', driveId: 'D1', itemId: 'I1', fetchImpl: (u, o) => mock.fetch(u, o), sleep: async () => {} });
  return new LedgerWorkbook(client);
}
const nonBlank = rows => rows.filter(r => r.some(v => v !== '' && v !== null));
const entry = (over = {}) => ({ date: '2026-09-06', type: 'Deposit', category: 'Membership', amount: '25.50', description: 'Test add', chequeNum: '', notes: '', ...over });

test('two members add at the same moment: both rows land with unique numbers and both are told the truth', async () => {
  const mock = new MockWorkbook(fixture);
  const A = member(mock), B = member(mock);
  await A.load(); await B.load();
  let interleaved = false;
  let bResult;
  mock.beforeRespond = async e => {
    if (!interleaved && e.method === 'POST' && /rows\/add/.test(e.url) && e.sessionId === A.client.sessionId) {
      interleaved = true;
      bResult = await B.addTransaction(entry({ description: 'B adds' }), { marker: makeMarker() });   // B completes inside A's write window
    }
  };
  const aResult = await A.addTransaction(entry({ description: 'A adds' }), { marker: makeMarker() });
  const rows = nonBlank(mock.table('LOG_Table').rows);
  assert.equal(rows.length, 14);
  const ids = rows.map(r => r[0]);
  assert.equal(new Set(ids).size, ids.length, 'unique numbers');
  assert.ok(rows.some(r => r[6] === 'A adds') && rows.some(r => r[6] === 'B adds'));
  assert.equal(bResult.txn.id, 14);
  assert.equal(aResult.txn.id, 15);
  assert.equal(aResult.renumberedFrom, 14);
});

test('A deletes #13 while B corrects #2 (which sits above #13): B succeeds; A either succeeds or reports the exact row its delete hit', async () => {
  const mock = new MockWorkbook(fixture);
  const A = member(mock), B = member(mock);
  const snapA = await A.load(); const snapB = await B.load();
  const target = snapA.transactions.find(t => t.id === 13);
  const refB = snapB.transactions.find(t => t.id === 2);
  const before = nonBlank(mock.table('LOG_Table').rows).map(r => `${r[0]}|${r[1]}`);
  let interleaved = false;
  mock.beforeRespond = async e => {
    if (!interleaved && e.method === 'DELETE' && e.sessionId === A.client.sessionId) {
      interleaved = true;
      await B.updateTransaction(refB, { description: 'B corrected #2' });   // moves #2 to the end, shifting rows above #13 up by one
    }
  };
  let aErr = null;
  try { await A.deleteTransaction(target); } catch (e) { aErr = e; }
  let rows = nonBlank(mock.table('LOG_Table').rows);
  const after = rows.map(r => `${r[0]}|${r[1]}`);
  if (aErr) {
    // B's correction moved #2 to the end of the table and shifted #13 up; A's delete hit whatever now sat at #13's old index.
    assert.ok(aErr instanceof UnresolvedOperationError, String(aErr));
    assert.equal(aErr.kind, 'wrong-row-deleted');
    assert.ok(rows.some(r => r[0] === 13), 'A\'s target was not deleted');
    const lost = before.filter(x => !after.includes(x));
    const victimKey = `${aErr.details.victim.id}|${aErr.details.victim.timestamp}`;
    assert.ok(lost.length <= 1 && (lost.length === 0 || lost[0] === victimKey), 'no unexplained loss: any missing row is the one named in the report');
    // The report can only carry the values THIS device last saw for the victim; B's newer text was never seen by A.
    assert.equal(aErr.details.victimLastKnownOnly, true);
    if (!rows.some(r => `${r[0]}|${r[1]}` === victimKey)) {
      await A.reappendRow(aErr.details.victim.values, { confirm: true });   // explicit, member-confirmed
      rows = nonBlank(mock.table('LOG_Table').rows);
    }
    assert.equal(rows.filter(r => r[0] === 2).length, 1, 'exactly one #2 after the explicit repair');
  } else {
    assert.equal(rows.filter(r => r[0] === 2).length, 1);
    assert.equal(rows.find(r => r[0] === 2)[6], 'B corrected #2', 'B\'s correction survived');
    assert.ok(!rows.some(r => r[0] === 13));
    const lost = before.filter(x => !after.includes(x));
    assert.deepEqual(lost, [`13|${target.timestamp}`]);
  }
});

test('a change made directly in Excel between loading and deleting is a conflict, and appears after reload', async () => {
  const mock = new MockWorkbook(fixture);
  const A = member(mock);
  const snap = await A.load();
  const target = snap.transactions.find(t => t.id === 7);
  mock.table('LOG_Table').setCell(target.rowIndex, 5, 121);   // typed in Excel
  await assert.rejects(() => A.deleteTransaction(target), e => e instanceof ConflictError && e.current.amount === 121);
  const fresh = (await A.load()).transactions.find(t => t.id === 7);
  assert.equal(fresh.amount, 121);
  await A.deleteTransaction(fresh);
  assert.ok(!mock.table('LOG_Table').rows.some(r => r[0] === 7));
});

test('two members correcting the same record: the second sees a conflict instead of overwriting', async () => {
  const mock = new MockWorkbook(fixture);
  const A = member(mock), B = member(mock);
  const refA = (await A.load()).transactions.find(t => t.id === 5);
  const refB = (await B.load()).transactions.find(t => t.id === 5);
  await A.updateTransaction(refA, { amount: 380 });
  await assert.rejects(() => B.updateTransaction(refB, { amount: 390 }), e => e instanceof ConflictError && e.current.amount === 380);
  assert.equal(mock.table('LOG_Table').rows.filter(r => r[0] === 5).length, 1);
  assert.equal(mock.table('LOG_Table').rows.find(r => r[0] === 5)[5], 380);
});

test('the sorted helper table ends consistent with LOG after both members finish', async () => {
  const mock = new MockWorkbook(fixture);
  const A = member(mock), B = member(mock);
  await A.load(); await B.load();
  await A.addTransaction(entry({ description: 'A1', date: '2026-02-10' }), { marker: makeMarker() });
  await B.addTransaction(entry({ description: 'B1', date: '2026-01-02' }), { marker: makeMarker() });
  const r = await A.rebuildSorted();
  assert.equal(r.consistent, true);
  const log = nonBlank(mock.table('LOG_Table').rows).slice().sort((a, b) => a[2] - b[2]).map(r => r[1]);
  assert.deepEqual(mock.table('LOG_Sorted_Table').rows.map(r => r[1]), log);
});
