import { test } from 'node:test';
import assert from 'node:assert/strict';
import { MockWorkbook } from '../../site/js/workbook/mock-excel.js';

import { SAMPLE_WORKBOOK as fixture } from '../../site/js/workbook/sample-workbook.js';
const BASE = 'https://graph.microsoft.com/v1.0/drives/D1/items/I1';
const json = async (res) => (res.status === 204 ? null : res.json());
const call = (mock, method, path, body, headers = {}) => mock.fetch(BASE + path, {
  method, headers: { Authorization: 'Bearer t', 'Content-Type': 'application/json', ...headers }, body: body === undefined ? undefined : JSON.stringify(body),
});

test('sessions: create, use, close; a stale session id is rejected as invalidSessionReCreatable', async () => {
  const mock = new MockWorkbook(fixture);
  const res = await call(mock, 'POST', '/workbook/createSession', { persistChanges: true });
  assert.equal(res.status, 201);
  const { id } = await res.json();
  assert.ok(id);
  const ok = await call(mock, 'GET', "/workbook/worksheets/ConfigHidden/range(address='A1')", undefined, { 'workbook-session-id': id });
  assert.equal(ok.status, 200);
  assert.deepEqual((await ok.json()).values, [[2026]]);
  mock.expireSession();
  const stale = await call(mock, 'GET', "/workbook/worksheets/ConfigHidden/range(address='A1')", undefined, { 'workbook-session-id': id });
  assert.equal(stale.status, 404);
  const err = (await stale.json()).error;
  assert.equal(err.innerError.code, 'invalidSessionReCreatable');
  const closed = await call(mock, 'POST', '/workbook/closeSession', {}, { 'workbook-session-id': id });
  assert.equal(closed.status, 204);
});

test('table body range: 26 rows with a blank first row, dates as serials, text formatted', async () => {
  const mock = new MockWorkbook(fixture);
  const res = await call(mock, 'GET', '/workbook/tables/LOG_Table/dataBodyRange');
  assert.equal(res.status, 200);
  const range = await res.json();
  assert.equal(range.address, 'LOG!A2:I14');
  assert.equal(range.rowCount, 13);
  assert.equal(range.columnCount, 9);
  assert.deepEqual(range.values[0], ['', '', '', '', '', '', '', '', '']);
  assert.equal(range.values[1][2], 46027);
  assert.equal(range.text[1][2], '2026-01-05');
  assert.equal(range.text[1][5], '$100.00');
  assert.equal(range.numberFormat[1][2], 'yyyy-mm-dd');
});

test('rows/add appends without inheriting number formats and returns the new index', async () => {
  const mock = new MockWorkbook(fixture);
  const res = await call(mock, 'POST', '/workbook/tables/LOG_Table/rows/add', { values: [[30, '2026-09-06T01:00:00.000Z', 46271, 'Deposit', 'Membership', 25, 'Test', '', '']] });
  assert.equal(res.status, 200);
  const body = await res.json();
  assert.equal(body.index, 13);
  const range = await json(await call(mock, 'GET', '/workbook/tables/LOG_Table/dataBodyRange'));
  assert.equal(range.rowCount, 14);
  assert.equal(range.values[13][0], 30);
  assert.equal(range.numberFormat[13][2], 'General', 'new rows do not inherit formats in the mock');
  const rowRange = await json(await call(mock, 'PATCH', '/workbook/tables/LOG_Table/rows/itemAt(index=13)/range', { numberFormat: [[null, null, 'yyyy-mm-dd', null, null, '$#,##0.00', null, null, null]] }));
  assert.equal(rowRange.numberFormat[0][2], 'yyyy-mm-dd');
  assert.equal(rowRange.text[0][2], '2026-09-06');
});

test('deleting a row by index shifts later rows up and shrinks the table', async () => {
  const mock = new MockWorkbook(fixture);
  const before = await json(await call(mock, 'GET', '/workbook/tables/LOG_Table/dataBodyRange'));
  const thirdId = before.values[3][0];
  const res = await call(mock, 'DELETE', '/workbook/tables/LOG_Table/rows/2');
  assert.equal(res.status, 200);
  const after = await json(await call(mock, 'GET', '/workbook/tables/LOG_Table/dataBodyRange'));
  assert.equal(after.rowCount, 12);
  assert.equal(after.values[2][0], thirdId);
  const alt = await call(mock, 'DELETE', '/workbook/tables/LOG_Table/rows/itemAt(index=0)');
  assert.equal(alt.status, 200);
  assert.equal((await json(await call(mock, 'GET', '/workbook/tables/LOG_Table/dataBodyRange'))).rowCount, 11);
});

test('range PATCH writes plain cells, skips nulls, and refuses formula cells (protected sheets)', async () => {
  const mock = new MockWorkbook(fixture);
  const r = await call(mock, 'PATCH', "/workbook/worksheets/ENTRY/range(address='B23')", { values: [[100]] });
  assert.equal(r.status, 200);
  assert.deepEqual((await json(await call(mock, 'GET', "/workbook/worksheets/ENTRY/range(address='B23')"))).values, [[100]]);
  const partial = await call(mock, 'PATCH', "/workbook/worksheets/LOG/range(address='A3:B3')", { values: [[null, 'kept-id']] });
  assert.equal(partial.status, 200);
  const row = await json(await call(mock, 'GET', "/workbook/worksheets/LOG/range(address='A3:B3')"));
  assert.equal(row.values[0][0], 1, 'null leaves the cell alone');
  assert.equal(row.values[0][1], 'kept-id');
  const bad = await call(mock, 'PATCH', "/workbook/worksheets/January/range(address='C34')", { values: [[0]] });
  assert.equal(bad.status, 403);
  assert.equal((await bad.json()).error.innerError.code, 'accessDenied');
  const outside = await call(mock, 'PATCH', "/workbook/worksheets/LOG_Sorted/range(address='A40:I40')", { values: [[1, 2, 3, 4, 5, 6, 7, 8, 9]] });
  assert.equal(outside.status, 400, 'writing outside a table body on a data sheet is refused so the adapter must use rows/add');
});

test('month and annual sheets are computed from LOG and change after a row is added', async () => {
  const mock = new MockWorkbook(fixture);
  const jan = await json(await call(mock, 'GET', "/workbook/worksheets/January/range(address='A2:O40')"));
  assert.equal(jan.values[0][0], 'January 2026 Transactions');
  assert.equal(jan.values[32][0], 'SUBTOTAL');
  assert.equal(jan.values[32][2], 180.25);
  assert.equal(jan.values[38][1], 5019.75, 'closing balance');
  await call(mock, 'POST', '/workbook/tables/LOG_Table/rows/add', { values: [[30, 'x', 46023, 'Deposit', 'Membership', 100, 'Jan dep', '', '']] });
  const jan2 = await json(await call(mock, 'GET', "/workbook/worksheets/January/range(address='A2:O40')"));
  assert.equal(jan2.values[32][5], 200, 'Membership subtotal now 200');
  assert.equal(jan2.values[38][1], 5119.75);
  const annual = await json(await call(mock, 'GET', "/workbook/worksheets/Annual%202026/range(address='A3:N19')"));
  assert.equal(annual.values[0][0], 'Category');
  assert.equal(annual.values[13][0], 'Net');
});

test('failNext injects one failure then recovers; afterApply applies the write first', async () => {
  const mock = new MockWorkbook(fixture);
  mock.failNext({ match: /rows\/add/, status: 429, code: 'TooManyRequests', innerCode: 'tooManyRequestsUncategorized', retryAfter: 3 });
  const r1 = await call(mock, 'POST', '/workbook/tables/LOG_Table/rows/add', { values: [[30, 'a', 46271, 'Deposit', 'Membership', 1, 'x', '', '']] });
  assert.equal(r1.status, 429);
  assert.equal(r1.headers.get('Retry-After'), '3');
  const r2 = await call(mock, 'POST', '/workbook/tables/LOG_Table/rows/add', { values: [[30, 'a', 46271, 'Deposit', 'Membership', 1, 'x', '', '']] });
  assert.equal(r2.status, 200);
  mock.failNext({ match: /rows\/add/, networkError: true, afterApply: true });
  await assert.rejects(() => call(mock, 'POST', '/workbook/tables/LOG_Table/rows/add', { values: [[31, 'b', 46271, 'Deposit', 'Membership', 1, 'y', '', '']] }), TypeError);
  const range = await json(await call(mock, 'GET', '/workbook/tables/LOG_Table/dataBodyRange'));
  assert.equal(range.rowCount, 15, 'both adds applied (the second despite the network error)');
});

test('log records every request; snapshot equality detects any change', async () => {
  const mock = new MockWorkbook(fixture);
  const snap = mock.snapshot();
  await call(mock, 'GET', '/workbook/tables/LOG_Table/dataBodyRange');
  assert.equal(mock.log.length, 1);
  assert.equal(mock.log[0].method, 'GET');
  assert.deepEqual(mock.snapshot(), snap);
  await call(mock, 'PATCH', "/workbook/worksheets/ENTRY/range(address='B23')", { values: [[1]] });
  assert.notDeepEqual(mock.snapshot(), snap);
  const meta = await json(await call(mock, 'GET', '?$select=name,size,lastModifiedDateTime,webUrl,eTag,file'));
  assert.match(meta.name, /\.xlsx$/);
});

test('$batch runs sub-requests in order, honours dependsOn failures with 424, and fires beforeRespond for each', async () => {
  const mock = new MockWorkbook(fixture);
  const seen = [];
  mock.beforeRespond = async e => { seen.push(e.method); };
  const res = await mock.fetch('https://graph.microsoft.com/v1.0/$batch', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ requests: [
    { id: '1', method: 'GET', url: '/drives/D1/items/I1/workbook/tables/LOG_Table/dataBodyRange' },
    { id: '2', method: 'DELETE', url: '/drives/D1/items/I1/workbook/tables/LOG_Table/rows/99', dependsOn: ['1'] },
    { id: '3', method: 'DELETE', url: '/drives/D1/items/I1/workbook/tables/LOG_Table/rows/1', dependsOn: ['2'] },
  ] }) });
  assert.equal(res.status, 200);
  const body = await res.json();
  assert.deepEqual(body.responses.map(r => r.status), [200, 400, 424]);
  assert.equal(body.responses[0].body.rowCount, 13);
  assert.deepEqual(seen, ['POST', 'GET', 'DELETE']);
  assert.equal(mock.table('LOG_Table').rows.length, 13, 'the dependent delete did not run');
});

test('exportState/importState reproduce the workbook (test-mode persistence across a page reload); onChange fires on mutations', async () => {
  const mock = new MockWorkbook(fixture);
  let changes = 0;
  mock.onChange = () => { changes++; };
  mock.table('LOG_Table').append([14, 'persist', 46271, 'Deposit', 'Membership', 5, 'Persisted row', '', '']);
  mock.setCell('ENTRY', 'B23', 6000);
  assert.equal(changes, 2);
  const state = JSON.parse(JSON.stringify(mock.exportState()));
  const copy = new MockWorkbook(fixture).importState(state);
  assert.deepEqual(copy.table('LOG_Table').rows, mock.table('LOG_Table').rows);
  assert.equal(copy.getCell('ENTRY', 'B23'), 6000);
  const body = await json(await call(copy, 'GET', '/workbook/tables/LOG_Table/dataBodyRange'));
  assert.equal(body.rowCount, 14);
  assert.equal(body.values.at(-1)[6], 'Persisted row');
  const month = await json(await call(copy, 'GET', "/workbook/worksheets/January/range(address='B37:B37')"));
  assert.equal(month.values[0][0], 6000, 'computed sheets follow the imported cells');
});
