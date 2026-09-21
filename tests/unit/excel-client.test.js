import { test } from 'node:test';
import assert from 'node:assert/strict';
import { MockWorkbook } from '../../site/js/workbook/mock-excel.js';
import { ExcelClient, ExcelApiError, classifyError } from '../../site/js/workbook/excel-client.js';

import { SAMPLE_WORKBOOK as fixture } from '../../site/js/workbook/sample-workbook.js';

function makeClient(mock, extra = {}) {
  const slept = [];
  const client = new ExcelClient({
    getToken: async () => 'tok', driveId: 'D1', itemId: 'I1',
    fetchImpl: (u, o) => mock.fetch(u, o), sleep: async ms => { slept.push(ms); }, ...extra,
  });
  return { client, slept };
}

test('classifyError reads second-level codes and Retry-After', () => {
  const c = classifyError(404, { error: { code: 'ItemNotFound', innerError: { code: 'invalidSessionReCreatable' } } }, new Headers());
  assert.equal(c.innerCode, 'invalidSessionReCreatable');
  assert.equal(c.sessionInvalid, true);
  assert.equal(c.recreatable, true);
  const t = classifyError(429, { error: { code: 'TooManyRequests', innerError: { code: 'tooManyRequestsUncategorized' } } }, new Headers({ 'Retry-After': '7' }));
  assert.equal(t.retryAfter, 7);
  assert.equal(t.retryableRead, true);
  const k = classifyError(409, { error: { code: 'Conflict', innerError: { code: 'accessConflict' } } }, new Headers());
  assert.equal(k.conflict, true);
  assert.equal(k.retryableRead, false);
});

test('opens a session lazily, sends the session header on every call, and closes it', async () => {
  const mock = new MockWorkbook(fixture);
  const { client } = makeClient(mock);
  const range = await client.getRange('ConfigHidden', 'A1');
  assert.deepEqual(range.values, [[2026]]);
  assert.equal(mock.log[0].url.endsWith('/workbook/createSession'), true);
  assert.ok(mock.log[1].sessionId, 'session header present');
  assert.equal(mock.log[1].sessionId, client.sessionId);
  await client.closeSession();
  assert.equal(mock.log.at(-1).url.endsWith('/workbook/closeSession'), true);
  assert.equal(client.sessionId, null);
});

test('handles the 202 long-running session creation by polling', async () => {
  const mock = new MockWorkbook(fixture);
  mock.slowSessionCreation = 2;   // two "running" polls before "succeeded"
  const { client, slept } = makeClient(mock);
  await client.openSession();
  assert.ok(client.sessionId);
  assert.ok(slept.length >= 2, 'polled with sleeps');
  const ops = mock.log.filter(e => /\/workbook\/operations\//.test(e.url));
  assert.equal(ops.length, 3);
});

test('requests are strictly sequential per client', async () => {
  const mock = new MockWorkbook(fixture);
  let release;
  const held = new Promise(r => { release = r; });
  let first = true;
  mock.beforeRespond = async entry => { if (first && /dataBodyRange/.test(entry.url)) { first = false; await held; } };
  const { client } = makeClient(mock);
  await client.openSession();
  const p1 = client.getTableBody('LOG_Table');
  const p2 = client.getRange('ENTRY', 'B23');
  await new Promise(r => setTimeout(r, 10));
  assert.equal(mock.log.filter(e => /ENTRY/.test(e.url)).length, 0, 'second request waits for the first');
  release();
  await Promise.all([p1, p2]);
  assert.equal(mock.log.filter(e => /ENTRY/.test(e.url)).length, 1);
});

test('an expired session on a read is recreated and the read retried once', async () => {
  const mock = new MockWorkbook(fixture);
  const { client } = makeClient(mock);
  await client.openSession();
  const old = client.sessionId;
  mock.expireSession();
  const range = await client.getRange('ENTRY', 'B23');
  assert.deepEqual(range.values, [[5000]]);
  assert.notEqual(client.sessionId, old);
  assert.equal(mock.log.filter(e => /createSession/.test(e.url)).length, 2);
});

test('an expired session on a write is NOT retried; the error says the session was invalid', async () => {
  const mock = new MockWorkbook(fixture);
  const { client } = makeClient(mock);
  await client.openSession();
  mock.expireSession();
  await assert.rejects(() => client.addTableRows('LOG_Table', [[30, 'm', 46271, 'Deposit', 'Membership', 1, 'x', '', '']]), e => e instanceof ExcelApiError && e.sessionInvalid === true && e.ambiguous === false);
  assert.equal(mock.log.filter(e => /rows\/add/.test(e.url)).length, 1, 'exactly one attempt');
  assert.equal((await client.getTableBody('LOG_Table')).rowCount, 13, 'nothing was written');
});

test('429 on a read waits for Retry-After and retries once; on a write it surfaces retryAfter without retrying', async () => {
  const mock = new MockWorkbook(fixture);
  const { client, slept } = makeClient(mock);
  await client.openSession();
  mock.failNext({ match: /dataBodyRange/, status: 429, code: 'TooManyRequests', innerCode: 'tooManyRequestsUncategorized', retryAfter: 2 });
  const body = await client.getTableBody('LOG_Table');
  assert.equal(body.rowCount, 13);
  assert.ok(slept.includes(2000));
  mock.failNext({ match: /rows\/add/, status: 429, code: 'TooManyRequests', innerCode: 'tooManyRequestsUncategorized', retryAfter: 5 });
  await assert.rejects(() => client.addTableRows('LOG_Table', [[30, 'm', 46271, 'Deposit', 'Membership', 1, 'x', '', '']]), e => e instanceof ExcelApiError && e.retryAfter === 5 && e.ambiguous === false);
  assert.equal(mock.log.filter(e => /rows\/add/.test(e.url)).length, 1);
});

test('a network failure during a write is reported as ambiguous; during a read it is retried once', async () => {
  const mock = new MockWorkbook(fixture);
  const { client } = makeClient(mock);
  await client.openSession();
  mock.failNext({ match: /rows\/add/, networkError: true, afterApply: true });
  await assert.rejects(() => client.addTableRows('LOG_Table', [[30, 'm', 46271, 'Deposit', 'Membership', 1, 'x', '', '']]), e => e instanceof ExcelApiError && e.ambiguous === true && e.status === 0);
  mock.failNext({ match: /dataBodyRange/, networkError: true });
  const body = await client.getTableBody('LOG_Table');
  assert.equal(body.rowCount, 14, 'the ambiguous write had applied; the read retried after the network error');
});

test('5xx on a write is ambiguous; 504 on a read is retried once; 409 accessConflict is a conflict', async () => {
  const mock = new MockWorkbook(fixture);
  const { client } = makeClient(mock);
  await client.openSession();
  mock.failNext({ match: /rows\/add/, status: 503, code: 'ServiceUnavailable', innerCode: 'serviceUnavailableUncategorized', retryAfter: 1 });
  await assert.rejects(() => client.addTableRows('LOG_Table', [[30, 'm', 46271, 'Deposit', 'Membership', 1, 'x', '', '']]), e => e.ambiguous === true && e.retryAfter === 1);
  mock.failNext({ match: /ConfigHidden/, status: 504, code: 'GatewayTimeout', innerCode: 'gatewayTimeoutUncategorized' });
  assert.deepEqual((await client.getRange('ConfigHidden', 'A1')).values, [[2026]]);
  mock.failNext({ match: /B23/, status: 409, code: 'Conflict', innerCode: 'accessConflict' });
  await assert.rejects(() => client.patchRange('ENTRY', 'B23', { values: [[1]] }), e => e.conflict === true && e.ambiguous === false);
});

test('item metadata and calculate use the documented paths', async () => {
  const mock = new MockWorkbook(fixture);
  const { client } = makeClient(mock);
  const meta = await client.getItemMeta();
  assert.match(meta.name, /\.xlsx$/);
  await client.calculate('Full');
  const calc = mock.log.find(e => /application\/calculate/.test(e.url));
  assert.equal(calc.method, 'POST');
  assert.deepEqual(JSON.parse(calc.body), { calculationType: 'Full' });
});

test('batch: sequential sub-requests share the session, return per-item results, and a transport failure is ambiguous when the batch mutates', async () => {
  const mock = new MockWorkbook(fixture);
  const { client } = makeClient(mock);
  await client.openSession();
  const results = await client.batch([
    { method: 'GET', path: '/workbook/tables/LOG_Table/dataBodyRange?$select=address,values,rowCount' },
    { method: 'DELETE', path: '/workbook/tables/LOG_Table/rows/$/itemAt(index=1)' },
  ]);
  assert.equal(results.length, 2);
  assert.equal(results[0].status, 200);
  assert.equal(results[0].body.rowCount, 13, 'the read shows the table as it was immediately before the delete');
  assert.equal(results[1].status, 204, 'the published example answers 204 No Content');
  assert.equal((await client.getTableBody('LOG_Table')).rowCount, 12);
  const batchEntry = mock.log.find(e => /\$batch$/.test(e.url));
  assert.ok(batchEntry.sessionId, 'the batch request itself carries the session header');
  mock.failNext({ match: /\$batch$/, networkError: true });
  await assert.rejects(() => client.batch([{ method: 'DELETE', path: '/workbook/tables/LOG_Table/rows/$/itemAt(index=1)' }]), e => e instanceof ExcelApiError && e.ambiguous === true);
  assert.equal(mock.log.filter(e => /\$batch$/.test(e.url)).length, 2, 'a failed batch is never resent');
});

test('writes are never retried automatically, whatever helper is used', async () => {
  const mock = new MockWorkbook(fixture);
  const { client } = makeClient(mock);
  await client.openSession();
  mock.failNext({ match: /worksheets\/ENTRY/, networkError: true, afterApply: true });
  await assert.rejects(() => client.patchRange('ENTRY', 'B23', { values: [[7]] }), e => e.ambiguous === true);
  assert.equal(mock.log.filter(e => e.method === 'PATCH' && /ENTRY/.test(e.url)).length, 1);
  mock.failNext({ match: /rows\/itemAt/, status: 503, code: 'ServiceUnavailable', innerCode: 'serviceUnavailableUncategorized', retryAfter: 1 });
  await assert.rejects(() => client.patchTableRowRange('LOG_Table', 1, { numberFormat: [[null, null, 'yyyy-mm-dd', null, null, null, null, null, null]] }), e => e.ambiguous === true);
  assert.equal(mock.log.filter(e => e.method === 'PATCH' && /rows\/itemAt/.test(e.url)).length, 1);
  mock.failNext({ match: /calculate/, networkError: true });
  await assert.rejects(() => client.calculate(), e => e.status === 0);
  assert.equal(mock.log.filter(e => /calculate/.test(e.url)).length, 1);
});

test('a refused request is logged with its status, code, inner code and message, so the next live failure is recorded rather than reconstructed', async () => {
  const mock = new MockWorkbook(fixture);
  const log = [];
  const { client } = makeClient(mock, { log: e => log.push(e) });
  await client.getTableBody('LISTS_Categories');                                   // one success, for contrast
  mock.failNext({ match: /LOG_Table\/dataBodyRange/, status: 400, code: 'InvalidArgument', innerCode: 'invalidArgument', message: 'The argument is invalid or missing or has an incorrect format.' });
  await assert.rejects(() => client.getTableBody('LOG_Table'));
  const e = log.find(x => x.status === 400);
  assert.ok(e, 'the refusal was logged');
  assert.equal(e.code, 'InvalidArgument');
  assert.equal(e.innerCode, 'invalidArgument');
  assert.match(e.error, /argument is invalid/);
  assert.ok(log.some(x => x.status === 200 && x.code === undefined), 'a success carries no error fields');
});
