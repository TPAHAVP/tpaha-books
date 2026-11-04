// The request log must hold every outcome a member may need to report, including the ones that happen before
// a session exists (session creation, asynchronous polling) and the ones hidden inside a JSON batch. It must
// never hold a token, a session id, a polling URL that carries an operation/session key, or cell contents.
// Review finding of 2026-11-03: a createSession failure left the log empty.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { MockWorkbook } from '../../site/js/workbook/mock-excel.js';
import { ExcelClient, ExcelApiError, describeLogEntry } from '../../site/js/workbook/excel-client.js';
import { SAMPLE_WORKBOOK as fixture } from '../../site/js/workbook/sample-workbook.js';

const makeClient = (mock, log) => new ExcelClient({ getToken: async () => 'tok-secret', driveId: 'D1', itemId: 'I1', fetchImpl: (u, o) => mock.fetch(u, o), sleep: async () => {}, log: e => log.push(e) });
const lines = log => log.map(describeLogEntry);
const assertClean = log => {
  const text = JSON.stringify(log) + '\n' + lines(log).join('\n');
  assert.doesNotMatch(text, /tok-secret/, 'no token');
  assert.doesNotMatch(text, /session-\d+/, 'no session id');
  assert.doesNotMatch(text, /op\d+|key='/, 'no operation key or session resource key');
  assert.doesNotMatch(text, /sessions?\/[A-Za-z0-9]/, 'no polling URL with an id');
};

test('a refused createSession is logged with status, request id and error codes', async () => {
  const mock = new MockWorkbook(fixture);
  const log = [];
  const client = makeClient(mock, log);
  mock.failNext({ match: 'createSession', status: 403, code: 'AccessDenied', innerCode: ['accessDenied'], message: 'Access denied' });
  await assert.rejects(() => client.openSession(), e => e instanceof ExcelApiError && e.status === 403 && e.code === 'AccessDenied' && e.innerCode === 'accessDenied');
  const entry = log.find(e => e.path === '/workbook/createSession');
  assert.ok(entry, `createSession must appear in the log; got ${JSON.stringify(log)}`);
  assert.equal(entry.method, 'POST');
  assert.equal(entry.status, 403);
  assert.ok(entry.requestId, 'request id');
  assert.equal(entry.code, 'AccessDenied');
  assert.deepEqual(entry.innerCodes, ['accessDenied']);
  assert.equal(entry.error, 'Access denied');
  assert.match(describeLogEntry(entry), /^POST \/workbook\/createSession → 403 .*request-id .* code AccessDenied inner accessDenied — Access denied$/);
  assertClean(log);
});

test('a createSession that cannot reach the service is logged as a network error for each attempt', async () => {
  const mock = new MockWorkbook(fixture);
  const log = [];
  const client = makeClient(mock, log);
  mock.failNext({ match: 'createSession', networkError: true });
  mock.failNext({ match: 'createSession', networkError: true });
  await assert.rejects(() => client.openSession(), e => e instanceof ExcelApiError && e.status === 0);
  const entries = log.filter(e => e.path === '/workbook/createSession');
  assert.equal(entries.length, 2, 'one attempt and one retry, both recorded');
  assert.ok(entries.every(e => e.status === 0 && /Failed to fetch/.test(e.error)));
  assert.match(describeLogEntry(entries[0]), /^POST \/workbook\/createSession → network error/);
  assertClean(log);
});

test('a 504 on createSession is logged, then the one retry that opens the session', async () => {
  const mock = new MockWorkbook(fixture);
  const log = [];
  const client = makeClient(mock, log);
  mock.failNext({ match: 'createSession', status: 504, code: 'GatewayTimeout', innerCode: ['gatewayTimeout'], message: 'Gateway timeout' });
  await client.openSession();
  assert.deepEqual(log.map(e => e.status), [504, 201]);
  assert.equal(log[0].code, 'GatewayTimeout');
  assert.equal(log[1].event, 'session');
  assert.ok(log[1].requestId, 'the successful creation carries its request id');
  assert.match(describeLogEntry(log[1]), /^session opened \(201, \d+ ms\) request-id /);
  assertClean(log);
});

test('asynchronous session creation logs the 202, every poll with its outcome and the final read, without the operation key', async () => {
  const mock = new MockWorkbook(fixture);
  mock.slowSessionCreation = 1;
  const log = [];
  const client = makeClient(mock, log);
  await client.openSession();
  const statuses = log.map(e => `${e.event || e.method} ${e.path || ''} ${e.status} ${e.outcome || ''}`.replace(/\s+/g, ' ').trim());
  assert.deepEqual(statuses, [
    'POST /workbook/createSession 202',
    'GET /workbook/operations/… 200 running',
    'GET /workbook/operations/… 200 succeeded',
    'GET /workbook/sessionInfoResource(…) 200',
    'session 201',
  ], JSON.stringify(log));
  assert.ok(log.every(e => e.requestId), 'every step, the 202 included, carries its request id');
  assert.match(describeLogEntry(log[1]), /^GET \/workbook\/operations\/… → 200 \(\d+ ms\) request-id \S+ operation running$/);
  assertClean(log);
});

test('a session creation the service reports as failed is logged with the operation error codes', async () => {
  const mock = new MockWorkbook(fixture);
  mock.slowSessionCreation = 0;
  mock.sessionCreationOutcome = { status: 'failed', code: 'UnsupportedWorkbook', innerCode: ['unsupportedWorkbook'], message: 'The workbook cannot be opened' };
  mock.slowSessionCreation = 1;
  const log = [];
  const client = makeClient(mock, log);
  await assert.rejects(() => client.openSession(), e => e instanceof ExcelApiError && e.code === 'UnsupportedWorkbook' && e.innerCode === 'unsupportedWorkbook');
  const failed = log.find(e => e.outcome === 'failed');
  assert.ok(failed, JSON.stringify(log));
  assert.equal(failed.path, '/workbook/operations/…');
  assert.equal(failed.code, 'UnsupportedWorkbook');
  assert.deepEqual(failed.innerCodes, ['unsupportedWorkbook']);
  assert.equal(failed.error, 'The workbook cannot be opened');
  assert.match(describeLogEntry(failed), /operation failed code UnsupportedWorkbook inner unsupportedWorkbook — The workbook cannot be opened$/);
  assertClean(log);
});

test('a poll that is itself refused is logged with its status and codes', async () => {
  const mock = new MockWorkbook(fixture);
  mock.slowSessionCreation = 1;
  const log = [];
  const client = makeClient(mock, log);
  mock.failNext({ match: '/workbook/operations/', status: 503, code: 'ServiceUnavailable', innerCode: ['serviceUnavailable'], message: 'Try later' });
  await assert.rejects(() => client.openSession(), e => e instanceof ExcelApiError && e.status === 503 && e.code === 'ServiceUnavailable');
  const poll = log.find(e => e.path === '/workbook/operations/…');
  assert.ok(poll, JSON.stringify(log));
  assert.equal(poll.status, 503);
  assert.equal(poll.code, 'ServiceUnavailable');
  assert.deepEqual(poll.innerCodes, ['serviceUnavailable']);
  assertClean(log);
});

test('a batch refused as a whole keeps its error codes in the log, not only the outer status', async () => {
  const mock = new MockWorkbook(fixture);
  const log = [];
  const client = makeClient(mock, log);
  mock.failNext({ match: /\$batch$/, status: 500, code: 'InternalServerError', innerCode: ['internalServerErrorUncategorized'], message: 'Something went wrong' });
  await assert.rejects(() => client.batch([{ path: '/workbook/tables/LOG_Table/dataBodyRange' }]), e => e instanceof ExcelApiError && e.status === 500 && e.code === 'InternalServerError' && e.innerCode === 'internalServerErrorUncategorized');
  const entry = log.find(e => e.path === '/$batch');
  assert.equal(entry.status, 500);
  assert.equal(entry.code, 'InternalServerError');
  assert.deepEqual(entry.innerCodes, ['internalServerErrorUncategorized']);
  assert.equal(entry.error, 'Something went wrong');
  assertClean(log);
});

test('each sub-request of a batch is logged with its own status and, when refused, its error codes', async () => {
  const mock = new MockWorkbook(fixture);
  const log = [];
  const client = makeClient(mock, log);
  const results = await client.batch([
    { path: '/workbook/tables/LOG_Table/dataBodyRange?$select=rowCount' },
    { path: '/workbook/tables/NoSuchTable/dataBodyRange' },
    { path: '/workbook/tables/LOG_Table/dataBodyRange?$select=address' },
  ]);
  assert.equal(results[0].status, 200);
  assert.equal(results[1].status, 404);
  assert.equal(results[2].status, 424, 'depends on the failed one');
  const subs = log.filter(e => e.batch);
  assert.equal(subs.length, 3, JSON.stringify(log));
  assert.deepEqual(subs.map(e => [e.method, e.status]), [['GET', 200], ['GET', 404], ['GET', 424]]);
  assert.equal(subs[0].path, '/workbook/tables/LOG_Table/dataBodyRange?$select=rowCount');
  assert.equal(subs[1].code, 'ItemNotFound');
  assert.ok(subs[1].innerCodes.length >= 1);
  assert.ok(subs[1].error, 'the message is kept');
  assert.equal(subs[2].code, 'FailedDependency');
  assert.match(describeLogEntry(subs[1]), /^ {2}in batch: GET \/workbook\/tables\/NoSuchTable\/dataBodyRange → 404 .*code ItemNotFound/);
  const outer = log.find(e => e.path === '/$batch');
  assert.equal(outer.status, 200);
  assert.equal(outer.items, 3);
  assert.match(describeLogEntry(outer), /^POST \/\$batch \(3 requests\) → 200/);
  assertClean(log);
});
