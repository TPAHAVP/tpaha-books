// The published ledger page opens one workbook, the one pinned in config.js, for every member on every device.
// A workbook remembered earlier on the device (a test copy chosen through the old picker) is ignored and
// forgotten, and an access failure is reported as such: the page never searches for other copies.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { startupWorkbook, describeAccessFailure, rememberWorkbook, recallWorkbook } from '../../site/js/workbook/locate.js';

class MemStorage {
  constructor() { this.m = new Map(); }
  getItem(k) { return this.m.has(k) ? this.m.get(k) : null; }
  setItem(k, v) { this.m.set(k, String(v)); }
  removeItem(k) { this.m.delete(k); }
}
const PINNED = { name: 'TPAHA_2026.xlsx', driveId: 'b!DRIVE', itemId: '01ITEM' };
const config = ledger => ({ workbooks: { ledger, storage: { driveId: '', itemId: '', name: '' } } });
const who = { tenantId: 'T', homeAccountId: 'U1.T', app: 'ledger' };

test('fresh device and account: the pinned workbook is used, from the configuration', () => {
  const storage = new MemStorage();
  const r = startupWorkbook({ config: config(PINNED), storage, who });
  assert.deepEqual(r.workbook, { driveId: 'b!DRIVE', itemId: '01ITEM', name: 'TPAHA_2026.xlsx', source: 'config' });
  assert.equal(r.replaced, null);
  assert.equal(storage.m.size, 0, 'nothing is written to the device');
});

test('device remembering a test copy: the pinned workbook overrides it and the remembered choice is forgotten', () => {
  const storage = new MemStorage();
  rememberWorkbook(storage, who, { driveId: 'b!OTHER', itemId: '01COPY', name: 'TPAHA_2026 (1).xlsx' });
  const r = startupWorkbook({ config: config(PINNED), storage, who });
  assert.equal(r.workbook.itemId, '01ITEM');
  assert.equal(r.workbook.source, 'config');
  assert.equal(r.replaced.name, 'TPAHA_2026 (1).xlsx', 'the page can say what was ignored');
  assert.equal(recallWorkbook(storage, who), null, 'the test-copy selection is removed, so it cannot come back');
});

test('the remembered choice of another account on the same device is left alone', () => {
  const storage = new MemStorage();
  const other = { ...who, homeAccountId: 'U2.T' };
  rememberWorkbook(storage, other, { driveId: 'b!OTHER', itemId: '01COPY', name: 'copy.xlsx' });
  startupWorkbook({ config: config(PINNED), storage, who });
  assert.ok(recallWorkbook(storage, other), 'only the signed-in account\'s own record is touched');
});

test('no pinned workbook: nothing is opened and nothing is remembered or searched', () => {
  const storage = new MemStorage();
  rememberWorkbook(storage, who, { driveId: 'b!OTHER', itemId: '01COPY', name: 'copy.xlsx' });
  const r = startupWorkbook({ config: config({ driveId: '', itemId: '', name: '' }), storage, who });
  assert.equal(r.workbook, null, 'a remembered copy is never used in place of the configured workbook');
  assert.equal(r.replaced, null);
});

// 2026-11-05: the secretary could open the workbook in Excel yet the page got 403 accessDenied. A 403 from Graph
// can mean the site's sign-in permission does not reach a workbook in another member's OneDrive, not only that
// the workbook is unshared, so the wording must not assert that the member lacks sharing permission.
test('describeAccessFailure: 403 names the refusal and both possible causes, never asserting the member is unshared', () => {
  const t = describeAccessFailure({ status: 403, code: 'AccessDenied', innerCode: 'accessDenied', message: 'Access denied', requestId: 'req-1' }, 'TPAHA_2026.xlsx');
  assert.match(t.title, /Could not open TPAHA_2026\.xlsx/);
  assert.match(t.reason, /refused|access denied/i, 'states what Microsoft 365 did');
  assert.match(t.reason, /permission/i, 'names the site sign-in permission as a possible cause');
  assert.match(t.reason, /Excel/, 'says this can happen even when Excel opens the file');
  assert.match(t.reason, /shared/i, 'sharing remains a possible cause');
  assert.doesNotMatch(t.reason, /does not have access|has been withdrawn|Ask the treasurer to share/i, 'no assertion that the member lacks sharing permission');
  assert.match(t.detail, /403/);
  assert.match(t.detail, /accessDenied/);
  assert.match(t.detail, /req-1/);
  assert.match(t.scope, /does not search for other copies/i);
  assert.doesNotMatch(JSON.stringify(t), /b!|01ITEM/, 'no drive or item id in the wording');
  // The same wording when only the code says access was denied.
  assert.equal(describeAccessFailure({ code: 'accessDenied', message: 'Access denied' }, 'TPAHA_2026.xlsx').reason, t.reason);
});

test('describeAccessFailure: 404 says the configured workbook was not found; other failures state status and code', () => {
  assert.match(describeAccessFailure({ status: 404, code: 'itemNotFound', message: 'The resource could not be found.' }, 'TPAHA_2026.xlsx').reason, /not found|moved|renamed|deleted/i);
  const t = describeAccessFailure({ status: 503, code: 'ServiceUnavailable', innerCode: 'serviceUnavailable', message: 'Try later' }, 'TPAHA_2026.xlsx');
  assert.match(t.reason, /Try later/);
  assert.match(t.detail, /503/);
  const n = describeAccessFailure({ status: 0, message: 'Failed to fetch' }, 'TPAHA_2026.xlsx');
  assert.match(n.reason, /could not be reached|Failed to fetch/i);
});
