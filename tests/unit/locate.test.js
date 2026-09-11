import { test } from 'node:test';
import assert from 'node:assert/strict';
import { findCandidates, workbookKey, rememberWorkbook, recallWorkbook, forgetWorkbook, resolveConfigured } from '../../site/js/workbook/locate.js';

class MemStorage {
  constructor() { this.m = new Map(); }
  getItem(k) { return this.m.has(k) ? this.m.get(k) : null; }
  setItem(k, v) { this.m.set(k, String(v)); }
  removeItem(k) { this.m.delete(k); }
}

const searchResult = {
  value: [
    { id: 'ITEM-XLSX', name: 'TPAHA_2026 (1).xlsx', size: 104054, lastModifiedDateTime: '2026-09-06T01:23:54Z', webUrl: 'https://tenant-my.sharepoint.com/personal/x/Documents/TPAHA_2026%20(1).xlsx', parentReference: { driveId: 'DRIVE-ME', path: '/drive/root:/Documents' }, file: { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' } },
    { id: 'ITEM-CSV', name: 'TPAHA_2026 export.csv', size: 10, lastModifiedDateTime: '2026-09-01T00:00:00Z', webUrl: 'https://x/csv', parentReference: { driveId: 'DRIVE-ME', path: '/drive/root:' }, file: { mimeType: 'text/csv' } },
    { id: 'FOLDER', name: 'TPAHA_2026 folder', lastModifiedDateTime: '2026-09-01T00:00:00Z', webUrl: 'https://x/folder', parentReference: { driveId: 'DRIVE-ME', path: '/drive/root:' }, folder: { childCount: 2 } },
  ],
};
const sharedResult = {
  value: [
    { id: 'LOCAL-REF-1', name: 'TPAHA_2026 (1).xlsx', remoteItem: { id: 'ITEM-XLSX', name: 'TPAHA_2026 (1).xlsx', size: 104054, lastModifiedDateTime: '2026-09-06T01:23:54Z', webUrl: 'https://tenant-my.sharepoint.com/personal/x/Documents/TPAHA_2026%20(1).xlsx', parentReference: { driveId: 'DRIVE-ME' }, file: { mimeType: 'x' }, shared: { sharedBy: { user: { displayName: 'Treasurer' } } } } },
    { id: 'LOCAL-REF-2', name: 'Winter storage.xlsx', remoteItem: { id: 'ITEM-STORAGE', name: 'Winter storage.xlsx', size: 77339, lastModifiedDateTime: '2026-08-01T00:00:00Z', webUrl: 'https://x/storage', parentReference: { driveId: 'DRIVE-OTHER', path: '/drives/DRIVE-OTHER/root:/Shared' }, file: { mimeType: 'x' }, shared: { sharedBy: { user: { displayName: 'Vice President' } } } } },
    { id: 'LOCAL-REF-3', name: 'Photos', remoteItem: { id: 'ITEM-FOLDER', name: 'Photos', parentReference: { driveId: 'DRIVE-OTHER' }, folder: {} } },
  ],
};
const graphGet = async path => {
  if (path.includes('/search(')) return searchResult;
  if (path.includes('/sharedWithMe')) return sharedResult;
  throw new Error('unexpected ' + path);
};

test('findCandidates lists only .xlsx files, merges shared items, dedupes, and sorts newest first', async () => {
  const c = await findCandidates(graphGet, { query: 'TPAHA_2026' });
  assert.deepEqual(c.map(x => x.itemId), ['ITEM-XLSX', 'ITEM-STORAGE']);
  const mine = c[0];
  assert.equal(mine.driveId, 'DRIVE-ME');
  assert.equal(mine.extension, 'xlsx');
  assert.equal(mine.path, '/Documents');
  assert.equal(mine.source, 'mine');
  assert.equal(mine.size, 104054);
  assert.equal(c[1].source, 'shared');
  assert.equal(c[1].sharedBy, 'Vice President');
  assert.equal(workbookKey(mine), 'DRIVE-ME!ITEM-XLSX');
});

test('findCandidates tolerates a failing sharedWithMe call', async () => {
  const c = await findCandidates(async p => { if (p.includes('/search(')) return searchResult; throw new Error('403'); }, { query: 'TPAHA' });
  assert.equal(c.length, 1);
  assert.equal(c[0].warnings.length, 0);
  assert.equal(c.warnings.length, 1);
});

test('remember / recall / forget are per tenant, account and app', () => {
  const storage = new MemStorage();
  const who = { tenantId: 'T', homeAccountId: 'U1.T', app: 'ledger' };
  const choice = { driveId: 'DRIVE-ME', itemId: 'ITEM-XLSX', name: 'TPAHA_2026 (1).xlsx', webUrl: 'https://x' };
  rememberWorkbook(storage, who, choice);
  const back = recallWorkbook(storage, who);
  assert.equal(back.itemId, 'ITEM-XLSX');
  assert.ok(back.chosenAt);
  assert.equal(recallWorkbook(storage, { ...who, homeAccountId: 'U2.T' }), null);
  assert.equal(recallWorkbook(storage, { ...who, app: 'storage' }), null);
  forgetWorkbook(storage, who);
  assert.equal(recallWorkbook(storage, who), null);
});

test('resolveConfigured returns the configured workbook only when both ids are present', () => {
  assert.equal(resolveConfigured({ workbooks: { ledger: { driveId: '', itemId: '', name: '' } } }, 'ledger'), null);
  assert.deepEqual(resolveConfigured({ workbooks: { ledger: { driveId: 'D', itemId: 'I', name: 'n.xlsx' } } }, 'ledger'), { driveId: 'D', itemId: 'I', name: 'n.xlsx', source: 'config' });
  assert.equal(resolveConfigured({}, 'ledger'), null);
});
