import { test } from 'node:test';
import assert from 'node:assert/strict';
import { draftKey, DraftStore, DraftStorageError, installUnloadGuard, clearDraftsForAccount, combineDrafts } from '../../site/js/save/drafts.js';

class MemStorage {
  constructor() { this.m = new Map(); this.failWrites = false; }
  get length() { return this.m.size; }
  key(i) { return [...this.m.keys()][i] ?? null; }
  getItem(k) { return this.m.has(k) ? this.m.get(k) : null; }
  setItem(k, v) { if (this.failWrites) throw new DOMException('quota', 'QuotaExceededError'); this.m.set(k, String(v)); }
  removeItem(k) { this.m.delete(k); }
}
const ids = { tenantId: 'tenant-a', homeAccountId: 'user-1.tenant-a', workbookId: 'D1!I1', schema: 1 };

test('draft keys separate tenant, account, workbook and schema', () => {
  const base = draftKey(ids);
  assert.equal(draftKey({ ...ids }), base, 'deterministic');
  assert.notEqual(draftKey({ ...ids, tenantId: 'tenant-b' }), base);
  assert.notEqual(draftKey({ ...ids, homeAccountId: 'user-2.tenant-a' }), base);
  assert.notEqual(draftKey({ ...ids, workbookId: 'D1!I2' }), base);
  assert.notEqual(draftKey({ ...ids, schema: 2 }), base);
  assert.match(base, /^tpaha:draft:/);
  assert.throws(() => draftKey({ ...ids, homeAccountId: '' }), /account/i);
});

test('drafts round-trip, survive only for the same account, and clear', () => {
  const storage = new MemStorage();
  const mine = new DraftStore(storage, draftKey(ids));
  mine.write({ form: { description: 'half typed' } });
  assert.deepEqual(mine.read().form, { description: 'half typed' });
  assert.ok(mine.read().savedAt, 'stamped');
  const other = new DraftStore(storage, draftKey({ ...ids, homeAccountId: 'user-2.tenant-a' }));
  assert.equal(other.read(), null, 'another account cannot see the draft');
  mine.clear();
  assert.equal(mine.read(), null);
});

test('storage failures are surfaced, corrupt drafts are discarded', () => {
  const storage = new MemStorage();
  const store = new DraftStore(storage, draftKey(ids));
  storage.failWrites = true;
  assert.throws(() => store.write({ a: 1 }), DraftStorageError);
  storage.failWrites = false;
  storage.setItem(draftKey(ids), '{not json');
  assert.equal(store.read(), null);
  assert.equal(storage.getItem(draftKey(ids)), null, 'corrupt entry removed');
});

test('clearDraftsForAccount removes only that account\'s drafts', () => {
  const storage = new MemStorage();
  new DraftStore(storage, draftKey(ids)).write({ a: 1 });
  new DraftStore(storage, draftKey({ ...ids, workbookId: 'D1!I2' })).write({ a: 2 });
  const otherKey = draftKey({ ...ids, homeAccountId: 'user-2.tenant-a' });
  new DraftStore(storage, otherKey).write({ a: 3 });
  const removed = clearDraftsForAccount(storage, { tenantId: ids.tenantId, homeAccountId: ids.homeAccountId });
  assert.equal(removed, 2);
  assert.ok(storage.getItem(otherKey));
});

test('the unload guard warns only while there are unsaved edits and can be removed', () => {
  const handlers = {};
  const win = { addEventListener: (t, h) => { handlers[t] = h; }, removeEventListener: t => { delete handlers[t]; } };
  let unsaved = true;
  const uninstall = installUnloadGuard(() => unsaved, win);
  const ev = { preventDefault() { this.prevented = true; }, returnValue: undefined };
  handlers.beforeunload(ev);
  assert.equal(ev.prevented, true);
  assert.ok(ev.returnValue);
  unsaved = false;
  const ev2 = { preventDefault() { this.prevented = true; }, returnValue: undefined };
  handlers.beforeunload(ev2);
  assert.equal(ev2.prevented, undefined);
  uninstall();
  assert.equal(handlers.beforeunload, undefined);
});

test('T4: combineDrafts keeps each form part on its own: a saved entry does not erase a typed or pending correction', () => {
  const ref = { id: 2, timestamp: 't2', fingerprint: 'fp' };
  const typed = combineDrafts({ entry: { draft: null, state: 'saved', pending: null }, correction: { ref, draft: { description: 'Separate correction must survive' }, pending: null, hasUnsavedEdits: true } });
  assert.equal(typed.entry, null);
  assert.equal(typed.pending, null);
  assert.equal(typed.correction.draft.description, 'Separate correction must survive');
  assert.equal(typed.correction.marker, null);
  const pending = combineDrafts({ entry: { draft: { description: 'x' }, state: 'saved', pending: null }, correction: { ref, draft: { amount: '9' }, pending: { marker: 'op', payload: { amount: '9' } }, hasUnsavedEdits: true } });
  assert.equal(pending.entry, null, 'a saved entry is dropped');
  assert.deepEqual(pending.correction, { ref, marker: 'op', payload: { amount: '9' }, draft: { amount: '9' } });
  const entryOnly = combineDrafts({ entry: { draft: { description: 'typed' }, state: 'unsaved', pending: { marker: 'm', payload: { description: 'typed' } } }, correction: { ref, draft: null, pending: null, hasUnsavedEdits: false } });
  assert.deepEqual(entryOnly.entry, { description: 'typed' });
  assert.equal(entryOnly.pending.marker, 'm');
  assert.equal(entryOnly.correction, null);
  assert.equal(combineDrafts({ entry: { draft: null, state: 'idle', pending: null }, correction: null }), null, 'nothing to keep');
});
