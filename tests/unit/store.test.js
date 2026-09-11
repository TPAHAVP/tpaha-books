import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Store } from '../../site/js/store.js';

class MemStorage {
  constructor() { this.m = new Map(); }
  getItem(k) { return this.m.has(k) ? this.m.get(k) : null; }
  setItem(k, v) { this.m.set(k, String(v)); }
  removeItem(k) { this.m.delete(k); }
}
const createEmpty = () => ({ schema: 1, items: [], updatedAt: '' });
const opts = extra => ({ key: 't', fileName: 't.json', createEmpty, ...extra });

test('the Store is local-only: it no longer accepts a cloud backend', () => {
  assert.throws(() => new Store(opts({ backend: {}, storage: new MemStorage() })), /no longer syncs/);
});

test('local mode round-trips through storage synchronously (no dirty window)', async () => {
  const storage = new MemStorage();
  const s = new Store(opts({ storage }));
  await s.load();
  assert.deepEqual(s.doc.items, []);
  assert.equal(s.status.mode, 'local');
  await s.mutate(d => { d.items.push('a'); d.updatedAt = '1'; });
  assert.equal(s.status.dirty, false);
  assert.deepEqual(JSON.parse(storage.getItem('tpaha:t:t.json')).doc.items, ['a']);
  const s2 = new Store(opts({ storage }));
  await s2.load();
  assert.deepEqual(s2.doc.items, ['a']);
});

test('import/export JSON', async () => {
  const storage = new MemStorage();
  const s = new Store(opts({ storage }));
  await s.load();
  await s.importJson(JSON.stringify({ schema: 1, items: ['z'], updatedAt: '9' }));
  assert.deepEqual(JSON.parse(s.exportJson()).items, ['z']);
  await assert.rejects(() => s.importJson('{"nope":1}'), /not a valid/i);
  await assert.rejects(() => s.importJson('garbage'), /not valid JSON/i);
});
