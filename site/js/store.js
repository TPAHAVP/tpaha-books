// Local-only document cache used by the Winter Storage screen until it is connected to its workbook.
// It never syncs anywhere: the earlier OneDrive JSON path was removed because the workbook is the record.
export class Store {
  constructor({ key, fileName, backend = null, createEmpty, storage = globalThis.localStorage }) {
    if (backend) throw new Error('Store no longer syncs to OneDrive; the workbook is the record. Use the workbook adapter instead.');
    this.key = key;
    this.fileName = fileName;
    this.createEmpty = createEmpty;
    this.storage = storage;
    this.doc = null;
    this.listeners = [];
    this.status = { mode: 'local', dirty: false, offline: false, syncing: false, lastSynced: null, error: null, notices: [] };
  }

  onChange(fn) { this.listeners.push(fn); }
  _emit() { for (const fn of this.listeners) fn(this.doc, this.status); }
  _cacheKey() { return `tpaha:${this.key}:${this.fileName}`; }
  _readCache() {
    try { const raw = this.storage && this.storage.getItem(this._cacheKey()); return raw ? JSON.parse(raw) : null; } catch { return null; }
  }
  _writeCache() {
    try { this.storage && this.storage.setItem(this._cacheKey(), JSON.stringify({ doc: this.doc, etag: null, dirty: false })); }
    catch (e) { this.status.error = 'This browser could not keep the data (storage full or blocked).'; }
  }
  clearCache() { try { this.storage && this.storage.removeItem(this._cacheKey()); } catch { /* ignore */ } }

  async load() {
    const cached = this._readCache();
    this.doc = cached && cached.doc ? cached.doc : this.createEmpty();
    this._writeCache();
    this._emit();
    return this.doc;
  }

  /** Apply a change and write it to the browser immediately. Returns whatever fn returns. */
  async mutate(fn) {
    const result = fn(this.doc);
    this._writeCache();
    this._emit();
    return result;
  }

  async push() { /* nothing to push: local only */ }

  exportJson() { return JSON.stringify(this.doc, null, 1); }

  async importJson(text) {
    let parsed;
    try { parsed = JSON.parse(text); } catch { throw new Error('That file is not valid JSON.'); }
    const empty = this.createEmpty();
    for (const k of Object.keys(empty)) {
      if (k === 'updatedAt' || k === 'updatedBy') continue;
      if (!(k in parsed)) throw new Error(`That file is not a valid ${this.key} data file (missing "${k}").`);
    }
    this.doc = parsed;
    this._writeCache();
    this._emit();
  }
}
