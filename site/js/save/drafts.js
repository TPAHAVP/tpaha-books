// Unsaved-work drafts kept in the browser, namespaced by tenant, account, workbook and schema.
// A draft is never "saved": it is what the member typed and has not yet committed to the workbook.
export class DraftStorageError extends Error {
  constructor(message, cause) { super(message); this.name = 'DraftStorageError'; this.cause = cause; }
}

const seg = s => encodeURIComponent(String(s));

export function draftKey({ tenantId, homeAccountId, workbookId, schema = 1 }) {
  if (!tenantId) throw new Error('draftKey needs the tenant id');
  if (!homeAccountId) throw new Error('draftKey needs the signed-in account id');
  if (!workbookId) throw new Error('draftKey needs the workbook id');
  return `tpaha:draft:v${Number(schema)}:${seg(tenantId)}:${seg(homeAccountId)}:${seg(workbookId)}`;
}

export class DraftStore {
  constructor(storage, key) {
    this.storage = storage;
    this.key = key;
  }
  read() {
    let raw = null;
    try { raw = this.storage.getItem(this.key); } catch { return null; }
    if (raw === null || raw === undefined) return null;
    try { return JSON.parse(raw); }
    catch { try { this.storage.removeItem(this.key); } catch { /* ignore */ } return null; }
  }
  write(obj) {
    const payload = JSON.stringify({ ...obj, savedAt: new Date().toISOString() });
    try { this.storage.setItem(this.key, payload); }
    catch (e) { throw new DraftStorageError('Your unsaved changes could not be kept in this browser (storage is full or blocked). They are still on screen; save them to the workbook now.', e); }
  }
  clear() { try { this.storage.removeItem(this.key); } catch { /* ignore */ } }
}

/**
 * One record for everything a page has not confirmed in the workbook: the typed entry, its unconfirmed
 * operation, and the correction being typed or left unconfirmed. Each part is kept or dropped on its own,
 * so a form that finishes never erases another form's draft (review finding T4). Null = nothing to keep.
 */
export function combineDrafts({ entry = null, correction = null } = {}) {
  const rec = {
    entry: entry && entry.draft && entry.state !== 'saved' ? entry.draft : null,
    pending: entry && entry.pending ? { marker: entry.pending.marker, payload: entry.pending.payload } : null,
    correction: correction && correction.ref && (correction.pending || (correction.hasUnsavedEdits && correction.draft))
      ? { ref: correction.ref, marker: correction.pending ? correction.pending.marker : null, payload: correction.pending ? correction.pending.payload : correction.draft, draft: correction.draft || null }
      : null,
  };
  return rec.entry || rec.pending || rec.correction ? rec : null;
}

/** Removes every draft belonging to one tenant + account. Returns how many were removed. */
export function clearDraftsForAccount(storage, { tenantId, homeAccountId }) {
  const marker = `:${seg(tenantId)}:${seg(homeAccountId)}:`;
  const keys = [];
  for (let i = 0; i < storage.length; i++) {
    const k = storage.key(i);
    if (k && k.startsWith('tpaha:draft:') && k.includes(marker)) keys.push(k);
  }
  for (const k of keys) storage.removeItem(k);
  return keys.length;
}

/** Warns before the page closes while there are unsaved edits. Returns an uninstall function. */
export function installUnloadGuard(hasUnsaved, win = globalThis.window) {
  const handler = ev => {
    if (!hasUnsaved()) return undefined;
    ev.preventDefault();
    ev.returnValue = 'You have unsaved changes.';
    return ev.returnValue;
  };
  win.addEventListener('beforeunload', handler);
  return () => win.removeEventListener('beforeunload', handler);
}
