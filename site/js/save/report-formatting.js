// Unfinished monthly report formatting, kept in the browser so a reload does not forget it.
//
// A transaction can reach the workbook while the row-visibility pass that follows it does not finish: a lost
// connection, a slow service, or a change to the month sheets' protection. The transaction is saved and must
// never be written again; only the formatting is outstanding. This record names the months still to do.
//
// It holds no workbook content and no password: month numbers and a timestamp, nothing else.
export class ReportFormattingStorageError extends Error {
  constructor(message, cause) { super(message); this.name = 'ReportFormattingStorageError'; this.cause = cause; }
}
const seg = s => encodeURIComponent(String(s));

export function reportFormattingKey({ tenantId, homeAccountId, workbookId }) {
  if (!tenantId) throw new Error('reportFormattingKey needs the tenant id');
  if (!homeAccountId) throw new Error('reportFormattingKey needs the signed-in account id');
  if (!workbookId) throw new Error('reportFormattingKey needs the workbook id');
  return `tpaha:reportfmt:v1:${seg(tenantId)}:${seg(homeAccountId)}:${seg(workbookId)}`;
}

const clean = months => [...new Set((months || []).map(Number).filter(m => Number.isInteger(m) && m >= 1 && m <= 12))].sort((a, b) => a - b);

export class ReportFormattingStore {
  constructor(storage, key) { this.storage = storage; this.key = key; }
  /** { months: number[], message: string, at: string } or null. */
  read() {
    let raw = null;
    try { raw = this.storage.getItem(this.key); } catch { return null; }
    if (raw === null || raw === undefined) return null;
    try {
      const rec = JSON.parse(raw);
      const months = clean(rec && rec.months);
      return months.length ? { ...rec, months } : null;
    } catch { try { this.storage.removeItem(this.key); } catch { /* ignore */ } return null; }
  }
  /**
   * Adds months still to do. Returns the stored record, or null when there is nothing left.
   * Throws ReportFormattingStorageError when the browser refuses to keep it: a caller must say so rather than
   * let a member believe the reminder will survive a reload.
   */
  add(months, message, blockedBy) {
    const current = this.read();
    const merged = clean([...(current?.months || []), ...clean(months)]);
    if (!merged.length) return this.clear();
    const rec = { months: merged, message: message || current?.message || '', blockedBy: blockedBy || current?.blockedBy || 'formatting', at: new Date().toISOString() };
    try { this.storage.setItem(this.key, JSON.stringify(rec)); } catch (e) {
      throw new ReportFormattingStorageError('This browser could not remember that the monthly report rows still need updating (its storage is full or blocked). The reminder is on this page only: if you close or reload it, use Refresh and check the month in Excel.', e);
    }
    return rec;
  }
  /**
   * Reserves months before an operation writes anything, and returns `added`: only the months that were not
   * already outstanding. Releasing an operation that wrote nothing must release only those, or it erases an
   * older reminder that is still unfinished and whose rows are still wrong (review W2b).
   *
   * An existing record keeps its own message and reason. A reservation is made before anything is attempted,
   * so it knows less than a completed operation did, and must not overwrite what that operation reported.
   * Only verified completion clears months, through remove().
   */
  reserve(months, message, blockedBy) {
    const current = this.read();
    const before = new Set(current ? current.months : []);
    const wanted = clean(months);
    const added = wanted.filter(m => !before.has(m));
    const msg = (current && current.message) || message;
    const why = (current && current.blockedBy) || blockedBy;
    try {
      return { rec: this.add(wanted, msg, why), added };
    } catch (e) {
      e.added = added;                                                    // so the caller can still release its own
      e.fallback = { months: [...new Set([...before, ...wanted])], message: msg, blockedBy: why };
      throw e;
    }
  }
  /** Drops months that have since been formatted. Returns what is left, or null. */
  remove(months) {
    const current = this.read();
    if (!current) return null;
    const done = new Set(clean(months));
    const left = current.months.filter(m => !done.has(m));
    if (!left.length) return this.clear();
    const rec = { ...current, months: left };
    try { this.storage.setItem(this.key, JSON.stringify(rec)); } catch { /* the months left are still shown on this page */ }
    return rec;
  }
  clear() { try { this.storage.removeItem(this.key); } catch { /* ignore */ } return null; }
}
