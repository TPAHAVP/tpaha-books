// Unfinished monthly report formatting, kept in the browser so a reload does not forget it.
//
// A transaction can reach the workbook while the row-visibility pass that follows it does not finish: a lost
// connection, a slow service, or a change to the month sheets' protection. The transaction is saved and must
// never be written again; only the formatting is outstanding. This record names the months still to do.
//
// It holds no workbook content and no password: month numbers and a timestamp, nothing else.
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
  /** Adds months still to do. Returns the stored record, or null when there is nothing left. */
  add(months, message) {
    const merged = clean([...(this.read()?.months || []), ...clean(months)]);
    if (!merged.length) return this.clear();
    const rec = { months: merged, message: message || this.read()?.message || '', at: new Date().toISOString() };
    try { this.storage.setItem(this.key, JSON.stringify(rec)); } catch { /* the banner still shows in this page */ }
    return rec;
  }
  /** Drops months that have since been formatted. Returns what is left, or null. */
  remove(months) {
    const current = this.read();
    if (!current) return null;
    const done = new Set(clean(months));
    const left = current.months.filter(m => !done.has(m));
    if (!left.length) return this.clear();
    const rec = { ...current, months: left };
    try { this.storage.setItem(this.key, JSON.stringify(rec)); } catch { /* ignore */ }
    return rec;
  }
  clear() { try { this.storage.removeItem(this.key); } catch { /* ignore */ } return null; }
}
