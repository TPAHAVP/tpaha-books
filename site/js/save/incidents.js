// An "incident" is an unresolved write: a delete or correction whose read-back showed damage the app must
// not repair on its own (a wrong row removed, two copies of a record, a duplicate number). It is kept in the
// browser so that a page reload or Refresh does not forget it; writes stay paused until the adapter has
// verified the documented resolution conditions (or a member acknowledged an inconclusive case).
const seg = s => encodeURIComponent(String(s));

export function incidentKey({ tenantId, homeAccountId, workbookId }) {
  if (!tenantId) throw new Error('incidentKey needs the tenant id');
  if (!homeAccountId) throw new Error('incidentKey needs the signed-in account id');
  if (!workbookId) throw new Error('incidentKey needs the workbook id');
  return `tpaha:incident:v1:${seg(tenantId)}:${seg(homeAccountId)}:${seg(workbookId)}`;
}

export class IncidentStore {
  constructor(storage, key) { this.storage = storage; this.key = key; }
  read() {
    let raw = null;
    try { raw = this.storage.getItem(this.key); } catch { return null; }
    if (raw === null || raw === undefined) return null;
    try { const r = JSON.parse(raw); return r && r.kind ? r : null; }
    catch { try { this.storage.removeItem(this.key); } catch { /* ignore */ } return null; }
  }
  /** Stores the record; storage failures are reported to the caller (the incident is still in memory). */
  write(record) {
    try { this.storage.setItem(this.key, JSON.stringify(record)); return true; }
    catch { return false; }
  }
  clear() { try { this.storage.removeItem(this.key); } catch { /* ignore */ } }
}
