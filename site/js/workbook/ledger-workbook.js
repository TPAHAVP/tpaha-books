// Ledger adapter: turns website actions into targeted operations on the real TPAHA ledger workbook.
//
// Concurrency boundary (docs/workbook-mapping.md §3b has the full statement):
// - Graph Excel has no conditional write, lock or transaction. Appends never move other rows. Deleting a
//   row and renumbering a row act on an INDEX, so they are preceded by a fresh read that is evaluated
//   before the request is sent (the row at that index must be the intended one) and followed by a
//   read-back. A shift can still happen between that read and the request; the app cannot prevent it,
//   which is why the pilot runs with one designated writer (config.pilot.writerModel).
// - Rows are identified by TransactionID + Timestamp ("identity"); a change to a row's other cells is a
//   fingerprint change, never a reason to recreate or delete anything.
// - Corrections never overwrite a row by index: the corrected copy is appended (position-safe) and the
//   old copy is then deleted with the checked delete.
// - When a read-back does not add up, the adapter records an INCIDENT (UnresolvedOperationError) that
//   names the rows involved and pauses every write. load() does NOT clear it: verifyResolution() re-reads
//   the table and checks the documented conditions; only a verified (or, when nothing can be verified,
//   an explicitly acknowledged) resolution lifts the pause. Repairs are explicit calls (removeCopy,
//   reappendRow) made by a member after seeing the details; nothing is repaired by inference.
// - inspect*() methods are READ-ONLY; complete*() methods finish an operation that landed.
// - Nothing here retries a write. The client retries only GETs.
import { MONTH_NAMES, CATEGORIES, toCents } from '../ledger/model.js';
import { ExcelApiError } from './excel-client.js';
import { parseAddress, rangeAddress, serialToIso, isoToSerial, isSerialDate } from './excel-values.js';

export const LOG_COLUMNS = ['TransactionID', 'Timestamp', 'Date', 'Type', 'Category', 'Amount', 'Description', 'ChequeNum', 'Notes'];
export const DATE_FORMAT = 'yyyy-mm-dd';
export const MONEY_FORMAT = '$#,##0.00';
const FORMAT_ROW = [null, null, DATE_FORMAT, null, null, MONEY_FORMAT, null, null, null];
const SELF_TEST_DESCRIPTION = 'TPAHA Books connection test';
const MONTH_FIRST_ROW = 4, MONTH_LAST_ROW = 33;   // the transaction rows on every monthly report sheet
export const REPORT_FORMATTING_SAVED = 'Transaction saved. Excel report formatting still needs updating.';
export const REPORT_FORMATTING_DELETED = 'Transaction deleted. Excel report formatting still needs updating.';
const SORTED_WARNING = 'The monthly sheets may be out of date: the sorted helper table could not be confirmed. Press Refresh to retry.';
const FORMAT_WARNING = 'The number formats on the new row could not be confirmed (its date may display as a number until Excel is opened). The values are saved.';

export class ConflictError extends Error {
  constructor(reason, { expected = null, current = null, message } = {}) {
    super(message || (reason === 'missing' ? 'This record no longer exists in the workbook.' : 'Someone else changed this record in the workbook.'));
    this.name = 'ConflictError'; this.reason = reason; this.expected = expected; this.current = current;
  }
}
export class VerificationError extends Error {
  constructor(step, message, details = null) { super(message); this.name = 'VerificationError'; this.step = step; this.details = details; }
}
/** The workbook is in a state the adapter cannot explain or fix safely. Writes stop until the incident is resolved. */
export class UnresolvedOperationError extends Error {
  constructor(kind, message, details = {}) { super(message); this.name = 'UnresolvedOperationError'; this.kind = kind; this.details = details; }
}
/** A row already carries this operation id but with different content: two operations shared an id. */
export class MarkerCollisionError extends Error {
  constructor(marker, existing) {
    super(`Another entry in the workbook already uses this operation id (${marker}) with different content. Your entry was not written. Refresh and enter it again.`);
    this.name = 'MarkerCollisionError'; this.marker = marker; this.existing = existing;
  }
}

// ---- row helpers -----------------------------------------------------------------------------
export function normalizeRow(values) {
  const out = [];
  for (let c = 0; c < LOG_COLUMNS.length; c++) { const v = values ? values[c] : undefined; out.push(v === null || v === undefined ? '' : v); }
  return out;
}
export const isBlankRow = values => normalizeRow(values).every(v => v === '');
export function rowFingerprint(values) { return JSON.stringify(normalizeRow(values)); }
export function sameCell(a, b) {
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < 1e-9;
  if (typeof a === 'number' && typeof b === 'string' && b.trim() !== '' && Number.isFinite(Number(b))) return Math.abs(a - Number(b)) < 1e-9;
  if (typeof b === 'number' && typeof a === 'string') return sameCell(b, a);
  return String(a) === String(b);
}
export function sameRow(a, b) { const x = normalizeRow(a), y = normalizeRow(b); return x.every((v, i) => sameCell(v, y[i])); }
const sameMatrix = (a, b) => a.length === b.length && a.every((r, i) => sameRow(r, b[i]));
export const identityOf = t => `${t.id}|${t.timestamp}`;
const identityOfValues = v => `${v[0] === '' ? null : Number(v[0])}|${String(v[1])}`;
/** Identities that occur more than once (two copies of the same record). */
export function duplicateIdentities(txns) {
  const seen = new Map();
  for (const t of txns) seen.set(identityOf(t), (seen.get(identityOf(t)) || 0) + 1);
  return [...seen.entries()].filter(([, n]) => n > 1).map(([k]) => k);
}

export function rowToTxn(values, rowIndex) {
  const v = normalizeRow(values);
  const rawDate = v[2];
  return {
    rowIndex,
    id: v[0] === '' ? null : Number(v[0]),
    timestamp: String(v[1]),
    date: isSerialDate(rawDate) ? serialToIso(rawDate) : String(rawDate).slice(0, 10),
    type: String(v[3]), category: String(v[4]),
    amount: v[5] === '' ? 0 : Number(v[5]),
    description: String(v[6]), chequeNum: v[7] === '' ? '' : String(v[7]), notes: String(v[8]),
    values: v, fingerprint: rowFingerprint(v),
  };
}
const publicTxn = t => t ? ({ rowIndex: t.rowIndex, id: t.id, timestamp: t.timestamp, date: t.date, type: t.type, category: t.category, amount: t.amount, description: t.description, chequeNum: t.chequeNum, notes: t.notes, values: t.values, fingerprint: t.fingerprint }) : null;

/**
 * Operation id written into the Timestamp column (mapping §2): the entry time as ISO 8601 UTC with NINE fraction
 * digits (the millisecond plus six random digits). It stays a valid ISO 8601 date-time, so scripts that parse the
 * column with new Date() keep working, and two members pressing Save in the same millisecond get different ids.
 */
export function makeMarker() {
  const bytes = new Uint8Array(6);
  globalThis.crypto.getRandomValues(bytes);
  const iso = new Date().toISOString();                                             // yyyy-mm-ddThh:mm:ss.mmmZ
  return `${iso.slice(0, -1)}${Array.from(bytes, b => String(b % 10)).join('')}Z`;
}
function normalizeCheque(x) { const s = String(x ?? '').trim(); if (s === '') return ''; return /^\d+$/.test(s) ? Number(s) : s; }
export function entryToRow(entry, marker) {
  const date = String(entry.date || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('The date must be in yyyy-mm-dd form.');
  const amount = Math.round(Number(String(entry.amount ?? '').replace(/[$,\s]/g, '')) * 100) / 100;
  if (!Number.isFinite(amount) || !(amount > 0)) throw new Error('The amount must be greater than 0.');
  if (!['Deposit', 'Withdrawal'].includes(entry.type)) throw new Error('The type must be Deposit or Withdrawal.');
  const category = String(entry.category || '').trim();
  if (!category) throw new Error('A category is required.');
  return [null, marker, isoToSerial(date), entry.type, category, amount, String(entry.description || '').trim(), normalizeCheque(entry.chequeNum), String(entry.notes || '').trim()];
}
function mergeChanges(current, changes) {
  const v = [...current.values];
  if ('date' in changes) { const d = String(changes.date).trim(); if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new Error('The date must be in yyyy-mm-dd form.'); v[2] = isoToSerial(d); }
  if ('type' in changes) { if (!['Deposit', 'Withdrawal'].includes(changes.type)) throw new Error('The type must be Deposit or Withdrawal.'); v[3] = changes.type; }
  if ('category' in changes) v[4] = String(changes.category).trim();
  if ('amount' in changes) { const a = Math.round(Number(String(changes.amount).replace(/[$,\s]/g, '')) * 100) / 100; if (!(a > 0)) throw new Error('The amount must be greater than 0.'); v[5] = a; }
  if ('description' in changes) v[6] = String(changes.description ?? '').trim();
  if ('chequeNum' in changes) v[7] = normalizeCheque(changes.chequeNum);
  if ('notes' in changes) v[8] = String(changes.notes ?? '').trim();
  return v;
}
/** Same content apart from id: columns Date..Notes. */
export function payloadMatches(rowValues, sentValues) {
  const a = normalizeRow(rowValues), b = normalizeRow(sentValues);
  for (let c = 2; c < LOG_COLUMNS.length; c++) if (!sameCell(a[c], b[c])) return false;
  return true;
}
const nextIdFrom = txns => txns.reduce((m, t) => Math.max(m, Number.isFinite(t.id) ? t.id : 0), 0) + 1;
const num = x => (x === '' || x === null || x === undefined ? 0 : Number(x));
const describe = t => `#${t.id} ${t.date} ${t.description} ${t.amount}`;
const idOrLast = t => (Number.isFinite(t.id) ? t.id : Number.MAX_SAFE_INTEGER);
/**
 * The order LOG_Sorted must hold: the non-blank rows of LOG by Date, then by TransactionID, then by table
 * position. The tie-break matches the workbook's own Office Scripts, whose `ordered()` sorts by date serial then
 * transaction id and whose `verifySorted()` refuses to run at all when LOG_Sorted differs from that order.
 * Ordering by table position instead puts the two out of step as soon as this app corrects a row that shares a
 * date with a higher-numbered one, because a correction moves its row to the end of LOG. A row with no
 * transaction number sorts last and keeps table order.
 *
 * Exported so that every consumer applies one rule: the rebuild below and the read-only Diagnostics check used
 * to hold separate copies of it, and disagreed once the tie-break changed (review finding V1).
 */
export function expectedSortedRows(logValues) {
  return (logValues || [])
    .map((values, index) => rowToTxn(normalizeRow(values), index))
    .filter(t => !isBlankRow(t.values))
    .sort((a, b) => a.date.localeCompare(b.date) || idOrLast(a) - idOrLast(b) || a.rowIndex - b.rowIndex)
    .map(t => t.values);
}
/** Read-only: does LOG_Sorted already hold exactly those rows, in that order? Returns the counts for a message. */
export function sortedTableMatches(logValues, sortedValues) {
  const want = expectedSortedRows(logValues);
  const have = (sortedValues || []).map(normalizeRow);
  return { ok: sameMatrix(have, want), want, have };
}
const sortedTarget = log => expectedSortedRows(log.rows.map(r => r.values));
const rowFormatsOk = nf => Array.isArray(nf) && nf[2] === DATE_FORMAT && nf[5] === MONEY_FORMAT;
/** 1-12 from an ISO date, 0 when it cannot be read. */
const monthOfIso_or_number = v => (typeof v === 'number' ? (Number.isInteger(v) ? v : 0) : monthOfIso(v));
export const monthOfIso = iso => { const m = Number(String(iso || '').slice(5, 7)); return Number.isInteger(m) && m >= 1 && m <= 12 ? m : 0; };
/**
 * Consecutive report rows that should share a visibility, so the whole sheet is set with a couple of requests
 * instead of thirty. `populated[i]` is true when report row MONTH_FIRST_ROW + i has a date in it.
 */
export function visibilityRuns(populated) {
  const runs = [];
  populated.forEach((isPopulated, i) => {
    const row = MONTH_FIRST_ROW + i, hidden = !isPopulated, last = runs[runs.length - 1];
    if (last && last.hidden === hidden) last.last = row;
    else runs.push({ first: row, last: row, hidden });
  });
  return runs;
}
/** The record a screen shows when the transaction landed but its report formatting did not finish. */
const pendingReportFormatting = (vis, message) => (vis.ok ? null : { pending: true, months: vis.months, message, detail: vis.error });

// ---- adapter -----------------------------------------------------------------------------------
export class LedgerWorkbook {
  constructor(client, { tableName = 'LOG_Table', sortedTableName = 'LOG_Sorted_Table', listsTableName = 'LISTS_Categories', log } = {}) {
    this.client = client;
    this.table = tableName;
    this.sortedTable = sortedTableName;
    this.listsTable = listsTableName;
    this.logFn = log || (() => {});
    this.year = null;
    this.incident = null;      // { kind, message, details, at, resolved, resolution, checks }
    this.sortedStale = false;
    this._chain = Promise.resolve();   // operation queue: one complete operation at a time (T1)
    this._pending = 0;
    this._selfTestRunning = false;
  }

  /** The unresolved incident, or null. Every write is refused while it is set. */
  get halted() { return this.incident && !this.incident.resolved ? this.incident : null; }
  _guard() {
    if (this.halted) throw new UnresolvedOperationError('halted', `Changes are paused because an earlier change could not be verified (${this.halted.message}). Resolve it from the details shown before making more changes.`, { cause: this.halted.kind, causeDetails: this.halted.details });
  }
  _halt(err) {
    this.incident = { kind: err.kind, message: err.message, details: err.details || {}, at: new Date().toISOString(), resolved: false, resolution: null, checks: null };
    this.logFn({ op: 'halt', kind: err.kind });
    return err;
  }
  /** Re-arms an incident recorded earlier (before a page reload). Returns the unresolved incident or null. */
  restoreIncident(record) {
    if (record && record.kind && !record.resolved) this.incident = { ...record };
    return this.halted;
  }

  _bodyPath() { return this.client.tableBodyPath(this.table, 'address,values,rowCount'); }
  _parseBody(body) {
    const parsed = parseAddress(body.address);
    const rows = (body.values || []).map((values, index) => ({ index, values: normalizeRow(values) }));
    return { sheet: parsed.sheet, firstRow: parsed.r1, address: body.address, rowCount: body.rowCount, rows, txns: rows.filter(r => !isBlankRow(r.values)).map(r => rowToTxn(r.values, r.index)) };
  }
  async _readLog() { return this._parseBody(await this.client.getTableBody(this.table)); }

  /**
   * Reads everything the ledger screens need. Does NOT lift a pause; it may add one if the table itself is
   * inconsistent (two copies of a record). `expectedDuplicates` lists identities whose second copy is a known,
   * still-pending correction of this page (it will be finished or reported by that operation, not here).
   */
  async _load({ expectedDuplicates = [] } = {}) {
    const meta = await this.client.getItemMeta().catch(() => ({}));
    const yearRange = await this.client.getRange('ConfigHidden', 'A1');
    const year = Number(yearRange.values[0][0]);
    if (!Number.isInteger(year) || year < 2000) throw new VerificationError('year', `ConfigHidden!A1 should hold the workbook year, found "${yearRange.values[0][0]}".`);
    this.year = year;
    const prior = await this.client.getRange('ENTRY', 'B23');
    const lists = await this.client.getTableBody(this.listsTable);
    const categories = lists.values.filter(r => r[1] !== '' && r[1] !== null).map(r => ({ id: Number(r[0]), name: String(r[1]), type: String(r[2]), column: String(r[3] ?? '') }));
    const log = await this._readLog();
    const header = await this.client.getRange(log.sheet, rangeAddress(null, log.firstRow - 1, 0, log.firstRow - 1, LOG_COLUMNS.length - 1));
    const got = header.values[0].map(String);
    if (got.join('|') !== LOG_COLUMNS.join('|')) throw new VerificationError('headers', `The ${this.table} columns are ${got.join(', ')}; expected ${LOG_COLUMNS.join(', ')}.`);
    const notices = [];
    const allDup = duplicateIdentities(log.txns);
    const dupIds = allDup.filter(k => !expectedDuplicates.includes(k));
    if (dupIds.length && !this.halted) {
      const copies = log.txns.filter(t => dupIds.includes(identityOf(t)));
      this._halt(new UnresolvedOperationError('integrity', `The transaction table holds more than one copy of ${dupIds.length === 1 ? 'a record' : `${dupIds.length} records`} (${[...new Set(copies.map(t => `#${t.id}`))].join(', ')}). Nothing will be written until this is checked.`, { copies: copies.map(publicTxn), identities: dupIds }));
    }
    const numbers = new Map();
    for (const k of new Set(log.txns.map(identityOf))) { const id = Number(k.split('|')[0]); numbers.set(id, (numbers.get(id) || 0) + 1); }
    const dupNumbers = [...numbers.entries()].filter(([, n]) => n > 1).map(([id]) => id);
    if (dupNumbers.length) notices.push(`Transaction numbers used more than once in the workbook: ${dupNumbers.map(n => `#${n}`).join(', ')}. The app tells them apart by their timestamps; the treasurer may want to renumber them in Excel.`);
    return { year, priorYearBalance: num(prior.values[0][0]), categories, transactions: log.txns, nextId: nextIdFrom(log.txns), meta, loadedAt: new Date().toISOString(), sortedStale: this.sortedStale, incident: this.halted, notices, pendingDuplicates: allDup.filter(k => expectedDuplicates.includes(k)) };
  }

  // ---- incidents ------------------------------------------------------------------------------
  /**
   * READ-ONLY. Re-reads the table and checks the conditions under which the current incident counts as
   * resolved. Returns { resolved, inconclusive, checks: [{ name, ok: true|false|null, detail }], incident }.
   * A verified resolution lifts the pause; an unverifiable one (ok === null somewhere) can only be acknowledged.
   */
  async _verifyResolution() {
    if (!this.incident) return { resolved: true, inconclusive: false, checks: [], incident: null };
    if (this.incident.resolved) return { resolved: true, inconclusive: false, checks: this.incident.checks || [], incident: this.incident };
    const log = await this._readLog();
    const d = this.incident.details || {};
    const checks = [];
    const check = (name, ok, detail) => checks.push({ name, ok, detail });
    const copiesOf = ident => log.txns.filter(t => identityOf(t) === ident);
    const dupIds = duplicateIdentities(log.txns);
    check('No record appears twice in the table', dupIds.length === 0, dupIds.length ? `still duplicated: ${dupIds.map(k => `#${k.split('|')[0]}`).join(', ')}` : 'every record appears once');
    switch (this.incident.kind) {
      case 'wrong-row-deleted':
        if (d.victim) {
          const n = copiesOf(identityOf(d.victim)).length;
          check(`Transaction #${d.victim.id} is present again`, n === 1, n === 1 ? 'present once' : n === 0 ? 'still missing' : 'present more than once');
        } else {
          check('The removed row is identified', null, 'the app could not tell which row was removed; compare with Excel version history');
        }
        if (d.duplicateCopy && d.identity) check(`Transaction #${d.id} has exactly one copy`, copiesOf(d.identity).length === 1, `${copiesOf(d.identity).length} copies`);
        break;
      case 'duplicate-copy':
        if (d.identity) check(`Transaction #${d.id ?? d.identity.split('|')[0]} has exactly one copy`, copiesOf(d.identity).length === 1, `${copiesOf(d.identity).length} copies`);
        break;
      case 'duplicate-marker': {
        const n = log.txns.filter(t => t.timestamp === d.marker).length;
        check('The operation id appears at most once', n <= 1, `${n} rows carry it`);
        break;
      }
      case 'duplicate-number': {
        const n = log.txns.filter(t => t.id === d.number).length;
        check(`Number #${d.number} is used by one row`, n <= 1, `${n} rows use it`);
        break;
      }
      case 'wrong-row-renumbered': {
        if (d.intended) {
          const me = log.txns.filter(t => t.timestamp === d.intended.timestamp);
          check(`Your entry (operation ${d.intended.timestamp}) is present with a unique number`, me.length === 1 && log.txns.filter(t => t.id === me[0].id).length === 1, me.length === 1 ? `now #${me[0].id}` : `${me.length} rows`);
        }
        if (d.possiblyAffected) {
          const row = log.txns.find(t => t.timestamp === d.possiblyAffected.timestamp);
          check(`Transaction ${d.possiblyAffected.timestamp} has its own number back`, Boolean(row) && row.id === d.previousNumber, row ? `now #${row.id}` : 'missing');
        }
        break;
      }
      default: break;   // 'integrity' and unknown kinds: the duplicate-identity check above decides
    }
    const failed = checks.some(c => c.ok === false), unknown = checks.some(c => c.ok === null);
    const now = new Date().toISOString();
    this.incident.checks = checks;
    this.incident.checkedAt = now;
    const resolved = !failed && !unknown;
    if (resolved) { this.incident.resolved = true; this.incident.resolvedAt = now; this.incident.resolution = 'verified'; this.logFn({ op: 'incident-resolved', kind: this.incident.kind }); }
    return { resolved, inconclusive: !failed && unknown, checks, incident: this.incident };
  }
  /** A member confirms, after checking the workbook (version history), that the incident is dealt with. Allowed only when verification cannot decide. */
  async _acknowledgeIncident({ confirm = false, by = '' } = {}) {
    if (!confirm) throw new Error('acknowledgeIncident needs explicit confirmation.');
    if (!this.incident || this.incident.resolved) return this.incident;
    const v = await this._verifyResolution();
    if (v.resolved) return this.incident;
    if (!v.inconclusive) throw new Error(`This cannot be acknowledged away while the table still shows a problem: ${v.checks.filter(c => c.ok === false).map(c => c.name).join('; ')}. Repair the table first.`);
    const now = new Date().toISOString();
    this.incident.resolved = true; this.incident.resolvedAt = now; this.incident.resolution = 'acknowledged'; this.incident.acknowledgedBy = by;
    this.logFn({ op: 'incident-acknowledged', kind: this.incident.kind });
    return this.incident;
  }

  // ---- append ------------------------------------------------------------------------------------
  async _applyFormats(rowIndex) {
    try { await this.client.patchTableRowRange(this.table, rowIndex, { numberFormat: [FORMAT_ROW] }); return null; }
    catch (e) { this.logFn({ op: 'formats-failed', message: e.message }); return FORMAT_WARNING; }
  }
  /** Rows carrying `marker`: none -> null; one with matching content -> it; one with other content -> collision; several -> unresolved. */
  _markerRow(log, marker, values) {
    const same = log.txns.filter(t => t.timestamp === marker);
    if (same.length > 1) throw this._halt(new UnresolvedOperationError('duplicate-marker', `Two rows in the workbook carry the same operation id (${marker}); one is a duplicate of the other. Nothing more will be written until this is checked.`, { marker, rows: same.map(publicTxn) }));
    if (same.length === 1 && values && !payloadMatches(same[0].values, values)) throw new MarkerCollisionError(marker, publicTxn(same[0]));
    return same[0] || null;
  }
  /**
   * Makes the transaction number of the just-added row `txn` unique. The renumbering PATCH is positional, so each
   * round re-reads the table and sends the PATCH only if our row (by operation id) still sits at that index.
   * Returns { txn, log, renumberedFrom }; records an incident and throws if it cannot.
   */
  async _ensureUniqueNumber(log, txn) {
    let cur = txn, curLog = log, renumberedFrom = null;
    for (let round = 0; round < 3; round++) {
      if (!curLog.txns.some(t => t.id === cur.id && t.timestamp !== cur.timestamp)) return { txn: cur, log: curLog, renumberedFrom };
      const fresh = nextIdFrom(curLog.txns);
      if (renumberedFrom === null) renumberedFrom = cur.id;
      const pre = await this._readLog();                                       // evaluated BEFORE the positional PATCH
      const at = pre.rows[cur.rowIndex];
      if (!at || String(at.values[1]) !== cur.timestamp) {                     // our row is no longer at that index: aim again, touch nothing
        const mine = pre.txns.filter(t => t.timestamp === cur.timestamp);
        if (mine.length !== 1) throw this._halt(new UnresolvedOperationError('duplicate-marker', `The row with operation id ${cur.timestamp} could not be located uniquely while fixing its number.`, { marker: cur.timestamp, rows: mine.map(publicTxn) }));
        cur = mine[0]; curLog = pre;
        continue;
      }
      try { await this.client.patchTableRowRange(this.table, cur.rowIndex, { values: [[fresh, null, null, null, null, null, null, null, null]] }); }
      catch (e) { if (!(e instanceof ExcelApiError && e.ambiguous)) throw e; }
      curLog = await this._readLog();
      const mine = curLog.txns.filter(t => t.timestamp === cur.timestamp);
      if (mine.length !== 1) throw this._halt(new UnresolvedOperationError('duplicate-marker', `The row with operation id ${cur.timestamp} could not be located uniquely after renumbering.`, { marker: cur.timestamp, rows: mine.map(publicTxn) }));
      // A pre-existing row that now carries `fresh` was renumbered by the shift, not by us.
      const victim = curLog.txns.find(t => t.id === fresh && t.timestamp !== cur.timestamp && pre.txns.some(p => p.timestamp === t.timestamp && p.id !== fresh));
      if (victim) {
        const was = pre.txns.find(p => p.timestamp === victim.timestamp);
        throw this._halt(new UnresolvedOperationError('wrong-row-renumbered', `While fixing a duplicate transaction number the rows moved, and transaction #${was.id} (${was.date}, ${was.description}) received number ${fresh} instead. Nothing more will be written until this is checked.`, { intended: publicTxn(cur), possiblyAffected: publicTxn(victim), previousNumber: was.id, number: fresh }));
      }
      cur = mine[0];
    }
    if (curLog.txns.some(t => t.id === cur.id && t.timestamp !== cur.timestamp)) {
      throw this._halt(new UnresolvedOperationError('duplicate-number', `The entry is in the workbook, but its transaction number #${cur.id} is also used by another entry and could not be made unique. Nothing more will be written until this is checked.`, { txn: publicTxn(cur), number: cur.id }));
    }
    return { txn: cur, log: curLog, renumberedFrom };
  }
  async _finishAppend(log, txn, alreadyExisted) {
    const warnings = [];
    const fmt = await this._applyFormats(txn.rowIndex);
    if (fmt) warnings.push(fmt);
    const u = await this._ensureUniqueNumber(log, txn);
    const sorted = await this._rebuildSorted();
    if (!sorted.consistent) warnings.push(SORTED_WARNING);
    const vis = await this._syncMonthVisibility([monthOfIso(u.txn.date)]);
    return { txn: u.txn, alreadyExisted, renumberedFrom: u.renumberedFrom, warnings, reportFormatting: pendingReportFormatting(vis, REPORT_FORMATTING_SAVED), formattedMonths: vis.months };
  }

  /** Submit Entry. `marker` is the operation id (makeMarker()); it makes checks and retries safe. */
  async _addTransaction(entry, { marker, user = '' } = {}) {
    this._guard();
    if (!marker || typeof marker !== 'string') throw new Error('An operation id (marker) is required for every add.');
    const values = entryToRow(entry, marker);
    let log = await this._readLog();
    const existing = this._markerRow(log, marker, values);
    if (existing) return this._finishAppend(log, existing, true);
    values[0] = nextIdFrom(log.txns);
    this.logFn({ op: 'add', id: values[0], user });
    await this.client.addTableRows(this.table, [values]);          // an ambiguous failure propagates: the caller inspects / resumes
    log = await this._readLog();
    const mine = this._markerRow(log, marker, values);
    if (!mine) throw new VerificationError('append-readback', `The entry was accepted but could not be found afterwards (operation ${marker}). Press Refresh; if it is missing, save again.`);
    return this._finishAppend(log, mine, false);
  }
  /**
   * READ-ONLY: what became of an add with this operation id?
   * { outcome: 'missing' } | { outcome: 'conflict', error } | { outcome: 'incomplete', txn, missingSteps, message } | { outcome: 'landed', result }
   */
  async _inspectAppend(marker, entry) {
    const values = entryToRow(entry, marker);
    const log = await this._readLog();
    const same = log.txns.filter(t => t.timestamp === marker);
    if (same.length === 0) return { outcome: 'missing' };
    if (same.length > 1) throw this._halt(new UnresolvedOperationError('duplicate-marker', `Two rows in the workbook carry the same operation id (${marker}); one is a duplicate of the other. Nothing more will be written until this is checked.`, { marker, rows: same.map(publicTxn) }));
    const row = same[0];
    if (!payloadMatches(row.values, values)) { const error = new MarkerCollisionError(marker, publicTxn(row)); return { outcome: 'conflict', error, message: error.message }; }
    const steps = [];
    if (log.txns.some(t => t.id === row.id && t.timestamp !== row.timestamp)) steps.push('a unique transaction number');
    const range = await this.client.getTableRowRange(this.table, row.rowIndex);
    if (!rowFormatsOk(range.numberFormat && range.numberFormat[0])) steps.push('the number formats of the new row');
    const sorted = await this.client.getTableBody(this.sortedTable);
    if (!sameMatrix(sorted.values.map(normalizeRow), sortedTarget(log))) steps.push('the sorted helper table the monthly sheets read');
    if (steps.length) return { outcome: 'incomplete', txn: publicTxn(row), missingSteps: steps, message: `Your entry is in the workbook as #${row.id}, but these finishing steps are not confirmed: ${steps.join('; ')}. Press Save to finish them; the entry itself is not written twice.` };
    return { outcome: 'landed', result: { txn: row, alreadyExisted: true, renumberedFrom: null, warnings: [] } };
  }
  /** Finishes an add that landed (formats, number, sorted table). Null if nothing landed. May write the finishing steps. */
  async _completeAppend(marker, entry) {
    this._guard();
    const values = entryToRow(entry, marker);
    const log = await this._readLog();
    const existing = this._markerRow(log, marker, values);
    if (!existing) return null;
    return this._finishAppend(log, existing, true);
  }
  async findAppended(marker) {
    const log = await this._readLog();
    return log.txns.find(t => t.timestamp === marker) || null;
  }

  // ---- delete / correct ----------------------------------------------------------------------
  _locate(log, ref) {
    const matches = log.txns.filter(t => t.id === ref.id && t.timestamp === ref.timestamp);
    if (!matches.length) throw new ConflictError('missing', { expected: ref });
    const t = matches.find(m => m.rowIndex === ref.rowIndex) || matches[0];
    if (t.fingerprint !== ref.fingerprint) throw new ConflictError('changed', { expected: ref, current: t });
    return t;
  }
  /**
   * Checked positional delete: a fresh read is evaluated first, and the DELETE is sent only if the row at the
   * index is exactly `cur` (identity + content). If the row has moved, the read is repeated (up to three times).
   * Returns { outcome: 'sent', pre, cur, ambiguous, error? } | { outcome: 'gone', pre } | { outcome: 'moving' } | { outcome: 'failed', pre, error }.
   */
  async _checkedDelete(cur) {
    for (let round = 0; round < 3; round++) {
      const pre = await this._readLog();
      const at = pre.rows[cur.rowIndex];
      const stillThere = at && Number(at.values[0]) === cur.id && String(at.values[1]) === cur.timestamp && rowFingerprint(at.values) === cur.fingerprint;
      if (!stillThere) {
        const now = pre.txns.find(t => identityOf(t) === identityOf(cur) && t.fingerprint === cur.fingerprint);
        if (!now) return { outcome: 'gone', pre };
        cur = now;                                                          // it moved: aim at the new index after another fresh read
        continue;
      }
      try { await this.client.deleteTableRow(this.table, cur.rowIndex); }
      catch (e) {
        if (!(e instanceof ExcelApiError)) throw e;
        if (e.ambiguous) return { outcome: 'sent', pre, cur, ambiguous: true, error: e };
        return { outcome: 'failed', pre, error: e };
      }
      return { outcome: 'sent', pre, cur, ambiguous: false, error: null };
    }
    return { outcome: 'moving' };
  }
  /**
   * After a delete: did exactly our target disappear? `pre` is the fresh read evaluated just before the DELETE.
   * Never restores anything. Returns { deleted, notices } or records an incident and throws.
   */
  _assessDelete({ log, cur, pre, post, ambiguous, purpose, extra = {} }) {
    const ident = identityOf(cur);
    const isTarget = t => identityOf(t) === ident && t.fingerprint === cur.fingerprint;
    const preTxns = pre ? pre.txns : log.txns;
    const has = (list, t) => list.some(q => identityOf(q) === identityOf(t) && (identityOf(t) !== ident || q.fingerprint === t.fingerprint));
    const targetStill = post.txns.some(isTarget);
    const removedByOthers = log.txns.filter(t => !isTarget(t) && !has(preTxns, t)).map(publicTxn);
    const missing = preTxns.filter(t => !isTarget(t) && !has(post.txns, t)).map(publicTxn);
    const halt = (victim, why) => this._halt(new UnresolvedOperationError('wrong-row-deleted',
      `${why} ${victim ? `The row that was removed is transaction #${victim.id} (${victim.date}, ${victim.description}, ${victim.amount}).` : 'The row that was removed could not be identified.'} It was NOT recreated automatically. Nothing more will be written until this is checked; a member can restore it explicitly from the details shown.`,
      { victim, victimLastKnownOnly: true, targetStillPresent: targetStill, removedByOthers, missing, duplicateCopy: purpose === 'correction' && targetStill, ...extra }));
    if (targetStill && missing.length === 0) return { deleted: false, notices: [] };
    if (targetStill) throw halt(missing.length === 1 ? missing[0] : null, 'The rows moved between the check and the delete, so the delete hit a different row.');
    const notices = [];
    if (removedByOthers.length) notices.push(`While you were working, someone else removed: ${removedByOthers.map(describe).join('; ')}.`);
    if (missing.length) notices.push(`Also missing since the delete: ${missing.map(describe).join('; ')}. Someone else may have deleted them; check Excel's version history if that is unexpected.`);
    if (ambiguous) notices.push('The answer to the delete was lost; the read-back confirmed the row is gone.');
    return { deleted: true, notices };
  }

  /** Delete a transaction identified by a reference from load(). */
  async _deleteTransaction(ref) {
    this._guard();
    const log = await this._readLog();
    const cur = this._locate(log, ref);
    const r = await this._checkedDelete(cur);
    if (r.outcome === 'gone') throw new ConflictError('missing', { expected: ref, message: `Transaction #${cur.id} was removed by someone else just now; nothing was deleted by you.` });
    if (r.outcome === 'moving') throw new VerificationError('delete-unconfirmed', `The table kept changing while the delete of #${cur.id} was being checked; nothing was deleted. Try again.`);
    if (r.outcome === 'failed') throw r.error;
    const post = await this._readLog();
    const outcome = this._assessDelete({ log, cur: r.cur, pre: r.pre, post, ambiguous: r.ambiguous, purpose: 'delete' });
    if (!outcome.deleted) throw new VerificationError('delete-unconfirmed', `The delete of #${cur.id} did not reach the workbook; nothing was changed. Try again.`);
    const sorted = await this._rebuildSorted();
    if (!sorted.consistent) outcome.notices.push(SORTED_WARNING);
    const vis = await this._syncMonthVisibility([monthOfIso(cur.date)]);
    return { deleted: true, notices: outcome.notices, reportFormatting: pendingReportFormatting(vis, REPORT_FORMATTING_DELETED), formattedMonths: vis.months };
  }

  /** Correct a transaction: append the corrected copy (same id and timestamp), then delete the old copy with the checked delete. */
  async _updateTransaction(ref, changes) {
    this._guard();
    const log = await this._readLog();
    const cur = this._locate(log, ref);
    const merged = mergeChanges(cur, changes);
    if (sameRow(merged, cur.values)) return cur;
    try { await this.client.addTableRows(this.table, [merged]); }
    catch (e) { if (!(e instanceof ExcelApiError && e.ambiguous)) throw e; }          // ambiguous: the read below decides
    return this._finishCorrection(await this._readLog(), cur, rowFingerprint(merged));
  }
  /**
   * READ-ONLY: what became of a correction of `ref` with `changes`?
   * { outcome: 'missing' } | { outcome: 'conflict', error } | { outcome: 'incomplete', txn, missingSteps, message } | { outcome: 'landed', result }
   */
  async _inspectCorrection(ref, changes) {
    const merged = mergeChanges(ref, changes);
    if (sameRow(merged, ref.values)) return { outcome: 'landed', result: { ...ref, warnings: [], notices: [] } };
    const newFp = rowFingerprint(merged), ident = identityOf(ref);
    const log = await this._readLog();
    const copies = log.txns.filter(t => identityOf(t) === ident);
    const fresh = copies.filter(t => t.fingerprint === newFp), olds = copies.filter(t => t.fingerprint === ref.fingerprint);
    if (fresh.length === 0) {
      if (copies.length === 1 && olds.length === 1) return { outcome: 'missing' };
      if (copies.length === 0) { const error = new ConflictError('missing', { expected: ref }); return { outcome: 'conflict', error, message: error.message }; }
      const error = new ConflictError('changed', { expected: ref, current: copies[0] });
      return { outcome: 'conflict', error, message: error.message };
    }
    if (fresh.length > 1 || copies.length > fresh.length + olds.length) {
      throw this._halt(new UnresolvedOperationError('duplicate-copy', `Transaction #${ref.id} has ${copies.length} copies in the workbook. Nothing more will be written until this is checked.`, { identity: ident, id: ref.id, timestamp: ref.timestamp, oldFingerprint: ref.fingerprint, newFingerprint: newFp, copies: copies.map(publicTxn) }));
    }
    const steps = [];
    if (olds.length) steps.push('removal of the old copy');
    const range = await this.client.getTableRowRange(this.table, fresh[0].rowIndex);
    if (!rowFormatsOk(range.numberFormat && range.numberFormat[0])) steps.push('the number formats of the corrected row');
    const sorted = await this.client.getTableBody(this.sortedTable);
    if (!sameMatrix(sorted.values.map(normalizeRow), sortedTarget(log))) steps.push('the sorted helper table the monthly sheets read');
    if (steps.length) return { outcome: 'incomplete', txn: publicTxn(fresh[0]), missingSteps: steps, message: `Your correction to #${ref.id} is in the workbook, but these finishing steps are not confirmed: ${steps.join('; ')}. Press Save to finish them; the correction is not written twice.` };
    return { outcome: 'landed', result: { ...fresh[0], warnings: [], notices: [] } };
  }
  /** Finishes a correction that landed (removes the old copy if still present, formats, sorted table). Null if nothing landed. */
  async _completeCorrection(ref, changes) {
    this._guard();
    const merged = mergeChanges(ref, changes);
    if (sameRow(merged, ref.values)) return ref;                                        // nothing needed to be written
    const newFp = rowFingerprint(merged);
    const log = await this._readLog();
    if (!log.txns.some(t => identityOf(t) === identityOf(ref) && t.fingerprint === newFp)) return null;
    return this._finishCorrection(log, ref, newFp, { resumed: true });
  }
  async _finishCorrection(log2, cur, newFp, { resumed = false } = {}) {
    const ident = identityOf(cur), oldFp = cur.fingerprint;
    const detail = { identity: ident, id: cur.id, timestamp: cur.timestamp, oldFingerprint: oldFp, newFingerprint: newFp };
    const warnings = [];
    const copies = log2.txns.filter(t => identityOf(t) === ident);
    const fresh = copies.filter(t => t.fingerprint === newFp), olds = copies.filter(t => t.fingerprint === oldFp);
    if (fresh.length === 0) throw new VerificationError('correction-append', `The correction to #${cur.id} did not reach the workbook; nothing was changed. Try again.`);
    if (fresh.length > 1 || copies.length > fresh.length + olds.length) {
      throw this._halt(new UnresolvedOperationError('duplicate-copy', `Transaction #${cur.id} now has ${copies.length} copies in the workbook. Nothing more will be written until this is checked.`, { ...detail, copies: copies.map(publicTxn) }));
    }
    const fmt = await this._applyFormats(fresh[0].rowIndex);
    if (fmt) warnings.push(fmt);
    if (olds.length === 0) {
      const sorted = await this._rebuildSorted();
      if (!sorted.consistent) warnings.push(SORTED_WARNING);
      const vis = await this._syncMonthVisibility([monthOfIso(cur.date), monthOfIso(fresh[0].date)]);
      return { ...fresh[0], warnings, notices: resumed ? [] : ['This record was deleted by someone else while you corrected it; your corrected copy is now the only one.'], reportFormatting: pendingReportFormatting(vis, REPORT_FORMATTING_SAVED), formattedMonths: vis.months };
    }
    const old = olds[0];
    const duplicate = (why, post) => this._halt(new UnresolvedOperationError('duplicate-copy', `Your correction to #${cur.id} is saved, but the old copy could not be removed (${why}). Two copies exist. Nothing more will be written until this is checked; the old copy can be removed explicitly from the details shown.`, { ...detail, copies: (post || log2).txns.filter(t => identityOf(t) === ident).map(publicTxn) }));
    const r = await this._checkedDelete(old);
    if (r.outcome === 'moving') throw duplicate('the table kept changing while the old copy was being located');
    if (r.outcome === 'failed') throw duplicate(r.error.message);
    const post = await this._readLog();
    if (r.outcome === 'sent') {
      const outcome = this._assessDelete({ log: log2, cur: r.cur, pre: r.pre, post, ambiguous: r.ambiguous, purpose: 'correction', extra: detail });
      if (!outcome.deleted) throw duplicate('the delete did not reach the workbook', post);
      const sorted = await this._rebuildSorted();
      if (!sorted.consistent) warnings.push(SORTED_WARNING);
      const result = post.txns.find(t => identityOf(t) === ident && t.fingerprint === newFp) || fresh[0];
      const vis = await this._syncMonthVisibility([monthOfIso(cur.date), monthOfIso(result.date)]);
      return { ...result, warnings, notices: outcome.notices, reportFormatting: pendingReportFormatting(vis, REPORT_FORMATTING_SAVED), formattedMonths: vis.months };
    }
    // 'gone': someone removed the old copy between our read and the checked delete
    if (post.txns.filter(t => identityOf(t) === ident).length !== 1) throw duplicate('the copies could not be reconciled', post);
    const sorted = await this._rebuildSorted();
    if (!sorted.consistent) warnings.push(SORTED_WARNING);
    const result = post.txns.find(t => identityOf(t) === ident && t.fingerprint === newFp) || fresh[0];
    const vis = await this._syncMonthVisibility([monthOfIso(cur.date), monthOfIso(result.date)]);
    return { ...result, warnings, notices: ['The old copy of this record was removed by someone else at the same time.'], reportFormatting: pendingReportFormatting(vis, REPORT_FORMATTING_SAVED), formattedMonths: vis.months };
  }

  // ---- monthly report row visibility -----------------------------------------------------------
  /**
   * Makes each monthly report sheet show exactly its populated rows, by the same rule the workbook's own
   * `refreshMonthlyVisibility()` uses: a row between 4 and 33 is shown when its date cell holds a value and
   * hidden when that cell is blank. Without this, a transaction added from the website lands in a row an
   * earlier script run had hidden, so the month sheet looks right on the website but misses the row when it is
   * opened or printed in Excel.
   *
   * It changes row visibility and nothing else. It never unprotects a sheet, never reads or sets a protection
   * password, and never writes a cell value: the month sheets are protected with "format rows" allowed, so a
   * plain `PATCH …/range(address='A4:A9') { rowHidden }` is permitted while content writes stay blocked
   * (docs/workbook-mapping.md §8). It never touches LOG_Table either, so however many times it runs it cannot
   * add, change or remove a transaction.
   *
   * Returns { ok, months, runs, error } and never throws into a save: a transaction that reached the workbook
   * stays saved even when its report formatting could not be finished.
   */
  async _syncMonthVisibility(months, { recalculate = true } = {}) {
    const wanted = [...new Set((months || []).map(monthOfIso_or_number).filter(m => m >= 1 && m <= 12))].sort((a, b) => a - b);
    if (!wanted.length) return { ok: true, months: [], runs: 0 };
    const expected = MONTH_LAST_ROW - MONTH_FIRST_ROW + 1;
    let runs = 0;
    try {
      // The report rows are formulas over LOG_Sorted; recalculate once so what is read reflects the change.
      if (recalculate) await this.client.calculate('Full');
      for (const m of wanted) {
        const name = MONTH_NAMES[m - 1];
        const col = await this.client.getRange(name, `A${MONTH_FIRST_ROW}:A${MONTH_LAST_ROW}`, 'address,values');
        const populated = (col.values || []).map(r => !(r[0] === '' || r[0] === null || r[0] === undefined));
        if (populated.length !== expected) return { ok: false, months: wanted, error: `The ${name} sheet returned ${populated.length} report rows instead of ${expected}.` };
        for (const run of visibilityRuns(populated)) {
          const address = `A${run.first}:A${run.last}`;
          await this.client.patchRange(name, address, { rowHidden: run.hidden });
          runs += 1;
          const back = await this.client.getRange(name, address, 'address,rowHidden');   // verify what was set
          if (back.rowHidden !== run.hidden) {
            const got = back.rowHidden === null || back.rowHidden === undefined ? 'a mix of both' : back.rowHidden ? 'hidden' : 'shown';
            return { ok: false, months: wanted, error: `${name} rows ${run.first} to ${run.last} should be ${run.hidden ? 'hidden' : 'shown'}; the workbook reports ${got}.` };
          }
        }
      }
    } catch (e) {
      return { ok: false, months: wanted, error: e.message };
    }
    return { ok: true, months: wanted, runs };
  }

  // ---- explicit repairs (member-confirmed, allowed while paused) ----------------------------
  /**
   * Removes one specific copy (identity + exact content) after a member confirmed it. When several IDENTICAL
   * copies exist, `allowIdentical` removes exactly one of them (which one does not matter) and proves the count.
   */
  async _removeCopy(identity, fingerprint, { confirm = false, allowIdentical = false } = {}) {
    if (!confirm) throw new Error('removeCopy needs explicit confirmation.');
    const log = await this._readLog();
    const isCopy = t => identityOf(t) === identity && t.fingerprint === fingerprint;
    const matches = log.txns.filter(isCopy);
    if (matches.length === 0) return { removed: false, reason: 'not-found' };
    if (matches.length > 1 && !allowIdentical) return { removed: false, reason: 'multiple' };
    const target = matches[matches.length - 1];
    const r = await this._checkedDelete(target);
    if (r.outcome === 'gone') return { removed: false, reason: 'not-found' };
    if (r.outcome === 'moving') return { removed: false, reason: 'moving' };
    if (r.outcome === 'failed') throw r.error;
    const post = await this._readLog();
    if (matches.length === 1) {
      const outcome = this._assessDelete({ log, cur: r.cur, pre: r.pre, post, ambiguous: r.ambiguous, purpose: 'repair' });
      if (!outcome.deleted) return { removed: false, reason: 'not-deleted' };
      await this._rebuildSorted();
      return { removed: true, notices: outcome.notices };
    }
    // identical copies: the count must drop by exactly one and no other row may be missing
    const lostOthers = r.pre.txns.filter(t => !isCopy(t) && !post.txns.some(q => identityOf(q) === identityOf(t) && q.fingerprint === t.fingerprint)).map(publicTxn);
    if (lostOthers.length) {
      throw this._halt(new UnresolvedOperationError('wrong-row-deleted', `Removing a duplicate copy of #${target.id} hit a different row: transaction #${lostOthers[0].id} (${lostOthers[0].date}, ${lostOthers[0].description}, ${lostOthers[0].amount}). It was NOT recreated automatically. Nothing more will be written until this is checked.`, { victim: lostOthers[0], victimLastKnownOnly: true, targetStillPresent: true, removedByOthers: [], missing: lostOthers, duplicateCopy: false }));
    }
    if (post.txns.filter(isCopy).length !== matches.length - 1) return { removed: false, reason: 'not-deleted' };
    await this._rebuildSorted();
    return { removed: true, notices: [] };
  }
  /** Re-appends a row from values a member has seen and confirmed (e.g. the row named in an incident). */
  async _reappendRow(values, { confirm = false } = {}) {
    if (!confirm) throw new Error('reappendRow needs explicit confirmation.');
    const v = normalizeRow(values);
    const ident = identityOfValues(v);
    let log = await this._readLog();
    const existing = log.txns.find(t => identityOf(t) === ident);
    if (existing) return existing;
    try { await this.client.addTableRows(this.table, [v]); }
    catch (e) { if (!(e instanceof ExcelApiError && e.ambiguous)) throw e; }
    log = await this._readLog();
    const back = log.txns.filter(t => identityOf(t) === ident);
    if (back.length !== 1) throw this._halt(new UnresolvedOperationError('duplicate-copy', `Restoring #${v[0]} left ${back.length} copies.`, { identity: ident, id: v[0], copies: back.map(publicTxn) }));
    await this._applyFormats(back[0].rowIndex);
    await this._rebuildSorted();
    return back[0];
  }

  // ---- sorted helper table ------------------------------------------------------------------
  /** Makes LOG_Sorted_Table equal to LOG_Table sorted by Date, re-checking against a fresh LOG read each round. */
  async _rebuildSorted() {
    let changed = false;
    for (let round = 0; round < 3; round++) {
      const log = await this._readLog();
      const want = sortedTarget(log);
      const body = await this.client.getTableBody(this.sortedTable);
      const current = body.values.map(normalizeRow);
      if (sameMatrix(current, want)) { this.sortedStale = false; return { changed, rows: want.length, consistent: true }; }
      changed = true;
      const p = parseAddress(body.address);
      const m = current.length, n = want.length, k = Math.min(m, n);
      try {
        if (k > 0) await this.client.patchRange(p.sheet, rangeAddress(null, p.r1, p.c1, p.r1 + k - 1, p.c1 + LOG_COLUMNS.length - 1), { values: want.slice(0, k), numberFormat: want.slice(0, k).map(() => FORMAT_ROW) });
        if (n > m) {
          await this.client.addTableRows(this.sortedTable, want.slice(m));
          await this.client.patchRange(p.sheet, rangeAddress(null, p.r1 + m, p.c1, p.r1 + n - 1, p.c1 + LOG_COLUMNS.length - 1), { numberFormat: want.slice(m).map(() => FORMAT_ROW) });
        } else {
          for (let i = m - 1; i >= n; i--) await this.client.deleteTableRow(this.sortedTable, i);
        }
      } catch (e) { this.logFn({ op: 'sorted-write-failed', message: e.message }); }
    }
    const log = await this._readLog();
    const body = await this.client.getTableBody(this.sortedTable);
    const consistent = sameMatrix(body.values.map(normalizeRow), sortedTarget(log));
    this.sortedStale = !consistent;
    return { changed, rows: log.txns.length, consistent };
  }

  // ---- reads for the screens ------------------------------------------------------------------
  async readMonth(month) {
    const name = MONTH_NAMES[Number(month) - 1];
    if (!name) throw new Error(`Bad month ${month}`);
    const r = await this.client.getRange(name, 'A2:O40');
    const v = r.values;
    const rows = [];
    for (let i = 2; i <= 31; i++) {
      const row = v[i] || [];
      if (row[0] === '' || row[0] === null || row[0] === undefined) continue;
      const cells = {};
      CATEGORIES.forEach((c, j) => { cells[c.name] = row[5 + j] === '' ? null : num(row[5 + j]); });
      rows.push({ date: isSerialDate(row[0]) ? serialToIso(row[0]) : String(row[0]), description: String(row[1] ?? ''), amount: num(row[2]), chequeNum: row[3] === '' ? '' : String(row[3]), cells, net: num(row[14]) });
    }
    const sub = v[32] || [];
    const byCategory = {};
    CATEGORIES.forEach((c, j) => { byCategory[c.name] = num(sub[5 + j]); });
    return { month: Number(month), name, title: String(v[0]?.[0] ?? ''), rows, subtotal: { amount: num(sub[2]), byCategory, net: num(sub[14]) },
      recon: { opening: num(v[35]?.[1]), deposits: num(v[36]?.[1]), expenses: num(v[37]?.[1]), closing: num(v[38]?.[1]) }, truncated: rows.length >= 30 };
  }
  async readAnnual() {
    const year = this.year || Number((await this.client.getRange('ConfigHidden', 'A1')).values[0][0]);
    const r = await this.client.getRange(`Annual ${year}`, 'A3:N19');
    return { year, header: r.values[0], rows: r.values.slice(1).filter(row => row[0] !== '' && row[0] !== null) };
  }
  async _setPriorYearBalance(value, expectedCurrent) {
    this._guard();
    const v = Math.round(Number(String(value).replace(/[$,\s]/g, '')) * 100) / 100;
    if (!Number.isFinite(v)) throw new Error('Enter a valid amount.');
    const cur = num((await this.client.getRange('ENTRY', 'B23')).values[0][0]);
    if (toCents(cur) !== toCents(expectedCurrent)) throw new ConflictError('changed', { expected: expectedCurrent, current: cur, message: `The prior-year balance in the workbook is now ${cur}, not ${expectedCurrent}.` });
    await this.client.patchRange('ENTRY', 'B23', { values: [[v]] });
    const back = num((await this.client.getRange('ENTRY', 'B23')).values[0][0]);
    if (toCents(back) !== toCents(v)) throw new VerificationError('prior-readback', `ENTRY!B23 read back as ${back} instead of ${v}.`);
    return v;
  }

  // ---- operation serialisation: one COMPLETE operation at a time per page (review finding T1) ------
  /**
   * Runs fn after every earlier operation of this adapter has finished, so the reads inside fn see fresh row
   * positions and no two positional requests are prepared against the same stale layout. Internal helpers call
   * the unlocked _implementations and never re-enter the queue, so nothing deadlocks. This serialises one page
   * only: other devices, tabs (see tab-guard.js), Excel and the old scripts are outside it.
   */
  _exclusive(fn) {
    this._pending += 1;
    const run = this._chain.then(() => fn()).finally(() => { this._pending -= 1; });
    this._chain = run.catch(() => {});
    return run;
  }
  /** True while an operation of this adapter (read or write) is queued or running. */
  get busy() { return this._pending > 0; }
  load(opts) { return this._exclusive(() => this._load(opts)); }
  verifyResolution() { return this._exclusive(() => this._verifyResolution()); }
  acknowledgeIncident(opts) { return this._exclusive(() => this._acknowledgeIncident(opts)); }
  addTransaction(entry, opts) { return this._exclusive(() => this._addTransaction(entry, opts)); }
  inspectAppend(marker, entry) { return this._exclusive(() => this._inspectAppend(marker, entry)); }
  completeAppend(marker, entry) { return this._exclusive(() => this._completeAppend(marker, entry)); }
  deleteTransaction(ref) { return this._exclusive(() => this._deleteTransaction(ref)); }
  updateTransaction(ref, changes) { return this._exclusive(() => this._updateTransaction(ref, changes)); }
  inspectCorrection(ref, changes) { return this._exclusive(() => this._inspectCorrection(ref, changes)); }
  completeCorrection(ref, changes) { return this._exclusive(() => this._completeCorrection(ref, changes)); }
  removeCopy(identity, fingerprint, opts) { return this._exclusive(() => this._removeCopy(identity, fingerprint, opts)); }
  reappendRow(values, opts) { return this._exclusive(() => this._reappendRow(values, opts)); }
  rebuildSorted() { return this._exclusive(() => this._rebuildSorted()); }
  syncMonthVisibility(months, opts) { return this._exclusive(() => this._syncMonthVisibility(months, opts)); }
  setPriorYearBalance(value, expectedCurrent) { return this._exclusive(() => this._setPriorYearBalance(value, expectedCurrent)); }

  // ---- self-test (test copy only) -------------------------------------------------------------
  /** Adds one labelled row, reads it back, deletes it, and compares the touched areas with their starting state. */
  async selfTest({ confirmTestCopy = false, onStep } = {}) {
    if (!confirmTestCopy) throw new Error('The self-test only runs on a test copy. Confirm that this workbook is the test copy first.');
    if (this._selfTestRunning) return { ok: false, steps: [], restored: false, error: 'The connection test is already running; wait for it to finish.', compared: [], notCompared: [] };
    this._selfTestRunning = true;
    try { return await this._selfTest({ onStep }); } finally { this._selfTestRunning = false; }
  }
  async _selfTest({ onStep } = {}) {
    const steps = [];
    const step = async (name, fn) => {
      try { const detail = await fn(); steps.push({ name, ok: true, detail: detail ?? '' }); if (onStep) onStep(steps.at(-1)); return detail; }
      catch (e) { steps.push({ name, ok: false, detail: e.message }); if (onStep) onStep(steps.at(-1)); throw e; }
    };
    const year = this.year || Number((await this.client.getRange('ConfigHidden', 'A1')).values[0][0]);
    const compared = ['LOG values and number formats', 'LOG_Sorted values and number formats', 'ENTRY!B23', 'January formulas (A4:O40)', `Annual ${year} formulas (A3:N19)`, 'table row counts'];
    const notCompared = ['other worksheets', 'formatting outside the compared ranges', 'table styles, names, protection and validation', 'workbook-level features',
      'row visibility on the monthly report sheets: the test row is added to December and removed again, and this app leaves December\'s report rows matching its contents'];
    const state = async () => {
      const log = await this.client.getTableBody(this.table);
      const sorted = await this.client.getTableBody(this.sortedTable);
      const prior = num((await this.client.getRange('ENTRY', 'B23')).values[0][0]);
      const jan = await this.client.getRange('January', 'A4:O40', 'address,formulas');
      const annual = await this.client.getRange(`Annual ${year}`, 'A3:N19', 'address,formulas');
      return { log: { values: log.values.map(normalizeRow), formats: log.numberFormat, rows: log.rowCount }, sorted: { values: sorted.values.map(normalizeRow), formats: sorted.numberFormat, rows: sorted.rowCount }, prior, jan: jan.formulas, annual: annual.formulas };
    };
    try {
      let snap = null;
      await step('Read the workbook', async () => { snap = await this.load(); if (this.halted) throw new Error(`The workbook has an unresolved problem: ${this.halted.message}`); return `${snap.transactions.length} transactions, year ${snap.year}, next id ${snap.nextId}`; });
      const before = await step('Record the starting state', async () => state());
      const marker = makeMarker();
      const deposit = snap.categories.find(c => c.type === 'Deposit') || { name: 'Membership' };
      const added = await step('Add one clearly labelled test row', () => this.addTransaction({ date: `${snap.year}-12-31`, type: 'Deposit', category: deposit.name, amount: 0.01, description: SELF_TEST_DESCRIPTION, chequeNum: '', notes: 'temporary row written by the connection test' }, { marker }));
      await step('Read the test row back', async () => { const f = await this.findAppended(marker); if (!f) throw new Error('The test row was not found after saving.'); return `row id ${f.id}, date ${f.date}, amount ${f.amount}`; });
      await step('Check the sorted table includes it', async () => { const body = await this.client.getTableBody(this.sortedTable); if (!body.values.some(r => r[1] === marker)) throw new Error('LOG_Sorted does not contain the test row.'); return 'present'; });
      await step('Delete the test row', () => this.deleteTransaction(added.txn));
      await step('Confirm it is gone', async () => { if (await this.findAppended(marker)) throw new Error('The test row is still present.'); return 'removed'; });
      const restored = await step('Compare with the starting state', async () => {
        const after = await state();
        if (JSON.stringify(after) !== JSON.stringify(before)) throw new Error(`The compared areas differ from their starting state (${compared.join(', ')}).`);
        return true;
      });
      return { ok: steps.every(s => s.ok), steps, restored: restored === true, compared, notCompared };
    } catch (e) {
      return { ok: false, steps, restored: false, error: e.message, compared, notCompared };
    }
  }
}
