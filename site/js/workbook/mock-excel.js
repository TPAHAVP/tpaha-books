// In-memory stand-in for the Microsoft Graph Excel API, driven by the same URLs the real client uses.
// It mirrors the ledger workbook: LOG_Table / LOG_Sorted_Table / LISTS_Categories tables, ENTRY!B23,
// ConfigHidden!A1, and month + annual sheets computed exactly like the workbook's formulas
// (rows from LOG_Sorted, subtotals from LOG). Used by unit tests, browser tests and the site's test mode.
import { CATEGORIES, DEPOSIT_CATEGORIES, WITHDRAWAL_CATEGORIES, MONTH_NAMES, toCents, fromCents } from '../ledger/model.js';
import { parseAddress, rangeAddress, indexToCol, serialToIso, formatText } from './excel-values.js';

const MONEY = '$#,##0.00';
const LOG_FORMATS = ['General', 'yyyy-mm-dd hh:mm:ss', 'yyyy-mm-dd', 'General', 'General', MONEY, 'General', 'General', 'General'];

function graphError(status, code, innerCode, message, extraHeaders = {}) {
  const body = { error: { code, message: message || code, innerError: { code: innerCode, 'request-id': `mock-${Date.now()}`, date: new Date().toISOString() } } };
  return new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json', 'request-id': body.error.innerError['request-id'], ...extraHeaders } });
}
const ok = (obj, status = 200) => obj === null ? new Response(null, { status: 204 }) : new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json', 'request-id': `mock-${Date.now()}` } });
const clone = o => JSON.parse(JSON.stringify(o));

export class MockWorkbook {
  constructor(fixture, { name = 'TPAHA_2026 (1).xlsx' } = {}) {
    this.fixture = clone(fixture);
    this.name = name;
    this.sheets = new Map();
    this.tables = new Map();
    this.sessions = new Set();
    this._sessionCounter = 0;
    this.slowSessionCreation = 0;
    this._operations = new Map();
    this.failures = [];
    this.log = [];
    this.beforeRespond = null;
    this.hidden = new Map();     // sheet name -> Set of hidden row numbers
    // As read from the real workbook on 2026-09-12: the month sheets are protected with row formatting allowed.
    this.protection = new Map(MONTH_NAMES.map(n => [n, { protected: true, options: { allowFormatRows: true, allowFormatCells: false, allowFormatColumns: false } }]));
    this.onChange = null;        // called after every mutation (test-mode persistence)
    this.lastModified = '2026-09-06T01:23:54Z';
    this.eTag = 1;
    this._build();
  }

  // ---- construction ----------------------------------------------------------------------
  _sheet(name, computed = false) {
    if (!this.sheets.has(name)) this.sheets.set(name, { name, cells: new Map(), computed });
    return this.sheets.get(name);
  }
  _key(r, c) { return `${r}:${c}`; }
  _get(sheet, r, c) { return sheet.cells.get(this._key(r, c)) || { v: '', nf: 'General' }; }
  _set(sheet, r, c, v, nf) {
    const cur = sheet.cells.get(this._key(r, c)) || { v: '', nf: 'General' };
    sheet.cells.set(this._key(r, c), { v: v === null || v === undefined ? '' : v, nf: nf === undefined ? cur.nf : nf });
  }
  _addTable(name, sheetName, header, rows, formats) {
    const sheet = this._sheet(sheetName);
    header.forEach((h, c) => this._set(sheet, 1, c, h, 'General'));
    rows.forEach((row, i) => row.forEach((v, c) => this._set(sheet, 2 + i, c, v, formats ? formats[c] : 'General')));
    this.tables.set(name, { name, sheet: sheetName, headerRow: 1, colStart: 0, colCount: header.length, bodyRows: rows.length, formats: formats || header.map(() => 'General') });
  }
  _build() {
    const f = this.fixture;
    this._set(this._sheet('ConfigHidden'), 1, 0, f.year, '0');
    this._set(this._sheet('ENTRY'), 23, 1, f.priorYearBalance, '"$"#,##0.00');
    this._addTable('LISTS_Categories', 'LISTS', ['CategoryID', 'CategoryName', 'Type', 'ColumnLetter'], f.categories, null);
    this._addTable('LOG_Table', 'LOG', f.logHeader, f.logRows, LOG_FORMATS);
    this._addTable('LOG_Sorted_Table', 'LOG_Sorted', f.logHeader, f.logSortedRows, LOG_FORMATS);
    for (const m of MONTH_NAMES) this._sheet(m, true);
    this._sheet(`Annual ${f.year}`, true);
  }

  // ---- test helpers ------------------------------------------------------------------------
  table(name) {
    const t = this.tables.get(name);
    if (!t) throw new Error(`No table ${name}`);
    const self = this;
    return {
      get rows() { return self._bodyValues(t); },
      setCell(i, c, v) { self._set(self.sheets.get(t.sheet), t.headerRow + 1 + i, t.colStart + c, v); self._touch(); },
      append(values) { self._appendRows(t, [values]); self._touch(); },
      remove(i) { self._deleteRow(t, i); self._touch(); },
    };
  }
  setCell(sheetName, address, value) { const a = parseAddress(address); this._set(this._sheet(sheetName), a.r1, a.c1, value); this._touch(); }
  /** Hidden row numbers on a sheet, ascending. */
  /** Test hook: model a month sheet re-protected with row formatting disallowed. */
  setAllowFormatRows(sheetName, allowed) {
    const p = this.protection.get(sheetName) || { protected: true, options: {} };
    this.protection.set(sheetName, { ...p, options: { ...p.options, allowFormatRows: allowed } });
  }
  hiddenRows(sheetName) { return [...(this.hidden.get(sheetName) || new Set())].sort((a, b) => a - b); }
  setHiddenRows(sheetName, rows) { this.hidden.set(sheetName, new Set(rows)); this._touch(); }
  getCell(sheetName, address) { const a = parseAddress(address); return this._get(this._sheet(sheetName), a.r1, a.c1).v; }
  failNext(spec) { this.failures.push(spec); }
  expireSession() { this.sessions.clear(); }
  snapshot() {
    const out = {};
    for (const [name, s] of this.sheets) if (!s.computed) out[name] = [...s.cells.entries()].sort(([a], [b]) => a.localeCompare(b));
    out.__tables = [...this.tables.values()].map(t => ({ name: t.name, bodyRows: t.bodyRows }));
    return out;
  }
  _touch() { this.eTag += 1; this.lastModified = new Date().toISOString(); if (this.onChange) { try { this.onChange(this); } catch { /* ignore */ } } }
  /** Plain-data copy of every stored cell and table size; importState(exportState()) reproduces the workbook. */
  exportState() {
    const sheets = {};
    for (const [name, s] of this.sheets) if (!s.computed) sheets[name] = [...s.cells.entries()];
    const hidden = {};
    for (const [name, set] of this.hidden) hidden[name] = [...set];
    return { sheets, hidden, tables: [...this.tables.values()].map(t => ({ name: t.name, bodyRows: t.bodyRows })), eTag: this.eTag, lastModified: this.lastModified };
  }
  importState(state) {
    if (!state || !state.sheets) return this;
    for (const [name, entries] of Object.entries(state.sheets)) { const s = this._sheet(name); s.cells = new Map(entries.map(([k, c]) => [k, { ...c }])); }
    for (const [name, rows] of Object.entries(state.hidden || {})) this.hidden.set(name, new Set(rows));
    for (const t of state.tables || []) { const table = this.tables.get(t.name); if (table) table.bodyRows = t.bodyRows; }
    if (state.eTag) this.eTag = state.eTag;
    if (state.lastModified) this.lastModified = state.lastModified;
    return this;
  }

  // ---- table helpers -----------------------------------------------------------------------
  _bodyValues(t) {
    const sheet = this.sheets.get(t.sheet);
    const rows = [];
    for (let i = 0; i < t.bodyRows; i++) rows.push(Array.from({ length: t.colCount }, (_, c) => this._get(sheet, t.headerRow + 1 + i, t.colStart + c).v));
    return rows;
  }
  _bodyAddress(t) { return rangeAddress(t.sheet, t.headerRow + 1, t.colStart, t.headerRow + Math.max(t.bodyRows, 1), t.colStart + t.colCount - 1); }
  _appendRows(t, values2D) {
    const sheet = this.sheets.get(t.sheet);
    const firstIndex = t.bodyRows;
    for (const row of values2D) {
      if (row.length !== t.colCount) throw new Error('row width mismatch');
      const r = t.headerRow + 1 + t.bodyRows;
      row.forEach((v, c) => this._set(sheet, r, t.colStart + c, v, 'General'));   // no format inheritance: the adapter must set formats
      t.bodyRows += 1;
    }
    return firstIndex;
  }
  _deleteRow(t, i) {
    if (i < 0 || i >= t.bodyRows) return false;
    const sheet = this.sheets.get(t.sheet);
    for (let k = i; k < t.bodyRows - 1; k++) {
      for (let c = 0; c < t.colCount; c++) {
        const below = this._get(sheet, t.headerRow + 2 + k, t.colStart + c);
        this._set(sheet, t.headerRow + 1 + k, t.colStart + c, below.v, below.nf);
      }
    }
    for (let c = 0; c < t.colCount; c++) sheet.cells.delete(this._key(t.headerRow + t.bodyRows, t.colStart + c));
    t.bodyRows -= 1;
    return true;
  }
  _tableAt(sheetName, r, c) {
    for (const t of this.tables.values()) if (t.sheet === sheetName && c >= t.colStart && c < t.colStart + t.colCount) return t;
    return null;
  }

  // ---- computed sheets (mirror the workbook formulas) --------------------------------------
  _txnsFrom(tableName) {
    const t = this.tables.get(tableName);
    return this._bodyValues(t).map((r, i) => ({ i, values: r })).filter(({ values }) => values[2] !== '' && values[2] !== null && values[2] !== undefined)
      .map(({ i, values: r }) => ({ index: i, id: r[0], timestamp: r[1], date: typeof r[2] === 'number' ? serialToIso(r[2]) : String(r[2]), serial: typeof r[2] === 'number' ? r[2] : null,
        type: r[3], category: r[4], amountCents: toCents(Number(r[5]) || 0), description: r[6], chequeNum: r[7], notes: r[8] }));
  }
  _monthTotals(month) {
    const year = this.fixture.year;
    const start = `${year}-${String(month).padStart(2, '0')}-01`, end = month === 12 ? `${year + 1}-01-01` : `${year}-${String(month + 1).padStart(2, '0')}-01`;
    const byCat = {}; for (const c of CATEGORIES) byCat[c.name] = 0;
    for (const t of this._txnsFrom('LOG_Table')) if (t.date >= start && t.date < end && t.category in byCat) byCat[t.category] += t.amountCents;
    const dep = DEPOSIT_CATEGORIES.reduce((s, n) => s + byCat[n], 0), wd = WITHDRAWAL_CATEGORIES.reduce((s, n) => s + byCat[n], 0);
    return { byCat, dep, wd, start, end };
  }
  _opening(month) {
    let bal = toCents(Number(this._get(this._sheet('ENTRY'), 23, 1).v) || 0);
    for (let m = 1; m < month; m++) { const t = this._monthTotals(m); bal += t.dep - t.wd; }
    return bal;
  }
  _monthGrid(name) {
    const month = MONTH_NAMES.indexOf(name) + 1;
    const year = this.fixture.year;
    const g = new Map();
    const put = (r, c, v) => g.set(this._key(r, c), v);
    put(2, 0, `${name} ${year} Transactions`); put(2, 5, 'DEPOSITS'); put(2, 9, 'WITHDRAWALS');
    ['Date', 'Description', 'Amount', 'Check#', '', ...DEPOSIT_CATEGORIES, ...WITHDRAWAL_CATEGORIES, '', 'Balance'].forEach((h, c) => put(3, c, h));
    const { byCat, dep, wd, start, end } = this._monthTotals(month);
    // Rows 4..33 come from LOG_Sorted via INDEX(COUNTIFS(date<start)+k): correct only when that table is sorted.
    const sorted = this._txnsFrom('LOG_Sorted_Table');
    const before = sorted.filter(t => t.date < start).length;
    const inMonth = sorted.filter(t => t.date >= start && t.date < end).length;
    let amountSum = 0;
    for (let k = 0; k < 30; k++) {
      const r = 4 + k;
      const t = k < inMonth ? sorted[before + k] : null;
      if (!t) { for (let c = 0; c < 15; c++) put(r, c, ''); continue; }
      put(r, 0, t.serial ?? t.date); put(r, 1, t.description); put(r, 2, fromCents(t.amountCents)); put(r, 3, t.chequeNum ?? ''); put(r, 4, '');
      amountSum += t.amountCents;
      let net = 0;
      CATEGORIES.forEach((c, j) => { const hit = t.category === c.name; put(r, 5 + j, hit ? fromCents(t.amountCents) : ''); if (hit) net = DEPOSIT_CATEGORIES.includes(c.name) ? t.amountCents : -t.amountCents; });
      put(r, 13, ''); put(r, 14, fromCents(net));
    }
    put(34, 0, 'SUBTOTAL'); put(34, 2, fromCents(amountSum));
    CATEGORIES.forEach((c, j) => put(34, 5 + j, fromCents(byCat[c.name])));
    put(34, 14, fromCents(dep - wd));
    put(36, 0, 'BANK RECONCILIATION');
    const opening = this._opening(month);
    put(37, 0, 'Opening Balance'); put(37, 1, fromCents(opening));
    put(38, 0, 'Add: Total Deposits'); put(38, 1, fromCents(dep));
    put(39, 0, 'Less: Total Expenses'); put(39, 1, fromCents(wd));
    put(40, 0, 'Closing Balance'); put(40, 1, fromCents(opening + dep - wd));
    return g;
  }
  _annualGrid() {
    const year = this.fixture.year;
    const g = new Map(); const put = (r, c, v) => g.set(this._key(r, c), v);
    put(1, 0, `${year} Financial Summary`);
    ['Category', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Year Total'].forEach((h, c) => put(3, c, h));
    const months = Array.from({ length: 12 }, (_, i) => this._monthTotals(i + 1));
    const row = (r, label, fn) => { put(r, 0, label); let total = 0; months.forEach((m, i) => { const v = fn(m); total += v; put(r, 1 + i, fromCents(v)); }); put(r, 13, fromCents(total)); };
    DEPOSIT_CATEGORIES.forEach((c, i) => row(4 + i, c, m => m.byCat[c]));
    row(8, 'Total Income', m => m.dep);
    WITHDRAWAL_CATEGORIES.forEach((c, i) => row(10 + i, c, m => m.byCat[c]));
    row(14, 'Total Expenses', m => m.wd);
    row(16, 'Net', m => m.dep - m.wd);
    const opening = toCents(Number(this._get(this._sheet('ENTRY'), 23, 1).v) || 0);
    put(18, 0, `Year Opening Balance (Jan 1, ${year})`); put(18, 1, fromCents(opening));
    const net = months.reduce((s, m) => s + m.dep - m.wd, 0);
    put(19, 0, `Year Closing Balance (Dec 31, ${year})`); put(19, 1, fromCents(opening + net));
    return g;
  }
  _computedCell(sheetName, r, c) {
    const grid = MONTH_NAMES.includes(sheetName) ? this._monthGrid(sheetName) : this._annualGrid();
    const v = grid.get(this._key(r, c));
    const nf = MONTH_NAMES.includes(sheetName) && r >= 4 && r <= 33 && c === 0 && typeof v === 'number' ? 'dd/mm/yyyy' : (typeof v === 'number' ? '"$"#,##0.00' : 'General');
    return { v: v === undefined ? '' : v, nf };
  }

  // ---- ranges ------------------------------------------------------------------------------
  _rangeJson(sheetName, a) {
    const sheet = this.sheets.get(sheetName);
    const values = [], text = [], numberFormat = [], formulas = [];
    for (let r = a.r1; r <= a.r2; r++) {
      const vr = [], tr = [], fr = [], ff = [];
      for (let c = a.c1; c <= a.c2; c++) {
        const cell = sheet.computed ? this._computedCell(sheetName, r, c) : this._get(sheet, r, c);
        vr.push(cell.v); fr.push(cell.nf); tr.push(formatText(cell.v, cell.nf));
        ff.push(sheet.computed && cell.v !== '' ? `=FORMULA(${sheetName}!${indexToCol(c)}${r})` : cell.v);
      }
      values.push(vr); text.push(tr); numberFormat.push(fr); formulas.push(ff);
    }
    const hiddenSet = this.hidden.get(sheetName) || new Set();
    let nHidden = 0;
    for (let r = a.r1; r <= a.r2; r++) if (hiddenSet.has(r)) nHidden += 1;
    const rows = a.r2 - a.r1 + 1;
    const rowHidden = nHidden === rows ? true : nHidden === 0 ? false : null;
    return { address: rangeAddress(sheetName, a.r1, a.c1, a.r2, a.c2), addressLocal: rangeAddress(sheetName, a.r1, a.c1, a.r2, a.c2), values, text, numberFormat, formulas, rowHidden,
      rowIndex: a.r1 - 1, columnIndex: a.c1, rowCount: a.r2 - a.r1 + 1, columnCount: a.c2 - a.c1 + 1 };
  }
  _patchRange(sheetName, a, body) {
    const sheet = this.sheets.get(sheetName);
    if (!sheet) return graphError(404, 'ItemNotFound', 'itemNotFound', `Worksheet ${sheetName} not found`);
    // Row visibility is a "format rows" change. The month sheets' protection allows it (checked against the
    // real workbook), so it is permitted here even on the protected, formula-driven sheets. Content is not.
    if (body.rowHidden !== undefined) {
      if (typeof body.rowHidden !== 'boolean') return graphError(400, 'BadRequest', 'invalidArgument', 'rowHidden must be a boolean');
      const prot = this.protection.get(sheetName);
      if (prot && prot.protected && prot.options && prot.options.allowFormatRows === false) {
        return graphError(403, 'AccessDenied', 'accessDenied', 'The worksheet is protected and does not allow formatting rows.');
      }
      const set = this.hidden.get(sheetName) || new Set();
      for (let r = a.r1; r <= a.r2; r++) { if (body.rowHidden) set.add(r); else set.delete(r); }
      this.hidden.set(sheetName, set);
      this._touch();
    }
    const touchesContent = ['values', 'formulas', 'numberFormat', 'columnHidden'].some(k => body[k] !== undefined);
    if (!touchesContent) return ok(this._rangeJson(sheetName, a));
    if (sheet.computed) return graphError(403, 'AccessDenied', 'accessDenied', 'The worksheet is protected; formula cells cannot be changed.');
    if (body.formulas) return graphError(400, 'BadRequest', 'invalidArgument', 'Formula writes are not permitted in this mock.');
    const rows = a.r2 - a.r1 + 1, cols = a.c2 - a.c1 + 1;
    for (const key of ['values', 'numberFormat']) {
      if (!body[key]) continue;
      if (body[key].length !== rows || body[key].some(r => r.length !== cols)) return graphError(400, 'BadRequest', 'invalidArgument', `${key} must match the range size`);
    }
    for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) {
      const t = this._tableAt(sheetName, r, c);
      if (t && (r === t.headerRow || r > t.headerRow + t.bodyRows)) return graphError(400, 'BadRequest', 'invalidArgument', `Cell ${indexToCol(c)}${r} is outside the table body; use rows/add`);
    }
    for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) {
      const v = body.values ? body.values[r - a.r1][c - a.c1] : null;
      const nf = body.numberFormat ? body.numberFormat[r - a.r1][c - a.c1] : null;
      if (v !== null && v !== undefined) this._set(sheet, r, c, v);
      if (nf !== null && nf !== undefined) this._set(sheet, r, c, this._get(sheet, r, c).v, nf);
    }
    this._touch();
    return ok(this._rangeJson(sheetName, a));
  }

  // ---- HTTP surface ------------------------------------------------------------------------
  async fetch(url, opts = {}) {
    const method = (opts.method || 'GET').toUpperCase();
    const u = new URL(url);
    const headers = new Headers(opts.headers || {});
    const sessionId = headers.get('workbook-session-id');
    const entry = { method, url: decodeURIComponent(u.pathname) + (u.search ? decodeURIComponent(u.search) : ''), sessionId, body: opts.body || null, at: Date.now() };
    this.log.push(entry);
    if (this.beforeRespond) await this.beforeRespond(entry);
    const path = entry.url;
    const wb = path.indexOf('/workbook');
    const rel = wb >= 0 ? path.slice(wb) : path.replace(/^.*\/items\/[^/?]+/, '');
    const body = opts.body ? JSON.parse(opts.body) : {};

    const matches = (f, p) => !f.match || (f.match instanceof RegExp ? f.match.test(p) : p.includes(String(f.match)));
    const injected = this.failures.findIndex(f => matches(f, path));
    const failure = injected >= 0 ? this.failures.splice(injected, 1)[0] : null;
    const fail = () => {
      if (failure.networkError) throw new TypeError('Failed to fetch');
      return graphError(failure.status || 500, failure.code || 'InternalServerError', failure.innerCode || 'internalServerErrorUncategorized', failure.message, failure.retryAfter ? { 'Retry-After': String(failure.retryAfter) } : {});
    };
    if (failure && !failure.afterApply) return fail();
    const respond = res => (failure && failure.afterApply ? fail() : res);

    // JSON batch: run sub-requests in order; a failed dependency yields 424 for dependants
    if (u.pathname.endsWith('/$batch') && method === 'POST') {
      const results = new Map();
      for (const r of body.requests || []) {
        const deps = r.dependsOn || [];
        if (deps.some(d => { const dr = results.get(String(d)); return !dr || dr.status >= 400; })) {
          results.set(String(r.id), { id: r.id, status: 424, headers: {}, body: { error: { code: 'FailedDependency', message: 'A request this one depends on failed.', innerError: { code: 'failedDependency' } } } });
          continue;
        }
        const subRes = await this.fetch(`${u.origin}/v1.0${r.url}`, { method: r.method, headers: r.headers || {}, body: r.body !== undefined ? JSON.stringify(r.body) : undefined });
        const txt = await subRes.text();
        let parsed = null;
        try { parsed = txt ? JSON.parse(txt) : null; } catch { parsed = txt; }
        results.set(String(r.id), { id: r.id, status: subRes.status, headers: Object.fromEntries(subRes.headers.entries()), body: parsed });
      }
      return respond(ok({ responses: (body.requests || []).map(r => results.get(String(r.id))) }));
    }

    // item metadata (no /workbook)
    if (wb < 0) {
      if (method === 'GET') return respond(ok({ id: 'I1', name: this.name, size: 104054, lastModifiedDateTime: this.lastModified, webUrl: 'https://example.invalid/mock-workbook', eTag: `"${this.eTag}"`, file: { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }, parentReference: { driveId: 'D1', path: '/drive/root:/Documents' } }));
      return graphError(405, 'MethodNotAllowed', 'methodNotAllowedUncategorized');
    }
    // sessions
    if (rel === '/workbook/createSession') {
      if (this.slowSessionCreation > 0) {
        const opId = `op${this._operations.size + 1}`;
        this._operations.set(opId, { remaining: this.slowSessionCreation });
        this.slowSessionCreation = 0;
        return new Response('{}', { status: 202, headers: { Location: `${u.origin}${u.pathname.replace(/\/workbook\/createSession$/, '')}/workbook/operations/${opId}`, 'Content-Type': 'application/json' } });
      }
      const id = `session-${++this._sessionCounter}`;
      this.sessions.add(id);
      return respond(ok({ id, persistChanges: body.persistChanges !== false }, 201));
    }
    if (/^\/workbook\/operations\//.test(rel)) {
      const opId = rel.split('/').pop();
      const op = this._operations.get(opId);
      if (!op) return graphError(404, 'ItemNotFound', 'itemNotFound');
      if (op.remaining > 0) { op.remaining -= 1; return ok({ id: opId, status: 'running' }); }
      return ok({ id: opId, status: 'succeeded', resourceLocation: `${u.origin}${u.pathname.replace(/\/workbook\/operations\/.*$/, '')}/workbook/sessionInfoResource(key='${opId}')` });
    }
    if (/^\/workbook\/sessionInfoResource/.test(rel)) {
      const id = `session-${++this._sessionCounter}`;
      this.sessions.add(id);
      return ok({ id, persistChanges: true });
    }
    if (rel === '/workbook/closeSession') { if (sessionId) this.sessions.delete(sessionId); return ok(null); }
    if (sessionId && !this.sessions.has(sessionId)) return graphError(404, 'ItemNotFound', 'invalidSessionReCreatable', 'The workbook session does not exist or has expired.');
    if (rel === '/workbook/application/calculate') return respond(ok(null));

    const prot = /^\/workbook\/worksheets\/([^/]+)\/protection$/.exec(rel.split('?')[0]);
    if (prot && method === 'GET') {
      const name = prot[1];
      if (!this.sheets.has(name)) return graphError(404, 'ItemNotFound', 'itemNotFound', `Worksheet ${name} not found`);
      const p = this.protection.get(name) || { protected: false, options: {} };
      return respond(ok({ protected: p.protected, options: p.options }));
    }

    // tables
    let m = /^\/workbook\/tables\/([^/]+)(?:\/(.*))?$/.exec(rel.split('?')[0]);
    if (m) {
      const t = this.tables.get(m[1]);
      if (!t) return graphError(404, 'ItemNotFound', 'itemNotFound', `Table ${m[1]} not found`);
      const sub = m[2] || '';
      if (sub === 'dataBodyRange' && method === 'GET') return respond(ok(this._rangeJson(t.sheet, parseAddress(this._bodyAddress(t)))));
      if (sub === 'rows' && method === 'GET') return respond(ok({ value: this._bodyValues(t).map((values, index) => ({ index, values })) }));
      if (sub === 'rows/add' && method === 'POST') {
        if (!Array.isArray(body.values) || !body.values.length) return graphError(400, 'BadRequest', 'invalidArgument', 'values required');
        const index = this._appendRows(t, body.values); this._touch();
        return respond(ok({ index, values: body.values }));
      }
      const rowM = /^rows\/(?:itemAt\(index=(\d+)\)|(\d+))(?:\/(range))?$/.exec(sub);
      if (rowM) {
        const idx = Number(rowM[1] ?? rowM[2]);
        if (idx < 0 || idx >= t.bodyRows) return graphError(400, 'BadRequest', 'invalidArgument', `Row index ${idx} out of range`);
        const r = t.headerRow + 1 + idx;
        const a = { r1: r, r2: r, c1: t.colStart, c2: t.colStart + t.colCount - 1 };
        if (rowM[3] === 'range') {
          if (method === 'GET') return respond(ok(this._rangeJson(t.sheet, a)));
          if (method === 'PATCH') return respond(this._patchRange(t.sheet, a, body));
        } else {
          if (method === 'DELETE') { this._deleteRow(t, idx); this._touch(); return respond(new Response('', { status: 200 })); }
          if (method === 'GET') return respond(ok({ index: idx, values: [this._bodyValues(t)[idx]] }));
        }
      }
      return graphError(404, 'ItemNotFound', 'itemNotFound', `Unsupported table path ${sub}`);
    }
    // worksheet ranges
    m = /^\/workbook\/worksheets\/([^/]+)\/range\(address='([^']+)'\)$/.exec(rel.split('?')[0]);
    if (m) {
      const sheetName = m[1].replace(/^'|'$/g, '');
      if (!this.sheets.has(sheetName)) return graphError(404, 'ItemNotFound', 'itemNotFound', `Worksheet ${sheetName} not found`);
      const a = parseAddress(m[2]);
      if (method === 'GET') return respond(ok(this._rangeJson(sheetName, a)));
      if (method === 'PATCH') return respond(this._patchRange(sheetName, a, body));
    }
    return graphError(404, 'ItemNotFound', 'itemNotFound', `Unsupported path ${rel}`);
  }
}
