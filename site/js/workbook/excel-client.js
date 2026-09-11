// Microsoft Graph Excel API client for one workbook.
// - One persistent session (workbook-session-id), recreated when Graph says it is re-creatable.
// - Strictly one request in flight per client instance (Microsoft's guidance). This does not
//   coordinate other tabs, devices, members, or edits made directly in Excel.
// - Only GET requests may be retried (once) after 504, 429/503 with Retry-After, a network error, or an
//   expired session. Every mutation (PATCH, POST, DELETE, batch containing one) is attempted exactly
//   once; when its outcome cannot be known the error carries `ambiguous: true` so the caller reads
//   the workbook before deciding anything.
const GRAPH = 'https://graph.microsoft.com/v1.0';
const enc = s => encodeURIComponent(String(s));

export class ExcelApiError extends Error {
  constructor({ message, status = 0, code = '', innerCode = '', retryAfter = null, requestId = null, ambiguous = false, sessionInvalid = false, conflict = false, method = '', path = '' }) {
    super(message || `${method} ${path} failed (${status || 'network'})`);
    this.name = 'ExcelApiError';
    Object.assign(this, { status, code, innerCode, retryAfter, requestId, ambiguous, sessionInvalid, conflict, method, path });
  }
}

/** Reads Microsoft's error envelope: top-level code, deepest innerError code, Retry-After. */
export function classifyError(status, body, headers) {
  const err = (body && body.error) || {};
  let inner = err.innerError;
  let innerCode = '';
  while (inner) { if (inner.code) innerCode = inner.code; inner = inner.innerError; }
  const ra = headers && headers.get ? headers.get('Retry-After') : (headers && headers['Retry-After']);
  const retryAfter = ra !== null && ra !== undefined && ra !== '' && Number.isFinite(Number(ra)) ? Number(ra) : null;
  const sessionInvalid = /^invalidSession/i.test(innerCode);
  const recreatable = /^invalidSessionReCreatable$/i.test(innerCode);
  const conflict = status === 409 || /accessConflict|conflictUncategorized|insertDeleteConflict|filteredRangeConflict/i.test(innerCode);
  const transient = [502, 503, 504, 429].includes(status) || /gatewayTimeout|transientFailure|serviceUnavailable|tooManyRequests/i.test(innerCode);
  return { code: err.code || '', innerCode, message: err.message || '', retryAfter, sessionInvalid, recreatable, conflict, retryableRead: transient || recreatable };
}

async function readJson(res) {
  const text = await res.text().catch(() => '');
  if (!text) return null;
  try { return JSON.parse(text); } catch { return { raw: text }; }
}

export class ExcelClient {
  constructor({ getToken, driveId, itemId, fetchImpl, sleep, log }) {
    if (!driveId || !itemId) throw new Error('ExcelClient needs driveId and itemId');
    this.getToken = getToken;
    this.driveId = driveId;
    this.itemId = itemId;
    this.fetch = fetchImpl || ((...a) => globalThis.fetch(...a));
    this.sleep = sleep || (ms => new Promise(r => setTimeout(r, ms)));
    this.logFn = log || (() => {});
    this.itemPath = `/drives/${enc(driveId)}/items/${enc(itemId)}`;
    this.base = GRAPH + this.itemPath;
    this.sessionId = null;
    this._queue = Promise.resolve();
  }

  _enqueue(fn) {
    const run = this._queue.then(fn, fn);
    this._queue = run.then(() => {}, () => {});
    return run;
  }

  // ---- sessions --------------------------------------------------------------------------
  async _authHeaders(extra = {}) {
    return { Authorization: `Bearer ${await this.getToken()}`, Accept: 'application/json', ...extra };
  }
  async _openSessionUnqueued() {
    for (let attempt = 0; attempt < 2; attempt++) {
      let res;
      try {
        res = await this.fetch(`${this.base}/workbook/createSession`, {
          method: 'POST', headers: await this._authHeaders({ 'Content-Type': 'application/json', Prefer: 'respond-async' }), body: JSON.stringify({ persistChanges: true }),
        });
      } catch (e) {
        if (attempt === 0) continue;
        throw new ExcelApiError({ message: `Could not reach Microsoft 365 to open the workbook: ${e.message}`, status: 0, method: 'POST', path: '/workbook/createSession' });
      }
      if (res.status === 201) { const j = await readJson(res); this.sessionId = j.id; this.logFn({ event: 'session', id: j.id }); return this.sessionId; }
      if (res.status === 202) { this.sessionId = await this._pollSession(res.headers.get('Location')); return this.sessionId; }
      if (res.status === 504 && attempt === 0) continue;
      const body = await readJson(res);
      const c = classifyError(res.status, body, res.headers);
      throw new ExcelApiError({ message: c.message || 'Could not open a workbook session', status: res.status, code: c.code, innerCode: c.innerCode, retryAfter: c.retryAfter, requestId: res.headers.get('request-id'), method: 'POST', path: '/workbook/createSession' });
    }
    throw new ExcelApiError({ message: 'Could not open a workbook session', method: 'POST', path: '/workbook/createSession' });
  }
  async _pollSession(location) {
    if (!location) throw new ExcelApiError({ message: 'Session creation returned 202 without a Location header', status: 202 });
    for (let i = 0; i < 120; i++) {
      await this.sleep(2000);
      const res = await this.fetch(location, { method: 'GET', headers: await this._authHeaders() });
      const j = await readJson(res);
      if (!res.ok) throw new ExcelApiError({ message: 'Session creation status check failed', status: res.status });
      if (j.status === 'succeeded') {
        const r2 = await this.fetch(j.resourceLocation, { method: 'GET', headers: await this._authHeaders() });
        const info = await readJson(r2);
        if (!r2.ok || !info || !info.id) throw new ExcelApiError({ message: 'Session created but its id could not be read', status: r2.status });
        return info.id;
      }
      if (j.status === 'failed') { const c = classifyError(500, j, new Headers()); throw new ExcelApiError({ message: c.message || 'The workbook could not be opened', status: 500, code: c.code, innerCode: c.innerCode }); }
    }
    throw new ExcelApiError({ message: 'Timed out waiting for the workbook session', status: 202 });
  }
  openSession() { return this._enqueue(async () => { if (!this.sessionId) await this._openSessionUnqueued(); return this.sessionId; }); }
  closeSession() {
    return this._enqueue(async () => {
      if (!this.sessionId) return;
      const id = this.sessionId;
      this.sessionId = null;
      try { await this.fetch(`${this.base}/workbook/closeSession`, { method: 'POST', headers: await this._authHeaders({ 'Content-Type': 'application/json', 'workbook-session-id': id }), body: '{}' }); }
      catch { /* best effort */ }
    });
  }

  // ---- single requests -------------------------------------------------------------------
  /** GET requests may be retried once; everything else is sent exactly once. */
  request(method, path, { body, headers = {}, session = true } = {}) {
    return this._enqueue(() => this._send(method, path, { body, headers, session, idempotent: method === 'GET' }, 0));
  }
  async _send(method, path, opts, attempt) {
    if (opts.session && !this.sessionId) await this._openSessionUnqueued();
    const url = path.startsWith('http') ? path : this.base + path;
    const h = await this._authHeaders(opts.headers);
    if (opts.session && this.sessionId) h['workbook-session-id'] = this.sessionId;
    if (opts.body !== undefined) h['Content-Type'] = 'application/json';
    const started = Date.now();
    let res;
    try {
      res = await this.fetch(url, { method, headers: h, body: opts.body === undefined ? undefined : JSON.stringify(opts.body) });
    } catch (e) {
      this.logFn({ method, path, status: 0, ms: Date.now() - started, error: e.message });
      if (opts.idempotent && attempt === 0) return this._send(method, path, opts, 1);
      throw new ExcelApiError({ message: `Network problem while contacting Microsoft 365 (${e.message}).`, status: 0, ambiguous: !opts.idempotent, method, path });
    }
    const requestId = res.headers.get('request-id');
    this.logFn({ method, path, status: res.status, ms: Date.now() - started, requestId });
    if (res.ok) return readJson(res);
    const body = await readJson(res);
    const c = classifyError(res.status, body, res.headers);
    if (c.recreatable && opts.session) {
      this.sessionId = null;
      if (opts.idempotent && attempt === 0) { await this._openSessionUnqueued(); return this._send(method, path, opts, 1); }
    } else if (opts.idempotent && attempt === 0 && c.retryableRead) {
      if (c.retryAfter) await this.sleep(Math.min(c.retryAfter, 30) * 1000);
      return this._send(method, path, opts, 1);
    }
    throw new ExcelApiError({
      message: c.message || `${method} ${path} failed (${res.status})`, status: res.status, code: c.code, innerCode: c.innerCode, retryAfter: c.retryAfter, requestId,
      ambiguous: !opts.idempotent && res.status >= 500, sessionInvalid: c.sessionInvalid, conflict: c.conflict, method, path,
    });
  }

  /**
   * JSON batch: sub-requests run in order on the server (each depends on the previous one).
   * Returns [{ status, body, headers, error }] in request order. Sent exactly once; a transport or
   * batch-level failure is `ambiguous` when any sub-request mutates.
   */
  batch(items) {
    return this._enqueue(async () => {
      if (!this.sessionId) await this._openSessionUnqueued();
      const mutates = items.some(it => (it.method || 'GET').toUpperCase() !== 'GET');
      const requests = items.map((it, i) => {
        const r = { id: String(i + 1), method: (it.method || 'GET').toUpperCase(), url: this.itemPath + it.path, headers: { 'workbook-session-id': this.sessionId, ...(it.headers || {}) } };
        if (it.body !== undefined) { r.body = it.body; r.headers['Content-Type'] = 'application/json'; }
        if (i > 0) r.dependsOn = [String(i)];
        return r;
      });
      const started = Date.now();
      let res;
      try {
        res = await this.fetch(`${GRAPH}/$batch`, { method: 'POST', headers: await this._authHeaders({ 'Content-Type': 'application/json', 'workbook-session-id': this.sessionId }), body: JSON.stringify({ requests }) });
      } catch (e) {
        this.logFn({ method: 'POST', path: '/$batch', status: 0, ms: Date.now() - started, error: e.message });
        throw new ExcelApiError({ message: `Network problem during a batch request (${e.message}).`, status: 0, ambiguous: mutates, method: 'POST', path: '/$batch' });
      }
      this.logFn({ method: 'POST', path: '/$batch', status: res.status, ms: Date.now() - started, requestId: res.headers.get('request-id') });
      if (!res.ok) {
        const body = await readJson(res);
        const c = classifyError(res.status, body, res.headers);
        throw new ExcelApiError({ message: c.message || `Batch request failed (${res.status})`, status: res.status, code: c.code, innerCode: c.innerCode, retryAfter: c.retryAfter, ambiguous: mutates && res.status >= 500, method: 'POST', path: '/$batch' });
      }
      const body = await readJson(res);
      const byId = new Map(((body && body.responses) || []).map(r => [String(r.id), r]));
      return items.map((it, i) => {
        const r = byId.get(String(i + 1)) || { status: 0, body: null, headers: {} };
        const headers = r.headers || {};
        const error = r.status >= 400 || r.status === 0 ? classifyError(r.status, r.body, new Headers(headers)) : null;
        if (error && /^invalidSession/i.test(error.innerCode)) this.sessionId = null;
        return { status: r.status, body: r.body ?? null, headers, error };
      });
    });
  }

  // ---- helpers ---------------------------------------------------------------------------
  getRange(sheet, address, select = 'address,values,text,numberFormat,rowIndex,columnIndex,rowCount,columnCount') {
    return this.request('GET', `/workbook/worksheets/${enc(sheet)}/range(address='${address}')${select ? `?$select=${select}` : ''}`);
  }
  patchRange(sheet, address, props) {
    return this.request('PATCH', `/workbook/worksheets/${enc(sheet)}/range(address='${address}')`, { body: props });
  }
  tableBodyPath(table, select = 'address,values,text,numberFormat,rowIndex,columnIndex,rowCount,columnCount') {
    return `/workbook/tables/${enc(table)}/dataBodyRange?$select=${select}`;
  }
  getTableBody(table) { return this.request('GET', this.tableBodyPath(table)); }
  addTableRows(table, values2D) { return this.request('POST', `/workbook/tables/${enc(table)}/rows/add`, { body: { values: values2D } }); }
  deleteRowPath(table, index) { return `/workbook/tables/${enc(table)}/rows/${Number(index)}`; }
  deleteTableRow(table, index) { return this.request('DELETE', this.deleteRowPath(table, index)); }
  rowRangePath(table, index) { return `/workbook/tables/${enc(table)}/rows/itemAt(index=${Number(index)})/range`; }
  getTableRowRange(table, index) { return this.request('GET', this.rowRangePath(table, index)); }
  patchTableRowRange(table, index, props) { return this.request('PATCH', this.rowRangePath(table, index), { body: props }); }
  calculate(type = 'Full') { return this.request('POST', '/workbook/application/calculate', { body: { calculationType: type } }); }
  getItemMeta() { return this.request('GET', '?$select=id,name,size,lastModifiedDateTime,webUrl,eTag,file,parentReference', { session: false }); }
}
