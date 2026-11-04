// Microsoft Graph Excel API client for one workbook.
// - One persistent session (workbook-session-id), recreated only when Graph says it is re-creatable
//   (second-level code invalidSessionReCreatable). Any other session error BLOCKS the client: the session
//   is dropped and nothing more is sent until clearBlock() — in the pages, a reload, which builds a new client.
// - Strictly one request in flight per client instance (Microsoft's guidance). This does not
//   coordinate other tabs, devices, members, or edits made directly in Excel.
// - Only GET requests may be retried (once) after 504, 429/503 with Retry-After, a network error, or a
//   re-creatable session error. Every mutation (PATCH, POST, DELETE, batch containing one) is attempted
//   exactly once; when its outcome cannot be known the error carries `ambiguous: true` so the caller reads
//   the workbook before deciding anything. A refused mutation (4xx) is definite: it was not applied.
// - The log it keeps (logFn) carries method, path, status, timing, request id and error codes: never a
//   token, never the session id, never cell contents.
const GRAPH = 'https://graph.microsoft.com/v1.0';
const enc = s => encodeURIComponent(String(s));

export class ExcelApiError extends Error {
  constructor({ message, status = 0, code = '', innerCode = '', innerCodes = [], retryAfter = null, requestId = null, ambiguous = false, sessionInvalid = false, recreatable = false, blocked = false, conflict = false, method = '', path = '' }) {
    super(message || `${method} ${path} failed (${status || 'network'})`);
    this.name = 'ExcelApiError';
    Object.assign(this, { status, code, innerCode, innerCodes, retryAfter, requestId, ambiguous, sessionInvalid, recreatable, blocked, conflict, method, path });
  }
}

/**
 * Reads Microsoft's error envelope. `innerCode` is the SECOND-level code, the one Microsoft documents and says
 * to act on; `innerCodes` is the whole chain. The paused-session envelope ("The target session is invalid.",
 * the refusal of 2026-11-03) nests a third level, InvalidOrTimedOutSession, beneath invalidSessionReCreatable;
 * classifying by the deepest code alone turned that documented, re-creatable error into an unknown 400.
 * The session flags therefore look at every level and at the top-level code.
 */
export function classifyError(status, body, headers) {
  const err = (body && body.error) || {};
  const innerCodes = [];
  for (let inner = err.innerError; inner; inner = inner.innerError) if (inner.code) innerCodes.push(String(inner.code));
  const innerCode = innerCodes[0] || '';
  const all = [String(err.code || ''), ...innerCodes];
  const ra = headers && headers.get ? headers.get('Retry-After') : (headers && headers['Retry-After']);
  const retryAfter = ra !== null && ra !== undefined && ra !== '' && Number.isFinite(Number(ra)) ? Number(ra) : null;
  const sessionInvalid = all.some(c => /^invalidSession/i.test(c));
  const recreatable = innerCodes.some(c => /^invalidSessionReCreatable$/i.test(c));
  const conflict = status === 409 || all.some(c => /accessConflict|conflictUncategorized|insertDeleteConflict|filteredRangeConflict/i.test(c));
  const transient = [502, 503, 504, 429].includes(status) || all.some(c => /gatewayTimeout|transientFailure|serviceUnavailable|tooManyRequests/i.test(c));
  return { code: err.code || '', innerCode, innerCodes, message: err.message || '', retryAfter, sessionInvalid, recreatable, conflict, retryableRead: transient || recreatable };
}

/** One log entry as a line of a request log. No token, no session id, no cell content is ever in an entry. */
export function describeLogEntry(e) {
  const codes = e.innerCodes && e.innerCodes.length > 1 ? e.innerCodes.join(' / ') : e.innerCode;
  const detail = `${e.requestId ? ' request-id ' + e.requestId : ''}${e.outcome ? ' operation ' + e.outcome : ''}${e.code ? ' code ' + e.code : ''}${codes ? ' inner ' + codes : ''}${e.error ? ' — ' + e.error : ''}`;
  if (e.event === 'session') return `session opened${e.status ? ` (${e.status}, ${e.ms} ms)` : ''}${detail}`;
  if (e.event === 'blocked') return `blocked: ${e.code || ''}${e.innerCodes && e.innerCodes.length ? ' / ' + e.innerCodes.join(' / ') : ''}${e.error ? ' — ' + e.error : ''}`;
  if (e.event) return String(e.event);
  const count = e.items !== undefined ? ` (${e.items} request${e.items === 1 ? '' : 's'})` : '';
  return `${e.batch ? '  in batch: ' : ''}${e.method} ${e.path}${count} → ${e.status || 'network error'}${e.ms !== undefined ? ` (${e.ms} ms)` : ''}${detail}`;
}
/** Log fields describing one refused response (codes and message only, never the body's cell contents). */
const errorFields = c => ({ code: c.code, innerCode: c.innerCode, innerCodes: c.innerCodes, error: c.message });
/** Paths of the asynchronous session-creation flow as logged: the operation and session keys are never written down. */
const OPERATION_PATH = '/workbook/operations/…';
const SESSION_INFO_PATH = '/workbook/sessionInfoResource(…)';

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
    this.blocked = null;          // { code, innerCode, innerCodes, message, requestId, at } after a non-recreatable session error
    this._queue = Promise.resolve();
  }

  _enqueue(fn) {
    const run = this._queue.then(fn, fn);
    this._queue = run.then(() => {}, () => {});
    return run;
  }

  // ---- blocking after a session error Microsoft does not call re-creatable ---------------------
  /**
   * Microsoft documents one session error the client may recover from by itself, invalidSessionReCreatable
   * ("The Microsoft Graph client can try to recreate a session and resume the work"). For every other
   * invalidSession* code it says further access to the session is not expected and recreating it is not
   * expected, or might not succeed: authentication, a workbook that cannot be found, a restricted or
   * unsupported workbook, a conflict with another client. Those block this client: the session is dropped,
   * and no request of any kind (single, batch, session creation) is sent until clearBlock(). The pages never
   * call clearBlock(); a reload constructs a new client once the member has dealt with the cause.
   */
  _block(c, requestId, method, path) {
    this.sessionId = null;
    this.blocked = { code: c.code, innerCode: c.innerCode, innerCodes: c.innerCodes || [], message: c.message, requestId, method, path, at: new Date().toISOString() };
    this.logFn({ event: 'blocked', code: c.code, innerCode: c.innerCode, innerCodes: c.innerCodes || [], error: c.message });
  }
  _blockedError(method, path) {
    const b = this.blocked;
    return new ExcelApiError({
      message: `Nothing is being sent to the workbook: an earlier request was refused with a session error this page is not allowed to recover from by itself (${b.innerCode || b.code}: ${b.message}). Reload the page once the cause is resolved.`,
      status: 0, code: b.code, innerCode: b.innerCode, innerCodes: b.innerCodes, requestId: b.requestId, sessionInvalid: true, blocked: true, method, path,
    });
  }
  _refuseIfBlocked(method, path) { if (this.blocked) throw this._blockedError(method, path); }
  /** Deliberate: lifts the block so the next request opens a new session. Not called by the pages (they reload). */
  clearBlock() { this.blocked = null; }

  // ---- sessions --------------------------------------------------------------------------
  async _authHeaders(extra = {}) {
    return { Authorization: `Bearer ${await this.getToken()}`, Accept: 'application/json', ...extra };
  }
  // Every attempt to open a session is logged like any other request (status, timing, request id, codes), so a
  // refusal that happens before any session exists still leaves evidence. Session ids, the Location of the
  // asynchronous operation and the session resource key are never logged.
  async _openSessionUnqueued() {
    const path = '/workbook/createSession';
    this._refuseIfBlocked('POST', path);
    for (let attempt = 0; attempt < 2; attempt++) {
      const started = Date.now();
      let res;
      try {
        res = await this.fetch(`${this.base}${path}`, {
          method: 'POST', headers: await this._authHeaders({ 'Content-Type': 'application/json', Prefer: 'respond-async' }), body: JSON.stringify({ persistChanges: true }),
        });
      } catch (e) {
        this.logFn({ method: 'POST', path, status: 0, ms: Date.now() - started, error: e.message });
        if (attempt === 0) continue;
        throw new ExcelApiError({ message: `Could not reach Microsoft 365 to open the workbook: ${e.message}`, status: 0, method: 'POST', path });
      }
      const requestId = res.headers.get('request-id');
      const ms = Date.now() - started;
      if (res.status === 201) { const j = await readJson(res); this.sessionId = j.id; this.logFn({ event: 'session', status: 201, ms, requestId }); return this.sessionId; }
      if (res.status === 202) {
        this.logFn({ method: 'POST', path, status: 202, ms, requestId });
        this.sessionId = await this._pollSession(res.headers.get('Location'));
        this.logFn({ event: 'session', status: 201, ms: Date.now() - started, requestId });
        return this.sessionId;
      }
      const body = await readJson(res);
      const c = classifyError(res.status, body, res.headers);
      this.logFn({ method: 'POST', path, status: res.status, ms, requestId, ...errorFields(c) });
      if (res.status === 504 && attempt === 0) continue;
      throw new ExcelApiError({ message: c.message || 'Could not open a workbook session', status: res.status, code: c.code, innerCode: c.innerCode, innerCodes: c.innerCodes, retryAfter: c.retryAfter, requestId, method: 'POST', path });
    }
    throw new ExcelApiError({ message: 'Could not open a workbook session', method: 'POST', path });
  }
  async _pollSession(location) {
    if (!location) throw new ExcelApiError({ message: 'Session creation returned 202 without a Location header', status: 202, method: 'POST', path: '/workbook/createSession' });
    for (let i = 0; i < 120; i++) {
      await this.sleep(2000);
      const started = Date.now();
      let res;
      try { res = await this.fetch(location, { method: 'GET', headers: await this._authHeaders() }); }
      catch (e) {
        this.logFn({ method: 'GET', path: OPERATION_PATH, status: 0, ms: Date.now() - started, error: e.message });
        throw new ExcelApiError({ message: `Could not reach Microsoft 365 while the workbook was being opened: ${e.message}`, status: 0, method: 'GET', path: OPERATION_PATH });
      }
      const requestId = res.headers.get('request-id');
      const j = await readJson(res);
      if (!res.ok) {
        const c = classifyError(res.status, j, res.headers);
        this.logFn({ method: 'GET', path: OPERATION_PATH, status: res.status, ms: Date.now() - started, requestId, ...errorFields(c) });
        throw new ExcelApiError({ message: c.message || 'Session creation status check failed', status: res.status, code: c.code, innerCode: c.innerCode, innerCodes: c.innerCodes, retryAfter: c.retryAfter, requestId, method: 'GET', path: OPERATION_PATH });
      }
      const outcome = (j && j.status) || 'unknown';
      if (outcome === 'failed') {
        const c = classifyError(500, j, new Headers());
        this.logFn({ method: 'GET', path: OPERATION_PATH, status: res.status, ms: Date.now() - started, requestId, outcome, ...errorFields(c) });
        throw new ExcelApiError({ message: c.message || 'The workbook could not be opened', status: 500, code: c.code, innerCode: c.innerCode, innerCodes: c.innerCodes, requestId, method: 'GET', path: OPERATION_PATH });
      }
      this.logFn({ method: 'GET', path: OPERATION_PATH, status: res.status, ms: Date.now() - started, requestId, outcome });
      if (outcome === 'succeeded') {
        const started2 = Date.now();
        let r2;
        try { r2 = await this.fetch(j.resourceLocation, { method: 'GET', headers: await this._authHeaders() }); }
        catch (e) {
          this.logFn({ method: 'GET', path: SESSION_INFO_PATH, status: 0, ms: Date.now() - started2, error: e.message });
          throw new ExcelApiError({ message: `Session created but its id could not be read: ${e.message}`, status: 0, method: 'GET', path: SESSION_INFO_PATH });
        }
        const info = await readJson(r2);
        const requestId2 = r2.headers.get('request-id');
        if (!r2.ok || !info || !info.id) {
          const c = classifyError(r2.status, info, r2.headers);
          this.logFn({ method: 'GET', path: SESSION_INFO_PATH, status: r2.status, ms: Date.now() - started2, requestId: requestId2, ...errorFields(c), error: c.message || 'no session id in the answer' });
          throw new ExcelApiError({ message: 'Session created but its id could not be read', status: r2.status, code: c.code, innerCode: c.innerCode, innerCodes: c.innerCodes, requestId: requestId2, method: 'GET', path: SESSION_INFO_PATH });
        }
        this.logFn({ method: 'GET', path: SESSION_INFO_PATH, status: r2.status, ms: Date.now() - started2, requestId: requestId2 });
        return info.id;
      }
    }
    throw new ExcelApiError({ message: 'Timed out waiting for the workbook session', status: 202, method: 'GET', path: OPERATION_PATH });
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
    this._refuseIfBlocked(method, path);
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
    if (res.ok) { this.logFn({ method, path, status: res.status, ms: Date.now() - started, requestId }); return readJson(res); }
    const body = await readJson(res);
    const c = classifyError(res.status, body, res.headers);
    // A refusal is logged with its code, inner code and message: the live log of 2026-09-20 had only the status,
    // and the code had to be left unrecorded.
    this.logFn({ method, path, status: res.status, ms: Date.now() - started, requestId, code: c.code, innerCode: c.innerCode, innerCodes: c.innerCodes, error: c.message });
    let blocked = false;
    if (c.sessionInvalid && opts.session) {
      this.sessionId = null;                                   // further access to that session is not expected (Microsoft)
      if (c.recreatable) {
        if (opts.idempotent && attempt === 0) { await this._openSessionUnqueued(); return this._send(method, path, opts, 1); }
      } else {
        this._block(c, requestId, method, path);
        blocked = true;
      }
    } else if (opts.idempotent && attempt === 0 && c.retryableRead) {
      if (c.retryAfter) await this.sleep(Math.min(c.retryAfter, 30) * 1000);
      return this._send(method, path, opts, 1);
    }
    throw new ExcelApiError({
      message: c.message || `${method} ${path} failed (${res.status})`, status: res.status, code: c.code, innerCode: c.innerCode, innerCodes: c.innerCodes, retryAfter: c.retryAfter, requestId,
      ambiguous: !opts.idempotent && res.status >= 500, sessionInvalid: c.sessionInvalid, recreatable: c.recreatable, blocked, conflict: c.conflict, method, path,
    });
  }

  /**
   * JSON batch: sub-requests run in order on the server (each depends on the previous one).
   * Returns [{ status, body, headers, error }] in request order. Sent exactly once; a transport or
   * batch-level failure is `ambiguous` when any sub-request mutates.
   */
  batch(items) {
    return this._enqueue(async () => {
      this._refuseIfBlocked('POST', '/$batch');
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
      const requestId = res.headers.get('request-id');
      const body = await readJson(res);
      if (!res.ok) {
        const c = classifyError(res.status, body, res.headers);
        this.logFn({ method: 'POST', path: '/$batch', items: items.length, status: res.status, ms: Date.now() - started, requestId, ...errorFields(c) });
        throw new ExcelApiError({ message: c.message || `Batch request failed (${res.status})`, status: res.status, code: c.code, innerCode: c.innerCode, innerCodes: c.innerCodes, retryAfter: c.retryAfter, requestId, ambiguous: mutates && res.status >= 500, method: 'POST', path: '/$batch' });
      }
      this.logFn({ method: 'POST', path: '/$batch', items: items.length, status: res.status, ms: Date.now() - started, requestId });
      const byId = new Map(((body && body.responses) || []).map(r => [String(r.id), r]));
      return items.map((it, i) => {
        const r = byId.get(String(i + 1)) || { status: 0, body: null, headers: {} };
        const headers = r.headers || {};
        const method = (it.method || 'GET').toUpperCase();
        const error = r.status >= 400 || r.status === 0 ? classifyError(r.status, r.body, new Headers(headers)) : null;
        // Each sub-request gets its own line: a refusal inside a 200 batch would otherwise leave no trace.
        this.logFn({ batch: true, method, path: it.path, status: r.status, requestId: headers['request-id'] || null, ...(error ? errorFields(error) : {}), ...(error && !r.status ? { error: 'no answer for this request in the batch response' } : {}) });
        if (error && error.sessionInvalid) {
          this.sessionId = null;
          if (!error.recreatable && !this.blocked) this._block(error, headers['request-id'] || null, method, it.path);
        }
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
  /**
   * How a table row is addressed, and why the DELETE and the range PATCH spell it differently.
   *
   * The first live connection test (2026-09-20) sent DELETE …/rows/26 — the form the "TableRow: delete"
   * reference page shows — and Microsoft answered **400**: "The API you are trying to use could not be found.
   * It may be available in a newer version of Excel." The request log records status and message; it did not
   * record the error code, so none is claimed here. workbookTableRow has no id property, only index and
   * values, so rows/N asks for a row by a key the resource does not have.
   *
   * The DELETE reproduces Microsoft's own worked example in "Working with Excel in Microsoft Graph", the one
   * published DELETE-by-position form, `$` segment included:
   *     DELETE …/workbook/tables('4')/rows/$/itemAt(index=6)   →   204 No Content
   * The range PATCH keeps the "TableRow: Range" reference form, rows/itemAt(index=N)/range, which succeeded in
   * that same live run (PATCH …/rows/itemAt(index=26)/range → 200). Whether the service accepts the DELETE
   * form is still to be shown live — docs/checkpoint-2-runbook.md Part 2a, one row. Evidence and links:
   * docs/workbook-mapping.md §3.
   */
  deleteRowPath(table, index) { return `/workbook/tables/${enc(table)}/rows/$/itemAt(index=${Number(index)})`; }
  deleteTableRow(table, index) { return this.request('DELETE', this.deleteRowPath(table, index)); }
  rowRangePath(table, index) { return `/workbook/tables/${enc(table)}/rows/itemAt(index=${Number(index)})/range`; }
  getTableRowRange(table, index) { return this.request('GET', this.rowRangePath(table, index)); }
  patchTableRowRange(table, index, props) { return this.request('PATCH', this.rowRangePath(table, index), { body: props }); }
  calculate(type = 'Full') { return this.request('POST', '/workbook/application/calculate', { body: { calculationType: type } }); }
  getWorksheetProtection(sheet) { return this.request('GET', `/workbook/worksheets/${enc(sheet)}/protection`); }
  getItemMeta() { return this.request('GET', '?$select=id,name,size,lastModifiedDateTime,webUrl,eTag,file,parentReference', { session: false }); }
}
