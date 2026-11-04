// Find and confirm which workbook the app should use, and remember the choice per tenant + account.
// Workbooks are addressed by drive id + item id (opaque ids, not sharing links). Access is still
// decided by Microsoft 365: a member who cannot open the file in OneDrive cannot open it here.
const SELECT = 'id,name,size,lastModifiedDateTime,webUrl,parentReference,file,folder';
const seg = s => encodeURIComponent(String(s));

export function workbookKey({ driveId, itemId }) { return `${driveId}!${itemId}`; }

function extOf(name) { const m = /\.([A-Za-z0-9]+)$/.exec(name || ''); return m ? m[1].toLowerCase() : ''; }
function pathOf(parentReference) {
  const p = (parentReference && parentReference.path) || '';
  const i = p.indexOf('root:');
  return i >= 0 ? (p.slice(i + 5) || '/') : '';
}
function toCandidate(item, source, sharedBy) {
  return {
    driveId: (item.parentReference && item.parentReference.driveId) || '',
    itemId: item.id,
    name: item.name || '',
    extension: extOf(item.name),
    path: pathOf(item.parentReference),
    size: item.size ?? null,
    lastModified: item.lastModifiedDateTime || '',
    webUrl: item.webUrl || '',
    source,
    sharedBy: sharedBy || '',
    warnings: [],
  };
}

/**
 * Lists .xlsx workbooks the signed-in member can see: their own OneDrive (by name) and files shared with them.
 * `graphGet(path)` performs an authenticated GET against Microsoft Graph v1.0 and returns parsed JSON.
 * The returned array carries a `warnings` property for partial failures.
 */
export async function findCandidates(graphGet, { query = 'TPAHA_2026' } = {}) {
  const found = [];
  const warnings = [];
  try {
    const r = await graphGet(`/me/drive/root/search(q='${seg(query.replace(/'/g, "''"))}')?$select=${SELECT}`);
    for (const it of (r && r.value) || []) if (it.file && extOf(it.name) === 'xlsx') found.push(toCandidate(it, 'mine'));
  } catch (e) { warnings.push(`Searching your OneDrive failed: ${e.message}`); }
  try {
    const r = await graphGet('/me/drive/sharedWithMe');
    for (const it of (r && r.value) || []) {
      const ri = it.remoteItem;
      if (!ri || !ri.file || extOf(ri.name) !== 'xlsx') continue;
      const by = ri.shared && ri.shared.sharedBy && ri.shared.sharedBy.user ? ri.shared.sharedBy.user.displayName : '';
      found.push(toCandidate(ri, 'shared', by));
    }
  } catch (e) { warnings.push(`Files shared with you could not be listed: ${e.message}`); }
  const byKey = new Map();
  for (const c of found) {
    const k = workbookKey(c);
    if (!byKey.has(k) || (byKey.get(k).source === 'shared' && c.source === 'mine')) byKey.set(k, c);
  }
  const list = [...byKey.values()].sort((a, b) => String(b.lastModified).localeCompare(String(a.lastModified)));
  list.warnings = warnings;
  return list;
}

const choiceKey = ({ tenantId, homeAccountId, app }) => `tpaha:workbook:${seg(app)}:${seg(tenantId || '-')}:${seg(homeAccountId || '-')}`;

export function rememberWorkbook(storage, who, choice) {
  const record = { driveId: choice.driveId, itemId: choice.itemId, name: choice.name || '', webUrl: choice.webUrl || '', chosenAt: new Date().toISOString() };
  storage.setItem(choiceKey(who), JSON.stringify(record));
  return record;
}
export function recallWorkbook(storage, who) {
  try {
    const raw = storage.getItem(choiceKey(who));
    if (!raw) return null;
    const obj = JSON.parse(raw);
    return obj && obj.driveId && obj.itemId ? obj : null;
  } catch { return null; }
}
export function forgetWorkbook(storage, who) { try { storage.removeItem(choiceKey(who)); } catch { /* ignore */ } }

/** The deployment may pin the workbook in config.js; both ids are required. */
export function resolveConfigured(config, app) {
  const w = config && config.workbooks && config.workbooks[app];
  return w && w.driveId && w.itemId ? { driveId: w.driveId, itemId: w.itemId, name: w.name || '', source: 'config' } : null;
}

/**
 * The workbook the ledger page opens at start-up: the pinned one, always. A choice this device remembered for
 * the signed-in account (a test copy picked through the old picker) is ignored and removed, so it cannot come
 * back if the pin is ever lifted. Without a pin nothing is opened: the page does not fall back to a remembered
 * copy and never searches. Returns { workbook, replaced } where `replaced` is the forgotten record, if any.
 */
export function startupWorkbook({ config, storage, who, app = 'ledger' }) {
  const pinned = resolveConfigured(config, app);
  if (!pinned) return { workbook: null, replaced: null };
  const remembered = recallWorkbook(storage, who);
  if (remembered) forgetWorkbook(storage, who);
  return { workbook: pinned, replaced: remembered };
}

/**
 * Words for a failure to open the configured workbook. `title`, `reason`, `scope` and `detail` are sentences
 * for the page; the drive and item ids are deliberately not among them.
 */
export function describeAccessFailure(err, name) {
  const e = err || {};
  const status = Number(e.status) || 0;
  const code = String(e.innerCode || e.code || '');
  const message = e.message ? String(e.message) : '';
  let reason;
  if (status === 403 || /accessDenied|forbidden|unauthenticated/i.test(code)) reason = `Your Microsoft account does not have access to this workbook, or the access has been withdrawn. Ask the treasurer to share ${name || 'the workbook'} with you, then try again.`;
  else if (status === 401) reason = 'Your sign-in is no longer valid. Sign out, sign in again, then try again.';
  else if (status === 404 || /itemNotFound|notFound/i.test(code)) reason = `The configured workbook was not found where the site expects it. It may have been moved, renamed or deleted. Nothing has been changed; tell the treasurer before doing anything else.`;
  else if (status === 0) reason = `Microsoft 365 could not be reached (${message || 'network problem'}). Check the connection and try again.`;
  else reason = `Microsoft 365 refused the request${message ? `: ${message}` : ''}. Try again in a moment; if it keeps failing, tell the treasurer.`;
  const parts = [status ? `HTTP ${status}` : 'no HTTP answer'];
  if (code) parts.push(`code ${code}`);
  if (e.requestId) parts.push(`request-id ${e.requestId}`);
  return {
    title: `Could not open ${name || 'the configured workbook'}`,
    reason,
    scope: 'This page opens only the workbook configured for the club and does not search for other copies. Nothing was written.',
    detail: parts.join(' · '),
  };
}
