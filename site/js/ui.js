// Shared formatting and small DOM helpers. Pure functions first (unit-tested), DOM helpers after.

export function fmtMoney(cents) {
  const n = Math.abs(Math.round(cents)) / 100;
  const s = '$' + n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return cents < 0 ? '-' + s : s;
}
export const fmtAmount = n => fmtMoney(Math.round(Number(n) * 100));

/** Numbers as Excel would show them; Excel's text values pass through unchanged. */
export function fmtNumberOrText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return String(Number(v.toFixed(6)));
  return String(v);
}

const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
export function fmtDate(iso, style = 'short') {
  if (!iso) return '';
  const [y, m, d] = iso.split('-').map(Number);
  return style === 'long' ? `${d} ${MON[m - 1]} ${y}` : `${d} ${MON[m - 1]}`;
}
export function todayIso() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
export function monthRange(year, month) {
  const last = new Date(Date.UTC(year, month, 0)).getUTCDate();
  const mm = String(month).padStart(2, '0');
  return [`${year}-${mm}-01`, `${year}-${mm}-${String(last).padStart(2, '0')}`];
}
export function splitPhone(s) {
  return String(s || '').split('/').map(x => x.trim()).filter(Boolean);
}
export function fmtTime(d) {
  return d ? new Date(d).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '';
}

// ---- DOM helpers ------------------------------------------------------------------------

/** h('div', { className, dataset, onClick, ...attrs }, ...children) */
export function h(tag, attrs = {}, ...children) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v === null || v === undefined || v === false) continue;
    if (k === 'className') el.className = v;
    else if (k === 'dataset') Object.assign(el.dataset, v);
    else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
    else if (k === 'html') el.innerHTML = v;
    else if (k === 'value') el.value = v;
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const c of children.flat(Infinity)) {
    if (c === null || c === undefined || c === false) continue;
    el.append(c instanceof Node ? c : document.createTextNode(String(c)));
  }
  return el;
}

export function toast(message, kind = 'ok') {
  let host = document.getElementById('toasts');
  if (!host) { host = h('div', { id: 'toasts', className: 'toasts', 'aria-live': 'polite' }); document.body.append(host); }
  const t = h('div', { className: `toast toast-${kind}`, role: 'status' }, message);
  host.append(t);
  requestAnimationFrame(() => t.classList.add('show'));
  setTimeout(() => { t.classList.remove('show'); setTimeout(() => t.remove(), 300); }, kind === 'error' ? 7000 : 3500);
}

export function confirmDialog(message, { okLabel = 'OK', danger = false, title = '' } = {}) {
  return new Promise(resolve => {
    let done = false;
    const finish = v => { if (done) return; done = true; dlg.close(); resolve(v); };
    const dlg = h('dialog', { className: 'confirm' },
      title ? h('h3', {}, title) : null,
      h('p', {}, message),
      h('div', { className: 'row-end' },
        h('button', { type: 'button', className: 'btn', onClick: () => finish(false) }, 'Cancel'),
        h('button', { type: 'button', className: danger ? 'btn btn-danger' : 'btn btn-primary', onClick: () => finish(true) }, okLabel)));
    dlg.addEventListener('close', () => { dlg.remove(); if (!done) { done = true; resolve(false); } });
    document.body.append(dlg);
    dlg.showModal();
  });
}

/** Tab buttons carry data-tab="x"; panels are #tab-x. Returns show(id). */
export function setupTabs(nav, onChange) {
  const buttons = [...nav.querySelectorAll('[data-tab]')];
  const show = id => {
    for (const b of buttons) {
      const on = b.dataset.tab === id;
      b.classList.toggle('active', on);
      b.setAttribute('aria-selected', on ? 'true' : 'false');
      const panel = document.getElementById(`tab-${b.dataset.tab}`);
      if (panel) panel.hidden = !on;
    }
    if (location.hash !== `#${id}`) history.replaceState(null, '', `#${id}`);
    if (onChange) onChange(id);
  };
  for (const b of buttons) b.addEventListener('click', () => show(b.dataset.tab));
  const wanted = location.hash.slice(1);
  show(buttons.some(b => b.dataset.tab === wanted) ? wanted : buttons[0].dataset.tab);
  return show;
}

export function downloadBlob(filename, blob) {
  const a = h('a', { href: URL.createObjectURL(blob), download: filename });
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}

export function readFileText(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(String(r.result));
    r.onerror = () => rej(r.error);
    r.readAsText(file);
  });
}

/** Human text for the Store status. */
export function syncStatusText(status) {
  if (status.mode === 'local') return 'Saved in this browser';
  if (status.error) return `Sync problem: ${status.error}`;
  if (status.syncing) return 'Saving to OneDrive…';
  if (status.offline) return status.dirty ? 'Offline. Changes will sync when back online.' : 'Offline. Showing the last copy.';
  if (status.dirty) return 'Unsaved changes…';
  return status.lastSynced ? `Synced to OneDrive ${fmtTime(status.lastSynced)}` : 'OneDrive';
}

/** Prints with a temporary @page rule (Safari has no named pages). */
export function printWithPage(pageRule) {
  const style = h('style', { html: `@page { ${pageRule} }` });
  document.head.append(style);
  const cleanup = () => { style.remove(); window.removeEventListener('afterprint', cleanup); };
  window.addEventListener('afterprint', cleanup);
  window.print();
  setTimeout(cleanup, 60_000);
}

/** Card shown in OneDrive mode when nobody is signed in. */
export function signInGate(auth, appName) {
  return h('section', { className: 'card gate' },
    h('h2', {}, `Sign in to use the ${appName}`),
    h('p', {}, 'Use your TPAHA Microsoft account. The data is read from and saved to the shared OneDrive folder.'),
    h('button', { type: 'button', className: 'btn btn-primary btn-lg', id: 'btn-signin', onClick: () => auth.signIn() }, 'Sign in with Microsoft'));
}

/** Card shown in OneDrive mode when no data folder link is configured. */
export function shareUrlGate(onSave) {
  const input = h('input', { type: 'url', id: 'share-url', placeholder: 'https://…sharepoint.com/…', required: true });
  return h('section', { className: 'card gate' },
    h('h2', {}, 'Where is the data folder?'),
    h('p', {}, 'Paste the OneDrive sharing link of the "TPAHA Books" folder. In OneDrive, select the folder, choose Share, then Copy link.'),
    h('form', { onSubmit: e => { e.preventDefault(); onSave(input.value.trim()); } },
      h('label', { className: 'field' }, 'Folder sharing link', input),
      h('button', { type: 'submit', className: 'btn btn-primary' }, 'Save link')));
}
