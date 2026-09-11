// Deployment configuration for TPAHA Books.
//
// Everything here is configuration, not a secret. The client and tenant ids identify the app registration and
// the organisation; who may sign in, and what they may open, is decided by Microsoft 365 at sign-in. There is
// no client secret in this site and there must never be one: this code runs in the member's browser, where
// nothing can be kept secret.
//
// TEST MODE. A page served from this machine runs against an in-memory sample workbook with made-up figures and
// never reaches Microsoft 365. That is how the automated tests and local development run. The published site
// always uses the ids below. The rule is the address the page was served from, not a flag or a query
// parameter, so nothing a member can click, type or paste can put the real site into test mode.
const LOCAL_HOSTS = new Set(['localhost', '127.0.0.1', '[::1]', '::1', '']);
const servedLocally = typeof location === 'undefined' || LOCAL_HOSTS.has(location.hostname);

// Entra app registration "TPAHA Books", completed 2026-09-11: single tenant; platform Single-page application
// with redirect URI https://tpahavp.github.io/tpaha-books/ (the site root, which is what js/auth.js derives);
// delegated Microsoft Graph User.Read and Files.ReadWrite with admin consent granted; "Assignment required"
// set to Yes with only the intended accounts assigned; implicit grant unchecked; public client flows disabled;
// no client secret and no certificate.
const ENTRA = {
  clientId: '6c5bf2f8-a50c-4908-ba93-535816364785',   // Application (client) ID
  tenantId: 'ef5ad4e3-2b18-4542-8fd2-ed7212cbcc93',   // Directory (tenant) ID: restricts sign-in to TPAHA
};

export const CONFIG = {
  clientId: servedLocally ? '' : ENTRA.clientId,
  tenantId: servedLocally ? '' : ENTRA.tenantId,
  workbooks: {
    // Leave blank to let the signed-in member pick the workbook (search by name, confirm, remembered per account).
    ledger: { driveId: '', itemId: '', name: '' },
    storage: { driveId: '', itemId: '', name: '' },
  },
  workbookSearch: 'TPAHA_2026',   // name used when searching the member's OneDrive for the ledger workbook
  allowWorkbookPicker: true,
  // Pilot writer model (docs/workbook-mapping.md §3b, docs/setup-guide.md E). 'single' = one designated writer at
  // a time: the page asks each session to confirm before its first Correct or Delete, refuses to write from a
  // second tab of the same browser, and states that it does not support simultaneous editing.
  pilot: {
    writerModel: 'single',
    note: 'Pilot: one designated writer at a time. Do not use this page in two tabs or on two devices, and do not edit the workbook in Excel or run its script buttons while it is in use. Simultaneous editing is not supported.',
  },
  appVersion: '2026.09.11',
  siteName: 'TPAHA Books',
};
