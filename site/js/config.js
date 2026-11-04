// Deployment configuration for TPAHA Books.
//
// Everything here is configuration, not a secret. The client and tenant ids identify the app registration and
// the organisation; who may sign in, and what they may open, is decided by Microsoft 365 at sign-in. There is
// no client secret in this site and there must never be one: this code runs in the member's browser, where
// nothing can be kept secret.
//
// TEST MODE, and exactly what decides it. The rule is a denylist of local hostnames, not an allowlist of one
// published address: a page served from localhost, 127.0.0.1, ::1 or from a file:// URL gets blank ids and runs
// against an in-memory sample workbook with made-up figures, never reaching Microsoft 365. That is how the
// automated tests and local development stay offline. **Every other hostname gets the ids below**, including a
// copy served from another machine's IP address or a fork published at a different address.
//
// That is deliberate and is not what keeps the app safe. Two other things do. The Entra registration accepts
// exactly one SPA redirect URI, so a sign-in started anywhere but https://tpahavp.github.io/tpaha-books/ is
// rejected by Microsoft before any token is issued. And "Assignment required" limits sign-in to the assigned
// accounts, who could already open the workbook themselves. A client id is an identifier, not a credential.
//
// The practical consequence for a member: nothing they can click, type or paste changes the mode, because it
// is not a flag or a query parameter. The practical consequence for a developer: serve the site from
// localhost, never from a LAN address, or the page will try to sign in for real.
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

// The club's treasurer workbook, pinned 2026-11-03: the published ledger page opens this file and no other, for
// every member on every device, as soon as they have signed in. The ids are Microsoft's opaque drive and item
// identifiers (shown on the Diagnostics page), not a sharing link and not a credential: a member who cannot open
// the file in OneDrive cannot open it here either. A workbook a device remembered earlier (a test copy chosen
// through the old picker) is ignored and forgotten. When this file cannot be opened the page says so and stops;
// it never searches for other copies. Locally served pages get blank ids and the in-memory sample (test mode).
const LEDGER_WORKBOOK = {
  name: 'TPAHA_2026.xlsx',
  driveId: 'b!O4R42hSWcEGFTomb8sxHJ9A0OrJC045NrcKyGO_FkY9O3XaCxPiCSKILopSmUmlS',
  itemId: '01WV6K6T2HOQUUIGZNTRFZ7PF7IDDUKMEM',
};
const NO_WORKBOOK = { driveId: '', itemId: '', name: '' };

export const CONFIG = {
  clientId: servedLocally ? '' : ENTRA.clientId,
  tenantId: servedLocally ? '' : ENTRA.tenantId,
  workbooks: {
    ledger: servedLocally ? NO_WORKBOOK : LEDGER_WORKBOOK,
    storage: NO_WORKBOOK,     // Winter Storage is out of scope; its page is unchanged
  },
  workbookSearch: 'TPAHA_2026',   // used only by the Diagnostics page's search; the ledger page never searches
  allowWorkbookPicker: false,     // the ledger page has no workbook picker and no "choose a different workbook
  // Pilot writer model (docs/workbook-mapping.md §3b, docs/setup-guide.md E). 'single' = one designated writer at
  // a time: the page asks each session to confirm before its first Correct or Delete, refuses to write from a
  // second tab of the same browser, and states that it does not support simultaneous editing.
  pilot: {
    writerModel: 'single',
    note: 'Pilot: one designated writer at a time. Do not use this page in two tabs or on two devices, and do not edit the workbook in Excel or run its script buttons while it is in use. Simultaneous editing is not supported.',
  },
  appVersion: '2026.11.03',       // change on every release: the service worker caches by this value
  siteName: 'TPAHA Books',
};
