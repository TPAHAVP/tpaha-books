// Deployment configuration for TPAHA Books.
// Everything here is configuration, not a secret: client and tenant ids identify the app and the
// organisation; access to the workbook is still decided by Microsoft 365 permissions at sign-in.
// An empty clientId puts the site in TEST MODE: an in-memory copy of the workbook, nothing saved to Microsoft 365.
export const CONFIG = {
  clientId: '',            // Entra app registration: "Application (client) ID"
  tenantId: '',            // Entra app registration: "Directory (tenant) ID" (restricts sign-in to TPAHA)
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
  appVersion: '2026.09.06',
  siteName: 'TPAHA Books',
};
