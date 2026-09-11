# TPAHA Books

A small web site that lets TPAHA board members work with the association's **Excel workbooks in
Microsoft 365** from an iPhone, an Android phone or an older computer's browser. The workbook stays
the record: the site signs the member in with their Microsoft 365 account, reads the workbook through
Microsoft Graph, shows phone-friendly forms, and writes the member's intended changes back into the
same workbook when they press **Save**, confirming only after Microsoft accepted the write and the
site read the result back.

- **Treasurer Ledger** (`site/ledger.html`): add, correct and delete transactions in `LOG_Table`,
  read the workbook's monthly sheets and annual summary, change the prior-year balance.
  Status: built and tested against an in-memory replica of the workbook; **not yet verified against
  Microsoft 365** (see `IMPLEMENTATION_STATUS.md`).
- **Winter Storage** (`site/storage.html`): the earlier screen, working on a copy kept in the browser.
  **Not connected to its workbook yet**; that is Phase 4.
- **Diagnostics** (`site/diagnostics.html`): identify the workbook, run read-only checks, and on a
  test copy run the add / read-back / delete / compare self-test (the page states exactly which areas it compares).

There is no database and no generated Excel file. Nothing in the repository contains records: the
workbooks and everything derived from them live in `data/`, which git ignores.

## How saving works (short version)

1. The site reads the transaction table fresh, right before writing.
2. Adding = one table-row append carrying an operation id (the entry time as an ISO timestamp with nine
   fraction digits, stored in the Timestamp column), then a number-format touch on the new row, a read-back
   that finds the row by that id, a uniqueness check on the transaction number, and a rebuild of the
   sorted helper table the monthly sheets depend on.
3. Correcting or deleting = locate the row by transaction number + timestamp, compare it with the copy
   the member loaded (a fingerprint of all nine cells); any difference is shown as "Someone else changed
   this record" and nothing is written. A correction appends the corrected copy (same number and
   timestamp), verifies it, then removes the old copy. Right before every delete (or renumbering) the
   table is read once more and the request is sent only if the intended row is still exactly at that
   index; otherwise the site re-aims or stops.
4. A lost answer from Microsoft is reported as "Could not save". "Check workbook" and the automatic check
   after a reload only READ and report what they found (not there / there but not finished / done /
   conflict). A Save then finishes or re-sends that same operation; nothing is written twice, and a newer
   draft typed in the meantime stays unsaved and visible until it is saved on its own.
5. One page runs one complete operation at a time (a second Save or Delete is refused while one runs, and
   the adapter queues anything that slips through), the ledger and Diagnostics pages share one same-browser
   tab guard, and every operation id comes from the same nine-digit generator.
6. Concurrency boundary: Microsoft Graph has no conditional write, lock or transaction. A shift that
   happens before the check read is caught; one that happens in the instant between the read and the
   request is detected afterwards by identity, named, and turned into an INCIDENT that pauses saving
   across Refresh and reload until the documented conditions are verified (or, when nothing can be
   verified, a member acknowledges after checking Excel's version history). Repairs are explicit and
   confirmed. Because that second case cannot be prevented from a browser, the pilot runs with **one
   designated writer at a time** and does not support simultaneous editing. Full statement in
   `docs/workbook-mapping.md` §3b.

Full mapping: `docs/workbook-mapping.md`. Design and plan: `docs/superpowers/`.

## Documents

| For | File |
|---|---|
| Administrator: hosting, Microsoft 365 app registration, restricting sign-in, first live test | `docs/setup-guide.md` |
| The day of the first live test: prerequisites, the run, what to send back, when to stop | `docs/checkpoint-2-runbook.md` |
| Board members: open, sign in, edit, Save, what the messages mean | `docs/member-guide.md` |
| Progress, what is verified, what is not, open questions | `IMPLEMENTATION_STATUS.md` |
| Website action → workbook operation | `docs/workbook-mapping.md` |

## Development

```
npm install                  # once
npm test                     # unit tests (node --test): Graph client, mock workbook, ledger adapter, save controller, drafts, models
npm run test:e2e             # Playwright: ledger + diagnostics + storage screens in test mode; desktop Chromium, iPhone-size Chromium, iPhone-size WebKit
npm run serve                # http://localhost:8787/  (test mode: in-memory sample workbook, nothing touches Microsoft 365)
python tools/build_seed.py   # only with the real .xlsx files present: writes data/ (git-ignored) fixtures used by the formula oracle tests
```

Test mode is what you get when a page is served from this machine. `site/js/config.js` supplies the app
registration only to the published host and blanks it for `localhost` and `127.0.0.1`, so local pages run
against a synthetic sample workbook (`site/js/workbook/sample-workbook.js`) served by `js/workbook/mock-excel.js`,
which answers the same Graph URLs the real client uses and mirrors the workbook's formulas. The rule is the
address the page came from, not a flag or a query parameter, so nothing a member can click, type or paste puts
the published site into test mode. `npm run verify:config` checks that rule and the rest of the sign-in
configuration without contacting Microsoft. The
formula oracle tests compare against values cached by Excel and skip automatically when the git-ignored
fixtures are absent. **Those four test files are themselves git-ignored** (`tests/unit/ledger-model.test.js`,
`tests/unit/storage-model.test.js`, `tests/unit/xlsx-export.test.js`, `tests/e2e/storage.spec.js`): they assert
figures taken from the association's real workbooks, so they stay on the treasurer's machine beside the
fixtures rather than in this repository. A clone therefore runs a smaller suite than the status report quotes,
and the difference is exactly those files.

Layout: `site/js/workbook/` (Graph Excel client, ledger adapter, mock, locator), `site/js/save/`
(save states, drafts), `site/js/ledger/` (screen + verified formula model used for validation and
cross-checks), `site/js/storage/` (storage screen, local copy), `tests/` (unit + Playwright).

To release, change `appVersion` in `site/js/config.js` and push; GitHub Pages deploys `site/`
through `.github/workflows/pages.yml`.
