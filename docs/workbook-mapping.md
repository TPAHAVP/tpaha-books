# Workbook mapping: TPAHA_2026 ledger

Source inspected read-only: `TPAHA_2026 (1).xlsx` in the project folder, dated 2026-09-06 (saved by Excel Online).

The online copy in the pilot OneDrive was **read** through this app for the first time on 2026-09-11: it was
identified and confirmed on the Diagnostics page, and the six read-only checks passed. That establishes its
identity and structure. It does **not** establish that its contents match this document's §5 findings, which
were made against the local file: the read-only checks do not look at the Annual sheet's formulas or at
duplicate transaction numbers, and the two copies are not the same snapshot. Nothing has been written to it.

## 1. Structure

| Worksheet | Purpose | Protected | Formulas | Edited by the app |
|---|---|---|---|---|
| ENTRY | Form cells for the old Office Script buttons; `B23` prior-year balance | yes (form cells unlocked) | none | `B23` only |
| January … December | Monthly statements built by formula from `LOG_Sorted_Table` and `LOG_Table` | yes | 404 each | never (read only) |
| ConfigHidden (hidden) | `A1` = workbook year (2026), named `WORKBOOK_YEAR` | no | none | never |
| LISTS (hidden) | `LISTS_Categories` table: CategoryID, CategoryName, Type, ColumnLetter; `DepositCategories` = E2:E5, `WithdrawalCategories` = F2:F5 | no | none | never |
| LOG (hidden) | `LOG_Table` (A1:I27): the transaction record | no | none | **yes**: add, correct, delete rows |
| LOG_Sorted (hidden) | `LOG_Sorted_Table` (A1:I26): copy of LOG sorted by Date; the month sheets read it | no | none | **yes**: rebuilt after every LOG change |
| Annual 2026 | Yearly summary by formula from `LOG_Table` | yes | 146 | never (read only) |

Named ranges: `WORKBOOK_YEAR`, `DepositCategories`, `WithdrawalCategories`, plus button anchors
(`SubmitButton`, `SearchButton`, `DeleteButton_Row17..20`) used only by the old scripts.

Buttons: three Office Script buttons on ENTRY (Run SubmitEntry, Run FindTransactions,
Run DeleteTransaction). Their code is stored in the owner's OneDrive, not in the file. Graph does
not run them; the web app reproduces their effect with the operations below. Their source was read on
2026-09-10; **§7 records what they do and what this app must satisfy to keep them working.**

## 2. LOG_Table columns

| # | Column | Type in cells | Number format | Notes |
|---|---|---|---|---|
| A | TransactionID | number | General | Positive integers, unique after the 2026-09-06 fix (three duplicates were renumbered; the figures are in the local notes). New ID = current maximum + 1. **Consequence:** if the highest-numbered row is deleted, its number can be reused by a later entry. Existing rows are never renumbered by a correction; the only renumbering is of a *just-added* row whose number turned out to be taken at the same instant (§3b). |
| B | Timestamp | text, ISO 8601 UTC. Rows written by the old scripts: three fraction digits, e.g. `2026-04-25T01:55:48.134Z`. **Rows written by the app: nine fraction digits** (the millisecond plus six random digits), e.g. `2026-09-06T14:03:11.482913057Z` | `yyyy-mm-dd hh:mm:ss` | Entry time, and for app rows the **operation id** that makes checks and retries safe. The extra digits exist because two members can press Save in the same millisecond (review finding R4). The value is still a valid ISO 8601 date-time (`new Date()` in JavaScript and Office Scripts parses it; Excel keeps it as text as before), but it is a **content-format change** to the column. **Confirmed on 2026-09-10 (§7): none of the three scripts reads, parses or compares this column**, so the nine-digit form does not affect them. They do rewrite it, with the value they read, as part of a whole-table rewrite; see §7. |
| C | Date | number (Excel serial) | `yyyy-mm-dd` | Transaction date. Must be written as a serial number, not text, or COUNTIFS/SUMIFS stop matching. |
| D | Type | text | General | `Deposit` or `Withdrawal` |
| E | Category | text | General | One of the eight names in LISTS |
| F | Amount | number | `$#,##0.00` | Positive |
| G | Description | text | General | |
| H | ChequeNum | number or blank | General | Numbers in the existing rows; text accepted |
| I | Notes | text or blank | General | |

Quirk: the table body's first row (sheet row 2) is empty. The app skips blank rows when reading
and never writes into them; new rows are appended at the end.

Row identity used by the app = TransactionID + Timestamp. Row fingerprint = all nine values.

## 3. Website action → workbook operation

All calls go through one persistent Graph workbook session (`workbook-session-id`), one request
at a time from one page, and every write is followed by a read-back before the app reports "Saved".
Reads may be retried once (network loss, 429/503/504 with `Retry-After`, expired session). **Writes are
sent exactly once**; when the answer to a write is lost the operation is reported as uncertain and is
resolved by reading (§3b), never by sending it again automatically. **Every positional request (a
DELETE of `rows/$/itemAt(index=n)`, a PATCH of `rows/itemAt(index=n)/range`) is preceded by a fresh read that is evaluated
before the request is sent**: the row at that index must be exactly the intended row (identity and
content), otherwise the app re-reads and re-aims, or stops. No `$batch` is used by the adapter
(the client still implements it; the earlier read+delete batch was withdrawn after the re-review because
the GET inside a batch cannot be evaluated before the DELETE runs).

**One complete operation at a time per page** (review finding T1): the adapter queues every operation
(add, correction, delete, finish, balance write, repair, and the reads that decide them) so that the next one
starts only after the previous one has finished its verification and helper-table work. Two Delete buttons
pressed in quick succession therefore cannot both plan a positional DELETE from the same stale layout: the
second waits and reads fresh positions. The page also refuses a second write action while one runs ("Another
change is still being saved") and greys the Delete buttons. This serialises one page only.

Two kinds of method exist for every operation that can be left half-done: **inspect** (read-only:
GETs only, returns `missing` / `landed` / `incomplete` / `conflict`) and **complete** (finishes an
operation that landed: number formats, uniqueness of the number, sorted table, removal of an old copy).
The page's "Check workbook" button and the automatic check after a reload call only *inspect*;
only a Save calls *complete* (or writes the operation if nothing landed).

**How a table row is addressed — corrected after the first live write test (2026-09-20).** The reference page
for [TableRow: delete](https://learn.microsoft.com/en-us/graph/api/tablerow-delete?view=graph-rest-1.0) shows
`DELETE …/tables/{id|name}/rows/{index}`, and that is what this app sent: `DELETE /workbook/tables/LOG_Table/rows/26`.
The live service answered **400** with the message *"The API you are trying to use could not be found. It may
be available in a newer version of Excel. Please refer to the documentation: …"*. The request log records the
status and the message; **it did not record an error code, and none is claimed here.** The connection test
stopped at its delete step and its test row stayed in the workbook as #30. (An earlier draft of this section
wrote "404 ApiNotFound", taken from other people's reports rather than from this log; corrected 2026-09-20.)

| Evidence | What it shows |
|---|---|
| The live request log, 2026-09-20 | `DELETE /workbook/tables/LOG_Table/rows/26 → 400`, message above. In the same run, `PATCH /workbook/tables/LOG_Table/rows/itemAt(index=26)/range → 200` |
| [workbookTableRow resource](https://learn.microsoft.com/en-us/graph/api/resources/workbooktablerow?view=graph-rest-1.0) | The row has **two properties, `index` and `values`, and no `id`**. `rows/{index}` asks for a row by a key the resource does not have; the SDK snippets on the delete page itself key on a *row id* (`Rows["{workbookTableRow-id}"]`) |
| [Working with Excel in Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0), "Delete table row" | **Microsoft's own worked example**, the one published DELETE-by-position form: `DELETE …/workbook/tables('4')/rows/$/itemAt(index=6)` → `204 No Content`. The rows listed on that same page carry `@odata.id`s of the form `…/rows/itemAt(0)` |
| [TableRowCollection: ItemAt](https://learn.microsoft.com/en-us/graph/api/tablerowcollection-itemat?view=graph-rest-1.0) and [TableRow: Range](https://learn.microsoft.com/en-us/graph/api/tablerow-range?view=graph-rest-1.0) | The reference form for reading a row by position and for its range: `rows/itemAt(index={index})` and `rows/itemAt(index={index})/range` |
| Microsoft Q&A: [DELETE rows/{index}](https://learn.microsoft.com/en-us/answers/questions/2149984/i-am-using-this-graph-api-delete-https-graph-micro) and [PATCH rows/{index}](https://learn.microsoft.com/en-us/answers/questions/1680085/error-apinotfound-trying-standard-update-workbook) | Others met the same message for both methods (their reports show the code `ApiNotFound`; ours did not record a code). The second thread's working form is `rows/$/ItemAt(index=…)` |

**The fix follows the documents, one form per operation.** The DELETE reproduces Microsoft's worked example
exactly, `$` segment included: `DELETE …/rows/$/itemAt(index=n)`, expecting 204. The range PATCH keeps the
reference form that succeeded live: `PATCH …/rows/itemAt(index=n)/range`. Nothing else about the delete
changed — the fresh read evaluated before sending, the identity-and-content check at that index, the
exactly-once send, and the read-back that decides an uncertain answer are as they were; a 4xx is a clean
refusal, thrown and never retried.

**The mock** now refuses `rows/N` for every method with 400 and the logged message (its error code left blank
rather than guessed), deletes only through the published `rows/$/itemAt(index=n)` form, answering 204, and
refuses `DELETE rows/itemAt(index=n)` with a refusal labelled as the mock's own, so the app cannot drift onto a
spelling no document shows. It had accepted `rows/N`, which is how the app reached the live test with it.
**Whether the service accepts the DELETE form is still to be shown live** — the targeted cleanup of #30 in
`docs/checkpoint-2-runbook.md` Part 2a is that proof, one row. The request log now records a refusal's error
code and inner code as well as its status, so the next one is recorded rather than reconstructed.

| Website action | Graph operations (in order) | Verification |
|---|---|---|
| Open / Refresh | `GET /drives/{d}/items/{i}` (name, last modified); `GET …/worksheets('ConfigHidden')/range(address='A1')`; `GET …/worksheets('ENTRY')/range(address='B23')`; `GET …/tables('LISTS_Categories')/dataBodyRange`; `GET …/tables('LOG_Table')/dataBodyRange`; header row read | Column headers checked; blank rows skipped; each row gets a fingerprint. **Integrity scan:** if the same identity appears twice (a copy pasted in Excel, or a correction interrupted on another device) the page records an incident and pauses writes (§3b); a copy that belongs to this page's own unfinished correction is not reported as damage because that correction will finish or report it. A Refresh never lifts a pause. |
| Search transactions | none (filters the loaded rows) | Refresh re-reads the table. |
| Add transaction (Submit Entry) | …then §8 report row visibility for the month of the new date. 1 `GET LOG_Table/dataBodyRange`: next number, and "is there already a row with my operation id?" (same content → the earlier attempt landed, finish it; different content → collision, refuse; more than one → incident); 2 `POST LOG_Table/rows/add` with `[[maxId+1, operationId, dateSerial, type, category, amount, description, chequeNum, notes]]`; 3 `GET LOG_Table/dataBodyRange` and find the row by operation id; 4 `PATCH LOG_Table/rows/itemAt(index=n)/range` numberFormat for Date and Amount, single attempt, a failure becomes a warning (values are saved); 5 uniqueness of the number (§3b); 6 rebuild LOG_Sorted | The row with the operation id exists exactly once, its content equals what was sent, its number is unique. |
| Check workbook (after an uncertain add, or automatically after a reload) — **read-only** | `GET LOG_Table/dataBodyRange`; if the row is there: `GET LOG_Table/rows/itemAt(index=n)/range` (number formats) and `GET LOG_Sorted_Table/dataBodyRange` | Outcome `missing` (Save will send it), `landed` (everything confirmed → Saved), `incomplete` (the row is there but formats / number / sorted table are not confirmed → "Not finished", Save finishes them) or `conflict` (another entry carries the id). Zero writes, proven by tests that count requests. |
| Correct transaction | …then §8 report row visibility for **both** the old and the new month. 1 `GET LOG_Table/dataBodyRange`; locate by identity, compare fingerprint with the loaded copy → mismatch or absent = "Someone else changed this record", nothing written; 2 `POST LOG_Table/rows/add` with the corrected copy (**same TransactionID and Timestamp**); 3 `GET` and verify exactly one corrected copy exists (else incident); 4 numberFormat touch; 5 checked delete of the old copy: `GET LOG_Table/dataBodyRange`, evaluate that the old copy still sits at its index (else re-read and re-aim, up to three times), then `DELETE LOG_Table/rows/$/itemAt(index=n)`; 6 `GET` and assess by identity (§3b); 7 rebuild LOG_Sorted | Read-back shows one copy with the new fingerprint and none with the old. The corrected row is now the last row of the table (the month sheets sort by date, so their output is unchanged). |
| Check workbook (after an uncertain correction) — **read-only** | `GET LOG_Table/dataBodyRange` (+ formats and sorted-table reads when the corrected copy exists) | `missing` / `landed` / `incomplete` (e.g. "removal of the old copy" still to do; Save finishes it) / `conflict` (record gone or changed by someone else). Whenever a correction completes, the page's stored record for that dialog takes the **corrected row** as its reference, together with any newer text typed meanwhile, so a later Save (or a save after a reload) is compared against the version that actually reached the workbook and not against the one from before (review finding U1). |
| Delete transaction | …then §8 report row visibility for the month of the deleted date. 1 `GET LOG_Table/dataBodyRange`; locate + fingerprint check as above; 2 checked delete as in step 5 above; 3 `GET` and assess by identity (§3b); 4 rebuild LOG_Sorted | The fresh read shows exactly the target at the index before the DELETE is sent; afterwards exactly the target is gone. |
| Rebuild LOG_Sorted (internal, after every change) | Up to three rounds of: 1 `GET LOG_Table/dataBodyRange` (fresh; non-blank rows sorted by Date, ties in table order); 2 `GET LOG_Sorted_Table/dataBodyRange`; 3 identical → done; else `PATCH worksheets('LOG_Sorted')/range(address=<overlap>)` values + numberFormat, `POST LOG_Sorted_Table/rows/add` for extra rows or `DELETE LOG_Sorted_Table/rows/$/itemAt(index=i)` for surplus rows (from the end). After the rounds, one more fresh comparison. | "Consistent" means the sorted table equals the **freshly re-read** LOG at the end. If still inconsistent the app shows "The monthly sheets may be out of date" and Diagnostics reports it; the next change or Refresh rebuilds again. |
| Month view | `GET …/worksheets('<Month>')/range(address='A2:O40')` | Displayed as the workbook computes it; the app computes the same figures from the loaded rows and flags any difference. |
| Annual view | `GET …/worksheets('Annual 2026')/range(address='A3:N19')` | Displayed as the workbook computes it; total income, total expenses, net, opening and closing balance are compared with the app's own calculation and any difference is shown. |
| Change prior-year balance | 1 `GET ENTRY!B23`; compare with the value loaded → mismatch = conflict; 2 `PATCH worksheets('ENTRY')/range(address='B23')` values; 3 read-back | Value equals what was sent. Sheet protection permits it: B23 is an unlocked cell. |
| Check whether an incident is resolved — **read-only** | `GET LOG_Table/dataBodyRange` | The documented conditions for that incident (§3b) are evaluated and shown one by one. |
| Explicit repairs (only from the paused band, each behind a confirmation) | Restore a named row: `POST LOG_Table/rows/add` with the values this page last saw, then read-back (refused if a row with that identity already exists). Remove a copy: checked delete of one copy identified by identity + exact content (for several identical copies, exactly one, with the count proven afterwards). | Followed by an automatic read-only resolution check. |
| Recalculate | `POST …/application/calculate` exists in the client but is **not called** by any screen. | — |
| Prove access (Diagnostics, test copy only) | load → record the starting state → add one row with Description `TPAHA Books connection test` → read back → check LOG_Sorted contains it → delete it → confirm gone → compare with the starting state | Compared: LOG values + number formats, LOG_Sorted values + number formats, `ENTRY!B23`, formulas of `January!A4:O40`, formulas of `Annual <year>!A3:N19`, both row counts. Not compared: other sheets, formatting outside those ranges, table styles, names, protection, validation, workbook-level features. Refused while any read-only check fails or the workbook has an unresolved incident. |

Never used: uploading a workbook file, `PUT …/content`, changing formulas, changing protection,
renaming tables or sheets, writing to month or annual sheets, `$batch`.

## 3b. Concurrency boundary, incidents, repairs and the pilot writer model

**What Graph gives us** (Microsoft's guidance, verified 2026-09-06: workbook sessions, one request at a
time, no conditional writes, no transactions):

- **No lock, no conditional write, no transaction** for Excel ranges or tables. Another writer's request
  can land between any two of ours. The race window is the gap between our evaluated read and our
  positional request, whatever its length.
- **Appends are positional-safe**: `rows/add` never moves another row. Two members adding at once both
  land. If both drew the same transaction number, the page that reads the duplicate renumbers **its own
  just-added row** to the next free number after a fresh read evaluated first (our row must still sit at
  that index), then verifies. At most three rounds; then an incident (`duplicate-number`). If a
  pre-existing row received the number instead, an incident (`wrong-row-renumbered`) names it.
- **Deletes and renumbering are positional**: `rows/$/itemAt(index=n)` (the delete) acts on whatever row is at that index when
  the request runs. The evaluated read removes every shift that happens before it; it cannot remove a
  shift that happens after it.

**Three classes of outcome, stated plainly** (tests in `tests/unit/ledger-workbook.test.js`,
`tests/unit/concurrency.test.js`, `tests/e2e/review-regressions.spec.js`):

1. **Prevented**: a shift that happens before the check read (another member deleted or corrected a row
   moments earlier) is seen by the fresh read; the app re-aims or stops; every transaction survives.
2. **Detected, not prevented**: a shift between the evaluated read and the DELETE/PATCH itself. The
   request hits a different row. The app detects this by identity (never by position), names the row
   with the values it last saw (`victimLastKnownOnly`), records an **incident**, pauses all writes, and
   repairs nothing by inference. **The unrelated row is lost until a member restores it explicitly**; if
   someone else had changed that row after this page last read it, that change is recoverable only from
   Excel's version history. Reporting this is not the same as preserving the row; it is why the pilot
   runs with one writer.
3. **Unverifiable**: the read-back cannot say which row was affected (victim unknown). The incident says
   so; only an explicit acknowledgement after comparing with version history lifts it.

**Incidents** (`UnresolvedOperationError`, kinds `wrong-row-deleted`, `wrong-row-renumbered`,
`duplicate-copy`, `duplicate-marker`, `duplicate-number`, `integrity`):

- Recorded by the adapter and **kept in the browser** (per tenant, account and workbook). A page reload
  restores it; Refresh re-reads the table but never lifts the pause; every write is refused while it is
  unresolved (the page shows "Changes are paused").
- **Resolution is verified**, never assumed: "Check whether this is resolved" re-reads the table and
  evaluates the conditions for that kind (no identity appears twice; the named row is present exactly
  once; the corrected record has exactly one copy; the operation id appears once; the number is used
  once; a renumbered row has its number back). All conditions true → resolved (verified), writes resume.
  Any false → still paused, the failing checks are listed. Any unverifiable (unknown victim) → the member
  may **acknowledge** after comparing with version history; acknowledgement is refused while a check fails.
- **Repairs are explicit**: *Restore #N as last seen* (re-append the named row's last-seen values; refused
  if a row with that identity exists) and *Remove this copy* / *Remove one of the identical copies*
  (checked delete of exactly one copy, count proven afterwards). Each runs only after a confirmation and
  is followed by an automatic read-only resolution check. Nothing is rolled back or recreated on its own.
- An unfinished **correction** is also persisted (record, intended change, operation id) so that a reload
  reopens it, checks read-only, and lets Save finish it (remove the old copy) rather than append again.

**Pilot writer model** (`config.pilot.writerModel = 'single'`):

- Because class 2 cannot be prevented from a browser, the pilot supports **one designated writer at a
  time** and says so on every screen. It does **not** support simultaneous editing.
- Enforced where a page can: before the first Correct or Delete of a session the member confirms nobody
  else is writing (another member, another device or tab, Excel, the old script buttons); a second tab of
  the same browser is detected and refuses to write, **on the ledger and on the Diagnostics page alike**
  (both join the same tab guard; the first TPAHA Books tab opened in a browser owns writing, the others
  read only, whichever page they are); within one page, one operation runs at a time (above); the banner
  states the rule. A browser without BroadcastChannel cannot detect other tabs; the page says so. Other devices, Excel and
  the old scripts cannot be detected by a page: the rule for them is an agreement, documented in the
  setup guide, and the incident machinery above is the safety net if it is broken.
- Options for going beyond one writer, each a board decision, none adopted here: (a) keep one writer at
  a time by agreement (current pilot); (b) route every change through one path by removing the old
  buttons and locking `LOG`, then add a serialising backend for the app (it would cover only writers that
  use it); (c) change the workbook so that no positional operation is needed (soft-delete via a status
  column; corrections and deletes become appends and single-cell writes located by identity; the month
  and annual formulas would have to exclude voided rows). The re-review's request to "present the choice
  rather than silently deciding" is answered in `IMPLEMENTATION_STATUS.md`.

## 4. Workbook identity

The pilot copy is named "TPAHA_2026 (1)" in the owner's OneDrive for Business. Before the first live operation the Diagnostics page lists candidate files from
`GET /me/drive/root/search(q='TPAHA_2026')` and `GET /me/drive/sharedWithMe`, showing name,
extension, folder path, size and last-modified time. The member confirms the `.xlsx` file; the
app then addresses it by drive id and item id (stored in that browser only, per tenant and
account). Ids are opaque and are not sharing links.

Done on 2026-09-11 by an assigned board account: the picker listed the intended `.xlsx`, it was confirmed, and
every content request that followed was a GET that returned 200. The only non-GET in that run is the workbook
session this app opens before its first read (`POST …/workbook/createSession`), which changes no cell.

## 5. Findings to report (not changed by the app)

Verified independently on the 2026-09-06 local copy:

- The Annual sheet problems reported on 2026-09-04 (Total Expenses summing itself in row 14,
  Net using an empty cell in N16, uncalculated February–December cells) are **fixed in this copy**:
  row 14 is `=SUM(x10:x13)`, N16 is `=N8-N14`, and all cells have values. The year total was checked against
  the app's own calculation and agreed; the figure itself stays in the local notes rather than here.
- The duplicate TransactionIDs reported on 2026-09-04 are **fixed in this copy**; the numbers are in the
  local notes.
- Remaining risk, not a defect: the month sheets show at most 30 transactions per month (rows
  4–33). The app shows all rows and flags a month that exceeds 30.
- Remaining dependency: the month sheets need `LOG_Sorted_Table` to be kept sorted. A future
  workbook change could replace it with a `SORTBY` formula, which would remove this app-side step.
  That is a workbook change and needs approval; it is not part of the integration.

Whether the **online** copy carries the same fixes is unknown until Diagnostics reads it.

## 6. Winter storage workbook

Not mapped for live use yet. The earlier extraction (`tools/extract_xlsx.py`) already lists its
entry cells per space (from the Complete Storage Swap script) and its formulas; the live test
workbook for storage is to be confirmed separately.

## 7. The workbook's own Office Scripts (source read 2026-09-10)

Cody supplied the source of the three ENTRY buttons. What matters for this app:

**They do not use the Timestamp column to identify anything.** SubmitEntry writes its own
`new Date().toISOString()` into it and never reads it back. FindTransactions reads whole rows but matches and
displays only Date, Description, Amount and TransactionID. DeleteTransaction finds its row by TransactionID and
confirms it against the Date, Description and Amount shown in `ENTRY!A17:D20`. The nine-fraction-digit
operation ids this app writes (§2) therefore do not affect them, and the concern carried since the first
review is closed.

**But every save and delete rewrites both tables completely.** Their shared `writeRows()` adds or deletes rows
until the count matches, then calls `setValues()` over the whole body of `LOG_Table` and `LOG_Sorted_Table`.
Two consequences:

- Rows written by this app are re-written by the scripts using the values they just read, including the
  Timestamp text. A `getValues()`/`setValues()` round trip preserves text, which is how the existing
  three-digit timestamps survive today. That the nine-digit form behaves the same is **not established for the real
  service**, and the ordinary Diagnostics connection test cannot establish it: that test never runs an Office
  Script, so it proves only this app's own write and read-back. The unit test that exercises a whole-table
  rewrite does so in the mock and proves the mock's behaviour only. Nor should a changed identity be assumed to
  announce itself: this app would notice one only where it looks for its own row by operation id, and a rewrite
  that happened later, between operations, would surface as a record that looks changed or missing rather than
  as an obvious failure. A separate, separately authorised live check is set out in
  `docs/checkpoint-2-runbook.md`.
- A script running while this app sits between its checked read and its positional request can move every row
  at once, not merely one. That is the concrete reason for the one-writer rule in §3b.

**They refuse to run at all unless the workbook passes two checks**, so anything this app leaves behind must
satisfy both:

| Script check | What it demands | How this app satisfies it |
|---|---|---|
| `validateLog` | every non-blank row has a unique, positive, safe-integer TransactionID; Date is an integer serial **inside the workbook year**; Type is Deposit or Withdrawal; Category matches that Type in LISTS; Amount is positive with at most two decimals | The entry form enforces the same rules before a save: `validateEntry` in `site/js/ledger/model.js` checks the year, and `entryToRow` writes an integer serial and rounds the amount to two decimals. Unit test: "every row this app writes passes the scripts' own validateLog, after an add, a correction and a delete". |
| `writeRows` (the shape it leaves) | both tables keep **at least one body row** — `Math.max(1, rows.length)` — so an empty ledger leaves one blank row, never a table with no body at all. `transactions()` filters that blank row out before every comparison, so it reads as zero transactions | This app keeps the same invariant: `sortedBodyRows()` returns the sorted rows, or one blank row when there are none, and the verification counts transactions rather than body rows. **This was a divergence, found on reviewing `RefreshReports` and fixed 2026-09-12**: the rebuild previously deleted `LOG_Sorted` down to zero rows when the last transaction was deleted, a shape the workbook's own scripts never produce and whose effect on the month formulas and on Graph's `dataBodyRange` was untested. Unit tests: the six "empty ledger" tests. |
| `verifySorted` | `LOG_Sorted` equals `LOG` sorted by **Date serial, then TransactionID**, compared as whole rows. One rule, `expectedSortedRows()`, is exported from the adapter and used by both the rebuild and the read-only Diagnostics check, after the two briefly held separate copies and disagreed (review finding V1) | This app now sorts LOG_Sorted the same way. **This was a defect, fixed 2026-09-10**: it previously broke ties by table position, which diverges as soon as a correction moves a row to the end of LOG, and the scripts would then have refused to run until someone ran RefreshReports. Unit test: "LOG_Sorted keeps the order the scripts require … including after a correction moves a row to the end of LOG". |

**Two differences that are not defects but must be known:**

- **Hidden rows on the month sheets.** After a save or delete the scripts run `refreshMonthlyVisibility()`,
  which hides every month-sheet row 4–33 whose date cell is empty and shows the rest. This app never writes to
  the month sheets, so a transaction it adds can land in a row left hidden by an earlier script run. The web
  pages are unaffected, because they read the cells directly whether hidden or not, but **someone opening or
  printing that month sheet in Excel could miss the row** until a script button is pressed. A decision for the
  board. **Since taken, and built:** the app now sets those rows itself (§8), and `RefreshReports` — reviewed
  2026-09-12 — is the manual fallback when it cannot. The app needs no password to do it, because the month
  sheets allow row formatting; the script uses the password only because it unprotects first.
  Telling members to press Submit or Delete merely to refresh visibility is not acceptable: those buttons change
  data.
- **Thirty transactions a month.** SubmitEntry refuses to add a 31st transaction in a month. This app allows it
  and warns that the month sheet displays at most 30. A 31st added here does not fail the scripts' checks, but
  SubmitEntry will then refuse to add more in that month.

**The fourth script, `RefreshReports`, read on 2026-09-12.** It is the repair button the other two scripts point
at ("Run RefreshReports before entering or deleting transactions"), and it does exactly four things: rebuild
`LOG_Sorted` from `LOG` with `writeRows(sorted, ordered(rows))`, recalculate, `verifySorted`, and
`refreshMonthlyVisibility()`. It **never calls `addRow` or `deleteRowsAt` on `LOG_Table`** and prints "No
transactions were added or deleted."

Two things follow, and both matter here:

- **It is safe to recommend as a repair**, unlike Run SubmitEntry and Run DeleteTransaction. It is the right
  answer when the month rows are wrong and the website cannot fix them (§8).
- **It is still a writer.** It rewrites `LOG_Sorted` and the month sheets' row visibility, so pressing it while
  a website save is in flight is a second writer on the same tables, exactly like the other two buttons.
- **It cannot prove the website's row identity survives a script.** That question is about `LOG_Table`, which
  only `SubmitEntry` and `DeleteTransaction` rewrite. `RefreshReports` leaves that table untouched, so it is
  no evidence either way; the separately authorised check in the runbook must use one of the other two.

It also refuses to run at all if `validateLog` fails — every transaction must have a unique positive whole
number, a date inside the workbook's year, a Deposit/Withdrawal type, a category matching that type in LISTS,
and a positive amount of at most two decimal places. Anything this app writes must satisfy that or the
workbook's own buttons stop working; the unit test "every row this app writes passes the scripts' own
validateLog" covers it, and the ported helpers were re-checked against this third copy of the shared block on
2026-09-12 — identical apart from indentation.

**Security note, unrelated to this app:** `refreshMonthlyVisibility()` reads the worksheet-protection password
from a cell on the ENTRY sheet and uses it to unprotect and re-protect each month sheet. That password is
therefore stored in plain text in the workbook, readable by anyone who can open it. The exact cell is recorded
in the local note beside the script source, not here. This app never touches that cell; the only ENTRY cell it
reads or writes is `B23`, the prior-year balance. Excel sheet protection is a guard against accidental edits
rather than real security, but the treasurer should know it is there.

The source of the three scripts is kept for review in the git-ignored `data/office-scripts/` folder, with a
note on why it is not in source control.

## 8. Monthly report row visibility

**The problem.** Rows 4 to 33 of each month sheet are formulas over `LOG_Sorted_Table`; a row shows a
transaction when its date cell has a value and is blank otherwise. The workbook's own scripts finish every save
by running `refreshMonthlyVisibility()`, which hides the blank rows and shows the rest. This app did not, so a
transaction it added landed in a row an earlier script run had hidden: correct on the website, missing when the
month sheet was opened or printed in Excel.

### What Microsoft Graph actually supports (checked 2026-09-12)

| Need | Graph v1.0 | Evidence |
|---|---|---|
| Change row visibility | **Yes.** `PATCH …/worksheets/{id}/range(address='A4:A9')` accepts `rowHidden` (Boolean) in the request body, alongside `values`, `formulas` and `numberFormat`. Delegated permission: `Files.ReadWrite` | [Update range](https://learn.microsoft.com/en-us/graph/api/range-update) |
| Read row visibility back | **Yes.** The same range resource exposes `rowHidden`, which reads true when every row in the range is hidden, false when none is, and null for a mix | [Update range](https://learn.microsoft.com/en-us/graph/api/range-update) |
| Recalculate | **Yes.** `POST …/workbook/application/calculate` | already in the client |
| Read the report cells | **Yes.** Ordinary range reads | §3 |
| Unprotect a sheet **with a password** | **No.** `POST …/protection/unprotect` takes no request body at all | [WorksheetProtection: unprotect](https://learn.microsoft.com/en-us/graph/api/worksheetprotection-unprotect) |
| Re-protect **with a password** | **No.** `POST …/protection/protect` takes `options` only | [WorksheetProtection: protect](https://learn.microsoft.com/en-us/graph/api/worksheetprotection-protect) |
| Read the protection options | Yes, but `options` and `protected` are **read-only** | [workbookWorksheetProtection](https://learn.microsoft.com/en-us/graph/api/resources/worksheetprotection) |

Office Scripts and Graph are **not** the same surface. `protection.unprotect(password)` and
`protection.protect(options, password)` exist in Office Scripts; the Graph equivalents take no password. A
straight port of `refreshMonthlyVisibility()` is therefore impossible: this app could not unprotect a
password-protected sheet, and even if it could, re-protecting would silently drop the password.

### What the workbook itself allows (read-only inspection of the local copy, 2026-09-12)

| Sheet | Protected | Password | "Format rows" |
|---|---|---|---|
| January … December | yes | **no** | **allowed** |
| ENTRY | yes | yes | blocked |
| Annual, LOG, LOG_Sorted, ConfigHidden, LISTS | see §1 | no | blocked or unprotected |

The month sheets are protected with row formatting **allowed**, which is precisely the permission needed to
hide and show a row. So the unprotect step is not needed at all.

**This is the local copy, and it does not prove the online one.** The pilot workbook may carry different
protection. There is no need to guess: the first save that formats rows on the online copy either succeeds, in
which case row formatting is allowed there too, or is refused, in which case the app reads the sheet's
protection back and says so in as many words. The runbook makes that the live check.

### The approach taken
After the transaction and the sorted helper table are verified, and only then:

0. Only once the sorted helper table has been **verified**. The report rows are formulas over that table, so
   formatting from an unverified one would hide or show the wrong rows and then report success. When it cannot
   be verified, nothing is touched and the months are recorded as unfinished work instead.
1. `POST …/workbook/application/calculate` once, so the report formulas reflect the change before they are read.
2. For each affected month (the month of an added or deleted transaction; **both** months for a correction that
   moves one): `GET …/worksheets('<Month>')/range(address='A4:A33')?$select=address,values`.
3. Group the rows into runs that share a visibility and `PATCH …/range(address='A<first>:A<last>')` with
   `{ "rowHidden": true|false }`. A month normally needs two requests: one for the populated block, one for the
   blank rows below it.
4. Verify each run by reading `rowHidden` back; a mixed answer, or the wrong one, is a failure.

**It never unprotects a sheet, never reads or writes the cell holding the protection password, never changes a
protection option, and never touches `LOG_Table` or `LOG_Sorted_Table`.** Row visibility is the only thing it
changes, which is why the retry below cannot duplicate a transaction however often it runs. It runs inside the
same operation queue and the same single-writer tab guard as everything else, and never on page load.

### When it does not finish
The months affected are written into this browser **before the first workbook change**, not after it, so a page
closed midway still knows what is outstanding; they are cleared only once the formatting is verified. If the
browser refuses to keep that record, the member is told plainly rather than left believing it was kept.

An operation that turns out to write nothing — a submit that was ignored as a duplicate, or one that stops at a
conflict — takes back **only the months it added itself**. If September was already outstanding from an earlier
change, a later September operation that never writes leaves that reminder exactly as it found it, wording and
all. Nothing but verified formatting clears a month, because nothing else has fixed the rows.

The transaction is already in the workbook and is never written again. The screen says **"Transaction saved.
Excel report formatting still needs updating."** (for a delete, "Transaction deleted."), the months still to do
are kept in this browser so a reload does not forget them, and a **Finish report formatting** button finishes
them. That button repairs the sorted helper table first when that is what is blocking, then
formats; the banner says which of the two it is. Either way it writes to the helper table and the month sheets
only and **never to `LOG_Table`**, so it cannot add or remove a transaction however often it runs. Saving is not paused: unfinished formatting is cosmetic, not damage to a
record. "Check workbook" and the check that runs after a reload stay read-only and have nothing to do with it.

### The limitation that remains
This depends on the month sheets keeping "Format rows" allowed. If someone re-protects them with that
disallowed, the `PATCH` fails and **the app cannot repair it**: Graph cannot supply a password to unprotect,
and changing a protection option is not something this app should do uninvited. The app says so in the message,
naming the sheet and the setting, because it reads the protection back to explain the refusal.

The recovery is in Excel, and `RefreshReports` is now a reviewed answer for it (§7): that script unprotects each
month sheet with the workbook's own password, sets rows 4–33 by the same rule this app uses, and re-protects
with the options it found. It adds and deletes nothing. Press it **only when no website save is in flight**, as
it writes `LOG_Sorted` and the month sheets.

Two other recoveries work: allow "Format rows" again on the month sheets and press **Finish report formatting**
on the website, or show the rows by hand in Excel. **Do not tell anyone to press Run SubmitEntry or Run
DeleteTransaction for this**: those buttons add and delete transactions.

### The alternative that was assessed and not taken
Leaving all thirty report rows visible on every month sheet would need no protection permission and no
visibility writes. It was rejected: every month sheet would then print thirty rows regardless of content,
mostly blank, which changes what the treasurer's monthly report looks like on paper. Hiding blank rows is the
workbook's existing design, and the approach above matches it rather than replacing it.
