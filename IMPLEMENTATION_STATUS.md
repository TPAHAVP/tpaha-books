# TPAHA Books — implementation status

Working folder: `C:\Users\Cody\Desktop\Projects\excel to HTML idea` (the implementation Codex reviewed and
the only one maintained). Another copy exists under `Desktop\Codex projets\excel to HTML idea`; it is not
deleted, replaced or written to without Cody's instruction.

Rules kept throughout: no passwords, tokens, client secrets, sharing links or real records in this file
or in the repository; nothing has connected to Microsoft 365; no production workbook has been touched.

---

## 2026-09-20 (fourth) — Correction to the entry below, from the actual request log. **For review; not published; no workbook change**

The reviewer supplied the request log of the failed run. It reads: `DELETE /workbook/tables/LOG_Table/rows/26 →
400`, with the message "The API you are trying to use could not be found. It may be available in a newer
version of Excel." **The status was 400, not 404, and the log recorded no error code.** The entry below, as
first written, said "404 `ApiNotFound`" — a status and a code taken from Microsoft Q&A threads describing
other people's failures, not from this log, which had not reached me. That was wrong to write down as fact and
is corrected everywhere: the client comment, the mock, the tests, the mapping, the runbook and the entry below.

**The DELETE form is reconciled with Microsoft's own example.** "Working with Excel in Microsoft Graph" — the
official overview — deletes a table row as `DELETE …/tables('4')/rows/$/itemAt(index=6)`, answering 204 No
Content. That is the one published DELETE-by-position form, so the app now reproduces it exactly, `$` segment
included. The range PATCH keeps `rows/itemAt(index=n)/range`, the reference form that succeeded in the live
run. The fresh identity-and-content check before the send, the exactly-once send and the read-back are
unchanged.

**The mock claims nothing it cannot know.** `rows/N` is refused with 400 and the logged message, code blank.
Only `rows/$/itemAt(index=n)` deletes, answering 204 as the example does. `DELETE rows/itemAt(index=n)` — a
form no document shows — is refused with a code that names the mock (`MockNotModelled`), so the app cannot
drift onto it and no one mistakes that refusal for the service's. Whether the service accepts the published
form is for the live cleanup of #30 to show.

**So the next failure is recorded, not reconstructed:** the client now logs a refusal's error code, inner code
and message alongside its status, and Diagnostics' Copy log shows them.

### Tests — 158 unit (was 156), 189 browser, 21 config, all passing
Changed: the replays assert 400 and the logged message, and no code; every DELETE the app sends must match the
published form and no PATCH may leave the reference form; the mock's refusal of `rows/N` is asserted at 400
with the message only. New: the reference-form DELETE is refused as not modelled while its range PATCH still
works; a refused request is logged with status, code, inner code and message, and a success carries none.

Not published. The workbook is untouched since the refused delete.

---

## 2026-09-20 (third) — **First live write test: the DELETE was refused.** Cause found, fix made, mock corrected. **For review; not published; no workbook change**

### What happened
The connection test (runbook Part 2 step 4) was run on the test copy. Sign-in, picker, all six read-only
checks, the add of the test row, its read-back and the sorted-table check passed. **The delete was refused
with 400**: `DELETE /workbook/tables/LOG_Table/rows/26 → 400`, "The API you are trying to use could not be found.
It may be available in a newer version of Excel." The log recorded the status and message and **no error code**.
The app sent the DELETE once, retried nothing, paused nothing, and reported that the row was still present —
the designed behaviour for a clean refusal, and it held. The test row remains as **transaction #30**. Its add
also rebuilt `LOG_Sorted` and set December's report rows.

*(Corrected 2026-09-20 from the request log. The first draft of this entry said "404 `ApiNotFound`", taken from
other people's reports; see the correction entry above.)*

### Cause
The app sent `DELETE …/tables/LOG_Table/rows/{index}` — exactly the form the reference page for "TableRow:
delete" documents. The live service does not accept it. The evidence (mapping §3, with links):
- `workbookTableRow` has **no `id`** — only `index` and `values`. `rows/{index}` asks for a row by a key the
  resource does not have; the SDK snippets on the delete page itself key on a row *id*.
- **Microsoft's own worked example** in "Working with Excel in Microsoft Graph" deletes a row by position as
  `DELETE …/tables('4')/rows/$/itemAt(index=6)` → 204 No Content — the one published DELETE-by-position form.
- The reference form for reading a row and its range is `rows/itemAt(index=N)` / `…/range`.
- Microsoft Q&A carries the same message for both DELETE and PATCH on `rows/{index}` (reported there with the
  code `ApiNotFound`; our log recorded no code), with `rows/$/ItemAt(index=…)` as the working form.
- In the **same live run**, the app's `PATCH …/rows/itemAt(index=N)/range` succeeded.

**Why the mock never objected:** it accepted both `rows/N` and `rows/itemAt(index=N)`. The tests proved the
app's logic, not the service's routing — which the status file has said all along, and this is what that
caveat looks like when it bites.

### The fix (`site/js/workbook/excel-client.js`, `site/js/workbook/mock-excel.js`)
- One form per operation, each from its own document: the DELETE reproduces the worked example exactly,
  `rows/$/itemAt(index=N)` (expecting 204); the range PATCH keeps the reference form that succeeded live,
  `rows/itemAt(index=N)/range`.
- **Unchanged, deliberately:** the fresh read evaluated before the DELETE, the identity-and-content check at
  that index, the exactly-once send, and the read-back that decides an uncertain answer. A 4xx refusal is
  `failed` (thrown, nothing retried, no incident); a lost answer is `ambiguous` (resolved by reading).
- **The mock now refuses `rows/N`** with 400 and the logged message, for DELETE, GET and PATCH, its error code
  left blank rather than guessed. It deletes only through the published `rows/$/itemAt(index=N)` form
  (answering 204) and refuses `DELETE rows/itemAt(index=N)` with a refusal labelled as the mock's own. It also
  gained a `method` filter on injected failures, so a replay can target the DELETE and not the number-format
  PATCH that shares the path.
- **The request log now records a refusal's error code, inner code and message**, not only its status, so the
  next live failure is recorded rather than reconstructed from other people's reports.

### Tests — 156 unit (was 152), 189 browser, 21 config, all passing
- Three tests that hard-coded `rows/N` were corrected to the working form; the R5 failure-injection regex too.
- **New:** `rows/N` is refused for every method and changes nothing; every DELETE the app sends — a delete, a
  correction's cleanup, a helper-table shrink — uses `itemAt` and never `rows/N`; the live refusal replayed
  against a delete (one DELETE, intact row, no incident, the fresh check still first, clean reload); and the
  live refusal replayed against the connection test itself (stops at "Delete the test row", `restored: false`,
  no later step runs, the test row remains exactly once).
- Honest limit: **the mock now encodes the evidenced form; it does not prove Microsoft accepts
  `DELETE …/rows/itemAt(index=N)`.** The documentation lists no DELETE form for `itemAt`. The targeted cleanup
  below is the proof, one row at a time.

### Targeted cleanup of #30 — runbook Part 2a, not yet authorised
Read-only first (website: exactly one test row, #30, no bands, count 26; Excel: #30 last in LOG, December row
4 showing it). Then **one** checked delete of #30 from the website, which also rebuilds `LOG_Sorted` and
re-hides December. Then verify read-only (25 on the website and in Diagnostics; #30 gone from both tables in
Excel; December as before). If the delete is refused again, stop and send the log — the fallback is the
workbook's own Run DeleteTransaction with 30 in the ENTRY form, pressed only on the reviewer's word. The
connection test must **not** be used for cleanup: it adds a new row and deletes that one. Only after #30 is
gone may the connection test be re-run, on a separate go-ahead.

Not published. The workbook has not been touched since the refused delete.

---

## 2026-09-20 (later) — **Published.** Part 0 run in full; every verification passed. **No workbook write**

Cody authorised publishing. Runbook Part 0 was followed step by step:

| Step | Result |
|---|---|
| 0.1 tree and suites | `git status` empty; `npm test` 152/0; Playwright 189 passed; `verify:config` 21/0 |
| 0.2 contents | 8 commits `9e464cd`…`e5a913a`, 15 files: site code, tests, docs. Private-file check (`data/`, `.xlsx`, scripts) printed nothing. After-hashes re-derived at HEAD matched the runbook table |
| 0.3 push | `6e19eb0..e5a913a  main -> main` |
| 0.4 deployment | `gh` is not installed on this machine, so the live site was polled instead: `js/save/report-formatting.js` went from 404 to **200 about ten seconds after the push** |
| 0.5 served files | `report-formatting.js` 200 `49c85f6f…02cf0` ✓ · `ledger-workbook.js` 200 `2cddfca7…261a` ✓ (was `20cd704c…b7ea`) · `config.js` 200 `bc25ac6c…ce48` ✓ unchanged · `ledger.html` 200 with the report band element present |
| 0.6 | Nobody signed in |

**Deployed commit: `e5a913a`.** The live site at `https://tpahavp.github.io/tpaha-books/` now carries the
monthly report-row feature and every review fix through the identity-check cleanup.

No workbook write has been run. Parts 2, 2b and 3 wait on a separate authorisation.

---

## 2026-09-20 — Script-identity check: cleanup completed, and made conditional. **Nothing published; no live write**

Reviewer's correction: the separately authorised script-identity check deleted the Excel throwaway afterwards
but never the website's own test entry, so a "clean" run would have left one test transaction in the workbook.
It also did not say what to do with cleanup when the comparison fails.

Fixed in `docs/checkpoint-2-runbook.md`:
- **Step 0** records the transaction count before anything is added, so "back to the original" has a number.
- **After an identical comparison** with no paused band and no incident: delete the website's entry from the
  website, delete the Excel throwaway from the website too (one writer, no further script runs), Refresh and
  confirm the count matches step 0, confirm in Excel, report both deletions and both counts. Deleting the
  website's row is itself evidence — the app finds it by the identity that was just rewritten.
- **If the text differs, or a paused band or incident is showing: stop before cleanup.** Nothing is deleted,
  no Refresh, no retry, no workbook button. The rows are the evidence and deleting them destroys it; the
  reviewer says how cleanup happens.
- Part 3.3 now refers to that cleanup and covers the stopped case.

Documentation only; no code changed. Tests unchanged from `7740e6c`: 152 unit, 189 browser, 21 config.

---

## 2026-09-12 (seventh) — Deployment verification now uses a file this release actually changes. **Nothing published; no live write**

Three corrections to the walkthrough, at Cody's direction.

**1. The deployment check was verifying the wrong file.** Part 0.5 hashed the deployed `js/config.js` against
the committed one. That file is **not changed by this release**, so it matches whether or not the new code
deployed — the check would have passed over a deployment that never happened, and the write test would then
have run against the old code. Prerequisite D, which records the same check from 2026-09-11, now says so too.

Part 0.5 now uses files this release does change, with the before-values measured against the live site today:

| File | Before | After |
|---|---|---|
| `js/save/report-formatting.js` | **404**, the file does not exist on the live site | 200, `49C85F6F…02CF0` |
| `js/workbook/ledger-workbook.js` | 200, `20CD704C…B7EA` | 200, `2CDDFCA7…261A` |

The new file's presence proves the new code is live; the changed file's hash proves the old build was replaced
rather than a cached copy served. `js/config.js` is still checked, but as a thing that must **not** move. The
requests carry a `?nocache=` value so a cache cannot answer for the server, and the walkthrough says to
re-derive the hashes with `git show HEAD:…` if further commits land before it is run.

**2. Prerequisite B said "the three Office Scripts".** There are four; `RefreshReports` was read on 2026-09-12.
Corrected, with both dates.

**3. The "Hidden month rows — still to decide" note is gone.** It was settled and built, and prerequisite B
already said so two paragraphs above — the walkthrough contradicted itself. Removing it also restores the
count: Part 1 says five prerequisites and now lists five, A to E, where it had six.

No code changed; tests unchanged from `7740e6c`: 152 unit, 189 browser, 21 config.

---

## 2026-09-12 (sixth) — The deployment and test-copy walkthrough is written. **Nothing published; no live write**

Codex approved `7740e6c`. `docs/checkpoint-2-runbook.md` is now the whole procedure, one step at a time, and
three things were added or corrected.

**Part 0, publishing — prepared, not run.** Confirm the tree is clean and all three suites pass; list exactly
what will go out and prove nothing private is in it (`git diff --name-only origin/main..HEAD` filtered for
`data/`, `.xlsx` and the scripts must print nothing); push; watch the Pages workflow; then verify what is being
*served* rather than what was pushed — the site loads with no Test mode banner, and the deployed `js/config.js`
hashes to `BC25AC…CE48`, the reviewed file. Undo is `git revert` and push, never a force-push.

**Part 3, cleanup — new.** The connection test cleans up after itself and Part 2b's transaction is deleted by
hand; this part is how both are proven. On the website: the count matches what it started at, no
*TPAHA Books connection test* row, no Part 2b transaction, no green band, no paused band. In Excel: the table
as before, the month sheet hidden exactly as at Part 2b step 1, Version history accounted for, workbook closed.
Then sign out and close every other tab. A cleanup that cannot be completed is a finding to report, not
something to tidy away.

**A numbering defect fixed.** Part 2b had two steps numbered 4 — the RefreshReports fallback I added on
2026-09-12 and the original "check the row is visible in Excel". Someone following it on the day would have
done the fallback before the check it is a fallback *for*. Renumbered so the check comes first (step 4) and the
repair second (step 5), and the closing summary now refers to step 7 rather than step 6.

Part 4 (what to send back) now also asks for the Part 2b observations, whether RefreshReports was needed, and
the transaction count before and after.

No code changed; tests unchanged from `7740e6c`: 152 unit, 189 browser, 21 config.

---

## 2026-09-12 (fifth) — The empty `LOG_Sorted` divergence is fixed. **For review; not published, no live write**

The finding recorded in the entry below is closed. This app now keeps the invariant the workbook's own
`writeRows()` keeps.

**The rule, in one place.** `sortedBodyRows(logValues)` returns the sorted transaction rows, or — when there are
none — exactly one blank row. The rebuild writes that body; the read-only check compares against it. The
transaction rule itself, `expectedSortedRows()`, is unchanged and still returns no phantom row, so the two
questions "what should the table hold" and "which transactions are there" stay separate.

**Verification counts transactions, not rows.** `sortedTableMatches()` now returns `transactions` (real rows
wanted) and `present` (real rows found) beside `want`/`have`, and Diagnostics reports those. An empty ledger
reads as "consistent (0 transactions)", never as one row. That is the same arithmetic the scripts do:
`verifySorted` compares `ordered(log)` with `transactions(sorted)`, and `transactions()` filters all-blank rows.

**The placeholder is written as values only.** Giving an empty row the date and money number formats would be
writing formatting the workbook's own script never writes, so the rebuild omits `numberFormat` when every row
it is writing is blank.

### Tests — six new, all failing against the previous code
Verified by reverting `sortedBodyRows` to its old return value: **146 passed, 6 failed**; with the fix,
**152 passed, 0 failed**.

| Test | What it pins |
|---|---|
| deleting the last transaction | `LOG_Sorted` holds exactly one blank row, not none; `LOG_Table` also keeps a body row |
| the check accepts the placeholder | `ok: true`, `transactions: 0`, `present: 0`; and `expectedSortedRows([])` is still `[]` |
| loading and rebuilding an empty ledger | the ledger loads with no transactions; `rebuildSorted()` reports consistent, `changed: false`, `rows: 0`, and sends **no write at all** — no repair loop |
| a table left with no body row | repaired to the placeholder, and the writes carry **no `numberFormat`** |
| the first transaction afterwards | replaces the placeholder instead of adding beside it — one row, no leftover blank — and deleting it again returns to the placeholder |
| the Office Scripts' own checks | `validateLog` and `verifySorted` both accept the empty shape; `transactions(sorted)` is `[]`; the shape equals what `writeRows()` would leave, so pressing Run RefreshReports next changes nothing |

W1, W2 and W2b are untouched: their tests still pass unchanged.

### Also in this pass: what the script-identity check can and cannot use
`RefreshReports` never writes `LOG_Table`, so running it **cannot** show whether a script rewrite preserves the
website's row identity. Only `SubmitEntry` and `DeleteTransaction` rewrite that table, through
`saveChange()` → `writeRows(log, after)`, and both change transactions — which is why that check needs its own
authorisation and a test copy. The runbook now names the button to press, says to use a throwaway transaction
and delete it afterwards, and records that `RefreshReports` copying the Timestamp into `LOG_Sorted` is a
weaker, different signal that must not be logged as the identity check. Mapping §7 says the same.

### Tests (2026-09-12, this folder)
- `npm test`: **152 passed, 0 failed** (was 146).
- `npx playwright test`: **189 passed, 0 failed** (unchanged).
- `npm run verify:config`: 21 passed.

Not published, and no live write has been run.

---

## 2026-09-12 (fourth) — `RefreshReports` supplied and reviewed. **No code change; not published, no live write**

Codex approved `68db038` (W2b). Cody then supplied the fourth Office Script. It is saved beside the other three
in the git-ignored `data/office-scripts/`, and it has been read in full.

### What it does
`main()` reads `LOG_Table`, runs `validateLog`, writes `LOG_Sorted_Table` with `writeRows(sorted,
ordered(rows))`, recalculates, runs `verifySorted`, then `refreshMonthlyVisibility()`, and prints "Reports
refreshed. No transactions were added or deleted." It calls neither `addRow` nor `deleteRowsAt` on
`LOG_Table`. The claim in its own output matches its code.

### What it confirms
- **The row rule this app implements is the workbook's own rule**, not an inference: rows 4–33 of all twelve
  month sheets, hidden exactly when the row's date cell in column A is empty. `refreshMonthlyVisibility()` is
  the behaviour `_syncMonthVisibility()` reproduces over Graph.
- **The ported helpers are right.** The shared block appears a third time here; `diff -w` against SubmitEntry's
  copy is empty, so the ports in `tests/unit/ledger-workbook.test.js` are now checked against three copies.
- **Why the script needs the password and this app does not.** It unprotects each month sheet, sets the rows,
  and re-protects with `getSavedOptions()`. This app never unprotects: it relies on the month sheets allowing
  row formatting, which the local copy does. Still unproven on the online copy — Part 2b of the runbook.

### What it changes
1. **The recovery guidance now has a reviewed name.** Previously: allow "Format rows" again by hand, or show
   rows manually, with `RefreshReports` explicitly *not* recommended because it was unread. Now **Run
   RefreshReports** is the documented repair when the month rows are wrong and the website cannot fix them —
   it is the one workbook button that changes no transactions. Run SubmitEntry and Run DeleteTransaction remain
   off-limits for this.
2. **A new live-test rule.** `RefreshReports` is still a writer: it rewrites `LOG_Sorted` and every month
   sheet's row visibility. Pressing it during a website save is a second writer on the same tables. The runbook
   now says to press it only when no save is in flight, and to record that it was needed.
3. **The hidden-month-row decision is closed.** It was "accept the limitation or build it"; it was built, and
   `RefreshReports` is the fallback rather than the plan.
4. **The separate script-identity check is no longer blocked on an unread script.** All four have been read; it
   now waits only on Cody's authorisation.

### One divergence found, not yet fixed
`writeRows()` keeps `LOG_Sorted_Table` at **at least one body row**, blank when there are no transactions
(`Math.max(1, rows.length)`). This app's `_rebuildSorted()` deletes down to **zero** rows in that case.
Verified against the mock: deleting all thirteen sample transactions leaves `LOG_Table` with one blank row and
`LOG_Sorted_Table` with none.

Its reach is small and it cannot arise in the pilot — the workbook holds 25 transactions and the test plan never
empties it — but it is a real difference from the workbook's own invariant, and what a zero-row table does to
the month sheets' formulas, or to Graph's `dataBodyRange`, has not been tested. Recommended fix: match the
script and leave one blank row. **Done on 2026-09-12** — see the entry above.

No code changed. Tests unchanged from `68db038`: 146 unit, 189 browser, 21 config.

---

## 2026-09-12 (third) — Review fix W2b: a conflicting operation no longer erases older unfinished work. **For review; not published, no live write**

Codex reviewed `d5a0faa`, accepted W1, and found that W2 was still wrong in one case. Reproduced exactly as
reported: existing `[9]`, reserve `[9]`, release `[9]` → `null`.

**What went wrong.** `reserveReportFormatting` returned every month it was asked for, not only the months it
actually added. September could already be outstanding from an earlier change whose rows were never fixed; a new
September correction or delete reserved September again, hit a conflict before writing anything, and released
September — deleting a reminder that was still true. The rows stayed wrong and nothing said so any more.

**The fix.** A reservation now knows which pending work is its own. `ReportFormattingStore.reserve()` reads what
is already outstanding, merges the wanted months, and returns `added`: only those that were not outstanding
before. A release is given `added`, so it can only ever take back work that same operation created. An existing
record also keeps its own message and reason — a reservation is made before anything is attempted, so it knows
less than the completed operation that wrote the record did, and must not overwrite it. Clearing a month is
still only ever done by `remove()` after verified completion.

| Case | Before | After |
|---|---|---|
| September outstanding, September operation conflicts | reminder deleted, rows still wrong | reminder kept, wording kept |
| September outstanding, correction to October conflicts | both deleted | September kept; October, which that operation added and never wrote, released |
| Nothing outstanding, operation conflicts | released | released (unchanged) |
| Formatting verified complete | cleared | cleared (unchanged) |

### Regression evidence
- **Browser (the one asked for):** an existing pending September, then a conflicting delete in September. The
  conflict dialog appears, no transaction is deleted, the banner still names September with its original
  wording, and it is still there after a reload with **zero writes at startup**; pressing Finish report
  formatting then completes it with the transaction table untouched. A second test covers the cross-month
  correction. Both **fail against the previous code** — verified by restoring the old return value and
  re-running: 2 failed, then 2 passed with the fix in place.
- **Unit:** five tests on the store, including the reported sequence (`add([9])`, `reserve([9])` → `added: []`,
  `remove([])` → September survives), the cross-month case, the no-prior-work case, verified completion still
  clearing, and a storage refusal still reporting what it would have added.

### Tests (2026-09-12, this folder)
- `npm test`: **146 passed, 0 failed** (was 141).
- `npx playwright test`: **189 passed, 0 failed** (was 183) = 63 per project × 3.
- `npm run verify:config`: 21 passed.

Still not published, and no live write has been run.

---

## 2026-09-12 (later) — Review fixes W1 and W2 on the report-row work. **For review; not published, no live write**

Codex reviewed `9e464cd` and found two defects. Both are fixed, with regression tests that fail against the
previous code.

| # | Finding | Fix | Regression evidence |
|---|---|---|---|
| W1 | Formatting ran even when the sorted helper table came back unverified, so rows could be set from stale report figures and then reported as complete (`reportFormatting: null`, `formattedMonths: [9]`) | `_finishReportRows(months, sortedConsistent)` now gates every write path. When the helper table is not verified, **no visibility request is sent**, the months are recorded as unfinished with `blockedBy: 'sorted-table'`, and `formattedMonths` is empty. The retry `finishReportRows()` rebuilds the helper table first and only then formats, and the banner says which of the two it is doing | Unit: an add, a delete and a correction under a helper table that cannot be written all leave the report rows untouched and name the reason; the retry repairs then formats and is idempotent; a month sheet re-protected against row formatting is reported with the setting named |
| W2 | The months were recorded only after the operation returned, so closing or reloading between the transaction landing and the formatting returning lost the unfinished work, worst of all after a delete | The months are written **before the first workbook change** and cleared **only after verified completion**. A call that provably writes nothing (ignored submit, conflict) releases its reservation — and, after the W2b fix below, only the part of it that call added. A storage refusal is now surfaced to the member instead of swallowed, and the page keeps the record in memory for the rest of the visit | Browser: holding the visibility request open, reloading mid-operation, and finding the work still offered afterwards with zero writes at startup; plus an ordinary save showing no banner, and a stubbed storage failure producing a visible message while the save still succeeds |

Also in this pass, at the reviewer's direction:
- **Removed the advice to press a workbook button for recovery.** Run SubmitEntry and Run DeleteTransaction
  change transactions; recommending them as a formatting repair was wrong. Recovery is now: allow "Format rows"
  again in Excel and press Finish report formatting, or show the rows by hand. (`RefreshReports` was reviewed on
  2026-09-12 and is now the recommended repair — see the entry above.)
- **The local copy's protection does not prove the online copy's.** Said so in the mapping and made it an
  explicit live check in the runbook: the first save that formats rows either succeeds, or is refused and the
  app reads the protection back and names the setting.
- **`RefreshReports`** has since arrived and been reviewed (2026-09-12, the entry above). It confirms the row
  rule and becomes the documented repair; the guidance in this entry was written before that and is superseded
  there.

### Tests (2026-09-12, this folder)
- `npm test`: **141 passed, 0 failed** (was 137).
- `npx playwright test`: **183 passed, 0 failed** (was 177) = 61 per project × 3.
- `npm run verify:config`: 21 passed.
- One existing browser assertion was corrected rather than the code: the retry may now read the transaction
  table and repair the helper table, so the test asserts the sharper property, that it never **writes**
  `LOG_Table`.

Still not published, and no live write has been run.

---

## 2026-09-12 — Monthly report rows are now updated by the app. **For review; not published, no live write run**

A transaction saved from the website used to land in a month-sheet row that an earlier Office Script run had
hidden: right on the website, missing when the sheet was opened or printed in Excel. The website now finishes
each save by setting which rows that month shows. Full detail in `docs/workbook-mapping.md` §8.

### Feasibility was checked first, against current Microsoft documentation
| Need | Graph v1.0 | Source |
|---|---|---|
| Change row visibility | **Supported.** `PATCH …/range(address=…)` accepts `rowHidden` (Boolean); delegated `Files.ReadWrite` | [Update range](https://learn.microsoft.com/en-us/graph/api/range-update) |
| Read it back to verify | **Supported.** `rowHidden` reads true when all rows in the range are hidden, false when none is, null for a mix | same |
| Recalculate | **Supported.** `POST …/workbook/application/calculate` | already in the client |
| Unprotect **with a password** | **Not supported.** The action takes no request body at all | [unprotect](https://learn.microsoft.com/en-us/graph/api/worksheetprotection-unprotect) |
| Re-protect **with a password** | **Not supported.** `protect` takes `options` only | [protect](https://learn.microsoft.com/en-us/graph/api/worksheetprotection-protect) |

Office Scripts and Graph are different surfaces, and this is where they differ: `unprotect(password)` and
`protect(options, password)` exist in Office Scripts and have no Graph equivalent. A straight port of
`refreshMonthlyVisibility()` was therefore impossible.

**It turned out not to be needed.** A read-only inspection of the local workbook shows the twelve month sheets
are protected **without a password** and with **"Format rows" allowed**, which is exactly the permission
required to hide and show a row. ENTRY is the sheet with a password, and this app does not touch its
protection. So the chosen approach never unprotects anything and never needs a password at all.

### What it does
After the transaction and the sorted helper table are verified: recalculate once; read `A4:A33` of each
affected month; group the rows into runs sharing a visibility and PATCH `rowHidden` per run (normally two
requests for a month); read each run back and check it. Affected months are the month of an added or deleted
transaction, and **both** months when a correction moves one. It runs inside the existing operation queue and
single-writer tab guard, and never on page load.

It changes row visibility and nothing else: no protection call, no read or write of the cell holding the
password, no request to `LOG_Table` or `LOG_Sorted_Table`. That is what makes the retry below incapable of
duplicating a transaction, and it is asserted by tests rather than only claimed.

### When it does not finish
The transaction stays saved and is never rewritten. The screen shows **"Transaction saved. Excel report
formatting still needs updating."** (a delete says "Transaction deleted."), the outstanding months are kept in
this browser so a reload does not forget them, and a **Finish report formatting** button runs only the
visibility pass for those months. Saving is **not** paused, because unfinished formatting is cosmetic rather
than damage to a record; that is deliberately different from an incident. "Check workbook" and the check after
a reload remain read-only and are unrelated to this.

**Protection recovery, stated honestly.** If the month sheets are ever re-protected with "Format rows"
disallowed, the PATCH fails and **the app cannot repair it**: Graph cannot supply a password to unprotect, and
silently changing a protection option is not something this app should do. It does read the sheet's protection
back to explain the refusal by name. Recovery is manual in Excel: allow "Format rows" again, then press Finish
report formatting; or show the rows by hand. **Not** by pressing Run SubmitEntry or Run DeleteTransaction,
which add and delete transactions. (Since 2026-09-12, Run RefreshReports is also a reviewed repair: it changes
no transactions. See the entry above.)

### The alternative that was assessed
Leaving all thirty rows visible needs no permission and no writes, and was rejected: every month sheet would
print thirty mostly blank rows, changing the treasurer's paper report. Hiding blank rows is the workbook's
existing design.

### Files changed
| File | Change |
|---|---|
| `site/js/workbook/ledger-workbook.js` | `_syncMonthVisibility`, `syncMonthVisibility` (queued), `visibilityRuns`, `monthOfIso`; wired into the add, delete and both correction paths; self-test now lists report row visibility as not compared |
| `site/js/save/report-formatting.js` | new: the outstanding-months record, holding month numbers and a timestamp, no workbook content and no password |
| `site/js/ledger/app.js` | the green band, its persistence, and the explicit Finish button |
| `site/ledger.html`, `site/css/app.css` | the band |
| `site/js/workbook/mock-excel.js` | row visibility, modelled on the protection actually read from the workbook: `rowHidden` PATCH allowed on the protected month sheets, content writes still refused |
| `tests/unit/ledger-workbook.test.js` | six new tests; the "no formula-sheet writes" assertions now distinguish content writes from the permitted visibility writes |
| `tests/e2e/report-visibility.spec.js` | new, five tests |
| `docs/workbook-mapping.md` §8, `docs/member-guide.md`, `docs/checkpoint-2-runbook.md` Part 2b | the approach, what a member sees, and the live procedure |

### Tests (2026-09-12, this folder)
- `npm test`: **137 passed, 0 failed** (was 131). New: run grouping and `monthOfIso`; an add showing and hiding
  the right rows of its month and no other; a delete of a month's only transaction hiding all thirty; a
  correction across months fixing both; a failed format leaving the transaction saved with the exact message
  and a retry that finishes it; and idempotence plus the absence of any transaction-table, protection or ENTRY
  request during a sync.
- `npx playwright test`: **177 passed, 0 failed** (was 162) = 59 per project × 3. The partial-failure test
  asserts the message, that it survives a reload, that **zero writes happen when the page opens**, and that the
  retry leaves exactly one transaction.
- `npm run verify:config`: 21 passed.

**Mock tests are not evidence that Graph supports this.** The documentation above is the API evidence; the mock
was taught to allow a `rowHidden` PATCH on a protected sheet because the real workbook's protection allows it,
and that modelling is itself an assumption until the live test runs.

### Live test procedure
`docs/checkpoint-2-runbook.md` **Part 2b**: note the hidden rows on a month sheet in Excel, close it, add one
labelled transaction on the website, confirm no green band, reopen in Excel and confirm the row is visible and
prints, delete it on the website, and confirm the sheet returns to its earlier state. The Diagnostics
connection test already exercises the same write on December, so it is the first live proof if it passes.

### Not done, and outstanding
- **Not published and no live write run**, as instructed. The change is committed locally for review only.
- **`RefreshReports` was never supplied.** The reference used is `refreshMonthlyVisibility()`, which is the
  routine that actually does this work and appears in full in both `SubmitEntry` and `DeleteTransaction`. If
  `RefreshReports` does anything beyond it, this may need revisiting.
- Months left stale by saves made **before** this change are not repaired automatically; a save touching that
  month fixes it, or the workbook's own button does.

---

## 2026-09-11 (live, read-only) — First connection to Microsoft 365 succeeded. **Partial Checkpoint 2 evidence, not approval**

Reported by Cody after running the Diagnostics page against the pilot copy `TPAHA_2026 (1).xlsx`. The account
used was an assigned board account (the Vice-President); the person's name is kept out of this file because it
is published, and is in the local notes. **No write of any kind was run, and the connection write test was not
started.**

### What happened
| | |
|---|---|
| Sign-in | An assigned account completed the Microsoft redirect and returned to the site |
| Workbook | Identified and confirmed on the Diagnostics page |
| Read-only checks | All six passed |
| Transaction table | Read back with the row count the treasurer expected; the exact figure is in the local notes |
| Sorted helper table | Matches the transaction table |
| Requests | Every logged content request was a GET returning 200 |
| Write test | Not run |

The one non-GET such a run makes is the workbook session this app opens before its first read
(`POST …/workbook/createSession`), which changes no cell. Everything else was a read.

### What this establishes, for the first time against the real service
- The **hosted sign-in works end to end**: the MSAL redirect round-trip on the GitHub Pages address, the
  registration, admin consent, the single-tenant restriction and assignment all function together, and a token
  was issued and accepted by Microsoft Graph. None of that could be shown by any amount of local testing.
- **Graph can open this workbook and this app can read it.** The adapter's read path works against real data:
  the column headers match, the table body parses, the year cell and the prior-year balance cell are the
  expected shapes, and the tables are found by name.
- **The sorted helper table in the online copy already satisfies the date-then-transaction-number order** that
  the workbook's own Office Scripts require and that the V1 fix aligned this app to. Had the two disagreed,
  that check would have failed.
- No throttling, no session error, no missing sheet or table: every request returned 200.

### What it does not establish, and must not be read as
- **Nothing about writing.** No row was added, corrected or deleted; no number format was set; no helper-table
  rebuild ran. Everything in `docs/workbook-mapping.md` §3 about `rows/add`, the checked delete, the
  `itemAt(index=n)` format PATCH and the rebuild loop remains unverified against the real service.
- **Nothing about the nine-digit timestamp.** Whether Excel stores it unchanged, and whether an Office Script's
  whole-table rewrite preserves it, is still open. That is the separately authorised check in the runbook.
- **Nothing about the Office Scripts**, which were not run.
- **Nothing about refusal.** Only an assigned account signed in. That an **unassigned** account is turned away
  with AADSTS50105 has not been observed and should be.
- **Nothing about phones.** The browser suite uses two engines, never a device.
- **Nothing about §5 of the mapping.** Those findings (the Annual sheet fixes, the duplicate transaction
  numbers) were made against the local file. The six read-only checks do not examine the Annual sheet's
  formulas or look for duplicate numbers, and the online copy is not the same snapshot as the local one. Worth
  confirming before the pilot, because duplicate numbers would make the workbook's own scripts refuse to run.

### Where this leaves Checkpoint 2
Part of the evidence is in hand. Checkpoint 2 is not met and no approval follows from this. Still required:
`RefreshReports` supplied and reviewed; the hidden-month-row decision; Cody's authorisation; then the
connection write test and the separate script-identity check, with their logs and before-and-after identity
evidence.

---

## 2026-09-11 (published) — The configured site is live; sign-in testing is next

Pushed `b18dd2e..7625121` on Cody's instruction: `db30a38` (the registration configuration) and `7625121`
(a correction to how the test-mode host rule is described). GitHub Pages redeployed from the new tip.

### Deployment verified, all of it read-only
| Check | Result |
|---|---|
| `/`, `/ledger.html`, `/diagnostics.html`, `/storage.html` | HTTP 200 |
| Live `js/config.js` against the committed file at `7625121` | byte-identical |
| `appVersion` served | `2026.09.11`, so phones fetch the new files |
| Corrected host-rule wording present in the deployed file | yes |
| Configuration checks (`npm run verify:config`) | 21 passed, 0 failed |
| `npm test` / `npx playwright test` before publishing | 131 and 162 passed, 0 failed |

Nothing here contacted Microsoft 365, no workbook was opened, and the Diagnostics write test was not run.

### What publishing changed
The live site is no longer a test-mode page. It now offers Microsoft sign-in to the five assigned accounts and,
once a member identifies and confirms a workbook, can read and write it. Reading and writing are separate
things and only the latter needs care: signing in, picking a workbook, browsing the ledger and running the six
read-only Diagnostics checks issue GET requests only. A write happens when someone presses Save, Correct,
Delete or the Diagnostics connection test.

### The wording correction
The host rule is a denylist of local hostnames, not an allowlist of the published one: `localhost`,
`127.0.0.1`, `::1` and `file://` pages get blank ids, and **every other hostname gets them**, including a LAN
address or a fork published elsewhere. That is deliberate; the client id is an identifier, not a credential.
What confines a real sign-in to this site is the single SPA redirect URI in the registration, which Microsoft
checks before issuing a token, together with assignment being required. `tools/verify-config.mjs` now asserts
that shape with eight hostname cases rather than only describing it.

### Next, and the limits on it
Sign-in testing, which Codex is walking Cody through. Worth confirming: an assigned account completes the
Microsoft redirect and returns to the site; an unassigned account is refused by Microsoft (AADSTS50105, "not
assigned to a role for the application"); the workbook picker lists the intended file and the six read-only
checks pass against it. **The Diagnostics connection test stays untouched** until `RefreshReports` has been
supplied and reviewed, the hidden-month-row decision is made, and Cody authorises the run. The separate
script-identity check in the runbook is likewise unauthorised.

---

## 2026-09-11 (later) — Entra app registration complete; site configured, not yet published

Cody completed the app registration and supplied its identifiers. `site/js/config.js` now carries them. No
workbook write has run, nothing has connected to Microsoft 365 from here, and the change is **committed locally
but not pushed**, so the live page is still the test-mode one.

### The registration, as reported
| Setting | Value |
|---|---|
| Application (client) ID | `6c5bf2f8-a50c-4908-ba93-535816364785` |
| Directory (tenant) ID | `ef5ad4e3-2b18-4542-8fd2-ed7212cbcc93` |
| Platform and redirect | Single-page application, `https://tpahavp.github.io/tpaha-books/` |
| Delegated permissions | `User.Read` and `Files.ReadWrite`, admin consent granted |
| Assignment required | Yes, with all five intended accounts assigned |
| Implicit grant / public client flows | unchecked / disabled |
| Client secret | none, and there must never be one |

Both ids are configuration rather than secrets, which is why they live in a file the site publishes: they
identify the app and the organisation, while who may sign in and what they may open stays with Microsoft 365.

### A problem this raised, and how it was solved
Test mode was triggered by an empty `clientId`. Filling the id in would have taken the whole browser suite
offline with the real registration and made local development impossible without editing the file back and
forth. `site/js/config.js` now **blanks the ids for local hostnames** (`localhost`, `127.0.0.1`, `::1`, and
a `file://` page) and supplies them for **every other hostname**. It is a denylist of local addresses, not an
allowlist of the published one: a copy served from a LAN address or a fork published elsewhere would also carry
the client id. That is deliberate, and the id is not what protects anything. The registration accepts exactly
one SPA redirect URI, so Microsoft rejects a sign-in begun anywhere else before issuing a token, and assignment
is required. The rule depends on the address the page was served from rather than a flag or query parameter, so
nothing a member can click, type or paste changes it, and the tests keep running offline against the sample.

### Verification, none of it touching Microsoft
`tools/verify-config.mjs` (new, `npm run verify:config`): **17 checks, all passing**. It checks that both ids
are well-formed GUIDs matching the registration and differ from each other; that a page served from
`localhost` or `127.0.0.1` gets no ids while other hostnames, including the published site, a future custom
domain and a LAN address, get them; that the
redirect URI the app derives from each of the four pages equals the registered SPA redirect exactly; that the
scopes requested are exactly the two consented; that the authority is the single tenant rather than
`/common`; that sign-in is redirect-based so no implicit grant is needed; that no client secret, certificate
or private key exists anywhere in the published site; that no workbook is pinned, so a member must still
identify and confirm one; and that the writer model is still single-writer.

What it cannot do, stated plainly: it cannot prove the registration in Entra still matches these values, that
consent is still granted, or that an assigned account can actually sign in. Those are Checkpoint 2 checks
against the real service.

- `npm test`: **131 passed, 0 failed**.
- `npx playwright test`: **162 passed, 0 failed**, confirming local test mode survived the change.
- `appVersion` bumped to `2026.09.11` so phones fetch the new files.

### The decision this leaves with Cody
Publishing this commit is the moment the live site stops being a test-mode page and becomes a real sign-in page
that can read and write the chosen workbook. That is the intended next stage, but it should be a deliberate
push rather than one that rides along with something else. After it, confirm that an assigned account reaches
the site and an unassigned one is refused with Microsoft's "not assigned" message.

### Still outstanding, unchanged
`RefreshReports` has not been supplied, and the hidden-month-row decision is still open. Both are needed
before the pilot, neither blocks publishing the configuration.

---

## 2026-09-11 — Hosting live; Entra app registration next

No code changed for this entry. Nothing has connected to Microsoft 365; no live write has run.

### Hosting (setup guide part A): done
| | |
|---|---|
| Repository | `https://github.com/TPAHAVP/tpaha-books` (public, association-owned) |
| Published commit | `b18dd2e`, one commit, 65 files. The ref GitHub holds matches the commit that passed the pre-upload scans, checked after the push |
| Site address | `https://tpahavp.github.io/tpaha-books/` — this is the **redirect URI** for the app registration |
| State when opened | **Test mode**, as intended: `clientId` in `site/js/config.js` is still empty, so there is no sign-in and no Microsoft 365 access from the published site |
| Deployment | `.github/workflows/pages.yml` publishes the `site/` folder on every push to `main`; the documentation and tests are not served |

### What preparing the upload turned up
Scanning before the first push found three kinds of real information that would otherwise have been published,
all now held back by `.gitignore` and still present on the treasurer's machine:

1. **Real figures inside four test files.** `tests/unit/ledger-model.test.js` (9 tests),
   `tests/unit/storage-model.test.js` (14), `tests/unit/xlsx-export.test.js` (3) and
   `tests/e2e/storage.spec.js` (9) assert values read from the real workbooks, among them the closing balance,
   the year net, a utility bill, the storage revenue and the space count. They need the git-ignored fixtures and
   skip without them, so they are useless to a clone. A clone therefore runs a smaller suite than the figures
   quoted in this report, and the difference is exactly those four files. `README.md` says so.
2. **Real figures in the documents.** `docs/workbook-mapping.md` named the year net, real cheque numbers and the
   transaction-number range; those are now described without the values. `docs/superpowers/`, the superseded
   design material from the first approach, quoted several more and is excluded whole.
3. **The worksheet-protection password's location.** The mapping and this report named the exact cell. They now
   say only that it sits in a cell on the ENTRY sheet, with the location kept in the local note beside the
   script source. One bare reference survived the first pass, was caught by review on commit `8944dfb`, and was
   removed by amending before anything was pushed; `8944dfb` never reached GitHub.

Also changed: a placeholder in `site/js/ui.js` that named the association's SharePoint host is now generic. The
storage rates and space count shipped in `site/js/storage/model.js` were left alone: they are application
defaults that any visitor to the storage page already sees, not workbook records.

### Where Checkpoint 2 preparation stands
| Prerequisite | State |
|---|---|
| A. Checkpoint 1 approved | done, 2026-09-08 |
| B. The three Office Scripts read | done, 2026-09-10. **`RefreshReports` still outstanding** |
| C. Writer arrangement and numbering policy | accepted, 2026-09-08 |
| C2. Hidden month rows | **outstanding**: accept a documented manual refresh, or ask for an implementation change |
| D. Hosting | **done, 2026-09-11** (above) |
| D. Entra app registration, assignment, `config.js` | **next** (setup guide B–D). Codex is walking Cody through it |
| E. A quiet test copy on the day | on the day |

The live Diagnostics run and the separately authorised script-identity check follow once B, C2 and D are
complete and Cody authorises them. Neither is authorised by any review so far.

---

## 2026-09-10 (sign-off) — **V1 closed by independent review; Checkpoint 2 preparation continues**

Sign-off read: `Desktop\Codex projets\excel to HTML idea\review-script-compatibility-signoff\SCRIPT_COMPATIBILITY_SIGNOFF.md`.

> "The V1 sorting remediation is approved. No further blocking findings were identified in this follow-up.
> Checkpoint 1 remains approved; hosting and Microsoft account preparation can proceed."

No redesign of the sorting fix is requested.

### Verified independently by the reviewer
- One ordering rule, `expectedSortedRows()`, serves both the adapter's rebuild and `sortedTableMatches()` used
  by Diagnostics. Correct same-day transactions pass after a correction moves their table positions; date-only
  ordering is still rejected.
- Their original V1 reproduction, rerun unchanged, passes.
- `npm test` **131 passed, 0 failed, 0 skipped**; project browser suite **162 passed, 0 failed**; with their own
  reproduction, **163 passed**. Their server ran on port 8801 against this folder. Two browser engines, no
  physical phone.
- **They read the saved script source** in `data/office-scripts/` and compared its helpers with the unit-test
  ports. Their conclusion matches what is recorded in `docs/workbook-mapping.md` §7: the ports reflect the
  originals for valid app-generated transactions; there is no timestamp-specific parser or matching rule in the
  three scripts; whole-row comparisons do include the timestamp value, and the writing scripts rewrite it along
  with the other cells. Retaining that source is what made this check possible.

### Explicitly still not established
Reading the source does not prove the live Office Scripts runtime preserves the exact nine-digit timestamp
through a whole-table rewrite. The reviewer confirmed the runbook now separates that check from the ordinary
Diagnostics self-test and marks it as awaiting authorisation, and that the hidden-row guidance no longer uses
Submit or Delete to refresh a report. This sign-off closes a code finding. It does not authorise live writes or
establish production readiness.

### Remaining steps, in the reviewer's order
1. Supply and review **RefreshReports** before the pilot.
2. Decide how to handle hidden month rows. A manual refresh needs the reviewed refresh script confirmed to
   reveal populated rows, and the resulting Excel view or printout checked. Automatic visibility updates would
   be a separate change.
3. Complete hosting and Microsoft sign-in and account setup (`docs/setup-guide.md` A–D).
4. Once those are done **and Cody authorises the live checks**, run the test-copy Diagnostics sequence and the
   separately documented script-identity check, and return the logs and the exact before and after identity
   evidence for Checkpoint 2 review.

---

## 2026-09-10 (later) — Script-compatibility follow-up review: V1 fixed (closed by the sign-off above)

Review read: `Desktop\Codex projets\excel to HTML idea\review-script-compatibility\SCRIPT_COMPATIBILITY_REVIEW.md`.
The adapter's new date-then-number ordering was accepted; one missed consumer of the rule was found.

| # | Finding | What changed | Regression evidence |
|---|---|---|---|
| V1 (P2) | The Diagnostics read-only check still sorted by date alone, keeping table position for ties, so after a correction it called the adapter's correct output inconsistent and disabled the connection test on valid data. The reverse was also true: the old order would have passed the check while the workbook's own scripts rejected it. | The ordering rule is now one exported function, `expectedSortedRows()`, with `sortedTableMatches()` beside it for the read-only comparison, both in `site/js/workbook/ledger-workbook.js`. The rebuild and the Diagnostics check call it, so they cannot drift apart again. Blank rows are handled in that one place. The check stays read-only and was not weakened; its label now names the required order. | Unit: "V1: one shared rule orders LOG_Sorted, so the rebuild and the read-only Diagnostics check cannot disagree" and "a row with no transaction number sorts last within its date and keeps table order". Browser (`tests/e2e/diagnostics.spec.js`): "Diagnostics accepts the date-then-number order the adapter writes, after a same-day correction" (all six checks pass, the test button enables) and "Diagnostics still rejects a helper table ordered by date alone, with the numbers reversed on one date". The reviewer's own reproduction, unchanged, now passes. |

### Evidence corrections the review asked for
- **The Diagnostics connection test never runs an Office Script.** Claims that it would prove a script's
  whole-table rewrite preserves a nine-digit timestamp have been removed from the mapping, the status entry
  above and the runbook. The unit test that exercises such a rewrite is renamed to say it covers the mock only.
- **A changed identity is no longer described as failing loudly.** The app notices one only where it looks for
  its own row by operation id; a rewrite between operations would look like a changed or missing record.
- **A separate, separately authorised live identity check** is now written into
  `docs/checkpoint-2-runbook.md`: save a labelled entry, record its exact Timestamp text, run the reviewed
  script action with nothing in flight, re-read and compare. It is marked as authorised by nobody yet.
- **Hidden month rows**: the guidance no longer says to press a script button to refresh visibility. It calls
  for a documented manual refresh with a reviewed `RefreshReports`, verified in Excel first, or a separate
  implementation change.
- **The scripts' source is retained for review** in the git-ignored `data/office-scripts/` folder, so the
  ports in the unit tests can be checked against the originals. It is kept out of source control because it
  documents the workbook's internal structure, including where the sheet-protection password sits.

### Tests run (2026-09-10, this folder)
- `npm test`: **131 passed, 0 failed, 0 skipped**.
- `npx playwright test`: **162 passed, 0 failed** = 54 per project × 3 projects.
- The reviewer's V1 reproduction, unchanged: **1 passed**.

### Still open for Cody
`RefreshReports` has still not been supplied, and the hidden-month-row decision is still open. Hosting and
account preparation can continue meanwhile.

---

## 2026-09-10 — The workbook's Office Scripts reviewed; one compatibility defect found and fixed

Cody supplied the source of SubmitEntry, FindTransactions and DeleteTransaction (Checkpoint 2 prerequisite B).
Findings and the resulting change are recorded in `docs/workbook-mapping.md` §7.

| Question | Answer | Consequence |
|---|---|---|
| Do they read, compare or parse the Timestamp column? | **No.** SubmitEntry writes its own `new Date().toISOString()`; FindTransactions matches on Date only; DeleteTransaction matches on TransactionID plus the Date, Description and Amount shown on the form. | The nine-fraction-digit operation id is safe. The concern carried since review R4 is closed. |
| Do they rewrite the column? | **Yes**, as part of a whole-table rewrite: `writeRows()` sets values over the entire body of both tables on every save and delete, with the values it just read. | Text round-trips, which is how today's three-digit timestamps survive. Whether the nine-digit form does the same is **not established for the real service**, and the ordinary connection test cannot establish it because it never runs an Office Script. A separately authorised live identity check is set out in the runbook. A changed identity should not be assumed to announce itself. |
| Do they impose rules this app must satisfy? | **Yes, two.** `validateLog` rejects the whole workbook unless every row has a unique positive integer id, an integer date serial inside the workbook year, a valid type and category pairing, and an amount with at most two decimals. `verifySorted` rejects it unless LOG_Sorted equals LOG sorted by **date, then transaction number**, compared as whole rows. | Our entry validation already met `validateLog`. **`verifySorted` did not hold**: see the defect below. |
| Does DeleteTransaction delete positionally? | It rewrites both tables wholesale rather than deleting one row by index. | Worse for concurrency, not better: a script can move every row at once. It reinforces the one-writer rule rather than changing it. |

### Defect found and fixed (2026-09-10)
`sortedTarget()` in `site/js/workbook/ledger-workbook.js` ordered LOG_Sorted by date and then **table
position**; the scripts require date and then **transaction number**. The two agree while rows are only
appended, and diverge as soon as this app corrects a row that shares a date with a higher-numbered one,
because a correction moves that row to the end of LOG. The scripts would then have thrown "LOG and LOG_Sorted
disagree. Run RefreshReports before entering or deleting transactions" and refused to run, until someone ran
RefreshReports. Fixed by matching their tie-break. A row with no transaction number sorts last and keeps table
order.

### Regression evidence (3 new unit tests, ported from the scripts' own helpers)
- "Office Scripts: LOG_Sorted keeps the order the scripts require (date, then transaction number), including
  after a correction moves a row to the end of LOG" — builds exactly that divergence and asserts the scripts'
  `verifySorted` accepts the result.
- "Office Scripts: every row this app writes passes the scripts' own validateLog, after an add, a correction
  and a delete" — a faithful port of `validateLog`/`validateEntry` run over the workbook the app leaves behind.
- "Office Scripts: a nine-digit operation id survives the whole-table rewrite the scripts perform, and the app
  still finds its row".

### Tests run (2026-09-10, this folder)
- `npm test`: **129 passed, 0 failed, 0 skipped** (126 before, plus the three above).
- `npx playwright test`: **156 passed, 0 failed** across the three configurations. The sort change altered no
  screen behaviour.

### Two items now open for Cody
1. **`RefreshReports`**: a fourth script, named in the other two scripts' error messages, whose source has not
   been read. It presumably rebuilds LOG_Sorted and month-row visibility; if it writes LOG_Sorted it needs the
   same review. Send it before the pilot.
2. **Hidden month rows**: the scripts hide empty month-sheet rows and this app does not unhide them, so a
   transaction added on the web can stay hidden when that month sheet is opened or printed in Excel until a
   that month sheet is opened or printed in Excel. The web pages themselves are unaffected. Either accept it for
   the pilot with a documented manual refresh using a reviewed `RefreshReports`, verified in Excel first, or ask
   for the app to unhide rows, which would be a separate implementation change writing to protected sheets using
   the password stored in a cell on the ENTRY sheet (named in the local notes). Not decided here. Pressing Submit or Delete to refresh visibility is not
   an acceptable instruction.

### Security note passed on, not caused by this app
The worksheet-protection password is stored in plain text in a cell on the ENTRY sheet (named in the local notes) and read from there by
`refreshMonthlyVisibility()`. Anyone who can open the workbook can read it. This app uses `ENTRY!B23` only and
never touches that cell.

---

## 2026-09-08 (sign-off) — **Checkpoint 1 approved; Microsoft 365 verification pending**

Codex sign-off read: `Desktop\Codex projets\excel to HTML idea\review-checkpoint-1-signoff\CHECKPOINT_1_SIGNOFF.md`.

> "Checkpoint 1 is approved for the locally tested Treasurer Ledger implementation and its proposed
> single-writer test-copy pilot. No further blocking findings were identified in this remediation review.
> U1 is closed; the previously accepted T1–T4 fixes remain accepted."

**What the approval does not cover**, in the reviewer's words and mine: it does not establish that the real
Microsoft 365 connection works, does not approve production use, and does not approve simultaneous board
editing. Graph sessions, permissions, real workbook formats and formulas, hosted sign-in, and behaviour on
actual phones remain unverified. The browser guard and the operation queue coordinate one page in one
browser; they do not coordinate an external Excel session, the old script buttons, another browser or
another device, so the single-writer rule stands for the pilot. Winter Storage remains browser-only and
awaits its own workbook connection and review.

### Independently reproduced by the reviewer (their run, their server on port 8799)
| Check | Result |
|---|---|
| `npm test` in this folder | 126 passed, 0 failed, 0 skipped |
| Project browser suite | 156 passed, 0 failed |
| The previously failing U1 reproduction, unchanged | 1 passed |
| Combined browser run | 157 passed, 0 failed |
| Overlapping-delete verification | passed |

Browser configurations: desktop Chromium, iPhone-sized Chromium, iPhone-sized WebKit. Two engines, no
physical device. Nothing connected to Microsoft 365; no live write was performed by the reviewer or here.

### Review history closed by this sign-off
| Pass | Findings | Outcome |
|---|---|---|
| Checkpoint 1 review (2026-09-06/07) | R1–R6 | fixed, accepted |
| Re-review (2026-09-07) | S1–S4 | fixed, accepted |
| Round 3 (2026-09-07) | T1–T4 | fixed, accepted |
| Round 4 (2026-09-08) | U1 | fixed, closed |
| Sign-off (2026-09-08) | none | **Checkpoint 1 approved** |

### Checkpoint 2 preparation: state of play
`docs/checkpoint-2-runbook.md` (new) is the single page to follow on the day: five prerequisites, the run
itself, what to send back, the stop conditions, and what the test does and does not establish.

| Prerequisite | Owner | State |
|---|---|---|
| A. Checkpoint 1 approved | — | **done** (this entry) |
| B. The three Office Scripts read for Timestamp handling | Cody (owner account) | **done**, 2026-09-10 (see the entry above): none of them reads the Timestamp column; one compatibility defect found and fixed. `RefreshReports` still to send |
| C. Writer arrangement and numbering policy accepted | Cody / the board | **done**, 2026-09-08 (below) |
| D. Hosting, app registration, assignment, `config.js` filled in | Cody | **outstanding**; `docs/setup-guide.md` A–D. No password, token or client secret is ever supplied to the developer |
| E. The test copy quiet during the run | Cody | on the day |

Nothing in the code changes for Checkpoint 2: the accepted fixes are not being redesigned. `config.js` still
carries an empty `clientId`, which is what keeps every page in test mode until step D fills it in.

### Decisions taken by Cody, 2026-09-08
- **Writer arrangement: one writer at a time**, as built and reviewed. One designated tester works at a time
  on the test copy; nobody edits the workbook in Excel or runs its script buttons meanwhile. The site enforces
  what a browser can see (banner, a confirmation before the first Correct or Delete, refusal to write from a
  second tab, one complete operation at a time) and nothing beyond that. No code change; `config.pilot.writerModel`
  stays `'single'`. Simultaneous board editing remains unsupported and unapproved.
- **Transaction numbering: keep highest existing number plus one**, with its consequence accepted: deleting the
  highest-numbered transaction frees that number for a later entry. Numbers stay unique at any moment; when two
  members add at the same instant the later row is renumbered to the next free number and the member is told.
  Existing entries are never renumbered by a correction. No code change; the member guide already states this.

---

## 2026-09-08 (later) — Checkpoint 1 round-4 remediation (U1), returned for sign-off

Review read: `Desktop\Codex projets\excel to HTML idea\review-checkpoint-1-round4\CHECKPOINT_1_REVIEW_ROUND4.md`
("T1–T4 are resolved in the reviewed paths. One remaining correction-recovery fix is required before closing
Checkpoint 1."). Implemented and tested locally against the in-memory Graph replica. **Nothing has been
verified against Microsoft 365**; no live write, no production connection.

| # | Finding | What changed | Class | Regression evidence (all passing) |
|---|---|---|---|---|
| U1 (P2) | A correction that saved while the member kept typing stored the newer draft together with the **old** row reference, so after a reload the newer text hit a false "Someone else changed this record"; the offered Reload record action then discarded that text | `afterEditSave()` persists the combined record **immediately after** `state.editingRef` becomes the saved row, in the unsaved branch as well as the saved one, so the stored reference always names the version that actually reached the workbook. The same path serves a pending attempt found to have landed (`inspect` → `resume`), because both end in `afterEditSave()`. The fingerprint conflict check is untouched: an external change is still reported. `reloadEditedRecord()` no longer replaces typed text silently: when the reloaded row differs from what is on screen it shows both versions and offers **Keep what I typed** (the reference moves to the row just read and shown, the text stays unsaved), **Use the workbook version**, or **Cancel**. | Prevented | Browser (`tests/e2e/round4.spec.js`, 4 tests): "text typed while a correction saves is stored with the SAVED row as its reference; after a reload it saves without a false conflict" (persisted reference `First correction saved`, zero writes on startup, one copy of #2 carrying the newer text, every unrelated row byte-identical); "an unconfirmed correction that landed is finished after a reload, and newer text then saves against the finished row"; "a real conflict is still reported, and Reload record asks before replacing typed text; keeping it saves onto the row just read"; "Reload record can also replace the typed text with the workbook version, on purpose". The reviewer's own `reviewer.spec.mjs`, which failed before, now passes and logs the corrected reference. |

### Evidence label corrected (asked for by the review)
The round-3 test titled "a pending (unconfirmed) correction is still persisted after an unrelated entry Save
succeeds" never performed an entry Save. It is renamed to what it covers: "a pending (unconfirmed) correction
survives a reload and a cancelled discard, and stays persisted". The claimed sequence is no longer reachable
through the UI, and the test file says so: since T1 the page refuses a second write while one is running, and a
pending correction keeps its modal dialog open, so an entry Save cannot overlap it. The composition of a saved
entry with a pending correction is proven by the unit test "T4: combineDrafts keeps each form part on its own".

### Tests run (2026-09-08, this folder)
- `npm test`: **126 passed, 0 failed, 0 skipped**.
- `npx playwright test`: **156 passed, 0 failed** = 52 per project × 3 projects (desktop Chromium,
  iPhone-13-size Chromium, iPhone-13-size WebKit): ledger 13, review regressions 12, round 3 7, round 4 4,
  diagnostics 4, launcher 3, storage 9. Two engines emulated; no physical device.
- Reviewer artifacts rerun: `reviewer.spec.mjs` (U1) **passes**, printing `Persisted reference after earlier
  save: First correction saved`; `verify-overlapping-deletes.mjs` (T1) passes, every other row preserved exactly.

### Unchanged
Still mocked / unverified live (Graph behaviour, formats, sessions, throttling, the online copy's layout, the
old Office Scripts' handling of the Timestamp column, a real iPhone, the hosted MSAL redirect); no live writes.
The writer-model and transaction-number decisions remain Cody's: the recommendation is turn-taking for the
pilot, with clear notification when a saved transaction receives the next available number, and the explicit
note that max-plus-one can reuse the former highest number after a deletion.

### Requires Cody
1. **Sign-off review** of: `site/js/ledger/app.js` (`afterEditSave`, `reloadEditedRecord`, `reloadChoiceDialog`),
   `tests/e2e/round4.spec.js`, the renamed round-3 test, `docs/member-guide.md`.
2. Decisions: writer model; transaction-number policy.
3. Before any live write: the three Office Scripts' source for Timestamp compatibility; a quiet test copy;
   Microsoft account and hosting setup (setup guide A–D). Account and hosting preparation may proceed now; the
   live write test is not authorised by any review so far.

---

## 2026-09-08 — Checkpoint 1 round-3 remediation (T1–T4), returned for re-review

Review read: `Desktop\Codex projets\excel to HTML idea\review-checkpoint-1-round3\CHECKPOINT_1_REVIEW_ROUND3.md`
(verdict "changes required before approving the live-write pilot"). Implemented and tested locally against
the in-memory Graph replica. **Nothing has been verified against Microsoft 365**; no live write, no production
connection.

| # | Finding | What changed | Class | Regression evidence (all passing) |
|---|---|---|---|---|
| T1 (P1) | One person in one tab could overlap two deletes; both planned their positional DELETE from the same stale layout and an unrelated row (#4) was lost | The adapter now runs **one complete operation at a time**: every public operation (add, correction, delete, inspect, finish, balance write, repairs, sorted rebuild, load, resolution check) is queued behind the previous one and starts only after its verification and helper-table work; internal helpers never re-enter the queue, so nothing deadlocks; a queued operation re-checks the pause when it starts. The page refuses a second Save, Delete or repair while one runs ("Another change is still being saved…"), greys the Delete buttons, and lets typing continue. The connection test refuses to overlap itself, and Diagnostics refuses to change the workbook while it runs. This serialises one page only; the one-writer rule for other tabs, devices, Excel and scripts stands. | Prevented (within one page) | Unit: "T1: two deletes confirmed from one page at once are serialised … only the two intended rows disappear"; "T1: mixed overlapping operations (add, correction, delete) … every unrelated row survives"; "T1: a queued operation re-checks the pause … the connection test refuses to overlap itself". Browser (`tests/e2e/round3.spec.js`): "T1: a second Delete confirmed while the first is still reading is refused; … #3, #4 survive"; "T1: Save is refused while another change is running, and typing stays possible". Reviewer reproduction rerun (`Promise.allSettled` of two deletes): both fulfilled, remaining `[1,4,5,…,13]`, #4 present, no incident. |
| T2 (P2) | Diagnostics ignored the tab guard, so the ledger and Diagnostics could both write in one browser | Diagnostics joins the same tab guard (same channel). The first TPAHA Books tab opened in a browser, whichever page, owns writing; in any other tab Diagnostics keeps its read-only checks but the write test is disabled **and** its handler refuses even a forced click; the ledger already refused. Browsers without BroadcastChannel are told the page cannot detect other tabs. | Prevented (same browser) | Browser: "T2: with the ledger open first, Diagnostics in a second tab keeps its read-only checks but refuses the write test, even when forced" (zero writes); "T2: with Diagnostics open first, the ledger in a second tab is the one that refuses to write". |
| T3 (P2) | The ledger's Save never received the nine-digit id generator; it used millisecond timestamps, so two identical saves at one clock value merged | The entry controller is given `makeMarker` explicitly, and the controller's default is now the same safe generator (`defaultMarker`), so no caller can fall back to a bare timestamp. One pending attempt keeps its id; each new operation gets a fresh one; payload collision checks unchanged. | Prevented | Unit: "T3: the default operation id is ISO 8601 with nine fraction digits and differs even when the clock does not move" (frozen `Date`); "T3: one pending attempt keeps its marker across retries, and each distinct new operation gets a fresh one". Browser: "T3: with the wall clock frozen, two distinct identical Saves create two rows with distinct nine-digit operation ids, and a retried lost attempt still makes one row" (asserts the id the real Save button wrote: `2026-09-07T12:00:00.000dddddd Z`). |
| T4 (P2) | A successful entry Save cleared the whole draft record, erasing a separate correction's persisted draft | Persistence is a single combined record computed by `combineDrafts()` in `drafts.js`; every success, reset and cancel path re-persists through it and no form ever clears the record directly. A saved entry drops only its own part; a typed or pending correction stays. | Prevented | Unit: "T4: combineDrafts keeps each form part on its own: a saved entry does not erase a typed or pending correction". Browser: "T4: a correction typed while an entry is saving stays persisted after the entry succeeds, and reopens with its exact text after a reload"; "T4: a pending (unconfirmed) correction is still persisted after an unrelated entry Save succeeds". |

### Tests run (2026-09-08, this folder)
- `npm test`: **126 passed, 0 failed** (ledger adapter 35, save controller 16, drafts 6, incidents + tab guard 3, Graph client 12, mock 10, concurrency 5, locator 4, store 3, storage model 14, ledger model 9, UI 6, export 3).
- `npx playwright test`: **144 passed, 0 failed** = 48 per project × 3 projects (desktop Chromium, iPhone-13-size Chromium, iPhone-13-size WebKit): ledger 13, review regressions 12, round 3 7, diagnostics 4, launcher 3, storage 9. Two engines emulated; no physical device.
- Round-3 reviewer reproduction (`reproductions.mjs` shape) rerun: safe outcome (above). The four reviewer browser demonstrations assert the defects; their safe counterparts are the seven tests in `tests/e2e/round3.spec.js`.

### Unchanged from the previous entry
Still mocked / unverified live (Graph behaviour, formats, sessions, throttling, the online copy's layout, the
old Office Scripts' handling of the Timestamp column, a real iPhone, the hosted MSAL redirect); no live writes;
the writer-model and transaction-number decisions remain Cody's (recommendation: turn-taking for the pilot,
clear notification when a saved transaction receives the next available number, and the explicit note that
max-plus-one can reuse the former highest number after a deletion).

### Requires Cody
1. **Re-review** of: `site/js/workbook/ledger-workbook.js` (operation queue), `site/js/save/save-controller.js`
   (`defaultMarker`), `site/js/save/drafts.js` (`combineDrafts`), `site/js/ledger/app.js`, `site/js/diagnostics.js`,
   `tests/unit/{ledger-workbook,save-controller,drafts}.test.js`, `tests/e2e/round3.spec.js`, `docs/workbook-mapping.md`
   §3/§3b, `docs/member-guide.md`, `docs/setup-guide.md` E, `README.md`.
2. Decisions: writer model; transaction-number policy (both unchanged, see the previous entry).
3. Before any live write: the three Office Scripts' source for Timestamp compatibility; a quiet test copy;
   Microsoft account and hosting setup (setup guide A–D).

---

## 2026-09-07 (later) — Checkpoint 1 re-review remediation (S1–S4), returned for re-review

Review read: `Desktop\Codex projets\excel to HTML idea\review-checkpoint-1-rereview\CHECKPOINT_1_REREVIEW.md`
(verdict "changes required"). Everything below was implemented and tested locally against the in-memory
Graph replica. **Nothing has been verified against Microsoft 365.** The live write self-test was not run and
no production records were connected.

### Vocabulary used from here on (as the re-review asked)
- **Prevented**: the app saw the danger before sending a request and sent nothing harmful.
- **Detected**: the app found afterwards, by row identity, that a request hit the wrong row; the damage is
  named and saving is paused. **Detection is not preservation**: the row is gone until a member restores it.
- **Resolved**: the app re-read the table and every documented condition for that incident held (verified),
  or nothing could be verified and a member acknowledged after checking Excel's version history.
- **Unverified live behaviour**: everything about the real Microsoft 365 service.

### Finding-by-finding

| # | Finding (P1) | What changed | Class | Regression evidence (all passing) |
|---|---|---|---|---|
| S1 | Wrong-row deletion remained; the batch GET was not evaluated before the DELETE; reporting is not preserving | The read+delete `$batch` is withdrawn. Every positional request (DELETE of a row, PATCH that renumbers a row) is now preceded by a **fresh read that is evaluated first**: the row at the index must be exactly the intended one (identity and content); if it moved, the app re-reads and re-aims (three rounds), if it is gone it stops. The residual race (a shift in the instant between that read and the request) is **detected**, named, paused, never repaired by guessing. **Writer model**: `config.pilot.writerModel = 'single'`: a banner on every ledger screen, a confirmation before the first Correct/Delete of a session, refusal to write from a second tab of the same browser, and documentation that simultaneous editing is not supported. The three ways beyond one writer are presented below for Cody's decision, not chosen. | Shift before the read: **prevented**. Shift after the read: **detected** only. | Unit: "S1: another member deletes #2 just before our check … every other transaction survives"; "S1: a correction whose old copy moved just before the check … every transaction survives"; "S1: renumbering re-reads before the positional PATCH … no historical row renumbered"; residual race tests renamed to say "detected … never repaired by guessing" and extended with verified resolution. Browser: "pilot writer model" ×2. Re-review reproduction rerun (shift injected at the DELETE itself): still detected, #3 still lost, reported as such, not as preserved. |
| S2 | A newer draft was erased on reload when an older save had landed | `restorePending` no longer touches the current draft or its generation; `_landed` declares the draft saved only when its content equals the payload that landed. A newer draft stays on screen, persisted and unsaved, and becomes its own operation on the next Save. Test mode now keeps the in-memory workbook in `sessionStorage`, so browser tests reload against the same workbook state, as the re-review asked. | Prevented | Unit: "S2: a restored pending operation and a NEWER draft …", "S2: the same payload restored after a reload …", "S2: … keeps the newer edits unsaved even if the generation happens to match". Browser: "S2: a newer draft survives a reload when the older attempt had landed: startup checking writes nothing, the newer text stays unsaved, and each version is saved once". Re-review reproduction rerun: `state=unsaved`, draft "Later unsaved version" kept. |
| S3 | "Check workbook" performed writes (formats, renumber, sorted table, old-copy delete), including on automatic recovery | The controller now has two callbacks: `inspect` (read-only) used by `check()` and by the automatic check after a reload, and `resume` (may finish) used only by a Save. The adapter gained `inspectAppend` and `inspectCorrection`, which issue GETs only and return `missing` / `landed` / `incomplete` (with the unfinished steps) / `conflict`. An incomplete operation is shown as "Not finished: … press Save to finish"; the Save button reads "Finish saving". | Prevented | Unit: "S3: check() is read-only …", "S3: Save after an incomplete check finishes the operation through resume, once", "S3: inspectAppend and inspectCorrection issue only GETs …", "S3: with the SaveController, Check after a failed correction only reads …" (asserts zero non-GET requests). Browser: "S3: Check workbook only reads …" and the S2 test assert zero writes since page load until Save. Re-review reproduction rerun: `mutations during check = 0`. |
| S4 | Refresh (and `load()`) removed the pause without resolving anything; incidents were not persisted; correction recovery state was not persisted | `load()` never clears the pause. An **incident** record (kind, message, details, checks) is kept in the browser per tenant/account/workbook and re-armed on reload. `verifyResolution()` (read-only) evaluates documented conditions per kind (no identity twice; named row present once; corrected record has one copy; operation id once; number used once; renumbered row has its number back) and lifts the pause only when all hold; `acknowledgeIncident()` is allowed only when a condition is unverifiable and refused while any fails. `load()` also scans for duplicate identities already in the table (pasted in Excel or left by another device) and pauses (`integrity`), except for this page's own known unfinished correction. Unfinished corrections (record, change, id) are persisted and reopened after a reload with a read-only check. The band shows the checks (OK / Not yet / Cannot tell), keeps "Refresh (stays paused)", and offers Dismiss only once resolved. | Prevented (state loss); resolution verified | Unit: "R4/S4: … load() does not lift the pause, a verified explicit repair does", "S4: load() pauses on duplicate copies already in the workbook …", "S4: acknowledgement is refused while a check fails …", plus the extended residual-race tests. Browser: "S1/S4: … Refresh and a page reload keep the pause; only a verified explicit restore lifts it", "S4: duplicate copies already in the workbook pause saving on load …", "S4: an unconfirmed correction survives a reload …". Re-review reproduction rerun: `paused after load = true`, `verifyResolution.resolved = false` with the three failing checks listed. |

### Also changed in this pass
- **Timestamp / operation id format** changed from `ISO#random` to **ISO 8601 with nine fraction digits**
  (`2026-09-06T14:03:11.482913057Z`), which still parses with `new Date()`; the re-review's compatibility
  concern about the old Office Scripts is narrowed to "do they parse or compare this column?", listed as a
  pre-live check for Cody in the setup guide (F). Their source is not in the file; nothing was changed in it.
- `removeCopy(…, { allowIdentical: true })` removes exactly one of several identical copies and proves the
  count afterwards (needed to resolve an `integrity` incident from a pasted duplicate).
- Mock: `exportState`/`importState`/`onChange` for test-mode persistence across a reload.
- Drafts schema 3: `{ entry, pending, correction }`.

### The writer-model decision (for Cody)
Simultaneous editing of the transaction table cannot be made safe from a browser: Graph has no lock,
conditional write or transaction, and deletes/renumbering are positional. The pilot is therefore
**single-writer, on the test copy**, enforced where a page can (banner, per-session confirmation before the
first Correct/Delete, second-tab refusal) and by agreement where it cannot (other devices, Excel, the old
buttons). Options beyond that, each with its remaining risk:
1. **Turn-taking by agreement (current).** Risk: a broken agreement can still produce a detected-only
   incident; recovery of a row changed by someone else in that instant needs Excel version history.
2. **One path for every change**: remove the three buttons, lock `LOG`, and add a serialising backend for
   the app. Risk: a real backend to host and secure; direct Excel edits by the treasurer would also have to
   stop or go through it; it does not help until *all* other writers are removed.
3. **Workbook redesign without positional operations**: a status column marks voided rows; corrections and
   deletes become appends plus single-cell writes located by identity; month and annual formulas exclude
   voided rows. Risk: a workbook change needing the treasurer's approval and re-verification of 550+
   formulas; the single-cell write is still positional but its misdirection would void the wrong row
   reversibly rather than delete it.
The recommendation is option 1 for the pilot and a decision on 2 or 3 only if the pilot shows turn-taking
is impractical.

### Tests run (2026-09-07, this folder)
- `npm test`: **120 passed, 0 failed** (Graph client 12, mock 10, ledger adapter 32, concurrency 5, save
  controller 14, incidents + tab guard 3, drafts 5, locator 4, store 3, storage model 14, ledger model 9,
  UI 6, export 3; the fixture-dependent formula-oracle tests ran because the git-ignored fixtures are present).
- `npx playwright test`: **123 passed, 0 failed** = 41 per project × 3 projects (desktop Chromium,
  iPhone-13-size Chromium, iPhone-13-size WebKit): ledger 13, review regressions 12, diagnostics 4,
  launcher 3, storage 9.
- Re-review reproductions (`review-checkpoint-1-rereview/reproductions.mjs`, adapted to the current
  API because `verify` was split into `inspect`/`resume`): S2 and S3 no longer reproduce; S4 no longer
  reproduces (the pause survives `load()`); S1 with the shift injected at the DELETE itself still loses the
  row and is reported as **detected, not prevented**; with the shift injected before the check read,
  every row is preserved.
- The reviewer's browser demonstration (draft loss after a persisted append and reload) has its safe
  counterpart in `tests/e2e/review-regressions.spec.js` ("S2 …"), using the persisted test-mode workbook.

### Still mocked, not verified against Microsoft 365
- Whether a table read immediately followed by a row DELETE behaves as modelled (no reordering inside the
  service), `rows/itemAt(index=n)/range` PATCH, number-format inheritance on appended rows, session
  lifetime, throttling, how quickly a read reflects another writer's change, the Annual and month sheet
  layouts of the online copy, the real Timestamp values.
- Whether the old Office Scripts tolerate nine-digit timestamps (pre-live check for Cody).
- A real iPhone (WebKit emulation is not iOS Safari), the MSAL redirect on the hosted address, GitHub Pages.

### Not started
Live verification (needs the app registration), single-tester pilot on the test copy, winter-storage
adapter (Phase 4), year rollover (excluded until saving is proven).

### Changes made to any real workbook
None. Nothing has connected to Microsoft 365.

### Requires Cody
1. **Checkpoint 1 re-review** of: `docs/workbook-mapping.md` (§2 Timestamp, §3, §3b), `docs/member-guide.md`,
   `docs/setup-guide.md` (E, F), `README.md`, `site/js/config.js` (pilot), `site/js/workbook/ledger-workbook.js`,
   `site/js/save/save-controller.js`, `site/js/save/incidents.js`, `site/js/save/tab-guard.js`,
   `site/js/ledger/app.js`, `tests/unit/{save-controller,ledger-workbook,concurrency,incidents,mock-excel}.test.js`,
   `tests/e2e/review-regressions.spec.js`, `tests/e2e/ledger.spec.js`.
2. Decision: the writer model (option 1 recommended for the pilot).
3. Decision (still open): transaction-number behaviour (max + 1; only a just-added row is renumbered).
4. Before the live write test: check what the three Office Scripts do with the Timestamp column (setup guide F).
5. Only after re-review: hosting account, app registration, assignment (setup guide A–C), then Diagnostics on
   the test copy with one writer (Checkpoint 2).

---

## 2026-09-07 — Checkpoint 1 review remediation (R1–R6)

Reviewed by Codex the same day ("changes required", S1–S4 above). Statements below that the re-review showed
to be wrong or overstated are marked **[corrected]**.

| # | Finding (P1) | What changed | Status after the re-review |
|---|---|---|---|
| R1 | A correction could destroy a different transaction and falsely claim it was restored | Correction became append + verified delete with identity-based assessment and no restoration. **[corrected]** The read+delete batch did not evaluate the read before the DELETE; the earlier claim "no longer reproduces" was an overstatement, since the reviewer's shape still lost the row. Now: prevented when the shift precedes the check read; detected (not prevented) when it falls between the read and the request; incidents verified before saving resumes. |
| R2 | Deleting one transaction duplicated another member's legitimate correction | Identity-based assessment; other rows are never "restored" from stale values. Unchanged; still passing. |
| R3 | The ordinary Save duplicated an append after a lost response | Pending operation recorded before the write; every Save/Retry resumes under the same marker. Unchanged, plus S3's read-only check. |
| R4 | A timestamp collision silently lost a different member's append | Unique operation id; collision refused. Format changed again (ISO with nine fraction digits) for script compatibility. |
| R5 | The client automatically retried correction writes | Only GET retried. Unchanged; still passing. |
| R6 | Later correction edits were hidden and stopped counting as unsaved | Dialog stays open unless the controller ends `saved`; later saves target the corrected copy. Unchanged; corrections' pending state now also persisted (S4). |

Other items from that pass that still hold: `completeCorrection`, the paused band (now incident-based),
the Annual cross-check, Diagnostics gating and compared/not-compared lists, the concurrency tests with two
adapters on one mock, the documentation corrections listed in the previous version of this entry.

---

## 2026-09-06 — Milestone 1: workbook mapping, Graph adapter, ledger screen on the workbook (Checkpoint 1)

Reviewed by Codex on 2026-09-06/07: "Changes required" (R1–R6). Statements below that the reviews showed
to be wrong are marked **[corrected]** and superseded by the entries above.

### Completed (implemented and tested locally against an in-memory replica; nothing yet against Microsoft 365)
- **Workbook mapping** `docs/workbook-mapping.md`: sheets, tables, columns, cell formats, website action → Graph
  operation, verification for each, findings. Read-only inspection of the local copy with openpyxl.
- **Graph Excel client** `site/js/workbook/excel-client.js`: persistent session, 202 long-running session
  creation, strictly sequential requests, Microsoft's second-level error codes, `Retry-After`, session
  re-creation. **[corrected]** the original retried some writes (R5); now only GET is retried. `$batch` is
  implemented and tested but no longer used by the adapter.
- **In-memory workbook mock** `site/js/workbook/mock-excel.js`: answers the same URLs, mirrors the month and
  annual sheet formulas, refuses writes to formula cells and outside table bodies, injects failures.
- **Ledger adapter** `site/js/workbook/ledger-workbook.js`. **[corrected]** twice; see the entries above.
- **Save workflow** `site/js/save/save-controller.js` and `site/js/save/drafts.js`. **[corrected]** twice (R3, S2/S3).
- **Ledger screen** `site/ledger.html`, `site/js/ledger/app.js`. **[corrected]** the Annual cross-check was missing;
  the Correct dialog hid later edits (R6); Refresh lifted the pause (S4).
- **Diagnostics page**: sign-in status, identify workbook, six read-only checks, self-test gated by the checks
  passing and by typing the workbook name; compared / not-compared lists.
- **Removed**: the JSON-file OneDrive store and its Graph file backend, the ledger's generated Excel/CSV
  export, JSON import/export. `site/js/store.js` is a browser-only cache for the storage screen.
- **Fixtures**: everything derived from the real workbooks lives in git-ignored `data/fixtures/`; tests that
  need them skip when absent; client, mock and adapter tests use a synthetic sample.

### Status of the five earlier defects (from the 2026-09-06 directive)
| # | Finding | Status |
|---|---|---|
| 1 | Lost edit while another save uploads | Fixed: `SaveController` ignores a second submit while saving, keeps later edits, ends in `unsaved` if anything changed during the save. |
| 2 | Storage conflict picks a whole record by timestamp | Ledger: per-row fingerprint check with a conflict dialog. Storage screen: the merge path is gone; its future adapter will use the same approach. |
| 3 | Sign-out clears data without protecting unsaved edits | Fixed: `requestSignOut` asks first (also when an incident is paused); drafts cleared only for the signing-out account. |
| 4 | Cache keys ignore account/tenant/workbook | Fixed: `draftKey`/`incidentKey` include tenant, account and workbook; workbook choice remembered per app + tenant + account. |
| 5 | Mocks do not prove cloud correctness | Still true: the mock speaks Graph's HTTP surface so the real client and adapter are exercised, but **cloud correctness is unproven** until the Diagnostics self-test runs on the real copy (Checkpoint 2). |

### Exact workbook operations implemented (all via Graph v1.0, inside one session, sequential)
`GET /drives/{d}/items/{i}` (metadata) · `POST …/workbook/createSession` / `closeSession` ·
`GET …/worksheets/{s}/range(address=…)` · `PATCH …/worksheets/{s}/range(address=…)` (values, numberFormat) ·
`GET …/tables/{t}/dataBodyRange` · `POST …/tables/{t}/rows/add` · `DELETE …/tables/{t}/rows/{n}` ·
`GET/PATCH …/tables/{t}/rows/itemAt(index=n)/range`. `POST /$batch` and `POST …/application/calculate` exist
in the client and are **not called**. Sheets written: `LOG` (table rows), `LOG_Sorted` (table rows),
`ENTRY!B23`. Nothing else is ever written.

### Known problems, limitations, open questions
- **Concurrency**: see §3b of the mapping and the writer-model decision above.
- **`LOG_Sorted` dependency**: the month sheets are only right after the sorted helper table is rebuilt; the
  adapter rebuilds it after every change against a fresh LOG read and Diagnostics checks its consistency.
- **Month sheets show at most 30 rows** (workbook design); the screen shows all rows and warns.
- **Number formats on appended rows**: set explicitly by the adapter, single attempt; to be confirmed live.
- Workbook identity and content of the **online** copy unverified.
- The Winter Storage screen still shows the earlier browser-only behaviour, clearly labelled as not connected.

---

## 2026-09-06 (earlier) — Milestone 0: assessment
Superseded by Milestone 1; kept for the record: assessment, workbook re-inspection, Graph documentation
verification, mapping and plan written.
