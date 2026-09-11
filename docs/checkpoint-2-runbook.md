# Checkpoint 2 runbook: the first live test on the test copy

One page to follow on the day. Checkpoint 1 was approved on 2026-09-08 for the locally tested implementation
and the single-writer test-copy pilot. **That approval says nothing about Microsoft 365 itself**: sign-in,
permissions, sessions, row formats and real writes are exactly what this test is for.

Everything below happens on the copy named **TPAHA_2026 (1)** in its owner's OneDrive. Not the association's
live records. Nobody is asked for a password by the site, and no password, token or client secret is ever sent
to the developer.

---

## Part 1. Before the day: five prerequisites

All five must be true before the write test is started. Any one missing means stop.

### A. Checkpoint 1 approved
Done, 2026-09-08 (`IMPLEMENTATION_STATUS.md`).

### B. The three Office Scripts — **read on 2026-09-10**

Their source was supplied and checked. Result: **none of the three reads, parses or compares the Timestamp
column**, so the app's nine-digit operation ids do not affect them. One real incompatibility was found and
fixed the same day (the app was sorting the helper table by table position where the scripts require date then
transaction number, which would have made the buttons refuse to run after certain corrections). Full record in
`docs/workbook-mapping.md` §7.

Two items from that review are still open and belong with the decisions below:

- **`RefreshReports`**, a fourth script named in the others' error messages, has not been read. Copy it the
  same way and send it before the pilot.
- **Hidden month rows.** The scripts hide empty month-sheet rows; this app does not unhide them, so a
  transaction added on the web can stay hidden when that month sheet is opened or printed in Excel. Decide
  whether to accept that for the pilot, with a documented manual refresh using `RefreshReports` once its source
  has been reviewed and it is confirmed to reveal populated rows. Check the Excel view or printout before
  relying on it. Do not press Submit or Delete merely to refresh visibility: those buttons change data.

### C. Two decisions — **accepted by Cody on 2026-09-08**
- **Writer arrangement: one writer at a time.** One designated tester works at a time; nobody edits the
  workbook in Excel or runs its script buttons meanwhile. The site enforces what a browser can see (a banner,
  a confirmation before the first Correct or Delete, refusal to write from a second tab, one operation at a
  time) and nothing beyond that. Simultaneous editing is not supported.
- **Transaction numbering: highest existing number plus one**, with its consequence accepted: **if the
  highest-numbered transaction is deleted its number can be given to a later entry**. Two members adding at
  the same moment both land, and the later one is renumbered to the next free number and told. Existing
  entries are never renumbered by a correction.

Nothing needs changing for either decision; both describe the behaviour already built and reviewed.

### C2. Hidden month rows — **still to decide** (see prerequisite B)
Not a blocker for the setup work, but settle it before the pilot runs.

### D. Hosting and Microsoft setup finished
**Done on 2026-09-11, but not yet published.** The site is hosted at `https://tpahavp.github.io/tpaha-books/`
from `https://github.com/TPAHAVP/tpaha-books`, the Entra app registration is complete, and `site/js/config.js`
now carries its ids (verified by `npm run verify:config`, 17 checks). The commit has not been pushed, so the
live page is still the test-mode one. **Publishing that commit is what turns the live site into a real sign-in
page that can reach the workbook**; do it deliberately, and confirm afterwards that an assigned account can
sign in and an unassigned one is refused.

`docs/setup-guide.md` parts A to D: the site published over HTTPS, the Entra app registration (single tenant,
SPA redirect URI = the site address, `Files.ReadWrite` + `User.Read`, admin consent, **no client secret**),
"Assignment required" set to Yes with only the intended accounts assigned, and `clientId` / `tenantId` filled
into `site/js/config.js`. Confirm an unassigned account is refused at sign-in.

### E. The workbook is quiet
Nobody has the test copy open for editing in Excel, no script is running, and no second TPAHA Books tab or
device is in use while the test runs.

---

## Part 2. On the day: the run

Before starting, open the workbook in Excel once and note **File → Info → Version history**: the newest
version and its time. That is the "before" marker.

1. Open `<site address>diagnostics.html` and press **Sign in with Microsoft**. Expect the Microsoft sign-in
   page, then a return to Diagnostics showing the signed-in account.
2. **Identify the workbook**: press **Find**. Confirm the row whose name is exactly `TPAHA_2026 (1).xlsx`,
   checking the folder, size and last-changed time. Press **Choose**.
3. **Read-only checks** run by themselves: workbook opens, transaction table with the expected columns,
   sorted helper table, workbook year, prior-year balance, sorted table consistent with the transaction
   table. **All six must show OK.** While any shows Problem the write test cannot be started and the page
   says why. This is also the first real evidence about Graph: capture it either way.
4. **The write test.** Type the workbook name exactly as shown, press **Run the connection test**. It adds
   one row described *TPAHA Books connection test*, reads it back, checks the sorted table contains it,
   deletes it, confirms it is gone, and compares these areas with their state before the test: the
   transaction table's values and number formats, the sorted helper table's values and number formats,
   `ENTRY!B23`, the formulas of `January!A4:O40` and of `Annual 2026!A3:N19`, and both tables' row counts.
   It does **not** compare other worksheets, formatting outside those ranges, table styles, names,
   protection, data validation, or workbook-level features; the page lists both lists.
5. **Copy log** and keep it. It holds method, path, status, timing and request ids. No records.
6. Open the workbook in Excel again. Compare **Version history** with the "before" marker, look at the
   versions the test created, and confirm the transaction table looks untouched and no row described
   *TPAHA Books connection test* remains.

## Part 3. What to send back for Checkpoint 2

- The read-only check results (all six lines).
- The connection test result line, plus its "compared" and "not compared" lists.
- The request log from **Copy log**.
- What Version history showed.
- The answers from prerequisite B.

## A separate, separately authorised check: does a script preserve a row's identity?

**Not part of the run above, and not authorised by any review so far.** Do it only when Cody says so, after
`RefreshReports` and the other scripts have been reviewed, and never while a web page has a save in flight.

Why it is needed: the scripts rewrite the whole transaction table on every save and delete, including the
Timestamp text this app uses to recognise its own rows. Nothing tested so far proves that a real Office Script
preserves a nine-fraction-digit timestamp. The connection test in Part 2 never runs a script, and the unit test
that simulates a rewrite proves only the mock's behaviour.

On the test copy, one step at a time:

1. Save one clearly labelled transaction from the web page. Wait until the line says Saved.
2. Read that row's Timestamp text and write it down **exactly**, character for character. The LOG sheet in
   Excel shows it.
3. With no web page saving anything, run the reviewed script action once.
4. Read the row again and compare the Timestamp text with what you wrote down.

Identical means the identity survives and the pilot can rely on it. Any difference means stop: do not retry an
unconfirmed write, and send the before and after text. A changed identity is a design question, not a retry.

## Stop conditions

Stop and report, rather than retrying, if any of these happen:

- any read-only check shows Problem;
- the connection test stops at a step, or ends without saying the compared areas were restored;
- the page reports **Changes are paused** (an unresolved change): send the details it names, and do not
  press repair buttons on the real copy without agreeing first;
- a row described *TPAHA Books connection test* is still in the workbook afterwards. Note it, leave it, and
  send the log; deleting it by hand is fine once the log has been captured.

## What this test does and does not establish

It establishes that this account can sign in, that the app can open **this** workbook through Microsoft
Graph, and that one add, one read-back and one delete behave as the app expects on the real service,
with the compared areas unchanged.

It never runs one of the workbook's own Office Scripts, so it says nothing about what those do to rows this app
has written; that is the separate check above. It does not establish anything about the live records, about two
people working at once, about phones (the
browser tests used two engines, never a physical device), or about any workbook other than this copy. Winter
Storage is still browser-only and is not part of this test.
