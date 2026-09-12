# Checkpoint 2 runbook: the first live test on the test copy

One page to follow on the day. Checkpoint 1 was approved on 2026-09-08 for the locally tested implementation
and the single-writer test-copy pilot. **That approval says nothing about Microsoft 365 itself**: sign-in,
permissions, sessions, row formats and real writes are exactly what this test is for.

Everything below happens on the copy named **TPAHA_2026 (1)** in its owner's OneDrive. Not the association's
live records. Nobody is asked for a password by the site, and no password, token or client secret is ever sent
to the developer.

---

## Part 0. Publish the reviewed code — **prepared, not yet run**

Codex approved `7740e6c` on 2026-09-12. The unpublished work is `9e464cd`, `d5a0faa`, `68db038`, `fdab0a1`,
`7740e6c` — the monthly report rows and the four review fixes that followed — plus the commit that added this
Part 0. Step 0.2 lists them; read what it prints rather than trusting this sentence.

**Nothing in Part 0 touches the workbook.** It puts reviewed code on the website. The workbook test is Part 2
and needs its own go-ahead. Run Part 0 only when Cody says to publish.

**0.1 — Confirm the tree is clean and every suite passes.** In the project folder:

| Command | Expect |
|---|---|
| `git status --short` | no output at all |
| `npm test` | `pass 152`, `fail 0` |
| `npx playwright test` | `189 passed` |
| `npm run verify:config` | `21 passed, 0 failed` |

Stop if any number is lower or anything is uncommitted. A failing suite is not published.

**0.2 — Look at exactly what will go out.**

- `git log --oneline origin/main..HEAD` — expect the commits above and nothing you do not recognise.
- `git diff --stat origin/main..HEAD` — site code, tests and documentation only: no workbook, no `data/`.
- `git diff --name-only origin/main..HEAD | grep -iE "data/|\.xlsx|office-script"` — expect **no output**.

Stop if that last command prints anything: something private is staged, and it must be removed before pushing.
Public history cannot be un-published.

**0.3 — Push.** `git push origin main`

**0.4 — Watch the deployment finish.** GitHub → the repository → **Actions** → "Deploy TPAHA Books to GitHub
Pages". Wait for the green tick. If it fails, the previously deployed site stays up: nothing is broken for
anyone, so read the log rather than pushing again.

**0.5 — Verify what is being served, not what was pushed.**

1. Hard-refresh `https://tpahavp.github.io/tpaha-books/` (Ctrl+F5). The ledger page loads and there is **no**
   "Test mode" banner — on a published address the site is in connected mode.
2. Check the deployed configuration file is byte-for-byte the reviewed one. In PowerShell:
   ```
   Invoke-WebRequest https://tpahavp.github.io/tpaha-books/js/config.js -OutFile $env:TEMP\deployed-config.js
   Get-FileHash $env:TEMP\deployed-config.js -Algorithm SHA256
   ```
   Expect `BC25AC6C2BE1D4983C4844B5BC61C714CE9354C82C50C11281A9060CB499CE48`. Stop if it differs.
3. Delete the download: `Remove-Item $env:TEMP\deployed-config.js`

**0.6 — Leave it signed out.** Do not sign in yet. Publishing is finished; the workbook test is a separate
authorisation.

**If it has to be undone:** Pages serves whatever is on `main`. Revert the commit you want gone
(`git revert <sha>`) and push; the workflow redeploys the previous state. Never force-push — it rewrites public
history rather than correcting it.

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

Both items from that review are now closed:

- **`RefreshReports`** was supplied and read on 2026-09-12. It rebuilds `LOG_Sorted` from `LOG`, recalculates,
  verifies, and refreshes month-row visibility. It **adds and deletes no transactions**, so it is the one
  workbook button that is safe to press as a repair. It is still a writer: press it only when no website save
  is in flight. Mapping §7.
- **Hidden month rows.** No longer a standing limitation: the app sets those rows itself after every save,
  delete and correction, and says so when it cannot (mapping §8). Two things about that are still unproven
  online and are checked in Part 2b below — that the pilot workbook's month sheets allow row formatting, and
  that the rows Excel then shows are the right ones. Do not press Submit or Delete to refresh visibility:
  those buttons change data.

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
**Done and published, 2026-09-11.** The site is live at `https://tpahavp.github.io/tpaha-books/` from
`https://github.com/TPAHAVP/tpaha-books`, the Entra app registration is complete, and `site/js/config.js`
carries its ids (verified by `npm run verify:config`, 21 checks). The deployed `js/config.js` was checked
byte-for-byte against the committed file. The live site now offers Microsoft sign-in.

**Steps 1 to 3 of Part 2 were completed on 2026-09-11 and passed**: an assigned account signed in, the picker
listed the intended `.xlsx`, it was confirmed, and all six read-only checks passed with every content request
returning 200. Still to confirm by hand: that an **unassigned** account is refused by Microsoft (AADSTS50105),
and the same sign-in on a real phone. Step 4, the write test, remains unauthorised.

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

## Part 2b. Live test of the monthly report rows

New in 2026-09-12 and not yet run against the real workbook. Do this only with the write test authorised, since
it saves and deletes a real transaction on the test copy. The connection test in step 4 already exercises the
same write on December, so if that passed, this is the confirmation in Excel.

0. **This is also the check that the online copy's protection allows row formatting.** The local copy allows it;
   that proves nothing about the pilot workbook. If the online sheets disallow it, step 3 will show the green
   band with a message naming the sheet and the setting, and that is the answer.
1. Open `TPAHA_2026 (1).xlsx` in Excel and go to a month sheet with few transactions, for example September.
   Note which of rows 4 to 33 are hidden. Close the workbook so the website is the only writer.
2. On the website, add one clearly labelled transaction dated in that month. Wait for **Saved to workbook**.
3. Confirm no green band appears. If it does, note its wording, press **Finish report formatting**, and note
   what happens. A message naming "Format rows" means the online copy's protection disallows it; send that
   wording rather than changing any protection setting yourself.
4. Open the workbook in Excel again, go to that month sheet, and confirm the new transaction's row is **visible**
   and the blank rows below it are hidden. Use File, Print, Preview and confirm the row appears there too.
5. **Only if the rows are wrong in Excel** and the website cannot fix them: the repair is **Run RefreshReports**
   in Excel — reviewed 2026-09-12, it adds and deletes no transactions. Close the workbook's save first and
   press it once. Do not press Run SubmitEntry or Run DeleteTransaction: those change transactions. Record that
   you needed it, because needing it is itself a result worth reporting.
6. Back on the website, delete that transaction. Confirm the website removes it.
7. Reopen the workbook in Excel and confirm the row is hidden again and the sheet looks as it did in step 1.

What to send: which rows were hidden at step 1 and step 7, whether the green band appeared, and whether the
printed preview showed the new row. Stop and report if the row stays hidden in Excel after step 4, or if the
green band cannot be cleared.

## Part 3. Cleanup: put the test copy back as you found it

Do this on the same day, before reporting. The connection test in Part 2 step 4 cleans up after itself; Part 2b
adds a real transaction that you delete by hand. This part is how you prove both.

**3.1 — On the website, with the workbook closed in Excel.**

1. Press **Refresh**.
2. The transaction count matches the number you started with. Write down both.
3. No transaction described *TPAHA Books connection test* is listed. Search for it.
4. The transaction you added in Part 2b is gone.
5. No green "report formatting" band. If one is showing, press **Finish report formatting** once and wait for
   it to clear. If it will not clear, stop and report it — do not keep pressing.
6. No red paused band and no unresolved incident. If either is showing, stop and report it. Do not clear it
   yourself: a paused band means the app believes the workbook needs a person to look.

**3.2 — In Excel, opening the workbook once.**

1. The transaction table holds the rows it held before, and no test rows.
2. The month sheet used in Part 2b looks as it did at Part 2b step 1 — the same rows hidden.
3. **File → Info → Version history**: the versions created today are the ones this test explains. Note how many.
4. Close the workbook. Do not leave it open in Excel.

**3.3 — Leave nothing running.**

1. Sign out of the website.
2. Close any second tab or device that had the site open.
3. If the separately authorised script-identity check was run, the throwaway transaction from its step 3 is
   deleted, and say so in the report.

**If cleanup cannot be completed** — a test row will not delete, the count is wrong, or a band will not clear —
stop there and report it with the request log. A workbook left in an unexpected state is a finding, not a
failure to hide, and it is why this is a test copy.

## Part 4. What to send back for Checkpoint 2

- The read-only check results (all six lines).
- The connection test result line, plus its "compared" and "not compared" lists.
- The request log from **Copy log**.
- What Version history showed, and how many versions today's test created.
- From Part 2b: which rows were hidden before and after, whether the green band appeared and what it said,
  whether the print preview showed the new row, and whether Run RefreshReports was needed.
- From Part 3: the transaction count before and after, and confirmation that no test row remains.
- The answers from prerequisite B.

## A separate, separately authorised check: does a script preserve a row's identity?

**Not part of the run above, and not authorised by any review so far.** All four scripts have now been read
(2026-09-12), so the remaining condition is Cody's word: do it only when he says so, and never while a web page
has a save in flight.

Why it is needed: `SubmitEntry` and `DeleteTransaction` rewrite the **whole of `LOG_Table`** on every save and
delete — `saveChange()` calls `writeRows(log, after)`, which sets every body row — and that includes the
Timestamp text this app uses to recognise its own rows. Nothing tested so far proves a real Office Script
preserves a nine-fraction-digit timestamp. The connection test in Part 2 never runs a script, and the unit test
that simulates a rewrite proves only the mock's behaviour.

**`RefreshReports` cannot stand in for this check.** It never writes `LOG_Table` — that is exactly what makes it
safe as a repair — so running it proves nothing about whether a rewrite of that table preserves the website's
row identity. Only the two buttons that rewrite `LOG_Table` can answer the question, and both of them change
transactions. That is the awkward part of this check and the reason it needs its own authorisation and a test
copy: proving the identity survives means letting a transaction-changing button run once.

(`RefreshReports` does copy the Timestamp text into `LOG_Sorted` through the same `writeRows()` helper, so if it
ever came back altered there, that would be worth reporting. It is a weaker, different signal: it would break
`verifySorted` rather than the app's identity matching, which reads `LOG_Table`. Do not record it as the
identity check.)

On the test copy, one step at a time:

1. Save one clearly labelled transaction from the web page. Wait until the line says Saved.
2. Read that row's Timestamp text and write it down **exactly**, character for character. The LOG sheet in
   Excel shows it.
3. With no web page saving anything, press **Run SubmitEntry** once to add a throwaway transaction of your own
   from the ENTRY form (or **Run DeleteTransaction** once on that throwaway). Either rewrites every row of
   `LOG_Table`, which is the point: it is the website's row you are watching, not the one you entered.
4. Read the website's row again and compare its Timestamp text with what you wrote down.
5. Delete the throwaway transaction afterwards, and say in the report that you did.

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
