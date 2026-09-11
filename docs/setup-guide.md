# One-time setup guide (administrator)

This is done once, by whoever administers TPAHA's Microsoft 365 and by whoever will own the hosting
account. Board members never do any of this; their guide is `docs/member-guide.md`.

Nothing in this guide asks for anyone's password, and nothing here changes the workbook.

## What the pieces are

| Piece | What it is | Cost |
|---|---|---|
| Hosting | A web address that serves the site's files over HTTPS | Free tier |
| App registration | Microsoft's record that "TPAHA Books" is allowed to ask members to sign in and to use their OneDrive access on their behalf | Free |
| Assignment | The list of accounts allowed to sign in to TPAHA Books | Free |
| The workbook | The Excel file in OneDrive. It stays where it is. The site reads and writes it through Microsoft's Excel service when a signed-in member presses Save. | — |

## A. Hosting (GitHub Pages)

1. Create a GitHub account controlled by the association (use a TPAHA mailbox, not a personal one).
2. Create a repository named `tpaha-books`. Public is fine: the repository contains the site's code only. Real records live in the workbook and are never in the repository (`data/` and the `.xlsx` files are ignored by git).
3. Upload this project to the repository (GitHub's "uploading an existing project" instructions), or ask the developer to push it.
4. In the repository: **Settings → Pages → Source: GitHub Actions**. The workflow in `.github/workflows/pages.yml` publishes the `site/` folder on every push to `main`.
5. After the first run the site address appears, for example `https://tpaha.github.io/tpaha-books/`. Note it exactly, including the trailing slash: it is the **redirect URI** in part B.

Limits (GitHub's published figures): sites up to 1 GB, a soft limit of 100 GB of traffic per month, 10 builds per hour (not applicable with the Actions workflow). The site is a few megabytes and the board is a handful of people, so these are not a concern. HTTPS is provided for github.io addresses. The site itself is public on the internet: that is acceptable because the pages hold no records; records are only shown to a member who has signed in and who already has access to the workbook.

Alternatives with the same properties: Cloudflare Pages or Netlify (both free for this size and both allow a private repository). A purchased domain is unnecessary; `books.tpaha.ca` could be added later if wanted.

## B. Register the app in Microsoft 365 (Entra)

1. Go to <https://entra.microsoft.com> and sign in with a TPAHA administrator account.
2. **Identity → Applications → App registrations → New registration.**
3. Name: `TPAHA Books`. Supported account types: **Accounts in this organizational directory only (TPAHA only – Single tenant)**.
4. Redirect URI: platform **Single-page application (SPA)**, value = the site address from part A (with trailing slash). Add a second one later for a custom domain if you adopt one. Click **Register**.
5. On the Overview page copy two values for part D: **Application (client) ID** and **Directory (tenant) ID**. These are identifiers, not secrets.
6. **API permissions → Add a permission → Microsoft Graph → Delegated permissions**: tick `Files.ReadWrite` and `User.Read`, then **Add permissions**. Then click **Grant admin consent for TPAHA** so members are not each asked to consent.
7. **Never create a client secret or certificate for this app.** The site runs entirely in the member's browser and does not need one; a secret in a web page would be public.

### What `Files.ReadWrite` means, honestly

It is the least-privileged permission Microsoft lists for the Excel workbook API. It lets the site, *while a member is signed in and only as that member*, read and change files that this member can already open in OneDrive or SharePoint. It does not grant access to anyone else's files and it grants nothing when nobody is signed in. The site only ever uses it on the one workbook the member has confirmed on the Diagnostics page or that is pinned in the configuration. There is no narrower delegated permission that still allows writing to an Excel table.

## C. Restrict sign-in to approved board accounts

The tenant setting in step B3 already blocks every account outside TPAHA. To limit it further to named board members:

1. **Identity → Applications → Enterprise applications → TPAHA Books → Properties**: set **Assignment required?** to **Yes**, then **Save**.
2. **Users and groups → Add user/group**: select each board member who should use the site, **Select**, then **Assign**. (Assigning a group instead of individuals needs a Microsoft Entra ID P1 licence; individuals work with any licence.)
3. An account that is not assigned sees a Microsoft error at sign-in ("the signed in user is not assigned to a role for the application", AADSTS50105) and never reaches the site's data.

Remember: sign-in permission and workbook permission are separate. A member also needs edit access to the workbook itself in OneDrive/SharePoint, exactly as if they opened it in Excel. Sharing the workbook is a decision for the association, not something the site does.

## D. Configure the site

**Done on 2026-09-11.** `site/js/config.js` carries the Application (client) ID and the Directory (tenant) ID
from part B. Both are identifiers, not secrets, which is why they sit in a file the site publishes; who may
sign in and what they may open is still decided by Microsoft 365. There is no client secret and there must
never be one.

One thing to understand before changing that file. The configuration hands those ids only to the **published
host**; a page served from a developer's own machine (`localhost`, `127.0.0.1`) gets blank ids and therefore
runs in test mode against the in-memory sample workbook. That is how the automated tests and local development
stay offline. The rule is the address the page was served from, not a flag or a query parameter, so nothing a
member can click, type or paste can put the real site into test mode.

Run `npm run verify:config` after any change here. It checks the id formats, the host rule, the redirect URI
the app actually derives, the scopes requested, and that no secret has crept into the site. It never contacts
Microsoft.

Leave `workbooks.ledger` blank for now: each member picks and confirms the workbook once (Diagnostics page or first open) and the choice is remembered in their browser, per account. Once the association has decided which workbook is the live one, its drive id and item id (shown on the Diagnostics page) can be pinned here so nobody has to choose.

Commit and push. Pages redeploys within a couple of minutes. Change `appVersion` whenever you deploy so phones
pick up the new files; it is `2026.09.11` for this change.

**What publishing this change does.** Until it is pushed, the live site is a harmless test-mode page. Once it
is, the published site asks assigned members to sign in and can read and write the workbook they choose. That
is the intended next stage, but pick the moment deliberately rather than letting it ride along with an
unrelated push.

## E. The pilot workbook and the one-writer rule

The pilot is the copy named "TPAHA_2026 (1)" in its owner's OneDrive. Only that owner can reach it until it is shared. Do not move the copy to a shared area or share it more widely without agreement.

**Writer model for the pilot: one designated writer at a time** (accepted by Cody on 2026-09-08). Microsoft Graph offers no lock or conditional write for Excel tables, so two people changing rows of the transaction table at the same instant can make a delete or correction hit the wrong row. The app reads and evaluates the table right before every such request, detects and names a misdirected one afterwards, pauses, and offers explicit repairs, but it cannot prevent a shift that happens in the instant between its read and its request (`docs/workbook-mapping.md` §3b). The pilot therefore does **not** support simultaneous editing, and the app says so:

- a blue band on every ledger screen states the rule (`config.js` → `pilot.writerModel: 'single'`, `pilot.note`);
- before the first Correct or Delete of a browser session the member confirms that nobody else is writing (another member on the page, another device or tab, Excel, the old script buttons);
- a second tab of the same browser is detected and refuses to write; this covers the ledger and the Diagnostics page together (whichever TPAHA Books tab opened first in that browser owns writing; the Diagnostics write test is refused in any other tab, while its read-only checks still run). A browser without BroadcastChannel cannot detect other tabs and says so; use one tab there;
- within one page, operations run one at a time: a second Save, Delete or repair is refused while one is running, and the adapter queues anything that slips through so every operation reads fresh row positions first;
- what a page cannot detect is covered by agreement: while a tester is using the page, nobody opens the workbook for editing in Excel or runs its script buttons. A single-tester pilot on the test copy is the recommended first step; a second tester joins only for turn-taking, never at the same time.

Going beyond one writer is a board decision with three options, none adopted here: keep turn-taking by agreement; remove the old buttons and lock `LOG` so that every change goes through one path (then a serialising backend for the app becomes meaningful); or change the workbook so that no positional operation is needed (a status column for voided rows, with the month and annual formulas excluding them). `IMPLEMENTATION_STATUS.md` presents these with their remaining risks.

## F. First live test (Checkpoint 2)

**Use `docs/checkpoint-2-runbook.md` on the day.** It carries the five prerequisites (including how to read
the three Office Scripts and what to check in them), the run itself, what to send back, and the stop
conditions. The steps below are the same run, in short.

1. Open `<site address>diagnostics.html`, press **Sign in with Microsoft**.
2. **Identify workbook**: press Find, confirm the row whose name is exactly `TPAHA_2026 (1).xlsx` (check the folder, size and last-modified time), press Choose.
3. **Read-only checks** run automatically: workbook opens, transaction table present with the expected columns, sorted helper table present, workbook year, prior-year balance, sorted table consistent with the transaction table. All six must show OK. **While any check shows Problem, the write test cannot be started** (its button stays off and the page says why); nothing is written to a workbook that fails a check.
4. **Prove access**: type the workbook name exactly as shown, press **Run the connection test**. It adds one row described "TPAHA Books connection test", reads it back, deletes it, rebuilds the sorted helper table, and compares **exactly these areas** with their state before the test: the transaction table's values and number formats, the sorted helper table's values and number formats, `ENTRY!B23`, the formulas of `January!A4:O40` and of `Annual <year>!A3:N19`, and both tables' row counts. It does **not** compare other worksheets, formatting outside those ranges, table styles, names, protection, data validation or workbook-level features; the page lists what was and was not compared. Then open the workbook in Excel, use **File → Info → Version history** to look at the versions the test created, and confirm nothing else changed.
5. Press **Copy log** and paste the log into the status report or send it to the developer. The log contains request ids and timings, no records.

If any step fails, stop there; nothing further is attempted, and the log says which step and why.

Before step 4 (the first live write), two things the developer cannot check from here:

- **The old Office Scripts and the Timestamp column: done on 2026-09-10.** None of the three reads, compares or parses that column, so the app's nine-digit timestamps do not affect them. One incompatibility was found and fixed (helper-table sort order). Still to send before the pilot: the source of the fourth script, **RefreshReports**, named in the others' error messages. See `docs/workbook-mapping.md` §7.
- **The workbook is quiet.** Nobody has it open for editing in Excel, and no script is running, while the Diagnostics test runs (one-writer rule, part E).

## G. Going live later (Checkpoint 3)

Only after the pilot: agree which workbook is the live record, share it with the assigned members, pin its ids in `config.js`, remove the test copies, and re-run the Diagnostics read-only checks (not the write test) against the live workbook.
