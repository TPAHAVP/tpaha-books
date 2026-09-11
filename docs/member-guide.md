# Board member guide: TPAHA Books

TPAHA Books is a web page that lets you add and correct treasurer ledger entries from your phone or
computer. Everything you save goes straight into the association's Excel workbook in OneDrive, the
same file the treasurer opens in Excel. There is no separate copy.

## The pilot rule: one writer at a time

During the pilot, **one person writes at a time**. Do not use the page in two tabs or on two devices at
once, and do not type in the workbook in Excel or press its old script buttons while someone is using the
page. The page says so in a blue band, asks you to confirm it before your first Correct or Delete, and
refuses to save from a second tab of the same browser (the Diagnostics page counts as a tab too), and runs your own changes one at a time. The reason: Excel Online cannot lock a row for one
person, so two people changing rows at the same moment can affect the wrong row. The page detects that
and pauses (see below), but it cannot always undo it.

## Open it

1. Open the link you were given (add it to your phone's home screen for one-tap access: in Safari, Share → Add to Home Screen; in Chrome, the menu → Add to Home screen).
2. Press **Sign in with Microsoft** and sign in with your TPAHA account, the same one you use for email. You will never be asked for your password by the site itself; the sign-in page is Microsoft's.
3. The first time, confirm which workbook to use when asked (name, folder and last-changed time are shown). Choose the one the treasurer told you to use.

If you see "not assigned" from Microsoft, your account has not been added yet; ask the administrator.

## Add a transaction

Fill in the date, Deposit or Withdrawal, the category, the amount, a description, and the cheque number if any. Press **Save to workbook**.

Watch the line under the button:

| It says | It means |
|---|---|
| **Unsaved changes** | You have typed something that is not in the workbook yet. |
| **Saving to the workbook…** | Your entry is being written and then read back. Leave the page open. |
| **Checking the workbook (reading only)…** | The site is looking whether an earlier attempt reached the workbook. It writes nothing while checking. |
| **Saved to workbook · #31** | Done. Microsoft confirmed the write and the site read the row back. #31 is its transaction number. |
| **Could not save** | The workbook was not changed, or the answer was lost on the way. Press **Check workbook** (reads only) or **Save**: the site first looks in the workbook for that same attempt before it writes anything, so nothing is written twice. |
| **Not finished: your entry is in the workbook as #31, but…** | The row is there; its finishing steps (number formats, number check, the sorted helper table) are not confirmed. The button now says **Finish saving**; press it. The entry itself is not written again. |
| **Someone else changed this record** | Another member (or someone in Excel) changed the same row since you opened it, so nothing was written. Press **Reload record** to read it again: if what you typed differs from the version now in the workbook, the site shows both and asks whether to **keep what you typed** (it will then be saved onto the row as it now stands) or to **use the workbook version**. Nothing you typed is replaced without that question. |
| **Another change is still being saved…** (short message) | One change at a time: the previous save, delete or repair has not finished. Wait a moment and press again. You can keep typing meanwhile. |
| **Changes are paused** (red band at the top) | A change could not be verified; see "If something needs checking" below. |

You need an internet connection to save. If you lose signal, what you typed stays on the screen until you can save; it is not saved anywhere until the line says **Saved to workbook**. If you close the page or your phone reloads it, what you typed and any unconfirmed attempt come back; the site checks the workbook (reading only) and tells you what it found before you save again.

If you keep typing while a save is running, or after an attempt that was not confirmed, the newer text stays on screen and stays **Unsaved**. It is never reported as saved because an older attempt landed: the older attempt is finished first, then you press Save again for the newer text.

## Find, correct, delete

- The list shows the workbook's transactions. Use the date range or the search box to narrow it. **Refresh** re-reads the workbook, so changes made by others or directly in Excel appear.
- **Correct** opens the row so you can fix a typo or amount. Save works the same way as above. If you keep typing while a correction is saving, the saved version goes into the workbook and your newer text stays on screen as unsaved, aimed at that saved version; it survives a reload and saves normally when you press Save again. The corrected row keeps its number, but in the workbook it moves to the end of the transaction table (the site writes the corrected copy, checks it, then removes the old copy); the monthly sheets are unaffected because they sort by date. If the page is reloaded halfway, the correction reopens and is checked before anything more is written.
- **Delete** asks you to confirm, then removes the row from the workbook. Right before removing it the site reads the table once more and removes the row only if it is still exactly where it expects; otherwise it looks again or stops. If the row changed since you loaded it, you get the "someone else changed this record" message instead.

Transaction numbers: a new entry gets the highest existing number plus one. Two members adding at the same moment both land; if they drew the same number, one of them is renumbered to the next free number and told. If the highest-numbered transaction is deleted, its number can be given to a later entry. Existing entries are never renumbered when corrected.

## If something needs checking

If, despite the one-writer rule, two changes hit the same part of the table at the same instant, a delete or correction can affect the wrong row. The site compares the table before and after every such change and then:

- shows a red band **Changes are paused** with exactly what it knows: which row was affected (with the values this device last saw), whether two copies of a record now exist, and a list of checks that must pass before saving resumes;
- refuses every save until those checks pass. **Refresh and reloading the page keep the pause**; the record of the problem is kept in your browser;
- never repairs anything on its own. Where a repair is possible, the band offers a button (**Restore #N as last seen**, **Remove this copy**) that acts only after you confirm. Values restored this way are what *this device* last saw; a change made by someone else in between would not be included, so when in doubt open the workbook in Excel and use **File → Info → Version history** first;
- **Check whether this is resolved** re-reads the workbook and shows each check as OK, Not yet, or Cannot tell. When all are OK the band turns green ("Resolved, verified") and saving resumes. If something cannot be told from the table (the removed row could not be identified), a second button lets you mark it resolved after you compared the workbook with version history; the site refuses that shortcut while any check still fails.

Tell the treasurer whenever this band appears.

## Months and Annual

These tabs show what the workbook's own monthly sheets and annual summary calculate, read straight from the workbook. A small note tells you whether the site's own arithmetic agrees; if it ever says they differ, tell the treasurer.

## Signing out

Use **Sign out** on the Settings tab, especially on a shared computer. If you have unsaved changes, an unconfirmed attempt, or a paused change, the site asks first.

## What the site never does

It never keeps a separate database, never emails or copies the records anywhere, never stores your password, and never changes the workbook's formulas or layout. Everything it writes is a row in the transaction table, the sorted helper table the monthly sheets read, or the prior-year balance cell, and it checks each write by reading it back.
