# Correction refused with "The target session is invalid." (2026-11-03) — investigation and local fix

Read-only investigation of the first live correction attempt on the test copy, then a local, unpublished fix
with tests. No OneDrive write, no publish, no cleanup. The test row (#71, April 30) is left exactly as the
browser left it; its correction dialog is still open on Cody's screen.

Vocabulary as in the handoff: **verified** = from code, a file or a log on this machine; **reported** = Cody's
statement; **inferred** = the best-supported reading of reported facts, not proven; **not established** = cannot
be settled from here.

## 1. What happened (reported)

1. Baseline 63 transactions. Adding test row #71 succeeded; the count became 64.
2. Correcting #71 (amount and description changed) failed. The dialog said
   `Could not save: The target session is invalid.`
3. Pressing **Check workbook** once then showed the message that begins "Could not save: the answer from
   Microsoft was lost, so this may or may not be in the workbook."
4. No retry, no delete. Nobody else was editing (reported).

There is **no request log for this attempt**: before this fix `ledger.html` constructed its `ExcelClient`
without a `log` function, so unlike `diagnostics.html` it recorded nothing to copy. The status, code and
inner codes of the live refusal were therefore **not captured**. Everything below about the live envelope and
the stage at which it struck is **inferred**.

## 2. The envelope Microsoft sends for this message (verified shape; live envelope inferred)

Microsoft's own pages do not contain the sentence. A Microsoft Q&A question from 2021-08-24
(<https://learn.microsoft.com/en-us/answers/questions/525599/is-there-a-way-to-unpause-a-microsoft-graph-excel>)
quotes the full envelope for exactly this message, **three levels deep**:

```
HTTP 400 BadRequest
code: InvalidSession            message: The target session is invalid.
  innerError code: invalidSessionReCreatable
                   message: The session specified in the request does not exist or is invalid due to a transient error.
    innerError code: InvalidOrTimedOutSession
                     message: We noticed that you haven't been interacting with this workbook, so we paused your session.
```

That this is what the service sent on 2026-11-03 is inferred from the identical top-level message; the live
inner codes were not recorded. Microsoft's error-handling page
(<https://learn.microsoft.com/en-us/graph/workbook-error-handling>) says to act on the **second-level** code,
that codes are case-insensitive, that message strings are diagnostic only, and for
`invalidSessionReCreatable`: "The Microsoft Graph client can try to recreate a session and resume the work.
Further access to the session specified in the failed request is not expected." No Microsoft page states the
inactivity timeout; the Q&A asker cites "about 7 minutes" for non-persistent sessions, unsourced.

## 3. The classifier defect (verified from code)

Before the fix, `classifyError` in `site/js/workbook/excel-client.js` walked to the **deepest**
`innerError.code` and tested only that:

```js
while (inner) { if (inner.code) innerCode = inner.code; inner = inner.innerError; }   // -> 'InvalidOrTimedOutSession'
const sessionInvalid = /^invalidSession/i.test(innerCode);                            // false
const recreatable  = /^invalidSessionReCreatable$/i.test(innerCode);                  // false
```

Fed the three-level envelope above, both flags were **false**, `retryableRead` was false, and the top-level
`InvalidSession` code was never consulted (the new unit test "classifyError: the paused-session envelope …"
fails against the old code with exactly that: `sessionInvalid` false). Consequences in `_send`:

- the dead session id was **kept** (`this.sessionId = null` ran only when `c.recreatable`);
- a GET was **not** retried on a new session;
- the error was thrown with `sessionInvalid: false`, `ambiguous: false` (status 400 < 500).

Every later request of that `ExcelClient` carried the same paused session id, so it failed the same way until
the page was reloaded (a reload constructs a new client). The existing unit tests passed because the mock
emitted a two-level body whose deepest code is `invalidSessionReCreatable`.

## 4. Why the two messages appeared (mechanism verified from code; stage inferred)

Correction path: `updateTransaction` in `ledger-workbook.js` **reads LOG first** (`_readLog`, a GET), then
POSTs the corrected copy, then reads again, formats, checked-deletes the old copy, rebuilds LOG_Sorted and
re-syncs the month rows. A paused-session refusal hits the first request that uses the session after the
pause. If the envelope was the one in §2, that was most likely the pre-read GET, in which case the POST was
never sent and nothing was written. This is **inferred**: with no log, the alternative (the POST landed and
the pause hit a later read in `_finishCorrection`, leaving the corrected copy and the old copy both in LOG)
is not excluded, and both produce the same two messages. A read on a fresh session settles which occurred.

Either way the refusal propagated to `SaveController.submit`: not a conflict, `ambiguous` false, so the state
text was "Could not save: The target session is invalid." and the pending operation was kept.

**Check workbook** called `inspectCorrection`, whose first step is another `_readLog` GET on the same client,
still carrying the paused id. It was refused identically. `SaveController.check` treated *any* exception from
`inspect` as `needsVerification = true`, and `saveStateText` in `app.js` rendered that flag as "the answer from
Microsoft was lost". So the second message described the check's own failure, not a lost write.

## 5. Defects found

1. **Verified, root cause:** `classifyError` classified by the deepest inner code; Microsoft's documented
   recreatable error, when wrapped in a third level, became an unknown 400.
2. **Verified:** the session id was dropped only for `recreatable`; for every other `invalidSession*` code
   (Microsoft: "further access to the session … is not expected") it was kept and reused.
3. **Verified:** `check()` conflated "the check failed to read" with "the write's answer was lost", and the
   page wording followed.
4. **Verified:** no request log on the ledger page, so the live refusal's codes were not recorded.

## 6. Fix applied locally (unpublished; nothing written to OneDrive)

- `excel-client.js` `classifyError`: `innerCode` is the **second-level** code (Microsoft's rule; the existing
  test `innerCode === 'invalidSessionReCreatable'` still holds); `innerCodes` is the whole chain for the log;
  `sessionInvalid` looks at every level and the top-level code; `recreatable` means some level equals
  `invalidSessionReCreatable`.
- Automatic recovery **only where Microsoft permits it**: on `invalidSessionReCreatable` the session is dropped
  and a GET is retried once on a new session; a write is never retried and is reported as refused (definite,
  not ambiguous). Any other `invalidSession*` code **blocks the client**: the session is dropped, `blocked`
  records code, inner codes, message and request id, and every later request (single, batch, session
  creation) is refused locally without being sent, until `clearBlock()`. The pages never call `clearBlock()`;
  a reload constructs a new client once the member has dealt with the cause.
- `save-controller.js`: a failed check sets `checkFailed`; the pending attempt and its marker stay.
- `app.js` wording for that state: "Could not check the workbook: <reason> Your earlier save attempt remains
  unconfirmed; the check wrote nothing. Press "Check workbook" to read again." (or "Reload the page, then …"
  when the client is blocked). The "answer from Microsoft was lost" sentence is now only for ambiguous writes.
- Request log on the ledger page: the client is created with a `log` function; entries hold time, method,
  path, status, timing, request id and error codes; "Copy request log" on the Settings tab copies them. The
  session-opened event no longer carries the session id (the diagnostics page's log line changed with it).
  Tokens, session ids and record contents are never logged.
- Review round 2: session creation was absent from the log (a reproduced `createSession` failure left it
  empty). Now every creation attempt, the 202 of an asynchronous creation, each poll with its outcome and the
  operation's error codes when it fails, and the final session read are logged with status, timing, request id
  and codes; the operation key and the session resource key are replaced by `…`. A batch refused as a whole keeps
  its codes, and each sub-request is logged on its own line with its status and codes. The log is also shown on
  the ledger page itself ("Show request log": readonly text, "Download request log" as a text file); when the
  clipboard rejects or is missing, "Copy request log" shows the same entries there, selected, instead of
  pointing at Diagnostics (whose log is that page's own).
- Not changed: keep-alive/`refreshSession` timers. After the fix every write path begins with a GET that
  recreates a paused session, so a correction after idle time proceeds without a visible error.

## 7. Tests added

`tests/unit/session-errors.test.js` (11, all failed before the fix except the guard for a non-recreatable
inner code) and `tests/e2e/session.spec.js` (4, run on desktop Chrome, iPhone-size Chromium and WebKit):

- the three-level envelope classifies as recreatable by the second-level code, with the chain preserved;
- paused session on a read → recreated once, read retried once; on a write → one attempt, no replay, session
  dropped, next read opens a new session;
- a non-recreatable session error blocks the client: GET, POST, PATCH, batch, `openSession`, `calculate` are
  all refused without any request reaching the service, until `clearBlock()`; the same from inside a batch;
- the log never carries a session id or token and records the inner-code chain;
- a session paused **before** a correction: recovered on the correction's initial read, correction succeeds,
  written once, one copy;
- the paused envelope refusing the **append itself**: nothing written, nothing replayed, Check reads on a new
  session and reports "not in the workbook yet", Save then writes once;
- a blocking session error **after the corrected copy was appended**: the attempt stays pending with its marker
  and payload, the blocked client neither checks nor writes, a fresh client's check reports "incomplete" and
  Save finishes it with exactly one append in the whole sequence;
- the controller reports a failed check as such, keeps the attempt pending, and clears the flag once a later
  check answers;
- browser: the three correction scenarios with the real wording, plus the request log's content and the copy
  control.

Results on 2026-11-03: unit 169 passed (158 before), browser 201 passed.

Round 2 added `tests/unit/request-log.test.js` (8: refused createSession; network error on both attempts; 504
then success; asynchronous creation's 202, polls and final read without the operation key; creation reported
failed with its codes; a refused poll; a batch refused as a whole; per-sub-request lines) and
`tests/e2e/request-log.spec.js` (4 × 3 browsers: clipboard rejecting, clipboard absent, Show/Hide kept in step
with later requests, downloaded file equal to the page text). Every unit test failed before the change. Results:
unit 177 passed, browser 213 passed.

## 8. What settles today's outcome (read-only, for Cody's decision)

Reloading `ledger.html` constructs a new client and a new session. The persisted correction draft comes back
as a pending operation and the page runs **Check workbook automatically, reading only**. Expected results:
"This is not in the workbook yet. Press Save to send it." (nothing was written) or an incomplete/duplicate
report (the POST landed). Either way the row and the dialog's text are preserved; nothing is written until
Save or Delete is pressed. The test row is not cleaned up by this. Note that the reload will run the published
site, not this local fix, until the fix is reviewed and published.
