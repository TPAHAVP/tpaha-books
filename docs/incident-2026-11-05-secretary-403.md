# Secretary refused with `403 accessDenied` on the pinned workbook (2026-11-05) — evidence and local fix

Reported: the secretary's TPAHA account can open the real `TPAHA_2026.xlsx` in Excel through Chrome,
but the published ledger page shows the access-error card with `HTTP 403 · code accessDenied ·
request-id 39e4acf2-12c2-4049-85ef-9abc6f3ae2bf`. The workbook lives in the Vice President's OneDrive.
Nothing was written; the page stops at the first read. No workbook, Entra permission or consent was changed
in this investigation and nothing is published.

## 1. What the page asks Microsoft for (verified from code)

- `site/js/auth.js` requested the delegated scopes `Files.ReadWrite` and `User.Read` (the registration of
  2026-09-11 carries the same two, with admin consent).
- `site/js/workbook/excel-client.js` addresses the pinned workbook as
  `/drives/{driveId}/items/{itemId}` (`this.itemPath`), never `/me/drive/...`. The drive id is the
  Vice President's OneDrive. The first request the page cannot do without is
  `POST .../workbook/createSession`; the metadata read before it is allowed to fail silently
  (`ledger-workbook.js`, `getItemMeta().catch(() => ({}))`).
- So for every member except the workbook's owner, the page reads a file in **another user's drive** with a
  token whose only Files scope is `Files.ReadWrite`.

## 2. What Microsoft documents (verified 2026-11-05 against learn.microsoft.com)

| Delegated permission | Microsoft's description | Admin consent required |
|---|---|---|
| `Files.ReadWrite` | "Allows the app to read, create, update, and delete the signed-in user's files." | No |
| `Files.ReadWrite.All` | "Allows the app to read, create, update, and delete all files the signed-in user can access." | No |
| `Files.Read.All` | "Allows the app to read all files the signed-in user can access." | No |

Source: [OneDrive permission scopes](https://learn.microsoft.com/en-us/onedrive/developer/rest-api/concepts/permissions_reference?view=odsp-graph-online)
and the [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference).

The same OneDrive page adds, verbatim: *"Note that for personal accounts, Files.Read and Files.ReadWrite also
grant access to files shared with the signed-in user."* The note is limited to personal accounts; TPAHA
accounts are work accounts. Its example usage pairs `Files.ReadWrite` with the member's own drive
(`PUT /me/drive/root/children/...`) and `Files.ReadWrite.All` with "Write a file shared with the user"
(`PUT /users/rgregg@contoso.com/drive/root/children/...`), that is, another user's drive.

The [sharedWithMe reference](https://learn.microsoft.com/en-us/graph/api/drive-sharedwithme?view=graph-rest-1.0)
states, verbatim: *"You can't access shared items returned from this API if the request doesn't contain one
of the `*.All` permissions."* (Its least-privileged delegated permission is `Files.Read.All`.)

The [createSession reference](https://learn.microsoft.com/en-us/graph/api/workbook-createsession?view=graph-rest-1.0)
lists `Files.ReadWrite` as least privileged, but documents only `/me/drive/...` request paths, the signed-in
user's own drive. The [driveItem get reference](https://learn.microsoft.com/en-us/graph/api/driveitem-get?view=graph-rest-1.0)
documents `/drives/{drive-id}/items/{item-id}` and lists `Files.ReadWrite.All` among its permissions.

A matching report, unofficial: [microsoft/entrabot#135](https://github.com/microsoft/entrabot/issues/135)
(2026-10-31) describes a token with `Files.ReadWrite` but no `Files.Read.All` getting `403 accessDenied` on
files in another user's OneDrive that were shared with the signed-in identity; the suggested remedy is adding
a delegated `*.All` scope.

## 3. Conclusion (evidence-based, to be confirmed by the permission change)

The 403 is consistent with the token scope, not with sharing: the secretary can open the file in Excel, which
uses the member's SharePoint permission directly, while the page's token carries a scope Microsoft defines as
the signed-in user's own files. The owner of the drive (the Vice President) is the only member for whom
`Files.ReadWrite` reaches this workbook, which is why the earlier tests from that account passed.

Confirmation is one step away and is Cody's to take: add the delegated permission and consent (section 5),
then publish the matching site change. A 403 that persists after both would point at sharing or Conditional
Access instead, and the request id above (plus the new one) is what to quote to Microsoft.

What was wrong in the page's wording: the 403 text asserted *"Your Microsoft account does not have access to
this workbook, or the access has been withdrawn. Ask the treasurer to share ..."*. That is false here and would
have sent the secretary to the treasurer for a sharing change that would not have helped.

## 4. Local changes (unpublished, for review)

- `site/js/auth.js`: `SCOPES = ['Files.ReadWrite.All', 'User.Read']`. Still delegated, still the single-tenant
  authority, still redirect flow; "Assignment required" and the account restrictions are untouched. The pin
  (`TPAHA_2026.xlsx`, drive id, item id) is unchanged.
- `site/js/workbook/locate.js` `describeAccessFailure`: a 403 now reads *"Microsoft 365 refused this page's
  request to open TPAHA_2026.xlsx (access denied). That happens when the permission the site signs in with does
  not reach a workbook kept in another member's OneDrive, even if you can open the same file in Excel, or when
  the workbook is not shared with your account. Nothing was changed. Tell the treasurer and quote the details
  below."* The details line (`HTTP 403 · code accessDenied · request-id …`) is unchanged.
- `tools/verify-config.mjs`: expects `Files.ReadWrite.All` + `User.Read`; two new checks (requests
  `Files.ReadWrite.All`; does not also request the narrower `Files.ReadWrite` or `Files.Read`). 32 checks.
- `site/js/config.js`: `appVersion` 2026.11.05 so the service worker fetches the new `auth.js` on phones
  (scripts are cache-first); registration comment updated. `docs/setup-guide.md` step B6 and the permission
  section rewritten for `Files.ReadWrite.All`.
- Tests, written first and watched fail: `tests/unit/pinned-workbook.test.js` (the 403 wording must name the
  refusal, the sign-in permission, Excel and sharing, and must not say "does not have access" or "Ask the
  treasurer to share"); `tests/e2e/pinned-workbook.spec.js` (the access card must not carry the old assertion).

## 5. The exact permission change needed in Entra (not done; Cody's decision)

Entra admin centre → App registrations → **TPAHA Books** → **API permissions**:

1. **Add a permission → Microsoft Graph → Delegated permissions → `Files.ReadWrite.All`** ("Have full access to
   all files user can access") → Add permissions.
2. **Grant admin consent for TPAHA**. Delegated `Files.ReadWrite.All` does not itself require admin consent,
   but the tenant's consent policy may, and the granted consent is what lets every member's token renew
   silently without a prompt.
3. Optionally, after the site is published with the new scope: **remove `Files.ReadWrite`** so the registration
   lists exactly what the site requests. Keep `User.Read`.

Nothing else: no application permission, no client secret, no change to "Assignment required", no change to the
workbook's sharing.

Order matters. Entra first, then publish: with the permission consented, the current site keeps working
(its old scope is still consented) and the new site obtains a `Files.ReadWrite.All` token silently through the
refresh token. Published first, the new site would send every member to a consent prompt
(`AADSTS65001`, "Need admin approval" for non-admins) until consent is granted.

After publishing: a device that already signed in holds a token for the old scope; MSAL finds no cached
access token for the new scope, renews silently, and if that fails the page redirects once to sign in again.
The secretary should then see the ledger; if not, the Settings tab's request log carries the new request id.
