# Decisions and open verifications (v0.1.0)

The task is `bmad-project/docs/plans/2026-11-22-microsoft-365-extension-task.md`.
This file records where the implementation differs from it, and which Graph or Core
behaviour has not yet been checked against a live tenant.

## Deviations from the task

1. **App-only credentials are stored per connection**, not as project provider
   secrets. `project_connection_secrets_save` attaches `secret_refs.provider_secrets`
   to the newest connection of the installation only. The task's §B4.2 was amended.
2. **`test_files_connection`.** Operation ids are extension-wide, and mail owns
   `test_connection`.
3. **`check_mail_access`** is a setup action, not a workflow operation. It is not
   effectful and only reads and drafts in the control mailbox.
4. **`list_site_libraries` needs a saved connection.** The Add-library flow is:
   save with `site_url`, run the action, fill in `drive_id`. Until then the other
   files operations return `invalid_connection`. Confirm this flow in the Core UI.
5. **`site_search_user`** points at `operation_id: list_site_libraries`. A resource
   query needs an existing operation, and there is no separate site-search operation.
6. **Send resume** deletes **every** draft carrying the step key, not only the
   abandoned ones, then builds a fresh draft. This is safe because nothing was sent.
7. **Before writing anything**, the send flow checks Sent Items for the key. If the
   email was already sent, it returns `confirmation: already_sent`.
8. **`createReply` / `createForward`** set the idempotency header as a named property
   in the internet-headers property set, because those calls take no
   `internetMessageHeaders`.
9. **`append_table_rows` with a `dedupe_conflict`** rejects the whole batch, so the
   failure truthfully reports no external effect.
10. **Excel `allow_formulas`** (default false) and **`write_range include_header`**
    are additions to the task.
11. **Free-text mail `query` cannot be combined with filters**: the combination
    returns `search_unsupported`.
12. **`save_trigger_settings` must not be used before Core A7.3.** Today
    `project_connection_ingress_patch` edits the newest connection and retypes it.
13. **`poll_mailbox`** declares `contract.serialization: {resource: mail_sync}`. Core
    requires it on every `connection_poll` action, and its lock key is per connection.
14. **Allowed upload hosts** also include `outlook.office365.com`, in addition to the
    task's `outlook.office.com`.

## To verify on the live tenant (acceptance §B8)

**Mail**
- Graph accepts `singleValueExtendedProperties` in `createReply` /
  `createReplyAll` / `createForward` bodies.
- The named internet-header property is sent as an `X-` header.
- The extended-property `$filter` works on Drafts and Sent Items, and the sent copy
  keeps it.
- The immutable id survives the move to Sent Items.
- Two preferences in one `Prefer` header are honoured.
- `contains(subject, …)` works in `$filter` together with `$orderby receivedDateTime`.
- nextLinks keep the `/me/` or `/users/{mailbox}/` path; if Graph rewrites the path
  to a GUID, cursors fail with `invalid_cursor`.
- The attachment upload-session host. Any host other than `outlook.office.com` /
  `outlook.office365.com` fails with `egress_host_not_allowed`.
- The meaning of attachment `size`.
- `DELETE` on a draft.
- A missing Archive folder returns 404 or 400.

**Files**
- Our `quickXorHash` equals Graph's.
- `@microsoft.graph.downloadUrl` comes back on an item GET.
- `Range` requests work on the download URL (otherwise files over 8 MiB fail with
  `response_too_large`).
- `?@microsoft.graph.conflictBehavior` is accepted on a simple PUT.
- `GET` / `DELETE` on an `uploadUrl` work without auth.
- Search results carry `parentReference.path`.
- `/sites/{host}:/{path}` resolves nested and team sites.
- `createLink` returns 200 for an existing link and 201 for a new one.
- When a `fail` conflict is reported during an upload session.
- Upload deadline: 240 s. A 100 MB upload on a slow link may need more.

**Excel**
- `$select` on `range(address=…)`, `usedRange(valuesOnly=true)`,
  `headerRowRange`, `dataBodyRange`.
- Listing tables per worksheet.
- Percent-encoding of worksheet and table names.
- A leading apostrophe stores text.
- `formulas` returns the formula text.
- A table with no data rows.
- Search API hit fields (`parentReference.driveId` / `siteId`).
- `rows/add` response `index`.
- The Excel request size limit, assumed to be about 4 MB.

**Trigger** (needs Core Part A): the contract is in `docs/TRIGGER_CONTRACT.md` §6
and in the task's A1–A7.
