# Decisions and open verifications

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

## v0.2.0

The task is `bmad-project/docs/plans/2026-11-22-microsoft-365-extension-v0.2-task.md`
(Part D; Core Part C is built separately on the Core release branch). The full
deviation and live-verify lists are in `docs/v0.2-calendar-notes.md` and
`docs/v0.2-teams-notes.md`.

### Decisions taken during the build

1. **`User.ReadBasic.All`** is added to the Teams scopes. @mentions resolve addresses
   with `GET /users/{address}`; it needs no admin consent. The task was amended.
2. **`event.run_url`** (task C3.4) was added to Core Part C during the build: channel
   providers had no way to link to the run. Without it, Teams messages and cards have
   no link.
3. **Every Teams post is HTML.** A text body is escaped, so the hidden marker and
   mentions work for both kinds. File links are plain links, not Graph attachments.
4. **The Teams channel picker** lists the channels of every joined team (first 50
   teams), as `team_id/channel_id`: Core passes no dependent field to a resource query.
5. **Calendar re-run guard.** `create_event` looks for its key before posting, so a
   re-run returns `already_exists` instead of a second event (one extra GET per create).
6. **An appointment without attendees is deleted** by `cancel_event`; `cancel` is a
   meeting action.
7. **After an ambiguous cancel**, an event that still reads as active is
   `timeout_unknown`: Graph's event has no parent folder to detect Deleted Items.
8. **Delivery without Core's `idempotency_key`** fails with `idempotency_key_missing`
   before any write.
9. **`calendar_ops/`** is the package name; `calendar` would shadow the standard library.

### To verify on the live tenant (D7)

**Calendar**
- `GET /events/{id}` after the organizer cancels (404, or `isCancelled`).
- A repeated POST with the same `transactionId`.
- Whether the extended-property filter combines with a date filter.
- IANA zone names in `dateTimeTimeZone.timeZone` on create and update.
- Whether the RBAC scope limits whose schedules `getSchedule` can read app-only.
- `allowedOnlineMeetingProviders` on calendars with and without Teams.

**Teams**
- Whether the HTML-comment marker survives in stored messages (else switch
  `teams.markers.DEFAULT_MARKER` to the zero-width marker).
- Channel message order in `list_messages`.
- Creating a 1:1 chat with an alias or an unknown address.
- Graph's answer to an oversized message.
- The real Workflows webhook host and status codes; that the card appears.

**Notifications and upgrade**
- One event through a mail route and a Teams route of this provider; no duplicates on
  re-delivery.
- System email (reset, welcome, email change) from `noreply@`.
- Upgrading an installed v0.1 (kind `tool_provider`) to v0.2: connections, OAuth
  grants and the mail-trigger cursor survive.
