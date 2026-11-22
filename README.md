# Microsoft 365

Microsoft 365 extension version 0.1.0 adds Outlook mail, OneDrive for Business /
SharePoint Online files and Excel workbooks to project workflows through Microsoft
Graph v1.0, plus a **New Outlook email** workflow trigger. Every connection is
project-scoped, and every workflow step selects its connection explicitly. There is
no default account.

Logo: `icon.png` (the neutral Flow Steward package mark).

## Requirements

- Flow Steward Core speaking `extension_host_v2`.
- The new-mail trigger also needs the Core changes "Part A" of the Microsoft 365
  task (background polls that keep their connection, receive a token or secrets,
  and deliver each email exactly once). On an older Core every operation works, but
  do not switch the trigger on.
- `flowsteward-extension-sdk` 0.4.0 (provided by the host).

## Two ways to connect

| | Signed-in user (delegated) | Organisation app (app-only) |
|---|---|---|
| Who | a person clicks *Connect with Microsoft* | an admin registers an app once |
| Mail | own mailbox, or a shared mailbox the user can open (`mailbox`) | any mailbox the app is allowed, by address |
| Files | own OneDrive, a SharePoint site or a drive ID | one SharePoint library per connection |
| Excel | **yes** (only this way) | no: Microsoft does not allow app-only workbook editing |
| If the person leaves | the connection stops | nothing changes |

Connection types: `m365_mail_user`, `m365_mail_app`, `m365_files_user`,
`m365_files_app`, `m365_excel_user`. Add as many of each as you need, for example
`invoices@` and `cmr@` as two `m365_mail_app` connections.

Work or school accounts only; personal Microsoft accounts are not supported.

## Setup

The **Setup guide** tab has the full steps. In short:

1. Register a **multi-tenant** app in Microsoft Entra ID. A single-tenant app fails
   with `tenant_not_multitenant` (AADSTS50194). Redirect URI = the exact Flow Steward
   callback URL. Create a client secret and note its expiry.
2. Delegated permissions: `User.Read offline_access Mail.ReadWrite Mail.Send
   Mail.ReadWrite.Shared Mail.Send.Shared Files.ReadWrite Files.ReadWrite.All
   Sites.ReadWrite.All`.
3. App-only:
   - **do not** grant the Entra application permissions `Mail.ReadWrite` or
     `Mail.Send`. They cover every mailbox and add up with Exchange restrictions;
   - grant mail through **Exchange RBAC for Applications**, scoped to the chosen
     mailboxes;
   - for files, use `Sites.Selected` with per-site grants.
4. App-only credentials (tenant ID, client ID, client secret) are entered **on each
   app-only connection**. Core's project-level provider secrets attach to only one
   connection, so they cannot be shared across mailboxes.
5. **Check read/write access** on an app-only mailbox reads and drafts in a *control
   mailbox* that must stay unreachable. It does **not** test send permission.

Never put a client secret, access token or refresh token in a workflow, log, issue or
support message.

## Operations

**Mail** (`m365_mail_user`, `m365_mail_app`):
- reads: `test_connection`, `list_mail_folders`, `search_messages`, `get_message`,
  `get_attachment` (→ artifact);
- writes: `send_message`, `reply_message`, `forward_message`, `set_message_flags`,
  `move_messages`, `delete_messages` (to Deleted Items only; already there →
  `changed: false`);
- setup action: `check_mail_access` (app-only).

**Files** (`m365_files_user`, `m365_files_app`):
- reads: `test_files_connection`, `list_drive_items`, `search_drive_files`,
  `download_file` (→ artifact), `list_site_libraries`;
- writes: `upload_file` (simple PUT up to 4 MiB, upload session above that),
  `create_folder` (idempotent), `create_share_link`.

**Excel** (`m365_excel_user`): `find_workbooks`, `list_workbook_objects`,
`read_range` (→ dataset), `write_range`, `append_table_rows`.

Every write is suppressed in Test Run.

### Resource queries

Which query serves which connection type:

| Query | Type |
|---|---|
| `mail_folders_user` | `m365_mail_user` |
| `mail_folders_app` | `m365_mail_app` |
| `drive_folders_user`, `site_libraries_user`, `site_search_user` | `m365_files_user` |
| `drive_folders_app`, `site_libraries_app` | `m365_files_app` |

The workflow step editor of Core v0.9.23 shows pickers only through
`external_resource_picker`. In v0.1 these queries serve the extension's own pages;
step fields take typed or mapped values.

### Choosing an Excel workbook

A step's `workbook` input is exactly one of:
- `{drive_id, drive_item_id}`. `find_workbooks` returns a ready `workbook` object to
  map.
- `{sharing_link}`: *Share → Copy link* in Excel. The signed-in user must already have
  access. An address-bar URL returns `workbook_link_unsupported`.

Only `.xlsx` in OneDrive for Business or a SharePoint Online library is supported.
Anything else returns `workbook_unsupported`.

## Write guarantees

A write reports success only after **positive confirmation**. If the outcome cannot
be confirmed, it returns `timeout_unknown` and is never repeated automatically.

- **Sending** (`send_message`, `reply_message`, `forward_message`) works through a
  draft. Draft and attachments are reconciled by an idempotency key and SHA-256
  before the send.
  - If adding an attachment stays unknown, the email is **not sent**. The step fails
    with `attachment_state_unknown` (retryable).
  - **Give send steps a retry policy**: the automatic step retry deletes the abandoned
    draft and sends one fresh email.
  - **Restart job** runs the **whole workflow** again with new keys. Writes from
    earlier steps (uploads, Excel rows, moves, other emails) can be repeated. Its
    one-email guarantee holds only when the first job failed before its send. If a
    draft could not be deleted, it stays in Drafts marked `abandoned`.
  - Attachment rows use the key `filename`, not `name` (Core reads `name` as a
    binding key).
  - Large outgoing attachments from **shared** mailboxes hit a known Microsoft issue.
    If it happens, the step fails with `attachment_upload_failed`; send from the
    user's own mailbox, or share a link instead.
- **`append_table_rows`** with `dedupe_column`:
  - an existing key with equal values is skipped;
  - a key with different values fails the whole batch with `dedupe_conflict`, and
    nothing is written;
  - a key repeated inside one batch is `invalid_payload`.
  - **Not protected against concurrency:** Excel has no lock, so two runs with the same
    key at the same moment can both append. Use one trigger workflow per table, with no
    parallel runs. Recommended for every trigger-driven append.
  - Values are compared after Excel's own conversion. Dates and percentages may
    compare unequal; that gives a false `timeout_unknown` or `dedupe_conflict`, never a
    false success.
  - Text starting with `= + - @` is written as text unless `allow_formulas: true`.
    This guards against formula injection from email content.
- **`write_range`** writes the values once. After an unclear result it re-reads the
  range: equal values → success, anything else → `timeout_unknown`.
- **`upload_file`** after an unclear result with `fail` or `replace`: confirmed by size
  and hash, otherwise `timeout_unknown`. With `rename` it is always `timeout_unknown`.

## New Outlook email trigger

- **Off** for every mailbox until switched on (Mail tab). Checks the chosen folder
  every 60 s (minimum 30 s).
- One workflow watches one mailbox; for two mailboxes, create two workflows. The
  trigger filters only by exact sender (`from`). Put "subject contains" in a workflow
  condition.
- The event carries `message_id` (immutable), `internet_message_id`, `from`,
  `subject`, `received_at`, `has_attachments`, `folder`, `mailbox`. The workflow reads
  body and attachments itself.
- **Covered:** crashes and restarts of Core, the worker or the extension; repeated
  pages; concurrent workers; a lost or expired delta link. Each email starts the
  workflow once.
- **Not covered:**
  - mail received before the first check after switching on, or while the trigger is
    off;
  - mail moved or deleted before a check sees it;
  - an old message imported during a sync reset.
- Contract with Core: `docs/TRIGGER_CONTRACT.md`.

## Error codes

Every error code is in `m365/errors.py::SAFE_ERROR_CODES`. The most common ones:

| Code | Meaning |
|---|---|
| `authorization_required` | The user's access expired or was revoked. Reconnect. |
| `permission_denied` | An admin must grant the missing permission. |
| `mailbox_access_denied` | Exchange RBAC does not include this mailbox. |
| `mailbox_not_found` | No such mailbox. |
| `invalid_client`, `client_secret_expired` | Fix the app credentials on the connection. |
| `tenant_not_multitenant` | Register the app as multi-tenant. |
| `rate_limited`, `provider_unavailable`, `network_failure` | Retryable. |
| `timeout_unknown` | A write could not be confirmed. Check before repeating it. |

## Development

- Contract source of truth: `contracts/parts/<service>/*.yaml`. After a change run
  `python .github/scripts/merge_contracts.py .`; `tests/test_contracts_merged.py`
  fails on drift.
- Tests use fake Graph servers only. Run them in the Flow Steward `fs-audit`
  container, not on the host.
- Design decisions and details still to confirm against a live tenant are in
  `docs/DECISIONS.md`.
