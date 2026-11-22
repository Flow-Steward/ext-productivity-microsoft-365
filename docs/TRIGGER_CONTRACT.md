# New-mail trigger: what the extension sends to Core

This file describes the `poll_mailbox` background action (`mail_sync/`) at the
Core boundary. Background: task §B5.2 and Part A (A1, A2, A3, A5, A6) in
`bmad-project/docs/plans/2026-11-22-microsoft-365-extension-task.md`.

Where Core 0.9.23 and Part A differ, the extension follows **Part A**. "Today"
below means Core 0.9.23 (`worker_inbound_polling.py`, `response_envelope.py`,
`command_services_common.py`).

## 1. Settings (connection `ingress_config`)

The `save_trigger_settings` platform action writes these settings. The poll only reads them.

| Field | Values | Default | Meaning |
|---|---|---|---|
| `mail_trigger_mode` | `off`, `polling` | `off` (A6) | `background.mode_field`. The extension polls only when the value is exactly `polling`. Any other value, or no value, answers `disabled` and makes no Graph call. |
| `mail_poll_interval_seconds` | integer ≥ 30 | 60 | `background.interval_field`. |
| `mail_trigger_folder` | well-known name (`inbox`) or folder id | `inbox` | The folder to watch. |

Suggested `ingress_schema` for both mail connection types. The lead owns
`contracts/connection_types.yaml`. Quote `'off'`, because YAML 1.1 reads a bare
`off` as `false`.

```yaml
ingress_schema:
  type: object
  properties:
    mail_trigger_mode: {type: string, enum: ['off', polling], default: 'off'}
    mail_poll_interval_seconds: {type: integer, minimum: 30, default: 60}
    mail_trigger_folder: {type: string, default: inbox, maxLength: 512}
    last_inbound_sync_at: {type: string}
    mail_sync: {type: object}
```

The runtime owns `last_inbound_sync_at` and `mail_sync`. The settings action never
writes them.

## 2. Inputs the poll reads

The poll reads its inputs from **`action.target.connection`**, or from
`target.connection`. This is the object operations already receive (A2: "pass the
connection's resolved secrets to the action as operations do"):

```json
{
  "connection_id": "…",
  "connection_type_id": "m365_user | m365_mail_app",
  "config": {"mailbox": "…", "tenant_id": "…", "client_id": "…"},
  "secrets": {"access_token": "…"} | {"client_secret": "…"},
  "ingress_config": {"mail_trigger_mode": "polling", "mail_sync": {…}, …},
  "state_version": 17,
  "config_generation": 4
}
```

- `state_version` and `config_generation` may be integers or strings. The extension
  echoes them **unchanged** and never increments them.
- **Fallback:** if `ingress_config`, `state_version` or `config_generation` is missing
  from `target.connection`, the poll reads it from
  `runtime_context.resources.project_connection`. That is the row Core hydrates
  today through `runtime_resources.project_connection`. The fallback applies only
  when that row has the same `connection_id`.
- **Today:** the worker passes only `target: {"connection_id": …}`. The poll then
  fails with `invalid_connection` ("A Microsoft 365 connection must be selected").
  This is deliberate: without A2 there is no token and no secret.

## 3. What the poll returns

### 3.1 Success with progress

The poll read at least one page, or recorded a reset:

```json
{
  "ok": true,
  "message": "Read 3 page(s); 2 new email(s).",
  "result": {
    "status": "polled | partial",
    "phase": "baseline | catch_up | live | recovery",
    "pages": 3,
    "processed_updates": 2,
    "events_emitted": 2,
    "caught_up": true,
    "sync_reset": false,
    "sync_started_at": "2026-11-22T10:00:00Z",
    "last_reset_at": "",
    "last_recovered_at": "",
    "error_code": ""
  },
  "events": [ <event>, … ],
  "host_mutations": [
    {
      "kind": "project_connection_upsert",
      "payload": {
        "connection_id": "<polled connection>",
        "connection_type": "<its own type, unchanged>",
        "ingress_config": { <ingress_config as read>, "last_inbound_sync_at": "<now>", "mail_sync": <state> },
        "expected_state_version": <as received>,
        "expected_config_generation": <as received>
      }
    }
  ]
}
```

- `status: "partial"` means an error stopped the poll after some progress, for
  example a 429 on the third page. The events and the state up to the last complete
  page are still returned, and `result.error_code` names the error. The same code is
  stored in `mail_sync.last_error_code`.
- `result` and the events never contain Graph links. The links exist only in
  `ingress_config.mail_sync`.

### 3.2 Idle

There are no events and no host mutation:

```json
{"ok": true, "message": "…", "result": {"status": "disabled | deferred", "processed_updates": 0, "events_emitted": 0}, "events": []}
```

- `disabled`: the mode is not `polling`.
- `deferred`: the 20 s budget ran out before the first page. This happens only when
  the app-only token request itself was slow.

### 3.3 Failure before any progress

The poll returns the shared error envelope (`m365/errors.py`). It has no events and
no host mutation:

```json
{"ok": false, "result": {}, "error_code": "permission_denied", "error": "…", "errors": [...], "retryable": false, …}
```

The codes are listed in `contracts/parts/mail_sync/actions.yaml`. The ones specific
to the trigger are `folder_not_found` (the folder was deleted) and `sync_reset`
(Graph reset the sync state repeatedly within one poll).

### 3.4 Event

One event per new email:

```json
{
  "event_type": "microsoft365.mail.message.received",
  "source_id": "m365:<connection_id>:<immutable message id>",
  "correlation_id": "<same as source_id>",
  "source": "microsoft365.mail",
  "severity": "info",
  "summary": "New email from billing@acme.test in cmr@acme.test",
  "payload": {
    "connection_id": "…",
    "mailbox": "cmr@acme.test",
    "message_id": "<immutable id>",
    "internet_message_id": "<…@…>",
    "from": "billing@acme.test",
    "subject": "…",
    "received_at": "2026-11-22T10:01:00Z",
    "has_attachments": true,
    "folder": "inbox"
  }
}
```

- `from` is the sender **address in lower case**. Core matches trigger config by
  exact string equality, so the trigger's `from` field must be lower case, and its
  help text says so.
- For a delegated connection without `mailbox`, `mailbox` holds the user's address.
  The poll reads it once with `GET /me`, when a sync starts.
- `connection_id` is in the payload so that the trigger config `connection_id`
  matches it. One workflow watches one mailbox.

### 3.5 State (`ingress_config.mail_sync`)

```json
{
  "schema": 1,
  "phase": "baseline | catch_up | live | recovery",
  "sync_started_at": "…Z", "watermark": "…Z", "query_from": "…Z",
  "delta_next_link": "", "delta_link": "", "query_next_link": "",
  "recovery_started_at": "", "last_reset_at": "", "last_recovered_at": "", "reset_count": 0,
  "last_error_code": "", "last_error_at": "",
  "config_generation": "4", "folder": "inbox", "mailbox": "cmr@acme.test", "mailbox_address": "cmr@acme.test"
}
```

For the Mail page, use `phase`, `last_inbound_sync_at` (the last poll),
`last_reset_at` and `last_recovered_at` ("sync was reset and recovered"), and
`last_error_code`.

## 4. Algorithm summary (`mail_sync/machine.py`)

- **First request.** The baseline starts with
  `GET …/mailFolders/{folder}/messages/delta?changeType=created&$select=id,internetMessageId,from,subject,receivedDateTime,hasAttachments`.
  Graph keeps `changeType` and `$select` in every `nextLink` and `deltaLink`. The
  poll follows those links verbatim.
- **Headers.** Every call sends `Prefer: IdType="ImmutableId", odata.maxpagesize=100`.
- **Catch-up and recovery query.** The query is
  `GET …/messages?$filter=receivedDateTime ge {query_from}&$orderby=receivedDateTime asc&$top=100`.
  It is ascending, so mail that arrives while the poll is paging lands after the
  pages already read.
- **Delivery filter.** Only mail with `receivedDateTime ≥ sync_started_at` is
  emitted, in every phase. `sync_started_at` is the poll's UTC time, rounded down to
  the second, and is set when a sync starts.
- **When a sync starts.** A new sync begins when there is no state, or when the
  stored state's `config_generation`, folder or mailbox differs from the current
  ones.
- **Recovery.** A 410, `SyncStateNotFound`, or a stored link that Graph rejects
  (400, or a link that is not a Graph link) starts a recovery, as does a phase that
  needs a link it does not have.
  - The recovery query starts at `query_from = max(sync_started_at, watermark - 15 min) - 5 min`.
  - A new baseline follows. Its catch-up starts at `recovery_started_at - 5 min`.
  - Within one poll there are at most two resets.
  - A 404 is `folder_not_found`, not a reset.
- **Budget.**
  - Each poll has 20 s. It starts a page only while at least 3 s are left, and
    starts no new page after 200 events.
  - The Graph transport deadline is the remaining budget, with a minimum of 5 s.
  - The poll returns the state as it was right after the last complete page.
    That is the "progress after each page" checkpoint. Core commits it in one
    transaction together with all the events read up to it.

## 5. What the extension relies on from Core (Part A)

1. **A1, per-connection write.** Apply the host mutation to the polled connection
   (`payload.connection_id`, which is the worker's `connection_id`), never to the
   newest connection of the installation.
2. **A1, compare-and-set.** Under the connection row lock, compare
   `expected_config_generation` and `expected_state_version` with the row.
   - Same generation, same version: commit the events, the dispatch records and the
     state, and bump `state_version`.
   - Same generation, newer version: commit the events and drop the state.
   - Different generation, or the connection is disabled or switched off: drop both
     the events and the state.
   - `expected_*` is `null` only when Core sent no versions. That happens today,
     not after Part A.
3. **A1, generation bumps.** Bump `config_generation` when `mail_trigger_mode`
   changes to or from `polling`, when `mail_trigger_folder` changes, and when the
   mailbox changes, the connection is reconnected, or it is disabled.
   - **Switching off clears the state** only through this bump. The extension
     starts a new sync, with a new `sync_started_at`, as soon as the stored state's
     generation differs from the current one.
   - Without the bump, switching off and on again would continue the old delta
     link, and mail received while the trigger was off would be delivered.
4. **A2.** Call `poll_mailbox` only for the action's `connection_type_ids`
   (`m365_user`, `m365_mail_app`). Pass the refreshed `access_token` (`_user`)
   or the connection's own `client_secret` (`_app`) in `target.connection.secrets`,
   together with `ingress_config` and the two versions. The action declares no
   `provider_secrets`.
5. **A3.** The event is unique on (`project_id`, `provider_key`, `source_id`). A
   replayed page or a repeat within one result writes no second event. The
   extension also drops repeats within one poll.
6. **A5.** Events from a background result must reach
   `start_provider_event_workflows` with this `source_id`.
   - Today the worker calls `_process_extension_response_envelope` without
     `event_command_service_factory`. The events therefore go through
     `emit_extension_event` and never through `ingest_provider_events`.
   - `supported_events` lists `microsoft365.mail.message.received`
     (`contracts/parts/mail_sync/supported_events.yaml`).
7. **A6.** Write `mail_trigger_mode: off` into `ingress_config` when a mail
   connection is created. Today Core reads a missing mode as `auto`; the extension
   still refuses to poll in that case.
8. **Serialization.** `contract.serialization.resource: mail_sync`. The lock key
   includes the polled `connection_id` (`extension_operation_lock._serialization_key`),
   so one mailbox is never polled twice at once, while different mailboxes poll in
   parallel. Today Core requires `serialization` on every `connection_poll` action.

## 6. Differences from Core's current shape (deviations)

| # | Extension sends or declares | Core today | Needed from Core |
|---|---|---|---|
| 1 | Reads `target.connection` with secrets, `ingress_config` and versions | `target = {connection_id}` | A2 |
| 2 | `expected_state_version`, `expected_config_generation` in the `project_connection_upsert` payload | Ignored; the newest connection is written (`connection_id` is not in the worker's `audit`) | A1 |
| 3 | Action-level `connection_type_ids` on `poll_mailbox` | Not read; every connection of the install is polled | A2 |
| 4 | `ingress_config` as a full copy (settings as read, plus the new `mail_sync` and `last_inbound_sync_at`) | Replaces `ingress_config` whole | See note (a) |
| 5 | `save_trigger_settings` uses `config.connection_type_ids` (two types) | `project_connection_ingress_patch` takes one `connection_type_id`, patches the newest connection and **overwrites its type** with that one | See note (b) |
| 6 | Failure without progress returns `ok: false` with no state | The interval is measured from `last_inbound_sync_at`, so a failing connection is retried on every 15 s scan | See note (c) |

Notes:

- **(a) Lost interval change.** If the interval changes while a poll is running,
  the poll's copy writes back the old interval. The interval does not bump the
  generation, so the compare-and-set does not catch this. Recommendation: for
  `poll_mailbox`, Core applies only the runtime-owned keys `mail_sync` and
  `last_inbound_sync_at` from the result, and keeps the declared settings fields
  of the current row.
- **(b) The settings save must change before release.** It must:
  - address the selected connection (`target.connection_id`);
  - keep that connection's own type;
  - validate against that type's `ingress_schema`;
  - bump `config_generation` under the A1 row lock (point 3 of section 5).

  Until then, do not ship `save_trigger_settings`. With two mail connections, it
  would change the wrong one, and with a single declared type it would retype an
  `_app` connection.
- **(c) Retries after failure.** Recommendation: Core records the time of a failed
  poll too, and applies the interval to it.

## 7. Guarantee boundaries (README and help text)

**Covered:**

- crashes and restarts of Core, the worker or the extension;
- repeated pages and concurrent workers;
- a lost or expired delta link, in any phase.

**Not covered:**

- mail received before `sync_started_at`;
- mail moved out of the folder or deleted before a poll sees it;
- mail whose `receivedDateTime` is older than the recovery fetch window;
- time while the trigger is switched off.
