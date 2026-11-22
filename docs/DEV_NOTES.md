# Developer notes: how a service package plugs in

The authoritative task is `~/www/bmad-project/docs/plans/2026-11-22-microsoft-365-extension-task.md`
(sections B4–B8). These notes describe the shared code that already exists.

## Shared runtime (`m365/`)

- `m365/errors.py`: `M365Error(code, message, *, http_status, retry_after_seconds,
  ambiguous, retryable, failure_class)`. Only codes in `SAFE_ERROR_CODES` are allowed;
  add a new one there if you need it. `ambiguous=True` means a write may have taken
  effect, and the framework reports `timeout_unknown`. A non-ambiguous failure of a
  write reports `failed` + `definitely_no_external_effect: true`. Retry facts
  (`retryable`, `failure_class`) come from the code; pass them explicitly for
  `attachment_state_unknown` (`retryable=True, failure_class="provider"`).
- `m365/connection.py`: `resolve_connection(payload, service=...)` returns
  `ConnectionContext` (`connection_id`, `connection_type_id`, `service`, `mode`
  `user|app`, `config`, `mailbox`, `tenant_id`, `client_id`). `mailbox_base(ctx)` gives
  `/me` or `/users/{mailbox}`.
- `m365/auth.py`: `transport_for(ctx)` returns a `GraphTransport` with a delegated
  token or an app-only client-credentials token.
- `m365/transport.py`: `GraphTransport`
  - `get_json(path, query=..., headers=..., response_limit=..., purpose=...)` for
    reads, with automatic read retries.
  - `request(method, path_or_url, *, json_body | raw_body + content_type, headers,
    response_limit, retry="read"|"idempotent"|"none", purpose, authorized=True,
    ok_statuses=...)` returns `GraphResponse(status, headers, body)` with `.json()`.
  - **Retry rule (B6):** only reads and the allow-listed idempotent writes
    (`set_message_flags`, `move_messages`, `delete_messages`, `create_folder`) may
    use `retry="idempotent"`. Every other write MUST pass `retry="none"`.
  - `path_or_url` is a Graph path starting with `/` (prefixed with
    `https://graph.microsoft.com/v1.0`), an absolute `@odata.nextLink`/`deltaLink`
    (Graph host only), or, with `authorized=False`, a pre-authenticated transfer URL
    on an allowed host (`*.sharepoint.com`, `outlook.office.com`). Never log those
    URLs or return them in results.
  - Errors are mapped centrally by `map_http_error` (401 → `authorization_required`,
    403 → `permission_denied` / `mailbox_access_denied`, 404 → `item_not_found`, 409 →
    `name_conflict`, 410/SyncStateNotFound → `sync_reset`, 429 → `rate_limited`, 5xx
    on a write → ambiguous `timeout_unknown`). Re-map `item_not_found` to the specific
    `*_not_found` in your package where it helps.
  - `IMMUTABLE_ID_PREFER` is the header value; send `Prefer: IdType="ImmutableId"` on
    every mail call.
- `m365/runtime.py`:
  - Declare `Operation(operation_id, service, handler, effectful, fields)` rows.
    `fields` lists every allowed input except `connection_ref`.
  - The handler receives `OperationContext` (`input`, `connection`, `transport`
    (lazy), `test_mode`, `idempotency_key`, `payload`) and returns the `result` dict.
  - For effectful operations, `result` may carry `external_effect_status` and
    `definitely_no_external_effect`. Defaults are `succeeded` / `false`.
  - Test Run suppression of effectful operations happens before the handler runs.

## What each service package exports

`<service>/operations.py` must export:

- `OPERATIONS: dict[str, Operation]`
- `QUERY_IDS: frozenset[str]`, the resource query ids it serves
- `handle_query(payload) -> dict`, which answers `{"ok": True, "result": {"items":
  [...], "next_cursor": "", "total": n}}`. The payload has `mode: query`,
  `query.query_id`, `query.context.connection_id`, `query.search`, `query.cursor`,
  and the selected connection under `target.connection`. Follow
  `~/www/ext-google-workspace/gmail/operations.py::handle_runtime` for the shape.

## Contract fragments

Do **not** edit `extension.yaml`, `contracts/connection_types.yaml` or the
`ui/actions/actions.yaml` root. Write your fragments under `contracts/parts/<service>/`:

- `operations.yaml`: a list of operation entries in the shape of
  `~/www/ext-google-workspace/contracts/operation_manifest.yaml`. No
  `required_oauth_scopes`; `connection_type_ids` lists both types of the service
  (Excel: only `m365_excel_user`).
- `resource_queries.yaml`: list entries in the shape of
  `~/www/ext-google-workspace/contracts/resource_queries.yaml`, **one entry per
  connection type**.
- `step_ui.yaml`: list entries in the shape of
  `~/www/ext-google-workspace/contracts/step_ui_manifest.yaml`. No
  `resource_picker`; use text/json/select/number/boolean/datetime widgets.
- `actions.yaml`: list entries in the shape of
  `~/www/ext-google-workspace/ui/actions/actions.yaml`.
- `artifact_policies.yaml`: entries for operations that read or write artifacts or
  datasets.
- `external_effects.yaml`: one entry per effectful operation (`effect_kind`,
  `channel`, `idempotency: required`, `test_mode_behavior: suppress`).

The lead merges the fragments.

## Tests

- Write tests under `tests/test_<service>_*.py`, against fake openers or fake
  transports, with no network.
- **Do not run pytest on the host** (project rule). The lead runs the whole suite in
  Docker through `lint.sh --checks extension-static,extension-tests`.
- You may run `python3 -m py_compile` and `python3 -m ruff check` / `ruff format` on
  your own files.
- Use `acme`-style fixture names. Never put `flowsteward.<x>` identities in tests.
