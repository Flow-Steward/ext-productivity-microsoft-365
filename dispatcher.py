"""Closed service dispatcher for the Microsoft 365 extension."""

from __future__ import annotations

from collections.abc import Callable, Mapping
from typing import Any

from flowsteward_extension_sdk import runtime_context_from_payload

from calendar_ops.operations import OPERATIONS as CALENDAR_OPERATIONS
from calendar_ops.operations import QUERY_IDS as CALENDAR_QUERY_IDS
from calendar_ops.operations import handle_query as handle_calendar_query
from excel.operations import OPERATIONS as EXCEL_OPERATIONS
from excel.operations import QUERY_IDS as EXCEL_QUERY_IDS
from excel.operations import handle_query as handle_excel_query
from files.operations import OPERATIONS as FILES_OPERATIONS
from files.operations import QUERY_IDS as FILES_QUERY_IDS
from files.operations import handle_query as handle_files_query
from identity import handle_verify_identity
from m365.connection import mapping, text
from m365.errors import error_response
from m365.runtime import run_operation
from mail.operations import OPERATIONS as MAIL_OPERATIONS
from mail.operations import QUERY_IDS as MAIL_QUERY_IDS
from mail.operations import handle_query as handle_mail_query
from mail_sync.poll import handle_poll_mailbox
from teams.operations import OPERATIONS as TEAMS_OPERATIONS
from teams.operations import QUERY_IDS as TEAMS_QUERY_IDS
from teams.operations import handle_query as handle_teams_query
from notify.deliver import handle_deliver_notification, is_delivery_request

RuntimeHandler = Callable[[dict[str, Any]], dict[str, Any]]

OPERATIONS = {
    **MAIL_OPERATIONS,
    **FILES_OPERATIONS,
    **EXCEL_OPERATIONS,
    **CALENDAR_OPERATIONS,
    **TEAMS_OPERATIONS,
}
SPECIAL_ACTIONS: dict[str, RuntimeHandler] = {
    "verify_microsoft_identity": handle_verify_identity,
    "poll_mailbox": handle_poll_mailbox,
}
QUERY_HANDLERS: dict[str, RuntimeHandler] = {
    **dict.fromkeys(MAIL_QUERY_IDS, handle_mail_query),
    **dict.fromkeys(FILES_QUERY_IDS, handle_files_query),
    **dict.fromkeys(EXCEL_QUERY_IDS, handle_excel_query),
    **dict.fromkeys(CALENDAR_QUERY_IDS, handle_calendar_query),
    **dict.fromkeys(TEAMS_QUERY_IDS, handle_teams_query),
}

_REGISTRIES = (
    MAIL_OPERATIONS,
    FILES_OPERATIONS,
    EXCEL_OPERATIONS,
    CALENDAR_OPERATIONS,
    TEAMS_OPERATIONS,
)
_overlap = {
    operation_id
    for index, registry in enumerate(_REGISTRIES)
    for other in _REGISTRIES[index + 1 :]
    for operation_id in set(registry) & set(other)
}
if _overlap or set(SPECIAL_ACTIONS) & set(OPERATIONS):
    raise RuntimeError(f"duplicate operation ids: {sorted(_overlap)}")


def dispatch_runtime(payload: dict[str, Any]) -> dict[str, Any]:
    # The host envelope is checked before anything is read or sent.
    try:
        runtime_context_from_payload(payload)
    except ValueError as exc:
        response = error_response("invalid_payload", str(exc))
        response["definitely_no_external_effect"] = True
        return response
    # Core's notification delivery request carries `action: deliver_notification`
    # with `event` and `route` at the top level, not a workflow operation envelope.
    if is_delivery_request(payload):
        return handle_deliver_notification(payload)
    mode = payload.get("mode")
    if mode == "query":
        query_id = text(mapping(payload.get("query")).get("query_id"), limit=128)
        handler = QUERY_HANDLERS.get(query_id)
        if handler is None:
            return error_response("invalid_payload", "Unknown Microsoft 365 resource query")
        return handler(payload)
    if mode != "action":
        return error_response("invalid_payload", "Request mode must be action or query")
    action_id = _action_id(payload)
    special = SPECIAL_ACTIONS.get(action_id)
    if special is not None:
        return special(payload)
    return run_operation(payload, OPERATIONS)


def _action_id(payload: Mapping[str, Any]) -> str:
    action = mapping(payload.get("action"))
    return text(action.get("action_id") or action.get("operation_id"), limit=128)


__all__ = ["OPERATIONS", "QUERY_HANDLERS", "SPECIAL_ACTIONS", "dispatch_runtime"]
