"""Outlook calendar operation registry and the calendar picker queries (task D2)."""

from __future__ import annotations

from typing import Any

from m365.auth import transport_for
from m365.connection import mapping, text
from m365.errors import M365Error, error_response
from m365.runtime import Operation, TransportFactory, run_operation

from .access_check import check_calendar_access
from .reads import (
    GET_EVENT_FIELDS,
    LIST_EVENTS_FIELDS,
    calendar_connection_test,
    calendar_items,
    get_event,
    list_calendars,
    list_events,
)
from .schedule import FREE_TIME_FIELDS, find_free_time
from .writes import (
    CANCEL_FIELDS,
    CREATE_FIELDS,
    RESPOND_FIELDS,
    UPDATE_FIELDS,
    cancel_event,
    create_event,
    respond_to_event,
    update_event,
)

SERVICE = "calendar"


def _op(
    operation_id: str,
    handler: Any,
    fields: set[str] | frozenset[str] = frozenset(),
    *,
    effectful: bool = False,
) -> Operation:
    return Operation(operation_id, SERVICE, handler, effectful=effectful, fields=frozenset(fields))


OPERATIONS: dict[str, Operation] = {
    row.operation_id: row
    for row in (
        _op("test_calendar_connection", calendar_connection_test),
        _op("list_calendars", list_calendars),
        _op("list_events", list_events, LIST_EVENTS_FIELDS),
        _op("get_event", get_event, GET_EVENT_FIELDS),
        _op("create_event", create_event, CREATE_FIELDS, effectful=True),
        _op("update_event", update_event, UPDATE_FIELDS, effectful=True),
        _op("cancel_event", cancel_event, CANCEL_FIELDS, effectful=True),
        _op("respond_to_event", respond_to_event, RESPOND_FIELDS, effectful=True),
        _op("find_free_time", find_free_time, FREE_TIME_FIELDS),
        # Setup action on the Calendar page: two reads, never a write. Not a workflow
        # operation.
        _op("check_calendar_access", check_calendar_access, {"control_mailbox"}),
    )
}

#: Resource query id → the one connection type it serves.
CALENDAR_QUERIES = {"calendars_user": "m365_user", "calendars_app": "m365_calendar_app"}
QUERY_IDS = frozenset(CALENDAR_QUERIES)


def handle_query(
    payload: dict[str, Any], *, transport_factory: TransportFactory = transport_for
) -> dict[str, Any]:
    try:
        query = mapping(payload.get("query"))
        query_id = text(query.get("query_id"), limit=128)
        expected_type = CALENDAR_QUERIES.get(query_id)
        if expected_type is None:
            raise M365Error("invalid_payload", "Unknown Outlook calendar query")
        connection_id = text(mapping(query.get("context")).get("connection_id"), limit=256)
        target = mapping(payload.get("target"))
        connection = mapping(target.get("connection"))
        if not connection_id or not connection:
            raise M365Error(
                "invalid_connection", "The calendar picker needs a selected calendar connection"
            )
        type_id = text(
            connection.get("connection_type_id") or connection.get("connection_type"), limit=128
        )
        if type_id != expected_type:
            raise M365Error(
                "invalid_connection", f"{query_id} serves {expected_type} connections only"
            )
        search = text(query.get("search"), limit=256).lower()
        action_payload = {
            **payload,
            "mode": "action",
            "action": {
                "action_id": "list_calendars",
                "input": {"connection_ref": connection_id},
                "target": target,
            },
        }
        response = run_operation(action_payload, OPERATIONS, transport_factory=transport_factory)
        if response.get("ok") is not True:
            return response
        items = calendar_items(mapping(response.get("result")))
        if search:
            items = [row for row in items if search in row["label"].lower()]
        return {"ok": True, "result": {"items": items, "next_cursor": "", "total": len(items)}}
    except M365Error as exc:
        response = error_response(exc.code, exc.message)
        response.update(exc.retry_facts())
        return response
    except Exception:
        return error_response("internal_error", "The Outlook calendar query failed")


__all__ = ["OPERATIONS", "QUERY_IDS", "handle_query"]
