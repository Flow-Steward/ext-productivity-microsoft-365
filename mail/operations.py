"""Outlook mail operation registry and the folder picker queries."""

from __future__ import annotations

from typing import Any

from m365.auth import transport_for
from m365.connection import mapping, text
from m365.errors import M365Error, error_response
from m365.runtime import Operation, TransportFactory, run_operation

from .access_check import check_mail_access
from .download import get_attachment
from .mutations import delete_messages, move_messages, set_message_flags
from .reads import (
    SEARCH_FIELDS,
    connection_test,
    folder_items,
    get_message,
    list_mail_folders,
    search_messages,
)
from .send import (
    FORWARD_FIELDS,
    REPLY_FIELDS,
    SEND_FIELDS,
    forward_message,
    reply_message,
    send_message,
)

SERVICE = "mail"


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
        _op("test_connection", connection_test),
        _op("list_mail_folders", list_mail_folders),
        _op("search_messages", search_messages, SEARCH_FIELDS),
        _op(
            "get_message",
            get_message,
            {
                "message_id",
                "include_body",
                "include_headers",
                "include_attachment_metadata",
                "include_raw_html",
            },
        ),
        _op("get_attachment", get_attachment, {"message_id", "attachment_id"}),
        _op("send_message", send_message, SEND_FIELDS, effectful=True),
        _op("reply_message", reply_message, REPLY_FIELDS, effectful=True),
        _op("forward_message", forward_message, FORWARD_FIELDS, effectful=True),
        _op(
            "set_message_flags",
            set_message_flags,
            {"message_ids", "is_read", "flag_status"},
            effectful=True,
        ),
        _op("move_messages", move_messages, {"message_ids", "destination_folder"}, effectful=True),
        _op("delete_messages", delete_messages, {"message_ids"}, effectful=True),
        # Setup action on the Mail page (B4.3): creates and deletes a probe draft in the
        # control mailbox, never sends. Not a workflow operation.
        _op("check_mail_access", check_mail_access, {"control_mailbox"}),
    )
}

#: Resource query id → the one connection type it serves.
FOLDER_QUERIES = {"mail_folders_user": "m365_mail_user", "mail_folders_app": "m365_mail_app"}
QUERY_IDS = frozenset(FOLDER_QUERIES)


def handle_query(
    payload: dict[str, Any], *, transport_factory: TransportFactory = transport_for
) -> dict[str, Any]:
    try:
        query = mapping(payload.get("query"))
        query_id = text(query.get("query_id"), limit=128)
        expected_type = FOLDER_QUERIES.get(query_id)
        if expected_type is None:
            raise M365Error("invalid_payload", "Unknown Outlook resource query")
        connection_id = text(mapping(query.get("context")).get("connection_id"), limit=256)
        target = mapping(payload.get("target"))
        connection = mapping(target.get("connection"))
        if not connection_id or not connection:
            raise M365Error(
                "invalid_connection", "The folder picker needs a selected mail connection"
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
                "action_id": "list_mail_folders",
                "input": {"connection_ref": connection_id},
                "target": target,
            },
        }
        response = run_operation(action_payload, OPERATIONS, transport_factory=transport_factory)
        if response.get("ok") is not True:
            return response
        items = folder_items(mapping(response.get("result")))
        if search:
            items = [row for row in items if search in row["label"].lower()]
        return {"ok": True, "result": {"items": items, "next_cursor": "", "total": len(items)}}
    except M365Error as exc:
        response = error_response(exc.code, exc.message)
        response.update(exc.retry_facts())
        return response
    except Exception:
        return error_response("internal_error", "The Outlook folder query failed")


__all__ = ["OPERATIONS", "QUERY_IDS", "handle_query"]
