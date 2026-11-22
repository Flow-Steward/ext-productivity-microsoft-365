"""Microsoft Teams operation registry and the team / channel / chat picker queries."""

from __future__ import annotations

from typing import Any

from m365.auth import transport_for
from m365.connection import mapping, resolve_connection, text
from m365.errors import M365Error, error_response
from m365.runtime import Operation, TransportFactory

from .posting import (
    POST_CHANNEL_FIELDS,
    POST_CHAT_FIELDS,
    REPLY_FIELDS,
    post_channel_message,
    post_chat_message,
    reply_to_message,
)
from .reads import (
    LIST_CHATS_FIELDS,
    LIST_MESSAGES_FIELDS,
    MAX_LIMIT,
    joined_teams,
    list_channels,
    list_chats,
    list_messages,
    list_teams,
    project_chat,
    project_team,
    scan_pages,
    team_channels,
    test_teams_connection,
)

SERVICE = "teams"
USER_TYPE = "m365_teams_user"
#: The channel picker reads the channels of at most this many joined teams.
MAX_PICKER_TEAMS = 50


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
        _op("test_teams_connection", test_teams_connection),
        _op("list_teams", list_teams),
        _op("list_channels", list_channels, {"team_id"}),
        _op("list_chats", list_chats, LIST_CHATS_FIELDS),
        _op("list_messages", list_messages, LIST_MESSAGES_FIELDS),
        _op("post_channel_message", post_channel_message, POST_CHANNEL_FIELDS, effectful=True),
        _op("reply_to_message", reply_to_message, REPLY_FIELDS, effectful=True),
        _op("post_chat_message", post_chat_message, POST_CHAT_FIELDS, effectful=True),
    )
}

#: Resource query id → the one connection type it serves.
QUERY_TYPES = {"teams_user": USER_TYPE, "channels_user": USER_TYPE, "chats_user": USER_TYPE}
QUERY_IDS = frozenset(QUERY_TYPES)


def handle_query(
    payload: dict[str, Any], *, transport_factory: TransportFactory = transport_for
) -> dict[str, Any]:
    try:
        query = mapping(payload.get("query"))
        query_id = text(query.get("query_id"), limit=128)
        expected_type = QUERY_TYPES.get(query_id)
        if expected_type is None:
            raise M365Error("invalid_payload", "Unknown Microsoft Teams resource query")
        params = mapping(query.get("params"))
        search = text(params.get("search", query.get("search")), limit=256).lower()
        cursor = text(params.get("cursor", query.get("cursor")), limit=8192)
        raw_limit = params.get("limit", query.get("limit"))
        limit = (
            raw_limit
            if isinstance(raw_limit, int) and not isinstance(raw_limit, bool) and raw_limit > 0
            else 25
        )
        limit = min(limit, MAX_LIMIT)
        context_id = text(mapping(query.get("context")).get("connection_id"), limit=256)
        connection = resolve_connection(payload, service=SERVICE, connection_ref=context_id)
        if connection.connection_type_id != expected_type:
            raise M365Error(
                "invalid_connection", f"{query_id} serves {expected_type} connections only"
            )
        transport = transport_factory(connection)
        next_cursor = ""
        if query_id == "teams_user":
            found = _team_items(transport, search)
        elif query_id == "channels_user":
            found = _channel_items(transport, search)
        else:
            found, next_cursor = _chat_items(
                transport, connection.connection_id, search, cursor, limit
            )
        return {
            "ok": True,
            "result": {"items": found, "next_cursor": next_cursor, "total": len(found)},
        }
    except M365Error as exc:
        return error_response(exc.code, exc.message, **exc.retry_facts())
    except Exception:
        return error_response("internal_error", "The Microsoft Teams resource query failed")


def _team_items(transport: Any, search: str) -> list[dict[str, Any]]:
    items = []
    for row in joined_teams(transport):
        team = project_team(row)
        if search and search not in team["display_name"].lower():
            continue
        items.append({"id": team["team_id"], "label": team["display_name"]})
    return items


def _channel_items(transport: Any, search: str) -> list[dict[str, Any]]:
    """``team_id/channel_id`` values (the notification route format) labelled Team / Channel."""
    items = []
    for row in joined_teams(transport)[:MAX_PICKER_TEAMS]:
        team = project_team(row)
        try:
            channels = team_channels(transport, team["team_id"])
        except M365Error as exc:
            if exc.code in {"team_not_found", "permission_denied"}:
                continue  # an archived or restricted team must not break the picker
            raise
        for channel in channels:
            label = f"{team['display_name']} / {channel['display_name']}"
            if search and search not in label.lower():
                continue
            items.append({"id": channel["target"], "label": label})
    return items


def _chat_items(
    transport: Any, connection_id: str, search: str, cursor: str, limit: int
) -> tuple[list[dict[str, Any]], str]:
    def accept(row: dict[str, Any]) -> dict[str, Any] | None:
        chat = project_chat(row)
        names = ", ".join(m["display_name"] or m["email"] for m in chat["members"])
        label = chat["topic"] or names or chat["chat_id"]
        if search and search not in f"{label} {names}".lower():
            return None
        return {"id": chat["chat_id"], "label": label, "chat_type": chat["chat_type"]}

    return scan_pages(
        transport,
        path="/me/chats",
        query={"$expand": "members", "$top": "50"},
        cursor=cursor,
        scope={"op": "chats_user", "connection": connection_id, "search": search},
        accept=accept,
        limit=limit,
        purpose="Teams chats",
    )


__all__ = ["OPERATIONS", "QUERY_IDS", "handle_query"]
