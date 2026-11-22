"""Teams reads (task D3): connection test, joined teams, channels, chats and messages."""

from __future__ import annotations

from collections.abc import Callable, Mapping
from typing import Any
from urllib.parse import urlencode

from m365.connection import mapping
from m365.errors import M365Error
from m365.runtime import OperationContext
from m365.transport import GRAPH_BASE

from .common import (
    MAX_BODY_CHARS,
    channel_path,
    chat_path,
    email_input,
    html_to_text,
    int_input,
    list_all,
    not_found_as,
    provider_text,
    team_id_input,
    text_input,
    thread_id_input,
)
from .cursors import decode_cursor, encode_cursor, fingerprint
from .markers import strip_markers

MAX_TEAMS = 1000
MAX_CHANNELS = 1000
GRAPH_PAGE_SIZE = 50
#: A filtered list reads at most this many Graph pages per call, then returns a cursor.
MAX_SCAN_PAGES = 10
DEFAULT_LIMIT = 25
MAX_LIMIT = 50

TEAM_SELECT = "id,displayName,description"
CHANNEL_SELECT = "id,displayName,description,membershipType,webUrl"

LIST_CHATS_FIELDS = frozenset({"topic_contains", "member", "limit", "cursor"})
LIST_MESSAGES_FIELDS = frozenset({"team_id", "channel_id", "chat_id", "limit", "cursor"})

Accept = Callable[[dict[str, Any]], dict[str, Any] | None]


# -- operations ----------------------------------------------------------------------------


def test_teams_connection(ctx: OperationContext) -> dict[str, Any]:
    me = ctx.transport.get_json(
        "/me", query={"$select": "id,displayName,userPrincipalName,mail"}, purpose="Teams user"
    )
    teams = list_all(
        ctx.transport,
        "/me/joinedTeams",
        query={"$select": "id"},
        max_items=MAX_TEAMS,
        purpose="Teams joined teams",
    )
    return {
        "user_id": provider_text(me.get("id"), limit=64),
        "display_name": provider_text(me.get("displayName"), limit=256),
        "user_principal_name": provider_text(
            me.get("userPrincipalName") or me.get("mail"), limit=320
        ),
        "auth_mode": "user",
        "joined_team_count": len(teams),
    }


def list_teams(ctx: OperationContext) -> dict[str, Any]:
    teams = [project_team(row) for row in joined_teams(ctx.transport)]
    return {"teams": teams, "count": len(teams)}


def list_channels(ctx: OperationContext) -> dict[str, Any]:
    team_id = team_id_input(ctx.input.get("team_id"))
    channels = team_channels(ctx.transport, team_id)
    return {"team_id": team_id, "channels": channels, "count": len(channels)}


def list_chats(ctx: OperationContext) -> dict[str, Any]:
    topic = text_input(ctx.input.get("topic_contains"), "topic_contains", max_length=255).lower()
    member = ""
    if ctx.input.get("member") not in (None, ""):
        member = email_input(ctx.input.get("member"), "member").lower()
    limit = int_input(
        ctx.input.get("limit"), "limit", default=DEFAULT_LIMIT, minimum=1, maximum=MAX_LIMIT
    )
    cursor = text_input(ctx.input.get("cursor"), "cursor", max_length=8192)

    def accept(row: dict[str, Any]) -> dict[str, Any] | None:
        chat = project_chat(row)
        if topic and topic not in chat["topic"].lower():
            return None
        if member and not any(m["email"].lower() == member for m in chat["members"]):
            return None
        return chat

    path = "/me/chats"
    query = {"$expand": "members", "$top": str(GRAPH_PAGE_SIZE)}
    scope = {
        "op": "list_chats",
        "connection": ctx.connection.connection_id,
        "topic": topic,
        "member": member,
    }
    chats, next_cursor = scan_pages(
        ctx.transport,
        path=path,
        query=query,
        cursor=cursor,
        scope=scope,
        accept=accept,
        limit=limit,
        purpose="Teams chats",
    )
    return {"chats": chats, "next_cursor": next_cursor, "count": len(chats)}


def list_messages(ctx: OperationContext) -> dict[str, Any]:
    raw = ctx.input
    limit = int_input(
        raw.get("limit"), "limit", default=DEFAULT_LIMIT, minimum=1, maximum=MAX_LIMIT
    )
    cursor = text_input(raw.get("cursor"), "cursor", max_length=8192)
    has_channel = raw.get("team_id") not in (None, "") or raw.get("channel_id") not in (None, "")
    has_chat = raw.get("chat_id") not in (None, "")
    if has_channel == has_chat:
        raise M365Error("invalid_payload", "Give team_id and channel_id, or chat_id")
    query: dict[str, str] = {"$top": str(GRAPH_PAGE_SIZE)}
    if has_chat:
        chat_id = thread_id_input(raw.get("chat_id"), "chat_id")
        path = f"{chat_path(chat_id)}/messages"
        query["$orderby"] = "createdDateTime desc"
        scope = {"op": "list_messages", "chat": chat_id}
        not_found = ("chat_not_found", "The chat was not found or the user is not a member")
    else:
        team_id = team_id_input(raw.get("team_id"))
        channel_id = thread_id_input(raw.get("channel_id"), "channel_id")
        path = f"{channel_path(team_id, channel_id)}/messages"
        scope = {"op": "list_messages", "team": team_id, "channel": channel_id}
        not_found = ("channel_not_found", "The team or channel was not found")
    scope["connection"] = ctx.connection.connection_id
    with not_found_as(*not_found):
        messages, next_cursor = scan_pages(
            ctx.transport,
            path=path,
            query=query,
            cursor=cursor,
            scope=scope,
            accept=project_message,
            limit=limit,
            purpose="Teams messages",
        )
    return {"messages": messages, "next_cursor": next_cursor, "count": len(messages)}


# -- shared reads (operations and resource queries) -------------------------------------------


def joined_teams(transport: Any) -> list[dict[str, Any]]:
    return list_all(
        transport,
        "/me/joinedTeams",
        query={"$select": TEAM_SELECT},
        max_items=MAX_TEAMS,
        purpose="Teams joined teams",
    )


def team_channels(transport: Any, team_id: str) -> list[dict[str, Any]]:
    with not_found_as("team_not_found", "The team was not found or the user is not a member"):
        rows = list_all(
            transport,
            f"/teams/{team_id}/channels",
            query={"$select": CHANNEL_SELECT},
            max_items=MAX_CHANNELS,
            purpose="Teams channels",
        )
    return [project_channel(team_id, row) for row in rows]


def scan_pages(
    transport: Any,
    *,
    path: str,
    query: Mapping[str, str],
    cursor: str,
    scope: Mapping[str, Any],
    accept: Accept,
    limit: int,
    purpose: str,
) -> tuple[list[dict[str, Any]], str]:
    """Collect up to ``limit`` accepted rows; the cursor resumes at the next unread row."""
    query_fp = fingerprint(scope)
    if cursor:
        link, offset = decode_cursor(cursor, path_prefix=path, query_fingerprint=query_fp)
    else:
        link, offset = f"{GRAPH_BASE}{path}?{urlencode(dict(query))}", 0
    results: list[dict[str, Any]] = []
    for _page in range(MAX_SCAN_PAGES):
        page = transport.get_json(link, purpose=purpose)
        rows = page.get("value")
        if not isinstance(rows, list):
            raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid list")
        next_link = page.get("@odata.nextLink")
        next_link = next_link if isinstance(next_link, str) and next_link else ""
        for index in range(offset, len(rows)):
            row = rows[index]
            projected = accept(dict(row)) if isinstance(row, Mapping) else None
            if projected is None:
                continue
            results.append(projected)
            if len(results) >= limit:
                if index + 1 < len(rows):
                    return results, _cursor(link, index + 1, path, query_fp)
                return results, _cursor(next_link, 0, path, query_fp) if next_link else ""
        if not next_link:
            return results, ""
        link, offset = next_link, 0
    return results, _cursor(link, 0, path, query_fp)


def _cursor(link: str, offset: int, path: str, query_fp: str) -> str:
    return encode_cursor(link, offset, path_prefix=path, query_fingerprint=query_fp)


# -- projections ----------------------------------------------------------------------------------


def project_team(row: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "team_id": provider_text(row.get("id"), limit=64),
        "display_name": provider_text(row.get("displayName"), limit=256),
        "description": provider_text(row.get("description"), limit=1024),
    }


def project_channel(team_id: str, row: Mapping[str, Any]) -> dict[str, Any]:
    channel_id = provider_text(row.get("id"), limit=512)
    return {
        "channel_id": channel_id,
        "team_id": team_id,
        "display_name": provider_text(row.get("displayName"), limit=256),
        "description": provider_text(row.get("description"), limit=1024),
        "membership_type": provider_text(row.get("membershipType"), limit=32) or "standard",
        "target": f"{team_id}/{channel_id}",
    }


def project_chat(row: Mapping[str, Any]) -> dict[str, Any]:
    members = []
    for raw in row.get("members") or []:
        if isinstance(raw, Mapping):
            members.append(
                {
                    "display_name": provider_text(raw.get("displayName"), limit=256),
                    "email": provider_text(raw.get("email"), limit=320),
                    "user_id": provider_text(raw.get("userId"), limit=64),
                }
            )
    return {
        "chat_id": provider_text(row.get("id"), limit=512),
        "topic": provider_text(row.get("topic"), limit=512),
        "chat_type": provider_text(row.get("chatType"), limit=32),
        "web_url": provider_text(row.get("webUrl"), limit=2048),
        "last_updated_at": provider_text(row.get("lastUpdatedDateTime"), limit=64) or None,
        "members": members[:250],
    }


def project_message(row: dict[str, Any]) -> dict[str, Any] | None:
    """A user message with its text body; system events and deleted messages are skipped."""
    if row.get("messageType") not in (None, "message") or row.get("deletedDateTime"):
        return None
    body = mapping(row.get("body"))
    content = strip_markers(provider_text(body.get("content"), limit=MAX_BODY_CHARS * 2))
    if str(body.get("contentType") or "").lower() == "html":
        content = html_to_text(content)
    sender = mapping(row.get("from"))
    user = mapping(sender.get("user"))
    app = mapping(sender.get("application"))
    who = user or app
    return {
        "message_id": provider_text(row.get("id"), limit=128),
        "reply_to_id": provider_text(row.get("replyToId"), limit=128) or None,
        "created_at": provider_text(row.get("createdDateTime"), limit=64) or None,
        "last_modified_at": provider_text(row.get("lastModifiedDateTime"), limit=64) or None,
        "subject": provider_text(row.get("subject"), limit=512),
        "body_text": strip_markers(content),
        "importance": provider_text(row.get("importance"), limit=16) or "normal",
        "sender": {
            "display_name": provider_text(who.get("displayName"), limit=256),
            "user_id": provider_text(who.get("id"), limit=64),
            "kind": "user" if user else ("application" if app else "unknown"),
        },
        "web_url": provider_text(row.get("webUrl"), limit=2048),
    }


__all__ = [
    "LIST_CHATS_FIELDS",
    "LIST_MESSAGES_FIELDS",
    "joined_teams",
    "list_channels",
    "list_chats",
    "list_messages",
    "list_teams",
    "project_channel",
    "project_chat",
    "project_message",
    "project_team",
    "scan_pages",
    "team_channels",
    "test_teams_connection",
]
