"""Teams posts (task D3): channel messages, thread replies and chat messages.

Contract: Graph has no idempotency key for chat messages, so an ambiguous post (no
response, 5xx, timeout) is **never** sent again. Every POST of a message uses
``retry="none"``. As a best effort the body carries a hidden marker derived from the
step key (:mod:`teams.markers`); after an ambiguous post the latest messages of the
target are read and a message with the marker counts as posted
(``confirmation: found_by_marker``). Otherwise the step fails with ``timeout_unknown``.

A body above Graph's ~28 KB message limit fails with ``message_too_large`` before
anything is sent.
"""

from __future__ import annotations

import json
import time
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from html import escape
from typing import Any
from urllib.parse import quote

from m365.connection import mapping
from m365.errors import M365Error
from m365.runtime import OperationContext

from .common import (
    MAX_BODY_CHARS,
    MAX_SUBJECT_CHARS,
    channel_path,
    chat_path,
    email_input,
    file_links_input,
    mentions_input,
    message_id_input,
    not_found_as,
    provider_message_id,
    provider_text,
    provider_thread_id,
    team_id_input,
    text_input,
    text_to_html,
    thread_id_input,
)
from .markers import DEFAULT_MARKER, MarkerStrategy, strip_markers

#: Graph: "the message size limit is approximately 28 KB" (whole request body).
MAX_MESSAGE_BYTES = 28 * 1024
CONFIRM_ATTEMPTS = 3
CONFIRM_DELAY_SECONDS = 2.0
CONFIRM_TOP = 20

#: Indirection point for tests.
sleep: Callable[[float], None] = time.sleep

_BODY_FIELDS = frozenset({"body_text", "body_html", "mentions", "file_links"})
POST_CHANNEL_FIELDS = _BODY_FIELDS | {"team_id", "channel_id", "subject"}
REPLY_FIELDS = _BODY_FIELDS | {"team_id", "channel_id", "message_id"}
POST_CHAT_FIELDS = _BODY_FIELDS | {"chat_id", "user"}


# -- targets ---------------------------------------------------------------------------------


@dataclass(frozen=True)
class TeamsTarget:
    """Where a message goes: a channel, a channel thread (``reply_to``) or a chat."""

    team_id: str = ""
    channel_id: str = ""
    chat_id: str = ""
    reply_to: str = ""

    @property
    def kind(self) -> str:
        if self.chat_id:
            return "chat"
        return "reply" if self.reply_to else "channel"

    def messages_path(self) -> str:
        if self.chat_id:
            return f"{chat_path(self.chat_id)}/messages"
        base = f"{channel_path(self.team_id, self.channel_id)}/messages"
        return f"{base}/{quote(self.reply_to, safe='')}/replies" if self.reply_to else base

    def recent_query(self) -> dict[str, str]:
        query = {"$top": str(CONFIRM_TOP)}
        if self.chat_id:
            query["$orderby"] = "createdDateTime desc"
        return query

    def not_found(self) -> tuple[str, str]:
        if self.chat_id:
            return "chat_not_found", "The chat was not found or the user is not a member"
        if self.reply_to:
            return "message_not_found", "The team, channel or message to reply to was not found"
        return "channel_not_found", "The team or channel was not found"


def parse_target(value: Any) -> TeamsTarget:
    """``team_id/channel_id`` or a chat id (``19:...``), as in a notification route."""
    if not isinstance(value, str) or not value.strip():
        raise M365Error(
            "channel_required", "A Teams channel (team_id/channel_id) or chat ID is required"
        )
    raw = value.strip()
    if "/" in raw:
        team, _, channel = raw.partition("/")
        return TeamsTarget(
            team_id=team_id_input(team, "channel"),
            channel_id=thread_id_input(channel, "channel_id"),
        )
    return TeamsTarget(chat_id=thread_id_input(raw, "chat_id"))


# -- message body ----------------------------------------------------------------------------


@dataclass(frozen=True)
class Mention:
    address: str
    user_id: str
    display_name: str


def resolve_mentions(transport: Any, addresses: Sequence[str]) -> list[Mention]:
    """Look up each address (``/users/{address}``) for the mention's user id."""
    result: list[Mention] = []
    for address in addresses:
        try:
            with not_found_as("user_not_found", f"No Microsoft 365 user has the address {address}"):
                row = transport.get_json(
                    f"/users/{quote(address, safe='@.')}",
                    query={"$select": "id,displayName"},
                    purpose="Teams mention lookup",
                )
        except M365Error as exc:
            if exc.code == "permission_denied":
                raise M365Error(
                    "permission_denied",
                    "Mentions need the delegated permission User.ReadBasic.All",
                    http_status=exc.http_status,
                ) from exc
            raise
        user_id = provider_text(row.get("id"), limit=64)
        if not user_id:
            raise M365Error("user_not_found", f"No Microsoft 365 user has the address {address}")
        name = provider_text(row.get("displayName"), limit=256) or address
        result.append(Mention(address, user_id, name))
    return result


def build_message(
    html_body: str,
    *,
    marker_key: str,
    subject: str = "",
    mentions: Sequence[Mention] = (),
    file_links: Sequence[tuple[str, str]] = (),
    marker: MarkerStrategy = DEFAULT_MARKER,
) -> dict[str, Any]:
    """The Graph ``chatMessage`` body; raises ``message_too_large`` above the limit."""
    parts: list[str] = []
    if mentions:
        tags = [
            f'<at id="{index}">{escape(item.display_name, quote=False)}</at>'
            for index, item in enumerate(mentions)
        ]
        parts.append(f"<p>{' '.join(tags)}</p>")
    parts.append(strip_markers(html_body))
    for url, name in file_links:
        parts.append(f'<p><a href="{escape(url)}">{escape(name, quote=False)}</a></p>')
    content = marker.embed("".join(parts), marker_key)
    message: dict[str, Any] = {"body": {"contentType": "html", "content": content}}
    if subject:
        message["subject"] = subject
    if mentions:
        message["mentions"] = [
            {
                "id": index,
                "mentionText": item.display_name,
                "mentioned": {
                    "user": {
                        "id": item.user_id,
                        "displayName": item.display_name,
                        "userIdentityType": "aadUser",
                    }
                },
            }
            for index, item in enumerate(mentions)
        ]
    check_size(message)
    return message


def check_size(message: Mapping[str, Any]) -> None:
    size = len(json.dumps(message, separators=(",", ":"), ensure_ascii=False).encode("utf-8"))
    if size > MAX_MESSAGE_BYTES:
        raise M365Error(
            "message_too_large",
            f"The Teams message is {size} bytes; Teams accepts about {MAX_MESSAGE_BYTES} bytes",
        )


def body_input(raw: Mapping[str, Any]) -> str:
    """The body as HTML from exactly one of ``body_text`` / ``body_html``."""
    body_text = text_input(
        raw.get("body_text"),
        "body_text",
        max_length=MAX_BODY_CHARS,
        allow_newlines=True,
        strip=False,
    )
    body_html = text_input(
        raw.get("body_html"),
        "body_html",
        max_length=MAX_BODY_CHARS,
        allow_newlines=True,
        strip=False,
    )
    if bool(body_text.strip()) == bool(body_html.strip()):
        raise M365Error("invalid_payload", "Give body_text or body_html (exactly one)")
    return body_html if body_html.strip() else text_to_html(body_text)


# -- the post ---------------------------------------------------------------------------------


def post_message(
    transport: Any,
    target: TeamsTarget,
    message: Mapping[str, Any],
    *,
    marker_key: str,
    marker: MarkerStrategy = DEFAULT_MARKER,
) -> dict[str, Any]:
    """POST once; an ambiguous outcome is resolved by the marker or ends in timeout_unknown."""
    check_size(message)
    try:
        with not_found_as(*target.not_found()):
            response = transport.request(
                "POST",
                target.messages_path(),
                json_body=dict(message),
                retry="none",
                purpose="Teams post",
                ok_statuses=(200, 201),
            )
    except M365Error as exc:
        if exc.code == "file_too_large":  # Graph 413
            raise M365Error(
                "message_too_large", "Teams refused the message as too large", http_status=413
            ) from exc
        if not exc.ambiguous:
            raise
        found = find_marked_message(transport, target, marker_key, marker=marker)
        if found is None:
            raise M365Error(
                "timeout_unknown",
                f"Teams did not confirm the post; check the {target.kind} before posting again",
                http_status=exc.http_status,
                ambiguous=True,
            ) from exc
        return post_result(target, found, "found_by_marker")
    try:
        row = response.json()
    except M365Error:
        row = {}  # Teams accepted the message; only its echo is unreadable
    return post_result(target, row, "posted")


def find_marked_message(
    transport: Any,
    target: TeamsTarget,
    marker_key: str,
    *,
    marker: MarkerStrategy = DEFAULT_MARKER,
) -> dict[str, Any] | None:
    """The newest message of the target that carries the marker, or None."""
    for attempt in range(CONFIRM_ATTEMPTS):
        if attempt:
            sleep(CONFIRM_DELAY_SECONDS)
        try:
            page = transport.get_json(
                target.messages_path(), query=target.recent_query(), purpose="Teams post lookup"
            )
        except M365Error:
            continue
        rows = page.get("value")
        for row in rows if isinstance(rows, list) else []:
            if not isinstance(row, Mapping):
                continue
            if marker.found(str(mapping(row.get("body")).get("content") or ""), marker_key):
                return dict(row)
    return None


def post_result(target: TeamsTarget, row: Mapping[str, Any], confirmation: str) -> dict[str, Any]:
    """The success result; a malformed id in Graph's echo never turns a post into a failure."""
    try:
        message_id = provider_message_id(row.get("id"))
    except M365Error:
        message_id = ""
    return {
        "message_id": message_id,
        "team_id": target.team_id,
        "channel_id": target.channel_id,
        "chat_id": target.chat_id,
        "reply_to_id": target.reply_to,
        "created_at": provider_text(row.get("createdDateTime"), limit=64) or None,
        "web_url": provider_text(row.get("webUrl"), limit=2048),
        "confirmation": confirmation,
        "external_effect_status": "succeeded",
        "definitely_no_external_effect": False,
    }


def post_html(
    transport: Any,
    target: TeamsTarget,
    html_body: str,
    *,
    marker_key: str,
    subject: str = "",
    marker: MarkerStrategy = DEFAULT_MARKER,
) -> dict[str, Any]:
    """Post an HTML body to a channel, thread or chat (helper for the notification provider)."""
    if not marker_key:
        raise M365Error("idempotency_key_missing", "The post has no idempotency key")
    if subject and target.kind != "channel":
        subject = ""  # Graph accepts a subject on channel root posts only
    message = build_message(html_body, marker_key=marker_key, subject=subject, marker=marker)
    return post_message(transport, target, message, marker_key=marker_key, marker=marker)


# -- operations -----------------------------------------------------------------------------------


def post_channel_message(ctx: OperationContext) -> dict[str, Any]:
    raw = ctx.input
    target = TeamsTarget(
        team_id=team_id_input(raw.get("team_id")),
        channel_id=thread_id_input(raw.get("channel_id"), "channel_id"),
    )
    subject = text_input(raw.get("subject"), "subject", max_length=MAX_SUBJECT_CHARS)
    return _post(ctx, target, subject=subject)


def reply_to_message(ctx: OperationContext) -> dict[str, Any]:
    raw = ctx.input
    target = TeamsTarget(
        team_id=team_id_input(raw.get("team_id")),
        channel_id=thread_id_input(raw.get("channel_id"), "channel_id"),
        reply_to=message_id_input(raw.get("message_id")),
    )
    return _post(ctx, target)


def post_chat_message(ctx: OperationContext) -> dict[str, Any]:
    raw = ctx.input
    has_chat = raw.get("chat_id") not in (None, "")
    has_user = raw.get("user") not in (None, "")
    if has_chat == has_user:
        raise M365Error("invalid_payload", "Give chat_id or user (exactly one)")
    if has_chat:
        return _post(ctx, TeamsTarget(chat_id=thread_id_input(raw.get("chat_id"), "chat_id")))
    user = email_input(raw.get("user"), "user")
    # Validate and size the message before the chat is created.
    key = ctx.idempotency_key
    message = _message(ctx, key)
    target = TeamsTarget(chat_id=one_on_one_chat(ctx.transport, user))
    return post_message(ctx.transport, target, message, marker_key=key)


def one_on_one_chat(transport: Any, address: str) -> str:
    """The 1:1 chat with ``address``; Graph returns the existing chat if there is one."""
    me = transport.get_json("/me", query={"$select": "id"}, purpose="Teams user")
    my_id = provider_text(me.get("id"), limit=64)
    if not my_id:
        raise M365Error("provider_unavailable", "Microsoft Graph returned no user ID")
    member_type = "#microsoft.graph.aadUserConversationMember"
    body = {
        "chatType": "oneOnOne",
        "members": [
            {
                "@odata.type": member_type,
                "roles": ["owner"],
                "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{my_id}')",
            },
            {
                "@odata.type": member_type,
                "roles": ["owner"],
                "user@odata.bind": "https://graph.microsoft.com/v1.0/users('"
                + quote(address, safe="@.")
                + "')",
            },
        ],
    }
    try:
        # Creating a 1:1 chat is idempotent on Graph's side, so the transport may repeat it.
        response = transport.request(
            "POST", "/chats", json_body=body, retry="idempotent", purpose="Teams chat"
        )
    except M365Error as exc:
        if exc.http_status in {400, 404} and not exc.ambiguous:
            raise M365Error(
                "user_not_found",
                f"Teams could not start a chat with {address}; check the address",
                http_status=exc.http_status,
            ) from exc
        if exc.ambiguous:
            raise M365Error(
                "provider_unavailable",
                "Teams did not confirm the chat; no message was posted",
                http_status=exc.http_status,
                retryable=True,
                failure_class="provider",
            ) from exc
        raise
    return provider_thread_id(response.json().get("id"))


def _message(ctx: OperationContext, key: str, *, subject: str = "") -> dict[str, Any]:
    raw = ctx.input
    html_body = body_input(raw)
    links = file_links_input(raw.get("file_links"))
    addresses = mentions_input(raw.get("mentions"))
    # Size check before any lookup: the body alone may already be too large.
    build_message(html_body, marker_key=key, subject=subject, file_links=links)
    mentions = resolve_mentions(ctx.transport, addresses) if addresses else []
    return build_message(
        html_body, marker_key=key, subject=subject, mentions=mentions, file_links=links
    )


def _post(ctx: OperationContext, target: TeamsTarget, *, subject: str = "") -> dict[str, Any]:
    key = ctx.idempotency_key
    message = _message(ctx, key, subject=subject)
    return post_message(ctx.transport, target, message, marker_key=key)


__all__ = [
    "MAX_MESSAGE_BYTES",
    "POST_CHANNEL_FIELDS",
    "POST_CHAT_FIELDS",
    "REPLY_FIELDS",
    "Mention",
    "TeamsTarget",
    "build_message",
    "check_size",
    "find_marked_message",
    "one_on_one_chat",
    "parse_target",
    "post_channel_message",
    "post_chat_message",
    "post_html",
    "post_message",
    "post_result",
    "reply_to_message",
    "resolve_mentions",
]
