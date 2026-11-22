"""Read operations: connection test, calendars, calendar view and one event."""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any

from m365.connection import mapping
from m365.errors import M365Error
from m365.runtime import OperationContext
from mail.common import Mailbox, bool_input, int_input, not_found_as, text_input
from mail.cursors import decode_cursor, encode_cursor, fingerprint

from .common import (
    CALENDAR_NOT_FOUND,
    CALENDAR_SELECT,
    EVENT_NOT_FOUND,
    EVENT_RESPONSE_LIMIT,
    EVENT_SELECT,
    MAX_LIST_DAYS,
    PREFER_TEXT_BODY,
    PREFER_UTC,
    SMALL_RESPONSE_LIMIT,
    calendar_id_input,
    calendar_out,
    calendar_path,
    client_for,
    event_id_input,
    event_out,
    event_path,
    iso_z,
    window_input,
)
from .zones import zone_input

MAX_CALENDARS = 500
MAX_BODY_BYTES = 1024 * 1024
LIST_EVENTS_FIELDS = frozenset({"calendar_id", "start", "end", "time_zone", "limit", "cursor"})
GET_EVENT_FIELDS = frozenset({"event_id", "include_body"})


# -- test_calendar_connection -------------------------------------------------------------


def calendar_connection_test(ctx: OperationContext) -> dict[str, Any]:
    client = client_for(ctx.transport, ctx.connection)
    address = ctx.connection.mailbox
    display_name: str | None = None
    if client.base == "/me":
        me = ctx.transport.request(
            "GET",
            "/me",
            query={"$select": "displayName,mail,userPrincipalName"},
            headers=Mailbox.headers(),
            response_limit=SMALL_RESPONSE_LIMIT,
            retry="read",
            purpose="Microsoft 365 profile",
        ).json()
        address = str(me.get("mail") or me.get("userPrincipalName") or "")
        display_name = str(me.get("displayName") or "") or None
    with not_found_as("mailbox_not_found", "The mailbox or its calendar was not found"):
        row = client.get(
            "/calendar",
            query={"$select": CALENDAR_SELECT},
            response_limit=SMALL_RESPONSE_LIMIT,
            purpose="Outlook default calendar",
        )
    calendar = calendar_out(row)
    return {
        "mailbox_address": address,
        "display_name": display_name,
        "auth_mode": ctx.connection.mode,
        "calendar_id": calendar["calendar_id"],
        "calendar_name": calendar["name"],
        "can_edit": calendar["can_edit"],
        "allowed_online_meeting_providers": calendar["allowed_online_meeting_providers"],
        "teams_meetings_allowed": calendar["teams_meetings_allowed"],
    }


# -- list_calendars -----------------------------------------------------------------------


def list_calendars(ctx: OperationContext) -> dict[str, Any]:
    client = client_for(ctx.transport, ctx.connection)
    rows = client.list_all(
        "/calendars",
        query={"$select": CALENDAR_SELECT, "$top": "100"},
        max_items=MAX_CALENDARS,
        purpose="Outlook calendars",
    )
    calendars = [calendar_out(row) for row in rows if row.get("id")]
    return {"calendars": calendars, "count": len(calendars)}


def calendar_items(result: Mapping[str, Any]) -> list[dict[str, Any]]:
    """Picker rows: the default calendar first, then the others by name."""
    rows = sorted(
        (row for row in result.get("calendars") or [] if isinstance(row, Mapping)),
        key=lambda row: (not row.get("is_default"), str(row.get("name") or "").lower()),
    )
    items = []
    for row in rows:
        label = str(row.get("name") or row.get("calendar_id"))
        if row.get("is_default"):
            label = f"{label} (default)"
        items.append({"calendar_id": row["calendar_id"], "label": label})
    return items


# -- list_events ----------------------------------------------------------------------------


def list_events(ctx: OperationContext) -> dict[str, Any]:
    client = client_for(ctx.transport, ctx.connection)
    raw = ctx.input
    zone = zone_input(raw["time_zone"], "time_zone")[1] if raw.get("time_zone") else None
    start, end = window_input(raw, zone, max_days=MAX_LIST_DAYS, strict=False)
    normalized: dict[str, Any] = {
        "calendar_id": calendar_id_input(raw.get("calendar_id")),
        "start": iso_z(start),
        "end": iso_z(end),
        "limit": int_input(raw.get("limit"), "limit", default=50, minimum=1, maximum=100),
    }
    key = fingerprint(client.base, {"operation": "list_events", **normalized})
    cursor = text_input(raw.get("cursor"), "cursor", max_length=8192)
    with not_found_as("calendar_not_found", CALENDAR_NOT_FOUND):
        if cursor:
            link = decode_cursor(cursor, base=client.base, query_fingerprint=key)
            page = client.next_page(link, prefer=(PREFER_UTC,), purpose="Outlook calendar view")
        else:
            page = client.get(
                f"{calendar_path(normalized['calendar_id'])}/calendarView",
                query={
                    "startDateTime": normalized["start"],
                    "endDateTime": normalized["end"],
                    "$select": EVENT_SELECT,
                    "$orderby": "start/dateTime",
                    "$top": str(normalized["limit"]),
                },
                prefer=(PREFER_UTC,),
                response_limit=EVENT_RESPONSE_LIMIT,
                purpose="Outlook calendar view",
            )
    rows = page.get("value")
    if not isinstance(rows, list):
        raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid event list")
    events = [event_out(row) for row in rows[: normalized["limit"]] if isinstance(row, Mapping)]
    next_link = page.get("@odata.nextLink")
    next_cursor = (
        encode_cursor(next_link, base=client.base, query_fingerprint=key)
        if isinstance(next_link, str) and next_link
        else ""
    )
    return {
        "events": events,
        "count": len(events),
        "next_cursor": next_cursor,
        "window_start_utc": normalized["start"],
        "window_end_utc": normalized["end"],
    }


# -- get_event --------------------------------------------------------------------------------


def get_event(ctx: OperationContext) -> dict[str, Any]:
    client = client_for(ctx.transport, ctx.connection)
    event_id = event_id_input(ctx.input.get("event_id"))
    include_body = bool_input(ctx.input.get("include_body"), "include_body", default=True)
    select = EVENT_SELECT + (",body" if include_body else "")
    with not_found_as("event_not_found", EVENT_NOT_FOUND):
        row = client.get(
            event_path(event_id),
            query={"$select": select},
            prefer=(PREFER_UTC, PREFER_TEXT_BODY),
            response_limit=EVENT_RESPONSE_LIMIT,
            purpose="Outlook event",
        )
    event = event_out(row)
    if include_body:
        content = str(mapping(row.get("body")).get("content") or "")
        encoded = content.encode("utf-8")
        event["body_text_truncated"] = len(encoded) > MAX_BODY_BYTES
        event["body_text"] = (
            encoded[:MAX_BODY_BYTES].decode("utf-8", errors="ignore")
            if event["body_text_truncated"]
            else content
        )
    return {"event": event}


__all__ = [
    "GET_EVENT_FIELDS",
    "LIST_EVENTS_FIELDS",
    "calendar_connection_test",
    "calendar_items",
    "get_event",
    "list_calendars",
    "list_events",
]
