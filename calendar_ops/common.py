"""Shared helpers of the calendar package: limits, time input, the calendar client and
the event output shape.

Every read asks Graph for UTC (``Prefer: outlook.timezone="UTC"``) and immutable ids.
Results carry UTC plus the event's original zone (task D2, "Time zones").
"""

from __future__ import annotations

import re
from collections.abc import Iterator, Mapping
from contextlib import contextmanager
from datetime import UTC, datetime, timedelta, tzinfo
from typing import Any

from m365.connection import ConnectionContext, mailbox_base, mapping
from m365.errors import M365Error
from mail.common import (
    Mailbox,
    address_out,
    graph_id,
    not_found_as,
    recipients_input,
    segment,
    text_input,
)
from mail.send import KEY_PROPERTY_ID

from .zones import lookup_zone, to_local

#: The step key lives in the v0.1 mail property set (``FlowStewardIdempotencyKey``).
EVENT_KEY_PROPERTY_ID = KEY_PROPERTY_ID
PREFER_UTC = 'outlook.timezone="UTC"'
PREFER_TEXT_BODY = 'outlook.body-content-type="text"'
TEAMS_PROVIDER = "teamsForBusiness"
MAX_LIST_DAYS = 31
#: Graph's getSchedule window must be shorter than 62 days.
MAX_SCHEDULE_DAYS = 62
MAX_SCHEDULE_ADDRESSES = 20
MAX_ATTENDEES = 500
MAX_SUBJECT_CHARS = 255
MAX_BODY_CHARS = 1_000_000
MAX_LOCATION_CHARS = 255
MAX_COMMENT_CHARS = 10_000
MAX_REMINDER_MINUTES = 40_320
EVENT_RESPONSE_LIMIT = 4 * 1024 * 1024
SMALL_RESPONSE_LIMIT = 64 * 1024
EVENT_NOT_FOUND = "The event was not found in this calendar"
CALENDAR_NOT_FOUND = "The calendar was not found in this mailbox"
_GRAPH_TIME = re.compile(r"^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})(?:\.\d+)?(Z|[+-]\d{2}:\d{2})?$")
EVENT_SELECT = (
    "id,iCalUId,seriesMasterId,type,subject,bodyPreview,start,end,originalStartTimeZone,"
    "originalEndTimeZone,isAllDay,isCancelled,isOrganizer,organizer,attendees,location,"
    "isOnlineMeeting,onlineMeetingProvider,onlineMeeting,showAs,responseStatus,webLink,"
    "isReminderOn,reminderMinutesBeforeStart,importance,sensitivity"
)
CALENDAR_SELECT = (
    "id,name,color,isDefaultCalendar,canEdit,canShare,owner,allowedOnlineMeetingProviders,"
    "defaultOnlineMeetingProvider"
)


# -- the calendar client ----------------------------------------------------------------


def client_for(transport: Any, connection: ConnectionContext) -> Mailbox:
    return Mailbox(transport, mailbox_base(connection))


def calendar_path(calendar_id: str) -> str:
    """The default calendar, or one calendar of the mailbox by id."""
    return f"/calendars/{segment(calendar_id)}" if calendar_id else "/calendar"


def events_path(calendar_id: str) -> str:
    return f"/calendars/{segment(calendar_id)}/events" if calendar_id else "/events"


def event_path(event_id: str) -> str:
    return f"/events/{segment(event_id)}"


def read_event(client: Mailbox, event_id: str, *, purpose: str) -> dict[str, Any]:
    with not_found_as("event_not_found", EVENT_NOT_FOUND):
        return client.get(
            event_path(event_id),
            query={"$select": EVENT_SELECT},
            prefer=(PREFER_UTC, PREFER_TEXT_BODY),
            response_limit=EVENT_RESPONSE_LIMIT,
            purpose=purpose,
        )


def read_calendar(client: Mailbox, calendar_id: str, *, purpose: str) -> dict[str, Any]:
    with not_found_as("calendar_not_found", CALENDAR_NOT_FOUND):
        return client.get(
            calendar_path(calendar_id),
            query={"$select": CALENDAR_SELECT},
            response_limit=SMALL_RESPONSE_LIMIT,
            purpose=purpose,
        )


@contextmanager
def as_read_failure() -> Iterator[None]:
    """A POST that only reads (``getSchedule``) never reports ``timeout_unknown``."""
    try:
        yield
    except M365Error as exc:
        if not exc.ambiguous:
            raise
        raise M365Error(
            "provider_unavailable" if exc.http_status else "network_failure",
            "Microsoft 365 did not answer the request",
            http_status=exc.http_status,
            retry_after_seconds=exc.retry_after_seconds,
        ) from exc


# -- input ------------------------------------------------------------------------------


def calendar_id_input(value: Any) -> str:
    if value is None or value == "":
        return ""
    return graph_id(value, "calendar_id")


def event_id_input(value: Any) -> str:
    return graph_id(value, "event_id")


def moment_input(value: Any, field: str, zone: tzinfo | None) -> datetime:
    """An ISO 8601 date-time as an aware UTC instant.

    With an offset (``Z``, ``+02:00``) the instant is absolute; without one it is a local
    time in ``zone``, which is then required.
    """
    if not isinstance(value, str) or not value.strip() or len(value) > 64:
        raise M365Error("invalid_payload", f"{field} must be an ISO 8601 date-time")
    try:
        parsed = datetime.fromisoformat(value.strip().replace("Z", "+00:00"))
    except ValueError as exc:
        raise M365Error("invalid_payload", f"{field} must be an ISO 8601 date-time") from exc
    if parsed.tzinfo is None:
        if zone is None:
            raise M365Error(
                "invalid_payload", f"{field} has no offset: set time_zone or add an offset"
            )
        parsed = parsed.replace(tzinfo=zone)
    return parsed.astimezone(UTC).replace(microsecond=0)


def window_input(
    raw: Mapping[str, Any], zone: tzinfo | None, *, max_days: int, strict: bool
) -> tuple[datetime, datetime]:
    """``start``/``end`` of a read window; ``strict`` means shorter than ``max_days``."""
    start = moment_input(raw.get("start"), "start", zone)
    end = moment_input(raw.get("end"), "end", zone)
    if end <= start:
        raise M365Error("invalid_payload", "end must be later than start")
    limit = timedelta(days=max_days)
    if (end - start >= limit) if strict else (end - start > limit):
        bound = "shorter than" if strict else "at most"
        raise M365Error("range_too_large", f"The window must be {bound} {max_days} days")
    return start, end


def iso_z(moment: datetime) -> str:
    return moment.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def iso_local(moment: datetime) -> str:
    return moment.strftime("%Y-%m-%dT%H:%M:%S")


def graph_time(moment: datetime, zone_name: str, zone: tzinfo) -> dict[str, str]:
    """A Graph ``dateTimeTimeZone`` that keeps the caller's zone name."""
    return {"dateTime": iso_local(to_local(moment, zone)), "timeZone": zone_name}


def attendees_input(raw: Mapping[str, Any]) -> list[dict[str, Any]] | None:
    """``attendees`` (required) plus ``optional_attendees``; ``None`` when neither is given."""
    if raw.get("attendees") is None and raw.get("optional_attendees") is None:
        return None
    required = recipients_input(raw.get("attendees"), "attendees", maximum=MAX_ATTENDEES)
    optional = recipients_input(
        raw.get("optional_attendees"), "optional_attendees", maximum=MAX_ATTENDEES
    )
    seen = {row["emailAddress"]["address"].lower() for row in required}
    rows = [{**row, "type": "required"} for row in required]
    rows.extend(
        {**row, "type": "optional"}
        for row in optional
        if row["emailAddress"]["address"].lower() not in seen
    )
    if len(rows) > MAX_ATTENDEES:
        raise M365Error("invalid_payload", f"An event can have at most {MAX_ATTENDEES} attendees")
    return rows


def body_input(raw: Mapping[str, Any]) -> dict[str, str] | None:
    body_text = text_input(
        raw.get("body_text"), "body_text", max_length=MAX_BODY_CHARS, allow_newlines=True
    )
    body_html = text_input(
        raw.get("body_html"), "body_html", max_length=MAX_BODY_CHARS, allow_newlines=True
    )
    if body_text and body_html:
        raise M365Error("invalid_payload", "Use body_text or body_html, not both")
    if not (body_text or body_html):
        return None
    return {"contentType": "HTML" if body_html else "Text", "content": body_html or body_text}


# -- provider output --------------------------------------------------------------------


def parse_graph_time(raw: Any) -> datetime | None:
    """A Graph ``dateTimeTimeZone`` as an aware UTC instant (``None`` if unreadable)."""
    value = mapping(raw)
    match = _GRAPH_TIME.match(str(value.get("dateTime") or ""))
    if not match:
        return None
    try:
        parsed = datetime.fromisoformat(match.group(1))
    except ValueError:
        return None
    if match.group(2):
        offset = match.group(2).replace("Z", "+00:00")
        return datetime.fromisoformat(match.group(1) + offset).astimezone(UTC)
    zone = lookup_zone(str(value.get("timeZone") or "UTC"))
    if zone is None:
        return None
    return parsed.replace(tzinfo=zone).astimezone(UTC)


def _local(moment: datetime | None, zone_name: str) -> str | None:
    zone = lookup_zone(zone_name)
    if moment is None or zone is None:
        return None
    return iso_local(to_local(moment, zone))


def attendees_out(raw: Any) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for item in raw[:MAX_ATTENDEES] if isinstance(raw, list) else []:
        address = address_out(item)
        if address is None:
            continue
        status = mapping(mapping(item).get("status"))
        rows.append(
            {
                **address,
                "type": str(mapping(item).get("type") or "required"),
                "response": str(status.get("response") or "none"),
            }
        )
    return rows


def event_out(row: Mapping[str, Any]) -> dict[str, Any]:
    start = parse_graph_time(row.get("start"))
    end = parse_graph_time(row.get("end"))
    start_zone = str(row.get("originalStartTimeZone") or "")
    end_zone = str(row.get("originalEndTimeZone") or start_zone)
    reminder = row.get("reminderMinutesBeforeStart")
    return {
        "event_id": str(row.get("id") or ""),
        "ical_uid": str(row.get("iCalUId") or ""),
        "series_master_id": str(row.get("seriesMasterId") or "") or None,
        "event_type": str(row.get("type") or "singleInstance"),
        "subject": str(row.get("subject") or ""),
        "body_preview": str(row.get("bodyPreview") or "")[:512],
        "start_utc": iso_z(start) if start else None,
        "end_utc": iso_z(end) if end else None,
        "start_time_zone": start_zone,
        "end_time_zone": end_zone,
        "start_local": _local(start, start_zone),
        "end_local": _local(end, end_zone),
        "is_all_day": bool(row.get("isAllDay")),
        "is_cancelled": bool(row.get("isCancelled")),
        "is_organizer": bool(row.get("isOrganizer")),
        "organizer": address_out(row.get("organizer")),
        "attendees": attendees_out(row.get("attendees")),
        "location": str(mapping(row.get("location")).get("displayName") or ""),
        "is_online_meeting": bool(row.get("isOnlineMeeting")),
        "online_meeting_provider": str(row.get("onlineMeetingProvider") or "unknown"),
        "join_url": str(mapping(row.get("onlineMeeting")).get("joinUrl") or "") or None,
        "show_as": str(row.get("showAs") or "unknown"),
        "response": str(mapping(row.get("responseStatus")).get("response") or "none"),
        "web_link": str(row.get("webLink") or ""),
        "reminder_enabled": bool(row.get("isReminderOn")),
        "reminder_minutes": reminder if isinstance(reminder, int) else None,
        "importance": str(row.get("importance") or "normal"),
        "sensitivity": str(row.get("sensitivity") or "normal"),
    }


def calendar_out(row: Mapping[str, Any]) -> dict[str, Any]:
    providers = row.get("allowedOnlineMeetingProviders")
    allowed = (
        [str(item) for item in providers if isinstance(item, str)]
        if isinstance(providers, list)
        else []
    )
    owner = mapping(row.get("owner"))
    return {
        "calendar_id": str(row.get("id") or ""),
        "name": str(row.get("name") or ""),
        "color": str(row.get("color") or ""),
        "is_default": bool(row.get("isDefaultCalendar")),
        "can_edit": bool(row.get("canEdit")),
        "owner": (
            {"address": str(owner.get("address") or ""), "name": str(owner.get("name") or "")}
            if owner
            else None
        ),
        "allowed_online_meeting_providers": allowed,
        "default_online_meeting_provider": str(row.get("defaultOnlineMeetingProvider") or ""),
        "teams_meetings_allowed": TEAMS_PROVIDER in allowed,
    }


__all__ = [
    "CALENDAR_NOT_FOUND",
    "EVENT_KEY_PROPERTY_ID",
    "EVENT_NOT_FOUND",
    "EVENT_SELECT",
    "MAX_LIST_DAYS",
    "MAX_SCHEDULE_ADDRESSES",
    "MAX_SCHEDULE_DAYS",
    "PREFER_TEXT_BODY",
    "PREFER_UTC",
    "TEAMS_PROVIDER",
    "as_read_failure",
    "attendees_input",
    "attendees_out",
    "body_input",
    "calendar_id_input",
    "calendar_out",
    "calendar_path",
    "client_for",
    "event_id_input",
    "event_out",
    "event_path",
    "events_path",
    "graph_time",
    "iso_local",
    "iso_z",
    "moment_input",
    "parse_graph_time",
    "read_calendar",
    "read_event",
    "window_input",
]
