"""Event writes: create, update, cancel and respond (task D2).

Every write uses ``retry="none"``; the transport never repeats it. An ambiguous write
(no response, 5xx, timeout) is confirmed by reading, never by writing again:

- ``create_event`` carries the step key in a single-value extended property (the v0.1
  mail property set) and as ``transactionId``. After an ambiguous POST the event
  collection is searched with the documented ``singleValueExtendedProperties/Any(...)``
  filter alone; the start time is checked here, on the returned events. Exactly one
  match with the requested start is success; none or several is ``timeout_unknown``.
  The POST is never repeated.
- ``update_event`` re-reads the event: every given field equal is success.
- ``cancel_event`` re-reads the event: gone or ``isCancelled`` is success.
- ``respond_to_event`` re-reads the attendee's copy: the requested response (or, for a
  decline, the event gone) is success.
"""

from __future__ import annotations

import re
import time
from collections.abc import Callable, Mapping
from datetime import UTC, datetime, tzinfo
from typing import Any

from m365.connection import mapping
from m365.errors import M365Error
from m365.runtime import OperationContext
from mail.common import (
    Mailbox,
    bool_input,
    int_input,
    not_found_as,
    odata_string,
    text_input,
)
from mail.common import (
    address_out as _address_out,
)

from .common import (
    CALENDAR_NOT_FOUND,
    EVENT_KEY_PROPERTY_ID,
    EVENT_NOT_FOUND,
    EVENT_RESPONSE_LIMIT,
    EVENT_SELECT,
    MAX_COMMENT_CHARS,
    MAX_LOCATION_CHARS,
    MAX_REMINDER_MINUTES,
    MAX_SUBJECT_CHARS,
    PREFER_TEXT_BODY,
    PREFER_UTC,
    TEAMS_PROVIDER,
    attendees_input,
    body_input,
    calendar_id_input,
    calendar_out,
    client_for,
    event_id_input,
    event_out,
    event_path,
    events_path,
    graph_time,
    iso_z,
    moment_input,
    read_calendar,
    read_event,
)
from .zones import zone_input

LOOKUP_ATTEMPTS = 3
LOOKUP_DELAY_SECONDS = 2.0
MAX_KEY_MATCHES = 10

#: Indirection point for tests.
sleep: Callable[[float], None] = time.sleep

_EVENT_FIELDS = frozenset(
    {
        "subject",
        "body_text",
        "body_html",
        "start",
        "end",
        "time_zone",
        "attendees",
        "optional_attendees",
        "location",
        "reminder_minutes",
        "reminder_enabled",
    }
)
CREATE_FIELDS = _EVENT_FIELDS | {"calendar_id", "online_meeting"}
UPDATE_FIELDS = _EVENT_FIELDS | {"event_id"}
CANCEL_FIELDS = frozenset({"event_id", "comment"})
RESPOND_FIELDS = frozenset({"event_id", "response", "send_response", "comment"})
#: respond_to_event action → the attendee's ``responseStatus.response`` afterwards.
RESPONSES = {
    "accept": "accepted",
    "tentativelyAccept": "tentativelyAccepted",
    "decline": "declined",
}


# -- input -----------------------------------------------------------------------------


def _times(raw: Mapping[str, Any]) -> tuple[datetime, datetime, str, tzinfo]:
    zone_name: str = "UTC"
    zone: tzinfo = UTC
    explicit_zone: tzinfo | None = None
    if raw.get("time_zone") not in (None, ""):
        zone_name, zone = zone_input(raw.get("time_zone"), "time_zone")
        explicit_zone = zone
    start = moment_input(raw.get("start"), "start", explicit_zone)
    end = moment_input(raw.get("end"), "end", explicit_zone)
    if end <= start:
        raise M365Error("invalid_payload", "end must be later than start")
    return start, end, zone_name, zone


def plan_event(raw: Mapping[str, Any], *, create: bool) -> tuple[dict[str, Any], dict[str, Any]]:
    """The Graph body and the expected values used to confirm an ambiguous write."""
    body: dict[str, Any] = {}
    expected: dict[str, Any] = {}
    if create or raw.get("subject") is not None:
        subject = text_input(
            raw.get("subject"), "subject", max_length=MAX_SUBJECT_CHARS, required=create
        )
        body["subject"] = expected["subject"] = subject
    content = body_input(raw)
    if content is not None:
        body["body"] = content
        expected["body_html" if content["contentType"] == "HTML" else "body_text"] = content[
            "content"
        ]
    has_start = raw.get("start") not in (None, "")
    has_end = raw.get("end") not in (None, "")
    if create or has_start or has_end:
        if not (has_start and has_end):
            raise M365Error("invalid_payload", "Set both start and end")
        start, end, zone_name, zone = _times(raw)
        body["start"] = graph_time(start, zone_name, zone)
        body["end"] = graph_time(end, zone_name, zone)
        expected["start_utc"], expected["end_utc"] = iso_z(start), iso_z(end)
    elif raw.get("time_zone") not in (None, ""):
        raise M365Error("invalid_payload", "time_zone is only used together with start and end")
    attendees = attendees_input(raw)
    if attendees is not None:
        body["attendees"] = attendees
        expected["attendees"] = _attendee_set(attendees)
    if raw.get("location") is not None:
        location = text_input(raw.get("location"), "location", max_length=MAX_LOCATION_CHARS)
        body["location"] = {"displayName": location}
        expected["location"] = location
    enabled = bool_input(raw.get("reminder_enabled"), "reminder_enabled")
    if raw.get("reminder_minutes") is not None:
        minutes = int_input(
            raw.get("reminder_minutes"),
            "reminder_minutes",
            default=15,
            minimum=0,
            maximum=MAX_REMINDER_MINUTES,
        )
        if enabled is False:
            raise M365Error("invalid_payload", "reminder_minutes needs the reminder enabled")
        body["reminderMinutesBeforeStart"] = expected["reminder_minutes"] = minutes
        enabled = True
    if enabled is not None:
        body["isReminderOn"] = expected["reminder_enabled"] = enabled
    return body, expected


def _attendee_set(rows: list[dict[str, Any]]) -> set[tuple[str, str]]:
    return {
        (str(mapping(row.get("emailAddress")).get("address") or "").lower(), str(row.get("type")))
        for row in rows
    }


def _comment(raw: Mapping[str, Any]) -> str:
    return text_input(
        raw.get("comment"), "comment", max_length=MAX_COMMENT_CHARS, allow_newlines=True
    )


# -- Graph calls ---------------------------------------------------------------------------


def _write(
    client: Mailbox,
    method: str,
    path: str,
    *,
    json_body: Mapping[str, Any] | None,
    purpose: str,
    ok_statuses: tuple[int, ...] = (200, 201, 202, 204),
) -> dict[str, Any]:
    response = client.transport.request(
        method,
        f"{client.base}{path}",
        json_body=json_body,
        raw_body=None if json_body is not None else b"",
        headers=Mailbox.headers(PREFER_UTC),
        response_limit=EVENT_RESPONSE_LIMIT,
        retry="none",
        purpose=purpose,
        ok_statuses=ok_statuses,
    )
    try:
        return response.json()
    except M365Error:
        return {}


def _reread(client: Mailbox, event_id: str, *, with_body: bool = False) -> dict[str, Any] | None:
    """The event after an ambiguous write; ``None`` when it no longer exists."""
    try:
        return client.get(
            event_path(event_id),
            query={"$select": EVENT_SELECT + (",body" if with_body else "")},
            prefer=(PREFER_UTC, PREFER_TEXT_BODY),
            response_limit=EVENT_RESPONSE_LIMIT,
            purpose="Outlook event confirmation",
        )
    except M365Error as exc:
        if exc.code == "item_not_found":
            return None
        raise


def _unknown(message: str, cause: M365Error) -> M365Error:
    return M365Error("timeout_unknown", message, http_status=cause.http_status, ambiguous=True)


def _success(confirmation: str, *, changed: bool = True, **fields: Any) -> dict[str, Any]:
    return {
        **fields,
        "confirmation": confirmation,
        "changed": changed,
        "external_effect_status": "succeeded",
        "definitely_no_external_effect": not changed,
    }


# -- create_event -----------------------------------------------------------------------------


def create_event(ctx: OperationContext) -> dict[str, Any]:
    key = ctx.idempotency_key
    raw = ctx.input
    calendar_id = calendar_id_input(raw.get("calendar_id"))
    body, expected = plan_event(raw, create=True)
    online = bool_input(raw.get("online_meeting"), "online_meeting", default=False)
    client = client_for(ctx.transport, ctx.connection)
    collection = events_path(calendar_id)
    if online:
        calendar = calendar_out(read_calendar(client, calendar_id, purpose="Outlook calendar"))
        if not calendar["teams_meetings_allowed"]:
            raise M365Error(
                "online_meeting_unsupported",
                "This calendar does not allow Teams meetings (teamsForBusiness)",
            )
        body["isOnlineMeeting"] = True
        body["onlineMeetingProvider"] = TEAMS_PROVIDER
    # A re-run of the same step finds its event instead of creating a second one.
    existing = _find_by_key(client, collection, key)
    if existing:
        return _resolve_matches(existing, expected, "already_exists", cause=None)
    body["transactionId"] = key
    body["singleValueExtendedProperties"] = [{"id": EVENT_KEY_PROPERTY_ID, "value": key}]
    try:
        with not_found_as("calendar_not_found", CALENDAR_NOT_FOUND):
            row = _write(client, "POST", collection, json_body=body, purpose="Outlook event create")
    except M365Error as exc:
        if not exc.ambiguous:
            raise
        return _confirm_create(client, collection, key, expected, exc)
    return _created(event_out(row), "created")


def _created(event: dict[str, Any], confirmation: str) -> dict[str, Any]:
    return _success(
        confirmation,
        event_id=event["event_id"],
        join_url=event["join_url"],
        event=event,
    )


def _key_filter(key: str) -> str:
    return (
        f"singleValueExtendedProperties/Any(ep: ep/id eq {odata_string(EVENT_KEY_PROPERTY_ID)} "
        f"and ep/value eq {odata_string(key)})"
    )


def _find_by_key(client: Mailbox, collection: str, key: str) -> list[dict[str, Any]]:
    """Events carrying the step key. Only the documented property filter is sent."""
    page = client.get(
        collection,
        query={
            "$filter": _key_filter(key),
            "$expand": (
                "singleValueExtendedProperties($filter=id eq "
                f"{odata_string(EVENT_KEY_PROPERTY_ID)})"
            ),
            "$select": EVENT_SELECT,
            "$top": str(MAX_KEY_MATCHES),
        },
        prefer=(PREFER_UTC,),
        response_limit=EVENT_RESPONSE_LIMIT,
        purpose="Outlook event lookup",
    )
    rows = page.get("value")
    if not isinstance(rows, list):
        raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid event list")
    # Defence in depth: keep only rows whose expanded property really carries the key.
    return [dict(row) for row in rows if isinstance(row, Mapping) and _carries_key(row, key)]


def _carries_key(row: Mapping[str, Any], key: str) -> bool:
    for prop in row.get("singleValueExtendedProperties") or []:
        prop = mapping(prop)
        if (
            str(prop.get("id") or "").lower() == EVENT_KEY_PROPERTY_ID.lower()
            and prop.get("value") == key
        ):
            return True
    return False


def _confirm_create(
    client: Mailbox, collection: str, key: str, expected: Mapping[str, Any], cause: M365Error
) -> dict[str, Any]:
    for attempt in range(LOOKUP_ATTEMPTS):
        try:
            matches = _find_by_key(client, collection, key)
        except M365Error:
            matches = []
        if matches:
            return _resolve_matches(matches, expected, "extended_property", cause=cause)
        if attempt + 1 < LOOKUP_ATTEMPTS:
            sleep(LOOKUP_DELAY_SECONDS)
    raise _unknown(
        "Outlook did not confirm the event create and no event with this step's key was "
        "found; check the calendar before creating it again",
        cause,
    )


def _resolve_matches(
    matches: list[dict[str, Any]],
    expected: Mapping[str, Any],
    confirmation: str,
    *,
    cause: M365Error | None,
) -> dict[str, Any]:
    events = [event_out(row) for row in matches]
    ids = ", ".join(event["event_id"] for event in events)
    dated = [event for event in events if event["start_utc"] == expected.get("start_utc")]
    if len(events) == 1 and len(dated) == 1:
        return _created(dated[0], confirmation)
    if len(events) > 1:
        message = f"Several events carry this step's key: {ids}"
    else:
        message = f"The event with this step's key ({ids}) has a different start time"
    raise M365Error(
        "timeout_unknown",
        f"{message}; check the calendar",
        http_status=cause.http_status if cause else None,
        ambiguous=True,
    )


# -- update_event -------------------------------------------------------------------------------


def update_event(ctx: OperationContext) -> dict[str, Any]:
    event_id = event_id_input(ctx.input.get("event_id"))
    body, expected = plan_event(ctx.input, create=False)
    if not body:
        raise M365Error("invalid_payload", "Set at least one field to change")
    client = client_for(ctx.transport, ctx.connection)
    try:
        with not_found_as("event_not_found", EVENT_NOT_FOUND):
            row = _write(
                client,
                "PATCH",
                event_path(event_id),
                json_body=body,
                purpose="Outlook event update",
            )
    except M365Error as exc:
        if not exc.ambiguous:
            raise
        return _confirm_update(client, event_id, expected, exc)
    event = event_out(row) if row.get("id") else {"event_id": event_id}
    return _success(
        "patched", event_id=event["event_id"] or event_id, updated_fields=sorted(expected)
    )


def _confirm_update(
    client: Mailbox, event_id: str, expected: Mapping[str, Any], cause: M365Error
) -> dict[str, Any]:
    try:
        row = _reread(client, event_id, with_body="body_text" in expected)
    except M365Error as exc:
        raise _unknown(
            f"Outlook did not confirm the update of event {event_id}; check the event", cause
        ) from exc
    if row is None:
        raise _unknown(f"Event {event_id} is gone after an unconfirmed update", cause)
    different = mismatched_fields(row, expected)
    if different:
        raise _unknown(
            f"Outlook did not confirm the update of event {event_id}; these fields differ: "
            f"{', '.join(different)}",
            cause,
        )
    return _success(
        "reread", event_id=str(row.get("id") or event_id), updated_fields=sorted(expected)
    )


def _normalized(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def mismatched_fields(row: Mapping[str, Any], expected: Mapping[str, Any]) -> list[str]:
    """Fields of ``expected`` that the re-read event does not show (sorted)."""
    event = event_out(row)
    actual: dict[str, Any] = {
        "subject": event["subject"],
        "start_utc": event["start_utc"],
        "end_utc": event["end_utc"],
        "location": event["location"],
        "reminder_enabled": event["reminder_enabled"],
        "reminder_minutes": event["reminder_minutes"],
        "attendees": {(row["address"].lower(), row["type"]) for row in event["attendees"]},
    }
    different = []
    for name, value in expected.items():
        if name == "body_html":
            # Outlook rewrites stored HTML, so an HTML body can never be compared.
            different.append(name)
        elif name == "body_text":
            content = str(mapping(row.get("body")).get("content") or "")
            if _normalized(content) != _normalized(str(value)):
                different.append(name)
        elif actual.get(name) != value:
            different.append(name)
    return sorted(different)


# -- cancel_event ---------------------------------------------------------------------------------


def cancel_event(ctx: OperationContext) -> dict[str, Any]:
    event_id = event_id_input(ctx.input.get("event_id"))
    comment = _comment(ctx.input)
    client = client_for(ctx.transport, ctx.connection)
    try:
        row = read_event(client, event_id, purpose="Outlook event")
    except M365Error as exc:
        if exc.code != "event_not_found":
            raise
        return _success("already_gone", changed=False, event_id=event_id, attendees_notified="none")
    if row.get("isCancelled") is True:
        return _success(
            "already_cancelled", changed=False, event_id=event_id, attendees_notified="none"
        )
    if row.get("isOrganizer") is not True:
        raise M365Error(
            "not_organizer",
            "Only the organizer can cancel this event; use Respond to event with decline",
        )
    organizer = (mapping(_address_out(row.get("organizer"))).get("address") or "").lower()
    attendees = [
        item for item in event_out(row)["attendees"] if item["address"].lower() != organizer
    ]
    try:
        with not_found_as("event_not_found", EVENT_NOT_FOUND):
            if attendees:
                _write(
                    client,
                    "POST",
                    f"{event_path(event_id)}/cancel",
                    json_body={"comment": comment} if comment else {},
                    purpose="Outlook event cancel",
                )
            else:
                # An appointment without attendees has nobody to notify: delete it.
                _write(
                    client,
                    "DELETE",
                    event_path(event_id),
                    json_body=None,
                    purpose="Outlook event delete",
                )
    except M365Error as exc:
        if not exc.ambiguous:
            raise
        return _confirm_cancel(client, event_id, exc)
    return _success(
        "cancel_accepted" if attendees else "deleted",
        event_id=event_id,
        attendees_notified="requested" if attendees else "none",
    )


def _confirm_cancel(client: Mailbox, event_id: str, cause: M365Error) -> dict[str, Any]:
    try:
        row = _reread(client, event_id)
    except M365Error as exc:
        raise _unknown(
            f"Outlook did not confirm the cancellation of event {event_id}; check the calendar",
            cause,
        ) from exc
    if row is None or row.get("isCancelled") is True:
        return _success("reread", event_id=event_id, attendees_notified="unknown")
    raise _unknown(
        f"Outlook did not confirm the cancellation of event {event_id}; it is still active",
        cause,
    )


# -- respond_to_event -------------------------------------------------------------------------------


def respond_to_event(ctx: OperationContext) -> dict[str, Any]:
    event_id = event_id_input(ctx.input.get("event_id"))
    action = ctx.input.get("response")
    if action not in RESPONSES:
        raise M365Error("invalid_payload", "response must be accept, tentativelyAccept or decline")
    send = bool_input(ctx.input.get("send_response"), "send_response", default=True)
    comment = _comment(ctx.input)
    target = RESPONSES[str(action)]
    client = client_for(ctx.transport, ctx.connection)
    row = read_event(client, event_id, purpose="Outlook event")
    if row.get("isOrganizer") is True:
        raise M365Error("invalid_payload", "The organizer cannot respond to their own event")
    current = str(mapping(row.get("responseStatus")).get("response") or "")
    if current == target:
        return _success(
            "already_responded",
            changed=False,
            event_id=event_id,
            response=target,
            response_sent="not_sent",
        )
    request: dict[str, Any] = {"sendResponse": bool(send)}
    if comment:
        request["comment"] = comment
    try:
        with not_found_as("event_not_found", EVENT_NOT_FOUND):
            _write(
                client,
                "POST",
                f"{event_path(event_id)}/{action}",
                json_body=request,
                purpose="Outlook event response",
            )
    except M365Error as exc:
        if not exc.ambiguous:
            raise
        return _confirm_response(client, event_id, target, bool(send), exc)
    return _success(
        "response_accepted",
        event_id=event_id,
        response=target,
        response_sent="requested" if send else "not_requested",
    )


def _confirm_response(
    client: Mailbox, event_id: str, target: str, send: bool, cause: M365Error
) -> dict[str, Any]:
    try:
        row = _reread(client, event_id)
    except M365Error as exc:
        raise _unknown(
            f"Outlook did not confirm the response to event {event_id}; check the event", cause
        ) from exc
    if row is None:
        confirmed = target == "declined"
    else:
        confirmed = str(mapping(row.get("responseStatus")).get("response") or "") == target
    if not confirmed:
        raise _unknown(
            f"Outlook did not confirm the response to event {event_id}; check the event", cause
        )
    return _success(
        "reread",
        event_id=event_id,
        response=target,
        response_sent="unknown" if send else "not_requested",
    )


__all__ = [
    "CANCEL_FIELDS",
    "CREATE_FIELDS",
    "RESPOND_FIELDS",
    "RESPONSES",
    "UPDATE_FIELDS",
    "cancel_event",
    "create_event",
    "mismatched_fields",
    "plan_event",
    "respond_to_event",
    "update_event",
]
