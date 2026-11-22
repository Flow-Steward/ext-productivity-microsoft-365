"""Shared fakes for the calendar tests: an in-memory Outlook calendar behind a fake
transport. This module holds no tests.

Faults are injected per action (``create``, ``patch``, ``cancel``, ``delete``,
``respond``, ``lookup``, ``schedule``): ``lost`` (the effect happens, the response is
lost), ``lost_twice`` (two effects, response lost), ``dropped`` (no effect, ambiguous),
``ok`` (normal, to script a later fault), ``refused`` (a confirmed 400), ``throttled`` (a confirmed 429) and ``unavailable`` (a
confirmed GET failure, 503).
"""

from __future__ import annotations

import json
import re
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import UTC, datetime
from typing import Any
from urllib.parse import unquote

from calendar_ops.common import EVENT_KEY_PROPERTY_ID
from calendar_ops.zones import lookup_zone
from m365.errors import M365Error
from m365.transport import GraphResponse

STEP_EXECUTION_ID = "6a1b2c3d-4e5f-4a6b-8c7d-000000000201"
MAILBOX = "cmr@acme.example"
APP_BASE = f"/users/{MAILBOX}"
USER_CONNECTION: dict[str, Any] = {
    "connection_id": "conn-acme-calendar",
    "connection_type_id": "m365_user",
    "config": {},
    "secrets": {"access_token": "opaque-token"},
}
APP_CONNECTION: dict[str, Any] = {
    "connection_id": "conn-acme-calendar",
    "connection_type_id": "m365_calendar_app",
    "config": {
        "tenant_id": "acme.onmicrosoft.com",
        "client_id": "11111111-2222-3333-4444-555555555555",
        "mailbox": MAILBOX,
        "control_mailbox": "ceo@acme.example",
    },
    "secrets": {"client_secret": "app-secret"},
}
#: (connection, Graph base) for parametrised tests over both auth modes.
MODES = [(USER_CONNECTION, "/me"), (APP_CONNECTION, APP_BASE)]
MODE_IDS = ["user", "app"]


def action_payload(
    action_id: str,
    operation_input: dict[str, Any] | None = None,
    *,
    connection: dict[str, Any] | None = None,
    test_mode: bool = False,
    step_execution_id: str = STEP_EXECUTION_ID,
) -> dict[str, Any]:
    return {
        "contract_version": "extension_host_v2",
        "mode": "action",
        "runtime_context": {
            "group_id": None,
            "for_each_id": None,
            "step_id": "book_meeting",
            "step_execution_id": step_execution_id,
            "test_mode": test_mode,
        },
        "action": {
            "action_id": action_id,
            "input": {"connection_ref": "conn-acme-calendar", **(operation_input or {})},
            "target": {"connection": dict(connection or USER_CONNECTION)},
        },
    }


@dataclass
class Call:
    method: str
    url: str
    query: dict[str, Any]
    json_body: Any
    headers: dict[str, str]
    retry: str


@dataclass
class Event:
    id: str
    calendar_id: str
    subject: str
    start: datetime
    end: datetime
    zone: str = "UTC"
    body: str = ""
    attendees: list[dict[str, Any]] = field(default_factory=list)
    location: str = ""
    is_organizer: bool = True
    organizer: str = MAILBOX
    is_cancelled: bool = False
    response: str = "organizer"
    reminder_on: bool = True
    reminder_minutes: int = 15
    online: bool = False
    props: dict[str, str] = field(default_factory=dict)
    transaction_id: str = ""


def ambiguous(what: str) -> M365Error:
    return M365Error("timeout_unknown", f"{what} did not complete", http_status=504, ambiguous=True)


def utc(text: str) -> datetime:
    return datetime.fromisoformat(text.replace("Z", "+00:00")).astimezone(UTC)


def graph_utc(moment: datetime) -> dict[str, str]:
    return {"dateTime": moment.strftime("%Y-%m-%dT%H:%M:%S.0000000"), "timeZone": "UTC"}


def _json(value: Any, status: int = 200) -> GraphResponse:
    return GraphResponse(status, {}, json.dumps(value).encode("utf-8"))


class FakeCalendar:
    """An in-memory mailbox with two calendars, reachable at ``base``."""

    def __init__(self, base: str = "/me") -> None:
        self.base = base
        self.calls: list[Call] = []
        self.faults: dict[str, list[str]] = {}
        self.counter = 0
        self.events: dict[str, Event] = {}
        self.calendars: dict[str, dict[str, Any]] = {
            "CAL-DEFAULT": {
                "id": "CAL-DEFAULT",
                "name": "Calendar",
                "color": "auto",
                "isDefaultCalendar": True,
                "canEdit": True,
                "owner": {"name": "Ana Acme", "address": MAILBOX},
                "allowedOnlineMeetingProviders": ["teamsForBusiness"],
                "defaultOnlineMeetingProvider": "teamsForBusiness",
            },
            "CAL-HOLIDAYS": {
                "id": "CAL-HOLIDAYS",
                "name": "Holidays",
                "isDefaultCalendar": False,
                "canEdit": True,
                "owner": {"name": "Ana Acme", "address": MAILBOX},
                "allowedOnlineMeetingProviders": [],
                "defaultOnlineMeetingProvider": "unknown",
            },
        }
        #: The ``value`` list getSchedule answers with.
        self.schedule: list[dict[str, Any]] = []
        #: When set, the lookup ignores the property filter (returns every event).
        self.lookup_ignores_filter = False

    # -- scripting ----------------------------------------------------------------------

    def fault(self, op: str, *modes: str) -> None:
        self.faults.setdefault(op, []).extend(modes)

    def add_event(self, calendar_id: str = "CAL-DEFAULT", **fields: Any) -> Event:
        fields.setdefault("subject", "Existing")
        fields.setdefault("start", utc("2026-11-25T09:00:00Z"))
        fields.setdefault("end", utc("2026-11-25T10:00:00Z"))
        event = Event(id=self._id(), calendar_id=calendar_id, **fields)
        self.events[event.id] = event
        return event

    def writes(self) -> list[Call]:
        return [
            call for call in self.calls if call.method != "GET" and "getSchedule" not in call.url
        ]

    def remaining_seconds(self) -> float:
        return 200.0

    # -- transport ------------------------------------------------------------------------

    def request(
        self,
        method: str,
        path_or_url: str,
        *,
        query: dict[str, Any] | None = None,
        json_body: Any = None,
        raw_body: bytes | None = None,
        content_type: str | None = None,
        headers: dict[str, str] | None = None,
        response_limit: int = 4 * 1024 * 1024,
        retry: str,
        purpose: str,
        authorized: bool = True,
        ok_statuses: Any = (200, 201, 202, 204),
    ) -> GraphResponse:
        assert retry in {"read", "idempotent", "none"}
        if method == "GET":
            assert retry == "read"
        call = Call(method, path_or_url, dict(query or {}), json_body, dict(headers or {}), retry)
        self.calls.append(call)
        assert authorized
        path = path_or_url
        if path.startswith("https://graph.microsoft.com/v1.0"):
            return self._next_page(path)
        if path == "/me" and method == "GET":
            return _json({"displayName": "Ana Acme", "mail": MAILBOX, "userPrincipalName": MAILBOX})
        assert path.startswith(self.base + "/"), path
        rest = path[len(self.base) :]
        for pattern, verb, handler in self._routes():
            match = re.fullmatch(pattern, rest)
            if match and verb == method:
                return handler(
                    *[unquote(g) for g in match.groups()], query=query or {}, body=json_body
                )
        raise AssertionError(f"unrouted {method} {path}")

    def _routes(self) -> list[tuple[str, str, Callable[..., GraphResponse]]]:
        return [
            (r"/calendar/getSchedule", "POST", self._get_schedule),
            (r"/calendar/calendarView", "GET", lambda **kw: self._view("CAL-DEFAULT", **kw)),
            (r"/calendars/([^/]+)/calendarView", "GET", self._view),
            (r"/calendar", "GET", lambda **kw: self._calendar("CAL-DEFAULT")),
            (r"/calendars", "GET", self._list_calendars),
            (r"/calendars/([^/]+)", "GET", lambda cid, **kw: self._calendar(cid)),
            (r"/events", "GET", lambda **kw: self._lookup("CAL-DEFAULT", **kw)),
            (r"/calendars/([^/]+)/events", "GET", self._lookup),
            (r"/events", "POST", lambda **kw: self._create("CAL-DEFAULT", **kw)),
            (r"/calendars/([^/]+)/events", "POST", self._create),
            (r"/events/([^/]+)/cancel", "POST", self._cancel),
            (r"/events/([^/]+)/(accept|tentativelyAccept|decline)", "POST", self._respond),
            (r"/events/([^/]+)", "GET", self._get_event),
            (r"/events/([^/]+)", "PATCH", self._patch),
            (r"/events/([^/]+)", "DELETE", self._delete),
        ]

    # -- helpers ----------------------------------------------------------------------------

    def _id(self) -> str:
        self.counter += 1
        return f"AAMkEVT{self.counter:04d}"

    def _apply(self, op: str, effect: Callable[[], GraphResponse]) -> GraphResponse:
        modes = self.faults.get(op)
        mode = modes.pop(0) if modes else None
        if mode is None or mode == "ok":
            return effect()
        if mode == "lost":
            effect()
            raise ambiguous(op)
        if mode == "lost_twice":
            effect()
            effect()
            raise ambiguous(op)
        if mode == "dropped":
            raise ambiguous(op)
        if mode == "refused":
            raise M365Error("invalid_payload", f"{op} refused", http_status=400)
        if mode == "throttled":
            raise M365Error("rate_limited", f"{op} throttled", http_status=429)
        if mode == "unavailable":
            raise M365Error("provider_unavailable", f"{op} unavailable", http_status=503)
        raise AssertionError(mode)

    def _event(self, event_id: str) -> Event:
        event = self.events.get(event_id)
        if event is None:
            raise M365Error("item_not_found", "not found", http_status=404)
        return event

    def row(self, event: Event, *, expand: bool = False) -> dict[str, Any]:
        row: dict[str, Any] = {
            "id": event.id,
            "iCalUId": f"UID-{event.id}",
            "type": "singleInstance",
            "subject": event.subject,
            "bodyPreview": event.body[:100],
            "body": {"contentType": "text", "content": event.body},
            "start": graph_utc(event.start),
            "end": graph_utc(event.end),
            "originalStartTimeZone": event.zone,
            "originalEndTimeZone": event.zone,
            "isAllDay": False,
            "isCancelled": event.is_cancelled,
            "isOrganizer": event.is_organizer,
            "organizer": {"emailAddress": {"address": event.organizer, "name": "Organizer"}},
            "attendees": event.attendees,
            "location": {"displayName": event.location},
            "isOnlineMeeting": event.online,
            "onlineMeetingProvider": "teamsForBusiness" if event.online else "unknown",
            "onlineMeeting": {"joinUrl": f"https://teams.example/{event.id}"}
            if event.online
            else None,
            "showAs": "busy",
            "responseStatus": {"response": event.response, "time": "0001-01-01T00:00:00Z"},
            "webLink": f"https://outlook.example/{event.id}",
            "isReminderOn": event.reminder_on,
            "reminderMinutesBeforeStart": event.reminder_minutes,
            "importance": "normal",
            "sensitivity": "normal",
        }
        if expand:
            row["singleValueExtendedProperties"] = [
                {"id": prop_id, "value": value} for prop_id, value in event.props.items()
            ]
        return row

    # -- calendars --------------------------------------------------------------------------

    def _calendar(self, calendar_id: str) -> GraphResponse:
        if calendar_id not in self.calendars:
            raise M365Error("item_not_found", "calendar", http_status=404)
        return _json(self.calendars[calendar_id])

    def _list_calendars(self, **_: Any) -> GraphResponse:
        return _json({"value": list(self.calendars.values())})

    def _view(self, calendar_id: str, *, query: dict[str, Any], **_: Any) -> GraphResponse:
        if calendar_id not in self.calendars:
            raise M365Error("item_not_found", "calendar", http_status=404)
        start, end = utc(query["startDateTime"]), utc(query["endDateTime"])
        rows = sorted(
            (
                event
                for event in self.events.values()
                if event.calendar_id == calendar_id and event.start < end and event.end > start
            ),
            key=lambda event: event.start,
        )
        return self._page(
            [self.row(event) for event in rows], int(query.get("$top", 10)), 0, calendar_id, query
        )

    def _page(
        self,
        rows: list[dict[str, Any]],
        top: int,
        skip: int,
        calendar_id: str,
        query: dict[str, Any],
    ) -> GraphResponse:
        chunk = rows[skip : skip + top]
        body: dict[str, Any] = {"value": chunk}
        if skip + top < len(rows):
            body["@odata.nextLink"] = (
                f"https://graph.microsoft.com/v1.0{self.base}/calendars/{calendar_id}/calendarView"
                f"?startDateTime={query['startDateTime']}&endDateTime={query['endDateTime']}"
                f"&$top={top}&$skip={skip + top}"
            )
        return _json(body)

    def _next_page(self, link: str) -> GraphResponse:
        match = re.search(
            r"/calendars/([^/]+)/calendarView\?startDateTime=([^&]+)&endDateTime=([^&]+)&\$top=(\d+)&\$skip=(\d+)",
            link,
        )
        assert match, link
        calendar_id, start, end, top, skip = match.groups()
        query = {"startDateTime": start, "endDateTime": end}
        rows = sorted(
            (
                event
                for event in self.events.values()
                if event.calendar_id == calendar_id
                and event.start < utc(end)
                and event.end > utc(start)
            ),
            key=lambda event: event.start,
        )
        return self._page(
            [self.row(event) for event in rows], int(top), int(skip), calendar_id, query
        )

    # -- events -------------------------------------------------------------------------------

    def _lookup(self, calendar_id: str, *, query: dict[str, Any], **_: Any) -> GraphResponse:
        def effect() -> GraphResponse:
            match = re.search(r"ep/value eq '((?:[^']|'')*)'", query.get("$filter", ""))
            assert match, query
            assert " and start" not in query["$filter"], (
                "no date filter next to the property filter"
            )
            key = match.group(1).replace("''", "'")
            rows = [
                self.row(event, expand=True)
                for event in self.events.values()
                if event.calendar_id == calendar_id
                and (self.lookup_ignores_filter or event.props.get(EVENT_KEY_PROPERTY_ID) == key)
            ]
            return _json({"value": rows})

        return self._apply("lookup", effect)

    def _create(self, calendar_id: str, *, body: dict[str, Any], **_: Any) -> GraphResponse:
        if calendar_id not in self.calendars:
            raise M365Error("item_not_found", "calendar", http_status=404)
        created: list[Event] = []

        def effect() -> GraphResponse:
            zone = lookup_zone(body["start"]["timeZone"])
            assert zone is not None
            start = datetime.fromisoformat(body["start"]["dateTime"]).replace(tzinfo=zone)
            end = datetime.fromisoformat(body["end"]["dateTime"]).replace(tzinfo=zone)
            event = Event(
                id=self._id(),
                calendar_id=calendar_id,
                subject=body.get("subject", ""),
                start=start.astimezone(UTC),
                end=end.astimezone(UTC),
                zone=body["start"]["timeZone"],
                body=(body.get("body") or {}).get("content", ""),
                attendees=[
                    {**row, "status": {"response": "none"}} for row in body.get("attendees", [])
                ],
                location=(body.get("location") or {}).get("displayName", ""),
                online=bool(body.get("isOnlineMeeting")),
                props={
                    row["id"]: row["value"] for row in body.get("singleValueExtendedProperties", [])
                },
                transaction_id=body.get("transactionId", ""),
            )
            if "isReminderOn" in body:
                event.reminder_on = body["isReminderOn"]
            if "reminderMinutesBeforeStart" in body:
                event.reminder_minutes = body["reminderMinutesBeforeStart"]
            self.events[event.id] = event
            created.append(event)
            return _json(self.row(event), 201)

        return self._apply("create", effect)

    def _get_event(self, event_id: str, **_: Any) -> GraphResponse:
        return _json(self.row(self._event(event_id)))

    def _patch(self, event_id: str, *, body: dict[str, Any], **_: Any) -> GraphResponse:
        event = self._event(event_id)

        def effect() -> GraphResponse:
            if "subject" in body:
                event.subject = body["subject"]
            if "body" in body:
                event.body = body["body"]["content"]
            if "location" in body:
                event.location = body["location"]["displayName"]
            if "start" in body:
                zone = lookup_zone(body["start"]["timeZone"])
                assert zone is not None
                event.start = (
                    datetime.fromisoformat(body["start"]["dateTime"])
                    .replace(tzinfo=zone)
                    .astimezone(UTC)
                )
                event.end = (
                    datetime.fromisoformat(body["end"]["dateTime"])
                    .replace(tzinfo=zone)
                    .astimezone(UTC)
                )
                event.zone = body["start"]["timeZone"]
            if "attendees" in body:
                event.attendees = [
                    {**row, "status": {"response": "none"}} for row in body["attendees"]
                ]
            if "isReminderOn" in body:
                event.reminder_on = body["isReminderOn"]
            if "reminderMinutesBeforeStart" in body:
                event.reminder_minutes = body["reminderMinutesBeforeStart"]
            return _json(self.row(event))

        return self._apply("patch", effect)

    def _delete(self, event_id: str, **_: Any) -> GraphResponse:
        self._event(event_id)

        def effect() -> GraphResponse:
            self.events.pop(event_id, None)
            return GraphResponse(204, {}, b"")

        return self._apply("delete", effect)

    def _cancel(self, event_id: str, **_: Any) -> GraphResponse:
        event = self._event(event_id)

        def effect() -> GraphResponse:
            assert event.is_organizer
            # The organizer's copy leaves the calendar (moved to Deleted Items).
            self.events.pop(event_id, None)
            return GraphResponse(202, {}, b"")

        return self._apply("cancel", effect)

    def _respond(
        self, event_id: str, action: str, *, body: dict[str, Any], **_: Any
    ) -> GraphResponse:
        event = self._event(event_id)
        assert "sendResponse" in body

        def effect() -> GraphResponse:
            event.response = {
                "accept": "accepted",
                "tentativelyAccept": "tentativelyAccepted",
                "decline": "declined",
            }[action]
            if action == "decline" and self.decline_removes:
                self.events.pop(event_id, None)
            return GraphResponse(202, {}, b"")

        return self._apply("respond", effect)

    decline_removes = False

    def _get_schedule(self, *, body: dict[str, Any], **_: Any) -> GraphResponse:
        def effect() -> GraphResponse:
            return _json({"value": self.schedule})

        return self._apply("schedule", effect)


class ScriptedGraph:
    """Answers requests in order from a script of dicts or exceptions; records calls."""

    def __init__(self, *script: Any) -> None:
        self.script = list(script)
        self.calls: list[Call] = []

    def remaining_seconds(self) -> float:
        return 200.0

    def request(self, method: str, path_or_url: str, **kwargs: Any) -> GraphResponse:
        self.calls.append(
            Call(
                method,
                path_or_url,
                dict(kwargs.get("query") or {}),
                kwargs.get("json_body"),
                dict(kwargs.get("headers") or {}),
                kwargs["retry"],
            )
        )
        if not self.script:
            raise AssertionError(f"unexpected {method} {path_or_url}")
        step = self.script.pop(0)
        if isinstance(step, Exception):
            raise step
        return _json(step)


__all__ = [
    "APP_BASE",
    "APP_CONNECTION",
    "MAILBOX",
    "MODES",
    "MODE_IDS",
    "STEP_EXECUTION_ID",
    "USER_CONNECTION",
    "FakeCalendar",
    "ScriptedGraph",
    "action_payload",
    "ambiguous",
    "graph_utc",
    "utc",
]
