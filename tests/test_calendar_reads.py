"""Calendar reads in both auth modes: connection test, calendars, calendar view with its
31-day limit and cursor, one event, the picker query, and the time-zone rules."""

from __future__ import annotations

from datetime import timedelta
from typing import Any

import pytest
from test_calendar_fakes import (
    APP_CONNECTION,
    MODE_IDS,
    MODES,
    USER_CONNECTION,
    FakeCalendar,
    ScriptedGraph,
    action_payload,
    utc,
)

from calendar_ops.operations import OPERATIONS, handle_query
from calendar_ops.zones import WINDOWS_ZONES, lookup_zone
from m365.errors import M365Error
from m365.runtime import run_operation


def _run(
    graph: Any, action_id: str, operation_input: dict[str, Any] | None = None, **kwargs: Any
) -> dict[str, Any]:
    payload = action_payload(action_id, operation_input, **kwargs)
    return run_operation(payload, OPERATIONS, transport_factory=lambda _connection: graph)


def _seed(fake: FakeCalendar) -> None:
    fake.add_event(
        subject="Berlin sync",
        start=utc("2026-11-25T08:00:00Z"),
        end=utc("2026-11-25T09:00:00Z"),
        zone="W. Europe Standard Time",
        body="Agenda: numbers",
    )
    fake.add_event(
        subject="NY call",
        start=utc("2026-11-26T15:00:00Z"),
        end=utc("2026-11-26T16:00:00Z"),
        zone="Eastern Standard Time",
    )
    fake.add_event(
        subject="Later", start=utc("2026-11-27T10:00:00Z"), end=utc("2026-11-27T11:00:00Z")
    )
    fake.add_event(calendar_id="CAL-HOLIDAYS", subject="Holiday")


# -- test_calendar_connection -------------------------------------------------------------


@pytest.mark.parametrize(("connection", "base"), MODES, ids=MODE_IDS)
def test_connection_reports_the_calendar_and_teams_support(connection: Any, base: str) -> None:
    fake = FakeCalendar(base)
    response = _run(fake, "test_calendar_connection", connection=connection)

    assert response["ok"] is True, response
    result = response["result"]
    assert result["mailbox_address"] == "cmr@acme.example"
    assert result["auth_mode"] == ("user" if base == "/me" else "app")
    assert result["display_name"] == ("Ana Acme" if base == "/me" else None)
    assert result["calendar_id"] == "CAL-DEFAULT"
    assert result["teams_meetings_allowed"] is True
    assert result["allowed_online_meeting_providers"] == ["teamsForBusiness"]
    assert fake.calls[-1].url == f"{base}/calendar"
    assert ("/me" in [call.url for call in fake.calls]) == (base == "/me")


def test_an_unknown_app_mailbox_is_mailbox_not_found() -> None:
    graph = ScriptedGraph(M365Error("item_not_found", "missing", http_status=404))
    response = _run(graph, "test_calendar_connection", connection=APP_CONNECTION)
    assert response["error_code"] == "mailbox_not_found"


def test_app_only_rbac_refusal_is_mailbox_access_denied() -> None:
    graph = ScriptedGraph(M365Error("mailbox_access_denied", "denied", http_status=403))
    response = _run(graph, "test_calendar_connection", connection=APP_CONNECTION)
    assert response["error_code"] == "mailbox_access_denied"


def test_an_app_connection_without_mailbox_is_invalid() -> None:
    connection = {**APP_CONNECTION, "config": {**APP_CONNECTION["config"], "mailbox": ""}}
    response = _run(ScriptedGraph(), "test_calendar_connection", connection=connection)
    assert response["error_code"] == "invalid_connection"


def test_a_mail_connection_cannot_run_calendar_operations() -> None:
    connection = {**USER_CONNECTION, "connection_type_id": "m365_mail_user"}
    response = _run(ScriptedGraph(), "list_calendars", connection=connection)
    assert response["error_code"] == "invalid_connection"


# -- list_calendars --------------------------------------------------------------------------


@pytest.mark.parametrize(("connection", "base"), MODES, ids=MODE_IDS)
def test_list_calendars(connection: Any, base: str) -> None:
    fake = FakeCalendar(base)
    result = _run(fake, "list_calendars", connection=connection)["result"]
    assert result["count"] == 2
    by_id = {row["calendar_id"]: row for row in result["calendars"]}
    assert by_id["CAL-DEFAULT"]["is_default"] is True
    assert by_id["CAL-HOLIDAYS"]["teams_meetings_allowed"] is False
    assert fake.calls[0].url == f"{base}/calendars"


@pytest.mark.parametrize(("connection", "base"), MODES, ids=MODE_IDS)
def test_calendar_picker_lists_the_default_first(connection: Any, base: str) -> None:
    fake = FakeCalendar(base)
    query_id = "calendars_user" if base == "/me" else "calendars_app"
    payload = {
        "contract_version": "extension_host_v2",
        "mode": "query",
        "runtime_context": action_payload("x")["runtime_context"],
        "query": {"query_id": query_id, "context": {"connection_id": "conn-acme-calendar"}},
        "target": {"connection": dict(connection)},
    }
    response = handle_query(payload, transport_factory=lambda _connection: fake)
    assert response["ok"] is True, response
    assert response["result"]["items"] == [
        {"calendar_id": "CAL-DEFAULT", "label": "Calendar (default)"},
        {"calendar_id": "CAL-HOLIDAYS", "label": "Holidays"},
    ]


def test_calendar_picker_refuses_the_other_connection_type() -> None:
    payload = {
        "query": {"query_id": "calendars_app", "context": {"connection_id": "conn-acme-calendar"}},
        "target": {"connection": dict(USER_CONNECTION)},
    }
    response = handle_query(payload, transport_factory=lambda _connection: FakeCalendar())
    assert response["error_code"] == "invalid_connection"


# -- list_events -----------------------------------------------------------------------------


@pytest.mark.parametrize(("connection", "base"), MODES, ids=MODE_IDS)
def test_list_events_uses_calendar_view_in_utc_with_the_original_zone(
    connection: Any, base: str
) -> None:
    fake = FakeCalendar(base)
    _seed(fake)
    response = _run(
        fake,
        "list_events",
        {"start": "2026-11-25T00:00:00Z", "end": "2026-11-28T00:00:00Z"},
        connection=connection,
    )

    assert response["ok"] is True, response
    result = response["result"]
    assert [event["subject"] for event in result["events"]] == ["Berlin sync", "NY call", "Later"]
    berlin = result["events"][0]
    assert berlin["start_utc"] == "2026-11-25T08:00:00Z"
    assert berlin["start_time_zone"] == "W. Europe Standard Time"
    assert berlin["start_local"] == "2026-11-25T09:00:00"
    assert result["events"][1]["start_local"] == "2026-11-26T10:00:00"
    assert result["next_cursor"] == ""
    call = fake.calls[0]
    assert call.url == f"{base}/calendar/calendarView"
    assert call.query["startDateTime"] == "2026-11-25T00:00:00Z"
    assert call.query["endDateTime"] == "2026-11-28T00:00:00Z"
    assert 'outlook.timezone="UTC"' in call.headers["Prefer"]
    assert 'IdType="ImmutableId"' in call.headers["Prefer"]


def test_list_events_on_another_calendar() -> None:
    fake = FakeCalendar()
    _seed(fake)
    result = _run(
        fake,
        "list_events",
        {
            "calendar_id": "CAL-HOLIDAYS",
            "start": "2026-11-25T00:00:00Z",
            "end": "2026-11-26T00:00:00Z",
        },
    )["result"]
    assert [event["subject"] for event in result["events"]] == ["Holiday"]
    assert fake.calls[0].url == "/me/calendars/CAL-HOLIDAYS/calendarView"


def test_list_events_unknown_calendar_is_calendar_not_found() -> None:
    response = _run(
        FakeCalendar(),
        "list_events",
        {"calendar_id": "CAL-NOPE", "start": "2026-11-25T00:00:00Z", "end": "2026-11-26T00:00:00Z"},
    )
    assert response["error_code"] == "calendar_not_found"


def test_list_events_allows_exactly_31_days_and_refuses_more_before_any_call() -> None:
    fake = FakeCalendar()
    start = utc("2026-11-01T00:00:00Z")
    ok = _run(
        fake,
        "list_events",
        {"start": "2026-11-01T00:00:00Z", "end": (start + timedelta(days=31)).isoformat()},
    )
    assert ok["ok"] is True, ok
    graph = ScriptedGraph()
    refused = _run(
        graph,
        "list_events",
        {"start": "2026-11-01T00:00:00Z", "end": "2026-12-02T00:00:01Z"},
    )
    assert refused["error_code"] == "range_too_large"
    assert graph.calls == []


def test_list_events_rejects_inverted_or_zone_less_windows() -> None:
    inverted = _run(
        ScriptedGraph(),
        "list_events",
        {"start": "2026-11-02T00:00:00Z", "end": "2026-11-01T00:00:00Z"},
    )
    assert inverted["error_code"] == "invalid_payload"
    naive = _run(
        ScriptedGraph(),
        "list_events",
        {"start": "2026-11-01T00:00:00", "end": "2026-11-02T00:00:00"},
    )
    assert naive["error_code"] == "invalid_payload"


def test_list_events_reads_local_times_in_the_given_zone() -> None:
    fake = FakeCalendar()
    response = _run(
        fake,
        "list_events",
        {
            "start": "2026-11-25T00:00:00",
            "end": "2026-11-26T00:00:00",
            "time_zone": "Europe/Berlin",
        },
    )
    assert response["result"]["window_start_utc"] == "2026-11-24T23:00:00Z"
    assert fake.calls[0].query["endDateTime"] == "2026-11-25T23:00:00Z"


@pytest.mark.parametrize(("connection", "base"), MODES, ids=MODE_IDS)
def test_list_events_cursor_continues_the_same_query(connection: Any, base: str) -> None:
    fake = FakeCalendar(base)
    _seed(fake)
    window = {"start": "2026-11-25T00:00:00Z", "end": "2026-11-28T00:00:00Z", "limit": 2}
    first = _run(fake, "list_events", window, connection=connection)["result"]
    assert [event["subject"] for event in first["events"]] == ["Berlin sync", "NY call"]
    assert first["next_cursor"].startswith("m365m1.")
    assert "graph.microsoft.com" not in first["next_cursor"]

    second = _run(
        fake, "list_events", {**window, "cursor": first["next_cursor"]}, connection=connection
    )["result"]
    assert [event["subject"] for event in second["events"]] == ["Later"]
    assert second["next_cursor"] == ""

    other = _run(
        fake,
        "list_events",
        {**window, "end": "2026-11-29T00:00:00Z", "cursor": first["next_cursor"]},
        connection=connection,
    )
    assert other["error_code"] == "invalid_cursor"


# -- get_event --------------------------------------------------------------------------------


@pytest.mark.parametrize(("connection", "base"), MODES, ids=MODE_IDS)
def test_get_event_with_text_body(connection: Any, base: str) -> None:
    fake = FakeCalendar(base)
    _seed(fake)
    event_id = next(iter(fake.events))
    response = _run(fake, "get_event", {"event_id": event_id}, connection=connection)

    assert response["ok"] is True, response
    event = response["result"]["event"]
    assert event["event_id"] == event_id
    assert event["body_text"] == "Agenda: numbers"
    assert event["body_text_truncated"] is False
    assert event["end_utc"] == "2026-11-25T09:00:00Z"
    call = fake.calls[0]
    assert call.url == f"{base}/events/{event_id}"
    assert "body" in call.query["$select"].split(",")
    assert 'outlook.body-content-type="text"' in call.headers["Prefer"]


def test_get_event_without_body_and_unknown_event() -> None:
    fake = FakeCalendar()
    _seed(fake)
    event_id = next(iter(fake.events))
    event = _run(fake, "get_event", {"event_id": event_id, "include_body": False})["result"][
        "event"
    ]
    assert "body_text" not in event
    missing = _run(fake, "get_event", {"event_id": "AAMkNOPE"})
    assert missing["error_code"] == "event_not_found"
    bad = _run(fake, "get_event", {"event_id": "../x"})
    assert bad["error_code"] == "invalid_payload"


# -- time zones -------------------------------------------------------------------------------


def test_every_windows_zone_resolves() -> None:
    unresolved = [name for name, iana in WINDOWS_ZONES.items() if lookup_zone(name) is None]
    assert unresolved == []
    assert lookup_zone("w. europe standard time") is not None
    assert lookup_zone("Europe/Berlin") is not None
    assert lookup_zone("tzone://Microsoft/Utc") is not None
    for bad in ("", "Mars/Olympus", "../etc/passwd", "Customized Time Zone"):
        assert lookup_zone(bad) is None
