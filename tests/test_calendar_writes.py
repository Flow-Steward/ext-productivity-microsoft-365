"""Calendar writes in both auth modes, and every ambiguous path (task D2, D7):

- create_event: the step key in a single-value extended property and as transactionId,
  the documented property filter alone, a client-side start check, never a second POST;
- the Teams online-meeting check before anything is created;
- update / cancel / respond confirmed by re-reading.
"""

from __future__ import annotations

from typing import Any

import pytest
from test_calendar_fakes import (
    APP_CONNECTION,
    MAILBOX,
    MODE_IDS,
    MODES,
    STEP_EXECUTION_ID,
    FakeCalendar,
    action_payload,
    utc,
)

from calendar_ops import writes
from calendar_ops.common import EVENT_KEY_PROPERTY_ID
from calendar_ops.operations import OPERATIONS
from m365.runtime import run_operation

KEY = f"create_event:{STEP_EXECUTION_ID}"
DOCUMENTED_FILTER = (
    f"singleValueExtendedProperties/Any(ep: ep/id eq '{EVENT_KEY_PROPERTY_ID}' "
    f"and ep/value eq '{KEY}')"
)


@pytest.fixture(autouse=True)
def _no_waiting(monkeypatch: pytest.MonkeyPatch) -> list[float]:
    waits: list[float] = []
    monkeypatch.setattr(writes, "sleep", waits.append)
    return waits


def _run(
    fake: FakeCalendar, action_id: str, operation_input: dict[str, Any], **kwargs: Any
) -> dict[str, Any]:
    payload = action_payload(action_id, operation_input, **kwargs)
    return run_operation(payload, OPERATIONS, transport_factory=lambda _connection: fake)


def _create_input(**extra: Any) -> dict[str, Any]:
    return {
        "subject": "Booking 42",
        "start": "2026-11-25T10:00:00",
        "end": "2026-11-25T10:30:00",
        "time_zone": "Europe/Berlin",
        "attendees": ["buyer@acme.example"],
        "optional_attendees": [{"address": "ops@acme.example", "name": "Ops"}],
        "location": "Room 1",
        "body_text": "See you there.",
        "reminder_minutes": 10,
        **extra,
    }


def _posts(fake: FakeCalendar) -> list[Any]:
    return [call for call in fake.calls if call.method == "POST" and call.url.endswith("/events")]


def _lookups(fake: FakeCalendar) -> list[Any]:
    return [call for call in fake.calls if call.method == "GET" and "$filter" in call.query]


def _assert_no_transport_retry(fake: FakeCalendar) -> None:
    for call in fake.writes():
        assert call.retry == "none", f"{call.method} {call.url} may be retried by the transport"


def _ambiguous(response: dict[str, Any]) -> None:
    assert response["ok"] is False
    assert response["error_code"] == "timeout_unknown"
    assert response["external_effect_status"] == "timeout_unknown"
    assert response["definitely_no_external_effect"] is False
    assert response["retryable"] is False


# -- create_event ----------------------------------------------------------------------------


@pytest.mark.parametrize(("connection", "base"), MODES, ids=MODE_IDS)
def test_create_event_carries_the_key_twice_and_keeps_the_zone(connection: Any, base: str) -> None:
    fake = FakeCalendar(base)
    response = _run(fake, "create_event", _create_input(), connection=connection)

    assert response["ok"] is True, response
    assert response["external_effect_status"] == "succeeded"
    assert response["definitely_no_external_effect"] is False
    result = response["result"]
    assert result["confirmation"] == "created"
    assert result["changed"] is True
    [event] = fake.events.values()
    assert result["event_id"] == event.id
    assert event.props == {EVENT_KEY_PROPERTY_ID: KEY}
    assert event.transaction_id == KEY
    [post] = _posts(fake)
    assert post.url == f"{base}/events"
    body = post.json_body
    assert body["start"] == {"dateTime": "2026-11-25T10:00:00", "timeZone": "Europe/Berlin"}
    assert body["end"] == {"dateTime": "2026-11-25T10:30:00", "timeZone": "Europe/Berlin"}
    assert body["attendees"] == [
        {"emailAddress": {"address": "buyer@acme.example"}, "type": "required"},
        {"emailAddress": {"address": "ops@acme.example", "name": "Ops"}, "type": "optional"},
    ]
    assert body["isReminderOn"] is True and body["reminderMinutesBeforeStart"] == 10
    assert "isOnlineMeeting" not in body
    assert result["event"]["start_utc"] == "2026-11-25T09:00:00Z"
    assert result["event"]["start_time_zone"] == "Europe/Berlin"
    assert result["event"]["start_local"] == "2026-11-25T10:00:00"
    # The re-run guard runs once before the POST, with the documented filter alone.
    [lookup] = _lookups(fake)
    assert lookup.url == f"{base}/events"
    assert lookup.query["$filter"] == DOCUMENTED_FILTER
    assert EVENT_KEY_PROPERTY_ID in lookup.query["$expand"]
    _assert_no_transport_retry(fake)
    for call in fake.calls:
        assert 'IdType="ImmutableId"' in call.headers.get("Prefer", "")


def test_create_event_accepts_utc_slot_times_and_windows_zones() -> None:
    fake = FakeCalendar()
    response = _run(
        fake,
        "create_event",
        {
            "subject": "Slot",
            "start": "2026-11-25T09:00:00Z",
            "end": "2026-11-25T09:30:00Z",
            "time_zone": "W. Europe Standard Time",
        },
    )
    assert response["ok"] is True, response
    body = _posts(fake)[0].json_body
    assert body["start"] == {
        "dateTime": "2026-11-25T10:00:00",
        "timeZone": "W. Europe Standard Time",
    }
    utc_only = FakeCalendar()
    _run(
        utc_only,
        "create_event",
        {"subject": "UTC", "start": "2026-11-25T09:00:00Z", "end": "2026-11-25T09:30:00Z"},
    )
    assert _posts(utc_only)[0].json_body["start"] == {
        "dateTime": "2026-11-25T09:00:00",
        "timeZone": "UTC",
    }


def test_create_event_is_suppressed_in_test_mode() -> None:
    fake = FakeCalendar()
    response = _run(fake, "create_event", _create_input(), test_mode=True)
    assert response["external_effect_status"] == "suppressed"
    assert response["definitely_no_external_effect"] is True
    assert fake.calls == []


@pytest.mark.parametrize(
    ("change", "message"),
    [
        ({"time_zone": None}, "offset"),
        ({"end": "2026-11-25T09:00:00"}, "later than start"),
        ({"time_zone": "Mars/Olympus"}, "known"),
        ({"subject": ""}, "subject"),
        ({"end": None}, "both start and end"),
        ({"body_html": "<p>x</p>"}, "not both"),
        ({"reminder_enabled": False}, "reminder"),
    ],
)
def test_create_event_input_errors_reach_no_graph(change: dict[str, Any], message: str) -> None:
    fake = FakeCalendar()
    response = _run(fake, "create_event", _create_input(**change))
    assert response["error_code"] == "invalid_payload", response
    assert message in response["error"]
    assert response["definitely_no_external_effect"] is True
    assert fake.calls == []


@pytest.mark.parametrize(("connection", "base"), MODES, ids=MODE_IDS)
def test_lost_create_is_confirmed_by_the_extended_property(connection: Any, base: str) -> None:
    fake = FakeCalendar(base)
    fake.fault("create", "lost")
    response = _run(fake, "create_event", _create_input(), connection=connection)

    assert response["ok"] is True, response
    assert response["result"]["confirmation"] == "extended_property"
    assert response["external_effect_status"] == "succeeded"
    assert len(fake.events) == 1
    assert response["result"]["event_id"] == next(iter(fake.events))
    assert len(_posts(fake)) == 1, "an ambiguous create is never re-POSTed"
    assert [call.query["$filter"] for call in _lookups(fake)] == [DOCUMENTED_FILTER] * 2
    _assert_no_transport_retry(fake)


def test_dropped_create_is_timeout_unknown_without_a_second_post(_no_waiting: list[float]) -> None:
    fake = FakeCalendar()
    fake.fault("create", "dropped")
    response = _run(fake, "create_event", _create_input())

    _ambiguous(response)
    assert "check the calendar" in response["error"]
    assert fake.events == {}
    assert len(_posts(fake)) == 1
    assert len(_lookups(fake)) == 1 + writes.LOOKUP_ATTEMPTS
    assert _no_waiting == [writes.LOOKUP_DELAY_SECONDS] * (writes.LOOKUP_ATTEMPTS - 1)


def test_a_failing_lookup_after_an_ambiguous_create_stays_unknown() -> None:
    fake = FakeCalendar()
    fake.fault("create", "lost")
    # The guard before the POST succeeds; every lookup after the lost create fails.
    fake.fault("lookup", "ok", "unavailable", "unavailable", "unavailable")
    response = _run(fake, "create_event", _create_input())
    _ambiguous(response)
    assert len(_posts(fake)) == 1


def test_two_events_with_the_key_are_timeout_unknown_with_all_ids() -> None:
    fake = FakeCalendar()
    fake.fault("create", "lost_twice")
    response = _run(fake, "create_event", _create_input())

    _ambiguous(response)
    assert len(fake.events) == 2
    for event_id in fake.events:
        assert event_id in response["error"]
    assert len(_posts(fake)) == 1


def test_a_lookup_that_ignores_the_filter_cannot_confirm_a_foreign_event() -> None:
    fake = FakeCalendar()
    fake.add_event(
        subject="Someone else's", start=utc("2026-11-25T09:00:00Z"), end=utc("2026-11-25T09:30:00Z")
    )
    fake.lookup_ignores_filter = True
    fake.fault("create", "dropped")
    response = _run(fake, "create_event", _create_input())
    _ambiguous(response)


def test_a_keyed_event_with_another_start_is_not_confirmed() -> None:
    """The date check runs client-side on the property matches."""
    fake = FakeCalendar()
    fake.add_event(
        subject="Booking 42",
        start=utc("2026-11-26T09:00:00Z"),
        end=utc("2026-11-26T09:30:00Z"),
        props={EVENT_KEY_PROPERTY_ID: KEY},
    )
    response = _run(fake, "create_event", _create_input())
    _ambiguous(response)
    assert "different start" in response["error"]
    assert _posts(fake) == []


def test_a_rerun_step_finds_its_event_and_does_not_create_another() -> None:
    fake = FakeCalendar()
    first = _run(fake, "create_event", _create_input())
    again = _run(fake, "create_event", _create_input())

    assert again["ok"] is True, again
    assert again["result"]["confirmation"] == "already_exists"
    assert again["result"]["event_id"] == first["result"]["event_id"]
    assert len(fake.events) == 1
    assert len(_posts(fake)) == 1


def test_another_step_execution_creates_its_own_event() -> None:
    fake = FakeCalendar()
    _run(fake, "create_event", _create_input())
    _run(
        fake,
        "create_event",
        _create_input(),
        step_execution_id="6a1b2c3d-4e5f-4a6b-8c7d-000000000299",
    )
    assert len(fake.events) == 2


def test_a_refused_or_throttled_create_has_no_effect() -> None:
    for mode, code in (("refused", "invalid_payload"), ("throttled", "rate_limited")):
        fake = FakeCalendar()
        fake.fault("create", mode)
        response = _run(fake, "create_event", _create_input())
        assert response["error_code"] == code
        assert response["external_effect_status"] == "failed"
        assert response["definitely_no_external_effect"] is True
        assert len(_lookups(fake)) == 1, "a confirmed refusal needs no lookup"


def test_create_on_an_unknown_calendar_is_calendar_not_found() -> None:
    fake = FakeCalendar()
    response = _run(fake, "create_event", _create_input(calendar_id="CAL-NOPE"))
    assert response["error_code"] == "calendar_not_found"
    assert response["definitely_no_external_effect"] is True


# -- online meetings ----------------------------------------------------------------------------


@pytest.mark.parametrize(("connection", "base"), MODES, ids=MODE_IDS)
def test_online_meeting_on_a_teams_calendar(connection: Any, base: str) -> None:
    fake = FakeCalendar(base)
    response = _run(fake, "create_event", _create_input(online_meeting=True), connection=connection)

    assert response["ok"] is True, response
    body = _posts(fake)[0].json_body
    assert body["isOnlineMeeting"] is True
    assert body["onlineMeetingProvider"] == "teamsForBusiness"
    assert response["result"]["join_url"].startswith("https://teams.example/")
    assert fake.calls[0].url == f"{base}/calendar"
    assert "allowedOnlineMeetingProviders" in fake.calls[0].query["$select"]


def test_online_meeting_on_a_calendar_without_teams_is_refused_before_anything() -> None:
    fake = FakeCalendar()
    response = _run(
        fake, "create_event", _create_input(online_meeting=True, calendar_id="CAL-HOLIDAYS")
    )
    assert response["error_code"] == "online_meeting_unsupported"
    assert response["definitely_no_external_effect"] is True
    assert [(call.method, call.url) for call in fake.calls] == [
        ("GET", "/me/calendars/CAL-HOLIDAYS")
    ]
    assert fake.events == {}


def test_online_meeting_on_an_unknown_calendar_is_calendar_not_found() -> None:
    fake = FakeCalendar()
    response = _run(
        fake, "create_event", _create_input(online_meeting=True, calendar_id="CAL-NOPE")
    )
    assert response["error_code"] == "calendar_not_found"


# -- update_event ---------------------------------------------------------------------------------


def _meeting(fake: FakeCalendar, **fields: Any) -> str:
    fields.setdefault(
        "attendees",
        [
            {
                "emailAddress": {"address": "buyer@acme.example"},
                "type": "required",
                "status": {"response": "none"},
            }
        ],
    )
    return fake.add_event(**fields).id


@pytest.mark.parametrize(("connection", "base"), MODES, ids=MODE_IDS)
def test_update_event_patches_only_the_given_fields(connection: Any, base: str) -> None:
    fake = FakeCalendar(base)
    event_id = _meeting(fake)
    response = _run(
        fake,
        "update_event",
        {
            "event_id": event_id,
            "subject": "Moved",
            "start": "2026-11-25T11:00:00Z",
            "end": "2026-11-25T12:00:00Z",
        },
        connection=connection,
    )

    assert response["ok"] is True, response
    result = response["result"]
    assert result["confirmation"] == "patched"
    assert result["updated_fields"] == ["end_utc", "start_utc", "subject"]
    [patch] = fake.writes()
    assert (patch.method, patch.url) == ("PATCH", f"{base}/events/{event_id}")
    assert set(patch.json_body) == {"subject", "start", "end"}
    assert fake.events[event_id].subject == "Moved"
    _assert_no_transport_retry(fake)


def test_update_event_has_no_notify_attendees_input() -> None:
    fake = FakeCalendar()
    response = _run(
        fake,
        "update_event",
        {"event_id": _meeting(fake), "subject": "x", "notify_attendees": False},
    )
    assert response["error_code"] == "invalid_payload"
    assert "notify_attendees" in response["error"]


def test_lost_update_is_confirmed_by_re_reading_every_field() -> None:
    fake = FakeCalendar()
    event_id = _meeting(fake)
    fake.fault("patch", "lost")
    response = _run(
        fake,
        "update_event",
        {
            "event_id": event_id,
            "subject": "Moved",
            "body_text": "New  agenda\n",
            "location": "Room 2",
            "attendees": ["buyer@acme.example", "cfo@acme.example"],
            "reminder_minutes": 5,
            "start": "2026-11-25T12:00:00",
            "end": "2026-11-25T13:00:00",
            "time_zone": "Europe/Berlin",
        },
    )
    assert response["ok"] is True, response
    assert response["result"]["confirmation"] == "reread"
    assert response["external_effect_status"] == "succeeded"
    assert len(fake.writes()) == 1


def test_dropped_update_is_timeout_unknown_with_the_differing_fields() -> None:
    fake = FakeCalendar()
    event_id = _meeting(fake, location="Room 1")
    fake.fault("patch", "dropped")
    response = _run(
        fake, "update_event", {"event_id": event_id, "subject": "Moved", "location": "Room 1"}
    )
    _ambiguous(response)
    assert "subject" in response["error"]
    assert "location" not in response["error"]
    assert len(fake.writes()) == 1


def test_an_html_body_cannot_be_confirmed_after_an_ambiguous_update() -> None:
    fake = FakeCalendar()
    event_id = _meeting(fake)
    fake.fault("patch", "lost")
    response = _run(fake, "update_event", {"event_id": event_id, "body_html": "<p>New</p>"})
    _ambiguous(response)
    assert "body_html" in response["error"]


def test_an_event_gone_after_an_ambiguous_update_is_unknown() -> None:
    fake = FakeCalendar()
    event_id = _meeting(fake)
    fake.fault("patch", "dropped")
    original = fake._apply

    def apply(op: str, effect: Any) -> Any:
        if op == "patch":
            fake.events.pop(event_id)
        return original(op, effect)

    fake._apply = apply  # type: ignore[method-assign]
    response = _run(fake, "update_event", {"event_id": event_id, "subject": "Moved"})
    _ambiguous(response)
    assert "gone" in response["error"]


@pytest.mark.parametrize(
    "update",
    [{}, {"start": "2026-11-25T11:00:00Z"}, {"time_zone": "Europe/Berlin", "subject": "x"}],
)
def test_update_event_input_errors(update: dict[str, Any]) -> None:
    fake = FakeCalendar()
    event_id = _meeting(fake)
    response = _run(fake, "update_event", {"event_id": event_id, **update})
    assert response["error_code"] == "invalid_payload"
    assert fake.calls == []


def test_update_of_an_unknown_event_is_event_not_found() -> None:
    response = _run(FakeCalendar(), "update_event", {"event_id": "AAMkNOPE", "subject": "x"})
    assert response["error_code"] == "event_not_found"
    assert response["definitely_no_external_effect"] is True


# -- cancel_event -------------------------------------------------------------------------------------


@pytest.mark.parametrize(("connection", "base"), MODES, ids=MODE_IDS)
def test_cancel_event_by_the_organizer_with_a_comment(connection: Any, base: str) -> None:
    fake = FakeCalendar(base)
    event_id = _meeting(fake)
    response = _run(
        fake, "cancel_event", {"event_id": event_id, "comment": "Sorry"}, connection=connection
    )

    assert response["ok"] is True, response
    result = response["result"]
    assert result["confirmation"] == "cancel_accepted"
    assert result["attendees_notified"] == "requested"
    assert result["changed"] is True
    [cancel] = fake.writes()
    assert (cancel.method, cancel.url) == ("POST", f"{base}/events/{event_id}/cancel")
    assert cancel.json_body == {"comment": "Sorry"}
    assert event_id not in fake.events
    _assert_no_transport_retry(fake)


def test_a_non_organizer_gets_not_organizer_and_nothing_is_written() -> None:
    fake = FakeCalendar()
    event_id = _meeting(
        fake, is_organizer=False, organizer="boss@acme.example", response="accepted"
    )
    response = _run(fake, "cancel_event", {"event_id": event_id})
    assert response["error_code"] == "not_organizer"
    assert "decline" in response["error"]
    assert response["definitely_no_external_effect"] is True
    assert fake.writes() == []


def test_cancel_of_an_already_cancelled_or_missing_event_changes_nothing() -> None:
    fake = FakeCalendar()
    cancelled = _meeting(fake, is_cancelled=True)
    first = _run(fake, "cancel_event", {"event_id": cancelled})["result"]
    gone = _run(fake, "cancel_event", {"event_id": "AAMkGONE"})["result"]
    assert (first["changed"], first["confirmation"]) == (False, "already_cancelled")
    assert (gone["changed"], gone["confirmation"]) == (False, "already_gone")
    assert fake.writes() == []


def test_an_appointment_without_attendees_is_deleted() -> None:
    fake = FakeCalendar()
    event_id = fake.add_event(attendees=[]).id
    result = _run(fake, "cancel_event", {"event_id": event_id})["result"]
    assert result["confirmation"] == "deleted"
    assert result["attendees_notified"] == "none"
    assert [(call.method, call.url) for call in fake.writes()] == [
        ("DELETE", f"/me/events/{event_id}")
    ]


def test_lost_cancel_is_success_once_the_event_is_gone() -> None:
    fake = FakeCalendar()
    event_id = _meeting(fake)
    fake.fault("cancel", "lost")
    response = _run(fake, "cancel_event", {"event_id": event_id})
    assert response["ok"] is True, response
    assert response["result"]["confirmation"] == "reread"
    assert response["result"]["attendees_notified"] == "unknown"
    assert len(fake.writes()) == 1


def test_lost_cancel_is_success_when_the_event_reads_as_cancelled() -> None:
    fake = FakeCalendar()
    event = fake.add_event(
        attendees=[{"emailAddress": {"address": "buyer@acme.example"}, "type": "required"}]
    )
    fake.fault("cancel", "dropped")
    original = fake._apply

    def apply(op: str, effect: Any) -> Any:
        if op == "cancel":
            event.is_cancelled = True
        return original(op, effect)

    fake._apply = apply  # type: ignore[method-assign]
    response = _run(fake, "cancel_event", {"event_id": event.id})
    assert response["result"]["attendees_notified"] == "unknown"


def test_dropped_cancel_with_the_event_still_active_is_timeout_unknown() -> None:
    fake = FakeCalendar()
    event_id = _meeting(fake)
    fake.fault("cancel", "dropped")
    response = _run(fake, "cancel_event", {"event_id": event_id})
    _ambiguous(response)
    assert "still active" in response["error"]
    assert len(fake.writes()) == 1


# -- respond_to_event ---------------------------------------------------------------------------------


def _invitation(fake: FakeCalendar, response: str = "notResponded") -> str:
    return _meeting(fake, is_organizer=False, organizer="boss@acme.example", response=response)


@pytest.mark.parametrize(("connection", "base"), MODES, ids=MODE_IDS)
@pytest.mark.parametrize(
    ("action", "expected"),
    [("accept", "accepted"), ("tentativelyAccept", "tentativelyAccepted"), ("decline", "declined")],
)
def test_respond_to_event(connection: Any, base: str, action: str, expected: str) -> None:
    fake = FakeCalendar(base)
    event_id = _invitation(fake)
    response = _run(
        fake,
        "respond_to_event",
        {"event_id": event_id, "response": action, "comment": "Thanks"},
        connection=connection,
    )

    assert response["ok"] is True, response
    result = response["result"]
    assert (result["response"], result["response_sent"]) == (expected, "requested")
    assert result["confirmation"] == "response_accepted"
    [post] = fake.writes()
    assert post.url == f"{base}/events/{event_id}/{action}"
    assert post.json_body == {"sendResponse": True, "comment": "Thanks"}
    _assert_no_transport_retry(fake)


def test_respond_without_sending_the_response() -> None:
    fake = FakeCalendar()
    event_id = _invitation(fake)
    result = _run(
        fake,
        "respond_to_event",
        {"event_id": event_id, "response": "decline", "send_response": False},
    )["result"]
    assert result["response_sent"] == "not_requested"
    assert fake.writes()[0].json_body == {"sendResponse": False}


def test_a_response_already_given_is_not_sent_again() -> None:
    fake = FakeCalendar()
    event_id = _invitation(fake, response="accepted")
    result = _run(fake, "respond_to_event", {"event_id": event_id, "response": "accept"})["result"]
    assert (result["changed"], result["response_sent"]) == (False, "not_sent")
    assert fake.writes() == []


def test_lost_response_is_confirmed_on_the_attendees_copy() -> None:
    fake = FakeCalendar()
    event_id = _invitation(fake)
    fake.fault("respond", "lost")
    response = _run(
        fake, "respond_to_event", {"event_id": event_id, "response": "tentativelyAccept"}
    )
    assert response["ok"] is True, response
    assert response["result"]["confirmation"] == "reread"
    assert response["result"]["response_sent"] == "unknown"
    assert len(fake.writes()) == 1


def test_lost_decline_counts_when_the_event_left_the_calendar() -> None:
    fake = FakeCalendar()
    fake.decline_removes = True
    event_id = _invitation(fake)
    fake.fault("respond", "lost")
    response = _run(fake, "respond_to_event", {"event_id": event_id, "response": "decline"})
    assert response["ok"] is True, response
    assert response["result"]["response_sent"] == "unknown"


def test_lost_accept_with_the_event_gone_is_unknown() -> None:
    fake = FakeCalendar()
    event_id = _invitation(fake)
    fake.fault("respond", "dropped")
    original = fake._apply

    def apply(op: str, effect: Any) -> Any:
        if op == "respond":
            fake.events.pop(event_id, None)
        return original(op, effect)

    fake._apply = apply  # type: ignore[method-assign]
    response = _run(fake, "respond_to_event", {"event_id": event_id, "response": "accept"})
    _ambiguous(response)


def test_dropped_response_is_timeout_unknown() -> None:
    fake = FakeCalendar()
    event_id = _invitation(fake)
    fake.fault("respond", "dropped")
    response = _run(fake, "respond_to_event", {"event_id": event_id, "response": "accept"})
    _ambiguous(response)
    assert len(fake.writes()) == 1


def test_respond_input_errors() -> None:
    fake = FakeCalendar()
    event_id = _invitation(fake)
    bad = _run(fake, "respond_to_event", {"event_id": event_id, "response": "maybe"})
    assert bad["error_code"] == "invalid_payload"
    own = _meeting(fake)
    organizer = _run(fake, "respond_to_event", {"event_id": own, "response": "accept"})
    assert organizer["error_code"] == "invalid_payload"
    missing = _run(fake, "respond_to_event", {"event_id": "AAMkNOPE", "response": "accept"})
    assert missing["error_code"] == "event_not_found"
    assert fake.writes() == []


def test_writes_in_app_mode_stay_in_the_connection_mailbox() -> None:
    fake = FakeCalendar(f"/users/{MAILBOX}")
    _run(fake, "create_event", _create_input(), connection=APP_CONNECTION)
    assert all(call.url.startswith(f"/users/{MAILBOX}/") for call in fake.calls)
