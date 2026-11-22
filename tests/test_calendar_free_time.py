"""find_free_time (task D2): getSchedule in both modes, busy from scheduleItems and the
availability view, working hours in a time zone, and the rule that an unknown or
errored address yields complete: false with no common slots."""

from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any

import pytest
from test_calendar_fakes import MODE_IDS, MODES, FakeCalendar, action_payload, graph_utc, utc

from calendar_ops.operations import OPERATIONS
from m365.runtime import run_operation

#: 2026-11-23 is a Monday; Berlin is UTC+1 in November.
WINDOW = {
    "start": "2026-11-23T00:00:00",
    "end": "2026-11-24T00:00:00",
    "time_zone": "Europe/Berlin",
    "working_hours_start": "09:00",
    "working_hours_end": "12:00",
    "duration_minutes": 30,
}
VIEW_START = utc("2026-11-22T23:00:00Z")
QUARTERS = 96


def _view(*busy: tuple[str, str, str]) -> str:
    """A 15-minute view of the window: '0' except the given (start, end, digit) spans."""
    chars = ["0"] * QUARTERS
    for start, end, digit in busy:
        first = int((utc(start) - VIEW_START) / timedelta(minutes=15))
        last = int((utc(end) - VIEW_START) / timedelta(minutes=15))
        for index in range(first, last):
            chars[index] = digit
    return "".join(chars)


def _item(status: str, start: str, end: str) -> dict[str, Any]:
    return {"status": status, "start": graph_utc(utc(start)), "end": graph_utc(utc(end))}


def _schedules() -> list[dict[str, Any]]:
    return [
        {
            "scheduleId": "ana@acme.example",
            "availabilityView": _view(("2026-11-23T08:00:00Z", "2026-11-23T09:00:00Z", "2")),
            "scheduleItems": [_item("busy", "2026-11-23T08:00:00Z", "2026-11-23T09:00:00Z")],
        },
        {
            "scheduleId": "Bob@Acme.example",
            "availabilityView": _view(("2026-11-23T10:00:00Z", "2026-11-23T10:30:00Z", "1")),
            "scheduleItems": [_item("tentative", "2026-11-23T10:00:00Z", "2026-11-23T10:30:00Z")],
        },
        # Only the view says busy: 11:30-11:45 local.
        {
            "scheduleId": "carl@acme.example",
            "availabilityView": _view(("2026-11-23T10:30:00Z", "2026-11-23T10:45:00Z", "2")),
            "scheduleItems": [],
        },
        # Working elsewhere does not block.
        {
            "scheduleId": "dan@acme.example",
            "availabilityView": _view(("2026-11-23T09:00:00Z", "2026-11-23T09:30:00Z", "4")),
            "scheduleItems": [
                _item("workingElsewhere", "2026-11-23T09:00:00Z", "2026-11-23T09:30:00Z")
            ],
        },
    ]


ADDRESSES = ["ana@acme.example", "bob@acme.example", "carl@acme.example", "dan@acme.example"]


def _run(fake: FakeCalendar, operation_input: dict[str, Any], **kwargs: Any) -> dict[str, Any]:
    payload = action_payload("find_free_time", operation_input, **kwargs)
    return run_operation(payload, OPERATIONS, transport_factory=lambda _connection: fake)


def _locals(result: dict[str, Any]) -> list[str]:
    return [slot["start_local"][11:16] for slot in result["common_slots"]]


@pytest.mark.parametrize(("connection", "base"), MODES, ids=MODE_IDS)
def test_common_slots_avoid_busy_tentative_and_view_busy_time(connection: Any, base: str) -> None:
    fake = FakeCalendar(base)
    fake.schedule = _schedules()
    response = _run(fake, {**WINDOW, "addresses": ADDRESSES}, connection=connection)

    assert response["ok"] is True, response
    result = response["result"]
    assert result["complete"] is True
    assert result["unresolved_addresses"] == []
    assert _locals(result) == ["10:00", "10:30"]
    assert result["common_slots"][0] == {
        "start_utc": "2026-11-23T09:00:00Z",
        "end_utc": "2026-11-23T09:30:00Z",
        "start_local": "2026-11-23T10:00:00",
        "end_local": "2026-11-23T10:30:00",
    }
    assert result["time_zone"] == "Europe/Berlin"
    assert result["window_start_utc"] == "2026-11-22T23:00:00Z"
    ana = result["schedules"][0]
    assert ana["resolved"] is True
    assert ana["busy"][0]["start_local"] == "2026-11-23T09:00:00"
    assert "_busy" not in ana
    [call] = fake.calls
    assert (call.method, call.url) == ("POST", f"{base}/calendar/getSchedule")
    assert call.retry == "idempotent", "getSchedule only reads, so it may be repeated"
    assert call.json_body == {
        "schedules": ADDRESSES,
        "startTime": {"dateTime": "2026-11-22T23:00:00", "timeZone": "UTC"},
        "endTime": {"dateTime": "2026-11-23T23:00:00", "timeZone": "UTC"},
        "availabilityViewInterval": 15,
    }
    assert 'outlook.timezone="UTC"' in call.headers["Prefer"]


def test_a_windows_zone_gives_the_same_slots() -> None:
    fake = FakeCalendar()
    fake.schedule = _schedules()
    result = _run(fake, {**WINDOW, "time_zone": "W. Europe Standard Time", "addresses": ADDRESSES})[
        "result"
    ]
    assert _locals(result) == ["10:00", "10:30"]


@pytest.mark.parametrize(("connection", "base"), MODES, ids=MODE_IDS)
def test_an_errored_address_blocks_every_common_slot(connection: Any, base: str) -> None:
    fake = FakeCalendar(base)
    fake.schedule = [
        *_schedules()[:2],
        {
            "scheduleId": "ghost@acme.example",
            "availabilityView": "",
            "scheduleItems": [],
            "error": {
                "message": "The specified user could not be found",
                "responseCode": "ErrorInvalidUser",
            },
        },
    ]
    response = _run(
        fake,
        {**WINDOW, "addresses": ["ana@acme.example", "bob@acme.example", "ghost@acme.example"]},
        connection=connection,
    )

    assert response["ok"] is True, response
    result = response["result"]
    assert result["complete"] is False
    assert result["common_slots"] == []
    assert result["slot_count"] == 0
    assert result["unresolved_addresses"] == ["ghost@acme.example"]
    ghost = result["schedules"][2]
    assert (ghost["resolved"], ghost["unresolved_reason"]) == (False, "error")
    assert ghost["error_message"] == "The specified user could not be found"
    # The resolved people still show their busy time, for display only.
    assert result["schedules"][0]["busy"][0]["status"] == "busy"
    assert result["schedules"][1]["busy"][0]["status"] == "tentative"


def test_a_question_mark_in_the_view_is_unknown_not_free() -> None:
    fake = FakeCalendar()
    view = list(_view())
    view[40] = "?"
    fake.schedule = [
        _schedules()[0],
        {"scheduleId": "bob@acme.example", "availabilityView": "".join(view), "scheduleItems": []},
    ]
    result = _run(fake, {**WINDOW, "addresses": ["ana@acme.example", "bob@acme.example"]})["result"]
    assert result["complete"] is False
    assert result["common_slots"] == []
    assert result["unresolved_addresses"] == ["bob@acme.example"]
    assert result["schedules"][1]["unresolved_reason"] == "unknown_availability"


@pytest.mark.parametrize(
    ("row", "reason"),
    [
        (None, "no_schedule"),
        (
            {"scheduleId": "bob@acme.example", "availabilityView": "", "scheduleItems": []},
            "no_data",
        ),
        (
            {"scheduleId": "bob@acme.example", "availabilityView": "0" * 10, "scheduleItems": []},
            "no_data",
        ),
        (
            {
                "scheduleId": "bob@acme.example",
                "availabilityView": "0" * QUARTERS,
                "scheduleItems": [{"status": "busy", "start": {"dateTime": "soon"}, "end": {}}],
            },
            "invalid_schedule",
        ),
    ],
)
def test_missing_or_partial_data_is_never_free(row: dict[str, Any] | None, reason: str) -> None:
    fake = FakeCalendar()
    fake.schedule = [_schedules()[0], *([row] if row else [])]
    result = _run(fake, {**WINDOW, "addresses": ["ana@acme.example", "bob@acme.example"]})["result"]
    assert result["complete"] is False
    assert result["common_slots"] == []
    assert result["unresolved_addresses"] == ["bob@acme.example"]
    assert result["schedules"][1]["unresolved_reason"] == reason


def test_at_most_20_addresses() -> None:
    fake = FakeCalendar()
    many = [f"user{index}@acme.example" for index in range(21)]
    response = _run(fake, {**WINDOW, "addresses": many})
    assert response["error_code"] == "invalid_payload"
    assert fake.calls == []
    fake.schedule = []
    twenty = _run(fake, {**WINDOW, "addresses": many[:20]})
    assert twenty["ok"] is True
    assert len(fake.calls[0].json_body["schedules"]) == 20
    duplicates = _run(fake, {**WINDOW, "addresses": ["ana@acme.example", "ANA@acme.example"]})
    assert duplicates["ok"] is True
    assert fake.calls[1].json_body["schedules"] == ["ana@acme.example"]


def test_the_window_must_be_shorter_than_62_days() -> None:
    start = datetime(2026, 11, 1)
    fake = FakeCalendar()
    refused = _run(
        fake,
        {
            **WINDOW,
            "addresses": ["ana@acme.example"],
            "start": start.isoformat(),
            "end": (start + timedelta(days=62)).isoformat(),
        },
    )
    assert refused["error_code"] == "range_too_large"
    assert fake.calls == []
    accepted = _run(
        fake,
        {
            **WINDOW,
            "addresses": ["ana@acme.example"],
            "start": start.isoformat(),
            "end": (start + timedelta(days=62, minutes=-15)).isoformat(),
        },
    )
    assert accepted["ok"] is True, accepted
    assert len(fake.calls) == 1


def test_weekends_are_skipped_unless_included() -> None:
    weekend = {
        **WINDOW,
        "start": "2026-11-21T00:00:00",
        "end": "2026-11-23T00:00:00",
        "addresses": ["sam@acme.example"],
    }
    view_start = utc("2026-11-20T23:00:00Z")
    row = {"scheduleId": "sam@acme.example", "availabilityView": "0" * 192, "scheduleItems": []}
    fake = FakeCalendar()
    fake.schedule = [row]
    skipped = _run(fake, weekend)["result"]
    assert (skipped["complete"], skipped["common_slots"]) == (True, [])
    included = _run(fake, {**weekend, "include_weekends": True})["result"]
    assert included["slot_count"] == 12
    assert fake.calls[0].json_body["startTime"]["dateTime"] == view_start.strftime(
        "%Y-%m-%dT%H:%M:%S"
    )


def test_max_slots_truncates() -> None:
    fake = FakeCalendar()
    fake.schedule = [
        {"scheduleId": "sam@acme.example", "availabilityView": "0" * QUARTERS, "scheduleItems": []}
    ]
    result = _run(fake, {**WINDOW, "addresses": ["sam@acme.example"], "max_slots": 2})["result"]
    assert (result["slot_count"], result["truncated"]) == (2, True)
    full = _run(fake, {**WINDOW, "addresses": ["sam@acme.example"], "slot_interval_minutes": 60})[
        "result"
    ]
    assert (_locals(full), full["truncated"]) == (["09:00", "10:00", "11:00"], False)


def test_a_window_starting_inside_a_quarter_is_aligned_but_slots_stay_inside() -> None:
    fake = FakeCalendar()
    row = {"scheduleId": "sam@acme.example", "availabilityView": "0" * 64, "scheduleItems": []}
    fake.schedule = [row]
    result = _run(
        fake,
        {
            **WINDOW,
            "start": "2026-11-23T09:07:00",
            "end": "2026-11-23T11:00:00",
            "addresses": ["sam@acme.example"],
        },
    )["result"]
    assert fake.calls[0].json_body["startTime"]["dateTime"] == "2026-11-23T08:00:00"
    assert fake.calls[0].json_body["endTime"]["dateTime"] == "2026-11-23T10:00:00"
    assert _locals(result) == ["09:30", "10:00", "10:30"]


def test_an_unanswered_get_schedule_is_a_read_failure_not_timeout_unknown() -> None:
    fake = FakeCalendar()
    fake.fault("schedule", "dropped")
    response = _run(fake, {**WINDOW, "addresses": ["ana@acme.example"]})
    assert response["error_code"] == "provider_unavailable"
    assert "external_effect_status" not in response
    assert response["retryable"] is True


@pytest.mark.parametrize(
    "change",
    [
        {"duration_minutes": 7},
        {"slot_interval_minutes": 20},
        {"working_hours_start": "9:00"},
        {"working_hours_start": "09:10"},
        {"working_hours_start": "12:00"},
        {"addresses": []},
        {"addresses": ["not-an-address"]},
        {"time_zone": "Atlantis/Capital"},
        {"include_weekends": "yes"},
    ],
)
def test_find_free_time_input_errors(change: dict[str, Any]) -> None:
    fake = FakeCalendar()
    response = _run(fake, {**WINDOW, "addresses": ["ana@acme.example"], **change})
    assert response["error_code"] == "invalid_payload", response
    assert fake.calls == []


def test_find_free_time_runs_in_test_mode() -> None:
    fake = FakeCalendar()
    fake.schedule = _schedules()
    response = _run(fake, {**WINDOW, "addresses": ADDRESSES}, test_mode=True)
    assert response["result"]["complete"] is True
