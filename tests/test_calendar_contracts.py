"""The calendar contract fragments agree with the calendar registry (task D2)."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import yaml

from calendar_ops.common import EVENT_KEY_PROPERTY_ID
from calendar_ops.operations import OPERATIONS, QUERY_IDS
from m365.connection import CONNECTION_TYPES
from m365.errors import SAFE_ERROR_CODES
from mail.send import KEY_PROPERTY_ID, PROPERTY_SET_GUID

ROOT = Path(__file__).resolve().parents[1]
PARTS = ROOT / "contracts" / "parts" / "calendar"
CALENDAR_TYPES = ["m365_user", "m365_calendar_app"]
EFFECTFUL = {op for op, row in OPERATIONS.items() if row.effectful}


def _load(name: str) -> list[dict[str, Any]]:
    value = yaml.safe_load((PARTS / name).read_text(encoding="utf-8"))
    assert isinstance(value, list)
    return value


def test_registry_covers_task_d2() -> None:
    assert set(OPERATIONS) == {
        "test_calendar_connection",
        "list_calendars",
        "list_events",
        "get_event",
        "create_event",
        "update_event",
        "cancel_event",
        "respond_to_event",
        "find_free_time",
        "check_calendar_access",
    }
    assert EFFECTFUL == {"create_event", "update_event", "cancel_event", "respond_to_event"}
    assert all(row.service == "calendar" for row in OPERATIONS.values())
    assert {CONNECTION_TYPES[t] for t in CALENDAR_TYPES} == {
        ("*", "user"),
        ("calendar", "app"),
    }


def test_the_package_never_shadows_the_stdlib_calendar_module() -> None:
    import calendar

    assert hasattr(calendar, "monthrange")
    assert not (ROOT / "calendar").exists()


def test_the_event_key_lives_in_the_v01_mail_property_set() -> None:
    assert EVENT_KEY_PROPERTY_ID == KEY_PROPERTY_ID
    assert PROPERTY_SET_GUID in EVENT_KEY_PROPERTY_ID


def test_operations_serve_both_calendar_types_without_oauth_scopes() -> None:
    rows = _load("operations.yaml")
    assert {row["operation_id"] for row in rows} == set(OPERATIONS) - {"check_calendar_access"}
    for row in rows:
        op = row["operation_id"]
        assert row["connection_type_ids"] == CALENDAR_TYPES
        assert "required_oauth_scopes" not in row
        inputs = {item["name"] for item in row["inputs"]}
        assert inputs - {"connection_ref"} == set(OPERATIONS[op].fields), op
        assert set(row["error_codes"]) <= SAFE_ERROR_CODES, op
        assert ("timeout_unknown" in row["error_codes"]) == (op in EFFECTFUL), op
        assert row["operation_kind"] == ("action" if op in EFFECTFUL else "source")
        outputs = [item["name"] for item in row["outputs"]]
        assert [f"result.{name}" for name in outputs] == row["result_fields"], op


def test_no_notify_attendees_input_anywhere() -> None:
    for name in ("operations.yaml", "step_ui.yaml", "actions.yaml"):
        assert "notify_attendees" not in (PARTS / name).read_text(encoding="utf-8")
    assert "notify_attendees" not in OPERATIONS["update_event"].fields


def test_specific_error_codes_are_declared_where_they_occur() -> None:
    rows = {row["operation_id"]: row for row in _load("operations.yaml")}
    assert "online_meeting_unsupported" in rows["create_event"]["error_codes"]
    assert "calendar_not_found" in rows["create_event"]["error_codes"]
    assert "not_organizer" in rows["cancel_event"]["error_codes"]
    assert "range_too_large" in rows["list_events"]["error_codes"]
    assert "range_too_large" in rows["find_free_time"]["error_codes"]
    for op in ("get_event", "update_event", "respond_to_event"):
        assert "event_not_found" in rows[op]["error_codes"], op


def test_one_calendar_query_per_connection_type() -> None:
    rows = _load("resource_queries.yaml")
    assert {row["query_id"] for row in rows} == QUERY_IDS
    assert sorted(row["connection_type_id"] for row in rows) == sorted(CALENDAR_TYPES)
    assert all(row["operation_id"] == "list_calendars" for row in rows)


def test_every_calendar_write_is_an_idempotent_suppressed_external_effect() -> None:
    rows = _load("external_effects.yaml")
    assert {row["operation_id"] for row in rows} == EFFECTFUL
    for row in rows:
        assert row["channel"] == "microsoft_outlook_calendar"
        assert row["idempotency"] == "required"
        assert row["test_mode_behavior"] == "suppress"


def test_actions_match_the_registry() -> None:
    rows = {row["action_id"]: row for row in _load("actions.yaml")}
    assert set(rows) == set(OPERATIONS)
    assert rows["check_calendar_access"]["workflow_visible"] is False
    assert rows["check_calendar_access"]["connection_type_ids"] == ["m365_calendar_app"]
    operations = {row["operation_id"]: row for row in _load("operations.yaml")}
    for action_id, row in rows.items():
        assert set(row["error_codes"]) <= SAFE_ERROR_CODES, action_id
        if action_id in EFFECTFUL:
            assert row["test_mode_behavior"] == "suppress"
            assert row["mutates_platform"] is True
        else:
            assert row["mutates_platform"] is False
        if action_id in operations:
            params = {item["name"] for item in row["parameters"]}
            assert params == {item["name"] for item in operations[action_id]["inputs"]}
            assert row["error_codes"] == operations[action_id]["error_codes"]


def test_step_ui_fields_are_declared_inputs() -> None:
    inputs = {
        row["operation_id"]: {item["name"] for item in row["inputs"]}
        for row in _load("operations.yaml")
    }
    forms = _load("step_ui.yaml")
    assert {form["operation_id"] for form in forms} == set(inputs)
    for form in forms:
        for field in form["fields"]:
            assert field["name"] in inputs[form["operation_id"]], (
                form["operation_id"],
                field["name"],
            )
            assert field["widget_type"] != "resource_picker"


def test_calendar_writes_never_use_transport_retries() -> None:
    """Writes are confirmed by reading; only the getSchedule read may be repeated."""
    for path in sorted((ROOT / "calendar_ops").glob("*.py")):
        source = path.read_text(encoding="utf-8")
        uses = re.findall(r'retry="(idempotent|read)"', source)
        if path.name == "schedule.py":
            assert uses == ["idempotent"]
        else:
            assert "idempotent" not in uses, path.name
    writes = (ROOT / "calendar_ops" / "writes.py").read_text(encoding="utf-8")
    assert 'retry="none"' in writes
