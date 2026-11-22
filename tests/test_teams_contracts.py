"""The Teams contract fragments agree with the Teams registry (v0.2 D3)."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import yaml

from m365.connection import CONNECTION_TYPES
from m365.errors import SAFE_ERROR_CODES
from teams.operations import OPERATIONS, QUERY_IDS

ROOT = Path(__file__).resolve().parents[1]
PARTS = ROOT / "contracts" / "parts" / "teams"
EFFECTFUL = {op for op, row in OPERATIONS.items() if row.effectful}


def _load(name: str) -> list[dict[str, Any]]:
    value = yaml.safe_load((PARTS / name).read_text(encoding="utf-8"))
    assert isinstance(value, list)
    return value


def test_registry_covers_task_d3() -> None:
    assert set(OPERATIONS) == {
        "test_teams_connection",
        "list_teams",
        "list_channels",
        "list_chats",
        "post_channel_message",
        "reply_to_message",
        "post_chat_message",
        "list_messages",
    }
    assert EFFECTFUL == {"post_channel_message", "reply_to_message", "post_chat_message"}
    assert all(row.service == "teams" for row in OPERATIONS.values())
    assert CONNECTION_TYPES["m365_user"] == ("*", "user")
    assert CONNECTION_TYPES["m365_teams_webhook"] == ("teams", "webhook")


def test_operations_serve_the_delegated_type_only() -> None:
    rows = _load("operations.yaml")
    assert {row["operation_id"] for row in rows} == set(OPERATIONS)
    for row in rows:
        op = row["operation_id"]
        assert row["connection_type_ids"] == ["m365_user"], op
        assert "required_oauth_scopes" not in row
        inputs = {item["name"] for item in row["inputs"]}
        assert inputs - {"connection_ref"} == set(OPERATIONS[op].fields), op
        assert set(row["error_codes"]) <= SAFE_ERROR_CODES, op
        assert ("timeout_unknown" in row["error_codes"]) == (op in EFFECTFUL), op
        assert ("message_too_large" in row["error_codes"]) == (op in EFFECTFUL), op
        assert row["operation_kind"] == ("action" if op in EFFECTFUL else "source")
        outputs = {item["name"] for item in row["outputs"]}
        assert {field.removeprefix("result.") for field in row["result_fields"]} == outputs, op


def test_every_post_is_an_idempotent_suppressed_teams_effect() -> None:
    rows = _load("external_effects.yaml")
    assert {row["operation_id"] for row in rows} == EFFECTFUL
    for row in rows:
        assert row["channel"] == "microsoft_teams"
        assert row["effect_kind"] == "message_delivery"
        assert row["idempotency"] == "required"
        assert row["test_mode_behavior"] == "suppress"


def test_one_query_per_picker_on_the_delegated_type() -> None:
    rows = _load("resource_queries.yaml")
    assert {row["query_id"] for row in rows} == QUERY_IDS
    assert {row["connection_type_id"] for row in rows} == {"m365_user"}
    assert all(row["operation_id"] in OPERATIONS for row in rows)


def test_actions_match_the_registry() -> None:
    rows = {row["action_id"]: row for row in _load("actions.yaml")}
    operations = {row["operation_id"]: row for row in _load("operations.yaml")}
    assert set(rows) == set(OPERATIONS)
    for action_id, row in rows.items():
        assert set(row["error_codes"]) == set(operations[action_id]["error_codes"])
        assert {p["name"] for p in row["parameters"]} == {
            i["name"] for i in operations[action_id]["inputs"]
        }
        if action_id in EFFECTFUL:
            assert row["test_mode_behavior"] == "suppress"
            assert row["mutates_platform"] is True
        else:
            assert row["mutates_platform"] is False


def test_step_ui_fields_are_declared_inputs() -> None:
    inputs = {
        row["operation_id"]: {item["name"] for item in row["inputs"]}
        for row in _load("operations.yaml")
    }
    forms = _load("step_ui.yaml")
    assert {form["operation_id"] for form in forms} == set(OPERATIONS)
    for form in forms:
        for field in form["fields"]:
            assert field["name"] in inputs[form["operation_id"]], (form["operation_id"], field)
            assert field["widget_type"] != "resource_picker"


def test_message_posts_never_use_transport_retries() -> None:
    """D3: a message POST is never repeated; only the idempotent 1:1 chat create may be."""
    source = (ROOT / "teams" / "posting.py").read_text(encoding="utf-8")
    assert source.count('retry="idempotent"') == 1
    assert re.search(r'"POST",\s*"/chats",\s*json_body=body,\s*retry="idempotent"', source)
    assert source.count('retry="none",') == 1  # the message POST, in post_message
    for path in sorted((ROOT / "teams").glob("*.py")):
        if path.name != "posting.py":
            assert 'retry="idempotent"' not in path.read_text(encoding="utf-8"), path.name
