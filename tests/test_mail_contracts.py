"""The mail contract fragments agree with the mail registry (B4.4, B5.1, B6)."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import yaml

from m365.errors import SAFE_ERROR_CODES
from mail import download, outgoing
from mail.operations import OPERATIONS, QUERY_IDS

ROOT = Path(__file__).resolve().parents[1]
PARTS = ROOT / "contracts" / "parts" / "mail"
MAIL_TYPES = ["m365_user", "m365_mail_app"]
EFFECTFUL = {op for op, row in OPERATIONS.items() if row.effectful}


def _load(name: str) -> list[dict[str, Any]]:
    value = yaml.safe_load((PARTS / name).read_text(encoding="utf-8"))
    assert isinstance(value, list)
    return value


def test_registry_covers_task_b51() -> None:
    assert set(OPERATIONS) == {
        "test_connection",
        "list_mail_folders",
        "search_messages",
        "get_message",
        "get_attachment",
        "send_message",
        "reply_message",
        "forward_message",
        "set_message_flags",
        "move_messages",
        "delete_messages",
        "check_mail_access",
    }
    assert EFFECTFUL == {
        "send_message",
        "reply_message",
        "forward_message",
        "set_message_flags",
        "move_messages",
        "delete_messages",
    }
    assert all(row.service == "mail" for row in OPERATIONS.values())


def test_operations_serve_both_mail_types_without_oauth_scopes() -> None:
    rows = _load("operations.yaml")
    assert {row["operation_id"] for row in rows} == set(OPERATIONS) - {"check_mail_access"}
    for row in rows:
        assert row["connection_type_ids"] == MAIL_TYPES
        assert "required_oauth_scopes" not in row
        inputs = {item["name"] for item in row["inputs"]}
        assert inputs - {"connection_ref"} == set(OPERATIONS[row["operation_id"]].fields), row[
            "operation_id"
        ]
        assert set(row["error_codes"]) <= SAFE_ERROR_CODES, row["operation_id"]
        assert ("timeout_unknown" in row["error_codes"]) == (row["operation_id"] in EFFECTFUL)
        assert row["operation_kind"] == ("action" if row["operation_id"] in EFFECTFUL else "source")


def test_one_folder_query_per_connection_type() -> None:
    rows = _load("resource_queries.yaml")
    assert {row["query_id"] for row in rows} == QUERY_IDS
    assert sorted(row["connection_type_id"] for row in rows) == sorted(MAIL_TYPES)
    assert all(row["operation_id"] == "list_mail_folders" for row in rows)


def test_every_mail_write_is_an_idempotent_suppressed_external_effect() -> None:
    rows = _load("external_effects.yaml")
    assert {row["operation_id"] for row in rows} == EFFECTFUL
    assert all(
        row["idempotency"] == "required" and row["test_mode_behavior"] == "suppress" for row in rows
    )


def test_actions_match_the_registry() -> None:
    rows = {row["action_id"]: row for row in _load("actions.yaml")}
    assert set(rows) == set(OPERATIONS)
    assert rows["check_mail_access"]["workflow_visible"] is False
    assert rows["check_mail_access"]["connection_type_ids"] == ["m365_mail_app"]
    for action_id, row in rows.items():
        assert set(row["error_codes"]) <= SAFE_ERROR_CODES, action_id
        if action_id in EFFECTFUL:
            assert row["test_mode_behavior"] == "suppress"
            assert row["mutates_platform"] is True


def test_step_ui_fields_are_declared_inputs() -> None:
    inputs = {
        row["operation_id"]: {item["name"] for item in row["inputs"]}
        for row in _load("operations.yaml")
    }
    for form in _load("step_ui.yaml"):
        assert form["operation_id"] in inputs
        for field in form["fields"]:
            assert field["name"] in inputs[form["operation_id"]], (
                form["operation_id"],
                field["name"],
            )
            assert field["widget_type"] != "resource_picker"


def test_artifact_policies_use_the_runtime_binding_keys() -> None:
    rows = {row["operation_id"]: row for row in _load("artifact_policies.yaml")}
    assert rows["get_attachment"]["outputs"][0]["binding_key"] == download.BINDING_KEY
    for op in ("send_message", "reply_message", "forward_message"):
        [policy] = rows[op]["inputs"]
        assert policy["binding_key"] == outgoing.BINDING_KEY
        assert policy["from_input"] == "attachments"
        assert policy["required"] is False


def test_only_the_allow_listed_writes_may_use_transport_retries() -> None:
    """B6: draft create, attachment add and send never reach the generic retry path."""
    for path in sorted((ROOT / "mail").glob("*.py")):
        source = path.read_text(encoding="utf-8")
        uses = re.findall(r'retry="(idempotent|read)"', source)
        if path.name == "mutations.py":
            assert "idempotent" in uses
        else:
            assert "idempotent" not in uses, path.name
    send_source = (ROOT / "mail" / "send.py").read_text(encoding="utf-8")
    assert 'retry="idempotent"' not in send_source
    assert send_source.count('retry="none"') >= 8
