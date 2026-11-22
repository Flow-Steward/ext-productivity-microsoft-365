"""The files contract fragments agree with the code (B4.4, B8 manifest rules)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

yaml = pytest.importorskip("yaml")

PARTS = Path(__file__).resolve().parents[1] / "contracts" / "parts" / "files"
FILES_TYPES = ["m365_user", "m365_files_app"]


def _load(name: str) -> list[dict[str, Any]]:
    data = yaml.safe_load((PARTS / name).read_text(encoding="utf-8"))
    assert isinstance(data, list)
    return data


def test_operation_manifest_matches_the_registry() -> None:
    from files.operations import OPERATIONS
    from m365.errors import SAFE_ERROR_CODES

    entries = {entry["operation_id"]: entry for entry in _load("operations.yaml")}
    assert set(entries) == set(OPERATIONS)
    for operation_id, entry in entries.items():
        operation = OPERATIONS[operation_id]
        assert entry["connection_type_ids"] == FILES_TYPES, operation_id
        assert "required_oauth_scopes" not in entry, operation_id
        names = [row["name"] for row in entry["inputs"]]
        assert names[0] == "connection_ref"
        assert set(names[1:]) == set(operation.fields), operation_id
        assert set(entry["error_codes"]) <= SAFE_ERROR_CODES, operation_id
        assert entry["operation_kind"] == ("action" if operation.effectful else "source")


def test_actions_mirror_the_operations() -> None:
    operations = {entry["operation_id"]: entry for entry in _load("operations.yaml")}
    for action in _load("actions.yaml"):
        operation = operations[action["action_id"]]
        assert [p["name"] for p in action["parameters"]] == [i["name"] for i in operation["inputs"]]
        assert action["result_fields"] == [
            field.removeprefix("result.") for field in operation["result_fields"]
        ]
        assert action["error_codes"] == operation["error_codes"]
        effectful = operation["operation_kind"] == "action"
        assert action["mutates_platform"] is effectful
        assert (action.get("test_mode_behavior") == "suppress") is effectful


def test_every_write_is_an_external_effect_that_test_run_suppresses() -> None:
    from files.operations import OPERATIONS

    effects = _load("external_effects.yaml")
    assert {row["operation_id"] for row in effects} == {
        op_id for op_id, op in OPERATIONS.items() if op.effectful
    }
    for row in effects:
        assert row["idempotency"] == "required"
        assert row["test_mode_behavior"] == "suppress"


def test_one_resource_query_per_connection_type() -> None:
    from files.operations import QUERY_TYPES

    queries = _load("resource_queries.yaml")
    operations = {entry["operation_id"] for entry in _load("operations.yaml")}
    assert {row["query_id"]: row["connection_type_id"] for row in queries} == QUERY_TYPES
    for row in queries:
        assert row["operation_id"] in operations
        assert isinstance(row["connection_type_id"], str)


def test_step_ui_uses_plain_widgets_for_declared_inputs() -> None:
    operations = {entry["operation_id"]: entry for entry in _load("operations.yaml")}
    allowed = {"text", "json", "select", "number", "boolean", "datetime", "artifact_picker"}
    forms = _load("step_ui.yaml")
    assert {form["operation_id"] for form in forms} == set(operations)
    for form in forms:
        inputs = {row["name"] for row in operations[form["operation_id"]]["inputs"]}
        for field in form["fields"]:
            assert field["name"] in inputs, (form["operation_id"], field["name"])
            assert field["widget_type"] in allowed
            assert "resource_query_id" not in field


def test_artifact_policies_use_the_transfer_binding_keys() -> None:
    from files.transfer import DOWNLOAD_BINDING_KEY, UPLOAD_BINDING_KEY

    policies = {row["operation_id"]: row for row in _load("artifact_policies.yaml")}
    assert policies["download_file"]["outputs"][0]["binding_key"] == DOWNLOAD_BINDING_KEY
    assert policies["upload_file"]["inputs"][0]["binding_key"] == UPLOAD_BINDING_KEY
    assert policies["upload_file"]["inputs"][0]["from_input"] == "artifact_handle"
