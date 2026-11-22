"""The Excel contract fragments agree with the Excel package."""

from __future__ import annotations

from pathlib import Path

import pytest
import yaml

from excel.datasets import INPUT_BINDING_KEY, OUTPUT_BINDING_KEY
from excel.operations import OPERATIONS
from m365.errors import SAFE_ERROR_CODES

PARTS = Path(__file__).resolve().parents[1] / "contracts" / "parts" / "excel"
WRITES = {"write_range", "append_table_rows"}


def _load(name: str) -> list[dict]:
    value = yaml.safe_load((PARTS / name).read_text(encoding="utf-8"))
    assert isinstance(value, list)
    return value


def _by_id(rows: list[dict], key: str) -> dict[str, dict]:
    result = {row[key]: row for row in rows}
    assert len(result) == len(rows)
    return result


def test_operation_fragment_matches_the_registry() -> None:
    operations = _by_id(_load("operations.yaml"), "operation_id")

    assert set(operations) == set(OPERATIONS)
    for operation_id, row in operations.items():
        assert row["connection_type_ids"] == ["m365_excel_user"]
        assert "required_oauth_scopes" not in row
        inputs = {item["name"] for item in row["inputs"]}
        assert inputs - {"connection_ref"} == set(OPERATIONS[operation_id].fields)
        assert set(row["error_codes"]) <= SAFE_ERROR_CODES
        assert row["operation_kind"] == ("action" if operation_id in WRITES else "source")
        assert (operation_id in WRITES) == OPERATIONS[operation_id].effectful
        assert "timeout_unknown" in row["error_codes"] or operation_id not in WRITES
    assert {item["name"] for item in operations["find_workbooks"]["inputs"]} == {
        "connection_ref",
        "name",
        "limit",
        "cursor",
    }
    assert "dedupe_conflict" in operations["append_table_rows"]["error_codes"]
    for operation_id in set(OPERATIONS) - {"find_workbooks"}:
        assert "workbook_link_unsupported" in operations[operation_id]["error_codes"]


def test_actions_fragment_exposes_find_workbooks_and_suppresses_writes() -> None:
    actions = _by_id(_load("actions.yaml"), "action_id")
    operations = _by_id(_load("operations.yaml"), "operation_id")

    assert set(actions) == set(OPERATIONS)
    assert actions["find_workbooks"]["mutates_platform"] is False
    for action_id, action in actions.items():
        parameters = {item["name"] for item in action["parameters"]}
        assert parameters == {item["name"] for item in operations[action_id]["inputs"]}
        assert action["error_codes"] == operations[action_id]["error_codes"]
        if action_id in WRITES:
            assert action["mutates_platform"] is True
            assert action["test_mode_behavior"] == "suppress"


def test_external_effects_cover_exactly_the_writes() -> None:
    effects = _by_id(_load("external_effects.yaml"), "operation_id")

    assert set(effects) == WRITES
    for effect in effects.values():
        assert effect["idempotency"] == "required"
        assert effect["test_mode_behavior"] == "suppress"


def test_artifact_policies_use_the_package_dataset_bindings() -> None:
    policies = _by_id(_load("artifact_policies.yaml"), "operation_id")

    assert policies["read_range"]["outputs"][0]["binding_key"] == OUTPUT_BINDING_KEY
    for operation_id in WRITES:
        policy = policies[operation_id]["inputs"][0]
        assert policy["binding_key"] == INPUT_BINDING_KEY
        assert policy["from_input"] == "dataset_handle"
        assert policy["kind"] == "dataset"


@pytest.mark.parametrize("name", ["step_ui.yaml"])
def test_step_ui_has_no_picker_and_covers_every_input(name) -> None:
    steps = _by_id(_load(name), "operation_id")
    operations = _by_id(_load("operations.yaml"), "operation_id")

    assert set(steps) == set(OPERATIONS)
    for operation_id, step in steps.items():
        fields = {field["name"] for field in step["fields"]}
        assert fields == {item["name"] for item in operations[operation_id]["inputs"]} - {
            "connection_ref"
        }
        for field in step["fields"]:
            assert field["widget_type"] != "resource_picker"
            assert "resource_query_id" not in field
            assert "external_resource_picker" not in field
