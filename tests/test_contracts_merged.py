"""The merged contract files must equal the service fragments in contracts/parts/."""

from __future__ import annotations

from pathlib import Path

import pytest
import yaml

ROOT = Path(__file__).resolve().parents[1]
PARTS = ROOT / "contracts" / "parts"
SERVICES = ["mail", "mail_sync", "files", "excel"]


def _parts(name: str) -> list:
    rows: list = []
    for service in SERVICES:
        path = PARTS / service / f"{name}.yaml"
        if path.exists():
            rows.extend(yaml.safe_load(path.read_text(encoding="utf-8")) or [])
    return rows


def _load(path: str) -> dict:
    return yaml.safe_load((ROOT / path).read_text(encoding="utf-8"))


@pytest.mark.parametrize(
    ("part", "path", "key"),
    [
        ("operations", "contracts/operation_manifest.yaml", "operations"),
        ("resource_queries", "contracts/resource_queries.yaml", "queries"),
        ("step_ui", "contracts/step_ui_manifest.yaml", "forms"),
        ("artifact_policies", "contracts/artifact_policies.yaml", "artifact_policies"),
        ("external_effects", "extension.yaml", "external_effects"),
        ("workflow_triggers", "extension.yaml", "workflow_triggers"),
        ("supported_events", "extension.yaml", "supported_events"),
        ("artifact_policies", "extension.yaml", "artifact_policies"),
    ],
)
def test_merged_file_matches_parts(part: str, path: str, key: str) -> None:
    assert _load(path)[key] == _parts(part), "run python .github/scripts/merge_contracts.py ."


def test_actions_are_the_parts_plus_identity() -> None:
    actions = _load("ui/actions/actions.yaml")["actions"]
    assert actions[0]["action_id"] == "verify_microsoft_identity"
    assert actions[1:] == _parts("actions")


def test_every_connection_type_capability_is_a_known_operation_or_action() -> None:
    known = {
        row["operation_id"] for row in _load("contracts/operation_manifest.yaml")["operations"]
    }
    known |= {row["action_id"] for row in _load("ui/actions/actions.yaml")["actions"]}
    for row in _load("contracts/connection_types.yaml")["connection_types"]:
        missing = set(row["capabilities"]) - known
        assert not missing, (row["connection_type_id"], missing)


def test_no_required_oauth_scopes_on_operations() -> None:
    for row in _load("contracts/operation_manifest.yaml")["operations"]:
        assert "required_oauth_scopes" not in row, row["operation_id"]


def test_mail_poll_mode_defaults_off() -> None:
    for row in _load("contracts/connection_types.yaml")["connection_types"]:
        if row["connection_type_id"].startswith("m365_mail_"):
            mode = row["ingress_schema"]["properties"]["mail_trigger_mode"]
            assert mode["default"] == "off"


def test_oauth_intents_request_files_readwrite_for_files_and_excel() -> None:
    for row in _load("contracts/connection_types.yaml")["connection_types"]:
        if row["connection_type_id"] in {"m365_files_user", "m365_excel_user"}:
            assert "Files.ReadWrite" in row["auth"]["scopes"]["required"]
