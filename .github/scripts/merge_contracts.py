"""Merge contracts/parts/<service>/*.yaml into the contract files and extension.yaml.

The parts are the source of truth; tests/test_contracts_merged.py fails on drift.
Run: python .github/scripts/merge_contracts.py .
"""

import sys
from pathlib import Path

import yaml

root = Path(sys.argv[1])
parts = root / "contracts" / "parts"
SERVICES = ["mail", "mail_sync", "files", "excel", "calendar", "teams", "notify"]


def load(name):
    rows = []
    for service in SERVICES:
        path = parts / service / f"{name}.yaml"
        if path.exists():
            data = yaml.safe_load(path.read_text()) or []
            assert isinstance(data, list), path
            rows.extend(data)
    return rows


def check_unique(rows, key, label):
    seen = set()
    for row in rows:
        value = row[key] if isinstance(row, dict) else row
        assert value not in seen, f"duplicate {label}: {value}"
        seen.add(value)


def dump(path, data):
    path.write_text(yaml.safe_dump(data, sort_keys=False, allow_unicode=True, width=100))


operations = load("operations")
check_unique(operations, "operation_id", "operation")
queries = load("resource_queries")
check_unique(queries, "query_id", "query")
step_ui = load("step_ui")
check_unique(step_ui, "operation_id", "step ui")
actions = load("actions")
check_unique(actions, "action_id", "action")
artifacts = load("artifact_policies")
effects = load("external_effects")
triggers = load("workflow_triggers")
events = load("supported_events")

dump(root / "contracts" / "operation_manifest.yaml", {"operations": operations})
dump(root / "contracts" / "resource_queries.yaml", {"queries": queries})
dump(root / "contracts" / "step_ui_manifest.yaml", {"forms": step_ui})
dump(root / "contracts" / "artifact_policies.yaml", {"artifact_policies": artifacts})
actions_path = root / "ui" / "actions" / "actions.yaml"
current = yaml.safe_load(actions_path.read_text()) if actions_path.exists() else {"actions": []}
# Core-owned actions not produced by a service part (post-connect identity).
own = [a for a in current.get("actions", []) if a["action_id"] == "verify_microsoft_identity"]
dump(actions_path, {"actions": own + actions})
manifest_path = root / "extension.yaml"
manifest = yaml.safe_load(manifest_path.read_text())
manifest["external_effects"] = effects
manifest["workflow_triggers"] = triggers
manifest["supported_events"] = events
manifest["artifact_policies"] = artifacts
dump(manifest_path, manifest)
print(
    len(operations),
    "operations,",
    len(queries),
    "queries,",
    len(actions),
    "actions,",
    len(effects),
    "effects,",
    len(triggers),
    "triggers",
)
