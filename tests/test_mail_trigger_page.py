"""v0.2.7: a page that switches the new-email trigger on or off per mailbox.

Before this, `save_trigger_settings` existed but no page called it, so the trigger
could only be switched through the API.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

from dispatcher import QUERY_HANDLERS, dispatch_runtime
from mail_sync.settings_query import QUERY_ID, handle_trigger_settings_query

ROOT = Path(__file__).resolve().parents[1]
MAIL_TYPES = ["m365_user", "m365_mail_app"]


def _yaml(relative: str) -> Any:
    return yaml.safe_load((ROOT / relative).read_text(encoding="utf-8"))


def _payload(connection: dict[str, Any], resources: dict[str, Any] | None = None) -> dict:
    return {
        "contract_version": "extension_host_v2",
        "mode": "query",
        "runtime_context": {
            "group_id": None,
            "for_each_id": None,
            "step_id": None,
            "step_execution_id": None,
            "resources": resources or {},
        },
        "query": {"query_id": QUERY_ID, "context": {"connection_id": connection["connection_id"]}},
        "target": {"connection": connection},
    }


def _connection(ingress: dict[str, Any] | None = None) -> dict[str, Any]:
    row: dict[str, Any] = {"connection_id": "conn-1", "connection_type": "m365_user"}
    if ingress is not None:
        row["ingress_config"] = ingress
    return row


def test_a_mailbox_without_settings_reads_as_off_on_the_inbox() -> None:
    response = handle_trigger_settings_query(_payload(_connection()))

    assert response["ok"] is True
    assert response["values"] == {"mail_trigger_mode": "off", "mail_trigger_folder": "inbox"}
    assert response["items"] == [{"trigger": "Off", "folder": "inbox", "last_check": ""}]


def test_a_switched_on_mailbox_shows_its_folder_and_last_check() -> None:
    ingress = {
        "mail_trigger_mode": "polling",
        "mail_trigger_folder": "archive",
        "last_inbound_sync_at": "2026-11-23T10:00:00Z",
    }

    response = handle_trigger_settings_query(_payload(_connection(ingress)))

    assert response["values"] == {"mail_trigger_mode": "polling", "mail_trigger_folder": "archive"}
    assert response["items"][0] == {
        "trigger": "On",
        "folder": "archive",
        "last_check": "2026-11-23T10:00:00Z",
    }


def test_the_settings_come_from_the_hydrated_row_of_the_same_connection_only() -> None:
    same = {"connection_id": "conn-1", "ingress_config": {"mail_trigger_mode": "polling"}}
    other = {"connection_id": "conn-2", "ingress_config": {"mail_trigger_mode": "polling"}}

    assert (
        handle_trigger_settings_query(_payload(_connection(), {"project_connection": same}))[
            "values"
        ]["mail_trigger_mode"]
        == "polling"
    )
    assert (
        handle_trigger_settings_query(_payload(_connection(), {"project_connection": other}))[
            "values"
        ]["mail_trigger_mode"]
        == "off"
    )


def test_the_query_needs_a_selected_connection() -> None:
    payload = _payload(_connection())
    payload.pop("target")

    response = handle_trigger_settings_query(payload)

    assert response["ok"] is False
    assert response["error_code"] == "invalid_connection"


def test_the_dispatcher_routes_the_query() -> None:
    assert QUERY_HANDLERS[QUERY_ID] is handle_trigger_settings_query
    response = dispatch_runtime(_payload(_connection({"mail_trigger_mode": "polling"})))
    assert response["values"]["mail_trigger_mode"] == "polling"


def test_the_page_picks_a_mail_connection_and_wires_load_and_save() -> None:
    manifest = _yaml("extension.yaml")
    assert manifest["query_manifest"] == "ui/queries/queries.yaml"
    # Core accepts a query that runs with a connection only from host contract 1.7.0 on.
    assert manifest["runtime"]["compatibility"]["platform_min"] == "1.7.0"
    assert "pages/mail-trigger.yaml" in [
        page["ref"] for page in _yaml("ui/ui_manifest.yaml")["pages"]
    ]

    page = _yaml("ui/pages/mail-trigger.yaml")
    assert page["connection"]["connection_type_ids"] == MAIL_TYPES

    queries = {row["query_id"]: row for row in _yaml("ui/queries/queries.yaml")["queries"]}
    query = queries[QUERY_ID]
    assert query["connection"]["connection_type_ids"] == MAIL_TYPES
    assert query["handler"]["config"]["runtime_resources"]["project_connection"] is True

    form = _yaml("ui/components/mail_trigger_form.yaml")
    assert form["data"] == {"load_query": QUERY_ID, "save_action": "save_trigger_settings"}
    actions = {row["action_id"]: row for row in _yaml("ui/actions/actions.yaml")["actions"]}
    save = actions["save_trigger_settings"]
    assert save["handler"]["operation"] == "project_connection_ingress_patch"
    saved_fields = set(save["handler"]["config"]["fields"])
    form_fields = {field["path"] for field in form["ui_schema"]["fields"]}
    assert form_fields <= saved_fields
    # No interval field: a workflow on a schedule covers any other rhythm.
    assert "mail_poll_interval_seconds" not in form_fields
    modes = {option["value"] for option in form["ui_schema"]["fields"][0]["props"]["options"]}
    assert modes == {"off", "polling"}
