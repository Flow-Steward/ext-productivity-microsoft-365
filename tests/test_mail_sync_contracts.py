"""Contract fragments of the new-mail trigger (task B5.2, A2, A6; B8 manifest bullets)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

from mail_sync import state
from mail_sync.machine import EVENT_TYPE

PARTS = Path(__file__).resolve().parents[1] / "contracts" / "parts" / "mail_sync"
MAIL_TYPES = ["m365_user", "m365_mail_app"]


def load(name: str) -> Any:
    return yaml.safe_load((PARTS / name).read_text(encoding="utf-8"))


def action(action_id: str) -> dict[str, Any]:
    (row,) = [row for row in load("actions.yaml") if row["action_id"] == action_id]
    return row


def test_poll_mailbox_is_a_mail_only_background_poll() -> None:
    row = action("poll_mailbox")
    assert row["connection_type_ids"] == MAIL_TYPES
    assert row["handler"]["mode"] == "extension"
    assert row["workflow_visible"] is False
    background = row["contract"]["background"]
    assert set(background) == {"kind", "interval_field", "mode_field", "transport_policy"}
    assert background["kind"] == "connection_poll"
    assert background["mode_field"] == state.MODE_FIELD
    assert background["interval_field"] == state.INTERVAL_FIELD
    assert row["contract"]["serialization"] == {"resource": "mail_sync"}
    assert row["handler"]["config"]["host_mutations"] == ["project_connection_upsert"]


def test_poll_mailbox_declares_no_action_level_provider_secrets() -> None:
    resources = action("poll_mailbox")["handler"]["config"]["runtime_resources"]
    assert "provider_secrets" not in resources
    assert resources["project_connection"] is True


def test_save_trigger_settings_patches_only_the_trigger_fields() -> None:
    row = action("save_trigger_settings")
    assert row["handler"] == {
        "mode": "platform",
        "operation": "project_connection_ingress_patch",
        "config": {
            "connection_type_ids": MAIL_TYPES,
            "fields": [state.MODE_FIELD, state.INTERVAL_FIELD, state.FOLDER_FIELD],
        },
    }
    assert "mail_sync" not in row["handler"]["config"]["fields"]
    assert [parameter["name"] for parameter in row["parameters"]] == row["handler"]["config"][
        "fields"
    ]


def test_mode_values_default_off() -> None:
    assert state.MODE_OFF == "off"
    assert state.MODE_ON == "polling"


def test_trigger_needs_a_connection_and_filters_by_sender_only() -> None:
    (trigger,) = load("workflow_triggers.yaml")
    assert trigger["trigger_type_id"] == "message_received"
    assert trigger["provider_key"] == "flowsteward.microsoft-365"
    assert trigger["event_name"] == EVENT_TYPE
    schema = trigger["config_schema"]
    assert schema["required"] == ["connection_id"]
    assert set(schema["properties"]) == {"connection_id", "from"}
    assert schema["additionalProperties"] is False
    help_text = " ".join(field["help_text"] for field in trigger["ui_schema"]["fields"])
    assert "condition" in help_text and "subject contains" in help_text
    assert "One workflow watches one mailbox" in help_text


def test_trigger_event_is_a_supported_event() -> None:
    assert load("supported_events.yaml") == [EVENT_TYPE]
    (trigger,) = load("workflow_triggers.yaml")
    assert trigger["event_name"] in load("supported_events.yaml")
