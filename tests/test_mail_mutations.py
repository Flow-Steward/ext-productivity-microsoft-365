"""Idempotent mail writes (B5.1, B6): flags, move and move-to-Deleted-Items."""

from __future__ import annotations

from typing import Any

import pytest
from test_mail_fakes import FakeOutlook, action_payload

from m365.runtime import run_operation
from mail.operations import OPERATIONS


def _run(
    fake: FakeOutlook, action_id: str, operation_input: dict[str, Any], **kwargs: Any
) -> dict[str, Any]:
    payload = action_payload(action_id, operation_input, **kwargs)
    return run_operation(payload, OPERATIONS, transport_factory=lambda _connection: fake)


def test_set_message_flags_patches_absolute_values_with_idempotent_retry() -> None:
    fake = FakeOutlook()
    one, two = fake.add_message(), fake.add_message()
    response = _run(
        fake,
        "set_message_flags",
        {"message_ids": [one.id, two.id, one.id], "is_read": True, "flag_status": "complete"},
    )

    assert response["ok"] is True, response
    assert response["external_effect_status"] == "succeeded"
    assert [row["status"] for row in response["result"]["results"]] == ["succeeded", "succeeded"]
    patches = [call for call in fake.calls if call.method == "PATCH"]
    assert [call.json_body for call in patches] == [
        {"isRead": True, "flag": {"flagStatus": "complete"}}
    ] * 2
    assert all(call.retry == "idempotent" for call in patches)
    assert one.is_read and one.flag == "complete"


@pytest.mark.parametrize(
    "operation_input",
    [
        {"message_ids": ["AAMk-MSG-1"]},
        {"message_ids": ["AAMk-MSG-1"], "flag_status": "starred"},
        {"message_ids": []},
        {"message_ids": ["bad id"], "is_read": True},
        {"message_ids": "AAMk-MSG-1", "is_read": True},
    ],
)
def test_flag_input_is_closed(operation_input: dict[str, Any]) -> None:
    fake = FakeOutlook()
    response = _run(fake, "set_message_flags", operation_input)
    assert response["error_code"] == "invalid_payload"
    assert response["definitely_no_external_effect"] is True
    assert fake.calls == []


def test_move_returns_the_immutable_id_and_reports_unchanged_messages() -> None:
    fake = FakeOutlook()
    inbox_message = fake.add_message("inbox")
    archived = fake.add_message("archive")
    response = _run(
        fake,
        "move_messages",
        {"message_ids": [inbox_message.id, archived.id], "destination_folder": "archive"},
    )

    assert response["ok"] is True, response
    results = response["result"]["results"]
    assert [(row["message_id"], row["changed"]) for row in results] == [
        (inbox_message.id, True),
        (archived.id, False),
    ]
    assert response["result"]["changed_count"] == 1
    assert inbox_message.folder == FakeOutlook.FOLDERS["archive"]
    moves = [call for call in fake.calls if call.url.endswith("/move")]
    assert len(moves) == 1 and moves[0].retry == "idempotent"
    assert moves[0].json_body == {"destinationId": FakeOutlook.FOLDERS["archive"]}


def test_move_to_a_missing_folder_is_folder_not_found() -> None:
    fake = FakeOutlook()
    message = fake.add_message()
    response = _run(
        fake, "move_messages", {"message_ids": [message.id], "destination_folder": "AAMk-F-GONE"}
    )
    assert response["error_code"] == "folder_not_found"
    assert response["definitely_no_external_effect"] is True


def test_delete_moves_to_deleted_items_and_accepts_messages_already_there() -> None:
    fake = FakeOutlook()
    live = fake.add_message("inbox")
    deleted = fake.add_message("deleteditems")
    response = _run(fake, "delete_messages", {"message_ids": [live.id, deleted.id]})

    assert response["ok"] is True, response
    assert response["external_effect_status"] == "succeeded"
    assert [row["changed"] for row in response["result"]["results"]] == [True, False]
    assert live.folder == FakeOutlook.FOLDERS["deleteditems"]
    assert not any(call.method == "DELETE" for call in fake.calls), "never a permanent delete"


def test_delete_of_a_message_already_in_deleted_items_is_success_unchanged() -> None:
    fake = FakeOutlook()
    deleted = fake.add_message("deleteditems")
    response = _run(fake, "delete_messages", {"message_ids": [deleted.id]})
    assert response["ok"] is True
    assert response["result"]["changed"] is False
    assert response["external_effect_status"] == "succeeded"


def test_only_missing_messages_fail_with_message_not_found() -> None:
    fake = FakeOutlook()
    response = _run(fake, "delete_messages", {"message_ids": ["AAMk-GONE"]})
    assert response["ok"] is False
    assert response["error_code"] == "message_not_found"
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True


def test_a_partial_change_is_never_reported_as_no_effect() -> None:
    fake = FakeOutlook()
    live = fake.add_message("inbox")
    response = _run(fake, "delete_messages", {"message_ids": [live.id, "AAMk-GONE"]})

    assert response["ok"] is True
    statuses = [
        (row["requested_message_id"], row["status"]) for row in response["result"]["results"]
    ]
    assert statuses == [(live.id, "succeeded"), ("AAMk-GONE", "failed")]
    assert response["external_effect_status"] == "timeout_unknown"
    assert response["definitely_no_external_effect"] is False


def test_an_unconfirmed_move_stops_the_batch() -> None:
    fake = FakeOutlook()
    first, second = fake.add_message(), fake.add_message()
    fake.fault("move", "dropped")
    response = _run(
        fake,
        "move_messages",
        {"message_ids": [first.id, second.id], "destination_folder": "archive"},
    )

    assert response["ok"] is True
    assert [row["status"] for row in response["result"]["results"]] == [
        "timeout_unknown",
        "not_attempted",
    ]
    assert response["external_effect_status"] == "timeout_unknown"
    assert response["definitely_no_external_effect"] is False


@pytest.mark.parametrize(
    ("action_id", "operation_input"),
    [
        ("set_message_flags", {"message_ids": ["AAMk-MSG-1"], "is_read": True}),
        ("move_messages", {"message_ids": ["AAMk-MSG-1"], "destination_folder": "archive"}),
        ("delete_messages", {"message_ids": ["AAMk-MSG-1"]}),
    ],
)
def test_test_run_suppresses_mailbox_writes(
    action_id: str, operation_input: dict[str, Any]
) -> None:
    fake = FakeOutlook()
    response = _run(fake, action_id, operation_input, test_mode=True)
    assert response["external_effect_status"] == "suppressed"
    assert fake.calls == []
