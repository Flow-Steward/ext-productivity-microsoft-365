"""Read/write access check against a control mailbox (B4.3)."""

from __future__ import annotations

from typing import Any

from test_mail_fakes import APP_CONNECTION, USER_CONNECTION, ScriptedGraph, action_payload

from m365.errors import M365Error
from m365.runtime import run_operation
from mail.access_check import BROAD_MESSAGE
from mail.operations import OPERATIONS

CONTROL = "/users/ceo@acme.example"


def _run(
    graph: ScriptedGraph, operation_input: dict[str, Any] | None = None, connection: Any = None
) -> dict[str, Any]:
    payload = action_payload(
        "check_mail_access", operation_input, connection=connection or APP_CONNECTION
    )
    return run_operation(payload, OPERATIONS, transport_factory=lambda _connection: graph)


def _denied() -> M365Error:
    return M365Error("mailbox_access_denied", "denied", http_status=403)


def test_a_restricted_app_is_denied_read_and_write_in_the_control_mailbox() -> None:
    graph = ScriptedGraph(_denied(), _denied())
    response = _run(graph)

    assert response["ok"] is True, response
    result = response["result"]
    assert result["verdict"] == "restricted"
    assert result["broad_access"] is False
    assert (result["read"], result["write"]) == ("denied", "denied")
    assert result["send_permission_tested"] is False
    assert "Mail.Send" in result["note"]
    assert [(c.method, c.url) for c in graph.calls] == [
        ("GET", f"{CONTROL}/mailFolders/inbox"),
        ("POST", f"{CONTROL}/messages"),
    ]
    assert graph.calls[1].retry == "none"
    assert not any(c.url.endswith(("/send", "/sendMail")) for c in graph.calls), (
        "the check never sends"
    )


def test_broad_access_is_reported_and_the_probe_draft_is_deleted() -> None:
    graph = ScriptedGraph({"id": "F-INBOX"}, {"id": "AAMk-PROBE"}, {})
    response = _run(graph)

    result = response["result"]
    assert result["verdict"] == "broad_access"
    assert result["broad_access"] is True
    assert result["message"] == BROAD_MESSAGE
    assert result["cleanup_failed"] is False
    assert (graph.calls[2].method, graph.calls[2].url) == (
        "DELETE",
        f"{CONTROL}/messages/AAMk-PROBE",
    )


def test_read_only_leak_is_still_broad_access() -> None:
    graph = ScriptedGraph({"id": "F-INBOX"}, _denied())
    result = _run(graph)["result"]
    assert result["verdict"] == "broad_access"
    assert (result["read"], result["write"]) == ("allowed", "denied")


def test_ambiguous_probe_create_is_resolved_by_its_marker() -> None:
    graph = ScriptedGraph(
        _denied(),
        M365Error("timeout_unknown", "lost", http_status=504, ambiguous=True),
        {"value": [{"id": "AAMk-PROBE"}]},
        {},
    )
    result = _run(graph)["result"]
    assert result["write"] == "allowed"
    assert result["cleanup_failed"] is False
    assert "FlowStewardIdempotencyKey" in graph.calls[2].query["$filter"]


def test_an_unknown_control_mailbox_is_inconclusive() -> None:
    graph = ScriptedGraph(
        M365Error("mailbox_not_found", "no", http_status=404),
        M365Error("mailbox_not_found", "no", http_status=404),
    )
    result = _run(graph)["result"]
    assert result["verdict"] == "inconclusive"
    assert result["read_error_code"] == "mailbox_not_found"


def test_the_control_mailbox_comes_from_input_before_config() -> None:
    graph = ScriptedGraph(_denied(), _denied())
    result = _run(graph, {"control_mailbox": "audit@acme.example"})["result"]
    assert result["control_mailbox"] == "audit@acme.example"
    assert graph.calls[0].url == "/users/audit@acme.example/mailFolders/inbox"


def test_the_check_is_only_for_app_only_connections() -> None:
    graph = ScriptedGraph()
    response = _run(graph, {"control_mailbox": "ceo@acme.example"}, connection=USER_CONNECTION)
    assert response["error_code"] == "invalid_connection"
    assert graph.calls == []


def test_the_control_mailbox_must_differ_from_the_connection_mailbox() -> None:
    graph = ScriptedGraph()
    response = _run(graph, {"control_mailbox": "CMR@acme.example"})
    assert response["error_code"] == "invalid_payload"
    assert graph.calls == []
