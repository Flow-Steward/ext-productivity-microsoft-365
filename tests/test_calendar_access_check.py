"""check_calendar_access (task D2): two reads, own calendar allowed and the control
mailbox's calendar refused. It never writes."""

from __future__ import annotations

from typing import Any

from test_calendar_fakes import APP_CONNECTION, USER_CONNECTION, ScriptedGraph, action_payload

from calendar_ops.access_check import BROAD_MESSAGE
from calendar_ops.operations import OPERATIONS
from m365.errors import M365Error
from m365.runtime import run_operation

OWN = "/users/cmr@acme.example/calendar"
CONTROL = "/users/ceo@acme.example/calendar"


def _run(
    graph: ScriptedGraph, operation_input: dict[str, Any] | None = None, connection: Any = None
) -> dict[str, Any]:
    payload = action_payload(
        "check_calendar_access", operation_input, connection=connection or APP_CONNECTION
    )
    return run_operation(payload, OPERATIONS, transport_factory=lambda _connection: graph)


def _denied() -> M365Error:
    return M365Error("mailbox_access_denied", "denied", http_status=403)


def test_a_scoped_app_reads_its_own_calendar_and_is_refused_on_the_control_mailbox() -> None:
    graph = ScriptedGraph({"id": "CAL"}, _denied())
    response = _run(graph)

    assert response["ok"] is True, response
    result = response["result"]
    assert result["verdict"] == "restricted"
    assert result["broad_access"] is False
    assert (result["own_read"], result["control_read"]) == ("allowed", "denied")
    assert result["control_read_error_code"] == "mailbox_access_denied"
    assert result["write_tested"] is False
    assert [(call.method, call.url) for call in graph.calls] == [("GET", OWN), ("GET", CONTROL)]


def test_reading_the_control_mailbox_is_broad_access() -> None:
    result = _run(ScriptedGraph({"id": "CAL"}, {"id": "CEO-CAL"}))["result"]
    assert result["verdict"] == "broad_access"
    assert result["broad_access"] is True
    assert result["message"] == BROAD_MESSAGE


def test_an_unreadable_own_calendar_is_inconclusive() -> None:
    result = _run(ScriptedGraph(_denied(), _denied()))["result"]
    assert result["verdict"] == "inconclusive"
    assert "own calendar" in result["message"]


def test_an_unknown_control_mailbox_is_inconclusive() -> None:
    graph = ScriptedGraph({"id": "CAL"}, M365Error("mailbox_not_found", "no", http_status=404))
    result = _run(graph)["result"]
    assert result["verdict"] == "inconclusive"
    assert result["control_read_error_code"] == "mailbox_not_found"


def test_the_control_mailbox_comes_from_input_before_config() -> None:
    graph = ScriptedGraph({"id": "CAL"}, _denied())
    result = _run(graph, {"control_mailbox": "audit@acme.example"})["result"]
    assert result["control_mailbox"] == "audit@acme.example"
    assert graph.calls[1].url == "/users/audit@acme.example/calendar"


def test_the_check_is_only_for_app_only_connections_and_another_mailbox() -> None:
    graph = ScriptedGraph()
    user = _run(graph, {"control_mailbox": "ceo@acme.example"}, connection=USER_CONNECTION)
    assert user["error_code"] == "invalid_connection"
    same = _run(graph, {"control_mailbox": "CMR@acme.example"})
    assert same["error_code"] == "invalid_payload"
    assert graph.calls == []
