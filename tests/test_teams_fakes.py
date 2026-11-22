"""Shared fakes for the Teams tests: payloads and a scripted Graph transport.

This module holds no tests. A script step is a dict (JSON response), an exception
(raised), or a callable that receives the recorded :class:`Call` and returns either.
"""

from __future__ import annotations

import json
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

from m365.errors import M365Error
from m365.transport import GraphResponse

STEP_EXECUTION_ID = "6a1b2c3d-4e5f-4a6b-8c7d-000000000301"
TEAM_ID = "0b8c2f3e-1d2a-4c5b-9e8f-7a6b5c4d3e2f"
CHANNEL_ID = "19:a1b2c3d4e5f6@thread.tacv2"
CHAT_ID = (
    "19:7e1f0000-aaaa-bbbb-cccc-000000000001_7e1f0000-aaaa-bbbb-cccc-000000000002@unq.gbl.spaces"
)
USER_CONNECTION: dict[str, Any] = {
    "connection_id": "conn-acme-teams",
    "connection_type_id": "m365_user",
    "config": {},
    "secrets": {"access_token": "opaque-token"},
}
WEBHOOK_URL = (
    "https://prod-42.westeurope.logic.azure.com:443/workflows/abc123/triggers/manual/paths/"
    "invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=TOPSECRETSIG"
)


def action_payload(
    action_id: str,
    operation_input: dict[str, Any] | None = None,
    *,
    connection: dict[str, Any] | None = None,
    test_mode: bool = False,
    step_execution_id: str = STEP_EXECUTION_ID,
) -> dict[str, Any]:
    return {
        "contract_version": "extension_host_v2",
        "mode": "action",
        "runtime_context": {
            "group_id": None,
            "for_each_id": None,
            "step_id": "notify_ops",
            "step_execution_id": step_execution_id,
            "test_mode": test_mode,
        },
        "action": {
            "action_id": action_id,
            "input": {"connection_ref": "conn-acme-teams", **(operation_input or {})},
            "target": {"connection": dict(connection or USER_CONNECTION)},
        },
    }


def query_payload(
    query_id: str, *, search: str = "", cursor: str = "", connection: dict[str, Any] | None = None
) -> dict[str, Any]:
    conn = dict(connection or USER_CONNECTION)
    return {
        "contract_version": "extension_host_v2",
        "mode": "query",
        "runtime_context": {
            "group_id": None,
            "for_each_id": None,
            "step_id": "picker",
            "step_execution_id": STEP_EXECUTION_ID,
        },
        "query": {
            "query_id": query_id,
            "search": search,
            "cursor": cursor,
            "context": {"connection_id": conn["connection_id"]},
        },
        "target": {"connection": conn},
    }


@dataclass
class Call:
    method: str
    url: str
    query: dict[str, Any]
    json_body: Any
    retry: str


def ambiguous(what: str = "Teams post") -> M365Error:
    return M365Error("timeout_unknown", f"{what} did not complete", http_status=504, ambiguous=True)


def not_found() -> M365Error:
    return M365Error("item_not_found", "not found", http_status=404)


class ScriptedGraph:
    """Answers requests in order from a script; records every call."""

    def __init__(self, *script: Any) -> None:
        self.script = list(script)
        self.calls: list[Call] = []

    def remaining_seconds(self) -> float:
        return 200.0

    def get_json(self, path: str, **kwargs: Any) -> dict[str, Any]:
        return self.request("GET", path, retry="read", **kwargs).json()

    def request(self, method: str, path_or_url: str, **kwargs: Any) -> GraphResponse:
        call = Call(
            method,
            path_or_url,
            dict(kwargs.get("query") or {}),
            kwargs.get("json_body"),
            kwargs["retry"],
        )
        self.calls.append(call)
        if not self.script:
            raise AssertionError(f"unexpected {method} {path_or_url}")
        step = self.script.pop(0)
        if callable(step) and not isinstance(step, Exception):
            step = step(call)
        if isinstance(step, Exception):
            raise step
        return GraphResponse(201 if method == "POST" else 200, {}, json.dumps(step).encode())

    @property
    def posts(self) -> list[Call]:
        return [call for call in self.calls if call.method == "POST"]


def echo_posted(graph: ScriptedGraph, message_id: str = "1700000000001") -> Callable[[Call], Any]:
    """A list response holding the last posted body, as Teams would store it."""

    def step(_call: Call) -> dict[str, Any]:
        posted = graph.posts[-1].json_body
        return {
            "value": [
                {"id": "1699999999999", "body": {"contentType": "html", "content": "<p>other</p>"}},
                {
                    "id": message_id,
                    "createdDateTime": "2026-11-22T10:00:00Z",
                    "webUrl": "https://teams.microsoft.com/l/message/x",
                    "body": posted["body"],
                },
            ]
        }

    return step


__all__ = [
    "CHANNEL_ID",
    "CHAT_ID",
    "STEP_EXECUTION_ID",
    "TEAM_ID",
    "USER_CONNECTION",
    "WEBHOOK_URL",
    "Call",
    "ScriptedGraph",
    "action_payload",
    "ambiguous",
    "echo_posted",
    "not_found",
    "query_payload",
]
