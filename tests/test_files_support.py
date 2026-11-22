"""Shared fakes for the files tests: a fake Graph opener and payload builders.

The real ``GraphTransport`` runs against :class:`FakeGraph`, so the tests see the
exact URLs, headers and retry modes the package sends. No network is used.
"""

from __future__ import annotations

import io
import json
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any
from urllib.error import HTTPError

from conftest import HOST_CONTRACT_VERSION, ROOT_STEP_SCOPE

from m365.transport import GraphTransport

GRAPH = "https://graph.microsoft.com/v1.0"
TOKEN = "token-acme-files"  # noqa: S105 - a fake bearer for the fake server
SCOPE = {"account_id": "account-acme", "project_id": "project-acme"}
SITE_URL = "https://acme.sharepoint.com/sites/Finance"
SITE_ID = (
    "acme.sharepoint.com,11111111-aaaa-bbbb-cccc-000000000001,22222222-aaaa-bbbb-cccc-000000000002"
)
DRIVE_ID = "b!acmeFinanceDocs_01"
DOWNLOAD_URL = (
    "https://acme.sharepoint.com/_layouts/15/download.aspx?UniqueId=acme&tempauth=SECRET-DL"
)
UPLOAD_URL = "https://acme.sharepoint.com/_api/v2.0/drives/x/items/y/uploadSession?guid=acme&tempauth=SECRET-UP"

USER_CONNECTION: dict[str, Any] = {
    "connection_id": "conn-files-user",
    "connection_type_id": "m365_files_user",
    "config": {},
    "secrets": {"access_token": TOKEN},
}
USER_SITE_CONNECTION: dict[str, Any] = {
    **USER_CONNECTION,
    "config": {"drive_target": SITE_URL},
}
APP_CONNECTION: dict[str, Any] = {
    "connection_id": "conn-files-app",
    "connection_type_id": "m365_files_app",
    "config": {
        "tenant_id": "acme.onmicrosoft.com",
        "client_id": "11111111-2222-3333-4444-555555555555",
        "site_url": SITE_URL,
        "drive_id": DRIVE_ID,
    },
    "secrets": {"client_secret": "secret-acme"},
}


class FakeResponse:
    def __init__(self, status: int, body: bytes, headers: dict[str, str]) -> None:
        self.status = status
        self.headers = headers
        self._stream = io.BytesIO(body)

    def read(self, size: int = -1) -> bytes:
        return self._stream.read(size)

    def getcode(self) -> int:
        return self.status

    def __enter__(self) -> FakeResponse:
        return self

    def __exit__(self, *_args: Any) -> None:
        return None


@dataclass
class Sent:
    method: str
    url: str
    headers: dict[str, str]
    body: bytes | None
    purpose: str

    def json(self) -> Any:
        return json.loads(self.body or b"{}")


def _contains(part: str) -> Callable[[str], bool]:
    def matches(candidate: str) -> bool:
        return part in candidate

    return matches


Responder = Any  # (status, body) | (status, body, headers) | Exception | callable(Sent)


@dataclass
class FakeGraph:
    routes: list[tuple[str, Callable[[str], bool], list[Responder]]] = field(default_factory=list)
    sent: list[Sent] = field(default_factory=list)

    def on(self, method: str, url: str | Callable[[str], bool], *responses: Responder) -> None:
        """Answer ``method`` on a URL containing ``url`` (or matching a predicate).

        Responses are used in order; the last one repeats.
        """
        matcher = url if callable(url) else _contains(url)
        self.routes.append((method.upper(), matcher, list(responses)))

    def __call__(self, request: Any, *, timeout_seconds: float, purpose: str) -> FakeResponse:
        sent = Sent(
            method=request.get_method(),
            url=request.full_url,
            headers={key.lower(): value for key, value in request.header_items()},
            body=request.data,
            purpose=purpose,
        )
        self.sent.append(sent)
        for method, matcher, responses in self.routes:
            if method == sent.method and matcher(sent.url):
                spec = responses.pop(0) if len(responses) > 1 else responses[0]
                if callable(spec) and not isinstance(spec, BaseException):
                    spec = spec(sent)
                return self._answer(sent, spec)
        raise AssertionError(f"unexpected request {sent.method} {sent.url}")

    @staticmethod
    def _answer(sent: Sent, spec: Responder) -> FakeResponse:
        if isinstance(spec, BaseException):
            raise spec
        status, body, *rest = spec
        headers = dict(rest[0]) if rest else {}
        raw = body if isinstance(body, bytes) else json.dumps(body).encode()
        if status >= 400:
            raise HTTPError(sent.url, status, "error", headers, io.BytesIO(raw))  # type: ignore[arg-type]
        return FakeResponse(status, raw, headers)

    def requests(self, method: str = "", part: str = "") -> list[Sent]:
        return [
            sent
            for sent in self.sent
            if (not method or sent.method == method) and (not part or part in sent.url)
        ]


class RecordingTransport(GraphTransport):
    """The real transport, recording each request's retry mode."""

    def __init__(self, graph: FakeGraph) -> None:
        super().__init__(TOKEN, opener=graph, sleep=lambda _seconds: None)
        self.calls: list[tuple[str, str, str, bool]] = []

    def request(self, method: str, path_or_url: str, **kwargs: Any):  # type: ignore[override]
        self.calls.append(
            (method.upper(), path_or_url, kwargs.get("retry", ""), kwargs.get("authorized", True))
        )
        return super().request(method, path_or_url, **kwargs)


@dataclass
class Harness:
    graph: FakeGraph = field(default_factory=FakeGraph)
    transports: list[RecordingTransport] = field(default_factory=list)

    def factory(self, _connection: Any) -> RecordingTransport:
        transport = RecordingTransport(self.graph)
        self.transports.append(transport)
        return transport

    @property
    def calls(self) -> list[tuple[str, str, str, bool]]:
        return [call for transport in self.transports for call in transport.calls]

    def run(
        self,
        action_id: str,
        operation_input: dict[str, Any],
        *,
        connection: dict[str, Any] | None = None,
        artifacts: dict[str, Any] | None = None,
        test_mode: bool = False,
    ) -> dict[str, Any]:
        from files.operations import OPERATIONS
        from m365.runtime import run_operation

        payload = action_payload(
            action_id,
            operation_input,
            connection=connection or USER_CONNECTION,
            artifacts=artifacts,
            test_mode=test_mode,
        )
        return run_operation(payload, OPERATIONS, transport_factory=self.factory)

    def query(
        self,
        query_id: str,
        *,
        connection: dict[str, Any] | None = None,
        search: str = "",
        cursor: str = "",
        limit: int = 50,
    ) -> dict[str, Any]:
        from files.operations import handle_query

        selected = connection or USER_CONNECTION
        payload = {
            "contract_version": HOST_CONTRACT_VERSION,
            "mode": "query",
            "runtime_context": dict(ROOT_STEP_SCOPE),
            "query": {
                "query_id": query_id,
                "context": {"connection_id": selected["connection_id"]},
                "params": {"search": search, "cursor": cursor, "limit": limit},
            },
            "target": {"connection": selected},
        }
        return handle_query(payload, transport_factory=self.factory)


def action_payload(
    action_id: str,
    operation_input: dict[str, Any],
    *,
    connection: dict[str, Any],
    artifacts: dict[str, Any] | None = None,
    test_mode: bool = False,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "contract_version": HOST_CONTRACT_VERSION,
        "mode": "action",
        "host": {"scope": dict(SCOPE)},
        "runtime_context": {**ROOT_STEP_SCOPE, "test_mode": test_mode},
        "action": {
            "action_id": action_id,
            "input": operation_input,
            "target": {"connection": connection},
        },
    }
    if artifacts is not None:
        payload["artifacts"] = artifacts
    return payload


def upload_artifacts(size: int, *, mime_type: str = "application/pdf") -> dict[str, Any]:
    return {
        "inputs": [
            {
                "artifact_id": "cmr-scan",
                "artifact_handle": "artifact:cmr-scan",
                "role": "input",
                "binding_key": "files_upload_artifact",
                "content_type": mime_type,
                "size_bytes": size,
                "scope": dict(SCOPE),
                "access": {
                    "transport": "presigned_url",
                    "mode": "read",
                    "download_url": "https://artifacts.acme.test/cmr-scan",
                    "max_size_bytes": max(size, 1),
                },
            }
        ]
    }


def drive_item(
    item_id: str,
    name: str,
    *,
    folder: bool = False,
    size: int = 0,
    parent_path: str = "/drives/b!acmeFinanceDocs_01/root:",
    quick_xor: str | None = None,
    sha256: str | None = None,
    extra: dict[str, Any] | None = None,
) -> dict[str, Any]:
    item: dict[str, Any] = {
        "id": item_id,
        "name": name,
        "size": size,
        "webUrl": f"https://acme.sharepoint.com/sites/Finance/Shared%20Documents/{name}",
        "lastModifiedDateTime": "2026-11-22T10:00:00Z",
        "parentReference": {"driveId": DRIVE_ID, "id": "parent-1", "path": parent_path},
    }
    if folder:
        item["folder"] = {"childCount": 0}
    else:
        hashes: dict[str, str] = {}
        if quick_xor:
            hashes["quickXorHash"] = quick_xor
        if sha256:
            hashes["sha256Hash"] = sha256
        item["file"] = {"mimeType": "application/pdf", "hashes": hashes}
    item.update(extra or {})
    return item


def graph_error(status: int, code: str = "generalException") -> tuple[int, dict[str, Any]]:
    return status, {"error": {"code": code, "message": "acme"}}


def test_fake_graph_reports_unexpected_requests() -> None:
    graph = FakeGraph()
    transport = RecordingTransport(graph)
    try:
        transport.get_json("/me/drive", purpose="probe")
    except AssertionError as exc:
        assert "unexpected request GET" in str(exc)
    else:  # pragma: no cover - the fake must refuse unknown routes
        raise AssertionError("unknown routes must fail loudly")
