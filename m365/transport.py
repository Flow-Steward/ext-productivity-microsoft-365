"""Microsoft Graph transport: bounded I/O, closed error mapping, explicit retry modes.

Retry rule (task B6): the transport retries 429/503/504 only for ``retry="read"``
and ``retry="idempotent"`` requests. Every other write passes ``retry="none"`` and
handles an ambiguous outcome itself.
"""

from __future__ import annotations

import json
import time
from collections.abc import Callable, Iterable, Mapping
from dataclasses import dataclass
from typing import Any
from urllib.error import HTTPError
from urllib.parse import urlencode, urlparse
from urllib.request import Request

from flowsteward_extension_sdk import PinnedPeerError, open_pinned_url, parse_retry_after

from .errors import M365Error

GRAPH_HOST = "graph.microsoft.com"
GRAPH_BASE = f"https://{GRAPH_HOST}/v1.0"
LOGIN_HOST = "login.microsoftonline.com"
#: Hosts of pre-authenticated transfer URLs Graph hands out (task B6).
TRANSFER_HOST_SUFFIXES = (".sharepoint.com",)
TRANSFER_HOSTS = frozenset({"outlook.office.com"})
CHUNK_BYTES = 64 * 1024
ERROR_RESPONSE_LIMIT = 64 * 1024
DEFAULT_RESPONSE_LIMIT = 4 * 1024 * 1024
RETRYABLE_STATUSES = frozenset({429, 503, 504})
RETRY_MODES = frozenset({"read", "idempotent", "none"})
MAX_ATTEMPTS = 4
MAX_RETRY_WAIT_SECONDS = 30.0
IMMUTABLE_ID_PREFER = 'IdType="ImmutableId"'

Opener = Callable[..., Any]


@dataclass(frozen=True)
class GraphResponse:
    status: int
    headers: dict[str, str]
    body: bytes

    def json(self) -> dict[str, Any]:
        if not self.body:
            return {}
        try:
            value = json.loads(self.body.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise M365Error("provider_unavailable", "Microsoft Graph returned invalid JSON") from exc
        if not isinstance(value, dict):
            raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid response")
        return value


class GraphTransport:
    def __init__(
        self,
        access_token: str,
        *,
        opener: Opener = open_pinned_url,
        timeout_seconds: float = 30.0,
        deadline_seconds: float = 240.0,
        sleep: Callable[[float], None] = time.sleep,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        token = str(access_token or "").strip()
        if not token:
            raise M365Error("authorization_required", "Microsoft 365 authorization is required")
        self._token = token
        self._opener = opener
        self._timeout = max(1.0, float(timeout_seconds))
        self._sleep = sleep
        self._clock = clock
        self._deadline = clock() + max(5.0, float(deadline_seconds))

    # -- public API -------------------------------------------------------

    def remaining_seconds(self) -> float:
        return max(0.0, self._deadline - self._clock())

    def get_json(
        self,
        path: str,
        *,
        query: Mapping[str, Any] | None = None,
        headers: Mapping[str, str] | None = None,
        response_limit: int = DEFAULT_RESPONSE_LIMIT,
        purpose: str,
    ) -> dict[str, Any]:
        return self.request(
            "GET", path, query=query, headers=headers, response_limit=response_limit,
            retry="read", purpose=purpose,
        ).json()

    def request(
        self,
        method: str,
        path_or_url: str,
        *,
        query: Mapping[str, Any] | None = None,
        json_body: Mapping[str, Any] | list[Any] | None = None,
        raw_body: bytes | None = None,
        content_type: str | None = None,
        headers: Mapping[str, str] | None = None,
        response_limit: int = DEFAULT_RESPONSE_LIMIT,
        retry: str,
        purpose: str,
        authorized: bool = True,
        ok_statuses: Iterable[int] = (200, 201, 202, 204),
    ) -> GraphResponse:
        if retry not in RETRY_MODES:
            raise ValueError("unknown retry mode")
        method = str(method or "GET").upper()
        if method != "GET" and retry == "read":
            raise ValueError("read retry is only for GET")
        url = self._url(path_or_url, query=query, authorized=authorized)
        body = raw_body
        request_headers: dict[str, str] = {"Accept": "application/json"}
        if json_body is not None:
            body = json.dumps(json_body, separators=(",", ":")).encode("utf-8")
            request_headers["Content-Type"] = "application/json"
        elif content_type:
            request_headers["Content-Type"] = content_type
        if authorized:
            request_headers["Authorization"] = f"Bearer {self._token}"
        for key, value in (headers or {}).items():
            if key.lower() == "authorization":
                continue
            request_headers[str(key)] = str(value)
        ok = frozenset(int(status) for status in ok_statuses)
        attempt = 0
        while True:
            attempt += 1
            try:
                return self._send(
                    Request(url, data=body, method=method, headers=request_headers),
                    response_limit=response_limit,
                    purpose=purpose,
                    ok_statuses=ok,
                    method=method,
                )
            except M365Error as exc:
                wait = self._retry_wait(exc, retry=retry, attempt=attempt)
                if wait is None:
                    raise
                self._sleep(wait)

    # -- internals --------------------------------------------------------

    def _url(self, path_or_url: str, *, query: Mapping[str, Any] | None, authorized: bool) -> str:
        raw = str(path_or_url or "")
        if raw.startswith("https://"):
            parsed = urlparse(raw)
            host = (parsed.hostname or "").lower()
            if authorized:
                # Absolute Graph URLs only come from @odata.nextLink / deltaLink.
                if host != GRAPH_HOST or not parsed.path.startswith("/v1.0/"):
                    raise M365Error("invalid_cursor", "The continuation link is not a Graph link")
                url = raw
            else:
                if not is_transfer_host(host):
                    raise M365Error(
                        "egress_host_not_allowed",
                        "Microsoft Graph returned a transfer URL on a host that is not allowed",
                    )
                url = raw
        else:
            if not authorized:
                raise ValueError("unauthorized requests need an absolute transfer URL")
            if not raw.startswith("/") or "//" in raw or ".." in raw.split("?", 1)[0]:
                raise M365Error("invalid_payload", "Invalid Microsoft Graph path")
            url = f"{GRAPH_BASE}{raw}"
        if query:
            pairs = [(str(k), str(v)) for k, v in query.items() if v is not None]
            if pairs:
                url = f"{url}{'&' if '?' in url else '?'}{urlencode(pairs)}"
        return url

    def _send(
        self,
        request: Request,
        *,
        response_limit: int,
        purpose: str,
        ok_statuses: frozenset[int],
        method: str,
    ) -> GraphResponse:
        timeout = min(self._timeout, max(1.0, self.remaining_seconds()))
        try:
            with self._opener(request, timeout_seconds=timeout, purpose=purpose) as response:
                status = int(getattr(response, "status", None) or response.getcode() or 200)
                headers = {k.lower(): v for k, v in dict(response.headers or {}).items()}
                body = read_bounded(response, response_limit)
        except HTTPError as exc:
            try:
                error_body = read_bounded(exc, ERROR_RESPONSE_LIMIT)
            except Exception:
                error_body = b""
            raise map_http_error(
                int(exc.code),
                error_body,
                method=method,
                retry_after=_retry_after(exc.headers),
            ) from exc
        except M365Error:
            raise
        except PinnedPeerError as exc:
            raise M365Error("network_failure", f"{purpose} failed") from exc
        except (TimeoutError, OSError) as exc:
            ambiguous = method != "GET"
            raise M365Error(
                "timeout_unknown" if ambiguous else "network_failure",
                f"{purpose} did not complete",
                ambiguous=ambiguous,
            ) from exc
        if status not in ok_statuses:
            raise map_http_error(status, body[:ERROR_RESPONSE_LIMIT], method=method, retry_after=None)
        return GraphResponse(status=status, headers=headers, body=body)

    def _retry_wait(self, exc: M365Error, *, retry: str, attempt: int) -> float | None:
        if retry == "none" or attempt >= MAX_ATTEMPTS:
            return None
        if exc.http_status not in RETRYABLE_STATUSES:
            return None
        wait = exc.retry_after_seconds
        if wait is None:
            wait = min(2.0 ** attempt, MAX_RETRY_WAIT_SECONDS)
        wait = min(float(wait), MAX_RETRY_WAIT_SECONDS)
        if wait >= self.remaining_seconds() - 1.0:
            return None
        return wait


def is_transfer_host(host: str) -> bool:
    host = (host or "").lower()
    return host in TRANSFER_HOSTS or any(
        host.endswith(suffix) and len(host) > len(suffix) for suffix in TRANSFER_HOST_SUFFIXES
    )


def read_bounded(response: Any, limit: int) -> bytes:
    buffer = bytearray()
    while True:
        chunk = response.read(min(CHUNK_BYTES, limit + 1 - len(buffer)))
        if not chunk:
            break
        buffer.extend(chunk)
        if len(buffer) > limit:
            raise M365Error("response_too_large", "The Microsoft 365 response exceeds its safe limit")
    return bytes(buffer)


def _retry_after(headers: Any) -> float | None:
    try:
        value = headers.get("Retry-After") if headers is not None else None
    except Exception:
        return None
    seconds = parse_retry_after(value) if value else None
    return float(seconds) if seconds is not None else None


def graph_error_code(body: bytes) -> tuple[str, str]:
    """Return (Graph error code, inner code) from an error body, both possibly empty."""
    try:
        payload = json.loads(body.decode("utf-8") or "{}")
    except (UnicodeDecodeError, json.JSONDecodeError):
        return "", ""
    error = payload.get("error") if isinstance(payload, dict) else None
    if isinstance(error, str):
        # OAuth token endpoint shape: {"error": "invalid_client", "error_description": "AADSTS..."}
        description = str(payload.get("error_description") or "")
        return error, description.split(":", 1)[0].strip()
    if not isinstance(error, dict):
        return "", ""
    inner = error.get("innerError") or error.get("innererror")
    inner_code = str(inner.get("code") or "") if isinstance(inner, dict) else ""
    return str(error.get("code") or ""), inner_code


_MAILBOX_CODES = frozenset({"ErrorInvalidUser", "MailboxNotEnabledForRESTAPI", "MailboxNotHostedInExchangeOnline", "ResourceNotFound"})
_RBAC_CODES = frozenset({"ErrorAccessDenied", "ErrorExchangeApplicationNotEnabled"})


def map_http_error(
    status: int, body: bytes, *, method: str, retry_after: float | None
) -> M365Error:
    code, inner = graph_error_code(body)
    blob = f"{code} {inner}"
    ambiguous_write = method != "GET" and status in {500, 502, 503, 504}
    if "AADSTS50194" in blob:
        return M365Error("tenant_not_multitenant", "Register the Entra app as multi-tenant", http_status=status)
    if "AADSTS7000222" in blob:
        return M365Error("client_secret_expired", "The app's client secret has expired", http_status=status)
    if "AADSTS7000215" in blob or code == "invalid_client":
        return M365Error("invalid_client", "The app's client ID or secret is not valid", http_status=status)
    if status == 401 or code in {"invalid_grant", "InvalidAuthenticationToken"}:
        return M365Error("authorization_required", "Microsoft 365 authorization is required", http_status=status)
    if status == 403:
        if code in _RBAC_CODES:
            return M365Error("mailbox_access_denied", "The app is not allowed to use this mailbox", http_status=status)
        return M365Error(
            "permission_denied",
            "Microsoft 365 denied access; an administrator must grant the missing permission",
            http_status=status,
        )
    if code in _MAILBOX_CODES and status in {400, 404}:
        return M365Error("mailbox_not_found", "The mailbox was not found", http_status=status)
    if status == 404:
        return M365Error("item_not_found", "The requested Microsoft 365 item was not found", http_status=status)
    if status == 409 or code == "nameAlreadyExists":
        return M365Error("name_conflict", "An item with this name already exists", http_status=status)
    if status == 410 or code in {"SyncStateNotFound", "syncStateNotFound", "resyncRequired"}:
        return M365Error("sync_reset", "The Graph sync state is no longer valid", http_status=status)
    if status == 413:
        return M365Error("file_too_large", "The content is too large for Microsoft 365", http_status=status)
    if status == 429:
        return M365Error("rate_limited", "Microsoft 365 is throttling requests", http_status=status, retry_after_seconds=retry_after)
    if status >= 500:
        if ambiguous_write:
            return M365Error(
                "timeout_unknown", "Microsoft 365 did not confirm the write", http_status=status,
                retry_after_seconds=retry_after, ambiguous=True,
            )
        return M365Error("provider_unavailable", "Microsoft 365 is unavailable", http_status=status, retry_after_seconds=retry_after)
    if status == 400:
        return M365Error("invalid_payload", "Microsoft 365 rejected the request", http_status=status)
    return M365Error("provider_unavailable", "Microsoft 365 returned an unexpected status", http_status=status)


__all__ = [
    "GRAPH_BASE",
    "GRAPH_HOST",
    "IMMUTABLE_ID_PREFER",
    "LOGIN_HOST",
    "GraphResponse",
    "GraphTransport",
    "graph_error_code",
    "is_transfer_host",
    "map_http_error",
    "read_bounded",
]
