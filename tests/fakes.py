"""Fake HTTP opener for transport-level tests (no network)."""

from __future__ import annotations

import io
import json
from collections.abc import Callable
from email.message import Message
from typing import Any
from urllib.error import HTTPError


class FakeResponse:
    def __init__(self, status: int, body: bytes, headers: dict[str, str] | None = None) -> None:
        self.status = status
        self._buffer = io.BytesIO(body)
        message = Message()
        for key, value in (headers or {}).items():
            message[key] = value
        self.headers = message

    def read(self, size: int = -1) -> bytes:
        return self._buffer.read(size)

    def getcode(self) -> int:
        return self.status

    def __enter__(self) -> FakeResponse:
        return self

    def __exit__(self, *exc: object) -> None:
        return None


def http_error(
    url: str, status: int, body: dict[str, Any] | bytes, headers: dict[str, str] | None = None
) -> HTTPError:
    raw = body if isinstance(body, bytes) else json.dumps(body).encode()
    message = Message()
    for key, value in (headers or {}).items():
        message[key] = value
    return HTTPError(url, status, "error", message, io.BytesIO(raw))


class FakeOpener:
    """Replays scripted responses and records every request."""

    def __init__(self, script: list[Callable[[Any], Any] | tuple[int, Any] | Exception]) -> None:
        self.script = list(script)
        self.requests: list[Any] = []

    def __call__(self, request: Any, *, timeout_seconds: float, purpose: str) -> FakeResponse:
        self.requests.append(request)
        if not self.script:
            raise AssertionError(f"unexpected request {request.get_method()} {request.full_url}")
        step = self.script.pop(0)
        if isinstance(step, Exception):
            raise step
        if callable(step):
            return step(request)
        status, body = step
        raw = body if isinstance(body, bytes) else json.dumps(body).encode()
        if status >= 400:
            raise http_error(request.full_url, status, raw)
        return FakeResponse(status, raw)
