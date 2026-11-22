"""Opaque search cursors that wrap ``@odata.nextLink`` without exposing it.

A cursor is bound to the query that produced it (folder, filters, limit and the
connection's mailbox) and only continues inside that mailbox on the Graph host.
"""

from __future__ import annotations

import base64
import hashlib
import json
from collections.abc import Mapping
from typing import Any
from urllib.parse import unquote, urlparse

from m365.errors import M365Error
from m365.transport import GRAPH_HOST

CURSOR_PREFIX = "m365m1."
MAX_CURSOR_LENGTH = 8192
_INVALID = "The search cursor is invalid or belongs to a different search"


def fingerprint(base: str, query: Mapping[str, Any]) -> str:
    canonical = json.dumps(
        {"base": base, "query": dict(query)}, sort_keys=True, separators=(",", ":")
    )
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()[:32]


def encode_cursor(next_link: str, *, base: str, query_fingerprint: str) -> str:
    _check_link(next_link, base)
    raw = json.dumps({"v": 1, "f": query_fingerprint, "l": next_link}, separators=(",", ":"))
    token = base64.urlsafe_b64encode(raw.encode("utf-8")).decode("ascii").rstrip("=")
    cursor = CURSOR_PREFIX + token
    if len(cursor) > MAX_CURSOR_LENGTH:
        raise M365Error("response_too_large", "The search cursor is too long")
    return cursor


def decode_cursor(cursor: str, *, base: str, query_fingerprint: str) -> str:
    """Return the Graph continuation link inside ``cursor``."""
    if (
        not isinstance(cursor, str)
        or not cursor.startswith(CURSOR_PREFIX)
        or len(cursor) > MAX_CURSOR_LENGTH
    ):
        raise M365Error("invalid_cursor", _INVALID)
    token = cursor[len(CURSOR_PREFIX) :]
    try:
        raw = base64.urlsafe_b64decode(token + "=" * (-len(token) % 4))
        payload = json.loads(raw.decode("utf-8"))
    except (ValueError, UnicodeDecodeError) as exc:
        raise M365Error("invalid_cursor", _INVALID) from exc
    if (
        not isinstance(payload, dict)
        or payload.get("v") != 1
        or payload.get("f") != query_fingerprint
        or not isinstance(payload.get("l"), str)
    ):
        raise M365Error("invalid_cursor", _INVALID)
    link = payload["l"]
    _check_link(link, base)
    return link


def _check_link(link: str, base: str) -> None:
    """The link must stay on Graph v1.0 and inside the connection's mailbox."""
    parsed = urlparse(link)
    if (
        parsed.scheme != "https"
        or (parsed.hostname or "").lower() != GRAPH_HOST
        or parsed.username
        or parsed.port
    ):
        raise M365Error("invalid_cursor", _INVALID)
    path = unquote(parsed.path).lower()
    if ".." in path or "//" in path:
        raise M365Error("invalid_cursor", _INVALID)
    allowed: tuple[str, ...]
    if base == "/me":
        allowed = ("/v1.0/me/",)
    else:
        mailbox = unquote(base.removeprefix("/users/")).lower()
        allowed = (f"/v1.0/users/{mailbox}/", f"/v1.0/users('{mailbox}')/")
    if not path.startswith(allowed):
        raise M365Error("invalid_cursor", _INVALID)


__all__ = ["decode_cursor", "encode_cursor", "fingerprint"]
