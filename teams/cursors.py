"""Opaque Teams list cursors: a Graph page link plus an offset inside that page.

Filtered lists (chats by topic or member, messages without system events) can stop in
the middle of a Graph page; the cursor resumes at the same row. A cursor is bound to
the query that produced it and only continues under the same Graph path.
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

CURSOR_PREFIX = "m365t1."
MAX_CURSOR_LENGTH = 8192
_INVALID = "The cursor is invalid or belongs to a different list"


def fingerprint(scope: Mapping[str, Any]) -> str:
    canonical = json.dumps(dict(scope), sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()[:32]


def encode_cursor(link: str, offset: int, *, path_prefix: str, query_fingerprint: str) -> str:
    _check_link(link, path_prefix)
    raw = json.dumps(
        {"v": 1, "f": query_fingerprint, "l": link, "o": int(offset)}, separators=(",", ":")
    )
    cursor = CURSOR_PREFIX + base64.urlsafe_b64encode(raw.encode("utf-8")).decode("ascii").rstrip(
        "="
    )
    if len(cursor) > MAX_CURSOR_LENGTH:
        raise M365Error("response_too_large", "The cursor is too long")
    return cursor


def decode_cursor(cursor: str, *, path_prefix: str, query_fingerprint: str) -> tuple[str, int]:
    """Return ``(Graph page link, offset)`` inside ``cursor``."""
    if (
        not isinstance(cursor, str)
        or not cursor.startswith(CURSOR_PREFIX)
        or len(cursor) > MAX_CURSOR_LENGTH
    ):
        raise M365Error("invalid_cursor", _INVALID)
    token = cursor[len(CURSOR_PREFIX) :]
    try:
        payload = json.loads(base64.urlsafe_b64decode(token + "=" * (-len(token) % 4)))
    except (ValueError, UnicodeDecodeError) as exc:
        raise M365Error("invalid_cursor", _INVALID) from exc
    if (
        not isinstance(payload, dict)
        or payload.get("v") != 1
        or payload.get("f") != query_fingerprint
        or not isinstance(payload.get("l"), str)
        or not isinstance(payload.get("o"), int)
        or isinstance(payload.get("o"), bool)
        or not 0 <= payload["o"] <= 1000
    ):
        raise M365Error("invalid_cursor", _INVALID)
    _check_link(payload["l"], path_prefix)
    return payload["l"], payload["o"]


def _check_link(link: str, path_prefix: str) -> None:
    """The link must stay on Graph v1.0 under ``/v1.0{path_prefix}``."""
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
    prefix = f"/v1.0{unquote(path_prefix).lower()}"
    if path != prefix and not path.startswith(prefix.rstrip("/") + "/"):
        raise M365Error("invalid_cursor", _INVALID)


__all__ = ["CURSOR_PREFIX", "decode_cursor", "encode_cursor", "fingerprint"]
