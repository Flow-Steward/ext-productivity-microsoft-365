"""Drive item projection, opaque cursors and bounded page collection.

A cursor never carries a raw Graph URL to the caller in readable form and cannot
be replayed against another search: it is a base64 envelope of the page link
(relative to the Graph base) and an offset into that page, signed together with
the operation, the connection and the normalised input that produced it.
"""

from __future__ import annotations

import base64
import binascii
import hashlib
import hmac
import json
from collections.abc import Callable, Mapping
from typing import Any
from urllib.parse import unquote

from m365.connection import mapping, text
from m365.errors import M365Error
from m365.transport import GRAPH_BASE, GraphTransport

from .drives import relative_graph_link

ITEM_SELECT = "id,name,size,file,folder,root,parentReference,lastModifiedDateTime,webUrl"
CURSOR_VERSION = 1
CURSOR_DOMAIN = "flowsteward.microsoft-365.files.cursor.v1"
_CURSOR_KEY = b"flowsteward.microsoft-365.files.cursor.integrity.v1"
MAX_CURSOR_CHARS = 8192
MAX_SCAN_PAGES = 10

Accept = Callable[[dict[str, Any]], "dict[str, Any] | None"]


# -- projection ---------------------------------------------------------------


def parent_path(parent: Mapping[str, Any]) -> str | None:
    """``/drives/x/root:/A/B`` → ``/A/B``; ``…/root:`` → ``/``; unknown → ``None``."""
    raw = parent.get("path")
    if not isinstance(raw, str) or "root:" not in raw:
        return None
    tail = unquote(raw.split("root:", 1)[1])
    return "/" + tail.strip("/") if tail.strip("/") else "/"


def project_item(raw: Any) -> dict[str, Any]:
    data = mapping(raw)
    item = text(data.get("id"), limit=256)
    name = text(data.get("name"), limit=512)
    if not item or not name:
        raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid drive item")
    parent = mapping(data.get("parentReference"))
    file_facet = mapping(data.get("file"))
    hashes = mapping(file_facet.get("hashes"))
    is_root = isinstance(data.get("root"), Mapping)
    folder_of = parent_path(parent)
    if is_root:
        path: str | None = "/"
    elif folder_of is None:
        path = None
    else:
        path = folder_of.rstrip("/") + "/" + name
    size = data.get("size")
    child_count = mapping(data.get("folder")).get("childCount")
    web_url = text(data.get("webUrl"), limit=2048)
    return {
        "item_id": item,
        "name": name,
        "drive_id": text(parent.get("driveId"), limit=256) or None,
        "parent_id": text(parent.get("id"), limit=256) or None,
        "path": path,
        "is_folder": isinstance(data.get("folder"), Mapping) or is_root,
        "size": size if isinstance(size, int) and not isinstance(size, bool) else None,
        "mime_type": text(file_facet.get("mimeType"), limit=255) or None,
        "modified_at": text(data.get("lastModifiedDateTime"), limit=64) or None,
        "web_url": web_url if web_url.lower().startswith("https://") else None,
        "child_count": child_count
        if isinstance(child_count, int) and not isinstance(child_count, bool)
        else None,
        "quick_xor_hash": text(hashes.get("quickXorHash"), limit=128) or None,
    }


# -- cursors ------------------------------------------------------------------


def encode_cursor(*, scope: Mapping[str, Any], link: str | None, skip: int) -> str:
    payload = {
        "v": CURSOR_VERSION,
        "d": CURSOR_DOMAIN,
        "scope": dict(scope),
        "link": link,
        "skip": int(skip),
    }
    document = {"p": payload, "s": _sign(payload)}
    cursor = base64.urlsafe_b64encode(_canonical(document)).decode("ascii").rstrip("=")
    if len(cursor) > MAX_CURSOR_CHARS:
        raise M365Error("response_too_large", "The page cursor is too large")
    return cursor


def decode_cursor(value: Any, *, scope: Mapping[str, Any]) -> tuple[str | None, int]:
    if not isinstance(value, str) or not value or len(value) > MAX_CURSOR_CHARS:
        raise M365Error("invalid_cursor", "The cursor is invalid")
    try:
        padded = value + "=" * (-len(value) % 4)
        document = json.loads(base64.urlsafe_b64decode(padded.encode("ascii")))
        payload = document["p"]
        signature = document["s"]
    except (ValueError, KeyError, TypeError, binascii.Error, UnicodeError) as exc:
        raise M365Error("invalid_cursor", "The cursor is invalid") from exc
    if (
        not isinstance(payload, dict)
        or not isinstance(signature, str)
        or not hmac.compare_digest(signature, _sign(payload))
        or payload.get("v") != CURSOR_VERSION
        or payload.get("d") != CURSOR_DOMAIN
    ):
        raise M365Error("invalid_cursor", "The cursor is invalid")
    if payload.get("scope") != dict(scope):
        raise M365Error("invalid_cursor", "The cursor belongs to a different listing or search")
    link = payload.get("link")
    skip = payload.get("skip")
    if (link is not None and (not isinstance(link, str) or not link.startswith("/"))) or (
        isinstance(skip, bool) or not isinstance(skip, int) or skip < 0
    ):
        raise M365Error("invalid_cursor", "The cursor is invalid")
    return link, skip


def _canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=True).encode()


def _sign(payload: Mapping[str, Any]) -> str:
    return hmac.new(_CURSOR_KEY, _canonical(payload), hashlib.sha256).hexdigest()


# -- page collection ----------------------------------------------------------


def collect(
    transport: GraphTransport,
    *,
    first_path: str,
    first_query: Mapping[str, Any] | None,
    scope: Mapping[str, Any],
    cursor: str,
    limit: int,
    accept: Accept,
    purpose: str,
    max_pages: int = MAX_SCAN_PAGES,
) -> tuple[list[dict[str, Any]], str | None]:
    """Read pages until ``limit`` accepted items, the end, or ``max_pages`` pages.

    Returns the items and the next opaque cursor (``None`` at the end). A cursor
    can point into the middle of a page, so client-side filters never lose items.
    """
    link, skip = decode_cursor(cursor, scope=scope) if cursor else (None, 0)
    items: list[dict[str, Any]] = []
    pages = 0
    while True:
        if link is None:
            page = transport.get_json(first_path, query=first_query, purpose=purpose)
        else:
            page = transport.get_json(GRAPH_BASE + link, purpose=purpose)
        values = page.get("value")
        if values is None:
            values = []
        if not isinstance(values, list):
            raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid page")
        raw_next = page.get("@odata.nextLink")
        next_link = relative_graph_link(raw_next) if raw_next else None
        for index in range(skip, len(values)):
            accepted = accept(mapping(values[index]))
            if accepted is None:
                continue
            items.append(accepted)
            if len(items) >= limit:
                if index + 1 < len(values):
                    return items, encode_cursor(scope=scope, link=link, skip=index + 1)
                return items, (
                    encode_cursor(scope=scope, link=next_link, skip=0) if next_link else None
                )
        pages += 1
        skip = 0
        if not next_link:
            return items, None
        link = next_link
        if pages >= max_pages:
            return items, encode_cursor(scope=scope, link=link, skip=0)


__all__ = [
    "ITEM_SELECT",
    "collect",
    "decode_cursor",
    "encode_cursor",
    "parent_path",
    "project_item",
]
