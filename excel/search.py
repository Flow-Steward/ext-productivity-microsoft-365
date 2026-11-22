"""``find_workbooks``: name search for .xlsx workbooks with the Microsoft Search API.

``POST /search/query`` with ``entityTypes: [driveItem]`` searches the signed-in
user's OneDrive and the SharePoint sites they can open. Site and library names are
not part of a search hit, so they are looked up once per distinct drive and site
(best effort, bounded).
"""

from __future__ import annotations

import base64
import hashlib
import json
import re
from collections.abc import Mapping
from typing import Any

from m365.connection import text
from m365.errors import M365Error
from m365.runtime import OperationContext
from m365.transport import GraphTransport

from .errors import without_effect
from .workbook import WORKBOOK_EXTENSION, segment

FIND_FIELDS = frozenset({"name", "limit", "cursor"})
DEFAULT_LIMIT = 25
MAX_LIMIT = 25
MAX_FROM = 1_000
MAX_LOOKUPS = 20
_ID = re.compile(r"^[A-Za-z0-9!_\-]{1,512}$")
_SITE_ID = re.compile(r"^[A-Za-z0-9.\-]{1,253},[0-9a-fA-F\-]{36},[0-9a-fA-F\-]{36}$")
#: Characters that carry KQL syntax; they are replaced by spaces.
_KQL_SYNTAX = re.compile(r"[\"'():*<>=\\{}\[\]^~]")
_CURSOR_VERSION = "fw1"


def _query_text(value: Any) -> str:
    raw = text(value, limit=255)
    cleaned = _KQL_SYNTAX.sub(" ", raw)
    words = []
    for word in cleaned.split():
        # Bare boolean operators would change the query; lower case keeps them as words.
        words.append(word.lower() if word in {"AND", "OR", "NOT", "NEAR"} else word)
    query = " ".join(words)
    if not query:
        raise M365Error("invalid_payload", "name must contain searchable text")
    return query


def _limit(value: Any) -> int:
    if value is None:
        return DEFAULT_LIMIT
    if isinstance(value, bool) or not isinstance(value, int) or not 1 <= value <= MAX_LIMIT:
        raise M365Error("invalid_payload", f"limit must be between 1 and {MAX_LIMIT}")
    return value


def _fingerprint(query: str) -> str:
    return hashlib.sha256(query.encode("utf-8")).hexdigest()[:16]


def encode_cursor(query: str, offset: int) -> str:
    raw = json.dumps({"v": _CURSOR_VERSION, "q": _fingerprint(query), "from": offset})
    return base64.urlsafe_b64encode(raw.encode("utf-8")).decode("ascii").rstrip("=")


def decode_cursor(value: Any, query: str) -> int:
    if value is None or value == "":
        return 0
    raw = text(value, limit=512)
    try:
        decoded = json.loads(base64.urlsafe_b64decode(raw + "=" * (-len(raw) % 4)))
    except (ValueError, TypeError) as exc:
        raise M365Error("invalid_cursor", "The cursor is invalid") from exc
    offset = decoded.get("from") if isinstance(decoded, dict) else None
    if (
        not isinstance(decoded, dict)
        or decoded.get("v") != _CURSOR_VERSION
        or decoded.get("q") != _fingerprint(query)
        or isinstance(offset, bool)
        or not isinstance(offset, int)
        or not 0 < offset <= MAX_FROM
    ):
        raise M365Error("invalid_cursor", "The cursor does not belong to this search")
    return offset


def _container(body: Mapping[str, Any]) -> Mapping[str, Any]:
    responses = body.get("value")
    if not isinstance(responses, list):
        raise M365Error("provider_unavailable", "Microsoft Search returned an invalid response")
    for response in responses:
        containers = response.get("hitsContainers") if isinstance(response, Mapping) else None
        if isinstance(containers, list):
            for container in containers:
                if isinstance(container, Mapping):
                    return container
    return {}


class _Lookups:
    """Best-effort, bounded lookups of drive (library) and site names."""

    def __init__(self, transport: GraphTransport) -> None:
        self._transport = transport
        self._drives: dict[str, Mapping[str, Any]] = {}
        self._sites: dict[str, Mapping[str, Any]] = {}
        self._calls = 0

    def _get(self, path: str, select: str, purpose: str) -> Mapping[str, Any]:
        if self._calls >= MAX_LOOKUPS:
            return {}
        self._calls += 1
        try:
            return self._transport.get_json(path, query={"$select": select}, purpose=purpose)
        except M365Error as exc:
            if exc.code == "authorization_required":
                raise
            return {}

    def drive(self, drive_id: str) -> Mapping[str, Any]:
        if drive_id not in self._drives:
            self._drives[drive_id] = self._get(
                f"/drives/{segment(drive_id)}", "id,name,driveType,webUrl", "Microsoft Graph drive"
            )
        return self._drives[drive_id]

    def site(self, site_id: str) -> Mapping[str, Any]:
        if not _SITE_ID.match(site_id):
            return {}
        if site_id not in self._sites:
            self._sites[site_id] = self._get(
                f"/sites/{segment(site_id).replace('%2C', ',')}",
                "id,displayName,webUrl",
                "Microsoft Graph site",
            )
        return self._sites[site_id]


def _hit_item(hit: Any, lookups: _Lookups) -> dict[str, Any] | None:
    resource = hit.get("resource") if isinstance(hit, Mapping) else None
    if not isinstance(resource, Mapping):
        return None
    name = text(resource.get("name"), limit=400)
    item_id = text(resource.get("id"), limit=512)
    parent = resource.get("parentReference")
    parent = parent if isinstance(parent, Mapping) else {}
    drive_id = text(parent.get("driveId"), limit=512)
    if (
        not name.lower().endswith(WORKBOOK_EXTENSION)
        or not _ID.match(item_id)
        or not _ID.match(drive_id)
    ):
        return None
    drive = lookups.drive(drive_id)
    if text(drive.get("driveType"), limit=64).lower() == "personal":
        return None
    site_id = text(parent.get("siteId"), limit=600)
    site = lookups.site(site_id) if site_id else {}
    return {
        "drive_id": drive_id,
        "drive_item_id": item_id,
        "name": name,
        "site": text(site.get("displayName"), limit=400),
        "site_url": text(site.get("webUrl"), limit=2048),
        "library": text(drive.get("name"), limit=400),
        "web_url": text(resource.get("webUrl"), limit=2048),
        # Ready to map into the ``workbook`` input of the other Excel operations.
        "workbook": {"drive_id": drive_id, "drive_item_id": item_id},
    }


def find_workbooks(context: OperationContext) -> dict[str, Any]:
    raw = context.input
    query = _query_text(raw.get("name"))
    size = _limit(raw.get("limit"))
    offset = decode_cursor(raw.get("cursor"), query)
    transport = context.transport
    request = {
        "requests": [
            {
                "entityTypes": ["driveItem"],
                "query": {"queryString": f"{query} filetype:xlsx"},
                "from": offset,
                "size": size,
            }
        ]
    }
    try:
        # A Search API query is a read (B6), so it may use the transport retries.
        response = transport.request(
            "POST",
            "/search/query",
            json_body=request,
            retry="idempotent",
            purpose="Microsoft Search query",
        )
    except M365Error as exc:
        raise without_effect(exc, "Microsoft Search query") from exc
    container = _container(response.json())
    hits = container.get("hits")
    hits = hits if isinstance(hits, list) else []
    lookups = _Lookups(transport)
    items = [item for item in (_hit_item(hit, lookups) for hit in hits) if item is not None]
    total = container.get("total")
    more = container.get("moreResultsAvailable") is True
    next_offset = offset + len(hits)
    next_cursor = (
        encode_cursor(query, next_offset) if more and hits and next_offset <= MAX_FROM else ""
    )
    return {
        "workbooks": items,
        "next_cursor": next_cursor,
        "total": total if isinstance(total, int) and not isinstance(total, bool) else len(items),
    }


__all__ = ["FIND_FIELDS", "decode_cursor", "encode_cursor", "find_workbooks"]
