"""Workbook addressing, resolution and workbook sessions.

A workbook is always addressed internally as ``drive_id`` + ``drive_item_id``. The
``workbook`` input takes exactly one of two forms (task B5.4):

1. ``{"drive_id": ..., "drive_item_id": ...}``;
2. ``{"sharing_link": "https://..."}``, resolved with
   ``GET /shares/u!{base64url(link)}/driveItem``. The link is only an address: the
   signed-in user must already have access, so no redeem header is sent.
"""

from __future__ import annotations

import base64
import re
from collections.abc import Iterator, Mapping
from contextlib import contextmanager
from dataclasses import dataclass
from typing import Any
from urllib.parse import quote, urlparse

from m365.connection import text
from m365.errors import M365Error
from m365.transport import GraphTransport

from .errors import remap_not_found, without_effect

WORKBOOK_EXTENSION = ".xlsx"
_ID = re.compile(r"^[A-Za-z0-9!_\-]{1,512}$")
_ITEM_SELECT = "id,name,file,parentReference,webUrl"
_LINK_HINT = "use Share → Copy link, or the IDs (drive_id and drive_item_id)"
SESSION_HEADER = "workbook-session-id"


@dataclass(frozen=True)
class Workbook:
    drive_id: str
    drive_item_id: str
    name: str
    web_url: str

    @property
    def base(self) -> str:
        return f"/drives/{segment(self.drive_id)}/items/{segment(self.drive_item_id)}/workbook"

    def summary(self) -> dict[str, Any]:
        return {
            "drive_id": self.drive_id,
            "drive_item_id": self.drive_item_id,
            "name": self.name,
            "web_url": self.web_url,
        }


def segment(value: str) -> str:
    """Percent-encode one Graph path segment (worksheet, table or column name, IDs)."""
    encoded = quote(value, safe="!")
    if ".." in encoded:
        # The shared transport refuses ".." in a path; encode the dots instead.
        encoded = encoded.replace(".", "%2E")
    return encoded


def validated_workbook(value: Any) -> dict[str, str]:
    if not isinstance(value, Mapping):
        raise M365Error(
            "invalid_payload",
            "workbook must be an object with drive_id and drive_item_id, or with sharing_link",
        )
    keys = set(value)
    if keys == {"drive_id", "drive_item_id"}:
        drive_id = text(value.get("drive_id"), limit=512)
        item_id = text(value.get("drive_item_id"), limit=512)
        if not _ID.match(drive_id) or not _ID.match(item_id):
            raise M365Error("invalid_payload", "workbook drive_id or drive_item_id is invalid")
        return {"drive_id": drive_id, "drive_item_id": item_id}
    if keys == {"sharing_link"}:
        link = text(value.get("sharing_link"), limit=2048)
        parsed = urlparse(link)
        if parsed.scheme != "https" or not parsed.hostname or " " in link:
            raise M365Error("invalid_payload", "workbook sharing_link must be an https link")
        return {"sharing_link": link}
    raise M365Error(
        "invalid_payload",
        "workbook takes exactly one form: drive_id + drive_item_id, or sharing_link",
    )


def share_token(link: str) -> str:
    encoded = base64.urlsafe_b64encode(link.encode("utf-8")).decode("ascii").rstrip("=")
    return f"u!{encoded}"


def resolve_workbook(transport: GraphTransport, value: Any) -> Workbook:
    address = validated_workbook(value)
    query = {"$select": _ITEM_SELECT}
    if "sharing_link" in address:
        try:
            item = transport.get_json(
                f"/shares/{share_token(address['sharing_link'])}/driveItem",
                query=query,
                purpose="Microsoft Graph sharing link lookup",
            )
        except M365Error as exc:
            if exc.code in {"item_not_found", "invalid_payload"}:
                raise M365Error(
                    "workbook_link_unsupported",
                    f"Microsoft 365 did not accept this link; {_LINK_HINT}",
                    http_status=exc.http_status,
                ) from exc
            if exc.code == "permission_denied":
                raise M365Error(
                    "permission_denied",
                    "The signed-in user has no access to this workbook; the link is not a grant",
                    http_status=exc.http_status,
                ) from exc
            raise
    else:
        try:
            item = transport.get_json(
                f"/drives/{segment(address['drive_id'])}/items/{segment(address['drive_item_id'])}",
                query=query,
                purpose="Microsoft Graph workbook lookup",
            )
        except M365Error as exc:
            if exc.code == "invalid_payload":
                raise M365Error(
                    "workbook_not_found", "The workbook was not found", http_status=exc.http_status
                ) from exc
            raise remap_not_found(exc, "workbook_not_found", "The workbook was not found") from exc
    return _workbook_from_item(item)


def _workbook_from_item(item: Mapping[str, Any]) -> Workbook:
    parent = item.get("parentReference")
    parent = parent if isinstance(parent, Mapping) else {}
    item_id = text(item.get("id"), limit=512)
    drive_id = text(parent.get("driveId"), limit=512)
    name = text(item.get("name"), limit=400)
    if not _ID.match(item_id) or not _ID.match(drive_id):
        raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid drive item")
    if not isinstance(item.get("file"), Mapping) or not name.lower().endswith(WORKBOOK_EXTENSION):
        raise M365Error(
            "workbook_unsupported", "Only .xlsx workbooks in OneDrive for Business or SharePoint"
        )
    if text(parent.get("driveType"), limit=64).lower() == "personal":
        raise M365Error("workbook_unsupported", "Excel REST does not support OneDrive Personal")
    return Workbook(
        drive_id=drive_id,
        drive_item_id=item_id,
        name=name,
        web_url=text(item.get("webUrl"), limit=2048),
    )


@contextmanager
def workbook_session(transport: GraphTransport, workbook: Workbook) -> Iterator[dict[str, str]]:
    """Open a persistent workbook session and always try to close it.

    Yields the headers to send with every call of the session.
    """
    try:
        response = transport.request(
            "POST",
            f"{workbook.base}/createSession",
            json_body={"persistChanges": True},
            retry="none",
            purpose="Excel workbook session",
        )
    except M365Error as exc:
        exc = without_effect(exc, "Excel workbook session")
        raise remap_not_found(exc, "workbook_not_found", "The workbook was not found") from exc
    session_id = text(response.json().get("id"), limit=8192)
    if not session_id:
        raise M365Error("provider_unavailable", "Excel did not return a workbook session")
    headers = {SESSION_HEADER: session_id}
    try:
        yield headers
    finally:
        try:
            transport.request(
                "POST",
                f"{workbook.base}/closeSession",
                headers=headers,
                retry="none",
                purpose="Excel workbook session close",
            )
        except M365Error:
            # The session expires on its own; a failed close never changes the outcome.
            pass


__all__ = [
    "SESSION_HEADER",
    "Workbook",
    "resolve_workbook",
    "segment",
    "share_token",
    "validated_workbook",
    "workbook_session",
]
