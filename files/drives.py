"""Choosing the drive of a files connection, and SharePoint site/library discovery.

Delegated (``m365_user``), config ``drive_target``:

- empty → the signed-in user's OneDrive (``/me/drive``);
- a SharePoint site URL → the site is resolved, then its default library
  (``/sites/{site-id}/drive``);
- anything else → a drive ID (``/drives/{id}``).

App-only (``m365_files_app``), config ``site_url`` + ``drive_id``: there is no site
search with ``Sites.Selected``, so the admin enters the site URL and picks one of the
site's libraries (``list_site_libraries``); every call then uses ``/drives/{drive_id}``.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any
from urllib.parse import quote, unquote, urlparse

from m365.connection import ConnectionContext, mapping, text
from m365.errors import M365Error
from m365.transport import GRAPH_BASE, GraphTransport

from . import paths

SITE_SELECT = "id,displayName,name,webUrl"
DRIVE_SELECT = "id,name,driveType,webUrl,description,quota"
LIBRARY_PAGES_MAX = 5
SITE_HOST_SUFFIX = ".sharepoint.com"


@dataclass(frozen=True)
class Site:
    site_id: str
    name: str
    web_url: str

    def as_result(self) -> dict[str, Any]:
        return {"site_id": self.site_id, "site_name": self.name, "site_web_url": self.web_url}


@dataclass(frozen=True)
class DriveTarget:
    #: Graph path prefix of the drive: ``/me/drive``, ``/sites/{id}/drive`` or ``/drives/{id}``.
    base: str
    #: Known up front only for an explicit drive ID; otherwise read from item responses.
    drive_id: str = ""
    site: Site | None = None


def parse_site_url(value: Any, field: str = "site_url") -> tuple[str, str]:
    """Return ``(hostname, server-relative path)`` of a SharePoint Online site URL."""
    raw = value.strip() if isinstance(value, str) else ""
    if not raw or len(raw) > 2048 or any(ord(char) < 32 for char in raw):
        raise M365Error("invalid_payload", f"{field} must be a SharePoint site URL")
    parsed = urlparse(raw)
    host = (parsed.hostname or "").lower()
    if (
        parsed.scheme != "https"
        or parsed.username
        or parsed.password
        or parsed.port
        or parsed.query
        or parsed.fragment
        or not host.endswith(SITE_HOST_SUFFIX)
        or len(host) <= len(SITE_HOST_SUFFIX)
    ):
        raise M365Error(
            "invalid_payload",
            f"{field} must be a site URL like https://contoso.sharepoint.com/sites/Finance",
        )
    segments = [unquote(segment) for segment in parsed.path.split("/") if segment]
    if any(segment in {".", ".."} or "/" in segment for segment in segments):
        raise M365Error("invalid_payload", f"{field} has an invalid path")
    return host, "/".join(segments)


def resolve_site(transport: GraphTransport, site_url: Any) -> Site:
    host, path = parse_site_url(site_url)
    graph_path = f"/sites/{host}"
    if path:
        graph_path += ":/" + "/".join(quote(segment, safe="") for segment in path.split("/"))
    try:
        raw = transport.get_json(
            graph_path, query={"$select": SITE_SELECT}, purpose="Microsoft Graph site lookup"
        )
    except M365Error as exc:
        if exc.code in paths.NOT_FOUND_CODES or exc.code == "invalid_payload":
            raise M365Error(
                "site_not_found", "The SharePoint site was not found", http_status=exc.http_status
            ) from exc
        raise
    return project_site(raw)


def project_site(raw: Any) -> Site:
    data = mapping(raw)
    site_id = text(data.get("id"), limit=512)
    if not site_id:
        raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid site")
    return Site(
        site_id=site_id,
        name=text(data.get("displayName"), limit=512) or text(data.get("name"), limit=512),
        web_url=_https(data.get("webUrl")),
    )


def site_base(site: Site) -> str:
    return f"/sites/{quote(site.site_id, safe=',.-')}"


def list_site_libraries(transport: GraphTransport, site: Site) -> list[dict[str, Any]]:
    """All document libraries of a site; several libraries always need an explicit choice."""
    libraries: list[dict[str, Any]] = []
    page = transport.get_json(
        f"{site_base(site)}/drives",
        query={"$select": "id,name,driveType,webUrl,description"},
        purpose="Microsoft Graph site libraries",
    )
    for _ in range(LIBRARY_PAGES_MAX):
        for raw in page.get("value") or []:
            libraries.append(project_library(raw))
        next_link = page.get("@odata.nextLink")
        if not isinstance(next_link, str) or not next_link:
            break
        page = transport.get_json(next_link, purpose="Microsoft Graph site libraries")
    return libraries


def project_library(raw: Any) -> dict[str, Any]:
    data = mapping(raw)
    drive = text(data.get("id"), limit=256)
    if not drive:
        raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid library")
    return {
        "drive_id": drive,
        "name": text(data.get("name"), limit=512),
        "drive_type": text(data.get("driveType"), limit=64),
        "web_url": _https(data.get("webUrl")),
        "description": text(data.get("description"), limit=1024),
    }


def project_drive(raw: Any) -> dict[str, Any]:
    data = mapping(raw)
    drive = text(data.get("id"), limit=256)
    if not drive:
        raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid drive")
    quota = mapping(data.get("quota"))
    quota_result: dict[str, Any] | None = None
    if quota:
        quota_result = {
            "total": _int_or_none(quota.get("total")),
            "used": _int_or_none(quota.get("used")),
            "remaining": _int_or_none(quota.get("remaining")),
            "state": text(quota.get("state"), limit=32) or None,
        }
    return {
        "drive_id": drive,
        "drive_name": text(data.get("name"), limit=512),
        "drive_type": text(data.get("driveType"), limit=64),
        "drive_web_url": _https(data.get("webUrl")),
        "quota": quota_result,
    }


def resolve_drive(connection: ConnectionContext, transport: GraphTransport) -> DriveTarget:
    if connection.is_app:
        drive = app_drive_id(connection)
        return DriveTarget(base=f"/drives/{paths.quote_id(drive)}", drive_id=drive)
    target = text(connection.config.get("drive_target"), limit=2048)
    if not target:
        return DriveTarget(base="/me/drive")
    if target.lower().startswith("https://"):
        site = resolve_site(transport, target)
        return DriveTarget(base=f"{site_base(site)}/drive", site=site)
    try:
        drive = paths.drive_id(target, "drive_target")
    except M365Error as exc:
        raise M365Error(
            "invalid_connection",
            "drive_target must be empty, a SharePoint site URL or a drive ID",
        ) from exc
    return DriveTarget(base=f"/drives/{paths.quote_id(drive)}", drive_id=drive)


def app_drive_id(connection: ConnectionContext) -> str:
    site_url = text(connection.config.get("site_url"), limit=2048)
    drive = text(connection.config.get("drive_id"), limit=256)
    if not site_url or not drive:
        raise M365Error(
            "invalid_connection",
            "The app-only files connection needs a site URL and a chosen library (drive_id)",
        )
    try:
        parse_site_url(site_url)
        return paths.drive_id(drive, "drive_id")
    except M365Error as exc:
        raise M365Error("invalid_connection", exc.message) from exc


def get_drive(transport: GraphTransport, target: DriveTarget) -> dict[str, Any]:
    try:
        raw = transport.get_json(
            target.base, query={"$select": DRIVE_SELECT}, purpose="Microsoft Graph drive"
        )
    except M365Error as exc:
        if exc.code in paths.NOT_FOUND_CODES:
            raise M365Error(
                "drive_not_found", "The drive or library was not found", http_status=exc.http_status
            ) from exc
        raise
    return project_drive(raw)


def drive_belongs_to_site(drive_web_url: str, site: Site) -> bool:
    """A library's web URL sits under its site's web URL (compared decoded, case-insensitive)."""
    if not drive_web_url or not site.web_url:
        return True  # nothing to compare; the drive ID itself is authoritative
    drive_url = unquote(drive_web_url).rstrip("/").lower()
    site_url = unquote(site.web_url).rstrip("/").lower()
    return drive_url == site_url or drive_url.startswith(site_url + "/")


def relative_graph_link(link: Any) -> str:
    """Strip the Graph base from an ``@odata.nextLink``; refuse anything else."""
    if not isinstance(link, str) or not link.startswith(GRAPH_BASE + "/") or len(link) > 6000:
        raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid page link")
    return link[len(GRAPH_BASE) :]


def _https(value: Any) -> str:
    url = text(value, limit=2048)
    return url if url.lower().startswith("https://") else ""


def _int_or_none(value: Any) -> int | None:
    return value if isinstance(value, int) and not isinstance(value, bool) else None


__all__ = [
    "DriveTarget",
    "Site",
    "app_drive_id",
    "drive_belongs_to_site",
    "get_drive",
    "list_site_libraries",
    "parse_site_url",
    "project_drive",
    "project_library",
    "project_site",
    "relative_graph_link",
    "resolve_drive",
    "resolve_site",
    "site_base",
]
