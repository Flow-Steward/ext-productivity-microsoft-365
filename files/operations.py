"""OneDrive for Business / SharePoint Online operations and resource queries.

Both connection types share every operation:

- ``m365_user`` (delegated): config ``drive_target`` (empty = own OneDrive, a
  SharePoint site URL = the site's default library, or a drive ID);
- ``m365_files_app`` (app-only, ``Sites.Selected`` or ``Files.ReadWrite.All``): config
  ``tenant_id``, ``client_id``, ``site_url`` and ``drive_id``.

``search_drive_files`` differs by mode: delegated connections use the Graph drive
search (recursive, name and content), app-only connections filter names inside one
folder without recursion, because ``Sites.Selected`` grants no search.

Transport retry rule (B6): reads use ``retry="read"``, ``create_folder`` uses
``retry="idempotent"`` (a repeat meets a name conflict, which is success), and
``upload_file`` / ``create_share_link`` use ``retry="none"``.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping
from typing import Any
from urllib.parse import quote

from m365.auth import transport_for
from m365.connection import ConnectionContext, mapping, resolve_connection, text
from m365.errors import M365Error, error_response
from m365.runtime import Operation, OperationContext
from m365.transport import GraphTransport

from . import drives, items, paths, transfer
from .drives import DriveTarget

SERVICE = "files"
USER_TYPE = "m365_user"
APP_TYPE = "m365_files_app"
DEFAULT_LIMIT = 50
MAX_LIMIT = 200
LINK_TYPES = ("view", "edit")
LINK_SCOPES = ("organization", "anonymous")

TransportFactory = Callable[[ConnectionContext], GraphTransport]


# -- helpers ------------------------------------------------------------------


def _limit(raw: Mapping[str, Any]) -> int:
    return paths.bounded_int(
        raw.get("limit"), field="limit", minimum=1, maximum=MAX_LIMIT, default=DEFAULT_LIMIT
    )


def _cursor(raw: Mapping[str, Any]) -> str:
    value = raw.get("cursor")
    if value in (None, ""):
        return ""
    return paths.bounded_text(value, field="cursor", maximum=items.MAX_CURSOR_CHARS)


def _folder_not_found(exc: M365Error) -> M365Error:
    if exc.code in paths.NOT_FOUND_CODES:
        return M365Error(
            "folder_not_found", "The folder was not found", http_status=exc.http_status
        )
    return exc


def _folder_items(
    transport: GraphTransport,
    target: DriveTarget,
    *,
    folder_item: str,
    folder_path: str,
    scope: Mapping[str, Any],
    cursor: str,
    limit: int,
    accept: items.Accept,
    purpose: str,
    page_size: int,
) -> tuple[list[dict[str, Any]], str | None]:
    try:
        return items.collect(
            transport,
            first_path=paths.children_ref(
                target.base, folder_item=folder_item, folder_path_value=folder_path
            ),
            first_query={"$select": items.ITEM_SELECT, "$top": page_size},
            scope=scope,
            cursor=cursor,
            limit=limit,
            accept=accept,
            purpose=purpose,
        )
    except M365Error as exc:
        raise _folder_not_found(exc) from exc


def _scope(context: OperationContext, binding: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "operation": context.operation_id,
        "connection_id": context.connection.connection_id,
        "input": dict(binding),
    }


def _name_filter(query: str, extensions: list[str]) -> Callable[[str], bool]:
    needle = query.casefold()

    def matches(name: str) -> bool:
        if needle and needle not in name.casefold():
            return False
        if extensions:
            return "." in name and name.rsplit(".", 1)[1].lower() in extensions
        return True

    return matches


# -- operations ---------------------------------------------------------------


def check_connection(context: OperationContext) -> dict[str, Any]:
    connection = context.connection
    transport = context.transport
    target = drives.resolve_drive(connection, transport)
    site = target.site
    if connection.is_app:
        site = drives.resolve_site(transport, connection.config.get("site_url"))
    drive = drives.get_drive(transport, target)
    if connection.is_app and site is not None:
        if not drives.drive_belongs_to_site(drive["drive_web_url"], site):
            raise M365Error(
                "invalid_connection", "The chosen library does not belong to the configured site"
            )
    return {
        "mode": connection.mode,
        **drive,
        "site_id": site.site_id if site else None,
        "site_name": site.name if site else None,
        "site_web_url": site.web_url if site else None,
    }


def list_drive_items(context: OperationContext) -> dict[str, Any]:
    raw = context.input
    folder_item, folder_path = paths.item_address(
        raw, id_field="folder_id", path_field="folder_path", required=False
    )
    limit = _limit(raw)
    cursor = _cursor(raw)
    binding = {"folder_id": folder_item, "folder_path": folder_path, "limit": limit}
    transport = context.transport
    target = drives.resolve_drive(context.connection, transport)
    found, next_cursor = _folder_items(
        transport,
        target,
        folder_item=folder_item,
        folder_path=folder_path,
        scope=_scope(context, binding),
        cursor=cursor,
        limit=limit,
        accept=items.project_item,
        purpose="Microsoft Graph folder listing",
        page_size=limit,
    )
    return {"items": found, "next_cursor": next_cursor}


def search_drive_files(context: OperationContext) -> dict[str, Any]:
    raw = context.input
    query = paths.bounded_text(raw.get("query"), field="query", maximum=255)
    folder_item, folder_path = paths.item_address(
        raw, id_field="folder_id", path_field="folder_path", required=False
    )
    extensions = paths.file_types(raw["file_types"]) if raw.get("file_types") else []
    limit = _limit(raw)
    cursor = _cursor(raw)
    binding = {
        "query": query,
        "folder_id": folder_item,
        "folder_path": folder_path,
        "file_types": extensions,
        "limit": limit,
    }
    transport = context.transport
    target = drives.resolve_drive(context.connection, transport)
    scope = _scope(context, binding)
    if context.connection.is_app:
        # Sites.Selected grants no search: filter names inside one folder, no recursion.
        name_ok = _name_filter(query, extensions)

        def accept_app(value: dict[str, Any]) -> dict[str, Any] | None:
            item = items.project_item(value)
            return item if not item["is_folder"] and name_ok(item["name"]) else None

        found, next_cursor = _folder_items(
            transport,
            target,
            folder_item=folder_item,
            folder_path=folder_path,
            scope=scope,
            cursor=cursor,
            limit=limit,
            accept=accept_app,
            purpose="Microsoft Graph folder name filter",
            page_size=MAX_LIMIT,
        )
        return {"items": found, "next_cursor": next_cursor, "search_mode": "folder_name_filter"}
    return _delegated_search(
        transport, target, query, folder_item, folder_path, extensions, scope, cursor, limit
    )


def _delegated_search(
    transport: GraphTransport,
    target: DriveTarget,
    query: str,
    folder_item: str,
    folder_path: str,
    extensions: list[str],
    scope: Mapping[str, Any],
    cursor: str,
    limit: int,
) -> dict[str, Any]:
    if ".." in query:
        raise M365Error("invalid_payload", "query must not contain ..")
    scope_path = ""
    if folder_item:
        try:
            folder = items.project_item(
                transport.get_json(
                    paths.item_ref(target.base, item=folder_item),
                    query={"$select": items.ITEM_SELECT},
                    purpose="Microsoft Graph folder",
                )
            )
        except M365Error as exc:
            raise _folder_not_found(exc) from exc
        if not folder["is_folder"] or folder["path"] is None:
            raise M365Error("folder_not_found", "The search folder was not found")
        scope_path = folder["path"]
    elif folder_path != "/":
        scope_path = folder_path
    prefix = scope_path.rstrip("/").casefold() + "/" if scope_path and scope_path != "/" else ""
    extension_ok = _name_filter("", extensions)

    def accept(value: dict[str, Any]) -> dict[str, Any] | None:
        item = items.project_item(value)
        if item["is_folder"] or not extension_ok(item["name"]):
            return None
        if prefix:
            # Graph searches the whole drive; keep only results under the folder.
            if item["path"] is None or not item["path"].casefold().startswith(prefix):
                return None
        return item

    escaped = query.replace("'", "''")
    found, next_cursor = items.collect(
        transport,
        first_path=f"{target.base}/root/search(q='{quote(escaped, safe='')}')",
        first_query={"$top": min(MAX_LIMIT, max(limit, DEFAULT_LIMIT))},
        scope=scope,
        cursor=cursor,
        limit=limit,
        accept=accept,
        purpose="Microsoft Graph drive search",
    )
    return {"items": found, "next_cursor": next_cursor, "search_mode": "drive_search"}


def download_file(context: OperationContext) -> dict[str, Any]:
    item, path = paths.item_address(
        context.input, id_field="item_id", path_field="item_path", required=True
    )
    # The output grant is checked before any token is requested.
    transfer.output_grant_limit(context.payload)
    transport = context.transport
    target = drives.resolve_drive(context.connection, transport)
    return transfer.download_file(transport, context.payload, target, item=item, path=path)


def upload_file(context: OperationContext) -> dict[str, Any]:
    raw = context.input
    handle = paths.artifact_handle_input(raw.get("artifact_handle"))
    folder_item, folder_path = paths.item_address(
        raw, id_field="folder_id", path_field="folder_path", required=False
    )
    filename = paths.name(raw.get("filename"), "filename")
    behavior = paths.choice(
        raw.get("conflict_behavior"),
        field="conflict_behavior",
        allowed=transfer.CONFLICT_BEHAVIORS,
        default="fail",
    )
    # The artifact grant is checked before any token is requested.
    artifact = transfer.input_artifact(context.payload, handle)
    transport = context.transport
    target = drives.resolve_drive(context.connection, transport)
    destination = transfer.UploadTarget(
        target=target,
        folder_item=folder_item,
        folder_path=folder_path,
        filename=filename,
        conflict_behavior=behavior,
    )
    return transfer.upload_file(transport, context.payload, destination, artifact)


def create_folder(context: OperationContext) -> dict[str, Any]:
    raw = context.input
    parent_item, parent_path = paths.item_address(
        raw, id_field="parent_folder_id", path_field="parent_folder_path", required=False
    )
    folder_name = paths.name(raw.get("name"), "name")
    transport = context.transport
    target = drives.resolve_drive(context.connection, transport)
    child = paths.child_ref(
        target.base, parent_item=parent_item, parent_path=parent_path, child=folder_name
    )
    try:
        response = transport.request(
            "POST",
            paths.children_ref(target.base, folder_item=parent_item, folder_path_value=parent_path),
            json_body={
                "name": folder_name,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "fail",
            },
            retry="idempotent",
            purpose="Microsoft Graph create folder",
            ok_statuses=(200, 201),
        )
    except M365Error as exc:
        if exc.code == "name_conflict":
            existing = _existing_folder(transport, child)
            if existing is None:
                raise M365Error(
                    "name_conflict", "An item with this name exists and is not a folder"
                ) from exc
            return {**_folder_result(existing, target), "changed": False}
        if exc.ambiguous:
            existing = _existing_folder(transport, child)
            if existing is None:
                raise
            return {**_folder_result(existing, target), "changed": True}
        raise _folder_not_found(exc) from exc
    return {**_folder_result(items.project_item(response.json()), target), "changed": True}


def _existing_folder(transport: GraphTransport, child: str) -> dict[str, Any] | None:
    try:
        item = items.project_item(
            transport.get_json(
                child, query={"$select": items.ITEM_SELECT}, purpose="Microsoft Graph folder"
            )
        )
    except M365Error:
        return None
    return item if item["is_folder"] else None


def _folder_result(item: Mapping[str, Any], target: DriveTarget) -> dict[str, Any]:
    return {
        "item_id": item["item_id"],
        "drive_id": item["drive_id"] or target.drive_id or None,
        "name": item["name"],
        "path": item["path"],
        "web_url": item["web_url"],
    }


def create_share_link(context: OperationContext) -> dict[str, Any]:
    raw = context.input
    item, path = paths.item_address(raw, id_field="item_id", path_field="item_path", required=True)
    link_type = paths.choice(
        raw.get("link_type"), field="link_type", allowed=LINK_TYPES, default="view"
    )
    scope = paths.choice(
        raw.get("scope"), field="scope", allowed=LINK_SCOPES, default="organization"
    )
    transport = context.transport
    target = drives.resolve_drive(context.connection, transport)
    try:
        response = transport.request(
            "POST",
            f"{paths.item_ref(target.base, item=item, path=path)}/createLink",
            json_body={"type": link_type, "scope": scope},
            retry="none",
            purpose="Microsoft Graph create sharing link",
            ok_statuses=(200, 201),
        )
    except M365Error as exc:
        raise paths.item_missing(exc) from exc
    permission = response.json()
    link = mapping(permission.get("link"))
    url = text(link.get("webUrl"), limit=2048)
    if not url.lower().startswith("https://"):
        raise M365Error(
            "timeout_unknown",
            "Microsoft 365 returned no sharing link; check the item's sharing before retrying",
            ambiguous=True,
        )
    return {
        "permission_id": text(permission.get("id"), limit=256) or None,
        "link_url": url,
        "link_type": text(link.get("type"), limit=32) or link_type,
        "scope": text(link.get("scope"), limit=32) or scope,
        # Graph returns 200 with an existing link of the same type and scope.
        "changed": response.status == 201,
    }


def list_site_libraries(context: OperationContext) -> dict[str, Any]:
    raw = context.input
    connection = context.connection
    site_url = raw.get("site_url")
    if site_url in (None, ""):
        configured = text(connection.config.get("site_url"), limit=2048) or text(
            connection.config.get("drive_target"), limit=2048
        )
        if not configured.lower().startswith("https://"):
            raise M365Error("invalid_payload", "site_url is required")
        site_url = configured
    drives.parse_site_url(site_url)
    transport = context.transport
    site = drives.resolve_site(transport, site_url)
    libraries = drives.list_site_libraries(transport, site)
    return {**site.as_result(), "libraries": libraries, "total": len(libraries)}


_ITEM_FIELDS = frozenset({"folder_id", "folder_path", "limit", "cursor"})

OPERATIONS: dict[str, Operation] = {
    "test_files_connection": Operation("test_files_connection", SERVICE, check_connection),
    "list_drive_items": Operation(
        "list_drive_items", SERVICE, list_drive_items, fields=_ITEM_FIELDS
    ),
    "search_drive_files": Operation(
        "search_drive_files",
        SERVICE,
        search_drive_files,
        fields=_ITEM_FIELDS | {"query", "file_types"},
    ),
    "download_file": Operation(
        "download_file", SERVICE, download_file, fields=frozenset({"item_id", "item_path"})
    ),
    "upload_file": Operation(
        "upload_file",
        SERVICE,
        upload_file,
        effectful=True,
        fields=frozenset(
            {"artifact_handle", "folder_id", "folder_path", "filename", "conflict_behavior"}
        ),
    ),
    "create_folder": Operation(
        "create_folder",
        SERVICE,
        create_folder,
        effectful=True,
        fields=frozenset({"parent_folder_id", "parent_folder_path", "name"}),
    ),
    "create_share_link": Operation(
        "create_share_link",
        SERVICE,
        create_share_link,
        effectful=True,
        fields=frozenset({"item_id", "item_path", "link_type", "scope"}),
    ),
    "list_site_libraries": Operation(
        "list_site_libraries", SERVICE, list_site_libraries, fields=frozenset({"site_url"})
    ),
}


# -- resource queries ---------------------------------------------------------

#: query id → connection type it serves (Core resolves one type per query).
QUERY_TYPES: dict[str, str] = {
    "drive_folders_user": USER_TYPE,
    "drive_folders_app": APP_TYPE,
    "site_libraries_user": USER_TYPE,
    "site_libraries_app": APP_TYPE,
    "site_search_user": USER_TYPE,
}
QUERY_IDS = frozenset(QUERY_TYPES)


def handle_query(
    payload: dict[str, Any], *, transport_factory: TransportFactory = transport_for
) -> dict[str, Any]:
    try:
        query = mapping(payload.get("query"))
        query_id = text(query.get("query_id"), limit=128)
        expected_type = QUERY_TYPES.get(query_id)
        if expected_type is None:
            raise M365Error("invalid_payload", "Unknown Microsoft 365 files resource query")
        params = mapping(query.get("params"))
        search = text(params.get("search", query.get("search")), limit=2048)
        cursor = text(params.get("cursor", query.get("cursor")), limit=items.MAX_CURSOR_CHARS)
        raw_limit = params.get("limit", query.get("limit"))
        limit = (
            raw_limit
            if isinstance(raw_limit, int) and not isinstance(raw_limit, bool) and raw_limit > 0
            else DEFAULT_LIMIT
        )
        limit = min(limit, MAX_LIMIT)
        context_id = text(mapping(query.get("context")).get("connection_id"), limit=256)
        connection = resolve_connection(payload, service=SERVICE, connection_ref=context_id)
        if connection.connection_type_id != expected_type:
            raise M365Error("invalid_connection", "This picker does not serve this connection type")
        transport = transport_factory(connection)
        scope = {"query_id": query_id, "connection_id": connection.connection_id, "search": search}
        if query_id.startswith("drive_folders_"):
            found, next_cursor = _folder_query(transport, connection, search, scope, cursor, limit)
        elif query_id.startswith("site_libraries_"):
            found, next_cursor = _library_query(transport, connection, search), None
        else:
            found, next_cursor = _site_search_query(transport, search, scope, cursor, limit)
        return {
            "ok": True,
            "result": {"items": found, "next_cursor": next_cursor or "", "total": len(found)},
        }
    except M365Error as exc:
        return error_response(exc.code, exc.message, **exc.retry_facts())
    except Exception:
        return error_response("internal_error", "The Microsoft 365 resource query failed")


def _folder_query(
    transport: GraphTransport,
    connection: ConnectionContext,
    search: str,
    scope: Mapping[str, Any],
    cursor: str,
    limit: int,
) -> tuple[list[dict[str, Any]], str | None]:
    """Folders of the root, or of the path typed as search (``/Invoices``); non-recursive."""
    target = drives.resolve_drive(connection, transport)
    if search.startswith("/"):
        folder_path, name_filter = paths.folder_path(search, "search"), ""
    else:
        folder_path, name_filter = "/", search
    name_ok = _name_filter(name_filter, [])

    def accept(value: dict[str, Any]) -> dict[str, Any] | None:
        item = items.project_item(value)
        if not item["is_folder"] or not name_ok(item["name"]):
            return None
        label = item["path"] or item["name"]
        return {"id": item["item_id"], "label": label, "path": item["path"]}

    try:
        return items.collect(
            transport,
            first_path=paths.children_ref(target.base, folder_path_value=folder_path),
            first_query={"$select": items.ITEM_SELECT, "$top": MAX_LIMIT},
            scope=scope,
            cursor=cursor,
            limit=limit,
            accept=accept,
            purpose="Microsoft Graph folder picker",
        )
    except M365Error as exc:
        raise _folder_not_found(exc) from exc


def _library_query(
    transport: GraphTransport, connection: ConnectionContext, search: str
) -> list[dict[str, Any]]:
    """Libraries of the app connection's site, or of the site URL typed as search."""
    if search.lower().startswith("https://"):
        site_url, name_filter = search, ""
    elif connection.is_app:
        site_url, name_filter = text(connection.config.get("site_url"), limit=2048), search
    else:
        site_url = text(connection.config.get("drive_target"), limit=2048)
        name_filter = search
        if not site_url.lower().startswith("https://"):
            return []  # a delegated picker needs a site URL to list libraries
    site = drives.resolve_site(transport, site_url)
    name_ok = _name_filter(name_filter, [])
    return [
        {"id": library["drive_id"], "label": library["name"], "web_url": library["web_url"]}
        for library in drives.list_site_libraries(transport, site)
        if name_ok(library["name"])
    ]


def _site_search_query(
    transport: GraphTransport,
    search: str,
    scope: Mapping[str, Any],
    cursor: str,
    limit: int,
) -> tuple[list[dict[str, Any]], str | None]:
    """Delegated only: ``GET /sites?search=`` (app-only Sites.Selected has no site search)."""
    if not search:
        return [], None

    def accept(value: dict[str, Any]) -> dict[str, Any] | None:
        site = drives.project_site(value)
        if not site.web_url:
            return None
        return {"id": site.web_url, "label": site.name or site.web_url, "site_id": site.site_id}

    return items.collect(
        transport,
        first_path="/sites",
        first_query={"search": search, "$select": drives.SITE_SELECT},
        scope=scope,
        cursor=cursor,
        limit=limit,
        accept=accept,
        purpose="Microsoft Graph site search",
    )


__all__ = ["OPERATIONS", "QUERY_IDS", "QUERY_TYPES", "handle_query"]
