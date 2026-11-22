"""Read operations: connection test, folders, message search and message details."""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any

from m365.connection import mapping
from m365.errors import M365Error
from m365.runtime import OperationContext

from .common import (
    BODY_PREFER_HTML,
    BODY_PREFER_TEXT,
    WELL_KNOWN_FOLDERS,
    Mailbox,
    address_out,
    addresses_out,
    bool_input,
    datetime_input,
    email_address,
    flag_status_out,
    folder_ref,
    graph_id,
    int_input,
    not_found_as,
    odata_string,
    segment,
    text_input,
)
from .cursors import decode_cursor, encode_cursor, fingerprint

MAX_FOLDERS = 1000
MAX_FOLDER_DEPTH = 10
MAX_BODY_BYTES = 1024 * 1024
MAX_HEADERS = 200
MAX_HEADER_VALUE = 4096
MAX_ATTACHMENT_ROWS = 250
MESSAGE_RESPONSE_LIMIT = 8 * 1024 * 1024
_FOLDER_SELECT = "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount"
_SUMMARY_SELECT = (
    "id,internetMessageId,conversationId,subject,from,toRecipients,ccRecipients,"
    "receivedDateTime,sentDateTime,isRead,isDraft,hasAttachments,importance,flag,"
    "parentFolderId,bodyPreview"
)
_MESSAGE_NOT_FOUND = "The message was not found in this mailbox"
_FOLDER_NOT_FOUND = "The mail folder was not found in this mailbox"
#: Folders listed as picker shortcuts by their well-known name.
PICKER_WELL_KNOWN = ("inbox", "archive", "sentitems", "drafts", "deleteditems", "junkemail")


def mailbox(ctx: OperationContext) -> Mailbox:
    return Mailbox.for_connection(ctx.transport, ctx.connection)


# -- test_connection --------------------------------------------------------------


def connection_test(ctx: OperationContext) -> dict[str, Any]:
    client = mailbox(ctx)
    display_name: str | None = None
    address = ctx.connection.mailbox
    if client.base == "/me":
        me = ctx.transport.request(
            "GET",
            "/me",
            query={"$select": "displayName,mail,userPrincipalName"},
            headers=Mailbox.headers(),
            response_limit=64 * 1024,
            retry="read",
            purpose="Microsoft 365 profile",
        ).json()
        address = str(me.get("mail") or me.get("userPrincipalName") or "")
        display_name = str(me.get("displayName") or "") or None
    with not_found_as("mailbox_not_found", "The mailbox was not found"):
        inbox = client.get(
            "/mailFolders/inbox",
            query={"$select": "id,totalItemCount,unreadItemCount"},
            response_limit=64 * 1024,
            purpose="Outlook Inbox",
        )
    return {
        "mailbox_address": address,
        "display_name": display_name,
        "auth_mode": ctx.connection.mode,
        "inbox_total": _count(inbox.get("totalItemCount")),
        "inbox_unread": _count(inbox.get("unreadItemCount")),
    }


def _count(value: Any) -> int:
    return value if isinstance(value, int) and not isinstance(value, bool) and value >= 0 else 0


# -- list_mail_folders ------------------------------------------------------------


def list_mail_folders(ctx: OperationContext) -> dict[str, Any]:
    client = mailbox(ctx)
    well_known_ids: dict[str, str] = {}
    for name in PICKER_WELL_KNOWN:
        try:
            row = client.get(
                f"/mailFolders/{name}",
                query={"$select": "id"},
                response_limit=64 * 1024,
                purpose="Outlook folder",
            )
        except M365Error as exc:
            if exc.code in {"item_not_found", "invalid_payload"}:
                continue  # e.g. no Archive folder in this mailbox
            raise
        folder_id = row.get("id")
        if isinstance(folder_id, str) and folder_id:
            well_known_ids[folder_id] = name
    folders: list[dict[str, Any]] = []
    truncated = _walk(
        client, "/mailFolders", path="", depth=0, parent="", out=folders, well_known=well_known_ids
    )
    shortcuts = [
        {
            "folder_id": name,
            "display_name": WELL_KNOWN_FOLDERS[name],
            "path": WELL_KNOWN_FOLDERS[name],
            "well_known_name": name,
            "is_well_known_shortcut": True,
        }
        for name in PICKER_WELL_KNOWN
        if name in set(well_known_ids.values())
    ]
    return {"well_known": shortcuts, "folders": folders, "truncated": truncated}


def _walk(
    client: Mailbox,
    collection: str,
    *,
    path: str,
    depth: int,
    parent: str,
    out: list[dict[str, Any]],
    well_known: Mapping[str, str],
) -> bool:
    rows = client.list_all(
        collection,
        query={"$select": _FOLDER_SELECT, "$top": "100"},
        max_items=MAX_FOLDERS,
        purpose="Outlook folders",
    )
    truncated = False
    for row in rows:
        if len(out) >= MAX_FOLDERS:
            return True
        folder_id = str(row.get("id") or "")
        name = str(row.get("displayName") or "")
        if not folder_id:
            continue
        full_path = f"{path} / {name}" if path else name
        children = _count(row.get("childFolderCount"))
        out.append(
            {
                "folder_id": folder_id,
                "display_name": name,
                "path": full_path,
                "parent_folder_id": parent or None,
                "well_known_name": well_known.get(folder_id),
                "depth": depth,
                "total_count": _count(row.get("totalItemCount")),
                "unread_count": _count(row.get("unreadItemCount")),
                "child_folder_count": children,
                "is_well_known_shortcut": False,
            }
        )
        if children:
            if depth + 1 >= MAX_FOLDER_DEPTH:
                truncated = True
                continue
            truncated = (
                _walk(
                    client,
                    f"/mailFolders/{segment(folder_id)}/childFolders",
                    path=full_path,
                    depth=depth + 1,
                    parent=folder_id,
                    out=out,
                    well_known=well_known,
                )
                or truncated
            )
    return truncated


def folder_items(result: Mapping[str, Any]) -> list[dict[str, Any]]:
    """Picker rows: well-known shortcuts first, then every folder by path."""
    items: list[dict[str, Any]] = []
    for row in result.get("well_known") or []:
        items.append(
            {"folder_id": row["folder_id"], "label": f"{row['display_name']} (well-known)"}
        )
    for row in result.get("folders") or []:
        items.append({"folder_id": row["folder_id"], "label": row["path"]})
    return items


# -- search_messages ----------------------------------------------------------------

SEARCH_FIELDS = frozenset(
    {
        "folder",
        "from",
        "subject_contains",
        "unread",
        "has_attachments",
        "received_after",
        "received_before",
        "query",
        "limit",
        "cursor",
    }
)


def _search_query(raw: Mapping[str, Any]) -> dict[str, Any]:
    folder = folder_ref(raw.get("folder"), "folder", default="inbox")
    sender = raw.get("from")
    normalized: dict[str, Any] = {
        "folder": folder,
        "from": email_address(sender, "from") if sender not in (None, "") else "",
        "subject_contains": text_input(
            raw.get("subject_contains"), "subject_contains", max_length=255
        ),
        "unread": bool_input(raw.get("unread"), "unread"),
        "has_attachments": bool_input(raw.get("has_attachments"), "has_attachments"),
        "received_after": datetime_input(raw.get("received_after"), "received_after"),
        "received_before": datetime_input(raw.get("received_before"), "received_before"),
        "query": text_input(raw.get("query"), "query", max_length=255),
        "limit": int_input(raw.get("limit"), "limit", default=25, minimum=1, maximum=100),
    }
    if normalized["query"]:
        if any(char in normalized["query"] for char in '"\\'):
            raise M365Error(
                "invalid_payload", "query must not contain double quotes or backslashes"
            )
        filters = (
            "from",
            "subject_contains",
            "unread",
            "has_attachments",
            "received_after",
            "received_before",
        )
        if any(normalized[name] not in ("", None) for name in filters):
            raise M365Error(
                "search_unsupported",
                "Microsoft Graph cannot combine a free-text query with filters; use one or the other",
            )
    if (
        normalized["received_after"]
        and normalized["received_before"]
        and normalized["received_after"] >= normalized["received_before"]
    ):
        raise M365Error("invalid_payload", "received_after must be earlier than received_before")
    return normalized


def _graph_query(normalized: Mapping[str, Any]) -> dict[str, str]:
    query = {"$select": _SUMMARY_SELECT, "$top": str(normalized["limit"])}
    if normalized["query"]:
        query["$search"] = f'"{normalized["query"]}"'
        return query
    # Graph requires an $orderby property to lead the $filter, so the date clause comes first.
    clauses = [f"receivedDateTime ge {normalized['received_after'] or '1900-01-01T00:00:00Z'}"]
    if normalized["received_before"]:
        clauses.append(f"receivedDateTime lt {normalized['received_before']}")
    if normalized["from"]:
        clauses.append(f"from/emailAddress/address eq {odata_string(normalized['from'])}")
    if normalized["subject_contains"]:
        clauses.append(f"contains(subject,{odata_string(normalized['subject_contains'])})")
    if normalized["unread"] is not None:
        clauses.append(f"isRead eq {'false' if normalized['unread'] else 'true'}")
    if normalized["has_attachments"] is not None:
        clauses.append(f"hasAttachments eq {'true' if normalized['has_attachments'] else 'false'}")
    query["$filter"] = " and ".join(clauses)
    query["$orderby"] = "receivedDateTime desc"
    return query


def search_messages(ctx: OperationContext) -> dict[str, Any]:
    client = mailbox(ctx)
    normalized = _search_query(ctx.input)
    key = fingerprint(client.base, normalized)
    cursor = text_input(ctx.input.get("cursor"), "cursor", max_length=8192)
    with not_found_as("folder_not_found", _FOLDER_NOT_FOUND):
        if cursor:
            link = decode_cursor(cursor, base=client.base, query_fingerprint=key)
            page = client.next_page(link, purpose="Outlook message search")
        else:
            page = client.get(
                f"/mailFolders/{segment(normalized['folder'])}/messages",
                query=_graph_query(normalized),
                purpose="Outlook message search",
            )
    rows = page.get("value")
    if not isinstance(rows, list):
        raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid message list")
    messages = [summary(row) for row in rows[: normalized["limit"]] if isinstance(row, Mapping)]
    next_link = page.get("@odata.nextLink")
    next_cursor = (
        encode_cursor(next_link, base=client.base, query_fingerprint=key)
        if isinstance(next_link, str) and next_link
        else ""
    )
    return {"messages": messages, "next_cursor": next_cursor, "count": len(messages)}


def summary(row: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "message_id": str(row.get("id") or ""),
        "internet_message_id": str(row.get("internetMessageId") or ""),
        "conversation_id": str(row.get("conversationId") or ""),
        "subject": str(row.get("subject") or ""),
        "from": address_out(row.get("from")),
        "to": addresses_out(row.get("toRecipients")),
        "cc": addresses_out(row.get("ccRecipients")),
        "received_at": row.get("receivedDateTime"),
        "sent_at": row.get("sentDateTime"),
        "is_read": bool(row.get("isRead")),
        "is_draft": bool(row.get("isDraft")),
        "has_attachments": bool(row.get("hasAttachments")),
        "importance": str(row.get("importance") or "normal"),
        "flag_status": flag_status_out(row.get("flag")),
        "folder_id": str(row.get("parentFolderId") or ""),
        "preview": str(row.get("bodyPreview") or "")[:512],
    }


# -- get_message ----------------------------------------------------------------------


def get_message(ctx: OperationContext) -> dict[str, Any]:
    client = mailbox(ctx)
    message_id = graph_id(ctx.input.get("message_id"), "message_id")
    include_body = bool_input(ctx.input.get("include_body"), "include_body", default=True)
    include_headers = bool_input(ctx.input.get("include_headers"), "include_headers", default=True)
    include_attachments = bool_input(
        ctx.input.get("include_attachment_metadata"), "include_attachment_metadata", default=True
    )
    include_html = bool_input(ctx.input.get("include_raw_html"), "include_raw_html", default=False)
    select = _SUMMARY_SELECT + ",sender,bccRecipients,replyTo"
    if include_body:
        select += ",body"
    if include_headers:
        select += ",internetMessageHeaders"
    path = f"/messages/{segment(message_id)}"
    with not_found_as("message_not_found", _MESSAGE_NOT_FOUND):
        row = client.get(
            path,
            query={"$select": select},
            prefer=(BODY_PREFER_TEXT,),
            response_limit=MESSAGE_RESPONSE_LIMIT,
            purpose="Outlook message",
        )
        message = summary(row)
        message.pop("preview", None)
        message["sender"] = address_out(row.get("sender"))
        message["bcc"] = addresses_out(row.get("bccRecipients"))
        message["reply_to"] = addresses_out(row.get("replyTo"))
        if include_body:
            body = mapping(row.get("body"))
            message["body_text"], message["body_text_truncated"] = _bounded(
                str(body.get("content") or "")
            )
        if include_html:
            html_row = client.get(
                path,
                query={"$select": "body"},
                prefer=(BODY_PREFER_HTML,),
                response_limit=MESSAGE_RESPONSE_LIMIT,
                purpose="Outlook message HTML",
            )
            html_body = mapping(html_row.get("body"))
            message["body_html"], message["body_html_truncated"] = _bounded(
                str(html_body.get("content") or "")
            )
        if include_headers:
            message["headers"] = _headers(row.get("internetMessageHeaders"))
        if include_attachments:
            message["attachments"] = list_attachment_metadata(client, message_id)
    return {"message": message}


def _bounded(value: str) -> tuple[str, bool]:
    encoded = value.encode("utf-8")
    if len(encoded) <= MAX_BODY_BYTES:
        return value, False
    return encoded[:MAX_BODY_BYTES].decode("utf-8", errors="ignore"), True


def _headers(raw: Any) -> list[dict[str, str]]:
    if not isinstance(raw, list):
        return []
    rows = []
    for item in raw[:MAX_HEADERS]:
        if isinstance(item, Mapping):
            rows.append(
                {
                    "name": str(item.get("name") or "")[:256],
                    "value": str(item.get("value") or "")[:MAX_HEADER_VALUE],
                }
            )
    return rows


_ATTACHMENT_TYPES = {
    "#microsoft.graph.fileAttachment": "file",
    "#microsoft.graph.itemAttachment": "item",
    "#microsoft.graph.referenceAttachment": "reference",
}


def attachment_type(row: Mapping[str, Any]) -> str:
    return _ATTACHMENT_TYPES.get(str(row.get("@odata.type") or ""), "unknown")


def list_attachment_metadata(client: Mailbox, message_id: str) -> list[dict[str, Any]]:
    rows = client.list_all(
        f"/messages/{segment(message_id)}/attachments",
        query={"$select": "id,name,contentType,size,isInline"},
        max_items=MAX_ATTACHMENT_ROWS,
        purpose="Outlook attachments",
    )
    return [
        {
            "attachment_id": str(row.get("id") or ""),
            "name": str(row.get("name") or ""),
            "content_type": str(row.get("contentType") or ""),
            "size": _count(row.get("size")),
            "is_inline": bool(row.get("isInline")),
            "attachment_type": attachment_type(row),
        }
        for row in rows
    ]


__all__ = [
    "SEARCH_FIELDS",
    "attachment_type",
    "folder_items",
    "get_message",
    "list_attachment_metadata",
    "list_mail_folders",
    "search_messages",
    "summary",
    "connection_test",
]
