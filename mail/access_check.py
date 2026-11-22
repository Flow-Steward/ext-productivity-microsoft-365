"""``check_mail_access`` (task B4.3): can this app-only connection reach a control mailbox?

Two real Graph calls with the app token against a mailbox that must stay out of reach:
a read of its Inbox, and a draft create (deleted again at once). Either one succeeding
means the app has broader mail access than its Exchange RBAC scope. The check covers
``Mail.ReadWrite`` only: it never sends mail, so it proves nothing about ``Mail.Send``.
"""

from __future__ import annotations

import uuid
from typing import Any
from urllib.parse import quote

from m365.errors import M365Error
from m365.runtime import OperationContext

from .common import Mailbox, email_address, odata_string, provider_id, segment
from .send import KEY_PROPERTY_ID

DENIED_CODES = frozenset({"permission_denied", "mailbox_access_denied", "authorization_required"})
SEND_NOTE = (
    "This check covers mail read and write (Mail.ReadWrite) only. It does not test send "
    "permission: an app with a broad Entra Mail.Send grant still passes it and can send as "
    "other mailboxes. Verify send restrictions separately."
)
BROAD_MESSAGE = (
    "This app can read or change mailboxes outside its scope; remove the Entra "
    "Mail.ReadWrite application permission."
)


def check_mail_access(ctx: OperationContext) -> dict[str, Any]:
    if not ctx.connection.is_app:
        raise M365Error(
            "invalid_connection", "The access check is only for app-only mail connections"
        )
    raw = ctx.input.get("control_mailbox") or ctx.connection.config.get("control_mailbox")
    if not raw:
        raise M365Error("invalid_payload", "Enter a control mailbox that this app must not reach")
    control = email_address(raw, "control_mailbox")
    if control.lower() == ctx.connection.mailbox.lower():
        raise M365Error(
            "invalid_payload", "The control mailbox must differ from the connection's mailbox"
        )
    client = Mailbox(ctx.transport, f"/users/{quote(control, safe='@.')}")
    read, read_code = _read_probe(client)
    write, write_code, cleanup_failed = _write_probe(client)
    broad = "allowed" in {read, write}
    if broad:
        verdict, message = "broad_access", BROAD_MESSAGE
    elif read == "denied" and write == "denied":
        verdict, message = (
            "restricted",
            "The app could neither read nor change the control mailbox.",
        )
    else:
        verdict, message = (
            "inconclusive",
            "The check could not decide; see the read and write results.",
        )
    return {
        "control_mailbox": control,
        "verdict": verdict,
        "broad_access": broad,
        "read": read,
        "read_error_code": read_code,
        "write": write,
        "write_error_code": write_code,
        "cleanup_failed": cleanup_failed,
        "send_permission_tested": False,
        "message": message,
        "note": SEND_NOTE,
    }


def _read_probe(client: Mailbox) -> tuple[str, str]:
    try:
        client.get(
            "/mailFolders/inbox",
            query={"$select": "id"},
            response_limit=64 * 1024,
            purpose="Access check read",
        )
    except M365Error as exc:
        return ("denied" if exc.code in DENIED_CODES else "inconclusive"), exc.code
    return "allowed", ""


def _write_probe(client: Mailbox) -> tuple[str, str, bool]:
    marker = f"access-check:{uuid.uuid4()}"
    draft = {
        "subject": "Flow Steward access check (safe to delete)",
        "body": {
            "contentType": "Text",
            "content": "Created and deleted by a Flow Steward access check.",
        },
        "singleValueExtendedProperties": [{"id": KEY_PROPERTY_ID, "value": marker}],
    }
    try:
        response = client.write(
            "POST", "/messages", json_body=draft, retry="none", purpose="Access check write"
        )
    except M365Error as exc:
        if exc.ambiguous:
            return _resolve_ambiguous(client, marker)
        return ("denied" if exc.code in DENIED_CODES else "inconclusive"), exc.code, False
    try:
        draft_id = provider_id(response.json().get("id"))
    except M365Error:
        return "allowed", "", True
    return "allowed", "", not _delete(client, draft_id)


def _resolve_ambiguous(client: Mailbox, marker: str) -> tuple[str, str, bool]:
    try:
        page = client.get(
            "/mailFolders/drafts/messages",
            query={
                "$filter": (
                    f"singleValueExtendedProperties/Any(ep: ep/id eq {odata_string(KEY_PROPERTY_ID)} "
                    f"and ep/value eq {odata_string(marker)})"
                ),
                "$select": "id",
            },
            purpose="Access check lookup",
        )
    except M365Error as exc:
        return ("denied" if exc.code in DENIED_CODES else "inconclusive"), "timeout_unknown", False
    rows = [row for row in page.get("value") or [] if isinstance(row, dict)]
    if not rows:
        return "inconclusive", "timeout_unknown", False
    cleanup_failed = False
    for row in rows:
        try:
            cleanup_failed = not _delete(client, provider_id(row.get("id"))) or cleanup_failed
        except M365Error:
            cleanup_failed = True
    return "allowed", "", cleanup_failed


def _delete(client: Mailbox, message_id: str) -> bool:
    try:
        client.write(
            "DELETE",
            f"/messages/{segment(message_id)}",
            retry="none",
            purpose="Access check cleanup",
        )
    except M365Error as exc:
        return exc.code == "item_not_found"
    return True


__all__ = ["BROAD_MESSAGE", "SEND_NOTE", "check_mail_access"]
