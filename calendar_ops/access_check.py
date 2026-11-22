"""``check_calendar_access`` (task D2): is this app-only calendar connection scoped?

Two reads with the app token: the default calendar of the connection's mailbox (it must
succeed) and the default calendar of a control mailbox outside the Exchange RBAC scope
(it must be refused). The check never writes. Exchange's ``Application
Calendars.ReadWrite`` role grants read and write together, so a refused read on the
control mailbox is the expected outcome of a correct scope; a successful one means the
app also holds an Entra calendar application permission.
"""

from __future__ import annotations

from typing import Any
from urllib.parse import quote

from m365.errors import M365Error
from m365.runtime import OperationContext
from mail.common import Mailbox, email_address

DENIED_CODES = frozenset({"permission_denied", "mailbox_access_denied", "authorization_required"})
WRITE_NOTE = (
    "This check only reads. Exchange RBAC grants calendar read and write together "
    "(Application Calendars.ReadWrite), so a refused read on the control mailbox shows "
    "the scope works; it does not create or change any event."
)
BROAD_MESSAGE = (
    "This app can read calendars outside its scope; remove the Entra Calendars.Read / "
    "Calendars.ReadWrite application permission and use Exchange RBAC only."
)


def check_calendar_access(ctx: OperationContext) -> dict[str, Any]:
    if not ctx.connection.is_app:
        raise M365Error(
            "invalid_connection", "The access check is only for app-only calendar connections"
        )
    raw = ctx.input.get("control_mailbox") or ctx.connection.config.get("control_mailbox")
    if not raw:
        raise M365Error("invalid_payload", "Enter a control mailbox that this app must not reach")
    control = email_address(raw, "control_mailbox")
    own = ctx.connection.mailbox
    if control.lower() == own.lower():
        raise M365Error(
            "invalid_payload", "The control mailbox must differ from the connection's mailbox"
        )
    own_read, own_code = _read_probe(ctx.transport, own, "Calendar access check (own mailbox)")
    control_read, control_code = _read_probe(
        ctx.transport, control, "Calendar access check (control mailbox)"
    )
    broad = control_read == "allowed"
    if broad:
        verdict, message = "broad_access", BROAD_MESSAGE
    elif own_read == "allowed" and control_read == "denied":
        verdict, message = (
            "restricted",
            "The app reads its own calendar and is refused on the control mailbox.",
        )
    elif own_read == "denied":
        verdict, message = (
            "inconclusive",
            "The app cannot read the connection's own calendar; check the Exchange RBAC "
            "management scope.",
        )
    else:
        verdict, message = (
            "inconclusive",
            "The check could not decide; see the read results.",
        )
    return {
        "control_mailbox": control,
        "verdict": verdict,
        "broad_access": broad,
        "own_read": own_read,
        "own_read_error_code": own_code,
        "control_read": control_read,
        "control_read_error_code": control_code,
        "write_tested": False,
        "message": message,
        "note": WRITE_NOTE,
    }


def _read_probe(transport: Any, mailbox: str, purpose: str) -> tuple[str, str]:
    client = Mailbox(transport, f"/users/{quote(mailbox, safe='@.')}")
    try:
        client.get("/calendar", query={"$select": "id"}, response_limit=64 * 1024, purpose=purpose)
    except M365Error as exc:
        return ("denied" if exc.code in DENIED_CODES else "inconclusive"), exc.code
    return "allowed", ""


__all__ = ["BROAD_MESSAGE", "WRITE_NOTE", "check_calendar_access"]
