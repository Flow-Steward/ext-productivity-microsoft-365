"""Resolve the selected connection from a host payload.

Five connection types exist (see README): ``m365_mail_user``, ``m365_mail_app``,
``m365_files_user``, ``m365_files_app`` and ``m365_excel_user``. Delegated
(``_user``) types carry a Core-managed OAuth ``access_token``; app-only (``_app``)
types carry their own ``client_secret`` plus ``tenant_id``/``client_id`` config.
"""

from __future__ import annotations

import re
from collections.abc import Mapping
from dataclasses import dataclass, field
from typing import Any

from .errors import M365Error

CONNECTION_TYPES: dict[str, tuple[str, str]] = {
    "m365_mail_user": ("mail", "user"),
    "m365_mail_app": ("mail", "app"),
    "m365_files_user": ("files", "user"),
    "m365_files_app": ("files", "app"),
    "m365_excel_user": ("excel", "user"),
}

_GUID = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")
_TENANT_DOMAIN = re.compile(r"^(?=.{3,253}$)[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$")
_MAILBOX = re.compile(r"^[^@\s/\\?#%]{1,64}@[A-Za-z0-9.-]{3,253}$")


def mapping(value: Any) -> dict[str, Any]:
    return dict(value) if isinstance(value, Mapping) else {}


def text(value: Any, *, limit: int = 2048) -> str:
    if not isinstance(value, str):
        return ""
    stripped = value.strip()
    if len(stripped) > limit or any(ord(char) < 32 for char in stripped):
        return ""
    return stripped


@dataclass(frozen=True)
class ConnectionContext:
    connection_id: str
    connection_type_id: str
    service: str
    mode: str
    config: dict[str, Any] = field(default_factory=dict)
    access_token: str = field(default="", repr=False)
    client_secret: str = field(default="", repr=False)

    @property
    def is_app(self) -> bool:
        return self.mode == "app"

    @property
    def tenant_id(self) -> str:
        return text(self.config.get("tenant_id"), limit=256)

    @property
    def client_id(self) -> str:
        return text(self.config.get("client_id"), limit=256)

    @property
    def mailbox(self) -> str:
        return text(self.config.get("mailbox"), limit=320)


def connection_payload(payload: Mapping[str, Any]) -> dict[str, Any]:
    for candidate in (
        mapping(mapping(payload.get("action")).get("target")).get("connection"),
        mapping(payload.get("target")).get("connection"),
    ):
        if isinstance(candidate, Mapping) and candidate:
            return dict(candidate)
    raise M365Error("invalid_connection", "A Microsoft 365 connection must be selected")


def resolve_connection(
    payload: Mapping[str, Any],
    *,
    service: str,
    connection_ref: str = "",
) -> ConnectionContext:
    """Return the selected connection, checked against the requested service."""
    connection = connection_payload(payload)
    connection_id = text(connection.get("connection_id") or connection.get("id"), limit=256)
    if connection_ref and connection_id and connection_id != connection_ref:
        raise M365Error(
            "invalid_connection", "The selected connection does not match connection_ref"
        )
    type_id = text(
        connection.get("connection_type_id") or connection.get("connection_type"), limit=128
    )
    kind = CONNECTION_TYPES.get(type_id)
    if kind is None or kind[0] != service:
        raise M365Error(
            "invalid_connection", f"This operation needs a Microsoft 365 {service} connection"
        )
    config = mapping(connection.get("config") or connection.get("connection_config"))
    secrets = mapping(connection.get("secrets"))
    mode = kind[1]
    context = ConnectionContext(
        connection_id=connection_id,
        connection_type_id=type_id,
        service=service,
        mode=mode,
        config=config,
        access_token=text(secrets.get("access_token"), limit=16384) if mode == "user" else "",
        client_secret=text(secrets.get("client_secret"), limit=4096) if mode == "app" else "",
    )
    _validate(context)
    return context


def _validate(context: ConnectionContext) -> None:
    if context.mode == "user":
        if not context.access_token:
            raise M365Error("authorization_required", "Microsoft 365 authorization is required")
    else:
        if not valid_tenant(context.tenant_id) or not _GUID.match(context.client_id):
            raise M365Error(
                "invalid_connection", "The app-only connection needs a tenant ID and client ID"
            )
        if not context.client_secret:
            raise M365Error("invalid_client", "The app-only connection has no client secret")
    if context.service == "mail":
        mailbox = context.mailbox
        if context.mode == "app" and not mailbox:
            raise M365Error("invalid_connection", "The app-only mail connection needs a mailbox")
        if mailbox and not _MAILBOX.match(mailbox):
            raise M365Error("invalid_connection", "The mailbox address is invalid")


def valid_tenant(value: str) -> bool:
    return bool(value) and bool(_GUID.match(value) or _TENANT_DOMAIN.match(value))


def mailbox_base(context: ConnectionContext) -> str:
    """Graph path prefix of the connection's mailbox: ``/me`` or ``/users/{mailbox}``."""
    from urllib.parse import quote

    if context.mailbox:
        return f"/users/{quote(context.mailbox, safe='@.')}"
    if context.is_app:
        raise M365Error("invalid_connection", "The app-only mail connection needs a mailbox")
    return "/me"


__all__ = [
    "CONNECTION_TYPES",
    "ConnectionContext",
    "connection_payload",
    "mailbox_base",
    "mapping",
    "resolve_connection",
    "text",
    "valid_tenant",
]
