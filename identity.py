"""Post-connect identity verification for delegated connections.

Core passes only the fresh access token to this action. The action calls
``GET /me`` and never decodes the token (Microsoft treats Graph tokens as opaque).
The stable identity is ``/me.id``.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from flowsteward_extension_sdk import open_pinned_url

from m365.connection import mapping, text
from m365.errors import M365Error, error_response
from m365.transport import GraphTransport

ME_RESPONSE_LIMIT = 64 * 1024


def _access_token(payload: dict[str, Any]) -> str:
    action = mapping(payload.get("action"))
    connection = mapping(mapping(action.get("target")).get("connection"))
    secrets = mapping(connection.get("secrets"))
    if set(secrets) != {"access_token"}:
        return ""
    return text(secrets.get("access_token"), limit=16384)


def handle_verify_identity(
    payload: dict[str, Any], *, opener: Callable[..., Any] = open_pinned_url
) -> dict[str, Any]:
    token = _access_token(payload)
    if not token:
        return error_response("invalid_payload", "A single connection access token is required")
    try:
        me = GraphTransport(token, opener=opener, deadline_seconds=30.0).get_json(
            "/me",
            query={"$select": "id,userPrincipalName,mail,displayName"},
            response_limit=ME_RESPONSE_LIMIT,
            purpose="Microsoft Graph /me",
        )
    except M365Error as exc:
        return error_response(exc.code, exc.message, **exc.retry_facts())
    subject = text(me.get("id"), limit=256)
    upn = text(me.get("userPrincipalName"), limit=320)
    mail = text(me.get("mail"), limit=320)
    if not subject or not (upn or mail):
        return error_response("provider_unavailable", "Microsoft 365 returned no usable identity")
    display = mail or upn
    return {
        "ok": True,
        "result": {
            "subject": subject,
            "verified_display_name": text(me.get("displayName"), limit=256) or display,
            "verified_email": display,
            "safe_metadata": {"user_principal_name": upn},
            "health": "ok",
        },
    }


__all__ = ["handle_verify_identity"]
