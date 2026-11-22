"""Access tokens for both connection modes.

Delegated connections receive a Core-managed, already refreshed ``access_token``.
App-only connections run the OAuth client-credentials grant here, because Core's
OAuth engine only supports ``authorization_code``. The token stays in memory for
this invocation; it is never persisted, logged or decoded.
"""

from __future__ import annotations

import json
from collections.abc import Callable
from typing import Any
from urllib.error import HTTPError
from urllib.parse import quote, urlencode
from urllib.request import Request

from flowsteward_extension_sdk import PinnedPeerError, open_pinned_url

from .connection import ConnectionContext, valid_tenant
from .errors import M365Error
from .transport import (
    ERROR_RESPONSE_LIMIT,
    LOGIN_HOST,
    GraphTransport,
    map_http_error,
    read_bounded,
)

GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default"
TOKEN_RESPONSE_LIMIT = 64 * 1024

Opener = Callable[..., Any]


def app_token(context: ConnectionContext, *, opener: Opener = open_pinned_url) -> str:
    if not context.is_app:
        raise ValueError("app_token is only for app-only connections")
    tenant = context.tenant_id
    if not valid_tenant(tenant):
        raise M365Error("invalid_connection", "The app-only connection needs a valid tenant ID")
    url = f"https://{LOGIN_HOST}/{quote(tenant, safe='.-')}/oauth2/v2.0/token"
    body = urlencode(
        {
            "grant_type": "client_credentials",
            "client_id": context.client_id,
            "client_secret": context.client_secret,
            "scope": GRAPH_DEFAULT_SCOPE,
        }
    ).encode("ascii")
    request = Request(
        url,
        data=body,
        method="POST",
        headers={"Accept": "application/json", "Content-Type": "application/x-www-form-urlencoded"},
    )
    try:
        with opener(request, timeout_seconds=20.0, purpose="Microsoft Entra token") as response:
            raw = read_bounded(response, TOKEN_RESPONSE_LIMIT)
    except HTTPError as exc:
        try:
            error_body = read_bounded(exc, ERROR_RESPONSE_LIMIT)
        except Exception:
            error_body = b""
        mapped = map_http_error(int(exc.code), error_body, method="GET", retry_after=None)
        if mapped.code == "authorization_required" or (
            mapped.code == "invalid_payload" and int(exc.code) in {400, 401}
        ):
            # A token-endpoint 400/401 without a recognised AADSTS code is a client problem.
            raise M365Error(
                "invalid_client",
                "Microsoft Entra rejected the app credentials",
                http_status=int(exc.code),
            ) from exc
        raise mapped from exc
    except (PinnedPeerError, TimeoutError, OSError) as exc:
        raise M365Error("network_failure", "Microsoft Entra could not be reached") from exc
    try:
        payload = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise M365Error(
            "provider_unavailable", "Microsoft Entra returned an invalid token response"
        ) from exc
    token = payload.get("access_token") if isinstance(payload, dict) else None
    if not isinstance(token, str) or not token.strip():
        raise M365Error("provider_unavailable", "Microsoft Entra returned no access token")
    return token.strip()


def transport_for(
    context: ConnectionContext,
    *,
    opener: Opener = open_pinned_url,
    deadline_seconds: float = 240.0,
) -> GraphTransport:
    token = app_token(context, opener=opener) if context.is_app else context.access_token
    return GraphTransport(token, opener=opener, deadline_seconds=deadline_seconds)


__all__ = ["GRAPH_DEFAULT_SCOPE", "app_token", "transport_for"]
