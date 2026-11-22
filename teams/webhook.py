"""Teams Workflows webhook client for ``m365_teams_webhook`` connections (task D4).

The connection's secret ``webhook_url`` points at a Power Automate flow ("When a Teams
webhook request is received"). The client POSTs an Adaptive Card payload to it, without
Graph auth, under these rules:

* the URL must be ``https`` on a host that ends with one of
  :data:`ALLOWED_WEBHOOK_HOST_SUFFIXES` (Core does not enforce wildcard egress yet);
* redirects are never followed: a 3xx is ``webhook_rejected``;
* 2xx means the flow **accepted** the request (``confirmation: accepted_by_webhook``,
  no message id): a later flow step posts the card;
* a lost response, timeout, reset or 5xx is ambiguous: ``timeout_unknown``, never
  repeated (the webhook has no idempotency, a repeat could post the card twice);
* 4xx is ``webhook_rejected`` with ``definitely_no_external_effect``;
* the URL is a secret: no message, log or result contains it or any part of it, and
  exceptions are raised without the underlying cause (which would hold the URL).

The default opener is the SDK's pinned opener (DNS pinning, private addresses refused,
redirects disabled). Tests inject their own.
"""

from __future__ import annotations

import http.client
import json
import socket
from collections.abc import Callable, Mapping
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlsplit
from urllib.request import Request

from flowsteward_extension_sdk import PinnedPeerError, open_pinned_url, parse_retry_after

from m365.connection import ConnectionContext
from m365.errors import M365Error

#: Hosts of Teams Workflows / Power Automate webhook triggers.
ALLOWED_WEBHOOK_HOST_SUFFIXES: tuple[str, ...] = (
    ".environment.api.powerplatform.com",
    ".logic.azure.com",
)
#: Teams renders cards up to about 28 KB.
MAX_WEBHOOK_PAYLOAD_BYTES = 28 * 1024
RESPONSE_DRAIN_LIMIT = 16 * 1024
WEBHOOK_TIMEOUT_SECONDS = 20.0
_PURPOSE = "Teams webhook"

Opener = Callable[..., Any]


def check_webhook_url(value: Any) -> str:
    """The URL if it is https on an allowed host; otherwise an error that never echoes it."""
    if not isinstance(value, str) or not value.strip() or len(value) > 4096:
        raise M365Error("invalid_connection", "The webhook connection has no valid webhook URL")
    url = value.strip()
    if any(char.isspace() or ord(char) < 32 or ord(char) == 127 for char in url):
        raise M365Error("invalid_connection", "The webhook URL is malformed")
    try:
        parsed = urlsplit(url)
        port = parsed.port
    except ValueError:
        raise M365Error("invalid_connection", "The webhook URL is malformed") from None
    if parsed.scheme.lower() != "https":
        raise M365Error("invalid_connection", "The webhook URL must use https")
    host = (parsed.hostname or "").lower()
    if parsed.username is not None or parsed.password is not None or port not in (None, 443):
        raise M365Error("invalid_connection", "The webhook URL is malformed")
    if not is_allowed_webhook_host(host):
        raise M365Error(
            "egress_host_not_allowed",
            "The webhook URL is not a Teams Workflows (Power Automate) webhook address",
        )
    return url


def is_allowed_webhook_host(host: str) -> bool:
    host = (host or "").lower()
    if not host or any(not (char.isalnum() or char in ".-") for char in host):
        return False
    return any(
        host.endswith(suffix) and len(host) > len(suffix) and not host.startswith(".")
        for suffix in ALLOWED_WEBHOOK_HOST_SUFFIXES
    )


def post_webhook_card(
    webhook_url: str,
    payload: Mapping[str, Any],
    *,
    opener: Opener = open_pinned_url,
    timeout_seconds: float = WEBHOOK_TIMEOUT_SECONDS,
) -> dict[str, Any]:
    """POST ``payload`` (see :func:`teams.cards.webhook_payload`) once to the webhook.

    Returns the accepted result; raises :class:`M365Error` with ``webhook_rejected``
    (nothing happened), ``timeout_unknown`` (ambiguous, never retried here),
    ``network_failure`` (the request never left), ``message_too_large`` or a URL error.
    """
    url = check_webhook_url(webhook_url)
    body = json.dumps(dict(payload), separators=(",", ":")).encode("utf-8")
    if len(body) > MAX_WEBHOOK_PAYLOAD_BYTES:
        raise M365Error(
            "message_too_large",
            f"The card is {len(body)} bytes; Teams accepts about {MAX_WEBHOOK_PAYLOAD_BYTES} bytes",
        )
    request = Request(
        url,
        data=body,
        method="POST",
        headers={"Content-Type": "application/json", "Accept": "application/json"},
    )
    status, retry_after = _send(request, opener=opener, timeout_seconds=timeout_seconds)
    return _classify(status, retry_after)


def post_webhook_for_connection(
    connection: ConnectionContext,
    payload: Mapping[str, Any],
    *,
    opener: Opener = open_pinned_url,
    timeout_seconds: float = WEBHOOK_TIMEOUT_SECONDS,
) -> dict[str, Any]:
    """:func:`post_webhook_card` with the URL of an ``m365_teams_webhook`` connection."""
    if connection.mode != "webhook" or connection.service != "teams":
        raise M365Error(
            "notification_connection_incompatible", "This is not a Teams webhook connection"
        )
    return post_webhook_card(
        connection.webhook_url, payload, opener=opener, timeout_seconds=timeout_seconds
    )


def _send(request: Request, *, opener: Opener, timeout_seconds: float) -> tuple[int, float | None]:
    """Return (status, Retry-After); every exception is replaced, never chained."""
    error: M365Error
    try:
        with opener(request, timeout_seconds=timeout_seconds, purpose=_PURPOSE) as response:
            status = int(getattr(response, "status", None) or response.getcode() or 0)
            retry_after = _retry_after(getattr(response, "headers", None))
            if 200 <= status < 300:
                try:
                    response.read(RESPONSE_DRAIN_LIMIT)
                except Exception:
                    pass  # the status already says the flow accepted the request
            return status, retry_after
    except HTTPError as exc:
        status, retry_after = int(exc.code), _retry_after(exc.headers)
        try:
            exc.close()
        except Exception:
            pass
        return status, retry_after
    except PinnedPeerError:
        error = _not_sent()
    except URLError as exc:
        reason = exc.reason
        if isinstance(reason, PinnedPeerError | ConnectionRefusedError | socket.gaierror):
            error = _not_sent()
        else:
            error = _ambiguous()
    except (TimeoutError, OSError, http.client.HTTPException):
        error = _ambiguous()
    except Exception:
        # Unknown failure: the request may have left. Its text may hold the URL.
        error = _ambiguous()
    raise error


def _classify(status: int, retry_after: float | None) -> dict[str, Any]:
    if 200 <= status < 300:
        return {
            "delivered": True,
            "confirmation": "accepted_by_webhook",
            "provider_message_id": "",
            "http_status": status,
            "external_effect_status": "succeeded",
            "definitely_no_external_effect": False,
        }
    if 300 <= status < 400:
        raise M365Error(
            "webhook_rejected",
            "The webhook answered with a redirect, which is not followed",
            http_status=status,
        )
    if status == 429:
        raise M365Error(
            "webhook_rejected",
            "The webhook is throttling requests; nothing was posted",
            http_status=status,
            retry_after_seconds=retry_after,
            retryable=True,
            failure_class="provider",
        )
    if 400 <= status < 500:
        raise M365Error(
            "webhook_rejected",
            "The webhook refused the request; check that the flow is on and its trigger "
            "allows anyone",
            http_status=status,
        )
    raise M365Error(
        "timeout_unknown",
        "The webhook did not confirm the request; check the channel before sending again",
        http_status=status,
        ambiguous=True,
    )


def _not_sent() -> M365Error:
    return M365Error("network_failure", "The webhook could not be reached; nothing was sent")


def _ambiguous() -> M365Error:
    return M365Error(
        "timeout_unknown",
        "The webhook response was lost; check the channel before sending again",
        ambiguous=True,
    )


def _retry_after(headers: Any) -> float | None:
    try:
        value = headers.get("Retry-After") if headers is not None else None
    except Exception:
        return None
    seconds = parse_retry_after(value) if value else None
    return float(seconds) if seconds is not None else None


__all__ = [
    "ALLOWED_WEBHOOK_HOST_SUFFIXES",
    "MAX_WEBHOOK_PAYLOAD_BYTES",
    "check_webhook_url",
    "is_allowed_webhook_host",
    "post_webhook_card",
    "post_webhook_for_connection",
]
