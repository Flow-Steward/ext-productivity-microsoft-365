"""Core's ``deliver_notification`` action, routed by the connection type.

The request is Core's local delivery request: ``event``, ``route`` (with
``route_config``), ``action_payload.target.connection`` and the top-level
``idempotency_key`` (task C3). Nothing here logs the subject or body: system email
events are ``payload_sensitive`` and their body carries one-time links.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping
from typing import Any

from m365.auth import transport_for
from m365.connection import connection_payload, mapping, resolve_connection, text
from m365.errors import M365Error, error_from_exception
from m365.runtime import TransportFactory, test_mode
from mail.common import Mailbox
from mail.send import SendFlow, plan_send

from .content import email_subject_and_body, event_content

ACTION_ID = "deliver_notification"

#: Connection type → how a route addresses it (the C1 ``notification`` block). A signed-in
#: account delivers email; Teams channels are reached through the Workflows webhook.
RECIPIENT = "recipient"
CONNECTION_CHANNEL = "connection_channel"
NOTIFICATION_TYPES: dict[str, str] = {
    "m365_user": RECIPIENT,
    "m365_mail_app": RECIPIENT,
    "m365_teams_webhook": CONNECTION_CHANNEL,
}

ChannelPoster = Callable[..., dict[str, Any]]

_SPLIT = re.compile(r"[,;\s]+")


def is_delivery_request(payload: Mapping[str, Any]) -> bool:
    action = payload.get("action")
    if isinstance(action, str):
        return action.strip() == ACTION_ID
    action_map = mapping(action)
    return (
        text(action_map.get("action_id") or action_map.get("operation_id"), limit=128) == ACTION_ID
    )


def _result(*, delivered: bool, provider_message_id: str = "", **extra: Any) -> dict[str, Any]:
    response: dict[str, Any] = {
        "ok": True,
        "delivered": delivered,
        "provider_message_id": provider_message_id,
        "result": {"delivered": delivered, "provider_message_id": provider_message_id, **extra},
        "errors": [],
    }
    response.update(extra)
    return response


def _failure(exc: M365Error) -> dict[str, Any]:
    response = error_from_exception(exc, effectful=True)
    response["delivered"] = False
    response["provider_message_id"] = ""
    return response


def _incompatible(message: str) -> M365Error:
    return M365Error("notification_connection_incompatible", message, retryable=False)


def _recipients(value: Any) -> list[Any]:
    if value is None:
        return []
    if isinstance(value, str):
        return [part for part in _SPLIT.split(value) if part]
    if isinstance(value, list):
        return value
    raise M365Error("invalid_payload", "Notification recipients must be a list or text")


def handle_deliver_notification(
    payload: dict[str, Any],
    *,
    transport_factory: TransportFactory = transport_for,
    webhook_poster: ChannelPoster | None = None,
) -> dict[str, Any]:
    try:
        if test_mode(payload):
            return _result(
                delivered=False,
                external_effect_status="suppressed",
                definitely_no_external_effect=True,
            )
        key = text(payload.get("idempotency_key"), limit=256)
        if not key:
            raise M365Error(
                "idempotency_key_missing",
                "Core did not send a delivery idempotency key; nothing was sent",
                retryable=False,
            )
        connection = connection_payload(payload)
        type_id = text(
            connection.get("connection_type_id") or connection.get("connection_type"), limit=128
        )
        address = NOTIFICATION_TYPES.get(type_id)
        if address is None:
            raise _incompatible("This connection type cannot deliver notifications")
        event = mapping(payload.get("event"))
        route_config = mapping(mapping(payload.get("route")).get("route_config"))
        step_key = f"{ACTION_ID}:{key}"
        if address == RECIPIENT:
            return _deliver_mail(payload, event, route_config, step_key, transport_factory)
        if route_config.get("to") or route_config.get("recipients"):
            raise _incompatible("A channel connection does not take recipients")
        return _deliver_webhook(payload, event, route_config, webhook_poster)
    except M365Error as exc:
        return _failure(exc)
    except Exception:
        return _failure(M365Error("internal_error", "The notification delivery failed"))


def _deliver_mail(
    payload: dict[str, Any],
    event: Mapping[str, Any],
    route_config: Mapping[str, Any],
    step_key: str,
    transport_factory: TransportFactory,
) -> dict[str, Any]:
    if route_config.get("channel"):
        raise _incompatible("A mail connection takes recipients, not a channel")
    to = _recipients(route_config.get("to") or route_config.get("recipients"))
    if not to:
        raise M365Error("recipient_required", "The route has no recipient", retryable=False)
    subject, body = email_subject_and_body(event, route_config)
    plan = plan_send(
        {
            "to": to,
            "cc": _recipients(route_config.get("cc")),
            "bcc": _recipients(route_config.get("bcc")),
            "subject": subject,
            "body_text": body,
        },
        step_key,
    )
    context = resolve_connection(payload, service="mail")
    flow = SendFlow(payload, Mailbox.for_connection(transport_factory(context), context), step_key)
    sent = flow.execute(plan, [])
    return _result(
        delivered=True,
        provider_message_id=str(sent.get("message_id") or ""),
        confirmation=str(sent.get("confirmation") or "sent"),
        external_effect_status="succeeded",
        definitely_no_external_effect=False,
    )


def _deliver_webhook(
    payload: dict[str, Any],
    event: Mapping[str, Any],
    route_config: Mapping[str, Any],
    webhook_poster: ChannelPoster | None,
) -> dict[str, Any]:
    if route_config.get("channel"):
        raise _incompatible("This webhook posts to its own channel; the route must not name one")
    context = resolve_connection(payload, service="teams")
    poster = webhook_poster or _default_webhook_poster()
    poster(context.webhook_url, content=event_content(event))
    return _result(
        delivered=True,
        provider_message_id="",
        confirmation="accepted_by_webhook",
        external_effect_status="succeeded",
        definitely_no_external_effect=False,
    )


def _default_webhook_poster() -> ChannelPoster:
    from teams.notify import post_webhook_notification

    return post_webhook_notification


__all__ = ["ACTION_ID", "NOTIFICATION_TYPES", "handle_deliver_notification", "is_delivery_request"]
