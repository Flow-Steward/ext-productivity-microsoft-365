"""Teams side of ``deliver_notification`` (task D4), called by ``notify.deliver``.

* :func:`post_notification`: ``m365_teams_user`` routes. ``target`` is the route's
  ``channel`` (``team_id/channel_id`` or a chat id); ``key`` is Core's idempotency key.
  The post follows :mod:`teams.posting`: never repeated, ``timeout_unknown`` unless the
  hidden marker is found.
* :func:`post_webhook_notification`: ``m365_teams_webhook`` connections; an Adaptive
  Card through :mod:`teams.webhook`.

``content`` is any object with ``title``, ``severity``, ``summary``, ``run_url`` and
``details`` (``notify.content.NotificationContent``). Nothing here logs it.
"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Any, Protocol

from flowsteward_extension_sdk import open_pinned_url

from .cards import notification_card, notification_html, webhook_payload
from .markers import DEFAULT_MARKER, MarkerStrategy
from .posting import parse_target, post_html
from .webhook import WEBHOOK_TIMEOUT_SECONDS, Opener, post_webhook_card


class NotificationContentLike(Protocol):
    @property
    def title(self) -> str: ...

    @property
    def severity(self) -> str: ...

    @property
    def summary(self) -> str: ...

    @property
    def run_url(self) -> str: ...

    @property
    def details(self) -> Sequence[tuple[str, str]]: ...


def post_notification(
    transport: Any,
    *,
    target: str,
    content: NotificationContentLike,
    key: str,
    marker: MarkerStrategy = DEFAULT_MARKER,
) -> dict[str, Any]:
    """Post the event as an HTML message; returns :func:`teams.posting.post_result`."""
    html_body = notification_html(
        title=content.title,
        severity=content.severity,
        summary=content.summary,
        link=content.run_url,
        facts=tuple(content.details or ()),
    )
    return post_html(transport, parse_target(target), html_body, marker_key=key, marker=marker)


def post_webhook_notification(
    webhook_url: str,
    *,
    content: NotificationContentLike,
    opener: Opener = open_pinned_url,
    timeout_seconds: float = WEBHOOK_TIMEOUT_SECONDS,
) -> dict[str, Any]:
    """POST the event as an Adaptive Card; ``confirmation: accepted_by_webhook`` on 2xx."""
    card = notification_card(
        title=content.title,
        severity=content.severity,
        summary=content.summary,
        link=content.run_url,
        facts=tuple(content.details or ()),
    )
    return post_webhook_card(
        webhook_url, webhook_payload(card), opener=opener, timeout_seconds=timeout_seconds
    )


__all__ = ["NotificationContentLike", "post_notification", "post_webhook_notification"]
