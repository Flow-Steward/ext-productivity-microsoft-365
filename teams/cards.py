"""Notification content for Teams: an Adaptive Card (webhook) and an HTML message (Graph).

Both carry the same facts from Core's event (task D4): title, severity, summary, a link
to the run page and optional detail facts (workflow, step, error).
"""

from __future__ import annotations

from collections.abc import Sequence
from html import escape
from typing import Any
from urllib.parse import urlsplit

MAX_TITLE_CHARS = 256
MAX_SUMMARY_CHARS = 4000
MAX_FACTS = 10
DEFAULT_LINK_LABEL = "Open in Flow Steward"
CARD_CONTENT_TYPE = "application/vnd.microsoft.card.adaptive"
CARD_VERSION = "1.4"

#: Severity → Adaptive Card TextBlock color; unknown severities render as ``default``.
SEVERITY_COLORS = {
    "critical": "attention",
    "error": "attention",
    "failure": "attention",
    "warning": "warning",
    "success": "good",
    "info": "default",
}
_HTML_COLORS = {"attention": "#C4314B", "warning": "#C19C00", "good": "#237B4B"}


def _short(value: Any, limit: int) -> str:
    text = str(value or "").strip()
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


def _severity(value: Any) -> tuple[str, str]:
    label = _short(value, 32).lower() or "info"
    return label, SEVERITY_COLORS.get(label, "default")


def _facts(facts: Sequence[tuple[str, str]]) -> list[tuple[str, str]]:
    rows = []
    for item in list(facts)[:MAX_FACTS]:
        if isinstance(item, tuple | list) and len(item) == 2:
            key, value = _short(item[0], 64), _short(item[1], 1024)
            if key and value:
                rows.append((key, value))
    return rows


def safe_link(value: Any) -> str:
    """An http(s) URL without credentials or whitespace, else ``""`` (the link is dropped)."""
    if not isinstance(value, str) or len(value) > 2048:
        return ""
    url = value.strip()
    try:
        parsed = urlsplit(url)
    except ValueError:
        return ""
    if (
        parsed.scheme not in {"https", "http"}
        or not parsed.hostname
        or parsed.username
        or any(char.isspace() or ord(char) < 32 for char in url)
    ):
        return ""
    return url


def notification_card(
    *,
    title: str,
    severity: str,
    summary: str,
    link: str = "",
    link_label: str = DEFAULT_LINK_LABEL,
    facts: Sequence[tuple[str, str]] = (),
) -> dict[str, Any]:
    """An Adaptive Card (the ``content`` object) with an "Open in Flow Steward" button."""
    label, color = _severity(severity)
    card: dict[str, Any] = {
        "type": "AdaptiveCard",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": CARD_VERSION,
        "body": [
            {
                "type": "TextBlock",
                "text": _short(title, MAX_TITLE_CHARS) or "Flow Steward",
                "weight": "Bolder",
                "size": "Medium",
                "wrap": True,
            },
            {
                "type": "TextBlock",
                "text": label.upper(),
                "color": color.capitalize() if color != "default" else "Default",
                "weight": "Bolder",
                "size": "Small",
                "spacing": "None",
            },
            {"type": "TextBlock", "text": _short(summary, MAX_SUMMARY_CHARS), "wrap": True},
        ],
    }
    rows = _facts(facts)
    if rows:
        card["body"].append(
            {"type": "FactSet", "facts": [{"title": k, "value": v} for k, v in rows]}
        )
    url = safe_link(link)
    if url:
        card["actions"] = [
            {
                "type": "Action.OpenUrl",
                "title": _short(link_label, 64) or DEFAULT_LINK_LABEL,
                "url": url,
            }
        ]
    return card


def webhook_payload(card: dict[str, Any]) -> dict[str, Any]:
    """The envelope a Teams Workflows "webhook request received" flow expects."""
    return {
        "type": "message",
        "attachments": [{"contentType": CARD_CONTENT_TYPE, "contentUrl": None, "content": card}],
    }


def notification_html(
    *,
    title: str,
    severity: str,
    summary: str,
    link: str = "",
    link_label: str = DEFAULT_LINK_LABEL,
    facts: Sequence[tuple[str, str]] = (),
) -> str:
    """The same facts as an HTML body for a Graph channel or chat post."""
    label, color = _severity(severity)
    style = f' style="color:{_HTML_COLORS[color]}"' if color in _HTML_COLORS else ""
    summary_html = escape(_short(summary, MAX_SUMMARY_CHARS), quote=False).replace("\n", "<br>")
    parts = [
        f"<p><strong>{escape(_short(title, MAX_TITLE_CHARS) or 'Flow Steward', quote=False)}</strong></p>",
        f"<p><strong{style}>{escape(label.upper(), quote=False)}</strong></p>",
        f"<p>{summary_html}</p>",
    ]
    for key, value in _facts(facts):
        parts.append(
            f"<p><strong>{escape(key, quote=False)}:</strong> {escape(value, quote=False)}</p>"
        )
    url = safe_link(link)
    if url:
        text = escape(_short(link_label, 64) or DEFAULT_LINK_LABEL, quote=False)
        parts.append(f'<p><a href="{escape(url)}">{text}</a></p>')
    return "".join(parts)


__all__ = [
    "CARD_CONTENT_TYPE",
    "DEFAULT_LINK_LABEL",
    "SEVERITY_COLORS",
    "notification_card",
    "notification_html",
    "safe_link",
    "webhook_payload",
]
