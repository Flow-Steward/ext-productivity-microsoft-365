"""Neutral notification content built from a Core event.

Core gives an email route its standard subject (``route_config.subject``) and body
(``event.payload.message``). Channel routes get no formatting from Core, so the
title, severity, summary and run link are built here.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from typing import Any
from urllib.parse import urlsplit

from m365.connection import mapping, text

MAX_SUMMARY_CHARS = 4000

_TITLES = {
    "job.failed": "Job failed",
    "job.completed": "Job completed",
    "job.awaiting_approval": "Job is waiting for approval",
    "approval.requested": "Approval requested",
    "approval.resolved": "Approval resolved",
    "notification.delivery.failed": "Notification delivery failed",
}


@dataclass(frozen=True)
class NotificationContent:
    title: str
    severity: str
    summary: str
    run_url: str
    details: tuple[tuple[str, str], ...]


def _line(value: Any, limit: int = 512) -> str:
    return text(value, limit=limit) if isinstance(value, str) else ""


def safe_run_url(value: Any) -> str:
    """Core's own run link: an absolute http(s) URL without credentials, or nothing."""
    if not isinstance(value, str) or len(value) > 2048:
        return ""
    candidate = value.strip()
    parts = urlsplit(candidate)
    if parts.scheme not in {"http", "https"} or not parts.hostname:
        return ""
    if parts.username or parts.password or any(ord(char) < 32 for char in candidate):
        return ""
    return candidate


def event_content(event: Mapping[str, Any]) -> NotificationContent:
    payload = mapping(event.get("payload"))
    event_type = _line(event.get("event_type"), 128)
    source_type = _line(payload.get("source_event_type"), 128)
    if source_type in _TITLES and event_type not in _TITLES:
        event_type = source_type
    title = _TITLES.get(event_type) or (
        event_type.replace(".", " ").replace("_", " ").capitalize() or "Notification"
    )
    workflow = (
        _line(payload.get("workflow_display_name"))
        or _line(payload.get("workflow_name"))
        or _line(event.get("workflow_id"))
    )
    if workflow:
        title = f"{title}: {workflow}"
    severity = (_line(event.get("severity"), 32) or "info").lower()
    raw_summary = payload.get("message")
    summary = raw_summary.strip() if isinstance(raw_summary, str) else ""
    if len(summary) > MAX_SUMMARY_CHARS:
        summary = summary[: MAX_SUMMARY_CHARS - 1] + "…"
    details = tuple(
        (label, value)
        for label, value in (
            ("Step", _line(payload.get("step_label")) or _line(event.get("step_id"))),
            ("Due", _line(payload.get("due_at"))),
            ("Error", _line(payload.get("error"), 1024)),
        )
        if value
    )
    return NotificationContent(
        title=title,
        severity=severity,
        summary=summary,
        run_url=safe_run_url(event.get("run_url")),
        details=details,
    )


def email_subject_and_body(
    event: Mapping[str, Any], route_config: Mapping[str, Any]
) -> tuple[str, str]:
    """Core's standard subject and body when present, else a neutral fallback."""
    content = event_content(event)
    subject = (
        _line(route_config.get("subject"), 255) or f"[{content.severity.upper()}] {content.title}"
    )
    body = content.summary
    if not body:
        lines = [content.title, ""]
        lines += [f"{label}: {value}" for label, value in content.details]
        if content.run_url:
            lines += ["", f"Open the run: {content.run_url}"]
        body = "\n".join(lines).strip()
    return subject, body + "\n"


__all__ = [
    "NotificationContent",
    "email_subject_and_body",
    "event_content",
    "safe_run_url",
]
