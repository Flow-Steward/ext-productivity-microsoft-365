"""Sync state of one mail connection and the settings that bind it.

The state lives in the connection's ``ingress_config`` under :data:`STATE_KEY`.
Core passes the connection's ``state_version`` and ``config_generation`` beside
it (see ``docs/TRIGGER_CONTRACT.md``); the extension echoes both back for the
compare-and-set (task A1) and never invents or increments them.
"""

from __future__ import annotations

import re
from collections.abc import Mapping
from dataclasses import asdict, dataclass, replace
from datetime import UTC, datetime, timedelta
from typing import Any

#: ``ingress_config`` key of the sync state.
STATE_KEY = "mail_sync"
#: ``ingress_config`` key Core's poll scheduler reads to apply the interval.
LAST_SYNC_KEY = "last_inbound_sync_at"
#: Trigger settings written by ``save_trigger_settings`` into ``ingress_config``.
MODE_FIELD = "mail_trigger_mode"
INTERVAL_FIELD = "mail_poll_interval_seconds"
FOLDER_FIELD = "mail_trigger_folder"
MODE_ON = "polling"
MODE_OFF = "off"
DEFAULT_FOLDER = "inbox"
STATE_SCHEMA = 1

PHASES = ("baseline", "catch_up", "live", "recovery")
#: The fetch window starts this much before its boundary (clock skew, late indexing).
FETCH_SKEW = timedelta(minutes=5)
#: A recovery query reaches this far behind the watermark.
RECOVERY_OVERLAP = timedelta(minutes=15)

_LINK_LIMIT = 8192
_FOLDER = re.compile(r"^[A-Za-z0-9=_+\-.]{1,512}$")


def utc_now() -> datetime:
    return datetime.now(UTC)


def floor_second(value: datetime) -> datetime:
    """UTC, without microseconds. Graph's ``receivedDateTime`` has whole seconds."""
    return value.astimezone(UTC).replace(microsecond=0)


def iso(value: datetime) -> str:
    return floor_second(value).strftime("%Y-%m-%dT%H:%M:%SZ")


def parse_time(value: Any) -> datetime | None:
    if not isinstance(value, str) or not value.strip() or len(value) > 64:
        return None
    raw = value.strip()
    if raw.endswith(("Z", "z")):
        raw = raw[:-1] + "+00:00"
    try:
        parsed = datetime.fromisoformat(raw)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=UTC)
    return parsed.astimezone(UTC)


def valid_folder(value: str) -> bool:
    """A well-known folder name (``inbox``) or a Graph folder id."""
    return bool(_FOLDER.match(value or ""))


@dataclass(frozen=True)
class SyncState:
    """One connection's sync progress. Times are UTC, whole seconds."""

    phase: str
    #: Delivery boundary, set once per enablement: only mail received at or after it is emitted.
    sync_started_at: datetime
    #: Newest ``receivedDateTime`` already emitted; starts at ``sync_started_at``.
    watermark: datetime
    #: Fetch boundary of the current or next catch-up / recovery query.
    query_from: datetime
    delta_next_link: str = ""
    delta_link: str = ""
    query_next_link: str = ""
    #: Set when a recovery starts; cleared when the following catch-up reaches live.
    recovery_started_at: datetime | None = None
    last_reset_at: datetime | None = None
    last_recovered_at: datetime | None = None
    reset_count: int = 0
    last_error_code: str = ""
    last_error_at: datetime | None = None
    # Binding facts: a state of another configuration is never continued.
    config_generation: str = ""
    folder: str = DEFAULT_FOLDER
    mailbox: str = ""
    #: The address events report; for ``/me`` it is resolved once at sync start.
    mailbox_address: str = ""

    def to_dict(self) -> dict[str, Any]:
        raw = asdict(self)
        out: dict[str, Any] = {"schema": STATE_SCHEMA}
        for key, value in raw.items():
            out[key] = iso(value) if isinstance(value, datetime) else value
            if value is None:
                out[key] = ""
        return out

    @classmethod
    def from_dict(cls, value: Any) -> SyncState | None:
        """Read a stored state; ``None`` when there is no usable state.

        A damaged state whose delivery boundary is still readable continues in
        recovery, so a corrupt cursor never moves the boundary forward.
        """
        if not isinstance(value, Mapping) or not value:
            return None
        started = parse_time(value.get("sync_started_at"))
        if started is None:
            return None
        binding: dict[str, Any] = {
            "config_generation": _text(value.get("config_generation"), 128),
            "folder": _text(value.get("folder"), 512) or DEFAULT_FOLDER,
            "mailbox": _text(value.get("mailbox"), 320),
            "mailbox_address": _text(value.get("mailbox_address"), 320),
        }
        history: dict[str, Any] = {
            "last_reset_at": parse_time(value.get("last_reset_at")),
            "last_recovered_at": parse_time(value.get("last_recovered_at")),
            "reset_count": _count(value.get("reset_count")),
            "last_error_code": _text(value.get("last_error_code"), 64),
            "last_error_at": parse_time(value.get("last_error_at")),
        }
        watermark = parse_time(value.get("watermark")) or started
        watermark = max(watermark, started)
        phase = value.get("phase")
        query_from = parse_time(value.get("query_from"))
        links = {
            "delta_next_link": _link(value.get("delta_next_link")),
            "delta_link": _link(value.get("delta_link")),
            "query_next_link": _link(value.get("query_next_link")),
        }
        if phase not in PHASES or query_from is None or None in links.values():
            # Readable boundary, unreadable progress: recover from the boundary.
            damaged = cls(
                phase="live",
                sync_started_at=started,
                watermark=watermark,
                query_from=started - FETCH_SKEW,
                recovery_started_at=parse_time(value.get("recovery_started_at")),
                **binding,
                **history,
            )
            return replace(damaged, delta_link="", delta_next_link="")
        return cls(
            phase=str(phase),
            sync_started_at=started,
            watermark=watermark,
            query_from=query_from,
            delta_next_link=links["delta_next_link"] or "",
            delta_link=links["delta_link"] or "",
            query_next_link=links["query_next_link"] or "",
            recovery_started_at=parse_time(value.get("recovery_started_at")),
            **binding,
            **history,
        )


def _text(value: Any, limit: int) -> str:
    if isinstance(value, bool) or value is None:
        return ""
    if isinstance(value, int):
        return str(value)
    if not isinstance(value, str):
        return ""
    stripped = value.strip()
    return stripped if len(stripped) <= limit else ""


def _link(value: Any) -> str | None:
    """A stored continuation link: ``""`` when absent, ``None`` when damaged."""
    if value is None or value == "":
        return ""
    if not isinstance(value, str) or len(value) > _LINK_LIMIT:
        return None
    return value if value.startswith("https://") else None


def _count(value: Any) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        return 0
    return max(0, value)


def generation_text(value: Any) -> str:
    """Core's ``config_generation`` as a comparable string (``""`` when absent)."""
    return _text(value, 128)


__all__ = [
    "DEFAULT_FOLDER",
    "FETCH_SKEW",
    "FOLDER_FIELD",
    "INTERVAL_FIELD",
    "LAST_SYNC_KEY",
    "MODE_FIELD",
    "MODE_OFF",
    "MODE_ON",
    "PHASES",
    "RECOVERY_OVERLAP",
    "STATE_KEY",
    "STATE_SCHEMA",
    "SyncState",
    "floor_second",
    "generation_text",
    "iso",
    "parse_time",
    "utc_now",
    "valid_folder",
]
