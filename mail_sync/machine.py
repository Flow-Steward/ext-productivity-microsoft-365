"""Pure sync state machine for the new-mail trigger (task B5.2).

No I/O here. The driver asks :func:`plan_request` what to fetch, fetches it,
and hands the page to :func:`apply_page`, which returns the next state and the
events of that page. A lost or expired cursor goes through
:func:`enter_recovery`. :func:`run_poll` is the budgeted loop around them; it
takes the Graph client and the clocks as arguments.

Phases:

1. ``baseline``: page through delta (``changeType=created``) without emitting,
   until the ``deltaLink`` arrives; then ``catch_up``.
2. ``catch_up``: query ``receivedDateTime ge query_from`` page by page and emit
   mail received at or after ``sync_started_at``; then ``live``.
3. ``live``: follow the delta link and emit, with the same delivery filter.
4. ``recovery`` (410, ``SyncStateNotFound`` or a lost link, in any phase): query
   from ``max(sync_started_at, watermark - 15 min) - 5 min``, then a new
   baseline whose catch-up starts at ``recovery start - 5 min``.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Mapping
from dataclasses import dataclass, field, replace
from datetime import datetime
from typing import Any, Protocol

from m365.errors import M365Error

from .state import (
    DEFAULT_FOLDER,
    FETCH_SKEW,
    RECOVERY_OVERLAP,
    SyncState,
    floor_second,
    iso,
    parse_time,
)

EVENT_TYPE = "microsoft365.mail.message.received"
EVENT_SOURCE = "microsoft365.mail"
SOURCE_ID_PREFIX = "m365"
SUBJECT_LIMIT = 1000
MESSAGE_ID_LIMIT = 1024
#: A poll that already holds this many events starts no further page.
MAX_EVENTS_PER_POLL = 200
#: A page is started only when at least this much of the budget is left.
PAGE_RESERVE_SECONDS = 3.0
#: Recovery restarts within one poll (a second 410 right after a reset) are bounded.
MAX_RESETS_PER_POLL = 2

#: Errors on a *stored* link that mean the link is lost, not that the mailbox failed.
_LOST_LINK_CODES = frozenset({"sync_reset", "invalid_cursor", "invalid_payload"})


@dataclass(frozen=True)
class FetchPlan:
    """What the next request is. ``kind`` is one of:

    ``delta_start`` (new delta with ``changeType=created``), ``query_start`` (new
    ``receivedDateTime ge query_from`` query), ``link`` (follow a stored
    ``nextLink``/``deltaLink`` verbatim) or ``lost`` (the phase needs a link the
    state does not have).
    """

    kind: str
    link: str = ""
    query_from: datetime | None = None


@dataclass(frozen=True)
class Page:
    messages: tuple[Mapping[str, Any], ...] = ()
    next_link: str = ""
    delta_link: str = ""


@dataclass(frozen=True)
class EventBinding:
    connection_id: str
    mailbox: str
    folder: str = DEFAULT_FOLDER


class SyncClient(Protocol):
    def fetch(self, plan: FetchPlan) -> Page: ...


def initial_state(
    *,
    now: datetime,
    config_generation: str,
    folder: str,
    mailbox: str,
    mailbox_address: str,
) -> SyncState:
    """The state of the first poll after the switch is turned on."""
    started = floor_second(now)
    return SyncState(
        phase="baseline",
        sync_started_at=started,
        watermark=started,
        query_from=started - FETCH_SKEW,
        config_generation=config_generation,
        folder=folder,
        mailbox=mailbox,
        mailbox_address=mailbox_address,
    )


def state_matches(state: SyncState, *, config_generation: str, folder: str, mailbox: str) -> bool:
    """Whether a stored state belongs to the current configuration.

    A different ``config_generation`` (switch turned off and on, folder or
    mailbox changed, reconnect) or a different folder or mailbox starts a new
    sync with a new ``sync_started_at``.
    """
    return (
        state.config_generation == config_generation
        and state.folder == folder
        and state.mailbox == mailbox
    )


def plan_request(state: SyncState) -> FetchPlan:
    if state.phase == "baseline":
        if state.delta_next_link:
            return FetchPlan("link", link=state.delta_next_link)
        return FetchPlan("delta_start")
    if state.phase in {"catch_up", "recovery"}:
        if state.query_next_link:
            return FetchPlan("link", link=state.query_next_link)
        return FetchPlan("query_start", query_from=state.query_from)
    link = state.delta_next_link or state.delta_link
    return FetchPlan("link", link=link) if link else FetchPlan("lost")


def enter_recovery(state: SyncState, *, now: datetime) -> SyncState:
    """Start (or restart) a recovery after a lost or expired cursor."""
    moment = floor_second(now)
    recovery_started_at = state.recovery_started_at or moment
    query_from = max(state.sync_started_at, state.watermark - RECOVERY_OVERLAP) - FETCH_SKEW
    return replace(
        state,
        phase="recovery",
        query_from=query_from,
        query_next_link="",
        delta_next_link="",
        delta_link="",
        recovery_started_at=recovery_started_at,
        last_reset_at=moment,
        reset_count=state.reset_count + 1,
    )


def apply_page(
    state: SyncState, page: Page, *, binding: EventBinding, now: datetime
) -> tuple[SyncState, list[dict[str, Any]]]:
    """The next state and the events of one fully read page."""
    if state.phase == "baseline":
        if page.next_link:
            return replace(state, delta_next_link=page.next_link), []
        if page.delta_link:
            return (
                replace(
                    state,
                    phase="catch_up",
                    delta_link=page.delta_link,
                    delta_next_link="",
                    query_next_link="",
                ),
                [],
            )
        raise _missing_link()
    events, watermark = emit_messages(state, page.messages, binding=binding)
    if state.phase in {"catch_up", "recovery"}:
        if page.next_link:
            return replace(state, query_next_link=page.next_link, watermark=watermark), events
        if state.phase == "catch_up":
            recovered = state.recovery_started_at is not None
            return (
                replace(
                    state,
                    phase="live",
                    query_next_link="",
                    watermark=watermark,
                    recovery_started_at=None,
                    last_recovered_at=floor_second(now) if recovered else state.last_recovered_at,
                ),
                events,
            )
        started = state.recovery_started_at or floor_second(now)
        return (
            replace(
                state,
                phase="baseline",
                query_next_link="",
                delta_next_link="",
                delta_link="",
                query_from=started - FETCH_SKEW,
                recovery_started_at=started,
                watermark=watermark,
            ),
            events,
        )
    if page.next_link:
        return replace(state, delta_next_link=page.next_link, watermark=watermark), events
    if page.delta_link:
        return (
            replace(state, delta_link=page.delta_link, delta_next_link="", watermark=watermark),
            events,
        )
    raise _missing_link()


def emit_messages(
    state: SyncState, messages: Iterable[Mapping[str, Any]], *, binding: EventBinding
) -> tuple[list[dict[str, Any]], datetime]:
    """Events for the messages at or after the delivery boundary, and the new watermark."""
    events: list[dict[str, Any]] = []
    watermark = state.watermark
    for message in messages:
        if not isinstance(message, Mapping) or "@removed" in message:
            continue
        message_id = message.get("id")
        if not isinstance(message_id, str) or not message_id.strip():
            continue
        if len(message_id) > MESSAGE_ID_LIMIT:
            continue
        received = parse_time(message.get("receivedDateTime"))
        if received is None or received < state.sync_started_at:
            continue
        events.append(build_event(message, received=received, binding=binding))
        watermark = max(watermark, floor_second(received))
    return events, watermark


def source_id(connection_id: str, message_id: str) -> str:
    return f"{SOURCE_ID_PREFIX}:{connection_id}:{message_id}"


def build_event(
    message: Mapping[str, Any], *, received: datetime, binding: EventBinding
) -> dict[str, Any]:
    """Metadata-only event; the workflow fetches body and attachments itself."""
    message_id = str(message["id"]).strip()
    key = source_id(binding.connection_id, message_id)
    sender = _mapping(_mapping(message.get("from")).get("emailAddress"))
    from_address = _clean(sender.get("address"), 320).lower()
    payload = {
        "connection_id": binding.connection_id,
        "mailbox": binding.mailbox,
        "message_id": message_id,
        "internet_message_id": _clean(message.get("internetMessageId"), 998),
        "from": from_address,
        "subject": _clean(message.get("subject"), SUBJECT_LIMIT),
        "received_at": iso(received),
        "has_attachments": message.get("hasAttachments") is True,
        "folder": binding.folder,
    }
    who = from_address or "an unknown sender"
    where = binding.mailbox or "the mailbox"
    return {
        "event_type": EVENT_TYPE,
        "source_id": key,
        "correlation_id": key,
        "source": EVENT_SOURCE,
        "severity": "info",
        "summary": f"New email from {who} in {where}",
        "payload": payload,
    }


@dataclass
class PollRun:
    """Result of one budgeted poll step."""

    state: SyncState
    events: list[dict[str, Any]] = field(default_factory=list)
    pages: int = 0
    resets: int = 0
    caught_up: bool = False
    error: M365Error | None = None

    @property
    def progressed(self) -> bool:
        """Whether there is anything to commit: a page was read or a reset recorded."""
        return self.pages > 0 or self.resets > 0


def run_poll(
    state: SyncState,
    client: SyncClient,
    *,
    binding: EventBinding,
    now: Callable[[], datetime],
    remaining_seconds: Callable[[], float],
    max_events: int = MAX_EVENTS_PER_POLL,
) -> PollRun:
    """Advance ``state`` page by page while the budget lasts.

    Every page is applied whole or not at all, so the returned state is always a
    checkpoint right after a complete page: Core commits it together with the
    events read up to it. An error stops the loop and keeps the progress made
    before it.
    """
    run = PollRun(state=state)
    seen: set[str] = set()
    while len(run.events) < max_events and remaining_seconds() >= PAGE_RESERVE_SECONDS:
        plan = plan_request(run.state)
        if plan.kind == "lost":
            if not _may_reset(run):
                break
            run.state = enter_recovery(run.state, now=now())
            run.resets += 1
            continue
        try:
            page = client.fetch(plan)
            next_state, events = apply_page(run.state, page, binding=binding, now=now())
        except M365Error as exc:
            lost = exc.code == "sync_reset" or (
                plan.kind == "link" and exc.code in _LOST_LINK_CODES
            )
            if lost and _may_reset(run):
                run.state = enter_recovery(run.state, now=now())
                run.resets += 1
                continue
            run.error = exc
            break
        run.state = next_state
        run.pages += 1
        for event in events:
            if event["source_id"] not in seen:
                seen.add(event["source_id"])
                run.events.append(event)
        if run.state.phase == "live" and page.delta_link and not page.next_link:
            run.caught_up = True
            break
    return run


def _may_reset(run: PollRun) -> bool:
    if run.resets < MAX_RESETS_PER_POLL:
        return True
    run.error = M365Error("sync_reset", "The Graph sync state was reset repeatedly")
    return False


def _missing_link() -> M365Error:
    return M365Error(
        "provider_unavailable", "Microsoft Graph returned a page without a continuation link"
    )


def _mapping(value: Any) -> Mapping[str, Any]:
    return value if isinstance(value, Mapping) else {}


def _clean(value: Any, limit: int) -> str:
    if not isinstance(value, str):
        return ""
    cleaned = "".join(" " if ord(char) < 32 or 127 <= ord(char) < 160 else char for char in value)
    return cleaned.strip()[:limit]


__all__ = [
    "EVENT_TYPE",
    "MAX_EVENTS_PER_POLL",
    "PAGE_RESERVE_SECONDS",
    "EventBinding",
    "FetchPlan",
    "Page",
    "PollRun",
    "SyncClient",
    "apply_page",
    "build_event",
    "emit_messages",
    "enter_recovery",
    "initial_state",
    "plan_request",
    "run_poll",
    "source_id",
    "state_matches",
]
