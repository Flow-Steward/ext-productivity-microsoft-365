"""``poll_mailbox``: the background action behind the new-mail trigger (task B5.2).

One call runs one budgeted poll step for one mail connection and answers with:

- ``events``: one ``microsoft365.mail.message.received`` event per new email,
  ``source_id = m365:<connection_id>:<immutable message id>``;
- ``host_mutations``: one ``project_connection_upsert`` that writes the new sync
  state into this connection's ``ingress_config`` and carries the
  ``state_version`` / ``config_generation`` the poll started from, for Core's
  compare-and-set (task A1).

Where Core puts the inputs and how it must treat the answer is written down in
``docs/TRIGGER_CONTRACT.md``.
"""

from __future__ import annotations

import time
from collections.abc import Callable, Mapping
from dataclasses import dataclass, replace
from datetime import datetime
from typing import Any

from flowsteward_extension_sdk import host_mutation

from m365.auth import transport_for
from m365.connection import (
    ConnectionContext,
    connection_payload,
    mailbox_base,
    mapping,
    resolve_connection,
)
from m365.errors import M365Error, error_from_exception
from m365.transport import GraphTransport

from .graph import MailSyncGraph
from .machine import EventBinding, PollRun, initial_state, run_poll, state_matches
from .state import (
    DEFAULT_FOLDER,
    FOLDER_FIELD,
    LAST_SYNC_KEY,
    MODE_FIELD,
    MODE_ON,
    STATE_KEY,
    SyncState,
    floor_second,
    generation_text,
    iso,
    utc_now,
    valid_folder,
)

ACTION_ID = "poll_mailbox"
#: One poll stays well inside the worker's 30 s call limit.
POLL_BUDGET_SECONDS = 20.0
#: GraphTransport never accepts a shorter overall deadline.
MIN_TRANSPORT_DEADLINE = 5.0

TransportFactory = Callable[[ConnectionContext, float], GraphTransport]


@dataclass(frozen=True)
class PollInputs:
    """What Core passes besides the credentials."""

    connection: dict[str, Any]
    ingress: dict[str, Any]
    mode: str
    folder: str
    #: Echoed unchanged for the compare-and-set; ``None`` when Core sent none.
    state_version: int | str | None
    config_generation: int | str | None


def read_poll_inputs(payload: Mapping[str, Any]) -> PollInputs:
    """Read the selected connection, its ``ingress_config`` and its versions.

    Primary location: ``action.target.connection`` (or ``target.connection``),
    the object operations already receive, with ``ingress_config``,
    ``state_version`` and ``config_generation`` beside ``config`` and
    ``secrets``. Fallback for each of the three: the hydrated row at
    ``runtime_context.resources.project_connection`` of the same connection.
    """
    connection = connection_payload(payload)
    row = mapping(
        mapping(mapping(payload.get("runtime_context")).get("resources")).get("project_connection")
    )
    connection_id = str(connection.get("connection_id") or connection.get("id") or "").strip()
    row_id = str(row.get("connection_id") or row.get("id") or "").strip()
    if row and connection_id and row_id and row_id != connection_id:
        # A snapshot of another connection is never used for this one.
        row = {}
    ingress = mapping(connection.get("ingress_config")) or mapping(row.get("ingress_config"))
    mode = str(ingress.get(MODE_FIELD) or "").strip().lower()
    folder = str(ingress.get(FOLDER_FIELD) or "").strip() or DEFAULT_FOLDER
    return PollInputs(
        connection=connection,
        ingress=ingress,
        mode=mode,
        folder=folder,
        state_version=_version(connection, row, "state_version"),
        config_generation=_version(connection, row, "config_generation"),
    )


def handle_poll_mailbox(
    payload: dict[str, Any],
    *,
    transport_factory: TransportFactory | None = None,
    now: Callable[[], datetime] = utc_now,
    monotonic: Callable[[], float] = time.monotonic,
    budget_seconds: float = POLL_BUDGET_SECONDS,
) -> dict[str, Any]:
    deadline = monotonic() + budget_seconds

    def remaining() -> float:
        return max(0.0, deadline - monotonic())

    try:
        inputs = read_poll_inputs(payload)
        if inputs.mode != MODE_ON:
            # Polling is off unless explicitly switched on (A6); nothing is read or written.
            return _idle_response("disabled", "The new-mail trigger is switched off.")
        context = resolve_connection(payload, service="mail")
        if not context.connection_id:
            raise M365Error("invalid_connection", "The mail connection has no connection id")
        if not valid_folder(inputs.folder):
            raise M365Error("invalid_connection", "The trigger folder setting is invalid")
        factory = transport_factory or _default_transport
        graph = MailSyncGraph(
            factory(context, remaining()),
            mailbox_base=mailbox_base(context),
            folder=inputs.folder,
        )
        state = _current_state(inputs, context, graph, now=now)
        binding = EventBinding(
            connection_id=context.connection_id,
            mailbox=state.mailbox_address or state.mailbox,
            folder=inputs.folder,
        )
        run = run_poll(state, graph, binding=binding, now=now, remaining_seconds=remaining)
    except M365Error as exc:
        return error_from_exception(exc, effectful=False)
    if not run.progressed:
        if run.error is not None:
            return error_from_exception(run.error, effectful=False)
        return _idle_response("deferred", "The poll budget ran out before the first page.")
    return poll_response(inputs, context, run, now=now())


def _current_state(
    inputs: PollInputs,
    context: ConnectionContext,
    graph: MailSyncGraph,
    *,
    now: Callable[[], datetime],
) -> SyncState:
    generation = generation_text(inputs.config_generation)
    mailbox = context.mailbox.lower()
    stored = SyncState.from_dict(inputs.ingress.get(STATE_KEY))
    if stored is not None and state_matches(
        stored, config_generation=generation, folder=inputs.folder, mailbox=mailbox
    ):
        return stored
    # First poll after the switch was turned on, or a state of another configuration.
    address = mailbox or graph.mailbox_address()
    return initial_state(
        now=now(),
        config_generation=generation,
        folder=inputs.folder,
        mailbox=mailbox,
        mailbox_address=address,
    )


def poll_response(
    inputs: PollInputs, context: ConnectionContext, run: PollRun, *, now: datetime
) -> dict[str, Any]:
    """The success envelope: result summary, events and the state mutation."""
    state = run.state
    if run.error is not None:
        state = replace(state, last_error_code=run.error.code, last_error_at=floor_second(now))
    else:
        state = replace(state, last_error_code="", last_error_at=None)
    ingress = {**inputs.ingress, LAST_SYNC_KEY: iso(now), STATE_KEY: state.to_dict()}
    mutation = host_mutation(
        "project_connection_upsert",
        {
            "connection_id": context.connection_id,
            "connection_type": context.connection_type_id,
            "ingress_config": ingress,
            "expected_state_version": inputs.state_version,
            "expected_config_generation": inputs.config_generation,
        },
    )
    emitted = len(run.events)
    result: dict[str, Any] = {
        "status": "partial" if run.error is not None else "polled",
        "phase": state.phase,
        "pages": run.pages,
        "processed_updates": emitted,
        "events_emitted": emitted,
        "caught_up": run.caught_up,
        "sync_reset": run.resets > 0,
        "sync_started_at": iso(state.sync_started_at),
        "last_reset_at": iso(state.last_reset_at) if state.last_reset_at else "",
        "last_recovered_at": iso(state.last_recovered_at) if state.last_recovered_at else "",
        "error_code": run.error.code if run.error is not None else "",
    }
    message = f"Read {run.pages} page(s); {emitted} new email(s)."
    if run.error is not None:
        message = f"{message} Stopped early: {run.error.code}."
    return {
        "ok": True,
        "message": message,
        "result": result,
        "events": list(run.events),
        "host_mutations": [mutation],
    }


def _idle_response(status: str, message: str) -> dict[str, Any]:
    return {
        "ok": True,
        "message": message,
        "result": {"status": status, "processed_updates": 0, "events_emitted": 0},
        "events": [],
    }


def _default_transport(context: ConnectionContext, remaining_seconds: float) -> GraphTransport:
    return transport_for(
        context, deadline_seconds=max(MIN_TRANSPORT_DEADLINE, float(remaining_seconds))
    )


def _version(primary: Mapping[str, Any], fallback: Mapping[str, Any], key: str) -> int | str | None:
    for source in (primary, fallback):
        value = source.get(key)
        if isinstance(value, bool):
            continue
        if isinstance(value, int):
            return value
        if isinstance(value, str) and value.strip() and len(value.strip()) <= 128:
            return value.strip()
    return None


__all__ = [
    "ACTION_ID",
    "POLL_BUDGET_SECONDS",
    "PollInputs",
    "handle_poll_mailbox",
    "poll_response",
    "read_poll_inputs",
]
