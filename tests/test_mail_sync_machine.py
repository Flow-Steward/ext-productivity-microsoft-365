"""Pure sync state machine of the new-mail trigger (task B5.2, B8 sync bullets)."""

from __future__ import annotations

from dataclasses import replace
from datetime import UTC, datetime, timedelta

import pytest

from m365.errors import M365Error
from mail_sync.machine import (
    EVENT_TYPE,
    MAX_RESETS_PER_POLL,
    EventBinding,
    FetchPlan,
    Page,
    apply_page,
    enter_recovery,
    initial_state,
    plan_request,
    run_poll,
    source_id,
    state_matches,
)
from mail_sync.state import SyncState

START = datetime(2026, 11, 22, 10, 0, 0, 700000, tzinfo=UTC)
STARTED = START.replace(microsecond=0)
BINDING = EventBinding(connection_id="conn-acme-1", mailbox="cmr@acme.test", folder="inbox")
DELTA_NEXT = "https://graph.microsoft.com/v1.0/users/cmr@acme.test/mailFolders/inbox/messages/delta?$skiptoken=a"
DELTA_LINK = "https://graph.microsoft.com/v1.0/users/cmr@acme.test/mailFolders/inbox/messages/delta?$deltatoken=b"
QUERY_NEXT = (
    "https://graph.microsoft.com/v1.0/users/cmr@acme.test/mailFolders/inbox/messages?$skip=100"
)


def at(minutes: float) -> str:
    return (STARTED + timedelta(minutes=minutes)).strftime("%Y-%m-%dT%H:%M:%SZ")


def message(message_id: str, minutes: float, **extra: object) -> dict[str, object]:
    return {
        "id": message_id,
        "receivedDateTime": at(minutes),
        "internetMessageId": f"<{message_id}@acme.test>",
        "subject": f"Invoice {message_id}",
        "hasAttachments": True,
        "from": {"emailAddress": {"name": "Acme Billing", "address": "Billing@Acme.test"}},
        **extra,
    }


def fresh() -> SyncState:
    return initial_state(
        now=START,
        config_generation="7",
        folder="inbox",
        mailbox="cmr@acme.test",
        mailbox_address="cmr@acme.test",
    )


def live_state() -> SyncState:
    return replace(fresh(), phase="live", delta_link=DELTA_LINK)


# -- initial state, sync_started_at and the delivery filter -------------------


def test_initial_state_sets_boundaries_once_from_the_first_poll() -> None:
    state = fresh()
    assert state.phase == "baseline"
    assert state.sync_started_at == STARTED
    assert state.watermark == STARTED
    assert state.query_from == STARTED - timedelta(minutes=5)
    assert plan_request(state) == FetchPlan("delta_start")


def test_state_round_trip_keeps_sync_started_at() -> None:
    state = replace(fresh(), delta_next_link=DELTA_NEXT)
    restored = SyncState.from_dict(state.to_dict())
    assert restored == state
    assert restored is not None and restored.sync_started_at == STARTED


def test_state_of_another_configuration_does_not_match() -> None:
    state = fresh()
    assert state_matches(state, config_generation="7", folder="inbox", mailbox="cmr@acme.test")
    assert not state_matches(state, config_generation="8", folder="inbox", mailbox="cmr@acme.test")
    assert not state_matches(
        state, config_generation="7", folder="archive", mailbox="cmr@acme.test"
    )
    assert not state_matches(
        state, config_generation="7", folder="inbox", mailbox="invoices@acme.test"
    )


def test_damaged_progress_keeps_the_boundary_and_recovers() -> None:
    stored = {**fresh().to_dict(), "phase": "nonsense", "delta_link": "http://evil.test/x"}
    state = SyncState.from_dict(stored)
    assert state is not None
    assert state.sync_started_at == STARTED
    assert plan_request(state).kind == "lost"


def test_unreadable_state_without_boundary_is_no_state() -> None:
    assert SyncState.from_dict({"phase": "live"}) is None
    assert SyncState.from_dict(None) is None


def test_delivery_filter_emits_only_mail_at_or_after_sync_started_at() -> None:
    state = replace(fresh(), phase="catch_up")
    page = Page(
        messages=(
            message("old", -3),
            message("edge", 0),
            message("new", 2),
            {"id": "gone", "@removed": {"reason": "deleted"}},
            {"receivedDateTime": at(4)},
        )
    )
    next_state, events = apply_page(state, page, binding=BINDING, now=START)
    assert [event["payload"]["message_id"] for event in events] == ["edge", "new"]
    assert next_state.watermark == STARTED + timedelta(minutes=2)


# -- baseline -> catch-up -> live ---------------------------------------------


def test_baseline_pages_emit_nothing_and_save_delta_progress() -> None:
    state = fresh()
    state, events = apply_page(
        state, Page(messages=(message("m1", 1),), next_link=DELTA_NEXT), binding=BINDING, now=START
    )
    assert events == []
    assert state.phase == "baseline"
    assert state.delta_next_link == DELTA_NEXT
    assert plan_request(state) == FetchPlan("link", link=DELTA_NEXT)
    state, events = apply_page(
        state, Page(messages=(message("m2", 1),), delta_link=DELTA_LINK), binding=BINDING, now=START
    )
    assert events == []
    assert state.phase == "catch_up"
    assert state.delta_link == DELTA_LINK
    assert state.delta_next_link == ""
    assert plan_request(state) == FetchPlan(
        "query_start", query_from=STARTED - timedelta(minutes=5)
    )


def test_catch_up_saves_query_progress_then_goes_live() -> None:
    state = replace(fresh(), phase="catch_up", delta_link=DELTA_LINK)
    state, events = apply_page(
        state, Page(messages=(message("m1", 1),), next_link=QUERY_NEXT), binding=BINDING, now=START
    )
    assert len(events) == 1
    assert state.query_next_link == QUERY_NEXT
    assert plan_request(state) == FetchPlan("link", link=QUERY_NEXT)
    state, events = apply_page(
        state, Page(messages=(message("m2", 2),)), binding=BINDING, now=START
    )
    assert len(events) == 1
    assert state.phase == "live"
    assert state.query_next_link == ""
    assert plan_request(state) == FetchPlan("link", link=DELTA_LINK)


def test_live_follows_next_link_then_keeps_the_new_delta_link() -> None:
    state = live_state()
    state, events = apply_page(
        state, Page(messages=(message("m1", 3),), next_link=DELTA_NEXT), binding=BINDING, now=START
    )
    assert len(events) == 1
    assert state.delta_next_link == DELTA_NEXT
    assert plan_request(state) == FetchPlan("link", link=DELTA_NEXT)
    new_link = DELTA_LINK + "c"
    state, events = apply_page(
        state, Page(messages=(message("m2", 4),), delta_link=new_link), binding=BINDING, now=START
    )
    assert state.delta_link == new_link
    assert state.delta_next_link == ""
    assert state.watermark == STARTED + timedelta(minutes=4)


def test_page_without_continuation_link_is_rejected() -> None:
    with pytest.raises(M365Error) as raised:
        apply_page(live_state(), Page(messages=()), binding=BINDING, now=START)
    assert raised.value.code == "provider_unavailable"


# -- events ---------------------------------------------------------------------


def test_source_id_format() -> None:
    assert source_id("conn-acme-1", "AAMkAD=") == "m365:conn-acme-1:AAMkAD="


def test_event_carries_metadata_only() -> None:
    _, events = apply_page(
        live_state(),
        Page(messages=(message("AAMkImmutable", 1),), delta_link=DELTA_LINK),
        binding=BINDING,
        now=START,
    )
    (event,) = events
    assert event["event_type"] == EVENT_TYPE == "microsoft365.mail.message.received"
    assert event["source_id"] == "m365:conn-acme-1:AAMkImmutable"
    assert event["payload"] == {
        "connection_id": "conn-acme-1",
        "mailbox": "cmr@acme.test",
        "message_id": "AAMkImmutable",
        "internet_message_id": "<AAMkImmutable@acme.test>",
        "from": "billing@acme.test",
        "subject": "Invoice AAMkImmutable",
        "received_at": at(1),
        "has_attachments": True,
        "folder": "inbox",
    }


# -- recovery -------------------------------------------------------------------


def test_recovery_before_the_first_event_uses_sync_started_at() -> None:
    state = enter_recovery(live_state(), now=START + timedelta(minutes=30))
    assert state.watermark == STARTED
    assert state.phase == "recovery"
    assert state.query_from == STARTED - timedelta(minutes=5)
    assert state.delta_link == state.delta_next_link == state.query_next_link == ""
    assert state.last_reset_at == STARTED + timedelta(minutes=30)


def test_recovery_reaches_fifteen_minutes_behind_the_watermark() -> None:
    state = replace(live_state(), watermark=STARTED + timedelta(hours=2))
    state = enter_recovery(state, now=START + timedelta(hours=3))
    assert state.query_from == STARTED + timedelta(hours=2) - timedelta(minutes=20)


def test_recovery_then_new_baseline_then_catch_up_from_recovery_start() -> None:
    reset_at = START + timedelta(hours=1)
    state = enter_recovery(live_state(), now=reset_at)
    state, events = apply_page(
        state, Page(messages=(message("lost-during-reset", 50),)), binding=BINDING, now=reset_at
    )
    assert [event["payload"]["message_id"] for event in events] == ["lost-during-reset"]
    assert state.phase == "baseline"
    assert state.query_from == reset_at.replace(microsecond=0) - timedelta(minutes=5)
    assert plan_request(state) == FetchPlan("delta_start")
    state, _ = apply_page(state, Page(delta_link=DELTA_LINK), binding=BINDING, now=reset_at)
    assert state.phase == "catch_up"
    recovered_at = reset_at + timedelta(minutes=2)
    state, _ = apply_page(state, Page(), binding=BINDING, now=recovered_at)
    assert state.phase == "live"
    assert state.recovery_started_at is None
    assert state.last_recovered_at == recovered_at.replace(microsecond=0)


def test_second_reset_keeps_the_earliest_recovery_start() -> None:
    first = START + timedelta(hours=1)
    state = enter_recovery(live_state(), now=first)
    state = enter_recovery(state, now=first + timedelta(minutes=10))
    assert state.recovery_started_at == first.replace(microsecond=0)
    assert state.reset_count == 2


# -- run_poll: budget, resets, errors -----------------------------------------


class ScriptedClient:
    """Answers FetchPlans from a script; an M365Error entry is raised."""

    def __init__(self, script: list[tuple[FetchPlan | str, object]]) -> None:
        self.script = list(script)
        self.plans: list[FetchPlan] = []

    def fetch(self, plan: FetchPlan) -> Page:
        self.plans.append(plan)
        expected, answer = self.script.pop(0)
        if isinstance(expected, str):
            assert plan.kind == expected
        else:
            assert plan == expected
        if isinstance(answer, M365Error):
            raise answer
        assert isinstance(answer, Page)
        return answer


def budget(seconds_per_call: float, total: float):
    spent = {"value": 0.0}

    def remaining() -> float:
        value = total - spent["value"]
        spent["value"] += seconds_per_call
        return value

    return remaining


def clock():
    return START + timedelta(minutes=1)


def test_run_poll_goes_from_baseline_to_live_in_one_poll_when_budget_allows() -> None:
    client = ScriptedClient(
        [
            ("delta_start", Page(next_link=DELTA_NEXT)),
            (FetchPlan("link", link=DELTA_NEXT), Page(delta_link=DELTA_LINK)),
            ("query_start", Page(messages=(message("arrived-during-baseline", 1),))),
            (FetchPlan("link", link=DELTA_LINK), Page(delta_link=DELTA_LINK + "2")),
        ]
    )
    run = run_poll(fresh(), client, binding=BINDING, now=clock, remaining_seconds=lambda: 20.0)
    assert run.pages == 4
    assert run.caught_up
    assert run.state.phase == "live"
    assert [event["payload"]["message_id"] for event in run.events] == ["arrived-during-baseline"]


def test_run_poll_stops_under_the_time_budget_and_keeps_page_progress() -> None:
    pages = [("delta_start", Page(next_link=DELTA_NEXT + "1"))] + [
        (FetchPlan("link", link=f"{DELTA_NEXT}{index}"), Page(next_link=f"{DELTA_NEXT}{index + 1}"))
        for index in range(1, 10)
    ]
    client = ScriptedClient(pages)
    run = run_poll(fresh(), client, binding=BINDING, now=clock, remaining_seconds=budget(5.0, 20.0))
    # 20, 15, 10, 5 seconds left start a page; 0 left does not.
    assert run.pages == 4
    assert run.state.phase == "baseline"
    assert run.state.delta_next_link == f"{DELTA_NEXT}4"
    # The next poll continues from the saved link.
    assert plan_request(run.state) == FetchPlan("link", link=f"{DELTA_NEXT}4")


def test_run_poll_starts_no_page_after_the_event_cap() -> None:
    state = live_state()
    client = ScriptedClient(
        [
            (
                FetchPlan("link", link=DELTA_LINK),
                Page(messages=tuple(message(f"m{i}", 1) for i in range(3)), next_link=DELTA_NEXT),
            )
        ]
    )
    run = run_poll(
        state, client, binding=BINDING, now=clock, remaining_seconds=lambda: 20.0, max_events=2
    )
    assert run.pages == 1
    assert len(run.events) == 3
    assert run.state.delta_next_link == DELTA_NEXT


def test_run_poll_drops_repeated_messages_within_one_poll() -> None:
    state = replace(fresh(), phase="catch_up", delta_link=DELTA_LINK)
    client = ScriptedClient(
        [
            ("query_start", Page(messages=(message("same", 1),))),
            (
                FetchPlan("link", link=DELTA_LINK),
                Page(messages=(message("same", 1),), delta_link=DELTA_LINK),
            ),
        ]
    )
    run = run_poll(state, client, binding=BINDING, now=clock, remaining_seconds=lambda: 20.0)
    assert [event["source_id"] for event in run.events] == ["m365:conn-acme-1:same"]


@pytest.mark.parametrize(
    ("state", "first_plan"),
    [
        (replace(fresh(), delta_next_link=DELTA_NEXT), FetchPlan("link", link=DELTA_NEXT)),
        (fresh(), FetchPlan("delta_start")),
        (
            replace(fresh(), phase="catch_up", query_next_link=QUERY_NEXT),
            FetchPlan("link", link=QUERY_NEXT),
        ),
        (live_state(), FetchPlan("link", link=DELTA_LINK)),
        (
            replace(enter_recovery(live_state(), now=START), query_next_link=QUERY_NEXT),
            FetchPlan("link", link=QUERY_NEXT),
        ),
    ],
    ids=["baseline-link", "baseline-start", "catch_up", "live", "recovery"],
)
def test_410_in_each_phase_recovers_without_losing_mail(
    state: SyncState, first_plan: FetchPlan
) -> None:
    gone = M365Error("sync_reset", "gone", http_status=410)
    client = ScriptedClient(
        [
            (first_plan, gone),
            ("query_start", Page(messages=(message("sent-during-reset", 1),))),
            ("delta_start", Page(delta_link=DELTA_LINK)),
            ("query_start", Page(messages=(message("sent-during-reset", 1),))),
            (FetchPlan("link", link=DELTA_LINK), Page(delta_link=DELTA_LINK + "2")),
        ]
    )
    run = run_poll(state, client, binding=BINDING, now=clock, remaining_seconds=lambda: 20.0)
    assert run.resets == 1
    assert run.error is None
    assert run.state.phase == "live"
    assert run.state.last_recovered_at is not None
    assert [event["payload"]["message_id"] for event in run.events] == ["sent-during-reset"]
    recovery_query = client.plans[1]
    assert recovery_query.query_from == STARTED - timedelta(minutes=5)


def test_lost_link_in_live_starts_recovery() -> None:
    state = replace(live_state(), delta_link="")
    client = ScriptedClient([("query_start", Page())])
    run = run_poll(state, client, binding=BINDING, now=clock, remaining_seconds=budget(9.0, 20.0))
    assert run.resets == 1
    assert run.state.phase == "baseline"


def test_rejected_stored_link_counts_as_lost() -> None:
    client = ScriptedClient(
        [
            (FetchPlan("link", link=DELTA_LINK), M365Error("invalid_cursor", "not a Graph link")),
            ("query_start", Page()),
        ]
    )
    run = run_poll(
        live_state(), client, binding=BINDING, now=clock, remaining_seconds=budget(9.0, 20.0)
    )
    assert run.resets == 1
    assert run.state.phase == "baseline"


def test_repeated_resets_are_bounded() -> None:
    gone = M365Error("sync_reset", "gone", http_status=410)
    client = ScriptedClient(
        [(FetchPlan("link", link=DELTA_LINK), gone)] + [("query_start", gone)] * MAX_RESETS_PER_POLL
    )
    run = run_poll(live_state(), client, binding=BINDING, now=clock, remaining_seconds=lambda: 20.0)
    assert run.resets == MAX_RESETS_PER_POLL
    assert run.error is not None and run.error.code == "sync_reset"
    assert run.state.phase == "recovery"


def test_error_after_progress_keeps_the_completed_pages() -> None:
    client = ScriptedClient(
        [
            (
                FetchPlan("link", link=DELTA_LINK),
                Page(messages=(message("m1", 1),), next_link=DELTA_NEXT),
            ),
            (
                FetchPlan("link", link=DELTA_NEXT),
                M365Error("rate_limited", "slow down", http_status=429),
            ),
        ]
    )
    run = run_poll(live_state(), client, binding=BINDING, now=clock, remaining_seconds=lambda: 20.0)
    assert run.pages == 1
    assert run.error is not None and run.error.code == "rate_limited"
    assert run.state.delta_next_link == DELTA_NEXT
    assert len(run.events) == 1
    assert run.progressed


def test_folder_not_found_is_an_error_not_a_reset() -> None:
    client = ScriptedClient(
        [
            (
                FetchPlan("link", link=DELTA_LINK),
                M365Error("folder_not_found", "gone", http_status=404),
            )
        ]
    )
    run = run_poll(live_state(), client, binding=BINDING, now=clock, remaining_seconds=lambda: 20.0)
    assert run.resets == 0
    assert run.error is not None and run.error.code == "folder_not_found"
    assert not run.progressed
