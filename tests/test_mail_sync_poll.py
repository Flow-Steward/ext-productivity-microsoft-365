"""``poll_mailbox`` end to end against a fake Graph transport (task B5.2, B8)."""

from __future__ import annotations

import json
from datetime import UTC, datetime, timedelta
from typing import Any

import pytest

from m365.errors import M365Error
from mail_sync.graph import SELECT_FIELDS, MailSyncGraph
from mail_sync.machine import FetchPlan
from mail_sync.poll import handle_poll_mailbox

T0 = datetime(2026, 11, 22, 10, 0, 0, tzinfo=UTC)
BASE = "https://graph.microsoft.com/v1.0/users/cmr@acme.test/mailFolders/inbox/messages"
NEXT = [f"{BASE}/delta?$skiptoken=page{index}" for index in range(5)]
DELTA = [f"{BASE}/delta?$deltatoken=sync{index}" for index in range(5)]
CLIENT_ID = "11111111-2222-3333-4444-555555555555"
SCOPE = {"group_id": None, "for_each_id": None, "step_id": None, "step_execution_id": None}


class FakeTime:
    """Monotonic time that advances by ``per_call`` seconds on every Graph call."""

    def __init__(self, per_call: float = 6.0) -> None:
        self.value = 0.0
        self.per_call = per_call

    def __call__(self) -> float:
        return self.value


class FakeTransport:
    """Routes: ``delta`` (new delta), ``query`` (received-date query), ``me``, or a link."""

    def __init__(self, routes: dict[str, list[Any]], clock: FakeTime | None = None) -> None:
        self.routes = {key: list(value) for key, value in routes.items()}
        self.calls: list[dict[str, Any]] = []
        self.clock = clock

    def get_json(self, path: str, *, query=None, headers=None, purpose: str, **_: Any):
        self.calls.append(
            {"path": path, "query": dict(query or {}), "headers": dict(headers or {})}
        )
        if self.clock is not None:
            self.clock.value += self.clock.per_call
        if path.startswith("https://"):
            key = path
        elif path.endswith("/messages/delta"):
            key = "delta"
        elif path.endswith("/messages"):
            key = "query"
        else:
            key = "me"
        assert self.routes.get(key), f"unexpected Graph call {key}"
        answer = self.routes[key].pop(0)
        if isinstance(answer, M365Error):
            raise answer
        return answer


def received(minutes: float) -> str:
    return (T0 + timedelta(minutes=minutes)).strftime("%Y-%m-%dT%H:%M:%SZ")


def msg(message_id: str, minutes: float) -> dict[str, Any]:
    return {
        "id": message_id,
        "receivedDateTime": received(minutes),
        "internetMessageId": f"<{message_id}@acme.test>",
        "subject": f"CMR {message_id}",
        "hasAttachments": True,
        "from": {"emailAddress": {"address": "Carrier@Acme.test"}},
    }


def page(*messages: dict[str, Any], next_link: str = "", delta_link: str = "") -> dict[str, Any]:
    body: dict[str, Any] = {"value": list(messages)}
    if next_link:
        body["@odata.nextLink"] = next_link
    if delta_link:
        body["@odata.deltaLink"] = delta_link
    return body


def app_connection(ingress: dict[str, Any] | None = None, **extra: Any) -> dict[str, Any]:
    connection = {
        "connection_id": "conn-cmr",
        "connection_type_id": "m365_mail_app",
        "config": {
            "tenant_id": "acme.onmicrosoft.com",
            "client_id": CLIENT_ID,
            "mailbox": "cmr@acme.test",
        },
        "secrets": {"client_secret": "acme-test-secret"},
        "ingress_config": {"mail_trigger_mode": "polling", **(ingress or {})},
        "state_version": 4,
        "config_generation": 2,
    }
    connection.update(extra)
    return connection


def payload(
    connection: dict[str, Any], *, resources: dict[str, Any] | None = None
) -> dict[str, Any]:
    return {
        "contract_version": "extension_host_v2",
        "mode": "action",
        "runtime_context": {**SCOPE, "resources": resources or {}},
        "action": {"action_id": "poll_mailbox", "target": {"connection": connection}, "input": {}},
    }


def poll(
    connection: dict[str, Any],
    transport: FakeTransport,
    *,
    now: datetime = T0,
    clock: FakeTime | None = None,
    resources: dict[str, Any] | None = None,
) -> dict[str, Any]:
    factory_calls: list[float] = []

    def factory(context: Any, remaining: float) -> Any:
        factory_calls.append(remaining)
        return transport

    response = handle_poll_mailbox(
        payload(connection, resources=resources),
        transport_factory=factory,
        now=lambda: now,
        monotonic=clock or FakeTime(),
    )
    response["_factory_calls"] = factory_calls
    return response


def saved_ingress(response: dict[str, Any]) -> dict[str, Any]:
    (mutation,) = response["host_mutations"]
    return mutation["payload"]["ingress_config"]


def event_ids(response: dict[str, Any]) -> list[str]:
    return [event["payload"]["message_id"] for event in response["events"]]


# -- the whole flow over several polls ----------------------------------------


def test_baseline_catch_up_live_across_polls_under_the_budget() -> None:
    clock = FakeTime(per_call=6.0)  # three pages fit into 20 s, a fourth does not
    first = FakeTransport(
        {
            "delta": [page(msg("before-1", -60), next_link=NEXT[1])],
            NEXT[1]: [page(msg("before-2", -50), next_link=NEXT[2])],
            NEXT[2]: [page(next_link=NEXT[3])],
        },
        clock,
    )
    response = poll(app_connection(), first, clock=clock)
    assert response["ok"] is True
    assert response["events"] == []
    assert response["result"]["phase"] == "baseline"
    assert response["result"]["pages"] == 3
    state = saved_ingress(response)["mail_sync"]
    assert state["delta_next_link"] == NEXT[3]
    assert state["sync_started_at"] == "2026-11-22T10:00:00Z"
    assert state["watermark"] == state["sync_started_at"]

    # changeType=created and the full $select go on the first delta request only;
    # links are followed verbatim; every call prefers immutable ids.
    assert first.calls[0]["query"] == {"changeType": "created", "$select": SELECT_FIELDS}
    assert all(call["query"] == {} for call in first.calls[1:])
    assert all('IdType="ImmutableId"' in call["headers"]["Prefer"] for call in first.calls)

    clock = FakeTime(per_call=6.0)
    second = FakeTransport(
        {
            NEXT[3]: [page(delta_link=DELTA[0])],
            "query": [page(msg("before-start", -2), msg("during-baseline", 0.5))],
            DELTA[0]: [page(msg("live-1", 1), delta_link=DELTA[1])],
        },
        clock,
    )
    connection = app_connection(saved_ingress(response), state_version=5)
    response = poll(connection, second, now=T0 + timedelta(minutes=1), clock=clock)
    assert event_ids(response) == ["during-baseline", "live-1"]
    assert response["result"]["phase"] == "live"
    assert response["result"]["caught_up"] is True
    assert second.calls[1]["query"]["$filter"] == "receivedDateTime ge 2026-11-22T09:55:00Z"
    assert second.calls[1]["query"]["$orderby"] == "receivedDateTime asc"
    state = saved_ingress(response)["mail_sync"]
    assert state["sync_started_at"] == "2026-11-22T10:00:00Z"  # set once
    assert state["delta_link"] == DELTA[1]
    assert state["watermark"] == received(1)

    third = FakeTransport({DELTA[1]: [page(msg("live-2", 3), delta_link=DELTA[2])]})
    connection = app_connection(saved_ingress(response), state_version=6)
    response = poll(connection, third, now=T0 + timedelta(minutes=4))
    assert event_ids(response) == ["live-2"]
    assert response["events"][0]["source_id"] == "m365:conn-cmr:live-2"
    assert saved_ingress(response)["mail_sync"]["sync_started_at"] == "2026-11-22T10:00:00Z"


def test_result_and_events_never_carry_graph_links() -> None:
    transport = FakeTransport(
        {
            "delta": [page(next_link=NEXT[1])],
            NEXT[1]: [page(delta_link=DELTA[0])],
            "query": [page()],
            DELTA[0]: [page(delta_link=DELTA[1])],
        }
    )
    response = poll(app_connection(), transport)
    visible = json.dumps({"result": response["result"], "events": response["events"]})
    assert "https://" not in visible
    assert "deltatoken" not in visible


# -- the shape Core receives ----------------------------------------------------


def test_state_goes_back_as_one_upsert_with_the_compare_and_set_echo() -> None:
    transport = FakeTransport(
        {
            "delta": [page(delta_link=DELTA[0])],
            "query": [page(msg("m1", 1))],
            DELTA[0]: [page(delta_link=DELTA[1])],
        }
    )
    connection = app_connection({"mail_poll_interval_seconds": 120, "mail_trigger_folder": "inbox"})
    response = poll(connection, transport, now=T0 + timedelta(seconds=30))
    (mutation,) = response["host_mutations"]
    assert mutation["kind"] == "project_connection_upsert"
    body = mutation["payload"]
    assert body["connection_id"] == "conn-cmr"
    assert body["connection_type"] == "m365_mail_app"
    assert body["expected_state_version"] == 4
    assert body["expected_config_generation"] == 2
    ingress = body["ingress_config"]
    assert ingress["mail_trigger_mode"] == "polling"
    assert ingress["mail_poll_interval_seconds"] == 120
    assert ingress["last_inbound_sync_at"] == "2026-11-22T10:00:30Z"
    assert ingress["mail_sync"]["config_generation"] == "2"
    assert ingress["mail_sync"]["phase"] == "live"
    (event,) = response["events"]
    assert event["event_type"] == "microsoft365.mail.message.received"
    assert event["source_id"] == "m365:conn-cmr:m1"
    assert set(event["payload"]) == {
        "connection_id",
        "mailbox",
        "message_id",
        "internet_message_id",
        "from",
        "subject",
        "received_at",
        "has_attachments",
        "folder",
    }
    assert event["payload"]["mailbox"] == "cmr@acme.test"
    assert event["payload"]["from"] == "carrier@acme.test"
    assert response["result"]["processed_updates"] == 1


@pytest.mark.parametrize("mode", [None, "", "off", "auto", "webhook"])
def test_poll_does_nothing_unless_the_switch_is_on(mode: str | None) -> None:
    ingress = {} if mode is None else {"mail_trigger_mode": mode}
    connection = app_connection()
    connection["ingress_config"] = ingress
    transport = FakeTransport({})
    response = poll(connection, transport)
    assert response["ok"] is True
    assert response["result"]["status"] == "disabled"
    assert response["events"] == []
    assert "host_mutations" not in response
    assert response["_factory_calls"] == []
    assert transport.calls == []


def test_new_config_generation_starts_a_new_sync() -> None:
    old = FakeTransport(
        {
            "delta": [page(delta_link=DELTA[0])],
            "query": [page()],
            DELTA[0]: [page(delta_link=DELTA[1])],
        }
    )
    stored = saved_ingress(poll(app_connection(), old))
    later = T0 + timedelta(hours=2)
    clock = FakeTime(6.0)
    transport = FakeTransport(
        {
            "delta": [page(msg("while-off", 30), next_link=NEXT[1])],
            NEXT[1]: [page(next_link=NEXT[2])],
            NEXT[2]: [page(next_link=NEXT[3])],
        },
        clock,
    )
    response = poll(app_connection(stored, config_generation=3), transport, now=later, clock=clock)
    state = saved_ingress(response)["mail_sync"]
    assert transport.calls[0]["query"]["changeType"] == "created"
    assert state["sync_started_at"] == "2026-11-22T12:00:00Z"
    assert state["config_generation"] == "3"
    assert response["events"] == []


def test_changed_folder_starts_a_new_sync_on_the_new_folder() -> None:
    old = FakeTransport(
        {
            "delta": [page(delta_link=DELTA[0])],
            "query": [page()],
            DELTA[0]: [page(delta_link=DELTA[1])],
        }
    )
    stored = saved_ingress(poll(app_connection(), old))
    stored["mail_trigger_folder"] = "AAMkFolder="
    transport = FakeTransport(
        {
            "delta": [page(delta_link=DELTA[2])],
            "query": [page()],
            DELTA[2]: [page(delta_link=DELTA[3])],
        }
    )
    response = poll(app_connection(stored), transport)
    assert (
        transport.calls[0]["path"]
        == "/users/cmr@acme.test/mailFolders/AAMkFolder%3D/messages/delta"
    )
    assert saved_ingress(response)["mail_sync"]["folder"] == "AAMkFolder="


# -- errors and recovery through the handler ----------------------------------


def test_failure_before_any_page_is_an_error_without_state() -> None:
    transport = FakeTransport(
        {"delta": [M365Error("permission_denied", "consent", http_status=403)]}
    )
    response = poll(app_connection(), transport)
    assert response["ok"] is False
    assert response["error_code"] == "permission_denied"
    assert "host_mutations" not in response


def test_failure_after_progress_commits_the_progress() -> None:
    stored = {"mail_sync": _live_state(DELTA[0])}
    transport = FakeTransport(
        {
            DELTA[0]: [page(msg("m1", 1), next_link=NEXT[1])],
            NEXT[1]: [M365Error("rate_limited", "slow down", http_status=429)],
        }
    )
    response = poll(app_connection(stored), transport)
    assert response["ok"] is True
    assert response["result"]["status"] == "partial"
    assert response["result"]["error_code"] == "rate_limited"
    assert event_ids(response) == ["m1"]
    state = saved_ingress(response)["mail_sync"]
    assert state["delta_next_link"] == NEXT[1]
    assert state["last_error_code"] == "rate_limited"


def test_410_in_live_recovers_and_delivers_mail_sent_during_the_reset() -> None:
    stored = {"mail_sync": _live_state(DELTA[0], watermark=received(40))}
    transport = FakeTransport(
        {
            DELTA[0]: [M365Error("sync_reset", "gone", http_status=410)],
            "query": [page(msg("sent-during-reset", 50)), page(msg("sent-during-reset", 50))],
            "delta": [page(delta_link=DELTA[1])],
            DELTA[1]: [page(delta_link=DELTA[2])],
        }
    )
    response = poll(app_connection(stored), transport, now=T0 + timedelta(hours=1))
    assert event_ids(response) == ["sent-during-reset"]
    assert response["result"]["sync_reset"] is True
    assert response["result"]["last_reset_at"] == "2026-11-22T11:00:00Z"
    assert response["result"]["last_recovered_at"] == "2026-11-22T11:00:00Z"
    recovery_filter = transport.calls[1]["query"]["$filter"]
    assert recovery_filter == "receivedDateTime ge 2026-11-22T10:20:00Z"  # 40 - 15 - 5 min
    catch_up_filter = transport.calls[3]["query"]["$filter"]
    assert catch_up_filter == "receivedDateTime ge 2026-11-22T10:55:00Z"  # reset - 5 min


def test_delegated_connection_resolves_its_address_once() -> None:
    connection = {
        "connection_id": "conn-pat",
        "connection_type_id": "m365_mail_user",
        "config": {},
        "secrets": {"access_token": "acme-test-token"},
        "ingress_config": {"mail_trigger_mode": "polling"},
        "state_version": 1,
        "config_generation": 1,
    }
    transport = FakeTransport(
        {
            "me": [{"mail": "Pat@Acme.test"}],
            "delta": [page(delta_link=DELTA[0])],
            "query": [page(msg("m1", 1))],
            DELTA[0]: [page(delta_link=DELTA[1])],
        }
    )
    response = poll(connection, transport)
    assert transport.calls[0]["path"] == "/me"
    assert transport.calls[1]["path"] == "/me/mailFolders/inbox/messages/delta"
    assert response["events"][0]["payload"]["mailbox"] == "pat@acme.test"
    connection["ingress_config"] = saved_ingress(response)
    again = FakeTransport({DELTA[1]: [page(msg("m2", 2), delta_link=DELTA[2])]})
    response = poll(connection, again)
    assert [call["path"] for call in again.calls] == [DELTA[1]]
    assert response["events"][0]["payload"]["mailbox"] == "pat@acme.test"


def test_inputs_fall_back_to_the_hydrated_row_of_the_same_connection() -> None:
    connection = app_connection()
    for key in ("ingress_config", "state_version", "config_generation"):
        connection.pop(key)
    row = {
        "connection_id": "conn-cmr",
        "ingress_config": {"mail_trigger_mode": "polling"},
        "state_version": 9,
        "config_generation": "g-4",
    }
    transport = FakeTransport(
        {
            "delta": [page(delta_link=DELTA[0])],
            "query": [page()],
            DELTA[0]: [page(delta_link=DELTA[1])],
        }
    )
    response = poll(connection, transport, resources={"project_connection": row})
    body = response["host_mutations"][0]["payload"]
    assert body["expected_state_version"] == 9
    assert body["expected_config_generation"] == "g-4"


def test_a_row_of_another_connection_is_ignored() -> None:
    connection = app_connection()
    connection.pop("ingress_config")
    row = {"connection_id": "conn-invoices", "ingress_config": {"mail_trigger_mode": "polling"}}
    response = poll(connection, FakeTransport({}), resources={"project_connection": row})
    assert response["result"]["status"] == "disabled"


def test_files_connection_is_refused() -> None:
    connection = app_connection(connection_type_id="m365_files_app")
    response = poll(connection, FakeTransport({}))
    assert response["ok"] is False
    assert response["error_code"] == "invalid_connection"


def test_missing_connection_is_refused() -> None:
    body = payload(app_connection())
    body["action"]["target"] = {"connection_id": "conn-cmr"}
    response = handle_poll_mailbox(body, transport_factory=lambda *_: FakeTransport({}))
    assert response["ok"] is False
    assert response["error_code"] == "invalid_connection"


def test_deleted_folder_is_folder_not_found() -> None:
    transport = FakeTransport({DELTA[0]: [M365Error("item_not_found", "gone", http_status=404)]})
    graph = MailSyncGraph(transport, mailbox_base="/users/cmr@acme.test", folder="inbox")  # type: ignore[arg-type]
    with pytest.raises(M365Error) as raised:
        graph.fetch(FetchPlan("link", link=DELTA[0]))
    assert raised.value.code == "folder_not_found"


def _live_state(delta_link: str, *, watermark: str = "2026-11-22T10:00:00Z") -> dict[str, Any]:
    return {
        "schema": 1,
        "phase": "live",
        "sync_started_at": "2026-11-22T10:00:00Z",
        "watermark": watermark,
        "query_from": "2026-11-22T09:55:00Z",
        "delta_link": delta_link,
        "delta_next_link": "",
        "query_next_link": "",
        "config_generation": "2",
        "folder": "inbox",
        "mailbox": "cmr@acme.test",
        "mailbox_address": "cmr@acme.test",
    }
