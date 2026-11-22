"""deliver_notification (task D4): routing by connection type, Core's key, mail path,
and the closed failure envelope."""

from __future__ import annotations

from typing import Any

import pytest
from test_mail_fakes import APP_CONNECTION, USER_CONNECTION, FakeOutlook

import dispatcher
from mail import send
from notify.content import event_content, safe_run_url
from notify.deliver import handle_deliver_notification

CORE_KEY = "a" * 64
RESET_LINK = "https://fs.example/reset?token=SECRET-RESET-TOKEN"


@pytest.fixture(autouse=True)
def _no_waiting(monkeypatch: pytest.MonkeyPatch) -> None:
    now = [0.0]

    def fake_sleep(seconds: float) -> None:
        now[0] += seconds

    monkeypatch.setattr(send, "sleep", fake_sleep)
    monkeypatch.setattr(send, "clock", lambda: now[0])


def delivery(
    connection: dict[str, Any],
    route_config: dict[str, Any],
    *,
    event: dict[str, Any] | None = None,
    key: str | None = CORE_KEY,
    test_mode: bool = False,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "contract_version": "extension_host_v2",
        "runtime_context": {
            "group_id": None,
            "for_each_id": None,
            "step_id": None,
            "step_execution_id": None,
            "test_mode": test_mode,
        },
        "action": "deliver_notification",
        "action_payload": {"target": {"connection": dict(connection)}},
        "event": event
        or {
            "event_id": "evt-1",
            "event_type": "job.failed",
            "severity": "error",
            "payload": {"message": "Job 42 failed at step Upload.", "workflow_name": "Invoices"},
            "run_url": "https://fs.example/jobs/42",
        },
        "route": {"route_id": "route-1", "route_config": dict(route_config)},
        "provider": {"provider_key": "flowsteward.microsoft-365"},
    }
    if key is not None:
        payload["idempotency_key"] = key
    return payload


def run(payload: dict[str, Any], fake: Any = None, **kwargs: Any) -> dict[str, Any]:
    return handle_deliver_notification(
        payload, transport_factory=lambda _connection: fake, **kwargs
    )


def test_mail_route_sends_one_email_with_core_subject_and_body() -> None:
    fake = FakeOutlook()
    response = run(
        delivery(USER_CONNECTION, {"to": "ops@acme.example", "subject": "Job failed: Invoices"}),
        fake,
    )
    assert response["ok"] is True and response["delivered"] is True
    assert response["provider_message_id"]
    [sent] = fake.sent()
    assert sent.subject == "Job failed: Invoices"
    assert "Job 42 failed" in sent.body["content"]


def test_redelivery_with_the_same_core_key_does_not_send_twice() -> None:
    fake = FakeOutlook()
    payload = delivery(USER_CONNECTION, {"to": ["ops@acme.example"]})
    first = run(payload, fake)
    second = run(payload, fake)
    assert len(fake.sent()) == 1
    assert second["delivered"] is True
    assert second["provider_message_id"] == first["provider_message_id"]
    assert second["result"]["confirmation"] == "already_sent"


def test_a_different_core_key_is_a_different_email() -> None:
    fake = FakeOutlook()
    run(delivery(USER_CONNECTION, {"to": "ops@acme.example"}), fake)
    run(delivery(USER_CONNECTION, {"to": "ops@acme.example"}, key="b" * 64), fake)
    assert len(fake.sent()) == 2


def test_app_only_mail_sends_from_the_connection_mailbox() -> None:
    fake = FakeOutlook(base="/users/cmr@acme.example")
    response = run(delivery(APP_CONNECTION, {"to": "ops@acme.example"}), fake)
    assert response["delivered"] is True
    assert all("/users/cmr@acme.example/" in call.url for call in fake.calls)


def test_missing_core_key_refuses_before_any_write() -> None:
    fake = FakeOutlook()
    response = run(delivery(USER_CONNECTION, {"to": "ops@acme.example"}, key=None), fake)
    assert response["ok"] is False
    assert response["error_code"] == "idempotency_key_missing"
    assert response["definitely_no_external_effect"] is True
    assert fake.calls == []


def test_recipient_required() -> None:
    response = run(delivery(USER_CONNECTION, {}), FakeOutlook())
    assert response["error_code"] == "recipient_required"
    assert response["delivered"] is False


def test_mail_route_with_a_channel_is_incompatible() -> None:
    response = run(
        delivery(USER_CONNECTION, {"to": "ops@acme.example", "channel": "t/c"}), FakeOutlook()
    )
    assert response["error_code"] == "notification_connection_incompatible"


def test_non_notification_connection_type_is_incompatible() -> None:
    files = {
        "connection_id": "conn-x",
        "connection_type_id": "m365_files_app",
        "config": {
            "tenant_id": "acme.onmicrosoft.com",
            "client_id": "11111111-2222-3333-4444-555555555555",
            "site_url": "https://acme.sharepoint.com/sites/ops",
            "drive_id": "b!drive",
        },
        "secrets": {"client_secret": "app-secret"},
    }
    response = run(delivery(files, {"to": "ops@acme.example"}), FakeOutlook())
    assert response["error_code"] == "notification_connection_incompatible"


def test_test_mode_suppresses_without_calls() -> None:
    fake = FakeOutlook()
    response = run(delivery(USER_CONNECTION, {"to": "ops@acme.example"}, test_mode=True), fake)
    assert response["delivered"] is False
    assert response["external_effect_status"] == "suppressed"
    assert fake.calls == []


def test_ambiguous_send_reports_timeout_unknown() -> None:
    fake = FakeOutlook()
    fake.fault("send", "dropped")
    response = run(delivery(USER_CONNECTION, {"to": "ops@acme.example"}), fake)
    assert response["ok"] is False and response["delivered"] is False
    assert response["error_code"] == "timeout_unknown"
    assert response["external_effect_status"] == "timeout_unknown"
    assert response["definitely_no_external_effect"] is False
    assert response["retryable"] is False


def test_sensitive_body_is_never_in_the_response() -> None:
    fake = FakeOutlook()
    event = {
        "event_id": "evt-reset",
        "event_type": "system.password_reset",
        "payload_sensitive": True,
        "payload": {"message": f"Reset your password: {RESET_LINK}"},
    }
    response = run(
        delivery(USER_CONNECTION, {"to": "user@acme.example", "subject": "Reset"}, event=event),
        fake,
    )
    assert response["delivered"] is True
    assert "SECRET-RESET-TOKEN" not in repr(response)
    assert "SECRET-RESET-TOKEN" in fake.sent()[0].body["content"]


def test_dispatcher_routes_the_delivery_request(monkeypatch: pytest.MonkeyPatch) -> None:
    seen: list[dict[str, Any]] = []

    def fake_handler(payload: dict[str, Any]) -> dict[str, Any]:
        seen.append(payload)
        return {"ok": True, "delivered": True}

    monkeypatch.setattr(dispatcher, "handle_deliver_notification", fake_handler)
    payload = delivery(USER_CONNECTION, {"to": "ops@acme.example"})
    assert dispatcher.dispatch_runtime(payload) == {"ok": True, "delivered": True}
    assert seen == [payload]


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("https://fs.example/jobs/42", "https://fs.example/jobs/42"),
        ("http://localhost:8080/jobs/1", "http://localhost:8080/jobs/1"),
        ("javascript:alert(1)", ""),
        ("https://user:pw@fs.example/jobs/1", ""),
        ("/jobs/1", ""),
        (None, ""),
    ],
)
def test_run_url_is_only_a_plain_http_link(value: Any, expected: str) -> None:
    assert safe_run_url(value) == expected


def test_channel_content_uses_title_severity_and_link() -> None:
    content = event_content(
        {
            "event_type": "approval.requested",
            "severity": "warning",
            "payload": {"workflow_display_name": "Payments", "message": "Approve batch 7"},
            "run_url": "https://fs.example/jobs/7",
        }
    )
    assert content.title == "Approval requested: Payments"
    assert content.severity == "warning"
    assert content.summary == "Approve batch 7"
    assert content.run_url == "https://fs.example/jobs/7"


TEAMS_WEBHOOK = {
    "connection_id": "conn-hook",
    "connection_type_id": "m365_teams_webhook",
    "config": {"channel_label": "Operations"},
    "secrets": {"webhook_url": "https://x.environment.api.powerplatform.com/hook?sig=S3CR3T"},
}


def test_webhook_posts_to_its_own_channel() -> None:
    seen: list[str] = []

    def poster(url: str, **_kwargs: Any) -> dict[str, Any]:
        seen.append(url)
        return {"delivered": True}

    response = run(delivery(TEAMS_WEBHOOK, {}), None, webhook_poster=poster)
    assert response["delivered"] is True
    assert response["result"]["confirmation"] == "accepted_by_webhook"
    assert response["provider_message_id"] == ""
    assert seen == [TEAMS_WEBHOOK["secrets"]["webhook_url"]]
    assert "S3CR3T" not in repr(response)


def test_webhook_route_naming_a_channel_is_refused() -> None:
    response = run(
        delivery(TEAMS_WEBHOOK, {"channel": "team-1/other"}),
        None,
        webhook_poster=lambda *_a, **_k: {"delivered": True},
    )
    assert response["error_code"] == "notification_connection_incompatible"
    assert response["definitely_no_external_effect"] is True
