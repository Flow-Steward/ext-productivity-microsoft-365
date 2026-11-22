from __future__ import annotations

from fakes import FakeOpener

from identity import handle_verify_identity
from m365.errors import M365Error
from m365.runtime import Operation, run_operation


def _identity_payload(secrets: dict) -> dict:
    return {"action": {"target": {"connection": {"secrets": secrets}}}}


def test_identity_uses_me_and_never_needs_token_claims() -> None:
    opener = FakeOpener(
        [
            (
                200,
                {
                    "id": "user-1",
                    "userPrincipalName": "ana@acme.example",
                    "mail": "ana@acme.example",
                    "displayName": "Ana",
                },
            )
        ]
    )
    response = handle_verify_identity(_identity_payload({"access_token": "opaque"}), opener=opener)
    assert response["ok"] is True
    assert response["result"]["subject"] == "user-1"
    assert response["result"]["verified_email"] == "ana@acme.example"
    assert "tenant" not in str(response["result"]).lower()


def test_identity_requires_exactly_one_token() -> None:
    response = handle_verify_identity(
        _identity_payload({"access_token": "a", "refresh_token": "b"})
    )
    assert response["ok"] is False


def _operation_payload(*, test_mode: bool = False, operation: str = "do_write") -> dict:
    return {
        "mode": "action",
        "runtime_context": {"test_mode": test_mode},
        "action": {
            "action_id": operation,
            "input": {"connection_ref": "conn-1"},
            "target": {
                "connection": {
                    "connection_id": "conn-1",
                    "connection_type_id": "m365_user",
                    "secrets": {"access_token": "t"},
                }
            },
        },
    }


def test_effectful_operation_is_suppressed_in_test_mode() -> None:
    called = []
    registry = {
        "do_write": Operation(
            "do_write", "mail", lambda ctx: called.append(1) or {}, effectful=True
        )
    }
    response = run_operation(
        _operation_payload(test_mode=True), registry, transport_factory=lambda c: None
    )
    assert response["external_effect_status"] == "suppressed"
    assert called == []


def test_unknown_input_field_is_rejected() -> None:
    registry = {"do_write": Operation("do_write", "mail", lambda ctx: {}, effectful=True)}
    payload = _operation_payload()
    payload["action"]["input"]["extra"] = 1
    response = run_operation(payload, registry, transport_factory=lambda c: None)
    assert response["error_code"] == "invalid_payload"
    assert response["definitely_no_external_effect"] is True


def test_attachment_state_unknown_is_failed_and_retryable() -> None:
    def handler(ctx):
        raise M365Error(
            "attachment_state_unknown", "unknown", retryable=True, failure_class="provider"
        )

    registry = {"do_write": Operation("do_write", "mail", handler, effectful=True)}
    response = run_operation(_operation_payload(), registry, transport_factory=lambda c: None)
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True
    assert response["retryable"] is True
    assert response["failure_class"] == "provider"
