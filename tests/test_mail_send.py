"""Draft-then-send (B5.1, B8): keyed drafts, lookup after an ambiguous create,
attachment reconciliation, the unknown-attachment rule with its resume path, and
confirmed / unconfirmed ambiguous sends."""

from __future__ import annotations

import json
from typing import Any

import pytest
from fakes import FakeOpener, http_error
from test_mail_fakes import (
    STEP_EXECUTION_ID,
    FakeOutlook,
    action_payload,
    input_artifacts,
    sha,
    streamer_for,
)

from m365.runtime import run_operation
from m365.transport import GraphTransport
from mail import outgoing, send
from mail.operations import OPERATIONS
from mail.send import HEADER_NAME, HEADER_PROPERTY_ID, KEY_PROPERTY_ID, STATE_PROPERTY_ID

KEY = f"send_message:{STEP_EXECUTION_ID}"
PDF = b"%PDF-1.7 invoice 42"
CSV = b"sku,qty\nA-1,3\n"


@pytest.fixture(autouse=True)
def _no_waiting(monkeypatch: pytest.MonkeyPatch) -> list[float]:
    now = [0.0]

    def fake_sleep(seconds: float) -> None:
        now[0] += seconds

    monkeypatch.setattr(send, "sleep", fake_sleep)
    monkeypatch.setattr(send, "clock", lambda: now[0])
    return now


def _run(
    fake: FakeOutlook,
    action_id: str,
    operation_input: dict[str, Any],
    *,
    monkeypatch: pytest.MonkeyPatch | None = None,
    files: dict[str, tuple[str, str, bytes]] | None = None,
    **kwargs: Any,
) -> dict[str, Any]:
    artifacts = None
    if files:
        assert monkeypatch is not None
        monkeypatch.setattr(outgoing, "artifact_streamer", streamer_for(files))
        artifacts = input_artifacts(files)
    payload = action_payload(action_id, operation_input, artifacts=artifacts, **kwargs)
    return run_operation(payload, OPERATIONS, transport_factory=lambda _connection: fake)


def _send_input(**extra: Any) -> dict[str, Any]:
    return {
        "to": ["buyer@acme.example"],
        "subject": "Invoice 42",
        "body_text": "Attached.",
        **extra,
    }


def _creates(fake: FakeOutlook) -> list[Any]:
    return [
        call
        for call in fake.calls
        if call.method == "POST"
        and call.url.endswith(("/messages", "/createReply", "/createReplyAll", "/createForward"))
    ]


def _sends(fake: FakeOutlook) -> list[Any]:
    return [call for call in fake.calls if call.method == "POST" and call.url.endswith("/send")]


def _assert_no_transport_retry(fake: FakeOutlook) -> None:
    for call in fake.writes():
        assert call.retry == "none", f"{call.method} {call.url} may be retried by the transport"


# -- the plain path -------------------------------------------------------------------


def test_send_creates_a_keyed_draft_then_sends_it() -> None:
    fake = FakeOutlook()
    response = _run(
        fake, "send_message", _send_input(cc=[{"address": "ops@acme.example", "name": "Ops"}])
    )

    assert response["ok"] is True, response
    assert response["external_effect_status"] == "succeeded"
    assert response["definitely_no_external_effect"] is False
    assert response["result"]["confirmation"] == "send_accepted"
    assert len(fake.sent_log) == 1
    sent = fake.messages[fake.sent_log[0]]
    assert sent.props[KEY_PROPERTY_ID] == KEY
    assert sent.props[f"header:{HEADER_NAME}"] == KEY
    assert response["result"]["message_id"] == sent.id
    create = _creates(fake)[0]
    assert create.json_body["singleValueExtendedProperties"] == [
        {"id": KEY_PROPERTY_ID, "value": KEY}
    ]
    assert create.json_body["internetMessageHeaders"] == [{"name": HEADER_NAME, "value": KEY}]
    assert create.json_body["body"] == {"contentType": "Text", "content": "Attached."}
    _assert_no_transport_retry(fake)
    for call in fake.calls:
        assert 'IdType="ImmutableId"' in call.headers.get("Prefer", "")


def test_property_set_guid_is_fixed() -> None:
    assert (
        KEY_PROPERTY_ID
        == "String {b9fd5293-1be4-448e-b08f-b4467c228ad8} Name FlowStewardIdempotencyKey"
    )
    assert (
        STATE_PROPERTY_ID
        == "String {b9fd5293-1be4-448e-b08f-b4467c228ad8} Name FlowStewardDraftState"
    )


def test_test_run_suppresses_every_send() -> None:
    fake = FakeOutlook()
    for action_id, operation_input in (
        ("send_message", _send_input()),
        ("reply_message", {"message_id": "AAMk-MSG-1"}),
        ("forward_message", {"message_id": "AAMk-MSG-1", "to": ["a@acme.example"]}),
    ):
        response = _run(fake, action_id, operation_input, test_mode=True)
        assert response["external_effect_status"] == "suppressed"
        assert response["definitely_no_external_effect"] is True
    assert fake.calls == []


@pytest.mark.parametrize(
    ("operation_input", "message"),
    [
        ({"subject": "no recipient"}, "recipient"),
        (_send_input(body_html="<p>x</p>"), "not both"),
        (_send_input(to=["not-an-address"]), "invalid email"),
        (_send_input(importance="urgent"), "importance"),
        (_send_input(attachments=[{"artifact_handle": "a1", "name": "x.pdf"}]), "artifact handle"),
        (
            _send_input(attachments=[{"artifact_handle": "a1", "filename": "../x.pdf"}]),
            "safe file name",
        ),
    ],
)
def test_send_input_is_validated_before_any_call(
    operation_input: dict[str, Any], message: str
) -> None:
    fake = FakeOutlook()
    response = _run(fake, "send_message", operation_input)
    assert response["ok"] is False
    assert response["error_code"] == "invalid_payload"
    assert message in response["error"]
    assert response["definitely_no_external_effect"] is True
    assert fake.calls == []


def test_unknown_send_field_is_rejected() -> None:
    response = _run(FakeOutlook(), "send_message", _send_input(save_to_sent_items=False))
    assert response["error_code"] == "invalid_payload"


# -- step 1: lookup after an ambiguous create -------------------------------------------


def test_ambiguous_create_with_no_draft_found_retries_the_create_once() -> None:
    fake = FakeOutlook()
    fake.fault("create", "dropped")
    response = _run(fake, "send_message", _send_input())

    assert response["ok"] is True, response
    assert len(_creates(fake)) == 2
    assert len(fake.sent_log) == 1
    assert fake.drafts() == []
    _assert_no_transport_retry(fake)


def test_ambiguous_create_with_one_draft_found_uses_it() -> None:
    fake = FakeOutlook()
    fake.fault("create", "lost")
    response = _run(fake, "send_message", _send_input())

    assert response["ok"] is True, response
    assert len(_creates(fake)) == 1
    assert len(fake.sent_log) == 1


def test_ambiguous_create_with_several_drafts_keeps_the_oldest() -> None:
    fake = FakeOutlook()
    fake.fault("create", "lost_twice")
    response = _run(fake, "send_message", _send_input())

    assert response["ok"] is True, response
    assert len(_creates(fake)) == 1
    assert fake.sent_log == ["AAMk-DRAFT-0001"]
    assert fake.drafts() == []


def test_a_create_that_lands_late_is_removed_before_the_send() -> None:
    fake = FakeOutlook()
    fake.fault("create", "late")
    creates = []

    def land_on_second_create(call: Any) -> None:
        if call in _creates(fake) and not creates:
            creates.append(call)
        elif call.method == "POST" and call.url == "/me/messages":
            fake.land_late()

    fake.hook = land_on_second_create
    response = _run(fake, "send_message", _send_input())

    assert response["ok"] is True, response
    assert len(fake.sent_log) == 1
    assert fake.drafts() == []  # the late draft was deleted, not sent


def test_two_ambiguous_creates_without_a_draft_fail_retryably_and_send_nothing() -> None:
    fake = FakeOutlook()
    fake.fault("create", "dropped", "dropped")
    response = _run(fake, "send_message", _send_input())

    assert response["ok"] is False
    assert response["error_code"] == "provider_unavailable"
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True
    assert response["retryable"] is True
    assert response["failure_class"] == "provider"
    assert fake.sent_log == []


def test_confirmed_create_failure_sends_nothing() -> None:
    fake = FakeOutlook()
    fake.fault("create", "refused")
    response = _run(fake, "send_message", _send_input())

    assert response["ok"] is False
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True
    assert fake.sent_log == [] and fake.drafts() == []


# -- step 2: attachments and reconciliation ----------------------------------------------


def test_inline_attachments_are_added_and_reconciled(monkeypatch: pytest.MonkeyPatch) -> None:
    fake = FakeOutlook()
    files = {
        "art-1": ("invoice.pdf", "application/pdf", PDF),
        "art-2": ("stock.csv", "text/csv", CSV),
    }
    response = _run(
        fake,
        "send_message",
        _send_input(attachments=["art-1", {"artifact_handle": "art-2", "filename": "lines.csv"}]),
        monkeypatch=monkeypatch,
        files=files,
    )

    assert response["ok"] is True, response
    assert response["result"]["attachment_count"] == 2
    sent = fake.messages[fake.sent_log[0]]
    assert sorted((a["name"], sha(a["content"])) for a in sent.attachments) == sorted(
        [("invoice.pdf", sha(PDF)), ("lines.csv", sha(CSV))]
    )
    _assert_no_transport_retry(fake)


def test_reconciliation_removes_a_duplicate(monkeypatch: pytest.MonkeyPatch) -> None:
    fake = FakeOutlook()
    fake.fault("attach", "duplicated")
    files = {"art-1": ("invoice.pdf", "application/pdf", PDF)}
    response = _run(
        fake,
        "send_message",
        _send_input(attachments=["art-1"]),
        monkeypatch=monkeypatch,
        files=files,
    )

    assert response["ok"] is True, response
    sent = fake.messages[fake.sent_log[0]]
    assert [a["content"] for a in sent.attachments] == [PDF]
    assert any(call.method == "DELETE" and "/attachments/" in call.url for call in fake.calls)


def test_reconciliation_adds_a_missing_attachment(monkeypatch: pytest.MonkeyPatch) -> None:
    fake = FakeOutlook()
    fake.fault("attach", "phantom")
    files = {"art-1": ("invoice.pdf", "application/pdf", PDF)}
    response = _run(
        fake,
        "send_message",
        _send_input(attachments=["art-1"]),
        monkeypatch=monkeypatch,
        files=files,
    )

    assert response["ok"] is True, response
    sent = fake.messages[fake.sent_log[0]]
    assert [a["content"] for a in sent.attachments] == [PDF]
    adds = [c for c in fake.calls if c.method == "POST" and c.url.endswith("/attachments")]
    assert len(adds) == 2


def test_same_name_files_are_told_apart_by_content(monkeypatch: pytest.MonkeyPatch) -> None:
    fake = FakeOutlook()
    other = b"%PDF-1.7 invoice 43"
    files = {
        "art-1": ("invoice.pdf", "application/pdf", PDF),
        "art-2": ("invoice.pdf", "application/pdf", other),
    }
    fake.fault("attach", "duplicated")  # the first file lands twice under the same name
    response = _run(
        fake,
        "send_message",
        _send_input(attachments=["art-1", "art-2"]),
        monkeypatch=monkeypatch,
        files=files,
    )

    assert response["ok"] is True, response
    sent = fake.messages[fake.sent_log[0]]
    assert sorted(sha(a["content"]) for a in sent.attachments) == sorted([sha(PDF), sha(other)])
    assert any(call.url.endswith("/$value") for call in fake.calls), "content was not compared"


def test_an_ambiguous_add_that_landed_is_resolved_and_sent(monkeypatch: pytest.MonkeyPatch) -> None:
    fake = FakeOutlook()
    fake.fault("attach", "lost")
    files = {"art-1": ("invoice.pdf", "application/pdf", PDF)}
    response = _run(
        fake,
        "send_message",
        _send_input(attachments=["art-1"]),
        monkeypatch=monkeypatch,
        files=files,
    )

    assert response["ok"] is True, response
    assert [a["content"] for a in fake.messages[fake.sent_log[0]].attachments] == [PDF]


def test_forward_keeps_the_original_attachments_out_of_reconciliation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake = FakeOutlook()
    original = fake.add_message(subject="CMR 7")
    fake.attach(original, "cmr.pdf", "application/pdf", b"%PDF-cmr")
    files = {"art-1": ("invoice.pdf", "application/pdf", PDF)}
    response = _run(
        fake,
        "forward_message",
        {
            "message_id": original.id,
            "to": ["carrier@acme.example"],
            "comment": "See attached",
            "attachments": ["art-1"],
        },
        monkeypatch=monkeypatch,
        files=files,
    )

    assert response["ok"] is True, response
    sent = fake.messages[fake.sent_log[0]]
    assert sorted(a["name"] for a in sent.attachments) == ["cmr.pdf", "invoice.pdf"]
    create = _creates(fake)[0]
    assert create.url.endswith("/createForward")
    assert create.json_body["comment"] == "See attached"
    assert create.json_body["message"]["toRecipients"] == [
        {"emailAddress": {"address": "carrier@acme.example"}}
    ]
    props = create.json_body["message"]["singleValueExtendedProperties"]
    assert {"id": KEY_PROPERTY_ID, "value": f"forward_message:{STEP_EXECUTION_ID}"} in props
    assert {"id": HEADER_PROPERTY_ID, "value": f"forward_message:{STEP_EXECUTION_ID}"} in props


def test_reply_uses_create_reply_all_with_the_key() -> None:
    fake = FakeOutlook()
    original = fake.add_message(subject="Invoice")
    response = _run(
        fake, "reply_message", {"message_id": original.id, "reply_all": True, "comment": "Thanks"}
    )

    assert response["ok"] is True, response
    create = _creates(fake)[0]
    assert create.url == f"/me/messages/{original.id}/createReplyAll"
    assert create.json_body["comment"] == "Thanks"
    assert {
        "id": KEY_PROPERTY_ID,
        "value": f"reply_message:{STEP_EXECUTION_ID}",
    } in create.json_body["message"]["singleValueExtendedProperties"]


def test_reply_to_a_missing_message_is_message_not_found() -> None:
    fake = FakeOutlook()
    response = _run(fake, "reply_message", {"message_id": "AAMk-MISSING"})
    assert response["error_code"] == "message_not_found"
    assert response["definitely_no_external_effect"] is True


# -- the unknown-attachment rule and its resume path --------------------------------------


def test_unknown_attachment_add_abandons_the_draft_and_signals_a_retry(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake = FakeOutlook()
    fake.fault("attach", "late")
    files = {"art-1": ("invoice.pdf", "application/pdf", PDF)}
    response = _run(
        fake,
        "send_message",
        _send_input(attachments=["art-1"]),
        monkeypatch=monkeypatch,
        files=files,
    )

    assert response["ok"] is False
    assert response["error_code"] == "attachment_state_unknown"
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True
    assert response["retryable"] is True
    assert response["failure_class"] == "provider"
    assert "AAMk-DRAFT-0001" in response["error"]
    assert fake.sent_log == []
    patch = [c for c in fake.calls if c.method == "PATCH"][0]
    assert patch.json_body == {
        "singleValueExtendedProperties": [{"id": STATE_PROPERTY_ID, "value": "abandoned"}]
    }
    assert fake.drafts() == []

    fake.land_late()  # the in-flight add lands on a deleted draft and fails harmlessly
    retried = _run(
        fake,
        "send_message",
        _send_input(attachments=["art-1"]),
        monkeypatch=monkeypatch,
        files=files,
    )

    assert retried["ok"] is True, retried
    assert len(fake.sent_log) == 1
    assert fake.sent_log[0] != "AAMk-DRAFT-0001"
    assert [a["content"] for a in fake.messages[fake.sent_log[0]].attachments] == [PDF]


def test_retried_step_deletes_an_abandoned_draft_it_could_not_delete_before(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake = FakeOutlook()
    fake.fault("attach", "late")
    fake.fault("delete_message", "refused")
    files = {"art-1": ("invoice.pdf", "application/pdf", PDF)}
    first = _run(
        fake,
        "send_message",
        _send_input(attachments=["art-1"]),
        monkeypatch=monkeypatch,
        files=files,
    )

    assert first["error_code"] == "attachment_state_unknown"
    assert "stays in Drafts" in first["error"]
    [abandoned] = fake.drafts()
    assert abandoned.props[STATE_PROPERTY_ID] == "abandoned"
    fake.land_late()  # the late add lands on the abandoned draft: it must still never be sent

    retried = _run(
        fake,
        "send_message",
        _send_input(attachments=["art-1"]),
        monkeypatch=monkeypatch,
        files=files,
    )

    assert retried["ok"] is True, retried
    assert abandoned.id not in fake.messages
    assert len(fake.sent_log) == 1 and fake.sent_log[0] != abandoned.id


def test_retried_step_fails_again_when_the_abandoned_draft_cannot_be_deleted(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake = FakeOutlook()
    fake.fault("attach", "late")
    fake.fault("delete_message", "refused", "refused")
    files = {"art-1": ("invoice.pdf", "application/pdf", PDF)}
    _run(
        fake,
        "send_message",
        _send_input(attachments=["art-1"]),
        monkeypatch=monkeypatch,
        files=files,
    )
    retried = _run(
        fake,
        "send_message",
        _send_input(attachments=["art-1"]),
        monkeypatch=monkeypatch,
        files=files,
    )

    assert retried["error_code"] == "attachment_state_unknown"
    assert retried["external_effect_status"] == "failed"
    assert retried["retryable"] is True and retried["failure_class"] == "provider"
    assert fake.sent_log == []


def test_restart_job_uses_a_new_key_and_a_fresh_draft(monkeypatch: pytest.MonkeyPatch) -> None:
    fake = FakeOutlook()
    fake.fault("attach", "late")
    fake.fault("delete_message", "refused")
    files = {"art-1": ("invoice.pdf", "application/pdf", PDF)}
    _run(
        fake,
        "send_message",
        _send_input(attachments=["art-1"]),
        monkeypatch=monkeypatch,
        files=files,
    )
    [abandoned] = fake.drafts()

    restarted = _run(
        fake,
        "send_message",
        _send_input(attachments=["art-1"]),
        monkeypatch=monkeypatch,
        files=files,
        step_execution_id="6a1b2c3d-4e5f-4a6b-8c7d-000000000202",
    )

    assert restarted["ok"] is True, restarted
    assert len(fake.sent_log) == 1 and fake.sent_log[0] != abandoned.id
    assert abandoned.id in fake.messages  # another key: left in Drafts, marked abandoned


# -- large attachments through an upload session ---------------------------------------------


def test_large_attachment_uses_an_unauthenticated_upload_session(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake = FakeOutlook()
    big = bytes(range(256)) * (14 * 1024)  # 3.5 MiB
    files = {"art-big": ("scan.pdf", "application/pdf", big)}
    response = _run(
        fake,
        "send_message",
        _send_input(attachments=["art-big"]),
        monkeypatch=monkeypatch,
        files=files,
    )

    assert response["ok"] is True, response
    puts = [c for c in fake.calls if c.method == "PUT"]
    assert len(puts) == 2 and all(c.authorized is False for c in puts)
    assert puts[0].headers["Content-Range"] == f"bytes 0-{send.UPLOAD_CHUNK_BYTES - 1}/{len(big)}"
    assert [a["content"] for a in fake.messages[fake.sent_log[0]].attachments] == [big]
    assert "authtoken" not in json.dumps(response)
    _assert_no_transport_retry(fake)


def test_ambiguous_final_upload_range_is_an_unknown_attachment(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake = FakeOutlook()
    big = b"x" * (3 * 1024 * 1024 + 10)
    fake.fault("chunk_final", "late")
    files = {"art-big": ("scan.pdf", "application/pdf", big)}
    response = _run(
        fake,
        "send_message",
        _send_input(attachments=["art-big"]),
        monkeypatch=monkeypatch,
        files=files,
    )

    assert response["error_code"] == "attachment_state_unknown"
    assert response["definitely_no_external_effect"] is True
    assert fake.sent_log == []


def test_refused_upload_session_is_attachment_upload_failed(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    fake = FakeOutlook()
    fake.fault("session", "refused")
    files = {"art-big": ("scan.pdf", "application/pdf", b"y" * (3 * 1024 * 1024))}
    response = _run(
        fake,
        "send_message",
        _send_input(attachments=["art-big"]),
        monkeypatch=monkeypatch,
        files=files,
    )

    assert response["error_code"] == "attachment_upload_failed"
    assert "README" in response["error"]
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True
    assert fake.sent_log == [] and fake.drafts() == []
    assert "authtoken" not in json.dumps(response)


# -- step 3: ambiguous sends ---------------------------------------------------------------


def test_ambiguous_send_confirmed_by_the_immutable_id() -> None:
    fake = FakeOutlook()
    fake.fault("send", "lost")
    response = _run(fake, "send_message", _send_input())

    assert response["ok"] is True, response
    assert response["result"]["confirmation"] == "immutable_id"
    assert response["external_effect_status"] == "succeeded"
    assert len(_sends(fake)) == 1


def test_ambiguous_send_confirmed_by_a_sent_items_search() -> None:
    fake = FakeOutlook()
    fake.fault("send", "lost")
    fake.ids_change_on_send = True
    response = _run(fake, "send_message", _send_input())

    assert response["ok"] is True, response
    assert response["result"]["confirmation"] == "sent_items_search"
    assert response["result"]["message_id"] == fake.sent()[0].id
    assert len(_sends(fake)) == 1


@pytest.mark.parametrize("mode", ["dropped", "lost_but_draft_visible"])
def test_unconfirmed_send_is_timeout_unknown_and_never_resent(mode: str) -> None:
    fake = FakeOutlook()
    if mode == "dropped":
        fake.fault("send", "dropped")
    else:
        fake.fault("send", "lost")
        fake.send_invisible = True  # a visible draft is no proof that nothing went out
    response = _run(fake, "send_message", _send_input())

    assert response["ok"] is False
    assert response["error_code"] == "timeout_unknown"
    assert response["external_effect_status"] == "timeout_unknown"
    assert response["definitely_no_external_effect"] is False
    assert response["retryable"] is False
    assert "AAMk-DRAFT-0001" in response["error"]
    assert len(_sends(fake)) == 1


def test_confirmed_send_refusal_fails_cleanly_and_removes_the_draft() -> None:
    fake = FakeOutlook()
    fake.fault("send", "throttled")
    response = _run(fake, "send_message", _send_input())

    assert response["error_code"] == "rate_limited"
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True
    assert response["retryable"] is True
    assert fake.drafts() == []


def test_a_sent_copy_with_the_key_is_not_sent_again() -> None:
    fake = FakeOutlook()
    previous = fake.add_message("sentitems", props={KEY_PROPERTY_ID: KEY})
    response = _run(fake, "send_message", _send_input())

    assert response["ok"] is True
    assert response["result"] == {
        "message_id": previous.id,
        "confirmation": "already_sent",
        "attachment_count": 0,
        "external_effect_status": "succeeded",
        "definitely_no_external_effect": False,
    }
    assert _sends(fake) == [] and _creates(fake) == []


# -- B6 retry allow-list against the real transport ---------------------------------------


def test_real_transport_never_repeats_an_ambiguous_draft_create() -> None:
    opener = FakeOpener(
        [
            (200, {"value": []}),  # Sent Items: not sent yet
            (200, {"value": []}),  # Drafts: nothing from an earlier attempt
            http_error("u", 503, {"error": {"code": "ServiceUnavailable"}}),  # create: ambiguous
            (200, {"value": [{"id": "AAMk-D1", "createdDateTime": "2026-11-22T10:00:00Z"}]}),
            (200, {"value": [{"id": "AAMk-D1", "createdDateTime": "2026-11-22T10:00:00Z"}]}),
            (202, b""),  # send
        ]
    )
    transport = GraphTransport("opaque-token", opener=opener, sleep=lambda _s: None)
    payload = action_payload("send_message", _send_input())
    response = run_operation(payload, OPERATIONS, transport_factory=lambda _c: transport)

    assert response["ok"] is True, response
    methods = [
        (r.get_method(), r.full_url.split("?")[0].rsplit("/", 1)[-1]) for r in opener.requests
    ]
    assert methods == [
        ("GET", "messages"),
        ("GET", "messages"),
        ("POST", "messages"),
        ("GET", "messages"),
        ("GET", "messages"),
        ("POST", "send"),
    ]


def test_real_transport_never_repeats_a_send_after_503() -> None:
    draft = {"id": "AAMk-D1", "createdDateTime": "2026-11-22T10:00:00Z"}
    opener = FakeOpener(
        [
            (200, {"value": []}),
            (200, {"value": []}),
            (201, draft),
            (200, {"value": [draft]}),
            http_error("u", 503, {"error": {"code": "ServiceUnavailable"}}),  # send: ambiguous
            (200, {"id": "AAMk-D1", "isDraft": False}),  # confirmation by immutable id
        ]
    )
    transport = GraphTransport("opaque-token", opener=opener, sleep=lambda _s: None)
    response = run_operation(
        action_payload("send_message", _send_input()),
        OPERATIONS,
        transport_factory=lambda _c: transport,
    )

    assert response["ok"] is True, response
    assert response["result"]["confirmation"] == "immutable_id"
    assert sum(1 for r in opener.requests if r.full_url.endswith("/send")) == 1
    assert opener.requests[-1].get_header("Prefer").startswith('IdType="ImmutableId"')
