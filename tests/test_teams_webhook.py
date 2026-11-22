"""Teams Workflows webhook client (v0.2 D4) and the notification card builders.

The webhook URL is a secret: no error, result or traceback may contain it.
"""

from __future__ import annotations

import http.client
import json
import socket
import traceback
from pathlib import Path
from typing import Any
from urllib.error import URLError

import pytest
from fakes import FakeOpener, FakeResponse, http_error
from flowsteward_extension_sdk import PinnedPeerError, open_pinned_url
from test_teams_fakes import WEBHOOK_URL

from m365.connection import ConnectionContext, resolve_connection
from m365.errors import M365Error, error_from_exception
from teams import webhook
from teams.cards import notification_card, notification_html, safe_link, webhook_payload
from teams.webhook import (
    ALLOWED_WEBHOOK_HOST_SUFFIXES,
    check_webhook_url,
    is_allowed_webhook_host,
    post_webhook_card,
    post_webhook_for_connection,
)

SECRET_PARTS = ("TOPSECRETSIG", "abc123", "prod-42", "workflows", "sig=")
CARD = webhook_payload(
    notification_card(
        title="Job failed", severity="error", summary="Step 3 failed", link="https://fs.example/r/1"
    )
)


def _assert_no_secret(exc: BaseException) -> None:
    rendered = "".join(traceback.format_exception(exc))
    envelope = (
        json.dumps(error_from_exception(exc, effectful=True)) if isinstance(exc, M365Error) else ""
    )
    for text in (str(exc), repr(exc), rendered, envelope):
        assert WEBHOOK_URL not in text
        for part in SECRET_PARTS:
            assert part not in text, part


def _post(*script: Any) -> tuple[dict[str, Any] | M365Error, FakeOpener]:
    opener = FakeOpener(list(script))
    try:
        return post_webhook_card(WEBHOOK_URL, CARD, opener=opener), opener
    except M365Error as exc:
        _assert_no_secret(exc)
        return exc, opener


# -- URL and host checks ---------------------------------------------------------------------------


def test_allowed_suffixes_are_a_module_constant() -> None:
    assert ALLOWED_WEBHOOK_HOST_SUFFIXES == (
        ".environment.api.powerplatform.com",
        ".logic.azure.com",
    )


@pytest.mark.parametrize(
    "host",
    [
        "prod-42.westeurope.logic.azure.com",
        "default0a1b2c3d.4e.environment.api.powerplatform.com",
        "PROD-42.WestEurope.Logic.Azure.com",
    ],
)
def test_allowed_hosts(host: str) -> None:
    assert is_allowed_webhook_host(host)
    assert check_webhook_url(f"https://{host}/workflows/x/triggers/manual?sig=s")


@pytest.mark.parametrize(
    "host",
    [
        "logic.azure.com",
        ".logic.azure.com",
        "evil.example",
        "logic.azure.com.evil.example",
        "evillogic.azure.com",
        "prod.logic.azure.com.",
        "environment.api.powerplatform.com",
        "x.api.powerplatform.com",
        "10.0.0.1",
        "",
    ],
)
def test_denied_hosts(host: str) -> None:
    assert not is_allowed_webhook_host(host)


@pytest.mark.parametrize(
    ("url", "code"),
    [
        (
            "http://prod-42.westeurope.logic.azure.com/workflows/abc123?sig=TOPSECRETSIG",
            "invalid_connection",
        ),
        ("ftp://prod-42.westeurope.logic.azure.com/workflows/abc123", "invalid_connection"),
        (
            "https://user:pw@prod-42.westeurope.logic.azure.com/workflows/abc123",
            "invalid_connection",
        ),
        ("https://prod-42.westeurope.logic.azure.com:8443/workflows/abc123", "invalid_connection"),
        ("https://prod-42.westeurope.logic.azure.com:abc/workflows/abc123", "invalid_connection"),
        ("https://prod-42.westeurope.logic.azure.com/work flows/abc123", "invalid_connection"),
        ("https://evil.example/workflows/abc123?sig=TOPSECRETSIG", "egress_host_not_allowed"),
        (
            "https://prod-42.westeurope.logic.azure.com.evil.example/workflows/abc123",
            "egress_host_not_allowed",
        ),
        ("", "invalid_connection"),
    ],
)
def test_refused_urls_send_nothing_and_never_echo_the_url(url: str, code: str) -> None:
    opener = FakeOpener([])
    with pytest.raises(M365Error) as raised:
        post_webhook_card(url, CARD, opener=opener)
    assert raised.value.code == code
    assert opener.requests == []
    if url:
        rendered = "".join(traceback.format_exception(raised.value))
        assert url not in rendered
        assert "TOPSECRETSIG" not in rendered and "abc123" not in rendered


def test_the_default_opener_is_the_sdk_pinned_opener_without_redirects() -> None:
    assert post_webhook_card.__kwdefaults__["opener"] is open_pinned_url


# -- outcomes -----------------------------------------------------------------------------------------


@pytest.mark.parametrize("status", [200, 202, 204])
def test_2xx_is_accepted_by_the_webhook(status: int) -> None:
    result, opener = _post((status, b"") if status == 204 else (status, {"ok": True}))
    assert result == {
        "delivered": True,
        "confirmation": "accepted_by_webhook",
        "provider_message_id": "",
        "http_status": status,
        "external_effect_status": "succeeded",
        "definitely_no_external_effect": False,
    }
    [request] = opener.requests
    assert request.get_method() == "POST"
    assert request.full_url == WEBHOOK_URL
    assert request.get_header("Content-type") == "application/json"
    assert "Authorization" not in request.headers
    assert json.loads(request.data) == CARD


@pytest.mark.parametrize("status", [301, 302, 307, 308])
def test_a_redirect_is_refused(status: int) -> None:
    # Raised as HTTPError by a no-redirect opener, or returned as a plain response.
    for step in (
        http_error(WEBHOOK_URL, status, b"", {"Location": "https://evil.example/x"}),
        lambda _r, s=status: FakeResponse(s, b""),
    ):
        outcome, opener = _post(step)
        assert isinstance(outcome, M365Error)
        assert outcome.code == "webhook_rejected"
        assert outcome.ambiguous is False
        assert len(opener.requests) == 1, "the redirect must not be followed"


@pytest.mark.parametrize("status", [400, 401, 403, 404, 410])
def test_4xx_is_rejected_with_no_external_effect(status: int) -> None:
    outcome, opener = _post((status, {"error": {"message": WEBHOOK_URL}}))
    assert isinstance(outcome, M365Error)
    assert outcome.code == "webhook_rejected"
    envelope = error_from_exception(outcome, effectful=True)
    assert envelope["definitely_no_external_effect"] is True
    assert envelope["external_effect_status"] == "failed"
    assert envelope["retryable"] is False
    assert len(opener.requests) == 1


def test_429_is_rejected_but_retryable() -> None:
    error = http_error(WEBHOOK_URL, 429, b"", {"Retry-After": "12"})
    outcome, _opener = _post(error)
    assert isinstance(outcome, M365Error)
    assert outcome.code == "webhook_rejected"
    envelope = error_from_exception(outcome, effectful=True)
    assert envelope["definitely_no_external_effect"] is True
    assert envelope["retryable"] is True
    assert envelope["retry_after_seconds"] == 12.0


@pytest.mark.parametrize(
    "step",
    [
        (500, b""),
        (502, b""),
        (503, b""),
        (504, b""),
        TimeoutError("timed out"),
        socket.timeout("timed out"),
        ConnectionResetError(104, "Connection reset by peer"),
        http.client.RemoteDisconnected("Remote end closed connection without response"),
        http.client.IncompleteRead(b"partial"),
        URLError(TimeoutError("timed out")),
        URLError(ConnectionResetError(104, "reset")),
    ],
)
def test_a_lost_response_is_timeout_unknown_and_never_retried(step: Any) -> None:
    outcome, opener = _post(step)
    assert isinstance(outcome, M365Error)
    assert outcome.code == "timeout_unknown"
    assert outcome.ambiguous is True
    envelope = error_from_exception(outcome, effectful=True)
    assert envelope["external_effect_status"] == "timeout_unknown"
    assert envelope["definitely_no_external_effect"] is False
    assert envelope["retryable"] is False
    assert len(opener.requests) == 1


@pytest.mark.parametrize(
    "step",
    [
        PinnedPeerError(f"Teams webhook: address of {WEBHOOK_URL} not allowed"),
        URLError(ConnectionRefusedError(111, "Connection refused")),
        URLError(socket.gaierror(-2, "Name or service not known")),
    ],
)
def test_a_request_that_never_left_is_a_network_failure(step: Any) -> None:
    outcome, _opener = _post(step)
    assert isinstance(outcome, M365Error)
    assert outcome.code == "network_failure"
    assert outcome.ambiguous is False
    assert error_from_exception(outcome, effectful=True)["definitely_no_external_effect"] is True


def test_errors_carrying_the_url_are_not_chained() -> None:
    outcome, _opener = _post(URLError(f"bad {WEBHOOK_URL}"))
    assert isinstance(outcome, M365Error)
    assert outcome.__cause__ is None
    assert outcome.__context__ is None


def test_an_oversized_card_is_refused_before_sending() -> None:
    opener = FakeOpener([])
    big = webhook_payload({"type": "AdaptiveCard", "body": [{"text": "x" * 40_000}]})
    with pytest.raises(M365Error) as raised:
        post_webhook_card(WEBHOOK_URL, big, opener=opener)
    assert raised.value.code == "message_too_large"
    assert opener.requests == []


# -- connection --------------------------------------------------------------------------------------


def test_the_webhook_connection_keeps_the_url_out_of_its_repr() -> None:
    payload = {
        "action": {
            "target": {
                "connection": {
                    "connection_id": "conn-hook",
                    "connection_type_id": "m365_teams_webhook",
                    "config": {"channel_label": "Operations"},
                    "secrets": {"webhook_url": WEBHOOK_URL},
                }
            }
        }
    }
    connection = resolve_connection(payload, service="teams")
    assert connection.mode == "webhook"
    assert "TOPSECRETSIG" not in repr(connection)

    opener = FakeOpener([(202, b"")])
    result = post_webhook_for_connection(connection, CARD, opener=opener)
    assert result["confirmation"] == "accepted_by_webhook"


def test_a_user_connection_cannot_post_to_a_webhook() -> None:
    connection = ConnectionContext("c", "m365_user", "teams", "user", access_token="t")
    with pytest.raises(M365Error) as raised:
        post_webhook_for_connection(connection, CARD, opener=FakeOpener([]))
    assert raised.value.code == "notification_connection_incompatible"


def test_the_module_never_logs() -> None:
    source = Path(webhook.__file__).read_text(encoding="utf-8")
    assert "logging" not in source and "print(" not in source


# -- cards ---------------------------------------------------------------------------------------------


def test_card_has_title_severity_summary_and_an_open_button() -> None:
    card = notification_card(
        title="Job failed", severity="Error", summary="Step 3 failed", link="https://fs.example/r/1"
    )
    assert card["type"] == "AdaptiveCard"
    assert card["version"] == "1.4"
    texts = [block["text"] for block in card["body"]]
    assert texts == ["Job failed", "ERROR", "Step 3 failed"]
    assert card["body"][1]["color"] == "Attention"
    assert card["actions"] == [
        {"type": "Action.OpenUrl", "title": "Open in Flow Steward", "url": "https://fs.example/r/1"}
    ]
    envelope = webhook_payload(card)
    assert envelope["type"] == "message"
    assert envelope["attachments"][0]["contentType"] == "application/vnd.microsoft.card.adaptive"
    assert envelope["attachments"][0]["content"] is card


def test_card_drops_unsafe_links_and_unknown_severities_render_default() -> None:
    card = notification_card(title="t", severity="weird", summary="s", link="javascript:alert(1)")
    assert "actions" not in card
    assert card["body"][1]["color"] == "Default"
    assert safe_link("http://flowsteward.localhost:8080/runs/1") == (
        "http://flowsteward.localhost:8080/runs/1"
    )
    assert safe_link("https://user:pw@fs.example/") == ""


def test_html_notification_escapes_every_field() -> None:
    html_body = notification_html(
        title="<b>x</b>",
        severity="warning",
        summary="a & b\nc",
        link='https://fs.example/r?a=1&b="2"',
    )
    assert "<b>x</b>" not in html_body
    assert "&lt;b&gt;x&lt;/b&gt;" in html_body
    assert "a &amp; b<br>c" in html_body
    assert 'href="https://fs.example/r?a=1&amp;b=&quot;2&quot;"' in html_body
    assert "WARNING" in html_body
