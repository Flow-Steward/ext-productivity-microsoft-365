from __future__ import annotations

import pytest
from fakes import FakeOpener, http_error

from m365.errors import M365Error, error_from_exception
from m365.transport import GraphTransport, is_transfer_host, map_http_error


def _transport(opener: FakeOpener, **kwargs) -> GraphTransport:
    return GraphTransport("token-value", opener=opener, sleep=lambda _s: None, **kwargs)


def test_get_json_prefixes_graph_base_and_sends_bearer() -> None:
    opener = FakeOpener([(200, {"id": "1"})])
    assert _transport(opener).get_json("/me", purpose="me") == {"id": "1"}
    request = opener.requests[0]
    assert request.full_url == "https://graph.microsoft.com/v1.0/me"
    assert request.get_header("Authorization") == "Bearer token-value"


def test_read_retries_on_503_with_retry_after() -> None:
    opener = FakeOpener(
        [
            http_error("u", 503, {"error": {"code": "serviceNotAvailable"}}, {"Retry-After": "1"}),
            (200, {"value": []}),
        ]
    )
    assert _transport(opener).get_json("/me/messages", purpose="list") == {"value": []}
    assert len(opener.requests) == 2


def test_write_with_retry_none_is_never_repeated_and_is_ambiguous() -> None:
    opener = FakeOpener([http_error("u", 503, {"error": {"code": "x"}})])
    with pytest.raises(M365Error) as caught:
        _transport(opener).request(
            "POST", "/me/sendMail", json_body={}, retry="none", purpose="send"
        )
    assert len(opener.requests) == 1
    assert caught.value.code == "timeout_unknown"
    assert caught.value.ambiguous is True
    envelope = error_from_exception(caught.value, effectful=True)
    assert envelope["external_effect_status"] == "timeout_unknown"
    assert envelope["definitely_no_external_effect"] is False
    assert envelope["retryable"] is False


def test_idempotent_write_is_retried() -> None:
    opener = FakeOpener(
        [http_error("u", 429, {"error": {"code": "TooManyRequests"}}), (200, {"id": "m"})]
    )
    response = _transport(opener).request(
        "PATCH", "/me/messages/m", json_body={"isRead": True}, retry="idempotent", purpose="flags"
    )
    assert response.json() == {"id": "m"}
    assert len(opener.requests) == 2


def test_read_retry_mode_rejects_writes() -> None:
    with pytest.raises(ValueError):
        _transport(FakeOpener([])).request("POST", "/x", retry="read", purpose="bad")


@pytest.mark.parametrize(
    "url",
    [
        "https://evil.example.com/v1.0/me",
        "https://graph.microsoft.com/beta/me",
    ],
)
def test_absolute_links_must_be_graph_v1(url: str) -> None:
    with pytest.raises(M365Error) as caught:
        _transport(FakeOpener([])).get_json(url, purpose="next")
    assert caught.value.code == "invalid_cursor"


def test_transfer_url_has_no_authorization_header() -> None:
    opener = FakeOpener([(200, b"bytes")])
    _transport(opener).request(
        "GET",
        "https://acme.sharepoint.com/download?x=1",
        retry="read",
        purpose="dl",
        authorized=False,
    )
    assert opener.requests[0].get_header("Authorization") is None


def test_transfer_url_on_other_host_is_refused() -> None:
    with pytest.raises(M365Error) as caught:
        _transport(FakeOpener([])).request(
            "GET", "https://files.example.net/x", retry="read", purpose="dl", authorized=False
        )
    assert caught.value.code == "egress_host_not_allowed"


@pytest.mark.parametrize(
    ("host", "allowed"),
    [
        ("acme.sharepoint.com", True),
        ("acme-my.sharepoint.com", True),
        ("outlook.office.com", True),
        ("sharepoint.com", False),
        ("acme.sharepoint.com.evil.net", False),
    ],
)
def test_transfer_hosts(host: str, allowed: bool) -> None:
    assert is_transfer_host(host) is allowed


@pytest.mark.parametrize(
    ("status", "body", "code"),
    [
        (401, {"error": {"code": "InvalidAuthenticationToken"}}, "authorization_required"),
        (403, {"error": {"code": "ErrorAccessDenied"}}, "mailbox_access_denied"),
        (403, {"error": {"code": "Authorization_RequestDenied"}}, "permission_denied"),
        (404, {"error": {"code": "ErrorInvalidUser"}}, "mailbox_not_found"),
        (404, {"error": {"code": "itemNotFound"}}, "item_not_found"),
        (409, {"error": {"code": "nameAlreadyExists"}}, "name_conflict"),
        (410, {"error": {"code": "SyncStateNotFound"}}, "sync_reset"),
        (429, {"error": {"code": "TooManyRequests"}}, "rate_limited"),
        (
            400,
            {"error": "invalid_client", "error_description": "AADSTS7000222: expired"},
            "client_secret_expired",
        ),
        (
            400,
            {"error": "invalid_request", "error_description": "AADSTS50194: single tenant"},
            "tenant_not_multitenant",
        ),
    ],
)
def test_error_mapping(status: int, body: dict, code: str) -> None:
    import json

    assert (
        map_http_error(status, json.dumps(body).encode(), method="GET", retry_after=None).code
        == code
    )


def test_tokens_never_appear_in_errors() -> None:
    opener = FakeOpener([http_error("u", 500, {"error": {"code": "x", "message": "token-value"}})])
    with pytest.raises(M365Error) as caught:
        _transport(opener, deadline_seconds=5).get_json("/me", purpose="me")
    assert "token-value" not in caught.value.message
