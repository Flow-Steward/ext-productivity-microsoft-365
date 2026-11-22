from __future__ import annotations

from urllib.parse import parse_qs

import pytest
from fakes import FakeOpener, http_error

from m365.auth import app_token
from m365.connection import mailbox_base, resolve_connection
from m365.errors import M365Error

TENANT = "11111111-2222-3333-4444-555555555555"
CLIENT = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"


def _payload(type_id: str, *, config: dict | None = None, secrets: dict | None = None) -> dict:
    return {
        "action": {
            "target": {
                "connection": {
                    "connection_id": "conn-1",
                    "connection_type_id": type_id,
                    "config": config or {},
                    "secrets": secrets or {},
                }
            }
        }
    }


def test_user_mail_connection_defaults_to_me() -> None:
    context = resolve_connection(
        _payload("m365_mail_user", secrets={"access_token": "t"}), service="mail"
    )
    assert context.mode == "user"
    assert mailbox_base(context) == "/me"


def test_user_mail_connection_with_shared_mailbox() -> None:
    context = resolve_connection(
        _payload(
            "m365_mail_user", config={"mailbox": "cmr@acme.example"}, secrets={"access_token": "t"}
        ),
        service="mail",
    )
    assert mailbox_base(context) == "/users/cmr@acme.example"


def test_app_mail_connection_requires_mailbox_and_secret() -> None:
    with pytest.raises(M365Error) as caught:
        resolve_connection(
            _payload(
                "m365_mail_app",
                config={"tenant_id": TENANT, "client_id": CLIENT},
                secrets={"client_secret": "s"},
            ),
            service="mail",
        )
    assert caught.value.code == "invalid_connection"
    with pytest.raises(M365Error) as caught:
        resolve_connection(
            _payload(
                "m365_mail_app",
                config={"tenant_id": TENANT, "client_id": CLIENT, "mailbox": "a@acme.example"},
            ),
            service="mail",
        )
    assert caught.value.code == "invalid_client"


def test_service_mismatch_is_rejected() -> None:
    with pytest.raises(M365Error) as caught:
        resolve_connection(
            _payload("m365_files_user", secrets={"access_token": "t"}), service="mail"
        )
    assert caught.value.code == "invalid_connection"


def test_connection_ref_must_match() -> None:
    with pytest.raises(M365Error):
        resolve_connection(
            _payload("m365_mail_user", secrets={"access_token": "t"}),
            service="mail",
            connection_ref="other",
        )


def test_missing_delegated_token_requires_authorization() -> None:
    with pytest.raises(M365Error) as caught:
        resolve_connection(_payload("m365_excel_user"), service="excel")
    assert caught.value.code == "authorization_required"


def _app_context():
    return resolve_connection(
        _payload(
            "m365_mail_app",
            config={"tenant_id": TENANT, "client_id": CLIENT, "mailbox": "cmr@acme.example"},
            secrets={"client_secret": "shh"},
        ),
        service="mail",
    )


def test_app_token_uses_client_credentials_on_tenant_endpoint() -> None:
    opener = FakeOpener([(200, {"access_token": "app-token", "token_type": "Bearer"})])
    assert app_token(_app_context(), opener=opener) == "app-token"
    request = opener.requests[0]
    assert request.full_url == f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"
    form = parse_qs(request.data.decode())
    assert form["grant_type"] == ["client_credentials"]
    assert form["scope"] == ["https://graph.microsoft.com/.default"]
    assert request.get_header("Authorization") is None


def test_app_token_expired_secret() -> None:
    opener = FakeOpener(
        [
            http_error(
                "u",
                401,
                {
                    "error": "invalid_client",
                    "error_description": "AADSTS7000222: The provided client secret keys are expired.",
                },
            )
        ]
    )
    with pytest.raises(M365Error) as caught:
        app_token(_app_context(), opener=opener)
    assert caught.value.code == "client_secret_expired"
