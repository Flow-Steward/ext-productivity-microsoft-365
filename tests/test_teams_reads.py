"""Teams reads (v0.2 D3): connection test, teams, channels, chats, messages and pickers."""

from __future__ import annotations

from typing import Any

import pytest
from test_teams_fakes import (
    CHANNEL_ID,
    CHAT_ID,
    TEAM_ID,
    USER_CONNECTION,
    ScriptedGraph,
    action_payload,
    not_found,
    query_payload,
)

from m365.errors import M365Error
from m365.runtime import run_operation
from teams.cursors import CURSOR_PREFIX
from teams.markers import HtmlCommentMarker, ZeroWidthMarker
from teams.operations import OPERATIONS, handle_query


def _run(
    graph: Any, action_id: str, operation_input: dict[str, Any] | None = None, **kwargs: Any
) -> dict[str, Any]:
    payload = action_payload(action_id, operation_input, **kwargs)
    return run_operation(payload, OPERATIONS, transport_factory=lambda _connection: graph)


def _chat(chat_id: str, topic: str | None, *members: tuple[str, str]) -> dict[str, Any]:
    return {
        "id": chat_id,
        "topic": topic,
        "chatType": "group" if topic else "oneOnOne",
        "webUrl": f"https://teams.microsoft.com/l/chat/{chat_id}",
        "lastUpdatedDateTime": "2026-11-20T08:00:00Z",
        "members": [
            {"displayName": name, "email": email, "userId": f"id-{name.lower()}"}
            for name, email in members
        ],
    }


def _msg(message_id: str, content: str, **extra: Any) -> dict[str, Any]:
    row = {
        "id": message_id,
        "messageType": "message",
        "createdDateTime": f"2026-11-22T10:{message_id[-2:]}:00Z",
        "lastModifiedDateTime": f"2026-11-22T10:{message_id[-2:]}:00Z",
        "subject": None,
        "importance": "normal",
        "body": {"contentType": "html", "content": content},
        "from": {"user": {"id": "u-ana", "displayName": "Ana Acme"}},
        "webUrl": f"https://teams.microsoft.com/l/message/{message_id}",
    }
    row.update(extra)
    return row


# -- test_teams_connection, list_teams, list_channels -------------------------------------------


def test_connection_reports_the_user_and_joined_team_count() -> None:
    graph = ScriptedGraph(
        {"id": "u-svc", "displayName": "Ops Bot", "userPrincipalName": "ops@acme.example"},
        {
            "value": [{"id": "t1"}],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/joinedTeams?$skip=1",
        },
        {"value": [{"id": "t2"}]},
    )
    response = _run(graph, "test_teams_connection")

    assert response["ok"] is True, response
    assert response["result"] == {
        "user_id": "u-svc",
        "display_name": "Ops Bot",
        "user_principal_name": "ops@acme.example",
        "auth_mode": "user",
        "joined_team_count": 2,
    }
    assert [call.url for call in graph.calls][:2] == ["/me", "/me/joinedTeams"]


def test_connection_without_a_token_is_authorization_required() -> None:
    connection = {**USER_CONNECTION, "secrets": {}}
    response = _run(ScriptedGraph(), "test_teams_connection", connection=connection)
    assert response["error_code"] == "authorization_required"


def test_a_mail_connection_is_refused_for_teams_operations() -> None:
    connection = {**USER_CONNECTION, "connection_type_id": "m365_mail_user"}
    response = _run(ScriptedGraph(), "list_teams", connection=connection)
    assert response["error_code"] == "invalid_connection"


def test_list_teams_projects_joined_teams() -> None:
    graph = ScriptedGraph(
        {"value": [{"id": TEAM_ID, "displayName": "Operations", "description": None}]}
    )
    response = _run(graph, "list_teams")

    assert response["result"] == {
        "teams": [{"team_id": TEAM_ID, "display_name": "Operations", "description": ""}],
        "count": 1,
    }
    assert graph.calls[0].query == {"$select": "id,displayName,description"}


def test_list_channels_returns_notification_targets() -> None:
    graph = ScriptedGraph(
        {
            "value": [
                {"id": CHANNEL_ID, "displayName": "General", "membershipType": "standard"},
                {
                    "id": "19:priv@thread.tacv2",
                    "displayName": "Secret",
                    "membershipType": "private",
                },
            ]
        }
    )
    response = _run(graph, "list_channels", {"team_id": TEAM_ID.upper()})

    assert response["ok"] is True, response
    result = response["result"]
    assert result["team_id"] == TEAM_ID
    assert result["count"] == 2
    assert result["channels"][0]["target"] == f"{TEAM_ID}/{CHANNEL_ID}"
    assert result["channels"][1]["membership_type"] == "private"
    assert graph.calls[0].url == f"/teams/{TEAM_ID}/channels"


def test_list_channels_of_an_unknown_team_is_team_not_found() -> None:
    response = _run(ScriptedGraph(not_found()), "list_channels", {"team_id": TEAM_ID})
    assert response["error_code"] == "team_not_found"


@pytest.mark.parametrize("team_id", ["", "not-a-guid", 42, f"{TEAM_ID}/x"])
def test_list_channels_validates_the_team_id(team_id: Any) -> None:
    response = _run(ScriptedGraph(), "list_channels", {"team_id": team_id})
    assert response["error_code"] == "invalid_payload"


def test_unknown_input_fields_are_refused() -> None:
    response = _run(ScriptedGraph(), "list_teams", {"team_id": TEAM_ID})
    assert response["error_code"] == "invalid_payload"


# -- list_chats -------------------------------------------------------------------------------------


def test_list_chats_filters_by_topic_and_member() -> None:
    page = {
        "value": [
            _chat("19:a@thread.v2", "Release planning", ("Ana", "ana@acme.example")),
            _chat("19:b@thread.v2", "Lunch", ("Ana", "ana@acme.example")),
            _chat("19:c@thread.v2", "Release retro", ("Bob", "bob@acme.example")),
        ]
    }
    graph = ScriptedGraph(page)
    response = _run(
        graph, "list_chats", {"topic_contains": "RELEASE", "member": "Ana@acme.example"}
    )

    assert response["ok"] is True, response
    assert [chat["chat_id"] for chat in response["result"]["chats"]] == ["19:a@thread.v2"]
    assert response["result"]["next_cursor"] == ""
    first = graph.calls[0].url
    assert first.startswith("https://graph.microsoft.com/v1.0/me/chats?")
    assert "%24expand=members" in first


def test_list_chats_cursor_resumes_inside_the_same_page() -> None:
    page = {"value": [_chat(f"19:c{i}@thread.v2", f"Topic {i}") for i in range(3)]}
    graph = ScriptedGraph(page, page)
    first = _run(graph, "list_chats", {"limit": 2})
    cursor = first["result"]["next_cursor"]

    assert [c["chat_id"] for c in first["result"]["chats"]] == [
        "19:c0@thread.v2",
        "19:c1@thread.v2",
    ]
    assert cursor.startswith(CURSOR_PREFIX)

    second = _run(graph, "list_chats", {"limit": 2, "cursor": cursor})
    assert [c["chat_id"] for c in second["result"]["chats"]] == ["19:c2@thread.v2"]
    assert second["result"]["next_cursor"] == ""
    assert graph.calls[1].url == graph.calls[0].url  # the same page, read again


def test_list_chats_cursor_follows_the_next_link() -> None:
    next_link = "https://graph.microsoft.com/v1.0/me/chats?$expand=members&$skiptoken=abc"
    graph = ScriptedGraph(
        {"value": [_chat("19:x@thread.v2", "Ops")], "@odata.nextLink": next_link},
        {"value": [_chat("19:y@thread.v2", "Ops 2")]},
    )
    response = _run(graph, "list_chats", {"limit": 2})
    assert [c["chat_id"] for c in response["result"]["chats"]] == [
        "19:x@thread.v2",
        "19:y@thread.v2",
    ]
    assert graph.calls[1].url == next_link


def test_a_cursor_of_other_filters_is_refused() -> None:
    page = {"value": [_chat(f"19:c{i}@thread.v2", "Ops") for i in range(3)]}
    first = _run(ScriptedGraph(page), "list_chats", {"limit": 1, "topic_contains": "ops"})
    cursor = first["result"]["next_cursor"]
    response = _run(ScriptedGraph(), "list_chats", {"limit": 1, "cursor": cursor})
    assert response["error_code"] == "invalid_cursor"


def test_a_forged_cursor_is_refused() -> None:
    response = _run(ScriptedGraph(), "list_chats", {"cursor": CURSOR_PREFIX + "eyJ2IjoxfQ"})
    assert response["error_code"] == "invalid_cursor"


def test_list_chats_validates_the_member_address() -> None:
    response = _run(ScriptedGraph(), "list_chats", {"member": "not an address"})
    assert response["error_code"] == "invalid_payload"


# -- list_messages ---------------------------------------------------------------------------------


def test_list_channel_messages_newest_first_with_markers_stripped() -> None:
    key = "post_channel_message:step-1"
    graph = ScriptedGraph(
        {
            "value": [
                _msg("1700000000002", HtmlCommentMarker().embed("<p>Deploy <b>done</b></p>", key)),
                _msg("1700000000001", ZeroWidthMarker().embed("Hello&nbsp;team", key)),
                _msg("1700000000000", "", messageType="systemEventMessage"),
                _msg("1699999999999", "gone", deletedDateTime="2026-11-22T09:00:00Z"),
            ]
        }
    )
    response = _run(graph, "list_messages", {"team_id": TEAM_ID, "channel_id": CHANNEL_ID})

    assert response["ok"] is True, response
    messages = response["result"]["messages"]
    assert [m["message_id"] for m in messages] == ["1700000000002", "1700000000001"]
    assert messages[0]["body_text"] == "Deploy done"
    assert messages[1]["body_text"] == "Hello team"
    assert "fs:" not in str(messages) and "⁠" not in str(messages)
    assert messages[0]["sender"] == {"display_name": "Ana Acme", "user_id": "u-ana", "kind": "user"}
    url = graph.calls[0].url
    assert url.startswith(
        f"https://graph.microsoft.com/v1.0/teams/{TEAM_ID}/channels/{CHANNEL_ID}/messages?"
    )


def test_list_chat_messages_orders_newest_first_and_pages() -> None:
    next_link = f"https://graph.microsoft.com/v1.0/chats/{CHAT_ID}/messages?$skiptoken=n1"
    graph = ScriptedGraph(
        {
            "value": [_msg("1700000000003", "c"), _msg("1700000000002", "b")],
            "@odata.nextLink": next_link,
        }
    )
    response = _run(graph, "list_messages", {"chat_id": CHAT_ID, "limit": 2})

    assert response["result"]["count"] == 2
    assert "%24orderby=createdDateTime+desc" in graph.calls[0].url
    cursor = response["result"]["next_cursor"]
    graph2 = ScriptedGraph({"value": [_msg("1700000000001", "a")]})
    second = _run(graph2, "list_messages", {"chat_id": CHAT_ID, "limit": 2, "cursor": cursor})
    assert graph2.calls[0].url == next_link
    assert [m["body_text"] for m in second["result"]["messages"]] == ["a"]


def test_a_chat_cursor_cannot_read_another_chat() -> None:
    next_link = f"https://graph.microsoft.com/v1.0/chats/{CHAT_ID}/messages?$skiptoken=n1"
    graph = ScriptedGraph({"value": [_msg("1700000000003", "c")], "@odata.nextLink": next_link})
    cursor = _run(graph, "list_messages", {"chat_id": CHAT_ID, "limit": 1})["result"]["next_cursor"]
    response = _run(
        ScriptedGraph(), "list_messages", {"chat_id": "19:other@thread.v2", "cursor": cursor}
    )
    assert response["error_code"] == "invalid_cursor"


@pytest.mark.parametrize(
    ("given", "code"),
    [
        ({"chat_id": CHAT_ID}, "chat_not_found"),
        ({"team_id": TEAM_ID, "channel_id": CHANNEL_ID}, "channel_not_found"),
    ],
)
def test_list_messages_not_found_codes(given: dict[str, str], code: str) -> None:
    response = _run(ScriptedGraph(not_found()), "list_messages", given)
    assert response["error_code"] == code


@pytest.mark.parametrize(
    "given",
    [
        {},
        {"team_id": TEAM_ID},
        {"chat_id": CHAT_ID, "team_id": TEAM_ID, "channel_id": CHANNEL_ID},
        {"chat_id": "not-a-chat"},
    ],
)
def test_list_messages_needs_one_target(given: dict[str, str]) -> None:
    response = _run(ScriptedGraph(), "list_messages", given)
    assert response["error_code"] == "invalid_payload"


# -- resource queries --------------------------------------------------------------------------------


def _query(graph: Any, query_id: str, **kwargs: Any) -> dict[str, Any]:
    return handle_query(query_payload(query_id, **kwargs), transport_factory=lambda _c: graph)


def test_teams_picker_filters_by_name() -> None:
    graph = ScriptedGraph(
        {
            "value": [
                {"id": TEAM_ID, "displayName": "Operations"},
                {"id": "t2", "displayName": "Sales"},
            ]
        }
    )
    response = _query(graph, "teams_user", search="oper")
    assert response["result"]["items"] == [{"id": TEAM_ID, "label": "Operations"}]


def test_channels_picker_lists_targets_and_skips_unreadable_teams() -> None:
    other = "1b8c2f3e-1d2a-4c5b-9e8f-7a6b5c4d3e2f"
    graph = ScriptedGraph(
        {
            "value": [
                {"id": TEAM_ID, "displayName": "Operations"},
                {"id": other, "displayName": "Old"},
            ]
        },
        {"value": [{"id": CHANNEL_ID, "displayName": "Alerts"}]},
        M365Error("permission_denied", "no", http_status=403),
    )
    response = _query(graph, "channels_user")
    assert response["ok"] is True, response
    assert response["result"]["items"] == [
        {"id": f"{TEAM_ID}/{CHANNEL_ID}", "label": "Operations / Alerts"}
    ]


def test_chats_picker_matches_member_names() -> None:
    graph = ScriptedGraph(
        {
            "value": [
                _chat("19:a@thread.v2", None, ("Ana Acme", "ana@acme.example")),
                _chat("19:b@thread.v2", "Ops", ("Bob", "bob@acme.example")),
            ]
        }
    )
    response = _query(graph, "chats_user", search="ana")
    assert response["result"]["items"] == [
        {"id": "19:a@thread.v2", "label": "Ana Acme", "chat_type": "oneOnOne"}
    ]


def test_queries_refuse_other_connection_types_and_unknown_ids() -> None:
    mail = {**USER_CONNECTION, "connection_type_id": "m365_mail_user"}
    assert (
        _query(ScriptedGraph(), "teams_user", connection=mail)["error_code"] == "invalid_connection"
    )
    assert _query(ScriptedGraph(), "nope")["error_code"] == "invalid_payload"
