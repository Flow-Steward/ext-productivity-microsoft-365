"""Mail reads (B5.1): connection test, folders and picker queries, search filters and
cursors, message details, and attachment download into an artifact."""

from __future__ import annotations

import base64
import json
from typing import Any

import pytest
from test_mail_fakes import APP_CONNECTION, USER_CONNECTION, ScriptedGraph, action_payload, sha

from m365.errors import M365Error
from m365.runtime import run_operation
from mail import download
from mail.cursors import CURSOR_PREFIX, fingerprint
from mail.operations import OPERATIONS, handle_query
from mail.reads import _search_query


def _run(
    graph: Any, action_id: str, operation_input: dict[str, Any] | None = None, **kwargs: Any
) -> dict[str, Any]:
    payload = action_payload(action_id, operation_input, **kwargs)
    return run_operation(payload, OPERATIONS, transport_factory=lambda _connection: graph)


def _not_found() -> M365Error:
    return M365Error("item_not_found", "not found", http_status=404)


# -- test_connection ----------------------------------------------------------------------


def test_connection_of_a_signed_in_user_reads_the_profile_and_inbox_counts() -> None:
    graph = ScriptedGraph(
        {
            "displayName": "Ana Acme",
            "mail": "ana@acme.example",
            "userPrincipalName": "ana@acme.example",
        },
        {"id": "F-INBOX", "totalItemCount": 12, "unreadItemCount": 3},
    )
    response = _run(graph, "test_connection")

    assert response["ok"] is True, response
    assert response["result"] == {
        "mailbox_address": "ana@acme.example",
        "display_name": "Ana Acme",
        "auth_mode": "user",
        "inbox_total": 12,
        "inbox_unread": 3,
    }
    assert [call.url for call in graph.calls] == ["/me", "/me/mailFolders/inbox"]


def test_connection_of_an_app_mailbox_uses_the_users_path_only() -> None:
    graph = ScriptedGraph({"id": "F-INBOX", "totalItemCount": 1, "unreadItemCount": 0})
    response = _run(graph, "test_connection", connection=APP_CONNECTION)

    assert response["ok"] is True, response
    assert response["result"]["mailbox_address"] == "cmr@acme.example"
    assert response["result"]["display_name"] is None
    assert response["result"]["auth_mode"] == "app"
    assert [call.url for call in graph.calls] == ["/users/cmr@acme.example/mailFolders/inbox"]


def test_connection_to_a_missing_mailbox_is_mailbox_not_found() -> None:
    response = _run(ScriptedGraph(_not_found()), "test_connection", connection=APP_CONNECTION)
    assert response["error_code"] == "mailbox_not_found"


# -- folders --------------------------------------------------------------------------------

_FOLDER_SCRIPT = (
    {"id": "F-INBOX"},
    _not_found(),  # no Archive folder in this mailbox
    {"id": "F-SENT"},
    {"id": "F-DRAFTS"},
    {"id": "F-DELETED"},
    {"id": "F-JUNK"},
    {
        "value": [
            {
                "id": "F-INBOX",
                "displayName": "Inbox",
                "childFolderCount": 1,
                "totalItemCount": 5,
                "unreadItemCount": 2,
            },
            {"id": "F-SENT", "displayName": "Sent Items", "childFolderCount": 0},
        ],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/mailFolders?$skip=2",
    },
    {"value": [{"id": "F-DRAFTS", "displayName": "Drafts", "childFolderCount": 0}]},
    {
        "value": [
            {"id": "F-PROC", "displayName": "Processed", "childFolderCount": 0, "totalItemCount": 9}
        ]
    },
)


def test_list_mail_folders_nests_folders_and_marks_well_known_ones() -> None:
    graph = ScriptedGraph(*_FOLDER_SCRIPT)
    response = _run(graph, "list_mail_folders")

    assert response["ok"] is True, response
    result = response["result"]
    assert [row["folder_id"] for row in result["well_known"]] == [
        "inbox",
        "sentitems",
        "drafts",
        "deleteditems",
        "junkemail",
    ]
    by_id = {row["folder_id"]: row for row in result["folders"]}
    assert by_id["F-PROC"]["path"] == "Inbox / Processed"
    assert by_id["F-PROC"]["parent_folder_id"] == "F-INBOX"
    assert by_id["F-PROC"]["depth"] == 1
    assert by_id["F-INBOX"]["well_known_name"] == "inbox"
    assert by_id["F-DRAFTS"]["well_known_name"] == "drafts"
    assert result["truncated"] is False
    assert graph.calls[7].url == "https://graph.microsoft.com/v1.0/me/mailFolders?$skip=2"
    assert graph.calls[8].url == "/me/mailFolders/F-INBOX/childFolders"


def _query_payload(query_id: str, connection: dict[str, Any], search: str = "") -> dict[str, Any]:
    payload = action_payload("list_mail_folders", connection=connection)
    payload.pop("action")
    payload["mode"] = "query"
    payload["query"] = {
        "query_id": query_id,
        "context": {"connection_id": connection["connection_id"]},
        "search": search,
    }
    payload["target"] = {"connection": connection}
    return payload


def test_folder_picker_query_lists_shortcuts_then_paths() -> None:
    graph = ScriptedGraph(*_FOLDER_SCRIPT)
    response = handle_query(
        _query_payload("mail_folders_user", USER_CONNECTION), transport_factory=lambda _c: graph
    )

    assert response["ok"] is True, response
    items = response["result"]["items"]
    assert items[0] == {"folder_id": "inbox", "label": "Inbox (well-known)"}
    assert {"folder_id": "F-PROC", "label": "Inbox / Processed"} in items
    assert response["result"]["next_cursor"] == ""
    assert response["result"]["total"] == len(items)


def test_folder_picker_query_filters_by_search_text() -> None:
    graph = ScriptedGraph(*_FOLDER_SCRIPT)
    response = handle_query(
        _query_payload("mail_folders_user", USER_CONNECTION, search="proc"),
        transport_factory=lambda _c: graph,
    )
    assert response["result"]["items"] == [{"folder_id": "F-PROC", "label": "Inbox / Processed"}]


def test_folder_picker_query_serves_exactly_one_connection_type() -> None:
    graph = ScriptedGraph()
    response = handle_query(
        _query_payload("mail_folders_app", USER_CONNECTION), transport_factory=lambda _c: graph
    )
    assert response["error_code"] == "invalid_connection"
    assert graph.calls == []


# -- search -----------------------------------------------------------------------------------

_ROW = {
    "id": "AAMk-MSG-1",
    "subject": "Invoice 42",
    "from": {"emailAddress": {"address": "billing@acme.example", "name": "Billing"}},
    "toRecipients": [{"emailAddress": {"address": "cmr@acme.example", "name": ""}}],
    "receivedDateTime": "2026-11-20T09:00:00Z",
    "isRead": False,
    "hasAttachments": True,
    "flag": {"flagStatus": "flagged"},
    "parentFolderId": "F-INBOX",
    "bodyPreview": "Please find attached",
}
_NEXT = "https://graph.microsoft.com/v1.0/me/mailFolders('inbox')/messages?$skiptoken=abc"
_FILTERS = {
    "from": "o'brien@acme.example",
    "subject_contains": "Invoice",
    "unread": True,
    "has_attachments": True,
    "received_after": "2026-11-01T00:00:00+02:00",
}


def test_search_builds_an_ordered_filter_and_an_opaque_cursor() -> None:
    graph = ScriptedGraph({"value": [_ROW], "@odata.nextLink": _NEXT})
    response = _run(graph, "search_messages", _FILTERS)

    assert response["ok"] is True, response
    query = graph.calls[0].query
    assert graph.calls[0].url == "/me/mailFolders/inbox/messages"
    assert query["$filter"] == (
        "receivedDateTime ge 2026-10-31T22:00:00Z and from/emailAddress/address eq 'o''brien@acme.example' "
        "and contains(subject,'Invoice') and isRead eq false and hasAttachments eq true"
    )
    assert query["$orderby"] == "receivedDateTime desc"
    assert query["$top"] == "25"
    message = response["result"]["messages"][0]
    assert message["message_id"] == "AAMk-MSG-1"
    assert message["from"] == {"address": "billing@acme.example", "name": "Billing"}
    assert message["flag_status"] == "flagged"
    cursor = response["result"]["next_cursor"]
    assert cursor.startswith(CURSOR_PREFIX)
    assert "graph.microsoft.com" not in cursor and "skiptoken" not in cursor

    graph_two = ScriptedGraph({"value": []})
    second = _run(graph_two, "search_messages", {**_FILTERS, "cursor": cursor})
    assert second["ok"] is True, second
    assert graph_two.calls[0].url == _NEXT
    assert graph_two.calls[0].query == {}
    assert second["result"]["next_cursor"] == ""


def test_a_cursor_only_continues_the_same_search() -> None:
    graph = ScriptedGraph({"value": [_ROW], "@odata.nextLink": _NEXT})
    cursor = _run(graph, "search_messages", _FILTERS)["result"]["next_cursor"]
    other = ScriptedGraph()
    response = _run(
        other, "search_messages", {**_FILTERS, "subject_contains": "Order", "cursor": cursor}
    )
    assert response["error_code"] == "invalid_cursor"
    assert other.calls == []


@pytest.mark.parametrize(
    "link",
    [
        "https://graph.microsoft.com/v1.0/users/ceo@acme.example/messages",
        "https://graph.example/v1.0/me/messages",
        "https://graph.microsoft.com/beta/me/messages",
    ],
)
def test_a_forged_cursor_cannot_leave_the_mailbox(link: str) -> None:
    key = fingerprint("/me", _search_query({}))
    raw = json.dumps({"v": 1, "f": key, "l": link}).encode()
    cursor = CURSOR_PREFIX + base64.urlsafe_b64encode(raw).decode().rstrip("=")
    graph = ScriptedGraph()
    response = _run(graph, "search_messages", {"cursor": cursor})
    assert response["error_code"] == "invalid_cursor"
    assert graph.calls == []


def test_app_mailbox_cursor_accepts_the_mailbox_path() -> None:
    link = "https://graph.microsoft.com/v1.0/users/cmr@acme.example/mailFolders/inbox/messages?$skip=25"
    graph = ScriptedGraph({"value": [], "@odata.nextLink": link})
    response = _run(graph, "search_messages", {}, connection=APP_CONNECTION)
    assert response["ok"] is True, response
    assert response["result"]["next_cursor"].startswith(CURSOR_PREFIX)


def test_free_text_query_uses_search_without_filter_or_order() -> None:
    graph = ScriptedGraph({"value": []})
    response = _run(
        graph, "search_messages", {"query": "invoice 42", "folder": "archive", "limit": 5}
    )

    assert response["ok"] is True, response
    assert graph.calls[0].url == "/me/mailFolders/archive/messages"
    assert graph.calls[0].query["$search"] == '"invoice 42"'
    assert "$filter" not in graph.calls[0].query and "$orderby" not in graph.calls[0].query


@pytest.mark.parametrize(
    ("operation_input", "code"),
    [
        ({"query": "invoice", "unread": True}, "search_unsupported"),
        ({"query": 'say "hi"'}, "invalid_payload"),
        ({"received_after": "2026-11-01"}, "invalid_payload"),
        (
            {"received_after": "2026-11-02T00:00:00Z", "received_before": "2026-11-01T00:00:00Z"},
            "invalid_payload",
        ),
        ({"limit": 101}, "invalid_payload"),
        ({"folder": "../inbox"}, "invalid_payload"),
        ({"unread": "yes"}, "invalid_payload"),
    ],
)
def test_search_input_is_closed(operation_input: dict[str, Any], code: str) -> None:
    graph = ScriptedGraph()
    response = _run(graph, "search_messages", operation_input)
    assert response["error_code"] == code
    assert graph.calls == []


def test_search_in_a_missing_folder_is_folder_not_found() -> None:
    response = _run(ScriptedGraph(_not_found()), "search_messages", {"folder": "AAMk-F-GONE"})
    assert response["error_code"] == "folder_not_found"


# -- get_message --------------------------------------------------------------------------------


def test_get_message_returns_text_body_html_headers_and_attachment_types() -> None:
    graph = ScriptedGraph(
        {
            **_ROW,
            "body": {"contentType": "text", "content": "Hello\nInvoice attached"},
            "internetMessageHeaders": [{"name": "X-Mailer", "value": "Acme"}],
            "replyTo": [{"emailAddress": {"address": "ap@acme.example"}}],
        },
        {"body": {"contentType": "html", "content": "<p>Hello</p>"}},
        {
            "value": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "id": "A1",
                    "name": "inv.pdf",
                    "contentType": "application/pdf",
                    "size": 300,
                    "isInline": False,
                },
                {
                    "@odata.type": "#microsoft.graph.itemAttachment",
                    "id": "A2",
                    "name": "fwd",
                    "size": 90,
                },
                {
                    "@odata.type": "#microsoft.graph.referenceAttachment",
                    "id": "A3",
                    "name": "link",
                    "size": 5,
                },
            ]
        },
    )
    response = _run(graph, "get_message", {"message_id": "AAMk-MSG-1", "include_raw_html": True})

    assert response["ok"] is True, response
    message = response["result"]["message"]
    assert message["body_text"] == "Hello\nInvoice attached"
    assert message["body_text_truncated"] is False
    assert message["body_html"] == "<p>Hello</p>"
    assert message["headers"] == [{"name": "X-Mailer", "value": "Acme"}]
    assert message["reply_to"] == [{"address": "ap@acme.example", "name": ""}]
    assert [a["attachment_type"] for a in message["attachments"]] == ["file", "item", "reference"]
    assert 'outlook.body-content-type="text"' in graph.calls[0].headers["Prefer"]
    assert 'IdType="ImmutableId"' in graph.calls[0].headers["Prefer"]
    assert 'outlook.body-content-type="html"' in graph.calls[1].headers["Prefer"]
    assert "internetMessageHeaders" in graph.calls[0].query["$select"]


def test_get_message_truncates_a_huge_body() -> None:
    graph = ScriptedGraph({**_ROW, "body": {"content": "й" * (600 * 1024)}})
    response = _run(
        graph,
        "get_message",
        {
            "message_id": "AAMk-MSG-1",
            "include_headers": False,
            "include_attachment_metadata": False,
        },
    )
    message = response["result"]["message"]
    assert message["body_text_truncated"] is True
    assert len(message["body_text"].encode("utf-8")) <= 1024 * 1024
    assert "headers" not in message and "attachments" not in message


def test_get_missing_message_is_message_not_found() -> None:
    response = _run(ScriptedGraph(_not_found()), "get_message", {"message_id": "AAMk-GONE"})
    assert response["error_code"] == "message_not_found"


# -- get_attachment -------------------------------------------------------------------------------


@pytest.fixture
def written(monkeypatch: pytest.MonkeyPatch) -> list[bytes]:
    captured: list[bytes] = []

    def writer(_payload: Any, chunks: Any, **kwargs: Any) -> dict[str, Any]:
        assert kwargs["binding_key"] == "attachment_artifact_handle"
        captured.append(b"".join(chunks))
        return {"artifact_handle": "artifact-acme-1"}

    monkeypatch.setattr(download, "output_grant_limit", lambda _payload: 1024 * 1024)
    monkeypatch.setattr(download, "artifact_writer", writer)
    return captured


_FILE_META = {
    "@odata.type": "#microsoft.graph.fileAttachment",
    "id": "A1",
    "name": "../inv.pdf",
    "contentType": "application/pdf",
    "size": 320,
    "isInline": False,
}


def test_get_attachment_streams_the_file_into_an_artifact(written: list[bytes]) -> None:
    content = b"%PDF-1.7 acme"
    graph = ScriptedGraph(_FILE_META, content)
    response = _run(graph, "get_attachment", {"message_id": "AAMk-MSG-1", "attachment_id": "A1"})

    assert response["ok"] is True, response
    assert written == [content]
    assert response["result"]["attachment_artifact_handle"] == "artifact-acme-1"
    assert response["result"]["attachment_metadata"] == {
        "message_id": "AAMk-MSG-1",
        "attachment_id": "A1",
        "original_filename": "inv.pdf",
        "content_type": "application/pdf",
        "size_bytes": len(content),
        "sha256": sha(content),
        "is_inline": False,
    }
    assert graph.calls[1].url == "/me/messages/AAMk-MSG-1/attachments/A1/$value"
    assert "contentBytes" not in json.dumps(response)


@pytest.mark.parametrize("kind", ["itemAttachment", "referenceAttachment"])
def test_attached_messages_and_links_are_unsupported(written: list[bytes], kind: str) -> None:
    graph = ScriptedGraph({**_FILE_META, "@odata.type": f"#microsoft.graph.{kind}"})
    response = _run(graph, "get_attachment", {"message_id": "AAMk-MSG-1", "attachment_id": "A1"})
    assert response["error_code"] == "attachment_type_unsupported"
    assert written == []


def test_an_attachment_over_the_artifact_limit_is_refused(
    monkeypatch: pytest.MonkeyPatch, written: list[bytes]
) -> None:
    monkeypatch.setattr(download, "output_grant_limit", lambda _payload: 8)
    graph = ScriptedGraph({**_FILE_META, "size": 40}, b"x" * 9)
    response = _run(graph, "get_attachment", {"message_id": "AAMk-MSG-1", "attachment_id": "A1"})
    assert response["error_code"] == "attachment_too_large"
    assert written == []


def test_missing_attachment_is_attachment_not_found(written: list[bytes]) -> None:
    response = _run(
        ScriptedGraph(_not_found()),
        "get_attachment",
        {"message_id": "AAMk-MSG-1", "attachment_id": "A9"},
    )
    assert response["error_code"] == "attachment_not_found"


def test_attachment_without_an_output_grant_fails_before_any_call() -> None:
    graph = ScriptedGraph()
    response = _run(graph, "get_attachment", {"message_id": "AAMk-MSG-1", "attachment_id": "A1"})
    assert response["error_code"] == "artifact_output_unavailable"
    assert graph.calls == []
