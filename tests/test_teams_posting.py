"""Teams posts (v0.2 D3): bodies, mentions, file links, size limit and ambiguous posts.

The contract under test: an ambiguous post is never repeated. It is confirmed only by
the hidden step marker in the latest messages; otherwise the step reports
``timeout_unknown`` with a possible external effect.
"""

from __future__ import annotations

import json
from typing import Any

import pytest
from fakes import FakeOpener, FakeResponse
from test_teams_fakes import (
    CHANNEL_ID,
    CHAT_ID,
    TEAM_ID,
    ScriptedGraph,
    action_payload,
    ambiguous,
    echo_posted,
    not_found,
)

from m365.errors import M365Error
from m365.runtime import run_operation
from m365.transport import GraphTransport
from teams import posting
from teams.markers import (
    DEFAULT_MARKER,
    HtmlCommentMarker,
    ZeroWidthMarker,
    marker_token,
    strip_markers,
)
from teams.operations import OPERATIONS
from teams.posting import (
    MAX_MESSAGE_BYTES,
    TeamsTarget,
    build_message,
    parse_target,
    post_html,
)

KEY_PREFIX = "post_channel_message:"
POSTED = {
    "id": "1700000000001",
    "createdDateTime": "2026-11-22T10:00:00Z",
    "webUrl": "https://teams.microsoft.com/l/message/1700000000001",
}


@pytest.fixture(autouse=True)
def _no_sleep(monkeypatch: pytest.MonkeyPatch) -> list[float]:
    slept: list[float] = []
    monkeypatch.setattr(posting, "sleep", slept.append)
    return slept


def _run(
    graph: Any, action_id: str, operation_input: dict[str, Any] | None = None, **kwargs: Any
) -> dict[str, Any]:
    payload = action_payload(action_id, operation_input, **kwargs)
    return run_operation(payload, OPERATIONS, transport_factory=lambda _connection: graph)


def _channel_input(**extra: Any) -> dict[str, Any]:
    return {"team_id": TEAM_ID, "channel_id": CHANNEL_ID, **extra}


# -- post_channel_message ---------------------------------------------------------------------


def test_channel_post_sends_html_once_with_the_step_marker() -> None:
    graph = ScriptedGraph(POSTED)
    response = _run(
        graph,
        "post_channel_message",
        _channel_input(subject="Deploy", body_html="<p>Deploy <b>done</b></p>"),
    )

    assert response["ok"] is True, response
    assert response["external_effect_status"] == "succeeded"
    assert response["definitely_no_external_effect"] is False
    result = response["result"]
    assert result["message_id"] == "1700000000001"
    assert result["confirmation"] == "posted"
    assert (result["team_id"], result["channel_id"], result["chat_id"]) == (TEAM_ID, CHANNEL_ID, "")
    [post] = graph.posts
    assert post.url == f"/teams/{TEAM_ID}/channels/{CHANNEL_ID}/messages"
    assert post.retry == "none"
    assert post.json_body["subject"] == "Deploy"
    content = post.json_body["body"]["content"]
    assert post.json_body["body"]["contentType"] == "html"
    assert content.startswith("<p>Deploy <b>done</b></p>")
    assert DEFAULT_MARKER.found(content, KEY_PREFIX + "6a1b2c3d-4e5f-4a6b-8c7d-000000000301")
    # The marker is a hash of the key: the step execution id never reaches Teams.
    assert "6a1b2c3d" not in content


def test_text_body_is_escaped_html_with_line_breaks() -> None:
    graph = ScriptedGraph(POSTED)
    _run(graph, "post_channel_message", _channel_input(body_text="a < b\nnext & last"))
    content = graph.posts[0].json_body["body"]["content"]
    assert strip_markers(content) == "a &lt; b<br>next &amp; last"


def test_mentions_are_resolved_by_address_and_placed_first() -> None:
    graph = ScriptedGraph(
        {"id": "u-ana", "displayName": "Ana <Ops>"},
        {"id": "u-bob", "displayName": "Bob"},
        POSTED,
    )
    response = _run(
        graph,
        "post_channel_message",
        _channel_input(
            body_text="Please check",
            mentions=["ana@acme.example", "ANA@acme.example", "bob@acme.example"],
        ),
    )

    assert response["ok"] is True, response
    assert [call.url for call in graph.calls[:2]] == [
        "/users/ana@acme.example",
        "/users/bob@acme.example",
    ]
    body = graph.posts[0].json_body
    assert body["body"]["content"].startswith(
        '<p><at id="0">Ana &lt;Ops&gt;</at> <at id="1">Bob</at></p>Please check'
    )
    assert body["mentions"][0] == {
        "id": 0,
        "mentionText": "Ana <Ops>",
        "mentioned": {
            "user": {"id": "u-ana", "displayName": "Ana <Ops>", "userIdentityType": "aadUser"}
        },
    }


def test_an_unknown_mention_is_user_not_found_and_nothing_is_posted() -> None:
    graph = ScriptedGraph(not_found())
    response = _run(
        graph, "post_channel_message", _channel_input(body_text="hi", mentions=["x@acme.example"])
    )
    assert response["error_code"] == "user_not_found"
    assert response["definitely_no_external_effect"] is True
    assert graph.posts == []


def test_file_links_are_rendered_as_links_not_uploads() -> None:
    graph = ScriptedGraph(POSTED)
    link = 'https://acme.sharepoint.com/:x:/g/abc?e=1&x="2"'
    _run(
        graph,
        "post_channel_message",
        _channel_input(body_text="Report", file_links=[{"url": link, "name": "Q3.xlsx"}]),
    )
    content = strip_markers(graph.posts[0].json_body["body"]["content"])
    assert content.endswith(
        '<p><a href="https://acme.sharepoint.com/:x:/g/abc?e=1&amp;x=&quot;2&quot;">Q3.xlsx</a></p>'
    )
    assert "attachments" not in graph.posts[0].json_body


@pytest.mark.parametrize(
    "links", [["http://acme.example/x"], ["javascript:alert(1)"], [{"url": "https://a.b", "x": 1}]]
)
def test_file_links_must_be_https(links: Any) -> None:
    response = _run(
        ScriptedGraph(), "post_channel_message", _channel_input(body_text="x", file_links=links)
    )
    assert response["error_code"] == "invalid_payload"


@pytest.mark.parametrize(
    "bodies", [{}, {"body_text": "a", "body_html": "<p>b</p>"}, {"body_text": "   "}]
)
def test_exactly_one_body_is_required(bodies: dict[str, str]) -> None:
    response = _run(ScriptedGraph(), "post_channel_message", _channel_input(**bodies))
    assert response["error_code"] == "invalid_payload"
    assert response["definitely_no_external_effect"] is True


def test_a_missing_channel_is_channel_not_found_without_effect() -> None:
    response = _run(
        ScriptedGraph(not_found()), "post_channel_message", _channel_input(body_text="x")
    )
    assert response["error_code"] == "channel_not_found"
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True


def test_test_mode_suppresses_the_post() -> None:
    graph = ScriptedGraph()
    response = _run(graph, "post_channel_message", _channel_input(body_text="x"), test_mode=True)
    assert response["result"]["external_effect_status"] == "suppressed"
    assert graph.calls == []


def test_a_throttled_post_is_retryable_and_had_no_effect() -> None:
    throttled = M365Error("rate_limited", "slow down", http_status=429, retry_after_seconds=7)
    response = _run(ScriptedGraph(throttled), "post_channel_message", _channel_input(body_text="x"))
    assert response["error_code"] == "rate_limited"
    assert response["retryable"] is True
    assert response["retry_after_seconds"] == 7
    assert response["definitely_no_external_effect"] is True


# -- size limit -----------------------------------------------------------------------------------


def test_a_body_above_the_limit_is_refused_before_any_request() -> None:
    graph = ScriptedGraph()
    response = _run(
        graph, "post_channel_message", _channel_input(body_text="x" * (MAX_MESSAGE_BYTES + 1))
    )
    assert response["error_code"] == "message_too_large"
    assert response["definitely_no_external_effect"] is True
    assert graph.calls == []


def test_the_limit_counts_utf8_bytes_and_mentions() -> None:
    # 9,000 three-byte characters: 27,000 bytes, under the limit on characters alone.
    body = "\u20ac" * 9_000
    message = build_message(body, marker_key="k")
    assert len(body) < MAX_MESSAGE_BYTES
    assert message["body"]["content"].startswith(body)
    graph = ScriptedGraph()
    response = _run(graph, "post_channel_message", _channel_input(body_text="\u20ac" * 9_600))
    assert response["error_code"] == "message_too_large"


def test_a_graph_413_is_message_too_large() -> None:
    too_large = M365Error("file_too_large", "too large", http_status=413)
    response = _run(ScriptedGraph(too_large), "post_channel_message", _channel_input(body_text="x"))
    assert response["error_code"] == "message_too_large"
    assert response["definitely_no_external_effect"] is True


# -- ambiguous posts ---------------------------------------------------------------------------------


def test_ambiguous_post_with_the_marker_found_succeeds_without_reposting() -> None:
    graph = ScriptedGraph(POSTED)
    graph.script = [ambiguous(), echo_posted(graph, "1700000000042")]
    response = _run(graph, "post_channel_message", _channel_input(body_text="Deploy done"))

    assert response["ok"] is True, response
    assert response["result"]["confirmation"] == "found_by_marker"
    assert response["result"]["message_id"] == "1700000000042"
    assert len(graph.posts) == 1
    lookup = graph.calls[1]
    assert lookup.method == "GET"
    assert lookup.url == f"/teams/{TEAM_ID}/channels/{CHANNEL_ID}/messages"


def test_ambiguous_post_without_the_marker_is_timeout_unknown(_no_sleep: list[float]) -> None:
    other = {"value": [{"id": "1", "body": {"content": "<p>someone else</p>"}}]}
    graph = ScriptedGraph(
        ambiguous(), other, M365Error("rate_limited", "x", http_status=429), other
    )
    response = _run(graph, "post_channel_message", _channel_input(body_text="Deploy done"))

    assert response["error_code"] == "timeout_unknown"
    assert response["external_effect_status"] == "timeout_unknown"
    assert response["definitely_no_external_effect"] is False
    assert response["retryable"] is False
    assert len(graph.posts) == 1, "an ambiguous post must never be repeated"
    assert len(_no_sleep) == posting.CONFIRM_ATTEMPTS - 1


def test_another_steps_marker_does_not_confirm_the_post() -> None:
    foreign = HtmlCommentMarker().embed("<p>Deploy done</p>", "post_channel_message:other-step")
    graph = ScriptedGraph(
        ambiguous(), *[{"value": [{"id": "9", "body": {"content": foreign}}]}] * 3
    )
    response = _run(graph, "post_channel_message", _channel_input(body_text="Deploy done"))
    assert response["error_code"] == "timeout_unknown"


def test_the_real_transport_never_repeats_an_ambiguous_post() -> None:
    """A 504 on the POST through GraphTransport: one POST, then reads only."""
    opener = FakeOpener([(504, {"error": {"code": "GatewayTimeout"}})])
    transport = GraphTransport("token", opener=opener, sleep=lambda _s: None)

    def lookup(request: Any) -> Any:
        assert request.get_method() == "GET"
        posted = opener.requests[0].data.decode()
        token = marker_token("k-1")
        assert token in posted
        row = {"id": "1700000000007", "body": {"content": f"<p>x</p><!-- fs:{token} -->"}}
        return FakeResponse(200, json.dumps({"value": [row]}).encode())

    opener.script.append(lookup)
    result = post_html(
        transport, TeamsTarget(team_id=TEAM_ID, channel_id=CHANNEL_ID), "<p>x</p>", marker_key="k-1"
    )
    assert result["confirmation"] == "found_by_marker"
    assert [r.get_method() for r in opener.requests] == ["POST", "GET"]


def test_zero_width_strategy_is_swappable() -> None:
    marker = ZeroWidthMarker()
    graph = ScriptedGraph(ambiguous())
    graph.script.append(echo_posted(graph))
    target = TeamsTarget(chat_id=CHAT_ID)
    result = post_html(graph, target, "<p>hi</p>", marker_key="key-9", marker=marker)

    content = graph.posts[0].json_body["body"]["content"]
    assert "<!--" not in content
    assert marker.found(content, "key-9")
    assert not marker.found(content, "key-8")
    assert strip_markers(content) == "<p>hi</p>"
    assert result["confirmation"] == "found_by_marker"
    assert graph.calls[1].query == {"$top": "20", "$orderby": "createdDateTime desc"}


def test_markers_in_caller_html_are_removed_before_posting() -> None:
    spoof = HtmlCommentMarker().embed("<p>x</p>", "post_channel_message:victim")
    message = build_message(spoof, marker_key="mine")
    assert not HtmlCommentMarker().found(message["body"]["content"], "post_channel_message:victim")
    assert HtmlCommentMarker().found(message["body"]["content"], "mine")


# -- reply_to_message ---------------------------------------------------------------------------------


def test_reply_posts_to_the_thread_without_a_subject() -> None:
    graph = ScriptedGraph(POSTED)
    response = _run(
        graph, "reply_to_message", _channel_input(message_id="1699999999999", body_text="ack")
    )
    assert response["ok"] is True, response
    assert response["result"]["reply_to_id"] == "1699999999999"
    post = graph.posts[0]
    assert post.url == f"/teams/{TEAM_ID}/channels/{CHANNEL_ID}/messages/1699999999999/replies"
    assert "subject" not in post.json_body
    assert post.retry == "none"


def test_reply_refuses_a_subject_and_maps_a_missing_thread() -> None:
    response = _run(
        ScriptedGraph(),
        "reply_to_message",
        _channel_input(message_id="1", body_text="x", subject="s"),
    )
    assert response["error_code"] == "invalid_payload"
    response = _run(
        ScriptedGraph(not_found()),
        "reply_to_message",
        _channel_input(message_id="1", body_text="x"),
    )
    assert response["error_code"] == "message_not_found"


def test_ambiguous_reply_is_looked_up_in_the_thread() -> None:
    graph = ScriptedGraph(ambiguous())
    graph.script.append(echo_posted(graph))
    response = _run(graph, "reply_to_message", _channel_input(message_id="1699", body_text="ack"))
    assert response["result"]["confirmation"] == "found_by_marker"
    assert graph.calls[1].url.endswith("/messages/1699/replies")


# -- post_chat_message ------------------------------------------------------------------------------


def test_chat_post_to_an_existing_chat() -> None:
    graph = ScriptedGraph(POSTED)
    response = _run(graph, "post_chat_message", {"chat_id": CHAT_ID, "body_text": "hi"})
    assert response["ok"] is True, response
    assert response["result"]["chat_id"] == CHAT_ID
    assert graph.posts[0].url == f"/chats/{CHAT_ID}/messages"


def test_chat_post_to_a_user_creates_or_reuses_the_one_on_one_chat() -> None:
    graph = ScriptedGraph({"id": "u-me"}, {"id": CHAT_ID, "chatType": "oneOnOne"}, POSTED)
    response = _run(graph, "post_chat_message", {"user": "ana@acme.example", "body_text": "hi"})

    assert response["ok"] is True, response
    assert response["result"]["chat_id"] == CHAT_ID
    create, post = graph.posts
    assert create.url == "/chats"
    assert create.retry == "idempotent"  # Graph returns the existing 1:1 chat
    binds = [member["user@odata.bind"] for member in create.json_body["members"]]
    assert binds == [
        "https://graph.microsoft.com/v1.0/users('u-me')",
        "https://graph.microsoft.com/v1.0/users('ana@acme.example')",
    ]
    assert create.json_body["chatType"] == "oneOnOne"
    assert post.url == f"/chats/{CHAT_ID}/messages"
    assert post.retry == "none"


def test_chat_post_to_an_unknown_user_is_user_not_found() -> None:
    graph = ScriptedGraph({"id": "u-me"}, M365Error("invalid_payload", "bad bind", http_status=400))
    response = _run(graph, "post_chat_message", {"user": "nobody@acme.example", "body_text": "hi"})
    assert response["error_code"] == "user_not_found"
    assert response["definitely_no_external_effect"] is True


def test_an_oversized_chat_message_creates_no_chat() -> None:
    graph = ScriptedGraph()
    response = _run(
        graph,
        "post_chat_message",
        {"user": "ana@acme.example", "body_text": "x" * (MAX_MESSAGE_BYTES + 10)},
    )
    assert response["error_code"] == "message_too_large"
    assert graph.calls == []


@pytest.mark.parametrize(
    "given", [{"body_text": "x"}, {"chat_id": CHAT_ID, "user": "a@acme.example", "body_text": "x"}]
)
def test_chat_post_needs_exactly_one_target(given: dict[str, str]) -> None:
    response = _run(ScriptedGraph(), "post_chat_message", given)
    assert response["error_code"] == "invalid_payload"


def test_missing_chat_is_chat_not_found() -> None:
    response = _run(
        ScriptedGraph(not_found()), "post_chat_message", {"chat_id": CHAT_ID, "body_text": "x"}
    )
    assert response["error_code"] == "chat_not_found"


# -- helpers for the notification provider ---------------------------------------------------------


def test_parse_target_accepts_channel_and_chat_routes() -> None:
    assert parse_target(f"{TEAM_ID}/{CHANNEL_ID}") == TeamsTarget(
        team_id=TEAM_ID, channel_id=CHANNEL_ID
    )
    assert parse_target(CHAT_ID) == TeamsTarget(chat_id=CHAT_ID)
    with pytest.raises(M365Error) as missing:
        parse_target("")
    assert missing.value.code == "channel_required"
    with pytest.raises(M365Error) as bad:
        parse_target("general")
    assert bad.value.code == "invalid_payload"


def test_post_html_needs_a_key_and_drops_the_subject_outside_channels() -> None:
    with pytest.raises(M365Error) as missing:
        post_html(ScriptedGraph(), TeamsTarget(chat_id=CHAT_ID), "<p>x</p>", marker_key="")
    assert missing.value.code == "idempotency_key_missing"

    graph = ScriptedGraph(POSTED)
    result = post_html(
        graph, TeamsTarget(chat_id=CHAT_ID), "<p>x</p>", marker_key="core-key", subject="Alert"
    )
    assert "subject" not in graph.posts[0].json_body
    assert result["confirmation"] == "posted"
    assert result["external_effect_status"] == "succeeded"

    graph = ScriptedGraph(POSTED)
    post_html(
        graph,
        TeamsTarget(team_id=TEAM_ID, channel_id=CHANNEL_ID),
        "<p>x</p>",
        marker_key="core-key",
        subject="Alert",
    )
    assert graph.posts[0].json_body["subject"] == "Alert"


def test_post_html_ambiguous_without_marker_raises_timeout_unknown() -> None:
    graph = ScriptedGraph(ambiguous(), {"value": []}, {"value": []}, {"value": []})
    with pytest.raises(M365Error) as raised:
        post_html(graph, TeamsTarget(chat_id=CHAT_ID), "<p>x</p>", marker_key="core-key")
    assert raised.value.code == "timeout_unknown"
    assert raised.value.ambiguous is True
    assert len(graph.posts) == 1


def test_an_unreadable_echo_of_a_successful_post_is_still_a_success() -> None:
    graph = ScriptedGraph({"id": "not a valid id!"})
    response = _run(graph, "post_channel_message", _channel_input(body_text="x"))
    assert response["ok"] is True, response
    assert response["result"]["message_id"] == ""
    assert response["result"]["confirmation"] == "posted"
    assert response["definitely_no_external_effect"] is False
