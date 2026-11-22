"""Teams notification adapters (v0.2 D4) used by ``notify.deliver``."""

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any

import pytest
from fakes import FakeOpener
from test_teams_fakes import CHANNEL_ID, CHAT_ID, TEAM_ID, WEBHOOK_URL, ScriptedGraph, ambiguous

from m365.errors import M365Error
from teams import posting
from teams.markers import DEFAULT_MARKER, strip_markers
from teams.notify import post_notification, post_webhook_notification


@dataclass(frozen=True)
class Content:
    title: str = "Job failed: Invoices"
    severity: str = "error"
    summary: str = "Step parse_pdf failed"
    run_url: str = "https://fs.example/runs/42"
    details: tuple[tuple[str, str], ...] = (("Step", "parse_pdf"), ("Error", "bad <pdf>"))


@pytest.fixture(autouse=True)
def _no_sleep(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(posting, "sleep", lambda _s: None)


def test_channel_notification_posts_html_with_the_core_key_marker() -> None:
    graph = ScriptedGraph({"id": "1700000000001", "createdDateTime": "2026-11-22T10:00:00Z"})
    result = post_notification(
        graph, target=f"{TEAM_ID}/{CHANNEL_ID}", content=Content(), key="core-idem-1"
    )

    assert result["message_id"] == "1700000000001"
    assert result["confirmation"] == "posted"
    [post] = graph.posts
    assert post.url == f"/teams/{TEAM_ID}/channels/{CHANNEL_ID}/messages"
    content = post.json_body["body"]["content"]
    assert DEFAULT_MARKER.found(content, "core-idem-1")
    html_body = strip_markers(content)
    assert "<strong>Job failed: Invoices</strong>" in html_body
    assert "ERROR" in html_body
    assert "<strong>Error:</strong> bad &lt;pdf&gt;" in html_body
    assert '<a href="https://fs.example/runs/42">Open in Flow Steward</a>' in html_body


def test_chat_notification_and_ambiguous_outcome() -> None:
    graph = ScriptedGraph(ambiguous(), {"value": []}, {"value": []}, {"value": []})
    with pytest.raises(M365Error) as raised:
        post_notification(graph, target=CHAT_ID, content=Content(), key="core-idem-2")
    assert raised.value.code == "timeout_unknown"
    assert len(graph.posts) == 1
    assert graph.posts[0].url == f"/chats/{CHAT_ID}/messages"


def test_webhook_notification_posts_an_adaptive_card() -> None:
    opener = FakeOpener([(202, b"")])
    result = post_webhook_notification(WEBHOOK_URL, content=Content(), opener=opener)

    assert result["confirmation"] == "accepted_by_webhook"
    assert result["provider_message_id"] == ""
    body: dict[str, Any] = json.loads(opener.requests[0].data)
    card = body["attachments"][0]["content"]
    assert [block["type"] for block in card["body"]] == [
        "TextBlock",
        "TextBlock",
        "TextBlock",
        "FactSet",
    ]
    assert card["body"][3]["facts"][1] == {"title": "Error", "value": "bad <pdf>"}
    assert card["actions"][0]["url"] == "https://fs.example/runs/42"


def test_webhook_notification_refuses_a_foreign_host() -> None:
    opener = FakeOpener([])
    with pytest.raises(M365Error) as raised:
        post_webhook_notification(
            "https://hooks.example/abc?sig=x", content=Content(), opener=opener
        )
    assert raised.value.code == "egress_host_not_allowed"
    assert opener.requests == []
