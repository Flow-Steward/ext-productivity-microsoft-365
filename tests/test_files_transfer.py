"""download_file and upload_file: transfer URLs, fragments, and the ambiguous-upload rule."""

from __future__ import annotations

import hashlib
import json
from typing import Any

import pytest
from test_files_support import (
    APP_CONNECTION,
    DOWNLOAD_URL,
    DRIVE_ID,
    UPLOAD_URL,
    Harness,
    drive_item,
    graph_error,
    upload_artifacts,
)

from files.quickxor import quick_xor_hash

APP_BASE = f"/drives/{DRIVE_ID}"


@pytest.fixture(autouse=True)
def _no_waits(monkeypatch: pytest.MonkeyPatch) -> None:
    from files import transfer

    monkeypatch.setattr(transfer, "sleep", lambda _seconds: None)


def _ok(response: dict) -> dict:
    assert response["ok"] is True, response
    return response["result"]


def _secrets_absent(value: Any) -> None:
    blob = json.dumps(value)
    assert "tempauth" not in blob
    assert "SECRET-" not in blob


# -- download -----------------------------------------------------------------


def _download_harness(content: bytes, *, url: str = DOWNLOAD_URL) -> Harness:
    harness = Harness()
    meta = drive_item("01CMR", "../cmr.pdf", size=len(content))
    meta["@microsoft.graph.downloadUrl"] = url

    def ranged(sent: Any) -> tuple[int, bytes, dict[str, str]]:
        start, end = sent.headers["range"].removeprefix("bytes=").split("-")
        return 206, content[int(start) : int(end) + 1], {}

    harness.graph.on("GET", f"{APP_BASE}/root:/Inbound/cmr.pdf:", (200, meta))
    harness.graph.on("GET", url, ranged)
    return harness


@pytest.fixture
def writer(monkeypatch: pytest.MonkeyPatch) -> list[bytes]:
    from files import transfer

    captured: list[bytes] = []

    def write(_payload: Any, chunks: Any, **kwargs: Any) -> dict[str, Any]:
        assert kwargs["binding_key"] == "files_download_artifact"
        body = b"".join(chunks)
        captured.append(body)
        return {
            "artifact_handle": "artifact:downloaded",
            "size_bytes": len(body),
            "sha256": hashlib.sha256(body).hexdigest(),
        }

    monkeypatch.setattr(transfer, "output_grant_limit", lambda _payload: 10_000_000)
    monkeypatch.setattr(transfer, "write_artifact_stream", write)
    return captured


def test_download_streams_ranges_from_the_transfer_url_without_authorization(
    writer: list[bytes], monkeypatch: pytest.MonkeyPatch
) -> None:
    from files import config

    monkeypatch.setattr(config, "DOWNLOAD_RANGE_BYTES", 4)
    content = b"CMR-0042-scan"
    harness = _download_harness(content)
    response = harness.run(
        "download_file", {"item_path": "/Inbound/cmr.pdf"}, connection=APP_CONNECTION
    )
    result = _ok(response)
    assert writer == [content]
    assert result["artifact_handle"] == "artifact:downloaded"
    assert result["filename"] == "cmr.pdf"
    assert result["size"] == len(content)
    assert result["sha256"] == hashlib.sha256(content).hexdigest()
    assert result["item_id"] == "01CMR" and result["drive_id"] == DRIVE_ID
    transfers = harness.graph.requests("GET", "tempauth=SECRET-DL")
    assert [sent.headers["range"] for sent in transfers] == [
        "bytes=0-3",
        "bytes=4-7",
        "bytes=8-11",
        "bytes=12-12",
    ]
    assert all("authorization" not in sent.headers for sent in transfers)
    assert all(call[3] is False for call in harness.calls if "tempauth" in call[1])
    _secrets_absent(response)


def test_download_refuses_a_transfer_host_outside_the_egress_list(writer: list[bytes]) -> None:
    harness = _download_harness(b"x", url="https://downloads.evil.example/f?tempauth=SECRET-DL")
    response = harness.run(
        "download_file", {"item_path": "/Inbound/cmr.pdf"}, connection=APP_CONNECTION
    )
    assert response["error_code"] == "egress_host_not_allowed"
    assert not harness.graph.requests(part="evil.example")
    _secrets_absent(response)


def test_download_larger_than_the_grant_is_refused_before_transfer(
    writer: list[bytes], monkeypatch: pytest.MonkeyPatch
) -> None:
    from files import transfer

    monkeypatch.setattr(transfer, "output_grant_limit", lambda _payload: 3)
    harness = _download_harness(b"four")
    response = harness.run(
        "download_file", {"item_path": "/Inbound/cmr.pdf"}, connection=APP_CONNECTION
    )
    assert response["error_code"] == "file_too_large"
    assert not harness.graph.requests(part="tempauth")


def test_download_without_output_grant_fails_before_any_token() -> None:
    harness = Harness()
    response = harness.run("download_file", {"item_id": "01CMR"}, connection=APP_CONNECTION)
    assert response["error_code"] == "artifact_output_unavailable"
    assert harness.transports == []


def test_download_of_a_folder_is_refused(writer: list[bytes]) -> None:
    harness = Harness()
    harness.graph.on(
        "GET", "/me/drive/items/01DIR", (200, drive_item("01DIR", "Inbound", folder=True))
    )
    assert harness.run("download_file", {"item_id": "01DIR"})["error_code"] == "invalid_payload"


def test_short_download_range_is_an_incomplete_download(
    writer: list[bytes], monkeypatch: pytest.MonkeyPatch
) -> None:
    harness = Harness()
    meta = drive_item("01CMR", "cmr.pdf", size=10)
    meta["@microsoft.graph.downloadUrl"] = DOWNLOAD_URL
    harness.graph.on("GET", "/me/drive/items/01CMR", (200, meta))
    harness.graph.on("GET", DOWNLOAD_URL, (206, b"short"))
    response = harness.run("download_file", {"item_id": "01CMR"})
    assert response["error_code"] == "provider_unavailable"
    _secrets_absent(response)


# -- upload: helpers ----------------------------------------------------------


@pytest.fixture
def artifact_source(monkeypatch: pytest.MonkeyPatch):
    from files import transfer

    holder: dict[str, bytes] = {"content": b""}

    def stream(_payload: Any, **kwargs: Any):
        assert kwargs["binding_key"] == "files_upload_artifact"
        assert kwargs["artifact_id"] == "artifact:cmr-scan"
        data = holder["content"]
        for offset in range(0, len(data), 3):
            yield data[offset : offset + 3]

    monkeypatch.setattr(transfer, "stream_artifact_bytes", stream)
    return holder


def _upload(
    harness: Harness, content: bytes, *, behavior: str | None = None, folder: str = "/Inbound"
) -> dict[str, Any]:
    operation_input: dict[str, Any] = {
        "artifact_handle": "artifact:cmr-scan",
        "folder_path": folder,
        "filename": "cmr 0042.pdf",
    }
    if behavior:
        operation_input["conflict_behavior"] = behavior
    return harness.run(
        "upload_file",
        operation_input,
        connection=APP_CONNECTION,
        artifacts=upload_artifacts(len(content)),
    )


TARGET = f"{APP_BASE}/root:/Inbound/cmr%200042.pdf:"


def _uploaded(content: bytes, **extra: Any) -> dict[str, Any]:
    return drive_item(
        "01UP",
        "cmr 0042.pdf",
        size=len(content),
        parent_path=f"{APP_BASE}/root:/Inbound",
        quick_xor=quick_xor_hash(content),
        **extra,
    )


# -- upload: simple PUT -------------------------------------------------------


def test_small_upload_is_one_put_with_conflict_behavior_and_no_retry(artifact_source) -> None:
    content = b"%PDF-acme"
    artifact_source["content"] = content
    harness = Harness()
    harness.graph.on("PUT", f"{TARGET}/content", (201, _uploaded(content)))
    response = _upload(harness, content, behavior="replace")
    result = _ok(response)
    assert response["external_effect_status"] == "succeeded"
    assert result["item_id"] == "01UP" and result["drive_id"] == DRIVE_ID
    assert result["path"] == "/Inbound/cmr 0042.pdf"
    assert result["sha256"] == hashlib.sha256(content).hexdigest()
    assert result["confirmed_after_timeout"] is False
    put = harness.graph.requests("PUT")[0]
    assert "%40microsoft.graph.conflictBehavior=replace" in put.url
    assert put.body == content
    assert put.headers["content-type"] == "application/pdf"
    assert put.headers["authorization"].startswith("Bearer ")
    assert [call[2] for call in harness.calls if call[0] != "GET"] == ["none"]


def test_default_conflict_behavior_is_fail_and_conflict_is_definite(artifact_source) -> None:
    content = b"%PDF-acme"
    artifact_source["content"] = content
    harness = Harness()
    harness.graph.on("PUT", f"{TARGET}/content", graph_error(409, "nameAlreadyExists"))
    response = _upload(harness, content)
    assert response["error_code"] == "name_conflict"
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True
    assert "conflictBehavior=fail" in harness.graph.requests("PUT")[0].url


def test_ambiguous_simple_upload_is_confirmed_by_path_size_and_quick_xor(artifact_source) -> None:
    content = b"%PDF-acme"
    artifact_source["content"] = content
    harness = Harness()
    harness.graph.on("PUT", f"{TARGET}/content", TimeoutError("lost response"))
    harness.graph.on("GET", TARGET, (200, _uploaded(content)))
    response = _upload(harness, content)
    result = _ok(response)
    assert result["confirmed_after_timeout"] is True
    assert response["external_effect_status"] == "succeeded"
    assert len(harness.graph.requests("PUT")) == 1


@pytest.mark.parametrize(
    "found",
    [
        graph_error(404, "itemNotFound"),
        (200, None),  # replaced below with a different hash
    ],
)
def test_unconfirmed_ambiguous_upload_is_timeout_unknown(artifact_source, found) -> None:
    content = b"%PDF-acme"
    artifact_source["content"] = content
    harness = Harness()
    harness.graph.on("PUT", f"{TARGET}/content", graph_error(504))
    if found[1] is None:
        found = (200, _uploaded(b"%PDF-othr"))
    harness.graph.on("GET", TARGET, found)
    response = _upload(harness, content, behavior="replace")
    assert response["error_code"] == "timeout_unknown"
    assert response["external_effect_status"] == "timeout_unknown"
    assert response["definitely_no_external_effect"] is False
    assert len(harness.graph.requests("PUT")) == 1
    assert len(harness.graph.requests("GET", TARGET)) == 3


def test_ambiguous_upload_falls_back_to_sha256_when_graph_reports_it(artifact_source) -> None:
    content = b"%PDF-acme"
    artifact_source["content"] = content
    harness = Harness()
    harness.graph.on("PUT", f"{TARGET}/content", graph_error(503))
    found = drive_item(
        "01UP",
        "cmr 0042.pdf",
        size=len(content),
        sha256=hashlib.sha256(content).hexdigest().upper(),
    )
    harness.graph.on("GET", TARGET, (200, found))
    assert _ok(_upload(harness, content))["confirmed_after_timeout"] is True


def test_ambiguous_rename_upload_is_always_timeout_unknown(artifact_source) -> None:
    content = b"%PDF-acme"
    artifact_source["content"] = content
    harness = Harness()
    harness.graph.on("PUT", f"{TARGET}/content", TimeoutError("lost response"))
    response = _upload(harness, content, behavior="rename")
    assert response["error_code"] == "timeout_unknown"
    assert response["external_effect_status"] == "timeout_unknown"
    assert not harness.graph.requests("GET")


def test_missing_destination_folder_is_folder_not_found(artifact_source) -> None:
    content = b"%PDF-acme"
    artifact_source["content"] = content
    harness = Harness()
    harness.graph.on("PUT", f"{TARGET}/content", graph_error(404, "itemNotFound"))
    response = _upload(harness, content)
    assert response["error_code"] == "folder_not_found"
    assert response["definitely_no_external_effect"] is True


def test_upload_without_artifact_grant_fails_before_any_token() -> None:
    harness = Harness()
    response = harness.run(
        "upload_file",
        {"artifact_handle": "artifact:cmr-scan", "filename": "cmr.pdf"},
        connection=APP_CONNECTION,
    )
    assert response["error_code"] == "artifact_input_unavailable"
    assert harness.transports == []


# -- upload: session ----------------------------------------------------------


@pytest.fixture
def small_session(monkeypatch: pytest.MonkeyPatch) -> None:
    from files import config

    monkeypatch.setattr(config, "SIMPLE_UPLOAD_MAX_BYTES", 4)
    monkeypatch.setattr(config, "UPLOAD_CHUNK_BYTES", 8)


def test_production_fragment_size_is_a_multiple_of_320_kib() -> None:
    from files import config

    assert config.UPLOAD_CHUNK_BYTES % (320 * 1024) == 0
    assert config.UPLOAD_CHUNK_BYTES <= 60 * 1024 * 1024
    assert config.SIMPLE_UPLOAD_MAX_BYTES == 4 * 1024 * 1024


def _session_harness(content: bytes, *fragment_answers: Any) -> Harness:
    harness = Harness()
    harness.graph.on("POST", f"{TARGET}/createUploadSession", (200, {"uploadUrl": UPLOAD_URL}))
    harness.graph.on("PUT", UPLOAD_URL, *fragment_answers)
    return harness


def test_large_upload_uses_a_session_with_unauthorized_ranged_fragments(
    artifact_source, small_session
) -> None:
    content = b"0123456789abcdefXYZ"  # 19 bytes → fragments of 8, 8, 3
    artifact_source["content"] = content
    harness = _session_harness(
        content,
        (202, {"nextExpectedRanges": ["8-"]}),
        (202, {"nextExpectedRanges": ["16-"]}),
        (201, _uploaded(content)),
    )
    response = _upload(harness, content, behavior="fail")
    result = _ok(response)
    assert result["item_id"] == "01UP"
    session = harness.graph.requests("POST")[0]
    assert session.json() == {
        "item": {"@microsoft.graph.conflictBehavior": "fail", "name": "cmr 0042.pdf"}
    }
    fragments = harness.graph.requests("PUT")
    assert [sent.headers["content-range"] for sent in fragments] == [
        "bytes 0-7/19",
        "bytes 8-15/19",
        "bytes 16-18/19",
    ]
    assert b"".join(sent.body or b"" for sent in fragments) == content
    assert all("authorization" not in sent.headers for sent in fragments)
    assert {call[2] for call in harness.calls if call[0] in {"POST", "PUT"}} == {"none"}
    _secrets_absent(response)


def test_ambiguous_middle_fragment_resumes_from_session_status(
    artifact_source, small_session
) -> None:
    content = b"0123456789abcdefXYZ"
    artifact_source["content"] = content
    harness = _session_harness(
        content,
        TimeoutError("lost"),
        (202, {"nextExpectedRanges": ["8-"]}),
        (202, {"nextExpectedRanges": ["16-"]}),
        (201, _uploaded(content)),
    )
    harness.graph.on("GET", UPLOAD_URL, (200, {"nextExpectedRanges": ["0-"]}))
    result = _ok(_upload(harness, content))
    assert result["confirmed_after_timeout"] is False
    ranges = [sent.headers["content-range"] for sent in harness.graph.requests("PUT")]
    assert ranges == ["bytes 0-7/19", "bytes 0-7/19", "bytes 8-15/19", "bytes 16-18/19"]


def test_middle_fragment_failure_cancels_the_session_and_is_definite(
    artifact_source, small_session
) -> None:
    content = b"0123456789abcdefXYZ"
    artifact_source["content"] = content
    harness = _session_harness(
        content, (202, {"nextExpectedRanges": ["8-"]}), graph_error(416, "invalidRange")
    )
    harness.graph.on("DELETE", UPLOAD_URL, (204, b""))
    response = _upload(harness, content)
    assert response["ok"] is False
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True
    cancel = harness.graph.requests("DELETE")
    assert len(cancel) == 1 and "authorization" not in cancel[0].headers
    _secrets_absent(response)


def test_ambiguous_last_fragment_is_confirmed_by_reading_the_target(
    artifact_source, small_session
) -> None:
    content = b"0123456789abcdefXYZ"
    artifact_source["content"] = content
    harness = _session_harness(
        content,
        (202, {"nextExpectedRanges": ["8-"]}),
        (202, {"nextExpectedRanges": ["16-"]}),
        graph_error(504),
    )
    harness.graph.on("GET", TARGET, (200, _uploaded(content)))
    result = _ok(_upload(harness, content, behavior="replace"))
    assert result["confirmed_after_timeout"] is True
    assert not harness.graph.requests("DELETE")


def test_ambiguous_last_fragment_with_rename_is_timeout_unknown(
    artifact_source, small_session
) -> None:
    content = b"0123456789abcdefXYZ"
    artifact_source["content"] = content
    harness = _session_harness(
        content,
        (202, {"nextExpectedRanges": ["8-"]}),
        (202, {"nextExpectedRanges": ["16-"]}),
        TimeoutError("lost"),
    )
    response = _upload(harness, content, behavior="rename")
    assert response["external_effect_status"] == "timeout_unknown"
    assert response["definitely_no_external_effect"] is False
    assert not harness.graph.requests("GET")
    assert not harness.graph.requests("DELETE")


def test_ambiguous_session_start_is_a_definite_failure(artifact_source, small_session) -> None:
    content = b"0123456789abcdefXYZ"
    artifact_source["content"] = content
    harness = Harness()
    harness.graph.on("POST", f"{TARGET}/createUploadSession", TimeoutError("lost"))
    response = _upload(harness, content)
    assert response["external_effect_status"] == "failed"
    assert response["definitely_no_external_effect"] is True
    assert not harness.graph.requests("PUT")


def test_upload_url_on_a_foreign_host_is_refused_before_any_bytes(
    artifact_source, small_session
) -> None:
    content = b"0123456789abcdefXYZ"
    artifact_source["content"] = content
    harness = Harness()
    harness.graph.on(
        "POST",
        f"{TARGET}/createUploadSession",
        (200, {"uploadUrl": "https://upload.evil.example/s?tempauth=SECRET-UP"}),
    )
    response = _upload(harness, content)
    assert response["error_code"] == "egress_host_not_allowed"
    assert response["definitely_no_external_effect"] is True
    assert not harness.graph.requests(part="evil.example")
    _secrets_absent(response)


def test_short_artifact_stream_never_sends_the_last_fragment(
    artifact_source, small_session, monkeypatch: pytest.MonkeyPatch
) -> None:
    content = b"0123456789abcdefXYZ"
    artifact_source["content"] = content[:12]  # the grant says 19 bytes
    harness = _session_harness(content, (202, {"nextExpectedRanges": ["8-"]}))
    harness.graph.on("DELETE", UPLOAD_URL, (204, b""))
    response = harness.run(
        "upload_file",
        {
            "artifact_handle": "artifact:cmr-scan",
            "folder_path": "/Inbound",
            "filename": "cmr 0042.pdf",
        },
        connection=APP_CONNECTION,
        artifacts=upload_artifacts(len(content)),
    )
    assert response["error_code"] == "artifact_input_unavailable"
    assert response["definitely_no_external_effect"] is True
    assert len(harness.graph.requests("PUT")) == 1
