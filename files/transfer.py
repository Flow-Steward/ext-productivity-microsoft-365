"""File transfers: download to an artifact, upload from an artifact.

Pre-authenticated transfer URLs (``@microsoft.graph.downloadUrl`` and an upload
session's ``uploadUrl``) are called with ``authorized=False``, so the Graph token
never reaches them. They are kept in local variables only: never logged, never put
in a result or an error.

Upload outcome rules (task B5.3):

- every upload request runs with ``retry="none"``;
- a fragment before the last one cannot create the file, so a failure there cancels
  the session and is a definite failure;
- an ambiguous result of the request that completes the file (the simple ``PUT`` or
  the last fragment) is confirmed by reading the target path: the item must have the
  artifact's size and hash (``quickXorHash``, or ``sha256Hash`` when Graph reports
  it). Anything else is ``timeout_unknown``. With ``conflict_behavior: rename`` the
  final name is unknown, so an ambiguous result is always ``timeout_unknown``.
"""

from __future__ import annotations

import hashlib
import os
import re
import time
from collections.abc import Callable, Iterable, Iterator, Mapping
from dataclasses import dataclass, field
from typing import Any

from flowsteward_extension_sdk import (
    ArtifactAccessError,
    PinnedPeerError,
    find_artifact_descriptor,
    stream_artifact_bytes,
    write_artifact_stream,
)
from flowsteward_extension_sdk.http import verified_platform_grant_signature

from m365.connection import mapping, text
from m365.errors import M365Error
from m365.transport import GraphTransport

from . import config, paths
from .drives import DriveTarget
from .items import ITEM_SELECT, project_item
from .progress import TransferProgress, transfer_progress
from .quickxor import QuickXorHash

UPLOAD_BINDING_KEY = "files_upload_artifact"
DOWNLOAD_BINDING_KEY = "files_download_artifact"
GRANT_CONTENT_TYPE = "application/octet-stream"
CONFLICT_BEHAVIORS = ("fail", "rename", "replace")
CONFIRM_ATTEMPTS = 3
CONFIRM_INTERVAL_SECONDS = 2.0
_MIME = re.compile(r"^[A-Za-z0-9!#$&^_.+-]{1,127}/[A-Za-z0-9!#$&^_.+-]{1,127}$")
_OUTPUT_UNAVAILABLE = "The approved artifact output is unavailable"
_INPUT_UNAVAILABLE = "The approved input artifact is unavailable"

#: Patched in tests; real runs wait between confirmation reads.
sleep: Callable[[float], None] = time.sleep


# -- artifact grants ----------------------------------------------------------


def scope_matches(payload: Mapping[str, Any], descriptor: Mapping[str, Any]) -> bool:
    host_scope = mapping(mapping(payload.get("host")).get("scope"))
    grant_scope = mapping(descriptor.get("scope"))
    for key in ("account_id", "project_id"):
        expected = str(host_scope.get(key) or "").strip()
        if not expected or str(grant_scope.get(key) or "").strip() != expected:
            return False
    return True


def output_grant_limit(payload: dict[str, Any]) -> int:
    """The signed output grant's byte limit; any doubt means no output."""
    try:
        descriptor = find_artifact_descriptor(
            payload, binding_key=DOWNLOAD_BINDING_KEY, role="output"
        )
        access = descriptor.get("access")
        max_size = access.get("max_size_bytes") if isinstance(access, Mapping) else None
        if (
            descriptor.get("role") != "output"
            or descriptor.get("binding_key") != DOWNLOAD_BINDING_KEY
            or descriptor.get("content_type") != GRANT_CONTENT_TYPE
            or not isinstance(access, Mapping)
            or access.get("transport") != "presigned_url"
            or access.get("mode") not in {"write", "read_write"}
            or isinstance(max_size, bool)
            or not isinstance(max_size, int)
            or max_size <= 0
            or not verified_platform_grant_signature(descriptor, dict(access))
            or not scope_matches(payload, descriptor)
        ):
            raise ArtifactAccessError("invalid output grant")
        return max_size
    except (ArtifactAccessError, TypeError, ValueError) as exc:
        raise M365Error("artifact_output_unavailable", _OUTPUT_UNAVAILABLE) from exc


@dataclass(frozen=True)
class InputArtifact:
    handle: str
    size: int
    mime_type: str


def input_artifact(payload: dict[str, Any], handle: str) -> InputArtifact:
    try:
        descriptor = find_artifact_descriptor(
            payload, artifact_id=handle, binding_key=UPLOAD_BINDING_KEY, role="input"
        )
        access = descriptor.get("access")
        size = descriptor.get("size_bytes")
        max_size = access.get("max_size_bytes") if isinstance(access, Mapping) else None
        mime_type = descriptor.get("content_type")
        if (
            descriptor.get("role") != "input"
            or descriptor.get("binding_key") != UPLOAD_BINDING_KEY
            or not isinstance(access, Mapping)
            or access.get("transport") != "presigned_url"
            or access.get("mode") not in {"read", "read_write"}
            or isinstance(size, bool)
            or not isinstance(size, int)
            or size < 0
            or isinstance(max_size, bool)
            or not isinstance(max_size, int)
            or max_size <= 0
            or size > max_size
            or not isinstance(mime_type, str)
            or not _MIME.match(mime_type)
            or not scope_matches(payload, descriptor)
        ):
            raise ArtifactAccessError("invalid input grant")
    except (ArtifactAccessError, TypeError, ValueError) as exc:
        raise M365Error("artifact_input_unavailable", _INPUT_UNAVAILABLE) from exc
    if size > config.artifact_max_bytes():
        raise M365Error("file_too_large", "The file is larger than this server allows")
    return InputArtifact(handle=handle, size=size, mime_type=mime_type)


def safe_filename(value: str) -> str:
    name = os.path.basename(value.replace("\\", "/"))
    name = re.sub(r"[\x00-\x1f\x7f]", "", name).strip().strip(".")
    normalized = name.encode("utf-8")[:255].decode("utf-8", errors="ignore").strip()
    return normalized or "download.bin"


# -- download -----------------------------------------------------------------


def download_file(
    transport: GraphTransport,
    payload: dict[str, Any],
    target: DriveTarget,
    *,
    item: str,
    path: str,
) -> dict[str, Any]:
    limit = min(config.artifact_max_bytes(), output_grant_limit(payload))
    # No $select: the pre-authenticated download URL is an instance annotation.
    try:
        raw = transport.get_json(
            paths.item_ref(target.base, item=item, path=path),
            purpose="Microsoft Graph file metadata",
        )
    except M365Error as exc:
        raise paths.item_missing(exc) from exc
    projected = project_item(raw)
    file_facet = mapping(raw.get("file"))
    if projected["is_folder"] or not file_facet:
        raise M365Error("invalid_payload", "The item is a folder, not a file")
    size = projected["size"]
    if size is None or size < 0:
        raise M365Error("provider_unavailable", "Microsoft Graph omitted the file size")
    if size > limit:
        raise M365Error("file_too_large", "The file is larger than the artifact limit")
    download_url = raw.get("@microsoft.graph.downloadUrl")
    if not isinstance(download_url, str) or not download_url.startswith("https://"):
        raise M365Error("provider_unavailable", "Microsoft Graph returned no download location")
    filename = safe_filename(projected["name"])
    mime_type = projected["mime_type"] or GRANT_CONTENT_TYPE
    if not _MIME.match(mime_type):
        mime_type = GRANT_CONTENT_TYPE
    state: dict[str, Any] = {"size": 0, "error": None}
    progress = transfer_progress("Downloading", filename, total_bytes=size)
    chunks = _ranged_download(transport, download_url, size=size, state=state, progress=progress)
    try:
        written = write_artifact_stream(
            payload, chunks, binding_key=DOWNLOAD_BINDING_KEY, content_type=GRANT_CONTENT_TYPE
        )
    except M365Error:
        raise
    except Exception as exc:
        if isinstance(state["error"], M365Error):
            raise state["error"] from exc
        raise M365Error("artifact_output_unavailable", _OUTPUT_UNAVAILABLE) from exc
    if isinstance(state["error"], M365Error):
        raise state["error"]
    handle = written.get("artifact_handle") if isinstance(written, Mapping) else None
    written_size = written.get("size_bytes") if isinstance(written, Mapping) else None
    sha256 = written.get("sha256") if isinstance(written, Mapping) else None
    if (
        not isinstance(handle, str)
        or not handle.strip()
        or written_size != state["size"]
        or (sha256 is not None and not isinstance(sha256, str))
    ):
        raise M365Error("artifact_output_unavailable", _OUTPUT_UNAVAILABLE)
    return {
        "artifact_handle": handle.strip(),
        "filename": filename,
        "mime_type": mime_type,
        "size": state["size"],
        "sha256": sha256 or None,
        "item_id": projected["item_id"],
        "drive_id": projected["drive_id"] or target.drive_id or None,
        "modified_at": projected["modified_at"],
        "web_url": projected["web_url"],
    }


def _ranged_download(
    transport: GraphTransport,
    url: str,
    *,
    size: int,
    state: dict[str, Any],
    progress: TransferProgress,
) -> Iterator[bytes]:
    try:
        start = 0
        while start < size:
            end = min(start + config.DOWNLOAD_RANGE_BYTES, size) - 1
            response = transport.request(
                "GET",
                url,
                headers={"Range": f"bytes={start}-{end}"},
                response_limit=config.DOWNLOAD_RANGE_BYTES,
                retry="read",
                purpose="Microsoft 365 file download",
                authorized=False,
                ok_statuses=(200, 206),
            )
            body = response.body
            if response.status == 200:
                # The server ignored the range: only a complete small file is acceptable.
                if start != 0 or len(body) != size:
                    raise M365Error("provider_unavailable", "The file download was incomplete")
            elif len(body) != end - start + 1:
                raise M365Error("provider_unavailable", "The file download was incomplete")
            state["size"] += len(body)
            progress(state["size"], False)
            yield body
            if response.status == 200:
                break
            start = end + 1
        if state["size"] != size:
            raise M365Error("provider_unavailable", "The file download was incomplete")
        progress(state["size"], True)
    except M365Error as exc:
        state["error"] = exc
        raise


# -- upload -------------------------------------------------------------------


@dataclass
class UploadDigest:
    size: int = 0
    _sha256: Any = field(default_factory=hashlib.sha256)
    _quick: QuickXorHash = field(default_factory=QuickXorHash)

    def update(self, chunk: bytes) -> None:
        self.size += len(chunk)
        self._sha256.update(chunk)
        self._quick.update(chunk)

    @property
    def sha256(self) -> str:
        return str(self._sha256.hexdigest())

    @property
    def quick_xor(self) -> str:
        return self._quick.b64digest()


@dataclass(frozen=True)
class UploadTarget:
    target: DriveTarget
    folder_item: str
    folder_path: str
    filename: str
    conflict_behavior: str

    def child(self) -> str:
        return paths.child_ref(
            self.target.base,
            parent_item=self.folder_item,
            parent_path=self.folder_path,
            child=self.filename,
        )


def upload_file(
    transport: GraphTransport,
    payload: dict[str, Any],
    destination: UploadTarget,
    artifact: InputArtifact,
) -> dict[str, Any]:
    digest = UploadDigest()
    progress = transfer_progress("Uploading", destination.filename, total_bytes=artifact.size)
    source = _artifact_chunks(payload, artifact)
    if artifact.size <= config.SIMPLE_UPLOAD_MAX_BYTES:
        return _simple_upload(transport, destination, artifact, source, digest, progress)
    return _session_upload(transport, destination, artifact, source, digest, progress)


def _artifact_chunks(payload: dict[str, Any], artifact: InputArtifact) -> Iterator[bytes]:
    try:
        stream = stream_artifact_bytes(
            payload,
            artifact_id=artifact.handle,
            binding_key=UPLOAD_BINDING_KEY,
            chunk_size=config.ARTIFACT_READ_CHUNK_BYTES,
            timeout_seconds=30.0,
        )
        total = 0
        for raw in stream:
            if not isinstance(raw, (bytes, bytearray, memoryview)):
                raise M365Error(
                    "artifact_input_unavailable", "The input artifact stream is invalid"
                )
            chunk = bytes(raw)
            total += len(chunk)
            if total > artifact.size:
                raise M365Error(
                    "artifact_input_unavailable", "The input artifact exceeds its signed size"
                )
            if chunk:
                yield chunk
        if total != artifact.size:
            raise M365Error("artifact_input_unavailable", "The input artifact stream is incomplete")
    except (ArtifactAccessError, PinnedPeerError, OSError) as exc:
        raise M365Error("artifact_input_unavailable", _INPUT_UNAVAILABLE) from exc


def _regroup(chunks: Iterable[bytes], size: int) -> Iterator[bytes]:
    buffer = bytearray()
    for chunk in chunks:
        buffer.extend(chunk)
        while len(buffer) >= size:
            yield bytes(buffer[:size])
            del buffer[:size]
    if buffer:
        yield bytes(buffer)


def _simple_upload(
    transport: GraphTransport,
    destination: UploadTarget,
    artifact: InputArtifact,
    source: Iterator[bytes],
    digest: UploadDigest,
    progress: TransferProgress,
) -> dict[str, Any]:
    body = b"".join(source)
    digest.update(body)
    try:
        response = transport.request(
            "PUT",
            f"{destination.child()}/content",
            query={"@microsoft.graph.conflictBehavior": destination.conflict_behavior},
            raw_body=body,
            content_type=artifact.mime_type,
            retry="none",
            purpose="Microsoft Graph file upload",
            ok_statuses=(200, 201),
        )
    except M365Error as exc:
        if exc.ambiguous:
            return _confirm_ambiguous(transport, destination, digest, exc)
        raise _remap_upload_error(exc) from exc
    progress(digest.size, True)
    return _upload_result(response.json(), destination, digest, confirmed_after_timeout=False)


def _session_upload(
    transport: GraphTransport,
    destination: UploadTarget,
    artifact: InputArtifact,
    source: Iterator[bytes],
    digest: UploadDigest,
    progress: TransferProgress,
) -> dict[str, Any]:
    try:
        session = transport.request(
            "POST",
            f"{destination.child()}/createUploadSession",
            json_body={
                "item": {
                    "@microsoft.graph.conflictBehavior": destination.conflict_behavior,
                    "name": destination.filename,
                }
            },
            retry="none",
            purpose="Microsoft Graph upload session",
            ok_statuses=(200, 201),
        ).json()
    except M365Error as exc:
        if exc.ambiguous:
            # A session holds no file; nothing exists until the last fragment arrives.
            raise M365Error(
                "provider_unavailable",
                "Microsoft 365 did not start the upload session",
                http_status=exc.http_status,
            ) from exc
        raise _remap_upload_error(exc) from exc
    upload_url = session.get("uploadUrl")
    if not isinstance(upload_url, str) or not upload_url.startswith("https://"):
        raise M365Error("provider_unavailable", "Microsoft Graph returned no upload location")
    offset = 0
    final_response = None
    try:
        for fragment in _regroup(source, config.UPLOAD_CHUNK_BYTES):
            digest.update(fragment)
            end = offset + len(fragment) - 1
            is_last = end + 1 == artifact.size
            try:
                response = _put_fragment(transport, upload_url, fragment, offset, artifact.size)
            except M365Error as exc:
                if is_last:
                    if exc.ambiguous:
                        return _confirm_ambiguous(transport, destination, digest, exc)
                    raise _remap_upload_error(exc) from exc
                if not exc.ambiguous:
                    raise _definite_failure(exc) from exc
                response = _resume_fragment(
                    transport, upload_url, fragment, offset, artifact.size, exc
                )
            offset = end + 1
            progress(offset, False)
            if is_last:
                final_response = response
        if final_response is None:
            raise M365Error("artifact_input_unavailable", "The input artifact stream is incomplete")
        if final_response.status not in {200, 201}:
            ambiguous = M365Error(
                "timeout_unknown", "Microsoft 365 did not confirm the upload", ambiguous=True
            )
            return _confirm_ambiguous(transport, destination, digest, ambiguous)
    except M365Error as exc:
        if not exc.ambiguous:
            _cancel_session(transport, upload_url)
        raise
    progress(digest.size, True)
    return _upload_result(final_response.json(), destination, digest, confirmed_after_timeout=False)


def _put_fragment(
    transport: GraphTransport, url: str, fragment: bytes, offset: int, total: int
) -> Any:
    end = offset + len(fragment) - 1
    return transport.request(
        "PUT",
        url,
        raw_body=fragment,
        content_type="application/octet-stream",
        headers={"Content-Range": f"bytes {offset}-{end}/{total}"},
        retry="none",
        purpose="Microsoft 365 upload fragment",
        authorized=False,
        ok_statuses=(200, 201, 202),
    )


def _resume_fragment(
    transport: GraphTransport,
    url: str,
    fragment: bytes,
    offset: int,
    total: int,
    original: M365Error,
) -> Any:
    """An ambiguous middle fragment: ask the session what it holds, resend at most once."""
    expected = _next_expected(transport, url)
    end = offset + len(fragment) - 1
    if expected == end + 1:
        return None  # the fragment landed
    if expected != offset:
        raise _definite_failure(original)
    try:
        return _put_fragment(transport, url, fragment, offset, total)
    except M365Error as exc:
        raise _definite_failure(exc) from exc


def _next_expected(transport: GraphTransport, url: str) -> int | None:
    try:
        status = transport.request(
            "GET",
            url,
            retry="read",
            purpose="Microsoft 365 upload session status",
            authorized=False,
            ok_statuses=(200,),
        ).json()
    except M365Error:
        return None
    ranges = status.get("nextExpectedRanges")
    if not isinstance(ranges, list) or not ranges or not isinstance(ranges[0], str):
        return None
    start = ranges[0].split("-", 1)[0].strip()
    return int(start) if start.isdigit() else None


def _cancel_session(transport: GraphTransport, url: str) -> None:
    try:
        transport.request(
            "DELETE",
            url,
            retry="none",
            purpose="Microsoft 365 upload session cancel",
            authorized=False,
            ok_statuses=(200, 202, 204),
        )
    except M365Error:
        pass  # an abandoned session expires on its own and never becomes a file


def _definite_failure(exc: M365Error) -> M365Error:
    """A failure before the last fragment: the file was not created."""
    if not exc.ambiguous:
        return _remap_upload_error(exc)
    return M365Error(
        "provider_unavailable",
        "The upload was interrupted before the file was created",
        http_status=exc.http_status,
    )


def _remap_upload_error(exc: M365Error) -> M365Error:
    if exc.code in paths.NOT_FOUND_CODES:
        return M365Error(
            "folder_not_found", "The destination folder was not found", http_status=exc.http_status
        )
    return exc


def _confirm_ambiguous(
    transport: GraphTransport,
    destination: UploadTarget,
    digest: UploadDigest,
    original: M365Error,
) -> dict[str, Any]:
    """Positive confirmation by target path, size and hash; otherwise ``timeout_unknown``."""
    if destination.conflict_behavior != "rename":
        for attempt in range(CONFIRM_ATTEMPTS):
            try:
                raw = transport.get_json(
                    destination.child(),
                    query={"$select": ITEM_SELECT},
                    purpose="Microsoft Graph upload confirmation",
                )
            except M365Error:
                raw = None
            if raw is not None and content_matches(raw, digest):
                return _upload_result(raw, destination, digest, confirmed_after_timeout=True)
            if attempt + 1 < CONFIRM_ATTEMPTS and transport.remaining_seconds() > 10.0:
                sleep(CONFIRM_INTERVAL_SECONDS)
    raise M365Error(
        "timeout_unknown",
        "Microsoft 365 did not confirm the upload; check the folder before running it again",
        http_status=original.http_status,
        ambiguous=True,
    )


def content_matches(raw: Mapping[str, Any], digest: UploadDigest) -> bool:
    size = raw.get("size")
    if isinstance(size, bool) or size != digest.size:
        return False
    hashes = mapping(mapping(raw.get("file")).get("hashes"))
    quick = text(hashes.get("quickXorHash"), limit=128)
    if quick:
        return quick == digest.quick_xor
    sha256 = text(hashes.get("sha256Hash"), limit=128)
    if sha256:
        return sha256.lower() == digest.sha256
    return False


def _upload_result(
    raw: Any,
    destination: UploadTarget,
    digest: UploadDigest,
    *,
    confirmed_after_timeout: bool,
) -> dict[str, Any]:
    item = project_item(raw)
    return {
        "item_id": item["item_id"],
        "drive_id": item["drive_id"] or destination.target.drive_id or None,
        "name": item["name"],
        "path": item["path"],
        "size": item["size"] if item["size"] is not None else digest.size,
        "mime_type": item["mime_type"],
        "web_url": item["web_url"],
        "quick_xor_hash": item["quick_xor_hash"],
        "sha256": digest.sha256,
        "conflict_behavior": destination.conflict_behavior,
        "confirmed_after_timeout": confirmed_after_timeout,
        "changed": True,
    }


__all__ = [
    "CONFLICT_BEHAVIORS",
    "DOWNLOAD_BINDING_KEY",
    "UPLOAD_BINDING_KEY",
    "InputArtifact",
    "UploadDigest",
    "UploadTarget",
    "content_matches",
    "download_file",
    "input_artifact",
    "output_grant_limit",
    "upload_file",
]
