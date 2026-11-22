"""Outgoing attachments: artifact grants → verified bytes with a SHA-256 per file."""

from __future__ import annotations

import hashlib
import mimetypes
import os
import re
from collections.abc import Callable, Iterable, Iterator, Mapping
from dataclasses import dataclass, field
from typing import Any

from flowsteward_extension_sdk import (
    ArtifactAccessError,
    PinnedPeerError,
    find_artifact_descriptor,
    report_progress,
    stream_artifact_bytes,
)

from m365.errors import M365Error

from .download import safe_filename

BINDING_KEY = "attachments"
#: Graph takes an attachment inline (base64 in one POST) below 3 MB; larger ones need an upload session.
INLINE_LIMIT_BYTES = 3 * 1024 * 1024
MAX_FILE_BYTES = 150 * 1024 * 1024
MAX_TOTAL_BYTES = 150 * 1024 * 1024
MAX_ATTACHMENTS = 20
STREAM_CHUNK_BYTES = 1024 * 1024
_UNAVAILABLE = "An attachment artifact is unavailable"
_MIME = re.compile(r"^[A-Za-z0-9!#$&^_.+-]{1,127}/[A-Za-z0-9!#$&^_.+-]{1,127}$")
_ROW_KEYS = frozenset({"artifact_handle", "artifact_id", "filename", "content_type"})

#: Indirection point for tests.
artifact_streamer: Callable[..., Iterable[bytes]] = stream_artifact_bytes


@dataclass
class OutgoingAttachment:
    handle: str
    filename: str
    content_type: str
    size: int
    sha256: str
    content: bytes | None = field(default=None, repr=False)

    @property
    def inline(self) -> bool:
        return self.size < INLINE_LIMIT_BYTES

    @property
    def match_key(self) -> tuple[str, str]:
        return (self.filename, normalized_content_type(self.content_type))


def normalized_content_type(value: Any) -> str:
    return str(value or "").strip().lower().split(";", 1)[0].strip()


def attachment_rows(value: Any) -> list[dict[str, str]]:
    """Validate the ``attachments`` input: artifact handles or ``{artifact_handle, filename?, content_type?}``."""
    if value is None:
        return []
    if not isinstance(value, list) or len(value) > MAX_ATTACHMENTS:
        raise M365Error(
            "invalid_payload", f"attachments must be a list of at most {MAX_ATTACHMENTS} artifacts"
        )
    rows: list[dict[str, str]] = []
    for raw in value:
        if isinstance(raw, str):
            raw = {"artifact_handle": raw}
        if not isinstance(raw, Mapping) or set(raw) - _ROW_KEYS:
            raise M365Error(
                "invalid_payload",
                "Each attachment is an artifact handle or {artifact_handle, filename, content_type}",
            )
        handle = raw.get("artifact_handle") or raw.get("artifact_id")
        if (
            not isinstance(handle, str)
            or not handle.strip()
            or len(handle) > 512
            or any(ord(c) < 32 for c in handle)
        ):
            raise M365Error("invalid_payload", "An attachment has no valid artifact handle")
        row = {"handle": handle.strip()}
        filename = raw.get("filename")
        if filename is not None:
            if (
                not isinstance(filename, str)
                or filename != os.path.basename(filename.replace("\\", "/"))
                or not 1 <= len(filename.encode("utf-8")) <= 255
                or filename.strip() != filename
                or filename in {".", ".."}
                or any(ord(c) < 32 or ord(c) == 127 for c in filename)
            ):
                raise M365Error(
                    "invalid_payload", "An attachment filename must be a safe file name"
                )
            row["filename"] = filename
        content_type = raw.get("content_type")
        if content_type is not None:
            if not isinstance(content_type, str) or not _MIME.match(content_type.strip()):
                raise M365Error("invalid_payload", "An attachment content_type is invalid")
            row["content_type"] = content_type.strip().lower()
        rows.append(row)
    return rows


def load_attachments(
    payload: dict[str, Any], rows: list[dict[str, str]]
) -> list[OutgoingAttachment]:
    """Read every artifact once: exact size, SHA-256, and the bytes of small files."""
    result: list[OutgoingAttachment] = []
    total = 0
    for index, row in enumerate(rows):
        descriptor = _descriptor(payload, row["handle"])
        size = descriptor["size_bytes"]
        total += size
        if size > MAX_FILE_BYTES or total > MAX_TOTAL_BYTES:
            raise M365Error(
                "attachment_too_large", "The attachments exceed the Outlook message size limit"
            )
        filename = row.get("filename") or safe_filename(
            descriptor.get("filename") or descriptor.get("name")
        )
        content_type = row.get("content_type") or _content_type(
            descriptor.get("content_type"), filename
        )
        report_progress(f"Reading attachment {filename}", done=index, total=len(rows))
        digest = hashlib.sha256()
        kept = bytearray() if size < INLINE_LIMIT_BYTES else None
        for chunk in read_chunks(payload, row["handle"], expected_size=size):
            digest.update(chunk)
            if kept is not None:
                kept.extend(chunk)
        result.append(
            OutgoingAttachment(
                handle=row["handle"],
                filename=filename,
                content_type=content_type,
                size=size,
                sha256=digest.hexdigest(),
                content=bytes(kept) if kept is not None else None,
            )
        )
    return result


def read_chunks(payload: dict[str, Any], handle: str, *, expected_size: int) -> Iterator[bytes]:
    """Stream one artifact and insist on its declared size."""
    seen = 0
    try:
        for raw in artifact_streamer(
            payload, artifact_id=handle, chunk_size=STREAM_CHUNK_BYTES, timeout_seconds=30.0
        ):
            if not isinstance(raw, (bytes, bytearray, memoryview)):
                raise M365Error("artifact_input_unavailable", _UNAVAILABLE)
            chunk = bytes(raw)
            seen += len(chunk)
            if seen > expected_size:
                raise M365Error(
                    "artifact_input_unavailable", "An attachment artifact is larger than declared"
                )
            if chunk:
                yield chunk
    except (ArtifactAccessError, PinnedPeerError, OSError) as exc:
        raise M365Error("artifact_input_unavailable", _UNAVAILABLE) from exc
    if seen != expected_size:
        raise M365Error("artifact_input_unavailable", "An attachment artifact is incomplete")


def _descriptor(payload: dict[str, Any], handle: str) -> dict[str, Any]:
    try:
        # Look up by handle only: the SDK matches a binding key before later handles.
        descriptor = find_artifact_descriptor(payload, artifact_id=handle, role="input")
    except (ArtifactAccessError, TypeError, ValueError) as exc:
        raise M365Error("artifact_input_unavailable", _UNAVAILABLE) from exc
    access = descriptor.get("access")
    size = descriptor.get("size_bytes")
    max_size = access.get("max_size_bytes") if isinstance(access, Mapping) else None
    if (
        descriptor.get("role") not in (None, "input")
        or descriptor.get("binding_key") != BINDING_KEY
        or not isinstance(access, Mapping)
        or access.get("transport") != "presigned_url"
        or access.get("mode") not in {"read", "read_write"}
        or isinstance(size, bool)
        or not isinstance(size, int)
        or size < 0
        or (
            max_size is not None
            and (isinstance(max_size, bool) or not isinstance(max_size, int) or size > max_size)
        )
    ):
        raise M365Error("artifact_input_unavailable", _UNAVAILABLE)
    return descriptor


def _content_type(declared: Any, filename: str) -> str:
    candidate = normalized_content_type(declared)
    if candidate and _MIME.match(candidate) and candidate != "application/octet-stream":
        return candidate
    guessed = mimetypes.guess_type(filename)[0]
    return guessed.lower() if guessed else "application/octet-stream"


__all__ = [
    "BINDING_KEY",
    "INLINE_LIMIT_BYTES",
    "MAX_ATTACHMENTS",
    "OutgoingAttachment",
    "attachment_rows",
    "load_attachments",
    "normalized_content_type",
    "read_chunks",
]
