"""``get_attachment``: one file attachment → an approved artifact output."""

from __future__ import annotations

import hashlib
import os
import re
from collections.abc import Callable, Iterator, Mapping
from typing import Any

from flowsteward_extension_sdk import (
    ArtifactAccessError,
    find_artifact_descriptor,
    report_progress,
    write_artifact_stream,
)
from flowsteward_extension_sdk.http import verified_platform_grant_signature

from m365.errors import M365Error
from m365.runtime import OperationContext

from .common import Mailbox, graph_id, not_found_as, segment
from .reads import attachment_type

BINDING_KEY = "attachment_artifact_handle"
GRANT_CONTENT_TYPE = "application/octet-stream"
#: Outlook's largest message; an attachment cannot be bigger.
MAX_ATTACHMENT_BYTES = 150 * 1024 * 1024
WRITE_CHUNK_BYTES = 1024 * 1024
_DEFAULT_NAME = "attachment.bin"
_OUTPUT_UNAVAILABLE = "The approved artifact output is unavailable"
_MIME = re.compile(r"^[A-Za-z0-9!#$&^_.+-]{1,127}/[A-Za-z0-9!#$&^_.+-]{1,127}$")

#: Indirection points for tests.
artifact_writer: Callable[..., dict[str, Any]] = write_artifact_stream


def get_attachment(ctx: OperationContext) -> dict[str, Any]:
    limit = min(MAX_ATTACHMENT_BYTES, output_grant_limit(ctx.payload))
    message_id = graph_id(ctx.input.get("message_id"), "message_id")
    attachment_id = graph_id(ctx.input.get("attachment_id"), "attachment_id")
    client = Mailbox.for_connection(ctx.transport, ctx.connection)
    path = f"/messages/{segment(message_id)}/attachments/{segment(attachment_id)}"
    report_progress("Reading the attachment")
    with not_found_as("attachment_not_found", "The attachment was not found on this message"):
        meta = client.get(
            path,
            query={"$select": "id,name,contentType,size,isInline"},
            response_limit=64 * 1024,
            purpose="Outlook attachment metadata",
        )
        kind = attachment_type(meta)
        if kind != "file":
            raise M365Error(
                "attachment_type_unsupported",
                "Only file attachments can be downloaded; attached messages and cloud links are not supported",
            )
        declared = meta.get("size")
        if (
            isinstance(declared, int)
            and not isinstance(declared, bool)
            and declared > limit + 64 * 1024
        ):
            # ``size`` includes Exchange property overhead, so allow some slack before refusing.
            raise M365Error(
                "attachment_too_large", "The attachment is larger than the artifact limit"
            )
        filename = safe_filename(meta.get("name"))
        report_progress(f"Downloading {filename}")
        try:
            content = client.get_bytes(
                f"{path}/$value", response_limit=limit, purpose="Outlook attachment content"
            )
        except M365Error as exc:
            if exc.code == "response_too_large":
                raise M365Error(
                    "attachment_too_large", "The attachment is larger than the artifact limit"
                ) from exc
            raise
    digest = hashlib.sha256(content).hexdigest()
    content_type = str(meta.get("contentType") or "").strip().lower().split(";", 1)[0]
    if not _MIME.match(content_type):
        content_type = GRANT_CONTENT_TYPE
    report_progress(f"Saving {filename} ({_size_text(len(content))})")
    try:
        written = artifact_writer(
            ctx.payload,
            _chunks(content),
            binding_key=BINDING_KEY,
            content_type=GRANT_CONTENT_TYPE,
        )
    except Exception as exc:
        raise M365Error("artifact_output_unavailable", _OUTPUT_UNAVAILABLE) from exc
    handle = written.get("artifact_handle") if isinstance(written, Mapping) else None
    if not isinstance(handle, str) or not handle:
        raise M365Error("artifact_output_unavailable", _OUTPUT_UNAVAILABLE)
    return {
        "attachment_artifact_handle": handle,
        "attachment_metadata": {
            "message_id": message_id,
            "attachment_id": attachment_id,
            "original_filename": filename,
            "content_type": content_type,
            "size_bytes": len(content),
            "sha256": digest,
            "is_inline": bool(meta.get("isInline")),
        },
    }


def output_grant_limit(payload: dict[str, Any]) -> int:
    try:
        descriptor = find_artifact_descriptor(payload, binding_key=BINDING_KEY, role="output")
        access = descriptor.get("access")
        if not isinstance(access, dict):
            raise ArtifactAccessError("missing access")
        max_size = access.get("max_size_bytes")
        if (
            descriptor.get("role") != "output"
            or descriptor.get("binding_key") != BINDING_KEY
            or descriptor.get("content_type") != GRANT_CONTENT_TYPE
            or access.get("transport") != "presigned_url"
            or access.get("mode") not in {"write", "read_write"}
            or not isinstance(max_size, int)
            or isinstance(max_size, bool)
            or max_size <= 0
            or not verified_platform_grant_signature(descriptor, access)
        ):
            raise ArtifactAccessError("invalid grant")
        return max_size
    except (ArtifactAccessError, TypeError, ValueError) as exc:
        raise M365Error("artifact_output_unavailable", _OUTPUT_UNAVAILABLE) from exc


def safe_filename(value: Any) -> str:
    name = os.path.basename(str(value or "").replace("\\", "/"))
    name = re.sub(r"[\x00-\x1f\x7f]", "", name).strip().strip(".")
    encoded = name.encode("utf-8")[:240]
    return encoded.decode("utf-8", errors="ignore").strip() or _DEFAULT_NAME


def _chunks(content: bytes) -> Iterator[bytes]:
    for offset in range(0, len(content), WRITE_CHUNK_BYTES):
        yield content[offset : offset + WRITE_CHUNK_BYTES]


def _size_text(size_bytes: int) -> str:
    if size_bytes >= 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    return f"{max(1, round(size_bytes / 1024))} KB"


__all__ = ["BINDING_KEY", "get_attachment", "output_grant_limit", "safe_filename"]
