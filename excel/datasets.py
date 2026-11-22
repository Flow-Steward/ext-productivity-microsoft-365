"""Dataset grants for Excel: NDJSON rows in (writes) and out (reads)."""

from __future__ import annotations

import json
import os
from collections.abc import Callable, Iterable, Iterator, Mapping
from typing import Any

from flowsteward_extension_sdk import (
    ArtifactAccessError,
    find_dataset_descriptor,
    stream_dataset_bytes,
    write_dataset_stream,
)
from flowsteward_extension_sdk.http import verified_platform_grant_signature

from m365.errors import M365Error

from .ranges import valid_cell

OUTPUT_BINDING_KEY = "excel_range_dataset"
INPUT_BINDING_KEY = "excel_input_dataset"
DATASET_CONTENT_TYPE = "application/x-ndjson"
DEFAULT_MAX_BYTES = 200 * 1024 * 1024
MAX_LINE_BYTES = 2 * 1024 * 1024

_OUTPUT_UNAVAILABLE = "The approved dataset output is unavailable"
_INPUT_UNAVAILABLE = "The approved input dataset is unavailable"

DatasetWriter = Callable[..., dict[str, Any]]


def artifact_max_bytes() -> int:
    raw = str(os.environ.get("FS_EXTENSION_ARTIFACT_MAX_BYTES") or "").strip()
    try:
        value = int(raw) if raw else DEFAULT_MAX_BYTES
    except ValueError:
        return DEFAULT_MAX_BYTES
    return value if value > 0 else DEFAULT_MAX_BYTES


def _scope_matches(payload: Mapping[str, Any], descriptor: Mapping[str, Any]) -> bool:
    host = payload.get("host")
    host_scope = host.get("scope") if isinstance(host, Mapping) else None
    grant_scope = descriptor.get("scope")
    if not isinstance(host_scope, Mapping) or not isinstance(grant_scope, Mapping):
        return False
    return all(
        str(host_scope.get(field) or "").strip()
        and str(host_scope.get(field) or "").strip() == str(grant_scope.get(field) or "").strip()
        for field in ("account_id", "project_id")
    )


# -- output ---------------------------------------------------------------


def output_grant_limit(payload: dict[str, Any]) -> int:
    """Return the byte limit of the approved output dataset grant."""
    try:
        descriptor = find_dataset_descriptor(payload, binding_key=OUTPUT_BINDING_KEY, role="output")
        access = descriptor.get("access")
        max_size = access.get("max_size_bytes") if isinstance(access, Mapping) else None
        if (
            descriptor.get("binding_key") != OUTPUT_BINDING_KEY
            or descriptor.get("content_type") != DATASET_CONTENT_TYPE
            or not isinstance(access, Mapping)
            or access.get("transport") != "presigned_url"
            or access.get("mode") not in {"write", "read_write"}
            or isinstance(max_size, bool)
            or not isinstance(max_size, int)
            or max_size <= 0
            or not verified_platform_grant_signature(descriptor, dict(access))
            or not _scope_matches(payload, descriptor)
        ):
            raise ArtifactAccessError("invalid dataset grant")
        return min(max_size, artifact_max_bytes())
    except (ArtifactAccessError, TypeError, ValueError) as exc:
        raise M365Error("dataset_output_unavailable", _OUTPUT_UNAVAILABLE) from exc


def write_output(
    payload: dict[str, Any],
    chunks: Iterable[bytes],
    *,
    expected_bytes: Callable[[], int],
    writer: DatasetWriter | None = None,
) -> str:
    try:
        written = (writer or write_dataset_stream)(
            payload,
            chunks,
            binding_key=OUTPUT_BINDING_KEY,
            content_type=DATASET_CONTENT_TYPE,
        )
    except M365Error:
        raise
    except Exception as exc:
        raise M365Error("dataset_output_unavailable", _OUTPUT_UNAVAILABLE) from exc
    handle = written.get("dataset_handle") if isinstance(written, Mapping) else None
    size = written.get("size_bytes") if isinstance(written, Mapping) else None
    if (
        not isinstance(handle, str)
        or not handle.strip()
        or isinstance(size, bool)
        or not isinstance(size, int)
        or size != expected_bytes()
    ):
        raise M365Error("dataset_output_unavailable", _OUTPUT_UNAVAILABLE)
    return handle.strip()


# -- input ----------------------------------------------------------------


def _columns(value: Any) -> list[str]:
    if isinstance(value, Mapping):
        if set(value) != {"columns"}:
            raise ArtifactAccessError("invalid dataset schema")
        value = value.get("columns")
    if not isinstance(value, list) or not 1 <= len(value) <= 16_384:
        raise ArtifactAccessError("invalid dataset schema")
    result: list[str] = []
    for row in value:
        name = row.get("name") if isinstance(row, Mapping) else None
        if (
            not isinstance(name, str)
            or not 1 <= len(name) <= 255
            or name in result
            or any(ord(char) < 32 or ord(char) == 127 for char in name)
        ):
            raise ArtifactAccessError("invalid dataset schema")
        result.append(name)
    return result


def read_input_rows(
    payload: dict[str, Any], dataset_handle: str, *, max_rows: int
) -> tuple[list[str], list[dict[str, Any]]]:
    """Return ``(columns, rows)`` of the approved NDJSON input dataset.

    Every row is an object whose keys are a subset of the declared columns.
    """
    try:
        descriptor = find_dataset_descriptor(
            payload, dataset_id=dataset_handle, binding_key=INPUT_BINDING_KEY, role="input"
        )
        access = descriptor.get("access")
        size = descriptor.get("size_bytes")
        row_count = descriptor.get("row_count")
        limit = artifact_max_bytes()
        if (
            descriptor.get("binding_key") != INPUT_BINDING_KEY
            or descriptor.get("content_type") != DATASET_CONTENT_TYPE
            or not isinstance(access, Mapping)
            or access.get("transport") != "presigned_url"
            or access.get("mode") not in {"read", "read_write"}
            or isinstance(size, bool)
            or not isinstance(size, int)
            or not 0 <= size <= limit
            or isinstance(row_count, bool)
            or not isinstance(row_count, int)
            or row_count < 0
            or not _scope_matches(payload, descriptor)
        ):
            raise ArtifactAccessError("invalid dataset grant")
        columns = _columns(descriptor.get("schema"))
        if row_count > max_rows:
            raise M365Error("range_too_large", f"The input dataset has more than {max_rows} rows")
        chunks = stream_dataset_bytes(
            payload,
            dataset_id=dataset_handle,
            binding_key=INPUT_BINDING_KEY,
            chunk_size=1024 * 1024,
            timeout_seconds=30.0,
        )
        rows = list(_rows(chunks, columns=columns, limit=limit))
    except M365Error:
        raise
    except (ArtifactAccessError, OSError, TypeError, ValueError) as exc:
        raise M365Error("dataset_input_unavailable", _INPUT_UNAVAILABLE) from exc
    if len(rows) != row_count:
        raise M365Error("dataset_input_unavailable", "The input dataset row count is invalid")
    return columns, rows


def _rows(chunks: Iterable[bytes], *, columns: list[str], limit: int) -> Iterator[dict[str, Any]]:
    buffer = bytearray()
    total = 0
    allowed = set(columns)
    for raw in chunks:
        if not isinstance(raw, (bytes, bytearray, memoryview)):
            raise M365Error("dataset_input_unavailable", "The input dataset stream is invalid")
        buffer.extend(raw)
        total += len(raw)
        if total > limit or len(buffer) > MAX_LINE_BYTES:
            raise M365Error("range_too_large", "The input dataset is too large")
        while b"\n" in buffer:
            line, _, rest = buffer.partition(b"\n")
            buffer = bytearray(rest)
            if line.strip():
                yield _row(bytes(line), allowed)
    if bytes(buffer).strip():
        yield _row(bytes(buffer), allowed)


def _row(line: bytes, allowed: set[str]) -> dict[str, Any]:
    try:
        value = json.loads(line.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise M365Error(
            "dataset_input_unavailable", "The input dataset contains invalid NDJSON"
        ) from exc
    if not isinstance(value, dict) or set(value) - allowed:
        raise M365Error("dataset_input_unavailable", "The input dataset row shape is invalid")
    if not all(valid_cell(cell) for cell in value.values()):
        raise M365Error("dataset_input_unavailable", "The input dataset contains an invalid cell")
    return value


__all__ = [
    "DATASET_CONTENT_TYPE",
    "INPUT_BINDING_KEY",
    "OUTPUT_BINDING_KEY",
    "artifact_max_bytes",
    "output_grant_limit",
    "read_input_rows",
    "write_output",
]
