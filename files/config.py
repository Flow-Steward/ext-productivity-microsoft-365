"""Operator-configurable limits of the files package."""

from __future__ import annotations

import os
from collections.abc import Mapping

EXTENSION_ARTIFACT_MAX_BYTES_ENV = "FS_EXTENSION_ARTIFACT_MAX_BYTES"
DEFAULT_EXTENSION_ARTIFACT_MAX_BYTES = 200 * 1024 * 1024
MIN_EXTENSION_ARTIFACT_MAX_BYTES = 1024 * 1024
MAX_EXTENSION_ARTIFACT_MAX_BYTES = 10 * 1024 * 1024 * 1024

#: Files up to this size go up in one PUT; larger ones through an upload session.
SIMPLE_UPLOAD_MAX_BYTES = 4 * 1024 * 1024
#: Upload-session fragments must be a multiple of 320 KiB (Graph rule).
UPLOAD_FRAGMENT_UNIT = 320 * 1024
UPLOAD_CHUNK_BYTES = 20 * UPLOAD_FRAGMENT_UNIT
#: Downloads are read in ranges, so a large file never sits in memory at once.
DOWNLOAD_RANGE_BYTES = 8 * 1024 * 1024
ARTIFACT_READ_CHUNK_BYTES = 1024 * 1024


def artifact_max_bytes(environ: Mapping[str, str] | None = None) -> int:
    values = os.environ if environ is None else environ
    try:
        parsed = int(str(values.get(EXTENSION_ARTIFACT_MAX_BYTES_ENV)).strip())
    except (TypeError, ValueError):
        parsed = DEFAULT_EXTENSION_ARTIFACT_MAX_BYTES
    return max(MIN_EXTENSION_ARTIFACT_MAX_BYTES, min(parsed, MAX_EXTENSION_ARTIFACT_MAX_BYTES))


__all__ = [
    "ARTIFACT_READ_CHUNK_BYTES",
    "DEFAULT_EXTENSION_ARTIFACT_MAX_BYTES",
    "DOWNLOAD_RANGE_BYTES",
    "EXTENSION_ARTIFACT_MAX_BYTES_ENV",
    "SIMPLE_UPLOAD_MAX_BYTES",
    "UPLOAD_CHUNK_BYTES",
    "UPLOAD_FRAGMENT_UNIT",
    "artifact_max_bytes",
]
