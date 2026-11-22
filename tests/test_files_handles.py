"""upload_file takes a handle or Core's file object from a workflow file input."""

from __future__ import annotations

import pytest


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("art_123", "art_123"),
        ("artifact:art_123", "artifact:art_123"),
        ({"artifact_id": "art_123", "filename": "a.xlsx", "size_bytes": 10}, "art_123"),
        ({"artifact_handle": "art_456"}, "art_456"),
    ],
)
def test_upload_accepts_a_handle_or_cores_file_object(value: object, expected: str) -> None:
    from files.paths import artifact_handle_input

    assert artifact_handle_input(value) == expected


@pytest.mark.parametrize("value", [None, "", {}, {"filename": "a.xlsx"}, 42])
def test_upload_refuses_a_missing_handle(value: object) -> None:
    from files.paths import artifact_handle_input
    from m365.errors import M365Error

    with pytest.raises(M365Error):
        artifact_handle_input(value)
