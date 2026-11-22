"""QuickXorHash against a direct transcription of Microsoft's reference algorithm."""

from __future__ import annotations

import base64
import random

import pytest

from files.quickxor import QuickXorHash, quick_xor_hash

_M64 = (1 << 64) - 1


def _reference(chunks: list[bytes]) -> str:
    """Line-by-line port of the published C# QuickXorHash (three 64-bit cells)."""
    data = [0, 0, 0]
    length = 0
    shift_so_far = 0
    for array in chunks:
        size = len(array)
        index = shift_so_far // 64
        offset = shift_so_far % 64
        for i in range(min(size, 160)):
            last = index == 2
            bits = 32 if last else 64
            if offset <= bits - 8:
                for j in range(i, size, 160):
                    data[index] = (data[index] ^ (array[j] << offset)) & _M64
            else:
                second = 0 if last else index + 1
                low = bits - offset
                xored = 0
                for j in range(i, size, 160):
                    xored ^= array[j]
                data[index] = (data[index] ^ (xored << offset)) & _M64
                data[second] = (data[second] ^ (xored >> low)) & _M64
            offset += 11
            while offset >= bits:
                index = 0 if last else index + 1
                offset -= bits
        shift_so_far = (shift_so_far + 11 * (size % 160)) % 160
        length += size
    out = bytearray(20)
    out[0:8] = data[0].to_bytes(8, "little")
    out[8:16] = data[1].to_bytes(8, "little")
    out[16:20] = data[2].to_bytes(8, "little")[:4]
    for i, byte in enumerate(length.to_bytes(8, "little")):
        out[12 + i] ^= byte
    return base64.b64encode(bytes(out)).decode("ascii")


def test_empty_content_hash_is_twenty_zero_bytes() -> None:
    assert quick_xor_hash(b"") == "AAAAAAAAAAAAAAAAAAAAAAAAAAA="


@pytest.mark.parametrize("size", [1, 2, 19, 159, 160, 161, 1000, 81919, 81920, 81921, 250_000])
def test_fast_hash_matches_reference_for_any_chunking(size: int) -> None:
    rng = random.Random(size)
    data = rng.randbytes(size)
    for cut in sorted({0, 1, 7, size // 3, size}):
        parts = [data[:cut], data[cut:]]
        hasher = QuickXorHash()
        for part in parts:
            hasher.update(part)
        assert hasher.b64digest() == _reference(parts)
    assert quick_xor_hash(data) == _reference([data])


def test_hash_depends_on_length_and_position() -> None:
    assert quick_xor_hash(b"\x00") != quick_xor_hash(b"")
    assert quick_xor_hash(b"ab") != quick_xor_hash(b"ba")
