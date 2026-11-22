"""Microsoft QuickXorHash, the content hash OneDrive for Business and SharePoint report.

Graph reports ``file.hashes.quickXorHash`` (base64) for business drives; SHA-256 is
not available there. The algorithm XORs every byte into a 160-bit value, the byte at
stream position ``k`` rotated left by ``(11 * k) mod 160`` bits, and finally XORs the
64-bit little-endian stream length into the last eight bytes.

Because the shift only depends on ``k mod 160``, bytes at the same position modulo 160
can be XORed together first. The implementation folds whole aligned blocks as big
integers, so hashing a 100 MB upload costs a few thousand integer operations instead of
a Python loop over every byte.
"""

from __future__ import annotations

import base64

WIDTH_BITS = 160
SHIFT = 11
_ROW_BYTES = WIDTH_BITS  # bytes until the shift pattern repeats
_BLOCK_BYTES = _ROW_BYTES * 512
_ROW_BITS = _ROW_BYTES * 8
_ROW_MASK = (1 << _ROW_BITS) - 1
_WIDTH_MASK = (1 << WIDTH_BITS) - 1


class QuickXorHash:
    def __init__(self) -> None:
        self._acc = 0
        self._pending = b""
        self._length = 0

    def update(self, data: bytes | bytearray | memoryview) -> None:
        chunk = bytes(data)
        if not chunk:
            return
        self._length += len(chunk)
        buffer = self._pending + chunk if self._pending else chunk
        full = len(buffer) - len(buffer) % _BLOCK_BYTES
        acc = self._acc
        view = memoryview(buffer)
        for offset in range(0, full, _BLOCK_BYTES):
            acc ^= int.from_bytes(view[offset : offset + _BLOCK_BYTES], "little")
        self._acc = acc
        self._pending = bytes(view[full:])

    def digest(self) -> bytes:
        acc = self._acc ^ int.from_bytes(self._pending, "little")
        columns = 0
        while acc:
            columns ^= acc & _ROW_MASK
            acc >>= _ROW_BITS
        value = 0
        for index in range(_ROW_BYTES):
            byte = (columns >> (8 * index)) & 0xFF
            if not byte:
                continue
            shift = (index * SHIFT) % WIDTH_BITS
            rotated = (byte << shift) | (byte >> (WIDTH_BITS - shift)) if shift else byte
            value ^= rotated & _WIDTH_MASK
        out = bytearray(value.to_bytes(WIDTH_BITS // 8, "little"))
        for index, length_byte in enumerate(self._length.to_bytes(8, "little")):
            out[WIDTH_BITS // 8 - 8 + index] ^= length_byte
        return bytes(out)

    def b64digest(self) -> str:
        return base64.b64encode(self.digest()).decode("ascii")


def quick_xor_hash(data: bytes) -> str:
    hasher = QuickXorHash()
    hasher.update(data)
    return hasher.b64digest()


__all__ = ["QuickXorHash", "quick_xor_hash"]
