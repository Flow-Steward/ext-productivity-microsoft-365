"""Hidden step markers in Teams message bodies (task D3, best effort).

Graph has no idempotency key for chat messages. A post carries a hidden marker derived
from the step's idempotency key so that, after an ambiguous post, the latest messages
can be searched for it. The marker never carries the key itself, only a short hash.

Two interchangeable strategies:

* :class:`HtmlCommentMarker` (default): ``<!-- fs:<token> -->`` at the end of the body.
* :class:`ZeroWidthMarker`: the token's bits as zero-width characters between word
  joiners, for tenants where the comment is stripped (decided by live tests).

:func:`strip_markers` removes markers of every strategy, so output never shows one,
whichever strategy wrote it.
"""

from __future__ import annotations

import hashlib
import re
from typing import Protocol

TOKEN_HEX_CHARS = 24
_ZERO = "\u200b"  # zero width space
_ONE = "\u200c"  # zero width non-joiner
_FENCE = "\u2060"  # word joiner
_COMMENT = re.compile(rf"<!--\s*fs:([0-9a-f]{{{TOKEN_HEX_CHARS}}})\s*-->")
_ZERO_WIDTH = re.compile(f"{_FENCE}([{_ZERO}{_ONE}]{{{TOKEN_HEX_CHARS * 4}}}){_FENCE}")
#: An HTML-escaped comment, in case Teams stores it as text.
_ESCAPED_COMMENT = re.compile(rf"&lt;!--\s*fs:([0-9a-f]{{{TOKEN_HEX_CHARS}}})\s*--&gt;")


def marker_token(key: str) -> str:
    """A short, stable hash of the step key; the key itself never leaves the extension."""
    digest = hashlib.sha256(f"flowsteward.teams.marker:{key}".encode()).hexdigest()
    return digest[:TOKEN_HEX_CHARS]


class MarkerStrategy(Protocol):
    name: str

    def embed(self, html_body: str, key: str) -> str:
        """Return the body with the marker for ``key`` appended."""

    def found(self, content: str, key: str) -> bool:
        """Whether ``content`` (a stored message body) carries the marker for ``key``."""


class HtmlCommentMarker:
    name = "html_comment"

    def embed(self, html_body: str, key: str) -> str:
        return f"{html_body}<!-- fs:{marker_token(key)} -->"

    def found(self, content: str, key: str) -> bool:
        token = marker_token(key)
        return any(
            match.group(1) == token
            for pattern in (_COMMENT, _ESCAPED_COMMENT)
            for match in pattern.finditer(content or "")
        )


class ZeroWidthMarker:
    name = "zero_width"

    def embed(self, html_body: str, key: str) -> str:
        return f"{html_body}{_encode_zero_width(marker_token(key))}"

    def found(self, content: str, key: str) -> bool:
        token = marker_token(key)
        return any(
            _decode_zero_width(match.group(1)) == token
            for match in _ZERO_WIDTH.finditer(content or "")
        )


def _encode_zero_width(token: str) -> str:
    bits = "".join(f"{int(char, 16):04b}" for char in token)
    return _FENCE + "".join(_ONE if bit == "1" else _ZERO for bit in bits) + _FENCE


def _decode_zero_width(value: str) -> str:
    bits = "".join("1" if char == _ONE else "0" for char in value)
    return "".join(f"{int(bits[i : i + 4], 2):x}" for i in range(0, len(bits), 4))


def strip_markers(content: str) -> str:
    """Remove the markers of every strategy (used on all message output and input)."""
    if not content:
        return content
    for pattern in (_COMMENT, _ESCAPED_COMMENT, _ZERO_WIDTH):
        content = pattern.sub("", content)
    return content


#: Swap here (one line) if live tests show that the tenant strips HTML comments.
DEFAULT_MARKER: MarkerStrategy = HtmlCommentMarker()

__all__ = [
    "DEFAULT_MARKER",
    "HtmlCommentMarker",
    "MarkerStrategy",
    "ZeroWidthMarker",
    "marker_token",
    "strip_markers",
]
