"""A1 addresses, cell values and value comparison for Excel ranges."""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from typing import Any

from m365.errors import M365Error

#: Excel limits: 16,384 columns (XFD) and 1,048,576 rows.
MAX_COLUMN = 16_384
MAX_ROW = 1_048_576
MAX_CELL_TEXT = 32_767

_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$")
_NUMERIC = re.compile(r"^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$")
#: Leading characters that make Excel parse a typed value as a formula.
_FORMULA_LEADERS = ("=", "+", "-", "@")


@dataclass(frozen=True)
class Bounds:
    start_column: int
    start_row: int
    end_column: int
    end_row: int

    @property
    def rows(self) -> int:
        return self.end_row - self.start_row + 1

    @property
    def columns(self) -> int:
        return self.end_column - self.start_column + 1

    @property
    def is_cell(self) -> bool:
        return self.rows == 1 and self.columns == 1

    def address(self) -> str:
        start = f"{column_name(self.start_column)}{self.start_row}"
        if self.is_cell:
            return start
        return f"{start}:{column_name(self.end_column)}{self.end_row}"

    def row_slice(self, first_row: int, last_row: int) -> Bounds:
        return Bounds(self.start_column, first_row, self.end_column, last_row)


def column_number(letters: str) -> int:
    value = 0
    for char in letters.upper():
        value = value * 26 + ord(char) - 64
    return value


def column_name(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _cell(value: str) -> tuple[int, int]:
    match = _CELL.match(value)
    if not match:
        raise M365Error("range_invalid", "The address must be an A1 cell or range such as B2:D20")
    column = column_number(match.group(1))
    row = int(match.group(2))
    if column > MAX_COLUMN or row > MAX_ROW:
        raise M365Error("range_invalid", "The address is outside the Excel grid")
    return column, row


def parse_address(value: Any) -> Bounds:
    """Parse a cell-bounded A1 address without a sheet name (``B2`` or ``B2:D20``)."""
    text = value.strip() if isinstance(value, str) else ""
    if not text or len(text) > 32 or "!" in text:
        raise M365Error(
            "range_invalid",
            "The address must be an A1 cell or range such as B2:D20, without a sheet name",
        )
    parts = text.split(":")
    if len(parts) > 2:
        raise M365Error("range_invalid", "The address must be an A1 cell or range such as B2:D20")
    first = _cell(parts[0])
    last = _cell(parts[1]) if len(parts) == 2 else first
    if first[0] > last[0] or first[1] > last[1]:
        raise M365Error("range_invalid", "The address bounds are reversed")
    return Bounds(first[0], first[1], last[0], last[1])


def parse_graph_address(value: Any) -> tuple[str, Bounds]:
    """Split a Graph range address (``'My Sheet'!A1:D20``) into sheet name and bounds."""
    text = value if isinstance(value, str) else ""
    if "!" not in text or len(text) > 400:
        raise M365Error("provider_unavailable", "Excel returned an invalid range address")
    sheet, _, address = text.rpartition("!")
    if len(sheet) >= 2 and sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    if not sheet:
        raise M365Error("provider_unavailable", "Excel returned an invalid range address")
    try:
        return sheet, parse_address(address)
    except M365Error as exc:
        raise M365Error("provider_unavailable", "Excel returned an invalid range address") from exc


def range_segment(bounds: Bounds) -> str:
    return f"range(address='{bounds.address()}')"


# -- values ---------------------------------------------------------------


def valid_cell(value: Any) -> bool:
    if value is None or isinstance(value, (bool, int)):
        return True
    if isinstance(value, float):
        return math.isfinite(value)
    return isinstance(value, str) and len(value) <= MAX_CELL_TEXT


def provider_values(value: Any, *, rows: int, columns: int) -> list[list[Any]]:
    """Validate a Graph ``values`` matrix of exactly ``rows`` x ``columns``."""
    if not isinstance(value, list) or len(value) != rows:
        raise M365Error("provider_unavailable", "Excel returned values outside the requested range")
    result: list[list[Any]] = []
    for row in value:
        if not isinstance(row, list) or len(row) != columns or not all(map(valid_cell, row)):
            raise M365Error("provider_unavailable", "Excel returned invalid range values")
        result.append([None if cell == "" else cell for cell in row])
    return result


def is_numeric_text(value: str) -> bool:
    return bool(_NUMERIC.match(value.strip()))


def outgoing_cell(value: Any, *, allow_formulas: bool) -> Any:
    """Turn one input cell into the value sent to Excel.

    ``None`` clears the cell (Graph ignores ``null`` cells, so an empty string is
    sent). Without ``allow_formulas`` a text that Excel would parse as a formula, or
    that starts with an apostrophe, is prefixed with an apostrophe so Excel stores it
    as text.
    """
    if value is None:
        return ""
    if (
        not allow_formulas
        and isinstance(value, str)
        and value[:1] in (*_FORMULA_LEADERS, "'")
        and not is_numeric_text(value)
    ):
        return "'" + value
    return value


def _canonical(value: Any) -> tuple[str, Any]:
    if value is None or value == "":
        return ("empty", None)
    if isinstance(value, bool):
        return ("bool", value)
    if isinstance(value, (int, float)):
        return ("number", float(value))
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"true", "false"}:
            return ("bool", lowered == "true")
        if is_numeric_text(value):
            number = float(value.strip())
            if math.isfinite(number):
                return ("number", number)
        return ("text", value)
    return ("other", repr(value))


def cells_equal(expected: Any, actual: Any, formula: Any = None) -> bool:
    """Compare a requested cell with what Excel stored.

    Excel coerces typed input: numeric text becomes a number and ``TRUE``/``FALSE``
    become booleans, so both sides are compared after the same coercion. A formula
    cell matches when its formula text equals the requested text. Anything Excel
    converts in other ways (dates, percentages) compares unequal, which only ever
    turns a confirmation into ``timeout_unknown``.
    """
    left = _canonical(expected)
    # An empty formula says nothing; only a stored formula is compared.
    candidates = (actual,) if formula is None else (actual, formula)
    for candidate in candidates:
        right = _canonical(candidate)
        if left[0] != right[0]:
            continue
        if left[0] == "number":
            if math.isclose(left[1], right[1], rel_tol=1e-12, abs_tol=1e-12):
                return True
        elif left == right:
            return True
    return False


def rows_equal(expected: list[Any], actual: list[Any], formulas: list[Any] | None = None) -> bool:
    width = max(len(expected), len(actual))
    for index in range(width):
        want = expected[index] if index < len(expected) else None
        got = actual[index] if index < len(actual) else None
        formula = formulas[index] if formulas is not None and index < len(formulas) else None
        if not cells_equal(want, got, formula):
            return False
    return True


def header_name(value: Any) -> str:
    """Return the display text of a header cell."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def dataset_headers(values: list[Any]) -> list[str]:
    result: list[str] = []
    used: set[str] = set()
    for index, value in enumerate(values, start=1):
        base = header_name(value)[:255]
        if not base or any(ord(char) < 32 or ord(char) == 127 for char in base):
            base = f"column_{index}"
        candidate = base
        suffix = 2
        while candidate in used:
            candidate = f"{base}_{suffix}"
            suffix += 1
        result.append(candidate)
        used.add(candidate)
    return result


def key_text(value: Any) -> str:
    """Canonical text of a dedupe key, matching Excel's coercion of typed input."""
    kind, canonical = _canonical(value)
    if kind == "empty":
        return ""
    if kind == "number":
        return repr(float(canonical))
    if kind == "bool":
        return "TRUE" if canonical else "FALSE"
    return str(canonical)


__all__ = [
    "MAX_CELL_TEXT",
    "Bounds",
    "cells_equal",
    "column_name",
    "column_number",
    "dataset_headers",
    "header_name",
    "is_numeric_text",
    "key_text",
    "outgoing_cell",
    "parse_address",
    "parse_graph_address",
    "provider_values",
    "range_segment",
    "rows_equal",
    "valid_cell",
]
