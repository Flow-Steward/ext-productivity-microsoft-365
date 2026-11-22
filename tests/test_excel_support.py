"""Shared fakes for the Excel tests: an in-memory workbook behind a fake transport.

This module holds no tests. Test modules import it by path so it works with any
pytest import mode.
"""

from __future__ import annotations

import json
import re
from collections.abc import Callable
from typing import Any
from urllib.parse import unquote

from m365.errors import M365Error

DRIVE_ID = "b!acmeDrive_01"
ITEM_ID = "01ACMEBOOK"
SHARING_LINK = "https://acme.sharepoint.com/:x:/s/finance/EacmeLink?e=abc"
SESSION_ID = "session-acme-1"
BASE = f"/drives/{DRIVE_ID}/items/{ITEM_ID}/workbook"

_CELL = re.compile(r"^([A-Z]+)([0-9]+)$")
_NUMERIC = re.compile(r"^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$")


def col_num(letters: str) -> int:
    value = 0
    for char in letters:
        value = value * 26 + ord(char) - 64
    return value


def col_name(number: int) -> str:
    result = ""
    while number:
        number, rest = divmod(number - 1, 26)
        result = chr(65 + rest) + result
    return result


def parse(address: str) -> tuple[int, int, int, int]:
    parts = address.split(":")
    first = _CELL.match(parts[0])
    last = _CELL.match(parts[-1])
    assert first and last, address
    return (
        col_num(first.group(1)),
        int(first.group(2)),
        col_num(last.group(1)),
        int(last.group(2)),
    )


def addr(c1: int, r1: int, c2: int, r2: int) -> str:
    return f"{col_name(c1)}{r1}:{col_name(c2)}{r2}"


def excel_coerce(value: Any) -> tuple[Any, Any]:
    """Mimic Excel's handling of a typed value: returns (value, formula)."""
    if isinstance(value, str):
        if value.startswith("'"):
            return value[1:], value[1:]
        if value.startswith("="):
            return 0, value
        if _NUMERIC.match(value.strip()):
            number = float(value)
            return number, number
        if value.strip().lower() in {"true", "false"}:
            flag = value.strip().lower() == "true"
            return flag, flag
        return value, value
    return value, value


class Sheet:
    def __init__(self, name: str) -> None:
        self.name = name
        self.cells: dict[tuple[int, int], Any] = {}
        self.formulas: dict[tuple[int, int], Any] = {}

    def set(self, column: int, row: int, value: Any) -> None:
        stored, formula = excel_coerce(value)
        if stored == "" or stored is None:
            self.cells.pop((column, row), None)
            self.formulas.pop((column, row), None)
        else:
            self.cells[(column, row)] = stored
            self.formulas[(column, row)] = formula

    def put_raw(self, column: int, row: int, value: Any) -> None:
        self.cells[(column, row)] = value
        self.formulas[(column, row)] = value

    def matrix(self, c1: int, r1: int, c2: int, r2: int, *, formulas: bool = False):
        source = self.formulas if formulas else self.cells
        return [[source.get((c, r), "") for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]


class Table:
    def __init__(
        self, name: str, sheet: Sheet, columns: list[str], *, column: int = 1, row: int = 1
    ):
        self.name = name
        self.sheet = sheet
        self.columns = columns
        self.start_column = column
        self.header_row = row
        self.count = 0
        for index, header in enumerate(columns):
            sheet.put_raw(column + index, row, header)

    @property
    def end_column(self) -> int:
        return self.start_column + len(self.columns) - 1

    def body_rows(self) -> tuple[int, int]:
        first = self.header_row + 1
        return first, first + max(self.count, 1) - 1

    def add_raw(self, values: list[Any]) -> None:
        row = self.header_row + 1 + self.count
        for index, value in enumerate(values):
            self.sheet.put_raw(self.start_column + index, row, value)
        self.count += 1

    def add(self, values: list[Any]) -> None:
        row = self.header_row + 1 + self.count
        for index, value in enumerate(values):
            self.sheet.set(self.start_column + index, row, value)
        self.count += 1


class Fault:
    """Raise ``error`` for matching calls; ``apply`` lands the write first."""

    def __init__(
        self,
        method: str,
        marker: str,
        error: M365Error,
        *,
        apply: bool,
        times: int = 1,
        land_after: int | None = None,
        on_raise: Callable[[], None] | None = None,
    ):
        self.method = method
        self.marker = marker
        self.error = error
        self.apply = apply
        self.times = times
        #: Land the write only after this many later calls (a slow write).
        self.land_after = land_after
        #: Called when the fault fires, for example to land part of a write later.
        self.on_raise = on_raise


def ambiguous(status: int | None = 504) -> M365Error:
    return M365Error(
        "timeout_unknown",
        "Microsoft 365 did not confirm the write",
        http_status=status,
        ambiguous=True,
    )


class FakeExcel:
    """A fake GraphTransport that serves one workbook."""

    def __init__(self) -> None:
        self.calls: list[dict[str, Any]] = []
        self.sheets: dict[str, Sheet] = {"Data": Sheet("Data")}
        self.tables: dict[str, Table] = {}
        self.item: dict[str, Any] = {
            "id": ITEM_ID,
            "name": "Acme Ledger.xlsx",
            "file": {"mimeType": "application/vnd.ms-excel"},
            "webUrl": "https://acme.sharepoint.com/sites/finance/Ledger.xlsx",
            "parentReference": {"driveId": DRIVE_ID, "driveType": "documentLibrary"},
        }
        self.share_error: M365Error | None = None
        self.faults: list[Fault] = []
        #: Writes that land only after this many later calls.
        self.delayed: list[tuple[int, Callable[[], None]]] = []
        self.remaining = 200.0
        self.search_body: dict[str, Any] = {"value": []}
        self.drives: dict[str, dict[str, Any]] = {}
        self.sites: dict[str, dict[str, Any]] = {}

    # -- GraphTransport surface ------------------------------------------

    def remaining_seconds(self) -> float:
        return self.remaining

    def get_json(self, path: str, *, query=None, headers=None, response_limit=0, purpose: str = ""):
        return self.request(
            "GET", path, query=query, headers=headers, retry="read", purpose=purpose
        ).json()

    def request(
        self,
        method: str,
        path: str,
        *,
        query=None,
        json_body=None,
        headers=None,
        response_limit=0,
        retry: str,
        purpose: str = "",
        **_kwargs,
    ):
        from m365.transport import GraphResponse

        self.calls.append(
            {
                "method": method,
                "path": path,
                "query": dict(query or {}),
                "body": json_body,
                "headers": dict(headers or {}),
                "retry": retry,
            }
        )
        self._tick_delayed()
        fault = next(
            (f for f in self.faults if f.method == method and f.marker in path and f.times > 0),
            None,
        )
        if fault is not None:
            fault.times -= 1
            if fault.apply:
                self._route(method, path, json_body)
            elif fault.land_after is not None:
                self.delayed.append(
                    (fault.land_after, lambda: self._route(method, path, json_body))
                )
            if fault.on_raise is not None:
                fault.on_raise()
            raise fault.error
        body = self._route(method, path, json_body)
        return GraphResponse(status=200, headers={}, body=json.dumps(body).encode("utf-8"))

    # -- helpers -----------------------------------------------------------

    def calls_matching(self, method: str, marker: str) -> list[dict[str, Any]]:
        return [c for c in self.calls if c["method"] == method and marker in c["path"]]

    def _tick_delayed(self) -> None:
        remaining = []
        for count, action in self.delayed:
            if count <= 0:
                action()
            else:
                remaining.append((count - 1, action))
        self.delayed = remaining

    def _route(self, method: str, path: str, body: Any) -> dict[str, Any]:
        if path == f"/drives/{DRIVE_ID}/items/{ITEM_ID}" and method == "GET":
            return dict(self.item)
        if path.startswith("/drives/") and "/items/" in path and "/workbook" not in path:
            raise M365Error("item_not_found", "missing", http_status=404)
        if path.startswith("/shares/u!") and path.endswith("/driveItem"):
            if self.share_error is not None:
                raise self.share_error
            return dict(self.item)
        if path == "/search/query" and method == "POST":
            return self.search_body
        if path.startswith("/drives/") and path.count("/") == 2:
            drive_id = unquote(path.split("/")[2])
            if drive_id in self.drives:
                return self.drives[drive_id]
            raise M365Error("item_not_found", "missing", http_status=404)
        if path.startswith("/sites/"):
            site_id = unquote(path.split("/", 2)[2])
            if site_id in self.sites:
                return self.sites[site_id]
            raise M365Error("item_not_found", "missing", http_status=404)
        assert path.startswith(BASE), path
        rest = path[len(BASE) :]
        if rest == "/createSession":
            assert body == {"persistChanges": True}
            return {"id": SESSION_ID, "persistChanges": True}
        if rest == "/closeSession":
            return {}
        if rest == "/worksheets":
            return {
                "value": [
                    {"id": f"{{{index}}}", "name": name, "position": index, "visibility": "Visible"}
                    for index, name in enumerate(self.sheets)
                ]
            }
        match = re.match(r"^/worksheets/([^/]+)(/.*)$", rest)
        if match:
            sheet = self.sheets.get(unquote(match.group(1)))
            if sheet is None:
                raise M365Error("item_not_found", "missing", http_status=404)
            return self._sheet_route(method, sheet, match.group(2), body)
        match = re.match(r"^/tables/([^/]+)(/.*)$", rest)
        if match:
            table = self.tables.get(unquote(match.group(1)))
            if table is None:
                raise M365Error("item_not_found", "missing", http_status=404)
            return self._table_route(method, table, match.group(2), body)
        raise AssertionError(f"unexpected call {method} {path}")

    def _sheet_route(self, method: str, sheet: Sheet, rest: str, body: Any) -> dict[str, Any]:
        if rest == "/tables":
            return {
                "value": [
                    {
                        "id": f"{{t-{t.name}}}",
                        "name": t.name,
                        "showHeaders": True,
                        "showTotals": False,
                    }
                    for t in self.tables.values()
                    if t.sheet is sheet
                ]
            }
        if rest.startswith("/usedRange"):
            if not sheet.cells:
                return {"address": f"{sheet.name}!A1"}
            columns = [c for c, _ in sheet.cells]
            rows = [r for _, r in sheet.cells]
            return {
                "address": f"'{sheet.name}'!{addr(min(columns), min(rows), max(columns), max(rows))}"
            }
        match = re.match(r"^/range\(address='([A-Z0-9:]+)'\)$", rest)
        assert match, rest
        c1, r1, c2, r2 = parse(match.group(1))
        if method == "GET":
            return {
                "values": sheet.matrix(c1, r1, c2, r2),
                "formulas": sheet.matrix(c1, r1, c2, r2, formulas=True),
            }
        assert method == "PATCH"
        values = body["values"]
        assert len(values) == r2 - r1 + 1 and all(len(row) == c2 - c1 + 1 for row in values)
        for dr, row in enumerate(values):
            for dc, value in enumerate(row):
                sheet.set(c1 + dc, r1 + dr, value)
        return {"address": f"{sheet.name}!{match.group(1)}"}

    def _table_route(self, method: str, table: Table, rest: str, body: Any) -> dict[str, Any]:
        if rest == "/headerRowRange":
            return {"values": [list(table.columns)]}
        if rest == "/range":
            last = table.header_row + max(table.count, 1)
            return {
                "address": f"{table.sheet.name}!"
                + addr(table.start_column, table.header_row, table.end_column, last)
            }
        if rest == "/dataBodyRange":
            first, last = table.body_rows()
            return {
                "address": f"{table.sheet.name}!"
                + addr(table.start_column, first, table.end_column, last)
            }
        if rest == "/rows/add" and method == "POST":
            assert body["index"] is None
            start = table.count
            for row in body["values"]:
                assert len(row) == len(table.columns)
                table.add(row)
            return {"index": start, "values": body["values"][:1]}
        raise AssertionError(f"unexpected table call {method} {rest}")

    def add_table(self, name: str, columns: list[str], *, sheet: str = "Data") -> Table:
        table = Table(name, self.sheets[sheet], columns)
        self.tables[name] = table
        return table


def payload(
    operation_id: str, action_input: dict[str, Any], *, test_mode: bool = False
) -> dict[str, Any]:
    return {
        "mode": "action",
        "action": {
            "action_id": operation_id,
            "input": {"connection_ref": "conn-acme-excel", **action_input},
            "target": {
                "connection": {
                    "connection_id": "conn-acme-excel",
                    "connection_type_id": "m365_excel_user",
                    "config": {},
                    "secrets": {"access_token": "acme-delegated-token"},
                }
            },
        },
        "runtime_context": {"test_mode": test_mode},
    }


def ids_workbook() -> dict[str, str]:
    return {"drive_id": DRIVE_ID, "drive_item_id": ITEM_ID}


def run(fake: FakeExcel, operation_id: str, action_input: dict[str, Any], **kwargs: Any):
    from excel.operations import OPERATIONS
    from m365.runtime import run_operation

    return run_operation(
        payload(operation_id, action_input, **kwargs),
        OPERATIONS,
        transport_factory=lambda _context: fake,
    )
