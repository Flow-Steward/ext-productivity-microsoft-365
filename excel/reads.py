"""Excel reads: workbook objects, range values and range → NDJSON dataset."""

from __future__ import annotations

import json
from collections.abc import Iterator, Mapping
from typing import Any

from m365.connection import text
from m365.errors import M365Error
from m365.runtime import OperationContext
from m365.transport import GraphTransport

from .datasets import output_grant_limit, write_output
from .errors import remap_not_found
from .ranges import (
    Bounds,
    dataset_headers,
    parse_address,
    parse_graph_address,
    provider_values,
    range_segment,
)
from .workbook import Workbook, resolve_workbook, segment

MAX_ROWS = 100_000
MAX_CELLS = 1_000_000
CELLS_PER_REQUEST = 20_000
MAX_ROWS_PER_REQUEST = 5_000
RANGE_RESPONSE_LIMIT = 16 * 1024 * 1024
PREVIEW_ROWS = 100
PREVIEW_BYTES = 256 * 1024
MAX_WORKSHEETS = 200
MAX_NAME = 255

READ_RANGE_FIELDS = frozenset({"workbook", "worksheet", "address", "table", "header_mode"})
LIST_OBJECTS_FIELDS = frozenset({"workbook"})


def name_input(value: Any, field: str, *, limit: int = MAX_NAME) -> str:
    result = text(value, limit=limit)
    if not result:
        raise M365Error("invalid_payload", f"{field} is invalid")
    return result


def worksheet_path(workbook: Workbook, worksheet: str) -> str:
    return f"{workbook.base}/worksheets/{segment(worksheet)}"


def table_path(workbook: Workbook, table: str) -> str:
    return f"{workbook.base}/tables/{segment(table)}"


def worksheet_not_found(exc: M365Error) -> M365Error:
    return remap_not_found(exc, "worksheet_not_found", "The worksheet was not found")


def table_not_found(exc: M365Error) -> M365Error:
    return remap_not_found(exc, "table_not_found", "The table was not found")


def rows_per_request(columns: int) -> int:
    return max(1, min(MAX_ROWS_PER_REQUEST, CELLS_PER_REQUEST // max(1, columns)))


def read_matrix(
    transport: GraphTransport,
    workbook: Workbook,
    worksheet: str,
    bounds: Bounds,
    *,
    headers: Mapping[str, str] | None = None,
    with_formulas: bool = False,
) -> Iterator[tuple[list[list[Any]], list[list[Any]] | None]]:
    """Yield ``(values, formulas)`` row slices of ``bounds``, bounded per request."""
    step = rows_per_request(bounds.columns)
    select = "values,formulas" if with_formulas else "values"
    for first in range(bounds.start_row, bounds.end_row + 1, step):
        part = bounds.row_slice(first, min(bounds.end_row, first + step - 1))
        try:
            body = transport.get_json(
                f"{worksheet_path(workbook, worksheet)}/{range_segment(part)}",
                query={"$select": select},
                headers=headers,
                response_limit=RANGE_RESPONSE_LIMIT,
                purpose="Excel range read",
            )
        except M365Error as exc:
            raise worksheet_not_found(exc) from exc
        values = provider_values(body.get("values"), rows=part.rows, columns=part.columns)
        formulas = (
            provider_values(body.get("formulas"), rows=part.rows, columns=part.columns)
            if with_formulas
            else None
        )
        yield values, formulas


def range_address(
    transport: GraphTransport,
    path: str,
    *,
    headers: Mapping[str, str] | None = None,
    purpose: str,
) -> tuple[str, Bounds]:
    body = transport.get_json(path, query={"$select": "address"}, headers=headers, purpose=purpose)
    return parse_graph_address(body.get("address"))


# -- list_workbook_objects --------------------------------------------------


def _collection(body: Mapping[str, Any]) -> list[Mapping[str, Any]]:
    rows = body.get("value")
    if not isinstance(rows, list):
        raise M365Error("provider_unavailable", "Excel returned an invalid collection")
    return [row for row in rows if isinstance(row, Mapping)]


def list_workbook_objects(context: OperationContext) -> dict[str, Any]:
    transport = context.transport
    workbook = resolve_workbook(transport, context.input.get("workbook"))
    sheets_body = transport.get_json(
        f"{workbook.base}/worksheets",
        query={"$select": "id,name,position,visibility"},
        purpose="Excel worksheet list",
    )
    worksheets: list[dict[str, Any]] = []
    for row in _collection(sheets_body):
        name = text(row.get("name"), limit=MAX_NAME)
        if not name:
            continue
        position = row.get("position")
        worksheets.append(
            {
                "id": text(row.get("id"), limit=256),
                "name": name,
                "position": position if isinstance(position, int) else None,
                "visibility": text(row.get("visibility"), limit=32),
            }
        )
    truncated = len(worksheets) > MAX_WORKSHEETS
    worksheets = worksheets[:MAX_WORKSHEETS]
    tables: list[dict[str, Any]] = []
    for sheet in worksheets:
        try:
            body = transport.get_json(
                f"{worksheet_path(workbook, sheet['name'])}/tables",
                query={"$select": "id,name,showHeaders,showTotals"},
                purpose="Excel table list",
            )
        except M365Error as exc:
            raise worksheet_not_found(exc) from exc
        for row in _collection(body):
            name = text(row.get("name"), limit=MAX_NAME)
            if not name:
                continue
            tables.append(
                {
                    "id": text(row.get("id"), limit=256),
                    "name": name,
                    "worksheet": sheet["name"],
                    "show_headers": row.get("showHeaders") is not False,
                    "show_totals": row.get("showTotals") is True,
                }
            )
    return {
        "workbook": workbook.summary(),
        "worksheets": worksheets,
        "tables": tables,
        "truncated": truncated,
    }


# -- read_range -------------------------------------------------------------


def validated_read_input(raw: Mapping[str, Any]) -> dict[str, Any]:
    has_sheet = raw.get("worksheet") is not None
    has_table = raw.get("table") is not None
    if has_sheet == has_table:
        raise M365Error("invalid_payload", "Exactly one of worksheet or table is required")
    header_mode = raw.get("header_mode", "first_row")
    if header_mode not in {"first_row", "none"}:
        raise M365Error("invalid_payload", "header_mode must be first_row or none")
    result: dict[str, Any] = {"workbook": raw.get("workbook"), "header_mode": header_mode}
    if has_table:
        if raw.get("address") is not None:
            raise M365Error("invalid_payload", "address is only allowed with worksheet")
        result["table"] = name_input(raw.get("table"), "table")
    else:
        result["worksheet"] = name_input(raw.get("worksheet"), "worksheet", limit=31)
        if raw.get("address") is not None:
            result["address"] = parse_address(raw.get("address"))
    return result


def _cell_type(value: Any) -> str:
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (int, float)):
        return "number"
    return "string"


def read_range(context: OperationContext) -> dict[str, Any]:
    normalized = validated_read_input(context.input)
    byte_limit = output_grant_limit(context.payload)
    transport = context.transport
    workbook = resolve_workbook(transport, normalized["workbook"])
    if "table" in normalized:
        try:
            worksheet, bounds = range_address(
                transport,
                f"{table_path(workbook, normalized['table'])}/range",
                purpose="Excel table range",
            )
        except M365Error as exc:
            raise table_not_found(exc) from exc
    else:
        worksheet = normalized["worksheet"]
        address: Bounds | None = normalized.get("address")
        if address is not None:
            bounds = address
        else:
            try:
                _, bounds = range_address(
                    transport,
                    f"{worksheet_path(workbook, worksheet)}/usedRange(valuesOnly=true)",
                    purpose="Excel used range",
                )
            except M365Error as exc:
                raise worksheet_not_found(exc) from exc
    header_rows = 1 if normalized["header_mode"] == "first_row" else 0
    if bounds.rows > MAX_ROWS + header_rows or bounds.rows * bounds.columns > MAX_CELLS + (
        header_rows * bounds.columns
    ):
        raise M365Error("range_too_large", "The requested Excel range is too large")
    state: dict[str, Any] = {
        "headers": [] if header_rows else dataset_headers([None] * bounds.columns),
        "header_pending": bool(header_rows),
        "types": {},
        "row_count": 0,
        "bytes": 0,
        "preview": [],
        "preview_bytes": 0,
    }

    def chunks() -> Iterator[bytes]:
        for values, _ in read_matrix(transport, workbook, worksheet, bounds):
            for row in values:
                if state["header_pending"]:
                    state["headers"] = dataset_headers(row)
                    state["header_pending"] = False
                    continue
                record = dict(zip(state["headers"], row, strict=True))
                for name, cell in record.items():
                    if cell is None:
                        continue
                    kind = _cell_type(cell)
                    previous = state["types"].get(name)
                    state["types"][name] = kind if previous in (None, kind) else "string"
                encoded = (
                    json.dumps(record, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
                    + b"\n"
                )
                state["row_count"] += 1
                state["bytes"] += len(encoded)
                if state["bytes"] > byte_limit:
                    raise M365Error("range_too_large", "The Excel range exceeds the dataset limit")
                if (
                    len(state["preview"]) < PREVIEW_ROWS
                    and state["preview_bytes"] + len(encoded) <= PREVIEW_BYTES
                ):
                    state["preview"].append(record)
                    state["preview_bytes"] += len(encoded)
                yield encoded

    handle = write_output(context.payload, chunks(), expected_bytes=lambda: state["bytes"])
    if state["header_pending"]:
        state["headers"] = dataset_headers([None] * bounds.columns)
    schema = [
        {"name": name, "type": state["types"].get(name, "string"), "nullable": True}
        for name in state["headers"]
    ]
    return {
        "dataset_handle": handle,
        "row_count": state["row_count"],
        "column_count": len(state["headers"]),
        "schema": schema,
        "preview": state["preview"],
        "workbook": workbook.summary(),
        "worksheet": worksheet,
        "address": bounds.address(),
    }


__all__ = [
    "LIST_OBJECTS_FIELDS",
    "READ_RANGE_FIELDS",
    "list_workbook_objects",
    "name_input",
    "range_address",
    "read_matrix",
    "read_range",
    "table_not_found",
    "table_path",
    "validated_read_input",
    "worksheet_not_found",
    "worksheet_path",
]
