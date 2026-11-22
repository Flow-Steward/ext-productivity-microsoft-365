"""Excel writes: ``write_range`` and ``append_table_rows``.

Both run inside a persistent workbook session and send their single write with
``retry="none"`` (task B6). An ambiguous outcome (timeout, 5xx, unreadable
response) is never repeated; it is confirmed by a bounded re-read instead:

* ``write_range``: the range holds the requested values → success, otherwise
  ``timeout_unknown``.
* ``append_table_rows``: with ``dedupe_column``, every appended key is present with
  equal values → success, otherwise ``timeout_unknown`` with a per-row status;
  without ``dedupe_column`` an ambiguous append is always ``timeout_unknown``.
"""

from __future__ import annotations

import json
import time
from collections.abc import Callable, Mapping
from typing import Any

from m365.errors import M365Error
from m365.runtime import OperationContext
from m365.transport import GraphTransport

from .datasets import read_input_rows
from .errors import ExcelRowsError
from .ranges import (
    MAX_COLUMN,
    MAX_ROW,
    Bounds,
    cells_equal,
    header_name,
    key_text,
    outgoing_cell,
    parse_address,
    range_segment,
    rows_equal,
    valid_cell,
)
from .reads import (
    MAX_CELLS,
    MAX_ROWS,
    name_input,
    range_address,
    read_matrix,
    table_not_found,
    table_path,
    worksheet_not_found,
    worksheet_path,
)
from .workbook import Workbook, resolve_workbook, workbook_session

MAX_INLINE_ROWS = 500
MAX_INLINE_COLUMNS = 100
MAX_INLINE_CELLS = 20_000
MAX_WRITE_CELLS = 100_000
MAX_WRITE_BODY_BYTES = 4 * 1024 * 1024
MAX_APPEND_ROWS = 500
WRITE_RESPONSE_LIMIT = 64 * 1024 * 1024
#: Waits before each confirmation re-read, bounded by the operation deadline.
REREAD_WAITS = (0.0, 2.0, 5.0, 10.0)
REREAD_RESERVE_SECONDS = 15.0
MAX_REPORTED_KEYS = 10

WRITE_RANGE_FIELDS = frozenset(
    {
        "workbook",
        "worksheet",
        "address",
        "dataset_handle",
        "inline_rows",
        "include_header",
        "allow_formulas",
    }
)
APPEND_FIELDS = frozenset(
    {"workbook", "table", "dataset_handle", "inline_rows", "dedupe_column", "allow_formulas"}
)

#: Replaced in tests so the bounded re-read does not sleep.
sleep: Callable[[float], None] = time.sleep


def _flag(raw: Mapping[str, Any], field: str) -> bool:
    value = raw.get(field, False)
    if value is None:
        return False
    if not isinstance(value, bool):
        raise M365Error("invalid_payload", f"{field} must be boolean")
    return value


def _data_source(raw: Mapping[str, Any]) -> str:
    has_dataset = raw.get("dataset_handle") is not None
    has_inline = raw.get("inline_rows") is not None
    if has_dataset == has_inline:
        raise M365Error(
            "invalid_payload", "Exactly one of dataset_handle or inline_rows is required"
        )
    return "dataset" if has_dataset else "inline"


def _cell(value: Any) -> Any:
    if not valid_cell(value):
        raise M365Error("invalid_payload", "inline_rows contains an invalid cell")
    return value


def _inline_matrix(value: Any) -> list[list[Any]]:
    if not isinstance(value, list) or not 1 <= len(value) <= MAX_INLINE_ROWS:
        raise M365Error("invalid_payload", f"inline_rows must hold 1 to {MAX_INLINE_ROWS} rows")
    result: list[list[Any]] = []
    cells = 0
    for row in value:
        if not isinstance(row, list) or not 1 <= len(row) <= MAX_INLINE_COLUMNS:
            raise M365Error("invalid_payload", "inline_rows must be a list of value lists")
        cells += len(row)
        if cells > MAX_INLINE_CELLS:
            raise M365Error("invalid_payload", "inline_rows has too many cells")
        result.append([_cell(cell) for cell in row])
    return result


def _inline_records(value: Any) -> list[dict[str, Any]]:
    if not isinstance(value, list) or not 1 <= len(value) <= MAX_APPEND_ROWS:
        raise M365Error("invalid_payload", f"inline_rows must hold 1 to {MAX_APPEND_ROWS} rows")
    result: list[dict[str, Any]] = []
    for row in value:
        if (
            not isinstance(row, Mapping)
            or not 1 <= len(row) <= MAX_INLINE_COLUMNS
            or not all(isinstance(key, str) and 1 <= len(key) <= 255 for key in row)
        ):
            raise M365Error(
                "invalid_payload", "inline_rows must be objects keyed by table column names"
            )
        result.append({key: _cell(cell) for key, cell in row.items()})
    return result


def _deadline_waits(transport: GraphTransport) -> list[float]:
    waits: list[float] = []
    budget = transport.remaining_seconds() - REREAD_RESERVE_SECONDS
    for wait in REREAD_WAITS:
        if budget < wait:
            break
        waits.append(wait)
        budget -= wait
    return waits or [0.0]


def _is_unconfirmed(exc: M365Error) -> bool:
    # A response that was too large to read came with a 2xx status: the write
    # landed, but its body is unknown, so it is confirmed like an ambiguous one.
    return exc.ambiguous or exc.code == "response_too_large"


# -- write_range ------------------------------------------------------------


def validated_write_input(raw: Mapping[str, Any]) -> dict[str, Any]:
    source = _data_source(raw)
    include_header = _flag(raw, "include_header")
    if include_header and source != "dataset":
        raise M365Error("invalid_payload", "include_header is only allowed with dataset_handle")
    result: dict[str, Any] = {
        "workbook": raw.get("workbook"),
        "worksheet": name_input(raw.get("worksheet"), "worksheet", limit=31),
        "address": parse_address(raw.get("address")),
        "include_header": include_header,
        "allow_formulas": _flag(raw, "allow_formulas"),
    }
    if source == "dataset":
        result["dataset_handle"] = name_input(raw.get("dataset_handle"), "dataset_handle")
    else:
        result["inline_rows"] = _inline_matrix(raw.get("inline_rows"))
    return result


def _write_matrix(context: OperationContext, normalized: Mapping[str, Any]) -> list[list[Any]]:
    if "inline_rows" in normalized:
        return [list(row) for row in normalized["inline_rows"]]
    columns, records = read_input_rows(
        context.payload, normalized["dataset_handle"], max_rows=MAX_WRITE_CELLS
    )
    matrix = [[record.get(column) for column in columns] for record in records]
    if normalized["include_header"]:
        matrix.insert(0, list(columns))
    return matrix


def write_target(address: Bounds, matrix: list[list[Any]]) -> tuple[Bounds, list[list[Any]]]:
    """Return the replaced rectangle and the requested values padded to fill it.

    A single-cell address anchors the values; a rectangle is replaced as a whole and
    its cells beyond the values are cleared.
    """
    height = len(matrix)
    width = max((len(row) for row in matrix), default=0)
    if address.is_cell:
        if not height or not width:
            raise M365Error("invalid_payload", "There are no values to write")
        target = Bounds(
            address.start_column,
            address.start_row,
            address.start_column + width - 1,
            address.start_row + height - 1,
        )
        if target.end_column > MAX_COLUMN or target.end_row > MAX_ROW:
            raise M365Error("range_invalid", "The values do not fit into the worksheet")
    else:
        if height > address.rows or width > address.columns:
            raise M365Error(
                "range_invalid",
                f"{height} x {width} values do not fit into {address.address()}",
            )
        target = address
    if target.rows * target.columns > MAX_WRITE_CELLS:
        raise M365Error("range_too_large", f"A write is limited to {MAX_WRITE_CELLS} cells")
    padded = [
        [*row, *([None] * (target.columns - len(row)))]
        for row in [*matrix, *([[]] * (target.rows - height))]
    ]
    return target, padded


def _range_matches(
    transport: GraphTransport,
    workbook: Workbook,
    worksheet: str,
    target: Bounds,
    expected: list[list[Any]],
    headers: Mapping[str, str] | None,
) -> bool:
    row = 0
    for values, formulas in read_matrix(
        transport, workbook, worksheet, target, headers=headers, with_formulas=True
    ):
        for index, actual in enumerate(values):
            formula_row = formulas[index] if formulas is not None else None
            if not rows_equal(expected[row], actual, formula_row):
                return False
            row += 1
    return row == len(expected)


def _confirm_range(
    transport: GraphTransport,
    workbook: Workbook,
    worksheet: str,
    target: Bounds,
    expected: list[list[Any]],
    session_headers: Mapping[str, str],
) -> bool:
    headers: Mapping[str, str] | None = session_headers
    for wait in _deadline_waits(transport):
        if wait:
            sleep(wait)
        try:
            if _range_matches(transport, workbook, worksheet, target, expected, headers):
                return True
        except M365Error:
            # The session may be what failed; the next attempt reads without it.
            headers = None
    return False


def write_range(context: OperationContext) -> dict[str, Any]:
    normalized = validated_write_input(context.input)
    matrix = _write_matrix(context, normalized)
    target, expected = write_target(normalized["address"], matrix)
    allow_formulas = normalized["allow_formulas"]
    outgoing = [
        [outgoing_cell(cell, allow_formulas=allow_formulas) for cell in row] for row in expected
    ]
    body = {"values": outgoing}
    if len(json.dumps(body, ensure_ascii=False).encode("utf-8")) > MAX_WRITE_BODY_BYTES:
        raise M365Error("range_too_large", "The values exceed the Excel request size limit")
    transport = context.transport
    workbook = resolve_workbook(transport, normalized["workbook"])
    worksheet = normalized["worksheet"]
    confirmation = "response"
    with workbook_session(transport, workbook) as session_headers:
        try:
            transport.request(
                "PATCH",
                f"{worksheet_path(workbook, worksheet)}/{range_segment(target)}",
                json_body=body,
                headers=session_headers,
                response_limit=WRITE_RESPONSE_LIMIT,
                retry="none",
                purpose="Excel range write",
            )
        except M365Error as exc:
            if not _is_unconfirmed(exc):
                raise worksheet_not_found(exc) from exc
            if not _confirm_range(
                transport, workbook, worksheet, target, expected, session_headers
            ):
                raise M365Error(
                    "timeout_unknown",
                    "Excel did not confirm the write and the range does not hold the requested "
                    "values; it was not repeated",
                    http_status=exc.http_status,
                    ambiguous=True,
                ) from exc
            confirmation = "reread"
    return {
        "workbook": workbook.summary(),
        "worksheet": worksheet,
        "address": target.address(),
        "row_count": target.rows,
        "column_count": target.columns,
        "confirmation": confirmation,
    }


# -- append_table_rows ------------------------------------------------------


def validated_append_input(raw: Mapping[str, Any]) -> dict[str, Any]:
    source = _data_source(raw)
    result: dict[str, Any] = {
        "workbook": raw.get("workbook"),
        "table": name_input(raw.get("table"), "table"),
        "allow_formulas": _flag(raw, "allow_formulas"),
        "dedupe_column": None,
    }
    if raw.get("dedupe_column") is not None:
        result["dedupe_column"] = name_input(raw.get("dedupe_column"), "dedupe_column")
    if source == "dataset":
        result["dataset_handle"] = name_input(raw.get("dataset_handle"), "dataset_handle")
    else:
        result["inline_rows"] = _inline_records(raw.get("inline_rows"))
    return result


def _append_records(
    context: OperationContext, normalized: Mapping[str, Any]
) -> list[dict[str, Any]]:
    if "inline_rows" in normalized:
        return [dict(row) for row in normalized["inline_rows"]]
    _, records = read_input_rows(
        context.payload, normalized["dataset_handle"], max_rows=MAX_APPEND_ROWS
    )
    if not records:
        raise M365Error("invalid_payload", "The input dataset has no rows")
    return records


def batch_keys(records: list[dict[str, Any]], dedupe_column: str | None) -> list[str]:
    """Return each row's dedupe key; reject empty and repeated keys before any write."""
    if dedupe_column is None:
        return [""] * len(records)
    keys: list[str] = []
    seen: dict[str, int] = {}
    for index, record in enumerate(records):
        key = key_text(record.get(dedupe_column))
        if not key:
            raise M365Error(
                "invalid_payload", f"Row {index + 1} has no value in dedupe column {dedupe_column}"
            )
        if key in seen:
            raise M365Error(
                "invalid_payload",
                f"Rows {seen[key] + 1} and {index + 1} repeat the dedupe key {key[:100]!r}",
            )
        seen[key] = index
        keys.append(key)
    return keys


def _table_headers(
    transport: GraphTransport, workbook: Workbook, table: str, headers: Mapping[str, str]
) -> list[str]:
    try:
        body = transport.get_json(
            f"{table_path(workbook, table)}/headerRowRange",
            query={"$select": "values"},
            headers=headers,
            purpose="Excel table header",
        )
    except M365Error as exc:
        raise table_not_found(exc) from exc
    values = body.get("values")
    if not isinstance(values, list) or len(values) != 1 or not isinstance(values[0], list):
        raise M365Error("provider_unavailable", "Excel returned an invalid table header")
    names = [header_name(value) for value in values[0]]
    if not names or not all(names) or len(set(names)) != len(names):
        raise M365Error("provider_unavailable", "Excel returned an invalid table header")
    return names


def _existing_rows(
    transport: GraphTransport,
    workbook: Workbook,
    table: str,
    columns: list[str],
    dedupe_column: str,
    wanted: set[str],
    headers: Mapping[str, str] | None,
) -> dict[str, list[dict[str, Any]]]:
    """Read the table body and return the rows whose key is in ``wanted``."""
    try:
        worksheet, bounds = range_address(
            transport,
            f"{table_path(workbook, table)}/dataBodyRange",
            headers=headers,
            purpose="Excel table body",
        )
    except M365Error as exc:
        raise table_not_found(exc) from exc
    if bounds.columns != len(columns):
        raise M365Error("provider_unavailable", "The table changed while it was being read")
    if bounds.rows > MAX_ROWS or bounds.rows * bounds.columns > MAX_CELLS:
        raise M365Error("range_too_large", "The table is too large to check dedupe_column")
    key_index = columns.index(dedupe_column)
    found: dict[str, list[dict[str, Any]]] = {}
    for values, formulas in read_matrix(
        transport, workbook, worksheet, bounds, headers=headers, with_formulas=True
    ):
        for index, row in enumerate(values):
            key = key_text(row[key_index])
            if key in wanted:
                formula_row = formulas[index] if formulas is not None else row
                found.setdefault(key, []).append(
                    {
                        "values": dict(zip(columns, row, strict=True)),
                        "formulas": dict(zip(columns, formula_row, strict=True)),
                    }
                )
    return found


def _record_matches(record: Mapping[str, Any], existing: list[dict[str, Any]]) -> bool:
    """Equal values in every column the input row provides."""
    for row in existing:
        if all(
            cells_equal(value, row["values"].get(name), row["formulas"].get(name))
            for name, value in record.items()
        ):
            return True
    return False


def _row_report(index: int, key: str, status: str, reason: str = "") -> dict[str, Any]:
    entry: dict[str, Any] = {"index": index, "key": key[:256], "status": status}
    if reason:
        entry["reason"] = reason
    return entry


def append_table_rows(context: OperationContext) -> dict[str, Any]:
    normalized = validated_append_input(context.input)
    records = _append_records(context, normalized)
    dedupe_column = normalized["dedupe_column"]
    keys = batch_keys(records, dedupe_column)
    table = normalized["table"]
    transport = context.transport
    workbook = resolve_workbook(transport, normalized["workbook"])
    with workbook_session(transport, workbook) as session_headers:
        columns = _table_headers(transport, workbook, table, session_headers)
        unknown = sorted({name for record in records for name in record} - set(columns))
        if unknown:
            raise M365Error(
                "invalid_payload",
                f"Columns not in table {table}: {', '.join(name[:64] for name in unknown[:10])}",
            )
        if dedupe_column is not None and dedupe_column not in columns:
            raise M365Error(
                "invalid_payload", f"dedupe_column {dedupe_column} is not a table column"
            )
        statuses: list[dict[str, Any]] = []
        to_add: list[int] = []
        if dedupe_column is not None:
            existing = _existing_rows(
                transport, workbook, table, columns, dedupe_column, set(keys), session_headers
            )
            conflicts: list[int] = []
            for index, (record, key) in enumerate(zip(records, keys, strict=True)):
                if key not in existing:
                    to_add.append(index)
                elif _record_matches(record, existing[key]):
                    statuses.append(_row_report(index, key, "skipped", "already_present"))
                else:
                    conflicts.append(index)
            if conflicts:
                reported = ", ".join(
                    repr(keys[index][:100]) for index in conflicts[:MAX_REPORTED_KEYS]
                )
                conflict_set = set(conflicts)
                skipped_set = {row["index"] for row in statuses}
                rows = [
                    _row_report(
                        index,
                        key,
                        "conflict"
                        if index in conflict_set
                        else "skipped"
                        if index in skipped_set
                        else "not_written",
                    )
                    for index, key in enumerate(keys)
                ]
                raise ExcelRowsError(
                    "dedupe_conflict",
                    f"{len(conflicts)} row(s) have a dedupe key that is already present with "
                    f"different values: {reported}. Nothing was written.",
                    result=_append_result(
                        workbook, table, records, rows, appended=0, skipped=0, changed=False
                    ),
                )
        else:
            to_add = list(range(len(records)))
        skipped = len(records) - len(to_add)
        if not to_add:
            result = _append_result(
                workbook, table, records, statuses, appended=0, skipped=skipped, changed=False
            )
            result["definitely_no_external_effect"] = True
            return result
        allow_formulas = normalized["allow_formulas"]
        values = [
            [
                outgoing_cell(records[index].get(column), allow_formulas=allow_formulas)
                for column in columns
            ]
            for index in to_add
        ]
        confirmation = "response"
        first_row_index: int | None = None
        try:
            response = transport.request(
                "POST",
                f"{table_path(workbook, table)}/rows/add",
                json_body={"index": None, "values": values},
                headers=session_headers,
                response_limit=WRITE_RESPONSE_LIMIT,
                retry="none",
                purpose="Excel table append",
            )
            index_value = response.json().get("index")
            if isinstance(index_value, int) and not isinstance(index_value, bool):
                first_row_index = index_value
        except M365Error as exc:
            if not _is_unconfirmed(exc):
                raise table_not_found(exc) from exc
            confirmed = _confirm_append(
                transport,
                workbook,
                table,
                columns,
                dedupe_column,
                {index: (records[index], keys[index]) for index in to_add},
                session_headers,
            )
            if len(confirmed) != len(to_add):
                rows = sorted(
                    [
                        *statuses,
                        *(
                            _row_report(
                                index,
                                keys[index],
                                "confirmed" if index in confirmed else "unknown",
                            )
                            for index in to_add
                        ),
                    ],
                    key=lambda row: row["index"],
                )
                message = (
                    "Excel did not confirm the append; it was not repeated"
                    if dedupe_column is None
                    else f"Excel did not confirm the append; {len(confirmed)} of {len(to_add)} "
                    "row(s) are confirmed and the rest may still land; it was not repeated"
                )
                raise ExcelRowsError(
                    "timeout_unknown",
                    message,
                    http_status=exc.http_status,
                    ambiguous=True,
                    result=_append_result(
                        workbook,
                        table,
                        records,
                        rows,
                        appended=len(confirmed),
                        skipped=skipped,
                        changed=True,
                    ),
                ) from exc
            confirmation = "reread"
    rows = sorted(
        [*statuses, *(_row_report(index, keys[index], "appended") for index in to_add)],
        key=lambda row: row["index"],
    )
    result = _append_result(
        workbook, table, records, rows, appended=len(to_add), skipped=skipped, changed=True
    )
    result["confirmation"] = confirmation
    result["first_row_index"] = first_row_index
    return result


def _confirm_append(
    transport: GraphTransport,
    workbook: Workbook,
    table: str,
    columns: list[str],
    dedupe_column: str | None,
    pending: dict[int, tuple[dict[str, Any], str]],
    session_headers: Mapping[str, str],
) -> set[int]:
    """Return the input indexes whose rows a bounded re-read finds with equal values."""
    confirmed: set[int] = set()
    if dedupe_column is None:
        # Without a key, absence and presence prove nothing.
        return confirmed
    headers: Mapping[str, str] | None = session_headers
    wanted = {key for _, key in pending.values()}
    for wait in _deadline_waits(transport):
        if wait:
            sleep(wait)
        try:
            existing = _existing_rows(
                transport, workbook, table, columns, dedupe_column, wanted, headers
            )
        except M365Error:
            headers = None
            continue
        for index, (record, key) in pending.items():
            if key in existing and _record_matches(record, existing[key]):
                confirmed.add(index)
        if len(confirmed) == len(pending):
            break
    return confirmed


def _append_result(
    workbook: Workbook,
    table: str,
    records: list[dict[str, Any]],
    rows: list[dict[str, Any]],
    *,
    appended: int,
    skipped: int,
    changed: bool,
) -> dict[str, Any]:
    return {
        "workbook": workbook.summary(),
        "table": table,
        "row_count": len(records),
        "appended_count": appended,
        "skipped_count": skipped,
        "changed": changed,
        # "none": nothing was sent; "response"/"reread" are set after a write.
        "confirmation": "none",
        "first_row_index": None,
        "rows": rows,
    }


__all__ = [
    "APPEND_FIELDS",
    "WRITE_RANGE_FIELDS",
    "append_table_rows",
    "batch_keys",
    "validated_append_input",
    "validated_write_input",
    "write_range",
    "write_target",
]
