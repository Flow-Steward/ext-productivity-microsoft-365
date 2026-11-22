"""``find_free_time``: common free slots of up to 20 people through ``getSchedule``.

Busy time is the union of ``scheduleItems`` (busy, tentative, out of office, unknown)
and the ``availabilityView`` digits 1-3. The view is requested in 15-minute steps on a
window aligned to quarter hours (UTC), so it lines up with slots on quarter hours.

An address whose schedule has an ``error``, no view, a view shorter than the window,
or any character other than 0-4 (``?`` and the like) is never treated as free: it is
listed in ``unresolved_addresses``, the result is ``complete: false`` and
``common_slots`` is empty. The resolved addresses keep their busy times for display.
"""

from __future__ import annotations

import bisect
import re
from datetime import UTC, date, datetime, time, timedelta, tzinfo
from typing import Any

from m365.connection import mapping
from m365.errors import M365Error
from m365.runtime import OperationContext
from mail.common import Mailbox, bool_input, int_input, recipients_input

from .common import (
    MAX_SCHEDULE_ADDRESSES,
    MAX_SCHEDULE_DAYS,
    PREFER_UTC,
    as_read_failure,
    client_for,
    iso_local,
    iso_z,
    parse_graph_time,
    window_input,
)
from .zones import to_local, zone_input

VIEW_INTERVAL_MINUTES = 15
SCHEDULE_RESPONSE_LIMIT = 8 * 1024 * 1024
MAX_BUSY_ROWS = 1000
MAX_SLOTS = 100
#: ``scheduleItems`` statuses that block a slot; ``free`` and ``workingElsewhere`` do not.
BUSY_STATUSES = frozenset({"busy", "tentative", "oof", "unknown"})
#: ``availabilityView``: 0 free, 1 tentative, 2 busy, 3 out of office, 4 working elsewhere.
VIEW_BUSY = frozenset("123")
VIEW_KNOWN = frozenset("01234")
_VIEW_STATUS = {"1": "tentative", "2": "busy", "3": "oof"}
_CLOCK = re.compile(r"^([01]\d|2[0-3]):([0-5]\d)$")
FREE_TIME_FIELDS = frozenset(
    {
        "addresses",
        "start",
        "end",
        "time_zone",
        "duration_minutes",
        "working_hours_start",
        "working_hours_end",
        "slot_interval_minutes",
        "include_weekends",
        "max_slots",
    }
)

Interval = tuple[datetime, datetime]


# -- input ------------------------------------------------------------------------------


def _addresses(value: Any) -> list[str]:
    rows = recipients_input(value, "addresses", maximum=MAX_SCHEDULE_ADDRESSES)
    if not rows:
        raise M365Error("invalid_payload", "addresses needs at least one address")
    return [row["emailAddress"]["address"] for row in rows]


def _clock(value: Any, field: str, default: str) -> time:
    raw = default if value in (None, "") else value
    match = _CLOCK.match(raw) if isinstance(raw, str) else None
    if not match or int(match.group(2)) % VIEW_INTERVAL_MINUTES:
        raise M365Error(
            "invalid_payload", f"{field} must be HH:MM on a quarter hour, for example 09:00"
        )
    return time(int(match.group(1)), int(match.group(2)))


def _multiple(value: int, field: str, step: int) -> int:
    if value % step:
        raise M365Error("invalid_payload", f"{field} must be a multiple of {step}")
    return value


def _floor_quarter(moment: datetime) -> datetime:
    return moment.replace(minute=moment.minute - moment.minute % VIEW_INTERVAL_MINUTES, second=0)


def _ceil_quarter(moment: datetime) -> datetime:
    floor = _floor_quarter(moment)
    return floor if floor == moment else floor + timedelta(minutes=VIEW_INTERVAL_MINUTES)


# -- the operation ------------------------------------------------------------------------


def find_free_time(ctx: OperationContext) -> dict[str, Any]:
    raw = ctx.input
    zone_name, zone = zone_input(raw.get("time_zone"), "time_zone", default="UTC")
    addresses = _addresses(raw.get("addresses"))
    start, end = window_input(raw, zone, max_days=MAX_SCHEDULE_DAYS, strict=True)
    duration = _multiple(
        int_input(
            raw.get("duration_minutes"), "duration_minutes", default=30, minimum=5, maximum=1440
        ),
        "duration_minutes",
        5,
    )
    interval = _multiple(
        int_input(
            raw.get("slot_interval_minutes"),
            "slot_interval_minutes",
            default=30,
            minimum=15,
            maximum=240,
        ),
        "slot_interval_minutes",
        VIEW_INTERVAL_MINUTES,
    )
    day_start = _clock(raw.get("working_hours_start"), "working_hours_start", "09:00")
    day_end = _clock(raw.get("working_hours_end"), "working_hours_end", "17:00")
    if day_end <= day_start:
        raise M365Error(
            "invalid_payload", "working_hours_end must be later than working_hours_start"
        )
    weekends = bool_input(raw.get("include_weekends"), "include_weekends", default=False)
    max_slots = int_input(
        raw.get("max_slots"), "max_slots", default=20, minimum=1, maximum=MAX_SLOTS
    )

    view_start, view_end = _floor_quarter(start), _ceil_quarter(end)
    rows = _get_schedule(ctx, addresses, view_start, view_end)
    view_length = int((view_end - view_start) / timedelta(minutes=VIEW_INTERVAL_MINUTES))
    schedules = [
        _schedule(address, rows.get(address.lower()), view_start, view_length, zone)
        for address in addresses
    ]
    unresolved = [row["address"] for row in schedules if not row["resolved"]]
    slots: list[dict[str, str]] = []
    truncated = False
    if not unresolved:
        busy = _merge([interval for row in schedules for interval in row.pop("_busy")])
        slots, truncated = _common_slots(
            busy,
            start=start,
            end=end,
            zone=zone,
            day_start=day_start,
            day_end=day_end,
            duration=timedelta(minutes=duration),
            step=timedelta(minutes=interval),
            weekends=bool(weekends),
            limit=max_slots,
        )
    for row in schedules:
        row.pop("_busy", None)
    return {
        "complete": not unresolved,
        "common_slots": slots,
        "slot_count": len(slots),
        "truncated": truncated,
        "unresolved_addresses": unresolved,
        "schedules": schedules,
        "time_zone": zone_name,
        "window_start_utc": iso_z(start),
        "window_end_utc": iso_z(end),
        "duration_minutes": duration,
    }


def _get_schedule(
    ctx: OperationContext, addresses: list[str], view_start: datetime, view_end: datetime
) -> dict[str, dict[str, Any]]:
    client = client_for(ctx.transport, ctx.connection)
    body = {
        "schedules": addresses,
        "startTime": {"dateTime": iso_local(view_start.astimezone(UTC)), "timeZone": "UTC"},
        "endTime": {"dateTime": iso_local(view_end.astimezone(UTC)), "timeZone": "UTC"},
        "availabilityViewInterval": VIEW_INTERVAL_MINUTES,
    }
    with as_read_failure():
        # getSchedule is a POST that only reads, so the transport may repeat it.
        response = ctx.transport.request(
            "POST",
            f"{client.base}/calendar/getSchedule",
            json_body=body,
            headers=Mailbox.headers(PREFER_UTC),
            response_limit=SCHEDULE_RESPONSE_LIMIT,
            retry="idempotent",
            purpose="Outlook free/busy",
            ok_statuses=(200,),
        )
    values = response.json().get("value")
    if not isinstance(values, list):
        raise M365Error("provider_unavailable", "Microsoft Graph returned an invalid schedule")
    rows: dict[str, dict[str, Any]] = {}
    for value in values:
        row = mapping(value)
        schedule_id = str(row.get("scheduleId") or "").strip().lower()
        if schedule_id:
            rows[schedule_id] = row
    return rows


def _schedule(
    address: str,
    row: dict[str, Any] | None,
    view_start: datetime,
    view_length: int,
    zone: tzinfo,
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "address": address,
        "resolved": False,
        "unresolved_reason": None,
        "error_message": None,
        "availability_view": "",
        "busy": [],
        "_busy": [],
    }
    if row is None:
        result["unresolved_reason"] = "no_schedule"
        return result
    error = row.get("error")
    if error:
        result["unresolved_reason"] = "error"
        result["error_message"] = str(mapping(error).get("message") or "")[:256] or None
        return result
    view = str(row.get("availabilityView") or "")
    result["availability_view"] = view[: view_length + 1]
    if not view or len(view) < view_length:
        result["unresolved_reason"] = "no_data"
        return result
    if any(char not in VIEW_KNOWN for char in view[:view_length]):
        result["unresolved_reason"] = "unknown_availability"
        return result
    busy: list[tuple[datetime, datetime, str]] = []
    for item in row.get("scheduleItems") or []:
        item = mapping(item)
        status = str(item.get("status") or "unknown")
        if status not in BUSY_STATUSES:
            continue
        item_start = parse_graph_time(item.get("start"))
        item_end = parse_graph_time(item.get("end"))
        if item_start is None or item_end is None or item_end <= item_start:
            result["unresolved_reason"] = "invalid_schedule"
            return result
        busy.append((item_start, item_end, status))
    items = list(busy)
    for span in _view_spans(view[:view_length], view_start):
        # A view span already covered by a schedule item adds nothing to show.
        if not any(start <= span[0] and span[1] <= end for start, end, _ in items):
            busy.append(span)
    result["resolved"] = True
    result["_busy"] = [(item_start, item_end) for item_start, item_end, _ in busy]
    result["busy"] = [
        {
            "start_utc": iso_z(item_start),
            "end_utc": iso_z(item_end),
            "start_local": iso_local(to_local(item_start, zone)),
            "end_local": iso_local(to_local(item_end, zone)),
            "status": status,
        }
        for item_start, item_end, status in sorted(busy)[:MAX_BUSY_ROWS]
    ]
    return result


def _view_spans(view: str, view_start: datetime) -> list[tuple[datetime, datetime, str]]:
    """Contiguous busy runs of the view (same digit) as (start, end, status)."""
    step = timedelta(minutes=VIEW_INTERVAL_MINUTES)
    spans: list[tuple[datetime, datetime, str]] = []
    for index, char in enumerate(view):
        if char not in VIEW_BUSY:
            continue
        start = view_start + index * step
        status = _VIEW_STATUS[char]
        if spans and spans[-1][1] == start and spans[-1][2] == status:
            spans[-1] = (spans[-1][0], start + step, status)
        else:
            spans.append((start, start + step, status))
    return spans


def _merge(intervals: list[Interval]) -> list[Interval]:
    merged: list[Interval] = []
    for start, end in sorted(intervals):
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def _overlaps(busy: list[Interval], starts: list[datetime], start: datetime, end: datetime) -> bool:
    index = bisect.bisect_left(starts, end) - 1
    return index >= 0 and busy[index][1] > start


def _common_slots(
    busy: list[Interval],
    *,
    start: datetime,
    end: datetime,
    zone: tzinfo,
    day_start: time,
    day_end: time,
    duration: timedelta,
    step: timedelta,
    weekends: bool,
    limit: int,
) -> tuple[list[dict[str, str]], bool]:
    starts = [interval[0] for interval in busy]
    slots: list[dict[str, str]] = []
    day: date = to_local(start, zone).date()
    last_day = to_local(end, zone).date()
    while day <= last_day:
        if weekends or day.weekday() < 5:
            local = datetime.combine(day, day_start)
            closing = datetime.combine(day, day_end)
            while local + duration <= closing:
                slot_start = local.replace(tzinfo=zone).astimezone(UTC)
                slot_end = slot_start + duration
                if (
                    start <= slot_start
                    and slot_end <= end
                    and to_local(slot_end, zone) <= closing
                    and not _overlaps(busy, starts, slot_start, slot_end)
                ):
                    if len(slots) == limit:
                        return slots, True
                    slots.append(
                        {
                            "start_utc": iso_z(slot_start),
                            "end_utc": iso_z(slot_end),
                            "start_local": iso_local(to_local(slot_start, zone)),
                            "end_local": iso_local(to_local(slot_end, zone)),
                        }
                    )
                local += step
        day += timedelta(days=1)
    return slots, False


__all__ = ["FREE_TIME_FIELDS", "find_free_time"]
