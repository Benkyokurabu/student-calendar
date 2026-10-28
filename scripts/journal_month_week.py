"""Keep retired month/week fields empty without losing carryover accounting."""
from datetime import date
import json
import re

from journal_slots import slot_columns, counter_address

ACCOUNTING_MONTHS_CELL = "HD4"  # Existing hidden helper column, outside lesson forms.
STORED_WEEKS_CELL = "HD5"
MONTH_WEEK_BLANK_FROM = date(2026, 10, 28)


def _stored_values(ws, address):
    raw = ws[address].value
    if raw in (None, ""):
        return {}
    months = json.loads(raw)
    if not isinstance(months, dict):
        raise ValueError(f"{ws.title}!{address}: invalid stored header values")
    return months


def slot_accounting_month(ws, left_col):
    """Read old visible months as well as the new hidden carryover metadata."""
    value = ws.cell(3, left_col + 3).value
    if isinstance(value, str) and value == "=$E$3":
        value = ws["E3"].value
    if value not in (None, "") and not str(value).startswith("="):
        return value
    return _stored_values(ws, ACCOUNTING_MONTHS_CELL).get(str(left_col))


def set_slot_accounting_month(ws, left_col, month):
    months = _stored_values(ws, ACCOUNTING_MONTHS_CELL)
    months[str(left_col)] = month
    ws[ACCOUNTING_MONTHS_CELL] = json.dumps(months, separators=(",", ":"))
    ws.column_dimensions["HD"].hidden = True
    ws.cell(3, left_col + 3).value = None


def slot_week_value(ws, left_col):
    value = ws.cell(3, left_col + 5).value
    if value not in (None, "") and not str(value).startswith("="):
        return value
    return _stored_values(ws, STORED_WEEKS_CELL).get(str(left_col))


def set_slot_week_value(ws, left_col, value):
    weeks = _stored_values(ws, STORED_WEEKS_CELL)
    weeks[str(left_col)] = value
    ws[STORED_WEEKS_CELL] = json.dumps(weeks, ensure_ascii=False, separators=(",", ":"))
    ws.column_dimensions["HD"].hidden = True
    ws.cell(3, left_col + 5).value = None


def clear_month_week_headers(ws, *, from_date=None):
    """Clear values/formulas only; preserve printed labels, styles and old lessons."""
    match = re.fullmatch(r"(\d{4})-(\d{2})", ws.title)
    period = tuple(map(int, match.groups())) if match else None
    columns = slot_columns(ws)
    # Resolve all references before clearing E3, which later month cells may use.
    months = {left: slot_accounting_month(ws, left) for left in columns}
    weeks = {left: slot_week_value(ws, left) for left in columns}
    for left in columns:
        if from_date is not None and period is not None:
            if period < (from_date.year, from_date.month):
                continue
            if period == (from_date.year, from_date.month):
                days = [ws.cell(row, left).value for row in (11, 31, 51)]
                if not any(str(day).isdigit() and int(day) >= from_date.day for day in days):
                    continue
        if months[left] is not None:
            set_slot_accounting_month(ws, left, months[left])
        if weeks[left] is not None:
            set_slot_week_value(ws, left, weeks[left])
        if left == 2 and isinstance(weeks[left], (int, float)):
            ws[counter_address(ws, 3)] = weeks[left]
        ws.cell(3, left + 3).value = None
        ws.cell(3, left + 5).value = None
