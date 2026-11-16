"""Reading test cases from, and writing results to, an Excel workbook.

Each test sheet has a header row (row 1) and one test case per row below it.
Columns are found by header name, so their order doesn't matter:

    Input:    ID | Description | Action | Phone | Params
    Expected: Expected Status | Expected Response
    Output:   Actual Status | Actual Response | Result | Details

Sheets named "Summary" or "Config" are never treated as test sheets.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Iterable

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

INPUT_HEADERS = ["ID", "Description", "Action", "Phone", "Params"]
EXPECTED_HEADERS = ["Expected Status", "Expected Response"]
OUTPUT_HEADERS = ["Actual Status", "Actual Response", "Result", "Details"]
ALL_HEADERS = INPUT_HEADERS + EXPECTED_HEADERS + OUTPUT_HEADERS

NON_TEST_SHEETS = {"summary", "config"}

PASS_FILL = PatternFill("solid", start_color="C6EFCE")  # Excel's standard "good" green
FAIL_FILL = PatternFill("solid", start_color="FFC7CE")  # Excel's standard "bad" red
WRAP = Alignment(wrap_text=True, vertical="top")


class WorkbookError(Exception):
    """Raised when a sheet is missing required columns or input is invalid."""


@dataclass
class TestCase:
    """One row of a test sheet."""

    __test__ = False  # tell pytest this isn't a test class despite the name

    sheet: str
    row: int
    case_id: str
    description: str
    action: str
    phone: str
    params: dict[str, str]
    expected_status: tuple[int, ...]  # any of these passes; empty = don't check
    expected_response: str


@dataclass
class Outcome:
    """What happened when a test case ran."""

    case: TestCase
    actual_status: int | None
    actual_response: str
    passed: bool
    details: list[str] = field(default_factory=list)


# ---- Helpers -------------------------------------------------------------

def select_sheets(wb: Workbook, names: Iterable[str] | None = None) -> list[Worksheet]:
    """Return the requested sheets, or every test sheet if none are named.

    Raises:
        WorkbookError: If a requested sheet doesn't exist.
    """
    if names:
        missing = [n for n in names if n not in wb.sheetnames]
        if missing:
            raise WorkbookError(f"sheet(s) not found: {', '.join(missing)}")
        return [wb[n] for n in names]
    return [ws for ws in wb.worksheets if ws.title.lower() not in NON_TEST_SHEETS]


def header_map(ws: Worksheet) -> dict[str, int]:
    """Map header names in row 1 to column numbers.

    Raises:
        WorkbookError: If any required header is missing.
    """
    found = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
    missing = [h for h in ALL_HEADERS if h not in found]
    if missing:
        raise WorkbookError(f"sheet '{ws.title}' is missing column(s): {', '.join(missing)}")
    return found


def parse_params(text: str | None) -> dict[str, str]:
    """Parse a Params cell: ``key=value`` pairs separated by new lines or ';'."""
    params: dict[str, str] = {}
    for part in (text or "").replace(";", "\n").splitlines():
        if "=" in part:
            key, value = part.split("=", 1)
            params[key.strip().lower()] = value.strip()
    return params


def parse_row_spec(spec: str | None) -> set[int] | None:
    """Parse a row selection like ``"2-5,8"`` into ``{2, 3, 4, 5, 8}``.

    Returns None (meaning "all rows") for an empty spec.

    Raises:
        WorkbookError: On malformed input.
    """
    if not spec:
        return None
    rows: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        try:
            if "-" in part:
                start, end = (int(x) for x in part.split("-", 1))
                if start > end:
                    raise ValueError
                rows.update(range(start, end + 1))
            elif part:
                rows.add(int(part))
        except ValueError:
            raise WorkbookError(f"invalid row selection '{part}' (use e.g. 2-5,8)") from None
    return rows


def parse_status(text: str) -> tuple[int, ...]:
    """Parse an Expected Status cell into the accepted HTTP status codes.

    Accepts one code (``200``) or several (``200,404,409``; spaces, ``;``,
    ``/`` or ``|`` also work). Blank means "don't check the status".

    Excel (US locale) turns ``200,404,409`` into the number 200404409, and may
    store 200 as 200.0, so numbers are normalized and long digit strings are
    split back into 3-digit codes.

    Raises:
        ValueError: If any part isn't a valid HTTP status code.
    """
    codes: list[int] = []
    for token in re.split(r"[,;/|\s]+", text.strip()):
        if not token:
            continue
        digits = str(int(float(token)))  # "200.0" -> "200"; rejects non-numbers
        if len(digits) > 3 and len(digits) % 3 == 0:
            parts = [digits[i:i + 3] for i in range(0, len(digits), 3)]
        else:
            parts = [digits]
        for part in parts:
            code = int(part)
            if not 100 <= code <= 599:
                raise ValueError(part)
            codes.append(code)
    return tuple(dict.fromkeys(codes))  # drop duplicates, keep order


# ---- Reading -------------------------------------------------------------

def load_cases(ws: Worksheet, rows: set[int] | None = None) -> list[TestCase]:
    """Read test cases from a sheet, in row order.

    Rows with an empty Action cell are skipped (blank lines, section notes).
    """
    cols = header_map(ws)
    cases: list[TestCase] = []
    for r in range(2, ws.max_row + 1):
        if rows is not None and r not in rows:
            continue

        def cell(name: str) -> str:
            value = ws.cell(row=r, column=cols[name]).value
            return "" if value is None else str(value).strip()

        action = cell("Action").lower()
        if not action:
            continue

        status_text = cell("Expected Status")
        try:
            expected_status = parse_status(status_text)
        except (ValueError, OverflowError):
            raise WorkbookError(
                f"{ws.title}!row {r}: Expected Status '{status_text}' is not a valid HTTP status "
                "(use e.g. 200, or 200,404 for any of several)") from None

        cases.append(TestCase(
            sheet=ws.title, row=r, case_id=cell("ID"), description=cell("Description"),
            action=action, phone=cell("Phone"), params=parse_params(cell("Params")),
            expected_status=expected_status, expected_response=cell("Expected Response"),
        ))
    return cases


# ---- Writing -------------------------------------------------------------

def write_outcome(ws: Worksheet, outcome: Outcome) -> None:
    """Write one outcome into its row's output columns and color the Result."""
    cols = header_map(ws)
    r = outcome.case.row
    ws.cell(row=r, column=cols["Actual Status"], value=outcome.actual_status)
    ws.cell(row=r, column=cols["Actual Response"], value=outcome.actual_response).alignment = WRAP
    result = ws.cell(row=r, column=cols["Result"], value="PASS" if outcome.passed else "FAIL")
    result.fill = PASS_FILL if outcome.passed else FAIL_FILL
    result.font = Font(bold=True)
    ws.cell(row=r, column=cols["Details"], value="\n".join(outcome.details)).alignment = WRAP


def write_summary(wb: Workbook, outcomes: list[Outcome]) -> None:
    """Replace the Summary sheet with totals per sheet and a list of failures."""
    if "Summary" in wb.sheetnames:
        del wb["Summary"]
    ws = wb.create_sheet("Summary", 0)  # first tab, so it's what you see on open

    ws.append(["Run at", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws.append([])
    ws.append(["Sheet", "Run", "Passed", "Failed"])
    for c in ws[3]:
        c.font = Font(bold=True)

    by_sheet: dict[str, list[Outcome]] = {}
    for o in outcomes:
        by_sheet.setdefault(o.case.sheet, []).append(o)
    for sheet, items in by_sheet.items():
        passed = sum(o.passed for o in items)
        ws.append([sheet, len(items), passed, len(items) - passed])

    total_passed = sum(o.passed for o in outcomes)
    ws.append(["TOTAL", len(outcomes), total_passed, len(outcomes) - total_passed])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)

    failures = [o for o in outcomes if not o.passed]
    if failures:
        ws.append([])
        ws.append(["Failures", "Row", "ID", "Details"])
        for c in ws[ws.max_row]:
            c.font = Font(bold=True)
        for o in failures:
            ws.append([o.case.sheet, o.case.row, o.case.case_id, "; ".join(o.details)])

    for col, width in zip("ABCD", (24, 10, 10, 80)):
        ws.column_dimensions[col].width = width
