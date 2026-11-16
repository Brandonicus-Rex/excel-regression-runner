"""Unit tests for reading and writing workbooks. No server needed."""

import pytest
from openpyxl import Workbook, load_workbook

from runner.workbook import (
    WorkbookError,
    Outcome,
    header_map,
    load_cases,
    parse_params,
    parse_row_spec,
    parse_status,
    select_sheets,
    write_outcome,
    write_summary,
)
from samples.make_sample import build_sample


@pytest.fixture
def sample(tmp_path):
    return load_workbook(build_sample(tmp_path / "sample.xlsx"))


def test_select_sheets_skips_config_and_summary(sample):
    assert [ws.title for ws in select_sheets(sample)] == ["Lifecycle", "Usage & Notifications", "Negative"]


def test_select_sheets_unknown_name(sample):
    with pytest.raises(WorkbookError):
        select_sheets(sample, ["Nope"])


def test_header_map_reports_missing_columns():
    wb = Workbook()
    wb.active.append(["ID", "Action"])
    with pytest.raises(WorkbookError, match="missing column"):
        header_map(wb.active)


def test_load_cases_parses_row(sample):
    case = load_cases(sample["Lifecycle"])[1]  # LC-01 (row 2 is the LC-00 setup row)
    assert (case.row, case.action, case.phone, case.expected_status) == (3, "create", "5550100101", (201,))
    assert case.params == {"plan": "basic", "features": "voicemail"}


def test_load_cases_row_filter(sample):
    rows = [c.row for c in load_cases(sample["Lifecycle"], {3, 5})]
    assert rows == [3, 5]


def test_bad_expected_status(sample):
    sample["Lifecycle"]["F3"] = "two hundred"
    with pytest.raises(WorkbookError, match="not a valid HTTP status"):
        load_cases(sample["Lifecycle"])


@pytest.mark.parametrize("text,expected", [
    ("200", (200,)),
    ("200.0", (200,)),                     # Excel's float storage
    ("200,404,409", (200, 404, 409)),
    ("200, 404 / 409 | 409", (200, 404, 409)),  # mixed separators, duplicate dropped
    ("200404409", (200, 404, 409)),        # Excel turned "200,404,409" into a number
    ("", ()),                              # blank = don't check
])
def test_parse_status(text, expected):
    assert parse_status(text) == expected


@pytest.mark.parametrize("bad", ["abc", "99", "600", "2004", "200,fail"])
def test_parse_status_rejects_invalid(bad):
    with pytest.raises(ValueError):
        parse_status(bad)


def test_every_sheet_cleans_up_at_start_and_end(sample):
    """Each sheet closes every number it creates both before and after its tests."""
    for ws in select_sheets(sample):
        cases = load_cases(ws)
        created = {c.phone for c in cases if c.action == "create" and c.expected_status == (201,)}
        setup = [c for c in cases[:len(created)]]
        # Start: tolerant setup rows close every number before anything else runs.
        assert {c.phone for c in setup} == created, ws.title
        for c in setup:
            assert (c.action, c.expected_status, c.expected_response) == ("close", (200, 404, 409), ""), c.case_id
        # End: every created number is closed again after its last use.
        for phone in created:
            last_use = max(i for i, c in enumerate(cases) if c.phone == phone)
            assert cases[last_use].action in ("close", "get"), (ws.title, phone)
            assert any(c.action == "close" and c.phone == phone for c in cases[len(setup):]), (ws.title, phone)


@pytest.mark.parametrize("text,expected", [
    ("a=1\nb = 2", {"a": "1", "b": "2"}),
    ("a=1; b=x=y", {"a": "1", "b": "x=y"}),   # only the first '=' splits
    ("", {}),
])
def test_parse_params(text, expected):
    assert parse_params(text) == expected


def test_parse_row_spec():
    assert parse_row_spec("2-4, 7") == {2, 3, 4, 7}
    assert parse_row_spec("") is None
    for bad in ("x", "5-2"):
        with pytest.raises(WorkbookError):
            parse_row_spec(bad)


def test_write_outcome_and_summary(sample):
    ws = sample["Lifecycle"]
    case = load_cases(ws, {2})[0]
    write_outcome(ws, Outcome(case, 500, "error=boom", False, ["expected status 201, got 500"]))
    assert ws["J2"].value == "FAIL" and ws["H2"].value == 500

    write_summary(sample, [Outcome(case, 500, "", False, ["x"])])
    summary = sample["Summary"]
    assert sample.sheetnames[0] == "Summary"
    assert [c.value for c in summary[5]] == ["TOTAL", 1, 0, 1]
