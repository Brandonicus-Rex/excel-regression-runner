"""Command-line entry point.

Examples::

    # Run every test sheet, write results to tests_results.xlsx
    python -m runner.cli samples/regression_tests.xlsx

    # Run rows 2-5 of one sheet (this is what the Excel macro calls)
    python -m runner.cli tests.xlsm --sheet Lifecycle --rows 2-5 --output out.xlsx

Exit codes: 0 = all passed, 1 = at least one test failed, 2 = setup error
(bad arguments, missing sheet or column, unreadable workbook).
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from openpyxl import load_workbook

from runner.client import ApiClient, ClientConfig
from runner.engine import run_case
from runner.workbook import (
    WorkbookError,
    load_cases,
    parse_row_spec,
    select_sheets,
    write_outcome,
    write_summary,
)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="excel-regression-runner",
                                description="Run API regression tests defined in an Excel workbook.")
    p.add_argument("workbook", type=Path, help="Workbook with test sheets (.xlsx or .xlsm)")
    p.add_argument("--sheet", action="append", dest="sheets",
                   help="Sheet to run (repeatable). Default: every test sheet")
    p.add_argument("--rows", help="Rows to run, e.g. 2-5,8. Default: all rows")
    p.add_argument("--output", type=Path,
                   help="Where to save results. Default: <workbook>_results.xlsx next to the input")
    p.add_argument("--base-url", help="API base URL (default: $API_BASE_URL)")
    p.add_argument("-v", "--verbose", action="store_true")
    return p


def default_output(path: Path) -> Path:
    """tests.xlsx -> tests_results.xlsx. Results never overwrite the input."""
    return path.with_name(f"{path.stem}_results.xlsx")


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    logging.basicConfig(level=logging.INFO if args.verbose else logging.WARNING,
                        format="%(levelname)s %(message)s")
    try:
        from dotenv import load_dotenv  # optional: read settings from .env
        load_dotenv()
    except ImportError:
        pass

    output = args.output or default_output(args.workbook)
    try:
        rows = parse_row_spec(args.rows)
        # keep_vba=False: the results copy is a plain .xlsx even if the input
        # is a macro-enabled .xlsm, so it can't carry or trigger macros.
        wb = load_workbook(args.workbook)
        sheets = select_sheets(wb, args.sheets)
        cases_by_sheet = [(ws, load_cases(ws, rows)) for ws in sheets]
        config = ClientConfig.from_env(args.base_url)
    except (WorkbookError, RuntimeError, OSError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    outcomes = []
    with ApiClient(config) as client:
        for ws, cases in cases_by_sheet:
            for case in cases:
                outcome = run_case(client, case)
                write_outcome(ws, outcome)
                outcomes.append(outcome)

    write_summary(wb, outcomes)
    try:
        wb.save(output)
    except OSError as exc:
        # Most often the output file is open in Excel.
        print(f"error: could not save {output}: {exc}", file=sys.stderr)
        return 2

    passed = sum(o.passed for o in outcomes)
    print(f"{passed}/{len(outcomes)} passed -> {output}")
    for o in outcomes:
        if not o.passed:
            print(f"  FAIL {o.case.sheet}!{o.case.row} {o.case.case_id}: {'; '.join(o.details)}")
    return 0 if passed == len(outcomes) else 1


if __name__ == "__main__":
    sys.exit(main())
