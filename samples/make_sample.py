"""Generate samples/regression_tests.xlsx: a ready-to-run example workbook.

Run from the project root::

    python -m samples.make_sample

All phone numbers are fictional (the 555-01xx range reserved for examples).
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from runner.workbook import ALL_HEADERS, EXPECTED_HEADERS, INPUT_HEADERS

OUTPUT = Path(__file__).resolve().parent / "regression_tests.xlsx"

# Cleanup rows pass whether they closed the account (200), it was already
# closed (409), or it never existed (404): either way the data ends up clean.
# Each sheet cleans up at the start (leftovers from an interrupted round)
# and at the end (so the next round starts clean).
CLEANUP_STATUS = "200,404,409"

# Each row: ID, Description, Action, Phone, Params, Expected Status, Expected Response
SHEETS: dict[str, list[tuple]] = {
    "Lifecycle": [
        ("LC-00", "Setup: close 5550100101 if left over from an earlier round", "close", "5550100101", "",
         CLEANUP_STATUS, ""),
        ("LC-01", "Create basic subscriber", "create", "5550100101", "plan=basic\nfeatures=voicemail", 201,
         "status=active\nplan=basic\nfeatures=voicemail"),
        ("LC-02", "Add a feature", "add_feature", "5550100101", "feature=mms", 200,
         "features=mms\nfeatures=voicemail"),
        ("LC-03", "Suspend service", "suspend", "5550100101", "", 200, "status=suspended"),
        ("LC-04", "Usage is blocked while suspended", "usage", "5550100101", "type=sms\namount=1", 409, ""),
        ("LC-05", "Restore service", "restore", "5550100101", "", 200, "status=active"),
        ("LC-06", "Remove a feature", "remove_feature", "5550100101", "feature=voicemail", 200,
         "!features=voicemail\nfeatures=mms"),
        ("LC-07", "Close the account", "close", "5550100101", "", 200, "status=closed"),
        ("LC-08", "Closed account reads back as closed", "get", "5550100101", "", 200, "status=closed"),
    ],
    "Usage & Notifications": [
        ("UN-00A", "Setup: close 5550100102 if left over from an earlier round", "close", "5550100102", "",
         CLEANUP_STATUS, ""),
        ("UN-00B", "Setup: close 5550100103 if left over from an earlier round", "close", "5550100103", "",
         CLEANUP_STATUS, ""),
        ("UN-01", "Create basic subscriber (1000 MB data)", "create", "5550100102", "plan=basic", 201,
         "remaining.data_mb=1000"),
        ("UN-02", "50% usage: no notification yet", "usage", "5550100102", "type=data_mb\namount=500", 200,
         "remaining.data_mb=500\n!notifications=*"),
        ("UN-03", "Cross 80%: exactly one warning", "usage", "5550100102", "type=data_mb\namount=350", 200,
         "once:notifications=*80%*\nremaining.data_mb=150"),
        ("UN-04", "Cross 100%: exhausted, no repeat of 80% warning", "usage", "5550100102",
         "type=data_mb\namount=200", 200,
         "once:notifications=*exhausted*\n!notifications=*80%*\nremaining.data_mb=0"),
        ("UN-05", "Create unlimited subscriber", "create", "5550100103", "plan=unlimited", 201,
         "remaining.data_mb=unlimited"),
        ("UN-06", "Unlimited plan never notifies", "usage", "5550100103", "type=data_mb\namount=99999", 200,
         "!notifications=*\nremaining.data_mb=unlimited"),
        # Cleanup: close the subscribers this sheet created so the next round starts clean.
        ("UN-07", "Clean up: close the account created in UN-01", "close", "5550100102", "", CLEANUP_STATUS, ""),
        ("UN-08", "Clean up: close the account created in UN-05", "close", "5550100103", "", CLEANUP_STATUS, ""),
    ],
    "Negative": [
        ("NG-00", "Setup: close 5550100105 if left over from an earlier round", "close", "5550100105", "",
         CLEANUP_STATUS, ""),
        ("NG-01", "Unknown subscriber", "get", "5550100199", "", 404, ""),
        ("NG-02", "Invalid phone number", "create", "12345", "plan=basic", 422, ""),
        ("NG-03", "Unknown plan", "create", "5550100104", "plan=gold", 422, ""),
        ("NG-04", "Create for duplicate check", "create", "5550100105", "plan=plus", 201, "status=active"),
        ("NG-05", "Duplicate phone number rejected", "create", "5550100105", "plan=plus", 409, ""),
        ("NG-06", "Invalid usage type", "usage", "5550100105", "type=fax\namount=1", 422, ""),
        ("NG-07", "Zero usage amount rejected", "usage", "5550100105", "type=sms\namount=0", 422, ""),
        ("NG-08", "Clean up: close the account created in NG-04", "close", "5550100105", "", CLEANUP_STATUS, ""),
    ],
}

HEADER_FILLS = {
    **{h: PatternFill("solid", start_color="DDEBF7") for h in INPUT_HEADERS},     # inputs: blue
    **{h: PatternFill("solid", start_color="FFF2CC") for h in EXPECTED_HEADERS},  # expected: yellow
}
WIDTHS = {"ID": 8, "Description": 38, "Action": 15, "Phone": 13, "Params": 20, "Expected Status": 10,
          "Expected Response": 34, "Actual Status": 10, "Actual Response": 40, "Result": 9, "Details": 50}


def build_sample(path: Path = OUTPUT) -> Path:
    """Write the sample workbook to ``path`` and return the path."""
    wb = Workbook()
    wb.remove(wb.active)
    wrap = Alignment(wrap_text=True, vertical="top")

    for name, rows in SHEETS.items():
        ws = wb.create_sheet(name)
        ws.append(ALL_HEADERS)
        for col, header in enumerate(ALL_HEADERS, start=1):
            cell = ws.cell(row=1, column=col)
            cell.font = Font(bold=True)
            cell.fill = HEADER_FILLS.get(header, PatternFill("solid", start_color="E2EFDA"))  # outputs: green
            ws.column_dimensions[cell.column_letter].width = WIDTHS[header]
        for row in rows:
            ws.append(list(row))
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = wrap
            row[3].number_format = "@"  # Phone as text, so leading zeros survive
            row[5].number_format = "@"  # Expected Status as text, so "200,404,409" stays a list
        ws.freeze_panes = "C2"  # keep headers and IDs visible while scrolling

    # Config sheet read by the Excel macro (vba/RunSelected.bas).
    cfg = wb.create_sheet("Config")
    cfg.append(["Python executable", r"C:\path\to\excel-regression-runner\.venv\Scripts\python.exe"])
    cfg.append(["Project folder", r"C:\path\to\excel-regression-runner"])
    cfg.append([])
    cfg.append(["Edit B1 and B2 to match your machine. Credentials go in the project's .env file, never here."])
    cfg.column_dimensions["A"].width = 20
    cfg.column_dimensions["B"].width = 70

    wb.save(path)
    return path


if __name__ == "__main__":
    print(f"wrote {build_sample()}")
