# excel-regression-runner

![CI](https://github.com/Brandonicus-Rex/excel-regression-runner/actions/workflows/ci.yml/badge.svg)

Spreadsheet-driven API regression testing. Testers write cases as rows in Excel,
and a Python runner executes them against a REST API, compares the responses with
the expected results, and writes PASS/FAIL back into the workbook with a summary sheet.

An optional Excel macro lets you select rows, click a button and see results in place,
while all the logic stays in Python where it can be tested.

```
 Excel workbook ──► runner.cli ──► REST API
  (test rows)          │
                       ▼
 results workbook ◄── compare expected vs actual
  (PASS/FAIL, details, Summary sheet)
```

A mock subscriber/usage API is included, so everything runs locally with no external systems.

## Features

- One test case per row; columns are found by header name, so order doesn't matter
- Wildcard expected results, plus "must not appear" and "exactly once" rules
- Word-level diff in the Details column when a value doesn't match
- Run all sheets, one sheet, or specific rows (`--rows 2-5,8`)
- Setup and cleanup rows, so test rounds can be repeated back to back
- Results go to a new workbook; the input is never modified
- Summary sheet with per-sheet counts and a list of failures
- Exit codes for CI: `0` all passed, `1` failures, `2` setup error
- Credentials from environment variables / `.env`, never from the workbook

## Workbook format

Each test sheet has a header row and one case per row. Sheets named `Summary` or `Config` are skipped.

| Column | Purpose |
|---|---|
| ID, Description | For people; shown in the summary |
| Action | `create`, `get`, `suspend`, `restore`, `close`, `add_feature`, `remove_feature`, `usage` |
| Phone | Subscriber phone number |
| Params | `key=value` pairs, one per line (or separated by `;`) |
| Expected Status | HTTP status code, e.g. `201`, or several that all count as a pass, e.g. `200,404,409` (blank = don't check) |
| Expected Response | Rules, one per line (see below; blank = don't check) |
| Actual Status, Actual Response, Result, Details | Written by the runner |

### Expected response rules

The JSON response is flattened into `key=value` lines (`remaining.data_mb=150`,
one `features=...` line per list item). Each rule line is checked against those lines,
case-insensitively, with `*` and `?` wildcards:

| Rule | Passes when |
|---|---|
| `status=active` | at least one line matches |
| `!notifications=*` | no line matches |
| `once:notifications=*80%*` | exactly one line matches |

Every rule is checked and every failure is reported, not just the first.

### Cleanup rows

Each sample sheet closes the subscribers it uses at both ends:

- **Setup rows at the top** close any account left open by an earlier round that
  stopped partway, so a round never depends on the previous one finishing.
- **Cleanup rows at the bottom** close what the sheet created, so the next round starts clean.

The mock API lets a closed number be created again. Cleanup rows use Expected Status
`200,404,409`: they pass whether they closed the account, it was already closed, or it
never existed. The test suite covers repeated rounds and recovery from an interrupted one.

## Quick start

Requires Python 3.10 to 3.13.

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate    macOS/Linux: source .venv/bin/activate
python -m pip install -r requirements.txt
cp .env.example .env          # Windows: copy .env.example .env

# Terminal 1: start the mock API
python -m uvicorn mock_server.app:app --port 8000

# Terminal 2: run the sample workbook
python -m runner.cli samples/regression_tests.xlsx
```

Output:

```
28/28 passed -> samples/regression_tests_results.xlsx
```

More examples:

```bash
python -m runner.cli samples/regression_tests.xlsx --sheet Lifecycle
python -m runner.cli samples/regression_tests.xlsx --sheet Lifecycle --rows 2-4
python -m runner.cli samples/regression_tests.xlsx --output reports/run.xlsx -v
```

To rebuild the sample workbook: `python -m samples.make_sample`.

## Run from Excel (optional)

`samples/regression_tests.xlsm` has a small macro built in: double-click a test's ID cell,
or select rows and click **Run selected**, and the results appear in the same sheet.
Set the two paths on its Config sheet first. To add the macro to your own workbook, see
[vba/README-VBA.md](vba/README-VBA.md).

## Tests

```bash
python -m pytest
```

Unit tests cover the comparison rules and workbook handling. End-to-end tests start
the mock API on a free port and run the sample workbook through the CLI.

## Project layout

```
runner/
  cli.py          command-line entry point
  workbook.py     read test rows, write results and the summary sheet
  engine.py       run one case: call the API, then compare
  actions.py      maps the Action column to API requests
  validators.py   flatten JSON and apply the expected-response rules
  client.py       HTTP client with retries; settings from the environment
mock_server/app.py  in-memory FastAPI subscriber/usage API
samples/          sample workbooks (.xlsx, and .xlsm with the macro) and the generator
vba/              optional Excel front-end macro
tests/
```

## Adding an action

Add a function to `runner/actions.py` that takes `(client, phone, params)` and returns
the response, then register it in `ACTIONS`. It's immediately usable from the Action column.

## License

MIT
