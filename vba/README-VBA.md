# Excel front-end macro

Run tests straight from the workbook, with the results filled in beside each row:

- **Double-click a test's ID cell** to run that one test.
- **Select any rows and click "Run selected"** to run several at once.

The macro saves the workbook, runs the Python runner for just those rows, and copies the
results back. All test logic stays in Python.

`samples/regression_tests.xlsm` is ready to use: just set the two paths on its Config sheet
(step 4 below). To add the macro to your own workbook, follow all the steps.

## Setup

1. Complete the Quick start in the main README (virtual environment and `.env`).
2. Open your test workbook in Excel and save it as **Excel Macro-Enabled Workbook (.xlsm)**.
3. Add the code (press **Alt+F11** to open the VBA editor):
   - **File > Import File...** and choose `vba/RunSelected.bas`.
   - Double-click **ThisWorkbook** in the Project panel, then paste the whole of
     `vba/ThisWorkbook.cls` into the code window. (Don't import this one: Excel would create
     a separate class instead of updating ThisWorkbook.)
4. Add a sheet named `Config` (the sample workbook already has one):
   - `B1`: full path to the venv's Python, e.g. `C:\projects\excel-regression-runner\.venv\Scripts\python.exe`
   - `B2`: full path to the project folder, e.g. `C:\projects\excel-regression-runner`
5. Press **Alt+F8**, run **AddRunButtons** once to put a "Run selected" button on every test
   sheet, then save.

Keep credentials in the project's `.env` file. Never put them in the workbook.

## Use

1. Start the API you're testing (or the mock: `python -m uvicorn mock_server.app:app --port 8000`).
2. Either double-click a test's **ID** cell, or select rows (any cells in them; hidden rows are
   skipped) and click **Run selected**.

Actual Status, Actual Response, Result and Details are filled in, with the Result cell shaded
green or red. A pass shows briefly in the status bar; failures pop up a message.

Double-clicking anywhere other than the ID column edits the cell as usual.

## Troubleshooting

- **Alt+F8 shows no macros**: the module isn't in this workbook. In the VBA editor (Alt+F11,
  then Ctrl+R) check that **RunSelected** is under this workbook's Modules, and that the file
  is saved as `.xlsm`. Saving as `.xlsx` silently removes macros.
- **Double-clicking an ID cell just edits it**: the ThisWorkbook code is missing (step 3).
- **"Check the paths on the Config sheet"**: B1 or B2 doesn't exist.
- **"The runner could not complete"**: open the log file named in the message (in `%TEMP%`).
  Common causes are a missing `.env`, the API not running, or a sheet missing a required column.
- **Macros disabled**: click **Enable Content** on the yellow bar, or keep the file in a Trusted Location.

The macro calls `python.exe -m runner.cli` rather than any `.exe` launcher in the venv,
so it works on machines where application control blocks those launchers.
