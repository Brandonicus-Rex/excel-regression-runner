Attribute VB_Name = "RunSelected"
' RunSelected: thin Excel front-end for the Python regression runner.
'
' Two ways to run tests from a test sheet:
'   - Double-click a test's ID cell to run that one row
'     (needs the small handler in vba\ThisWorkbook.cls).
'   - Select any rows and click the "Run selected" button
'     (run AddRunButtons once to add a button to every test sheet).
'
' Either way the macro:
'   1. saves the workbook so Python reads the latest inputs,
'   2. runs `python -m runner.cli` for just those rows,
'   3. copies the result columns back into the sheet.
'
' All the real work (HTTP calls, comparisons, reporting) lives in Python
' where it can be unit tested. Paths come from the Config sheet:
'   Config!B1 = python.exe inside the project's virtual environment
'   Config!B2 = the project folder (where runner\ lives)
' Credentials stay in the project's .env file, never in the workbook.

Option Explicit

Private Const RESULT_COLUMNS As String = "Actual Status|Actual Response|Result|Details"

' ---- Entry points ----------------------------------------------------------

' Button: run every visible selected row on the active sheet.
Public Sub RunSelectedTests()
    Dim ws As Worksheet, rows As Object, cell As Range
    Set ws = ActiveSheet
    If Not IsTestSheet(ws) Then MsgBox "Select rows on a test sheet first.", vbExclamation: Exit Sub

    Set rows = CreateObject("Scripting.Dictionary")
    If TypeName(Selection) = "Range" Then
        If Not Intersect(Selection.EntireRow, ws.UsedRange) Is Nothing Then
            For Each cell In Intersect(Selection.EntireRow, ws.UsedRange).Columns(1).Cells
                If cell.Row >= 2 And Not cell.EntireRow.Hidden Then rows(cell.Row) = True
            Next cell
        End If
    End If
    If rows.Count = 0 Then MsgBox "Select at least one test row (row 2 or below).", vbExclamation: Exit Sub

    RunRows ws, rows
End Sub

' Double-click: run a single row (called from ThisWorkbook).
Public Sub RunRow(ByVal ws As Worksheet, ByVal rowNumber As Long)
    Dim rows As Object
    If Trim$(CStr(ws.Cells(rowNumber, HeaderColumn(ws, "Action")).Value)) = "" Then
        MsgBox "Row " & rowNumber & " has no Action, so there's nothing to run.", vbExclamation
        Exit Sub
    End If
    Set rows = CreateObject("Scripting.Dictionary")
    rows(rowNumber) = True
    RunRows ws, rows
End Sub

' One-time setup: add a "Run selected" button to row 1 of every test sheet,
' just right of the last header. Safe to run again (old buttons are replaced).
Public Sub AddRunButtons()
    Dim ws As Worksheet, btn As Object, anchor As Range
    For Each ws In ThisWorkbook.Worksheets
        If IsTestSheet(ws) Then
            For Each btn In ws.Buttons
                If btn.Caption = "Run selected" Then btn.Delete
            Next btn
            If ws.Rows(1).RowHeight < 24 Then ws.Rows(1).RowHeight = 24
            Set anchor = ws.Cells(1, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1)
            Set btn = ws.Buttons.Add(anchor.Left + 4, anchor.Top + 2, 100, 20)
            btn.Caption = "Run selected"
            btn.OnAction = "RunSelectedTests"
            btn.Placement = xlMove  ' moves with its cell, keeps its size
        End If
    Next ws
End Sub

' ---- Shared helpers (also used by ThisWorkbook) -----------------------------

' A test sheet is any sheet with an "Action" header, except Config and Summary.
Public Function IsTestSheet(ByVal ws As Object) As Boolean
    If TypeName(ws) <> "Worksheet" Then Exit Function
    If LCase$(ws.Name) = "config" Or LCase$(ws.Name) = "summary" Then Exit Function
    IsTestSheet = HeaderColumn(ws, "Action") > 0
End Function

' Column number of a header in row 1, or 0 if not found.
Public Function HeaderColumn(ByVal ws As Worksheet, ByVal header As String) As Long
    Dim found As Range
    Set found = ws.Rows(1).Find(What:=header, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not found Is Nothing Then HeaderColumn = found.Column
End Function

' ---- Runner ------------------------------------------------------------------

' Run the given rows (a Dictionary keyed by row number) and copy back the results.
Private Sub RunRows(ws As Worksheet, rows As Object)
    Dim pythonExe As String, projectDir As String, rowSpec As String
    Dim outFile As String, logFile As String, cmd As String, stamp As String
    Dim exitCode As Long

    pythonExe = Trim$(ThisWorkbook.Worksheets("Config").Range("B1").Value)
    projectDir = Trim$(ThisWorkbook.Worksheets("Config").Range("B2").Value)
    If pythonExe = "" Or projectDir = "" Then GoTo BadConfig
    If Dir(pythonExe) = "" Or Dir(projectDir, vbDirectory) = "" Then GoTo BadConfig

    rowSpec = CompressRows(rows)
    ThisWorkbook.Save  ' Python reads the file on disk, not what's in memory

    stamp = Format$(Now, "yyyymmdd_hhnnss")
    outFile = Environ$("TEMP") & "\regression_results_" & stamp & ".xlsx"
    logFile = Environ$("TEMP") & "\regression_log_" & stamp & ".txt"

    ' cmd /s /c "..." keeps the inner quotes intact for paths with spaces.
    ' "python -m" is used because some machines block the venv's .exe launchers.
    cmd = "cmd /s /c """ & "cd /d " & Q(projectDir) & " && " & Q(pythonExe) & " -m runner.cli " & _
          Q(ThisWorkbook.FullName) & " --sheet " & Q(ws.Name) & " --rows " & rowSpec & _
          " --output " & Q(outFile) & " > " & Q(logFile) & " 2>&1" & """"

    Application.StatusBar = "Running " & ws.Name & " rows " & rowSpec & "..."
    exitCode = CreateObject("WScript.Shell").Run(cmd, 0, True)  ' 0 = hidden window, True = wait
    Application.StatusBar = False

    ' Exit codes from the runner: 0 = all passed, 1 = some failed, 2 = setup error.
    If exitCode > 1 Or Dir(outFile) = "" Then
        MsgBox "The runner could not complete (exit code " & exitCode & ")." & vbCrLf & _
               "See the log: " & logFile, vbCritical
        Exit Sub
    End If

    CopyResults ws, rows, outFile
    Kill outFile
    Kill logFile
    If exitCode = 0 Then
        Application.StatusBar = "Passed: " & ws.Name & " rows " & rowSpec
        Application.OnTime Now + TimeValue("00:00:05"), "ClearStatusBar"
    Else
        MsgBox "Some tests failed. See the Details column.", vbExclamation
    End If
    Exit Sub

BadConfig:
    MsgBox "Check the paths on the Config sheet (B1 = python.exe, B2 = project folder).", vbExclamation
End Sub

' Copy the output columns (found by header name) for the given rows from the results file.
Private Sub CopyResults(ws As Worksheet, rows As Object, outFile As String)
    Dim results As Workbook, src As Worksheet, header As Variant, r As Variant
    Dim srcCol As Long, dstCol As Long

    Application.ScreenUpdating = False
    Set results = Workbooks.Open(outFile, ReadOnly:=True)
    Set src = results.Worksheets(ws.Name)

    For Each header In Split(RESULT_COLUMNS, "|")
        srcCol = HeaderColumn(src, CStr(header))
        dstCol = HeaderColumn(ws, CStr(header))
        If srcCol > 0 And dstCol > 0 Then
            For Each r In rows.Keys
                ws.Cells(r, dstCol).Value = src.Cells(r, srcCol).Value
                If header = "Result" Then  ' carry over the green/red PASS/FAIL fill
                    ws.Cells(r, dstCol).Interior.Color = src.Cells(r, srcCol).Interior.Color
                End If
            Next r
        End If
    Next header

    results.Close SaveChanges:=False
    Application.ScreenUpdating = True
End Sub

' Turn row numbers into a compact spec such as "2-5,8".
Private Function CompressRows(rows As Object) As String
    Dim keys As Variant, i As Long, j As Long, tmp As Variant
    Dim startRow As Long, prevRow As Long, spec As String

    keys = rows.Keys
    For i = LBound(keys) To UBound(keys) - 1  ' small lists, so a simple sort is fine
        For j = i + 1 To UBound(keys)
            If keys(j) < keys(i) Then tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
        Next j
    Next i

    startRow = keys(0): prevRow = keys(0)
    For i = 1 To UBound(keys)
        If keys(i) = prevRow + 1 Then
            prevRow = keys(i)                  ' still in a consecutive run
        Else
            spec = AddRange(spec, startRow, prevRow)
            startRow = keys(i): prevRow = keys(i)
        End If
    Next i
    CompressRows = AddRange(spec, startRow, prevRow)
End Function

' Append "a" or "a-b" to a comma-separated spec.
Private Function AddRange(spec As String, firstRow As Long, lastRow As Long) As String
    Dim part As String
    If firstRow = lastRow Then part = CStr(firstRow) Else part = firstRow & "-" & lastRow
    If spec = "" Then AddRange = part Else AddRange = spec & "," & part
End Function

Public Sub ClearStatusBar()
    Application.StatusBar = False
End Sub

' Wrap a value in double quotes for the command line.
Private Function Q(text As String) As String
    Q = """" & text & """"
End Function
