' ThisWorkbook code: double-click a test's ID cell to run that test.
'
' Do NOT use File > Import for this file (Excel would create a separate
' class instead of updating ThisWorkbook). Instead:
'   1. Alt+F11, then double-click "ThisWorkbook" in the Project panel.
'   2. Paste this whole file into the code window that opens.
'
' Only double-clicks in the ID column of a test sheet run a test, so
' double-click editing works normally everywhere else.

Option Explicit

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Not RunSelected.IsTestSheet(Sh) Then Exit Sub
    If Target.Row < 2 Or Target.Cells.Count > 1 Then Exit Sub
    If Target.Column <> RunSelected.HeaderColumn(Sh, "ID") Then Exit Sub

    Cancel = True  ' don't open the ID cell for editing
    RunSelected.RunRow Sh, Target.Row
End Sub
