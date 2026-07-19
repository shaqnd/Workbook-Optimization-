' ============================================================================
'  AUTO SHOW/HIDE BUILDING COLUMNS
' ----------------------------------------------------------------------------
'  This makes the building columns appear/disappear automatically based on the
'  "Number of Buildings (1-10)" cell (C8) on the Subject Property tab.
'
'  HOW TO INSTALL (one time, ~30 seconds):
'    1. Open the workbook in Excel and enable macros/content if prompted.
'    2. Press Alt+F11 to open the VBA editor.
'    3. In the Project pane (left), double-click the sheet named
'       "Subject Property"  (it may show as "Sheet1 (Subject Property)").
'    4. Paste EVERYTHING below the line marked >>>> into that sheet's code pane.
'    5. Press Alt+Q to return to Excel. Change C8 (Number of Buildings) once
'       to trigger it (e.g. type 1 then your real count).
'
'  It hides/shows the matching columns on BOTH the Subject Property tab and the
'  Cost Approach MB tab, so the two stay in sync. It is safe to keep the
'  conditional-formatting fallback on as well.
' ============================================================================
' >>>> PASTE FROM HERE <<<<
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("C8")) Is Nothing Then Exit Sub
    ApplyBuildingColumns
End Sub

Public Sub ApplyBuildingColumns()
    Dim n As Long, i As Long
    n = 1
    If IsNumeric(Me.Range("C8").Value) Then n = CLng(Me.Range("C8").Value)
    If n < 1 Then n = 1
    If n > 10 Then n = 10

    Application.ScreenUpdating = False
    ' Building i lives in column (2 + i): Building 1 = C, ... Building 10 = L.
    For i = 1 To 10
        Me.Columns(2 + i).Hidden = (i > n)
    Next i

    On Error Resume Next   ' keep going even if the MB sheet was renamed/removed
    For i = 1 To 10
        ThisWorkbook.Worksheets("Cost Approach MB").Columns(2 + i).Hidden = (i > n)
    Next i
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub
