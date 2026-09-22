Attribute VB_Name = "mod_report_tools"
Option Explicit
' Excel Zero to Hero - Module 7 sample toolkit: report formatting.
' Import: Alt+F11 > File > Import File. Run: Alt+F8. Save as .xlsm to keep macros.

Sub FormatReport()
    ' Turns the active sheet into a readable report in one click.
    Dim ws As Worksheet
    Set ws = ActiveSheet
    With ws.Range("A1").CurrentRegion
        ' Header row
        With .Rows(1)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)
        End With
        ' Number columns get thousands separators
        .Columns(2).NumberFormat = "#,##0"
        ' Autofit everything
        .Columns.AutoFit
    End With
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    MsgBox "Report formatted. Freeze panes on.", vbInformation
End Sub

Sub AddSummaryRow()
    ' Appends a bold Total row under the current region's first numeric column.
    Dim ws As Worksheet, lastRow As Long
    Set ws = ActiveSheet
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count + 1
    ws.Cells(lastRow, 1).Value = "TOTAL"
    ws.Cells(lastRow, 2).Formula = "=SUM(" & ws.Range("B1").CurrentRegion.Columns(2).Address(False, False) & ")"
    ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, 2)).Font.Bold = True
End Sub
