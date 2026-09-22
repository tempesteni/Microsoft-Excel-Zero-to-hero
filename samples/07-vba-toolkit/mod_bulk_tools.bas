Attribute VB_Name = "mod_bulk_tools"
Option Explicit
' Excel Zero to Hero - Module 7 sample toolkit: bulk operations.
' Import: Alt+F11 > File > Import File. Run: Alt+F8. Save as .xlsm to keep macros.

Sub RemoveEmptyRows()
    ' Deletes fully blank rows in the current region bottom-up (bottom-up keeps indexes valid).
    Dim ws As Worksheet, r As Long
    Set ws = ActiveSheet
    For r = ws.Range("A1").CurrentRegion.Rows.Count To 2 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

Sub ExportEachSheetToCSV()
    ' Saves every worksheet as its own CSV next to this workbook.
    Dim ws As Worksheet, folder As String
    folder = ThisWorkbook.Path & "\"
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        ws.Copy
        ActiveWorkbook.SaveAs folder & ws.Name & ".csv", xlCSV
        ActiveWorkbook.Close False
    Next ws
    Application.DisplayAlerts = True
    MsgBox "Exported " & ThisWorkbook.Worksheets.Count & " sheets to " & folder, vbInformation
End Sub

Sub RenameSheetsFromColumnA()
    ' Renames Sheet1..N using the values in A1 of each sheet (trimmed, max 31 chars).
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Len(Trim$(ws.Range("A1").Value)) > 0 Then
            ws.Name = Left$(Trim$(ws.Range("A1").Value), 31)
        End If
    Next ws
End Sub
