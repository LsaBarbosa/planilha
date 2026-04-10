Option Explicit

Public Sub EnsureIndexSheet()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_INDEX)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_INDEX
    End If

    ws.Cells(1, 1).Value = "Metric"
    ws.Cells(1, 2).Value = "Value"
    ws.Rows(1).Font.Bold = True
End Sub

Public Sub SyncIndexSheet()
    Dim ws As Worksheet
    Dim totalStudies As Long
    Dim activeStudies As Long
    Dim lastRow As Long
    Dim rowIndex As Long

    EnsureStudiesSheet
    EnsureIndexSheet

    Set ws = ThisWorkbook.Worksheets(SHEET_INDEX)

    lastRow = GetLastStudyRow()
    totalStudies = Application.Max(0, lastRow - 1)

    For rowIndex = 2 To lastRow
        If UCase$(Trim$(CStr(ThisWorkbook.Worksheets(SHEET_STUDIES).Cells(rowIndex, STUDY_COL_STATUS).Value))) = STATUS_ACTIVE Then
            activeStudies = activeStudies + 1
        End If
    Next rowIndex

    ws.Cells(2, 1).Value = "TotalStudies"
    ws.Cells(2, 2).Value = totalStudies
    ws.Cells(3, 1).Value = "ActiveStudies"
    ws.Cells(3, 2).Value = activeStudies
    ws.Cells(4, 1).Value = "TotalScore"
    ws.Cells(4, 2).Value = CalculateTotalScore()
    ws.Cells(5, 1).Value = "LastSync"
    ws.Cells(5, 2).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub
