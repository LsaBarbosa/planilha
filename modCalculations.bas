Option Explicit

Public Sub RecalculateStudyScores()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim rawValue As Variant
    Dim status As String

    EnsureStudiesSheet
    Set ws = ThisWorkbook.Worksheets(SHEET_STUDIES)
    lastRow = GetLastStudyRow()

    For rowIndex = 2 To lastRow
        status = UCase$(Trim$(CStr(ws.Cells(rowIndex, STUDY_COL_STATUS).Value)))
        rawValue = ws.Cells(rowIndex, STUDY_COL_VALUE).Value

        If status = STATUS_ACTIVE And IsNumeric(rawValue) Then
            ws.Cells(rowIndex, STUDY_COL_SCORE).Value = CDbl(rawValue) * 2
        Else
            ws.Cells(rowIndex, STUDY_COL_SCORE).Value = vbNullString
        End If
    Next rowIndex
End Sub

Public Function CalculateTotalScore() As Double
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim scoreValue As Variant
    Dim total As Double

    EnsureStudiesSheet
    Set ws = ThisWorkbook.Worksheets(SHEET_STUDIES)
    lastRow = GetLastStudyRow()

    total = 0
    For rowIndex = 2 To lastRow
        scoreValue = ws.Cells(rowIndex, STUDY_COL_SCORE).Value
        If IsNumeric(scoreValue) Then
            total = total + CDbl(scoreValue)
        End If
    Next rowIndex

    CalculateTotalScore = total
End Function
