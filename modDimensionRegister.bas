Option Explicit

Public Sub EnsureStudiesSheet()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_STUDIES)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_STUDIES
    End If

    If ws.Cells(1, STUDY_COL_ID).Value = vbNullString Then
        ws.Cells(1, STUDY_COL_ID).Value = "StudyID"
        ws.Cells(1, STUDY_COL_NAME).Value = "StudyName"
        ws.Cells(1, STUDY_COL_VALUE).Value = "Value"
        ws.Cells(1, STUDY_COL_STATUS).Value = "Status"
        ws.Cells(1, STUDY_COL_UPDATED_AT).Value = "UpdatedAt"
        ws.Cells(1, STUDY_COL_SCORE).Value = "Score"
        ws.Rows(1).Font.Bold = True
    End If
End Sub

Public Function GetLastStudyRow() As Long
    Dim ws As Worksheet
    EnsureStudiesSheet
    Set ws = ThisWorkbook.Worksheets(SHEET_STUDIES)
    GetLastStudyRow = ws.Cells(ws.Rows.Count, STUDY_COL_ID).End(xlUp).Row
    If GetLastStudyRow < 1 Then GetLastStudyRow = 1
End Function

Public Function FindStudyRowById(ByVal studyId As String) As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long

    EnsureStudiesSheet
    Set ws = ThisWorkbook.Worksheets(SHEET_STUDIES)
    lastRow = GetLastStudyRow()

    For rowIndex = 2 To lastRow
        If StrComp(CStr(ws.Cells(rowIndex, STUDY_COL_ID).Value), studyId, vbTextCompare) = 0 Then
            FindStudyRowById = rowIndex
            Exit Function
        End If
    Next rowIndex

    FindStudyRowById = 0
End Function

Public Function AddStudyRow(ByVal studyId As String, ByVal studyName As String, ByVal value As Double, ByVal status As String) As Long
    Dim ws As Worksheet
    Dim targetRow As Long

    EnsureStudiesSheet
    Set ws = ThisWorkbook.Worksheets(SHEET_STUDIES)

    If FindStudyRowById(studyId) <> 0 Then
        Err.Raise vbObjectError + 1004, "AddStudyRow", "Study ID já existe: " & studyId
    End If

    targetRow = GetLastStudyRow() + 1
    ws.Cells(targetRow, STUDY_COL_ID).Value = studyId
    ws.Cells(targetRow, STUDY_COL_NAME).Value = studyName
    ws.Cells(targetRow, STUDY_COL_VALUE).Value = value
    ws.Cells(targetRow, STUDY_COL_STATUS).Value = status
    ws.Cells(targetRow, STUDY_COL_UPDATED_AT).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(targetRow, STUDY_COL_SCORE).Value = vbNullString

    AddStudyRow = targetRow
End Function
