Option Explicit

Public Sub UpdateStudyValue(ByVal studyId As String, ByVal value As Double)
    Dim ws As Worksheet
    Dim rowIndex As Long

    rowIndex = FindStudyRowById(studyId)
    If rowIndex = 0 Then
        Err.Raise vbObjectError + 1001, "UpdateStudyValue", "Study ID não encontrado: " & studyId
    End If

    Set ws = ThisWorkbook.Worksheets(SHEET_STUDIES)
    ws.Cells(rowIndex, STUDY_COL_VALUE).Value = value
    ws.Cells(rowIndex, STUDY_COL_UPDATED_AT).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Public Sub SetStudyStatus(ByVal studyId As String, ByVal status As String)
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim normalizedStatus As String

    normalizedStatus = UCase$(Trim$(status))
    If normalizedStatus <> STATUS_ACTIVE And normalizedStatus <> STATUS_INACTIVE Then
        Err.Raise vbObjectError + 1002, "SetStudyStatus", "Status inválido: " & status
    End If

    rowIndex = FindStudyRowById(studyId)
    If rowIndex = 0 Then
        Err.Raise vbObjectError + 1003, "SetStudyStatus", "Study ID não encontrado: " & studyId
    End If

    Set ws = ThisWorkbook.Worksheets(SHEET_STUDIES)
    ws.Cells(rowIndex, STUDY_COL_STATUS).Value = normalizedStatus
    ws.Cells(rowIndex, STUDY_COL_UPDATED_AT).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub
