Option Explicit

Public Function ValidateWorkbook(ByRef message As String) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim status As String

    On Error GoTo ValidationError

    EnsureConfigSheet
    EnsureStudiesSheet
    EnsureIndexSheet

    Set ws = ThisWorkbook.Worksheets(SHEET_STUDIES)
    If ws.Cells(1, STUDY_COL_ID).Value <> "StudyID" Then
        message = "Cabeçalho StudyID inválido na planilha Studies."
        ValidateWorkbook = False
        Exit Function
    End If
    If ws.Cells(1, STUDY_COL_NAME).Value <> "StudyName" Then
        message = "Cabeçalho StudyName inválido na planilha Studies."
        ValidateWorkbook = False
        Exit Function
    End If
    If ws.Cells(1, STUDY_COL_STATUS).Value <> "Status" Then
        message = "Cabeçalho Status inválido na planilha Studies."
        ValidateWorkbook = False
        Exit Function
    End If

    lastRow = GetLastStudyRow()
    For rowIndex = 2 To lastRow
        status = UCase$(Trim$(CStr(ws.Cells(rowIndex, STUDY_COL_STATUS).Value)))
        If status <> vbNullString And status <> STATUS_ACTIVE And status <> STATUS_INACTIVE Then
            message = "Status inválido encontrado na linha " & rowIndex & "."
            ValidateWorkbook = False
            Exit Function
        End If
    Next rowIndex

    message = "Estrutura validada com sucesso."
    ValidateWorkbook = True
    Exit Function

ValidationError:
    message = "Erro de validação: " & Err.Description
    ValidateWorkbook = False
End Function
