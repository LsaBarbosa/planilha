Option Explicit

Public Sub Main()
    On Error GoTo Handler

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    InitializeWorkbook

    If GetLastStudyRow() = 1 Then
        CreateStudy "Estudo inicial", 10
        CreateStudy "Estudo secundário", 20
    End If

    RecalculateStudyScores
    SyncIndexSheet

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Sistema inicializado com sucesso.", vbInformation, APP_NAME
    Exit Sub

Handler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Erro: " & Err.Description, vbCritical, APP_NAME
End Sub

Public Sub RunQualityChecks()
    MsgBox RunReleaseChecks(), vbInformation, APP_NAME
End Sub

Public Sub RecalculateAndSync()
    RecalculateStudyScores
    SyncIndexSheet
    MsgBox "Recalculo e sincronização concluídos.", vbInformation, APP_NAME
End Sub
