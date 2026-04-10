Option Explicit

Public Function GenerateStudyId() As String
    Static seeded As Boolean
    Dim suffix As Long

    If Not seeded Then
        Randomize Timer
        seeded = True
    End If

    suffix = CLng((Rnd * 9000) + 1000)
    GenerateStudyId = "ST-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(suffix, "0000")
End Function

Public Function CreateStudy(ByVal studyName As String, ByVal value As Double) As String
    Dim id As String

    If Trim$(studyName) = vbNullString Then
        Err.Raise vbObjectError + 1000, "CreateStudy", "Nome do estudo é obrigatório."
    End If

    id = GenerateStudyId()
    AddStudyRow id, studyName, value, STATUS_ACTIVE
    CreateStudy = id
End Function
