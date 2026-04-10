Option Explicit

Public Function RunReleaseChecks() As String
    Dim validationMessage As String
    Dim isValid As Boolean

    isValid = ValidateWorkbook(validationMessage)

    If isValid Then
        RunReleaseChecks = "OK | " & validationMessage
    Else
        RunReleaseChecks = "FAIL | " & validationMessage
    End If
End Function
