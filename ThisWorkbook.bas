Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    InitializeWorkbook
    On Error GoTo 0
End Sub
