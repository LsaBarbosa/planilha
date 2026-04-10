Option Explicit

Public Sub InitializeWorkbook()
    EnsureConfigSheet
    EnsureStudiesSheet
    EnsureIndexSheet
    SetConfigValue "app_name", APP_NAME
    SetConfigValue "app_version", APP_VERSION
    SetConfigValue "initialized_at", Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Public Sub EnsureConfigSheet()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_CONFIG
    End If

    ws.Cells(1, COL_KEY).Value = "Key"
    ws.Cells(1, COL_VALUE).Value = "Value"
End Sub

Public Function GetConfigValue(ByVal key As String, Optional ByVal defaultValue As String = "") As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long

    EnsureConfigSheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)

    lastRow = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    For rowIndex = 2 To lastRow
        If StrComp(CStr(ws.Cells(rowIndex, COL_KEY).Value), key, vbTextCompare) = 0 Then
            GetConfigValue = CStr(ws.Cells(rowIndex, COL_VALUE).Value)
            Exit Function
        End If
    Next rowIndex

    GetConfigValue = defaultValue
End Function

Public Sub SetConfigValue(ByVal key As String, ByVal value As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long

    EnsureConfigSheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)

    lastRow = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    For rowIndex = 2 To lastRow
        If StrComp(CStr(ws.Cells(rowIndex, COL_KEY).Value), key, vbTextCompare) = 0 Then
            ws.Cells(rowIndex, COL_VALUE).Value = value
            Exit Sub
        End If
    Next rowIndex

    ws.Cells(lastRow + 1, COL_KEY).Value = key
    ws.Cells(lastRow + 1, COL_VALUE).Value = value
End Sub
