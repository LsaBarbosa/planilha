Attribute VB_Name = "modConfig"
Option Explicit

Public Const APP_VERSION As String = "F01-WorkbookBaseUI-v1"
Public Const SHEET_INSTRUCTIONS As String = "Instructions"
Public Const SHEET_STUDY_INDEX As String = "Study Index"
Public Const SHEET_DIM_REGISTER As String = "Dimension Register"
Public Const SHEET_TEMPLATE As String = "TEMPLATE"

Public Const HEADER_FIRST_ROW As Long = 1
Public Const HEADER_LABEL_ROW As Long = 1
Public Const HEADER_VALUE_ROW As Long = 2
Public Const TABLE_HEADER_ROW_STANDARD As Long = 5
Public Const TABLE_FIRST_DATA_ROW_STANDARD As Long = 6
Public Const STANDARD_INITIAL_ROWS As Long = 30

Public Const STUDY_HEADER_LABEL_ROW As Long = 4
Public Const STUDY_HEADER_VALUE_ROW As Long = 5
Public Const TABLE_HEADER_ROW_STUDY As Long = 8
Public Const TABLE_FIRST_DATA_ROW_STUDY As Long = 9
Public Const STUDY_INITIAL_ROWS As Long = 10

Public Sub InstallFeatureF01()
    Dim wb As Workbook

    On Error GoTo CleanFail
    Set wb = ThisWorkbook

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    modWorkbookBaseUI.BuildWorkbookBaseUI wb
    ApplyApplicationWindowDefaults wb
    ActivateEntrySheet wb

CleanExit:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "F01 installation failed: " & Err.Description, vbCritical + vbOKOnly, "Workbook Base & UI"
    Resume CleanExit
End Sub

Public Sub OnWorkbookOpen()
    On Error Resume Next
    ApplyApplicationWindowDefaults ThisWorkbook
    EnsureReservedSheetOrder ThisWorkbook
    ActivateEntrySheet ThisWorkbook
End Sub

Public Sub ApplyApplicationWindowDefaults(ByVal wb As Workbook)
    Dim ws As Worksheet

    On Error Resume Next
    Application.DisplayFormulaBar = True

    For Each ws In wb.Worksheets
        ApplySheetWindowDefaults ws
    Next ws
End Sub

Public Sub ApplySheetWindowDefaults(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = True
End Sub

Public Sub ActivateEntrySheet(ByVal wb As Workbook)
    On Error Resume Next
    wb.Worksheets(SHEET_INSTRUCTIONS).Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub

Public Sub EnsureReservedSheetOrder(ByVal wb As Workbook)
    Dim instructionsWs As Worksheet
    Dim studyIndexWs As Worksheet
    Dim dimRegisterWs As Worksheet
    Dim templateWs As Worksheet

    Set instructionsWs = EnsureSheet(wb, SHEET_INSTRUCTIONS)
    Set studyIndexWs = EnsureSheet(wb, SHEET_STUDY_INDEX)
    Set dimRegisterWs = EnsureSheet(wb, SHEET_DIM_REGISTER)
    Set templateWs = EnsureSheet(wb, SHEET_TEMPLATE)

    instructionsWs.Move Before:=wb.Worksheets(1)
    studyIndexWs.Move After:=instructionsWs
    dimRegisterWs.Move After:=studyIndexWs
    templateWs.Move After:=dimRegisterWs
End Sub

Public Function EnsureSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureSheet Is Nothing Then
        Set EnsureSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function
