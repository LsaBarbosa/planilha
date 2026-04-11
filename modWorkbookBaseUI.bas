Attribute VB_Name = "modWorkbookBaseUI"
Option Explicit

Public Sub BuildWorkbookBaseUI(ByVal wb As Workbook)
    Dim wsInstructions As Worksheet
    Dim wsStudyIndex As Worksheet
    Dim wsDimRegister As Worksheet
    Dim wsTemplate As Worksheet

    Set wsInstructions = modConfig.EnsureSheet(wb, modConfig.SHEET_INSTRUCTIONS)
    Set wsStudyIndex = modConfig.EnsureSheet(wb, modConfig.SHEET_STUDY_INDEX)
    Set wsDimRegister = modConfig.EnsureSheet(wb, modConfig.SHEET_DIM_REGISTER)
    Set wsTemplate = modConfig.EnsureSheet(wb, modConfig.SHEET_TEMPLATE)

    BuildInstructionsSheet wsInstructions
    BuildStudyIndexSheet wsStudyIndex
    BuildDimensionRegisterSheet wsDimRegister
    BuildTemplateSheet wsTemplate

    modConfig.EnsureReservedSheetOrder wb
End Sub

Public Sub BuildInstructionsSheet(ByVal ws As Worksheet)
    modSheetFormatting.ResetSheetCanvas ws
    modSheetFormatting.ApplyBaseVisualStandards ws
    modSheetFormatting.SetStandardColumnWidths ws

    ws.Name = modConfig.SHEET_INSTRUCTIONS

    ws.Range("A1:P1").RowHeight = 24
    modSheetFormatting.FormatTitle ws.Range("A1:J1"), "Tolerance Study Workbook - Instructions"
    modSheetFormatting.AddLogoPlaceholder ws, "L1", 180, 28

    modSheetFormatting.FormatSectionHeader ws.Range("A3:P3"), "Minimum operating flow"

    WriteInstructionLine ws, 5, "1", "Register dimensions in the Dimension Register sheet."
    WriteInstructionLine ws, 6, "2", "Create a new study from the TEMPLATE-based flow in later sprints."
    WriteInstructionLine ws, 7, "3", "Fill Dim. ID and Direction in each study row."
    WriteInstructionLine ws, 8, "4", "Review Totals, RSS Sigma, RSS Fail %, and Worst Case."
    WriteInstructionLine ws, 9, "5", "Read the consolidated result in Study Index."

    modSheetFormatting.FormatSectionHeader ws.Range("A11:P11"), "Engineering conventions"
    WriteInstructionText ws, 13, "Default study reading starts from the left side of the stack-up."
    WriteInstructionText ws, 14, "Switch Direction changes the interpretation of Max and Min according to the project convention."
    WriteInstructionText ws, 15, "Visible workbook labels remain in technical English."

    modSheetFormatting.FormatSectionHeader ws.Range("A17:P17"), "Sprint 1 scope delivered"
    WriteInstructionText ws, 19, "Base sheets created in the correct order: Instructions, Study Index, Dimension Register, TEMPLATE."
    WriteInstructionText ws, 20, "Document header, footer, visual theme, and sheet scaffolding are installed."
    WriteInstructionText ws, 21, "Workbook calculations, study factory, and governance routines are scheduled for later features."

    ws.Range("A24:P24").Merge
    ws.Range("A24").Value = "Run modConfig.InstallFeatureF01 to rebuild the base UI."
    ws.Range("A24").Font.Italic = True
    ws.Range("A24").HorizontalAlignment = xlLeft

    modSheetFormatting.ApplyOuterFrame ws.Range("A3:P24")
End Sub

Public Sub BuildStudyIndexSheet(ByVal ws As Worksheet)
    Dim headers As Variant
    Dim headerRange As Range
    Dim dataRange As Range
    Dim footerTopRow As Long

    headers = Array("GAP No.", "Study Description", "Target", "Nominal", "Max", "Min", "Worst Case Status", "RSS Fail %")

    modSheetFormatting.ResetSheetCanvas ws
    modSheetFormatting.ApplyBaseVisualStandards ws
    modSheetFormatting.SetStandardColumnWidths ws

    ws.Name = modConfig.SHEET_STUDY_INDEX
    modSheetFormatting.DrawDocumentHeader ws, "P"
    modSheetFormatting.AddLogoPlaceholder ws, "Q1", 135, 38

    ws.Range("A4:P4").Merge
    ws.Range("A4").Value = "Study Index"
    ws.Range("A4").Font.Bold = True
    ws.Range("A4").Font.Size = 12
    ws.Range("A4").HorizontalAlignment = xlLeft

    Set headerRange = ws.Range("A" & modConfig.TABLE_HEADER_ROW_STANDARD & ":H" & modConfig.TABLE_HEADER_ROW_STANDARD)
    modSheetFormatting.WriteHeadersHorizontal ws, modConfig.TABLE_HEADER_ROW_STANDARD, 1, headers
    modSheetFormatting.FormatHeaderRow headerRange

    Set dataRange = ws.Range("A" & modConfig.TABLE_FIRST_DATA_ROW_STANDARD & ":H" & modConfig.TABLE_FIRST_DATA_ROW_STANDARD + modConfig.STANDARD_INITIAL_ROWS - 1)
    modSheetFormatting.FormatDataInputRange dataRange.Columns(1)
    modSheetFormatting.FormatDataCalculatedRange dataRange.Columns(2).Resize(, 7)

    ws.Range("A" & modConfig.TABLE_FIRST_DATA_ROW_STANDARD & ":A" & modConfig.TABLE_FIRST_DATA_ROW_STANDARD + modConfig.STANDARD_INITIAL_ROWS - 1).HorizontalAlignment = xlCenter
    ws.Range("D" & modConfig.TABLE_FIRST_DATA_ROW_STANDARD & ":H" & modConfig.TABLE_FIRST_DATA_ROW_STANDARD + modConfig.STANDARD_INITIAL_ROWS - 1).NumberFormat = "0.000"

    footerTopRow = modConfig.TABLE_FIRST_DATA_ROW_STANDARD + modConfig.STANDARD_INITIAL_ROWS + 2
    modSheetFormatting.DrawDocumentFooter ws, footerTopRow, "P"
    modSheetFormatting.ApplyOuterFrame ws.Range("A1:P" & footerTopRow + 1)
End Sub

Public Sub BuildDimensionRegisterSheet(ByVal ws As Worksheet)
    Dim headers As Variant
    Dim headerRange As Range
    Dim dataRange As Range
    Dim footerTopRow As Long

    headers = Array("Item No.", "Part Name", "Part Number", "Drawing Page", "Revision", _
                    "Dimension Description", "Dim. ID", "Max", "Min", "Nominal", _
                    "TOL (+/-)", "Trig. Mean", "Trig. Ref.", "Units")

    modSheetFormatting.ResetSheetCanvas ws
    modSheetFormatting.ApplyBaseVisualStandards ws
    modSheetFormatting.SetStandardColumnWidths ws

    ws.Name = modConfig.SHEET_DIM_REGISTER
    modSheetFormatting.DrawDocumentHeader ws, "P"
    modSheetFormatting.AddLogoPlaceholder ws, "Q1", 135, 38

    ws.Range("A4:N4").Merge
    ws.Range("A4").Value = "Dimension Register"
    ws.Range("A4").Font.Bold = True
    ws.Range("A4").Font.Size = 12
    ws.Range("A4").HorizontalAlignment = xlLeft

    Set headerRange = ws.Range("A" & modConfig.TABLE_HEADER_ROW_STANDARD & ":N" & modConfig.TABLE_HEADER_ROW_STANDARD)
    modSheetFormatting.WriteHeadersHorizontal ws, modConfig.TABLE_HEADER_ROW_STANDARD, 1, headers
    modSheetFormatting.FormatHeaderRow headerRange

    Set dataRange = ws.Range("A" & modConfig.TABLE_FIRST_DATA_ROW_STANDARD & ":N" & modConfig.TABLE_FIRST_DATA_ROW_STANDARD + modConfig.STANDARD_INITIAL_ROWS - 1)
    modSheetFormatting.FormatDataInputRange dataRange

    ws.Range("H" & modConfig.TABLE_FIRST_DATA_ROW_STANDARD & ":K" & modConfig.TABLE_FIRST_DATA_ROW_STANDARD + modConfig.STANDARD_INITIAL_ROWS - 1).NumberFormat = "0.000"
    ws.Range("N" & modConfig.TABLE_FIRST_DATA_ROW_STANDARD & ":N" & modConfig.TABLE_FIRST_DATA_ROW_STANDARD + modConfig.STANDARD_INITIAL_ROWS - 1).Validation.Delete

    footerTopRow = modConfig.TABLE_FIRST_DATA_ROW_STANDARD + modConfig.STANDARD_INITIAL_ROWS + 2
    modSheetFormatting.DrawDocumentFooter ws, footerTopRow, "P"
    modSheetFormatting.ApplyOuterFrame ws.Range("A1:P" & footerTopRow + 1)
End Sub

Public Sub BuildTemplateSheet(ByVal ws As Worksheet)
    Dim headerRange As Range
    Dim dataRange As Range
    Dim footerTopRow As Long
    Dim studyHeaders As Variant
    Dim studyTableHeaders As Variant

    studyHeaders = Array("GAP No.", "STUDY DESCRIPTION", "ACCEPTANCE LIMIT", "REFERENCE TYPE", "SWITCH DIRECTION", "ENGINEERING JUSTIFICATION")
    studyTableHeaders = Array("Drawing No.", "Revision", "Dimension Description", "Dim. ID", "Direction", "Max", "Min", "Nominal", "TOL (+/-)", "Total Tol.", "Distribution", "σ")

    modSheetFormatting.ResetSheetCanvas ws
    modSheetFormatting.ApplyBaseVisualStandards ws
    modSheetFormatting.SetStandardColumnWidths ws

    ws.Name = modConfig.SHEET_TEMPLATE
    modSheetFormatting.DrawDocumentHeader ws, "P"
    modSheetFormatting.AddLogoPlaceholder ws, "Q1", 135, 38

    Set headerRange = ws.Range("A" & modConfig.STUDY_HEADER_LABEL_ROW & ":P" & modConfig.STUDY_HEADER_LABEL_ROW)
    modSheetFormatting.FormatHeaderRow headerRange

    ws.Range("A4:B4").Merge: ws.Range("A4").Value = studyHeaders(0)
    ws.Range("C4:F4").Merge: ws.Range("C4").Value = studyHeaders(1)
    ws.Range("G4:H4").Merge: ws.Range("G4").Value = studyHeaders(2)
    ws.Range("I4:J4").Merge: ws.Range("I4").Value = studyHeaders(3)
    ws.Range("K4:L4").Merge: ws.Range("K4").Value = studyHeaders(4)
    ws.Range("M4:P4").Merge: ws.Range("M4").Value = studyHeaders(5)

    modSheetFormatting.FormatMergedInputRange ws.Range("A5:B5"), "NEW GAP"
    modSheetFormatting.FormatMergedInputRange ws.Range("C5:F5"), "DESCRIPTION OF GAP"
    modSheetFormatting.FormatMergedInputRange ws.Range("G5:H5"), "0.000"
    modSheetFormatting.FormatMergedInputRange ws.Range("I5:J5"), "MAX"
    modSheetFormatting.FormatMergedInputRange ws.Range("K5:L5"), "Left"
    modSheetFormatting.FormatMergedInputRange ws.Range("M5:P5"), ""

    ws.Range("A7:L7").Merge
    ws.Range("A7").Value = "Study Template"
    ws.Range("A7").Font.Bold = True
    ws.Range("A7").Font.Size = 12
    ws.Range("A7").HorizontalAlignment = xlLeft

    Set headerRange = ws.Range("A" & modConfig.TABLE_HEADER_ROW_STUDY & ":L" & modConfig.TABLE_HEADER_ROW_STUDY)
    modSheetFormatting.WriteHeadersHorizontal ws, modConfig.TABLE_HEADER_ROW_STUDY, 1, studyTableHeaders
    modSheetFormatting.FormatHeaderRow headerRange

    Set dataRange = ws.Range("A" & modConfig.TABLE_FIRST_DATA_ROW_STUDY & ":L" & modConfig.TABLE_FIRST_DATA_ROW_STUDY + modConfig.STUDY_INITIAL_ROWS - 1)
    modSheetFormatting.FormatDataCalculatedRange dataRange.Columns(1).Resize(, 3)
    modSheetFormatting.FormatDataInputRange dataRange.Columns(4).Resize(, 2)
    modSheetFormatting.FormatDataCalculatedRange dataRange.Columns(6).Resize(, 7)

    ws.Range("F" & modConfig.TABLE_FIRST_DATA_ROW_STUDY & ":L" & modConfig.TABLE_FIRST_DATA_ROW_STUDY + modConfig.STUDY_INITIAL_ROWS - 1).NumberFormat = "0.000"

    footerTopRow = modConfig.TABLE_FIRST_DATA_ROW_STUDY + modConfig.STUDY_INITIAL_ROWS + 2
    modSheetFormatting.DrawDocumentFooter ws, footerTopRow, "P"
    modSheetFormatting.ApplyOuterFrame ws.Range("A1:P" & footerTopRow + 1)
End Sub

Private Sub WriteInstructionLine(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal stepNumber As String, ByVal instructionText As String)
    ws.Range("A" & rowNumber & ":B" & rowNumber).Merge
    ws.Range("A" & rowNumber).Value = stepNumber
    ws.Range("A" & rowNumber).Interior.Color = modSheetFormatting.COLOR_DARK_BLUE
    ws.Range("A" & rowNumber).Font.Color = modSheetFormatting.COLOR_WHITE
    ws.Range("A" & rowNumber).Font.Bold = True
    ws.Range("A" & rowNumber).HorizontalAlignment = xlCenter
    ws.Range("C" & rowNumber & ":P" & rowNumber).Merge
    ws.Range("C" & rowNumber).Value = instructionText
    ws.Range("C" & rowNumber).Interior.Color = modSheetFormatting.COLOR_LIGHT_BLUE
    ws.Range("C" & rowNumber).HorizontalAlignment = xlLeft
    ws.Range("A" & rowNumber & ":P" & rowNumber).Borders.LineStyle = xlContinuous
    ws.Range("A" & rowNumber & ":P" & rowNumber).Borders.Color = modSheetFormatting.COLOR_BORDER
End Sub

Private Sub WriteInstructionText(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal instructionText As String)
    ws.Range("A" & rowNumber & ":P" & rowNumber).Merge
    ws.Range("A" & rowNumber).Value = instructionText
    ws.Range("A" & rowNumber).Interior.Color = modSheetFormatting.COLOR_LIGHT_GRAY
    ws.Range("A" & rowNumber).HorizontalAlignment = xlLeft
    ws.Range("A" & rowNumber & ":P" & rowNumber).Borders.LineStyle = xlContinuous
    ws.Range("A" & rowNumber & ":P" & rowNumber).Borders.Color = modSheetFormatting.COLOR_BORDER
End Sub
