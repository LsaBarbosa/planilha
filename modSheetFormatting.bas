Attribute VB_Name = "modSheetFormatting"
Option Explicit

Private Const SHAPE_ROUNDED_RECTANGLE As Long = 5

Public Const COLOR_DARK_BLUE As Long = &H7B4E22
Public Const COLOR_LIGHT_BLUE As Long = 15984090
Public Const COLOR_LIGHT_GRAY As Long = 15132390
Public Const COLOR_BORDER As Long = 11382189
Public Const COLOR_WHITE As Long = 16777215
Public Const COLOR_BLACK As Long = 0

Private Const FONT_DEFAULT As String = "Calibri"
Private Const FONT_TITLE As String = "Calibri"
Private Const FONT_SIZE_DEFAULT As Double = 10
Private Const FONT_SIZE_HEADER As Double = 10
Private Const FONT_SIZE_TITLE As Double = 16

Public Sub ResetSheetCanvas(ByVal ws As Worksheet)
    Dim i As Long

    On Error Resume Next
    ws.Cells.Clear
    ws.Cells.ClearFormats
    ws.Cells.Validation.Delete
    On Error GoTo 0
    ws.Cells.Font.Name = FONT_DEFAULT
    ws.Cells.Font.Size = FONT_SIZE_DEFAULT
    ws.Cells.Font.Color = COLOR_BLACK
    ws.Cells.HorizontalAlignment = xlCenter
    ws.Cells.VerticalAlignment = xlCenter

    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    On Error GoTo 0
End Sub

Public Sub ApplyBaseVisualStandards(ByVal ws As Worksheet)
    With ws.Cells
        .Font.Name = FONT_DEFAULT
        .Font.Size = FONT_SIZE_DEFAULT
        .Font.Color = COLOR_BLACK
    End With

    ws.Rows.RowHeight = 20
    ws.Cells.WrapText = False
    ws.Cells.Borders.LineStyle = xlNone
    modConfig.ApplySheetWindowDefaults ws
End Sub

Public Sub AddLogoPlaceholder(ByVal ws As Worksheet, ByVal leftCell As String, ByVal widthPoints As Double, ByVal heightPoints As Double)
    Dim anchor As Range
    Dim shp As Shape

    DeleteShapeIfExists ws, "corpLogoPlaceholder"

    Set anchor = ws.Range(leftCell)
    Set shp = ws.Shapes.AddShape(SHAPE_ROUNDED_RECTANGLE, anchor.Left, anchor.Top, widthPoints, heightPoints)

    With shp
        .Name = "corpLogoPlaceholder"
        .Fill.ForeColor.RGB = RGB(33, 78, 122)
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 1.5
        .TextFrame2.TextRange.Characters.Text = "CORPORATE LOGO"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Name = FONT_DEFAULT
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub

Public Sub DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Public Sub FormatTitle(ByVal target As Range, ByVal textValue As String)
    target.Merge
    target.Value = textValue

    With target
        .Font.Name = FONT_TITLE
        .Font.Size = FONT_SIZE_TITLE
        .Font.Bold = True
        .Font.Color = RGB(34, 78, 123)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub

Public Sub FormatSectionHeader(ByVal rng As Range, ByVal textValue As String)
    rng.Merge
    rng.Value = textValue

    With rng
        .Interior.Color = RGB(34, 78, 123)
        .Font.Color = COLOR_WHITE
        .Font.Bold = True
        .Font.Size = FONT_SIZE_HEADER
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = COLOR_BORDER
    End With
End Sub

Public Sub FormatLabelCell(ByVal rng As Range, ByVal textValue As String)
    rng.Value = textValue

    With rng
        .Interior.Color = RGB(34, 78, 123)
        .Font.Color = COLOR_WHITE
        .Font.Bold = True
        .Font.Size = FONT_SIZE_HEADER
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = COLOR_BORDER
    End With
End Sub

Public Sub FormatInputCell(ByVal rng As Range, Optional ByVal defaultValue As String = "")
    rng.Value = defaultValue

    With rng
        .Interior.Color = COLOR_WHITE
        .Font.Color = COLOR_BLACK
        .Font.Bold = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Locked = False
        .Borders.LineStyle = xlContinuous
        .Borders.Color = COLOR_BORDER
    End With
End Sub

Public Sub FormatCalculatedCell(ByVal rng As Range, Optional ByVal defaultValue As String = "")
    rng.Value = defaultValue

    With rng
        .Interior.Color = COLOR_LIGHT_GRAY
        .Font.Color = COLOR_BLACK
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Locked = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = COLOR_BORDER
    End With
End Sub

Public Sub FormatHeaderRow(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(34, 78, 123)
        .Font.Color = COLOR_WHITE
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = COLOR_BORDER
    End With
End Sub

Public Sub FormatDataInputRange(ByVal rng As Range)
    With rng
        .Interior.Color = COLOR_LIGHT_BLUE
        .Font.Color = COLOR_BLACK
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = COLOR_BORDER
        .Locked = False
    End With
End Sub

Public Sub FormatDataCalculatedRange(ByVal rng As Range)
    With rng
        .Interior.Color = COLOR_LIGHT_GRAY
        .Font.Color = COLOR_BLACK
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = COLOR_BORDER
        .Locked = True
    End With
End Sub

Public Sub DrawDocumentHeader(ByVal ws As Worksheet, ByVal headerLastColumn As String)
    Dim labelColumns As Variant
    Dim valueRanges As Variant
    Dim i As Long

    labelColumns = Array("A1", "C1", "G1", "K1", "M1", "O1")
    valueRanges = Array("A2:B2", "C2:F2", "G2:J2", "K2:L2", "M2:N2", "O2:P2")

    ws.Range("A1:" & headerLastColumn & "2").RowHeight = 20

    FormatLabelCell ws.Range(labelColumns(0)), "JOB"
    FormatLabelCell ws.Range(labelColumns(1)), "SUBJECT"
    FormatLabelCell ws.Range(labelColumns(2)), "NOTES"
    FormatLabelCell ws.Range(labelColumns(3)), "CREATED BY"
    FormatLabelCell ws.Range(labelColumns(4)), "CHECKED BY"
    FormatLabelCell ws.Range(labelColumns(5)), "DATE"

    For i = LBound(valueRanges) To UBound(valueRanges)
        FormatMergedInputRange ws.Range(valueRanges(i))
    Next i
End Sub

Public Sub DrawDocumentFooter(ByVal ws As Worksheet, ByVal topRow As Long, ByVal footerLastColumn As String)
    Dim labelAddresses As Variant
    Dim valueRanges As Variant
    Dim i As Long

    labelAddresses = Array("A" & topRow, "D" & topRow, "F" & topRow, "I" & topRow, "L" & topRow, "O" & topRow)
    valueRanges = Array("A" & topRow + 1 & ":C" & topRow + 1, _
                        "D" & topRow + 1 & ":E" & topRow + 1, _
                        "F" & topRow + 1 & ":H" & topRow + 1, _
                        "I" & topRow + 1 & ":K" & topRow + 1, _
                        "L" & topRow + 1 & ":N" & topRow + 1, _
                        "O" & topRow + 1 & ":P" & topRow + 1)

    FormatLabelCell ws.Range(labelAddresses(0)), "Doc No."
    FormatLabelCell ws.Range(labelAddresses(1)), "REV."
    FormatLabelCell ws.Range(labelAddresses(2)), "Created by"
    FormatLabelCell ws.Range(labelAddresses(3)), "Checked by"
    FormatLabelCell ws.Range(labelAddresses(4)), "Approved by"
    FormatLabelCell ws.Range(labelAddresses(5)), "Date"

    For i = LBound(valueRanges) To UBound(valueRanges)
        FormatMergedInputRange ws.Range(valueRanges(i))
    Next i

    ws.Range("A" & topRow & ":" & footerLastColumn & topRow + 1).Borders.LineStyle = xlContinuous
    ws.Range("A" & topRow & ":" & footerLastColumn & topRow + 1).Borders.Color = COLOR_BORDER
End Sub

Public Sub ApplyOuterFrame(ByVal rng As Range)
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = COLOR_BORDER
        .Weight = xlThin
    End With
End Sub

Public Sub AutoFitReadableColumns(ByVal ws As Worksheet, ByVal fromColumn As String, ByVal toColumn As String)
    ws.Range(fromColumn & ":" & toColumn).Columns.AutoFit
End Sub

Public Sub SetStandardColumnWidths(ByVal ws As Worksheet)
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 12
    ws.Columns("K").ColumnWidth = 16
    ws.Columns("L").ColumnWidth = 16
    ws.Columns("M").ColumnWidth = 16
    ws.Columns("N").ColumnWidth = 14
    ws.Columns("O").ColumnWidth = 14
    ws.Columns("P").ColumnWidth = 14
End Sub


Public Sub FormatMergedInputRange(ByVal rng As Range, Optional ByVal defaultValue As String = "")
    rng.Merge
    FormatInputCell rng, defaultValue
End Sub

Public Sub FormatMergedCalculatedRange(ByVal rng As Range, Optional ByVal defaultValue As String = "")
    rng.Merge
    FormatCalculatedCell rng, defaultValue
End Sub

Public Sub WriteHeadersHorizontal(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal firstColumn As Long, ByVal headerValues As Variant)
    Dim i As Long

    For i = LBound(headerValues) To UBound(headerValues)
        ws.Cells(rowNumber, firstColumn + i).Value = headerValues(i)
    Next i
End Sub
