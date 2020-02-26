Option Explicit

' oldschool_sheet.bas
' This code should be put in a module.

' in Immediate Window (View->Immediate Window or ctrl-g), type:
' call OldSchoolMenu()
' to start.

Const FORMATTING_RANGE = True

' the following options will be efective if FORMATTING_RANGE = True
Const FORMATTING_RANGE_FONT_NAME = "Consolas"
Const FORMATTING_RANGE_FONT_SIZE = 12
Const FORMATTING_RANGE_WRAP_TEXT = False
Const FORMATTING_RANGE_ROW_HEIGHT = 14.4

Const DEFAULT_RANGE = "BB200"


Type ColorScheme
    headerBgColor As Long
    headerFgColor As Long
    contentBgColor As Long
    contentFgColor As Long
    activeCellBgColor As Long
    activeCellFgColor As Long
End Type


Public Function GetColorsOfLotus() As ColorScheme
    Dim activeColors As ColorScheme

    activeColors.headerBgColor = RGB(51, 204, 204)  ' 33CCCC, blue
    activeColors.headerFgColor = RGB(0, 0, 0)  ' black
    activeColors.contentBgColor = RGB(0, 0, 0)  ' black
    activeColors.contentFgColor = RGB(255, 255, 255)  ' 33CCCC, blue
    activeColors.activeCellBgColor = RGB(51, 204, 204)  ' 33CCCC, blue
    activeColors.activeCellFgColor = RGB(255, 255, 255)  ' white

    GetColorsOfLotus = activeColors
End Function


Public Function GetColors() As ColorScheme
    ' Returns the active colors.
    Dim activeColors As ColorScheme
    activeColors = GetColorsOfLotus()
    GetColors = activeColors
End Function


Public Function DetectMaxColumns() As Long
    ' Find max column number until white.
    Dim j As Long
    Dim jmax As Long
    Dim cellColor As Long
    ' find max rows until white
    For j = 1 To Rows.Count
        cellColor = Cells(1, j).Interior.color
        If cellColor = RGB(255, 255, 255) Then
            Exit For
        Else
            jmax = j
        End If
    Next j
    ' Debug.Print "GetMaxColumns():" & jmax
    DetectMaxColumns = jmax
End Function


Public Function DetectMaxRows() As Long
    ' Find max row number until white.
    Dim i As Long
    Dim imax As Long
    Dim cellColor As Long
    For i = 1 To Rows.Count
        cellColor = Cells(i, 1).Interior.color
        If cellColor = RGB(255, 255, 255) Then
            Exit For
        Else
            imax = i
        End If
    Next i
    ' Debug.Print "GetMaxRows():" & imax
    DetectMaxRows = imax
End Function


Public Sub PaintHeaderColumn()
    ' Paints the old school title column.
    ' On Error GoTo 0  ' to disable error handling.
    Dim i As Long
    Dim cellColor As Long
    Dim imax As Long
    Dim address1 As String
    Dim range1 As String
    imax = DetectMaxRows()

    If imax = 0 Then
        ' this is not an old school sheet, simply exit.
        Exit Sub
    End If

    ' if you add a new sheet, and it is not a OldSchool sheet,
    ' the following line would raise an error.
    ' https://stackoverflow.com/questions/17980854/vba-runtime-error-1004-application-defined-or-object-defined-error-when-select
    address1 = Cells(imax, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    ' Debug.Print "PaintHeaderColumn():address1:" & address1

    range1 = "A1:" & address1
    ' Debug.Print "PaintHeaderColumn():range1:" & range1
    
    Range(range1).Interior.color = GetColors.headerBgColor
    Range(range1).Font.color = GetColors.headerFgColor
End Sub


Public Sub PaintHeaderRow()
    ' Paints the old school title row.
    ' On Error GoTo 0  ' to disable error handling.
    Dim j As Long
    Dim jmax As Long
    Dim address1 As String
    Dim range1 As String
    jmax = DetectMaxColumns()

    If jmax = 0 Then
        ' this is not an old school sheet, simply exit.
        Exit Sub
    End If

    ' if you add a new sheet, and it is not a OldSchool sheet,
    ' the following line would raise an error.
    ' https://stackoverflow.com/questions/17980854/vba-runtime-error-1004-application-defined-or-object-defined-error-when-select
    address1 = Cells(1, jmax).Address(RowAbsolute:=False, ColumnAbsolute:=False)

    range1 = "A1:" & address1
    Range(range1).Interior.color = GetColors.headerBgColor
    Range(range1).Font.color = GetColors.headerFgColor
End Sub


Public Function PaintHeaders()
    ' Paints the old school title row and column.
    Call PaintHeaderRow
    Call PaintHeaderColumn
End Function


Public Sub PaintContent()
    ' Paints the old school range, but only the content area.
    Dim address1 As String
    Dim range1 As String

    Dim i As Long
    Dim j As Long
    Dim imax As Long
    Dim jmax As Long

    imax = DetectMaxRows()
    jmax = DetectMaxColumns()

    If Not (imax > 0 And jmax > 0) Then
        Exit Sub
    End If

    address1 = Cells(imax, jmax).Address(RowAbsolute:=False, ColumnAbsolute:=False)

    range1 = "B2:" & address1
    Range(range1).Interior.color = GetColors.contentBgColor
    Range(range1).Font.color = GetColors.contentFgColor
End Sub


Sub FormatRange()
    ' Formats the range.
    Dim maxRowNumber As Long
    Dim maxColNumber As Long
    Dim sheet As Worksheet
    Dim targetRange As String  ' a string like "3:9"

    maxRowNumber = DetectMaxRows()
    maxColNumber = DetectMaxColumns()
    targetRange = "A1:" & Cells(maxRowNumber, maxColNumber).Address
    ' Cells(r, c).Address
    
    Range(targetRange).Font.Name = FORMATTING_RANGE_FONT_NAME
    Range(targetRange).Font.Size = FORMATTING_RANGE_FONT_SIZE
    Range(targetRange).WrapText = FORMATTING_RANGE_WRAP_TEXT
    Range(targetRange).VerticalAlignment = xlTop
    Range(targetRange).rowHeight = FORMATTING_RANGE_ROW_HEIGHT

    ' Range(targetRange).Select
    ' With Selection.Font
    '     .Name = "Consolas"
    '     .Size = 12
    '     .Strikethrough = False
    '     .Superscript = False
    '     .Subscript = False
    '     .OutlineFont = False
    '     .Shadow = False
    '     .Underline = xlUnderlineStyleNone
    '     .TintAndShade = 0
    '     .ThemeFont = xlThemeFontNone
    ' End With
    
    ' With Selection
    '     .HorizontalAlignment = xlGeneral
    '     .VerticalAlignment = xlTop
    '     .WrapText = True
    '     .Orientation = 0
    '     .AddIndent = False
    '     .IndentLevel = 0
    '     .ShrinkToFit = False
    '     .ReadingOrder = xlContext
    '     .MergeCells = False
    ' End With

End Sub


Sub PaintInitially()
    ' Paints a range initially in the active sheet.
    ' This sub is called from OldSchoolMenu()
    Dim default As String
    Dim title As String
    Dim range1 As String
    Dim answer As String
    
    Dim activeColors As ColorScheme
    activeColors = GetColors()

    ' paint screen
    title = "Paint Content"
    default = DEFAULT_RANGE
    answer = InputBox("From A1, paint the sheet until which cell?", title, default)
    range1 = "A1:" & answer
        
    Range(range1).Interior.color = GetColors.contentBgColor
    Call PaintHeaders
    Call PaintContent
End Sub


Sub OldSchoolMenu()
    ' Displays a simple Old School menu for easy usage
    Dim default As String
    Dim range1 As String
    
    Dim title As String
    title = "Old School Menu"
    
    Dim prompt As String
    prompt = prompt & "p - initial paint of the range" & vbCrLf
    prompt = prompt & "f - format range" & vbCrLf
    prompt = prompt & "q - quit" & vbCrLf

    Dim answer As String
    answer = InputBox(prompt, title)
    answer = LCase(Trim(answer))
    If answer = "p" Then
        Call PaintInitially
    ElseIf answer = "f" Then
        Call FormatRange
    End If
End Sub

