Option Explicit

' The contents of this file should be placed in the Workbook itself.

' in Immediate Window (View->Immediate Window or ctrl-g), type:
' OldSchoolMenu()
' to start.


Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)
    ' Runs whenever the selection changes, cursor or mouse move.

    Dim maxRowNumber As Long
    Dim maxColNumber As Long

    maxRowNumber = DetectMaxRows()
    maxColNumber = DetectMaxColumns()

    If ActiveCell.Row <= maxRowNumber Then
        If ActiveCell.Column <= maxColNumber Then
            ' if we are the in old school range

            Call PaintContent

            ActiveCell.Interior.color = GetColors().activeCellBgColor
            ' ActiveCell.Font.color = GetColors.activeCellFgColor

            Call PaintHeaders
            Call FormatRange
        End If
    End If
End Sub

