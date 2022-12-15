Sub QVDELETECOLUMN()
    '
    ' QVDELETECOLUMN Macro
    ' Macro recorded 06.04.2011 by r.manov
    '

    '

    Application.ScreenUpdating = False
    Cells(3, 1).Select
    Do While Selection.Interior.ColorIndex > 1
        ActiveCell.Offset(0, 1).Select
        If ActiveCell.ColumnWidth <= 2 Then ActiveCell.EntireColumn.Delete
        Loop

        Columns("B:B").Select
        Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

        Cells(1, 1).Select
        ActiveWorkbook.Save
        Application.ScreenUpdating = True
End Sub
