Sub SORTBE()
    '
    ' SORTBE Macro
    ' Macro recorded 07.10.2010 by r.manov
    '
    ' Keyboard Shortcut: Ctrl+Shift+E
    '
    Application.ScreenUpdating = False
    Rows("5:1700").Select

    'PARVO SORT PO COLUMN A
    Selection.Sort Key1:=Range("A5"), Order1:=xlAscending, Header:=xlNo, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
    DataOption1:=xlSortNormal

    'SA6TINSKI SORT PO COLUMN BE
    Selection.Sort Key1:=Range("BE5"), Order1:=xlAscending, Header:=xlNo, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
    DataOption1:=xlSortNormal
    Range("A2").Select
    ActiveWindow.ScrollRow = 1
    Application.ScreenUpdating = True
End Sub