Sub SORTKA()
    '
    ' SORTKA Macro
    '
    ' Keyboard Shortcut: Ctrl+Shift+A
    '
    Application.ScreenUpdating = False
    Rows("5:1700").Select
    Selection.Sort Key1:=Range("A5"), Order1:=xlAscending, Header:=xlNo, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
    DataOption1:=xlSortNormal
    Range("A2").Select
    ActiveWindow.ScrollRow = 1
    Application.ScreenUpdating = True
End Sub