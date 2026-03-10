Sub SORTDS()
    '
    ' Macro2 Macro
    '

    '
    Rows("5:1700").Select
    Selection.Sort Key1:=Range("DS5"), Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
    DataOption1:=xlSortNormal

End Sub
