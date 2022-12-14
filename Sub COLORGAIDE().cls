Sub COLORGAIDE()
    '
    ' COLORGAIDE Macro
    '

    '
    Range("E4").Select
    For I = 0 To 56
        Selection.Interior.ColorIndex = I
        ActiveCell.Offset(1, 0).Select
    Next I

End Sub
