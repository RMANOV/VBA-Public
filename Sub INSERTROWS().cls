Sub INSERTROWS()
    '
    ' INSERTROWS Macro
    ' Macro recorded 10.7.2010 by r.manov
    '
    ' Keyboard Shortcut: Ctrl+Shift+I
    '
    Dim Rng
    Rng = InputBox("Enter number of rows required.")
    Range(ActiveCell.Offset(0, 0), ActiveCell.Offset(Rng - 1, 0)).Select
    Selection.EntireRow.Insert
End Sub
