Sub FIRSTGREEN()
    '
    ' FIRSTGREEN Macro
    ' Macro recorded 20.01.2011 by r.manov
    '

    '

    Application.ScreenUpdating = False
    ActiveCell.Select
    Do Until Selection.Interior.ColorIndex = 35 Or Selection.Interior.ColorIndex = 4 Or Selection.Interior.ColorIndex = 6
        ActiveCell.Offset(0, 1).Select
    Loop
    Application.ScreenUpdating = True
    ActiveCell.Select
End Sub
