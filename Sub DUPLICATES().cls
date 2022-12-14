Sub DUPLICATES()
    '
    ' DUPLICATES Macro
    ' Macro recorded 07.7.2010 by r.manov
    '
    ' Keyboard Shortcut: Ctrl+Shift+D
    '
    Application.ScreenUpdating = False
    Rng = Selection.Rows.Count
    For I = Rng To 1 Step -1
        myCheck = ActiveCell
        ActiveCell.Offset(1, 0).Select
        For J = 1 To I
            If ActiveCell = myCheck Then
                Selection.Font.Bold = True
                Selection.Font.ColorIndex = 3
            End If
            ActiveCell.Offset(1, 0).Select
        Next J
        ActiveCell.Offset(-I, 0).Select
    Next I
    Application.ScreenUpdating = True
End Sub
