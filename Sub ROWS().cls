Sub ROW()
    '
    ' ROW Macro
    ' Macro recorded 07.10.2010 by r.manov
    '
    ' Keyboard Shortcut: Ctrl+Shift+K
    '
    Application.ScreenUpdating = False
    ActiveCell.Select
    With ActiveWindow
        .ScrollRow = ActiveCell.ROW
        ActiveCell.EntireRow.Select
    End With
    Application.ScreenUpdating = True
End Sub