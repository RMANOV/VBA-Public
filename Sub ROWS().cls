Sub ROW()
    '
    ' ROW Macro
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