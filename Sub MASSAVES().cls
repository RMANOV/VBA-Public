Sub MASSAVES()
    '
    ' MASSAVES Macro
    '
    ' Keyboard Shortcut: Ctrl+Shift+S
    '

'     Application.ScreenUpdating = False

'     W = Workbooks.Count
'     For I = 1 To W
'         ActiveWorkbook.Save
'         Application.ActiveWindow.ActivateNext
'     Next I

'     Application.ScreenUpdating = True
' End Sub
Application.ScreenUpdating = False
    While Boolean = True
        ActiveWorkbook.Save
        Application.ActiveWindow.ActivateNext
        
    End While
Application.ScreenUpdating = True
End Sub
