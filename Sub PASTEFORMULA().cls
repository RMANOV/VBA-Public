Sub PASTEFORMULA()
    '
    ' PASTEFORMULA Macro
    '

    '
    Application.ScreenUpdating = False
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
    Application.ScreenUpdating = True
End Sub