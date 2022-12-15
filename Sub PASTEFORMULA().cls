Sub PASTEFORMULA()
    '
    ' PASTEFORMULA Macro
    ' Macro recorded 16.12.2010 by r.manov
    '

    '
    Application.ScreenUpdating = False
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
    Application.ScreenUpdating = True
End Sub