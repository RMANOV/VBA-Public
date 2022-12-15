Sub PASTEVALUESMULTIPLY()
    '
    ' PASTEVALUESMULTIPLY Macro
    ' Macro recorded 16.12.2010 by r.manov
    '

    '
    Application.ScreenUpdating = False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply, _
    SkipBlanks:=False, Transpose:=False
    Application.ScreenUpdating = True
End Sub