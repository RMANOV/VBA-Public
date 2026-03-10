Sub PASTEVALUESMULTIPLY()
    '
    ' PASTEVALUESMULTIPLY Macro
    '

    '
    Application.ScreenUpdating = False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply, _
    SkipBlanks:=False, Transpose:=False
    Application.ScreenUpdating = True
End Sub