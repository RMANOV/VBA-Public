Sub PASTEVALUESDIVIDE()
    '
    ' PASTEVALUESDIVIDE Macro
    '

    '
    Application.ScreenUpdating = False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlDivide, _
    SkipBlanks:=False, Transpose:=False
    Application.ScreenUpdating = True
End Sub