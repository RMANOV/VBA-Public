Sub PASTEVALUESDIVIDE()
    '
    ' PASTEVALUESDIVIDE Macro
    ' Macro recorded 24.06.2011 by r.manov
    '

    '
    Application.ScreenUpdating = False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlDivide, _
    SkipBlanks:=False, Transpose:=False
    Application.ScreenUpdating = True
End Sub