Sub PASTEFORMULAANDDESTRUCT()
    '
    ' PASTEFORMULAANDDESTRUCT Macro
    ' Macro recorded 16.12.2010 by r.manov
    '

    '
    Application.ScreenUpdating = False
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Selection.Replace What:="#N/A", Replacement:="0", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Application.ScreenUpdating = True
End Sub
