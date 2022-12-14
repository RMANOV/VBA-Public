

Sub Moving_old_data_to_old_table()
    '
    ' Moving_old_data_to_old_table Macro
    '
    ' Keyboard Shortcut: Ctrl+q
    Application.ScreenUpdating = False

    Range("BH1:DH3").Select
    Range("DH3").Activate
    Selection.Copy

    Range("C1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False


    Range("BH5:DH1700").Select
    Range("DH1700").Activate
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.ScrollColumn = 3
    Range("C5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteComments, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
    Range("BH5:DH1700").Select
    Range("DH1700").Activate
    Application.CutCopyMode = False
    Selection.ClearContents
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13434828
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Replace What:="", Replacement:="0", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Range("BH5:DH1700").Select
    Range("DH1700").Activate

    Application.ScreenUpdating = True
End Sub