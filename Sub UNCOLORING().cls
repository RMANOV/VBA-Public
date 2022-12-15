Sub UNCOLORING()
    '
    ' UNCOLORING Macro
    ' Macro recorded 17.11.2010 by r.manov
    '
    ' Keyboard Shortcut: Ctrl+Shift+U
    '
    Application.ScreenUpdating = False
    Sheets(Array("METRO", "BILLA", "KAUFLAND", "LIDL", "CARREFOUR", "T-MARKET", _
    "PICCADILLY", "PENNY", "FANTASTIKO", "EVROPA", "DISTRIBUCIA", "SUM")).Select
    Application.ReplaceFormat.Interior.ColorIndex = xlNone
    Application.FindFormat.Interior.ColorIndex = 35
    Selection.Replace What:="", Replacement:="", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=True, _
    ReplaceFormat:=True
    Application.FindFormat.Interior.ColorIndex = 4
    Selection.Replace What:="", Replacement:="", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=True, _
    ReplaceFormat:=True
    Application.FindFormat.Interior.ColorIndex = 6
    Selection.Replace What:="", Replacement:="", LookAt:=xlWhole, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=True, _
    ReplaceFormat:=True
    Range("A2").Select
    Sheets("METRO").Select
    Application.ScreenUpdating = True
End Sub
