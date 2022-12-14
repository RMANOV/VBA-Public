Sub CHISTA4IZPALNENIE()
    '
    ' CHISTA4IZPALNENIE Macro
    ' Macro recorded 29.10.2010 by r.manov
    '

    '
    Sheets(Array("CHART 1", "ÏÎ ÂÅÐÈÃÈ", "ÎÁÎÁÙÅÍÈÅ", "ÎÁÎÁÙÅÍÈÅ - 2012", "ÏÎ ÂÅÐÈÃÈ - 2012", "CHART 1 - 2012 ", _
    "ÎÁÎÁÙÅÍÈÅ - 2011", "ÏÎ ÂÅÐÈÃÈ - 2011", "CHART - 2011")).Select

    Application.DisplayAlerts = False
    ActiveWindow.SelectedSheets.Delete
    Application.DisplayAlerts = True
    Rows("61:90").Select
    Selection.Delete Shift:=xlUp
    Range("A2").Select
    Sheets("ËÎÂÅ×").Select
    Rows("248:280").Select
    Selection.Delete Shift:=xlUp
    Range("A2").Select
    ActiveWorkbook.Save
End Sub
