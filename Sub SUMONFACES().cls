Sub SUMONFACES()
    '
    ' SUMONFACES Macro
    ' Macro recorded 16.03.2011 by r.manov
    '
    ' Keyboard Shortcut: Ctrl+Shift+S
    '

    Application.ScreenUpdating = False
    'SELECTIRA NA4ALNA KLETKA V SEDM.PRODAJBI
    Windows("01-SEDMI4NI PRODAJBI.xlsm").Activate
    Sheets("METRO").Select
 SUMARIZE:
    Cells(1711, "C").Select
    'PO4ISTVA SUMIRA6TITE REDOVE
    Selection.Resize(3, 150).Select
    Selection.ClearContents

    'ZAPO4VA SUMIRANE NA KARLOVO
    Windows("HELP - PROGNOZI.xls").Activate
    Sheets("SUM ON FACES (2)").Select
    Cells(5, "K").Select
 KARLOVO:
    Do Until ActiveCell.Value = 0
        N = ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
        Windows("01-SEDMI4NI PRODAJBI.xlsm").Activate
        'ZAPO4VA SA6INSKOTO TARSENE NA ART.NOMER
        Columns(1).Select
        Selection.Find(What:=N, After:=ActiveCell, LookIn:=xlValues, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
        br = ActiveCell.ROW
        If N = 0 Or br < 5 Or Cells(br, "BD").Value + Cells(br, "DI").Value = 0 Then
            'BELEJKA 20.08.2013 - PONEJE IMA PROMIANA V FUNKCIITE NA KOLKONI "DI" I "BC", ZA PROVERKATA SE POLZVAT PREDNITE KOLONI "BC" I "DH".
            Windows("HELP - PROGNOZI.xls").Activate
            GoTo KARLOVO
        End If
        'IZBIRA REDA NA PRODUKTA I GO KOPIRA
        Cells(br, 3).Select
        Selection.Resize(1, 110).Select
        Selection.Copy
        Cells(1712, 3).Select
        Selection.Resize(1, 110).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, _
        SkipBlanks:=False, Transpose:=False
        Windows("HELP - PROGNOZI.xls").Activate
        Sheets("SUM ON FACES (2)").Select
        ActiveCell.Select
    Loop


    'STARTIRA IZ4ISLIAVANE NA LOVE4
    'ZAPO4VA SUMIRANE NA LOVE4
    Windows("HELP - PROGNOZI.xls").Activate
    Sheets("SUM ON FACES (2)").Select
    Cells(5, "L").Select
 LOVE4:
    Do Until ActiveCell.Value = 0
        N = ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
        Windows("01-SEDMI4NI PRODAJBI.xlsm").Activate
        'ZAPO4VA SA6INSKOTO TARSENE NA ART.NOMER
        Columns(1).Select
        Selection.Find(What:=N, After:=ActiveCell, LookIn:=xlValues, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
        br = ActiveCell.ROW
        If N = 0 Or br < 5 Or Cells(br, "BD").Value + Cells(br, "DI").Value = 0 Then
            'BELEJKA 20.08.2013 - PONEJE IMA PROMIANA V FUNKCIITE NA KOLONI "DI" I "BC", ZA PROVERKATA SE POLZVAT PREDNITE KOLONI "BC" I "DH".
            Windows("HELP - PROGNOZI.xls").Activate
            GoTo LOVE4
        End If
        'IZBIRA REDA NA PRODUKTA I GO KOPIRA
        Cells(br, 3).Select
        Selection.Resize(1, 110).Select
        Selection.Copy
        Cells(1711, 3).Select
        Selection.Resize(1, 110).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, _
        SkipBlanks:=False, Transpose:=False
        Windows("HELP - PROGNOZI.xls").Activate
        Sheets("SUM ON FACES (2)").Select
        ActiveCell.Select
    Loop

    'PRAVI SUMIRA6T RED , SMENIA SHIITA I ZAPO4VA OTNA4ALO
    Windows("01-SEDMI4NI PRODAJBI.xlsm").Activate
    'PRAVI SUMIRA6T RED I UNI6TOJAVA FORMULATA
    Range("C1713").Select
    ActiveCell.FormulaR1C1 = "=R[-2]C+R[-1]C"
    Selection.Copy
    Selection.Resize(1, 110).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    Cells(2, 1).Select
    'SMENIA SHIITA I ZAPO4VA OTNA4ALO
    ActiveSheet.Next.Select
    If ActiveSheet.Name = ("SUM") Then End

        GoTo SUMARIZE

        Application.ScreenUpdating = True
End Sub
