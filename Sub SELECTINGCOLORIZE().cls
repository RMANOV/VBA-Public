Sub SELECTINGCOLORIZE()
    '
    ' SELECTINGCOLORIZE Macro
    ' Macro recorded 19.03.2011 by r.manov
    '
    ' Keyboard Shortcut: Ctrl+Shift+G
    '

    Application.ScreenUpdating = False
    D = Selection.Rows.Count
    Selection.Name = "RNG"
    Range("RNG").Select
    R = ActiveCell.ROW
    K = ActiveCell.Column
    If Selection.Interior.ColorIndex < 1 Then
        Do While Selection.Interior.ColorIndex < 1
            ActiveCell.Offset(0, 1).Select
        Loop
    End If
    Do Until Selection.Interior.ColorIndex < 1
        ActiveCell.Offset(0, 1).Select
    Loop
    ActiveCell.Offset(0, -1).Select
    K2 = ActiveCell.Column
    Cells(R, K).Select
    B = K2 - K + 1
    Selection.Resize(D, B).Select
    If Not ActiveWorkbook.Name = ("01-SEDMI4NI PRODAJBI.xlsm") Then
        Selection.Copy
    End If
    Application.ScreenUpdating = True
End Sub
