

Sub DecemberCorrector()
'r.manov - 09.05.2023
Application.ScreenUpdating = False
' select file "01-SEDMI4NI PRODAJBI.xlsm" and make it active
Workbooks("01-SEDMI4NI PRODAJBI.xlsm").Activate
' select sheet "METRO" in "01-SEDMI4NI PRODAJBI" and make it active
Sheets("METRO").Activate
' select cell "A1" in "METRO" and make it active

CorrectorDecember:

    Range("DD5").Activate
    'while active row is not 1700 - do
    Do While ActiveCell.Row <> 1700
        R = ActiveCell.Row
        K = ActiveCell.Column
        'if 'DI' is equal to '0' - go to next row
        If Cells(R, K + 5).Value = 0 Then
        ActiveCell.Offset(1, 0).Activate
        
        'if 'DI' is not equal to '0' - do
        Else
            R = ActiveCell.Row
            K = ActiveCell.Column
        'copy value from 'DU' and multiply with that number range from 'Dd' to 'Dh' - do this while 'du' is equal to number in range '0.9' to'1.001'
        Do While Cells(R, K + 17).Value >= 0.9 And Cells(R, K + 17).Value < 1.0000
            R = R0
            K = K0

            Cells(R, K + 17).Copy
            'check if current cell is 0 -go to next row
            If Cells(R, K + 17).Value = 0 Then ActiveCell.Offset(1, 0).Activate
            'if current cell is not 0 - check if color is 36 or 39 - go to next cell
            If Cells(R, K).Interior.ColorIndex = 36 Or Cells(R, K).Interior.ColorIndex = 39 Then ActiveCell.Offset(0, 1).Activate
            'if current cell is not 0 and color is not 36 or 39 - check if cells in range 'ay' to'bc' have color 36 or 39
            If Cells(R, K - 53).Interior.ColorIndex <> 36 And Cells(R, K - 53).Interior.ColorIndex <> 39 Then
                For K = 108 To K + 4
                    Cells(R, K).PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply, SkipBlanks _
                            :=False, Transpose:=False
                    ActiveCell.Offset(0, 1).Activate
                Next K
                'if column count is equal to 112 - go to next row
                If K > 112 Then
                    ActiveCell.Offset(1, -5).Activate
                    Exit Do
                End If
            Else
                ActiveCell.Offset(1, 0).Activate
            End If
        Loop
        End If
    Loop
'if row is 1700 - go to next sheet
'if sheet is not "EVROPA", "PICCADYLLY", "CARREFOUR", "PENNY", "SUM" - go to next sheet
If ActiveSheet.Name <> "EVROPA" And ActiveSheet.Name <> "PICCADYLLY" And ActiveSheet.Name <> "CARREFOUR" And ActiveSheet.Name <> "PENNY" And ActiveSheet.Name <> "SUM" Then
    ActiveSheet.Next.Select
End If
'star from begining
GoTo CorrectorDecember
Application.ScreenUpdating = True
End Sub
