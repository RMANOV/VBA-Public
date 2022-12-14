Sub BEGININGLEVELSNEWYEAR()
    '
    ' BEGININGLEVELSNEWYEAR Macro
    '

    '

    Application.ScreenUpdating = False

    'DA SE VGRADI ZA6TITA ZA SPRENI PRODUKTI
    'OT4ITANE NA TRENDA

    Windows("01-SEDMI4NI PRODAJBI.xlsm").Activate
    Sheets("METRO").Select
    Cells(5, 3).Select
    R = ActiveCell.ROW
    K = ActiveCell.Column

 NEXTSELECT:
    Cells(R, 3).Select
    'PRESKA4A PRAZNITE REDOVE
    Do Until Cells(R, 54).Value <> 0
        ActiveCell.Offset(1, 0).Select
        R = R + 1
        If R = 1000 Then
            ActiveSheet.Next.Select
            If ActiveSheet.Name = ("SUM") Or ActiveSheet.Name = ("EVROPA") Or ActiveSheet.Name = ("PICCADILLY") Or ActiveSheet.Name = ("CARREFOUR") Or ActiveSheet.Name = ("PENNY") Then End
                R = 5
            End If
        Loop

        'SELEKTIRA CELIA RED DO 49 SEDMICA VKLU4ITELNO
        Selection.Resize(1, 50).Select

        'ZAPO4VA IZ4ISLIAVANE NA SREDNATA NA PRODAJBITE
        CNT2 = Selection.Count 'KOLI4ESTVO NA KLETKI S REALNI PRODAJBI
        AVGSALES = 0
        BRLILAVI = 0
        BRAKTIVNI = 0

        Cells(R, 3).Select
        For I = 0 To CNT2 - 1
            If ActiveCell.Interior.ColorIndex = 39 Or ActiveCell.Interior.ColorIndex = 36 Then BRLILAVI = BRLILAVI + 1
                If ActiveCell.Interior.ColorIndex <> 39 And ActiveCell.Interior.ColorIndex <> 36 And ActiveCell.Value <> 0 Then
                    AVGSALES = AVGSALES + ActiveCell.Value
                    BRAKTIVNI = BRAKTIVNI + 1
                End If
                ActiveCell.Offset(0, 1).Select
            Next I

            If BRAKTIVNI = 0 Then AVGSALES = (AVGSALES / BRLILAVI) Else AVGSALES = (AVGSALES / BRAKTIVNI)
                ' Calculate time-weighted average - last weeks are more important - only for the BRAKTIVNI
                AVGSALES = Application.WorksheetFunction.WeightedAverage(Selection, 1, 1, 1, 1)



                Cells(R, 60).Value = AVGSALES
                R = R + 1
                K = K + 1
                If R = 1000 Then
                    ActiveSheet.Next.Select
                    If ActiveSheet.Name = ("SUM") Or ActiveSheet.Name = ("EVROPA") Or ActiveSheet.Name = ("PICCADILLY") Or ActiveSheet.Name = ("CARREFOUR") Or ActiveSheet.Name = ("PENNY") Then End
                        R = 5
                    End If
                    GoTo NEXTSELECT




                    Application.ScreenUpdating = True
End Sub
