Sub DeleteBlankCOLUMNs1()

    'Deletes the entire COLUMN within the selection if the ENTIRE COLUMN contains no data.



    'We use Long in case they have over 32,767 COLUMNs selected.

    Dim I As Long



    'We turn off calculation and screenupdating to speed up the macro.

    With Application

        .Calculation = xlCalculationManual

        .ScreenUpdating = False



        'We work backwards because we are deleting COLUMNs.

        For I = Selection.Columns.Count To 1 Step -1

            If WorksheetFunction.CountA(Selection.Columns(I)) = 0 Then

                Selection.Columns(I).EntireColumn.Delete

            End If

        Next I



        .Calculation = xlCalculationAutomatic

        .ScreenUpdating = True

    End With

End Sub
