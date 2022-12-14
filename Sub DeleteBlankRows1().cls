Sub DeleteBlankRows1()

    'Deletes the entire row within the selection if the ENTIRE row contains no data.



    'We use Long in case they have over 32,767 rows selected.

    Dim I As Long



    'We turn off calculation and screenupdating to speed up the macro.

    With Application

        .Calculation = xlCalculationManual

        .ScreenUpdating = False



        'We work backwards because we are deleting rows.

        For I = Selection.Rows.Count To 1 Step -1

            If WorksheetFunction.CountA(Selection.Rows(I)) = 0 Then

                Selection.Rows(I).EntireRow.Delete

            End If

        Next I



        .Calculation = xlCalculationAutomatic

        .ScreenUpdating = True

    End With

End Sub
