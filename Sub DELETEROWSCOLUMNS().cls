Sub DELETEROWSCOLUMNS()
    '
    ' DELETEROWSCOLUMNS Macro
    ' Macro recorded 29.11.2010 by r.manov
    '

    '
    Dim I As Long
    Dim J As Long


    With Application

        .Calculation = xlCalculationManual

        .ScreenUpdating = False

        For J = Selection.Columns.Count To 1 Step -1

            If WorksheetFunction.CountA(Selection.Columns(J)) = 0 Then

                Selection.Columns(J).EntireColumn.Delete

            End If

        Next J

        For I = Selection.Rows.Count To 1 Step -1

            If WorksheetFunction.CountA(Selection.Rows(I)) = 0 Then

                Selection.Rows(I).EntireRow.Delete

            End If

        Next I
        .Calculation = xlCalculationAutomatic

        .ScreenUpdating = True

    End With

End Sub
