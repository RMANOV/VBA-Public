Sub DELETEDUPLICATES()
    '
    ' DELETEDUPLICATES Macro
    ' Macro recorded 10.7.2010 by r.manov
    '
    ' Keyboard Shortcut: Ctrl+Shift+Z
    '
    Dim R As Long
    Dim N As Long
    Dim V As Variant
    Dim Rng As Range

    On Error GoTo EndMacro
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual


    Set Rng = Application.Intersect(ActiveSheet.UsedRange, _
    ActiveSheet.Columns(ActiveCell.Column))

    Application.StatusBar = "Processing Row: " & Format(Rng.ROW, "#,##0")

    N = 0
    For R = Rng.Rows.Count To 2 Step -1
        If R Mod 500 = 0 Then
            Application.StatusBar = "Processing Row: " & Format(R, "#,##0")
        End If

        V = Rng.Cells(R, 1).Value
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Note that COUNTIF works oddly with a Variant that is equal to vbNullString.
        ' Rather than pass in the variant, you need to pass in vbNullString explicitly.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If V = vbNullString Then
            If Application.WorksheetFunction.CountIf(Rng.Columns(1), vbNullString) > 1 Then
                Rng.Rows(R).EntireRow.Delete
                N = N + 1
            End If
        Else
            If Application.WorksheetFunction.CountIf(Rng.Columns(1), V) > 1 Then
                Rng.Rows(R).EntireRow.Delete
                N = N + 1
            End If
        End If
    Next R

 EndMacro:

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Duplicate Rows Deleted: " & CStr(N)

End Sub
