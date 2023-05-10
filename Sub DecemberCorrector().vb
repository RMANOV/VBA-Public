Sub DecemberCorrector()
    ' r.manov - 09.05.2023
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Set wb = Workbooks("01-SEDMI4NI PRODAJBI.xlsm")
    
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Not IsError(Application.Match(ws.Name, Array("METRO", "EVROPA", "PICCADYLLY", "CARREFOUR", "PENNY", "SUM"), 0)) Then
            CorrectorDecember ws
        End If
    Next ws
    
    Application.ScreenUpdating = True
End Sub

Sub CorrectorDecember(ws As Worksheet)
    Dim R As Long, K As Long
    
    For R = 5 To 1699
        If ws.Cells(R, 112).Value <> 0 Then
            If ws.Cells(R, 125).Value >= 0.9 And ws.Cells(R, 129).Value < 1.01 Then
                If ws.Cells(R, 61).Interior.ColorIndex <> 36 And ws.Cells(R, 61).Interior.ColorIndex <> 39 Then
                    For K = 108 To 112
                        If ws.Cells(R, K).Interior.ColorIndex <> 36 And ws.Cells(R, K).Interior.ColorIndex <> 39 Then
                            ' ws.Cells(R, K).Value = ws.Cells(R, K).Value * ws.Cells(R, 125).Value
                            'sqrt(sqrt(sqrt(Cells(R, 125).Value)))
                            ws.Cells(R, K).Value = ws.Cells(R, K).Value * Sqr(Sqr(Sqr(ws.Cells(R, 125).Value)))
                        End If
                    Next K
                End If
            End If
        End If
    Next R
End Sub
