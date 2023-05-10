Sub DecemberCorrector()
    ' r.manov - 09.05.2023
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Set wb = Workbooks("01-SEDMI4NI PRODAJBI.xlsm")
    
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If IsError(Application.Match(ws.Name, Array("EVROPA", "PICCADYLLY", "CARREFOUR", "PENNY", "SUM"), 0)) Then
            CorrectorDecember ws
        End If
    Next ws
    
    Application.ScreenUpdating = True
End Sub

Sub CorrectorDecember(ws As Worksheet)
    Dim R As Long, K As Long
    Dim skipRow As Boolean
    
    For R = 5 To 1699
        'if r1 and r2 are empty - change sheet
        'if r1 and r2 are not empty - continue
        If ws.Cells(R, 1).Value = "" And ws.Cells(R, 2).Value = "" Then
            Exit For
        End If
        
        skipRow = False
        
        If ws.Cells(R, 113).Value <> 0 Then
            If Not (ws.Cells(R, 125).Value >= 0.9 And ws.Cells(R, 125).Value <= 1.001) Then
                
                For K = 51 To 55
                    If ws.Cells(R, K).Interior.ColorIndex = 36 Or ws.Cells(R, K).Interior.ColorIndex = 39 Then
                        skipRow = True
                        Exit For
                    End If
                Next K
                
                If Not skipRow Then
                    For K = 108 To 112
                        If ws.Cells(R, K).Interior.ColorIndex = 36 Or ws.Cells(R, K).Interior.ColorIndex = 39 Then
                            skipRow = True
                            Exit For
                        End If
                    Next K
                End If
                
                If Not skipRow Then
                    Do
                        ' ws.Cells(R, 108).Resize(1, 5).Value = Application.WorksheetFunction.Transpose(Application.WorksheetFunction.Transpose(ws.Cells(R, 108).Resize(1, 5).Value)) * ws.Cells(R, 125).Value
                        ' ws.Cells(R, 125).Value = Application.WorksheetFunction.Max(0, ws.Cells(R, 125).Value - Application.WorksheetFunction.Sum(ws.Cells(R, 108).Resize(1, 5).Value))
                        'select and copy cell 125
                        ws.Cells(R, 125).Select
                        ws.Cells(R, 125).Copy
                        'select cell 108
                        ws.Cells(R, 108).Select
                        'resize the selection to 5 columns to 112
                        ws.Cells(R, 108).Resize(1, 5).Select
                        'multiply the values in the selection with the value in cell 125 at once
                        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply, SkipBlanks:=False, Transpose:=False
                        'check if the value in cell 125 is between 0.9 and 1.001 - if not - copy again new value of cell 125 and repeat 5.3 and 5.4
                    Loop While Not (ws.Cells(R, 125).Value >= 0.9 And ws.Cells(R, 125).Value <= 1.001)
                End If
            End If
        End If
    Next R
End Sub

'0.check if sheet name is EVROPA or PICCADYLLY or CARREFOUR or PENNY or SUM - starting from METRO - if it is not then continue, but if it is then skip the sheet
'1.check if value in column 113 is not 0 - if it is 0 then skip the row, but if it is not 0 then continue
'2.check if value in column 125 is between 0.9 and 1.001 - if it is not then continue, but if it is then skip the row
'3.check if any of the cells in range 51-55 are not colored with color 36 or 39 - if they are not then continue, but if they are then skip the row
'4.check if any of the cells in range 108-112 are not colored with color 36 or 39 - if they are not then continue, but if they are then skip the row
'5.if all of the above are true then multiply the values in range 108-112 with the value in column 125 until the value in column 125 is between 0.9 and 1.001
'5.1 select cell 108
'5.2 resize the selection to 5 columns to 112
'5.3 multiply the values in the selection with the value in cell 125
'5.4 check if the value in cell 125 is between 0.9 and 1.001 - if not - copy again new value of cell 125 and repeat 5.3 and 5.4
'6.value in cell 125 is decreased by every multiplication, so after every multiplication cycle the value in column 125 is checked again and if it is not between 0.9 and 1.001 then the multiplication cycle is repeated
