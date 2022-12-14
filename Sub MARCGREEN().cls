Sub MARCGREEN()
    '
    ' MARCGREEN Macro
    '

    '
    Selection.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434828
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    'SPOKOINO ZELENO
    'Selection.Interior.ColorIndex = 35

    'PREDVARITELNO JALTO
    'Selection.Interior.ColorIndex = 36

    'NASITENO ZELENO
    'Selection.Interior.ColorIndex = 4
End Sub
