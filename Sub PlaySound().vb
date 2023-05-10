Sub PlaySound()
    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    'increase the volume to 70%
    oShell.SendKeys "^{F8}"
    oShell.Run "C:\Windows\Media\ding.wav", 1, False
End Sub
