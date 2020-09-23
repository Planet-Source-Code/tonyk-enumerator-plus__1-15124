Attribute VB_Name = "modToggleProxy"
Option Explicit

Public Sub ToggleProxy()
    Shell ("start rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,3"), vbHide
    Sleep (500)
    SendKeys ("L")
    Sleep (700)
    SendKeys ("{TAB}")
    SendKeys ("{TAB}")
    SendKeys ("{x}")
    SendKeys ("{ENTER}")
    Sleep (700)
    SendKeys ("{TAB}")
    SendKeys ("{TAB}")
    SendKeys ("{TAB}")
    SendKeys ("{TAB}")
    SendKeys ("{TAB}")
    SendKeys ("{TAB}")
    SendKeys ("{TAB}")
    SendKeys ("{ENTER}")
    
    'End
End Sub

Sub Sleep(ByVal MillaSec As Long, Optional ByVal DeepSleep As Boolean = False)
    Dim tStart#, Tmr#
    tStart = Timer
    While Tmr < (MillaSec / 1000)
        Tmr = Timer - tStart
        If DeepSleep = False Then DoEvents
    Wend
End Sub
