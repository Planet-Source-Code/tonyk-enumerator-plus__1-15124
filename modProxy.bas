Attribute VB_Name = "modProxy"
Option Explicit

Dim A$, z As Integer, P$, k As Integer, B$
Global strIP As String

Public Sub ParseIt()
    Dim i As Integer
    Dim x As Integer
    Dim y As Integer
    If phOnOff = 0 Then
        Exit Sub
    End If
    hProxyAddress = iProxy
    psText() = Split(hProxyAddress, ";", -1)
    dCount = UBound(psText)
    For i = 0 To dCount
        frmMain.lstParsedItems.AddItem psText(i)
        ps2Text() = Split(psText(i), "=", -1)
        d2Count = UBound(ps2Text)
        For x = 0 To d2Count - 1
            frmMain.lstProtocol.AddItem ps2Text(x)
            ps3Text() = Split(psText(i), ":", -1)
            d3Count = UBound(ps3Text)
            For y = 1 To d3Count
               frmMain.lstPorts.AddItem ps3Text(y)
            Next y
        Next x
    Next i
    cSep = ":"
    eSep = "="
    scSep = ";"
    sSep = cSep & eSep
    
    P$ = sSep
    A$ = hProxyAddress
    For z = 2 To Len(P$)
        k = InStr(A$, Mid$(P$, z, 1))
        Do While k
            Mid$(A$, k, 1) = Left$(P$, 1)
            k = InStr(A$, Mid$(P$, z, 1))
        Loop
    Next
    Do While Len(A$)
        k = InStr(A$, Left$(P$, 1))
        If k = 1 Then
            A$ = Mid$(A$, 2)
        ElseIf k Then
            B$ = Left$(A$, k - 1)
            'lblIP = B$
            frmMain.lstTemp.AddItem B$
            
            A$ = Mid$(A$, k + 1)
        Else
            A$ = ""
        End If
    Loop
    
    frmMain.lstIP.AddItem B$
    strIP = B$
End Sub
