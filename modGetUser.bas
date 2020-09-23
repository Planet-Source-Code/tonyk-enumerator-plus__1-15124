Attribute VB_Name = "modGetUser"
Option Explicit
Dim tUser As String
Declare Function w32_WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpszLocalName As String, ByVal lpszUserName As String, lpcchBuffer As Long) As Long

Public Function GetUserName() As String
    Dim lpUserName As String, lpnLength As Long, lResult As Long
    'Create a buffer
    lpUserName = String(256, Chr$(0))
    'Get the network user
    lResult = w32_WNetGetUser(vbNullString, lpUserName, 256)
    If lResult = 0 Then
        lpUserName = Left$(lpUserName, InStr(1, lpUserName, Chr$(0)) - 1)
        GetUserName = UCase(lpUserName)
    Else
        GetUserName = "No User Found !"
    End If
End Function
