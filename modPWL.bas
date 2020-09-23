Attribute VB_Name = "modPWL"
Option Explicit

Declare Function WNetEnumCachedPasswords Lib "mpr.dll" (ByVal s As String, ByVal i As Integer, ByVal B As Byte, ByVal proc As Long, ByVal l As Long) As Long
Type PASSWORD_CACHE_ENTRY
    cbEntry As Integer
    cbResource As Integer
    cbPassword As Integer
    iEntry As Byte
    nType As Byte
    abResource(1 To 1024) As Byte
    End Type

Public Function CallBack(x As PASSWORD_CACHE_ENTRY, ByVal lSomething As Long) As Integer
    Dim nLoop As Integer
    Dim cString As String
    Dim Resource As String
    Dim ResType As String
    Dim Password As String
    ResType = x.nType

    For nLoop = 1 To x.cbResource
        If x.abResource(nLoop) <> 0 Then
            cString = cString & Chr(x.abResource(nLoop))
        Else
            cString = cString & " "
        End If
    Next

    Resource = cString
    cString = ""

    For nLoop = x.cbResource + 1 To (x.cbResource + x.cbPassword)
        If x.abResource(nLoop) <> 0 Then
            cString = cString & Chr(x.abResource(nLoop))
        Else
            cString = cString & " "
        End If
    Next
    
    Password = cString
    cString = ""
    frmMain.List1.AddItem " R: " & Resource & " T:" & x.nType
    frmMain.List2.AddItem Password
    CallBack = True
End Function

Public Sub GetPasswords()
    Dim nLoop As Integer
    Dim cString As String
    Dim lLong As Long
    Dim bByte As Byte
    bByte = &HFF
    nLoop = 0
    lLong = 0
    cString = ""
    Call WNetEnumCachedPasswords(cString, nLoop, bByte, AddressOf CallBack, lLong)
End Sub


