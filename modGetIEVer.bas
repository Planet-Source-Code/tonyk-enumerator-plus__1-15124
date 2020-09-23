Attribute VB_Name = "modGetIEVer"
Option Explicit

Private Declare Function DllGetVersion Lib "Shlwapi.dll" _
        (pdvi As DLLVERSIONINFO) As Long

Private Const NOERROR = 0

Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
End Type

Global rtnStr As String
Global lMajor As Long, lMinor As Long, lBuild As Long

Public Function GetIEVersion(ByRef lMajor As Long, _
                             ByRef lMinor As Long, _
                             Optional ByRef lBuild As Long) As String
    
    Dim tDLLVerInfo As DLLVERSIONINFO
    Dim r As Long
    On Error Resume Next
    ' Reset version info
    lMajor = 0: lMinor = 0: lBuild = 0
    
    ' Set the cbSize of the DLLVERSIONINFO
    ' structure as this needs to be filled
    ' before calling the version info function
    tDLLVerInfo.cbSize = Len(tDLLVerInfo)
    
    ' Call the function that will return the
    ' version info
    r = DllGetVersion(tDLLVerInfo)
    
    If r = NOERROR Then
        ' Return a string and values. First the values
        With tDLLVerInfo
            lMajor = .dwMajor
            lMinor = .dwMinor
            lBuild = .dwBuildNumber
        End With
        ' ...and the string
        GetIEVersion = lMajor & "." & lMinor & "." & lBuild
    Else
        ' There was an error.. Might be because
        ' IE isn't installed.
        GetIEVersion = "ERROR"
    End If
    
End Function


