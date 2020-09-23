Attribute VB_Name = "modGetWinVer"
Option Explicit

Public Declare Function GetVersionEx Lib "kernel32" Alias _
       "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 '  Maintenance string for PSS usage
End Type

' dwPlatforID Constants
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Dim tOSVer As OSVERSIONINFO
Global wVer, wBld As String

Public Function GetWinVer() As String
   ' First set length of OSVERSIONINFO
   ' structure size
   tOSVer.dwOSVersionInfoSize = Len(tOSVer)
   ' Get version information
   GetVersionEx tOSVer
   ' Determine OS type
   With tOSVer
      
      Select Case .dwPlatformId
         Case VER_PLATFORM_WIN32_NT
            ' This is an NT version (NT/2000)
            ' If dwMajorVersion >= 5 then
            ' the OS is Win2000
            If .dwMajorVersion >= 5 Then
               wVer = "Windows 2000"
            Else
               wVer = "Windows NT"
            End If
         Case Else
            ' This is Windows 95/98/ME
            If .dwMajorVersion >= 5 Then
               wVer = "Windows ME"
            ElseIf .dwMajorVersion = 4 And .dwMinorVersion > 0 Then
               wVer = "Windows 98"
            Else
               wVer = "Windows 95"
            End If
        ' End Select
         ' Check for service pack
         
          If InStr(.szCSDVersion, "C") Then
             wVer = wVer & "OSR2"
          Else
             wVer = wVer
          End If
                
               ' Case Is = 10

          If InStr(.szCSDVersion, "A") Then
             wVer = wVer & " SE"
          Else
             wVer = wVer
          End If
                
               ' Case Is = 90
         wVer = wVer
         End Select
         'wVer = wVer & " " & Left(.szCSDVersion, _
                         ' InStr(1, .szCSDVersion, Chr$(0)))
                          
         ' Get OS version
         wBld = .dwMajorVersion & "." & _
                          .dwMinorVersion & "." & .dwBuildNumber
        
    End With

End Function


