Attribute VB_Name = "modChkIEDll"
Option Explicit

Global CheckThisDLL As String
Global DLLExists As Boolean

Public Sub ChkForDLL()
'Ok here we check for the shlwapi.dll which is not on systems with
'versions of IE 4.x and lower.  This will prevent the error "cannot
'find file shlwapi.dll" when run from computers with lower versions of IE.

    CheckThisDLL = Dir("C:\Windows\System\shlwapi.dll")
    If CheckThisDLL = "" Then
        DLLExists = False
    Else
        DLLExists = True
    End If
End Sub

