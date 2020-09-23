Attribute VB_Name = "modRegistry"
'PUT THIS IN A .BAS!!!
'
' Easiest Read/Write to Registry
' Kevin Mackey
' LimpiBizkit@aol.com
'
Option Explicit

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1 ' Unicode nul terminated String
Public Const REG_DWORD = 4 ' 32-bit number
Public Const KEY_QUERY_VALUE = &H1&
Public Const KEY_SET_VALUE = &H2&
Public Const KEY_CREATE_SUB_KEY = &H4&
Public Const KEY_ENUMERATE_SUB_KEYS = &H8&
Public Const KEY_NOTIFY = &H10&
Public Const KEY_CREATE_LINK = &H20&
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Public Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Public Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Public Const KEY_EXECUTE = KEY_READ
Public Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Public Const DisplayErrorMsg = False

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long

Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte

Global hProxyAddress, pOnOff As String
Global dCount, d2Count, d3Count, d4Count As String
Global psText() As String
Global ps2Text() As String
Global ps3Text() As String
Global ps4Text() As String
Global cSep, eSep, scSep As String
Global sSep As String

Public Sub SaveKey(hKey As Long, strPath As String)
    Dim r As Long
    Dim keyhand&
    r = RegCreateKey(hKey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

Public Function GetString(hKey As Long, strPath As String, strValue As String)
    'EXAMPLE:
    '
    'text1.text = getstring(HKEY_CURRENT_USE
    '     R, "Software\VBW\Registry", "String")
    '
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim r As Long
    Dim lValueType
    
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)

        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))

            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub

Function GetDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    'EXAMPLE:
    '
    'text1.text = getdword(HKEY_CURRENT_USER
    '     , "Software\VBW\Registry", "Dword")
    '
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    ' Get length/data type
    'GetDword = 1
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDword = lBuf
        End If
   ' Else
       ' GetDword = 1
        'Call errlog("GetDWORD-" & strPath, Fals
        '     e)
    End If
    r = RegCloseKey(keyhand)
End Function

Function GetDWORDValue(SubKey As String, Entry As String) As Variant
    Call ParseKey(SubKey, MainKeyHandle)
    If MainKeyHandle Then
       rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
       If rtn = ERROR_SUCCESS Then 'if the key could be opened then
          rtn = RegQueryValueExA(hKey, Entry, 0, REG_DWORD, lBuffer, 4) 'get the value from the registry
          If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
             rtn = RegCloseKey(hKey)  'close the key
             GetDWORDValue = lBuffer  'return the value
          Else                        'otherwise, if the value couldnt be retreived
             GetDWORDValue = "Error"  'return Error to the user
             If DisplayErrorMsg = True Then 'if the user wants errors displayed
                MsgBox ErrorMsg(rtn)        'tell the user what was wrong
             End If
          End If
       Else 'otherwise, if the key couldnt be opened
          GetDWORDValue = "Error"        'return Error to the user
          If DisplayErrorMsg = True Then 'if the user wants errors displayed
             MsgBox ErrorMsg(rtn)        'tell the user what was wrong
          End If
       End If
    End If
End Function

Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    'EXAMPLE"
    '
    'Call SaveDword(HKEY_CURRENT_USER, "Soft
    '     ware\VBW\Registry", "Dword", text1.text)
    '
    '
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then
    '     Call errlog("SetDWORD", False)
    r = RegCloseKey(keyhand)
End Function

Function SetDWORDValue(SubKey As String, Entry As String, Value As Long)
    Call ParseKey(SubKey, MainKeyHandle)
    If MainKeyHandle Then
       rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_WRITE, hKey) 'open the key
       If rtn = ERROR_SUCCESS Then
          rtn = RegSetValueExA(hKey, Entry, 0, REG_DWORD, Value, 4)
          If Not rtn = ERROR_SUCCESS Then
             If DisplayErrorMsg = True Then
                MsgBox ErrorMsg(rtn)
             End If
          End If
          rtn = RegCloseKey(hKey) 'close the key
       Else 'if there was an error opening the key
          If DisplayErrorMsg = True Then 'if the user want errors displayed
             MsgBox ErrorMsg(rtn) 'display the error
          End If
       End If
    End If
End Function

Public Function DeleteKey(ByVal hKey As Long, ByVal strKey As String)
    'EXAMPLE:
    '
    'Call DeleteKey(HKEY_CURRENT_USER, "Soft
    '     ware\VBW")
    '
    Dim r As Long
    r = RegDeleteKey(hKey, strKey)
End Function

Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "So
    '     ftware\VBW\Registry", "Dword")
    '
    Dim keyhand As Long
    Dim r As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

Private Sub ParseKey(KeyName As String, Keyhandle As Long)
    rtn = InStr(KeyName, "\") 'return if "\" is contained in the Keyname
    If Left(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
       MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + KeyName 'display error to the user
       Exit Sub 'exit the procedure
    ElseIf rtn = 0 Then 'if the Keyname contains no "\"
       Keyhandle = GetMainKeyHandle(KeyName)
       KeyName = "" 'leave Keyname blank
    Else 'otherwise, Keyname contains "\"
       Keyhandle = GetMainKeyHandle(Left(KeyName, rtn - 1)) 'seperate the Keyname
       KeyName = Right(KeyName, Len(KeyName) - rtn)
    End If
End Sub

Function ErrorMsg(lErrorCode As Long) As String
    Dim GetErrorMsg
    'If an error does accurr, and the user wants error messages displayed, then
    'display one of the following error messages
    Select Case lErrorCode
           Case 1009, 1015
                GetErrorMsg = "The Registry Database is corrupt!"
           Case 2, 1010
                GetErrorMsg = "Bad Key Name"
           Case 1011
                GetErrorMsg = "Can't Open Key"
           Case 4, 1012
                GetErrorMsg = "Can't Read Key"
           Case 5
                GetErrorMsg = "Access to this key is denied"
           Case 1013
                GetErrorMsg = "Can't Write Key"
           Case 8, 14
                GetErrorMsg = "Out of memory"
           Case 87
                GetErrorMsg = "Invalid Parameter"
           Case 234
                GetErrorMsg = "There is more data than the buffer has been allocated to hold."
           Case Else
                GetErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
    End Select
End Function

Function GetMainKeyHandle(MainKeyName As String) As Long
    Select Case MainKeyName
           Case "HKEY_CLASSES_ROOT"
                GetMainKeyHandle = HKEY_CLASSES_ROOT
           Case "HKEY_CURRENT_USER"
                GetMainKeyHandle = HKEY_CURRENT_USER
           Case "HKEY_LOCAL_MACHINE"
                GetMainKeyHandle = HKEY_LOCAL_MACHINE
           Case "HKEY_USERS"
                GetMainKeyHandle = HKEY_USERS
           Case "HKEY_PERFORMANCE_DATA"
                GetMainKeyHandle = HKEY_PERFORMANCE_DATA
           Case "HKEY_CURRENT_CONFIG"
                GetMainKeyHandle = HKEY_CURRENT_CONFIG
           Case "HKEY_DYN_DATA"
                GetMainKeyHandle = HKEY_DYN_DATA
    End Select
End Function

