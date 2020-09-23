Attribute VB_Name = "modGetPCName"
Option Explicit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function GetPCName() As String
    Dim pcnStr As String
    'Create a buffer
    pcnStr = String(255, Chr$(0))
    'Get the computer name
    GetComputerName pcnStr, 255
    'remove the unnecessary chr$(0)'s
    pcnStr = Left$(pcnStr, InStr(1, pcnStr, Chr$(0)) - 1)
    'Show the computer name
   GetPCName = pcnStr
End Function
  
