Attribute VB_Name = "modRegInfo"
Option Explicit

Global pKey, RegPer, RegOrg, PID As String
Global WinSysDir, WinInstDir, iProxy, phOnOff As String
Global caOnOff As String

Public Sub GetRegInfo()
    RegPer = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
    RegOrg = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization")
    PID = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProductId")
    pKey = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProductKey")
    WinSysDir = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Setup", "SysDir")
    WinInstDir = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Setup", "SourcePath")
    iProxy = GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyServer")
    phOnOff = GetDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyEnable")
    caOnOff = GetDword(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Network", "DisablePwdCaching")
End Sub
