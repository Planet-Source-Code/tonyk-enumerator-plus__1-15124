Attribute VB_Name = "modSysResource"
Option Explicit

Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2

Declare Function GetFreeResources Lib "RSRC32" Alias "_MyGetFreeSystemResources32@4" (ByVal lWhat As Long) As Long

Global strFSR, strFGDIR, strFUR As String

