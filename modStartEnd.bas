Attribute VB_Name = "modStartEnd"
Option Explicit

Public Sub main()
    frmMain.Show
End Sub

Public Sub QuitProg(Optional ByVal Force As Boolean = False)
    Dim I As Long
    On Error Resume Next
    For I = Forms.Count - 1 To 0 Step -1
        Unload Forms(I) ' Triggers QueryUnload and Form_Unload
         ' If we aren't in Force mode and the
         ' unload failed, stop the shutdown.
         If Not Force Then
            If Forms.Count > I Then
               Exit Sub
            End If
         End If
     Next I
      ' If we are in Force mode OR all
      ' forms unloaded, close all files.
     If Force Or (Forms.Count = 0) Then Close
      ' If we are in Force mode AND all
      ' forms not unloaded, end.
     If Force Or (Forms.Count > 0) Then End
End Sub

