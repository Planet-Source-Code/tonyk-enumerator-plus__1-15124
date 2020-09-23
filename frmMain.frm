VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enumerator Plus "
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMain.frx":0442
   ScaleHeight     =   6840
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save All Information "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5115
      TabIndex        =   20
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print All Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5115
      TabIndex        =   19
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   18
      Top             =   6430
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "User/Computer Information"
      ForeColor       =   &H00000080&
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   6855
      Begin VB.Label lblProxyStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Proxy is "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3600
         TabIndex        =   46
         Top             =   2805
         Width           =   3255
      End
      Begin VB.Label lblWinInstDir 
         BackStyle       =   0  'Transparent
         Caption         =   "Windows Install Directory"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   45
         Top             =   2400
         Width           =   3855
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Win. Install Dir:"
         Height          =   255
         Left            =   480
         TabIndex        =   44
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblWinSysDir 
         BackStyle       =   0  'Transparent
         Caption         =   "Windows System Directory"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   43
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Windows Sys. Dir:"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblProdKey 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Key Goes Here if there is one."
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   41
         Top             =   1920
         Width           =   3855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Key:"
         Height          =   255
         Left            =   680
         TabIndex        =   40
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblProdID 
         BackStyle       =   0  'Transparent
         Caption         =   "Product ID Goes Here"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   39
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Product ID:"
         Height          =   255
         Left            =   800
         TabIndex        =   38
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblSysRes 
         BackStyle       =   0  'Transparent
         Caption         =   "% of System Resources That Are Free"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   37
         Top             =   4005
         Width           =   4815
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "System Resources:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   4005
         Width           =   1695
      End
      Begin VB.Label lblIEVer 
         BackStyle       =   0  'Transparent
         Caption         =   "IE 4.x or Earlier"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   35
         Top             =   2805
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "I.E. Version:"
         Height          =   255
         Left            =   720
         TabIndex        =   34
         Top             =   2805
         Width           =   1095
      End
      Begin VB.Label lblScreenRes 
         BackStyle       =   0  'Transparent
         Caption         =   "1000 X 1000"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5280
         TabIndex        =   33
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Screen Resolution:"
         Height          =   255
         Left            =   3480
         TabIndex        =   32
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  'Transparent
         Caption         =   "UserName"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   31
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lblPCName 
         BackStyle       =   0  'Transparent
         Caption         =   "ComputerName"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   30
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblWinVer 
         BackStyle       =   0  'Transparent
         Caption         =   "Windows 2000 SP1"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   29
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label lblWinBuild 
         BackStyle       =   0  'Transparent
         Caption         =   "Windows Build"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   28
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblRegOrg 
         BackStyle       =   0  'Transparent
         Caption         =   "Registered To Organization"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   27
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label lblRegPers 
         BackStyle       =   0  'Transparent
         Caption         =   "Registered To Person(s)"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label lblMAC 
         BackStyle       =   0  'Transparent
         Caption         =   "NIC's Machine Address"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label lblIP 
         BackStyle       =   0  'Transparent
         Caption         =   "PC IP Address"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label lblUptime 
         BackStyle       =   0  'Transparent
         Caption         =   "Counting "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "PC Uptime:"
         Height          =   255
         Left            =   825
         TabIndex        =   10
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Machine Address:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "IP Address:"
         Height          =   255
         Left            =   780
         TabIndex        =   8
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Registered To:"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Windows Build:"
         Height          =   255
         Left            =   450
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Windows Version:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Name:"
         Height          =   255
         Left            =   375
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         Height          =   255
         Left            =   780
         TabIndex        =   1
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdGetCache 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Get  Passwords"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5115
      MouseIcon       =   "frmMain.frx":0884
      TabIndex        =   17
      Top             =   240
      Width           =   1815
   End
   Begin VB.ListBox List7 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1620
      ItemData        =   "frmMain.frx":0CC6
      Left            =   75
      List            =   "frmMain.frx":0CC8
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      Top             =   240
      Width           =   4905
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   4388
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   4388
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List3 
      Height          =   255
      Left            =   4388
      TabIndex        =   13
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List4 
      Height          =   255
      Left            =   4388
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List5 
      Height          =   255
      Left            =   4388
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List6 
      Height          =   255
      Left            =   4388
      TabIndex        =   16
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Left            =   5468
      Top             =   6360
   End
   Begin VB.ListBox lstParsedItems 
      Height          =   255
      Left            =   480
      TabIndex        =   47
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox lstTemp 
      Height          =   255
      Left            =   480
      TabIndex        =   48
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox lstProtocol 
      Height          =   255
      Left            =   480
      TabIndex        =   49
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox lstPorts 
      Height          =   255
      Left            =   480
      TabIndex        =   50
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox lstIP 
      Height          =   255
      Left            =   480
      TabIndex        =   51
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblOnOff 
      BackStyle       =   0  'Transparent
      Caption         =   "On"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4450
      TabIndex        =   53
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lblCacheStat 
      BackStyle       =   0  'Transparent
      Caption         =   "Password Caching is "
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2640
      TabIndex        =   52
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2228
      TabIndex        =   23
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "KlineSoft"
      BeginProperty Font 
         Name            =   "Pepita MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   2775
      TabIndex        =   21
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label lblCachePass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cached Passwords"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   195
      TabIndex        =   7
      Top             =   0
      Width           =   1620
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Visible         =   0   'False
      Begin VB.Menu mnuToggle 
         Caption         =   "&Toggle Proxy On/Off"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount& Lib "kernel32" ()
Dim pHold() As String
Dim Hold, Hold2, Hold3 As String
Dim dTmp, tTmp, lTmp, sTmp, lsTmp, fsTmp As String
Dim l1Hold, tOnOff, nl As String
Dim strText, h1, h2 As String
Dim strDelimeter, nType As String
Dim tRes, tType, tUname, tPass As String
Dim hRes, hType, hUname, hPass As String
Dim prxHold As Long
Dim CheckThisFile As String
Dim FileExists As Boolean

Private Sub cmdPrint_Click()
    Dim x As Integer
     'print routine here
    Printer.FontName = "Times New Roman"
    Printer.FontSize = 12
    Printer.FontBold = False         'no bold
    Printer.FontItalic = False       'no italic
    Printer.FontUnderline = False    'no underline
    Printer.FontStrikethru = False   'no strike
    Printer.ForeColor = RGB(0, 0, 0) 'color black
    Printer.FontBold = True
    Printer.FontUnderline = True
    'print a couple of blank lines for a top margin
    Printer.Print ""
    Printer.Print ""
    Printer.Print "Cached Passwords" & "  --- " & Date & "  " & Time
    Printer.Print ""
    Printer.FontBold = False
    Printer.FontUnderline = False
    'Next 3 lines just Print everything in list7
    For x = 0 To List7.ListCount - 1
       Printer.Print List7.List(x)
    Next x
    
    Printer.NewPage 'Starts a new page.
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.Print ""
    Printer.Print ""
    Printer.Print "System Information" & " --- " & Date & "  " & Time
    Printer.Print ""
    Printer.FontBold = False
    Printer.FontUnderline = False
    Printer.Print "User Name:   " & lblUser.Caption
    Printer.Print ""
    Printer.Print "Computer Name:   " & lblPCName.Caption
    Printer.Print ""
    Printer.Print "O/S Version:   " & lblWinVer.Caption
    Printer.Print ""
    Printer.Print "O/S Build:   " & lblWinBuild.Caption
    Printer.Print ""
    Printer.Print "Internet Explorer Version:   " & lblIEVer.Caption & "    " & lblProxyStatus.Caption
    Printer.Print ""
    Printer.Print "Windows Registered To:   " & lblRegPers.Caption
    Printer.Print ""
    Printer.Print "Windows Registered Organization:   " & lblRegOrg.Caption
    Printer.Print ""
    Printer.Print "Windows Product ID:   " & lblProdID.Caption
    Printer.Print ""
    Printer.Print "Windows Product Key:   " & lblProdKey.Caption
    Printer.Print ""
    Printer.Print "Windows System Directory Location:   " & lblWinSysDir.Caption
    Printer.Print ""
    Printer.Print "Windows Install Directory Location:   " & lblWinInstDir.Caption
    Printer.Print ""
    Printer.Print "PC's IP Address:   " & lblIP.Caption
    Printer.Print ""
    Printer.Print "NIC's Machine Address:   " & lblMAC.Caption
    Printer.Print ""
    Printer.Print "PC's Screen Resolution:   " & lblScreenRes.Caption
    Printer.Print ""
    Printer.Print "Available System Resources:   " & lblSysRes.Caption
    Printer.Print ""
    Printer.EndDoc
    MsgBox ("Now Prining The Information"), vbOKOnly
End Sub

Private Sub cmdSave_Click()
    Dim x As Integer
'This little routine simply appends data to a text file that is
'located in the same directory as the application.  It names it
'SysInfo.txt.
    Open App.Path & "\SysInfo.TXT" For Append As #1  'Open the file for append
    Write #1, Date & "   " & Time     'Date and Time Stamps the begining of the textfile
    For x = 0 To List7.ListCount - 1  'Loops through the listbox and writes entry to file until end of listbox
        If List7.List(x) = " " Then
            Write #1, "_____________________________________"  'Used to seperate Cache Entries
        Else
            Write #1, List7.List(x)
        End If
    Next x  'Loop back to the For x and get next item in list7
    Write #1, "User Name:   " & lblUser.Caption
    Write #1, "Computer Name:   " & lblPCName.Caption
    Write #1, "O/S Version:   " & lblWinVer.Caption
    Write #1, "O/S Build:   " & lblWinBuild.Caption
    Write #1, "Internet Explorer Version:   " & lblIEVer.Caption & "    " & lblProxyStatus.Caption
    Write #1, "Windows Registered To:   " & lblRegPers.Caption
    Write #1, "Windows Registered Organization:   " & lblRegOrg.Caption
    Write #1, "Windows Product ID:   " & lblProdID.Caption
    Write #1, "Windows Product Key:   " & lblProdKey.Caption
    Write #1, "Windows System Directory Location:   " & lblWinSysDir.Caption
    Write #1, "Windows Install Directory Location:   " & lblWinInstDir.Caption
    Write #1, "PC's IP Address:   " & lblIP.Caption
    Write #1, "NIC's Machine Address:   " & lblMAC.Caption
    Write #1, "PC's Screen Resolution:   " & lblScreenRes.Caption
    Write #1, "Available System Resources:   " & lblSysRes.Caption
    Close #1   'Close the file
    MsgBox ("Data Has Been Saved To " & App.Path & "\SysInfo.txt  "), vbOKOnly
End Sub

Private Sub Form_Load()
    Dim x
    Dim y
    
    On Error Resume Next
    Timer1.Interval = 1
    'The following 2 lines add the Program's Version info to the Window title.
    Me.Caption = Me.Caption & App.Major & "." & App.Minor & "." & _
    App.Revision
    
    lblUser.Caption = GetUserName 'call the GetUserName function from modGetUser
    lblPCName.Caption = GetPCName
    GetWinVer
    lblWinVer.Caption = wVer
    lblWinBuild.Caption = wBld
    
    ChkForDLL  'Calls sub ChkForDLL from modChkIEDll
    If DLLExists = True Then
        lblIEVer.Caption = GetIEVersion(lMajor, lMinor, lBuild) 'Calls GetIEVersion function from modGetIEVer
    End If
    
    GetRegInfo  'Calls public sub GetRegInfo from modRegInfo
    'Begin check to see if pasword cacheing has been disabled
    If caOnOff = 0 Then 'It's still enabled which is Windows default
        lblOnOff.ForeColor = &H8000& 'Green
        lblOnOff = "On"
    Else
        lblOnOff.ForeColor = &HC0&   'Red
        lblOnOff = "Off"  'Someone has disabled pass cache in registry
    End If
    'End check for disabled password cacheing
    
    lblRegPers.Caption = RegPer  'Shows Registered owner
    lblRegOrg.Caption = RegOrg   'Shows Registered Organization
    lblProdID = PID              'Shows Windows Product ID
    lblProdKey = pKey            'Shows Windows Product Key if there is one
    lblWinSysDir = WinSysDir     'Shows Windows System Directory
    lblWinInstDir = WinInstDir   'Shows Windows Install Directory (Where the CAB files are)
    
    ParseIt  'Calls the ParseIt sub in modProxy and parses proxy info.
    prxHold = phOnOff  'sets prxHold to either 1 for enabled or 0 for disabled
    If phOnOff = 1 Then  'Proxy is on
        tOnOff = "On"    'Sets label to ON
    Else
        tOnOff = "Off"  'Proxy is off
        strIP = "N/A"   'Set label to OFF
    End If
    
    lblProxyStatus = "Proxy is  " & tOnOff & "  at  " & strIP 'Displays proxy status and Address
    
    lblIP.Caption = GetIPAddress  'Calls GetIPAddress from modGetIP
    lblMAC.Caption = GetMACAddress() 'Calls GetMACAddress from modGetMAC
    
    'These 3 lines get and display the screen resolution
    x = Str$(GetSystemMetrics(SM_CXSCREEN))  'Screen Width
    y = Str$(GetSystemMetrics(SM_CYSCREEN))  'Screen Height
    lblScreenRes = x & "  X " & y
    
    'Next 4 lines get and display Available System Resources
    strFSR = "Free: " + CStr(GetFreeResources(GFSR_SYSTEMRESOURCES)) + "%"
    strFGDIR = "  GDI: " + CStr(GetFreeResources(GFSR_GDIRESOURCES)) + "%"
    strFUR = "  User: " + CStr(GetFreeResources(GFSR_USERRESOURCES)) + "%"
    lblSysRes = strFSR & strFGDIR & strFUR
    
     If Err Then
        MsgBox "Error: " & Err.Number & Chr(10) & Err.Description, vbOKOnly + vbCritical, "Error: " & Err.Number
     End If
     
End Sub

Private Sub cmdGetCache_Click()
    Dim i, B As Integer
    
    If lblOnOff = "Off" Then
        MsgBox ("Password Cacheing Has Been Turned Off"), vbOKOnly
        Exit Sub  'Exit cause there is nothing to process.
     End If
        
    Call GetPasswords  'calls the routine to get data into 2 hidden list boxes.
 
 'Begin code to parse username from hidden listbox 2 where they would look like
 'user:pass and splits it into user in listbox 3 and password in listbox 4.
 'it will also look for type 6 (dialup) from listbox 1 and parse the password
 'from the end of the string just after the last "\" and will add this to
 'listbox 3.  If there is no username then it wirtes "No Username" to listbox3.
 
    For i = 0 To List2.ListCount - 1
            Hold = List2.List(i)
            l1Hold = List1.List(i)
            pHold() = Split(Hold, ":", -1)
            dTmp = InStr(1, Hold, ":", vbBinaryCompare)
            tTmp = 0
            tTmp = Right(l1Hold, 1) 'Get the last charector in list1
            If tTmp = "6" Then  'if 6 then its a dialup type
                tTmp = 1  'temp string used like a switch
            Else
                tTmp = 0
            End If
            If dTmp = 0 Then
                List4.AddItem Hold
                
                If tTmp = 0 Then
                    List3.AddItem "No Username" 'Display this if no username found
                Else
                    strDelimeter = "\"
                    
                    If InStrRev(l1Hold, strDelimeter) > 0 Then
                         sTmp = Mid(l1Hold, InStrRev(l1Hold, strDelimeter) + 1)
                         lsTmp = Len(sTmp) - 3
                         fsTmp = Left(sTmp, lsTmp)
                         List3.AddItem fsTmp
                    End If
                 End If
         
            GoTo NoDel
            End If
            Hold2 = pHold(B)
            Hold2 = Left(Hold2, dTmp)
            List3.AddItem Hold2
            Hold3 = pHold(B)
            lTmp = Len(Hold) - dTmp
            Hold3 = Right(Hold, lTmp)
            List4.AddItem Hold3
        GoTo NoDel
NoDel:
    Next i
'End code to parse username:password from hidden listbox2 to hidden
'listboxes 3 and 4.

 'Begin code to parse the type (X.nType) from the 1st listbox.
    For i = 0 To List1.ListCount - 1
        sTmp = List1.List(i)
        If Right(sTmp, 1) = "3" Then nType = "3 (Share)"
        If Right(sTmp, 1) = "4" Then nType = "4 (MAPI)"
        If Right(sTmp, 1) = "6" Then nType = "6 (DialUP)"
        If Right(sTmp, 1) = "9" Then nType = "19 (Internet Explorer)"
        If Right(sTmp, 1) <> "9" Then
            lsTmp = Len(sTmp) - 3
        Else
            lsTmp = Len(sTmp) - 4
        End If
        fsTmp = Left(sTmp, lsTmp)
        List5.AddItem fsTmp
        List6.AddItem nType
    Next i
 'End parse type code.
    
 'Begin code to write all above gathered data to a textbox with readable
 'formatting.
    Dim eCount As String
    eCount = 1
    For i = 0 To List5.ListCount - 1
        nl = "Entry " & eCount  'Header and counter
        tRes = List5.List(i)    'Get RAW Resource Info
        h1 = Len(tRes)          'Count len of resource info
        h2 = h1 - 3             'Subtract 3 from length an put in h2
        tRes = Right(tRes, h2)  'Get stripped (finished) resource data
        hRes = " RESOURCE: "
        tType = List6.List(i)
        hType = "            TYPE:  "
        tUname = List3.List(i)
        hUname = " USERNAME:  "
        tPass = List4.List(i)
        hPass = "PASSWORD:  "
        'next 6 lines populate the final listbox (list7) with info gathered above.
        List7.AddItem nl
        List7.AddItem hRes & tRes
        List7.AddItem hType & tType
        List7.AddItem hUname & tUname
        List7.AddItem hPass & tPass
        List7.AddItem " "
        eCount = eCount + 1  'Just adds 1 to counter
    Next i  'Loops back to the For I until all items in list5 have been processed
End Sub

Private Sub cmdQuit_Click()
    QuitProg   'Calls QuitProg sub from modStartStop
End Sub

Private Sub GradientFill() ' Thanks to John Coleman for this gradient code
    'This just gives the form its cool color.
    Dim i As Long
    Dim c As Integer
    Dim r As Double
    r = ScaleHeight * 2.3
    If ScaleHeight = 0 Then GoTo errHand
    For i = 0 To ScaleHeight
        c = Abs(220 * Sin(i / r))
        Me.Line (0, i)-(ScaleWidth, i), RGB(c, c, c + 30)
    Next
errHand: Exit Sub
End Sub

Private Sub Form_Resize() 'This calls the above sub to give form its color.
    'GradientFill
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuToggle_Click()
'This menu item is hidden but does work.  The only problem is you
'have to restart the program to see the status switch so I am still
'working on this.
    If prxHold = 1 Then
        prxHold = 0
    ElseIf phOnOff = 0 Then
        prxHold = 1
    End If
    Call SetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyEnable", prxHold)
    'GetRegInfo
   
End Sub

Private Sub Timer1_Timer()
    Dim strHours As Long, strMinutes As Long, strSeconds As Long
    Dim strRaw As Long
    'The line below gets the tickcount and then divides
    'it by 1000 so you get the total whole number of seconds
    strRaw = GetTickCount \ 1000
    'Finds number of whole hours
    strHours = strRaw \ 3600
    'subtracts the number of whole hours
    strRaw = strRaw - (strHours * 3600)
    'finds the number of whole minutes
    strMinutes = strRaw \ 60
    'subtract the number of minutes
    strRaw = strRaw - (strMinutes * 60)
    'sets the last part as seconds
    strSeconds = strRaw
    '"Below" put a "0" in front of the number
    'if the number of seconds/minutes is below 10
    If strMinutes < 10 Then
        If strSeconds < 10 Then
            lblUptime.Caption = strHours & ":0" & strMinutes & ":0" & strSeconds
        Else
            lblUptime.Caption = strHours & ":0" & strMinutes & ":" & strSeconds
        End If
    Else
        If strSeconds < 10 Then
            lblUptime.Caption = strHours & ":" & strMinutes & ":0" & strSeconds
        Else
            lblUptime.Caption = strHours & ":" & strMinutes & ":" & strSeconds
        End If
    End If
End Sub


