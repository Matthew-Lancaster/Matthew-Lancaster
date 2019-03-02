VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form CID_Run_Me 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CID Run Me"
   ClientHeight    =   1725
   ClientLeft      =   21540
   ClientTop       =   16665
   ClientWidth     =   5895
   Icon            =   "frmCentral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   115
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   WindowState     =   1  'Minimized
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Arh So "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   6030
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2085
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Timer Timer6 
      Interval        =   25
      Left            =   3375
      Top             =   1065
   End
   Begin VB.ListBox List2 
      Height          =   4155
      Left            =   6030
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   255
      Width           =   2910
   End
   Begin RichTextLib.RichTextBox RTB1 
      CausesValidation=   0   'False
      Height          =   3225
      Left            =   -15
      TabIndex        =   14
      Top             =   1710
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   5689
      _Version        =   393217
      BackColor       =   13857626
      Enabled         =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      OLEDropMode     =   0
      TextRTF         =   $"frmCentral.frx":0E42
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer5 
      Interval        =   10000
      Left            =   2955
      Top             =   1065
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   2955
      Top             =   645
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3375
      Top             =   645
   End
   Begin VB.ListBox List1 
      BackColor       =   &H0080FFFF&
      Height          =   3180
      Left            =   -15
      TabIndex        =   13
      Top             =   4980
      Visible         =   0   'False
      Width           =   5925
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   3375
      Top             =   225
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Monitor"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1320
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   -15
      Width           =   1320
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   30
      TabIndex        =   7
      Top             =   4485
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2955
      Top             =   225
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Music."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   0
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   -15
      Width           =   1320
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H007D9C34&
      Caption         =   "111"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   2640
      TabIndex        =   8
      Top             =   0
      Width           =   2025
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00DF705B&
      Caption         =   " - Known -   Processes ...I Like..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1020
      Left            =   4665
      TabIndex        =   12
      Top             =   0
      Width           =   1230
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00DF705B&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   690
      Left            =   4665
      TabIndex        =   11
      Top             =   1020
      Width           =   1230
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H007B45FA&
      Caption         =   "Monitor On"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   1305
      TabIndex        =   10
      Top             =   375
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404040&
      Caption         =   "Keyboard"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   2355
      TabIndex        =   6
      Top             =   1320
      Width           =   1290
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404040&
      Caption         =   "Mouse"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H007D9C34&
      Caption         =   "Key"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   3630
      TabIndex        =   4
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H007D9C34&
      Caption         =   "Mouse"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   390
      Left            =   960
      TabIndex        =   3
      Top             =   1320
      Width           =   1395
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H007B5B37&
      Caption         =   "6/6/06 6:6:6P"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   585
      Left            =   -15
      TabIndex        =   2
      Top             =   735
      Width           =   4680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H007B45FA&
      Caption         =   " Music On"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   1305
   End
   Begin VB.Menu Menu_File 
      Caption         =   "&File"
      Begin VB.Menu Menu_NoteGoldWinamp 
         Caption         =   "Gold Winamp Log - NotePad"
      End
      Begin VB.Menu Menu_NoteCallWinamp 
         Caption         =   "Caller Winamp Log - NotePad"
      End
   End
   Begin VB.Menu Menu_Execute 
      Caption         =   "&Execute"
      Begin VB.Menu Menu_Pulse 
         Caption         =   "Pulse"
      End
      Begin VB.Menu Menu_Date1991 
         Caption         =   "Date1991"
      End
      Begin VB.Menu Menu_Mp3Tags 
         Caption         =   "MP3 Tags"
      End
      Begin VB.Menu Menu_CreativePlayCenter 
         Caption         =   "Creative PlayCenter"
      End
      Begin VB.Menu Menu_AttachmentExtractor 
         Caption         =   "Attachment Extractor"
      End
      Begin VB.Menu Menu_Alcohol120 
         Caption         =   "Alcohol 120%"
      End
   End
   Begin VB.Menu Menu_Menu 
      Caption         =   "&Menu"
      Begin VB.Menu Menu_KeyboardM 
         Caption         =   "Keyboard Menu"
      End
      Begin VB.Menu Menu_KnowProcess 
         Caption         =   "Known Processes I Like"
      End
      Begin VB.Menu Menu_StopCurrent 
         Caption         =   "Stop After Current"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu Menu_Help 
      Caption         =   "&Help"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu Menu_HelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "CID_Run_Me"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd _
        As Long, lpPoint As POINTAPI)
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function RasEnumEntriesA Lib "rasapi32.dll" _
    (ByVal Reserved As String, ByVal lpszPhonebook As String, _
    lprasentryname As Any, lpcb As Long, lpcEntries As Long) _
    As Long



Private Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpcb As Long, lpcConnections As Long) As Long





Private Const sLocation As String = "CID"

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
'SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_ON
Private Const MONITOR_ON = -1&
Private Const MONITOR_OFF = 2&
Private Const SC_MONITORPOWER = &HF170&
Private Const WM_SYSCOMMAND = &H112
Private Const ipc_isplaying As Long = 104

Private Const WINAMP_BUTTON2 = 40045
Private Const WINAMP_BUTTON3 = 40046
Private Const WM_CLOSE = &H10
Private Const WM_USER = &H400
Private Const WM_COMMAND = &H111
Private Const GW_HWNDNEXT = 2

Private Type RASENTRYNAME95
    dwSize As Long
    szEntryName(256) As Byte
End Type


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type



Public Mousey As Long
Public Keyy As Long
Public Killtime As Long

Public Sluty As Long
Public Sluty2 As Long

Public Tibulate As Double

Public Startuptime As Double
Public Postimer As Date

Dim Music_Off_Timer As Date
Dim Monitor2_Off_Timer As Date
Public Skip2 As Long

Public ScrWidth_Old As Double

Public HeightOffSet As Long
Public Ew2 As Date

Dim MsgResult As Long
Dim XcNul As Long
Dim LhRet As Long

Public Atg$
Public Atg2$

Public CurProcHwnd2 As Long
Public TameBeast2 As Long

Public Chink55 As Long

Public ListBox55

Public Qwerty2 As Long

Public Mousey3 As Long
Public Mouse55 As Long
Public KeyBack

Public TogW As Long
Public TogW2 As Long

Public Xmp1 As Long
Public Ymp1 As Long

Public Comstr As Long
Public Wq5 As Long
Public Now4 As Date
Public Bdf3 As Long

Public Xmarks As Date

Public Sps As Long
Public Spd As Long

Public FormLoad1 As Long

Public AgesEmpire As Long

Public FirstRun As Long

Public MattsTimer5 As Date

Public AgPop As Date

Public Dbon As Double
Public DbonPop As Long

Public MajorOnTime As Date
Public MinorOnTime As Date

Public Rem5 As Long

'Public NoMonOff As long
'Public NoMusic As long
Public NoMonOffX As Long
Public NoMusicX As Long

Public Dbon2 As Long

Public DayCounter As Long
Public DayCheck As Date
Public Daymouse As Long
Public Daykeyy As Long

Public Gabh As Long
Public Gabh2 As Long
Public Gabh3 As Long
Public GabhTimer As Date
Public GabhTimer2 As Date

Public DayMouseTimer As Date

Public TotalDayMouse As Long
Public TotalDayKeyy As Long

Public AsusTimer20 As Date

Public MattsTimer5Old As Date
Public MattsTimer2Old As Date
Public PeaShoot As Double
Public XRad As Long
Public OldXRad As Long
Public OldWState As Long

Public WLFN1$
Public WLFN2$

Public Timer5Div As Date

Private Type FILETIME
   LowDateTime          As Long
   HighDateTime         As Long
End Type


Private Type WIN32_FIND_DATA
   dwFileAttributes     As Long
   ftCreationTime       As FILETIME
   ftLastAccessTime     As FILETIME
   ftLastWriteTime      As FILETIME
   nFileSizeHigh        As Long
   nFileSizeLow         As Long
   dwReserved0          As Long
   dwReserved1          As Long
   cFileName            As String * 260  'MUST be set to 260
   cAlternate           As String * 14
End Type


Public FirstTime5Run As Long

Public LastCurProcHwnd As Long


Public FirstRun2 As Long


Public NoMonMusChange As Long

Public PlimBurn1 As Long
Public PlimBurn2 As Long

Public tagad2perm$
Public wedf$
Public wedg$
Public Chit22 As Long
Public Chit24 As Long

Public ReRun As Long

Public LockInToMp3 As Long
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------

Public QueTheAmpTest As Date

Public AutoConVpn As Date

Public OldAsh$
Public OldAshPI As Integer

Public Agust As Date ' timer for sub bashing for winamp stopper if diff album
Public Agust2 As Date

Public XtX1
Public XtX2

Public Label10BackColor
Public Label21BackColor
Public Label23BackColor
Public InitHeight
Public EmSQ


Public Sub Init()


NoMonOff2 = 0
NoMonOff = 0
NoMusic = 0
ProcessId24 = 0
ProcessId22 = 0
frontpagepid = 0
ProcessId25Ssam = 0

mdlProcess.List_ActiveProcess
mdlProcess.List_ActiveModules

End Sub


Private Sub Command2_Click()

If Command2.BackColor = QBColor(12) Then
    Command2.BackColor = QBColor(10)
    Menu_StopCurrent.Checked = True
    Exit Sub
End If

If Command2.BackColor = QBColor(10) Then
    Command2.BackColor = QBColor(12)
    Menu_StopCurrent.Checked = False
End If

End Sub




Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Call Label6Hit

End Sub

Sub Label6Hit()

If OptionsMenu.Top <> 0 Then
    OptionsMenu.WindowState = 0
    OptionsMenu.Refresh
End If

If OptionsMenu.WindowState <> 0 Then
    OptionsMenu.Width = CID_Run_Me.Width
    OptionsMenu.Top = CID_Run_Me.Top - OptionsMenu.Height
    OptionsMenu.Left = Screen.Width - OptionsMenu.Width - OffSetGoogle
End If

If OptionsMenu.Visible = False Then
    OptionsMenu.Show
    OptionsMenu.List1.SetFocus
    Exit Sub
End If

If OptionsMenu.Visible = True Then
    OptionsMenu.Visible = False
End If

End Sub

Private Sub Command1_Click()
If Music_Off_Timer > 0 Then Amon = 1
If Amon = 0 Then
    Music_Off_Timer = 0
    winamp_player_off
    Music_Off_Timer = Now + TimeSerial(12, 0, 0)
End If
If Amon = 1 Then Music_Off_Timer = 0: Amon = 0

End Sub

Private Sub Command3_Click()
If Monitor2_Off_Timer > 0 Then Amon = 1
Monitor2_Off_Timer = Now + TimeSerial(12, 0, 0)
If Amon = 1 Then Monitor2_Off_Timer = 0: Amon = 0

End Sub


Public Sub Form_Load()

If App.PrevInstance = True Then
    
    Cheese = 3
    
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    
    Exit Sub

End If



MMControl1.Notify = True
MMControl1.Wait = True
MMControl1.Shareable = False
MMControl1.DeviceType = "WaveAudio"
MMControl1.fileName = App.Path + "\005.Wav"

'Open the MCI WaveAudio device.
MMControl1.Command = "Open"

'MMControl1.Command = "Prev"
'MMControl1.Command = "Play"






Dim Hwnd23
Dim HwndCTLTask3
Dim Rect23 As RECT
HwndCTLTask3 = FindWindow("_GD_Sidebar", vbNullString)
If HwndCTLTask3 > 0 Then
    Hwnd23 = GetWindowRect(HwndCTLTask3, Rect23)
End If
If Rect23.Top = 0 Then OffSetGoogle = 0
If Rect23.Top > 0 Then OffSetGoogle = (Rect23.Right - Rect23.Left) * Screen.TwipsPerPixelY


Monitor2_Off_Timer = Now + TimeSerial(24, 0, 0)
Music_Off_Timer = Now + TimeSerial(24, 0, 0)

Agust = Now

AsusTimer20 = Now + TimeSerial(0, 0, 30)

MattsTimer5 = 0

Cheese = True

Load OptionsMenu

MajorOnTime = TimeSerial(0, 5, 0)
MinorOnTime = TimeSerial(0, 2, 0)

FirstRun = True
        
FormLoad1 = 1

ReDim HwndArray(200)
ReDim HwndArra2(200)
ReDim GetWinLen(200)

For R = 0 To 200
    HwndArray(R) = -2
    HwndArra2(R) = -2
Next




DUN_Services sArray


If Dir$(App.Path + "\AutoVPN.txt") <> "" Then
    Open App.Path + "\AutoVPN.txt" For Input As #1
    Line Input #1, egg$
    Close #1
    AutoConVpn = DateValue(egg$) + TimeValue(egg$)
End If




CID_Run_Me.Caption = "CID Run Me - Ver " + Trim$(str$(App.Major)) + "." + Trim$(str$(App.Minor)) + "." + Trim$(str$(App.Revision))

Startuptime = Now
Postimer = Now + TimeSerial(0, 1, 30)
CID_Run_Me.WindowState = 0
Tibulate = Now + TimeSerial(0, 1, 0)


CID_Run_Me.Height = (Label5.Top + Label5.Height + 39) * Screen.TwipsPerPixelY

InitHeight = CID_Run_Me.Height


CID_Run_Me.Left = Screen.Width - CID_Run_Me.Width - OffSetGoogle
HeightOffSet = 400
CID_Run_Me.Top = Screen.Height - (CID_Run_Me.Height + HeightOffSet)

'MsgBox (str$(Screen.Height) + vbCr + str$(CID_Run_Me.Height))
CID_Run_Me.WindowState = 0

CID_Run_Me.Refresh

OptionsMenu.Width = CID_Run_Me.Width
OptionsMenu.Top = CID_Run_Me.Top - OptionsMenu.Height
OptionsMenu.Left = Screen.Width - OptionsMenu.Width - OffSetGoogle


If Dir$(App.Path + "\winamp-list.txt") <> "" Then
    freef5 = FreeFile
    Open App.Path + "\winamp-list.txt" For Input As #freef5
    Do
        Line Input #freef5, Atg$
    Loop Until EOF(freef5)
    Close #freef5
    Atg$ = Mid$(Atg$, 24)

End If



QueTheAmpTest = Now + TimeSerial(0, 0, 20)


Call Init

Load frmReceiver
Load frmSender_CallID

If vbCompiled2 = True Then
    Load DSkeybd_F
    MajorOnTime = TimeSerial(0, 5, 0)
    MinorOnTime = TimeSerial(0, 2, 0)

Else
    Timer4.Enabled = True
    
    MajorOnTime = TimeSerial(0, 1, 0)
    MinorOnTime = TimeSerial(0, 0, 20)

End If






free5 = FreeFile
Open "C:\Callerid\ci logs\2cidrunme.txt" For Input As #free5
Do
    Line Input #free5, aed$
    If InStr(aed$, "Last System ReBoot Time") Then
        wett = 1
        Qwerty2 = 0
        Mouse55 = 0
    End If
    
    If wett = 1 Then
        If InStr(aed$, "Keyboard =") Then
            Qwerty2 = Qwerty2 + Val(Mid$(aed$, InStr(aed$, "Keyboard =") + 10))
        End If
        If InStr(aed$, "Mouse =") Then
            Mouse55 = Mouse55 + Val(Mid$(aed$, InStr(aed$, "Mouse =") + 7, 11))
        End If
    End If

Loop Until EOF(free5)

Close free5

If Dir$("C:\Callerid\ci logs\2cidrunmedaycount.txt") <> "" Then
    free5 = FreeFile
    Open "C:\Callerid\ci logs\2cidrunmedaycount.txt" For Input As #free5
    
    Do
        Line Input #free5, aed$
        If InStr(aed$, "Day Counter") Then
            DayCheck = DateValue(Mid$(aed$, 15))
            Line Input #free5, aed$
            Daymouse = Val(Mid$(aed$, 15))
            Line Input #free5, aed$
            Daykeyy = Val(Mid$(aed$, 15))
        End If
    Loop Until EOF(free5)
End If

If DayCheck = 0 Then DayCheck = Now

Close #free5

If Dir$(App.Path + "\WinAmpLog\Winamp-list-CallerID.txt") <> "" Then
    freef5 = FreeFile
    Open App.Path + "\WinAmpLog\Winamp-list-CallerID.txt" For Input As #freef5
    Do
        Line Input #freef5, Atg$
    Loop Until EOF(freef5)
    Close #freef5
    Atg$ = Trim(Mid$(Atg$, 23))
End If

If Dir$(App.Path + "\WinAmpLog\Winamp-list-GoldAmp.txt") <> "" Then
    freef5 = FreeFile
    Open App.Path + "\WinAmpLog\Winamp-list-GoldAmp.txt" For Input As #freef5
    Do
        Line Input #freef5, Atg2$
    Loop Until EOF(freef5)
    Close #freef5
    Atg2$ = Trim(Mid$(Atg2$, 23))
End If


frmSender_CallID.txtMsg.Text = "KeyCounter " + str$(Qwerty2 + Keyy)
Call frmSender_CallID.cmdSend_Click

frmSender_CallID.txtMsg.Text = "MouseCounter " + str$(Mouse55 + Mousey)
Call frmSender_CallID.cmdSend_Click


FirstRun = False

Call proc24

frmSender_CallID.txtMsg.Text = "Request Play " + Time$
Call frmSender_CallID.cmdSend_Click

If CheckT1$ = "2" Then
    Call ClingDing
End If

Label10BackColor = Label10.BackColor
Label21BackColor = Label21.BackColor
Label23BackColor = Label23.BackColor

FirstRun2 = False

End Sub


Private Sub Form_Unload(Cancel As Integer)


If Cheese = 3 Then

    'cheese = False
    
    'Dim Form2 As Form
    'For Each Form2 In Forms
    '    Unload Form2
    '    Set Form2 = Nothing
    'Next Form2

    Exit Sub

End If

If OptionsMenu.Visible = -1 Then
    OptionsMenu.Visible = False
    Cancel = True
    Exit Sub
End If

'If CID_Run_Me.Height = InitHeight Then
'    CID_Run_Me.Height = 2055
'    Cancel = True
'    ScrWidth_Old = 0
'    Exit Sub
'End If


gug = FreeFile
Open "C:\Callerid\CI LOGS\2CidRunME.txt" For Append As #gug

Print #gug, "Start = "; Format$(Startuptime, "DD/MM/YYYY HH:MM:SS"); " ";
Print #gug, "End = "; Format$(Now, "DD/MM/YYYY HH:MM:SS"); " ";
wds$ = "         "
LSet wds$ = Trim(str(Mousey))
Print #gug, "Mouse = "; wds$; " ";
wds$ = Trim(str(Keyy))
Print #gug, "Keyboard = "; wds$
Close #gug

gug = FreeFile
Open "C:\Callerid\CI Logs\2CidRunMeDayCount.txt" For Append As #gug
Print #gug, "Day Counter  = "; Now
Print #gug, "Day Mouse    = "; Daymouse
Print #gug, "Day Keyboard = "; Daykeyy
Close #gug


If CID_Run_Me.WindowState = 0 Then CID_Run_Me.WindowState = 1: Cancel = True: Exit Sub


Cheese = False

Dim Form1 As Form

For Each Form1 In Forms
    Unload Form1
    Set Form1 = Nothing
Next Form1

End Sub


Private Sub Label1_Click()
Call Command1_Click
End Sub

Private Sub Label14_Click()
Call Command3_Click
End Sub

Private Sub Label15_Click()

On Error Resume Next

we1 = InStr(RTB1.Text, "WinAmp Caller ID")
we3 = InStr(we1, RTB1.Text, "-----" + vbCr) + 5

RTB1.SelStart = we1 - 1
RTB1.SelLength = we3 - we1

'RTB1.SelBold = True
RTB1.SelFontSize = 8
RTB1.SelColor = &H80FFFF


we1 = InStr(RTB1.Text, "WinAmp Gold Amp")
we3 = InStr(we1, RTB1.Text, "-----" + vbCr) + 5

RTB1.SelStart = we1 - 1
RTB1.SelLength = we3 - we1

'RTB1.SelBold = True
'RTB1.SelFontSize = 8
RTB1.SelColor = &H80FFFF



'Call Label16_Click
we1 = InStr(RTB1.Text, Trim(str(Val(LastCurProcHwnd))))
we2 = InStrRev(RTB1.Text, "Title Name", we1)
we3 = InStr(we1, RTB1.Text, "-----" + vbCr) + 5

RTB1.SelStart = we2 - 1
RTB1.SelLength = we3 - we2

'RTB1.SelBold = True
'RTB1.SelFontSize = 8
RTB1.SelColor = &HFFFFFF


End Sub

Private Sub Label16_Click()

ListBox55 = Not ListBox55

jag = 3210

If ListBox55 = True Then
    CID_Run_Me.Height = InitHeight + jag
Else
    CID_Run_Me.Height = InitHeight
End If

CID_Run_Me.Left = Screen.Width - CID_Run_Me.Width - OffSetGoogle
CID_Run_Me.Top = Screen.Height - (CID_Run_Me.Height + HeightOffSet)

OptionsMenu.Width = CID_Run_Me.Width
OptionsMenu.Top = CID_Run_Me.Top - OptionsMenu.Height
OptionsMenu.Left = Screen.Width - OptionsMenu.Width - OffSetGoogle

'CID_Run_Me.Refresh


ScrWidth_Old = 0
End Sub

Function vbCompiled(Optional hWndMain As Variant, Optional Buffer As String) As Boolean

Dim nRtn As Long

Buffer = Space$(256)

nRtn = GetModuleFileNameA(hWndMain, Buffer, Len(Buffer))
Buffer = UCase(Left(Buffer, nRtn))

End Function


Function vbCompiled2(Optional hWndMain As Variant, Optional Buffer As String) As Boolean

Dim nRtn As Long

Buffer = Space$(256)

nRtn = GetModuleFileNameA(0&, Buffer, Len(Buffer))

Buffer = UCase(Left(Buffer, nRtn))

If Right(Buffer, 8) = "\VB6.EXE" Then
    vbCompiled2 = False
Else
    vbCompiled2 = True
End If

End Function


Private Sub List1_DblClick()
Call Label16_Click
End Sub

Private Sub Menu_Date1991_Click()
Shell "D:\VB6\VB-NT\Date1991\date1991.exe", vbNormalFocus
End Sub
Private Sub Menu_Pulse_Click()
Shell "D:\VB6\VB-NT\pulse\pulse.exe", vbNormalFocus
End Sub
Private Sub Menu_Alcohol120_Click()
Shell "R:\Program Files\Alcohol Soft\Alcohol 120\Alcohol.exe", vbNormalFocus
End Sub

Private Sub Menu_AttachmentExtractor_Click()
Shell "D:\VB6\VB-NT\Attachment Extractor 672673312002\Attachment Extractor.exe", vbNormalFocus
End Sub

Private Sub Menu_CreativePlayCenter_Click()
Shell "R:\Program Files\Creative\SBAudigy\PlayCenter2\CTPlay2.exe", vbNormalFocus
End Sub

Private Sub Menu_KeyboardM_Click()
Call Label6Hit
End Sub

Private Sub Menu_KnowProcess_Click()
Call Label16_Click
End Sub

Private Sub Menu_Mp3Tags_Click()
Shell "D:\VB6\VB-NT\MP3 Tags\MP3 Tags.exe", vbNormalFocus
End Sub

Private Sub Menu_NoteCallWinamp_Click()
'Shell "explorer D:\VB6\VB-NT\Cid-Run-Me\WinAmpLog\Winamp-list-CallerAmp.txt"
Shell "notepad D:\VB6\VB-NT\Cid-Run-Me\WinAmpLog\Winamp-list-CallerID.txt", vbNormalFocus
End Sub

Private Sub Menu_NoteGoldWinamp_Click()
Shell "notepad D:\VB6\VB-NT\Cid-Run-Me\WinAmpLog\Winamp-list-GoldAmp.txt", vbNormalFocus
End Sub


Private Sub Menu_StopCurrent_Click()

Call Command2_Click

End Sub

Private Sub Timer1_Timer()


'If QueTheAmpTest < Now Then
'    If ProcessId22 = 0 Then
'    ProcessId22 = Shell("R:\program files\winamp caller\winamp.exe", vbMinimizedNoFocus)
'    Chit22 = 1
'    ReRun = 1
'    LockInToMp3 = True
'    QueTheAmpTest = Now + TimeSerial(0, 0, 30)
'    End If
'End If

If Cheese = 3 Then
    W = MsgBox("Dont Run Twice.. Bye Bye..", vbInformation)
    

    End
End If

Dim pCount As Long
Dim CurProcHwnd As Long
Dim wfile2 As String
Dim wfile As String

Call FindCursor(x!, y!)

For i = 1 To 5
    bdf = GetAsyncKeyState(i)
    If bdf <> 0 Then MattsTimer2 = Now + MinorOnTime
Next



'Dim CurProcHwnd As Long
CurProcHwnd = GetForegroundWindow

If CurProcHwnd = 0 Then Exit Sub

If LockInToMp3 = True Then
    ReRun = 1
'    Chit22 = 1
End If

Dim Peet2 As Long
Dim Peet4 As Long

    
    

'If Tagad > 0 And ReRun = 0 And Curprochwnd = Me.hWnd Then CurProchwnd2 = Me.hWnd

Peet3$ = ""
Peet33$ = ""
If List2.ListCount > 0 Then
    For R = List2.ListCount - 1 To 0 Step -1
        peet1$ = List2.List(R)
        Peet2 = Val(Mid$(peet1$, 1, 9))
        Peet4 = Val(Mid$(peet1$, 11, 8))
        If GetWindow(Peet2, GW_HWNDFIRST) = 0 Then
            List2.RemoveItem R
            ReRun = 1
        Else
            Peet3$ = Peet3$ + str$(Peet2)
            Peet33$ = Peet33$ + str$(Peet4)
        End If
    Next
End If



If CurProcHwnd <> CurProcHwnd2 Or ReRun = 1 Or FirstRun2 = False Then
'start the stuff------------
        
If FindWindow(vbNullString, "Download") > 0 Then Call dss_download
If FindWindow(vbNullString, "Security Information") > 0 Then
    w21 = Timer + 0.1
    w22 = Now + TimeSerial(0, 0, 2)
    Do
        DoEvents
    Loop Until w22 < Now Or w21 < Timer
    AppActivate "Security Information", 0
    SendKeys "{enter}", True
End If
        
If FindWindow(vbNullString, "Microsoft Excel") > 0 Then
Dim HwndMicE As Long
HwndMicE = FindWindow(vbNullString, "Microsoft Excel")
Dim MeRyu3 As RECT
wert = GetWindowRect(HwndMicE, MeRyu3)
If MeRyu3.Top < 40 Then

MoveWindow HwndMicE, MeRyu3.Left, 131, MeRyu3.Right - MeRyu3.Left, MeRyu3.Bottom - MeRyu3.Top, True

End If

'MoveWindow Me.hWnd, Rect22me.Left, Rect22.Bottom, Rect22me.Right - Rect22me.Left, Rect22me.Bottom - Rect22me.Top, True

End If
    
        
        
        
If CurProcHwnd <> Me.hWnd Then LastCurProcHwnd = CurProcHwnd
        
        
        

    If InStr(Peet3$, str(CurProcHwnd)) = 0 Then
        totss = cProcesses.Convert(CurProcHwnd, Otss, cnFromhWnd Or cnToProcessID)
        List2.AddItem Format$(CurProcHwnd, "000000000") + " " + Format$(Otss, "0000000") + " "
        Peet3$ = Peet3$ + str$(CurProcHwnd)
        Peet33$ = Peet33$ + str$(Otss)
    
    End If



    FlingGrater1$ = ""
    FlingGrater2$ = ""

    NoMonOff2 = 0
    NoMonOff = 0
    NoMusic = 0
    frontpagepid = 0
    ProcessId25Ssam = 0

    Dim lhTmp As Long
    
    ash$ = ""
        
    List1.Clear
    RTB1.Text = ""
    
        
    Chit24 = 0
        
    If LockInToMp3 = False Then Chit22 = 0
        
    
    For R = 0 To List2.ListCount - 1
        
        peet1$ = List2.List(R)
        Peet2 = Val(Mid$(peet1$, 1, 9))
        Peet4 = Val(Mid$(peet1$, 11, 8))
            
        lhTmp = GetWindow(Peet2, GW_HWNDFIRST)
        ash$ = GetWindowTitle(Peet2)
           
        If CurProcHwnd = Peet2 Then
            lsidx = R
        End If
            
            
            wfile2 = GetFileFromHwnd(Peet2)
            
            Dim Ret As Long, sText As String
            sText = Space(255)
            Ret = GetClassName(Peet2, sText, 255)
            sText = Left$(sText, Ret)
            List1.AddItem "Title Name:-" + ash$
            List1.AddItem "Class Name:-" + sText
            List1.AddItem "File Name :-" + wfile2
            List1.AddItem "Hwnd:--" + str(Peet2)
            List1.AddItem "PID :--" + str(Peet4)
            List1.AddItem "-------------------------"
                
            Otss = Peet4
            
            Call Perfect(wfile2, 1)
            
            If NoMusic = 1 And OldNoMusic = 0 Then FlingGrater1$ = wfile2
            If NoMonOff = 1 And OldNoMonOff = 0 Then FlingGrater2$ = wfile2
            
            OldNoMusic = NoMusic
            OldNoMonOff = NoMonOff
            
            If InStr(LCase(wfile2), "amp caller\winamp.exe") Then
                ProcessId22 = Otss
                Chit22 = 1
                LockInToMp3 = False
            End If
                
            If InStr(LCase(wfile2), "amp gold\winamp.exe") Then
                ProcessId24 = Otss
                Chit24 = 1
            End If

            
                
        Next
    
        
        
        If (NoMusic = 1 And Cmdv = 1) Then NoMusic = 0
    
        If Ebuy = 0 Then Ebuyer = 0
    
        NoMonOffX = NoMonOff
        NoMusicX = NoMusic

        Call proc24
        
        List1.AddItem "Special Processes........."
        
        If ProcessId22 > 0 Then
            ash$ = GetWindowTitle(Winamp22Handle2)
            If ash$ = "" Then ProcessId22 = 0
        End If
        
        If ProcessId22 > 0 Then
            If InStr(Peet3$, str(Winamp22Handle2)) = 0 Then
                List2.AddItem Format$(Winamp22Handle2, "000000000") + " " + Format$(ProcessId22, "0000000") + " "
                Peet3$ = Peet3$ + str(Winamp22Handle2)
                Peet33$ = Peet33$ + str$(ProcessId22)
            End If
            wfile2 = GetFileFromHwnd(Winamp22Handle2)
            profile22$ = wfile2
            sText = Space(255)
            Ret = GetClassName(Winamp22Handle2, sText, 255)
            sText = Left$(sText, Ret)
            List1.AddItem "WinAmp Caller ID......"
            List1.AddItem profile22$
            List1.AddItem "Title Name--" + ash$
            List1.AddItem "Class Name--" + sText
            List1.AddItem "Hwnd:--" + str(Winamp22Handle2)
            List1.AddItem "PID :--" + str(ProcessId22)
            List1.AddItem "-------------------------"
            LockInToMp3 = False
            Chit22 = 1
        End If
        
        If ProcessId24 > 0 Then
            ash$ = GetWindowTitle(Winamp24Handle2)
            If ash$ = "" Then ProcessId24 = 0
        End If

        If ProcessId24 > 0 Then
            If InStr(Peet3$, str(Winamp24Handle2)) = 0 Then
                List2.AddItem Format$(Winamp24Handle2, "000000000") + " " + Format$(ProcessId24, "0000000") + " "
                Peet3$ = Peet3$ + str(Winamp24Handle2)
                Peet33$ = Peet33$ + str$(ProcessId24)
            End If
            wfile2 = GetFileFromHwnd(Winamp24Handle2)
            profile24$ = wfile2
            sText = Space(255)
            Ret = GetClassName(Winamp24Handle2, sText, 255)
            sText = Left$(sText, Ret)
            List1.AddItem "WinAmp Gold Amp......."
            List1.AddItem profile24$
            List1.AddItem "Title Name--" + ash$
            List1.AddItem "Class Name--" + sText
            List1.AddItem "Hwnd:--" + str(Winamp24Handle2)
            List1.AddItem "PID :--" + str(ProcessId24)
            List1.AddItem "-------------------------"
            Chit24 = 1
        End If
        
        ReRun = 0
    
        For R1 = 0 To List1.ListCount - 1
            RTB1.Text = RTB1.Text + List1.List(R1) + vbCr
        Next
        
        If FlingGrater1$ <> "" Then
            RTB1.Text = RTB1.Text + "NoMusic  := " + FlingGrater1$ + vbCr
        End If
        If FlingGrater2$ <> "" Then
            RTB1.Text = RTB1.Text + "NoMonOff := " + FlingGrater2$ + vbCr
        End If

    CurProcHwnd2 = CurProcHwnd
    
    Call Label15_Click
    
End If


For iCtr = 0 To UBound(sArray)
    Skip2 = FindWindow(vbNullString, sArray(iCtr))
    If Skip2 > 0 Then Exit For
Next

If Skip2 = 0 Then Skip2 = FindWindow(vbNullString, "Let The Backup Comence...")

If RASCount() Or Skip2 > 0 Then
    NoMusic = 1: NoMonOff = 1
    NoMusicX = 1: NoMonOffX = 1
End If



FirstRun2 = True

If Chit22 = 0 Then
'    ProcessId22 = 0
End If

'If Chit24 = 0 Then ProcessId24 = 0

w14 = GetWindow(Winamp22Handle2, GW_HWNDFIRST)
w15 = GetWindow(Winamp24Handle2, GW_HWNDFIRST)
If w14 < 1 And Winamp22Handle2 > 0 Then
    Winamp22Handle2 = 0: ProcessId22 = 0: ReRun = 1
End If
If w15 < 1 And Winamp24Handle2 > 0 Then
    Winamp24Handle2 = 0: ProcessId24 = 0: ReRun = 1
End If

If ProcessId22 > 0 Then wedf$ = "+" Else wedf$ = ""
If ProcessId24 > 0 Then wedg$ = "@" Else wedg$ = ""


If Label15.Caption <> Trim(str(List2.ListCount)) + wedf$ + wedg$ Then
    Label15.Caption = Trim(str(List2.ListCount)) + wedf$ + wedg$
End If

Call proc24

End Sub

Private Sub Timer2_Timer()



If CID_Run_Me.Top + CID_Run_Me.Height + HeightOffSet > Screen.Height Or _
    OptionsMenu.Top + OptionsMenu.Height + HeightOffSet > Screen.Height Then
    If XRad <> CID_Run_Me.Top + OptionsMenu.Top Then
        If GetAsyncKeyState(1) < 0 Then
            PeaShoot = Timer + 0.4
        End If
    End If
End If




XRad = CID_Run_Me.Top + OptionsMenu.Top




If (PeaShoot < Timer And PeaShoot > 0) Or PeaShoot - 100 > Timer Or (OldWState <> CID_Run_Me.WindowState) Then
    PeaShoot = 0
    
    If CID_Run_Me.WindowState = 0 Then
        CID_Run_Me.Left = Screen.Width - CID_Run_Me.Width - OffSetGoogle
        CID_Run_Me.Top = Screen.Height - (CID_Run_Me.Height + HeightOffSet)

        CID_Run_Me.Refresh

        OptionsMenu.Width = CID_Run_Me.Width
        OptionsMenu.Top = CID_Run_Me.Top - OptionsMenu.Height
        OptionsMenu.Left = Screen.Width - OptionsMenu.Width - OffSetGoogle
    End If
    
    If OptionsMenu.WindowState = 0 And CID_Run_Me.WindowState <> 0 Then
        'OptionsMenu.Width = CID_Run_Me.Width
        OptionsMenu.Top = Screen.Height - (HeightOffSet) - OptionsMenu.Height
        OptionsMenu.Left = Screen.Width - OptionsMenu.Width - OffSetGoogle
    End If
End If

'OldXRad = XRad
OldWState = CID_Run_Me.WindowState

If Day(DayCheck) <> Day(Now) Then
    
    If Day(DayCheck) = Day(Now) Then chucky = 1
    
    DayCheck = Now
    
    gug = FreeFile
    Open "C:\Callerid\ci logs\2CidRunMeDayCount.txt" For Append As #gug
    Print #gug, "End Of Day Counter  = "; Now
    Print #gug, "Day Mouse           = "; Daymouse
    Print #gug, "Day Keyboard        = "; Daykeyy
    Close #gug
    
    gug = FreeFile
    Open "C:\Callerid\ci logs\2CidRunMeDayCountLog.txt" For Append As #gug
    
    d$ = "          "
    u$ = "          "
    RSet d$ = Trim(str$(Daymouse))
    RSet u$ = Trim(str$(Daykeyy))
    
    wer$ = Format$(Now, "DD/MM/YY HH:MM:SS") + " - " + "Mouse = " + d$ + " @ " + "Keyboard = " + u$
    Print #gug, wer$
    Close #gug
    
    OptionsMenu.List1.AddItem wer$
    
    DayMouseTimer = Now + TimeSerial(0, 0, 10)
    TotalDayMouse = Daymouse
    TotalDayKeyy = Daykeyy
    Daymouse = 0
    Daykeyy = 0
    
    If chucky = 0 Then
        gug = FreeFile
        Open "C:\Callerid\ci logs\2CidRunMeDayCount.txt" For Output As #gug
        Print #gug, "Day Counter  = "; Now
        Print #gug, "Day Mouse    = "; Daymouse
        Print #gug, "Day Keyboard = "; Daykeyy
        Close #gug
    End If
End If


If AsusTimer20 > Now Then Call Proc25Asus


If DayMouseTimer < Now And DayMouseTimer > 0 And (Daykeyy > 0 Or Daymouse > 0) Then
    DayMouseTimer = 0
    TotalDayKeyy = 0
    TotalDayMouse = 0
    Label21.BackColor = Label21BackColor
    Label23.BackColor = Label23BackColor
End If


W = 0
If MattsTimer2 > Now Then W = DateDiff("s", Now, MattsTimer2)

If W > 0 Then
    If MattsTimer5 > Now Then
        If Label10.BackColor <> Label10.BackColor Then Label10.BackColor = Label10.BackColor
    Else
        If Label10.BackColor <> &HFF0000 Then Label10.BackColor = &HFF0000
    End If
End If

If W = 0 And MattsTimer5 > Now Then
    W = DateDiff("s", Now, MattsTimer5):
    If Label10.BackColor <> &H0 Then Label10.BackColor = &H0
End If

Label10.Caption = Trim$(str$(W))

Label2.Caption = Format$(Now, "DDD D-MM-YY H:MM:SSa/p")

If Xmarks < Now And Xmarks > 0 Then
    Xmarks = 0
    Label23.BackColor = Label23BackColor
End If

If NoMonMusChange <> (NoMonOff2 + 5) + (NoMonOff + 2) + (NoMusic + 4) Then
If NoMonOff = 0 Then Command3.BackColor = &HFFFF&
If NoMonOff = 1 Then Command3.BackColor = &HFF00&
If NoMonOff2 = 1 Then Command3.BackColor = &HFF80FF
If NoMusic = 0 Then Command1.BackColor = &HFFFF&
If NoMusic = 1 Then Command1.BackColor = &HFF00&
NoMonMusChange = (NoMonOff2 + 5) + (NoMonOff + 2) + (NoMusic + 4)
End If


If Screen.Width = 19200 And Screen.Height = 15360 And ScrWidth_Old <> 19200 Then

    CID_Run_Me.Left = Screen.Width - CID_Run_Me.Width - OffSetGoogle
    HeightOffSet = 400
    CID_Run_Me.Top = Screen.Height - (CID_Run_Me.Height + HeightOffSet)

End If



ScrWidth_Old = Screen.Width

Call dss_download


If Postimer < Now And Postimer > 0 And CID_Run_Me.WindowState = 0 Then
    Postimer = 0
    CID_Run_Me.Left = Screen.Width - CID_Run_Me.Width - OffSetGoogle
    CID_Run_Me.Top = Screen.Height - (CID_Run_Me.Height + HeightOffSet)
End If

If Tibulate > 0 And Tibulate < Now Then
    Tibulate = 0 ': CID_Run_Me.WindowState = 0
End If
If FormLoad1 = 1 Then
       MattsTimer2 = Now + MinorOnTime + MinorOnTime
    FormLoad1 = 0
End If

If Mid$(Time$, 4) = "20:00" Then
    MattsTimer5 = Now + TimeSerial(0, 2, 0)
End If

If Mid$(Time$, 4) = "40:00" Then
    MattsTimer5 = Now + TimeSerial(0, 2, 0)
End If

If Mid$(Time$, 4) = "00:00" Then
    MattsTimer5 = Now + TimeSerial(0, 8, 0)
End If

Dim MattsLockMonOff7AM

Dim Hg As Date

Hg = TimeValue(str$(Now))

If Hg < TimeSerial(7, 0, 0) Then
    MattsLockMonOff7AM = True
Else
    MattsLockMonOff7AM = False
End If

'Label3.Caption = str(NoMonoff)

If MattsTimer5 < MattsTimer2 And MattsTimer2 > 0 Then MattsTimer5 = Now - 3
If MattsTimer2 < MattsTimer5 And MattsTimer5 > 0 Then MattsTimer2 = Now - 3

If (MattsTimer5 < Now And MattsTimer5 > 0) And _
    (MattsTimer2 < Now And MattsTimer2 > 0) Then
'        MsgBox ("Ping")
        
        
        
    If Monitor2_Off_Timer = 0 Then
 '       MsgBox ("Ping")
        If NoMonOff = 0 And NoMonOff2 = 0 Then
            MattsTimer5 = 0: MattsTimer2 = 0
            
            SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_OFF
        End If
    End If
End If


'If MattsTimer5 < Now Then MattsTimer5 = 0
'If MattsTimer2 < Now Then MattsTimer2 = 0


If MattsTimer2 > Now And MattsTimer2Old <> MattsTimer2 Then
    MattsLockMonOff7AM = False
End If

If MattsTimer2 > Now And MattsTimer2Old < Now Or _
   MattsTimer5 > Now And MattsTimer5Old < Now Then
    If MattsLockMonOff7AM = False Then
        SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_ON
    End If
End If



MattsTimer2Old = MattsTimer2
MattsTimer5Old = MattsTimer5


'---------------------------

If Music_Off_Timer < Now Then Music_Off_Timer = 0
If f12_off_timer < Now Then f12_off_timer = 0
If Monitor2_Off_Timer < Now Then Monitor2_Off_Timer = 0

If Music_Off_Timer = 0 And Label1.Caption <> "Music On" Then Label1.Caption = "Music On": Command1.Caption = "Music"
If Monitor2_Off_Timer = 0 And Label14.Caption <> "Monitor On" Then Label14.Caption = "Monitor On": Command3.Caption = "Monitor"

If Music_Off_Timer <> 0 Then
    If Command1.Caption <> "Music Off" Then Command1.Caption = "Music Off"
    Label1.Caption = Format$(Music_Off_Timer - Now, "HH:MM:SS")
End If

If Monitor2_Off_Timer <> 0 Then
    If Command3.Caption <> "Monitor Off" Then Command3.Caption = "Monitor Off"
    Label14.Caption = Format$(Monitor2_Off_Timer - Now, "HH:MM:SS")
End If

    

End Sub

Private Sub Timer3_Timer()
  
  
'Screen.TwipsPerPixelY
Dim Hwnd23
Dim HwndCTLTask3
Dim Rect23 As RECT
HwndCTLTask3 = FindWindow("_GD_Sidebar", vbNullString)
If HwndCTLTask3 > 0 Then
    Hwnd23 = GetWindowRect(HwndCTLTask3, Rect23)
End If
If Rect23.Top = 0 Then OffSetGoogle = 0
If Rect23.Top > 0 Then OffSetGoogle = (Rect23.Right - Rect23.Left) * Screen.TwipsPerPixelY
If OldRect23 > 0 And HwndCTLTask3 = 0 Then CID_Run_Me.Left = Screen.Width - CID_Run_Me.Width - OffSetGoogle

OldRect23 = HwndCTLTask3
  
  
  
  
'Caller Amp
If Winamp22Handle2 > 0 Then

    ash$ = GetWindowTitle(Winamp22Handle2)

    If Trim(ash$) <> "" Then
        If InStr(ash$, "- Winamp") > 0 Then
            ash$ = Trim(Left(ash$, InStr(ash$, "- Winamp") - 1))
        End If
        If Atg$ <> ash$ Then
            If XtX1 = 0 And XtX1 > -1 Then XtX1 = Now
            If XtX1 + TimeSerial(0, 0, 10) < Now And XtX1 > 0 Then
                freef5 = FreeFile
                Open App.Path + "\WinAmpLog\Winamp-list-CallerID.txt" For Append As #freef5
                Print #freef5, Format$(XtX1, "DD/MM/YYYY HH:MM:SS ampm ") + "[Play ] " + ash$
                Close #freef5
                Atg$ = ash$
                XtX1 = 0
            End If
        End If
    End If
End If
  
'Gold Amp
If Winamp24Handle2 = 0 Then Exit Sub

ash$ = GetWindowTitle(Winamp24Handle2)

If Trim(ash$) <> "" Then
    pdq$ = "- Winamp"
    If InStr(ash$, pdq$) = 0 Then pdq$ = "Winamp 5.094": Exit Sub
    ash$ = Trim(Left(ash$, InStr(ash$, pdq$) - 1))
    
    eth = InStr(ash$, ". ")
    'eth2 = Val(Mid$(ash$, 1, eth - 1))
    
    hix = 0
        
    If OldAsh$ <> ash$ Then
        If OldAsh$ <> "" Then
            'eth3 = InStr(OldAsh$, ". ")
            'eth4 = Val(Mid$(OldAsh$, 1, eth3 - 1))
            For R = eth To Len(ash$)
                If Mid$(OldAsh$, R, 1) <> Mid$(ash$, R, 1) Then Exit For
            Next
            If R > Len(ash$) Then R = Len(ash$)
            
            If Asc(Mid$(ash$, R, 1)) < 48 Or _
                Asc(Mid$(ash$, R, 1)) > 57 Or _
                R = eth + 2 Then hix = 1
            
            If Mid$(ash$, eth, 4) = "----" Then hix = 0
        
        End If
        
        'If eth2 < eth4 Or _
        '    (eth2 = eth4 + 1 Or _
        '        eth2 = eth4 + 2) Then hix = 1
        
            
        If hix = 1 Then
            If Now - Agust > TimeSerial(0, 0, 30) And Agust2 < Now Then
                Call Bashing
                
                freef5 = FreeFile
                Open App.Path + "\WinAmpLog\Winamp-list-GoldAmp.txt" For Append As #freef5
                Print #freef5, Format$(Now, "DD/MM/YYYY HH:MM:SS ampm ") + "[End  ] " + Atg2$
                Close #freef5

                MMControl1.Command = "Prev"
                MMControl1.Command = "Play"
                EmSQ = 1
            End If
        End If
        
        OldAsh$ = ash$
    
    End If

    If Atg2$ <> ash$ Then
        If XtX2 = 0 And XtX2 > -1 Then XtX2 = Now
        If XtX2 + TimeSerial(0, 0, 10) < Now And XtX2 > 0 Then
            freef5 = FreeFile
            Open App.Path + "\WinAmpLog\Winamp-list-GoldAmp.txt" For Append As #freef5
            Print #freef5, Format$(XtX2, "DD/MM/YYYY HH:MM:SS ampm ") + "[Play ] " + ash$
            Close #freef5
            If InStr(ash$, "-----") = 0 Then Agust = Now
            Atg2$ = ash$
            XtX2 = 0
        End If
    End If
End If

End Sub



Private Sub Timer4_Timer()

For i = 5 To 255
    
    bdf = GetAsyncKeyState(i)
 
    If bdf <> 0 Then
        kbbdf = True
        Call CID_Run_Me.lastpress
    End If
Next

End Sub

Private Sub Timer6_Timer()

'Private Declare Function GetAsyncKeyState Lib "user32" _
        (ByVal vKey As Long) As Integer

bdfh2 = GetAsyncKeyState(1)
If bdfh2 <> 0 Then
    CurProcHwnd3 = GetForegroundWindow
    If Winamp24Handle2PL = CurProcHwnd3 Then
        Agust2 = Now + TimeSerial(0, 0, 2)
    End If
End If

End Sub


Public Sub proc22()

'Caller Amp
If (ProcessId22 <> oldprocessid22 And ProcessId22 > 0) Or PlimBurn1 = False Then
    Winamp22Handle2 = hWndFromProcID(ProcessId22)
    ReRun = 1
End If

If ProcessId22 > 0 And Winamp22Handle2 < 1 Then
    PlimBurn1 = False
Else
    PlimBurn1 = True
End If

MsgResult = SendMessage(Winamp22Handle2, WM_USER, 0, ByVal ipc_isplaying)

'// IPC_ISPLAYING returns the status of playback.
'// If it returns 1, it is playing. if it returns 3, it is paused,
'// if it returns 0, it is not playing.

If MsgResult = 1 Then
    If XtX1 = -1 Then XtX1 = 0
Else
    XtX1 = -1
End If

oldprocessid22 = ProcessId22

End Sub



Public Sub proc24()

'Gold Amp
Call proc22
    
'    winamp24handle2 = hWndFromProcID(ProcessId24)
If (ProcessId24 <> oldprocessid24 And ProcessId24 > 0) Or PlimBurn2 = False Then
    tagotd = 1
    Winamp24Handle2 = hWndFromProcID(ProcessId24)
    Winamp24Handle2PL = hWndFromProcIDPlayList(ProcessId24)
End If

If ProcessId24 > 0 And Winamp24Handle2 < 1 Then
    PlimBurn2 = False
Else
    PlimBurn2 = True
End If

oldprocessid24 = ProcessId24

MsgResult = SendMessage(Winamp24Handle2, WM_USER, 0, ByVal ipc_isplaying)

'// IPC_ISPLAYING returns the status of playback.
'// If it returns 1, it is playing. if it returns 3, it is paused,
'// if it returns 0, it is not playing.

If MsgResult = 1 And EmSQ = 1 Then
    EmSQ = 0
    freef5 = FreeFile
    Open App.Path + "\WinAmpLog\Winamp-list-GoldAmp.txt" For Append As #freef5
    Print #freef5, Format$(Now, "DD/MM/YYYY HH:MM:SS ampm ") + "[Play ] " + Atg2$
    Close #freef5
End If


If MsgResult = 1 Then
    NoMonOff = 1
    NoMusic = 1
    If XtX2 = -1 Then XtX2 = 0
Else
    EmSQ = 1
    NoMonOff = NoMonOffX
    NoMusic = NoMusicX
    Agust = Now - 4
    XtX2 = -1
End If

If Rem5 <> MsgResult Then
    Rem5 = MsgResult
    Call Timer2_Timer
End If

If tagotd = 1 Then ReRun = 1


End Sub



Public Sub Proc25Asus()

Exit Sub

Dim wed As Long

wed = FindWindow(vbNullString, "Asus Probe 2")

'Set groot = CreateObject(wed.Root)

'If wed > 0 Then
'If wed.WindowState = vbMinimized Then Me.Hide

Dim Ryu3 As RECT

Dim wer2 As Long
Dim wert As Long


wer2 = GetWindow(wed, wCmd)

wert = GetWindowRect(hWnd, Ryu3)

'ryu3.left
'ryu3.top
'ryu3.bottom
'groot.WindowState = vbMinimized
    
ShowWindow wed, SW_HIDE

ShowWindow wed, SW_MINIMIZE
'ShowWindow wed, SW_HIDE
ShowWindow wed, SW_SHOWNORMAL
ShowWindow wed, SW_NORMAL
ShowWindow wed, SW_SHOWMINIMIZED
ShowWindow wed, SW_SHOWMAXIMIZED
ShowWindow wed, SW_MAXIMIZE
ShowWindow wed, SW_SHOWNOACTIVATE
ShowWindow wed, SW_SHOW
ShowWindow wed, SW_MINIMIZE
ShowWindow wed, SW_SHOWMINNOACTIVE
ShowWindow wed, SW_SHOWNA
ShowWindow wed, SW_RESTORE
ShowWindow wed, SW_SHOWDEFAULT
ShowWindow wed, SW_FORCEMINIMIZE
ShowWindow wed, SW_MAX


'End If


End Sub

Public Sub FindCursor(ByRef x!, ByRef y!)

Dim P As POINTAPI

GetCursorPos P

x = P.x

y = P.y

If x <> 1 And y <> 1 And DbonPop = 1 And Dbon2 = 0 Then
    Dbon = Timer + 0.1
    Dbon2 = 1
End If

If x = 1 And y = 1 Then
    Dbon = Timer + 0.7: DbonPop = 1
End If

If Dbon - 100 > Timer Then
    Dbon = Timer
End If

If FirstRun = True Then
    Xmp1 = x
    Ymp1 = y
End If

If DbonPop = 1 And Dbon < Timer Then
    Dbon2 = 0
    DbonPop = 0
    Xmp1 = x
    Ymp1 = y
End If

If (x <> Xmp1 Or y <> Ymp1) And Dbon < Timer Then
    
    Mousey = Mousey + Sqr(Abs(Xmp1 - x) ^ 2 + Abs(Ymp1 - y) ^ 2)
    Daymouse = Daymouse + Sqr(Abs(Xmp1 - x) ^ 2 + Abs(Ymp1 - y) ^ 2)
    
    Xmarks = Now - 1
    MattsTimer2 = Now + MinorOnTime
    Call mystery
     
    If Mousey3 <> Now Then
        frmSender_CallID.txtMsg.Text = "MouseCounter " + str$(Mouse55 + Mousey)
        Call frmSender_CallID.cmdSend_Click
        Mousey3 = Now
    End If
    
    If Mousey > 0 Then
        
        Call ClingDing
        
    End If

    MattsTimer2 = Now + MinorOnTime

End If

Xmp1 = P.x

Ymp1 = P.y

End Sub

Public Sub ClingDing()

If TotalDayMouse > 0 Then
    Label21.Caption = Trim(str$(TotalDayMouse))
    Label23.Caption = Trim(str$(TotalDayKeyy))
    Label21.BackColor = QBColor(12)
    Label23.BackColor = QBColor(12)
    Exit Sub
End If


If CheckT1$ = "1" Then
    Label21.Caption = Trim(str$(Mousey))
    Label23.Caption = Trim(str$(Keyy))
End If
        
If CheckT1$ = "2" Then
    Label21.Caption = Trim(str$(Mouse55 + Mousey))
    Label23.Caption = Trim(str$(Qwerty2 + Keyy))
End If
        
If CheckT1$ = "3" Then
    Label21.Caption = Trim(str$(Daymouse))
    Label23.Caption = Trim(str$(Daykeyy))
End If




End Sub




Public Sub dss_download()

Dim nMaxCount As Long
Dim lpClassName As String
Dim hand3 As Long
Dim cch As Long
Dim lpString As String
  
handle2 = FindWindow(vbNullString, "Download")

filehwnd2$ = ""
If handle2 > 0 Then filehwnd2$ = GetFileFromHwnd(handle2)

If filehwnd2$ = "R:\Program Files\Olympus\DSSPlayer2002\DictWnd.exe" Then
    
    On Local Error Resume Next
    AppActivate "Download", False
    On Local Error GoTo 0

    Ew = Now + TimeSerial(0, 0, 2)
    Do
        DoEvents
        ash$ = GetActiveWindowTitle(False)
    Loop Until Ew < Now Or ash$ = "Download"


    SendKeys "{enter}", False

End If

End Sub




Public Sub Bashing()

If Command2.BackColor = QBColor(12) Then Exit Sub

If ProcessId24 > 0 Then
    
    MsgResult = SendMessage(Winamp24Handle2, WM_USER, 0, ByVal ipc_isplaying)

    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.

    'if playing stop plyaing it
    If MsgResult = 1 Then
     
        MsgResult = SendMessage(Winamp24Handle2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
        
    End If
End If
        
End Sub

Public Sub winamp_player_off()

'If Music_Off_Timer > 0 Then Exit Sub
    
If ProcessId22 > 0 Then
    
    MsgResult = SendMessage(Winamp22Handle2, WM_USER, 0, ByVal ipc_isplaying)

    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.

    If MsgResult = 1 Then
     
        MsgResult = SendMessage(Winamp22Handle2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
        
    End If
End If
        
End Sub

Function RASCount() As Integer
  Dim lpRasConn(0 To 1) As Long ' dummy buffer area
  Dim rc As Long ' return code
  Dim lpcb As Long ' buffer size
  Dim lpcConnections As Long ' connection count

  lpRasConn(0) = 32 ' each returned item is at least 32 bytes long
  lpcb = 0 ' set total number of usable bytes in the buffer to zero
  
  rc = RasEnumConnections(lpRasConn(0), lpcb, lpcConnections)
  RASCount = lpcConnections ' return connection count

End Function

Public Sub DUN_Services(DUN_Array() As String)
    
    'Pass in Empty array for DUN_Array
    Dim S As Long, Ln As Long, conname As String, i As Long
    Dim R(255) As RASENTRYNAME95
    R(0).dwSize = 264
    S = 256 * R(0).dwSize
    
    Call RasEnumEntriesA(vbNullString, vbNullString, R(0), S, Ln)
    Ln = Ln - 1
    ReDim DUN_Array(200)
    If Ln = 0 Then Exit Sub
    
    pigbum = Ln
    
    For i = 0 To Ln
        conname = StrConv(R(i).szEntryName(), vbUnicode)
        If InStr(conname, "Broadband") > 0 Or _
        InStr(conname, "Bluetooth") > 0 Then
            For i2 = i + 1 To Ln
                R(i2 - 1) = R(i2)
            Next
        pigbum = pigbum - 1
        End If
    Next
    
    Ln = pigbum
    
    For i = 0 To Ln
        conname = StrConv(R(i).szEntryName(), vbUnicode)
        DUN_Array(i) = Left$(conname, InStr(conname, _
          vbNullChar) - 1)
    Next i
    
    pt = i
    For i = 0 To Ln
        conname = StrConv(R(i).szEntryName(), vbUnicode)
        DUN_Array(i + pt - 1) = "Connect " + Left$(conname, InStr(conname, _
          vbNullChar) - 1)
    Next i
    
    pt = pt + i
    For i = 0 To Ln
        conname = StrConv(R(i).szEntryName(), vbUnicode)
        DUN_Array(i + pt - 1) = "Connecting " + Left$(conname, InStr(conname, _
          vbNullChar) - 1) + "..."
    Next i

    pt = pt + i
    For i = 0 To Ln
        conname = StrConv(R(i).szEntryName(), vbUnicode)
        DUN_Array(i + pt - 1) = "Connecting to " + Left$(conname, InStr(conname, _
          vbNullChar) - 1)
    Next i



'Skip2 = FindWindow(vbNullString, "Connect BTOW - Surftime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connect BTOW - Daytime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting BTOW - Daytime...")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting BTOW - Surftime...")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting to BTOW - Daytime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting to BTOW - Surftime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Dial-up Connection")
    
    
    
    DUN_Array(pt + i - 1) = "Dial-up Connection"


    ReDim Preserve DUN_Array(pt + i - 1)




End Sub


Public Sub Winamp_Player_On()

aw$ = globalhot$

   'ras listesini yuklesi baslar
   


If NoMusic <> 0 Or (Music_Off_Timer > Now And Music_Off_Timer > 0) Then Exit Sub
        

    For iCtr = 0 To UBound(sArray)
        Skip2 = FindWindow(vbNullString, sArray(iCtr))
        If Skip2 > 0 Then Exit For
    Next


'Skip2 = FindWindow(vbNullString, "Connect BTOW - Surftime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connect BTOW - Daytime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting BTOW - Daytime...")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting BTOW - Surftime...")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting to BTOW - Daytime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting to BTOW - Surftime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Dial-up Connection")
     
     
     
     
     
     
     
If Skip2 > 0 Then NoMusic = 1

If ProcessId22 = 0 Then
    
    If LockInToMp3 = False Then
        ProcessId22 = Shell("R:\program files\winamp caller\winamp.exe", vbMinimizedNoFocus)
    End If
    Chit22 = 1
    ReRun = 1
    LockInToMp3 = True
    Call proc22
    
Else

    MsgResult = SendMessage(Winamp22Handle2, WM_USER, 0, ByVal ipc_isplaying)
    
    If MsgResult = 0 Or MsgResult = 3 Then
        
           MsgResult = SendMessage(Winamp22Handle2, WM_COMMAND, WINAMP_BUTTON2, ByVal XcNul)
    
    End If

End If

End Sub



Public Sub mystery()
    
    'Call Winamp_Player_On
    
If kbbdf = True Then
'frmReceiver.txtTarget.Text = ""
    Keyy = Keyy + 1
    Daykeyy = Daykeyy + 1
    
    frmSender_CallID.txtMsg.Text = "KeyCounter " + str$(Keyy + Qwerty2)
    Call frmSender_CallID.cmdSend_Click
    
    Xmarks = Now
    
    Call ClingDing
    
    If TotalDayMouse = 0 And TotalDayKeyy = 0 Then Label23.BackColor = &HFFC0C0
    
    MattsTimer2 = Now + MinorOnTime

End If


If Trim(globalhot$) = "" Then Exit Sub

If InStr(globalhot$, "Hit Request") Then
    frmSender_CallID.txtMsg.Text = "KeyCounter " + str$(Qwerty2 + Keyy)
    Call frmSender_CallID.cmdSend_Click
    frmSender_CallID.txtMsg.Text = "MouseCounter " + str$(Mouse55 + Mousey)
    Call frmSender_CallID.cmdSend_Click
End If

If InStr(globalhot$, "winamp play") Then
    Call Winamp_Player_On
    MattsTimer5 = Now + TimeSerial(0, 3, 0)
End If


If InStr(globalhot$, "winamp stop") Then
    Call winamp_player_off
    If MattsTimer5 > Now Then MattsTimer5 = Now + TimeSerial(0, 1, 30)
End If

If InStr(globalhot$, "Monitor To The Fucking A") Then
    'Call winamp_player_off
    MattsTimer5 = Now + TimeSerial(0, 3, 0)
End If

globalhot$ = ""

End Sub

Public Sub lastpress()

kbbdf = True

If kbbdf = True Then Call mystery

kbbdf = flase

End Sub

Private Sub Timer5_Timer()

If AutoConVpn < Now Then
    AutoConVpn = Now + 30
    Shell "D:\VB6\VB-NT\Autoconnect payasgo\Auto_Connect_VPN_PAG.exe", vbNormalFocus

    AutoConVpn = Now + 32
    AutoConVpn = DateSerial(Year(AutoConVpn), Month(AutoConVpn), 1)
    AutoConVpn = AutoConVpn + TimeSerial(21, 0, 0)
    Open App.Path + "\AutoVPN.txt" For Output As #1
    Print #1, str(AutoConVpn)
    Close #1


End If

If Timer5Div > Now Then Exit Sub

Timer5Div = Int(Now) + TimeSerial(Hour(Now) + 1, 0, 0)


If Dir$("R:\Documents and Settings\Matt Lan\NTUSER.DAT.LOG") = "" Then FirstTime5Run = 1


If FirstTime5Run = 0 Then

    FirstTime5Run = 1


    WLFN1$ = "R:\Documents and Settings\Matt Lan\NTUSER.DAT.LOG"
    WLFN2$ = "G:\Documents and Settings\Matt Lan\NTUSER.DAT.LOG"
    WLFN3$ = "R:\Documents and Settings\Matt Lan\Local Settings\Application Data\Microsoft\Windows\UsrClass.dat.LOG"
    WLFN4$ = "G:\Documents and Settings\Matt Lan\Local Settings\Application Data\Microsoft\Windows\UsrClass.dat.LOG"

    wsfn1$ = GetShortName(WLFN1$)
    wsfn2$ = GetShortName(WLFN2$)
    wsfn3$ = GetShortName(WLFN3$)
    wsfn4$ = GetShortName(WLFN4$)

    If wsfn2$ = "" Then
        Open WLFN2$ For Output As #1: Close #1
        MsgBox ("Backup File :-" + vbCr + WLFN2$ + vbCr + "Didn't Exist - So Created an Empty File" + vbCr + "Be Sure You Do A Dos Level Backup P.D.Quick..")
        wsfn2$ = GetShortName(WLFN2$)
    End If

    If wsfn4$ = "" Then
        Open WLFN4$ For Output As #1: Close #1
        MsgBox ("Backup File :-" + vbCr + WLFN4$ + vbCr + "Didn't Exist - So Created an Empty File" + vbCr + "Be Sure You Do A Dos Level Backup P.D.Quick..")
        wsfn4$ = GetShortName(WLFN4$)
    End If

    Open "C:\bat\vb_code\LSFNSupr.txt" For Output As #1
    Print #1, wsfn1$
    Print #1, wsfn2$
    Print #1, wsfn3$
    Print #1, wsfn4$
    Close #1

End If





Dim MvarByte1 As Long
Dim MvarByte2 As Long
Dim HwndTest1 As Long
Dim HwndTest2 As Long

Dim GetMeFileData1 As WIN32_FIND_DATA
Dim GetMeFileData2 As WIN32_FIND_DATA

HwndTest1 = FindFirstFile(WLFN1$, GetMeFileData1)

HwndTest2 = FindClose(HwndTest)

HwndTest1 = FindFirstFile(WLFN2$, GetMeFileData2)

HwndTest2 = FindClose(HwndTest)

If GetMeFileData1.nFileSizeHigh = 0 Then
    MvarByte1 = GetMeFileData1.nFileSizeLow
    Else
    MvarByte1 = GetMeFileData1.nFileSizeHigh
End If

If GetMeFileData2.nFileSizeHigh = 0 Then
    MvarByte2 = GetMeFileData2.nFileSizeLow
    Else
    MvarByte2 = GetMeFileData2.nFileSizeHigh
End If
                

If MvarByte1 < MvarByte2 Then
    MsgBox ("After An Hourly Check Following Results Found:-" + vbCr + WLFN1$ + vbCr + WLFN2$ + vbCr + "Destination File Bigger than Original" + vbCr + "Action Taken Will Be The Chop..")
    
    Dim Part1$
    filefree1 = FreeFile
    Open WLFN2$ For Binary As filefree1
    Part1$ = Input(MvarByte1, filefree1)
    Close #filefree1
    
    Kill WLFN2$
    
    filefree1 = FreeFile
    Open WLFN2$ For Output As filefree1
    Print #filefree1, Part1;
    Close #filefree1

End If





HwndTest1 = FindFirstFile(WLFN3$, GetMeFileData1)

HwndTest2 = FindClose(HwndTest)

HwndTest1 = FindFirstFile(WLFN4$, GetMeFileData2)

HwndTest2 = FindClose(HwndTest)

If GetMeFileData1.nFileSizeHigh = 0 Then
    MvarByte1 = GetMeFileData1.nFileSizeLow
    Else
    MvarByte1 = GetMeFileData1.nFileSizeHigh
End If

If GetMeFileData2.nFileSizeHigh = 0 Then
    MvarByte2 = GetMeFileData2.nFileSizeLow
    Else
    MvarByte2 = GetMeFileData2.nFileSizeHigh
End If
                

If MvarByte1 < MvarByte2 Then
    MsgBox ("After An Hourly Check Following Results Found:-" + vbCr + WLFN3$ + vbCr + WLFN4$ + vbCr + "Destination File Bigger than Original" + vbCr + "Action Taken Will Be The Chop..")
    
    filefree1 = FreeFile
    Open WLFN3$ For Binary As filefree1
    Part1$ = Input(MvarByte1, filefree1)
    Close #filefree1
    
    Kill WLFN4$
    
    filefree1 = FreeFile
    Open WLFN4$ For Output As filefree1
    Print #filefree1, Part1;
    Close #filefree1
    
End If




End Sub
