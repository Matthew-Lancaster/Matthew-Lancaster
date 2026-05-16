VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "RS232 Reed Contact Switch Door"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   660
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer7Test 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   8925
      Top             =   3675
   End
   Begin VB.Timer TimerControl 
      Interval        =   500
      Left            =   9150
      Top             =   2805
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7530
      Top             =   3480
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Left            =   7440
      Top             =   2880
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6390
      Top             =   1845
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5550
      Top             =   2100
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2400
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      Handshaking     =   1
      NullDiscard     =   -1  'True
      RTSEnable       =   -1  'True
   End
   Begin MCI.MMControl MMControl2 
      Height          =   960
      Left            =   2265
      TabIndex        =   1
      Top             =   4050
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1693
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4350
      Top             =   1740
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5250
      Left            =   -30
      TabIndex        =   0
      Top             =   1095
      Width           =   18270
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   11610
      TabIndex        =   7
      Top             =   555
      Width           =   6615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Since Last Door Close - "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   -15
      TabIndex        =   6
      Top             =   555
      Width           =   10620
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   540
      Left            =   10620
      TabIndex        =   5
      Top             =   555
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Height          =   540
      Left            =   13500
      TabIndex        =   4
      Top             =   0
      Width           =   270
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "*The One* Are *IN*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   13770
      TabIndex        =   3
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "RS232 Missing Pulse Detector Starting Up"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   -15
      TabIndex        =   2
      Top             =   0
      Width           =   13500
   End
   Begin VB.Menu Mnu_VB 
      Caption         =   "VB ME"
   End
   Begin VB.Menu Mnu_log 
      Caption         =   "LOGG"
   End
   Begin VB.Menu Mnu_loggFolder 
      Caption         =   "! LOGG FOLDER"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XONow, ONow, ONow2, E, D, F1, F2, F3, F4, Buffer, AreYouIn, FlagOpen, Count2, Count1, LowTest, B
Dim HandShakePass, Timer7TestVar, Ticker1, CountR, CountR2
Dim DoorTime, TICK2, Dex, HighNow, HighNow2, Status2, Status
Dim YESUNLOAD, RSXErrors

Dim InPutDate, TestDate, Year5, MonthStart, DayCount1, DayCountMonth, WholeYear1, DayCountT As Integer
Dim Month5, WeeksSinceYear, WeeksSinceInput, ResultYearDate, Month7, WeeksPlusDays

'mscomm
Dim OCTSHoldingVAR, RUNME, RUNME2, TS, QQ, OldList1Count, CountD



Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
   (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
   
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hWnd As Long, ByVal wmessage As Long, ByVal wParam As Long, _
   lParam As Any) As Long
   
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
   (ByVal hWnd As Long, ByVal wmessage As Long, ByVal wParam As Long, _
   lParam As Any) As Long
Private Declare Function PostMessageLng Lib "user32" Alias "PostMessageA" _
   (ByVal hWnd As Long, ByVal wmessage As Long, ByVal wParam As Long, _
   ByVal lParam As Long) As Long

Private Const WM_COMMAND = &H111
Private Const WM_COPYDATA = &H4A

Private Const WINAMP_BUTTON1 = 40044
Private Const WINAMP_BUTTON2 = 40045
Private Const WINAMP_BUTTON3 = 40046
Private Const WINAMP_BUTTON4 = 40047
Private Const WINAMP_BUTTON5 = 40048

Private Const WM_USER = &H400
Private Const WM_WA_IPC = WM_USER

Private Const IPC_GETVERSION = 0
Private Const IPC_PLAYFILE = 100
Private Const IPC_DELETE = 101
Private Const IPC_STARTPLAY = 102
Private Const IPC_CHDIR = 103
Private Const IPC_ISPLAYING = 104
Private Const IPC_GETOUTPUTTIME = 105
Private Const IPC_JUMPTOTIME = 106
Private Const IPC_WRITEPLAYLIST = 120

Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Public Enum enumPlay
    playPlaying = 1
    playPaused = 3
    playStoped = 0
End Enum

Public Enum enumButtons
    PreviousTrackButton = 40044
    NextTrackButton = 40048
    PlayButton = 40045
    PauseUnpauseButton = 40046
    StopButton = 40047
    StopAfterCurrentTrack = 40147
    FadeoutAndStop = 40157
    FastForward5Seconds = 40148
    FastRewind5Seconds = 40144
    StartOfPlaylist = 40154
    GoToEndOfPlaylist = 40158
    OpenFileDialog = 40029
    OpenUrlDialog = 40155
    OpenFileInfoBox = 40188
    SetTimeDisplayModeToElapsed = 40037
    SetTimeDisplayModeToRemaining = 40038
    TogglePreferencesScreen = 40012
    OpenVisualizationOptions = 40190
    OpenVisualizationPlugInOptions = 40191
    ExecuteCurrentVisualizationPlugIn = 40192
    ToggleAboutBox = 40041
    ToggleTitleAutoscrolling = 40189
    ToggleAlwaysOnTop = 40019
    ToggleWindowshade = 40064
    TogglePlaylistWindowshade = 40266
    ToggleDoublesizeMode = 40165
    ToggleEq = 40036
    TogglePlaylistEditor = 40040
    ToggleMainWindowVisible = 40258
    ToggleMinibrowser = 40298
    ToggleEasymove = 40186
    RaiseVolumeBy1Perc = 40058
    LowerVolumeBy1Perc = 40059
    ToggleRepeat = 40022
    ToggleShuffle = 40023
    OpenJumpToTimeDialog = 40193
    OpenJumpToFileDialog = 40194
    OpenSkinSelector = 40219
    ConfigureCurrentVisualizationPlugIn = 40221
    ReloadTheCurrentSkin = 40291
    CloseWinamp = 40001
    MovesBack10TracksInPlaylist = 40197
    ShowTheEditBookmarks = 40320
    AddsCurrentTrackAsABookmark = 40321
    PlayAudioCd = 40323
    LoadAPresetFromEq = 40253
    SaveAPresetToEqf = 40254
    OpensLoadPresetsDialog = 40172
    OpensAutoLoadPresetsDialog = 40173
    LoadDefaultPreset = 40174
    OpensSavePresetDialog = 40175
    OpensAutoLoadSavePreset = 40176
    OpensDeletePresetDialog = 40178
    OpensDeleteAnAutoLoadPresetDialog = 40180
End Enum

Private m_sWinampDir As String
Private m_cPlaylist As Collection




Private Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

'On Error Resume Next
'MkDir App.Path + "\" + GetComputerName
'On Error GoTo 0
'App.Path + "\" + GetComputerName + "



Sub AutoSizeForm()


'Code to Auto Size Form based on controls used Including Menu Bar Sketchy Style
'Make sure form set to scale Twips
'put in your form load

Me.Show

x = 1
y = 1
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True And Control.Visible = True Then
        If Control.Width + Control.Left > x Then x = Control.Width + Control.Left
        If Control.Height + Control.Top > y Then y = Control.Height + Control.Top
        If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
    End If
Next
On Error GoTo 0

Me.Width = (x + 80)
If mnu = 1 Then
    pluso = 720 ': pluso = 1100 'Sometimes Different
Else
    pluso = 450
End If

Me.Height = (y + pluso)
Me.Refresh
DoEvents

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2




End Sub


Function FileInUse(ByVal strFilePath As String) As Boolean
  
  Dim hFile As Long
  Dim FileInfo  As OFSTRUCT
  
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function
  If Right(strFilePath, 1) <> Chr(0) Then strFilePath = strFilePath & Chr(0)
  
  FileInfo.cBytes = Len(FileInfo)
  hFile = OpenFile(strFilePath, FileInfo, OF_SHARE_EXCLUSIVE)
  If hFile = -1 And Err.LastDllError = 32 Then
    FileInUse = True
  Else
    CloseHandle hFile
  End If
  
End Function

Sub FileInUseDelay()
        
    Tx$ = TS
    
        
    QQ = Now + TimeSerial(10, 0, 0)
    Do
        'DoEvents
        ii = FileInUse(Tx$)
        If ii = True Then
            Label3.BackColor = &HFF&

            Sleep (100)
        End If
    Loop Until ii = False Or QQ < Now
    QQ = 0
    
    If ii = True Then
        List1.AddItem "Trouble File Stuck In Use " + vbCrLf + Tx$ + vbCrLf + "Longer than 10 Min"
        List1.ListIndex = List1.ListCount - 1

        Stop
    End If
    Label3.BackColor = &HFF00&
End Sub



Private Sub Form_Load()


'05-Nov-09 08:49:00a

'InPutDate = DateValue("03-Nov-09 06:52:12a") + TimeValue("03-Nov-09 06:52:12a")
'TestDate = DateValue("05-Nov-09 08:49:00a") + TimeValue("05-Nov-09 08:49:00a")
'Call FindTimeInfo

'dex
'Stop

On Error Resume Next
MkDir App.Path + "\" + GetComputerName
MkDir App.Path + "\" + GetComputerName + "\RS232 DOOR REED DATA"
On Error GoTo 0

XONow = -2

Me.BackColor = QBColor(1)

HandShakePass = -2

Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False

'NEED A PORT SCANNER IN HERE
MSComm1.CommPort = 1
'MSComm1.PortOpen = True

'MSComm1.Settings = "9600,N,8,1"

' Tell the control to read entire buffer when Input
' is used.
MSComm1.InputLen = 0

On Error Resume Next
'MSComm1.RTSEnable = True
'MSComm1.DTREnable = True

Start2:


DoEvents
MSComm1.PortOpen = True

If Err.Number = 8002 Then MsgBox "Port 2 RS232 - Wont OPEN -- MUST BE COMM 1": End
If Err.Number <> 0 Then MsgBox "Port 2 RS232 - Wont OPEN UnKnown Error #" + Str(Err.Number) + " - " + Err.Description: End



'rf = Now + TimeSerial(0, 0, 2)
'Do
    'DoEvents
'    If MSComm1.CTSHolding = False Then Sleep 100
'Loop Until MSComm1.CTSHolding = True Or rf < Now

DoEvents

If MSComm1.CTSHolding = False Then
    MSComm1.PortOpen = False
    Label6 = Trim(Str(Val(Label6) + 1))

    GoTo Start2
    
    MsgBox "Start With Door OPEN - EndE": End
End If



MMControl2.Notify = True
MMControl2.Wait = True
MMControl2.Shareable = False
MMControl2.DeviceType = "WaveAudio"
MMControl2.FileName = "D:\VB6\VB-NT\00_Best_VB_01\ClipBoard Logger\01 Woarble Tone 5 Mins.wav"
MMControl2.Command = "Open"

'GoTo Test



TS = App.Path + "\" + GetComputerName + "\RS232 DOOR REED DATA\RS232 Reed Door Entry Exit - " + Format$(Now, "yyyy-mm-dd DDD") + ".txt"

TICK = "----------------------------------------------------------------------------------"
List1.AddItem TICK


TICK = "### Prog - RS232 Reed Contact Switch Door ------------------"
List1.AddItem TICK

TICK = "### Prog - ROOM 14 - 3 AYMER ROAD - BN3 4GB -----------"
List1.AddItem TICK

 

TICK = "### Prog Start ---- " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p") + " -----------------"
List1.AddItem TICK

TICK = "----------------------------------------------------------------------------------"
List1.AddItem TICK

If Err.Number = 0 Then
    
    TICK = "Comm " + Format$(MSComm1.CommPort, "00") + " Modem Working"
    Else
    TICK = "Comm " + Format$(MSComm1.CommPort, "00") + " Modem Not Working"
End If

List1.AddItem TICK


TICK = "----------------------------------------------------------------------------------"
List1.AddItem TICK

CountD = 1




OCTSHoldingVAR = -10

RUNME = -2
RUNME2 = -2
    

'MSComm1.DTREnable = True

'MSComm1.RTSEnable = True



AreYouIn = True




Test:

Call AutoSizeForm


If IsIDE = False Then Me.WindowState = 1

Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False


List1.AddItem "Testing Very Great Golden Handshaking..."
DoEvents
Call Timer1_Timer

Timer5.Enabled = True

Label6 = "RSX #" + Trim(Str(RSXErrors)) + Format(Now, " ddd hh:mm:ssa/p")





End Sub

Private Sub Form_Resize()

On Error Resume Next
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Exit Sub
    
    


If Me.WindowState <> 1 And IsIDE = False Then
    Me.WindowState = 1
    Cancel = True
    Exit Sub
End If


Timer1.Enabled = False

'MSComm1.RTSEnable = False
'DoEvents
'MSComm1.DTREnable = False
'DoEvents

MSComm1.PortOpen = False
'DoEvents

TICK = "----------------------------------------------------------------------------------"
List1.AddItem TICK

Call Timer1_Timer


Call FileInUseDelay

TICK = "### Prog End ---- " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p") + " ---------------------"
fr1 = FreeFile
Open TS For Append Lock Read Write As #fr1
Print #fr1, TICK
Print #fr1, "----------------------------------------------------------------------------------"
Print #fr1,
Print #fr1,

Close #fr1



If B = 1 Then Kill TS



End Sub

Private Sub Label2_DblClick()
CountR = CountR + 1
If CountR < 3 Then Exit Sub
CountR = 0
Label2 = "*The One* Are *IN*"
AreYouIn = True
FlagOpen = False

End Sub

Private Sub Label3_DblClick()
'If IsIDE = False Then Exit Sub
CountR2 = CountR2 + 1
If CountR2 < 4 Then Exit Sub
B = 1
Label3.BackColor = &HC00000
End Sub

Private Sub Label5_Click()
'Label5.
End Sub

Private Sub List1_DblClick()
If InStr(List1.List(List1.ListIndex), "BN3") Then
List1.List(List1.ListIndex) = "### Prog - ROOM % - ** PRIVATE ** -------------------------------"
End If

End Sub

Private Sub Mnu_Log_Click()

If AreYouIn = False Then Exit Sub

Shell "notepad " + TS, vbNormalFocus

End Sub

Private Sub Mnu_loggFolder_Click()
If AreYouIn = False Then Exit Sub
Shell "explorer /e, " + App.Path + "\RS232 DOOR REED DATA", vbNormalFocus


End Sub


Private Sub Mnu_VB_Click()

If AreYouIn = False Then Exit Sub

Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
YESUNLOAD = True
Unload Me

End Sub
Private Sub MSComm1_OnComm()

'    If Timer5.Enabled = True Then Exit Sub

   
   Select Case MSComm1.CommEvent
   ' Handle each event or error by placing
   ' code below each case statement

   ' Errors
      Case comEventBreak   ' A Break was received.
      pf = 1
      
      Case comEventFrame   ' Framing Error
      pf = 1
      Case comEventOverrun   ' Data Lost.
      pf = 1
      Case comEventRxOver   ' Receive buffer overflow.
      pf = 1
      Case comEventRxParity   ' Parity Error.
      pf = 1
      Case comEventTxFull   ' Transmit buffer full.
      pf = 1
      Case comEventDCB   ' Unexpected error retrieving DCB]
      pf = 1

   ' Events
      
      Case comEvCD   ' Change in the CD line.
      
      pf = 1
        
'        Label1 = " - " + Format(Now, " DDD DD-MMM-YY HH:MM:SSa/p") + " - " + Format(Timer, "0.00000000")
'        List1.AddItem "Change in the CD line -- " + Str(MSComm1.CDHolding) + Label1

'        Me.WindowState = 0
'        DoEvents
      
'      MsgBox "Change in the CD line."
        
      
      Case comEvCTS   ' Change in the CTS line.
      pf = 1
      
      
'        Me.WindowState = 0
'        DoEvents
      
'        MsgBox "Change in the CTS line."
      
'        If Timer1.Enabled = True Then Call DealDoor
'        Label1 = " - " + Format(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " - " + Format(Timer, "0.000")
'        List1.AddItem Format(Count2, "00000 ") + "Change in the CTS line -- " + Str(MSComm1.CTSHolding) + Label1
        
        Count2 = Count2 + 1
        
        'If Timer1.Enabled = True Then Call DealDoor
        
        If Timer1.Enabled = True And Timer7TestVar = False Then Call DealDoor
  
            

        If Timer7TestVar = True And LowTest = 2 Then
            Label4 = "@"
            LowTest = 3
        End If
        
        If Timer7TestVar = True And LowTest = 1 Then
            Label4 = "#"
            LowTest = 2
        End If
        
        
        
        
        
        Case comEvDSR   ' Change in the DSR line.
        pf = 1
'        Count2 = Count2 + 1
      
'        MsgBox "Change in the DSR line."
        
'        Label1 = " - " + Format(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " - " + Format(Timer, "0.000")
'        List1.AddItem Format(Count2, "00000 ") + "Change in the DSR line -- " + Str(MSComm1.DSRHolding) + Label1
        
'        Count2 = Count2 + 1
      
        Count2 = Count2 + 1
        
        DoorTime = Now
        
        If Timer1.Enabled = True And Timer7TestVar = False Then Call DealDoor
      
      
      Case comEvRing   ' Change in the Ring Indicator.
      pf = 1
      
      
      Case comEvReceive   ' Received RThreshold # of
      pf = 1
                        ' chars.
             
'             Buffer = Buffer & MSComm1.Input

      
      Case comEvSend   ' There are SThreshold number of
      pf = 1
                     ' characters in the transmit
                     ' buffer.
        a = a
      
      Case comEvEOF   ' An EOF charater was found in
      pf = 1
                     ' the input stream
    
'        Stop
   
   
   End Select

If pf = 0 Then

    Label1 = " - " + Format(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " - " + Format(Timer, "0.000")
    List1.AddItem "MSComm1_OnComm -- UnKonwn Event " + Str(MSComm1.CommEvent) + Label1
    'Stop

End If

List1.ListIndex = List1.ListCount - 1

End Sub

Sub FindTimeInfo()

'InPutDate = DateValue("01-12-2008")
'TestDate = DateValue("31-05-2009")

If InPutDate = 0 Then Stop

DayCountT = DateDiff("d", InPutDate, TestDate)

Year5 = 0
For r5 = Year(InPutDate) + 1 To Year(TestDate) + 1
    If DateSerial(r5, Month(InPutDate), Day(InPutDate)) < Now Then Year5 = Year5 + 1
Next
For r5 = 1 To -2 Step -1
    XX = DateSerial(Year(TestDate), Month(TestDate) + r5, Day(InPutDate))
    MonthStart = XX
    If XX <= TestDate Then Exit For
Next
Month5 = 0
Month7 = 0
For r5 = 1 To 10000
    XX = DateSerial(Year(InPutDate), Month(InPutDate) + r5, Day(InPutDate))
    If Year(XX) <> oxx Then Month7 = 0
    oxx = Year(XX)
    If XX <= TestDate Then Month5 = Month5 + 1: Month7 = Month7 + 1
    If XX >= TestDate Then Exit For
Next

For r5 = Year(TestDate) To 1 Step -1
    If DateSerial(r5, Month(InPutDate), Day(InPutDate)) < Now Then
        DayCount1 = DateDiff("d", DateSerial(r5, Month(InPutDate), Day(InPutDate)), TestDate)
        DayCountMonth = DateDiff("d", MonthStart, TestDate)
        WholeYear1 = DateDiff("d", DateSerial(r5, Month(InPutDate), Day(InPutDate)), DateSerial(r5 + 1, Month(InPutDate), Day(InPutDate)))
        'Month5 = DateDiff("m", DateSerial(r5, Month(InPutDate), Day(InPutDate)), TestDate)
        WeeksSinceYear = DateDiff("w", DateSerial(r5, Month(InPutDate), Day(InPutDate)), TestDate) - 1
        WeeksSinceInput = DateDiff("w", InPutDate, TestDate)
        WeeksPlusDays = DateDiff("d", InPutDate, TestDate) - (DateDiff("w", InPutDate, TestDate) * 7)
                        
                
        ResultYearDate = Format$(Year5 + DayCount1 / WholeYear1, ".000")
        Exit For
    End If
Next


tMm = DateDiff("s", InPutDate, TestDate)

F1 = Int((tMm / 60 ^ 2) / 24)

If F1 > 0 Then tMm = tMm - (DateDiff("s", Now, Now + 1) * F1)

F2 = Int((tMm / 60 ^ 2))
F3 = Int(tMm / 60)
F4 = tMm Mod 60

'EXAMPLE
Dex = Trim(Str(F1)) + " Days " + Format$(TimeSerial(0, F3, F4), "H:MM:SS")



End Sub


Sub DealDoor()

If Timer7TestVar = True Then Exit Sub

CTSHoldingvar = MSComm1.CTSHolding

If OCTSHoldingVAR <> CTSHoldingvar Then
    
    If CTSHoldingvar = False Then
        Status = "Door Open": FlagOpen = True
        Status2 = "Open"
        If RUNME2 <> -2 Then List1.AddItem ""
    Else
        Status2 = "Closed"
        Status = "Door Closed":
    End If



    If ONow = 0 Then
        TSH = App.Path + "\RS232 DOOR REED DATA\LAST DOOR CLOSE.TXT"
        Open TSH For Input As #1
            Do
            Line Input #1, LLL2
            If LLL2 <> "" Then LLL = LLL2
            Loop Until EOF(1)
        Close #1
    
        ONow = DateValue(LLL) + TimeValue(LLL)
        HAONow = ONow
    Else
    
        'If Label2 <> "*The One* Are *IN*" Then ONow2 = Now
        
        HAONow = ONow
        ONow = Now
    
    End If
        
    'ONow = Now
    'XONow = ONow
    
    'TEMP WORKAROUND IN CASE = 0
    If HAONow = 0 Then HAONow = Now
    
    InPutDate = HAONow
    TestDate = Now
    Call FindTimeInfo
    
    'If RUNME2 <> -2 Then ONow = D
    
    
    If CTSHoldingvar = False And AreYouIn = False Then
        'Me.WindowState = 0
    End If
    
    Dex = Trim(Str(F1)) + " Days " + Format$(TimeSerial(0, F3, F4), "H:MM:SS")
    TICK2 = Dex
    'FindTimeInfo
    TICK = Format(CountD, "0000") + " -- " + Format(ONow, "DDD DD-MMM-YY HH:MM:SSa/p") + "   " + TICK2 + " - " + Status
    
    If RUNME2 = -2 Then TICK = TICK + " -- Start Up Result": RUNME2 = 0
    
    List1.AddItem TICK

    If CTSHoldingvar = False Then CountD = CountD + 1

    
    If RUNME <> -2 Then
        WAWIN = FindWindow("Winamp v1.x", vbNullString)
        If WAWIN <> 0 Then messageResult = SendMessage(WAWIN, WM_USER, 0, ByVal IPC_ISPLAYING)

        'Public Sub CommandPause()
        If messageResult = 1 And WAWIN <> 0 And IsIDE = False Then SendMessage WAWIN, WM_COMMAND, WINAMP_BUTTON3, ByVal 0

        Sleep 400
    End If
    
    RUNME = 0
    MMControl2.Command = "prev"
    MMControl2.Command = "Play"
    
End If

OCTSHoldingVAR = CTSHoldingvar

'MSComm1.RTSEnable = True

End Sub


'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************

Private Sub Timer1_Timer()


If ONow > 0 And Timer1.Enabled = True Then
    InPutDate = ONow
    TestDate = Now
    Call FindTimeInfo
    Dex = Trim(Str(F1)) + " Days " + Format$(TimeSerial(0, F3, F4), "H:MM:SS")
    Label5 = "Since Last Door Close - " + Dex

    'TEMP WORK AROUND
    HighNow = Now 'CHK SAFE IN CASE ONOW = 0
    If DateDiff("s", ONow, Now) > HighNow2 Then
        HighNow = ONow: HighNow2 = DateDiff("s", ONow, Now)
    Else
        
        TSX = App.Path + "\" + GetComputerName + "\RS232 DOOR REED DATA\Last Door Lenght Time Logg.txt"
        
        'Call FileInUseDelay2(TSX)
    
        Open TSX For Append As #1
            Print #1, Format(HighNow, "DDD dd-mm-yyyy hh:mm:ss") + " --- " + Format(Now, "DDD dd-mm-yyyy hh:mm:ss");
            InPutDate = HighNow
            TestDate = Now
            Call FindTimeInfo
            Dex = " --- " + Trim(Str(F1)) + " Days " + Format$(TimeSerial(0, F3, F4), "H:MM:SS") + " -- " + Status2
            Print #1, Dex
        Close #1
        
        HighNow = 0
        HighNow2 = 0
    
    End If

End If

If List1.ListCount <> OldList1Count Then
    
   If XONow = 0 And Timer1.Enabled = True Then
        TSX = App.Path + "\RS232 DOOR REED DATA\LAST DOOR CLOSE.txt"
        If FileInUse(TSX) = True Then Exit Sub
    
        Open TSX For Append As #1
            Print #1, Format(ONow, "dd-mm-yyyy hh:mm:ss")
        Close #1
    End If
    
    'If XONow = -1 Then XONow = 0
    If XONow = -2 And Timer1.Enabled = True Then XONow = 0
    
    If FileInUse(Tx$) = True Then Exit Sub
    
    fr1 = FreeFile
    Open TS For Append Lock Read Write As #fr1
        For r = OldList1Count To List1.ListCount - 1
            Print #fr1, List1.List(r)
        Next
    Close #fr1
End If


OldList1Count = List1.ListCount

End Sub

Private Sub Timer2_Timer()

Exit Sub

MSComm1.RTSEnable = False
DoEvents
Sleep 50

MSComm1.PortOpen = False
DoEvents
Sleep 50
Do
DoEvents
Sleep 50
Loop Until MSComm1.PortOpen = False

MSComm1.PortOpen = True

Do
    DoEvents
    Sleep 50
Loop Until MSComm1.PortOpen = True
    
    
DoEvents
MSComm1.RTSEnable = True

Do
    DoEvents
    Sleep 50
Loop Until MSComm1.PortOpen = True


End Sub


Private Sub Timer3_Timer()
             
If Timer7TestVar = True Then Exit Sub
             
             
If ONow + TimeSerial(0, 2, 0) < Now And FlagOpen = True Then

        Label2 = "*The One* Are *Out*": AreYouIn = False
        'Me.WindowState = 0

End If
             
On Error Resume Next
             
Buffer = Buffer & MSComm1.Input

If Buffer <> "" Then
    Buffer = ""
    Label1 = "RS232 Armed and Very Very Dangerous ** " + Format(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " **"
End If

MSComm1.Output = "A"


End Sub

Private Sub Timer4_Timer()
Exit Sub

If Second(Now) Mod 2 <> 0 Then Exit Sub

MSComm1.RTSEnable = Not MSComm1.RTSEnable
        
Count2 = Count2 + 1
Label1 = " - " + Format(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " - " + Format(Timer, "0.000")
List1.AddItem Format(Count2, "00000 ") + "Toggle RTSEnable line -- " + Str(MSComm1.RTSEnable) + Label1

List1.ListIndex = List1.ListCount - 1


End Sub


Private Sub Timer5_Timer()

Timer5.Interval = 50

If MSComm1.PortOpen = False Then
    MSComm1.PortOpen = True
    MSComm1.RTSEnable = True
    Count1 = 0
End If


'If Second(Now) Mod 2 <> 0 Then Exit Sub

MSComm1.RTSEnable = Not MSComm1.RTSEnable

If Count1 = 0 Then Count2 = 0
If Count2 > Count1 Then Count2 = Count1

Label1 = "Very Great Golden Handshaking - Test =" + Str(Count1) + " Result =" + Str(Count2)

Count1 = Count1 + 1
        
If Count1 >= 4 And Count2 < Count1 - 2 Then
        RSXErrors = RSXErrors + 1
        Label6 = "RSX #" + Trim(Str(RSXErrors)) + Format(Now, " ddd hh:mm:ssa/p")


        MSComm1.PortOpen = False
        Exit Sub
'    HandShakePass = False
'    Timer5.Enabled = False
End If

If Count2 >= 4 Then
    HandShakePass = True
    Timer5.Enabled = False
End If
    


'Label1 = " - " + Format(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " - " + Format(Timer, "0.000")
'List1.AddItem Format(Count2, "00000 ") + "Toggle RTSEnable line -- " + Str(MSComm1.RTSEnable) + Label1



List1.ListIndex = List1.ListCount - 1


End Sub

Private Sub Timer7Test_Timer()
'    Timer7Test.Enabled = False

'    Exit Sub

If MSComm1.PortOpen = False Then
    MSComm1.PortOpen = True
    MSComm1.RTSEnable = True
    Timer7TestVar = False
    Ticker1 = 0
    LowTest = 0
End If

If LowTest > 0 Then
    
    If LowTest = 1 And MSComm1.CTSHolding <> False Then
        MSComm1.PortOpen = False
        Exit Sub
    End If
    
    If LowTest = 2 Then
        Ticker1 = 0
        Label4 = "#"
        MSComm1.RTSEnable = True
    
    End If
    If LowTest = 3 Then
        LowTest = 0
        Ticker1 = 0
        Label4 = "0"
        Timer7TestVar = False
        Timer7Test.Interval = 250
    
    End If

    If LowTest = 2 Then LowTest = 3


    If Ticker1 > 15 Then ' SET HERE MORE = LESS CHANCE ERROR BUT LESS SECURE
        'Label6 = Trim(Str(Val(Label6) + 1))
        RSXErrors = RSXErrors + 1
        Label6 = "RSX #" + Trim(Str(RSXErrors)) + Format(Now, " ddd hh:mm:ssa/p")
        MSComm1.PortOpen = False
        Exit Sub
    End If


End If


Ticker1 = Ticker1 + 1



If Ticker1 > 80 Then ' AND HERE - SET HERE MORE = LESS CHANCE ERROR BUT LESS SECURE
    'Timer7Test.Enabled = False
    MSComm1.PortOpen = False
    Exit Sub

    'Style = vbYesNo + vbCritical + vbDefaultButton2 + vbMsgBoxSetForeground
    'Response = MsgBox("NOT Working " + vbCrLf + Format$(Now, "DD-MMM HH:MM:SS") + vbCrLf + "Description " + Err.Description + vbCrLf + "Err Number" + Str(Err.Number) + vbCrLf + "Errr Source " + Err.Source + vbCrLf + "STOP YES/NO", Style)
    'If Response = vbYes Then Stop
End If



If Timer7TestVar = True Then Exit Sub

If MSComm1.CTSHolding = False Then Exit Sub


If Ticker1 >= 100 Then
    Label4 = "100": Label4.BackColor = &HFF&
    'Beep
End If

SETTEST = 50

If Ticker1 < 100 Then Label4 = Format(SETTEST + 1 - Ticker1, "0")

If Ticker1 > SETTEST Then
    
    LowTest = 1
    Label4 = "*"
    Timer7Test.Interval = 1
    Timer7TestVar = True
    MSComm1.RTSEnable = False

End If

End Sub





Private Sub TimerControl_Timer()

If HandShakePass > -2 Then
Timer5.Enabled = False

If HandShakePass = False Then
    Lel1 = " - " + Format(Now, "DDD DD-MMM-YY HH:MM:SSa/p")
    List1.AddItem "Great - Golden Handshake Test -- Failed" + Lel1
    Style = vbYesNo + vbCritical + vbDefaultButton2 + vbMsgBoxSetForeground
    Response = MsgBox("Golden Handshake Test -- Failed" + vbCrLf + "STOP YES/NO", Style)
    If Response = vbYes Then Stop: End
    TimerControl.Enabled = False
End If


If HandShakePass = True Then
    List1.AddItem "The Very Great Golden Handshaking Test -- Result Art -- Awesome"
    Count2 = 0
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer5.Enabled = False
    Timer7Test.Enabled = True
    TimerControl.Enabled = False
    MSComm1.RTSEnable = True


End If
End If


End Sub
