VERSION 5.00
Begin VB.Form AI 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Active Idle..."
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   795
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer11 
      Interval        =   100
      Left            =   6810
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   5775
      Top             =   1635
   End
   Begin VB.Timer Timer3 
      Interval        =   50
      Left            =   5730
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5760
      Top             =   270
   End
   Begin VB.Label Label23 
      Caption         =   "23"
      Enabled         =   0   'False
      Height          =   270
      Left            =   6990
      TabIndex        =   15
      Top             =   2145
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label21 
      Caption         =   "21"
      Enabled         =   0   'False
      Height          =   270
      Left            =   6960
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   780
      TabIndex        =   13
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "0"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   4680
      TabIndex        =   12
      Top             =   2610
      Width           =   1965
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Mouse Clicks"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   3120
      TabIndex        =   11
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   2610
      Width           =   1545
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Idle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   780
      TabIndex        =   10
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   360
      Width           =   2100
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   780
      TabIndex        =   9
      Top             =   720
      Width           =   2100
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Active"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2895
      TabIndex        =   8
      ToolTipText     =   "View keyboard and Mouse Log With Notepad"
      Top             =   360
      Width           =   2100
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2895
      TabIndex        =   7
      ToolTipText     =   "F12 to Command Prompt - Disable"
      Top             =   720
      Width           =   2100
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Now"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   780
      TabIndex        =   5
      Top             =   1440
      Width           =   2100
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2895
      TabIndex        =   4
      ToolTipText     =   "F12 to Command Prompt - Disable"
      Top             =   1440
      Width           =   2100
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Tott"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   1440
      Width           =   765
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Day"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2895
      TabIndex        =   1
      ToolTipText     =   "F12 to Command Prompt - Disable"
      Top             =   1080
      Width           =   2100
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   780
      TabIndex        =   0
      Top             =   1080
      Width           =   2100
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB"
   End
   Begin VB.Menu Mnu_File 
      Caption         =   "File"
      Begin VB.Menu Mnu_Now2 
         Caption         =   "Now2 Logg"
      End
      Begin VB.Menu Menu_EndDay 
         Caption         =   "End Of Day Logg"
      End
   End
   Begin VB.Menu Mnu_IdleOver 
      Caption         =   "Idle Over Time"
   End
   Begin VB.Menu MNU_LOGGFOLDER 
      Caption         =   "Logg Folder"
   End
End
Attribute VB_Name = "AI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OXXA, OHwndw, HALO, HALO2, LockMove, RamDrive, FS, OHwnd
Dim TTGIG
Private Declare Function GetForegroundWindow Lib "user32" () As Long
'Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'------------------
'Required for
'Function FileInUse(ByVal strFilePath As String) As Boolean

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public MouseFirstRun As Integer, OldWinStat
Public KeyHittsOld, Keyy

Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40

Private Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

'--------------


Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uaction&, ByVal uparam&, ByRef lpvParam As Any, ByVal fuwinini&) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long   'Parameter: form.(property) = SetCursorPos(x,y) in Pixels.  'MODULE 1121

Public Lk1, Lk2, Lk3, Lk4, Lab2$

Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Type POINTAPI
        x As Long
        Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public NotePad2RT2
Public VolumeSet1, VolumeSet2, OldMsgResult, TimeDelay, OldVolumeSet
Public IdleOnNday1$, IdleOnNday2$
Public Timer11Var, AppTitle, EasyNow

Public QQX1 As Long, QQD1 As Long, QQN1 As Long
Public QQX2 As Long, QQD2 As Long, QQN2 As Long
Public QQ91 As Date, QQT1 As Date, QQZ1 As Date
Public QQ92 As Date, QQT2 As Date, QQZ2 As Date
Public OldDayNow
Public IdleMin1 As Date, IdleMin2 As Date, TTQ, SetFlagZero5, SetFlagZeroDay
Public WExpand, MMove2, MMove2Time, Mouse10CLicks, MouseClicks, MainIdleTime, IsOldIdleTime
Public IdleOff As Date, IdleOn As Date, OldMark, OldMark2
Public IdleOffT As Date, IdleOnT As Date
Public IdleOnN1$, IdleOnN2$, Fr0$, IdleOnN3$, IdleOnN4$
Public IdleOnN5$, IdleOnN6$
Public ThisSysUptime3 As Date
Public IdleT As Date, IdleS As Date, IdleMin As Date, IdleT2 As Date, IdleT2Flag

Public Xmp5 As Long
Public Ymp5 As Long

Private Sub Form_Click()
End
End Sub

Private Sub Form_Load()
OHwnd = 20

Lk4 = -1
wc = 0: hc = 0
Me.BackColor = 0

Set FS = CreateObject("Scripting.FileSystemObject")

On Local Error Resume Next
For Each Control In AI
    If Control.Enabled = True Then
        If Control.Left + Control.Width > wc Then wc = Control.Left + Control.Width
        If Control.Top + Control.Height > hc Then hc = Control.Top + Control.Height
    End If
Next

AI.Width = wc + 90 + hq
AI.Height = hc + 700

IdleOn = 0
IdleOff = 0
RKeyPlus = 99

On Error Resume Next
FR = FreeFile
Open App.Path + "\00_Text_Data_02\Idle and Active Logg Day.txt" For Input As #FR
Line Input #FR, a1$
Line Input #FR, a2$
Line Input #FR, a3$
Close #FR
If Err.Number > 0 Then
    On Error GoTo 0
    FR = FreeFile
    Open App.Path + "\00_Text_Data_02\Idle and Active Logg Day.bak.txt" For Input As #FR
    Line Input #FR, a1$
    Line Input #FR, a2$
    Line Input #FR, a3$
    Close #FR
End If
On Error GoTo 0


If DateValue(a1$) = Int(Now) Then
QQN1 = Val(a2$)
QQN2 = Val(a3$)
End If


On Error Resume Next
Open App.Path + "\00_Text_Data_02\Idle and Active Logg Tot.txt" For Input As #1
Line Input #1, a1$
Line Input #1, a2$
Line Input #1, a3$
Close #1
If Err.Number > 0 Then
    On Error GoTo 0
    FR = FreeFile
    Open App.Path + "\00_Text_Data_02\Idle and Active Logg Tot.bak.txt" For Input As #FR
    Line Input #FR, a1$
    Line Input #FR, a2$
    Line Input #FR, a3$
    Close #FR
End If
On Error GoTo 0

QQD1 = Val(a2$)
QQD2 = Val(a3$)



On Local Error Resume Next
Set dc = FS.Drives
For Each D In dc
    dr = D.DriveLetter
    n = D.VolumeName
    If InStr(n, "RAM") > 0 Then Exit For
Next
On Local Error GoTo 0
RamDrive = dr


Me.Show
Me.WindowState = 1

End Sub

Private Sub Form_Resize()
Call Resize_Code

End Sub

Public Sub Resize_Code()
    Call Timer1_Timer
End Sub

Private Sub Label2_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\DARK\DARK.exe", vbNormalFocus

End Sub

Private Sub Label5_Click()
Call ATidalX.Mnu_Idel_Click
End Sub

Private Sub Label7_Click()
Call ATidalX.Mnu_Idle_Now_Click
End Sub

Private Sub Menu_EndDay_Click()
Shell "notepad " + App.Path + "\00_Text_Data_02\Idle and Active Logger End Day.txt", vbNormalFocus

End Sub

Private Sub Mnu_IdleOver_Click()
Shell "notepad " + App.Path + "\00_Text_Data_02\Idle and Active Logger Idle Over Time.txt", vbNormalFocus

'Shell "D:\VB6\VB-NT\00_Best_VB_01\NotePad5\Notepad.exe " + App.Path + "\00_Text_Data_02\Idle and Active Logger Idle Over Time.txt", vbNormalFocus

End Sub

Private Sub MNU_LOGGFOLDER_Click()
Shell "EXPLORER " + App.Path, vbNormalFocus
End Sub

Private Sub Mnu_Now2_Click()
Shell "notepad " + App.Path + "\00_Text_Data_02\Idle and Active Logger Now2.txt", vbNormalFocus

End Sub

Private Sub MNU_VB_Click()

If IsIDE = True Then Stop: End

Me.Visible = False
Me.Enabled = False
Me.WindowState = 1
'Unload Me
Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
Unload Me
End

End Sub

Private Sub Timer1_Timer()

qq8 = GetIdleM(5)




If TTGIG > 5 Then Exit Sub
TTGIG = TTGIG + 1

Rf = FindWindow(vbNullString, "Drives_Gig")
If Rf = 0 Then
    Shell "D:\VB6\VB-NT\00_Best_VB_01\Drives Gig\Drives_Gig.exe", vbNormalNoFocus
End If

End Sub

Function IsIdleTime(iq)
IsIdleTime = False
If MainIdleTime = 0 Then Exit Function
If DateDiff("s", MainIdleTime, Now) > iq Then IsIdleTime = True

End Function

Function GetIdleM(Mark3)

If IdleMin1 = 0 Then
    IdleMin1 = Now
    IdleMin2 = Now
    syss = 1
End If

RR = IsIdleTime(1)

If RR = True Then
    qq3 = DateDiff("s", IdleMin1, Now)
    QQX1 = QQX1 + qq3
    QQD1 = QQD1 + qq3
    QQN1 = QQN1 + qq3
End If
If RR = False Then
    qq3 = DateDiff("s", IdleMin2, Now)
    QQX2 = QQX2 + qq3
    QQD2 = QQD2 + qq3
    QQN2 = QQN2 + qq3
    AI.Label33.BackColor = &H800000
    AI.Label8.BackColor = &H800000
    AI.Label1.BackColor = &H800000
    AI.Label35.BackColor = &H400000
    AI.Label10.BackColor = &H400000
    AI.Label3.BackColor = &H400000
Else
    AI.Label35.BackColor = &H800000
    AI.Label10.BackColor = &H800000
    AI.Label3.BackColor = &H800000
    AI.Label33.BackColor = &H400000
    AI.Label8.BackColor = &H400000
    AI.Label1.BackColor = &H400000
End If

qq3 = QQX1
qq4 = qq3 Mod 60 'seconds
qq5 = (qq3 \ 60) Mod 60 'Minutes
qq6 = (qq3 \ (60 ^ 2)) Mod 24 'Hours
qq7 = (qq3 \ (24 * (60 ^ 2))) 'Days
qq9 = qq7 + TimeSerial(qq6, qq5, qq4)

qq3 = QQX2
qq4 = qq3 Mod 60 'seconds
qq5 = (qq3 \ 60) Mod 60 'Minutes
qq6 = (qq3 \ (60 ^ 2)) Mod 24 'Hours
qq7 = (qq3 \ (24 * (60 ^ 2))) 'Days
qq10 = qq7 + TimeSerial(qq6, qq5, qq4)

QQ91 = qq9
QQ92 = qq10

qq3 = QQD1
qq4 = qq3 Mod 60 'seconds
qq5 = (qq3 \ 60) Mod 60 'Minutes
qq6 = (qq3 \ (60 ^ 2)) Mod 24 'Hours
qq7 = (qq3 \ (24 * (60 ^ 2))) 'Days
qq9 = qq7 + TimeSerial(qq6, qq5, qq4)

qq3 = QQD2
qq4 = qq3 Mod 60 'seconds
qq5 = (qq3 \ 60) Mod 60 'Minutes
qq6 = (qq3 \ (60 ^ 2)) Mod 24 'Hours
qq7 = (qq3 \ (24 * (60 ^ 2))) 'Days
qq10 = qq7 + TimeSerial(qq6, qq5, qq4)

QQT1 = qq9
QQT2 = qq10

qq3 = QQN1
qq4 = qq3 Mod 60 'seconds
qq5 = (qq3 \ 60) Mod 60 'Minutes
qq6 = (qq3 \ (60 ^ 2)) Mod 24 'Hours
qq7 = (qq3 \ (24 * (60 ^ 2))) 'Days
qq9 = qq7 + TimeSerial(qq6, qq5, qq4)

qq3 = QQN2
qq4 = qq3 Mod 60 'seconds
qq5 = (qq3 \ 60) Mod 60 'Minutes
qq6 = (qq3 \ (60 ^ 2)) Mod 24 'Hours
qq7 = (qq3 \ (24 * (60 ^ 2))) 'Days
qq10 = qq7 + TimeSerial(qq6, qq5, qq4)

QQZ1 = qq9
QQZ2 = qq10

IdleMin1 = Now
IdleMin2 = Now

IdleOnN7$ = IdleOnN1$
IdleOnN8$ = IdleOnN2$
IdleOnN9$ = IdleOnN3$

IdleOnN10$ = IdleOnN4$
IdleOnN11$ = IdleOnN5$
IdleOnN12$ = IdleOnN6$

If IsOldIdleTime <> RR Then
    TTQ = Now + TimeSerial(0, 0, 5)
End If

IdleOvertime = TimeSerial(0, 10, 0)
idleover15mins = 0
If QQ91 > IdleOvertime Then
    idleover15mins = QQ91
End If

If IsOldIdleTime <> RR And SetFlagZero5 = True Then
    SetFlagZero5 = False
    QQX1 = 0: QQ91 = 0
    QQX2 = 0: QQ92 = 0
    flagextra = True
    'IdleOnN7$ = IdleOnN1$
End If

'SetFlagZeroDay = True
flag2extra = False
If IsOldIdleTime <> RR And SetFlagZeroDay = True Then
    SetFlagZeroDay = False
    QQN1 = 0 ': QQZ1 = 0
    QQN2 = 0 ': QQZ2 = 0
    flagextra = True
    flag2extra = True
    IdleOnNday1$ = AI.Label35 'idle
    IdleOnNday2$ = AI.Label33 'active
End If

If Day(Now) <> OldDayNow And OldDayNow > 0 Then
    SetFlagZeroDay = True
End If

OldDayNow = Day(Now)

If TTQ < Now And TTQ > 0 Then
    SetFlagZero5 = True
End If

IsOldIdleTime = RR
    AI.Label2 = Format$(MainIdleTime, "DDD DD MMM YY -- HH:MM:SSa/p")
    

    If RR = -1 Or IdleOnN1$ = "" Then IdleOnN1$ = Format$(Int(QQ91), "0d ") + Format$(QQ91, "HH:MM:SS")
    IdleOnN2$ = Format$(Int(QQT1), "0d ") + Format$(QQT1, "HH:MM:SS")
    IdleOnN3$ = Format$(Int(QQZ1), "0d ") + Format$(QQZ1, "HH:MM:SS")
    If RR = -1 Then AI.Label10 = Format$(QQ91, "HH:MM:SS") 'Top
    AI.Label35 = Format$(QQZ1, "HH:MM:SS") 'Middle
    AI.Label3 = IdleOnN2$ 'Bottom
    
    If RR = 0 Or IdleOnN4$ = "" Then
        IdleOnN4$ = Format$(Int(QQ92), "0d ") + Format$(QQ92, "HH:MM:SS")
    End If
    IdleOnN5$ = Format$(Int(QQT2), "0d ") + Format$(QQT2, "HH:MM:SS")
    IdleOnN6$ = Format$(Int(QQZ2), "0d ") + Format$(QQZ2, "HH:MM:SS")
    If RR = 0 Then AI.Label8 = Format$(QQ92, "HH:MM:SS")
    AI.Label33 = Format$(QQZ2, "HH:MM:SS")
    AI.Label1 = IdleOnN5$

hot = 0

If Mark3 = 5 And flagextra = False Then Exit Function

If Mark3 > 2 Or syss = 1 Then flagextra = True

If Mark3 = 4 Then Mark3 = OldMark2: hot = 1
If RR = True Then Desc$ = " -- Idle"
If RR = False Then Desc$ = " -- Active"
If Mark3 = 3 Then Desc$ = " -- Program Close"
If hot = 1 Then Desc$ = " -- End Of Day"

If syss = 1 Then
    Desc$ = " -- Progam Start"
    IdleOnN7$ = IdleOnN1$
    IdleOnN8$ = IdleOnN2$
    IdleOnN9$ = IdleOnN3$

    IdleOnN10$ = IdleOnN4$
    IdleOnN11$ = IdleOnN5$
    IdleOnN12$ = IdleOnN6$
End If


If Mark3 = 3 Then Mark3 = 1

'If IdleOnN1$ = "" Then syss = 0: Exit Function
If idleover15mins > 0 And RR = False Then
    IdleOnNx1$ = Format$(Int(idleover15mins), "0d ") + Format$(idleover15mins, "HH:MM:SS")
    fr40$ = Format$(Now - idleover15mins, "DDD DD-MM-YY HH:MM:SS ") + " --- "
    fr40$ = fr40 + Format$(Now, "DD-MM-YY HH:MM:SS ") + " --- "
    fr40$ = fr40$ + IdleOnNx1$ + " "
    fr40$ = fr40$ + "Idle " 'Over Time = " + Format$(IdleOvertime, "HH:MM:SS")
    
    tx2$ = App.Path + "\00_Text_Data_02\Idle and Active Logger Idle Over Time.txt"
    Call FileInUseDelay(tx2$)
    FR = FreeFile
    Open tx2$ For Append As #FR
    Print #FR, fr40$
    Close #FR
End If

'---------------------------------
Exit Function

If flag2extra = True Then
    fr40$ = Format$(Now, "DD-MM-YYYY HH:MM:SS ") + " --- "
    fr40$ = fr40$ + IdleOnNday1$ + " -- Idle -- "
    fr40$ = fr40$ + IdleOnNday2$ + " -- Active -- Start New Day"
    tx2$ = App.Path + "\00_Text_Data_02\Idle and Active Logger End Day.txt"
    Call FileInUseDelay(tx2$)
    FR = FreeFile
    Open tx2$ For Append As #FR
    Print #FR, fr40$
    Close #FR

End If

If flagextra = True Then

Fr0$ = Format$(Now, "DD-MM-YYYY HH:MM:SS ") + " --- "
Fr0$ = Fr0$ + IdleOnN7$ + " -- Idle -- "
Fr0$ = Fr0$ + IdleOnN10$ + " -- Active -- " + Desc$

Frx0$ = Fr0$ + Format$(Nx, " 000000") + Format$(Ny, " 000000") + Str(Int(MouseClicks))

Fr1$ = Format$(Now, "DD-MM-YYYY HH:MM:SS ") + " --- "
Fr1$ = Fr1$ + IdleOnN8$ + " -- Idle -- "
Fr1$ = Fr1$ + IdleOnN11$ + " -- Active -- " + Desc$

'On Local Error GoTo ErrHandle
tx2$ = App.Path + "\00_Text_Data_02\Idle and Active Logger Now.txt"
Call FileInUseDelay(tx2$)
FR = FreeFile
Open tx2$ For Append As #FR
Print #FR, Fr0$
Close #FR

tx2$ = App.Path + "\00_Text_Data_02\Idle and Active Logger Now2.txt"
Call FileInUseDelay(tx2$)
FR = FreeFile
Open tx2$ For Append As #FR
Print #FR, Frx0$
Close #FR

tx2$ = App.Path + "\00_Text_Data_02\Idle and Active Logger Day.txt"
Call FileInUseDelay(tx2$)
FR = FreeFile
Open tx2$ For Append As #FR
Print #FR, Fr1$
Close #FR

End If

If Second(Now) <> OldSecond Then
OldSecond = Second(Now)
'On Local Error GoTo ErrHandle
tx3$ = App.Path + "\00_Text_Data_02\Idle and Active Logg Day.bak.txt"
tx2$ = App.Path + "\00_Text_Data_02\Idle and Active Logg Day.txt"
Call FileInUseDelay(tx3$)
FS.copyFile tx2$, tx3$
Call FileInUseDelay(tx2$)
FR = FreeFile
Open tx2$ For Output As #FR
Print #FR, Int(Now)
Print #FR, QQN1
Print #FR, QQN2
Close #FR

tx2$ = App.Path + "\00_Text_Data_02\Idle and Active Logg Tot.txt"
tx3$ = App.Path + "\00_Text_Data_02\Idle and Active Logg Tot.bak.txt"
Call FileInUseDelay(tx3$)
FS.copyFile tx2$, tx3$
Call FileInUseDelay(tx2$)
FR = FreeFile
Open tx2$ For Output As #FR
Print #FR, Int(Now)
Print #FR, QQD1
Print #FR, QQD2
Close #FR
End If

OldMark2 = Mark3

Exit Function
ErrHandle:
DoEvents
'Resume

End Function


Private Sub Label21_Change()
'Mouse
Vcodez = 0
Call KeyOrMouse(1)

End Sub

Private Sub Label23_Change()
'Key
KeyIdleTime = Now + TimeSerial(0, 0, 30)

'If Dir$("D:\TEMP\KeyHit-Tidal-VCodes.txt") = "" Then popk = 1 Else popk = 0

Call KeyOrMouse(2)

End Sub


Sub KeyOrMouse(KeyMouse)

If Lab2$ = Label21.Caption + Label23.Caption Then Exit Sub
Lab2$ = Label21.Caption + Label23.Caption


If KeyMouse = 1 Then
    AI.Label13 = Int(MouseClicks)
End If
gogo = False
If KeyMouse = 1 And MouseClicks > 3 Then gogo = True
If KeyMouse = 2 Then gogo = True

If MainIdleTime = 0 Then MainIdleTime = Now

If gogo = True Then
    
    'never equal zero except start
    'if it does it put back up before get here again
    'If KeyMouse = 2 Then
    qq8 = GetIdleM(2)
    'End If
    MainIdleTime = Now
    'easy = True
    
    'Test7 = 0
    'Call SetLockMax
    'Call ProgressLock
End If

End Sub


Public Function GetWindowState(ByVal lngHwnd As Long) As Integer
    
    If IsWindow(lngHwnd) = 1 Then
        GetWindowState = -1
        If IsIconic(lngHwnd) <> 0 Then
            GetWindowState = vbMinimized
        ElseIf IsZoomed(lngHwnd) <> 0 Then
            GetWindowState = vbMaximized
        End If
    End If

End Function

Public Sub FindCursor(x, Y)

Dim P As POINTAPI

GetCursorPos P
'   return x and y co-ordinate
x = P.x ' / GetSystemMetrics(0) * Screen.Width
'   for current cursor position
Y = P.Y '/ GetSystemMetrics(1) * Screen.Height

End Sub


Private Sub MouseDetectMove()

FindCursor Nx, Ny

'This Will Happen to Mouse If User Is Logged off
'If Nx = 0 And Ny = 0 Then
'LockSSaver = Now + LockSaverDelay
'Winamp Video
'Call SetLockMax
'Call ProgressLock
'End If

If Nx = 0 Or Ny = 0 Then Exit Sub


'If Quake = 0 Then
    If (Nx <> Xmp5 Or Ny <> Ymp5) And (Nx <> ScreenWidth And Ny <> ScreenHeight) Then
        'Call WinonPoint
        Mousey = Mousey + Sqr(Abs(Xmp5 - Nx) ^ 2 + Abs(Ymp5 - Ny) ^ 2)
        MouseClicks = Sqr(Abs(Xmp5 - Nx) ^ 2 + Abs(Ymp5 - Ny) ^ 2)
                    
        If MouseFirstRun = 0 Then
            MouseFirstRun = 1
            Mousey = 0
        End If
        
        If MouseClicks > 3 Then
            Label21.Caption = Str$(Mousey)
        End If
        
        Xmp5 = Nx: Ymp5 = Ny
    
    End If
'End If

End Sub


Private Sub Timer2_Timer()

Dim Hwndw As Long
Dim HwndW2 As Long

Dim Rect4 As RECT
Dim Rect5 As RECT


Hwndw = FindWindow(vbNullString, "Tidal Information...")



If Hwndw = 0 And OHwnd <> 0 Then
    'If Me.WindowState = 0 Then Me.WindowState = 1
End If

If Hwndw <> 0 And OHwnd = 0 Then
    'If Me.WindowState = 1 Then Me.WindowState = 0
End If

OHwnd = Hwndw


If GetWindowState(Hwndw) <> -1 Then Exit Sub

If (GetAsyncKeyState(1) And GetForegroundWindow = Me.hwnd) Then
    LockMove = Now + Timer + 0.1
     Exit Sub
End If
   
If LockMove > Now + Timer Then
     Exit Sub
End If



If OHwndw <> Hwndw Then settid = 1
OHwndw = Hwndw
If Hwndw > 0 Then
Hwnd24 = GetWindowRect(Hwndw, Rect4)
On Local Error Resume Next

    If AI.Left <> (Rect4.Left * Screen.TwipsPerPixelY - AI.Width) Then
            AI.Left = (Rect4.Left * Screen.TwipsPerPixelY - AI.Width)
            Hwndw = FindWindow(vbNullString, "BBC 24 Hour Weather")
            Hwnd24 = GetWindowRect(Hwndw, Rect5)
            d1 = Rect5.Bottom * Screen.TwipsPerPixelX
            d2 = Rect4.Top * Screen.TwipsPerPixelX
            If d1 > d2 Then d2 = d1
            AI.Top = d2
    End If

If Lk2 <> Rect4.Top + RECT.Left And Rect4.Top > 0 And WindowVisible(Hwndw) = True Then
    If Rect4.Top <> 400 And Rect4.Left <> 640 Then
    
        Hwndw = FindWindow(vbNullString, "BBC 24 Hour Weather")
        xxa = GetWindowState(Hwndw)
        If Hwndw > 0 And xxa <> -1 Then
        AI.Top = (Rect4.Top * Screen.TwipsPerPixelX)
        End If
        If Hwndw = 0 Then
            AI.Top = (Rect4.Top * Screen.TwipsPerPixelX)
        End If
        AI.Left = (Rect4.Left * Screen.TwipsPerPixelY - AI.Width)
        Lk2 = Rect4.Top + Rect4.Left
        pd = 2
    End If
End If
End If
HwndW2 = Hwndw

Hwndw = FindWindow(vbNullString, "BBC 24 Hour Weather")
xxa = GetWindowState(Hwndw)
If Hwndw > 0 And xxa <> OXXA Then
'OldWinStat = xxa
'Lk1 = -1
Hwnd24 = GetWindowRect(Hwndw, Rect5)
If Rect5.Top > 0 Or Lk3 <> Rect5.Top + Rect5.Left Then
    AI.Top = (Rect5.Bottom * Screen.TwipsPerPixelX)
    Lk3 = Rect5.Top + Rect5.Left
    If HwndW2 = 0 Then
    AI.Left = (Rect5.Right * Screen.TwipsPerPixelX) - AI.Width
    
    End If
End If
pd2 = 1
End If

If OXXA <> xxa Then pd2 = 1: Lk4 = 0
OXXA = xxa
'If settid = 1 Then pd2 = 1: Lk4 = 0

If Lk1 + Lk2 + Lk3 <> Lk4 And pd2 = 1 Then
    If pd2 = 0 Then AI.Top = (420)
    If pd = 0 And Hwndw = 0 Then AI.Left = (Screen.Width - AI.Width)
    If xxa <> -1 And HwndW2 = 0 Then AI.Left = (Screen.Width - AI.Width)
    
    Lk4 = Lk1 + Lk2 + Lk3
    If xxa = 1 And pd2 = 1 And pd <> 2 Then
        If HwndW2 = 0 Then
        AI.Top = (480)
        Else
        AI.Top = (Rect4.Top * Screen.TwipsPerPixelX)
        
        End If
    End If
End If

If hwnd2 = 0 And hwnd > 0 And xxa <> -1 And HALO = 0 Then
        AI.Left = (Screen.Width) - AI.Width
        AI.Top = 480
HALO = 1
End If
If hwnd2 <> 0 And xxa = -1 And HALO = 1 Then HALO = 0
If settid = 1 And HwndW2 = 0 And HALO2 = 0 Then
    AI.Left = (Screen.Width) - AI.Width
    AI.Top = 480
    HALO2 = 1
End If
If HwndW2 > 0 And HALO2 = 1 Then HALO2 = 0

End Sub

Private Sub Timer3_Timer()

Call MouseDetectMove

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

Sub FileInUseDelay(Tx$)
        
    qq = Now + TimeSerial(0, 5, 0)
    Do
'        DoEvents
        ii = FileInUse(Tx$)
        If ii = True Then Sleep (50)
    Loop Until ii = False Or qq < Now
    
    If ii = True Then
        MsgBox "Trouble File Stuck In Use " + vbCrLf + Tx$ + vbCrLf + "Longer than 5 min"
        Stop
    End If
End Sub


Private Sub Timer11_Timer()
On Local Error GoTo Errend
frg = FreeFile
Open RamDrive + ":\temp\KeyHit-Tidal.txt" For Input As #frg
Line Input #frg, kb$
Close #frg



KeyHitts = Val(kb$)
If KeyHittsOld = 0 Then
    KeyHittsOld = KeyHitts - 1
    If KeyHittsOld < 0 Then KeyHittsOld = 0
End If
If KeyHittsOld > KeyHitts Then KeyHittsOld = 0

If KeyHittsOld = KeyHitts Then Exit Sub


Keyy = Keyy + (KeyHitts - KeyHittsOld)
'Daykeyy = Daykeyy + (KeyHitts - KeyHittsOld)

Label23 = Str(Keyy)

KeyHittsOld = KeyHitts


'Kbbdf = True
'KBPress2 = vcode
'LastPressTrig = True
'LastPressTrigTimer.Enabled = True
'Call CID_Run_Me.Lastpress

Errend:
Close #frg
End Sub




Property Let WindowVisible(hwnd As Long, New_Value As Boolean)
    If hwnd = 0 Then Exit Property
     
    Call ShowWindow(hwnd, IIf(New_Value, SW_SHOW, SW_HIDE))

End Property

Property Get WindowVisible(hwnd As Long) As Boolean

    If hwnd = 0 Then Exit Property
    WindowVisible = (IsWindowVisible(hwnd) > 0)

End Property


