VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mp3Rec"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameNew 
      Caption         =   "Start New File Every"
      Height          =   1695
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
      Begin VB.ComboBox cmbMinutes 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   240
         Width           =   660
      End
      Begin VB.CheckBox CheckPaus 
         Caption         =   "Pause If Level Below"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox tbxDB 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Text            =   "tbxDB"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox tbxSec 
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Text            =   "tbxSec"
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblMinutes 
         AutoSize        =   -1  'True
         Caption         =   "Minutes"
         Height          =   195
         Left            =   840
         TabIndex        =   35
         Top             =   300
         Width           =   555
      End
      Begin VB.Label lbldBFor 
         AutoSize        =   -1  'True
         Caption         =   "dB For More Than"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   34
         Top             =   1005
         Width           =   1290
      End
      Begin VB.Label lbldBFor 
         AutoSize        =   -1  'True
         Caption         =   "Seconds"
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   33
         Top             =   1365
         Width           =   630
      End
   End
   Begin VB.Frame FrameStart 
      Caption         =   "Start Writing At (Time)"
      Height          =   1695
      Left            =   120
      TabIndex        =   23
      Top             =   1320
      Width           =   2055
      Begin VB.TextBox tbxHours 
         Height          =   285
         Left            =   720
         TabIndex        =   24
         Text            =   "tbxHours"
         Top             =   1200
         Width           =   615
      End
      Begin VB.OptionButton optDur 
         Caption         =   "For                  Hours"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   1230
         Width           =   1815
      End
      Begin VB.OptionButton optDur 
         Caption         =   "Forever"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox tbxTimeStart 
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Text            =   "tbxTimeStart"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lbKeep 
         AutoSize        =   -1  'True
         Caption         =   "Keep Writing To Disk"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1515
      End
   End
   Begin VB.Frame FrameFull 
      Height          =   1575
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   4455
      Begin VB.TextBox tbxFull 
         Height          =   285
         Left            =   1080
         TabIndex        =   20
         Text            =   "tbxFull"
         Top             =   165
         Width           =   615
      End
      Begin VB.OptionButton optFull 
         Caption         =   "Delete Oldest Log Folder"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   525
         Width           =   2175
      End
      Begin VB.OptionButton optFull 
         Caption         =   "Stop Writing"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   765
         Width           =   1215
      End
      Begin VB.TextBox tbxOld 
         Height          =   285
         Left            =   2640
         TabIndex        =   17
         Text            =   "tbxOld"
         Top             =   1155
         Width           =   615
      End
      Begin VB.CheckBox CheckOld 
         Caption         =   "Delete Log Folders Older Than                   Days"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   3975
      End
      Begin VB.Label lblFull 
         AutoSize        =   -1  'True
         Caption         =   "If Less Than                 MB On Disk"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   210
         Width           =   2505
      End
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Stop Write"
      Height          =   195
      Index           =   3
      Left            =   780
      TabIndex        =   15
      Top             =   840
      Width           =   1065
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Start Write"
      Height          =   195
      Index           =   2
      Left            =   780
      TabIndex        =   14
      Top             =   600
      Width           =   1065
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Auto"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   735
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Off"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   735
   End
   Begin VB.Frame FrameBitRate 
      Caption         =   "Bitrate"
      Height          =   735
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   1815
      Begin VB.ComboBox cmbBitRate 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   270
         Width           =   660
      End
      Begin VB.OptionButton optStereo 
         Caption         =   "Mono"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   7
         Top             =   150
         Width           =   855
      End
      Begin VB.OptionButton optStereo 
         Caption         =   "Stereo"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   6
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4080
      Top             =   0
   End
   Begin VB.Frame frameVu 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      Begin VB.Label lblVu 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4695
      End
   End
   Begin VB.Label lblWrite 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Write On"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1860
      TabIndex        =   11
      Top             =   975
      Width           =   870
   End
   Begin VB.Label lblAuto 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Auto Off"
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   1860
      TabIndex        =   10
      Top             =   735
      Width           =   870
   End
   Begin VB.Label lblCapture 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Capture On"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1860
      TabIndex        =   9
      Top             =   495
      Width           =   870
   End
   Begin VB.Label lbldB 
      AutoSize        =   -1  'True
      Caption         =   "-60            -50            -40            -30            -20            -10            0"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   300
      Width           =   4680
   End
   Begin VB.Label lblVu 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---- Thu 11-Mar-2010 11:41:46a ----  ----
'Dancer!
'Http://www.rarewares.org/dancer/dancer.php?f=285

'---- Thu 11-Mar-2010 11:41:55a ----  ----
'Mp3Rec
'Http://www.heimskringla.com/software/mp3rec/index.html

'---- Thu 11-Mar-2010 11:41:59a ----  ----
'Heimskringla Software
'Http://www.heimskringla.com/software/index.html



Option Explicit
Option Compare Text

Dim MeSw As Single
Dim SampRate As Integer
Dim PausNow As Boolean
Dim AudioLimit As Double
Dim TimeLimit As Double
Dim TimeCheck As Date
'

Private Function FindOldestFolder(InFolder As String) As String
  
  Dim RetStr As String
  Dim Older As String
  
  If Right$(InFolder, 1) <> "\" Then InFolder = InFolder & "\"
  
  RetStr = Dir$(InFolder & "*.*", vbDirectory)
  Older = sE
  
  Do While RetStr <> sE
    If RetStr <> "." And RetStr <> ".." Then
      If (GetAttr(InFolder & RetStr) And vbDirectory) = vbDirectory Then
        'It's a folder
        If Older = sE Then Older = RetStr
        If Older > RetStr Then Older = RetStr
        ' Compare names, not age
      End If
    End If
    
    RetStr = Dir$
  Loop
  
  FindOldestFolder = Older

End Function

Private Sub SetVar()

  If Capturing Then
    lblCapture.ForeColor = vbRed
    lblCapture.Caption = "Capture On"
    cmbBitRate.Enabled = False
    optStereo(1).Enabled = False
    optStereo(2).Enabled = False
  Else
    lblCapture.ForeColor = vbGreen
    lblCapture.Caption = "Capture Off"
    cmbBitRate.Enabled = True
    optStereo(1).Enabled = True
    optStereo(2).Enabled = True
    frameVu.Width = 0
  End If
  
  If Auto Then
    lblAuto.ForeColor = vbRed
    lblAuto.Caption = "Auto On"
  Else
    lblAuto.ForeColor = vbGreen
    lblAuto.Caption = "Auto Off"
  End If
  
  If Writing Then
    lblWrite.ForeColor = vbRed
    lblWrite.Caption = "Write On"
  Else
    lblWrite.ForeColor = vbGreen
    lblWrite.Caption = "Write Off"
  End If
  
End Sub

Public Sub Vu()
  
  Static i As Long
  Static x As Long
  Static VuNow As Long
  Static VuMax As Long
  Static Norm As Double
  Static dBNorm As Double
  
  For x = 1 To UBound(sBuff)
    'DoEvents
    VuNow = Abs(sBuff(x))
    If VuMax < VuNow Then VuMax = VuNow
  Next x
  
  i = i + 1
  If i = 10 Then i = 0
  
  If i = 0 Then
    Norm = VuMax / 32768#
    dBNorm = (20 * Log(Norm) / Log(10#)) / 60 + 1
    If dBNorm < 0 Then dBNorm = 0
    
    frameVu.Width = dBNorm * MeSw
    lblVu(1).BackColor = RGB(512 * dBNorm, 512 * (1 - dBNorm), 0)
    VuMax = 0
  
    If Norm > AudioLimit Then
      PausNow = False
      TimeCheck = Now
    Else
      If (Now - TimeCheck > TimeLimit) And PauseOnSilence Then
        PausNow = True
      End If
    End If
    
  End If
  
End Sub

Private Sub CheckOld_Click()
  KillOld = (CheckOld.Value = vbChecked)
End Sub

Private Sub CheckPaus_Click()
  PauseOnSilence = (CheckPaus.Value = vbChecked)
  PausNow = False
  SetVar
End Sub

Private Sub cmbBitRate_Click()
  BitRate = CInt(cmbBitRate.Text)
  
  If Chan = 1 Then
    'Mono
    Select Case BitRate
    Case 8, 16
      SampRate = 1
    Case 24, 32, 144
      SampRate = 2
    Case Else
      SampRate = 4
    End Select
    
  Else
    'Stereo
    Select Case BitRate
    Case 8 To 32
      SampRate = 1
    Case 40 To 64, 144
      SampRate = 2
    Case Else
      SampRate = 4
    End Select
    
  End If
  
End Sub

Private Sub cmbMinutes_Click()
  
  MinuteChange = CLng(cmbMinutes.Text)
  
End Sub

Private Sub Form_Load()
  Dim i As Long
  
'
'For MPEG1 (sampling frequencies of 32, 44.1 and 48 kHz)
'        32,40,48,56,64,80,96,112,128,    160,192,224,256,320
'For MPEG2 (sampling frequencies of 16, 22.05 and 24 kHz)
'8,16,24,32,40,48,56,64,80,96,112,128,144,160
  cmbBitRate.AddItem "8"
  cmbBitRate.AddItem "16"
  cmbBitRate.AddItem "24"
  cmbBitRate.AddItem "32"
  cmbBitRate.AddItem "40"
  cmbBitRate.AddItem "48"
  cmbBitRate.AddItem "56"
  cmbBitRate.AddItem "64"
  cmbBitRate.AddItem "80"
  cmbBitRate.AddItem "96"
  cmbBitRate.AddItem "112"
  cmbBitRate.AddItem "128"
  cmbBitRate.AddItem "144"
  cmbBitRate.AddItem "160"
  cmbBitRate.AddItem "192"
  cmbBitRate.AddItem "224"
  cmbBitRate.AddItem "256"
  cmbBitRate.AddItem "320"
  
  For i = 0 To cmbBitRate.ListCount - 1
    If CStr(BitRate) = cmbBitRate.List(i) Then
      cmbBitRate.Text = cmbBitRate.List(i)
    End If
  Next
  If Not IsNumeric(cmbBitRate.Text) Then
    cmbBitRate.Text = cmbBitRate.List(0)
  End If
  
  cmbMinutes.AddItem "1"
  cmbMinutes.AddItem "5"
  cmbMinutes.AddItem "10"
  cmbMinutes.AddItem "15"
  cmbMinutes.AddItem "20"
  cmbMinutes.AddItem "30"
  cmbMinutes.AddItem "60"
  cmbMinutes.AddItem "120"
  cmbMinutes.AddItem "180"
  cmbMinutes.AddItem "240"
  
  For i = 0 To cmbMinutes.ListCount - 1
    If CStr(MinuteChange) = cmbMinutes.List(i) Then
      cmbMinutes.Text = cmbMinutes.List(i)
    End If
  Next
  If Not IsNumeric(cmbMinutes.Text) Then
    cmbMinutes.Text = cmbMinutes.List(0)
  End If
  cmbMinutes.Text = cmbMinutes.List(0)
  
  tbxHours_LostFocus
  tbxTimeStart_LostFocus
  tbxSec_LostFocus
  tbxDB_LostFocus
  tbxFull_LostFocus
  tbxOld_LostFocus
  
  optStereo(Chan).Value = True
  
  If DurTot = FOREVER Then
    optDur(1).Value = True
    tbxHours.Text = "2"
  Else
    optDur(0).Value = True
  End If
  
  If KillIfFull Then
    optFull(0).Value = True
  Else
    optFull(1).Value = True
  End If
  
  If Auto Then
    optMode(1).Value = True
  Else
    optMode(3).Value = True
  End If
  
  If PauseOnSilence Then
    CheckPaus.Value = vbChecked
  Else
    CheckPaus.Value = vbUnchecked
  End If
  
  If KillOld Then
    CheckOld.Value = vbChecked
  Else
    CheckOld.Value = vbUnchecked
  End If
CheckOld.Value = vbUnchecked
  
  MeSw = Me.ScaleWidth
  
  lblVu(0).Move 0, 0, MeSw, 300
  frameVu.Move 0, 0, 0, 300
  lblVu(1).Move 0, 0, MeSw, 300
  
  SetVar
  
optFull(1) = vbChecked
CheckPaus = vbUnchecked
'optDur(0) = vbChecked
End Sub

Private Sub Form_Unload(Cancel As Integer)
  StopCapture
  Set frmMain = Nothing
End Sub

Private Sub optDur_Click(Index As Integer)
  
  tbxHours.Enabled = optDur(0).Value
  
  If optDur(1).Value Then
    DurTot = FOREVER
  Else
    tbxHours_LostFocus
  End If
  
End Sub

Private Sub optFull_Click(Index As Integer)
  KillIfFull = (Index = 0)
End Sub

Private Sub optMode_Click(Index As Integer)

  Select Case Index
  Case 0 ' Off
    Auto = False
    StopWrite
    StopCapture
    
  Case 1 ' Auto
    Auto = True
    If Not StartCapture(Chan, SampRate) Then StopCapture
    
  Case 2 ' Start Write
    Auto = False
    If Not StartCapture(Chan, SampRate) Then StopCapture
    If Not StartWrite Then StopWrite
    
  Case 3 ' Stop  Write
    Auto = False
    If Not StartCapture(Chan, SampRate) Then StopCapture
    StopWrite
    
  End Select
  
  PausNow = False
  TimeCheck = Now
  SetVar
  
End Sub

Private Sub optStereo_Click(Index As Integer)
  Chan = Index
  cmbBitRate_Click
End Sub

Private Sub tbxDB_LostFocus()
  
  If IsNumeric(tbxDB.Text) Then
    dB = -Abs(CDbl(tbxDB.Text))
  End If
  
  tbxDB.Text = CStr(dB)
  AudioLimit = 10 ^ (dB / 20)

End Sub

Private Sub tbxFull_LostFocus()
  
  If IsNumeric(tbxFull.Text) Then
    DiskMin = Abs(CDbl(tbxFull.Text))
  End If
  
  tbxFull.Text = CStr(DiskMin)

End Sub

Private Sub tbxHours_LostFocus()
  
  If IsNumeric(tbxHours.Text) Then
    DurTot = Abs(CDbl(tbxHours.Text))
  End If
  
  tbxHours.Text = CStr(DurTot)
  
End Sub

Private Sub tbxOld_LostFocus()

  If IsNumeric(tbxOld.Text) Then
    Old = Abs(CDbl(tbxOld.Text))
  End If
  
  tbxOld.Text = CStr(Old)
  
End Sub

Private Sub tbxSec_LostFocus()
  
  If IsNumeric(tbxSec.Text) Then
    SecSilence = Abs(CDbl(tbxSec.Text))
  End If
  
  tbxSec.Text = CStr(SecSilence)
  TimeLimit = SecSilence / 86400
  
End Sub

Private Sub tbxTimeStart_LostFocus()

  If IsDate(tbxTimeStart.Text) Then
    TimeStart = CDate(tbxTimeStart.Text)
  End If
  
  tbxTimeStart.Text = Format$(TimeStart, "ddddd ttttt")

End Sub

Private Sub Timer1_Timer()
  
  Static JustStarted As Boolean
  Static OldMin As Integer
  Static NewMin As Integer
  Static Oldest As String
  Static OldDate As Date
  
  If Not Capturing Then Exit Sub
  
  If ABufferIsFull Then ReadBuffer
  
  ' Delete oldest folders or stop
  NewMin = Minute(Time)
  If OldMin <> NewMin Then
    If NewMin Mod 5 = 2 Then
      If ExistFile(LogFolder) Then
        Oldest = FindOldestFolder(LogFolder)
        
        If ExistFile(LogFolder & Oldest) Then
          
          If IsDate(Oldest) Then
            OldDate = CDate(Oldest)
          Else
            OldDate = CDate("1111-11-11")
          End If
          
          Oldest = LogFolder & Oldest
          
          If OldDate < Now - 1 Then ' Can't be too new
            If OldDate < Now - Old Then
              ' If older than limit then Delete Oldest
              If CheckOld.Value = vbChecked Then
                DelTree Oldest
              End If
            ElseIf FreeSpaceInfo < DiskMin Then
              If optFull(0).Value Then  ' If Stop then
                optMode(3).Value = True ' Stop
              Else                      ' If delete then
                'Delete Oldest
                DelTree Oldest
              End If
            End If
          End If
          
        End If
        
      End If
    End If
  End If
  OldMin = NewMin
  
  If Not Auto Then Exit Sub
  
  ' Write if within time limits ---------------
  If Now < TimeStart Or Now > TimeStart + DurTot / 24 Then
    If Writing Then StopWrite: SetVar
    Exit Sub
  End If
  
  ' Pause if silence --------------------------
  If Not PausNow And Not Writing Then StartWrite: SetVar
  If PausNow And Writing Then StopWrite: SetVar
  
  ' Start a new file every x minutes ----------
  If Not PausNow And Writing Then
    If (Int(Time * 1440) Mod MinuteChange) = 0 Then
      If Not JustStarted Then
        Call StopWrite: SetVar
        If Not Writing Then StartWrite: SetVar 'Start again
        JustStarted = Writing
      End If
    Else
      JustStarted = False
    End If
  End If
  
End Sub
