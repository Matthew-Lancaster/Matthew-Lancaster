VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VU LOGGER"
   ClientHeight    =   5985
   ClientLeft      =   150
   ClientTop       =   525
   ClientWidth     =   12735
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   12735
   Begin VB.Timer Timer4 
      Interval        =   500
      Left            =   4740
      Top             =   4185
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   390
      Left            =   -15
      TabIndex        =   45
      Top             =   2190
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4050
      Top             =   2955
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   3420
      Top             =   2970
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   2850
      Top             =   2955
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5070
      TabIndex        =   44
      Top             =   1590
      Width           =   5070
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   6285
      TabIndex        =   43
      Top             =   630
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   -30
      TabIndex        =   36
      Top             =   2565
      Width           =   12645
   End
   Begin VB.Frame FrameNew 
      Caption         =   "Start New File Every"
      Height          =   1695
      Left            =   3165
      TabIndex        =   3
      Top             =   5595
      Visible         =   0   'False
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
      Left            =   1005
      TabIndex        =   23
      Top             =   5595
      Visible         =   0   'False
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
      Left            =   1005
      TabIndex        =   16
      Top             =   7275
      Visible         =   0   'False
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
      Left            =   1665
      TabIndex        =   15
      Top             =   5115
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Start Write"
      Height          =   195
      Index           =   2
      Left            =   1665
      TabIndex        =   14
      Top             =   4875
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Auto"
      Height          =   195
      Index           =   1
      Left            =   1005
      TabIndex        =   13
      Top             =   5115
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Off"
      Height          =   195
      Index           =   0
      Left            =   1005
      TabIndex        =   12
      Top             =   4875
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame FrameBitRate 
      Caption         =   "Bitrate"
      Height          =   735
      Left            =   3645
      TabIndex        =   5
      Top             =   4755
      Visible         =   0   'False
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
         Top             =   -15
         Width           =   4695
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5070
      TabIndex        =   42
      Top             =   1065
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5085
      TabIndex        =   41
      Top             =   570
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1650
      TabIndex        =   40
      Top             =   585
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1650
      TabIndex        =   39
      Top             =   1590
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1650
      TabIndex        =   38
      Top             =   1095
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   15
      TabIndex        =   37
      Top             =   570
      Width           =   450
   End
   Begin VB.Label lblWrite 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Write On"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2745
      TabIndex        =   11
      Top             =   5250
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblAuto 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Auto Off"
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   2745
      TabIndex        =   10
      Top             =   5010
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblCapture 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Capture On"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2745
      TabIndex        =   9
      Top             =   4770
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lbldB 
      AutoSize        =   -1  'True
      Caption         =   "-60            -50            -40            -30            -20            -10            0"
      Height          =   195
      Left            =   30
      TabIndex        =   4
      Top             =   285
      Width           =   4695
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
   Begin VB.Menu MNU_NORM 
      Caption         =   "NORM LOGG"
      Begin VB.Menu MNU_OPEN_LOGG 
         Caption         =   "OPEN  LOGG"
      End
      Begin VB.Menu MNU_LOGG_FOLDER 
         Caption         =   "LOGG FOLDER"
      End
   End
   Begin VB.Menu MNU_PEAK 
      Caption         =   "PEAK LOGG"
      Begin VB.Menu MNU_PEAK_LOGG 
         Caption         =   "PEAK LOGG"
      End
      Begin VB.Menu MNU_PEAK_LOGG_FOLDER 
         Caption         =   "PEEK LOGG FOLDER"
      End
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB ME"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIND     SET_PEAK_HERE

Dim i, XX2, TT, FF

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wmessage As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Const MONITOR_ON = -1&
Private Const MONITOR_OFF = 2&
Private Const SC_MONITORPOWER = &HF170&
Private Const WM_SYSCOMMAND = &H112
Private Const ipc_isplaying As Long = 104
Private Const IPC_GETPLAYLISTFILE  As Long = 211
Private Const IPC_GETOUTPUTTIME  As Long = 105
Private Const IPC_GETINFO As Long = 126
Private Const WINAMP_BUTTON2 = 40045
Private Const WINAMP_BUTTON3 = 40046
Private Const WM_CLOSE = &H10
Private Const WM_USER = &H400
Private Const WM_COMMAND = &H111
Private Const GW_HWNDNEXT = 2

Dim Fr, VUM, VUMT, OMIN, TM, OTM, APP_PATH2, APP_PATH3
Dim OVUM
Dim OTM2, OMIN2
Dim VTX, VTY, DECP, TXV, VUMX
Dim RPEAK1, RPEAK2, RPEAK3, RPEAK4, RPEAK5
Dim sBuffPeak(102), xPointBuff
Dim sBuffPeakHOUR()
Dim XXT, Arrow
Dim LOGGSTRING, OLOGGSTRING
Dim TTMOUSEMOVE
Dim F, RD1, RD2
Dim GHO, GHO2, GH, LOGGS
Dim WINAMP2, MSGRESULT1, OMSGRESULT1


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
'Option Compare Text

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
  Static X As Long
  Static VuNow As Long
  Static VuMax As Long
  Static Norm As Double
  Static dBNorm As Double

  
  On Error Resume Next
  For X = 1 To UBound(sBuff)
    VuNow = Abs(sBuff(X))
    If VuMax < VuNow Then VuMax = VuNow
  Next X
  On Error GoTo 0
  
'  i = 1 / 0
  
  If Me.WindowState = 0 Then
    RPEAK4 = 1000
    For X = 0 To UBound(sBuffPeak)
        If RPEAK4 > sBuffPeak(X) Then
        If sBuffPeak(X) > 0 Then RPEAK4 = sBuffPeak(X)
        End If
    Next X
    End If
    
    For X = List2.ListCount - 1 To 1 Step -1
        XXT = DateValue(List2.List(X)) + TimeValue(List2.List(X))
        If XXT < Now - TimeSerial(0, 15, 0) Then
            List2.RemoveItem (X)
            Label6.BackColor = &H80FF80
        End If
    Next X
  
    RPEAK5 = List2.ListCount
    
    '## 100 BUFFER AVERAGE NX
  
    Norm = VuMax / 32768#
    On Error Resume Next
    dBNorm = (20 * Log(Norm) / Log(10)) / 60 + 1
    On Error GoTo 0
    If dBNorm < 0 Then dBNorm = 0
    
    'frameVu.Width = lblVu(1).Width / Screen.TwipsPerPixelX 'dBNorm * MeSw
    frameVu.Width = dBNorm * MeSw
    lblVu(1).BackColor = RGB(512 * dBNorm, 512 * (1 - dBNorm), 0)
    VuMax = 0
    
    'WE HAVE TO REVERSE THIS COZ THE MATHS
    'IS STILL CORRECT AS SHOWS DB SCALE
    
    VUM = 60 - (frameVu.Width / lblVu(0).Width) * 60
    VUMX = Round(VUM, 3)
    Label1 = VUMX
    
    '10 SEC
    If VUM < RPEAK1 Then RPEAK1 = VUMX: Label2 = Format(RPEAK1, "0.000") + " -- 10 Sec"
    '1 MIN
    If VUM < RPEAK2 Then RPEAK2 = VUMX: Label3 = Format(RPEAK2, "0.000") + " -- 30 Sec"
    '2 SEC
    If VUM < RPEAK3 Then RPEAK3 = VUMX: Label4 = Format(RPEAK3, "0.000") + " -- 0.5 Sec"
    
    Label5 = Format(RPEAK4, "0.000") + " --" + Str(UBound(sBuffPeak)) + " Buffer Peak"
    Label6 = Trim(Str(RPEAK5)) + " --" + " THUD's <" + Trim(Str(Arrow)) + "db - in Last 15 Mins"
    
    sBuffPeak(xPointBuff) = VUM
    xPointBuff = xPointBuff + 1
    If xPointBuff > UBound(sBuffPeak) Then xPointBuff = 0
    
    
    'PEAKT1 = 40 'dupe use of this later
    'Arrow = PEAKT1
    PEAKT2 = (60 - PEAKT1) - 1
    '>HIGHER THAN = YES THIS RESULT IS LOUDER THAN LAST
    If OVUM > RPEAK3 And RPEAK3 <= PEAKT1 Then
        'Reset
        If RPEAK3 < Arrow Then
            List2.AddItem Format(Now, "DD-mm-yyyy hh:mm:ss")
        End If
    
    End If
    
    
    WINAMP2 = FindWindow("Winamp v1.x", vbNullString)
    If WINAMP2 > 0 Then
        MSGRESULT1 = SendMessage(WINAMP2, WM_USER, 0, ByVal ipc_isplaying)
        If OMSGRESULT1 <> MSGRESULT1 Then
            If Me.WindowState = 1 Then Me.WindowState = 0
            If Me.WindowState = 0 Then Me.WindowState = 1
        End If

        
        OMSGRESULT1 = MSGRESULT1
        If MSGRESULT1 = 1 Then
            OVUM = RPEAK3
            Exit Sub
        End If
    End If
    
    
    
    'VOICE LOGG GRAB ALL IN SAME THIS BIT
    RD1 = 0
    If FS.FILEEXISTS("D:\TEMP\VOICE SPEAK LOGG.TXT") = True Then
    Set F = FS.getfile("D:\TEMP\VOICE SPEAK LOGG.TXT")
    RD1 = F.datelastmodified

err.Clear
ResumeHere1:
        If err.Number > 0 Then Sleep 50
        On Error GoTo ResumeHere1
        Call FileInUseDelay("D:\TEMP\VOICE SPEAK LOGG.TXT", False)
        Fr = FreeFile
        Open "D:\TEMP\VOICE SPEAK LOGG.TXT" For Input Lock Write As #Fr
        'FUTURE OF ResumeHere1

        On Error GoTo 0

        LOGGSTRING = Input(LOF(Fr), Fr)
'        Debug.Print LOGGSTRING
        LOGGSTRING = Replace(LOGGSTRING, " - ", " --- ", , 1)
'        Debug.Print LOGGSTRING
        Close #Fr
        'LOGGSTRING = "1#123456789" + vbCrLf
        'LOGGSTRING = LOGGSTRING + "2#123456789" + vbCrLf
        'LOGGSTRING = LOGGSTRING + "3#123456789" + vbCrLf
        'LOGGSTRING = LOGGSTRING + "4#"
        
        GHO = 0
        GHO2 = 1
        Do
            GH = InStr(GHO + 1, LOGGSTRING, vbCrLf)
            If GH = 0 Then
                If Len(LOGGSTRING) > GHO + 1 Then
                    GH = Len(LOGGSTRING) + 1
                End If
            End If
            If GH = 0 Then Exit Do
            If GHO > 0 Then GHO = GHO - 1
            LOGGS = Mid(LOGGSTRING, GHO2, (GH - GHO2))
            'Debug.Print "--" + LOGGS + "--"
            List1.AddItem LOGGS
            If TTMOUSEMOVE < Now Then
                List1.ListIndex = List1.ListCount - 1
            End If
            GHO = GH
            GHO2 = GH + 2
        Loop Until GH = 0
        If Dir("D:\TEMP\VOICE SPEAK LOGG.TXT") <> "" Then
            Kill "D:\TEMP\VOICE SPEAK LOGG.TXT"
        End If
    End If
    
    
    'END VOICE LOGG START VU M LOGG
    
    'Reset
    Fr = FreeFile
    TM = APP_PATH2 + "\VU LOGG " + Format(Now, "YYYY-MM-DD") + ".txt"
    
    
    If FS.FILEEXISTS(TM) = False Then OTM = ""
err.Clear
ResumeHere2:
    If err.Number > 0 Then Sleep 50
    On Error GoTo ResumeHere2
    On Error Resume Next
    Call FileInUseDelay(TM, True)
    Fr = FreeFile
    Open TM For Append Lock Write As #Fr
    'FUTURE OF ResumeHere2
    
    'On Error GoTo 0
    If OTM <> TM Then
        OTM = TM
        Print #Fr, "# VU METER LOGGER FOR DAY STARTS"
        Print #Fr, "# --"
        Print #Fr, "# " + Format(Now, "DDD DD MMM YYYY HH:MM:SS")
        Print #Fr, "# --"
        Print #Fr, "# Logging is -60 to 0 db"
        Print #Fr, "# --"
    End If
    
    If OMIN <> Minute(Now) Then
        OMIN = Minute(Now)
        Print #Fr, "# TIME - " + Format(Now, "HH:MM")
    End If
    
    DECP = 2
    VUMT = "-" + Format(VUM, "00." + String(DECP, "0"))
    'Debug.Print VUMT
    
    Print #Fr, VUMT
    Close Fr
    
    
    DECP = 2
    'VUM = Round(VUM, DECP)
    'PEAKT1 = 39
    Arrow = PEAKT1
    PEAKT2 = (60 - PEAKT1) - 1
    '>HIGHER THAN = YES THIS RESULT IS LOUDER THAN LAST
    If OVUM > RPEAK3 And RPEAK3 <= PEAKT1 Then
    'Reset
'    If RPEAK3 < Arrow Then
'        List2.AddItem Format(Now, "DD-mm-yyyy hh:mm:ss")
'    End If
    
    Fr = FreeFile
    TM = APP_PATH3 + "\VU LOGG " + Format(Now, "YYYY-MM-DD") + ".txt"
    If FS.FILEEXISTS(TM) = False Then OTM2 = ""
    
    
'--- LOGGING PEAK
    
err.Clear
ResumeHere3:
    If err.Number > 0 Then Sleep 50
    On Error GoTo ResumeHere3
    On Error Resume Next

    Call FileInUseDelay(TM, True)
    Fr = FreeFile
    Open TM For Append Lock Write As #Fr
    'FUTURE OF ResumeHere3

    'On Error GoTo 0
    If OTM2 <> TM Then
        OTM2 = TM
        Print #Fr, "# --"
        Print #Fr, "# VU METER LOGGER FOR DAY STARTS ## LOGGER MODE = PEAK MODE"
        Print #Fr, "# " + Format(Now, "DDD DD MMM YYYY HH:MM:SS")
        Print #Fr, "# --"
    End If
    
    VTX = Int(((60 - RPEAK3) - PEAKT2))
    VUMT = "-" + Format(RPEAK3, "00." + String(DECP, "0"))
    VTY = String(VTX, "-")
    'Debug.Print VUM
    'Debug.Print VUMT + " - " + Format(Now, "HH:MM:SS") + " # " + VTY
    
    
    
    

    TT = DateDiff("s", XX2, Now)
    FF = Format(Int(Now) + TimeSerial(0, 0, TT), "HH:MM:SS") + "''s Gap "
    FF = Replace(FF, "00:", "")
    XX2 = Now
    
    TXV = VUMT + " - " + Format(Now, "DD MMM - HH:MM:SS") + " # -- " + FF + " # " + VTY
    Print #Fr, TXV
    Close Fr
    
    List1.AddItem TXV
    If TTMOUSEMOVE < Now Then
        List1.ListIndex = List1.ListCount - 1
    End If
    
    Do While List1.ListCount > 5000
        List1.RemoveItem (0)
    Loop
    
    
    End If
    
    OVUM = RPEAK3
    
err.Clear
ResumeHere4:
    If err.Number > 0 Then Sleep 100
    On Error GoTo ResumeHere4
    If FS.FILEEXISTS("D:\TEMP\VOICE SPEAK LOGG.TXT") Then
        Set F = FS.getfile("D:\TEMP\VOICE SPEAK LOGG.TXT")
        RD2 = F.datelastmodified
        If RD2 = RD1 Or RD1 = 0 Then
            Kill "D:\TEMP\VOICE SPEAK LOGG.TXT"
        End If
    End If
    'FUTURE OF ResumeHere4
    On Error GoTo 0
  
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
    
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    
    Call Record.Main
    XX2 = Now
    
    
    'SET_PEAK_HERE
    PEAKT1 = 39
    Arrow = PEAKT1 'usual same as PEAKT1
    RPEAK1 = 1000
    RPEAK2 = 1000
    RPEAK3 = 1000
    Set FS = CreateObject("Scripting.FileSystemObject")
    
    App.Title = App.EXEName
    frmMain.Caption = App.EXEName

    
    
    
    
    If Dir("M:\TEMP", vbDirectory) <> "" Then
        APP_PATH2 = "M:\00 Loggers Text\VU METER LOGGER"
        APP_PATH3 = "M:\00 Loggers Text\VU METER LOGGER PEAK"
        Else
        APP_PATH2 = App.Path + "\VU METER LOGGER"
        APP_PATH3 = App.Path + "\VU METER LOGGER PEAK"
    End If

    TM = APP_PATH2 + "\VU LOGG " + Format(Now, "YYYY-MM-DD") + ".txt"
    OTM = TM
    TM = APP_PATH3 + "\VU LOGG " + Format(Now, "YYYY-MM-DD") + ".txt"
    OTM2 = TM
    Call FileInUseDelay("D:\temp\VU LOGGER PATHS.txt", True)
    Fr = FreeFile
    Open "D:\temp\VU LOGGER PATHS.txt" For Output As #Fr
        Print #Fr, OTM
        Print #Fr, OTM2
    Close #Fr


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
    
    cmbBitRate.Text = 128
  
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
  
  'Me.Width = lblVu(1).Width
  
  MeSw = 4695 'Me.ScaleWidth
  
  lblVu(0).Move 0, 0, MeSw, 300
  frameVu.Move 0, 0, 0, 300
  lblVu(1).Move 0, 0, MeSw, 300
  
  SetVar
  
optFull(1) = vbChecked
CheckPaus = vbUnchecked
'optDur(0) = vbChecked

On Error Resume Next
If IsIDE = False Then Me.WindowState = 1
'Me.WindowState = 1

Me.Width = List1.Left + List1.Width
Me.Height = List1.Top + List1.Height + 800
DoEvents
Call Form_Resize
On Error GoTo 0

Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True

Call Form_Resize
End Sub

Private Sub Form_Resize()
On Error Resume Next

Me.Width = List1.Left + List1.Width
Me.Height = List1.Top + List1.Height + 800
Me.Top = Screen.Height / 2 - Me.Height / 2
'Me.Left = 5000 + (Screen.Width / 2 - Me.Width / 2)
DoEvents
Me.Left = Screen.Width - Me.Width

On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
  StopCapture
  Set frmMain = Nothing
End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label5_Click()
'Label5
End Sub

Private Sub Label6_Click()
'LABEL6
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TTMOUSEMOVE = Now + TimeSerial(0, 0, 5)
End Sub

Private Sub MNU_LOGG_FOLDER_Click()
TM = APP_PATH2 + "\VU LOGG " + Format(Now, "YYYY-MM-DD") + ".txt"
Shell "Explorer.exe /select, " + TM, vbNormalFocus
'Shell "Explorer.exe /select, " + APP_PATH2 + "\", vbNormalFocus

End Sub

Private Sub MNU_OPEN_LOGG_Click()
TM = APP_PATH2 + "\VU LOGG " + Format(Now, "YYYY-MM-DD") + ".txt"
Call FileInUseDelay(TM, True)
Shell "NOTEPAD " + TM, vbMaximizedFocus

End Sub

Private Sub MNU_PEAK_LOGG_Click()
TM = APP_PATH3 + "\VU LOGG " + Format(Now, "YYYY-MM-DD") + ".txt"
Shell "NOTEPAD " + TM, vbMaximizedFocus

End Sub

Private Sub MNU_PEAK_LOGG_FOLDER_Click()
TM = APP_PATH3 + "\VU LOGG " + Format(Now, "YYYY-MM-DD") + ".txt"
Shell "Explorer.exe /select, " + TM, vbNormalFocus

End Sub

Private Sub MNU_VB_Click()
If IsIDE = True Then Unload Me: Exit Sub

Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
End
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

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)



If KeyCode = 13 Then
    List1.AddItem "# TEXT -- " + Text1
    If TTMOUSEMOVE < Now Then
        List1.ListIndex = List1.ListCount - 1
    End If
If Trim(Text1 = "") Then Exit Sub
err.Clear
ResumeHere5:
    If err.Number > 0 Then Sleep 50
    On Error GoTo ResumeHere5

    Call FileInUseDelay(TM, True)
    Fr = FreeFile
    Open TM For Append Lock Write As #Fr
    'FUTURE OF ResumeHere5
    On Error GoTo 0

    
    Print #Fr, "# TEXT -- " + Text1
    Close #Fr
    Text1 = ""
End If
KeyCode = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
  
  Static JustStarted As Boolean
  Static OldMin As Integer
  Static NewMin As Integer
  Static Oldest As String
  Static OldDate As Date

    
    
  If Not Capturing Then
    Exit Sub
  End If
  'If ABufferIsFull Then
    ReadBuffer
  'End If
    
    Exit Sub
  
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

Private Sub Timer2_Timer()
If Int(Timer) Mod 10 = 0 Then RPEAK1 = 1000
If Int(Timer) Mod 30 = 0 Then RPEAK2 = 1000
If Int(Timer * 2) Mod 1 = 0 Then RPEAK3 = 1000
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

Private Sub Timer3_Timer()
WINAMP2 = FindWindow("Winamp v1.x", vbNullString)
If WINAMP2 > 0 Then
    MSGRESULT1 = SendMessage(WINAMP2, WM_USER, 0, ByVal ipc_isplaying)
    If MSGRESULT1 = 1 Then
    Timer1.Enabled = False
    Else
    Timer1.Enabled = True
    
    End If
End If

End Sub


Public Sub ReadBuffer()

  Static LB As Long
  Dim k As Long
  
  k = LB
  For i = 1 To NUM_BUFFERS
    k = k + 1
    If k > NUM_BUFFERS Then k = 1
    
    If inHdr(k).dwFlags And WHDR_DONE Then
      LB = k
      rc = waveInAddBuffer(hWaveIn, inHdr(k), Len(inHdr(k)))
      If rc <> 0 Then
        MsgBox "waveInAddBuffer failed"
        Exit Sub
      End If

      CopyMemory sBuff(1), ByVal inHdr(k).lpData, BufferSize
      
  
  End If
Next
  
'THIS LOOP RUNS ONCE
  

'TRICKY
'IF WE CALL LIKE THIS ON ERROR WILL COME BACK TO THIS CALL
'Call frmMain.Vu

'IF WE CALL LIKE THIS ON ERROR WILL STOP IN THE SUB
Call Vu
'THIS MEANS WILL HAVE THIS SUB IN THE FORM WHERE CALL GOES
    
    
    
ABufferIsFull = False

End Sub

Private Sub Timer4_Timer()
    On Error Resume Next
    TT = DateDiff("s", XX2, Now)
    ProgressBar1.Max = 60
    ProgressBar1.Min = 0
    ProgressBar1.Value = TT
    ProgressBar1.Left = 0
    ProgressBar1.Width = Me.Width
End Sub
