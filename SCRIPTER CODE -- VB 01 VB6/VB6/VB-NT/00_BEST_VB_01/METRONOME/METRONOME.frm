VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "METRONOME"
   ClientHeight    =   2955
   ClientLeft      =   165
   ClientTop       =   885
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   4875
      Top             =   2025
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4875
      Top             =   1500
   End
   Begin MCI.MMControl MMControl1 
      Height          =   465
      Left            =   2550
      TabIndex        =   2
      Top             =   675
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   820
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   150
      TabIndex        =   1
      Top             =   600
      Width           =   2265
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   390
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   688
      _Version        =   393216
      BorderStyle     =   1
      LargeChange     =   1
      Max             =   4000
      SelStart        =   888
      Value           =   888
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
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
      Left            =   6300
      TabIndex        =   5
      Top             =   675
      Width           =   2235
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Random"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   2625
      TabIndex        =   4
      Top             =   1950
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "OFF ON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   2625
      TabIndex        =   3
      Top             =   1275
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB ME"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim H1, H2, RANDOMH, SLIDERHH, MM

Private Sub File1_Click()
Shell "explorer /select, " + App.Path, vbNormalFocus
End Sub

Private Sub Form_Load()
End
'Timer1.Enabled = False
'Timer2.Enabled = False

'Exit Sub
If App.PrevInstance = True Then End

File1.Path = App.Path + "\WAVE\"
File1.Pattern = "*.WAV"
Debug.Print File1.Path
'Stop
    
    ' Set properties needed by MCI to open.
    MMControl1.Notify = True
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.DeviceType = "WaveAudio"
    Debug.Print File1.List(1)
    
    MMControl1.FileName = File1.Path + "\" + File1.List(2)
    Debug.Print File1.Path + "\" + File1.List(2)
    
    
    
    ' Open the MCI WaveAudio device.
    MMControl1.Command = "Open"
    'MMControl1.Command = "Play"


'H1 = True 'RANDOM
Slider1.Value = 1000
SLIDERHH = Slider1.Value
Me.WindowState = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
MMControl1.Command = "Close"
End
End Sub

Private Sub Label2_Click()

H1 = Not H1

End Sub

Private Sub MNU_VB_Click()
Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
End
End Sub

Private Sub Slider1_Scroll()
Timer1.Interval = 0
Timer1.Interval = Slider1.Value * 2

Timer2.Interval = 0
Timer2.Interval = Slider1.Value * 2

SLIDERHH = Slider1.Value

End Sub

Private Sub Timer1_Timer()
    
    
    If MM < Now Then
        'If MM = 0 Then MM = Now + TimeSerial(0, 1, 0): Exit Sub
        MM = Now + TimeSerial(0, 1, 0)
        MMControl1.Command = "Prev"
        MMControl1.Command = "Play"
        
    
    End If
    
    Exit Sub
    
    
    Timer1.Enabled = False
    Timer1.Interval = 0
    If H1 = True Then
        RANDOMH = Int((Rnd(1) * 150) + 1)
    Else
        RANDOMH = 0
    End If

    Timer1.Interval = (SLIDERHH) + RANDOMH
    
    Slider1.Value = Timer1.Interval
    Timer2.Interval = Timer1.Interval
    
    Timer2.Enabled = True
    Label3 = Str(Timer1.Interval + RANDOMH)
    
    MMControl1.Command = "Prev"
    MMControl1.Command = "Play"
    
End Sub

Private Sub Timer2_Timer()
    
    Timer2.Enabled = False
    Timer2.Interval = 0
    Timer2.Interval = (SLIDERHH) + RANDOMH

    Timer1.Enabled = True
    
    'If H1 = True And H2 < Now Then
    '    H2 = Now + TimeSerial(0, 0, Int((Rnd(1) * 2) + 1))
    'End If
    'If H1 = True And H2 >= Now Then Exit Sub
    
    MMControl1.Command = "Prev"
    MMControl1.Command = "Play"
    
End Sub

