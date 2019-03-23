VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{9E12D4F7-7B6A-409A-B2E3-3425F2E3B086}#1.0#0"; "vaxvoipocx.ocx"
Begin VB.Form TalkChatForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Talk Chat"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "TalkChatForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4800
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   3840
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8414
            Object.ToolTipText     =   "TalkChat Status"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Container 
      BackColor       =   &H000080FF&
      Height          =   1215
      Left            =   158
      ScaleHeight     =   1155
      ScaleWidth      =   4425
      TabIndex        =   15
      Top             =   150
      Width           =   4485
      Begin VAXVOIPOCXLib.VaxVoipOcx VOIP 
         Height          =   435
         Left            =   1230
         TabIndex        =   20
         Top             =   210
         Visible         =   0   'False
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   767
         _StockProps     =   0
      End
      Begin VB.CommandButton BttnTalk 
         Caption         =   "Talk"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   285
         Picture         =   "TalkChatForm.frx":044A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   180
         Width           =   800
      End
      Begin VB.CommandButton BttnHold 
         Caption         =   "Hold"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   1755
         Picture         =   "TalkChatForm.frx":0754
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   180
         Width           =   800
      End
      Begin VB.CommandButton BttnCut 
         Caption         =   "Cut"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   3240
         Picture         =   "TalkChatForm.frx":0B96
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   180
         Width           =   800
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   1710
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2265
      Left            =   38
      TabIndex        =   0
      Top             =   1470
      Width           =   4725
      Begin MSWinsockLib.Winsock sckTalkChat 
         Left            =   1890
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   3262
         Picture         =   "TalkChatForm.frx":0FD8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   14
         Top             =   240
         Width           =   480
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   982
         Picture         =   "TalkChatForm.frx":1422
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   13
         Top             =   240
         Width           =   480
      End
      Begin VB.Frame Frame3 
         Height          =   780
         Left            =   2370
         TabIndex        =   8
         Top             =   1020
         Width           =   2265
         Begin MSComctlLib.Slider MicSlider 
            Height          =   585
            Left            =   30
            TabIndex        =   9
            Top             =   120
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1032
            _Version        =   393216
            LargeChange     =   2
            TickStyle       =   2
            TickFrequency   =   0
            TextPosition    =   1
         End
      End
      Begin VB.CheckBox MicMute 
         Caption         =   "Mute &Microphone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2370
         TabIndex        =   4
         Top             =   1860
         Width           =   1875
      End
      Begin VB.CheckBox VolumeMute 
         Caption         =   "Mute &Volume"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   90
         TabIndex        =   3
         Top             =   1860
         Width           =   1545
      End
      Begin VB.Frame Frame2 
         Height          =   780
         Left            =   90
         TabIndex        =   1
         Top             =   1020
         Width           =   2265
         Begin MSComctlLib.Slider VolumeSlider 
            Height          =   615
            Left            =   30
            TabIndex        =   2
            Top             =   120
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1085
            _Version        =   393216
            LargeChange     =   10
            TickStyle       =   2
            TickFrequency   =   0
            TextPosition    =   1
         End
      End
      Begin VB.Label lblMicrophone 
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3150
         TabIndex        =   12
         Top             =   780
         Width           =   495
      End
      Begin VB.Label lblMin 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   2550
         TabIndex        =   11
         Top             =   780
         Width           =   105
      End
      Begin VB.Label lblMax 
         AutoSize        =   -1  'True
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   4200
         TabIndex        =   10
         Top             =   780
         Width           =   315
      End
      Begin VB.Label lblVolume 
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1020
         TabIndex        =   7
         Top             =   780
         Width           =   495
      End
      Begin VB.Label lblMax 
         AutoSize        =   -1  'True
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   1860
         TabIndex        =   6
         Top             =   780
         Width           =   315
      End
      Begin VB.Label lblMin 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   780
         Width           =   105
      End
   End
End
Attribute VB_Name = "TalkChatForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Obj As New clsVolume
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Sub BttnCut_Click()
    VOIP.StopTalk
    BttnTalk.Enabled = True
    BttnCut.Enabled = False
    BttnHold.Enabled = False
End Sub

Private Sub BttnHold_Click()
    If BttnHold.Caption = "Hold" Then
        BttnHold.Caption = "Talk"
        VOIP.StopTalk
    Else
        BttnHold.Caption = "Hold"
        VOIP.StartTalk RemoteForm.sckIdentifier.RemoteHost, 1006
    End If
End Sub

Private Sub BttnTalk_Click()
    
    'start Voice conversation
    VOIP.StartTalk RemoteForm.sckIdentifier.RemoteHost, 1006
    BttnTalk.Enabled = False
    BttnHold.Enabled = True
    BttnCut.Enabled = True
End Sub

Private Sub Form_Load()
    'Me.Show
    'set VOIP component licence key
    VOIP.SetLicenceKey ("EVALUATE")
    
    VolumeSlider.Min = Obj.MinWaveVolume
    VolumeSlider.Max = Obj.MaxWaveVolume
    VolumeSlider.Value = Obj.WaveVolume
    VolumeSlider.TickFrequency = Obj.MaxWaveVolume / 100
    VolumeSlider.LargeChange = Obj.MaxWaveVolume * 10 / 100
    VolumeSlider.SmallChange = Obj.MaxWaveVolume * 5 / 100
    lblVolume.Caption = Format(VolumeSlider.Value / VolumeSlider.Max, "###0%")
    VolumeMute.Value = Obj.WaveMute
    
    MicSlider.Min = Obj.MinMicVolume
    MicSlider.Max = Obj.MaxMicVolume
    MicSlider.Value = Obj.MicroVolume
    MicSlider.TickFrequency = Obj.MaxMicVolume / 100
    MicSlider.LargeChange = Obj.MaxMicVolume * 10 / 100
    MicSlider.SmallChange = Obj.MaxMicVolume * 5 / 100
    lblMicrophone.Caption = Format(MicSlider.Value / MicSlider.Max, "###0%")
    MicMute.Value = Obj.MicroMute
End Sub

Private Sub MicMute_Click()
    If MicMute.Value = 1 Then
        MicMute.ForeColor = vbRed
        Obj.MicroMute = True
    Else
        MicMute.ForeColor = vbBlue
        Obj.MicroMute = False
    End If
End Sub

Private Sub MicSlider_Change()
    Obj.MicroVolume = MicSlider.Value
    lblMicrophone.Caption = Format(MicSlider.Value / MicSlider.Max, "###0%")
End Sub

Private Sub MicSlider_Scroll()
    Obj.MicroVolume = MicSlider.Value
    lblMicrophone.Caption = Format(MicSlider.Value / MicSlider.Max, "###0%")
End Sub

Private Sub Timer1_Timer()
Dim lStatus As Long
    
    'get the current connection status
    lStatus = VOIP.GetConnectionStatus
    
    'depending on status code display status message
    'and also set the status flags appropriately
    Select Case lStatus
        Case 100
            StatusBar1.Panels(1).Text = "NORMAL"
        Case 101
            StatusBar1.Panels(1).Text = "ESTABLISHING_CONNECTION"
        Case 102
            StatusBar1.Panels(1).Text = "CONNECTION_ESTABLISHED"
        Case 103
            StatusBar1.Panels(1).Text = "COULD_NOT_ESTABLISH"
            Call BttnCut_Click
        Case 104
            StatusBar1.Panels(1).Text = "VOICE_CONVERSATION_CLOSED"
            Call BttnCut_Click
        Case 105
            StatusBar1.Panels(1).Text = "CANNOT_OPEN_SOUND_OUTPUT_DEVICE"
            Call BttnCut_Click
        Case 106
            StatusBar1.Panels(1).Text = "CANNOT_OPEN_SOUND_INPUT_DEVICE"
            Call BttnCut_Click
        Case 107
            StatusBar1.Panels(1).Text = "INPUT_DEVICE_PROBLEM"
            Call BttnCut_Click
        Case 108
            StatusBar1.Panels(1).Text = "OUTPUT_DEVICE_PROBLEM"
            Call BttnCut_Click
        Case 109
            StatusBar1.Panels(1).Text = "CONNECTION_LOST"
            Call BttnCut_Click
    End Select

End Sub

Private Sub VolumeMute_Click()
    If VolumeMute.Value = 1 Then
        VolumeMute.ForeColor = vbRed
        Obj.WaveMute = True
    Else
        VolumeMute.ForeColor = vbBlue
        Obj.WaveMute = False
        'Beep
    End If
End Sub

Private Sub VolumeSlider_Change()
    Obj.WaveVolume = VolumeSlider.Value
    lblVolume.Caption = Format(VolumeSlider.Value / VolumeSlider.Max, "###0%")
    'Beep
End Sub

Private Sub VolumeSlider_Scroll()
    Obj.WaveVolume = VolumeSlider.Value
    lblVolume.Caption = Format(VolumeSlider.Value / VolumeSlider.Max, "###0%")
End Sub
