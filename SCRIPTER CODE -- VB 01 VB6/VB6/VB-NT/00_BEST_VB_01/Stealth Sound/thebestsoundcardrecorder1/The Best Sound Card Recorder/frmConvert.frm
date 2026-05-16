VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConvert 
   BorderStyle     =   0  'None
   Caption         =   "Convert .WAV to .MP3"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   3000
      Width           =   4095
   End
   Begin MSComDlg.CommonDialog cdgBrowse 
      Left            =   3000
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "X"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtWaveFile 
      Height          =   525
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label lblWave 
      BackStyle       =   0  'Transparent
      Caption         =   ".WAV File:"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Convert .WAV to .MP3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    On Error GoTo Error_Handler
    cdgBrowse.CancelError = True
    cdgBrowse.Filter = "Wave Files(.WAV)|*.wav|All Files|*.*"
    cdgBrowse.Flags = &H2 Or &H400
    cdgBrowse.ShowOpen
    txtWaveFile.Text = cdgBrowse.Filename
Error_Handler:
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdConvert_Click()
    Dim FileNum As Integer
    On Error GoTo NotFinished
    cmdConvert.Caption = "Please Wait...Now Converting..."
    cmdConvert.Enabled = False
    Shell App.Path & "\lame.exe " & Chr(34) & txtWaveFile.Text & Chr(34), vbHide
Start:
    Sleep 1000
    FileNum = FreeFile
    Open App.Path & "\lame.exe" For Binary Access Write As #FileNum
    Close #FileNum
    MsgBox txtWaveFile.Text & " was successfully converted to " & txtWaveFile.Text & ".mp3", vbInformation, "Success"
    cmdConvert.Caption = "Convert"
    cmdConvert.Enabled = True
    Exit Sub
NotFinished:     GoTo Start
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim Y As Integer
    Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Me.AutoRedraw = True
    Me.DrawStyle = 6
    Me.DrawMode = 13
    Me.DrawWidth = 13
    Me.ScaleMode = 3

    Me.ScaleHeight = (256)
    For i = 0 To 255
        Me.Line (0, Y)-(Me.Width, Y + 1), (vbBlack) - (i * (vbBlack - vbGreen) / 255), BF
        Y = Y + 1
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
End Sub
