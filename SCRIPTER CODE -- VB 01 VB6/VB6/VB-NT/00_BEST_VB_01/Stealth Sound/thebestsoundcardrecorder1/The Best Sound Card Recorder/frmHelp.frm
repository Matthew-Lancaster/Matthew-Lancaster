VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Help"
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "X"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtHelp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      IMEMode         =   3  'DISABLE
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   6255
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "Help"
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
      TabIndex        =   2
      Top             =   0
      Width           =   6500
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim Y As Integer
    Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

    txtHelp.Text = "The wave file produced, can't be played." & vbCrLf & vbCrLf & _
                   "To fix this problem:" & vbCrLf & vbCrLf & _
                   "You may need to download a music converter that supports the wave file's format. (Go to http://www.rarewares.org/dancer/dancer.php?f=33 and extract the lame.exe file into this directory and click the " & Chr(34) & "Convert .WAV to .MP3" & Chr(34) & " button on this program's main window)" & vbCrLf & _
                   "------------------------------------------------------------------------------" & vbCrLf & _
                   "The volume is way too low or can't be heard at all on the recordings." & vbCrLf & vbCrLf & _
                   "To fix this problem:" & vbCrLf & vbCrLf & _
                   "1. Doubleclick on the speaker icon on the system tray and change the sliders to where you want them." & vbCrLf & vbCrLf & _
                   "Or you can click Start, click Run, and type in " & """" & "SNDVOL32" & """ and press ENTER." & vbCrLf & vbCrLf & _
                   "2. Experiment with the sliders until you get the right volume."

    Me.AutoRedraw = True
    Me.DrawStyle = 6
    Me.DrawMode = 13
    Me.DrawWidth = 13
    Me.ScaleMode = 3

    Me.ScaleHeight = (256)
    For i = 0 To 255
        Me.Line (0, Y)-(Me.Width, Y + 1), (vbYellow) - (i * (vbYellow - vbBlack) / 255), BF
        Y = Y + 1
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call ReleaseCapture
        Call SendMessage(Me.hwnd, &HA1, 2, 0)
    End If
End Sub
