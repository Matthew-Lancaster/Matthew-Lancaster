VERSION 5.00
Begin VB.Form frmCapPal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capture Palette"
   ClientHeight    =   2580
   ClientLeft      =   2205
   ClientTop       =   1890
   ClientWidth     =   3720
   Icon            =   "CapPal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtColors 
      Height          =   285
      Left            =   765
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "256"
      Top             =   1245
      Width           =   420
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   2505
      TabIndex        =   2
      Top             =   2010
      Width           =   960
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   360
      Left            =   1335
      TabIndex        =   1
      Top             =   2010
      Width           =   960
   End
   Begin VB.CommandButton cmdFrame 
      Caption         =   "&Frame"
      Height          =   360
      Left            =   165
      TabIndex        =   0
      Top             =   2010
      Width           =   960
   End
   Begin VB.Label lblFrames 
      Alignment       =   2  'Center
      Caption         =   "0 Frames"
      Height          =   600
      Left            =   1320
      TabIndex        =   5
      Top             =   1290
      Width           =   2280
   End
   Begin VB.Label lblColors 
      Caption         =   "Colors:"
      Height          =   240
      Left            =   135
      TabIndex        =   4
      Top             =   1275
      Width           =   510
   End
   Begin VB.Label lblStatic 
      Caption         =   $"CapPal.frx":000C
      Height          =   870
      Left            =   135
      TabIndex        =   3
      Top             =   105
      Width           =   3435
   End
End
Attribute VB_Name = "frmCapPal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'form level flag to indicate whether
'we need to close the palette capture on unload
Private fManual As Boolean
'form level flag to record number of frames captured in manual mode
Private numManFrames As Long

Private Sub Form_Load()
'load num pal colors from registry
    txtColors.Text = GetSetting(App.Title, "palette", "numcolors", "256")
End Sub

Private Sub cmdFrame_Click()
    fManual = True
    Call capPaletteManual(frmMain.capwnd, False, Val(txtColors.Text))
    numManFrames = numManFrames + 1
    lblFrames.Caption = numManFrames & " Frames"
    cmdCancel.Caption = "&Close"
End Sub

Private Sub cmdStart_Click()
    Const numFrames As Long = 100 'change this value to sample more or less frames
    numManFrames = 0 'reset manual capture counter if necessary
    fManual = False
    lblFrames.Caption = "Sampling - please wait..."
    lblFrames.Refresh
    cmdFrame.Enabled = False
    Call capPaletteAuto(frmMain.capwnd, numFrames, Val(txtColors.Text))
    lblFrames.Caption = "Finished! - " & numFrames & " frames sampled"
    cmdFrame.Enabled = True
    cmdCancel.Caption = "&OK"
End Sub

Private Sub txtColors_KeyPress(KeyAscii As Integer)
    'allow backspace key
    If KeyAscii = 8 Then Exit Sub
    'logic to keep the user input valid
    If KeyAscii < 48 Then KeyAscii = 0
    If KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtColors_LostFocus()
    'Input Filter
    If Val(txtColors.Text) < 16 Then txtColors.Text = 16
    If Val(txtColors.Text) > 256 Then txtColors.Text = 256
End Sub

Private Sub cmdCancel_Click()
    If fManual Then
        'close manual palette capture by sending false
        Call capPaletteManual(frmMain.capwnd, False, Val(txtColors.Text))
    End If
    If cmdCancel.Caption <> "&Cancel" Then 'save num colors to registry
        Call SaveSetting(App.Title, "palette", "numcolors", txtColors.Text)
    End If
    Unload Me
End Sub
