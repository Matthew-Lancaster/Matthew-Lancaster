VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ball Shader"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbPreset 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2415
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1080
      Width           =   2280
   End
   Begin VB.CommandButton cmdDraw 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Draw"
      Height          =   375
      Left            =   2415
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   2325
   End
   Begin VB.HScrollBar hscBlue 
      Height          =   255
      Left            =   2415
      Max             =   255
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.HScrollBar hscGreen 
      Height          =   255
      Left            =   2415
      Max             =   255
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.HScrollBar hscRed 
      Height          =   255
      Left            =   2415
      Max             =   255
      TabIndex        =   1
      Top             =   0
      Value           =   255
      Width           =   1695
   End
   Begin VB.PictureBox picCircle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   0
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   162
      TabIndex        =   0
      Top             =   0
      Width           =   2430
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   2415
      Top             =   1800
      Width           =   2280
   End
   Begin VB.Label lblBlue 
      BackColor       =   &H80000007&
      Caption         =   "Blue"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   720
      Width           =   645
   End
   Begin VB.Label lblGreen 
      BackColor       =   &H80000007&
      Caption         =   "Green"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   360
      Width           =   645
   End
   Begin VB.Label lblRed 
      BackColor       =   &H80000007&
      Caption         =   "Red"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   0
      Width           =   645
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbPreset_Click()
Select Case Format(cmbPreset.Text, "<")
    Case "red"
        hscRed.Value = 255
        hscGreen.Value = 0
        hscBlue.Value = 0
    Case "green"
        hscRed.Value = 0
        hscGreen.Value = 255
        hscBlue.Value = 0
    Case "blue"
        hscRed.Value = 0
        hscGreen.Value = 0
        hscBlue.Value = 255
    Case "yellow"
        hscRed.Value = 255
        hscGreen.Value = 255
        hscBlue.Value = 0
    Case "orange"
        hscRed.Value = 255
        hscGreen.Value = 150
        hscBlue.Value = 0
    Case "pink"
        hscRed.Value = 255
        hscGreen.Value = 0
        hscBlue.Value = 255
    Case "cyan"
        hscRed.Value = 0
        hscGreen.Value = 255
        hscBlue.Value = 255
    Case "purple"
        hscRed.Value = 125
        hscGreen.Value = 0
        hscBlue.Value = 125
    Case "white"
        hscRed.Value = 255
        hscGreen.Value = 255
        hscBlue.Value = 255
    Case "grey"
        hscRed.Value = 150
        hscGreen.Value = 150
        hscBlue.Value = 150
    Case "presets..."
        Exit Sub
End Select
cmdDraw_Click
End Sub

Private Sub cmdDraw_Click()
picCircle.Picture = picCircle.Image
picCircle.Cls
RedStep = 255 / (hscRed.Value + 1)
GreenStep = 255 / (hscGreen.Value + 1)
BlueStep = 255 / (hscBlue.Value + 1)
For i = 0 To 255
    picCircle.ForeColor = RGB(i / RedStep, i / GreenStep, i / BlueStep)
    picCircle.FillColor = picCircle.ForeColor
    radius = (255 - (i)) / 4
    picCircle.Circle (picCircle.ScaleWidth / 2, picCircle.ScaleHeight / 2), radius
Next i
Clipboard.Clear
Clipboard.SetData picCircle.Picture, vbCFBitmap
End Sub

Private Sub Form_Load()
picCircle.Picture = picCircle.Image
'cmbPreset.Text = "Presets..."
cmdDraw_Click
cmbPreset.Width = hscRed.Width
End Sub
