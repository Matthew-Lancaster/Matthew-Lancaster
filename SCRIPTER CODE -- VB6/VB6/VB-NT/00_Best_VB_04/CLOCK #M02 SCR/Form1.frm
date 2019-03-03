VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00560107&
   BorderStyle     =   0  'None
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   2565
      Left            =   150
      ScaleHeight     =   2505
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2850
      Top             =   1500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TT, R, TV
Const T As Double = 57.29577951

Private Sub Form_Click()
End
End Sub

Private Sub Form_Load()
Me.Height = 2048
Me.Width = 2048
Me.BackColor = &H560107
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.ScaleHeight = 1000
Me.ScaleWidth = 1000

Me.ScaleHeight = 1000
Me.ScaleWidth = 1000

Me.Top = 0
Me.Left = 0
Me.Width = 1000
Me.Height = 1000

End Sub

Private Sub Picture1_Click()
End
End Sub

Private Sub Timer1_Timer()
Dim H As Long, M As Long, S As Long                     'time units
Dim Hd As Double, Md As Double, Sd As Double            'Degrees
Dim Hr As Double, Mr As Double, Sr As Double            'Radians

Me.Cls
Me.Cls

H = Hour(Time): M = Minute(Time): S = Second(Time)

If H >= 12 Then H = H - 12

Hd = H * 30
Hd = Hd + M / 2
Md = M * 6
Sd = S * 6

Hd = Hd - 90: Md = Md - 90: Sd = Sd - 90

If Hd < 0 Then Hd = Hd + 360
If Md < 0 Then Md = Md + 360
If Sd < 0 Then Sd = Sd + 360

Hr = Hd / T: Mr = Md / T: Sr = Sd / T

Call CLOCKNUMBERHOUR
Call CLOCKNUMBERMINUTE


TT = Me.ScaleHeight / 2
Me.DrawWidth = 13
Me.Line (TT, TT)-(Me.ScaleHeight / 2 + ((Me.ScaleHeight / 2) * 0.5 * Cos(Hr)), Me.ScaleWidth / 2 + ((Me.ScaleHeight / 2) * 0.5 * Sin(Hr))), vbWhite
Me.Line (TT, TT)-(Me.ScaleHeight / 2 + ((Me.ScaleHeight / 2) * 0.6 * Cos(Mr)), Me.ScaleWidth / 2 + ((Me.ScaleHeight / 2) * 0.6 * Sin(Mr))), vbBlue
Me.DrawWidth = 13
Me.Line (TT, TT)-(Me.ScaleHeight / 2 + ((Me.ScaleHeight / 2) * 0.7 * Cos(Sr)), Me.ScaleWidth / 2 + ((Me.ScaleHeight / 2) * 0.7 * Sin(Sr))), vbGreen

Me.DrawWidth = 8
Me.Circle (Me.ScaleWidth / 2, Me.ScaleHeight / 2), 5, vbWhite
End Sub


Sub CLOCKNUMBERHOUR()

Dim H As Long, M As Long, S As Long                     'time units
Dim Hd As Double, Md As Double, Sd As Double            'Degrees
Dim Hr As Double, Mr As Double, Sr As Double            'Radians

For R = 0 To 12

'Me.Cls
H = R: M = Minute(0): S = Second(0)

If H >= 12 Then H = H - 12

Hd = H * 30
Hd = Hd + M / 2
Md = M * 6
Sd = S * 6

Hd = Hd - 90: Md = Md - 90: Sd = Sd - 90

If Hd < 0 Then Hd = Hd + 360
If Md < 0 Then Md = Md + 360
If Sd < 0 Then Sd = Sd + 360

Hr = Hd / T: Mr = Md / T: Sr = Sd / T




TT = Me.ScaleHeight / 2
TV = 370 'Int((ME.ScaleHeight / 2) / 1.39)
Me.FillColor = vbWhite
Me.FillStyle = vbSolid
Me.DrawWidth = 1
'Line (TT, TT)-(Me.ScaleHeight / 2 + ((Me.ScaleHeight / 2) * 0.5 * Cos(Hr)), Me.ScaleWidth / 2 + ((Me.ScaleHeight / 2) * 0.5 * Sin(Hr))), vbGreen
Me.Circle (TT + (Cos(Hr) * TV), TT + (Sin(Hr) * TV)), 13, vbWhite


Next

End Sub

Sub CLOCKNUMBERMINUTE()
Exit Sub
Dim H As Long, M As Long, S As Long                     'time units
Dim Hd As Double, Md As Double, Sd As Double            'Degrees
Dim Hr As Double, Mr As Double, Sr As Double            'Radians

For R = 0 To 59

'Me.Cls
H = 1: M = R: S = Second(0)

If H >= 12 Then H = H - 12

Hd = H * 30
Hd = Hd + M / 2
Md = M * 6
Sd = S * 6

Hd = Hd - 90: Md = Md - 90: Sd = Sd - 90

If Hd < 0 Then Hd = Hd + 360
If Md < 0 Then Md = Md + 360
If Sd < 0 Then Sd = Sd + 360

Hr = Hd / T: Mr = Md / T: Sr = Sd / T


TT = Me.ScaleHeight / 2
'TT = TT + Me.ScaleHeight
TV = 370 'Int((ME.ScaleHeight / 2) / 1.39)

Me.FillColor = vbWhite
Me.FillStyle = vbSolid
Me.DrawWidth = 1
'Line (TT, TT)-(Me.ScaleHeight / 2 + ((Me.ScaleHeight / 2) * 0.5 * Cos(Hr)), Me.ScaleWidth / 2 + ((Me.ScaleHeight / 2) * 0.5 * Sin(Hr))), vbGreen
Me.Circle (TT + (Cos(Mr) * TV), TT + (Sin(Mr) * TV)), 3, vbWhite


Next

End Sub

