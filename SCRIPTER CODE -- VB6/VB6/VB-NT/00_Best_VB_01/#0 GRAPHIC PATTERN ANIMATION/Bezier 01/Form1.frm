VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4110
      Left            =   0
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   372
      TabIndex        =   6
      Top             =   435
      Width           =   5640
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   4680
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.CommandButton Zufallskurve 
         Caption         =   "Zufallskurve"
         Height          =   345
         Left            =   60
         TabIndex        =   3
         Top             =   45
         Width           =   1290
      End
      Begin VB.CheckBox Clear1 
         Caption         =   "Clear"
         Height          =   195
         Left            =   4095
         TabIndex        =   2
         Top             =   0
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Thin"
         Height          =   150
         Left            =   4095
         TabIndex        =   1
         Top             =   225
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Copyright © by Jürgen Anke (1998)"
         Height          =   210
         Left            =   1485
         TabIndex        =   5
         Top             =   105
         Width           =   4485
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Width           =   1365
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5880
      Top             =   1290
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim WTrue, HWTrue, KWTrue, TW1, TW2, TW3, FR
Dim i, RR, TT, TW, OTW, DD, DF, i2

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function PolyBezier Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long

Dim ptBez1() As POINTAPI
Dim ptBez2() As POINTAPI

Sub DrawBezier()
    
    Call PolyBezier(Picture1.hdc, ptBez1(0), UBound(ptBez1) - 1)
    Call PolyBezier(Picture1.hdc, ptBez2(0), UBound(ptBez2) - 1)
    
End Sub

Private Sub ColorCycle()


DD = -3
DF = Abs(DD)
WTrue = WTrue + 0.09
TW1 = 127 * Sin(WTrue) + 127
TW2 = 127 * Cos(WTrue * 1.3) + 127
TW3 = 127 * Sin(WTrue * 1.6) + 127
   
Picture1.ForeColor = RGB(TW1, TW2, TW3)   ' Set drawing color.
'picture1.ForeColor = vbWhite
'Picture1.ForeColor = RGB(255 - KWTrue, 255 - HWTrue, 255 - WTrue)
End Sub


Sub RandomPoints()
    
    Dim MaxX As Long, MaxY As Long
    TW = 10
    Call ColorCycle
    
    
    If OTW <> TW Then
        ReDim ptBez1(TW)
        ReDim ptBez2(TW)
    End If
    
    'Picture1.ForeColor = vbBlack
    If FR = True And Clear1 = vbChecked Then
    Call PolyBezier(Picture1.hdc, ptBez1(0), UBound(ptBez1))
    'Call PolyBezier(Picture1.hdc, ptBez2(0), UBound(ptBez2) - 1)
    End If
    FR = True
    Picture1.ForeColor = vbWhite

    If Check1 = vbChecked Then
    Picture1.DrawWidth = 3
    Else
    Picture1.DrawWidth = 10
    
    End If

    
    
    OTW = TW
    'Randomize
    MaxX = Picture1.ScaleWidth
    MaxY = Picture1.ScaleHeight
    RR = RR + 0.01
    TT = RR

    For i = 0 To TW
            TT = TT + 0.001 + (Cos(RR / 5)) + (RR / 5)
  
                ptBez1(i).X = ((MaxX / 2) * Sin(TT) + (MaxX / 2)) / 1#
                ptBez1(i).Y = ((MaxY / 2) * Cos(TT * 1.05) + (MaxY / 2)) / 1.01


    Next
    
    ptBez1(TW).X = ptBez1(0).X
    ptBez1(TW).Y = ptBez1(0).Y
    
    
   
    
End Sub


Sub RandomPoints_MACH1()
    
    Picture1.ForeColor = vbBlack
    If FR = True And Clear1 = vbChecked Then
    Call PolyBezier(Picture1.hdc, ptBez1(0), UBound(ptBez1) - 1)
    Call PolyBezier(Picture1.hdc, ptBez2(0), UBound(ptBez2) - 1)
    End If
    FR = True
    Picture1.ForeColor = vbWhite

    If Check1 = vbChecked Then
    Picture1.DrawWidth = 3
    Else
    Picture1.DrawWidth = 10
    
    End If

    Dim MaxX As Long, MaxY As Long
    TW = 8
    Call ColorCycle
    If OTW <> TW Then
        ReDim ptBez1(TW)
        ReDim ptBez2(TW)
    End If
    OTW = TW
    'Randomize
    MaxX = Picture1.ScaleWidth
    MaxY = Picture1.ScaleHeight
    RR = RR + 0.01
    TT = RR
    For i2 = 1 To 1
        For i = 0 To TW
            TT = TT + 0.001 + (Cos(RR / 5)) + (RR / 5)
            If i2 = 1 Then
                ptBez1(i).X = ((MaxX / 2) * Sin(TT) + (MaxX / 2)) / 1#
                ptBez1(i).Y = ((MaxY / 2) * Cos(TT * 1.05) + (MaxY / 2)) / 1.01
            End If
            If i2 = 2 Then
                ptBez2(i).X = ((MaxX / 2) * Sin(TT) + (MaxX / 2)) / 1#
                ptBez2(i).Y = ((MaxY / 2) * Cos(TT * 1.05) + (MaxY / 2)) / 1.01
        End If
    Next
    Next
    
End Sub



Private Sub Timer1_Timer()
Call Zufallskurve_Click
End Sub

Private Sub Zufallskurve_Click()

    RandomPoints
    DrawBezier
    
End Sub
Private Sub Form_Load()
    
    RR = 10
    Me.Width = Screen.Width '/ Screen.TwipsPerPixelX - 100
    Me.Height = Screen.Height - 1100 ' / Screen.TwipsPerPixelY - 100
    Me.Top = 400
    Me.Left = 0
    Picture1.ForeColor = vbWhite
    Picture1.DrawWidth = 3
    RandomPoints
    Me.Show
End Sub


Private Sub Form_Resize()
Picture1.Move Picture1.Left, Picture1.Top, ScaleWidth, ScaleHeight - Picture2.Height

If Me.WindowState Mod 2 = 0 Then Call TIMEROUT(True): Exit Sub
If Me.WindowState = 1 Then Call TIMEROUT(False): Exit Sub

End Sub

Private Sub TIMEROUT(TF)
    On Local Error Resume Next
    For Each Form In Forms
        For Each Control In Controls
            If Control.Interval > 0 Then Control.Enabled = TF
        Next
    Next
End Sub
