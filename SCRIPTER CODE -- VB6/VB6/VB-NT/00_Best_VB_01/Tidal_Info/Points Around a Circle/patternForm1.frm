VERSION 5.00
Begin VB.Form Vector 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vector Pattern"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1815
      Top             =   1425
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   960
   End
End
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim R, Theta, n, i, j, TT, GG
Public X0, Y0, XN, YN, IT3, N0, ExitVector, V As Double
Dim Vert%()


Private Sub Form_Load()

ExitVector = False
ReDim Vert%(2000)

'INPUT "enter radius (range 10-170)"; r
'INPUT "enter number of vertices (range 3-50)"; v

End Sub

Private Sub Timer1_Timer()
'frmPassLock.Hide

If ExitVector = True Then Exit Sub

DrawWidth = 5 ' Use wider brush.
frmPassLock.DrawWidth = (frmPassLock.Slider2.Value / 2) + 1

R = 370 'Size
If V = 0 Then V = 1

'For v = 1 To 500 Step 5
V = V + 0.01
If V > 40 Then V = V + 0.1
If V > 80 Then V = V + 1
If V > 160 Then V = V + 2
If V > 320 Then V = V + 4
If V > 1000 Then V = 1

'Form1.Height/2
X0 = (Screen.Width / Screen.TwipsPerPixelX) / 2
Y0 = (Screen.Height / Screen.TwipsPerPixelY) / 2
X0 = X0 + 20
Y0 = Y0 - 50

Theta = 6.28319 / V: n = V - 1

'THETA = 5 / v: N = v - 1
For i = 1 To n
DoEvents
Vert%(2 * i - 1) = R * 1.5 * Cos(Theta * i) + X0
Vert%(2 * i) = R * Sin(Theta * i) + Y0
PSet (Vert%(2 * i - 1) - 40, 50 + Vert%(2 * i)), 1
Next i
N0 = n

For j = 1 To n

DoEvents

WK = WK + TwD1
If WK > 255 Then TwD1 = -6: WK = WK + TwD1
If WK < 1 Then TwD1 = 6: WK = WK + TwD1
HW = HW + TwD2
If HW > 255 Then TwD2 = -7: HW = HW + TwD2
If HW < 1 Then TwD2 = 7: HW = HW + TwD2
KW = KW + TwD3
If KW > 255 Then TwD3 = -8: KW = KW + TwD3
If KW < 1 Then TwD3 = 8: KW = KW + TwD3
   
frmPassLock.ForeColor = RGB(KW, HW, WK)    ' Set drawing color.

IT3 = 1
ExitVector = True

Do
    DoEvents
    If App.title = "Tidal Information..." Then ExitVector = False
Loop Until ExitVector = False

If App.title = "Tidal Information..." Then
    Timer1.Enabled = False
    Timer2.Enabled = False
    Exit Sub
End If


X0 = XN: Y0 = YN: N0 = N0 - 1
Next j

'Next

'LOCATE 1, 1
'INPUT "Try another (Y/N)"; A$
'IF A$ = "Y" OR A$ = "y" GOTO 40
'CLS
'END

End Sub

Private Sub Timer2_Timer()
If ExitVector = False Then Exit Sub

'For it3 = 1 To N0
'DoEvents

XN = Vert%(2 * IT3 - 1): YN = Vert%(2 * IT3)

If frmPassLock.Check7.Value = vbUnchecked Then
frmPassLock.Line (X0 - 40, 50 + Y0)-(XN - 40, 50 + YN)
End If

If frmPassLock.Check7.Value = vbChecked Then
    TT = TT + 1
    If TT >= 27 Then
        TT = 27
        If GG = False Then
            GG = True
            ReDim Preserve ptBez1(26)
        End If
    End If
    If TT < 27 Then
        ReDim Preserve ptBez1(TT - 1)
    End If
    For i = TT - 2 To 0 Step -1
        If ptBez1(i).x > 0 Then
            ptBez1(i + 1).x = ptBez1(i).x
            ptBez1(i + 1).Y = ptBez1(i).Y
        End If
    Next
    ptBez1(0).x = XN - 40
    ptBez1(0).Y = 50 + YN
    Call PolyBezier(frmPassLock.HDC, ptBez1(0), UBound(ptBez1) - 1)

End If



IT3 = IT3 + 1

If IT3 > N0 Then
    ExitVector = False
End If

'Next i

End Sub
