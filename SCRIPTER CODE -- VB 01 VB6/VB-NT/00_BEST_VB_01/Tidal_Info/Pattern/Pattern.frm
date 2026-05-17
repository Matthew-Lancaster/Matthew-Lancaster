VERSION 5.00
Begin VB.Form YinYang 
   AutoRedraw      =   -1  'True
   Caption         =   "Yin Yang Pattern"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   211
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1575
      Top             =   2025
   End
End
Attribute VB_Name = "YinYang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pi, TwD1, TwD2, TwD3, BackColor4, X3, Y3, dw, Size, Steps1, Dot_Or_Line, TTI
Dim KW, HW, WK, RR3, TRZ, LW, X4, Y4, StartX, i, TT, GG
'public XPud


Private Sub Form_Load()
pi = 4 * Atn(1)
TwD1 = 0.5
TwD2 = 0.7
TwD3 = 0.9
BackColor4 = RGB(255, 255, 255)
BackColor4 = RGB(0, 0, 0)
'Form1.Height = 7000
'Form1.Width = 6000
X3 = (Screen.Width / Screen.TwipsPerPixelX) / 2
Y3 = (Screen.Height / Screen.TwipsPerPixelY) / 2
Y3 = Y3 - 33

frmPassLock.ForeColor = RGB(0, 0, 0)

Size = 380
Size = Size - (dw / 2.6)
Steps1 = 0.01: Steps2 = 0.03 'tri
Steps1 = 0.02: Steps2 = 0.04 'tri
Dot_Or_Line = 1 'line
'Dot_Or_Line = 2 'dot
ReDim ptBez1(26)
GG = False
TT = 0
End Sub

Private Sub Timer1_Timer()

'frmPassLock.hide

'h1 = -10: he1 = Abs(h1)
'h2 = -20: he2 = Abs(h2)
'h3 = -30: he3 = Abs(h3)

'wk = wk + twd1
'If wk >= 255 - he1 Then twd1 = h1
'If wk < he1 Then twd1 = he1
'hw = hw + twd1 * 1.2
'If hw >= 255 - he2 Then twd2 = h2
'If hw < he2 Then twd2 = he2
'kw = kw + twd1 * 1.5
'If kw >= 255 - he3 Then twd3 = h2
'If kw < he3 Then twd3 = he3
'frmPassLock.Hide
TwD1 = TwD1 + 0.005 + ((frmPassLock.Slider5.Value / 1000))
TTI = TwD1 + (frmPassLock.Slider5.Value / 500)

'1=Colour speed
'2=draw Width
'3=Spin Speed
'4=Seed colur
'5=Colour Sine

KW = 125 * Sin(TTI) + 125
HW = 125 * Sin(TTI * 2) + 125
WK = 125 * Sin(TTI * 5) + 125

RR3 = RR3 + Steps1 + ((frmPassLock.Slider3.Value / 500))

TRZ = TRZ + Steps2 + ((frmPassLock.Slider3.Value / 500)) + (frmPassLock.Slider4.Value / 8000)
LW = Size * Sin(TRZ)
X4 = (LW * 1.7) * Sin(RR3) + X3
Y4 = (LW * 1) * Cos(RR3) + Y3
'x5 = Size * Sin(rr3 - 3) + x3
'y5 = Size * Cos(rr3 - 3) + y3
     
frmPassLock.DrawWidth = frmPassLock.Slider2.Value
     
     
     
     
     
If XPud = False Then
    XPud = True
    frmPassLock.ForeColor = RGB(0, 0, 0)
    frmPassLock.PSet (X4, Y4)
'    frmPassLock.ForeColor = RGB(Int(KW), Int(HW), Int(WK))   ' Set drawing color.
End If
     
If frmPassLock.Check2.Value = vbChecked Then
    frmPassLock.PSet (X4, Y4)
End If
        
If frmPassLock.Check3.Value = vbChecked Then
    frmPassLock.DrawWidth = 3
    frmPassLock.Circle (X4, Y4), frmPassLock.Slider2.Value
End If
        
If frmPassLock.Check4.Value = vbChecked Then
    frmPassLock.Line Step(0, 0)-(X4, Y4)
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
    ptBez1(0).x = X4
    ptBez1(0).Y = Y4
    Call PolyBezier(frmPassLock.HDC, ptBez1(0), UBound(ptBez1) - 1)

End If


frmPassLock.ForeColor = RGB(Int(KW), Int(HW), Int(WK))   ' Set drawing color.

End Sub
