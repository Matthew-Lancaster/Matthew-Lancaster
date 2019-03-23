VERSION 5.00
Begin VB.Form Keleidoscope 
   AutoRedraw      =   -1  'True
   Caption         =   "Pattern"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   737
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer HighRes 
      Interval        =   1
      Left            =   1755
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1350
      Top             =   570
   End
End
Attribute VB_Name = "Keleidoscope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Visual Basic For Dos Easy Converted
'Displays an animated tunnel fly through
'With kelidoscope of colours
'By Matt Lan 11-Apr-2004



Public awat, pi, lk, cv, tyu, cb, c, c2, h
Public wk, hw, kw, Size, x3, y3


Private Sub Form_GotFocus()
'Timer1.Enabled = True
HighRes.Enabled = True
End Sub

Private Sub Form_Load()

pi = 4 * Atn(1)
lk = 150
cv = 0.0000001
cv = 0.00002



pi = 4 * Atn(1)
twd1 = 2
twd2 = 2.1
twd3 = 2.2
BackColor = RGB(255, 255, 255)
BackColor = RGB(0, 0, 0)
Me.Height = 7000
Me.Width = 6000
x3 = (Screen.Width / Screen.TwipsPerPixelX) / 2
y3 = (Screen.Height / Screen.TwipsPerPixelY) / 2
y3 = y3 - 40

DW = 50
DrawWidth = DW    ' Use wider brush.




Size = 320
Size = Size - (DW / 2.6)

Exit Sub

Steps1 = 0.01: Steps2 = 0.03 'tri
Steps1 = 0.02: Steps2 = 0.09 'tri
Dot_Or_Line = 1 'line
'Dot_Or_Line = 2'dot

End Sub

Private Sub Form_LostFocus()
Timer1.Enabled = False
HighRes.Enabled = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
      'PSet (x, y)      ' Draw a point.
  'x3 = x: y3 = y
 '  DrawWidth = 90      ' Use wider brush.
   ForeColor = RGB(kw, hw, w)   ' Set drawing color.
End Sub


Private Sub Timer1_Timer()


h1 = -10: he1 = Abs(h1)
h2 = -20: he2 = Abs(h2)
h3 = -30: he3 = Abs(h3)

wk = wk + twd1
If wk >= 255 - he1 Then twd1 = h1
If wk < he1 Then twd1 = he1
hw = hw + twd2
If hw >= 255 - he2 Then twd2 = h2
If hw < he2 Then twd2 = he2
kw = kw + twd3
If kw >= 255 - he3 Then twd3 = h2
If kw < he3 Then twd3 = he3

ForeColor = RGB(Int(kw), Int(hw), Int(wk))   ' Set drawing color.

rr3 = rr3 + Steps1
tr = tr + Steps2
lw = Size * Sin(tr)
x4 = (lw) * Sin(rr3) + x3
y4 = (lw) * Cos(rr3) + y3
'x5 = Size * Sin(rr3 - 3) + x3
'y5 = Size * Cos(rr3 - 3) + y3
     
If Dot_Or_Line = 1 Then Line Step(0, 0)-(x4, y4)
If Dot_Or_Line = 2 Then PSet (x4, y4)    ' Draw a point.
 
   
   'Circle (x4, y4), 10
   'DrawWidth = 30    ' Use wider brush.
 '  ForeColor = RGB(0, 255, 0) 'RGB(kw, hw, w)   ' Set drawing color.

End Sub




Sub HighRes_timer()


Rem FOR h = 0 TO lk STEP cv

 h = h + cv
 'If h > lk Then Exit Sub

    Rem IF h > 1 THEN h = h + .0001: GOTO jump
    Rem IF h <= 1 THEN h = h + .0002: GOTO jump

jump:

'    LOCATE 1, 1: Print Format$(h, "0000.000000000000000"); " Of"; lk; " Keep Pressing key to Exit "

    q = h
    tyu = tyu + 0.003
    r = tyu
    Rem r = 0
    wk = 10
    hw = 10
    kw = 10
    c = cb / 20: Rem c / 10000
    c2 = cb: Rem c2 / 10000

    cb = cb + 0.01

    cb2 = (100000 + (Sin((cb)))) / 2

    c = c + cb2
    c2 = c2 + cb2

    Do
        Rem c2 = c2 + (cv)
        Rem gh = 300 * (SIN(c2)) + 100

        'c = c + 0.001: Rem (cv * 1000) + ((COS(c2)) / gh)
        'If c > 15 Then c = 0
        'c1 = 14 * (Abs(Cos(c)))
        
        r = r + q
        w = ((r / (pi)) / ((h)))  '(h * 2))'h * 1.5))
        x = (w * 1.7) * Sin(r * 5) + x3
        y = (w) * Cos(r * 5) + y3

        'h1 = -10: he1 = Abs(h1)
        'h2 = -20: he2 = Abs(h2)
        'h3 = -30: he3 = Abs(h3)

        'wk = wk + twd1
        'If wk >= 255 - he1 Then twd1 = h1
        'If wk < he1 Then twd1 = he1
        'hw = hw + twd2
        'If hw >= 255 - he2 Then twd2 = h2
        'If hw < he2 Then twd2 = he2
        'kw = kw + twd3
        'If kw >= 255 - he3 Then twd3 = h2
        'If kw < he3 Then twd3 = he3

        c = c + 0.05
        kw = 200 * Sin(c) + 220
        hw = 200 * Sin(c * 1.1) + 220
        wk = 200 * Sin(c * 1.2) + 220

        ForeColor = RGB(Int(kw), Int(hw), Int(wk))   ' Set drawing color.
        
        PSet (x, y)
        Rem CIRCLE (x, y), 2, c1
        Rem LINE STEP(0, 0)-(x, y), c1
        s1 = 0
        Rem wl = 100
        Rem IF x < wl OR y < wl THEN s1 = 1
        Rem IF x > 640 - wl OR y > 480 - wl THEN s1 = 1
        Rem IF x > 590 OR y > 430 THEN s1 = 1

        If w > Size Then s = s + 1 Else s = 0

    Loop Until s > 100
    s = 0

End Sub

