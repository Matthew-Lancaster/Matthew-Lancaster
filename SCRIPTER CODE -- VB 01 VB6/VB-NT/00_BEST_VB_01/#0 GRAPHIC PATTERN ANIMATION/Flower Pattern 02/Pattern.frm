VERSION 5.00
Begin VB.Form PatternFrm 
   AutoRedraw      =   -1  'True
   Caption         =   "Pattern"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   692
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2160
      Top             =   7680
   End
End
Attribute VB_Name = "PatternFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
pi = 4 * Atn(1)
tw1 = 1
tw2 = 1.1
tw3 = 1.2
BackColor = RGB(255, 255, 255)
BackColor = RGB(0, 0, 0)
Me.Height = 7000
Me.Width = 6000
x3 = (Screen.Width / Screen.TwipsPerPixelX) / 2
y3 = (Screen.Height / Screen.TwipsPerPixelY) / 2
y3 = y3 - 33

DW = 150
DrawWidth = DW    ' Use wider brush.


Size = 350
Size = Size - (DW / 2.6)
Steps1 = 0.01: Steps2 = 0.03 'tri
Steps1 = 0.02: Steps2 = 0.09 'tri
Dot_Or_Line = 1 'line
'Dot_Or_Line = 2'dot

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

w = w + tw1
If w >= 255 - he1 Then tw1 = h1
If w < he1 Then tw1 = he1
hw = hw + tw2
If hw >= 255 - he2 Then tw2 = h2
If hw < he2 Then tw2 = he2
kw = kw + tw3
If kw >= 255 - he3 Then tw3 = h2
If kw < he3 Then tw3 = he3

ForeColor = RGB(Int(kw), Int(hw), Int(w))   ' Set drawing color.

rr3 = rr3 + Steps1
tr = tr + Steps2
lw = Size * Sin(tr)
x4 = lw * Sin(rr3) + x3
y4 = lw * Cos(rr3) + y3
'x5 = Size * Sin(rr3 - 3) + x3
'y5 = Size * Cos(rr3 - 3) + y3
     
If Dot_Or_Line = 1 Then Line Step(0, 0)-(x4, y4)
If Dot_Or_Line = 2 Then PSet (x4, y4)    ' Draw a point.
 
   
   'Circle (x4, y4), 10
   'DrawWidth = 30    ' Use wider brush.
 '  ForeColor = RGB(0, 255, 0) 'RGB(kw, hw, w)   ' Set drawing color.

End Sub
