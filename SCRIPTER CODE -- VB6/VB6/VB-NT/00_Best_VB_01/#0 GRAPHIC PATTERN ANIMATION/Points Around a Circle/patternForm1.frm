VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public w As Double
Public hw As Double
Public kw As Double
Public tw1 As Double
Public tw2 As Double
Public tw3 As Double

Public v As Double
Dim vert%(2000)


Private Sub Form_Load()

'Dim vert%(2000)
'Dim v As Double
'KEY OFF
'CLS
'INPUT "enter radius (range 10-170)"; r
'INPUT "enter number of vertices (range 3-50)"; v
'CLS
'v = 1

End Sub

Private Sub Timer1_Timer()

DrawWidth = 1 ' Use wider brush.
r = 370 'Size
If v = 0 Then v = 1

'For v = 1 To 500 Step 5
v = v + 0.01
If v > 40 Then v = v + 0.1
If v > 80 Then v = v + 1
If v > 160 Then v = v + 2
If v > 320 Then v = v + 4
If v > 1000 Then v = 1



'Form1.Height/2
x0 = (Screen.Width / Screen.TwipsPerPixelX) / 2
y0 = (Screen.Height / Screen.TwipsPerPixelY) / 2
x0 = x0 + 20
y0 = y0 - 50
'PSet (x0, y0) 'Centre

THETA = 6.28319 / v: N = v - 1


'THETA = 5 / v: N = v - 1
For I = 1 To N
vert%(2 * I - 1) = r * 1.5 * Cos(THETA * I) + x0
vert%(2 * I) = r * Sin(THETA * I) + y0
PSet (vert%(2 * I - 1) - 40, 50 + vert%(2 * I)), 1
Next I
N0 = N
For J = 1 To N



w = w + tw1
If w > 255 Then tw1 = -6: w = w + tw1
If w < 1 Then tw1 = 6: w = w + tw1
hw = hw + tw2
If hw > 255 Then tw2 = -7: hw = hw + tw2
If hw < 1 Then tw2 = 7: hw = hw + tw2
kw = kw + tw3
If kw > 255 Then tw3 = -8: kw = kw + tw3
If kw < 1 Then tw3 = 8: kw = kw + tw3
   
ForeColor = RGB(kw, hw, w)   ' Set drawing color.





For I = 1 To N0





XN = vert%(2 * I - 1): YN = vert%(2 * I)
Line (x0 - 40, 50 + y0)-(XN - 40, 50 + YN)
Next I
x0 = XN: y0 = YN: N0 = N0 - 1
Next J
'Next

'LOCATE 1, 1
'INPUT "Try another (Y/N)"; A$
'IF A$ = "Y" OR A$ = "y" GOTO 40
'CLS
'END

End Sub
