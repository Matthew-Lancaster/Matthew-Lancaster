VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Pattern"
   ClientHeight    =   4176
   ClientLeft      =   60
   ClientTop       =   732
   ClientWidth     =   5112
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   341
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4170
      Left            =   0
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   4176
      ScaleWidth      =   5112
      TabIndex        =   0
      Top             =   0
      Width           =   5115
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2160
      Top             =   7680
   End
   Begin VB.Menu Mnu_Crop 
      Caption         =   "Crop"
   End
   Begin VB.Menu Mnu_Save 
      Caption         =   "Save"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ff1
Public ff2
Public ft1
Public tt1, tt2, tt3

Public xxl1, xxh1, yyl1, yyh1

Public ra1, ra2, ra3

Public noin As Integer
Public atim As Double
Public handle As Integer
Public w As Integer
Public hw As Double
Public kw As Double
Public tw1 As Double
Public tw2 As Double
Public tw3 As Double
Public rr4 As Double
Public x3, y3, x8
Public x4 As Single
Public y4 As Single
Public x5 As Single
Public y5 As Single
Public pi As Double
Public tr

Private Sub Command1_Click()
End Sub




Private Sub Form_Load()

'handle = Shell("c:\vbdos\vbdos /ah /RUN c:\coolcid\cibas /Es /S:0 ", vbNormalFocus)

pi = 4 * Atn(1)
tw1 = 2
tw2 = 2.4
tw3 = 2.8
Form1.Height = 11000
Form1.Width = 11000
Form1.Top = 200
Form1.Left = 5000

'ff1 = 10000 / 315
'ff2 = 9650 / 315


'Picture1.Align = 1
'Picture1.Align = 2
'Picture1.Align = 3
'Picture1.Align = 4
'Picture1.Align = 0
Picture1.Width = (Form1.Width / Screen.TwipsPerPixelX) - 15

'Picture1.ScaleWidth = 100
'Picture1.ScaleHeight = 100
Picture1.Width = (Form1.Width / Screen.TwipsPerPixelX) - 15
Picture1.Height = (Form1.Height / Screen.TwipsPerPixelY) - 60
Picture1.Top = 0
Picture1.Left = 0
Picture1.DrawWidth = 70
Picture1.BackColor = RGB(255, 255, 255)
'Picture1.BackColor = RGB(0, 255, 0)
'Picture1.BackColor = RGB(0, 0, 0)
'Picture1.ScaleHeight = Form1.Height
'Picture1.ScaleLeft = 0
'Picture1.ScaleTop = 0
'Picture1.ScaleWidth = Form1.Width

x3 = (Picture1.Width * Screen.TwipsPerPixelX) / 2 '* Screen.TwipsPerPixelX '/ ff1 ' 315 'left 31.7
y3 = (Picture1.Height * Screen.TwipsPerPixelX) / 2 ' * Screen.TwipsPerPixelY '/ ff2 ' 315 'up

x3 = x3 - 20
y3 = y3 - 20
x8 = x3 * 1.016 'radius
xxl1 = Picture1.Width * Screen.TwipsPerPixelX
yyl1 = Picture1.Height * Screen.TwipsPerPixelY




End Sub

Timer1.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
      'PSet (x, y)      ' Draw a point.
  'x3 = x: y3 = y
 '  DrawWidth = 90      ' Use wider brush.
   
   
'   ForeColor = RGB(kw, hw, w)   ' Set drawing color.
End Sub



Private Sub Form_Unload(Cancel As Integer)
'Call Form_Terminate
End Sub

Private Sub Mnu_Crop_Click()

'Picture1.Top = xxl1
'Picture1.Left = yyl1
Picture1.Top = 0
Picture1.Left = 0

dw = Picture1.DrawWidth + 9
rw1 = (yyh1 - yyl1) / Screen.TwipsPerPixelY
rw2 = (xxh1 - xxl1) / Screen.TwipsPerPixelX
rw1 = rw1 + dw
rw2 = rw2 + dw
Picture1.Height = rw1 + 10 '(((yyh1 - yyl1) + Picture1.DrawWidth + 9)) - 150
Picture1.Width = rw2 + 30 '(((xxh1 - xxl1) + Picture1.DrawWidth + 9)) - 150
'Picture1.Align = 0

End Sub

Private Sub Mnu_Save_Click()
Timer1.Enabled = False

'Clipboard.Clear
'Clipboard.SetData (Picture1.Image  PictureBox)
'Picture1.Picture SavePicture(App.Path + "kk.bmp")
'SavePicture Picture1.Image, App.Path + "\YinYang3 BlkBG.bmp"
SavePicture Picture1.Image, App.Path + "\Yin-Yanger " + Format(Now, "YYYY-MM-DD HH-MM-SS") + ".bmp"

Shell "explorer " + App.Path
End
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Timer1.Enabled = False

'Clipboard.Clear
'Clipboard.SetData (Picture1.Image  PictureBox)
'Picture1.Picture SavePicture(App.Path + "kk.bmp")
'SavePicture Picture1.Image, App.Path + "\YinYang1.bmp"
'Shell "explorer " + App.Path

End Sub

Private Sub Timer1_Timer()

h1 = -10: he1 = Abs(h1)
h2 = -20: he2 = Abs(h2)
h3 = -30: he3 = Abs(h3)

tg = 255
hg12 = 0
w = w + tw1
If w >= tg - he1 Then tw1 = h1
If w < he1 + hg12 Then tw1 = he1
hw = hw + tw2
If hw >= tg - he2 Then tw2 = h2
If hw < he2 + hg12 Then tw2 = he2
kw = kw + tw3
If kw >= tg - he3 Then tw3 = h2
If kw < he3 + hg12 Then tw3 = he3


tt1 = 0.05
tt2 = 0.08
tt3 = 0.12
ra1 = ra1 + tt1
ra2 = ra2 + tt2
ra3 = ra3 + tt3

rr1 = 127 * Sin(ra1) + 127
rr2 = 127 * Sin(ra2) + 127
rr3 = 127 * Cos(ra3) + 127
Picture1.ForeColor = RGB(Int(rr1), Int(rr2), Int(rr3)) 'w   ' Set drawing color.

rr4 = rr4 + 0.005 * 2
tr = tr + 0.01 * 3



lw = (x8 / 1.2) * Sin(tr)
y4 = lw * Sin(rr4) + x3
x4 = lw * Cos(rr4) + y3
'x5 = 200 * Sin(rr3 - 3) + x3
'y5 = 200 * Cos(rr3 - 3) + y3

'x3 = Form1.Height / 2 '/ ff1 ' 315 'left 31.7
'y3 = Form1.Width / 2 '/ ff2 ' 315 'up

Picture1.PSet (x4, y4)  ' Draw a point.
  
If x4 < xxl1 Then xxl1 = x4
If x4 > xxh1 Then xxh1 = x4
If y4 < yyl1 Then yyl1 = y4
If y4 > yyh1 Then yyh1 = y4
  
  
  
  'Line (x4, y4)-(x5, y5)
   'Circle (x4, y4), 10
 '  ForeColor = RGB(0, 255, 0) 'RGB(kw, hw, w)   ' Set drawing color.

End Sub
