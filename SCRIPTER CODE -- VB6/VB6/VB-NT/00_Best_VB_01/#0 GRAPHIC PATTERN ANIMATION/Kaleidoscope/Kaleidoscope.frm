VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Keleidoscope 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Pattern"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   760
   WindowState     =   1  'Minimized
   Begin MSComctlLib.Slider Slider1 
      Height          =   10695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   18865
      _Version        =   393216
      Orientation     =   1
      Max             =   100
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1800
      Top             =   2160
   End
   Begin VB.Timer HighRes 
      Interval        =   1
      Left            =   1755
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   1080
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   10695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   18865
      _Version        =   393216
      Orientation     =   1
      Max             =   100
   End
   Begin VB.Menu MNU_VBME 
      Caption         =   "VB ME"
      NegotiatePosition=   3  'Right
   End
   Begin VB.Menu MNU_CHKBOX 
      Caption         =   "CHECK BOX"
      Begin VB.Menu MNU_REVERSE 
         Caption         =   "REVERSE"
      End
   End
   Begin VB.Menu MNU_END 
      Caption         =   "EXIT"
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



Public awat, pi, lk, cv, tyu, cb, c, c2, h, CSINSLIDER As Double
Public wk, hw, kw, Size, x3, y3, TEN




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If IsIDE = True And KeyCode = 27 Then End
End Sub

Private Sub Form_Load()

If App.PrevInstance = True Then End



pi = 4 * Atn(1)
lk = 150




pi = 4 * Atn(1)
twd1 = 4
twd2 = 4.1
twd3 = 4.2
twd3 = 4.5
BackColor = RGB(255, 255, 255)
BackColor = RGB(0, 0, 0)
Me.Height = 7000
Me.Width = 6000
x3 = (Screen.Width / Screen.TwipsPerPixelX) / 2
y3 = (Screen.Height / Screen.TwipsPerPixelY) / 2
y3 = y3 - 40

DW = 30 ' -- 10 20 30 40 50 60 70 80 90 100
DW = 40
DrawWidth = DW    ' Use wider brush.
Slider1.Value = DW



Size = 320
Size = Size - (DW / 2.6)
'XT = -5
Me.Show
DoEvents
Me.Show
'Exit Sub

Steps1 = 0.01: Steps2 = 0.03 'tri
Steps1 = 0.02: Steps2 = 0.09 'tri
Steps1 = 0.0001: Steps2 = 0.03 'tri
Steps1 = 0.01: Steps2 = 0.09 'tri
Steps1 = 0.11: Steps2 = 0.19 'tri
Steps1 = 0.0001: Steps2 = 0.4009 'tri
Dot_Or_Line = 1 'line
Dot_Or_Line = 2 'dot

If IsIDE = True Then
    Me.WindowState = vbNormal
    DoEvents
    Me.Show
End If


End Sub


Private Sub Form_LostFocus()

Me.WindowState = vbMinimized

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
      'PSet (x, y)      ' Draw a point.
  'x3 = x: y3 = y
 '  DrawWidth = 90      ' Use wider brush.
'   ForeColor = RGB(kw, hw, w)   ' Set drawing color.
End Sub




Private Sub Form_Resize()
'Exit Sub
If Me.WindowState = vbNormal Then
'    Me.Show
'    DoEvents
    Me.Height = Screen.Height - 450 '7000 ' 7755   '7000
    Me.Width = Screen.Width '11520    '6000
    Me.Top = 500
    Slider1.Height = (Me.Height - 1800) / Screen.TwipsPerPixelY
    Slider2.Height = (Me.Height - 1800) / Screen.TwipsPerPixelY
    Slider2.Left = Slider1.Width
End If
If Me.WindowState Mod 2 = 0 Then Call TIMEROUT(True): Exit Sub
If Me.WindowState = 1 Then Call TIMEROUT(False): Exit Sub


End Sub

Sub TIMEROUT(TF)
    On Error Resume Next
    For Each Form In Forms
        For Each Control In Controls
            If Control.Interval > 0 Then Control.Enabled = TF
        Next
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Me.WindowState = 1 Then Exit Sub

Cancel = True
Me.WindowState = 1
End Sub

Private Sub MNU_END_Click()
End
End Sub

Private Sub MNU_REVERSE_Click()
MNU_REVERSE.Checked = Not MNU_REVERSE.Checked
End Sub

Private Sub MNU_VBME_Click()
    Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
    End
End Sub


Sub SET_DRAW_COLOR()
'Do
c = c + 0.05
c = c + (Slider2.Value / 0.0005)
kw = 127 * Sin(c) + 127
hw = 127 * Sin(c * 1.1) + 127
wk = 127 * Sin(c * 1.2) + 127


If hw < 4 Or hw > 251 And 1 = 2 Then

Else
    If kw > 200 Then
        kw = 254 - kw
    End If
    If wk > 200 Then
        wk = 254 - wk
    End If
End If
'Loop Until EXPLOD = 0

ForeColor = RGB(Int(kw), Int(hw), Int(wk))   ' Set drawing color.

End Sub

Private Sub Slider1_Scroll()
DW = Slider1.Value
If DW = 0 Then DW = 1
Me.DrawWidth = DW
End Sub

Private Sub Slider2_Click()
CSINSLIDER = Slider2.Value
End Sub

Private Sub Timer1_Timer()

SET_DRAW_COLOR


rr3 = rr3 + Steps1
tr = tr + Steps2
lw = Size * Cos(tr)
x4 = (lw) * Sin(rr3) + x3
y4 = (lw) * Cos(rr3) + y3
'x5 = Size * Sin(rr3 - 3) + x3
'y5 = Size * Cos(rr3 - 3) + y3
'
If Dot_Or_Line = 1 Then Line Step(0, 0)-(x4, y4)
If Dot_Or_Line = 2 Then PSet (x4, y4)    ' Draw a point.
'Line Step(0, 0)-(x4, y4)
'PSet (x4, y4)    ' Draw a point.
   
   'Circle (x4, y4), 10
   'DrawWidth = 30    ' Use wider brush.
 '  ForeColor = RGB(0, 255, 0) 'RGB(kw, hw, w)   ' Set drawing color.

End Sub




Sub HighRes_timer()

cv = 0.0000001
cv = 0.00002
cv = 0.00009
cv = 0.00019
Rem FOR h = 0 TO lk STEP cv'-1* 2 + 1

 'h = h + (cv * (MNU_REVERSE.Checked * 2 + 1))
 h = h - cv
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

'        c = c + 0.05
'        kw = 200 * Sin(c) + 220
'        hw = 200 * Sin(c * 1.1) + 220
'        wk = 200 * Sin(c * 1.2) + 220
        SET_DRAW_COLOR
        
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

Private Sub Timer2_Timer()
'Exit Sub

' --- HALF HOUR
' --- 60 SEC
' --- MULTIPLY 10

' --- MULTIPLY 3
' --- HALF HOUR


If TEN >= 600 * 3 Then Me.WindowState = vbMinimized: Timer2.Enabled = False: TEN = -4
If TEN > 0 Then TEN = TEN + 1


End Sub
'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************

