VERSION 5.00
Begin VB.Form F1 
   AutoRedraw      =   -1  'True
   Caption         =   "Pattern"
   ClientHeight    =   6936
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   15240
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   578
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer TIMER_GO 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2352
      Top             =   936
   End
   Begin VB.Timer HighRes 
      Enabled         =   0   'False
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
Attribute VB_Name = "F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Visual Basic For Dos Easy Converted
'Displays an animated tunnel fly through
'With kelidoscope of colours
'By Matt Lan 11-Apr-2004
Dim SET_REDRAW_LABEL_TIMER_MOD_2
Dim I(20)
Dim FS_FONT_S(20)
Dim XI
Dim LABEL8_TIMER_BEGIN_VAR
Dim LABEL8_TIMER_CALIBRATE_XI
Dim X_TIMER
Dim OXI
Dim SET_REDRAW_LABEL8_TIMER_MOD_2
Dim XI_TEXT

Public awat, pi, lk, cv, tyu, cb, c, c2, h
Public wk, hw, kw, Size, x3, y3

Private Sub Form_GotFocus()
'Timer1.Enabled = True
HighRes.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then End
End Sub

Private Sub Form_Load()

If App.PrevInstance = True Then End

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
'Picture1.BackColor = Me.BackColor
Me.Height = 7000
Me.Width = 6000
x3 = (Screen.Width / Screen.TwipsPerPixelX) / 2
y3 = (Screen.Height / Screen.TwipsPerPixelY) / 2
y3 = y3 - 40

DW = 50
DrawWidth = DW    ' Use wider brush.
'Picture1.DrawWidth = Me.DrawWidth
Size = 320
Size = Size - (DW / 2.6)

TIMER_GO.Enabled = False
TIMER_GO.Interval = 5000
TIMER_GO.Enabled = True



Exit Sub

Steps1 = 0.01: Steps2 = 0.03 'tri
Steps1 = 0.02: Steps2 = 0.09 'tri
Dot_Or_Line = 1 'line
'Dot_Or_Line = 2'dot

End Sub

Sub TIMER_GO_Timer()

Me.WindowState = vbMaximized

TIMER_GO.Enabled = False

'Picture1.Top = 0
'Picture1.Height = Me.Height
'Picture1.Width = Me.Width
'Picture1.Left = 0
'Picture1.ScaleMode = Me.ScaleMode
'Picture1.Visible = False

FS_FONT_S(1) = 35
FS_FONT_S(2) = 32
FS_FONT_S(3) = 29
FS_FONT_S(4) = 29
FS_FONT_S(5) = 29
FS_FONT_S(6) = 22
FS_FONT_S(7) = 24
I(1) = "1.. 2020-04-28--09-57-00__SCREENCASTIFY_.MP4"
I(2) = "2.. 2020_04_25 Sat_Apr 17_02_20__MAH02571__ OLD STEINE BRIGHTON_.MP4"
I(3) = "3-3.. 2020_04_27 Mon_Apr 18_54_31__MAH02629__ LAGOON SEAFRONT.MP4"
I(4) = "3-4.. 2020_04_27 Mon_Apr 19_10_12__MAH02631__ LAGOON SEAFRONT.MP4"
I(5) = "3-5.. 2020_04_27 Mon_Apr 19_16_50__MAH02634__ LAGOON SEAFRONT.MP4"
I(6) = "4-SIXER..  2020_04_28 Tue_Apr 11_31_41__MAH02635__ GARDEN WHEN SCREENCASTIFY.MP4"
I(7) = "3-7.. 2020_04_28 Tue_Apr 11_34_56__MAH02637__ GARDEN WHEN SCREENCASTIFY.MP4"
I(8) = "____ _SEVEN_ _7_ _VIDEO_ _25TH_ _TO_ _28_ _OF_ _APRIL_ ____" + vbCrLf

F2.Label3.Caption = "1.. SCREENCASTIFY"
F2.Label4.Caption = "2.. OLD STEINE BRIGHTON"
F2.Label5.Caption = "3.. LAGOON SEAFRONT 3-5"
F2.Label7.Caption = "4.. GARDEN            SIXER-7"

For R1 = 1 To 7
    LINETEXT = I(R1)
    LINETEXT = Replace(LINETEXT, "_.MP4", "")
    LINETEXT = Replace(LINETEXT, ".MP4", "")
    LINETEXT = Replace(LINETEXT, "-", "_")
    LINETEXT = Replace(LINETEXT, "3_3", "3-3")
    LINETEXT = Replace(LINETEXT, "3_4", "3-4")
    LINETEXT = Replace(LINETEXT, "3_5", "3-5")
    LINETEXT = Replace(LINETEXT, "4_S", "4-S")
    LINETEXT = Replace(LINETEXT, "3_7", "3-7")
    I(R1) = LINETEXT
Next

F2.Label1.Caption = I(8)
F2.Label1.Font = "ARIAL BLACK"
F2.Label1.FontSize = 32
F2.Label1.AutoSize = True
F2.Label1.Refresh
F2.Label1.AutoSize = False
F2.Label1.Height = F2.Label1.Height - 27
F2.Label1.Top = -12
F2.Label1.Left = 0
F2.Label1.Width = (Me.Width / Screen.TwipsPerPixelX)
F2.Label1.ForeColor = RGB(255, 255, 255)
F2.Label1.Alignment = 2
F2.Label1.BackStyle = 0

F2.Label2.Caption = I(1)
F2.Label2.Font = F2.Label1.Font
F2.Label2.FontSize = FS_FONT_S(1)
F2.Label2.AutoSize = True
F2.Label2.Refresh
F2.Label2.AutoSize = False
'F2.Label2.Height = F2.Label2.Height
F2.Label2.Top = 48
F2.Label2.Left = F2.Label1.Left
F2.Label2.Width = F2.Label1.Width
F2.Label2.ForeColor = F2.Label1.ForeColor
F2.Label2.Alignment = F2.Label1.Alignment
F2.Label2.BackStyle = 0
' F2.Label2.BackColor
' F2.Label1.BackStyle = 1
' F2.Label1.Visible = False

' --------------------
' F2.Label LEFT BOX CODE

F2.Label3.Font = F2.Label1.Font
F2.Label3.FontSize = 18
F2.Label3.AutoSize = True
F2.Label3.Refresh
F2.Label3.AutoSize = False
F2.Label3.Top = 170
F2.Label3.Left = 10
F2.Label3.Width = 800
'F2.Label3.Height = 800
F2.Label3.ForeColor = F2.Label1.ForeColor
F2.Label3.Alignment = 0 ' LEFT
F2.Label3.Refresh

HH = F2.Label3.Height - 50

F2.Label4.Font = F2.Label3.Font
F2.Label4.FontSize = F2.Label3.FontSize
F2.Label4.Top = F2.Label3.Top + HH
F2.Label4.Left = F2.Label3.Left
F2.Label4.Width = F2.Label3.Width
F2.Label4.Height = F2.Label3.Height
F2.Label4.ForeColor = F2.Label3.ForeColor
F2.Label4.Alignment = 0 ' LEFT

F2.Label5.Font = F2.Label3.Font
F2.Label5.FontSize = F2.Label3.FontSize
F2.Label5.Top = F2.Label4.Top + HH
F2.Label5.Left = F2.Label3.Left
F2.Label5.Width = F2.Label3.Width
F2.Label5.Height = F2.Label3.Height
F2.Label5.ForeColor = F2.Label3.ForeColor
F2.Label5.Alignment = 0 ' LEFT

F2.Label7.Font = F2.Label3.Font
F2.Label7.FontSize = F2.Label3.FontSize
F2.Label7.Top = F2.Label5.Top + HH
F2.Label7.AutoSize = True
F2.Label7.Refresh
F2.Label7.AutoSize = False
F2.Label7.Left = F2.Label3.Left
F2.Label7.Width = F2.Label3.Width
' F2.Label7.Height = F2.Label3.Height
F2.Label7.ForeColor = F2.Label3.ForeColor
F2.Label7.Alignment = 0 ' LEFT


F2.Label8.Font = F2.Label3.Font
F2.Label8.FontSize = 15
F2.Label8.Top = F2.Label2.Top + HH + HH + HH
F2.Label8.AutoSize = True
F2.Label8.Refresh
F2.Label8.AutoSize = False
F2.Label8.Left = 1680 ' F2.Label3.Left ' 1700
F2.Label8.Width = 200
F2.Label8.Height = F2.Label8.Height * 7
F2.Label8.ForeColor = F2.Label3.ForeColor
F2.Label8.Alignment = 0 ' LEFT

'F2.Label1.Visible = False
'F2.Label2.Visible = False
'F2.Label3.Visible = False
'F2.Label4.Visible = False
'F2.Label5.Visible = False
'F2.Label7.Visible = False
'F2.Label8.Visible = False

Me.BackColor = RGB(1, 1, 50)
'Picture1.BackColor = Me.BackColor

X_TIMER = Timer
LABEL8_TIMER_BEGIN_VAR = Timer

HighRes.Enabled = True
HighRes.Interval = 1

F2.Show

End Sub

Private Sub Form_LostFocus()
    Timer1.Enabled = False
    HighRes.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    'PSet (x, y)      ' Draw a point.
    'x3 = x: y3 = y
    'DrawWidth = 90      ' Use wider brush.
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

' ForeColor = RGB(Int(kw), Int(hw), Int(wk))   ' Set drawing color.
Picture1.ForeColor = RGB(Int(kw), Int(hw), Int(wk))    ' Set drawing color.

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
    R = tyu
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
        
        R = R + q
        w = ((R / (pi)) / ((h)))  '(h * 2))'h * 1.5))
        X = (w * 1.7) * Sin(R * 5) + x3
        Y = (w) * Cos(R * 5) + y3

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

        F1.ForeColor = RGB(Int(kw), Int(hw), Int(wk))    ' Set drawing color.
        ' Picture1.ForeColor = RGB(Int(kw), Int(hw), Int(wk))   ' Set drawing color.
        F1.PSet (X, Y)
        ' Picture1.PSet (x, y)
        
        Rem CIRCLE (x, y), 2, c1
        Rem LINE STEP(0, 0)-(x, y), c1
        s1 = 0
        Rem wl = 100
        Rem IF x < wl OR y < wl THEN s1 = 1
        Rem IF x > 640 - wl OR y > 480 - wl THEN s1 = 1
        Rem IF x > 590 OR y > 430 THEN s1 = 1
        
        If XI = 0 Then XI = 2
        Dim XI_VAR(10)
        XI_VAR(2) = 15#
        XI_VAR(3) = 10.3
        XI_VAR(4) = 10.4
        XI_VAR(5) = 10.5
        XI_VAR(6) = 10.51222
        XI_VAR(7) = 10.00007
        XI_VAR(8) = 10.8
        
'        XI_VAR(2) = 2
        XI_VAR(3) = 40
        XI_VAR(4) = 40
        XI_VAR(5) = 40
        XI_VAR(6) = 40
        XI_VAR(7) = 40
        XI_VAR(8) = 40
        
        
        If XI < 7 Then XI_TEXT = " >----"
        If XI > 7 And XI_TEXT = " >----" Then
            XI_TEXT = " >++<"
        End If
        
        If XI > 7 Then
            XI = 7
        End If
        
        
        If XI <> OXI Then
            LABEL8_TIMER_CALIBRATE_XI = LABEL8_TIMER_CALIBRATE
        End If
        OXI = XI
        
        LABEL8_TIMER_CALIBRATE = Timer - LABEL8_TIMER_BEGIN_VAR
        If Timer * 4 Mod 2 = 0 Then
            If SET_REDRAW_LABEL8_TIMER_MOD_2 = False Then
                OLD_TIMER_1 = Timer
                VAR_MINUS = Abs(XI_VAR(XI) - Abs(LABEL8_TIMER_CALIBRATE - (XI_VAR(XI) + LABEL8_TIMER_CALIBRATE_XI)))
                VAR_2 = Abs(LABEL8_TIMER_CALIBRATE - (XI_VAR(XI) + LABEL8_TIMER_CALIBRATE_XI))
                F2.Label8.Caption = Format(LABEL8_TIMER_CALIBRATE, "0.00") + vbCr + Format(VAR_2, "0.00") + XI_TEXT + vbCr + Format(VAR_MINUS, "0.00") + vbCr + Trim(Str(XI_VAR(XI))) + vbCr + "--" + vbCr + Format(LABEL8_TIMER_CALIBRATE_XI + XI_VAR(XI), "0.00") + " --" + Str(XI) + ".."
                SET_REDRAW_LABEL8_TIMER_MOD_2 = True
            End If
        Else
            SET_REDRAW_LABEL8_TIMER_MOD_2 = False
        End If
        
        If Timer * 2 Mod 2 = 0 Then
            If SET_REDRAW_LABEL_TIMER_MOD_2 = False Then
                X2_TIMER = 4
                If XI < 8 Then
                    If LABEL8_TIMER_CALIBRATE > (XI_VAR(XI) + LABEL8_TIMER_CALIBRATE_XI) Then
                        F2.Label2.Caption = I(XI): X_TIMER = Timer
                        F2.Label2.FontSize = FS_FONT_S(XI)

                        XI = XI + 1
                    End If
                End If
                
                SET_REDRAW_LABEL_TIMER_MOD_2 = True
            End If
        Else
            SET_REDRAW_LABEL_TIMER_MOD_2 = False
        End If

        If w > Size Then s = s + 1 Else s = 0

    Loop Until s > 100
    s = 0

End Sub

