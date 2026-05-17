VERSION 5.00
Begin VB.Form Keleidoscope 
   Caption         =   "Pattern"
   ClientHeight    =   3156
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   4188
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   263
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer HighRes 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1755
      Top             =   1170
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

Option Explicit


Dim Awat, pi, LK, CV, TYU, cb, C, H, Rg5, Modi, BackColor4, Q, TT, GG, i
Dim WK, HW, KW, Size, X3, Y3, Gh, Cdr, Cat, Cbz
Dim HF, W, x, y, Tag2, TTI, Steps1, CsD1, CsD2, CsD3, S1, S
Dim TwD1 As Double
Dim TwD2 As Double
Dim TwD3 As Double
Dim TwT1
            
Dim TwX1
Dim TwX2
Dim TwX3
Dim TTI2, CoS1, CoS2, CoS3

            

Private Sub Form_Load()

'frmPassLock.hide

pi = 4 * Atn(1)
LK = 150
'Cv = 0.000001
CV = 0.0000005

H = 0.5

pi = 4 * Atn(1)
TwD1 = 2
TwD2 = 2.1
TwD3 = 2.2
BackColor4 = RGB(255, 255, 255)
BackColor4 = RGB(0, 0, 0)
Me.Height = 7000


Me.Width = 6000
X3 = (Screen.Width / Screen.TwipsPerPixelX) / 2
Y3 = (Screen.Height / Screen.TwipsPerPixelY) / 2
Y3 = Y3 - 20

dw = 20

'frmPassLock.DrawWidth = dw    ' Use wider brush.




Size = 330
Size = Size - (dw / 2.6)

TT = 0
GG = 0

Exit Sub

Steps1 = 0.01: Steps2 = 0.03 'tri
Steps1 = 0.02: Steps2 = 0.09 'tri
Dot_Or_Line = 1 'line
'Dot_Or_Line = 2'dot

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
      'PSet (x, y)      ' Draw a point.
  'x3 = x: y3 = y
 '  DrawWidth = 90      ' Use wider brush.
   'ForeColor = RGB(kw, hw, W)   ' Set drawing color.
End Sub





Sub HighRes_Timer()

    'If Me.Visible = False Then HighRes.Enabled = False

    H = H + CV + (frmPassLock.Slider3.Value / 900000) 'Spin Speed
 
    Q = H
    
    TYU = TYU + 0.0003
    Rg5 = TYU
    
    'wk = 10
    'hw = 10
    'kw = 10


    Cbz = Cbz + 0.0000009
            
    If frmPassLock.Check6.Value = vbChecked Then 'seedless checked
        TwD1 = cb
        cb = cb + 0.00000005 + (frmPassLock.Slider1.Value / 1000) 'colour speed
    Else
        cb = cb + 0.00000005 + (frmPassLock.Slider1.Value / 5000) 'colour speed
    End If
    If Err.Number > 0 Then cb = 0: Err.Clear
    If frmPassLock.Check1.Value = vbUnchecked Then 'seedless checked
'        cb2 = cb '(((frmPassLock.Slider4.Value) / 500000) + (cb)) 'seed colour
        TwD1 = cb
        Cdr = cb
    Else
'         Cdr = cb
    End If
    'c2 = c2 + cb2
'    frmPassLock.PSet (x3, y3)
    HF = True
    On Local Error Resume Next
    Do
        
        Modi = Modi + 1
        If Modi = 300 Then
            Modi = 0 ':      DoEvents
'            Call ATidalX.Timer7WinAmpVolIDE_Timer
'            Call ATidalX.Timer8WinAmpVol_Timer
        End If
        
        
        
        
        Rem c2 = c2 + (cv)
        Rem gh = 300 * (SIN(c2)) + 100

        'c = c + 0.001: Rem (cv * 1000) + ((COS(c2)) / gh)
        'If c > 15 Then c = 0
        'c1 = 14 * (Abs(Cos(c)))
        
        Rg5 = Rg5 + Q
        'If Err.Number > 0 Then Rg5 = 0: Err.Clear
        
        W = ((Rg5 / 5) / ((H)))  '(h * 2))'h * 1.5))
        x = (W * 1.7) * Sin(Rg5) + X3
        y = (W) * Cos(Rg5) + Y3

        
        'If frmPassLock.Check5.Value = vbUnchecked Then 'exited keli
            'sine keli
            If frmPassLock.Check1.Value = vbUnchecked Then Tag2 = 0.005
            If frmPassLock.Check1.Value = vbChecked Then Tag2 = 0.05
        'End If
        
        
        If frmPassLock.Check5.Value = vbChecked Then
            Cat = Cat + 0.00005 '+ frmPassLock.Slider4.Value / 500000
            TwD1 = TwD1 + (Sin(Cat) / 90)
        End If
        
        

        
        '1=Colour speed - Zoom in out
        '2=draw Width
        '3=Spin Speed
        '4=Seed colur
        '5=Colour Sine
        
        'If frmPassLock.Check1.Value = vbChecked Then tti = tti + (Cdr * 10)

        'tti = tti/10
        
        'tth = (frmPassLock.Slider5.Value / 5000)
        
        'ttl = frmPassLock.Slider5.Value / 5000
    
        
        'frmPassLock.Slider2.Enabled = False
        'frmPassLock.Slider4.Enabled = False
        
        'tti = tti * 10
        'tti = tti
        
        If frmPassLock.Check6.Value = vbChecked Then
        
        TwD1 = TwD1 + Tag2
        TTI = TwD1 + Cdr
        CsD1 = (frmPassLock.Slider5.Value / 1000)
        CsD2 = ((CsD1 * 3))
        CsD3 = ((CsD1 * 5))
        KW = 120 * Cos(TTI) + 120
        HW = 120 * Sin((TTI * (CsD2))) + 120
        WK = 120 * Cos((TTI * (CsD3))) + 120
        Else
        
        TwT1 = frmPassLock.Slider5.Value / 50000
           
        TTI2 = TTI2 + 0.0001
           
        CoS1 = Abs(Cos(TTI2))
        CoS2 = Abs(Sin(TTI2 * 5))
        CoS3 = Abs(Cos(TTI2 * 9))
        
        TwD1 = TwT1 + CoS1
        If HW > 255 Then
            TwX1 = (TwD1 * -1)
        End If
        If HW < 1 Or TwX1 = 0 Then TwX1 = (TwD1)
        HW = HW + TwX1
            
        
        '10^1.2
        TwD2 = TwT1 + (CoS2 * 2)
        If HW > 255 Then
            TwX2 = (TwD2 * -1)
        End If
        If HW < 1 Or TwX2 = 0 Then TwX2 = (TwD2)
        HW = HW + TwX2
            
        
        TwD3 = TwT1 + (CoS3 * 3)
        If KW > 255 Then TwX3 = (TwD3 * -1)
        If KW < 1 Or TwX3 = 0 Then TwX3 = (TwD3)
        KW = KW + TwX3
        
        End If
 
 '       hw = 125 * Sin(tti * ((frmPassLock.Slider5.Value / 100))) + 125
  '      wk = 125 * Sin(tti * ((frmPassLock.Slider5.Value / 50))) + 125


        frmPassLock.ForeColor = RGB(Int(KW), Int(HW), Int(WK))
                
        frmPassLock.DrawWidth = frmPassLock.Slider2.Value
        
        If XPud = False Then
        
        XPud = True: frmPassLock.PSet (x, y)
        End If
        If frmPassLock.Check2.Value = vbChecked Then
            frmPassLock.PSet (x, y)
        End If
        
        If frmPassLock.Check3.Value = vbChecked Then
            frmPassLock.DrawWidth = 2
            frmPassLock.Circle (x, y), frmPassLock.Slider2.Value
        End If
        
        If frmPassLock.Check4.Value = vbChecked Then
            If HF = True Then
                frmPassLock.PSet (x, y)
                HF = False
            End If
            frmPassLock.Line Step(0, 0)-(x, y)
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
                    ptBez1(i + 1).y = ptBez1(i).y
                End If
            Next
            ptBez1(0).x = x
            ptBez1(0).y = y
            Call PolyBezier(frmPassLock.HDC, ptBez1(0), UBound(ptBez1) - 1)

End If

        
        
        
        
        
        
        
        S1 = 0
        Rem wl = 100
        Rem IF x < wl OR y < wl THEN s1 = 1
        Rem IF x > 640 - wl OR y > 480 - wl THEN s1 = 1
        Rem IF x > 590 OR y > 430 THEN s1 = 1

        If W > Size Then S = S + 1 Else S = 0

    Loop Until S > 100
    S = 0

End Sub

