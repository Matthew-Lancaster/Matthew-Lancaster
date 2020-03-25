Attribute VB_Name = "modFormulae"
Option Explicit

Public Function LargeMod(n1 As Double, n2 As Long) As Long



End Function

Public Sub CalcLines(nFrac As Long, ByVal StartLine As Long, ByVal EndLine As Long, Optional OnlyOneLine As Boolean = False)
Dim rX As Double, rY As Double, nX As Long, nY As Long, nLine As Long
Dim SX As Double, SY As Double, EX As Double, EY As Double
Dim dX As Double, dY As Double

Dim X0 As Double, Y0 As Double, X1 As Double, Y1 As Double
Dim X2 As Double, Y2 As Double
Dim XY0 As Double, XY1 As Double
Dim z0 As tComplex, z1 As tComplex
Dim zTmp1 As tComplex, zTmp2 As tComplex
Dim zTmp3 As tComplex, zTmp4 As tComplex

Dim nW As Double, nH As Double
Dim BailOut As Double, i As Double, MaxI As Long
Dim P0 As Double, P1 As Double, P2 As Double, P3 As Double, P4 As Double

'Local misc calculation vars:
'-----------------------------------------------------------------------------------
Dim Var1 As Double, Var2 As Double, Var3 As Double, Var4 As Double, Var5 As Double
Dim Var6 As Double, Var7 As Double, Var8 As Double, Var9 As Double, Var10 As Double
'-----------------------------------------------------------------------------------

'Get some xtra priority for calculation:
SetThreadPriority App.ThreadID, THREAD_BASE_PRIORITY_MAX

'Retrieve parameters:
P0 = Frac(nFrac).Param(0)
P1 = Frac(nFrac).Param(1)
P2 = Frac(nFrac).Param(2)
P3 = Frac(nFrac).Param(3)
P4 = Frac(nFrac).Param(4)

'Choose which formula to compute:
Select Case Frac(nFrac).nType

'****************************************************************************************
'******************************* M A N D E L B R O T ************************************
'****************************************************************************************
Case TYPE_MANDELBROT

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((CDbl(nX) * dX) / CDbl(nW))
        rY = SY + ((CDbl(nY) * dY) / CDbl(nH))
        i = 0
        X0 = rX + P0
        Y0 = rY + P1
        
        While i <= MaxI
            
            X1 = (X0 * X0) - (Y0 * Y0) + rX
            Y1 = (2 * X0 * Y0) + rY
            
            'XY1 = (X1 * X1) + (Y1 * Y1)
            'If XY1 > BailOut Then
            If X1 * X1 + Y1 * Y1 > BailOut Then
                GoTo BailOut_Mandel
            Else
                i = i + 1
                X0 = X1
                Y0 = Y1
            End If
        Wend
        
        '------------------
        'Inside mandelbrot set
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_Mandel
BailOut_Mandel:
        '------------------
        
        'Outside mandelbrot set
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (X0 * X0) + (Y0 * Y0)
            .nI = i
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_Mandel:
    Next nX
Next nY

'****************************************************************************************
'******************************* M A N D E L B R O T - Z POW ****************************
'****************************************************************************************
Case TYPE_MANDELZPOW

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY
If P2 < 1 Or P2 > 200 Then P2 = 1

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((CDbl(nX) * dX) / CDbl(nW))
        rY = SY + ((CDbl(nY) * dY) / CDbl(nH))
        i = 0
        z0.r = rX + P0
        z0.i = rY + P1
        
        While i <= MaxI
            
            z1 = c_pow(z0, CLng(P2))
            z1.r = z1.r + rX
            z1.i = z1.i + rY
            
            XY1 = (z1.r * z1.r) + (z1.i * z1.i)
            If XY1 > BailOut Then
                GoTo BailOut_MandelZPow
            Else
                i = i + 1
                z0.r = z1.r
                z0.i = z1.i
            End If
        Wend
        
        '------------------
        'Inside mandelbrotZPow set
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = z0.r
            .Y0 = z0.i
            .X1 = z1.r
            .Y1 = z1.i
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_MandelZPow
BailOut_MandelZPow:
        '------------------
        
        'Outside mandelbrotZPow set
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (z0.r * z0.r) + (z0.i * z0.i)
            .nI = i
            .X0 = z0.r
            .Y0 = z0.i
            .X1 = z1.r
            .Y1 = z1.i
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_MandelZPow:
    Next nX
Next nY

'****************************************************************************************
'************************************** MANDEL^3 ****************************************
'****************************************************************************************
Case TYPE_MANDEL3

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((nX * dX) / nW)
        rY = SY + ((nY * dY) / nH)
        i = 0
        X0 = rX + P0
        Y0 = rY + P1
        
        While i <= MaxI
            X1 = (X0 * X0 * X0) - (3 * X0 * Y0 * Y0) + rX
            Y1 = (3 * X0 * X0 * Y0) - (Y0 * Y0 * Y0) + rY
            
            XY1 = (X1 * X1) + (Y1 * Y1)
            If XY1 > BailOut Then
                GoTo BailOut_Mandel3
            Else
                i = i + 1
                X0 = X1
                Y0 = Y1
            End If
        Wend
        
        '------------------
        'Inside Mandel^3 set
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_Mandel3
BailOut_Mandel3:
        '------------------
        'Outside Mandel^3 set
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (X0 * X0) + (Y0 * Y0)
            .nI = i
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_Mandel3:
    Next nX
Next nY

'****************************************************************************************
'************************************** MANDEL^4 ****************************************
'****************************************************************************************
Case TYPE_MANDEL4

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((nX * dX) / nW)
        rY = SY + ((nY * dY) / nH)
        i = 0
        X0 = rX + P0
        Y0 = rY + P1
        
        While i <= MaxI
            X1 = (X0 * X0 * X0 * X0) - (6 * X0 * X0 * Y0 * Y0) + (Y0 * Y0 * Y0 * Y0) + rX
            Y1 = (4 * X0 * X0 * X0 * Y0) - (4 * X0 * Y0 * Y0 * Y0) + rY
            
            XY1 = (X1 * X1) + (Y1 * Y1)
            If XY1 > BailOut Then
                GoTo BailOut_Mandel4
            Else
                i = i + 1
                X0 = X1
                Y0 = Y1
            End If
        Wend
        
        '------------------
        'Inside Mandel^4 set
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_Mandel4
BailOut_Mandel4:
        '------------------
        'Outside Mandel^4 set
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (X0 * X0) + (Y0 * Y0)
            .nI = i
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_Mandel4:
    Next nX
Next nY


'****************************************************************************************
'************************************** J U L I A ***************************************
'****************************************************************************************
Case TYPE_JULIA

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((nX * dX) / nW)
        rY = SY + ((nY * dY) / nH)
        i = 0
        X0 = rX
        Y0 = rY
        
        While i <= MaxI
            X1 = (X0 * X0) - (Y0 * Y0) + P0
            Y1 = (2 * X0 * Y0) + P1
            XY1 = (X1 * X1) + (Y1 * Y1)
            If XY1 > BailOut Then
                GoTo BailOut_Julia
            Else
                i = i + 1
                X0 = X1
                Y0 = Y1
            End If
        Wend
        
        '------------------
        'Inside julia set
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_Julia
BailOut_Julia:
        '------------------
        'Outside julia set
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (X0 * X0) + (Y0 * Y0)
            .nI = i
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_Julia:
    Next nX
Next nY

'****************************************************************************************
'************************************** J U L I A  - Z POW ******************************
'****************************************************************************************
Case TYPE_JULIAZPOW

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY
If P2 < 1 Or P2 > 200 Then P2 = 1

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((nX * dX) / nW)
        rY = SY + ((nY * dY) / nH)
        i = 0
        z0.r = rX
        z0.i = rY
        
        While i <= MaxI
            'X1 = pow_r(X0, Y0, CLng(P2)) + P0
            'Y1 = pow_i(X0, Y0, CLng(P2)) + P1
            z1 = c_pow(z0, CLng(P2))
            z1.r = z1.r + P0
            z1.i = z1.i + P1
            
            XY1 = (z1.r * z1.r) + (z1.i * z1.i)
            If XY1 > BailOut Then
                GoTo BailOut_JuliaZPow
            Else
                i = i + 1
                z0.r = z1.r
                z0.i = z1.i
            End If
        Wend
        
        '------------------
        'Inside juliaZPow set
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = z0.r
            .Y0 = z0.i
            .X1 = z1.r
            .Y1 = z1.i
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_JuliaZPow
BailOut_JuliaZPow:
        '------------------
        'Outside juliaZPow set
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (z0.r * z0.r) + (z0.i * z0.i)
            .nI = i
            .X0 = z0.r
            .Y0 = z0.i
            .X1 = z1.r
            .Y1 = z1.i
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_JuliaZPow:
    Next nX
Next nY


'****************************************************************************************
'******************************* M A N D E L B R O T - R I N G S ************************
'****************************************************************************************
Case TYPE_MANDELRINGS

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

Var6 = P0
If Var6 <= 0 Then Var6 = 0.2

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        
        X0 = rX
        Y0 = rY
        rX = SX + ((CDbl(nX) * dX) / CDbl(nW))
        rY = SY + ((CDbl(nY) * dY) / CDbl(nH))
        Var1 = 0
        Var2 = 2

        'cMagnitude = Var2
        'zMagnitude = Var3
        'dRadius_HI = Var4
        'dRadius_LO = Var5
        'dRadius = Var6
                
        Var2 = Sqr(rX * rX + rY * rY)
        Var4 = Var2 + Var6
        Var5 = Var2 - Var6

        For i = 1 To Frac(nFrac).MaxI
            X1 = X0 * X0 - Y0 * Y0 + rX
            Y1 = 2 * X0 * Y0 + rY
            
            X0 = X1
            Y0 = Y1

            Var3 = Sqr(X1 * X1 + Y1 * Y1)
            If (Var3 < Var4 And Var3 > Var5 And i > 1) Then
                Var1 = Sqr(Var1) + (1 - Abs(Var3 - Var2) / Var6)
            End If
            If Var3 > BailOut Then
                Exit For
            End If
        Next i
        
        With Frac(nFrac).Cell(nX, nLine)
            If Var1 > 0 Then
                .nI = Sqr(Var1) * 5 '500
                .X0 = X0
                .Y0 = Y0
                .X1 = X1
                .Y1 = Y1
                .rX = rX
                .rY = rY
                .State = 0
            Else
                .nI = 0
                .X0 = X0
                .Y0 = Y0
                .X1 = X1
                .Y1 = Y1
                .rX = rX
                .rY = rY
                .State = 1
            End If
        End With
    Next nX
Next nY


'****************************************************************************************
'******************************* BarnsleyMandel 1 ***************************************
'****************************************************************************************
Case TYPE_BARNSLEYM1


With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((CDbl(nX) * dX) / CDbl(nW))
        rY = SY + ((CDbl(nY) * dY) / CDbl(nH))
        i = 0
        X0 = rX + P0
        Y0 = rY + P1
        
        While i <= MaxI
            If X0 < 0 Then
                X0 = X0 + 1
            Else
                X0 = X0 - 1
            End If

            X1 = X0 * rX - Y0 * rY
            Y1 = X0 * rY + Y0 * rX

            X0 = X1
            Y0 = Y1
            
            XY1 = (X1 * X1) + (Y1 * Y1)
            If XY1 > BailOut Then
                GoTo BailOut_BarnsleyM1
            Else
                i = i + 1
                X0 = X1
                Y0 = Y1
            End If
        Wend
        
        '------------------
        'Inside BarnsleyMandel set
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_BarnsleyM1
BailOut_BarnsleyM1:
        '------------------
        
        'Outside BarnsleyMandel set
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (X0 * X0) + (Y0 * Y0)
            .nI = i
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_BarnsleyM1:
    Next nX
Next nY

'****************************************************************************************
'******************************* BarnsleyMandel 4 ***************************************
'****************************************************************************************
Case TYPE_BARNSLEYM4

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((CDbl(nX) * dX) / CDbl(nW))
        rY = SY + ((CDbl(nY) * dY) / CDbl(nH))
        i = 0
        X0 = rX + P0
        Y0 = rY + P1
        While i <= MaxI
            
            If X0 < P0 Then X0 = X0 + P1
            If X0 > P0 Then X0 = X0 - P1
            
            X1 = (X0 * X0) - (Y0 * Y0) + rX
            Y1 = (2 * X0 * Y0) + rY
            
            XY1 = (X1 * X1) + (Y1 * Y1)
            If XY1 > BailOut Then
                GoTo BailOut_MandelBarnsley
            Else
                i = i + 1
                X0 = X1
                Y0 = Y1
            End If
        Wend
        
        '------------------
        'Inside MandelBarnsley4 set
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_MandelBarnsley
BailOut_MandelBarnsley:
        '------------------
        
        'Outside MandelBarnsley set
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (X0 * X0) + (Y0 * Y0)
            .nI = i
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_MandelBarnsley:
    Next nX
Next nY

'****************************************************************************************
'******************************* BARNSLEYJULIA 1 ****************************************
'****************************************************************************************
Case TYPE_BARNSLEYJ1

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((CDbl(nX) * dX) / CDbl(nW))
        rY = SY + ((CDbl(nY) * dY) / CDbl(nH))
        i = 0
        X0 = rX
        Y0 = rY
        
        While i <= MaxI
            If X0 >= 0 Then
                X0 = X0 - 1
                X1 = X0 * P0 - Y0 * P1
                Y1 = X0 * P1 + Y0 * P0
            Else
                X0 = X0 + 1
                X1 = X0 * P0 - Y0 * P1
                Y1 = X0 * P1 + Y0 * P0
            End If
            
            XY1 = (X1 * X1) + (Y1 * Y1)
            If XY1 > BailOut Then
                GoTo BailOut_Barnsleyj1
            Else
                i = i + 1
                X0 = X1
                Y0 = Y1
            End If
        Wend
        
        '------------------
        'Inside barnsleyj1 set
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_Barnsleyj1
BailOut_Barnsleyj1:
        '------------------
        
        'Outside barnsleyj1 set
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (X0 * X0) + (Y0 * Y0)
            .nI = i
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_Barnsleyj1:
    Next nX
Next nY
        
'****************************************************************************************
'******************************* BARNSLEYJULIA 2 ****************************************
'****************************************************************************************
Case TYPE_BARNSLEYJ2

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((CDbl(nX) * dX) / CDbl(nW))
        rY = SY + ((CDbl(nY) * dY) / CDbl(nH))
        i = 0
        X0 = rX
        Y0 = rY
        
        While i <= MaxI
            If (X0 * P1 + P0 * Y0) >= 0 Then
                X0 = X0 - 1
                X1 = X0 * P0 - Y0 * P1
                Y1 = X0 * P1 + Y0 * P0
            Else
                X0 = X0 + 1
                X1 = X0 * P0 - Y0 * P1
                Y1 = X0 * P1 + Y0 * P0
            End If
            
            XY1 = (X1 * X1) + (Y1 * Y1)
            If XY1 > BailOut Then
                GoTo BailOut_Barnsleyj2
            Else
                i = i + 1
                X0 = X1
                Y0 = Y1
            End If
        Wend
        
        '------------------
        'Inside barnsleyj2 set
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_Barnsleyj2
BailOut_Barnsleyj2:
        '------------------
        
        'Outside barnsleyj2 set
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (X0 * X0) + (Y0 * Y0)
            .nI = i
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_Barnsleyj2:
    Next nX
Next nY

'****************************************************************************************
'******************************* BARNSLEYJULIA 3 ****************************************
'****************************************************************************************
Case TYPE_BARNSLEYJ3

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((CDbl(nX) * dX) / CDbl(nW))
        rY = SY + ((CDbl(nY) * dY) / CDbl(nH))
        i = 0
        X0 = rX
        Y0 = rY
        
        While i <= MaxI
            If X0 > 0 Then
                X1 = X0 * X0 - Y0 * Y0 - 1
                Y1 = 2 * X0 * Y0
            Else
                X1 = X0 * X0 - Y0 * Y0 - 1 + P0 * X0
                Y1 = 2 * X0 * Y0 + P1 * X0
            End If
            
            XY1 = (X1 * X1) + (Y1 * Y1)
            If XY1 > BailOut Then
                GoTo BailOut_Barnsleyj3
            Else
                i = i + 1
                X0 = X1
                Y0 = Y1
            End If
        Wend
        
        '------------------
        'Inside barnsleyj3 set
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_Barnsleyj3
BailOut_Barnsleyj3:
        '------------------
        
        'Outside barnsleyj3 set
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (X0 * X0) + (Y0 * Y0)
            .nI = i
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_Barnsleyj3:
    Next nX
Next nY

'****************************************************************************************
'******************************* BARNSLEYMANJUL *****************************************
'****************************************************************************************

Case TYPE_BARNSLEYMANJUL

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((CDbl(nX) * dX) / CDbl(nW))
        rY = SY + ((CDbl(nY) * dY) / CDbl(nH))
        i = 0
        X0 = rX
        Y0 = rY
        
        While i <= MaxI
            
            If X0 > 0 Then
                X0 = X0 - 1
            Else
                X0 = X0 + 1
            End If
            
            X1 = pow_r(X0, Y0, 3) + ((rX + 4 * P0) / 5)
            Y1 = pow_i(X0, Y0, 3) + ((rY + 4 * P1) / 5)
            
            XY1 = (X1 * X1) + (Y1 * Y1)
            If XY1 >= BailOut Then
                GoTo BailOut_BarnsleyManJul
            Else
                i = i + 1
                X0 = X1
                Y0 = Y1
            End If
        Wend
        
        '------------------
        'Inside BarnsleyManJul set
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_BarnsleyManJul
BailOut_BarnsleyManJul:
        '------------------
        
        'Outside BarnsleyManJul set
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (X0 * X0) + (Y0 * Y0)
            .nI = i
            .X0 = X0
            .Y0 = Y0
            .X1 = X1
            .Y1 = Y1
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_BarnsleyManJul:
    Next nX
Next nY

'****************************************************************************************
'******************************* S P I D E R ********************************************
'****************************************************************************************
Case TYPE_SPIDER

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((CDbl(nX) * dX) / CDbl(nW))
        rY = SY + ((CDbl(nY) * dY) / CDbl(nH))
        i = 0
        z0.r = rX + P0
        z0.i = rY + P1
        zTmp1.r = rX
        zTmp1.i = rY
        While i <= MaxI
            
            z1.r = z0.r * z0.r - z0.i * z0.i
            z1.i = 2 * z0.r * z0.i
            z1.r = z1.r + zTmp1.r
            z1.i = z1.i + zTmp1.i
            
            zTmp1.r = zTmp1.r / 2
            zTmp1.i = zTmp1.i / 2
            zTmp1.r = zTmp1.r + z1.r
            zTmp1.i = zTmp1.i + z1.i
            
            XY1 = (z1.r * z1.r) + (z1.i * z1.i)
            If XY1 > BailOut Then
                GoTo BailOut_Spider
            Else
                i = i + 1
                z0.r = z1.r
                z0.i = z1.i
            End If
        Wend
        
        '------------------
        'Inside Spider set
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = z0.r
            .Y0 = z0.i
            .X1 = z1.r
            .Y1 = z1.i
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_Spider
BailOut_Spider:
        '------------------
        
        'Outside Spider set
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (z0.r * z0.r) + (z0.i * z0.i)
            .nI = i
            .X0 = z0.r
            .Y0 = z0.i
            .X1 = z1.r
            .Y1 = z1.i
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_Spider:
    Next nX
Next nY

'****************************************************************************************
'******************************* MANDELLAMBDA *******************************************
'****************************************************************************************
Case TYPE_MANDELLAMBDA

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((CDbl(nX) * dX) / CDbl(nW))
        rY = SY + ((CDbl(nY) * dY) / CDbl(nH))
        i = 0
        z0.r = 0.5 + P0
        z0.i = 0 + P1
        zTmp1.r = rX    'Lambda real part
        zTmp1.i = rY    'Lambda imaginary part
        zTmp2.r = 1
        zTmp2.i = 0
        While i <= MaxI
            
            z1 = c_mulc(z0, zTmp1) ' Zn * Lambda
            z1 = c_mulc(z1, c_sub(zTmp2, z0))
            
            XY1 = (z1.r * z1.r) + (z1.i * z1.i)
            If XY1 > BailOut Then
                GoTo BailOut_MandelLambda
            Else
                i = i + 1
                z0.r = z1.r
                z0.i = z1.i
            End If
        Wend
        
        '------------------
        'Inside MandelLambda set
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = z0.r
            .Y0 = z0.i
            .X1 = z1.r
            .Y1 = z1.i
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_MandelLambda
BailOut_MandelLambda:
        '------------------
        
        'Outside MandelLambda set
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (z0.r * z0.r) + (z0.i * z0.i)
            .nI = i
            .X0 = z0.r
            .Y0 = z0.i
            .X1 = z1.r
            .Y1 = z1.i
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_MandelLambda:
    Next nX
Next nY

'****************************************************************************************
'******************************* SIERPINSKI *********************************************
'****************************************************************************************
Case TYPE_SIERPINSKI

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((CDbl(nX) * dX) / CDbl(nW))
        rY = SY + ((CDbl(nY) * dY) / CDbl(nH))
        i = 0
        z0.r = rX
        z0.i = rY
        While i <= MaxI
            
            If z0.i > 0.5 Then
                z1.r = 2 * z0.r
                z1.i = 2 * z0.i - 1
            ElseIf z0.r > 0.5 Then
                z1.r = 2 * z0.r - 1
                z1.i = 2 * z0.i
            Else
                z1.r = 2 * z0.r
                z1.i = 2 * z0.i
            End If
            
            XY1 = (z1.r * z1.r) + (z1.i * z1.i)
            If XY1 > BailOut Then
                GoTo BailOut_Sierpinski
            Else
                i = i + 1
                z0.r = z1.r
                z0.i = z1.i
            End If
        Wend
        
        '------------------
        'Inside Sierpinski set
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = z0.r
            .Y0 = z0.i
            .X1 = z1.r
            .Y1 = z1.i
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_Sierpinski
BailOut_Sierpinski:
        '------------------
        
        'Outside Sierpinski set
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (z0.r * z0.r) + (z0.i * z0.i)
            .nI = i
            .X0 = z0.r
            .Y0 = z0.i
            .X1 = z1.r
            .Y1 = z1.i
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_Sierpinski:
    Next nX
Next nY

'****************************************************************************************
'******************************* N E W T O N  *******************************************
'****************************************************************************************
Case TYPE_NEWTON

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((CDbl(nX) * dX) / CDbl(nW))
        rY = SY + ((CDbl(nY) * dY) / CDbl(nH))
        i = 0
        z0.r = rX
        z0.i = rY
        While i <= MaxI
            
            z1 = c_mul(c_pow(z0, CLng(P0)), CLng(P0 - 1))
            z1.r = z1.r - 1
            zTmp1 = c_mul(c_pow(z0, CLng(P0 - 1)), CLng(P0))
            z1 = c_divc(z1, zTmp1)
            
            'XY1 = z1.r * z1.r + z1.i * z1.i
            'XY0 = z0.r * z0.r + z0.i * z0.i
            X1 = z1.r - z0.r
            Y1 = z1.i - z0.i
            XY1 = X1 * X1 + Y1 * Y1
            If XY1 < BailOut Or XY1 > 1E+20 Then
                GoTo BailOut_Newton
            Else
                i = i + 1
                z0.r = z1.r
                z0.i = z1.i
            End If
        Wend
        
        '------------------
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = z0.r
            .Y0 = z0.i
            .X1 = z1.r
            .Y1 = z1.i
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_Newton
BailOut_Newton:
        '------------------
        
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (z0.r * z0.r) + (z0.i * z0.i)
            .nI = i
            .X0 = z0.r
            .Y0 = z0.i
            .X1 = z1.r
            .Y1 = z1.i
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_Newton:
    Next nX
Next nY

'****************************************************************************************
'******************************* NEWTON-JULIA-NOVA  *******************************************
'****************************************************************************************
Case TYPE_NEWTONJULIANOVA

With Frac(nFrac)
    SX = .SX: SY = .SY
    EX = .EX: EY = .EY
    nW = .nW: nH = .nH
    MaxI = .MaxI
    BailOut = .BailOut
End With

dX = EX - SX
dY = EY - SY

For nY = StartLine To EndLine
    If OnlyOneLine Then
        nLine = 0
    Else
        nLine = nY
    End If
    For nX = 0 To nW - 1
        rX = SX + ((CDbl(nX) * dX) / CDbl(nW))
        rY = SY + ((CDbl(nY) * dY) / CDbl(nH))
        i = 0
        z0.r = 1
        z0.i = 0
        zTmp3.r = 1
        zTmp3.i = 0
        While i <= MaxI
            
            zTmp2 = z0
            zTmp1 = c_mulc(z0, z0)
            z1 = c_sub(z0, c_divc(c_sub(c_mulc(z0, zTmp1), zTmp3), c_mul(zTmp1, 3)))
            z1.r = z1.r + P0
            z1.i = z1.i + P1
            
            'z=z-((z*z1-1)/(3*z1))+c;
            'z3 = z-z2;
            'd = sum_sqrs_z3();

            
            'XY1 = z1.r * z1.r + z1.i * z1.i
            'XY0 = z0.r * z0.r + z0.i * z0.i
            X1 = z1.r - z0.r
            Y1 = z1.i - z0.i
            XY1 = X1 * X1 + Y1 * Y1
            If XY1 < BailOut Or XY1 > 1E+20 Then
                GoTo BailOut_NewtonJuliaNova
            Else
                i = i + 1
                z0.r = z1.r
                z0.i = z1.i
            End If
        Wend
        
        '------------------
        With Frac(nFrac).Cell(nX, nLine)
            .nI = 0
            .X0 = z0.r
            .Y0 = z0.i
            .X1 = z1.r
            .Y1 = z1.i
            .rX = rX
            .rY = rY
            .State = 1
        End With
        '------------------
        GoTo CalcNext_NewtonJuliaNova
BailOut_NewtonJuliaNova:
        '------------------
        
        With Frac(nFrac).Cell(nX, nLine)
            XY0 = (z0.r * z0.r) + (z0.i * z0.i)
            .nI = i
            .X0 = z0.r
            .Y0 = z0.i
            .X1 = z1.r
            .Y1 = z1.i
            .rX = rX
            .rY = rY
            .State = 0
        End With
        '------------------
CalcNext_NewtonJuliaNova:
    Next nX
Next nY


End Select

SetThreadPriority App.ThreadID, THREAD_PRIORITY_NORMAL

End Sub
