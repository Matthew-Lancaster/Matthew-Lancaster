Attribute VB_Name = "modComplex"
Option Explicit

Public Cnt As Long
Public Tmp1 As Double, Tmp2 As Double, Tmp3 As Double, Tmp4 As Double
'----------------------------------------------------------------------------------


'**************************************************
'Standard derived math functions (rational numbers):
Public Function Cosh(x As Double) As Double
    Cosh = 0.5 * (Exp(x) + Exp(-x))
End Function
Public Function Sinh(x As Double) As Double
    Sinh = 0.5 * (Exp(x) - Exp(-x))
End Function
'**************************************************

'----------------------------------------------------------------------------------

'**********************************************************************************
'C O M P L E X - M A T H - F U N C T I O N S :
'**********************************************************************************
Public Function exp_r(a As Double, b As Double) As Double
    exp_r = Exp(a) * Cos(b)
End Function
Public Function exp_i(a As Double, b As Double) As Double
    exp_i = Exp(a) * Sin(b)
End Function
Public Function cosh_r(a As Double, b As Double) As Double
    cosh_r = ((Exp(a) + Exp(-a)) / 2) * Cos(b)
End Function
Public Function cosh_i(a As Double, b As Double) As Double
    cosh_i = ((Exp(a) - Exp(-a)) / 2) * Sin(b)
End Function
Public Function mul_r(a As Double, b As Double, c As Double, d As Double) As Double
    mul_r = (a * c) - (b * d)
End Function
Public Function mul_i(a As Double, b As Double, c As Double, d As Double) As Double
    mul_i = (a * d) + (b * c)
End Function
Public Function pow_r(ByVal a As Double, ByVal b As Double, n As Long) As Double
    Tmp1 = a
    Tmp2 = b
    For Cnt = 1 To n - 1
        Tmp3 = (Tmp1 * a) - (Tmp2 * b)
        Tmp4 = (Tmp1 * b) + (Tmp2 * a)
        Tmp1 = Tmp3
        Tmp2 = Tmp4
    Next Cnt
    pow_r = Tmp1
End Function
Public Function pow_i(ByVal a As Double, ByVal b As Double, n As Long) As Double
    Tmp1 = a
    Tmp2 = b
    For Cnt = 1 To n - 1
        Tmp3 = (Tmp1 * a) - (Tmp2 * b)
        Tmp4 = (Tmp1 * b) + (Tmp2 * a)
        Tmp1 = Tmp3
        Tmp2 = Tmp4
    Next Cnt
    pow_i = Tmp2
End Function

Public Function c_add(z1 As tComplex, z2 As tComplex) As tComplex
    c_add.r = z1.r + z2.r
    c_add.i = z1.i + z2.i
End Function
Public Function c_sub(z1 As tComplex, z2 As tComplex) As tComplex
    c_sub.r = z1.r - z2.r
    c_sub.i = z1.i - z2.i
End Function
Public Function c_mul(z As tComplex, fact As Double) As tComplex
    c_mul.r = z.r * fact
    c_mul.i = z.i * fact
End Function
Public Function c_mulc(z1 As tComplex, z2 As tComplex) As tComplex
    c_mulc.r = z1.r * z2.r - z1.i * z2.i
    c_mulc.i = z1.i * z2.r + z1.r * z2.i
End Function
Public Function c_divc(z1 As tComplex, z2 As tComplex) As tComplex
    Tmp1 = z2.r * z2.r + z2.i * z2.i
    If Tmp1 <> 0 Then
        c_divc.r = (z1.r * z2.r + z1.i * z2.i) / Tmp1
        c_divc.i = (z1.i * z2.r - z1.r * z2.i) / Tmp1
    End If
End Function
Public Function c_pow(z As tComplex, n As Long) As tComplex
    Tmp1 = z.r
    Tmp2 = z.i
    For Cnt = 1 To n - 1
        Tmp3 = (Tmp1 * z.r) - (Tmp2 * z.i)
        Tmp4 = (Tmp1 * z.i) + (Tmp2 * z.r)
        Tmp1 = Tmp3
        Tmp2 = Tmp4
    Next Cnt
    c_pow.r = Tmp1
    c_pow.i = Tmp2
End Function
Public Function c_powc(z As tComplex, w As tComplex) As tComplex
    c_powc = c_exp(c_mulc(w, c_log(z)))
End Function
Public Function c_exp(z As tComplex) As tComplex
    c_exp.r = Exp(z.r) * Cos(z.i)
    c_exp.i = Sin(z.i)
End Function
Public Function c_log(z As tComplex) As tComplex
    If z.r <> 0 Then
        c_log.r = 0.5 * Log(z.r * z.r + z.i * z.i)
        c_log.i = Atn(z.i / z.r)
    End If
End Function
Public Function c_sin(z As tComplex) As tComplex
    c_sin.r = Sin(z.r) * 0.5 * (Exp(z.i) + Exp(-z.i))
    c_sin.i = Cos(z.r) * 0.5 * (Exp(z.i) - Exp(-z.i))
End Function
Public Function c_cos(z As tComplex) As tComplex
    c_cos.r = Cos(z.r) * 0.5 * (Exp(z.i) + Exp(-z.i))
    c_cos.i = -(Sin(z.r) * 0.5 * (Exp(z.i) - Exp(-z.i)))
End Function
Public Function c_tan(z As tComplex) As tComplex
    Tmp1 = Cos(2 * z.r) + (0.5 * (Exp(2 * z.i) + Exp(-2 * z.i)))
    c_tan.r = Sin(2 * z.r) / Tmp1
    c_tan.i = (0.5 * (Exp(2 * z.i) - Exp(-2 * z.i))) / Tmp1
End Function
Public Function c_sinh(z As tComplex) As tComplex
    c_sinh.r = 0.5 * (Exp(z.r) - Exp(-z.r)) * Cos(z.i)
    c_sinh.i = 0.5 * (Exp(z.r) + Exp(-z.r)) * Sin(z.i)
End Function
Public Function c_cosh(z As tComplex) As tComplex
    c_cosh.r = 0.5 * (Exp(z.r) + Exp(-z.r)) * Cos(z.i)
    c_cosh.i = 0.5 * (Exp(z.r) - Exp(-z.r)) * Sin(z.i)
End Function
Public Function c_cotan(z As tComplex) As tComplex
    Tmp1 = Cosh(2 * z.i) - Cos(2 * z.r)
    c_cotan.r = Sin(2 * z.r) / Tmp1
    c_cotan.i = -Sinh(2 * z.i) / Tmp1
End Function
Public Function c_cotanh(z As tComplex) As tComplex
    Tmp1 = Cosh(2 * z.r) - Cos(2 * z.i)
    c_cotanh.r = Sinh(2 * z.r) / Tmp1
    c_cotanh.i = -Sin(2 * z.i) / Tmp1
End Function
Public Function c_mag(z As tComplex) As Double
    c_mag = z.r * z.r + z.i * z.i
End Function
Public Function c_abs(z As tComplex) As Double
    c_abs = Sqr(z.r * z.r + z.i * z.i)
End Function
'**********************************************************************************
'**********************************************************************************
