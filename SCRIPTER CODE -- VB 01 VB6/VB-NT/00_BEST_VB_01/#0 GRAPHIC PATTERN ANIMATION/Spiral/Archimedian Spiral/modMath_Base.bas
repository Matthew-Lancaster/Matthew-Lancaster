Attribute VB_Name = "modMath_Base"
'' Property Editor Active-X Control
''
'' Version 2.0
''
''
'' Basic Math Functions

'' Copyright (C) 2001, 2002 GUInerd.
'' All Rights Reserved.

'' Unauthorized Use Strictly Prohibited.

Option Explicit

Public Const Pi = 3.14159265358979

Public Type CARTESEAN2D
    
    Origin As POINTAPI
    Points As POINTAPI
    
    Arc As Double
    Distance As Double

End Type

Public Function ModStep(Number, Modulus)
    On Error Resume Next
    
    If (Number < Modulus) Then
        ModStep = Number + (4 - Modulus)
    ElseIf (Number Mod Modulus) Then
        ModStep = Number + (Modulus - (Number Mod Modulus))
    Else
        ModStep = Number
    End If
    
End Function

Public Function Swab(ByVal Value As Long) As Long
    
    Dim vN1(0 To 3), _
        vn2(0 To 3)
        
    Dim c As Long
    
    CopyMemory vN1(0), Value, 4&
    
    vn2(3) = vN1(0)
    vn2(2) = vN1(1)
    vn2(1) = vN1(2)
    vn2(0) = vN1(3)
    
    CopyMemory c, vn2(0), Value
    Swab = c
    
End Function

Public Function Max(ParamArray Numbers()) As Variant

    Dim i As Long, _
        j As Long, _
        l As Long
        
    Dim c As Variant, _
        n As Variant
        
    On Error Resume Next
    
    l = -1&
    l = LBound(Numbers)
    If (l = -1&) Then Exit Function
    
    j = UBound(Numbers)
    
    c = Numbers(l)
    For i = l To j
    
        n = Numbers(i)
        If (n > c) Then c = n
    
    Next i

    Max = c
    
End Function

Public Function Min(ParamArray Numbers()) As Variant

    Dim i As Long, _
        j As Long, _
        l As Long
        
    Dim c As Variant, _
        n As Variant
        
    On Error Resume Next
    
    l = -1&
    l = LBound(Numbers)
    If (l = -1&) Then Exit Function
    
    j = UBound(Numbers)
    
    c = Numbers(l)
    For i = l To j
    
        n = Numbers(i)
        If (n < c) Then c = n
    
    Next i

    Min = c
    
End Function

Public Function Floor(Number) As Variant
    Dim pResult As Variant, _
        n As Variant
        
    n = Round(Number)
    
    If (Number >= 0) Then
        If (n > Number) Then
            n = n - 1
        End If
    Else
        If (n < Number) Then
            n = n + 1
        End If
    End If
    
    Floor = n
    
End Function

Public Function Ceil(Number) As Variant
    Dim pResult As Variant, _
        n As Variant
    
    n = Round(Number)
    
    If (Number >= 0) Then
        If (n < Number) Then
            n = n + 1
        End If
    Else
        If (n > Number) Then
            n = n - 1
        End If
    End If
    
    Ceil = n
    
End Function

Public Sub GetLinearPoints(lpChart As CARTESEAN2D)
    
    Dim c As Double, _
        f As Double, _
        d As Double
    
    Dim q As Long, _
        h As Long
    
    If (lpChart.Distance = 0) Or (lpChart.Arc = -1#) Then
        CopyMemory lpChart.Points, lpChart.Origin, Len(lpChart.Points)
        Exit Sub
    End If
    
    c = lpChart.Arc
    
    If (c < 0) Then c = c + 360
    If (c = 360) Then c = 0
            
    q = Floor(c / 90)
        
    If (q = (c / 90)) Then
        q = q - 0.1
        If (q < 0) Then q = 3
    End If
    
    c = c * (Pi / 180)
    c = Tan(c)
    
    d = Sqr((lpChart.Distance * lpChart.Distance) / (1 + (c * c)))

    '' now convert from classical cartesian geometry to "Geometry95"
    
    f = (d)
    d = (f * c)
    
    If (q = 2) Or (q = 1) And (lpChart.Arc <> 90) Then
        d = -d
    End If
            
    If (lpChart.Arc = 270) And (q = 3) Then
        d = -d
    End If
    
    If (q = 0) Or (q = 3) Then
        f = -f
    End If
    
    '' got the points
    
    With lpChart
        .Points.x = .Origin.x + d
        .Points.y = .Origin.y + f
    End With
    
End Sub

Public Sub GetDistanceArc(lpChart As CARTESEAN2D)

    Dim lpTrue As POINTAPI
    
    Dim a As Double, _
        b As Double, _
        c As Double, _
        d As Double
    
    Dim q As Integer
    
    '' convert Geometry95 to classical geometry
    
    lpTrue.x = lpChart.Points.x - lpChart.Origin.x
    lpTrue.y = lpChart.Origin.y - lpChart.Points.y
    
    a = lpTrue.x
    b = lpTrue.y
    
    If (a = 0) And (b = 0) Then
        lpChart.Distance = 0#
        CopyMemory lpChart.Points, lpChart.Origin, 8&
        Exit Sub
    End If
    
    '' yeah, Pythagoras!
    
    c = Sqr((a * a) + (b * b))
    lpChart.Distance = c
    
    If (lpTrue.x >= 0&) And (lpTrue.y >= 0&) Then
        q = 0&
    ElseIf (lpTrue.x >= 0&) And (lpTrue.y < 0&) Then
        q = 1&
    ElseIf (lpTrue.y >= 0&) And (lpTrue.x < 0&) Then
        q = 3&
    Else
        q = 2&
    End If
    
    c = 0#
    
    If (b = 0) Then
        c = (180 * -(a < 0)) + 90
    ElseIf (a = 0) Then
        c = (180 * -(b < 0))
    Else
        
        c = Atn(b / a)
        c = c * (180 / Pi)
                
        c = Abs(c)
        
        '' now convert from classical cartesian geometry to "Geometry95"
        
        If ((q + 1) Mod 2) Then c = Abs(90 - c)
        
        c = c + (q * 90)
        
        If (c >= 360) Then
            c = c - 360
        ElseIf (c < 0) Then
            c = c + 360
        End If
        
    End If
    
    lpChart.Arc = c

End Sub


Public Function PrecisePi(n)

    Dim a, b, c, d, e, f
    Static sum
    
    If n = 0@ Then sum = 0@
    
    a = 0@
    b = 0@
    c = 0@
    d = 0@
    e = 0@
    f = 0@
    
    a = 4 / ((8 * n) + 1)
    b = 2 / ((8 * n) + 4)
    c = 1 / ((8 * n) + 5)
    d = 1 / ((8 * n) + 6)
    
    e = (1 / 16)
    
    f = sum + ((a - b - c - d) * (e ^ n))
    sum = f
    
    PrecisePi = f

End Function
