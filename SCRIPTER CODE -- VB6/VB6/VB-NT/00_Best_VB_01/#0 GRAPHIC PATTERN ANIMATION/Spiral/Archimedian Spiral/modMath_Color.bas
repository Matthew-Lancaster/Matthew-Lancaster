Attribute VB_Name = "modMath_Color"
Option Explicit
Option Base 0

Public Const SF_VALUE = &H0
Public Const SF_SATURATION = &H1
Public Const SF_HUE = &H2

Public Type PALETTEINFOEX
    
    FlipHorizontal As Boolean
    FlipVertical As Boolean
    
    GrayTones As Boolean
    
    StaticField As Integer
    Value As Integer
    
    BlockSize As Integer
    
End Type

Public Type RGBDATA
    Blue As Byte
    Green As Byte
    Red As Byte
End Type

Public Type BGRDATA
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Public Type CMYDATA
    Cyan As Byte
    Magenta As Byte
    Yellow As Byte
End Type

Public Type HSVDATA
    Hue As Integer
    Saturation As Integer
    Value As Integer
End Type

Public Type HSLDATA
    Hue As Integer
    Saturation As Integer
    Lightness As Integer
End Type

Public Function Cex(ByVal Color As Long) As String
    
    Dim s As String
    
    s = Hex(Color)
    Cex = "&H" + String(6 - Len(s), "0") + s
    
End Function

Public Function GetRGB(ByVal crColor As Long, Red, Green, Blue)

    Dim i As Long
    
    i = (crColor And &HFF&)
    Red = (i And &HFF&)
    
    i = ((crColor And &HFF00&) / &H100&)
    Green = (i And &HFF&)
    
    i = ((crColor And &HFF0000) / &H10000)
    Blue = (i And &HFF&)
    
End Function

'' Single Convert ColorRef to RGB

Public Sub ColorToRGB(ByVal Color As Long, Bits As RGBDATA)
    GetRGB Color, Bits.Red, Bits.Green, Bits.Blue
End Sub

'' Single Convert RGB to ColorRef

Public Function RGBToColor(Bits As RGBDATA) As Long
    RGBToColor = RGB(Bits.Red, Bits.Green, Bits.Blue)
End Function

'' Single Convert ColorRef to RGB-reversed

Public Sub ColorToBGR(ByVal Color As Long, Bits As BGRDATA)
    CopyMemory Bits, Color, 3&
End Sub

'' Single Convert RGB-reversed to ColorRef

Public Function BGRToColor(Bits As BGRDATA) As Long
    Dim c As Long
    
    CopyMemory c, Bits, 3&
    BGRToColor = c
End Function

Public Sub ColorToHSV(ByVal crColor As Long, hsv As HSVDATA)
    On Error Resume Next
        
    Dim r As Double, _
        s As Double, _
        t As Double
        
    Dim n As Double
        
    Dim a As Long, _
        b As Long
        
    Dim f As Long, _
        g As Long
                
    Dim cr As RGBDATA
    
    ColorToRGB crColor, cr
    
    a = Min(cr.Red, cr.Green, cr.Blue)
    b = Max(cr.Red, cr.Green, cr.Blue)
        
    If (a = b) Then
        
        hsv.Hue = -1
        
        r = (a / 255)
        
        With hsv
            Select Case r
            
                Case Is <= 0.5
                    
                    .Saturation = 360
                    .Value = (510 * r)
            
                Case Is <= 1
                
                    r = 1 - r
                    
                    .Value = 255
                    .Saturation = (720 * r)
                
            End Select
        
        End With
        
        
        
        Exit Sub
        
    End If
    
    hsv.Value = Round((b / 255) * 255)
    
    Select Case b
        
        Case cr.Red
            f = 0
            
            If (a = cr.Blue) Then
                r = (cr.Green / b) * 60
                f = r
            Else
                r = (cr.Blue / b) * 60
                f = 360 - r
            End If
        
        Case cr.Green
            f = 120
            
            If (a = cr.Blue) Then
                r = (cr.Red / b) * 60
                f = f - r
            Else
                r = (cr.Blue / b) * 60
                f = f + r
            End If
        
        Case cr.Blue
            f = 240
                        
            If (a = cr.Green) Then
                r = (cr.Red / b) * 60
                f = f + r
                
                If (f >= 360) Then f = f - 360
            Else
                r = (cr.Green / b) * 60
                f = f - r
            End If
    
    End Select
    
    s = (a / 255) * 360
    g = 360 - s
        
    hsv.Hue = f
    hsv.Saturation = g
        
End Sub

Public Function HSVToColor(hsv As HSVDATA) As Long

    On Error Resume Next

    Dim a As Integer, _
        b As Integer, _
        c As Integer
        
    Dim d As Single, _
        e As Single
        
    Dim i As Long, _
        j As Long
        
    Dim n As Integer
        
    Dim cr As RGBDATA
    
    If (hsv.Hue = 360) Then hsv.Hue = 0
    
    If (hsv.Hue = -1) Then
        
        
        With hsv
            If (.Saturation / 360) > (.Value / 255) Then
                .Saturation = 360
                e = (.Value / 510)
            Else
                .Value = 255
                e = 1 - (.Saturation / 720)
            End If
        
        End With
        
        a = e * 255
        HSVToColor = RGB(a, a, a)
        
        Exit Function
    End If
    
    If (hsv.Hue = 360) Then hsv.Hue = 0
    
    a = hsv.Value
    d = hsv.Saturation
    
    d = 360 - d
    
    d = (d / 360) * a
    c = d
    
    d = (hsv.Hue / 360)
    n = Floor(d * 6)
    
    Select Case n
    
        Case 0, 6:    '' 0, 360 - Red
            '' b is delta -> green
            e = hsv.Hue
            e = (e / 60)
            
            e = e * (a - c)
            e = e + c
            
            b = Floor(e)
            j = RGB(a, b, c)
        
        Case 1:       '' 60 - Yellow
            '' b is delta <- red
            e = hsv.Hue - 60
            e = ((60 - e) / 60)
            
            e = e * (a - c)
            e = e + c
            
            b = Floor(e)
            j = RGB(b, a, c)
        
        Case 2:       '' 120 - Green
            '' b is delta -> blue
            e = hsv.Hue - 120
            e = (e / 60)
            
            e = e * (a - c)
            e = e + c
            
            b = Floor(e)
            j = RGB(c, a, b)
        
        Case 3:       '' 180 - Cyan
            '' b is delta <- green
            e = hsv.Hue - 180
            e = ((60 - e) / 60)
            
            e = e * (a - c)
            e = e + c
            
            b = Floor(e)
            j = RGB(c, b, a)
        
        Case 4:       '' 240 - Blue
            '' b is delta -> red
            e = hsv.Hue - 240
            e = (e / 60)
            
            e = e * (a - c)
            e = e + c
            
            b = Floor(e)
            j = RGB(b, c, a)
                
        Case 5:       '' 300 - Magenta
            '' b is delta <- blue
            e = hsv.Hue - 300
            e = ((60 - e) / 60)
            
            e = e * (a - c)
            e = e + c
            
            b = Floor(e)
            j = RGB(a, c, b)
        
    End Select
      
    HSVToColor = j
    
End Function

Public Sub ColorToCMY(ByVal Color As Long, cmy As CMYDATA)
''
    Dim r As Long, _
        g As Long, _
        b As Long
        
    Dim x As Long
        
    GetRGB Color, r, g, b
    
    x = Max(r, g, b)
    
    r = Abs(1 - r)
    g = Abs(1 - g)
    b = Abs(1 - b)
    
    cmy.Magenta = r
    cmy.Yellow = g
    cmy.Cyan = b
            
End Sub

Public Function CMYToColor(cmy As CMYDATA) As Long
''
    Dim c As Long, _
        m As Long, _
        y As Long, _
        k As Long
        
    Dim r As Long, _
        g As Long, _
        b As Long
        
    c = cmy.Cyan
    m = cmy.Magenta
    y = cmy.Yellow
        
    r = Abs((1 - m))
    g = Abs((1 - y))
    b = Abs((1 - c))
    
    CMYToColor = RGB(r, g, b)
    
End Function

Public Sub GetPercentages(ByVal crColor As Long, dpRed As Double, dpGreen As Double, dpBlue As Double)
    
    Dim vR As Double, _
        vB As Double, _
        vG As Double
        
    Dim vIn(0 To 2) As Byte, _
        sR As Double, _
        sG As Double, _
        sB As Double

    Dim d As Double
        
    GetRGB crColor, vIn(0), vIn(1), vIn(2)
    
    vR = CDbl(vIn(0))
    vG = CDbl(vIn(1))
    vB = CDbl(vIn(2))
            
    If (vR > vG) Then d = vR Else d = vG
    If (vB > d) Then d = vB
    
    If (d = 0&) Then d = 255#
    
    dpRed = (vR / d)
    dpGreen = (vG / d)
    dpBlue = (vB / d)
        
End Sub

Public Function ColorVariance(ByVal Color1 As Long, ByVal Color2 As Long) As Double
    Dim rgb1 As RGBDATA, _
        rgb2 As RGBDATA
        
    Dim vR As Double, _
        vG As Double, _
        vB As Double
        
    Dim vO As Double
        
    ColorToRGB Color1, rgb1
    ColorToRGB Color2, rgb2
    
    vR = Abs(CDbl(rgb2.Red) - CDbl(rgb1.Red))
    vG = Abs(CDbl(rgb2.Green) - CDbl(rgb1.Green))
    vB = Abs(CDbl(rgb2.Blue) - CDbl(rgb1.Blue))
    
    vO = (vR + vG + vB) / 3#
    ColorVariance = vO
    
End Function

Public Function AbsTone(ByVal crColor As Long, Optional ByVal Percent As Double = 1#) As Long
    
    Dim pR As Double, _
        pB As Double, _
        pG As Double
    
    Dim sR As Double, _
        sB As Double, _
        sG As Double
    
    GetPercentages crColor, pR, pG, pB
        
    sR = (pR) * (255 * Percent)
    sG = (pG) * (255 * Percent)
    sB = (pB) * (255 * Percent)
        
    AbsTone = RGB(CInt(sR), CInt(sG), CInt(sB))
    
End Function

Public Function GetAverageColor(ByVal Color1 As Long, ByVal Color2 As Long, Optional ByVal Track As Long = 100, Optional ByVal Value As Long = 50) As Long

    Dim r1 As Double, r2 As Double, r3 As Double, _
        g1 As Double, g2 As Double, g3 As Double, _
        b1 As Double, b2 As Double, b3 As Double
        
    Dim x As Double, _
        y As Double
        
    GetRGB Color1, r1, g1, b1
    GetRGB Color2, r2, g2, b2
    
    x = Max(r1, r2)
    y = Min(r1, r2)
    
    r3 = r1 + Floor(CDbl((r2 - r1) / Track) * Value)
    
    x = Max(g1, g2)
    y = Min(g1, g2)
    
    g3 = g1 + Floor(CDbl((g2 - g1) / Track) * Value)
    
    x = Max(b1, b2)
    y = Min(b1, b2)
    
    b3 = b1 + Floor(CDbl((b2 - b1) / Track) * Value)
    
    GetAverageColor = RGB(CInt(r3), CInt(g3), CInt(b3))

End Function

Public Function GrayTone(ByVal crColor As Long) As Long

    Dim lpRGB As RGBDATA, _
        a As Long, _
        b As Long, _
        c As Long
        
    
    ColorToRGB crColor, lpRGB
    
    a = lpRGB.Red
    b = lpRGB.Green
    c = lpRGB.Blue
    
    a = (a + b + c) / 3
    
    GrayTone = RGB(a, a, a)
    
End Function

            
Public Sub WheelData(ByVal cx As Long, ByVal cy As Long, lpWheel As CARTESEAN2D)

    Dim dia As Double
        
    If (cx < cy) Then dia = cx Else dia = cy
    
    dia = dia - 12#
    
    lpWheel.Origin.x = Round(cx / 2#)
    lpWheel.Origin.y = Round(cy / 2#)
    
    lpWheel.Distance = (dia / 2)
    
End Sub








