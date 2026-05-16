Attribute VB_Name = "modDraw_Base"
''' GUInerd Standard Menu System
''' Version 4.0

'''

''' Common Base-Level Graphics Routines


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit
Option Base 0

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

'' Global Variables for Unicode (Used By the Text Functions)

Public Unicode_Checked As Boolean
Public Using_Unicode As Boolean

Public Sub CheckUnicode()
    Dim lpOS As OSVERSIONINFO
    
    lpOS.OSVersionInfoSize = Len(lpOS)
    GetVersion lpOS
    
    If (lpOS.PlatformId = VER_PLATFORM_WIN32_NT) Then
        Using_Unicode = True
    Else
        Using_Unicode = False
    End If
    
    Unicode_Checked = True
    
End Sub

''' Draw 3D Borders

'' Heavy Sunken (3D Controls)

Public Function Draw3DFrame_Sunken(ByVal hDC As Long, lpRect As RECT, Optional ByVal dwPenStyle As Long = PS_SOLID)
    Dim hPen As Long, _
        x As Long, _
        y As Long, _
        lpPoint As POINTAPI
    
    Dim hOldPen As Long
    
    hPen = CreatePen(dwPenStyle, 2&, GetSysColor(COLOR_3DDKSHADOW))
    hOldPen = SelectObject(hDC, hPen)
    
    With lpRect
        MoveToEx hDC, .Left + 1, .Bottom - 1, lpPoint
        LineTo hDC, .Left + 1, .Top + 1
        
        LineTo hDC, .Left, .Top + 1
        LineTo hDC, .Right - 1, .Top + 1
    End With
    
    SelectObject hDC, hOldPen
    DeleteObject hPen
    
    hPen = CreatePen(dwPenStyle, 1&, GetSysColor(COLOR_BTNFACE))
    SelectObject hDC, hPen
    
    With lpRect
        MoveToEx hDC, .Right - 2&, .Top, lpPoint
        LineTo hDC, .Right - 2&, .Bottom - 2&
        LineTo hDC, .Left, .Bottom - 2&
    End With
    
    SelectObject hDC, hOldPen
    DeleteObject hPen
    
    hPen = CreatePen(dwPenStyle, 2&, GetSysColor(COLOR_BTNHIGHLIGHT))
    SelectObject hDC, hPen
    
    With lpRect
        MoveToEx hDC, .Right, .Top, lpPoint
        LineTo hDC, .Right, .Bottom
        LineTo hDC, .Left, .Bottom
        
    End With
        
    SelectObject hDC, hOldPen
    DeleteObject hPen
    
End Function


'' Light Sunken (Selected Items)

Public Function Draw3DBox_Sunken(ByVal hDC As Long, lpRect As RECT, Optional ByVal dwPenStyle As Long = PS_SOLID)
    Dim hPen As Long, _
        x As Long, _
        y As Long, _
        lpPoint As POINTAPI
    
    Dim hOldPen As Long
    
    hPen = CreatePen(dwPenStyle, 1&, GetSysColor(COLOR_3DDKSHADOW))
    hOldPen = SelectObject(hDC, hPen)
    
    With lpRect
        MoveToEx hDC, .Left, .Bottom - 1, lpPoint
        LineTo hDC, .Left, .Top
        
        LineTo hDC, .Left, .Top
        LineTo hDC, .Right, .Top
    End With
    
    SelectObject hDC, hOldPen
    DeleteObject hPen
    
    hPen = CreatePen(dwPenStyle, 1&, GetSysColor(COLOR_BTNFACE))
    SelectObject hDC, hPen
    
    With lpRect
        MoveToEx hDC, .Right - 1&, .Top, lpPoint
        LineTo hDC, .Right - 1&, .Bottom - 1&
        LineTo hDC, .Left, .Bottom - 1&
    End With
    
    SelectObject hDC, hOldPen
    DeleteObject hPen
    
    hPen = CreatePen(dwPenStyle, 1&, GetSysColor(COLOR_BTNHIGHLIGHT))
    SelectObject hDC, hPen
    
    With lpRect
        MoveToEx hDC, .Right, .Top, lpPoint
        LineTo hDC, .Right, .Bottom
        LineTo hDC, .Left, .Bottom
        
    End With
        
    SelectObject hDC, hOldPen
    DeleteObject hPen
    
End Function


'' Heavy Raised (3D Buttons, Forms or Menus)

Public Function Draw3DFrame_Raised(ByVal hDC As Long, lpRect As RECT)

    Dim hPen As Long, _
        hOldPen As Long
    
    Dim lpPoint As POINTAPI
        
    hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNHIGHLIGHT))
    
    hOldPen = SelectObject(hDC, hPen)
    
    MoveToEx hDC, lpRect.Right, lpRect.Top, lpPoint
    LineTo hDC, lpRect.Left, lpRect.Top
    LineTo hDC, lpRect.Left, lpRect.Bottom
    
    SelectObject hDC, hOldPen
    DeleteObject hPen
    
    hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNSHADOW))
    
    SelectObject hDC, hPen
    
    MoveToEx hDC, lpRect.Right - 1, lpRect.Top, lpPoint
    LineTo hDC, lpRect.Right - 1, lpRect.Bottom - 1
    LineTo hDC, lpRect.Left, lpRect.Bottom - 1
    
    SelectObject hDC, hOldPen
    DeleteObject hPen

    hPen = CreatePen(PS_SOLID, 1, 0&)
    
    SelectObject hDC, hPen
    
    MoveToEx hDC, lpRect.Right, lpRect.Top, lpPoint
    LineTo hDC, lpRect.Right, lpRect.Bottom
    LineTo hDC, lpRect.Left, lpRect.Bottom
    
    SelectObject hDC, hOldPen
    DeleteObject hPen

End Function

'' Light Raised (Highlighted/Bordered)

Public Function Draw3DBox_Raised(ByVal hDC As Long, lpRect As RECT)

    Dim hPen As Long, _
        hOldPen As Long
    
    Dim lpPoint As POINTAPI
        
    hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNHIGHLIGHT))
    
    hOldPen = SelectObject(hDC, hPen)
    
    MoveToEx hDC, lpRect.Right, lpRect.Top, lpPoint
    LineTo hDC, lpRect.Left, lpRect.Top
    LineTo hDC, lpRect.Left, lpRect.Bottom
    
    SelectObject hDC, hOldPen
    DeleteObject hPen
    
    hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNSHADOW))
    
    SelectObject hDC, hPen
    
    MoveToEx hDC, lpRect.Right, lpRect.Top, lpPoint
    LineTo hDC, lpRect.Right, lpRect.Bottom
    LineTo hDC, lpRect.Left, lpRect.Bottom
    
    SelectObject hDC, hOldPen
    DeleteObject hPen

End Function

'' Dotted Work Frame

Public Function DrawWorkFrame(ByVal hDC As Long, lpRect As RECT)
    Dim hPen As Long, _
        hOldPen As Long
        
    Dim lpPoint As POINTAPI
        
    hPen = CreatePen(PS_DOT, 1&, GetSysColor(COLOR_HIGHLIGHT))
    hOldPen = SelectObject(hDC, hPen)
    
    With lpRect
        MoveToEx hDC, .Left, .Top, lpPoint

        LineTo hDC, .Right, .Top
        LineTo hDC, .Right, .Bottom
        LineTo hDC, .Left, .Bottom
        LineTo hDC, .Left, .Top
    
        MoveToEx hDC, .Left + 1, .Top + 1, lpPoint

        LineTo hDC, .Right - 1, .Top + 1
        LineTo hDC, .Right - 1, .Bottom - 1
        LineTo hDC, .Left + 1, .Bottom - 1
        LineTo hDC, .Left + 1, .Top + 1
    
    End With
    
    SelectObject hDC, hOldPen
    DeleteObject hPen
    
End Function


'' Bevel Line


Public Function BevelLine(ByVal hDC As Long, PointA As POINTAPI, PointB As POINTAPI, Optional ByVal Vertical As Boolean, Optional ByVal Inverse As Boolean)

    Dim hPen As Long, _
        lpPoint As POINTAPI, _
        hOldPen As Long
        
    hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNHIGHLIGHT))
    hOldPen = SelectObject(hDC, hPen)
    
    MoveToEx hDC, PointA.x, PointA.y, lpPoint
    LineTo hDC, PointB.x, PointB.y
    
    SelectObject hDC, hOldPen
    DeleteObject hPen
    
    hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNSHADOW))
    SelectObject hDC, hPen
        
    If (Vertical = True) Then
        
        If Inverse = True Then
            MoveToEx hDC, PointA.x - 1, PointA.y, lpPoint
            LineTo hDC, PointB.x - 1, PointB.y
        Else
            MoveToEx hDC, PointA.x + 1, PointA.y, lpPoint
            LineTo hDC, PointB.x + 1, PointB.y
        End If
    Else
        If Inverse = True Then
            MoveToEx hDC, PointA.x, PointA.y + 1, lpPoint
            LineTo hDC, PointB.x, PointB.y + 1
        Else
            MoveToEx hDC, PointA.x, PointA.y - 1, lpPoint
            LineTo hDC, PointB.x, PointB.y - 1
        End If
    End If
    
    SelectObject hDC, hOldPen
    DeleteObject hPen

    MoveToEx hDC, lpPoint.x, lpPoint.y, lpPoint
    
End Function



'''' Special Items

''' Render Graphics

'' Render Standard Picture (Icon/Bitmap/Cursor)

Public Function Render_Picture(ByVal hDC As Long, ByVal objPicture As StdPicture, _
                               lpRect As RECT, _
                               Optional ByVal AsDisabled As Boolean, _
                               Optional ByVal crBackColor As Long = -1&) As Long
                               
    Dim fAsDisabled As Boolean, _
        crColor As Long
    
    crColor = crBackColor
    fAsDisabled = AsDisabled
                               
    Select Case objPicture.Type
        Case vbPicTypeBitmap
            Render_Picture = Render_Image(hDC, objPicture.Handle, IMAGE_BITMAP, lpRect, fAsDisabled, crColor)
    
        Case vbPicTypeIcon
            Render_Picture = Render_Image(hDC, objPicture.Handle, IMAGE_ICON, lpRect, fAsDisabled, crColor)

    End Select

End Function

'' Render an Image

Public Function Render_Image(ByVal hDC As Long, ByVal hImage As Long, _
                              ByVal ImageType As Long, _
                              lpRect As RECT, _
                              Optional ByVal AsDisabled As Boolean, _
                              Optional ByVal crBackColor As Long = -1&) As Long

    Dim crColor As Long, _
        lpBrush As LOGBRUSH, _
        hBrush As Long
        
    Dim cx As Long, _
        cy As Long
        
    Dim cbFlags As Long
        
    Dim hPicNew As Long
        
        
        With lpRect
            cx = (.Right - .Left) + 1
            cy = (.Bottom - .Top) + 1
                        
            Select Case ImageType
                Case IMAGE_BITMAP
                    cbFlags = DST_BITMAP
                    
                Case IMAGE_ICON, IMAGE_CURSOR
                    cbFlags = DST_ICON
                    
            End Select
                                
            If (AsDisabled = True) Then
                hPicNew = GetSmallPicture(hImage, ImageType, cx, cy, &HFFFFFF)
            Else
                hPicNew = GetSmallPicture(hImage, ImageType, cx, cy, crBackColor)
            End If
            
            cbFlags = (cbFlags Or (DSS_DISABLED And AsDisabled))
            
            DrawState hDC, 0&, 0&, hPicNew, 0&, .Left, .Top, cx, cy, cbFlags
        
            If ImageType = IMAGE_BITMAP Then
                DeleteObject hPicNew
            Else
                DestroyIcon hPicNew
            End If
            
        
        End With
        
End Function


'''' Text Rendering / Measuring Functions


'' Unicode Text Measure

Public Function MeasureText_API(ByVal hDC As Long, ByVal lpszText As String, _
                                lpSize As SIZEAPI) As Long
        
    Dim cMeasure As SIZEAPI
        
    Dim lSize As Long, _
        lpPtr As Long
                
    Dim sStr As String
                
    If (Unicode_Checked = False) Then CheckUnicode
    
    sStr = lpszText
    lSize = Len(sStr)
    
    If (Using_Unicode = True) Then
        lpPtr = StrPtr(sStr)
        
        GetTextExtentPoint32W hDC, lpPtr, lSize, cMeasure
    Else
        sStr = StrConv(sStr, vbFromUnicode)
        lpPtr = StrPtr(sStr)
        
        GetTextExtentPoint32 hDC, lpPtr, lSize, cMeasure
    End If

    lpSize.cx = cMeasure.cx
    lpSize.cy = cMeasure.cy
        
    If (lSize <> 0&) Then
        MeasureText_API = True
    End If
    
End Function


'' Unicode Char Widths

Public Function GetCharWidth_API(ByVal hDC As Long, ByVal CharFirst As Long, ByVal CharLast As Long, _
                                 Widths() As Long) As Long
        
    Dim b As Boolean, _
        wIn() As Long
        
    Dim i As Long
    
    i = (CharLast - CharFirst)
    ReDim wIn(0 To i)
                
    If (Unicode_Checked = False) Then CheckUnicode
    
    If (Using_Unicode = True) Then
        b = GetCharWidth32W(hDC, CharFirst, CharLast, wIn(0))
    Else
        b = GetCharWidth32(hDC, CharFirst, CharLast, wIn(0))
    End If

    Widths = wIn
    GetCharWidth_API = b
    
End Function

'' Unicode Text Write Functions

'' TextOut
Public Function TextOut_API(ByVal hDC As Long, ByVal lpszText As String, _
                            ByVal x As Long, ByVal y As Long) As Long
        
    Dim lSize As Long, _
        lpPtr As Long, _
        n As Long
               
    Dim szText As String
                
    If (Unicode_Checked = False) Then CheckUnicode
        
    lSize = Len(lpszText)
    
    If (Using_Unicode = True) Then
        lpPtr = StrPtr(lpszText)
        TextOutW hDC, x, y, lpPtr, lSize
    Else
        szText = StrConv(lpszText, vbFromUnicode)
        lpPtr = StrPtr(szText)
        
        TextOut hDC, x, y, lpPtr, lSize
    End If
    
End Function

'' Draw-Disabled
Public Function DrawDisabled_API(ByVal hDC As Long, ByVal lpszText As String, _
                                 ByVal x As Long, ByVal y As Long) As Long
        
    Dim lSize As Long, _
        lpPtr As Long
    
    Dim szText As String
    
    lpPtr = 0&
    
    If (Unicode_Checked = False) Then CheckUnicode
            
    lSize = Len(lpszText)
    
    If (lSize = 0&) Then
        Exit Function
    End If
    
    If (Using_Unicode = True) Then
        lpPtr = StrPtr(lpszText)
    
        DrawStateW hDC, 0&, 0&, lpPtr, lSize, x, y, 0&, 0&, (DST_TEXT Or DSS_DISABLED)
    Else
        szText = StrConv(lpszText, vbFromUnicode)
        lpPtr = StrPtr(szText)
        
        DrawState hDC, 0&, 0&, lpPtr, lSize, x, y, 0&, 0&, (DST_TEXT Or DSS_DISABLED)
    End If
    
End Function
                            
'' DrawText
Public Function DrawText_API(ByVal hDC As Long, ByVal lpszText As String, _
                             lpRect As RECT) As Long

    Dim lSize As Long, _
        lpPtr As Long
                
    Dim szText As String
                
    If (Unicode_Checked = False) Then CheckUnicode
    
    lSize = Len(lpszText)
    
    If (Using_Unicode = True) Then
        lpPtr = StrPtr(lpszText)
        DrawTextW hDC, lpPtr, lSize, lpRect, 0&
    Else
        szText = StrConv(lpszText, vbFromUnicode)
        lpPtr = StrPtr(szText)
        
        DrawText hDC, lpPtr, lSize, lpRect, 0&
    End If
    
End Function


'''' Code for manipulating colors and bitmaps

'' Turn a color into the Red, Green and Blue constituents
'' in separate variables

Public Sub GetRGB(ByVal Color As Long, Red As Byte, Green As Byte, Blue As Byte)
    Dim c As Long
    
    c = (Color And &HFF&)
    Red = CByte(c)
    
    c = ((Color And &HFF00&) / &H100&)
    Green = CByte(c)
    
    c = ((Color And &HFF0000) / &H10000)
    Blue = CByte(c)
    
End Sub

'' Batch convert ColorRefs into their constituents
'' saved as BITDATA

Public Sub AClrBit(Colors() As Long, Bits() As BITDATA)
    Dim i As Long, _
        j As Long
        
    On Error Resume Next
    i = -1&
    i = UBound(Colors)
    If i = -1& Then Exit Sub
    
    ReDim Bits(0 To i)
    For j = 0 To i
        ClrBit Colors(j), Bits(j)
    Next j
End Sub

'' vice-versa

Public Sub ABitClr(Bits() As BITDATA, Colors() As Long)
    Dim i As Long, _
        j As Long
        
    On Error Resume Next
    i = -1&
    i = UBound(Bits)
    If i = -1& Then Exit Sub
    
    ReDim Colors(0 To i)
    For j = 0 To i
        Colors(j) = BitClr(Bits(j))
    Next j
    
End Sub

'' Single Convert ColorRef to RGB

Public Sub ClrBit(ByVal Color As Long, Bits As BITDATA)
    GetRGB Color, Bits.Red, Bits.Green, Bits.Blue
End Sub

'' Single Convert RGB to ColorRef

Public Function BitClr(Bits As BITDATA) As Long
    BitClr = RGB(Bits.Red, Bits.Green, Bits.Blue)
End Function

'' Make a gradient using a device
Public Function ColorGrad(ByVal Color1 As Long, ByVal Color2 As Long, ByVal Steps As Long) As Long()

    Dim b As BITDATA, _
        n As BITDATA, _
        o As BITDATA

    Dim factR As Single, _
        factG As Single, _
        factB As Single
        
    Dim i As Long, _
        j As Long
        
    Dim cGrad() As Long
        
    j = (Steps - 1)
    ReDim cGrad(0 To j)
    
    ClrBit Color1, n
    ClrBit Color2, o
    
    factR = ((CLng(o.Red) - CLng(n.Red)) / (Steps - 1))
    factG = ((CLng(o.Green) - CLng(n.Green)) / (Steps - 1))
    factB = ((CLng(o.Blue) - CLng(n.Blue)) / (Steps - 1))
        
    For i = 0 To j
    
        b.Red = n.Red + (factR * i)
        b.Green = n.Green + (factG * i)
        b.Blue = n.Blue + (factB * i)
        
        With b
            cGrad(i) = (.Red) + (.Green * &H100&) + (.Blue * &H10000)
        End With
    Next i
    
    ColorGrad = cGrad
    
End Function

'' Expand a palette of colors
Public Function ExpandPalette(Palette() As Long, ByVal Steps As Long) As Long()

    Dim i As Long, _
        j As Long
    
    Dim x As Long, _
        y As Long
          
    Dim b As BITDATA, _
        s As Single
                
    Dim n As Long, _
        f As Long
    
    Dim cNew() As Long, _
        cCopy() As Long
    
    On Error Resume Next
    
    i = -1&
    i = UBound(Palette)
    If (i = -1&) Then
        Exit Function
    ElseIf ((i + 1) = Steps) Then
        ExpandPalette = Palette
        Exit Function
    ElseIf ((i + 1) > Steps) Then
        'ExpandPalette = ContractPalette(Palette, Steps)
        Exit Function
    End If
                        
    j = (Steps - 1)
    ReDim cNew(0 To j)
    
    s = (Steps / (i + 1))
    i = i - 1
    
    y = Round(s) + 1
    
    For f = 0 To i
    
        If (f = i) Then y = (j - n) + 1
    
        cCopy = ColorGrad(Palette(f), Palette(f + 1), y)
        CopyMemory cNew(n), cCopy(0), Len(cCopy(0)) * y
        n = n + (y - 1)
        
        Erase cCopy
        
    Next f
    
    ExpandPalette = cNew
    
End Function

'' Fill an area with a gradient

Public Function GradientFill(ByVal hDC As Long, ByVal crColor1 As Long, ByVal crColor2 As Long, lpRect As RECT, Optional ByVal fVertical As Boolean)

    Dim lpPoint As POINTAPI, _
        rcFill As RECT
        
    Dim lpBrush As LOGBRUSH, _
        hBrush As Long
        
    Dim i As Long, _
        j As Long
        
    Dim dx As Long, _
        dy As Long
        
    Dim Colors() As Long
    
    i = GetActualColor(crColor1)
    j = GetActualColor(crColor2)
    
    With lpRect
        dx = (.Right - .Left) + 1
        dy = (.Bottom - .Top) + 1
    End With
    
    CopyMemory rcFill, lpRect, 16&
    
    lpBrush.lbStyle = BS_SOLID
    If (fVertical = True) Then
        Colors = ColorGrad(i, j, dy)
    Else
        Colors = ColorGrad(i, j, dx)
    End If

    j = -1&
    j = UBound(Colors)
    
    For i = 0& To j
        
        lpBrush.lbColor = Colors(i)
        hBrush = CreateBrushIndirect(lpBrush)
        
        With rcFill
            If (fVertical = True) Then
                .Top = (lpRect.Top + i)
                .Bottom = .Top + 1&
            Else
                .Left = (lpRect.Left + i)
                .Right = .Left + 1&
            End If
        End With
        
        FillRect hDC, rcFill, hBrush
        DeleteObject hBrush
        
    Next i

    Erase Colors

End Function

Public Function GetAverageColor(ByVal Color1 As Long, ByVal Color2 As Single) As Long
    Dim Bits(0 To 2) As BITDATA, _
        df As Single
    
    ClrBit Color1, Bits(0)
    ClrBit Color2, Bits(1)
    
    If (Round(Color2) <> Color2) Then
    
        df = (CSng(Bits(0).Red) * Color2)
        Bits(2).Red = Round(df)
    
        df = (CSng(Bits(0).Green) * Color2)
        Bits(2).Green = Round(df)
    
        df = (CSng(Bits(0).Blue) * Color2)
        Bits(2).Blue = Round(df)
    
    Else
        df = (CSng(Bits(0).Red) + CSng(Bits(1).Red)) / 2&
        Bits(2).Red = Round(df)
    
        df = (CSng(Bits(0).Green) + CSng(Bits(1).Green)) / 2&
        Bits(2).Green = Round(df)
    
        df = (CSng(Bits(0).Blue) + CSng(Bits(1).Blue)) / 2&
        Bits(2).Blue = Round(df)
    End If

    GetAverageColor = BitClr(Bits(2))

End Function

Private Function GetPlainCaption(ByVal vData As String) As String
    Dim i As Long, _
        j As Long
        
    Dim s1 As String, _
        s2 As String
        
    i = Len(vData)
    
    j = InStr(1, vData, "&")
    
    s1 = Mid(vData, 1, j - 1)
    s2 = Mid(vData, j + 1)
    
    GetPlainCaption = s1 + s2
    
End Function

''' Copyright (C) 2001, 2002 Nathan Moschkin

''' ****************** NOT FOR COMMERCIAL USE *****************
''' Inquire if you would like to use this code commercially.
''' Unauthorized recompilation and/or re-release for commercial
''' use is strictly prohibited.
'''
''' please send changes made to code to me at the address, below,
''' if you plan on making those changes publicly available.

''' e-mail questions or comments to nmosch@tampabay.rr.com



