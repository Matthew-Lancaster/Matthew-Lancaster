Attribute VB_Name = "modDIBTools"
Option Explicit

Public Sub DIBDims(ByVal Width As Long, ByVal Height As Long, lpInfo As BITMAPINFO)

    Dim i As Long

    With lpInfo.bmiHeader
        .biWidth = Width
        .biHeight = Height
        
        i = ModStep(Height, 4&)
        i = i * (ModStep(Width * 3&, 4&))
                
        .biSizeImage = i
    End With

End Sub

Public Sub DIBInfo(ByVal hBitmap As Long, lpInfo As BITMAPINFO)
    Dim lpBitmap As BITMAP, _
        cx As Long, _
        cy As Long
    
    GetObject hBitmap, Len(lpBitmap), lpBitmap
    
    cx = lpBitmap.bmWidth
    cy = lpBitmap.bmHeight
    
    With lpInfo.bmiHeader
        .biSize = Len(lpInfo) - 4&
        .biBitCount = 24&
        .biPlanes = 1&
        
        DIBDims cx, cy, lpInfo
    End With
    
End Sub

Public Sub GetBits(ByVal aHDC As Long, ByVal hBitmap As Long, lpBits() As Byte)
    Dim lpInfo As BITMAPINFO, _
        c As Long, _
        cy As Long
    
    DIBInfo hBitmap, lpInfo
    
    cy = lpInfo.bmiHeader.biHeight
    c = lpInfo.bmiHeader.biSizeImage - 1&
    
    ReDim Preserve lpBits(0 To c)
    
    GetDIBits aHDC, hBitmap, 0&, cy, lpBits(0&), lpInfo, 0&
    
End Sub

Public Sub SetBits(ByVal aHDC As Long, ByVal hBitmap As Long, lpBits() As Byte)
    Dim lpInfo As BITMAPINFO, _
        c As Long, _
        cy As Long
    
    DIBInfo hBitmap, lpInfo
    
    cy = lpInfo.bmiHeader.biHeight
    c = lpInfo.bmiHeader.biSizeImage - 1&
    
    SetDIBits aHDC, hBitmap, 0&, cy, lpBits(0&), lpInfo, 0&
    
End Sub

Public Sub GetScanLine(lpSection As DIBDATA, ByVal nLine As Long, lpLine() As RGBDATA)

    Dim lpBytes() As Byte, _
        i As Long, _
        j As Long
        
    j = lpSection.Info.bmiHeader.biWidth
    ReDim Preserve lpLine(0 To (j - 1))
    
    i = ModStep(j * 3, 4)
        
    j = (lpSection.Info.bmiHeader.biHeight - 1&) - nLine

    j = j * i
    CopyMemory lpLine(0), lpSection.Bytes(j), i
    
End Sub

Public Sub SetScanLine(lpSection As DIBDATA, ByVal nLine As Long, lpLine() As RGBDATA)

    Dim lpBytes() As Byte, _
        i As Long, _
        j As Long
        
    j = lpSection.Info.bmiHeader.biWidth
    i = ModStep(j * 3, 4)
        
    j = (lpSection.Info.bmiHeader.biHeight - 1&) - nLine

    j = j * i
    CopyMemory lpSection.Bytes(j), lpLine(0), i
    
End Sub

Public Sub GetLine(lpSection As DIBDATA, ByVal nLine As Long, lpLine() As RGBDATA)
    
    Dim lpBytes() As Byte, _
        i As Long, _
        j As Long
        
    j = lpSection.Info.bmiHeader.biWidth
    ReDim Preserve lpLine(0 To (j - 1))
    
    i = ModStep(j * 3, 4) - 1
    ReDim lpBytes(0 To i)
        
    i = (lpSection.Info.bmiHeader.biHeight - 1&) - nLine
    
    GetDIBits lpSection.hDC, lpSection.hBitmap, i, 1&, lpBytes(0), lpSection.Info, 0&
    
    j = (j * 3&)
    
    CopyMemory lpLine(0), lpBytes(0), j
    Erase lpBytes
    
End Sub

Public Sub SetLine(lpSection As DIBDATA, ByVal nLine As Long, lpLine() As RGBDATA)
    
    Dim lpBytes() As Byte, _
        i As Long, _
        j As Long
        
    j = lpSection.Info.bmiHeader.biWidth
    
    i = ModStep(j * 3, 4) - 1
    ReDim lpBytes(0 To i)
        
    j = (j * 3&)
    CopyMemory lpBytes(0), lpLine(0), j
    
    i = (lpSection.Info.bmiHeader.biHeight - 1&) - nLine
    
    SetDIBits lpSection.hDC, lpSection.hBitmap, i, 1&, lpBytes(0), lpSection.Info, 0&
    Erase lpBytes
    
End Sub


Public Sub Darken(ByVal hDC As Long, lpRect As RECT, ByVal Value As Double, Optional ByVal BackColor As Long = -1&, Optional ByVal MaskColor As Long = &HC0C0C)

    Dim lpLine() As RGBDATA, _
        lpDIB As DIBDATA
        
    Dim hBitmap As Long, _
        hBmpTemp As Long
        
    Dim x As Long, _
        y As Long
        
    GdiFlush
        
    hBmpTemp = CreateCompatibleBitmap(hDC, 1, 1)
    hBitmap = SelectObject(hDC, hBmpTemp)
    
    With lpDIB
        .hDC = hDC
        .hBitmap = hBitmap
    End With
    
    DIBInfo hBitmap, lpDIB.Info
    
    GetBits hDC, hBitmap, lpDIB.Bytes
        
    With lpRect
        For y = .Top To .Bottom
            GetScanLine lpDIB, y, lpLine
            
            For x = .Left To .Right
                If (BackColor <> -1&) Then
                    If (RGBToColor(lpLine(x)) = MaskColor) Then
                        ColorToRGB BackColor, lpLine(x)
                    Else
                        SetTone lpLine(x), Value
                    End If
                Else
                    SetTone lpLine(x), Value
                End If
            Next x
            
            SetScanLine lpDIB, y, lpLine
        Next y
        
    End With
    
    SetBits hDC, hBitmap, lpDIB.Bytes
    Erase lpDIB.Bytes
    
    SelectObject hDC, hBitmap
    DeleteObject hBmpTemp

End Sub

Public Sub Unmask(ByVal hDC As Long, lpRect As RECT, ByVal BackColor As Long, Optional ByVal MaskColor As Long = &HC0C0C0)

    Dim lpLine() As RGBDATA, _
        lpDIB As DIBDATA
        
    Dim hBitmap As Long, _
        hBmpTemp As Long
        
    Dim x As Long, _
        y As Long
        
    GdiFlush
        
    hBmpTemp = CreateCompatibleBitmap(hDC, 1, 1)
    hBitmap = SelectObject(hDC, hBmpTemp)
    
    With lpDIB
        .hDC = hDC
        .hBitmap = hBitmap
    End With
    
    DIBInfo hBitmap, lpDIB.Info
    GetBits hDC, hBitmap, lpDIB.Bytes
        
    With lpRect
        For y = .Top To .Bottom
            GetScanLine lpDIB, y, lpLine
            
            For x = .Left To .Right
                If (RGBToColor(lpLine(x)) = MaskColor) Then
                    ColorToRGB BackColor, lpLine(x)
                End If
            Next x
            
            SetScanLine lpDIB, y, lpLine
        Next y
        
    End With
    
    SetBits hDC, hBitmap, lpDIB.Bytes
    Erase lpDIB.Bytes
    
    SelectObject hDC, hBitmap
    DeleteObject hBmpTemp

End Sub

Public Sub UnmaskGrayTones(ByVal hDC As Long, lpRect As RECT, ByVal Color As Long, Optional ByVal Exclude As Long = -1&)

    Dim lpLine() As RGBDATA, _
        lpDIB As DIBDATA
        
    Dim hBitmap As Long, _
        hBmpTemp As Long
        
    Dim x As Long, _
        y As Long
        
    Dim n As Long
        
    GdiFlush
        
    hBmpTemp = CreateCompatibleBitmap(hDC, 1, 1)
    hBitmap = SelectObject(hDC, hBmpTemp)
    
    With lpDIB
        .hDC = hDC
        .hBitmap = hBitmap
    End With
    
    DIBInfo hBitmap, lpDIB.Info
    GetBits hDC, hBitmap, lpDIB.Bytes
        
    With lpRect
        For y = .Top To .Bottom
            GetScanLine lpDIB, y, lpLine
            
            For x = .Left To .Right
                n = RGBToColor(lpLine(x))
                
                If (Exclude <> -1&) Then
                    If (n <> Exclude) Then
                        SetTone lpLine(x), ((n And &HFF&) / 255), Color
                    End If
                Else
                    
                    SetTone lpLine(x), ((n And &HFF&) / 255), Color
                End If
                
            Next x
            
            SetScanLine lpDIB, y, lpLine
        Next y
        
    End With
    
    SetBits hDC, hBitmap, lpDIB.Bytes
    Erase lpDIB.Bytes
    
    SelectObject hDC, hBitmap
    DeleteObject hBmpTemp

End Sub




