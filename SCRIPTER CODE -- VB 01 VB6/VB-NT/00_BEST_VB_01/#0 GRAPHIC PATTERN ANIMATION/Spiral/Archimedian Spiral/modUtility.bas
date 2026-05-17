Attribute VB_Name = "modUtility"
''' GUInerd Standard Menu System
''' Version 4.0

'''

''' Common Miscellaneous Utility Routines


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit





Public Function GetOlePicture(ByVal hImage As Long, ByVal ImageType As Long, Optional ByVal Own As Boolean) As StdPicture

    Dim varPicData As PICTDESCBMP, _
        varIconData As PICTDESCICON
        
    Dim varIPicture As StdPicture
    Dim varOLEID As GUID
    Dim varhBmp As Long
    
    With varOLEID
    
' {7BF80981-BF32-101A-8BBB-00AA00300CAB}
        
        .Data1 = &H7BF80981
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0&) = &H8B
        .Data4(1) = &HBB
        .Data4(3) = &HAA
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
        
    End With
    
    If (ImageType = IMAGE_BITMAP) Then
    
        varPicData.SizeofStruct = Len(varPicData)
        varPicData.hBitmap = hImage
        varPicData.PicType = 1
            
        OleCreatePictureIndirect VarPtr(varPicData), varOLEID, Own, varIPicture
    
    ElseIf (ImageType = IMAGE_ICON) Or (ImageType = IMAGE_CURSOR) Then
    
        varIconData.SizeofStruct = Len(varIconData)
        varIconData.hIcon = hImage
        varIconData.PicType = 3
            
        OleCreatePictureIndirect ByVal VarPtr(varIconData), varOLEID, True, varIPicture
    End If
        
    On Error Resume Next
    
    If Not varIPicture Is Nothing Then
        Set GetOlePicture = varIPicture
    End If

End Function

'' Copy one bitmap to another, changing the bitplanes from 24 to 1.

Public Function GetMonoPicture(ByVal hImage As Long, ByVal dwType As Long, _
                               Optional ByVal cx As Long = 16&, _
                               Optional ByVal cy As Long = 16&) As Long
    Dim hImageNew As Long
        
    If dwType = IMAGE_ICON Then
        
        hImageNew = CopyImage(hImage, IMAGE_ICON, cx, cy, &H41)
        
    Else
    
        hImageNew = CopyImage(hImage, IMAGE_BITMAP, cx, cy, &H41)
    
    End If
    
    GetMonoPicture = hImageNew
    
End Function

'' Function to get a 16x16 representation of the picture

Public Function GetSmallPicture(ByVal hImage As Long, ByVal dwType As Long, Optional ByVal cx As Long = 16&, Optional ByVal cy As Long = 16&, Optional ByVal crUnMask As Long = -1&) As Long
    Dim hDC As Long, _
        nDC As Long
    
    Dim rcPicture As RECT
    
    GetSmallPicture = CopyImage(hImage, dwType, cx, cy, LR_COPYFROMRESOURCE)
    
    If (crUnMask <> -1&) Then
        If (dwType = IMAGE_BITMAP) Then
            hDC = GetDC(0&)
            nDC = CreateCompatibleDC(hDC)
            ReleaseDC 0&, hDC
            
            rcPicture.Right = cx
            rcPicture.Bottom = cy
            
            SelectObject nDC, GetSmallPicture
            Unmask nDC, rcPicture, &HC0C0C0, crUnMask
            DeleteDC nDC
        End If
    End If
    
End Function

Public Function MakePoints(ByVal lParam As Long) As POINTAPI

    Dim p As POINTAPI
    
    p.x = (lParam And &HFFFF&)
    p.y = lParam / &H10000
    
    MakePoints = p
    
End Function

Public Function CopyToBitmap(ByVal hImage As Long, ByVal ImageType As Integer) As StdPicture
    Dim newDC As Long, _
        newDC2 As Long, _
        newBitmap As Long, _
        dDC As Long
        
    Dim lpBitmap As BITMAP, _
        lpIcon As ICONINFO
        
    Dim iWidth As Long, _
        iHeight As Long
    
    dDC = GetDC(0&)
    newDC = CreateCompatibleDC(dDC)
    
    If (ImageType = 1&) Then        ' Bitmap
        GetObject hImage, Len(lpBitmap), lpBitmap
        
        iWidth = lpBitmap.bmWidth
        iHeight = lpBitmap.bmHeight
    Else                            ' Icon
        GetIconInfo hImage, lpIcon
        GetObject lpIcon.hbmColor, Len(lpBitmap), lpBitmap
    End If
    
    iWidth = lpBitmap.bmWidth
    iHeight = lpBitmap.bmHeight
    
    newBitmap = CreateCompatibleBitmap(dDC, iWidth, iHeight)
    
    SelectObject newDC, newBitmap
    
    If (ImageType = 1&) Then
        newDC2 = CreateCompatibleDC(dDC)
        SelectObject newDC2, hImage
        
        BitBlt newDC, 0&, 0&, iWidth, iHeight, newDC2, 0&, 0&, SRCCOPY
        
        DeleteDC newDC2
    Else
    
        DrawIconEx newDC, 0&, 0&, hImage, iWidth, iHeight, 0&, 0&, DI_NORMAL
    
    End If
    
    DeleteDC newDC
    Set CopyToBitmap = GetOlePicture(newBitmap, IMAGE_BITMAP, True)
    
End Function

Public Function Unmask(ByVal hDC As Long, lpRect As RECT, ByVal MaskColor As Long, ByVal BackColor As Long)

    Dim i As Long, _
        j As Long
        
    Dim cx As Long, _
        cy As Long
        
    Dim x As Long, _
        y As Long
        
    Dim nPixel As Long
        
    x = lpRect.Left
    cx = lpRect.Right - x
    
    y = lpRect.Top
    cy = lpRect.Bottom - y
        
    For i = x To (x + cx)
        For j = y To (y + cy)
        
            nPixel = GetPixel(hDC, i, j)
            If (nPixel = MaskColor) Then
                SetPixel hDC, i, j, BackColor
            End If
        
        Next j
    Next i

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



