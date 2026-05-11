Attribute VB_Name = "PowerReSize"
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, iPic As StdPicture) As Long
Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type RGBtype
    B As Byte
    R As Byte
    G As Byte
End Type

Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Const DIB_RGB_COLORS = 0&
Public Const BI_RGB = 0&

Type BITMAPINFOHEADER
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type

'*****************************************************
'PowerResize
' Returns a resized version of Img with the new dimensions passed.
'
'Written by Mark Gordon aka msg555
'1/16/06
'Free to use/sell/whatever
'*****************************************************
Public Function PowerResize(Img As StdPicture, newWidth As Long, newHeight As Long) As StdPicture
    Debug.Assert Img.Type = vbPicTypeBitmap 'Image must be a bitmap
        
    Dim SrcBmp As BITMAP
    GetObject Img.handle, Len(SrcBmp), SrcBmp
        
    Dim srcBI As BITMAPINFO
    With srcBI.bmiHeader
        .biSize = Len(srcBI.bmiHeader)
        .biWidth = SrcBmp.bmWidth
        .biHeight = -SrcBmp.bmHeight
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
    End With

    'Create Source Bit Array
    Dim SrcBits() As RGBQUAD
    ReDim SrcBits(0 To SrcBmp.bmWidth - 1, 0 To SrcBmp.bmHeight - 1) As RGBQUAD

    'Grab Source Bits
    Dim lDc As Long
    lDc = CreateCompatibleDC(0)
    GetDIBits lDc, Img.handle, 0, SrcBmp.bmHeight, SrcBits(0, 0), srcBI, DIB_RGB_COLORS
    DeleteDC lDc

    'Create Destination Bit Array
    Dim DblDstBits() As Double
    ReDim DblDstBits(0 To 3, 0 To newWidth - 1, 0 To newHeight - 1) As Double

    'Multipliers
    Dim xMult As Double, yMult As Double
    xMult = newWidth / SrcBmp.bmWidth
    yMult = newHeight / SrcBmp.bmHeight

    'Traversing variables
    Dim X As Long, XX As Long
    Dim Y As Long, YY As Long
    
    'Low/High scan X/Y
    Dim lsX As Double, hsX As Double
    Dim lsY As Double, hsY As Double
    
    Dim OverlapWidth As Double
    Dim OverlapHeight As Double
    Dim Overlap As Double
    
    For X = 0 To SrcBmp.bmWidth - 1
        lsX = X * xMult
        hsX = X * xMult + xMult
        For Y = 0 To SrcBmp.bmHeight - 1
            lsY = Y * yMult
            hsY = Y * yMult + yMult
            For XX = Fix(lsX) To IIf(Fix(hsX) = hsX, Fix(hsX), Fix(hsX + 1)) - 1
                For YY = Fix(lsY) To IIf(Fix(hsY) = hsY, Fix(hsY), Fix(hsY + 1)) - 1
                    OverlapWidth = 1
                    OverlapHeight = 1
                    
                    If XX < lsX Then OverlapWidth = 1# - (lsX - XX)
                    If XX + 1# > hsX Then OverlapWidth = OverlapWidth - (XX + 1# - hsX)
                    If YY < lsY Then OverlapHeight = 1# - (lsY - YY)
                    If YY + 1# > hsY Then OverlapHeight = OverlapHeight - (YY + 1# - hsY)
                    
                    Overlap = OverlapHeight * OverlapWidth
                    
                    DblDstBits(0, XX, YY) = DblDstBits(0, XX, YY) + SrcBits(X, Y).rgbRed * Overlap
                    DblDstBits(1, XX, YY) = DblDstBits(1, XX, YY) + SrcBits(X, Y).rgbGreen * Overlap
                    DblDstBits(2, XX, YY) = DblDstBits(2, XX, YY) + SrcBits(X, Y).rgbBlue * Overlap
                    DblDstBits(3, XX, YY) = DblDstBits(3, XX, YY) + Overlap
                Next
            Next
        Next
    Next
    
    Dim DstBits() As RGBQUAD
    ReDim DstBits(0 To newWidth - 1, 0 To newHeight - 1) As RGBQUAD
    
    For X = 0 To newWidth - 1
        For Y = 0 To newHeight - 1
            DstBits(X, Y).rgbRed = Round(DblDstBits(0, X, Y) / DblDstBits(3, X, Y))
            DstBits(X, Y).rgbGreen = Round(DblDstBits(1, X, Y) / DblDstBits(3, X, Y))
            DstBits(X, Y).rgbBlue = Round(DblDstBits(2, X, Y) / DblDstBits(3, X, Y))
        Next
    Next
    
    Dim dstBI As BITMAPINFO
    With dstBI.bmiHeader
        .biSize = Len(dstBI.bmiHeader)
        .biWidth = newWidth
        .biHeight = -newHeight
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
    End With
    
    Dim hBmp As Long
    hBmp = CreateBitmap(newWidth, newHeight, 1, 32, ByVal 0)

    SetDIBits 0, hBmp, 0, newHeight, DstBits(0, 0), dstBI, DIB_RGB_COLORS

    Dim IGuid As Guid
    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    Dim PicDst As PictDesc
    With PicDst
        .cbSizeofStruct = Len(PicDst)
        .hImage = hBmp
        .picType = vbPicTypeBitmap
    End With
    
    OleCreatePictureIndirect PicDst, IGuid, True, PowerResize
End Function

