VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Base 0

'Class Name:    cImage.cls
'
'Description:   This class creates and gives access to a DIBSection for the
'               purpose of displaying and editing a digital image.
'

Private Type SAFEARRAYBOUND
    cElements         As Long
    lLbound           As Long
End Type
Private Type SAFEARRAY2D
    cDims             As Integer
    fFeatures         As Integer
    cbElements        As Long
    cLocks            As Long
    pvData            As Long
    Bounds(0 To 1)    As SAFEARRAYBOUND
End Type
Private Type RGBQUAD
    rgbBlue           As Byte
    rgbGreen          As Byte
    rgbRed            As Byte
    rgbReserved       As Byte
End Type
Private Type BITMAPINFOHEADER
    biSize            As Long
    biWidth           As Long
    biHeight          As Long
    biPlanes          As Integer
    biBitCount        As Integer
    biCompression     As Long
    biSizeImage       As Long
    biXPelsPerMeter   As Long
    biYPelsPerMeter   As Long
    biClrUsed         As Long
    biClrImportant    As Long
End Type
Private Type BITMAPFILEHEADER
    bfType            As Integer
    bfSize            As Long
    bfReserved1       As Integer
    bfReserved2       As Integer
    bfOffBits         As Long
End Type
Private Type BITMAP
    bmType            As Long
    bmWidth           As Long
    bmHeight          As Long
    bmWidthBytes      As Long
    bmPlanes          As Integer
    bmBitsPixel       As Integer
    bmBits            As Long
End Type
Private Type BITMAPINFO
    bmiHeader         As BITMAPINFOHEADER
    bmiColors(255)    As RGBQUAD
End Type
Private Type POINTAPI
        x As Long
        y As Long
End Type


Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function CreateDIBSection2 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long   'lplpVoid changed to ByRef
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, pRGBQuad As RGBQUAD) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long
Private Declare Function GetDIBits256 Lib "gdi32" Alias "GetDIBits" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)


Private Const BLACKONWHITE    As Long = 1 'nStretchMode constants for
Private Const COLORONCOLOR    As Long = 3 '  SetStretchBltMode() API function
Private Const HALFTONE        As Long = 4 'HALFTONE not supported in Win 95, 98, ME

Private Const BI_RGB          As Long = 0&
Private Const BI_RLE4         As Long = 2&
Private Const BI_RLE8         As Long = 1&
Private Const DIB_RGB_COLORS  As Long = 0

Private m_hDIb        As Long       ' Handle to the current DIBSection
Private m_hBmpOld     As Long       ' Handle to the old bitmap in the DC, for clear up
Private m_hDC         As Long       ' Handle to the Device context holding the DIBSection
Private m_Ptr         As Long       ' Address of memory pointing to the DIBSection's bits
Private m_BI          As BITMAPINFO ' Type containing the Bitmap information
Private m_RGB(255)    As RGBQUAD



Private Sub Clear()
    If (m_hDC <> 0) Then
        If (m_hDIb <> 0) Then
            SelectObject m_hDC, m_hBmpOld
            DeleteObject m_hDIb
        End If
        DeleteObject m_hDC
    End If
    m_hDC = 0
    m_hDIb = 0
    m_hBmpOld = 0
    m_Ptr = 0
End Sub
Private Sub Class_Terminate()
    Clear
End Sub



'====================================================================================
'                                PUBLIC PROPERTIES
'====================================================================================
Public Property Get Width() As Long
    Width = m_BI.bmiHeader.biWidth
End Property
Public Property Get Height() As Long
    Height = m_BI.bmiHeader.biHeight
End Property
Public Property Get BitCount() As Integer
    BitCount = m_BI.bmiHeader.biBitCount
End Property
Public Property Get hDC() As Long
    hDC = m_hDC
End Property
Public Property Get DIBitsPtr() As Long
    DIBitsPtr = m_Ptr
End Property
Public Property Get BytesPerScanLine() As Long
    Select Case m_BI.bmiHeader.biBitCount ' Scans must align on 4-byte boundaries
    Case 1:    BytesPerScanLine = ((m_BI.bmiHeader.biWidth - 1) \ 8 + 4) And &HFFFFFFFC
    Case 4:    BytesPerScanLine = ((m_BI.bmiHeader.biWidth - 1) \ 2 + 4) And &HFFFFFFFC
    Case 8:    BytesPerScanLine = (m_BI.bmiHeader.biWidth + 3) And &HFFFFFFFC
    Case Else: BytesPerScanLine = (m_BI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
    End Select
End Property



'====================================================================================
'                             DIMENSION / COLOR DEPTH
'====================================================================================
Public Function Create(lWidth As Long, lHeight As Long, iBitCount As Integer) As Boolean
    Clear                        'Set Dimensions and BitCount in this cImage
    Select Case iBitCount
    Case 24
        m_hDC = CreateCompatibleDC(0)
    Case 1, 4, 8
        Dim lHDCDesk As Long
        lHDCDesk = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        m_hDC = CreateCompatibleDC(lHDCDesk)
        DeleteDC lHDCDesk
    End Select
    If m_hDC = 0 Then Exit Function
    With m_BI.bmiHeader
        .biSize = Len(m_BI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = iBitCount
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine * .biHeight
    End With
    If iBitCount <> 24 Then ' Create a default grayscale palette
        Dim i As Long
        Dim c As Long
        c = 2 ^ iBitCount - 1
        For i = 0 To c
            With m_BI.bmiColors(i)
                .rgbBlue = i * 255# / c
                .rgbGreen = .rgbBlue
                .rgbRed = .rgbBlue
            End With
        Next i
    End If
    m_hDIb = CreateDIBSection2(m_hDC, m_BI, DIB_RGB_COLORS, m_Ptr, 0, 0)
    If m_hDIb = 0 Then
        DeleteObject m_hDC
    Else
        m_hBmpOld = SelectObject(m_hDC, m_hDIb)
        Create = True
    End If
End Function



'====================================================================================
'                                 LOAD/COPY IMAGE
'====================================================================================
Public Function CopyStdPicture(ByRef TheStdPicture As StdPicture, Optional iBitCount As Integer) As Boolean
    Dim lHDC         As Long
    Dim lhDCDesktop  As Long
    Dim lhBmpOld     As Long
    Dim tBMP         As BITMAP
    Dim CopyPalette  As Boolean

    GetObjectAPI TheStdPicture.handle, Len(tBMP), tBMP

    CopyPalette = (iBitCount = 0)
    If CopyPalette Then
        iBitCount = tBMP.bmBitsPixel
        If iBitCount = 16 Then iBitCount = 24
    End If

    If Not Create(tBMP.bmWidth, tBMP.bmHeight, iBitCount) Then Exit Function

    If m_BI.bmiHeader.biBitCount = 24 Then
        lhDCDesktop = GetDC(GetDesktopWindow())
    Else
        lhDCDesktop = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    End If
    If lhDCDesktop = 0 Then Exit Function

    lHDC = CreateCompatibleDC(lhDCDesktop)
    DeleteDC lhDCDesktop
    If lHDC = 0 Then Exit Function
    lhBmpOld = SelectObject(lHDC, TheStdPicture.handle)
    If m_BI.bmiHeader.biBitCount = 24 Then
        BitBlt m_hDC, 0, 0, m_BI.bmiHeader.biWidth, m_BI.bmiHeader.biHeight, lHDC, 0, 0, vbSrcCopy
    Else
        If CopyPalette Then
            Dim lC As Long
            Dim C2 As Long
            C2 = 2 ^ m_BI.bmiHeader.biBitCount
            lC = GetDIBColorTable(lHDC, 0, C2, m_RGB(0))
            If (lC > 0) Then SetDIBColorTable m_hDC, 0, lC, m_RGB(0)
        End If
        GetDIBits256 lHDC, TheStdPicture.handle, 0, tBMP.bmHeight, ByVal m_Ptr, m_BI, DIB_RGB_COLORS
    End If
    SelectObject lHDC, lhBmpOld
    DeleteObject lHDC
    CopyStdPicture = True
End Function
Public Function CopyHDC(ByVal lHDC As Long, lWidth As Long, lHeight As Long, Optional ByVal iBitCount As Integer, Optional lSrcLeft As Long, Optional lSrcTop As Long) As Boolean
    Dim C1 As Long
    If iBitCount = 0 Then
        C1 = GetDIBColorTable(lHDC, 0, 256, m_RGB(0))
        Select Case C1
        Case 1 To 2:     iBitCount = 1
        Case 3 To 16:    iBitCount = 4
        Case 17 To 256:  iBitCount = 8
        Case Else:       iBitCount = 24
        End Select
    End If
    If Not Create(lWidth, lHeight, iBitCount) Then Exit Function
    If C1 > 0 Then SetDIBColorTable m_hDC, 0, C1, m_RGB(0)
    BitBlt m_hDC, 0, 0, lWidth, lHeight, lHDC, lSrcLeft, lSrcTop, vbSrcCopy
    CopyHDC = True
End Function

Public Function CopyPalletHDC(ByVal lHDC As Long) As Boolean
    Dim g As Long

    g = GetDIBColorTable(lHDC, 0, 2 ^ m_BI.bmiHeader.biBitCount, m_RGB(0))
    If g > 0 Then CopyPalletHDC = (g = SetDIBColorTable(m_hDC, 0, g, m_RGB(0)))
End Function



'====================================================================================
'                              PAINT/PASTE SECTIONS
'====================================================================================
Public Sub PaintHDC(lHDC As Long, Optional lDestLeft As Long, Optional lDestTop As Long, Optional eRop As RasterOpConstants = vbSrcCopy)
    BitBlt lHDC, lDestLeft, lDestTop, m_BI.bmiHeader.biWidth, m_BI.bmiHeader.biHeight, m_hDC, 0, 0, eRop
End Sub



'====================================================================================
'                               DISPLAY FUNCTIONS
'====================================================================================
'The following functions return modified versions of this class for display purposes.
'They are not meant to be used as reliable image processing routines, because the
'PlgBlt() and StretchBlt() API calls are not precise.

Public Function Greyscale() As cImage
    Set Greyscale = New cImage        'Return 8 bit Greyscale version of this cImage
    Greyscale.Create m_BI.bmiHeader.biWidth, m_BI.bmiHeader.biHeight, 8
    BitBlt Greyscale.hDC, 0, 0, m_BI.bmiHeader.biWidth, m_BI.bmiHeader.biHeight, m_hDC, 0, 0, vbSrcCopy
End Function

Public Function Resample(lWidth As Long, lHeight As Long) As cImage
    Set Resample = New cImage         'Return a resized version of this cImage
    Resample.Create lWidth, lHeight, m_BI.bmiHeader.biBitCount
    If m_BI.bmiHeader.biBitCount <> 24 Then Resample.CopyPalletHDC m_hDC

    If (lWidth = m_BI.bmiHeader.biWidth) And (lHeight = m_BI.bmiHeader.biHeight) Then
       'Just return a copy
        BitBlt Resample.hDC, 0, 0, lWidth, lHeight, m_hDC, 0, 0, vbSrcCopy
    Else
       'HALFTONE gives better quality at slower speed, but it's unsupported in Win 95, 98, ME.
       'If we can't use HALFTONE, use COLORONCOLOR.  The default BLACKONWHITE is unacceptable.
        If SetStretchBltMode(Resample.hDC, HALFTONE) = 0 Then SetStretchBltMode Resample.hDC, COLORONCOLOR
        StretchBlt Resample.hDC, 0, 0, lWidth, lHeight, m_hDC, 0, 0, m_BI.bmiHeader.biWidth, m_BI.bmiHeader.biHeight, vbSrcCopy
    End If
End Function

Public Function Mirror(Vertical As Boolean) As cImage
    Dim MyPoint(2) As POINTAPI 'Return a mirror image of this cImage

    If Vertical Then
        MyPoint(0).x = 0
        MyPoint(0).y = m_BI.bmiHeader.biHeight
        MyPoint(1).x = m_BI.bmiHeader.biWidth
        MyPoint(1).y = m_BI.bmiHeader.biHeight
        MyPoint(2).x = 0
        MyPoint(2).y = 0
    Else
        MyPoint(0).x = m_BI.bmiHeader.biWidth
        MyPoint(0).y = 0
        MyPoint(1).x = 0
        MyPoint(1).y = 0
        MyPoint(2).x = m_BI.bmiHeader.biWidth
        MyPoint(2).y = m_BI.bmiHeader.biHeight
    End If

    Set Mirror = New cImage
    Mirror.Create m_BI.bmiHeader.biWidth, m_BI.bmiHeader.biHeight, m_BI.bmiHeader.biBitCount
    If m_BI.bmiHeader.biBitCount <> 24 Then Mirror.CopyPalletHDC m_hDC
    PlgBlt Mirror.hDC, MyPoint(0), m_hDC, 0, 0, m_BI.bmiHeader.biWidth, m_BI.bmiHeader.biHeight, 0, 0, 0
End Function

Public Function Rotate(ByVal Degrees As Long) As cImage
    Dim NewWidth     As Long 'Return version of this cImage rotated Degrees
    Dim NewHeight    As Long
    Dim MyPoint(2)   As POINTAPI

    Degrees = Degrees Mod 360
    If Degrees < 0 Then Degrees = Degrees + 360

    Select Case Degrees
    Case 90
        MyPoint(0).x = 0
        MyPoint(0).y = m_BI.bmiHeader.biWidth
        MyPoint(1).x = 0
        MyPoint(1).y = 0
        MyPoint(2).x = m_BI.bmiHeader.biHeight
        MyPoint(2).y = m_BI.bmiHeader.biWidth
        NewWidth = m_BI.bmiHeader.biHeight
        NewHeight = m_BI.bmiHeader.biWidth
    Case 180
        MyPoint(0).x = m_BI.bmiHeader.biWidth
        MyPoint(0).y = m_BI.bmiHeader.biHeight
        MyPoint(1).x = 0
        MyPoint(1).y = m_BI.bmiHeader.biHeight
        MyPoint(2).x = m_BI.bmiHeader.biWidth
        MyPoint(2).y = 0
        NewWidth = m_BI.bmiHeader.biWidth
        NewHeight = m_BI.bmiHeader.biHeight
    Case 270
        MyPoint(0).x = m_BI.bmiHeader.biHeight
        MyPoint(0).y = 0
        MyPoint(1).x = m_BI.bmiHeader.biHeight
        MyPoint(1).y = m_BI.bmiHeader.biWidth
        MyPoint(2).x = 0
        MyPoint(2).y = 0
        NewWidth = m_BI.bmiHeader.biHeight
        NewHeight = m_BI.bmiHeader.biWidth
    Case Else
        Exit Function
    End Select

    Set Rotate = New cImage
    Rotate.Create NewWidth, NewHeight, m_BI.bmiHeader.biBitCount
    If m_BI.bmiHeader.biBitCount <> 24 Then Rotate.CopyPalletHDC m_hDC
    PlgBlt Rotate.hDC, MyPoint(0), m_hDC, 0, 0, m_BI.bmiHeader.biWidth, m_BI.bmiHeader.biHeight, 0, 0, 0
End Function

