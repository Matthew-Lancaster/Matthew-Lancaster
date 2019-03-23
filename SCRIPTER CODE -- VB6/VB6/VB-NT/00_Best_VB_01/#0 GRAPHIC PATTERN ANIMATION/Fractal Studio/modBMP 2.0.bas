Attribute VB_Name = "modBMP"
Option Explicit

Public Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Public Declare Function CreateDIBS Lib "gdi32" Alias "CreateDIBSection" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long

Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type

Type SafeArray1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0) As SafeArrayBound
End Type

Type tBitmap
    FileHeader As BITMAPFILEHEADER
    InfoHeader As BITMAPINFOHEADER
    Data() As Byte
    W As Long
    H As Long
    W32 As Long
    tblOffsetX() As Long    'X offset
    tblOffsetY() As Long    'Y offset
End Type

Type tMemBMP
    hdc As Long
    hBMP As Long
    W As Long
    H As Long
    OldBMP As Long
End Type

Type tDIBSection
    BMP As tBitmap
    mBMP As tMemBMP
    ArrayInfo As SafeArray1D
    memPointer As Long
    Ready As Boolean
End Type

Global MyBMP  As tBitmap
Global MyMemBMP As tMemBMP

Public Function SaveBMP(fPath As String, BMP As tBitmap) As Boolean
On Error GoTo Err_SaveBMP

Dim fNum As Integer

If Dir(fPath) <> vbNullString Then Kill fPath

fNum = FreeFile
Open fPath For Binary As fNum
    Put #fNum, 1, BMP.FileHeader
    Put #fNum, , BMP.InfoHeader
    Put #fNum, , BMP.Data
Close #fNum

SaveBMP = True
Exit Function

Err_SaveBMP:
SaveBMP = False

End Function

Public Function CreateBMP(W As Long, H As Long, Optional InitData As Boolean = True) As tBitmap
Dim BMP As tBitmap, W32 As Long, Cnt As Long
Dim ColTableSize As Long


W32 = W * 4
With BMP.FileHeader
    .bfReserved1 = 0
    .bfReserved2 = 0
    .bfSize = Len(BMP.FileHeader) + Len(BMP.InfoHeader) + (W32 * H)
    .bfType = 19778
    .bfOffBits = Len(BMP.FileHeader) + Len(BMP.InfoHeader)
End With

With BMP.InfoHeader
    .biBitCount = 32
    .biClrImportant = 0
    .biClrUsed = 0
    .biCompression = 0
    .biHeight = H
    .biPlanes = 1
    .biSize = 40
    .biSizeImage = W32 * H
    .biWidth = W
    .biXPelsPerMeter = 0
    .biYPelsPerMeter = 0
End With

BMP.W = W
BMP.H = H
BMP.W32 = W32
If InitData Then
    ReDim BMP.Data(0 To (W32 * H) - 1)
End If
ReDim BMP.tblOffsetY(0 To H - 1)
ReDim BMP.tblOffsetX(0 To W - 1)

For Cnt = 0 To BMP.H - 1
    BMP.tblOffsetY(Cnt) = (BMP.H - Cnt - 1) * BMP.W32
Next Cnt

For Cnt = 0 To BMP.W - 1
    BMP.tblOffsetX(Cnt) = Cnt * 4
Next Cnt

CreateBMP = BMP

End Function

Public Function CreateMemBMP(BMP As tBitmap) As tMemBMP
Dim TmpBMP As tMemBMP
Dim BMPInfo As BITMAPINFO
Dim ret As Long

With TmpBMP
    .W = BMP.W
    .H = BMP.H
    .hdc = CreateCompatibleDC(0)
    .hBMP = CreateCompatibleBitmap(GetWindowDC(GetDesktopWindow), .W, .H)
    BMPInfo.bmiHeader = BMP.InfoHeader
    ret = SetDIBits(.hdc, .hBMP, 0, BMP.H, BMP.Data(0), BMPInfo, DIB_RGB_COLORS)
    .OldBMP = SelectObject(.hdc, .hBMP)
    If .hdc = 0 Or .hBMP = 0 Or ret = 0 Then
        MsgBox "Unable to create DIB.." & vbCrLf & vbCrLf & "[Requested Size:]" & _
               "W > " & Trim(Str(BMP.W)) & vbCrLf & "H > " & Trim(Str(BMP.H)), vbExclamation Or vbOKOnly
        aData.Flag_Stop = True
    End If
End With

If TmpBMP.hdc <> 0 And TmpBMP.hBMP <> 0 Then 'Succeeded
    CreateMemBMP = TmpBMP
End If

End Function

Public Function DIBInsertData(DestMemBMP As tMemBMP, SrcBMP As tBitmap) As Boolean
Dim ret As Long
Dim BMPInfo As BITMAPINFO

BMPInfo.bmiHeader = SrcBMP.InfoHeader
With DestMemBMP
    ret = SetDIBits(.hdc, .hBMP, 0, SrcBMP.H, SrcBMP.Data(0), BMPInfo, DIB_RGB_COLORS)
End With

DIBInsertData = ret
End Function

Public Sub SetPoint(BMP As tBitmap, x As Long, Y As Long, R As Byte, G As Byte, B As Byte)
Dim Pos As Long

If x < BMP.W - 1 And Y < BMP.H - 1 Then
    With BMP
        Pos = .tblOffsetY(Y) + .tblOffsetX(x)
        .Data(Pos) = B
        .Data(Pos + 1) = G
        .Data(Pos + 2) = R
    End With
End If

End Sub

Public Sub CleanUp()

FreeMemBMP MyMemBMP

End Sub
Public Sub FreeMemBMP(Var As tMemBMP)

SelectObject Var.hdc, Var.OldBMP
DeleteObject Var.hBMP
DeleteObject Var.OldBMP
DeleteDC Var.hdc

End Sub

Public Sub FreeBitmapData(Var As tBitmap)

With Var
    ReDim .Data(0)
    ReDim .tblOffsetX(0)
    ReDim .tblOffsetY(0)
End With

End Sub

Public Function LoadBMP(fPath As String, BMP As tBitmap) As Boolean
Dim fNum As Integer, Cnt As Long

If Dir(fPath) = "" Then GoTo LoadBMP_Failed

fNum = FreeFile
Open fPath For Binary As fNum
    With BMP
        Get #fNum, 1, .FileHeader
        Get #fNum, , .InfoHeader
        .W = .InfoHeader.biWidth
        .W32 = ((.W * 3) \ 4) * 4: If .W32 < (.W * 3) Then .W32 = .W32 + 4
        .H = .InfoHeader.biHeight
        
        ReDim .tblOffsetX(0 To .W - 1)
        ReDim .tblOffsetY(0 To .H - 1)
        
        For Cnt = 0 To .H - 1
            .tblOffsetY(Cnt) = (.H - Cnt - 1) * .W32
        Next Cnt

        For Cnt = 0 To .W - 1
            .tblOffsetX(Cnt) = Cnt * 3
        Next Cnt
        
        ReDim .Data(0 To (.W32 * .InfoHeader.biHeight * 3))
        Get #fNum, , .Data
    End With
Close #fNum

LoadBMP = True

Exit Function
LoadBMP_Failed:

LoadBMP = False

End Function

Public Function CreateDIB(nW As Long, nH As Long, DIB As tDIBSection) As Boolean
Dim BMPInf As BITMAPINFO

DIB.BMP = CreateBMP(nW, nH)
Erase DIB.BMP.Data

BMPInf.bmiHeader = DIB.BMP.InfoHeader

DIB.mBMP.W = nW
DIB.mBMP.H = nH
DIB.mBMP.hdc = CreateCompatibleDC(0)
DIB.mBMP.hBMP = CreateDIBS(DIB.mBMP.hdc, BMPInf, DIB_RGB_COLORS, DIB.memPointer, 0, 0)
DIB.mBMP.OldBMP = SelectObject(DIB.mBMP.hdc, DIB.mBMP.hBMP)

End Function

Public Sub PrepareDIBSection(DIB As tDIBSection)

If DIB.Ready Then Exit Sub

With DIB.ArrayInfo
    .cbElements = 0
    .cDims = 1
    .Bounds(0).lLbound = 0
    .Bounds(0).cElements = DIB.BMP.W32 * DIB.BMP.H
    .pvData = DIB.memPointer
End With

CopyMemory ByVal VarPtrArray(DIB.BMP.Data()), VarPtr(DIB.ArrayInfo), 4
DIB.Ready = True

End Sub

Public Sub FinishDIBSection(DIB As tDIBSection)

If DIB.Ready = False Then Exit Sub

CopyMemory ByVal VarPtrArray(DIB.BMP.Data), 0&, 4
DIB.Ready = False

End Sub

Public Sub FreeDIB(DIB As tDIBSection)

If DIB.Ready Then FinishDIBSection DIB
Erase DIB.BMP.Data
Erase DIB.BMP.tblOffsetX
Erase DIB.BMP.tblOffsetY
SelectObject DIB.mBMP.hdc, DIB.mBMP.OldBMP
DeleteDC DIB.mBMP.hdc
DeleteObject DIB.mBMP.hBMP

End Sub
