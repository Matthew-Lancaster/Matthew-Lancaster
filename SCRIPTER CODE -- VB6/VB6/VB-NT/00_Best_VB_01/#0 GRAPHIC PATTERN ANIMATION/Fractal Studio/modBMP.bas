Attribute VB_Name = "modBMP"

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
    hDC As Long
    hBMP As Long
    W As Long
    H As Long
    OldBMP As Long
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

Public Function CreateBMP(W As Long, H As Long) As tBitmap
Dim BMP As tBitmap, W32 As Long, Cnt As Long
Dim ColTableSize As Long


W32 = ((W * 3) \ 4) * 4: If W32 < (W * 3) Then W32 = W32 + 4

With BMP.FileHeader
    .bfReserved1 = 0
    .bfReserved2 = 0
    .bfSize = Len(BMP.FileHeader) + Len(BMP.InfoHeader) + (W32 * H)
    .bfType = 19778
    .bfOffBits = Len(BMP.FileHeader) + Len(BMP.InfoHeader)
End With

With BMP.InfoHeader
    .biBitCount = 24
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
ReDim BMP.Data(0 To (W32 * H) - 1)
ReDim BMP.tblOffsetY(0 To H - 1)
ReDim BMP.tblOffsetX(0 To W - 1)

For Cnt = 0 To BMP.H - 1
    BMP.tblOffsetY(Cnt) = (BMP.H - Cnt - 1) * BMP.W32
Next Cnt

For Cnt = 0 To BMP.W - 1
    BMP.tblOffsetX(Cnt) = Cnt * 3
Next Cnt

CreateBMP = BMP

End Function

Public Function CreateMemBMP(BMP As tBitmap) As tMemBMP
Dim TmpBMP As tMemBMP
Dim BMPInfo As BITMAPINFO

With TmpBMP
    .W = BMP.W
    .H = BMP.H
    .hDC = CreateCompatibleDC(GetWindowDC(GetDesktopWindow))
    .hBMP = CreateCompatibleBitmap(frmMain.hDC, .W, .H)
    BMPInfo.bmiHeader = BMP.InfoHeader
    SetDIBits .hDC, .hBMP, 0, BMP.H, BMP.Data(0), BMPInfo, DIB_RGB_COLORS
    .OldBMP = SelectObject(.hDC, .hBMP)
End With

If TmpBMP.hDC > 0 And TmpBMP.hBMP > 0 Then 'Succeeded
    CreateMemBMP = TmpBMP
End If

End Function

Public Function CreateMemBMP_Partial(BMP As tBitmap, nStartLine As Long, nLines As Long) As tMemBMP
Dim TmpBMP As tMemBMP
Dim BMPInfo As BITMAPINFO

With TmpBMP
    .W = BMP.W
    .H = BMP.H
    .hDC = CreateCompatibleDC(GetWindowDC(GetDesktopWindow))
    .hBMP = CreateCompatibleBitmap(frmMain.hDC, .W, .H)
    BMPInfo.bmiHeader = BMP.InfoHeader
    SetDIBits .hDC, .hBMP, 0, nLines, BMP.Data(BMP.tblOffsetY(nStartLine)), BMPInfo, DIB_RGB_COLORS
    .OldBMP = SelectObject(.hDC, .hBMP)
End With

If TmpBMP.hDC > 0 And TmpBMP.hBMP > 0 Then 'Succeeded
    CreateMemBMP_Partial = TmpBMP
End If

End Function

Public Function DIBInsertData(DestMemBMP As tMemBMP, SrcBMP As tBitmap) As Boolean
Dim Ret As Long
Dim BMPInfo As BITMAPINFO

BMPInfo.bmiHeader = SrcBMP.InfoHeader
With DestMemBMP
    Ret = SetDIBits(.hDC, .hBMP, 0, SrcBMP.H, SrcBMP.Data(0), BMPInfo, DIB_RGB_COLORS)
End With

InsertData = Ret
End Function

Public Sub SetPoint(BMP As tBitmap, X As Long, Y As Long, R As Byte, G As Byte, B As Byte)
Dim Pos As Long

If X < BMP.W - 1 And Y < BMP.H - 1 Then
    With BMP
        Pos = .tblOffsetY(Y) + .tblOffsetX(X)
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

SelectObject Var.hDC, Var.OldBMP
DeleteObject Var.hBMP
DeleteObject Var.OldBMP
DeleteDC Var.hDC

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
