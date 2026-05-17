Attribute VB_Name = "modParamFile"
Option Explicit

Public Const PARAM_VER001 = 1
Public XXT
Type tFourCC
    b1 As Byte
    b2 As Byte
    b3 As Byte
    b4 As Byte
End Type

Type RGB_In_File
    R As Byte
    G As Byte
    B As Byte
End Type

Type tParamFile_SimpleProps_V01
    fccIdent As tFourCC         'Identification of file type
    sVersionID As Long          'Which version of the file format this is
    nType As Integer            'Fractal Type
    nW As Long                  'Output width in pixels
    nH As Long                  'Output height in pixels
    nDirectFileSize As Long     'BMP output size
    SX As Double                'Starting X coordinate
    SY As Double                'Starting Y coordinate
    EX As Double                'Ending X coordinate
    EY As Double                'Ending Y coordinate
    nBailout As Double          'Bailout value in calculation
    nMaxI As Double             'Maximum iteration count
    bLinearContinuity As Byte   'Whether the colors are interpolated linearly
    nPar0 As Double             'Calculation Param 0
    nPar1 As Double             'Calculation Param 1
    nPar2 As Double             'Calculation Param 2
    nPar3 As Double             'Calculation Param 3
    nPar4 As Double             'Calculation Param 4
    nLinesToDraw As Long        'Number of lines to render in one chunk
    nColPeriod As Long          'Palette interpolation interval
    FilterMapOut As tFilterVars 'Outside mapping filter
    FilterMapIn As tFilterVars  'Inside mapping filter
    bPalAnimNormalize As Byte   'Whether to normalize values before color cycling
    bNormIntervalSpec As Byte   'Whether a normalization range has been specified
    NormInterval As Long        'The normalization interval
End Type
Type tParamFile_ArrayInfo_V01
    nPalLen As Long             'Number of palette elements
    nFiltersOut As Long         'Number of outside filters
    nFiltersIn As Long          'Number of inside filters
    nZooms As Long              'Number of zooms
End Type

Public Sub SaveParamFile(nFrac As Long, szFile As String)
Dim fNum As Integer
Dim pF As tParamFile_SimpleProps_V01
Dim pArrInfo As tParamFile_ArrayInfo_V01
Dim ArrBuf() As Byte, nBufSize As Long, nPos As Long
Dim Cnt As Long
If szFile = "" Then Exit Sub
'Fill param file data structure:
With Frac(nFrac)

    pF.fccIdent.b1 = Asc("F")
    pF.fccIdent.b2 = Asc("D")
    pF.fccIdent.b3 = Asc("E")
    pF.fccIdent.b4 = Asc("F")
    
    pF.sVersionID = PARAM_VER001
    pF.nType = .nType 'frmParam.cmbFractType.TEXT 'frmParam.cmbFractType.LISTINDEX
    pF.nW = .nW
    pF.nH = .nH
    pF.nDirectFileSize = .nDirectFileSize
    pF.SX = .SX
    pF.SY = .SY
    pF.EX = .EX
    pF.EY = .EY
    pF.nBailout = .BailOut
    pF.nMaxI = .MaxI
    pF.bLinearContinuity = Bool2Byte(.LinearContinuity)
    pF.nPar0 = .Param(0)
    pF.nPar1 = .Param(1)
    pF.nPar2 = .Param(2)
    pF.nPar3 = .Param(3)
    pF.nPar4 = .Param(4)
    pF.nLinesToDraw = .LinesToDraw
    pF.nColPeriod = .nColPeriod
    pF.FilterMapOut = .FilterMapOut
    pF.FilterMapIn = .FilterMapIn
    pF.bPalAnimNormalize = Bool2Byte(.bPalAnimNormalize)
    pF.bNormIntervalSpec = Bool2Byte(.bPalAnimNormIntervalSpecified)
    pF.NormInterval = .nPalAnimNormInterval
    
    pArrInfo.nPalLen = .nPalLen
    pArrInfo.nFiltersOut = .nFiltersOut
    pArrInfo.nFiltersIn = .nFiltersIn
    pArrInfo.nZooms = .nZooms

    'Prepare array buffer:
    nBufSize = .nPalLen * 3
    'nBufSize = nBufSize + .nFiltersOut * 43
    'nBufSize = nBufSize + .nFiltersIn * 43
    'nBufSize = nBufSize + .nZooms * 33
    ReDim ArrBuf(0 To nBufSize - 1)
    
    nPos = 0
    'Fill in the palette:
    For Cnt = 0 To .nPalLen - 1
        ArrBuf(nPos) = .Pal(Cnt).R
        ArrBuf(nPos + 1) = .Pal(Cnt).G
        ArrBuf(nPos + 2) = .Pal(Cnt).B
        nPos = nPos + 3
    Next Cnt
End With

'Delete file if it exists (user should've been prompted for overwrite before this):
If FileExist(szFile) Then Kill szFile   'FileExist = defined function

'Write the file:
fNum = FreeFile
Open szFile For Binary As fNum
    Put #fNum, 1, pF
    Put #fNum, , pArrInfo
    Put #fNum, , ArrBuf
    Put #fNum, , Frac(nFrac).FilterOut
    Put #fNum, , Frac(nFrac).FilterIn
    Put #fNum, , Frac(nFrac).Zoom
Close fNum

End Sub

Public Sub LoadParamFile(szFile As String, Optional bShow As Boolean = True)
Dim fNum As Integer
Dim pF As tParamFile_SimpleProps_V01
Dim pArrInfo As tParamFile_ArrayInfo_V01
Dim FilePal() As RGB_In_File
Dim nFrac As Long

If Not FileExist(szFile) Then
    MsgBox "The specified file doesn't exist. Check your HDD or send bug report to author"
    Exit Sub
End If

'Extract basic info, wait with arrays.
fNum = FreeFile
Open szFile For Binary As fNum
Get #fNum, 1, pF
Get #fNum, , pArrInfo

'Check the file:
With pF.fccIdent
    If .b1 <> Asc("F") Or .b2 <> Asc("D") Or .b3 <> Asc("E") Or .b4 <> Asc("F") Then
        'Invalid file type, that's for sure.
        MsgBox "Invalid file type!"
        Close #fNum
        Exit Sub
    End If
End With
If pF.sVersionID <> PARAM_VER001 Then
    MsgBox "This file was saved with an old version of the program. Sorry to say it, but backwards compatiblity has not been implemented yet.. The file you're trying to open has version no " & Trim(Str(pF.sVersionID))
    Close #fNum
    Exit Sub
End If

'Make new fractal from par file:
nFrac = AddFractal(False)
With Frac(nFrac)
    .nType = pF.nType
    .nW = pF.nW
    .nH = pF.nH
    .nDirectFileSize = pF.nDirectFileSize
    .SX = pF.SX
    .SY = pF.SY
    .EX = pF.EX
    .EY = pF.EY
    .BailOut = pF.nBailout
    .MaxI = pF.nMaxI
    .LinearContinuity = pF.bLinearContinuity
    .Param(0) = pF.nPar0
    .Param(1) = pF.nPar1
    .Param(2) = pF.nPar2
    .Param(3) = pF.nPar3
    .Param(4) = pF.nPar4
    .LinesToDraw = pF.nLinesToDraw
    .nColPeriod = pF.nColPeriod
    .FilterMapOut = pF.FilterMapOut
    .FilterMapIn = pF.FilterMapIn
    .bPalAnimNormalize = pF.bPalAnimNormalize
    .bPalAnimNormIntervalSpecified = pF.bNormIntervalSpec
    .nPalAnimNormInterval = pF.NormInterval
    
    .nPalLen = pArrInfo.nPalLen
    .nFiltersOut = pArrInfo.nFiltersOut
    .nFiltersIn = pArrInfo.nFiltersIn
    .nZooms = pArrInfo.nZooms
    
    ReDim .Pal(0 To .nPalLen - 1)
    ReDim FilePal(0 To .nPalLen - 1)
    If .nFiltersOut > 0 Then ReDim .FilterOut(0 To .nFiltersOut - 1)
    If .nFiltersIn > 0 Then ReDim .FilterIn(0 To .nFiltersIn - 1)
    ReDim .Zoom(0 To .nZooms)
    
    'Read into arrays:
    Get #fNum, , FilePal
    Get #fNum, , .FilterOut
    Get #fNum, , .FilterIn
    Get #fNum, , .Zoom

    'Close the file:
    Close #fNum
    
    'Fix some additional stuff in struct:
    ReDim .Cell(0 To .nW - 1, 0 To .nH - 1)
    .InUse = True
    .nIndex = nFrac
    .ChunkBMP = CreateBMP(.nW, .LinesToDraw)
    .PalDisp = CreateMemBMP(CreateBMP(frmParam.imgPal.ScaleWidth, frmParam.imgPal.ScaleHeight))
    .BMP = CreateBMP(.nW, .nH)
    .BB = CreateMemBMP(.BMP)
    .SavedFractal = True
    .szPathFractal = szFile
    .UseDispersionControl = True 'default
    
    'Insert palette data into fract pal array:
    For Cnt = 0 To .nPalLen - 1
        .Pal(Cnt).R = FilePal(Cnt).R
        .Pal(Cnt).G = FilePal(Cnt).G
        .Pal(Cnt).B = FilePal(Cnt).B
    Next Cnt
    XXT = 1
    'frmParam.cmbFractType.ListIndex = pF.nType
    XXT = 0
End With

InitWin Frac(nFrac)
Frac(nFrac).Win.wWin.Caption = szFile
aData.nActive = nFrac

If bShow Then
    SizeFracWin Frac(nFrac)
    Frac(nFrac).Win.wWin.Show
    ShowParams Frac(nFrac)
    Calculate nFrac, True
End If

 

End Sub

Public Function Bool2Byte(Arg As Boolean) As Byte
    If Arg Then
        Bool2Byte = 1
    Else
        Bool2Byte = 0
    End If
End Function

Public Function FileExist(sFile As String) As Boolean
    FileExist = (Dir(sFile, vbHidden Or vbReadOnly Or vbSystem Or vbArchive) <> "")
End Function
