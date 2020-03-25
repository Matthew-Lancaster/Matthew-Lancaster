Attribute VB_Name = "modAVI"
Option Explicit

'************** A V I - F I L E - F O R M A T *******
'****************************************************
'*
'* "RIFF" "AVI " "LIST" /list length/
'*    {Header list
'*        AVI header
'*        Stream List
'*            {
'*            Stream Header
'*            Stream Format
'*            Stream Additional Data
'*            }
'*    "MOVI" - Data chunks
'*        "00DB" /chunk data/     (for each frame)
'*    "idx1"
'*        Index entry chunk       (for each keyframe)
'*    }
'*
'****************************************************
'****************************************************

Type tRECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type

Type tAVIIndexEntry
    fccCKID As String * 4
    dwFlags As Long
    dwChunkOffset As Long
    dwChunkLength As Long
End Type

Type tMainAVIHeader
    dwMicroSecPerFrame As Long
    dwMaxBytesPerSec As Long
    dwReserved1 As Long
    dwFlags As Long
    dwTotalFrames As Long
    dwInitialFrames As Long
    dwStreams As Long
    dwSuggestedBufferSize As Long
    dwWidth As Long
    dwHeight As Long
    dwScale As Long
    dwRate As Long
    dwStart As Long
    dwLength As Long
End Type

Type tAVIStreamHeader
    fccType As String * 4
    fccHandler As String * 4
    dwFlags As Long
    dwReserved1 As Long
    dwInitialFrames As Long
    dwScale As Long
    dwRate As Long
    dwStart As Long
    dwLength As Long
    dwSuggestedBufferSize As Long
    dwQuality As Long
    dwSampleSize As Long
    rcFrame As tRECT
End Type

Type tAVIChunk
    StreamID As String * 4
    Data() As Byte
End Type

Type tAVI
    Header As String
    MainH As tMainAVIHeader         'Main AVI Header
    StrH As tAVIStreamHeader        'The Video Stream Header
    szFile As String                'The file to write to
    Created As Boolean              'Whether the file has been created yet
    nOffset As Long                 'The offset into the file (to next chunk)
    BMP As tBitmap                  'A Bitmap structure used for initializing format data
    BMPInfo As BITMAPINFO           'BITMAPINFO structure used as Stream Format chunk
    IndexEntry() As tAVIIndexEntry  'AVI Index entries (for each  keyframe)
    nFrames As Long                 'Total number of frames
    fNum As Integer                 'File number while writing to file
End Type

Global Const LENGTH_AVIHEADER = 56
Global Const LENGTH_STREAMHEADER = 56
Global Const LENGTH_STREAMFORMAT = 40
Global Const LENGTH_INDEXENTRY = 16
Global Const LENGTH_HEADERPADDED = 2048

Global Const AVIF_TRUSTCKTYPE = &H800
Global Const AVIF_HASINDEX = &H10
Global Const AVIIF_KEYFRAME = &H10
'-------------------------------------'

Public Sub InitAVI(AVI As tAVI, W As Long, H As Long, mSpF As Long, nFrames As Long)
Dim nFPS As Integer, Cnt As Long

'MAX FPS:
'------------------
nFPS = 1000 \ mSpF
If nFPS > nFrames Then nFPS = nFrames
'------------------

With AVI
    AVI.BMP = CreateBMP(W, H)
    AVI.nFrames = nFrames
    ReDim AVI.IndexEntry(0 To nFrames - 1)
End With

With AVI.MainH
    .dwMicroSecPerFrame = mSpF * 1000
    .dwMaxBytesPerSec = AVI.BMP.W32 * H * nFPS
    .dwReserved1 = 0
    .dwFlags = AVIF_HASINDEX Or AVIF_TRUSTCKTYPE
    .dwTotalFrames = nFrames
    .dwInitialFrames = 0
    .dwStreams = 1
    .dwSuggestedBufferSize = AVI.BMP.W32 * H
    .dwWidth = W
    .dwHeight = H
    .dwScale = 0
    .dwRate = 0
    .dwStart = 0
    .dwLength = 0
End With

With AVI.StrH
    .fccType = "vids"
    .fccHandler = DW2Str(0)
    .dwFlags = 0
    .dwReserved1 = 0
    .dwInitialFrames = 0
    .dwScale = 1
    .dwRate = 1000 / mSpF
    .dwStart = 0
    .dwLength = nFrames
    .dwSuggestedBufferSize = AVI.MainH.dwSuggestedBufferSize
    .dwQuality = 0
    .dwSampleSize = 0
    With .rcFrame
        .Left = 0
        .Top = 0
        .Right = W
        .Bottom = H
    End With
End With

For Cnt = 0 To nFrames - 1
    With AVI.IndexEntry(Cnt)
        .fccCKID = "00db"
        .dwFlags = AVIIF_KEYFRAME
        .dwChunkOffset = 4 + (Cnt * (8 + AVI.BMP.W32 * AVI.BMP.H))
        .dwChunkLength = AVI.BMP.W32 * H
    End With
Next Cnt

With AVI.BMPInfo
    .bmiHeader = AVI.BMP.InfoHeader
End With

End Sub

Public Sub WriteAVIHeader(AVI As tAVI, szFile As String)
Dim fNum As Integer, Buf As String
Dim nLen As Long, PadSize As Long, nFrameSize As Long, Cnt As Long

fNum = FreeFile
Open szFile For Binary As fNum
    With AVI
    
        '----------------------- HEADERS ----------------------------------
        '------------------------------------------------------------------
        
        'RIFF form + AVI Main Header:
        nFrameSize = .BMP.W32 * .BMP.H
        nLen = LENGTH_HEADERPADDED + .nFrames * ((2 * 4) + nFrameSize) + .nFrames * _
               LENGTH_INDEXENTRY
        Buf = "RIFF" & DW2Str(nLen) & "AVI " & "LIST"
        nLen = 10 * 4 + LENGTH_AVIHEADER + LENGTH_STREAMHEADER + LENGTH_STREAMFORMAT
        Buf = Buf & DW2Str(nLen) & "hdrlavih" & DW2Str(LENGTH_AVIHEADER)
        Put #fNum, 1, Buf
        Put #fNum, , .MainH
        
        'Stream Header:
        nLen = 5 * 4 + LENGTH_STREAMHEADER + LENGTH_STREAMFORMAT
        Buf = "LIST" & DW2Str(nLen) & "strlstrh" & DW2Str(LENGTH_STREAMHEADER)
        Put #fNum, , Buf
        Put #fNum, , .StrH
        
        'Stream Format
        Buf = "strf" & DW2Str(LENGTH_STREAMFORMAT)
        Put #fNum, , Buf
        Put #fNum, , .BMPInfo.bmiHeader
    
        'JUNK section (2K padding) :
        nLen = (2048 - (3 * 4)) - (LENGTH_AVIHEADER + LENGTH_STREAMHEADER + _
                                   LENGTH_STREAMFORMAT + (17 * 4))
        Buf = "JUNK" & DW2Str(nLen) & String(nLen, 0)
        Put #fNum, , Buf
    
        '----------------- VIDEO DATA -------------------------------------
        '------------------------------------------------------------------
        
        'Write start of video data to file:
        nLen = .nFrames * (8 + nFrameSize) + 4
        Buf = "LIST" & DW2Str(nLen) & "movi"
        Put #fNum, 2048 - (3 * 4) + 1, Buf

        '----------------- INDEX DATA -------------------------------------
        '------------------------------------------------------------------
    
        Buf = "idx1" & DW2Str(LENGTH_INDEXENTRY * .nFrames)
        nLen = (2048 + .nFrames * (8 + nFrameSize)) + 1
        Put #fNum, nLen, Buf
        
        For Cnt = 0 To .nFrames - 1
            nLen = (2048 + .nFrames * (8 + nFrameSize) + 8) + (LENGTH_INDEXENTRY * Cnt) + 1
            Put #fNum, nLen, .IndexEntry(Cnt)
        Next Cnt

        '----------------- 1K PADDING -------------------------------------
        '------------------------------------------------------------------
    
        nLen = 2048 + 4 + (.nFrames * (8 + nFrameSize)) + 8 + (.nFrames * LENGTH_INDEXENTRY)
        PadSize = nLen \ 1024
        If (PadSize * 1024) <> nLen Then
            PadSize = PadSize + 1
        End If
        PadSize = (PadSize * 1024) - nLen
    
        Buf = String(PadSize, 0)
        Put #fNum, nLen + 1, Buf
    
        '------------------------------------------------------------------
        '------------------------------------------------------------------
        
        'Turn Created flag on:
        .Created = True
        'Specify file:
        .szFile = szFile
        
    End With
Close #fNum

End Sub

Public Sub WriteAVIFrame(AVI As tAVI, BMP As tBitmap, nFrame As Long, IsOpen As Boolean, CloseFile As Boolean)
Dim nFrameSize As Long, Buf As String

'Check frame validity before writing:
'----------------------------------------------
If Not AVI.Created Then
    MsgBox "The AVI file has not yet been created."
    Exit Sub
End If
If nFrame > AVI.nFrames - 1 Then
    MsgBox "The AVI file was initialized to hold a lower number of frames."
    Exit Sub
End If
If BMP.W <> AVI.BMP.W Or BMP.H <> AVI.BMP.H Then
    MsgBox "The AVI was initialized to a different frame size."
    Exit Sub
End If
'----------------------------------------------

If Not IsOpen Then
    AVI.fNum = FreeFile
    Open AVI.szFile For Binary As AVI.fNum
End If
    nFrameSize = BMP.W32 * BMP.H
    Buf = "00db" & DW2Str(nFrameSize)
    Put AVI.fNum, 2048 + nFrame * (8 + nFrameSize) + 1, Buf
    Put AVI.fNum, 2048 + nFrame * (8 + nFrameSize) + 8 + 1, BMP.Data
If CloseFile Then
    Close AVI.fNum
End If

End Sub

Public Function DW2Str(dwNum As Long) As String
Dim b(0 To 3) As Long

b(0) = dwNum And 255
b(1) = (dwNum And 65280) \ 256
b(2) = (dwNum And 16711680) \ 65536
b(3) = (dwNum And 2130706432) \ 16777216
'All except the most significant bit, which causes overflow in VB..

DW2Str = Chr$(b(0)) & Chr$(b(1)) & Chr$(b(2)) & Chr$(b(3))

End Function
