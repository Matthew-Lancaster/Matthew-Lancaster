Attribute VB_Name = "mVFW"
'****************************************************************
'*  VB file:   VFW.bas... VB32 wrapper for Win32 Video For Windows
'*                           functions.
'*  created:        1998 by Ray Mercer
'*
'*  last modified:  12/2/98 by Ray Mercer (added comments)
'*
'*
'*  a Visual Basic translation of Microsoft's vfw.h file which is
'*  a part of the Win32 Platform SDK
'*
'*  Copyright (c) 1998 Ray Mercer.  All rights reserved.
'****************************************************************

Option Explicit

'------------------------------------------------------------------
'  Messages which can be sent to an AVICAP window
'------------------------------------------------------------------

Public Const WM_USER As Long = &H400
Public Const WM_CAP_START As Long = WM_USER

Public Const WM_CAP_GET_CAPSTREAMPTR As Long = WM_CAP_START + 1

Public Const WM_CAP_SET_CALLBACK_ERROR As Long = WM_CAP_START + 2
Public Const WM_CAP_SET_CALLBACK_STATUS As Long = WM_CAP_START + 3
Public Const WM_CAP_SET_CALLBACK_YIELD As Long = WM_CAP_START + 4
Public Const WM_CAP_SET_CALLBACK_FRAME As Long = WM_CAP_START + 5
Public Const WM_CAP_SET_CALLBACK_VIDEOSTREAM As Long = WM_CAP_START + 6
Public Const WM_CAP_SET_CALLBACK_WAVESTREAM As Long = WM_CAP_START + 7
Public Const WM_CAP_GET_USER_DATA As Long = WM_CAP_START + 8
Public Const WM_CAP_SET_USER_DATA As Long = WM_CAP_START + 9
    
Public Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP_START + 10
Public Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP_START + 11
Public Const WM_CAP_DRIVER_GET_NAME As Long = WM_CAP_START + 12
Public Const WM_CAP_DRIVER_GET_VERSION As Long = WM_CAP_START + 13
Public Const WM_CAP_DRIVER_GET_CAPS As Long = WM_CAP_START + 14

Public Const WM_CAP_FILE_SET_CAPTURE_FILE As Long = WM_CAP_START + 20
Public Const WM_CAP_FILE_GET_CAPTURE_FILE As Long = WM_CAP_START + 21
Public Const WM_CAP_FILE_ALLOCATE As Long = WM_CAP_START + 22
Public Const WM_CAP_FILE_SAVEAS As Long = WM_CAP_START + 23
Public Const WM_CAP_FILE_SET_INFOCHUNK As Long = WM_CAP_START + 24
Public Const WM_CAP_FILE_SAVEDIB As Long = WM_CAP_START + 25

Public Const WM_CAP_EDIT_COPY As Long = WM_CAP_START + 30

Public Const WM_CAP_SET_AUDIOFORMAT As Long = WM_CAP_START + 35
Public Const WM_CAP_GET_AUDIOFORMAT As Long = WM_CAP_START + 36

Public Const WM_CAP_DLG_VIDEOFORMAT As Long = WM_CAP_START + 41
Public Const WM_CAP_DLG_VIDEOSOURCE As Long = WM_CAP_START + 42
Public Const WM_CAP_DLG_VIDEODISPLAY As Long = WM_CAP_START + 43
Public Const WM_CAP_GET_VIDEOFORMAT As Long = WM_CAP_START + 44
Public Const WM_CAP_SET_VIDEOFORMAT As Long = WM_CAP_START + 45
Public Const WM_CAP_DLG_VIDEOCOMPRESSION As Long = WM_CAP_START + 46

Public Const WM_CAP_SET_PREVIEW As Long = WM_CAP_START + 50
Public Const WM_CAP_SET_OVERLAY As Long = WM_CAP_START + 51
Public Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP_START + 52
Public Const WM_CAP_SET_SCALE As Long = WM_CAP_START + 53
Public Const WM_CAP_GET_STATUS As Long = WM_CAP_START + 54
Public Const WM_CAP_SET_SCROLL As Long = WM_CAP_START + 55

Public Const WM_CAP_GRAB_FRAME As Long = WM_CAP_START + 60
Public Const WM_CAP_GRAB_FRAME_NOSTOP As Long = WM_CAP_START + 61

Public Const WM_CAP_SEQUENCE As Long = WM_CAP_START + 62
Public Const WM_CAP_SEQUENCE_NOFILE As Long = WM_CAP_START + 63
Public Const WM_CAP_SET_SEQUENCE_SETUP As Long = WM_CAP_START + 64
Public Const WM_CAP_GET_SEQUENCE_SETUP As Long = WM_CAP_START + 65
Public Const WM_CAP_SET_MCI_DEVICE As Long = WM_CAP_START + 66
Public Const WM_CAP_GET_MCI_DEVICE As Long = WM_CAP_START + 67
Public Const WM_CAP_STOP As Long = WM_CAP_START + 68
Public Const WM_CAP_ABORT As Long = WM_CAP_START + 69

Public Const WM_CAP_SINGLE_FRAME_OPEN As Long = WM_CAP_START + 70
Public Const WM_CAP_SINGLE_FRAME_CLOSE As Long = WM_CAP_START + 71
Public Const WM_CAP_SINGLE_FRAME As Long = WM_CAP_START + 72

Public Const WM_CAP_PAL_OPEN As Long = WM_CAP_START + 80
Public Const WM_CAP_PAL_SAVE As Long = WM_CAP_START + 81
Public Const WM_CAP_PAL_PASTE As Long = WM_CAP_START + 82
Public Const WM_CAP_PAL_AUTOCREATE As Long = WM_CAP_START + 83
Public Const WM_CAP_PAL_MANUALCREATE As Long = WM_CAP_START + 84

Public Const WM_CAP_SET_CALLBACK_CAPCONTROL As Long = WM_CAP_START + 85

' ------------------------------------------------------------------
'  VFW UDTS (from vfw.h)
' ------------------------------------------------------------------

Type VFWPOINT 'strange name to avoid collision with other POINT UDTs
        x As Long
        y As Long
End Type

Type CAPDRIVERCAPS
    wDeviceIndex As Long                    '// Driver index in system.ini
    fHasOverlay As Long                     '// Can device overlay?
    fHasDlgVideoSource As Long              '// Has Video source dlg?
    fHasDlgVideoFormat As Long              '// Has Format dlg?
    fHasDlgVideoDisplay As Long             '// Has External out dlg?
    fCaptureInitialized As Long             '// Driver ready to capture?
    fDriverSuppliesPalettes As Long         '// Can driver make palettes?
    hVideoIn As Long                        '// Driver In channel
    hVideoOut As Long                       '// Driver Out channel
    hVideoExtIn As Long                     '// Driver Ext In channel
    hVideoExtOut As Long                    '// Driver Ext Out channel
End Type

Type CAPSTATUS
    uiImageWidth As Long                    '// Width of the image
    uiImageHeight As Long                   '// Height of the image
    fLiveWindow As Long                     '// Now Previewing video?
    fOverlayWindow As Long                  '// Now Overlaying video?
    fScale As Long                          '// Scale image to client?
    ptScroll As VFWPOINT                    '// Scroll position
    fUsingDefaultPalette As Long            '// Using default driver palette?
    fAudioHardware As Long                  '// Audio hardware present?
    fCapFileExists As Long                  '// Does capture file exist?
    dwCurrentVideoFrame As Long             '// # of video frames cap'td
    dwCurrentVideoFramesDropped As Long     '// # of video frames dropped
    dwCurrentWaveSamples As Long            '// # of wave samples cap'td
    dwCurrentTimeElapsedMS As Long          '// Elapsed capture duration
    hPalCurrent As Long                     '// Current palette in use
    fCapturingNow As Long                   '// Capture in progress?
    dwReturn As Long                        '// Error value after any operation
    wNumVideoAllocated As Long              '// Actual number of video buffers
    wNumAudioAllocated As Long              '// Actual number of audio buffers
End Type

Public Const AVSTREAMMASTER_AUDIO As Long = 0  '/* Audio master (VFW 1.0, 1.1) */
Public Const AVSTREAMMASTER_NONE  As Long = 1  '/* No master */
Type CAPTUREPARMS
    dwRequestMicroSecPerFrame As Long       '// Requested capture rate
    fMakeUserHitOKToCapture As Long         '// Show "Hit OK to cap" dlg?
    wPercentDropForError As Long            '// Give error msg if > (10% default)
    fYield As Long                          '// Capture via background task?
    dwIndexSize As Long                     '// Max index size in frames (32K default)
    wChunkGranularity As Long               '// Junk chunk granularity (2K default)
    fUsingDOSMemory As Long                 '// Use DOS buffers? (obsolete)
    wNumVideoRequested As Long              '// # video buffers, If 0, autocalc
    fCaptureAudio As Long                   '// Capture audio?
    wNumAudioRequested As Long              '// # audio buffers, If 0, autocalc
    vKeyAbort As Long                       '// Virtual key causing abort
    fAbortLeftMouse As Long                 '// Abort on left mouse?
    fAbortRightMouse As Long                '// Abort on right mouse?
    fLimitEnabled As Long                   '// Use wTimeLimit?
    wTimeLimit As Long                      '// Seconds to capture
    fMCIControl As Long                     '// Use MCI video source?
    fStepMCIDevice As Long                  '// Step MCI device?
    dwMCIStartTime As Long                  '// Time to start in MS
    dwMCIStopTime As Long                   '// Time to stop in MS
    fStepCaptureAt2x As Long                '// Perform spatial averaging 2x
    wStepCaptureAverageFrames As Long       '// Temporal average n Frames
    dwAudioBufferSize As Long               '// Size of audio bufs (0 = default)
    fDisableWriteCache As Long              '// Attempt to disable write cache
    AVStreamMaster As Long                  '// Which stream controls length?
End Type
                                            

Type CAPINFOCHUNK
    fccInfoID As Long                       '// Chunk ID, "ICOP" for copyright
    lpData As Long                          '// pointer to data
    cbData As Long                          '// size of lpData
End Type

Type VIDEOHDR
    lpData As Long                          '// address of video buffer
    dwBufferLength As Long                  '// size, in bytes, of the Data buffer
    dwBytesUsed As Long                     '// see below
    dwTimeCaptured As Long                  '// see below
    dwUser As Long                          '// user-specific data
    dwFlags As Long                         '// see below
    dwReserved(3) As Long                   '// reserved; do not use
End Type


'//BITMAP DEFINES (from mmsystem.h)
Public Type BITMAPINFOHEADER
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

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors() As Long 'array of RGBQUADs
End Type


'// Capture Function Declares
Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" _
                                        (ByVal lpszWindowName As String, _
                                        ByVal dwStyle As Long, _
                                        ByVal x As Long, _
                                        ByVal y As Long, _
                                        ByVal nWidth As Long, _
                                        ByVal nHeight As Long, _
                                        ByVal hwndParent As Long, _
                                        ByVal nID As Long) As Long 'returns HWND

Declare Function capGetDriverDescription Lib "avicap32.dll" Alias "capGetDriverDescriptionA" _
                                        (ByVal dwDriverIndex As Long, _
                                        ByVal lpszName As String, _
                                        ByVal cbName As Long, _
                                        ByVal lpszVer As String, _
                                        ByVal cbVer As Long) As Long 'returns C BOOL

'VFW "customized" File Dialogs
'these are declared in the CmnDlg.bas file since they are actually OpenFile/SaveFIle dialogs

'Private Declare Function GetOpenFileNamePreview Lib "msvfw32.dll" _
'    Alias "GetOpenFileNamePreviewA" (filestruct As OPENFILENAME) As Long
'Private Declare Function GetSaveFileNamePreview Lib "msvfw32.dll" _
'    Alias "GetSaveFileNamePreviewA" (filestruct As OPENFILENAME) As Long

'//General WINAPI Declares
Private Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" _
                                            (ByVal hWnd As Long, _
                                            ByVal wMsg As Long, _
                                            ByVal wParam As Long, _
                                            ByVal lParam As Long) As Long
Private Declare Function SendMessageAsAny Lib "user32" Alias "SendMessageA" _
                                            (ByVal hWnd As Long, _
                                            ByVal wMsg As Long, _
                                            ByVal wParam As Long, _
                                            ByRef lParam As Any) As Long
Private Declare Function SendMessageAsString Lib "user32" Alias "SendMessageA" _
                                            (ByVal hWnd As Long, _
                                            ByVal wMsg As Long, _
                                            ByVal wParam As Long, _
                                            ByVal lParam As String) As Long
'// ------------------------------------------------------------------
'// IDs for status and error callbacks
'// ------------------------------------------------------------------

Public Const IDS_CAP_BEGIN As Long = 300              '/* "Capture Start" */
Public Const IDS_CAP_END As Long = 301                '/* "Capture End" */

Public Const IDS_CAP_INFO As Long = 401               '/* "%s" */
Public Const IDS_CAP_OUTOFMEM As Long = 402           '/* "Out of memory" */
Public Const IDS_CAP_FILEEXISTS As Long = 403         '/* "File '%s' exists -- overwrite it?" */
Public Const IDS_CAP_ERRORPALOPEN As Long = 404       '/* "Error opening palette '%s'" */
Public Const IDS_CAP_ERRORPALSAVE As Long = 405       '/* "Error saving palette '%s'" */
Public Const IDS_CAP_ERRORDIBSAVE As Long = 406       '/* "Error saving frame '%s'" */
Public Const IDS_CAP_DEFAVIEXT As Long = 407          '/* "avi" */
Public Const IDS_CAP_DEFPALEXT As Long = 408          '/* "pal" */
Public Const IDS_CAP_CANTOPEN As Long = 409           '/* "Cannot open '%s'" */
Public Const IDS_CAP_SEQ_MSGSTART As Long = 410       '/* "Select OK to start capture\nof video sequence\nto %s." */
Public Const IDS_CAP_SEQ_MSGSTOP As Long = 411        '/* "Hit ESCAPE or click to end capture" */
                
Public Const IDS_CAP_VIDEDITERR As Long = 412         '/* "An error occurred while trying to run VidEdit." */
Public Const IDS_CAP_READONLYFILE As Long = 413       '/* "The file '%s' is a read-only file." */
Public Const IDS_CAP_WRITEERROR As Long = 414         '/* "Unable to write to file '%s'.\nDisk may be full." */
Public Const IDS_CAP_NODISKSPACE As Long = 415        '/* "There is no space to create a capture file on the specified device." */
Public Const IDS_CAP_SETFILESIZE As Long = 416        '/* "Set File Size" */
Public Const IDS_CAP_SAVEASPERCENT As Long = 417      '/* "SaveAs: %2ld%%  Hit Escape to abort." */
                
Public Const IDS_CAP_DRIVER_ERROR As Long = 418       '/* Driver specific error message */

Public Const IDS_CAP_WAVE_OPEN_ERROR As Long = 419    '/* "Error: Cannot open the wave input device.\nCheck sample size, frequency, and channels." */
Public Const IDS_CAP_WAVE_ALLOC_ERROR As Long = 420   '/* "Error: Out of memory for wave buffers." */
Public Const IDS_CAP_WAVE_PREPARE_ERROR As Long = 421 '/* "Error: Cannot prepare wave buffers." */
Public Const IDS_CAP_WAVE_ADD_ERROR As Long = 422     '/* "Error: Cannot add wave buffers." */
Public Const IDS_CAP_WAVE_SIZE_ERROR As Long = 423    '/* "Error: Bad wave size." */
                
Public Const IDS_CAP_VIDEO_OPEN_ERROR As Long = 424   '/* "Error: Cannot open the video input device." */
Public Const IDS_CAP_VIDEO_ALLOC_ERROR As Long = 425  '/* "Error: Out of memory for video buffers." */
Public Const IDS_CAP_VIDEO_PREPARE_ERROR As Long = 426 '/* "Error: Cannot prepare video buffers." */
Public Const IDS_CAP_VIDEO_ADD_ERROR As Long = 427    '/* "Error: Cannot add video buffers." */
Public Const IDS_CAP_VIDEO_SIZE_ERROR As Long = 428   '/* "Error: Bad video size." */
                
Public Const IDS_CAP_FILE_OPEN_ERROR As Long = 429    '/* "Error: Cannot open capture file." */
Public Const IDS_CAP_FILE_WRITE_ERROR As Long = 430   '/* "Error: Cannot write to capture file.  Disk may be full." */
Public Const IDS_CAP_RECORDING_ERROR As Long = 431    '/* "Error: Cannot write to capture file.  Data rate too high or disk full." */
Public Const IDS_CAP_RECORDING_ERROR2 As Long = 432   '/* "Error while recording" */
Public Const IDS_CAP_AVI_INIT_ERROR As Long = 433     '/* "Error: Unable to initialize for capture." */
Public Const IDS_CAP_NO_FRAME_CAP_ERROR As Long = 434 '/* "Warning: No frames captured.\nConfirm that vertical sync interrupts\nare configured and enabled." */
Public Const IDS_CAP_NO_PALETTE_WARN As Long = 435    '/* "Warning: Using default palette." */
Public Const IDS_CAP_MCI_CONTROL_ERROR As Long = 436  '/* "Error: Unable to access MCI device." */
Public Const IDS_CAP_MCI_CANT_STEP_ERROR As Long = 437 '/* "Error: Unable to step MCI device." */
Public Const IDS_CAP_NO_AUDIO_CAP_ERROR As Long = 438 '/* "Error: No audio data captured.\nCheck audio card settings." */
Public Const IDS_CAP_AVI_DRAWDIB_ERROR As Long = 439  '/* "Error: Unable to draw this data format." */
Public Const IDS_CAP_COMPRESSOR_ERROR As Long = 440   '/* "Error: Unable to initialize compressor." */
Public Const IDS_CAP_AUDIO_DROP_ERROR As Long = 441   '/* "Error: Audio data was lost during capture, reduce capture rate." */
                
'/* status string IDs */
Public Const IDS_CAP_STAT_LIVE_MODE As Long = 500      '/* "Live window" */
Public Const IDS_CAP_STAT_OVERLAY_MODE As Long = 501   '/* "Overlay window" */
Public Const IDS_CAP_STAT_CAP_INIT As Long = 502       '/* "Setting up for capture - Please wait" */
Public Const IDS_CAP_STAT_CAP_FINI As Long = 503       '/* "Finished capture, now writing frame %ld" */
Public Const IDS_CAP_STAT_PALETTE_BUILD As Long = 504  '/* "Building palette map" */
Public Const IDS_CAP_STAT_OPTPAL_BUILD As Long = 505   '/* "Computing optimal palette" */
Public Const IDS_CAP_STAT_I_FRAMES As Long = 506       '/* "%d frames" */
Public Const IDS_CAP_STAT_L_FRAMES As Long = 507       '/* "%ld frames" */
Public Const IDS_CAP_STAT_CAP_L_FRAMES As Long = 508   '/* "Captured %ld frames" */
Public Const IDS_CAP_STAT_CAP_AUDIO As Long = 509      '/* "Capturing audio" */
Public Const IDS_CAP_STAT_VIDEOCURRENT As Long = 510   '/* "Captured %ld frames (%ld dropped) %d.%03d sec." */
Public Const IDS_CAP_STAT_VIDEOAUDIO As Long = 511     '/* "Captured %d.%03d sec.  %ld frames (%ld dropped) (%d.%03d fps).  %ld audio bytes (%d,%03d sps)" */
Public Const IDS_CAP_STAT_VIDEOONLY As Long = 512      '/* "Captured %d.%03d sec.  %ld frames (%ld dropped) (%d.%03d fps)" */

'Translations of C- "Message Cracker" Macros to VB (declared in vfw.h)

Function capSetCallbackOnError(ByVal hCapWnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnError = SendMessageAsLong(hCapWnd, WM_CAP_SET_CALLBACK_ERROR, 0&, lpProc)
End Function
Function capSetCallbackOnStatus(ByVal hCapWnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnStatus = SendMessageAsLong(hCapWnd, WM_CAP_SET_CALLBACK_STATUS, 0&, lpProc)
End Function
Function capSetCallbackOnYield(ByVal hCapWnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnYield = SendMessageAsLong(hCapWnd, WM_CAP_SET_CALLBACK_YIELD, 0&, lpProc)
End Function
Function capSetCallbackOnFrame(ByVal hCapWnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnFrame = SendMessageAsLong(hCapWnd, WM_CAP_SET_CALLBACK_FRAME, 0&, lpProc)
End Function
Function capSetCallbackOnVideoStream(ByVal hCapWnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnVideoStream = SendMessageAsLong(hCapWnd, WM_CAP_SET_CALLBACK_VIDEOSTREAM, 0&, lpProc)
End Function
Function capSetCallbackOnWaveStream(ByVal hCapWnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnWaveStream = SendMessageAsLong(hCapWnd, WM_CAP_SET_CALLBACK_WAVESTREAM, 0&, lpProc)
End Function
Function capSetCallbackOnCapControl(ByVal hCapWnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnCapControl = SendMessageAsLong(hCapWnd, WM_CAP_SET_CALLBACK_CAPCONTROL, 0&, lpProc)
End Function
Function capSetUserData(ByVal hCapWnd As Long, ByVal lUser As Long) As Boolean
   capSetUserData = SendMessageAsLong(hCapWnd, WM_CAP_SET_USER_DATA, 0&, lUser)
End Function
Function capGetUserData(ByVal hCapWnd As Long) As Long
   capGetUserData = SendMessageAsLong(hCapWnd, WM_CAP_GET_USER_DATA, 0&, 0&)
End Function
Function capDriverConnect(ByVal hCapWnd As Long, Optional ByVal i As Long = 0&) As Boolean
   capDriverConnect = SendMessageAsLong(hCapWnd, WM_CAP_DRIVER_CONNECT, i, 0&)
End Function
Function capDriverDisconnect(ByVal hCapWnd As Long) As Boolean
   capDriverDisconnect = SendMessageAsLong(hCapWnd, WM_CAP_DRIVER_DISCONNECT, 0&, 0&)
End Function
Function capDriverGetName(ByVal hCapWnd As Long) As String
   'returns driver name as VB string
   Dim szBuffer As String
   
   szBuffer = String$(128, 0)
   Call SendMessageAsString(hCapWnd, WM_CAP_DRIVER_GET_NAME, 128, szBuffer)
   capDriverGetName = Left$(szBuffer, InStr(szBuffer, vbNullChar) - 1)
End Function
Function capDriverGetVersion(ByVal hCapWnd As Long) As String
   'returns version as VB string
   Dim szBuffer As String
   Dim retVal As Boolean
   
   szBuffer = String$(128, 0)
   retVal = SendMessageAsString(hCapWnd, WM_CAP_DRIVER_GET_VERSION, 128, szBuffer)
    If 0 <> retVal Then
        capDriverGetVersion = Left$(szBuffer, InStr(szBuffer, vbNullChar) - 1)
    End If
End Function
Function capDriverGetCaps(ByVal hCapWnd As Long, ByRef Caps As CAPDRIVERCAPS) As Boolean
   'fills CAPDRIVERCAPS UDT
   capDriverGetCaps = SendMessageAsAny(hCapWnd, WM_CAP_DRIVER_GET_CAPS, Len(Caps), Caps)
End Function
Function capFileSetCaptureFile(ByVal hCapWnd As Long, ByVal FilePath As String) As Boolean
   capFileSetCaptureFile = SendMessageAsString(hCapWnd, WM_CAP_FILE_SET_CAPTURE_FILE, 0&, FilePath)
End Function
Function capFileGetCaptureFile(ByVal hCapWnd As Long) As String
   'returns full path of capture file as VB string
   Dim szBuffer As String
   Dim retVal As Boolean

   szBuffer = String$(128, 0)
   retVal = SendMessageAsString(hCapWnd, WM_CAP_FILE_GET_CAPTURE_FILE, 128, szBuffer)
    If retVal Then
        capFileGetCaptureFile = Left$(szBuffer, InStr(szBuffer, vbNullChar) - 1)
    End If
End Function
Function capFileAlloc(ByVal hCapWnd As Long, ByVal dwSize As Long) As Boolean
   capFileAlloc = SendMessageAsLong(hCapWnd, WM_CAP_FILE_ALLOCATE, 0&, dwSize)
End Function
Function capFileSaveAs(ByVal hCapWnd As Long, ByVal FilePath As String) As Boolean
   capFileSaveAs = SendMessageAsString(hCapWnd, WM_CAP_FILE_SAVEAS, 0&, FilePath)
End Function
Function capFileSetInfoChunk(ByVal hCapWnd As Long, ByRef InfChunk As CAPINFOCHUNK) As Boolean
   capFileSetInfoChunk = SendMessageAsAny(hCapWnd, WM_CAP_FILE_SET_INFOCHUNK, 0&, InfChunk)
End Function
Function capFileSaveDIB(ByVal hCapWnd As Long, ByVal FilePath As String) As Boolean
   capFileSaveDIB = SendMessageAsString(hCapWnd, WM_CAP_FILE_SAVEDIB, 0&, FilePath)
End Function
Function capEditCopy(ByVal hCapWnd As Long) As Boolean
   capEditCopy = SendMessageAsLong(hCapWnd, WM_CAP_EDIT_COPY, 0&, 0&)
End Function
Function capSetAudioFormat(ByVal hCapWnd As Long, ByRef wavFormat As WAVEFORMATEX, ByVal WavFormatSize As Long) As Boolean
   capSetAudioFormat = SendMessageAsAny(hCapWnd, WM_CAP_SET_AUDIOFORMAT, WavFormatSize, wavFormat)
End Function
Function capSetAudioFormatAsArray(ByVal hCapWnd As Long, ByVal wavFormat As Long, ByVal WavFormatSize As Long) As Boolean
   capSetAudioFormatAsArray = SendMessageAsLong(hCapWnd, WM_CAP_SET_AUDIOFORMAT, WavFormatSize, wavFormat)
End Function
Function capGetAudioFormat(ByVal hCapWnd As Long, ByRef wavFormat As WAVEFORMATEX, ByVal WavFormatSize As Long) As Long
   capGetAudioFormat = SendMessageAsAny(hCapWnd, WM_CAP_GET_AUDIOFORMAT, WavFormatSize, wavFormat)
End Function
Function capGetAudioFormatAsArray(ByVal hCapWnd As Long, ByVal wavFormat As Long, ByVal WavFormatSize As Long) As Long
   capGetAudioFormatAsArray = SendMessageAsLong(hCapWnd, WM_CAP_GET_AUDIOFORMAT, WavFormatSize, wavFormat)
End Function
Function capGetAudioFormatSize(ByVal hCapWnd As Long) As Long
   capGetAudioFormatSize = SendMessageAsLong(hCapWnd, WM_CAP_GET_AUDIOFORMAT, 0&, 0&)
End Function
Function capDlgVideoFormat(ByVal hCapWnd As Long) As Boolean
   capDlgVideoFormat = SendMessageAsLong(hCapWnd, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
End Function
Function capDlgVideoSource(ByVal hCapWnd As Long) As Boolean
   capDlgVideoSource = SendMessageAsLong(hCapWnd, WM_CAP_DLG_VIDEOSOURCE, 0&, 0&)
End Function
Function capDlgVideoDisplay(ByVal hCapWnd As Long) As Boolean
   capDlgVideoDisplay = SendMessageAsLong(hCapWnd, WM_CAP_DLG_VIDEODISPLAY, 0&, 0&)
End Function
Function capDlgVideoCompression(ByVal hCapWnd As Long) As Boolean
   capDlgVideoCompression = SendMessageAsLong(hCapWnd, WM_CAP_DLG_VIDEOCOMPRESSION, 0&, 0&)
End Function
Function capGetVideoFormat(ByVal hCapWnd As Long, ByRef BmpFormat As BITMAPINFO, ByVal CapFormatSize As Long) As Long
   capGetVideoFormat = SendMessageAsAny(hCapWnd, WM_CAP_GET_VIDEOFORMAT, CapFormatSize, BmpFormat)
End Function
Function capGetVideoFormatSize(ByVal hCapWnd As Long) As Long
   capGetVideoFormatSize = SendMessageAsLong(hCapWnd, WM_CAP_GET_VIDEOFORMAT, 0&, 0&)
End Function
Function capSetVideoFormat(ByVal hCapWnd As Long, ByRef BmpFormat As BITMAPINFO, ByVal CapFormatSize As Long) As Boolean
   capSetVideoFormat = SendMessageAsAny(hCapWnd, WM_CAP_SET_VIDEOFORMAT, CapFormatSize, BmpFormat)
End Function
Function capPreview(ByVal hCapWnd As Long, ByVal f As Boolean) As Boolean
   capPreview = SendMessageAsLong(hCapWnd, WM_CAP_SET_PREVIEW, -(f), 0&) 'convert the VB Boolean to a C BOOL with the - sign
End Function
Function capPreviewRate(ByVal hCapWnd As Long, ByVal wMS As Long) As Boolean
   capPreviewRate = SendMessageAsLong(hCapWnd, WM_CAP_SET_PREVIEWRATE, wMS, 0&)
End Function
Function capOverlay(ByVal hCapWnd As Long, ByVal f As Boolean) As Boolean
   capOverlay = SendMessageAsLong(hCapWnd, WM_CAP_SET_OVERLAY, -(f), 0&)
End Function
Function capPreviewScale(ByVal hCapWnd As Long, ByVal f As Boolean) As Boolean
   capPreviewScale = SendMessageAsLong(hCapWnd, WM_CAP_SET_SCALE, -(f), 0&)
End Function
Function capGetStatus(ByVal hCapWnd As Long, ByRef capStat As CAPSTATUS) As Boolean
   capGetStatus = SendMessageAsAny(hCapWnd, WM_CAP_GET_STATUS, Len(capStat), capStat)
End Function
Function capSetScrollPos(ByVal hCapWnd As Long, ByRef pt As VFWPOINT) As Boolean
   capSetScrollPos = SendMessageAsAny(hCapWnd, WM_CAP_SET_SCROLL, 0&, pt)
End Function
Function capGrabFrame(ByVal hCapWnd As Long) As Boolean
   capGrabFrame = SendMessageAsLong(hCapWnd, WM_CAP_GRAB_FRAME, 0&, 0&)
End Function
Function capGrabFrameNoStop(ByVal hCapWnd As Long) As Boolean
   capGrabFrameNoStop = SendMessageAsLong(hCapWnd, WM_CAP_GRAB_FRAME_NOSTOP, 0&, 0&)
End Function
Function capCaptureSequence(ByVal hCapWnd As Long) As Boolean
   capCaptureSequence = SendMessageAsLong(hCapWnd, WM_CAP_SEQUENCE, 0&, 0&)
End Function
Function capCaptureSequenceNoFile(ByVal hCapWnd As Long) As Boolean
   capCaptureSequenceNoFile = SendMessageAsLong(hCapWnd, WM_CAP_SEQUENCE_NOFILE, 0&, 0&)
End Function
Function capCaptureStop(ByVal hCapWnd As Long) As Boolean
   capCaptureStop = SendMessageAsLong(hCapWnd, WM_CAP_STOP, 0&, 0&)
End Function
Function capCaptureAbort(ByVal hCapWnd As Long) As Boolean
   capCaptureAbort = SendMessageAsLong(hCapWnd, WM_CAP_ABORT, 0&, 0&)
End Function
Function capCaptureSingleFrameOpen(ByVal hCapWnd As Long) As Boolean
   capCaptureSingleFrameOpen = SendMessageAsLong(hCapWnd, WM_CAP_SINGLE_FRAME_OPEN, 0&, 0&)
End Function
Function capCaptureSingleFrameClose(ByVal hCapWnd As Long) As Boolean
   capCaptureSingleFrameClose = SendMessageAsLong(hCapWnd, WM_CAP_SINGLE_FRAME_CLOSE, 0&, 0&)
End Function
Function capCaptureSingleFrame(ByVal hCapWnd As Long) As Boolean
   capCaptureSingleFrame = SendMessageAsLong(hCapWnd, WM_CAP_SINGLE_FRAME, 0&, 0&)
End Function
Function capCaptureGetSetup(ByVal hCapWnd As Long, ByRef capParms As CAPTUREPARMS) As Boolean
   capCaptureGetSetup = SendMessageAsAny(hCapWnd, WM_CAP_GET_SEQUENCE_SETUP, Len(capParms), capParms)
End Function
Function capCaptureSetSetup(ByVal hCapWnd As Long, ByRef capParms As CAPTUREPARMS) As Boolean
   capCaptureSetSetup = SendMessageAsAny(hCapWnd, WM_CAP_SET_SEQUENCE_SETUP, Len(capParms), capParms)
End Function
Function capSetMCIDeviceName(ByVal hCapWnd As Long, ByVal DeviceName As String) As Boolean
   'DeviceName = DeviceName & Chr$(0) 'null-terminate the string just to be safe
   capSetMCIDeviceName = SendMessageAsString(hCapWnd, WM_CAP_SET_MCI_DEVICE, 0&, DeviceName)
End Function
Function capGetMCIDeviceName(ByVal hCapWnd As Long) As String
   'returns device name as VB string (default name is "")
   Dim dwSize As Long
   Dim szBuffer As String
   
   dwSize = 128 'MCISTRING_MAX
   szBuffer = String$(dwSize, 0)
   Call SendMessageAsString(hCapWnd, WM_CAP_GET_MCI_DEVICE, dwSize, szBuffer)
   capGetMCIDeviceName = Left$(szBuffer, InStr(szBuffer, vbNullChar) - 1)
End Function
Function capPaletteOpen(ByVal hCapWnd As Long, ByVal FilePath As String) As Boolean
   capPaletteOpen = SendMessageAsString(hCapWnd, WM_CAP_PAL_OPEN, 0&, FilePath)
End Function
Function capPaletteSave(ByVal hCapWnd As Long, ByVal FilePath As String) As Boolean
   capPaletteSave = SendMessageAsString(hCapWnd, WM_CAP_PAL_SAVE, 0&, FilePath)
End Function
Function capPalettePaste(ByVal hCapWnd As Long) As Boolean
   capPalettePaste = SendMessageAsLong(hCapWnd, WM_CAP_PAL_PASTE, 0&, 0&)
End Function
Function capPaletteAuto(ByVal hCapWnd As Long, ByVal iFrames As Long, ByVal iColors As Long) As Boolean
   'iColors should not be greater than 256
   Debug.Assert iColors < 257
   capPaletteAuto = SendMessageAsLong(hCapWnd, WM_CAP_PAL_AUTOCREATE, iFrames, iColors)
End Function
Function capPaletteManual(ByVal hCapWnd As Long, ByVal f As Boolean, ByVal iColors As Long) As Boolean
   'iColors should not be greater than 256
   Debug.Assert iColors < 257
   capPaletteManual = SendMessageAsLong(hCapWnd, WM_CAP_PAL_MANUALCREATE, -(f), iColors)
End Function

