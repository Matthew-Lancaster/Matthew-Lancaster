Attribute VB_Name = "AVICAP"
'*
'* Author: E. J. Bantz Jr.
'* Copyright: None, use and distribute freely ...
'* E-Mail: ej@bantz.com
'* Web: http://ej.bantz.com/video

'// ------------------------------------------------------------------
'//  Windows API Constants / Types / Declarations
'// ------------------------------------------------------------------
Public Const WM_USER = &H400
Private Type POINTAPI
        x As Long
        y As Long
End Type
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SendMessageS Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As String) As Long

'// ------------------------------------------------------------------
'//  Window Messages  WM_CAP... which can be sent to an AVICAP window
'// ------------------------------------------------------------------

'// Defines start of the message range
Public Const WM_CAP_START = WM_USER

Public Const WM_CAP_GET_CAPSTREAMPTR = WM_CAP_START + 1

Public Const WM_CAP_SET_CALLBACK_ERROR = WM_CAP_START + 2
Public Const WM_CAP_SET_CALLBACK_STATUS = WM_CAP_START + 3
Public Const WM_CAP_SET_CALLBACK_YIELD = WM_CAP_START + 4
Public Const WM_CAP_SET_CALLBACK_FRAME = WM_CAP_START + 5
Public Const WM_CAP_SET_CALLBACK_VIDEOSTREAM = WM_CAP_START + 6
Public Const WM_CAP_SET_CALLBACK_WAVESTREAM = WM_CAP_START + 7
Public Const WM_CAP_GET_USER_DATA = WM_CAP_START + 8
Public Const WM_CAP_SET_USER_DATA = WM_CAP_START + 9
    
Public Const WM_CAP_DRIVER_CONNECT = WM_CAP_START + 10
Public Const WM_CAP_DRIVER_DISCONNECT = WM_CAP_START + 11
Public Const WM_CAP_DRIVER_GET_NAME = WM_CAP_START + 12
Public Const WM_CAP_DRIVER_GET_VERSION = WM_CAP_START + 13
Public Const WM_CAP_DRIVER_GET_CAPS = WM_CAP_START + 14

Public Const WM_CAP_FILE_SET_CAPTURE_FILE = WM_CAP_START + 20
Public Const WM_CAP_FILE_GET_CAPTURE_FILE = WM_CAP_START + 21
Public Const WM_CAP_FILE_ALLOCATE = WM_CAP_START + 22
Public Const WM_CAP_FILE_SAVEAS = WM_CAP_START + 23
Public Const WM_CAP_FILE_SET_INFOCHUNK = WM_CAP_START + 24
Public Const WM_CAP_FILE_SAVEDIB = WM_CAP_START + 25

Public Const WM_CAP_EDIT_COPY = WM_CAP_START + 30

Public Const WM_CAP_SET_AUDIOFORMAT = WM_CAP_START + 35
Public Const WM_CAP_GET_AUDIOFORMAT = WM_CAP_START + 36

Public Const WM_CAP_DLG_VIDEOFORMAT = WM_CAP_START + 41
Public Const WM_CAP_DLG_VIDEOSOURCE = WM_CAP_START + 42
Public Const WM_CAP_DLG_VIDEODISPLAY = WM_CAP_START + 43
Public Const WM_CAP_GET_VIDEOFORMAT = WM_CAP_START + 44
Public Const WM_CAP_SET_VIDEOFORMAT = WM_CAP_START + 45
Public Const WM_CAP_DLG_VIDEOCOMPRESSION = WM_CAP_START + 46

Public Const WM_CAP_SET_PREVIEW = WM_CAP_START + 50
Public Const WM_CAP_SET_OVERLAY = WM_CAP_START + 51
Public Const WM_CAP_SET_PREVIEWRATE = WM_CAP_START + 52
Public Const WM_CAP_SET_SCALE = WM_CAP_START + 53
Public Const WM_CAP_GET_STATUS = WM_CAP_START + 54
Public Const WM_CAP_SET_SCROLL = WM_CAP_START + 55

Public Const WM_CAP_GRAB_FRAME = WM_CAP_START + 60
Public Const WM_CAP_GRAB_FRAME_NOSTOP = WM_CAP_START + 61

Public Const WM_CAP_SEQUENCE = WM_CAP_START + 62
Public Const WM_CAP_SEQUENCE_NOFILE = WM_CAP_START + 63
Public Const WM_CAP_SET_SEQUENCE_SETUP = WM_CAP_START + 64
Public Const WM_CAP_GET_SEQUENCE_SETUP = WM_CAP_START + 65
Public Const WM_CAP_SET_MCI_DEVICE = WM_CAP_START + 66
Public Const WM_CAP_GET_MCI_DEVICE = WM_CAP_START + 67
Public Const WM_CAP_STOP = WM_CAP_START + 68
Public Const WM_CAP_ABORT = WM_CAP_START + 69

Public Const WM_CAP_SINGLE_FRAME_OPEN = WM_CAP_START + 70
Public Const WM_CAP_SINGLE_FRAME_CLOSE = WM_CAP_START + 71
Public Const WM_CAP_SINGLE_FRAME = WM_CAP_START + 72

Public Const WM_CAP_PAL_OPEN = WM_CAP_START + 80
Public Const WM_CAP_PAL_SAVE = WM_CAP_START + 81
Public Const WM_CAP_PAL_PASTE = WM_CAP_START + 82
Public Const WM_CAP_PAL_AUTOCREATE = WM_CAP_START + 83
Public Const WM_CAP_PAL_MANUALCREATE = WM_CAP_START + 84

'// Following added post VFW 1.1
Public Const WM_CAP_SET_CALLBACK_CAPCONTROL = WM_CAP_START + 85

'// Defines end of the message range
Public Const WM_CAP_END = WM_CAP_SET_CALLBACK_CAPCONTROL

'// ------------------------------------------------------------------
'//  Structures
'// ------------------------------------------------------------------
Type CAPDRIVERCAPS
    wDeviceIndex As Long '               // Driver index in system.ini
    fHasOverlay As Long '                // Can device overlay?
    fHasDlgVideoSource As Long '         // Has Video source dlg?
    fHasDlgVideoFormat As Long '         // Has Format dlg?
    fHasDlgVideoDisplay As Long '        // Has External out dlg?
    fCaptureInitialized As Long '        // Driver ready to capture?
    fDriverSuppliesPalettes As Long '    // Can driver make palettes?
    hVideoIn As Long '                   // Driver In channel
    hVideoOut As Long '                  // Driver Out channel
    hVideoExtIn As Long '                // Driver Ext In channel
    hVideoExtOut As Long '               // Driver Ext Out channel
End Type

Type CAPSTATUS
    uiImageWidth As Long                    '// Width of the image
    uiImageHeight As Long                   '// Height of the image
    fLiveWindow As Long                     '// Now Previewing video?
    fOverlayWindow As Long                  '// Now Overlaying video?
    fScale As Long                          '// Scale image to client?
    ptScroll As POINTAPI                    '// Scroll position
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

Type CAPTUREPARMS
    dwRequestMicroSecPerFrame As Long       '// Requested capture rate
    fMakeUserHitOKToCapture As Long         '// Show "Hit OK to cap" dlg?
    wPercentDropForError As Long            '// Give error msg if > (10%)
    fYield As Long                          '// Capture via background task?
    dwIndexSize As Long                     '// Max index size in frames (32K)
    wChunkGranularity As Long               '// Junk chunk granularity (2K)
    fUsingDOSMemory As Long                 '// Use DOS buffers?
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
End Type

Type CAPINFOCHUNK
    fccInfoID As Long                       '// Chunk ID, "ICOP" for copyright
    lpData As Long                          '// pointer to data
    cbData As Long                          '// size of lpData
End Type

Type VIDEOHDR
    lpData As Long '// address of video buffer
    dwBufferLength As Long '// size, in bytes, of the Data buffer
    dwBytesUsed As Long '// see below
    dwTimeCaptured As Long '// see below
    dwUser As Long '// user-specific data
    dwFlags As Long '// see below
    dwReserved(3) As Long '// reserved; do not use}
End Type

'// The two functions exported by AVICap
Declare Function capCreateCaptureWindowA Lib "avicap32.dll" ( _
    ByVal lpszWindowName As String, _
    ByVal dwStyle As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Integer, _
    ByVal hWndParent As Long, ByVal nID As Long) As Long
Declare Function capGetDriverDescriptionA Lib "avicap32.dll" ( _
    ByVal wDriver As Integer, _
    ByVal lpszName As String, _
    ByVal cbName As Long, _
    ByVal lpszVer As String, _
    ByVal cbVer As Long) As Boolean

'// ------------------------------------------------------------------
'// String IDs from status and error callbacks
'// ------------------------------------------------------------------

Public Const IDS_CAP_BEGIN = 300              '/* "Capture Start" */
Public Const IDS_CAP_END = 301                '/* "Capture End" */

Public Const IDS_CAP_INFO = 401               '/* "%s" */
Public Const IDS_CAP_OUTOFMEM = 402           '/* "Out of memory" */
Public Const IDS_CAP_FILEEXISTS = 403         '/* "File '%s' exists -- overwrite it?" */
Public Const IDS_CAP_ERRORPALOPEN = 404       '/* "Error opening palette '%s'" */
Public Const IDS_CAP_ERRORPALSAVE = 405       '/* "Error saving palette '%s'" */
Public Const IDS_CAP_ERRORDIBSAVE = 406       '/* "Error saving frame '%s'" */
Public Const IDS_CAP_DEFAVIEXT = 407          '/* "avi" */
Public Const IDS_CAP_DEFPALEXT = 408          '/* "pal" */
Public Const IDS_CAP_CANTOPEN = 409           '/* "Cannot open '%s'" */
Public Const IDS_CAP_SEQ_MSGSTART = 410       '/* "Select OK to start capture\nof video sequence\nto %s." */
Public Const IDS_CAP_SEQ_MSGSTOP = 411        '/* "Hit ESCAPE or click to end capture" */
                
Public Const IDS_CAP_VIDEDITERR = 412         '/* "An error occurred while trying to run VidEdit." */
Public Const IDS_CAP_READONLYFILE = 413       '/* "The file '%s' is a read-only file." */
Public Const IDS_CAP_WRITEERROR = 414         '/* "Unable to write to file '%s'.\nDisk may be full." */
Public Const IDS_CAP_NODISKSPACE = 415        '/* "There is no space to create a capture file on the specified device." */
Public Const IDS_CAP_SETFILESIZE = 416        '/* "Set File Size" */
Public Const IDS_CAP_SAVEASPERCENT = 417      '/* "SaveAs: %2ld%%  Hit Escape to abort." */
                
Public Const IDS_CAP_DRIVER_ERROR = 418       '/* Driver specific error message */

Public Const IDS_CAP_WAVE_OPEN_ERROR = 419    '/* "Error: Cannot open the wave input device.\nCheck sample size, frequency, and channels." */
Public Const IDS_CAP_WAVE_ALLOC_ERROR = 420   '/* "Error: Out of memory for wave buffers." */
Public Const IDS_CAP_WAVE_PREPARE_ERROR = 421 '/* "Error: Cannot prepare wave buffers." */
Public Const IDS_CAP_WAVE_ADD_ERROR = 422     '/* "Error: Cannot add wave buffers." */
Public Const IDS_CAP_WAVE_SIZE_ERROR = 423    '/* "Error: Bad wave size." */
                
Public Const IDS_CAP_VIDEO_OPEN_ERROR = 424   '/* "Error: Cannot open the video input device." */
Public Const IDS_CAP_VIDEO_ALLOC_ERROR = 425  '/* "Error: Out of memory for video buffers." */
Public Const IDS_CAP_VIDEO_PREPARE_ERROR = 426 '/* "Error: Cannot prepare video buffers." */
Public Const IDS_CAP_VIDEO_ADD_ERROR = 427    '/* "Error: Cannot add video buffers." */
Public Const IDS_CAP_VIDEO_SIZE_ERROR = 428   '/* "Error: Bad video size." */
                
Public Const IDS_CAP_FILE_OPEN_ERROR = 429    '/* "Error: Cannot open capture file." */
Public Const IDS_CAP_FILE_WRITE_ERROR = 430   '/* "Error: Cannot write to capture file.  Disk may be full." */
Public Const IDS_CAP_RECORDING_ERROR = 431    '/* "Error: Cannot write to capture file.  Data rate too high or disk full." */
Public Const IDS_CAP_RECORDING_ERROR2 = 432   '/* "Error while recording" */
Public Const IDS_CAP_AVI_INIT_ERROR = 433     '/* "Error: Unable to initialize for capture." */
Public Const IDS_CAP_NO_FRAME_CAP_ERROR = 434 '/* "Warning: No frames captured.\nConfirm that vertical sync interrupts\nare configured and enabled." */
Public Const IDS_CAP_NO_PALETTE_WARN = 435    '/* "Warning: Using default palette." */
Public Const IDS_CAP_MCI_CONTROL_ERROR = 436  '/* "Error: Unable to access MCI device." */
Public Const IDS_CAP_MCI_CANT_STEP_ERROR = 437 '/* "Error: Unable to step MCI device." */
Public Const IDS_CAP_NO_AUDIO_CAP_ERROR = 438 '/* "Error: No audio data captured.\nCheck audio card settings." */
Public Const IDS_CAP_AVI_DRAWDIB_ERROR = 439  '/* "Error: Unable to draw this data format." */
Public Const IDS_CAP_COMPRESSOR_ERROR = 440   '/* "Error: Unable to initialize compressor." */
Public Const IDS_CAP_AUDIO_DROP_ERROR = 441   '/* "Error: Audio data was lost during capture, reduce capture rate." */
                
'/* status string IDs */
Public Const IDS_CAP_STAT_LIVE_MODE = 500      '/* "Live window" */
Public Const IDS_CAP_STAT_OVERLAY_MODE = 501   '/* "Overlay window" */
Public Const IDS_CAP_STAT_CAP_INIT = 502       '/* "Setting up for capture - Please wait" */
Public Const IDS_CAP_STAT_CAP_FINI = 503       '/* "Finished capture, now writing frame %ld" */
Public Const IDS_CAP_STAT_PALETTE_BUILD = 504  '/* "Building palette map" */
Public Const IDS_CAP_STAT_OPTPAL_BUILD = 505   '/* "Computing optimal palette" */
Public Const IDS_CAP_STAT_I_FRAMES = 506       '/* "%d frames" */
Public Const IDS_CAP_STAT_L_FRAMES = 507       '/* "%ld frames" */
Public Const IDS_CAP_STAT_CAP_L_FRAMES = 508   '/* "Captured %ld frames" */
Public Const IDS_CAP_STAT_CAP_AUDIO = 509      '/* "Capturing audio" */
Public Const IDS_CAP_STAT_VIDEOCURRENT = 510   '/* "Captured %ld frames (%ld dropped) %d.%03d sec." */
Public Const IDS_CAP_STAT_VIDEOAUDIO = 511     '/* "Captured %d.%03d sec.  %ld frames (%ld dropped) (%d.%03d fps).  %ld audio bytes (%d,%03d sps)" */
Public Const IDS_CAP_STAT_VIDEOONLY = 512      '/* "Captured %d.%03d sec.  %ld frames (%ld dropped) (%d.%03d fps)" */
Function capSetCallbackOnError(ByVal lwnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnError = SendMessage(lwnd, WM_CAP_SET_CALLBACK_ERROR, 0, lpProc)
End Function
Function capSetCallbackOnStatus(ByVal lwnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnStatus = SendMessage(lwnd, WM_CAP_SET_CALLBACK_STATUS, 0, lpProc)
End Function
Function capSetCallbackOnYield(ByVal lwnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnYield = SendMessage(lwnd, WM_CAP_SET_CALLBACK_YIELD, 0, lpProc)
End Function
Function capSetCallbackOnFrame(ByVal lwnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnFrame = SendMessage(lwnd, WM_CAP_SET_CALLBACK_FRAME, 0, lpProc)
End Function
Function capSetCallbackOnVideoStream(ByVal lwnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnVideoStream = SendMessage(lwnd, WM_CAP_SET_CALLBACK_VIDEOSTREAM, 0, lpProc)
End Function
Function capSetCallbackOnWaveStream(ByVal lwnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnWaveStream = SendMessage(lwnd, WM_CAP_SET_CALLBACK_WAVESTREAM, 0, lpProc)
End Function
Function capSetCallbackOnCapControl(ByVal lwnd As Long, ByVal lpProc As Long) As Boolean
   capSetCallbackOnCapControl = SendMessage(lwnd, WM_CAP_SET_CALLBACK_CAPCONTROL, 0, lpProc)
End Function
Function capSetUserData(ByVal lwnd As Long, ByVal lUser As Long) As Boolean
   capSetUserData = SendMessage(lwnd, WM_CAP_SET_USER_DATA, 0, lUser)
End Function
Function capGetUserData(ByVal lwnd As Long) As Long
   capGetUserData = SendMessage(lwnd, WM_CAP_GET_USER_DATA, 0, 0)
End Function
Function capDriverConnect(ByVal lwnd As Long, ByVal i As Integer) As Boolean
   capDriverConnect = SendMessage(lwnd, WM_CAP_DRIVER_CONNECT, i, 0)
End Function
Function capDriverDisconnect(ByVal lwnd As Long) As Boolean
   capDriverDisconnect = SendMessage(lwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0)
End Function
Function capDriverGetName(ByVal lwnd As Long, ByVal szName As Long, ByVal wSize As Integer) As Boolean
   capDriverGetName = SendMessage(lwnd, YOURCONSTANTMESSAGE, wSize, szName)
End Function
Function capDriverGetVersion(ByVal lwnd As Long, ByVal szVer As Long, ByVal wSize As Integer) As Boolean
   capDriverGetVersion = SendMessage(lwnd, WM_CAP_DRIVER_GET_VERSION, wSize, szVer)
End Function
Function capDriverGetCaps(ByVal lwnd As Long, ByVal S As Long, ByVal wSize As Integer) As Boolean
   capDriverGetCaps = SendMessage(lwnd, WM_CAP_DRIVER_GET_CAPS, wSize, S)

End Function
Function capFileSetCaptureFile(ByVal lwnd As Long, szName As String) As Boolean
   capFileSetCaptureFile = SendMessageS(lwnd, WM_CAP_FILE_SET_CAPTURE_FILE, 0, szName)
End Function
Function capFileGetCaptureFile(ByVal lwnd As Long, ByVal szName As Long, wSize As String) As Boolean
   capFileGetCaptureFile = SendMessageS(lwnd, WM_CAP_FILE_SET_CAPTURE_FILE, wSize, szName)
End Function
Function capFileAlloc(ByVal lwnd As Long, ByVal dwSize As Long) As Boolean
   capFileAlloc = SendMessage(lwnd, WM_CAP_FILE_ALLOCATE, 0, dwSize)
End Function
Function capFileSaveAs(ByVal lwnd As Long, szName As String) As Boolean
   capFileSaveAs = SendMessageS(lwnd, WM_CAP_FILE_SAVEAS, 0, szName)
End Function
Function capFileSetInfoChunk(ByVal lwnd As Long, ByVal lpInfoChunk As Long) As Boolean
   capFileSetInfoChunk = SendMessage(lwnd, WM_CAP_FILE_SET_INFOCHUNK, 0, lpInfoChunk)
End Function
Function capFileSaveDIB(ByVal lwnd As Long, ByVal szName As Long) As Boolean
   capFileSaveDIB = SendMessage(lwnd, WM_CAP_FILE_SAVEDIB, 0, szName)
End Function
Function capEditCopy(ByVal lwnd As Long) As Boolean
   capEditCopy = SendMessage(lwnd, WM_CAP_EDIT_COPY, 0, 0)
End Function
Function capSetAudioFormat(ByVal lwnd As Long, ByVal S As Long, ByVal wSize As Integer) As Boolean
   capSetAudioFormat = SendMessage(lwnd, WM_CAP_SET_AUDIOFORMAT, wSize, S)
End Function
Function capGetAudioFormat(ByVal lwnd As Long, ByVal S As Long, ByVal wSize As Integer) As Long
   capGetAudioFormat = SendMessage(lwnd, WM_CAP_GET_AUDIOFORMAT, wSize, S)
End Function
Function capGetAudioFormatSize(ByVal lwnd As Long) As Long
   capGetAudioFormatSize = SendMessage(lwnd, WM_CAP_GET_AUDIOFORMAT, 0, 0)
End Function
Function capDlgVideoFormat(ByVal lwnd As Long) As Boolean
   capDlgVideoFormat = SendMessage(lwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0)
End Function
Function capDlgVideoSource(ByVal lwnd As Long) As Boolean
   capDlgVideoSource = SendMessage(lwnd, WM_CAP_DLG_VIDEOSOURCE, 0, 0)
End Function
Function capDlgVideoDisplay(ByVal lwnd As Long) As Boolean
   capDlgVideoDisplay = SendMessage(lwnd, WM_CAP_DLG_VIDEODISPLAY, 0, 0)
End Function
Function capDlgVideoCompression(ByVal lwnd As Long) As Boolean
   capDlgVideoCompression = SendMessage(lwnd, WM_CAP_DLG_VIDEOCOMPRESSION, 0, 0)
End Function
Function capGetVideoFormat(ByVal lwnd As Long, ByVal S As Long, ByVal wSize As Integer) As Long
   capGetVideoFormat = SendMessage(lwnd, WM_CAP_GET_VIDEOFORMAT, wSize, S)
End Function
Function capGetVideoFormatSize(ByVal lwnd As Long) As Long
   capGetVideoFormatSize = SendMessage(lwnd, WM_CAP_GET_VIDEOFORMAT, 0, 0)
End Function
Function capSetVideoFormat(ByVal lwnd As Long, ByVal S As Long, ByVal wSize As Integer) As Boolean
   capSetVideoFormat = SendMessage(lwnd, WM_CAP_SET_VIDEOFORMAT, wSize, S)
End Function
Function capPreview(ByVal lwnd As Long, ByVal F As Boolean) As Boolean
   capPreview = SendMessage(lwnd, WM_CAP_SET_PREVIEW, F, 0)
End Function
Function capPreviewRate(ByVal lwnd As Long, ByVal wMS As Integer) As Boolean
   capPreviewRate = SendMessage(lwnd, WM_CAP_SET_PREVIEWRATE, wMS, 0)
End Function
Function capOverlay(ByVal lwnd As Long, ByVal F As Boolean) As Boolean
   capOverlay = SendMessage(lwnd, WM_CAP_SET_OVERLAY, F, 0)
End Function
Function capPreviewScale(ByVal lwnd As Long, ByVal F As Boolean) As Boolean
   capPreviewScale = SendMessage(lwnd, WM_CAP_SET_SCALE, F, 0)
End Function
Function capGetStatus(ByVal lwnd As Long, ByVal S As Long, ByVal wSize As Integer) As Boolean
   capGetStatus = SendMessage(lwnd, WM_CAP_GET_STATUS, wSize, S)
End Function
Function capSetScrollPos(ByVal lwnd As Long, ByVal lpP As Long) As Boolean
   capSetScrollPos = SendMessage(lwnd, WM_CAP_SET_SCROLL, 0, lpP)
End Function
Function capGrabFrame(ByVal lwnd As Long) As Boolean
   capGrabFrame = SendMessage(lwnd, WM_CAP_GRAB_FRAME, 0, 0)
End Function
Function capGrabFrameNoStop(ByVal lwnd As Long) As Boolean
   capGrabFrameNoStop = SendMessage(lwnd, WM_CAP_GRAB_FRAME_NOSTOP, 0, 0)
End Function
Function capCaptureSequence(ByVal lwnd As Long) As Boolean
   capCaptureSequence = SendMessage(lwnd, WM_CAP_SEQUENCE, 0, 0)
End Function
Function capCaptureSequenceNoFile(ByVal lwnd As Long) As Boolean
   capCaptureSequenceNoFile = SendMessage(lwnd, WM_CAP_SEQUENCE_NOFILE, 0, 0)
End Function
Function capCaptureStop(ByVal lwnd As Long) As Boolean
   capCaptureStop = SendMessage(lwnd, WM_CAP_STOP, 0, 0)
End Function
Function capCaptureAbort(ByVal lwnd As Long) As Boolean
   capCaptureAbort = SendMessage(lwnd, WM_CAP_ABORT, 0, 0)
End Function
Function capCaptureSingleFrameOpen(ByVal lwnd As Long) As Boolean
   capCaptureSingleFrameOpen = SendMessage(lwnd, WM_CAP_SINGLE_FRAME_OPEN, 0, 0)
End Function
Function capCaptureSingleFrameClose(ByVal lwnd As Long) As Boolean
   capCaptureSingleFrameClose = SendMessage(lwnd, WM_CAP_SINGLE_FRAME_CLOSE, 0, 0)
End Function
Function capCaptureSingleFrame(ByVal lwnd As Long) As Boolean
   capCaptureSingleFrame = SendMessage(lwnd, WM_CAP_SINGLE_FRAME, 0, 0)
End Function
Function capCaptureGetSetup(ByVal lwnd As Long, ByVal S As Long, ByVal wSize As Integer) As Boolean
   capCaptureGetSetup = SendMessage(lwnd, WM_CAP_GET_SEQUENCE_SETUP, wSize, S)
End Function
Function capCaptureSetSetup(ByVal lwnd As Long, ByVal S As Long, ByVal wSize As Integer) As Boolean
   capCaptureSetSetup = SendMessage(lwnd, WM_CAP_SET_SEQUENCE_SETUP, wSize, S)
End Function
Function capSetMCIDeviceName(ByVal lwnd As Long, ByVal szName As Long) As Boolean
   capSetMCIDeviceName = SendMessage(lwnd, WM_CAP_SET_MCI_DEVICE, 0, szName)
End Function
Function capGetMCIDeviceName(ByVal lwnd As Long, ByVal szName As Long, ByVal wSize As Integer) As Boolean
   capGetMCIDeviceName = SendMessage(lwnd, WM_CAP_GET_MCI_DEVICE, wSize, szName)
End Function
Function capPaletteOpen(ByVal lwnd As Long, ByVal szName As Long) As Boolean
   capPaletteOpen = SendMessage(lwnd, WM_CAP_PAL_OPEN, 0, szName)
End Function
Function capPaletteSave(ByVal lwnd As Long, ByVal szName As Long) As Boolean
   capPaletteSave = SendMessage(lwnd, WM_CAP_PAL_SAVE, 0, szName)
End Function
Function capPalettePaste(ByVal lwnd As Long) As Boolean
   capPalettePaste = SendMessage(lwnd, WM_CAP_PAL_PASTE, 0, 0)
End Function
Function capPaletteAuto(ByVal lwnd As Long, ByVal iFrames As Integer, ByVal iColor As Long) As Boolean
   capPaletteAuto = SendMessage(lwnd, WM_CAP_PAL_AUTOCREATE, iFrames, iColors)
End Function
Function capPaletteManual(ByVal lwnd As Long, ByVal fGrab As Boolean, ByVal iColors As Long) As Boolean
   capPaletteManual = SendMessage(lwnd, WM_CAP_PAL_MANUALCREATE, fGrab, iColors)
End Function

