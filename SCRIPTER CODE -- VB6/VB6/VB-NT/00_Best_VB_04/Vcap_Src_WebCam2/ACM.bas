Attribute VB_Name = "mACM"
'****************************************************************
'*  VB file:   ACM.bas... VB32 wrapper for Win32 ACM functions
'*
'*  created:        1998 by Ray Mercer
'*  last modified:  12/2/98 by Ray Mercer (added comments)
'*
'*
'*  Some useful defines and functions for using the Windows
'*  Audio Compression Manager from Visual Basic
'*
'*  Copyright (c) 1998 Ray Mercer.  All rights reserved.
'****************************************************************

Option Explicit

'WAVEFORMATEX is from MS Voice sample on VB5 CD
Public Const MAXEXTRABYTES = 3          ' Maximum (Extra Bytes + 1) In Non PCM Wave Formats...
Type WAVEFORMATEX
    wFormatTag As Integer       ' format type
    nChannels As Integer        ' number of channels (i.e. mono, stereo, etc.)
    nSamplesPerSec As Long      ' sample rate
    nAvgBytesPerSec As Long     ' for buffer estimation
    nBlockAlign As Integer      ' block size of data
    wBitsPerSample As Integer   ' Bits Per Sample
    cbSize As Integer           ' Size Of (FACT CHUNK)
    xBytes(MAXEXTRABYTES) As Byte ' (FACT CHUNK)
End Type

Type WAVEHDR
    lpData As Long              ' pointer to locked data buffer
    dwBufferLength As Long      ' length of data buffer
    dwBytesRecorded As Long     ' used for input only
    dwUser As Long              ' for client's use
    dwFlags As Long             ' assorted flags (see defines)
    dwLoops As Long             ' loop control counter
    wavehdr_tag As Long         ' reserved for driver
    Reserved As Long            ' reserved for driver
    hData As Long               ' handle to locked data buffer
End Type


'//AUDIO DEFINES (from msacm.h)
Public Const ACMFORMATTAGDETAILS_FORMATTAG_CHARS As Long = 48
Public Const ACMFORMATDETAILS_FORMAT_CHARS As Long = 128
Public Type ACMFORMATTAGDETAILSTRUCT
    cbStruct As Long
    dwFormatTagIndex As Long
    dwFormatTag As Long
    cbFormatSize As Long
    fdwSupport As Long
    cStandardFormats As Long
    szFormatTag As String * ACMFORMATTAGDETAILS_FORMATTAG_CHARS
End Type
Public Declare Function acmFormatTagDetails Lib "msacm32.dll" Alias "acmFormatTagDetailsA" _
                                (ByVal had As Long, _
                                    ByRef paftd As ACMFORMATTAGDETAILSTRUCT, _
                                    ByVal fdwDetails As Long) As Long  'returns MMResult
Public Const ACMFORMATCHOOSE_STYLEF_INITTOWFXSTRUCT  As Long = &H40&
Public Type ACMFORMATCHOOSESTRUCT
    cbStruct As Long
    fdwStyle As Long
    hwndOwner As Long
    pwfx As Long ' can be WAVEFORMATEX
    cbwfx As Long
    pszTitle As String
    szFormatTag As String * ACMFORMATTAGDETAILS_FORMATTAG_CHARS
    szFormat As String * ACMFORMATDETAILS_FORMAT_CHARS
    pszName As String
    cchName As Long
    fdwEnum As Long
    pwfxEnum As Long ' can be a WAVEFORMATEX
    hInstance As Long
    pszTemplateName As Long 'can be a string
    lCustData As Long
    pfnHook As Long
End Type
Public Declare Function acmFormatChoose Lib "msacm32.dll" Alias "acmFormatChooseA" _
                                        (ByRef acmFormat As ACMFORMATCHOOSESTRUCT) As Long  ' returns MMResult

Public Const ACM_METRIC_MAX_SIZE_FORMAT  As Long = 50
Public Const ACM_FORMATENUMF_HARDWARE    As Long = &H400000
Public Const ACM_FORMATENUMF_INPUT      As Long = &H800000
Public Declare Function acmMetrics Lib "msacm32.dll" _
                                        (ByVal hao As Long, _
                                        ByVal uMetric As Long, _
                                        ByVal pMetric As Long) As Long ' returns MMResult

'****************************************************************
'* FUNCTION SetAudioFormatDlg()
'* ===============
'* By Ray Mercer, 1998
'*
'* Displays a system-defined dialog which allows the
'* user to select an ACM audio format
'*
'* INPUTS:
'* ownerHwnd - hWnd of parent window or 0& if no modality is desired
'*
'* RETURNS:
'* <none>
'****************************************************************
Sub SetAudioFormatDlg(ByVal ownerHwnd As Long)

Dim retVal As Long
Dim dwSize As Long
Dim acmFormat As ACMFORMATCHOOSESTRUCT
Dim hFormatData As Long 'handle to memory location
Dim pFormatData As Long 'pointer to locked memory buffer
Dim maxLenWavFormat As Long
Dim curLenWavFormat As Long
Dim formatTag As String
Dim formatAttributes As String

'alloc string buffers for format tags (note- I haven't implemented saving and restoring these
' because this doesn't seem to be implemented in VidCap32 either - excercise for reader :-)
formatTag = String$(ACMFORMATTAGDETAILS_FORMATTAG_CHARS, 0)
formatAttributes = String$(ACMFORMATDETAILS_FORMAT_CHARS, 0)
' Ask the ACM what the largest wave format is.....
retVal = acmMetrics(0&, ACM_METRIC_MAX_SIZE_FORMAT, VarPtr(maxLenWavFormat))
If MMSYSERR_NOERROR = retVal Then
    'allocate mem "C-style" (this is easier than working with huge structs)
    hFormatData = GlobalAlloc(GMEM_MOVEABLE Or GMEM_SHARE Or GMEM_ZEROINIT, maxLenWavFormat) ' Allocate Global Memory
    Debug.Assert hFormatData <> 0
    pFormatData = GlobalLock(hFormatData)    ' Lock Memory handle
    Debug.Assert pFormatData <> 0
    'Get the current audio format size
    curLenWavFormat = capGetAudioFormatSize(frmMain.capwnd)
    Debug.Assert curLenWavFormat <> 0
    'Get the current audio format
    retVal = capGetAudioFormatAsArray(frmMain.capwnd, pFormatData, curLenWavFormat)
    With acmFormat
        .cbStruct = Len(acmFormat)
        .fdwStyle = ACMFORMATCHOOSE_STYLEF_INITTOWFXSTRUCT
        .fdwEnum = ACM_FORMATENUMF_HARDWARE Or ACM_FORMATENUMF_INPUT
        .hwndOwner = ownerHwnd
        .pwfx = pFormatData
        .cbwfx = curLenWavFormat
        .pszTitle = "Select Sound Format"
        .szFormatTag = formatTag  'returns valid data but not used (see note on format tag details above)
        .szFormat = formatAttributes 'ditto
        .hInstance = 0&  'careful! this should be null unless you are using templates
    End With
    retVal = acmFormatChoose(acmFormat)
    If MMSYSERR_NOERROR = retVal Then
        Call CopyPTRtoLONG(dwSize, pFormatData + 16, 4) 'pull the length field out of the byte array
        retVal = capSetAudioFormatAsArray(frmMain.capwnd, pFormatData, dwSize)
        'macros return FALSE on error
        If 0 = retVal Then MsgBox "Could not set new audio format", vbInformation, App.Title
    End If
    'free the locked mem
    Call GlobalUnlock(pFormatData)
    Call GlobalFree(hFormatData)
Else
    MsgBox "Windows ACM returned an error", vbInformation, App.Title
End If

End Sub
