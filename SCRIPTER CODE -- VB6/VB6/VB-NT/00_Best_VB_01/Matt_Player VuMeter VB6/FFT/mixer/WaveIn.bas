Attribute VB_Name = "Module2"
' THis module contains all of the WAVE declares and constants.
' Author: Matt Gillmore (SCO_STINKS@YAHOO.com)

Option Explicit

Type WaveFormatEx
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Type WaveHdr
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    reserved As Long
End Type

Type Stereo
    L As Integer
    R As Integer
End Type

Type FFTStereo
    L As Single
    R As Single
End Type

Type WavBuf
    Hdr As WaveHdr
    Arr(1023) As Stereo '23 msec (43Hz and 4 kB Size)
    FFT(1023) As FFTStereo
End Type

Type WaveInCaps
    ManufacturerID As Integer      'wMid
    ProductID As Integer       'wPid
    DriverVersion As Long       'MMVERSIONS vDriverVersion
    ProductName(31) As Byte 'szPname[MAXPNAMELEN]
    Formats As Long
    Channels As Integer
    reserved As Integer
End Type

Public Const WAVE_INVALIDFORMAT = &H0&                 '/* invalid format */
Public Const WAVE_FORMAT_1M08 = &H1&                   '/* 11.025 kHz, Mono,   8-bit
Public Const WAVE_FORMAT_1S08 = &H2&                   '/* 11.025 kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_1M16 = &H4&                   '/* 11.025 kHz, Mono,   16-bit
Public Const WAVE_FORMAT_1S16 = &H8&                   '/* 11.025 kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_2M08 = &H10&                  '/* 22.05  kHz, Mono,   8-bit
Public Const WAVE_FORMAT_2S08 = &H20&                  '/* 22.05  kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_2M16 = &H40&                  '/* 22.05  kHz, Mono,   16-bit
Public Const WAVE_FORMAT_2S16 = &H80&                  '/* 22.05  kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_4M08 = &H100&                 '/* 44.1   kHz, Mono,   8-bit
Public Const WAVE_FORMAT_4S08 = &H200&                 '/* 44.1   kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_4M16 = &H400&                 '/* 44.1   kHz, Mono,   16-bit
Public Const WAVE_FORMAT_4S16 = &H800&                 '/* 44.1   kHz, Stereo, 16-bit

Public Const WAVE_FORMAT_PCM = 1

Public Const WHDR_DONE = &H1&              '/* done bit */
Public Const WHDR_PREPARED = &H2&          '/* set if this header has been prepared */
Public Const WHDR_BEGINLOOP = &H4&         '/* loop start block */
Public Const WHDR_ENDLOOP = &H8&           '/* loop end block */
Public Const WHDR_INQUEUE = &H10&          '/* reserved for driver */

Public Const WIM_OPEN = &H3BE
Public Const WIM_CLOSE = &H3BF
Public Const WIM_DATA = &H3C0

Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, WaveHdrPointer As WaveHdr, ByVal WaveHdrStructSize As Long) As Long
Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, WaveHdrPointer As WaveHdr, ByVal WaveHdrStructSize As Long) As Long
Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, WaveHdrPointer As WaveHdr, ByVal WaveHdrStructSize As Long) As Long

Declare Function waveInGetNumDevs Lib "winmm" () As Long
Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, WaveInCapsPointer As WaveInCaps, ByVal WaveInCapsStructSize As Long) As Long

Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, WaveFormatExPointer As WaveFormatEx, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal Flags As Long) As Long
Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
