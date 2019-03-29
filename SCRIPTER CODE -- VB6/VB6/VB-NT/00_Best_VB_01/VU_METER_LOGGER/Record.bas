Attribute VB_Name = "Record"

Option Explicit
'Option Compare Text
Public PEAKT1, PEAKT2

Public FS
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const FOREVER As Long = 24000000
Public Const sE As String = ""
Public Const sQ As String = """"
Public Const AudioLog As String = "\Audio_Logg\"

Public BitRate As Integer
Public Capturing As Boolean
Public Writing As Boolean
Public sBuff() As Integer
Public ABufferIsFull As Boolean
Public DurTot As Double
Public Old As Double
Public TimeStart As Date
Public SecSilence As Double
Public dB As Double
Public DiskMin As Double
Public KillIfFull As Boolean
Public KillOld As Boolean
Public Chan As Integer
Public MinuteChange As Long
Public PauseOnSilence As Boolean
Public Auto As Boolean
Public LogFolder As String

Private Const WAVE_FORMAT_PCM = 1
Private Const WAVE_MAPPER = -1&

Private Const CALLBACK_FUNCTION = &H30000
Private Const MM_WIM_DATA = &H3C0
Private Const WIM_DATA = MM_WIM_DATA
Public Const WHDR_DONE = &H1         '  done bit
Private Const GMEM_FIXED = &H0         ' Global Memory Flag used by GlobalAlloc functin
Public Const NUM_BUFFERS = 1
Private Const DEVICEID = WAVE_MAPPER
Private Const MMSYSERR_NOERROR = 0

Private Type WaveInCaps
    ManufacturerID As Integer      'wMid
    ProductID As Integer           'wPid
    DriverVersion As Long          'MMVERSIONS vDriverVersion
    ProductName(1 To 32) As Byte   'szPname[MAXPNAMELEN]
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type

Private Type WaveHdr
    lpData As Long          ' Address of the waveform buffer.
    dwBufferLength As Long  ' Length, in bytes, of the buffer.
    dwBytesRecorded As Long ' When the header is used in input, this member specifies how much
                            ' data is in the buffer.
    dwUser As Long          ' User data.
    dwFlags As Long         ' Flags supplying information about the buffer. Set equal to zero.
    dwLoops As Long         ' Number of times to play the loop. Set equal to zero.
    lpNext As Long          ' Not used
    Reserved As Long        ' Not used
End Type

Private Type WaveFormatEx
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize As Integer
End Type

Private Declare Function waveInGetNumDevs Lib "winmm" () As Long
Private Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long

Private Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WaveFormatEx, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WaveHdr, ByVal uSize As Long) As Long
Private Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WaveHdr, ByVal uSize As Long) As Long
Private Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Public Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WaveHdr, ByVal uSize As Long) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public BufferSize As Long
Private i As Long
Public rc As Long
Private msg As String * 200

Private hReadPipe As Long
Private hWritePipe As Long
Private NumBytesDone As Long

Public hWaveIn As Long
Private wformat As WaveFormatEx
Private hmem(1 To NUM_BUFFERS) As Long
Public inHdr(1 To NUM_BUFFERS) As WaveHdr
'
Private Const STD_ERROR_HANDLE = -12&
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const MAX_PATH = 260

Private Declare Function CreatePipe Lib "kernel32" _
    (phReadPipe As Long, _
    phWritePipe As Long, _
    ByVal lpPipeAttributes As Long, _
    ByVal nSize As Long) As Long

Private Declare Function SetStdHandle Lib "kernel32" _
    (ByVal nStdHandle As Long, _
    ByVal nHandle As Long) As Long

Private Declare Function WriteFile Lib "kernel32" _
    (ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Type Int64
    Lo As Long
    Hi As Long
End Type

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Int64, lpTotalNumberOfBytes As Int64, lpTotalNumberOfFreeBytes As Int64) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBpSor As Long, lpNumberOfFreeClers As Long, lpTotalNumberOfClusters As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" _
    Alias "GetShortPathNameA" ( _
    ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long _
    ) As Long
'


Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long

Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40




Public Sub DelTree(ByVal DelFolder As String)
  
  Dim BufLen As Long
  Dim RetVal As Long
  Dim ShortFolder As String
  
  On Error GoTo ExitSub
  
  
  BufLen = MAX_PATH
  ShortFolder = Space$(BufLen)
  RetVal = GetShortPathName(DelFolder, ShortFolder, BufLen)
  ShortFolder = Left$(ShortFolder, RetVal)
  
  'On Error Resume Next
  'Shell "command.com /c deltree /y " & ShortFolder, vbHide
  'Shell "cmd.exe /c rmdir /s /q " & ShortFolder, vbHide
  'On Error GoTo 0
  
ExitSub:
End Sub

Sub Main()
  
  Dim ArrCmd As Variant
  Dim i As Integer
  Dim C As String
  Dim tmp As String
  Dim Pos As Integer
  
  Capturing = False
  Writing = False
  Auto = False
  LogFolder = App.Path & AudioLog
  LogFolder = "D:\MP3REC" & AudioLog
  BitRate = 8
  Chan = 1
  TimeStart = Now
  DurTot = FOREVER
  MinuteChange = 5
  PauseOnSilence = True
  dB = -40
  SecSilence = 10
  DiskMin = 100
  KillIfFull = True
  KillOld = True
  Old = 90
  
  ArrCmd = Split(Command, "/")
  
  For i = LBound(ArrCmd) To UBound(ArrCmd)
    C = Trim$(ArrCmd(i))
    tmp = sE
    If Left$(C, 4) = "auto" Then
      '/auto = Enter Auto Mode. Default: Auto off.
      Auto = True
    End If
    If Left$(C, 3) = "log" Then
      '/logx = logfolder. Default: AppPath\AudioLog\
      If Len(C) > 3 Then
        tmp = Trim$(Mid$(C, 4))
        If Right$(tmp, 1) <> "\" Then tmp = tmp & "\"
        Pos = InStr(1, tmp, AudioLog)
        If Pos > 0 Then tmp = Left$(tmp, Pos - 1)
        If ExistFile(tmp) Then
          If GetAttr(tmp) And vbDirectory Then
            LogFolder = tmp & Mid$(AudioLog, 2)
          End If
        End If
      End If
    End If
    If Left$(C, 2) = "br" Then
      '/brx = BitRate x kbits/sec. Default: 8
      If Len(C) > 2 Then tmp = Mid$(C, 3)
      If IsNumeric(tmp) Then BitRate = CInt(tmp)
    End If
    If Left$(C, 2) = "st" Then
      '/st = Stereo. Default: Mono
      Chan = 2
    End If
    If Left$(C, 3) = "dur" Then
      ' /durx = total Duration of Writing x Hours. Default: Forever
      If Len(C) > 3 Then tmp = Mid$(C, 4)
      If IsNumeric(tmp) Then DurTot = CDbl(tmp)
    End If
    If Left$(C, 2) = "nf" Then
      ' /nfx = New File every x minutes.Default: 5
      If Len(C) > 2 Then tmp = Mid$(C, 3)
      If IsNumeric(tmp) Then MinuteChange = CInt(tmp)
    End If
    If Left$(C, 2) = "dp" Then
      '/dp = Don't Pause on silence. Default: Pause on silence
      PauseOnSilence = False
    End If
    If Left$(C, 2) = "db" Then
      ' /dbx = x dB limit for pause. Default: -40
      If Len(C) > 2 Then tmp = Mid$(C, 3)
      If IsNumeric(tmp) Then dB = -Abs(CDbl(tmp))
    End If
    If Left$(C, 3) = "sec" Then
      ' /secx = x Seconds limit for pause. Default: 10
      If Len(C) > 3 Then tmp = Mid$(C, 4)
      If IsNumeric(tmp) Then SecSilence = CDbl(tmp)
    End If
    If Left$(C, 4) = "disk" Then
      ' /diskx = Min Disk limit x MB. Default: 100
      If Len(C) > 4 Then tmp = Mid$(C, 5)
      If IsNumeric(tmp) Then DiskMin = CDbl(tmp)
    End If
    If Left$(C, 2) = "sf" Then
      '/sf = Stop if disk is Full. Default: Delete oldest folder.
      KillIfFull = False
    End If
    If Left$(C, 2) = "dk" Then
      '/dk = Don't Kill old files. Default: Kill old files
      KillOld = False
    End If
    If Left$(C, 3) = "old" Then
      ' /oldx = Kill folders Older than x days. Default: 90
      If Len(C) > 3 Then tmp = Mid$(C, 4)
      If IsNumeric(tmp) Then Old = CDbl(tmp)
    End If
  Next
  
    frmMain.Show
    DoEvents

End Sub

Public Function StartCapture(ByVal Ch As Integer, _
                          ByVal Sps As Integer) _
                          As Boolean
  Dim Caps As WaveInCaps
  Dim CFrmt As Long
  Dim BufTime As Double
  
  If Capturing Then
    StartCapture = True
    Exit Function
  Else
    StartCapture = False
  End If
  
  If Sps < 2 Then Sps = 1
  If Sps > 2 Then Sps = 4
  
  With wformat
    .wFormatTag = WAVE_FORMAT_PCM
    .nChannels = Ch
    .nSamplesPerSec = Sps * 11025&
    .wBitsPerSample = 16
    .nBlockAlign = (.nChannels * .wBitsPerSample) \ 8
    .nAvgBytesPerSec = .nBlockAlign * .nSamplesPerSec
    .cbSize = 0
  
'WAVE_FORMAT_1M08 = &H1&   11.025 kHz, Mono,   8-bit
'WAVE_FORMAT_1S08 = &H2&   11.025 kHz, Stereo, 8-bit
'WAVE_FORMAT_1M16 = &H4&   11.025 kHz, Mono,   16-bit
'WAVE_FORMAT_1S16 = &H8&   11.025 kHz, Stereo, 16-bit
'WAVE_FORMAT_2M08 = &H10&  22.05  kHz, Mono,   8-bit
'WAVE_FORMAT_2S08 = &H20&  22.05  kHz, Stereo, 8-bit
'WAVE_FORMAT_2M16 = &H40&  22.05  kHz, Mono,   16-bit
'WAVE_FORMAT_2S16 = &H80&  22.05  kHz, Stereo, 16-bit
'WAVE_FORMAT_4M08 = &H100& 44.1   kHz, Mono,   8-bit
'WAVE_FORMAT_4S08 = &H200& 44.1   kHz, Stereo, 8-bit
'WAVE_FORMAT_4M16 = &H400& 44.1   kHz, Mono,   16-bit
'WAVE_FORMAT_4S16 = &H800& 44.1   kHz, Stereo, 16-bit
    CFrmt = .nChannels
    If .wBitsPerSample = 16 Then CFrmt = 4 * CFrmt
    If .nSamplesPerSec = 22050 Then CFrmt = &H10& * CFrmt
    If .nSamplesPerSec = 44100 Then CFrmt = &H100& * CFrmt
    
    BufTime = 0.1
    BufferSize = .nBlockAlign * CLng(BufTime * .nSamplesPerSec)
                                   ' 0.1 seconds
    
  End With
  
  ReDim sBuff(1 To BufferSize \ 2) ' 2 bytes/integer
  
  rc = waveInGetDevCaps(DEVICEID, VarPtr(Caps), Len(Caps))
  If rc <> 0 Then
    waveInGetErrorText rc, msg, Len(msg)
    MsgBox msg
    Exit Function
  End If
  
  ' Check if device can do this format
  If Caps.Formats And CFrmt Then
  Else
    MsgBox "The default audio Capture device can't do this"
    Exit Function
  End If
  
  For i = 1 To NUM_BUFFERS
    hmem(i) = GlobalAlloc(&H40, BufferSize) 'BufferSize
    inHdr(i).lpData = GlobalLock(hmem(i))
    inHdr(i).dwBufferLength = BufferSize
    inHdr(i).dwFlags = 0
    inHdr(i).dwLoops = 0
  Next

  rc = waveInOpen(hWaveIn, DEVICEID, wformat, _
              AddressOf waveInProc, 0, CALLBACK_FUNCTION)
  If rc <> 0 Then
    waveInGetErrorText rc, msg, Len(msg)
    MsgBox msg
    Exit Function
  End If

  For i = 1 To NUM_BUFFERS
    rc = waveInPrepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
    If rc <> 0 Then
      waveInGetErrorText rc, msg, Len(msg)
      MsgBox msg
    End If
  Next
  
  For i = 1 To NUM_BUFFERS
    
    rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
    If rc <> 0 Then
      waveInGetErrorText rc, msg, Len(msg)
      MsgBox msg
    End If
    
    CopyMemory sBuff(1), ByVal inHdr(i).lpData, BufferSize
    
  Next
  
  rc = waveInStart(hWaveIn)
  If rc <> 0 Then
    waveInGetErrorText rc, msg, Len(msg)
    MsgBox msg
  End If
  
  StartCapture = True
  Capturing = True
  
  ABufferIsFull = False
  frmMain.Timer1.Interval = 1000 * BufTime
  
End Function

Public Function StartWrite() As Boolean
  
  Dim Options As String
  Dim DestFile As String
  Dim DateTag As String
  Dim TimeTag As String
  Dim YearTag As String
  Dim Ahora As Date
  
  If Writing Then
    StartWrite = True
    Exit Function
  Else
    StartWrite = False
  End If
  
  If Not Capturing Then Exit Function
  
  Writing = False
  
  Ahora = Now
  YearTag = Format$(Ahora, "yyyy")
  DateTag = Format$(Ahora, "yyyy-mm-dd")
  TimeTag = Format$(Ahora, "HH-NN-SS")
  
  DestFile = LogFolder
  If Not ExistFile(DestFile) Then MkDir DestFile
  DestFile = DestFile & DateTag & "\"
  If Not ExistFile(DestFile) Then MkDir DestFile
  DestFile = DestFile & TimeTag & ".mp3"

  With wformat
    
    Options = "-r -x -h -s " & CStr(.nSamplesPerSec / 1000)
      'raw, swapbyte, hi qual, SampleRate
    Options = Options & " -b " & CStr(BitRate) 'Bitrate
'    Options = Options & " --bitwidth " & CStr(.wBitsPerSample)
    
    If .nChannels = 1 Then
      Options = Options & " -m m" 'Mono
    ElseIf BitRate > 128 Then
      Options = Options & " -m s" 'Stereo
    Else
      Options = Options & " -m j" 'Joint Stereo
    End If
    
    Options = Options & " --ta " & "Date:_" & DateTag
    Options = Options & " --tt " & "Time:_" & TimeTag
    Options = Options & " --tl " & "AudioLog_Mp3Rec"
    Options = Options & " --tc " & "http://www.heimskringla.com/"
    Options = Options & " --ty " & YearTag
    Options = Options & " --tg " & "Soundtrack"
    
  End With
  
  rc = CreatePipe(hReadPipe, hWritePipe, 0&, BufferSize)
  If rc = 0 Then MsgBox "CreatePipe failed": Exit Function
  
  rc = SetStdHandle(STD_INPUT_HANDLE, hReadPipe)
  If rc = 0 Then MsgBox "SetStdHandle hReadPipe failed": Exit Function
  
  rc = Shell(App.Path + "\lame.exe " & Options _
                          & " - " & DestFile, vbMinimizedNoFocus)
  If rc = 0 Then MsgBox "Error starting lame": Exit Function
  
  Do While FindWindow("tty", "lame") = 0 And _
                        FindWindow("ConsoleWindowClass", _
                        App.Path & "\lame.exe") = 0
    If Now - Ahora > 5 / 86400 Then
      MsgBox "Can't find lame at startup"
      Exit Function
    End If
  Loop
  
  Writing = True
  
  StartWrite = True
  
End Function

Public Function StopWrite()
  
  Writing = False
  
  rc = CloseHandle(hWritePipe)
'  If rc = 0 Then MsgBox "CloseHandle(hWritePipe) failed"
  rc = CloseHandle(hReadPipe)
'  If rc = 0 Then MsgBox "CloseHandle(hReadPipe) failed"

End Function

Public Sub StopCapture()

  If Writing Then StopWrite

  Capturing = False
  
  rc = waveInReset(hWaveIn)
  rc = waveInStop(hWaveIn)
  
  For i = 1 To NUM_BUFFERS
    rc = waveInUnprepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
    GlobalFree hmem(i)
  Next
  
  rc = waveInClose(hWaveIn)
  
  frmMain.Timer1.Interval = 1000
  
End Sub

Private Sub waveInProc(ByVal hwi As Long, _
                        ByVal uMsg As Long, _
                        ByVal dwInstance As Long, _
                        ByVal dwParam1 As Long, _
                        ByVal dwParam2 As Long)

  If uMsg = WIM_DATA Then
    ABufferIsFull = True
  Else
    ABufferIsFull = False
  End If
  
End Sub

Public Sub ReadBufferBAS()

  Static LB As Long
  Dim k As Long
  
  k = LB
  For i = 1 To NUM_BUFFERS
    k = k + 1
    If k > NUM_BUFFERS Then k = 1
    
    If inHdr(k).dwFlags And WHDR_DONE Then
      LB = k
      rc = waveInAddBuffer(hWaveIn, inHdr(k), Len(inHdr(k)))
      If rc <> 0 Then
        MsgBox "waveInAddBuffer failed"
        Exit Sub
      End If

      CopyMemory sBuff(1), ByVal inHdr(k).lpData, BufferSize
      
  
  End If
Next
  
'THIS LOOP RUNS ONCE
  

'TRICKY
'IF WE CALL LIKE THIS ON ERROR WILL COME BACK TO THIS CALL
'Call frmMain.Vu

'IF WE CALL LIKE THIS ON ERROR WILL STOP IN THE SUB
'Call Vu
'THIS MEANS WILL HAVE THIS SUB IN THE FORM WHERE CALL GOES
    
    
    
ABufferIsFull = False

End Sub


Public Function ExistFile(TestPath As String) As Boolean

  If TestPath = sE Then
    ExistFile = False
    Exit Function
  End If
  
  On Error GoTo ErrFix
  ExistFile = (Dir$(TestPath, vbReadOnly Or vbDirectory) <> sE)
  On Error GoTo 0
  
Exit Function

ErrFix:
  Dim ErrNum As Integer
  ErrNum = err.Number
  If ErrNum >= 52 And ErrNum <= 76 Then
    ExistFile = False
    err.Clear
    Resume Next
  Else
    MsgBox "Error Number: " & ErrNum & vbCrLf & _
            err.Description
    End
  End If
  
End Function

Private Function CInt64(Lo As Long, Hi As Long) As Double
'This function converts the Int64 data type to a double
    Dim dblLo As Double, dblHi As Double
    If Lo < 0 Then dblLo = 2 ^ 32 + Lo Else dblLo = Lo
    If Hi < 0 Then dblHi = 2 ^ 32 + Hi Else dblHi = Hi
    CInt64 = dblLo + dblHi * 2 ^ 32
End Function

Public Function FreeSpaceInfo() As Double
  
  Dim Avail As Int64, Tot As Int64, Fr As Int64
  Dim SpCl As Long, BpS As Long, FreeCl As Long, TotCl As Long
  
  If GetProcAddress(GetModuleHandle("kernel32.dll"), _
                            "GetDiskFreeSpaceExA") <> 0 Then                                 '
  
    If GetDiskFreeSpaceEx(App.Path, Avail, Tot, Fr) Then
      FreeSpaceInfo = CInt64(Avail.Lo, Avail.Hi)
    Else
      FreeSpaceInfo = -1
      Debug.Print "GetDiskFreeSpaceEx failed"
    End If
    
  Else
    
    If GetDiskFreeSpace(Left$(App.Path, 3), SpCl, BpS, _
          FreeCl, TotCl) <> 0 Then
      FreeSpaceInfo = CDbl(FreeCl) * SpCl * BpS
    Else
      FreeSpaceInfo = -1
      Debug.Print "GetDiskFreeSpace failed"
    End If
    
  End If
  
End Function


'Type OFSTRUCT
'  cBytes     As Byte
'  fFixedDisk As Byte
'  nErrCode   As Integer
'  Reserved1  As Integer
'  Reserved2  As Integer
'  szPathName As String * 128
'End Type

'Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long

'Const OF_SHARE_COMPAT = &H0
'Const OF_SHARE_EXCLUSIVE = &H10
'Const OF_SHARE_DENY_WRITE = &H20
'Const OF_SHARE_DENY_READ = &H30
'Const OF_SHARE_DENY_NONE = &H40

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Function FileInUse(ByVal strFilePath As String) As Boolean
  
  Dim hFile As Long
  Dim FileInfo  As OFSTRUCT
  
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function
  If Right(strFilePath, 1) <> Chr(0) Then strFilePath = strFilePath & Chr(0)
  
  FileInfo.cBytes = Len(FileInfo)
  hFile = OpenFile(strFilePath, FileInfo, OF_SHARE_EXCLUSIVE)
  If hFile = -1 And err.LastDllError = 32 Then
    FileInUse = True
  Else
    CloseHandle hFile
  End If
  
End Function

Sub FileInUseDelay(Txt, Complain)
    
    Dim QQ
        
    QQ = Now + TimeSerial(0, 8, 0)
    Do While FileInUse(Txt) = True Or QQ < Now
        DoEvents
        Sleep (20)
    Loop
    
    If QQ < Now And Complain = True Then
        MsgBox "Trouble File Stuck In Use " + vbCrLf + Txt + vbCrLf + "Longer than 8 min"
        Stop
    End If
End Sub


