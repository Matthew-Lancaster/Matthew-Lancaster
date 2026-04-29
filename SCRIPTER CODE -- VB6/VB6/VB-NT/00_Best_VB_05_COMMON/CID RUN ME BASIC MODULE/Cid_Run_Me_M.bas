Attribute VB_Name = "Cid_Run_Me_M"
Public Idle_Timer

'Put This in a Bas Mobule
'Public m_CRC As clsCRC

Public FS, F, RamDrive
Dim TTFArray(30)

Public Declare Function Putfocus _
        Lib "user32" _
        Alias "SetFocus" _
        (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public CurProcHwnd As Long, MustUnload

'Public Declare Function FindWindowDLL Lib "user32" Alias "FindWindowA" _
'   (ByVal lpClassName As String, _
'     ByVal lpWindowName As String) _
'    As Long
    
    
Public NoteHard1()
Public NoteHard2()

Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Public ScreenTwipsX, ScreenTwipsY, ScreenWidthX, ScreenHeightY, Idle_Timer_Proc

Public VcodeTT$
Public KeyLoggDate
Public WinampArray(30)
Public WinampArray2(30)
Public NotePadaRay(300)
Public M5, K5

Public Declare Function SetFocuses Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

'Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction&, ByVal uParam&, ByRef lpvParam As Any, ByVal fuWinIni&) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Declare Function ShowOwnedPopups Lib "user32" (ByVal hwnd As Long, ByVal fShow As Long) As Long

Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function closewindow Lib "user32" Alias "CloseWindow" (ByVal hWnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long



'Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowDLL Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long



'Public Declare Function FindWindowDDL _
'        Lib "user32" _
'        Alias "FindWindowA" _
'        (ByVal lpClassName As Long, _
'        ByVal lpWindowName As Long) As Long


Public Declare Function GetModuleFileNameA Lib "kernel32" _
         (ByVal hModule As Long, ByVal lpFileName As String, _
         ByVal nSize As Long) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'#########################################################
'# Convert() releated definitions
'#########################################################
    '#### Functions/Consts used for GetHWnd() (part of Convert())
Private Declare Function GetDesktopWindow Lib "user32" () As Long
'Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


Private Type FILETIME
   LowDateTime          As Long
   HighDateTime         As Long
End Type


Public Type WIN32_FIND_DATA
   dwFileAttributes     As Long
   ftCreationTime       As FILETIME
   ftLastAccessTime     As FILETIME
   ftLastWriteTime      As FILETIME
   nFileSizeHigh        As Long
   nFileSizeLow         As Long
   dwReserved0          As Long
   dwReserved1          As Long
   cFileName            As String * 260  'MUST be set to 260
   cAlternate           As String * 14
End Type






Public StopMe As Long
Public Snatch

'Public cProcesses As New clsCnProc

Public NoMusic As Long
Public NoMonOff As Long
Public NoMonOff2 As Long
'Public NoMmusicEx As long
'Public NoMonoffEx As long
Public Ebuyer As Long
Public Ebuy As Long
Public OTSS As Long
Public Cmdv As Long
Public Tagad As Long

Public ProcessId22 As Long
Public ProcessId23 As Long
Public ProcessId24 As Long
Public ProcessId25Ssam As Long

Public VirusActive As Long

Public Ebuyer3() As Long
Public Ebuyer4() As Long

Public GetWinLen() As Long
Public HwndArray() As Long
Public HwndArra2() As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type


Public globalhot$


'Const PROCESS_VM_READ = &H10
'Const PROCESS_QUERY_INFORMATION = &H400

'mattstimer2
Public MajorOnTime As Date
Public MinorOnTime As Date

Public MattsTimer2 As Date






Public Monitor_Timer As Date


Public sArray() As String
Public iCtr As Long

Public Peet3$
Public Peet33$

Public FlingGrater1$
Public FlingGrater2$

Public CheckT1$

Public OffSetGoogle
Public OldRect23
'Required for
'Private Declare Function FileInUse Lib "kernel32" (ByVal strFilePath As String) As Boolean

Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Rect8 As RECT
Public MeRyu3 As RECT
Public MeRyu4 As RECT
Public MeRyu5 As RECT
Private m_sPattern As String
Private m_lhFind As Long



Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Private Const PROCESS_CREATE_PROCESS = &H80
Private Const PROCESS_CREATE_THREAD = &H2
Private Const PROCESS_DUP_HANDLE = &H40
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const PROCESS_SET_QUOTA = &H100
Private Const PROCESS_SET_INFORMATION = &H200
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_VM_OPERATION = &H8
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_VM_WRITE = &H20
Private Const SYNCHRONIZE = &H100000

'--------------------------------------------- use these for *******************
Private Const NORMAL_PRIORITY_CLASS As Long = &H20
Private Const IDLE_PRIORITY_CLASS As Long = &H40
Private Const HIGH_PRIORITY_CLASS As Long = &H80
Private Const REALTIME_PRIORITY_CLASS As Long = &H100
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000


Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadId As Long
End Type

Private Type STARTUPINFO
  cb As Long
  lpReserved As Long
  lpDesktop As Long
  lpTitle As Long
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Byte
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

'Public Type STARTUPINFO
'   cb As Long
'   lpReserved As String
'   lpDesktop As String
'   lpTitle As String
'   dwX As Long
'   dwY As Long
'   dwXSize As Long
'   dwYSize As Long
'   dwXCountChars As Long
'   dwYCountChars As Long
'   dwFillAttribute As Long
'   dwFlags As Long
'   wShowWindow As Integer
'   cbReserved2 As Integer
'   lpReserved2 As Long
'   hStdInput As Long
'   hStdOutput As Long
'   hStdError As Long
'End Type
'
'Public Type PROCESS_INFORMATION
'   hProcess As Long
'   hThread As Long
'   dwProcessId As Long
'   dwThreadId As Long
'End Type

'Const SW_SHOWNORMAL = 1

'Public Declare Function MoveWindow _
'        Lib "user32" _
'        (ByVal hWnd As Long, _
'         ByVal X As Long, _
'         ByVal Y As Long, _
'         ByVal nWidth As Long, _
'         ByVal nHeight As Long, _
'         ByVal bRepaint As Long) As Long

'Public Declare Function FindWindow _
'        Lib "user32.dll" _
'            Alias "FindWindowA" _
'            (ByVal lpClassName As Long, _
'            ByVal lpWindowName As Long) As Long
            
            
Public Declare Function GetParent _
        Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetParent _
        Lib "user32" _
        (ByVal hWndChild As Long, _
         ByVal hWndNewParent As Long) As Long

Public Declare Function CreateProcessA _
        Lib "kernel32" _
        (ByVal lpApplicationName As String, _
         ByVal lpCommandLine As String, _
         ByVal lpProcessAttributes As Long, _
         ByVal lpThreadAttributes As Long, _
         ByVal bInheritHandles As Long, _
         ByVal dwCreationFlags As Long, _
         ByVal lpEnvironment As Long, _
         ByVal lpCurrentDirectory As String, _
         lpStartupInfo As STARTUPINFO, _
         lpProcessInformation As PROCESS_INFORMATION) As Long

Public Declare Function WaitForSingleObject _
        Lib "kernel32" _
        (ByVal hHandle As Long, _
        ByVal dwMilliseconds As Long) As Long

'Public Declare Function CloseHandle Lib "kernel32" _
        (ByVal hObject As Long) As Long

'Public Const NORMAL_PRIORITY_CLASS = &H20&


Private Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Any, ByVal lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
'Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long


Private Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Public Declare Function SetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Boolean
'---------------------------------------------

Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40

Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type



Public EasyDoesIt

Public TrueTerminate

'Private Type POINTAPI
'    X As Long
'    Y As Long
'End Type

Private Enum WindowStates
  SW_HIDE = 0
  SW_SHOWNORMAL = 1
  SW_NORMAL = 1
  SW_SHOWMINIMIZED = 2
  SW_SHOWMAXIMIZED = 3
  SW_MAXIMIZE = 3
  SW_SHOWNOACTIVATE = 4
  SW_SHOW = 5
  SW_MINIMIZE = 6
  SW_SHOWMINNOACTIVE = 7
  SW_SHOWNA = 8
  SW_RESTORE = 9
  SW_SHOWDEFAULT = 10
  SW_FORCEMINIMIZE = 11
  SW_MAX = 11
End Enum


Public Declare Function SendMessage _
        Lib "user32" _
        Alias "SendMessageA" _
        (ByVal hwnd As Long, _
         ByVal wMsg As Long, _
         ByVal wParam As Long, _
         ByVal lParam As Any) As Long

Private Const MONITOR_ON = -1&
Private Const MONITOR_OFF = 2&
Private Const SC_MONITORPOWER = &HF170&
Public Const WM_SYSCOMMAND = &H112
Public Const ipc_isplaying As Long = 104

Public Const WINAMP_BUTTON2 = 40045
Public Const WINAMP_BUTTON3 = 40046
'Public Const WM_CLOSE = &H10
'Public Const WM_USER = &H400
Public Const WM_COMMAND = &H111
'Private Const GW_HWNDNEXT = 2
Public Const Nexttrackbutton = 40048
Public Const Raisevolumeby1percent = 40058
Public Const Lowervolumeby1percent = 40059
Public Const WINAMP_FFWD5S = 40060                  '// fast forwards 5 seconds

'#define WINAMP_OPTIONS_EQ               40036 // toggles the EQ window
'#define WINAMP_OPTIONS_PLEDIT           40040 // toggles the playlist window
'#define WINAMP_VOLUMEUP                 40058 // turns the volume up a little
'#define WINAMP_VOLUMEDOWN               40059 // turns the volume down a little
'#define WINAMP_FFWD5S                   40060 // fast forwards 5 seconds
'#define WINAMP_REW5S                    40061 // rewinds 5 seconds

Public MsgResult As Long
Public XcNul As Long
Public LhRet As Long


'Type OFSTRUCT
'  cBytes     As Byte
'  fFixedDisk As Byte
'  nErrCode   As Integer
'  Reserved1  As Integer
'  Reserved2  As Integer
'  szPathName As String * 128
'End Type

'Const OF_SHARE_COMPAT = &H0
'Const OF_SHARE_EXCLUSIVE = &H10
'Const OF_SHARE_DENY_WRITE = &H20
'Const OF_SHARE_DENY_READ = &H30
'Const OF_SHARE_DENY_NONE = &H40

'Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
'Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'

'Option Explicit

' Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
 Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
 Const INFINITE = -1&




Function GetWindowClass(ByVal hwnd As Long) As String
    Dim Ret As Long, sText As String
    sText = Space(255)
    Ret = GetClassName(hwnd, sText, 255)
    sText = Left$(sText, Ret)
   GetWindowClass = sText
End Function


Public Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim L As Long
   Dim S As String
   L = GetWindowTextLength(hwnd)
   S = Space(L + 1)
   GetWindowText hwnd, S, L + 1
   GetWindowTitle = Left$(S, L)
End Function



Public Function GetWindowState(ByVal lngHwnd As Long) As Integer
    If IsWindow(lngHwnd) = 1 Then
        GetWindowState = -1
    If IsIconic(lngHwnd) <> 0 Then
        GetWindowState = vbMinimized
    ElseIf IsZoomed(lngHwnd) <> 0 Then
        GetWindowState = vbMaximized
    End If
    End If
End Function

'Public Function GetFileFromHwnd(lngHwnd) As String
'
''MsgBox getfilefromhwnd(Me.hwnd)
'
'Dim lngProcess&, hProcess&, bla&, C&
'Dim strFile As String
'Dim X
'
'strFile = String$(256, 0)
'X = GetWindowThreadProcessId(lngHwnd, lngProcess)
'hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
'X = EnumProcessModules(hProcess, bla, 4&, C)
'C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
'GetFileFromHwnd = Left(strFile, C)
'
'End Function


'Public Function GetFileFromProc(lngProcess) As String
'
''MsgBox getfilefromhwnd(Me.hwnd)
''Dim lngProcess&, hProcess&, bla&, C&
'Dim hProcess&, bla&, C&
'Dim strFile As String
'Dim X
'
'strFile = String$(256, 0)
''x = GetWindowThreadProcessId(lngHwnd, lngProcess)
'hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
'X = EnumProcessModules(hProcess, bla, 4&, C)
'C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
'GetFileFromProc = Left(strFile, C)
'
'End Function




Public Sub Perfect(Ew1$, LockStop)

Dim dickweed
Dim R
Dim Ash$
Dim totss As Boolean
Dim OutputData As Long
Dim totss2 As Long
Dim EWD$

Snatch = 0

EWD$ = Ew1$
Ew1$ = LCase$(Ew1$)

If InStr(Ew1$, LCase("fs.exe")) Then
    Snatch = 1
End If

If InStr(Ew1$, LCase("nero.exe")) Then
    NoMonOff = 1
    NoMusic = 1
    Snatch = 1
End If


If InStr(Ew1$, LCase("LeechGet.exe")) Then
    NoMonOff2 = 1
    NoMusic = 1
    Snatch = 1
End If

'''''outlook Express
If InStr(Ew1$, LCase("msimn.exe")) Then
    NoMonOff = 1
    Snatch = 1
End If

If InStr(Ew1$, LCase("Del_empty_folders.exe")) Then
    NoMonOff = 1
    Snatch = 1
End If

If InStr(Ew1$, LCase("IceView.exe")) Then
    NoMonOff = 1
    Snatch = 1
End If

If InStr(Ew1$, "realplay.exe") Then
    NoMonOff = 1: NoMusic = 1
    Snatch = 1
End If

If InStr(Ew1$, LCase("CTRec.exe")) Then
    NoMonOff = 1: NoMusic = 1
    Snatch = 1
End If

If InStr(Ew1$, LCase("PowerDVD.exe")) Then
    NoMonOff = 1: NoMusic = 1
    Snatch = 1
End If

If InStr(Ew1$, LCase("PowerVCD.exe")) Then
    NoMonOff = 1: NoMusic = 1
    Snatch = 1
End If


If InStr(Ew1$, LCase("SeriousSam.exe")) Then
    NoMonOff = 1
    NoMusic = 1
    ProcessId25Ssam = OTSS
    Snatch = 1
End If

If InStr(Ew1$, "p Video\winamp.exe") Then
    NoMonOff = 1: NoMusic = 1
    Snatch = 1
End If

If InStr(Ew1$, "p beavis\winamp.exe") Then
    NoMonOff = 1: NoMusic = 1
    Snatch = 1
End If

If InStr(Ew1$, "p midi\winampexe") Then
    NoMonOff = 1
    Snatch = 1
End If

If InStr(Ew1$, LCase("FRONTPG.EXE")) Then
    NoMonOff = 1
    frontpagepid = OTSS
    Snatch = 1
End If

If InStr(Ew1$, "cmd.exe") Then
    NoMonOff = 1
    Cmdv = 1
    Snatch = 1
End If

If InStr(Ew1$, "iexplore.exe") Then
    NoMonOff = 1: NoMusic = 1
    Ebuy = 1
    Snatch = 1
End If

If InStr(LCase(Ew1$), "amp caller\winamp.exe") Then
    ProcessId22 = OTSS
    Snatch = 1
End If
                
If InStr(LCase(Ew1$), "winamp.exe") Then
    ProcessId24 = OTSS
    Snatch = 1
End If

If Snatch = 1 And LockStop = 0 Then
    'dickweed = 0
   
    
    
    If InStr(Peet33$, str(OTSS)) = 0 Then
        etotss = cProcesses.Convert(CurProcHwnd, OTSS, cnFromhWnd Or cnToProcessID)
        totss = cProcesses.Convert(OTSS, totss2, cnTohWnd Or cnFromProcessID)
        CID_Run_Me.List2.AddItem Format$(totss2, "000000000") + " " + Format$(OTSS, "0000000") + " "
        Peet33$ = Peet33$ + str(OTSS)
        Peet3$ = Peet3$ + str(totss2)
    End If

    
    
    
    'totss = cProcesses.Convert(Otss, totss2, cnTohWnd Or cnFromProcessID)
                            
    'dickweed = 0
    
    'For r = 0 To Tagad
    '    If totss2 = HwndArray(r) Then
    '        dickweed = 1
    '        Exit For
    '    End If
    'Next

    'If dickweed = 0 Then
    '    For r = 0 To Tagad
    '        If HwndArray(r) < 0 Then
    '            HwndArray(r) = totss2
    '            HwndArra2(r) = Otss
    '            Tagad = Tagad + 1
    '            Exit For
    '
    '        End If
    '    Next
    'End If
End If

Ew1$ = EWD$
End Sub





Function FindWinPartNero(NeroDupe$) As Long
Dim Ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long


Dim NeroArray(30)
            
Dim cText As String
            

'Find the first window

'need this is you gona use this procedure get from CIDRun ME
'an another one
test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)

    
FindWinPartNero = False
    
NeroDupe$ = ""
    
Do While test_hwnd <> 0
    
    
    Ash$ = GetWindowTitle(test_hwnd)
    cText = Space$(255)
    ghj$ = GetClassName(test_hwnd, cText, 255)
    
    If InStr(Ash$, "Nero Wave Editor") And InStr(cText, "Afx:10000000:0") > 0 Then
        FindWindowPart = test_hwnd
        
        
        neros = neros + 1
        NeroArray(neros) = Ash$
        If neros > 1 Then
            For R = 1 To neros
                If NeroArray(R) = Ash$ And R <> neros And Ash$ <> "Nero Wave Editor" Then
                    NeroDupe$ = Ash$
                    EXITDUPE = 1
                Exit For
                End If
            Next
        End If
        
        If EXITDUPE = 1 Then Exit Do
    End If
    
'retrieve the next window
test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop
FindWinPartNero = neros

End Function





Function FindWinPartCodeBreak(TTF) As String

Dim Ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

'DONE IN DECLARES
'Dim TTFArray(30)
            
Dim cText As String
            
            
For R = 1 To UBound(TTFArray)
    If TTFArray(R) <> 0 Then
        If InStr(GetWindowTitle(TTFArray(R)), TTF) = 0 Then
            TTFArray(R) = 0
        End If
    End If
Next
            
If Idle_Timer + TimeSerial(0, 0, 5) > Now Then Exit Function
            
            
            
'Find the first window

'need this is you gona use this procedure get from CIDRun ME
'an another one
test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)

'FindWinPartCodeBreak = ""
    
Do While test_hwnd <> 0
    
    Ash$ = GetWindowTitle(test_hwnd)
    cText = Space$(255)
    ghj$ = GetClassName(test_hwnd, cText, 255)
    
    If InStr(Ash$, TTF) Then
        
        For R = 1 To 30
                If TTFArray(R) = test_hwnd Then
                    Exit For
                End If
                
                If TTFArray(R) = 0 Then
                    TTFArray(R) = test_hwnd
                    FindWinPartCodeBreak = Ash$
                    Exit For
                
                End If
                             
        Next
    End If
    
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop



End Function




Function FindWinPart(Dupe$, TF) As Long

Dim Ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long
Dim NeroArray(30)
Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one
test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)

FindWinPart = False
    
Do While test_hwnd <> 0
    
    Ash$ = GetWindowTitle(test_hwnd)
    cText = Space$(255)
    ghj$ = GetClassName(test_hwnd, cText, 255)
    
    If InStr(LCase(Ash$), LCase(Dupe$)) > 0 Then
        FindWinPart = test_hwnd
        hwnd9 = GetWindowRect(test_hwnd, Rect8)

'
'GoTo Jump5:
'        If TF = True Then
'
'Hx = 2 - (Rect8.Right - Rect8.Left) ' + 45
'hy = 40
'hw = Rect8.Right - Rect8.Left
'HH = Rect8.Bottom - Rect8.Top
'MoveWindow test_hwnd, Hx, hy, hw, HH, True
'
'
'        Else
'
'If InStr(ash$, "Writing") > 0 Then TG = 30 Else TG = 0
'
'Hx = 0 + TG '- (Rect5.Right - Rect5.Left) + 45
'hy = 40 + TG
'hw = Rect8.Right - Rect8.Left
'HH = Rect8.Bottom - Rect8.Top
'MoveWindow test_hwnd, Hx, hy, hw, HH, True
'
'        End If
'Jump5:
        
        'EXITDUPE = 1
'
'        GoTo skipme
'        neros = neros + 1
'        NeroArray(neros) = ash$
'        If neros > 1 Then
'            For R = 1 To neros
'                If NeroArray(R) = ash$ And R <> neros And ash$ <> "Nero Wave Editor" Then
'                    NeroDupe$ = ash$
'                    EXITDUPE = 1
'                Exit For
'                End If
'            Next
'        End If
'skipme:
'
        
        If EXITDUPE = 1 Then Exit Do
    End If
    
'retrieve the next window
test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop

'FindWinPartNero = neros

End Function



Function FindWinPartFront() As Long
FindWinPartFront = False

Dim Ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one
test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)
br$ = ""
Do While test_hwnd <> 0
    
    hwnd9 = GetWindowRect(test_hwnd, Rect8)
        Ash$ = GetWindowTitle(test_hwnd)
'        If InStr(ash$, "Double") > 0 Then Stop
        gws = GetWindowState(test_hwnd)
        If gws = -1 Then gws2$ = "Window State Normal"
        If gws = vbMinimized Then gws2$ = "Window State Minimized"
        If gws = vbMaximized Then gws2$ = "Window State Maximized"

    
    If (Rect8.Top > 0 And Rect8.Left > 0) Or gws = vbMinimized Then
        Ash$ = GetWindowTitle(test_hwnd)
        If Ash$ <> "" And InStr(br$, "-- " + Ash$ + " -- ") = 0 Then
            br$ = br$ + "-- " + Ash$ + " -- "
            On Error Resume Next
            AppActivate Ash$
            On Error GoTo 0
            huge = huge + 1
            ef = SetForegroundWindow(hwnd9)
            ef = Putfocus(hwnd9)
            ecute = Timer + 0.1
            Do
            If Timer < ecute - 30 Then Exit Do
            Loop Until Timer > ecute
        End If
    End If
        
'retrieve the next window
test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop
MsgBox str(huge) + " Windows Brought To Front"
FindWinPartFront = True

End Function



Function FindHighestHandle() As Long
'FindWinPartFront = False

Dim Ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one
test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)
br$ = ""
Do While test_hwnd <> 0
    

    
    Ash$ = GetWindowTitle(test_hwnd)
    If InStr(Ash$, "Program Manager") > 0 Then
        'etotss = cProcesses.Convert(test_hwnd, Ot$, cnFromhWnd Or cnToexe)
    End If
    
    If test_hwnd > HighTesthWnd Then
        HighTesthWnd = test_hwnd: Ash$ = GetWindowTitle(test_hwnd)
        'etotss = cProcesses.Convert(test_hwnd, Ot$, cnFromhWnd Or cnToexe)

    End If
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop
'MsgBox str(huge) + " Windows Brought To Front"
FindHighestHandle = HighTesthWnd

'Date 06-10-08
'"[System Process]"
'1593707806   -- "Program Manager"
'This SysUptime 61 Days 0h 3m 54s 532 mil
'Last SysUptime 29 Days 3h 1m 49s 203 mil




End Function






Function FindWinPartNotePad() As Long
FindWinPartNotePad = False

Dim Ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String

'Find the first window
huge = 0
'need this is you gona use this procedure get from CIDRun ME an another one
test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)
br$ = ""
Do While test_hwnd <> 0
    
    Ash$ = GetWindowClass(test_hwnd)
            
        If Ash$ = "Notepad2" Then
            gws = GetWindowState(test_hwnd)
            If gws = -1 Then gws2$ = "Window State Normal"
            If gws = vbMinimized Then gws2$ = "Window State Minimized"
            If gws = vbMaximized Then
            gws2$ = "Window State Maximized"
            End If
            If gws = -1 Then
                huge = huge + 1
                ReDim Preserve NoteHard2(huge)
                NoteHard2(huge) = test_hwnd
            End If
        End If
        
'retrieve the next window
test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop

'after chking all items in notehard2 we should have deled items in list that should not be there
'after that we should add new items in hard2 to list if they are not there which they wont be after 1st chk
'we chk what items in hard needed coz we del the ones not need when 1st chk


For R1 = 0 To CID_Run_Me.Lst1.ListCount - 1
    Key = 0
    For R2 = 1 To UBound(NoteHard2)
        If Val(CID_Run_Me.Lst1.List(R1)) = NoteHard2(R2) Then
            Keg = 1
        End If
    Next
    If Keg = 0 Then
        CID_Run_Me.Lst1.RemoveItem (R1)
    End If
Next



For R1 = 0 To CID_Run_Me.Lst1.ListCount - 1
    Key = 0
    For R2 = 1 To UBound(NoteHard2)
        If Val(CID_Run_Me.Lst1.List(R1)) = NoteHard2(R2) Then
            Keg = 1
            NoteHard2(R2) = 0
        End If
    Next
Next




For R1 = 1 To UBound(NoteHard2)
    If NoteHard2(R1) > 0 Then
        CID_Run_Me.Lst1.AddItem str(NoteHard2(R1))
    End If
Next


Cnt = 0
For R = 0 To CID_Run_Me.Lst1.ListCount - 1
    
    
    Xxt1 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
    wert1 = GetWindowRect(Xxt1, MeRyu4)
    xxt2 = FindWindow("Shell_TrayWnd", vbNullString)
    wert2 = GetWindowRect(xxt2, MeRyu5)
    
    'XXrXS(XXRsCNT) = Xxr
    
    xwid = Screen.Width / Screen.TwipsPerPixelX

'    MoveWindow Xxr, 0, MeRyu4.Bottom, xwid, MeRyu5.Top - MeRyu4.Bottom, True



    Xxt1 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
    xxt2 = FindWindow("Shell_TrayWnd", vbNullString)
    wert1 = GetWindowRect(Xxt1, MeRyu4)
    wert2 = GetWindowRect(xxt2, MeRyu5)
    
    oop = (MeRyu4.Right / CID_Run_Me.Lst1.ListCount) / 1.25
    gws = GetWindowState(Val(CID_Run_Me.Lst1.List(R)))
    If R > 0 And gws = -1 Then Cnt = Cnt + oop
    
    SW = Screen.Width / Screen.TwipsPerPixelX
    HH = 10
    'MoveWindow Xxr, 0, hh, SW, MeRyu5.Top - hh, True
    
    If gws = -1 Then
        
        MoveWindow Val(CID_Run_Me.Lst1.List(R)), 0, HH, SW, MeRyu5.Top - HH, True
        'MoveWindow Val(CID_Run_Me.Lst1.List(R)), Cnt, MeRyu4.Bottom, oop, MeRyu5.Top - MeRyu4.Bottom, True
    
    End If
Next

FindWinPartNotePad = huge

End Function


Sub DIR_ON_ANOTHER_DRIVE(TEXTDRI)
    On Error Resume Next
    DD = TEXTDRI
    If Right$(DD, 1) <> "\" Then DD = DD + "\"
    R = Len(DD)
    Do
                
        t = InStrRev(DD, "\")
        DD = Mid$(DD, 1, t - 1)
        R = t
        Err.Clear
        If t > 3 Then MkDir DD
        YY = Err.Description
        UU = Err.Number
    
    Loop Until t <= 3
'End

End Sub



 
 
Function FileInUse(ByVal strFilePath As String) As Boolean
  
  Dim hFile As Long
  Dim FileInfo  As OFSTRUCT
  
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function
  If Right(strFilePath, 1) <> Chr(0) Then strFilePath = strFilePath & Chr(0)
  
  FileInfo.cBytes = Len(FileInfo)
  hFile = OpenFile(strFilePath, FileInfo, OF_SHARE_EXCLUSIVE)
  If hFile = -1 And Err.LastDllError = 32 Then
    FileInUse = True
  Else
    CloseHandle hFile
  End If
  
End Function

 Public Function shellAndWait(ByVal fileName As String, LL) As Long
     Dim executionStatus As Long
     Dim ProcessHandle As Long
     Dim ReturnValue As Long
     
     'Execute the application/file
     executionStatus = Shell(fileName, LL)
     'executionStatus = Shell(fileName, vbHide)

     'Get the Process Handle
     ProcessHandle = OpenProcess(&H100000, True, executionStatus)

     'Wait till the application is finished
     ReturnValue = WaitForSingleObject(ProcessHandle, INFINITE)

     'Send the Return Value Back
     shellAndWait = ReturnValue

 End Function





Public Function ExecCmd(cmdLine$) As Long
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim Ret&

     ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)
                          
                          'NORMAL_PRIORITY_CLASS, _

    ' Start the shelled application:
    Ret& = CreateProcessA(vbNullString, cmdLine$, _
                          0&, 0&, 1&, _
                          IDLE_PRIORITY_CLASS, _
                          0&, vbNullString, _
                          start, proc)
    ' --- let it start - this seems important
    Call WaitForSingleObject(proc.hProcess, 1500)
'    Call WaitForSingleObject(proc.hProcess, INFINITE)
    
    
    If Ret Then
       ExecCmd = InstanceToWnd(proc.dwProcessId)
       'Me.Print Ret, ExecCmd
       Call CloseHandle(proc.hThread)
       Call CloseHandle(proc.hProcess)
    End If
    
End Function


Function InstanceToWnd(ByVal target_pid As Long) As Long
    Dim test_hwnd As Long, test_pid As Long
    'Find the first window
    test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)
    Do While test_hwnd <> 0
       'Check if the window isn't a child
       If GetParent(test_hwnd) = 0 Then
         'Get the window's thread
         test_thread_id = GetWindowThreadProcessId(test_hwnd, _
                             test_pid)
         If test_pid = target_pid Then
            InstanceToWnd = test_hwnd
            Exit Do
         End If
       End If
       'retrieve the next window
       test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    Loop
End Function








