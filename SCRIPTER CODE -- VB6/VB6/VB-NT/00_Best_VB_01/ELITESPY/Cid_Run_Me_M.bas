Attribute VB_Name = "Cid_Run_Me_M"
Public MsgBox_11, EXIT_TRUE

Private Const GW_HWNDNEXT = 2
Public APP_NAME_RELOAD_IT_ER____
Public APP_NAME_RELOAD_IT_ER_EXE As String
Public VAR_ARRAY
Public BLOCK_RUN_1()
Public BLOCK_RUN_2()
Public BLOCK_RUN_3()

Public FS, F, RamDrive

'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer

Private Type POINTAPI
        X As Long
        Y As Long
End Type


Public CurProchWnd As Long, MustUnload


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long
Public NoteHard1()
Public NoteHard2()

Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Public ScreenTwipsX, ScreenTwipsY, ScreenWidthX, ScreenHeightY, Idle_Timer_Proc

Public VcodeTT$
Public KeyLoggDate
Public WinampArray(30)
Public WinampArray2(30)
Public NotePadaRay(300)
Public M5, K5

Public Declare Function SetFocuses Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction&, ByVal uParam&, ByRef lpvParam As Any, ByVal fuWinIni&) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Declare Function ShowOwnedPopups Lib "user32" (ByVal hWnd As Long, ByVal fShow As Long) As Long

Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
'Declare Function closewindow Lib "user32" Alias "CloseWindow" (ByVal hWnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long


Public Declare Function GetModuleFileNameA Lib "kernel32" _
         (ByVal hModule As Long, ByVal lpFileName As String, _
         ByVal nSize As Long) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long




Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


Private Type FILETIME
   LowDateTime          As Long
   HighDateTime         As Long
End Type


Private Type WIN32_FIND_DATA
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

Public cProcesses As New clsCnProc

Public NoMusic As Long
Public NoMonOff As Long
Public NoMonOff2 As Long
'Public NoMmusicEX As long
'Public NoMonoffEX As long
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
Public hWndArray() As Long
Public hWndArra2() As Long

'Public Type POINTAPI
'    X As Long
'    Y As Long
'End Type


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

Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long

Public Type RECT
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


Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long



'http://www.vbforums.com/showthread.php?673134-RESOLVED-Minimum-height-for-menu-bar-to-be-visible
'-------------- MENU SIZE DECLARE
'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'----------------------------------------------------------------------------------
Private Type MENUBARINFO
  cbSize As Long
  rcBar As RECT
  hMenu As Long
  hWndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type
Private MenuInfo As MENUBARINFO
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
'----------------------------------------------------------------------------------


Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long


'------------------------------------
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16


Public Function CreateFolderTree(ByVal sPath As String) As Boolean
    Dim nPos As Integer

    On Error GoTo CreateFolderTreeError
    
    nPos = InStr(sPath, "\")
    While nPos > 0
        
        '----------------------------------------------------------------------------
        'ROUTINE TAKEN FROM
        '----------------------------------------------------------------------------
        'SEND_TO_SCRIPT_IRFAR - Microsoft Visual Basic [design] - [mdlFileSys (Code)]
        'D:\VB6\VB-NT\00_Send_To\Send To Text List of Files & Sub Folders IRFAR\#0 Send To Text List of Files and Sub Folders IRFAN.exe
        '----------------------------------------------------------------------------
        'MODIFIED A BIT DIR COMMAND REPLACE MORE COMPLEX WAY
        '----------------------------------------------------------------------------
        
        'If Not FolderExists(Left$(sPath, nPos - 1)) Then
        
        If Dir((Left$(sPath, nPos - 1)), vbDirectory) = "" Then
            MkDir Left$(sPath, nPos - 1)
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
    'If Not FolderExists(sPath) Then MkDir sPath
    If Dir(sPath, vbDirectory) = "" Then MkDir sPath
    
    CreateFolderTree = True
    Exit Function

CreateFolderTreeError:
    Exit Function
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

Public Function GetFileFromHwnd(lngHwnd) As String

'-----------------------------------------------------------
'REPAIR DONE NOW GET FULL PATH OF EXE BY USE THE OPENPROCESS
'COPY PASTE THE DECLARE AND DONE
'Thu 16 February 2017 03:11:44----------
'-----------------------------------------------------------
'-----------------------------------------------------------

'------------------------------------------------------------------
'------------------------------------------------------------------
'TEMPORARY AS FULL EXE WAS NOT WORKING FOR A BIT
'------------------------------------------------------------------
'MsgBox getfilefromhWnd(Me.hWnd)
'Var = cProcesses.Convert(lnghWnd, vfile, cnFromhWnd Or cnToEXE)
'Var = cProcesses.Convert(lnghWnd, PID, cnFromhWnd Or cnToProcessID)
'Var = cProcesses.Convert(PID, vfile, cnFromProcessID Or cnToEXE)
'GetFileFromhWnd = vfile
'Exit Function
'------------------------------------------------------------------
'------------------------------------------------------------------

Dim lngProcess&, hProcess&, bla&, C&
Dim strFile As String
Dim X

strFile = String$(256, 0)
X = GetWindowThreadProcessId(lngHwnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
X = EnumProcessModules(hProcess, bla, 4&, C)
C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
GetFileFromHwnd = Left(strFile, C)
'------------------------------------------------------------------

End Function

Public Function GetFileFromProc(lngProcess) As String

Dim hProcess&, bla&, C&
Dim strFile As String
Dim X

strFile = String$(256, 0)
'X = GetWindowThreadProcessId(lngHwnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
X = EnumProcessModules(hProcess, bla, 4&, C)
C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
GetFileFromProc = Left(strFile, C)

End Function




Public Sub Perfect(Ew1$, LockStop)

Dim dickweed
Dim R
Dim ash$
Dim totss As Boolean
Dim OutputData As Long
Dim totss2 As Long

Snatch = 0

ewd$ = Ew1$
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
   
    
    
    If InStr(Peet33$, Str(OTSS)) = 0 Then
        etotss = cProcesses.Convert(CurProchWnd, OTSS, cnFromhWnd Or cnToProcessID)
        totss = cProcesses.Convert(OTSS, totss2, cnTohWnd Or cnFromProcessID)
        CID_Run_Me.List2.AddItem Format$(totss2, "000000000") + " " + Format$(OTSS, "0000000") + " "
        Peet33$ = Peet33$ + Str(OTSS)
        Peet3$ = Peet3$ + Str(totss2)
    End If

    
    
    
    'totss = cProcesses.Convert(Otss, totss2, cnTohWnd Or cnFromProcessID)
                            
    'dickweed = 0
    
    'For r = 0 To Tagad
    '    If totss2 = hWndArray(r) Then
    '        dickweed = 1
    '        Exit For
    '    End If
    'Next

    'If dickweed = 0 Then
    '    For r = 0 To Tagad
    '        If hWndArray(r) < 0 Then
    '            hWndArray(r) = totss2
    '            hWndArra2(r) = Otss
    '            Tagad = Tagad + 1
    '            Exit For
    '
    '        End If
    '    Next
    'End If
End If

Ew1$ = ewd$
End Sub





Function FindWinPartNero(NeroDupe$) As Long
Dim ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long


Dim NeroArray(30)
            
Dim cText As String
            

'Find the first window

'need this is you gona use this procedure get from CIDRun ME
'an another one
test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)

    
FindWinPartNero = False
    
NeroDupe$ = ""
    
Do While test_hwnd <> 0
    
    
    ash$ = GetWindowTitle(test_hwnd)
    cText = Space$(255)
    ghj$ = GetClassName(test_hwnd, cText, 255)
    
    If InStr(ash$, "Nero Wave Editor") And InStr(cText, "Afx:10000000:0") > 0 Then
        FindWindowPart = test_hwnd
        
        
        neros = neros + 1
        NeroArray(neros) = ash$
        If neros > 1 Then
            For R = 1 To neros
                If NeroArray(R) = ash$ And R <> neros And ash$ <> "Nero Wave Editor" Then
                    NeroDupe$ = ash$
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

Public Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim L As Long
   Dim S As String
   L = GetWindowTextLength(hWnd)
   S = Space(L + 1)
   GetWindowText hWnd, S, L + 1
   GetWindowTitle = Left$(S, L)
End Function


Private Function GetActiveWindow(ByVal ReturnParent As Boolean) As Long
   Dim i As Long
   Dim j As Long
   i = GetForegroundWindow
   If ReturnParent Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
      i = j
   End If
   GetActiveWindow = i
End Function

Function GetParentTitle(ByVal Handle As Long) As String
   Dim i As Long
   Dim j As Long, TTx1 As String
   i = Handle
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
      i = j
   TTx1 = GetWindowTitle(i)
   GetParentTitle = TTx1

End Function

'Function GetWindowTitle(ByVal hWnd As Long) As String
'   Dim l As Long
'   Dim S As String
'   l = GetWindowTextLength(hWnd)
'   S = Space(l + 1)
'   GetWindowText hWnd, S, l + 1
'   GetWindowTitle = Left$(S, l)
'End Function
Function GetWindowClass(ByVal hWnd As Long) As String
    Dim Ret As Long, sText As String
    sText = Space(255)
    Ret = GetClassName(hWnd, sText, 255)
    sText = Left$(sText, Ret)
   GetWindowClass = sText
End Function


Function FindWinPart(Dupe$, TF) As Long
Dim ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long
Dim NeroArray(30)
Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one
test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)

FindWinPart = False
    
Do While test_hwnd <> 0
    
    ash$ = GetWindowTitle(test_hwnd)
    cText = Space$(255)
    ghj$ = GetClassName(test_hwnd, cText, 255)
    
    If InStr(ash$, Dupe$) > 0 Then
        FindWinPart = test_hwnd
        HWnd9 = GetWindowRect(test_hwnd, Rect8)
GoTo Jump5:
        If TF = True Then

HX = 2 - (Rect8.Right - Rect8.Left) ' + 45
HY = 40
HW = Rect8.Right - Rect8.Left
HH = Rect8.Bottom - Rect8.Top
MoveWindow test_hwnd, HX, HY, HW, HH, True

        
        Else
        
If InStr(ash$, "Writing") > 0 Then Tg = 30 Else Tg = 0
        
HX = 0 + Tg '- (Rect5.Right - Rect5.Left) + 45
HY = 40 + Tg
HW = Rect8.Right - Rect8.Left
HH = Rect8.Bottom - Rect8.Top
MoveWindow test_hwnd, HX, HY, HW, HH, True
        
        End If
Jump5:
        
        'EXITDUPE = 1
        
        GoTo skipme
        neros = neros + 1
        NeroArray(neros) = ash$
        If neros > 1 Then
            For R = 1 To neros
                If NeroArray(R) = ash$ And R <> neros And ash$ <> "Nero Wave Editor" Then
                    NeroDupe$ = ash$
                    EXITDUPE = 1
                Exit For
                End If
            Next
        End If
skipme:
        If EXITDUPE = 1 Then Exit Do
    End If
    
'retrieve the next window
test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop

'FindWinPartNero = neros

End Function



Function FindWinPartFront() As Long
FindWinPartFront = False

Dim ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one
test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
BR$ = ""
Do While test_hwnd <> 0
    
    HWnd9 = GetWindowRect(test_hwnd, Rect8)
        ash$ = GetWindowTitle(test_hwnd)
'        If InStr(ash$, "Double") > 0 Then Stop
        GWS = GetWindowState(test_hwnd)
        If GWS = -1 Then gws2$ = "Window State Normal"
        If GWS = vbMinimized Then gws2$ = "Window State Minimized"
        If GWS = vbMaximized Then gws2$ = "Window State Maximized"

    
    If (Rect8.Top > 0 And Rect8.Left > 0) Or GWS = vbMinimized Then
        ash$ = GetWindowTitle(test_hwnd)
        If ash$ <> "" And InStr(BR$, "-- " + ash$ + " -- ") = 0 Then
            BR$ = BR$ + "-- " + ash$ + " -- "
            On Error Resume Next
            AppActivate ash$
            On Error GoTo 0
            Huge = Huge + 1
            EF = SetForegroundWindow(HWnd9)
            'ef = Putfocus(hWnd9)
            ECUTE = Timer + 0.1
            Do
            If Timer < ECUTE - 30 Then Exit Do
            Loop Until Timer > ECUTE
        End If
    End If
        
'retrieve the next window
test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop
MsgBox Str(Huge) + " Windows Brought To Front"
FindWinPartFront = True

End Function


'
'
'Function FindWinPart(Dupe$, TF) As Long
'Dim ash$
'Dim test_hWnd As Long, _
'    test_pid As Long, _
'    test_thread_id As Long
'Dim NeroArray(30)
'Dim cText As String
'
''Find the first window
'
''need this is you gona use this procedure get from CIDRun ME an another one
'test_hWnd = FindWindow2(ByVal 0&, ByVal 0&)
'
'FindWinPart = False
'
'Do While test_hWnd <> 0
'
'    ash$ = GetWindowTitle(test_hWnd)
'    cText = Space$(255)
'    ghj$ = GetClassName(test_hWnd, cText, 255)
'
'    If InStr(ash$, Dupe$) > 0 Then
'        FindWinPart = test_hWnd
'        hWnd9 = GetWindowRect(test_hWnd, Rect8)
'GoTo Jump5:
'        If TF = True Then
'
'Hx = 2 - (Rect8.Right - Rect8.Left) ' + 45
'hy = 40
'hw = Rect8.Right - Rect8.Left
'hh = Rect8.Bottom - Rect8.Top
'MoveWindow test_hWnd, Hx, hy, hw, hh, True
'
'
'        Else
'
'If InStr(ash$, "Writing") > 0 Then Tg = 30 Else Tg = 0
'
'Hx = 0 + Tg '- (Rect5.Right - Rect5.Left) + 45
'hy = 40 + Tg
'hw = Rect8.Right - Rect8.Left
'hh = Rect8.Bottom - Rect8.Top
'MoveWindow test_hWnd, Hx, hy, hw, hh, True
'
'        End If
'Jump5:
'
'        'EXITDUPE = 1
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
'        If EXITDUPE = 1 Then Exit Do
'    End If
'
''retrieve the next window
'test_hWnd = GetWindow(test_hWnd, GW_hWndNEXT)
'
'Loop
'
''FindWinPartNero = neros
'
'End Function


'
'Function FindWinPartFront() As Long
'FindWinPartFront = False
'
'Dim ash$
'Dim test_hWnd As Long, _
'    test_pid As Long, _
'    test_thread_id As Long
'
'Dim cText As String
'
''Find the first window
'
''need this is you gona use this procedure get from CIDRun ME an another one
'test_hWnd = FindWindow2(ByVal 0&, ByVal 0&)
'br$ = ""
'Do While test_hWnd <> 0
'
'    hWnd9 = GetWindowRect(test_hWnd, Rect8)
'        ash$ = GetWindowTitle(test_hWnd)
''        If InStr(ash$, "Double") > 0 Then Stop
'        gws = GetWindowState(test_hWnd)
'        If gws = -1 Then gws2$ = "Window State Normal"
'        If gws = vbMinimized Then gws2$ = "Window State Minimized"
'        If gws = vbMaximized Then gws2$ = "Window State Maximized"
'
'
'    If (Rect8.Top > 0 And Rect8.Left > 0) Or gws = vbMinimized Then
'        ash$ = GetWindowTitle(test_hWnd)
'        If ash$ <> "" And InStr(br$, "-- " + ash$ + " -- ") = 0 Then
'            br$ = br$ + "-- " + ash$ + " -- "
'            On Error Resume Next
'            AppActivate ash$
'            On Error GoTo 0
'            Huge = Huge + 1
'            ef = SetForegroundWindow(hWnd9)
'            'ef = Putfocus(hWnd9)
'            ecute = Timer + 0.1
'            Do
'            If Timer < ecute - 30 Then Exit Do
'            Loop Until Timer > ecute
'        End If
'    End If
'
''retrieve the next window
'test_hWnd = GetWindow(test_hWnd, GW_hWndNEXT)
'
'Loop
'MsgBox Str(Huge) + " Windows Brought To Front"
'FindWinPartFront = True
'
'End Function
'

'
'Function FindHighestHandle() As Long
''FindWinPartFront = False
'
'Dim ash$
'Dim test_hWnd As Long, _
'    test_pid As Long, _
'    test_thread_id As Long
'
'Dim cText As String
'
''Find the first window
'
''need this is you gona use this procedure get from CIDRun ME an another one
'test_hWnd = FindWindow2(ByVal 0&, ByVal 0&)
'br$ = ""
'Do While test_hWnd <> 0
'
'
'
'    ash$ = GetWindowTitle(test_hWnd)
'    If InStr(ash$, "Program Manager") > 0 Then
'        'etotss = cProcesses.Convert(test_hWnd, Ot$, cnFromhWnd Or cnToexe)
'    End If
'
'    If test_hWnd > hightesthWnd Then
'        hightesthWnd = test_hWnd: ash$ = GetWindowTitle(test_hWnd)
'        'etotss = cProcesses.Convert(test_hWnd, Ot$, cnFromhWnd Or cnToexe)
'
'    End If
'    'retrieve the next window
'    test_hWnd = GetWindow(test_hWnd, GW_hWndNEXT)
'
'Loop
''MsgBox str(huge) + " Windows Brought To Front"
'FindHighestHandle = hightesthWnd
'
''Date 06-10-08
''"[System Process]"
''1593707806   -- "Program Manager"
''This SysUptime 61 Days 0h 3m 54s 532 mil
''Last SysUptime 29 Days 3h 1m 49s 203 mil
'
'
'
'
'End Function







Function FindWinPart_SEARCHER_hWnd_TO_EXE(SEARCH_STRING) As Long
    
    Dim Text_TEMP_ER As String
    Dim test_hwnd As Long, _
        test_pid As Long, _
        test_thread_id As Long
    
    Dim cText As String
    
    'Find the first window
    test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
    Huge = 0
    Do While test_hwnd <> 0
        
        Text_TEMP_ER = GetFileFromHwnd(test_hwnd)
        'If InStr(GetWindowTitle(test_hWnd), SEARCH_STRING) > 0 Then
        If InStr(UCase(Text_TEMP_ER), UCase(SEARCH_STRING)) > 0 Then
            Huge = Huge + 1
            
            'Debug.Print Str(Huge) + "__" + Str(test_hWnd) + "__" + Text_TEMP_ER
            'Debug.Print Str(Huge) + "__" + Str(test_hWnd) + "__" + GetWindowTitle(test_hWnd)
            'Debug.Print Str(Huge) + "__" + Str(test_hWnd) + "__" + GetWindowClass(test_hWnd)
            'Debug.Print vbCrLf
            'EXIT SUB
            Huge = Huge
        End If
            
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    
    Loop
End Function




Function FindWinPart_SEARCHER(SEARCH_STRING) As Long
    FindWinPart_SEARCHER = False
    
    Dim test_hwnd As Long, _
        test_pid As Long, _
        test_thread_id As Long
    
    Dim cText As String
    Dim CLASS_NAME As String
    Dim XGO
    Dim CLASS_NAME_______________
    
    'Find the first window
    test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
    Do While test_hwnd <> 0
        CLASS_NAME_______________ = GetWindowClass(test_hwnd)
        If InStr(UCase(CLASS_NAME_______________), UCase(SEARCH_STRING)) > 0 Then XGO = True
        If InStr(UCase(GetWindowTitle(test_hwnd)), UCase(SEARCH_STRING)) > 0 Then XGO = True
        If XGO = True Then
            FindWinPart_SEARCHER = test_hwnd: Exit Function
        End If
            
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    
    Loop
End Function





Function FindWinPartNotePad() As Long
FindWinPartNotePad = False

Dim ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String

'Find the first window
Huge = 0
'need this is you gona use this procedure get from CIDRun ME an another one
test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
BR$ = ""
Do While test_hwnd <> 0
    
    ash$ = GetWindowClass(test_hwnd)
            
        If ash$ = "Notepad2" Then
            GWS = GetWindowState(test_hwnd)
            If GWS = -1 Then gws2$ = "Window State Normal"
            If GWS = vbMinimized Then gws2$ = "Window State Minimized"
            If GWS = vbMaximized Then
            gws2$ = "Window State Maximized"
            End If
            If GWS = -1 Then
                Huge = Huge + 1
                ReDim Preserve NoteHard2(Huge)
                NoteHard2(Huge) = test_hwnd
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
        CID_Run_Me.Lst1.AddItem Str(NoteHard2(R1))
    End If
Next


Cnt = 0
For R = 0 To CID_Run_Me.Lst1.ListCount - 1
    Xxt1 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
    xxt2 = FindWindow("Shell_TrayWnd", vbNullString)
    wert1 = GetWindowRect(Xxt1, MeRyu4)
    wert2 = GetWindowRect(xxt2, MeRyu5)
    oop = (MeRyu4.Right / CID_Run_Me.Lst1.ListCount) / 1.25
    GWS = GetWindowState(Val(CID_Run_Me.Lst1.List(R)))
    If R > 0 And GWS = -1 Then Cnt = Cnt + oop
    If GWS = -1 Then MoveWindow Val(CID_Run_Me.Lst1.List(R)), Cnt, MeRyu4.Bottom, oop, MeRyu5.Top - MeRyu4.Bottom, True
Next

FindWinPartNotePad = Huge

End Function



Function FindHighestHandle() As Long
'FindWinPartFront = False

Dim ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one
test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
BR$ = ""
Do While test_hwnd <> 0
    

    
    ash$ = GetWindowTitle(test_hwnd)
    If InStr(ash$, "Program Manager") > 0 Then
        'etotss = cProcesses.Convert(test_hWnd, Ot$, cnFromhWnd Or cnToexe)
    End If
    
    If test_hwnd > hightesthWnd Then
        hightesthWnd = test_hwnd: ash$ = GetWindowTitle(test_hwnd)
        'etotss = cProcesses.Convert(test_hWnd, Ot$, cnFromhWnd Or cnToexe)

    End If
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop
'MsgBox str(huge) + " Windows Brought To Front"
FindHighestHandle = hightesthWnd

'Date 06-10-08
'"[System Process]"
'1593707806   -- "Program Manager"
'This SysUptime 61 Days 0h 3m 54s 532 mil
'Last SysUptime 29 Days 3h 1m 49s 203 mil




End Function





'
'Function FindWinPartNotePad() As Long
'FindWinPartNotePad = False
'
'Dim ash$
'Dim test_hWnd As Long, _
'    test_pid As Long, _
'    test_thread_id As Long
'
'Dim cText As String
'
''Find the first window
'Huge = 0
''need this is you gona use this procedure get from CIDRun ME an another one
'test_hWnd = FindWindow2(ByVal 0&, ByVal 0&)
'br$ = ""
'Do While test_hWnd <> 0
'
'    ash$ = GetWindowClass(test_hWnd)
'
'        If ash$ = "Notepad2" Then
'            gws = GetWindowState(test_hWnd)
'            If gws = -1 Then gws2$ = "Window State Normal"
'            If gws = vbMinimized Then gws2$ = "Window State Minimized"
'            If gws = vbMaximized Then
'            gws2$ = "Window State Maximized"
'            End If
'            If gws = -1 Then
'                Huge = Huge + 1
'                ReDim Preserve NoteHard2(Huge)
'                NoteHard2(Huge) = test_hWnd
'            End If
'        End If
'
''retrieve the next window
'test_hWnd = GetWindow(test_hWnd, GW_hWndNEXT)
'
'Loop
'
''after chking all items in notehard2 we should have deled items in list that should not be there
''after that we should add new items in hard2 to list if they are not there which they wont be after 1st chk
''we chk what items in hard needed coz we del the ones not need when 1st chk
'
'
'For R1 = 0 To CID_Run_Me.Lst1.ListCount - 1
'    Key = 0
'    For R2 = 1 To UBound(NoteHard2)
'        If Val(CID_Run_Me.Lst1.List(R1)) = NoteHard2(R2) Then
'            Keg = 1
'        End If
'    Next
'    If Keg = 0 Then
'        CID_Run_Me.Lst1.RemoveItem (R1)
'    End If
'Next
'
'
'
'For R1 = 0 To CID_Run_Me.Lst1.ListCount - 1
'    Key = 0
'    For R2 = 1 To UBound(NoteHard2)
'        If Val(CID_Run_Me.Lst1.List(R1)) = NoteHard2(R2) Then
'            Keg = 1
'            NoteHard2(R2) = 0
'        End If
'    Next
'Next
'
'
'
'
'For R1 = 1 To UBound(NoteHard2)
'    If NoteHard2(R1) > 0 Then
'        CID_Run_Me.Lst1.AddItem Str(NoteHard2(R1))
'    End If
'Next
'
'
'Cnt = 0
'For R = 0 To CID_Run_Me.Lst1.ListCount - 1
'    Xxt1 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
'    xxt2 = FindWindow("Shell_TrayWnd", vbNullString)
'    wert1 = GetWindowRect(Xxt1, MeRyu4)
'    wert2 = GetWindowRect(xxt2, MeRyu5)
'    oop = (MeRyu4.Right / CID_Run_Me.Lst1.ListCount) / 1.25
'    gws = GetWindowState(Val(CID_Run_Me.Lst1.List(R)))
'    If R > 0 And gws = -1 Then Cnt = Cnt + oop
'    If gws = -1 Then MoveWindow Val(CID_Run_Me.Lst1.List(R)), Cnt, MeRyu4.Bottom, oop, MeRyu5.Top - MeRyu4.Bottom, True
'Next
'
'FindWinPartNotePad = Huge
'
'End Function
'





