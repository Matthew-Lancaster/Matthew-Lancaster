Attribute VB_Name = "Cid_Run_Me_M"
Public Reload_Me

Public Path_And_FileName As String
Public Path_Folder As String


Public Idle_Timer

Public FS, F, RamDrive
Dim TTFArray(30)

Public CurProcHwnd As Long, MustUnload
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long
Public NoteHard1()
Public NoteHard2()

Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Public ScreenTwipsX, ScreenTwipsY, ScreenWidthX, ScreenHeightY

Public VcodeTT$
Public KeyLoggDate
Public WinampArray(30)
Public WinampArray2(30)
Public NotePadaRay(300)
Public M5, K5

Public Declare Function SetFocuses Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

'Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction&, ByVal uParam&, ByRef lpvParam As Any, ByVal fuWinIni&) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Declare Function ShowOwnedPopups Lib "user32" (ByVal hWnd As Long, ByVal fShow As Long) As Long

Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
'Declare Function closewindow Lib "user32" Alias "CloseWindow" (ByVal hWnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

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

Public Type POINTAPI
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
Public MeRect As RECT
Private m_sPattern As String
Private m_lhFind As Long






Public Function DirExist(sPath As String) As Boolean
    DirExist = (Dir(sPath, vbDirectory) <> "")
End Function

Public Function MkDirNested(sFullPath As String) As Boolean
    On Error GoTo ErrHandler
    Dim LnNextSlash As Integer
    Dim LnStartPos As Integer
    Dim LsCurDir As String
    ' Set first char
    LnStartPos = 1
    ' validates path syntax
    If (Right$(sFullPath, 1) <> "\") Then sFullPath = (sFullPath & "\")
    Do
        LnNextSlash = InStr(LnStartPos, sFullPath, "\")
        If (LnNextSlash >= LnStartPos) Then
            LsCurDir = Left$(sFullPath, LnNextSlash)
            If (Not DirExist(LsCurDir)) Then
               ' Create the dir
               MkDir LsCurDir
            End If
             ' Check if it's the last char and exit if true
             LnStartPos = (LnNextSlash + 1)
             If (LnStartPos >= Len(sFullPath)) Then Exit Do
        End If
    Loop
    MkDirNested = True
    Exit Function
ErrHandler:
    MkDirNested = False
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

'MsgBox getfilefromhwnd(Me.hwnd)

Dim lngProcess&, hProcess&, bla&, C&
Dim strFile As String
Dim x

strFile = String$(256, 0)
x = GetWindowThreadProcessId(lngHwnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
x = EnumProcessModules(hProcess, bla, 4&, C)
C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
GetFileFromHwnd = Left(strFile, C)

End Function


Public Function GetFileFromProc(lngProcess) As String

'MsgBox getfilefromhwnd(Me.hwnd)
'Dim lngProcess&, hProcess&, bla&, C&
Dim hProcess&, bla&, C&
Dim strFile As String
Dim x

strFile = String$(256, 0)
'x = GetWindowThreadProcessId(lngHwnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
x = EnumProcessModules(hProcess, bla, 4&, C)
C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
GetFileFromProc = Left(strFile, C)

End Function




Public Sub Perfect(Ew1$, LockStop)

Dim dickweed
Dim R
Dim Window_Title_String
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

Ew1$ = ewd$
End Sub





Function FindWinPartNero(NeroDupe$) As Long

Dim Window_Title_String
Dim Test_HWND As Long, _
    Test_pid As Long, _
    Test_thread_id As Long


Dim NeroArray(30)
            
Dim cText As String
            

'Find the first window

'need this is you gona use this procedure get from CIDRun ME
'an another one
Test_HWND = FindWindow2(ByVal 0&, ByVal 0&)

    
FindWinPartNero = False
    
NeroDupe$ = ""
    
Do While Test_HWND <> 0
    
    
    Window_Title_String = GetWindowTitle(Test_HWND)
    cText = Space$(255)
    ghj$ = GetClassName(Test_HWND, cText, 255)
    
    If InStr(Window_Title_String, "Nero Wave Editor") And InStr(cText, "Afx:10000000:0") > 0 Then
        FindWindowpart = Test_HWND
        
        
        neros = neros + 1
        NeroArray(neros) = Window_Title_String
        If neros > 1 Then
            For R = 1 To neros
                If NeroArray(R) = Window_Title_String And R <> neros And Window_Title_String <> "Nero Wave Editor" Then
                    NeroDupe$ = Window_Title_String
                    EXITDUPE = 1
                Exit For
                End If
            Next
        End If
        
        If EXITDUPE = 1 Then Exit Do
    End If
    
'retrieve the next window
Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)

Loop
FindWinPartNero = neros

End Function


Function FindWinPartTOP(Top) As Long
Dim TITLETXT, CLASSTXT As String
Dim Test_HWND As Long, _
    Test_pid As Long, _
    Test_thread_id As Long


Dim cText As String
            

'Find the first window

'need this is you gona use this procedure get from CIDRun ME
'an another one
Test_HWND = FindWindow2(ByVal 0&, ByVal 0&)

    
FindWinPartTOP = False

'Debug.Print "Start ------------"

Do While Test_HWND <> 0
    
    
    TITLETXT = GetWindowTitle(Test_HWND)
    cText = Space$(255)
    CLASSTXT = GetClassName(Test_HWND, cText, 255)
    CLASSTXT = StripNulls(cText)
    
    If InStr(CLASSTXT, "BaseBar") > 0 Then
        
            Tx = Tx + 1
'            Debug.Print str(txt) + "GetWindowTitle   " + GetWindowTitle(Test_HWND)
'            Debug.Print GetWindowClass(Test_HWND)
            TTT = GetWindowRect(Test_HWND, MeRect)
            
'            Debug.Print Test_HWND
'            Debug.Print MeRect.Bottom
'            Debug.Print GetParentTOP(Test_HWND)
            FindWinPartTOP = Test_HWND
                        
            A = A
        
        
        
    End If
    
'retrieve the next window
Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)

Loop



End Function





Function FindWinPartCodeBreak(TTF) As String

Dim Window_Title_String
Dim Test_HWND As Long, _
    Test_pid As Long, _
    Test_thread_id As Long

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
Test_HWND = FindWindow2(ByVal 0&, ByVal 0&)

'FindWinPartCodeBreak = ""
    
Do While Test_HWND <> 0
    
    Window_Title_String = GetWindowTitle(Test_HWND)
    cText = Space$(255)
    ghj$ = GetClassName(Test_HWND, cText, 255)
    
    If InStr(Window_Title_String, TTF) Then
        
        For R = 1 To 30
                If TTFArray(R) = Test_HWND Then
                    Exit For
                End If
                
                If TTFArray(R) = 0 Then
                    TTFArray(R) = Test_HWND
                    FindWinPartCodeBreak = Window_Title_String
                    Exit For
                
                End If
                             
        Next
    End If
    
    'retrieve the next window
    Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)

Loop



End Function



Function FindWinPart_ANY_STRING(TTF) As Long

Dim Window_Title_String
Dim Test_HWND As Long, _
    Test_pid As Long, _
    Test_thread_id As Long

Dim cText As String
            
            
FindWinPart_ANY_STRING = 0

'Need this is you gona use this procedure get from CIDRun ME and another one
'Find the first window
Test_HWND = FindWindow2(ByVal 0&, ByVal 0&)

Do While Test_HWND <> 0
    
    Window_Title_String = GetWindowTitle(Test_HWND)
    cText = Space$(255)
    ghj$ = GetClassName(Test_HWND, cText, 255)
    
    If InStr(Window_Title_String, TTF) Then
        
        FindWinPart_ANY_STRING = Test_HWND
        Exit Function
        
    End If
    
    'retrieve the next window
    Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)

Loop



End Function





Function FindWinPart(Dupe$, TF) As Long

Dim Window_Title_String
Dim Test_HWND As Long, _
    Test_pid As Long, _
    Test_thread_id As Long
Dim NeroArray(30)
Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one
Test_HWND = FindWindow2(ByVal 0&, ByVal 0&)

FindWinPart = False
    
Do While Test_HWND <> 0
    
    Window_Title_String = GetWindowTitle(Test_HWND)
    cText = Space$(255)
    ghj$ = GetClassName(Test_HWND, cText, 255)
    
    If InStr(Window_Title_String, Dupe$) > 0 Then
        FindWinPart = Test_HWND
        hwnd9 = GetWindowRect(Test_HWND, Rect8)


GoTo Jump5:
        If TF = True Then

Hx = 2 - (Rect8.Right - Rect8.Left) ' + 45
hy = 40
hw = Rect8.Right - Rect8.Left
hh = Rect8.Bottom - Rect8.Top
MoveWindow Test_HWND, Hx, hy, hw, hh, True

        
        Else
        
If InStr(Window_Title_String, "Writing") > 0 Then tg = 30 Else tg = 0
        
Hx = 0 + tg '- (Rect5.Right - Rect5.Left) + 45
hy = 40 + tg
hw = Rect8.Right - Rect8.Left
hh = Rect8.Bottom - Rect8.Top
MoveWindow Test_HWND, Hx, hy, hw, hh, True
        
        End If
Jump5:
        
        'EXITDUPE = 1
        
        GoTo skipme
        neros = neros + 1
        NeroArray(neros) = Window_Title_String
        If neros > 1 Then
            For R = 1 To neros
                If NeroArray(R) = Window_Title_String And R <> neros And Window_Title_String <> "Nero Wave Editor" Then
                    NeroDupe$ = Window_Title_String
                    EXITDUPE = 1
                Exit For
                End If
            Next
        End If
skipme:
        
        
        If EXITDUPE = 1 Then Exit Do
    End If
    
'retrieve the next window
Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)

Loop

'FindWinPartNero = neros

End Function


Function FindWindow_Get_All_Explorer() As String

FindWindow_Get_All_Explorer = ""

Dim WINDOW_TITLE
Dim Test_HWND As Long, _
    Test_pid As Long, _
    Test_thread_id As Long

Dim cText As String
Huge = 0

Test_HWND = FindWindow2(ByVal 0&, ByVal 0&)
BR$ = ""
Do While Test_HWND <> 0
    
    WINDOW_TITLE = GetWindowTitle(Test_HWND)
    
    '--------------------------------------------------
    'C:\Windows\explorer.exe
    '--------------------------------------------------
    If GetWindowClass(Test_HWND) = "CabinetWClass" Then
        Huge = Huge + 1
        BR$ = BR$ + WINDOW_TITLE + vbCrLf + vbCrLf
    End If
        
    'retrieve the next window
    Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)

Loop

BR2$ = "EXPLORER SESSION CURRENTLY WORKING ON MAYBE REBOOT OR TASKKILL" + vbCrLf
BR2$ = BR2$ + "C:\Windows\explorer.exe" + vbCrLf
BR2$ = BR2$ + "FIND ON CLASS NAME __CabinetWClass__" + vbCrLf
BR2$ = BR2$ + Trim(str(Huge)) + " __ COUNT __ EXPLORER Window SET FOUND PASTE CLIPBOARD" + vbCrLf + vbCrLf
BR2$ = BR2$ + BR$


Clipboard.Clear
Clipboard.SetText BR2$
FindWindow_Get_All_Explorer = BR2$

MsgBox BR2$

End Function




Function FindWinPartFront() As Long
FindWinPartFront = False

Dim Window_Title_String
Dim Test_HWND As Long, _
    Test_pid As Long, _
    Test_thread_id As Long

Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one
Test_HWND = FindWindow2(ByVal 0&, ByVal 0&)
BR$ = ""
Do While Test_HWND <> 0
    
    hwnd9 = GetWindowRect(Test_HWND, Rect8)
        Window_Title_String = GetWindowTitle(Test_HWND)
'        If InStr(Window_Title_String, "Double") > 0 Then Stop
        gws = GetWindowState(Test_HWND)
        If gws = -1 Then gws2$ = "Window State Normal"
        If gws = vbMinimized Then gws2$ = "Window State Minimized"
        If gws = vbMaximized Then gws2$ = "Window State Maximized"

    
    If (Rect8.Top > 0 And Rect8.Left > 0) Or gws = vbMinimized Then
        Window_Title_String = GetWindowTitle(Test_HWND)
        If Window_Title_String <> "" And InStr(BR$, "-- " + Window_Title_String + " -- ") = 0 Then
            BR$ = BR$ + "-- " + Window_Title_String + " -- "
            On Error Resume Next
            AppActivate Window_Title_String
            On Error GoTo 0
            Huge = Huge + 1
            ef = SetForegroundWindow(hwnd9)
            ef = Putfocus(hwnd9)
            ecute = Timer + 0.1
            Do
            If Timer < ecute - 30 Then Exit Do
            Loop Until Timer > ecute
        End If
    End If
        
'retrieve the next window
Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)

Loop
MsgBox str(Huge) + " Windows Brought To Front"
FindWinPartFront = True

End Function



Function FindHighestHandle() As Long
'FindWinPartFront = False

Dim Window_Title_String
Dim Test_HWND As Long, _
    Test_pid As Long, _
    Test_thread_id As Long

Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one
Test_HWND = FindWindow2(ByVal 0&, ByVal 0&)
BR$ = ""
Do While Test_HWND <> 0
    

    
    Window_Title_String = GetWindowTitle(Test_HWND)
    If InStr(Window_Title_String, "Program Manager") > 0 Then
        'etotss = cProcesses.Convert(Test_HWND, Ot$, cnFromhWnd Or cnToexe)
    End If
    
    If Test_HWND > hightesthwnd Then
        hightesthwnd = Test_HWND: Window_Title_String = GetWindowTitle(Test_HWND)
        'etotss = cProcesses.Convert(Test_HWND, Ot$, cnFromhWnd Or cnToexe)

    End If
    'retrieve the next window
    Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)

Loop
'MsgBox str(huge) + " Windows Brought To Front"
FindHighestHandle = hightesthwnd

'Date 06-10-08
'"[System Process]"
'1593707806   -- "Program Manager"
'This SysUptime 61 Days 0h 3m 54s 532 mil
'Last SysUptime 29 Days 3h 1m 49s 203 mil




End Function

Function FindWinPartMYCOMPUTER() As Long
FindWinPartMYCOMPUTER = False


Dim Window_Title_String
Dim Test_HWND As Long, _
    Test_pid As Long, _
    Test_thread_id As Long

Dim cText As String

'Find the first window
Huge = 0
'need this is you gona use this procedure get from CIDRun ME an another one
Test_HWND = FindWindow2(ByVal 0&, ByVal 0&)
BR$ = ""
Do While Test_HWND <> 0
    
            
        If GetWindowClass(Test_HWND) = "BaseBar" Then
            Debug.Print GetWindowTitle(Test_HWND)
            Debug.Print GetWindowClass(Test_HWND)

            TTT = GetWindowRect(Test_HWND, MeRect)
            Debug.Print Test_HWND
            Debug.Print MeRect.Bottom
            Debug.Print GetParentTOP(Test_HWND)

            
            'gws = GetWindowState(Test_HWND)
            'If gws = -1 Then gws2$ = "Window State Normal"
            'If gws = vbMinimized Then gws2$ = "Window State Minimized"
            'If gws = vbMaximized Then
            'gws2$ = "Window State Maximized"
            'End If
            'If gws = -1 Then
                'huge = huge + 1
                'ReDim Preserve NoteHard2(huge)
                'NoteHard2(huge) = Test_HWND
            'End If
        End If
        
        'retrieve the next window
        Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)

Loop

FindWinPartMYCOMPUTER = Test_HWND

'after chking all items in notehard2 we should have deled items in list that should not be there
'after that we should add new items in hard2 to list if they are not there which they wont be after 1st chk
'we chk what items in hard needed coz we del the ones not need when 1st chk



End Function




Function FindWinPartNotePad() As Long
FindWinPartNotePad = False

Dim Window_Title_String
Dim Test_HWND As Long, _
    Test_pid As Long, _
    Test_thread_id As Long

Dim cText As String

'Find the first window
Huge = 0
'need this is you gona use this procedure get from CIDRun ME an another one
Test_HWND = FindWindow2(ByVal 0&, ByVal 0&)
BR$ = ""
Do While Test_HWND <> 0
    
    Window_Title_String = GetWindowClass(Test_HWND)
            
        If Window_Title_String = "Notepad2" Then
            gws = GetWindowState(Test_HWND)
            If gws = -1 Then gws2$ = "Window State Normal"
            If gws = vbMinimized Then gws2$ = "Window State Minimized"
            If gws = vbMaximized Then
            gws2$ = "Window State Maximized"
            End If
            If gws = -1 Then
                Huge = Huge + 1
                ReDim Preserve NoteHard2(Huge)
                NoteHard2(Huge) = Test_HWND
            End If
        End If
        
'retrieve the next window
Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)

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
    hh = 10
    'MoveWindow Xxr, 0, hh, SW, MeRyu5.Top - hh, True
    
    If gws = -1 Then
        
        MoveWindow Val(CID_Run_Me.Lst1.List(R)), 0, hh, SW, MeRyu5.Top - hh, True
        'MoveWindow Val(CID_Run_Me.Lst1.List(R)), Cnt, MeRyu4.Bottom, oop, MeRyu5.Top - MeRyu4.Bottom, True
    
    End If
Next

FindWinPartNotePad = Huge

End Function






