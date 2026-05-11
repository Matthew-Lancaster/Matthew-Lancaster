Attribute VB_Name = "Tidal_Mod"
'MoveWindow HwndCTLTask3, Rect3.Left - (Rect4.Right - Rect4.Left), Rect3.Top, Rect4.Right - Rect4.Left, Rect4.Bottom - Rect4.Top, True

Public Declare Function MoveWindow _
        Lib "user32" _
        (ByVal hwnd As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal bRepaint As Long) As Long


Private Type POINTAPI
        X As Long
        Y As Long
End Type




Public drive2$
Public drive3$

Public welong As Double, welat As Double

Public speedtime As Date

Public ash5$
Public procid2 As Long
Public procid3 As Long
Public hWnd2 As Long
Public hwnd3 As Long

Public Equinox5 As Date
Public Equinox6 As Date
Public tooleq1 As Date
Public tooleq2 As Date
Public tooleq3 As Date
Public tooleq4 As Date
Public processid25ssam As Long
Public frontpagepid As Long

Public yahoo As Integer

Public globalhot$


Public cProcesses As New clsCnProc





Public Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long




Public Declare Function closewindow Lib "user32" Alias "CloseWindow" (ByVal hwnd As Long) As Long

Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long

'Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long  'MODULE 1141
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142

Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As MODULEENTRY32) As Long

Public Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExA" (ByVal lpFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Boolean

Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwProcessId As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const PROCESS_TERMINATE = &H1

Public Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Boolean
'Public Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Boolean
Private Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long

Private Const MAX_PATH = 260

Public Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModule As String * 256    'MAX_MODULE_NAME32 + 1
    szExePath As String * MAX_PATH
End Type

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Public Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type



Public Declare Function CloseHandle Lib "kernel32" _
        (ByVal hObject As Long) As Long

Public Declare Function FindWindow _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long
Public Declare Function GetParent _
        Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetParent _
        Lib "user32" _
        (ByVal hWndChild As Long, _
         ByVal hWndNewParent As Long) As Long
Public Declare Function GetWindowThreadProcessId _
        Lib "user32" _
        (ByVal hwnd As Long, _
        lpdwProcessId As Long) As Long
Public Declare Function GetWindow _
        Lib "user32" _
        (ByVal hwnd As Long, _
        ByVal wCmd As Long) As Long
Public Declare Function LockWindowUpdate _
        Lib "user32" _
        (ByVal hwndLock As Long) As Long
Public Declare Function GetDesktopWindow _
        Lib "user32" () As Long
Public Declare Function DestroyWindow _
        Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function TerminateProcess _
        Lib "kernel32" _
        (ByVal hProcess As Long, _
        ByVal uExitCode As Long) As Long
Public Declare Function GetCurrentProcess _
        Lib "kernel32" () As Long
Public Declare Function Putfocus _
        Lib "user32" _
        Alias "SetFocus" _
        (ByVal hwnd As Long) As Long
Private Declare Function SendMessage _
        Lib "user32" _
        Alias "SendMessageA" _
        (ByVal hwnd As Long, _
         ByVal wMsg As Long, _
         ByVal wParam As Long, _
         ByVal lParam As Any) As Long
        

Public Const GW_HWNDNEXT = 2
'public Const WM_QUIT As Long = &H12
Public Const WM_CLOSE = &H10
Public Const WM_SYSCOMMAND As Long = &H112
Public Const SC_CLOSE As Long = &HF060&

Public ash$
Public subattack As Long
Public subattack2 As Long
Public subattackproid As Long
Public subattackproid2 As Long

Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long

Public Const TH32CS_SNAPHEAPLIST As Long = &H1
Public Const TH32CS_SNAPPROCESS As Long = &H2
Public Const TH32CS_SNAPTHREAD As Long = &H4
Public Const TH32CS_SNAPMODULE As Long = &H8
Public Const TH32CS_SNAPALL As Long = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public Const TH32CS_INHERIT As Long = &H80000000

Public lProcessID As Long

Public process() As PROCESSENTRY32
Public Module() As MODULEENTRY32
Public Thread() As THREADENTRY32
Public ProcessID(0 To 999) As Long


Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction&, ByVal uParam&, ByRef lpvParam As Any, ByVal fuWinIni&) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long  'Parameter: form.(property) = SetCursorPos(x,y) in Pixels.  'MODULE 1121

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40




Public Function TerminateForms()
QuitForms = True
TrueTerminate = True

On Local Error Resume Next
    Dim Form As Form
    For Each Form In Forms
        'al = Form.name
        Unload Form
        Set Form = Nothing
    Next Form

End Function


Private Function AlwaysOnTop(ByVal hwnd As Long)  'Makes a form always on top
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Private Function NotAlwaysOnTop(ByVal hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function


Public Function Function_Exist(ByVal sModule As String, ByVal sFunction As String) As Boolean
'on local error GoTo VB_Error
    '// Just stuff for checking funtions.
    Dim hHandle As Long
    
    '// Get the handle
    hHandle = GetModuleHandle(sModule)
    
    If hHandle = 0 Then
        '// Raise not found error
        'If Err.LastDllError <> ERROR_MOD_NOT_FOUND Then Call Err_Dll(Err.LastDllError, "GetModuleHandle failed ::: MOD_NOT_FOUND", sLocation, "Function_Exist")
        
        hHandle = LoadLibraryEx(sModule, 0&, 0&)
        ': If hHandle = 0 Then Call Err_Dll(Err.LastDllError, "LoadLibraryEx failed", sLocation, "Function_Exist")
        
        '// Now see if there is adress
        If GetProcAddress(hHandle, sFunction) = 0 Then
            '  Call Err_Dll(Err.LastDllError, "GetProcAdress failed", sLocation, "Function_Exist")
            Function_Exist = False
        Else
            Function_Exist = True
        End If
        
      '  If FreeLibrary(hHandle) = False Then Call Err_Dll(Err.LastDllError, "FreeLibrary failed", sLocation, "Function_Exist")
    Else
        If GetProcAddress(hHandle, sFunction) = 0 Then
          '    Call Err_Dll(Err.LastDllError, "GetProcAdress failed", sLocation, "Function_Exist")
            Function_Exist = Function_Exist
        Else
            Function_Exist = True
        End If
    End If
    
'Exit Function
'VB_Error:
'    Err_Vb Err.Number, Err.Description, sLocation, "Function_Exist"
'Resume Next
End Function


Public Function Module32_Enum(ByRef Module() As MODULEENTRY32, Optional ByVal lProcessID As Long) As Long
'on local error GoTo VB_Error

    '// To get modules
    
    ReDim Module(0)
    
    If Function_Exist("kernel32.dll", "CreateToolhelp32Snapshot") = True Then
        Dim MODULEENTRY32 As MODULEENTRY32
        Dim hSnapShot As Long
        Dim lModule As Long
        
        hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, lProcessID)
        ': If hSnapShot = INVALID_HANDLE_VALUE Then Call Err_Dll(Err.LastDllError, "CreateToolHelp32Snapshoot failed", sLocation, "Module32_Enum")
        
        MODULEENTRY32.dwSize = Len(MODULEENTRY32)
        If Module32First(hSnapShot, MODULEENTRY32) = False Then
            Module32_Enum = -1
             ' Call Err_Dll(Err.LastDllError, "Module32First failed", sLocation, "Module32_Enum")
            
           ' If CloseHandle(hSnapShot) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Module32_Enum")
            Exit Function
        Else
            ReDim Module(lModule)
            Module(lModule) = MODULEENTRY32
        End If
        
        Do
            If Module32Next(hSnapShot, MODULEENTRY32) = False Then
                Exit Do
            Else

                lModule = lModule + 1
                ReDim Preserve Module(lModule)
                Module(lModule) = MODULEENTRY32
            End If
        Loop
        
       ' If CloseHandle(hSnapShot) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Module32_Enum")   'Call Error_API(Err.LastDllError, sLocation & "\Module32_Enum", "CloseHandle")
        
        Module32_Enum = lModule
    End If
    
'Exit Function
'VB_Error:
'    Err_Vb Err.Number, Err.Description, sLocation, "Module32_Enum"
'Resume Next
End Function

Function InstanceToWnd2(ByVal target_pid As Long) As Long
    Dim test_hwnd As Long, _
        test_pid As Long, _
        test_thread_id As Long
    'Find the first window
    test_hwnd = FindWindow(ByVal 0&, ByVal 0&)
    Do While test_hwnd <> 0
   '    ash$ = GetWindowTitle(test_hwnd)
   '    If InStr(ash$, "Tidal Info Ver") And test_hwnd <> hwnd3 Then
   '     InstanceToWnd2 = test_hwnd: Exit Do
   '    End If
       'Check if the window isn't a child
       If GetParent(test_hwnd) = 0 Then
         'Get the window's thread
         test_thread_id = GetWindowThreadProcessId(test_hwnd, _
                             test_pid)
         If test_pid = target_pid Then
            'InstanceToWnd2 = test_hwnd
            ash$ = GetWindowTitle(test_hwnd)
'            If InStr(ash$, "Tidal Info Ver") Then
            If InStr(ash$, "Tidal_Info") Then
             InstanceToWnd2 = test_hwnd: Exit Do
            End If
            'Exit Do
         End If
       End If
       'retrieve the next window
       test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    Loop
End Function

Function InstanceToWnd(ByVal target_pid As Long) As Long

    Dim test_hwnd As Long, _
        test_pid As Long, _
        test_thread_id As Long
    'Find the first window
    test_hwnd = FindWindow(ByVal 0&, ByVal 0&)
    subattack = 0
    Do While test_hwnd <> 0
       ash$ = GetWindowTitle(test_hwnd)
       If InStr(UCase$(ash$), "TIDAL") Then
       If test_hwnd <> subattack2 Then
       If InStr(ash$, "Tidal_info - Microsoft Visual Basic") = False Then
       subattack = test_hwnd
         test_thread_id = GetWindowThreadProcessId(test_hwnd, _
                             test_pid)
        If test_pid <> subattachproid Then subattackproid2 = test_pid: Exit Do
       
       
       End If
       End If
       End If
       'Check if the window isn't a child
       If GetParent(test_hwnd) = 0 Then
         'Get the window's thread
    '     test_thread_id = GetWindowThreadProcessId(test_hwnd, _
                             test_pid)
        
        ' If test_pid = target_pid Then
        '    InstanceToWnd = test_hwnd
            'Exit Do
     '   If test_pid <> subattachproid Then subattackproid2 = test_pid: Exit Do
         
         
        ' End If
       End If
       'retrieve the next window
       test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    Loop
End Function

'Function GetWindowTitle(ByVal hwnd As Long) As String
'   Dim l As Long
'   Dim S As String
'   l = GetWindowTextLength(hwnd)
'   S = Space(l + 1)
'   GetWindowText hwnd, S, l + 1
'   GetWindowTitle = Left$(S, l)
'End Function




Private Sub Process_Kill(P_ID As Long)
    '// Kill the wanted process
    'Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long

    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, P_ID)
    ': If hProcess = 0 Then Call Err_Dll(Err.LastDllError, "OpenProcess failed", sLocation, "Kill_Process")
    
    If GetExitCodeProcess(hProcess, lExitCode) = False Then a5 = a5
    If TerminateProcess(hProcess, lExitCode) = False Then a5 = a5
    If CloseHandle(hProcess) = False Then a5 = a5
End Sub
Public Function Process32_Enum(ByRef process() As PROCESSENTRY32) As Long
'on local error GoTo VB_Error

    '// Get the most wanted, processes

    ReDim process(0)
    
    Dim PROCESSENTRY32 As PROCESSENTRY32
    Dim hSnapShot As Long
    Dim lProcess As Long
    
    
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    ': If hSnapShot = INVALID_HANDLE_VALUE Then Call Err_Dll(Err.LastDllError, "CreateToolHelp32Snapshoot failed ::: INVALID_HANDLE_VALUE", sLocation, "Process32_Enum")
    
    PROCESSENTRY32.dwSize = Len(PROCESSENTRY32)
    If Process32First(hSnapShot, PROCESSENTRY32) = False Then
        Process32_Enum = -1
          'Call Err_Dll(Err.LastDllError, "Process32First failed", sLocation, "Process32_Enum")
        
       ' If CloseHandle(hSnapShot) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Process32_Enum")
        CloseHandle (hSnapShot) ' = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Process32_Enum")
        Exit Function
    Else
        ReDim process(lProcess)
        process(lProcess) = PROCESSENTRY32
    End If

    Do
        If Process32Next(hSnapShot, PROCESSENTRY32) = False Then
            Exit Do
        Else
            
            lProcess = lProcess + 1
            ReDim Preserve process(lProcess)
            process(lProcess) = PROCESSENTRY32
        End If
    Loop
    
    'If CloseHandle(hSnapShot) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Process32_Enum")   '(Err.LastDllError, sLocation & "\Process32_Enum", "CloseHandle")
    CloseHandle (hSnapShot) ' = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Process32_Enum")   '(Err.LastDllError, sLocation & "\Process32_Enum", "CloseHandle")
    
    Process32_Enum = lProcess
    
Exit Function

'VB_Error:
'    Err_Vb Err.Number, Err.Description, sLocation, "Process32_Enum"
'Resume Next
End Function

Function EnumeratingProcesses()

    'Processes List

    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    uProcess.dwSize = Len(uProcess)
    R = Process32First(hSnapShot, uProcess)
    Do While R
        Ew$ = Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0))
        R = Process32Next(hSnapShot, uProcess)
    Loop
    CloseHandle hSnapShot
End Function



Public Sub List_ActiveProcess()
    '// Same
        
    Dim lCount  As Long
    Dim pFile   As String
    Dim NodeX   As Node
    
    Dim Ew$
    VirusActive = 0
    
    lCount = Process32_Enum(process())
    
    Dim i As Long
    'With TreeView
     '   .Nodes.Clear
        For i = 0 To lCount
           ' Set NodeX = .Nodes.Add(, , CStr("PROCESS_" & I), FileName_Parse(Process(I).szExeFile), 1)
            ProcessID(i) = process(i).th32ProcessID
         '   NodeX.Tag = Process(I).szExeFile

'ew$ = Process(I).szExeFile
'If InStr(ew$, "ccApp.exe") Then
'virusactive = 1
'End If



        Next i
   
End Sub






Public Sub List_ActiveModules()
    '// At the end we have all the modules, listed in the form

    Dim lCount As Long
    Dim pCount As Long
    Dim pFile As String
    Dim NodeX As Node
    Dim i As Long
    Dim O As Long
    Dim lSize As Long
    
    Dim Cmdv As Integer
    Dim Ew$
    
    pCount = Process32_Enum(process())
    
    'nof12 = 0
    Dim egghead As Integer
    egghead = 0
    
    For O = 0 To pCount
        lCount = Module32_Enum(Module(), ProcessID(O))

'If ProcessID(O) = procid3 Then
'ash5$ = ew$
'End If
        
        For i = 0 To lCount
'           With TreeView
              '  Set NodeX = .Nodes.Add("PROCESS_" & O, tvwChild, CStr("MODULE_" & O & "NUM_" & I), Module(I).szModule, 2)
             '   NodeX.Tag = Module(I).szExePath & "\" & Module(I).szModule
           Ew$ = Module(i).szExePath & "\" & Module(i).szModule



'If InStr(UCase$(ew$), "C:\UTILS\TIDAL.EXE") Then
If InStr(UCase$(Ew$), "C:\UTILS\TIDAL.EXE") Then
procid2 = ProcessID(O)
hWnd2 = InstanceToWnd2(procid2)
'hwnd3 = hwnd3
'Dim hProcess As Long
'Dim lExitCode As Long
'hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, procid2)
'If hProcess = 0 Then Call Err_Dll(Err.LastDllError, "OpenProcess failed", sLocation, "Kill_Process")
    

'CloseProgram hProcess, 0, False
End If
           
If InStr(Ew$, "SeriousSam.exe") Then
'nomonoff = 1
'nomusic = 1
processid25ssam = ProcessID(O)
End If
           

If InStr(Ew$, "FRONTPG.EXE") Then
'nomonoff = 1
frontpagepid = ProcessID(O)
egghead = 1
End If

           
           
           
           
 '           End With
        Next i
    Next O

If egghead = 0 Then frontpagepid = 0


End Sub




Public Sub xxxyyy2()

'If xxx2 < 0 Then If DIALER.ShowInTaskbar = False Then dialer.ShowInTaskbar = True
'If xxx2 >= 0 Then If DIALER.ShowInTaskbar = True Then DIALER.ShowInTaskbar = False

End Sub


Function CloseProgram(Optional ByVal ProcessHandle As Long, Optional ByVal ThreadHandle As Long, Optional CloseThread As Boolean = True) As Boolean
  
  Dim ReturnValue As Long
  Dim ExitCode    As Long

  If CloseThread = True Then
    ReturnValue = GetExitCodeThread(ThreadHandle, ExitCode)
    If ReturnValue <> 0 Then
  '    TerminateProcess ProcessHandle, ExitCode
       TerminateThread ThreadHandle, ExitCode
      CloseProgram = True
    Else
      CloseProgram = False
    End If
    
  Else
    ReturnValue = GetExitCodeProcess(ProcessHandle, ExitCode)
    If ReturnValue <> 0 Then
   '   TerminateThread ThreadHandle, ExitCode
       TerminateProcess ProcessHandle, ExitCode
      CloseProgram = True
    Else
      CloseProgram = False
    End If
  End If
  Exit Function
  
End Function

Function frontpage(ByVal target_pid As Long) As Long
    Dim ash$
    Dim test_hwnd As Long, _
        test_pid As Long, _
        test_thread_id As Long
    'Find the first window
    test_hwnd = FindWindow(ByVal 0&, ByVal 0&)
    subattackrealplay = 0
    frontpage = 0
    Do While test_hwnd <> 0
       ash$ = GetWindowTitle(test_hwnd)
       If InStr(ash$, "MatthewLan.Com Web") Then
       frontpage = 1
       Exit Do
       End If
       If InStr(ash$, "Matt.Lan BT Web") And InStr(ash$, "MatthewLan.Com Web") = 0 Then
       frontpage = 2
       Exit Do
       End If
       'Check if the window isn't a child
       If GetParent(test_hwnd) = 0 Then
         'Get the window's thread
        ' test_thread_id = GetWindowThreadProcessId(test_hwnd, _
                             test_pid)
         If test_pid = target_pid Then
            'frontpage = test_hwnd
            'Exit Do
         End If
       End If
       'retrieve the next window
       test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    Loop
End Function




Public Function GetActiveWindowTitle(ByVal ReturnParent As Boolean) As String
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
   GetActiveWindowTitle = GetWindowTitle(i)
End Function

Public Function GetActiveWindow(ByVal ReturnParent As Boolean) As Long
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

