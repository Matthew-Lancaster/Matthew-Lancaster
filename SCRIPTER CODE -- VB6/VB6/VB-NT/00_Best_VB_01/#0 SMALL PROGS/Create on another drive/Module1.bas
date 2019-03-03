Attribute VB_Name = "Module1"
'CREATE DIR ON ANOTHER DRIVE CODE

Public FS

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long



'Public Declare Function ShellExecute _
        Lib "shell32.dll" _
        Alias "ShellExecuteA" _
        (ByVal hwnd As Long, _
         ByVal lpOperation As String, _
         ByVal lpFile As String, _
         ByVal lpParameters As String, _
         ByVal lpDirectory As String, _
         ByVal nShowCmd As Long) As Long

Const SW_SHOWNORMAL = 1

'Public Declare Function MoveWindow _
'        Lib "user32" _
'        (ByVal hwnd As Long, _
'         ByVal x As Long, _
'         ByVal y As Long, _
'         ByVal nWidth As Long, _
'         ByVal nHeight As Long, _
'         ByVal bRepaint As Long) As Long

' ------------------------------------------------------------------------
' ------------------------------------------------------------------------

' ------- Find the first window
' ------- MISSING DLL VERSION



Public Declare Function FindWindow _
        Lib "user32.dll" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long


'Private Declare Function FindWindow _
'        Lib "user32" _
'        Alias "FindWindowA" _
'        (ByVal lpClassName As Long, _
'        ByVal lpWindowName As Long) As Long
'
'' ------------------------------------------------------------------------
'' ------------------------------------------------------------------------
'
Private Declare Function GetParent _
        Lib "user32" (ByVal hWnd As Long) As Long
        
'Public Declare Function SetParent _
        Lib "user32" _
        (ByVal hWndChild As Long, _
        ByVal hWndNewParent As Long) As Long

Private Declare Function GetWindowThreadProcessId _
        Lib "user32" _
        (ByVal hWnd As Long, _
        lpdwProcessId As Long) As Long
        
Private Declare Function GetWindow _
        Lib "user32" _
        (ByVal hWnd As Long, _
        ByVal wCmd As Long) As Long
'Public Declare Function LockWindowUpdate _
'        Lib "user32" _
'        (ByVal hwndLock As Long) As Long
'Public Declare Function GetDesktopWindow _
'        Lib "user32" () As Long
'Public Declare Function DestroyWindow _
'        Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function TerminateProcess _
'        Lib "kernel32" _
'        (ByVal hProcess As Long, _
'        ByVal uExitCode As Long) As Long
'Public Declare Function GetCurrentProcess _
'        Lib "kernel32" () As Long
'
'Public Declare Function Putfocus _
'        Lib "user32" _
'        Alias "SetFocus" _
'        (ByVal hwnd As Long) As Long
'Public Declare Function SendMessage _
'        Lib "user32" _
'        Alias "SendMessageA" _
'        (ByVal hWnd As Long, _
'         ByVal wMsg As Long, _
'         ByVal wParam As Long, _
'         ByVal lParam As Any) As Long
        
Const GW_HWNDNEXT = 2   ' in another
'Public Const WM_QUIT As Long = &H12
'Public Const WM_CLOSE = &H10
'Private Const WM_SYSCOMMAND As Long = &H112
'Public Const SC_CLOSE As Long = &HF060&

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
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
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type


'Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA"
Private Declare Function CreateProcessA _
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

Private Declare Function WaitForSingleObject _
        Lib "kernel32" _
        (ByVal hHandle As Long, _
        ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
        (ByVal hObject As Long) As Long

Const NORMAL_PRIORITY_CLASS = &H20&

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Const INFINITE = -1&

'EXCUTE AND WAIT TO SHOW WORKS HERE

Public Function ExecCmd(cmdLine$) As Long
    Dim Test_HWND
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim Ret&

     ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)

    ' Start the shelled application:
    Ret& = CreateProcessA(vbNullString, cmdLine$, _
                          0&, 0&, 1&, _
                          NORMAL_PRIORITY_CLASS, _
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
    
    'WAIT TO SHOW SILL REQUIRED
    'if you Want an EXPLOREr SelEct And already got same folder
    'open from before
    'you won't get a new instance window
    'Even So it Does get Brought to Foreground from a Minmize
    'But Won't Do Select Folder Your Request or Makevisible
    'that Select
    
    If ExecCmd > 0 Then
        Do
        'If IsWindowVisible(ExecCmd) = True Then Exit Do
        If IsWindow(ExecCmd) = 1 Then Exit Do
        
        Sleep 100
        Loop Until 1 = 2
    End If
    
    
    'Cant do this FIND HWND FROM PROCESS INFO
    If ExecCmd = 0 Then
        cmdLine$ = "HWND Not Available"
        ExecCmd = 0
'        SHALL WE WAIT FOR AN EXPLORER TO ARRIVE IN FOREGROUND
'        H PROCESS OPIATE PROCESS
'        FORE GROUND - FORE ROUND - ROUND WINDOW - RAINBOW TV

'
'        '\\Linda-pc\d acer\VB6-EXE'S SYNC
'       ExecCmd = InstanceToWnd(proc.hProcess)
'       'Me.Print Ret, ExecCmd
'
'       Call CloseHandle(proc.hThread)
'       Call CloseHandle(proc.hProcess)
        
    End If
    
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

Public Function ExecCmdWAIT(cmdLine$) As Long
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim Ret&

     ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)

    ' Start the shelled application:
    Ret& = CreateProcessA(vbNullString, cmdLine$, _
                          0&, 0&, 1&, _
                          NORMAL_PRIORITY_CLASS, _
                          0&, vbNullString, _
                          start, proc)
    
    
    
    
    ' --- let it start - this seems important
'    Call WaitForSingleObject(proc.hProcess, 500)
    
    Call WaitForSingleObject(proc.hProcess, INFINITE)
    
    If Ret Then
       ExecCmdWAIT = InstanceToWnd(proc.dwProcessId)
       'Me.Print Ret, ExecCmd
       
       Call CloseHandle(proc.hThread)
       Call CloseHandle(proc.hProcess)
    End If
    
    
    
End Function

'OR THIS

Public Function shellAndWait(ByVal fileName As String) As Long
    Dim executionStatus As Long
    Dim ProcessHandle As Long
    Dim ReturnValue As Long
    'Execute the application/file
    'executionStatus = Shell(fileName, vbNormalFocus)
    executionStatus = Shell(fileName, vbNormalFocus)

    'Get the Process Handle
    ProcessHandle = OpenProcess(&H100000, True, executionStatus)

    'Wait till the application is finished
    ReturnValue = WaitForSingleObject(ProcessHandle, INFINITE)

    'Send the Return Value Back
    shellAndWait = ReturnValue

   'THERE IS NO PROPER EXIT CODE
   'JUST IF LAUNCH OKAY

End Function

Public Function shellAndWaitToShow(ByVal fileName As String) As Long

'EXCUTE AND WAIT TO SHOW WORKS HERE
'Public Function ExecCmd(cmdLine$) As Long

End Function


Function InstanceToWnd(ByVal target_pid As Long) As Long
    Dim Test_HWND As Long, _
        test_pid As Long, _
        test_thread_id As Long
    'Find the first window
    Test_HWND = FindWindow(ByVal 0&, ByVal 0&)
    Do While Test_HWND <> 0
       'Check if the window isn't a child
       If GetParent(Test_HWND) = 0 Then
         'Get the window's thread
         test_thread_id = GetWindowThreadProcessId(Test_HWND, target_pid)
         If test_pid = target_pid Then
            InstanceToWnd = Test_HWND
            Exit Do
         End If
       End If
       'retrieve the next window
       Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)
    Loop
End Function


'FIND IF A EXPLORER SELCT HAS BEEN BROUGHT FORWARD
Function OneForeGroundWnd(ByVal target_Class As String) As Long
    Dim Test_HWND As Long, _
        test_pid As Long, _
        test_id As Long
    'Find the first window
    Test_HWND = FindWindow(ByVal 0&, ByVal 0&)
    Do While Test_HWND <> 0
       'Check if the window isn't a child
       If GetParent(Test_HWND) = 0 Then
         'Get the window's thread
         test_id = GetForegroundWindow
         
         If GetWindowClass(test_id) = target_Class Then
            OneForeGroundWnd = Test_HWND
            Exit Do
         End If
       End If
       'retrieve the next window
       Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)
    Loop
End Function


Function GetWindowClass(ByVal hWnd As Long) As String
    Dim Ret As Long, sText As String
    sText = Space(255)
    Ret = GetClassName(hWnd, sText, 255)
    sText = Left$(sText, Ret)
    GetWindowClass = sText
End Function

