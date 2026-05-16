Attribute VB_Name = "Module3"
Option Explicit
Const INFINITE = -1&

' J French - 27th Nov 2003
' Shell and Re-Parent
' hacked from MS and KPD
' Add Two Command Buttons
Public subattackrealplay
Public frontpagepid As Long
Public Winamp24Handle2PL As Long
Public Winamp24Handle2 As Long
Public Winamp22Handle2 As Long
Public oldprocessid24 As Long
Public oldprocessid22 As Long


Public Declare Function ShellExecute _
        Lib "shell32.dll" _
        Alias "ShellExecuteA" _
        (ByVal hWnd As Long, _
         ByVal lpOperation As String, _
         ByVal lpFile As String, _
         ByVal lpParameters As String, _
         ByVal lpDirectory As String, _
         ByVal nShowCmd As Long) As Long

Const SW_SHOWNORMAL = 1

Public Declare Function MoveWindow _
        Lib "user32" _
        (ByVal hWnd As Long, _
         ByVal x As Long, _
         ByVal y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal bRepaint As Long) As Long

Public Declare Function FindWindow _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long
        
Public Declare Function GetParent _
        Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetParent _
        Lib "user32" _
        (ByVal hWndChild As Long, _
         ByVal hWndNewParent As Long) As Long
Public Declare Function GetWindowThreadProcessId _
        Lib "user32" _
        (ByVal hWnd As Long, _
        lpdwProcessId As Long) As Long
Public Declare Function GetWindow _
        Lib "user32" _
        (ByVal hWnd As Long, _
        ByVal wCmd As Long) As Long
Public Declare Function LockWindowUpdate _
        Lib "user32" _
        (ByVal hwndLock As Long) As Long
Public Declare Function GetDesktopWindow _
        Lib "user32" () As Long
Public Declare Function DestroyWindow _
        Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function TerminateProcess _
        Lib "kernel32" _
        (ByVal hProcess As Long, _
        ByVal uExitCode As Long) As Long
Public Declare Function GetCurrentProcess _
        Lib "kernel32" () As Long

Public Declare Function Putfocus _
        Lib "user32" _
        Alias "SetFocus" _
        (ByVal hWnd As Long) As Long
'Public Declare Function SendMessage _
'        Lib "user32" _
'        Alias "SendMessageA" _
'        (ByVal hWnd As Long, _
'         ByVal wMsg As Long, _
'         ByVal wParam As Long, _
'         ByVal lParam As Any) As Long
        
'Public Const GW_HWNDNEXT = 2   ' in another
'public Const WM_QUIT As Long = &H12
Public Const WM_CLOSE = &H10
Public Const WM_SYSCOMMAND As Long = &H112
Public Const SC_CLOSE As Long = &HF060&

Public Type STARTUPINFO
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

Public Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type


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

Public Declare Function CloseHandle Lib "kernel32" _
        (ByVal hObject As Long) As Long

Public Const NORMAL_PRIORITY_CLASS = &H20&


Dim mWnd As Long

Function InstanceToWnd(ByVal target_pid As Long) As Long
    Dim Test_HWND As Long, _
        Test_pid As Long, _
        Test_thread_id As Long
    'Find the first window
    Test_HWND = FindWindow(ByVal 0&, ByVal 0&)
    Do While Test_HWND <> 0
       'Check if the window isn't a child
       If GetParent(Test_HWND) = 0 Then
         'Get the window's thread
         Test_thread_id = GetWindowThreadProcessId(Test_HWND, _
                             Test_pid)
         If Test_pid = target_pid Then
            InstanceToWnd = Test_HWND
            Exit Do
         End If
       End If
       'retrieve the next window
       Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)
    Loop
End Function



Function InstanceToWnd2(ByVal target_pid As Long) As Long
    Dim Window_Title_String
    Dim Test_HWND As Long, _
        Test_pid As Long, _
        Test_thread_id As Long
    'Find the first window
    Test_HWND = FindWindow(ByVal 0&, ByVal 0&)
    subattackrealplay = 0
    Do While Test_HWND <> 0
       Window_Title_String = GetWindowTitle(Test_HWND)
       If Mid$(Window_Title_String, 1, 11) = "RealPlayer:" Then
       subattackrealplay = 1
       End If
       'Check if the window isn't a child
       If GetParent(Test_HWND) = 0 Then
         'Get the window's thread
        ' Test_thread_id = GetWindowThreadProcessId(Test_HWND, _
                             Test_pid)
         If Test_pid = target_pid Then
            InstanceToWnd2 = Test_HWND
            'Exit Do
         End If
       End If
       'retrieve the next window
       Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)
    Loop
End Function


Function frontpage(ByVal target_pid As Long) As Long
    Dim Window_Title_String
    Dim Test_HWND As Long, _
        Test_pid As Long, _
        Test_thread_id As Long
    'Find the first window
    Test_HWND = FindWindow(ByVal 0&, ByVal 0&)
    'subattackrealplay = 0
    frontpage = 0
    Do While Test_HWND <> 0
       Window_Title_String = GetWindowTitle(Test_HWND)
       If InStr(Window_Title_String, "MatthewLan.Com Web") Then
       frontpage = 1
       Exit Do
       End If
       If InStr(Window_Title_String, "Matt.Lan BT Web") And InStr(Window_Title_String, "MatthewLan.Com Web") = 0 Then
       frontpage = 2
       Exit Do
       End If
       'Check if the window isn't a child
       If GetParent(Test_HWND) = 0 Then
         'Get the window's thread
        ' Test_thread_id = GetWindowThreadProcessId(Test_HWND, _
                             Test_pid)
         If Test_pid = target_pid Then
            'frontpage = Test_HWND
            'Exit Do
         End If
       End If
       'retrieve the next window
       Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)
    Loop
End Function








Public Function ExecCmd(cmdLine$) As Long
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
    Call WaitForSingleObject(proc.hProcess, 500)
    If Ret Then
       ExecCmd = InstanceToWnd(proc.dwProcessId)
       'Me.Print Ret, ExecCmd
       Call CloseHandle(proc.hThread)
       Call CloseHandle(proc.hProcess)
    End If
    
End Function

'OR THIS
 
Public Function ShellAndWait(ByVal fileName As String) As Long
     Dim executionStatus As Long
     Dim ProcessHandle As Long
     Dim ReturnValue As Long
     'Execute the application/file
     'executionStatus = Shell(fileName, vbNormalFocus)
     executionStatus = Shell(fileName, vbHide)

     'Get the Process Handle
     ProcessHandle = OpenProcess(&H100000, True, executionStatus)

     'Wait till the application is finished
     ReturnValue = WaitForSingleObject(ProcessHandle, INFINITE)

     'Send the Return Value Back
     ShellAndWait = ReturnValue

 End Function

