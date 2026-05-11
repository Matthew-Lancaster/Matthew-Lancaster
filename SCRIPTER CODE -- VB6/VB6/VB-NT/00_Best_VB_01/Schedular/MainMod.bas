Attribute VB_Name = "MainMod"
Public ART1(), ART2(), ART3()

Public Totter As Long, FS

Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwProcessID As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const PROCESS_TERMINATE = &H1

Public process() As PROCESSENTRY32
Public Module() As MODULEENTRY32
Public Thread() As THREADENTRY32
Public ProcessID(0 To 999) As Long

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

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID  As Long
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



'Private Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Any, ByVal lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetSystemMenu Lib "user32.dll" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public cProcesses As New clsCnProc


Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142

Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

'Public cProcesses As New clsCnProc

Public XShow1


Public Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long

    
Public JH(70)
    
Const GW_HWNDNEXT = 2

Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2

Public TotalAll As Long

Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wmessage As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_CLOSE = &H10


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

Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long

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



Function FindWinPart(tt$, xx1) As Long
FindWinPart = 0

Dim ash$
Dim Test_Hwnd As Long, _
    Test_PID As Long, _
    test_thread_id As Long

Dim cText As String
'Dim Rect8 As RECT

'Find the first window
xx = 0
'need this is you gona use this procedure get from CIDRun ME an another one
Test_Hwnd = FindWindow2(ByVal 0&, ByVal 0&)
br$ = ""
Do While Test_Hwnd <> 0
    
        'hwnd9 = GetWindowRect(test_hwnd, Rect8)
        ash$ = GetWindowTitle(Test_Hwnd)
'        If InStr(ash$, "Double") > 0 Then Stop
'        gws = GetWindowState(test_hwnd)
'        If gws = -1 Then gws2$ = "Window State Normal"
'        If gws = vbMinimized Then gws2$ = "Window State Minimized"
'        If gws = vbMaximized Then gws2$ = "Window State Maximized"

    
    'If (Rect8.Top > 0 And Rect8.Left > 0) Or gws = vbMinimized Then
        ash$ = GetWindowTitle(Test_Hwnd)
        If ash$ <> "" And InStr(ash$, tt$) > 0 Then
            FindWinPart = Test_Hwnd
            xx = xx + 1
            If xx = xx1 Then Exit Do
        End If
    'End If
        
'retrieve the next window
Test_Hwnd = GetWindow(Test_Hwnd, GW_HWNDNEXT)

Loop
'MsgBox Str(huge) + " Windows Brought To Front"

End Function

Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim l As Long
   Dim S As String
   l = GetWindowTextLength(hwnd)
   S = Space(l + 1)
   GetWindowText hwnd, S, l + 1
   GetWindowTitle = Left$(S, l)
End Function


Public Sub Process_Kill(P_ID As Long)
    '// Kill the wanted process
    
    Dim hProcess As Long
    Dim lExitCode As Long
    'NORMAL_PRIORITY_CLASS
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, P_ID) ': If hProcess = 0 Then Call Err_Dll(Err.LastDllError, "OpenProcess failed", sLocation, "Kill_Process")
    
    DD = GetExitCodeProcess(hProcess, lExitCode) '= False 'Then Call Err_Dll(Err.LastDllError, "GetExitCodeProcess failed", sLocation, "Kill_Process")
    DD = TerminateProcess(hProcess, lExitCode) '= False 'Then Call Err_Dll(Err.LastDllError, "TerminateProcess failed", sLocation, "Kill_Process")
    
    DD = CloseHandle(hProcess) '= False 'Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Kill_Process")
End Sub


