Attribute VB_Name = "WindowVisible"
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Declare Function GetWindowThreadProcessId _
        Lib "user32" _
        (ByVal hwnd As Long, _
        lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16

Private Const MAX_PATH = 260

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long

Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal x As Long, ByVal y As Long, _
        ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
        Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME


Private Const PROCESS_ALL_ACCESS = &H1F0FFF

'Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" _
    (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3



'Private Const PROCESS_QUERY_INFORMATION = &H400
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


Private sEXEName As String
Private hProcess As Long
Private lProcessID As Long
Private bIsNT4 As Boolean
Private bIsOpen As Boolean


'Private Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Any, ByVal lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
'Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpExitCode As Long) As Long
'Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetSystemMenu Lib "user32.dll" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


Public Enum WindowStates
  SW_HIDE = 0
  SW_SHOWNORMAL = 1
  SW_normal = 1
  SW_SHOWMINIMIZED = 2
  SW_SHOWMAXIMIZED = 3
  SW_MAXIMIZE = 3
  SW_SHOWNOACTIVATE = 4
  SW_SHOW = 5
  SW_MINIMIZE = 6
  SW_SHOWMINNOACTIVE = 7
  SW_SHOWNA = 8
  sw_RESTORE = 9
  SW_SHOWDEFAULT = 10
  SW_FORCEMINIMIZE = 11
  SW_MAX = 11
End Enum





Property Let WindowVisible(hwnd As Long, New_Value As Boolean)
    If hwnd = 0 Then Exit Property
    'TT = cProcesses.Convert(hwnd, otss, cnFromhWnd Or cnToProcessID)
    'If otss <> Totter And XShow1 = False Then Exit Property

    Call ShowWindow(hwnd, IIf(New_Value, SW_SHOW, SW_HIDE))

End Property

Property Get WindowVisible(hwnd As Long) As Boolean

    If hwnd = 0 Then Exit Property
    'tt = cProcesses.Convert(hwnd, otss, cnFromhWnd Or cnToProcessID)
    'If otss <> Totter And XShow1 = False Then Exit Sub
    WindowVisible = (IsWindowVisible(hwnd) > 0)

End Property

Public Function FindWindowLike(Caption As String) As Long

  Dim Length As Long
  Dim CurrWnd As Long
  Dim ListItem As String
 
  CurrWnd = GetWindow(GetActiveWindow, GW_HWNDFIRST)
  Do While (CurrWnd <> 0)
    Length = GetWindowTextLength(CurrWnd)
   
    If (Length > 0) Then
      ListItem = Space$(Length + 1)
      Length = GetWindowText(CurrWnd, ListItem, Length + 1)
      ListItem = Left$(ListItem, Length)
     
      If (ListItem Like Caption) Then
        FindWindowLike = CurrWnd
        Exit Function
      End If
    End If
   
    CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
  Loop

End Function


Private Sub Form_Load3()
App.TaskVisible = False
    Dim lhWnd As Long
 
    lhWnd = FindWindow("ALT:0062B0E0", vbNullString)      ''''''''' HERE I PUT THE CLASS NAME "alt...."
    If (lhWnd > 0) Then
      WindowVisible(lhWnd) = False
    Else
      Call MsgBox("No window could be found.", vbExclamation)
      Exit Sub
    End If
    End
End Sub



Public Function GetFileFromHwnd(lngHwnd) As String

Dim lngProcess&, hProcess&, bla&, C&
Dim strFile As String, FileN
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
'Dim lngProcess&
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



