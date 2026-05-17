Attribute VB_Name = "ControlPanelModule"
Option Explicit
Private Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_SCREENSAVE = &HF140&
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
Private Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up

Public Sub SwapMouse(Button As String)
    If Button = "Left" Then
        SwapMouseButton False
    Else
        SwapMouseButton True
    End If
End Sub

Public Sub ClipPointer()
Dim Rect As Rect
    Rect.Left = 0 '((Screen.Width / 2) - 100)
    Rect.Top = 0 '((Screen.Height / 2) - 100)
    Rect.Right = 100 '(Rect.Left + 100)
    Rect.Bottom = 100 '(Rect.Top + 100)
    ClipCursor Rect
End Sub
Public Sub UnClipPointer()
Dim Rect As Rect
    Rect.Left = 0 '((Screen.Width / 2) - 100)
    Rect.Top = 0 '((Screen.Height / 2) - 100)
    Rect.Right = Screen.Width '(Rect.Left + 100)
    Rect.Bottom = Screen.Height '(Rect.Top + 100)
    ClipCursor Rect
End Sub

Public Sub SetMouseDoubleClickTime(ByVal T As Integer)
On Error Resume Next
    SetDoubleClickTime T
End Sub

Public Sub OpenCDDoor()
    mciSendString "Set CDAudio door open", vbNullString, 0&, 0&
End Sub
Public Sub CloseCDDoor()
    mciSendString "Set CDAudio door closed", vbNullString, 0&, 0&
End Sub

Public Sub RunScreenSaver()
    SendMessage GetDesktopWindow(), WM_SYSCOMMAND, SC_SCREENSAVE, 0&
End Sub

Public Function GetUser() As String
Dim User As String, retVal As Long
    User = String(100, Chr$(0))
    GetUserName User, 100
    User = Left(User, InStr(User, Chr(0)) - 1)
    GetUser = User
End Function

Public Function GetLoginTime() As String
Dim MillSec#, Days#, Hours#, Min#, Sec#
Dim Msg As String

    MillSec = GetTickCount()
    Sec = MillSec / 1000
    
    Days = Sec \ (24& * 3600&)
    
    If Days > 0 Then Sec = Sec - (24& * 3600& * Days) 'Not Tested

    Hours = Sec \ 3600
    If Hours > 0 Then Sec = Sec - (3600 * Hours)
    Min = Sec \ 60
    Sec = Sec Mod 60

    Msg = "Windows has been running for" & vbCrLf & vbCrLf
    Msg = Msg & Days & " Day(s), " & Hours & " Hour(s), " & Min & " Minute(s), " & Sec & " Second(s)"
    GetLoginTime = Msg
End Function



Public Function SystemInfo() As String
Dim SysInfo As SYSTEM_INFO
Dim OSVer As OSVERSIONINFO
Dim txtData As String, str As String
    
    'GetSystemInfo SysInfo
    OSVer.dwOSVersionInfoSize = Len(OSVer)
    GetVersionEx OSVer
    
    Select Case OSVer.dwPlatformId
        Case 0
            txtData = "OS Name: " & "Windows 32s" & vbCrLf
        Case 1
            txtData = "OS Name: " & "Windows 95/98" & vbCrLf
        Case 2
            txtData = "OS Name: " & "Windows NT" & vbCrLf
    End Select
    
    txtData = txtData & "Version: " & Trim(OSVer.dwMajorVersion) & "." & Trim(OSVer.dwMinorVersion) & vbCrLf
    txtData = txtData & "OS Manufacturer: Microsoft Corporation" & vbCrLf
    
    str = String(255, Chr(0))
    GetComputerName str, 255
    str = Left(str, InStr(1, str, Chr(0)) - 1)
    txtData = txtData & "System Name: " & str & vbCrLf
    
    str = String(255, Chr(0))
    GetWindowsDirectory str, 255
    str = Left(str, InStr(1, str, Chr(0)) - 1)
    txtData = txtData & "System Directory: " & str & vbCrLf
    
    SystemInfo = txtData
End Function

Public Sub CrezyMouse()
Dim CButt As Long, Extra As Long
    mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0&, 0&, CButt, Extra
End Sub
