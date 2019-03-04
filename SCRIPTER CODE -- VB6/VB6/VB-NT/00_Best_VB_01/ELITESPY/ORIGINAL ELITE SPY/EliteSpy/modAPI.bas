Attribute VB_Name = "modAPI"
'--------------------------------------------------------------------------------
'    Component  : modAPI
'    Project    : EliteSpy
'
'    Description: Module for API declarations
'
'    Author     : Andrea Batina
'    Modified   : 31/10/2001
'--------------------------------------------------------------------------------
Option Explicit

' Public types
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

' Public API declarations
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

' Public constants for GetWindowLong API declaration
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

' Public constants for window styles
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_CHILD = &H40000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_SYSMENU = &H80000
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const WM_CLOSE = &H10

' Private constants for extended window styles
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_STATICEDGE = &H20000

' Public constants for SetWindowPos API declaration
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

' Public constants for ShowWindow API declaration
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5

' Public constants for set window on top
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

' Public constants for CreateToolhelp32Snapshot API declaration
Public Const TH32CS_SNAPPROCESS = &H2&

Private m_sPattern As String
Private m_lhFind As Long

Public Function EnumWinProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim k As Long
    Dim sName As String
    
    ' Fill buffer
    sName = Space$(128)
    ' Get window caption
    k = GetWindowText(hwnd, sName, 128)
    If k > 0 Then
        ' Trim caption
        sName = Left$(sName, k)
        If lParam = 0 Then sName = UCase(sName)
        ' If names match
        If sName Like m_sPattern Then
            ' Return window handle
            m_lhFind = hwnd
            ' Exit function
            EnumWinProc = 0
            Exit Function
        End If
    End If
    EnumWinProc = 1
End Function

Public Function FindWindowHwnd(sWild As String) As Long
    ' Save window title
    m_sPattern = UCase$(sWild)
    ' enumerate all open windows
    EnumWindows AddressOf EnumWinProc, False
    ' Return window handle
    FindWindowHwnd = m_lhFind
End Function

Public Sub GenerateCode(ByVal nCodeType As Integer, ByVal sWindowName As String)
    Dim sFile As String
    Dim F As Integer
    Dim sData As String
        
    ' Get selected code type
    Select Case nCodeType
        Case 0: sFile = "FindWindow.esc"
        Case 1: sFile = "CloseWindow.esc"
        Case 2: sFile = "MinimizeWindow.esc"
        Case 3: sFile = "MaximizeWindow.esc"
        Case 4: sFile = "ChangeTitle.esc"
        Case 5: sFile = "ActivateWindow.esc"
        Case 6: sFile = "EnableDisableWindow.esc"
        Case 7: sFile = "HideShowWindow.esc"
    End Select
    
    With frmgeneratedCode
        ' Load file
        .txtCode.LoadFile App.Path & "\Codes\" & sFile, rtfRTF
        ' Replace temp string with window title
        .txtCode.TextRTF = Replace(.txtCode.TextRTF, "+++---+++", sWindowName)
        ' Remove window from top
        SetWindowPos frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        frmMain.chkOnTop.Value = 0
        ' Show form
        .Show , frmMain
    End With
End Sub
