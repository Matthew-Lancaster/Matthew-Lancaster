Attribute VB_Name = "TaskManagerModule"
Option Explicit

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long

Private TaskList As String

Function GetTaskList() As String
    TaskList = ""
    'TaskList = "GetTaskList#"
    EnumWindows AddressOf EnumWindowsProc, ByVal 0&
    GetTaskList = TaskList
End Function

Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
Dim sSave As String, Ret As Long
    
    Ret = GetWindowTextLength(hwnd)
    sSave = Space(Ret)
    GetWindowText hwnd, sSave, Ret + 1
    
    TaskList = TaskList & sSave
    If Len(sSave) > 0 Then
        If IsIconic(hwnd) Then
            TaskList = TaskList & "=Minimize#"
        ElseIf IsZoomed(hwnd) Then
            TaskList = TaskList & "=Maximize#"
        ElseIf IsWindowVisible(hwnd) Then
            TaskList = TaskList & "=Not Active#"
        ElseIf IsWindowEnabled(hwnd) Then
            TaskList = TaskList & "=Active#"
        End If
    End If
    EnumWindowsProc = True
End Function

Sub KillTask(TaskName As String)
Dim hwnd As Long
Const WM_CLOSE = &H10
Const WM_ENDSESSION = &H16
    hwnd = FindWindow(vbNullString, TaskName)
    PostMessage hwnd, WM_CLOSE, 0&, 0&
End Sub
