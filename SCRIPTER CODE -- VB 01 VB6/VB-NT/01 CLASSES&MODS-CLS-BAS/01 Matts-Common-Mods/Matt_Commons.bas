Attribute VB_Name = "Matt_Commons"
'------------------------------------
'Outlook_Express_Philo

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long




Public URL2$
Public URLLoc$


      Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
         lpRect As RECT) As Long

      Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
         ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
         ByVal nHeight As Long, ByVal bRepaint As Long) As Long

      Declare Function GetDesktopWindow Lib "user32" () As Long

      Declare Function EnumWindows Lib "user32" _
         (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

      Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent _
         As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

      Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId _
         As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long

      Declare Function GetWindowThreadProcessId Lib "user32" _
         (ByVal hwnd As Long, lpdwProcessId As Long) As Long

      Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
         (ByVal hwnd As Long, ByVal lpClassName As String, _
         ByVal nMaxCount As Long) As Long

      Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
         (ByVal hwnd As Long, ByVal lpString As String, _
         ByVal cch As Long) As Long

      Public TopCount As Integer     ' Number of Top level Windows
      Public ChildCount As Long   ' Number of Child Windows
      Public ThreadCount As Long  ' Number of Thread Windows

Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWndParent As Long, _
     ByVal hWndChildAfter As Long, _
     ByVal lpszClass As String, _
     ByVal lpszTitle As String) _
    As Long


Public Hwnd22



Public CompName$
Public Rabbit As Date
Public Cheese


Public dks As Integer
Public sluty As Integer
Public sluty2 As Integer
Public sluty3 As Integer
Public ebuyer22 As Integer
Public rnowt2 As Date

Public OldRect2 As RECT
Public OldRect3 As RECT
Public OldRect4 As RECT
Public OldRect5 As RECT
Public OldRect6 As RECT

Public Abcd255 As Integer
Public Abcd355 As Integer
Public Abcd455 As Integer
Public Abcd655 As Integer

Public OldAbcd As Integer

Public SoggyPie As Integer

Public QuitForms As Integer
Public Soff As Integer
Public Toff As Integer

Public Gooze1
Public Gooze2
Public Gooze3
Public Gooze4
Public Gooze5
Public Gooze6
Public Gooze7
Public Gooze8

Public Qkey
Public Qkey2 As Date

Public Count2Check

Public Nb() As Long
Public ab$()

Public Enum WindowStates
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


Public ash2$

Public Const GW_CHILD = 5


Public URLadt$
Public URLadt2$



Public Function GetWindowState(ByVal lngHWND As Long) As Integer
    If IsWindow(lngHWND) = 0 Then
        GetWindowState = -1
    ElseIf IsIconic(lngHWND) <> 0 Then
        GetWindowState = vbMinimized
    ElseIf IsZoomed(lngHWND) <> 0 Then
        GetWindowState = vbMaximized
    End If
End Function



