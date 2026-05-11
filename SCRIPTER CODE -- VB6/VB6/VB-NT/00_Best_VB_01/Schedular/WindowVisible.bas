Attribute VB_Name = "WindowVisible"
Private Const SW_HIDE = 0
  Private Const SW_SHOW = 5
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

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" _
    (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_CHILD = 5



Property Let WindowVisible(hwnd As Long, New_Value As Boolean)
    If hwnd = 0 Then Exit Property
    tt = cProcesses.Convert(hwnd, otss, cnFromhWnd Or cnToProcessID)
    If otss <> Totter And XShow1 = False Then Exit Property

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



