Attribute VB_Name = "CLEAR_DEBUG_AREA"
Option Explicit


' ------------------------------------------------------------------------------------------
' VB6 CLEAR THE DEBUG AREA - Google Search
' https://www.google.com/search?q=VB6+CLEAR+THE+DEBUG+AREA&oq=VB6+CLEAR+THE+DEBUG+AREA&aqs=chrome..69i57.7772j0j7&sourceid=chrome&ie=UTF-8
' --------
' [RESOLVED] How to clear Immediate Window in IDE-VBForums
' http://www.vbforums.com/showthread.php?672465-RESOLVED-How-to-clear-Immediate-Window-in-IDE
' ----
' [ Wednesday 01:17:20 Am_18 September 2019 ]
' ------------------------------------------------------------------------------------------


Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" _
    (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
     ByVal lpsz2 As String) As Long
Private Declare Function PostMessage Lib "User32" Alias "PostMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
     ByVal lParam As Long) As Long
Private Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const WM_ACTIVATE As Long = &H6

Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_CANCEL = &H3
Private Const VK_CONTROL = &H11

Public Sub ClearInmediateWindow()
Dim lWinVB As Long, lWinEmmediate As Long

    lWinVB = FindWindow("wndclass_desked_gsk", vbNullString)
    'Last param depends on languages, use your inmediate window caption:
    lWinEmmediate = FindWindowEx(lWinVB, ByVal 0&, "VbaWindow", "Immediate")

    PostMessage lWinEmmediate, WM_ACTIVATE, 1, 0&
    keybd_event VK_CANCEL, 0, 0, 0  ' (This is Control-Break)
    keybd_event VK_CONTROL, 0, 0, 0 'Select All
    keybd_event vbKeyA, 0, 0, 0 'Select All
    keybd_event vbKeyDelete, 0, 0, 0 'Clear
    
    keybd_event vbKeyA, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
    keybd_event vbKeyF5, 0, 0, 0   'Continue execution
    keybd_event vbKeyF5, 0, KEYEVENTF_KEYUP, 0
End Sub


