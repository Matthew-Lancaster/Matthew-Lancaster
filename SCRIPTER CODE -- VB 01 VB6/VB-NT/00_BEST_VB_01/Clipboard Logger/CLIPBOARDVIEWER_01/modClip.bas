Attribute VB_Name = "modClip"
Option Explicit

Public mNextClip As Long, mPrevHandle As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetClipboardViewer Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ChangeClipboardChain Lib "user32" (ByVal hwnd As Long, ByVal hWndNext As Long) As Long
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_DRAWCLIPBOARD = &H308
Public Const GWL_WNDPROC = (-4)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_USERDATA = (-21)
Public Const WM_LBUTTONDBLCLK = &H203


Sub MAIN()

Load Form1

End Sub

Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Select Case Msg
    Case WM_DRAWCLIPBOARD
        'The clipboard is changed.
        'A trick here, send a double click message to _
        the usercontrol and then raise ClipboardChanged event
        SendMessage hwnd, WM_LBUTTONDBLCLK, 0, 0
        SendMessage mNextClip, Msg, wParam, lParam
    Case WM_CHANGECBCHAIN
        'Another clipboard viewer closed
        If mNextClip = wParam Then
            mNextClip = lParam
        Else
            SendMessage mNextClip, Msg, wParam, lParam
        End If
End Select

WndProc = CallWindowProc(mPrevHandle, hwnd, Msg, wParam, lParam)

End Function

Public Sub SubClass(mHandle As Long, mAddress As Long)

mPrevHandle = GetWindowLong(mHandle, GWL_WNDPROC)
SetWindowLong mHandle, GWL_WNDPROC, mAddress

End Sub

Public Sub UnSubClass(mHandle As Long)

SetWindowLong mHandle, GWL_WNDPROC, mPrevHandle

End Sub
