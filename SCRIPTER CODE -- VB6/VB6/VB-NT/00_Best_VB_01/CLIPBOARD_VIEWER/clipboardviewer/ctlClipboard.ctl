VERSION 5.00
Begin VB.UserControl ctlClipboard 
   CanGetFocus     =   0   'False
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   765
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ctlClipboard.ctx":0000
   ScaleHeight     =   750
   ScaleWidth      =   765
   ToolboxBitmap   =   "ctlClipboard.ctx":030A
End
Attribute VB_Name = "ctlClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*******************************************************
' Name      : Clipboard Viewer
' Purpose   : Monitor the activity of clipboard.
'           : It is more efficiently than using a Timer
'           : to monitor the clipboard.
' Author    : Season (season_@hotmail.com)
'*******************************************************

Private mStarted As Boolean

Event ClipboardChanged()

Private Sub UserControl_DblClick()

'This dblclick event is only called by modClip.bas _
because the control is invisible at runtime
RaiseEvent ClipboardChanged

End Sub

Private Sub UserControl_Resize()

UserControl.Width = 480
UserControl.Height = 480

End Sub

Private Sub UserControl_Terminate()

EndClipboardViewer

End Sub

Public Sub StartClipboardViewer()

If Not mStarted Then
    'The SetClipboardViewer() adds the specified window _
    to the chain of clipboard viewers. Clipboard viewer _
    windows receive a WM_DRAWCLIPBOARD message whenever _
    the content of the clipboard changes.
    'Subclass the window to get the WM_DRAWCLIPBOARD message.
    
    mNextClip = SetClipboardViewer(UserControl.hwnd)
    SubClass UserControl.hwnd, AddressOf WndProc
    mStarted = True
End If

End Sub

Public Sub EndClipboardViewer()

If mStarted Then
    'The ChangeClipboardChain() removes a specified window _
    from the chain of clipboard viewers.
    
    UnSubClass UserControl.hwnd
    ChangeClipboardChain UserControl.hwnd, mNextClip
    mStarted = False
End If

End Sub
