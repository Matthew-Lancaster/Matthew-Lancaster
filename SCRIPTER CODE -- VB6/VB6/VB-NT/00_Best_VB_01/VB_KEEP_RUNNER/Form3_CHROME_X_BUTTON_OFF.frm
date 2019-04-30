VERSION 5.00
Begin VB.Form Form3_CHROME_X_BUTTON_OFF 
   Caption         =   "Form4"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form2"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form3_CHROME_X_BUTTON_OFF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----
'Classic VB - How can I Hide/Unhide the Form's "X" Button-VBForums
'http://www.vbforums.com/showthread.php?368754-Classic-VB-How-can-I-Hide-Unhide-the-Form-s-quot-X-quot-Button
'----

Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd _
As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu _
 As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd _
As Long) As Long
 
Private Const MF_BYCOMMAND = &H0&
Private Const SC_CLOSE = &HF060&
 
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long

 
Public Sub SetXState(hWnd, blnState As Boolean)
    Dim hMenu As Long
 
    hMenu = GetSystemMenu(hWnd, blnState)
    Call RemoveMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
    Call DrawMenuBar(hWnd)
 
End Sub

Public Sub SET_CHROME_WINDOW_X_BUTTON_CLOSE_TO_OFF()

XhWnd_OFF = FindWindow("Chrome_WidgetWin_1", vbNullString)

XhWnd_OFF = Me.hWnd

Call SetXState(XhWnd_OFF, False)

End Sub
