VERSION 5.00
Begin VB.Form Form4_DISABLE_CLOSE_BUTTON 
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
Attribute VB_Name = "Form4_DISABLE_CLOSE_BUTTON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'From
'Mon 15-Oct-2018 21:43:38
'TO
'Tue 16-Oct-2018 02:42:49
'----
'Hide Max, Min and Close Buttons - Visual Basic (Microsoft) Win32API - Tek-Tips
'https://www.tek-tips.com/viewthread.cfm?qid=330760
'----

'----
'Classic VB - How can I Hide/Unhide the Form's "X" Button-VBForums
'http://www.vbforums.com/showthread.php?368754-Classic-VB-How-can-I-Hide-Unhide-the-Form-s-quot-X-quot-Button
'----

'----
'Hide Max, Min and Close Buttons - Visual Basic (Microsoft) Win32API - Tek-Tips
'https://www.tek-tips.com/viewthread.cfm?qid=330760
'----

'----
'Warn when Closing Multiple Open Tabs in Chrome, Firefox, Edge and Internet Explorer • Raymond.CC
'https://www.raymond.cc/blog/warn-when-closing-multiple-opened-tabs-in-google-chrome/
'----

'----
'Lock it - Chrome Web Store
'https://chrome.google.com/webstore/detail/lock-it/magjmoeipkknhmpcojpeomlfgaofhfho?hl=en-GB
'--------
'Chrome Close Lock - Chrome Web Store
'https://chrome.google.com/webstore/detail/chrome-close-lock/plcabbfeeokakkmdecdccmibahigjkno?hl=en-GB
'--------
'Don 't Close Window With Last Tab - Chrome Web Store
'https://chrome.google.com/webstore/detail/dont-close-window-with-la/dlnpfhfhmkiebpnlllpehlmklgdggbhn?hl=en-GB
'--------
'disable the close button on a form vb6 - Google Search
'https://www.google.co.uk/search?q=disable+the+close+button+on+a+form+vb6&oq=DISABLE+THE+CLOSE+BUTTON+&aqs=chrome.4.0j69i57j0l4.9013j0j7&sourceid=chrome&ie=UTF-8
'----

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE = (-16)

Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_SYSMENU = &H80000

Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu _
 As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal HWND _
As Long) As Long
 
Private Declare Function GetSystemMenu Lib "user32" (ByVal HWND _
As Long, ByVal bRevert As Long) As Long

Private Const MF_BYCOMMAND = &H0&
Private Const SC_CLOSE = &HF060&

Private Const MF_BYPOSITION = &H400&

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long

Dim XHWND_OFF As Long



Private Sub Form_Load()

' -------------------------------------
' IT'S CODE BUT NOT WORK ON CHROME
' VB.NET
' -------------------------------------

' XHWND_OFF = FindWindow("Chrome_WidgetWin_1", vbNullString)
' XHWND_OFF = FindWindow("Chrome_RenderWidgetHostHWND", vbNullString)
' XHWND_OFF = FindWindow("Notepad++", vbNullString)
' XHWND_OFF = FindWindow("HashMyFiles", vbNullString)

XHWND_OFF = FindWindow("Chrome_WidgetWin_1", vbNullString)

Call SET_BUTTON_CHROME(HWND)

End


XHWND_OFF = FindWindow("Chrome_WidgetWin_1", vbNullString)
' XHWND_OFF = FindWindow("Notepad++", vbNullString)
Call DisableCloseButton(XHWND_OFF)

' Call SetXState(XHWND_OFF, True)


'Call SET_BUTTON_CHROME


End

End Sub

Public Sub SetXState(HWND As Long, blnState As Boolean)
    Dim hMenu As Long
 
    hMenu = GetSystemMenu(HWND, blnState)
    Call RemoveMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
    Call DrawMenuBar(HWND)
 
End Sub


Public Sub SET_BUTTON_CHROME(HWND As Long)

    'Call Form4_DISABLE_CLOSE_BUTTON.SET_BUTTON_CHROME

    Dim dwStyle As Long

    'Dim XHWND_OFF As Long

    XHWND_OFF = HWND
    'XHWND_OFF = Form1.hWnd

    dwStyle = GetWindowLong(XHWND_OFF, GWL_STYLE)
    dwStyle = dwStyle And Not WS_MINIMIZEBOX
    dwStyle = dwStyle And Not WS_MAXIMIZEBOX
    
    dwStyle = dwStyle And Not WS_SYSMENU
    
    
    'dwStyle = dwStyle And WS_SYSMENU
'
    Call SetWindowLong(XHWND_OFF, GWL_STYLE, dwStyle)

    dwStyle = GetWindowLong(XHWND_OFF, GWL_STYLE)
    
    ' -------------
    ' ENABLE BUTTON
    ' -------------
    'dwStyle = dwStyle Or WS_MINIMIZEBOX
    'dwStyle = dwStyle Or WS_MAXIMIZEBOX
    
    ' dwStyle = dwStyle Or WS_MINIMIZEBOX
    ' dwStyle = dwStyle Or WS_MAXIMIZEBOX
'
    Call SetWindowLong(XHWND_OFF, GWL_STYLE, dwStyle)


    ' ------------------------------
    ' FALSE DISABLED
    ' TRUE ENABLED
    ' ------------------------------
    
    Call SetXState(XHWND_OFF, False)
    ' Call SetXState(XHWND_OFF, True)


End Sub


Public Sub DisableCloseButton(HWND As Long)
'----
'FreeVBCode code snippet: Disable the Close Button on a Form
'http://www.freevbcode.com/ShowCode.asp?ID=2448
'----
'ANOTHER WAY TO REMOVE CLOSE BUTTON
'NOT MUCH HELP PUT IT BACK THIS WAY

'PURPOSE: Removes X button from a form
'EXAMPLE: DisableCloseButton Me
'RETURNS: True if successful, false otherwise
'NOTES:   Also removes Exit Item from
'         Control Box Menu


    Dim lHndSysMenu As Long
    Dim lAns1 As Long, lAns2 As Long
    
    
    lHndSysMenu = GetSystemMenu(HWND, 0)

    'remove close button
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)

   'Remove seperator bar
    ' lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
    
    'Return True if both calls were successful
    ' DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)

End Sub
