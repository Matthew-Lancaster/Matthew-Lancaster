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
'----
'Hide Max, Min and Close Buttons - Visual Basic (Microsoft) Win32API - Tek-Tips
'https://www.tek-tips.com/viewthread.cfm?qid=330760
'----

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE = (-16)

Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_SYSMENU = &H80000

Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu _
 As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd _
As Long) As Long
 
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd _
As Long, ByVal bRevert As Long) As Long

Private Const MF_BYCOMMAND = &H0&
Private Const SC_CLOSE = &HF060&


Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long

Private Sub Form_Load()

'    Dim dwStyle As Long
'
'    dwStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
'    dwStyle = dwStyle And Not WS_MINIMIZEBOX
'    dwStyle = dwStyle And Not WS_MAXIMIZEBOX
'    dwStyle = dwStyle And Not WS_SYSMENU
'    Call SetWindowLong(Me.hWnd, GWL_STYLE, dwStyle)

End Sub

Public Sub SetXState(hWnd, blnState As Boolean)
    Dim hMenu As Long
 
    hMenu = GetSystemMenu(hWnd, blnState)
    Call RemoveMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
    Call DrawMenuBar(hWnd)
 
End Sub


Public Sub SET_BUTTON_CHROME()

    'Call Form4_DISABLE_CLOSE_BUTTON.SET_BUTTON_CHROME

    Dim dwStyle As Long

    'Dim XhWnd_OFF As Long

    XhWnd_OFF = FindWindow("Chrome_WidgetWin_1", vbNullString)
    ' XhWnd_OFF = FindWindow("Chrome_RenderWidgetHosthWnd", vbNullString)
    ' XhWnd_OFF = FindWindow("Notepad++", vbNullString)
    

    'XhWnd_OFF = Form1.hWnd
    
    ' MsgBox XhWnd_OFF

    dwStyle = GetWindowLong(XhWnd_OFF, GWL_STYLE)
    dwStyle = dwStyle And Not WS_MINIMIZEBOX
    dwStyle = dwStyle And Not WS_MAXIMIZEBOX
    
    ' dwStyle = dwStyle And Not WS_SYSMENU
    
    
    'dwStyle = dwStyle And WS_SYSMENU
'
    Call SetWindowLong(XhWnd_OFF, GWL_STYLE, dwStyle)

    dwStyle = dwStyle Or WS_MINIMIZEBOX
    dwStyle = dwStyle Or WS_MAXIMIZEBOX
    
    Call SetWindowLong(XhWnd_OFF, GWL_STYLE, dwStyle)


    ' ------------------------------
    ' FALSE DISABLED
    ' TRUE ENABLED
    ' ------------------------------
    
    Call SetXState(XhWnd_OFF, False)
    ' Call SetXState(XhWnd_OFF, True)


End Sub
