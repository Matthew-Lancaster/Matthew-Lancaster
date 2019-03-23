Option Strict Off
Option Explicit On
Friend Class Form4_DISABLE_CLOSE_BUTTON
	Inherits System.Windows.Forms.Form
	'----
	'Hide Max, Min and Close Buttons - Visual Basic (Microsoft) Win32API - Tek-Tips
	'https://www.tek-tips.com/viewthread.cfm?qid=330760
	'----
	
	Private Declare Function GetWindowLong Lib "user32"  Alias "GetWindowLongA"(ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
	Private Declare Function SetWindowLong Lib "user32"  Alias "SetWindowLongA"(ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
	
	Private Const GWL_STYLE As Short = (-16)
	
	Private Const WS_MAXIMIZEBOX As Integer = &H10000
	Private Const WS_MINIMIZEBOX As Integer = &H20000
	Private Const WS_SYSMENU As Integer = &H80000
	
	Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
	Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Integer) As Integer
	
	Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
	
	Private Const MF_BYCOMMAND As Integer = &H0
	Private Const SC_CLOSE As Integer = &HF060
	
	
	Private Declare Function FindWindow Lib "user32"  Alias "FindWindowA"(ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
	
	Private Sub Form4_DISABLE_CLOSE_BUTTON_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		' -------------------------------------
		' IT'S CODE BUT NOT WORK ON 64 BIT ITEM
		' VB.NET
		' -------------------------------------
		
		'    Dim dwStyle As Long
		'
		'    dwStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
		'    dwStyle = dwStyle And Not WS_MINIMIZEBOX
		'    dwStyle = dwStyle And Not WS_MAXIMIZEBOX
		'    dwStyle = dwStyle And Not WS_SYSMENU
        '    Call SetWindowLong(Me.hWnd, GWL_STYLE, dwStyle)

        Call SET_BUTTON_CHROME()
		
	End Sub
	
    Public Sub SetXState(ByRef hWnd As Long, ByRef blnState As Boolean)
        Dim hMenu As Integer

        'UPGRADE_WARNING: Couldn't resolve default property of object hWnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        hMenu = GetSystemMenu(hWnd, blnState)
        Call RemoveMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
        'UPGRADE_WARNING: Couldn't resolve default property of object hWnd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Call DrawMenuBar(hWnd)
        End
    End Sub
	
	
	Public Sub SET_BUTTON_CHROME()
		Dim XHWND_OFF As Object
		
		'Call Form4_DISABLE_CLOSE_BUTTON.SET_BUTTON_CHROME
		
		Dim dwStyle As Integer
		
		'Dim XHWND_OFF As Long
		
		'UPGRADE_WARNING: Couldn't resolve default property of object XHWND_OFF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		XHWND_OFF = FindWindow("Chrome_WidgetWin_1", vbNullString)
		' XHWND_OFF = FindWindow("Chrome_RenderWidgetHostHWND", vbNullString)
        ' XHWND_OFF = FindWindow("Notepad++", vbNullString)
		
		
		'XHWND_OFF = Form1.hWnd
		
		' MsgBox XHWND_OFF
		
		'UPGRADE_WARNING: Couldn't resolve default property of object XHWND_OFF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		dwStyle = GetWindowLong(XHWND_OFF, GWL_STYLE)
		dwStyle = dwStyle And Not WS_MINIMIZEBOX
		dwStyle = dwStyle And Not WS_MAXIMIZEBOX
		
		' dwStyle = dwStyle And Not WS_SYSMENU
		
		
		'dwStyle = dwStyle And WS_SYSMENU
		'
		'UPGRADE_WARNING: Couldn't resolve default property of object XHWND_OFF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call SetWindowLong(XHWND_OFF, GWL_STYLE, dwStyle)
		

        dwStyle = GetWindowLong(XHWND_OFF, GWL_STYLE)
        ' -------------
		' ENABLE BUTTON
		' -------------
        ' dwStyle = dwStyle Or WS_MINIMIZEBOX
        ' dwStyle = dwStyle Or WS_MAXIMIZEBOX
		
		'UPGRADE_WARNING: Couldn't resolve default property of object XHWND_OFF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Call SetWindowLong(XHWND_OFF, GWL_STYLE, dwStyle)
		
		
		' ------------------------------
		' FALSE DISABLED
		' TRUE ENABLED
		' ------------------------------
		
        Call SetXState(XHWND_OFF, False)
		' Call SetXState(XHWND_OFF, True)
		
		
	End Sub
End Class