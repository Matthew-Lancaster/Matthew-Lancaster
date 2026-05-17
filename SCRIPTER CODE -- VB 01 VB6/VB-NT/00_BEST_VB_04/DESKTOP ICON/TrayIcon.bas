Attribute VB_Name = "TraIconDeclares"
'::::::::::::::::::::::::::::::::::::::::::
Public Const DM_DISPLAYFLAGS = &H200000
Public Const DM_DISPLAYFREQUENCY = &H400000
Public Const CDS_FORCE As Long = &H80000000
Public Const CDS_RESET = &H40000000
Public Const HWND_BROADCAST = &HFFFF&
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_PALETTECHANGED = &H311
Public Const WM_DISPLAYCHANGE = &H7E
Public Const WM_SETTINGCHANGE = &H1A
Public Const BITSPIXEL = 12
Public Const HORZRES = 8
Public Const VERTSIZE = 6
Public Const WM_WININICHANGE = &H1A
Public Const GWL_WNDPROC = (-4)
Public Const WM_USER = &H400
Public Const SW_HIDE = 0    ' Hide Window
Public Const SW_SHOW = 5    ' Show Window
  
Public Const NIM_ADD = 0
Public Const NIM_MODIFY = 1
Public Const NIM_DELETE = 2
Public Const NIF_MESSAGE = 1
Public Const NIF_ICON = 2
Public Const NIF_TIP = 4

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200
Public Const WM_NULL = &H0
Public Const WM_PAINT = &HF
Public Const WM_SYSCOMMAND = &H112
Public Const SC_HOTKEY = &HF150&
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const MOD_WIN = &H8

Type NOTIFYICONDATA
  cbSize              As Long
  hwnd                As Long
  uID                 As Long
  uFlags              As Long
  uCallbackMessage    As Long
  hIcon               As Long
  szTip               As String * 64
End Type

Declare Function Shell_NotifyIconA Lib "shell32" _
(ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Declare Function PostMessage Lib "user32" _
 Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
 ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long

'Subclassing declares
Public OldWindowProc As Long
Public Declare Function FindWindowEx Lib "user32" _
    Alias "FindWindowExA" (ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
    ByVal nCmdShow As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetCurrentTime Lib "kernel32" Alias "GetTickCount" () As Long
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Public changeWidth As Long
Public changeHeight As Long

Public pixelDepth As Long
Public InIde As Long
Public WeDidIt As Long

Public Function NewWindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    *** Warning ***
'Subclassing is dangerous. The debugger does not work well when a new
'WindowProc is installed. If you halt the program instead of unloading its
'forms, it will crash and so will the Visual Basic IDE. Save your work often!
'Don't say you weren't warned.


Select Case msg
Case WM_WININICHANGE
    RefreshIconPositions
    Form1.Text1.Text = "WM_WININICHANGE detected"
    NewWindowProc = CallWindowProc( _
        OldWindowProc, Form1.hwnd, msg, wParam, _
        lParam)
Case WM_PALETTECHANGED
NewWindowProc = 0
'DO NOT PROCESS MESSAGE AFTER RECEIVING WM_PALETTECHANGED
 
Case WM_HOTKEY 'the one and only global hotkey for our window


        
        Form1.WindowState = 0
        Form1.Show
'we won't pass this message   no need


Case WM_DISPLAYCHANGE 'the display resolution has changed

NewWindowProc = CallWindowProc( _
        OldWindowProc, Form1.hwnd, msg, wParam, _
        lParam)
'::::::::::::::::::::::::::::::::::
If WeDidIt = 0 Then
    
    'lparam contains width and height vars of change
     
    Dim rcVar As POINTAPI2
    CopyMemory rcVar, lParam, 4
    changeWidth = rcVar.x
    changeHeight = rcVar.y
    ChangeForm.Show
    
    
    
    Pause 3000 'give the screen time to settle
    
    Dim lprect As RECT
    Dim DeskHwnd As Long
    DeskHwnd = GetDesktopWindow()
    ss& = GetWindowRect(DeskHwnd, lprect) 'get the size
    nwidth& = lprect.Right             'of the Desktop
    nheight& = lprect.Bottom
    nbits% = GetDeviceCaps(Form1.hdc, BITSPIXEL)
    PlaceIconsInPosition nbits%, nwidth&, nheight&  ' restore icons to new positions
    Form1.Text1.Text = "Detected mode change"
    Unload ChangeForm
Else 'we did it and we have already done the above
        
    WeDidIt = 0
End If


Case Else
'hand messages back to default handler
    NewWindowProc = CallWindowProc( _
        OldWindowProc, Form1.hwnd, msg, wParam, _
        lParam)
End Select



End Function
Public Sub TaskBarHideUnhide()
    Dim lTaskBarHWND As Long
    Dim lRet As Long
    ' Show / hide the taskbar
    lTaskBarHWND = FindWindow("Shell_TrayWnd", "")
    gg& = IsWindowVisible(lTaskBarHWND)
    
    If gg& = 1 Then 'visible
        lRet = ShowWindow(lTaskBarHWND, SW_HIDE)
        
        Pause 100
        ' get the desktop device context for debug print to desktop
        hwndSrc& = 0
        hSrcDC& = GetDC(hwndSrc&) 'same as GetDc(Null)  Desktop dc
        
        If InIde = 0 Then
        Test$ = " TaskBar Hidden by DeskTop Icon Tools 3.02  - ***WARNING*** HOTKEY DOES NOT WORK IN VB IDE. "
        Else
        Test$ = " TaskBar Hidden by DeskTop Icon Tools 3.02  -  Press Ctrl-Shift-D to bring Control Window Up. "
        End If
        
        zs& = TextOut(hSrcDC&, 0, StartHeight - 20, Test$, Len(Test$))
        ' **** IMPORTANT **** release resources back to Windows
        Dmy& = ReleaseDC(hwndSrc&, hSrcDC&)
    Else 'hidden
        lRet = ShowWindow(lTaskBarHWND, SW_SHOW)
    End If
    
    If lRet < 0 Then
    '
    ' Handle error from api
    '
    End If

   
End Sub

Public Sub DeskTopHide(ByVal bShow As Boolean) 'NOT USED
    Dim lDesktopHwnd As Long
    Dim lFlags As Long
'
' Show / Hide the Desktop Icons  ACTUALLY the whole Listview
'

    On Error Resume Next

    lDesktopHwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)

    If lDesktopHwnd = 0 Then
        ' raise an error ! You have no desktop !!!
        Exit Sub
    End If
    
    lFlags = IIf(bShow, SW_SHOW, SW_HIDE)
    
    'ShowWindow lDesktopHwnd, lFlags
    ShowWindow 0&, lFlags
    
End Sub

Public Sub Pause(mSeconds As Long)

Dim CurrentTime
CurrentTime = GetCurrentTime
Do
ssd = DoEvents()
Loop While GetCurrentTime - CurrentTime <= mSeconds Or GetCurrentTime - CurrentTime < 0



End Sub
Public Sub PlaceIconsInPosition(pixeldepth2 As Integer, nwidth2 As Long, nheight2 As Long)  ' change pixel depth
If nwidth2 = CurrentWidth And nheight2 = CurrentHeight Then
'Same Width and Height don't change Icon position
Else
 Form1.Text1.Text = "Restored Icons"
    Dim i As Long
    Dim diff2 As Single
    Dim diff As Single
                     'restore to fit icons into the
    Dim tempone      'same pattern we found in FindIcons
    Dim temptwo      'but scaled to fit the screen
    diff = StartWidth / (nwidth2)
    diff2 = StartHeight / (nheight2)
    
    For i = 0 To icount - 1
        tempone = CLng(IconPosition(i).x / diff)
        temptwo = CLng(IconPosition(i).y / diff2)
        Call SendMessageByLong(hdesk, LVM_SETITEMPOSITION, i, _
        CLng(tempone + temptwo * &H10000))
    Next i

End If

'This damn project must owe me a couple of hundred hrs now
'Still here?
Form1.Text2.Text = Str(nwidth2) + " *" + Str(nheight2) + "  " + Str(pixeldepth2) + " bit"
FindIcons   'reload vars and show the user
            
End Sub


Function MakeWord(ByVal bHi As Byte, ByVal bLo As Byte) As Integer 'NOT USED

    If bHi And &H80 Then
        MakeWord = (((bHi And &H7F) * 256) + bLo) Or &H8000
    Else
        MakeWord = (bHi * 256) + bLo
    End If
End Function

Function MakeDWord(wHi As Integer, wLo As Integer) As Long 'NOT USED

    If wHi And &H8000& Then
        MakeDWord = (((wHi And &H7FFF&) * 65536) Or (wLo And &HFFFF&)) Or &H80000000
    Else
        MakeDWord = (wHi * 65535) + wLo
    End If
End Function


