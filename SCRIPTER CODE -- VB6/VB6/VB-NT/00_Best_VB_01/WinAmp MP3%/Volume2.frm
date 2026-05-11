VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form WinAmpMP3 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "WinAmp_MP3_VideoBar"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   Icon            =   "Volume2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1650
      Top             =   945
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   375
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "WinAmpMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XSS As Long, TT5
Dim TSHOW

Private Enum WindowStates
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

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long


Public FS, XT, OFULLSCREEN, GFW, OGFW, LockMove
Public OX, OY, ODit, Dit, HAMMER, MEUP, TAG3, TAG4, FIRSTRUN, FIRSTRUN2, FULLSCREEN, Win_Main, Win_Video
'HAMMER HAPPY
Dim Rt6 As RECT
Dim Rt3 As RECT
Dim Rect2 As RECT
Dim Tag2
Private Const GW_HWNDNEXT = 20
Private Const GW_CHILD = 5

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" _
        (lpPoint As POINTAPI) As Long
             


Dim Filespec, Adate1, Bdate1, MP3
Dim MeTop, Show2, HH
Public GG2
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Public ForeG, ForeG2
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MoveWindow _
        Lib "user32" _
        (ByVal hWnd As Long, _
         ByVal x As Long, _
         ByVal y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal bRepaint As Long) As Long



Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
    
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect _
   As RECT) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


'-----------------------------------------------------

Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long  'MODULE 1141
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long

Private Const MONITOR_ON = -1&
Private Const MONITOR_OFF = 2&
Private Const SC_MONITORPOWER = &HF170&
Private Const WM_SYSCOMMAND = &H112
Private Const ipc_isplaying As Long = 104
Private Const IPC_GETPLAYLISTFILE  As Long = 211
Private Const IPC_GETOUTPUTTIME  As Long = 105
Private Const IPC_GETINFO As Long = 126
Private Const WINAMP_BUTTON2 = 40045
Private Const WINAMP_BUTTON3 = 40046
Private Const WM_CLOSE = &H10
Private Const WM_USER = &H400
Private Const WM_COMMAND = &H111
'Private Const GW_HWNDNEXT = 2
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long




Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer



'----------------------------------------------------

    
Public Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function

Public Function NotAlwaysOnTop(ByVal hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function


Private Sub Form_Load()



If App.PrevInstance = True Then End



XX2 = FindWindow(vbNullString, "WinAmp_MP3_VideoBar")
If XX2 > 0 And XX2 <> Me.hWnd Then
    MsgBox "ALEADY RUNNING FROM EXE MAY BE PROBLEM CODING"
    
End If


TT5 = Now + TimeSerial(0, 10, 0)

'If GetComputerName = "55-88-HAPPY" Then End

App.Title = App.EXEName + ".exe"


Set FS = CreateObject("Scripting.FileSystemObject")
FIRSTRUN = True
Me.Visible = False
ProgressBar1.Min = 0
ProgressBar1.Max = 100
ProgressBar2.Min = 0
ProgressBar2.Max = 100
'Me.Show

ProgressBar1.ToolTipText = "PBAR 1 OF 2 -- " + App.Path + " -- " + App.EXEName
ProgressBar2.ToolTipText = "PBAR 2 OF 2 -- " + App.Path + " -- " + App.EXEName



'Me.Width = 0
Me.Height = 650

'Call PosiXY



AlwaysOnTop (Me.hWnd)
'NotAlwaysOnTop (ATidalX.hWnd)
Call OpenMixer(0)
Call Timer1_Timer
Show2 = False


End Sub

Sub PosiXY()



XT = False
Win_Main = FindWindow("Winamp v1.x", vbNullString)
'If Win_Main = 0 Then End
If Win_Main = 0 Then
    Me.Hide
    'TSHOW = 0
    'ShowWindow Me.hWnd, SW_HIDE
    Exit Sub
End If

Win_Video = FindWindow("Winamp Video", "Winamp Video")

If Win_Video > 0 Then XT = IsWindowVisible(Win_Video): If XT > 1 Then Stop
'If XT = 0 Then End
If XT = 0 Then
    Me.Hide
    'TSHOW = 0
    'ShowWindow Me.hWnd, SW_HIDE
    
    Exit Sub
End If
    

If Win_Video > 0 And XT = False Then
    Win_Main = FindWinPart("Winamp Video")
    If Win_Video > 0 Then XT = 1
End If


If Win_Video > 0 Then msgresult = SendMessage(Win_Video, WM_USER, 0, ByVal ipc_isplaying)
If msgresult = 0 And Win_Main > 0 Then
'b = Win_Video: Win_Video = Win_Main: Win_Main = b
msgresult = SendMessage(Win_Main, WM_USER, 0, ByVal ipc_isplaying)

End If
'// IPC_ISPLAYING returns the status of playback.
'// If it returns 1, it is playing. if it returns 3, it is paused,
'// if it returns 0, it is not playing.

If msgresult = 0 Then XT = 0: Win_Video = 0

If msgresult = 1 Then
    
TT5 = Now + TimeSerial(0, 2, 0)
'Call Form_Resize2
'AlwaysOnTop (Me.hWnd)

'Me.Show 'ShowWindow Me.hWnd, SW_SHOW
'Call Form_Resize
'Me.Top
'Me.LEFT
'ME.HEIGHT
'Me.WIDTH


End If



If Win_Main > 0 Then pow = GetWindowState(Win_Main) = vbMinimized

If (XT = 0 And Show2 = True) Or pow = True Or Win_Main = 0 Then
    Me.Hide: Show2 = False
    'Var o= NotAlwaysOnTop(Me.hWnd)
End If
If Win_Video = False Or XT = 0 Then Exit Sub
If pow = True Then
    Exit Sub
End If
If (XT = 1 And Show2 = False) Then
    TSHOW = True
    'Me.Show:
    Show2 = True
    'Var = NotAlwaysOnTop(Me.hWnd)
End If

GFW = GetForegroundWindow
If OGFW <> GFW And GFW = Win_Video Or GFW = Win_Main Then
If Me.Visible = True Then
If GetAsyncKeyState(1) Then
    LockMove = Now + Timer + 0.3
     'Exit Sub
End If
  
If LockMove < Now + Timer Then
    Me.SetFocus
End If

Exit Sub
End If
End If

OGFW = GFW

'SAVE THIS BIT LATER
If msgresult = 0 Then Exit Sub


'If xt = 0 Then Exit Sub
'If xt = 0 Then Stop

'DONT LIKE THIS to much
're = GetWindowRect(Win_Video, Rect2)
'jj = Win_Video + Rect2.Left + Rect2.Top + Rect2.Right + Rect2.Bottom
'If jj + XT = ForeG Then Exit Sub
'ForeG = jj + XT

'xt = FindWindow("Winamp Video", "Winamp Video")
If Win_Video > 0 Then
    
    
    'XT6 = FindWindow("Winamp Video", vbNullString)
    re = GetWindowRect(Win_Video, Rt6)
    
    FULLSCREEN = False
    Tag2 = 0
    If (Rt6.Left = 0 Or Rt6.Top = 0) And (Rt6.Right = 1280 Or Rt6.Bottom = 800) Then
        Tag2 = 1
        FULLSCREEN = True
        'XT = FindWindow(vbNullString, "WinAmp_MP3")
        Var = AlwaysOnTop(Me.hWnd)
    
    End If
    
    If OFULLSCREEN <> FULLSCREEN Then
        If FULLSCREEN = False Then
            'XT = FindWindow(vbNullString, "WinAmp_MP3")
            Var = AlwaysOnTop(Me.hWnd)
        End If
    End If
    
    OFULLSCREEN = FULLSCREEN
    
    
    If Tag2 = 0 Then HAMMER = Now - 1
    
    If Show2 = False Then
        Me.Show: Show2 = True
        DoEvents
        WinAmpMP3.Visible = True
        'If FULLSCREEN = True Then Var = AlwaysOnTop(Me.hWnd)
    End If
    
    
    Call FindCursor2(x, y)
    If OX <> x Or OY <> y And Tag2 = 1 Then
        If HAMMER = 0 Then MEUP = 28 '- Me.Hide
        HAMMER = Now + TimeSerial(0, 0, 2)
    End If
    If HAMMER < Now And HAMMER > 0 Then
        HAMMER = 0
        MEUP = 0
        'Me.Show
    End If
    OX = x: OY = y
    'If Tag2 = 0 Then MEUP = 0
    If FIRSTRUN = True Then MEUP = 0: FIRSTRUN = False
    
    
    'HwndCTLTask4 = FindWindow("MOM Class", vbNullString)
    HwndCTLTask4 = FindWindow(vbNullString, "WindowsFormsParkingWindow")

    re = GetWindowRect(HwndCTLTask4, Rt3)
    HGT1 = (Rt3.Bottom - Rt3.Top) - 5
    HwndCTLTask2 = FindWindow("Shell_TrayWnd", vbNullString)
    re = GetWindowRect(HwndCTLTask2, Rt3)
    HGT2 = Rt3.Bottom - Rt3.Top
    
    re = GetWindowRect(Win_Main, Rt3)
    'If Win_Main <> 20908472 Then Stop
    wid = (Screen.Width / Screen.TwipsPerPixelX) - Rt3.Right
    
    

    
    'XT! is Video Screen
    'move the winamp player
    HGT = (Screen.Height / Screen.TwipsPerPixelY) - (HGT1 + (HGT2)) '+ 300))
    HGT = HGT - Me.Height / Screen.TwipsPerPixelY
    HGT = HGT  '+ 20 '+30 FOR AJUSTMENT
    'If Tag2 = 0 Then MoveWindow Win_Video, Rt3.Right, HGT1 - 45, wid, HGT + 30, True: TAG3 = 0
    If Tag2 = 0 Then MoveWindow Win_Video, Rt3.Right, HGT1, wid, HGT, True: TAG3 = 0
    XSS = Rt3.Right
    
    'HIH = Me.Height / Screen.TwipsPerPixelX
    
    XT = FindWindow(vbNullString, Me.Caption)
    
    If Tag2 = 0 Then
        'This wll move our program about in switch full screen
        'NOT FULL SCREEN
        op3 = 1
        
        'Dim RT3VIDEO As RECT
        're = GetWindowRect(Win_Main, RT3VIDEO)
        
        'OUR VB WINDOW
        're = GetWindowRect(XT, Rt3)
        'MoveWindow XT, Rt3.Right + 1, 769 - hih, wid, hih, True: TAG4 = 0
        
        ADJUST = 10
        Var = AlwaysOnTop(Me.hWnd)
        
        
        Me.Height = 590
        Me.Top = (HGT + 25) * Screen.TwipsPerPixelY
        Me.Width = wid * Screen.TwipsPerPixelX
        Me.Left = Rt3.Right * Screen.TwipsPerPixelX
        
        Call Form_Resize2
        If TSHOW = True Then
            TSHOW = 0
            Me.Show
            DoEvents
            AlwaysOnTop (Me.hWnd)
        End If
        
        'MoveWindow XT, XSS, HGT + ADJUST, wid, Rt3.Bottom - Rt3.Top, True: TAG4 = 0
        
        
        'If Rt3.Right <> 275 Then Stop
        'End If
        'MoveWindow xt, 0, 800 - (Me.Height / Screen.TwipsPerPixelY) - MEUP, 1280, hih, True
    
    Else
        
        'This wll move our program about in switch full screen
        'Me.Visible = False
        'If in Full Screen dont work
        MoveWindow XT, 0, 800 - (Me.Height / Screen.TwipsPerPixelY) - MEUP, 1280, HIH, True
        
    End If
    
    
    End If


End Sub

Private Sub Form_Resize2()

On Error Resume Next

ProgressBar1.Top = 0
ProgressBar1.Left = 0
ProgressBar1.Height = (Me.Height / 2) + 10
If Me.Width > 0 Then ProgressBar1.Width = Me.Width
ProgressBar2.Top = ProgressBar1.Height - 10
ProgressBar2.Left = ProgressBar1.Left
ProgressBar2.Height = (Me.Height / 2) + 10
ProgressBar2.Width = ProgressBar1.Width

End Sub

Private Sub ProgressBar1_Click()
If GG2 = 0 Then GG2 = Now + TimeSerial(0, 0, 1): HH = 0
If GG2 < Now Then GG2 = 0: HH = 0: Exit Sub
HH = HH + 1
If HH = 2 Then End
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
    End
End If

End Sub

Private Sub ProgressBar2_Click()
Call ProgressBar1_Click
End Sub


Private Sub ProgressBar2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
    End
End If

End Sub

Private Sub Timer1_Timer()

Timer1.Interval = 50

YY = 0
If FindWindow(vbNullString, "WinAmp_MP3_VideoBar - Microsoft Visual Basic [design] - [WinAmpMP3 (Code)]") > 0 Then YY = 1
If FindWindow(vbNullString, "WinAmp_MP3_VideoBar - Microsoft Visual Basic [break] - [WinAmpMP3 (Code)]") > 0 Then YY = 1
If FindWindow(vbNullString, "WinAmp_MP3_VideoBar - Microsoft Visual Basic [run] - [WinAmpMP3 (Code)]") > 0 Then YY = 1
If YY > 0 And IsIDE = False Then End


If FIRSTRUN2 = True Then Call PosiXY

If TT5 < Now Then End

FIRSTRUN2 = True

'Filespec = "U:\Temp\MP3%.txt"
'Set F = FS.GetFile(Filespec)

'Adate1 = F.datelastmodified
'If Adate1 > Bdate1 Then
'    On Error Resume Next
'    Open Filespec For Input Lock Write As #1
'    MP3 = Val(Input(LOF(1), 1))
'    Close #1
'End If

If Win_Video > 0 Then msgresult = SendMessage(Win_Main, WM_USER, 0, ByVal ipc_isplaying)
If msgresult = 0 And Win_Main > 0 Then
    'b = Win_Video: Win_Video = Win_Main: Win_Main = b
    msgresult = SendMessage(Win_Main, WM_USER, 0, ByVal ipc_isplaying)
End If

If Win_Video = 0 Or Win_Main = 0 Or msgresult = 0 Then Exit Sub

On Error Resume Next

'If Err.Number > 0 Then Exit Sub

'Bdate1 = Adate1


'Lenght Song
MsgResult3 = SendMessage(Win_Main, WM_USER, 1, ByVal IPC_GETOUTPUTTIME)

'Posi in Song
MsgResult2 = SendMessage(Win_Main, WM_USER, 0, ByVal IPC_GETOUTPUTTIME)


If MsgResult3 = 0 Then Exit Sub

ProgressBar1.Max = MsgResult3 * 1000

ProgressBar2.Value = GetVolume(SPEAKER)
'ProgressBar1.Value = Abs(MP3)
ProgressBar1.Value = MsgResult2


End Sub


Public Sub FindCursor2(x, y)
Dim P As POINTAPI
GetCursorPos P
x = P.x
y = P.y
End Sub






Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim l As Long
   Dim s As String
   l = GetWindowTextLength(hWnd)
   s = Space(l + 1)
   GetWindowText hWnd, s, l + 1
   GetWindowTitle = Left$(s, l)
End Function



Function FindWinPart(Search) As Long
FindWinPart = False
Dim Ash$, HUGE As Long
Dim Test_Hwnd As Long, test_pid As Long, test_thread_id As Long
Dim cText As String

'Find the first window
'need this is you gona use this procedure get from CIDRun ME an another one
Test_Hwnd = FindWindow2(ByVal 0&, ByVal 0&)
Do While Test_Hwnd <> 0
    DoEvents
    teWin_Video = GetWindowClass(Test_Hwnd) 'Mwnd will get reset here if none
    If teWin_Video <> "" And InStr(teWin_Video, "Winamp Video") > 0 Then
        XT = IsWindowVisible(Test_Hwnd)
        If XT = 1 Then
            HUGE = Test_Hwnd
            Win_Video = HUGE
            Exit Do
        End If
    End If
    Test_Hwnd = GetWindow(Test_Hwnd, GW_HWNDNEXT)
Loop

If HUGE = 0 Then
    Win_Video = 0
    Win_Main = 0
    'End
    Exit Function
End If


Test_Hwnd = Win_Video
Do While Test_Hwnd <> 0
    DoEvents
    teWin_Video = GetWindowClass(Test_Hwnd) 'Mwnd will get reset here if none

    If teWin_Video <> "" And InStr(teWin_Video, "Winamp v1.x") > 0 Then
        HUGE = Test_Hwnd
        Win_Main = HUGE
        Exit Do
    End If
    Test_Hwnd = GetWindow(Test_Hwnd, GW_HWNDNEXT)
Loop

FindWinPart = Win_Main
'a = a

 Exit Function



End Function



Function GetActiveWindowClass(ReturnParent As Long) As String
    GetActiveWindowClass = GetWindowClass(ReturnParent)
End Function

Function GetWindowClass(ByVal hWnd As Long) As String
    Dim ret As Long, sText As String
    sText = Space(255)
    ret = GetClassName(hWnd, sText, 255)
    sText = Left$(sText, ret)
   GetWindowClass = sText
End Function






Public Function GetFileFromHwnd(lngHwnd) As String

'messageBox getfilefromhwnd(Me.hwnd)

Dim lngProcess&, hProcess&, bla&, C&
Dim strFile As String
Dim x

strFile = String$(256, 0)
'Temp Out
'x = GetWindowThreadProcessId(lngHwnd, lngProcess)
'hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
'x = EnumProcessModules(hProcess, bla, 4&, C)
'C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
'GetFileFromHwnd = Left(strFile, C)

End Function




Public Function GetWindowState(ByVal lngHwnd As Long) As Integer
    If IsWindow(lngHwnd) = 1 Then
        GetWindowState = -1
        If IsIconic(lngHwnd) <> 0 Then
            GetWindowState = vbMinimized
        ElseIf IsZoomed(lngHwnd) <> 0 Then
            GetWindowState = vbMaximized
        End If
    End If
End Function


'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
'IsIDE = False
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function

'***********************************************
'***********************************************

