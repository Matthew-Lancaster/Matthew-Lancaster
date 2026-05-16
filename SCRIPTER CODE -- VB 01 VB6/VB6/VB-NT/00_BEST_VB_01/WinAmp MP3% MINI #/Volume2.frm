VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form WinAmpMP3MINI 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "WinAmp_MP3_MINI"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "Volume2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   2175
      Top             =   1680
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2460
      Top             =   1005
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "WinAmpMP3MINI"
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
      Interval        =   100
      Left            =   1650
      Top             =   945
   End
End
Attribute VB_Name = "WinAmpMP3MINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FS, XT, OFULLSCREEN, GFW, OGFW, LockMove, ENDE
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
Dim METOP, MELEFT, Show2, HH
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

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function


'MATT-555ROIDS
'55-88-HAPPY


'----------------------------------------------------

    
Public Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function

Public Function NotAlwaysOnTop(ByVal hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function

Private Sub Form_Load()

If App.PrevInstance = True Then End

ProgressBar1.ToolTipText = App.Path + " -- " + App.EXEName

Me.Caption = App.EXEName

Set FS = CreateObject("Scripting.FileSystemObject")
FIRSTRUN = True
Me.Visible = False

If App.PrevInstance = True Then End
ProgressBar1.Min = 0
ProgressBar1.Max = 100
Me.Show

METOP = 270
MELEFT = 680
If GetComputerName = "55-88-HAPPY" Then
    METOP = 20
    MELEFT = 10
End If


Me.Top = METOP
Me.Left = MELEFT
ProgressBar1.Top = 0
ProgressBar1.Left = 0
ProgressBar1.Height = 140
ProgressBar1.Width = 1950
Me.Width = ProgressBar1.Width
Me.Height = ProgressBar1.Height



End Sub

Sub PosiXY()

XT = False
Win_Main = FindWindow("Winamp v1.x", vbNullString)


msgresult = SendMessage(Win_Main, WM_USER, 0, ByVal ipc_isplaying)
'// IPC_ISPLAYING returns the status of playback.
'// If it returns 1, it is playing. if it returns 3, it is paused,
'// if it returns 0, it is not playing.


OGFW = GFW

'SAVE THIS BIT LATER
If msgresult = 0 Then Exit Sub


End Sub

Private Sub Form_Resize()

Exit Sub

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

ENDE = ENDE + 1

If ENDE > 2 Then End

End Sub


Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
    End
End If

End Sub

Private Sub Timer1_Timer()


Win_Main = FindWindow("Winamp v1.x", vbNullString)

If Win_Main = 0 Then End

'Filespec = "U:\Temp\MP3%.txt"
'Set F = FS.GetFile(Filespec)

'Adate1 = F.datelastmodified
'If Adate1 > Bdate1 Then
'    On Error Resume Next
'    Open Filespec For Input Lock Write As #1
'    MP3 = Val(Input(LOF(1), 1))
'    Close #1
'End If

msgresult = SendMessage(Win_Main, WM_USER, 0, ByVal ipc_isplaying)


If msgresult = 0 Then Exit Sub

On Error Resume Next


'Lenght Song
MsgResult3 = SendMessage(Win_Main, WM_USER, 1, ByVal IPC_GETOUTPUTTIME)

'Posi in Song
MsgResult2 = SendMessage(Win_Main, WM_USER, 0, ByVal IPC_GETOUTPUTTIME)


If MsgResult3 = 0 Then Exit Sub

ProgressBar1.Max = MsgResult3 * 1000

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


Private Sub Timer2_Timer()
'qq = GetForegroundWindow
'If qq = ForeG Then Exit Sub
'ForeG = qq

'If FindWindow(vbNullString, "Extreme Lock Build: 2011") > 0 Then
''    Exit Sub
 '   Me.Top = Screen.Height - (Me.Height) - 10
 '   Me.Left = 10
'
'Else

Me.Top = METOP
Me.Left = Left

'

'End If
AlwaysOnTop (Me.hWnd)
'NotAlwaysOnTop (ATidalX.hWnd)

End Sub

Private Sub Timer3_Timer()
ENDE = 0
End Sub
