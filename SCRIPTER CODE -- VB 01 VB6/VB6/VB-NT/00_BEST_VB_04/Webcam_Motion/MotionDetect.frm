VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "WebCam Motion Capture"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   720
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   7800
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Timer TimerRAPID 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7080
      Top             =   3120
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7080
      Top             =   1680
   End
   Begin VB.Timer MouseKeyboardDetectPress 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7080
      Top             =   2640
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7080
      Top             =   2160
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7080
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7080
      Top             =   720
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      DrawWidth       =   3
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   6630
      TabIndex        =   0
      Top             =   600
      Width           =   6690
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Picture Image Not Captured"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3810
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "left click to record, right click to stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.Menu MNU_MENU 
      Caption         =   "MENU"
      Begin VB.Menu MNU_SYSTEM_PROPERTIES 
         Caption         =   "System Properties - Enable Disable the Webcam If Problem"
      End
      Begin VB.Menu Mnu_Task_Schedualr 
         Caption         =   "Task Schedular"
      End
      Begin VB.Menu MNU_ADD_TASK 
         Caption         =   "Add Me To Task Schedular"
      End
      Begin VB.Menu Mnu_StopHour 
         Caption         =   "Stop for 12 Hours Since Last Picture Image"
      End
      Begin VB.Menu Mnu_StopHourAbort 
         Caption         =   "Abort Stop Since Last Image"
      End
      Begin VB.Menu MNU_SHOW_LOGG 
         Caption         =   "SHOW FORM EVENTS LOGG"
      End
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VIEW 
      Caption         =   "EXPLORER"
   End
   Begin VB.Menu MNU_IRFAN 
      Caption         =   "IRFAN"
   End
   Begin VB.Menu MNU_TAKE_ANOTHER 
      Caption         =   "ANOTHER"
   End
   Begin VB.Menu MNU_RAPID 
      Caption         =   "RAPID FIRE"
   End
   Begin VB.Menu MNU_DELETE 
      Caption         =   "KILL"
   End
   Begin VB.Menu MNU_LAST_IMAGE_TIME 
      Caption         =   "LAST IMAGE TIME"
   End
   Begin VB.Menu MNU_STAY 
      Caption         =   "STAY"
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT ="
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'BEEP

Dim IPATH

Dim NOT_DO_MORE

Dim MNU_STAY_VAR


Dim OLD_LIST1_COUNT
Dim REQUEST_TO_UNLOAD

Dim Lab1Var
Dim START_TIMER
Dim START_TIMER_FORM

Dim DONT_RUN_TIMER3_TWICE
Dim DONT_RUN_TIMER3_IF_BEGIN_NOT_SET_RUN

Dim MouseAndKey_Opperate_Ever1
Dim MouseAndKey_Opperate_Ever2
Dim MouseClicksPicture1




'MAKE DETECT IF EXE IS NOT AS UPTODATE AS VB PROJECT
Dim OWS
Dim XMP5, YMP5

Dim XMP7, YMP7
Dim Picture1MouseClicks

Dim FR1

Dim CUR_WINDOW

Dim VARCENTER
Dim vFile ', DONT_DO_THAT_AGAIN
Dim Control, Result
Dim TTY, FS, SubPath
Dim DNAME
Public SavePicturePath, XCOUNT
Private hCapWnd, Tried_To_Disconnect_Webcam

'Description = "VB version of VidCap32.exe by Ray Mercer"
'CompatibleMode = "0"
'MajorVer = 1
'MinorVer = 1
'RevisionVer = 1
'AutoIncrementVer = 1
'ServerSupportFiles = 0
'VersionComments="A VB-only "clone" of MS VidCap32.exe by Ray Mercer <raymer@shrinkwrapvb.com>"
''VersionCompanyName = "ShrinkwrapVB <http://www.shrinkwrapvb.com>"
'VersionFileDescription = "vbVidCap - Video Capture Application"
'VersionLegalCopyright = "Copyright (C) 1998-2000 by Ray Mercer"
'VersionLegalTrademarks = "vbVidCap Copyright (C) 1998-2000 by Ray Mercer"
'VersionProductName = "vbVidCap"
'CondComp = "USECALLBACKS = 1"




'Firstly id like to say that im no expert with webcams, i only know how to capture and stop it
'The major aim of this code, for me, was to make the motion detection algorithm itself
'BYE BYE Someone better came along


'FOR WEBCAM DECLARATIONS


Dim Cmd, Whitecomp, WhiteCompCount, PPL, WHITEGO
Private Declare Function GetWindowTextLength Lib "user32.DLL" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long  'MODULE 1141
Private Declare Function GetWindowText Lib "user32.DLL" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Private Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long

Dim TS As String, QQ, Txw$, ii, WhiteFactor2, Cat, WhiteFactor, Oi, TYX

Private Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Dim ET, TBak, TY, Tx, HH

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal nID As Long) As Long
Private Declare Function capGetDriverDescription Lib "avicap32.dll" Alias "capGetDriverDescriptionA" _
                                        (ByVal dwDriverIndex As Long, _
                                        ByVal lpszName As String, _
                                        ByVal cbName As Long, _
                                        ByVal lpszVer As String, _
                                        ByVal cbVer As Long) As Long 'returns C BOOL


Private mCapHwnd As Long

Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private Const GET_FRAME As Long = 1084
Private Const COPY As Long = 1054

'declarations
Dim P() As Long
Dim POn() As Boolean

Dim inten As Integer

Dim i As Integer, j As Integer

Dim Ri As Long, Wo As Long
Dim RealRi As Long

Dim C As Long, c2 As Long

Dim R As Integer, G As Integer, B As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer

Dim Tppx As Single, Tppy As Single
Dim Tolerance As Integer

Dim RealMov As Integer

Dim Counter As Integer

Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim LastTime As Long


Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long



Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long



Private Declare Function IsIconic Lib "user32.DLL" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32.DLL" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.DLL" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long



Const SW_SHOWNORMAL = 1

' ------------------------------------------------------------------------
' ------------------------------------------------------------------------

' ------- Find the first window
' ------- MISSING DLL VERSION
'Private Declare Function FindWindow _
'        Lib "user32.DLL" _
'        Alias "FindWindowA" _
'        (ByVal lpClassName As Long, _
'        ByVal lpWindowName As Long) As Long

'Private Declare Function FindWindow _
'        Lib "user32" _
'        Alias "FindWindowA" _
'        (ByVal lpClassName As Long, _
'        ByVal lpWindowName As Long) As Long

'' ------------------------------------------------------------------------
'' ------------------------------------------------------------------------

Private Declare Function GetParent _
        Lib "user32" (ByVal hWnd As Long) As Long
        
'Public Declare Function SetParent _
'        Lib "user32" _
'        (ByVal hWndChild As Long, _
'         ByVal hWndNewParent As Long) As Long
'
Private Declare Function GetWindowThreadProcessId _
        Lib "user32" _
        (ByVal hWnd As Long, _
        lpdwProcessId As Long) As Long
        
'Private Declare Function GetWindow _
'        Lib "user32" _
'        (ByVal hwnd As Long, _
'        ByVal wCmd As Long) As Long

'Const GW_HWNDNEXT = 2

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

Private Declare Function CreateProcessA _
        Lib "kernel32" _
        (ByVal lpApplicationName As String, _
         ByVal lpCommandLine As String, _
         ByVal lpProcessAttributes As Long, _
         ByVal lpThreadAttributes As Long, _
         ByVal bInheritHandles As Long, _
         ByVal dwCreationFlags As Long, _
         ByVal lpEnvironment As Long, _
         ByVal lpCurrentDirectory As String, _
         lpStartupInfo As STARTUPINFO, _
         lpProcessInformation As PROCESS_INFORMATION) As Long

'Private Declare Function CloseHandle Lib "kernel32" _
'        (ByVal hObject As Long) As Long

Const NORMAL_PRIORITY_CLASS = &H20&

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Const INFINITE = -1&

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer

Private Type POINTAPI
        x As Long
        y As Long
End Type


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function WindowFromPoint _
                 Lib "user32" (ByVal xPoint As Long, _
                               ByVal yPoint As Long) As Long





'-------------------------------------
'DECLARATIONS
'-------------------------------------



Private Function GetWindowState(ByVal lngHwnd As Long) As Integer
    If IsWindow(lngHwnd) = 1 Then
        GetWindowState = -1
        If IsIconic(lngHwnd) <> 0 Then
            GetWindowState = vbMinimized
        ElseIf IsZoomed(lngHwnd) <> 0 Then
            GetWindowState = vbMaximized
        End If
    End If
End Function

Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function



Private Sub CmboSource_Change()

End Sub

Private Sub Form_Load()

START_TIMER = Timer

IPATH = App.Path + "\" + GetComputerName + "-" + GetUserName
If Dir(IPATH, vbDirectory) = "" Then
    MkDir App.Path + "\" + GetComputerName + "-" + GetUserName
End If

If InStr(Command$, "-REBOOT-CLEAR") > 0 Then
    If Dir(IPATH + "\JOB COMPLETE TO END RESULT--.TXT") <> "" Then
        Kill IPATH + "\JOB COMPLETE TO END RESULT--.TXT"
    End If
End If

If Dir(IPATH + "\JOB COMPLETE TO END RESULT--.TXT") <> "" Then

    '----------------------------------
    'JOB NOT COMPLETE SUCCESS LAST TIME
    'NOT DO ANYMORE UNTIL REBOOT
    '----------------------------------
    
    NOT_DO_MORE = True

End If


FR1 = FreeFile
Open IPATH + "\JOB COMPLETE TO END RESULT--.TXT" For Output As #FR1
Close #FR1


List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Form Start ----"


If App.PrevInstance = True Then
    End
End If

'Beep

CUR_WINDOW = GetWindowState(GetForegroundWindow)

If UCase(Command$) = "TASK" Then
    
    Call MNU_ADD_TASK_Click
    End
    Unload Me
    Exit Sub
End If


Dim Xxr, Xxp2, DateX, DateX2, DateX3, HaltHour
Set FS = CreateObject("Scripting.FileSystemObject")

If Dir(App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\Stop for an Hour.txt") <> "" Then
    'Last Image Taken
    FR1 = FreeFile
    Open App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\WebCam % Pic Data2.txt" For Input As #FR1
        Line Input #FR1, DateX
    Close #FR1
    DateX = Mid(DateX, 5)
    DateX2 = DateValue(DateX) + TimeValue(DateX)
    DateX2 = Now ' OVERRIDE ABOVE USE TIME FROM NOW NOT LAST PIC
    'Time to Halt for an Hour
    FR1 = FreeFile
    Open App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\Stop for an Hour.txt" For Input As #FR1
        Line Input #FR1, DateX
    Close #FR1
    DateX3 = DateValue(DateX) + TimeValue(DateX)
    If DateX3 > DateX2 Then
        HaltHour = True
        Mnu_StopHour.Checked = True
        
        Unload Me
        Exit Sub
    Else
        Mnu_StopHour.Checked = False
        Kill App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\Stop for an Hour.txt"
    End If
End If


If FindWinPart = True Then End

'If FindWindow(vbNullString, "vbVidCap") > 0 Then End

Form1.Hide
Form1.WindowState = vbNormal

Call MouseKeyboardDetectPress_Timer

If IsIDE = False Then
    'Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /runexit """ + App.Path + "\" + App.EXEName + ".vbp""", vbHide
    'End
End If

If IsIDE = False Then
    Form1.Caption = "WebCam Motion Capture"
Else
    Form1.Caption = "WebCam Motion Capture In IDE"
End If

ChDir App.Path

Tppx = Screen.TwipsPerPixelX
Tppy = Screen.TwipsPerPixelY



List1.left = 0
List1.top = 0
List1.Width = 640 * Tppx
List1.Height = 480 * Tppy
List1.Visible = True

'set up the visual stuff
Picture1.Width = 640 * Tppx
Picture1.Height = 480 * Tppy
'Picture1.ScaleWidth = 640 * 5
'Picture1.ScaleHeight = 480 * 5
'Picture1.Refresh
'DoEvents

Form1.Width = Picture1.Width + 150
VARCENTER = False
Form1.Height = Picture1.Height + (54 * Tppy) ' - ACCOUNT FOR TOP MENU BAR ENABLED

Picture1.left = 0
Picture1.top = 0

Picture1.Visible = False
List1.Visible = True

Form2.Picture2.Visible = False
Form3.Picture3.Visible = False




'Inten is the measure of how many pixels are going to be recognized. I highly dont recommend
'setting it lower than this, i have a 3.0 GHz PC and it starts to lag a little. On this setting,
'every 15th pixel is checked
inten = 15 '8 '' HIGHER = LESS CHECKS -- grid chk
'The tolerance of recognizing the pixel change
Tolerance = 35 'less is more

ReDim POn(640 / inten, 480 / inten)
ReDim P(640 / inten, 480 / inten)

'Exit Sub

Cmd = Command$
'Cmd = "Quite"
If IsIDE = True Or Cmd = "" Then
    'Me.WindowState = 0
Else
'    Form1.Hide
End If

'Me.WindowState


'Form1.Show
DoEvents


Form1.WindowState = vbMinimized
'Form1.Show
Form1.Hide


MouseKeyboardDetectPress.Enabled = True


If HaltHour = False And NOT_DO_MORE = False Then
    '--------------------------------------------------------------
    'START OF PROGRAM MAIN ROUTINE
    Timer1.Enabled = True
    '--------------------------------------------------------------
Else
    Timer1.Enabled = False
    
    ' STILL OPPERATE TIMER2 IF NOT TIMER1
    Timer2.Enabled = True
    'AND TIMER5 WHEN TIMER2
    Timer5.Enabled = False
    Timer5.Interval = 15000
    Timer5.Enabled = True

    
    'IF DONT_DO_THAT_AGAIN = False
    Form1.WindowState = vbMinimized
    Form1.Show
End If

Call MouseKeyboardDetectPress_Timer


Dim TIMER_SPEED
START_TIMER_FORM = Timer
TIMER_SPEED = Str(Timer - START_TIMER_FORM) + " Secs Form Load"

List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Form Start Finished --" + TIMER_SPEED



End Sub
Private Sub SaveWebCamSingleFrame(Filename As String)
'Dim FileName As String
Dim retVal As Boolean

'retVal = VBGetSaveFileName(FileName, _
'                            filter:="DIB Bitmap Files (*.bmp)|*.bmp", _
'                            DlgTitle:="Save Single Frame", _
'                            DefaultExt:="bmp", _
'                            Owner:=Me.hWnd)
'retVal = VBGetSaveFileName(FileName, _
'                            filter:="DIB Bitmap Files (*.jpg)|*.jpg", _
'                            DlgTitle:="Save Single Frame", _
'                            DefaultExt:="jpg", _
'                            Owner:=Me.hwnd)
'If False <> retVal Then
    'hCapWnd = capwnd
    retVal = capFileSaveDIB(hCapWnd, Filename)
    If True <> retVal Then
        'MsgBox "Problem saving frame", vbInformation, App.Title
        Filename = ""
    End If
'End If
End Sub

Private Sub Form_Resize()

'    If Me.WindowState = vbMinimized Then
'            Lab1Var = Label1.Caption
'    End If



'If OWS = vbMinimized And Me.WindowState <> vbMinimized Then
'    DONT_DO_THAT_AGAIN = True
'End If
'OWS = Me.WindowState


'If Timer2.Enabled = True And 1 = 2 Then
'
'    If OWS = Me.WindowState Then
'        DONT_DO_THAT_AGAIN = True
'    End If
'    OWS = Me.WindowState
'End If

If Me.WindowState <> vbNormal Then Exit Sub

If VARCENTER = True Then Exit Sub

Me.top = Screen.Height / 2 - Me.Height / 2 '+400
Me.left = Screen.Width / 2 - Me.Width / 2 '+400
VARCENTER = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Cancel = True
'XCOUNT = 1000
'Call Timer2_Timer
'Exit Sub

Dim XV
REQUEST_TO_UNLOAD = True
If Timer1.Enabled = True Then

    'SET ALL TIMERS IN ALL FORMS TO ENABLED = FALSE
    On Error Resume Next
        For i = 0 To Forms.Count - 1
            For Each Control In Forms(i).Controls
                If InStr(UCase(Control.Name), "TIMER") > 0 Then
                    Control.Enabled = False
                End If
            Next
        Next i
    On Error GoTo 0


    XV = Now + TimeSerial(0, 0, 18)
    Do
        DoEvents
        If Timer1.Enabled = True Then Sleep 10
        'If Timer1.Enabled = True Then Sleep 20
        'Timer1.Enabled = False
    Loop Until Timer1.Enabled = False Or XV <= Now
'    If xv = 0 Then MsgBox "Timer1 Still Engaged Problem End Anyway"
End If





'SET ALL TIMERS IN ALL FORMS TO ENABLED = FALSE
On Error Resume Next
    For i = 0 To Forms.Count - 1
        For Each Control In Forms(i).Controls
            If InStr(UCase(Control.Name), "TIMER") > 0 Then
                Control.Enabled = False
            End If
        Next
    Next i
On Error GoTo 0
   
'Call capDriverDisconnect(hCapWnd) ': hCapWnd = 0
Result = capDriverDisconnect(hCapWnd) ': hCapWnd = 0
If Result = False Then Call DestroyWindow(hCapWnd)  ': hCapWnd = 0

'disconnect VFW driver
'Call mVFW.capDriverDisconnect(hCapWnd)
'destroy CapWnd
'If hCapWnd <> 0 Then Call DestroyWindow(hCapWnd) ': hCapWnd = 0
'STILL DON'T WORK AND CLEAR THE hCapWnd
   
   
Dim Form As Form
For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next Form

Reset
Close

On Error Resume Next
If NOT_DO_MORE = False Then
    Kill IPATH + "\JOB COMPLETE TO END RESULT--.TXT"
End If

Set Form1 = Nothing


End Sub



Private Sub Label4_Click()
End Sub

Private Sub MNU_ADD_TASK_Click()
Dim OPTSHELL$

OPTSHELL$ = "C:\WINDOWS\SYSTEM32\SCHTASKS /Delete /TN ""# " + App.EXEName + """ /F"
'Shell OPTSHELL$, vbHide
shellAndWait OPTSHELL$
'Debug.Print OPTSHELL$
OPTSHELL$ = "C:\WINDOWS\SYSTEM32\SCHTASKS /Create /S """ + GetComputerName + """ /RU """ + GetComputerName + "\" + GetUserName + """ /RP "" "" /SC MINUTE /MO 5 /TN ""# " + App.EXEName + """ /TR """ + App.Path + "\" + App.EXEName + ".exe"" /ST 00:00:00"
Shell OPTSHELL$, vbHide

'Debug.Print OPTSHELL$

End Sub

Private Sub MNU_DELETE_Click()

Dim FILESPEC
'MsgBox "Irfan - Also Load Explorer Minimized - Before Exit", vbMsgBoxSetForeground
If SavePicturePath <> "" Then

    On Error Resume Next
    Kill SavePicturePath
    If Err.Number = 0 Then
        MsgBox "DELETE WITH SUCCESS IMAGE" + vbCrLf + SavePicturePath
    Else
        MsgBox "DELETE ERROR " + vbCrLf + Err.Description
    End If
Else
        MsgBox "Last Image Not a Saved One - Un-Able to Delete Last Image"
End If
Timer3.Enabled = False
Timer5.Enabled = False

End Sub

Private Sub MNU_EXIT_Click()
Unload Me
End Sub

Private Sub MNU_IRFAN_Click()
Timer3.Enabled = False
Timer5.Enabled = False

Dim FILESPEC
'MsgBox "Irfan - Also Load Explorer Minimized - Before Exit", vbMsgBoxSetForeground
If SavePicturePath <> "" Then
'    Shell "Explorer.exe /select, " + SavePicturePath, vbMinimizedNoFocus
    Shell "C:\Program Files\IrfanView\i_view32.exe """ + SavePicturePath + """ /fs /silent", vbMaximizedFocus
    'Unload Me
    Exit Sub
End If

If LAST_FILE_DATE_PATH = "" Then
    
    'Private Sub MNU_LAST_WEBCAM_PIC_Click()
    'ScanPath.SHOW
    
    'XdATE2 = 0
    'ScanPath.chk_LIST_VIEW_SHORT_5 = vbChecked
    
    LAST_FILE_DATE_TIME = DateSerial(100, 1, 1)
    
    ScanPath.cboMask = "*.JPG"
    ScanPath.chkSubFolders = vbChecked
    
    ScanPath.TxtPath = "D:\0 00 Art Loggers - WEBCAM\"
    
    ScanPath.chkCopyMemory.Value = vbChecked
    
    Call ScanPath.CMDScan_NO_LIST_FAST_Click
    
    FILESPEC = LAST_FILE_DATE_PATH 'ScanPath.lblCount7
    
    'Set F = FS.getfile((Filespec1))
    'ADATE1 = F.datelastmodified
    
    'ScanPath.lblCount7 = ""
    'ScanPath.ListView1.ListItems.Clear
    
    'ScanPath.TxtPath = "D:\0 00 Art Loggers - WEBCAM\"
    'Call ScanPath.cmdScan_Click
    
    
    'Filespec2 = ScanPath.lblCount7
    'If Filespec2 <> "" Then
    '    Set F = FS.getfile((Filespec2))
    '    ADATE2 = F.datelastmodified
    '    If ADATE1 > ADATE2 Then
    '        FileSpec = Filespec1
    '    Else
    '        FileSpec = Filespec2
    '
    '    End If
    'Else
    '    FileSpec = Filespec1
    'End If

End If

'Me.WindowState = vbMinimized
'If MNU_MESSAGE_BOXES.Checked = False Then
'    MsgBox "FOUND LATEST IMAGE Clipboarded - LOAD Explorer Minimized AS Well as IrFan Maximized To View" + vbCrLf + "FILES FOUND =" + Str(tFileCount) + vbCrLf, vbMsgBoxSetForeground
'End If
'Me.WindowState = vbMinimized

'Shell "Explorer.exe /select, " + FILESPEC, vbMinimizedNoFocus

Shell "C:\Program Files\IrfanView\i_view32.exe """ + FILESPEC + """ /fs /silent", vbMaximizedFocus

'Me.WindowState = vbMinimized


'Unload Me

End Sub

Private Sub MNU_RAPID_Click()
If TimerRAPID.Enabled = False Then
    TimerRAPID.Enabled = True
    MNU_RAPID.Caption = "**Rapid Strobe"
Else
    TimerRAPID.Enabled = False
    MNU_RAPID.Caption = "Rapid Strobe"
End If
End Sub

Private Sub MNU_SHOW_LOGG_Click()

If List1.Visible = True Then
    List1.Visible = False
    Exit Sub
End If

If List1.Visible = False Then
    List1.Visible = True
End If


End Sub

Private Sub MNU_STAY_Click()
    MNU_STAY_VAR = True
    Timer2.Enabled = False
    MNU_EXIT.Caption = "END IF WANT"
    Me.WindowState = vbMinimized
End Sub

Private Sub Mnu_StopHour_Click()
FR1 = FreeFile
Open App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\Stop for an Hour.txt" For Output As #FR1
    'Print #FR1, Format(Now + TimeSerial(1, 0, 0), "YYYY-MM-DD HH:MM:SS")
    Print #FR1, Format(Now + TimeSerial(12, 0, 0), "YYYY-MM-DD HH:MM:SS")
Close #FR1
Unload Me
End Sub

Private Sub Mnu_StopHourAbort_Click()
Kill App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\Stop for an Hour.txt"
Unload Me
End Sub

Private Sub MNU_SYSTEM_PROPERTIES_Click()
'Shell "sysTEM.cpl", vbMaximizedFocus

vFile = "sysdm.cpl"
Result = ShellExecute(Me.hWnd, vbNullString, vFile, vbNullString, vbNullString, vbNormal) 'vbMaximized 'vbNormal
'Debug.Print Result
Unload Me
End Sub

Private Sub MNU_TAKE_ANOTHER_Click()

If Timer1.Enabled = True Then Exit Sub

START_TIMER = Timer

'DONT RUN TIMER3 EVER AGAIN AFTER TIMER3 AND MORE CODE OF THIS ONE
DONT_RUN_TIMER3_TWICE = True

'SET ALL TIMERS IN ALL FORMS TO ENABLED = FALSE
'BUT NOT OUR FORM
On Error Resume Next
    For i = 0 To Forms.Count - 1
        If Forms(i).Name <> Me.Name Then
            For Each Control In Forms(i).Controls
                If InStr(UCase(Control.Name), "TIMER") > 0 Then
                    Control.Enabled = False
                End If
            Next
        End If
    Next i
On Error GoTo 0
   
Dim Form As Form
For Each Form In Forms
    If Form.Name <> "Form1" Then
        Unload Form
        Set Form = Nothing
    End If
Next Form

Picture1 = Nothing

XCOUNT = 0

Timer1.Enabled = True

End Sub

Private Sub Mnu_Task_Schedualr_Click()
'http://am-techzone.blogspot.co.uk/2011/07/windows-shortcuts.html

'Shell "Taskschd.msc", vbNormalFocus
'Shell "SCHTASKS /Create /S system /U """ + GetUserName + """ /P "" "" /SC MINUTE /MD 10 /TN """ + App.EXEName + """ /TR """ + App.Path + "\" + App.EXEName

Shell "control schedtasks", vbNormalFocus


End Sub

Private Sub MNU_VB_Click()
Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
Unload Me
End Sub

Private Sub MNU_VIEW_Click()

'Shell "Explorer.exe /select, " + SavePicturePath, vbNormalFocus
Timer3.Enabled = False
Timer5.Enabled = False

'Unload Me

Timer3.Enabled = False
Timer5.Enabled = False

Dim FILESPEC
'MsgBox "Irfan - Also Load Explorer Minimized - Before Exit", vbMsgBoxSetForeground
If SavePicturePath <> "" Then
    Shell "Explorer.exe /select, " + SavePicturePath, vbNormalFocus
'    Shell "C:\Program Files\IrfanView\i_view32.exe """ + SavePicturePath + """ /fs /silent", vbMaximizedFocus
    'Unload Me
    Exit Sub
End If

If LAST_FILE_DATE_PATH = "" Then
    
    'Private Sub MNU_LAST_WEBCAM_PIC_Click()
    'ScanPath.SHOW
    
    'XdATE2 = 0
    'ScanPath.chk_LIST_VIEW_SHORT_5 = vbChecked
    
    LAST_FILE_DATE_TIME = DateSerial(100, 1, 1)
    
    ScanPath.cboMask = "*.JPG"
    ScanPath.chkSubFolders = vbChecked
    
    ScanPath.TxtPath = "D:\0 00 Art Loggers - WEBCAM\"
    
    ScanPath.chkCopyMemory.Value = vbChecked
    
    Call ScanPath.CMDScan_NO_LIST_FAST_Click
    
    FILESPEC = LAST_FILE_DATE_PATH 'ScanPath.lblCount7
    
    'Set F = FS.getfile((Filespec1))
    'ADATE1 = F.datelastmodified
    
    'ScanPath.lblCount7 = ""
    'ScanPath.ListView1.ListItems.Clear
    
    'ScanPath.TxtPath = "D:\0 00 Art Loggers - WEBCAM\"
    'Call ScanPath.cmdScan_Click
    
    
    'Filespec2 = ScanPath.lblCount7
    'If Filespec2 <> "" Then
    '    Set F = FS.getfile((Filespec2))
    '    ADATE2 = F.datelastmodified
    '    If ADATE1 > ADATE2 Then
    '        FileSpec = Filespec1
    '    Else
    '        FileSpec = Filespec2
    '
    '    End If
    'Else
    '    FileSpec = Filespec1
    'End If

End If


Shell "Explorer.exe /select, " + FILESPEC, vbNormalFocus

'Shell "C:\Program Files\IrfanView\i_view32.exe """ + FILESPEC + """ /fs /silent", vbMaximizedFocus




End Sub


Private Sub Picture1_Click()
'Me.WindowState = vbMinimized
'DONT_DO_THAT_AGAIN = True

'DONT RUN TIMER3 EVER AGAIN AFTER TIMER5 AND MORE CODE OF THIS ONE
DONT_RUN_TIMER3_TWICE = True
Timer3.Enabled = False


End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'Picture1.top = 800

If Picture1MouseClicks = -1 Then Exit Sub

If XMP7 = 0 And YMP7 = 0 Then
    XMP7 = x: YMP7 = y
End If


    If (x <> XMP7 Or y <> YMP7) Then
        
        Picture1MouseClicks = Picture1MouseClicks + Sqr(Abs(XMP7 - x) ^ 2 + Abs(YMP7 - y) ^ 2)
                    
                    
    
        'Label1 = Str(Val(Int(MouseClicks))) + " --" + Str(Val(Int(Mousey)))
        'Label1 = Lab1Var + " --" + Str(Val(Int(Picture1MouseClicks)))
        
'        If Me.WindowState = vbMinimized Then
'            Lab1Var = Label1
'        End If
        
        
        If Picture1MouseClicks > 200000 Then
        
            Timer5.Enabled = False
            Timer3.Enabled = False
            DONT_RUN_TIMER3_TWICE = True
            Picture1MouseClicks = -1
        End If
        
        XMP5 = x: YMP5 = y
    
    End If



End Sub

'START OF THE PROGRAM MAIN ROUTINE
Private Sub Timer1_Timer()

List1.Visible = True
Picture1.Visible = False

Dim FF$
MNU_LAST_IMAGE_TIME.Caption = "Operation Begin"
Call MouseKeyboardDetectPress_Timer

List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Main Code Timer1 Start"
List1.ListIndex = List1.ListCount - 1

If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

Timer1.Interval = 1

Call MouseKeyboardDetectPress_Timer

List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Load WebCam Form Start"
List1.ListIndex = List1.ListCount - 1

If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub
Load frmMain

List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Load WebCam Form Finished"
List1.ListIndex = List1.ListCount - 1

If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

Call MouseKeyboardDetectPress_Timer

'frmMain.Show
'Timer1.Enabled = False
'Exit Sub
'Unload frmMain
'Sleep 15000
'End


'App.path  "\" + GetComputerName + "-" + GetUserName +"\VBDataNoArchive\WebCamTemp.Jpg"

List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Get WebCam Form HWND Handle"
List1.ListIndex = List1.ListCount - 1

If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

hCapWnd = frmMain.capwnd

Call MouseKeyboardDetectPress_Timer


List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Grab Frame Command Begin"
List1.ListIndex = List1.ListCount - 1

If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

MNU_LAST_IMAGE_TIME.Caption = "Grab Frame Command Begin"
Call MouseKeyboardDetectPress_Timer

Call capGrabFrame(hCapWnd)

Call MouseKeyboardDetectPress_Timer

'Call capGrabFrame(hCapWnd)
'Call capGrabFrame(hCapWnd)
'Call capGrabFrame(hCapWnd)
'Call capGrabFrame(hCapWnd)

List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Get WebCame Driver Name Command Begin"
List1.ListIndex = List1.ListCount - 1

If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

MNU_LAST_IMAGE_TIME.Caption = "Get WebCame Driver Name"
Call MouseKeyboardDetectPress_Timer

DNAME = capDriverGetName(hCapWnd)

Call MouseKeyboardDetectPress_Timer


If FS.FolderExists(App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive") = False Then
    MkDir App.Path + "\" + GetComputerName + "-" + GetUserName
    MkDir App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive"
End If

Dim FILEPATH2 As String
Dim FILEPATH3 As String

FILEPATH2 = App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\WebCamTemp.Jpg"
'Kill FILEPATH2
FILEPATH3 = FILEPATH2

'RETURNS FILEPATH2 = "" IF FALSE IMAGE

List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Save WebCam Single Frame Begin"
List1.ListIndex = List1.ListCount - 1
If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

MNU_LAST_IMAGE_TIME.Caption = "Save WebCam Single Frame"
Call MouseKeyboardDetectPress_Timer

Call SaveWebCamSingleFrame(FILEPATH2)

Call MouseKeyboardDetectPress_Timer

'FALSE CODE EXPECTED NOT A FILE MAYBE DEL FIRST
'If Dir(App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\WebCamTemp.Jpg") = "" Then

If FILEPATH2 = "" Then

    List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "WebCam Single Frame Not a Image Present"
    List1.ListIndex = List1.ListCount - 1
    If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

    MNU_LAST_IMAGE_TIME.Caption = "Single Frame Not a Image"
    Call MouseKeyboardDetectPress_Timer


'    If CUR_WINDOW = vbMaximized Then Unload Me: Exit Sub  'End

    Timer1.Enabled = False
    'OPERATE TIMER2 AFTER TIMER1
    'Timer2.Enabled = True
    
    'DONT_DO_THAT_AGAIN = False
    
    'If Timer5.Enabled = False Then
        'Me.Show
        'Me.WindowState = vbNormal
    'End If
    
    List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Get Last Saved WebCam Image"
    List1.ListIndex = List1.ListCount - 1
If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

    MNU_LAST_IMAGE_TIME.Caption = "Get Last Saved WebCam"
    Call MouseKeyboardDetectPress_Timer
    
    
    List1.Visible = False
    Picture1.Visible = True
    On Error Resume Next
    Set Picture1 = LoadPicture(FILEPATH3)
    On Error GoTo 0
    Dim f
    Set f = FS.getfile(FILEPATH3)
    
    Dim StepLine, StepLineStart, StepLineStartX
    
    With Picture1
        '.AutoRedraw = True
        StepLine = 20 * Screen.TwipsPerPixelY
        StepLineStart = 1 * Screen.TwipsPerPixelY ' --- Top Line
        StepLineStartX = 1 * Screen.TwipsPerPixelX ' --- Left Line
        .FontBold = False
        .FontSize = 18
        .ForeColor = vbGreen
        .CurrentX = StepLineStartX
        .CurrentY = StepLineStart
        Picture1.Print "Picture Image Not Captured"

        .CurrentX = StepLineStartX
        .CurrentY = StepLineStart + StepLine * 1
        Picture1.Print "Last Picture Image Time -- " + Format$(f.datelastmodified, "DDD DD-MMM-YYYY HH:MM:SSa/p");

        .Visible = True
    End With
        
    Me.Refresh
    
    List1.Visible = False
    
    Exit Sub
End If


'Numbers money count dracular

'Unload Form1
'Unload frmMain
'Exit Sub


'STARTCAM

List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Load Last Saved WebCam Image"
List1.ListIndex = List1.ListCount - 1
If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

MNU_LAST_IMAGE_TIME.Caption = "Load Last Saved WebCam"
Call MouseKeyboardDetectPress_Timer

On Error Resume Next
Set Picture1 = LoadPicture(App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\ArrayPic.jpg")
On Error GoTo 0

List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "1 of 2 -- Compare Motion and Brightness WebCam Image"
List1.ListIndex = List1.ListCount - 1
If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub
MNU_LAST_IMAGE_TIME.Caption = "1 of 2 - Compare Motion"
Call MouseKeyboardDetectPress_Timer

'#1 OF 2
Call ComParePic


List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Load Back Last Saved WebCam Image into a PictureBox"
List1.ListIndex = List1.ListCount - 1
If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub
MNU_LAST_IMAGE_TIME.Caption = "Load Back Last Saved WebCam"
Call MouseKeyboardDetectPress_Timer

On Error Resume Next
'Call GET_PICTURE_SAVED_TO_FILE_FROM_CAMERA
Set Picture1 = LoadPicture(App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\WebCamTemp.Jpg")
On Error GoTo 0


Form2.Picture2.Width = Picture1.Width
Form2.Picture2.Height = Picture1.Height '- 8

Form2.Picture2.Picture = Picture1.Picture

'Call ResizePictureBoxToImage(Form2.Picture2.Picture, Tppx, Tppy) 'twipWd As Integer, twipHt As Integer)
'CALL ResizePictureBoxToImage2TEST



List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Disconnect WebCam Device"
List1.ListIndex = List1.ListCount - 1
If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

MNU_LAST_IMAGE_TIME.Caption = "Disconnect WebCam Device"
Call MouseKeyboardDetectPress_Timer

Result = capDriverDisconnect(hCapWnd) ': hCapWnd = 0
If Result = False Then
    List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Destroy Window of WebCam Device Because Disconnect Failed"
    List1.ListIndex = List1.ListCount - 1
    If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

    MNU_LAST_IMAGE_TIME.Caption = "Destroy WebCam Device"
    Call MouseKeyboardDetectPress_Timer
    
    Call DestroyWindow(hCapWnd)  ': hCapWnd = 0
End If


List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "2 of 2 -- Compare Motion and Brightness WebCam Image"
List1.ListIndex = List1.ListCount - 1
If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

MNU_LAST_IMAGE_TIME.Caption = "2 of 2 - Compare Motion"
Call MouseKeyboardDetectPress_Timer
'Why Twice? YES
'#2 OF 2
Call ComParePic

Call MouseKeyboardDetectPress_Timer




'disconnect VFW driver
'Call mVFW.capDriverDisconnect(hCapWnd)
'destroy CapWnd
'If hCapWnd <> 0 Then Call DestroyWindow(hCapWnd) ': hCapWnd = 0
'STILL DON'T WORK AND CLEAR THE hCapWnd

'Load Form1
Load Form2
Unload Form3

Call MouseKeyboardDetectPress_Timer


List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Sizing Picture Images"
List1.ListIndex = List1.ListCount - 1
If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

MNU_LAST_IMAGE_TIME.Caption = "Sizing Picture Images"
Call MouseKeyboardDetectPress_Timer

'Debug.Print "------------"
For i = 0 To Forms.Count - 1
    If InStr(Forms(i).Name, "Form1") > 0 And Len(Forms(i).Name) = 5 Then
        If IsIDE = True Or Cmd = "" Then
            Forms(i).Show
        Else
            Forms(i).Hide
        End If
    End If
    
    If Forms(i).Name = "Form2" Or Forms(i).Name = "Form3" Then
        Forms(i).Width = Form1.Width
        Forms(i).Height = Form1.Height - (19 * Screen.TwipsPerPixelY) ' - ACCOUNT FOR TOP MENU BAR NOT ENABLED
    End If
    
    For Each Control In Forms(i).Controls
        If InStr(Control.Name, "Picture") > 0 Then
            With Control
                .Visible = True
                .Height = Picture1.Height
                .Width = Picture1.Width
                .left = Picture1.left
                .top = Picture1.top
            End With
        End If
    Next
Next



Label1.Caption = Int(Wo / (Ri + Wo) * 100) & " % movement" & vbCrLf & "Real Movement: " & RealRi & vbCrLf _
& "Completed in: " & GetTickCount - LastTime


For i = 0 To Forms.Count - 1
    
    Call MouseKeyboardDetectPress_Timer
    
    For Each Control In Forms(i).Controls
        If Control.Name = "Picture1" Or Control.Name = "Picture2" Then
             
            'State all Statistics %
            With Control
                '.AutoRedraw = True
                StepLine = 14 * Screen.TwipsPerPixelY
                StepLineStart = 435 * Screen.TwipsPerPixelY ' --- Top Line
                StepLineStartX = 470 * Screen.TwipsPerPixelX ' --- Left Line
                .FontBold = False
                .FontSize = 10
                .ForeColor = vbGreen
                .CurrentX = StepLineStartX
                .CurrentY = StepLineStart
                'Control.Print "640x480"
                Tx = Wo / (Ri + Wo) * 100
                If Tx = 100 Then Tx = 99
                If Tx < 10 Then HH = " "
                .CurrentX = StepLineStartX
                .CurrentY = StepLineStart + StepLine * 1
                Control.Print "Motion " + HH;
                .ForeColor = vbWhite
                Control.Print Format$(Wo / (Ri + Wo) * 100, "0") & "%";
                .ForeColor = vbGreen
                Control.Print " White ";
                .ForeColor = vbWhite
                Control.Print Str(Int(WhiteFactor))
                .CurrentX = StepLineStartX
                .CurrentY = StepLineStart + StepLine * 2
                .ForeColor = vbGreen
                Control.Print Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p")
                '.AutoRedraw = False

            End With
        End If
    Next
Next i

'why you have to have -- Picture1.Print -- in With Statement
'http://computer-programming-forum.com/16-visual-basic/f37ebaab7a8e292f.htm

'ANSWER ----------------
'MATT.LAN@BTINTERNET.COM
'BUT YOU CAN HAVE ABOVE

'--------------------------------------------------------------------------



'Call Copy_Blend_Picture_Box





'For i = 0 To Forms.Count - 1
'    For Each Control In Forms(i).Controls
'        If Control.Name = "Picture3" Then
'            With Control
'                .AutoRedraw = Picture1.AutoRedraw
'                .ScaleHeight = Picture1.ScaleHeight
'                .ScaleWidth = Picture1.ScaleWidth
'                .Height = Form3.Height
'                .left = Picture1.left
'                .top = Picture1.top
'                .Width = Form3.Width
'                .Visible = True
'            End With
'        End If
'    Next
'Next



'With Form3.Picture3
'    .Height = Picture1.Height '480 * Tppy
'    .left = Picture1.left
'    .top = Picture1.top
'    .Width = Picture1.Width '640 * Tppx
'    .Visible = True
'End With


Dim D

'SubPath = "M:\0 00 Art"

'Set FS = CreateObject("Scripting.FileSystemObject")


'If FS.DriveExists("M") = False Then
'    SubPath = "D:\0 00 Art Loggers - WEBCAM\"
'Else
    
'    Set D = FS.GetDrive("D:\")
'    If D.ISREADY = False Then
'        SubPath = "D:\0 00 Art Loggers - WEBCAM\"
'    Else
'        SubPath = "M:\0 00 Art Loggers - WEBCAM\"
'    End If
'End If
    
    
List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "File System Work"
List1.ListIndex = List1.ListCount - 1
If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

MNU_LAST_IMAGE_TIME.Caption = "File System Work"
Call MouseKeyboardDetectPress_Timer
    
SubPath = "D:\0 00 Art Loggers - WEBCAM"
If FS.FolderExists(SubPath) = False Then MkDir SubPath
SubPath = "D:\0 00 Art Loggers - WEBCAM\WEBCAM"
If FS.FolderExists(SubPath) = False Then MkDir SubPath

'If FS.FOLDEREXISTS(SubPath + "\01 Loggers") = False Then MkDir SubPath + "\01 Loggers"
'If FS.FOLDEREXISTS(SubPath + "\01 Loggers\Web_Cam") = False Then MkDir SubPath + "\01 Loggers\Web_Cam"


FF$ = "WebCapture_" + Format$(Now, "YYYY-MM-DD-DDD")
If FS.FolderExists(SubPath + "\" + FF$) = False Then MkDir SubPath + "\" + FF$


On Error Resume Next
TTY = 0
Do
    TTY = TTY + 1
    FileInUseDelay App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\Web_Cam Pic Always.jpg"
    Err.Clear
    SavePicture Picture1, App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\Web_Cam Pic Always.jpg"
Loop Until Err.Number = 0 Or TTY > 20

If TTY > 20 Then
    MsgBox "ERROR SAVE FILE 20 TIMES TRY" + vbCrLf + "SavePicture Picture1, App.path + \VBDataNoArchive\Web_Cam Pic Always.jpg"
    Unload Me
    Exit Sub
End If

FileInUseDelay App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\WebCam % Pic Data.txt"
Open App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\WebCam % Pic Data.txt" For Append As #1
    Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p") + " - " + Format$(Wo / (Ri + Wo) * 100, "0") & "%"
Close #1

Call MouseKeyboardDetectPress_Timer

FileInUseDelay App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\WebCam % Pic Data2.txt"
Open App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\WebCam % Pic Data2.txt" For Output As #1
    Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS")
Close #1

Call MouseKeyboardDetectPress_Timer

'WHATS THAT DIRTY THING FOR
'--------------------------
'Whitecomp = WhiteFactor
'WhiteCompCount = 0
'If Whitecomp < 35 Then
'    Open App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\WhiteFactorCount.txt" For Input As #1
'        Line Input #1, PPL
'    Close #1
'    Open App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\WhiteFactorCount.txt" For Output As #1
'        Print #1, Str(Val(PPL) + 1)
'    Close #1
'    WhiteCompCount = Val(PPL) + 1
'End If
'
'If Now - Int(Now) < TimeSerial(9, 0, 0) Then
'    Open App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\WhiteFactorCount.txt" For Output As #1
'        Print #1, "1"
'    Close #1
'End If
'
'WHITEGO = True
'
'If Now - Int(Now) < TimeSerial(9, 0, 0) And Whitecomp < 35 Then WHITEGO = False

'If WhiteCompCount > 3 Then WHITEGO = False

'PERCENT COMPARE MOTION
'Wo / (Ri + Wo) * 100 >= 2
'+
'WhiteFactor
Dim WHITEFACTOR_VAR, MOTION_VAR
Dim PICTURESAVED

WHITEFACTOR_VAR = 40
MOTION_VAR = 1.58 '1.92 '3.4
If Wo / (Ri + Wo) * 100 >= MOTION_VAR And WhiteFactor >= WHITEFACTOR_VAR Then
    
    FileInUseDelay App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\ArrayPic.jpg"
    SavePicture Picture1, App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\ArrayPic.jpg"
    
    Call MouseKeyboardDetectPress_Timer
    
    Dim TNow
    TNow = Now
    MNU_LAST_IMAGE_TIME.Caption = "IM-Saved " + Format(TNow, "DDD HH:MM:SS")
    Call MouseKeyboardDetectPress_Timer
    
    TS = SubPath + "\" + FF$ + "\WebCapture- " + Format$(TNow, "YYYY-MM-DD HH-MM-SS DDD ") + "- " + Format$(Wo / (Ri + Wo) * 100, "0") & "% Compare -- " + Format$(WhiteFactor, "0.0") & "%-WhiteFactor"
    TS = TS + " -- " + DNAME + ".jpg"
    FileInUseDelay TS
    SavePicture Form2.Picture2.Image, TS
'    SavePicture Form1.Picture1.Image, TS
    SavePicturePath = TS
    PICTURESAVED = True
    
    List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Picture Image Saved Within Boundrys of Brightness and Motion"
    List1.ListIndex = List1.ListCount - 1
    If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

'    MNU_LAST_IMAGE_TIME.Caption = "Picture Image Saved Motion Okay"

    
Else
    
    List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "Picture Image Not Saved Not Within Boundrys of Brightness and Motion"
    List1.ListIndex = List1.ListCount - 1
    If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

'    MNU_LAST_IMAGE_TIME.Caption = "Picture Image Not Saved Motion Not Okay"

    MNU_LAST_IMAGE_TIME.Caption = "Not Saved"
    Call MouseKeyboardDetectPress_Timer
    
    Set f = FS.getfile(App.Path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\ArrayPic.jpg")

    With Picture1
        '.AutoRedraw = True
        StepLine = 20 * Screen.TwipsPerPixelY
        StepLineStart = 1 * Screen.TwipsPerPixelY ' --- Top Line
        StepLineStartX = 1 * Screen.TwipsPerPixelX ' --- Left Line
        .FontBold = True
        .FontSize = 18
        .ForeColor = vbWhite
        .CurrentX = StepLineStartX
        .CurrentY = StepLineStart
        Picture1.Print "Picture Image Not Saved"
    
        .CurrentX = StepLineStartX
        .CurrentY = StepLineStart + StepLine * 1
        Picture1.Print "Last Picture Image Saved -- " + Format$(f.datelastmodified, "DDD DD-MMM-YYYY HH:MM:SSa/p")
        .CurrentX = StepLineStartX
        .CurrentY = StepLineStart + StepLine * 2
        If WhiteFactor < WHITEFACTOR_VAR Then
            Picture1.Print "White Factor = " + Format(WhiteFactor, "0.00") + " <" + Str(WHITEFACTOR_VAR)
        Else
            Picture1.Print "Motion = " + Format((Wo / (Ri + Wo) * 100), "0.00") + " <" + Str(MOTION_VAR)
        End If
    
        .Visible = True
    End With
    
    List1.Visible = False

    Me.Refresh
End If



'If SavePicturePath = "" Or CUR_WINDOW = vbMaximized Then
'If CUR_WINDOW = vbMaximized Then
'    Me.Show
'    Me.WindowState = vbMinimized 'SavePicturePath = SubPath + "\WEBCAM\" + FF$
'
'    'WHY
'    'DONT_DO_THAT_AGAIN = True
'End If

'TASK IS USED WHEN WANT TO ADD TO SCHEDUAL TASK
'If (IsIDE = True Or (Cmd = "" And UCase(Cmd) <> "TASK") Or SavePicturePath <> "") _
 And CUR_WINDOW <> vbMaximized Then

Call MouseKeyboardDetectPress_Timer

'IdeTestAsEXE = True
'Dim TZ
'TZ = 50
'Do

Call MouseKeyboardDetectPress_Timer
'    TZ = TZ - 1
'Loop Until TZ = 0



Dim Test
Test = False

'-----------------------------------
If IsIDE = True And Test = True Then
    Me.WindowState = vbNormal
End If


If ((Cmd = "" And UCase(Cmd) <> "TASK")) And PICTURESAVED = True Then
    'Me.Show
    Me.WindowState = vbNormal

End If

If PICTURESAVED = True Then

    'Timer5.Enabled = False
    'Timer5.Interval = 5000
    'Timer5.Enabled = True
    
End If








If Form1.Visible = False Then Unload Me

If IsIDE = IsIDE And MNU_LAST_IMAGE_TIME.Caption = "Not Saved" Then
    
    'SHOW PICTURE BREIF AS POSSIABLE IF NOT SAVED
    'BUT NOT AFTER MOUSE KEY INTERVENTION ON FORM
    If MouseAndKey_Opperate_Ever2 = False Then
        Me.WindowState = vbMinimized
    End If

End If


'--------------------------------
MouseClicksPicture1 = 0
Timer1.Enabled = False
'TIMER2 AFTER TIMER1
Timer2.Enabled = True
'AND TIMER5 WHEN TIMER2

Timer5.Enabled = False
Timer5.Interval = 10000
Timer5.Enabled = True

'OPERATE TIMER3 IF WANTED
'BUT NOT WHEN TAKE ANOTHER SHOT FROM MENU
If DONT_RUN_TIMER3_TWICE = False Then
    If MouseAndKey_Opperate_Ever1 = True Then
        
        Timer3.Enabled = False
        Timer3.Interval = 3000
        Timer3.Enabled = True
        'LEFT A LOT OF JUNK CODE IN HERE OVER THAT IN REM
        'CAREFUL OF IN TIMER2
        'SET ALL TIMERS IN ALL FORMS TO ENABLED = FALSE
        'WHY
        
    End If
End If


Dim TIMER_SPEED, TIMER_SPEED2
TIMER_SPEED = Str(Timer - START_TIMER) + " Secs Total Execution Operation"
If START_TIMER_FORM > 0 Then
    TIMER_SPEED2 = Str(Timer - START_TIMER_FORM) + " Secs"
End If
List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "End of Timer1 Main Code --" + TIMER_SPEED
List1.ListIndex = List1.ListCount - 1
If REQUEST_TO_UNLOAD = True Then Timer1.Enabled = False: Unload Me: Exit Sub

If START_TIMER_FORM > 0 Then
    List1.AddItem Format(Now, "DDD DD-MMM-YYYY HH:MM:SS ") + "End of Timer1 Main Code Minus Time in Form Load --" + TIMER_SPEED2
    List1.ListIndex = List1.ListCount - 1
End If
START_TIMER_FORM = -1

List1.AddItem "-------------------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Visible = False

Exit Sub

Errcode2:

DoEvents
Resume Next

End Sub

Sub ComParePic()
Dim TX2
WhiteFactor2 = 0
Ri = 0 'right
Wo = 0 'wrong

'LastTime = GetTickCount
Cat = 0
For i = 0 To 640 / inten - 1
    
    Call MouseKeyboardDetectPress_Timer

    For j = 0 To 480 / inten - 1
    'get a point
    C = Picture1.Point(i * inten * Tppx, j * inten * Tppy)
    'analyze it, Red, Green, Blue
        R = C Mod 256
        G = (C \ 256) Mod 256
        B = (C \ 256 \ 256) Mod 256
        
        WhiteFactor = R + G + B
        WhiteFactor2 = WhiteFactor2 + WhiteFactor
        TX2 = TX2 + 1
    
    'recall what the point was one step before this
    c2 = P(i, j)
        'analyze it
        R2 = c2 Mod 256
        G2 = (c2 \ 256) Mod 256
        B2 = (c2 \ 256 \ 256) Mod 256
        
    'main comparison part... if each R, G and B are somewhat same, then it pixel is same still
    'in a perfect camera and software tolerance should theoretically be 1 but this isnt true...
    If Abs(R - R2) < Tolerance And Abs(G - G2) < Tolerance And Abs(B - B2) < Tolerance Then
    'pixel remained same
    Ri = Ri + 1
    'Pon stores a boolean if the pixel changed or didnt, to be used to detect REAL movement
    POn(i, j) = True
    
    Else
    'Pixel changed
    Wo = Wo + 1
    'make a red dor
    P(i, j) = Picture1.Point(i * inten * Tppx, j * inten * Tppy)
    Picture1.PSet (i * inten * Tppx, j * inten * Tppy), vbRed
    POn(i, j) = False
    End If
    
    Next j
    
Next i

RealRi = 0

For i = 1 To 640 / inten - 2
    For j = 1 To 480 / inten - 2
    If POn(i, j) = False Then
        'Real movement is simply occuring when all 4 pixels around one pixel changed
        'Simply put, If this pixel is changed and all around it changed too, then this is a real
        'movement
        If POn(i, j + 1) = False Then
            If POn(i, j - 1) = False Then
                If POn(i + 1, j) = False Then
                    If POn(i - 1, j) = False Then
                    RealRi = RealRi + 1
                    Picture1.PSet (i * inten * Tppx, j * inten * Tppy), vbGreen
                    End If
                End If
            End If
        End If
        
    End If
        
        
    Next j
Next i

WhiteFactor2 = WhiteFactor2 / TX2

WhiteFactor = WhiteFactor2

'BlendPicture Picture2, Picture1

End Sub

Sub ResizePictureBoxToImage(pic As PictureBox, twipWd As Integer, twipHt As Integer)
 Exit Sub
 ' This code assumes that all units are in twips.  If
 ' not, you must convert it to twips before calling
 ' this routine.  This also assumes that the image
 ' was blt'ed to 0,0.
 Dim BorderHt As Integer, BorderWd As Integer
 BorderWd = pic.Width - pic.ScaleWidth
 BorderHt = pic.Height - pic.ScaleHeight
 pic.Move pic.left, pic.top, twipWd + BorderWd, _
   twipHt + BorderHt
End Sub

Sub Copy_Blend_Picture_Box()
Exit Sub
'ALL THIS WORKS TO COPY A PICTURE
'REMEMBER SET THE SIZE BEFORE LOAD A PICTURE
'I WAS WONDERING WHY I COULDN'T SAVE THE IMAGE WITH TEXT DRAWN OVER
'BUT WORKED THAT OUT IN END
'SO ONLY NEEDED 2 PICTURE BOX RATHER THAN 3
'BECAUSE 1ST HAD LOT TOO MANY GRAPHICS DRAWN OVER

'Form3.Picture3.AutoRedraw = True

'Form3.Picture3.PaintPicture Form2.Picture2.Picture, 0, 0, , , , , _
    Form3.Picture3.ScaleWidth, Form3.Picture3.ScaleHeight
    With Form3.Picture3
        .Width = Form2.Picture2.Width
        .Height = Form2.Picture2.Height
        .Refresh
        .AutoRedraw = True
        .PaintPicture Form2.Picture2.Image, 0, 0 ', .ScaleWidth, .ScaleHeight
        .AutoRedraw = False
    End With
'Form3.Picture3.AutoRedraw = False

Exit Sub


'Dim TX2
'WhiteFactor2 = 0
'Ri = 0 'right
'Wo = 0 'wrong

'LastTime = GetTickCount
'Cat = 0
For i = 0 To 640  'Form1.Picture1.Width / Tppx '640
    For j = 0 To 480  'Form1.Picture1.Height / Tppy '480
    'DoEvents
    'get a point
    C = Form2.Picture2.Point(i * Tppx, j * Tppy)
    'c2 = Form3.Picture3.Point(i * Tppx, j * Tppy)
    
'    .ForeColor = vbGreen
'    .CurrentX = 445 * Screen.TwipsPerPixelX
'    '.CurrentX = 545 * Screen.TwipsPerPixelX
'    .CurrentY = 445 * Screen.TwipsPerPixelY
    
    Form3.Picture3.PSet (i * Tppx, j * Tppy), C
    
    'analyze it, Red, Green, Blue
'        R = C Mod 256
'        G = (C \ 256) Mod 256
'        B = (C \ 256 \ 256) Mod 256
        
'        WhiteFactor = R + G + B
'        WhiteFactor2 = WhiteFactor2 + WhiteFactor
'        TX2 = TX2 + 1
    
    'recall what the point was one step before this
'    c2 = P(i, j)
'        'analyze it
'        R2 = c2 Mod 256
'        G2 = (c2 \ 256) Mod 256
'        B2 = (c2 \ 256 \ 256) Mod 256
'
    'main comparison part... if each R, G and B are somewhat same, then it pixel is same still
    'in a perfect camera and software tolerance should theoretically be 1 but this isnt true...
'    If Abs(R - R2) < Tolerance And Abs(G - G2) < Tolerance And Abs(B - B2) < Tolerance Then
    'pixel remained same
'    Ri = Ri + 1
    'Pon stores a boolean if the pixel changed or didnt, to be used to detect REAL movement
'    POn(i, j) = True
    
    'Else
    'Pixel changed
'    Wo = Wo + 1
    'make a red dor
'    P(i, j) = Picture1.Point(i * inten * Tppx, j * inten * Tppy)
'    Picture1.PSet (i * inten * Tppx, j * inten * Tppy), vbRed
'    POn(i, j) = False
    'End If
    
    Next j
    
Next i

'RealRi = 0
'
'For i = 1 To 640 / inten - 2
'    For j = 1 To 480 / inten - 2
'    If POn(i, j) = False Then
'        'Real movement is simply occuring when all 4 pixels around one pixel changed
'        'Simply put, If this pixel is changed and all around it changed too, then this is a real
'        'movement
'        If POn(i, j + 1) = False Then
'            If POn(i, j - 1) = False Then
'                If POn(i + 1, j) = False Then
'                    If POn(i - 1, j) = False Then
'                    RealRi = RealRi + 1
'                    Picture1.PSet (i * inten * Tppx, j * inten * Tppy), vbGreen
'                    End If
'                End If
'            End If
'        End If
'
'    End If
'
'
'    Next j
'Next i

'WhiteFactor2 = WhiteFactor2 / TX2


'WhiteFactor = WhiteFactor2


'BlendPicture Picture2, Picture1



End Sub





'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************






Function FileInUse(ByVal strFilePath As String) As Boolean
  
  Dim hFile As Long
  Dim FileInfo  As OFSTRUCT
  
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function
  If Right(strFilePath, 1) <> Chr(0) Then strFilePath = strFilePath & Chr(0)
  
  FileInfo.cBytes = Len(FileInfo)
  hFile = OpenFile(strFilePath, FileInfo, OF_SHARE_EXCLUSIVE)
  If hFile = -1 And Err.LastDllError = 32 Then
    FileInUse = True
  Else
    CloseHandle hFile
  End If
  
End Function

Sub FileInUseDelay(Txw$)
        
    QQ = Now + TimeSerial(0, 1, 30)
    Do
'        DoEvents
        ii = FileInUse(Txw$)
        If ii = True Then Sleep (200)
    Loop Until ii = False Or QQ < Now
    
    If ii = True Then
        'MsgBox "Trouble File Stuck In Use " + vbCrLf + Txw$ + vbCrLf + "Longer than 1 min 30 secs"
        Stop
    End If
End Sub




Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim l As Long
   Dim s As String
   l = GetWindowTextLength(hWnd)
   s = Space(l + 1)
   GetWindowText hWnd, s, l + 1
   GetWindowTitle = left$(s, l)
End Function



Function FindWinPart() As Long
If IsIDE = True Then Exit Function
Dim ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String
Dim Huge
'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one
test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
Do While test_hwnd <> 0
    
        ash$ = GetWindowTitle(test_hwnd)
        If ash$ <> "" And InStr(ash$, "WebCamMatts - Micro") > 0 Then
            Huge = test_hwnd
            Exit Do
        End If
        
'retrieve the next window
test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop

FindWinPart = False
If Huge > 0 Then FindWinPart = True

End Function



Private Sub Timer2_Timer()

Dim Control 'Option Explict Set
Dim SET_TIME
Dim DROPMINIMIZED

SET_TIME = 240

XCOUNT = XCOUNT + 1
'480 = 4 MINUTES

'IF DONT_DO_THAT_AGAIN = True THEN WONT EVER MINIMIZE

'If Me.WindowState = vbMinimized And XCOUNT <= 5 Then DONT_DO_THAT_AGAIN = True

'DROPMINIMIZED = 8

'If XCOUNT = DROPMINIMIZED And DONT_DO_THAT_AGAIN = False Then Me.WindowState = vbMinimized

'LOOK TIMEOUT TO MIIMIZE
'---
'If DONT_DO_THAT_AGAIN = False Then
'    If CUR_WINDOW = vbMaximized And XCOUNT = DROPMINIMIZED Then  'And DONT_DO_THAT_AGAIN = False Then
'        DONT_DO_THAT_AGAIN = True
'        Me.WindowState = vbMinimized
'    End If
'End If

If XCOUNT > SET_TIME Then XCOUNT = SET_TIME
MNU_EXIT.Caption = "END ME" + Str(SET_TIME - XCOUNT) + " SEC'S" '" = 4 MIN COUNTDOWN -- -:¦:-•:*'''''*:•-:¦:-"

If SavePicturePath <> "" Then
'TRY THIS BUT TEXT MERGE'S WHEN OVERWRITE AND SET BACKGROUND THEN ALL TEXT ON PICTURE DISAPPEAR'S
'    With Picture1
'    '.BackColor = vbBlack
'    .ForeColor = vbGreen
'    .CurrentX = 1 * Screen.TwipsPerPixelX
'    .CurrentY = 1 * Screen.TwipsPerPixelY
'    Picture1.Print "#########################";
'    .ForeColor = vbWhite
'    .CurrentX = 1 * Screen.TwipsPerPixelX
'    .CurrentY = 1 * Screen.TwipsPerPixelY
'    Picture1.Print "EXIT =" + Str(480 - XCOUNT);
'    End With


If hCapWnd > 0 And Tried_To_Disconnect_Webcam = False Then
    Result = capDriverDisconnect(hCapWnd) ': hCapWnd = 0
    If Result = False Then Call DestroyWindow(hCapWnd) ': hCapWnd = 0

    Dim XZAG
    Tried_To_Disconnect_Webcam = True
    
    'SET ALL TIMERS IN ALL FORMS TO ENABLED = FALSE
    'BUT NOT OUR FORM
    On Error Resume Next
        For i = 0 To Forms.Count - 1
            If Forms(i).Name <> Me.Name Then
                For Each Control In Forms(i).Controls
                    If InStr(UCase(Control.Name), "TIMER") > 0 Then
                        'BUT KEEP OUR'S GOING
                 '       XZAG = 0
                 '       If InStr(UCase(Control.Name), "TIMER2") = 0 Then XZAG = 1
                 '       If InStr(UCase(Control.Name), "TIMER3") = 0 Then XZAG = 1
                 '       If InStr(UCase(Control.Name), "TIMER5") = 0 Then XZAG = 1
                 '       If XZAG = 0 Then
                            Control.Enabled = False
                 '       End If
                    End If
                Next
            End If
        Next i
    On Error GoTo 0
    
End If
End If

'SavePicturePath = "" = HASN'T FINISHED SUB ROUTINE YET
'If SavePicturePath = "" Or XCOUNT < SET_TIME Then Exit Sub

If XCOUNT < SET_TIME Then Exit Sub

Unload Me

End Sub


Sub GET_PICTURE_FROM_CAMERA()
'Get the picture from camera.. the main part
'SendMessage mCapHwnd, GET_FRAME, 0, 0


'On Error Resume Next
'TYX = Clipboard.GetFormat(vbCFText)
'If TYX = True Then TBak = Clipboard.GetText
'Clipboard.Clear

'SendMessage mCapHwnd, COPY, 0, 0
'Exit Sub

'GJax = Now + TimeSerial(0, 0, 5)
'Do
'    DoEvents
'    TY = Clipboard.GetFormat(vbCFBitmap)
'    If Clipboard.GetFormat(vbCFBitmap) = True Then Exit Do
'Loop Until TY = True Or GJax < Now
'If TY = False Then Form1.Show: DoEvents: STOPCAM: 'MsgBox "Web_Cam Not On"


'If TY = True Then
    'Picture1.Picture = Clipboard.GetData
'    STOPCAM
'End If

'Picture1.Picture = frmMain.Picture

'If TY = False Or Err.Number > 0 Then
'If TYX = True Then
'    Clipboard.Clear
'    Clipboard.SetText (TBak)
'End If
'Unload Form1: Exit Sub
'End If
    
'If TBak <> "" Then
'    Clipboard.Clear
'    Clipboard.SetText (TBak)
'End If

End Sub


Sub ResizePictureBoxToImage2TEST()
'Dim bmp As Bitmap
'bmp = New Bitmap(App.path + "\" + GetComputerName + "-" + GetUserName + "\VBDataNoArchive\WebCamTemp.Jpg")
'
'If bmp.Width < picBox.Image.Width Then picBox.Width = bmp.Width: If bmp.Height < picBox.Image.Height Then picBox.Height = bmp.Height
'picBox.Invalidate() : picBox.Refresh()

'Form2.Picture2.Picture.SetBounds(x,y,width,height)

'On Error GoTo 0
End Sub



Public Function ExecCmd(cmdLine$) As Long
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim Ret&

     ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)

    ' Start the shelled application:
    Ret& = CreateProcessA(vbNullString, cmdLine$, _
                          0&, 0&, 1&, _
                          NORMAL_PRIORITY_CLASS, _
                          0&, vbNullString, _
                          start, proc)
    
    
    
    
    ' --- let it start - this seems important
    'Call WaitForSingleObject(proc.hProcess, 500)
    
    Call WaitForSingleObject(proc.hProcess, INFINITE)
    
    If Ret Then
       ExecCmd = InstanceToWnd(proc.dwProcessId)
       'Me.Print Ret, ExecCmd
       
       Call CloseHandle(proc.hThread)
       Call CloseHandle(proc.hProcess)
    End If
    
    
    
End Function

Public Function ExecCmdWAIT(cmdLine$) As Long
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim Ret&

     ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)

    ' Start the shelled application:
    Ret& = CreateProcessA(vbNullString, cmdLine$, _
                          0&, 0&, 1&, _
                          NORMAL_PRIORITY_CLASS, _
                          0&, vbNullString, _
                          start, proc)
    
    
    
    
    ' --- let it start - this seems important
'    Call WaitForSingleObject(proc.hProcess, 500)
    
    Call WaitForSingleObject(proc.hProcess, INFINITE)
    
    If Ret Then
       ExecCmdWAIT = InstanceToWnd(proc.dwProcessId)
       'Me.Print Ret, ExecCmd
       
       Call CloseHandle(proc.hThread)
       Call CloseHandle(proc.hProcess)
    End If
    
    
    
End Function

'OR THIS

Public Function shellAndWait(ByVal Filename As String) As Long
    Dim executionStatus As Long
    Dim ProcessHandle As Long
    Dim ReturnValue As Long
    'Execute the application/file
    'executionStatus = Shell(fileName, vbNormalFocus)
    executionStatus = Shell(Filename, vbHide)

    'Get the Process Handle
    ProcessHandle = OpenProcess(&H100000, True, executionStatus)

    'Wait till the application is finished
    ReturnValue = WaitForSingleObject(ProcessHandle, INFINITE)

    'Send the Return Value Back
    shellAndWait = ReturnValue

   'THERE IS NO PROPER EXIT CODE
   'JUST IF LAUNCH OKAY

End Function


Function InstanceToWnd(ByVal target_pid As Long) As Long
    Dim test_hwnd As Long, _
        test_pid As Long, _
        test_thread_id As Long
    'Find the first window
    test_hwnd = FindWindow(ByVal 0&, ByVal 0&)
    Do While test_hwnd <> 0
       'Check if the window isn't a child
       If GetParent(test_hwnd) = 0 Then
         'Get the window's thread
         test_thread_id = GetWindowThreadProcessId(test_hwnd, test_pid)
         If test_pid = target_pid Then
            InstanceToWnd = test_hwnd
            Exit Do
         End If
       End If
       'retrieve the next window
       test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    Loop
End Function


Sub MouseAndKey_Opperate(LOGICK)

    'TIMER3 1 SEC OPPERATE TO MINIMIZE QUICKER AT START IF MOUSE START MOVING
    'TIMER3 1 SEC OPPERATE TO MINIMIZE AT START QUICKER IF MOUSE START MOVING
    'QUICK MINIMIZE TIMER WHATEVER IF MOUSE KEY BACKGROUND WORK GOING ON
    
    If LOGICK = True Then
                    
        'IF WHILE IN FORM
        If Me.hWnd = GetForegroundWindow Then
            MouseAndKey_Opperate_Ever2 = True
        End If
        MouseAndKey_Opperate_Ever1 = True
        
    
    End If

    
    
    'TIMER5 5 SEC OPPERATE TO MINIMIZE AT START LESS QUICKER IF MOUSE NOT START MOVING
    'SEE FORM LOAD AND SEE PICTURESAVED Var
    'WHEN TIMER5 ENDS IT WILL SORT OPERATION OUT
    
    'Timer5.Enabled = False
    'Timer5.Interval = 5000
    'Timer5.Enabled = True

    'DON'T OPPERATE TIMER 5 UNLESS TIMER1 FINISHED
    'If Timer1.Enabled = True Then Exit Sub






End Sub



Private Sub MouseKeyboardDetectPress_Timer()

DoEvents


'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer

'Private Type POINTAPI
'        X As Long
'        Y As Long
'End Type

'MouseAndKey_Opperate = False

Dim KeyVar, i

If MouseDetectMove = True Then
    MouseAndKey_Opperate (True)
    Exit Sub
End If

i = 0
For i = 1 To 255
    KeyVar = GetAsyncKeyState(i)
    If KeyVar < -300 Then
        MouseAndKey_Opperate (True)
        Exit For
    End If
Next

End Sub


Private Sub Timer5_Timer()
    

'TIMER3 - Minimize also at start LESS QUICK without a key
'TIMER5 - 5 SECS IF MOVE AND ALL WHILE WITH A KEY
'TIMER5 - OBVIOUS IF KEY HAS MAYBE BEEN POPPED UP
'TIMER5 - HOPE DONT DO THAT AGAIN WILL TELL IF KEYPRESSED

'If DONT_DO_THAT_AGAIN = False Then
                          

''DON'T OPPERATE TIMER 5 UNLESS TIMER3 FINISHED
''IF WE POPPED OUR WINDOW UP DON'T LET TIMER3 DOWN IT AFTER
'If Timer1.Enabled = True Then Exit Sub

'IF TIMER3 DONE AND KEYPRESS=TRUE ABORT USE TIMER5

'Debug.Print Timer3.Enabled
'Debug.Print MouseAndKey_Opperate_Ever1
'Debug.Print MouseAndKey_Opperate_Ever2


'WHY
If Timer3.Enabled = False And MouseAndKey_Opperate_Ever2 = True Then Exit Sub
' WHILE FORM SHOWING
If MouseAndKey_Opperate_Ever2 = True Then Exit Sub


If Me.WindowState <> vbMinimized Then

    Me.WindowState = vbMinimized

End If

Timer5.Enabled = False

'DONT RUN TIMER3 EVER AGAIN AFTER TIMER3 AND MORE CODE OF THIS ONE
DONT_RUN_TIMER3_TWICE = True

End Sub

Public Sub FindCursor(x, y)

Dim P As POINTAPI

GetCursorPos P
'   return x and y co-ordinate
x = P.x ' / GetSystemMetrics(0) * Screen.Width
'   for current cursor position
y = P.y '/ GetSystemMetrics(1) * Screen.Height

End Sub

Private Function MouseDetectMove()

MouseDetectMove = False
Dim NX, NY, MouseClicks
FindCursor NX, NY

'This Will Happen to Mouse If User Is Logged off
'If Nx = 0 And Ny = 0 Then
'LockSSaver = Now + LockSaverDelay
'Winamp Video
'Call SetLockMax
'Call ProgressLock
'End If

'If Nx = 0 Or Ny = 0 Then Exit Sub

'If Me.WindowState = vbMinimized Then
'    MouseClicks = 0
'    Mousey = 0
'    Mnu_Exit.Caption = "Exit"
'    Exit Sub
'End If

If XMP5 = 0 And YMP5 = 0 Then
    XMP5 = NX: YMP5 = NY
End If


    If (NX <> XMP5 Or NY <> YMP5) Then 'And (Nx <> ScreenWidth And Ny <> ScreenHeight) Then
        'Call WinonPoint
        'Mousey = Mousey + Sqr(Abs(Xmp5 - NX) ^ 2 + Abs(Ymp5 - NY) ^ 2)
        MouseClicks = Sqr(Abs(XMP5 - NX) ^ 2 + Abs(YMP5 - NY) ^ 2)
                    
        'WITHIN THE FORM
        If TimerRAPID.Enabled = True Then
            
            If Me.hWnd = GetWindowParentHWND(WindowFromPoint(NX, NY)) Then
                
                MouseClicksPicture1 = MouseClicksPicture1 + Sqr(Abs(XMP5 - NX) ^ 2 + Abs(YMP5 - NY) ^ 2)
            End If
        End If
                    
'        If MouseFirstRun = 0 Then
'            MouseFirstRun = 1
'            Mousey = 0
'        End If
    
    'Form2.Label1 = Str(Val(Int(MouseClicks))) + " --" + Str(Val(Int(Mousey)))
        
        If MouseClicks > 3 Then
        'If Mousey > 3 Or MouseClicks > 3 Then

'           Label21.Caption = Str$(Mousey)
            'MNU_EXIT.Caption = "Exit *"
'            If DONT_DO_THAT_AGAIN = False Then
'                'Me.WindowState = vbMinimized
'                DONT_DO_THAT_AGAIN = True
'                Timer3.Enabled = True
'            End If
'            Timer5.Enabled = False
'            Timer5.Interval = 5000
'            Timer5.Enabled = True
            
            MouseDetectMove = True
        End If
        
        XMP5 = NX: YMP5 = NY
    
    End If

End Function

'Function GetWindowTitle(ByVal hWnd As Long) As String
'   Dim l As Long
'   Dim S As String
'   l = GetWindowTextLength(hWnd)
'   S = Space(l + 1)
'   GetWindowText hWnd, S, l + 1
'   GetWindowTitle = left$(S, l)
'End Function

Public Function GetWindowParentHWND(ByVal ParentVar As Long) As Long
   Dim i As Long
   Dim j As Long
   i = ParentVar
   'If ParentVar Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
      i = j
   'End If
   GetWindowParentHWND = i
End Function

Sub Timer3_Timer()

    If DONT_RUN_TIMER3_TWICE = True Then Exit Sub
    
    If Me.WindowState <> vbMinimized Then 'And Me.hWnd <> GetForegroundWindow Then

        Me.WindowState = vbMinimized
        Timer3.Enabled = False
    End If
        
    DONT_RUN_TIMER3_TWICE = True

End Sub


Private Sub TimerRAPID_Timer()

    If Timer1.Enabled = True Then Exit Sub
        
    If MouseClicksPicture1 = 0 Then Exit Sub
    
    'Label1 = MouseClicksPicture1
    
    'Picture1.top = 800
    
    MouseClicksPicture1 = 0
    
    Timer3.Enabled = False
    Timer5.Enabled = False
    Timer2.Enabled = False
    XCOUNT = 0
    MNU_EXIT.Caption = "END ME"
    
    Call MNU_TAKE_ANOTHER_Click

End Sub
