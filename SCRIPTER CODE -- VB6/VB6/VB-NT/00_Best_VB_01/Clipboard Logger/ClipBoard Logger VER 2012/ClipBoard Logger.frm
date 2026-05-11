VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "ClipBoard Logger"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   14430
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ClipBoard Logger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   14430
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   8490
      Top             =   2310
   End
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   8430
      Top             =   1590
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   45
      Left            =   3495
      ScaleHeight     =   45
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   6990
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   4260
      ScaleHeight     =   120
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   7005
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Picture2 
      Height          =   195
      Left            =   6000
      ScaleHeight     =   135
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   7035
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Height          =   285
      Left            =   5295
      ScaleHeight     =   225
      ScaleWidth      =   360
      TabIndex        =   6
      Top             =   6930
      Visible         =   0   'False
      Width           =   420
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   5775
      TabIndex        =   4
      Top             =   7320
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Test Clip"
      Height          =   390
      Left            =   4410
      TabIndex        =   3
      Top             =   7290
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Text"
      Height          =   390
      Left            =   3120
      TabIndex        =   2
      Top             =   7290
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear ClipBoard"
      Height          =   390
      Left            =   1380
      TabIndex        =   1
      Top             =   7290
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   8460
      Top             =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clip all"
      Height          =   390
      Left            =   45
      TabIndex        =   0
      Top             =   7290
      Visible         =   0   'False
      Width           =   1320
   End
   Begin MCI.MMControl MMControl2 
      Height          =   330
      Left            =   10005
      TabIndex        =   10
      Top             =   7290
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   7710
      Left            =   -60
      TabIndex        =   5
      Top             =   -15
      Width           =   14445
      _ExtentX        =   25479
      _ExtentY        =   13600
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"ClipBoard Logger.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu Mnu_Exit 
      Caption         =   "Exit"
   End
   Begin VB.Menu Mnu_Clips 
      Caption         =   "Clips"
      Begin VB.Menu Mnu_ClipAll 
         Caption         =   "Clip All"
      End
      Begin VB.Menu Mnu_ClearClipBoard 
         Caption         =   "Clear ClipBoard"
      End
      Begin VB.Menu Mnu_Clear_Text 
         Caption         =   "Clear Text"
      End
      Begin VB.Menu Mnu_Test_Clip 
         Caption         =   "Test Clip"
      End
   End
   Begin VB.Menu Mnu_LoggExplorer 
      Caption         =   "Open Logg Explorer"
   End
   Begin VB.Menu Mnu_Loggs 
      Caption         =   "Loggs Menu"
      Begin VB.Menu Mnu_Open_Recent 
         Caption         =   "Open Recent Trim Logg"
      End
      Begin VB.Menu Mnu_Open_Recent_Hex 
         Caption         =   "Hex Open Recent Trim Logg"
      End
      Begin VB.Menu Mnu_Open_Logg 
         Caption         =   "Open This Logg"
      End
      Begin VB.Menu Mnu_Open_Total 
         Caption         =   "Open Total Logg"
      End
      Begin VB.Menu Mnu_StripLogg 
         Caption         =   "Open Strip Logg"
      End
      Begin VB.Menu Mnu_ShellT_Todays 
         Caption         =   "Shell T -- Todays Logg"
      End
      Begin VB.Menu Mnu_Shell_T 
         Caption         =   "Shell T Total Logg"
      End
   End
   Begin VB.Menu Mnu_Minimize 
      Caption         =   "Minimize"
   End
   Begin VB.Menu Mnu_Max 
      Caption         =   "Max"
   End
   Begin VB.Menu Mnu_Norm 
      Caption         =   "Norm"
   End
   Begin VB.Menu Mnu_VB6 
      Caption         =   "VB ME"
   End
   Begin VB.Menu Mnu_Options 
      Caption         =   "OPTIONS"
      Begin VB.Menu Mnu_SoundOn 
         Caption         =   "SOUND ON"
      End
      Begin VB.Menu Mnu_Edit_Sound 
         Caption         =   "EDIT SOUND"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DO A LOAD WITH HEX VIEW FOR CODES
'HERE
'C:\Program Files\XVI32\XVI32.exe

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long


Dim OFH, a1 As String, B1 As String, F1 As String

Private Declare Function FlatSB_SetScrollPos Lib "comctl32" (ByVal hWnd As Long, ByVal code As Long, ByVal nPos As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "comctl32" (ByVal hWnd As Long, ByVal fnBar As Long, lpsi As SCROLLINFO) As Boolean
Private Declare Function FlatSB_SetScrollInfo Lib "comctl32" (ByVal hWnd As Long, ByVal fnBar As Long, lpsi As SCROLLINFO, ByVal fRedraw As Boolean) As Long

Const SB_HORZ = 0
Const SB_VERT = 1
Const SB_BOTH = 3
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Const SIF_RANGE = &H1
Const SIF_PAGE = &H2
Const SIF_POS = &H4
Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS)


Public RRR, XXMOUSEMOVE, OLDTTS, Picture3W, Picture3H, Easy2, Star, Rr$, RrS$, Tx1$, CountR2$
Public AD$, TRemGG$
Public CountR
Public Start
Public Pic1$, Pic2$, OT1
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetParent _
        Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long  'MODULE 1141
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2
Const GW_CHILD = 5

Private Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lClose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Dim Filexxx$, CurProcHwnd, TTT
Dim Rect1 As RECT

Dim TS1 As String, TS2 As String, TS3 As String, QQ, Txw$, ii, FF$, XX, YY
'Private Declare Function GetForegroundWindow Lib "user32" () As Long

'Public cProcesses As New clsCnProc
    '#### Functions/Consts used for GetHWnd() (part of Convert())
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'Private Const GW_HWNDNEXT = 2
'Private Const GW_CHILD = 5
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long



Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

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

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Dim TJax, GJax, ET, TBak, TY, Tx, HH

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal nID As Long) As Long

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
Dim LastTime As Long, TT




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



Private Sub Command1_Click()

Clipboard.Clear
Clipboard.SetText Text1.Text

End Sub

Private Sub Command2_Click()
Clipboard.Clear

End Sub

Private Sub Command3_Click()
    
    Text1.Text = ""
    CountR = 0

End Sub

Private Sub Command4_Click()
Clipboard.Clear
Clipboard.SetText Format$((Now), "DDD DD-MMM-YY HH:MM:SS")
End Sub

Private Sub Form_Activate()
'MsgBox "hello"

End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End

'Picture2.Picture = LoadPicture(App.Path + "\VBIcon4.bmp")
    
Set FS = CreateObject("Scripting.FileSystemObject")
    
Dim FileSpec, TT As Long
FileSpec = App.Path + "\" + App.EXEName + ".vbp"

If IsIDE = False And Dir$(FileSpec) <> "" Then
    'TT = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /runexit """ + FileSpec + """", vbMinimizedNoFocus)
    'End
End If
    
    
Open App.Path + "\VBIcon4.bmp" For Binary As #1
Pic1$ = Space$(LOF(1))
Get #1, 1, Pic1$
Close #1


On Error Resume Next
MkDir App.Path + "\Data\" + GetComputerName
MkDir App.Path + "\Data\" + GetComputerName + "\Day-Data"
On Error GoTo 0


Open App.Path + "\Data\" + GetComputerName + "\00_ClipBoard_Total--TRIM.txt" For Binary As #1
If LOF(1) > 1 * 1024 ^ 2 Then
    Pic1$ = Space$((1 * 1024 ^ 2) + 1)
    Seek 1, LOF(1) - (1 * 1024 ^ 2)
    Get #1, , Pic1$
    Close #1

'---------------------
'Count =

ii = InStr(Pic1$, "---------------------" + vbLf + "Count =")
Pic1$ = Mid$(Pic1$, ii)

Kill App.Path + "\Data\" + GetComputerName + "\00_ClipBoard_Total--TRIM.txt"
Open App.Path + "\Data\" + GetComputerName + "\00_ClipBoard_Total--TRIM.txt" For Binary As #1

Put #1, , Pic1$
Close #1

'Simple Copy File
Set FS = CreateObject("Scripting.FileSystemObject")
a1 = App.Path + "\Data\" + GetComputerName + "\00_ClipBoard_Total--TRIM.txt"
B1 = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\00_ClipBoard_Total--TRIM.txt"
'fs.CopyFile ,
ShellFileCopy a1, B1



End If
Close #1

Pic1$ = ""



Tx1$ = App.Path + "\Data\" + GetComputerName + "\#OutClipChunck.Txt"
DumVar = IsFileOpenDelay(Tx1$)
fr1 = FreeFile
Open Tx1$ For Binary As #fr1
AD$ = Input$(LOF(fr1), fr1)
Close #fr1

Text1 = ""

If Dir(App.Path + "\Data\" + GetComputerName + "\Start.txt") <> "" Then
    Open App.Path + "\Data\" + GetComputerName + "\Start.txt" For Input As #1
    Line Input #1, Star
    Close #1
End If

Start = Int(Now)
If Star = "" Then Star = Now
If DateValue(Star) = Start Then
    fr1 = FreeFile
    Open App.Path + "\Data\" + GetComputerName + "\#ClipBoard.Txt" For Binary As #fr1
    RrS$ = Input$(LOF(fr1), fr1)
    Close #fr1
    Err.Clear
    On Error Resume Next
    Text1.Text = Text1.Text + RrS$
    If Err.Number > 0 Then
        MsgBox "Not Adding to CLip Board"
    End If
    On Error GoTo 0
    Tx1$ = App.Path + "\Data\" + GetComputerName + "\#Count.Txt"
    fr1 = FreeFile
    If Dir(Tx1) <> "" Then
        Open Tx1$ For Input As #fr1
        Line Input #fr1, CountR2$
        Close #fr1
    End If
    CountR = Val(CountR2$)

Else
    tq = App.Path + "\Data\" + GetComputerName + "\#ClipBoard_Old.Txt"
    If Dir(tq) <> "" Then
    Kill tq
    Name App.Path + "\Data\" + GetComputerName + "\#ClipBoard.Txt" As App.Path + "\Data\" + GetComputerName + "\#ClipBoard_Old.Txt"
'    Open App.Path + "\Data\" + GetComputerName + "\#ClipBoard.Txt" For Output As #fr1
'    Close #fr1
    End If
    CountR = 0
End If

Open App.Path + "\Data\" + GetComputerName + "\Start.txt" For Output As #1
Print #1, Format$(Start, "DD-MM-YYYY")
Close #1


'after load most import recent load a backlogg for viewing
fr1 = FreeFile
Open App.Path + "\Data\" + GetComputerName + "\#ClipBoard_Old.Txt" For Binary As #fr1
RrS$ = Input$(LOF(fr1), fr1)
Close #fr1
Err.Clear
On Error Resume Next
Text1.Text = RrS$ + Text1.Text
If Err.Number > 0 Then
    MsgBox "Not Adding to CLip Board"
End If





Form1.Show
    
'Text1.SelStart = 0
'Text1.SelLength = Len(Text1)
'Text1.Font.Size = 12
'Text1.SelColor = &HFF00&
    
    
    
' Set properties needed by MCI to open.
'MMControl1.Notify = True
'MMControl1.Wait = True
'MMControl1.Shareable = False
'MMControl1.DeviceType = "WaveAudio"
'MMControl1.FileName = "D:\VB6\VB-NT\00_Best_VB_01\ClipBoard Logger\01 Woarble Tone 5 Mins.wav"
'' Open the MCI WaveAudio device.
'MMControl1.Command = "Open"
   
'MMControl1.Command = "prev"
'MMControl1.Command = "Play"

MMControl2.Notify = True
MMControl2.Wait = True
MMControl2.Shareable = False
MMControl2.DeviceType = "WaveAudio"
'MMControl2.FileName = App.Path + "\Camera1a_2_8kHz.wav"
'MMControl2.FileName = App.Path + "\Shot-12 Mix to Shot-18.wav"


MMControl2.FileName = App.Path + "\01 Woarble Tone 5 Mins.wav"
MMControl2.Command = "Open"

'D:\Wave's\Camera1a_2_8kHz.wav

'MMControl2.Command = "prev"
'MMControl2.Command = "Play"



Call zzLoad_Checks


DoEvents

'Form1.Width = Text1.Width + 80
'Form1.Height = Text1.Height + Text1.Top + 680

DoEvents
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2


'Call Mnu_Norm_Click



Form1.WindowState = 1
Text1.Font.Size = 14

Text1.SelStart = Len(Text1.Text)


Call Mnu_Norm_Click



Load FrmJoy



End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then HHH = Now + TimeSerial(0, 10, 0) 'MsgBox Str(Form1.Top)
Call Mnu_Norm_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)

'D:\MY WEBS\

'Unload FrmJoy
'Unload frmSender

Call zzSave_Checks

If Me.WindowState <> 1 Then
    Me.WindowState = 1
    'test may have to put back form need reseting
    'Unload FrmJoy
    Cancel = True
    Exit Sub
End If



End
End Sub

Private Sub Mnu_Clear_Text_Click()
Call Command3_Click
End Sub

Private Sub Mnu_ClearClipBoard_Click()
Call Command2_Click
End Sub

Private Sub Mnu_ClipAll_Click()
Call Command1_Click
End Sub

Private Sub Mnu_Edit_Sound_Click()

Dim TT As String

TT = App.Path + "\01 Woarble Tone 5 Mins 02.wav"

TT = FindShortPath(TT)

Shell """C:\Program Files\Cool2000\cool2000.exe"" """ + TT + """", vbNormalFocus

End Sub

Private Sub Mnu_Exit_Click()
MMControl2.Command = "close"

Tx1$ = App.Path + "\Data\" + GetComputerName + "\#OutClipChunck.Txt"
DumVar = IsFileOpenDelay(Tx1$)
fr1 = FreeFile
Open Tx1$ For Output As #fr1
Print #fr1, AD$;
Close #fr1

End
Unload Form1
End Sub

Private Sub Mnu_LoggExplorer_Click()
Shell "Explorer.exe /e," + App.Path + "\Data", vbNormalFocus
End Sub

Public Sub Mnu_Max_Click()
On Error Resume Next
Exit Sub
'Me.WindowState = 2
Form1.Left = 0
Form1.Top = 420
Form1.Width = Screen.Width - 3800
Form1.Height = Screen.Height - 4700

Text1.Left = 0
Text1.Top = 0
Text1.Width = Form1.Width - 90
Text1.Height = Form1.Height - 800

If Text1.Font.Size < 14 Then
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
    Text1.Font.Size = 18
    Text1.SelStart = Len(Text1)
End If

'Me.SetFocus

End Sub

Private Sub Mnu_Minimize_Click()
Me.WindowState = 1
End Sub

Private Sub Mnu_Norm_Click()
On Error Resume Next



'Me.WindowState = 2
Form1.Left = 0
Form1.Top = 420

'If GetComputerName = "MATT-555ROIDS" Then
Form1.Width = Screen.Width - 3800
Form1.Height = Screen.Height - 1500
'End If


DoEvents

Text1.Left = 0
Text1.Top = 0
Text1.Width = Form1.Width - 90
Text1.Height = Form1.Height - 800


If GetComputerName = "MATT-555ROIDS" Then
'    Text1.Height = 10000
'    Text1.Width = 18000
End If
If GetComputerName = "55-88-HAPPY" Then
'    Text1.Height = 7000
'    Text1.Width = 14000
End If



'If Text1.Font.Size > 14 Then
'    Text1.SelStart = 0
'    Text1.SelLength = Len(Text1)
'    Text1.Font.Size = 12
'    Text1.SelStart = Len(Text1)
'End If

'Me.SetFocus


End Sub

Private Sub Mnu_Open_Logg_Click()
Shell "notepad " + App.Path + "\Data\" + GetComputerName + "\Day-Data\ClipBoard-" + Format$(Start, "YYYY-MM-DD") + ".Txt", vbNormalFocus
End Sub

Private Sub Mnu_Open_Recent_Click()
Shell "notepad " + App.Path + "\Data\" + GetComputerName + "\00_ClipBoard_Total--TRIM.txt", vbNormalFocus
End Sub

Private Sub Mnu_Open_Recent_Hex_Click()
'C:\Program Files\XVI32\XVI32.exe
Shell "C:\Program Files\XVI32\XVI32.exe """ + App.Path + "\Data\" + GetComputerName + "\00_ClipBoard_Total--TRIM.txt""", vbNormalFocus
End Sub

Private Sub Mnu_Open_Total_Click()
Shell "notepad " + App.Path + "\Data\" + GetComputerName + "\00_ClipBoard_Total.txt", vbNormalFocus
End Sub

Private Sub Mnu_Shell_T_Click()
ee$ = GetShortName(App.Path + "\Data\" + GetComputerName + "\00_ClipBoard_Total.txt")
Shell "C:\Utils\T.com " + ee$, vbNormalFocus
End Sub

Private Sub Mnu_ShellTRecent_Click()

End Sub

Private Sub Mnu_ShellT_Todays_Click()
ee$ = GetShortName(App.Path + "\Data\" + GetComputerName + "\#ClipBoard.Txt")
Shell "C:\Utils\T.com " + ee$, vbNormalFocus

End Sub

Private Sub Mnu_SoundOn_Click()

If Mnu_SoundOn.Checked = True Then Mnu_SoundOn.Checked = False: Exit Sub

Mnu_SoundOn.Checked = True

End Sub

Private Sub Mnu_StripLogg_Click()
Shell "notepad " + App.Path + "\Data\" + GetComputerName + "\00_ClipBoard_Tot_Strip.txt", vbNormalFocus

End Sub

Private Sub Mnu_Test_Clip_Click()
Call Command4_Click
End Sub



Private Sub Mnu_VB6_Click()
Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\ClipBoard Logger.vbp""", vbNormalFocus
End

End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
RRR = Now + TimeSerial(0, 0, 3)








Exit Sub


Dim SI As SCROLLINFO




FlatSB_SetScrollPos Text1.hWnd, SB_VERT, 60, False
SI.cbSize = Len(SI)
SI.fMask = SIF_ALL
FlatSB_GetScrollInfo Text1.hWnd, SB_VERT, SI
SI.nPos = SI.nPos - 10
SI.nPos = 50
FlatSB_SetScrollInfo Text1.hWnd, SB_VERT, SI, True

'Text1.AutoRedraw = True


End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    RRR = Now + TimeSerial(0, 0, 3)
    HHH = Now + TimeSerial(0, 2, 0)
    
    XXMOUSEMOVE = Now + TimeSerial(0, 4, 0)

End Sub

Private Sub Timer1_Timer()

'If Me.WindowState = 1 Then
'Exit Sub
'End If

On Error GoTo Errortrap

If HHH < Now And HHH > 0 Then
    Me.WindowState = 1: HHH = 0
End If


OO = False
If OO = False Then OO = Clipboard.GetFormat(vbCFText)
If OO = False Then OO = Clipboard.GetFormat(vbCFBitmap)

'MAJOR WORKAROUND

If OO = False Then Exit Sub





If OO = False Then OO = Clipboard.GetFormat(vbCFDIB)
If OO = False Then OO = Clipboard.GetFormat(vbCFEMetafile)
If OO = False Then OO = Clipboard.GetFormat(vbCFFiles)
If OO = False Then OO = Clipboard.GetFormat(vbCFLink)
If OO = False Then OO = Clipboard.GetFormat(vbCFMetafile)
If OO = False Then OO = Clipboard.GetFormat(vbCFPalette)
If OO = False Then OO = Clipboard.GetFormat(vbCFRTF)
If OO = False Then OO = Clipboard.GetFormat(-15694) 'Pasted Objects on form

If OO = False Then
    For R = -30000 To 32000
        If OO = False Then
            OO = Clipboard.GetFormat(R)
        Else
            Exit For
        End If
    Next
End If

If OO = False Then
    If TRemGG$ <> "" Then
        Clipboard.Clear
        Clipboard.SetText TRemGG$
        Exit Sub
    End If
End If

'= DO NOT DO
If 1 = 2 Then
'If Clipboard.GetFormat(vbCFBitmap) And OT1 <> vbCFBitmap and 1=2 Then
    Picture1.Picture = Clipboard.GetData(vbCFBitmap)
    
    Xx1 = Picture1.Picture.Width
    yy1 = Picture1.Picture.Height
    'hh1 = Picture1.Picture.hPal
    hh2 = Picture1.Picture.Type
    If Xx1 = 423 And yy1 = 423 And hh1 = 0 And hh2 = 1 Then
        SavePicture Picture1.Picture, App.Path + "\VBIcon.bmp"
        Open App.Path + "\VBIcon.bmp" For Binary As #1
        Pic2$ = Space$(LOF(1))
        Get #1, 1, Pic2$
        Close #1
        If Pic1$ = Pic2$ And TRemGG$ <> "" Then
            Clipboard.Clear
            Clipboard.SetText TRemGG$
            If Mnu_SoundOn.Checked = True Then
                MMControl2.Command = "prev"
                MMControl2.Command = "Play"
            End If
            Exit Sub
        End If
    
    End If
End If

'OT1 = Clipboard.GetFormat(vbCFBitmap)

If Clipboard.GetFormat(vbCFText) = True Then
    gg$ = Clipboard.GetText
    If Easy2 = True Then Picture4.Picture = LoadPicture(): Easy2 = False
Else

If Clipboard.GetFormat(vbCFBitmap) = True Then
Picture4.Picture = Clipboard.GetData
Easy2 = True
If Picture3W = Picture4.Width And Picture3H = Picture4.Height Then Exit Sub

Call PictureClip
Exit Sub

End If

End If

'gg$ = Clipboard.GetText

If Len(gg$) > 15000 Then Exit Sub

If AD$ = gg$ Then Exit Sub

If RRR > Now Then
Clipboard.Clear
Clipboard.SetText gg$, vbCFText
AD$ = Clipboard.GetText
Exit Sub
End If

'If InStr(Text1.Text, gg$) Then Exit Sub

On Error GoTo 0

Ash$ = GetActiveWindowTitle(True)

xxrt = GetForegroundWindow
Do
xxr5 = xxrt
xxrt = GetParent(xxrt)
Loop Until xxrp = 0
xxrp = xxr5
ash2$ = ""
ash2$ = GetWindowTitle(xxrp)

CountR = CountR + 1

frmSender.URLLogging

Dim UrlWork
If URL <> "" Then
    UrlWork = UrlWork + "In Internet Explorer -- " + vbLf
    UrlWork = UrlWork + "WinTitle = " + WinTitle + vbLf
    UrlWork = UrlWork + "URL = " + URL + vbLf
End If

tq$ = vbLf
Rr$ = "---------------------" + tq$
Rr$ = Rr$ + "Count = " + Format$(CountR, "000") + " -- "
Rr$ = Rr$ + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") + tq$
Rr$ = Rr$ + "---------------------" + tq$

If UrlWork <> "" Then
    Rr$ = Rr$ + UrlWork
    Rr$ = Rr$ + "---------------------" + tq$
End If

If ash2$ <> "" And ash2$ <> Ash$ Then
Rr$ = Rr$ + "Parent Window = " + ash2$ + tq$
Rr$ = Rr$ + "---------------------" + tq$
End If
Rr$ = Rr$ + Ash$ + tq$
Rr$ = Rr$ + "---------------------" + tq$
Rr$ = Rr$ + gg$ + tq$
Rr$ = Rr$ + "---------------------" + tq$
'rr$ = rr$ + "x" + vbCrLf

TRemGG$ = gg$

OFH = Text1.SelStart

Err.Clear
On Error Resume Next
Text1.Text = Text1.Text + Rr$

If Err.Number > 0 Then
    MsgBox "Not Adding to CLip Board"
Else
    If Mnu_SoundOn.Checked = True Then
        MMControl2.Command = "prev"
        MMControl2.Command = "Play"
    End If
End If

On Error GoTo 0

If Form1.Width = 2400 Then
Text1.SelStart = Len(Text1.Text)
Else
Text1.SelStart = Len(Text1.Text)
Text1.SelStart = OFH
End If

If XXMOUSEMOVE < Now Then
    Text1.SelStart = Len(Text1.Text)
End If

On Error GoTo Errortrap2

Tx1$ = App.Path + "\Data\" + GetComputerName + "\Day-Data\ClipBoard-" + Format$(Start, "YYYY-MM-DD") + ".Txt"
DumVar = IsFileOpenDelay(Tx1$)
fr1 = FreeFile
Open Tx1$ For Append As #fr1
Print #fr1, Rr$;
Close #fr1

Tx1$ = App.Path + "\Data\" + GetComputerName + "\#ClipBoard.Txt"
DumVar = IsFileOpenDelay(Tx1$)
fr1 = FreeFile
Open Tx1$ For Append As #fr1
Print #fr1, Rr$;
Close #fr1

Tx1$ = App.Path + "\Data\" + GetComputerName + "\#Count.Txt"
DumVar = IsFileOpenDelay(Tx1$)
fr1 = FreeFile
Open Tx1$ For Output As #fr1
Print #fr1, CountR
Close #fr1

Tx1$ = App.Path + "\Data\" + GetComputerName + "\00_ClipBoard_Total.txt"
DumVar = IsFileOpenDelay(Tx1$)
fr1 = FreeFile
Open Tx1$ For Append As #fr1
Print #fr1, Rr$;
Close #fr1

Tx1$ = App.Path + "\Data\" + GetComputerName + "\00_ClipBoard_Total--TRIM.txt"
DumVar = IsFileOpenDelay(Tx1$)
fr1 = FreeFile
Open Tx1$ For Append As #fr1
Print #fr1, Rr$;
Close #fr1


Tx1$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\00_ClipBoard_Total--TRIM.txt"
DumVar = IsFileOpenDelay(Tx1$)
fr1 = FreeFile
Open Tx1$ For Append As #fr1
Print #fr1, Rr$;
Close #fr1






Tx1$ = App.Path + "\Data\" + GetComputerName + "\00_ClipBoard_Tot_Strip.txt "
DumVar = IsFileOpenDelay(Tx1$)
fr1 = FreeFile
Open Tx1$ For Append As #fr1
Print #fr1, gg$
'Print #fr1, vbCrLf
Close #fr1

AD$ = gg$

Exit Sub
Errortrap:
DoEvents
Exit Sub
Resume

'If File In Use
Errortrap2:

MsgBox "Error " + Err.Description

'Sleep 500
Stop
Resume
Exit Sub

End Sub

Public Function GetActiveWindowTitle(ByVal ReturnParent As Boolean) As String
   Dim i As Long
   Dim j As Long
   i = GetForegroundWindow
   ReturnParent = False
   If ReturnParent Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
   i = j
   End If
   GetActiveWindowTitle = GetWindowTitle(i)
End Function

Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim l As Long
   Dim S As String
   l = (hWnd)
   S = Space(GetWindowTextLength(l))
   GetWindowText hWnd, S, l + 1
   'Prob here sudden start go wrong but look i is handle so how can text lenght be so
   GetWindowTitle = S 'Left$(s, l)
End Function

Sub PictureClip()

CurProcHwnd = GetForegroundWindow
Filexxx$ = GetFileFromHwnd(CurProcHwnd)
Hwnd9 = GetWindowRect(CurProcHwnd, Rect1)


Dim Ip
On Error Resume Next

Picture3.Picture = Clipboard.GetData

u1 = Picture3.Width
u2 = Picture3.Height
t1 = ((Rect1.Right - Rect1.Left) * Screen.TwipsPerPixelX) - 60
t2 = ((Rect1.Bottom - Rect1.Top) * Screen.TwipsPerPixelY) + 60
t3 = ((Rect1.Right - Rect1.Left) * Screen.TwipsPerPixelX) + 60
t4 = ((Rect1.Bottom - Rect1.Top) * Screen.TwipsPerPixelY) + 60
t5 = ((Rect1.Right - Rect1.Left) * Screen.TwipsPerPixelX)
t6 = ((Rect1.Bottom - Rect1.Top) * Screen.TwipsPerPixelY) + 30

Picture3W = u1
Picture3H = u2
Tag = 0
If u1 = t1 And u2 = t2 Then Tag = 1
If u1 = t3 And u2 = t4 Then Tag = 1
If u1 = t5 And u2 = t6 Then Tag = 1
If u1 >= 19260 And u2 >= 12060 Then Tag = 1: tag2 = 2
If Tag = 0 Then
Picture3.Picture = LoadPicture()
Exit Sub
End If

Clipboard.Clear
Picture3W = 0
Picture3H = 0

If Mnu_SoundOn.Checked = True Then
    MMControl2.Command = "prev"
    MMControl2.Command = "Play"
End If

art$ = "D:\0 00 Art Loggers\"
qq2$ = "Hot-Key-App-Shots\": qq3$ = "App"
If tag2 = 2 Then qq2$ = "Hot-Key-Screen-Shots\": qq3$ = "Screen"
FF$ = "HotKey" + qq3$ + "Capture_" + Format$(Now, "YYYY-MM-DD-DDD")
On Error Resume Next
'Error More Like Below MkDir "D:\00 Art\" + qq2$ + FF$
MkDir art$
MkDir art$ + qq2$
Err.Clear
MkDir art$ + qq2$ + FF$
If Err.Number = 76 Then
    art$ = "C:\0 00 Art Loggers\"
    MkDir art$
    MkDir art$ + qq2$
    MkDir art$ + qq2$ + FF$
End If
On Error GoTo 0
FileInUseDelay App.Path + "\VBDataNoArchive\HotKey-" + qq3$ + "-Shot Pic.jpg"
SavePicture Picture3.Picture, App.Path + "\VBDataNoArchive\HotKey-" + qq3$ + "-Shot Pic.jpg"
FileInUseDelay App.Path + "\VBDataNoArchive\HotKey-" + qq3$ + "-Shot % Pic Data.txt"
Open App.Path + "\VBDataNoArchive\HotKey-" + qq3$ + "-Shot % Pic Data.txt" For Append As #1
Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p")
Close #1
'If Wo / (Ri + Wo) * 100 >= 1 Then
    'FileInUseDelay App.Path + "\VBDataNoArchive\HotKey-" + qq3$ + "-Shot ArrayPic.jpg"
    'SavePicture Picture3.Picture, App.Path + "\VBDataNoArchive\HotKey-" + qq3$ + "-Shot ArrayPic.jpg"
    TS1 = art$ + qq2$ + FF$ + "\HotKey" + qq3$ + "Capture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + ".jpg"
    'FileInUseDelay TS1
    SavePicture Picture3.Picture, TS1
    If tag2 = 0 Then
    TS2 = art$ + qq2$ + FF$ + "\HotKey" + qq3$ + "Capture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + ".txt"
    FileInUseDelay TS2
    Open TS2 For Output As #1
    Print #1, "---------------------"
    
    If tag2 = 0 Then Print #1, "Screen HotKey Application Shot FileName ="
    If tag2 = 2 Then Print #1, "Screen HotKey Screen Shot FileName ="
    
    Print #1, TS1
    Print #1, "---------------------"
    Print #1, "Application Exe = "; Filexxx$
    Print #1, "---------------------"
    Print #1, "Application Title = "; GetWindowTitle(CurProcHwnd)
    Print #1, "---------------------"
    Print #1, "Application Class Title = "; GetWindowClass(CurProcHwnd)
    Print #1, "---------------------"
    Print #1, "Current Handle Hwnd ="; CurProcHwnd
    Print #1, "---------------------"
    Print #1, "Resolution -----"
    Print #1, "---------------------"
    Print #1, "Top ="; Rect1.Top
    Print #1, "Bottom ="; Rect1.Bottom
    Print #1, "Left ="; Rect1.Left
    Print #1, "Right ="; Rect1.Right
    Print #1, "---------------------"
    Print #1, "Width ="; Rect1.Right - Rect1.Left
    Print #1, "Height ="; Rect1.Bottom - Rect1.Top
    Print #1, "---------------------"
    Close #1
    End If

' Clear the third picture (not necessary if not visible).
Picture3.Picture = LoadPicture()


Exit Sub
Errcode2:
DoEvents
Resume Next

End Sub




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
        MsgBox "Trouble File Stuck In Use " + vbCrLf + Txw$ + vbCrLf + "Longer than 1 min 30 secs"
        Stop
    End If
End Sub




Public Function GetFileFromHwnd(lngHwnd) As String

'MsgBox getfilefromhwnd(Me.hwnd)

Dim lngProcess&, hProcess&, bla&, C&
Dim strFile As String
Dim x

strFile = String$(256, 0)
x = GetWindowThreadProcessId(lngHwnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
x = EnumProcessModules(hProcess, bla, 4&, C)
C = GetModuleFileNameExA(hProcess, bla, strFile, Len(strFile))
GetFileFromHwnd = Left(strFile, C)

End Function


Function GetWindowClass(ByVal hWnd As Long) As String
    Dim Ret As Long, sText As String
    sText = Space(255)
    Ret = GetClassName(hWnd, sText, 255)
    sText = Left$(sText, Ret)
    GetWindowClass = sText
End Function




Function IsFileAlreadyOpen(FileName As String) As Boolean
    Dim hFile As Long
    Dim lastErr As Long
    hFile = -1
    lastErr = 0
    hFile = lOpen(FileName, &H10)
    If hFile = -1 Then
        lastErr = Err.LastDllError
    Else
        lClose (hFile)
    End If
    IsFileAlreadyOpen = (hFile = -1) And (lastErr = 32)
End Function

Function IsFileOpenDelay(Tx1$)
    QQ = Now + TimeSerial(0, 4, 0)
    Do
        DoEvents
        ii = IsFileAlreadyOpen(Tx1$)
        If ii = True Then Sleep (200)
    Loop Until ii = False Or QQ < Now
    If ii = True Then
        MsgBox "Trouble File Stuck In Use " + vbCrLf + Tx1$ + vbCrLf + "Longer than 4 Min"
        Stop
    End If
End Function

Private Sub Timer2_Timer()

Call FrmJoy.Timer1_Timerxxx
Exit Sub
End Sub






'----------------       Latest Version Of Save Chks
'#### REMEBER SWITCHES

'Sub zzLoad_Checks()
'Put This In Declarations
'Dim OCheckQ
'Very Nice Code Will Save all your Check Boxes an Menu Checks and Values you can filter
'some out -- works best as I Know

'if you use menu checks you have to set them yourself on clicks
'If Mnu_VB.Checked = True Then Mnu_VB.Checked = False: Exit Sub
'If Mnu_VB.Checked = False Then Mnu_VB.Checked = True

'3 Parts
' ---
'1 Load
'2 Save
'3 Timer ' This Chks for Changes using XOR Hash
'   Best way I know with Timer Hardly CPU Usage Unfriendly

'zzCheckTimer.Enabled = True
'Dont Have Timer Enabled on Form Load

'Call Timer to Run On Form Unload ----- call zzCheckTimer_Timer
'Then you could set timer slow like 10 Secs - I Do

'-------------------------------
Sub zzLoad_Checks()

'zzCheckTimer.Enabled = False

Dim Th()
ReDim Th(Me.Controls.Count * 3)

On Error Resume Next
i = 0

Open App.Path + "\0TextData\ChkSettings.txt" For Input As #1
Do
    Line Input #1, vv$
    Th(i) = vv$
    i = i + 1
    Line Input #1, vv$
    Th(i) = Val(vv$)
    i = i + 1
    Line Input #1, vv$
    Th(i) = Val(vv$)
    i = i + 1
Loop Until EOF(1)
Close #1
    
tit = i
For Each Control In Me.Controls
    XX = 1
    
    ppi = LCase(Control.Name)
    If InStr(ppi, LCase("Combo")) Then XX = 0
    If InStr(ppi, LCase("Check")) Then XX = 0
    If InStr(ppi, LCase("Mnu")) Then XX = 0
    If InStr(ppi, LCase("Menu")) Then XX = 0
    
    
    xxd = -1
    For R = 0 To tit
        If Control.Name = Th(R) Then
            xxd = R: Exit For
        End If
    Next
    
    If xxd > 0 And XX = 0 Then
        On Error Resume Next
        If Th(xxd + 1) <> 0 Then Control.Value = Th(xxd + 1)
        If Th(xxd + 2) <> 0 Then Control.Checked = Th(xxd + 2)
        'Dont Let People Play Around If you Want to Hard Code In Designer To Enable Chk This
        'This Lets Happen Eg If Th(xxd + 2) <> 0
        
        a1 = 0
        a1 = Control.Value
        B1 = 0
        B1 = Control.Checked
        OCheckQ = OCheckQ Xor (a1 + B1)
        On Error GoTo 0
    End If
Next
'zzCheckTimer.Enabled = True
End Sub

'-------------------------------
Sub zzSave_Checks()
Dim a1, B1 As Integer
On Error Resume Next
If FS.FolderExists = False Then MkDir App.Path + "\0TextData"
Open App.Path + "\0TextData\ChkSettings.txt" For Output As #1

For Each Control In Me.Controls
    Err.Clear
        a1 = 0
        a1 = Control.Value
        a2 = Err.Number
    Err.Clear
        B1 = 0
        B1 = Control.Checked
        B2 = Err.Number
    
    If a2 = 0 Or B2 = 0 Then
        Print #1, Control.Name
        Print #1, a1
        Print #1, B1
    End If
Next
Close #1

End Sub



'-------------------------------
Private Sub zzCheckTimer_Timer()

On Error Resume Next
For Each Control In Me.Controls
    Err.Clear
        a1 = 0
        a1 = Control.Value
    Err.Clear
        B1 = 0
        B1 = Control.Checked
        checkq = checkq Xor (a1 + B1)
Next

If checkq = OCheckQ Then Exit Sub
OCheckQ = checkq
Call zzSave_Checks

End Sub



Private Sub Timer3_Timer()
    filespec1 = App.Path + "\01 Woarble Tone 5 Mins 01.wav"
    Set F = FS.GetFile((filespec1))
    Adate1 = F.datelastmodified
    Filespec2 = App.Path + "\01 Woarble Tone 5 Mins 02.wav"
    Set F = FS.GetFile((Filespec2))
    Adate2 = F.datelastmodified
    If Adate2 > Adate1 Then
    
    ATidalX.MMControl1.Command = "Close"
    
    FS.CopyFile Filespec2, filespec1
    
    ATidalX.MMControl1.Notify = True
    ATidalX.MMControl1.Wait = True
    ATidalX.MMControl1.Shareable = False
    ATidalX.MMControl1.DeviceType = "waveaudio"
    ATidalX.MMControl1.FileName = App.Path + "\01 Woarble Tone 5 Mins 01.wav"
    ATidalX.MMControl1.Command = "Open"
            
    End If

End Sub

Public Function FindShortPath(strFileName As String) As String
    Dim lngRes As Long, strPath As String
    strPath = String$(165, 0)
    lngRes = GetShortPathName(strFileName, strPath, 164)
    FindShortPath = Left$(strPath, lngRes)
End Function

