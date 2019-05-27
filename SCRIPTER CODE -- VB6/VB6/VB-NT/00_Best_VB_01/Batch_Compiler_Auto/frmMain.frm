VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.Ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form frmMain 
   Caption         =   "MAster BATch VB6 Compiler"
   ClientHeight    =   8412
   ClientLeft      =   60
   ClientTop       =   108
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8412
   ScaleWidth      =   12000
   WindowState     =   1  'Minimized
   Begin RichTextLib.RichTextBox TxtLog 
      Height          =   285
      Left            =   1950
      TabIndex        =   23
      Top             =   165
      Width           =   1410
      _ExtentX        =   2477
      _ExtentY        =   508
      _Version        =   393217
      BackColor       =   16777215
      HideSelection   =   0   'False
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   4155
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   572
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5910
      Top             =   5715
   End
   Begin VB.Frame framCover 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   525
      Left            =   5160
      TabIndex        =   15
      Top             =   885
      Visible         =   0   'False
      Width           =   2625
      Begin VB.ListBox lstUpdates 
         Height          =   1980
         IntegralHeight  =   0   'False
         Left            =   660
         TabIndex        =   18
         Top             =   930
         Width           =   4815
      End
      Begin VB.Frame fraButtons 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   375
         Left            =   660
         TabIndex        =   16
         Top             =   3090
         Width           =   4815
         Begin VB.PictureBox cmdCompile 
            BackColor       =   &H000000FF&
            Height          =   1000
            Left            =   0
            ScaleHeight     =   948
            ScaleWidth      =   948
            TabIndex        =   13
            Top             =   0
            Width           =   1000
         End
         Begin VB.PictureBox cmdBack 
            BackColor       =   &H000000FF&
            Height          =   1000
            Left            =   0
            ScaleHeight     =   948
            ScaleWidth      =   948
            TabIndex        =   14
            Top             =   0
            Width           =   1000
         End
         Begin VB.Label lblCount 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 Project(s)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1410
            TabIndex        =   17
            Top             =   90
            Width           =   1635
         End
      End
      Begin VB.Shape Border 
         Height          =   375
         Left            =   2730
         Top             =   60
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scanning Projects ... "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   75
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.Label lblUpdate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Updates Needed for The Following Projects:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   660
         TabIndex        =   19
         Top             =   630
         Width           =   3645
      End
   End
   Begin VB.TextBox txtLog2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   4245
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   5115
      Width           =   1215
   End
   Begin VB.Frame framCompile 
      Height          =   3765
      Left            =   30
      TabIndex        =   1
      Top             =   810
      Visible         =   0   'False
      Width           =   7965
      Begin VB.Frame frmSlide 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2715
         Left            =   60
         TabIndex        =   2
         Top             =   120
         Width           =   7815
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            Height          =   2695
            Left            =   0
            ScaleHeight     =   2652
            ScaleWidth      =   720
            TabIndex        =   4
            Top             =   0
            Width           =   765
            Begin VB.PictureBox picArrow 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   1368
               Index           =   1
               Left            =   60
               Picture         =   "frmMain.frx":05FF
               ScaleHeight     =   1368
               ScaleWidth      =   324
               TabIndex        =   21
               Top             =   480
               Width           =   324
            End
            Begin VB.PictureBox picArrow 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   1710
               Index           =   0
               Left            =   150
               ScaleHeight     =   1716
               ScaleWidth      =   408
               TabIndex        =   9
               Top             =   480
               Width           =   405
            End
            Begin VB.Image Image2 
               Height          =   384
               Index           =   1
               Left            =   60
               Picture         =   "frmMain.frx":0C47
               Top             =   2220
               Width           =   384
            End
            Begin VB.Image Image1 
               Height          =   384
               Index           =   1
               Left            =   0
               Picture         =   "frmMain.frx":0DDE
               Top             =   0
               Width           =   384
            End
            Begin VB.Image Image1 
               Height          =   480
               Index           =   0
               Left            =   90
               Top             =   0
               Width           =   480
            End
            Begin VB.Image Image2 
               Height          =   480
               Index           =   0
               Left            =   150
               Top             =   2220
               Width           =   480
            End
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   285
            Left            =   1380
            TabIndex        =   3
            Top             =   1800
            Width           =   6405
            _ExtentX        =   11303
            _ExtentY        =   508
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            Caption         =   "-------"
            Height          =   195
            Left            =   1440
            TabIndex        =   12
            Top             =   1380
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Average Compile Time:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   11
            Top             =   1170
            Width           =   1950
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Current Project:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1470
            TabIndex        =   8
            Top             =   30
            Width           =   1350
         End
         Begin VB.Label lblProject 
            AutoSize        =   -1  'True
            Caption         =   "-------"
            Height          =   195
            Left            =   1470
            TabIndex        =   7
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Current EXE File:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1470
            TabIndex        =   6
            Top             =   600
            Width           =   1350
         End
         Begin VB.Label lblEXE 
            AutoSize        =   -1  'True
            Caption         =   "-------"
            Height          =   195
            Left            =   1470
            TabIndex        =   5
            Top             =   810
            Width           =   420
         End
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstProj 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   550
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Project Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "EXE Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label7 
      BackColor       =   &H008080FF&
      Caption         =   "Label7"
      Height          =   528
      Left            =   8304
      TabIndex        =   26
      Top             =   2484
      Width           =   624
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      ForeColor       =   &H80000005&
      Height          =   576
      Left            =   8292
      TabIndex        =   25
      Top             =   1632
      Width           =   696
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label5"
      Height          =   585
      Left            =   8190
      TabIndex        =   24
      Top             =   585
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Batch Profile"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Batch Profile..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Batch Profile..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save Batch Profile &As..."
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuCompile 
      Caption         =   "&Compile"
      Begin VB.Menu mnuAutoCompile 
         Caption         =   "&Auto Compile"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompileSingle 
         Caption         =   "Compile &Selected Project(s)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuCompileAll 
         Caption         =   "&Compile All"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_FOLDER_ME 
      Caption         =   "FOLDER ME"
   End
   Begin VB.Menu mnuQSave 
      Caption         =   "&Quick Save"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuLogPopup 
      Caption         =   "&Event Log"
      Begin VB.Menu mnuPrintLog 
         Caption         =   "&Print Log..."
      End
      Begin VB.Menu mnuSaveLog 
         Caption         =   "&Save Log..."
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearLog 
         Caption         =   "Clear Log"
      End
   End
   Begin VB.Menu Mnu_ACompile 
      Caption         =   "Auto-Compile"
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "Run VB"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_RUN_FAVS 
      Caption         =   "RUN FAVS AFTER"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_EXIT_AFTER 
      Caption         =   "EXIT AFTER"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'main process  Sub CompileList(vIndexes As String)


Option Explicit

Public EXIT_TRUE

Dim EXIT_ON_END, END_STATE
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Dim x As Integer, ProFullPath, ProEXEPath, ENDCODE
Dim PROGCCOMPILED, RETURNvAL

Dim Bads
Private Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lClose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Dim DaysToScan2, AccessToFile, QQ, II3
Dim Tx$, Tx2$, DumVar, LV, LV2, LV3, LV4, TV2, HO, HOX, HOX2, FS, FileSpec1, ADate1, A, XXF

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Private Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
'Private Declare Function lClose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long

Dim We, A1$, B1$, DD$, TF$, II$, ET, IQ, S, RR, II2, ExitRun, Tagg, BadRun, TN
Dim NoNo ' add dont compile if running
Dim FR1, FR2 ' FreeFile add dont compile if running
Dim Jh7
Dim mIndexes As String
Public EditResult As Integer
Dim SaveLocation As String
Dim mChanged As Boolean
Dim mCompiling As Boolean
Dim CancelLoop As Boolean

Private Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

Const OF_READ = &H0
Const OF_WRITE = &H1
Const OF_READWRITE = &H2
Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40

Enum BOOL
   FALSE_ = 0
   TRUE_ = 1
End Enum

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As BOOL


'----------------------------------------------------
'----------------------------------------------------
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


Private Sub cmdBack_Click()
  framCover.Visible = False
  LockWindow lstProj.hWnd, False
End Sub

Private Sub cmdCancel_Click()
  
  'Set Cancel Variable to cancel current compilation loop
  CancelLoop = True
  
End Sub

Private Sub cmdCompile_Click()
framCover.Visible = False
LockWindow lstProj.hWnd, False
'Pause 0.1
If mIndexes = "" Then
    BadRun = True
  
    Tx$ = App.Path + "\#_Data\" + GetComputerName + "\" + GetComputerName + "__BatchComplierTextForVBCompiler.txt"
    DumVar = IsFileOpenDelay(Tx$)
    FR1 = FreeFile
    Open Tx$ For Input As #FR1
        Line Input #FR1, AXx$
    Close #FR1
    
    Call AddAlertToLogSimple(TxtLog)
    Exit Sub
End If
CompileList Left$(mIndexes, Len(mIndexes) - 1)
End Sub


Private Sub Form_Load()

'ScanPath.Show

On Error GoTo EXIT_ERROR

If App.PrevInstance = True Then End
Dim Filespec, TT

Filespec = App.Path + "\" + App.EXEName + ".vbp"
If IsIDE = False And Dir$(Filespec) <> "" Then
    'TT = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /runexit """ + Filespec + """", vbMinimizedNoFocus)
    'End
End If

Set FS = CreateObject("Scripting.FileSystemObject")



TxtLog.Font.Size = 9
TxtLog.SelStart = 0
TxtLog.SelLength = Len(TxtLog)
TxtLog.SelColor = Label6.BackColor


' Set properties needed by MCI to open.
MMControl1.Notify = True
MMControl1.Wait = True
MMControl1.Shareable = False
MMControl1.DeviceType = "WaveAudio"

'Loads the first wave in any DIR for sound effect
ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask = "*.wav"
ScanPath.txtPath.Text = App.Path
ScanPath.chkSort = vbChecked
Call ScanPath.cmdScan_Click



A1$ = ScanPath.ListView1.ListItems.Item(1).SubItems(1)
B1$ = ScanPath.ListView1.ListItems.Item(1)
MMControl1.Filename = A1$ + B1$
ScanPath.ListView1.ListItems.Clear
ScanPath.chkSort = vbUnchecked



CodeREsize


On Error Resume Next
frmMain.WindowState = 1

'DoEvents
If Command$ <> "" Then
    frmMain.WindowState = 1
    'frmMain.Visible = False
    frmMain.Show
    
Else
    frmMain.Show
    DoEvents
End If
    
DoEvents
    
frmMain.WindowState = 1
DoEvents

On Error GoTo 0
On Error GoTo EXIT_ERROR
On Error GoTo 0


VBPath = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
If Dir(VBPath) = "" Then
    VBPath = "C:\Program Files (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
End If


DoEvents

Call OpenProfileScan("")

'Using the Projects Array, Populate the list view
Call PopulateProjects

'Exit Sub
Call mnuAutoCompile_Click

If Me.WindowState = 0 Then
    'MMControl1.Command = "Open"
    'MMControl1.Command = "Play"
    '    Call WaitWavFinish
    'MMControl1.Command = "Close"
End If

If Bads = 0 Then
    TxtLog.Text = TxtLog.Text + vbCrLf + "***** Finished ******" + vbCrLf
Else
    TxtLog.Text = TxtLog.Text + vbCrLf + "***** Finished ******   **** " + Str(Bads) + " Errors" + vbCrLf
End If
TxtLog.Refresh
DoEvents


Dim Hx, Hx2

Hx = 0
Do
Hx = InStr(Hx + 1, TxtLog.Text, "Errors")
If Hx > 0 Then
    Hx2 = InStr(Hx, TxtLog.Text, "<<")
    
    If Hx2 > 0 Then
    TxtLog.SelStart = Hx - 1
    TxtLog.SelLength = Hx2 - Hx + 2
    TxtLog.SelColor = Label6.BackColor 'QBColor(12)
    End If
    Hx = Hx2
End If

Loop Until Hx = 0

TxtLog.SelStart = 0 'Len(TxtLog.Text)

TxtLog.Refresh
DoEvents


' On Error Resume Next ' FOR NOW
'HERE NO
lstProj.SetFocus

'Search this thats where they are highlighted in selection
'lstProj.ListItems.Item(R).Selected = True


'Exit Sub


'----------------- END

Me.Caption = "MAster BATch VB6 Compiler -- WORK DONE"
If Command$ <> "" And IsIDE = False Then
    'Call MNU_RUN_FAVS_Click
    End
End If

'SET WHEN RUN FAVS ON EXIT
If EXIT_ON_END = True Then Unload Me


END_STATE = True


If BadRun = True And Command$ = "" And IsIDE = False Then
    Me.WindowState = 0
    Exit Sub
End If





Exit Sub









'-----NEW END




If BadRun = False And IsIDE = False Then
    
    If PROGCCOMPILED = True And Command$ = "" Then
        If FindWindow(vbNullString, "Tidal Information...") = 0 Then
            Shell "D:\VB6\VB-NT\00_Best_VB_01\Fav Progs\Run Fav Programs.exe GOGO", vbNormalFocus
        End If
    End If
    
    Unload Me
    
End If

If PROGCCOMPILED = True And Command$ = "" Then
    If FindWindow(vbNullString, "Tidal Information...") = 0 Then
        Shell "D:\VB6\VB-NT\00_Best_VB_01\Run Fav Progs\Run Fav Programs.exe GOGO", vbNormalFocus
    End If
End If
    
    
    
    
Unload Me
    



If Me.WindowState = 1 And BadRun = False Then Unload frmMain

Dim RF
RF = FindWindow(vbNullString, "#00_Schedule_Timer")
If RF = 0 Then
    Shell "D:\VB6\VB-NT\00_Best_VB_01\#00_Schedule_Timer\#00_Schedule_Timer-ACER.exe", vbNormalNoFocus
End If


lstProj.BackColor = Label5.BackColor




'If err.Number > 0 Then Unload frmMain


'ENDCODE = True
'ONLY IF SOMEINK COMPILED

'MsgBox "h"
'Unload frmMain


EXIT_ERROR:
MsgBox "ERROR HAPPEN DURING FORM_LOAD" + vbCrLf + vbCrLf + err.Description

End Sub


Public Sub Form_Resize()

TxtLog.Refresh
txtLog2.Refresh

Me.Refresh

If ENDCODE = True Then Exit Sub

Dim x
  'If IsIDE = False Then Exit Sub
  'Resizeing Controls and column headers
    On Error Resume Next

    Me.Top = (Screen.Height / 2 - Me.Height / 2) - 300
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Height = 7800 '9900

  lstProj.Move 0, 0, ScaleWidth, (ScaleHeight / 100) * 50
  TxtLog.Move 0, lstProj.Height, ScaleWidth, (ScaleHeight / 100) * 49.7
  framCompile.Move 0, -60, ScaleWidth, lstProj.Height
  frmSlide.Move 60, 240, framCompile.Width - 120
  'splitter1.Width = ScaleWidth
  'splitter1.Top = txtLog.Top - splitter1.Height
  
  'If framCover.Visible Then splitter1_AfterScroll
  
  For x = 1 To lstProj.ColumnHeaders.Count
    lstProj.ColumnHeaders(x).Width = (lstProj.Width - 120) / lstProj.ColumnHeaders.Count
  Next
  
  

  
'  If Height < 5145 Then
'    LockWindow Me.hWnd, True
'    Height = 5145
'    LockWindow Me.hWnd, False
'  End If
  
'frmMain.Show
'DoEvents
  
  
End Sub

Sub CodeREsize()
Exit Sub
Dim x, y, Control, MNU, PlusO
  'Code to Auto Size Form based on controls used Including Menu Bar Sketchy Style
'Make sure form set to scale Twips
x = 1
y = 1
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True And Control.Visible = True Then
        If Control.Width + Control.Left > x Then x = Control.Width + Control.Left
        If Control.Height + Control.Top > y Then y = Control.Height + Control.Top
        If InStr(Control.Name, "Mnu_") > 0 Then MNU = 1
    End If
Next
On Error GoTo 0

Me.Width = (x + 80)
If MNU = 1 Then
    PlusO = 720: PlusO = 1100 'Sometimes Different
Else
    PlusO = 450
End If

Me.Height = (y + PlusO)
Me.Refresh
DoEvents

'Me.Left = (Screen.Width - Me.Width)
'Me.Top = (Screen.Height - Me.Height)


End Sub

Private Sub Form_Unload(Cancel As Integer)

'MsgBox "Hi11"



If ExitRun = True Then
    Shell App.Path + "\ExitRun\ExitRun.exe", vbMinimizedNoFocus
End If

'MsgBox Str(Now)

End

Unload ScanPath
Reset
    
End Sub




Sub AddProject(Vfile)
  
  Dim pFile As String
  pFile = Vfile
  
  
Dim TG
Dim Readln As String
Dim Num As Integer
Dim Path32 As String
Dim ExeName32 As String
  
  'If UBound(mProjects) = 25 Then Stop
  'AutoIncrementVer=1
    
    
    
    
    'ServerSupportFiles=0
    
'VersionComments="Matt.lan@btinternet.com"
'VersionCompanyName="Magik -- The One --"
'VersionLegalCopyright="Yes Like That"
'VersionLegalTrademarks="<º)))><¸..·´XNo5X`·..¸><(((º>"

'FavorPentiumPro(tm)=-1
'BoundsCheck=-1
'FlPointCheck=-1
'FDIVCheck=-1
'UnroundedFP=-1






  Reset
  
  'Retreive Compile information from VBP file
  FR1 = FreeFile
  Open Vfile For Input As #FR1
    Do While Not EOF(FR1)
      Line Input #FR1, Readln
        'If InStr(Readln, "WebCamMatts.exe") > 0 Then Stop
      
      If Left(Readln, 9) = "ExeName32" Then
        ExeName32 = Replace(Readln, "ExeName32=" & Chr(34), "")
        ExeName32 = Left$(ExeName32, Len(ExeName32) - 1)
      ElseIf Left(Readln, 6) = "Path32" Then
        Path32 = Replace(Readln, "Path32=" & Chr(34), "")
        Path32 = Left$(Path32, Len(Path32) - 1)
      End If
    Loop
  Close #FR1
  
  'Project hasnt been compiled yet?
  If ExeName32 = "" Then
    'MsgBox "Please compile this project manually in VB First."
    Exit Sub
  End If
        
        'If InStr(ExeName32, "WebCamMatts.exe") > 0 Then Stop
'        If InStr(ExeName32, "Linker2") > 0 Then Stop
        'D:\VB6\VB-NT\00_Best_VB_01\Linker2
  
  'Get Realtive Path to the project file
  Path32 = GetRelativePath(Path32, Left$(pFile, Len(pFile) - (Len(GetFileName(pFile))) - 1))
    
  Dim EXEName As String
  If InStr(ExeName32, ":\") = 0 Then
      EXEName = Path32 & "\" & ExeName32
      EXEName = Replace(EXEName, "\\", "\")
    Else
    EXEName = ExeName32
    EXEName = Replace(EXEName, "\\", "\")
  End If
  
  
On Error GoTo 0
  'Check to see if the project already exists in the list
  'If UBound(mProjects()) > 0 Then
  '  Dim XI As Integer
  ' For XI = 0 To UBound(mProjects)
  '    If LCase$(mProjects(XI).ProjectFullPath) = LCase(pFile) Then
  '      Exit Sub
  '    End If
  '  Next
  'End If
  

mProjects(UBound(mProjects)).ProjectFullPath = pFile
mProjects(UBound(mProjects)).ProjectName = GetFileName(mProjects(UBound(mProjects)).ProjectFullPath)
  
mProjects(UBound(mProjects)).ExeFullPath = EXEName
mProjects(UBound(mProjects)).EXEName = GetFileName(mProjects(UBound(mProjects)).ExeFullPath)


ReDim Preserve mProjects(UBound(mProjects) + 1)
   
   
  
End Sub

Sub UpdateProjectInfo(Vfile)
    
    
    Dim LLO, AK3, Num2, R, FF, Num, TG
    FR1 = FreeFile
    Open Vfile For Input As #FR1
        TG = Input(LOF(FR1), FR1)
    Close #FR1
    Dim Ak1(10)
    Dim Ak2(10)
    AK3 = False
    Ak1(1) = "AutoIncrementVer=1"
    Ak1(2) = "FavorPentiumPro(tm)=-1"
    Ak1(3) = "BoundsCheck=1"
    Ak1(4) = "FlPointCheck=-1"
    Ak1(5) = "FDIVCheck=-1"
    Ak1(6) = "UnroundedFP=0"
    Ak1(7) = "VersionComments=""Matt.lan@btinternet.com -- Last Compile = " + Format(Now, "DDD DD-MMM-YYYY HH-MM-SSa/p") + """"
    Ak1(8) = "VersionCompanyName=""Magik -- The One -- Hammer HardCore Coder"""
    Ak1(9) = "VersionLegalCopyright=""Yes -- Loving It -- Loving It -- Loving It"""
    Ak1(10) = "VersionLegalTrademarks=""<º)))><¸..·´XNo5X`·..¸><(((º>"""
    
    FR1 = FreeFile
    Open Vfile + ".Bak" For Output Lock Write As #FR1
    FR2 = FreeFile
    Open Vfile For Input As #FR2
    Do
        Line Input #FR2, LLO
        For R = 1 To 10
            FF = UCase(Mid$(Ak1(R), 1, InStr(Ak1(R), "=")))
            If InStr(UCase(LLO), FF) > 0 And Ak1(R) <> LLO Then
                LLO = Ak1(R)
                Ak2(R) = True
                AK3 = True
            End If
        Next
        Print #FR1, LLO
    Loop Until EOF(FR1)
    Close #FR2, #FR1
    
    If AK3 = True Then FS.CopyFile Vfile + ".Bak", Vfile
    Kill Vfile + ".Bak"
    
    If AK3 = True Then
        FR2 = FreeFile
        Open Vfile + ".Bak" For Output Lock Write As #FR2
        FR1 = FreeFile
        Open Vfile For Input As #FR1
        Do
            Line Input #FR1, LLO
            If Trim(LLO) = "" And AK3 = True Then
                AK3 = False
                For R = 1 To 10
                    If Ak2(R) = False Then
                        Print #FR2, Ak1(R)
                    End If
                Next
            LLO = ""
            End If
            Print #FR2, LLO
        Loop Until EOF(FR1)
    Close #FR2, #FR1
  
    FS.CopyFile Vfile + ".Bak", Vfile
    Kill Vfile + ".Bak"
    End If

End Sub

Function GetRelativePath(findPath As String, startPath As String) As String
  
  
    If findPath = "" Then GetRelativePath = startPath
    If findPath <> "" Then GetRelativePath = findPath
    If InStr(findPath, "..\") = 0 Then Exit Function
    GetRelativePath = startPath
    Exit Function
    
    If InStr(findPath, "..\") Then Stop
  
  
  
  
  'Exit Function
  
  Dim L As Integer
  Dim I As Integer
  Dim Backs As Integer
  Dim vstartPath As String
  
    GetRelativePath = startPath
    If findPath = startPath Then Exit Function
  
  vstartPath = startPath
  
  'Find out how many BackDirs (..\) there are
  L = Len(findPath)
  findPath = Replace(findPath, "..\", "")
  Backs = (L - Len(findPath)) / 3
  
  'Back up BACKS BackDirs
  For I = 1 To Backs
    If I = 1 Then
      L = InStrRev(vstartPath, "\")
    Else
      L = InStrRev(vstartPath, "\", L - 1)
    End If
    vstartPath = Left(vstartPath, L - 1)
  Next
  
  GetRelativePath = vstartPath & "\" & Trim(Replace(findPath, Chr(34), ""))
  
End Function

Sub Dostatus(vProjName As String, vEXEName As String, vPercent As Single)
  
  'If not compiling, this function shouldnt be called
  If Not (framCompile.Visible) Then Exit Sub
  
  'Display current compile information
  lblProject = vProjName
  lblEXE = vEXEName
  ProgressBar1.Value = Int(vPercent)
End Sub

Sub AllControlsEnabled(vEnabled As Boolean)
  mnuFile.Enabled = vEnabled
  mnuLogPopup.Enabled = vEnabled
  mnuCompile.Enabled = vEnabled
  lstProj.Visible = vEnabled
  framCompile.Visible = Not (vEnabled)
  If vEnabled = False Then
    mnuQSave.Enabled = False
    mnuSave.Enabled = False
  Else
    mnuQSave.Enabled = mChanged
    mnuSave.Enabled = mChanged
  End If
End Sub

Sub CompileList(vIndexes As String)
    
Dim CCD
  
  Dim Seconds As Long
  Dim sTime As Long
  Dim Indexes() As String
  
  Indexes() = Split(vIndexes, ",")
  
  mCompiling = True
  
  Dim x As Integer
  Dim I As Integer
  Dim CMD As String
  Dim Cnt As Integer
  Dim Cnt2 As Integer
  Dim Total As Integer
  'Disable Controls
  
  
    On Error Resume Next
    Kill App.Path + "\Error Logs\Make-Error-Log-*.Txt"
    On Error GoTo 0
  
  
  AllControlsEnabled (False)
  
  'Set splitter1.Control_Top_Or_Left = framCompile
  'Set splitter1.Control_Bottom_Or_Right = txtLog
  
  DoEvents

  Cnt = UBound(Indexes) + 1
  Total = Cnt
  'Set Mousepoint to Hourglass
  MousePointer = 11

Dim Vfile
For x = 0 To Cnt - 1
    Vfile = mProjects(Val(Indexes(x))).ProjectFullPath
    Call UpdateProjectInfo(Vfile)
Next
   
ReDim Info(Cnt)
Dim InfoR()
ReDim InfoR(Cnt)
Dim InfoI
Dim ErrorPro()
ReDim ErrorPro(Cnt)


  'Go Through List
  For x = 0 To Cnt - 1
    
    'Check for Cancel Being Pressed
    If CancelLoop Then
      
      'Add the Alert to the Log
      Call AddAlertToLog(TxtLog, "Operation Aborted by User")
      
      'Reset Cancel Variable
      CancelLoop = False
      
      'Stop the Loop
      Exit For
      
    End If
    
      
    sTime = GetTickCount()
    
    'Increment cnt2
    Cnt2 = Cnt2 + 1
  
    'Show 'Compiling' Arrow
    'picArrow.Visible = True
    
    'Let Process Finish
    DoEvents
    
    
    GoTo jump9
    
    If InStr(mProjects(Val(Indexes(x))).ExeFullPath, "BatchCompiler.exe") > 0 Then
        RR = mProjects(Val(Indexes(x))).ExeFullPath
        II2 = InStrRev(RR, ".")
        mProjects(Val(Indexes(x))).ExeFullPath = Mid$(mProjects(Val(Indexes(x))).ExeFullPath, 1, II2 - 1) + "-2.exe"
        ExitRun = True
        'MsgBox "01 " + Str(Now)
    End If
    
jump9:
    
    NoNo = 0
    
    If InStr(mProjects(Val(Indexes(x))).ExeFullPath, "BatchCompiler.exe") > 0 Then NoNo = 1
    
    
    If FileInUse(mProjects(Val(Indexes(x))).ExeFullPath) = True And NoNo = 0 Then
        NoNo = 1
        InfoI = "Errors -----------------------------" + vbCrLf + "Cannot Compile EXE Already Running" + vbCrLf + mProjects(Val(Indexes(x))).ExeFullPath + vbCrLf + "-----------------------------<<" + vbCrLf
        'InfoR(x) = InfoI
        ErrorPro(x) = mProjects(Val(Indexes(x))).ProjectName
        TxtLog.Text = TxtLog.Text + vbCrLf + InfoI
        BadRun = True
        Bads = Bads + 1
    End If

'    If InStr(mProjects(Val(Indexes(x))).ExeFullPath, "BatchCompiler") > 0 Then
'        NoNo = 1
'        ExitRun = True
'    End If
    
    'If InStr(mProjects(Val(Indexes(x))).ExeFullPath, "BatchCompiler-2.exe") > 0 Then
    
    'If FileInUse2(mProjects(Val(Indexes(X))).ProjectFullPath) = True Then
    '    NoNo = 1
    '    txtLog.Text = txtLog.Text + vbCrLf + "-----------------------------" + vbCrLf + "Cannot Compile EXE Already Running" + vbCrLf + mProjects(Val(Indexes(X))).ExeFullPath + vbCrLf + "-----------------------------"
    'End If
    
'    If InStr(mProjects(Val(Indexes(x))).ExeFullPath, "BatchCompiler.exe") > 0 Then NoNo = 1
    
    
    If NoNo = 0 Then
    'Build the command to shell
    CCD = CCD + 1
    DD$ = Mid$(mProjects(Val(Indexes(x))).ExeFullPath, InStrRev(mProjects(Val(Indexes(x))).ExeFullPath, "\") + 1)
    TF$ = App.Path + "\Error Logs\Make-Error-Log-" + DD$ + "-" + Format$(CCD, "0000") + ".Txt"
    If FileExists(TF$) = True Then
        On Error Resume Next
        Kill TF$
        On Error GoTo 0
    End If
    
    Call FileInUseDelay(TF$)
    FR1 = FreeFile
    Open TF$ For Output Lock Read As #FR1
    Print #FR1, "-----------------------------------"
    Print #FR1, Now
    Print #FR1, "-----------------------------------"
    Print #FR1, Left(mProjects(Val(Indexes(x))).ProjectFullPath, InStrRev(mProjects(Val(Indexes(x))).ProjectFullPath, "\"))
    Print #FR1, "-----------------------------------"
    Print #FR1, mProjects(Val(Indexes(x))).ProjectFullPath
    Print #FR1, "-----------------------------------"
    Print #FR1, mProjects(Val(Indexes(x))).ExeFullPath
    Print #FR1, "-----------------------------------"
    Print #FR1, DD$
    Print #FR1, "-----------------------------------"
    Print #FR1, TF$
    Print #FR1, "-----------------------------------"
    Close #FR1
    
    'Info(X) = "-----" + vbCrLf + Left(mProjects(Val(Indexes(X))).ProjectFullPath, InStrRev(mProjects(Val(Indexes(X))).ProjectFullPath, "\")) + vbCrLf
    'Info(X) = Info(X) + mProjects(Val(Indexes(X))).ProjectFullPath + vbCrLf
    'Info(X) = Info(X) + mProjects(Val(Indexes(X))).ExeFullPath + vbCrLf
    'Info(X) = Info(X) + DD$ + vbCrLf
    'Info(X) = Info(X) + TF$ + vbCrLf
    
    CMD = VBPath & " /out """ + TF$ & """ " & " /make " + Chr(34) & mProjects(Val(Indexes(x))).ProjectFullPath & Chr(34) & " " & Chr(34) & mProjects(Val(Indexes(x))).ExeFullPath & Chr(34)
    PROGCCOMPILED = True
    
    'Update Status Frame
    Call Dostatus(mProjects(Val(Indexes(x))).ProjectFullPath, mProjects(Val(Indexes(x))).ExeFullPath, (Cnt2 / Total) * 100)
    On Error GoTo 0
    
    'Shell the command to Complie
    'Call Shell(CMD)
          
    RETURNvAL = shellAndWait(CMD)
          
          
    'Add Compile to Log
    Call CompileToLog(TxtLog, mProjects(Val(Indexes(x))).ProjectFullPath, mProjects(Val(Indexes(x))).ExeFullPath, Cnt2, Cnt, True)
    
    End If
    
    'Hide 'Compiling' Arrow
    'picArrow.Visible = False
    
    'Pause - this is for effect so we actuall see the arrow when
    'Compiling relatively small projects
    'Pause (0.1)
      
    Seconds = Seconds + (GetTickCount - sTime)
    
    lblTime = Format((Seconds / Cnt2) / 1000, "0.0000") & " seconds"
    
    'Let Process Finish
    DoEvents
    
Next

For R = 1 To lstProj.ListItems.Count
    lstProj.ListItems.Item(R).Selected = False
Next

'Check the Loggs for errors

Pause 0.2
CCD = 0

For x = 0 To Cnt - 1

'    DD$ = Mid$(mProjects(Val(Indexes(X))).ExeFullPath, InStrRev(mProjects(Val(Indexes(X))).ExeFullPath, "\") + 1)
'    TF$ = App.Path + "\Error Logs\Make-Error-Log-" + DD$ + ".Txt"

    If DD$ <> "BatchCompiler.exe" Then
        CCD = CCD + 1
    End If
    CCD = 1
    DD$ = Mid$(mProjects(Val(Indexes(x))).ExeFullPath, InStrRev(mProjects(Val(Indexes(x))).ExeFullPath, "\") + 1)
    TF$ = App.Path + "\Error Logs\Make-Error-Log-" + DD$ + "-" + Format$(CCD, "0000") + ".Txt"
    If FileExists(TF$) = True Then
        DumVar = IsFileOpenDelay(TF$)
'        ET = 0
'        Do
'            DoEvents
'            ET = ET + 1
'        Loop Until FileSize(TF$) > 0
        'If ET > 1 Then Stop
        FR1 = FreeFile
        Open TF$ For Input As FR1
            II$ = Input(LOF(FR1), FR1)
            IQ = LOF(FR1)
        Close #FR1
        II$ = II$ + "---<<"
        
        
         ProFullPath = mProjects(Val(Indexes(x))).ProjectFullPath
         ProEXEPath = mProjects(Val(Indexes(x))).ExeFullPath

        
        Call UpdateFileLoggs
        
        
        'If II$ = "" Then Stop
        If InStr(II$, "succeeded") = 0 Then
             'If InStr(DD$, "BatchCompiler-2.exe") = 0 Then
             If InStr(DD$, "BatchCompiler.exe") = 0 Then
                'For R = 1 To lstProj.ListItems.Count
                '    If lstProj.ListItems.Item(R) = mProjects(Val(Indexes(x))).ProjectName Then
                '    If ErrorPro(x) = True Then
                '        lstProj.ListItems.Item(R).Selected = True
                '    End If
                'Next
                ErrorPro(x) = mProjects(Val(Indexes(x))).ProjectName
                Dim YY1, YY2
                YY1 = Len(TxtLog.Text)
                TxtLog.Text = TxtLog.Text + vbCrLf + "Error -----------------------------** Errors In Log" + vbCrLf + mProjects(Val(Indexes(x))).ExeFullPath + vbCrLf + II$ ' Info(X) + vbCrLf + II$
                
                'If InfoR(x) = "" Then
                '    TxtLog.Text = TxtLog.Text + vbCrLf + "Error -----------------------------** Errors In Log" + vbCrLf + mProjects(Val(Indexes(x))).ExeFullPath + vbCrLf + II$ ' Info(X) + vbCrLf + II$
                'Else
                '    TxtLog.Text = TxtLog.Text + vbCrLf + "Error -----------------------------** Errors In Log " + Info(x) + "---<<" + vbCrLf
                'End If
                
                YY2 = Len(TxtLog.Text)
                Bads = Bads + 1
                BadRun = True
            End If
        End If
            
                        
        
        If InStr(II$, "succeeded") > 0 Then
            On Error Resume Next
            Kill TF$
            On Error GoTo 0
            'Run the Exe After its been compiled
            Tagg = 0
            If DD$ = "BatchCompiler.exe" Then Tagg = 1
            If DD$ = "BatchCompiler-2.exe" Then Tagg = 1
            If DD$ = "ExitRun.exe" Then Tagg = 1
            'But Not If
            '#DontReRunCompiler.txt
            DD$ = Mid$(mProjects(Val(Indexes(x))).ExeFullPath, 1, InStrRev(mProjects(Val(Indexes(x))).ExeFullPath, "\"))
            DD$ = DD$ + "#DontReRunCompiler.txt"
            Dim BackAPath, A1
            A1 = InStrRev(DD$, "\")
            A1 = InStrRev(DD$, "\", A1 - 1)
            
            BackAPath = Mid$(DD$, 1, A1) + "#DontReRunCompiler.txt"
            If Dir$(DD$) <> "" Then
                Tagg = 1
            End If
            If Dir$(BackAPath) <> "" Then
                Tagg = 1
            End If
                
            Tagg = 1 'LOOKHERE
            
            
            
            
            Dim TXD
            TXD = LCase(" WMICPU2 MINI.exeVolumeBar.exeVolumeBar WinAmp.exeWinAmp MP3%.exeFast Clipboard.exeDrive_Detach.exeURL Logger.exeActive Idle.exeDrives_Gig.exeCid -RunMe.exeBBCWeather.exeTidal.exeRun Fav Programs.exe")
            If InStr(TXD, LCase(mProjects(Val(Indexes(x))).EXEName)) > 0 Then Tagg = 1
            
            
            
            If mProjects(Val(Indexes(x))).ExeFullPath = "D:\VB6\VB-NT\00_Best_VB_01\Auto Start Menu\Auto Start Menu.exe" Then
                'FS.CopyFile "D:\VB6\VB-NT\00_Best_VB_01\Auto Start Menu\Auto Start Menu.exe", "D:\TEMP\Auto Start Menu.exe"
            End If
            
            If Tagg = 0 Then
                'RARE USED
                'Run the Exe After its been compiled
                'Shell mProjects(Val(Indexes(x))).ExeFullPath, vbNormalNoFocus
            End If
        
            On Error Resume Next
            DumVar = IsFileOpenDelay(Tx$)
            Tx2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\Programming\BatchComplierLoggerCount2.txt"
            DumVar = IsFileOpenDelay(Tx2$)
            FS.CopyFile Tx$, Tx2$
            On Error GoTo 0
            
        
        End If
    End If

Next

If Dir(App.Path + "\#_Data", vbDirectory) = "" Then
    On Error Resume Next
    MkDir App.Path + "\#_Data"
    If err.Number > 0 Then
        MsgBox "ERROR MAKE DATA FOLDER" + vbCrLf + vbCrLf + App.Path + "\#_Data" + vbCrLf + vbCrLf + err.Description + vbCrLf + vbCrLf + "END"
        End
    End If
    On Error GoTo 0
End If

If Dir(App.Path + "\#_Data\" + GetComputerName, vbDirectory) = "" Then
    On Error Resume Next
    MkDir App.Path + "\#_Data\" + GetComputerName
    If err.Number > 0 Then
        MsgBox "ERROR MAKE DATA FOLDER" + vbCrLf + vbCrLf + App.Path + "\#_Data\" + GetComputerName + vbCrLf + vbCrLf + err.Description + vbCrLf + vbCrLf + "END"
        End
    End If
    On Error GoTo 0
End If



Tx$ = App.Path + "\#_Data\" + GetComputerName + "\" + GetComputerName + "__BatchComplierTextForVBCompiler.txt"
If Dir(Tx$) <> "" Then
    DumVar = IsFileOpenDelay(Tx$)
    FR1 = FreeFile
    Open Tx$ For Input As #FR1
        Line Input #FR1, AXx$
    Close #FR1
End If
'STRING LOOKS LIKE THIS
'*** Total  --- VB Projects Compiled in (13) Days  41

Do
AccessToFile = True
For x = 0 To Cnt - 1
    DD$ = mProjects(Val(Indexes(x))).ExeFullPath
    'If IsFileAlreadyOpen(DD$) Then AccessToFile = False: MsgBox "Delay"
    'If FileInUse(DD$) = True Then AccessToFile = False: MsgBox "Delay2"
Next
Loop Until AccessToFile = True


Call AddAlertToLog(TxtLog, "Finished Compiling " & Cnt2 & " Projects", (Seconds / Cnt2) / 1000)

'use this to fill report at end
On Error Resume Next
For x = 0 To Cnt - 1
    For R = 1 To lstProj.ListItems.Count - 1
        If ErrorPro(x) = lstProj.ListItems.Item(R) Then
            lstProj.ListItems.Item(R).Selected = True: Exit For
        End If
    Next
Next


'Call UpdateFileLoggs


'Re-Enable Controls
AllControlsEnabled (True)
mCompiling = False
  
'Reset Mousepoint to default
MousePointer = 0

lstProj.SetFocus
End Sub

Sub CompileFunc(vSelectedOnly As Boolean)
    
  Dim Seconds As Long
  Dim sTime As Long
  
  mCompiling = True
  
  Dim x As Integer
  Dim I As Integer
  Dim CMD As String
  Dim Cnt As Integer
  Dim Cnt2 As Integer
  Dim Total As Integer
  'Disable Controls
  
  AllControlsEnabled (False)
  
  'Set splitter1.Control_Top_Or_Left = framCompile
  'Set splitter1.Control_Bottom_Or_Right = txtLog
  
  DoEvents

  'Count Selected Items, if needed.
  If vSelectedOnly Then
    For x = 1 To lstProj.ListItems.Count
      If lstProj.ListItems(x).Selected Then Cnt = Cnt + 1
    Next
  End If

  Total = IIf(vSelectedOnly, Cnt, lstProj.ListItems.Count)

  'Set Mousepoint to Hourglass
  MousePointer = 11

  'Go Through List
  For x = 1 To lstProj.ListItems.Count
    
    'Check for Cancel Being Pressed
    If CancelLoop Then
      
      'Add the Alert to the Log
      Call AddAlertToLog(TxtLog, "Operation Aborted by User")
      
      'Reset Cancel Variable
      CancelLoop = False
      
      'Stop the Loop
      Exit For
      
    End If
    
    If (vSelectedOnly And lstProj.ListItems(x).Selected) Or Not (vSelectedOnly) Then
      
      sTime = GetTickCount()
      
      'Increment cnt2
      Cnt2 = Cnt2 + 1
    
      'Show 'Compiling' Arrow
      'picArrow.Visible = True
      
      'Let Process Finish
      DoEvents
      
      'Find Array's Index Value
      I = Val(Replace(LCase$(lstProj.ListItems(x).Key), "key:", ""))
      
      'Build the command to shell
      CMD = VBPath & " /make " & Chr(34) & mProjects(I).ProjectFullPath & Chr(34) & " " & Chr(34) & mProjects(I).ExeFullPath & Chr(34)
      
      'Update Status Frame
      Call Dostatus(mProjects(I).ProjectFullPath, mProjects(I).ExeFullPath, (Cnt2 / Total) * 100)
      
      
      
      'Shell the command
    '  Call Shell(CMD)
        RETURNvAL = shellAndWait(CMD)
            
      'Add Compile to Log
      Call CompileToLog(TxtLog, mProjects(I).ProjectFullPath, mProjects(I).ExeFullPath, Cnt2, IIf(vSelectedOnly, Cnt, lstProj.ListItems.Count))
      
      'Hide 'Compiling' Arrow
      'picArrow.Visible = False
      
      'Pause - this is for effect so we actuall see the arrow when
      'Compiling relatively small projects
      Pause (0.1)
        
      Seconds = Seconds + (GetTickCount - sTime)
      
      lblTime = Format((Seconds / Cnt2) / 1000, "0.0000") & " seconds"
    End If
    
    'Let Process Finish
    DoEvents
    
  Next

  'Set splitter1.Control_Top_Or_Left = lstProj
  'Set splitter1.Control_Bottom_Or_Right = txtLog

  Call AddAlertToLog(TxtLog, "Finished Compiling " & Cnt2 & " Projects", (Seconds / Cnt2) / 1000)

  'Re-Enable Controls
  AllControlsEnabled (True)
  mCompiling = False
  
  'Reset Mousepoint to default
  MousePointer = 0
End Sub

Private Sub lstProj_DblClick()
Dim MM
MM = lstProj.SelectedItem.Text

For We = 1 To ScanPath.ListView1.ListItems.Count
    A1$ = ScanPath.ListView1.ListItems.Item(We).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(We)
        If InStr(A1$ + B1$, MM) > 0 Then
            Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE """ + A1$ + B1$ + """", vbNormalFocus
            End
        End If
Next



End Sub


Private Sub lstProj_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button <> 2 Then Exit Sub
Dim MM
MM = lstProj.SelectedItem.Text

For We = 1 To ScanPath.ListView1.ListItems.Count
    A1$ = ScanPath.ListView1.ListItems.Item(We).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(We)
        If InStr(A1$ + B1$, MM) > 0 Then
            Shell "explorer.exe /e, /select, """ + A1$ + B1$ + """", vbNormalFocus
            End
        End If
Next


End Sub

Private Sub lstUpdates_Click()
  lstUpdates.ListIndex = -1
End Sub

Private Sub Mnu_ACompile_Click()
Call mnuAutoCompile_Click
End Sub


Private Sub MNU_EXIT_AFTER_Click()
EXIT_ON_END = True

DoEvents

If END_STATE = True Then Unload Me
End Sub

Private Sub MNU_EXIT_Click()
EXIT_ON_END = True

DoEvents

If END_STATE = True Then Unload Me
End Sub

Private Sub MNU_VB_FOLDER_Click()
Shell "EXPLORER /SELECT, " + App.Path + "\" + App.EXEName + ".vbp", vbMaximizedFocus
End Sub

Private Sub MNU_VB_ME_Click()
    
    Dim CODER_VBP_FILE_NAME_2
    Dim VB_1, VB_2, VB_3

    VB_1 = "C:\PROGRAM FILES\Microsoft Visual Studio\VB98\VB6.EXE"
    VB_2 = "C:\PROGRAM FILES (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
    If Dir(VB_1) <> "" Then VB_3 = VB_1
    If Dir(VB_2) <> "" Then VB_3 = VB_2
    '------------------------------------------------------
    'ADD MICROSFOT SCRIPTING RUNTIME FOR HERE IN REFERENCES
    '------------------------------------------------------
    Dim objShell
    Set objShell = CreateObject("Wscript.Shell")
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".vbp"
    objShell.Run """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
    Set objShell = Nothing
    
    EXIT_TRUE = True
    Unload Me
    Exit Sub
    
End Sub

Private Sub MNU_RUN_FAVS_Click()

Me.WindowState = 1
DoEvents

Shell "D:\VB6\VB-NT\00_Best_VB_01\Run Fav Progs\Run Fav Programs.exe GOGO", vbNormalFocus

EXIT_ON_END = True

DoEvents

If END_STATE = True Then Unload Me

End Sub

Private Sub Mnu_VB_Click()
Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\BatchCompiler.vbp""", vbNormalFocus
End

End Sub

Private Sub mnuAutoCompile_Click()
  SaveProfile (True)
  
  cmdBack.Enabled = False
  cmdCompile.Enabled = False
  
  mIndexes = ""
  lstUpdates.Clear
  lblCount = "0 Project(s)"
  
  'framCover.ZOrder 0
  'framCover.Visible = True
  
  Pause 0.5
  
  Dim Updates As Integer
  Const Tags As String = ",UserControl,Form,Module,Class,ResFile32,"
  Dim Tag As String
  Dim Num As Integer
  Dim EXE_Date As Date
  Dim Readln As String
  Dim File As String
  Dim Data As String
  Dim x As Integer
  Dim VBPDir As String
  Dim Indexes As String
  Dim Module_Date  As Date
  
  LockWindow lstProj.hWnd, True
  
  For x = 0 To UBound(mProjects())
    Do
    Jh7 = 0
    If Dir$(mProjects(x).ProjectFullPath) = "" Then
    x = x + 1: Jh7 = 1
    'MsgBox "Cant Find Project" + mProjects(X).ProjectFullPath ': End
    End If
    Loop Until Jh7 = 0 Or x = UBound(mProjects())
    
    With mProjects(x)
      VBPDir = GetDirectoryName(.ProjectFullPath)
      If Right(VBPDir, 1) = "\" Then VBPDir = Left(VBPDir, Len(VBPDir) - 1)
      
     ' If Dir$(.ExeFullPath) = "" Then MsgBox "Cant Find Exe " + .ExeFullPath: End
      'Debug.Print mProjects(X).ExeFullPath
       
       If Dir$(.ExeFullPath) <> "" Then EXE_Date = GetFileDate(.ExeFullPath) Else EXE_Date = 0
      
      
      If Dir$(.ProjectFullPath) <> "" Then
      FR1 = FreeFile
      Open .ProjectFullPath For Input As #FR1
        Do While Not EOF(FR1)
          Line Input #FR1, Readln
          If InStr(1, Readln, "=") > 0 Then
            Tag = Left$(Readln, InStr(1, Readln, "=") - 1)
            Data = Mid$(Readln, InStr(1, Readln, "=") + 1)
            Data = Mid$(Data, InStr(1, Data, ";") + 1)
            If InStr(1, Tags, Tag) > 0 Then
              If InStr(1, Data, "..") > 0 Then
                File = Trim(Replace(GetRelativePath(GetDirectoryName(Data), VBPDir), Chr(34), "")) & Trim(GetFileName(Data))
              Else
                If Mid(Trim(Data), 2, 1) = ":" Then
                  File = Data
                Else
                  If Mid$(Data, 1, 1) = """" Then Data = Mid$(Data, 2, Len(Data) - 2)
                  File = Trim(VBPDir) & "\" & Trim(Data)
                End If
              End If
              File = Replace(File, "\\", "\")
              Module_Date = GetFileDate(File)
              If DateDiff("s", EXE_Date, Module_Date) > 0 Then
                Updates = Updates + 1
                lstUpdates.AddItem " " & .ProjectFullPath
                mIndexes = mIndexes & x & ","
                Close #FR1
                Exit Do
              End If
            End If
          End If
        Loop
      Close #FR1
    End If
    End With
  Next
  
  If mIndexes = "" Then
    framCover.Visible = False
    LockWindow lstProj.hWnd, False
    Pause 0.1
    'MsgBox "None of the projects listed show changes." & vbCrLf & vbCrLf & "All Projects Up-To-Date", vbInformation + vbOKCancel, "Projects Up-To-Date"
    'BadRun = True
    Call CloseAllFiles
    
    
    
    Tx$ = App.Path + "\#_Data\" + GetComputerName + "\" + GetComputerName + "__BatchComplierTextForVBCompiler.txt"
    If Dir(Tx$) <> "" Then
        DumVar = IsFileOpenDelay(Tx$)
        FR1 = FreeFile
        Open Tx$ For Input As #FR1
            Line Input #FR1, AXx$
        Close #FR1
    End If
            
    Call AddAlertToLogSimple(TxtLog)
  Else
    lblCount = Trim(Str(Updates)) & " Project(s)"
    cmdBack.Enabled = True
    cmdCompile.Enabled = True
    'Pause 0.1
    Call CloseAllFiles
    Call cmdCompile_Click
  
  End If
  
  
End Sub

Private Sub mnuClearLog_Click()
  TxtLog.Text = ""
End Sub

Private Sub mnuCompileSingle_Click()
  
  'Compile Select Projects
  Call CompileFunc(True)

End Sub

Sub mnuCompileAll_Click()

  'Compile All Projects
  Call CompileFunc(False)

End Sub

Private Sub mnuExit_Click()

  'Unload the form, causing the application to end
  Unload Me

End Sub

Private Sub mnuNew_Click()
  
  'Reset the save location, so we are prompted next time.
  SaveLocation = ""
  
  'Clear List
  lstProj.ListItems.Clear
  
  'Reset Caption
  DoCaption ("")
  
  'Clear Projects Array
  ReDim mProjects(0)
  
  TxtLog.Text = ""
  
  'Set Changed=False
  Call DoChanged(False)
  
End Sub


Sub OptionalTouchFiles()
'Exit Sub

ScanPath.cboMask = "*.bas;*.cls;*.frm"
'ScanPath.cboDate./ = 0
'ScanPath.DTPicker1(0) = Now - (DaysToScan + 1) 'All in Last Ten Days
ScanPath.DTPicker1(0) = vbNullString

ScanPath.ListView1.ListItems.Clear

ScanPath.txtPath.Text = "D:\VB6\VB-NT"
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "X:\00 Lists-Common-Words\Acronyms Code"
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "X:\00 Lists-Common-Words\Double Word Code"
Call ScanPath.cmdScan_Click

ReDim mProjects(0)
            
Dim Xdate, OldA1, TF, FC, F1

For We = 1 To ScanPath.ListView1.ListItems.Count
    A1$ = ScanPath.ListView1.ListItems.Item(We).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(We)
    
    If A1$ <> OldA1 Then
    OldA1 = A1$
    Xdate = 0
    End If
    
    FileSpec1 = A1$ + B1$
    Set F = FS.GetFile((FileSpec1))
    ADate1 = F.DateLastModified
    If ADate1 > Xdate Then
        Xdate = ADate1
    End If
    
    If We + 1 < ScanPath.ListView1.ListItems.Count Then
        If ScanPath.ListView1.ListItems.Item(We + 1).SubItems(1) <> OldA1 Then
            Set F = FS.GetFolder(A1$)
            Set FC = F.Files
            For Each F1 In FC
                DoEvents
                If InStr(LCase(Right(F1.Name, 4)), ".vbp") > 0 Then
                    
                    If Xdate > F1.DateLastModified Then
                        TF = LastModifiedToCurrent(A1$ + F1.Name)
                    End If
                End If
            Next
        End If
    End If
Next


End Sub

Sub OpenProfileScan(Vfile As String)

'Use this if you ever rename a folder then use a search program to replace file names
'then you want to complie those projects the .vbp will have to be date touched
'if any forms are seen to be higher dates than project
Call OptionalTouchFiles


DaysToScan = 15

ScanPath.cboMask = "*.vbp"
ScanPath.cboDate.ListIndex = 0
ScanPath.DTPicker1(0) = Now - (DaysToScan + 1) 'All in Last Ten Days

ScanPath.ListView1.ListItems.Clear

ScanPath.txtPath.Text = "D:\VB6\VB-NT"
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "X:\00 Lists-Common-Words\Acronyms Code"
Call ScanPath.cmdScan_Click
ScanPath.txtPath.Text = "X:\00 Lists-Common-Words\Double Word Code"
Call ScanPath.cmdScan_Click

ReDim mProjects(0)
            
            
For We = 1 To ScanPath.ListView1.ListItems.Count
    A1$ = ScanPath.ListView1.ListItems.Item(We).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(We)
    
    FileSpec1 = A1$ + B1$
    Set F = FS.GetFile((FileSpec1))
    ADate1 = F.DateLastModified
    ScanPath.ListView1.ListItems.Item(We).SubItems(2) = Format(ADate1, "YYYY-MM-DD HH:MM:SS")
    
Next

ScanPath.ListView1.SortOrder = lvwDescending
ScanPath.ListView1.SortKey = 2
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False
            
'Exit Sub

For We = 1 To ScanPath.ListView1.ListItems.Count
    A1$ = ScanPath.ListView1.ListItems.Item(We).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(We)
    
    FileSpec1 = A1$ + B1$
    Set F = FS.GetFile((FileSpec1))
    ADate1 = F.DateLastModified
    If ADate1 > Now - (DaysToScan) Then
        AddProject (A1$ + B1$)
    'Else
    'A = A
    End If
Next

ReDim Preserve mProjects(UBound(mProjects) - 1)

End Sub

Sub SaveProfile(Optional vSaveAs As Boolean)
  On Error GoTo err
  
  SaveLocation = LastProfile 'App.Path + "\Matts NewAlphaCommScripts.BPR"
  
  'If SaveLocation is blank, or 'Save As' was clicked, show the dialog
  If SaveLocation = "" Or vSaveAs Then
    
    CD.Filter = "BPR files (*.bpr)|*.BPR"
    
    'Trigger Error to Exit sub if Cancel was Pressed
    CD.CancelError = True
    
    CD.Filename = ""
    
    'Show save Dialog box
    'CD.ShowSave
    'SaveLocation = App.Path + "\Matts NewAlphaCommScripts.BPR"
    
    'Set the current Save location to the file selected
    'SaveLocation = CD.FileName
    
  End If
  
  Dim fNum As String
  
  'Find Next Open File Number
  FR1 = FreeFile
  'Check for over-write
  Dim Answer As Integer
  If vSaveAs And Dir(SaveLocation) <> "" Then
    
    'If trying to save over an existing file, prompt to confirm.
    'Answer = MsgBox("The file already exists. OK to over-write?", vbYesNo + vbQuestion, "Over-Write?")
    
    Answer = vbYes
  
  ElseIf vSaveAs Then
    
    'if the user just presses 'Save' and we already know SaveLocation, just save
    'over existing file.
    Answer = vbYes
    
  End If
  
  'Should we delete the existing file?
  If Answer = vbYes Then
    
    'If we want to delete the file, lets do it.
    If Dir(SaveLocation) <> "" Then Kill SaveLocation
    
  Else
  
    'if not, we can not save.
    Exit Sub
    
  End If
   
  'Simple File Write to save settings
  Dim x As Integer

  FR1 = FreeFile
  'Open Output file
  Open SaveLocation For Output As #FR1
    
    'Loop through projects
    For x = 0 To UBound(mProjects)
    
      'Print Data to file
      Print #FR1, mProjects(x).ProjectFullPath
      Print #FR1, mProjects(x).ExeFullPath
      
    Next
    
  'Close File
  Close #FR1
  
  'Update Caption
  DoCaption Left(GetFileName(SaveLocation), Len(GetFileName(SaveLocation)) - 4)
  
  'SaveSettings
  SaveSetting "BatchCompile", "Settings", "LastProfile", SaveLocation
  
  'Since we've just saved, there are no changes to this project
  Call DoChanged(False)
  
err:
End Sub

Sub DoChanged(vChanged As Boolean)
  'Set mChanged Varialable
  mChanged = vChanged
  
  'Set Control(s) appropriately
  mnuQSave.Enabled = vChanged
  mnuSave.Enabled = vChanged
End Sub

Sub PopulateProjects()
  Dim x As Integer
  Dim LI As ListItem
  
  LockWindow lstProj.hWnd, True
  
  'Delete all Current Items
  lstProj.ListItems.Clear
  
  'populate listview with array data
  For x = 0 To UBound(mProjects())
    
    'Set List Item Text (Cell 1)
    Set LI = lstProj.ListItems.Add(x + 1, "Key:" & x, mProjects(x).ProjectName)
    
    'Set List Item's SubItem Text (Cell 2)
    LI.SubItems(1) = mProjects(x).EXEName
    
  Next
  
  LockWindow lstProj.hWnd, False
  
End Sub


Private Sub mnuQSave_Click()
  
  SaveProfile (True)

End Sub

Private Sub mnuSave_Click()
  
  SaveProfile (True)
  
End Sub

Private Sub mnuSaveAs_Click()
  
  'Call Save As Routine
  SaveProfile (True)
  
End Sub

Sub DoCaption(vText As String)
  
  'Set Application's Caption
  If vText = "" Then
    
    'Case to reset text to default
    Caption = " Batch VB6 Compiler"
  
  Else
    
    'Case to display information
    Caption = " Batch VB6 Compiler - " & vText
  
  End If

End Sub

Private Sub mnuSaveLog_Click()
  On Error GoTo err
    
  CD.Filter = "LOG files (*.log)|*.LOG"
  
    'Trigger Error to Exit sub if Cancel was Pressed
    CD.CancelError = True
    
    CD.Filename = App.Path & "\" & Format(Date, "mmddyy") & "-" & Format(Time, "hhmm AMPM")
    
    'Show save Dialog box
    CD.ShowSave
    
    Dim Num As Integer
    Dim File As String
    File = CD.Filename
    
    FR1 = FreeFile()
    Open File For Output As #FR1
      Print #FR1, TxtLog.Text
    Close #FR1
    
err:
End Sub

Private Sub Timer1_Timer()
'If Me.WindowState <> 0 Then Exit Sub
'If GetForegroundWindow <> Me.hWnd Then Exit Sub

'On Error Resume Next
'txtLog.SetFocus

End Sub

Private Sub txtLog_Change()
  mnuLogPopup.Enabled = Not (Trim(TxtLog.Text) = "")
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
  hFile = OpenFile(strFilePath, FileInfo, &H1)
  If hFile = -1 And err.LastDllError = 32 Then
    FileInUse = True
  End If

CloseHandle hFile
  
End Function

Function FileInUse2(ByVal strFilePath As String) As Boolean
  
  Dim hFile As Long
  Dim FileInfo  As OFSTRUCT
  
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function
  If Right(strFilePath, 1) <> Chr(0) Then strFilePath = strFilePath & Chr(0)
  
  FileInfo.cBytes = Len(FileInfo)
  hFile = OpenFile(strFilePath, FileInfo, &H40)
If InStr(strFilePath, "Batch") Then Stop
If hFile = -1 And err.LastDllError = 32 Then
    FileInUse2 = True
End If
CloseHandle hFile
'Stop
'Const OF_READ = &H0
'Const OF_WRITE = &H1
'Const OF_READWRITE = &H2
'Const OF_SHARE_COMPAT = &H0
'Const OF_SHARE_EXCLUSIVE = &H10
'Const OF_SHARE_DENY_WRITE = &H20
'Const OF_SHARE_DENY_READ = &H30
'Const OF_SHARE_DENY_NONE = &H40

  
End Function

Sub FileInUseDelay(Tx$)
        
    QQ = Now + TimeSerial(0, 1, 30)
    Do
'        DoEvents
        II = FileInUse(Tx$)
        If II = True Then Sleep (200)
    Loop Until II = False Or QQ < Now
    
    If II = True Then
        MsgBox "Trouble File Stuck In Use " + vbCrLf + Tx$ + vbCrLf + "Longer than 1 min 30 secs"
        Stop
    End If
End Sub

Function IsFileAlreadyOpen(Filename As String) As Boolean
    Dim hFile As Long
    Dim lastErr As Long
    hFile = -1
    lastErr = 0
    hFile = lOpen(Filename, &H10)
    If hFile = -1 Then
        lastErr = err.LastDllError
    Else
        lClose (hFile)
    End If
    IsFileAlreadyOpen = (hFile = -1) And (lastErr = 32)
End Function

Function IsFileOpenDelay(Tx$)
    QQ = Now + TimeSerial(0, 4, 0)
    Do
        DoEvents
        II3 = IsFileAlreadyOpen(Tx$)
        If II3 = True Then Sleep (200)
    Loop Until II3 = False Or QQ < Now
    If II3 = True Then
        MsgBox "Trouble File Stuck In Use " + vbCrLf + Tx$ + vbCrLf + "Longer than 4 Min"
        Stop
    End If
End Function



Function FileExists(ByVal strFilePath As String) As Boolean
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function





Private Sub WaitWavFinish()

Do
    DoEvents
Loop Until MMControl1.Mode = 525

End Sub

Private Sub txtLog_LostFocus()
'txtLog.SetFocus

End Sub


Sub UpdateFileLoggs()
            
    Dim LVS
    Dim R_RESULT
    Dim GOOD
    Dim TX1
            
    TX1 = Tx$ = App.Path + "\#_Data\" + GetComputerName
    Tx$ = App.Path + "\#_Data\" + GetComputerName + "\" + GetComputerName + "__BatchComplierLogger.txt"
    
    R_RESULT = CreateFolderTree(TX1)
    If R_RESULT = True Then
        GOOD = "YES"
    Else
        MsgBox "UNABLE TO MAKE FOLDER " + vbCrLf + Tx$
    End If
    
    DumVar = IsFileOpenDelay(Tx$)
    FR1 = FreeFile
    Open Tx$ For Append As #FR1
        TN = Now
        Print #FR1, Format$(TN, "DDD DD-MM-YYYY HH:MM:SS "); " -- " + ProFullPath
        Print #FR1, Format$(TN, "DDD DD-MM-YYYY HH:MM:SS "); " -- " + ProEXEPath
    Close #FR1
        
        
    On Error Resume Next
    'Compilied projects count new thing include ones not compile succeed
    Tx$ = App.Path + "\#_Data\" + GetComputerName + "\" + GetComputerName + "__BatchComplierLoggerDatesCount.txt"
    DumVar = IsFileOpenDelay(Tx$)
    FR1 = FreeFile
    Open Tx$ For Input As #FR1
        Line Input #FR1, LVS
    Close #FR1
    LV2 = Val(LVS) + 1
        
    Tx$ = App.Path + "\#_Data\" + GetComputerName + "\" + GetComputerName + "__BatchComplierLoggerDates.txt"
    DumVar = IsFileOpenDelay(Tx$)
    FR1 = FreeFile
    Open Tx$ For Append As #FR1
        Print #FR1, Format$(TN, "DD-MM-YYYY HH:MM:SS") + Str(LV2)
    Close #FR1

    Tx$ = App.Path + "\#_Data\" + GetComputerName + "\" + GetComputerName + "__BatchComplierLoggerDatesCount.txt"
    DumVar = IsFileOpenDelay(Tx$)
    FR1 = FreeFile
    Open Tx$ For Output As #FR1
        Print #FR1, LV2
    Close #FR1
    
    Tx$ = App.Path + "\#_Data\" + GetComputerName + "\" + GetComputerName + "__BatchComplierDaysToScan.txt"
    DumVar = IsFileOpenDelay(Tx$)
    FR1 = FreeFile
    Open Tx$ For Input As #FR1
        Line Input #FR1, LV
    Close #FR1
    DaysToScan2 = LV
    
    Tx$ = App.Path + "\#_Data\" + GetComputerName + "\" + GetComputerName + "__BatchComplierDaysToScan.txt"
    DumVar = IsFileOpenDelay(Tx$)
    FR1 = FreeFile
    Open Tx$ For Output As #FR1
        Print #FR1, DaysToScan
    Close #FR1
    
    LV = 0
    Tx$ = App.Path + "\#_Data\" + GetComputerName + "\" + GetComputerName + "__BatchComplierLoggerSeek.txt"
    DumVar = IsFileOpenDelay(Tx$)
    FR1 = FreeFile
    Open Tx$ For Input As #FR1
        If Not (EOF(1)) Then Line Input #FR1, LV
    Close #FR1
    HOX = Val(LV)
    
    If Val(DaysToScan2) <> DaysToScan Then HOX = 1
    
    LV3 = Now - DaysToScan
    Tx$ = App.Path + "\#_Data\" + GetComputerName + "\" + GetComputerName + "__BatchComplierLoggerDates.txt"
    DumVar = IsFileOpenDelay(Tx$)
    FR1 = FreeFile
    Open Tx$ For Input Lock Write As #FR1
        If HOX = 0 Then HOX = 1
        Seek #FR1, HOX
        Do
        HOX2 = Seek(1)
        Line Input #FR1, LV
        err.Clear
        On Error Resume Next
        LV4 = DateValue(Mid(LV, 1, 19)) + TimeValue(Mid(LV, 1, 19))
        If err.Number > 0 Then
            On Error GoTo 0
            Line Input #FR1, LV
            LV4 = DateValue(Mid(LV, 1, 19)) + TimeValue(Mid(LV, 1, 19))
        End If
        
        TV2 = Val(Mid(LV, 20))
        If LV4 >= LV3 And HO = 0 Then
            HO = TV2
            HOX = HOX2 - 10
            If HOX < 0 Then HOX = 1
        Exit Do
        End If
        Loop Until EOF(FR1)
    Close #FR1
    
    Tx$ = App.Path + "\#_Data\" + GetComputerName + "\" + GetComputerName + "__BatchComplierLoggerSeek.txt"
    DumVar = IsFileOpenDelay(Tx$)
    FR1 = FreeFile
    Open Tx$ For Output As #FR1
        Print #FR1, HOX
    Close #FR1

    Tx$ = App.Path + "\#_Data\" + GetComputerName + "\" + GetComputerName + "__BatchComplierLoggerCount1.txt"
    DumVar = IsFileOpenDelay(Tx$)
    FR1 = FreeFile
    Open Tx$ For Output Lock Write As #FR1
        Print #FR1, "Total  --- VB Projects Compiled in (" + Trim(Str(DaysToScan)) + ") Days =" & Str(LV2 - (HO - 1))
        Print #FR1, "Unique - VB Projects Compiled in (" + Trim(Str(DaysToScan)) + ") Days =" & Str(ScanPath.ListView1.ListItems.Count)
    Close #FR1

    Tx$ = App.Path + "\#_Data\" + GetComputerName + "\" + GetComputerName + "__BatchComplierLoggerCount2.txt"
    DumVar = IsFileOpenDelay(Tx$)
    FR1 = FreeFile
    Open Tx$ For Output Lock Write As #FR1
        
        Print #FR1, "Last Project Compiled " + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p")
        XXF = DateDiff("d", LV4, Now)
        
        'Botch Fix
        If XXF < 1 Then XXF = 15 'DaysToScan
        
        'for compiler leave this one
        AXx$ = "*** Total  --- VB Projects Compiled in (" + Trim(Str(XXF)) + ") Days " & Str(LV2 - (HO - 1)) & vbCrLf
        
        Print #FR1, "Total -- VB Projects Compiled in (" + Trim(Str(XXF)) + ") Days" & Str(LV2 - (HO - 1))
        Print #FR1, "Unique VB Projects Compiled in (" + Trim(Str(DaysToScan)) + ") Days" & Str(ScanPath.ListView1.ListItems.Count)
    Close #FR1

    Tx$ = App.Path + "\#_Data\" + GetComputerName + "\" + GetComputerName + "__BatchComplierTextForVBCompiler.txt"
    DumVar = IsFileOpenDelay(Tx$)
    FR1 = FreeFile
    Open Tx$ For Output As #FR1
        Print #FR1, AXx$
    Close #FR1

End Sub

Public Function CreateFolderTree(ByVal sPath As String) As Boolean
    Dim nPos As Integer

    On Error GoTo CreateFolderTreeError
    
    nPos = InStr(sPath, "\")
    While nPos > 0
        
        '----------------------------------------------------------------------------
        'ROUTINE TAKEN FROM
        '----------------------------------------------------------------------------
        'SEND_TO_SCRIPT_IRFAR - Microsoft Visual Basic [design] - [mdlFileSys (Code)]
        'D:\VB6\VB-NT\00_Send_To\Send To Text List of Files & Sub Folders IRFAR\#0 Send To Text List of Files and Sub Folders IRFAN.exe
        '----------------------------------------------------------------------------
        'MODIFIED A BIT DIR COMMAND REPLACE MORE COMPLEX WAY
        '----------------------------------------------------------------------------
        
        'If Not FolderExists(Left$(sPath, nPos - 1)) Then
        
        If Dir((Left$(sPath, nPos - 1)), vbDirectory) = "" Then
            MkDir Left$(sPath, nPos - 1)
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
    'If Not FolderExists(sPath) Then MkDir sPath
    If Dir(sPath, vbDirectory) = "" Then MkDir sPath
    
    CreateFolderTree = True
    Exit Function

CreateFolderTreeError:
    Exit Function
End Function


