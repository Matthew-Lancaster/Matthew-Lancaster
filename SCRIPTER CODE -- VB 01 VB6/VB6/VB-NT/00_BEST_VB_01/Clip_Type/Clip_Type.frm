VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.Ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.Ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form Form1 
   Caption         =   "Clip Type Form"
   ClientHeight    =   7572
   ClientLeft      =   168
   ClientTop       =   1068
   ClientWidth     =   13608
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Clip_Type.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7572
   ScaleWidth      =   13608
   Begin VB.Timer Timer_FORM_LOAD 
      Interval        =   1
      Left            =   9060
      Top             =   4248
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   2688
      Left            =   60
      TabIndex        =   17
      Top             =   792
      Width           =   4668
      _ExtentX        =   8234
      _ExtentY        =   4741
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Clip_Type.frx":030A
   End
   Begin VB.Timer Timer_1_MINUTE 
      Interval        =   1000
      Left            =   10404
      Top             =   3624
   End
   Begin VB.Timer Timer_SAVED 
      Interval        =   1
      Left            =   9000
      Top             =   3636
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   10044
      Top             =   3612
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   9720
      Top             =   3636
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2016
      Left            =   5880
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9372
      Top             =   3612
   End
   Begin MCI.MMControl MMControl1 
      Height          =   336
      Left            =   4140
      TabIndex        =   5
      Top             =   3648
      Visible         =   0   'False
      Width           =   4368
      _ExtentX        =   7684
      _ExtentY        =   572
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   10164
      Top             =   2640
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   10152
      Top             =   2040
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   720
      TabIndex        =   16
      Top             =   0
      Width           =   675
   End
   Begin VB.Label GRAB_SMS_FILE 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   15
      TabIndex        =   15
      Top             =   375
      Width           =   630
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "FILE DOWN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2505
      TabIndex        =   14
      Top             =   375
      Width           =   6780
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   675
      TabIndex        =   13
      Top             =   375
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1110
      TabIndex        =   12
      Top             =   375
      Width           =   915
   End
   Begin VB.Label Label_Eddie 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Eddie"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5700
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label_UnKnown 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "UnKnown"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6615
      TabIndex        =   10
      Top             =   0
      Width           =   1350
   End
   Begin VB.Label Mum_G 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Group"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3285
      TabIndex        =   9
      Top             =   0
      Width           =   915
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   11010
      TabIndex        =   8
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1425
      TabIndex        =   6
      Top             =   0
      Width           =   675
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "MS Word"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8025
      TabIndex        =   4
      Top             =   0
      Width           =   1260
   End
   Begin VB.Label Label_dad 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Dad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5055
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label_Hosso 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Hoss"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4275
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label_Mum 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Mum"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2505
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   675
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT END"
   End
   Begin VB.Menu MUN_EXIT_UNLOAD 
      Caption         =   "EXIT UNLOAD"
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "!! VB"
   End
   Begin VB.Menu MNU_VBFOLDER 
      Caption         =   "!! VB FOLDER"
   End
   Begin VB.Menu MNU_ME_ON_TOP 
      Caption         =   "ME ON TOP"
   End
   Begin VB.Menu MNU_ADMINISTRATOR 
      Caption         =   "ADMINISTRATOR"
   End
   Begin VB.Menu MNU_SAVED 
      Caption         =   "SAVED=TRUE "
   End
   Begin VB.Menu MNU_FILE1 
      Caption         =   "FILE 1"
      Begin VB.Menu MNU_NOTEPAD 
         Caption         =   "NOTEPAD EDIT"
      End
      Begin VB.Menu MNU_INFO_RAPID 
         Caption         =   "INFO RAPID"
      End
   End
   Begin VB.Menu MNU_FILE2 
      Caption         =   "FILE2"
      Begin VB.Menu MNU_GRAB_SMS_FILE 
         Caption         =   "!! GRAB A SMS FILE"
      End
      Begin VB.Menu Mnu_Clue 
         Caption         =   "!! Clued Up"
         Shortcut        =   ^C
      End
      Begin VB.Menu Mnu_Logg 
         Caption         =   "!! Logg"
         Shortcut        =   ^L
      End
      Begin VB.Menu Mnu_Job 
         Caption         =   "!! Jobs"
         Shortcut        =   ^J
      End
      Begin VB.Menu Mnu_Prog 
         Caption         =   "!! Programming"
         Shortcut        =   ^P
      End
      Begin VB.Menu Mnu_UNET 
         Caption         =   "!! USENET"
         Shortcut        =   ^U
      End
      Begin VB.Menu Mnu_KadMan 
         Caption         =   "!! KADMAN"
         Shortcut        =   ^K
      End
      Begin VB.Menu Mnu_Conflict 
         Caption         =   "!! &CONFLICT"
         Shortcut        =   ^O
      End
      Begin VB.Menu MNU_MY_HARDCORE_FEMALE 
         Caption         =   "!! MY FEMALE"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu MNU_LS_LAST_LOAD 
      Caption         =   "FILE _LOAD LAST"
   End
   Begin VB.Menu MNU_LS_SAVE_LAST 
      Caption         =   "FILE _SAVE LAST"
   End
   Begin VB.Menu Mnu_Emailmnu 
      Caption         =   "EMAIL"
      Begin VB.Menu Mnu_Email 
         Caption         =   "!! Email MAZ EDD"
      End
      Begin VB.Menu Mnu_Email_Val 
         Caption         =   "!! Email VAL"
      End
      Begin VB.Menu Mnu_Email_Gill 
         Caption         =   "!! Email Gill"
      End
      Begin VB.Menu Mnu_Email_Judy 
         Caption         =   "!! Email Judy"
      End
      Begin VB.Menu Mnu_Email_Andy 
         Caption         =   "!! Email Andy"
      End
   End
   Begin VB.Menu Mnu_FixCaps 
      Caption         =   "!! --- FIX CAPS - PRO"
   End
   Begin VB.Menu Mnu_Caps 
      Caption         =   "! ALL CAPS"
   End
   Begin VB.Menu Mnu_Fix1st_Letters 
      Caption         =   "! FIX 1ST LETT"
   End
   Begin VB.Menu Mnu_Fix1st_LettersNOLOW 
      Caption         =   "! FIX 1ST LETT No Low ---"
   End
   Begin VB.Menu Mnu_CLip 
      Caption         =   "!! CLIP"
   End
   Begin VB.Menu Mnu_Blue 
      Caption         =   "!! BLUE"
   End
   Begin VB.Menu Mnu_BToT 
      Caption         =   "!! Blue ToT"
   End
   Begin VB.Menu Mnu_Blue_Fol 
      Caption         =   "!! BLUE FOLDER"
   End
   Begin VB.Menu Mnu_Dont_Alter 
      Caption         =   "!! Dont Alter On Save"
   End
   Begin VB.Menu Mnu_SendBT_W5 
      Caption         =   "!! SEND TO BLUET W5"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_SendBT_K8 
      Caption         =   "!! SEND TO BLUET K8"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_SendBT_W100i 
      Caption         =   "!! SEND TO W100i"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'WORK TO DO ON THIS ' File1

Dim XVB_DATE




Dim COUNTIP()
Dim FILESAVENAME1_SAME
Dim FILESAVENAME1_SAME_DONE_READY

Public UNLOAD_BEGUN
Public CLIPBOARD_EDITED

Public FS, WINAMPEDIT
Private Const ipc_isplaying As Long = 104
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Dim TU As Double, WINSTATETX, OLDMR
Const WINAMP_BUTTON3 = 40046
Const WM_CLOSE = &H10
Const WM_USER = &H400
Const WM_COMMAND = &H111
Private XcNul As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

Dim LOWERTIMER, OTEXT
Dim OLENL
Dim TIMERSAVESAMEFILE, FILESAVENAME1, FILESAVENAME2
Dim BlueDontSaveOnExit, SMSLEN, EDITED2, EDITED3
Dim BlueDontSaveWhenOnExit
Dim Test, BLUESEND, Source1
Public EDITED
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim Buffer1 As String, FF1, EGG, OGG, ClipedText, DontAlter
Dim Buffer2 As String
Dim A1 As String, A2 As String



Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

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



Private Sub Mnu_Exit_Click()
End
End Sub

Private Sub MNU_VB_FOLDER_Click()
    Beep
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + CODER_EXE_FILE_NAME_1 + """", vbNormalFocus
    EXIT_TRUE = True
    
    Me.WindowState = 1

'    Unload Me
    'End
End Sub
Private Sub MNU_ADMINISTRATOR_Click()
    'Call MNU_ADMINISTRATOR_Click
    Beep
    MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR MODE ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "MODE ON", "MODE OFF")
    End If
End Sub
'-------------------------------------------
'-------------------------------------------
Private Sub Form_Load()

Set FS = CreateObject("Scripting.FileSystemObject")
        
' Set properties needed by MCI to open.
MMControl1.Notify = True
MMControl1.Wait = True
MMControl1.Shareable = False
MMControl1.DeviceType = "WaveAudio"

'Loads the first wave in any DIR for sound effect
'ChDir App.Path
DD = Dir(App.Path + "\*.WAV")
MMControl1.FileName = App.Path + "\" + DD

sWindowsFolder = GetSpecialfolder(36)

sMyDocsFolder = GetSpecialfolder(5)
Dim R As Long
'On Error Resume Next
'For R = 0 To 120
'    Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
'Next
'End
    
'Set poSendMail = New clsSendMail


EDITED = False
CLIPBOARD_EDITED = False
Me.Show
'Timer_FORM_LOAD.Enabled = True

End Sub

Private Sub Form_Resize()
'If Me.Visible = False Then Exit Sub
WINSTATETX = Me.WindowState
Me.SetFocus
End Sub

Private Sub MNU_ME_ON_TOP_Click()

K = AlwaysOnTop(Me.hWnd)

End Sub

Private Sub Timer_FORM_LOAD_Timer()

Timer_FORM_LOAD.Enabled = False
Call ME_TOP_MOVE

RTB1.Width = 14000
RTB1.Height = 8000
RTB1.Top = Label4.Top + Label4.Height
RTB1.Left = 0
RTB1.Font.Size = 20
RTB1.Font.Name = "ARIEL"
RTB1.Font.Bold = True
RTB1.Font.Weight = 1000
RTB1.Enabled = True ' THIS CONTROL SHOW ENABLE BUT IS NOT UNTIL FORM FULL LOAD UP WE ENABLER
DoEvents
X = 1
Y = 1
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True And Control.Visible = True Then
        'Debug.Print Control.Name
        'If Control.Name = "RTB1" Then MsgBox Control.Name
        If Control.Width + Control.Left > X Then X = Control.Width + Control.Left
        If Control.Height + Control.Top > Y Then Y = Control.Height + Control.Top
        If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
    End If
Next
On Error GoTo 0

Me.Width = (X + 90)
If mnu = 1 Then
    pluso = 950 ': pluso = 1100 'Sometimes Different
Else
    pluso = 450
End If

Me.Height = (Y + pluso)

'K = AlwaysOnTop(Me.hWnd)
Form1.SetFocus
RTB1.SetFocus

RTB1 = Format(Now, "dd-MMM hh:mm") + vbCrLf
RTB1.SelStart = Len(RTB1)

End Sub

Sub ME_TOP_MOVE()

Me.Left = (Screen.Width - Me.Width) / 2
Me_Top = (Screen.Height - Me.Height) / 2
Me_Top = Me_Top - 1000
If Me_Top < 0 Then Me_Top = 0
Me.Top = Me_Top

End Sub

Private Sub Form_Unload(Cancel As Integer)

Call ENDSAVE


Unload Form2

If EDITED = False Then Exit Sub

UNLOAD_BEGUN = True

If FF1 >= 1 And BlueDontSaveOnExit = False Then

'Call Mnu_Blue_Click

Mnu_SendBT_W5_Click
TRIAG = 1

End If



If Trim(RTB1.Text) = "" Then Exit Sub

If Len(RTB1) < 200 And ClipedText = False And DontAlter = False Then Call Mnu_Fix1st_LettersNOLOW_Click


If CLIPBOARD_EDITED = True Then
    Clipboard.Clear
    Sleep 400
    Clipboard.SetText (RTB1.Text)
    CLIPBOARD_EDITED = False
End If


'RTB1 = "**** Quick Type **** " + Format(Now, "DDD DD-MMM-YY HH:MM:SS") + " ****" + vbCrLf + RTB1

Source1 = App.Path + "\TextToday__" + GetComputerName + "_.txt"
FileInUseDelay (Source1)
FR1 = FreeFile
Open Source1 For Append As #FR1
    Print #FR1, RTB1.Text
Close #FR1

'GO LOGG NOW NOT CLUE

'NO BOTH


'NOT REQUIRED NOT A SUB ROUTINE ANYWHERE BUT LOAD OR SAVE ONE ONLY
'------------------------------------------------------------------------------------------------------------------
'Call Mnu_Clue_Click
'If TRIG = 0 Then Call Mnu_Logg_Click


End
End Sub


Private Sub Label_Eddie_Click()
    
    RTB1.Text = "Dear Eddie -- " + Format(Now, "DDD DD-MMM-YY HH:MMa/p") + vbCrLf + RTB1.Text + vbCrLf + "Bro" + vbCrLf + "Matthiaoso"
    EDITED = True

    DoEvents

    If Mid(RTB1, 1, 1) = vbCr Then RTB1 = Mid(RTB1, 3)

    FF1 = 7
    EDITED = False
    'CLIPBOARD_EDITED = False
End Sub


Private Sub Label_Mum_Click()
    
    RTB1.Text = "Dear Mum -- " + Format(Now, "DDD DD-MMM-YY HH:MMa/p") + vbCrLf + RTB1.Text + vbCrLf + "Love xxy" + vbCrLf + "Matthiaoso"
    EDITED = True

    DoEvents

    If Mid(RTB1, 1, 1) = vbCr Then RTB1 = Mid(RTB1, 3)

    FF1 = 1
    EDITED = False
    'CLIPBOARD_EDITED = False
End Sub

Private Sub Label_Hosso_Click()

    RTB1.Text = "Hi Liz / Dale / Frank - " + Format(Now, "DD-MMM-YY HH:MMa/p") + vbCrLf + RTB1.Text + vbCrLf + "Regards" + vbCrLf + "Matthiaoso"
    EDITED = True

    DoEvents

    If Mid(RTB1, 1, 1) = vbCr Then RTB1 = Mid(RTB1, 3)

    FF1 = 2
    EDITED = False
    'CLIPBOARD_EDITED = False

End Sub

Private Sub Label_Dad_Click()

    RTB1.Text = "Hi Dad -- " + Format(Now, "DDD DD-MMM-YY HH:MMa/p") + vbCrLf + RTB1.Text + vbCrLf + "Matthiaoso"
    EDITED = True

    DoEvents

    If Mid(RTB1, 1, 1) = vbCr Then RTB1 = Mid(RTB1, 3)

    FF1 = 3
    EDITED = False
    'CLIPBOARD_EDITED = False

End Sub


Private Sub Label_UnKnown_Click()

    RTB1.Text = "Dear Sir -- " + Format(Now, "DDD DD-MMM-YY HH:MMa/p") + vbCrLf + RTB1.Text + vbCrLf + "Yours Sincerely" + vbCrLf + "Matt" + vbCrLf + "A K Respects" + vbCrLf + "Matthiaoso"
    EDITED = True

    DoEvents

    If Mid(RTB1, 1, 1) = vbCr Then RTB1 = Mid(RTB1, 3)

    FF1 = 5
    EDITED = False
    'CLIPBOARD_EDITED = False

End Sub




Private Sub Label2_Click()
'DOWN
If Form2.File1.ListIndex = -1 Then
    Call GRAB_SMS_FILE_CLICK
    Exit Sub
End If

Form2.File1.ListIndex = Form2.File1.ListIndex - 1
XX8 = Form2.File1.List(Form2.File1.ListIndex)
FILESMSGET = Form2.File1.Path + "\" + XX8
Label4 = XX8

Form2.Hide
FR = FreeFile
Open FILESMSGET For Input As #FR
    II = Input(LOF(FR), FR)
Close FR
Form1.RTB1 = II
EDITED = False
CLIPBOARD_EDITED = False

End Sub

Private Sub Label3_Click()

If Form2.File1.ListIndex = -1 Then
    Call GRAB_SMS_FILE_CLICK
    Exit Sub
End If


Form2.File1.ListIndex = Form2.File1.ListIndex + 1
XX8 = Form2.File1.List(Form2.File1.ListIndex)
FILESMSGET = Form2.File1.Path + "\" + XX8
Label4 = XX8


Form2.Hide
FR = FreeFile
Open FILESMSGET For Input As #FR
    II = Input(LOF(FR), FR)
Close FR
Form1.RTB1 = II
EDITED = False
CLIPBOARD_EDITED = False

End Sub

Private Sub Label5_Click()

Source1 = "D:\TEMP\Word.txt"
FileInUseDelay (Source1)
FR1 = FreeFile
Open "D:\TEMP\Word.txt" For Output As #FR1
Print #FR1, RTB1.Text;
Close #FR1

Shell "C:\Program Files\Microsoft Office\OFFICE11\WINWORD.EXE D:\TEMP\Word.txt", vbNormalFocus

Me.WindowState = 1

End Sub

Private Sub Mnu_Blue_Click()

Call Mnu_Fix1st_LettersNOLOW_Click

A1 = App.Path + "\Temp Join__" + GetComputerName + "_.txt"
If FF1 = 1 Then XXI = "MUM": xxi2 = "M"
If FF1 = 2 Then XXI = "HOSSO": xxi2 = "H"
If FF1 = 3 Then XXI = "DAD": xxi2 = "D"
If FF1 = 4 Then XXI = "MUM&GROUP": xxi2 = "MG"
If FF1 = 5 Then XXI = "UNKNOWN": xxi2 = "U"
If FF1 = 7 Then XXI = "EDD": xxi2 = "E"



A2 = sMyDocsFolder + "\03 NotePad Files\00 TOP\01 My Logg.txt"

Do
    DUM = JoinFiles(A1, A2, i)
    
    Sleep 100
    DoEvents
Loop Until DUM = True


Source1 = sMyDocsFolder + "\00 BlueTooth SMS's From CLipType\xBlue" + xxi2 + Format$(Now, "YYMMDDHHMMSS") + " --- " + XXI + ".txt"
FileInUseDelay (Source1)
FR1 = FreeFile
Open Source1 For Output Lock Write As #FR1
    Print #FR1, RTB1.Text;
Close #FR1





'ts = App.Path + "\Send To ScanPath_Merge_Text_Files\## Blue TOTAL.txt"
'ts = App.Path + "\Send To ScanPath_Merge_Text_Files\#0 Scan an Merge Text Files Clip Type.exe"
'Shell ts, vbMinimizedNoFocus

Shell "explorer /select, " + Source1, vbMinimizedNoFocus

'Shell "explorer D:\# MY DOCS\# 01 My Documents\00 BlueTooth SMS's From CLipType", vbNormalFocus

If BLUESEND = True Then BLUESEND = False: Exit Sub

Unload Me


End Sub

Private Sub Mnu_Blue_Fol_Click()

'AUTO FIND BTOOTH FOLDER -- SMS GILLIAN

Shell "explorer D:\# MY DOCS\# 01 My Documents\00 BlueTooth SMS's From CLipType", vbNormalFocus

If Trim(RTB1.Text) = "" Then End


Me.WindowState = 1

End Sub

Private Sub Mnu_BToT_Click()

tS = sMyDocsFolder + "\00 BlueTooth SMS's From CLipType\## Blue TOTAL.txt"
Shell "C:\WINDOZE\system32\notepad.exe " + tS, vbNormalFocus

If Trim(RTB1.Text) = "" Then End

Me.WindowState = 1



End Sub

Private Sub Mnu_Caps_Click()

If FF1 = 0 Then RTB1 = UCase(RTB1)

If FF1 > 0 Then
ts1 = InStr(RTB1, Chr(13))
ts2 = InStrRev(RTB1, Chr(13))
ts2 = InStrRev(RTB1, Chr(13), ts2 - 1)
RTB1 = Mid$(RTB1, 1, ts1) + UCase(Mid$(RTB1, ts1, ts2 - ts1)) + Mid(RTB1, ts2)
End If
End Sub


Private Sub Mnu_CLip_Click()

If CLIPBOARD_EDITED = True Then
    Clipboard.Clear
    Sleep 400
    Clipboard.SetText (RTB1.Text)
    CLIPBOARD_EDITED = False
End If
Me.WindowState = 1

End Sub

Private Sub Mnu_Clue_Click()

Source1 = sMyDocsFolder + "\03 NotePad Files\00 TOP\Clued Up.txt"
FileInUseDelay (Source1)
FR1 = FreeFile
Open sMyDocsFolder + "\03 NotePad Files\00 TOP\Clued Up.txt" For Append Lock Write As #FR1
Print #FR1, "-- "; Format(Now, "DDD DD-MMM-YY HH:MM:SS") + " --"
Print #FR1, RTB1.Text
Print #FR1, "-----"
Close #FR1

'RTB1 = ""


End Sub

Private Sub Mnu_Conflict_Click()

Source1 = sMyDocsFolder + "\03 NotePad Files\00 TOP\Plagiarism & Conflict.txt"
FileInUseDelay (Source1)
Open Source1 For Append Lock Write As #1
Print #1, "-- "; Format(Now, "DDD DD-MMM-YY HH:MM:SS") + " --"
Print #1, RTB1.Text
Print #1, "-----"
Close #1
RTB1 = ""
Unload Me

End Sub

Private Sub Mnu_Dont_Alter_Click()
DontAlter = True
End Sub

Private Sub Mnu_Email_Andy_Click()
On Error Resume Next

Form1.SetFocus
DoEvents

'Exit Sub

MAPISession1.DownLoadMail = False
MAPISession1.SignOn
MAPIMessages1.SessionID = MAPISession1.SessionID
MAPIMessages1.Compose


MAPIMessages1.RecipIndex = 0
MAPIMessages1.RecipDisplayName = "Matthiaoso"
MAPIMessages1.RecipAddress = "Matt.Lan@BTinternet.com"

MAPIMessages1.RecipIndex = 1
MAPIMessages1.RecipDisplayName = "Andy MacMillan"
MAPIMessages1.RecipAddress = "AndyMac11@NTLworld.com"




text5$ = ""
text5$ = text5$ + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p") + vbCrLf
text5$ = text5$ + "--------------------------------------------" + vbCrLf
text5$ = text5$ + RTB1.Text + vbCrLf
text5$ = text5$ + "--------------------------------------------" + vbCrLf

A5$ = sMyDocsFolder + "\01 Email Settings\Signatures\Signature-Nortons2Small.txt"
A5$ = sMyDocsFolder + "\01 Email Settings\Signatures\Signature-Nortons2.txt"

Source1 = A5$
FileInUseDelay (Source1)

fx = FreeFile
    Open A5$ For Input As #fx
    Seek fx, 5
    ttg = Input(LOF(fx) - 5, fx)
    Close #fx

StrConv text5$, vbFromUnicode
StrConv ttg, vbFromUnicode

text5$ = text5$ + ttg + vbCrLf

halo = Mid$(RTB1.Text, 1, 25)
For R = 1 To Len(halo)
    If Mid$(halo, R, 2) = vbCrLf Then Mid$(halo, R, 2) = "  "
Next


MAPIMessages1.MsgSubject = "Msg From Matty Lethal -- " + halo

'MsgBox text5$

'Stop
'End
'StrConv text5$, vbUnicode

MAPIMessages1.MsgReceiptRequested = True

StrConv text5$, vbUnicode
MAPIMessages1.MsgType = "IPM.Microsoft Mail.Note"
'MAPIMessages1.MsgType = "Text/HTML"
MAPIMessages1.MsgNoteText = text5$


MAPIMessages1.Send False
MAPISession1.SignOff
Unload Me

End Sub

Private Sub Mnu_Email_Click()
'Even that lot of wires one john done was a intercept
On Error Resume Next

Form1.SetFocus
DoEvents

'Exit Sub

MAPISession1.DownLoadMail = False
MAPISession1.SignOn
MAPIMessages1.SessionID = MAPISession1.SessionID
MAPIMessages1.Compose

MAPIMessages1.MsgReceiptRequested = True

MAPIMessages1.RecipIndex = 0
MAPIMessages1.RecipDisplayName = "Matt Lan"
MAPIMessages1.RecipAddress = "Matt.Lan@btinternet.com"

'This is Edd's Main One
MAPIMessages1.RecipIndex = 1
MAPIMessages1.RecipDisplayName = "Eddie"
MAPIMessages1.RecipAddress = "Lancaster.E@sky.com"

MAPIMessages1.RecipIndex = 2
MAPIMessages1.RecipDisplayName = "Marianne"
MAPIMessages1.RecipAddress = "Marianne.Farley@btinternet.com"
'Resolve recipient name
'MAPIMessages1.AddressResolveUI = True
'MAPIMessages1.ResolveName
'Create the message

MAPIMessages1.MsgSubject = "IP Mail"

text5$ = ""
text5$ = text5$ + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p") + vbCrLf
text5$ = text5$ + "--------------------------------------------" + vbCrLf
text5$ = text5$ + RTB1.Text + vbCrLf
text5$ = text5$ + "--------------------------------------------" + vbCrLf

A5$ = sMyDocsFolder + "\01 Email Settings\Signatures\Signature-Nortons2Small.txt"
A5$ = sMyDocsFolder + "\01 Email Settings\Signatures\Signature-Nortons2.txt"

Source1 = A5$
FileInUseDelay (Source1)
fx = FreeFile
    Open A5$ For Input As #fx
    Seek fx, 5
    ttg = Input(LOF(fx) - 5, fx)
    Close #fx

StrConv text5$, vbFromUnicode
StrConv ttg, vbFromUnicode

text5$ = text5$ + ttg + vbCrLf

'text5$ = text5$ + "Sweet"


halo = Mid$(RTB1.Text, 1, 25)
For R = 1 To Len(halo)
    If Mid$(halo, R, 2) = vbCrLf Then Mid$(halo, R, 2) = "  "
Next

MAPIMessages1.MsgSubject = "Msg From Matty Lethal -- " + halo

'MsgBox text5$

'Stop
'End
'StrConv text5$, vbUnicode
StrConv text5$, vbUnicode

'MAPIMessages1.MsgType = "Text/HTML"


'text7 = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" + vbCrLf
'text7 = text7 + "<HTML><HEAD>" + vbCrLf
'text7 = text7 + "<META content=3D""text/html; charset=3Diso-8859-1"" =" + vbCrLf
'text7 = text7 + "http-equiv=3DContent-Type>" + vbCrLf
'text7 = text7 + "<META name=3DGENERATOR content=3D""MSHTML 8.00.6001.18783"">" + vbCrLf
'text7 = text7 + "<STYLE></STYLE>" + vbCrLf
'text7 = text7 + "</HEAD>" + vbCrLf
'text7 = text7 + "<BODY bgColor=3D#ffffff>" + vbCrLf
'text7 = text7 + "<DIV><FONT color=3D#000080>" + vbCrLf

MAPIMessages1.MsgType = "IPM.Microsoft Mail.Note"
'MAPIMessages1.MsgType = "Text/HTML"
MAPIMessages1.MsgNoteText = text5

'Add attachment
'MAPIMessages1.AttachmentPathName = "c:\zxcvzxcv.zip"
'Send the message

MAPIMessages1.Send False
MAPISession1.SignOff
Unload Me
End Sub

Private Sub Mnu_Email_Gill_Click()

On Error Resume Next

Form1.SetFocus
DoEvents

'Exit Sub

MAPISession1.DownLoadMail = False
MAPISession1.SignOn
MAPIMessages1.SessionID = MAPISession1.SessionID
MAPIMessages1.Compose


MAPIMessages1.RecipIndex = 0
MAPIMessages1.RecipDisplayName = "Matthiaoso"
MAPIMessages1.RecipAddress = "Matt.Lan@btinternet.com"

MAPIMessages1.RecipIndex = 1
MAPIMessages1.RecipDisplayName = "Gillian"
MAPIMessages1.RecipAddress = "gillian.hyland@ntlworld.com"

MAPIMessages1.RecipIndex = 2
MAPIMessages1.RecipDisplayName = "Lee"
MAPIMessages1.RecipAddress = "lee@kayotix.com"

MAPIMessages1.RecipIndex = 3
MAPIMessages1.RecipDisplayName = "Mark"
MAPIMessages1.RecipAddress = "mark.nick@ntlworld.com"

'Resolve recipient name
'MAPIMessages1.AddressResolveUI = True
'MAPIMessages1.ResolveName
'Create the message



text5$ = ""
text5$ = text5$ + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p") + vbCrLf
text5$ = text5$ + "--------------------------------------------" + vbCrLf
text5$ = text5$ + RTB1.Text + vbCrLf
text5$ = text5$ + "--------------------------------------------" + vbCrLf

A5$ = sMyDocsFolder + "\01 Email Settings\Signatures\Signature-Nortons2Small.txt"
A5$ = sMyDocsFolder + "\01 Email Settings\Signatures\Signature-Nortons2.txt"
Source1 = A5$
FileInUseDelay (Source1)
fx = FreeFile
    Open A5$ For Input As #fx
    Seek fx, 5
    ttg = Input(LOF(fx) - 5, fx)
    Close #fx

StrConv text5$, vbFromUnicode
StrConv ttg, vbFromUnicode

text5$ = text5$ + ttg + vbCrLf
'text5$ = text5$ + "Sweet"

halo = Mid$(RTB1.Text, 1, 25)
For R = 1 To Len(halo)
    If Mid$(halo, R, 2) = vbCrLf Then Mid$(halo, R, 2) = "  "
Next


MAPIMessages1.MsgSubject = "Msg From Matty Lethal -- " + halo

'MsgBox text5$

'Stop
'End
'StrConv text5$, vbUnicode

MAPIMessages1.MsgReceiptRequested = True

StrConv text5$, vbUnicode
MAPIMessages1.MsgType = "IPM.Microsoft Mail.Note"
'MAPIMessages1.MsgType = "Text/HTML"
MAPIMessages1.MsgNoteText = text5$

'MAPIMessages1.body = text5$
'.HTMLBody


'Add attachment
'MAPIMessages1.AttachmentPathName = "c:\zxcvzxcv.zip"
'Send the message

MAPIMessages1.Send False
MAPISession1.SignOff
Unload Me

End Sub

Private Sub Mnu_Email_Judy_Click()

On Error Resume Next

Form1.SetFocus
DoEvents

'Exit Sub

MAPISession1.DownLoadMail = False
MAPISession1.SignOn
MAPIMessages1.SessionID = MAPISession1.SessionID
MAPIMessages1.Compose


MAPIMessages1.RecipIndex = 0
MAPIMessages1.RecipDisplayName = "Matthiaoso"
MAPIMessages1.RecipAddress = "Matt.Lan@btinternet.com"

MAPIMessages1.RecipIndex = 1
MAPIMessages1.RecipDisplayName = "Judy"
MAPIMessages1.RecipAddress = "jalees@easynet.co.uk"

'MAPIMessages1.RecipIndex = 2
'MAPIMessages1.RecipDisplayName = "Marianne"
'MAPIMessages1.RecipAddress = "Marianne.Farley@btinternet.com"
'Resolve recipient name
'MAPIMessages1.AddressResolveUI = True
'MAPIMessages1.ResolveName
'Create the message

MAPIMessages1.MsgSubject = "IP Mail"

text5$ = ""
text5$ = text5$ + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p") + vbCrLf
text5$ = text5$ + "--------------------------------------------" + vbCrLf
text5$ = text5$ + RTB1.Text + vbCrLf
text5$ = text5$ + "--------------------------------------------" + vbCrLf

A5$ = sMyDocsFolder + "\01 Email Settings\Signatures\Signature-Nortons2Small.txt"
A5$ = sMyDocsFolder + "\01 Email Settings\Signatures\Signature-Nortons2.txt"
Source1 = A5$
FileInUseDelay (Source1)
fx = FreeFile
    Open A5$ For Input As #fx
    Seek fx, 5
    ttg = Input(LOF(fx) - 5, fx)
    Close #fx

StrConv text5$, vbFromUnicode
StrConv ttg, vbFromUnicode

text5$ = text5$ + ttg + vbCrLf
'text5$ = text5$ + "Sweet"

halo = Mid$(RTB1.Text, 1, 25)
For R = 1 To Len(halo)
    If Mid$(halo, R, 2) = vbCrLf Then Mid$(halo, R, 2) = "  "
Next


MAPIMessages1.MsgSubject = "Msg From Matty Lethal -- " + halo

'MsgBox text5$

'Stop
'End
'StrConv text5$, vbUnicode

MAPIMessages1.MsgReceiptRequested = True



'MAPIMessages1.AsHTML = bHtml                   ' Optional, default = FALSE, send mail as html or plain text
'MAPIMessages1.Receipt = True

'MAPIMessages1.EncodeType = MIME_ENCODE
'MAPIMessages1.ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
'MAPIMessages1.EncodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
'MAPIMessages1.Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
'MAPIMessages1.Receipt = bReceipt                         ' Optional, default = FALSE

StrConv text5$, vbUnicode

MAPIMessages1.MsgType = "IPM.Microsoft Mail.Note"
'MAPIMessages1.MsgType = "Text/HTML"
MAPIMessages1.MsgNoteText = text5$
'MAPIMessages1.body = text5$
'.HTMLBody


'Add attachment
'MAPIMessages1.AttachmentPathName = "c:\zxcvzxcv.zip"
'Send the message

MAPIMessages1.Send False
MAPISession1.SignOff
Unload Me

End Sub

Private Sub Mnu_Email_Val_Click()
On Error Resume Next

Form1.SetFocus
DoEvents

'Exit Sub

MAPISession1.DownLoadMail = False
MAPISession1.SignOn
MAPIMessages1.SessionID = MAPISession1.SessionID
MAPIMessages1.Compose


MAPIMessages1.RecipIndex = 0
MAPIMessages1.RecipDisplayName = "Matthiaoso"
MAPIMessages1.RecipAddress = "Matt.Lan@btinternet.com"

MAPIMessages1.RecipIndex = 1
MAPIMessages1.RecipDisplayName = "Valerie de Schaller"
MAPIMessages1.RecipAddress = "ValAndFree@yahoo.com"

'MAPIMessages1.RecipIndex = 2
'MAPIMessages1.RecipDisplayName = "Marianne"
'MAPIMessages1.RecipAddress = "Marianne.Farley@btinternet.com"
'Resolve recipient name
'MAPIMessages1.AddressResolveUI = True
'MAPIMessages1.ResolveName
'Create the message

MAPIMessages1.MsgSubject = "IP Mail"

text5$ = ""
text5$ = text5$ + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p") + vbCrLf
text5$ = text5$ + "--------------------------------------------" + vbCrLf
text5$ = text5$ + RTB1.Text + vbCrLf
text5$ = text5$ + "--------------------------------------------" + vbCrLf

A5$ = sMyDocsFolder + "\01 Email Settings\Signatures\Signature-Nortons2Small.txt"
A5$ = sMyDocsFolder + "\01 Email Settings\Signatures\Signature-Nortons2.txt"
Source1 = A5$
FileInUseDelay (Source1)
fx = FreeFile
    Open A5$ For Input As #fx
    Seek fx, 5
    ttg = Input(LOF(fx) - 5, fx)
    Close #fx

StrConv text5$, vbFromUnicode
StrConv ttg, vbFromUnicode

text5$ = text5$ + ttg + vbCrLf
'text5$ = text5$ + "Sweet"


halo = Mid$(RTB1.Text, 1, 25)
For R = 1 To Len(halo)
    If Mid$(halo, R, 2) = vbCrLf Then Mid$(halo, R, 2) = "  "
Next

MAPIMessages1.MsgSubject = "Msg From Matty Lethal -- " + halo

'MsgBox text5$

'Stop
'End
'StrConv text5$, vbUnicode

MAPIMessages1.MsgReceiptRequested = True



'MAPIMessages1.AsHTML = bHtml                   ' Optional, default = FALSE, send mail as html or plain text
'MAPIMessages1.Receipt = True

'MAPIMessages1.EncodeType = MIME_ENCODE
'MAPIMessages1.ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
'MAPIMessages1.EncodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
'MAPIMessages1.Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
'MAPIMessages1.Receipt = bReceipt                         ' Optional, default = FALSE

StrConv text5$, vbUnicode
MAPIMessages1.MsgType = "IPM.Microsoft Mail.Note"
'MAPIMessages1.MsgType = "Text/HTML"
MAPIMessages1.MsgNoteText = text5$
'MAPIMessages1.body = text5$
'.HTMLBody


'Add attachment
'MAPIMessages1.AttachmentPathName = "c:\zxcvzxcv.zip"
'Send the message

MAPIMessages1.Send False
MAPISession1.SignOff
Unload Me
End Sub




Private Sub Mnu_Fix1st_Letters_Click()

    
Dim EE As String

EE = RTB1.Text
   
If EGG = 0 Then EE = LCase(EE)


Do
    tt2 = InStr(tt2 + 1, EE, " ")
    If tt2 > 0 Then Mid$(EE, tt2 + 1, 1) = UCase(Mid$(EE, tt2 + 1, 1))
    tt3 = InStr(tt3 + 1, EE, vbCrLf)
    If tt3 > 0 Then Mid$(EE, tt2 + 1, 1) = UCase(Mid$(EE, tt2 + 1, 1))
Loop Until tt2 = 0

tt2 = 0
Do
    tt2 = InStr(tt2 + 1, EE, vbLf)
    If tt2 > 0 And tt2 + 1 <= Len(EE) Then Mid$(EE, tt2 + 1, 1) = UCase(Mid$(EE, tt2 + 1, 1))
Loop Until tt2 = 0

tt2 = 0
Do
    tt2 = InStr(tt2 + 1, EE, "-")
    If tt2 > 0 And tt2 + 1 <= Len(EE) Then Mid$(EE, tt2 + 1, 1) = UCase(Mid$(EE, tt2 + 1, 1))
Loop Until tt2 = 0
tt2 = 0
Do
    tt2 = InStr(tt2 + 1, EE, "(")
    If tt2 > 0 And tt2 + 1 <= Len(EE) Then Mid$(EE, tt2 + 1, 1) = UCase(Mid$(EE, tt2 + 1, 1))
Loop Until tt2 = 0


Do
    DoEvents
    vf = InStr(EE, " i ")
    If vf = 0 Then vf = InStr(EE, vbLf + "i ")
    If vf = 0 Then vf = InStr(EE, vbCr + "i ")
    If vf > 0 Then
        Mid$(EE, vf + 1, 1) = "I"
    End If
Loop Until vf = 0 Or vf >= Len(EE)


RTB1.Text = EE
EDITED = True

End Sub





Private Sub Mnu_Fix1st_LettersNOLOW_Click()

EGG = 1
Call Mnu_Fix1st_Letters_Click
EGG = 0
End Sub

Private Sub Mnu_FixCaps_Click()

'Me.WindowState = 1

Dim VF4, TY
Dim TT As String

TT = RTB1.Text

TT = " " + LCase(TT)


    'Mid$(TT, 2, 1) = UCase(Mid$(TT, 2, 1))
    vf = 0
    Do
        DoEvents
            
           VFX = vf
            vf = InStr(vf + 1, TT, vbLf)
            If vf = 0 Then
            If VFX >= Len(TT) Then Exit Do
            vf = Len(TT) + 2
            End If
                        
            
            YY5 = Mid$(TT, VFX + 1, (vf - VFX) - 2)
            
            
'            If InStr(YY5, "vf2") > 0 Then Stop
            
            '1ST ONE DONE
            For R = 1 To Len(YY5)
                i = Mid(YY5, R, 1)
                Mid(YY5, R, 1) = UCase(Mid(YY5, R, 1)): VF7 = R
                If Mid(YY5, R, 1) <> i Then
                    Exit For
                End If
            Next
            
            
            '+ Chr(9) + vbLf
            
            
            
            CHKA = ".,;:?(!-<""[{~¦+"
            For R1 = 1 To Len(CHKA)
                GOMORE = 0
                TY = Mid$(CHKA, R1, 1)
                VF4 = InStr(YY5, TY)
'                If TY = ";" And VF4 > 0 Then Stop
                If VF4 > 0 Then
                    For R2 = VF4 To Len(YY5)
                        i = Mid(YY5, R2, 1)
                        If IsNumeric(i) Then
                            Exit For
                        End If
                        Mid(YY5, R2, 1) = UCase(Mid(YY5, R2, 1))
                        If Mid(YY5, R2, 1) <> i Then
                            GOMORE = 1
                            Exit For
                        End If
                    Next
                End If
            Next
        
            Mid(TT, VFX + 1, Len(YY5)) = YY5 '+ vbCrLf
        
            'VF = VFX + 1
    Loop Until vf = 0 Or vf - 1 > Len(TT)


    'FIX WHAT HAPPEN_S AFTER A CHAR EVENT FOUND
    '-------------------------------------------------------------------------
    Dim XPOINTER2, XPOINTER3
    XPOINTER = 1
    CHKA = ".,;:?(!-<""[{~¦+"
    ReDim COUNTIP(Len(CHKA))
    Do
        For R1 = 1 To Len(CHKA)
            TY = Mid$(CHKA, R1, 1)
            XPOINTER4 = InStr(XPOINTER, TT, TY)
            COUNTIP(R1) = XPOINTER4
        Next
        'WE ARE CHECK FOR A FEW SYMBOL AND EACH EVENT HAS RIASED OR LOWER XPOINTER
        'SO WE WANTING THE LOWEST OF CHECK TO BACK CHECK UNTIL DONE
        XPOINTER3 = Len(TT)
        XPOINTER2 = 0
        For R1 = 1 To Len(CHKA)
            If COUNTIP(R1) > 0 Then
                If COUNTIP(R1) < XPOINTER3 Then
                    XPOINTER2 = COUNTIP(R1)
                    XPOINTER3 = XPOINTER2
                End If
            End If
        Next
        
        If XPOINTER2 = 0 Then Exit Do
            
        For R2 = XPOINTER2 To Len(TT)
            CHRX = UCase(Mid(TT, R2, 1))
            If Asc(CHRX) >= 65 And Asc(CHRX) <= 65 + 25 Then
                Mid(TT, R2, 1) = CHRX
                Exit For
            End If
        Next
        XPOINTER = XPOINTER2 + 1
    Loop Until XPOINTER2 = 0
    
    
    
'    'FIX WHAT HAPPEN_S AFTER A STRING EVENT FOUND
'    '----------------------------------------------------------------------------
'    Dim XPOINTER2, XPOINTER3
'    XPOINTER = 1
'
'    Dim REPLACESTRING(2)
'    REPLACESTRING(1) = "Hi "
'    REPLACESTRING(2) = "Hi There "
'
'    ReDim COUNTIP(2)
'    COUNTIPX = 0
'    Do
'        For R1 = 1 To 2
'            TY = REPLACESTRING(R1)
'            XPOINTER4 = InStr(XPOINTER, TT, TY)
'            COUNTIP(R1) = XPOINTER4
'        Next
'        'WE ARE CHECK FOR A FEW SYMBOL AND EACH EVENT HAS RIASED OR LOWER XPOINTER
'        'SO WE WANTING THE LOWEST OF CHECK TO BACK CHECK UNTIL DONE
'        XPOINTER3 = Len(TT)
'        XPOINTER2 = 0
'        For R1 = 1 To Len(CHKA)
'            If COUNTIP(R1) > 0 Then
'                If COUNTIP(R1) < XPOINTER3 Then
'                    XPOINTER2 = COUNTIP(R1)
'                    XPOINTER3 = XPOINTER2
'                End If
'            End If
'        Next
'
'        If XPOINTER2 = 0 Then Exit Do
'
'        For R2 = XPOINTER2 To Len(TT)
'            CHRX = UCase(Mid(TT, R2, 1))
'            If Asc(CHRX) >= 65 And Asc(CHRX) <= 65 + 25 Then
'                Mid(TT, R2, 1) = CHRX
'                Exit For
'            End If
'        Next
'        XPOINTER = XPOINTER2 + 1
'    Loop Until XPOINTER2 = 0
    
    
'    TTL = LCase(TT)
'    Dim REPLACESTRING(2)
'    REPLACESTRING(1) = "Hi There "
'    REPLACESTRING(2) = "Hi "
'    For R1 = 1 To 2
'    For R2 = 65 To 65 + 25
'    TT = Replace(LCase(TT), LCase(REPLACESTRING(R1) + Chr(R2)), REPLACESTRING(R1) + Chr(R2))
'    Next
'    Next
    
    
    

Do
    DoEvents
    vf = InStr(TT, " i ")
    If vf = 0 Then vf = InStr(TT, vbLf + "i ")
    If vf = 0 Then vf = InStr(TT, vbCr + "i ")
    If vf > 0 Then
        Mid$(TT, vf + 1, 1) = "I"
    End If
Loop Until vf = 0 Or vf >= Len(TT)


TT = Mid(TT, 2)

RTB1.Text = TT
EDITED = True


End Sub

Private Sub GRAB_SMS_FILE_CLICK()
Form2.File1.Path = "D:\# MY DOCS\# 01 My Documents\00 BlueTooth SMS's From CLipType"

Form2.Show
Form2.SetFocus

AlwaysOnTop (Form2.hWnd)


End Sub

Private Sub MNU_INFO_RAPID_Click()
'D:\VB6\VB-NT\00_Best_VB_01\Clip_Type\TEXTBOX

II = " /nologo /DirD:\VB6\VB-NT\00_Best_VB_01\Clip_Type\TEXTBOX\"
Shell "C:\Program Files\seRapid\seRapid.exe " + II, vbNormalFocus
Me.WindowState = 1

'END

End Sub

Private Sub Mnu_Job_Click()

A1 = App.Path + "\Temp Join__" + GetComputerName + "_.txt"
A2 = sMyDocsFolder + "\03 NotePad Files\00 TOP\Jobs.txt"

DUM = JoinFiles(A1, A2, i)
If DUM = True Then Unload Me

End Sub

Private Sub Mnu_KadMan_Click()

A1 = App.Path + "\Temp Join__" + GetComputerName + "_.txt"
A2 = sMyDocsFolder + "\03 NotePad Files\01 TOP\Kadaitcha Man Techy NFO.txt"

DUM = JoinFiles(A1, A2, i)
If DUM = True Then Unload Me

End Sub

Private Sub Mnu_Logg_Click()

'always goes to clue already end prog
'Call Mnu_Clue_Click

A1 = App.Path + "\Temp Join__" + GetComputerName + "_.txt"
A2 = sMyDocsFolder + "\03 NotePad Files\00 TOP\01 My Logg.txt"

DUM = JoinFiles(A1, A2, i)
If DUM = True Then

If UNLOAD_BEGUN = True Then Exit Sub

Unload Me

End If

End Sub

Private Sub MNU_LS_LAST_LOAD_Click()

FILESAVENAME1_SAME_DONE_READY = True

    If FILESAVENAME1_SAME = "" Then
        If FS.FolderExists(App.Path + "\TEXTBOX") = False Then
            MkDir App.Path + "\TEXTBOX"
        End If
        FILESAVENAME1_SAME = App.Path + "\TEXTBOX\TEXTBOX TEXT SAVE_SAME.txt"
    End If

If Dir(FILESAVENAME1_SAME) = "" Then
    Beep
    MsgBox "NOT SUCH A FILE HERE YET _ MUST SAVE ONE" + vbCrLf + vbCrLf + FILESAVENAME1_SAME
    Exit Sub
End If

FR = FreeFile
Open FILESAVENAME1_SAME For Input As #FR
    II = Input(LOF(FR), FR)
Close FR
Form1.RTB1 = II
EDITED = False
CLIPBOARD_EDITED = False

End Sub

Private Sub MNU_LS_SAVE_LAST_Click()

FILESAVENAME1_SAME_DONE_READY = True
Call ENDSAVE_SAME

End Sub

Private Sub MNU_MY_LOVELY_WOMAN_Click()



A1 = App.Path + "\Temp Join__" + GetComputerName + "_.txt"
A2 = sMyDocsFolder + "\03 NotePad Files\00 TOP\My Woman.txt"

DUM = JoinFiles(A1, A2, i)
If DUM = True Then Unload Me



Exit Sub



Source1 = sMyDocsFolder + "\03 NotePad Files\00 TOP\My Woman.txt"
FileInUseDelay (Source1)
FR1 = FreeFile
Open Source1 For Append Lock Write As #FR1
Print #FR1, "-- "; Format(Now, "DDD DD-MMM-YY HH:MM:SS") + " --"
Print #FR1, RTB1.Text
Print #FR1, "-----"
Close #FR1
Unload Me
End Sub

Private Sub MNU_NOSAVEEXIT_Click()


End
End Sub

Private Sub MNU_NOTEPAD_Click()

Shell "NOTEPAD.EXE " + FILESAVENAME1, vbNormalFocus
Me.WindowState = 1


End Sub

Private Sub Mnu_Prog_Click()

A1 = App.Path + "\Temp Join__" + GetComputerName + "_.txt"
A2 = sMyDocsFolder + "\03 NotePad Files\VB Code\Programming Snippets.txt"

DUM = JoinFiles(A1, A2, i)
If DUM = True Then Unload Me

End Sub


Function Check1stDate(Source1) As Boolean
    Buffer1 = " --" + vbCrLf + vbCrLf + vbCrLf + "-----" + vbCrLf
    Buffer2 = Buffert1
    Open Source1 For Binary Access Read As #1
    Get #1, 26, Buffer1
    Close #1
    If Buffer1 = Buffer2 Then Check1stDate = True
End Function

Private Sub MNU_SAVED_Click()
Call Timer_SAVED_Timer
End Sub

Private Sub Mnu_SendBT_W5_Click()
    
    BLUESEND = True
    Call Mnu_Blue_Click
    Shell "C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=6c:0e:0d:30:fc:86 -BD_NAME=W50x2D07919180203 -DEV_CLASS=00262746" + " """ + Source1 + """"
    BlueDontSaveOnExit = True

End Sub
Private Sub Mnu_SendBT_W880I_Click()
    
    BLUESEND = True
    Call Mnu_Blue_Click
    Shell "C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=6c:0e:0d:30:fc:86 -BD_NAME=W50x2D07919180203 -DEV_CLASS=00262746" + " """ + Source1 + """"
    BlueDontSaveOnExit = True

End Sub
Private Sub Mnu_SendBT_K8_Click()

    BLUESEND = True
    Call Mnu_Blue_Click
    Shell "C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=00:1d:28:2b:90:9c -BD_NAME=K80x2D07919180203 -DEV_CLASS=00262746" + " """ + Source1 + """"
    BlueDontSaveOnExit = True

End Sub

Private Sub Mnu_SendBT_W100i_Click()

    BLUESEND = True
    Call Mnu_Blue_Click
    Shell "C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=6c:0e:0d:66:d0:18 -BD_NAME=W100i -DEV_CLASS=00262746" + " """ + Source1 + """"
    BlueDontSaveOnExit = True

End Sub




Private Sub Mnu_UNET_Click()

A1 = App.Path + "\Temp Join__" + GetComputerName + "_.txt"
A2 = sMyDocsFolder + "\03 NotePad Files\Word Play\#0-Unet Notes.txt"

DUM = JoinFiles(A1, A2, i)
If DUM = True Then Unload Me

End Sub


Sub ENDSAVE_SAME()

On Error Resume Next



'REQUEST A SAVE ONLY AFTER A LOAD
'-------------------------------------------------------
If FILESAVENAME1_SAME_DONE_READY = False Then Exit Sub
'-------------------------------------------------------

    If FILESAVENAME1_SAME = "" Then
        If FS.FolderExists(App.Path + "\TEXTBOX") = False Then
            MkDir App.Path + "\TEXTBOX"
        End If
        FILESAVENAME1_SAME = App.Path + "\TEXTBOX\TEXTBOX TEXT SAVE_SAME.txt"
    End If
    
    Err.Clear
    FR = FreeFile
    Open FILESAVENAME1_SAME For Output As #FR
        Print #FR, RTB1.Text
    Close #FR
    
    If Err.Number > 0 Then
        MsgBox "Error Save Try Agian Press"
        Sleep 1000
    End If

End Sub


Sub ENDSAVE()

Call ENDSAVE_SAME

On Error Resume Next

Do
    
    If FILESAVENAME1 = "" Then
        If FS.FolderExists(App.Path + "\TEXTBOX") = False Then
            MkDir App.Path + "\TEXTBOX"
        End If
        FILESAVENAME1 = App.Path + "\TEXTBOX\TEXTBOX TEXT SAVE " + Format$(Now, "YYYY-MM-DD HH-MM-SS - DDD") + ".txt"
    End If
    
    If FILESAVENAME2 = "" Then
        If FS.FolderExists(App.Path + "\TEXTBOX BIGGER") = False Then
            MkDir App.Path + "\TEXTBOX BIGGER"
        
        End If
        FILESAVENAME2 = App.Path + "\TEXTBOX BIGGER\TEXTBOX TEXT SAVE -BIGGER-" + Format$(Now, "YYYY-MM-DD HH-MM-SS - DDD") + ".txt"
    End If
        
    Err.Clear
    FR = FreeFile
    Open FILESAVENAME1 For Output As #FR
        Print #FR, RTB1.Text
    Close #FR

    
    LENL = Len(RTB1.Text)
    If OLENL < LENL Then
        FR = FreeFile
        Open FILESAVENAME2 For Output As #FR
            Print #FR, RTB1.Text
        Close #FR
    End If
    
    If Err.Number > 0 Then
        MsgBox "Error Save Try Again Press"
        Sleep 1000
    End If

Loop Until Err.Number = 0
    
End Sub

Private Sub Mnu_VB_Click()

If CLIPBOARD_EDITED = True Then
    Clipboard.Clear
    Sleep 400
    Clipboard.SetText (RTB1.Text)
    CLIPBOARD_EDITED = False
End If


Call ENDSAVE


Call Mnu_VB_ME_Click

Unload Me

End

End Sub

Private Sub Mnu_VB_ME_Click()

VBPATH = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
If Dir(VBPATH) = "" Then
    VBPATH = "C:\Program Files (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
End If

Shell VBPATH + " """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
End

End Sub
'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)

    Dim FileSpec, FILE_SPEC_GO, i
    
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".vbp"
    i = CODER_VBP_FILE_NAME_2
    CODER_EXE_FILE_NAME_1 = i
    If InStr(i, " - 64 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(i, " - 64 BIT.vbp", ".vbp")
    If InStr(i, " - 32 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(i, " - 32 BIT.vbp", ".vbp")
    
    CODER_EXE_FILE_NAME_1 = Replace(CODER_EXE_FILE_NAME_1, ".vbp", ".exe")
    
    FileSpec = CODER_VBP_FILE_NAME_2
    If Dir(FileSpec) <> "" Then
        FILE_SPEC_GO = FileSpec
    Else
        If LCase(GetSpecialfolder(&H29)) = LCase("C:\Windows\SysWOW64") Then
            FileSpec = Replace(FileSpec, ".vbp", " - 64 BIT.vbp")
        Else
            FileSpec = Replace(FileSpec, ".vbp", " - 32 BIT.vbp")
        End If
    End If
    
    If Dir(FileSpec) <> "" Then FILE_SPEC_GO = FileSpec
    
    CODER_VBP_FILE_NAME_2 = FILE_SPEC_GO
End Sub


Private Sub MNU_VBFOLDER_Click()

Shell "explorer /select, " + FILESAVENAME1 + "\", vbNormalFocus
Me.WindowState = 1
End Sub

Private Sub Mum_G_Click()
    
    RTB1.Text = "Hi Mum & Group -- " + Format(Now, "DDD DD-MMM-YY HH:MMa/p") + vbCrLf + RTB1.Text + vbCrLf + "A K Respects" + vbCrLf + "Matthiaoso"
    EDITED = True

    DoEvents

    If Mid(RTB1, 1, 1) = vbCr Then RTB1 = Mid(RTB1, 3)

    FF1 = 4

End Sub

Private Sub MUN_EXIT_UNLOAD_Click()

Unload Me

End Sub

Private Sub RTB1_Change()

EDITED = True
EDITED2 = True
CLIPBOARD_EDITED = True

'If OTEXT <> RTB1.Text Then
'    OTEXT = RTB1.Text
    'If EDITED3 > Now Then
    'End If
    
'End If

LL = Len(RTB1.Text) - 400
If LL < 1 Then LL = 1

If InStr(RTB1.Text, "`") Then
    XZ = RTB1.SelStart
    'TCH = StrConv("`", vbUnicode)
    RTB1.Text = Replace(RTB1.Text, "`", "") 'SHOULD DO ONLY END
    RTB1.SelStart = XZ
    EDITED = True

End If
        
        
WINAMPEDIT = Now + TimeSerial(0, 0, 5)
        
        
        
EDITED3 = Now + TimeSerial(0, 0, 20)



If EDITED3 = -1 Then EDITED3 = 0

LOWERTIMER = Now + TimeSerial(0, 1, 0)

'If Mid(RTB1, 1, 1) = vbCr Then Stop: Exit Sub

BlueDontSaveOnExit = False

Label1 = Len(RTB1)
If SMSLEN - Len(RTB1) > 0 Then
    Label9 = Str(SMSLEN - Len(RTB1))
Else
    Label9 = 0
End If

gg = Len(RTB1)
If OGG > gg + 2 Or OGG < gg - 2 Then

ClipedText = True
End If
OGG = gg

'5 SMS
Label1.BackColor = 0

SMSLEN = 763

If Len(RTB1) > SMSLEN And FF1 > 0 Then

Label1.BackColor = QBColor(12)

'10 SMS
'If Len(RTB1) > 1530 And FF1 >0 Then

'Beep
    
    MMControl1.Command = "Open"
    MMControl1.Command = "Play"
        Call WaitWavFinish
    MMControl1.Command = "Close"



End If

If RTB1.SelStart > 1 Then

If (Mid$(RTB1, 1, 1) <> UCase(Mid$(RTB1, 1, 1))) And Len(RTB1) < 2 Then
    RTB1 = UCase(Mid$(RTB1, 1, 1)) + Mid$(RTB1, 2)
    RTB1.SelStart = 1
    RTB1.SelLength = Len(RTB1.Text)
End If
End If

'If Len(RTB1) = 1 Then
'RTB1 = Format(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + "-- Start Text --" + vbCrLf + RTB1
'RTB1.SelStart = Len(RTB1)
'End If

End Sub

Private Sub RTB1_KeyDown(KeyCode As Integer, Shift As Integer)

EDITED3 = Now + TimeSerial(0, 0, 20)

If KeyCode = 223 Then KeyCode = 0: Exit Sub
If KeyCode = 27 And IsIDE = True Then Unload Me: Exit Sub




Label6 = Trim(Str(RTB1.SelStart))

If KeyCode = 112 Then Unload Form1
'If KeyCode = 223 Then Me.WindowState = 1
If KeyCode = 17 Then Exit Sub


KeyCode = UCase(KeyCode)
If KeyCode = Asc("C") And Shift = 2 Then Call Mnu_Clue_Click
If KeyCode = Asc("L") And Shift = 2 Then Call Mnu_Logg_Click
If KeyCode = Asc("J") And Shift = 2 Then Call Mnu_Job_Click
If KeyCode = Asc("P") And Shift = 2 Then Call Mnu_Prog_Click
If KeyCode = Asc("U") And Shift = 2 Then Call Mnu_UNET_Click
If KeyCode = Asc("K") And Shift = 2 Then Call Mnu_KadMan_Click
If KeyCode = Asc("O") And Shift = 2 Then Call Mnu_Conflict_Click
If KeyCode = Asc("1") And Shift = 2 Then Call Mnu_Fix1st_LettersNOLOW_Click


'If Shift = 2 Then Stop
If FF1 >= 1 Then Exit Sub


If KeyCode = 77 And Shift = 2 Then FF1 = 1: Call Label_Mum_Click
'If KeyCode = 72 And Shift = 2 Then FF1 = 2: Call Label_Hosso_Click
'If KeyCode = 68 And Shift = 2 Then FF1 = 3: Call Label_Dad_Click
'If KeyCode = 68 And Shift = 2 Then FF1 = 7: Call Label_Eddie_Click
'If KeyCode = 68 And Shift = 2 Then FF1 = 4: Call Label_MumG_Click
'If KeyCode = 68 And Shift = 2 Then FF1 = 5: Call Label_UnKnown_Click

If KeyCode = 87 And Shift = 2 Then Call Label5_Click   'WORD



'NICE SAYING IN MY JOBS TODAY
'CANNOT CONTINUE UNTIL WE REDUCE EVERY POSIBILITY OF RISK THAT HAS A POISIBILITY OF OCCURING


End Sub


' #VBIDEUtils#************************************************
' * Programmer Name  : Waty Thierry
' * Web Site         : www.geocities.com/ResearchTriangle/6311/
' * E-Mail           : waty.thierry@usa.net
' * Date             : 14/06/99
' * Time             : 11:54
' * Comments         : Join two files together
' *
' *****************************************

Public Function JoinFiles(Source1 As String, Source2 As String, i) As Boolean
    'Join 1 in front of 2
   
   On Error GoTo ErrorHandler
    
    i = Check1stDate(Source1)
    Open A1 For Output Lock Write As #1
        Print #1, "-- "; Format(Now, "DDD DD-MMM-YY HH:MM:SS") + " --"
        Print #1, RTB1.Text
        Print #1,
        Print #1, "-----"
    Close #1
    
    FileInUseDelay (Source1)
    Open Source1 For Binary Access Read As #1
    FileInUseDelay (Source2)
    Open Source2 For Binary Access Read As #2
   
    If i = True Then offs = 41 Else offs = 1
   
   Buffer1 = Space(LOF(1))
   Get #1, , Buffer1
   
   Buffer2 = Space(LOF(2) - (offs - 1))
   Get #2, offs, Buffer2
   
   Close #1, #2
   
    FileInUseDelay (Source2)
    Kill Source2
    FileInUseDelay (Source2)
    Open Source2 For Binary Access Write As #2
    Put #2, , Buffer1
    Put #2, , Buffer2
    Close #2
   
JoinFiles = True

'RTB1 = ""

Exit Function
ErrorHandler:
    
    MsgBox "Error Join File TRY AGAIN" + vbCrLf + Source1 + vbCrLf + Source2
    
    Close #1
    Close #2
    Close #3

JoinFiles = False

End Function

Public Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Public Function NotAlwaysOnTop(ByVal hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function

Function FileInUse(ByVal strFilePath As String) As Boolean
  
  On Error Resume Next
  
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
        II = FileInUse(Txw$)
        If II = True Then Sleep (200)
    Loop Until II = False Or QQ < Now
    
    If II = True Then
        MsgBox "Trouble File Stuck In Use " + vbCrLf + Txw$ + vbCrLf + "Longer than 1 min 30 secs"
        Stop
    End If
End Sub


Private Sub WaitWavFinish()
    Do
        DoEvents
    Loop Until MMControl1.Mode = 525
End Sub

Private Sub RTB1_KeyUp(KeyCode As Integer, Shift As Integer)

Label6 = Trim(Str(RTB1.SelStart))

End Sub


Private Sub Timer1_Timer()
    
    Exit Sub
    
    Label7 = Format(Now, "ddd dd-mmm-yy hh:mm:ssa/p")

        TF2 = FindWindow("Winamp v1.x", vbNullString)
            
        MR = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)
        If OLDMR <> MR Then
        XZ = RTB1.SelStart
        
        
        If MR = 1 Then
        'RTB1.Text = RTB1.Text + LF + "<<  VIDEO PLAY >>" + vbCrLf
        End If
        If MR <> 1 Then
        'RTB1.Text = RTB1.Text + LF + "<< VIDEO STOP >>" + vbCrLf
        
        End If
        
        RTB1.SelStart = Len(RTB1.Text)

        
        End If
        OLDMR = MR


End Sub

Private Sub Timer2_Timer()

'If Me.WindowState = 1 Then WINAMPEDIT = 1
'
'    TF2 = FindWindow("Winamp v1.x", vbNullString)
'
'    If WINAMPEDIT > Now Then
'        messageResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)
'        If messageResult <> 1 Then
'            'MsgResult2 = SendMessage(TF2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
'        End If
'
'    Else
'
'    If WINAMPEDIT < Now Or WINAMPEDIT = 0 Then
'        messageResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)
'        If messageResult = 1 Then
'            'MsgResult2 = SendMessage(TF2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
'        End If
'        WINAMPEDIT = 0
'    End If
'
'End If


                
                

If Me.WindowState = 1 Then EDITED3 = 1



'-----------------------------------------------------------------------
'GOT SOME QUICKER SAVE THAN HERE GOING ON
'JUST AN EXTRA MESSURE CODER TESTER
'CHECK EXIT SAVE ALSO
'-----------------------------------------------------------------------
If EDITED = True And EDITED3 <= 0 Then EDITED3 = Now + TimeSerial(0, 0, 8)
'-----------------------------------------------------------------------

'  <  NOW
If EDITED3 < Now And EDITED3 > 0 Then
    
    LF = vbCrLf
    
    If Mid(RTB1.Text, Len(RTB1.Text) - 2) = vbCrLf Then
        LF = ""
    End If
    
    If InStrRev(RTB1.Text, " >>" + vbCrLf) > Len(RTB1.Text) - 5 = False Then
        XZ = RTB1.SelStart
        RTB1.Text = RTB1.Text + LF + "<< " + Format$(Now, "DD-MMM-YY -- HH:MM:SSa/p") + " >>" + vbCrLf
        RTB1.SelStart = Len(RTB1.Text)
        EDITED = True

    End If
    
    RTB1.SelStart = Len(RTB1.Text)
    
    OTEXT = RTB1.Text
    EDITED3 = -1
    
End If

If LOWERTIMER < Now And LOWERTIMER > 0 Then
    LOWERTIMER = 0
    'Me.WindowState = 1
End If

If EDITED2 = True Or EDITED = True Then
    EDITED2 = False

    Call ENDSAVE
    OTEXT = RTB1.Text

End If

End Sub

Private Sub Timer3_Timer()




KKEY = GetAsyncKeyState(223)
If Me.WindowState = 0 And KKEY < 0 Then
    If GetForegroundWindow <> Me.hWnd Then
        Me.SetFocus
        TU = Now + Timer + 0.2
        Exit Sub
    End If
End If


If KKEY < 0 And TU < Now + Timer Then
    Me.SetFocus
    TU = Now + Timer + 0.2
    
    If WINSTATETX = 1 Then Me.WindowState = 0: Call VIDEO: DoEvents: Me.SetFocus: Exit Sub

    
    
    If WINSTATETX = 0 Then Me.WindowState = 1: Call VIDEO
End If


End Sub



Sub VIDEO()

Exit Sub
        'vbNullString = EMPTY NULL VOID NOTHING
        X = FindWindow("Winamp Video", vbNullString)
'        XVID = Not XVID
'        If XVID = 0 Then
            
        If IsWindowVisible(X) = True Then
            ShowWindow X, SW_SHOWNORMAL
           Else
            ShowWindow X, SW_MINIMIZE
        
            TF2 = FindWindow("Winamp v1.x", vbNullString)
            messageResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)
            If messageResult = 1 Then
                'MsgResult2 = SendMessage(X, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
            End If
        End If
        
        'SetForegroundWindow X
        'lngI = SetFocuses(PopBack)
        'ingl = Putfocus(PopBack)
            
        

End Sub

'-------------------------------------------------
'-------------------------------------------------
Private Sub Timer_1_MINUTE_Timer()

Timer_1_MINUTE.Interval = 60000

Call VB_PROJECT_CHECKDATE

End Sub
'-------------------------------------------------
'-------------------------------------------------
Sub VB_PROJECT_CHECKDATE()

'TOP DELCLARE -------------
'DIM XVB_DATE
'--------------------------

If Mid(App.Path, 1, 1) <> "D" Then Exit Sub

PATH_FILE_NAME1 = App.Path + "\" + App.EXEName + ".EXE"
PATH_FILE_NAME2 = Replace(PATH_FILE_NAME1, "D:\VB6\", "D:\VB6-EXE'S\")

'---------------------------------------------------
'IF A NEW PROJECT NOT BEEN SYNC YET TO MIRROR FOLDER
'---------------------------------------------------
'MINOR WORK AROUND CREATE PATH
'---------------------------------------------------
If Dir(PATH_FILE_NAME2) = "" Then
    On Error Resume Next
        MkDir Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\"))
        FS.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
    If Err.Number > 0 Then Exit Sub
End If

Set F = FS.GetFile(PATH_FILE_NAME2)

VB_DATE = F.DateLastModified

'--------
'01 OF 02
'----------------------------------
'WRITE INFO ABOUT DATE CHANGE NEWER
'----------------------------------

If XVB_DATE < VB_DATE And XVB_DATE > 0 Then
    
    '-----------------------------------
    'RUN A VBS TO COPY OVER AND RELOADER
    '-----------------------------------
    
'    FR1 = FreeFile
'    Open App.Path + "\RELOAD AND COPY.VBS" For Output As #FR1


    '----------------------------------
    'Mon 10 April 2017 16:30:50--------
    '----------------------------------
    'PROJECT REFERANCE ----------------
    'wshom.ocx
    'IF HARD FIND DO BROWSER
    'IN SYSTEM32 AND SYSWOW64
    'DIR wshom.ocx /S
    '----------------------------------
    'WINDOWS SCRIPT HOST OBJECT MODEL
    'AFTER FIND ALSO HAVE TO SELECT HER
    '----------------------------------
    '----------------------------------
    'Mon 10 April 2017 15:45:11--------
    '----
    'VISUAL BASIC wshom.ocx NOT WORKING - Google Search
    'https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB721GB721&q=VISUAL+BASIC+wshom.ocx+NOT+WORKING&oq=VIUAL+BASIC+wshom.ocx+NOT+WORKING&gs_l=serp.3..30i10k1.1533.3987.0.4658.12.12.0.0.0.0.127.888.11j1.12.0....0...1c.1.64.serp..0.12.886...33i160k1j33i21k1.fYCPCpVPeVk
    '--------
    'WScript reference in VB
    'http://forums.codeguru.com/showthread.php?30458-WScript-reference-in-VB
    '----
    'Set objShell = WScript.CreateObject("WScript.Shell") - ------Not WORKER
    '----------------------------------
    '----------------------------------
    
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    
    WSHShell.Run """" + App.Path + "\RELOAD AND COPY.VBS" + """"
    Set WSHShell = Nothing
    Unload Me
    
End If

XVB_DATE = VB_DATE




End Sub


Private Sub Timer_SAVED_Timer()
    If EDITED = True Or EDITED2 = True Then
        If MNU_SAVED.Caption = "SAVED=TRUE " Then
            MNU_SAVED.Caption = "SAVED=FALSE"
        End If
    Else
        If MNU_SAVED.Caption = "SAVED=FALSE" Then
            MNU_SAVED.Caption = "SAVED=TRUE "
        End If
    End If
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



