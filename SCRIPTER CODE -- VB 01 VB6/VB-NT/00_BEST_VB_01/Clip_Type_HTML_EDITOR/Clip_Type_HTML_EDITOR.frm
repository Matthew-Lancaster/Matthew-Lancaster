VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Clip Type Form"
   ClientHeight    =   7575
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   13605
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   13605
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txt_2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   735
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Text            =   "Clip_Type_HTML_EDITOR.frx":0000
      Top             =   945
      Width           =   11460
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   5100
      Top             =   1860
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1125
      Top             =   3045
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   6495
      TabIndex        =   8
      Top             =   2130
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   660
      Top             =   3030
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   5355
      TabIndex        =   6
      Top             =   4950
      Visible         =   0   'False
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   825
      Left            =   10755
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Clip_Type_HTML_EDITOR.frx":0012
      Top             =   4665
      Width           =   2145
   End
   Begin VB.Label Lab_COLOR_BACKGROUND 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Lab COLOUR BACKGROUND"
      Height          =   570
      Left            =   4740
      TabIndex        =   19
      Top             =   3735
      Width           =   1065
   End
   Begin VB.Label REMARK_Lab1 
      Caption         =   "NOT USED TIMER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   705
      TabIndex        =   18
      Top             =   3540
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   720
      TabIndex        =   17
      Top             =   0
      Width           =   675
   End
   Begin VB.Label GRAB_SMS_FILE 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   15
      TabIndex        =   16
      Top             =   375
      Width           =   630
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "FILE DOWN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2505
      TabIndex        =   15
      Top             =   375
      Width           =   6780
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "UP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   675
      TabIndex        =   14
      Top             =   375
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "DOWN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1110
      TabIndex        =   13
      Top             =   375
      Width           =   915
   End
   Begin VB.Label Label_Eddie 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Eddie"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5700
      TabIndex        =   12
      Top             =   0
      Width           =   825
   End
   Begin VB.Label Label_UnKnown 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "UnKnown"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6615
      TabIndex        =   11
      Top             =   0
      Width           =   1350
   End
   Begin VB.Label Mum_G 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Group"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3285
      TabIndex        =   10
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
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   11010
      TabIndex        =   9
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1425
      TabIndex        =   7
      Top             =   0
      Width           =   675
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "MS Word"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8025
      TabIndex        =   5
      Top             =   0
      Width           =   1260
   End
   Begin VB.Label Label_dad 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Dad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5055
      TabIndex        =   4
      Top             =   0
      Width           =   585
   End
   Begin VB.Label Label_Hosso 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Hoss"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4275
      TabIndex        =   3
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label_Mum 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Mum"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2505
      TabIndex        =   2
      Top             =   0
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   675
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
      Begin VB.Menu MNU_MY_LOVELY_WOMAN 
         Caption         =   "!! MY WOMAN"
         Shortcut        =   ^W
      End
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
   Begin VB.Menu MNU_VB 
      Caption         =   "!! VB"
   End
   Begin VB.Menu MNU_VBFOLDER 
      Caption         =   "!! VB FOLDER"
   End
   Begin VB.Menu Mnu_FixCaps 
      Caption         =   "!! --- FIX CAPS - PRO"
   End
   Begin VB.Menu MNU_CALENDAR_SORTER 
      Caption         =   "!! CALENDAR SORTER"
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
   Begin VB.Menu MNU_NOSAVEEXIT 
      Caption         =   "!! EXIT NO SAVE"
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

Dim Txt_3

Public FS, WINAMPEDIT
Private Const ipc_isplaying As Long = 104
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

'Private Type SHITEMID
'    cb As Long
'    abID As Byte
'End Type
'Private Type ITEMIDLIST
'    mkid As SHITEMID
'End Type
'Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

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
Dim TEST, BLUESEND, Source1
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





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()

'Debug.Print Format$(Now, "YYYY-MM-DD HH-MM-SS - DDD") + ".txt"
        
Set FS = CreateObject("Scripting.FileSystemObject")
        
' Set properties needed by MCI to open.
MMControl1.Notify = True
MMControl1.Wait = True
MMControl1.Shareable = False
MMControl1.DeviceType = "WaveAudio"

Me.BackColor = Lab_COLOR_BACKGROUND.BackColor

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

'Code to Auto Size Form based on controls used Including Menu Bar Sketchy Style
'Make sure form set to scale Twips
'put in your form load

Me.Show
DoEvents

Call HEIGHT_SUB_WITH_VAR("LABEL", TEXT_INPUT_RETURN_VAR_NUMERIC, X, Y)
     
Me.Width = (X + 2000)


Txt_2.Left = 0
Txt_2.Top = TEXT_INPUT_RETURN_VAR_NUMERIC
Txt_2.HEIGHT = 800

Call HEIGHT_SUB_WITH_VAR("TXT", TEXT_INPUT_RETURN_VAR_NUMERIC, X, Y)

Text1.Left = 0
Text1.Top = TEXT_INPUT_RETURN_VAR_NUMERIC

Text1.FontSize = 10
Txt_2.FontSize = 10

Me.HEIGHT = (Y + 2000) '+ PLUSO)
Me.Refresh

Text1.Width = Me.Width - 320
Text1.HEIGHT = Me.HEIGHT - (LABEL_Y2 + PLUSO)
Text1.Width = Me.Width - 320
Txt_2.Width = Me.Width - 320

DoEvents

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.HEIGHT - Me.HEIGHT) / 2

'ACTIVATE

Me.Show
DoEvents

'Text1.SetFocus
If ISIDE = False Then K = AlwaysOnTop(Me.hWnd)
Form1.SetFocus
Text1.SetFocus

'Text1 = Format(Now, "dd-MMM hh:mm") + vbCrLf

Text1 = ""
Call MNU_CALENDAR_SORTER_Click

Text1.SelStart = Len(Text1)

EDITED = False

End Sub

Private Sub MNU_CALENDAR_SORTER_SUB()

TT = Text1.Text

TT = Replace(TT, "<iframe src=""", "")
TT = Replace(TT, """ style=""border-width:0"" width=""800"" height=""600"" frameborder=""0"" scrolling=""no""></iframe>""", "")
TT = Replace(TT, "&amp;", "&")

Text1 = Trim(TT)

INPUT_STRING = "?title="
i1 = InStr(Text1.Text, INPUT_STRING)
i2 = InStr(Text1.Text, "&mode=")
I3 = Len(INPUT_STRING)
Txt_2 = Mid(Text1, i1 + I3, i2 - (i1 + I3))

Txt_2 = Trim(Txt_2)
Txt_2 = Replace(Txt_2, "%20", " ")
Txt_2 = Replace(Txt_2, "%2C", ",")

Text1.Text = TT

'TITLE REPLACE
'?title=
'&mode=
';mode=

End Sub

Private Sub MNU_CALENDAR_SORTER_Click()

If Clipboard.GetFormat(vbCFText) = True Then
    Text1.Text = Clipboard.GetText
    Beep
    If Text1 <> "" Then
        If InStr(Text1, "?title=") > 0 Then
            If InStr(Text1, "mode=") > 0 Then
                Call MNU_CALENDAR_SORTER_SUB
            End If
        End If
    End If
End If

End Sub

Private Sub Txt_2_Change()

If KeyCode = 27 Then Unload Me

INPUT_STRING = "?title="
i1 = InStr(Text1, INPUT_STRING)
i2 = InStr(Text1, "&mode=")
I3 = Len(INPUT_STRING)

'Txt_3 = Mid(Text1, I1 + I3, i2 - (I1 + I3))

Txt_3 = Txt_2
Txt_3 = Trim(Txt_3)
Txt_3 = Replace(Txt_3, " ", "%20")
Txt_3 = Replace(Txt_3, ",", "%2C")

Dim Mid_Text_1, Mid_Text_2, Mid_Text_3
Mid_Text_1 = Mid(Text1, 1, i1 + I3 - 1)
Mid_Text_2 = Txt_3
Mid_Text_3 = Mid(Text1, i2)

Text1.Text = Mid_Text_1 + Mid_Text_2 + Mid_Text_3

End Sub


Function HEIGHT_SUB_WITH_VAR(TEXT_INPUT, TEXT_INPUT_RETURN_VAR_NUMERIC, X, Y)
X = 0: Y = 0: TEXT_INPUT_RETURN_VAR_NUMERIC = 0
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True And Control.Visible = True Then
        
        'X AND Y COVER ALL
        If Control.Width + Control.Left > X Then X = Control.Width + Control.Left
        If Control.HEIGHT + Control.Top > Y Then Y = Control.HEIGHT + Control.Top
        
        'TEXT_INPUT -- SPECIAL REQUEST
        If InStr(UCase(Control.Name), TEXT_INPUT) > 0 Then
            DEBUG1 = Control.Name
            If Control.HEIGHT + Control.Top > LABEL_RETURN Then
                LABEL_RETURN = Control.HEIGHT + Control.Top
                DEBUG3 = Control.Name
            End If
        End If

    End If
Next
On Error GoTo 0

TEXT_INPUT_RETURN_VAR_NUMERIC = LABEL_RETURN
'HEIGHT_SUB_WITH_VAR = LABEL_RETURN

End Function

Private Sub Form_Resize()

    WINSTATETX = Me.WindowState
    Me.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)

Call ENDSAVE
On Error Resume Next
Unload Form3
Unload Form2
Unload Me

Exit Sub

If EDITED = False Then Exit Sub

If FF1 >= 1 And BlueDontSaveOnExit = False Then

'Call Mnu_Blue_Click

Mnu_SendBT_W5_Click
TRIAG = 1

End If



If Trim(Text1.Text) = "" Then Exit Sub

If Len(Text1) < 200 And ClipedText = False And DontAlter = False Then Call Mnu_Fix1st_LettersNOLOW_Click

'Clipboard.Clear
'Clipboard.SetText (Text1.Text)



'Text1 = "**** Quick Type **** " + Format(Now, "DDD DD-MMM-YY HH:MM:SS") + " ****" + vbCrLf + Text1

Source1 = App.Path + "\TextToday.txt"
FileInUseDelay (Source1)
FR1 = FreeFile
Open App.Path + "\TextToday.txt" For Append As #FR1
Print #FR1, Text1.Text
Close #FR1

'GO LOGG NOW NOT CLUE

'NO BOTH

Call Mnu_Clue_Click
If TRIG = 0 Then Call Mnu_Logg_Click


End
End Sub


Private Sub Label_Eddie_Click()
    
    Text1.Text = "Dear Eddie -- " + Format(Now, "DDD DD-MMM-YY HH:MMa/p") + vbCrLf + Text1.Text + vbCrLf + "Bro" + vbCrLf + "Matthiaoso"

    DoEvents

    If Mid(Text1, 1, 1) = vbCr Then Text1 = Mid(Text1, 3)

    FF1 = 7
    EDITED = False

End Sub


Private Sub Label_Mum_Click()
    
    Text1.Text = "Dear Mum -- " + Format(Now, "DDD DD-MMM-YY HH:MMa/p") + vbCrLf + Text1.Text + vbCrLf + "Love xxy" + vbCrLf + "Matthiaoso"

    DoEvents

    If Mid(Text1, 1, 1) = vbCr Then Text1 = Mid(Text1, 3)

    FF1 = 1
    EDITED = False

End Sub

Private Sub Label_Hosso_Click()

    Text1.Text = "Hi Liz / Dale / Frank - " + Format(Now, "DD-MMM-YY HH:MMa/p") + vbCrLf + Text1.Text + vbCrLf + "Regards" + vbCrLf + "Matthiaoso"

    DoEvents

    If Mid(Text1, 1, 1) = vbCr Then Text1 = Mid(Text1, 3)

    FF1 = 2
    EDITED = False

End Sub

Private Sub Label_Dad_Click()

    Text1.Text = "Hi Dad -- " + Format(Now, "DDD DD-MMM-YY HH:MMa/p") + vbCrLf + Text1.Text + vbCrLf + "Matthiaoso"
    
    DoEvents

    If Mid(Text1, 1, 1) = vbCr Then Text1 = Mid(Text1, 3)

    FF1 = 3
    EDITED = False

End Sub


Private Sub Label_UnKnown_Click()

    Text1.Text = "Dear Sir -- " + Format(Now, "DDD DD-MMM-YY HH:MMa/p") + vbCrLf + Text1.Text + vbCrLf + "Yours Sincerely" + vbCrLf + "Matt" + vbCrLf + "A K Respects" + vbCrLf + "Matthiaoso"
    DoEvents

    If Mid(Text1, 1, 1) = vbCr Then Text1 = Mid(Text1, 3)

    FF1 = 5
    EDITED = False

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
Form1.Text1 = II
EDITED = False

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
Form1.Text1 = II
EDITED = False

End Sub

Private Sub Label5_Click()

Source1 = "D:\TEMP\Word.txt"
FileInUseDelay (Source1)
FR1 = FreeFile
Open "D:\TEMP\Word.txt" For Output As #FR1
Print #FR1, Text1.Text;
Close #FR1

Shell "C:\Program Files\Microsoft Office\OFFICE11\WINWORD.EXE D:\TEMP\Word.txt", vbNormalFocus

Me.WindowState = 1

End Sub

Private Sub Mnu_Blue_Click()

Call Mnu_Fix1st_LettersNOLOW_Click

A1 = App.Path + "\Temp Join.txt"
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
    Print #FR1, Text1.Text;
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

If Trim(Text1.Text) = "" Then End


Me.WindowState = 1

End Sub

Private Sub Mnu_BToT_Click()

tS = sMyDocsFolder + "\00 BlueTooth SMS's From CLipType\## Blue TOTAL.txt"
Shell "C:\WINDOZE\system32\notepad.exe " + tS, vbNormalFocus

If Trim(Text1.Text) = "" Then End

Me.WindowState = 1



End Sub


Private Sub Mnu_Caps_Click()

If FF1 = 0 Then Text1 = UCase(Text1)

If FF1 > 0 Then
ts1 = InStr(Text1, Chr(13))
ts2 = InStrRev(Text1, Chr(13))
ts2 = InStrRev(Text1, Chr(13), ts2 - 1)
Text1 = Mid$(Text1, 1, ts1) + UCase(Mid$(Text1, ts1, ts2 - ts1)) + Mid(Text1, ts2)
End If
End Sub


Private Sub Mnu_CLip_Click()

Clipboard.Clear
Clipboard.SetText (Text1.Text)
Me.WindowState = 1

End Sub

Private Sub Mnu_Clue_Click()

Source1 = sMyDocsFolder + "\03 NotePad Files\00 TOP\Clued Up.txt"
FileInUseDelay (Source1)
FR1 = FreeFile
Open sMyDocsFolder + "\03 NotePad Files\00 TOP\Clued Up.txt" For Append Lock Write As #FR1
Print #FR1, "-- "; Format(Now, "DDD DD-MMM-YY HH:MM:SS") + " --"
Print #FR1, Text1.Text
Print #FR1, "-----"
Close #FR1

'Text1 = ""


End Sub

Private Sub Mnu_Conflict_Click()

Source1 = sMyDocsFolder + "\03 NotePad Files\00 TOP\Plagiarism & Conflict.txt"
FileInUseDelay (Source1)
Open Source1 For Append Lock Write As #1
Print #1, "-- "; Format(Now, "DDD DD-MMM-YY HH:MM:SS") + " --"
Print #1, Text1.Text
Print #1, "-----"
Close #1
Text1 = ""
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
text5$ = text5$ + Text1.Text + vbCrLf
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

halo = Mid$(Text1.Text, 1, 25)
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
text5$ = text5$ + Text1.Text + vbCrLf
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


halo = Mid$(Text1.Text, 1, 25)
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
text5$ = text5$ + Text1.Text + vbCrLf
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

halo = Mid$(Text1.Text, 1, 25)
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
text5$ = text5$ + Text1.Text + vbCrLf
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

halo = Mid$(Text1.Text, 1, 25)
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
text5$ = text5$ + Text1.Text + vbCrLf
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


halo = Mid$(Text1.Text, 1, 25)
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


EE = Text1.Text
   
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




Text1.Text = EE

End Sub





Private Sub Mnu_Fix1st_LettersNOLOW_Click()

EGG = 1
Call Mnu_Fix1st_Letters_Click
EGG = 0
End Sub

Private Sub Mnu_FixCaps_Click()

'Me.WindowState = 1

Dim TT As String

TT = Text1.Text

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
            
            
            
            CHKA = ".,;:?(!-<""[{~¦"
            For R1 = 1 To Len(CHKA)
                GOMORE = 0
                ty = Mid$(CHKA, R1, 1)
                vf4 = InStr(YY5, ty)
'                If TY = ";" And VF4 > 0 Then Stop
                If vf4 > 0 Then
                    For r2 = vf4 To Len(YY5)
                        i = Mid(YY5, r2, 1)
                        If IsNumeric(i) Then
                            Exit For
                        End If
                        Mid(YY5, r2, 1) = UCase(Mid(YY5, r2, 1))
                        If Mid(YY5, r2, 1) <> i Then
                            GOMORE = 1
                            Exit For
                        End If
                    Next
                End If
            Next
        
            Mid(TT, VFX + 1, Len(YY5)) = YY5 '+ vbCrLf
        
            'VF = VFX + 1
    Loop Until vf = 0 Or vf - 1 > Len(TT)



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

Text1.Text = TT



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

A1 = App.Path + "\Temp Join.txt"
A2 = sMyDocsFolder + "\03 NotePad Files\00 TOP\Jobs.txt"

DUM = JoinFiles(A1, A2, i)
If DUM = True Then Unload Me

End Sub

Private Sub Mnu_KadMan_Click()

A1 = App.Path + "\Temp Join.txt"
A2 = sMyDocsFolder + "\03 NotePad Files\01 TOP\Kadaitcha Man Techy NFO.txt"

DUM = JoinFiles(A1, A2, i)
If DUM = True Then Unload Me

End Sub

Private Sub Mnu_Logg_Click()

'always goes to clue already end prog
'Call Mnu_Clue_Click

A1 = App.Path + "\Temp Join.txt"
A2 = sMyDocsFolder + "\03 NotePad Files\00 TOP\01 My Logg.txt"

DUM = JoinFiles(A1, A2, i)
If DUM = True Then Unload Me

End Sub

Private Sub MNU_MY_LOVELY_WOMAN_Click()



A1 = App.Path + "\Temp Join.txt"
A2 = sMyDocsFolder + "\03 NotePad Files\00 TOP\My Woman.txt"

DUM = JoinFiles(A1, A2, i)
If DUM = True Then Unload Me



Exit Sub



Source1 = sMyDocsFolder + "\03 NotePad Files\00 TOP\My Woman.txt"
FileInUseDelay (Source1)
FR1 = FreeFile
Open Source1 For Append Lock Write As #FR1
Print #FR1, "-- "; Format(Now, "DDD DD-MMM-YY HH:MM:SS") + " --"
Print #FR1, Text1.Text
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

A1 = App.Path + "\Temp Join.txt"
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

A1 = App.Path + "\Temp Join.txt"
A2 = sMyDocsFolder + "\03 NotePad Files\Word Play\#0-Unet Notes.txt"

DUM = JoinFiles(A1, A2, i)
If DUM = True Then Unload Me

End Sub

Sub ENDSAVE()

On Error Resume Next

Do
    
    If FILESAVENAME1 = "" Then
        If FS.FOLDEREXISTS(App.Path + "\TEXTBOX") = False Then
            MkDir App.Path + "\TEXTBOX"
        End If
        FILESAVENAME1 = App.Path + "\TEXTBOX\TEXTBOX TEXT SAVE " + Format$(Now, "YYYY-MM-DD HH-MM-SS - DDD") + ".txt"
    End If
    
    If FILESAVENAME2 = "" Then
        If FS.FOLDEREXISTS(App.Path + "\TEXTBOX BIGGER") = False Then
            MkDir App.Path + "\TEXTBOX BIGGER"
        
        End If
        FILESAVENAME2 = App.Path + "\TEXTBOX BIGGER\TEXTBOX TEXT SAVE -BIGGER-" + Format$(Now, "YYYY-MM-DD HH-MM-SS - DDD") + ".txt"
    End If
        
    Err.Clear
    FR = FreeFile
    Open FILESAVENAME1 For Output As #FR
        Print #FR, Text1.Text
    Close #FR

    
    LENL = Len(Text1.Text)
    If OLENL < LENL Then
        FR = FreeFile
        Open FILESAVENAME2 For Output As #FR
            Print #FR, Text1.Text
        Close #FR
    End If
    
    If Err.Number > 0 Then
        MsgBox "Error Save Try Agian Press"
        Sleep 1000
    End If

Loop Until Err.Number = 0
    
End Sub

Private Sub Mnu_VB_Click()

'Clipboard.Clear
'Clipboard.SetText (Text1.Text)

Call ENDSAVE


Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
End

End Sub

Private Sub MNU_VBFOLDER_Click()

Shell "explorer /select, " + FILESAVENAME1 + "\", vbNormalFocus
Me.WindowState = 1
End Sub

Private Sub Mum_G_Click()
    
    Text1.Text = "Hi Mum & Group -- " + Format(Now, "DDD DD-MMM-YY HH:MMa/p") + vbCrLf + Text1.Text + vbCrLf + "A K Respects" + vbCrLf + "Matthiaoso"
    
    DoEvents

    If Mid(Text1, 1, 1) = vbCr Then Text1 = Mid(Text1, 3)

    FF1 = 4

End Sub

Private Sub Text1_Change()

EDITED = True
EDITED2 = True


'If OTEXT <> Text1.Text Then
'    OTEXT = Text1.Text
    'If EDITED3 > Now Then
    'End If
    
'End If

LL = Len(Text1.Text) - 400
If LL < 1 Then LL = 1

If InStr(Text1.Text, "`") Then
    XZ = Text1.SelStart
    'TCH = StrConv("`", vbUnicode)
    Text1.Text = Replace(Text1.Text, "`", "") 'SHOULD DO ONLY END
    Text1.SelStart = XZ
End If
        
        
WINAMPEDIT = Now + TimeSerial(0, 0, 5)
        
        
        
EDITED3 = Now + TimeSerial(0, 0, 20)



If EDITED3 = -1 Then EDITED3 = 0

LOWERTIMER = Now + TimeSerial(0, 1, 0)

'If Mid(Text1, 1, 1) = vbCr Then Stop: Exit Sub

BlueDontSaveOnExit = False

Label1 = Len(Text1)
If SMSLEN - Len(Text1) > 0 Then
    Label9 = Str(SMSLEN - Len(Text1))
Else
    Label9 = 0
End If

gg = Len(Text1)
If OGG > gg + 2 Or OGG < gg - 2 Then

ClipedText = True
End If
OGG = gg

'5 SMS
Label1.BackColor = 0

SMSLEN = 763

If Len(Text1) > SMSLEN And FF1 > 0 Then

Label1.BackColor = QBColor(12)

'10 SMS
'If Len(Text1) > 1530 And FF1 >0 Then

'Beep
    
    MMControl1.Command = "Open"
    MMControl1.Command = "Play"
        Call WaitWavFinish
    MMControl1.Command = "Close"



End If

If Text1.SelStart > 1 Then

If (Mid$(Text1, 1, 1) <> UCase(Mid$(Text1, 1, 1))) And Len(Text1) < 2 Then
    Text1 = UCase(Mid$(Text1, 1, 1)) + Mid$(Text1, 2)
    Text1.SelStart = 1
    Text1.SelLength = Len(Text1.Text)
End If
End If

'If Len(Text1) = 1 Then
'Text1 = Format(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + "-- Start Text --" + vbCrLf + Text1
'Text1.SelStart = Len(Text1)
'End If

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then Unload Me

EDITED3 = Now + TimeSerial(0, 0, 20)

If KeyCode = 223 Then KeyCode = 0: Exit Sub




Label6 = Trim(Str(Text1.SelStart))

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
        Print #1, Text1.Text
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

'Text1 = ""

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

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)

Label6 = Trim(Str(Text1.SelStart))



End Sub

Private Sub Timer1_Timer()
    
    Exit Sub
    
    Label7 = Format(Now, "ddd dd-mmm-yy hh:mm:ssa/p")

        TF2 = FindWindow("Winamp v1.x", vbNullString)
            
        MR = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)
        If OLDMR <> MR Then
        XZ = Text1.SelStart
        
        
        If MR = 1 Then
        'Text1.Text = Text1.Text + LF + "<<  VIDEO PLAY >>" + vbCrLf
        End If
        If MR <> 1 Then
        'Text1.Text = Text1.Text + LF + "<< VIDEO STOP >>" + vbCrLf
        
        End If
        
        Text1.SelStart = Len(Text1.Text)

        
        End If
        OLDMR = MR


End Sub

Private Sub Timer2_Timer()

Exit Sub

If Me.WindowState = 1 Then WINAMPEDIT = 1

    TF2 = FindWindow("Winamp v1.x", vbNullString)
    
    If WINAMPEDIT > Now Then
        messageResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)
        If messageResult <> 1 Then
            'MsgResult2 = SendMessage(TF2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
        End If
    
    Else
    
    If WINAMPEDIT < Now Or WINAMPEDIT = 0 Then
        messageResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)
        If messageResult = 1 Then
            'MsgResult2 = SendMessage(TF2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
        End If
        WINAMPEDIT = 0
    End If
    
End If


                
                

If Me.WindowState = 1 Then EDITED3 = 1




'  <  NOW
If EDITED3 < Now And EDITED3 > 0 Then
    
    LF = vbCrLf
    
    If Mid(Text1.Text, Len(Text1.Text) - 2) = vbCrLf Then
        LF = ""
    End If
    
    If InStrRev(Text1.Text, " >>" + vbCrLf) > Len(Text1.Text) - 5 = False Then
        XZ = Text1.SelStart
        Text1.Text = Text1.Text + LF + "<< " + Format$(Now, "DD-MMM-YY -- HH:MM:SSa/p") + " >>" + vbCrLf
        Text1.SelStart = Len(Text1.Text)
    End If
    
    Text1.SelStart = Len(Text1.Text)
    
    OTEXT = Text1.Text
    EDITED3 = -1
    
End If

If LOWERTIMER < Now And LOWERTIMER > 0 Then
    LOWERTIMER = 0
    'Me.WindowState = 1
End If

If EDITED2 = True Then
    EDITED2 = False

    Call ENDSAVE
    OTEXT = Text1.Text

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
    
    'VIDEO NOT USER -- EXIT SUB
    
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
