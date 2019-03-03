VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Object = "{019FD1D9-2A35-4D5E-AB4E-BA793946941B}#1.0#0"; "ClipboardViewer.ocx"
Begin VB.Form FRM_ClipTest 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clipboard Viewer"
   ClientHeight    =   6468
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   16116
   Icon            =   "frmClipTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6468
   ScaleWidth      =   16116
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Text3 
      Height          =   2892
      Left            =   4320
      TabIndex        =   15
      Top             =   1320
      Width           =   4212
      _ExtentX        =   7430
      _ExtentY        =   5101
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmClipTest.frx":1272
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   2892
      Left            =   0
      TabIndex        =   14
      Top             =   1320
      Width           =   4212
      _ExtentX        =   7430
      _ExtentY        =   5101
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmClipTest.frx":12FD
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   732
      Left            =   10800
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   2532
      _ExtentX        =   4466
      _ExtentY        =   1291
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmClipTest.frx":1388
   End
   Begin VB.Timer Timer_Test_Logic 
      Interval        =   1000
      Left            =   10956
      Top             =   384
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Text"
      Height          =   495
      Left            =   9036
      TabIndex        =   4
      Top             =   1656
      Width           =   1320
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy Picture 2"
      Height          =   495
      Index           =   1
      Left            =   9036
      TabIndex        =   3
      Top             =   1020
      Width           =   1320
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   432
      Index           =   1
      Left            =   10416
      Picture         =   "frmClipTest.frx":1413
      ScaleHeight     =   384
      ScaleWidth      =   360
      TabIndex        =   2
      Top             =   996
      Width           =   408
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy Picture 1"
      Height          =   495
      Index           =   0
      Left            =   9036
      TabIndex        =   1
      Top             =   408
      Width           =   1320
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   432
      Index           =   0
      Left            =   10416
      Picture         =   "frmClipTest.frx":1661
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   0
      Top             =   372
      Width           =   432
   End
   Begin ClipboardViewer.ctlClipboard ctlClipboard1 
      Left            =   10956
      Top             =   1008
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label1 
      Caption         =   "Compare if not the Same"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   5
      Left            =   5052
      TabIndex        =   12
      Top             =   72
      Visible         =   0   'False
      Width           =   648
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Compare if not the Same"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   4
      Left            =   4368
      TabIndex        =   11
      Top             =   72
      Visible         =   0   'False
      Width           =   648
   End
   Begin VB.Label Label1 
      Caption         =   "Compare when not the Same"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   3
      Left            =   36
      TabIndex        =   10
      Top             =   72
      Width           =   4140
   End
   Begin VB.Label Label_SIZE_2 
      Caption         =   "SIZE 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4344
      TabIndex        =   9
      Top             =   876
      Width           =   4140
   End
   Begin VB.Label Label_SIZE_1 
      Caption         =   "SIZE 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   36
      TabIndex        =   8
      Top             =   888
      Width           =   4140
   End
   Begin VB.Label Label1 
      Caption         =   "Text in Clipboard -- BEFORE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   4344
      TabIndex        =   7
      Top             =   468
      Width           =   4140
   End
   Begin VB.Label Label1 
      Caption         =   "Text in Clipboard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   1
      Left            =   36
      TabIndex        =   6
      Top             =   468
      Width           =   4140
   End
   Begin VB.Label Label1 
      Caption         =   "Picture in Clipboard"
      Height          =   276
      Index           =   0
      Left            =   9072
      TabIndex        =   5
      Top             =   2316
      Width           =   1608
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1800
      Left            =   9072
      Stretch         =   -1  'True
      Top             =   2616
      Width           =   2460
   End
   Begin VB.Menu Mnu_Exit 
      Caption         =   "Exit"
   End
   Begin VB.Menu MNU_API_RESET 
      Caption         =   "API"
   End
   Begin VB.Menu Mnu_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_ADMINISTRATOR 
      Caption         =   "ADMIN"
   End
   Begin VB.Menu MNU_CLIP_TIME 
      Caption         =   "GIVE ME TIME -->"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_COMPARE 
      Caption         =   "COMPARE 1 2"
   End
End
Attribute VB_Name = "FRM_ClipTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EXIT_TRUE

'I = MsgBox("ClipBoard API Has Stopped and Gone Missing" + vbCrLf + "Use the Menu Option *INFO* to Diagnostic and Reload It" + vbCrLf + "This Can Happen If ChkDsk Unlocked All Handles to The Hard Drive and the ClipboardViewer.ocx Driver Couldn't Get Access", vbMsgBoxSetForeground)
'ClipboardViewer.ocx Driver
'ClipboardViewer.oca Driver -- TRUE NAME EXT
'THE VERSION USE IS
'ClipboardViewer.oca
'C:\Program Files\Microsoft Visual Studio\VB98\ClipboardViewer.oca   0x488   29/08/2015 22:56:01 03/04/2016 21:48:47 A   10,240  *           *           0x00120089  1,696   7244    VB6.EXE C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE   oca 16.56%
'MUST BE THE REGISTER VERSION
'AS ONE IN ROOT OF APP EXE DON'T LOCATE TO USE

Dim i, MSGQ, IR
Dim CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2
Dim FileSpec, FILE_SPEC_GO

Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long

' Public constants for set window on top
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

' Public constants for SetWindowPos API declaration
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Option Explicit

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


'Public HOOK_CLIPBOARD_API_lOADED As Boolean

Private Sub cmdCopy_Click(Index As Integer)

Select Case Index
    Case 0
        Clipboard.SetData Picture1(0).Picture, vbCFDIB
    Case 1
        Clipboard.SetData Picture1(1).Picture, vbCFDIB
End Select

End Sub

Private Sub Command1_Click()

Clipboard.SetText Text1.Text

End Sub

Private Sub ctlClipboard1_ClipboardChanged()

Timer_API_Test_Logick_Var2 = Now
TIME_LAST_CLIPBOARD = Now

'-------------------------------TEST2017
'Call Form1.COMPARE_FOR_EXE


'If EXECUTE_TIMER_ENABLED = False Then Exit Sub

'---------------------------------------------------------
'DELAY RELOAD API CLIPPER AS FREQENT AFTER IDLE FEW SECOND
'TO MUCH LONGER 8 MINUTE WHEN CLIPBOARD DONE
'---------------------------------------------------------

CLIPBOARD_ACTIVITY_MOMENT = Now + TimeSerial(0, 8, 0)

'-------------------------------TEST2017
'Call Form1.Timer1_Timer

'Exit Sub

If Clipboard.GetFormat(vbCFText) = True Then
    Text3.Text = Text2.Text
    On Error Resume Next
    Do
        Err.Clear
        Text2.Text = Clipboard.GetText
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
    
End If

If Clipboard.GetFormat(vbCFBitmap) = True Then
    On Error Resume Next
    Do
        Err.Clear
        Image1.Picture = Clipboard.GetData(vbCFBitmap)
        If Err.Number > 0 Then Sleep 200
    Loop Until Err.Number = 0
    On Error GoTo 0
End If

End Sub

Private Sub Form_Click()
If IsIDE = True Then End
End Sub

Private Sub Form_Load()

If App.PrevInstance = True Then End

EXIT_TRUE = False

'HOOK_CLIPBOARD_API_lOADED = True

'-------------------------------TEST2017
'Form1.Mnu_API_Reload_Status.Caption = "The API Form Clipper Logger Sub Call Loaded @ " + Format(Now, "DDD DD-MM-YYYY HH:MM:SS")

'-------------------------------TEST2017
'If Form1.Mnu_API_UNLoad_Status.Caption = "The API Form Clipper Logger Sub Call UN-Loaded @ " Then
'    Form1.Mnu_API_UNLoad_Status.Caption = "The API Form Clipper Logger Sub Call UN-Loaded @ NOT HAPPEN YET NOT A TIME SET AT MOMENT"
'End If

'Me.Hide
'Start to monitor the clipboard

ctlClipboard1.StartClipboardViewer

'-----------------------------
O_VAL_MINUTE_API = Now
'Call Form1.MNU_API_RESET_SUB
'-----------------------------

Call Text2_Change

Call MNU_ADMINISTRATOR_Click

SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

Me.Width = Text3.Left + Text3.Width + 200

Me.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)


If Me.WindowState <> vbMinimized And EXIT_TRUE = False Then Cancel = True: Me.WindowState = vbMinimized: Exit Sub

'HOOK_CLIPBOARD_API_lOADED = False

'-------------------------------TEST2017
'Form1.Mnu_API_UNLoad_Status.Caption = "The API Form Clipper Logger Sub Call UN-Loaded @ " + Format(Now, "DDD DD-MM-YYYY HH:MM:SS")

'Stop the clipboard viewer
ctlClipboard1.EndClipboardViewer

'Call Form1.MNU_API_RESET_Click

End Sub

Private Sub Label1_Click(Index As Integer)
'Label1
Call MNU_COMPARE_Click
End Sub

'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
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
    If Dir(FileSpec) = "" Then FILE_SPEC_GO = CODER_VBP_FILE_NAME_2
    
    If Dir(FILE_SPEC_GO) = "" Then MsgBox "ERROR NOT EXISTOR _ " + FILE_SPEC_GO
    
    CODER_VBP_FILE_NAME_2 = FILE_SPEC_GO
End Sub

Private Sub MNU_API_RESET_Click()
EXIT_TRUE = True
Unload Me
DoEvents
Load FRM_ClipTest

End Sub

Private Sub Mnu_Exit_Click()
EXIT_TRUE = True
Unload Me
End Sub

Private Sub Mnu_VB_ME_Click()
    
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
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
    objShell.Run """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
    Set objShell = Nothing
    
    EXIT_TRUE = True
    Unload Me
    Exit Sub
    
    
    Beep
    
'    i = MsgBox(GetSpecialfolder(38))
    
    If GetSpecialfolder(38) = "" Then
        MsgBox "YOU HAVE TO SWITCH ADMIN MODE ON" + vbCrLf + vbCrLf + "THE COMMAND GetSpecialfolder(38) REQUIRE ADMINISTRATOR" + vbCrLf + vbCrLf + "RESULT RETURN NOTHING HERE GetSpecialfolder(38)__<" + GetSpecialfolder(38) + ">__ RESULT RETURNED NOTHING", vbMsgBoxSetForeground
        Exit Sub
    End If
    
    If IsIDE = True Then
        If IsIDE = True Then
            MSGQ = vbYesNoCancel
        Else
            MSGQ = vbOKOnly
        End If
        IR = MsgBox("NOT IN IDE ARE YOU SURE", MSGQ, vbMsgBoxSetForeground)
        If IR <> vbYes Then Exit Sub
    End If
    
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
    Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + CODER_VBP_FILE_NAME_2 + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    End
End Sub
Private Sub MNU_VB_FOLDER_Click()
    Beep
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
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
'---------------------------------------------------
'Private Type ITEMIDLIST
'    mkid As SHITEMID
'End Type
'Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
'Function GetSpecialfolder(ByVal CSIDL As Long) As String
'    Dim R As Long
'    Dim IDL As ITEMIDLIST
'    R = SHGetSpecialFolderLocation(100, CSIDL, IDL)
'    If R = NOERROR Then
'        Path$ = Space$(512)
'        R = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
'        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
'        Exit Function
'    End If
'    GetSpecialfolder = ""
'End Function
'---------------------------------------------------

Private Sub Text2_Change()

Dim SAME_VALUE

If Text2.Text = Text3.Text Then
    SAME_VALUE = "SAME IS TRUE"
    Label1(3).BackColor = Label1(5).BackColor
Else
    SAME_VALUE = "SAME IS FALSE"
    Label1(3).BackColor = Label1(4).BackColor
End If

Label_SIZE_1.Caption = "SIZE =" + Str(Len(Text2.Text)) + " -- " + SAME_VALUE
Label_SIZE_2.Caption = "SIZE =" + Str(Len(Text3.Text)) + " -- " + SAME_VALUE
End Sub

Private Sub Text3_Change()
    Call Text2_Change
End Sub

Public Sub Timer_Test_Logic_Timer()

Timer_Test_Logic.Enabled = False
Timer_Test_Logic.Interval = 10000

Exit Sub

'REMOVE THIS ONE AUG 08 2K SIXTEEN


'HOOK_CLIPBOARD_API_lOADED = True

If IsIDE = True Then
    Timer_Test_Logic.Interval = 4000
Else
    Timer_Test_Logic.Interval = 1000
End If

Timer_API_Test_Logick_Var1 = Now

End Sub


Private Sub MNU_COMPARE_Click()
    
    Dim FILENAME_COMPARE_1
    Dim FILENAME_COMPARE_2
    Dim COMPARE_EXE_1, COMPARE_EXE_2
    Dim FR1
    Beep
    
    FILENAME_IN_USE_CHECK = App.Path + "\# DATA\" + GetComputerName + "\Data\"
    If Dir(FILENAME_IN_USE_CHECK, vbDirectory) = "" Then
        CreateFolderTree FILENAME_IN_USE_CHECK
    End If
    
    
    FILENAME_IN_USE_CHECK = App.Path + "\# DATA\" + GetComputerName + "\Data\#COMPARE _" + GetComputerName + "_ 1 DUPE CHECKER.Txt"
    'DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_COMPARE_1 = FILENAME_IN_USE_CHECK
    FR1 = FreeFile
    On Error Resume Next
    Open FILENAME_IN_USE_CHECK For Output As #FR1
        Print #FR1, Text2.Text;
    Close #FR1

    FILENAME_IN_USE_CHECK = App.Path + "\# DATA\" + GetComputerName + "\Data\#COMPARE _" + GetComputerName + "_ 2 DUPE CHECKER.Txt"
    'DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_COMPARE_2 = FILENAME_IN_USE_CHECK
    FR1 = FreeFile
    On Error Resume Next
    Open FILENAME_IN_USE_CHECK For Output As #FR1
        Print #FR1, Text3.Text;
    Close #FR1

    Beep
    Shell "C:\Program Files (x86)\WinMerge\WinMergeU.exe" + " """ + FILENAME_COMPARE_1 + """ """ + FILENAME_COMPARE_2 + """"
    Beep
    Me.WindowState = vbMinimized

End Sub
