VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3192
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3192
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer_MAIN 
      Interval        =   1
      Left            =   180
      Top             =   696
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Com1"
      Height          =   900
      Left            =   4020
      TabIndex        =   3
      Top             =   45
      Width           =   630
   End
   Begin VB.PictureBox ProBar1 
      Height          =   525
      Left            =   150
      ScaleHeight     =   480
      ScaleWidth      =   4212
      TabIndex        =   2
      Top             =   2625
      Width           =   4260
   End
   Begin VB.PictureBox RTB1 
      Height          =   2160
      Left            =   645
      ScaleHeight     =   2112
      ScaleWidth      =   3048
      TabIndex        =   0
      Top             =   450
      Width           =   3090
   End
   Begin VB.Label Label3 
      Caption         =   "Label2"
      Height          =   450
      Left            =   3780
      TabIndex        =   5
      Top             =   1650
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   450
      Left            =   3780
      TabIndex        =   4
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   360
      Left            =   555
      TabIndex        =   1
      Top             =   45
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PATH_FILE


Public g_hInst As Long
Public g_hIcon As Long
Public hwndWinamp As Long
Dim g_szSongtitle As String * 255




'Private Declare Function SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_ON

'Private Declare Function SendMessage _
'        Lib "user32" _
'        Alias "SendMessageA" _
'        (ByVal hwnd As Long, _
'         ByVal wMsg As Long, _
'         ByVal wParam As Long, _
'         ByVal lParam As Any) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
        
        
        
Private Const MONITOR_ON = -1&
Private Const MONITOR_OFF = 2&
Private Const SC_MONITORPOWER = &HF170&
Private Const WM_SYSCOMMAND = &H112
Private Const ipc_isplaying As Long = 104

Private Const WINAMP_BUTTON2 = 40045
Private Const WINAMP_BUTTON3 = 40046
Private Const WM_CLOSE = &H10
Private Const WM_USER = &H400
Private Const WM_COMMAND = &H111
Private Const GW_HWNDNEXT = 2
Dim MsgResult As Long
Dim XcNul As Long
Dim LhRet As Long
Private Const GET_PLAY_POSITION = &H105 'If Data is 0, returns the position In milliseconds of playback. If Data is 1, returns current track length In seconds. Returns -1 If Not playing Or If an Error occurs.


Public at2$
'GET_PLAY_POSITION


Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText at2$, vbCFText
End
End Sub


'
'Function WinMain(ByVal hInstance As Long, _
'    ByVal hPrevInstance As Long, _
'    lpCmdLine As String, _
'    ByVal iCmdShow As Long) As Long
'
''Local lResult As Long
''Dim cds As COPYDATA
'Dim pptr As Long
'hwndWinamp = FindWindow("Winamp v1.x", "")
'If hwndWinamp = 0 Then MsgBox "Winamp is not loaded!", MB_ICONINFORMATION, "Warning!"
'Exit Function
''End If
''lResult = GetWindowText(hwndWinamp, g_szSongtitle, SizeOf(g_szSongtitle))
'MsgBox g_szSongtitle
''SendMessage hwndWinamp, WM_USER, CloseWinamp, 0
'End Function
'


Sub GET_PATH_OR_FILE_PATH()

Set FS = CreateObject("Scripting.FileSystemObject")

AT_TEMP = Command$
AT_TEMP = Replace(AT_TEMP, """", "")

'CLIPBOARD DONE LATER
'If LINE_CLIP = "" Then
'    LINE_CLIP = Clipboard.GetText
'    'IF DIR(LINE_CLIP,vbDirectory )
'End If

'---------
'TEST MODE -- FILE
'---------
AT_ISIDE = "C:\TEMP\bcdinfo.txt"
If IsIDE = True Then AT_TEMP = AT_ISIDE
'---------
'TEST MODE -- FOLDER
'---------
'AT_ISIDE = "C:\TEMP\"
'If IsIDE = True Then AT_TEMP = AT_ISIDE


'Me.Height = Label2.Top + Label2.Height + (830)
'Me.Width = Label3.Left + Label3.Width + (300)


'--------------------------
'FIND COMMANDLINE - SEND TO
'--------------------------
If FS.FolderExists(AT_TEMP) = True Then
    PATH_FILE = AT_TEMP
    SENDTO = "SEND TO RIGHT CLICK"
End If
If FS.FileExists(AT_TEMP) = True Then
    PATH_FILE = AT_TEMP
    File = True
    SENDTO = "SEND TO RIGHT CLICK"
End If

'---------------------
'NONE SEND TO CMD LINE -- THEN
'IS CLIPBOARD ANY
'---------------------
AT_TEMP = Replace(Trim(Clipboard.GetText), """", "")
If PATH_FILE = "" Then
    
    If FS.FolderExists(AT_TEMP) = True Then
        PATH_FILE = AT_TEMP
        SENDTO = "CLIPBOARD FIND PICK"
    End If
        
    If FS.FileExists(AT_TEMP) = True Then
        PATH_FILE = AT_TEMP
        SENDTO = "CLIPBOARD FIND PICK"
        File = True
    End If
End If
    
Label1 = PATH_FILE

If File = True Then
    'PATH_FILE = Mid(AT_TEMP, 1, InStrRev(AT_TEMP, "\"))
'    Label3 = Mid(PATH_FILE, 1, InStrRev(PATH_FILE, "\"))

End If


'MNU_STATUS.Caption = "  -- SOURCE GIVEN IS -- " + SENDTO

'If FILE = True Then
'TIMER1.Enabled = True
'End If


'--------------------------
'STILL NONE AFTER CLIPBOARD
'OR
'SEND TO
'USE REMEMBER FILE
'--------------------------

'If PATH_FILE = "" Then
'    MsgBox " NOTHING WITH VALID FOLDER OF FILE NAME GIVEN" + vbCrLf + "COMMAND$" + vbCrLf + Command$ + vbCrLf + "OR CLIPBOARD" + vbCrLf + AT_TEMP
'    End
'End If


'Label1

'ChDrive w$
'ChDir w$

'If FILE = False Then
'    Shell "explorer /e, /select, " + PATH_FILE, vbNormalFocus
'    End
'Else
'    Exit Sub
'    Shell "explorer /e, /select, " + PATH_FILE, vbNormalFocus
'End If


End Sub

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
  'Test = False
End Function
'***********************************************

Private Sub Timer_MAIN_Timer()
'at$ = Command$

'at$ = "D:\VB6\VB-NT\00_Send_To\Send To Filename To Text File\temp.zip"

'at$ = """E:\04 Music ---\03 My Music Zen\01 Banging Tunes\2005\0001 - Dj Snort\Dj Snort-Tube Ride.mp3"""

GET_PATH_OR_FILE_PATH

'CMD_LINE = Replace(Command$, """", "")

CMD_LINE = PATH_FILE

'If IsIDE = True And CMD_LINE = "" Then
'    CMD_LINE = "W:\MP_ROOT\100ANV01 -- AUG 15 -- CATSLE EDDIE JO TOBY EVIE"
'End If
'Kill "D:\VB6\VB-NT\00_Send_To\Send To Filename To Text File\Send To Filename To Text File.TXT"
'CMD_LINE = "D:\VB6\VB-NT\00_Send_To\Send To Filename To Text File\Send To Filename To Text File.frm"


If CMD_LINE = "" Then
    'Me.Show
    DoEvents
    
    MsgBox "Not File " + vbCrLf + CMD_LINE + vbCrLf + Command$, vbMsgBoxSetForeground
End If

Set m_fsoObject = New FileSystemObject

If m_fsoObject.FileExists(CMD_LINE) = True Then
    R$ = Left$(CMD_LINE, Len(CMD_LINE) - 4) + ".txt"
    If Dir$(R$) <> "" Then
        MsgBox "File Already Exist" + vbCrLf + R$
    End If
    
    If Dir$(R$) = "" Then
        Open R$ For Output As #1
            Print #1, "------------------------------------------------"
            Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS A/P")
            Print #1, CMD_LINE
            Print #1, R$
            Print #1, "------------------------------------------------"
        Close #1
        DONE_YES = True
    End If
End If

If m_fsoObject.FolderExists(CMD_LINE) = True And DONE_YES = False Then
    FileName = Mid(CMD_LINE, InStrRev(CMD_LINE, "\") + 1)
    R$ = CMD_LINE + "\" + FileName + ".txt"
    Debug.Print R$
    Open R$ For Output As #1
        Print #1, "------------------------------------------------"
        Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS a/p")
        Print #1, CMD_LINE
        Print #1, R$
        Print #1, "------------------------------------------------"
    Close #1
End If

NOTEPAD1 = "C:\PROGRAM FILES\NOTEPAD++\NOTEPAD++.EXE"
NOTEPAD2 = "C:\PROGRAM FILES (X86)\NOTEPAD++\NOTEPAD++.EXE"

If Dir(NOTEPAD1) <> "" Then NOTEPAD = NOTEPAD1
If Dir(NOTEPAD2) <> "" Then NOTEPAD = NOTEPAD2

If NOTEPAD <> "" And R$ <> "" Then
    '---------------------------------------------
    'SET REFERENCES __ MICROSOFT SCRIPTING RUNTIME
    '---------------------------------------------
    VBS_LAUNCHER_NAME = NOTEPAD
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
        WSHShell.Run """" + VBS_LAUNCHER_NAME + """ """ + R$ + """"
    Set WSHShell = Nothing
    
    'YOU REQUIRE ADMIN TO RUN SHELL
    'USE THE ABOVE INSTEAD
    '------------------------------
'    On Error Resume Next
'    I = Shell("""" + NOTEPAD + """ """ + R$ + """", vbNormalFocus)
'    If I = 0 Then
'        MsgBox "SHELL NOTEPAD++ DID NOT WORK MOSTLY YOU REQUIRE ADMIN"
'    End If
End If


If NOTEPAD = "" And R$ <> "" Then
    MsgBox "NOTEPAD++ NOT FOUND -- TO FIND HERE" + vbCrLf + NOTEPAD1 + vbCrLf + "OR" + vbCrLf + NOETPAD2
End If


'Unload Me
'Exit Sub


'RTB1.LoadFile tt$, rtfText

Unload Me

End Sub


