VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Com1"
      Height          =   900
      Left            =   4020
      TabIndex        =   3
      Top             =   45
      Width           =   630
   End
   Begin MSComctlLib.ProgressBar ProBar1 
      Height          =   525
      Left            =   150
      TabIndex        =   2
      Top             =   2625
      Width           =   4260
      _ExtentX        =   7510
      _ExtentY        =   928
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
      Scrolling       =   1
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   2160
      Left            =   645
      TabIndex        =   0
      Top             =   450
      Width           =   3090
      _ExtentX        =   5455
      _ExtentY        =   3821
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"ClipBoard.frx":0000
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




Private Sub Form_Load()


'at$ = Command$

'at$ = "D:\VB6\VB-NT\00_Send_To\Send To Filename To Text File\temp.zip"

'at$ = """E:\04 Music ---\03 My Music Zen\01 Banging Tunes\2005\0001 - Dj Snort\Dj Snort-Tube Ride.mp3"""

CMD_LINE = Replace(Command$, """", "")


If ISIDE = True And CMD_LINE = "" Then
    CMD_LINE = "W:\MP_ROOT\100ANV01 -- AUG 15 -- CATSLE EDDIE JO TOBY EVIE"
End If


If CMD_LINE = "" Then MsgBox "No File ": End

'MsgBox R$

If Dir$(CMD_LINE) <> "" Then
    R$ = Left$(CMD_LINE, Len(CMD_LINE) - 4) + ".TXT"
    Open R$ For Output As #1
    Print #1, "------------------------------------------------"
    Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS A/P")
    Print #1, CMD_LINE
    Print #1, R$
    Print #1, "------------------------------------------------"
    Close #1
End If


If Dir$(CMD_LINE, vbDirectory) <> "" Then
    R$ = Left$(CMD_LINE, Len(CMD_LINE) - 4) + ".TXT"
    Open R$ For Output As #1
    Print #1, "------------------------------------------------"
    Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS a/p")
    Print #1, AT$
    Print #1, R$
    Print #1, "------------------------------------------------"
    Close #1
End If

NOTEPAD1 = "C:\PROGRAM FILES\NOTEPAD++\NOTEPAD++.EXE"
NOTEPAD2 = "C:\PROGRAM FILES (X86)\NOTEPAD++\NOTEPAD++.EXE"

If Dir(NOTEPAD1) <> "" Then NOTEPAD = NOTEPAD1
If Dir(NOTEPAD2) <> "" Then NOTEPAD = NOTEPAD2

If NOTEPAD <> "" Then
    Shell NOTEPAD + " " + R$, vbNormalFocus
Else
    MsgBox "NOTEPAD++ NOT FOUND -- TO FIND HERE" + vbCrLf + NOTEPAD1 + vbCrLf + "OR" + vbCrLf + NOETPAD2
End If


End
Exit Sub


RTB1.LoadFile tt$, rtfText



End Sub


Function WinMain(ByVal hInstance As Long, _
    ByVal hPrevInstance As Long, _
    lpCmdLine As String, _
    ByVal iCmdShow As Long) As Long

'Local lResult As Long
'Dim cds As COPYDATA
Dim pptr As Long
hwndWinamp = FindWindow("Winamp v1.x", "")
If hwndWinamp = 0 Then MsgBox "Winamp is not loaded!", MB_ICONINFORMATION, "Warning!"
Exit Function
'End If
'lResult = GetWindowText(hwndWinamp, g_szSongtitle, SizeOf(g_szSongtitle))
MsgBox g_szSongtitle
'SendMessage hwndWinamp, WM_USER, CloseWinamp, 0
End Function


