VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   630
      TabIndex        =   6
      Top             =   3330
      Width           =   3900
   End
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
      _ExtentX        =   7514
      _ExtentY        =   926
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
      _ExtentX        =   5450
      _ExtentY        =   3810
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


at$ = Command$
'at$ = "G:\00 DVD Decrypter\00 Done"
'MsgBox at$
'at$ = """E:\04 Music ---\03 My Music Zen\01 Banging Tunes\2005\0001 - Dj Snort\Dj Snort-Tube Ride.mp3"""

If at$ <> "" And Mid$(at$, 1, 1) = """" Then
    at$ = Mid$(at$, 2)
    at$ = Mid$(at$, 1, Len(at$) - 1)
End If
'Label1.Caption = at$
On Error Resume Next

ChDir at$

If Err.Number > 0 Then MsgBox "No Dir ": End
On Error GoTo 0

Dir1.Path = at$
yy$ = ""
For r1 = 0 To Dir1.ListCount
    tt$ = Dir1.List(r1)
    tt$ = Mid$(tt$, InStrRev(tt$, "\") + 1)
    yy$ = yy$ + tt$ + vbCrLf
Next

MsgBox yy$
Clipboard.Clear
Clipboard.SetText yy$

End
Exit Sub


Dim winamp, fs, currentTrack, trackPath, trackFolder, trackFile

Set winamp = CreateObject("gen_scripting.WinAmp")

Set fs = CreateObject("Scripting.FileSystemObject")

'winamp.ForwardFiveSeconds
'trackPath = winamp.GetPlaylistFile(currentTrack)
'currentTrack = winamp.GetPlaylistPosition()

'x = winamp.GetVersion()

'winamp.TogglePlaylist

'x = winamp.GetTrackPositionMilliseconds






'rtfRTF 0 (Default) RTF. The file loaded must be a valid .rtf file.
'rtfText 1 Text. The RichTextBox control loads any text file.


'tt$ = "C:\"
'tt$ = GetFolder(Me.hWnd, "Scan Path:", tt$)

'tt$ = "C:\HTTrack\01 Matts Web Httrack.txt"

'RTB1.LoadFile tt$, rtfText
'Open tt$ For Binary As #1
'Input(number, [#]filenumber)

'yy$ = Input(LOF(1), 1)
'Close #1

'Stop
'Clipboard.Clear
'Clipboard.SetText au$, vbCFText
'Clipboard.SetText RTB1.Text

Dim MsgResult As Long
Dim Winamp22Handle2 As Long
Dim Winamp22Handle3 As Long
Winamp22Handle3 = FindWindow("Winamp PE", vbNullString)
Winamp22Handle2 = FindWindow("Winamp v1.x", vbNullString)
    MsgResult = SendMessage(Winamp22Handle2, WM_USER, 0, ByVal ipc_isplaying)




    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.
'    AppActivate "Winamp Playlist Editor"

    If MsgResult = 1 Then
'        MsgResult = SendMessage(Winamp22Handle2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
    'XcNul = 0
        
        MsgResult = SendMessage(Winamp22Handle2, WM_USER, GET_PLAY_POSITION, ByVal XcNul)
'        x = winamp.GetTrackPositionMilliseconds()       ' Get current position in the track
        'MsgBox "Play pos " + Str(MsgResult)
    
    End If
'    If MsgResult = 0 Or MsgResult = 3 Then
'        MsgResult = SendMessage(Winamp22Handle2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
'    End If


at$ = Command$
'at$ = """E:\04 Music ---\03 My Music Zen\01 Banging Tunes\2005\0001 - Dj Snort\Dj Snort-Tube Ride.mp3"""
If at$ <> "" And Mid$(at$, 1, 1) = """" Then
at$ = Mid$(at$, 2)
at$ = Mid$(at$, 1, Len(at$) - 1)
End If
'Label1.Caption = at$


If at$ = "" Then MsgBox "No File ": End
r$ = Left$(at$, Len(at$) - 4) + ".txt"
'MsgBox R$

If Dir$(r$) = "" And r$ <> "" Then
Open r$ For Output As #1
Print #1, "------------------------------------------------"
Print #1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS a/p")
Print #1, at$
Print #1, r$
Print #1, "------------------------------------------------"
Close #1
Shell "notepad2.exe " + r$, vbNormalFocus
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


