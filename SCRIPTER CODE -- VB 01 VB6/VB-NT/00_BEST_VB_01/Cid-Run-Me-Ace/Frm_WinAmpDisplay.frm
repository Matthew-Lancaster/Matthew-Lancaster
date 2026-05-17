VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Frm_WinAmpDisplay 
   BackColor       =   &H00000000&
   Caption         =   "WinAmp Display"
   ClientHeight    =   2856
   ClientLeft      =   480
   ClientTop       =   924
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   2856
   ScaleWidth      =   11160
   Visible         =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   5940
      Top             =   765
   End
   Begin VB.PictureBox TextToSpeech1 
      Height          =   1155
      Left            =   525
      ScaleHeight     =   1104
      ScaleWidth      =   1092
      TabIndex        =   2
      Top             =   450
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   7725
      Top             =   1200
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   1875
      TabIndex        =   1
      Top             =   975
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   593
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer1 
      Interval        =   15000
      Left            =   8775
      Top             =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2865
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11115
   End
End
Attribute VB_Name = "Frm_WinAmpDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Public Voice As SpVoice
Public MusicM$
Public VoiceActive
Public VoiceActiveMp3
Public VoiceActiveMp32
Public VoiceActiveMp33 As Date
Public VoiceTimeWait



'Public Xxr4
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

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
Private Const Nexttrackbutton = 40048
Private Const Raisevolumeby1percent = 40058
Private Const Lowervolumeby1percent = 40059


Public MsgResult As Long
Public XcNul As Long
Public LhRet As Long

'------------




Private Sub Form_Activate()
Frm_WinAmpDisplay.Top = 400
Frm_WinAmpDisplay.Left = 0
Timer1.Enabled = True

End Sub

Public Sub Form_Load()

'Frm_WinAmpDisplay.Show
Frm_WinAmpDisplay.Top = 400
Frm_WinAmpDisplay.Left = 0

Call Frm_WinAmpDisplay.DisplayWinamp


MMControl1.Notify = True
MMControl1.Wait = True
MMControl1.Shareable = False
MMControl1.DeviceType = "WaveAudio"

MMControl1.fileName = "D:\Wave's\Cosmi\CANNON.WAV"

'Open the MCI WaveAudio device.
MMControl1.Command = "Open"

'MMControl1.Command = "Prev"
'MMControl1.Command = "Play"




End Sub

Public Sub DisplayWinamp()

Winamphand2 = FindWindow("Winamp v1.x", vbNullString)
WinampHand3 = FindWindow("Winamp PE", vbNullString)

'   Initialize the voice object

Set Voice = New SpVoice

Set Voice.Voice = Voice.GetVoices().Item(1)
If IsIDE() = True Then
    Set Voice.Voice = Voice.GetVoices().Item(2)
End If
'5=man
'4=lady

'swap 2 & 3 if when compiled & 4 & 5  -- 1 & 2

'   Set Voice.Voice = Voice.GetVoices().Item(?) 'sample TTS
'   Set Voice.Voice = Voice.GetVoices().Item(3) 'microsoft Sam
'   Set Voice.Voice = Voice.GetVoices().Item(2) 'microsoft Mike
'   Set Voice.Voice = Voice.GetVoices().Item(0) 'microsoft mary


'Voice.Status.RunningState

If Winamphand2 <> WinampHand2Rem Then
wfile2 = GetFileFromHwnd(Winamphand2)
    rt$ = Mid$(wfile2, 1, InStrRev(wfile2, "\"))
    Tt2 = InStrRev(wfile2, "\")
    R2$ = Mid$(wfile2, InStrRev(wfile2, "\", Tt2 - 1) + 1)
    R2$ = Mid$(R2$, 1, InStrRev(R2$, "\")) + " "
End If

WinampHand2Rem = Winamphand2

xer = 0
Xx$ = rt$
tty1$ = rt$ + "Current Play To Text File.txt"
tty1$ = Dir$(tty1$)
tty2$ = rt$ + "Current Play To Text File Append.txt"
tty3$ = rt$ + "Current Play To Text File -One.txt"

If tty1$ = "" Then

    On Error Resume Next
    
    Path_And_FileName = tty3$
    Call CID_Run_Me.FileInUseDelay(Path_And_FileName, 0)
    
    Path_And_FileName = tty3$
    f6 = FreeFile
    Open Path_And_FileName For Input As #f6
        Line Input #f6, Gg$
    Close #f6
    On Error GoTo 0

End If
If Gg$ = "" Then Gg$ = "-- WinAmp Not Loaded --"



If tty1$ <> "" Then
    Path_And_FileName = Xx$ + tty1$
    F4 = FreeFile
    Call CID_Run_Me.FileInUseDelay(Path_And_FileName, 0)
    Path_And_FileName = Xx$ + tty1$
    Open Path_And_FileName For Input As #F4
        Call CID_Run_Me.FileInUseDelay(tty2$, 1)
        f5 = FreeFile
        Path_And_FileName = tty2$
        Open tty2$ For Append As #f5
            Print #f5,
            Print #f5, Format$(Now, "DDD DD-MMM-YY HH:MM:SS ap")
            Print #f5, "--- -- --- -- -- -- -- -"
            Do
                Err.Clear
                On Local Error Resume Next
                Line Input #F4, Gg$
                xer = Err.Number
                Print #f5, Gg$
            Loop Until EOF(F4)
        Close #F4
    Print #f5,
    Close #f5
    
'    If GetUserName = "Matt3" Then
    
    Path_And_FileName = tty3$
    Call CID_Run_Me.FileInUseDelay(Path_And_FileName, 1)

    Path_And_FileName = tty2$
    f6 = FreeFile
    Open tty3$ For Output As #f6
        Print #f6, Gg$
    Close #f6

'    End If

    Kill Xx$ + tty1$

End If

If InStr(Gg$, "- Winamp") Then
    r = InStr(Gg$, "- Winamp")
    tt$ = Trim(Mid$(Gg$, 1, r - 1)) + Trim(Mid$(Gg$, r + 8))
    Gg$ = R2$ + tt$
End If

Xx = 0
If Frm_WinAmpDisplay.Label1.Caption = Gg$ Then Xx = 1

gh = InStr(Gg$, "K - ")
If gh > 0 Then
gx$ = Mid$(Gg$, 1, gh) + vbCrLf + Trim(Mid$(Gg$, gh + 4))
Else
gx$ = Gg$
End If

gh = InStrRev(Gg$, "\")
If gh > 0 Then
gx$ = Mid$(Gg$, 1, gh) + vbCrLf + Trim(Mid$(Gg$, gh + 1))
Else
gx$ = Gg$
End If



Frm_WinAmpDisplay.Label1.Caption = gx$

        


If InStr(Gg$, "[Paused]") > 0 Then
    Xx = 1
    If VoiceTimeWait = False Then MusicM$ = ""
End If

If InStr(Gg$, "[Stopped]") > 0 Then Xx = 1
If InStr(Gg$, "<Nothing>") > 0 Then Xx = 1
If InStr(Gg$, "-- WinAmp Not Loaded --") > 0 Then Xx = 1


ee = FS.FileExists("D:\TEMP\PauseTimeWinampFlag.txt")
If ee = True Then
If Xx = 0 Then Kill "D:\TEMP\PauseTimeWinampFlag.txt"
Xx = 1
End If

If Xx = 0 Then
'    Frm_WinAmpDisplay.WindowState = 0
'    Frm_WinAmpDisplay.Show
    
    'Here
    'Frm_WinAmpDisplay.Visible = True
    
    If Trim(Mid$(Gg$, gh + 4)) <> MusicM$ Then
        MusicM$ = Trim(Mid$(Gg$, gh + 4))
        If Now > VoiceActiveMp33 And VoiceActiveMp33 > 0 Then
            VoiceActiveMp33 = 0
            VoiceActive = True
            Call VoiceMP3Pause
            If CID_Run_Me.Mnu_Voiceonoff.Checked = False Then
                Voice.Speak Trim(Mid$(Gg$, gh + 4)), SVSFlagsAsync
            End If
        End If
    End If
End If
'If Xxr4 < Now And Xxr4 > 0 Then
'    MMControl1.Command = "Prev"
 '   MMControl1.Command = "Play" '
'End If
'in display routine


'Xxr4 = Now + TimeSerial(0, 0, 20)

Timer1.Enabled = True

End Sub


Private Sub Timer1_Timer()

'Frm_WinAmpDisplay.WindowState = 1
Timer1.Enabled = False
Frm_WinAmpDisplay.Visible = False
'Unload Frm_WinAmpDisplay
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Exit Sub


'Label15.Caption = Str(Voice.Status.RunningState)
If VoiceActive = True Then

    If Voice.Status.RunningState = 1 Then
        VoiceActive = False
        If VoiceActiveMp3 = 1 Then
            Winamp22Handle2 = FindWindow("Winamp v1.x", vbNullString)
            MsgResult2 = SendMessage(Winamp22Handle2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
            VoiceActive = False
            VoiceTimeWait = False
        End If
    End If

End If

End Sub
Sub VoiceMP3Pause()

    If VoiceActive = False Then Exit Sub
    'Winamp22Handle2 = FindWindow("Winamp PE", vbNullString)
    Winamp22Handle2 = FindWindow("Winamp v1.x", vbNullString)
    
    MsgResult = SendMessage(Winamp22Handle2, WM_USER, 0, ByVal ipc_isplaying)

    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.
    VoiceActiveMp3 = MsgResult
    If VoiceActiveMp3 = 1 Then
        MsgResult2 = SendMessage(Winamp22Handle2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
        VoiceTimeWait = Now + TimeSerial(0, 0, 2)
        VoiceTimeWait = True
    End If

End Sub

Private Sub Timer3_Timer()
    Timer3.Enabled = False
    Exit Sub
    
    Winamp22Handle2 = FindWindow("Winamp v1.x", vbNullString)
    
    MsgResult = SendMessage(Winamp22Handle2, WM_USER, 0, ByVal ipc_isplaying)

    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.
    If MsgResult <> VoiceActiveMp32 Then
    VoiceActiveMp33 = Now + TimeSerial(0, 0, 15)
    VoiceActiveMp32 = MsgResult
    End If
End Sub
