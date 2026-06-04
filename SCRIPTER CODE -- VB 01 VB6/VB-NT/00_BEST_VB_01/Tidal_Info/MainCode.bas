Attribute VB_Name = "aa_MainCode"
Public FORCE_TALKER_GO_WITHOUT_CHECKER_DUPE_IGNOR

Dim FILENAME_IN_USE_CHECK As String

Public MOTION_FAST_FORWARD_BUTTON_GO
Public MOTION_REVERSE_BUTTON_GO
Public JOY_VOLUME_UP_BUTTON_GO
Public JOY_VOLUME_DOWN_BUTTON_GO
Public JOY_PREVIOUS_TRACK_BUTTON_GO
Public JOY_NEXT_TRACK_BUTTON_GO
Public JOY_PAUSE_BUTTON_GO

Public JOY_NEXT_TRACK
Public JOY_CONTROLLER

Public SAY_ONCE_MOST_LIKELY_A_INCORECT_VOICE_DRIVER

Public ALL_FORM_REQUEST_TO_END
Public TrueTerminate
Public WSHShell


Public MOON_STATE_LIGHTER

Public VOXLIST_NOT_ACTIVE, VOXLIST_NOT_ACTIVE_1ST_RUN




Public Count_To_Min_Vox_Window

Public EXIT_FORM_SET

Public JOYPAD_SOCKETED_IN

Public Mwnd_OLD

Public IsVBrunningForCodeVAR

Public STRING_OF_VOICE_IN_QUE_PLAY
Public VoiceStreamActive

Public WINAMPBackOnAfterVoiceRequired
Public TALKTIME$

Public StringVoiceToPLayActive
Public StringVoiceToPLayActiveFLAGGED

Public oVolumeSet2

'Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, _
    ByVal lParam As Long) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As _
    Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

' The following variables are shared between the main ChildWindows procedure
' and the auxiliary (private) ChildWindows_CBK routine

' An array of Long holding the handle of all child windows.
Dim Windows() As Long
' The number of elements in the array.
Dim WindowsCount As Long




Public j As JOYINFO
Public Je As JOYCAPS
Public be As B
Public be1 As B


'Public Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
Public F, FileSpec, FS, RamDrive
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public WinAmpPowerHwnd, WAFastFF, TTz1, TTz2, TTz3, TTz4, TTz5, TTz6

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type


Public Abcd255 As Integer
Public Abcd355 As Integer
Public Abcd455 As Integer
Public Abcd655 As Integer

Public OldAbcd As Integer

Public Declare Function PolyBezier Lib "gdi32" (ByVal HDC As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
Public ptBez1() As POINTAPI

Private Type POINTAPI
        x As Long
        y As Long
End Type

Dim ExeGo
Public WinAmpChkIsPlayingHwnd
Public WinAmpChkIsPlayingHwndOld
Public InCode, Test
Public CapsStat, NumsStat, ScrosStat
Public lpStrUser
Public lpStrPassword

'------------------
'Required for
'Function FileInUse(ByVal strFilePath As String) As Boolean

Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long

Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40

Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'--------------

Public YinVectKeli$

Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Public OYYG, TYYG, LockDownX, LockDown2

Public OYYG2, TYYG2, LockVideoWinAmp
Public VCodeTT_, VCodeTTOld$

Public TTCount1, TTCount2
Public TTCount3, TTCount4
Public TTCount5, TTCount6
Public TTCount7, TTCount8
Public KeyClean
Public ADate1Old, ADate2Old, ADate3Old, ADate4Old

Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function Putfocus _
        Lib "user32" _
        Alias "SetFocus" _
        (ByVal hWnd As Long) As Long

Public Declare Function WindowFromPoint _
                 Lib "user32" (ByVal xPoint As Long, _
                               ByVal yPoint As Long) As Long



Public OldWinTitle, OldAsh, URL$, VolumeSet

Private Const WM_GETTEXTLENGTH = &HE
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5
Private Const WM_SETTEXT            As Long = &HC
Private Const WM_GETTEXT = &HD



Public SysStartTimeOldN, SysStartTimeNewN

Public Keyy As Long
Public Mousey As Long

Public NoteP, Np3
Public tty, SayTime, LockSSaver, LockSaverDelay, DeadLock
Public LockSaverDelay1, LockSaverDelay2, LockSaverDelay3
Public LockSaverDelay4 'LockSaverDelay4 '--- Short delay even shorter for IDE

'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

'SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_ON

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByRef lParam As Any) _
    As Long



Public XxG As Date

Public Sld1, Sld2, Sld3, Sld4, Sld5, Chked1, Chked2, Chked3, Chked4, Chked5, Chked6, Chked7, XPud
Public DChked5, DChked6, DChked7
'Pattern1---------
Public WK As Double
Public HW As Double
Public KW As Double
Public TwD1 As Double
Public TwD2 As Double
Public TwD3 As Double
Public V As Double

'--------------



'Flower Pattern
Public noin As Integer
Public atim As Double
Public Handle As Integer
'Public wk As Integer
'Public hw As Double
'Public kw As Double
'Public twd1 As Double
'Public twd2 As Double
'Public twd3 As Double
Public RR3 As Double
Public X3 As Single
Public Y3 As Single
Public X4 As Single
Public Y4 As Single
Public x5 As Single
Public y5 As Single
Public pi As Double
Public TRZ

Public Size, Steps1, Steps2, Dot_Or_Line, dw
'--------------------------


Private Const MONITOR_ON = -1&
Private Const MONITOR_OFF = 2&
Private Const SC_MONITORPOWER = &HF170&
Private Const WM_SYSCOMMAND = &H112
Public Const ipc_isplaying As Long = 104

Public Const IPC_GETPLAYLISTFILE  As Long = 211
Public Const IPC_GETOUTPUTTIME  As Long = 105
Public Const IPC_GETINFO As Long = 126

Public Const WINAMP_BUTTON1 = 40044 ' __ PLAY AROUND FIND A BUTTON
Public Const WINAMP_BUTTON2 = 40045 ' Start play if stopped
Public Const WINAMP_BUTTON3 = 40046 ' Start play from paused or pause if Playing
Public Const WINAMP_BUTTON4 = 40047 ' STOP BUTTON
Public Const WINAMP_BUTTON5 = 40048 ' __ PLAY AROUND FIND A BUTTON

'Public Const WM_CLOSE = &H10
'Public Const WM_USER = &H400
Public Const WM_COMMAND = &H111
'Private Const GW_HWNDNEXT = 2
Public Const Nexttrackbutton = 40048
Public Const WINAMP_PREVSONG = 40198
Public Const Raisevolumeby1percent = 40058
Public Const Lowervolumeby1percent = 40059
Public Const WINAMP_REW5S = 40061                    '// rewinds 5 seconds
Public Const WINAMP_FFWD5S = 40060                  '// fast forwards 5 seconds
Public Const ID_PE_CLOSE = 40224
'#define WINAMP_OPTIONS_EQ               40036 // toggles the EQ window
'#define WINAMP_OPTIONS_PLEDIT           40040 // toggles the playlist window
'#define WINAMP_VOLUMEUP                 40058 // turns the volume up a little
'#define WINAMP_VOLUMEDOWN               40059 // turns the volume down a little
'#define WINAMP_FFWD5S                   40060 // fast forwards 5 seconds
'#define WINAMP_REW5S                    40061 // rewinds 5 seconds



Public MsgResult As Long
Public XcNul As Long
Public LhRet As Long

'------------


'------------------------------------
'Outlook_Express_Philo
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type

'Option Explicit
Private Declare Sub GetKeyboardState Lib "user32" (ByVal lpKeyState As String)
Private Declare Sub SetKeyboardState Lib "user32" (ByVal lpKeyState As String)
' function switches caps lock on



Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long




Public URL2$
Public URLLoc$


      Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, _
         lpRect As RECT) As Long

      Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, _
         ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
         ByVal nHeight As Long, ByVal bRepaint As Long) As Long

      Declare Function GetDesktopWindow Lib "user32" () As Long

      Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadID _
         As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long

      Declare Function GetWindowThreadProcessId Lib "user32" _
         (ByVal hWnd As Long, lpdwProcessId As Long) As Long

      Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
         (ByVal hWnd As Long, ByVal lpClassName As String, _
         ByVal nMaxCount As Long) As Long

      Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
         (ByVal hWnd As Long, ByVal lpString As String, _
         ByVal cch As Long) As Long

      Public TopCount As Integer     ' Number of Top level Windows
      Public ChildCount As Long   ' Number of Child Windows
      Public ThreadCount As Long  ' Number of Thread Windows

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWndParent As Long, _
     ByVal hWndChildAfter As Long, _
     ByVal lpszClass As String, _
     ByVal lpszTitle As String) _
    As Long


Public Hwnd22



Public CompName$
Public Rabbit As Date
Public Cheese


Public dks As Integer
Public Sluty As Integer
Public Sluty2 As Integer
Public Sluty3 As Integer
Public Sluty4 As Integer
Public rnowt2 As Date


Public SoggyPie As Integer

Public QuitForms As Integer
Public Soff As Integer
Public Toff As Integer

Public Gooze1
Public Gooze2
Public Gooze3
Public Gooze4
Public Gooze5
Public Gooze6
Public Gooze7
Public Gooze8

Public Qkey
Public Qkey2 As Date

'Public Count2Check

Public Nb() As Long
Public ab$()

Public Enum WindowStates
  SW_HIDE = 0
  SW_SHOWNORMAL = 1
  SW_NORMAL = 1
  SW_SHOWMINIMIZED = 2
  SW_SHOWMAXIMIZED = 3
  SW_MAXIMIZE = 3
  SW_SHOWNOACTIVATE = 4
  SW_SHOW = 5
  SW_MINIMIZE = 6
  SW_SHOWMINNOACTIVE = 7
  SW_SHOWNA = 8
  SW_RESTORE = 9
  SW_SHOWDEFAULT = 10
  SW_FORCEMINIMIZE = 11
  SW_MAX = 11
End Enum


Public ash2$

'Public Const GW_CHILD = 5


Public URLadt$
Public URLadt2$


'Public Const HWND_TOPMOST = -1
'Public Const HWND_NOTOPMOST = -2
'Public Const SWP_NOACTIVATE = &H10
'Public Const SWP_NOMOVE = &H2
'Public Const SWP_NOSIZE = &H1
'Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
'Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GWL_WNDPROC = (-4)
Private Const IDANI_OPEN = &H1
Private Const IDANI_CLOSE = &H2
Private Const IDANI_CAPTION = &H3
Private Const WM_USER = &H400



Public Function IsFileOpenDelay(FILENAME_IN_USE_CHECK_2 As String)
    
    '---------------------------------------------------
    'REPEAT USE OF code
    'MAKE USE OF THIS ONE CODE SUB ROUTINE THIS ONE LOOK
    'BETTER MORE SIMPLE
    'IsFileOpenDelay
    '---------------------------------------------------
    'FileInUseDelay REDIRECTED HERE
    '---------------------------------------------------
    
    'And Here QQ_LIMIT = 90 --- FileInUseDelay
    Dim QQ_LIMIT, QQ, DD_DIFF, ii, O_DD_DIFF
    QQ_LIMIT = 80
    QQ = Now + TimeSerial(0, 0, QQ_LIMIT)
    
    ATidalX.MNU_FILE_STUCK_IN_USE.Visible = False
    DD_DIFF = 0
    Do
        DoEvents
        ii = IsFileAlreadyOpen(FILENAME_IN_USE_CHECK_2)
        If ii = True Then
            Sleep (200)
            DD_DIFF = DateDiff("S", Now, QQ)
            ATidalX.MNU_FILE_STUCK_IN_USE.Visible = True
            If DD_DIFF <> O_DD_DIFF Then
                ATidalX.MNU_FILE_STUCK_IN_USE.Caption = "IsFileOpenDelay Busy --" + str(DD_DIFF) + " Secs - File Stuck In Use Of Limit -" + str(QQ_LIMIT) + " -- " + FILENAME_IN_USE_CHECK_2
            End If
            O_DD_DIFF = DD_DIFF
        End If
    Loop Until ii = False Or QQ < Now
    
    If ii = False Then
        ATidalX.MNU_FILE_STUCK_IN_USE.Caption = "Last IsFileOpenDelay --" + str(QQ_LIMIT) + " Second"
    End If
    
    If ii = True Then
        ATidalX.MNU_FILE_STUCK_IN_USE.Caption = "Error IsFileOpenDelay --" + str(DD_DIFF) + " Second File Stuck In Use -- Of Limit --" + str(QQ_LIMIT) + " -- " + FILENAME_IN_USE_CHECK_2
        MSGBOX2 = "ERROR IsFileOpenDelay " + vbCrLf + FILENAME_IN_USE_CHECK_2 + vbCrLf + str(DD_DIFF) + " Second of --" + str(QQ_LIMIT) + " Limit"
        frm_MSGBOX.Timer1.Enabled = True
    End If
End Function





Public Function GetWindowState(ByVal lngHwnd As Long) As Integer
    
    'CareFul this It Can Be a Max window in the Minimized Posi
    If IsWindow(lngHwnd) = 1 Then
        GetWindowState = -1
        If IsIconic(lngHwnd) <> 0 Then
            GetWindowState = vbMinimized
        ElseIf IsZoomed(lngHwnd) <> 0 Then
            GetWindowState = vbMaximized
        End If
    End If


    '--- USE THIS IN FURTURE
'    If IsWindow(lngHwnd) = 1 Then
'        GetWindowState = -1
'    If IsIconic(lngHwnd) <> 0 Then
'        GetWindowState = vbMinimized
'    ElseIf IsZoomed(lngHwnd) <> 0 Then
'        GetWindowState = vbMaximized
'    End If
'    End If
'
'    RESULT 0 OR 1 -- 0 = vbMinimized
'    RESULT IN TRUE IS NOT VISIBLE SO EQUAL -2 FOR OUT RESULT
'    If IsWindowVisible(lngHwnd) = 0 Then
'        GetWindowState = -2
'    End If


End Function

Public Sub WinAmpPause()

    '-----------------------------------------------------------------
    'ONLY USE THIS NOT A PAUSE WHEN TIME NOT JOY PAD CONTROL LIKE HERE
    '-----------------------------------------------------------------
    'If ATidalX.MNU_NO_PAUSE.Checked = True Then Exit Sub

    If VoiceStreamActive = True Then Exit Sub
    If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then Exit Sub
    
    ATidalX.TIMER_WinAmpPause.Enabled = False


'    ii = Now + TimeSerial(0, 1, 0)
'    Do While VoiceStreamActive = True Or ii < Now
'        DoEvents
'        Sleep 100
'    Loop
'
'    If ii < Now Then
'        Debug.Print "WinAmpPause"
'        Debug.Print "Stuck in This Loop Longer Than 1 Min"
'    End If
'
    
    'LOAD IF NOT PLAYING
    'PROBLEM
    
    If FindWindow("Winamp v1.x", vbNullString) = 0 Then
        'Fr1 = FreeFile
        'Open "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\0TextData\WinAmpEXEFileActive.txt" For Input As #Fr1
        '    Line Input #Fr1, TT$
        'Close #Fr1

        'TT$ = "C:\Program Files\00 WinAmp's\Winamp 556 & Remote\winamp.exe"
        
        TT$ = "C:\Program Files\WinAmp\winamp.exe"
        If Dir(TT$) = "" Then
            TT$ = "C:\Program Files (X86)\WinAmp\winamp.exe"
        End If
        Shell TT$, vbNormalNoFocus
        Exit Sub
    
    End If

    Call PAUSEALLWINAMP

'-------------
Exit Sub
'-------------

    If MsgResult <> 0 Then
        'Start play from paused or pause if Playing
        MsgResult2 = SendMessage(WinAmpPowerHwnd, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
        Kill "C:\TEMP\Pause All For WinAMP-STOP.txt"
        F = FreeFile
        Open "C:\TEMP\Pause All For WinAMP-PLAY.txt" For Output As #F
        Close #F
    End If
    
    
    '=PAUSE SOUND
    If MsgResult = 1 Then
        ATidalX.MMControl6.Command = "Prev"
        ATidalX.MMControl6.Command = "Play"
    End If
    
    If MsgResult = 0 Then
        'start play if stopped
        MsgResult2 = SendMessage(WinAmpPowerHwnd, WM_COMMAND, WINAMP_BUTTON2, ByVal XcNul)
        Kill "C:\TEMP\Pause All For WinAMP-PLAY.txt"
        F = FreeFile
        Open "C:\TEMP\Pause All For WinAMP-STOP.txt" For Output As #F
        Close #F
    
    End If
            
            
            
            
End Sub

Sub ISWINAMP_PLAYING_DO_THE_ACTION()

    'DONT USE PAUSE IN THIS PROG
    XX = GetForegroundWindow
    If XX = FindWindow(vbNullString, "JPEG Encoder -- Shockingly Good Photo ART -- [-:¦:-•:*''' The One '''*:•-:¦:-]") Then
        Exit Sub
    End If

    test_hwnd = FindWindow("Winamp v1.x", vbNullString)

    If test_hwnd = 0 Then Exit Sub
    
    MsgResult = SendMessage(test_hwnd, WM_USER, 0, ByVal ipc_isplaying)
    'GONE ONTO PAUSE
    If MsgResult = 1 Then MSGX = 1
    
    WIN_TITLE_WINAMP = GetWindowTitle(test_hwnd)

    SET_GO = False
    If InStr(WIN_TITLE_WINAMP, "(FNOOB TECHNO RADIO) - Winamp") > 0 Then
        SET_GO = True
    End If
    If InStr(WIN_TITLE_WINAMP, "Nubreaks.com Radio") > 0 Then
        SET_GO = True
    End If
    
    If SET_GO = True Then
    
        ' Start Play from paused ALWAYS WHEN STREAMER
        ' -------------------------------
        If MsgResult = 3 Then
            ' STOP PLAY AND THEN PRESS PLAY
            ' -----------------------------
            MsgResult2 = SendMessage(test_hwnd, WM_COMMAND, WINAMP_BUTTON4, ByVal XcNul)
            MsgResult2 = SendMessage(test_hwnd, WM_COMMAND, WINAMP_BUTTON2, ByVal XcNul)
            Exit Sub
        End If
    
    End If
    
    If MsgResult <> 0 Then
        'Start play from paused or pause if Playing
        MsgResult2 = SendMessage(test_hwnd, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
    End If

    If MsgResult = 0 Then
        'start play if stopped
        MsgResult2 = SendMessage(test_hwnd, WM_COMMAND, WINAMP_BUTTON2, ByVal XcNul)
    End If

End Sub

Sub PAUSEALLWINAMP()



Call ISWINAMP_PLAYING_DO_THE_ACTION


Exit Sub


Dim cText As String
test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
    
Do While test_hwnd <> 0
    
    cText = Space$(255)
    G = GetClassName(test_hwnd, cText, 255)
    cText = StripNulls(cText)
    
    
    If cText = "Winamp v1.x" Then
    
        MsgResult = SendMessage(test_hwnd, WM_USER, 0, ByVal ipc_isplaying)
    
        'GONE ONTO PAUSE
        If MsgResult = 1 Then MSGX = 1
        
        If MsgResult <> 0 Then
            'Start play from paused or pause if Playing
            MsgResult2 = SendMessage(test_hwnd, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
        End If
    
        If MsgResult = 0 Then
            'start play if stopped
            MsgResult2 = SendMessage(test_hwnd, WM_COMMAND, WINAMP_BUTTON2, ByVal XcNul)
        End If
    End If
    
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop

On Error Resume Next

If MSGX <> 1 Then       '=PAUSE SOUND
    ATidalX.MMControl6.Command = "Prev"
    ATidalX.MMControl6.Command = "Play"
    
    If 1 = 2 Then
        Kill "C:\TEMP\Pause All For WinAMP-PLAY.txt"
        F = FreeFile
        Open "C:\TEMP\Pause All For WinAMP-STOP.txt" For Output As #F
        Close #F
    End If
    
Else
    
    If 1 = 2 Then
        Kill "C:\TEMP\Pause All For WinAMP-STOP.txt"
        F = FreeFile
        Open "C:\TEMP\Pause All For WinAMP-PLAY.txt" For Output As #F
        Close #F
    End If

End If



End Sub





Sub WinAmpNextVIDEOTrack()
    

End Sub

Sub WinAmpNextTrack()
    
    If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then Exit Sub
    If VoiceStreamActive = True Then Exit Sub
    ATidalX.TIMER_WinAmpNextTrack.Enabled = False
    
    'Next Track
    MsgResult2 = SendMessage(WinAmpPowerHwnd, WM_COMMAND, Nexttrackbutton, ByVal XcNul)

End Sub

Sub WinAmpPrevTrack()
        
    If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then Exit Sub
    If VoiceStreamActive = True Then Exit Sub
    ATidalX.TIMER_WinAmpPrevTrack.Enabled = False
        
    MsgResult2 = SendMessage(WinAmpPowerHwnd, WM_COMMAND, WINAMP_PREVSONG, ByVal XcNul)
    
    '--------------------
    Exit Sub
    '--------------------
    
    
    'Next Track
    pass1 = 0
    
    
    Do
        ash1$ = GetWindowTitle(WinAmpPowerHwnd)
        MsgResult2 = SendMessage(WinAmpPowerHwnd, WM_COMMAND, WINAMP_PREVSONG, ByVal XcNul)
        'Length Song
        MsgResult33 = SendMessage(WinAmpPowerHwnd, WM_USER, 1, ByVal IPC_GETOUTPUTTIME)

        'ash2$ = GetWindowTitle(WinAmpPowerHwnd)
        'If pass1 = 0 Then
        '    Do
        '        DoEvents
        '    Loop Until ash2$ <> ash1$
        'End If
        'pass1 = 1
        'ash2$ = Mid$(ash2$, InStr(ash2$, ". "), 5)
        
    'Loop Until ash2$ <> ". ---" And ash2$ <> ". 0#0"
    Loop Until MsgResult33 > 0
        
End Sub

Sub WinAmpFF()
    'Fast Forward Some
    MsgResult2 = SendMessage(WinAmpPowerHwnd, WM_COMMAND, WINAMP_FFWD5S, ByVal XcNul)

End Sub

Sub WinAmpRW()
    'Rewind Some
    MsgResult2 = SendMessage(WinAmpPowerHwnd, WM_COMMAND, WINAMP_REW5S, ByVal XcNul)

End Sub


Public Sub WinAmpCommands(NiceVar)
    
    If NiceVar = 0 Then Exit Sub
    
    Sluty3 = 0
    
    'WinAmpPowerHwnd = FindWindow("Winamp v1.x", vbNullString)
    'Winamp22Handle3 = FindWindow("Winamp PE", vbNullString)
         
    gcw2 = GetForegroundWindow
    Gcw$ = GetWindowTitle(gcw2)
    'weh = FindWindow("Outlook Express Browser Class", vbNullString)
    
    'MsgResult =
    'WinAmpPowerHwnd = WinAmpChkIsPlayingHwnd
    
    'If WinAmpPowerHwnd = 0 Then Exit Sub
    
    MsgResult = ATidalX.WinAmpChkIsPlaying
    
    If NiceVar = 2 Then
        ATidalX.TIMER_WinAmpPause.Enabled = True
    End If
    
    If ATidalX.WinAmpChkIsPlaying = 0 Then Exit Sub
    
    'If NiceVar = 3 Then Call aa_MainCode.WinAmpNextTrack
    'If NiceVar = 6 Then Call aa_MainCode.WinAmpPrevTrack
    
    If NiceVar = 3 Then ATidalX.TIMER_WinAmpNextTrack.Enabled = True
    If NiceVar = 6 Then ATidalX.TIMER_WinAmpPrevTrack.Enabled = True
    
    If ATidalX.WinAmpChkIsPlaying <> 1 Then Exit Sub
    If NiceVar = 4 Then
        Call WinAmpFF
    End If
    If NiceVar = 5 Then
        Call WinAmpRW
    End If

    If 1 = 2 And NiceVar = 8 Then
        ' Start play ANYWAY ----- if stopped
        ' WANT CHECK IF IN STREAM MODE
        ' AND THEN NOT REQUIRE PAUSE UNPAUSE JUST USE START PLAY AGAIN
        ' USE SAME PAUSE BUTTON THEN
        ' THIS FOR NOW
        ' JUST GRAB TITLE CHECK FOR FNOOBTECHNO EASYIER FOR NOW
        ' DON'T REALLY WANT EASIER RESTART ACCIDENTAL
        MsgResult2 = SendMessage(test_hwnd, WM_COMMAND, WINAMP_BUTTON2, ByVal XcNul)
    End If


End Sub


Public Sub KeyWindowCheck()
'Call SetKeys("CAPSLOCK_OFF")

'Mark Msg By Convo
'If Sluty4 = 2 Then
'Sluty4 = 0
'SendKeys "^{INSERT}", True
'SendKeys "^C", False
'MsgBox "done"
'End If


'If GetForegroundWindow = FindWindow("Outlook Express Browser Class", vbNullString) Then


'Mark Messages By Conversation
If Sluty4 = 1 Then
         
         ash$ = GetWindowTitle(weh)
         Dim XY2  As RECT
         er = GetWindowRect(weh, XY2)

         rt = XY2.Top
         
         'MsgBox Str(rt)

         'th = weh.WindowState
        
        If rt <> -32000 Then
            SendKeys "%v", True
            SendKeys "{right}", True
            SendKeys "{up}", True
            SendKeys "{enter}", True
            
        End If




Sluty4 = 0
Exit Sub
End If



If Sluty3 = 7 Then
    'np2 = FindWindow(vbNullString, GetSpecialfolder(CSIDL_PERSONAL) +"\Andy Latest.txt - Notepad2")
    'Np3 = FindWindow(vbNullString, "* "+GetSpecialfolder(CSIDL_PERSONAL)+"\Andy Latest.txt - Notepad2")
    Np2 = FindWindow(vbNullString, GetSpecialfolder(CSIDL_PERSONAL) + "\CallerID\2ABBREV.TXT - Notepad2")
    Np3 = FindWindow(vbNullString, "* " + GetSpecialfolder(CSIDL_PERSONAL) + "\CallerID\2ABBREV.TXT - Notepad2")
        
    If Np2 = 0 And Np3 = 0 Then 'And NoteP < Now Then
        tty = tty + 1
        'NoteP = Now + TimeSerial(0, 0, 5)
        'Shell "C:\WINDOZE\Notepad2.exe "+GetSpecialfolder(CSIDL_PERSONAL)+"\Andy Latest.txt", vbNormalFocus
            
        Shell "D:\VB6\VB-NT\NotePad instr top LF 2abbrev txt.exe", vbNormalFocus
        'MsgBox "s  --- " + Str(tty)
    End If
Sluty3 = 0
End If
    
    


If Sluty = 1 Then

ExeGo = False
    
ash$ = GetActiveWindowTitle(False)
   
Wie$ = " - Windows Internet Explorer"
Gcw$ = GetWindowTitle(GetForegroundWindow)

Call IEURL
'=url$

'wie$  gcw$

Np2 = FindWindow("TfmPictureView", vbNullString)
Np3 = FindWindow("TfmBrowser", vbNullString)

'If np2 <> GetForegroundWindow Then np2 = 0

If Np2 > 0 Then
    yy = GetWindowState(Np2)
    If yy = -1 Then
        ShowWindow Np3, SW_MINIMIZE
        ShowWindow Np2, SW_MINIMIZE
    End If

    If yy = vbMinimized Then
        ShowWindow Np3, SW_MAXIMIZE
        ShowWindow Np2, SW_MAXIMIZE
    End If
    Sluty = 0
    Exit Sub
        
End If
    

Sluty = 0
If dks = 1 And 1 = 2 Then
  
        dse = 0
        dse = 1
        On Local Error GoTo ers3
        'AppActivate "Message Rules", 0
        On Local Error GoTo 0
  
        If dse = 0 Then
            SendKeys "%l", True
            SendKeys "^a", True
            SendKeys "^c", True
            On Local Error GoTo ers3
            AppActivate "Notepad", 0
            On Local Error GoTo 0
            If dse = 1 Then MsgBox "Must open Notepad"
            
            If dse = 0 Then
                SendKeys "+{insert}", True
                SendKeys "--------------------------{enter}", True
                SendKeys "--------------------------{enter}", True
                SendKeys "--------------------------{enter}", True
                On Local Error GoTo ers3
                AppActivate "Message Rules", 0
                On Local Error GoTo 0
                SendKeys "{tab}", True
                SendKeys "{tab}", True
                SendKeys "{tab}", True
                SendKeys "{tab}", True
                SendKeys "{down}", True
            End If
            
        End If
    End If

    'Call CapsLockOFF
        
    Sleep 500
    DoEvents
    
    Dim abytBuffer(0 To 255) As Byte
    GetKeyboardState abytBuffer(0)

    CapsStat = abytBuffer(vbKeyCapital) > 0
    If CapsStat = True Then MsgBox "Caps Lock On": Exit Sub


        '------------------------------------
        'CheckCopyPaste
        
        'Notepad2
       
        If FindWindow("Notepad2", vbNullString) = GetForegroundWindow Then
            If Menu.CheckCopyPaste.Value = vbChecked Then
                'Gooze1 = Now + TimeSerial(0, 1, 0)
                SendKeys "^C", 0
                Exit Sub
            End If
        End If
        

        '------------------------------------
        'Indent
        If Menu.CheckIndent.Value = vbChecked Then
            Gooze1 = Now + TimeSerial(0, 1, 0)
            SendKeys "+{tab}{down}", 0
            Exit Sub
        End If
        
        '------------------------------------
        'IndentOut
        If Menu.CheckIndentOut.Value = vbChecked Then
            Gooze2 = Now + TimeSerial(0, 1, 0)
            SendKeys "{tab}{down}{home}", 0
            Exit Sub
        End If
        
        '------------------------------------
        'Publisher
        If Menu.CheckPublisher.Value = vbChecked Then
            Gooze5 = Now + TimeSerial(0, 2, 0)
            SendKeys "%e", True
            SendKeys "a", True
            SendKeys "{enter}", True
            Exit Sub
        End If
        
        '------------------------------------
        'IndentEmailOut
        If Menu.CheckIndentEmailOut.Value = vbChecked Then
            Gooze6 = Now + TimeSerial(0, 1, 0)
            SendKeys ">> {down}{home}", 0
            Exit Sub
        End If
        
       

          
        dse = 1
        ash2$ = GetActiveWindowTitle(False)
        test_hwnd = GetActiveWindow(False)
        
        
        
        
        'https://signin.ebay.co.uk/ws/eBayISAPI.dll?SignIn&UsingSSL=1&pUserId=&co_partnerId=2&siteid=3&ru=http%3A%2F%2Fmy.ebay.co.uk%2Fws%2FeBayISAPI.dll%3FMyEbayBeta%26MyeBay%3D%26guest%3D1&pageType=3984
        'Normal Login
        dh$ = "https://signin.ebay.co.uk"
        If InStr(URL$, dh$) > 0 And ExeGo = False Then
            'SendKeys "matt.lan@btinternet.com{tab}"
            SendKeys ""
            ExeGo = True
        End If
        
        dh$ = "https://orders.ebuyer.com"
        If InStr(URL$, dh$) > 0 And ExeGo = False Then
            'SendKeys "matt.lan@btinternet.com{tab}"
            SendKeys ""
            ExeGo = True
        End If
        
        
        '2nd Login
        'https://static.ebuyer.com/customer/account/index.html?action=bG9naW4=
        dh$ = "https://static.ebuyer.com/customer/account/index.html?action="
        If InStr(URL$, dh$) > 0 And ExeGo = False Then
            'SendKeys "matt.lan@btinternet.com{tab}"
            SendKeys ""
            ExeGo = True
        End If
        
        
  
   
  
'--------------------------------------------------------------------------------------------------------------------------
        
        xsop = 0
        '------------------------------------
        If Menu.CheckMumBT.Value = vbChecked Then
            xxz = 0
            If InStr(Gcw$, "Login - BT Yahoo!") > 0 Then xxz = 1
            If xxz = 1 Then
                xsop = 1
                SendKeys "linda.lancaster@****{tab}", 0
                SendKeys "{enter}", 0
                ExeGo = True
                
            End If
        End If
      
        '------------------------------------
        If Menu.CheckEddieBT.Value = vbChecked Then
            xxz = 0
            If InStr(Gcw$, "Login - BT Yahoo!") > 0 Then xxz = 1
            If xxz = 1 Then
                xsop = 1
                SendKeys "edward.lancaster@****{tab}", 0
                SendKeys "{enter}", 0
                ExeGo = True
            
            End If
        End If
      
        'MINE
        '------------------------------------
        If xsop = 0 Then
            xxz = 0
            If InStr(Gcw$, "Login - BT Yahoo!") > 0 Then xxz = 1
            If xxz = 1 Then
                xsop = 1
                SendKeys ""
                ExeGo = True
                
            End If
        End If
            
        'https://bt.edit.client.yahoo.com/bt/create_sub?.partner=bt-1&.master_login=matthew.lancaster%40btopenworld.com&.done=%2Faccounts&.scrumb=sXi6wDhX1Gq&.intl=uk
        If xsop = 0 Then
            xxz = 0
            If InStr(URL$, "https://bt.edit.client.yahoo.com") > 0 Then xxz = 1
            If xxz = 1 Then
                xsop = 1
                SendKeys ""
                ExeGo = True
                
            End If
        End If
            
        xxz = 0
        If InStr(Gcw$, "Verify Password") > 0 Then xxz = 1
        If xxz = 1 Then
            SendKeys ""
            ExeGo = True
        
            dse = 1
        End If
        
        
        
  
  
  
  
  
        '------------------------------------
        'Google Remove URL
        DF$ = "Remove URL"
  
        If InStr(ash$, DF$) > 0 Then
            SendKeys "matt.lan@btinternet.com{tab}", 0
            SendKeys "{enter}", 0
            ExeGo = True
        End If
        
        '------------------------------------
        'Google Alerts
        DF$ = "https://www.google.com/alerts/signin?hl=en"
  
        If InStr(URL$, DF$) > 0 And ExeGo = False Then
            SendKeys "rub.rim@****{tab}", 0
            SendKeys "{enter}", 0
            ExeGo = True
            
        End If
       
        '------------------------------------
        'Google Calendar
        DF$ = "https://www.google.com/accounts/"
  
        If InStr(URL$, DF$) > 0 And ExeGo = False Then
            'SendKeys "rub.rim@googlemail.com{tab}", 0
            SendKeys "{enter}", 0
            ExeGo = True
            
        End If
       
        '------------------------------------
        DF$ = "http://www.facebook.com/login.php"
        DF$ = "facebook.com/login.php"
        'https://login.facebook.com/login.php?login_attempt=1
  
        If InStr(URL$, DF$) > 0 And ExeGo = False Then
            SendKeys "{enter}", 0
            ExeGo = True
        End If
       
       'https://www.sainsburys.co.uk/sol/my_account/global_login.jsp?bmUID=1263461388791
        DF$ = "www.sainsburys.co.uk"
  
        If InStr(URL$, DF$) > 0 And ExeGo = False Then
            SendKeys "", 0
            ExeGo = True
        End If
       
       
       
       
       
        '------------------------------------
        'Right Moves
        
        DF$ = "http://www.rightmove.co.uk/action/ContactAgentAction"
  
        If InStr(URL$, DF$) > 0 And ExeGo = False Then
            SendKeys "Matthew Lancaster{tab}", 0
            SendKeys "matt.lan@btinternet.com{tab}", 0
            SendKeys "{tab}", 0
            SendKeys "East Sussex.{tab}", 0
            SendKeys "{tab}", 0
            ExeGo = True
        End If
       
        '------------------------------------
        'Latest Homes
        
        DF$ = "http://latesthomes.co.uk/emailoffice.aspx"
  
        If InStr(URL$, DF$) > 0 And ExeGo = False Then
            SendKeys "Matthew Lancaster{tab}", 0
            SendKeys "matt.lan@btinternet.com{tab}", 0
            SendKeys "{tab}", 0
            ExeGo = True
        End If
       
        

       
       
        '------------------------------------
        '
        DF$ = "http://www.homemove.org.uk"
        
        If InStr(URL$, DF$) > 0 And ExeGo = False Then
            SendKeys "{TAB}{TAB}{TAB}{DOWN 10}{TAB}{TAB}{ENTER}", True
            A = A
            ExeGo = True
        End If
       
        '------------------------------------
        'BT.Com Bill
        DF$ = "BT.com | Personal"
  
        If InStr(ash$, DF$) > 0 Then
            SendKeys "matthew.lancaster@btopenworld.com{tab}", 0
            SendKeys "", 0
            ExeGo = True
            
        End If
       
       
       

       
       
       
       '------------------------------------
       'Outlook E Copy to Folder Cool
        
       'weh = FindWindowPart("Outlook Express")

        weh = FindWindow(vbNullString, "Find Message")
        
        ash22$ = GetWindowTitle(weh)

        If weh = GetForegroundWindow Then fgw = 1 Else fgw = 0

        'If InStr(ash$, "Find Message") = 0 Then ash$ = ""

        If weh > 0 Then wnd2 = GetWindowState(weh)
        
        If weh = GetForegroundWindow Then fgw = 1 Else fgw = 0
        If wnd2 <> vbMinimized And fgw = 1 Then
            dse = 0
            On Local Error GoTo ers3
            AppActivate ash22$, 0
            On Local Error GoTo 0
  
            If dse = 0 Then
                SendKeys "%ep", True
                SendKeys "{enter}", True
                SendKeys "{up}", True
                ExeGo = True
                
            End If
        End If
       '------------------------------------
       'Outlook E Cool Folder Forward
        
       'weh = FindWindowPart("Outlook Express")

        weh = FindWindow("Outlook Express Browser Class", vbNullString)
         
        ash22$ = GetWindowTitle(weh)

        If weh = GetForegroundWindow Then fgw = 1 Else fgw = 0

        'If InStr(ash$, "Cool - Outlook Express - Main Identity") = 0 Then ash$ = ""


        If weh > 0 Then wnd2 = GetWindowState(weh)
        
        
        If weh = GetForegroundWindow Then fgw = 1 Else fgw = 0
        If wnd2 <> vbMinimized And fgw = 1 Then
            dse = 0
            On Local Error GoTo ers3
            AppActivate ash22$, 0
            On Local Error GoTo 0
  
            If dse = 0 Then
                SendKeys "^f", True
                SendKeys "matt.lan@btinternet.com", True
                SendKeys "****", True
                SendKeys "{up}", True
                ExeGo = True
            
            End If
        End If
       '------------------------------------
       'Outlook E InBox Spam
        
       'weh = FindWindowPart("Outlook Express")

        weh = FindWindow("Outlook Express Browser Class", vbNullString)
         
        ash22$ = GetWindowTitle(weh)

        'If InStr(ash$, "Deleted Items - Outlook Express - Main Identity") = 0 Then ash$ = ""

       'If weh > 0 Then weh2 = FindWindow(vbNullString, "Deleted Items - Outlook Express - Main Identity")

        If weh > 0 Then wnd2 = GetWindowState(weh)
        
        
        If weh = GetForegroundWindow Then fgw = 1 Else fgw = 0
        If wnd2 <> vbMinimized And fgw = 1 Then
            dse = 0
            On Local Error GoTo ers3
            AppActivate ash22$, 0
            On Local Error GoTo 0
  
            If dse = 0 Then
                SendKeys "****", True
                SendKeys "y", True
                ExeGo = True
                
            End If
        End If
        '------------------------------------
        'Outlook_Express_Philo
        dse = 1
        
        
        
        If wnd2 <> vbMinimized And fgw = 1 Then
        
            On Local Error Resume Next
            AppActivate "Alt.Philo - Outlook Express", 0
            qq1 = Err.Number
            Err.Clear
            AppActivate "00 Alt.SZ - Outlook Express", 0
            qq3 = Err.Number
        
            On Local Error GoTo 0
        
            If qq1 = 0 Or qq2 = 0 Or qq3 = 0 Then dse = 0
            If rt = 32000 Then dse = 1
           
  
            If dse = 0 Then
                SendKeys "^a", 0
                SendKeys "^c", 0
                ExeGo = True
            
            End If
        End If
        
        
        
        '------------------------------------
        'Vodafone
        easy5 = False
        DF$ = "vodafone"
        If InStr(ash$, DF$) > 0 Then easy5 = True
        DF$ = "Vodafone Log in"
        If InStr(ash$, DF$) > 0 Then easy5 = True
        DF$ = "Vodafone Login"
        If InStr(ash$, DF$) > 0 Then easy5 = True
        DF$ = "Vodafone UK - Login"
        If InStr(ash$, DF$) > 0 Then easy5 = True
        
        If easy5 = True Then
            SendKeys "****", True
            SendKeys "{tab}", True
            Sleep 100
            SendKeys "{tab}", True
            Sleep 100
            SendKeys "", True
            ExeGo = True
        End If
      
        
        '------------------------------------
        'Cpc
  
        If InStr(Gcw$, "CPC online") Then
  
            SendKeys "MatthewL2005", True
            SendKeys "{tab}", 0
            
            SendKeys "****"
            ExeGo = True
            
            Exit Sub
        End If
      
      
      'MsgBox ":"

        '------------------------------------
        'monday lottery
        DF$ = "Monday - The Charities Lottery"
        
        If InStr(ash$, DF$) > 0 Then
            SendKeys "lynlan{tab}", 0
            SendKeys "{enter}", 0
            ExeGo = True
            
        End If
          
        
          
        '------------------------------------
        'Paypal
        
        'active
        DF$ = "Welcome - Paypal"
        
        If InStr(ash$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
        End If
        
        'active
        DF$ = "Complete Your eBay Payment - PayPal"
        
        If InStr(ash$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
        End If
        
        'active
        DF$ = "Login - PayPal"
        
        If InStr(ash$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
            
        End If
        
        DF$ = "PayPal - Send Money"
        
        If InStr(ash$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
            
        End If
          
        DF$ = "PayPal - Welcome"
        
        If InStr(ash$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
            
        End If
         
        DF$ = "PayPal - Login"
        
        If InStr(ash$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
            
        End If
          
        DF$ = "PayPal - Password"
        
        If InStr(ash$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
            
        End If
          
        DF$ = "Choose a payment method - PayPal"
        
        If InStr(ash$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
            
        End If
          
        DF$ = "https://www.paypal.com"
        If InStr(URL$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
        End If
        
        DF$ = "https://www.securesuite.co.uk/"
        If InStr(URL$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
        End If
        
        'https://www.sainsburys.co.uk/groceries/index.jsp?bmUID=1263145117007
        DF$ = "https://www.sainsburys.co.uk/groceries/index.jsp"
        If InStr(URL$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
        End If
          
        'https://www.royalmail.com/portal/rm/ppmpaymentfull?pageId=ppmpayment_threed_sec&_requestid=37950
  
        DF$ = "https://www.royalmail.com/portal/rm/ppmpaymentfull"
        If InStr(URL$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
        End If
          
          
        '------------------------------------
        DF$ = "https://www.amazon"
  
        If InStr(URL$, DF$) > 0 And ExeGo = False Then
            'SendKeys "rub.rim@****{tab}", 0
            SendKeys "", 0
            ExeGo = True
            
        End If
       
        
                    
        '------------------------------------
        DF$ = "http://bthomehub.home"
        If InStr(URL$, DF$) > 0 And ExeGo = False Then
            SendKeys "", 0
            ExeGo = True
        End If
       
        
          
        
        '------------------------------------
        'Andy's Password for .RAR
          
        DF$ = "Password"
        
        If InStr(ash$, DF$) > 0 Then
            SendKeys "", 0
            ExeGo = True
            
        End If
          
        
        '------------------------------------
        'maplin
        DF$ = "maplin.co.uk"
          
        If InStr(URL$, DF$) > 0 Then
            SendKeys "matt.lan@btinternet.com", True
            SendKeys "{TAB}", True
            Sleep 200
            SendKeys "", True
            ExeGo = True
            
        End If
          
        
        'Credit card
        'df$ = "Alliance & Leicester Card Servicing Europe: Online Banking for UK Customers -"
        
        
        dse = 1
        DF$ = "Alliance & Leicester Card Servicing"
        dh$ = "https://www.bankcardservices.co.uk"
        
        If InStr(ash$, DF$) > 0 And InStr(URL$, dh$) > 0 Then dse = 0
        If dse = 0 And Menu.CheckMumAL.Value = vbUnchecked Then
            'SendKeys "{tab}{tab}{tab}{tab}{tab}{tab}{tab}", 0
            SendKeys "****{tab}", True
            SendKeys "", True
            ExeGo = True
            
        End If
        'Mum hasnt got credit Card Setup
        If dse = 0 And Menu.CheckMumAL.Value = vbChecked Then
            'SendKeys "{tab}{tab}{tab}{tab}{tab}{tab}{tab}", 0
            SendKeys "****{tab}", True
            SendKeys "", True
            ExeGo = True
            
        End If
        
        
    
      
        'A&L Debit Card Number For Payments-----------
        'Paying in Transfer Money
'        df$ = "https://www.bankcardservices.co.uk/NASApp/"
        DF$ = "Alliance & Leicester Card Servicing"
        dh$ = "type=payOnline"
        
        If InStr(ash$, DF$) > 0 And InStr(URL$, dh$) > 0 Then
        If Menu.CheckMumAL.Value = vbUnchecked Then
                'SendKeys "****{tab}", 0
                
                SendKeys "{tab}", 0
                SendKeys "{tab}", 0
                SendKeys "{tab}", 0
                SendKeys "{tab}", 0
                ExeGo = True
                
            End If
        End If
        
        If InStr(ash$, DF$) > 0 And InStr(URL$, dh$) > 0 Then
            If Menu.CheckMumAL.Value = vbChecked Then
                SendKeys "{tab}", 0
                SendKeys "{tab}", 0
                SendKeys "{tab}", 0
                SendKeys "{tab}", 0
                ExeGo = True
            
            End If
        End If
        
        'visa password extra check
        '

        DF$ = "https://www.bankcardservices.co.uk/NASApp/NetAccessXX/EFormSubmitProcess"
        dh$ = "type=payOnline"
        
        If InStr(ash$, DF$) > 0 And InStr(URL$, dh$) > 0 Then
        If Menu.CheckMumAL.Value = vbUnchecked Then
                'SendKeys "{tab}", 0
                SendKeys "{tab}{tab}", 0
                ExeGo = True
                
            End If
        End If
          
        
        DF$ = "https://www.securesuite.co.uk/mbna/tdsecure"
        dh$ = "type=payOnline"
        
        If InStr(ash$, DF$) > 0 And InStr(URL$, dh$) > 0 Then
                'SendKeys "{tab}", 0
                SendKeys "", 0
                ExeGo = True
        End If
          
          

        'digital promo
        DF$ = "Sage Pay"
        dh$ = "https://live.sagepay.com/gateway/service/authentication"
        
        If InStr(ash$, DF$) > 0 And InStr(URL$, dh$) > 0 Then
                'SendKeys "{tab}", 0
                SendKeys "{tab}{tab}", 0
                ExeGo = True
        End If
          

        'visa password extra check
        '

        DF$ = "        'visa password extra check"
        '
        dh$ = "type=payOnline"
        
        If InStr(ash$, DF$) > 0 And InStr(URL$, dh$) > 0 Then
        If Menu.CheckMumAL.Value = vbUnchecked Then
                'SendKeys "{tab}", 0
                SendKeys "{tab}{tab}", 0
                ExeGo = True
                
            End If
        End If
          
          
        'www%2Eplanet%2Dsource%2Dcode
        DF$ = "Sign In"
        dh1$ = "www%2Eplanet%2Dsource%2Dcode"
        If InStr(ash$, DF$) > 0 And InStr(URL$, dh1$) > 0 Then
                SendKeys "matt.lan@btinternet.com{tab}", True
                ExeGo = True
    End If
          
        'http://latesthomes.co.uk/emailalertsignup.aspx
        DF$ = "Latest Homes"
        dh1$ = "emailalertsignup"
        If InStr(ash$, DF$) > 0 And InStr(URL$, dh1$) > 0 Then
                SendKeys "matt.lan@btinternet.com{tab}", True
                ExeGo = True
        End If
          
          
          
        'Mine Debit
        dse = 1
        
        DF$ = "Alliance&Leicester - Online Banking"
        dh1$ = "login/PM4point1.asp"
        dh2$ = "alliance-leicester.co.uk/index.asp"
        If InStr(ash$, DF$) > 0 And InStr(URL$, dh2$) > 0 Then
            If Menu.CheckMumAL.Value = vbUnchecked Then
                'SendKeys "", True
                'New Type Code
                SendKeys "", True
                ExeGo = True
        End If
        End If
            
        If InStr(ash$, DF$) > 0 And InStr(URL$, dh1$) > 0 Then
            If Menu.CheckMumAL.Value = vbUnchecked Then
                SendKeys "", True
                ExeGo = True
            End If
        End If
        
        If InStr(ash$, DF$) > 0 And InStr(URL$, dh2$) > 0 Then
            If Menu.CheckMumAL.Value = vbChecked Then
                SendKeys "", True
                ExeGo = True
            End If
        End If
        
        If InStr(ash$, DF$) > 0 And InStr(URL$, dh2$) > 0 Then
            If Menu.CheckMumAL.Value = vbChecked Then
                SendKeys "", True
                ExeGo = True
            End If
        End If
        
        
          
        '------------------------------------
        
        DF$ = "The Co-operative Bank p.l.c"
          
        If InStr(ash$, DF$) > 0 Then
            SendKeys "{tab}", 0
            SendKeys "{tab}{tab}", 0
            SendKeys "", 0
            ExeGo = True
        End If
        
          
        '------------------------------------
        'Rswww Components
        
        DF$ = "rswww.com"
        
        If InStr(ash$, DF$) > 0 Then
            SendKeys "matthew2005{tab}", True
            SendKeys "", True
            
            MsgBox "matthew2005"
            ExeGo = True
        End If
        
        'dse = 0
        'On Local Error GoTo ers3
        'AppActivate "http://rswww.com homepage", 0
        'On Local Error GoTo 0
        
        'If dse = 0 Then
        '    SendKeys "matthew2005{tab}", True
        '    SendKeys "ditchit{enter}", True
        '
        'End If
          
        'Please log-in - W
        
        DF$ = "Please log-in - W"
          
        If InStr(ash$, DF$) > 0 Then
            SendKeys "matthew2005{tab}", True
            SendKeys "", True
            ExeGo = True
            
            'MsgBox "matthew2005 - ditchit"
        End If
                
        DF$ = "Member log-in - Win"
          
        If InStr(ash$, DF$) > 0 Then
            SendKeys "matthew2005{tab}", True
            SendKeys "", True
            ExeGo = True
            
            '[MsgBox "matthew2005 - ditchit"
        End If
                
        If frontpagepid > 0 Then
            dse = 1
            fronttid = frontpage(0)
            
            DF$ = "Name and Password Req"
            If InStr(ash$, DF$) > 0 Then dse = 0

            If fronttid = 1 Then
                If dse = 0 Then
                    SendKeys "+{tab}", True
                    SendKeys "matthewlan{tab}", True
                    SendKeys "", True
                    ExeGo = True
                    
                End If
            End If
            
            If fronttid = 2 Then
                If dse = 0 Then
                    SendKeys "+{tab}", True
                    SendKeys "matt.lan@btinternet.com{tab}", True
                    SendKeys "", True
                    ExeGo = True
            
                    
                End If
            End If
        End If
          
        
        dse5 = 0
        DF$ = "Yahoo! Login"
        If InStr(ash$, DF$) > 0 Then dse5 = 1

       DF$ = "Welcome, My!"
        If InStr(ash$, DF$) > 0 Then dse5 = 1
        
        DF$ = "Sign in to Yahoo!"
        If InStr(ash$, DF$) > 0 Then dse5 = 1
          
        'Current 08-02-2006
        DF$ = "My Yahoo!"
        If InStr(ash$, DF$) > 0 Then dse5 = 1
        
        'Current 08-02-2006
        DF$ = "Sign In - My Yahoo!"
        If InStr(ash$, DF$) > 0 Then dse5 = 1
         
        'Current 13-02-2006
        DF$ = "Sign In - Yahoo!"
        If InStr(ash$, DF$) > 0 Then dse5 = 1
        
        
        If dse5 = 1 And testnomore = 5 Then
            dse = 0
            
            yahoo = 0
            
            Load Yahoo_Login
            Yahoo_Login.Show
            Yahoo_Login.SetFocus
              
            
            Do
                DoEvents
            Loop Until yahoo <> 0
            
            
            If yahoo = 1 Then
                If dse = 0 Then
                    SendKeys "mattlancaster2000{tab}", 0
                    SendKeys "", 0
                    ExeGo = True
                    
                End If
            End If
            If yahoo = 2 Then
                If dse = 0 Then
                    SendKeys "brainbashing{tab}", 0
                    SendKeys "", 0
                    ExeGo = True
            
                End If
            End If
            If yahoo = 3 Then
                If dse = 0 Then
                    SendKeys "MattLan2008{tab}", 0
                    SendKeys "", 0
                    ExeGo = True
                End If
            End If
            If yahoo = 4 Then
                If dse = 0 Then
                    SendKeys "voidman2004xx{tab}", 0
                    SendKeys "", 0
                    ExeGo = True
            
                End If
            End If
            If yahoo = 5 Then
                If dse = 0 Then
                    SendKeys "", 0
                    ExeGo = True
                    
                End If
            End If
            
            If yahoo = 0 Then dse = 0
              
        End If
        
          
        DF$ = "Verify Password"
        If InStr(ash$, DF$) > 0 And ExeGo = False Then
            'SendKeys "kingking21{enter}", 0
            SendKeys "", True
            ExeGo = True
        End If
        
        
        DF$ = "login.yahoo.com"
        If InStr(URL$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
        End If

        
        
        DF$ = "Sign In - Bt Yahoo"
        If InStr(ash$, DF$) > 0 And ExeGo = False Then
            'SendKeys "kingking21{enter}", 0
            SendKeys ""
            ExeGo = True
        
        End If
          
        '------------------------------------
        'ebay
        'there is two of these one doesnt put the user name
        DF$ = "Please Sign In:"
        
        If InStr(ash$, DF$) > 0 And ExeGo = False Then
            'SendKeys "Voidman2004{tab}", 0
            SendKeys ""
            ExeGo = True
            
            'MSComm1.Output = "kingkingmagic48"
            
        End If
          
        '------------------------------------
        'ebay ' dont think used anymore replaced one below
        DF$ = "Sign In: My Fav"
        
        If InStr(ash$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
            
        End If
        
        '------------------------------------
        'ebay ' Current
        DF$ = "My Ebay"
        
        If InStr(ash$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
            
        End If
          
        '------------------------------------
        'ebay ' Current
        DF$ = "Sign in or register to continue"
        
        If InStr(ash$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
            
        End If
          
        '------------------------------------
        'ebay ' Current
        DF$ = "https://signin.ebay.co.uk"
        
        If InStr(URL$, DF$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
            
        End If
          
        '------------------------------------
        'ebay
        If InStr(ash$, "Sign In: Favourite Searches") > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
        End If
        '------------------------------------
        'ebay - Current on click to watch without pre log in
        
        
        'ash$ = GetActiveWindowTitle(False)
        'If InStr(ash$, "Sign In") > 0 And dse2 = 0 Then
        '    SendKeys "kingkingmagic48"
            
        'End If
        
        If InStr(ash$, "Welcome to eBay") > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
            
        End If

        
              
        If InStr(ash$, "MyPlesk.com") = 1 Then
            SendKeys "matthewlan.com{tab}"
            SendKeys ""
            ExeGo = True
            
        End If
          
        'not used changed maybe
        dse = 1
        'ash$ = GetActiveWindowTitle(False)
        If InStr(ash$, "Plesk 8.") = 1 Then dse = 0
            
        If dse = 0 Then
            SendKeys "matthewlan.com{tab}"
            SendKeys ""
            ExeGo = True
            'SendKeys ""
'            SendKeys ""
            
        End If
          
        dse = 1
        'ash$ = GetActiveWindowTitle(False)
        If InStr(ash$, "MyHostNed") = 1 Then dse = 0
            
        If dse = 0 Then
            SendKeys "****{tab}"
            ExeGo = True
            'SendKeys "{tab}"
           ' SendKeys ""
            'MsgBox ""
            'htucgo
            
        End If
            

        '#TESCO
        dh$ = "https://secure.tesco.com/"
        If InStr(URL$, dh$) > 0 And ExeGo = False Then
            SendKeys ""
           ' SENDKEYS0
            
            ExeGo = True
        End If
        
        
        '#Credit CARD MBNA
        dh$ = "https://orders.ebuyer.com/customer/shopping/index.html?action="
        If InStr(URL$, dh$) > 0 And ExeGo = False Then
            SendKeys ""
            ExeGo = True
        End If
        
        'Normal Login
        dh$ = "http://accounts.ebuyer.com/customer/orders"
        If InStr(URL$, dh$) > 0 And ExeGo = False Then
            'SendKeys "matt.lan@btinternet.com{tab}"
            SendKeys ""
            ExeGo = True
        End If
        
        '
        
      
      
        If 1 = 2 Then
            If processid25ssam > 0 Then
                SendKeys "please god{enter}", True
                SendKeys "please giveall{enter}", True
                SendKeys "please killall{enter}", True
                SendKeys "please open{enter}", True
                SendKeys "please fly{enter}", True
                SendKeys "please ghost{enter}", True
                SendKeys "please invisible{enter}", True
                SendKeys "please tellall{enter}", True
                SendKeys "please refresh{enter}", True
                ExeGo = True
                
            End If
        End If
      
          'Timer
        If ExeGo = False Then ATidalX.ShellKeyLoad.Enabled = True
      
End If
      
     
Exit Sub
    
errorz1:
esx = 1
Resume Next

ers3:
dse = 1
Resume Next


End Sub

Public Sub ErrorTrap(A1, A2$, a3)

If A1 = 0 Then Exit Sub

FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\ErrorLog.txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4

fe4 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #fe4
Print #fe4, "----" + Format$(Now, "DDD DD,MMM,YYYY HH:MM:SS")
Print #fe4, A1
Print #fe4, A2$
Print #fe4, a3
Print #fe4, "------------------"
Close fe4

End Sub

Public Sub UpDate5(tl)


tr$ = "Program Mostly Written By Matthew.Lancaster@btinternet.com 2005 " + vbCrLf + "Version "
tr$ = tr$ + Trim$(str$(App.Major)) + "." + Trim$(str$(App.Minor)) + "." + Trim$(str$(App.Revision)) + vbCrLf
tr$ = tr$ + "Located in " + App.Path + "\" + App.EXEName + ".EXE" + vbCrLf

MessageBox = MessageBox + tr$




'UpDateText.List1.Clear

tr$ = "Click / Toggle on the Time to Speed it Up, Stop it or Double Click to reset to current time." + vbCrLf
MessageBox = MessageBox + tr$
tr$ = "The Tide is only speeded up by 50% compared to the Moon." + vbCrLf
MessageBox = MessageBox + tr$
tr$ = "Click / Toggle on Next Moon Description to see Now an Next Description." + vbCrLf
MessageBox = MessageBox + tr$
tr$ = "Click / Toggle on the Date Of the Next Moon to see The Percent luminosity for Now & Next." + vbCrLf
MessageBox = MessageBox + tr$
tr$ = "Click on '100 % = Hide Tide / Full Moon' To get this Info Back" + vbCrLf
MessageBox = MessageBox + tr$
tr$ = "Click on Today an Tommorow for more Sunrise Info" + vbCrLf
MessageBox = MessageBox + tr$
tr$ = "DoD = Difference of Difference with the 100% Max Difference can be around Equinox" + vbCrLf
MessageBox = MessageBox + tr$
tr$ = "Click / Toggle on DoD to see when this and next Year's Equinox's are" + vbCrLf
MessageBox = MessageBox + tr$
tr$ = "Programs Longitude / Latitude = " + str$(welong) + ", " + str$(welat) + vbCrLf
MessageBox = MessageBox + tr$

'555 ROIDSGOLDRIM
tr$ = "Personalized For UnKnown Street POST CODE"
If CompName$ = "555ROIDSGOLDRIM" Then tr$ = "Personalized For Matt's Street"
If CompName$ = "MAGICRAM2EODUR" Then tr$ = "Personalized For Matt's Street"
If CompName$ = "EDDIESCIENTIST" Then tr$ = "Personalized For Eddie's Street"
If CompName$ = "NASA1234567890" Then tr$ = "Personalized For Marianne's Street"
If CompName$ = "MEACHELLE12345" Then tr$ = "Personalized For Meachelle's Street"

tr$ = tr$ + vbCrLf

MessageBox = MessageBox + tr$
tr$ = "Mouse Keyboard Counter That Keeps a Log of the Count" + vbCrLf
MessageBox = MessageBox + tr$
tr$ = "Mouse cursor Hiding Default 'On' can Switch Off When Playing games" + vbCrLf
MessageBox = MessageBox + tr$


MsgBox MessageBox

Exit Sub

tr$ = ""

tr$ = tr$ + "------------------------------"
tr$ = tr$ + vbCr + "------------------------------------------------------"

tr$ = tr$ + vbCr + "*** This Version Has Just Been Updated ***"

tr$ = tr$ + vbCr + "------------------------------------------------------"
tr$ = tr$ + vbCr + "------------------------------"

If tl = 1 Then
    
    UpDateText.Label3.Caption = tr$
    '555 ROIDSGOLDRIM
    If tl = 1 Then
        If CompName$ = "MAGICRAM2EODUR" Then
            Rabbit = Now + TimeSerial(0, 0, 2)
        Else
            Rabbit = Now + TimeSerial(0, 0, 10)
        End If
    End If
    If tl = 1 Then
        If CompName$ = "555ROIDSGOLDRIM" Then
            Rabbit = Now + TimeSerial(0, 0, 2)
        Else
            Rabbit = Now + TimeSerial(0, 0, 10)
        End If
    End If
End If
tr$ = ""
tr$ = tr$ + "------------------------------------------------------"
tr$ = tr$ + vbCr + "---------------------------------------------"
tr$ = tr$ + vbCr + "-----------------------------------"
tr$ = tr$ + vbCr + "------------------------"
tr$ = tr$ + vbCr + "-------------"
tr$ = tr$ + vbCr + "---"

If tl = 0 Then UpDateText.Label3.Caption = tr$

'If tl = 1 Then UpDateText.Show

tr$ = ""

End Sub

Sub UpDate66()

Exit Sub


qtu$ = "Tidal Info Ver 1.0."

ATidalX.Caption = "Tidal Info..."
ATidalX.Refresh

Dim weh As Long
Dim art As Long

Do
    puk = 0
    wef = App.Revision
'    Do
        weh = FindWindow(vbNullString, qtu$ + Trim(str(wef - 1)))
        
        If weh = 0 Then weh = FindWindow(vbNulString, "Tidal Information...")
        
        wef = wef - 1
        If weh > 0 Then
            
         'MsgBox "I"
   
            art = SendMessage(weh, WM_CLOSE, 0&, 0&)
            puk = 1
            If art > 0 Then MsgBox ("Probelms Removing Previous Instance's of this program that are running in memory now")
 '           xzz = 1
            End If
        
'    Loop Until wef < 80 Or xzz = 1
Loop Until puk = 0

'ATidalX.Caption = "Tidal Info Ver " + Trim$(Str$(App.Major)) + "." + Trim$(Str$(App.Minor)) + "." + Trim$(Str$(App.Revision))
ATidalX.Caption = "Tidal Info."


Dim Adate1 As Date, Adate2 As Date

Wet$ = App.EXEName
If Wet$ = "YTidal_Info2" Then Wet$ = "Tidal"
Wet$ = Wet$ + ".exe"
    
FileSpec = "C:\utils\" + Wet$
    
On Local Error Resume Next
Set F = FS.GetFile((FileSpec))
Adate1 = F.DateLastModified

If Err.Number > 0 Then Exit Sub
    


FileSpec = App.Path + "\" + Wet$
Set F = FS.GetFile((FileSpec))
Adate2 = F.DateLastModified

If Err.Number > 0 Then Exit Sub


If Adate1 <> Adate2 Then
    
    Wtime = Now + TimeSerial(0, 0, 2)
    Do
        fre5 = FreeFile
        Open "C:\utils\" + Wet$ For Output As #fre5
        Print #fre5, "S"
        Close #fre5
    Loop Until Err.Number = 0 Or Wtime < Now
    
    If Err.Number > 0 Then MsgBox "Please Exit the Other Tidal Info Program... Then Continue...": Cheese = 1

    If Err.Number = 0 Then

'Set F = fs.GetFile((FileSpec))
        FS.CopyFile App.Path + "\" + Wet$, "C:\utils\" + Wet$

        Shell "C:\utils\" + Wet$ + " update", vbNormalFocus
        
        End
    
    End If

End If



End Sub



Sub FindWinIExplorGet()

hwnd3 = FindWindowPartScan("Windows Internet Explorer")

ash$ = GetWindowTitle(hwnd3)

'http

hwnd3 = FindWindowPart("http")

ash$ = GetWindowTitle(hwnd3)

FindWinIExplor.txtDestCaption.Text = ash$
    
'Call FindWinIExplor.cmdSend_Click

rt = InStr(ash$, "- Windows Internet Explorer")

'ash$ = Mid$(ash$, 1, rt - 2)




End Sub





Private Sub object_NavigateComplete2( _
    ByVal pDisp As Object, _
    ByVal URL As Variant)
End Sub





Function FindWindowPartScan(title$) As Long

Dim ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long
'Find the first window
test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
    
FindWindowPartScan = False
    
Do While test_hwnd <> 0
    ash$ = GetWindowTitle(test_hwnd)
    ht = test_hwnd
    
    If InStr(ash$, title$) Then
    ash2$ = ""
        Do
            'test_hwnd = FindWindowEx(hWndForm, 0&, "ThunderTextBox", vbNullString)
            'test_hwnd = GetWindow(test_hwnd, GW_CHILD)
             
            'hWndChild = GetWindow(hWndChild, GW_HWNDNEXT)
            pit = test_hwnd
            test_hwnd = FindWindowEx(test_hwnd, pit, vbNullString, vbNullString)

            If (test_hwnd > 0) Then
                ash2$ = GetWindowTitle(test_hwnd)
            End If
            'test_hwnd = GetParent(test_hwnd)
            'test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
            test_hwnd = GetWindow(test_hwnd, GW_CHILD)
        Loop Until UCase$(Mid$(ash2$, 1, 4)) = UCase$("http") Or test_hwnd = 0
    
    End If
    
    
    
    test_hwnd = ht
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop

End Function


Function FindWindowPart(title$) As Long
Dim ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long
'Find the first window
test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
    
FindWindowPart = False
    
Do While test_hwnd <> 0
    ash$ = GetWindowTitle(test_hwnd)
    
    If InStr(ash$, title$) Then
        FindWindowPart = test_hwnd
        Exit Do
    End If
    
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
Loop

End Function







      Function EnumWinProc(ByVal lhWnd As Long, ByVal lParam As Long) _
         As Long
         
         
         Dim Retval As Long, ProcessID As Long, ThreadID As Long
         Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
         Dim WinClass As String, WinTitle As String

         Retval = GetClassName(lhWnd, WinClassBuf, 255)
         WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
         Retval = GetWindowText(lhWnd, WinTitleBuf, 255)
         WinTitle = StripNulls(WinTitleBuf)
         Print #1,
         Print #1, lhWnd
         Print #1, WinClass
         Print #1, WinTitle
         
         
         If InStr(UCase(WinTitle), "HTTP") > 0 Then MsgBox (WinTitle)
         If InStr(UCase(WinTitle), "BT YAHO") > 0 Then
         'MsgBox (WinTitle)
         A = 1
         End If
         If InStr(UCase(WinClass), "HTTP") > 0 Then MsgBox (WinClass)
         
         TopCount = TopCount + 1
         ' see the Windows Class and Title for each top level Window
'         Debug.Print "Top level Class = "; WinClass; ", Title = "; WinTitle
         ' Usually either enumerate Child or Thread Windows, not both.
         ' In this example, EnumThreadWindows may produce a very long list!
         Retval = EnumChildWindows(lhWnd, AddressOf EnumChildProc, lParam)
         ThreadID = GetWindowThreadProcessId(lhWnd, ProcessID)
         Retval = EnumThreadWindows(ThreadID, AddressOf EnumThreadProc, _
         lParam)
         EnumWinProc = True
      
       
      End Function

      Private Function EnumChildProc(ByVal lhWnd As Long, ByVal lParam As Long) _
         As Long
         Dim Retval As Long
         Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
         Dim WinClass As String, WinTitle As String
         Dim WinRect As RECT
         Dim WinWidth As Long, WinHeight As Long

         Retval = GetClassName(lhWnd, WinClassBuf, 255)
         WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
         Retval = GetWindowText(lhWnd, WinTitleBuf, 255)
         WinTitle = StripNulls(WinTitleBuf)
         
         hWndTextBox = FindWindowEx(lhWnd, 0&, WinClass, vbNullString)
         ash$ = GetWindowTitle(hWndTextBox)
         
         Print #1,
         Print #1, lhWnd
         Print #1, WinClass
         Print #1, WinTitle
         Print #1, hWndTextBox
         Print #1, ash$
         
         ChildCount = ChildCount + 1
         ' see the Windows Class and Title for each Child Window enumerated
         'Debug.Print "   Child Class = "; WinClass; ", Title = "; WinTitle
         If InStr(UCase(WinClass), "HTTP") Then MsgBox (WinClass)
         If InStr(UCase(WinTitle), "HTTP") Then MsgBox (WinTitle)
         If hWndTextBox <> 0 Then
         A = 1
         End If
         If WinClass <> "" Then
         'MsgBox (WinClass)
         A = 1
         End If
         If InStr(UCase(WinTitle), "HTTP") Then
         'MsgBox (WinTitle)
         A = 1
         End If
         
         ' You can find any type of Window by searching for its WinClass
         If WinClass = "ThunderTextBox" Then    ' TextBox Window
            
            'hWndTextBox = FindWindowEx(lhWnd, 0&, "ThunderTextBox", vbNullString)
            'ash$ = GetWindowTitle(hWndTextBox)
            'If ash$ <> "" Then
            'A = 1
            'End If
            'If InStr(UCase(ash$), "HTTP") Then MsgBox (ash$)
            
            Retval = GetWindowRect(lhWnd, WinRect)  ' get current size
            WinWidth = WinRect.Right - WinRect.Left ' keep current width
            WinHeight = (WinRect.Bottom - WinRect.Top) * 2 ' double height
            'RetVal = MoveWindow(lhWnd, 0, 0, WinWidth, WinHeight, True)
            EnumChildProc = False
         Else
            EnumChildProc = True
         End If
      End Function

      Function EnumThreadProc(ByVal lhWnd As Long, ByVal lParam As Long) _
         As Long
         Dim Retval As Long
         Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
         Dim WinClass As String, WinTitle As String
        

         Retval = GetClassName(lhWnd, WinClassBuf, 255)
         WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
         Retval = GetWindowText(lhWnd, WinTitleBuf, 255)
         WinTitle = StripNulls(WinTitleBuf)
         
         Print #1,
         Print #1, lhWnd
         Print #1, WinClass
         Print #1, WinTitle
         
         
         If InStr(UCase(WinTitle), "HTTP") > 0 Then MsgBox (WinTitle)
         ThreadCount = ThreadCount + 1
         ' see the Windows Class and Title for top level Window
'         Debug.Print "Thread Window Class = "; WinClass; ", Title = "; _
         WinTitle
         EnumThreadProc = True
      End Function

      Function StripNulls(OriginalStr As String) As String
         ' This removes the extra Nulls so String comparisons will work
         If (InStr(OriginalStr, Chr(0)) > 0) Then
            OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
         End If
         StripNulls = OriginalStr
      End Function


'Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
Private Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function




Public Sub Logger()
    
'   dw = frmIECtrl.LV1.ListItems.Count
   
'   If dw = 0 Then Exit Sub
   
'ae1 = xLocationName2$
'ae2 = xLocationURL2$
   
'If ae1 = "" And ae2 = "" Then Exit Sub
   
   
   
   
   
'Dim tmpIE As IE_Class
'Dim itmx As ListItem
'---------
xLocationName$ = ""
xLocationURL$ = ""
xLocationx1$ = ""
xLocationx2$ = ""
xLocationx3$ = ""
xLocationx4$ = ""
xLocationx5$ = ""
'---------
On Local Error Resume Next
'---------
If IEWin Is Nothing Then Exit Sub

'If IEWin.Count > 1 Then Stop
'---------

ww5 = 0
For rrt3 = 1 To IEWin.Count

Set tmpIE = IEWin.IE(rrt3)

If GetForegroundWindow = tmpIE.IEctl.hWnd Then ww5 = 1: Exit For

Next
If ww5 = 0 Then rrt3 = rrt3 - 1

If ww5 = 0 And InStr(ATidalX.Text1.Text, "URL Logger") <> 1 Then
    
    Set tmpIE = IEWin.IE(rrt3)
    
    Call Log2(rrt3)

    Exit Sub
End If


If InStr(ATidalX.Text1.Text, "URL Logger") = 1 Then

    For rrt3 = 1 To IEWin.Count

        Set tmpIE = IEWin.IE(rrt3)

        Call Log2(rrt3)

    Next

End If

End Sub

Public Sub Log2(rrt3)

Set tmpIE = IEWin.IE(rrt3)

'IEWin.IE.IEctl.ClientToWindow




'Weeee = tmpIE.IEctl.ClientToWindow



xLocationName = tmpIE.IEctl.LocationName
'xLocationURL = tmpIE.IEctl.Document
xLocationURL = tmpIE.IEctl.LocationURL
xLocationHWND = tmpIE.IEctl.hWnd
xLocationxBusy = tmpIE.IEctl.Busy
If xLocationxBusy = True Then xLocationxBusy = "True" Else xLocationxBusy = "False"
xLocationxFullProgPath = tmpIE.IEctl.FullName

ash$ = GetWindowTitle(CStr(xLocationHWND))

If (InStr(xLocationxFullProgPath, "\iexplore.exe") = 0 And _
    InStr(ash$, xLocationName) = 1) Or _
    (xLocationName <> ash$ And _
    UCase(Mid$(xLocationName, 1, 4)) = "HTTP") Then
    
    ash5 = InStr(ash$, " - Windows Internet Explorer") + 27
    ash51 = Len(ash$)
    If ash5 = ash51 And InStr(ash$, " - Windows Internet Explorer") > 0 Then
        ash2$ = Trim$(Mid$(ash$, 1, ash5 - 28))
    End If

    ash$ = ash2$

End If

'--------------
On Local Error GoTo 0
   
   
If ash$ + vbCrLf + xLocationURL = ATidalX.Text1.Text Then Exit Sub




        
      'egh$ = GetUserName
  
    FILENAME_IN_USE_CHECK = GetSpecialfolder(CSIDL_PERSONAL) + "\01 Palm HotSync\UrlLogg-long-" + GetComputerName + GetUserName + ".txt"
    FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
    Fr8 = FreeFile
    Open FILENAME_IN_USE_CHECK For Append As #Fr8
    Print #Fr8, "----" + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa ") + "---------------------"
    
    Print #Fr8, ash$
    Print #Fr8, xLocationName
    Print #Fr8, xLocationURL
    Print #Fr8, xLocationxFullProgPath
    Print #Fr8, "Busy = "; xLocationxBusy
    Close #Fr8
    
    FILENAME_IN_USE_CHECK = GetSpecialfolder(CSIDL_PERSONAL) + "\01 Palm HotSync\UrlLogg1-" + GetComputerName + GetUserName + ".txt"
    FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
    Fr8 = FreeFile
    Open FILENAME_IN_USE_CHECK For Append As #Fr8
    tui = 0
    ty5$ = "BT Yahoo! Home Page"
    If ty5$ = xLocationName Then tui = 1
    ty6$ = "Google Reader"
    If ty6$ = xLocationName Then tui = 1
    If ty5$ = ash$ Then tui = 1
    If ty6$ = ash$ Then tui = 1

    tc = 0
    TRem$ = Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa ")
    If TRem$ = Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa ") Then tc = 1
    If tui = 0 Then
       If tc = 0 Then Print #Fr8, "----" + Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa ") + "---------------------"
       Print #Fr8, ash$
       Print #Fr8, xLocationName
       Print #Fr8, xLocationURL
       Print #Fr8,
       Close #Fr8
    End If
Close #Fr8

ATidalX.Text1.Text = ash$ + vbCrLf + xLocationURL
ATidalX.Text1.Refresh

End Sub


Sub IEURL()
        
    Dim hWndForm        As Long
    Dim hWndTextBox     As Long
    Dim Retval As Long
    Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
    Dim WinClass As String, WinTitle As String, WinTitle2 As String

    ttg = 0
    hWndFormi = FindWindow("IEFrame", vbNullString)
    If hWndFormi = GetForegroundWindow Then ttg = 1
    hWndFormE = FindWindow("ExploreWClass", vbNullString)
    If hWndFormE = GetForegroundWindow Then ttg = 2
    If ttg = 1 Then hWndForm = hWndFormi
    If ttg = 2 Then hWndForm = hWndFormE
    
    URL$ = ""
    
    If ttg = 0 Or ttg = 2 Then
        Exit Sub
    End If
    
    
    
    Dim hWndIEChild As Long
    
    If ttg = 1 Then
        hWndIEChild = FindWindowEx(hWndForm, 0&, "WorkerW", vbNullString)
        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ReBarWindow32", vbNullString)
        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "Address Band Root", vbNullString)
        'hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ComboBoxEx32", vbNullString)
        'hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ComboBox", vbNullString)
        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "Edit", vbNullString)
    End If
    
    If ttg = 2 Then
        hWndIEChild = FindWindowEx(hWndForm, 0&, "WorkerW", vbNullString)
        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ReBarWindow32", vbNullString)
        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ComboBoxEx32", vbNullString)
        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ComboBox", vbNullString)
        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "Edit", vbNullString)
    End If
    
    
    
    If hWndIEChild <> 0 Then
        
    window_hwnd = hWndIEChild
    txtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, 0, 0)
        
    txtlen = SendMessage(window_hwnd, WM_GETTEXT, 255, ByVal WinTitleBuf)
    WinTitle = StripNulls(WinTitleBuf)
    ash$ = GetWindowTitle(hWndForm)
    isd = InStr(ash$, " - Windows Internet Explorer")
    If isd > 0 Then ash$ = Left$(ash$, isd - 1)
        
    If OldWinTitle <> WinTitle Or OldAsh <> ash$ Or 1 = 1 Then
        ash2$ = ash$
        WinTitle2 = WinTitle
        XX = 0
        
        If Trim(ash$) = "" And Trim(WinTitle2) = "" Then XX = 1
        If ash$ = "Windows Internet Explorer" And Trim(WinTitle2) = "" Then XX = 1
        If ash$ = "Windows Explorer" And Trim(WinTitle2) = "" Then XX = 1
        
        If XX = 0 Then
            'Call DisplayLogging("", ash2$, WinTitle2)
            URL$ = WinTitle2
        End If
    End If
    
    OldWinTitle = WinTitle
    OldAsh = ash$
    
End If
    
End Sub
Property Let WindowVisible(hWnd As Long, New_Value As Boolean)

    Call ShowWindow(hWnd, IIf(New_Value, SW_SHOW, SW_HIDE))

End Property

Property Get WindowVisible(hWnd As Long) As Boolean

    WindowVisible = (IsWindowVisible(hWnd) > 0)
  
End Property







Function SetKeys(Optional capslock As String)
' call SetKeys ("CAPSLOCK_OFF")


Dim keys As String
'Const VK_NUMLOCK = &H14



keys = Space$(256)
GetKeyboardState keys

Select Case capslock

Case "CAPSLOCK_ON"
keys = Left(keys, VK_CAPITAL) & Chr(1) _
& Right(keys, Len(keys) - (VK_CAPITAL + 1))
GoTo initiate

Case "CAPSLOCK_OFF"
keys = Left(keys, VK_CAPITAL) & Chr(0) _
& Right(keys, Len(keys) - (VK_CAPITAL + 1))
GoTo initiate
'-------------

Case "NUMLOCK_ON"
keys = Left(keys, VK_NUMLOCK) & Chr(0) _
& Right(keys, Len(keys) - (VK_NUMLOCK + 1))
GoTo initiate
'-------------

Case "SCROLL-LOCK_ON"
keys = Left(keys, VK_SCROLL) & Chr(0) _
& Right(keys, Len(keys) - (VK_SCROLL + 1))
SetKeyboardState keys
'-------------

Case "INSERT_ON"
keys = Left(keys, VK_INSERT) & Chr(0) _
& Right(keys, Len(keys) - (VK_INSERT + 1))
'VK_INSERT
GoTo initiate


Case "TOGGLE"
If &H1 = (Asc(Mid(keys, VK_CAPITAL + 1, 1)) And &H1) Then
keys = Left(keys, VK_CAPITAL) & Chr(0) _
& Right(keys, Len(keys) - (VK_CAPITAL + 1))
Else
keys = Left(keys, VK_CAPITAL) & Chr(1) _
& Right(keys, Len(keys) - (VK_CAPITAL + 1))
End If
GoTo initiate

Case Else

MsgBox "Not a value" & vbCrLf & _
"CAPSLOCK_ON" & vbCrLf & _
"CAPSLOCK_OFF" & vbCrLf & _
"OR " & "TOGGLE"

End Select

Exit Function

initiate:

SetKeyboardState keys

End Function





'Type OFSTRUCT
'  cBytes     As Byte
'  fFixedDisk As Byte
'  nErrCode   As Integer
'  Reserved1  As Integer
'  Reserved2  As Integer
'  szPathName As String * 128
'End Type

'Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long

'Const OF_SHARE_COMPAT = &H0
'Const OF_SHARE_EXCLUSIVE = &H10
'Const OF_SHARE_DENY_WRITE = &H20
'Const OF_SHARE_DENY_READ = &H30
'Const OF_SHARE_DENY_NONE = &H40

'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


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

Sub FileInUseDelay(TxT, Complain As Boolean)
        
    Dim QQ
    QQ = Now + TimeSerial(0, 8, 0)
    Do While FileInUse(TxT) = True Or QQ < Now
        DoEvents
        Sleep (20)
    Loop
    
    If QQ < Now And Complain = True Then
        MsgBox "Trouble File Stuck In Use " + vbCrLf + TxT + vbCrLf + "Longer than 8 min"
        Stop
    End If
End Sub



Sub Operating_System()
' Operating System

    Dim Enumerator As SWbemObjectSet
    Dim Object As SWbemObject
    Dim Item As ListItem

    On Error Resume Next
  


    
    MDIProcServ.WBEMServices.ExecNotificationQueryAsync MDIProcServ.eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
    Set Enumerator = MDIProcServ.WBEMServices.ExecQuery("Select * From Win32_Processor")
    
    If Err.Number = 91 Then
        MsgBox "Please CONNECT to Local or Remote computer first", vbOKOnly, "Connect please!"
        Exit Sub
    End If
    
    For Each Object In Enumerator
        With msfgProcessor
        cpudes = Object.Name
        End With
    Next

    t1 = InStr(cpudes, " ")
    t1 = InStr(t1 + 1, cpudes, " ")
    t1 = InStr(t1 + 1, cpudes, " ")

    cpudes = Mid$(cpudes, 1, t1) + Mid$(cpudes, InStrRev(cpudes, " ") + 1)

    
    MDIProcServ.WBEMServices.ExecNotificationQueryAsync MDIProcServ.eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
    Set Enumerator = MDIProcServ.WBEMServices.ExecQuery("Select * From Win32_OperatingSystem")
    
    If Err.Number = 91 Then
        MsgBox "Please CONNECT to Local or Remote computer first", vbOKOnly, "Connect please!"
        Exit Sub
    End If
    
    
      
        
    
    
    For Each Object In Enumerator
      With msfgOperatingSystem
      
        
        TTz1 = Object.LastBootUpTime
        TTz2 = Object.InstallDate
        TTz2 = Mid$(TTz2, 1, 4) + "-" + Mid$(TTz2, 5, 2) + "-" + Mid$(TTz2, 7, 2) + " " + Mid$(TTz2, 9, 2) + ":" + Mid$(TTz2, 11, 2) + ":" + Mid$(TTz2, 13, 2)
        TTz2 = DateValue(TTz2) + TimeValue(TTz2)
        TTz3 = Mid$(Object.Name, 1, InStr(Object.Name, "|") - 1) + " - Ver. " + Object.Version + " - SP" + Mid$(Object.CSDVersion, Len(Object.CSDVersion))
        'TTz3 = "Windows XP Pro" + " - Ver. " + Object.Version + " - SP" + Mid$(Object.CSDVersion, Len(Object.CSDVersion))
        TTz3 = "MS Win XP Pro" + " SP" + Mid$(Object.CSDVersion, Len(Object.CSDVersion))


'        msgbox TTz2

        Exit For
        
        .ColWidth(0) = 3000
        .ColWidth(1) = 3000
      
        .Row = 1
        .col = 0
        .Text = "BootDevice"
        .col = 1
        .Text = Object.BootDevice
        
        .Row = 2
        .col = 0
        .Text = "BuildNumber"
        .col = 1
        .Text = Object.BuildNumber
        
        .Row = 3
        .col = 0
        .Text = "BuildType"
        .col = 1
        .Text = Object.BuildType
        
        .Row = 4
        .col = 0
        .Text = "Caption"
        .col = 1
        .Text = Object.Caption
        
        .Row = 5
        .col = 0
        .Text = "CodeSet"
        .col = 1
        .Text = Object.CodeSet
        
        .Row = 6
        .col = 0
        .Text = "CountryCode"
        .col = 1
        .Text = Object.CountryCode
        
        .Row = 7
        .col = 0
        .Text = "CreationClassName"
        .col = 1
        .Text = Object.CreationClassName
        
        .Row = 8
        .col = 0
        .Text = "CSCreationClassName"
        .col = 1
        .Text = Object.CSCreationClassName
        
        .Row = 9
        .col = 0
        .Text = "CSDVersion"
        .col = 1
        .Text = Object.CSDVersion
        
        .Row = 10
        .col = 0
        .Text = "CSName"
        .col = 1
        .Text = Object.CSName
        
        .Row = 11
        .col = 0
        .Text = "CurrentTimeZone"
        .col = 1
        .Text = Object.CurrentTimeZone
        
        .Row = 12
        .col = 0
        .Text = "Debug"
        .col = 1
        .Text = Object.Debug
        
        .Row = 13
        .col = 0
        .Text = "Description"
        .col = 1
        .Text = Object.Description
        
        .Row = 14
        .col = 0
        .Text = "Distributed"
        .col = 1
        .Text = Object.Distributed
        
        .Row = 15
        .col = 0
        .Text = "ForegroundApplicationBoost"
        .col = 1
        .Text = Object.ForegroundApplicationBoost
        
        .Row = 16
        .col = 0
        .Text = "FreePhysicalMemory"
        .col = 1
        .Text = Object.FreePhysicalMemory
        
        .Row = 17
        .col = 0
        .Text = "FreeSpaceInPagingFiles"
        .col = 1
        .Text = Object.FreeSpaceInPagingFiles
        
        .Row = 18
        .col = 0
        .Text = "FreeVirtualMemory"
        .col = 1
        .Text = Object.FreeVirtualMemory
        
        .Row = 19
        .col = 0
        .Text = "InstallDate"
        .col = 1
        .Text = Object.InstallDate
        
        .Row = 20
        .col = 0
        .Text = "LastBootUpTime"
        .col = 1
        .Text = Object.LastBootUpTime
        
        .Row = 21
        .col = 0
        .Text = "LocalDateTime"
        .col = 1
        .Text = Object.LocalDateTime
        
        .Row = 22
        .col = 0
        .Text = "Locale"
        .col = 1
        .Text = Object.locale
        
        .Row = 23
        .col = 0
        .Text = "Manufacturer"
        .col = 1
        .Text = Object.Manufacturer
        
        .Row = 24
        .col = 0
        .Text = "MaxNumberOfProcesses"
        .col = 1
        .Text = Object.MaxNumberOfProcesses
        
        .Row = 25
        .col = 0
        .Text = "MaxProcessMemorySize"
        .col = 1
        .Text = Object.MaxProcessMemorySize
        
        .Row = 26
        .col = 0
        .Text = "Name"
        .col = 1
        .Text = Object.Name
        
        .Row = 27
        .col = 0
        .Text = "NumberOfLicensedUsers"
        .col = 1
        .Text = Object.NumberOfLicensedUsers
        
        .Row = 28
        .col = 0
        .Text = "NumberOfProcesses"
        .col = 1
        .Text = Object.NumberOfProcesses
        
        .Row = 29
        .col = 0
        .Text = "NumberOfUsers"
        .col = 1
        .Text = Object.NumberOfUsers
        
        .Row = 30
        .col = 0
        .Text = "Organization"
        .col = 1
        .Text = Object.Organization
        
        .Row = 31
        .col = 0
        .Text = "OSLanguage"
        .col = 1
        .Text = Object.OSLanguage
        
        .Row = 32
        .col = 0
        .Text = "OSProductSuite"
        .col = 1
        .Text = Object.OSProductSuite
        
        .Row = 33
        .col = 0
        .Text = "OSType"
        .col = 1
        .Text = Object.OSType
        
        .Row = 34
        .col = 0
        .Text = "OtherTypeDescription"
        .col = 1
        .Text = Object.OtherTypeDescription
        
        .Row = 35
        .col = 0
        .Text = "PlusProductID"
        .col = 1
        .Text = Object.PlusProductID
        
        .Row = 36
        .col = 0
        .Text = "PlusVersionNumber"
        .col = 1
        .Text = Object.PlusVersionNumber
        
        .Row = 37
        .col = 0
        .Text = "Primary"
        .col = 1
        .Text = Object.Primary
        
        .Row = 38
        .col = 0
        .Text = "QuantumLength"
        .col = 1
        .Text = Object.QuantumLength
        
        .Row = 39
        .col = 0
        .Text = "QuantumType"
        .col = 1
        .Text = Object.QuantumType
        
        .Row = 40
        .col = 0
        .Text = "RegisteredUser"
        .col = 1
        .Text = Object.RegisteredUser
        
        .Row = 41
        .col = 0
        .Text = "SerialNumber"
        .col = 1
        .Text = Object.SerialNumber
        
        .Row = 42
        .col = 0
        .Text = "ServicePackMajorVersion"
        .col = 1
        .Text = Object.ServicePackMajorVersion
        
        .Row = 43
        .col = 0
        .Text = "ServicePackMinorVersion"
        .col = 1
        .Text = Object.ServicePackMinorVersion
        
        .Row = 44
        .col = 0
        .Text = "SizeStoredInPagingFiles"
        .col = 1
        .Text = Object.SizeStoredInPagingFiles
        
        .Row = 45
        .col = 0
        .Text = "Status"
        .col = 1
        .Text = Object.Status
        
        .Row = 46
        .col = 0
        .Text = "SystemDevice"
        .col = 1
        .Text = Object.SystemDevice
        
        .Row = 47
        .col = 0
        .Text = "SystemDirectory"
        .col = 1
        .Text = Object.SystemDirectory
        
        .Row = 48
        .col = 0
        .Text = "TotalSwapSpaceSize"
        .col = 1
        .Text = Object.TotalSwapSpaceSize
        
        .Row = 49
        .col = 0
        .Text = "TotalVirtualMemorySize"
        .col = 1
        .Text = Object.TotalVirtualMemorySize
        
        .Row = 50
        .col = 0
        .Text = "TotalVisibleMemorySize"
        .col = 1
        .Text = Object.TotalVisibleMemorySize
        
        .Row = 51
        .col = 0
        .Text = "Version"
        .col = 1
        .Text = Object.Version
        
        .Row = 52
        .col = 0
        .Text = "WindowsDirectory"
        .col = 1
        .Text = Object.WindowsDirectory
      End With
    
    Exit For
    Next
    
    
    'MDIProcServ.WBEMServices.ExecNotificationQueryAsync MDIProcServ.eventSink, "Select * from __InstanceModificationEvent Within 2.0 Where TargetInstance Isa 'Win32_Service'"
    Set Enumerator = MDIProcServ.WBEMServices.ExecQuery("Select * From Win32_ComputerSystem")
    
    If Err.Number = 91 Then
        MsgBox "Please CONNECT to Local or Remote computer first", vbOKOnly, "Connect please!"
        Exit Sub
    End If
    
    For Each Object In Enumerator
      With msfgComputerSystem
      
        
        TTz4 = Trim(Object.Manufacturer + " " + Object.Model) + vbCrLf + cpudes + "  "
        Exit For
    
            .Row = 22
        .col = 0
        .Text = "Manufacturer"
        .col = 1
        .Text = Object.Manufacturer
        
        .Row = 23
        .col = 0
        .Text = "Model"
        .col = 1
        .Text = Object.Model
        
                
        
    End With
    Exit For
    Next
    
TTz5 = ""
    'Processor Information
Set obj = GetObject("winmgmts:").InstancesOf("Win32_Processor")
    For Each obj2 In obj
        'text3.Text = text3.Text & obj2.Caption & ENTER
        'text3.Text = text3.Text & "Speed: " & obj2.currentclockspeed & " Mhz" & ENTER
        TTz5 = Format(Val(obj2.CurrentClockSpeed) / 1000, "0.0") + "GHz "
    Next


'WRONG FOR AGES
'Set obj = GetObject("winmgmts:").InstancesOf("Win32_PhysicalMemory")
Set obj = GetObject("winmgmts:").InstancesOf("Win32_ComputerSystem")

Dim i As String
For Each obj2 In obj
        
        'WRONG FOR AGES
        'TTz5 = TTz5 + Trim(Trim(str(obj2.Capacity / 1024 ^ 3))) + " Gb Ram"
        TTz5 = TTz5 + Trim(Format(obj2.TotalPhysicalMemory / 1024 ^ 3, "0.00")) + "G Ram"
    
    Exit For
Next

        'msgbox TTz4 + " " + cpudes







End Sub



Sub DrivesGB()

Set FS = CreateObject("Scripting.FileSystemObject")

Dim Nuke4, Nuke5, RR, icount



Set FS = CreateObject("Scripting.FileSystemObject")
Dim d
On Local Error Resume Next
Set dc = FS.Drives
For Each d In dc
    Err.Clear
    DR = d.DriveLetter
    If d.Driveready = True Then
    DRX = d.DriveType

    



    
    IDD = 1
    
    'If DR = "Z" Then IDD = 1
    'If DRX = 1 Or DRX = 4 Or DRX = 3 Then IDD = 1
    
    If DRX = 2 Then IDD = 0
    
    If IDD = 0 Then n = d.VolumeName


    'Select Case d.DriveType
    'Case 0: t = "Unknown"
    'Case 1: t = "Removable"
    'Case 2: t = "Fixed"
    'Case 3: t = "Network"
    'Case 4: t = "CD-ROM"
    'Case 5: t = "RAM Disk"


    'IDD = 0
    If InStr(n, "MatthewLan") > 0 Then IDD = 1
    If InStr(n, "-TRU") > 0 Then IDD = 1
    
    
    If IDD = 0 Then
    
        If InStr(n, "BAK") > 0 And Err.Number = 0 Then
            bakspace = bakspace + d.TotalSize / 1024 ^ 3
        End If



        icount = icount + 1
        RR = RR + d.VolumeName + "-"
        Nuke4 = Nuke4 + d.TotalSize / 1024 ^ 3
        Nuke5 = Nuke5 + d.FreeSpace / 1024 ^ 3
    End If
    End If
Next
If Nuke5 < TotHd Then E$ = "ä"
If Nuke5 > TotHd And TotHd > 0 Then E$ = "ã"


TotHd = Nuke5

TTz6 = Trim(str(icount)) + " Drive's " + Format(Nuke4 / 1024, "0.00") + " TByte. Minus " + Format(bakspace / 1024, "0.00") + " TB Backup " + Format(Nuke5, "0") + "Gb Free"


End Sub






' Return an array of Long holding the handles of all the child windows
' of a given window. If hWnd = 0 it returns all the top-level windows.

Function ChildWindows(ByVal hWnd As Long) As Long()
    WindowsCount = 0
    If hWnd Then
        EnumChildWindows hWnd, AddressOf EnumWindows_CBK, 1
    Else
        EnumWindows AddressOf EnumWindows_CBK, 1
    End If
    ' Trim uninitialized elements and return to caller.
    ReDim Preserve Windows(WindowsCount) As Long
    ChildWindows = Windows()
End Function

' The callback routine, common to both EnumWindows and EnumChildWindows.

Private Function EnumWindows_CBK(ByVal hWnd As Long, ByVal lParam As Long) As _
    Long
    ' Make room in the array, if necessary.
    If WindowsCount = 0 Then
        ReDim Windows(100) As Long
    ElseIf WindowsCount >= UBound(Windows) Then
        ReDim Preserve Windows(WindowsCount + 100) As Long
    End If
    ' Store the new item.
    WindowsCount = WindowsCount + 1
    Windows(WindowsCount) = hWnd
    ' Return 1 to continue enumeration.
    EnumWindows_CBK = 1
End Function



