Attribute VB_Name = "A1_VarsMod"
Option Explicit

Public API_LOAD


Public TIMER_MISSING_PULSE_CLIPBOARD_01
Public TIMER_MISSING_PULSE_CLIPBOARD_02


Public LAST_CLIPBOARD_FILE
Public PATH_TO_CLIPBOARD_IMAGE_LOGGER_HOT_KEY As String
Public PATH_TO_CLIPBOARD_IMAGE_LOGGER_AUTO_SCR__ As String
Public PATH_TO_CLIPBOARD_IMAGE_LOGGER_AUTO_FORM_ As String
Public PATH_TO_CLIPBOARD_IMAGE_INFO_IN_APP As String
Public PATH_TO_CLIPBOARD_TEXT_INFO_APP_DAY_DATA As String
Public PATH_TO_CLIPBOARD_APP_DATA_1 As String
Public PATH_TO_CLIPBOARD_APP_DATA_2 As String
Public PATH_TO_CLIPBOARD_APP_DATA_3 As String
Public PATH_TO_CLIPBOARD_APP_DATA_4 As String
Public PATH_TO_CLIPBOARD_APP_DATA_5 As String
Public FILE_NAME_USER_HOTKEY_SHOT_PIC_DATA_NETWORK_COMMON As String
Public FILE_NAME_USER_HOTKEY_SHOT_PIC_DATA_LOCATION_CURRENT As String
Public FILE_NAME_USER_HOTKEY_SHOT_PIC_DATA_LOCATION As String



Public Declare Function WindowFromPoint _
                 Lib "user32" (ByVal xPoint As Long, _
                               ByVal yPoint As Long) As Long

Public O_VAL_MINUTE_API

Public SORTI
Public ERROR_LOGG_UPDATE_TIME, ERROR_LOGG_READ_STATUS

Public RETRY2
Public RETRY3
Public EXIT_TRUE


Public CLIPBOARD_ACTIVITY_MOMENT

Public FS
Public FSO

Public CPC_404_CPC

Public TestLowTimer

Public KEYBOARD_OR_MOUSE_ACTIVE_3_MIN
Public KEYBOARD_OR_MOUSE_ACTIVE_10_SEC
Public FILENAME_IN_USE_CHECK As String

Public TIME_LAST_CLIPBOARD_ERROR_MSG
Public TIME_LAST_CLIPBOARD
Public TIME_LAST_CLIPBOARD_Timer1

Public ClipFormatDesc2

Public Timer_API_Test_Logick_Var1 As Date
Public Timer_API_Test_Logick_Var2 As Date

Public HOOK_CLIPBOARD_API_lOADED

Public URL, WinTitle As String, HHH, EXECUTE_TIMER_ENABLED


Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long



'In a module
Public Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
  (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function joyGetPos Lib "Winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long

Private Const EM_GETTHUMB = &HBE
Private Const SB_THUMBPOSITION = &H4
Private Const WM_VSCROLL = &H115
Public Type JOYINFO
   x As Long
   y As Long
   Z As Long
   Buttons As Long
End Type

Public Type JOYCAPS
   wMid As Integer
   wPid As Integer
   szPname As String * 32
   wXmin As Long
   wXmax As Long
   wYmin As Long
   wYmax As Long
   wZmin As Long
   wZmax As Long
   wNumButtons As Long
   wPeriodMin As Long
   wPeriodMax As Long
 End Type

Public Declare Function joyGetDevCapsA Lib "Winmm.dll" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long
Public Declare Function joyGetDevCapsW Lib "Winmm.dll" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type B
B(10) As Boolean
End Type

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Const hWnd_TOPMOST = -1
Const hWnd_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Function GET_DRIVES_FIND_FOLDER(FOLDER As String)

GET_DRIVES_FIND_FOLDER = ""

Dim DC, d, S
Set DC = FSO.Drives
For Each d In DC
  S = S & d.DriveLetter
  If d.DriveType <> 2 Then 'FIXED DRIVE
    'Stop
    'n = d.ShareName
  ElseIf d.IsReady Then
    'N = N + d.DriveLetter 'd.VolumeName
    If FSO.FolderExists(d.DriveLetter + ":" + FOLDER) = True Then
        GET_DRIVES_FIND_FOLDER = d.DriveLetter + ":" + FOLDER
        Exit Function
    End If
    
    
  End If
  's = s & n & "<BR>"
Next

'GET_DRIVES = N
'GET_DRIVES_RESULT = N

End Function

Private Sub SPARE_BIT_OF_KITT()

Dim FOLDER_SENDTO_FAT32_STORE

'----------------
'FIND OUR FAT32 DRIVE BY SEARCH FOLDER SENDTO FOLDER
'----------------
'DEFAULT XP
Dim FOLDER_FIND As String
FOLDER_FIND = "\01 SendTo\"
'----------------
'SIX = WIN 7
'----------------
If GETWinNT_Ver = "WIN XP" Then FOLDER_FIND = "\01 SendTo 01- WIN XP"
If GETWinNT_Ver = "WIN 7" Then FOLDER_FIND = "\01 SendTo 02- WIN 7"
If GETWinNT_Ver = "WIN 7 - 64" Then FOLDER_FIND = "\01 SendTo 02- WIN 7"
If GETWinNT_Ver = "WIN 10" Then FOLDER_FIND = "\01 SendTo 03- WIN 10"
'----------------

'HEAVY TASK - SWITCH TO PROBLEM CAN HAPPEN
'If GETWinNT_Ver = "WIN 7" Then
'    If GetOsBitness = 64 Then FOLDER_FIND = "\01 SendTo 02- WIN 7 - 64"
'End If
''GetCpuBitness

'BETTER - BUT STILL USE FEW CPU CYCLES
If GETWinNT_Ver = "WIN 7" Then
    If GetOsArchitecture = 64 Then FOLDER_FIND = "\01 SendTo 02- WIN 7 - 64"
End If
'----------------
FOLDER_SENDTO_FAT32_STORE = GET_DRIVES_FIND_FOLDER(FOLDER_FIND)
'MNU_SEND_TO_FAT32_FOLDER.Caption = "EXPLORER -- " + FOLDER_SENDTO_FAT32_STORE

End Sub





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
    
    Form1.MNU_FILE_STUCK_IN_USE.Visible = False
    DD_DIFF = 0
    Do
        DoEvents
        ii = IsFileAlreadyOpen(FILENAME_IN_USE_CHECK_2)
        If ii = True Then
            Sleep (200)
            DD_DIFF = DateDiff("S", Now, QQ)
            Form1.MNU_FILE_STUCK_IN_USE.Visible = True
            If DD_DIFF <> O_DD_DIFF Then
                Form1.MNU_FILE_STUCK_IN_USE.Caption = "IsFileOpenDelay Busy --" + Str(DD_DIFF) + " Secs - File Stuck In Use Of Limit -" + Str(QQ_LIMIT) + " -- " + FILENAME_IN_USE_CHECK_2
            End If
            O_DD_DIFF = DD_DIFF
        End If
    Loop Until ii = False Or QQ < Now
    
    If ii = False Then
        Form1.MNU_FILE_STUCK_IN_USE.Caption = "Last IsFileOpenDelay --" + Str(QQ_LIMIT) + " Second"
    End If
    
    If ii = True Then
        Form1.MNU_FILE_STUCK_IN_USE.Caption = "Error IsFileOpenDelay --" + Str(DD_DIFF) + " Second File Stuck In Use -- Of Limit --" + Str(QQ_LIMIT) + " -- " + FILENAME_IN_USE_CHECK_2
        MSGBOX2 = "ERROR IsFileOpenDelay " + vbCrLf + FILENAME_IN_USE_CHECK_2 + vbCrLf + Str(DD_DIFF) + " Second of --" + Str(QQ_LIMIT) + " Limit"
        frm_MSGBOX.Timer1.Enabled = True
    End If
End Function


Private Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Private Function NotAlwaysOnTop(ByVal hWnd As Long)
    Dim flags
    SetWindowPos hWnd, hWnd_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function



'Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

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





Public Function API_CLIPBOARD_HOOK() As Integer

'API_CLIPBOARD_HOOK = HOOK_CLIPBOARD_API_lOADED

End Function




Public Function GetVerticalScrollPos(rtb As RichTextBox) As Long
  GetVerticalScrollPos = SendMessage(rtb.hWnd, EM_GETTHUMB, 0&, 0&)
End Function

Public Sub SetVerticalScrollPos(rtb As RichTextBox, Position As Long)
    On Error Resume Next
    'Position = 100
    SendMessage rtb.hWnd, WM_VSCROLL, SB_THUMBPOSITION + &H10000 * Position, Nothing
End Sub














Public Function GetBPress(num As Long, Buttons As B)
Dim S As Long, N, k, P
For N = 1 To 10
Buttons.B(N) = False
Next
If num = 0 Then Exit Function
If num = 1 Then
Buttons.B(1) = True
Exit Function
End If
If num = 2 Then
Buttons.B(2) = True
Exit Function
End If

Do
    N = 1
    S = num \ 2
    Do While S >= 2
    N = N + 1
    S = S \ 2
    Loop

    k = 2 ^ N
    Select Case k
    Case 1
    Buttons.B(1) = True
    Case 2
    Buttons.B(2) = True
    Case 4
    Buttons.B(3) = True
    Case 8
    Buttons.B(4) = True
    Case 16
    Buttons.B(5) = True
    Case 32
    Buttons.B(6) = True
    Case 64
    Buttons.B(7) = True
    Case 128
    Buttons.B(8) = True
    Case 256
    Buttons.B(9) = True
    Case 512
    Buttons.B(10) = True
    End Select

    P = num - k
    num = P
    
If P = 1 Then
    Buttons.B(1) = True
Exit Do
ElseIf P = 0 Then
Exit Do
End If
Loop
End Function

Public Function IsWinNT() As Boolean
    Dim OSVer As OSVERSIONINFO
    'Dim WinVer
    
    OSVer.dwOSVersionInfoSize = Len(OSVer)
    GetVersionEx OSVer
    IsWinNT = (OSVer.dwPlatformId = 2)
    'WIN VER AS REPORT BY CMD SHELL
    'WinVer = (OSVer.dwMajorVersion) + "." + (OSVer.dwMinorVersion) + "." + (OSVer.dwBuildNumber)
    
    'TAKE THE FIRST DIGIT ONLY WIN 7 IS MAJOR VER SIX
    'WinVer = (OSVer.dwMajorVersion)
    
    'SERVICE PACK NUMBER
    'WinVer = (OSVer.szCSDVersion)
    
End Function

Public Function GETWinNT_Ver() As String
    Dim OSVer As OSVERSIONINFO
    Dim WinVer
    Dim IsWinNT
    
    OSVer.dwOSVersionInfoSize = Len(OSVer)
    GetVersionEx OSVer
    IsWinNT = (OSVer.dwPlatformId = 2)
    
    GETWinNT_Ver = "WIN XP"
    
    If IsWinNT = True Then
        'WIN VER AS REPORT BY CMD SHELL
        'WinVer = Trim(Str(OSVer.dwMajorVersion)) + "." + Trim(Str(OSVer.dwMinorVersion)) + "." + Trim(Str(OSVer.dwBuildNumber))
        
        'TAKE THE FIRST DIGIT ONLY WIN 7 IS MAJOR VER SIX
        'WinVer = (OSVer.dwMajorVersion)
        'GETWinNT_Ver = (OSVer.dwMajorVersion)
        
        If OSVer.dwMajorVersion = 6 Then
            GETWinNT_Ver = "WIN 7"
        End If
        
        'NOT NUMERIC VALUE
        'SERVICE PACK NUMBER
        'WinVer = (OSVer.szCSDVersion)
    End If
    
    
    'WIN 10 IS ALSO VER --SIX-- AS IS WIN 7
    'WIN 10
    'CMD VER WIN 10 Microsoft Windows [Version 10.0.10586]
    'OSVer.dwMajorVersion -- SIX
    'OSVer.dwMinorVersion -- 2
    'OSVer.dwBuildNumber -- 9200
    If IsWinNT = True Then
        If OSVer.dwMajorVersion = 6 Then
        If OSVer.dwMinorVersion = 2 Then
        If OSVer.dwBuildNumber = 9200 Then
            GETWinNT_Ver = "WIN 10"
        End If
        End If
        End If
    End If
    
End Function


'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
  ' Test = False
End Function
'***********************************************


