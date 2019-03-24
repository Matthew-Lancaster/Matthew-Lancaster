VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "KAT MP3 RELOAD"
   ClientHeight    =   5904
   ClientLeft      =   588
   ClientTop       =   5448
   ClientWidth     =   12864
   Icon            =   "KAT MP3 DEL NAUGHT SOUND.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5904
   ScaleWidth      =   12864
   Begin VB.Timer Timer_1_MINUTE 
      Interval        =   1000
      Left            =   12432
      Top             =   780
   End
   Begin VB.DriveListBox Drive1 
      Height          =   288
      Left            =   6084
      TabIndex        =   9
      Top             =   1284
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Timer TIMER_SET_CAMERA_CHECK 
      Interval        =   1000
      Left            =   8604
      Top             =   588
   End
   Begin VB.Timer TIMER_SET_CAMERA_WITH_SECURITY_INFO 
      Interval        =   10
      Left            =   7944
      Top             =   612
   End
   Begin VB.Timer Timer_VICE_VERSA_TIMER_ON 
      Interval        =   59000
      Left            =   11820
      Top             =   936
   End
   Begin VB.Timer Timer_VB_MAXIMIZE 
      Interval        =   1
      Left            =   11436
      Top             =   1080
   End
   Begin VB.Timer Timer_COMMAND_RESTRICTOR_NOTEPAD_PICK_UP 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5952
      Top             =   2856
   End
   Begin VB.Timer TIMER_KILL_CMD_PREVIOUS_INSTANCE_SOFTLY_WHEN_TOO_MANY 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5424
      Top             =   2868
   End
   Begin VB.Timer Timer_EXPLORER_GONE 
      Interval        =   1000
      Left            =   9960
      Top             =   1032
   End
   Begin VB.Timer TIMER_MNU_CID_RELOADER_DISAPPEAR 
      Interval        =   5000
      Left            =   9528
      Top             =   1008
   End
   Begin VB.Timer TIMER_EXPLORER_IS_PRESENT_OPEN_AT_FOLDER_TRUE 
      Interval        =   1000
      Left            =   8715
      Top             =   1035
   End
   Begin VB.Timer Timer_Remove_Explorer_Libraries 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8265
      Top             =   1035
   End
   Begin VB.Timer TIMER_MouseDetectMove 
      Interval        =   10
      Left            =   7815
      Top             =   1035
   End
   Begin VB.Timer Timer_EXPLORER_RESET 
      Interval        =   1000
      Left            =   7350
      Top             =   1050
   End
   Begin VB.Timer Timer_Tidal_Information_Microsoft_Visual_Basic_break_SubRoutine 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6885
      Top             =   1050
   End
   Begin VB.Timer Timer_01_VB_HELP_BOX_02_MSDN 
      Interval        =   100
      Left            =   11004
      Top             =   1104
   End
   Begin VB.FileListBox File5_ARCHIVE_KAT_MP3 
      Height          =   264
      Left            =   8985
      TabIndex        =   8
      Top             =   2505
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5475
      Top             =   2160
   End
   Begin VB.Timer Timer_VICE_VERSA_ME 
      Interval        =   55000
      Left            =   5565
      Top             =   1545
   End
   Begin VB.FileListBox File4 
      Height          =   264
      Left            =   10212
      TabIndex        =   7
      Top             =   2052
      Visible         =   0   'False
      Width           =   1056
   End
   Begin VB.Timer Timer_TIME_HAPPEN 
      Interval        =   1
      Left            =   3840
      Top             =   1776
   End
   Begin VB.Timer Timer_Find_Window_SPEED 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4524
      Top             =   1776
   End
   Begin VB.Timer Timer_KAT_MOUSE_FOR_VB 
      Interval        =   100
      Left            =   5580
      Top             =   1056
   End
   Begin VB.FileListBox File3 
      Height          =   264
      Left            =   8844
      TabIndex        =   6
      Top             =   2052
      Visible         =   0   'False
      Width           =   1308
   End
   Begin VB.Timer TIMER_RETRY_WRITE_INFO_UNTIL_DONE2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4884
      Top             =   1428
   End
   Begin VB.Timer TIMER_RETRY_WRITE_INFO_UNTIL_DONE1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5220
      Top             =   1056
   End
   Begin VB.Timer WebCam_Matt_Timer 
      Interval        =   1000
      Left            =   4884
      Top             =   2112
   End
   Begin VB.Timer TIMER_DEL_THUMB_CAMAREA_MP4 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   4884
      Top             =   1776
   End
   Begin VB.Timer Timer_ME_CAMERA 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7392
      Top             =   612
   End
   Begin VB.FileListBox File2 
      Height          =   264
      Left            =   8856
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Timer zzCheckTimer 
      Enabled         =   0   'False
      Interval        =   59000
      Left            =   4536
      Top             =   2112
   End
   Begin VB.Timer Timer_Find_Window 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4536
      Top             =   1428
   End
   Begin VB.Timer Timer_DESKTOP_INI_DEL 
      Interval        =   1000
      Left            =   4536
      Top             =   1080
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1248
      Left            =   0
      TabIndex        =   3
      Top             =   3144
      Width           =   13635
   End
   Begin VB.Timer TIMER_SET_CAMERA_UP_TO_MARK_FOLDER 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4188
      Top             =   2112
   End
   Begin VB.DirListBox Dir2 
      Height          =   720
      Left            =   7752
      TabIndex        =   2
      Top             =   2052
      Visible         =   0   'False
      Width           =   1032
   End
   Begin VB.Timer Timer_W32TM 
      Interval        =   1000
      Left            =   3840
      Top             =   2112
   End
   Begin VB.Timer TIMER_DEL_CAMERA_EMPTY_MP4_MOV 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6492
      Top             =   624
   End
   Begin VB.DirListBox Dir1 
      Height          =   720
      Left            =   6588
      TabIndex        =   1
      Top             =   2052
      Visible         =   0   'False
      Width           =   1116
   End
   Begin VB.Timer Timer_Drives 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4200
      Top             =   1428
   End
   Begin VB.Timer Timer_DEL_WAVS 
      Interval        =   1000
      Left            =   3840
      Top             =   1428
   End
   Begin VB.Timer Timer_WINMERGE_KILL 
      Interval        =   1000
      Left            =   4188
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3852
      Top             =   1056
   End
   Begin VB.FileListBox FILE1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2472
      Left            =   0
      TabIndex        =   0
      Top             =   408
      Width           =   13635
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   13620
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "| EXIT |"
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "| VB ME |"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "| VB FOLDER |"
   End
   Begin VB.Menu MNU_ADMINISTRATOR 
      Caption         =   "ADMINISTRATOR MODE ON"
   End
   Begin VB.Menu MNU_FILE_DOWN 
      Caption         =   "DOWN"
   End
   Begin VB.Menu MNU_FILE_UP 
      Caption         =   "UP"
   End
   Begin VB.Menu Mnu_Mixer_Deck_Tunes 
      Caption         =   "Mixer Deck Tunes"
      Begin VB.Menu Mnu_Mixer_Deck_Tunes_K_Lite_Player_Classic 
         Caption         =   "Play Mixer Tunes -- hifi www.stream1.nubreaks.com.m3u -- With -- \Program Files (x86)\K-Lite Codec Pack\MPC-HC64\mpc-hc64.exe"
      End
   End
   Begin VB.Menu MNU_COMMAND_RESTRICTOR 
      Caption         =   "1 OF 2 _ EDIT_THE_COMMAND_RESTRICTOR"
   End
   Begin VB.Menu MNU_COMMAND_RESTRICTOR_ENABLED 
      Caption         =   "2 OF 2 _ COMMAND_RESTRICTOR_ENABLED_=_NOT"
   End
   Begin VB.Menu MNU_PROGRAM 
      Caption         =   "PROGRAM EXE MENU FILE_LOCATOR && CID_RUN_ME"
      Begin VB.Menu MNU_FILE_LOCATOR 
         Caption         =   "FILE LOCATOR"
      End
      Begin VB.Menu MNU_CID_RERUN_ME_RELOADER 
         Caption         =   "CID RUN ME __ ACTIVE RELOADER __ CHECK BOX SELECT __ HITT CHECK HERE"
      End
   End
   Begin VB.Menu MNU_CID_RELOADER 
      Caption         =   "CIR RUN_ME ACTIVE_RELOADER_WON'T_WHEN_FORM_NOT_MINIMIZED"
   End
   Begin VB.Menu MNU_ENABLE_FINDWINOW_TIMER 
      Caption         =   "| FIND WINDOW AND FOLDER EXPLORER MENU |"
      Begin VB.Menu MNU_ENABLE_FINDWINOW_TIMER_VAR 
         Caption         =   "FIND WINDOW MAIN ENGINE CORE - TIMER ENABLED"
         Checked         =   -1  'True
      End
      Begin VB.Menu MNU_OFFICE_TOOLBAR_USE 
         Caption         =   "OFFICE TOOLBAR RELOAD"
         Checked         =   -1  'True
      End
      Begin VB.Menu MNU_KAT_MP3_FOLDER 
         Caption         =   "KAT MP3 FOLDER EXPLORER"
      End
      Begin VB.Menu MNU_KAT_MP3_FOLDER_FILE_LOCATOR_PRO 
         Caption         =   "KAY MP3 FOLDER -- FILE LOCATOR PRO"
      End
      Begin VB.Menu MNU_KAT_MP3_RUN 
         Caption         =   "KAT MP3 RECORDER - RUN"
      End
      Begin VB.Menu MNU_SEND_TO_FAT32_FOLDER 
         Caption         =   "SEND TO - Fat32 Folder"
      End
      Begin VB.Menu MNU_SEND_TO_SYSTEM_FOLDER 
         Caption         =   "SEND TO - SYSTEM"
      End
   End
   Begin VB.Menu MNU_W_DRIVE 
      Caption         =   "W DRIVE WORKING"
   End
   Begin VB.Menu MNU_W_DRIVE_OFF 
      Caption         =   "W DRIVE OFF && GO"
   End
   Begin VB.Menu MNU_SECURITY_TO_CAMERA 
      Caption         =   "SECURITY TO CAMERA FILE"
   End
   Begin VB.Menu MNU_PASTE_VOLUME_LABEL_CAMERA 
      Caption         =   "PASTE VOLUME LABLE CAMERA ""MC - HX60V"""
   End
   Begin VB.Menu MNU_VICE_VERSA_TIMER_ON 
      Caption         =   "VICE-VERSA-TIMER-ON"
   End
   Begin VB.Menu MNU_RUN 
      Caption         =   "| RUN LOAD VICE VERSA |"
      Begin VB.Menu MNU_VICEVERSA_EXPLORER 
         Caption         =   "VICE VERSA Explorer -- FOLDER W DRIVE"
      End
      Begin VB.Menu MNU_VICEVERSA_W_SUBST 
         Caption         =   "VICE VERSA CAMERA - SETUP W DRIVE - SUBST"
      End
      Begin VB.Menu MNU__1 
         Caption         =   "--------"
      End
      Begin VB.Menu MNU_EXPLORER_SECURITY_INFO 
         Caption         =   "CAMERA SECURITY INFO COPY FROM HERE - DO EXPLORER -  App.Path + ""\CAMERA DATA"""
      End
      Begin VB.Menu MNU_EXPLAIN 
         Caption         =   "-------- TRYING TO FIND .volumename = ""##"" Or .volumename = ""MC - HX60V"" -- PASTE IT ON HARDCORE"
      End
      Begin VB.Menu MNU_VICEVERSA_REMOTE 
         Caption         =   "VICE VERSA SCRIPT - CAMERA REMOTE"
      End
      Begin VB.Menu MNU_VICEVERSA_LOCAL 
         Caption         =   "VICE VERSA SCRIPT - CAMERA LOCAL"
      End
      Begin VB.Menu MNU_VICEVERSA_ME 
         Caption         =   "VICE VERSA SCRIPT - CAMERA ME"
      End
   End
   Begin VB.Menu MNU_VICE_VERSA_ME 
      Caption         =   "** VICE VERSA ME **"
   End
   Begin VB.Menu MNU_VICE_VERSA_LOCAL 
      Caption         =   "** VICE VERSA LOCAL **"
   End
   Begin VB.Menu MNU_VICE_VERSA_NETWORK_SYNCRO 
      Caption         =   "** VICE VERSA NETWORK SYNCRO **"
   End
   Begin VB.Menu MNU_TIMER_SINCE_ACTIVITY 
      Caption         =   "TIMER SINCE ACTIVITY"
   End
   Begin VB.Menu MNU_WAV_ACTIVITY 
      Caption         =   "MNU WAV ACTIVITY - NOT YET"
   End
   Begin VB.Menu MNU_WAV_ACTIVITY_ARCHIVE 
      Caption         =   "MNU WAV ACTIVITY_ARCHIVE - NOT YET"
   End
   Begin VB.Menu MNU_AUTOHOTKEY_SAVE_AS 
      Caption         =   "\Autokey -- Save As Key Enter.ahk"
   End
   Begin VB.Menu MNU_CLIPBOARD_CONVERT_VODAFONE 
      Caption         =   "CLIPBOARD CONVERT HEX(20) TO HEX(32) TO A WHITE SPACE  FOR VODAFONE MMS MESSAGE TO EMAIL LINE BREAKER"
   End
   Begin VB.Menu MNU_GET_EXPLORER_CURRENT_SESSION_CLIPBOADIN 
      Caption         =   "GET_EXPLORER_CURRENT_SESSION_CLIPBOADIN"
   End
   Begin VB.Menu MNU_NOTEPAD__BOOT_KILLER 
      Caption         =   "NOTEPAD__BOOT_KILLER"
   End
   Begin VB.Menu MNU_BOOT_KILLER 
      Caption         =   "RUN BOOT_KILLER"
   End
   Begin VB.Menu MNU_TASK_KILLER 
      Caption         =   "TASK KILLER"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WSHShell

Dim SCRIPT_TO_RUN
Dim MSgBox_NOT_EXIST_SCRIPT_TO_RUN
Dim XVB_DATE

Dim KAT_MP3_IF_NOT_EXIST_AND_THEN_SHELL

'----------------------------------------------------------------------------
'PROBLEM HERE __ WHEN PATH IS NOT D DRIVE FOR CODE EXE OR NEW LOCATION TESTER
'VB_PROJECT_CHECKDATE
'----------------------------------------------------------------------------
'-------------------------------------------------
'PROBLEM HERE __ DO KILL BEORE LOAD ON XP OR OTHER
'SOME_THING_WRONG_DURING_MERGE_CONFLICT()
'-------------------------------------------------

'KAT MOUSE PROJECT
'---------------------------------------------
'SEEM THE KATMOUSE UNLOAD CAN ALSO STEAL FOCUS
'THAT NEED PUT BACK
'
'SOME PULL DOWN MENU REQUIRE KATMOUSE TO STAY
'OR LOAD TIME INVOLED IN RETURN TO FORM
'---------------------------------------------
'mweather.cfg
'------------
'KAT MP3 CODE
'------------

'-----------------------------------
'GETTING STARTED TIP NOTES FOR WIN 7
'START WIN 7 EXPLORER IN MY COMPUTER

'-----------------------------------
'01 - EXPLORER IN MY COMPUTER START
'-----------------------------------
'%windir%\explorer.exe /E,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}
'-----------------------------------
'02 - TURN OFF INDEXER - WIN 7 - UP
'-----------------------------------
'http://superuser.com/questions/73204/correct-way-to-disable-indexing-in-windows-7
'-----------------------------------
'CONTROL PANEL - SERVICES - WINDOWS SEARCH - DISABLE
'-----------------------------------

'JOB
'INFO RAPID ON SEND TO
'NETWORK AND LOCAL SEARCH ALL FILES
'FILE COMPARE OPTION CHOICE GOODSYNC
'SEND TO FILENAME TO CLIPBOARD NOT MSGBOX BIT LITTLE FORM
'MORE JOBS IN GOODSYNC
'-- ARCHIVE HISTORY VB CODE WORK IT OUT
'ASK RAM FOR ASUS
'voice mail message direct record

'CRC32 CHECK WANT ADD HEADER BLOCK WHEN FINISH
'AND MAKE A COPY TO HAVE AT EVENT OF CRASH MAYBE

'TIDAL IS NOT TO REPORT DARNESS % IN DAYLIGHT
'DONE BEFORE CONFLICT OF NOT SYNC IN TIME

'SYNC NEED A GOOD ARCHIVE METHOD ALL CHANGE IN CODE
'CHECK IT OUT IF ANOTHER JOB TAKE PRIRITY OF SAVED AND HISTORY FOLDER

Dim OHWnd_VB_CLIPPER_ERROR
Dim OHWnd_VB_LOADER


Dim O_DRIVE1_ListCount
Dim VAR_MC__HX60V_SET

Dim LOAD_FIRST_RUN_W_DRIVE_OFF

Dim O_mWnd_VB_VbaWindow_MAXIMIZE
Dim OHWnd_FINDER_1


Dim LOAD_TIME1
Dim LOAD_TIME2
Dim LOAD_TIME3


Dim OHWnd_TEAM_VIEWER

Dim FILENAME_PATH_COMMAND_RESTRICTOR
Dim COMMAND_RESTRICTOR_VALUE_______
Dim COMMAND_RESTRICTOR_CLASS_STRING

Dim COMMAND_RESTRICTOR_NOTEPAD_PICK_UP_DATE As Date


Dim ExplorerGone
Dim FORM_STAY_UP_WITH_MSGBOX

Dim TIME_CHECK_IF_NOT_BEEN_SET_BACKER As Date

Dim CID_RUN_ME_SHELL_LOADED

Dim I_Memmer




Dim EXPLORER_DETECT_IS_PRESENT

Dim OcWnd_ICACLS_SETTER_PERMISSION_HWND As Long

Dim OcWnd_GOOD_SYNC As Long
Dim OcWnd_GOOD_SYNC_CRASH As Long




Dim MouseDetectMove_FOR_EXPLORER_RESET
Dim ONx, ONy
Dim EXPLORER_RESET_TIMER

Dim Tidal_Information_Microsoft_Visual_Basic_break_
Dim Tidal_Information_Microsoft_Visual_Basic_break_Counter
Dim XTELL_ONCE

Dim W32X2, W32_COUNT

Dim About_Kat_MP3_Recorder
'-------
'MEMORY
'HISTORY
'-------
'PAST
'PRESENT
'FURTURE
'-------

Dim OcWnd


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

Const BM_CLICK = &HF5&

Private Const WM_SETTEXT            As Long = &HC
Private Const WM_GETTEXT As Long = &HD
Private Const WM_GETTEXTLENGTH As Long = &HE

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWndParent As Long, _
     ByVal hWndChildAfter As Long, _
     ByVal lpszClass As String, _
     ByVal lpszTitle As String) _
    As Long

'Private Declare Function GetForegroundWindow Lib "user32" () As Long

'Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long  'MODULE 1141

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByRef lParam As Any) _
    As Long
    
'Private Const WM_SETTEXT            As Long = &HC
'Private Const WM_GETTEXT = &HD

'Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142


Dim OTargetHwnd_AUDIT

Dim PHOTO_DATE_VAR

Dim O_WindowState

Dim CAMERA_EXPLORER_PATH

Dim VAR_BOTTOM_OBJECT

Dim KAT_MP3_PATH_TMP As String
Dim KAT_MP3_PATH_UNTOLD
Dim KAT_MP3_PATH As String, KAT_MP3_PATH1, KAT_MP3_PAT2
Dim SET_INTERVAL_DEL_WAVS

'Dim VVCOMLINE

Dim EXPLORER_LAUNCH
Dim FOLDER_FIND As String
Dim FOLDER_SENDTO_SYSTEM
Dim FOLDER_SENDTO_FAT32_STORE

Dim OXK3_KATMOUSE_PRESENCE_VISUAL_TIMER

Dim MOUSE_LEFT_BUTTON_DELAY
Dim OX33
Dim O_HAS_VB_BEEN_VISIBLE_TOGGLE
Dim OXK1
Dim RESULT_KATMOUSE_NOT_BEEN_LOADED_BY_US_TIME
Dim OXK1_KATMOUSE_PRESENCE_TIMER

Dim PATH_FILE_NAME1
Dim RESULT_KATMOUSE_BEEN_LOADED_TIME
Dim RESULT_KATMOUSE_BEEN_LOADED_BY_US

'---------
'MORE WORK
'---------
'HERE -- Call VB_PROJECT_CHECKDATE
'---------------------------------
'GOOD ONE --
'-----------
'Public Function POSTMESSAGE_CLOSE_ALL_PARENT(ByVal TargetHwnd As Long) As String
'---------------------------------------------------------------------------------
'TO DO FINISH OFF --
'COMPLETE -- FINDWIN PART VICE VERSA WITH HWND NUMBER SCRIPT DETECT PREVIOUS MATCH
'-------------------
'i = FindWinPart_VISIBLE("ViceVersa Pro - [", "", HWND_RETURN, VAR1, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)
'---------------------------------------------------------------------------------

Dim HWND_RETURN As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uaction&, ByVal uparam&, ByRef lpvParam As Any, ByVal fuwinini&) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long   'Parameter: form.(property) = SetCursorPos(x,y) in Pixels.  'MODULE 1121

Private Declare Function FindWindowDLL _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Dim FIRST_RUN_FIND_PID_KAT_MP3
Dim PID As Long

'USE
'MINICOMPARE - VAR
Dim TargetHwnd As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long  'MODULE 1135
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2   ' in another
Private Const WM_QUIT As Long = &H12
Private Const WM_CLOSE = &H10
Private Const WM_SYSCOMMAND As Long = &H112
Private Const SC_CLOSE As Long = &HF060&

Dim WAIT_FOR_OFFICE_TOOLBAR_TO_LOAD

Dim KAT_XY As RECT
Dim MSDN_LIB_RECT As RECT

Dim INTERVAL_Timer_DESKTOP_INI_DEL
Dim INTERVAL_TIMER_DEL_WAVS


Dim COUNT_SEC_WEBCAM_MATT_FLAG
Dim COUNT_SEC_WEBCAM_MATT
Dim FIRST_LIMIT_RESET_DONE

Dim OX2
Dim TRIG1_GO3, TRIG2_GO3
Dim W32TM_RUN_ONCE

Dim COUNT_UP_WAIT
Dim VAR_SHELL_SUBST_BEEN_SET
Dim VAR_SHELL_SUBST_BEEN_SET_FLAG2
'Dim XVB_DATE As Date

Dim MINICOMPARE, OTTX, OOTTX, OOTTX2
Dim ADATE As Date
Dim VARCENTER

Dim XX
Dim OX
'Dim VAR As LongZ
Dim COUNT_UP
Dim FS
Dim OXS
Dim XDATE

Dim PATHSET




Private Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long


Private cProcesses As New clsCnProc

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'---------
Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
'---------

'---------------------------------------------------------------

'http://www.vbforums.com/showthread.php?673134-RESOLVED-Minimum-height-for-menu-bar-to-be-visible
'-------------- MENU SIZE DECLARE
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type MENUBARINFO
  cbSize As Long
  rcBar As RECT
  hMenu As Long
  hwndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type
Private MenuInfo As MENUBARINFO
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hwnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long


Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long


'------------------------------------
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long




 

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

Private Sub Form_Activate()

VARCENTER = False
'TAKES IT FROM HERE NOT FORM LOAD
'Call Form_Resize

End Sub

Private Sub Form_Load()

If App.PrevInstance = True Then End

'Set WSHShell = CreateObject("WScript.Shell")

Dim WW
MNU_CLIPBOARD_CONVERT_VODAFONE.Visible = False
O_WindowState = -8

WW = "D:\KAT MP3 RECORDER\KAT MP3 "
If GetComputerName = "1-ASUS-EEE" Then PATHSET = WW + GetComputerName
If GetComputerName = "3-LINDA-PC" Then PATHSET = WW + GetComputerName
If GetComputerName = "5-ASUS-P2520LA" Then PATHSET = WW + GetComputerName
If GetComputerName = "4-ASUS-GL522VW" Then PATHSET = WW + GetComputerName
If GetComputerName = "7-ASUS-GL522VW" Then PATHSET = WW + GetComputerName

If GetComputerName = "1-ASUS-X5DIJ" Then
    PATHSET = WW + GetComputerName
    Me.Caption = Me.Caption + " -- WinMerge Fault Working on"
End If

If PATHSET = "" Then PATHSET = WW + GetComputerName

If Dir(PATHSET, vbDirectory) = "" Then
    On Error Resume Next
    MkDir ("D:\KAT MP3 RECORDER\")
    Err.Clear
    MkDir (PATHSET)
    
    If Err.Number > 0 Then
        MsgBox "ERROR CREATE FOLDER" + vbCrLf + Err.Description + vbCrLf + PATHSET
        Unload Me
        Exit Sub
    End If
    On Error GoTo 0
End If

If Dir(PATHSET + "\# ARCHIVE\", vbDirectory) = "" Then
    On Error Resume Next
    MkDir ("D:\KAT MP3 RECORDER\")
    Err.Clear
    MkDir (PATHSET + "\# ARCHIVE\")
    If Err.Number > 0 Then
        MsgBox "ERROR CREATE FOLDER" + vbCrLf + Err.Description + vbCrLf + PATHSET + "\# ARCHIVE\"
        Unload Me
        Exit Sub
    End If
    On Error GoTo 0
End If

File5_ARCHIVE_KAT_MP3.Path = PATHSET + "\# ARCHIVE\"
File5_ARCHIVE_KAT_MP3.Pattern = "*.*"

If Dir(PATHSET, vbDirectory) = "" Then
    
'    If Dir(PATHSET, vbDirectory) = "" Then
'
'        PATHSET = Command$
'        PATHSET = Replace(PATHSET, """", "")
'
'    End If
    
    If Dir(PATHSET, vbDirectory) = "" Then
        MsgBox ("---" + PATHSET + "---" + vbCrLf + "PATH IS NOT A VALID PATH")
        End
    End If
End If

FILE1.Path = PATHSET

FILE1.FileName = "*.*"

KAT_MP3_PATH1 = "C:\Program Files\Kat MP3 Recorder\Kat MP3 Recorder.EXE"
KAT_MP3_PATH2 = "C:\Program Files (X86)\Kat MP3 Recorder\Kat MP3 Recorder.EXE"
If Dir(KAT_MP3_PATH1) <> "" Then
    KAT_MP3_PATH = KAT_MP3_PATH1
End If
If Dir(KAT_MP3_PATH2) <> "" Then
    KAT_MP3_PATH = KAT_MP3_PATH2
End If

Set FS = CreateObject("Scripting.FileSystemObject")

If IsIDE = False Then Me.WindowState = vbMinimized
If IsIDE = True Then Me.WindowState = vbMinimized
If IsIDE = True Then Me.WindowState = vbNormal

'X2 = 100

MNU_TIMER_SINCE_ACTIVITY.Caption = "TIMER SINCE ACTIVITY - WAIT FOR FIRST CHANGE"

'----------------------------------------------------------------------
'AT FORM LOAD THE TIMER INTERVAL IS SET AFTER FORM_LOAD AND USE DEFAULT
'----------------------------------------------------------------------
'UNLESS RESET WITH ENABLE FALSE AND TRUE
'----------------------------------------------------------------------
Timer_DESKTOP_INI_DEL.Enabled = False
Timer_DESKTOP_INI_DEL.Interval = 1000
Timer_DESKTOP_INI_DEL.Enabled = True

'-----------------
'TRY IN BAS MODUAL
'-----------------
'Call KAT_MP3_RELOAD_VAR.zzLoad_Checks
'---------------------------
'OR FORM - UNTIL GET WORKING
'---------------------------
Call Frm_Save_Load_Checks.zzLoad_Checks

Me.zzCheckTimer.Enabled = True

Timer_Find_Window.Enabled = True
Timer_Drives.Enabled = True

'If IsIDE = True Then MNU_VB_ME.Enabled = False

MNU_KAT_MP3_FOLDER.Caption = "KAT MP3 FOLDER EXPLORER --- " + PATHSET
MNU_KAT_MP3_FOLDER_FILE_LOCATOR_PRO.Caption = "KAY MP3 FOLDER -- FILE LOCATOR PRO --- " + PATHSET

Timer_VICE_VERSA_ME.Interval = 5000
Timer_VICE_VERSA_ME.Enabled = True

Call MNU_EXPLORER_SECURITY_INFO_2

Call TIMER_EXPLORER_IS_PRESENT_OPEN_AT_FOLDER_TRUE_TIMER

'------------------------------------
'THE SWITCHES SET REMEMBER THIS VALUE
'NOT TO LEAVE RUNNER ANYMORE
'------------------------------------
'MNU_CID_RERUN_ME_RELOADER.Checked = True
                    

MNU_W_DRIVE.Enabled = False

LOAD_FIRST_RUN_W_DRIVE_OFF = True
Call MNU_W_DRIVE_OFF_Click



MNU_VICE_VERSA_ME.Enabled = False
MNU_VICE_VERSA_LOCAL.Enabled = False
Form1.MNU_VICE_VERSA_ME.Enabled = False

'TIMER_KILL_CMD_PREVIOUS_INSTANCE_SOFTLY_WHEN_TOO_MANY.Enabled = True
LOAD_TIME1 = True
LOAD_TIME2 = True
LOAD_TIME3 = True

Call MNU_COMMAND_RESTRICTOR_Click

'Timer_COMMAND_RESTRICTOR_NOTEPAD_PICK_UP.Enabled = True

Call MNU_ADMINISTRATOR_Click

End Sub




'Private Sub MNU_SPECIAL_FOLDER_PIN_TO_START_MENU_Click()
'
'Call SPECIAL_FOLDER_PIN_TO_START_MENU
'
'Shell "Explorer.exe " + FOLDER_PIN_TO_START_MENU, vbNormalFocus
'
'End Sub

Private Sub SPECIAL_FOLDER_DESKTOP_INI()
'---------------------
'DESKTOP.INI -- DELETE
'---------------------
'IT KEEP GETTING PUT BACK IF COPY NEW ICON TO DESKTOP
'AND IF YOU WANT TO SHOW SYSTEM FILES AND FOLDER IT SHOW THIS DESKTOP.INI ALSO
'ON THE DESKTOP
'---------------------
'NOT SURE HOW IT WORK - BUT OKAY WITHOUT
'---------------------------------------
'Call SHOW_DEBUG_SPECIALS

'------------
'WIN 7 32 BIT
'------------
' 16 -- C:\Users\MATT 01\Desktop
' 25 -- C:\Users\Public\Desktop

If FS.FileExists(GetSpecialfolder(16) + "\DESKTOP.INI") = True Then
    FS.DeleteFile (GetSpecialfolder(16) + "\DESKTOP.INI")
End If
If FS.FileExists(GetSpecialfolder(25) + "\DESKTOP.INI") = True Then
    FS.DeleteFile (GetSpecialfolder(25) + "\DESKTOP.INI")
End If

'----------------
'----------------
'MUCK ABOUT WITH DIR AND VBHIDDEN - HASSEL
'Kill GetSpecialfolder(16) + "\DESKTOP.INI" - NOT WORK HIDDEN
'If Dir(GetSpecialfolder(16) + "\DESKTOP.INI", vbHidden) <> "" Then
'----------------
'----------------
'Exit Sub
'----------------
'----------------

'----------------
'FIND OUR FAT32 DRIVE BY SEARCH FOLDER SENDTO FOLDER
'----------------
'DEFAULT XP
FOLDER_FIND = "\01 SendTo\"
'----------------
'SIX = WIN 7
'----------------------------------
EXPLORER_LAUNCH = False
Call MNU_SEND_TO_FAT32_FOLDER_Click
EXPLORER_LAUNCH = True
'----------------------------------


FOLDER_SENDTO_SYSTEM = GetSpecialfolder(9)
'SHOW_DEBUG_SPECIALS

MNU_SEND_TO_FAT32_FOLDER.Caption = "EXPLORER -- " + FOLDER_SENDTO_FAT32_STORE
MNU_SEND_TO_SYSTEM_FOLDER.Caption = "EXPLORER -- " + FOLDER_SENDTO_SYSTEM

End Sub





Public Function GetWindowState(ByVal lngHwnd As Long) As Integer
    If IsWindow(lngHwnd) = 1 Then
        GetWindowState = -1
    If IsIconic(lngHwnd) <> 0 Then
        GetWindowState = vbMinimized
    ElseIf IsZoomed(lngHwnd) <> 0 Then
        GetWindowState = vbMaximized
    End If
    End If
        
    'RESULT 0 OR 1 -- 0 = vbMinimized
    'RESULT IN TRUE IS NOT VISIBLE SO EQUAL -2 FOR OUT RESULT
    If IsWindowVisible(lngHwnd) = 0 Then
        GetWindowState = -2
    End If

    
End Function



Private Sub Form_Resize()

If O_WindowState <> Me.WindowState Then
    If Me.WindowState = vbMinimized Or Me.WindowState = 2 Then
        MNU_CID_RELOADER.Visible = False

        Exit Sub
    End If
'    REM OKAY THIS CODE FOR HERE MAX AS CODE IS ONLY INNER FORM
'    If Me.WindowState = vbMaximized Then Exit Sub
End If

MNU_CID_RELOADER.Visible = True
TIMER_MNU_CID_RELOADER_DISAPPEAR.Enabled = True

O_WindowState = Me.WindowState

List1.Top = FILE1.Top + FILE1.Height
'REMEMBER FROM FIRST TIME SET
'----------------------------
If VAR_BOTTOM_OBJECT = 0 Then
    VAR_BOTTOM_OBJECT = List1.Height + List1.Top
End If



VAR_WIDTH_OBJECT = FILE1.Width
If List1.Width > VAR_WIDTH_OBJECT Then VAR_WIDTH_OBJECT = List1.Width

'------------------------------------------------
'LOAD THE FORM TO SIZE OF INNER FORM AND MENU BAR
'------------------------------------------------
'SOMETIME DO OTHER WAY AROUND
'LOAD THE INNER CONTROLS TO SIZE OUTER FORM
'------------------------------------------------

Me.Width = VAR_WIDTH_OBJECT + 190
Menu_Height2 = Menu_Height
'Debug.Print Menu_Height
'Menu_Height2 = 21

Me.Height = ((VAR_BOTTOM_OBJECT + (Menu_Height2 * Screen.TwipsPerPixelY)) + 500 + 120)

If VARCENTER = True Then Exit Sub

    Me.Top = Screen.Height / 2 - Me.Height / 2 '+400
    Me.Left = Screen.Width / 2 - Me.Width / 2 '+400
    
    VARCENTER = True

End Sub




Private Sub Form_Unload(Cancel As Integer)

UNLOAD_FORM_FLAG = True
'If IsIDE = True Then Stop

'Call KAT_MP3_RELOAD_VAR.zzCheckTimer_Timer

Dim Form As Form
For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next Form

End Sub


Private Sub MNU_AUTOHOTKEY_SAVE_AS_Click()

EXE_NAME = "C:\SCRIPTER CODE -- AUTOKEY\Autokey -- Save As Key Enter.ahk"
i = ShellExecute(Me.hwnd, "open", EXE_NAME, FileName, vbNullString, conSwNormal)
Me.WindowState = vbMinimized

End Sub

Public Sub MNU_BOOT_KILLER_Click()
'FORM1.MNU_BOOT_KILLER_Click
BEEP
Me.WindowState = vbMinimized
'Shell "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-BOOT KILLER.BAT", vbMaximizedFocus
Shell "CMD /C START """" /REALTIME /MAX ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01 BOOT KILLER.BAT""", vbMaximizedFocus
BEEP
End Sub
Public Sub MNU_TASK_KILLER_Click()

'FORM1.MNU_TASK_KILLER_Click
BEEP
If RUN_FROM_AUTO_TASK_KILLER = True Then
    RUN_FROM_AUTO_TASK_KILLER = False
    i = FindWindow(vbNullString, "Administrator:  C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 03 PROCESS KILLER.BAT")
    If i > 0 Then Exit Sub
End If

Me.WindowState = vbMinimized
'Shell "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 03 PROCESS KILLER.BAT", vbMaximizedFocus

Shell "CMD /C START """" /REALTIME /MAX ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 03 PROCESS KILLER.BAT""", vbMaximizedFocus

Shell "CMD /C START """" /REALTIME /MAX ""D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe""", vbMaximizedFocus
BEEP
End Sub

Private Sub MNU_GET_EXPLORER_CURRENT_SESSION_CLIPBOADIN_Click()
Call FindWindow_Get_All_Explorer
End Sub

'Private Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long
Function FindWindow_Get_All_Explorer() As String

FindWindow_Get_All_Explorer = ""

Dim WINDOW_TITLE
Dim Test_HWND As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String
Huge = 0

Test_HWND = FindWindow2(ByVal 0&, ByVal 0&)
BR$ = ""
Do While Test_HWND <> 0
    
    WINDOW_TITLE = GetWindowTitle(Test_HWND)
    
    '--------------------------------------------------
    'C:\Windows\explorer.exe
    '--------------------------------------------------
    If GetWindowClass(Test_HWND) = "CabinetWClass" Then
        Huge = Huge + 1
        BR$ = BR$ + WINDOW_TITLE + vbCrLf + vbCrLf
    End If
        
    'retrieve the next window
    Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)

Loop

BR2$ = "EXPLORER SESSION CURRENTLY WORKING ON MAYBE REBOOT OR TASKKILL" + vbCrLf
BR2$ = BR2$ + "C:\Windows\explorer.exe" + vbCrLf
BR2$ = BR2$ + "FIND ON CLASS NAME __CabinetWClass__" + vbCrLf
BR2$ = BR2$ + Trim(Str(Huge)) + " __ COUNT __ EXPLORER Window SET FOUND PASTE CLIPBOARD" + vbCrLf + vbCrLf
BR2$ = BR2$ + BR$

Clipboard.Clear
Clipboard.SetText BR2$
FindWindow_Get_All_Explorer = BR2$

MsgBox BR2$

End Function



Private Sub MNU_COMMAND_RESTRICTOR_Click()

On Error Resume Next
MkDir App.Path + "\DATA\"

FILENAME_PATH_COMMAND_RESTRICTOR = App.Path + "\DATA\COMMAND_RESTRICTOR_INSTANCE_LIMITER_SETTING" + GetComputerName + ".TXT"
If Dir(FILENAME_PATH_COMMAND_RESTRICTOR) = "" Then
    FR1 = FreeFile
    Open FILENAME_PATH_COMMAND_RESTRICTOR For Output As #FR1
        Print #FR1, "; ----------------------------------------------------"
        Print #FR1, "; WED 15 FEB 2017 12:28 A ____ MATT.LAN@BTINTERNET.COM"
        Print #FR1, "; ----------------------------------------------------"
        Print #FR1, "; REM WITH __ ; __ "
        Print #FR1, "; ONLY ONE SEARCH AT THE MOMENT"
        Print #FR1, "; OF CLASS NAME THE HWND WILL BE REQUEST CLOSE POST MESSAGE KILL "
        Print #FR1, "; PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)"
        Print #FR1, "; THE PARENT NOT SERACH FOR YET "
        Print #FR1, "; SPECIFY IN LINE ABOVE UNREMMER HOW MANY INSTANCE WANT TO BE RESTRICTED TO"
        Print #FR1, "; THE HWND KILL WILL BE OLDER FIRST AS THAT IS WAY THE DESKTOP HWND LAY OUT ARE"
        Print #FR1, "; IN THE USUAL MANNER BUT AND IT GOES HIGH AND SORT ALGORITHM"
        Print #FR1, "; SOMETIME USE A LOWER HWND NUMERIC OR HWND HIGH"
        Print #FR1, "; EXAMPLE KILL CMD WINDOW DOS PROMPT FOR USE HERE"
        Print #FR1, "; ----------------------------------------------------"
        Print #FR1, "; YOU MIGHT HAVE SOME OTHER HIDDEN AND BE LOST"
        Print #FR1, "; IF YOU WANT THIS FILE RESET HERE BACK THE DELETER"
        Print #FR1, "; ----------------------------------------------------"
        Print #FR1, "; EDIT FILE AND TIMER WILL PICK UP CHANGER AND THEN"
        Print #FR1, "; STOP CHECKING AFTER UNTIL LOAD TIME"
        Print #FR1, "; ----------------------------------------------------"
        Print #FR1, "; SWAP A SET PAIR IN BY REMMER A PAIR AWAY"
        Print #FR1, "; WANTED IDEA FOR CODING WITH DOS COMMANDER"
        Print #FR1, "; ----------------------------------------------------"
        Print #FR1, "; CLASS NAME _________________________________________"
        Print #FR1, "; & HOW MANY INSTANCE WANT GONE LEFT WITH AFTER ______"
        Print #FR1, "; ----------------------------------------------------"
        Print #FR1, "ConsoleWindowClass"
        Print #FR1, "4"
    Close #FR1
End If

If LOAD_TIME1 <> True Then
    NOTEPAD = "C:\Program Files (x86)\Notepad++\notepad++.exe"
    If Dir(NOTEPAD) = "" Then
        BEEP
        MsgBox "NOT EXIST NOETPAD++ __ USER NORMAL NOTEPAD INSTEAD ___ MESSENGER ONLY ONE WHILE LOAD " + vbCrLf + vbCrLf + NOTEPAD, vbMsgBoxSetForeground
    Else
        NOTEPAD_OKAY = True
    End If
        
    If NOTEPAD_OKAY = False Then
        NOTEPAD = "C:\Windows\System32\notepad.exe"
        If Dir(NOTEPAD) = "" Then
            BEEP
            MsgBox "NOT EXIST NORMAL NOTEPAD___ " + vbCrLf + vbCrLf + NOTEPAD, vbMsgBoxSetForeground
        End If
    End If
End If

If Dir(FILENAME_PATH_COMMAND_RESTRICTOR) = "" Then

    BEEP
    MsgBox "KAT_MP3_LOADER __ FILENAME __ HAD AN ERROR CREATING " + FILENAME_PATH_COMMAND_RESTRICTOR, vbMsgBoxSetForeground
    LOAD_TIME1 = False
    Exit Sub

End If

If LOAD_TIME1 = True Then
    LOAD_TIME1 = False
    
    'Set F1 = FS.getfile(FILENAME_PATH_COMMAND_RESTRICTOR)
    'COMMAND_RESTRICTOR_NOTEPAD_PICK_UP_DATE = F1.datelastmodified
    
    Exit Sub
End If
    
BEEP
Shell NOTEPAD + " " + FILENAME_PATH_COMMAND_RESTRICTOR, vbMaximizedFocus
BEEP

'Set F1 = FS.getfile(FILENAME_PATH_COMMAND_RESTRICTOR)
'COMMAND_RESTRICTOR_NOTEPAD_PICK_UP_DATE = F1.datelastmodified
Timer_COMMAND_RESTRICTOR_NOTEPAD_PICK_UP.Enabled = True

End Sub

Private Sub MNU_COMMAND_RESTRICTOR_ENABLED_Click()

If MNU_COMMAND_RESTRICTOR_ENABLED.Caption = "2 OF 2 _ COMMAND_RESTRICTOR_ENABLED_=_YES" Then
    
    MNU_COMMAND_RESTRICTOR_ENABLED.Caption = "2 OF 2 _ COMMAND_RESTRICTOR_ENABLED_=_NOT"
    TIMER_KILL_CMD_PREVIOUS_INSTANCE_SOFTLY_WHEN_TOO_MANY.Enabled = False
    Exit Sub
End If

If MNU_COMMAND_RESTRICTOR_ENABLED.Caption = "2 OF 2 _ COMMAND_RESTRICTOR_ENABLED_=_NOT" Then
    MNU_COMMAND_RESTRICTOR_ENABLED.Caption = "2 OF 2 _ COMMAND_RESTRICTOR_ENABLED_=_YES"
    TIMER_KILL_CMD_PREVIOUS_INSTANCE_SOFTLY_WHEN_TOO_MANY.Enabled = True
    COMMAND_RESTRICTOR_NOTEPAD_PICK_UP_DATE = 0
    Timer_COMMAND_RESTRICTOR_NOTEPAD_PICK_UP.Enabled = True
    'Call MNU_COMMAND_RESTRICTOR_Click
    Exit Sub
End If

End Sub

Private Sub MNU_NOTEPAD__BOOT_KILLER_Click()

On Error Resume Next

BEEP
Me.WindowState = vbMinimized
Shell "C:\Program Files (x86)\Notepad++\notepad++.exe ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-BOOT KILLER.BAT""", vbMaximizedFocusBEEP


End Sub


Private Sub MNU_CID_RERUN_ME_RELOADER_Click()
    
    '------------------------------------
    'THE SWITCHES SET REMEMBER THIS VALUE
    'NOT TO LEAVE RUNNER ANYMORE
    '------------------------------------
    
    MNU_CID_RERUN_ME_RELOADER.Checked = Not MNU_CID_RERUN_ME_RELOADER.Checked

End Sub

Private Sub MNU_CLIPBOARD_CONVERT_VODAFONE_Click()

BEEP

Exit Sub



On Error GoTo ENDE_SUB:

i = Clipboard.GetFormat(vbCFRTF)
i = Clipboard.GetFormat(vbCFDIB)
i = Clipboard.GetFormat(vbCFText)

If i = False Then BEEP: Exit Sub

'DIM VODA_F_CONVERT AS

'--------------------------------------------------------------
'THE BINARY PASTE IS SOMEWHAT DIFFERENT WITH &HE2 + &H80 + &HA9
'AND WHEN READ IN IS ALREADY CONVERTED TO A SINGLE &H3F
'SO WE SHALL USE THAT WORK RATHER THAN COMPLICATE
'--------------------------------------------------------------
'VODA_F_CONVERT = Replace(VODA_F_CONVERT, &HE2 + &H80 + &HA9, "----")
'REDUNDANT
'--------------------------------------------------------------

T = 24
'Debug.Print &H 'Hex("24")

For R = 1 To 100
    'Debug.Print Hex(Mid(VODA_F_CONVERT), R, 1)
Next

VODA_F_CONVERT = Clipboard.GetText
Do
    
    If InStr(VODA_F_CONVERT, &H3F) > 0 Then
        VODA_F_CONVERT = Replace(VODA_F_CONVERT, &H3F, "----")
    End If
Loop Until InStr(VODA_F_CONVERT, &H3F) = 0

Clipboard.Clear
Clipboard.SetText VODA_F_CONVERT
BEEP


Exit Sub

ENDE_SUB:

DoEvents
Sleep 100
BEEP
Resume

End Sub

Private Sub MNU_EXPLAIN_Click()

On Error Resume Next
Clipboard.Clear
Clipboard.SetText "MC - HX60V"

If Err.Number > 0 Then
    BEEP
    MsgBox "ERROR CLIPBOARDING -- TRY AGAIN -- BEEP WILL BE SUCCESS"
    Exit Sub
End If

BEEP

Me.WindowState = vbMinimized

End Sub

Private Sub MNU_EXPLORER_SECURITY_INFO_Click()

Call MNU_EXPLORER_SECURITY_INFO_2

Shell "EXPLORER " + App.Path + "\CAMERA DATA", vbNormalFocus

End Sub

Private Sub MNU_EXPLORER_SECURITY_INFO_2()

TEST_STR = "App.Path + ""\CAMERA DATA"""
MNU_EXPLORER_SECURITY_INFO.Caption = Replace(MNU_EXPLORER_SECURITY_INFO.Caption, TEST_STR, App.Path + "\CAMERA DATA")

End Sub

Private Sub MNU_FILE_DOWN_Click()

SLEEP_SPEED = 30
'MNU_FILE_UP.Caption = "** UP **"
x = FILE1.ListIndex
If x = -1 Then x = 0
If Len(FILE1.List(x)) > 23 Then
    For R = x + 1 To FILE1.ListCount - 1
        FILE1.ListIndex = R
        FILE1.Selected(FILE1.ListIndex) = True
        Me.Refresh
        If Len(FILE1.List(R)) <= 23 Then ESCAPE_VAR = True: Exit For
        MNU_FILE_DOWN.Caption = "DOWN -- LOOK FOR END BLOCK --" + Str(R) + " of" + Str(FILE1.ListCount - 1)
        Me.Refresh
        DoEvents
        Sleep SLEEP_SPEED
    Next
End If
If ESCAPE_VAR = True Then Exit Sub

x = FILE1.ListIndex
For R = x + 1 To FILE1.ListCount - 1
    FILE1.ListIndex = R
    FILE1.Selected(FILE1.ListIndex) = True
    Me.Refresh
    MNU_FILE_DOWN.Caption = "DOWN -- LOOK FOR BEGIN --" + Str(R) + " of" + Str(FILE1.ListCount - 1)
    If Len(FILE1.List(R)) > 23 Then Exit For
    Me.Refresh
    DoEvents
    Sleep SLEEP_SPEED
Next

End Sub

Private Sub MNU_FILE_UP_Click()
SLEEP_SPEED = 30
'MNU_FILE_DOWN.Caption = "** UP **"

x = FILE1.ListIndex
If x >= FILE1.ListCount Then x = FILE1.ListCount - 1
If Len(FILE1.List(x)) > 23 Then
    For R = x - 1 To 0 Step -1
        FILE1.ListIndex = R
        FILE1.Selected(FILE1.ListIndex) = True
        MNU_FILE_UP.Caption = "UP -- LOOK FOR END BLOCK --" + Str(R) + " of" + Str(FILE1.ListCount - 1)
        If Len(FILE1.List(R)) <= 23 Then ESCAPE_VAR = True: Exit For
        Me.Refresh
        DoEvents
        Sleep SLEEP_SPEED
    Next
End If
If ESCAPE_VAR = True Then Exit Sub

x = FILE1.ListIndex
For R = x - 1 To 0 Step -1
    FILE1.ListIndex = R
    FILE1.Selected(FILE1.ListIndex) = True
    MNU_FILE_UP.Caption = "UP -- LOOK FOR BEGIN --" + Str(R) + " of" + Str(FILE1.ListCount - 1)
    If Len(FILE1.List(R)) > 23 Then Exit For
    DoEvents
    Me.Refresh
    Sleep SLEEP_SPEED
Next

End Sub

Private Sub MNU_PASTE_VOLUME_LABEL_CAMERA_Click()

    Call MNU_EXPLAIN_Click

End Sub

Private Sub MNU_SECURITY_TO_CAMERA_Click()
    TIMER_SET_CAMERA_WITH_SECURITY_INFO.Enabled = True
End Sub


Private Sub MNU_VICE_VERSA_LOCAL_Click()

    Call MNU_VICEVERSA_LOCAL_Click
    
    Me.WindowState = vbMinimized

End Sub

Private Sub MNU_VICE_VERSA_ME_CLICK()
    
    Form1.MNU_VICE_VERSA_ME.Visible = False

    Call MNU_VICEVERSA_ME_Click
    
    Me.WindowState = vbMinimized

End Sub


Private Sub MNU_ENABLE_FINDWINOW_TIMER_VAR_Click()

MNU_ENABLE_FINDWINOW_TIMER_VAR.Checked = Not MNU_ENABLE_FINDWINOW_TIMER_VAR.Checked

Timer_Find_Window.Enabled = MNU_ENABLE_FINDWINOW_TIMER_VAR.Checked
'Timer_Find_Window
End Sub

Private Sub MNU_EXIT_Click()
Unload Me
End Sub

Private Sub MNU_FILE_LOCATOR_Click()

    PID = Shell("C:\Program Files\Mythicsoft\FileLocator Pro\FileLocatorPro.exe", vbMaximizedFocus)

End Sub

Private Sub MNU_KAT_MP3_FOLDER_Click()
    MNU_KAT_MP3_FOLDER.Caption = "KAT MP3 FOLDER EXPLORER " + PATHSET
    PID = Shell("EXPLORER " + PATHSET, vbNormalNoFocus)
    
End Sub

Private Sub MNU_KAT_MP3_FOLDER_FILE_LOCATOR_PRO_Click()
    
    '------------------------------------------------
    'KEY TAPPING CROSS BETWEEN COWBOY CHINK BOOTS AND
    'BODY SNATCHER GRAVE BURIAL ALIVE
    '------------------------------------------------
    
    MNU_KAT_MP3_FOLDER_FILE_LOCATOR_PRO.Caption = "KAY MP3 FOLDER -- FILE LOCATOR PRO " + PATHSET

    PID = Shell("C:\Program Files\Mythicsoft\FileLocator Pro\FileLocatorPro.exe """ + PATHSET + """", vbMaximizedFocus)

End Sub

Private Sub MNU_KAT_MP3_RUN_Click()

    ' COUNT_UP = 0
    'WHY
    FIRST_LIMIT_RESET_DONE = True
    FIRST_RUN_FIND_PID_KAT_MP3 = True
    
    'WHY
    GO3 = 1
    
    ' TIME_HAPPEN --- COMMON VAR
    i = TIME_HAPPEN + " -- Kat MP3 Recorder.EXE - REQUEST TO RUN EXECUTED"
    
    ' PID HAS TO BE -1
    PID = -1
    Var = cProcesses.GetEXEID(PID, "C:\Program Files\Kat MP3 Recorder\Kat MP3 Recorder.EXE")
    If PID = -1 Then
        
        Shell KAT_MP3_PATH, vbMinimizedNoFocus
        COUNT_UP = Now
    
        Exit Sub
    Else
        i = TIME_HAPPEN + " -- Kat MP3 Recorder.EXE - REQUEST TO RUN EXECUTED -- ALREADY RUNNER"
    End If
    
    Label1.Caption = i
    List1.AddItem Label1.Caption
    
    BEEP
    'WHY
    GO3 = 0

End Sub

Private Sub Mnu_Mixer_Deck_Tunes_K_Lite_Player_Classic_Click()

'19:59:58 - TIME CD# 2
'19:59:59 - TIME CD# 1
'20:00:04 - Wednesday 10th of August
'20:01:19 - Day Light! 97%
'19:01:00 - MOON!,  55.34 %, Swell!
'19:10:39 - Day Light! 91.3%
'19:11:00 - TIME #07:11
'19:11:54 - MOON!,  55.42 %, Swell!
'19:30:04 - Wednesday 10th of August
'19:30:04 - Wednesday 10th of August
'19:45:00 - TIME #Quarter to  8
'19:45:19 - Day Light! 95.2%
'19:47:59 - Day Light! 95.5%
'19:59:58 - TIME CD# 2
'19:59:59 - TIME CD# 1
'20:17:17 - MOON!,  55.9 %, Swell!
'20:19:59 - Day Light! 99.1%
'20:22:44 - MOON!,  55.94 %, Swell!
'20:28:00 - TIME #08:28
'20:28:00 - Day Light! 100%
'20:28:11 - MOON!,  55.98 %, Swell!
'20:28:26 - Sun Set . 8   28 and  26 Seconds
'20:28:27 - DarkNess! 0%
'20:29:16 - DarkNess! .2%
'20:30:00 - TIME #08:30

'---- Play Mixer Tunes -- hifi www.stream1.nubreaks.com.m3u -- With -- \Program Files (x86)\K-Lite Codec Pack\MPC-HC64\mpc-hc64.exe

Dim I1, I2

I1 = """D:\0 00 MUSIC ---\04 My Music\01 Banging Tunes\00 Special\Sp 01 Will De Cruize - Ol Mate\# 00 hifi www.stream1.nubreaks.com.m3u"""
I2 = """C:\Program Files (x86)\K-Lite Codec Pack\MPC-HC64\mpc-hc64.exe"""
Shell I2 + " " + I1, vbNormalFocus

Me.WindowState = vbMinimized

End Sub

Public Sub MNU_OFFICE_TOOLBAR_USE_Click()

    MNU_OFFICE_TOOLBAR_USE.Checked = Not MNU_OFFICE_TOOLBAR_USE.Checked

'    Call KAT_MP3_RELOAD_VAR.zzCheckTimer_Timer
    'Call Frm_Save_Load_Checks.zzLoad_Checks
    Call Frm_Save_Load_Checks.zzCheckTimer_Timer
    Call zzCheckTimer_Timer
End Sub

Private Sub MNU_SEND_TO_FAT32_FOLDER_Click()

'----------------
'FIND OUR FAT32 DRIVE BY SEARCH FOLDER SENDTO FOLDER
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

FOLDER_FIND = FOLDER_FIND + "\"

'----------------
FOLDER_SENDTO_FAT32_STORE = GET_DRIVES_FIND_FOLDER(FOLDER_FIND)
'----------------

'---------------------------------------
If EXPLORER_LAUNCH = False Then Exit Sub

Shell "Explorer.exe " + FOLDER_SENDTO_FAT32_STORE, vbNormalFocus
'---------------------------------------

End Sub

Private Sub MNU_SEND_TO_SYSTEM_FOLDER_Click()
'FOLDER_SENDTO_FAT32_STORE = GET_DRIVES_FIND_FOLDER(FOLDER_FIND)
FOLDER_SENDTO_SYSTEM = GetSpecialfolder(9)

Shell "Explorer.exe " + FOLDER_SENDTO_SYSTEM, vbNormalFocus

End Sub

Private Sub MNU_TIMER_SINCE_ACTIVITY_Click()

'MNU_TIMER_SINCE_ACTIVITY.Caption = "TIMER SINCE ACTIVITY - WAIT FOR FIRST CHANGE"

BEEP: MsgBox "NOT HERE DONE YET"

End Sub

'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".vbp"
    i = CODER_VBP_FILE_NAME_2
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
Private Sub Mnu_VB_ME_Click()
    BEEP
    If GetSpecialfolder(38) = "" Then
        MsgBox "YOU HAVE TO SWITCH ADMIN MODE ON" + vbCrLf + vbCrLf + "THE COMMAND GetSpecialfolder(38) REQUIRE ADMINISTRATOR" + vbCrLf + vbCrLf + "RESULT RETURN NOTHING HERE GetSpecialfolder(38)__<" + GetSpecialfolder(38) + ">__ RESULT RETURNED NOTHING", vbMsgBoxSetForeground
        Exit Sub
    End If
    
    If IsIDE = True Then
        BEEP
        If IsIDE = True Then
            MSGQ = vbYesNoCancel
        Else
            MSGQ = vbOKOnly
        End If
        IR = MsgBox("NOT IN IDE ARE YOU SURE __ CANCLE TO STOP INSTEAD", MSGQ, vbMsgBoxSetForeground)
        If IR = vbCancel Then Stop
        If IR = vbNo Then Exit Sub
    End If
    
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
    Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + CODER_VBP_FILE_NAME_2 + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
End Sub
Private Sub MNU_VB_FOLDER_Click()
    BEEP
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    Shell "EXPLORER /SELECT, """ + CODER_VBP_FILE_NAME_2 + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
End Sub
Private Sub MNU_ADMINISTRATOR_Click()
    'Call MNU_ADMINISTRATOR_Click
    BEEP
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


Sub FIND_VICE_VERSA()
    
    VVCOMLINE1 = GetSpecialfolder(38) + "\VICEVERSA PRO\VICEVERSA.EXE"
    If Dir(VVCOMLINE1) <> "" Then
        VVCOMLINE = VVCOMLINE1
    End If
    
    VVCOMLINE2 = Replace(UCase(GetSpecialfolder(38)), " (X86)", "") + "\VICEVERSA PRO\VICEVERSA.EXE"
    If Dir(VVCOMLINE2) <> "" Then
        VVCOMLINE = VVCOMLINE2
    End If
    
    If VVCOMLINE = "" Then
        MsgBox "UNABLE TO FIND VICEVERSA PROGRAM @" + vbCrLf + VVCOMLINE1 + vbCrLf + "OR" + vbCrLf + VVCOMLINE2
        Exit Sub
    End If

End Sub


Private Sub MNU_VICE_VERSA_NETWORK_SYNCRO_Click()
    Call FIND_VICE_VERSA
    
    VICE_VERSA_SCRIPT_FILE = "C:\ViceVersa PRO\ASUS G TO ASUS P QUICK SYNC FOLDERING.fsf"
    If Dir(VICE_VERSA_SCRIPT_FILE) = "" Then
    
        MsgBox "UNABLE TO FIND" + vbCrLf + VICE_VERSA_SCRIPT_FILE + vbCrLf, vbMsgBoxSetForeground
        Exit Sub
    
    End If
    
    PID = Shell(VVCOMLINE + " """ + VICE_VERSA_SCRIPT_FILE + """", vbNormalFocus)

    Me.WindowState = vbMinimized

End Sub

Private Sub MNU_VICE_VERSA_TIMER_ON_Click()

 
Timer_VICE_VERSA_TIMER_ON.Enabled = True
If MNU_VICE_VERSA_TIMER_ON.Caption <> "VICE-VERSA-TIMER-ON" Then
    MNU_VICE_VERSA_TIMER_ON.Caption = "VICE-VERSA-TIMER-ON"
    'MNU_VICE_VERSA_TIMER_ON.Caption = "VICE-VERSA-TIMER-OFF-PRESS-BUTTON-TO-GO"
    
    'MNU_VICE_VERSA_TIMER_ON.Caption = "VICE-VERSA-TIMER-ON"
    Timer_VICE_VERSA_TIMER_ON.Enabled = False
    MNU_VICE_VERSA_TIMER_ON.Caption = "VICE-VERSA-TIMER-OFF__PRESS-BUTTON-TO__GO"

End If

End Sub

Private Sub MNU_VICEVERSA_EXPLORER_Click()

    If Dir(CAMERA_EXPLORER_PATH, vbDirectory) = "" Then
        MsgBox "W DRIVE FOLDER " + CAMERA_EXPLORER_PATH + " -- DOES NOT EXIST - ABANDON LOAD EXPLORER"
        Exit Sub
    End If
    
    Shell "EXPLORER ""W:\""", vbNormalFocus

    Me.WindowState = vbMinimized

End Sub

Private Sub MNU_VICEVERSA_LOCAL_Click()

    Call FIND_VICE_VERSA
    'D:\VICE_VERSA SCRIPT FILE\ViceVersa D DRIVE - WHOLE FOLDER SETS - QUICK.fsf
    VICE_VERSA_SCRIPT_FILE = "D:\VICE_VERSA SCRIPT FILE\ViceVersa PHOTO CAMERA - LOCAL.fsf"
    If Dir(VICE_VERSA_SCRIPT_FILE) = "" Then
    
        MsgBox "UNABLE TO FIND" + vbCrLf + VICE_VERSA_SCRIPT_FILE + vbCrLf, vbMsgBoxSetForeground
        Exit Sub
    
    End If
    
    PID = Shell(VVCOMLINE + " """ + VICE_VERSA_SCRIPT_FILE + """", vbNormalFocus)

    Me.WindowState = vbMinimized

End Sub

Private Sub MNU_VICEVERSA_ME_Click()
    Form3_Me_ViceVersa.Show
    Me.WindowState = vbMinimized
End Sub

Private Sub MNU_VICEVERSA_REMOTE_Click()

    Call FIND_VICE_VERSA
    
    VICE_VERSA_SCRIPT_FILE = "D:\VICE_VERSA SCRIPT FILE\ViceVersa PHOTO CAMERA - REMOTE.fsf"
    If Dir(VICE_VERSA_SCRIPT_FILE) = "" Then
    
        MsgBox "UNABLE TO FIND" + vbCrLf + VICE_VERSA_SCRIPT_FILE + vbCrLf, vbMsgBoxSetForeground
        Exit Sub
    
    End If
    
    PID = Shell(VVCOMLINE + " """ + VICE_VERSA_SCRIPT_FILE + """", vbNormalFocus)
    
    Me.WindowState = vbMinimized

End Sub

Private Sub MNU_VICEVERSA_W_SUBST_Click()
    
    If MNU_W_DRIVE_OFF.Caption = "W DRIVE OFF && GO" Then Exit Sub

    
    'I = GET_DRIVES_FOR_CAMERA
    If VAR_SHELL_SUBST_BEEN_SET = False Then
        MsgBox "THE W DRIVE SUBST IS NOT MADE - MAYBE UNABLE TO FIND CAMERA VOLUME NAME", vbMsgBoxSetForeground
    Else
        MsgBox "THE W DRIVE SUBST CREATED SUCCESS", vbMsgBoxSetForeground
    End If

End Sub

Private Sub MNU_W_DRIVE_OFF_Click()
    
    If LOAD_FIRST_RUN_W_DRIVE_OFF = True Then
        LOAD_FIRST_RUN_W_DRIVE_OFF = False
        MNU_W_DRIVE_OFF.Caption = "W DRIVE OFF && GO"
        Timer_VICE_VERSA_TIMER_ON.Enabled = False
        MNU_VICE_VERSA_TIMER_ON.Caption = "VICE-VERSA-TIMER-OFF__PRESS-BUTTON-TO__GO"
        'VAR_SHELL_SUBST_BEEN_SET = False
        BEEP
        
        If Dir("W:\", vbDirectory) <> "" Then Shell "SUBST W: /D", vbHide
        If Dir("MC - HX60V", vbVolume) <> "" Then Shell "SUBST W: /D", vbHide

        Exit Sub
    End If
    
    'W DRIVE OFF && GO
    If MNU_W_DRIVE_OFF.Caption = "W DRIVE OFF && GO" Then
        MNU_W_DRIVE_OFF.Caption = "W DRIVE GO"
        Timer_VICE_VERSA_TIMER_ON.Enabled = True
        MNU_VICE_VERSA_TIMER_ON.Caption = "VICE-VERSA-TIMER-ON"
        'VAR_SHELL_SUBST_BEEN_SET = False
        BEEP
        Exit Sub
    End If
    
    If MNU_W_DRIVE_OFF.Caption = "W DRIVE GO" Then
        MNU_W_DRIVE_OFF.Caption = "W DRIVE OFF && GO"
        
        BEEP
'        Exit Sub
    End If
    
    
    Shell "SUBST W: /D", vbHide
    VAR_SHELL_SUBST_BEEN_SET = False
    BEEP
 
    'Timer_VICE_VERSA_TIMER_ON.Enabled = False
    'If MNU_VICE_VERSA_TIMER_ON.Caption <> "VICE-VERSA-TIMER-ON" Then
        'MNU_VICE_VERSA_TIMER_ON.Caption = "VICE-VERSA-TIMER-ON"
        Timer_VICE_VERSA_TIMER_ON.Enabled = False
        MNU_VICE_VERSA_TIMER_ON.Caption = "VICE-VERSA-TIMER-OFF__PRESS-BUTTON-TO__GO"
    'End If


End Sub

Private Sub MNU_WAV_ACTIVITY_Click()
BEEP:     MsgBox "NONE HERE YET"
End Sub



Private Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim l As Long




   Dim S As String
   l = GetWindowTextLength(hwnd)
   S = Space(l + 1)
   GetWindowText hwnd, S, l + 1
   GetWindowTitle = Left$(S, l)
   
End Function

Private Function GetText(pHandle As Long) As String
       
    Dim Buffer As String
    Dim TextLength As Long
   
    TextLength = SendMessageAny(pHandle, WM_GETTEXTLENGTH, 0&, 0&)

    Buffer = String$(TextLength, 0&)
    SendMessageAny pHandle, WM_GETTEXT, TextLength + 1, Buffer
    GetText = Buffer
End Function




Sub SOME_THING_WRONG_DURING_MERGE_CONFLICT()
    Exit Sub
    
    '--------------------------------------------
    'WRONG MISSING PART OF CODE ABOVE
    'HAVE TO SHUT DOWN THESE ROUTINE FOR A MOMENT
    '--------------------------------------------
  
  
    '-----------------------------------------------------------
    'CHECK CLASS NAME IS VB IDE WINDOW WE WANT OR IF ANOTHER NOT
    '-----------------------------------------------------------
    'X2 = GetWindowTitle(X1)
    
    'I = FindWinPart("Tidal_Information - Microsoft Visual Basic [break]")
    
    
    
    '----------------------------------
    'TO BE REM OUT TEMPORALLY THE BELOW
    '----------------------------------
    
'    I = FindWinPart("Tidal_Information - Microsoft Visual Basic [break]")
'
'    If I > 0 Then
'        If Tidal_Information_Microsoft_Visual_Basic_break_ = False Then
'            Tidal_Information_Microsoft_Visual_Basic_break_ = True
'            BEEP
'            Timer_Tidal_Information_Microsoft_Visual_Basic_break_SubRoutine.Enabled = True
'        End If
'    End If
'End If
'
'If I = 0 Then
'    Tidal_Information_Microsoft_Visual_Basic_break_ = False
'    Timer_Tidal_Information_Microsoft_Visual_Basic_break_SubRoutine.Enabled = False
'End If

End Sub


Public Function WindowText(ByVal window_hwnd As Long) As String

    Dim txtlen              As Long

    If window_hwnd = 0 Then
        Exit Function
    End If

    txtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, ByVal 0, ByVal 0)
    If txtlen = 0 Then
        Exit Function
    End If

    WindowText = String$(txtlen, vbNullChar)
    SendMessage window_hwnd, WM_GETTEXT, ByVal (txtlen + 1), ByVal StrPtr(WindowText)

End Function


Private Sub Timer_01_VB_HELP_BOX_02_MSDN_Timer()

' ----------------------------------
' ALTERNATIVE LIKE HERE
' ----------------------------------
' Timer_01_VB_HELP_BOX_02_MSDN_Timer
' ---- OTHER -----------------------
' Timer_VB_MAXIMIZE
' ---- ALL THE VB ------------------
' ----------------------------------

'--------------------------------------------------------------------------------
'1..
'FIREFOX ADD AND SERACH ENGINE WITHOUT QUESTION ADD
'2..
'GOODSYNC
'GOOD_SYNC_MSGBOX_AT_END_IF_JOB_S_RUNNING_HWND = FindWindow("#32770", "GoodSync")
'3..
'ICACLS -- Error Applying Security
'--------------------------------------------------------------------------------


'FIREFOX ADD AND SERACH ENGINE WITHOUT QUESTION ADD
'--------------------------------------------------
FINDER = FindWindow("MozillaDialogClass", "Add Search Engine")
If FINDER > 0 And FINDER <> OHWnd_FINDER_1 Then

    '---------------------------------------
    'RESULT = PostMessage(FINDER, WM_CLOSE, 0&, 0&)
    '---------------------------------------
    'HWND_2 = GET_CHILD_TEST(MSDN_Lib)
    '---------------------------------------
    
'    cWnd = FindWindowEx(pwnd, 0, "BUTTON", "HELP")
'    EnableWindow cWnd, False

    pwnd = FINDER
    'GS_cWnd4 = 0
    'GS_cWnd1 = FindWindowEx(pwnd, 0, "Edit", vbNullString)
    'GS_cWnd2 = FindWindowEx(pwnd, 0, "Button4", "&OK")
    'GS_cWnd3 = FindWindowEx(pwnd, 0, "button", "&Cancel")
    
    GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
    
        
    If GS_cWnd2 > 0 Then
        '-------------------------------------------------
        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
        '-------------------------------------------------
        For R_REPEAT = 1 To 10
            GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
            If GS_cWnd2 = 0 Then Exit For
            SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
    
End If

OHWnd_FINDER_1 = FINDER






'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'GOODSYNC
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------




Dim GOOD_SYNC_MSGBOX_CRASH_HWND As Long

'GOODSYNC CRASH kill task not create report confidential and too large zip create
'and also hold up not continue happen restarter
'--------------------------------------------------------------------------------
'TESTER
'------
'GOOD_SYNC_MSGBOX_CRASH_HWND = FindWinPart("GoodSync2Go")
'--------------------------------------------------------------------------------
GOOD_SYNC_MSGBOX_CRASH_HWND = FindWindow("#32770", "Preparing Crash Report - GoodSync")
i3 = GOOD_SYNC_MSGBOX_CRASH_HWND
If i3 = 0 Then GOOD_SYNC_MSGBOX_CRASH_HWND = FindWindow("#32770", "GoodSync Crash")
i3 = GOOD_SYNC_MSGBOX_CRASH_HWND
If i3 > 0 And i3 <> OcWnd_GOOD_SYNC_CRASH Then
    
    'PID = -1
    'Var = cProcesses.GetEXEID(PID, "C:\Program Files\WinMerge\WinMergeU.exe")
    'Var = cProcesses.Convert(PID, OTSS, cnFromProcessID Or cnTohWnd)
    'Var = cProcesses.Convert(i3, PID, cnFromhWnd Or cnToProcessID)
    'Var = cProcesses.Convert(i3, PID, cnFromhWnd Or cnToProcessID)
    'Var = cProcesses.Convert(i3, PID, cnFromhWnd Or cnToProcessIDcnToEXE)
    'PID = -1
    'Var = cProcesses.Convert(i3, EXE_STRING, cnFromhWnd Or CNToEXE)
    
    '---------------------------------------------------------------
    '2017
    'LEACHED FROM ELITESPY TO GET PROPER FULL EXE NAME
    'INCLUDED IN CLASS ROUTINE TOGETHER
    'THIS IS NOT ENOUGH WHEN BE SHORT EXE NAME -- cnToEXE
    'MAYBE UPDATE WITH OTHER CODE NEARBY
    'NORMAL ROUTINE BUT ADD THE cProcesses. BECUASE IT IS IN A CLASS
    'AND CLASS MUST BE INITALISED -- AS IT IS DONE
    '---------------------------------------------------------------
    EXE_STRING = cProcesses.GetFileFromHwnd(i3)
    
    'Stop
    
    PID = -1
    Var = cProcesses.Convert(i3, PID, cnFromhWnd Or cnToProcessID)
    
    'Stop
    
    '---------------------------------------------------------------------
    'WORKING BUT NOT FULL EXE NAME PATH
    'Var = cProcesses.Convert(PID, EXE_STRING, cnFromProcessID Or cnToEXE)
    'Var = cProcesses.Convert(i3, PID, cnFromhWnd Or cnToEXE)
    '---------------------------------------------------------------------
    'VAR = GetEXEID(PID, ByRef sRunningEXE As String) As Boolean
    '---------------------------------------------------------------------
    
    
    
    Var = cProcesses.Process_Kill(PID)
    
    Sleep 1000
    
    Shell EXE_STRING, vbMinimizedNoFocus
    
    
    
    
    '---------------------------------------
    'RESULT = PostMessage(MSDN_Lib, WM_CLOSE, 0&, 0&)
    'Microsoft Visual Basic
    '#32770
    '---------------------------------------
    'INSTEAD OF METHOD 1 USE METHOD 2
    '---------------------------------------
    '1. DONT CLOSE THE HELP ALERT VB INFO
    '2. DISABLE THE HELP BUTTON ACIDENT FLICKER
    '---------------------------------------
    'HWND_2 = GET_CHILD_TEST(MSDN_Lib)
    '---------------------------------------
    'TEXT1 = GetWindowText GS_cWnd1, S, l + 1
    'TEXT1 = GetWindowTitle(GS_cWnd1)
    
    '----------------------------------------------------------------
    'TO GET THE TEXT OF A BUTTON OR MSGBOX ORGINAL SOURCE CREDIT HERE
    '----
    'GET THE TEXT OF A BUTTON WITH HWND IN VB6 - Google Search
    'https://www.google.co.uk/search?q=GET+THE+TEXT+OF+A+BUTTON+WITH+HWND+IN+VB6&oq=GET+THE+TEXT+OF+A+BUTTON+WITH+HWND+IN+VB6+&aqs=chrome..69i57.9997j0j7&sourceid=chrome&ie=UTF-8
    '--------
    'How you get the hwnd's of a text box or button of another window?-VBForums
    'http://www.vbforums.com/showthread.php?576117-How-you-get-the-hwnd-s-of-a-text-box-or-button-of-another-window
    '----
    '----------------------------------------------------------------
    
'    Dim GS_cWnd1 As Long
'
'    pwnd = i3
'    GS_cWnd4 = 0
'    GS_cWnd1 = FindWindowEx(pwnd, 0, "Edit", vbNullString) '"One or more jobs are running now:")
'    TEXT1 = GetText(GS_cWnd1)
'    If InStr(TEXT1, "One or more jobs are running now:") > 0 Then GS_cWnd4 = 1
'
'    GS_cWnd2 = FindWindowEx(pwnd, 0, "Button", "&OK")
'    GS_cWnd3 = FindWindowEx(pwnd, 0, "button", "&Cancel")
    
    
    'ENABLE BUTTON = FALSE -- LEARN BEOFRE
    'EnableWindow cWnd, False
    
    
    '---------------------------------------------------------
    'Const BM_CLICK = &HF5&
    
    'CREDIT TO LEARN THE PUSH BUTTON WITH THESE 2 OR 3 COMMAND
    
    '----
    'POSTMESSAGE TO PRESS BUTTON ON A MSGBOX VB6 - Google Search
    'https://www.google.co.uk/search?q=POSTMESSAGE+TO+PRESS+BUTTON+ON+A+MSGBOX+VB6&oq=POSTMESSAGE+TO+PRESS+BUTTON+ON+A+MSGBOX+VB6&aqs=chrome..69i57.13373j0j7&sourceid=chrome&ie=UTF-8
    '--------
    'VBA/VB.Net/VB6–Click Open/Save/Cancel Button on IE Download window – PART I
    'http://www.siddharthrout.com/2011/10/23/vbavb-netvb6click-opensavecancel-button-on-ie-download-window/
    '----
    '---------------------------------------------------------
    
'    If GS_cWnd2 > 0 And GS_cWnd3 > 0 And GS_cWnd4 = 1 Then
'
'        '-------------------------------------------------
'        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
'        '1ST RUN DEBUG FOCUSED TO BUTTON BUT NOT HITT IT IN UNTIL 2ND
'        'FOCUS PROBLEM FIRST
'        '-------------------------------------------------
'        For R_REPEAT = 1 To 3
'            SendMessage GS_cWnd2, BM_CLICK, 0, 0
'        Next
'    End If
    
    '---------------------------------------
    '31 AUG 2K SIXTEEN
    'LEARN
    '---------------------------------------
    'Find Window Tutorial Part1 <A must see> by Paul Zaczkowski (from psc cd)
    'https://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37467&lngWId=1
    '---------------------------------------
End If

OcWnd_GOOD_SYNC_CRASH = i3




'GOODSYNC
GOOD_SYNC_MSGBOX_AT_END_IF_JOB_S_RUNNING_HWND = FindWindow("#32770", "GoodSync")
i3 = GOOD_SYNC_MSGBOX_AT_END_IF_JOB_S_RUNNING_HWND
If i3 >= 0 And i3 <> OcWnd_GOOD_SYNC Then

    '---------------------------------------
    'RESULT = PostMessage(MSDN_Lib, WM_CLOSE, 0&, 0&)
    'Microsoft Visual Basic
    '#32770
    '---------------------------------------
    'INSTEAD OF METHOD 1 USE METHOD 2
    '---------------------------------------
    '1. DONT CLOSE THE HELP ALERT VB INFO
    '2. DISABLE THE HELP BUTTON ACIDENT FLICKER
    '---------------------------------------
    'HWND_2 = GET_CHILD_TEST(MSDN_Lib)
    '---------------------------------------
    'TEXT1 = GetWindowText GS_cWnd1, S, l + 1
    'TEXT1 = GetWindowTitle(GS_cWnd1)
    
    '----------------------------------------------------------------
    'TO GET THE TEXT OF A BUTTON OR MSGBOX ORGINAL SOURCE CREDIT HERE
    '----
    'GET THE TEXT OF A BUTTON WITH HWND IN VB6 - Google Search
    'https://www.google.co.uk/search?q=GET+THE+TEXT+OF+A+BUTTON+WITH+HWND+IN+VB6&oq=GET+THE+TEXT+OF+A+BUTTON+WITH+HWND+IN+VB6+&aqs=chrome..69i57.9997j0j7&sourceid=chrome&ie=UTF-8
    '--------
    'How you get the hwnd's of a text box or button of another window?-VBForums
    'http://www.vbforums.com/showthread.php?576117-How-you-get-the-hwnd-s-of-a-text-box-or-button-of-another-window
    '----
    '----------------------------------------------------------------
    
    Dim GS_cWnd1 As Long
    
    pwnd = i3
    GS_cWnd4 = 0
    GS_cWnd1 = FindWindowEx(pwnd, 0, "Edit", vbNullString) '"One or more jobs are running now:")
    TEXT1 = GetText(GS_cWnd1)
    If InStr(TEXT1, "One or more jobs are running now:") > 0 Then GS_cWnd4 = 1
    
    GS_cWnd2 = FindWindowEx(pwnd, 0, "Button", "&OK")
    GS_cWnd3 = FindWindowEx(pwnd, 0, "button", "&Cancel")
    
    
    'ENABLE BUTTON = FALSE -- LEARN BEOFRE
    'EnableWindow cWnd, False
    
    
    '---------------------------------------------------------
    'Const BM_CLICK = &HF5&
    
    'CREDIT TO LEARN THE PUSH BUTTON WITH THESE 2 OR 3 COMMAND
    
    '----
    'POSTMESSAGE TO PRESS BUTTON ON A MSGBOX VB6 - Google Search
    'https://www.google.co.uk/search?q=POSTMESSAGE+TO+PRESS+BUTTON+ON+A+MSGBOX+VB6&oq=POSTMESSAGE+TO+PRESS+BUTTON+ON+A+MSGBOX+VB6&aqs=chrome..69i57.13373j0j7&sourceid=chrome&ie=UTF-8
    '--------
    'VBA/VB.Net/VB6–Click Open/Save/Cancel Button on IE Download window – PART I
    'http://www.siddharthrout.com/2011/10/23/vbavb-netvb6click-opensavecancel-button-on-ie-download-window/
    '----
    '---------------------------------------------------------
    
    If GS_cWnd2 > 0 And GS_cWnd3 > 0 And GS_cWnd4 = 1 Then
        
        '-------------------------------------------------
        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
        '1ST RUN DEBUG FOCUSED TO BUTTON BUT NOT HITT IT IN UNTIL 2ND
        'FOCUS PROBLEM FIRST
        '-------------------------------------------------
        For R_REPEAT = 1 To 3
            SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
    
    '---------------------------------------
    '31 AUG 2K SIXTEEN
    'LEARN
    '---------------------------------------
    'Find Window Tutorial Part1 <A must see> by Paul Zaczkowski (from psc cd)
    'https://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37467&lngWId=1
    '---------------------------------------
End If

OcWnd_GOOD_SYNC = i3



'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
' ICACLS -- Error Applying Security
'-------------------------------------------------------------------------------

Dim ICACLS_cWnd1 As Long, ICACLS_cWnd2 As Long, ICACLS_cWnd3 As Long, i3_2 As Long
ICACLS_SETTER_PERMISSION_HWND = FindWindow("#32770", "Error Applying Security")
i3_2 = ICACLS_SETTER_PERMISSION_HWND
If i3_2 >= 0 Then 'And i3_2 <> OcWnd_ICACLS_SETTER_PERMISSION_HWND Then
    'Sleep 2
    
    pwnd = i3_2
    'ICACLS_cWnd4 = 0
    'ICACLS_cWnd2 = FindWindowEx(pWnd, 0, "button", "&Continue")
    
'    TEXT1 = GetText(ICACLS_cWnd2)
'    ICACLS_cWnd2 = FindWindowEx(ICACLS_cWnd2, 0, "static", vbNullString) '"One or more jobs are running now:")
'    TEXT1 = GetText(ICACLS_cWnd2)
'    ICACLS_cWnd2 = FindWindowEx(ICACLS_cWnd2, 0, "static", vbNullString) '"One or more jobs are running now:")
'    TEXT1 = GetText(ICACLS_cWnd2)
'    ICACLS_cWnd2 = FindWindowEx(ICACLS_cWnd1, 0, "Static", vbNullString) '"One or more jobs are running now:")
'    TEXT1 = GetText(ICACLS_cWnd1)
'    ICACLS_cWnd2 = FindWindowEx(ICACLS_cWnd2, 0, "Static", vbNullString) '"One or more jobs are running now:")
'    TEXT1 = GetText(ICACLS_cWnd1)
    
    'If InStr(TEXT1, "One or more jobs are running now:") > 0 Then ICACLS_cWnd4 = 1
    
    'ICACLS_cWnd2 = FindWindowEx(pWnd, 0, "Button", "&OK")
    'ICACLS_cWnd3 = FindWindowEx(pWnd, 0, "button", "&Cancel")
    ICACLS_cWnd2 = FindWindowEx(pwnd, 0, "button", "&Continue")
    
    If ICACLS_cWnd2 > 0 Then
        
        '-------------------------------------------------
        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
        '1ST RUN DEBUG FOCUSED TO BUTTON BUT NOT HITT IT IN UNTIL 2ND
        'FOCUS PROBLEM FIRST
        '-------------------------------------------------
        For R_REPEAT = 1 To 2
            ICACLS_cWnd2 = FindWindowEx(pwnd, 0, "button", "&Continue")
            If ICACLS_cWnd2 > 0 Then
                
                i = SendMessage(ICACLS_cWnd2, BM_CLICK, 0, 0)
                'I = SendMessageany(ICACLS_cWnd2, BM_CLICK, 0, 0)
                'I = PostMessage(ICACLS_cWnd2, BM_CLICK, 0&, 0&)
                'I = SetForegroundWindow(ICACLS_SETTER_PERMISSION_HWND)
                'If I > 0 Then
                    'Sleep 100
                'End If
                
                If I_Memmer < Now Then
                    BEEP
                    I_Memmer = Now + TimeSerial(0, 0, 3)
                            
                End If
            DoEvents
            End If
            Next
'            ICACLS_cWnd2 = FindWindowEx(pwnd, 0, "button", "&Continue")
'            If ICACLS_cWnd2 > 0 Then
                'AppActivate "Error Applying Security", True
                'Sleep 3

            
                'cWnd = FindWindowEx(pwnd, 0, "button", "Cancel")
                'EnableWindow cWnd, False
                'EnableWindow ICACLS_cWnd2, False
                

                
            
'            End If

    End If
    
    '---------------------------------------
    '31 AUG 2K SIXTEEN
    'LEARN
    '---------------------------------------
    'Find Window Tutorial Part1 <A must see> by Paul Zaczkowski (from psc cd)
    'https://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37467&lngWId=1
    '---------------------------------------
End If

OcWnd_ICACLS_SETTER_PERMISSION_HWND = i3_2

End Sub

Private Sub Timer_COMMAND_RESTRICTOR_NOTEPAD_PICK_UP_Timer()

On Error Resume Next

If IsIDE = True Then Stop
'CHEK SOMETIME DO MORE DRIVE ACCESS QUICKER WANT LESS
'BEGIN DEBUG HERE


If MNU_COMMAND_RESTRICTOR_ENABLED.Caption <> "2 OF 2 _ COMMAND_RESTRICTOR_ENABLED_=_YES" Then
    TIMER_KILL_CMD_PREVIOUS_INSTANCE_SOFTLY_WHEN_TOO_MANY.Enabled = True
    MNU_COMMAND_RESTRICTOR_ENABLED.Caption = "2 OF 2 _ COMMAND_RESTRICTOR_ENABLED_=_YES"
End If

'If MNU_COMMAND_RESTRICTOR_ENABLED.Caption = "2 OF 2 _ COMMAND_RESTRICTOR_ENABLED_=_NOT" Then
'    MNU_COMMAND_RESTRICTOR_ENABLED.Caption = "2 OF 2 _ COMMAND_RESTRICTOR_ENABLED_=_NOT"
'End If

If Dir(FILENAME_PATH_COMMAND_RESTRICTOR) = "" Then
    If Err.Number > 0 Then
        BEEP
        Call MNU_COMMAND_RESTRICTOR_Click
        Exit Sub
    End If
    BEEP
    MsgBox "KAT_MP3_LOADER __ FILENAME WAIT RETURN MODIFIED __ HAD AN ERROR CREATING AND DISAPPEARED " + FILENAME_PATH_COMMAND_RESTRICTOR, vbMsgBoxSetForeground
    Timer_COMMAND_RESTRICTOR_NOTEPAD_PICK_UP.Enabled = False
    Exit Sub
End If

If Err.Number > 0 Then Exit Sub

Set F1 = FS.GetFile(FILENAME_PATH_COMMAND_RESTRICTOR)
If COMMAND_RESTRICTOR_NOTEPAD_PICK_UP_DATE <> F1.DateLastModified Then
    COMMAND_RESTRICTOR_NOTEPAD_PICK_UP_DATE = F1.DateLastModified
    If Err.Number > 0 Then Exit Sub
    
    Err.Clear
    FR1 = FreeFile
    Open FILENAME_PATH_COMMAND_RESTRICTOR For Input As #FR1
    COMMAND_RESTRICTOR_VALUE_______ = ""
    COMMAND_RESTRICTOR_CLASS_STRING = ""
    If Err.Number > 0 Then Exit Sub
    'ERR.DESCRIPTION
    Do
        If Err.Number > 0 Then Exit Sub
        
        Line Input #FR1, LINE_GET
        If Mid(LINE_GET, 1, 1) <> ";" Then
            DONE_ONE = False
            If COMMAND_RESTRICTOR_CLASS_STRING = "" Then
                COMMAND_RESTRICTOR_CLASS_STRING = LINE_GET
                DONE_ONE = True
            End If
            If COMMAND_RESTRICTOR_VALUE_______ = "" And DONE_ONE = False Then
                COMMAND_RESTRICTOR_VALUE_______ = Val(LINE_GET)
            End If
        End If
    Loop Until EOF(FR1)
    Close #FR1
End If

NOT_CORRECT = False
If COMMAND_RESTRICTOR_VALUE_______ = "" Then NOT_CORRECT = True
If COMMAND_RESTRICTOR_CLASS_STRING = "" Then NOT_CORRECT = True

If NOT_CORRECT = False And LOAD_TIME2 = False Then
    LOAD_TIME2 = True
End If

If NOT_CORRECT = True And LOAD_TIME2 = True Then
    LOAD_TIME2 = FLASE
    MsgBox "KAT_MP3_LOADER __ MUST BE SOMETHING WRONG TWO VALUE __ CLASS NAME AND COUNTER __ AS NOT PICKED UP IN NOTEPAD " + FILENAME_PATH_COMMAND_RESTRICTOR, vbMsgBoxSetForeground
End If

TIMER_KILL_CMD_PREVIOUS_INSTANCE_SOFTLY_WHEN_TOO_MANY.Enabled = True

Timer_COMMAND_RESTRICTOR_NOTEPAD_PICK_UP.Enabled = False

'If MNU_COMMAND_RESTRICTOR_ENABLED.Caption = "2 OF 2 _ COMMAND_RESTRICTOR_ENABLED_=_NOT" Then
'    MNU_COMMAND_RESTRICTOR_ENABLED.Caption = "2 OF 2 _ COMMAND_RESTRICTOR_ENABLED_=_NOT"
'End If

End Sub

Private Sub Timer_DESKTOP_INI_DEL_Timer()
    
    If UNLOAD_FORM_FLAG = True Then Unload Me: Exit Sub
    
    INTERVAL_Timer_DESKTOP_INI_DEL_VAR = Hour(Now)
    
    If INTERVAL_Timer_DESKTOP_INI_DEL_VAR = INTERVAL_Timer_DESKTOP_INI_DEL Then
        Exit Sub
    End If
    
    INTERVAL_Timer_DESKTOP_INI_DEL = INTERVAL_Timer_DESKTOP_INI_DEL_VAR
    
    Timer_DESKTOP_INI_DEL.Interval = 1000
    
    Call SPECIAL_FOLDER_DESKTOP_INI

End Sub


Sub OFFICE_TOOLBAR()


Exit Sub
If MNU_OFFICE_TOOLBAR_USE.Checked = False Then Exit Sub


'ONLY RUN IN LATEST USER NAME
GO3 = False
If GetComputerName = "1-ASUS-EEE" Then GO3 = True
If GetComputerName = "1-ASUS-X5DIJ" And GetUserName = "MATT 04" Then GO3 = True
If GetComputerName = "3-LINDA-PC" Then GO3 = True

    If GO3 = True Then

    'If I > 40 Then End
    'I = ShellExecute(Me.hWnd, "open", "C:\WINDOWS\explorer.exe", FILENAME, vbNullString, conSwNormal)
    ''Shell "EXPLORER.EXE """ + FILENAME + """", vbNormalFocus
    
    Dim OFFICE_FILENAME(3) As String
    'OFFICE XP TOOLBAR SHORTCUT IN WIN 7
    OFFICE_FILENAME(1) = GetSpecialfolder(38) + "\Microsoft Office\Office10\MSOFFICE.EXE"
    
    'OFFICE XP TOOLBAR SHORTCUT
    OFFICE_FILENAME(2) = GetSpecialfolder(38) + "\Microsoft Office\Office10\OSA.EXE"
    'OFFICE_FILENAME(2) = SHELL  GetSpecialfolder(38) + "\Microsoft Office\Office10\OSA.EXE -b -l", vbNormalNoFocus
    'EXTRA\ PARAMETERS AT END --  -b -l", vbNormalNoFocus
    
    'OFFICE 2000 TOOLBAR SHORTCUT
    OFFICE_FILENAME(3) = GetSpecialfolder(38) + "\Microsoft Office\Office\1033\MSOFFICE.EXE"
    
    For R = 1 To 3
        If Dir(OFFICE_FILENAME(R)) <> "" Then
            'PID VAR -- HAS TO BE -1 for ENTRY
            Var = cProcesses.GetEXEID(-1, OFFICE_FILENAME(R))
            If Var = False Then
                Shell OFFICE_FILENAME(R), vbNormalNoFocus
                WAIT_FOR_OFFICE_TOOLBAR_TO_LOAD = True
                Exit Sub
            End If
        End If
    Next
    
    
    

End If

End Sub

'Public Function GetActiveWindowTitle(ByVal ReturnParent As Boolean) As String
Public Function POSTMESSAGE_CLOSE_ALL_PARENT(ByVal TargetHwnd As Long) As String
    
    '--------------------------------------
    'NOT USED AT MOMENT
    'IS TO DO WITH INFORAPID SERACH REPLACE
    '--------------------------------------
    
    Dim i As Long
    Dim j As Long
    Dim COUNTER5
    i = TargetHwnd
    Dim ARRAY5(20) As Long
    If TargetHwnd Then
        Do While i <> 0
            j = i
            COUNTER5 = COUNTER5 + 1
            ARRAY5(COUNTER5) = i
            i = GetParent(i)
        Loop
        i = j
    
    End If
    
    RESULT = PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)
    
    For R = COUNTER5 To 1 Step -1
    'Debug.Print ARRAY5(R)
        If ARRAY5(R) <> 0 Then
            RESULT = PostMessage(ARRAY5(R), WM_CLOSE, 0&, 0&)
        End If
    Next
    For R = 1 To COUNTER5
    'Debug.Print ARRAY5(R)
        If ARRAY5(R) <> 0 Then
            RESULT = PostMessage(ARRAY5(R), WM_CLOSE, 0&, 0&)
        End If
    Next
End Function


Private Sub Timer_Find_Window_Timer()
    
    
    If UNLOAD_FORM_FLAG = True Then Unload Me: Exit Sub
    TIME_HAPPEN = Format(Now, "MM-DD DDD-DD-MMM -- HH:MM:SS")

    '-----------------------------------------------------------------------------------
    TargetHwnd = FindWindow(vbNullString, "CID Run Me.")
    If TargetHwnd = 0 And CID_RUN_ME_SHELL_LOADED = False Then
        CID_RUN_ME_SHELL_LOADED = True
        
        '------------------------------------
        'THE SWITCHES SET REMEMBER THIS VALUE
        'NOT TO LEAVE RUNNER ANYMORE
        '------------------------------------

        If MNU_CID_RERUN_ME_RELOADER.Checked = True Then
            If 1 = 1 Or Me.WindowState <> vbMinimized Then
                RUN_CODE = "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\Cid-RunMe.exe"
                If Dir(RUN_CODE) <> "" Then
                    Shell RUN_CODE, vbNormalNoFocus
                    BEEP
                End If
            End If
        End If
    End If
    
    'WAIT FOR BECOME ACTIVE
    'BUT ERROR NOT RESPOND WITH RELOAD
    TargetHwnd = FindWindow(vbNullString, "CID Run Me.")
    If TargetHwnd > 0 Then
        CID_RUN_ME_SHELL_LOADED = False
    End If
    'WORK AROUND
    CID_RUN_ME_SHELL_LOADED = False


    TargetHwnd = FindWindow(vbNullString, "InfoRapid Search & Replace")
    If TargetHwnd > 0 Then
        i = GetWindowRect(TargetHwnd, KAT_XY)
        XIOP = 0
        If KAT_XY.Right - KAT_XY.Left < 350 Then XIOP = 1
        If KAT_XY.Bottom - KAT_XY.Top < 170 Then XIOP = 1
        If XIOP = 1 Then
            'TargetHwnd1 = FindWindow(vbNullString, "InfoRapid Search & Replace - [Search Results]")
            'RESULT = PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)
            
            'LINE IN BETWEEN DASH ARE REM-ED OUT
            '----------------------------------------
            'POSTMESSAGE_CLOSE_ALL_PARENT (TargetHwnd)
            '----------------------------------------
            
            'RESULT = PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)
            'Parent
            
            'LINE IN BETWEEN DASH ARE REM-ED OUT
            '----------------------------------------
            'TargetHwnd = FindWindow(vbNullString, "InfoRapid Search & Replace - [Search Results]")
            '----------------------------------------
            'RESULT = PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)
            '----------------------------------------
        
            'RESULT
            'REQUEST TO CLOSE WINDOW IS MADE
            'WINDOW IS CLOSED
            'AND THEN PARENT WINDOW REQUEST MADE
            'AND THEN ORGINAL MESSGE BOX WINDOW IS BACK AGAIN
            'KILL IT AND MAYBE LOSE DATA ENTERED
            'OR SENDKEYS FOCUS
            'TEST DRIVE KILL
            'PROBLEM WITH KILL MIGHT NOT BE ONE WANT BEEN WORKING ON
'            PID = -1
'            Var = cProcesses.GetEXEID(PID, "C:\Program Files\WinMerge\WinMergeU.exe")
'            Var = cProcesses.Process_Kill(PID)
            
            i = SetForegroundWindow(TargetHwnd)
            SendKeys "N", True
            
        End If
    End If
    
    TargetHwnd = FindWindow(vbNullString, "Microsoft Office XP component")
    If TargetHwnd <> 0 Then
        RESULT = PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)
    End If
    
    'Dim ReturnValue, I
    'ReturnValue = Shell("calc.exe", 1)   ' Run Calculator.
    'AppActivate ReturnValue    ' Activate the Calculator.

    '-----------------------------------------
    'when set permission everyone full control
    '-----------------------------------------
    'test set the audit everyone
    'and use checkbox tick all sub inherit from this
    'not this subfolder only
    '-----------------------------------------------
    If GetForegroundWindow <> OTargetHwnd_AUDIT Then
        OTargetHwnd_AUDIT = 0
    End If

    TargetHwnd = FindWindow(vbNullString, "Error Applying Security")
    TargetHwnd = 0
    If TargetHwnd <> 0 And OTargetHwnd_AUDIT <> TargetHwnd Then
        i = SetForegroundWindow(TargetHwnd)
        DoEvents
        'SendKeys "{enter}", True
        SendKeys "{enter}", True
        OTargetHwnd_AUDIT = TargetHwnd
        DoEvents
        Sleep 500
    End If
    
    If FindWindow(vbNullString, "Microsoft Office Shortcut Bar") = 0 Then
        If WAIT_FOR_OFFICE_TOOLBAR_TO_LOAD = False Then
            Call OFFICE_TOOLBAR
        End If
    Else
        WAIT_FOR_OFFICE_TOOLBAR_TO_LOAD = False
    End If

    TargetHwnd = FindWindow(vbNullString, "Kat MP3 Recorder.EXE - Application Error")
    If TargetHwnd > 0 Then
    '    I = SetForegroundWindow(FindWindow(vbNullString, "Kat MP3 Recorder.EXE - Application Error"))
    '    PostMessage
    '    SendKeys "{ESC}", True
        
        COUNT_UP = 0
        FIRST_LIMIT_RESET_DONE = True
        'RESULT = PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)
        
        GO3 = 1
        Label1.Caption = TIME_HAPPEN + " -- Kat MP3 Recorder.EXE - Application Error #01"
        List1.AddItem Label1.Caption
        'Shell KAT_MP3_PATH, vbMinimizedNoFocus
        Call KILL_KAT_MP3_AND_THEN_SHELL
        
        
        Call Timer1_Timer
    End If
    
    'DISABLE THIS ONE CONFLICT BOX SETTINGS DEFAULT
    'CPU USAGE METHOD -- JUMP STARTER
    TargetHwnd = FindWindow("#32770", "About Kat MP3 Recorder")
    If TargetHwnd = 0 Then 'And About_Kat_MP3_Recorder = False Then
        About_Kat_MP3_Recorder = True
        'About_Kat_MP3_Recorder = TargetHwnd
        
        'THE About Kat MP3 Recorder
        'SHOW KAT MP3 NOT LOADED AND WANT RELOAD
        
        COUNT_UP = 0
        FIRST_LIMIT_RESET_DONE = True
        'RESULT = PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)
        
        GO3 = 1
        
        'Call KILL_KAT_MP3_AND_THEN_SHELL
        
        If KAT_MP3_IF_NOT_EXIST_AND_THEN_SHELL = True Then
            TIME_HAPPEN = Format(Now, "MM-DD DDD-DD-MMM -- HH:MM:SS")
            Label1.Caption = TIME_HAPPEN + " -- Kat MP3 Recorder.EXE - About Kat MP3 Recorder - Show Not Loaded"
            List1.AddItem Label1.Caption
        End If
        
        'Shell KAT_MP3_PATH, vbMinimizedNoFocus
        
        'Call Timer1_Timer
        'Call KAT_MP3_IF_NOT_EXIST_AND_THEN_SHELL_FUNC
        
        Call KILL_KAT_MP3_AND_THEN_SHELL
    
    End If
    'If TargetHwnd = 0 And About_Kat_MP3_Recorder = True Then About_Kat_MP3_Recorder = False

    
    'DISABLE THIS ONE CONFLICT BOX SETTINGS DEFAULT
    TargetHwnd = FindWindow("#32770", "Kat MP3 Recorder")
    If TargetHwnd > 0 And 1 = 2 Then
        
        'X -- 642 -- Norm Use Window
        'Y -- 422 -- Norm Use Window
        i = GetWindowRect(TargetHwnd, KAT_XY)
        XIOP = 0
        If KAT_XY.Right - KAT_XY.Left < 500 Then XIOP = 1
        If KAT_XY.Bottom - KAT_XY.Top < 400 Then XIOP = 1
        If XIOP = 1 Then
            COUNT_UP = 0
            FIRST_LIMIT_RESET_DONE = True
            'RESULT = PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)
            
            GO3 = 1
            Label1.Caption = TIME_HAPPEN + " -- Kat MP3 Recorder.EXE - Application Error #02"
            List1.AddItem Label1.Caption
            
            'Call KAT_MP3_IF_NOT_EXIST_AND_THEN_SHELL
            
            Call KILL_KAT_MP3_AND_THEN_SHELL

            'Call Timer1_Timer
        End If
    
    End If
    
    If XX = 0 Then
        XX = FindWindow("WinMergeWindowClassW", vbNullString)
        If XX > 0 Then
            Call Timer_WINMERGE_KILL_Timer
        End If
    End If

End Sub







'Hi World of Screen Gazers
'
'Thought I Would show off my Code Today
'for About Five Hour
'That I 'm Largely Proud Of
'Enjoy at Will
'----
'This Code Will Remove KatMouse When Not Using in VB6 IDE
'and Reload When You Want to Use in the VB6 IDE
'Because Exclude Limitation and then Complex Include Set Up
'----
'Without Scroll Wheel in Visual Studio for Basic 6 Virtually Impossible
'Other Way Visual Basic Code is Okay With Scroll Wheel When Running
'----
'Once I Thought to Use a Code the Same Myself
'as I Found Code to Harness the Scroll Bar
'----
'But The Code Had Limit Fault Where a Large Box the Scroll Bar Handling was Clipped to Not Reach the Bottom
'----
'Dreadful Wanker on My Computer When I Could Be Exercise Outdoor
'I Will Get Some Time Moment for That Pleasure Activity One Day
'----
'I Promise I was Going to Send a Blog of My Win XP Setup Batch Script Next or Soon
'But This Beat Me to It
'Well I still Got Win 7 and 64 Bit Win7 and Win 10 to Do
'----
'----
'Also, My Batch of Check Disk Program Will Gets an Edit
'as the Sound Player Program is an XP Program and Win 7 Don't Use
'but Answer is Use a Copy of the Sound Player From XP in Win 7
'Code Look at Bit Better - Visual Stunning
'than Batch Command Language
'----
'
'Hope YOU Like IT
'Hope You Humanoid Are All Well
'And Birds and Bee'
'Living Like Like Like a Cave Man in the Bitter Element
'It Still Cold on a Summer Night
'
'Welcome to Suggest Me Another Idea of Code
'I Got Enough Anyhow
'Just a Spontaneous Brain Wave Idea Today
'to End the Problem to do With Make Code Easy
'Kay Mouse is a Tool to Allow Scroll a Window
'Where Such Thing as Visual Studio v6 Doesn't Allow
'
'I Like the Trickiness of It With a Timer SUB and Many Latching including other Time Set Variable
'
'It Is Quite a Brave Thing of Me to Show How I Write Rem Statement
'Which are Usual Too Shy for Me to Show my Code
'and then they never see the Light of Day
'Except in Running Version's
'-----
'Top Mark Prize Award for Read All My Rem Statement
'-----
'Nice to Use a Rem and Code Spell Checker Like Grammarly
'-----
'and Upper Case WORD-IN Work
'on Spell Check in Own Web Page Based Environment
'-----
'======================================================
'
'======================================================
'KAT MP3 VB6 CODE
'------------------------------------------------------
'I FEEL I WANT TO SHOW OFF MY WORK TODAY
'TO PROVE MY WORTH
'------------------------------------------------------
'FEEDBACK BE NICE - GLAD TO HEAR FROM YOU
'-- ESPECIALLY IF WAS BE AUTHOR OF KATMOUSE
'IN MY LANGAUGE
'FEEDBACK TO AUTHOR ALSO ENCOURAGED
'WELCOME I HAVE NOT YET
'------------------------------------------------------
'
'======================================================
'MATTHEW LANCASTER - -ROIDSRIM
'MATT.LAN@ BTINTERNET.COM
'======================================================
'--------
'katmouse
'--------
'Lockergnome Award Software.
'Informer Editor 's pick award
'Mouse wheel enhancement for Windows (screen shot)
'written by Eduard Hiti.
'http://ehiti.de/katmouse/
'------------------------------------------------------
'de.web@ katmouse
'------------------------------------------------------
'Eduard Hiti
'Dipl.-Informatiker (Uni)
'------------------------------------------------------
'http://ehiti.de/katmouse/
'http://ehiti.de/
'------------------------------------------------------
'
'======================================================
'HERE IS MY CODE FROM TODAY
'THURSDAY 26-SIX MAY 2kSixteen
'------------------------------------------------------
'ABOUT 5 HOUR WRITING TO 7PM
'------------------------------------------------------
'======================================================
'I BEEN A LONG FAN OF KAT MOUSE
'AND FOUND TO USE WITH VB6
'BUT MANY YEAR AFTER FOUND IT
'HAD NEVER WORK PROPER WHEN LOSE FOCUS
'AND HAD A SIDE EFFECT RESULT
'MORE AS WORKED IN An OPPOSITE WAY WHEN SCROLL WAS SMOOTH
'NOT IN FOCUS BUT NOT WHEN IN
'------------------------------------------------------
'MANY YEAR LATER RECENTLY An NEW VERSION COME ALONG
'WITH IMPROVE AS TALK BEFORE
'FOR WINDOWS 7
'------------------------------------------------------
'TODAY I GOT IDEA
'AS SETUP PROCEDURE IS A LONG PROCESS FOR VB USER
'AS IT IS NOT COMPATIBLE WITH MANY PROGRAM
'SO I DECIDED ON IDEA TO POST MESSAGE KILL PROCESS NICELY
'WHEN NOT USE THE VB IDE PROGRAMMER
'DON 'T SEE MANY DRAW BACK TO THAT
'AND EXTRA IF PROBLEM WITH SCROLL BAR AND OTHER FLIP TO WINDOW PROBLEM
'IN WIN 7 OR WIN 10
'TO MAKE WORK COMPLETE AND QUICK SETUP ALL GOT TO DO IT USE THE CLASS
'DRAG A DROP BUTTON TO EXCLUDE THEM - THAT SHOULD WORK NICELY
'------------------------------------------
'
'------------------------------------------------------
'THE PROGRAM LOAD TIME OF KAT MOUSE WITH THIS CODE
'IS VERY NICE ALSO
'BUT MY CODE HAS OTHER TIMER AND MAY DRAG IT OUT A BIT
'AS IT IS - SMALL FOOTPRINT PROGRAM
'SET YOUR TIMER WHAT YOU WANT
'------------------------------------------------------
'THE CODE -- END RESULT IS TESTED
'WITH REM EXPLANATION
'------------------------------------------------------
'IT DON 'T USE PROCESS ENUMERATE TO KEEP PROCESS USAGE LOW
'------------------------------------------------------
'MAX SPEED IS TO USE QUICK FINDWINDOW
'------------------------------------------------------
'ONE LAST ERROR LOOKING AT
'IN EXE MODE SEEM TO LEAVE BEHIND An ICON AS LIKE
'NOT PROCESS CLOSE NICELY LIKE POST MESSAGE COMMAND TO CLOSE
'
'---------------------------------------------------
'ADDED TWO CODE AFTER FIRST PUBLISH HOUR AGO
'ONE MAKE A QUICK KILL PROCESS POST MESSAGE CLOSE
'WHEN TIMER SHOW KAT MOUSE BEEN IN USE BUT NEVER
'BRING UP WINDOW
'FOR QUICK CLOSE ON EXIT PROGRAM CODE IDE ENVIRONMENT
'AND ANOTHER TO MAKE SURE HWND EXIST BUT THE SECOND
'POST MESSAGE CLOSE COMMAND IN SUB
'---------------------------------------------------
'
'------------------------------------------------------
'NOT SURE IF CAN MAKE MORE COST CODE FRIENDLY
'NEAR THE END I ALREADY DUMPED TWO FAIR SIZE ROUTINE
'BYE - KEEP YOU SPIRIT HIGH
'======================================================
'
'======================================================
'THE DECLARE VARIABLE AND FUNCTION API SCRIPT
'======================================================
'
'Dim OXK3_KATMOUSE_PRESENCE_VISUAL_TIMER
'Dim MOUSE_LEFT_BUTTON_DELAY
'Dim OX33
'Dim O_HAS_VB_BEEN_VISIBLE_TOGGLE
'Dim OXK1
'Dim RESULT_KATMOUSE_NOT_BEEN_LOADED_BY_US_TIME
'
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
'
'Private Declare Function FindWindowDLL _
'        Lib "user32" _
'() '        Alias "FindWindowA" _
'        (ByVal lpClassName As Long, _
'        ByVal lpWindowName As Long) As Long
'Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'
'======================================================
'THE DECLARE VARIABLE AND FUNCTION API SCRIPT
'======================================================
'
'------------------------------------------------------
'ONE FUNCTION REQUIRED -- GetWindowTitle
'------------------------------------------------------
'Private Function GetWindowTitle(ByVal hwnd As Long) As String
'   Dim l As Long
'   Dim S As String
'   l = GetWindowTextLength(hwnd)
'   S = Space(l + 1)
'   GetWindowText hwnd, S, l + 1
'   GetWindowTitle = Left$(S, l)
'End Function
'------------------------------------------------------
'======================================================
'
'======================================================
'------------------------------------------------------
'ADD THIS CODE SUB ROUTINE INTO A TIMER
'------------------------------------------------------
'======================================================




Private Sub Timer_KAT_MOUSE_FOR_VB_Timer()
    Timer_KAT_MOUSE_FOR_VB.Enabled = False
Exit Sub
    
    
Dim CLASS_NAME_TEXT As String

'----------------------------------------------------
'PRIORITY IF FOCUS IS KATMOUSE
'THEN DONT CLOSE APP BECAUSE NOT MATCH WITH VB
'----------------------------------------------------
'----------------------------------------------------
'KAT MOUSE
'----------------------------------------------------
'IT ABILITY TO CLOSE WORK WELL ON POST MESSAGE COMMAND
'AND THEN REMOVE ITSELF AND ICON ALSO
'FORCE TERMINATE LEAVE ICON BEHIND
'PROCESS-KILL
'----------------------------------------------------

'----------------------------------------------------------
'IS VB WINDOW HERE
X1 = FindWindow("wndclass_desked_gsk", vbNullString)
'IF NOT CHECK IS KAT MOUSE HERE
XK1 = FindWindow("KatMouseWindowClass", vbNullString)
If XK1 > 0 Then
    OXK1_KATMOUSE_PRESENCE_TIMER = Now
Else
    OXK1_KATMOUSE_PRESENCE_TIMER = 0
End If

XK2 = FindWindow(vbNullString, "KatMouse Properties")
If XK2 > 0 Then
    XX3 = IsWindowVisible(XK2)
    If XK3 = 1 Then
        OXK3_KATMOUSE_PRESENCE_VISUAL_TIMER = Now
        'NEVER GOING TO USE RESET THIS LATER
        'TOP UP IDEA
    End If
    Else
    If OXK3_KATMOUSE_PRESENCE_VISUAL_TIMER = 0 Then
        OXK3_KATMOUSE_PRESENCE_VISUAL_TIMER = Now
    End If
End If

'------------------------------------------------------
'MAKE QUICKER INSTANT KIL PROCESS
'------------------------------------------------------
'GIVE A FEW SECOND TO LOAD BUT IF NEVER BEEN VISUAL AND
'LOAD A WHILE MAKE INSTANT KILL PROCESS
'------------------------------------------------------
i = DateDiff("S", OXK3_KATMOUSE_PRESENCE_VISUAL_TIMER, OXK1_KATMOUSE_PRESENCE_TIMER)
'SHOULD MAKE A QUICK KILL PROCESS HERE
'Debug.Print I

'IF VB NOT AND KAT MOUSE TRUE HERE
'AND NOT USED A MOMENT
'THEN DO A CLOSE PROCESS OF KAT MOUSE

'----------------------------------------------------------
'WANT A QUICK EXIT IF VB NOT HERE
'BUT ALLOW TO USE KAT MOUSE
'----------------------------------------------------------

'IS VB WINDOW HERE
'X1 = FindWindow("wndclass_desked_gsk", vbNullString)
'IF NOT CHECK IS KAT MOUSE HERE
'XK2 = FindWindow(vbNullString, "KatMouse Properties")

If i > 40 And (XK1 > 0 And X1 = 0) Then
    RESULT = PostMessage(XK1, WM_CLOSE, 0&, 0&)
    'Beep
    '------------------------------------------
    'DONT WAIT  CLOSED BEFORE FLAG HAS DONE
    '------------------------------------------
    RESULT_KATMOUSE_BEEN_LOADED = False
    RESULT_KATMOUSE_BEEN_LOADED_BY_US = FLASE
    
    'RESET HERE BECAUSE
    'MAYBE MISLEAD ABOUT BEEN IN USE WHEN LOADED
    OXK3_KATMOUSE_PRESENCE_VISUAL_TIMER = 0
    Exit Sub
End If


If X1 = 0 And XK1 > 0 Then
    If MOUSE_LEFT_BUTTON_DELAY = 0 Then
        MOUSE_LEFT_BUTTON_DELAY = Now + TimeSerial(0, 0, 3)
    End If
End If

JUMP:
'----------------------------------------------------------

XK1 = FindWindow("KatMouseWindowClass", vbNullString)
'XK2 = FindWindow(vbNullString, "KatMouse_MainWindow")
X3 = GetForegroundWindow()
X4 = GetWindowTitle(X3)

'-------------------------------------
'NOT GOOD ENOUGH
'THINK GOT HANDLE TEXT OF PROJECT NAME
'-------------------------------------
'FIND PARENT
'-----------
'X5 = GetParent(X3)
'NOT WORK ALREADY AT BASE

'----------------------------------------------------------------------
'KAT MOUSE
'---------
'SCREW IT USE A DELAY BEFORE CLOSE AND CONDICTION
'IS KATMOUSE WINDOW NOT FOREGROUND
'----------------------------------------------------------------------
'BUT WE WANT SWAP QUICK IN VB -- PROBLEM
'---------------------------------------
'OKAY SET FLAG IF JUST LOADED AND THEN WAIT AND WITH CONDITION
'-------------------------------------------------------------

'------------------------------------------------------------------
'PROBLEM SOLVED NEED ALSO CHECK APP.CAPTION NAME OF FORM WORKING ON
'----------
'DEBUG ONLY
'----------
'------------------------------------------------------------------
'RESULT GOOD
'-----------
If X4 = Me.Caption Then
    Exit Sub
End If


'---------------------------------------------------
'TIMER SAFEGUARD KAT MOUSE NOT CLOSED WHILE USING IT
'---------------------------------------------------
'IF BEEN USE IT AND NOT USE VB THEN DELAY CLOSE
'---------------------------------------------------
'AND IF LOADED TO USE AND NOT USE VB ALSO SAFEGUARD
'AND NOT LOADED BY HERE
'----------------------------------------------------
'GUESS ALWAYS WANT TIMER TOP UP WHEN KAT MOUSE AROUND
'JUST BEEN ARRIVED AROUND - MORE LIKE IT
'----------------------------------------------------
'TOP UP -- ONLY ONCE -- WHEN LOAD SHELL BEGAN
'TO LOAD TO USE
'AS LONG AS NOT LOADED BY US
'
'STEP -- METHOD
'LOAD IT
'SIT IN TASKBAR
'
'JUST BEEN LOADED
'TAG WITH SOME TIME DELAY

'------------------------------------------------
'ANOTHER DIMENSION
'------------------------------------------------
'IF IS LOAD BY US - NOT
'------------------------------------------------
'AND HAS BEEN
'ONLY JUST LOADED BY DETECT NEW HWND
'------------------------------------------------
If RESULT_KATMOUSE_BEEN_LOADED_BY_US = False Then
    If OXK1 <> XK1 And XK1 > 0 Then
        OXK1 = XK1
        RESULT_KATMOUSE_NOT_BEEN_LOADED_BY_US_TIME = Now
        Exit Sub
    End If
End If

'RECENTLY LOADED OR VISIBLE 01 OF 02
If FindWindow(vbNullString, "KatMouse Properties") > 0 Then
    XX52 = IsWindowVisible(FindWindow(vbNullString, "KatMouse Properties"))
    If XX52 = 1 Then
        RESULT_KATMOUSE_BEEN_VISIBLE_TIME = Now
        Exit Sub
    End If
End If

'RECENTLY LOADED OR VISIBLE 02 OF 02
'DO HAS KAT MOUSE ONLY JUST LOADED
'AND CHECK IF VB WORKING ON
'-----------------------
'NOT SURE REQUIRE BECAUSE
'IF VB WOKING ON -- KAT MOUSE IS OKAY TO RUN
'AND IF JUST LOADED ALSO
'-----------------------
'OTHER WORDS
'KAT MOUSE JUST LOADED WILL RUN FOR 40 SEC
'IF NOT ISWINDOWVISIBLE
'OR WITHOUT VB IN USE AND THEN CLOSE
'WHAT WOULD WANT QUICK



'------------------------------
'CHECK VB OPEN AND FROM THEN CLOSE
'WHILE KAT MOUSE TIMER
'------------------------------
'VB CLASS WINDOW IS HERE OR NOT
'------------------------------
X1 = FindWindow("wndclass_desked_gsk", vbNullString)
XX52 = IsWindowVisible(X1)
'---------------------------
'VB WINDOW IS NOT HERE
'AND IS A CHANGE FROM BEFORE
'---------------------------
If XX52 = 0 And XX52 <> O_HAS_VB_BEEN_VISIBLE_TOGGLE Then
    'REMOVE VALUE OF TIME IF VB
    RESULT_KATMOUSE_NOT_BEEN_LOADED_BY_US_TIME = 0
End If
O_HAS_VB_BEEN_VISIBLE_TOGGLE = XX52
'---------------------------------------
'LINE ABOVE REMOVE VALUE OF TIME IF VB
'NOT VISIBLE AFTER IT HAS HAD HAS BEEN
'FOR NEXT CODE BELOW
'MAKE IT QUICK --
'WHEN THE KAT MOUSE WAS NOT LOADED BY US
'---------------------------------------
'AND HERE IF FOREGROUND WINDOW CHANGES
'X3 = GetForegroundWindow()
'BUT WE DON'T WANT FOREGROUND CHANGE AND WILL RESET EVERYTHING
'WILL ONLY WANT IF CHANGE FROM VB TO ANOTHER FOREGROUND

'USUAL THIS DOES EQUAL A TIME EVEN LOAD BY US
'MAYBE WHEN DEBUG FIRST RUN
'BUT
'WORKING WELL AT THIS END POINT
'LOAD KAT MOUSE YOURSELF OKAY
'CLOSE IT OKAY
'AND MOVE TO ANOTHER FORM AND THEN FOREGROUND CHANGE WILL
'LET IT KILL PROCESS
'BUT DOES IT EXIT WHEN DELAY TIME WHEN CLOSE AND LEFT
'WITHOUT ANOTHER FOREGROUND CHANGE
'LONGER THAN 40 SEC AND STILL THERE
'DOES NOT MATTER FOREGROUND WOULD CHANGE ANYHOW
'IT WILL GIVE TIME TO GET TO VB AGAIN
'AND LEFT CLICK ON ICON IS PROBLEM
'MEANT TO DO RIGHT CLICK FOR PROGRAM
'WE USE MOUSE BUTTON DETECT - NOW

'MOUSE BUTTON LEFT
'GOOD THING AS ENABLE IS WITH LEFT BUTTON BUT
'TOO QUICK WITHOUT DELAY
'----------------------------------------------------
'----------------------------------------------------
'RESULT FROM THAT IS WHEN to GO TO ENABLE DISABLE
'----------------------------------------------------
'AND LEAVE FEW SECOND
'DEATH
'GUESS YOU DON'T WANT IT WHEN to GO TO DISABLE
'BUT AT LEAST SEE IT WORK
'PROBLEM IF YOU EVER WANT KAT MOUSE FOR ANYTHING ELSE
'RELOAD WILL BE DONE WHEN to WANT VB
'----------------------------------------------------
'ALSO DIALOG BOX LIKE FIND AND SEARCH ARE MAKE
'KAT MOUSE LOSE PROCESS
'NOT ENOUGH TIME IN DAY
'PROGRAM DIY THAT BIT IF YOU WISH
'OR WAIT FOR ME LATER
'IF I THINK IT WORTH IT
'A FIND OR SEARCH BOX YOU DON'T REQUIRE KAT MOUSE
'------------------------------------------------

'I = 0 ': KASC = 0
'For I = 1 To 255
'    BDF = GetAsyncKeyState(I)
'    'NOT F5
'    If I <> 116 Then
'        If BDF < -300 Then
'            KCODE = I
'            'Debug.Print I, BDF
'            Exit For
'        End If
'    End If
'    'I=27 ESCAPE BDF =-32758
'Next

'TIMER_KAT_MOUSE_FOR_VB.Interval


'--------------------------------------------
If GetAsyncKeyState(1) Then
    MOUSE_LEFT_BUTTON_DELAY = Now + TimeSerial(0, 0, 3)
    OX34 = 0
    Exit Sub
End If
If MOUSE_LEFT_BUTTON_DELAY > Now Then
    Exit Sub
End If
'------------------------------------------------------------
'WHILE BUTTON DELAY ANY FOREGROUND WINDOW CHANGE REMOVE VALUE
'------------------------------------------------------------
If MOUSE_LEFT_BUTTON_DELAY > Now Then
    'FOREGROUND WINDOW CHECK X3
    If OX34 > 0 And OX34 <> X3 Then
        MOUSE_LEFT_BUTTON_DELAY = 0
    End If
    OX34 = X3
End If

'FOREGROUND WINDOW CHECK X3 WITH CHECK OF VB
If (X1 > 0 And OX33 = X1) And OX33 <> X3 Then
    RESULT_KATMOUSE_NOT_BEEN_LOADED_BY_US_TIME = 0
End If
OX33 = X3

XXTIME = TimeSerial(0, 0, 40)
If RESULT_KATMOUSE_NOT_BEEN_LOADED_BY_US_TIME + XXTIME > Now Then
    Exit Sub
End If
'---------------------------------------

'----------------------------------------------------------------
'NEXT STEP FOUND TO BE LOADED AND WITH TIME DELAY BEEN SET BEFORE
'----------------------------------------------------------------
XX22 = RESULT_KATMOUSE_BEEN_LOADED_TIME + TimeSerial(0, 0, 40)
If XX22 > Now Then
    '-------------------------------------------------------
    'EXTRA TOP UP ADDED ON TIME VAR
    'IF KAT MOUSE WINDOW STILL LOADED CONDITION
    'AND MORE IMPORTANT IS VISIBLE
    '-------------------------------------------------------
    'WE DO TO FIND IT IS VISIBLE
    '--------------------------------------------
    '---------------------------------------
    'KatMouse Properties
    'WON'T SHOW HWND UNLESS A VISIBLE WINDOW
    '---------------------------------------
    If XK1 > 0 Then
        If FindWindow(vbNullString, "KatMouse Properties") > 0 Then
            XX52 = IsWindowVisible(FindWindow(vbNullString, "KatMouse Properties"))
            If IsWindow(XX52) = 1 Then
                RESULT_KATMOUSE_BEEN_VISIBLE_TIME = Now
            End If
            RESULT_KATMOUSE_BEEN_LOADED_TIME = Now
            Exit Sub
        End If
    End If
End If

'---------------------------------------------
'DO FOREGROUND MATCH KAT MOUSE WINDOW 02 OF 02
If X3 = XK1 Then
    Exit Sub
End If
'----------
'SECURED IT
'UNDERSTAND FINDWINDOW VS GetWindowTitle -- STRING
'NOT WORK
'----------


'------------------------------
'VB CLASS WINDOW IS HERE OR NOT
'------------------------------
X1 = FindWindow("wndclass_desked_gsk", vbNullString)
'----------------------------------------------------
'BUT ONLY WHEN IN ISIDE = TRUE
'----------------------------------------------------
'OR MABE ANOTHER PROGRAM - GUESS - LATERS TO THAT ONE
'----------------------------------------------------

'---------------------------------------------------
'IS THERE An FORE_GROUND MATCH WITH VB PROGRAM WINDOW
'---------------------------------------------------
'X3 = GetForegroundWindow()
'---------------------------------------------------
'NOT REQUIRE TWICE AS ABOVE
'------------------------------------------------
'FIND THE WINDOW PRIZE TREASURE CHEST BUST BLOSSOM
'TRAIN SPOTTER
'------------------------------------------------
If X3 = X1 Then
    
    '-----------------------------------------------------------
    'CHECK CLASS NAME IS VB IDE WINDOW WE WANT OR IF ANOTHER NOT
    '-----------------------------------------------------------
    X2 = GetWindowTitle(X1)
    If InStr(X2, " - Microsoft Visual Basic [") > 0 Then
        
        '-----------------------------------------------
        'YES IN VB
        'AND THEN CHECK IF KAT MOUSE LOADED THAT WE WANT
        '-----------------------------------------------
        
        'KAT MOUSE
        'NOT HERE AND THEN WE GET IT
        If XK1 = 0 Then
        
            If Dir("C:\Program Files (x86)\KatMouse\KatMouse.exe") <> "" Then
                Shell "C:\Program Files (x86)\KatMouse\KatMouse.exe", vbNormalNoFocus ', vbNormalNoFocus
                RESULT_KATMOUSE_BEEN_LOADED = True
            End If
            If RESULT_KATMOUSE_BEEN_LOADED = False Then
                If Dir("C:\Program Files\KatMouse\KatMouse.exe") <> "" Then
                    Shell "C:\Program Files\KatMouse\KatMouse.exe", vbNormalNoFocus
                    RESULT_KATMOUSE_BEEN_LOADED = True
                    RESULT_KATMOUSE_BEEN_LOADED_BY_US = True
                End If
            End If
                
            '---------------------------------------------------
            'STORE VAR TOGGLE
            '---------------------------------------------------
            If RESULT_KATMOUSE_BEEN_LOADED = True Then
                
                '--------------------------
                'BETTER WAIT FOR IT TO LOAD
                '--------------------------
                XCOUNT_ERROR = 100
                Do
                    XCOUNT_ERROR = XCOUNT_ERROR - 1
                    DoEvents
                    X5 = FindWindow("KatMouseWindowClass", vbNullString)
                    If XCOUNT_ERROR = 0 Or X5 > 0 Then Exit Sub
                    Sleep 50
                Loop Until X5 > 0
                '--------------------------
                
            End If
            
            If FindWindow("KatMouseWindowClass", vbNullString) > 0 Then
                
                '----------------------------------------------------
                'WHEN FIRST TIME RUN MAKE NOTE MARK FLAG -- IF LOADED
                'WHEN WE WANTED IT
                '----------------------------------------------------
                
                RESULT_KATMOUSE_BEEN_LOADED = True
            
                '------------------------
                'QUIT HERE DON'T CONTINUE
                '------------------------
                Exit Sub
            End If
        End If
    
        '----------------------------------------------------
        'GONE THROUGH THAT BLOCK OF STATEMENT AND GOT TO HERE
        'AND THEN KAT MOUSE IS LOADED SET FLAG
        '----------------------------------------------------
        If FindWindow("KatMouseWindowClass", vbNullString) > 0 Then
            RESULT_KATMOUSE_BEEN_LOADED = True
        End If
            
        '---------------------------------------------
        'RESULT SET ABOVE MORE ALSO SO DOUBLE IF CHECK
        '---------------------------------------------
        If RESULT_KATMOUSE_BEEN_LOADED = True Then
            RESULT_KATMOUSE_BEEN_LOADED_TIME = Now
            Exit Sub
        End If
    End If
End If
    
'--------------------------------------
'FOREGROUND WINDOW DONT MATCH VB WINDOW
'FLIP IT OFF
'--------------------------------------

'--------------------------------
'X7 = GetWindowTitle(X1)
'X8 = GetWindowTitle(X3)
'X9 = GetWindowTitle(X3)
'cText = Space$(255)
'I = GetClassName(X3, cText, 255)
'X10 = cText
'X11 = GetParent(X3)
'--------------------------------

If RESULT_KATMOUSE_BEEN_LOADED_BY_US = False Then
    
    'X3 FOREGROUND
    'X1 VB WINDOW
    
    If X3 <> X1 Then
        '---------------------------------------------------
        'STORE VAR TOGGLE
        '---------------------------------------------------
        
        '-------------------------------------------------------------
        'CLOSE PROGRAM - KAT MOUSE WHEN DONT WANT IT
        'PROBLEM UNABLE TO EXCLUDE ALL AP PROGRAM AND INCLUDE ONLY ONE
        'FOR VB
        '-------------------------------------------------------------
        'ANOTHER INTEREST-ING THING FOR WIN 10 AND WIN 7 USER OF KAT MOUSE
        '----
        'YOU GET An HORIZONTAL SCROLL BAR MOVEMENT WHEN SCROLL DOWN
        'NOT WANTED
        'SO EXCLUDE CLASS HORIZONTAL SCROLL BAR IN KAT MOUSE
        'FIX-ED-NESS
        '----------------------------------------------------------------
        
        X5 = FindWindow("KatMouseWindowClass", vbNullString)
        If X5 = 0 Then Exit Sub
        '----------------
        'CLOSE THE WINDOW
        '----------------
        RESULT = PostMessage(X5, WM_CLOSE, 0&, 0&)
        '----------------
        '----------------
        '----------------
        '----------------
        
        '--------------------------------
        'BETTER WAIT FOR IT TO CLOSE EXIT
        '--------------------------------
        XCOUNT_ERROR = 100
        Do
            XCOUNT_ERROR = XCOUNT_ERROR - 1
            DoEvents
            X5 = FindWindow("KatMouseWindowClass", vbNullString)
            If XCOUNT_ERROR = 0 Or X5 = 0 Then Exit Sub
            Sleep 50
        Loop Until X5 = 0
        
        '------------------------------------------
        'BE SURE HAS CLOSED BEFORE FLAG HAS DONE
        '------------------------------------------
        If X5 = 0 Then
            RESULT_KATMOUSE_BEEN_LOADED = False
            RESULT_KATMOUSE_BEEN_LOADED_BY_US = FLASE
        End If
        
    End If
End If
    
End Sub

Private Sub TIMER_MNU_CID_RELOADER_DISAPPEAR_Timer()

MNU_CID_RELOADER.Visible = False
TIMER_MNU_CID_RELOADER_DISAPPEAR.Enabled = False

End Sub

Private Sub Timer_Remove_Explorer_Libraries_Timer()
'------------------------------
'CALLED FROM HERE --- BELOW ONE
'Timer_EXPLORER_RESET_Timer
'------------------------------

ONE_INSTANCE_ONLY = False
i = FindWinPart_VISIBLE("Libraries", "CabinetWClass", HWND_RETURN, VAR1, KILL_SOFTLY, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)

If HOW_MANY_INSTANCE_SHOWING = 1 Then
    Timer_Remove_Explorer_Libraries.Enabled = False
End If



TargetHwnd = i
RESULT = PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)
End Sub


Sub TIMER_KILL_CMD_PREVIOUS_INSTANCE_SOFTLY_WHEN_TOO_MANY_TIMER()

'--------------------------------------------------------
'USE WITHER HERE
'Sub Timer_EXPLORER_RESET_Timer
'--------------------------------------------------------
If MNU_COMMAND_RESTRICTOR_ENABLED.Caption <> "2 OF 2 _ COMMAND_RESTRICTOR_ENABLED_=_YES" Then
    MNU_COMMAND_RESTRICTOR_ENABLED.Caption = "2 OF 2 _ COMMAND_RESTRICTOR_ENABLED_=_YES"
End If

Dim hWndForm        As Long
Dim hWndTextBox     As Long
Dim Retval As Long
Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
Dim WinClass As String, WinTitle2 As String

'---------------------------------
HOW_MANY_INSTANCE_SHOWING = -1
'---------------------------------
ONE_INSTANCE_ONLY = False

KILL_SOFTLY = COMMAND_RESTRICTOR_VALUE_______
CLASS_NAME_KILL_SOFTLY = COMMAND_RESTRICTOR_CLASS_STRING

If KILL_SOFTLY = "" And CLASS_NAME_KILL_SOFTLY = "" Then Exit Sub

i = FindWinPart_VISIBLE("", CLASS_NAME_KILL_SOFTLY, HWND_RETURN, VAR1, KILL_SOFTLY, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)

If i > 0 Then EXPLORER_DETECT_IS_PRESENT = True
If HOW_MANY_INSTANCE_SHOWING > 4 Then

    '---------------------------------
    'SET ENTRY TO KILL WHAT NOT WANTER
    '---------------------------------
    HOW_MANY_INSTANCE_SHOWING = 4
    i = FindWinPart_VISIBLE("", "ConsoleWindowClass", HWND_RETURN, VAR1, KILL_SOFTLY, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)

End If

End Sub




Sub TIMER_EXPLORER_IS_PRESENT_OPEN_AT_FOLDER_TRUE_TIMER()

'--------------------------------------------------------
'USE WITHER HERE
'Sub Timer_EXPLORER_RESET_Timer
'--------------------------------------------------------

Dim hWndForm        As Long
Dim hWndTextBox     As Long
Dim Retval As Long
Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
Dim WinClass As String, WinTitle2 As String

ONE_INSTANCE_ONLY = False
EXPLORER_DETECT_IS_PRESENT = False
i = FindWinPart_VISIBLE("", "CabinetWClass", HWND_RETURN, VAR1, KILL_SOFTLY, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)
If i > 0 Then EXPLORER_DETECT_IS_PRESENT = True
i = FindWinPart_VISIBLE("", "ToolbarWindow32", HWND_RETURN, VAR1, KILL_SOFTLY, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)
If i > 0 Then EXPLORER_DETECT_IS_PRESENT = True

hWndFormE = FindWindow("CabinetWClass", vbNullString)

If hWndFormE > 0 Then
    hWndIEChild = FindWindowEx(hWndFormE, 0&, vbNullString, vbNullString)
    hWndIEChild1 = FindWindowEx(1, 0&, vbNullString, vbNullString)
    hWndIEChild = FindWindowEx(hWndFormE, 0&, "ShellTabWindowClass", vbNullString)
    hWndIEChild = FindWindowEx(hWndFormE, 0&, "WorkerW", vbNullString)
    hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ReBarWindow32", vbNullString)
    hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ComboBoxEx32", vbNullString)
    hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ComboBox", vbNullString)
    hWndIEChild = FindWindowEx(hWndIEChild, 0&, "Edit", vbNullString)

    window_hwnd = hWndIEChild
    txtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, 0, 0)
        
    txtlen = SendMessage(window_hwnd, WM_GETTEXT, 255, ByVal WinTitleBuf)
    WinTitle = StripNulls(WinTitleBuf)
    
    WinTitle2 = GetWindowTitle(hWndForm)
    WinTitle2 = GetWindowTitle(hWndIEChild)
End If

End Sub

Public Function StripNulls(OriginalStr As String) As String
        ' This removes the extra Nulls so String comparisons will work
        If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
        End If
        StripNulls = OriginalStr
End Function

'Function GetWindowTitle(ByVal hwnd As Long) As String
'   Dim l As Long
'   Dim S As String
'   l = GetWindowTextLength(hwnd)
'   S = Space(l + 1)
'   GetWindowText hwnd, S, l + 1
'   GetWindowTitle = Left$(S, l)
'End Function

Private Sub Timer_EXPLORER_RESET_Timer()

Exit Sub

'If EXPLORER_RESET_TIMER = 0 Then
'    EXPLORER_RESET_TIMER = Now + TimeSerial(0, 15, 0)
'End If
'If EXPLORER_RESET_TIMER > Now Then Exit Sub

'1 HOUR
If MouseDetectMove_FOR_EXPLORER_RESET = 0 Then Exit Sub
If MouseDetectMove_FOR_EXPLORER_RESET > Now Then Exit Sub

If HOW_MANY_INSTANCE_SHOWING = 1 Then
    Timer_Remove_Explorer_Libraries.Enabled = False
End If

If 1 = 1 Or Dir("C:\BAT\EXPLORER_RESET.BAT") = "" Then
    
    i = ""
    i = i + "TASKKILL /F /IM EXPLORER.EXE /T" + vbCrLf
    i = i + "EXPLORER.EXE" + vbCrLf
    i = i + "EXIT" + vbCrLf
    
    '-------------------------------------------------
    'HOW DO THIS ONE
    '-------------------------------------------------
    'I = I + "PAUSE" + vbCrLf
    'I = I + "TASKKILL /IF Libraries" + vbCrLf
    '-------------------------------------------------
    
    FR1 = FreeFile
    Open "C:\BAT\EXPLORER_RESET.BAT" For Output As #FR1
    Print #FR1, i
    Close #FR1
End If

'------------------------------
'/B RUN WITHOUT OPEN NEW WINDOW
'GET EXIT TO WORK PROPER
'------------------------------
'/C DONT WAIT TERMINATE AFTER
'/K OPPOSITE
'------------------------------

Shell "CMD /C START """", /B /REALTIME C:\BAT\EXPLORER_RESET.BAT", vbHide
Timer_Remove_Explorer_Libraries.Enabled = True
End Sub

Sub TIMER_SET_CAMERA_CHECK_TIMER()

XDrive1_ListCount = Drive1.ListCount
If XDrive1_ListCount = O_DRIVE1_ListCount Then Exit Sub

If XDrive1_ListCount <> O_DRIVE1_ListCount Then
    If VAR_MC__HX60V_SET = False Then
        O_DRIVE1_ListCount = XDrive1_ListCount
    End If
End If

If Dir("MC - HX60V", vbVolume) = "" Then
    TIMER_SET_CAMERA_CHECK.Interval = 10000
    If VAR_MC__HX60V_SET = True Then
        VAR_MC__HX60V_SET = False
    End If
End If

If Dir("MC - HX60V", vbVolume) <> "" Then
    If VAR_MC__HX60V_SET = False Then
        VAR_MC__HX60V_SET = True
        TIMER_SET_CAMERA_WITH_SECURITY_INFO.Enabled = True
        TIMER_SET_CAMERA_CHECK.Interval = 4000
    End If
End If

End Sub



Sub TIMER_SET_CAMERA_WITH_SECURITY_INFO_TIMER()
'Exit Sub ' here ----
PASS_DRIVE_AT = ""
TO_GO = False
For R = 69 To 89
    On Error Resume Next
    If Dir(Chr(R) + ":\DCIM", vbDirectory) <> "" Then TO_GO = True
    If Err.Number > 0 Then
        TIMER_SET_CAMERA_WITH_SECURITY_INFO.Enabled = False
        Exit Sub
    End If
    On Error GoTo 0
    If Dir(Chr(R) + ":\MP_ROOT", vbDirectory) <> "" Then TO_GO = True
    If Dir(Chr(R) + ":\", vbDirectory) <> "" Then
        If Dir(Chr(R) + ":\", vbVolume) = "MC - HX60V" Then
            TO_GO = True
        End If
    End If


    If TO_GO = True Then
        PASS_DRIVE_AT = Chr(R)
        '-----------------------------------
        Call COPY_MY_SECURITY_INFO_TO_CAMERA(PASS_DRIVE_AT)
        '-----------------------------------
        Exit For
            
    End If
Next

TIMER_SET_CAMERA_WITH_SECURITY_INFO.Enabled = False
Exit Sub

'--------------------------------------------------------------
'WAS GOING TO CHECK VOLUME NAME EXIST WITHOUT CHECK FOLDER PAIR
'BUT TAKEN CARE OF ABOVE
'ANOTHER WAY OF WORK AROUND
'--------------------------------------------------------------
'ENDER CODE Tue 04 April 2017 10:39:48----------
'--------------------------------------------------------------


If 1 = 2 And PASS_DRIVE_AT = "" Then
    If Dir("MC - HX60V", vbVolume) <> "" Then
        TO_GO = False
        For R = 69 To 89
            If Dir(Chr(R) + ":\", vbDirectory) <> "" Then
                If Dir(Chr(R) + ":\", vbVolume) = "MC - HX60V" Then
                    TO_GO = True
                End If
            End If
            
            
            If TO_GO = True Then
                PASS_DRIVE_AT = Chr(R)
                '-----------------------------------
                Call COPY_MY_SECURITY_INFO_TO_CAMERA(PASS_DRIVE_AT)
                '-----------------------------------
                Exit For
            End If
        Next
    End If
End If


TIMER_SET_CAMERA_WITH_SECURITY_INFO.Enabled = False

End Sub


Private Sub TIMER_SET_CAMERA_UP_TO_MARK_FOLDER_Timer()

If Timer_VICE_VERSA_TIMER_ON.Enabled = False Then Exit Sub

On Error Resume Next
Dir2.Path = "W:\DCIM"

If Err.Number > 0 Then
    ERR_COUNT1 = ERR_COUNT1 + 1
    If ERR_COUNT1 > 500 Then
        TIMER_SET_CAMERA_UP_TO_MARK_FOLDER.Enabled = False
    End If
    If Err.Number > 0 Then
        Exit Sub
    End If
End If
On Error GoTo 0


TIMER_SET_CAMERA_UP_TO_MARK_FOLDER.Enabled = False

Call LAST_FOLDER_CAMERA("W:\DCIM")
Call LAST_FOLDER_CAMERA("W:\MP_ROOT")

'-----------------------------------
PASS_DRIVE_AT = "W"
Call COPY_MY_SECURITY_INFO_TO_CAMERA(PASS_DRIVE_AT)
'-----------------------------------

'WRITE_INFO_TXT_IN_EMPTY_FOLDER
Call WRITE_INFO_TXT_IN_EMPTY_FOLDER("W:\DCIM")
Call WRITE_INFO_TXT_IN_EMPTY_FOLDER("W:\MP_ROOT")

End Sub

Sub COPY_MY_SECURITY_INFO_TO_CAMERA(PASS_DRIVE_AT)
    
    On Error Resume Next
    i = App.Path + "\CAMERA DATA"

    File2.Path = i
    If Err.Number > 0 Then Exit Sub
    File2.FileName = "*.*"
    
    For R1 = 0 To File2.ListCount - 1
        
        FN1 = File2.Path + "\" + File2.List(R1)
        If FS.FileExists(FN1) = True Then
            Set F1 = FS.GetFile(FN1)
        End If
        
        'FN2 = "W:\" + File2.List(R1)
        FN2 = PASS_DRIVE_AT + ":\" + File2.List(R1)
        If FS.FileExists(FN2) = True Then
            Set F2 = FS.GetFile(FN2)
        End If

        NOT_COPY = False
        If FS.FileExists(FN1) = True Then
            If FS.FileExists(FN2) = True Then
                If F1.DateLastModified = F2.DateLastModified Then
                    NOT_COPY = True
                End If
            End If
        End If
        
        If NOT_COPY = False Then
            FS.CopyFile FN1, FN2
        End If
    Next

End Sub

Function LATEST_DATE_FILE_ON_CAMERA_CHANGED()

If Timer_VICE_VERSA_TIMER_ON.Enabled = False Then Exit Function


'GET DATE FOR CHANGES IF VALUE IS NONE
If OPHOTO_DATE_VAR = 0 Then
    
    FOLDER_FILE_NAME = App.Path + "\DATE LAST FILE CHANGE ON CAMERA.TXT"
    If Dir(FOLDER_FILE_NAME) <> "" Then
        FR1 = FreeFile
        Open FOLDER_FILE_NAME For Input As #FR1
            Line Input #FR1, LINE_VAR_INFO_NOW
        Close #FR1
        OPHOTO_DATE_VAR = DateValue(LINE_VAR_INFO_NOW) + TimeValue(LINE_VAR_INFO_NOW)
    End If
End If

For RSEL = 1 To 2
    If RSEL = 1 Then FOLDER = "DCIM"
    If RSEL = 2 Then FOLDER = "MP_ROOT"
    
'    If InStr(FOLDER, "DCIM") > 0 Then
'        VAR_SIZE_LEN = 16
'    End If
'    If InStr(FOLDER, "MP_ROOT") > 0 Then
'        VAR_SIZE_LEN = 19
'    End If
    
    On Error Resume Next
    Dir2.Path = FOLDER
    
    For R = 0 To Dir2.ListCount - 1
        'Debug.Print Dir2.List(R)
        FLAG_UP_FILE_EXIST = -2
        File2.Path = Dir2.List(R)
        For R2 = 0 To File2.ListCount - 1
            Set F = FS.GetFile(File2.Path + "\" + File2.List(R2))
            PHOTO_DATE = F.DateLastModified
            If PHOTO_DATE > PHOTO_DATE_VAR Then
              PHOTO_DATE_VAR = PHOTO_DATE
            End If
        Next
    Next
Next
    
'STORE THE CHANGED DATE IF
If OPHOTO_DATE_VAR <> PHOTO_DATE_VAR Then
    
    FOLDER_NAME = Dir2.List(R)
    FOLDER_FILE_NAME = App.Path + "\DATE LAST FILE CHANGE ON CAMERA.TXT"
    
    FR1 = FreeFile
    Open FOLDER_FILE_NAME For Output As #FR1
        Print #FR1, Now
    Close #FR1
    
    LATEST_DATE_FILE_ON_CAMERA_CHANGED = True
    
End If
    
OPHOTO_DATE_VAR = PHOTO_DATE_VAR

End Function



Sub WRITE_INFO_TXT_IN_EMPTY_FOLDER(FOLDER)

If InStr(FOLDER, "DCIM") > 0 Then
    VAR_SIZE_LEN = 16
End If
If InStr(FOLDER, "MP_ROOT") > 0 Then
    VAR_SIZE_LEN = 19
End If

On Error Resume Next
Dir2.Path = FOLDER

For R = 0 To Dir2.ListCount - 1
    'Debug.Print Dir2.List(R)
    
    FLAG_UP_FILE_EXIST = -2
    
    File2.Path = Dir2.List(R)
    For R2 = 0 To File2.ListCount - 1
        If Len(Dir2.List(R)) > VAR_SIZE_LEN Then
            If FLAG_UP_FILE_EXIST = -2 Then
                FLAG_UP_FILE_EXIST = False
            End If
            If UCase(File2.List(R2)) = "INFO ----------.TXT" Then
                FLAG_UP_FILE_EXIST = True
            End If
        End If
    Next
        
    If File2.ListCount = 0 And Len(Dir2.List(R)) > VAR_SIZE_LEN Then
        FLAG_UP_FILE_EXIST = False
    End If
        
    If FLAG_UP_FILE_EXIST = False Then
    
        FOLDER_NAME = Dir2.List(R)
        FR1 = FreeFile
        Open FOLDER_NAME + "\INFO ----------.TXT" For Output As #FR1
            Print #FR1, "MUST MAKE FOLDER NOT EMPTY VICE VERSA"
            Print #FR1, FOLDER_NAME
            Print #FR1, Now
        Close #FR1
    
    End If
Next


End Sub

Sub LAST_FOLDER_CAMERA(FOLDER)

On Error Resume Next

If InStr(FOLDER, "DCIM") > 0 Then
    VAR_SIZE_LEN = 16
End If
If InStr(FOLDER, "MP_ROOT") > 0 Then
    VAR_SIZE_LEN = 19
End If

Dir2.Path = FOLDER

LAST_FOLDER = ""
LAST_FOLDER_MARK = True
For R = Dir2.ListCount - 1 To 0 Step -1
    'Debug.Print Dir2.List(R)
    File2.Path = Dir2.List(R)
    If File2.ListCount > 0 Then
        If Len(Dir2.List(R)) = VAR_SIZE_LEN Then
            LAST_FOLDER = Dir2.List(R)
            Exit For
        End If
    End If
    
    If Len(Dir2.List(R)) > VAR_SIZE_LEN Then
        LAST_FOLDER_MARK = False
    End If
Next

On Error Resume Next
    
If LAST_FOLDER_MARK = True Then
    FOLDER_NAME = LAST_FOLDER + " -- GOT TO HERE -- " + Format(Now, "YYYY-MM-DD -- HH-MM-SS")
    MkDir FOLDER_NAME
    If Err.Number > 0 Then
        MsgBox "ERROR MAKE MARKER FOLDER FOR CAMERA" + vbCrLf + FOLDER_NAME, vbMsgBoxSetForeground
        Exit Sub
    End If
    FR1 = FreeFile
    Open FOLDER_NAME + "\INFO ----------.TXT" For Output As #FR1
        Print #FR1, "MUST MAKE FOLDER NOT EMPTY VICE VERSA"
        Print #FR1, Now
    Close #FR1
    If Err.Number > 0 Then
        MsgBox "ERROR MAKE FILE IN FOLDER FOR CAMERA" + vbCrLf + FOLDER_NAME + "\INFO ----------.TXT", vbMsgBoxSetForeground
        Exit Sub
    End If
End If


'DELETE PREVIOUS MARKER ADD NEW
'IF NOTHING ON LAST ONE
For R = Dir2.ListCount - 2 To 0 Step -1
    'Debug.Print Dir2.List(R)
    If Len(Dir2.List(R)) > VAR_SIZE_LEN Then
        RmDir Dir2.List(R)
    End If
    If Len(Dir2.List(R)) = VAR_SIZE_LEN Then
        Exit For
    End If
Next



End Sub

Private Sub Timer_Tidal_Information_Microsoft_Visual_Basic_break_SubRoutine_Timer()

'-------------------------------------------------------------------------------------
'TEMP REMER OUT BECUASE CONFLICT MISSING CHUNNK ANOTHER ROUTINE WHEN MERGE CAME SYNCER
'-------------------------------------------------------------------------------------

Timer_Tidal_Information_Microsoft_Visual_Basic_break_SubRoutine.Enabled = False
Exit Sub


i = Tidal_Information_Microsoft_Visual_Basic_break_Counter
i = i + 1
Tidal_Information_Microsoft_Visual_Basic_break_Counter = i
If i > 100 Then
    Timer_Tidal_Information_Microsoft_Visual_Basic_break_SubRoutine.Enabled = False
End If
BEEP


End Sub

Private Sub Timer_TIME_HAPPEN_Timer()
TIME_HAPPEN = Format(Now, "MM-DD DDD-DD-MMM -- HH:MM:SS")
End Sub

Private Sub Timer_VB_MAXIMIZE_Timer()

'PURELY FOR VB ROUTINE IF ANOTHER ARE HERE

' ----------------------------------
' ALTERNATIVE LIKE HERE
' ----------------------------------
' Timer_01_VB_HELP_BOX_02_MSDN_Timer
' ---- OTHER -----------------------
' Timer_VB_MAXIMIZE
' ---- ALL THE VB ------------------
' ----------------------------------

'01 OF 02 DISABLE THE HELP BUTTON ACCIDENT PRESS
'FOCUS THE VB WINDOW ENTRY ON THE ESC KEY
'01 OF 02 THE MSDN BY CLOSE BY ESCAPE KEY


Dim mWnd_VB_VbaWindow_MAXIMIZE As Long
Dim GS_cWnd1 As Long, GS_cWnd2 As Long, GS_cWnd3 As Long, i3 As Long


'---------------------------------------
'01 OF 02 DISABLE THE HELP BUTTON ACCIDENT PRESS
'---------------------------------------
MSDN_Lib = FindWindow("#32770", "Microsoft Visual Basic")
If MSDN_Lib > 0 And MSDN_Lib <> OcWnd Then

    '---------------------------------------
    'RESULT = PostMessage(MSDN_Lib, WM_CLOSE, 0&, 0&)
    'Microsoft Visual Basic
    '#32770
    '---------------------------------------
    'INSTEAD OF METHOD 1 USE METHOD 2
    '---------------------------------------
    '1. DONT CLOSE THE HELP ALERT VB INFO
    '2. DISABLE THE HELP BUTTON ACIDENT FLICKER
    '---------------------------------------
    'HWND_2 = GET_CHILD_TEST(MSDN_Lib)
    '---------------------------------------
    
    pwnd = MSDN_Lib
    cWnd = FindWindowEx(pwnd, 0, "BUTTON", "HELP")
    EnableWindow cWnd, False
    
    '---------------------------------------
    '31 AUG 2K SIXTEEN
    'LEARN
    '---------------------------------------
    'Find Window Tutorial Part1 <A must see> by Paul Zaczkowski (from psc cd)
    'https://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37467&lngWId=1
    '---------------------------------------
End If

OcWnd = MSDN_Lib



'----------------------------------------
'FOCUS THE VB WINDOW ENTRY ON THE ESC KEY
'----------------------------------------

FindCursor Nx, Ny
mWnd_VB_VbaWindow_MAXIMIZE = WindowFromPoint(Nx, Ny)

CLASS_VAR = GetActiveWindowClass(mWnd_VB_VbaWindow_MAXIMIZE) 'Mwnd will get reset here if none
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Immediate" Then mWnd_VB_VbaWindow_MAXIMIZE = 0
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Locals" Then mWnd_VB_VbaWindow_MAXIMIZE = 0
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Watches" Then mWnd_VB_VbaWindow_MAXIMIZE = 0

'"VbaWindow" Or CLASS_VAR = "DesignerWindow" Then
If mWnd_VB_VbaWindow_MAXIMIZE > 0 Then
    If O_mWnd_VB_VbaWindow_MAXIMIZE <> mWnd_VB_VbaWindow_MAXIMIZE Then
        If CLASS_VAR = "VbaWindow" Or CLASS_VAR = "DesignerWindow" Then
            
            O_mWnd_VB_VbaWindow_MAXIMIZE = mWnd_VB_VbaWindow_MAXIMIZE
            BEEP
            ShowWindow mWnd_VB_VbaWindow_MAXIMIZE, SW_MAXIMIZE
        End If
    End If
End If
'---------------------------------------



'------------------------------------------------------------------------------
' METHOD 3
' 1 ---- MICROSOFT VISUAL BASIC 01 OF 02 -- MSDN HELP ARRIVE ESC KEY GONE
' 1 ---- MICROSOFT VISUAL BASIC 02 OF 02 -- HELP ARRIVE MAKE BUTTON NOT ENABLED
' 2 ---- GOODSYNC ASK TO CLOSE AT END
' 3 ---- SECURITY BIT SET BUTTON PUSH ---- Error Applying Security
'------------------------------------------------------------------------------

'----------------------------------------------------
'TEMPROARY MISSING CODE WHEN CONFLICT COMPARE SYNC
'----------------------------------------------------
'Call Timer_01_VB_SOUNDBEEP_WHEN_BREAK_POINT_IDE_Timer
'----------------------------------------------------

'---------------------------------------
'01 OF 02 THE MSDN BY CLOSE BY ESCAPE KEY
'---------------------------------------
'CARE NOT MORE THAN ONE HELP WINDOW OPEN
'PRESS ESCAPE KEY CODE 27 TO CLOSE
'---------------------------------------
If GetAsyncKeyState(27) Then
    MSDN_Library = FindWindow(vbNullString, "MSDN Library Visual Studio 6.0")
    If MSDN_Library > 0 Then
        '----------------
        'CLOSE THE WINDOW
        '----------------
        RESULT = PostMessage(MSDN_Library, WM_CLOSE, 0&, 0&)
    End If
End If



'----------------------------------------
'FOCUS THE VB WINDOW ENTRY ON THE ESC KEY
'----------------------------------------






FindCursor Nx, Ny
mWnd_VB_VbaWindow_MAXIMIZE = WindowFromPoint(Nx, Ny)

CLASS_VAR = GetActiveWindowClass(mWnd_VB_VbaWindow_MAXIMIZE) 'Mwnd will get reset here if none
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Immediate" Then mWnd_VB_VbaWindow_MAXIMIZE = 0
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Locals" Then mWnd_VB_VbaWindow_MAXIMIZE = 0
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Watches" Then mWnd_VB_VbaWindow_MAXIMIZE = 0

'"VbaWindow" Or CLASS_VAR = "DesignerWindow" Then
If mWnd_VB_VbaWindow_MAXIMIZE > 0 Then
    If O_mWnd_VB_VbaWindow_MAXIMIZE <> mWnd_VB_VbaWindow_MAXIMIZE Then
        If CLASS_VAR = "VbaWindow" Or CLASS_VAR = "DesignerWindow" Then
            
            O_mWnd_VB_VbaWindow_MAXIMIZE = mWnd_VB_VbaWindow_MAXIMIZE
            BEEP
            ShowWindow mWnd_VB_VbaWindow_MAXIMIZE, SW_MAXIMIZE
        End If
    End If
End If
'---------------------------------------




'---------------------------------------
'MAXIMIZE THE VB WINDOW ENTRY
'---------------------------------------
'MSDN_Lib = FindWindow("#32770", "Microsoft Visual Basic")
'IHWND = FindWindow(vbNullString, "wndclass_desked_gsk")
'I IHWND > 0
'    I_CH_HWND = FindWindowEx("VbaWindow", vbNullString)
'End If

'MAX WINDOWS WITHIN VB FOR ------------------------------------
'VbCode window To Max on Detect

FindCursor Nx, Ny
mWnd_VB_VbaWindow_MAXIMIZE = WindowFromPoint(Nx, Ny)

CLASS_VAR = GetActiveWindowClass(mWnd_VB_VbaWindow_MAXIMIZE) 'Mwnd will get reset here if none
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Immediate" Then mWnd_VB_VbaWindow_MAXIMIZE = 0
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Locals" Then mWnd_VB_VbaWindow_MAXIMIZE = 0
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Watches" Then mWnd_VB_VbaWindow_MAXIMIZE = 0

If mWnd_VB_VbaWindow_MAXIMIZE > 0 Then
    If O_mWnd_VB_VbaWindow_MAXIMIZE <> mWnd_VB_VbaWindow_MAXIMIZE Then
        If CLASS_VAR = "VbaWindow" Or CLASS_VAR = "DesignerWindow" Then
            
            O_mWnd_VB_VbaWindow_MAXIMIZE = mWnd_VB_VbaWindow_MAXIMIZE
            BEEP
            ShowWindow mWnd_VB_VbaWindow_MAXIMIZE, SW_MAXIMIZE
        End If
    End If
End If



'---------------------------------------
'02 OF 02 DISABLE THE HELP BUTTON ACCIDENT PRESS
'---------------------------------------
'2017_FEB_15 EARLY HOUR
'---------------------------------------
TEAM_VIEWER = FindWindow("#32770", "Sponsored session")
If TEAM_VIEWER > 0 And TEAM_VIEWER <> OHWnd_TEAM_VIEWER Then

    '---------------------------------------
    RESULT = PostMessage(TEAM_VIEWER, WM_CLOSE, 0&, 0&)
    'Microsoft Visual Basic
    '#32770
    '---------------------------------------
    'INSTEAD OF METHOD 1 USE METHOD 2
    '---------------------------------------
    '1. DONT CLOSE THE HELP ALERT VB INFO
    '2. DISABLE THE HELP BUTTON ACIDENT FLICKER
    '---------------------------------------
    'HWND_2 = GET_CHILD_TEST(MSDN_Lib)
    '---------------------------------------
    
'    pwnd = TEAM_VIEWER
'    cWnd = FindWindowEx(pwnd, 0, "BUTTON", "HELP")
'    EnableWindow cWnd, False
    
    pwnd = TEAM_VIEWER
    'GS_cWnd4 = 0
    'GS_cWnd1 = FindWindowEx(pwnd, 0, "Edit", vbNullString)
    
    'GS_cWnd2 = FindWindowEx(pwnd, 0, "Button4", "&OK")
    GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
    'GS_cWnd3 = FindWindowEx(pwnd, 0, "button", "&Cancel")
    
    
        
    If GS_cWnd2 > 0 Then
        '-------------------------------------------------
        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
        '1ST RUN DEBUG FOCUSED TO BUTTON BUT NOT HITT IT IN UNTIL 2ND
        'FOCUS PROBLEM FIRST
        '-------------------------------------------------
        For R_REPEAT = 1 To 3
            GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
            If GS_cWnd2 = 0 Then Exit For
            SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
    
    '---------------------------------------
    '31 AUG 2K SIXTEEN
    'LEARN
    '---------------------------------------
    'Find Window Tutorial Part1 <A must see> by Paul Zaczkowski (from psc cd)
    'https://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37467&lngWId=1
    '---------------------------------------
End If

OHWnd_TEAM_VIEWER = TEAM_VIEWER







'---------------------------------------
'03 OF 03 VISIAUL BASIC LOAD CRASH THE CLIPBOARD
'---------------------------------------
'2017_FEB_22 EARLY HOUR 01:43a
'---------------------------------------
'TYPE 01 OF 02
VB_LOADER = FindWindow("#32770", "Visual Component Manager")
If VB_LOADER > 0 And VB_LOADER <> OHWnd_VB_LOADER Then
    pwnd = VB_LOADER
    GS_cWnd1 = FindWindowEx(pwnd, 0, vbNullString, "Can't open Clipboard")
    GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
    If GS_cWnd1 > 0 And GS_cWnd2 > 0 Then
        '-------------------------------------------------
        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
        '1ST RUN DEBUG FOCUSED TO BUTTON BUT NOT HITT IT IN UNTIL 2ND
        'FOCUS PROBLEM FIRST
        '-------------------------------------------------
        For R_REPEAT = 1 To 4
            GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
            If GS_cWnd2 = 0 Then Exit For
            SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
End If



'TYPE 02 OF 02
VB_LOADER = FindWindow("#32770", "Data View")
If VB_LOADER > 0 And VB_LOADER <> OHWnd_VB_LOADER Then
    pwnd = VB_LOADER
    GS_cWnd1 = FindWindowEx(pwnd, 0, vbNullString, "Method '~' of object '~' failed")
    GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
    If GS_cWnd1 > 0 And GS_cWnd2 > 0 Then
        For R_REPEAT = 1 To 4
                GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
                If GS_cWnd2 = 0 Then Exit For
                SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
End If
'TYPE 02 OF 02
'VB_LOADER = FindWindow("#32770", "Add-In Toolbar")
VB_LOADER = FindWindow("#32770", vbNullString)
If VB_LOADER > 0 And VB_LOADER <> OHWnd_VB_LOADER Then
    'OHWnd_VB_CLIPPER_ERROR = VB_CLIPPER_ERROR
    OHWnd_VB_LOADER = VB_LOADER
    pwnd = VB_LOADER
    GS_cWnd1 = FindWindowEx(pwnd, 0, vbNullString, "Method '~' of object '~' failed")
    GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
    If GS_cWnd1 > 0 And GS_cWnd2 > 0 Then
        For R_REPEAT = 1 To 4
                GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
                If GS_cWnd2 = 0 Then Exit For
                SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
End If


VB_CLIPPER_ERROR = FindWindow("#32770", vbNullString)
If VB_CLIPPER_ERROR > 0 And VB_CLIPPER_ERROR <> OHWnd_VB_CLIPPER_ERROR Then
    OHWnd_VB_CLIPPER_ERROR = VB_CLIPPER_ERROR
    pwnd = VB_CLIPPER_ERROR
    GS_cWnd1 = FindWindowEx(pwnd, 0, vbNullString, "Can't open Clipboard")
    GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
    If GS_cWnd1 > 0 And GS_cWnd2 > 0 Then
        For R_REPEAT = 1 To 4
                GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
                If GS_cWnd2 = 0 Then Exit For
                SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
End If




'KILL ITSELF IF RUNNING OR GETTER IN THE WAY OF LEARN
If IsIDE = True Then
    VB_EXE = FindWindow("ThunderRT6FormDC", "KAT MP3 RELOAD")
    If VB_EXE > 0 And VB_EXE <> OHWnd_VB_EXE Then
        RESULT = PostMessage(VB_EXE, WM_CLOSE, 0&, 0&)
    End If
End If


'TYPE 02 OF 02
VB_LOADER = FindWindow("#32770", "Add-In Toolbar")
If 1 = 2 And VB_LOADER > 0 And VB_LOADER <> OHWnd_VB_LOADER Then
'Add-In Toolbar

    '---------------------------------------
    'RESULT = PostMessage(VB_LOADER, WM_CLOSE, 0&, 0&)
    'Microsoft Visual Basic
    '#32770
    '---------------------------------------
    'INSTEAD OF METHOD 1 USE METHOD 2
    '---------------------------------------
    '1. DONT CLOSE THE HELP ALERT VB INFO
    '2. DISABLE THE HELP BUTTON ACIDENT FLICKER
    '---------------------------------------
    'HWND_2 = GET_CHILD_TEST(MSDN_Lib)
    '---------------------------------------
    
'    pwnd = VB_LOADER
'    cWnd = FindWindowEx(pwnd, 0, "BUTTON", "HELP")
'    EnableWindow cWnd, False
    
    pwnd = VB_LOADER
    'GS_cWnd4 = 0
    'GS_cWnd1 = FindWindowEx(pwnd, 0, "Edit", vbNullString)
    
    GS_cWnd1 = FindWindowEx(pwnd, 0, vbNullString, "Data View")
    GS_cWnd1 = FindWindowEx(pwnd, 0, vbNullString, "Can't open Clipboard")
    'Can't open Clipboard
    
    'GS_cWnd2 = FindWindowEx(pwnd, 0, "Button4", "&OK")
    GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
    'GS_cWnd3 = FindWindowEx(pwnd, 0, "button", "&Cancel")
    If GS_cWnd1 > 0 And GS_cWnd2 > 0 Then
    
        
        '-------------------------------------------------
        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
        '1ST RUN DEBUG FOCUSED TO BUTTON BUT NOT HITT IT IN UNTIL 2ND
        'FOCUS PROBLEM FIRST
        '-------------------------------------------------
        For R_REPEAT = 1 To 4
            GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
            If GS_cWnd2 = 0 Then Exit For
            SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
    
    '---------------------------------------
    '31 AUG 2K SIXTEEN
    'LEARN
    '---------------------------------------
    'Find Window Tutorial Part1 <A must see> by Paul Zaczkowski (from psc cd)
    'https://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37467&lngWId=1
    '---------------------------------------
End If


'FILE ALREADY open ____ LOADED _ MOVE ON
'-----------------------------
VB_LOADER = FindWindow("#32770", "Microsoft Visual Basic")
If 1 = 2 And VB_LOADER > 0 And VB_LOADER <> OHWnd_VB_LOADER Then
    pwnd = VB_LOADER
    GS_cWnd1 = FindWindowEx(pwnd, 0, vbNullString, "OK")
    GS_cWnd1 = FindWindowEx(pwnd, 0, "Static", vbNullString)
    TEXT1 = WindowText(GS_cWnd1)
    
    GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
    If GS_cWnd1 > 0 And GS_cWnd2 > 0 Then
        '-------------------------------------------------
        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
        '1ST RUN DEBUG FOCUSED TO BUTTON BUT NOT HITT IT IN UNTIL 2ND
        'FOCUS PROBLEM FIRST
        '-------------------------------------------------
        For R_REPEAT = 1 To 4
            GS_cWnd2 = FindWindowEx(pwnd, 0, vbNullString, "OK")
            If GS_cWnd2 = 0 Then Exit For
            SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
End If












End Sub

Private Sub Timer_VICE_VERSA_TIMER_ON_Timer()

Timer_VICE_VERSA_TIMER_ON.Enabled = False
        
    'MNU_VICE_VERSA_TIMER_ON.Caption = "VICE-VERSA-TIMER-ON"
    Timer_VICE_VERSA_TIMER_ON.Enabled = False
    MNU_VICE_VERSA_TIMER_ON.Caption = "VICE-VERSA-TIMER-OFF__PRESS-BUTTON-TO__GO"

End Sub

Private Sub Timer_W32TM_Timer()

'---------------------------------------
'USE THE SYSTEM SYNC TIME REGULAR - NICE
'-----------------------------------------------------------------------------
'CHECK IT OUT - IN CMD LINE SOMETIMES FAIL BUT WINDOWS RIGHT CORNER DON'T FAIL
'-----------------------------------------------------------------------------

If UNLOAD_FORM_FLAG = True Then Unload Me: Exit Sub


If TIME_CHECK_IF_NOT_BEEN_SET_BACKER > 0 Then
    If TIME_CHECK_IF_NOT_BEEN_SET_BACKER + TimeSerial(0, 1, 0) < Now Then
        If TIME_BEFORE_DATE_SET_BACK = 0 Then
            BEEP
            Form2_TIME_SET_BACK.Caption = "THE TIME HAS BEEN SET BACK MORE THAN 10 SECOND AN WILL BE SET AGAIN IN 1 SECOND"
            Form2_TIME_SET_BACK.Show
            W32X2 = Now + TimeSerial(0, 0, 1)
            TIME_BEFORE_DATE_SET_BACK = TIME_CHECK_IF_NOT_BEEN_SET_BACKER
        End If
    End If
End If

TIME_CHECK_IF_NOT_BEEN_SET_BACKER = Now

'TIME_BEFORE_DATE_SET_BACK = 1

If Now > TIME_BEFORE_DATE_SET_BACK And TIME_BEFORE_DATE_SET_BACK > 0 Then

    BEEP
    TIME_BEFORE_DATE_SET_BACK = 0
    
    Form2_TIME_SET_BACK.TIMER_UNLOAD_FORM_SHOWING.Enabled = True
    
    A = A

End If


'EVERY TEN MIN TIMER

'X2 = Minute(Now)
'For R = 10 To 59 Step 10
'    If R >= X2 Then Exit For
'Next
'XTM = R
'
'If XTM <> X2 And TRIG2_GO3 = 1 Then TRIG2_GO3 = 0
'
'If TRIG2_GO3 = 0 Then
'    If X2 < OX2 Or XTM = X2 Then
'        TRIG1_GO3 = True
'        TRIG2_GO3 = 1
'    End If
'End If
'OX2 = X2
'
'If TRIG1_GO3 = False Then Exit Sub

If W32X2 = 0 Then
    W32X2 = Now + TimeSerial(0, 0, 30)
    Exit Sub
End If

If W32X2 < Now Then
    
    W32_COUNT = W32_COUNT + 1
    If W32_COUNT > 20 Then W32_COUNT = 20
    
    If W32_COUNT = 1 Then
        '------------------------------------------------------------
        'TEMP TAKEN OUT -- ERRRO OF SETUP CODE FROM WEB USELESS AGAIN
        'ONLY THE LAST COMMAND WORK SYNC
        '------------------------------------------------------------
        'Shell "CMD /C ""W32TM /REGISTER""", vbHide
        W32X2 = Now + TimeSerial(0, 0, 10)
        Exit Sub
    End If
    
    
    
    '----
    'TIME SERVERS UK - Google Search
    'https://www.google.co.uk/search?num=30&q=TIME+SERVERS+UK&oq=TIME+SERVERS+UK&gs_l=serp.3..0j0i22i30k1l4.5607.7437.0.7683.3.3.0.0.0.0.109.308.1j2.3.0....0...1c.1.64.serp..0.3.306.yu9yxB8wlf0
    '----
    'Stratum 2 NTP Time Servers
    '--------------------------
    'Hostname    IP Address  DNS
    '--------------------------------------
    'ntp2c.mcc.ac.uk 130.88.200.4    Yes
    'ntp.exnet.com   194.207.34.9    Yes
    'ntp1.sandvika.net   194.164.127.5   No
    'ntp.cis.strath.ac.uk Yes
    '--------------------------------------
    '5 more rows, 1 more column
    '--------------------------------------

    '-------------------
    '1..time.windows.com
    '2..ntp2c.mcc.ac.uk
    '3..ntp.exnet.com
    '-------------------
    'ABLE USE SIX TIME SERVER AS URL PAGE TALK
    'UNABLE TO UNDERSTNAD HOW NETWORK USE IS ENABLED
    'ON SITE VS OFF SITE
    'MAYBE IT IS ALREADY WORKING AND THAT IS WHAT IT DOES -- MAYBE
    '-------------------------------------------------------------
    '<QUOTE>
    'Each workstation and server in this network will try to locate a time source for synchronization. Using an internal algorithm designed to reduce network traffic, systems will make up to six attempts to find a time source. Here's a look at the order of these attempts:
    '
    'Parent domain controller (on-site)
    'Local domain controller (on-site)
    'Local PDC emulator (on-site)
    'Parent domain controller (off-site)
    'Local domain controller (off-site)
    'Local PDC emulator (off-site)
    '-----------------------------------------------
    
    If W32_COUNT = 2 Then
        '------------------------------------------------------------
        'TEMP TAKEN OUT -- ERRRO OF SETUP CODE FROM WEB USELESS AGAIN
        'ONLY THE LAST COMMAND WORK SYNC
        '------------------------------------------------------------
        'Shell "CMD /C ""W32tm /config /manualpeerlist:""time.windows.com"",""ntp2c.mcc.ac.uk"",""ntp.exnet.com"" /syncfromflags:manual""", vbHide
        W32X2 = Now + TimeSerial(0, 0, 10)
        Exit Sub
    End If

    If W32_COUNT = 3 Then
        '------------------------------------------------------------
        'TEMP TAKEN OUT -- ERRRO OF SETUP CODE FROM WEB USELESS AGAIN
        'ONLY THE LAST COMMAND WORK SYNC
        '------------------------------------------------------------
        'Shell "CMD /C ""W32tm /config /update""", vbHide
        W32X2 = Now + TimeSerial(0, 0, 10)
        Exit Sub
    End If
    
    
    'Sun 14 August 2016 17:00:00
    '----------------------------------------------------------------------------------------------------
    'ALL AS DESCRIBED HERE 4 STAGE
    '----
    'Synchronize time throughout your entire Windows network - Page 6040425 - TechRepublic
    'http://www.techrepublic.com/article/synchronize-time-throughout-your-entire-windows-network/6040425/
    '----
    '----------------------------------------------------------------------------------------------------


    'PAY FOR
    '----
    'Time synchronization in a local network
    'http://clocksynchro.com/
    '----
    
    If W32_COUNT >= 4 Then
        Shell "CMD /C ""W32TM /RESYNC""", vbHide
        'Shell "CMD /C ""W32TM /RESYNC """, vbNormalFocus
        If W32_COUNT < 20 Then
            W32X2 = Now + TimeSerial(0, 1, 0)
        Else
            W32X2 = Now + TimeSerial(0, 3, 0)
        End If
    End If
    
    
    'If W32_COUNT >= 3 Then
    
'    If TRIG1_GO3 = True Or W32TM_RUN_ONCE = False Then
'        W32TM_RUN_ONCE = True
        '    TRIG1_GO3 = False
            
            'START "" /REALTIME "W32TM" "/RESYNC"
            
            'K TO STAY - SEE RESULT
            'Shell "CMD /K ""W32TM /RESYNC""", vbNormalFocus
            
            'THE TIME.WINDOWS SERVER IS NOT ALWAYS GETTING THE CLOCK
            
        'Shell "CMD ""W32TM /RESYNC""", vbHide
            
        
    'End If

End If


End Sub

Private Sub Timer_WINMERGE_KILL_Timer()
    
    If GetComputerName <> "1-ASUS-X5DIJ" Then Exit Sub

    '------------------------------------
    'PROBLEM IN WIN XP FOR THIS STAY OPEN
    '------------------------------------
    'NOT WITH WIN 7 WIN 10 OKAY
    '------------------------------------
    
    If UNLOAD_FORM_FLAG = True Then Unload Me: Exit Sub
    
    If XX = 0 Then
        XX = FindWindow("WinMergeWindowClassW", vbNullString)
    End If
    If XX = 0 Then Exit Sub
    
    If GetWindowState(XX) <> -2 Then
        Exit Sub
    End If
    
    XX = 0

    'DO AT BEGINING DECLARE
    'Dim PID As Long

    'PID HAS TO BE -1
    PID = -1
    Var = cProcesses.GetEXEID(PID, "C:\Program Files\WinMerge\WinMergeU.exe")
    
    Var = cProcesses.Process_Kill(PID)

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
    
'    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    
    WSHShell.Run """" + App.Path + "\RELOAD AND COPY.VBS" + """"
    Set WSHShell = Nothing
    Unload Me
    
End If

XVB_DATE = VB_DATE

End Sub



Private Sub TIMER_RETRY_WRITE_INFO_UNTIL_DONE2_Timer()
    
    On Error Resume Next
    FR1 = FreeFile
    Open "D:\VB6-EXE'S\#00 RELOAD_APPEND_SCRIPT\## SCRIPT RELOAD APPEND.TXT" For Append As #FR1
        VARTIME = Format(Now, "YYYY MM DD HH MM SS") + " -- " + PATH_FILE_NAME1
    Close #FR1
    
    If Err.Number = 0 Then
        TIMER_RETRY_WRITE_INFO_UNTIL_DONE2.Enabled = False
    End If

End Sub


Private Sub Timer1_Timer()
    
If UNLOAD_FORM_FLAG = True Then Unload Me: Exit Sub

'If IsIDE = False Then Timer1.Interval = 1000
Timer1.Interval = 4000

Call VB_PROJECT_CHECKDATE
    
If KAT_MP3_PATH = "" And KAT_MP3_PATH_UNTOLD = False Then
    KAT_MP3_PATH_UNTOLD = True
    MsgBox "KAT MP3 APP PATH NOT FOUND" + vbCrLf + KAT_MP3_PATH1 + vbCrLf + "OR" + vbCrLf + KAT_MP3_PATH2, vbMsgBoxSetForeground
End If
If KAT_MP3_PATH_UNTOLD = True Then Exit Sub

    
TIME_HAPPEN = Format(Now, "MM-DD DDD-DD-MMM -- HH:MM:SS")

If FIRST_RUN_FIND_PID_KAT_MP3 = False Then
    FIRST_RUN_FIND_PID_KAT_MP3 = True
    'PID HAS TO BE -1
    PID = -1
    '---------------------------------------------------------
    'THE PATH WILL BE STRIPPED SO WE SHALL TEMP IT TO PRESERVE
    '---------------------------------------------------------
    KAT_MP3_PATH_TMP = KAT_MP3_PATH
    Var = cProcesses.GetEXEID(PID, KAT_MP3_PATH_TMP) '+ "C:\Program Files\Kat MP3 Recorder\Kat MP3 Recorder.EXE")
    If PID = -1 Then
        
        Shell KAT_MP3_PATH, vbMinimizedNoFocus
        COUNT_UP = Now
        Label1.Caption = TIME_HAPPEN + " -- Kat MP3 Recorder.EXE - NOT IN EXECUTION AT START OF ME PROGRAM"
        List1.AddItem Label1.Caption
        
        Exit Sub
    End If
End If

GO3 = 0

Set F = FS.GetFolder(FILE1.Path)

'-----------------------------------
'WHAT TO DO IF COUNT UP REACH TARGET
'-----------------------------------
'-----------------------------------
' "1-ASUS-X5DIJ" RECORDS FOR 5 MIN 300 SEC
' -------------- GOT 7 SECX DELAY TO ADD
' "3-ASUS-EEE"   RECORDS FOR 30 SEC
'-----------------------------------
'SET MINICOMPARE HAS TO BE A FAIR BIT HIGHER THAN THE RECORD LENGHT
'30 SEC WILL BE 55 ABOUT 25 HIGHER
'SET LIMIT HIGH
'CLOSE TO LIMIT MIGHT NEED ADJUSTING
'ACER 55 LIMIT AND MAX 53 SHOW
'-----------------------------------
'-----------------------------------
MINICOMPARE = 400
If GetComputerName = "1-ASUS-EEE" Then MINICOMPARE = 80
If GetComputerName = "1-ASUS-X5DIJ" Then MINICOMPARE = 315
If GetComputerName = "3-LINDA-PC" Then MINICOMPARE = 315 + 500 'ADD ANOTHER MULTI MINUTE MUM COMPUTER DO RUN SLOW LESS QUICK
If GetComputerName = "4-ASUS-GL522VW" Then MINICOMPARE = 315
If GetComputerName = "5-ASUS-P2520LA" Then MINICOMPARE = 315
If GetComputerName = "7-ASUS-GL522VW" Then MINICOMPARE = 315
'-----------------------------------


'------------------------------
TEST_DEBUG_VAR_FSASIZE = F.Size

If TEST_DEBUG_VAR_FSASIZE <> OXS And OXS > 0 Then

    COUNT_UP = Now
    Label1.Caption = TIME_HAPPEN + " -- FOLDER SIZE"
    Call COUNT_TIMER
    FIRST_LIMIT_RESET_DONE = True
    OXS = TEST_DEBUG_VAR_FSASIZE
    Set F = Nothing
    Exit Sub
End If
If OXS = 0 Then OXS = TEST_DEBUG_VAR_FSASIZE
Set F = Nothing

'FILE1.Refresh

x = FILE1.ListCount
If OX <> x And OX > 0 Then

    COUNT_UP = Now
    
    'WHY THIS - IS BLOCKING WORK PROPER
    COUNT_UP_WAIT = 1
    
    Call COUNT_TIMER
    FIRST_LIMIT_RESET_DONE = True
    Label1.Caption = TIME_HAPPEN + " -- FILE COUNT"

    OX = x
    Exit Sub
End If
If OX = 0 Then OX = x
If x = 0 Then Exit Sub

If FILE1.ListCount - 1 < 0 Then
    'FINDWINDOW
    
    If XTELL_ONCE = False Then
        MsgBox "TOO MANY FILE IN THIS PATH " + vbCrLf + VBCRL + FILE1.Path + " TAKING US THERE -- ONLY BE TOLD ONCE DURING RUN OR ELSE IF SITU RECTIFIED -- FUNCTION HERE TIMER1 WILL NOT RUN", vbMsgBoxSetForeground
        Shell "EXPLORER " + FILE1.Path, vbNormalFocus
    End If
    XTELL_ONCE = True
    Exit Sub
Else
    XTELL_ONCE = False
End If

Set F = FS.GetFile((FILE1.Path + "\" + FILE1.List(FILE1.ListCount - 1)))
If F.DateLastModified <> XDATE And XDATE <> 0 Then
    
    If FIRST_LIMIT_RESET_DONE = True Then

        COUNT_UP = Now
        COUNT_UP = F.DateLastModified
        Call COUNT_TIMER
        Label1.Caption = TIME_HAPPEN + " -- FILE DATE MODIFIED"

    End If
    
    XDATE = F.DateLastModified


    'If GetComputerName = "1-ASUS-X5DIJ" Then MINICOMPARE = 302 Else MINICOMPARE = 40
    
    If DateDiff("S", F.DateLastModified, Now) > MINICOMPARE Then
        RUN_CODE = True
    Else
        Exit Sub
    End If
End If

If XDATE = 0 Then XDATE = F.DateLastModified



If RUN_CODE = True Then
    COUNT_UP = 0
    FIRST_LIMIT_RESET_DONE = True
    
End If

TargetHwnd = FindWindow(vbNullString, "Kat MP3 Recorder.EXE - Application Error")
If TargetHwnd <> 0 Then

    'KAT MP3 STILL SEEM TO CONTINUE WITH THIS ERROR

    COUNT_UP = 0
    FIRST_LIMIT_RESET_DONE = True
    RESULT = PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)
    
'    I = SetForegroundWindow(FindWindow(vbNullString, "Kat MP3 Recorder.EXE - Application Error"))
'    PostMessage
'    SendKeys "{ESC}", True
    
    GO3 = 1
    Label1.Caption = TIME_HAPPEN + " -- Kat MP3 Recorder.EXE - Application Error"

End If

TargetHwnd = FindWindow("#32770", "Kat MP3 Recorder")
If TargetHwnd > 0 Then


    'CRASH WITH THIS ERROR
    i = GetWindowRect(TargetHwnd, KAT_XY)
    'NORMAL FIXED WIN SIZE -- HAS SAME FIND_WINDOW CLASS AND NAME WHEN ERROR MSGBOX
    XIOP = 0
    '642
    If KAT_XY.Right - KAT_XY.Left < 500 Then XIOP = 1
    '422
    If KAT_XY.Bottom - KAT_XY.Top < 400 Then XIOP = 1


    If XIOP = 1 Then
        COUNT_UP = 0
        FIRST_LIMIT_RESET_DONE = True
        
        RESULT = PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)
        
        GO3 = 1
        Label1.Caption = TIME_HAPPEN + " -- Kat MP3 Recorder - MSGBOX - The parameter is incorrect. CRASH"
    End If
End If



If GO3 = 1 Or (COUNT_UP > 0 And (DateDiff("S", COUNT_UP, Now) > MINICOMPARE) And COUNT_UP_WAIT = COUNT_UP_WAIT) Then

    
    If Label1.Caption = "" Or Label1.Caption = "Label1" Then Label1.Caption = TIME_HAPPEN + " -- NOT REASON SET"
    
    List1.AddItem Label1.Caption
        
    RESULT = PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)
    
    'WAIT 4 SEC
    XC = 0
    Do
        Sleep 1
        XC = XC + 1
    Loop Until XC = 4
    
    Call KILL_KAT_MP3_AND_THEN_SHELL

    
    COUNT_UP = 0
    COUNT_UP_WAIT = 0
    OOTTX2 = 0
    Exit Sub

End If

Call COUNT_TIMER

End Sub

Sub KILL_KAT_MP3_AND_THEN_SHELL()

    KAT_MP3_PATH_TMP = KAT_MP3_PATH
    
    'PID HAS TO BE -1
    PID = -1
    Var = cProcesses.GetEXEID(PID, KAT_MP3_PATH_TMP)
    
    'WANT TO KILL -- MIGHT ALREADY DISAPPEARED
    If PID <> -1 Then
        'PID HAS TO BE -1
        PID = -1
        Var = cProcesses.GetEXEID(PID, KAT_MP3_PATH_TMP)
        
        Var = cProcesses.Process_Kill(PID)
    End If
    
    Sleep 1
    

    
    'Call DEL_WAVS
    'Call DEL_WAVS
    
    Shell KAT_MP3_PATH, vbMinimizedNoFocus

    'Shell "D:\KAT MP3 RECORDER\KILL_MP3 ---- .BAT"
    Call RUN_SCRIPT_DEL_WAV_VBS

End Sub

Sub RUN_SCRIPT_DEL_WAV_VBS()
    
    '-----------------------------------------------------------
    SCRIPT_TO_RUN = App.Path + "\KILL_KAT MP3 _VB COPY ----.VBS"
    SCRIPT_TO_RUN = App.Path + "\KILL_KAT MP3 __QUIETLY.VBS"
    '
    If Dir(SCRIPT_TO_RUN) <> "" Then
        Set WSHShell = CreateObject("WScript.Shell")
        WSHShell.Run """" + SCRIPT_TO_RUN + """"
        Set WSHShell = Nothing
    End If

    If Dir(SCRIPT_TO_RUN) = "" Then
        If MSgBox_NOT_EXIST_SCRIPT_TO_RUN = False Then
            MSgBox_NOT_EXIST_SCRIPT_TO_RUN = True
            MsgBox "NOT EXIST __ ONCE MESSSENGER" + vbCrLf + vbCrLf + SCRIPT_TO_RUN, vbMsgBoxSetForeground
        End If
    End If

End Sub

'Public Sub KAT_MP3_IF_NOT_EXIST_AND_THEN_SHELL_FUNC()

'    Exit Sub
'
'    KAT_MP3_IF_NOT_EXIST_AND_THEN_SHELL = False
'
'    KAT_MP3_PATH_TMP = KAT_MP3_PATH
'
'    'PID HAS TO BE -1
'    PID = -1
'    Var = cProcesses.GetEXEID(PID, KAT_MP3_PATH_TMP)
'
'    'EXIST
'    If PID <> -1 Then
'        'Var = cProcesses.Process_Kill(PID)
'        'Exit Function
'    End If
'
'    'Sleep 1
'
'    'Shell KAT_MP3_PATH, vbMinimizedNoFocus
'
'    'Call KAT_MP3_IF_NOT_EXIST_AND_THEN_SHELL_FUNC
'    '-----------------------------------------
'
'    Call RUN_SCRIPT_DEL_WAV_VBS
'
'
'
'
''    If Dir("C:\SCRIPTER\SCRIPTER -- VBS\KILL_MP3 ---- .BAT") <> "" Then
''        Shell "C:\SCRIPTER\SCRIPTER -- VBS\KILL_MP3 ---- .BAT"
''    Else
''        If MSgBox_NOT_EXIST_C_SCRIPTER_SCRIPTER_VBS_KILL_MP3_BAT = False Then
''            MSgBox_NOT_EXIST_C_SCRIPTER_SCRIPTER_VBS_KILL_MP3_BAT = True
''            MsgBox "NOT EXIST" + vbCrLf + vbCrLf + "C:\SCRIPTER\SCRIPTER -- VBS\KILL_MP3 ---- .BAT", vbMsgBoxSetForeground
''        End If
''    End If
'
'
''    If Dir("C:\SCRIPTER\SCRIPTER -- VBS\KILL_MP3 ---- .BAT") <> "" Then
''        Shell "C:\SCRIPTER\SCRIPTER -- VBS\KILL_MP3 ---- .BAT"
''    Else
''        If MSgBox_NOT_EXIST_C_SCRIPTER_SCRIPTER_VBS_KILL_MP3_BAT = False Then
''            MSgBox_NOT_EXIST_C_SCRIPTER_SCRIPTER_VBS_KILL_MP3_BAT = True
''            MsgBox "NOT EXIST" + vbCrLf + vbCrLf + "C:\SCRIPTER\SCRIPTER -- VBS\KILL_MP3 ---- .BAT", vbMsgBoxSetForeground
''        End If
''    End If
'
'    'Call DEL_WAVS
'    'Call DEL_WAVS
'
'End Sub


Sub COUNT_TIMER()

If COUNT_UP = 0 Then Exit Sub
TTX = DateDiff("S", COUNT_UP, Now)

If TTX < OTTX Then
    OOTTX = OTTX
    
    
    'If FIRST_LIMIT_RESET_DONE = True Then
        'If COUNT_UP_WAIT = 1 Then
        '    If TTX > OOTTX2 And TTX < MINICOMPARE Then OOTTX3 = OTTX
        'End If
    'End If
End If
OTTX = TTX
If OOTTX2 < TTX Then OOTTX2 = TTX

MNU_TIMER_SINCE_ACTIVITY.Caption = "-- " + Trim(Str(TTX)) + " SECS -- LIMIT IS --" + Str(MINICOMPARE) + " -- LAST MAXIMUM WAS --" + Str(OOTTX) + " -- MAX UNDER LIMIT IS --" + Str(OOTTX2) + " --"



End Sub


 

Sub TIMER_DEL_CAMERA_EMPTY_MP4_MOV_TIMER()

        
    If Timer_VICE_VERSA_TIMER_ON.Enabled = False Then Exit Sub
    
    
    CAMERA_EXPLORER_PATH = "W:\DCIM"

    PATH_SET_TO_HX60V = False
    PATH_SET_TO_SOMETHING = False
    MUST_BE_FUJI_XP90_ = False
    MUST_BE_SONY_HX60V = False
    
    If Dir("W:\MP_ROOT", vbDirectory) <> "" Then
        CAMERA_EXPLORER_PATH = "W:\MP_ROOT"
        PATH_SET_TO_HX60V = True
        PATH_SET_TO_SOMETHING = True
    End If
    
    If Dir("W:\DCIM", vbDirectory) <> "" And PATH_SET_TO_HX60V <> True Then
        CAMERA_EXPLORER_PATH = "W:\DCIM"
        PATH_SET_TO_SOMETHING = True
        '-------------------------------
        MUST_BE_FUJI_XP90_ = True
        '-------------------------------
        Else
        MUST_BE_SONY_HX60V = True
    End If



    If PATH_SET_TO_SOMETHING = False Then Exit Sub
    'If Dir("W:\MP_ROOT", vbDirectory) = "" Then Exit Sub
    
    DoEvents
    Dir1.Path = "W:\DCIM"
    Dir1.Refresh
    DoEvents
    
    For R = 0 To Dir1.ListCount - 1
        DoEvents
        
        On Error Resume Next
            Set F = FS.GetFolder((Dir1.List(R)))
            If Err.Number > 0 Then Exit Sub
        On Error GoTo 0
    
        '19 IS NOT IF MARKER FOLDER
        If F.Size = 0 And Len(Dir1.List(R)) <= 19 Then
            
            On Error Resume Next
                RmDir Dir1.List(R)
            On Error GoTo 0
        End If
    Next
    
    TIMER_DEL_CAMERA_EMPTY_MP4_MOV.Enabled = False
    
    'Debug.Print GetSpecialfolder(38) + "\ViceVersa Pro 2\ViceVersa.exe ""D:\VICE_VERSA SCRIPT FILE\viceversa photo camera - me.fsf"""
    
'    If Dir(GetSpecialfolder(38) + "\ViceVersa Pro 2\ViceVersa.exe") <> "" Then
'        VVCOMLINE = GetSpecialfolder(38) + "\ViceVersa Pro 2\ViceVersa.exe"
'    End If
    
    'Call GetSpecialFolder_Show_Script_Debug(0)
    
    Call FIND_VICE_VERSA

    
'    VVCOMLINE1 = GetSpecialfolder(38) + "\VICEVERSA PRO\VICEVERSA.EXE"
'    If Dir(VVCOMLINE1) <> "" Then
'        VVCOMLINE = ""
'    End If
'
'    VVCOMLINE2 = Replace(UCase(GetSpecialfolder(38)), " (X86)", "") + "\VICEVERSA PRO\VICEVERSA.EXE"
'    If Dir(VVCOMLINE2) = "" Then
'        VVCOMLINE = ""
'    End If
'
'    If VVCOMLINE = "" Then
'        MsgBox "UNABLE TO FIND VICEVERSA PROGRAM @" + vbCrLf + VVCOMLINE1 + vbCrLf + "OR" + vbCrLf + VVCOMLINE2
'        Exit Sub
'    End If
    
    
    
    'PROBLEM CANNOT GET HWND OF EXPLORER WHEN GOT PROCESS ID
    'NORMAL PROGRAM OKAY
    'RESULT TALK IN WIN 7
    
    
    
'    CAMERA_EXPLORER_PATH = "W:\DCIM"
    
    ONE_INSTANCE_ONLY = True
    i = FindWinPart_VISIBLE(CAMERA_EXPLORER_PATH, "CabinetWClass", HWND_RETURN, VAR1, KILL_SOFTLY, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)
    
    HOW_MANY_INSTANCE_SHOWING_VAR = HOW_MANY_INSTANCE_SHOWING
    
    If HOW_MANY_INSTANCE_SHOWING = 0 Then
        
        'DISABLED THIS FOR A TIME
        'PID = Shell("EXPLORER " + CAMERA_EXPLORER_PATH, vbNormalNoFocus)
        
        ONE_INSTANCE_ONLY = True
        ESCAPEE_COUNT = 20
        Do
            'DISABLED THIS FOR A TIME
            Exit Do
            
            i = FindWinPart_VISIBLE(CAMERA_EXPLORER_PATH, "CabinetWClass", HWND_RETURN, VAR1, KILL_SOFTLY, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)
            
            '-------------------------------------------
            '1. FIRST IS SHELL PROCESS
            '2. THEN WINDOW MAKE VISIBLE
            '3. AND THEN HAS TO BE SETFOREGROUND FOCUS
            '-------------------------------------------
            '3. NOT GOOD IDEA AS STEAL FOCUS PROBLEM MAY LEAD
            '   ON TO TIMEOUT
            '----------------------------------------------------------------
            'STEAL FOCUS PROBLEM SORTED GOOD SEE END PARAGRAPGH IN THIS BLOCK
            'MUST NOT LOAD ANY QUICKER FOR END RESULT
            '----------------------------------------------------------------
            '
            '-------------------------------------------------
            'ONE AT LEAST HAS TO BE FOREGROUND
            'OR OUR SET FOCUS WILL OVER LAP AND BE UNDER
            '-------------------------------------------------
            'NOT TRUE SOMETHING CAN LOAD WITH FOCUS FROM SHELL
            'BUT NOTHING CAN STEAL FOCUS BACK AGAIN
            '-------------------------------------------------
            'GETFOREGROUND IS THEN BETTER WITH SOMETHING LOAD
            'FROM SHELL WITH FOCUS
            '-------------------------------------------------
            'NOT USE WITH THIS EXPLORER LOAD
            '-------------------------------------------------
            'FINISH END RESULT WITH WAIT EXPLORER TO LOAD ALSO AT BEGIN
            'AND THEN THAT SORT THE STEAL FOCUS PROBLEM OUT
            'RESULT
            'TOOK ALL DAY -- END -- THUR 19 MAY
            '-------------------------------------------------
            
            
            'If GetForegroundWindow_VAR = True Then
                If HOW_MANY_INSTANCE_SHOWING_VAR < HOW_MANY_INSTANCE_SHOWING Then
                    Exit Do
                End If
                If MATCH_HWND = True Then
                    Exit Do
                End If
            'End If
            
            DoEvents
            Sleep 500
            ESCAPEE_COUNT = ESCAPEE_COUNT - 1
        Loop Until ESCAPEE_COUNT <= 0
        'Debug.Print "TIMEOUT RETRY COUNTDOWN TAKEN FROM 20 SECONDS --" + Str(ESCAPEE_COUNT)
    
    End If
    
    
    
    
    
    '-------------------------------------------------
    'VICEVERSA ALWAYS A TRICKY PROBLEM
    'AS USE MANY FORM
    'THE PID CONVERSION TO HWND FROM SHELL
    'IS DIFFERENT TO THE FIND WINDOW FROM MAIN FORM
    'AND THE PARENT OF BOTH EQUAL THEY ARE MAIN PARENT
    'AND DON'T GO LOWER TO NEXT GW_HWNDNEXT
    'MAYBE TYR GO UP OTHER WAY LATER
    '-------------------------------------------------
    'MAKE DIFFERCULT TO FIND ONE CAME FROM SHELL
    '-------------------------------------------------
    'ANOTHER ANSWER
    'CHECK FOR PREVIOUS INSTANCE
    'WILL PREVENT TIMEOUT DELAY WHEN SHELL WAIT
    '--------------------------------------------------------------
    i = FindWinPart_VISIBLE("ViceVersa Pro - [", "", HWND_RETURN, VAR1, KILL_SOFTLY, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)
    HOW_MANY_INSTANCE_SHOWING_VAR = HOW_MANY_INSTANCE_SHOWING
    '-------------------------------------------
    '-------------------------------------------
    '1. FIRST IS SHELL PROCESS
    '2. THEN WINDOW MAKE VISIBLE
    '3. AND THEN HAS TO BE SETFOREGROUND FOCUS
    '-------------------------------------------
    
    VVCOMLINE = GetSpecialfolder(38) + "\VICEVERSA PRO\VICEVERSA.EXE"
    If Dir(GetSpecialfolder(38) + "\VICEVERSA PRO\VICEVERSA.EXE") <> "" Then
        VVCOMLINE = ""
    End If
    
    VVCOMLINE = Replace(UCase(GetSpecialfolder(38)), " (X86)", "") + "\VICEVERSA PRO\VICEVERSA.EXE"
    If Dir(VVCOMLINE) = "" Then
        VVCOMLINE = ""
    End If
    
    
    
    
    
    If MUST_BE_SONY_HX60V = True Then
        VICE_VERSA_SCRIPT_FILE = "D:\VICE_VERSA SCRIPT FILE\ViceVersa PHOTO CAMERA - LOCAL.fsf"
    End If
    
    If MUST_BE_FUJI_XP90_ = True Then
        VICE_VERSA_SCRIPT_FILE = "D:\VICE_VERSA SCRIPT FILE\ViceVersa PHOTO CAMERA - LOCAL _ FUJI XP90.fsf"
    End If
    
    If Dir(VICE_VERSA_SCRIPT_FILE) = "" Then
        MsgBox "UNABLE TO FIND" + vbCrLf + VICE_VERSA_SCRIPT_FILE + vbCrLf, vbMsgBoxSetForeground
        Exit Sub
    End If
    
    If LATEST_DATE_FILE_ON_CAMERA_CHANGED = False Then
        Exit Sub
    End If
    PID = Shell(VVCOMLINE + " """ + VICE_VERSA_SCRIPT_FILE + """", vbNormalFocus)

    i = cProcesses.Convert(PID, HWND_RETURN, cnFromProcessID Or cnTohWnd)
    
    '--------------------------------
    'EXAMPLE SHOW REULT OF GET PARENT
    'RESULT ENTRY SAME AS PARENT
    'HWND_2 = GET_PARENT(HWND_RETURN)
    '--------------------------------
    'HWND_2 = GET_CHILD_TEST(HWND_RETURN)
    '--------------------------------
    
    '----------------------------------------------------------
    'GONE IN FULL CIRCLE THE SHELL HWND RETURN IS SAME AS
    'ONE GET FROM FINDWINPART
    'NOT ALWAYS
    'Debug.Print "HWND FROM SHELL TO HWND =" + Str(HWND_RETURN)
    'RESULT -- IF LOAD QUICK WITH ANY DEBUG BREAKPOINT
    '----------------------------------------------------------
    'HAS RESULT WANT MATCH
    'IS KNOWN TO BE TOLD
    'THAT HWND CAN CHANGE IT VAULE
    '------------------------------------------------------------------------------
    'DONT ALWAYS PROVIDE RESULT WANT WHEN MORE THEN 3 LOADED ON LESS QUICK COMPUTER
    '------------------------------------------------------------------------------
    'CONFIRM THAT RETRY LONGER TIMEOUT
    'DONT ALWAYS WANT IT WHEN UPTO 3 OR MORE
    '---------------------------------------
    'TEST UPTO EIGHT - AND NOT RESULT
    '---------------------------------------------------------
    'AND AFTER BACK TO SHOW AGAIN NOT ALWAYS WORK EVEN WITH ONE
    'GUESS WAIT ON THE TIMEOUT OF 20 SEC
    '----------------------------------------------------------
    'LAST IDEA - COUNT THE INSTANCES SHOWING
    'OBLITERATE TIMEOUT USE
    'AND ALSO WITH MATCH_HWND AS WON'T CONFLICT IF ALREADY PREVIOUS INSTANCE
    '-----------------------------------------------------------------------
    
    'Call WAIT_WINDOW_VISIBLE_AFTER_SHELL(HWND_RETURN, TIMEOUT_WAIT_USED)
    
    ESCAPEE_COUNT = 20
    Do
        i = FindWinPart_VISIBLE("ViceVersa Pro - [", "", HWND_RETURN, VAR1, KILL_SOFTLY, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)
        
        '-------------------------------------------
        '1. FIRST IS SHELL PROCESS
        '2. THEN WINDOW MAKE VISIBLE
        '3. AND THEN HAS TO BE SETFOREGROUND FOCUS
        '-------------------------------------------
        '3. NOT GOOD IDEA AS STEAL FOCUS PROBLEM MAY LEAD
        '   ON TO TIMEOUT
        '------------------------------------------------------
        
        '-------------------------------------------------
        'ONE AT LEAST HAS TO BE FOREGROUND
        'OR OUR SET FOCUS WILL OVER LAP AND BE UNDER
        '-------------------------------------------------
        'NOT TRUE SOMETHING CAN LOAD WITH FOCUS FROM SHELL
        'BUT NOTHING CAN STEAL FOCUS BACK AGAIN
        '-------------------------------------------------
        'GETFOREGROUND IS THEN BETTER WITH SOMETHING LOAD
        'FROM SHELL WITH FOCUS
        '-------------------------------------------------
        
        'If GetForegroundWindow_VAR = True Then
            If HOW_MANY_INSTANCE_SHOWING_VAR < HOW_MANY_INSTANCE_SHOWING Then
                Exit Do
            End If
            If MATCH_HWND = True Then
                Exit Do
            End If
        'End If
        
        DoEvents
        Sleep 500
        ESCAPEE_COUNT = ESCAPEE_COUNT - 1
    Loop Until ESCAPEE_COUNT <= 0
'    Debug.Print "TIMEOUT RETRY COUNTDOWN TAKEN FROM 20 SECONDS --" + Str(ESCAPEE_COUNT)

    If ESCAPEE_COUNT = 0 Then
        'MsgBox "TIMEOUT WAIT REACH 0 FROM 10 SEC"
    End If
    
    If UNLOAD_FORM_FLAG = True Then Unload Me: Exit Sub
    
    Timer_ME_CAMERA.Enabled = True

End Sub




Sub TIMER_ME_CAMERA_TIMER()
    '-------------------------------
    'RUNNER BOTH HX60V AND FUJI XP90
    '-------------------------------
    
    If Timer_VICE_VERSA_TIMER_ON.Enabled = False Then Exit Sub

    
    Timer_ME_CAMERA.Enabled = False
    'Me.SetFocus
    DoEvents
    'LOST ABLE TO GAIN FOCUS TO OUR MAIN SCREN BACK
    'TRY SET WINDOWSTATE MINIMIZED AND THEN NORMAL AGAIN
    'Me.WindowState = vbMinimized
    'DoEvents
    'Me.WindowState = vbNormal
    'DoEvents
    'Me.SetFocus
    'DoEvents
    'AT LEAST IT FLASHES AT THIS STAGE
    'TRY COMMAND SET FOCUS
    'I = SetForegroundWindow(Me.hWnd)
    'DoEvents
    '-------------------------------------------------
    'OH CRAPPER WINDOWS 7 NOT SET TO ALLOW STEAL FOCUS
    'AT THE MOMENT
    '-------------------------------------------------
    'AND THEN NEXT RESULT
    'A NEW FORM ALSO CANNOT STEAL FOCUS
    '-------------------------------------------------
    
    F_SIZE_SET = False
        
    '---------------------------------------------------------------
    'HAS THE CAMERA DRIVE SD CARD BEEN UNPLUGED JUST BEFORE -- CHECK
    On Error Resume Next
    
    PATH_SET_TO_HX60V = False
    PATH_SET_TO_SOMETHING = False
    MUST_BE_FUJI_XP90_ = False
    MUST_BE_SONY_HX60V = False
    
    If Dir("W:\MP_ROOT", vbDirectory) <> "" Then
        Dir2.Path = "W:\MP_ROOT"
        PATH_SET_TO_HX60V = True
        PATH_SET_TO_SOMETHING = True
    End If
    
    If Dir("W:\DCIM", vbDirectory) <> "" And PATH_SET_TO_HX60V <> True Then
        Dir2.Path = "W:\DCIM"
        PATH_SET_TO_SOMETHING = True
        MUST_BE_FUJI_XP90_ = True
    Else
        MUST_BE_SONY_HX60V = True
    End If
    
    If PATH_SET_TO_SOMETHING = False Then Exit Sub
    '---------------------------------------------
    
    If Err.Number > 0 Then Exit Sub
    On Error GoTo 0
    'WHY NOT REMOVE ERROR CHECK SEE IF MORE FAULT
    '--------------------------------------------
    
    File2.FileName = "*.*"
    '-----------------
    'FIND ANY BIG FILE
    '-----------------
    For R1 = 0 To Dir2.ListCount - 1
        On Error Resume Next
            File2.Path = Dir2.List(R1)
            If Err.Number > 0 Then
                Timer_ME_CAMERA.Enabled = True
                Exit Sub
            End If
        On Error GoTo 0
            
        If File2.ListCount > 0 Then
            For R2 = 0 To File2.ListCount - 1
                On Error Resume Next
                    Set F = FS.GetFile(File2.Path + "\" + File2.List(R2))
                    If Err.Number > 0 Then
                        Timer_ME_CAMERA.Enabled = True
                        Exit Sub
                    End If
                On Error GoTo 0
                
                If Err.Number = 0 Then F_SIZE_SET = True
                
                '-----------------------------
                'BIGGER THAN
                '100 * MEG
                '-----------------------------
                'IN THE RIGHT DATE LAST 2 DAYS
                
                '----------------------------------------------------------
                'SOMETIMES ERROR TRAP REQUIRE FILE NOT HERE FOR SOME REASON
                '----------------------------------------------------------
                On Error Resume Next
                'BIGGER THAN
                '100 * MEG
                '90 * MEG
                If F.Size > 90 * (1024 ^ 2) Then
                    If Err.Number > 0 Then Exit Sub
                    
                    If F.DateLastModified > Now - 4 Then
                        FLAG_GO_ME = True
                    End If
                End If
                If Err.Number > 0 Then Exit Sub
                
            Next
        End If
    Next
    
    If FLAG_GO_ME = True Then
     'Me.WindowState = vbNormal
     'I = SetForegroundWindow(Me.hWnd)
     'DoEvents
     'Me.Refresh
     'DoEvents
     
     
     'ABANDON THIS
     'CHANGE PLAN
     'KEEP THIS
    If MUST_BE_SONY_HX60V = True Then
        Load Form3_Me_ViceVersa
    End If
    
    If MUST_BE_FUJI_XP90_ = True Then
        Load Form3_Me_ViceVersa
    End If
        
        'If MsgBox("ME VICE VERSA", vbYesNo Or vbMsgBoxSetForeground) = vbYes Then
 '           Shell VVCOMLINE + " ""D:\VICE_VERSA SCRIPT FILE\viceversa photo camera - me.fsf""", vbNormalFocus
'        End If
    
    End If


'    Shell GetSpecialfolder(38) + "\ViceVersa Pro 2\ViceVersa.exe ""D:\VICE_VERSA SCRIPT FILE\viceversa photo camera - me.fsf"" /autoexec /wait ", vbNormalFocus
'    Shell GetSpecialfolder(38) + "\ViceVersa Pro 2\ViceVersa.exe ""D:\VICE_VERSA SCRIPT FILE\viceversa photo camera - good.fsf"" /autoexec", vbNormalFocus
    
    'MOUSE TO KEYBOARD - GUI - VISUAL STUDIO
    'NOT ANY MOUSE SCROLL WIN 10
    
    
    '10 SEC
    '2 SEC
    
    '-------------------------------
    'RUNNER BOTH HX60V AND FUJI XP90
    '-------------------------------
    TIMER_DEL_THUMB_CAMAREA_MP4.Enabled = True

End Sub



Sub TIMER_DEL_THUMB_CAMAREA_MP4_TIMER()

    If Timer_VICE_VERSA_TIMER_ON.Enabled = False Then Exit Sub

    TIMER_DEL_THUMB_CAMAREA_MP4.Interval = 2000
    
    On Error Resume Next
    Dir2.Path = "W:\MP_ROOT"
    If Err.Number > 0 Then Exit Sub
    
    File2.FileName = "*.THM"
    '--------------------------------------------------
    'DELETE AND LEFT OVER THUMB OF MOVIE LEFT IN CAMERA
    '--------------------------------------------------
    'ONLY WHEN MOVIE NOT ASSOIATED WITH IT
    '--------------------------------------------------
    On Error Resume Next
    For R1 = 0 To Dir2.ListCount - 1
        File2.Path = Dir2.List(R1)
        If File2.ListCount > 0 Then
            For R2 = 0 To File2.ListCount - 1
                Set F1 = FS.GetFile(File2.Path + "\" + File2.List(R2))
                MATCH_FILENAME = Replace(File2.List(R2), ".THM", ".MP4")
                If FS.FileExists(File2.Path + "\" + MATCH_FILENAME) = False Then
                        Kill File2.Path + "\" + File2.List(R2)
                    Else
                        'Stop
                End If
            Next
        End If
    Next
    On Error GoTo 0

End Sub


Sub WAIT_WINDOW_VISIBLE_AFTER_SHELL(WAIT_STRING, HWND_RETURN, TIMEOUT_WAIT_USED)
    
    '## FindWinPart("ViceVersa Pro - [ViceVersa PHOTO CAMERA - GOOD.fsf] - (Registered)")
    'GET GOT THE HWND BEEN DONE
    'WAIT 10 SECOND
    ESCAPEE_COUNT = 20

    Do
        '-------
        'ALWAYS RETURN 0 PROBLEM -- FOR VICEVERSA MAYBE
        '----------------------------
        'SOLVED
        'WORKING -- THIS BIT AT LEAST
        '----------------------------
        i = FindWinPart_VISIBLE(WAIT_STRING, "", HWND_RETURN, VAR1, KILL_SOFTLY, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)
        If VAR1 = 1 And i = HWND_RETURN Then Exit Do
        If VAR1 = 1 And TIMEOUT_WAIT_USED = False Then
            Exit Do
        End If
        
        If UNLOAD_FORM_FLAG = True Then Unload Me: Exit Sub
        
        DoEvents
        Sleep 1000
        ESCAPEE_COUNT = ESCAPEE_COUNT - 1
    Loop Until (VAR1 = 1 And i = HWND_RETURN) Or ESCAPEE_COUNT <= 0
    
    'Debug.Print "TIMEOUT RETRY COUNTDOWN TAKEN FROM 20 SECONDS --" + Str(ESCAPEE_COUNT)
    '-----------------------------------------------------
    
    'End
    Exit Sub
    
    
    'GET GOT THE HWND BEEN DONE
    'WAIT 10 SECOND
    ESCAPEE_COUNT = 100
    Do
        'ALWAYS RETURN 0 PROBLEM -- FOR VICEVERSA MAYBE
        'MAYBE BECUASE HASSLE OF VICE VERSA CHANGE IT HANDLE QUICK AFTER LAUNCH
        
        If IsWindowVisible(HWND_RETURN) = 0 Then Exit Do
        DoEvents
        Sleep 100
        ESCAPEE_COUNT = ESCAPEE_COUNT - 1
    Loop Until (IsWindowVisible(HWND_RETURN) = 0) Or ESCAPEE_COUNT <= 0
    'Debug.Print "IsWindowVisible -- " + Str(ESCAPEE_COUNT)
    '-----------------------------------------------------
    ESCAPEE_COUNT = 100
    Do
        'IS WINDOW SHOWING 1 = TRUE
        If IsWindow(HWND_RETURN) = 1 Then Exit Do
        DoEvents
        Sleep 100
        ESCAPEE_COUNT = ESCAPEE_COUNT - 1
    Loop Until (IsWindow(HWND_RETURN) = 1) Or ESCAPEE_COUNT <= 0
    'Debug.Print "IsWindow -- " + Str(ESCAPEE_COUNT)
    '----------------------------------------------
    ESCAPEE_COUNT = 100
    Do
        'IS IsIconic = 0 ALWAYS FOR NORMAL FORM WINDOW
        If IsIconic(HWND_RETURN) = IsIconic(HWND_RETURN) Then Exit Do
        DoEvents
        Sleep 100
        ESCAPEE_COUNT = ESCAPEE_COUNT - 1
    Loop Until (IsIconic(HWND_RETURN) = IsIconic(HWND_RETURN)) Or ESCAPEE_COUNT <= 0
    'Debug.Print "IsIconic -- " + Str(ESCAPEE_COUNT)
    '----------------------------------------------
    ESCAPEE_COUNT = 100
    Do
        'IS IsZoomed = ALWAYS 0 UNLESS MAXIMISED - THINK MOMENT
        If IsZoomed(HWND_RETURN) = IsZoomed(HWND_RETURN) Then Exit Do
        DoEvents
        Sleep 100
        ESCAPEE_COUNT = ESCAPEE_COUNT - 1
    Loop Until (IsZoomed(HWND_RETURN) = IsZoomed(HWND_RETURN)) Or ESCAPEE_COUNT <= 0
    'Debug.Print "IsZoomed -- " + Str(ESCAPEE_COUNT)

End Sub


Sub DEL_WAVS()
    
    '------------------------------------------------------------------------------
    'DEL WAV HERE IS ONLY CODE OF IT OWN AND __ VBS DO THAT ON ALL COMPUTER NETWORK
    'OR ALL FOLDER OF NETWORK COMPUTER WHEN SYNC OVER
    '------------------------------------------------------------------------------

    '---------------------------------------------------------------------------------------------
    'GOT SOME MOVE TO ARCHIVE BUT NOT ANY DELETER
    '---------------------------------------------------------------------------------------------
    'Exit Sub
    '---------------------------------------------------------------------------------------------
    'MIGHT BE OKAY TO DELETER FOR A BIT AS NEW CODE MADE USE FILE LIST BOX BETTER THAN DIR COMMAND
    'NEW CODE TO DLETER IS VBS BUT SIZE IS ALL IN ONE
    '---------------------------------------------------------------------------------------------

    'THIS THE DEL AFTER TIME
    'CALLED FROM -- SUB Timer_DEL_WAVS_Timer()

    'FILE1.Refresh
    DoEvents
    
    FI2 = FILE1.ListCount - 100
    If FI2 < 0 Then FI2 = 0
    If FI2 >= 0 Then MNU_WAV_ACTIVITY.Visible = True
    
    For R = 0 To FI2 '0 Step -1
        DoEvents
        If UNLOAD_FORM_FLAG = True Then GoTo Exit_Sub22
        
        MIN_SIZE______ = 0
        If Right(FILE1.List(R), 4) = ".WAV" Or Right(FILE1.List(R), 4) = ".MP3" Then
            
            If GetComputerName = "1-ASUS-EEE" Then
                DEL_DATE_DAYS1 = 3
                DEL_DATE_DAYS2 = 180 'ARCHIVE
                MIN_SIZE______ = 300 ' 40000
                GetComputerVAR = True
            End If
            If GetComputerName = "1-ASUS-X5DIJ" Then
                DEL_DATE_DAYS1 = 3
                DEL_DATE_DAYS2 = 180
                MIN_SIZE______ = 300 '40000
                GetComputerVAR = True
            End If
            If GetComputerName = "3-LINDA-PC" Then
                GetComputerVAR = False
                DEL_DATE_DAYS1 = 3
                DEL_DATE_DAYS2 = 180
                MIN_SIZE______ = 1024 '40000
                GetComputerVAR = True
            End If
            If GetComputerName = "4-ASUS-GL522VW" Then
                DEL_DATE_DAYS1 = 3
                DEL_DATE_DAYS2 = 180
                MIN_SIZE______ = 1024 '40000
                GetComputerVAR = True
            End If
            If GetComputerName = "5-ASUS-P2520LA" Then
                DEL_DATE_DAYS1 = 3
                DEL_DATE_DAYS2 = 180
                MIN_SIZE______ = 1024
                GetComputerVAR = True
            End If
            If GetComputerName = "7-ASUS-GL522VW" Then
                DEL_DATE_DAYS1 = 3
                DEL_DATE_DAYS2 = 180
                MIN_SIZE______ = 1024 '40890 '40,856
                GetComputerVAR = True
            End If
            
            '-------------------------------------------------------------------------
            
            '------------------------------------------------------------------
            'THE RADIO RECORD PAIR ARE DIFFERENT SILENT RECORD RUN AT NONE SIZE
            'ASUS-EEE RADIO
            'ACER-ONE RADIO
            '------------------------------------------------------------------
            RADIO_GetComputerVAR = False
            If GetComputerName = "1-ASUS-EEE" Then RADIO_GetComputerVAR = True
            If GetComputerName = "3-LINDA-PC" Then RADIO_GetComputerVAR = True
    
            '-------------------------------------------------------------------------
            DoEvents
            '-------------------------------------------------------------------------
            
            '-------------------------------------------------------------------------
            'DELETE THE FILE BY DATE OF DATE 1 OF 2 EARLYEST IN MAIN FOLDER
            '-------------------------------------------------------------------------
            
            If 1 = 1 Then
            
                On Error Resume Next
                    Set F = FS.GetFile((FILE1.Path + "\" + FILE1.List(R)))
                    If Err.Number > 0 Then
                        GoTo Exit_Sub22
                    End If
                On Error GoTo 0
                
                '-----------------------------
                On Error Resume Next
                    ADATE = F.DateLastModified
                    If Err.Number > 0 Then
                        GoTo Exit_Sub22
                    End If
                On Error GoTo 0
                
                'KEEP A FOUR HOUR LOGG
                'ADJUST AGAIN TO MAKE 4 DAY
                'ADJUST AGAIN TO MAKE 8 DAY
                
                VAR_DONE = False
                
                'DEL_DATE_DAYS1 -- SHORT DATE
                
                If ADATE < Now - DEL_DATE_DAYS1 Then
                    'ONLY THE NORMAL SHORT FILE NAME LENGHT
                    'AS KEEP WHAT IS RENAME LABELED INFO
                    'ABOUT NEIL NEXT DOOR TALK ABOUT HAND ON THE VOLUME KNOB
                    MNU_STR_VAR1 = ""
                    If Len(FILE1.List(R)) <= 23 Then
                        On Error Resume Next
                            
                        '-------------------
                        'COPY TO THE ARCHIVE
                        '-------------------
                        FN1 = FILE1.Path + "\" + FILE1.List(R)
                        FN2 = FILE1.Path + "\# ARCHIVE\" + FILE1.List(R)
                        Err.Clear
                        MNU_STR_VAR1 = "MOVE TO ARCHIVE"
                        FS.MoveFile FN1, FN2
                        'ERR.DESCRIPTION
                        MNU_WAV_ACTIVITY.Caption = MNU_STR_VAR1 + Str(R) + " of" + Str(FILE1.ListCount) + " Counting UP To" + Str(FI2) + " When Stop Error or Problem " + FILE1.List(R) + " " + Err.Description
                        
                        If Err.Description = "File already exists" Then
                            Kill FN1
                            Err.Clear
                        End If
                        
                        
                        If Err.Number > 0 Then
                            GoTo Exit_Sub22
                        End If
                        
                        '-----------------------------
                        'AFTER COPY GOOD AND THEN KILL
                        '-----------------------------
                        'BUT NOT IS A RADIO FILE
                        'SORT THEM OUT WITH A SIZE CHECK
                        '-----------------------------
                        If MNU_STR_VAR1 = "MOVE TO ARCHIVE" Then MNU_STR_VAR2 = "MOVE ARCHIVE && "
                        If RADIO_GetComputerVAR = True Then
                            
                            '---------------------------------
                            'CHECK 2 HOUR AND DELETE IF NAUGHT
                            '---------------------------------
                            
                            If ADATE < Now - TimeSerial(2, 0, 0) And VAR_DONE = False Then
                                '-----------------------------
                                'AND CHECK 0 SIZE OR MAYBE LOW
                                '-----------------------------
                                If F.Size = 0 Then
                                    'TEMP OUT
                                    'MNU_WAV_ACTIVITY.Caption = MNU_STR_VAR2 + "DELETE RADIO MAIN FOLDER 2 DAY && 0 SIZE" + Str(R)
                                    
                                    
                                    'Kill FILE1.Path + "\" + FILE1.List(R)
                                    
                                    
                                End If
                                '-------------------------------------------------
                                'ABORT ON ANY ERROR COME BACK TO IT LATER ON TIMER
                                '-------------------------------------------------
                                If Err.Number > 0 Then
                                    GoTo Exit_Sub22
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next R
    
    
    
    '----------------------------------------------
    'DELETE BY SIZE MOST RECENT
    '----------------------------------------------
    If MIN_SIZE______ > 0 And File4.Path = FILE1.Path Then
        File4.Refresh
        DoEvents
        FI4 = File4.ListCount - 10
        If FI4 < 0 Then FI4 = 0
            For R = FI4 To 0 Step -1
            DoEvents
            If UNLOAD_FORM_FLAG = True Then GoTo Exit_Sub22
            
            If Right(FILE1.List(R), 4) = ".WAV" Or Right(FILE1.List(R), 4) = ".MP3" Then
                On Error Resume Next
                    Set F = FS.GetFile((File4.Path + "\" + File4.List(R)))
                    If Err.Number > 0 Then GoTo Exit_Sub22
                On Error GoTo 0
                
                On Error Resume Next
                    ASIZE = F.Size
                    If Err.Number > 0 Then GoTo Exit_Sub22
                On Error GoTo 0
            
                VAR_DONE = False
                
                If ASIZE <= MIN_SIZE______ Then
                    '-----------------------------
                    'NOT IF DESCRIPTOR ON FILENAME
                    '-----------------------------
                    If Len(File4.List(R)) <= 23 Then
                        On Error Resume Next
                        
                        MNU_WAV_ACTIVITY.Caption = "DELETE IN CURRENT USER ACTIVE FOLDER NOT ARCHIVE IT" + Str(R)
                        '-----------------------------
                        '-----------------------------
                        Err.Clear
                        Kill File4.Path + "\" + FILE1.List(R)
                        If Err.Number > 0 Then GoTo Exit_Sub22
                        
                        On Error GoTo 0
                    End If
                End If
            End If
        Next
    End If
    
    '----------------------------------------------
    'DELETE BY DATE THE ARCHIVE FOLDER
    '----------------------------------------------
    If File4.Path = FILE1.Path + "\# ARCHIVE" Then
        DoEvents
        File4.Refresh
        DoEvents
        FI4 = File4.ListCount - 10
        If FI4 < 0 Then FI4 = 0
        For R = FI4 To 0 Step -1
            
            DoEvents
            If UNLOAD_FORM_FLAG = True Then GoTo Exit_Sub22
            
            If Right(FILE1.List(R), 4) = ".WAV" Or Right(FILE1.List(R), 4) = ".MP3" Then
                
                On Error Resume Next
                    Set F = FS.GetFile((File4.Path + "\" + File4.List(R)))
                    If Err.Number > 0 Then GoTo Exit_Sub22
                On Error GoTo 0
                
                'TROUBLE GETTING THE DATE
                On Error Resume Next
                    ADATE = F.DateLastModified
                    If Err.Number > 0 Then GoTo Exit_Sub22
                On Error GoTo 0
                
                'KEEP A FOUR HOUR LOGG
                'ADJUST AGAIN TO MAKE 4 DAY
                'ADJUST AGAIN TO MAKE 8 DAY
                VAR_DONE = False
                
                'DEL_DATE_DAYS2 -- BIG DATE
                If ADATE < Now - DEL_DATE_DAYS2 Then
                    'ONLY THE NORMAL SHORT FILE NAME LENGHT
                    'AS KEEP WHAT IS RENAME LABELED INFO
                    'ABOUT NEIL NEXT DOOR TALK ABOUT HAND ON THE VOLUME KNOB
                    
                    If Len(File4.List(R)) <= 23 Then
                        On Error Resume Next
                        
                        MNU_WAV_ACTIVITY.Caption = "DELETE IN ARCHIVE FOLDER" + Str(R)
                        '-----------------------------
                        '-----------------------------
                        Err.Clear
                        Kill File4.Path + "\" + FILE1.List(R)
                        If Err.Number > 0 Then GoTo Exit_Sub22
                        
                        On Error GoTo 0
                    End If
                End If
            End If
        Next
    End If
    
Exit_Sub22:
    
MNU_WAV_ACTIVITY_ARCHIVE.Caption = "ARCHIVE FILE COUNT =" + Str(File5_ARCHIVE_KAT_MP3.ListCount) + " .WAV .MP3"
    
If Err.Number > 0 Then
    MNU_WAV_ACTIVITY.Caption = MNU_WAV_ACTIVITY.Caption + " -- WAV DELETE END WITH ERROR"
Else
    'MNU_WAV_ACTIVITY.Visible = False
End If
    
End Sub

Private Sub Timer_DEL_WAVS_Timer()

    If UNLOAD_FORM_FLAG = True Then Unload Me: Exit Sub

    INTERVAL_DEL_WAVS = Hour(Now)
    If INTERVAL_DEL_WAVS = SET_INTERVAL_DEL_WAVS Then
        Exit Sub
    End If
    SET_INTERVAL_DEL_WAVS = INTERVAL_DEL_WAVS
    
    Call RUN_SCRIPT_DEL_WAV_VBS
    
    '------------------------------------------------------------------------------
    'DEL WAV HERE IS ONLY CODE OF IT OWN AND __ VBS DO THAT ON ALL COMPUTER NETWORK
    'OR ALL FOLDER OF NETWORK COMPUTER WHEN SYNC OVER
    '------------------------------------------------------------------------------
    Call DEL_WAVS

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




'---------------------------------------------------------------
'http://www.vbforums.com/showthread.php?673134-RESOLVED-Minimum-height-for-menu-bar-to-be-visible
'-------------- MENU SIZE DECLARE
'Private Type rect
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'Private Type MENUBARINFO
'  cbSize As Long
'  rcBar As rect
'  hMenu As Long
'  hwndMenu As Long
'  fBarFocused As Boolean
'  fFocused As Boolean
'End Type
'Private MenuInfo As MENUBARINFO
'Private Const OBJID_MENU As Long = &HFFFFFFFD
'Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
'Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hwnd As Long, _
'ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
'Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Private Function Menu_Height()
 
    MenuInfo.cbSize = Len(MenuInfo)
    
    If GetMenuBarInfo(Me.hwnd, OBJID_MENU, 0, MenuInfo) Then
   
        With MenuInfo.rcBar
       
'            Debug.Print "Left: " & CStr(.Left)
'            Debug.Print "Right: " & CStr(.Right)
'            Debug.Print "Top: " & CStr(.Top)
'            Debug.Print "Bottom: " & CStr(.Bottom)
            'Menu_Height = CStr(.Top) + CStr(.Bottom)
            Menu_Height = CStr(.Bottom) - CStr(.Top)
            'Debug.Print Menu_Height

        End With
       
    End If
   
End Function
'------------------------------------------

Private Sub Timer_Drives_Timer()
    If UNLOAD_FORM_FLAG = True Then Unload Me: Exit Sub

    i = GET_DRIVES_FOR_CAMERA
End Sub

Function GET_DRIVES_FOR_CAMERA()

'-----------------------------
'SOMETIME ERROR READ DRIVE BUG
'-----------------------------
'---------------------
On Error GoTo END_EXIT
'---------------------
'--------------------------------------------
'NEED SOMETHING TO TELL IF NEW USB BEEN ADDED
'SAVE EVENT TIMER
'--------------------------------------------
'-----------------------------------------------------
'DON'T REQUIRE POLL IF CONNECTED - ONLY USE DIR(EXIST)
'-----------------------------------------------------
'MAYBE USE THAT THAN DRIVE POLL
'------------------------------

VAR_SHELL_SUBST_BEEN_SET_FLAG = False

Dim DC
Set DC = FS.Drives
For Each D In DC
    S = S & D.DriveLetter
    'If d.DriveType <> 2 Then 'FIXED DRIVE
    'If d.ISREADY Then
        'Debug.Print d.DriveLetter + " -- " + d.volumename + " -- " + Str(d.drivetype)
    'End If
    
    If D.DriveType = 1 Then 'CAMERA SD MEM CARD
        'Stop
        'n = d.ShareName
        'Else
        
        If D.IsReady Then
            N = N + D.DriveLetter 'd.VolumeName
            
            VOLUME_NAME_CAMERA_1 = "##"
            VOLUME_NAME_CAMERA_2 = "FUJI XP90"
            VOLUME_NAME_CAMERA_3 = "MC - HX60V"
            
            CAMERA_TO_GO = False
            If D.VolumeName = VOLUME_NAME_CAMERA_1 Then CAMERA_TO_GO = True
            If D.VolumeName = VOLUME_NAME_CAMERA_2 Then CAMERA_TO_GO = True
            If D.VolumeName = VOLUME_NAME_CAMERA_3 Then CAMERA_TO_GO = True
            
            If CAMERA_TO_GO = True Then
                VAR_SHELL_SUBST_BEEN_SET_FLAG = True
                VAR_SHELL_SUBST_BEEN_SET_FLAG2 = True
                
                If VAR_SHELL_SUBST_BEEN_SET = False Then
                    'Debug.Print d.DriveLetter + " -- " + d.volumename + " -- " + Str(d.drivetype)
                    
                    
                    If MNU_W_DRIVE_OFF.Caption <> "W DRIVE OFF && GO" Then
                                        
                        Shell "SUBST W: " + D.DriveLetter + ":\", vbHide
                    
                    End If
                    
                    
                    VAR_SHELL_SUBST_BEEN_SET = True
                    '------------------------------------------------------------------------
                    ' THESE ONE TIMER ROUTINE HERE ARE NOT ANYTHING TO DO WITH A DELETER FEATURE
                    ' USER A RENAME
                    ' LAUNCH VICE VERSA IS VIDEO
                    '------------------------------------------------------------------------
                    TIMER_DEL_CAMERA_EMPTY_MP4_MOV.Enabled = True
                    '------------------------------------------------------------------------
                    TIMER_SET_CAMERA_UP_TO_MARK_FOLDER.Enabled = True
                    MNU_W_DRIVE.Enabled = True
                    MNU_VICE_VERSA_ME.Enabled = True
                    MNU_VICE_VERSA_LOCAL.Enabled = True
                    Form1.MNU_VICE_VERSA_ME.Enabled = True

                    
                    Exit For
                End If
                Exit For
            End If
        End If
    End If
Next

GET_DRIVES = N
GET_DRIVES_RESULT = N
GET_DRIVES_FOR_CAMERA = N

If VAR_SHELL_SUBST_BEEN_SET_FLAG2 = True Then
    If VAR_SHELL_SUBST_BEEN_SET_FLAG = False Then
        '------------------------------------------
        TIMER_DEL_THUMB_CAMAREA_MP4.Enabled = False
        '------------------------------------------
        Shell "SUBST W: /D", vbHide
        VAR_SHELL_SUBST_BEEN_SET = False
        'I = GET_DRIVES_FOR_CAMERA
        MNU_W_DRIVE.Enabled = False
        MNU_VICE_VERSA_ME.Enabled = False
        MNU_VICE_VERSA_LOCAL.Enabled = False
        '------------------------------------------
        Form1.MNU_VICE_VERSA_ME.Enabled = False
        Timer_VICE_VERSA_ME.Enabled = False

    End If
End If

Set DC = Nothing

END_EXIT:

End Function

Private Sub Timer_VICE_VERSA_ME_Timer()
    
    Timer_VICE_VERSA_ME.Interval = 55000
    Form1.MNU_VICE_VERSA_ME.Visible = False
    Timer_VICE_VERSA_ME.Enabled = False

End Sub

Private Sub Timer2_Timer()

'file5.Path =



End Sub




Private Sub zzCheckTimer_Timer()

Dim CHECKQ
On Error Resume Next
'Debug.Print FORM_ME.Name

For Each Control In FORM_ME.Controls
    CHECK_A1 = 0
    CHECK_A1 = Control.Value
    CHECK_B1 = 0
    CHECK_B1 = Control.Checked
 '   Debug.Print Control.Name + Str(CHECK_A1) + Str(Val(CHECK_B1))
    CHECKQ = CHECKQ + Str(CHECK_A1) + Str(Val(CHECK_B1))
Next

'Debug.Print checkq + vbCrLf + OCheckQ

If CHECKQ = OCheckQ Then Exit Sub

OCheckQ = CHECKQ

'Debug.Print OCheckQ

Call Frm_Save_Load_Checks.zzSave_Checks

End Sub



Private Sub WebCam_Matt_Timer_Timer()

If ComputerName <> "1-ASUS-EEE" Then Exit Sub

COUNT_SEC_WEBCAM_MATT = COUNT_SEC_WEBCAM_MATT + 1
If COUNT_SEC_WEBCAM_MATT < 120 And COUNT_SEC_WEBCAM_MATT_FLAG = True Then Exit Sub
If FindWindow(vbNullString, "WebCam Motion Capture") > 0 Then
    COUNT_SEC_WEBCAM_MATT = 0
End If

If FindWindow(vbNullString, "WebCam Motion Capture") = 0 Then
    COUNT_SEC_WEBCAM_MATT_FLAG = True
    COUNT_SEC_WEBCAM_MATT = 0
    
    Shell "D:\VB6\VB-NT\00_Best_VB_04\Webcam_Motion\WebCamMatts.exe", vbNormalNoFocus
    
End If

End Sub

'Private Function GetWindowTitle(ByVal hwnd As Long) As String
'   Dim L As Long
'   Dim S As String
'   L = GetWindowTextLength(hwnd)
'   S = Space(L + 1)
'   GetWindowText hwnd, S, L + 1
'   GetWindowTitle = Left$(S, L)
'End Function

Function FindWinPart(FIND_STRING) As Long


Dim ASH$
Dim Test_HWND As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME
'an another one

'--------------------------------------------
Test_HWND = FindWindowDLL(ByVal 0&, ByVal 0&)
'----------------------------------------
'THIS ONE IS NORM USE
'BUT MODIFIED TO ONE ABOVE BY ME MATT.LAN@
'----------------------------------------
'IT DONT PRODUCE A RESULT
'Test_HWND = FindWindow(ByVal 0&, ByVal 0&)
'--------------------------------------------

FindWinPart = False
    
Do While Test_HWND <> 0
    
    
    ASH$ = GetWindowTitle(Test_HWND)
    
    'cText = Space$(255)
    'ghj$ = GetClassName(Test_HWND, cText, 255)
    If ASH$ <> "" Then
        If InStr("-" + ASH$, FIND_STRING) > 0 Then
            FindWinPart = Test_HWND
            Exit Do
        End If
    End If
    
    'retrieve the next window
    Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)
Loop


End Function



Function FindWinPart_VISIBLE(FIND_STRING, CLASS_TEXT, TARGET_HWND_CHECK, VAR1, KILL_SOFTLY, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY) As Long

'----------------
'FINDWINDOWPART
'FIND_WINDOW_PART
'----------------
Dim ASH$
Dim Test_HWND As Long, _
    test_pid As Long, _
    test_thread_id As Long
Dim cText As String

Dim Test_HWND2 As Long



'-----------------------------
VAR1 = 0
'-----------------------------
HOW_MANY_INSTANCE_SHOWING = -1
'-----------------------------
MATCH_HWND = False
'-----------------------------
HOW_MANY_INSTANCE_COUNT = 0
'-----------------------------
FindWinPart_VISIBLE = False
'-----------------------------
GetForegroundWindow_VAR = False
'-----------------------------
    


'Find the first window

'need this is you gona use this procedure get from CIDRun ME
'an another one

'------------------------------
'FOR THER USE SEARCH VICE_VERSA
'BEGIN WHERE
'TARGET_HWND_CHECK
'AND MIGHT FIND NEXT LOWER
'------------------------------
'Test_HWND = TARGET_HWND_CHECK
'-------------------------------------------------
'THIS NOT WORK BECAUSE RESULT WANT COME FROM LOWER
'AND THEN NEED BEGIN AT 0 LOOK AT NEXT LINE
'-------------------------------------------------

'--------------------------------------------
Test_HWND = FindWindowDLL(ByVal 0&, ByVal 0&)
'----------------------------------------
'THIS ONE IS NORM USE
'BUT MODIFIED TO ONE ABOVE BY ME MATT.LAN@
'----------------------------------------
'IT DONT PRODUCE A RESULT
'Test_HWND = FindWindow(ByVal 0&, ByVal 0&)
'--------------------------------------------

'--------------------------------------------
Do While Test_HWND <> 0
    
    
    ASH$ = GetWindowTitle(Test_HWND)
    
    cText = Space$(255)
    CLASS_V1 = GetClassName(Test_HWND, cText, 255)
    'TRIM THE NULLSTRING IS YOU WANTED
    
    CLASS_FLAG = True
    '---------------
    'GOT SOME CLASS_TEXT - PASSED
    '---------------
    If CLASS_TEXT <> "" Then
    '---------------
    'IS A MATCH TRUE
    '---------------
        If InStr("-" + cText, CLASS_TEXT) > 0 Then
            CLASS_FLAG = True
        Else
            CLASS_FLAG = False
        End If
    End If
    
    
    If ASH$ <> "" And CLASS_FLAG = True Then
        If InStr("-" + ASH$, FIND_STRING) > 0 Then
            HWND_COUNT = HWND_COUNT + 1
            
            VAR1 = IsWindow(Test_HWND)
'            Debug.Print Format(HWND_COUNT, "000") + " -- HWND FROM FIND_WIN_PART =" + Str(HWND_RETURN)
'            Debug.Print Format(HWND_COUNT, "000") + " -- " + Ash$
'            Debug.Print Format(HWND_COUNT, "000") + " -- IS_VISIBLE" + Str(VAR1)
            
            '----------------------------------------------
            'Test_HWND IS LOWER THAN TARGET_HWND_WHEN CHECK
            '----------------------------------------------
            If Test_HWND < TARGET_HWND_CHECK Then
                'Debug.Print Format(HWND_COUNT, "000") + " -- Test_HWND < TARGET_HWND_CHECK --" + Str(Test_HWND) + " -- " + Str(TARGET_HWND_CHECK) + " -- " + Str(TARGET_HWND_CHECK - Test_HWND) + " -- PROPER"
            End If
            '----------------------------------------------
            If Test_HWND > TARGET_HWND_CHECK Then
                'Debug.Print Format(HWND_COUNT, "000") + " -- Test_HWND > TARGET_HWND_CHECK --" + Str(Test_HWND) + " -- " + Str(TARGET_HWND_CHECK)
            End If
            '----------------------------------------------
            If TARGET_HWND_CHECK = Test_HWND Then
                'Debug.Print Format(HWND_COUNT, "000") + " -- FIND WINODW PART GOT MATCH HWND --" + Str(Test_HWND)
            End If
            '----------------------------------------------
            
'            If VAR1 = 1 And TARGET_HWND_CHECK = Test_HWND Then
'                Debug.Print Format(HWND_COUNT, "000") + " -- FIND WINODW PART GOT MATCH HWND --" + Str(Test_HWND)
'                FindWinPart_VISIBLE = Test_HWND
'                Exit Do
'            End If
            
            If VAR1 = 1 Then
                'Exit Do
            End If
        
            If HOW_MANY_INSTANCE_SHOWING = -1 And VAR1 = 1 Then
                HOW_MANY_INSTANCE_COUNT = HOW_MANY_INSTANCE_COUNT + 1
                '--------------------------------------
                If GetForegroundWindow = Test_HWND Then
                    GetForegroundWindow_VAR = True
                End If
                '--------------------------------------
            
            If KILL_SOFTLY > 0 Then
                If HOW_MANY_INSTANCE_COUNT > KILL_SOFTLY Then
                
                    RESULT = PostMessage(Test_HWND, WM_CLOSE, 0&, 0&)
                    
            
                End If
            End If
            
            End If
            
            '------------------------ -----
            'CHECKING FOR TARGET HWND MATCH
            '------------------------------
            'BETTER TO FIND EXACT TARGET WHEN WAIT FOR SHELL TO LOAD
            '-------------------------------------------------------
            If TARGET_HWND_CHECK = Test_HWND Then
'                HOW_MANY_INSTANCE_COUNT = 1
'                HOW_MANY_INSTANCE_SHOWING = HOW_MANY_INSTANCE_COUNT
                FindWinPart_VISIBLE = Test_HWND
            '-------------------------------------------------------
                MATCH_HWND = True
            '-------------------------------------------------------
                Exit Do
            End If
        
            'WANT FIRST OCCURANCE ONLY THEN HERE
            If ONE_INSTANCE_ONLY = True Then
                '-------------------------------
                'HOW_MANY_INSTANCE_COUNT = HOW_MANY_INSTANCE_COUNT + 1
                '-------------------------------
                ONE_INSTANCE_ONLY = False
                FindWinPart_VISIBLE = Test_HWND
                Exit Do
                '-------------------------------
            Else
                'TAKE IT TO THE END ALL OF THEM IF MULTI INSTANCE
                '------------------------------------------------
                FindWinPart_VISIBLE = Test_HWND
            End If
        End If
    End If
    
    'retrieve the next window
    Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)
Loop

If HOW_MANY_INSTANCE_SHOWING = -1 Then
    HOW_MANY_INSTANCE_SHOWING = HOW_MANY_INSTANCE_COUNT
End If

If Test_HWND = TARGET_HWND_CHECK Then Exit Function


'Exit Function

'FIND PARENT OF HWND FOR SHELL PID TO HWND COMPARE
'Test_HWND


'----------------
'ON GOING WORK
'----------------
'RESULT SAME AS ENTRY MUST BE MAIN PARENT
i = GET_PARENT(Test_HWND)
'----------------------------------------
'I = GET_CHILD_TEST(Test_HWND)
'----------------------------------------
'NOT ANY RESULT WE WANT HERE

KILL_SOFTLY = 0

End Function
'GET_CHILD_TEST
Public Function GET_CHILD_TEST(ByVal TargetHwnd As Long) As String

    Dim i As Long
    Dim j As Long
    Dim COUNTER5
    i = TargetHwnd
    Dim cText As String

    Do While i <> 0
        j = i
        i = GetWindow(i, GW_CHILD)
        'WHY EXIT HERE NOT FULL TEST RUN
        'If I = j Then Exit Do
        
        'DoEvents
'        Debug.Print I
        res = 0
        If GetWindowTitle(i) = "Help" Then res = 1
        'Debug.Print GetWindowTitle(I)
        cText = Space$(255)
        '--------------------------------------------
        'GET CLASS NAME IS RETURN RESULT IN PARAMETER
        'RETURN FROM FUNCTION IS A NUMERIC
        '--------------------------------------------
        CLASS_RESULT_NUMERIC = GetClassName(i, cText, 255)
        If res = 1 And InStr(cText, "Button") > 0 Then
        res = 2: Stop
        End If
        
        'cText = Space$(255)
        'ghj$ = GetClassName(Test_HWND, cText, 255)
        
        If res = 2 Then MsgBox "result"
        
        
        If i = TargetHwnd Then
            'Debug.Print "GET_CHILD_TEST --" + Str(TargetHwnd)
        End If
    Loop
    i = j
            
GET_CHILD_TEST = i


End Function

Public Function GET_PARENT(ByVal TargetHwnd As Long) As String

    Dim i As Long
    Dim j As Long
    Dim COUNTER5
    i = TargetHwnd
    
    If TargetHwnd Then
        Do While i <> 0
            j = i
            COUNTER5 = COUNTER5 + 1
            i = GetParent(i)
        Loop
        i = j
    
    End If
    
GET_PARENT = i
'RESULT SAME AS ENTRY MUST BE MAIN PARENT
'WHEN VICE VERSA SHELL HWND FROM PROCESS
End Function




Function GET_DRIVES_FIND_FOLDER(FOLDER As String)

'-------------------------
'WANT
'-------------------------
'FOLDER_SENDTO_FAT32_STORE
'-------------------------

GET_DRIVES_FIND_FOLDER = ""

Dim DC
Set DC = FS.Drives
For Each D In DC
  S = S & D.DriveLetter
  If D.DriveType <> 2 Then 'FIXED DRIVE
    'Stop
    'n = d.ShareName
  ElseIf D.IsReady Then
    'N = N + d.DriveLetter 'd.VolumeName
    If FS.FolderExists(D.DriveLetter + ":" + FOLDER) = True Then
        GET_DRIVES_FIND_FOLDER = D.DriveLetter + ":" + FOLDER
        Exit Function
    End If
  End If
  's = s & n & "<BR>"
Next

'GET_DRIVES = N
'GET_DRIVES_RESULT = N

End Function



Public Sub FindCursor(x, y)

Dim P As POINTAPI

GetCursorPos P
'   return x and y co-ordinate
x = P.x ' / GetSystemMetrics(0) * Screen.Width
'   for current cursor position
y = P.y '/ GetSystemMetrics(1) * Screen.Height

End Sub

Sub WinonPoint()

Exit Sub

'Get the window under the cursor
If IsIDE = False Then Exit Sub
FindCursor Nx, Ny
mWnd = WindowFromPoint(Nx, Ny)
WW$ = GetWindowTitle(mWnd)
XX = 0
If InStr(WW$, "Extreme Lock Build: 2011") > 0 Then XX = 1
If InStr(WW$, "Tidal_Info") > 0 Then XX = 1
If InStr(WW$, "Tidal Info") > 0 Then XX = 1
If XX = 1 And InStr(WW$, "(Code)") > 0 Then XX = 0

'if not in ide then load but not more than once
If TestKeyLoggOff = False And TestKeyLoggOff <> -2 Then
    Call DSkeybd_F.Command1_Click
    TestKeyLoggOff = -2
End If


'TestKeyLoggOff = True
'TestKeyLoggOff = False
If TestKeyLoggOff = True Then
    If XX = 1 And DSkeybd_F.IsHook = False Then
        Call DSkeybd_F.Command1_Click
    End If
    If XX = 0 And DSkeybd_F.IsHook = True Then
        Call DSkeybd_F.UnHookCommand_Click
    End If
End If

A = Me.hwnd

End Sub


Private Sub TIMER_MouseDetectMove_TIMER()

FindCursor Nx, Ny

If (Nx <> ONx Or Ny <> ONy) Then

    MouseDetectMove_FOR_EXPLORER_RESET = Now + TimeSerial(1, 0, 0)

End If

ONx = Nx
ONy = Ny

Exit Sub

'This Will Happen to Mouse If User Is Logged off
'If Nx = 0 And Ny = 0 Then
'    LockSSaver = Now + LockSaverDelay
'    'Winamp Video
'    Call SetLockMax
'    Call ProgressLock
'End If

'If Nx = 0 Or Ny = 0 Then Exit Sub


If Quake = 0 Then
    If (Nx <> Xmp5 Or Ny <> Ymp5) And (Nx <> ScreenWidth And Ny <> ScreenHeight) Then
        
Call WinonPoint
        
        
        'If Sluty3 = 0 Then StandBy = Now + TimeSerial(0, 0, 3)

'        If Tim < Now Then
 '           List1.AddItem Format$(Now, "HH:MM:SS") + Str(nx) + " -- " + Str(ny)
  '          List1.ListIndex = List1.ListCount - 1
   '         Tim = Now + TimeSerial(0, 0, 1)
    '    End If

        Mousey = Mousey + Sqr(Abs(Xmp5 - Nx) ^ 2 + Abs(Ymp5 - Ny) ^ 2)
        
        'Mousey2 = Sqr(Abs(Xmp5 - Nx) ^ 2 + Abs(Ymp5 - Ny) ^ 2)
        'If Mousey2 > 15 Then Mouse10CLicks = True Else Mouse10CLicks = False
        
        MouseClicks = Sqr(Abs(Xmp5 - Nx) ^ 2 + Abs(Ymp5 - Ny) ^ 2)
                    
        If MouseFirstRun = 0 Then
            MouseFirstRun = 1
            Mousey = 0
        End If
        
        If MouseClicks > 3 Then
            Label21.Caption = Str$(Mousey)
        End If
        
        If Easyride > Now Then Label23.BackColor = Wtoolcol23

        Startmouse = 1
        If DelayXYUpdates22 > Now Then DelayXYUpdates22 = Now - 1
    
        Xmp5 = Nx: Ymp5 = Ny
    
    End If
End If

If (Quake = 1 And (Nx <> ScreenWidth And Ny <> ScreenHeight)) Or (Easyride < Now And Quake = 1) Then
    If drive2$ <> "" And Hidecursor2 = False Then
        If CompName$ <> "MEACHELLE12345" Then
            SetCursorPos Xmp6, Ymp6
            Xmp5 = Xmp6: Ymp5 = Ymp6
            Quake = 0
            Easyride = Now - 1
        End If
    End If
End If

If Easyride < Now Then
    If Label22.BackColor = &HFFC0C0 Then Label22.BackColor = &HFF0000
    If Label23.BackColor = &HFFC0C0 Then
        Label23.BackColor = Wtoolcol23
    End If
End If

End Sub



Private Sub Timer_EXPLORER_GONE_Timer()


'If GETWinNT_Ver <> "WIN XP" Then
    'Timer_EXPLORER_GONE.Enabled = False
    'Exit Sub
'End If


'1st FIT IN ANOTHER SUBROUTINE
'Call MNU_CLIPBOARD_API_PUBLIC_VAR_HOOK_Click


'-------------------
'If Explorer Crashes
'-------------------

'----------------
'EXPLORER DESKTOP
'----------------
If FindWindow("Progman", "Program Manager") = 0 Then

    ExplorerGone = True
    
    Me.WindowState = Normal
    DoEvents
    Me.Refresh
    DoEvents
    Me.SetFocus
    DoEvents
    Me.Refresh
    DoEvents
    
    
    'ONLY REQUIRE WIN XP
    'FORM_STAY_UP_WITH_MSGBOX = True ___ NOT USE THIS PROGRAM CODE
    'I = MsgBox("Do You Want to Reload Explorer, Crash -- Disappeared", vbYesNo + vbMsgBoxSetForeground)
    'FORM_STAY_UP_WITH_MSGBOX = False

'    Me.WindowState = Normal

    i = vbYes

    If i = vbYes Then
        'Shell "c:\windows\Explorer.exe", vbNormalFocus
        
        'cmdLine$ = "c:\windows\Explorer.exe"
        'i = ExecCmd(cmdLine$)
    
        Shell "c:\windows\Explorer.exe", vbNormalFocus
    
        Do
            I2 = FindWindow("Progman", "Program Manager")
            DoEvents
            'Sleep 500
        Loop Until I2 <> 0
        A = A
    End If
    
    Exit Sub

End If

If FindWindow("Progman", "Program Manager") <> 0 And ExplorerGone = True Then

    'BRING WINDOWS FRONT
    'I = FindWinPartExplorerGone(False) ' True = Quite Mode Don't Display  Result
    i = FindWinPartExplorerGone(True) ' True = Quite Mode Don't Display  Result
'    Debug.Print str(i) + " Windows Brought Forward"
 
    'MNU_BRing_Front.Caption = "Bring All Windows Front -- Explorer Crash/Terminated @ " + Format(Now, "DD-MMM-YYYY HH:MM:SS")
    ExplorerGone = False

End If



End Sub





Function FindWinPartExplorerGone(Q) As Long
'AND BRING ALL WINDOWS FORWARD
'ONLY VB SUFFERS FROM THIS

FindWinPartExplorerGone = False

Dim RECT8 As RECT



Dim ASH$
Dim Test_HWND As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one

Dim Huge

Test_HWND = FindWindowDLL(ByVal 0&, ByVal 0&)
BR$ = ""
Do While Test_HWND <> 0
    
    
    DoEvents
    hwnd9 = GetWindowRect(Test_HWND, RECT8)
        ASH$ = GetWindowTitle(Test_HWND)
'        If InStr(Ash$, "Double") > 0 Then Stop
        gws = GetWindowState(Test_HWND)
'        If gws = -1 Then gws2$ = "Window State Normal"
'        If gws = vbMinimized Then gws2$ = "Window State Minimized"
'        If gws = vbMaximized Then gws2$ = "Window State Maximized"
'
'
    If (RECT8.Top > 0 And RECT8.Left > 0) Or gws = vbMinimized Then
        ASH$ = GetWindowTitle(Test_HWND)
        If ASH$ <> "" And InStr(BR$, "-- " + ASH$ + " -- ") = 0 Then
            BR$ = BR$ + "-- " + ASH$ + " -- "
'            On Error Resume Next
'            AppActivate Ash$, 0
'            On Error GoTo 0
            Huge = Huge + 1
            ef = SetForegroundWindow(hwnd9)
            'ef = Putfocus(hwnd9)
            ECUTE = Timer + 0.1
         
            Do
            
                DoEvents
            
                'Safety IN CASE TME RESETS BACK TO ZERO
                If Timer < ECUTE - 30 Then Exit Do
                
            Loop Until Timer > ECUTE
        
        End If
    End If
        
'retrieve the next window
Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)

Loop

If Q = False Then MsgBox Str(Huge) + " Windows Brought To Front", vbMsgBoxSetForeground

FindWinPartExplorerGone = Huge

End Function



