VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "KAT MP3 RELOAD"
   ClientHeight    =   5930
   ClientLeft      =   590
   ClientTop       =   3510
   ClientWidth     =   12070
   LinkTopic       =   "Form1"
   ScaleHeight     =   5930
   ScaleWidth      =   12070
   Begin VB.Timer Timer_Find_Window_SPEED 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5880
      Top             =   600
   End
   Begin VB.Timer Timer_KAT_MOUSE_FOR_VB 
      Interval        =   100
      Left            =   7200
      Top             =   840
   End
   Begin VB.FileListBox File3 
      Height          =   860
      Left            =   9840
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer TIMER_RETRY_WRITE_INFO_UNTIL_DONE2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   1320
   End
   Begin VB.Timer TIMER_RETRY_WRITE_INFO_UNTIL_DONE1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   840
   End
   Begin VB.Timer WebCam_Matt_Timer 
      Interval        =   1000
      Left            =   8385
      Top             =   3210
   End
   Begin VB.Timer TIMER_DEL_THUMB_CAMAREA_MP4 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6120
      Top             =   1560
   End
   Begin VB.Timer Timer_ME_CAMERA 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5520
      Top             =   3000
   End
   Begin VB.FileListBox File2 
      Height          =   260
      Left            =   8280
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer zzCheckTimer 
      Enabled         =   0   'False
      Interval        =   59000
      Left            =   5520
      Top             =   2520
   End
   Begin VB.Timer Timer_Find_Window 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5520
      Top             =   600
   End
   Begin VB.Timer Timer_DESKTOP_INI_DEL 
      Interval        =   1000
      Left            =   5520
      Top             =   1560
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   0
      TabIndex        =   3
      Top             =   3480
      Width           =   12132
   End
   Begin VB.Timer TIMER_SET_CAMERA_UP_TO_MARK_FOLDER 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4920
      Top             =   2520
   End
   Begin VB.DirListBox Dir2 
      Height          =   1440
      Left            =   8160
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer_W32TM 
      Interval        =   1000
      Left            =   3915
      Top             =   3870
   End
   Begin VB.Timer TIMER_DEL_CAMERA_EMPTY_MP4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4920
      Top             =   2040
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   6480
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer_Drives 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4920
      Top             =   1560
   End
   Begin VB.Timer Timer_DEL_WAVS 
      Interval        =   1000
      Left            =   3840
      Top             =   1560
   End
   Begin VB.Timer Timer_WINMERGE_KILL 
      Interval        =   1000
      Left            =   4320
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   1080
   End
   Begin VB.FileListBox FILE1 
      Height          =   2860
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   12090
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
      Width           =   12045
   End
   Begin VB.Menu MNU_ENABLE_FINDWINOW_TIMER 
      Caption         =   "* FIND WINDOW AND FOLDER EXPLOERER *"
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
      Begin VB.Menu MNU_SEND_TO_FAT32_FOLDER 
         Caption         =   "SEND TO - Fat32 Folder"
      End
      Begin VB.Menu MNU_SEND_TO_SYSTEM_FOLDER 
         Caption         =   "SEND TO - SYSTEM"
      End
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "* EXIT *"
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "* VB ME *"
   End
   Begin VB.Menu MNU_RUN 
      Caption         =   "* RUN LOAD *"
      Begin VB.Menu MNU_VICEVERSA_W_SUBST 
         Caption         =   "VICE VERSA CAMERA - SETUP W DRIVE - SUBST"
      End
      Begin VB.Menu MNU_VICEVERSA_EXPLORER 
         Caption         =   "VICE VERSA Explorer FOLDER W DRIVE"
      End
      Begin VB.Menu MNU_VICEVERSA_REMOTE 
         Caption         =   "VICE VERSA SCRIPT - CAMERA REMOTE"
      End
      Begin VB.Menu MNU_VICEVERSA_LOCAL 
         Caption         =   "VICE VERSA SCRIPT - CAMERA LOCAL"
      End
   End
   Begin VB.Menu MNU_TIMER_SINCE_ACTIVITY 
      Caption         =   "TIMER SINCE ACTIVITY"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'KAT MOUSE PROJECT
'---------------------------------------------
'SEEM THE KATMOUSE UNLOAD CAN ALSO STEAL FOCUS
'THAT NEED PUT BACK
'
'SOME PULL DOWN MENU REQUIRE KATMOUSE TO STAY
'OR LOAD TIME INVOLED IN RETURN TO FORM
'---------------------------------------------

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

'-------
'MEMORY
'HISTORY
'-------
'PAST
'PRESENT
'FURTURE
'-------
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
Dim Windows_Security As RECT
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
Dim XVB_DATE As Date

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

Dim WW

O_WindowState = -8


WW = "D:\KAT MP3 RECORDER\KAT MP3 "
If GetComputerName = "1-ASUS-EEE" Then PATHSET = WW + GetComputerName
If GetComputerName = "3-LINDA-PC" Then PATHSET = WW + GetComputerName
If GetComputerName = "5-ASUS-P2520LA" Then PATHSET = WW + GetComputerName

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
'D:\#0 1 INSTALLATIONS\Installation programs\# 00 New Install Progs\# Installed Now\#01 Matthew's\KAT MP3 RECORDER\kat-mp3-recorder-setup-free.exe




Set FS = CreateObject("Scripting.FileSystemObject")

If IsIDE = False Then Me.WindowState = vbMinimized
If IsIDE = True Then Me.WindowState = vbMinimized
If IsIDE = True Then Me.WindowState = vbNormal

X2 = 100

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

If FS.FILEEXISTS(GetSpecialfolder(16) + "\DESKTOP.INI") = True Then
    FS.DELETEFILE (GetSpecialfolder(16) + "\DESKTOP.INI")
End If
If FS.FILEEXISTS(GetSpecialfolder(25) + "\DESKTOP.INI") = True Then
    FS.DELETEFILE (GetSpecialfolder(25) + "\DESKTOP.INI")
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
    If Me.WindowState = vbMinimized Then Exit Sub
'    REM OKAY THIS CODE FOR HERE MAX AS CODE IS ONLY INNER FORM
'    If Me.WindowState = vbMaximized Then Exit Sub
End If
O_WindowState = Me.WindowState


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

Me.Height = ((VAR_BOTTOM_OBJECT + (Menu_Height2 * Screen.TwipsPerPixelY)) + 490)

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

Private Sub MNU_ENABLE_FINDWINOW_TIMER_VAR_Click()

MNU_ENABLE_FINDWINOW_TIMER_VAR.Checked = Not MNU_ENABLE_FINDWINOW_TIMER_VAR.Checked

Timer_Find_Window.Enabled = MNU_ENABLE_FINDWINOW_TIMER_VAR.Checked
'Timer_Find_Window
End Sub

Private Sub MNU_EXIT_Click()
Unload Me
End Sub

Private Sub MNU_KAT_MP3_FOLDER_Click()

    PID = Shell("EXPLORER " + PATHSET, vbNormalNoFocus)
    

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
End Sub

Private Sub MNU_VB_ME_Click()

'GetSpecialfolder(38) OR
'GetSpecialfolder(42)

'RUN PROGRAM AS ADMINISTRATOR IN WIN 10
'RIGHT CLICK -- PROPERTIES -- COMPATIABLE AND SET THERE TO ALWAYS
'OR PROBLEM WON'T DO SHELL PROPER


'MsgBox """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbMsgBoxSetForeground
'Shell GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE", vbNormalFocus

Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
EXIT_TRUE = True
Unload Me


End Sub

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


Private Sub MNU_VICEVERSA_EXPLORER_Click()

        If Dir(CAMERA_EXPLORER_PATH, vbDirectory) = "" Then
            MsgBox "W DRIVE FOLDER " + CAMERA_EXPLORER_PATH + " -- DOES NOT EXIST - ABANDON LOAD EXPLORER"
            Exit Sub
        End If
        
        Shell "EXPLORER ""W:\""", vbNormalFocus


End Sub

Private Sub MNU_VICEVERSA_LOCAL_Click()

    Call FIND_VICE_VERSA
    
    VICE_VERSA_SCRIPT_FILE = "D:\VICE_VERSA SCRIPT FILE\ViceVersa PHOTO CAMERA - LOCAL.fsf"
    If Dir(VICE_VERSA_SCRIPT_FILE) = "" Then
    
        MsgBox "UNABLE TO FIND" + vbCrLf + VICE_VERSA_SCRIPT_FILE + vbCrLf, vbMsgBoxSetForeground
        Exit Sub
    
    End If
    
    PID = Shell(VVCOMLINE + " """ + VICE_VERSA_SCRIPT_FILE + """", vbNormalFocus)

End Sub

Private Sub MNU_VICEVERSA_REMOTE_Click()

    Call FIND_VICE_VERSA
    
    VICE_VERSA_SCRIPT_FILE = "D:\VICE_VERSA SCRIPT FILE\ViceVersa PHOTO CAMERA - REMOTE.fsf"
    If Dir(VICE_VERSA_SCRIPT_FILE) = "" Then
    
        MsgBox "UNABLE TO FIND" + vbCrLf + VICE_VERSA_SCRIPT_FILE + vbCrLf, vbMsgBoxSetForeground
        Exit Sub
    
    End If
    
    
    PID = Shell(VVCOMLINE + " """ + VICE_VERSA_SCRIPT_FILE + """", vbNormalFocus)

End Sub

Private Sub MNU_VICEVERSA_W_SUBST_Click()
    
    'I = GET_DRIVES_FOR_CAMERA
    If VAR_SHELL_SUBST_BEEN_SET = False Then
        MsgBox "THE W DRIVE SUBST IS NOT MADE - MAYBE UNABLE TO FIND CAMERA VOLUME NAME", vbMsgBoxSetForeground
    Else
        MsgBox "THE W DRIVE SUBST CREATED SUCCESS", vbMsgBoxSetForeground
    End If

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
    
    Dim I As Long
    Dim j As Long
    Dim COUNTER5
    I = TargetHwnd
    Dim ARRAY5(20) As Long
    If TargetHwnd Then
        Do While I <> 0
            j = I
            COUNTER5 = COUNTER5 + 1
            ARRAY5(COUNTER5) = I
            I = GetParent(I)
        Loop
        I = j
    
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

    TargetHwnd = FindWindow(vbNullString, "InfoRapid Search & Replace")
    If TargetHwnd > 0 Then
        I = GetWindowRect(TargetHwnd, KAT_XY)
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
            
            I = SetForegroundWindow(TargetHwnd)
            SendKeys "N", True
            
            
        
        
        End If
    End If

    
    TargetHwnd = FindWindow(vbNullString, "Microsoft Office XP component")
    If TargetHwnd <> 0 Then
        RESULT = PostMessage(TargetHwnd, WM_CLOSE, 0&, 0&)
    End If
    
    
    TargetHwnd = FindWindow(vbNullString, "Windows Security")
    If TargetHwnd > 0 And 1 = 2 Then
        I = GetWindowRect(TargetHwnd, Windows_Security)
        XIOP = 0
        If Windows_Security.Right - Windows_Security.Left < 750 Then XIOP = 1
        If Windows_Security.Bottom - Windows_Security.Top < 350 Then XIOP = 1
        If XIOP = 1 Then
            On Error Resume Next
            DoEvents
            I = SetForegroundWindow(TargetHwnd)
            DoEvents
            SendKeys "Y", False
            DoEvents
            Sleep 400
            On Error GoTo 0
        End If
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
    
    TargetHwnd = FindWindow("#32770", "Kat MP3 Recorder")
    If TargetHwnd > 0 Then
        
        'X -- 642 -- Norm Use Window
        'Y -- 422 -- Norm Use Window
        I = GetWindowRect(TargetHwnd, KAT_XY)
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
            
            Call KILL_KAT_MP3_AND_THEN_SHELL
            
            'Shell KAT_MP3_PATH, vbMinimizedNoFocus
            
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
I = DateDiff("S", OXK3_KATMOUSE_PRESENCE_VISUAL_TIMER, OXK1_KATMOUSE_PRESENCE_TIMER)
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

If I > 40 And (XK1 > 0 And X1 = 0) Then
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

Private Sub TIMER_SET_CAMERA_UP_TO_MARK_FOLDER_Timer()

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

Call COPY_MY_SECURITY_INFO_TO_CAMERA

'WRITE_INFO_TXT_IN_EMPTY_FOLDER
Call WRITE_INFO_TXT_IN_EMPTY_FOLDER("W:\DCIM")
Call WRITE_INFO_TXT_IN_EMPTY_FOLDER("W:\MP_ROOT")


End Sub

Sub COPY_MY_SECURITY_INFO_TO_CAMERA()

    On Error Resume Next
    I = App.Path + "\CAMERA DATA"

    File2.Path = I
    If Err.Number > 0 Then Exit Sub
    File2.FileName = "*.*"
    
    For R1 = 0 To File2.ListCount - 1
        
        FN1 = File2.Path + "\" + File2.List(R1)
        If FS.FILEEXISTS(FN1) = True Then
            Set F1 = FS.getfile(FN1)
        End If
        
        FN2 = "W:\" + File2.List(R1)
        If FS.FILEEXISTS(FN2) = True Then
            Set F2 = FS.getfile(FN2)
        End If

        NOT_COPY = False
        If FS.FILEEXISTS(FN1) = True Then
            If FS.FILEEXISTS(FN2) = True Then
                If F1.datelastmodified = F2.datelastmodified Then
                    NOT_COPY = True
                End If
            End If
        End If
        
        If NOT_COPY = False Then
            
            FS.CopyFILE FN1, FN2
        
        
        End If
    Next
    


End Sub

Function LATEST_DATE_FILE_ON_CAMERA_CHANGED()

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
            Set F = FS.getfile(File2.Path + "\" + File2.List(R2))
            PHOTO_DATE = F.datelastmodified
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
    Open FOLDER_NAME + "\INFO.TXT" For Output As #FR1
        Print #FR1, "MUST MAKE FOLDER NOT EMPTY VICE VERSA"
        Print #FR1, Now
    Close #FR1
    If Err.Number > 0 Then
        MsgBox "ERROR MAKE FILE IN FOLDER FOR CAMERA" + vbCrLf + FOLDER_NAME + "\INFO.TXT", vbMsgBoxSetForeground
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

Private Sub Timer_W32TM_Timer()

'---------------------------------------
'USE THE SYSTEM SYNC TIME REGULAR - NICE
'-----------------------------------------------------------------------------
'CHECK IT OUT - IN CMD LINE SOMETIMES FAIL BUT WINDOWS RIGHT CORNER DON'T FAIL
'-----------------------------------------------------------------------------

If UNLOAD_FORM_FLAG = True Then Unload Me: Exit Sub

'EVRY TEN MIN TIMER

X2 = Minute(Now)
For R = 10 To 59 Step 10
    If R >= X2 Then Exit For
Next
XTM = R

If XTM <> X2 And TRIG2_GO3 = 1 Then TRIG2_GO3 = 0

If TRIG2_GO3 = 0 Then
    If X2 < OX2 Or XTM = X2 Then
        TRIG1_GO3 = True
        TRIG2_GO3 = 1
    End If
End If
OX2 = X2

If TRIG1_GO3 = False Then Exit Sub


If TRIG1_GO3 = True Or W32TM_RUN_ONCE = False Then
    W32TM_RUN_ONCE = True
    TRIG1_GO3 = False
    
    'START "" /REALTIME "W32TM" "/RESYNC"
    
    'K TO STAY - SEE RESULT
    'Shell "CMD /K ""W32TM /RESYNC""", vbNormalFocus
    
    'THE TIME.WINDOWS SERVER IS NOT ALWAYS GETTING THE CLOCK
    
    Shell "CMD ""W32TM /RESYNC""", vbHide
    

'    A = A

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


Sub VB_PROJECT_CHECKDATE()

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
        FS.CopyFILE PATH_FILE_NAME1, PATH_FILE_NAME2
    If Err.Number > 0 Then Exit Sub
End If


Set F = FS.getfile(PATH_FILE_NAME2)

VB_DATE = F.datelastmodified

'--------
'01 OF 02
'----------------------------------
'WRITE INFO ABOUT DATE CHNAGE NEWER
'----------------------------------

If XVB_DATE < VB_DATE And XVB_DATE > 0 Then
    'MsgBox "VB PROJECT DATE IS CHANGED MAYBE NEWER", vbMsgBoxSetForeground

    '------------------------------------------------
    'DO THE WRITE INFO ABOUT RELOAD MIRROR EXE FOLDER
    '------------------------------------------------
    TIMER_RETRY_WRITE_INFO_UNTIL_DONE1.Enabled = True
    TIMER_RETRY_WRITE_INFO_UNTIL_DONE2.Enabled = True

End If

XVB_DATE = VB_DATE

'--------
'02 OF 02
'----------------------------------------------------------------
'USE THIS CODE SNIP TO RELOAD ANY APPEND EXE IF WAITING FROM INFO
'----------------------------------------------------------------
'BUT NEVER RELOAD OUR OWN CODE -- LET ANOTHER PROGRAM
'----------------------------------------------------------------
    
If Dir("D:\VB6-EXE'S\#00 RELOAD_APPEND_SCRIPT\") = "" Then Exit Sub
    
File3.Path = "D:\VB6-EXE'S\#00 RELOAD_APPEND_SCRIPT\"
File3.FileName = "*.*"
File3.Refresh

On Error Resume Next
If File3.ListCount > 0 Then
    For R = 0 To File3.ListCount
        If InStr("-" + File3.List(R), "RELOAD ME -- ") > 0 Then
            
            FR1 = FreeFile
            Open File3.Path + "\" + File3.List(R) For Input As #FR1
                Line Input #FR1, RELOAD_VAR_TO_DO
            Close #FR1
            
            If Err.Number > 0 Then Exit Sub
            Debug.Print RELOAD_VAR_TO_DO
            
            Err.Clear
            
            '------------------------------------------------------
            'GREEN LIGHT GOT - AND THEN - ALL OVER FOR THIS PROGRAM
            '------------------------------------------------------
            'HANDED OVER TO EXE RELOADER SHELL
            '------------------------------------------------------
            If Dir(File3.Path + "\" + File3.List(R) + ".GREEN.TXT") = "" Then
                If UCase(RELOAD_VAR_TO_DO) <> UCase(PATH_FILE_NAME1) Then
                    FR1 = FreeFile
                    Open File3.Path + "\" + File3.List(R) + ".GREEN.TXT" For Output As #FR1
                    Close #FR1
                    If Err.Number > 0 Then Exit Sub
                    
                    Shell "D:\VB6\VB-NT\00_Best_VB_01\RELOAD_NETWORK_SYNC_EXE\RELOAD_Network_Sync_EXE_VB_MIRROR.exe """ + RELOAD_VAR_TO_DO + """"
                End If
            End If
        End If
    Next
End If

End Sub

Private Sub TIMER_RETRY_WRITE_INFO_UNTIL_DONE1_Timer()
    
    On Error Resume Next
    VARTIME_RELOAD = "RELOAD ME -- " + Format(Now, "YYYY MM DD HH MM SS")
    FR1 = FreeFile
    Open "D:\VB6-EXE'S\#00 RELOAD_APPEND_SCRIPT\" + VARTIME_RELOAD + ".TXT" For Output As #FR1
        Print #FR1, PATH_FILE_NAME1
        'Print #FR1, "APPENDING"
    Close #FR1
    
    If Err.Number = 0 Then
        TIMER_RETRY_WRITE_INFO_UNTIL_DONE1.Enabled = False
    End If

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
    
'D:\#0 1 INSTALLATIONS\Installation programs\# 00 New Install Progs\# Installed Now\#01 Matthew's\KAT MP3 RECORDER\kat-mp3-recorder-setup-free.exe
If KAT_MP3_PATH = "" And KAT_MP3_PATH_UNTOLD = False Then
    KAT_MP3_PATH_UNTOLD = True
    MsgBox "KAT MP3 APP PATH NOT FOUND" + vbCrLf + KAT_MP3_PATH1 + vbCrLf + "OR" + vbCrLf + KAT_MP3_PATH + vbCrLf + "PROGRAM ONLY USE IT FOR KNOWN COMPUTER NAME -- BUT SETUP PATH IS DONE BEFORE" + vbCrLf + "HAVE TO CORRECT THAT ONE IT DOES ALWAYS USE", vbMsgBoxSetForeground
End If
If KAT_MP3_PATH_UNTOLD = True Then Exit Sub

    
TIME_HAPPEN = Format(Now, "MM-DD DDD-DD-MMM -- HH:MM:SS")

If FIRST_RUN_FIND_PID_KAT_MP3 = False Then
    FIRST_RUN_FIND_PID_KAT_MP3 = True
    'PID HAS TO BE -1
    PID = -1
    Var = cProcesses.GetEXEID(PID, "C:\Program Files\Kat MP3 Recorder\Kat MP3 Recorder.EXE")
    If PID = -1 Then
        
        Shell KAT_MP3_PATH, vbMinimizedNoFocus
        COUNT_UP = Now
        Label1.Caption = TIME_HAPPEN + " -- Kat MP3 Recorder.EXE - NOT IN EXECUTION AT START OF ME PROGRAM"
        List1.AddItem Label1.Caption
        
        Exit Sub
    End If
End If

GO3 = 0

Set F = FS.getfOLDER(FILE1.Path)

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
MINICOMPARE = 80
If GetComputerName = "1-ASUS-X5DIJ" Then MINICOMPARE = 315
If GetComputerName = "5-ASUS-P2520LA" Then MINICOMPARE = 315
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

FILE1.Refresh

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

Set F = FS.getfile((FILE1.Path + "\" + FILE1.List(FILE1.ListCount - 1)))
If F.datelastmodified <> XDATE And XDATE <> 0 Then
    
    If FIRST_LIMIT_RESET_DONE = True Then

        COUNT_UP = Now
        COUNT_UP = F.datelastmodified
        Call COUNT_TIMER
        Label1.Caption = TIME_HAPPEN + " -- FILE DATE MODIFIED"

    End If
    
    XDATE = F.datelastmodified


    'If GetComputerName = "1-ASUS-X5DIJ" Then MINICOMPARE = 302 Else MINICOMPARE = 40
    
    If DateDiff("S", F.datelastmodified, Now) > MINICOMPARE Then
        RUN_CODE = True
    Else
        Exit Sub
    End If
End If

If XDATE = 0 Then XDATE = F.datelastmodified



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
    I = GetWindowRect(TargetHwnd, KAT_XY)
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



If GO3 = 1 Or (DateDiff("S", COUNT_UP, Now) > MINICOMPARE And COUNT_UP_WAIT = COUNT_UP_WAIT) Then

    If Label1.Caption = "" Then Label1.Caption = TIME_HAPPEN + " -- NOT REASON SET"
    
    List1.AddItem Label1.Caption
    
    
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
        Var = cProcesses.Process_Kill(PID)
    End If
    
    Sleep 1
    
    Call DEL_WAVS
    Call DEL_WAVS
    
    Shell KAT_MP3_PATH, vbMinimizedNoFocus

End Sub

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

Sub TIMER_DEL_CAMERA_EMPTY_MP4_TIMER()
    
    CAMERA_EXPLORER_PATH = "W:\DCIM"

    If Dir(CAMERA_EXPLORER_PATH, vbDirectory) = "" Then Exit Sub
    
    DoEvents
    Dir1.Path = "W:\MP_ROOT"
    Dir1.Refresh
    DoEvents
    
    For R = 0 To Dir1.ListCount - 1
        DoEvents
        
        On Error Resume Next
            Set F = FS.getfOLDER((Dir1.List(R)))
            If Err.Number > 0 Then Exit Sub
        On Error GoTo 0
    
        '19 IS NOT IF MARKER FOLDER
        If F.Size = 0 And Len(Dir1.List(R)) <= 19 Then
            
            On Error Resume Next
                RmDir Dir1.List(R)
            On Error GoTo 0
        End If
    Next
    
    TIMER_DEL_CAMERA_EMPTY_MP4.Enabled = False
    
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
    I = FindWinPart_VISIBLE(CAMERA_EXPLORER_PATH, "CabinetWClass", HWND_RETURN, VAR1, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)
    
    HOW_MANY_INSTANCE_SHOWING_VAR = HOW_MANY_INSTANCE_SHOWING
    
    If HOW_MANY_INSTANCE_SHOWING = 0 Then
        
        'DISABLED THIS FOR A TIME
        'PID = Shell("EXPLORER " + CAMERA_EXPLORER_PATH, vbNormalNoFocus)
        
        ONE_INSTANCE_ONLY = True
        ESCAPEE_COUNT = 20
        Do
            'DISABLED THIS FOR A TIME
            Exit Do
            
            I = FindWinPart_VISIBLE(CAMERA_EXPLORER_PATH, "CabinetWClass", HWND_RETURN, VAR1, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)
            
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
        Debug.Print "TIMEOUT RETRY COUNTDOWN TAKEN FROM 20 SECONDS --" + Str(ESCAPEE_COUNT)
    
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
    I = FindWinPart_VISIBLE("ViceVersa Pro - [", "", HWND_RETURN, VAR1, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)
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
    
    VICE_VERSA_SCRIPT_FILE = "D:\VICE_VERSA SCRIPT FILE\ViceVersa PHOTO CAMERA - LOCAL.fsf"
    If Dir(VICE_VERSA_SCRIPT_FILE) = "" Then
    
        MsgBox "UNABLE TO FIND" + vbCrLf + VICE_VERSA_SCRIPT_FILE + vbCrLf, vbMsgBoxSetForeground
        Exit Sub
    
    End If
    
    
    If LATEST_DATE_FILE_ON_CAMERA_CHANGED = False Then
        Exit Sub
    End If
    PID = Shell(VVCOMLINE + " """ + VICE_VERSA_SCRIPT_FILE + """", vbNormalFocus)

    I = cProcesses.Convert(PID, HWND_RETURN, cnFromProcessID Or cnTohWnd)
    
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
        I = FindWinPart_VISIBLE("ViceVersa Pro - [", "", HWND_RETURN, VAR1, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)
        
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
    Debug.Print "TIMEOUT RETRY COUNTDOWN TAKEN FROM 20 SECONDS --" + Str(ESCAPEE_COUNT)

    If ESCAPEE_COUNT = 0 Then
        'MsgBox "TIMEOUT WAIT REACH 0 FROM 10 SEC"
    End If
    
    If UNLOAD_FORM_FLAG = True Then Unload Me: Exit Sub
    
    Timer_ME_CAMERA.Enabled = True
    

End Sub

Sub TIMER_ME_CAMERA_TIMER()
    
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
    Dir2.Path = "W:\MP_ROOT"
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
                    Set F = FS.getfile(File2.Path + "\" + File2.List(R2))
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
                If F.Size > 100 * (1024 ^ 2) Then
                    If F.datelastmodified > Now - 2 Then
                        FLAG_GO_ME = True
                    End If
                End If
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
     Load Form3_Me_ViceVersa
        
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
    TIMER_DEL_THUMB_CAMAREA_MP4.Enabled = True

End Sub

Sub TIMER_DEL_THUMB_CAMAREA_MP4_TIMER()

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
                Set F1 = FS.getfile(File2.Path + "\" + File2.List(R2))
                MATCH_FILENAME = Replace(File2.List(R2), ".THM", ".MP4")
                If FS.FILEEXISTS(File2.Path + "\" + MATCH_FILENAME) = False Then
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
        I = FindWinPart_VISIBLE(WAIT_STRING, "", HWND_RETURN, VAR1, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY)
        If VAR1 = 1 And I = HWND_RETURN Then Exit Do
        If VAR1 = 1 And TIMEOUT_WAIT_USED = False Then
            Exit Do
        End If
        
        If UNLOAD_FORM_FLAG = True Then Unload Me: Exit Sub
        
        DoEvents
        Sleep 1000
        ESCAPEE_COUNT = ESCAPEE_COUNT - 1
    Loop Until (VAR1 = 1 And I = HWND_RETURN) Or ESCAPEE_COUNT <= 0
    
    Debug.Print "TIMEOUT RETRY COUNTDOWN TAKEN FROM 20 SECONDS --" + Str(ESCAPEE_COUNT)
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
    Debug.Print "IsWindowVisible -- " + Str(ESCAPEE_COUNT)
    '-----------------------------------------------------
    ESCAPEE_COUNT = 100
    Do
        'IS WINDOW SHOWING 1 = TRUE
        If IsWindow(HWND_RETURN) = 1 Then Exit Do
        DoEvents
        Sleep 100
        ESCAPEE_COUNT = ESCAPEE_COUNT - 1
    Loop Until (IsWindow(HWND_RETURN) = 1) Or ESCAPEE_COUNT <= 0
    Debug.Print "IsWindow -- " + Str(ESCAPEE_COUNT)
    '----------------------------------------------
    ESCAPEE_COUNT = 100
    Do
        'IS IsIconic = 0 ALWAYS FOR NORMAL FORM WINDOW
        If IsIconic(HWND_RETURN) = IsIconic(HWND_RETURN) Then Exit Do
        DoEvents
        Sleep 100
        ESCAPEE_COUNT = ESCAPEE_COUNT - 1
    Loop Until (IsIconic(HWND_RETURN) = IsIconic(HWND_RETURN)) Or ESCAPEE_COUNT <= 0
    Debug.Print "IsIconic -- " + Str(ESCAPEE_COUNT)
    '----------------------------------------------
    ESCAPEE_COUNT = 100
    Do
        'IS IsZoomed = ALWAYS 0 UNLESS MAXIMISED - THINK MOMENT
        If IsZoomed(HWND_RETURN) = IsZoomed(HWND_RETURN) Then Exit Do
        DoEvents
        Sleep 100
        ESCAPEE_COUNT = ESCAPEE_COUNT - 1
    Loop Until (IsZoomed(HWND_RETURN) = IsZoomed(HWND_RETURN)) Or ESCAPEE_COUNT <= 0
    Debug.Print "IsZoomed -- " + Str(ESCAPEE_COUNT)

End Sub


Sub DEL_WAVS()

    'THIS THE DEL AFTER TIME
    'CALLED FROM -- SUB Timer_DEL_WAVS_Timer()

    DoEvents
    FILE1.Refresh
    DoEvents
    
    For R = FILE1.ListCount - 1 To 0 Step -1
        DoEvents
        If UNLOAD_FORM_FLAG = True Then Exit Sub
        
        If Right(FILE1.List(R), 4) = ".WAV" Or Right(FILE1.List(R), 4) = ".MP3" Then
            
            GetComputerVAR = False
            If GetComputerName = "1-ASUS-EEE" Then
                GetComputerVAR = False
            End If
            If GetComputerName = "5-ASUS-P2520LA" Then
                DEL_DATE_DAYS = 40
                GetComputerVAR = True
            End If
            If GetComputerName = "3-ASUS-X5DIJ" Then
                DEL_DATE_DAYS = 8
                GetComputerVAR = True
            End If
            
            '-------------------------------------------------------------------------
            DoEvents
            '-------------------------------------------------------------------------
            
            '------------------------------------------------------------------
            'ASUS-EEE RADIO
            'ACER-ONE RADIO
            '------------------------------------------------------------------
            If GetComputerVAR = False Then
                
                On Error Resume Next
                    Set F = FS.getfile((FILE1.Path + "\" + FILE1.List(R)))
                    If Err.Number > 0 Then Exit Sub
                On Error GoTo 0
            
                CHECK_RUN = False
                If F.Size < 57 Then
                    
                    On Error Resume Next
                        Kill FILE1.Path + "\" + FILE1.List(R)
                        CHECK_RUN = True
                    On Error GoTo 0
                End If
        
            End If
    
    
            '-------------------------------------------------------------------------
            DoEvents
            '-------------------------------------------------------------------------
    
            '-------------------------------------------------------------------------
            'IS THE -- ASUS-X5DIJ AUDIO - DAY DELAY AND NOT NEED SIZE WHEN MOST RECENT
            '-- ASUS-X5DIJ AUDIO
            '-- ASUS-P2520LA AUDIO
            '-------------------------------------------------------------------------
            If GetComputerVAR = True Then
            
                On Error Resume Next
                    Set F = FS.getfile((FILE1.Path + "\" + FILE1.List(R)))
                    If Err.Number > 0 Then Exit Sub
                On Error GoTo 0
                
                
                'TROUBLE GETTING THE DATE
                On Error Resume Next
                    ADATE = F.datelastmodified
                    If Err.Number > 0 Then Exit Sub
                On Error GoTo 0
                
                'KEEP A FOUR HOUR LOGG
                'ADJUST AGAIN TO MAKE 4 DAY
                'ADJUST AGAIN TO MAKE 8 DAY
                If ADATE < Now - DEL_DATE_DAYS Then 'TimeSerial(4, 0, 0) Then
                    'ONLY THE NORMAL SHORT FILE NAME LENGHT
                    'AS KEEP WHAT IS RENAME LABELED INFO
                    'ABOUT NEIL NEXT DOOR TALK ABOUT HAND ON THE VOLUME KNOB
                    If Len(FILE1.List(R)) <= 23 Then
                        On Error Resume Next
                            Kill FILE1.Path + "\" + FILE1.List(R)
                        On Error GoTo 0
                    End If
                End If
                If ADATE < Now - TimeSerial(2, 0, 0) Then
                    If Len(FILE1.List(R)) <= 23 Then
                        If F.Size = 0 Then
                            On Error Resume Next
                                Kill FILE1.Path + "\" + FILE1.List(R)
                            On Error GoTo 0
                        End If
                    End If
                End If
                

            End If
        End If
    
    Next
    
    
    
    
    
End Sub

Private Sub Timer_DEL_WAVS_Timer()

    If UNLOAD_FORM_FLAG = True Then Unload Me: Exit Sub

    INTERVAL_DEL_WAVS = Hour(Now)
    If INTERVAL_DEL_WAVS = SET_INTERVAL_DEL_WAVS Then
        Exit Sub
    End If
    SET_INTERVAL_DEL_WAVS = INTERVAL_DEL_WAVS

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

    I = GET_DRIVES_FOR_CAMERA
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
        
        If D.ISREADY Then
            N = N + D.DriveLetter 'd.VolumeName
            
            If D.volumename = "##" Or D.volumename = "MC - HX60V" Then
                VAR_SHELL_SUBST_BEEN_SET_FLAG = True
                VAR_SHELL_SUBST_BEEN_SET_FLAG2 = True
                
                If VAR_SHELL_SUBST_BEEN_SET = False Then
                    'Debug.Print d.DriveLetter + " -- " + d.volumename + " -- " + Str(d.drivetype)
                    Shell "SUBST W: " + D.DriveLetter + ":\", vbHide
                    VAR_SHELL_SUBST_BEEN_SET = True
                    TIMER_DEL_CAMERA_EMPTY_MP4.Enabled = True
                    TIMER_SET_CAMERA_UP_TO_MARK_FOLDER.Enabled = True
                    
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
        '------------------------------------------
    End If
End If

Set DC = Nothing

END_EXIT:

End Function

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

If COMPUTERNAME <> "1-ASUS-EEE" Then Exit Sub

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


Private Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim l As Long
   Dim S As String
   l = GetWindowTextLength(hwnd)
   S = Space(l + 1)
   GetWindowText hwnd, S, l + 1
   GetWindowTitle = Left$(S, l)
End Function


Function FindWinPart(FIND_STRING) As Long


Dim Ash$
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
    
    
    Ash$ = GetWindowTitle(Test_HWND)
    
    'cText = Space$(255)
    'ghj$ = GetClassName(Test_HWND, cText, 255)
    If Ash$ <> "" Then
        If InStr("-" + Ash$, FIND_STRING) > 0 Then
            FindWindowPart = Test_HWND
            Exit Do
        End If
    End If
    
    'retrieve the next window
    Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)
Loop


End Function



Function FindWinPart_VISIBLE(FIND_STRING, CLASS_TEXT, TARGET_HWND_CHECK, VAR1, HOW_MANY_INSTANCE_SHOWING, MATCH_HWND, GetForegroundWindow_VAR, ONE_INSTANCE_ONLY) As Long

'----------------
'FINDWINDOWPART
'FIND_WINDOW_PART
'----------------
Dim Ash$
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
    
    
    Ash$ = GetWindowTitle(Test_HWND)
    
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
    
    
    If Ash$ <> "" And CLASS_FLAG = True Then
        If InStr("-" + Ash$, FIND_STRING) > 0 Then
            HWND_COUNT = HWND_COUNT + 1
            
            VAR1 = IsWindow(Test_HWND)
'            Debug.Print Format(HWND_COUNT, "000") + " -- HWND FROM FIND_WIN_PART =" + Str(HWND_RETURN)
'            Debug.Print Format(HWND_COUNT, "000") + " -- " + Ash$
'            Debug.Print Format(HWND_COUNT, "000") + " -- IS_VISIBLE" + Str(VAR1)
            
            '----------------------------------------------
            'Test_HWND IS LOWER THAN TARGET_HWND_WHEN CHECK
            '----------------------------------------------
            If Test_HWND < TARGET_HWND_CHECK Then
                Debug.Print Format(HWND_COUNT, "000") + " -- Test_HWND < TARGET_HWND_CHECK --" + Str(Test_HWND) + " -- " + Str(TARGET_HWND_CHECK) + " -- " + Str(TARGET_HWND_CHECK - Test_HWND) + " -- PROPER"
            End If
            '----------------------------------------------
            If Test_HWND > TARGET_HWND_CHECK Then
                'Debug.Print Format(HWND_COUNT, "000") + " -- Test_HWND > TARGET_HWND_CHECK --" + Str(Test_HWND) + " -- " + Str(TARGET_HWND_CHECK)
            End If
            '----------------------------------------------
            If TARGET_HWND_CHECK = Test_HWND Then
                Debug.Print Format(HWND_COUNT, "000") + " -- FIND WINODW PART GOT MATCH HWND --" + Str(Test_HWND)
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

'RESULT SAME AS ENTRY MUST BE MAIN PARENT
I = GET_PARENT(Test_HWND)
'----------------------------------------
I = GET_CHILD_TEST(Test_HWND)
'----------------------------------------
'NOT ANY RESULT WE WANT HERE

End Function
'GET_CHILD_TEST
Public Function GET_CHILD_TEST(ByVal TargetHwnd As Long) As String

    Dim I As Long
    Dim j As Long
    Dim COUNTER5
    I = TargetHwnd

    Do While I <> 0
        j = I
        I = GetWindow(I, GW_CHILD)
        If I = j Then Exit Do
        DoEvents
'        Debug.Print I
        
        If I = TargetHwnd Then
            Debug.Print "GET_CHILD_TEST --" + Str(TargetHwnd)
        End If
    Loop
    I = j
            
GET_CHILD_TEST = I


End Function

Public Function GET_PARENT(ByVal TargetHwnd As Long) As String

    Dim I As Long
    Dim j As Long
    Dim COUNTER5
    I = TargetHwnd
    
    If TargetHwnd Then
        Do While I <> 0
            j = I
            COUNTER5 = COUNTER5 + 1
            I = GetParent(I)
        Loop
        I = j
    
    End If
    
GET_PARENT = I
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
  ElseIf D.ISREADY Then
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
