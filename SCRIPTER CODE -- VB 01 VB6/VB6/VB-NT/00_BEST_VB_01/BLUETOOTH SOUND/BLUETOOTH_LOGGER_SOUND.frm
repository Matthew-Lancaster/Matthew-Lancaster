VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "BLUETOOTH SOUND"
   ClientHeight    =   3108
   ClientLeft      =   168
   ClientTop       =   540
   ClientWidth     =   10908
   LinkTopic       =   "Form1"
   ScaleHeight     =   3108
   ScaleWidth      =   10908
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer_1_MINUTE 
      Interval        =   1000
      Left            =   4920
      Top             =   876
   End
   Begin VB.Timer Timer_BLUE_EMAIL 
      Interval        =   1000
      Left            =   4788
      Top             =   456
   End
   Begin VB.FileListBox File1 
      Height          =   264
      Left            =   84
      TabIndex        =   2
      Top             =   2508
      Visible         =   0   'False
      Width           =   1224
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2568
      Top             =   900
   End
   Begin VB.Timer Timer_KEY_ASYNC 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2172
      Top             =   900
   End
   Begin MCI.MMControl MMControl1 
      Height          =   336
      Left            =   1776
      TabIndex        =   1
      Top             =   1332
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   593
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   900
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Left            =   1392
      TabIndex        =   3
      Top             =   0
      Width           =   1236
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   15
      TabIndex        =   0
      Top             =   -60
      Width           =   1305
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu Mnu_Explore_AUDIO_File 
      Caption         =   "Explorer Select __ APP Folder __ Open AUDIO FILE"
   End
   Begin VB.Menu MNU_EDIT_LOGG 
      Caption         =   "EDIT LOGG"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_EXIT_IN 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XVB_DATE


Dim START_TIME As Date

'VOLUME WANT

Dim NUM_VAR_INDEX

Dim OWindowState
Dim FORM_LOAD_DONE


Dim AUDIO

Dim VARCENTER, XCount, JOB_LoggFile


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


'http://www.vbforums.com/showthread.php?673134-RESOLVED-Minimum-height-for-menu-bar-to-be-visible
'-------------- MENU SIZE DECLARE
Private Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type MENUBARINFO
  cbSize As Long
  rcBar As rect
  hMenu As Long
  hwndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type
Private MenuInfo As MENUBARINFO
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long



Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
'----------------
'DECLARE ---- END
'----------------

'---------
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
'---------
'### USAGE
'sWindowsFolder = GetSpecialfolder(36)
'sMyDocsFolder = GetSpecialfolder(5)
'---------
'Dim R As Long
'On Error Resume Next
'For R = 0 To 120
'    Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
'Next
'End
'---------
'---------
Private Enum SpecialFolderConstants
    spfldDesktop = &H0
    spfldInternet = &H1
    spfldPrograms = &H2
    spfldControls = &H3
    spfldPrinters = &H4
    spfldPersonal = &H5
    spfldFavorites = &H6
    spfldStartUp = &H7
    spfldRecent = &H8
    spfldSendTo = &H9
    spfldBitBucket = &HA
    spfldStartMenu = &HB
    spfldDesktopDirectory = &H10
    spfldDrives = &H11
    spfldNetwork = &H12
    spfldNethood = &H13
    spfldFonts = &H14
    spfldTemplates = &H15
    spfldCommonStartMenu = &H16
    spfldCommonPrograms = &H17
    spfldCommonStartup = &H18
    spfldCommonDesktopDirectory = &H19
    spfldAppData = &H1A
    spfldPrintHood = &H1B
    spfldAltStartup = &H1D             '' DBCS
    spfldCommonAltStartup = &H1E       '' DBCS
    spfldCommonFavorites = &H1F
    spfldInternetCache = &H20
    spfldCookies = &H21
    spfldHistory = &H22
    spfldWindows = &H24
    spfldWindowSystem = &H25
    
    '&H27& My Pictures
    '40 User Profile
    '43 Common Files
    '46 All Users Templates
    '47 Administrative Tools
    '49 Network Connections
    
End Enum
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type




'ALL YOU WANT FOR SPECIAL FOLDER
'------------------------
'------------------------
'Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
'------------------------
'Private Type SHITEMID
'    cb As Long
'    abID As Byte
'End Type
'Private Type ITEMIDLIST
'    mkid As SHITEMID
'End Type
'Public Enum SpecialFolderConstants
'AND REST OF ABOVE LINE
'------------------------
'------------------------

Private Function GetSpecialfolder(CSIDL As Long) As String
    '##############################################################################################
    'Returns the Path to a "Special" Folder (i.e. Internet History)
    '##############################################################################################
    
    Dim IDL As ITEMIDLIST
    Dim lResult As Long
    Dim sPath As String
    
    lResult = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If lResult = 0 Then
        sPath = Space$(512)
        lResult = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        
        GetSpecialfolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
End Function
'------------------------
'------------------------




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

If FORM_LOAD_DONE = False Then Exit Sub

Timer2.Enabled = False

End Sub

Private Sub Form_GotFocus()
'Timer2.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then End
End Sub

Private Sub Form_Load()

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
Dim EXECUTE_FILE_NAME
'EXECUTE_FILE_NAME = "C:WINDOWS\EXPLORER.EXE " + EXECUTE_PARAM
EXECUTE_FILE_NAME = "C:\PStart\Progs\0_Nirsoft_Package\NirSoft\BlueToothView_Desc.VBS"
WSHShell.Run """" + EXECUTE_FILE_NAME + """"
Set WSHShell = Nothing


Timer_BLUE_EMAIL.Enabled = True


Exit Sub






'Exit Sub



'Beep

'@REMER
'------

'AUDIO = "D:\0 00 MUSIC ---\09 SOUND EFFECT\VB Wave's\Sound_Effects\Nautical\Sonar Copy.wav"
'
'If Dir(AUDIO) = "" Then
'    AUDIO = App.Path + "\AUDIO DATA\Sonar Copy.wav"
'End If
'
'If GetComputerName <> "1-ASUS-X5DIJ" Then
'    AUDIO = "D:\# MY DOCS\# 01 My Documents\My Music\Peter, Paul, And Mary Puff The Magic Dragon.wav"
'    If Dir(AUDIO) = "" Then
'        AUDIO = App.Path + "\AUDIO DATA\Peter, Paul, And Mary Puff The Magic Dragon.wav"
'    End If
'
'End If

START_TIME = Now


Dim NUM_VAR_STRING As String

If Dir(App.Path + "\AUDIO DATA\z# ---- INDEX_PLAY__" + GetComputerName + ".TXT") <> "" Then
    FR1 = FreeFile
    Open App.Path + "\AUDIO DATA\z# ---- INDEX_PLAY__" + GetComputerName + ".TXT" For Binary As #FR1
        NUM_VAR_STRING = Space(LOF(FR1))
        Get #FR1, 1, NUM_VAR_STRING
    Close #FR1
    NUM_VAR_INDEX = Val(NUM_VAR_STRING)
Else
    Beep
End If

File1.Path = App.Path + "\AUDIO DATA"
File1.Pattern = "*.WAV"

IWARPPER = NUM_VAR_INDEX
Do
    AUDIO = File1.List(NUM_VAR_INDEX)
    
    ' IS IT A REM IGNOR SKIP - DONT WANTING
    ' OR DO WE WANT RESET TRY ANOTHER
    If UCase(Mid(AUDIO, 1, 2)) = "Z#" Then
        NUM_VAR_INDEX = NUM_VAR_INDEX + 1
        AUDIO = ""
    End If
    
    If NUM_VAR_INDEX >= File1.ListCount Then
        NUM_VAR_INDEX = 0
    End If

    'WRAP AROUND --- IWRAPER FULL LOOP

Loop Until AUDIO <> "" And (UCase(Mid(AUDIO, 1, 2)) <> "Z#" Or NUM_VAR_INDEX = IWARPPER)

'AUDIO = File1.Path + "\" + File1.List(NUM_VAR_INDEX)
NUM_VAR_INDEX = NUM_VAR_INDEX + 1

Debug.Print Str(NUM_VAR_INDEX) + " -- " + AUDIO
Label2.Caption = AUDIO

If NUM_VAR_INDEX >= File1.ListCount Then NUM_VAR_INDEX = 0

'----------------------------------
FILETEXT = App.Path + "\AUDIO DATA\z# ---- INFO -- PUT IN FRONT TO REM SKIP IGNOR -- ONLY WAV FILE - OR EXTERNAL LOADER STOP ABILITY TO STOP MP3 DEMAND.TXT"

If IsIDE = True And Dir(FILETEXT) <> "" Then Kill FILETEXT

If Dir(FILETEXT) = "" Then
    On Error Resume Next
    TAP_COUNTER = 0
    TAP_COUNTER_LIMIT = 2000
    Do
        Err.Clear
        FR1 = FreeFile
        Open FILETEXT For Output As #FR1
            Print #FR1, "------------------------------------------------------------------------------------------------------------------"
            Print #FR1, FILETEXT
            Print #FR1,
            Print #FR1, "A NOTE FOR THE USER -- WILL BE MADE BY PROGRAM CODE IF NOT THERE " + Format(Now, "DDDD DDD DD MMMM YYYY HH:MM:SS")
            Print #FR1,
            Print #FR1, "------------------------------------------------------------------------------------------------------------------"
            
            
        
        Close #FR1
        
        If Err.Number = 0 Then Exit Do
        Sleep 200
    
        DoEvents
    Loop Until Err.Number = 0 Or TAP_COUNTER >= TAP_COUNTER_LIMIT
    
    If TAP_COUNTER = TAP_COUNTER_LIMIT Then ' ---- 2000
        Beep
    End If
End If
'----------------------------------


'----------------------------------
On Error Resume Next
TAP_COUNTER = 0
TAP_COUNTER_LIMIT = 2000
Do
    Err.Clear
    FR1 = FreeFile
    Open App.Path + "\AUDIO DATA\z# ---- INDEX_PLAY__" + GetComputerName + ".TXT" For Output As #FR1
        Print #FR1, Format(NUM_VAR_INDEX, "0");
    Close #FR1
    
    If Err.Number = 0 Then Exit Do
    Sleep 200

    DoEvents
Loop Until Err.Number = 0 Or TAP_COUNTER >= TAP_COUNTER_LIMIT

If TAP_COUNTER = TAP_COUNTER_LIMIT Then ' ---- 2000
    Beep
End If
'----------------------------------

If Dir(File1.Path + "\" + AUDIO) = "" Then
    'AUDIO = App.Path + "\AUDIO DATA\Sonar Copy.wav"
    'AUDIO = App.Path + "\AUDIO DATA\Peter, Paul, And Mary Puff The Magic Dragon.wav"
        
    i = MsgBox("FAULT WITH READ THE AUDIO DATA NOT AT PATH" + vbCrLf + vbCrLf + App.Path + "\AUDIO DATA\*.wav" + vbCrLf + vbCrLf + "WANT EXPLORER LOAD THERE -- DO YOU -- MAXIMIZED", vbYesNo + vbMsgBoxSetForeground)
    
    If i = vbYes Then
        Shell "Explorer.exe /E, """ + App.Path + "\AUDIO DATA\  """, vbMaximizedFocus
    End If
End If





Label1.FontSize = 32
Label1.Left = 0
Label1.Top = 0
Label1.AutoSize = False
'Label1.AutoSize = True
Me.Caption = "BlueTooth Sound -- " + Format(Now, "DDD DD MMM HH:MM:SS")
Label1.Caption = Format(Now, "DDD DD MMM HH:MM:SS")
Label1.AutoSize = False
'Label1.AutoSize = True
'Label1.AutoSize = False
DoEvents

Me.Width = 14000

Label1.Width = 14000
Label1.Refresh


Label2.FontSize = 12
Label2.Height = 1500
Label2.Left = 0
Label2.Top = Label1.Top + Label1.Height + 20
'Label1.AutoSize = False
'Label1.AutoSize = True
Label2.Caption = AUDIO
'Label1.AutoSize = False
'Label1.AutoSize = True
'Label1.AutoSize = False
DoEvents
Me.Width = 14000
Label2.Width = 14000
Label2.Refresh

'Me.Width = Label1.Width + 120
VARCENTER = False
Me.Height = Label2.Top + Label2.Height + 780
'VARCENTER = False


Call RESET_SETUP_SOUND_FILE_AND_PLAY("YESSOUND", File1.Path + "\" + AUDIO)


'----------------------------------------------

'RUN SOUND EVENT NOW REST OF CODE _ FORM LAYOUT

'----------------------------------------------




DoEvents

VARCENTER = False
Call Form_Resize
'Me.Width = Label1.Width + 1000
'Me.Height = Label1.Height + 900


'SETFORE
DoEvents
Me.Show
DoEvents
Me.Refresh
DoEvents
Me.Show
DoEvents


SetForegroundWindow (Me.hWnd)

FORM_LOAD_DONE = True


Exit Sub











vfile = App.Path + "\GoodSync_Logger CommandLine--" + GetComputerName + ".txt"
If Dir(vfile) = "" Then XX = 22
FR1 = FreeFile
Open vfile For Append As FR1
    Print #FR1, Format(Now, "YYYY-MM-DD HH-MM-SS") + " | """ + Command$ + """"
Close #FR1

'Me.WindowState = vbMinimized

If XX = 22 Then ShellExecute Me.hWnd, "open", vfile, vbNullString, vbNullString, conSwNormal



'Clipboard.Clear
'Clipboard.SetText Command$ + " ###"


'End



'XX = 59
'If IsNumeric(Command$) = True Then If Val(Command$) > XX Then XX = Val(Command$)
'If IsIDE = True Then XX = 8: Me.Show: Me.WindowState = vbNormal
'Me.Caption = Format(XX, "00") + " DELAYED LOADER"


'Label1.Caption = "OFFICE 2000 TOOLBAR DELAYED LOADER" + vbCrLf + "AND ALL *.LNK'S FROM" + vbCrLf + " E:\01 Start Menu\Programs\Startup\DELAYED_LOADER\ "
'Label1 = Label1 + vbCrLf + "SET THE COMMAND LINE TO SECONDS TO DELAYED LOAD" + vbCrLf + "DEFAULT = 30"
'If Command$ <> "" And IsNumeric(Command$) Then Label1 = Label1 + vbCrLf + "Current =" + Str(Val(Command$))
'If IsIDE = True Then Label1.Caption = Label1.Caption + vbCrLf + "IN ISIDE =" + Str(XX)
'Label1 = Label1 + vbCrLf + "AND IF THE SCHEDULAR FOR DEFRAGGLER ON DRIVE E - DON'T EXIST" + vbCrLf + "THEN IT IS LOADED"


Label1.Caption = "GoodSync Job @" + Format(Now, "DDD DD-MM-YYYY HH-MM-SS") + " Time By Script Program" + vbCrLf
'Label1 = Label1 + "########" + vbCrLf
'If Command$ = "" Then
'    'Label1 = Label1 + "Command Line = Empty - Program Test Run ---" + vbCrLf
'End If

'TEST LINE





If Command$ <> "" Then
    Label1 = Label1 + Command$ + vbCrLf
Else
    'not sure if this is how post ANALYZE - Is Meant to Be Taken From Post Sync
    'Label1 = Label1 + "--_POST ANALYZE_ --JOB NAME __ ""E to D #E"" --RESULT __ """" --LOGFILE __ ""C:\Documents and Settings\Matt 04\Application Data\GoodSync\_mirrors_\E-\2015-1110-133949-ASUS-BIGGER-E to D #E.log"" --LEFT __ ""E:"" --RIGHT __ ""D:\#\#E"" --LEFT FOLDER RESOLVED __ ""E:"" --RIGHT FOLDER RESOLVED __ ""D:/#/#E"" --CHANGED __ 0 --ERRORS __ 0 --CONFLICTS __ 0" + vbCrLf
    Label1 = Label1 + "--_POST ANALYZE_ --JOB NAME __ ""E SYNC ASUS EEE"" --RESULT __ ""Error analyzing E:: stopped by Stop By User"" --LEFT __ ""E:"" --RIGHT __ ""\\ASUS-EEE\EEE E"" --LOGFILE __ ""C:\Documents and Settings\Matt 04\Application Data\GoodSync\_mirrors_\E-\2015-1111-011908-ASUS-BIGGER-E SYNC ASUS EEE.log"" __ --LEFT FOLDER RESOLVED __ ""E:"" --RIGHT FOLDER RESOLVED __ ""//ASUS-EEE/EEE E"" --ERRORS __ 0 --CHANGED __ 32 --CONFLICTS __ 0"
    'Label1 = "--_POST ANALYZE_ --JOB NAME __ ""E SYNC ASUS EEE"" --RESULT __ ""Error analyzing E:: stopped by Stop By User"" --LEFT __ ""E:"" --RIGHT __ ""\\ASUS-EEE\EEE E"" --LOGFILE __ ""C:\Documents and Settings\Matt 04\Application Data\GoodSync\_mirrors_\E-\2015-1111-011908-ASUS-BIGGER-E SYNC ASUS EEE.log"" __ --LEFT FOLDER RESOLVED __ ""E:"" --RIGHT FOLDER RESOLVED __ ""//ASUS-EEE/EEE E"" --ERRORS __ 0 --CHANGED __ 32 --CONFLICTS __ 0"
End If

'TEST LINE
'Label1 = "--_POST SYNC_ --JOB NAME __ ""E to D #E"" --RESULT __ """" --LOGFILE __ ""C:\Documents and Settings\Matt 04\Application Data\GoodSync\_mirrors_\E-\2015-1110-133949-ASUS-BIGGER-E to D #E.log"" --LEFT __ ""E:"" --RIGHT __ ""D:\#\#E"" --LEFT FOLDER RESOLVED __ ""E:"" --RIGHT FOLDER RESOLVED __ ""D:/#/#E"" --CHANGED __ 0 --ERRORS __ 0 --CONFLICTS __ 0"""


'If InStr(Label1, "PRE ANALYZE") > 0 Then Me.Caption = Me.Caption + " - PRE ANALYZE"
If InStr(Label1, "--_PRE ANALYZE_") > 0 Then Me.Caption = Me.Caption + " --_PRE ANALYZE_"
If InStr(Label1, "--_POST ANALYZE_") > 0 Then Me.Caption = Me.Caption + " --_POST ANALYZE_"
If InStr(Label1, "--_POST SYNC_") > 0 Then Me.Caption = Me.Caption + " --_POST SYNC_"


'Label1 = "GoodSync Job Done @Sat 31-10-2015 10-37-02 --JOB NAME ""EDDIE ELEMENT 1TB - E"" --RESULT """" --LOGFILE ""C:\Documents and Settings\Matt 04\Application Data\GoodSync\_mirrors_\E-\2015-1031-103642-ASUS-BIGGER-EDDIE ELEMENT 1TB - E.log""  --LEFT ""E:"" --RIGHT ""=WD_E_1TB_3E:"" --LEFT FOLDER RESOLVED ""E:"" --RIGHT FOLDER RESOLVED ""Y:"" --CHANGED 0 -ERRORS 0 --CONFLICTS 0 --SYNCED-OK ITEMS POST 0 --SYNCED-OK FILE CHANGES POST 137"






Call Mnu_Fix1st_Letters_Click

'Clipboard.Clear
'Clipboard.SetText Label1

'Label1 = Replace(Label1, "########" + vbCrLf, "Command Line Without Modify = " + Command$ + vbCrLf)


If IsIDE = False Then
    FR1 = FreeFile
    Open App.Path + "\GoodSync_Logger--" + GetComputerName + ".txt" For Append As FR1
    
        Print #FR1, "-------------------------------------------"
        Print #FR1, "Command Line = """ + Command$ + """" + vbCrLf
        Print #FR1, "-------------------------------------------"
'        Print #fr1, "Time is Displayed From This Program Not Script Parameter"
'        Print #fr1,
        Print #FR1, Label1.Caption;
        Print #FR1, "-------------------------------------------"
        Print #FR1,
    
    Close #FR1
Else
    Me.WindowState = vbNormal
End If




'--SYNCED-OK ITEMS POST __ 55
'--SYNCED-OK FILE CHANGES POST __ 55


'--JOB NAME __ "D Microsoft Sync ASUS EEE D"

'Clipboard.Clear
'Clipboard.SetText Label1
'MsgBox "Ok"



AX1 = InStr(Label1, "--JOB NAME __ ")
If AX1 > 0 And InStr(Label1, "--_POST ANALYZE_") > 0 Then
    
    AX2 = InStr(AX1, Label1, vbCrLf)
    AX3 = AX2 - AX1
    JOB = Mid(Label1, AX1 + 2, AX3 - 4)
    
    Me.Caption = Replace(Me.Caption + " - " + JOB, "_", "")
    Me.Caption = Replace(Me.Caption, "--", "-- ")
    Me.Caption = Replace(Me.Caption, "- JOB", "-- JOB :-")
    Me.Caption = Me.Caption + """"
    
    AX1 = InStr(Label1, "--LOGFILE __ ")
    AX2 = InStr(AX1, Label1, vbCrLf)
    AX3 = AX2 - AX1
    
    JOB_LoggFile = Mid(Label1, AX1 + 14, AX3 - 19)
    
    '--LOGFILE __ ""C:\Documents and Settings\Matt 04\Application Data\GoodSync\_mirrors_\E-\2015-1111-011908-ASUS-BIGGER-E SYNC ASUS EEE.log""
    
    '--CHANGED __
    
    AX1 = InStr(Label1, "--CHANGED __ ")
    AX2 = InStr(AX1, Label1, vbCrLf)
    AX3 = AX2 - AX1
    JOB_CHANGED = Mid(Label1, AX1 + 2, AX3 - 2)
    JOB_CHANGED_VAL = Val(Mid(Label1, AX1 + 12, AX3 - 13))
    
    '--CONFLICTS __
    
    AX1 = InStr(Label1, "--CONFLICTS __ ")
    JOB_CONFLICTS = Mid(Label1, AX1 + 2)
    JOB_CONFLICTS_VAL = Val(Mid(Label1, AX1 + 14))

    If JOB_CHANGED_VAL > 20 Or JOB_CONFLICTS_VAL > 0 Then
        xyz = True
    Else
        'Beep
    End If
End If


'MsgBox "Ok"


MNU_EXIT_IN.Caption = "Exit in (- 8 -) Minutes"




If JOB_CONFLICTS_VAL > 0 Then
    ExitVar = "Wait ConFlicts"
    MNU_EXIT_IN.Caption = "*** " + ExitVar + " ***"
End If
If JOB_CHANGED_VAL > 20 Then
    If JOB_CONFLICTS_VAL > 0 Then
        ExitVar = ExitVar + " - and Changes" + Str(JOB_CHANGED_VAL) + "> 20 = True"
    Else
        ExitVar = ExitVar + "Changes" + Str(JOB_CHANGED_VAL) + " > 20 Trigger = True"
    End If
    MNU_EXIT_IN.Caption = "*** " + ExitVar + " ***"

End If

Label1 = Label1 + vbCrLf + "######## -------- " + MNU_EXIT_IN.Caption

If JOB_CONFLICTS_VAL > 0 Or JOB_CHANGED_VAL > 20 Then
    Label = Label + " -- Condition Decision Making Act"
End If

Label1.Left = 0
Label1.Top = 0

Label1.Refresh
'DoEvents

'Me.Width = Label1.Width + 120
VARCENTER = False
'Me.Height = Label1.Height + 780
'VARCENTER = False

'MsgBox Label1

'VARCENTER = False
'Call Form_Resize

XCount = 480
XCount = 40
XCount = 2
If xyz = True Then
'    Beep
    'VARCENTER = False
    Me.WindowState = vbNormal
    Timer1.Enabled = False
    'MsgBox "Ok"

Else
'    Beep
    Timer1.Enabled = True
    Call Timer1_Timer
End If

Timer1.Enabled = True

'NUM_VAR_INDEX = -1
'EXIT SUB --IS IN THE BLOCK SUBROUTINE

End Sub


Sub RESET_SETUP_SOUND_FILE_AND_PLAY(VARSOUND, AUDIO)

'Debug.Print AUDIO

On Error GoTo EXIT_END

If Dir(AUDIO) = "" Then
    Me.WindowState = vbNormal
    'MsgBox "SOUND FILE 01# IS MISSING" + vbCrLf + AUDIO, vbMsgBoxSetForeground
End If




If AUDIO <> "" And Dir(AUDIO) <> "" Then
    
    'SOUND FILE SETUP AND LATER SAME ROUTINE PLAY
    '--------------------------------------------
    MMControl1.Command = "Stop"
    MMControl1.Command = "Close"
    MMControl1.Notify = True
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.Filename = AUDIO
'    Debug.Print Path_Of_Sound_File
    
    'Path_Of_Sound(1) = App.Path + "\Camera1a_2_8kHz.wav"
    'MMControl1.fileName = App.Path + "\Camera1a_2_8kHz.wav"
    'MMControl1.fileName = App.Path + "\Shot-12 Mix to Shot-18.wav"
    'MMControl1.fileName = App.Path + "\01 Woarble Tone 5 Mins.wav"
    MMControl1.Command = "Open"

End If


    
     If VARSOUND <> "NOTSOUND" Then
                
        If AUDIO <> "" Then
            'MMControl1.Command = "prev"
            MMControl1.Command = "Play"
        End If
    
    End If



EXIT_END:

End Sub




Private Sub Form_Unload(Cancel As Integer)
    
    MMControl1.Command = "Stop"
    MMControl1.Command = "Close"

End Sub

Private Sub Label1_Change()


Call Form_Resize



End Sub


Private Sub Label1_Click()
    MMControl1.Command = "Stop"
    MMControl1.Command = "Close"
    
    End
End Sub



'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub MNU_EDIT_LOGG_Click()

'vfile = App.Path + "\GoodSync_Logger.txt"

'Me.WindowState = vbMinimized

'ShellExecute Me.hwnd, "open", vfile, vbNullString, vbNullString, conSwNormal


End Sub

Private Sub MNU_EXIT_IN_Click()
End
End Sub


Private Sub Mnu_Explore_AUDIO_File_Click()

'Shell "Explorer.exe /select, """ + AUDIO + """", vbNormalFocus

Shell "Explorer.exe /E, """ + App.Path + "\AUDIO DATA\  """, vbMaximizedFocus

End Sub

'Private Sub MNU_RUN_NOW_Click()
'XX = 0
'End Sub





Private Sub Form_Resize()

Exit Sub


If Timer_KEY_ASYNC.Enabled = True Then
    
    If Me.WindowState = vbNormal And OWindowState = vbMinimized Then
    
    Timer_KEY_ASYNC.Enabled = False
    
    MMControl1.Command = "Stop"
    MMControl1.Command = "Close"
    
    End If
    
    OWindowState = Me.WindowState

End If


'If Timer2.Enabled = True Then
'    If FORM_LOAD_DONE = True Then
'        Timer2.Enabled = False
'
'    Timer_KEY_ASYNC.Enabled = False
'
'    MMControl1.Command = "Stop"
'    MMControl1.Command = "Close"
    
    
'    End If
'End If

'RESIZE_AT_LOAD = True
'
'If RESIZE_AT_LOAD = True And Me.WindowState <> vbMinimized Then
'    Me.WindowState = vbNormal
'    Call MNU_Norm_Click
'    RESIZE_AT_LOAD = False
'End If


If Me.WindowState <> vbNormal Then Exit Sub



'Text1.Left = 0
'Text1.Top = 0

On Error Resume Next ' - CRASH WHEN MINITURE

'Text1.Width = Form1.Width - 70
'Text1.Height = Form1.Height - (Menu_Height * Screen.TwipsPerPixelY) + 930


'If Me.Top + Me.Left + Me.Width + Me.Height = OLTLWH Then Exit Sub
'OLTLWH = Me.Top + Me.Left + Me.Width + Me.Height
'
'Timer4.Enabled = False
'Timer4.Interval = 40
'Timer4.Enabled = True

'If VARCENTER = False Then

'If VARCENTER = False Then
    'Label1.AutoSize = True

'    Label1.Width = Me.Width + 120
'    Label1.Height = Me.Height + 120

'End If
'End If
'If VARCENTER = True Then
'    Label1.Width = Me.Width + 120
'    Label1.Height = Me.Height + 120
'End If

'If VARCENTER = True Then Exit Sub

'On Error Resume Next
'    Me.Width = Label1.Width + 80
'    VARCENTER = False
    'Me.Height = (Label1.Height - Menu_Height * Screen.TwipsPerPixelY) + 2800
'Me.Height = Label1.Height + 2500

Me.Width = Label2.Width + 250
'XXDD = Form1.Height - (Menu_Height * Screen.TwipsPerPixelY) - 500 - Label1.Height
'If XXDD > 0 Then Text1.Height = XXDD - HEIGHT_ADJUST
Me.Height = (Menu_Height * Screen.TwipsPerPixelY) + Label2.Top + Label2.Height + 500 + 100



Me.Top = Screen.Height / 2 - Me.Height / 2 '+400
Me.Left = Screen.Width / 2 - Me.Width / 2 '+400

VARCENTER = True
 '   Label1.AutoSize = False




Exit Sub


Exit Sub


'NOT USED SEE ABOVE

On Error Resume Next
    Me.Width = Label2.Width + 120
'    VARCENTER = False
    Me.Height = Label2.Height + 780
If Me.WindowState <> vbNormal Then

    Exit Sub

End If

If VARCENTER = True Then Exit Sub
Me.Top = Screen.Height / 2 - Me.Height / 2 '+400
Me.Left = Screen.Width / 2 - Me.Width / 2 '+400
VARCENTER = True

End Sub


Private Sub Mnu_Fix1st_Letters_Click()

    
Dim EE As String


'EE = Text1.Text
EE = Label1.Caption
   
'If EGG = 0 Then EE = LCase(EE)
'EE = LCase(EE)

'Mid$(EE, 1, 1) = UCase(Mid$(EE, 1, 1))
'
For r = 65 To 65 + 25
'    EE = Replace(EE, " " + LCase(Chr(R)), " " + UCase(Chr(R)))
'    EE = Replace(EE, vbLf + LCase(Chr(R)), vbLf + UCase(Chr(R)))
'    EE = Replace(EE, vbCr + LCase(Chr(R)), vbCr + UCase(Chr(R)))
'    EE = Replace(EE, "-" + LCase(Chr(R)), "-" + UCase(Chr(R)))
'    EE = Replace(EE, "(" + LCase(Chr(R)), "(" + UCase(Chr(R)))
'    EE = Replace(EE, "." + LCase(Chr(R)), "." + UCase(Chr(R)))
'    EE = Replace(EE, "\" + LCase(Chr(R)), "\" + UCase(Chr(R)))
'    EE = Replace(EE, "_" + LCase(Chr(R)), "_" + UCase(Chr(R)))
'    EE = Replace(EE, """" + LCase(Chr(R)), """" + UCase(Chr(R)))
'    EE = Replace(EE, "." + LCase(Chr(R)), "." + UCase(Chr(R)))
    EE = Replace(EE, "--" + LCase(Chr(r)), vbCrLf + "--" + LCase(Chr(r)))
    EE = Replace(EE, "--" + UCase(Chr(r)), vbCrLf + "--" + UCase(Chr(r)))
Next



Label1.Caption = EE

'Clipboard.Clear
'Clipboard.SetText AD$

'Me.WindowState = 1

'Text1.Text = EE

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
    Beep
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
    'End
End Sub
Private Sub MNU_VB_FOLDER_Click()
    Beep
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + CODER_EXE_FILE_NAME_1 + """", vbNormalFocus
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

Private Sub Timer_BLUE_EMAIL_Timer()

If FindWindow(vbNullString, "SendSMTP_BLUETOOTH_BATCH") > 0 Then Exit Sub

If Command$ <> "" Then MsgBox Command$

Shell "C:\PStart\Progs\SendSMTP_v2.19.0.1\SendSMTP_BLUETOOTH_BATCH.BAT", vbMinimizedNoFocus

End

End Sub

Private Sub Timer_KEY_ASYNC_Timer()

i = 0
For i = 0 To 255
    GET_KEY = GetAsyncKeyState(i)
'    If GET_KEY < -300 Then GET_KEY = I: Debug.Print GET_KEY: Exit For



    'NOT F5 FOR TEST
    If GET_KEY < -300 And i <> 116 Then
        GET_KEY = i
        'Debug.Print GET_KEY
        Exit For
    End If
Next

'Exit Sub

DO_FLAG = False
'PAGE UP
If GET_KEY = 33 Then DO_FLAG = True
'ALT
If GET_KEY = 18 Then DO_FLAG = True
'MOUSE WHELL PUSH
If GET_KEY = 16 Then DO_FLAG = True
'MOUSE RIGHT CLICK
If GET_KEY = 2 Then DO_FLAG = True
'ESCAPE
If GET_KEY = 27 Then DO_FLAG = True

If DO_FLAG = False Then Exit Sub

Timer_KEY_ASYNC.Enabled = False

MMControl1.Command = "Stop"
MMControl1.Command = "Close"

End Sub

Private Sub Timer1_Timer()
    
    i = DateDiff("n", Now, START_TIME + TimeSerial(8, 0, 0))
    
    'MNU_EXIT_IN.Caption = "EXIT IN" + Str(XCount) + " Sexonds - Total Eight Minutes"
    MNU_EXIT_IN.Caption = "EXIT IN" + Str(i) + " Minute - Total 8 Hour Lenght"
    
    If START_TIME + TimeSerial(8, 0, 0) < Now Then
    
        End
        
    End If
    
    Exit Sub

    XCount = XCount - 1

    MNU_EXIT_IN.Caption = "EXIT IN" + Str(XCount) + " Sexonds - Total Eight Minutes"
    Call Form_Resize

    If XCount = 0 Then End
End Sub





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
    
    If GetMenuBarInfo(Me.hWnd, OBJID_MENU, 0, MenuInfo) Then
   
        With MenuInfo.rcBar
       
'            Debug.Print "Left: " & CStr(.Left)
'            Debug.Print "Right: " & CStr(.Right)
'            Debug.Print "Top: " & CStr(.Top)
'            Debug.Print "Bottom: " & CStr(.Bottom)
            Menu_Height = CStr(.Bottom) - CStr(.Top)
        End With
       
    End If
   
End Function
'------------------------------------------

Private Sub Timer2_Timer()



Me.WindowState = vbMinimized
Timer2.Enabled = False



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
    
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    
    WSHShell.Run """" + App.Path + "\RELOAD AND COPY.VBS" + """"
    Set WSHShell = Nothing
    Unload Me
    
End If

XVB_DATE = VB_DATE

End Sub




