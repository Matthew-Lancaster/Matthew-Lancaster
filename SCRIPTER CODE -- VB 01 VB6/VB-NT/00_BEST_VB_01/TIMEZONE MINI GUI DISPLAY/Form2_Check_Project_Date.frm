VERSION 5.00
Begin VB.Form Project_Check_Date 
   BackColor       =   &H00808080&
   Caption         =   "Form2"
   ClientHeight    =   3540
   ClientLeft      =   192
   ClientTop       =   540
   ClientWidth     =   9084
   LinkTopic       =   "Form2"
   ScaleHeight     =   3540
   ScaleWidth      =   9084
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer_FORM_LOAD_RUN_WAIT 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1560
      Top             =   1536
   End
   Begin VB.Timer TIMER_VB_PROJECT_CHECKDATE_FORM_LOAD 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3564
      Top             =   1548
   End
   Begin VB.Timer TIMER_RUN_AND_THEN_HIDE_PRE_EVENT 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3132
      Top             =   1092
   End
   Begin VB.Timer Timer_RUN_AND_THEN_HIDE_DELAY 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   2424
      Top             =   1596
   End
   Begin VB.Timer Timer_HIDE 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1908
      Top             =   960
   End
   Begin VB.Timer Timer_VB_PROJECT_CHECKDATE 
      Interval        =   10000
      Left            =   1296
      Top             =   768
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "CHECK PROJECT DATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   3816
      TabIndex        =   1
      Top             =   912
      Width           =   3828
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "CHECK PROJECT DATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   168
      TabIndex        =   0
      Top             =   132
      Width           =   3828
   End
End
Attribute VB_Name = "Project_Check_Date"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MATTHEW LANCASTER
'MATT.LAN@BTINTERNET.COM
'
'-----------------------------------------------------------
'SOME MORE CODE TO HERE
'NOW LOOKS AT A FILE WHICH HOLD ALL MY NETOWRK COMPUTER NAME
'AND DOES UPDATE AROUND NETWORK COMPLETELY VERY QUICKLY
'THERE WAS ABUG A CALL TO PROCEDURE IN FORM
'DOES NOT LOAD FORM FOR TIMER
'THAT IS DOES WHEN THE TIMER OBJECT IS SET WITH INTERVAL
'-----------------------------------------------------------
'THIS FILE IN SYNCED WITH SYCRONIZER
'-----------------------------------------------------------
'D:\VB6\VB-NT\00_Best_VB_01\10 SYNCRONIZE\SYNCRONIZER.exe
'D:\VB6\VB-NT\00_Best_VB_01\10 SYNCRONIZE\SYNCRONIZER.VBP
'-----------------------------------------------------------
'TIME
'Sat 19-May-2018 11:49:57
'Sat 19-May-2018 13:38:29 1 hour, 48 minutes and 32 seconds
'-----------------------------------------------------------

'TIME
'Sat 19-May-2018 21:02:19
'Sat 19-May-2018 23:00:00 2 hour
'NETWORK SYNCRONIZATION WORK __ CURED A NIGGELY BUG BY REMOVING ONE
'BEEP AT END
'-----------------------------------------------------------
Dim COUNT_TIMER_OVER_RIDE

Dim VB_APP_PATH_NAME_2 As String
Dim PATH_FILE_NAME1 As String
Dim PATH_FILE_NAME2 As String

' Private cProcesses As New clsCnProc

' Dim m_CRC

Dim APP_PATH_APP_EXENAME_EXE_PATH As String
Dim DATE_OF_APP_WITHER_VB_EXE_AT_LOAD_PATH As String
Dim CRC_OF_APP_WITHER_VB_EXE_AT_LOAD As String
Dim WSHShell

Dim SET_FORM_PROJECT_CHECK_DATE_SOME_INFO_TRIGGER_SHOW

Dim CRC_OF_APP_EXE_AT_LOAD
Dim DATE_OF_APP_EXE_AT_LOAD

Dim FORM_LOAD_CHECK_PROJECT_DATE_DO_AH_ONCE

Const hWnd_TOPMOST = -1
Const hWnd_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long


Dim APP_EXENAME_DATE
Dim FSO
Public EXIT_TRUE
Public CHECK_PROJECT_DATE_IN_PROCESS


Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Type FILETIME
   LowDateTime          As Long
   HighDateTime         As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes     As Long
   ftCreationTime       As FILETIME
   ftLastAccessTime     As FILETIME
   ftLastWriteTime      As FILETIME
   nFileSizeHigh        As Long
   nFileSizeLow         As Long
   dwReserved0          As Long
   dwReserved1          As Long
   cFileName            As String * 260  'MUST be set to 260
   cAlternate           As String * 14
End Type

Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1

'-----------------------------------------------------------------
Private Declare Function GetVersionExA Lib "kernel32" _
(lpVersionInformation As OSVERSIONINFO) As Integer

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Const PROCESS_CREATE_PROCESS = &H80
Private Const PROCESS_CREATE_THREAD = &H2
Private Const PROCESS_DUP_HANDLE = &H40
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const PROCESS_SET_QUOTA = &H100
Private Const PROCESS_SET_INFORMATION = &H200
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_VM_OPERATION = &H8
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_VM_WRITE = &H20

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Private Type PROCESS_INFORMATION
  hProcess    As Long
  hThread     As Long
  dwProcessID As Long
  dwThreadID  As Long
End Type

Private Type STARTUPINFO
  cb              As Long
  lpReserved      As String
  lpDesktop       As String
  lpTitle         As String
  dwX             As Long
  dwY             As Long
  dwXSize         As Long
  dwYSize         As Long
  dwXCountChars   As Long
  dwYCountChars   As Long
  dwFillAttribute As Long
  dwFlags         As Long
  wShowWindow     As Integer
  cbReserved2     As Integer
  lpReserved2     As Long
  hStdInput       As Long
  hStdOutput      As Long
  hStdError       As Long
End Type

Private Enum Priorities
  p_RealTime = &H100
  p_Hight = &H80
  p_Normal = &H20
  p_Idle = &H40
End Enum

Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadID As Long) As Long
Private Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, hModule As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Any, ByVal lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpExitCode As Long) As Long
Private Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Boolean

Private Declare Function CloseHandle Lib "kernel32" _
        (ByVal hObject As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Const TH32CS_SNAPPROCESS = &H2&



' Private Type FILETIME
'     dwLowDate  As Long
'     dwHighDate As Long
' End Type
 
 Private Type SYSTEMTIME
     wYear      As Integer
     wMonth     As Integer
     wDayOfWeek As Integer
     wDay       As Integer
     wHour      As Integer
     wMinute    As Integer
     wSecond    As Integer
     wMillisecs As Integer
 End Type
 
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_WRITE = &H40000000
  
Private Declare Function CreateFile Lib "kernel32" Alias _
   "CreateFileA" (ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   ByVal lpSecurityAttributes As Long, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) _
   As Long

Private Declare Function LocalFileTimeToFileTime Lib _
    "kernel32" (lpLocalFileTime As FILETIME, _
    lpFileTime As FILETIME) As Long

Private Declare Function SetFileTime Lib "kernel32" _
    (ByVal hFile As Long, ByVal MullP As Long, _
    ByVal NullP2 As Long, lpLastWriteTime _
    As FILETIME) As Long

Private Declare Function SystemTimeToFileTime Lib _
    "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime _
    As FILETIME) As Long
 
' Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long


Function GetWindowsVersion()
Dim OSInfo As OSVERSIONINFO
Dim RetValue As Integer
    OSInfo.dwOSVersionInfoSize = 148
    OSInfo.szCSDVersion = Space$(128)
    RetValue = GetVersionExA(OSInfo)
    sngWindowsVersion = CSng((CStr(OSInfo.dwMajorVersion) & "." & CStr(OSInfo.dwMinorVersion)))
    strWindowsVersion = CStr(OSInfo.dwMajorVersion) & "." & CStr(OSInfo.dwMinorVersion) & "." & CStr(OSInfo.dwBuildNumber)
    GetWindowsVersion = OSInfo.dwMajorVersion
    GetWindowsVersion = CSng((CStr(OSInfo.dwMajorVersion) & "." & CStr(OSInfo.dwMinorVersion)))
  
    '----------------------------------------------------
    'WINDOWS XP PROBLEM _ CHANGE THE SCRIPT HERE
    'NOT TO RUN AS ADMIN OR BRING THE RUNAS DIALOG BOX UP
    'WIN XP = 5.1 _ WINDOWS 10 = 6.2
    '----------------------------------------------------

End Function
'-----------------------------------------------------------------

Private Sub Form_Load()
    
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set WSHShell = CreateObject("WScript.Shell")
    Project_Check_Date.Caption = App.EXEName + ".EXE"
    
    If Me.EXIT_TRUE = True Then
        Unload Me
        Exit Sub
    End If
    
   
End Sub

Private Sub Timer_FORM_LOAD_RUN_WAIT_Timer()

'    Timer_FORM_LOAD_RUN_WAIT.Enabled = False
'    Timer_VB_PROJECT_CHECKDATE.Enabled = True
'    TIMER_VB_PROJECT_CHECKDATE_FORM_LOAD.Enabled = True

End Sub


Private Sub Form_Unload(Cancel As Integer)

' Form.EXIT_TRUE

End Sub

Private Sub Label1_Change()
    
    If Me.EXIT_TRUE = True Then
        Unload Me
        Exit Sub
    End If
    
    Dim r
    
    ZZ1 = Split(Label1.Caption, vbCrLf)
    For r = 1 To UBound(ZZ1)
        If InStr(Label2.Caption, ZZ1(r)) = 0 Then
            Label2.Caption = Label2.Caption + vbCrLf + ZZ1(r)
        End If
    Next
    
    SET_FORM_PROJECT_CHECK_DATE_SOME_INFO_TRIGGER_SHOW = True
    PR_WIDTH = Project_Check_Date.Label2.width + 100
    Project_Check_Date.Label2.FontSize = 15
    On Error Resume Next
    Project_Check_Date.width = PR_WIDTH
    Project_Check_Date.height = Project_Check_Date.Label2.height + 800
    Project_Check_Date.Left = Screen.width / 2 - (PR_WIDTH / 2)
    
    Project_Check_Date.Label2.Left = 0
    Project_Check_Date.Label2.Top = 100
    
    Project_Check_Date.BackColor = RGB(150, 200, 150)
    Project_Check_Date.Refresh
    Project_Check_Date.Show
    Project_Check_Date.SetFocus
    RESULT_API = AlwaysOnTop(Me.hWnd)

End Sub


Private Sub Timer_HIDE_Timer()

Project_Check_Date.Label1 = ""
Timer_HIDE.Enabled = False
Me.Hide

End Sub

Public Sub Timer_VB_PROJECT_CHECKDATE_Timer()


    COUNT_TIMER_OVER_RIDE = COUNT_TIMER_OVER_RIDE + 1

    If Me.EXIT_TRUE = True Then
        Unload Me
        Exit Sub
    End If
 
    Dim VBS_LAUNCHER_NAME
    
    READY_TO_GO = False
    
    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32
    
    If APP_EXENAME_DATE = 0 Then
        PATH_FILE_NAME1 = App.Path + "\" + App.EXEName + ".EXE"
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set F = FSO.GetFile(PATH_FILE_NAME1)
        APP_EXENAME_DATE = F.DateLastModified
    End If
    
    PATH_FILE_NAME1 = App.Path + "\" + App.EXEName + ".EXE"
    PATH_FILE_NAME2 = Replace(PATH_FILE_NAME1, "D:\VB6\", "D:\VB6-EXE\")
    '---------------------------------------------------
    On Error Resume Next
    If Dir(PATH_FILE_NAME2) = "" Then
        If Dir(Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\")), vbDirectory) = "" Then
            ' CREATE THE VB-EXE FOLDER IF NOT THERE
            ' -------------------------------------
            CreateFolderTree Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\"))
        End If
    End If
    
    Err.Clear
    Set F1 = FSO.GetFile(PATH_FILE_NAME1)
    Set F2 = FSO.GetFile(PATH_FILE_NAME2)
    VB_EXE_DATE = F1.DateLastModified
    APP_EXENAME_DATE = F2.DateLastModified
    VB_EXE_SIZE = F1.Size
    APP_EXENAME_SIZE = F2.Size
    
    ' IS SAME DATE AND SAME SIZE -- EXIT
    ' IF MINE CHANGES THEN CHECK THE OTHER NETWORK SHARE
    ' ---------------------------------
    If COUNT_TIMER_OVER_RIDE < 10 Then
    If APP_EXENAME_DATE = VB_EXE_DATE And APP_EXENAME_SIZE = VB_EXE_SIZE Then
        Exit Sub
    End If
    End If
    
    ABLE_TO_EXIT = False
    
    
    ' ---------------------------------------------------------------------
    COUNT_TIMER_OVER_RIDE = 0
    ' ---------------------------------------------------------------------
    ' DOUBLE CHECK PERIODICLY TRY TO COPY FILE TO DESTINATION VB_EXE FOLDER
    ' ---------------------------------------------------------------------
   
    
    ' CAN READ IT SELF NOT WRITE TO SELF SO COPY OUT TO VB_EXE FOLDER IF DIFFERNT
    If APP_EXENAME_DATE < VB_EXE_DATE Then
        FSO.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
        ABLE_TO_EXIT = True
    End If
        
    PATH_FILE_NAME4 = "\\8-msi-gp62m-7rd\8_msi_gp62m_7rd_02_d_drive\" + Right(App.Path, Len(App.Path) - 3) + "\" + App.EXEName + ".EXE"
    PATH_FILE_NAME5 = Replace(PATH_FILE_NAME4, "\VB6\", "\VB6-EXE\")
    PATH_FILE_NAME4 = PATH_FILE_NAME1
    On Error Resume Next
    If Dir(PATH_FILE_NAME4) = "" And Dir(Mid(PATH_FILE_NAME4, 1, InStrRev(PATH_FILE_NAME4, "\")), vbDirectory) = "" Then
         CreateFolderTree Mid(PATH_FILE_NAME4, 1, InStrRev(PATH_FILE_NAME4, "\"))
    End If
    Set F1 = FSO.GetFile(PATH_FILE_NAME4)
    Set F2 = FSO.GetFile(PATH_FILE_NAME5)
    VB_EXE_DATE = F1.DateLastModified
    APP_EXENAME_DATE = F2.DateLastModified
    If APP_EXENAME_DATE < VB_EXE_DATE Then FSO.CopyFile PATH_FILE_NAME4, PATH_FILE_NAME5
    
    PATH_FILE_NAME4 = "\\9-asus-g815lm\9_asus_g815lm_02_d_drive\" + Right(App.Path, Len(App.Path) - 3) + "\" + App.EXEName + ".EXE"
    PATH_FILE_NAME5 = Replace(PATH_FILE_NAME4, "\VB6\", "\VB6-EXE\")
    PATH_FILE_NAME4 = PATH_FILE_NAME1
    On Error Resume Next
    If Dir(PATH_FILE_NAME4) = "" And Dir(Mid(PATH_FILE_NAME4, 1, InStrRev(PATH_FILE_NAME4, "\")), vbDirectory) = "" Then
         CreateFolderTree Mid(PATH_FILE_NAME4, 1, InStrRev(PATH_FILE_NAME4, "\"))
    End If
    Set F1 = FSO.GetFile(PATH_FILE_NAME4)
    Set F2 = FSO.GetFile(PATH_FILE_NAME5)
    VB_EXE_DATE = F1.DateLastModified
    APP_EXENAME_DATE = F2.DateLastModified
    If APP_EXENAME_DATE < VB_EXE_DATE Then FSO.CopyFile PATH_FILE_NAME4, PATH_FILE_NAME5
    
    ' \\8-msi-gp62m-7rd\8_msi_gp62m_7rd_02_d_drive\ ' VB6\VB-NT\00_BEST_VB_01
    ' \\9-asus-g815lm\9_asus_g815lm_02_d_drive\VB6\ ' VB-NT\00_BEST_VB_01
   
    ' COPIER DONE
    ' DON'T RELAUNCH PROGRAM UNLESS COPY WAS FROM VB_EXE FOLDER
    ' ---------------------------------------------------------
    If ABLE_TO_EXIT = True Then Exit Sub
        
        

    
    
    
    WxHex_FILE_2 = Hex(m_CRC.CalculateFile(PATH_FILE_NAME2))
    WxHex_FILE_1 = Hex(m_CRC.CalculateFile(PATH_FILE_NAME1))
    
    If WxHex_FILE_1 = WxHex_FILE_2 Then
        Exit Sub
    End If

    VBS_LAUNCHER_NAME = """" + "C:\SCRIPTER\SCRIPTER CODE -- VB 02 VBSCRIPT\VBS 40-RUN EXE 2 ARG_QUIET.VBS" + """" + " "
    ' THIS IS THE PARAM ARGUMENT -- THIS IS THE FILE TO BE LAUNCH
    VBS_LAUNCHER_NAME = VBS_LAUNCHER_NAME + """" + Replace(PATH_FILE_NAME1, " ", "*") + """" + " "
    ' THIS IS THE 1ST FILE TO COMPARE DATE AND CRC
    VBS_LAUNCHER_NAME = VBS_LAUNCHER_NAME + """" + Replace(PATH_FILE_NAME1, " ", "*") + """" + " "
    ' THIS IS THE 2ND FILE TO COMPARE DATE AND CRC
    VBS_LAUNCHER_NAME = VBS_LAUNCHER_NAME + """" + Replace(PATH_FILE_NAME2, " ", "*") + """"

    WSHShell.RUN VBS_LAUNCHER_NAME, ShowWindow_2, DontWaitUntilFinished
    
    Me.EXIT_TRUE = True
    Dim Form As Form
    On Error Resume Next
    For Each Form In Forms
        Form.EXIT_TRUE = Me.EXIT_TRUE
    Next Form
    For Each Form In Forms
        Err.Clear
        If Form.EXIT_TRUE = True Then
            If Err.Number = 0 Then
                EXIT_TRUE = True
                Exit For
            End If
        End If
    Next Form
    
    Dim i, Control
    'SET ALL TIMERS IN ALL FORMS ENABLED TO =FALSE
    On Error Resume Next
        For i = 0 To Forms.Count - 1
            For Each Control In Forms(i).Controls
                If InStr(UCase(Control.Name), "TIMER") > 0 Then
                    'Debug.Print Control.Name
                    Control.Enabled = False
                    DoEvents
                End If
            Next
        Next i
    On Error GoTo 0
    
    For r = 1 To 100
        DoEvents
    Next
    
    For Each Form In Forms
        Unload Form
    Next Form

    For r = 1 To 100
        DoEvents
    Next
    
    ' End

End Sub


Public Sub DATE_OF_APP_EXE_AT_LOAD_SUB()

End Sub



Public Sub DATE_OF_APP_WITHER_VB_EXE_AT_LOAD_SUB()
    
    Exit Sub
    
    ' WHY DATE WRONG AT LOAD
    ' ONLY GET DATE NOT LAUNCHER
    ' Exit Sub -- LATER IN MODULE
    
    Dim F1, F2
    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32

    If DATE_OF_APP_WITHER_VB_EXE_AT_LOAD_PATH = "" Then
        PATH_FILE_NAME1 = App.Path + "\" + App.EXEName + ".EXE"
        DATE_OF_APP_WITHER_VB_EXE_AT_LOAD_PATH = Replace(PATH_FILE_NAME1, "D:\VB6\", "D:\VB6-EXE\")
    End If
    If DATE_OF_APP_WITHER_VB_EXE_AT_LOAD = 0 Then
        Set F = FSO.GetFile(DATE_OF_APP_WITHER_VB_EXE_AT_LOAD_PATH)
        DATE_OF_APP_WITHER_VB_EXE_AT_LOAD = F.DateLastModified
    End If
    If DATE_OF_APP_EXE_AT_LOAD <> 0 Then
        If DATE_OF_APP_WITHER_VB_EXE_AT_LOAD <> 0 Then
            On Error Resume Next
            Set F1 = FSO.GetFile(APP_PATH_APP_EXENAME_EXE_PATH)
            Set F2 = FSO.GetFile(DATE_OF_APP_WITHER_VB_EXE_AT_LOAD)
            If F2.DateLastModified > F1.DateLastModified Then
                WxHex_FILE_1 = Hex(m_CRC.CalculateFile(APP_PATH_APP_EXENAME_EXE_PATH))
                If WxHex_FILE_1 <> CRC_OF_APP_EXE_AT_LOAD Then
                    If Err.Number = 0 Then
                        
                        ' KILL THE WSCRIPT NAME
                        ' FUNCTION & INCLUDE SUB
                        ' FIND_SCRIPTNAME_VAR = "\VBS - RELOAD AND COPY_"
                        ' -----------------------------------------------
                        Call WSCRIPT_SCRIPTNAME_RELOAD_KILLER
                        ' -------------------------------------------
                        APP_PATH_AND_EXE = App.Path + "\" + App.EXEName + ".EXE"
                        APP_PATH_AND_EXE = Replace(APP_PATH_AND_EXE, " ", "*")
                        VBS_LAUNCHER_NAME = "C:\SCRIPTER\SCRIPTER CODE -- VB 02 VBSCRIPT\VBS 40-RUN EXE 2 ARG_QUIET.VBS"
                        WSHShell.RUN """" + VBS_LAUNCHER_NAME + """" + " " + APP_PATH_AND_EXE, ShowWindow_2, DontWaitUntilFinished
                        End
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub TIMER_VB_PROJECT_CHECKDATE_FORM_LOAD_TIMER()
    
If Me.EXIT_TRUE = True Then
    Unload Me
    Exit Sub
End If
    
Call VB_PROJECT_CHECKDATE

Exit Sub

'TIMER_VB_PROJECT_CHECKDATE_FORM_LOAD.Enabled = False
'Timer_VB_PROJECT_CHECKDATE.Enabled = True

End Sub


Public Sub VB_PROJECT_CHECKDATE(Optional FORM_LOAD_VAR)
                


End Sub

Private Sub Timer_RUN_AND_THEN_HIDE_DELAY_Timer()
'    Timer_RUN_AND_THEN_HIDE_DELAY.Enabled = False
'
'    NotAlwaysOnTop Me.hWnd
'    Me.Visible = False
End Sub


Function ControlExists(ControlName As String, FormCheck As Form) As Boolean
    Dim strTest As String
    On Error Resume Next
    strTest = FormCheck(ControlName).Name
    ControlExists = (Err.Number = 0)
End Function

Sub ControlCall_Find(ControlName As String)

    On Error Resume Next
    Dim SUBNAME As String
    Dim Form As Form
    For Each Form In Forms
        Err.Clear
        strTest = Form(ControlName).Name
        ControlExistsAnswer = (Err.Number = 0)
        
        If ControlExistsAnswer = True Then
            Form.CHECK_PROJECT_DATE_IN_PROCESS = Me.CHECK_PROJECT_DATE_IN_PROCESS
            SUBNAME = ControlName + "_" + Mid(ControlName, 1, InStr(ControlName, "_") - 1)
            CallByName Form, SUBNAME, VbMethod
        End If
    Next
    On Error GoTo 0

End Sub


Function ControlCall(ControlName As String, FormCheck As Form) As Boolean
    Dim strTest As String, SUBNAME As String
    On Error Resume Next
    strTest = FormCheck(ControlName).Name
    ControlCall = (Err.Number = 0)
    If ControlCall = True Then
        ' ---------------------------------------------
        ' NEW LEARN TODAY 2ND ONE
        ' 1ST FIND FORM EXIST
        ' AND NOW CallByName TO RUN A ROUTINE BY STRING
        ' THAT MAKE TOTAL EVERY LEARN ABOUT CONTROL ARRAY AH
        ' ---------------------------------------------
        ' [ Saturday 10:40:20 Am_31 August 2019 ]
        ' ---------------------------------------------
        ' MAKE SURE CALL TO SUBROUTINE IS ABLE PUBLIC
        ' FROM ANOTHER FORM
        ' HA HA
        ' ---------------------------------------------
        ' [ Saturday 10:50:00 Am_31 August 2019 ]
        ' ---------------------------------------------
        SUBNAME = ControlName + "_" + Mid(ControlName, 1, InStr(ControlName, "_") - 1)
        CallByName FormCheck, SUBNAME, VbMethod
        'CallByName FormCheck(ControlName).Name, ControlName, VbLet
    End If
    
    ' ----
    ' Calling a Property or Method Using a String Name (Visual Basic) | Microsoft Docs
    ' https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/early-late-binding/calling-a-property-or-method-using-a-string-name
    ' --------
    ' VB Helper: HowTo: Call a subroutine by name in VB6
    ' https://www.vb-helper.com/howto_call_routine_by_name.html
    ' ----
    ' Note
    '
    ' While the CallByName function may be useful in some cases,
    ' you must weigh its usefulness
    ' against the performance implications — using CallByName to invoke a
    ' procedure is slightly slower than a late-bound call. If you are
    ' invoking a function that is called repeatedly, such as inside a loop,
    ' CallByName can have a severe effect on performance.


    'CallByName FormCheck(ControlName).Name, ControlName, VbLet
    'RESULT = CallByName(text1, "MousePointer", VbGet)
    'CallByName text1, "Move", VbMethod, 100, 100

    'Dim Math As New MathClass
    'Me.TextBox1.Text = CStr(CallByName(Math, Me.TextBox2.Text, Microsoft.VisualBasic.CallType.Method, TextBox1.Text))
    
    'Dim text1 As TextBox = CType(Me.Controls("text_1"), TextBox)
    
    'FormCheck.
    'Me.Controls.Find("control name", True).FirstOrDefault()
    'Dim oObj As Object ( CType(FormCheck.Controls("NAME_OF_CONTROL"), Control))
    'Obj.Property = Value
End Function

Public Function CreateFolderTree(ByVal sPath As String) As Boolean
    Dim nPos As Integer

    If Mid(sPath, Len(sPath), 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)

    On Error GoTo CreateFolderTreeError
    
    nPos = InStr(sPath, "\")
    While nPos > 0
        If Not FolderExists(Left$(sPath, nPos - 1)) Then
            MkDir Left$(sPath, nPos - 1)
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
    If Not FolderExists(sPath) Then MkDir sPath
    
    CreateFolderTree = True
    Exit Function

CreateFolderTreeError:
    Exit Function
End Function


Private Function FolderExists(sFolder As String) As Boolean
    '##############################################################################################
    'Returns True if the specified folder exists
    '##############################################################################################
    
    Dim WFD As WIN32_FIND_DATA
    Dim lResult As Long
    
    lResult = FindFirstFile(sFolder, WFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            FolderExists = True
        Else
            FolderExists = False
        End If
    End If
End Function


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

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
   Debug.Assert Not TestIDE(IsIDE)
'  TESTING
'  IsIDE = False
End Function
Private Function TestIDE(Test As Boolean) As Boolean
   Test = True
End Function
'***********************************************

Private Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Private Function NotAlwaysOnTop(ByVal hWnd As Long)
    Dim flags
    SetWindowPos hWnd, hWnd_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function

Sub WSCRIPT_SCRIPTNAME_RELOAD_KILLER()
    
    Dim objWMIService
    Dim colProcesses
    Dim i1
    Dim i2
    Dim objProcess
    Dim strScriptName
    Dim PID_Script As Long
    
    ' FUNCTION & INCLUDE SUB
    FIND_SCRIPTNAME_VAR = "\VBS - RELOAD AND COPY_"
    ' ---------------------------------------------
    
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process")
    
    i1 = 0 ' ANY PROGRAM WSCRIPT
    i2 = 0 ' MY  PROGRAM WSCRIPT
    For Each objProcess In colProcesses
        If Not (IsNull(objProcess.commandLine)) Then
            strScriptName = objProcess.commandLine
            If InStr(strScriptName, FIND_SCRIPTNAME_VAR) > 0 Then
                PID_Script = objProcess.ProcessID
                If PID_Script > 0 Then
                    Call Process_Kill(PID_Script)
                End If
            End If
        End If
    Next
End Sub

Public Sub Process_Kill(P_ID As Long)
    '// Kill the wanted process
    
    Dim hProcess As Long
    Dim lExitCode As Long
    'NORMAL_PRIORITY_CLASS
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, P_ID)
    'If hProcess = 0 Then Call Err_Dll(Err.LastDllError, "OpenProcess failed", sLocation, "Kill_Process")
    
    If GetExitCodeProcess(hProcess, lExitCode) = False Then
    'Call Err_Dll(Err.LastDllError, "GetExitCodeProcess failed", sLocation, "Kill_Process")
    End If
    If TerminateProcess(hProcess, lExitCode) = False Then
    'Call Err_Dll(Err.LastDllError, "TerminateProcess failed", sLocation, "Kill_Process")
    End If
    If CloseHandle(hProcess) = False Then
    'Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Kill_Process")
    End If
End Sub


Function FIND_SCRIPTNAME(PID_TEST)

Dim objWMIService
Dim colProcesses
Dim i1
Dim i2
Dim objProcess
Dim strScriptName
Dim PID_Script
FIND_SCRIPTNAME = ""

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process")

i1 = 0 ' ANY PROGRAM WSCRIPT
i2 = 0 ' MY  PROGRAM WSCRIPT
For Each objProcess In colProcesses
    If Not (IsNull(objProcess.commandLine)) Then
        'strScriptName = Trim(Right(objProcess.CommandLine, Len(objProcess.CommandLine) - InStrRev(objProcess.CommandLine, "\")))
        ' strScriptName = Left(strScriptName, Len(strScriptName) - 1)
        strScriptName = objProcess.commandLine
        strScriptName = Replace(strScriptName, """", "")
        PID_Script = objProcess.ProcessID
        If Val(PID_Script) = Val(PID_TEST) Then
            FIND_SCRIPTNAME = strScriptName
            Exit Function
        End If
    End If
Next

End Function
Function WSCRIPT_SCRIPTNAME(PID_TEST)

Dim objWMIService
Dim colProcesses
Dim i1
Dim i2
Dim objProcess
Dim strScriptName
Dim PID_Script
WSCRIPT_SCRIPTNAME = ""

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("select * from win32_process where name = 'wscript.exe'")

i1 = 0 ' ANY PROGRAM WSCRIPT
i2 = 0 ' MY  PROGRAM WSCRIPT
For Each objProcess In colProcesses
    If Not (IsNull(objProcess.commandLine)) Then
        strScriptName = Trim(Right(objProcess.commandLine, Len(objProcess.commandLine) - InStrRev(objProcess.commandLine, "\")))
        strScriptName = Left(strScriptName, Len(strScriptName) - 1)
        PID_Script = objProcess.ProcessID
        If Val(PID_Script) = Val(PID_TEST) Then
            WSCRIPT_SCRIPTNAME = strScriptName
            Exit Function
        End If
    End If
Next

End Function


Public Function SetFileDateTime(ByVal Filename As String, _
  ByVal TheDate As String) As Boolean
'************************************************
'PURPOSE:    Set File Date (and optionally time)
'            for a given file)
         
'PARAMETERS: TheDate -- Date to Set File's Modified Date/Time
'            FileName -- The File Name

'Returns:    True if successful, false otherwise
'************************************************
If Dir(Filename) = "" Then Exit Function
If Not IsDate(TheDate) Then Exit Function

Dim lFileHnd As Long
Dim lRet As Long

Dim typFileTime As FILETIME
Dim typLocalTime As FILETIME
Dim typSystemTime As SYSTEMTIME

With typSystemTime
    .wYear = Year(TheDate)
    .wMonth = Month(TheDate)
    .wDay = Day(TheDate)
    .wDayOfWeek = Weekday(TheDate) - 1
    .wHour = Hour(TheDate)
    .wMinute = Minute(TheDate)
    .wSecond = Second(TheDate)
End With

lRet = SystemTimeToFileTime(typSystemTime, typLocalTime)
lRet = LocalFileTimeToFileTime(typLocalTime, typFileTime)

lFileHnd = CreateFile(Filename, GENERIC_WRITE, _
    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, _
    OPEN_EXISTING, 0, 0)
    
lRet = SetFileTime(lFileHnd, ByVal 0&, ByVal 0&, _
         typFileTime)

CloseHandle lFileHnd
SetFileDateTime = lRet > 0

End Function




Sub SOME_INFO()

    ' -------------------------------------------------
    ' FORM2_CHECK_POJECT_DATE
    ' ----------------------------
    ' NOTHING REQUIRE CALL IN HERE
    ' ----------------------------
    ' FORM START AND FORM_LOAD RUN
    ' TRIGGER BY TIMER BEGIN
    ' ----------------------------
    ' Timer_VB_PROJECT_CHECKDATE.Interval = 1000
    ' -------------------------------------------------
    
    
    ' -------------------------------------------------
    ' USE THIS LINE SET IN FORM_LOAD OF MAIN START PROJECT
    ' AT THE BEGINNING
    ' -------------------------------------------------
    ' -------------------------------------------------
    ' IF WANT RENAME FORM FOR DIFFERENT PROJECT
    ' & ALSO KEEP UNIVERSAL FOR SYNC PURPOSE
    ' UNABLE TO RENAME FORM IN ANYWAY AND SYNC
    ' ANSWER CHANGE NAME WITHOUT FRM_ PRECEDE
    ' SO LAND AT TOP OF FORM LISTER NOT IN THE WAY
    ' SOME PROJECT HAVE FRM_ OR FORM_ SORT ORDER
    ' -------------------------------------------------
    ' 03 OF 03
    ' -------------------------------------------------
    ' DETECT PRESENCE OF FORM
    ' REQUEST ANSWER WILL LOAD THE FORM ALSO
    ' -------------------------------------------------
    ' On Error Resume Next
    ' Dim frmX As Form, FrmxName_T As String
    ' Dim FrmxName() As String
    ' ReDim Preserve FrmxName(10)
    ' i = 0
    ' i = i + 1: FrmxName(i) = "Project_Check_Date"  ' PREVIOUS  NAME USER
    ' i = i + 1: FrmxName(i) = "Frm_Project_Check_Date"    ' ATTEMPTED NAME USER
    ' i = i + 1: FrmxName(i) = "Project_Check_Date"        ' STANDARDISE
    ' ReDim Preserve FrmxName(i)
    ' For i = 1 To UBound(FrmxName)
    '     Err.Clear
    '     Set frmX = Forms.Add(FrmxName(i))
    '     If Err.Number = 0 Then
    '         FrmxName_T = FrmxName(i)
    '         Exit For
    '     End If
    ' Next
    ' -------------------------------------------------
    
    
    
    ' -------------------------------------------------
    ' TEST INFO -- ALL OVER HERE NOT READ FURTHER
    ' -------------------------------------------------
    ' IF WANT RENAME FORM FOR DIFFERENT PROJECT
    ' & ALSO KEEP UNIVERSAL FOR SYNC PURPOSE
    ' -------------------------------------------------
    ' ---------------------------------------
    ' 01 OF 03
    ' FIND FOR SOMETHING
    ' BUT ERROR THIS EXAMPLE RETURN BAD ERROR
    ' EVEN THROUGH ERROR HANDLING ON
    ' LOOK BEYOND REM FOR PROPER EXAMPLE
    ' ---------------------------------------
    ' Dim Form As Form
    ' For Each Form In Forms
    '    ' NONE FORM BEEN LOADER YET SO NOT SHOW HERE
    '    ' HOW CHECK FORM EXIST WITHOUT LOADER
    '    ' ------------------------------------------
    '    'If InStr(Form.Name, "Project_Check_Date") Then
    '        Call Form.VB_PROJECT_CHECKDATE("FORM LOAD")
    '    'End If
    '    'If InStr(Form.Name, "Frm_Project_Check_Date") Then
    '        Call Form.VB_PROJECT_CHECKDATE("FORM LOAD")
    '    'End If
    ' Next
    ' On Error GoTo 0

    ' 02 OF 03
    ' BIT OF CODE & ONLY FOR EXCEL
    ' ----------------------------
    ' On Error Resume Next
    ' Dim projectIndex As Integer
    ' Dim formName As String
    ' projectIndex = 2
    ' formName = "Project_Check_Date"
    ' UserFormExists = (Application.VBE.vbProjects(projectIndex).VBComponents(formName).Name = formName)
    ' On Error GoTo 0
    
    ' -------------------------------------------------
    ' IF WANT RENAME FORM FOR DIFFERENT PROJECT
    ' & ALSO KEEP UNIVERSAL FOR SYNC PURPOSE
    ' -------------------------------------------------
    ' 03 OF 03
    ' -------------------------------------------------
    ' RESULT GIVE SEE ABOVE
    ' -------------------------------------------------


End Sub


Sub SOME_INFO_02()

            ' STYLE BEFORE
'            Dim WSHShell
'            Set WSHShell = CreateObject("WScript.Shell")
'                WSHShell.Run """" + VBS_LAUNCHER_NAME + """", 1, False
'            Set WSHShell = Nothing
    
    '    If GetWindowsVersion <= 5.1 Then
    '        MsgBox "XP"
    '        Dim WSHShell
    '        Set WSHShell = CreateObject("WScript.Shell")
    '            WSHShell.Run """" + VBS_LAUNCHER_NAME + """", 1, False
    '        Set WSHShell = Nothing
            
            ' -------------------------------------------------------------------------------
            ' Shell VBS_LAUNCHER_NAME ____ NOT WORKING
            ' -------------------------------------------------------------------------------
            ' Eample of using ShellExecute
            ' ShellExecute hwnd, "open", pathandfile, vbNullString, vbNullString, conSwNormal
            ' -------------------------------------------------------------------------------
            'ShellExecute Me.hWnd, "open", VBS_LAUNCHER_NAME, vbNullString, vbNullString, conSwNormal
            
    '        Set oRunas = CreateObject("runas.clsrunas", GetComputerName)
    '        With oRunas
    '                .sDomain = "WORKGROUP"
    '                .sUserName = GetUserName
    '                .sPassword = " "
    '                .sCommand = "C:\Program Files\Olympus\DSSPlayer2002\DictWnd.exe"
    '        '       .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
    '        '             you must explictly use the relevant program for example if you wanted
    '        '             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
    '        '             you would have to use
    '        '             "c:\winnt\notepad.exe c:\text.txt"
    '                .RunAs 'Call the Run As method
    '        End With
    '        Set oRunas = Nothing
    '    End If


End Sub




