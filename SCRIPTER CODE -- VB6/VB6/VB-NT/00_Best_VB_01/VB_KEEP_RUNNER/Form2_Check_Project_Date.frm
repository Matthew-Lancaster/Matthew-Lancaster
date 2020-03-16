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
      Left            =   1536
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
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1296
      Top             =   792
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

Dim ONCE_STARTER_MODE


Dim VB_APP_PATH_NAME_2 As String
Dim PATH_FILE_NAME1 As String
Dim PATH_FILE_NAME2 As String

' Private cProcesses As New clsCnProc

' Dim m_CRC

Dim APP_PATH_APP_EXENAME_EXE_PATH As String
Dim DATE_OF_APP_WITHER_VB_EXE_AT_LOAD_PATH As String
Dim CRC_OF_APP_WITHER_VB_EXE_AT_LOAD As String
Dim WSHShell

Dim RZZD
Dim DISPLAY_RESULT_VAR_AFTER_FIRST_PASS

Dim SET_FORM_PROJECT_CHECK_DATE_SOME_INFO_TRIGGER_SHOW

Dim CRC_OF_APP_EXE_AT_LOAD
Dim DATE_OF_APP_EXE_AT_LOAD

Dim FORM_LOAD_CHECK_PROJECT_DATE_DO_AH_ONCE

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "Kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Dim XVB_DATE_SYNC_VB_PROJECT
Dim XVB_DATE
Dim APP_EXENAME_DATE
Dim FSO
Public EXIT_TRUE
Public CHECK_PROJECT_DATE_IN_PROCESS


Private Declare Function FindFirstFile Lib "Kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "Kernel32" (ByVal hFindFile As Long) As Long

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
Private Declare Function GetVersionExA Lib "Kernel32" _
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
  dwProcessId As Long
  dwThreadId  As Long
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

Private Declare Function Process32First Lib "Kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "Kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Private Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function TerminateProcess Lib "Kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, hModule As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Any, ByVal lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpExitCode As Long) As Long
Private Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Boolean

Private Declare Function CloseHandle Lib "Kernel32" _
        (ByVal hObject As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "Kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
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
  
Private Declare Function CreateFile Lib "Kernel32" Alias _
   "CreateFileA" (ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   ByVal lpSecurityAttributes As Long, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) _
   As Long

Private Declare Function LocalFileTimeToFileTime Lib _
    "Kernel32" (lpLocalFileTime As FILETIME, _
    lpFileTime As FILETIME) As Long

Private Declare Function SetFileTime Lib "Kernel32" _
    (ByVal hFile As Long, ByVal MullP As Long, _
    ByVal NullP2 As Long, lpLastWriteTime _
    As FILETIME) As Long

Private Declare Function SystemTimeToFileTime Lib _
    "Kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime _
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

    Timer_FORM_LOAD_RUN_WAIT.Enabled = False
    Timer_VB_PROJECT_CHECKDATE.Enabled = False
    TIMER_VB_PROJECT_CHECKDATE_FORM_LOAD.Enabled = True

End Sub


Private Sub Form_Unload(Cancel As Integer)

' Form.EXIT_TRUE

End Sub

Private Sub Label1_Change()
    
    If Me.EXIT_TRUE = True Then
        Unload Me
        Exit Sub
    End If
    
    If RZZD = 0 Then
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
    PR_WIDTH = Project_Check_Date.Label2.Width + 100
    Project_Check_Date.Label2.FontSize = 15
    On Error Resume Next
    Project_Check_Date.Width = PR_WIDTH
    Project_Check_Date.Height = Project_Check_Date.Label2.Height + 800
    Project_Check_Date.Left = Screen.Width / 2 - (PR_WIDTH / 2)
    
    Project_Check_Date.Label2.Left = 0
    Project_Check_Date.Label2.Top = 100
    
    Project_Check_Date.BackColor = RGB(150, 200, 150)
    Project_Check_Date.Refresh
    Project_Check_Date.Show
    Project_Check_Date.SetFocus
    RESULT_API = AlwaysOnTop(Me.hWnd)

End Sub


Public Sub Timer_VB_PROJECT_CHECKDATE_Timer()

    If Me.EXIT_TRUE = True Then
        Unload Me
        Exit Sub
    End If
    
    If ONCE_STARTER_MODE = True Then
    Exit Sub
    End If
    ONCE_STARTER_MODE = True
    
    ' --------------------------------------------------
    Call VB_PROJECT_CHECKDATE
    Call DATE_OF_APP_WITHER_VB_EXE_AT_LOAD_SUB
    Call DATE_OF_APP_EXE_AT_LOAD_SUB
    ' --------------------------------------------------


End Sub


Public Sub DATE_OF_APP_EXE_AT_LOAD_SUB()

    ' WHY DATE WRONG AT LOAD
    ' ONLY GET DATE NOT LAUNCHER
    ' Exit Sub -- LATER IN MODULE
    

    Set m_CRC = New clsCRC
    m_CRC.Algorithm = CRC32

    If APP_PATH_APP_EXENAME_EXE_PATH = "" Then
        PATH_FILE_NAME1 = App.Path + "\" + App.EXEName + ".EXE"
        APP_PATH_APP_EXENAME_EXE_PATH = PATH_FILE_NAME1
    End If
    If DATE_OF_APP_EXE_AT_LOAD = 0 Then
        Set F = FSO.GetFile(APP_PATH_APP_EXENAME_EXE_PATH)
        DATE_OF_APP_EXE_AT_LOAD = F.DateLastModified
    End If
    
    If CRC_OF_APP_EXE_AT_LOAD = "" Then
        'APP_PATH_APP_EXENAME_EXE_PATH
        'VB_APP_PATH_NAME_2
        If FSO.FolderExists(VB_APP_PATH_NAME_2) = True Then
        'If Dir(VB_APP_PATH_NAME_2) <> "" Then
            If FSO.FileExists(VB_APP_PATH_NAME_2) Then
            WxHex_FILE_2 = Hex(m_CRC.CalculateFile(VB_APP_PATH_NAME_2))
            End If
        End If
        WxHex_FILE_1 = Hex(m_CRC.CalculateFile(APP_PATH_APP_EXENAME_EXE_PATH))
        
        CRC_OF_APP_EXE_AT_LOAD = WxHex_FILE_1
    End If
    
'    Exit Sub
    
    If DATE_OF_APP_EXE_AT_LOAD <> 0 Then
        On Error Resume Next
        Set F = FSO.GetFile(APP_PATH_APP_EXENAME_EXE_PATH)
        
        If DATE_OF_APP_EXE_AT_LOAD < F.DateLastModified Then
            ' --------------------------------------------------
            ' THE CRC WAS IMPORTANT FOR WINDOWS-XP COMPUTER PAIR
            ' NOT SEE WHY IS DATE WRONG
            ' HERE LOOK AT OWN EXE FOR CONTENT SAME AS WHEN BOOT AR
            ' LIKE FOR INSTANCE BUT NOT ALWAYS POSIABLE
            ' THE EXE IS OVERWRITE WITH NEW VERSION
            ' UNDENETH ONE ALRREADY IN PLAY
            ' ONLY IF ACCESS GIVEN -- LOT TIME NOT ALLOW AR
            ' --------------------------------------------------
            If DATE_OF_APP_EXE_AT_LOAD > 0 Then
                If F.DateLastModified > 0 Then
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
                            VBS_LAUNCHER_NAME = "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 40-RUN EXE 2 ARG_QUIET.VBS"
                            WSHShell.Run """" + VBS_LAUNCHER_NAME + """" + " " + APP_PATH_AND_EXE, ShowWindow_2, DontWaitUntilFinished
                            End
                        End If
                    End If
                End If
            End If
        End If
    End If

End Sub



Public Sub DATE_OF_APP_WITHER_VB_EXE_AT_LOAD_SUB()
    
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
                        VBS_LAUNCHER_NAME = "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 40-RUN EXE 2 ARG_QUIET.VBS"
                        WSHShell.Run """" + VBS_LAUNCHER_NAME + """" + " " + APP_PATH_AND_EXE, ShowWindow_2, DontWaitUntilFinished
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
    
If RZZD = 0 Then
    RZZD = 1
    Call VB_PROJECT_CHECKDATE
    Exit Sub
End If

If DISPLAY_RESULT_VAR_AFTER_FIRST_PASS = True Then
    RZZD = 2
    Call VB_PROJECT_CHECKDATE
    RZZD = 1
End If
RZZD = -10

TIMER_VB_PROJECT_CHECKDATE_FORM_LOAD.Enabled = False
Timer_VB_PROJECT_CHECKDATE.Enabled = True

End Sub


Public Sub VB_PROJECT_CHECKDATE(Optional FORM_LOAD_VAR)
                
    ' FIRST PASS AND 2ND
    ' RZZD = RZZD + 1
                
    If 1 = 2 Then
        Exit Sub
    End If
    
    If ONCE_STARTER_MODE = True Then
    Exit Sub
    End If
    ONCE_STARTER_MODE = True
    
    
    RETRY_CRC_ERROR = 4
    
    Dim VB_DATE
    Dim I_TEXT
    Dim VBS_LAUNCHER_NAME
    
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
    'IF A NEW PROJECT NOT BEEN SYNC YET TO VB_EXE FOLDER
    '---------------------------------------------------
    On Error Resume Next
    If Dir(PATH_FILE_NAME2) = "" Then
        If Dir(Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\")), vbDirectory) = "" Then
            CreateFolderTree Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\"))
        End If
    End If
    Err.Clear
    Set F = FSO.GetFile(PATH_FILE_NAME1)
    APP_EXENAME_DATE = F.DateLastModified
    Set F = FSO.GetFile(PATH_FILE_NAME2)
    VB_EXE_DATE = F.DateLastModified
    
    ' NOT ANY FURTHER CHECK AFTER ONE
    ' IF VB_EXE NOT MATCH VB THEN CONTUNE AHEAD
    ' UNABLE STORE VARIABLE FOR ONE PASS
    ' AS FORM UNLOADER
    ' UNTIL BETTER AR
    ' -----------------------------------------
    If APP_EXENAME_DATE = VB_EXE_DATE Then
        Exit Sub
    End If
    
    
    If APP_EXENAME_DATE > VB_EXE_DATE Then
        WxHex_FILE_2 = Hex(m_CRC.CalculateFile(PATH_FILE_NAME2))
        For RSB = 1 To RETRY_CRC_ERROR
            If Me.EXIT_TRUE = True Then
                Unload Me
                Exit Sub
            End If
            WxHex_FILE_1 = Hex(m_CRC.CalculateFile(PATH_FILE_NAME1))
            If WxHex_FILE_1 <> WxHex_FILE_2 Then
                i = App.EXEName + ".EXE" + vbCrLf + "CHECK PROJECT DATE __ 01 OF 03 ----------"
                i = i + vbCrLf + PATH_FILE_NAME1 + vbCrLf + PATH_FILE_NAME2 + " "
                Project_Check_Date.Label1 = i
                
                If RZZD = 1 Then
                    DISPLAY_RESULT_VAR_AFTER_FIRST_PASS = True
                    Exit Sub
                End If

                FSO.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
            End If
            WxHex_FILE_2 = Hex(m_CRC.CalculateFile(PATH_FILE_NAME2))
            If WxHex_FILE_1 = WxHex_FILE_2 Then
                Set F = FSO.GetFile(PATH_FILE_NAME1)
                SetFileDateTime PATH_FILE_NAME2, F.DateLastModified
                Exit For
            End If
        Next
    End If
    
    FILE_NAME = "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\Multiple_Thread Port Scanner 02 CON\NETWORK_COMPUTER_NAME.txt"
    'If Dir(FILE_NAME) = "" Then MsgBox "FILE NOT FOUND" + vbCrLf + vbCrLf + FILE_NAME
    If Dir(FILE_NAME) <> "" Then
        FR1 = FreeFile
        Open FILE_NAME For Input As #FR1
        Do
            Line Input #FR1, LINE_COMPUTER_NAME_1
            If Me.EXIT_TRUE = True Then
                Unload Me
                Exit Sub
            End If
            If LINE_COMPUTER_NAME_1 <> "" Then
                LINE_COMPUTER_NAME_2 = Replace(LINE_COMPUTER_NAME_1, "-", "_")
                VB_APP_PATH_NAME_1 = App.Path
                VB_APP_PATH_NAME_1 = Replace(VB_APP_PATH_NAME_1, "D:\VB6\", "D:\VB6-EXE\")
                VB_APP_PATH_NAME_1 = Mid(VB_APP_PATH_NAME_1, 4)
                LINE_EXE_PATH = "\\" + LINE_COMPUTER_NAME_1 + "\" + LINE_COMPUTER_NAME_2 + "_02_d_drive\" + VB_APP_PATH_NAME_1
                VB_APP_PATH_NAME_2 = LINE_EXE_PATH + "\" + App.EXEName + ".EXE"
                If Dir(Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\")), vbDirectory) = "" Then
                    CreateFolderTree Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\"))
                End If
                
                Set FSO = CreateObject("Scripting.FileSystemObject")
                Set F = FSO.GetFile(PATH_FILE_NAME1)
                APP_EXENAME_DATE = F.DateLastModified
                VB_EXE_DATE = 0
                Set F = FSO.GetFile(VB_APP_PATH_NAME_2)
                VB_EXE_DATE = F.DateLastModified
                
                If APP_EXENAME_DATE > VB_EXE_DATE Then

                    WxHex_FILE_1 = Hex(m_CRC.CalculateFile(PATH_FILE_NAME1))
                    WxHex_FILE_2 = Hex(m_CRC.CalculateFile(VB_APP_PATH_NAME_2))
                    If WxHex_FILE_1 <> WxHex_FILE_2 Then
                        For RSB = 1 To RETRY_CRC_ERROR
                            If Me.EXIT_TRUE = True Then
                                Unload Me
                                Exit Sub
                            End If
                            WxHex_FILE_1 = Hex(m_CRC.CalculateFile(PATH_FILE_NAME1))
                            If WxHex_FILE_1 <> WxHex_FILE_2 Then
                                i = App.EXEName + ".EXE" + vbCrLf + "CHECK PROJECT DATE __ 02 OF 03  ----------"
                                i = i + vbCrLf + PATH_FILE_NAME1 + vbCrLf + VB_APP_PATH_NAME_2 + " "
                                Project_Check_Date.Label1 = i
                                If RZZD = 1 Then
                                    DISPLAY_RESULT_VAR_AFTER_FIRST_PASS = True
                                    Exit Sub
                                End If

                                FSO.CopyFile PATH_FILE_NAME1, VB_APP_PATH_NAME_2
                                
                            End If
                            WxHex_FILE_2 = Hex(m_CRC.CalculateFile(VB_APP_PATH_NAME_2))
                            If WxHex_FILE_1 = WxHex_FILE_2 Then
                                Set F = FSO.GetFile(PATH_FILE_NAME1)
                                SetFileDateTime VB_APP_PATH_NAME_2, F.DateLastModified
                                Exit For
                            End If
                        Next
                    End If
                End If
                'MsgBox "COPY "
'                MsgBox "COPY " + vbCrLf + vbCrLf + PATH_FILE_NAME1 + vbCrLf + vbCrLf + VB_APP_PATH_NAME_2, vbMsgBoxSetForeground
                'Exit Sub
            End If
        Loop Until EOF(FR1)
        Close FR1
    End If
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set F = FSO.GetFile(PATH_FILE_NAME1)
    XVB_DATE = F.DateLastModified
    Set F = FSO.GetFile(PATH_FILE_NAME2)
    VB_DATE = F.DateLastModified
    
    VBS_LAUNCHER_NAME = ""
    
    If XVB_DATE > VB_DATE And XVB_DATE > 0 And VB_DATE > 0 Then
        Me.Refresh
        FSO.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
        
        WxHex_FILE_1 = Hex(m_CRC.CalculateFile(PATH_FILE_NAME1))
        WxHex_FILE_2 = Hex(m_CRC.CalculateFile(PATH_FILE_NAME2))
        If WxHex_FILE_1 <> WxHex_FILE_2 Then
            For RSB = 1 To RETRY_CRC_ERROR
                If Me.EXIT_TRUE = True Then
                    Unload Me
                    Exit Sub
                End If
                WxHex_FILE_1 = Hex(m_CRC.CalculateFile(PATH_FILE_NAME1))
                If WxHex_FILE_1 <> WxHex_FILE_2 Then
                    i = App.EXEName + ".EXE" + vbCrLf + "CHECK PROJECT DATE __ 03 OF 03 ----------"
                    i = i + vbCrLf + PATH_FILE_NAME1 + vbCrLf + PATH_FILE_NAME2 + " "
                    Project_Check_Date.Label1 = i
                    If RZZD = 1 Then
                        DISPLAY_RESULT_VAR_AFTER_FIRST_PASS = True
                        Exit Sub
                    End If
                    FSO.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
                End If
                WxHex_FILE_2 = Hex(m_CRC.CalculateFile(PATH_FILE_NAME2))
                If WxHex_FILE_1 = WxHex_FILE_2 Then
                    Set F = FSO.GetFile(PATH_FILE_NAME1)
                    SetFileDateTime PATH_FILE_NAME2, F.DateLastModified
                    Exit For
                End If
            Next
        End If
    End If
    
    VBS_LAUNCHER_NAME = App.Path + "\VBS - RELOAD AND COPY_" + GetComputerName + ".VBS"
    READY_TO_GO = False
    
    '----------------------------------
    'WRITE INFO ABOUT DATE CHANGE NEWER
    '----------------------------------
    If XVB_DATE < VB_DATE And XVB_DATE > 0 Then
    
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
        'Set objShell = WScript.CreateObject("WScript.Shell")
        '----------------------------------
    
        ' ----------------------------------------------------------
        ' MAJOR IMPROVEMENT AND GOT WORKING 1ST TIME RUNNING UPDATER
        ' THE UPDATED WHEN NEW COMPILED COMES ALONG EXIT AND RELAUNCH
        ' IN VBSCRIPT FROM A MIRROR FOLDER VB-EXE
        ' ----------------------------------------------------------
        ' Thu 03-May-2018 09:25:29
        ' Thu 03-May-2018 11:25:00 -- 3 HOUR
        ' ----------------------------------------------------------
        
        PATH_FILE_NAME1 = App.Path + "\" + App.EXEName + ".EXE"
        PATH_FILE_NAME2 = Replace(PATH_FILE_NAME1, "D:\VB6\", "D:\VB6-EXE\")
        
        I_TEXT = ""
        I_TEXT = I_TEXT + "Dim FSO" + vbCrLf
        I_TEXT = I_TEXT + "Set FSO = CreateObject(""SCRIPTING.FILESYSTEMOBJECT"")" + vbCrLf
        I_TEXT = I_TEXT + "On Error Resume Next" + vbCrLf
        I_TEXT = I_TEXT + "Do" + vbCrLf
        I_TEXT = I_TEXT + "    Err.Clear" + vbCrLf
        I_TEXT = I_TEXT + "    FC_2 = """ + PATH_FILE_NAME2 + """" + vbCrLf
        I_TEXT = I_TEXT + "    FC_1 = """ + PATH_FILE_NAME1 + """" + vbCrLf
        I_TEXT = I_TEXT + "    FN_1 = Mid(FC_2, InStrRev(FC_1, ""\"") + 1)" + vbCrLf
        I_TEXT = I_TEXT + "    FSO.CopyFile FC_2, FC_1" + vbCrLf
        I_TEXT = I_TEXT + "    Set F = FSO.GetFile(FC_1)" + vbCrLf
        I_TEXT = I_TEXT + "    XVB_DATE_1 = F.DateLastModified" + vbCrLf
        I_TEXT = I_TEXT + "    Set F = FSO.GetFile(FC_2)" + vbCrLf
        I_TEXT = I_TEXT + "    APP_EXENAME_DATE = F.DateLastModified" + vbCrLf
        I_TEXT = I_TEXT + "    IF XVB_DATE_1 <> APP_EXENAME_DATE THEN X_COUNT = X_COUNT + 1" + vbCrLf
        'I_TEXT = I_TEXT + "    MSGBOX X_COUNT" + vbCrLf
        I_TEXT = I_TEXT + "    WScript.Sleep 1000" + vbCrLf
        I_TEXT = I_TEXT + "    ' TEN MINUTES SLEEPER" + vbCrLf
        I_TEXT = I_TEXT + "Loop Until XVB_DATE_1=APP_EXENAME_DATE Or X_COUNT > 600" + vbCrLf
        I_TEXT = I_TEXT + "If X_COUNT > 600 Then" + vbCrLf
        I_TEXT = I_TEXT + "    MsgBox ""ERROR COPY FILE RETRY COUNT ""+X_COUNT+"" RETRY 10 MINUTE LEARN FOR VB UPDATE"" + vbCrLf + FC_1" + vbCrLf
        I_TEXT = I_TEXT + "End If" + vbCrLf
        I_TEXT = I_TEXT + "Err.Clear" + vbCrLf
        I_TEXT = I_TEXT + "Set UAC = CreateObject(""SHELL.APPLICATION"")" + vbCrLf
        '----------------------------------------------------
        'WINDOWS XP PROBLEM _ CHANGE THE SCRIPT HERE
        'NOT TO RUN AS ADMIN OR BRING THE RUNAS DIALOG BOX UP
        'WIN XP = 5.1 _ WINDOWS 10 = 6.2
        '----------------------------------------------------
        If GetWindowsVersion > 5.1 Then
            I_TEXT = I_TEXT + "UAC.ShellExecute FC_1, """", """", ""RUNAS"", 1" + vbCrLf
        Else
            I_TEXT = I_TEXT + "UAC.ShellExecute FC_1" + vbCrLf
        End If
        I_TEXT = I_TEXT + "If Err.Number > 0 Then" + vbCrLf
        I_TEXT = I_TEXT + "    MsgBox ""ERROR LAUNCH VB PROGRAM FROM UPDATE"" + vbCrLf + FC_1 + vbCrLf + Err.Description" + vbCrLf
        I_TEXT = I_TEXT + "End If" + vbCrLf
        I_TEXT = I_TEXT + "WScript.Quit 0" + vbCrLf
        
'        Syntax
'        .ShellExecute "application", "parameters", "dir", "verb", window
'        .ShellExecute 'some program.exe', '"some parameters with spaces"', , "runas", 1
'        Run a batch script with elevated permissions, flag=runas:
'        Set objShell = CreateObject("Shell.Application")
'        objShell.ShellExecute "E:\demo\batchScript.cmd", "", "", "runas", 1
'        Run a VBScript with elevated permissions, flag=runas:
'        Set objShell = CreateObject("Shell.Application")
'        objShell.ShellExecute "cscript", "E:\demo\vbscript.vbs", "", "runas", 1
                
        If Dir(VBS_LAUNCHER_NAME) <> "" Then Kill VBS_LAUNCHER_NAME
        FR1 = FreeFile
        Open VBS_LAUNCHER_NAME For Output As #FR1
            Print #FR1, I_TEXT
        Close #FR1
        
        READY_TO_GO = True
        
    End If
    
    If READY_TO_GO = True Then
        
        '-------------------------
        'GETTING 6.2 IN WINDOWS 10
        'LOOKING FOR XP = 5.1
        '-------------------------
        If GetWindowsVersion > 5.1 Or 1 = 1 Then
            'MsgBox "10"
            
            APP_PATH_AND_EXE = App.Path + "\" + App.EXEName + ".EXE"
            APP_PATH_AND_EXE = Replace(APP_PATH_AND_EXE, " ", "*")
            ' ADD STAR TO CONCOTION AND THEY REMOVE OTHER PARAMETER IN
            ' Mon 17-Feb-2020 02:20:37
            
            ' KILL THE WSCRIPT NAME
            ' FUNCTION & INCLUDE SUB
            ' FIND_SCRIPTNAME_VAR = "\VBS - RELOAD AND COPY_"
            ' -----------------------------------------------
            Call WSCRIPT_SCRIPTNAME_RELOAD_KILLER
            VBS_LAUNCHER_NAME = "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 40-RUN EXE 2 ARG_QUIET.VBS"
            WSHShell.Run """" + VBS_LAUNCHER_NAME + """" + " " + APP_PATH_AND_EXE, ShowWindow_2, DontWaitUntilFinished
            End
            
        End If
    End If
    
    If SET_FORM_PROJECT_CHECK_DATE_SOME_INFO_TRIGGER_SHOW = True Then
        SET_FORM_PROJECT_CHECK_DATE_SOME_INFO_TRIGGER_SHOW = False
        
        TIMER_RUN_AND_THEN_HIDE_PRE_EVENT.Interval = 1000
        TIMER_RUN_AND_THEN_HIDE_PRE_EVENT.Enabled = True
        Timer_RUN_AND_THEN_HIDE_DELAY.Interval = 10000
        Timer_RUN_AND_THEN_HIDE_DELAY.Enabled = True
    End If
End Sub

Private Sub TIMER_RUN_AND_THEN_HIDE_PRE_EVENT_Timer()
    TIMER_RUN_AND_THEN_HIDE_PRE_EVENT.Enabled = False
    RZZD = -10
    Project_Check_Date.Label1 = App.EXEName + ".EXE" + vbCrLf + "CHECK PROJECT DATE __ DONE "

End Sub

Private Sub Timer_RUN_AND_THEN_HIDE_DELAY_Timer()
    Timer_RUN_AND_THEN_HIDE_DELAY.Enabled = False
    
    NotAlwaysOnTop Me.hWnd
    Me.Visible = False
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

Private Function CreateFolderTree(ByVal sPath As String) As Boolean
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
  
  'TESTING
'  IsIDE = False
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************


Private Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Private Function NotAlwaysOnTop(ByVal hWnd As Long)
    Dim flags
    SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, flags
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




