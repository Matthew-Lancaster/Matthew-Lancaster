VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4608
   ClientLeft      =   132
   ClientTop       =   780
   ClientWidth     =   5976
   Icon            =   "VB_KEEP_RUNNER_OPERATION_KILL_EVENT.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4608
   ScaleWidth      =   5976
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer_EXIT 
      Interval        =   100
      Left            =   1968
      Top             =   3516
   End
   Begin VB.Timer TIMER_CLIPBOARD_RETRY 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1344
      Top             =   3432
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   2640
      Left            =   444
      TabIndex        =   0
      Top             =   348
      Width           =   4524
      WordWrap        =   -1  'True
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_ADMINISTRATOR 
      Caption         =   "ADMINISTRATOR"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --------------------------------------------
' MATT.LAN@BTINTERNET.COM
' --------------------------------------------
' --------------------------------------------
' WRITE CODE EVENT TIME
' --------------------------------------------
' SESSION 001
' Thu 23-Jan-2020 14:52:06
' Thu 23-Jan-2020 15:30:00 -- 1 HOUR 38 MINUTE
' --------------------------------------------
' SESSION 002 EXTEND TIME DIGEST
' Thu 23-Jan-2020 14:52:06
' Thu 23-Jan-2020 17:10:00 -- 2 HOUR 18 MINUTE
' --------------------------------------------
' SESSION 003 ADD ROUTINE TO TERMINATE SOFTLY BEFORE BEGIN
' Thu 23-Jan-2020 17:10:00
' Thu 23-Jan-2020 17:34:00 -- 24 MINUTE
' --------------------------- TOTAL 2 HOUR 42 MINUTE
' --------------------------------------------

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer

'#####################################################################################
'# Private sub/function/type/const/etc definitions required by the class
'#####################################################################################
'#########################################################
'# BEGIN IT ALL
'#########################################################
'#########################################################
'# General definitions
'#########################################################
Private Declare Function TerminateProcess _
        Lib "Kernel32" _
        (ByVal hProcess As Long, _
        ByVal uExitCode As Long) As Long


Private Declare Sub CloseHandle Lib "Kernel32" (ByVal hPass As Long)
Private Const MAX_PATH = 260

Private sEXEName As String
Private hProcess As Long
Private lProcessID As Long
Private bIsNT4 As Boolean
Private bIsOpen As Boolean


Private Declare Function FindWindowDLL _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long

'#########################################################
'# Functions/Consts/Types used for GetVersionEx()
'#########################################################
Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long           '#### NT: Build Number, 9x: High-Order has Major/Minor ver, Low-Order has build
    PlatformID As Long
    szCSDVersion As String * 128    '#### NT: ie- "Service Pack 3", 9x: 'arbitrary additional information'
End Type


'#########################################################
'# Win32 (non-NT4) specific definitions
'#########################################################
Private Declare Function CreateToolhelp32Snapshot Lib "Kernel32" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "Kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "Kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Const TH32CS_SNAPPROCESS As Long = 2&
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
    szExeFile As String * MAX_PATH
End Type

    '#### Required private data members
Private uProcess As PROCESSENTRY32
Private hSnapShot As Long


'#########################################################
'# Windows NT4 specific definitions
'#    NOTE: Remember to distribute the PSAPI.DLL file.
'#########################################################
Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
'STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

    '#### Required private data members
Private Modules(1 To 200) As Long
Private lProcessIDs() As Long
Private lcbNeeded As Long
Private lRetVal As Long
Private i As Long


'#########################################################
'# Convert() releated definitions
'#########################################################
    '#### Functions/Consts used for GethWnd() (part of Convert())
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const GW_hWndNEXT = 2
Private Const GW_CHILD = 5

    '#### eNum used with Convert()
Public Enum cnProcessConv
    cnFromClass = 1
    cnFromEXE = 2
    cnFromhWnd = 4
    cnFromProcessID = 8
    cnFromTitle = 16
    cnToClass = 32
    cnToEXE = 64
    cnTohWnd = 128
    cnToProcessID = 256
    cnToTitle = 512
End Enum

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Const GW_OWNER = 4
Private Const WS_EX_TOOLWINDOW = &H80
Private Const WS_EX_APPWINDOW = &H40000
Private Const GA_ROOT = 2

Private Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
' Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Const PROCESS_CREATE_PROCESS = &H80
Private Const PROCESS_CREATE_THREAD = &H2
Private Const PROCESS_DUP_HANDLE = &H40
Private Const PROCESS_SET_QUOTA = &H100
Private Const PROCESS_SET_INFORMATION = &H200
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_VM_OPERATION = &H8
Private Const PROCESS_VM_WRITE = &H20

Private Type PROCESS_INFORMATION
  hProcess    As Long
  hThread     As Long
  dwProcessId As Long
  dwThreadId  As Long
End Type


'#####################################################################################
'# Private sub/function/type/const/etc definitions required by the class
'#####################################################################################
'#####################################################################################
'# END IT ALL
'#####################################################################################




Private Sub FORM_Load()
    
    Dim i As String
    Dim EXE_N As String
    
    ' -------------------------------------------
    ' IF ISIDE TERMINATE PREVIOUS INSTANCE OF EXE
    ' NOT BLOCK THE COMPILER _.EXE
    ' -------------------------------------------
    EXE_N = App.Path + "\" + App.EXEName + ".exe"
    If IsIDE = True Then
        A_RESULT = GetEXEID_KILL_ALL_INSTANCE(EXE_N)
        If A_RESULT > -1 Then
            Beep
            End
        End If
    End If
    ' -------------------------------------------
    
    VAR_COMMAND = Command$
    
    If IsIDE = True Then VAR_COMMAND = "002"
    
    i = ""
    i = i + App.EXEName + vbCrLf
    i = i + vbCrLf
    i = i + "OPERATION KILL EVENT" + vbCrLf
    i = i + vbCrLf
    i = i + Format(Now, "DDD DD-MMM-YYYY HH:MM:SS AM/PM") + vbCrLf
    i = i + vbCrLf
    If VAR_COMMAND = "001" Then
        Label1.FontSize = 20
        i = i + "CURRENT TITLE NOT RESPOND FOR 120 SECOND" + vbCrLf
        i = i + vbCrLf
        i = i + "1.. IF 7-ASUS __ THEN __ CLOSE GOOSYNC" + vbCrLf
        i = i + vbCrLf
        i = i + "2.. KILL_WSCRIPT_GLOBAL" + vbCrLf
        i = i + vbCrLf
        i = i + "3.. KILL_CMD _ HARD PROCESS KILL __ NOT TASKILL" + vbCrLf
        i = i + vbCrLf
        i = i + ""
        i = i + vbCrLf
        i = i + ""
        i = i + vbCrLf
        i = TRIM_END_VBCRLF(i)
    End If
    If VAR_COMMAND = "002" Then
        Label1.FontSize = 12
        i = i + "TICKER MISSING PULSE STRESS SYTEM TRIGGER 4 SECOND MISS REPEAT 4 TIME" + vbCrLf
        i = i + vbCrLf
        i = i + "1.. CLOSE GOOSYNC AND PORTABLE IF C-DRIVE" + vbCrLf
        i = i + vbCrLf
        i = i + "2.. KILL_WSCRIPT_GLOBAL" + vbCrLf
        i = i + vbCrLf
        i = i + "3.. KILL_CMD _ HARD PROCESS KILL __ NOT TASKILL" + vbCrLf
        i = i + vbCrLf
        i = i + "4.. BAT 01-NOT RESPONDER KILLER NOT FORCE_WAIT.BAT" + vbCrLf
        i = i + "5.. BAT 01-NOT RESPONDER KILLER FORCE_WAIT.BAT" + vbCrLf
        i = i + vbCrLf
        i = i + "ARRAY_KILL(10) SET" + vbCrLf
        i = i + "D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe" + vbCrLf
        i = i + "D:\VB6\VB-NT\00_Best_VB_01\CLIPBOARD_VIEWER\ClipBoard Viewer.exe" + vbCrLf
        i = i + "D:\VB6\VB-NT\00_Best_VB_01\CPU % OF A PROGRAM\CPU % INDIVIDUAL PROCESS.exe" + vbCrLf
        i = i + "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\Tidal.exe" + vbCrLf
        i = i + "C:\PStart\# NOT INSTALL REQUIRED\Tail\Tail.exe" + vbCrLf
        i = i + "C:\Program Files\Mozilla Firefox\firefox.exe" + vbCrLf
        i = i + "D:\VB6\VB-NT\00_Best_VB_01\URL Logger\URL Logger.exe" + vbCrLf
        i = i + vbCrLf
        i = i + "" + vbCrLf
        i = i + "" + vbCrLf
        i = i + "" + vbCrLf
        i = TRIM_END_VBCRLF(i)
    End If
    
    If VAR_COMMAND = "" Then
        i = i + "COMMAND$ __ NOT GIVEN" + vbCrLf
        i = i + vbCrLf
        i = i + ""
    End If
    
    Label1.Caption = i
    Label1.WordWrap = False
    Label1.FontBold = True
    Label1.FontName = "ARIAL"
    Label1.Top = 50
    Label1.Left = 80
    Me.Width = Label1.Width + 280
    Me.Height = Label1.Height + 900
    
    Call MNU_ADMINISTRATOR_Click
    
End Sub

Function TRIM_END_VBCRLF(i)
    ' -------------------------------------
    ' FUNCTION STATION EXTENSION
    ' SESSION MOTION SECTOR SEGMENT SECTION
    ' NUMBER RULE PASSAGE PATH
    ' UP THE GARDEN
    ' THIRD EYE
    ' -------------------------------------
    i = Trim(i)
    RI = 0
    For R = Len(i) To 1 Step -1
        O_RI = RI
        If Mid(i, R, 1) = vbCr Then RI = R
        If Mid(i, R, 1) = vbLf Then RI = R
        If O_RI = RI Then Exit For
    Next
    If RI > 0 Then i = Mid(i, 1, RI - 1)
    TRIM_END_VBCRLF = i                 ' --- I HAVE GO BACK THE FUNCTION
End Function

Private Sub MNU_EXIT_Click()
    End
End Sub

Private Sub Label1_Click()
    On Error Resume Next
    Me.WindowState = vbMinimized
    Err.Clear
    Clipboard.Clear
    Err.Clear
    Clipboard.SetText VARTIME_DIFFERENCE
    If Err.NUMEBR > 0 Then
        VAR_TIMER_CLIPBOARD_RETRY = "Label1_Click"
        TIMER_CLIPBOARD_RETRY.Enabled = True
    End If
    If Err.NUMEBR = 0 Then
        End
    End If
End Sub

Private Sub Timer_EXIT_Timer()
    If Me.hWnd = GetForegroundWindow Then
        
        If GetAsyncKeyState(27) Then     ' --------- ESCAPE KEY
            Call Label1_Click            ' --------- GET CLIPBOARD ITEM
            End
        End If
    End If
End Sub

Private Sub TIMER_CLIPBOARD_RETRY_Timer()

If VAR_TIMER_CLIPBOARD_RETRY = "MNU_GIVE_TIME_DIFF_Click" Then
    Call Label1_Click
End If

TIMER_CLIPBOARD_RETRY.Enabled = False

End Sub

Private Sub Mnu_VB_ME_Click()
    
    Dim CODER_VBP_FILE_NAME_2
    Dim VB_1, VB_2, VB_3
    VB_1 = "C:\PROGRAM FILES\Microsoft Visual Studio\VB98\VB6.EXE"
    VB_2 = "C:\PROGRAM FILES (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
    If Dir(VB_1) <> "" Then VB_3 = VB_1
    If Dir(VB_2) <> "" Then VB_3 = VB_2
    '------------------------------------------------------
    ' ADD
    ' MICROSOFT SCRIPTING RUNTIME
    ' FROM REFERENCES
    '------------------------------------------------------
    Dim objShell
    Set objShell = CreateObject("Wscript.Shell")
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
    objShell.Run """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
    Set objShell = Nothing
    
    EXIT_TRUE = True
    Unload Me
End Sub

Private Sub MNU_VB_FOLDER_Click()
    Beep
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
End Sub

Private Sub MNU_ADMINISTRATOR_Click()
    MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR MODE ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "MODE ON", "MODE OFF")
    End If
End Sub

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

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
  
  'TESTING
  'IsIDE = False
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************



' --------------------------------------------------------------------------------------------
' CODE SET SOURCE LOCATION REFERANCE
' --------------------------------------------------------------------------------------------
' D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\Process Info, Traversal & Conversion\clsCnProc.cls
' --------------------------------------------------------------------------------------------

'#########################################################
'# Opens the processes and moves to the first ProcessID
'#########################################################
Public Function OpenProcesses() As Boolean
        '#### If we're running under WinNT4
    If (bIsNT4) Then
        Dim lcb As Long

            '#### Init the API vars
        lcbNeeded = 96
        lcb = 8

            '#### While lcbNeeded is bigger then lcb, loop to find the correct size of lProcessIDs()
        Do While (lcb <= lcbNeeded)
                '#### Increase the size of lcb by 2, dividing the result by 4 to convert from bytes to processes
            lcb = lcb * 2
            ReDim lProcessIDs(lcb / 4)

                '#### If the return value is 0, error out
            If (EnumProcesses(lProcessIDs(1), lcb, lcbNeeded) = 0) Then
                GoTo OpenProcesses_Error
            End If
        Loop

            '#### Init i, set bIsOpen and move to the first process, returning the result of the NextProcess() call to the caller
        i = 1
        bIsOpen = True
        OpenProcesses = NextProcess()

        '#### Else we're running under a system that supports CreateToolhelp32Snapshot()
    Else
            '#### Setup hSnapShot, begin to setup the uProcess UDT
        hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
        uProcess.dwSize = Len(uProcess)

            '#### If hSnapShot was successfully setup
        If (hSnapShot <> 0) Then
                '#### Find the first hProcess and set the return value
            hProcess = ProcessFirst(hSnapShot, uProcess)
            OpenProcesses = (hProcess <> 0)

                '#### If a valid hProcess was found, set sEXEName and lProcessID
            If (OpenProcesses) Then
                sEXEName = TrimNull(uProcess.szExeFile)
                lProcessID = uProcess.th32ProcessID
            End If

            '#### Else hSnapShot was not successfully setup, so error out
        Else
            GoTo OpenProcesses_Error
        End If
    End If

        '#### Set bIsOpen
    bIsOpen = OpenProcesses
    Exit Function

OpenProcesses_Error:
    OpenProcesses = False
    bIsOpen = False
End Function

'#########################################################
Public Function GetEXEID_KILL_ALL_INSTANCE(ByRef sRunningEXE As String) As Long
        '#### Default the return value
        '#### If we're able to successfully open the processes
    Dim VAR
    
    ' MUST GIVE THE VARIABLE A VALUE OR ERROR EVENT AS EXIT
    ' -----------------------------------------------------
    GetEXEID_KILL_ALL_INSTANCE = -1
    
    If (OpenProcesses()) Then
            '#### Get the name of the EXE
        sRunningEXE = UCase(GetFileName(sRunningEXE))
            '#### Do while we still have processes
        Do
            If (InStr(1, UCase(sEXEName), UCase(sRunningEXE), vbBinaryCompare) > 0) Then
                VAR = Process_Kill(lProcessID)
                If GetEXEID_KILL_ALL_INSTANCE = -1 Then GetEXEID_KILL_ALL_INSTANCE = 0
                GetEXEID_KILL_ALL_INSTANCE = GetEXEID_KILL_ALL_INSTANCE + 1
            End If
        Loop While (NextProcess())
    End If

End Function

'#########################################################
'# Peals off the last filename element from the passed sPath and returns it to the caller
'#########################################################
Private Function GetFileName(ByVal sPath As String) As String
    Dim iLastSpace As Integer
    Dim iLastSlash As Integer

        '#### If sPath has a value, process it
    If (Len(sPath) > 0) Then
            '#### Deteremine the index of the last "\" and the following " " (if there is one)
        iLastSlash = InStrRev(sPath, "\", -1)
        iLastSpace = InStr(iLastSlash + 1, sPath, " ")

            '#### If a space was found in sPath, make sure we only remove the EXEs name and not any arguments
        If (iLastSpace > 0) Then
            'GetFileName = Mid(sPath, iLastSlash + 1, iLastSpace - iLastSlash - 1)

            '#### Else there were no arguments, so peal off the name
            
            'ABORT THIS CAUSING ERROR TRIM AWAY PART FILE EXE NAME
            GetFileName = Mid(sPath, iLastSlash + 1)
        Else
            GetFileName = Mid(sPath, iLastSlash + 1)
        End If
    End If
End Function

'#########################################################
'# Moves to the next lProcessID, setting sEXEName and lProcessID
'#########################################################
Public Function NextProcess() As Boolean
        '#### If there is currently process info open
    If (bIsOpen) Then
            '#### If we're running under WinNT4
        If (bIsNT4) Then
                '#### If we are still within the bounds of lProcessIDs()
            If (i <= (lcbNeeded / 4)) Then
                    '#### Setup hProcess and set the return value
                hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcessIDs(i))
                NextProcess = True

                    '#### If hProcess returned a valid value
                If (hProcess <> 0) Then
                        '#### Set lProcessID
                    lProcessID = lProcessIDs(i)

                        '#### If we are able to retrieve the module handels for the found hProcess
                    If (EnumProcessModules(hProcess, Modules(1), 200, 0) <> 0) Then
                            '#### Init the sEXEName buffer, retrieve the module name then trim it based on lRetVal
                        sEXEName = Space(MAX_PATH)
                        lRetVal = GetModuleFileNameExA(hProcess, Modules(1), sEXEName, 500)
                        sEXEName = Left(sEXEName, lRetVal)
                    End If

                    '#### Else hProcess did not return a valid value, so reset lProcessID and sEXEName
                Else
                    lProcessID = 0
                    sEXEName = ""
                End If

                    '#### Close hProcess and inc i for the next call
                Call CloseHandle(hProcess)
                i = i + 1

                '#### Else we're outside the bounds of lProcessIDs(), so return false
            Else
                NextProcess = False
            End If

            '#### Else we're running under a system that supports CreateToolhelp32Snapshot()
        Else
                '#### If the current hProcess is valid
            If (hProcess <> 0) Then
                    '#### Move hProcess to the next hProcess and set the return value
                hProcess = ProcessNext(hSnapShot, uProcess)
                NextProcess = (hProcess <> 0)

                    '#### If a valid hProcess was found, set sEXEName and lProcessID
                If (NextProcess) Then
                    sEXEName = uProcess.szExeFile
                    lProcessID = uProcess.th32ProcessID
                End If

                '#### Else the current hProcess is invalid, so return false
            Else
                NextProcess = False
            End If
        End If

        '#### Else we're not currently open, so return false
    Else
        NextProcess = False
    End If
End Function

'#########################################################
'# Trims the passed sString up to vbNull
'#########################################################
Private Function TrimNull(ByVal sString As String) As String
    Dim lIndex As Long

        '#### Default the return value and determine if there is a vbNull in sString
    TrimNull = sString
    lIndex = InStr(1, TrimNull, Chr(0), vbBinaryCompare)

        '#### If a vbNull was present, trim up to it and return
    If (lIndex > 0) Then TrimNull = Left(TrimNull, lIndex - 1)
End Function

Public Function Process_Kill(P_ID As Long) As Long
    '// Kill the wanted process
    
    ' ---------------------------------------------------------
    ' EXAMPLE ---- VAR = cProcesses.Process_Kill(PID)
    ' ---------------------------------------------------------
    ' USED TO BE THAT WHEN CLASS CODE SET ORIGINAL SOURCE FOUND
    ' ---------------------------------------------------------
    ' CLASS WAY NOT ANYMORE SO REMOVE
    ' cProcesses.
    ' AND NOT REQUIRE INITIALISE CLASS
    ' ---------------------------------------------------------
        
    
    Dim hProcess As Long
    Dim lExitCode As Long
    Dim XI As Long
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, P_ID)
    
    XI = TerminateProcess(hProcess, lExitCode) = False
    
    Call CloseHandle(hProcess)
    
End Function

