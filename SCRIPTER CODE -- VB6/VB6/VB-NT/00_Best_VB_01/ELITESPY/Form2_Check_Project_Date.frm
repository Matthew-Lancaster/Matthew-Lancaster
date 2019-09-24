VERSION 5.00
Begin VB.Form Project_Check_Date 
   BackColor       =   &H00808080&
   Caption         =   "Form2"
   ClientHeight    =   1476
   ClientLeft      =   192
   ClientTop       =   540
   ClientWidth     =   9084
   LinkTopic       =   "Form2"
   ScaleHeight     =   1476
   ScaleWidth      =   9084
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer_HIDE 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3360
      Top             =   1044
   End
   Begin VB.Timer Timer_VB_PROJECT_CHECKDATE 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2496
      Top             =   1128
   End
   Begin VB.Label Label1 
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
      Height          =   1176
      Left            =   168
      TabIndex        =   0
      Top             =   132
      Width           =   8736
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

Dim FORM_LOAD_CHECK_PROJECT_DATE_DO_AH_ONCE


Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long


Dim XVB_DATE_SYNC_VB_PROJECT
Dim XVB_DATE
Dim XVB_DATE_2
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

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1


'-----------------------------------------------------------------
Private Declare Function GetVersionExA Lib "kernel32" _
(lpVersionInformation As OSVERSIONINFO) As Integer

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type

Function GetWindowsVersion()
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
    sngWindowsVersion = CSng((CStr(osinfo.dwMajorVersion) & "." & CStr(osinfo.dwMinorVersion)))
    strWindowsVersion = CStr(osinfo.dwMajorVersion) & "." & CStr(osinfo.dwMinorVersion) & "." & CStr(osinfo.dwBuildNumber)
    GetWindowsVersion = osinfo.dwMajorVersion
    GetWindowsVersion = CSng((CStr(osinfo.dwMajorVersion) & "." & CStr(osinfo.dwMinorVersion)))
  
    '----------------------------------------------------
    'WINDOWS XP PROBLEM _ CHANGE THE SCRIPT HERE
    'NOT TO RUN AS ADMIN OR BRING THE RUNAS DIALOG BOX UP
    'WIN XP = 5.1 _ WINDOWS 10 = 6.2
    '----------------------------------------------------

End Function
'-----------------------------------------------------------------

Private Sub Form_Load()
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Timer_VB_PROJECT_CHECKDATE.Enabled = True
    
    Label1.FontSize = 30
    'List1.FontSize = 14
    
    FORM_LOAD_CHECK_PROJECT_DATE_DO_AH_ONCE = True
    
    Exit Sub

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
    ' ALSO AS NOTER ELSEWHERE THIS MODUAL
    ' -------------------------------------------------
    ' TOP DELCLARE ------------------------------------
    ' -------------------------------------------------
    ' Dim XVB_DATE
    ' Dim XVB_DATE_2
    ' Dim FSO
    ' Public EXIT_TRUE ' THIS IS USUALAR SET IN EACH FORM
    ' Public CHECK_PROJECT_DATE_IN_PROCESS ' SET EACH FORM USER AND COMMON BAS
    ' COMMON BAS NOT REALLY REQUIRE ONLY BETWEEN FORM
    ' BUT DOESN'T MOAN IF THERE AND EXPANSION EXIST
    
    ' REQUIRE HERE _TIMER IN MAIN FORM AND LABEL HERE
    ' THEM WILL BE FOUND AND USE IF THERE
    ' -------------------------------------------------
    ' Public Sub Timer_CHECK_PROJECT_DATE_IN_PROCESS_Timer()
    '     If Me.CHECK_PROJECT_DATE_IN_PROCESS = True Then
    '         Label_CHECK_PROJECT_DATE_IN_PROCESS.BackColor = &HFF00&
    '         ' Hex (Label1(4).BackColor) ' &H0000FF00& -- BRIGHT GREEN
    '         ' Hex (Label1(5).BackColor) ' &HC0FFFF&   -- PALE YELLOW
    '     Else
    '         Label_CHECK_PROJECT_DATE_IN_PROCESS.BackColor = &HC0FFFF
    '     End If
    '     Label_CHECK_PROJECT_DATE_IN_PROCESS.Refresh
    '     DoEvents
    ' End Sub
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
    ' i = i + 1: FrmxName(i) = "Form2_Check_Project_Date"  ' PREVIOUS  NAME USER
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
    '    'If InStr(Form.Name, "Form2_Check_Project_Date") Then
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
    ' formName = "Form2_Check_Project_Date"
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

Private Sub Timer_HIDE_Timer()
Me.Hide
Timer_HIDE.Enabled = False
End Sub

Private Sub Timer_VB_PROJECT_CHECKDATE_Timer()

    If Timer_VB_PROJECT_CHECKDATE.Interval <> 1000 Then
        Timer_VB_PROJECT_CHECKDATE.Interval = 1000
    End If
    
    Call VB_PROJECT_CHECKDATE("")
    
    If FORM_LOAD_CHECK_PROJECT_DATE_DO_AH_ONCE = True Then
        FORM_LOAD_CHECK_PROJECT_DATE_DO_AH_ONCE = False
        Project_Check_Date.Visible = False
    End If

End Sub


Public Sub VB_PROJECT_CHECKDATE(FORM_LOAD_VAR)
            
    Dim CONTROL_NAME As String
    Me.CHECK_PROJECT_DATE_IN_PROCESS = True
    CONTROL_NAME = "Timer_CHECK_PROJECT_DATE_IN_PROCESS"
    Call ControlCall_Find(CONTROL_NAME)
    DoEvents
    

    Timer_HIDE.Enabled = True

    ' -------------------------------------------------
    ' TOP DELCLARE ------------------------------------
    ' -------------------------------------------------
    ' Dim XVB_DATE
    ' Dim XVB_DATE_2
    ' Dim FSO
    ' Public EXIT_TRUE ' THIS IS USUALAR SET IN EACH FORM
    ' Public CHECK_PROJECT_DATE_IN_PROCESS ' SET EACH FORM USER AND COMMON BAS
    ' -------------------------------------------------
    ' USE THIS LINE IN FORM_LOAD OF MAIN START PROJECT
    ' AT THE BEGINNING
    ' Call Form2_Check_Project_Date.VB_PROJECT_CHECKDATE("FORM LOAD")
    ' -------------------------------------------------
    
    Dim PATH_FILE_NAME1
    Dim PATH_FILE_NAME2
    Dim VB_DATE
    Dim I_TEXT
    Dim VBS_LAUNCHER_NAME
    
    If Mid(App.Path, 1, 1) <> "D" Then
        Me.CHECK_PROJECT_DATE_IN_PROCESS = False
        Call ControlCall_Find(CONTROL_NAME)
        Exit Sub
    End If
    
    If XVB_DATE_2 = 0 Then
        Timer_VB_PROJECT_CHECKDATE.Interval = 1000
        PATH_FILE_NAME1 = App.Path + "\" + App.EXEName + ".EXE"
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set F = FSO.GetFile(PATH_FILE_NAME1)
        XVB_DATE_2 = F.DateLastModified
    End If
    
    
    If XVB_DATE_SYNC_VB_PROJECT = 0 Then
        FILE_NAME = App.Path + "\# DATA\VB Project EXE Date__" + GetComputerName + "_" + GetUserName + ".txt"
        FILE_NAME_PATH = App.Path + "\# DATA"
        If Dir(FILE_NAME_PATH, vbDirectory) = "" Then
            On Error Resume Next
            MkDir FILE_NAME_PATH
            On Error GoTo 0
        End If
        If Dir(FILE_NAME) <> "" Then
            On Error Resume Next
            FR1 = FreeFile
            Open FILE_NAME For Input As #FR1
                Line Input #FR1, XVB_DATE_SYNC_VB_PROJECT
            Close FR1
            On Error GoTo 0
        End If
        If XVB_DATE_SYNC_VB_PROJECT <> "" Then
            On Error Resume Next
            XVB_DATE_24 = DateValue(XVB_DATE_SYNC_VB_PROJECT) + TimeValue(XVB_DATE_SYNC_VB_PROJECT)
            On Error GoTo 0
        End If
        On Error Resume Next
        FR1 = FreeFile
        Open FILE_NAME For Output As #FR1
            Print #FR1, XVB_DATE_2
        Close FR1
        On Error GoTo 0
        If XVB_DATE_24 < XVB_DATE_2 Then
            COPY_OVER_NEW_EXE = True
        End If
    End If
    
    
    PATH_FILE_NAME1 = App.Path + "\" + App.EXEName + ".EXE"
    PATH_FILE_NAME2 = Replace(PATH_FILE_NAME1, "D:\VB6\", "D:\VB6-EXE\")
    '---------------------------------------------------
    'IF A NEW PROJECT NOT BEEN SYNC YET TO MIRROR FOLDER
    '---------------------------------------------------
    On Error Resume Next
    If Dir(PATH_FILE_NAME2) = "" Or COPY_OVER_NEW_EXE = True Then
        If Err.Number > 0 Then
            Me.CHECK_PROJECT_DATE_IN_PROCESS = False
            Call ControlCall_Find(CONTROL_NAME)
            Exit Sub
        End If
        
        If Dir(Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\")), vbDirectory) = "" Then
            If Err.Number > 0 Then
                Me.CHECK_PROJECT_DATE_IN_PROCESS = False
                Call ControlCall_Find(CONTROL_NAME)
                Exit Sub
            End If
            
            CreateFolderTree Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\"))
        End If
        Err.Clear
    
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set F = FSO.GetFile(PATH_FILE_NAME1)
        XVB_DATE_2 = F.DateLastModified
        Set F = FSO.GetFile(PATH_FILE_NAME2)
        XVB_DATE_3 = F.DateLastModified
        
        If XVB_DATE_2 > XVB_DATE_3 Then
            List1.AddItem "COPY TO >> " + PATH_FILE_NAME2
            List1.Refresh
            Me.Refresh
            FSO.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
        End If
        If Err.Number > 0 Then
            Me.CHECK_PROJECT_DATE_IN_PROCESS = False
            Call ControlCall_Find(CONTROL_NAME)
            Exit Sub
        End If
    
        
        FILE_NAME = "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\Multiple_Thread Port Scanner 02 CON\NETWORK_COMPUTER_NAME.txt"
        'If Dir(FILE_NAME) = "" Then MsgBox "FILE NOT FOUND" + vbCrLf + vbCrLf + FILE_NAME
        If Dir(FILE_NAME) <> "" Then
            FR1 = FreeFile
            Open FILE_NAME For Input As #FR1
            Do
                Line Input #FR1, LINE_COMPUTER_NAME_1
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
                    XVB_DATE_2 = F.DateLastModified
                    XVB_DATE_3 = 0
                    Set F = FSO.GetFile(VB_APP_PATH_NAME_2)
                    XVB_DATE_3 = F.DateLastModified
                    
                    If XVB_DATE_2 > XVB_DATE_3 Then
                    
                                            ' ----------------------------------------------------
                        ' IF START ANY NETWORK PROJECT
                        ' DATE CHECK BACK UP GO AND THEN SHOW HERE
                        ' THE FORM SHOW
                        ' HOLD UP COMPUTER FEW SECOND SMOEBODY MIGHT WONDER WHAT GOING ON
                        ' BESIDE COLOUR LABLE BOX CHANGE THAT NOT INDICATE ENOUGH
                        ' ----------------------------------------------------
                        If FORM_LOAD_CHECK_PROJECT_DATE_DO_AH_ONCE = True Then
                            Project_Check_Date.Show
                            Project_Check_Date.SetFocus
                            Dim Form
                            For Each Form In Forms
                                If Form.Caption <> "" Then
                                    Project_Check_Date.Caption = Form.Caption
                                End If
                                Exit For
                            Next
                            Project_Check_Date.Refresh
                            RESULT_API = AlwaysOnTop(Project_Check_Date.hwnd)
                            Project_Check_Date.Refresh
                            ' DoEvents
                        End If

                        List1.AddItem "COPY TO >> " + VB_APP_PATH_NAME_2
                        List1.Refresh
                        Me.Refresh
                        DoEvents
                        FSO.CopyFile PATH_FILE_NAME1, VB_APP_PATH_NAME_2
                                                
                    End If
                    'MsgBox "COPY "
    '                MsgBox "COPY " + vbCrLf + vbCrLf + PATH_FILE_NAME1 + vbCrLf + vbCrLf + VB_APP_PATH_NAME_2, vbMsgBoxSetForeground
                    'Exit Sub
                    
                End If
            Loop Until EOF(FR1)
            Close FR1
        End If
    End If
    
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set F = FSO.GetFile(PATH_FILE_NAME1)
    XVB_DATE = F.DateLastModified
    Set F = FSO.GetFile(PATH_FILE_NAME2)
    VB_DATE = F.DateLastModified
    
    VBS_LAUNCHER_NAME = ""
    
    If XVB_DATE > VB_DATE And XVB_DATE > 0 And VB_DATE > 0 Then
        List1.AddItem "COPY TO >> " + PATH_FILE_NAME2
        List1.Refresh
        Me.Refresh
        FSO.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
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
        I_TEXT = I_TEXT + "    XVB_DATE_2 = F.DateLastModified" + vbCrLf
        I_TEXT = I_TEXT + "    IF XVB_DATE_1 <> XVB_DATE_2 THEN X_COUNT = X_COUNT + 1" + vbCrLf
        'I_TEXT = I_TEXT + "    MSGBOX X_COUNT" + vbCrLf
        I_TEXT = I_TEXT + "    WScript.Sleep 1000" + vbCrLf
        I_TEXT = I_TEXT + "    ' TEN MINUTES SLEEPER" + vbCrLf
        I_TEXT = I_TEXT + "Loop Until XVB_DATE_1=XVB_DATE_2 Or X_COUNT > 600" + vbCrLf
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
            
            Dim WSHShell
            Set WSHShell = CreateObject("WScript.Shell")
                WSHShell.Run """" + VBS_LAUNCHER_NAME + """", 1, False
            Set WSHShell = Nothing
        End If
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
        
        
        If IsIDE = True Then
            Me.CHECK_PROJECT_DATE_IN_PROCESS = False
            Call ControlCall_Find(CONTROL_NAME)
            Exit Sub
        End If
        
        If FORM_LOAD_VAR = "FORM LOAD" Then End
        
        EXIT_TRUE = True
    
        For Each Form In Forms
            Unload Form
            Set Form = Nothing
        Next Form
        Unload Me
        
    End If
    DoEvents
    Me.CHECK_PROJECT_DATE_IN_PROCESS = False
    Call ControlCall_Find(CONTROL_NAME)
            
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


Private Function AlwaysOnTop(ByVal hwnd As Long)  'Makes a form always on top
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Private Function NotAlwaysOnTop(ByVal hwnd As Long)
    Dim flags
    SetWindowPos hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, flags
End Function


