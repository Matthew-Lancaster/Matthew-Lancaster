VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_1 
   Caption         =   "Form1"
   ClientHeight    =   6216
   ClientLeft      =   192
   ClientTop       =   540
   ClientWidth     =   11688
   Icon            =   "CPU % INDIVIDUAL PROCESS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6216
   ScaleWidth      =   11688
   Begin VB.Timer TIMER_MSGBOX_FROM_PROCESS_LASSO_OVER_MEMORY_SIZE 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3768
      Top             =   4092
   End
   Begin VB.Timer Timer_VB_PROJECT_CHECKDATE 
      Interval        =   1000
      Left            =   1875
      Top             =   4815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2592
      Left            =   9084
      TabIndex        =   6
      Top             =   1620
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   4572
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Timer TIMER_EXE_PICKUP 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   4815
   End
   Begin VB.Timer SOUND_PLAY_END_TIMER 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1020
      Top             =   4815
   End
   Begin VB.Timer SOUND_PLAY_TIMER 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   630
      Top             =   4800
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   2670
      TabIndex        =   5
      Top             =   5055
      Visible         =   0   'False
      Width           =   2835
      _ExtentX        =   6244
      _ExtentY        =   572
      _Version        =   393216
      PrevEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1308
      Left            =   540
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1284
   End
   Begin VB.Timer Timer_INIT 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   630
      Top             =   4455
   End
   Begin VB.ListBox List3 
      Height          =   240
      Left            =   1305
      TabIndex        =   0
      Top             =   4215
      Width           =   852
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1056
      Left            =   0
      TabIndex        =   3
      Top             =   852
      Visible         =   0   'False
      Width           =   408
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   3240
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   5790
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   630
      Top             =   4095
   End
   Begin MSComctlLib.ListView ListView_INDEX 
      Height          =   2790
      Left            =   1920
      TabIndex        =   7
      Top             =   825
      Width           =   1275
      _ExtentX        =   2244
      _ExtentY        =   4911
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   756
      Left            =   9096
      TabIndex        =   9
      Top             =   840
      Width           =   912
      _ExtentX        =   1609
      _ExtentY        =   1334
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E6FEFF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   492
      Width           =   7788
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EAFFEE&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   12
      TabIndex        =   2
      Top             =   36
      Width           =   7788
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_PROCESS_OBSERVE_MENU 
      Caption         =   "PROCESS_OBSERVE_MENU"
      Begin VB.Menu MNU_FLICKR 
         Caption         =   "FLICKR"
      End
      Begin VB.Menu MNU_DRIVERBOOSTER 
         Caption         =   "DRIVER BOOSTER __ IOBIT"
      End
      Begin VB.Menu MNU_GOOGLEDRIVESYNC 
         Caption         =   "GOOGLEDRIVESYNC.EXE"
      End
      Begin VB.Menu MNU_GOODSYNC 
         Caption         =   "Goodsync.exe"
      End
      Begin VB.Menu MNU_GOODSYNCV10 
         Caption         =   "Goodsync-v10.exe"
      End
      Begin VB.Menu MNU_GoodSync2Go 
         Caption         =   "GOODSYNC2Go"
      End
      Begin VB.Menu MNU_CCLEANER 
         Caption         =   "CCLEANER64 -- NOT WORK 64 BIT "
      End
      Begin VB.Menu POWER_DATA_RECOVERY_EXE 
         Caption         =   "POWER DATA RECOVERY.EXE"
      End
      Begin VB.Menu MNU_WSCRIPT 
         Caption         =   "WSCRIPT"
      End
      Begin VB.Menu MNU_DuplicateCleaner 
         Caption         =   "DuplicateCleaner"
      End
      Begin VB.Menu MNU_DuplicateCleanerV5 
         Caption         =   "Duplicate Cleaner v5"
      End
      Begin VB.Menu MNU_NeroVision 
         Caption         =   "NEROVISION"
      End
      Begin VB.Menu Mnu_Freemaker 
         Caption         =   "FREEMAKER"
      End
      Begin VB.Menu MNU_DLLHOST_EXE_WHEN_SET_SECRUITY_ON_FOLDER 
         Caption         =   "DLLHOST.EXE"
      End
      Begin VB.Menu MNU_EXPLORER_EXE 
         Caption         =   "EXPLORER.EXE"
      End
      Begin VB.Menu MNU_UNSTOPPABLE_COPIER 
         Caption         =   "RoadKil's Unstoppable Copier"
      End
      Begin VB.Menu MNU_FILE_LOCATOR 
         Caption         =   "File Locator Pro"
      End
      Begin VB.Menu MNU_AGENT_RANSACK 
         Caption         =   "Agent Ransack"
      End
      Begin VB.Menu MNU_FFMPEG 
         Caption         =   "FFMPEG && MOBILE MEDIA CONVERTOR"
      End
      Begin VB.Menu MNU_CMD 
         Caption         =   "CMD"
      End
      Begin VB.Menu MNU_CHKDSK 
         Caption         =   "CHKDSK.EXE"
      End
      Begin VB.Menu MNU_PARTITION_WIZARD 
         Caption         =   "Partition Wizard.exe"
      End
      Begin VB.Menu MNU_OUTLOOK_OFFICE 
         Caption         =   "OUTLOOK OFFICE 2007"
      End
      Begin VB.Menu MNU_HASH_MY_FILE 
         Caption         =   "HASH MY FILE"
      End
      Begin VB.Menu MNU_INSYNC 
         Caption         =   "INSYNC"
      End
      Begin VB.Menu MNU_HANDBRAKE 
         Caption         =   "HANDBRAKE"
      End
      Begin VB.Menu MNU_FormatConverter_EXE 
         Caption         =   "FormatConverter.exe HIKVISION"
      End
   End
   Begin VB.Menu MNU_KILL_WSCRIPT 
      Caption         =   "KILL WSCRIPT"
   End
   Begin VB.Menu MNU_MINI_DISPLAY 
      Caption         =   "SHOW MINI_DISPLAY"
   End
   Begin VB.Menu MNU_SOUND_TEST 
      Caption         =   "SOUND_TEST"
   End
   Begin VB.Menu MNU_AUDIO_ON 
      Caption         =   "AUDIO OFF"
   End
   Begin VB.Menu MNU_MENU_ITEM_COUNT 
      Caption         =   "MENU ITEM COUNT"
   End
End
Attribute VB_Name = "Form_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LAB_TIMEZONE_LOAD

Dim OLD_CPU_SINCE
Dim CPU_SINCE_ALARM

Dim MyProc_Alternative

Dim OLD_MyProc
Dim OLD_Proc_CreationDate_2
Dim PROC_RUN_TIME
Dim OLD_PROC_RUN_TIME

Public EXIT_TRUE
Dim CPU_AVERAGE_BELOW_ALARM
Dim ALARM_COUNTER
Dim O_CPU_SINCE
Dim OLD_WIDTH_ADJUST
Dim ListView1_ListItems_Counter
Dim WinState

Dim PROCESS_ID_TO_USER As Long
Dim COUNTER_VAR
Dim COUNT_FOR_ITEM_LIST
Dim ALARM

Private wmi As Object, Locator As Object
Dim SampleRate As Long
Private PrevCpuTime As Single

Private Const MAX_PATH As Long = 260

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
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const LVM_FIRST = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Private Const LVS_EX_FULLROWSELECT = &H20
     
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

' --------------------------------------------------------------------
' --------------------------------------------------------------------
' SET OF DECLARE WORTH HAVE ALL CODE PROJECT BEGGINER
' --------------------------------------------------------------------
' --------------------------------------------------------------------
' BAS MODULE DECLARE
' Private Declare Function EnumChildWindows Lib "user32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
' --------------------------------------------------------------------
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Type MENUBARINFO
  cbSize As Long
  rcBar As Rect
  hMenu As Long
  hWndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type
Private MenuInfo As MENUBARINFO
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
'----------------------------------------------------------------------------------
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
' --------------------------------------------------------------------
' --------------------------------------------------------------------
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
' --------------------------------------------------------------------
Const GWL_WNDPROC = -4
Const GWL_HINSTANCE = -6
Const GWL_HWNDPARENT = -8
Const GWL_STYLE = -16
Const GWL_EXSTYLE = -20
Const GWL_USERDATA = -21
Const GWL_ID = -12
' --------------------------------------------------------------------
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowDLL Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function FindWindow2 Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cchar As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
' --------------------------------------------------------------------
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' In a module
'Private Const IDC_HAND As Long = 32649&
'Private Const IDC_WAIT As Long = 32514&
'Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
'Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long


' HARD SEARCH TO FIND 2026-05-01
'----------------------------------------
'Solved: How to get global cursor info using GetCursorInfo | Experts Exchange
'----------------------------------------
'https://www.experts-exchange.com/questions/21213133/How-to-get-global-cursor-info-using-GetCursorInfo.html
'----------------------------------------
'Private Type POINTAPI
'    x As Long
'    y As Long
'End Type
Private Type PCURSORINFO
    cbSize As Long
    flags As Long
    hCursor As Long
    ptScreenPos As POINTAPI
End Type
Private Declare Function GetCursorInfo Lib "user32.dll" (ByRef pci As PCURSORINFO) As Long





' --------------------------------------------------------------------
' FUNCTION SET HERE
' --------------------------------------------------------------------
Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
' --------------------------------------------------------------------
Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
' --------------------------------------------------------------------
' --------------------------------------------------------------------
Public Function GetSpecialFolder_Show_Script_Debug(CSIDL As Long) As String
Dim r As Long
On Error Resume Next
For r = 0 To 120
    If Trim(GetSpecialfolder(r)) <> "" Then
        'Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
        'AAX = GetSpecialfolder(R)
    End If
Next
End
End Function
Public Function GetSpecialfolder(CSIDL As Long) As String
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
Private Function GetTitle(ByVal hWnd As Long) As String
    Dim sTitle                As String
    sTitle = String(MAX_LENGTH, Chr$(0))
    GetWindowText hWnd, sTitle, MAX_LENGTH
    i = InStr(1, sTitle, Chr$(0))
    GetTitle = Mid$(sTitle, 1, i - 1)
End Function
Private Function GetClass(ByVal hWnd As Long)
    Dim sClass                As String
    sClass = String(MAX_LENGTH, Chr$(0))
    GetClassName hWnd, sClass, MAX_LENGTH
    i = InStr(1, sClass, Chr$(0))
    GetClass = Mid$(sClass, 1, i - 1)
End Function
Private Function WindowClass(ByVal hWnd As Long) As String
    Const MAX_LEN As Byte = 255
    Dim strBuff As String, intLen As Integer
    strBuff = String(MAX_LEN, vbNullChar)
    intLen = GetClassName(hWnd, strBuff, MAX_LEN)
    WindowClass = Left(strBuff, intLen)
End Function
Public Function StripNulls(OriginalStr As String) As String
        ' This removes the extra Nulls so String comparisons will work
        If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
        End If
        StripNulls = OriginalStr
End Function
Public Function StripNulls_2(OriginalStr) As String
        ' This removes the extra Nulls so String comparisons will work
        If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
        End If
        StripNulls_2 = OriginalStr
End Function
Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim l As Long
   Dim s As String
   l = GetWindowTextLength(hWnd)
   s = Space(l + 1)
   GetWindowText hWnd, s, l + 1
   GetWindowTitle = Left$(s, l)
End Function
Public Function GetWindowState(ByVal lngHwnd As Long) As Integer
    If IsWindow(lngHwnd) = 1 Then
        GetWindowState = -1
    If IsIconic(lngHwnd) <> 0 Then
        GetWindowState = vbMinimized
    ElseIf IsZoomed(lngHwnd) <> 0 Then
        GetWindowState = vbMaximized
    End If
    End If
End Function
Public Sub vLaunch(vFile As String)
    Dim Tri
    Tri = 1
    If Mid$(vFile, 1, 5) = "Http:" Then Tri = 1
    If Mid$(vFile, 2, 2) = ":\" Then Tri = 2
    If Right$(vFile, 1) = "\" Then Tri = 3
    Select Case Tri
    Case 1
        ShellExecute Me.hWnd, "open", vFile, vbNullString, vbNullString, conSwNormal
    Case 2
        vFile = "file:\\" + vFile
        ShellExecute Me.hWnd, "open", vFile, vbNullString, vbNullString, conSwNormal
        'Shell "C:\Program Files\Internet Explorer\iexplore.exe " + vFile, vbNormalFocus
    Case 3
        Shell "Explorer.exe /e," + vFile, vbNormalFocus
    End Select
End Sub








Private Sub Form_Load()
    ' Me.Icon. "D:\VB6\VB-NT\00_Best_VB_01\CPU % OF A PROGRAM\Goldzig.ico"
    ' LOOK END OF THIS SUB
    If InStr(Command$, "googledrivesync__cpu_percent_vb") = 0 Then
        SKIP_APP_PREVINSTANCE = True
    End If
    
    If App.PrevInstance = True Then End
    
    Call Project_Check_Date.VB_PROJECT_CHECKDATE("FORM LOAD")
    
    'Call Form_VB_Check_Date.VB_PROJECT_CHECKDATE
    
    'Dim MyProc As String
    
'    MyProc = "Flickr.exe"
'    PROCESS_ID_TO_USER = FindProcessID(MyProc)
'    If PROCESS_ID_TO_USER = 0 Then
'        MyProc = "Goodsync-v10.exe"
'        Call MNU_GOODSYNC_Click
'    End If
'
    'SampleRate = 1 'in seconds
    'Timer1.Interval = SampleRate * 1000
    Timer1.Interval = 1
    
    Set Locator = CreateObject("WbemScripting.SWbemLocator")
    Set wmi = Locator.ConnectServer
    
    Label1.Top = 0
    Label1.Height = 1000
    Label2.Top = Label1.Top + Label1.Height - 12
    Label2.Height = 12 '300
    Label2.Visible = False
    ListView2.Top = Label2.Top + Label2.Height
    ListView2.Height = 1800
    ListView1.Top = ListView2.Top + ListView2.Height
    ListView_INDEX.Top = ListView2.Top + ListView2.Height
    ListView_INDEX.Left = 0

'    List1.Top = Label1.Top + Label1.Height
'    List2.Top = Label1.Top + Label1.Height
'    List4.Top = Label1.Top + Label1.Height
'    List2.Left = 0
'    List2.Width = 750
'    List4.Left = List2.Width
'    List4.Width = 2700
'    List1.Left = List4.Width + List4.Left
'    List1.Height = 6000
'    List2.Height = List1.Height
'    List4.Height = List1.Height
'    List1.Width = 7800
    ListView1.Width = 12000
    ListView2.Width = (ListView1.Width + ListView_INDEX.Width) - 780
    ListView1.Height = 5900
    ListView_INDEX.Width = 500
    ListView_INDEX.Height = ListView1.Height
    ListView1.Left = ListView_INDEX.Width
    ListView2.Left = 0
    
    Me.Width = ListView1.Left + ListView1.Width + 250
    Me.Height = ListView1.Top + ListView1.Height + Menu_Height * Screen.TwipsPerPixelY + 580
    
    'Taskbar at the Bottom
    '---------------------
    Dim RectTaskbar As Rect
    HwndTask = FindWindow("Shell_TrayWnd", vbNullString)
    HwndResult = GetWindowRect(HwndTask, RectTaskbar)
    ' CENTER ON SCREEN COMPENSATE TASKBAR AT BOTTOM
    '----------------------------------------------
    If Screen.Height - RectTaskbar.Top * Screen.TwipsPerPixelY < 2500 Then
        Me.Left = Screen.Width / 2 - Me.Width / 2
'        Me.Top = (Screen.Height - ((RectTaskbar.Bottom - RectTaskbar.Top) * Screen.TwipsPerPixelY)) / 2 - Me.Height / 2
    Else
        Me.Left = Screen.Width / 2 - Me.Width / 2
'        Me.Top = Screen.Height / 2 - Me.Height / 2
    End If
    
    Me.Top = 100
    
    Label1.Width = ListView1.Left + ListView1.Width
    Label2.Width = ListView1.Left + ListView1.Width
    
    ListView1.Font.Name = "ARIAL"
    ListView1.Font.Size = 11
    ListView2.Font.Name = "ARIAL"
    ListView2.Font.Size = 11
    ListView_INDEX.Font.Name = "ARIAL"
    ListView_INDEX.Font.Size = 11
    
    With ListView_INDEX
        .ColumnHeaders.Add , "#", "#", 1980, lvwColumnLeft
        .View = lvwReport
    End With
    With ListView1
        .ColumnHeaders.Add , "TIME", "TIME", 1300, lvwColumnLeft
        .ColumnHeaders.Add , "CPU TIME", "CPU TIME", 1250, lvwColumnLeft
        .ColumnHeaders.Add , "CPU AVERAGE", "CPU AVERAGE", 1700, lvwColumnLeft
        .ColumnHeaders.Add , "CPU SINCE", "CPU SINCE", 1400, lvwColumnLeft
        .ColumnHeaders.Add , "CPU %", "CPU %", 1000, lvwColumnLeft
        .ColumnHeaders.Add , "RUN TIME", "RUN TIME", 1590, lvwColumnLeft
        .ColumnHeaders.Add , "PROCESS CREATION TIME", "PROCESS CREATION TIME", 3200, lvwColumnLeft
        .View = lvwReport
    End With
    
    With ListView2
        .ColumnHeaders.Add , "TIME", "TIME", 2000, lvwColumnLeft
        .ColumnHeaders.Add , "PROCESS", "PROCESS", 2000, lvwColumnLeft
        .ColumnHeaders.Add , "RUN TIME", "RUN TIME", 3000, lvwColumnLeft
        .ColumnHeaders.Add , "PROCESS CREATION TIME", "PROCESS CREATION TIME", 3200, lvwColumnLeft
        .View = lvwReport
    End With
    
    SetFullRowSelection ListView1.hWnd, True
    
    COUNT_FOR_ITEM_LIST = 20
    For i = 1 To COUNT_FOR_ITEM_LIST
        List2.AddItem i
    With ListView_INDEX
        Set LV2 = .ListItems.Add(, , i)
    End With
    'Call LV_AutoSizeColumn(ListView_INDEX, ListView_INDEX.ColumnHeaders.Item(1))
Next

CPU_AVERAGE_BELOW_ALARM = 1

Call LABEL1_SETTING

'ICON = D:\VB6\VB Icon's\Icons\02 All BlueSky Geo Gorts\Goldzig.ico

If Val(GetOperatingSystem) > 6 Then
    'MNU_AUDIO_ON.Caption = "AUDIO IS ON"
End If

'MNU_VB_ME.Caption = "VB_ME_Ver_2018_" + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision)) ' + " _ Matt.Lan@btinternet.com"
MNU_VB_ME.Caption = "VB_ME"

Timer_INIT.Enabled = True

If IsIDE = False Then Me.WindowState = vbMinimized

If InStr(Command$, "googledrivesync__cpu_percent_vb") > 0 Then
    Timer1.Enabled = False
    MsgBox "C:\PROGRAM FILES\GOOGLE\DRIVE\GOOGLEDRIVESYNC.EXE" + vbCrLf + vbCrLf + "HAS BEEN DO MESSENGER FROM" + vbCrLf + "PROCESS LASSO"
    End
    ' -------------------------------------------------------------------
    ' Call MNU_GOOGLEDRIVESYNC_Click
    ' WHEN SETUP PROCEsS LASSO NOT USER DOUBEL QUOTE P-L NOT REQUIRE THEM
    ' -------------------------------------------------------------------
    ' TRY MAKE CODE SO ONLY RUN ONCE WHEN CPU USER HIGH BY
    ' GOOGLEDRIVESYNC
    ' AND THEN WIRTE FILE
    ' AND IF LOW RESULT
    ' IF AGAIN AND NOT HIGH
    ' IF WAIT BEEN LONG WHEN ASK TRIGGER
    ' -------------------------------------------------------------------
End If

' -----------------------------------
'MyProc = "UnstopCpy.exe"
'Call MNU_UNSTOPPABLE_COPIER_Click

'MyProc = "FileLocatorPro.exe"
'Call MNU_FILE_LOCATOR_Click

'MyProc = "cmd.exe"
'Call MNU_CMD_Click

'MyProc = "AgentRansack.exe"
'Call MNU_AGENT_RANSACK_Click
' -----------------------------------

'If PROCESS_ID_TO_USER = 0 Then
'    MyProc = "CCleaner64.exe"
'    Call MNU_CCLEANER64_Click
'End If

If PROCESS_ID_TO_USER = 0 Then
    MyProc = "FreemakeVC.exe"
    Call Mnu_Freemaker_Click
End If

If PROCESS_ID_TO_USER = 0 Then
    MyProc = "DuplicateCleaner.exe"
    Call MNU_DuplicateCleaner_Click
End If
If PROCESS_ID_TO_USER = 0 Then
    MyProc = "DuplicateCleaner.exe"
    Call MNU_DuplicateCleanerV5_Click
End If

If PROCESS_ID_TO_USER = 0 Then
    MyProc = "NeroVision.exe"
    Call Mnu_Freemaker_Click
End If

If PROCESS_ID_TO_USER = 0 Then
    MyProc = "FormatConverter.exe"
    Call MNU_FormatConverter_EXE_Click
End If

If PROCESS_ID_TO_USER = 0 Then
    MyProc = "PowerDataRecovery.exe"
    Call POWER_DATA_RECOVERY_EXE_Click
End If

If PROCESS_ID_TO_USER = 0 Then
    MyProc = "DriverBooster.exe"
    Call MNU_DRIVERBOOSTER_Click
End If


If PROCESS_ID_TO_USER = 0 Then
    Call MNU_GoodSync2Go_Click
End If

If PROCESS_ID_TO_USER = 0 Then
    Call MNU_GOODSYNCV10_Click
End If

'PROCESS_ID_TO_USER = 0
If PROCESS_ID_TO_USER = 0 Then
    '4G COMPUTER __ "Goodsync.exe"
    Call MNU_GOODSYNC_Click
End If



If GetComputerName = "4-ASUS-GL522VW" Then
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"
End If
If GetComputerName = "8-MSI-GP62M-7RD" Then
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"
End If
If GetComputerName = "9-ASUS-G815LM" Then
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"
End If


' Load LAB_TIMEZONE
LAB_TIMEZONE_LOAD = True

If Command$ <> "" Then
    TIMER_MSGBOX_FROM_PROCESS_LASSO_OVER_MEMORY_SIZE.Enabled = True
End If

TIMER_MSGBOX_FROM_PROCESS_LASSO_OVER_MEMORY_SIZE.Enabled = True

End Sub

Private Sub MNU_CCLEANER64_Click()

    MyProc_Alternative = True
    MyProc_Alternative = False
    ' IF FALSE HERE AND THEN IT WILL CHECK AGAINST AVERAGE
    ' SET THE LEVEL THAT FALL BELOW ALARM WITH
    ' CPU_SINCE_ALARM = 1
    MyProc = "CCleaner64.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 1
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"


End Sub



Private Sub TIMER_MSGBOX_FROM_PROCESS_LASSO_OVER_MEMORY_SIZE_Timer()
    TIMER_MSGBOX_FROM_PROCESS_LASSO_OVER_MEMORY_SIZE.Enabled = False
    If InStr(UCase(Command$), "____ OVER_8 GB") = 0 Then Exit Sub
    MsgBox UCase(Command$), vbMsgBoxSetForeground, App.EXEName + ".EXE"
End Sub

Sub LABEL1_SETTING()

LAB1_VAR = ""
LAB1_VAR = LAB1_VAR + "If Process _" + MyProc + "_ PID # _" + Trim(Str(PROCESS_ID_TO_USER)) + "_ Less to IDLE CPU-Time State 1 Per Sec then Alarm Finished Done"
LAB1_VAR = LAB1_VAR + vbCrLf
LAB1_VAR = LAB1_VAR + "Detected When Average is Below < " + Format(CPU_AVERAGE_BELOW_ALARM, "0.0000")
LAB1_VAR = LAB1_VAR + " And CPU Average <=" + Str(CPU_SINCE_ALARM) + " And By Every Other Two __ Alarm Audio Indicated by + Symbol"
LAB1_VAR = LAB1_VAR + vbCrLf + "The CPU Time Ought To Be More But VB6 Only Checker With One CPU"

If MyProc_Alternative = True Then
    LAB1_VAR = ""
    LAB1_VAR = LAB1_VAR + "If Process _" + MyProc + "_ PID # _" + Trim(Str(PROCESS_ID_TO_USER)) + "_ Less to IDLE CPU-Time State 1 Per Sec then Alarm"
    LAB1_VAR = LAB1_VAR + vbCrLf
    LAB1_VAR = LAB1_VAR + "When Average is Less < Than Before with Run Count >=10 And By Every Other 4 "
    LAB1_VAR = LAB1_VAR + " And CPU Average <=" + Str(CPU_SINCE_ALARM)
    LAB1_VAR = LAB1_VAR + " _ Alarm Audio Indicated by + Symbol"
    LAB1_VAR = LAB1_VAR + vbCrLf + "The CPU Time Ought To Be More But VB6 Only Checker With One CPU"
End If

Label1.Caption = LAB1_VAR
Me.Caption = App.EXEName + " _ Ver. 2018_" + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision)) ' + " _ Matt.Lan@btinternet.com"

End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then
        MNU_MINI_DISPLAY_Click
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

End

On Error Resume Next
For Each Form In Forms
Err.Clear
    If Form.EXIT_TRUE = True Then
        'Form.Name
        If Err.Number = 0 Then
            Me.EXIT_TRUE = True
        End If
    End If
Next Form
On Error GoTo 0

If IsIDE = False Then
    If Me.WindowState <> vbMinimized And Me.EXIT_TRUE = False Then
        Me.WindowState = vbMinimized
        Cancel = True
        Exit Sub
    End If
End If



If Me.EXIT_TRUE = True Then Cancel = False


'SET ALL TIMERS IN ALL FORMS ENABLED TO =FALSE
On Error Resume Next
    For i = 0 To Forms.Count - 1
        For Each Control In Forms(i).Controls
            If InStr(UCase(Control.Name), "TIMER") > 0 Then
                'Debug.Print Control.Name
                Control.Enabled = False
            End If
        Next
    Next i
On Error GoTo 0


For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next Form

End

End Sub

Private Sub List3_KeyDown(KeyCode As Integer, Shift As Integer)
If IsIDE = True And KeyCode = 27 Then End
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
If IsIDE = True And KeyCode = 27 Then End

End Sub

Private Sub MNU_AGENT_RANSACK_Click()
    ' C:\Program Files\Mythicsoft\Agent Ransack\AgentRansack.exe
    MyProc_Alternative = True
    MyProc = "AgentRansack.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0
    Call LABEL1_SETTING
End Sub

Private Sub MNU_AUDIO_ON_Click()

If MNU_AUDIO_ON.Caption = "AUDIO IS ON" Then
    MNU_AUDIO_ON.Caption = "AUDIO IS OFF"
Else
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"
End If

End Sub

Private Sub MNU_CHKDSK_Click()
    
    MyProc_Alternative = True
    MyProc = "chkdsk.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0
    Call LABEL1_SETTING

End Sub

Private Sub MNU_CMD_Click()

    ' C:\Windows\System32\cmd.exe
    MyProc_Alternative = True
    MyProc = "cmd.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0
    Call LABEL1_SETTING
End Sub

Private Sub MNU_DLLHOST_EXE_WHEN_SET_SECRUITY_ON_FOLDER_Click()
    'C:\Windows\System32\dllhost.exe
    MyProc_Alternative = True
    MyProc_Alternative = False
    ' IF FALSE HERE AND THEN IT WILL CHECK AGAINST AVERAGE
    ' SET THE LEVEL THAT FALL BELOW ALARM WITH
    ' CPU_SINCE_ALARM = 1
    MyProc = "dllhost.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 4
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"
End Sub

Private Sub MNU_EXPLORER_EXE_Click()
    'C:\Windows\System32\dllhost.exe
    MyProc_Alternative = True
    MyProc_Alternative = False
    ' IF FALSE HERE AND THEN IT WILL CHECK AGAINST AVERAGE
    ' SET THE LEVEL THAT FALL BELOW ALARM WITH
    ' CPU_SINCE_ALARM = 1
    MyProc = "explorer.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 4
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"
End Sub

Private Sub MNU_EXIT_Click()
EXIT_TRUE = True
Unload Me
'End
End Sub

Private Sub MNU_FFMPEG_Click()
    MyProc_Alternative = True
    MyProc_Alternative = False
    ' IF FALSE HERE AND THEN IT WILL CHECK AGAINST AVERAGE
    ' SET THE LEVEL THAT FALL BELOW ALARM WITH
    ' CPU_SINCE_ALARM = 0.1
    MyProc = "ffmpeg.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0.1
    Call LABEL1_SETTING
End Sub

Private Sub MNU_FILE_LOCATOR_Click()

    'C:\Program Files\Mythicsoft\FileLocator Pro\FileLocatorPro.exe

    MyProc_Alternative = True
    MyProc = "FileLocatorPro.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0
    Call LABEL1_SETTING

End Sub

Private Sub MNU_FLICKR_Click()
    
    MyProc_Alternative = True
    MyProc = "Flickr.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"

End Sub

Private Sub MNU_FormatConverter_EXE_Click()
    MyProc_Alternative = True
    MyProc_Alternative = False
    ' IF FALSE HERE AND THEN IT WILL CHECK AGAINST AVERAGE
    ' SET THE LEVEL THAT FALL BELOW ALARM WITH
    ' CPU_SINCE_ALARM = 1
    MyProc = "FormatConverter.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0.04
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"
End Sub

Private Sub Mnu_Freemaker_Click()
    MyProc_Alternative = True
    MyProc_Alternative = False
    ' IF FALSE HERE AND THEN IT WILL CHECK AGAINST AVERAGE
    ' SET THE LEVEL THAT FALL BELOW ALARM WITH
    ' CPU_SINCE_ALARM = 1
    MyProc = "FreemakeVC.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 1
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"
End Sub

Private Sub Mnu_NeroVision_Click()
    MyProc_Alternative = True
    MyProc_Alternative = False
    ' IF FALSE HERE AND THEN IT WILL CHECK AGAINST AVERAGE
    ' SET THE LEVEL THAT FALL BELOW ALARM WITH
    ' CPU_SINCE_ALARM = 1
    MyProc = "NeroVision.exe" ' C:\Program Files (x86)\Nero\Nero 2018\Nero Vision\NeroVision.exe
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 1
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"
End Sub


Private Sub MNU_DRIVERBOOSTER_Click()
    ' C:\Program Files (x86)\IObit\Driver Booster\7.4.0\DriverBooster.exe
    MyProc_Alternative = True
    MyProc_Alternative = False
    ' IF FALSE HERE AND THEN IT WILL CHECK AGAINST AVERAGE
    ' SET THE LEVEL THAT FALL BELOW ALARM WITH
    ' CPU_SINCE_ALARM = 1
    MyProc = "DriverBooster.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 1
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"
End Sub

Private Sub MNU_GOODSYNC_Click()
    
    MyProc_Alternative = True
    MyProc_Alternative = False
    MyProc = "Goodsync.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0.15
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"

End Sub

Private Sub MNU_GoodSync2Go_Click()
    MyProc_Alternative = True
    MyProc_Alternative = False
    MyProc = "GoodSync2Go.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0.03
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"
    
End Sub

Private Sub MNU_GOODSYNCV10_Click()
    MyProc_Alternative = True
    MyProc_Alternative = False
    MyProc = "Goodsync-v10.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0.15
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"
End Sub

Private Sub MNU_GOOGLEDRIVESYNC_Click()

    MyProc_Alternative = True
    MyProc_Alternative = False
    MyProc = "googledrivesync.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0.1
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"

End Sub

Private Sub MNU_HANDBRAKE_Click()

    MyProc_Alternative = True
    MyProc_Alternative = False
    MyProc = "HandBrake.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0.1
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"

End Sub

Private Sub MNU_HASH_MY_FILE_Click()

    MyProc_Alternative = True
    MyProc_Alternative = False
    MyProc = "hashmyfiles.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0.1
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"

End Sub

Private Sub MNU_INSYNC_Click()

    MyProc_Alternative = True
    MyProc_Alternative = False
    MyProc = "Insync.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0.1
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"

End Sub

Private Sub MNU_KILL_WSCRIPT_Click()
    Shell "TASKKILL /F /IM WSCRIPT.EXE /T", vbHide
End Sub

Private Sub MNU_MINI_DISPLAY_Click()
If Trim(LISTVIEW_ITEM_2) = "" Then Exit Sub
Load FORM_VOL
FORM_VOL.Show
End Sub


Private Sub MNU_OUTLOOK_OFFICE_Click()
    MyProc_Alternative = True
    MyProc_Alternative = False
    'C:\Program Files (x86)\Microsoft Office\Office12\OUTLOOK.EXE
    MyProc = "OUTLOOK.EXE"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0.8
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"
End Sub


Private Sub POWER_DATA_RECOVERY_EXE_Click()
    MyProc_Alternative = True
    MyProc_Alternative = False
    MyProc = "PowerDataRecovery.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0.4
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"
End Sub


Private Sub MNU_PARTITION_WIZARD_Click()
    
    MyProc_Alternative = True
    MyProc = "partitionwizard.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"

End Sub

Private Sub MNU_UNSTOPPABLE_COPIER_Click()
    
    'C:\HBCD\Programs\Programs\UnstopCpy.exe
    
    MyProc_Alternative = True
    MyProc = "UnstopCpy.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"
    
End Sub

Private Sub MNU_SOUND_TEST_Click()
    
    If MMControl1.Filename = "" Then
            MMControl1.Filename = App.Path + "\Windows Ringin.wav"
            MMControl1.Notify = True
            MMControl1.Wait = True
            MMControl1.Shareable = False
            MMControl1.DeviceType = "WaveAudio"
            MMControl1.Command = "Open"
    End If
        
    SOUND_PLAY_TIMER.Enabled = True
    
End Sub

Private Sub VB_RUN_NOT_WHEN_IDE_AND_THEN_SHOW()
' -----------------------------------------------------------------
' IS SHORT CUT NOT SET ADMINISTATOR FOR THIS PARTICALUR CODE IN EXE
' MODE IT WILL NOT RUN LOAD ITSELF ON THE VB COMPILER
' BUT AND WILL RUN OTHER PROGRAM
' AS TALK OTHER PROGRAM ARE OKAY BUT THEN MAYBE SET AS ADMIN
' AMSWER GIVEN
' AUG 28 SUN
' -----------------------------------------------------------------
If 1 = 1 And IsIDE = True Then
    MNU_VB.Enabled = False
    Beep
    MsgBox "NOT WHEN ISIDE = TRUE"
    Exit Sub
End If

Dim ReturnHwnd As Long
Dim i
'VB ONLY WANTS THE 1ST OF THE 2 HWND
'ReturnHwnd = FindWindowPartVB("ClipBoard Logger - Microsoft Visual Basic[")
'------------------------------------------------
'DONT NEED ABOVE USE THIS
X1 = FindWindow("wndclass_desked_gsk", vbNullString)
'------------------------------------------------
'FIND THE WINDOW PRIZE TREASURE CHEST BUST BLOSSOM
'TRAIN SPOTTER
'------------------------------------------------
'-----------------------------------------------------------
'CHECK CLASS NAME IS VB IDE WINDOW WE WANT OR IF ANOTHER NOT
'-----------------------------------------------------------
X2 = GetWindowTitle(X1)
If InStr(X2, " - Microsoft Visual Basic [") > 0 Then
    'RUN A WINDOW SPY
    WIN_SPY_NAME = "ClipBoard Logger"
    If InStr(X2, WIN_SPY_NAME) > 0 Then
    
        MsgBox "DON'T RUN VB IDE - LOADED"
        i = GetWindowState(X1)
        If i = vbMinimized Then
            SHOWVAR = SW_SHOWDEFAULT
            ShowWindow ReturnHwnd, SHOWVAR
        End If
        
        EXIT_TRUE = True
        Unload Me
        Exit Sub
    End If
End If
'------------------------------------------------
'BETTER LINE NEXT DON'T USE VB MENU IF ISIDE
'------------------------------------------------
'TEMP WORK AROUND
'OVER DRIVE
'OVER RIDE
'------------------------------------------------
'FINDWINDOW PART PROBLEM IN EXE AND WHEN IN WIN 7
'------------------------------------------------
'WIN 7 PROBLEM MUST USE EXTRA LAST LEFT SQUARE BRACKET OF SERACH END ABOVE
'------------------------------------------------
If ReturnHwnd > 0 Then
    If MsgBox("ReturnHwnd" + Str(ReturnHwnd) + " -- " + "MeHwnd" + Str(Me.hWnd) + vbCrLf + "VB Code " + vbCrLf + WIN_SPY_NAME + vbCrLf + " already Running - Sure Want to Run VB IDE", vbYesNo) = vbYes Then
        WANT_TO_RUN_ANYWAY = True
    End If
End If

If ReturnHwnd > 0 Then
    i = GetWindowState(ReturnHwnd1)
    If i = vbMinimized Then
'        SHOWVAR = SW_RESTORE
'        SHOWVAR = SW_SHOW
'        SHOWVAR = SW_SHOWNA
'        SHOWVAR = SW_SHOWDEFAULT
'        SHOWVAR = SW_SHOWNORMAL
'        SHOWVAR = SW_SHOWMAXIMIZED
        SHOWVAR = SW_SHOWDEFAULT
        ShowWindow ReturnHwnd, SHOWVAR
        'GUESS MAYBE
        'SetForegroundWindow (ReturnHwnd)
        DoEvents
    End If
   
    'MADE REDUNDANT CODE BY CONDICTION HERE AND ABOVE
    If WANT_TO_RUN_ANYWAY = False Then
        MsgBox "EXIT AS FOUND WINODW OF VB AND PUT TO SHOW FOCUS"
        EXIT_TRUE = True
        Unload Me
        Exit Sub
    End If
End If

Dim CODER_VBP_FILE_NAME_2
Dim VB_1, VB_2, VB_3
VB_1 = "C:\PROGRAM FILES\Microsoft Visual Studio\VB98\VB6.EXE"
VB_2 = "C:\PROGRAM FILES (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
If Dir(VB_1) <> "" Then VB_3 = VB_1
If Dir(VB_2) <> "" Then VB_3 = VB_2
'------------------------------------------------------
'ADD MICROSFOT SCRIPTING RUNTIME FOR HERE IN REFERENCES
'------------------------------------------------------
Dim OBJSHELL
Set OBJSHELL = CreateObject("Wscript.Shell")
CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
OBJSHELL.Run """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
Set OBJSHELL = Nothing
EXIT_TRUE = True
Unload Me
End Sub
Private Sub Mnu_VB_ME_Click()
    
    Call VB_RUN_NOT_WHEN_IDE_AND_THEN_SHOW
    
    Dim CODER_VBP_FILE_NAME_2
    Dim VB_1, VB_2, VB_3
    VB_1 = "C:\PROGRAM FILES\Microsoft Visual Studio\VB98\VB6.EXE"
    VB_2 = "C:\PROGRAM FILES (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
    If Dir(VB_1) <> "" Then VB_3 = VB_1
    If Dir(VB_2) <> "" Then VB_3 = VB_2
    '------------------------------------------------------
    'ADD MICROSFOT SCRIPTING RUNTIME FOR HERE IN REFERENCES
    '------------------------------------------------------
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
    Call CHECK_INTEGRITY_OF_VISUAL_BASIC_PROJECT_VBP(CODER_VBP_FILE_NAME_2)
    
    Dim OBJSHELL
    Set OBJSHELL = CreateObject("WSCRIPT.SHELL")
    OBJSHELL.Run """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
    Set OBJSHELL = Nothing
    
    EXIT_TRUE = True
    Unload Me
    Exit Sub
End Sub
Private Sub MNU_VB_FOLDER_Click()
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
End Sub
Sub CHECK_INTEGRITY_OF_VISUAL_BASIC_PROJECT_VBP(CODER_VBP_FILE_NAME_2)
    ' --------------------------------------------------------------------
    ' FIND IF THE IS NOT CORRECT VERSION AND SIMPLE PUT CORRECT
    ' ONE MY COMPUTER SEEM PROBLEM WITH BASIC
    ' TALK HAS WRONG VERISON ONCE EVERY FEW DAY
    ' AND EDIT IT TO WHAT NORM VERISON SUPPOSED TO BE AND FINE
    ' DISCOVERY -- IT SEEM TO HAPPEN JUST AFTER A COMPILE SOMETIME_
    ' [ Monday 15:37:30 Pm_17 June 2019 ]
    ' --------------------------------------------------------------------
    ' --------------------------------------------------------------------
    Dim r, FR
    Dim VAR_STRING As String

    FR = FreeFile
    Open CODER_VBP_FILE_NAME_2 For Binary As FR
        VAR_STRING = Space(LOF(FR))
        Get #FR, , VAR_STRING
    Close FR
    ' --------------------------------------------------------------
    ' # .VBP # .VBP" -- # MSCOMC -- # OCX -- # MSCOMCTL # CTL # .CTL
    ' --------------------------------------------------------------
    ' FOUND TO BE ERROR
    ' --------------------------------------------------------------
    ' Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0; MSCOMCTL.OCX
    ' --------------------------------------------------------------
    ' CORRECT VALUE
    ' --------------------------------------------------------------
    ' Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0; MSCOMCTL.OCX
    ' --------------------------------------------------------------
    XR1 = InStr(UCase(VAR_STRING), UCase("A1}#2.1; MSCOMCTL.OCX"))
    XR2 = InStr(UCase(VAR_STRING), UCase("A1}#2.#")) ' ---- "A1}#2.#0; mscomctl.OCX"
    XR3 = InStr(UCase(VAR_STRING), UCase("#0; MSCOMCTL.OCX"))
    If XR1 = 0 And XR2 > 0 And XR3 > 0 Then
        GET_01 = InStr(UCase(VAR_STRING), UCase("MSCOMCTL.OCX")) - XR2
        Mid(VAR_STRING, XR2, GET_01 + Len("MSCOMCTL.OCX")) = "A1}#2.1#0; MSCOMCTL.OCX"
        ' MsgBox "MSCOMCTL.OCX" + vbCrLf + "WRONG VERSION -- CHANGE TO" + vbCrLf + "#2.1# -- MSCOMCTL.OCX"), vbMsgBoxSetForeground
        GO_NEXT_IN = True
    End If
    If GO_NEXT_IN = True Then
        If Dir(CODER_VBP_FILE_NAME_2) <> "" Then
            Kill CODER_VBP_FILE_NAME_2
        End If
        FR = FreeFile
        Open CODER_VBP_FILE_NAME_2 For Binary As FR
            Put #FR, , VAR_STRING
        Close FR
    End If
End Sub
Private Sub MNU_FOLDER_DIGITAL_STILL_CAMERA_Click()
    ' Beep
    Me.WindowState = vbMinimized
    Dim FOLDER_Path
    FOLDER_Path = "D:\DSC\2015+SONY\2020 CyberShot HX60V\JPG"
    Dir1.Path = FOLDER_Path
    Dim TD
    For r = 0 To Dir1.ListCount - 1
        FIND_PATH_1 = Dir1.List(r)
        FIND_PATH_2 = Mid(FIND_PATH_1, InStrRev(FIND_PATH_1, "\"))
        If Mid(FIND_PATH_2, 1, 3) = "\20" Then
            TD = Dir1.List(r)
        End If
    Next
    FOLDER_Path = TD
    
    Shell "EXPLORER /SELECT, """ + FOLDER_Path + """", vbNormalFocus
    EXIT_TRUE = True
    ' Unload Me
    ' End
End Sub
Private Sub MNU_CAMERA_VIDEO_FOLDER_4G_Click()
    Me.WindowState = vbMinimized
    Shell "EXPLORER \\4-asus-gl522vw\4_asus_gl522vw_10_1_samsung_1tb\DSC\2015+SONY_MP4\2020 CyberShot HX60V __ MP4", vbMaximizedFocus
End Sub
Private Sub MNU_ADMINISTRATOR_Click()
    ' Call MNU_ADMINISTRATOR_Click
    MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR MODE ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "MODE ON", "MODE OFF")
    End If
End Sub
'-------------------------------------------
'-------------------------------------------


Private Sub MNU_WSCRIPT_Click()
    MyProc_Alternative = True
    MyProc = "Wscript.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0
    Call LABEL1_SETTING
End Sub
Private Sub MNU_DuplicateCleanerV5_Click()
    
    MyProc_Alternative = True
    MyProc_Alternative = False
    MyProc = "Duplicate Cleaner 5.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0.5
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"


End Sub

Private Sub MNU_DuplicateCleaner_Click()
    
'    MyProc_Alternative = True
'    MyProc = "DuplicateCleaner.exe"
'    PROCESS_ID_TO_USER = FindProcessID(MyProc)
'    CPU_SINCE_ALARM = 0
'    Call LABEL1_SETTING
    
    MyProc_Alternative = True
    MyProc_Alternative = False
    MyProc = "DuplicateCleaner.exe"
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    CPU_SINCE_ALARM = 0.5
    Call LABEL1_SETTING
    MNU_AUDIO_ON.Caption = "AUDIO IS ON"

End Sub


Private Sub TIMER_EXE_PICKUP_Timer()
    
    TIMER_EXE_PICKUP.Enabled = False
    PROCESS_ID_TO_USER = FindProcessID(MyProc)
    Call LABEL1_SETTING

End Sub

Private Sub Timer_INIT_Timer()
    
    On Error GoTo ENDER
    Timer_INIT.Enabled = False
    
    
    List3.Top = 10000
    
    List3.SetFocus
    
    If IsIDE = False Then Me.WindowState = vbMinimized
    
    Call SET_MENU_PADD_WORK
    
ENDER:
   
End Sub


Private Sub Timer1_Timer()
    
    If Timer1.Interval <> 1000 Then Timer1.Interval = 1000
    
    If PROCESS_ID_TO_USER = 0 Then
        TIMER_EXE_PICKUP.Enabled = True
        
        ' SOUND IF NOT A PROCESS
        
        ALARM = True
        
        ALARM_COUNTER = ALARM_COUNTER + 1

        If ALARM_COUNTER Mod 2 <> 0 Then ALARM = False
            
        If ALARM = True And InStr(MNU_AUDIO_ON.Caption, "AUDIO IS ON") > 0 Then
            If MMControl1.Filename = "" Then
                MMControl1.Filename = App.Path + "\Windows Ringin.wav"
                MMControl1.Notify = True
                MMControl1.Wait = True
                MMControl1.Shareable = False
                MMControl1.DeviceType = "WaveAudio"
                MMControl1.Command = "Open"
            End If
            SOUND_PLAY_TIMER.Enabled = True
        End If
        
        
        Exit Sub
    End If
    
    Dim Procs As Object, Proc As Object
    Dim CpuTime  As Single, Utilization As Single
    Set Procs = wmi.InstancesOf("Win32_Process")


'Set objService = GetObject("winmgmts:")
'Set objProcess = GetObject("winmgmts:{impersonationLevel=impersonate}//localhost")
'Set objItems = objProcess.ExecQuery("Select PercentProcessorTime from Win32_PerfFormattedData_PerfProc_Process where IDProcess=2684")
'For Each objItem In objItems
'Debug.Print objItem.PercentProcessorTime
'Next

    PROCESS_ID_TO_USER_FOUND = False
    For Each Proc In Procs

'        If Proc.ProcessID <> 0 Then  'System Idle Process
'
'        End If
        
        If Proc.ProcessID = PROCESS_ID_TO_USER Then
            PROCESS_ID_TO_USER_FOUND = True
            CpuTime = Proc.KernelModeTime / 10000000
            
            If PrevCpuTime <> 0 Then
                
                'Utilization = 1 - (CpuTime - PrevCpuTime) / SampleRate
                
                Proc_CreationDate = Mid(Proc.CreationDate, 1, 4) + "-"
                Proc_CreationDate = Proc_CreationDate + Mid(Proc.CreationDate, 5, 2) + "-"
                Proc_CreationDate = Proc_CreationDate + Mid(Proc.CreationDate, 7, 2) + " "
                Proc_CreationDate = Proc_CreationDate + Mid(Proc.CreationDate, 9, 2) + ":"
                Proc_CreationDate = Proc_CreationDate + Mid(Proc.CreationDate, 11, 2) + ":"
                Proc_CreationDate = Proc_CreationDate + Mid(Proc.CreationDate, 13, 2)
                Proc_CreationDate_3 = Proc_CreationDate + " ." + Mid(Proc.CreationDate, InStr(Proc.CreationDate, ".") + 1)
                Proc_CreationDate_3 = Mid(Proc_CreationDate_3, 1, InStr(Proc_CreationDate_3, "+") - 1)
                
                Proc_CreationDate_2 = DateValue(Proc_CreationDate) + TimeValue(Proc_CreationDate)
                
                If OLD_Proc_CreationDate_2 <> Proc_CreationDate_2 And OLD_Proc_CreationDate_2 <> 0 Then
                    LockWindowUpdate ListView2.hWnd
                    With ListView2
                        Set LV2 = .ListItems.Add(, , Format(Now, "DD-MM-YYYY - HH-MM-SS"))
                        LV2.SubItems(1) = OLD_MyProc
                        LV2.SubItems(2) = OLD_PROC_RUN_TIME
                        LV2.SubItems(3) = OLD_Proc_CreationDate_2
                    End With
                    LockWindowUpdate 0
                    For R_L = 1 To ListView2.ColumnHeaders.Count
                        Call LV_AutoSizeColumn(ListView2, ListView2.ColumnHeaders.item(R_L))
                    Next
                End If
                OLD_Proc_CreationDate_2 = Proc_CreationDate_2
                OLD_MyProc = MyProc
                
                CpuTime = Proc.KernelModeTime / 10000000
                'PROCESSOR TIME COULD BE HIGHER BUT VB6 IS ONLY LOOKING AT ONE CPU MAYBE
                
                F1 = Int((CpuTime / 60 ^ 2) / 24)    'DAY
                F2 = Int((CpuTime / 60 ^ 2)) Mod 24  'HOUR
                F3 = Int(CpuTime / 60) Mod 60        'MINUTE
                F4 = CpuTime Mod 60                  'SECONDS
                
                '0H 16M 14S
                '1H 55M
                
                DATEDIFF_2 = DateDiff("S", Proc_CreationDate_2, Now)
                
                F1 = Int((DATEDIFF_2 / 60 ^ 2) / 24)    'DAY
                F2 = Int((DATEDIFF_2 / 60 ^ 2)) Mod 24  'HOUR
                F3 = Int(DATEDIFF_2 / 60) Mod 60        'MINUTE
                F4 = DATEDIFF_2 Mod 60                  'SECONDS
                PROC_RUN_TIME = Format(F1, "00") & "d " & Format(F2, "00") & "h " & Format(F3, "00") & "m " & Format(F4, "00") & "s"
                If Format(F1, "00") = "00" Then
                    PROC_RUN_TIME = Replace(PROC_RUN_TIME, "00d ", "")
                End If
                
                OLD_PROC_RUN_TIME = PROC_RUN_TIME
'                If FORMAT(F2,"00") = "00" Then
'                    PROC_RUN_TIME = Replace(PROC_RUN_TIME, "0h ", "")
'                End If
'                If FORMAT(F3,"00") = "00" Then
'                    PROC_RUN_TIME = Replace(PROC_RUN_TIME, "0m ", "")
'                End If
                'Debug.Print Str(CpuTime) & "--" & Str(F1) & "--" & Str(F2) & "--" & Str(F3) & "--" & Str(F4) & "--" & Proc_CreationDate
                
                
                'For Each Process In GetObject("winmgmts:{impersonationLevel=impersonate}//localhost").ExecQuery("Select PercentProcessorTime,IDProcess from Win32_PerfFormattedData_PerfProc_Process where IDProcess=" + Trim(Str(PROCESS_ID_TO_USER)) + ",")
                '    LISTVIEW_ITEM_4 = Proc.PercentProcessorTime
                'Next
                
                '----------------------------
                'NOT RUN AT MOMENT OVER HEADS WITH SWITCH TO AND RETRY MSGBOX
                'SET NOT RUN WHEN MINIMIZED
                '----------------------------
                PercentProcessorTime_VAR = "0"
                If Me.WindowState <> 1 And 1 = 2 Then
                    Set objProcess = GetObject("winmgmts:{impersonationLevel=impersonate}//localhost")
                    Set objItems = objProcess.ExecQuery("Select PercentProcessorTime from Win32_PerfFormattedData_PerfProc_Process where IDProcess=+" + Format(PROCESS_ID_TO_USER))
                    For Each objItem In objItems
                        'Debug.Print objItem.PercentProcessorTime
                        PercentProcessorTime_VAR = objItem.PercentProcessorTime
                        
                    Next
                End If
                
                If O_CPU_SINCE > 0 Then
                    CPU_SINCE = CpuTime - O_CPU_SINCE
                Else
                    CPU_SINCE = 0
                End If
                O_CPU_SINCE = CpuTime
                If CPU_SINCE < 0 Then CPU_SINCE = 0
                
                CPU_SINCE_TOTAL = 0
                For r = 1 To ListView1.ListItems.Count
                    WO = ListView1.ListItems(r).SubItems(3)
                    CPU_SINCE_TOTAL = CPU_SINCE_TOTAL + Val(WO)
                Next
                CPU_SINCE_TOTAL = CPU_SINCE_TOTAL + Val(CPU_SINCE)
                If ListView1.ListItems.Count = 20 Then
                    WO = ListView1.ListItems(1).SubItems(3)
                    CPU_SINCE_TOTAL = CPU_SINCE_TOTAL - Val(WO)
                End If
                
                
                'LISTVIEW_ITEM_1 = Mid(Format(Now, "YYYY-MM-DD_ _HH-MM-SS a/p"), 1, 21) + "  "
                LISTVIEW_ITEM_1 = Mid(Format(Now, "HH-MM-SS a/p"), 1, 10) + "  "
                LISTVIEW_ITEM_2 = Format(CpuTime, "0.0000") + "  "
                LISTVIEW_ITEM_3 = Format(CPU_SINCE, "0.0000000") + "         "
                LISTVIEW_ITEM_4 = Format(CPU_SINCE_TOTAL, "0.0000") + "               "
                LISTVIEW_ITEM_5 = PercentProcessorTime_VAR + " %      "
                LISTVIEW_ITEM_6 = PROC_RUN_TIME + "  "
                LISTVIEW_ITEM_7 = Proc_CreationDate_3 + "  "
                
                
                LockWindowUpdate ListView1.hWnd
                'ListView1.ListItems(1).SubItems (1)
                With ListView1
                    Set LV2 = .ListItems.Add(, , LISTVIEW_ITEM_1)
                    LV2.SubItems(1) = LISTVIEW_ITEM_2
                    LV2.SubItems(3) = LISTVIEW_ITEM_3 ' 3 2 SWAP
                    LV2.SubItems(2) = LISTVIEW_ITEM_4
                    LV2.SubItems(4) = LISTVIEW_ITEM_5
                    LV2.SubItems(5) = LISTVIEW_ITEM_6
                    LV2.SubItems(6) = LISTVIEW_ITEM_7
                End With
                LockWindowUpdate 0
                
                ' Load LAB_TIMEZONE
                If LAB_TIMEZONE_LOAD = True Then
                    LAB_TIMEZONE_LOAD = False
                    Load FORM_VOL
                End If
                
                
                List1.AddItem CpuTime 'Format(Utilization, "0.0000%")
                List4.AddItem Mid(Format(Now, "YYYY-MM-DD_ _HH-MM-SS a/p"), 1, 21)

                If List1.ListCount > COUNT_FOR_ITEM_LIST Then List1.RemoveItem (0)
                If List4.ListCount > COUNT_FOR_ITEM_LIST Then List4.RemoveItem (0)

                If ListView1.ListItems.Count > COUNT_FOR_ITEM_LIST Then
                    LockWindowUpdate ListView1.hWnd
                    ListView1.ListItems.Remove (1)
                    LockWindowUpdate 0
                End If
            
            End If
            PrevCpuTime = CpuTime
        End If
    Next
    
'    NORMAL_RUNNER = True
'
'    If MyProc = "UnstopCpy.exe" Then NORMAL_RUNNER = False
'    If MyProc = "FileLocatorPro.exe" Then NORMAL_RUNNER = False
'    If MyProc = "AgentRansack.exe" Then NORMAL_RUNNER = False
'    If MyProc = "Goodsync-v10.exe" Then NORMAL_RUNNER = False
'    If MyProc = "ffmpeg.exe" Then NORMAL_RUNNER = False
'    If MyProc = "cmd.exe" Then NORMAL_RUNNER = False
'    If MyProc = "chkdsk.exe" Then NORMAL_RUNNER = False
    
    If MyProc_Alternative = False Then
        ALARM = False
        SET_GO = False
        
        
        ' ---------------------------------------------------------
        ' TAKE OUT THIS ONE -- CPU_SINCE = 0 -- MOSTLY WHENEVER = 0 AND MOD NUMBER LIKE TYPE THING
        ' EVERY OTHER 1 -- MAKE RINGER
        ' ---------------------------------------------------------
        'If CPU_SINCE = 0 And OLD_CPU_SINCE = 0 And CPU_SINCE_TOTAL < CPU_AVERAGE_BELOW_ALARM Then
     
'        If OLD_CPU_SINCE = 0 And CPU_SINCE_TOTAL < CPU_AVERAGE_BELOW_ALARM Then
'            SET_GO = True
'        End If
        
        OLD_CPU_SINCE = CPU_SINCE
        
        If CPU_SINCE_ALARM > 0 And CPU_SINCE_TOTAL < CPU_SINCE_ALARM Then
            SET_GO = True
        End If
        If SET_GO = True Then
            If ListView1.ListItems.Count > 18 - 1 Then
                ALARM_COUNTER = ALARM_COUNTER + 1
                If ALARM_COUNTER > 2 Then ALARM_COUNTER = 1
                If ALARM_COUNTER = 1 Then ALARM = True
                        
                If ListView1.ListItems.Count > 0 And ALARM = True Then
                    WO = Trim(ListView1.ListItems(ListView1.ListItems.Count).SubItems(2))
                    ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = WO + " +" + "             "
                    Call LV_AutoSizeColumn(ListView1, ListView1.ColumnHeaders.item(4))
                End If
            End If
        Else
            ALARM = False
            COUNTER_VAR = 0
            ALARM_COUNTER = 0
        End If
    End If
    
    If MyProc_Alternative = True Then
        ALARM = False
        START_LOOKER = 18
        START_LOOKER = START_LOOKER + 1
        If ListView1.ListItems.Count >= START_LOOKER Then
            ALARM = True
            RUNNING = False
            XV = (ListView1.ListItems.Count - START_LOOKER) + 1
            If XV <= 0 Then XV = 1
            
            OWO = Val(Trim(ListView1.ListItems(XV).SubItems(3)))
            
            For r = XV To ListView1.ListItems.Count
                RUNNING = True
                WO = Val(Trim(ListView1.ListItems(r).SubItems(3)))
                If WO > OWO Then
                    ALARM = False
                End If
                OWO = WO
            Next
            If RUNNING = False Then ALARM = False
                            
            If ALARM = False Then ALARM_COUNTER = -1
            If ALARM = True Then
                ALARM_COUNTER = ALARM_COUNTER + 1
            End If

            If ALARM_COUNTER Mod 8 <> 0 Then ALARM = False
            
            
            If ALARM = True Then
                WO = Trim(ListView1.ListItems(ListView1.ListItems.Count).SubItems(2))
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = WO + " +" + "             "
                Call LV_AutoSizeColumn(ListView1, ListView1.ColumnHeaders.item(4))
            End If
            End If
    End If
    
    
    WIDTH_ADJUST = Len(LISTVIEW_ITEM_1) + Len(LISTVIEW_ITEM_2) + Len(LISTVIEW_ITEM_3) + Len(LISTVIEW_ITEM_4) + Len(LISTVIEW_ITEM_5) + Len(LISTVIEW_ITEM_6) + Len(LISTVIEW_ITEM_7)
    If OLD_WIDTH_ADJUST <> WIDTH_ADJUST Then
        For R_L = 1 To ListView1.ColumnHeaders.Count
            Call LV_AutoSizeColumn(ListView1, ListView1.ColumnHeaders.item(R_L))
        Next
    End If
    OLD_WIDTH_ADJUST = WIDTH_ADJUST
    
    Label2.Caption = "CPU SINCE TOTAL AVERAGE =" + Str(CPU_SINCE_TOTAL)
    
    If PROCESS_ID_TO_USER_FOUND = False Then
        COUNTER_VAR = 0
        ALARM_MAIN = False
        TIMER_EXE_PICKUP.Enabled = True
        WinState = False
    End If
    
    
    APOINTER = 0
    Dim typCI As PCURSORINFO
    typCI.cbSize = Len(typCI)
    If GetCursorInfo(typCI) <>  0 Then
'        Debug.Print typCI.flags, typCI.hCursor, typCI.ptScreenPos.x &  ; "," &  typCI.ptScreenPos.y
'        Debug.Print typCI.hCursor
        APOINTER = typCI.hCursor
    End If
    
    ' VBHOURGLASS = 11


    
    If ALARM = True And InStr(MNU_AUDIO_ON.Caption, "AUDIO IS ON") > 0 And APOINTER <> 65543 Then
        If MMControl1.Filename = "" Then
            MMControl1.Filename = App.Path + "\Windows Ringin.wav"
            MMControl1.Notify = True
            MMControl1.Wait = True
            MMControl1.Shareable = False
            MMControl1.DeviceType = "WaveAudio"
            MMControl1.Command = "Open"
        End If
        SOUND_PLAY_TIMER.Enabled = True
    End If
    
    If ALARM = True Then
        If WinState = False Then
            Me.SetFocus
            'Me.WindowState = vbNormal
            WinState = True
        End If
    End If
    
    'REF
    '----
    'Cpu Utilization in vb6-VBForums
    'http://www.vbforums.com/showthread.php?718429-Cpu-Utilization-in-vb6
    '----
End Sub




Private Function FindProcessID(ByVal pExename As String) As Long

    Dim ProcessID As Long, hSnapShot As Long
    Dim uProcess As PROCESSENTRY32, rProcessFound As Long
    Dim Pos As Integer, szExename As String
    
    ' Create snapshot of current processes
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    ' Check if snapshot is valid
    If hSnapShot = -1 Then
        Exit Function
    End If
    'Initialize uProcess with correct size
    uProcess.dwSize = Len(uProcess)
    'Start looping through processes
    rProcessFound = ProcessFirst(hSnapShot, uProcess)
    Do While rProcessFound
        Pos = InStr(1, uProcess.szExeFile, vbNullChar)
        If Pos Then
            szExename = Left$(uProcess.szExeFile, Pos - 1)
        End If
        If LCase$(szExename) = LCase$(pExename) Then
            'Found it
            ProcessID = uProcess.th32ProcessID
            Exit Do
          Else
            'Wrong, so continue looping
            rProcessFound = ProcessNext(hSnapShot, uProcess)
        End If
    Loop
    CloseHandle hSnapShot
    FindProcessID = ProcessID
    
    'REF
    '----
    '[RESOLVED] Process ID from exe name-VBForums
    'http://www.vbforums.com/showthread.php?571312-RESOLVED-Process-ID-from-exe-name
    '----
End Function

Private Sub SOUND_PLAY_TIMER_Timer()
        
        If MMControl1.Mode = 526 Then Exit Sub
        
        MMControl1.Wait = True
        MMControl1.Command = "PREV"
        MMControl1.Command = "PLAY"

        SOUND_PLAY_TIMER.Enabled = False

'  -->        If MMControl1.Mode = 525 Then RichText1.Text = "Stopped."
'  -->        If MMControl1.Mode = 526 Then RichText1.Text = "Playing."
'  -->        If MMControl1.Mode = 527 Then RichText1.Text = "Recording."
'  -->        If MMControl1.Mode = 528 Then RichText1.Text = "Seeking."
'  -->        If MMControl1.Mode = 529 Then RichText1.Text = "Paused."
'  -->        If MMControl1.Mode = 530 Then RichText1.Text = "Ready."

End Sub

Private Sub SOUND_PLAY_END_TIMER_Timer()

        If MMControl1.Mode = 525 Then
            SOUND_PLAY_END_TIMER.Enabled = False
            MMControl1.Command = "CLOSE"
            MMControl1.Filename = ""
        End If

End Sub


Private Sub SetFullRowSelection(ByVal hWndListView, ByVal bFullRow As Boolean)
   SendMessage hWndListView, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, ByVal CLng(bFullRow)
End Sub

Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column _
 As ColumnHeader = Nothing)

 Dim C As ColumnHeader
 If Column Is Nothing Then
  For Each C In LV.ColumnHeaders
   SendMessage LV.hWnd, LVM_FIRST + 30, C.Index - 1, -1
  Next
 Else
  SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
 End If
 LV.Refresh
End Sub


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


Public Function FolderExists(sFolder As String) As Boolean
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


'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  'IsIDE = False
  'Exit Function
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************


Public Function GetOperatingSystem()
    Dim localHost       As String
    Dim objWMIService   As Variant
    Dim colOperatingSystems As Variant
    Dim objOperatingSystem As Variant
    On Error GoTo Error_Handler
    localHost = "." 'Technically could be run against remote computers, if allowed
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & localHost & "\root\cimv2")
    Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    For Each objOperatingSystem In colOperatingSystems
        GetOperatingSystem = objOperatingSystem.Caption & " " & objOperatingSystem.Version
        GetOperatingSystem = Format(Val(Mid(objOperatingSystem.Version, 1, InStr(objOperatingSystem.Version, "."))), "00")
        Exit Function
    Next
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
Error_Handler:
    MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: getOperatingSystem" & vbCrLf & _
           "Error Description: " & Err.Description, _
           vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function



Sub SET_MENU_PADD_WORK()

Call frmListMenu.SET_MENU_PADD_WORK

End Sub

' Add or subtract one from wait_counter. If
' wait_counter = 0, display the default cursor.
' Otherwise display the hourglass cursor.
Private Sub Hourglass(ByVal start_wait As Boolean)
Static wait_counter As Integer

Dim was_hourglass As Boolean
Dim now_hourglass As Boolean

    ' Record the current cursor state.
    was_hourglass = (wait_counter > 0)

'    ' Update start_wait.
'    If start_wait Then
'        wait_counter = wait_counter + 1
'    Else
'        wait_counter = wait_counter - 1
'    End If
'    If wait_counter < 0 Then wait_counter = 0
'
'    ' See if the cursor's status has changed.
'    ' We only set it if it has. Otherwise
'    ' repeatedly resetting the cursor with the
'    ' same value will make it flicker.
'    now_hourglass = (wait_counter > 0)
'    If now_hourglass <> was_hourglass Then
'        If wait_counter = 0 Then
'            MousePointer = vbDefault
'        Else
'            MousePointer = vbHourglass
'        End If
'    End If
'
'    ' Update the wait counter display.
'    ' Remove this in a real application.
'    lblWaits = wait_counter
End Sub
