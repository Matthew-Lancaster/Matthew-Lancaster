VERSION 5.00
Begin VB.Form FORM_1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Volume2"
   ClientHeight    =   4416
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   DrawStyle       =   5  'Transparent
   Icon            =   "Volume2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4416
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer_TIME_INFO__SEARCH_REPLACE_RESTORE 
      Interval        =   1000
      Left            =   2052
      Top             =   900
   End
   Begin VB.Timer TIMER_ESCAPE_KEY 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1824
      Top             =   2784
   End
   Begin VB.Timer TIMER_HWND_4 
      Interval        =   10
      Left            =   1044
      Top             =   2700
   End
   Begin VB.Timer MOUSE_HOVER 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2844
      Top             =   2436
   End
   Begin VB.Timer Timer_FORM_PROJECT_CHECK_DATE 
      Interval        =   2000
      Left            =   2880
      Top             =   1812
   End
   Begin VB.Timer TIMER_SET_POSITION 
      Interval        =   1
      Left            =   972
      Top             =   2136
   End
   Begin VB.Timer Timer_MOUSE_MOVER 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   924
      Top             =   1668
   End
   Begin VB.Timer Timer_TIMEZONE 
      Interval        =   1000
      Left            =   840
      Top             =   1008
   End
   Begin VB.Label LAB_TIMEZONE 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "TIMEZONE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   744
      TabIndex        =   0
      Top             =   324
      Width           =   1524
   End
End
Attribute VB_Name = "FORM_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================================
'# __ D:\VB6\VB-NT\00_Best_VB_01\TIMEZONE MINI GUI DISPLAY\
'# __
'# __ TIMEZONE MINI GUI DISPLAY.vbp
'# __ TIMEZONE MINI GUI DISPLAY.exe
'# __
'# BY Matthew __ Matt.Lan@Btinternet.com
'# __
'# START AND END TIME
'# Sun 09-Feb-2020 17:38:11
'# Sun 09-Feb-2020 18:40:00 -- 1 HOUR 2 MINUTE
'# __
'====================================================================

' -------------------------------------------------------------------
' DESCRIPTION
' -------------------------------------------------------------------
' SESSION 001
' -------------------------------------------------------------------
' DISPLAY THE TIMEZONE DATE FOR ANOTHER AREA
' THE OPERATION PROJECT TAKEN FROM MY OWN
' -------------------------------------------------------------------
' "D:\VB6\VB-NT\00_Best_VB_01\Winamp - Volume MINI\VolumeBar MINI.vbp"
' -------------------------------------------------------------------
' HAS A BORDERLESS WINDOW
' WITH CLOCK IN THERE -- TIMEZONE SHIFT
' THE LAST THING WAS TO MAKE MOUSE DOWN
' AND MAKE IT RELOCATABLE
' WHICH A CHALLENEG
' AS TIMER REQUIRE AFTER MOUSE DOWN EVENT
' AND THEN GET X Y CORDINATE
' AND APPLY TO MOVE LOCATE THE WINDOW ANYWHERE
' THE LEFT RIGHT Y CORDIDNATE WAS A BIT OFFSET SO MINUS FEW VALUE NUMBER
' AND GET IT EXTACT FLOATER
' -------------------------------------------------------------------
' DOUBLE CLICK TO EXIT
' AND A TOOLTIP FLOAT OVER TALK WHERE IT ARE
' -------------------------------------------------------------------

' -------------------------------------------------------------------
' LOCATION ON-LINE
' -------------------------------------------------------------------

Dim LATCH_TIMEZONE_HWND_BEGIN

Dim FSO

Dim FREE_BYTE_01
Dim DRIVE_CURRENT_STORE_1
Dim TIMER_VAR_STORE_1

Dim OLD_HOUR_NOW
Dim TOOLTIP_NOT_SHOW_AFTER_TIMER

Dim SYSTEM_START_TIME_STRING

Dim HWND_AND_TITLE_OLD

Dim HWND_4_OLD_INT
Dim HWND_BEFORE
Dim MOUSE_DOWN_EVENT_TIME_4 As Date
Dim DEBUG_ME_TOP__ME_LEFT
Dim HWND_4_TRIG

Private Const MAX_LENGTH = &HFF
Private Const WM_GETTEXTLENGTH = &HE
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Dim TITLE_HWND_4_OLD, TITLE_HWND_4, HWND_4_OLD

Dim MOUSE_DOWN_EVENT_TIME

Public cProcesses As New clsCnProc


Dim GS
Dim FREE_BYTE_02

Dim MOD_VALUE_INTERVAL_2

Dim TIMER_TIMEZONE_FIRST_RUN

Dim Explorer_Path_OLD, Explorer_Path

Dim STR_ME_TOP
Dim STR_ME_LEFT

Dim X_VAR
Dim Y_VAR

Dim ENDE
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Public ForeG

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
    
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "Kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
    
Private Declare Function EnumChildWindows Lib "user32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cchar As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cchar As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

 
Public FIND_CHILD_BUTTON

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
    
Private Declare Function GetParent _
        Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent _
        Lib "user32" _
        (ByVal hWndChild As Long, _
         ByVal hWndNewParent As Long) As Long
    
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
    
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
    
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
Private Declare Function FindFirstFile Lib "Kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "Kernel32" (ByVal hFindFile As Long) As Long
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
'Private Type SHITEMID
'    cb As Long
'    abID As Byte
'End Type
'Private Type ITEMIDLIST
'    mkid As SHITEMID
'End Type
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
'Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
'Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowDLL Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function FindWindow2 Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
'Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
'Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
'Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cchar As Long) As Long
'Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
' --------------------------------------------------------------------
'Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
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
   Dim L As Long
   Dim S As String
   L = GetWindowTextLength(hWnd)
   S = Space(L + 1)
   GetWindowText hWnd, S, L + 1
   GetWindowTitle = Left$(S, L)
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



    
    
    
    
    
    
'Function GetUserName() As String
'   Dim UserName As String * 255
'   Call GetUserNameA(UserName, 255)
'   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
'End Function
'Function GetComputerName() As String
'   Dim UserName As String * 255
'   Call GetComputerNameA(UserName, 255)
'   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
'End Function
    
    
Public Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function

Public Function NotAlwaysOnTop(ByVal hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function


Private Sub Form_Load()

    If App.PrevInstance = True Then End
    
    ' If IsIDE = False Then End
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'KILL ITSELF IN __.EXE KILL SOFTLY
    'WHILE ISIDE LEARN
    '---------------------------------
    If IsIDE = True Then
        A = cProcesses.GetEXEID_KILL_ALL_INSTANCE(App.Path + "\" + App.EXEName + ".exe")
        If A = "FOUND SOEMTHING" Then
            Beep
            End
        End If
    End If
    
    TIMER_TIMEZONE_FIRST_RUN = "D:\"
    
    Me.Caption = App.EXEName
    
    On Error Resume Next
        FILE_NAME = App.Path + "\" + App.EXEName + "_DATA_HOLDER_FOR_X_Y_CORDINATE_#NFS_EX_.TXT"
        If Dir(FILE_NAME) <> "" Then
            FR1 = FreeFile
            Open FILE_NAME For Input As #FR1
                Line Input #FR1, STR_ME_TOP
                Line Input #FR1, STR_ME_LEFT
            Close FR1
            
        End If
    On Error GoTo 0
    
    LAB_TIMEZONE.ToolTipText = UCase(App.Path + " ---- \ ---- " + App.EXEName + ".EXE")
    
    ' Call TIMER_SET_POSITION_Timer
    LAB_TIMEZONE.Caption = ""
    LAB_TIMEZONE.Visible = False
    Me.Visible = False
    
    If IsIDE = True Then
        TIMER_ESCAPE_KEY.Enabled = True
    End If
    
    TOOLTIP_NOT_SHOW_AFTER_TIMER = Now + TimeSerial(0, 2, 0)
    
End Sub

Private Sub Form_Resize()
Call TIMER_SET_POSITION_Timer
End Sub



Private Sub TIMER_ESCAPE_KEY_Timer()

    If GetAsyncKeyState(27) < 0 Then End

End Sub

Private Sub Timer_TIME_INFO__SEARCH_REPLACE_RESTORE_Timer()
    If Timer_TIME_INFO__SEARCH_REPLACE_RESTORE.Interval <> 100 Then
        Timer_TIME_INFO__SEARCH_REPLACE_RESTORE.Interval = 100
    End If
    
    ENTER_CAPTION_LAB = LAB_TIMEZONE.Caption
    
    CAPTION_LAB = ENTER_CAPTION_LAB
    If TIMER_VAR_STORE_1 <> "" Then
        TIMER_VAR_STORE_2 = TIMER_VAR_STORE_1
        TIMER_VAR_STORE_1 = UCase(Format(Now + TimeSerial(10, 0, 0), "DDD  HH : MM : SS.000  AM/PM"))
        TIMER_VAR_STORE_1 = Replace(TIMER_VAR_STORE_1, ".000", Format(Timer - Int(Timer), ".000"))
        If InStr(TIMER_VAR_STORE_1, "") Then
            CAPTION_LAB = Replace(CAPTION_LAB, TIMER_VAR_STORE_2, TIMER_VAR_STORE_1)
        End If
    End If
    
    
    If DRIVE_CURRENT_STORE_1 <> "" Then
        DRIVE_CURRENT_STORE_2 = DRIVE_CURRENT_STORE_1
        On Error Resume Next
        Set G = FSO.GetDrive(Mid(FREE_BYTE_01, 1, 1))
        XX_DRIVE_READY = False
        Err.Clear
        XX_DRIVE_READY = G.ISREADY
        If XX_DRIVE_READY = True And Err.Number = 0 Then
            NUKE1 = -1
            NUKE1 = G.freespace
            I4 = 0
            M2_GB = NUKE1 / 1024 ^ 3
            M2_MB = NUKE1 / 1024 ^ 2
            M2_KB = NUKE1 / 1024
            M2_B_ = NUKE1
            If M2_B_ > 0 Then I5 = "B": M2_8 = M2_B_: I4 = Format$(M2_8, "0")
            If M2_KB > 0 Then I5 = "K": M2_8 = M2_KB: I4 = Format$(M2_8, "0")
            If M2_MB > 0 Then I5 = "M": M2_8 = M2_MB: I4 = Format$(M2_8, "0")
            If M2_GB > 0 Then I5 = "G": M2_8 = M2_GB: I4 = Format$(M2_8, "0")

            If I5 = "G" Then I5 = "G"
            I4 = Format$(M2_8, "0")
            L4 = Format$(M2_8, "0.0000000")
            M4 = Format$(M2_8, "0.00")
            I5 = "-" + I5
            i2 = Chr(R1) + " " + I4 + I5 + "   "
            i2 = L4 + I5
            i2 = i2 + "   "   ' Chr(175) + UP UNDERSCORE
            DRIVE_CURRENT = i2
        End If
        DRIVE_CURRENT = Trim(DRIVE_CURRENT)
        
        If InStr(DRIVE_CURRENT_STORE_1, "") Then
            DRIVE_CURRENT_STORE_1 = FREE_BYTE_01 + " __ " + DRIVE_CURRENT
            CAPTION_LAB = Replace(CAPTION_LAB, DRIVE_CURRENT_STORE_2, DRIVE_CURRENT_STORE_1)
        End If
    End If

    If ENTER_CAPTION_LAB <> CAPTION_LAB Then
        LAB_TIMEZONE.Caption = CAPTION_LAB
    End If

End Sub
Private Sub Timer_TIMEZONE_Timer()

    Call SYSTEM_UPTIME


    ' ------------------------------------------------------------------
    ' FORM NOT QUIT PROPER
    ' STILL REMAIN WHEN HWND VALUE GOT NONE
    ' EXAMPLE AHK MENU BUTTON GO REMOVE ALL VISUAL BASIC
    ' 23-OCT-2020 02:36:29 FRI
    ' ------------------------------------------------------------------
    If FindWindow("ThunderFormDC", "TIMEZONE MINI GUI DISPLAY") > 0 Then
        LATCH_TIMEZONE_HWND_BEGIN = True
    End If
    
    If FindWindow("ThunderFormDC", "TIMEZONE MINI GUI DISPLAY") = 0 Then
        If LATCH_TIMEZONE_HWND_BEGIN = True Then
            End
        End If
    End If


   ' Timer_TIMEZONE.Interval = 1000
    
    If TOOLTIP_NOT_SHOW_AFTER_TIMER < Now Then
        LAB_TIMEZONE.ToolTipText = ""
    End If
    
    If Hour(Now) <> OLD_HOUR_NOW Then
        OLD_HOUR_NOW = Hour(Now)
        Call SYSTEM_UPTIME
    End If
    
    

    Dim FS
    Dim M1(10)
    Dim M2(10)
    Dim M3(10)
    
    i = i + 1: M1(i) = "1-ASUS-X5DIJ"
    i = i + 1: M1(i) = "2-ASUS-EEE"
    i = i + 1: M1(i) = "3-LINDA-PC"
    i = i + 1: M1(i) = "4-ASUS-GL522VW"
    i = i + 1: M1(i) = "5-ASUS-P2520LA"
    i = i + 1: M1(i) = "7-ASUS-GL522VW"
    i = i + 1: M1(i) = "8-MSI-GP62M-7RD"
    For r = 1 To i
        M2(r) = Replace(M1(r), "-", "_")
    Next
    i = 0
    i = i + 1: M3(i) = "1X"
    i = i + 1: M3(i) = "2E"
    i = i + 1: M3(i) = "3L"
    i = i + 1: M3(i) = "4G"
    i = i + 1: M3(i) = "5P"
    i = i + 1: M3(i) = "7G"
    i = i + 1: M3(i) = "8M"
    For r = 1 To i
        If GetComputerName = M1(i) Then
            SYMBOL_COMPUTERNAME = M3(i)
        End If
    Next
    
    ' 1X 2E 3L
    EXTRA_DETAIL_NOT_LOW_END_COMPUTER = True
    For r = 1 To 3
        If GetComputerName = M1(i) Then
            Timer_TIMEZONE.Interval = 20000
            ' ------------------------------------------------------------------
            ' EXTRA_DETAIL_NOT_LOW_END_COMPUTER = False
            ' ------------------------------------------------------------------
            ' LESS QUICK LOW END COMPUTER
            ' AND TIMER HAVE SLIP IN FUCTION WHERE SEARCH REPLEACE ONLY THAT ONE
            ' ------------------------------------------------------------------
            
        End If
    Next


    MOD_VALUE_INTERVAL_1 = 10
    MOD_VALUE_INTERVAL_2 = MOD_VALUE_INTERVAL_2 + 1
    MOD_VALUE_INTERVAL_2 = 0
    If MOD_VALUE_INTERVAL_2 Mod MOD_VALUE_INTERVAL_1 = 0 Then
        MOD_VALUE_INTERVAL_2 = 0
        GOOSYNC_PORTABLE = ""
        If GOOSYNC_PORTABLE = "" Then
            If FSO.FileExists("C:\SCRIPTOR DATA\VB_KEEP_RUNNER_IS_D_HDD_GOODSYNC2GO_RUNNER\D_HDD_GOODSYNC2GO_RUNNER_" + GetComputerName + ".TXT") = True Then
                GOOSYNC_PORTABLE = "GS-D-HDD __"
            End If
        End If
        If GOOSYNC_PORTABLE = "" Then
            If FSO.FileExists("C:\SCRIPTOR DATA\VB_KEEP_RUNNER_IS_C_HDD_GOODSYNC2GO_RUNNER\C_HDD_GOODSYNC2GO_RUNNER_" + GetComputerName + ".TXT") = True Then
                GOOSYNC_PORTABLE = "GS-OS-HDD"
            End If
        End If
    
        GS = GOOSYNC_PORTABLE
        'FOCUS
        
        hWndForm = FindWindow("CabinetWClass", vbNullString)
        Explorer_Path = StripNulls(GetWindowTitle(hWndForm))
        Explorer_Path = Replace(Explorer_Path, "This PC", "")
        
        If Mid(Explorer_Path, 2, 1) <> ":" Then
            If Mid(Explorer_Path, 1, 2) <> "\\" Then
                Explorer_Path = ""
            End If
        End If
            
        If Explorer_Path = "" Then
            ' -------------------------------------------------------
            ' IF NONE AT
            ' EXPLORER_PATH
            ' GIVE IT DEFAULT D DRIVE -- VIA TIMER_TIMEZONE_FIRST_RUN
            ' -------------------------------------------------------
            If TIMER_TIMEZONE_FIRST_RUN <> "" Then
                Explorer_Path_OLD = TIMER_TIMEZONE_FIRST_RUN
                TIMER_TIMEZONE_FIRST_RUN = ""
            End If
            If Explorer_Path_OLD <> "" Then
                Explorer_Path = Explorer_Path_OLD
            End If
        End If
        
        Explorer_Path_OLD = Explorer_Path
        TIMER_TIMEZONE_FIRST_RUN = ""
        
        If Explorer_Path <> "" Then
            Set FS = CreateObject("Scripting.FileSystemObject")
            If Mid(Explorer_Path, 1, 2) = "\\" Then
                X2 = InStr(Explorer_Path, "\\")
                If X2 > 0 Then
                    X21 = InStr(X2 + 2, Explorer_Path, "\")
                    X31 = InStr(X21 + 1, Explorer_Path, "\")
                    If X31 = 0 And X21 > 0 Then X31 = Len(Explorer_Path)
                    If X31 > 0 Then
                        Explorer_Path = Mid(Explorer_Path, 1, X31)
                    End If
                End If
            End If
            If Mid(Explorer_Path, 2, 1) = ":" Then
                Explorer_Path = Mid(Explorer_Path, 1, 1) + ":\"
            End If
            DRI = Explorer_Path
            On Error Resume Next
            Dim DRV_ARR(26)
            i = ""
            XXO = Time
            DRIVE_COUNTER_1 = 0
            For R1 = Asc("C") To Asc("Y")
                Set G = FS.GetDrive(Chr(R1)) ' DRI
                XX_DRIVE_READY = False
                Err.Clear
                XX_DRIVE_READY = G.ISREADY
                If XX_DRIVE_READY = True And Err.Number = 0 Then
                    NUKE1 = -1
                    NUKE1 = G.freespace
                    Call CHKDRIVE_SPACE_TO_LOAD_RECYCLEBIN(NUKE1, Chr(R1))
                    DRV_ARR(R1 - Asc("C")) = NUKE1
                    I4 = 0
                    M2_GB = NUKE1 / 1024 ^ 3
                    M2_MB = NUKE1 / 1024 ^ 2
                    M2_KB = NUKE1 / 1024
                    M2_B_ = NUKE1
                    If I4 = "0" Then I5 = "G": M2_8 = M2_GB: I4 = Format$(M2_8, "0")
                    If I4 = "0" Then I5 = "M": M2_8 = M2_MB: I4 = Format$(M2_8, "0")
                    If I4 = "0" Then I5 = "K": M2_8 = M2_KB: I4 = Format$(M2_8, "0")
                    If I4 = "0" Then I5 = "B": M2_8 = M2_B_: I4 = Format$(M2_8, "0")
                    If I5 = "G" Then I5 = "G"
                    I4 = Format$(M2_8, "0")
                    L4 = Format$(M2_8, "0.0000000")
                    M4 = Format$(M2_8, "0.00")
                    I5 = "-" + I5
                    
                    VOL_LAB = G.VOLUMENAME
                    If VOL_LAB = "OS" Then V_LAB = ""
                    If VOL_LAB = "MSI_1TB_DRIVE" Then V_LAB = "_[ M1TB]"
                    If VOL_LAB = "MSI_2TB_DRIVE" Then V_LAB = "_[ M2TB]"
                    If VOL_LAB = "1_SAMSUNG_4TB_D_4" Then V_LAB = "_[ WD4TB]"
                    If VOL_LAB = "MC - HX60V1" Then V_LAB = "_[ MC_V1]"
                    
                    i2 = Chr(R1) + " " + I4 + I5 + V_LAB + "   "
                    
                    If Mid(DRI, 1, 1) = Chr(R1) Then
                        i2 = L4 + I5
                        i2 = i2 + "   "   ' Chr(175) + UP UNDERSCORE
                        DRIVE_CURRENT = i2
                        i2 = V_LAB + Chr(R1) + " " + I4 + I5 + "   "
                        DRIVE_COUNTER_1 = DRIVE_COUNTER_1 + 1
                    Else
                    End If
                    i1 = i1 + i2
                End If
                Set G = Nothing
            Next
            DRIVE_CURRENT = Trim(DRIVE_CURRENT)
            
            On Error GoTo 0
            If Mid(Explorer_Path, 1, 2) = "\\" Then
                DISPLAY_HDD_NET = LCase(Explorer_Path)
                For r = 1 To UBound(M1)
                    DISPLAY_HDD_NET = Replace(DISPLAY_HDD_NET, LCase(M1(r)), M3(r))
                    DISPLAY_HDD_NET = Replace(DISPLAY_HDD_NET, LCase(M2(r)), "")
                    DISPLAY_HDD_NET = Replace(DISPLAY_HDD_NET, "\\", "")
                Next
                DISPLAY_HDD_NET = UCase(Replace(DISPLAY_HDD_NET, "_02_", ""))
                Explorer_Path = DISPLAY_HDD_NET
            End If
            ' ----------------------------------
            ' FUEL GUAGE EMPTY
            ' ----------------------------------
            FREE_BYTE_01 = Trim(Explorer_Path)
            FREE_BYTE_02 = Trim(i1)
            DRIVE_COUNTER_2 = 0
            For r = 1 To Len(FREE_BYTE_02)
                ' ---------------------------------------------------
                ' COUNTER WHERE SPLIT BETWEEN EACH DRIVE TEXT 3 SPACE
                ' ---------------------------------------------------
                If Mid(FREE_BYTE_02, r, 3) = "   " Then
                    DRIVE_COUNTER_2 = DRIVE_COUNTER_2 + 1
                End If
            Next
            DRIVE_COUNTER_4 = Int(DRIVE_COUNTER_2 / 2)
            If DRIVE_COUNTER_4 >= 8 Then
                DRIVE_COUNTER_2 = 0
                For r = 1 To Len(FREE_BYTE_02)
                    ' ---------------------------------------------------
                    ' COUNTER WHERE SPLIT BETWEEN EACH DRIVE TEXT 3 SPACE
                    ' ---------------------------------------------------
                    If Mid(FREE_BYTE_02, r, 3) = "   " Then
                        DRIVE_COUNTER_2 = DRIVE_COUNTER_2 + 1
                        If DRIVE_COUNTER_2 > DRIVE_COUNTER_4 Then
                            FREE_BYTE_02_POS_FINDER = r
                            Exit For
                        End If
                    End If
                Next
                FREE_BYTE_04 = ""
                For r = 1 To FREE_BYTE_02_POS_FINDER
                    FREE_BYTE_04 = FREE_BYTE_04 + Mid(FREE_BYTE_02, r, 1)
                Next
                ' ----------------------------------
                FREE_BYTE_04 = FREE_BYTE_04 + vbCrLf
                ' ----------------------------------
                For r = FREE_BYTE_02_POS_FINDER + 2 To Len(FREE_BYTE_02)
                    FREE_BYTE_04 = FREE_BYTE_04 + Mid(FREE_BYTE_02, r, 1)
                Next
                FREE_BYTE_02 = FREE_BYTE_04
            End If
        End If
    
    End If

    VAR_X = Len(FREE_BYTE_01) + Len(GS)
    If LAB_TIMEZONE.Caption = "" And VAR_X > 0 Then
        LAB_TIMEZONE.Visible = True
        Me.Visible = True
    End If

    'CPU MPC
    M_NAME = GetComputerName
    If M_NAME = "1-ASUS-X5DIJ" Then M_NAME_2 = "1X "
    If M_NAME = "2-ASUS-EEE" Then M_NAME_2 = "2E "
    If M_NAME = "3-LINDA-PC" Then M_NAME_2 = "3L "
    If M_NAME = "4-ASUS-GL522VW" Then M_NAME_2 = "4G "
    If M_NAME = "5-ASUS-P2520LA" Then M_NAME_2 = "5P "
    If M_NAME = "7-ASUS-GL522VW" Then M_NAME_2 = "7G "
    If M_NAME = "8-MSI-GP62M-7RD" Then M_NAME_2 = "8M "
    M_NAME = M_NAME_2
    
    If LAB_TIMEZONE.Alignment <> 0 Then
        LAB_TIMEZONE.Alignment = 0 ' LEFT -- 2 CENTER
    End If
    
    DRIVE_CURRENT_STORE_1 = FREE_BYTE_01 + " __ " + DRIVE_CURRENT
    i = ""
    i = i + M_NAME + " __ " + DRIVE_CURRENT_STORE_1 + " __ " + GS
    TIMER_VAR_STORE_1 = UCase(Format(Now + TimeSerial(10, 0, 0), "DDD  HH : MM : SS.000  AM/PM"))
    TIMER_VAR_STORE_1 = Replace(TIMER_VAR_STORE_1, ".000", Format(Timer - Int(Timer), ".000"))
    i = i + TIMER_VAR_STORE_1
    i = i + " __ SYS START __ " + SYSTEM_START_TIME_STRING
    i = i + DEBUG_ME_TOP__ME_LEFT + vbCr + FREE_BYTE_02
    
    CAPTION_LAB = i
    LAB_TIMEZONE.Caption = CAPTION_LAB
    
'    LAB_TIMEZONE.AutoSize = False
'    LAB_TIMEZONE.Refresh
'    LAB_TIMEZONE.AutoSize = True

    Me.Width = LAB_TIMEZONE.Width + 200
    Me.Height = LAB_TIMEZONE.Height

End Sub


Sub CHKDRIVE_SPACE_TO_LOAD_RECYCLEBIN(DRIVE_SPACE, DRIVE_LETTER)

If DRIVE_SPACE <> 0 Then Exit Sub

If FindWindow(vbNullString, "# Recycle Bin") = 0 Then
    Shell "EXPLORER shell:RecycleBinFolder", vbMaximizedFocus
End If

' If GET_RECYCLEBIN_SIZE > 0 Then
'    Shell "EXPLORER shell:RecycleBinFolder", vbNormalFocus
' End If

End Sub



Sub SYSTEM_UPTIME()

Load MDIProcServ

If MDIProcServ.TT_1VDT > 0 Or MDIProcServ.TT_2VDT > 0 Then
    Dim XI_STRING
    Dim YEAR_DIFF, MONTH_DIFF_MOD, YEAR_OF_MONTH_MOD_12
    DAY_CODE = Format(Int((DateDiff("h", MDIProcServ.TT_1VDT, Now) / 24)), "0")
    DAY_CODE_SHUTDOWN = Val(DAY_CODE)
    XI_STRING = ""
    'XI_STRING = XI_STRING + Str(DateDiff("h", MDIProcServ.TT_1VDT, Now)) + "HR -- "
    XI_STRING = XI_STRING + DAY_CODE + " DAY"
    XI_STRING = XI_STRING + " " + Format((DateDiff("h", MDIProcServ.TT_1VDT, Now) Mod 24), "0") + " HOUR"
    SYSTEM_START_TIME_STRING = XI_STRING
    ' Form1.Text_SYSTEM_START_TIME_02.Text = XI_STRING
End If

End Sub


Private Sub Timer_FORM_PROJECT_CHECK_DATE_Timer()

Call Project_Check_Date.VB_PROJECT_CHECKDATE("FORM LOAD")

Timer_FORM_PROJECT_CHECK_DATE.Interval = 60000

End Sub


Sub TIMER_SET_POSITION_Timer()
    
    On Error Resume Next
    ' Me.Show
    
    Me.WindowState = vbNormal
    If Err.Number > 0 Then Exit Sub
    
    
    If STR_ME_TOP = "" Then
        Me.Top = 150
        Me.Left = 680
    End If
    If STR_ME_TOP <> "" Then
        Me.Top = Val(STR_ME_TOP)
    End If
    If STR_ME_LEFT <> "" Then
        Me.Left = Val(STR_ME_LEFT)
    End If
    
    LAB_TIMEZONE.Top = 0
    LAB_TIMEZONE.Left = 0
'    LAB_TIMEZONE.Width = 100
    LAB_TIMEZONE.FontName = "Arial"
    FONT_SIZE = 15
    If GetComputerName = "1-ASUS-X5DIJ" Then FONT_SIZE = 13
    If GetComputerName = "2-ASUS-EEE" Then FONT_SIZE = 13
    LAB_TIMEZONE.FontSize = FONT_SIZE
    LAB_TIMEZONE.FontBold = True
    Call Timer_TIMEZONE_Timer
    ' LAB_TIMEZONE.AutoSize = False
    ' LAB_TIMEZONE.Refresh
    ' LAB_TIMEZONE.AutoSize = True
    LAB_TIMEZONE.Left = 100
    
    Me.Width = LAB_TIMEZONE.Width + 200
    Me.Height = LAB_TIMEZONE.Height
    
    If Err.Number > 0 Then Exit Sub
    AlwaysOnTop (Me.hWnd)

    If Err.Number = 0 Then
    TIMER_SET_POSITION.Enabled = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    FR1 = FreeFile
    Open App.Path + "\" + App.EXEName + "_DATA_HOLDER_FOR_X_Y_CORDINATE_#NFS_EX_.TXT" For Output As #FR1
        Print #FR1, Str(Me.Top)
        Print #FR1, Str(Me.Left)
    Close FR1
End Sub

Private Sub LAB_TIMEZONE_DblClick()
    Unload Me
    End
End Sub
Private Sub LAB_TIMEZONE_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Timer_MOUSE_MOVER.Enabled = True
    X_VAR = x
    Y_VAR = y
    
    If MOUSE_DOWN_EVENT_ACTION = "HIGH" Then
        MOUSE_DOWN_EVENT_TIME = 0
    End If
    If MOUSE_DOWN_EVENT_TIME = 0 Then
        MOUSE_DOWN_EVENT_TIME = Now + TimeSerial(0, 0, 4)
        MOUSE_DOWN_EVENT_TIME_4 = Now + TimeSerial(0, 0, 10)

        MOUSE_DOWN_EVENT_ACTION = "DOWN"
        If TOOLTIP_NOT_SHOW_AFTER_TIMER > Now Then
            LAB_TIMEZONE.ToolTipText = "HOLD MOUSE DOWN 3 SEC TO QUIT ---- DOUBLE CLICK END WHOLE APP"
        End If
        MOUSE_HOVER.Enabled = True
    End If
End Sub
Sub MOUSE_HOVER_TIMER()
    If MOUSE_DOWN_EVENT_ACTION = "DOWN" Then
        If MOUSE_DOWN_EVENT_TIME <= Now And MOUSE_DOWN_EVENT_TIME > 0 Then
            Unload Me
        End If
    End If
    If MOUSE_DOWN_EVENT_TIME <= Now And MOUSE_DOWN_EVENT_ACTION = "HIGH" Then
        MOUSE_DOWN_EVENT_TIME = 0
        If TOOLTIP_NOT_SHOW_AFTER_TIMER > Now Then
            LAB_TIMEZONE.ToolTipText = UCase(App.Path + " ---- \ ---- " + App.EXEName + ".EXE")
        End If
    End If
    If MOUSE_DOWN_EVENT_TIME_4 < Now And DEBUG_ME_TOP__ME_LEFT > "" Then
        DEBUG_ME_TOP__ME_LEFT = ""
    End If
End Sub
Private Sub LAB_TIMEZONE_MouseUP(Button As Integer, Shift As Integer, x As Single, y As Single)
    Timer_MOUSE_MOVER.Enabled = False
    R_BUTTON = 2
    If Button = R_BUTTON Then
        Call Mnu_VB_ME_Click
    End If
    MOUSE_DOWN_EVENT_ACTION = "HIGH"
    
    AlwaysOnTop (Me.hWnd)

End Sub

Private Sub Timer_MOUSE_MOVER_Timer()
    Dim tPA As POINTAPI
    ' Get cursor cordinates
    GetCursorPos tPA
    ' Set label caption to cursor cordinates
    ' lblCordi.Caption = "X: " & tPA.x & "  Y: " & tPA.y

    Me.Top = (tPA.y) * Screen.TwipsPerPixelY - (Y_VAR - 0)
    Me.Left = (tPA.x) * Screen.TwipsPerPixelX - (X_VAR + 98)
    STR_ME_TOP = Trim(Str(Me.Top))
    STR_ME_LEFT = Trim(Str(Me.Left))
    'HWND_4 = FindWindow("Chrome_WidgetWin_1", vbNullString)
    If HWND_4 > 0 Then
        TITLE_HWND_4 = GetTitle(HWND_4)
        CLASS_HWND_4 = GetClass(HWND_4)
        If InStr(CLASS_HWND_4, "Chrome_WidgetWin_1") > 0 Or MOUSE_DOWN_EVENT_TIME_4 > Now Then
            DEBUG_ME_TOP__ME_LEFT = Str(Me.Top) + " " + Str(Me.Left)
            MOUSE_DOWN_EVENT_TIME_4 = Now + TimeSerial(0, 0, 10)
        End If
    End If

    ' LAB_TIMEZONE
    If MOUSE_DOWN_EVENT_TIME > 0 Then
        If MOUSE_DOWN_EVENT_ACTION = "DOWN" Then
            MOUSE_DOWN_EVENT_TIME = Now + TimeSerial(0, 0, 4)

        End If
    End If
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Call Mnu_VB_ME_Click
    End If
    
    ENDE = ENDE + 1
    If ENDE > 2 Then End
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
    'End
End Sub

Private Sub MNU_CAMERA_VIDEO_FOLDER_4G_Click()
    Me.WindowState = vbMinimized
    Shell "EXPLORER \\4-asus-gl522vw\4_asus_gl522vw_10_1_samsung_1tb\DSC\2015+SONY_MP4\2020 CyberShot HX60V __ MP4", vbMaximizedFocus
End Sub



'-------------------------------------------
'-------------------------------------------
'---------------------------------------------------
'Public Type SHITEMID
'    cb As Long
'    abID As Byte
'End Type
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

Private Sub TIMER_HWND_4_Timer()
    Dim HWND_AND_TITLE
    HWND_4 = GetForegroundWindow
    If HWND_4 = 0 Then Exit Sub
    If HWND_4 = Me.hWnd Then Exit Sub
    
    TITLE_HWND_4 = GetTitle(HWND_4)
    HWND_AND_TITLE = Format(HWND_4, " 00000000000 ") + TITLE_HWND_4
    
    ' If InStr(HWND_4_OLD, Format(HWND_4, "00000000000")) = 0 Then
    If InStr(HWND_AND_TITLE_OLD, HWND_AND_TITLE) = 0 Then
        ' HWND_4_TRIG = False
        ' TITLE_HWND_4_OLD = ""
        ' HWND_AND_TITLE = ""
        
        ' HWND_BEFORE = HWND_4
        ' If GetParent_2(HWND_4) = Me.hWnd Then Exit Sub
        If InStr(UCase(HWND_AND_TITLE), "YOUR NOTIFICATIONS") > 0 Then
            Call SET_POS_FACEBOOK_HIDE_TOP_NOTIFY
        End If
    End If
    HWND_AND_TITLE_OLD = HWND_AND_TITLE
    
    
    Exit Sub
    
    TITLE_HWND_4 = GetTitle(HWND_4)
    HWND_AND_TITLE = Format(HWND_4, " 00000000000 ") + TITLE_HWND_4
    
    ' If HWND_4 = Me.hWnd Then Exit Sub
        
    ' If HWND_4 = HWND_4_OLD_INT Then Exit Sub
    HWND_4_OLD_INT = HWND_4
    HWND_4_OLD = HWND_4_OLD + Format(HWND_4, " 00000000000")
    If Len(HWND_4_OLD) >= 20 Then
        CHOP = Len(HWND_4_OLD) - 20
        If CHOP < 1 Then CHOP = 1
        HWND_4_OLD = Mid(HWND_4_OLD, CHOP + 1)
    End If
    
    ' HWND_4 = FindWindow("Chrome_WidgetWin_1", vbNullString)
    ' HWND_4 = GetForegroundWindow
    If HWND_4 > 0 Then
        TITLE_HWND_4 = GetTitle(HWND_4)
        CLASS_HWND_4 = GetClass(HWND_4)
    End If
    If InStr(UCase(TITLE_HWND_4), "YOUR NOTIFICATIONS") > 0 Then
        Call SET_POS_FACEBOOK_HIDE_TOP_NOTIFY
    End If
    
'    If InStr(UCase(TITLE_HWND_4), "YOUR NOTIFICATIONS") = 0 Then
'        If InStr(CLASS_HWND_4, "ThunderRT6FormDC") = 0 Then
'        If InStr(CLASS_HWND_4, "ThunderFormDC") = 0 Then
'            TITLE_HWND_4_OLD = ""
'            HWND_4_TRIG = False
'        End If
'        End If
'    End If
'    End If
'    End If
    
End Sub

Function GetParent_2(ByVal Handle As Long) As String
   Dim i As Long
   Dim j As Long, TTx1 As String
   i = Handle
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
      i = j
   '
   ' TTx1 = GetWindowTitle(i)
   GetParent_2 = i

End Function



Sub SET_POS_FACEBOOK_HIDE_TOP_NOTIFY()
    
    M_NAME = GetComputerName

'    If M_NAME = "2-ASUS-EEE" Then M_NAME_2 = "2E "
'    If M_NAME = "3-LINDA-PC" Then M_NAME_2 = "3L "
'    If M_NAME = "5-ASUS-P2520LA" Then tPA.x = 100: tPA.y = 100: X_VAR = 100: Y_VAR = 100
    X_VAR = 0
    If M_NAME = "1-ASUS-X5DIJ" Then X_VAR = 2595: Y_VAR = 3262
    If M_NAME = "4-ASUS-GL522VW" Then X_VAR = 3324: Y_VAR = 5518
    If M_NAME = "7-ASUS-GL522VW" Then X_VAR = 3324: Y_VAR = 5518
    If M_NAME = "8-MSI-GP62M-7RD" Then X_VAR = 3324: Y_VAR = 5518
    
    If X_VAR = 0 Then Exit Sub
    
    Me.Top = (X_VAR)
    Me.Left = (Y_VAR)

    ' SEARCH
    ' DEBUG.PRINT STR(ME.TOP) + "  " + STR(ME.LEFT)

End Sub

Private Sub Timer2_Timer()

    Timer2.Enabled = False
    
    ' qq = GetForegroundWindow
    ' If qq = ForeG Then Exit Sub
    ' ForeG = qq
    
    '
    'If FindWindow(vbNullString, "Extreme Lock Build: 2011") > 0 Then
    ''    Exit Sub
    '    Me.Top = Screen.Height - (Me.Height) - 10
    '    Me.Left = 10
    '
    'Else
    '    Me.Top = MeTop
    '    Me.Left = 680
    'End If
    AlwaysOnTop (Me.hWnd)
    'NotAlwaysOnTop (ATidalX.hWnd)

End Sub

Private Sub Timer3_Timer()
    Timer3.Enabled = False

    ENDE = 0
End Sub



Private Sub Mnu_UrlLoad_Click()
    If Mid$(LastURL, 2, 2) = ":\" Then
        Shell "Explorer.exe /e," + LastURL, vbNormalFocus
    Else
        vLaunch (LastURL)
    End If
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
Private Sub MNU_ADMINISTRATOR_Click()
    'Call MNU_ADMINISTRATOR_Click
    ' Beep
    MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR MODE ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "MODE ON", "MODE OFF")
    End If
End Sub



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


