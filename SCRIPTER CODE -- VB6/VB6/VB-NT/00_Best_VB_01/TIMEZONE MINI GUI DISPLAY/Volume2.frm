VERSION 5.00
Begin VB.Form FORM_1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
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
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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



Public cProcesses As New clsCnProc


Dim GS
Dim FREE_BYTE

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
    
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Type POINTAPI
        x As Long
        y As Long
End Type
    
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "Kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
    
Private Declare Function EnumChildWindows Lib "user32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cchar As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cchar As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'Const for Dialog (Form) Data
Const GWL_WNDPROC = -4
Const GWL_HINSTANCE = -6
Const GWL_HWNDPARENT = -8
Const GWL_STYLE = -16
Const GWL_EXSTYLE = -20
Const GWL_USERDATA = -21
Const GWL_ID = -12
 
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
    
    
Public Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function

Public Function NotAlwaysOnTop(ByVal hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, flags
End Function


Private Sub Form_Load()

    If App.PrevInstance = True Then End
    
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
End Sub

Private Sub Form_Resize()
Call TIMER_SET_POSITION_Timer
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
End Sub

Private Sub LAB_TIMEZONE_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Timer_MOUSE_MOVER.Enabled = True
    X_VAR = x
    Y_VAR = y
End Sub
Private Sub LAB_TIMEZONE_MouseUP(Button As Integer, Shift As Integer, x As Single, y As Single)
    Timer_MOUSE_MOVER.Enabled = False
    R_BUTTON = 2
    If Button = R_BUTTON Then
        Call MNU_VB_ME_Click
    End If
End Sub
Private Sub Timer_MOUSE_MOVER_Timer()
    Dim tPA As POINTAPI
    ' Get cursor cordinates
    GetCursorPos tPA
    ' Set label caption to cursor cordinates
    ' lblCordi.Caption = "X: " & tPA.x & "  Y: " & tPA.y

    Me.Left = (tPA.x) * Screen.TwipsPerPixelX - (X_VAR + 98)
    Me.Top = (tPA.y) * Screen.TwipsPerPixelY - (Y_VAR - 0)
    STR_ME_TOP = Trim(Str(Me.Top))
    STR_ME_LEFT = Trim(Str(Me.Left))
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
    End If
    
    ENDE = ENDE + 1
    If ENDE > 2 Then End
End Sub

Private Sub Timer_TIMEZONE_Timer()

   ' Timer_TIMEZONE.Interval = 1000

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


    MOD_VALUE_INTERVAL_1 = 10
    MOD_VALUE_INTERVAL_2 = MOD_VALUE_INTERVAL_2 + 1
    If MOD_VALUE_INTERVAL_2 Mod MOD_VALUE_INTERVAL_1 = 0 Then
        MOD_VALUE_INTERVAL_2 = 0
        GOOSYNC_PORTABLE = ""
        If Dir("C:\SCRIPTOR DATA\VB_KEEP_RUNNER_IS_D_HDD_GOODSYNC2GO_RUNNER\D_HDD_GOODSYNC2GO_RUNNER_" + GetComputerName + ".TXT") <> "" Then
            GOOSYNC_PORTABLE = "Gs-D-Hdd __ "
        End If
        If GOOSYNC_PORTABLE = "" Then
            If Dir("C:\SCRIPTOR DATA\VB_KEEP_RUNNER_IS_C_HDD_GOODSYNC2GO_RUNNER\C_HDD_GOODSYNC2GO_RUNNER_" + GetComputerName + ".TXT") <> "" Then
                GOOSYNC_PORTABLE = "Gs-c-Hdd __ "
            End If
        End If
    
        GS = GOOSYNC_PORTABLE
        
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
            Set G = FS.GetDrive(DRI)
            If G.ISREADY = False Then Exit Sub
            NUKE1 = G.freespace
            On Error GoTo 0
            'If nuke1 < MHD1 And nuke1 <> MHD1 Then E$ = "ã" Else E$ = "ä"
        '    If nuke1 < MHD1 Then E$ = "ä"
        '    If nuke1 > MHD1 And MHD1 > 0 Then E$ = "ã"
        '    If nuke1 = MHD1 Then E$ = Label13.Caption
        '    If MHD1 = 0 Then E$ = Label13.Caption
            If NUKE1 > 0 Then
            If Mid(Explorer_Path, 1, 2) = "\\" Then
                DISPLAY_HDD_NET = LCase(Explorer_Path)
                For r = 1 To i
                    DISPLAY_HDD_NET = Replace(DISPLAY_HDD_NET, LCase(M1(r)), M3(r))
                    DISPLAY_HDD_NET = Replace(DISPLAY_HDD_NET, LCase(M2(r)), "")
                    DISPLAY_HDD_NET = Replace(DISPLAY_HDD_NET, "\\", "")
                Next
                DISPLAY_HDD_NET = UCase(Replace(DISPLAY_HDD_NET, "_02_", ""))
                ' DISPLAY_HDD_NET = UCase(Replace(DISPLAY_HDD_NET, "DRIVE", "HDD"))
                
                Explorer_Path = DISPLAY_HDD_NET
            End If
                FREE_BYTE = Explorer_Path + " " + Format$(NUKE1 / 1024 ^ 3, "0.0000") + " Gb __ "
            End If
        End If
    
    End If

    VAR_X = Len(FREE_BYTE) + Len(GS)
    If LAB_TIMEZONE.Caption = "" And VAR_X > 0 Then
        LAB_TIMEZONE.Visible = True
        Me.Visible = True
    End If
    LAB_TIMEZONE.Caption = FREE_BYTE + GS + Format(Now + TimeSerial(10, 0, 0), "DDD  HH : MM : SS  AM/PM")

    'LAB_TIMEZONE.AutoSize = False
    'LAB_TIMEZONE.Refresh
    'LAB_TIMEZONE.AutoSize = True
    
    Me.Width = LAB_TIMEZONE.Width + 200
    Me.Height = LAB_TIMEZONE.Height

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




Private Sub MNU_VB_ME_Click()
    
    Dim CODER_VBP_FILE_NAME_2
    Dim VB_1, VB_2, VB_3
    VB_1 = "C:\PROGRAM FILES\Microsoft Visual Studio\VB98\VB6.EXE"
    VB_2 = "C:\PROGRAM FILES (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
    If Dir(VB_1) <> "" Then VB_3 = VB_1
    If Dir(VB_2) <> "" Then VB_3 = VB_2
    '------------------------------------------------------
    'ADD MICROSFOT SCRIPTING RUNTIME FOR HERE IN REFERENCES
    '------------------------------------------------------
    Dim objShell
    Set objShell = CreateObject("Wscript.Shell")
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
    objShell.Run """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
    Set objShell = Nothing
    
    EXIT_TRUE = True
    Unload Me
    Exit Sub
    
    
    Beep
    
'    i = MsgBox(GetSpecialfolder(38))
    
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
    End
End Sub
Private Sub MNU_VB_FOLDER_Click()
    Beep
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
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
'Private Type ITEMIDLIST
'    mkid As SHITEMID
'End Type
'Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Function GetSpecialfolder(ByVal CSIDL As Long) As String
    Dim r As Long
    Dim IDL As ITEMIDLIST
    r = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If r = NOERROR Then
        Path$ = Space$(512)
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function

'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".vbp"
    i = CODER_VBP_FILE_NAME_2
    CODER_EXE_FILE_NAME_1 = i
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
    If Dir(FileSpec) = "" Then FILE_SPEC_GO = CODER_VBP_FILE_NAME_2
    
    If Dir(FILE_SPEC_GO) = "" Then MsgBox "ERROR NOT EXISTOR _ " + FILE_SPEC_GO
    
    CODER_VBP_FILE_NAME_2 = FILE_SPEC_GO
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


