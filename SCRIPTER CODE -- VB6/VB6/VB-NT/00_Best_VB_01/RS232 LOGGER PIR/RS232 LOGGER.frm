VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Begin VB.Form DIALER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RS232_LOGGER"
   ClientHeight    =   3228
   ClientLeft      =   10056
   ClientTop       =   300
   ClientWidth     =   11352
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3228
   ScaleWidth      =   11352
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   WindowState     =   1  'Minimized
   Begin VB.Timer TIMER_FRONT_DOOR 
      Interval        =   1000
      Left            =   2604
      Top             =   708
   End
   Begin MSCommLib.MSComm MSComm4 
      Left            =   1764
      Top             =   204
      _ExtentX        =   974
      _ExtentY        =   974
      _Version        =   393216
      DTREnable       =   0   'False
   End
   Begin VB.Timer Timer_PIR 
      Interval        =   1000
      Left            =   2268
      Top             =   684
   End
   Begin VB.PictureBox RichTextBox1 
      Height          =   735
      Left            =   4944
      ScaleHeight     =   684
      ScaleWidth      =   1404
      TabIndex        =   4
      Top             =   180
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1668
      Left            =   444
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   900
      Width           =   1428
   End
   Begin VB.Timer TIMER_1 
      Interval        =   1000
      Left            =   1860
      Top             =   720
   End
   Begin VB.Timer TimerComm4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3516
      Top             =   1044
   End
   Begin VB.PictureBox MMControl1 
      Height          =   1170
      Left            =   4932
      ScaleHeight     =   1128
      ScaleWidth      =   2160
      TabIndex        =   2
      Top             =   1092
      Visible         =   0   'False
      Width           =   2208
   End
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   5880
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List2 
      Height          =   1008
      Left            =   1080
      TabIndex        =   0
      Top             =   5400
      Width           =   4695
   End
   Begin MSCommLib.MSComm MSComm3 
      Left            =   1164
      Top             =   192
      _ExtentX        =   995
      _ExtentY        =   995
      _Version        =   393216
      DTREnable       =   0   'False
   End
End
Attribute VB_Name = "DIALER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DOOR_OPEN_HAPPEN
Dim FOLDER_NAME, FILE_NAME_4
Dim FILE_NAME, X1, X2, A_NOW
Dim PROGRAM_LOAD
Public cProcesses As New clsCnProc
Dim OLD_VAR_DSR_3, VAR_DSR_3
Dim OLD_VAR_DSR_4, VAR_DSR_4

Public I_N_TAIL
Public I_N_NOTEPAD
Public I_N_AUTOHOTKEY


Private Const MONITOR_ON = -1&
Private Const MONITOR_OFF = 2&
Private Const SC_MONITORPOWER = &HF170&
Private Const WM_SYSCOMMAND = &H112

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Const WM_CLOSE = &H10
Private Const WM_USER = &H400
Private Const WM_COMMAND = &H111
Private Const GW_HWNDNEXT = 2


Private Sub Form_Load()

If App.PrevInstance = True Then End

'KILL ITSELF IN __.EXE KILL SOFTLY
'WHILE ISIDE LEARN
'---------------------------------
Dim VAR, PID As Long
If IsIDE = True Then
    PID = -1
    VAR = cProcesses.GetEXEID(PID, App.Path + "\" + App.EXEName + ".exe")
    If PID <> -1 Then
        'Call Process_HIGH_PRIORITY_CLASS(PID)
        VAR = cProcesses.Process_Kill(PID)
        Beep
        End
    End If
End If

I_N_TAIL = "C:\PStart\# NOT INSTALL REQUIRED\Tail\Tail.exe"
If Dir(I_N_TAIL) = "" Then
    ' MsgBox "THE EXE FILE __ NOT EXIST" + vbCrLf + vbCrLf + I_N_TAIL
    ' Beep
    ' Exit Sub
End If


PROGRAM_LOAD = True

OLD_VAR_DSR_3 = -10
OLD_VAR_DSR_4 = -10

Call TIMER_1_TIMER

'Me.Visible = True


End Sub


Sub TIMER_1_TIMER()

On Error Resume Next
For R = 1 To 16
    Err.Clear
    Me.MSComm3.CommPort = R
    'Me.MSComm3.PortOpen = False
    'DoEvents
    Me.MSComm3.Settings = "1200,N,8,1"
    Me.MSComm3.PortOpen = True
    Me.MSComm3.DTREnable = True

    
'    If Err.Number <> 8002 Then
'        Exit For
'    End If
    If Me.MSComm3.PortOpen = True Then
        VAR_DSR_3 = Me.MSComm3.DSRHolding
        Exit For
    End If
    
Next

' NEXT PORT IS MY FRONT DOOR OPEN LOGGER
' 1ST PORT DETECTED WILL BE FOR PIR 2ND DOOR
For R = Me.MSComm3.CommPort To 16
    Err.Clear
    Me.MSComm4.CommPort = R
    Me.MSComm4.PortOpen = False
    DoEvents
    Me.MSComm4.Settings = "1200,N,8,1"
    Me.MSComm4.PortOpen = True
    Me.MSComm4.DTREnable = True
    
    If Me.MSComm4.PortOpen = True Then
        VAR_DSR_4 = Me.MSComm4.DSRHolding
        Exit For
    End If
Next

' --------------------------------------
' NONE COMM PORT ALL 16 TESTER
' LEAVE HIGH TO KEEP LIGHT FOR SCREEN ON
' --------------------------------------
If R = 17 Then VAR_DSR_3 = 0
If Me.MSComm3.PortOpen = False Then
    VAR_DSR_3 = True
End If


End Sub


Private Sub Timer_PIR_Timer()
        
Dim FSO
On Error Resume Next
If Me.MSComm3.PortOpen = False Then Exit Sub
VAR_DSR_3 = Me.MSComm3.DSRHolding
'Debug.Print Time$ + " " + Str(VAR_DSR_3)
If Err.Number > 0 Or Err.Number = 8002 Then
    TIMER_1.Enabled = True
    VAR_DSR_3 = True
End If

If OLD_VAR_DSR_3 = VAR_DSR_3 Then Exit Sub

OLD_VAR_DSR_3 = VAR_DSR_3

Set FSO = CreateObject("Scripting.FileSystemObject")


' MsgBox Str(R) + " -- " + Str(VAR_DSR_3)
' Debug.Print Str(R) + " -- " + Str(VAR_DSR_3)

Dim AR(3)
AR(1) = "\\1-asus-x5dij\1_asus_x5dij_01_c_drive"
AR(2) = "\\2-asus-eee\2_asus_eee_01_c_drive"
AR(3) = "\\4-asus-gl522vw\4_asus_gl522vw_01_c_drive"

On Error Resume Next
If VAR_DSR_3 = True Then
    For R = 1 To UBound(AR)
        FOLDER_NAME = AR(R) + "\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY"
        FILE_NAME = FOLDER_NAME + "\Autokey -- 14-Brightness With Dimmer.txt"
        If FSO.FILEExists(FILE_NAME) = False Then
            If FSO.FOLDERExists(FOLDER_NAME) = False Then
                RESULT = CreateFolderTree(FOLDER_NAME)
            End If
            FR1 = FreeFile
            Open FILE_NAME For Output As #FR1
            Close #FR1
        End If
    Next
Else
    For R = 1 To UBound(AR)
        FILE_NAME = AR(R) + "\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-Brightness With Dimmer.txt"
        Kill FILE_NAME
    Next
End If

End Sub


Sub TIMER_FRONT_DOOR_TIMER()

Dim STRING_VAR As String

If Me.MSComm4.PortOpen = False Then Exit Sub
VAR_DSR_4 = Me.MSComm4.DSRHolding
Debug.Print Time$ + " " + Str(VAR_DSR_3)
On Error Resume Next
If Me.MSComm4.PortOpen = False Then
    VAR_DSR_4 = False
    TIMER_1.Enabled = True
End If
If Err.Number > 0 Or Err.Number = 8002 Then
    TIMER_1.Enabled = True
    'VAR_DSR_4 = 1
End If

If OLD_VAR_DSR_4 = VAR_DSR_4 Then Exit Sub

OLD_VAR_DSR_4 = VAR_DSR_4

Set FSO = CreateObject("Scripting.FileSystemObject")


' MsgBox Str(R) + " -- " + Str(VAR_DSR_3)
' Debug.Print Str(R) + " -- " + Str(VAR_DSR_3)

Dim AR(1)
'AR(1) = "\\1-asus-x5dij\1_asus_x5dij_01_c_drive"
'AR(2) = "\\2-asus-eee\2_asus_eee_01_c_drive"
AR(1) = "\\4-asus-gl522vw\4_asus_gl522vw_01_c_drive"
AR(1) = "C:"

FILE_NAME_2 = "RS232 FRONT DOOR.txt"
FILE_NAME_8 = "RS232 FRONT DOOR OPEN.txt"
FILE_NAME_9 = "RS232 FRONT DOOR CLOSE.txt"
FILE_NAME_4 = "RS232 FRONT DOOR LOGGER.txt"
PATH_2 = "VB6\VB-NT\00_Best_VB_01\Tidal_Info"



On Error Resume Next
If VAR_DSR_4 = True Then
    For R = 1 To UBound(AR)
        DOOR_OPEN_HAPPEN = True
        FOLDER_NAME = AR(1) + "\SCRIPTOR DATA\" + PATH_2
        FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_2
        If FSO.FILEExists(FILE_NAME) = False Then
            If FSO.FOLDERExists(FOLDER_NAME) = False Then
                RESULT = CreateFolderTree(FOLDER_NAME)
            End If
            FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_2
            FR1 = FreeFile
            Open FILE_NAME For Output As #FR1
            Close #FR1
            FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_8
            FR1 = FreeFile
            Open FILE_NAME For Output As #FR1
            Close #FR1
            FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_4
            A_NOW = Now
            If PROGRAM_LOAD = True Then
                Call CHECK_ARCHIVE_LOGGER
                PROGRAM_LOAD = False
                If X1 > X2 Then
                    Call WRITE_LOGGER_BEGIN
                    Call WRITE_LOGGER_OPEN_INFO
                End If
                NEXT_AFTER_PROGRAM_LOAD = True
            End If
            If PROGRAM_LOAD = False And NEXT_AFTER_PROGRAM_LOAD = False Then
               Call WRITE_LOGGER_OPEN_INFO
            End If
        End If
    Next
Else
    For R = 1 To UBound(AR)
        On Error Resume Next
        FOLDER_NAME = AR(1) + "\SCRIPTOR DATA\" + PATH_2
        FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_2
        Kill FILE_NAME
        
        If DOOR_OPEN_HAPPEN = True Then
            DOOR_OPEN_HAPPEN = False
            FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_9
            FR1 = FreeFile
            Open FILE_NAME For Output As #FR1
            Close #FR1
        End If
        
        FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_4
        
        A_NOW = Now
        If PROGRAM_LOAD = True Then
            PROGRAM_LOAD = False
            Call CHECK_ARCHIVE_LOGGER
            If X2 > X1 Then
                Call WRITE_LOGGER_BEGIN
                Call WRITE_LOGGER_CLOSE_INFO
            End If
            NEXT_AFTER_PROGRAM_LOAD = True
        End If
        If PROGRAM_LOAD = False And NEXT_AFTER_PROGRAM_LOAD = False Then
            Call WRITE_LOGGER_CLOSE_INFO
        End If
        
        
    Next
End If

If NEXT_AFTER_PROGRAM_LOAD = True Then
    NEXT_AFTER_PROGRAM_LOAD = False
    FOLDER_NAME = AR(1) + "\SCRIPTOR DATA\" + PATH_2
    FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_4
    Path_And_FileName = FILE_NAME
    If FindWinPart_ANY_STRING("Tail for Win32 - [Non-Workspace Files - " + Path_And_FileName + "]") = 0 Then
        If Dir(I_N_TAIL) <> "" Then
            Shell I_N_TAIL + " """ + Path_And_FileName + """", vbMinimized
        End If
    End If
End If


End Sub

Sub WRITE_LOGGER_OPEN_INFO()
    FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_4
    FR1 = FreeFile
    Open FILE_NAME For Append As #FR1
    Print #FR1, Format(A_NOW, "YYYY-MM-DD -- HH:MM:SS -- HH:MM:SS AM/PM -- DDD") + " -- DOOR OPEN"
    Close #FR1
End Sub
                
Sub WRITE_LOGGER_CLOSE_INFO()
    FILE_NAME = FOLDER_NAME + "\" + FILE_NAME_4
    FR1 = FreeFile
    Open FILE_NAME For Append As #FR1
    Print #FR1, Format(A_NOW, "YYYY-MM-DD -- HH:MM:SS -- HH:MM:SS AM/PM -- DDD") + " -- DOOR CLOSE"
    Close #FR1
End Sub

Sub WRITE_LOGGER_BEGIN()
    A1 = " -----------------------------------------"
    A2 = " -- RS232 LOGGER FRONT DOOR BEGIN"
    FR1 = FreeFile
    Open FILE_NAME For Append As #FR1
    Print #FR1, Format(A_NOW, "YYYY-MM-DD -- HH:MM:SS") + " - -" + Format(A_NOW, "HH:MM:SS AMPM -- DDD") + A1
    Print #FR1, Format(A_NOW, "YYYY-MM-DD -- HH:MM:SS") + " - -" + Format(A_NOW, "HH:MM:SS AMPM -- DDD") + A2
    Print #FR1, Format(A_NOW, "YYYY-MM-DD -- HH:MM:SS") + " - -" + Format(A_NOW, "HH:MM:SS AMPM -- DDD") + A1
    Close #FR1
End Sub

Sub CHECK_ARCHIVE_LOGGER()

    Dim STRING_VAR As String
    STRING_VAR = Space(2000)
    FR1 = FreeFile
    Open FILE_NAME For Binary As #FR1
    LOF_MINUS = LOF(FR1) - 2000
    If LOF_MINUS < 1 Then
        LOF_MINUS = 1
        STRING_VAR = LOF(FR1)
    End If
    Get #FR1, LOF_MINUS, STRING_VAR
    Close #FR1
    X1 = InStrRev(STRING_VAR, "-- DOOR CLOSE")
    X2 = InStrRev(STRING_VAR, "-- DOOR OPEN")

    STRING_VAR = ""

End Sub

Function FindWinPart_ANY_STRING(TTF) As Long

Dim Window_Title_String
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String
            
FindWinPart_ANY_STRING = 0

'Need this is you gonna use this procedure get from CIDRun ME and another one
'Find the first window
test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)

Do While test_hwnd <> 0
    
    Window_Title_String = GetWindowTitle(test_hwnd)
    cText = Space$(255)
    ghj$ = GetClassName(test_hwnd, cText, 255)
    
    If InStr(Window_Title_String, TTF) Then
        
        FindWinPart_ANY_STRING = test_hwnd
        Exit Function
        
    End If
    
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop

End Function



Function StripNulls(OriginalStr As String) As String
    'Removes NullStrings.
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function



'Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

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




Public Function CreateFolderTree(ByVal sPath As String) As Boolean
    Dim nPos As Integer

    On Error GoTo CreateFolderTreeError
    
    nPos = InStr(sPath, "\")
    If Mid(sPath, 1, 2) = "\\" Then
        nPos = 2
        For R = 1 To 3
            nPos = InStr(nPos + 1, sPath, "\")
        Next
    End If
    While nPos > 0
        
        '----------------------------------------------------------------------------
        'ROUTINE TAKEN FROM
        '----------------------------------------------------------------------------
        'SEND_TO_SCRIPT_IRFAR - Microsoft Visual Basic [design] - [mdlFileSys (Code)]
        'D:\VB6\VB-NT\00_Send_To\Send To Text List of Files & Sub Folders IRFAR\#0 Send To Text List of Files and Sub Folders IRFAN.exe
        '----------------------------------------------------------------------------
        'MODIFIED A BIT DIR COMMAND REPLACE MORE COMPLEX WAY
        '----------------------------------------------------------------------------
        
        'If Not FolderExists(Left$(sPath, nPos - 1)) Then
        
        If Dir((Left$(sPath, nPos - 1)), vbDirectory) = "" Then
            MkDir Left$(sPath, nPos - 1)
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
    'If Not FolderExists(sPath) Then MkDir sPath
    If Dir(sPath, vbDirectory) = "" Then MkDir sPath
    
    CreateFolderTree = True
    Exit Function

CreateFolderTreeError:
    Exit Function
End Function


Function GetActiveWindow(ByVal ReturnParent As Boolean) As Long
   Dim i As Long
   Dim j As Long
   i = GetForegroundWindow
   If ReturnParent Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
      i = j
   End If
   GetActiveWindow = i
End Function

Function GetParentTitle(ByVal Handle As Long) As String
   Dim i As Long
   Dim j As Long, TTx1 As String
   i = Handle
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
      i = j
   TTx1 = GetWindowTitle(i)
   GetParentTitle = TTx1

End Function

Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim L As Long
   Dim S As String
   L = GetWindowTextLength(hWnd)
   S = Space(L + 1)
   GetWindowText hWnd, S, L + 1
   GetWindowTitle = Left$(S, L)
End Function
Function GetWindowClass(ByVal hWnd As Long) As String
    Dim Ret As Long, sText As String
    sText = Space(255)
    Ret = GetClassName(hWnd, sText, 255)
    sText = Left$(sText, Ret)
   GetWindowClass = sText
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



