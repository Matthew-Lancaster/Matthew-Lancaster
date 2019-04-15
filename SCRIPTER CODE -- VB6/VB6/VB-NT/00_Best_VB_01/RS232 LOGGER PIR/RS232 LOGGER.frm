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
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1956
      Top             =   1140
   End
   Begin VB.PictureBox RichTextBox1 
      Height          =   735
      Left            =   2964
      ScaleHeight     =   684
      ScaleWidth      =   1404
      TabIndex        =   4
      Top             =   384
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1860
      Top             =   720
   End
   Begin VB.Timer TimerComm4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2280
      Top             =   720
   End
   Begin VB.PictureBox MMControl1 
      Height          =   1170
      Left            =   3192
      ScaleHeight     =   1128
      ScaleWidth      =   2160
      TabIndex        =   2
      Top             =   1344
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
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "DIALER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cProcesses As New clsCnProc
Dim OLD_VAR_DSR, VAR_DSR
Private Const MONITOR_ON = -1&
Private Const MONITOR_OFF = 2&
Private Const SC_MONITORPOWER = &HF170&
Private Const WM_SYSCOMMAND = &H112

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

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

OLD_VAR_DSR = -10

Call TIMER2_TIMER

Me.Visible = True

End Sub


Sub TIMER2_TIMER()

On Error Resume Next
For R = 1 To 16
    Err.Clear
    Me.MSComm3.CommPort = R
    Me.MSComm3.Settings = "1200,N,8,1"
    Me.MSComm3.PortOpen = True
    Me.MSComm3.DTREnable = True

    VAR_DSR = Me.MSComm3.DSRHolding
    
    If Err.Number <> 8002 Then
        Exit For
    End If
Next
If Err.Number = 0 Then
     Timer2.Enabled = False
End If

' --------------------------------------
' NONE COMM PORT ALL 16 TESTER
' LEAVE HIGH TO KEEP LIGHT FOR SCREEN ON
' --------------------------------------
If R = 17 Then VAR_DSR = 1

End Sub


Private Sub Timer1_Timer()
        
Dim FSO

VAR_DSR = Me.MSComm3.DSRHolding
If Err.Number > 0 Then
    Timer2.Enabled = True
    VAR_DSR = 1
End If

If Err.Number = 8002 Then
    VAR_DSR = 1
End If
If Err.Number > 0 Then
    VAR_DSR = 1
End If

If OLD_VAR_DSR = VAR_DSR Then Exit Sub

OLD_VAR_DSR = VAR_DSR

Set FSO = CreateObject("Scripting.FileSystemObject")


' MsgBox Str(R) + " -- " + Str(VAR_DSR)
' Debug.Print Str(R) + " -- " + Str(VAR_DSR)

Dim AR(3)
AR(1) = "\\1-asus-x5dij\1_asus_x5dij_01_c_drive"
AR(2) = "\\2-asus-eee\2_asus_eee_01_c_drive"
AR(3) = "\\4-asus-gl522vw\4_asus_gl522vw_01_c_drive"

On Error Resume Next
If VAR_DSR = True Then
    For R = 1 To 3
        FOLDER_NAME = AR(R) + "\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY"
        FILE_NAME = FOLDER_NAME + "\Autokey -- 14-Brightness With Dimmer.txt"
        If FSO.FILEExists(FILE_NAME) = FLASE Then
            If FSO.FOLDERExists(FOLDER_NAME) = False Then
                RESULT = CreateFolderTree(FOLDER_NAME)
            End If
            FR1 = FreeFile
            Open FILE_NAME For Output As #FR1
            Close #FR1
        End If
    Next
Else
    For R = 1 To 3
        FILE_NAME = AR(R) + "\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-Brightness With Dimmer.txt"
        Kill FILE_NAME
    Next
End If

End Sub



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

