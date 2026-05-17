VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FCFCFC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Process Commander"
   ClientHeight    =   4365
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3012.801
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4275
      TabIndex        =   0
      Top             =   3420
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4260
      TabIndex        =   1
      Top             =   3870
      Width           =   1245
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   3195
      ScaleHeight     =   1050
      ScaleWidth      =   2535
      TabIndex        =   8
      Top             =   2115
      Width           =   2535
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "www.planetsourcecode.com"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   90
         MouseIcon       =   "frmAbout.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   675
         Width           =   2340
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "www.freevbsource.com"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   90
         MouseIcon       =   "frmAbout.frx":015E
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   405
         Width           =   1920
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For VB programmers:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   90
         TabIndex        =   9
         Top             =   45
         Width           =   2040
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.hallsoft.tk"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   3285
      MouseIcon       =   "frmAbout.frx":02B0
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1800
      Width           =   1230
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0402
      Height          =   915
      Left            =   135
      TabIndex        =   6
      Top             =   3285
      Width           =   3930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All rights reserved"
      Height          =   225
      Left            =   3285
      TabIndex        =   5
      Top             =   1395
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright© Hallsoft"
      Height          =   225
      Left            =   3285
      TabIndex        =   4
      Top             =   1035
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Process Commander"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3285
      TabIndex        =   3
      Top             =   315
      Width           =   1815
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      Height          =   225
      Left            =   3285
      TabIndex        =   2
      Top             =   675
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   2100
      Left            =   45
      Picture         =   "frmAbout.frx":04AD
      Top             =   315
      Width           =   3300
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
'// Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         '// Unicode nul terminated string
Const REG_DWORD = 4                      '// 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
        Else
            GoTo SysInfoErr
        End If
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           '// Loop Counter
    Dim rc As Long                                          '// Return Code
    Dim hKey As Long                                        '// Handle To An Open Registry Key
    Dim hDepth As Long                                      '//
    Dim KeyValType As Long                                  '// Data Type Of A Registry Key
    Dim tmpVal As String                                    '// Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  '// Size Of Registry Key Variable
    '//------------------------------------------------------------
    '// Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '//------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) '// Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError               '// Handle Error...
    
    tmpVal = String$(1024, 0)                             '// Allocate Variable Space
    KeyValSize = 1024                                       '// Mark Variable Size
    
    '//------------------------------------------------------------
    '// Retrieve Registry Key Value...
    '//------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           '// Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)
    Else                                                    '// WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   '// Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType
    Case REG_SZ
        KeyVal = tmpVal
    Case REG_DWORD
        For i = Len(tmpVal) To 1 Step -1
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
        Next
        KeyVal = Format$("&h" + KeyVal)
    End Select
    
    GetKeyValue = True
    rc = RegCloseKey(hKey)
    Exit Function
    
GetKeyError:
    KeyVal = ""
    GetKeyValue = False
    rc = RegCloseKey(hKey)
End Function

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    File_Start "www.geocities.com/hallsoftware", "Open"
End Sub

Private Sub Label7_Click()
    File_Start "www.freevbsource.com", "Open"
End Sub

Private Sub Label8_Click()
    File_Start "www.planetsourcecode.com", "Open"
End Sub
