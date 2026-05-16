VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3225
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5550
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2225.952
   ScaleMode       =   0  'User
   ScaleWidth      =   5211.737
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   4980
      Top             =   2730
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4200
      TabIndex        =   1
      Top             =   1470
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Top             =   1050
      Width           =   1260
   End
   Begin PicClip.PictureClip pc2 
      Left            =   6300
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   50271
      _Version        =   393216
      Rows            =   50
      Picture         =   "frmAbout.frx":0CCA
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version Stamp Here"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   90
      TabIndex        =   9
      Top             =   3000
      Width           =   1515
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Information :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   8
      Top             =   1065
      Width           =   2625
   End
   Begin VB.Label LbEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail : netupdate@mrenigma.co.uk"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   90
      TabIndex        =   7
      Top             =   1545
      Width           =   2730
   End
   Begin VB.Label LbDevelover 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Darren Lawrence"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   1305
      Width           =   1665
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   4680
      Picture         =   "frmAbout.frx":36A0C
      Stretch         =   -1  'True
      Top             =   255
      Width           =   600
   End
   Begin VB.Image ImgPic 
      Height          =   720
      Left            =   150
      Picture         =   "frmAbout.frx":37123
      Top             =   210
      Width           =   720
   End
   Begin VB.Label LbNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":37FED
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   945
      Left            =   90
      TabIndex        =   4
      Top             =   1980
      Width           =   5190
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      X1              =   84.515
      X2              =   4991.06
      Y1              =   1304.511
      Y2              =   1304.511
   End
   Begin VB.Label LbEdition 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Professional Edition"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2565
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Label lblTitle2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NetUpdate Administrator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   345
      Left            =   900
      TabIndex        =   3
      Top             =   330
      Width           =   3600
   End
   Begin VB.Label LblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NetUpdate Administrator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   930
      TabIndex        =   5
      Top             =   345
      Width           =   3600
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
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
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Dim n As Integer

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
      Call StartSysInfo
End Sub

Private Sub cmdOk_Click()
      Unload Me
End Sub

Private Sub Form_Load()
      Me.Caption = "About " & App.FileDescription
      lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
      ' LblTitle.Caption = App.Title
End Sub

Public Sub StartSysInfo()
      On Error GoTo SysInfoErr
  
Dim rc As Long
Dim SysInfoPath As String
    
      ' Try To Get System Info Program Path\Name From Registry...
      If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
         ' Try To Get System Info Program Path Only From Registry...
      ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
         ' Validate Existance Of Known 32 Bit File Version
         If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
            ' Error - File Can Not Be Found...
         Else
            GoTo SysInfoErr
         End If
         ' Error - Registry Entry Can Not Be Found...
      Else
         GoTo SysInfoErr
      End If
    
      Call Shell(SysInfoPath, vbNormalFocus)
    
      Exit Sub
SysInfoErr:
      MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
Dim i As Long                                           ' Loop Counter
Dim rc As Long                                          ' Return Code
Dim hKey As Long                                        ' Handle To An Open Registry Key
Dim hDepth As Long                                      '
Dim KeyValType As Long                                  ' Data Type Of A Registry Key
Dim tmpval As String                                    ' Tempory Storage For A Registry Key Value
Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
      ' ------------------------------------------------------------
      ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
      ' ------------------------------------------------------------
      rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
      If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
      tmpval = String$(1024, 0)                             ' Allocate Variable Space
      KeyValSize = 1024                                       ' Mark Variable Size
    
      ' ------------------------------------------------------------
      ' Retrieve Registry Key Value...
      ' ------------------------------------------------------------
      rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
      KeyValType, tmpval, KeyValSize)    ' Get/Create Key Value
                        
      If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
      If (Asc(Mid(tmpval, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
      tmpval = left(tmpval, KeyValSize - 1)               ' Null Found, Extract From String
   Else                                                    ' WinNT Does NOT Null Terminate String...
      tmpval = left(tmpval, KeyValSize)                   ' Null Not Found, Extract String Only
   End If
   ' ------------------------------------------------------------
   ' Determine Key Value Type For Conversion...
   ' ------------------------------------------------------------
   Select Case KeyValType                                  ' Search Data Types...
      Case REG_SZ                                             ' String Registry Key Data Type
         KeyVal = tmpval                                     ' Copy String Value
      Case REG_DWORD                                          ' Double Word Registry Key Data Type
         For i = Len(tmpval) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpval, i, 1)))   ' Build Value Char. By Char.
         Next
         KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
   End Select
    
   GetKeyValue = True                                      ' Return Success
   rc = RegCloseKey(hKey)                                  ' Close Registry Key
   Exit Function                                           ' Exit
    
GetKeyError: ' Cleanup After An Error Has Occured...
   KeyVal = ""                                             ' Set Return Val To Empty String
   GetKeyValue = False                                     ' Return Failure
   rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


Private Sub Timer1_Timer()
      ' Timer for animated Windows XP logo
      On Error Resume Next

      Image1.Picture = pc2.GraphicCell(n)

      n = n + 1

      If n > 49 Then
         n = 0
         Image1.Picture = pc2.GraphicCell(0)
      End If
End Sub

