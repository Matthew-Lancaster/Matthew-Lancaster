VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Begin VB.Form DIALER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comm Port 4"
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
   WhatsThisHelp   =   -1  'True
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
      Interval        =   20
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
Private Const MONITOR_ON = -1&
Private Const MONITOR_OFF = 2&
Private Const SC_MONITORPOWER = &HF170&
Private Const WM_SYSCOMMAND = &H112

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long


Private Sub Form_Load()

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

MsgBox Str(R) + " -- " + Str(VAR_DSR)
Debug.Print Str(R) + " -- " + Str(VAR_DSR)
    
End

    
'F = Me.MSComm3.
'MsgBox F
'For R = 1 To 100000000
'    VAR_DSR = Me.MSComm3.DSRHolding
'    Debug.Print Str(VAR_DSR)
'Next

Exit Sub


' Flush the input buffer.
' MSComm2.InBufferCount = 0
Ci.MSComm2.InputLen = 1
    

End Sub



Function StripNulls(OriginalStr As String) As String
    'Removes NullStrings.
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function


