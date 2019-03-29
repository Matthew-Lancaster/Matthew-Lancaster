VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Ci 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caller ID Xxx xxX"
   ClientHeight    =   3420
   ClientLeft      =   10050
   ClientTop       =   585
   ClientWidth     =   9585
   ControlBox      =   0   'False
   FillColor       =   &H00FFC0C0&
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "CI.frx":0000
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3420
   ScaleWidth      =   9585
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "ReRun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6510
      TabIndex        =   27
      Top             =   240
      Width           =   1080
   End
   Begin VB.PictureBox DirectSS1 
      Height          =   420
      Left            =   2655
      ScaleHeight     =   360
      ScaleWidth      =   720
      TabIndex        =   26
      Top             =   1035
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Timer TimerStopMessage 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   8055
      Top             =   1620
   End
   Begin VB.Timer TimerComm1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8055
      Top             =   1215
   End
   Begin VB.Timer Timer11 
      Interval        =   35
      Left            =   8730
      Top             =   2850
   End
   Begin VB.Timer Timer10 
      Interval        =   1
      Left            =   9150
      Top             =   2430
   End
   Begin VB.Timer Timer9 
      Left            =   8730
      Top             =   2430
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9150
      Top             =   2025
   End
   Begin VB.Timer Timer4 
      Interval        =   500
      Left            =   8730
      Top             =   2025
   End
   Begin VB.TextBox datetime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "MON 00/JAN/1900 12:00:00 PM"
      Top             =   0
      Width           =   3030
   End
   Begin MCI.MMControl MMControl1 
      Height          =   795
      Left            =   1395
      TabIndex        =   24
      Top             =   2610
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1402
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer7 
      Interval        =   3000
      Left            =   9150
      Top             =   1605
   End
   Begin VB.Timer Timer6 
      Interval        =   200
      Left            =   8730
      Top             =   1605
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   9000
      Top             =   3300
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   7
      DTREnable       =   -1  'True
      BaudRate        =   1200
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9000
      Top             =   3855
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      Handshaking     =   3
      BaudRate        =   56000
   End
   Begin VB.CommandButton viewitone 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "&View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5625
      TabIndex        =   0
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton edit 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7575
      TabIndex        =   1
      Top             =   0
      Width           =   1995
   End
   Begin VB.CommandButton stats_Command 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "&Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6510
      TabIndex        =   2
      Top             =   0
      Width           =   1065
   End
   Begin VB.CommandButton EXITBOX 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5625
      TabIndex        =   3
      Top             =   225
      Width           =   885
   End
   Begin VB.TextBox BUSY 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4950
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   ".BUSY."
      Top             =   0
      Width           =   675
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   9600
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   0
      Width           =   1080
   End
   Begin VB.CommandButton logbut 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "&L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   9600
      TabIndex        =   7
      Top             =   480
      Width           =   360
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1395
      Left            =   9600
      TabIndex        =   8
      Top             =   2160
      Width           =   1440
   End
   Begin VB.Timer Timer5 
      Interval        =   500
      Left            =   9150
      Top             =   1185
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   8730
      Top             =   1185
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8730
      Top             =   765
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00400000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2940
      ItemData        =   "CI.frx":030A
      Left            =   0
      List            =   "CI.frx":030C
      TabIndex        =   20
      Top             =   480
      Width           =   9600
   End
   Begin VB.Label LabelRecs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1530
      TabIndex        =   23
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Rec's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1035
      TabIndex        =   25
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Ring11 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   8805
      TabIndex        =   22
      Top             =   240
      Width           =   765
   End
   Begin VB.Label dayfield 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Day Field"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label Ring13Left 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "On"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Ring14Right 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "On"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1575
      TabIndex        =   18
      Top             =   0
      Width           =   345
   End
   Begin VB.Label total_no_calls_label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Calls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2610
      TabIndex        =   9
      Top             =   240
      Width           =   405
   End
   Begin VB.Label Rec_no_label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   555
      TabIndex        =   11
      Top             =   240
      Width           =   480
   End
   Begin VB.Label rec_no_label 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Rec #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      TabIndex        =   12
      Top             =   240
      Width           =   555
   End
   Begin VB.Label Total_callers_label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Callers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3600
      TabIndex        =   13
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Total_callers_label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4215
      TabIndex        =   14
      Top             =   240
      Width           =   555
   End
   Begin VB.Label Spin_Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "/"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   4770
      TabIndex        =   15
      Top             =   255
      Width           =   180
   End
   Begin VB.Label Text_label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "No. of Rings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   7560
      TabIndex        =   16
      Top             =   240
      Width           =   1245
   End
   Begin VB.Label busybee_label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4965
      TabIndex        =   17
      Top             =   240
      Width           =   660
   End
   Begin VB.Label Total_no_calls_label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3015
      TabIndex        =   10
      Top             =   240
      Width           =   585
   End
   Begin VB.Menu Menu_File 
      Caption         =   "&File"
      Enabled         =   0   'False
      Begin VB.Menu Mnu_Std 
         Caption         =   "STd Codes Inet"
      End
      Begin VB.Menu Menu_1_FL_Explorer 
         Caption         =   "File Log Exporer"
      End
      Begin VB.Menu Menu_1_TelV2 
         Caption         =   "C:\CallerID\CI Logs\TelV2.txt"
      End
      Begin VB.Menu Menu_1_CILog 
         Caption         =   "C:\CallerID\CI Logs\CILog.txt"
      End
      Begin VB.Menu Menu_1_CID4 
         Caption         =   "C:\CallerID\CI Logs\CID4.TXT"
      End
      Begin VB.Menu Menu_BlackList 
         Caption         =   "C:\CallerID\CI Logs\CiBlacklist.txt"
      End
      Begin VB.Menu Menu_1_CIDUD6 
         Caption         =   "C:\CallerID\CI Logs\CIDUD6.TXT"
      End
      Begin VB.Menu Menu_1_CIDUD5OD 
         Caption         =   "C:\CallerID\CI Logs\CIDUD5OD.TXT"
      End
      Begin VB.Menu Menu_1_CIDud6RingCountLog 
         Caption         =   "C:\CallerID\CI Logs\CIDud6RingCountLog.txt"
      End
      Begin VB.Menu Menu_1_Cid4TimeUpdated 
         Caption         =   "C:\CallerID\CI Logs\Cid4TimeUpdated.txt"
      End
      Begin VB.Menu Menu_1_TOXIC 
         Caption         =   "C:\CallerID\CI Logs\TOXIC.TXT"
      End
   End
   Begin VB.Menu Menu_Edit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu Menu_Execute 
      Caption         =   "Execute"
      Begin VB.Menu Menu_Pulse 
         Caption         =   "Pulse"
      End
      Begin VB.Menu Menu_Date1991 
         Caption         =   "Date1991"
      End
      Begin VB.Menu Menu_Date1991Web 
         Caption         =   "Date1991 Web"
      End
      Begin VB.Menu Menu_Mp3Tags 
         Caption         =   "MP3 Tags"
      End
      Begin VB.Menu Menu_CreativePlaycenter 
         Caption         =   "Creative PlayCenter"
      End
      Begin VB.Menu Menu_AttachmentExtractor 
         Caption         =   "Attachment Extractor"
      End
      Begin VB.Menu Menu_alcohol120 
         Caption         =   "Alcohol 120%"
      End
      Begin VB.Menu Menu_IceView 
         Caption         =   "IceView"
      End
      Begin VB.Menu Menu_PhotoThumb 
         Caption         =   "Photo Thumb"
      End
      Begin VB.Menu Menu_DVDXCOPY 
         Caption         =   "DVD X Copy"
      End
   End
   Begin VB.Menu Menu_Menu 
      Caption         =   "Menu"
      Begin VB.Menu Menu_MenuInfo 
         Caption         =   "Menu Info"
      End
      Begin VB.Menu Menu_StartLogs 
         Caption         =   "Startup Logs"
      End
      Begin VB.Menu Menu_TestTrig 
         Caption         =   "Test Trig"
      End
      Begin VB.Menu Mnu_Dialer 
         Caption         =   "Dialer"
      End
      Begin VB.Menu Mnu_Winamp_p 
         Caption         =   "WinAmp Play"
      End
      Begin VB.Menu Mnu_Winamp_S 
         Caption         =   "Winamp Stop"
      End
   End
   Begin VB.Menu Menu_Internet 
      Caption         =   "Internet"
      Begin VB.Menu Menu_Ebay 
         Caption         =   "Ebay"
      End
      Begin VB.Menu Menu_Links 
         Caption         =   "Web Links Secure"
      End
      Begin VB.Menu Menu_Mondaylotto 
         Caption         =   "Monday Lotto"
      End
   End
   Begin VB.Menu Menu_View 
      Caption         =   "&View"
      Begin VB.Menu Menu_EntriesOfLogs 
         Caption         =   "Entries of logs..."
      End
      Begin VB.Menu Menu_MainTelephoneEntries 
         Caption         =   "Main Telephone Entries..."
      End
      Begin VB.Menu Menu_HighScoresEntries 
         Caption         =   "High scores entries..."
      End
      Begin VB.Menu Menu_BlackListInfo 
         Caption         =   "Black List Info"
      End
   End
   Begin VB.Menu Mnu_RRoids 
      Caption         =   "&-RR-"
   End
   Begin VB.Menu Mnu_OutPut 
      Caption         =   "OutPut"
      Begin VB.Menu Mnu_OutPut_CiStatsLog 
         Caption         =   "CiStatsLog"
      End
      Begin VB.Menu Mnu_OutPut_CiStatsLogFuzz 
         Caption         =   "CiStatsLogFuzz"
      End
   End
End
Attribute VB_Name = "Ci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Xxr1

'Searches
'CheckRs232(C, R34)
'DialCheckComPort - Global

'CommPortSet

Public Voice As SpVoice
Public speak$

Private Declare Function MoveWindow _
        Lib "user32" _
        (ByVal hWnd As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal bRepaint As Long) As Long


Public ScrWidth_Old As Double
Public Pagernumber$

Dim Kps() As Integer
'Dim HobNob$
Dim MattD$

Dim Tap2$                    ' for storing date data

Dim Sign As String * 1
Dim Counter3$
Dim TDate3$
Dim TTime3$
Dim HallFame As Integer
Dim Lsd4 As Double            'real-time ring indicator
Dim SampleComPort1 As Integer
Dim ADate As Double
Public pwe As Integer
Public SpinnerSec As Date

Public OldCIWinState

Public WithEvents MComm1 As MSComm
Attribute MComm1.VB_VarHelpID = -1



'--------------------------------------------------------
'--------------------------------------------------------


Public Sub BackUp()

'Exit Sub

BUSY.BackColor = Red1
Ci.BUSY.Visible = -1
teltext = FreeFile

On Local Error Resume Next
Kill "W:\CALLERID\CI LOGS\TelV2-9.txt"
Name "W:\CALLERID\CI LOGS\TelV2-8.txt" As "W:\CALLERID\CI LOGS\TelV2-9.txt"
Name "W:\CALLERID\CI LOGS\TelV2-7.txt" As "W:\CALLERID\CI LOGS\TelV2-8.txt"
Name "W:\CALLERID\CI LOGS\TelV2-6.txt" As "W:\CALLERID\CI LOGS\TelV2-7.txt"
Name "W:\CALLERID\CI LOGS\TelV2-5.txt" As "W:\CALLERID\CI LOGS\TelV2-6.txt"
Name "W:\CALLERID\CI LOGS\TelV2-4.txt" As "W:\CALLERID\CI LOGS\TelV2-5.txt"
Name "W:\CALLERID\CI LOGS\TelV2-3.txt" As "W:\CALLERID\CI LOGS\TelV2-4.txt"
Name "W:\CALLERID\CI LOGS\TelV2-2.txt" As "W:\CALLERID\CI LOGS\TelV2-3.txt"
Name "W:\CALLERID\CI LOGS\TelV2-1.txt" As "W:\CALLERID\CI LOGS\TelV2-2.txt"
Name "W:\CALLERID\CI LOGS\TelV2.txt" As "W:\CALLERID\CI LOGS\TelV2-1.txt"
On Local Error GoTo 0

On Local Error Resume Next
Kill "W:\CALLERID\CI LOGS\TelV2-9HighScore.txt"
Name "W:\CALLERID\CI LOGS\TelV2-8HighScore.txt" As "W:\CALLERID\CI LOGS\TelV2-9HighScore.txt"
Name "W:\CALLERID\CI LOGS\TelV2-7HighScore.txt" As "W:\CALLERID\CI LOGS\TelV2-8HighScore.txt"
Name "W:\CALLERID\CI LOGS\TelV2-6HighScore.txt" As "W:\CALLERID\CI LOGS\TelV2-7HighScore.txt"
Name "W:\CALLERID\CI LOGS\TelV2-5HighScore.txt" As "W:\CALLERID\CI LOGS\TelV2-6HighScore.txt"
Name "W:\CALLERID\CI LOGS\TelV2-4HighScore.txt" As "W:\CALLERID\CI LOGS\TelV2-5HighScore.txt"
Name "W:\CALLERID\CI LOGS\TelV2-3HighScore.txt" As "W:\CALLERID\CI LOGS\TelV2-4HighScore.txt"
Name "W:\CALLERID\CI LOGS\TelV2-2HighScore.txt" As "W:\CALLERID\CI LOGS\TelV2-3HighScore.txt"
Name "W:\CALLERID\CI LOGS\TelV2-1HighScore.txt" As "W:\CALLERID\CI LOGS\TelV2-2HighScore.txt"
Name "W:\CALLERID\CI LOGS\TelV2HighScore.txt" As "W:\CALLERID\CI LOGS\TelV2-1HighScore.txt"
On Local Error GoTo 0

On Local Error Resume Next
Kill "W:\CALLERID\CI LOGS\TelV2-9History.txt"
Name "W:\CALLERID\CI LOGS\TelV2-8History.txt" As "W:\CALLERID\CI LOGS\TelV2-9History.txt"
Name "W:\CALLERID\CI LOGS\TelV2-7History.txt" As "W:\CALLERID\CI LOGS\TelV2-8History.txt"
Name "W:\CALLERID\CI LOGS\TelV2-6History.txt" As "W:\CALLERID\CI LOGS\TelV2-7History.txt"
Name "W:\CALLERID\CI LOGS\TelV2-5History.txt" As "W:\CALLERID\CI LOGS\TelV2-6History.txt"
Name "W:\CALLERID\CI LOGS\TelV2-4History.txt" As "W:\CALLERID\CI LOGS\TelV2-5History.txt"
Name "W:\CALLERID\CI LOGS\TelV2-3History.txt" As "W:\CALLERID\CI LOGS\TelV2-4History.txt"
Name "W:\CALLERID\CI LOGS\TelV2-2History.txt" As "W:\CALLERID\CI LOGS\TelV2-3History.txt"
Name "W:\CALLERID\CI LOGS\TelV2-1History.txt" As "W:\CALLERID\CI LOGS\TelV2-2History.txt"
Name "W:\CALLERID\CI LOGS\TelV2History.txt" As "W:\CALLERID\CI LOGS\TelV2-1History.txt"
On Local Error GoTo 0

teltext = FreeFile
Open "W:\CALLERID\CI LOGS\TelV2History.txt" For Output As #teltext
For r = 0 To Ciform1.List2.ListCount - 1
    Print #teltext, Ciform1.List2.List(r)
Next
Close #teltext

teltext = FreeFile
Open "W:\CALLERID\CI LOGS\TelV2HighScore.txt" For Output As #teltext
For r = 0 To Ciform4.List2.ListCount - 1
    Print #teltext, Ciform4.List2.List(r)
Next
Close #teltext

teltext = FreeFile
Open "W:\CALLERID\CI LOGS\TelV2.txt" For Output As #teltext
For r = 0 To Ciform2.List2.ListCount - 1
    Print #teltext, Ciform2.List2.List(r)
Next
Close #teltext


BUSY.Visible = False

End Sub


Private Sub Command1_Click()
ReRun = True
Unload Ci

End Sub

Private Sub Edit_Click()

BUSY.Visible = -1
Which = 0

If Ci.List1.Visible = 0 Then Ci.List1.Visible = -1

'Ci.List1.SetFocus
Load CiEDIT
CiEDIT.Visible = -1
CiEDIT.Enabled = -1

DoIt2 = 1
CiEDIT.Speed_search_text.Enabled = -1
CiEDIT.Speed_search_text.SetFocus
BUSY.Visible = 0

CiEDIT.Top = Ci.Top
CiEDIT.Left = Ci.Left

End Sub



Private Sub ExitCiEdit()

CiEDIT.Visible = 0
CiEDIT.Enabled = 0
Unload CiEDIT
Ci.Enabled = True
'Ci.List1.SetFocus

End Sub

'***********************************************
'# Check, whether we are in the IDE
Public Function IsIDE() As Boolean
  
Debug.Assert Not TestIDE(IsIDE)

'# In the IDE, the forwarding of
'# keystrokes will be suppressed

End Function



Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function




Public Sub SpeakV(speak$)

'Public Voice As SpVoice
'Set Voice.Voice = Voice.GetVoices().Item(3) 'microsoft mike

'Set Voice.Voice = Voice.GetVoices().Item(3) 'sample TTS
'Set Voice.Voice = Voice.GetVoices().Item(2) 'microsoft Sam
'Set Voice.Voice = Voice.GetVoices().Item(0) 'microsoft mary
Set Voice = New SpVoice
If IsIDE = True Then Set Voice.Voice = Voice.GetVoices().Item(2) 'microsoft mike
If IsIDE = False Then Set Voice.Voice = Voice.GetVoices().Item(3) 'microsoft mike

If speak$ = "" Then speak$ = "Warriarr"


On Local Error Resume Next
Voice.speak speak$, SVSFlagsAsync

End Sub



Private Sub Form_Activate()



If Activated = True Then Exit Sub

Call Ci.SpeakV("")

Activated = True




GoTo skip1
ShowWindow wed, SW_MINIMIZE
ShowWindow wed, SW_HIDE
ShowWindow wed, SW_SHOWNORMAL
ShowWindow wed, SW_NORMAL
ShowWindow wed, SW_SHOWMINIMIZED
ShowWindow wed, SW_SHOWMAXIMIZED
ShowWindow wed, SW_MAXIMIZE
ShowWindow wed, SW_SHOWNOACTIVATE
ShowWindow wed, SW_SHOW
ShowWindow wed, SW_MINIMIZE
ShowWindow wed, SW_SHOWMINNOACTIVE
ShowWindow wed, SW_SHOWNA
ShowWindow wed, SW_RESTORE
ShowWindow wed, SW_SHOWDEFAULT
ShowWindow wed, SW_FORCEMINIMIZE
ShowWindow wed, SW_MAX
skip1:



OldRect22Top.Top = -50

BumSqeak = -1

'Normal
Ab1$ = Chr$(&H11) + Chr$(1) + Chr$(1) + Chr$(1) + Chr$(8)
'Test
Ab2$ = Chr$(&H80) + Chr$(&HC) + Chr$(2) + Chr$(&HA)
'RingBack
Ab3$ = Chr$(&H11) + Chr$(1) + Chr$(2) + Chr$(1) + Chr$(8)




Slashing_Rain = Timer

Call Ci.StartUpTimes("Start")

Conk2 = -1

Ci.BUSY.Visible = -1
Ci.BUSY.Refresh

Call SeXxYy
Call DIALER.DialerTop

Red1 = QBColor(12)
Yellow1 = QBColor(14)
White1 = QBColor(15)
Green1 = QBColor(10)
Blue1 = Ci.List1.BackColor

Call Qwerty

Cheese = True

Ci.Rec_no_label2.Caption = "0"

Timer1.Enabled = False
Timer3.Enabled = True   '--
Timer3.Interval = 1     '--
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False

Timer9.Enabled = True
Timer9.Interval = 1


Ci.Caption = "Caller ID - Ver. " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
         
Boat2$ = "OK"
       
Monk2$ = "off"

Yts = -3

Ci.dayfield.Caption = "Week Day"

AiX = 0

'HobNob$ = "Phone"

Puff = Now

If App.PrevInstance Then
    MsgBox "Bye bye!", vbInformation, App.Title & " is already running."
    End
End If


FirstRun0 = 0
FirstRun2 = 0

DoIt2 = 10

DoEvents

Call Timer3_Timer

HallFame = 110
RecList = 250
Clock = -1

FindPos = 45
CodePos = 42
CountPos = 37
TdatePos = 17
TdatePos2 = 25
TTimePos = 28

er = 0

'Ciform1.Show

Load Ciform1
Load Ciform2
Load Ciform4

Ci.Show

Call LoadFile

Ci.Refresh

FastFind2String$ = "?"

For rt = 0 To Ciform2.List2.ListCount
    FastFind2String$ = FastFind2String$ + Format$(rt, "00000") + "***v*" + Ciform2.List2.List(rt)
Next

Call StartUpTimes("Load File")







ReDim TimeUpdated(200) As Date
TimeUpdateReDim = 200
TimeUpdatedCount = 0

filef7 = FreeFile
Open "W:\CALLERID\CI LOGS\Cid4TimeUpdated.txt" For Input As #filef7
Do
    Line Input #filef7, lkj$
    date3 = TimeValue(lkj$) + DateValue(lkj$)
    TimeUpdated(TimeUpdatedCount) = date3
    TimeUpdatedCount = TimeUpdatedCount + 1
    If TimeUpdatedCount >= TimeUpdateReDim Then
        TimeUpdateReDim = TimeUpdateReDim + 100
        ReDim Preserve TimeUpdated(timepdateredim) As Date
    End If
  
Loop Until EOF(filef7)

Close #filef7



Call StartUpTimes("Load File CID4 Time Updated")



Ciform1.BUSY.Visible = 0
Ciform2.BUSY.Visible = 0
Ciform4.BUSY.Visible = 0

Ci.List1.AddItem "DATE     TIME     NUMBER           NAME QNTY THE DESCRIPTION "
Ci.List1.AddItem "컴컴컴컴쩡컴컴컴컫컴컴컴컴컴컴컴컴쩡컴컫컴컴쩡컴컴컴컴컴컴컴컴"
Ci.List1.AddItem "컴컴컴컴좔컴컴컴컨컴컴컴컴컴컴컴컴좔컴컨컴컴좔컴컴컴컴컴컴컴컴"
Ci.List1.AddItem "Recorded Numbers...NEW"
Ci.List1.AddItem "컴컴컴컴쩡컴컴컴컫컴컴컴컴컴컴컴컴쩡컴컫컴컴쩡컴컴컴컴컴컴컴컴"

Ci.List1.ListIndex = Ci.List1.ListCount - 1

Hurig$ = ""

ReDim Kps((Ciform1.List2.ListCount) + 2)

DoEvents

For re = 0 To Ciform1.List2.ListCount - 1
    Kps(re) = -1
Next




''''''''''''add history of Recorded Numbers...NEW

'got a Problem check the code number

For re = 0 To Ciform1.List2.ListCount - hup
    Linea$ = Ciform1.List2.List(re)
    If LTrim(Linea$) <> "" Then
        lineb$ = Linea$
        Kps(re) = re
        
        
        Call Sort_Ci_FastFind
        Number$ = Mid$(Linea$, 1, 16)
        Counter$ = "    "
        RSet Counter$ = Trim(Str$(Val(Mid$(Linea$, CountPos, 4))))
        Tdate$ = Mid$(Linea$, TdatePos, 6) + Mid$(Linea$, TdatePos2, 2)
        Ldate$ = Mid$(Linea$, TdatePos, 10)
        Ttime$ = Mid$(Linea$, TTimePos, 8)
        
        
        Word$ = "    "
        If Mid$(Code$, 1, 2) = "01" Then
            Word$ = Mid$(Find$, 1, 4)
            If Mid$(Number$, 1, 2) = "00" Then Word$ = "INTL"
        End If

        If Mid$(Code$, 1, 2) = "02" Then
            Word$ = Mid$(Find$, 1, 4)
        End If

        If InStr(Find$, RTrim(Number$)) = 1 Then Word$ = Mid$(Find$, 1, 4)
        
        
        
        
        
        DateLine = False
        If Ldate$ <> "00/00/1900" Then
            DateLine = DateValue(Ldate$) + TimeValue(Ttime$)

            'Hurig$ = Hurig$ + Str(DateLine) + " "
        
            Call TimeBeenUpdated(DateLine)
        End If
        
        If DateLine = True Then rh$ = Chr$(254) Else rh$ = "�"
        
        numberx$ = Number$
        If Len(Number$) > 8 And InStr("0123456789", Mid$(Number$, 1, 1)) > 0 Then
            Number$ = Mid$(Number$, 1, 5) + "-" + Mid$(Number$, 6)
            Number$ = Left$(Number$, Len(Number$) - 1)
        End If
'        Number$ = Mid$(Number$, 1, 5) + "-" + Mid$(Number$, 6)
'        Number$ = numberx$
 
        Ci.List1.AddItem Tdate$ + "�" + Ttime$ + rh$ + Number$ + "�" + Word$ + "�" + Counter$ + "�" + Find$ ', Ci.List1.ListCount - 7
        Number$ = numberx$
        
        If Ci.List1.ListCount = 70 Then Ci.List1.Refresh
        
    End If
Next




Ci.List1.ListIndex = List1.ListCount - 1



Call StartUpTimes("Add History Of Recorded Numbers New")


DoEvents

Ci.List1.AddItem "컴컴컴컴좔컴컴컴컨컴컴컴컴컴컴컴컴좔컴컨컴컴좔컴컴컴컴컴컴컴컴"
Ci.List1.AddItem "Recorded Numbers...OLD, Merged With Hall of Fame..."
Ci.List1.AddItem "컴컴컴컴쩡컴컴컴컫컴컴컴컴컴컴컴컴쩡컴컫컴컴쩡컴컴컴컴컴컴컴컴"

pwe = pwe + 3

Ci.List1.ListIndex = List1.ListCount - 1

' date sorting .. Recorded Numbers...OLD





' date sorting .. Recorded Numbers...OLD-------Dead
    
a$ = lineb$
    
Tdate$ = Mid$(a$, TdatePos, 10)
Ttime$ = Mid$(a$, TTimePos, 8)
            
If Tdate$ = "00/00/1900" Then
    nartrt5 = DateSerial(1900, 1, 1)
Else
    Nartrt1 = DateValue(Tdate$)
    nartrt5 = Nartrt1 + TimeValue(Ttime$)
End If

Call pok

'For poir2 = 0 To Ciform1.List2.ListCount - 1
For poir2 = 0 To De1 - 1
    'a$ = Ciform1.List2.List(poir2)
    
    a$ = Demagic$(poir2 + 1)
    
    If Trim(a$) = "" Then Exit For
    
    Tdate$ = Mid$(a$, TdatePos, 10)
    Ttime$ = Mid$(a$, TTimePos, 8)
            
          
            
    If Tdate$ = "00/00/1900" Then
        Nartrt2 = DateSerial(1900, 1, 1)
    Else
        Nartrt1 = DateValue(Tdate$)
        On Local Error Resume Next
        
        Nartrt2 = Nartrt1 + TimeValue(Ttime$)
        If Err.Number > 0 Then
            For r3 = 1 To Len(Ttime$)
                e2 = InStr("0123456789:", Mid$(Ttime$, r3, 1))
                If e2 = 0 Then Mid$(Ttime$, r3, 1) = "0"
            Next
        End If
        On Local Error GoTo 0
    
    End If
        
        
        
        
    Linea$ = a$
        
    If Nartrt2 < nartrt5 Then
        'If Len(Linea$) < 16 Then Stop
        CiScratch.List4.AddItem Format$(Nartrt2, "000000.000000000") + " " + Format$(999 - poir2, "0000") + " - " + Linea$
    End If
Next

For poir2 = 0 To Ciform4.List2.ListCount - 1
    a$ = Ciform4.List2.List(poir2)
    Tdate$ = Mid$(a$, TdatePos, 10)
    Ttime$ = Mid$(a$, TTimePos, 8)
            
    If Tdate$ = "00/00/1900" Then
        Nartrt2 = DateSerial(1900, 1, 1)
    Else
        Nartrt1 = DateValue(Tdate$)
        Nartrt2 = Nartrt1 + TimeValue(Ttime$)
    End If
        
    Linea$ = a$
        
    If Nartrt2 < nartrt5 Then
        'If Len(Linea$) < 16 Then Stop

        CiScratch.List4.AddItem Format$(Nartrt2, "000000.000000000") + " " + Format$(999 - poir2, "0000") + " * " + Linea$
    End If
Next


For rert = CiScratch.List4.ListCount - 1 To 0 Step -1
    lineb$ = CiScratch.List4.List(rert)
    Linea$ = Mid$(lineb$, 25)
    
'    Call Sort_Ci_Find
    Call Sort_Ci_FastFind
    
    Counter$ = Mid$(Linea$, CountPos, 4)
    RSet Counter$ = LTrim(Str(Val(Counter$)))
        
    Tdate$ = Mid$(Linea$, TdatePos, 6) + Mid$(Linea$, TdatePos2, 2)
    Ttime$ = Mid$(Linea$, TTimePos, 8)
    Ldate$ = Mid$(Linea$, TdatePos, 10)
    
    If Mid$(lineb$, 23, 1) = "-" Then Sign = "�" Else Sign = Chr$(254)
     
    'rh$ =  "�"
    If Ldate$ = "00/00/1900" Then Ldate$ = "01/01/1900"
    On Local Error Resume Next
    DateLine = DateValue(Ldate$) + TimeValue(Ttime$)
    
        
        If Err.Number > 0 Then
            For r3 = 1 To Len(Ttime$)
                e2 = InStr("0123456789:", Mid$(Ttime$, r3, 1))
                If e2 = 0 Then Mid$(Ttime$, r3, 1) = "0"
            Next
        End If
        On Local Error GoTo 0
    
    
    
    Call TimeBeenUpdated(DateLine)
    If DateLine = True Then rh$ = Chr$(254) Else rh$ = "�"

    numberx$ = Number$
    If Len(Number$) > 8 And InStr("0123456789", Mid$(Number$, 1, 1)) > 0 Then
        Number$ = Mid$(Number$, 1, 5) + "-" + Mid$(Number$, 6)
        Number$ = Left$(Number$, Len(Number$) - 1)
    End If
'    Number$ = Mid$(Number$, 1, 5) + "-" + Mid$(Number$, 6)
'    Number$ = numberx$
    
    Ci.List1.AddItem Tdate$ + "�" + Ttime$ + rh$ + Number$ + Sign + Word$ + "�" + Counter$ + "�" + Find$ ', Ci.List1.ListCount - 4
    Number$ = numberx$
    If SpinnerSec <> Now Then DoEvents: SpinnerSec = Now
Next




Call StartUpTimes("Sorting All Data For Dump Into CI List Box")


Ci.List1.AddItem "컴컴컴컴좔컴컴컴컨컴컴컴컴컴컴컴컴좔컴컨컴컴좔컴컴컴컴컴컴컴컴"
Ci.List1.AddItem "Hall of Fame..."
Ci.List1.AddItem "컴컴컴컴쩡컴컴컴컫컴컴컴컴컴컴컴컴쩡컴컫컴컴쩡컴컴컴컴컴컴컴컴"

Ci.List1.ListIndex = Ci.List1.ListCount - 1

Call HallFameCreate

Ci.List1.AddItem "컴컴컴컴좔컴컴컴컨컴컴컴컴컴컴컴컴좔컴컨컴컴좔컴컴컴컴컴컴컴컴"

Ci.List1.ListIndex = Ci.List1.ListCount - 1

Call StartUpTimes("Hall Of Fame Create")






ReDim timeupdater(0) As Date

 








Sign = Chr$(181)
Word$ = ""

dew$ = Dir$("W:\CALLERID\CI LOGS\update.txt")

If dew$ <> "" Then
    cilogbak = FreeFile
    Open "W:\CALLERID\CI LOGS\cilogbak.txt" For Output As #cilogbak
    cilog = FreeFile
    Open "W:\CALLERID\CI LOGS\cilog.txt" For Input As #cilog
    Do
        Line Input #cilog, Linea$
        r$ = Mid$(Linea$, 1, 44)
        Call Sort_Ci_FastFind
        r$ = r$ + Find$
        Print #cilogbak, r$
        
        DoEvents
    
    Loop Until EOF(cilog)
    Close #cilog
    Close #cilogbak

    Name "W:\CALLERID\CI LOGS\cilog.txt" As "W:\CALLERID\CI LOGS\cilog2.txt"
    Name "W:\CALLERID\CI LOGS\cilogbak.txt" As "W:\CALLERID\CI LOGS\cilog.txt"
    Name "W:\CALLERID\CI LOGS\cilog2.txt" As "W:\CALLERID\CI LOGS\cilogbak.txt"
    
    On Local Error Resume Next
    Kill "W:\CALLERID\CI LOGS\update.txt"
    On Local Error GoTo 0

End If

Bigdropkicker = 0
Yts = 0
heyhello = 0
Egghead = 0

'Ci.DirectSS1.Pitch = 200
'Ci.DirectSS1.Speed = 200

'Ci.DirectSS1.Style (0)
'Ci.DirectSS1.Speak "System Activated"

Call Ci.New_Number_Counter

Call StartUpTimes("New Number Counter")

Load CiMessage

Call StartUpTimes("Load CiMessage")




'temp
'DIALER.Show

Call StartUpTimes("Dialer Show")

Call DIALER.DialerTop

Load frmReceiver
Load frmSender

'frmReceiver.Show
'frmSender.Show

Load CiStats

Call StartUpTimes("Load CiStats")

free5 = FreeFile
Open "W:\CALLERID\ci logs\2cidrunme.txt" For Input As #free5
Do
    Line Input #free5, aed$
    If InStr(aed$, "Last System ReBoot Time") Then
        wees$ = Mid$(aed$, 1, 16)
        If InStr(wees$, "***") = 0 Then
            a = Now
            Last_Reboot = DateValue(wees$)
            Last_Reboot = Last_Reboot + TimeValue(wees$)
        End If
    End If
Loop Until EOF(free5)
Close #free5

Call StartUpTimes("Get Last System Reboot")

Ci.Ring14Right.Caption = ".?."
Ci.Ring13Left.BackColor = Yellow1
Ci.Ring13Left.ForeColor = 0
Ci.Ring14Right.BackColor = Yellow1
Ci.Ring14Right.ForeColor = 0
Ci.Ring13Left.Caption = ".?."

Dragon2St5$ = "...--??--... ---****---"
CiStats.onoffhook_label.Caption = Dragon2St5$
CiStats.onoffhook_label.BackColor = Yellow1
CiStats.onoffhook_label.ForeColor = 0
CiStats.Label31.BackColor = Yellow1
CiStats.Label32.BackColor = Yellow1

FirstRun2 = 1



Ci.MMControl1.Notify = True
Ci.MMControl1.Wait = True
Ci.MMControl1.Shareable = False
Ci.MMControl1.DeviceType = "WaveAudio"
Ci.MMControl1.FileName = "H:\wave's\Camera1a_2_8kHz.wav"
Ci.MMControl1.Command = "Open"
Ci.MMControl1.Command = "Play"

Ci.BUSY.Visible = 0

CiStats.Label34.Caption = Format$(Timer - Slashing_Rain, " 0.00000.00")


Call CheckRs232(C, R34)

Timer1.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Timer7.Enabled = True
Timer8.Enabled = True

Timer1.Interval = 600
Timer3.Interval = 1000
Timer4.Interval = 600
Timer5.Interval = 200
Timer6.Interval = 300
Timer7.Interval = 8000
Timer8.Interval = 1000
Timer9.Interval = 200

CiStats.Timerx.Enabled = True
CiStats.Timerx.Interval = 1000

DIALER.Timer1.Enabled = True
DIALER.Timer1.Interval = 20

Number$ = "*-*-*-\--/-*-*-*"




Call StartUpTimes("Check RS232")

Call StartUpTimes("Finished")


Slashing_Rain = Timer - Slashing_Rain

Ci.List1.SetFocus

Ci.List1.ListIndex = 0

Load Restarter

CiStats.Hide

End Sub


Sub StartUpTimes(PlimSole$)

Dim wet As Double

rt1$ = "      "

rt2$ = Format$((Timer - Slashing_Rain), "0.00")

RSet rt1$ = rt2$

wet = Val(Mid$(CiScratch.List1.List(CiScratch.List1.ListCount - 1), 9, 6))

rt3$ = "      "

rt4$ = Format$(Val((rt2$) - wet), "0.00")

If Val(rt4$) > 0 Then RSet rt3$ = rt4$

CiScratch.List1.AddItem Format$(Now, "HH:MM:SS") + rt1$ + rt3$ + " " + PlimSole$

End Sub


Private Sub TimeBeenUpdated(DateLine)
    
            
For r1 = 0 To TimeUpdatedCount
    If DateLine = TimeUpdated(r1) Then
        DateLine = True: e = 1: Exit For
    End If
Next

If e = 0 Then DateLine = False

End Sub




Private Sub HallFameCreate()

'----------------- add already obtained hi score nos

If Ciform4.List2.ListCount > 0 Then
    For rlp = Ci.List1.ListCount - 1 To 0 Step -1
        If Mid$(Ci.List1.List(rlp), 18, 1) = "�" Then Exit For
    Next

    For r = 0 To Ciform4.List2.ListCount - 1
        'If r = Ciform4.List2.ListCount - 1 Then Stop
        Linea$ = Ciform4.List2.List(r)
'        Number$ = Mid$(Linea$, 1, 16)
        If Linea$ <> "" Then
            Call Sort_Ci_FastFind
            
            
            
            
            'rh$ = "�"
            
            
            If Ldate$ <> "00/00/1900" Then
                DateLine = DateValue(Ldate$) + TimeValue(Ttime$)
            Else
                DateLine = DateSerial(1900, 1, 1)
            End If
            Call TimeBeenUpdated(DateLine)
            If DateLine = True Then rh$ = Chr$(254) Else rh$ = "�"
    
            
            
            
            numberx$ = Number$
            
            If Len(Number$) > 8 And InStr("0123456789", Mid$(Number$, 1, 1)) > 0 Then
                Number$ = Mid$(Number$, 1, 5) + "-" + Mid$(Number$, 6)
                Number$ = Left$(Number$, Len(Number$) - 1)
            End If
            
            ' Number$ = numberx$
            Ci.List1.AddItem Tdate$ + "�" + Ttime$ + rh$ + Number$ + "�" + Word$ + "�" + Counter$ + "�" + Find$ ', Ci.List1.ListCount - 1
            Number$ = numberx$
            If SpinnerSec <> Now Then DoEvents: SpinnerSec = Now

        End If
    Next
End If

Ci.List1.ListIndex = Ci.List1.ListCount - 1
Ci.Refresh



Linea$ = ""
If Ciform4.List2.ListCount < HallFame Then
    For r = Ciform4.List2.ListCount - 1 To 0 Step -1
        Linea$ = Ciform4.List2.List(r)
        If RTrim(Linea$) <> "" Then Exit For
    Next

    If RTrim(Linea$) <> "" Then jjj = Val(Mid$(Linea$, CountPos, 4)) Else jjj = 1220

    poi = 0
    hhh = 2000

    Do
        lll = 0
        Do
            kkk = 0

            For Rtd = poi To Ciform2.List2.ListCount - 1
                feed = 1
                If InStr(Ciform2.List2.List(Rtd), "L I S T    2") = 0 Then
                    Linea$ = Ciform2.List2.List(Rtd)
                    Counter$ = Mid$(Linea$, CountPos, 4)
                    www = Val(Counter$)
                    If www > lll And www < hhh Then lll = www
                    If www = 0 And Counter$ <> "0000" Then www = -1
                    If www = jjj Then kkk = 1: Exit For
                End If

                If poi = 0 And Rtd = Ciform2.List2.ListCount - 1 Then feed = 0
            Next Rtd

            If feed = 0 Then If lll > 0 Then jjj = lll: lll = 0
            If Rtd <> Ciform2.List2.ListCount - 1 Then poi = Rtd + 1
            If poi >= Ciform2.List2.ListCount - 1 And feed = 1 Then jjj = jjj - 1
            If poi >= Ciform2.List2.ListCount - 1 Then poi = 0
        Loop Until kkk = 1 Or Ciform4.List2.ListCount >= Ciform2.List2.ListCount - 1
        
        hhh = jjj
        mmm = 0

        For re = 0 To Ciform4.List2.ListCount - 1
            If Mid$(Linea$, 1, 16) = Mid$(Ciform4.List2.List(re), 1, 16) Then mmm = 1: Exit For
        Next re

        If mmm = 0 Then jjj = www: k$ = Linea$
        
        If k$ <> "" Then
            Ciform4.List2.AddItem k$, Ciform4.List2.ListCount
'            Call Sort_Ci_Find
            Call Sort_Ci_FastFind

            numberx$ = Number$
            If Len(Number$) > 8 And InStr("0123456789", Mid$(Number$, 1, 1)) > 0 Then
                Number$ = Mid$(Number$, 1, 5) + "-" + Mid$(Number$, 6)
                Number$ = Left$(Number$, Len(Number$) - 1)
            End If
'            Number$ = Mid$(Number$, 1, 5) + "-" + Mid$(Number$, 6)
        '    Number$ = numberx$
            Ci.List1.AddItem Tdate$ + "�" + Ttime$ + "�" + Number$ + "�" + Word$ + "�" + Counter$ + "�" + Find$ ', Ci.List1.ListCount - 1
            Number$ = numberx$
            If SpinnerSec <> Now Then DoEvents: SpinnerSec = Now
        End If

        k$ = ""
    Loop Until HallFame <= Ciform4.List2.ListCount Or Ciform4.List2.ListCount >= Ciform2.List2.ListCount - 1
End If


End Sub

Private Sub HallFameSub()

Tdate$ = Mid$(Tdate$, 1, 6) + Mid$(Tdate$, 9)
r = Ci.List1.ListCount - 2

For rlp = Ci.List1.ListCount To 0 Step -1
    If Mid$(Ci.List1.List(rlp), 18, 1) = "�" Then Exit For
Next

For rlt = Ci.List1.ListCount To 0 Step -1
    If Mid$(Ci.List1.List(rlt), 18, 1) = "�" Then Exit For
Next

ool = rlt
count2$ = Mid$(Ci.List1.List(rlp), 41, 4)

If Val(Counter$) > Val(count2$) Or Ciform4.List2.ListCount < HallFame Then
    hhh = 0

    For r = Ci.List1.ListCount - 1 To ool Step -1
        If Mid$(Ci.List1.List(r), 18, 1) = "�" Then hhh = r + 1: oop = r
        If Mid$(Ci.List1.List(r), 18, 1) = "�" Then
            count2$ = Mid$(Ci.List1.List(r), 41, 4)
            If Val(Counter$) < Val(count2$) Then hhh = r + 1: oop = r
        End If

        If hhh > 0 Then Exit For
    Next

    For wr = Ciform4.List2.ListCount - 1 To 0 Step -1
        number2$ = Mid$(Ci.List1.List(wr + hhh), 19, 16)
        Number3$ = Mid$(Ciform4.List2.List(wr), 1, 16)
        If Number3$ = Number$ Then Ciform4.List2.RemoveItem wr: jagx = wr
        If number2$ = Number$ Then Ci.List1.RemoveItem wr + hhh
    Next

    If hhh = 0 Then hhh = Ci.List1.ListCount - (HallFame + 2)
    RSet Counter$ = LTrim(Str(Val(Counter$)))
    
    numberx$ = Number$
    If Len(Number$) > 8 And InStr("0123456789", Mid$(Number$, 1, 1)) > 0 Then
        Number$ = Mid$(Number$, 1, 5) + "-" + Mid$(Number$, 6)
        Number$ = Left$(Number$, Len(Number$) - 1)
    End If
'    Number$ = Mid$(Number$, 1, 5) + "-" + Mid$(Number$, 6)
'    Number$ = numberx$
    Ci.List1.AddItem Tdate$ + "�" + Ttime$ + "�" + Number$ + "�" + Mid$(Word$, 1, 4) + "�" + Counter$ + "�" + Find$, hhh
    Number$ = numberx$

    count2$ = Mid$(Ci.List1.List(hhh), 41, 4)
    count2$ = Format$(Val(count2$), "0000")
    number2$ = Mid$(Ci.List1.List(hhh), 19, 16)

    If YuBt < Ciform2.List2.ListCount Then
        If YuBt > -1 Then
            code2$ = Mid$(Ciform2.List2.List(YuBt), CodePos, 2)
            ttime2$ = Mid$(Ciform2.List2.List(YuBt), TTimePos, 8)
            tdate2$ = Mid$(Ciform2.List2.List(YuBt), TdatePos, 10)
            find2$ = Mid$(Ciform2.List2.List(YuBt), FindPos)
        End If
    End If

    If oop = 0 Then oop = ool
    Ciform4.List2.AddItem number2$ + tdate2$ + " " + ttime2$ + " " + count2$ + " " + code2$ + " " + find2$, (oop - ool)

    If jagx <> (oop - ool) Then
        Message2$ = number2$ + " Moved up To Position " + Trim(Str((oop - ool) + 1)) + " From " + Trim(Str(jagx + 1) + " Count = " + Trim(count2$) + " - " + find2$)
        CiMessage.List1.AddItem Message2$
        CiMessage.Visible = True
    End If

End If

If Ciform4.List2.ListCount > HallFame Then
    Ciform4.List2.RemoveItem Ciform4.List2.ListCount - 1
End If

End Sub



Public Sub Form_Terminate()

For i = Forms.Count - 1 To 0 Step -1
    Forms(i).Enabled = False
Next
End
Cheese = False

For i = Forms.Count - 1 To 0 Step -1
qq = 0
If Forms(i).Name = "Restarter" Then
    qq = 1
End If
If ReRun = False Then qq = 0

If qq = 0 Then Unload Forms(i)
Next



End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Form_Terminate
'Call ReRunner
End Sub


Private Sub EXITBOX_Click()

Call Form_Terminate

End Sub


Private Sub Form_Load()
Ci.Left = Screen.Width - Ci.Width
Ci.Show
End Sub

Public Sub New_Number_Counter()
Dim aws51$(100)
tigstomp55 = 1

For ri1 = 0 To Ci.List1.ListCount - 1
    aws1$ = Mid$(Ci.List1.List(ri1), 46)
    If InStr(UCase$(aws1$), "NEW NUMBER") = 1 Then
        aws3$ = Mid$(Ci.List1.List(ri1), 19, 16)
        spop = 0
        
        For ri5 = 1 To tigstomp55
            If aws3$ = aws51$(ri5) Then spop = 1
        Next

        If spop = 0 Then aws51$(tigstomp55) = aws3$: tigstomp55 = tigstomp55 + 1
    End If

    If InStr(UCase$(aws1$), "Recorded Numbers...OLD") > 0 Then
        Exit For
    End If
Next

For r = 0 To Forms.Count - 1
    wed$ = Forms(r).Caption
    If wed$ = "Statistics" Then
        If tigstomp55 - 1 > 0 Then
            CiStats.New_Number_Count_Label.BackColor = Blue1
            CiStats.New_Number_Count_Label.ForeColor = White1
        Else
            CiStats.New_Number_Count_Label.BackColor = Yellow1
            CiStats.New_Number_Count_Label.ForeColor = 0
        End If
            
        CiStats.New_Number_Count_Label.Caption = Str(tigstomp55 - 1)
    
    End If
Next


End Sub






Private Sub List1_Click()

If FirstRun2 = 0 Then Exit Sub

Call Record_Number_text

Call Ring_Indicator

Ldate$ = CiMessage.Rings_label(0).Caption

On Local Error Resume Next

Ci.dayfield.Caption = Format$(DateValue(Ldate$), "DDDD")

End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)

Call Record_Number_text

Call Ring_Indicator


If KeyCode <> 16 And KeyCode <> 18 And KeyCode <> 84 Then

   ' Ci.MSComm1.Output = "" + vbCr

End If

If Pager1 > 0 And KeyCode <> 16 And KeyCode <> 18 And KeyCode <> 84 Then

    Pager1 = 0

End If

End Sub

Private Sub Menu_1_FL_Explorer_Click()
Shell "explorer W:\CALLERID\CI Logs", vbNormalFocus
End Sub

Private Sub Menu_1_CID4_Click()
Shell "notepad W:\CALLERID\CI Logs\CID4.txt", vbNormalFocus
End Sub

Private Sub Menu_1_Cid4TimeUpdated_Click()
Shell "notepad W:\CALLERID\CI Logs\Cid4TimeUpdated.txt", vbNormalFocus
End Sub

Private Sub Menu_1_CIDUD5OD_Click()
Shell "notepad W:\CALLERID\CI Logs\CIDUD5OD.txt", vbNormalFocus
End Sub

Private Sub Menu_1_CIDUD6_Click()
Shell "notepad W:\CALLERID\CI Logs\CIDUD6.txt", vbNormalFocus
End Sub

Private Sub Menu_1_CIDud6RingCountLog_Click()
Shell "notepad W:\CALLERID\CI Logs\CIDud6RingCountLog.txt", vbNormalFocus
End Sub

Private Sub Menu_1_CILog_Click()
Shell "notepad W:\CALLERID\CI Logs\CILog.txt", vbNormalFocus
End Sub

Private Sub Menu_1_TelV2_Click()
Shell "notepad W:\CALLERID\CI Logs\TelV2.txt", vbNormalFocus
End Sub

Private Sub Menu_1_TOXIC_Click()
Shell "notepad W:\CALLERID\CI Logs\TOXIC.txt", vbNormalFocus
End Sub

Private Sub Menu_BlackList_Click()
Shell "notepad W:\CALLERID\CI Logs\CIBlacklist.txt", vbNormalFocus

End Sub

Private Sub Menu_BlackListInfo_Click()
'Blacklist1
MsgBox ("0870 Blocked" + vbCr + "0845 Blocked" + vbCr + "Unblocked if Called on the 2nd Time on Trot" + vbCr + "2nd Time on trot counter reset to Zero after 15 Mins" + vbCr + "(Spam)'s Blocked" + vbCr + "(Not Spam)'s UnBlocked")

End Sub

Private Sub Menu_DVDXCOPY_Click()
Shell "H:\Program Files\321Studios\DVDXCopy\DVDXCopy.exe", vbNormalFocus

End Sub

Private Sub Menu_Ebay_Click()
Shell "explorer http://www.ebay.co.uk/", vbNormalFocus

End Sub

Private Sub Menu_EntriesOfLogs_Click()
XaGG1 = False
Ciform2.Visible = 0
Ciform2.Enabled = 0
Ciform4.Visible = 0
Ciform4.Enabled = 0
Ciform1.Visible = True
Ciform1.Enabled = True
Ciform1.Text2.SetFocus
End Sub

Private Sub Menu_HighScoresEntries_Click()
XaGG1 = False
Ciform1.Visible = 0
Ciform1.Enabled = 0
Ciform2.Visible = 0
Ciform2.Enabled = 0
Ciform4.Visible = True
Ciform4.Enabled = True
Ciform4.Text2.SetFocus

End Sub

Private Sub Menu_IceView_Click()
Shell "H:\Program Files\IceView\IceView.exe", vbNormalFocus
End Sub

Private Sub Menu_Links_Click()
Shell "explorer H:\My Documents\My Webs\MatthewLan.Com Web\Temp\bookmark_sec.htm", vbNormalFocus

End Sub

Private Sub Menu_MainTelephoneEntries_Click()
XaGG1 = False
Ciform1.Visible = 0
Ciform1.Enabled = 0
Ciform4.Visible = 0
Ciform4.Enabled = 0
Ciform2.Visible = True
Ciform2.Enabled = True
Ciform2.Text2.SetFocus

End Sub

Private Sub Menu_MenuInfo_Click()
Call MenuMode_click
End Sub
Private Sub Menu_Date1991_Click()
Shell "C:\VB6\VB-NT\Date1991\date1991.exe", vbNormalFocus
End Sub
Private Sub Menu_Date1991Web_Click()
Shell "C:\VB6\VB-NT\Date1991\date1991.exe Web", vbNormalFocus
End Sub

Private Sub Menu_Mondaylotto_Click()
Shell "explorer https://www.playmonday.com/CF/coupons.do"
End Sub

Private Sub Menu_PhotoThumb_Click()
Shell "H:\Program Files\PhotoThumb\PhotoThumb.exe", vbNormalFocus
End Sub

Private Sub Menu_Pulse_Click()
Shell "C:\VB6\VB-NT\pulse\pulse.exe", vbNormalFocus
End Sub
Private Sub Menu_Alcohol120_Click()
Shell "H:\Program Files\Alcohol Soft\Alcohol 120\Alcohol.exe", vbNormalFocus
End Sub

Private Sub Menu_AttachmentExtractor_Click()
Shell "C:\VB6\VB-NT\Attachment Extractor 672673312002\Attachment Extractor.exe", vbNormalFocus
End Sub

Private Sub Menu_CreativePlayCenter_Click()
Shell "H:\Program Files\Creative\SBAudigy\PlayCenter2\CTPlay2.exe", vbNormalFocus
End Sub

Private Sub Menu_KeyboardM_Click()
'Call Label6Hit
End Sub

Private Sub Menu_Mp3Tags_Click()
Shell "C:\VB6\VB-NT\MP3 Tags\MP3 Tags.exe", vbNormalFocus
End Sub


Private Sub Menu_StartLogs_Click()
Load CiScratch
CiScratch.Show
CiScratch.Left = DIALER.Left - CiScratch.Width
CiScratch.Top = CiStats.Top - CiScratch.Height
End Sub

Private Sub MenuMode_click()


If CiMessage.Visible = True Then
    CiMessage.Visible = False
    'Ci.List1.SetFocus
    Exit Sub
End If

CiMessage.Visible = True
'Ci.List1.SetFocus

End Sub




Public Sub Record_Number_text()
a = Ci.List1.ListIndex - 1
If a < 1 Then B = 0: GoTo jump3
b1$ = "컴컴컴컴좔컴컴컴컨컴컴컴컴컴컴컴컴좔컴컨컴컴좔컴컴컴컴컴컴컴컴"
b2$ = "컴컴컴컴쩡컴컴컴컫컴컴컴컴컴컴컴컴쩡컴컫컴컴쩡컴컴컴컴컴컴컴컴"
B = 0
If Ci.List1.List(a + 1) = b1$ Then GoTo jump3
If Ci.List1.List(a + 1) = b2$ Then GoTo jump3
If Ci.List1.List(a + 2) = b2$ Then GoTo jump3
For r = a To -1 Step -1
    B = B + 1
    If Ci.List1.List(r) = b2$ Then Exit For
Next
jump3:
A4$ = LTrim(Str(B))
Ci.Rec_no_label2.Caption = A4$
End Sub

Private Sub Ring_Indicator()

If FirstRun2 = 0 Then Exit Sub

Call CiMessage.Info_Command_click

End Sub



Private Sub SubOffTimer2()

If Tk <> 2 And (YuBt < We2 And YuBt > -1) Then

    Purple = 1

    If InStr(Find$, "NEW NUMBER") Then
        If Val(Counter$) = 1 Or Val(Counter$) = 2 Then
            Pager1 = Now + TimeSerial(0, 1, 0)
            Purple = 3
        End If
    End If

    If RTrim(Number$) = "07901893477" Then
        Purple = 1
        EricT68 = 1
        Pager1 = Now + TimeSerial(0, 0, 20)
    End If

    'If RTrim(Number$) = "07919180203" Then
    '    Purple = 1
    '    Pager1 = Now + TimeSerial(0, 0, 20)
    'End If

    If RTrim(Number$) = "XXX0 0XXX" Then
        Purple = 3
        Pager1 = Now + TimeSerial(0, 0, 20)
    End If

    If RTrim(Number$) = "WITHHELD" Then
        Pager1 = 0: Purple = 0

    End If

End If

'------------------ dans   ="07876266954"
'------------------ dave   ="01273324195"
'------------------ micheal="01273736459"


End Sub


Private Sub Menu_TestTrig_Click()
Aok = 1
Aok2 = 1
Call DIALER.Timer1_Timer
Aok2 = 0
End Sub

Private Sub Mnu_Dialer_Click()
DIALER.Show
End Sub

Sub LogFuzz()

f1 = FreeFile
Open "W:\CALLERID\CI Logs\000 Cistats_Log.txt" For Output As #f1
f2 = FreeFile
Open "W:\CALLERID\CI Logs\000 Cistats_Log_Fuzz.txt" For Output As #f2
For r = 1 To CiStats.List1.ListCount - 1
rg$ = CiStats.List1.List(r)

gg1$ = Mid$(rg$, 10, 2)
gg2$ = Mid$(rg$, 13, 2)
gg3$ = Mid$(rg$, 16, 2)
gg4$ = Mid$(rg$, 12, 1) + Mid$(rg$, 15, 1)
'If Trim(gg2$) <> "" Then Stop
If InStr("0123456789", Mid$(rg$, 1, 1)) Then

rg2$ = Mid$(rg$, 1, 9)
ff5$ = Format$(TimeSerial(Val(gg1$), Val(gg2$), Val(gg3$)), "hh:mm:ss")
rg2$ = rg2$ + " -- " + ff5$ + " --" + Mid$(rg$, 18)
rg$ = rg2$
End If

Print #f1, rg$
rt1 = InStr(rg$, "Out")
If rt1 = 0 Then rt1 = InStr(rg$, "In")
If rt1 > 0 Then rt2 = InStr(rt1, rg$, ">>")
fuzz$ = rg$
If rt1 > 0 Then
If rt2 = 0 Then rt2 = Len(rg$)
fuzz$ = Mid$(rg$, 1, rt1 + 2) + " --Fuzzed Number-- " + Mid$(rg$, rt2)
End If
Print #f2, fuzz$
Next
Close #f1
Close #f2

End Sub

Private Sub Mnu_OutPut_CiStatsLog_Click()

Call LogFuzz
Shell "Notepad W:\CALLERID\CI Logs\000 Cistats_Log.txt", vbNormalFocus

End Sub

Private Sub Mnu_OutPut_CiStatsLogFuzz_Click()

Call LogFuzz
Shell "Notepad W:\CALLERID\CI Logs\000 Cistats_Log_Fuzz.txt", vbNormalFocus

End Sub

Private Sub Mnu_RRoids_Click()
'Time_Label.Caption = "RRR Note"
tt$ = "c:\windows\system32\notepad.exe C:\Documents and Settings\Matt Lan\My Documents\01 Palm HotSync\WN-R-Words.txt"
'TY = cProcesses.Convert(Pid, hwnd3, TT$)
'TY = InstanceToWnd(TY, "WN-R-Words.txt")

'If TY = False Then Shell TT$, vbMinimizedNoFocus
Shell tt$, vbNormalFocus


End Sub

Private Sub Mnu_Std_Click()
Shell "explorer http://informatics.open.ac.uk/std"
End Sub

Private Sub Mnu_Winamp_p_Click()
    frmSender.txtMsg.Text = "winamp play"
    Call frmSender.cmdSend_Click
End Sub

Private Sub Mnu_Winamp_s_Click()
    frmSender.txtMsg.Text = "winamp stop"
    Call frmSender.cmdSend_Click
End Sub

'private sub  Ring11_label_click()
'msgbox "explain"
'8+- = both values the same
'+dialer
'> in progress
'end sub


Private Sub Stats_Command_Click()

'temp
CiStats.Show

Call DIALER.DialerTop

CiStats.List1.ListIndex = CiStats.List1.ListCount - 1

End Sub

Private Sub Timer1_Timer()


Timer1Vib = 1

'combuf modem think

If ComBuf <> 0 Then Exit Sub

If Lsd4 < Now And Lsd4 > 0 Then
    Lsd4 = 0: lsd4lock = 1
End If

If Lsd4 = 0 And lsd4lock = 1 And Ci.Enabled = True And Ci.List1.Visible = -1 Then
    Call Ring_Indicator
    lsd4lock = 0
End If

Dim dud5 As Integer
Dim Y As Double

PagerQuit = 4

If PagerCounter > 16 Or CallPagerCounter >= PagerQuit - 1 And Pager1 > 0 And Now > Pager1 Then
    d = 0

    If PagerCounter > 16 Then d = 1

    Pager1 = 0: PagerCounter = 0
    cidud5od% = FreeFile

    Open "W:\CALLERID\CI LOGS\ciDUD5od.txt" For Append As cidud5od%
    If d = 1 Then Print #cidud5od%, "-No more Retries for pager for this number--"
    If d = 0 Then Print #cidud5od%, "-No more Page's for this number--"
    Close #cidud5od%

    Purple = 0
End If


If Pager1 > 0 And Now > Pager1 And CallPagerCounter < PagerQuit Then
    
    BUSY.Visible = -1
    PagerCounter = PagerCounter + 1
    Pager1 = 0
    
    DoEvents
    
    Special = 0
    
    Dragon$ = "Dial Out"
       
    Pager1 = Now + TimeSerial(0, 1, 0)

    If EricT68 = 1 Then
    If Ci.MSComm1.PortOpen = True Then _
        Ci.MSComm1.Output = "ATX2DT 1470,079018934777,,,," + vbCr
        EricT68 = 0
    End If

    If EricT68 = 0 Then
        If Ci.MSComm1.PortOpen = True Then _
        Ci.MSComm1.Output = "ATX2DT 1470,079191802037,,,," + vbCr
    End If
    
    DoEvents

    ComBuf = 2
    
    Call Ci.Ring13_14(10)
   
End If

If ComBuf <> 0 Then Exit Sub

On Local Error Resume Next


If FirstRun0 = 1 Then
    
    MattD$ = ""

    Do
        DoEvents
        Call MComm1_OnComm
        MattD$ = m_Buffer
        If MattD$ <> "" Then Sleep 500
        aw = Ci.MSComm1.InBufferCount
    Loop Until aw = 0


    If MattD$ <> "" Then
        
        On Local Error GoTo 0

        MattD$ = RTrim(LTrim(MattD$))
        
        For rop = 1 To Len(MattD$)
            If Mid$(MattD$, rop, 1) = Chr$(&HA) Then Mid$(MattD$, rop, 1) = "-"
            If Mid$(MattD$, rop, 1) = Chr$(&HD) Then Mid$(MattD$, rop, 1) = "-"
        Next


        'If MattD$ <> "" Then

        If InStr(MattD$, "RING") Then
            Special = 0
            Ci.MSComm1.InBufferCount = 0
            m_Buffer = ""
            
            If Lsd4 > 0 Then
                Ring5 = Ring5 + 1
            End If

            If Lsd4 = 0 Then
                Ring5 = 1
            End If
            
            Lsd4 = Now + TimeSerial(0, 0, 10)
            Lsd44 = Now + TimeSerial(0, 0, 10)
            
            'CiMessage.Rings_label(7).Caption = LTrim(Str(Ring5))
            'Ci.Ring11.Caption = LTrim(Str(Ring5))
        
        
        End If

        cidud5od% = FreeFile
        Open "W:\CALLERID\CI LOGS\ciDUD5od.txt" For Append As cidud5od%
        cidud6% = FreeFile
        Open "W:\CALLERID\CI LOGS\ciDUD6.txt" For Append As cidud6%
        jumpme = 0

        'If InStr(MattD$, "NO DIALTONE") Or InStr(MattD$, "NO CARRIER") Then

        '    Ci.MSComm1.InBufferCount = 0
        '    ToolChest = 0
        '    'NoCheck2 = 0
        'End If





        If InStr(MattD$, "OK") Then

            Dragon$ = "Ok"
            Nutty = 0
            'ToolChest = 0
            'NoCheck2 = 0
            
        End If

        If (MattD$ = "-" Or InStr(MattD$, "RING") Or InStr(MattD$, "OK")) And Len(MattD$) < 29 Then

            m_Buffer = ""
            Print #cidud5od%, MattD$; String$(29 - Len(MattD$), "+"); Date$; "-"; Time$; "-"; Format$(Timer, "00000.00"); "-"; Number$; " "; Format$(Ring5, "00") '; " "; HobNob$
                    
            If InStr(MattD$, "RING") Then
                Print #cidud6%, MattD$; String$(29 - Len(MattD$), "+"); Date$; "-"; Time$; "-"; Format$(Timer, "00000.00"); "-"; Number$; " "; " "; Format$(Ring5, "00") '; " "; HobNob$
            End If
        
            jumpme = 1
        End If

        If jumpme = 0 Then
            m_Buffer = ""
            Print #cidud5od%, MattD$; String$(10, "+"); Date$; "-"; Time$; "-"; Format$(Timer, "00000.00"); "-"; Number$; " "; " "; Format$(Ring5, "00")   '; " "; HobNob$
        End If

        Close cidud6%
        Close cidud5od%

        jumpme = 0
        
        If InStr(MattD$, "NO CARRIER") Then
            Dragon$ = "No Carrier"
            Call Ci.Ring13_14(30)
        End If

        If InStr(MattD$, "NO DIALTONE") Then
            Dragon$ = "No Dialtone"
            Call Ci.Ring13_14(30)
        End If
        
        If InStr(MattD$, "BUSY") Then
            Dragon$ = "No Dialtone"
            Call Ci.Ring13_14(30)
        End If
    End If
End If

'MattD$ = Ar3$
'            If InStr(MattD$, "BUSY") Then
'                Dragon$ = "No Dialtone"
'                Call Ci.Ring13_14(30)
'            End If
'MattD$ = ""

If FirstRun2 = 1 Then Ci.BUSY.Visible = 0

Exit Sub

errors:
er = 0
Resume Next

End Sub


'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------
'--------------------------------------------------------


Public Sub AddNumberToCIList1()

'put black listing in here, this subroutine

'If NumberStringDisMiss < Now And NumberStringDisMiss > 0 Then
'    NumberStringDisMiss = 0
'    Number$ = "*-*-*-\/\/-*-*-*"
'End If
       
If FirstRun0 = 0 Then
    FirstRun0 = 1
    Clock = 1
End If

If Tk > 0 Then

    BUSY.Visible = -1
    Ci.List1.BackColor = Red1
            
    Call IncomingCall(ICResult, P1, P2, P3, Q$, From2$, Tnow, Number$, Word$)



    If ICResult = True Or Tk = 0 Then
        If Ciform1.Visible = -1 Then Call ExitCiForm1
        If Ciform2.Visible = -1 Then Call ExitCiForm2
        If Ciform4.Visible = -1 Then Call ExitCiForm4
    End If


    If ICResult = True Or Tk > 0 Then
            

        ' tk 1 detect a back door trig no.
        If Tk = 1 Then
            Tdate$ = Mid$(Date$, 4, 2) + "/" + Mid$(Date$, 1, 2) + "/" + Mid$(Date$, 9)
            Mdate$ = Mid$(Tdate$, 1, 6) + Mid$(Date$, 7)

            Ttime$ = Time$
            Number$ = "XXX0 0XXX       "
            Code$ = "01"

            'number$ = "WITHHELD        "
            'number$ = "PAYPHONE        "
            'number$ = "01273-885142    "
            'number$ = "01273-421969    "
            'number$ = "01273-775400    "
            'number$ = "01273-595992    "
            'number$ = "01273123456     "
            'number$ = "08701558822     "
            'number$ = "01273413291     "
            'number$ = "01273555555     "
            'code$ = "00"

            Counter$ = "0000"

            'word$ = "RINGBACK"
            'code$ = "02"

            Word$ = ""
            'code$ = "00"

            'word$ = "PAYPHONE"
            'code$ = "01"

            Call FastFind


            'find$
            'code$ = "01"
            'code$ = "00"

            Word$ = Mid$(Find$, 1, 4)
            Counter$ = Format$(Val(Counter$) + 1, "0000")
            RSet Counter$ = LTrim(Str(Val(Counter$)))
        End If

        ' tk = 2 detect a front door number from script file.
        ' tk = 2 and add item mode from ciedit.
        ' date format should be day/month/year
        If Tk = 2 Then
            Tk2 = 1
            Mdate$ = Mid$(Slop$, 1, 10)
            Tdate$ = Mid$(Slop$, 1, 6) + Mid$(Slop$, 9, 2)
            Ttime$ = Mid$(Slop$, 12, 8)
            number2$ = Mid$(Slop$, 21, 16)
            Number$ = "                "
            LSet Number$ = number2$
            Word$ = Mid$(Slop$, 38)
            If Word$ = "" Then
                Word$ = "    "
            End If

            If Val(Code$) < 1 Then Code$ = "00"

            ' in add item mode puts code 01 if you enter a description

            If RTrim(Word$) <> "" Then Code$ = "01"
            If InStr(UCase$(Word$), "RINGBACK") Then Code$ = "02"
            If InStr(UCase$(Word$), "RINGBACK") Then Code$ = "02"
            Counter$ = "0000"
            Call FastFind
            If Code$ = "01" Then Word$ = Mid$(Find$, 1, 4)
            If Code$ = "02" Then Word$ = Mid$(Find$, 1, 4)
            If InStr(Find$, RTrim(Number$)) Then Word$ = Mid$(Find$, 1, 4)
            Counter$ = Format$(Val(Counter$) + 1, "0000")
            RSet Counter$ = LTrim(Str(Val(Counter$)))
        End If







        ' P2 > 0 Detect a test number

        If P2 > 0 Then
            Mdate$ = Mid$(Date$, 4, 2) + "/" + Mid$(Date$, 1, 2) + "/" + Mid$(Date$, 7)
            Tdate$ = Mid$(Date$, 4, 2) + "/" + Mid$(Date$, 1, 2) + "/" + Mid$(Date$, 9)
            Ttime$ = Time$
            Number$ = "1234567890      "
            Code$ = "01"
            Counter$ = "0000"
            Word$ = "TEST"
            Call FastFind
            Counter$ = Format$(Val(Counter$) + 1, "0000")
            RSet Counter$ = LTrim(Str(Val(Counter$)))
        End If

            
        ' any other number not a test
            
        If P1 > 0 Or P3 > 0 Then
        
            Tdate$ = Format$(Tnow, "DD/MM/YY")
            Mdate$ = Format$(Tnow, "DD/MM/YYYY")
            Ttime$ = Format$(Tnow, "HH:MM:SS")
            
            gup = 0
            er = 0

            'HobNob$ = "-----"
                
            'HobNob$ = "Error"
           
    
            If Mid$(Tdate$, 3, 1) <> "/" Or Mid$(Tdate$, 6, 1) <> "/" Then Tdate$ = Mid$(Date$, 4, 2) + "/" + Mid$(Date$, 1, 2) + "/" + Mid$(Date$, 9)

            Code$ = "00"
            Counter$ = "0000"

            PK2 = InStr(Q$, Chr$(7))
            If PK2 > 0 Then Code$ = "01"
            If P3 > 0 Then Code$ = "02"
                
            If InStr(Number$, "XXX0 0XXX") Then Code$ = "01"
                
            
            '' the one to test for wrong finds on number
            Call FastFind
                
            If Code$ = "01" Then Word$ = Mid$(Find$, 1, 4)
            If Code$ = "02" Then Word$ = Mid$(Find$, 1, 4)
            
            If Mid$(Number$, 1, 2) = "00" Then
                Code$ = "01": Word$ = "INTL"
            End If
                
            Counter$ = Format$(Val(Counter$) + 1, "0000")
            RSet Counter$ = LTrim(Str(Val(Counter$)))

        End If
  




        If ICResult = False Or Tk > 0 Then


            cidud5od% = FreeFile
            Open "W:\CALLERID\CI LOGS\ciDUD5od.txt" For Append As cidud5od%
            cidud6% = FreeFile
            Open "W:\CALLERID\CI LOGS\ciDUD6.txt" For Append As cidud6%
            Print #cidud5od%, "Start Of New Number...";
            Print #cidud6%, "Start Of New Number...";
            If Tk > 0 Then Print #cidud6%, "Simulated number by trig button or Number from back door "; Tdate$; "-"; Ttime$; "-"; Format$(Timer, "00000.00"); "-"; Number$
            If P1 > 0 Then Print #cidud6%, "Normal                                                   "; Tdate$; "-"; Ttime$; "-"; Format$(Timer, "00000.00"); "-"; Number$
            If P2 > 0 Then Print #cidud6%, "Test                                                     "; Tdate$; "-"; Ttime$; "-"; Format$(Timer, "00000.00"); "-"; Number$
            If P3 > 0 Then Print #cidud6%, "Ringback                                                 "; Tdate$; "-"; Ttime$; "-"; Format$(Timer, "00000.00"); "-"; Number$
            If Tk > 0 Then Print #cidud5od%, "Simulated number by trig button or Number from back door "; Tdate$; "-"; Ttime$; "-"; Format$(Timer, "00000.00"); "-"; Number$
            If P1 > 0 Then Print #cidud5od%, "Normal                                                   "; Tdate$; "-"; Ttime$; "-"; Format$(Timer, "00000.00"); "-"; Number$
            If P2 > 0 Then Print #cidud5od%, "Test                                                     "; Tdate$; "-"; Ttime$; "-"; Format$(Timer, "00000.00"); "-"; Number$
            If P3 > 0 Then Print #cidud5od%, "Ringback                                                 "; Tdate$; "-"; Ttime$; "-"; Format$(Timer, "00000.00"); "-"; Number$

            Close cidud6%
            Close cidud5od%

            If ICResult = True Then
                CallInCheck = Now + TimeSerial(0, 0, 30)
                Lsd4 = 0
                If Pager1 > 0 Then Pager1 = Now + TimeSerial(0, 1, 0)
            End If
        End If
            


        'NumberStringDisMiss = Now + TimeSerial(0, 0, 30)
            
        Ring5 = 0
            
        Total_no_calls_label2.Caption = LTrim(Str(Val(Total_no_calls_label2.Caption) + 1))

        If TimeHasBeenUpdated = 1 Then rh$ = Chr$(254) Else rh$ = "�"
            
        If whop$ = Chr$(254) Then rh$ = Chr$(254)

        'If InStr(Number$, "595992") Then Stop
        
        
        numberx$ = Number$
        If Len(Number$) > 8 And InStr("0123456789", Mid$(Number$, 1, 1)) > 0 Then
            Number$ = Mid$(Number$, 1, 5) + "-" + Mid$(Number$, 6)
            Number$ = Left$(Number$, Len(Number$) - 1)
        End If
'        Number$ = Mid$(Number$, 1, 5) + "-" + Mid$(Number$, 6)
'        Number$ = numberx$
        
        Ci.List1.AddItem Tdate$ + "�" + Ttime$ + rh$ + Number$ + "�" + Mid$(Word$, 1, 4) + "�" + Counter$ + "�" + Find$, 2
        Number$ = numberx$

        Counter$ = Format$(Val(Counter$), "0000")
        Tdate$ = Mid$(Tdate$, 1, 6) + Mid$(Date$, 7)

        If Tk = 2 Then Tdate$ = Mdate$

        If Ciform1.List2.ListCount > RecList Then
            Ciform1.List2.RemoveItem Ciform1.List2.ListCount - 1
        End If

            
        'GoTo jumpkill1
        If YuBt > -1 Then
            If YuBt < Ciform2.List2.ListCount Then
                'Ciform1.List2.AddItem Ciform2.List2.List(YuBt), 0
                phub$ = Number$ + Tdate$ + " " + Ttime$ + " " + Counter$ + " " + Code$ + " " + Find$
                Ciform1.List2.AddItem phub$, 0

                If Tdate$ <> Format$(Now, "dd/mm/yyyy") Then
                    freeme = FreeFile
                    Open "W:\CALLERID\CI LOGS\cierror.txt" For Append As #freeme
                    Print #freeme, Tdate$, , Format$(Now, "dd/mm/yyyy")
                    Close #freeme
                End If

            Call FastFind
                If InStr(Ciform2.List2.List(YuBt), Number$) = 0 Then
                
                    MsgBox "Caller Id Dupe # Sort"
                    Stop
                
                End If
                '----------------------------------------
                '----------------------------------------
                '----------------------------------------
                '----------------------------------------
                '----------------------------------------
                '----------------------------------------
                '----------------------------------------
                
                If Trim(Mid$(Ciform2.List2.List(YuBt), 1, 16)) <> Trim(Number$) Then
                    MsgBox "Caller Id Dupe # Sort"
                    
                    Stop
                End If
                
                Ciform2.List2.List(YuBt) = Number$ + Tdate$ + " " + Ttime$ + " " + Counter$ + " " + Code$ + " " + Find$
                '----------------------------------------
                '----------------------------------------
                '----------------------------------------
                '----------------------------------------
                '----------------------------------------
                '----------------------------------------
                '----------------------------------------
            End If
        End If


        If YuBt = -1 Then

            'Ciform1.List2.AddItem Number$ + "00/00/1900 00:00:00 0000 " + Code$ + " " + Find$, 0
                
            phub$ = Number$ + Tdate$ + " " + Ttime$ + " " + Counter$ + " " + Code$ + " " + Find$
            Ciform1.List2.AddItem phub$, 0
            Ciform2.List2.AddItem phub$, 0
                
            We2 = We2 + 1

            Total_callers_label2.Caption = LTrim(Str(Val(Total_callers_label2.Caption) + 1))

        End If

        TDate3$ = Mid$(Ciform1.List2.List(0), TdatePos, 10)
        TTime3$ = Mid$(Ciform1.List2.List(0), TTimePos, 8)
        Counter3$ = Mid$(Ciform1.List2.List(0), CountPos, 4)


        'blacklist1----------------
       
        
        If Blacklist4 < Now Then Blacklist3 = 0
        
        If Blacklist2$ = Number$ Then
            Blacklist3 = Blacklist3 + 1
            Blacklist4 = Now + TimeSerial(0, 15, 0)
        Else
            Blacklist3 = 0
        End If
        
        
        If Mid$(Number$, 1, 4) = "0870" And Blacklist3 <> 1 Then
            BlackList1 = 1
            'Zenith, A&L
        End If
        
        If Mid$(Number$, 1, 4) = "0845" And Blacklist3 <> 1 Then
            BlackList1 = 1
            'Zenith, A&L
        End If
        
        Blacklist2$ = Number$
        
        If InStr(UCase$(Find$), "(SPAM)") Then
            BlackList1 = 1
        End If
        
        If InStr(UCase$(Find$), "(NOT SPAM)") Then
            BlackList1 = 0
        End If



        If Tk = 2 Or (YuBt >= We2) Then
            If RTrim(Number$) = "WITHHELD" Then
                Pager1 = 0
                'PICK UP EVERY fifth WITHHELD then every second after
                'withheld = withheld + 1
                'If withheld = 6 Then withheld = 4
            End If
        End If

        'my first mobile
        '07799-357496

        Call SubOffTimer2


        If CiMessage.Rings_label(10).Caption = "Off" Then
            'pager subroutine
            Pager1 = 0
        End If

        'over ride pager off'''put white listing in here
        'after the pager1 gets set to zero thus
        'undoing & overrideing & taking prioity to the
        'pager on off switch
        'WAIT 30 SECONDS THEN ALERT MOBILE THESE
        'NUMBERS CALLED
    
        If InStr(Find$, "NEW NUMBER") Then
            If Counter$ = "0001" Then
                Pager1 = Now + TimeSerial(0, 1, 0)
                Purple = 3
                Call Ci.New_Number_Counter
                RingTextBack$ = ""
                If InStr(Find$, "RINGBACK") Then
                    RingTextBack$ = "Ring Back. "
                End If
                'Ci.DirectSS1.speak RingTextBack$ + "New Number First Time Caller"
                Call Ci.SpeakV(RingTextBack$ + "New Number First Time Caller")
            End If
        End If

        'let rip numbers but still page to mobile....
        'WAIT 30 SECONDS THEN ALERT MOBILE THESE
        'NUMBERS CALLED
        If RTrim(Number$) = "01273419914" Then
            Pager1 = Now + TimeSerial(0, 1, 0): Purple = 2
        End If

        'Ltd eds
        If RTrim(Number$) = "01273571373" Then
            Pager1 = Now + TimeSerial(0, 1, 0): Purple = 2
        End If

        'kevin light
        If RTrim(Number$) = "01273732733" Then
            Pager1 = Now + TimeSerial(0, 1, 0): Purple = 2
        End If


        'no we have no override pager off if pager
        'control box says off
        If CiMessage.Rings_label(10).Caption = "Off" Then
            'pager subroutine
            Pager1 = 0
            Purple = 0
        End If
        
        
        Call HallFameSub
        
        Counter$ = Format$(Val(Counter$), "0000")

        cilog = FreeFile
        Open "W:\CALLERID\CI LOGS\cilog.txt" For Append As #cilog
        Print #cilog, Number$ + Mdate$ + " " + Ttime$ + " " + Counter$ + " " + Code$ + " " + Find$
        Close #cilog


        'do not backup when streaming a list of numbers in...
        If Tk <> 2 Then
            Call Ci.BackUp
        End If
        
        Ci.List1.ListIndex = 0
        
        If Number$ = Pagernumber$ Then
            CallPagerCounter = CallPagerCounter + 1
        Else
            CallPagerCounter = 0
        End If

        Pagernumber$ = Number$
    End If
End If

If ICResult = False And Tk = 0 And Q$ <> "" Then
    cidud = FreeFile
    Open "W:\CALLERID\CI LOGS\ciDUD.txt" For Append As #cidud
    If Number$ = "" Then Number$ = Space$(16)
    Print #cidud, Q$
    Print #cidud, "***"; Number$; "-"; Date$; "-"; Time$; "-"; Timer
    Close #cidud
End If

CiMessage.Label7.Caption = "Records Keeping" + Str$(Ciform1.List2.ListCount) + " Out of" + Str(RecList) + " Total"
Ci.LabelRecs.Caption = Str$(Ciform1.List2.ListCount) + " of" + Str(RecList)


If Ci.List1.BackColor = Red1 Then List1.BackColor = Blue1

Tk = 0

If Purple > 0 And Haze = 0 Then Haze = Purple

End Sub

Private Sub Timer2_Timer()

End Sub

Public Sub Timer3_Timer()

a$ = Format$(Now, "h:mm:ssa/p")
B$ = Format$(Now, "ddd d/mmm/yyyy")
C$ = B$ + " " + a$

Ci.datetime.Text = C$
   
'If FirstRun2 = 1 And lsd44 = 0 Then lsd44 = Now + TimeSerial(0, 0, 3)
   
If Lsd44 < Now And Lsd44 > 0 Then
    Lsd44 = 0
    '--- OK for the record once an for all then pluses +++ are
    '--- Collect on startup of dialer to save time
    '--- its easyier to note cidud6.log in here as thats
    '--- a big file so they are the ---
    '--- Just find a way of stop the duplicates later
    '--- You dont really need to write whole list just
    '--- append it sorts anyway
    wedss$ = Format$(DingALag, "0.0000000000") + Str$(DingALag) + " " + LastNumber$ + " " + Format$(Ring5, "000")
    If wedss$ <> Wassar1$ Then
        If DingALag > 0 Then CiScratch.List2.AddItem wedss$ + " --*"
        
        For rews = 1 To CiScratch.List2.ListCount - 1
            If rews > 1 Then easye$ = CiScratch.List2.List(rews - 2)
            easyg$ = CiScratch.List2.List(rews - 1)
            easyh$ = CiScratch.List2.List(rews)
            If easyh$ = "" Then Exit For
            If Mid$(easyh$, 1, 54) = Mid$(easyg$, 1, 54) And _
                Mid$(easyh$, Len(easyh$) - 1, 1) = Mid$(easyg$, Len(easyg$) - 1, 1) Then
                CiScratch.List2.RemoveItem (rews - 1)
            End If
        
            If rews > 1 Then
                If Mid$(easyh$, 1, 54) = Mid$(easye$, 1, 54) And _
                    Mid$(easyh$, Len(easyh$) - 1, 1) = Mid$(easye$, Len(easye$) - 1, 1) Then
                        CiScratch.List2.RemoveItem (rews - 2)
                End If
            End If
        Next
        
        cidud7% = FreeFile
        Open "W:\CALLERID\CI LOGS\cidud6RingCountLog.txt" For Output As #cidud7%
        For rews = 0 To CiScratch.List2.ListCount - 1
            'Print #cidud7%, Mid$(CiScratch.List2.List(rews), 1, Len(CiScratch.List2.List(rews)) - 4)
            Print #cidud7%, CiScratch.List2.List(rews)
        Next
        Close #cidud7%
        Call CiMessage.Info_Command_click
    End If
    Wassar1$ = wedss$
End If


If Lsd55 < Now And Lsd55 > 0 Then
    Lsd55 = 0
    wedss$ = Format$(DingALag, "0.0000000000") + Str$(DingALag) + " " + LastNumber2$ + " " + Format$(DialerRingCountPlus2 - 1, "000")
    If wedss$ <> Wassar2$ Then
        CiScratch.List2.AddItem wedss$ + " ++*"
        
        For rews = 1 To CiScratch.List2.ListCount - 1
            If rews > 1 Then easye$ = CiScratch.List2.List(rews - 2)
            easyg$ = CiScratch.List2.List(rews - 1)
            easyh$ = CiScratch.List2.List(rews)
            If easyh$ = "" Then Exit For
            If Mid$(easyh$, 1, 54) = Mid$(easyg$, 1, 54) And _
                Mid$(easyh$, Len(easyh$) - 1, 1) = Mid$(easyg$, Len(easyg$) - 1, 1) Then
                CiScratch.List2.RemoveItem (rews - 1)
            End If
        
            If rews > 1 Then
                If Mid$(easyh$, 1, 54) = Mid$(easye$, 1, 54) And _
                    Mid$(easyh$, Len(easyh$) - 1, 1) = Mid$(easye$, Len(easye$) - 1, 1) Then
                        CiScratch.List2.RemoveItem (rews - 2)
                End If
            End If
        Next
        
        cidud7% = FreeFile
        Open "W:\CALLERID\CI LOGS\cidud6RingCountLog.txt" For Output As #cidud7%
        For rews = 0 To CiScratch.List2.ListCount - 1
            Print #cidud7%, CiScratch.List2.List(rews)
        Next
        Close #cidud7%
        DialerRingCount = 0
        DialerRingCountPlus2 = 0
        Call CiMessage.Info_Command_click
    End If
    Wassar2$ = wedss$
End If



   
If FirstRun2 = 0 Then
    For r = 0 To Forms.Count - 1
        wed$ = Forms(r).Caption
        If wed$ = "CiEDIT" Then CiEDIT.datetime.Text = C$
        If wed$ = "Ciform1" Then Ciform1.datetime.Text = C$
        If wed$ = "Ciform2" Then Ciform2.datetime.Text = C$
        If wed$ = "Ciform4" Then Ciform4.datetime.Text = C$
        If wed$ = "Statistics" Then
            CiStats.datetime.Text = C$
            CiStats.Label34.Caption = Format$(Timer - Slashing_Rain, " 0.00000")
            CiStats.Label34.Refresh
            If FirstRun2 = 1 Then
                If Cutty > 0 Then
                    Label30.Caption = Format$(((Futty / Cutty) * 100), " 00.000000") + "%"
                End If
            End If
        
        End If
    Next
End If

If FirstRun2 = 1 Then
            
    If Ciform1.Visible = True Then Ciform1.datetime.Text = C$
    If Ciform2.Visible = True Then Ciform2.datetime.Text = C$
    If Ciform4.Visible = True Then Ciform4.datetime.Text = C$
    If CiEDIT.Visible = True Then CiEDIT.datetime.Text = C$
    If CiStats.Visible = True Then CiStats.datetime.Text = C$
    If CiStats.Visible = True Then
        
        'If Rebirth_Tog = True Then
        '    CiStats.Rebirth = Format$(Last_Reboot, "DDD d-m-yy h:m:ssa/p")
        
        'Else
'            Dim df As Double
'            df = Now - Last_Reboot
'            CiStats.Rebirth = Trim(Str(Int(df))) + " Days " + Trim(Str(Hour(df))) + "h " + Trim(Str(Minute(df))) + "m " + Trim(Str(Second(df))) + "s"
            'CiStats.Rebirth = SystemUptime.Days & " Days " & _
            'Format$(SystemUptime.Hours Mod 24, "00") & ":" & _
            'Format$(SystemUptime.Minutes, "00") & ":" & _
            'Format$(SystemUptime.Seconds, "00") & "." & _
            'Format$(SystemUptime.MSeconds, "000")
        
        
        'End If
    End If
End If


If FirstRun2 = 1 Then
    If Trim(CiStats.Qwerty_Label40) = "" Then
        
        frmSender.txtMsg.Text = "Hit Request " + Time$

        Call frmSender.cmdSend_Click

        If BadComms = 1 Then
            Mouse_Label.Caption = "Bad Comms"
            Qwerty_Label40.Caption = "Bad Comms"
        End If
    End If
End If



End Sub

''---In Module---'
'Private Type TimeConv
'    Days        As Long
'    Hours       As Long
'    Minutes     As Long
'    Seconds     As Long
'    MSeconds    As Long
'End Type
 
'Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
 
Sub timer11_timer()
            
If FirstRun2 = 0 Then Exit Sub

If Rebirth_Tog = False Then
    CiStats.Rebirth = Format$(Last_Reboot, "DDD d-m-yy h:m:ssa/p")
Else
    CiStats.Rebirth = SystemUptime.Days & " Days " & _
    Format$(SystemUptime.Hours Mod 24, "00") & ":" & _
    Format$(SystemUptime.Minutes, "00") & ":" & _
    Format$(SystemUptime.Seconds, "00") & "." & _
    Format$(SystemUptime.MSeconds, "000")
End If


End Sub


Private Function ConvertTime(ByVal Tick As Long) As TimeConv
 
ConvertTime.MSeconds = Tick Mod 1000
 
Tick = Tick \ 1000
ConvertTime.Days = ((Tick) \ (24 * (60 ^ 2)))
 
If ConvertTime.Days > 0 Then Tick = (Tick - 24 * (60 ^ 2)) * ConvertTime.Days
ConvertTime.Hours = Tick \ (60 ^ 2)
 
If ConvertTime.Hours > 0 Then Tick = Tick - ((60 ^ 2) * ConvertTime.Hours)
ConvertTime.Minutes = Tick \ 60
 
ConvertTime.Seconds = Tick Mod 60
 
End Function
 
 
Private Property Get SystemUptime() As TimeConv
SystemUptime = ConvertTime(GetTickCount)
End Property
 
 
'---Form Usage Example:---'
'Private Sub Timer2_Timer()
''Exit Sub
'Label4 = "System Uptime: " + vbCr & SystemUptime.Days & " Days, " & _
'        SystemUptime.Hours Mod 24 & ":" & _
'        SystemUptime.Minutes & ":" & _
'        SystemUptime.Seconds & "." & _
'        Format$(SystemUptime.MSeconds, "000") & "."
'End Sub

Private Sub timer4_timer()

'Call SeXxYy
Call DIALER.DialerTop

End Sub

Private Sub Timer5_Timer()

Call Magic

If Screen.Width <> 19200 And Screen.Height <> 15360 Then
    ScrWidth_Old = 1
End If


If Screen.Width = 19200 And Screen.Height = 15360 And ScrWidth_Old = 1 Then

    ScrWidth_Old = 0

    Call DIALER.DialerTop

End If

End Sub

Private Sub Timer6_timer()

'this was taken from magic2 bit modified
'this check Test Triggers and such Like Paging calls
'to let me know if a special number called the land
'line.

Timer6Vib = 1

Call Magic2

If ComBuf <> 2 Then Exit Sub


Call MComm1_OnComm


If m_Buffer = "" Then Exit Sub

Ar1a$ = m_Buffer
ets = 0
ets = InStr(Ar1a$, "NO CARRIER")
If ets = 0 Then ets = InStr(Ar1a$, "NO DIALTONE")
If ets = 0 Then ets = InStr(Ar1a$, "OK")
If ets = 0 Then ets = InStr(Ar1a$, "BUSY")

If ets = 0 Then Exit Sub

m_Buffer = ""

MattD$ = Ar1a$

Dragon$ = ""

If InStr(Ar1a$, "OK") Then Dragon$ = "No Carrier"
If InStr(Ar1a$, "NO CARRIER") Then Dragon$ = "No Carrier"
If InStr(Ar1a$, "NO DIALTONE") Then Dragon$ = "No Dialtone"
    
For rop = 1 To Len(MattD$)
    If Mid$(MattD$, rop, 1) = Chr$(&HA) Then Mid$(MattD$, rop, 1) = "+"
    If Mid$(MattD$, rop, 1) = Chr$(&HD) Then Mid$(MattD$, rop, 1) = "-"
Next


If Boat2$ <> Dragon$ Then

Boat2$ = Dragon$

    cidud5od% = FreeFile
    'Close: Stop
    Open "W:\CALLERID\CI LOGS\ciDUD5od.txt" For Append As cidud5od%
    mattd1$ = Mid$(MattD$, 1, ets - 1)
  
    Print #cidud5od%, mattd1$; String$(50 - Len(mattd1$), "+"); Date$; "-"; Time$; "-"; Format$(Timer, "00000.00"); "-"; Number$; " "; " "; Format$(Ring5, "00")
    
    mattd1$ = Mid$(MattD$, ets)
    wei = 50
    If Len(mattd1$) > 49 Then wei = Len(mattd1$) + 5
    
    Print #cidud5od%, mattd1$; String$(wei - Len(mattd1$), "+"); Date$; "-"; Time$; "-"; Format$(Timer, "00000.00"); "-"; Number$; " "; Format$(Ring5, "00")
    
    Close #cidud5od%
    
End If





MattD$ = mattd1$

If InStr(MattD$, "NO DIALTONE") Then

    If Purple > 0 And Purple <> Haze Then
        Purple = 0: Haze = 0
        Pager1 = 0
    End If

End If


If InStr(MattD$, "NO CARRIER") Then
    Pager1 = 0
    PagerCounter = 0
        

    If Purple >= 1 And Purple <= 3 Then
        Purple = Purple - 1
        Pager1 = Now + TimeSerial(0, 0, 5)
        If Purple = 1 Then Haze = 0
    End If
      
End If

Special = 0

ComBuf = 0

Call Ci.Ring13_14(40)

End Sub

Private Sub Timer7_Timer()
Ci.Refresh
Rebirth_Tog = Not Rebirth_Tog

If Rebirth_Tog = True Then
    Timer7.Interval = 14000
Else
    Timer7.Interval = 3000
End If

End Sub

Private Sub Timer8_Timer()

Call CiStats.StatExitCommand5


we1 = FindWindow(vbNullString, "CallerID - Microsoft Visual Basic [run]")
We2 = FindWindow(vbNullString, "CallerID - Microsoft Visual Basic [break]")

If We2 <> 0 Then Xxr1 = 0
If we > 0 And Xxr1 <> we1 Then
Xxr1 = we
'ShowWindow we, SW_MINIMIZE
'ShowWindow we, SW_SHOWMINNOACTIVE
'ShowWindow we, SW_FORCEMINIMIZE
'ShowWindow wed, SW_SHOWNORMAL
'ShowWindow we, SW_FORCEMINIMIZE
ShowWindow we, SW_SHOWMINIMIZED

End If


End Sub

Private Sub Timer9_Timer()

Timer9Vib = 1

Call CI2.Spin

If SpinnerSec <> Now Then
    Call StarWars(False)
    SpinnerSec = Now
    Exit Sub
End If


End Sub


Private Sub Timer10_Timer()

testjump = True
'testjump = False
If testjump = False Then Call StarWars(True)
If testjump = False Then Timer10.Enabled = False

End Sub

Private Sub TimerComm1_Timer()

'MsgBox "COM1: not available. Change the CommPort property to another port."

TimerComm1.Enabled = False
CiMessage.Show

'Ci.TimerStopMessage.Enabled = True
    
End Sub

Private Sub TimerStopMessage_Timer()
CiMessage.Hide
TimerStopMessage.Enabled = False

End Sub

Private Sub ViewItOne_click()

If Ci.List1.Visible = 0 Or Ci.Visible = 0 Then
    Ci.Enabled = True
    Ci.Visible = -1
    Ci.List1.Visible = -1
End If

Ciform1.Visible = -1
Ciform1.Enabled = -1
Ciform1.Text2.SetFocus
XaGG1 = True
End Sub





Public Sub Qwerty()

free5 = FreeFile
Open "W:\CALLERID\ci logs\2cidrunme.txt" For Input As #free5
wtage = 0
Do
    Line Input #free5, aed$
    
    If InStr(aed$, "Last System ReBoot Time") Then
        wtage = 1
        If qwerty2 > Qwerty2Ttt Then Qwerty2Ttt = qwerty2
        If mouse55 > Mouse55Ttt Then Mouse55Ttt = mouse55
        qwerty2 = 0
        mouse55 = 0
    End If
        
    
    If wtage = 1 Then
        If InStr(aed$, "Keyboard =") Then
            qwerty2 = qwerty2 + Val(Mid$(aed$, InStr(aed$, "Keyboard =") + 10))
        End If
        
        If InStr(aed$, "Mouse =") Then
            mouse55 = mouse55 + Val(Mid$(aed$, InStr(aed$, "Mouse =") + 7, 10))
        End If
    End If
Loop Until EOF(free5)

Close #free5


End Sub

Public Sub HitScore()

Dim gstring As Double
Dim sstring As Double

sstring = (Val(CiStats.Qwerty_Label40) / Qwerty2Ttt) * 100

If sstring >= 100 Then compute_key$ = Format$(((Val(CiStats.Qwerty_Label40) / Qwerty2Ttt) * 100), " 0") + "%K"
If sstring < 100 Then compute_key$ = Format$(((Val(CiStats.Qwerty_Label40) / Qwerty2Ttt) * 100), " 0.0") + "%K"
If sstring < 10 Then compute_key$ = Format$(((Val(CiStats.Qwerty_Label40) / Qwerty2Ttt) * 100), " 0.00") + "%K"
If sstring < 1 Then compute_key$ = Format$(((Val(CiStats.Qwerty_Label40) / Qwerty2Ttt) * 100), " 0.000") + "%K"


gstring = ((KillTheBeast * 1000) / Mouse55Ttt) * 100
If gstring < 100 Then
    compute_mos$ = Format$(gstring, " 0.0") + "%M"
Else
    compute_mos$ = Format$(gstring, " 0") + "%M"
End If

CiStats.High_Score_Label.Caption = compute_key$ + compute_mos$

End Sub


Public Sub mystery()

'hit request

If InStr(GlobalHot$, "KeyCounter") Then
    CiStats.Qwerty_Label40 = Str(Val(Mid$(GlobalHot$, 11)))
    Call HitScore
End If

If InStr(GlobalHot$, "MouseCounter") Then
    
    KillTheBeast = Val(Mid$(GlobalHot$, 13)) / 1000
    
    
    If KillTheBeast < 1000 Then hatehair$ = Format$(KillTheBeast, " 0") + "k"
    If KillTheBeast > 999 Then hatehair$ = Format$(KillTheBeast, " 0,000") + "k"
    If KillTheBeast > 999999 Then hatehair$ = Format$(KillTheBeast, " 0,000,000") + "k"
    
    CiStats.Mouse_Label = hatehair$
    Call HitScore
End If


If InStr(GlobalHot$, "Request Play") Then

    If Ring14Right.BackColor = Red1 Then
        frmSender.txtMsg.Text = "winamp play"
    Else
        frmSender.txtMsg.Text = "winamp stop"
    End If

    Call frmSender.cmdSend_Click

End If


End Sub


Public Sub Ring13_14(LuvIt As Integer)


Wes99 = Special2
Special2 = 0
    
If Dragon$ = "No Dialtone" Then
    
    Ci.Ring14Right.BackColor = Red1
    Ci.Ring14Right.ForeColor = White1
    Ci.Ring14Right.Caption = "Off"
            
    Ci.Ring13Left.BackColor = Red1
    Ci.Ring13Left.ForeColor = White1
    Ci.Ring13Left.Caption = "Off"
    Conk2 = -1
End If
            
If Dragon$ = "No Carrier" Then
        
    Ci.Ring13Left.Caption = "On"
    Ci.Ring13Left.BackColor = 0
    Ci.Ring13Left.ForeColor = White1
    Ci.Ring14Right.Caption = "On"
    Ci.Ring14Right.BackColor = 0
    Ci.Ring14Right.ForeColor = White1
    
End If
            
If Dragon$ = "Dial Out" Then
    If LuvIt = 10 Then PagerDial = 1 Else PagerDial = 0
    
    Ci.Ring13Left.Caption = "Off"
    
    If PagerDial = 1 Then
        Ci.Ring13Left.ForeColor = 0
        Ci.Ring13Left.BackColor = Green1
    End If
    
    
    If PagerDial = 0 Then
        If AlertType = 0 Then
            If Wes99 = 1 Then
                Ci.Ring13Left.BackColor = Blue1
                Ci.Ring13Left.ForeColor = White1
            Else
                Ci.Ring13Left.ForeColor = 0
                Ci.Ring13Left.BackColor = Green1
            End If
        End If
        
        If AlertType = 1 Then
            Ci.Ring13Left.BackColor = &HFF8080 'light blue
            Ci.Ring13Left.ForeColor = 0
        End If
    End If
End If


Call CiStats.CiStatOnOffHook

If Dragon$ <> "Dial Out" Or Ci.Ring14Right.BackColor = Red1 Then CID_Run_Me_Interface

Call TimerIntervalChange

If Dragon = "No Carrier" Then Nutty = 0


End Sub



Public Sub MComm1_OnComm()

If Ci.MSComm1.PortOpen = False Then Exit Sub

Dim Buffer As Variant
    'Select Case ci.MSComm1.CommEvent
        ' Event messages.
     '   Case comEvReceive
            Buffer = Ci.MSComm1.Input
            m_Buffer = m_Buffer + Buffer '& StrConv(Buffer, vbUnicode)
            
    'End Select
End Sub



Public Sub CheckRs232(C, R34)

'Ci.MSComm1.CommPort = 3 ' Internal Acer

'CommPortSet
'Modem
Ci.MSComm1.CommPort = 8 ' Robotics
Ci.MSComm1.Settings = "56000,N,8,1"




    
'Call Init(1, Me.MSComm1)

On Error Resume Next

Ci.MSComm1.PortOpen = True

DIALER.TimerComm4.Enabled = True

'CiMessage.Label8.Caption = "Comm " + Format$(MSComm1.CommPort, "00") + " Modem Working"

CIMessLB8$ = "Comm " + Format$(Ci.MSComm1.CommPort, "00") + " Modem Working"

If Err.Number > 0 Then
 '   CiMessage.Label8.Caption = "Comm " + Format$(MSComm1.CommPort, "00") + " Modem Not Working"
    CIMessLB8$ = "Comm " + Format$(Ci.MSComm1.CommPort, "00") + " Modem Not Working"
    TimerComm1.Enabled = True
    Exit Sub
End If

' Flush the input buffer.
Ci.MSComm1.InBufferCount = 0

'cirs232 = FreeFile
'Open "W:\CALLERID\CI LOGS\cirs232.txt" For Append As #cirs232
'If C <> 10 Then Print #cirs232, Str$(Now); " - Not O.K."'

'Close #cirs232

Dim Ads$(20)

'Ci.MSComm1.Output = "ATS0=0S7=1S6=8S11=50" + vbCr
Ads$(1) = "ATZ"
Ads$(2) = "AT+FCLASS=0"
Ads$(3) = "AT+VCID=0"
'Ads$(4) = "ATM=0"
Ads$(4) = "ATS6=2"
Ads$(5) = "ATS7=2"
Ads$(6) = "ATX=3"
Ads$(7) = "ATS11=50"

'Ci.Timer5.Enabled = False

For r = 1 To 7
    If Ci.MSComm1.PortOpen = True Then _
        Ci.MSComm1.Output = Ads$(r) + vbCr
        C = 0
        For ro = 1 To 5
            dummy = DoEvents()
            Call MComm1_OnComm
            FromModem$ = m_Buffer
            If InStr(FromModem$, "OK") Then C = 10
            If C <> 10 Then
                Sleep 500
            End If
            If C = 10 Then Exit For
        Next

    m_Buffer = ""

    If C = 0 Then Exit For

Next

If C <> 10 Then MsgBox "Modem Wont Respond to ATZ Init"

End Sub




Private Sub SeXxYy()










'''''''''''' End


Exit Sub



If Screen.Width <> 19200 Then Exit Sub
If GetAsyncKeyState(1) < 0 Then Exit Sub
dialertop12 = False
    
Dim Rect21 As RECT
Dim Rect22 As RECT
Dim Rect23 As RECT
Dim Rect22me As RECT
Dim Hwnd21 As Long
Dim Hwnd22 As Long
Dim Hwnd23 As Long



Hwnd21 = GetWindowRect(HwndCTLTask1, Rect21)
If Hwnd21 = 0 Then
    HwndCTLTask1 = FindWindow(vbNullString, "CTLTask")
    Hwnd21 = GetWindowRect(HwndCTLTask1, Rect21)
End If
    

Hwnd22 = GetWindowRect(HwndCTLTask2, Rect22)
If Hwnd22 = 0 Then
    HwndCTLTask2 = FindWindow(vbNullString, "Microsoft Office Shortcut Bar")
    Hwnd22 = GetWindowRect(HwndCTLTask2, Rect22)
End If

HwndCTLTask3 = FindWindow("_GD_Sidebar", vbNullString)
If HwndCTLTask3 > 0 Then
    Hwnd23 = GetWindowRect(HwndCTLTask3, Rect23)

'a=Rect23.Top
End If
If Rect23.Top = 0 Then Hwnd23 = 0
    
'Hwnd22 = GetWindowRect(HwndCTLTask2, Rect22)
'If Hwnd22 = 0 Then
'    HwndCTLTask2 = FindWindow(vbNullString, "Microsoft Office Shortcut Bar")
'    Hwnd22 = GetWindowRect(HwndCTLTask2, Rect22)
'End If
If Hwnd22 = 0 And PreHwnd22 > 0 Then
If Ci.WindowState = 0 Then
    Ci.Top = 0
    CiTopPos = 0
    dialertop12 = True
End If
End If



If Hwnd23 = 0 And PreHwnda23 > 0 Then
If Ci.WindowState = 0 Then
    'Hwnd22me = GetWindowRect(Me.hWnd, Rect22me)
    'MoveWindow Me.hWnd, Rect22me.Left - (Rect23.Right - Rect23.Left), Rect22.Bottom, Rect22me.Right - Rect22me.Left, Rect22me.Bottom - Rect22me.Top, True
    Ci.Left = Screen.Width - Ci.Width
    dialertop12 = True
End If
End If

PreHwnd22 = Hwnd22

If Hwnd23 > 0 And PreHwndb23 <> Rect23.Left And PreHwnda23 > 0 Then
    Hwnd22me = GetWindowRect(Me.hWnd, Rect22me)
'    MoveWindow Me.hWnd, Rect22me.Left - 30 - (Rect23.Right - Rect23.Left), Rect22me.Bottom, Rect22me.Right - Rect22me.Left, Rect22me.Bottom - Rect22me.Top, True
    'MoveWindow Me.hWnd, Rect23.Left - (Rect22me.Right - Rect22me.Left) + 1, Rect22me.Top, Rect22me.Right - Rect22me.Left, Rect22me.Bottom - Rect22me.Top, True
    Ci.Left = Rect23.Left * Screen.TwipsPerPixelX
    Stop
    dialertop12 = True
End If

PreHwnda23 = Hwnd23
PreHwndb23 = Rect23.Left

If Hwnd21 > 0 Then
    If Ci.Top = Pleet And TopToolBar1 <> Rect21.Bottom Then
        If Rect21.Bottom = 2 Or Rect21.Bottom > 300 Then Pleet = 0 Else Pleet = 390
        dialertop12 = True
    End If
    TopToolBar1 = Rect21.Bottom
End If
    
If Hwnd22 > 0 Then
    
    If Rect22.Top <> OldRect22Top.Top Or Rect22.Left <> OldRect22Top.Left Then
        
        CiTopPos = Screen.TwipsPerPixelY * Rect22.Bottom
        
        'If Rect22.Bottom = 18 Then PlimSole = 265
        'If Rect22.Bottom = 32 Then PlimSole = 460
        'If Rect22.Bottom = 2 Then PlimSole = 0
    
        
        
        If Rect22.Top > 2 Then
            CiTopPos = 0
            Ci.Top = 0
            dialertop12 = True
        End If
        
        If Rect22.Top <= 2 Then
        
            RectPib = RectME
        
            Hwnd22me = GetWindowRect(Me.hWnd, Rect22me)
            If Hwnd23 > 0 Then
                'MoveWindow Me.hWnd, (Rect22me.Right - Rect22me.Left) - 5 - (Rect23.Right - Rect23.Left), Rect22.Bottom, Rect22me.Right - Rect22me.Left, Rect22me.Bottom - Rect22me.Top, True
                Ci.Left = Rect23.Left * Screen.TwipsPerPixelX
                
                PreHwndb23 = Rect23.Left
                'Stop
                'Moves back from google side bar
            End If
            
            If Hwnd23 = 0 Then
                'MoveWindow Me.hWnd, Rect22me.Left, Rect22.Bottom, Rect22me.Right - Rect22me.Left, Rect22me.Bottom - Rect22me.Top, True
                Ci.Top = Rect22.Bottom * Screen.TwipsPerPixelX
                Ci.Left = Screen.Width - Ci.Width
            End If
            
            CiTopPos = Ci.Top
            Call DIALER.DialerTop
        
        End If
        
    
    End If
    
    OldRect22Top = Rect22
    
End If

'If Ci.WindowState = 1 And OldCIWinState = 0 Then
'    For i = Forms.Count - 1 To 0 Step -1
'        Forms(i).WindowState = 1
'    Next
'End If

'If Ci.WindowState = 0 And OldCIWinState = 1 Then
'    For i = 0 To Forms.Count - 1
'        Forms(i).WindowState = 0
'    Next
'End If

OldCIWinState = Ci.WindowState


If Ci.Top < CiTopPos Then
    
    
    If Ci.WindowState = 0 Then
        Ci.Top = CiTopPos
    
        If Hwnd23 > 0 Then
            Hwnd22me = GetWindowRect(Me.hWnd, Rect22me)
'            MoveWindow Me.hWnd, (Rect22me.Right - Rect22me.Left) - 5 - (Rect23.Right - Rect23.Left), Rect22.Bottom, Rect22me.Right - Rect22me.Left, Rect22me.Bottom - Rect22me.Top, True
            Ci.Left = Rect23.Left * Screen.TwipsPerPixelX
            dialertop12 = True
            PreHwndb23 = Rect23.Left
            'dialer.Top
        End If
    
        If Hwnd23 = 0 Then
        Ci.Left = Screen.Width - Ci.Width
        End If

    End If
        
    dialertop12 = True
End If


If dialertop12 = True Then
    Call DIALER.DialerTop
End If

End Sub


