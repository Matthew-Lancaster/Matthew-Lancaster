VERSION 5.00
Begin VB.Form Form_SEND_TO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008080&
   Caption         =   "SEND TO SCRIPT SUB FOLDER FILES SET"
   ClientHeight    =   6492
   ClientLeft      =   228
   ClientTop       =   588
   ClientWidth     =   7344
   Icon            =   "temp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6492
   ScaleWidth      =   7344
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   624
      Left            =   9000
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "EXPLORER CURRENT INI FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.04
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   600
      TabIndex        =   37
      Top             =   4680
      Width           =   5652
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAR SCAN SENDTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.04
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   600
      TabIndex        =   36
      Top             =   600
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAR GO FROM INI NUMERIC ONE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.04
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   600
      TabIndex        =   35
      Top             =   3000
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAR EDIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.04
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   600
      TabIndex        =   34
      Top             =   4080
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAR GO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.04
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   600
      TabIndex        =   33
      Top             =   2400
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAR XSCRIPT PROCESS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.04
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   600
      TabIndex        =   32
      Top             =   1800
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAR SCAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.04
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   600
      TabIndex        =   31
      Top             =   1200
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "Label4"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5280
      TabIndex        =   30
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   3960
      TabIndex        =   29
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   3480
      TabIndex        =   28
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   3480
      TabIndex        =   27
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   3720
      TabIndex        =   26
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   3480
      TabIndex        =   25
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   3720
      TabIndex        =   24
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   3720
      TabIndex        =   23
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   7
      Left            =   3960
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   8
      Left            =   3600
      TabIndex        =   21
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   9
      Left            =   3840
      TabIndex        =   20
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   10
      Left            =   3840
      TabIndex        =   19
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   11
      Left            =   4080
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   12
      Left            =   3840
      TabIndex        =   17
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   13
      Left            =   4080
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   14
      Left            =   4080
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   15
      Left            =   4440
      TabIndex        =   14
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   16
      Left            =   3600
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   17
      Left            =   3720
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   18
      Left            =   3720
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   19
      Left            =   3960
      TabIndex        =   10
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   20
      Left            =   3720
      TabIndex        =   9
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   21
      Left            =   3960
      TabIndex        =   8
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   22
      Left            =   3960
      TabIndex        =   7
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   23
      Left            =   4080
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   24
      Left            =   3840
      TabIndex        =   5
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   25
      Left            =   4080
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "000 SCAN TO GIVE "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.04
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   600
      TabIndex        =   3
      Top             =   0
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "OBJECTIVE FILE ONLY DIR ONLY CRC OPTION "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.04
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   600
      TabIndex        =   2
      Top             =   6480
      Width           =   5700
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "BEGIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.55
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB    "
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_PULL 
      Caption         =   "MENU PULL   "
      NegotiatePosition=   1  'Left
      Begin VB.Menu MNU_000_SCAN_GIVE 
         Caption         =   "000 --- SCAN SET TO GIVE"
      End
      Begin VB.Menu MNU_DIRECTORY_ONLY_01 
         Caption         =   "001 --- DIRECTORY ONLY"
      End
      Begin VB.Menu MNU_FILE_ONLY_02 
         Caption         =   "002 --- FILE ONLY"
      End
      Begin VB.Menu MNU_CRC_03 
         Caption         =   "003 --- CRC OPTION"
      End
      Begin VB.Menu MNU_CRC_NOT_PEND_DRIVE_04 
         Caption         =   "004 --- CRC -- DETECT NOT PEN DRIVE OPTION"
      End
      Begin VB.Menu MNU_CRC_NOT_PEND_DRIVE 
         Caption         =   "004 --- CRC -- DETECT BELOW TEN MEGABYTE"
      End
      Begin VB.Menu MNU_ALL_DRIVE_SET_05 
         Caption         =   "007 --- SWING ALL DRIVE SET FILING SYSTEM"
      End
   End
   Begin VB.Menu MNU_00 
      Caption         =   "IRFAN"
      Begin VB.Menu MNU_IRFAN_01 
         Caption         =   "001 --- IRFAN GO"
      End
      Begin VB.Menu MNU_IRFAN_02 
         Caption         =   "002 --- IRFAN TEMP"
      End
      Begin VB.Menu MNU_IRFAN_03_NOTEPAD 
         Caption         =   "003 --- NOTEPAD"
      End
      Begin VB.Menu MNU_IRFAN_04 
         Caption         =   "004 --- IRFAN NOT PLAY MOVIES"
      End
      Begin VB.Menu MNU_IRFAN_05 
         Caption         =   "005 --- EDIT XSCRIPT"
      End
      Begin VB.Menu MNU_IRFAN_07_NOT_uSE_XSCRIPT 
         Caption         =   "007 --- NOT USE XSCRIPT"
      End
      Begin VB.Menu MNU_IRFAN_08_EXPLORER_INI 
         Caption         =   "008 --- EXPLORER FOLDER + SELECT FILE FROM CUR POS INI"
      End
   End
   Begin VB.Menu MNU_RERUN 
      Caption         =   "* RERUN *"
   End
End
Attribute VB_Name = "Form_SEND_TO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'JOBS ON THE WAY GLOBALLY
'
'SYNC PROGRAM FOR HASHING WITH SUB DIR
'TO CHECK THING LIKE ALL .OCX FILE VB
'
'###
'DO SYNC A COMMON MOUDLE IN .BAS OR FORM OR .CLS
'MUST DO MY OWN AS NOT ANY BEEN DONE YET
'
'###
'GOOD SYNC TO USE DOUBLE JOB WITH ONE NOT CONFLICT MODE
'TEST GOOD SYNC CAN RUN AND KEEP HISTROY IN ONE JOB AND SAVE LATEST
'ONLY IN ANOTHER JOB AT SAME PATH LEVEL
'
'WORK TO DO A PATH LEVEL ON DUPE CHECKER
'
'MEDIA INFAR LOADER TO USE
'STORE INI COUNT EACH FOLDER
'
'JOB TO DO CHECK MOST OCX FILE IN APP.PATH AND SYSTEM FOR CRC32
'IS CORRECT WITH VERSION NUMBER
'
'GET MEMOERY FOR ASUS P
'GET MEMEORY FOR ASUS X
'ANSWER HOW REGISTER MODEL NUMBER WITH ASUS P
'ANSWER FIND HOW DO HDMI ASUS P
'
'CLIPBOARD TO REGRAB LAST CLIPBOARD WHEN HOLDING SOMETHING
'MAYBE
'
'CLIPBOARD

'MOST PROGRAMS WITH VIEW VB FOLDER AND WITH THE VB ME BEOFRE
'AND A DEL VBW
'
'WRITE A OWN CODE TO FILE SEARCH A ALL DRIVES AND NETWORK
'
'MAKE SURE ASUS P WORKING WELL WITH KAT MP3
'
'SEARCH ALL VB VBP FILE TO SEE ANY USE WRONG VERSIN NUMBER
'LIKE TIDAL
'
'BEGIN LEARN USE COMMAND LINE GOOD SYNC FOR OTHER COMPUTER BIG JOBS
'
'CHECK WHAT BEEN FORGETTING WHILE WORK HARD
'
'KAT MOUSE BETTER
'HOOVER LOAD UP
'CLOSE ICON PROPER
'
'CCSS ON NETWORK
'
'MINI KEYBOARD
'OR LOGITEC
'
'BIGGEST ENERGY SAVER
'--------------------
'-MUM TV
'AND CHECK HDMI WORKING

'-MINI DISPLAY ON FRONT DOOR

'-LIGHT IN FRONT ROOM QUICKER
'
'-DISPLAY ON PIR CCSS
'-EASY TEST ONE IF ONLY ONE DISLAY WITHOUT LENS

'-PIR FOR BATH ROOM

'-USB PORT SPLITTER MUST GO BACK CPC

'-TV STAND FOLD OUT NEAR BED

'-DRILL TO DO WITH BATTERYT

'-JOB FINISH DUPE CHECKER ERROR
'-

'WHY DO ICON IMAGE IN CUSTOMIZE TOOLBAR VB NOT REPLANCE AGAIN MORE

'FINISH HARD DRIVE PSU PROJECT
'WITH TEST OTHER UNIT
'BETTER JOB AT LOWER TO 12V WITHOUT HIGH 15V KICKING IN VOLT FIRST

'LONGER RECORD KAT MP3

'TIMER FOR ON A BED AND TO SAVER ENERGY

'AN MORE OR TYPE REED RELAY SETUP OF PIR
'ONE IN KITCHEN
'AND ONE FOR MONITOR BY BED BUT LONGER SO NOT DROP OUT WITH TIMER

'LONG USB TO BED UNIT
'MAYBE ETHERNET OPTION










'JOB CHECK AFTER 5TXSCRIPT PROCESSED THAT COUNT IS UPDATE EVERYWHERE

'------------------------------------------------------------
'JOB
'CODE REM OUT NOT WORK AT THE MOMENT
'HAVE TO MAKE IT PER USER COMPUTER
'------------------------------------------------------------
'IF SCAN TO DO SAME AS LAST DON'T RESET COUNT INI BACK NAUGHT
'Call MNU_INI_BACK_NAUGHT_Click
'------------------------------------------------------------





Dim INI_FILE_AT_LINE

Dim DUPE_LAST_PATH2_ERROR_VAR
Dim Form_SEND_TO_Label10_VAR_HOLD

Dim FLAG_TO_DO_XSCRIPT
Dim FILE19X_SCRIPT_DATE_MODIFIED

Dim AFTER_XSCRIPT_COUNTER
Dim BEFORE_XSCRIPT_COUNTER
Dim FILE_SCRIPT_COUNTER

Dim DUPE_LAST_PATH
Dim FOLDER_PATH_IRFANView


'Dim XTOG

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Dim FILEPATH

Dim FILEPATH_I_VIEW_INI
Dim X_SP_F_i_view_ini
Public FILE13_SCRIPT_TEMP_FOLDER

Dim INI_StartIndex
Dim RC_INI


Dim FFILEAR(), WORKFLAG
Dim OBJECTIVE_BRIEF
Dim i
Dim i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Dim OFILESPEC
Dim HHAR01()
Dim HHAR02()
Public FILE11, FILE12
Dim DAT1(), AT$
Dim QQ1 As Long
Private Declare Function TouchFileTimes Lib "imagehlp.dll" (ByVal FileHandle As Long, ByRef pSystemTime As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As BOOL
Enum BOOL
   FALSE_ = 0
   TRUE_ = 1
End Enum
'Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Private Declare Function TouchFileTimes Lib "imagehlp.dll" (ByVal FileHandle As Long, ByRef pSystemTime As Any) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
'Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'right order


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


Private Sub SET_VAR_FS()

Set FS = CreateObject("Scripting.FileSystemObject")

End Sub

Private Sub Form_Load()

If App.PrevInstance = True Then End

IRFAN_TO_DO_FROM_SEND_TO = False

Call SET_VAR_FS
ScanPath.Visible = False

Me.WindowState = vbMaximized

'01 OF 01
'FOLDER OF PROGRAM
Call GET_FOLDER_PATH_IRFANView

'----------------------------------
'FIRST LOAD OF THIS VAR HAPPEN HERE
'----------------------------------
FILE13_SCRIPT_TEMP_FOLDER = "C:\TEMP\IRFAN_SLIDESHOW-" + GetComputerName + ".txt"

Call XSCRIPT_CHUNK_DATE_UPDATE_SHOW_REQUIRE

'---------------------------------------------


'---------------------------------------------
'SEAECH -- Control.FontSize
'---------------------------------------------
For Each Control In Controls
    If UCase(Mid(Control.Name, 1, 5)) = "LABEL" Then
        Control.FontSize = 11
    End If
Next
'---------------------------------------------

'Label6.Caption = "IRFAN XSCRIPT PROCESS -- IF MANUAL EDIT DONE"


'---------------------------
MNU_IRFAN_03_NOTEPAD.Caption = "003 --- NOTEPAD++ -- C:\TEMP\IRFAN_SLIDESHOW-" + GetComputerName + ".txt"
Label8.Caption = "IRFAN EDIT -- NOTEPAD++ -- C:\TEMP\IRFAN_SLIDESHOW-" + GetComputerName + ".txt"
'---------------------------

'------------------------
'ONLY GET THE INI POINTER
'------------------------
Call MNU_INI_READ

'------------------------------------------------------------
'GET THE INI FILE AT POINTER -- AND UPDATE ALL LABELS ON FORM
'------------------------------------------------------------
Call Label0_Click




'If Command$ <> "" Then MNU_IRFAN_02_Click
'Call Test4




Exit Sub
End



'Call Test2
'Exit Sub
Call Test1
End
Exit Sub

'1 pound = 0.0714285714 stone
'    More about calculator.

'1 pound = 0.0714285714 stone
stone = 112 * 0.0714285714
stone = 212 * 0.0714285714
Debug.Print stone
'= 7.9999999968
'= 15.1428571368
 
 
 'More than a year ago, in July 2007, International Space Station astronauts threw an obsolete, refrigerator-sized ammonia reservoir overboard. Ever since, the 1400-lb piece of space junk has been circling Earth in a decaying orbit--and now it is about to reenter.
 
stone = 1400 * 0.0714285714
Debug.Print stone

 
 
 
 
' That 's cheating! You still don't know HOW it's done.

'256 ^ 0 = 1
'256 ^ 1 = 256
'256 ^ 2 = 65,536
'256 ^ 3 = 16,777,216

'So, 194.247.44.146 is ...

'146 * 256^0 = 146
'44 * 256^1 = 11,264
'247*256^2 = 16,187,392
'194*256^3 = 3,254,779,904
'
'Add those up, and ...

'146 + 11,264 + 16,187,392 + 3,254,779,904 = 3270978706
'
'3410960384  3410964479  UK  UNITED KINGDOM

Dim a0
Dim A1
a0 = 3270978706#
A1 = a0 / (256 ^ 3)
a2 = A1 - Int(A1) / 256 ^ 2
a3 = 3270978706# / 256 ^ 1
a4 = 3270978706# / 256 ^ 0

 
'203.79.32.0
'203.79.47.255
 
 
 
End Sub


Sub Test2()

Dim FS, Idate As Date, Hdate As String, File1 As String, FILE2 As String
Set FS = CreateObject("Scripting.FileSystemObject")

Call GetFileDates

File1 = Mid$(List1.List(List1.ListCount - 1), 21)
For r = List1.ListCount - 2 To 0 Step -1
FILE2 = Mid$(List1.List(r), 21)
If Mid$(List1.List(List1.ListCount - 1), 1, 19) <> Mid$(List1.List(r), 1, 19) Then
FS.COPYFILE File1, FILE2
TT = LastModifiedToCurrent(FILE2)
End If
Next

Call GetFileDates

End Sub

Public Sub GetFileDates()
Dim FS, Idate As Date, Hdate As String, File1 As String, FILE2 As String
Set FS = CreateObject("Scripting.FileSystemObject")
List1.Clear
File1 = "D:\VB6\VB-NT\00_Best_VB_01\Banging_Tunes\BangList Total Logg.html"
Call Date_File(File1, Idate, Hdate)
'dd = Hdate + " " + file1
List1.AddItem Hdate + " " + File1

For r = 3 To 25
On Error Resume Next
Err.Clear
Set g = FS.GetDrive(FS.GetDriveName(FS.GetAbsolutePathName(Chr$(r + 64) + ":")))
On Error GoTo 0
If InStr(RR, g.VolumeName) = 0 And Err.Number = 0 Then
RR = RR + g.VolumeName + "-"
If Mid(g.VolumeName, Len(g.VolumeName) - 1, 2) = "GB" Then

FILE2 = Chr$(r + 64) + ":\Banging\BangList Total Logg.html"
Call Date_File(FILE2, Idate, Hdate)
List1.AddItem Hdate + " " + FILE2

'fs.copyFile file1, file2


End If
'Nuke4 = Nuke4 + g.totalsize / 1024 ^ 3
'Nuke5 = Nuke5 + g.freespace / 1024 ^ 3
End If
Next

End Sub

Public Sub Date_File(Filespec2, Idate As Date, Hdate As String)
'Call Date_File(filespec2$, Idate)

Dim FS, F
Set FS = CreateObject("Scripting.FileSystemObject")

pdate = 0
FILESPEC = Filespec2
Idate = DateSerial(1980, 1, 1)
If FS.FileExists(FILESPEC) Then
Set F = FS.GETFILE(FILESPEC)
Idate = F.DateLastModified
End If
Hdate = Format$(Idate, "YYYY-MM-DD HH:MM:SS")


End Sub


Public Function FindFileSize(Filename As String) As Long
        On Error Resume Next
        Dim filedata As WIN32_FIND_DATA

        filedata = Findfile(Filename)
        
        If filedata.nFileSizeHigh = 0 Then
            FindFileSize = Format$(filedata.nFileSizeLow, "###,###,###")
        Else
            FindFileSize = Format$(filedata.nFileSizeHigh, "###,###,###")
        End If
End Function



Function LastModifiedToCurrent(duFile As String)
    Dim lngHandle As Long
    lngHandle = CreateFile(duFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If TouchFileTimes(lngHandle, ByVal 0&) = 0 Then
        'MsgBox "Error while changing file dates. Access Denied." + vbCrLf + duFile, vbCritical, "Modify Error"
    LastModifiedToCurrent = False
    Else
        'MsgBox "File date changed successfully.", vbInformation, "Modify Successful"
    LastModifiedToCurrent = True
    End If
    CloseHandle lngHandle
End Function

Sub Test1()

'restore header on wave fle
Open "m:\DWR 2008-11-27 00-36-03 -- 2008-11-28 13-02-14 -#- Janice Paul Short.wav" For Binary As #1
ll$ = "TTTT"
Get #1, 5, ll$
QQ1 = LOF(1)
Close #1
A1 = Asc(Mid$(ll, 1, 1))
a2 = Asc(Mid$(ll, 2, 1))
a3 = Asc(Mid$(ll, 3, 1))
a4 = Asc(Mid$(ll, 4, 1))
Mid(ll, 1, 1) = Chr(&H78)
Mid(ll, 2, 1) = Chr(&H10)
Mid(ll, 3, 1) = Chr(&H83)
Mid(ll, 4, 1) = Chr(&H23)

Mid(ll, 1, 1) = Chr(&H78)
Mid(ll, 2, 1) = Chr(&H94)
Mid(ll, 3, 1) = Chr(&H45)
Mid(ll, 4, 1) = Chr(&H0)

Open "M:\05 Media\Digital Wave Player M Drive\DR2008\DWR 2008-11-26 07-34-18 -- 2008-11-26 07-41-10 -#- New Anita Dingo Missed Steve On Linda On.wav" For Binary As #1
qq2 = LOF(1)
Close #1
Open "M:\05 Media\Digital Wave Player M Drive\DR2008\DWR 2008-11-26 09-39-09 -- 2008-11-27 00-35-50 -#- Steve Ibrag Linda Itching For Attack But Lets Clearly See Who Is Attacking First.wav" For Binary As #1
'qq2 = LOF(1)
jj$ = Space$(&H86)
Get #1, 1, jj$
Close #1

HH = &H78108323
HH = &H457894
gg1 = Hex$(QQ1 - 8)
gg2 = Hex$(qq2 - 8)
'4,560,000

Mid(ll, 1, 1) = Chr(&H78)
Mid(ll, 2, 1) = Chr(&H94)
Mid(ll, 3, 1) = Chr(&H45)
Mid(ll, 4, 1) = Chr(&H0)

Mid(ll, 1, 1) = Chr(&H78)
Mid(ll, 2, 1) = Chr(&H60)
Mid(ll, 3, 1) = Chr(&H8C)
Mid(ll, 4, 1) = Chr(&H56)

'QQ1 = QQ1 * 1.03025 '37.26,11 - correct from time files

'QQ1 = QQ1 * 1.032

QQ1 = QQ1 - 29000 '* 1.0000001 '36:20:11 - correct run lenght

gg3 = Hex$(Int(QQ1))
ll = ""
lDataSize = Len(gg3)
Dim i As Integer
i2 = 1
For i = lDataSize - 1 To 1 Step -2
'For i = 1 To lDataSize Step 2
    'bytearray(i2) = Val("&h" + Mid$(gg3, i, 2))
    ll = ll + Chr(Val("&h" + Mid$(gg3, i, 2)))
    i2 = i2 + 1
Next



'T1 = DateSerial(bytearray(1), bytearray(2), bytearray(3),bytearray(4))

T1 = DateDiff("h", DateSerial(2008, 11, 27) + TimeSerial(0, 36, 3), DateSerial(2008, 11, 28) + TimeSerial(13, 2, 14))
t2 = DateDiff("n", DateSerial(2008, 11, 27) + TimeSerial(0, 36, 3), DateSerial(2008, 11, 28) + TimeSerial(13, 2, 14)) Mod 60
t3 = DateDiff("s", DateSerial(2008, 11, 27) + TimeSerial(0, 36, 3), DateSerial(2008, 11, 28) + TimeSerial(13, 2, 14)) Mod 60

A1 = Asc(Mid$(ll, 1, 1))
'a2 = Asc(Mid$(ll, 2, 1))
'a3 = Asc(Mid$(ll, 3, 1))
'a4 = Asc(Mid$(ll, 4, 1))
getb$ = " "
Open "m:\DWR 2008-11-27 00-36-03 -- 2008-11-28 13-02-14 -#- Janice Paul Short.wav" For Binary As #1
'll$ = "TTTT"
Put #1, 1, jj$
Put #1, 5, ll
Put #1, &H7D, ll
'Get #1, LOF(1), getb$
'Put #1, LOF(1), Chr$(0)
Close #1

A = Hex$(Asc(getb$))
A = Len(ll)

Shell "C:\Program Files\Windows Media Player\mplayer2.exe ""m:\DWR 2008-11-27 00-36-03 -- 2008-11-28 13-02-14 -#- Janice Paul Short.wav""", vbNormalFocus

End
End Sub

Sub Test3()
'Form1.Hide
'ScanPath.Show
'ScanPath.TxtPath.Text = "M:\04 Music ---\04 My Music\01 Banging Tunes"
'ScanPath.TxtPath.Text = "M:\04 Music ---\04 My Music\01 Banging Tunes"
'ScanPath.cmdScanDir_Click
'
'For WE = 1 To ScanPath.ListView2.ListItems.Count
'    A1$ = ScanPath.ListView2.ListItems.Item(WE).SubItems(1)
'    B1$ = ScanPath.ListView2.ListItems.Item(WE)
'    A1$ = A1$ + B1$
'
'    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
'
'    et = InStr(A1$, "\200")
'    If et > 0 Then
'    If et = Len(A1$) - 4 Then
'    ef = InStrRev(A1$, "\", et - 1)
'
'    'If Len(a1$) > Len(ScanPath.txtPath.Text) + 1 And InStr(Len(ScanPath.txtPath.Text) + 2, a1$, "\") > 0 Then
'
'    tg1 = Mid$(A1$, ef, et - ef)
'    tg2 = Mid$(A1$, et + 1, 4)
'    tg3 = Mid$(A1$, 1, ef) + Mid$(tg1, 2, 7) + tg2 + " - " + Mid$(tg1, 16)
'    'ef = InStr(et + 6, a1$, "\")
'    'If ef = 0 Then
'    TY = Mid$(A1$, 1, et - 1)
'    'tx = Mid$(a1$, 1, Len(ScanPath.txtPath.Text) + 12) + " " + tg1 + " - " + Mid$(a1$, Len(ScanPath.txtPath.Text) + 14)
'    'ty = a1$ + "\" + Mid$(a1$,et+5 , 5)
'    On Error GoTo 0
''    MkDir ty
''    Name ty As tg3
''    MkDir tg3
'    'End If
'    Set FS = CreateObject("Scripting.FileSystemObject")
'    'ScanPath.txtPath.Text = a1$
'    'ScanPath.cmdScan_Click
'
'    Set F = FS.GetFolder(A1$)
'        Set fc = F.Files
'        For Each F1 In fc
'            's = s & f1.Name
'            's = s & vbCrLf
'    '    Next
'
'
'    'For we2 = 1 To ScanPath.ListView1.ListItems.Count
'    '    a1$ = ScanPath.ListView1.ListItems.Item(we2).SubItems(1)
'    '    b1$ = ScanPath.ListView1.ListItems.Item(we2)
'
'
'    Set F1 = FS.GETFILE(A1$ + "\" + F1.Name)
'    F1.Move tg3 + "\" + F1.Name
'    Next
'    'F.Move tg3
'    Set fc = F.SubFolders
'        For Each F1 In fc
'            Set F1 = FS.GetFolder(A1$ + "\" + F1.Name)
'            F1.Move tg3 + "\" + F1.Name
'
'        Next
'
'    RmDir A1$
'
'    End If
'    End If
'Next

End Sub

Private Sub Form_Unload(Cancel As Integer)

If UNLOAD_WITH_IRFAN_KILL = False Then

    Shell FOLDER_PATH_IRFANView + " /killmesoftly"

End If

Dim Control

'SET ALL TIMERS IN ALL FORMS ENABLED=FALSE
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
   
Dim Form As Form
For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next Form

'End

End Sub


Sub Test4()

Exit Sub
End
End Sub

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function

'***********************************************




Sub Label0_Click()

'THIS RUN AT FORM_LOAD

Me.Show
DoEvents

Set FS = CreateObject("Scripting.FileSystemObject")

Call GIVE_DRIVE_COMMAND_STRING

For Each Control In Controls
    If Control.Name = "MNU_000_SCAN_GIVE" Then
        Control.Caption = "000 --- OBJECTIVE SCAN - " + AT$
    End If
Next

'----------------------------------
'FIRST LOAD OF THIS VAR HAPPEN HERE
'----------------------------------
'FILE13_SCRIPT_TEMP_FOLDER = "C:\TEMP\IRFAN_SLIDESHOW-" + GetComputerName + ".txt"

If Dir(FILE13_SCRIPT_TEMP_FOLDER) <> "" Then
    Set F = FS.GETFILE(FILE13_SCRIPT_TEMP_FOLDER)

    '------------------------------------------------
    'DO GET LINE COUNT
    '------------------------------------------------
    'DO THIS IN CHUNK MODE AND COUNT VBCRLF - QUICKER
    '------------------------------------------------
    FF01 = FreeFile
    Open FILE13_SCRIPT_TEMP_FOLDER For Input As #FF01
    Do
        If Not (EOF(FF01)) Then Line Input #FF01, line3
        FILE_SCRIPT_COUNTER = FILE_SCRIPT_COUNTER + 1
        
        If Val(INI_StartIndex) = FILE_SCRIPT_COUNTER Then
            INI_FILE_AT_LINE = line3
        End If

    Loop Until EOF(FF01)
    Close #FF01
    Label11.Caption = "EXPLORER CURRENT INI FILE = " + INI_FILE_AT_LINE


End If

Label3.Caption = "SCAN TO GIVE -- " + AT$
If Dir(FILE13_SCRIPT_TEMP_FOLDER) <> "" Then
    
    
    Label5.Caption = "IRFAN -- " + FILE13_SCRIPT_TEMP_FOLDER + " -- FSIZE -" + Str(F.Size) + " -- LINE COUNT =" + Str(FILE_SCRIPT_COUNTER)
    'Label6.Caption = "IRFAN XSCRIPT PROCESS - FSIZE --" + Str(F.Size) + "LINE COUNT - BEFORE=" + Str(BEFORE_XSCRIPT_COUNTER) + " - AFTER=" + Str(AFTER_XSCRIPT_COUNTER)


End If

'Label2.FontSize = 15
Label2.WordWrap = False
Label2.AutoSize = True

Label2 = "OBJECTIVE " + vbCrLf
Label2 = Label2 + "------------------" + vbCrLf
Label2 = Label2 + "GIVE -- FILE SCRIPT AS BOTH OPTION COMPLEX AND NORMAL" + vbCrLf
Label2 = Label2 + "GIVE -- FILE ONLY OPTION" + vbCrLf
Label2 = Label2 + "GIVE -- DIRECTORY ONLY OPTION" + vbCrLf
Label2 = Label2 + "GIVE -- LATER -- CRC OPTION" + vbCrLf
Label2 = Label2 + "GIVE 2012-OCT-21 -- SEND TO NOW HAS WORKING WITH IRFAN" + vbCrLf
Label2 = Label2 + "       AND XSCRIPT DONT PROCESS WITH IT" + vbCrLf
Label2 = Label2 + "OPTION TO GIVE DRIVE WHOLE SELECTOR TO ROOT DOWN BY FILE SEND-TO" + vbCrLf

Label2 = Label2 + "------------------" + vbCrLf
Label2 = Label2 + "PRESS BUTTON AIM HERE - BEGIN AS OPTION NUMBER ONE" + vbCrLf
Label2 = Label2 + "PRESS BUTTON AIM HERE - BEGIN AS OPTION NUMBER ONE" + vbCrLf
Label2 = Label2 + "------------------"

'Label2.BackColor = QBColor(14)
'MsgBox Label2

'Label2.Top = 200
'Label2.Left = 0


Label9.Caption = "IRFAN GO - RESET INI TO NUMERIC #1 ONE -- FROM INI =" + Str(INI_StartIndex) + " Of " + Str(FILE_SCRIPT_COUNTER)

If FLAG_TO_DO_XSCRIPT = True Then
    Label7.Caption = "IRFAN GO - FROM INI #" + Str(INI_StartIndex) + " Of " + Str(FILE_SCRIPT_COUNTER) + " -- INI COUNT WILL RESET TO #1 AT RUN - FOLDER CHANGE HAPPEN"
Else
    Label7.Caption = "IRFAN GO - FROM INI #" + Str(INI_StartIndex) + " Of " + Str(FILE_SCRIPT_COUNTER)
End If


'-------------------
'LABEL CAPTIONS DONE
'-------------------

'-------------------
'FORM LAYOUT TO DO
'-------------------

On Error Resume Next
Me.Width = Label2.Width + 1000
Me.Height = Label2.Height + 4500
On Error GoTo 0

DoEvents


For Each Control In Controls
    If Mid(Control.Name, 1, 5) = "Label" Then
        Control.Left = 0
        Control.Width = Me.Width
    End If
Next




ReDim RAF(70) 'AMPLE
i = 0
'i = i + 1: RAF(i) = "3"
    i = i + 1: RAF(i) = "5"
    i = i + 1: RAF(i) = "10"
    i = i + 1: RAF(i) = "6"
    i = i + 1: RAF(i) = "7"
    i = i + 1: RAF(i) = "9"
    i = i + 1: RAF(i) = "8"
    i = i + 1: RAF(i) = "11"
    i = i + 1: RAF(i) = "2"
ReDim Preserve RAF(i)

Label3.Top = 550
HDIFF = 125
OI = Label3.Top + Label3.Height + HDIFF

For r = 1 To UBound(RAF)
    For Each Control In Controls
        If Mid(Control.Name, 1, 5) = "Label" Then
            If Mid(Control.Name, 6) = RAF(r) Then
                Control.Top = OI
                OI = Control.Top + Control.Height + HDIFF
            End If
        End If
    Next
Next


On Error Resume Next
Me.Height = Label2.Top + Label2.Height + 1000
On Error GoTo 0

'Label2.Left = (Me.Width - Label2.Width) / 2
'Label3.Left = (Me.Width - Label3.Width) / 2
'Label5.Left = (Me.Width - Label5.Width) / 2

'Label3.Width = Me.Width / 2
'Label3.Width = Me.Width / 2
'Label5.Width = Me.Width / 2

'100
'200
'300
'400
'500
'600
'700
'800
'900
'1000
'1100

'MsgBox (Me.Width - Label2.Width) / 2




On Error Resume Next
    Me.Top = 1000
    Me.Left = 2400
On Error GoTo 0


'------------------------------------------
'FORM LAYOUT WITH THE DRIVES THAT ARE READY
'------------------------------------------
Call FORM_LAYOUT_WITH_DRIVES_READY


End Sub


Sub GIVE_DRIVE_COMMAND_STRING()

DUPE_LAST_PATH2_ERROR_VAR = False
'FILEPATH_APP_TMP1 = App.Path + "\Last Folder.tmp"
'If Dir(FILEPATH_APP_TMP1) <> "" Then FILEPATH_APP_TMP = FILEPATH_APP_TMP1

APP_PATH_VAR = App.Path + "\IRFANVIEW -- " + GetComputerName
If Dir(APP_PATH_VAR, vbDirectory) = "" Then
    MkDir APP_PATH_VAR
End If


FILEPATH_APP_TMP1 = APP_PATH_VAR + "\Last Folder --" + GetComputerName + "--.tmp"
FILEPATH_APP_TMP = FILEPATH_APP_TMP1

'If Dir(FILEPATH_APP_TMP, vbDirectory) = "" Then

'End If

'COMMAND$
'CLIPBOARD
'REMEMBER FILE

AT$ = Command$
AT$ = Replace(AT$, """", "")

'If IsIDE = True Then AT$ = "\\4-asus-gl522vw\4-asus-g_asus_g_d\DSC\2015 2016\2016 Sony CyberShot DSC-HX60V"
'If IsIDE = True Then AT$ = "\\4-asus-gl522vw\4-asus-g_asus_g_d\DSC\2015 2016"
ATISIDE$ = "D:\DSC"


If FS.FolderExists(AT$) = False Then
'    MsgBox "COMMAND$ FOLDER DONT EXIST" + vbCrLf + Command$
End If


'--------------------------
'FIND COMMANDLINE - SEND TO
'--------------------------
If FS.FolderExists(AT$) = True Then
    'AT$ = Clipboard.GetText
    If FS.FolderExists(AT$) = True Then
        If Form_SEND_TO.Label10.Caption = "IRFAN SCAN SENDTO" Then
            Form_SEND_TO.Label10.Caption = "IRFAN SCAN SENDTO - FROM -- SEND TO -- COMMAND_LINE"
        End If
    End If
End If

'---------------------
'NONE SEND TO CMD LINE -- THEN
'IS CLIPBOARD ANY
'---------------------
If FS.FolderExists(AT$) = False Then
    AT_TEMP$ = Clipboard.GetText
    If FS.FolderExists(AT_TEMP$) = True Then
        Form_SEND_TO.Label10.Caption = "IRFAN SCAN SENDTO - FROM CLIPBOARD"
        AT$ = AT_TEMP$
    End If
End If
    
'--------------------------
'STILL NONE AFTER CLIPBOARD
'USE REMEMBER FILE
'--------------------------
If FS.FolderExists(AT$) = False Then
    
    If Dir(FILEPATH_APP_TMP1) <> "" Then
        
        '------------
        FR1 = FreeFile
        Open FILEPATH_APP_TMP For Input As #FR1
            F_COMPARE1 = Input$(LOF(FR1), FR1)
            F_COMPARE1 = Replace(F_COMPARE1, vbCrLf, "")
            AT_TEMP$ = F_COMPARE1
        Close #FR1
        If FS.FolderExists(AT_TEMP$) = True Then
            AT$ = AT_TEMP$
            Form_SEND_TO.Label10.Caption = "IRFAN SCAN SENDTO - FROM MEMORY"
        Else
            Form_SEND_TO.Label10.Caption = "IRFAN SCAN SENDTO - FROM MEMORY - FOLDER NOT EXIST"
        End If
    End If
End If
    
    
'-----------------------------------
'NOT A PATH FROM CLIPBOARD OR MEMORY
'-----------------------------------
If FS.FolderExists(AT$) = False Then
    If IsIDE = True Then
        i = MsgBox("USE ISIDE HARD CODE PATH" + vbCrLf + ATISIDE$, vbMsgBoxSetForeground + vbYesNo)
        If i = vbYes Then
            Form_SEND_TO.Label10.Caption = "IRFAN SCAN SENDTO - FROM HARDCODE ISIDE"
            AT$ = ATISIDE$
            USE_ISIDE = True
        End If
    End If
End If
    
    
'----------------------------------
If FS.FolderExists(AT$) = True Then

'01# -- TUE -- 3RD -- NOV 1998
'02# -- WED -- 11TH -- DEC 2002
'WITHOUT ONE DAY AWOL PRIOR
'PRIOR PREVIOUS TO 01# -- 1 YEAR CHEMICAL COSH FREE
'
'----
'
'The Patient - Patent Invention -- CREATOR - CREATIVE
'Intricate Interesting Intriguing Item - Aymer Road Talk Hanna
'COPYRIGHTS
'Day Light
'SKYLIGHT
'OUTLOOK
'A BEAUTY SPOT AS SKY BEAUTY SPOT
'GOOD FOR A CAMARA PHOTO - BEAUTY SPOT IS NORMALLY TERAFIRMA
'GREENHOUSE EFFECT
'GLASS HOUSES
'
'HUMAN RIGHTS
'ANIMAL RIGHTS
'REAL MC-COY - FAKE - PHONEY - CHARLTON - THE REAL DILL
'PSYCHIATRY AT LAST WORLD MENTALHEALTH WELLBEING WEEK
'ENDED BY A QUOTE FROM THE ORGANISATION -WHO- AND A DOCTOR
'STILL WANT TO BRING IT INTO THE SAME LINE AS
'AS MENTAL HEALTH WELLBEING AS PSYCHICAL HEALTH
'COMPARE EXAMPLE OF TWO
'
'A BROKEN MIND IS THE SAME AS A BROKEN PSYCHICAL HEALTH
'A PSYCHICAL BROKEN LEG IS THE SAME AS A BROKEN MIND HEAD
'WHILE LOST YOUR CAPACITY TO APPEAR TO HOLD INFOMATION LIKE A MEDIA STORAGE
'WOULD ACCORDING TO THE DOCUMENT MAKER PEN-PUSHER DOCTOR CHARLTON
'REQUIRED An ISOLATION HOSPITAL REMOVED FROM THE OUTSIDE WORLD COMMUNITY AS IF IT WERE LIKE A PSYCHICAL PROBLEM DISEASE
'AN INJURY IS NOT A DISEASE
'AN INJURY MIGHT LEAVE A SCAR
'AND MIND HAS UNLIMITED THINKER CAPACITY
'USE IT OR LOSE IT
'CAN 'T FLEX YOU MIND MUSCLE ENOUGH
'CHEMICAL COSH
'IS LIKE BEING TAKEN TO THE CLEANER
'LAUNDRY SERVICING
'MI -- MONEY lAUNDERING
'MI--  MENTAL INTELLIGENCE
'LITERACY HARDCORE
'RSI IS LIKE BEING TAKEN TO THE CLEANER
'GIVE SOMETHING
'TAKE SOMETHING
'THEY PEN-PUSHER GIVE A LOT OF SOMETHING TO BE TAKEN
'JACK AND THE BEANSTALK
'SPEND YOU MONEY - TRAVELLING AROUND THE PLANET
'PICK A COUNTRY
'SEND YOUR POLLUTION TO SOMEONE ELSE'S GARDEN OF EDEN
'
'BUY SOMETHING IN TRADE OF A MATERIAL ITEM
'AND YOU GOT SOMETHING TO TREASURE
'
'QUOTE - THE HISTORY OF A PHILOPINIASIM WAR
'THE ANSWER GAINED IN A HISTORY LESSON WAS
'TO BUY AND TRADE AND NOT SMASH IT IN WAR
'
'THROW YOUR MONEY AT TRAVEL EXPENSES
'SPENDING HOLIDAY
'GROUNDBREAKING Time
'LONG BREAK
'
'
'
'OPPONENTS OF THE PSYCHIATRY - MODEL
'PROPOSE MENTAL HEALTH WELLBEING IS A SOCIAL PROBLEM
'AS YOUR HEAR MANY WIZE THERAPY COUNCIL SUPPORT AGONY UNTIE
'OFTEN LOOK AT THE
'DEFENCE OF THE MIND
'AND PARENT ARE OFFEN RESPONSIBLE FOR LACK OF COMMUNICATION
'AND ANY OPPONENT OF PSYCHIATRY IS LABEL
'A ANTI - PSYCHIATRY
'COURTESY OF THE TIME
'DR LAING - INVENTED TO WORD
'SOURCE UNCOVERED TO ME BY DR TOMAUOS SZASZ
'AND THEN PSYCHIATRY THAT THEN AWARD THEMSELVES THE
'DOCTOR DOCUMENT MAKER - EXPERT IN SOCIAL BEHAVIOUR
'TRANSLATE TO An ANTI-SOCIAL
'
'A SOCIAL MODEL ALTERNATIVE TO A CHEMICAL COSH PSYCHIATRY THEORY
'AND THEN HOW IS A MINORITY DOCUMENT MAKER PEN-PUSHER
'BECOME PILL PUSHER
'THEY SEE THEMSELVES AS TOP OF THE EMPLOYMENT
'WHITER THAN WHITE
'AND THEN CALLED LABEL THEMSELF AS PRO STAFF AND EXPERT
'HOW IS THAT A SOCIAL BEHAVIOUR EXPERT IN A MINORITY
'AGAINST THE COMMONER EN-MASS
'AND A DOCUMENT MAKER - MAKES DOLLAR 100K UK POUND A YEAR
'LIVE A LIVER OF LUXURY LIFESTYLE
'
'TROUBLE WITH OWN A CAR VEHICAL
'COMPARED TO BAREFOOT
'YOU ALWAYS GET A BRAND NAME LOGO WITH A VEHICAL
'IN THE CLUB - INVEST YOUR MONEY IN-IT
'BURN MARK -- RED HOT POKER
'RUST BUCKET -- AT THE END OF THE DAY - RAINBOW
'POLISH SHINE
'
'HORSE POWER - SLAVE DRIVER
'WHAT DID YOUR LAST SLAVE DIE OF
'i 'M VERY DEMANDING FROM SLAVES LIKE YOU
'
'DOCUMENT MAKER
'YOU TALK THE TALK
'BUT DO YOU WALK THE WALK
'
'
'
'----
'CHEMICAL COSH -- CHEMICAL FREE
'TOBACCONIST - PSYCHIATRY NHS - INVESTOR IN THE TOBACCO
'INDUSTRY
'GOVERNMENT HEALTH WARNING FROM CHEIF SURGEON
'AND TWO FACED MESSAGE
'TWO SIDED CION
'LIKE ALCOHOLIC - WOULDN'T GET MORE FREE SUPPLY OF ALCOHOL
'LIKE FANNING THE FLAMES
'FUELING THE FLAME
'FANATIC
'LAB RATS OF A VIVISECTOR MENTALITY
'PUT A TEST AFTER PLY LAB RAT WITH ALCOHOL THEY THEN SWITCH WATER BACK ON THE MENU AND THE LAB RAT STILL PREFER ALCOHOLIC
'i WOULDN 'T PISS ON YOU IF YOU WERE ON FIRE
'TEACH A MAN HOW TO FISH
'SET A MAN ON FIRE AND BE WARM REST OF LIFE
'BROTHER BURN YOUR HOUSE DOWN - JIMI HENDRIX
'DISCO INFERNO
'DISC DISK - MONEY SPINNER - ELECTRON
'DOCUMENT MAKER
'WHITE COLOR CONSERVATIVE - POINT THEIR FINGER AT ME
'IT 'S MY LIFE LET ME LIVE IT THE WAY I WANT TO
'----
'TO BE SECTIONOR OR TO BE SECTIONIST
'OR TO BE SECTION
'Beware the victim who is the victimizer.
'Sexicutist
'Secticutionor
'Sexicutionor
'Sexicutioner
'Section --Erection
'Section - Hard Drive Segment
'Section -Vivisector
'SEXPERT -- THE CHICKEN OR THE EGG
'EGG -SPERT
'
'CHECKMATE
'
'Get Smart Blossom Who's the Boss Full House BILL HICKS QUOTE
'Tortoise and the Hare
'What 's Up Doc
'Bugs Bunny - Nasty Bugger
'The Hare of Another Day - Terrestrial Set Cartoon World
'The Hair of the Dog - Alcohol - Morning After
'
'LOTTERY SUCCESS - FRUIT MACHINE -- FULL ROW
'winning Number
'WINNING LINE -- winning ticket
'TOP SEQUENCE MEMBER NUMBER
'MIDI SEQUENCER
'IGNITE CAPACITY - BRAINWAVE - BRAIN CHILD
'IGNITE CAPACITY
'LIGHTNING STORM
'BRIGHT IDEA
'CREATIVE IDEA
'MASTER BATCH FILE CREATOR
'MASTER BATTERY CAPACITY
'THUNDER PANTS - FOSSIL FUEL - AEROPLANE
'REFILL UP - TIGER UNDER YOUR BONNET SHELL COMPANY
'A SHIFT OUT OF SYNC JET'S ON AN AEROPLANE AT LOW HEIGHT ALTITUDE NEAR LONDON
'HORSHAM
'IS LIKE A SHELL SHOCK
'STEREO AUDIO
'---------------------------------
'SAME SHIT DIFFERENT DAY
'ALWAYS FAULT FINDING AND A COMPLAINT
'UGLY MIND
'---------------------------------
'-----------------------------------------------------------
'PLAGERISM - NOT THE REAL -- COPY WITHOUT
'DOCUMENT MAKER - CHARLTON - NOT THE REAL DILL
'PERMISSION
'SEXISM - GENDA BENDA - CONNECTOR MATE - MAKE CONTACT - LEGO
'FEMINISM - MACHO-ISM - DOMINATRIX - MISTRESS DIANNA
'SEXIEST
'PEST
'SEX SELLS TO A PERVERT
'ADVERT - BUSINESS NOT GOT ANY NAME NOT GOT ANY BRAIN
'-----------------------------------------------------------
'PREFERENCE -ISM
'PARTIAL
'SWING BOTH WAYS - SWINGER
'DON 'T MIND IF I DO
'COMPARE -ISM
'RELATION -Shape - ISM
'DIVERER -ISM - CAN 'T MAKE YOU MIND UP
'OVER PREPAREDNESS - ISM
'UNDER PREPAREDNESS - ISM
'Normal PREPAREDNESS - ISM
'ABOVE NORMAL PREPAREDNESS-ISM
'HIGH PREPAREDNESS - ISM
'REAL-TIME PREPAREDNESS-ISM
'ABNORMAL PREPAREDNESS - ISM
'PROPER PRIOR PREPARATION PREVENTS PISS POOR PERFORMANCE
'-----------------------------------------------------------

    '-----------------------------------------
    'DON'T SAY THERE IS A PROBLEM WHEN MIGHT BE ANSWER ONE STEP AROUND THE CORNER
    'AS LIKE PROGRAMMING
    'AS LIKE PSYCHATRY ARE PROGRAMMING MY HEAD
    '-----------------------------------------

    On Error Resume Next
    If IsIDE = True Then On Error GoTo 0
    
    FILEPATH_APP_TMP_COPY = Replace(FILEPATH_APP_TMP, "--.tmp", "--BACKCOPY.tmp")
    
    If Dir(FILEPATH_APP_TMP_COPY) <> "" Then
        '------------
        FR1 = FreeFile
        Open FILEPATH_APP_TMP For Input As #FR1
            F_COMPARE1 = Input$(LOF(FR1), FR1)
        Close #FR1
        '------------
        FR1 = FreeFile
        Open FILEPATH_APP_TMP_COPY For Input As #FR1
            F_COMPARE2 = Input(LOF(FR1), FR1)
        Close #FR1
        '------------
        If F_COMPARE1 = F_COMPARE2 Then
            '--------------------
            DUPE_LAST_PATH = True
        End If
    End If
    
    If Err.Number > 0 Then
        DUPE_LAST_PATH2_ERROR_VAR = True
    End If
    
    If DUPE_LAST_PATH = False And Err.Number = 0 Then
        If Dir(FILEPATH_APP_TMP_COPY) = "" Then
        
        End If
        
        'DON'T EXIST PPROBLEM MAYBE WRITTEN LATER
        ''YES BECAUSE ABOVE IS READ INPUT ONLY
        If Dir(FILEPATH_APP_TMP) <> "" Then
            FS.COPYFILE FILEPATH_APP_TMP, FILEPATH_APP_TMP_COPY
        End If
    End If
End If
    
If FS.FolderExists(AT$) = False Then
    Form_SEND_TO.Label10.Caption = "IRFAN SCAN SENDTO - NOT VALID PATH FIND"
End If
'----------------------------------
    
    
    
    
'DO THE WRITE IF GOOD OF STORED MEM PATH - NOT IF ISIDE
If FS.FolderExists(AT$) = True And USE_ISIDE = False Then
    'i = AT$
    FR1 = FreeFile
    Open FILEPATH_APP_TMP1 For Output As #FR1
        Print #FR1, AT$
    Close FR1
        
Else
    OBJECTIVE_BRIEF = "SINGLE DRIVE SELECT VIA BUTTON"
End If

'---------------------------------
'STORE THIS APPEND INFO DATA LATER
'---------------------------------
Form_SEND_TO_Label10_VAR_HOLD = Form_SEND_TO.Label10.Caption


'If IsIDE = True Then
    'AT$ = "M:\#\#D\00 Pen Drives MOBILES"
'    OBJECTIVE_BRIEF = "IDE MODE SET  " + AT$
    'AT$ = Clipboard.GetText

'End If

'If IsIDE = True Then AT$ = "C:\TEMP\" + Dir("C:\TEMP\*.*")

End Sub

Private Sub LabDR_Click(Index As Integer)

If LabDR(Index).BackColor = Label4.BackColor Then Exit Sub
AT$ = Chr(Index + 65) + ":\"

Label3.Caption = "SCAN TO GIVE -- " + AT$



End Sub


Public Sub NAUGHT_NULL(VAR1, VAR2, VAR3)

'Call Form_SEND_TO.NAUGHT_NULL(i1, F.TotalSize, i2)
VAR1 = Replace(VAR1, "000,", "___,")
VAR1 = Replace(VAR1, " 00,", "___,")
VAR1 = Replace(VAR1, "  0,", "___,")

VAR1 = Replace(VAR1, "_,0", "_,_")
VAR1 = Replace(VAR1, "_,00", "_,__")
VAR1 = Replace(VAR1, "_,_0", "_,__")
VAR1 = Replace(VAR1, ",__0", ",___")

If VAR3 = -1 Then Exit Sub

For r = 1 To 8
    RULER = 1024 ^ r
    If r = 1 Then RULETEXT = "KILO RANGE"
    If r = 2 Then RULETEXT = "MEGA RANGE"
    If r = 3 Then RULETEXT = "GIGA RANGE"
    If r = 4 Then RULETEXT = "TERA - TETRA - TERRORIZOR - RANGE"
    If VAR2 >= RULER Then VAR3 = RULETEXT
Next


End Sub

Private Sub Label1_Click()
' --- BEGIN





'Dim AT$

Set FS = CreateObject("Scripting.FileSystemObject")

Call GIVE_DRIVE_COMMAND_STRING


'Set F = FS.GetDrive(Mid(AT$, 1)) '+ ":\")
Set F = FS.GetDrive(FS.GetDriveName(AT$))

Me.Hide

ScanPath.Show
DoEvents
ScanPath.SetFocus
DoEvents
ScanPath.Show
DoEvents


ScanPath.TxtPath.Text = AT$

FILE5TXT = ScanPath.TxtPath.Text
FILE5TXT = Replace(FILE5TXT, "\", "_")
FILE5TXT = Replace(FILE5TXT, ":", "_")
FILE5TXT = FILE5TXT + " VOL - " + F.VolumeName

'Set F = FS.GetFolder(folderspec)
's = F.DateCreated

's = s & "Created: " & F.DateCreated & vbCrLf
's = s & "Last Accessed: " & F.DateLastAccessed & vbCrLf
's = s & "Last Modified: " & F.DateLastModified
'MsgBox s, 0, "File Access Info"

'DAT1 = F.DateCreated
'DAT1 = F.DateLastModified
'DAT1 = F.DateLastAccessed

If Len(ScanPath.TxtPath.Text) = 3 Then
    SUBTEXTFOLDER = "#0 ACHIVE - SCRIPT FILE FOLDER AND SUB\"
    TDIR = ScanPath.TxtPath.Text + SUBTEXTFOLDER
    
    If Dir(TDIR, vbDirectory) = "" Then
        MkDir TDIR
    End If
    
'    TDIR = SUBTEXTFOLDER
End If

FILE11 = "#00 File List DIR & SUB HERE -- " + Format(Now, "YYYY_MM_DD_HH_MM_SS") + " -- COMPLEX -- " + FILE5TXT + ".txt"
FILE12 = "#00 File List DIR & SUB HERE -- " + Format(Now, "YYYY_MM_DD_HH_MM_SS") + " -- NORMAL ---- " + FILE5TXT + ".txt"
FILE13 = "#00 File List DIR & SUB HERE -- " + Format(Now, "YYYY_MM_DD_HH_MM_SS") + " -- COMPLEX -- " + FILE5TXT + "- DIRECTORY.txt"
FILE14 = "#00 File List DIR & SUB HERE -- " + Format(Now, "YYYY_MM_DD_HH_MM_SS") + " -- NORMAL ---- " + FILE5TXT + "- DIRECTORY.txt"

Dim ORGTAGFILE(4)
ORGTAGFILE(1) = FILE11
ORGTAGFILE(2) = FILE12
ORGTAGFILE(3) = FILE13
ORGTAGFILE(4) = FILE14

FILE21 = TDIR + FILE11  '" -- COMPLEX --  "
FILE22 = TDIR + FILE12  '" -- NORMAL ---- "
FILE23 = TDIR + FILE13  '" -- COMPLEX --  DIR"
FILE24 = TDIR + FILE14  '" -- NORMAL ---- DIR"

'ABOVE
'WORK DIM THE ARRAY HERE
'BELOW

ReDim FFILEAR(4)
FFILEAR(1) = FILE21
FFILEAR(2) = FILE22
FFILEAR(3) = FILE23
FFILEAR(4) = FILE24

FF01 = FreeFile
Open FILE21 For Output As #FF01
FF02 = FreeFile
Open FILE22 For Output As #FF02
FF03 = FreeFile
Open FILE23 For Output As #FF03
FF04 = FreeFile
Open FILE24 For Output As #FF04

Dim FFAR(4)
FFAR(1) = FF01
FFAR(2) = FF02
FFAR(3) = FF03
FFAR(4) = FF04



' -------------------------------------- DC SET
' -------------------------------------- DC SET
' -------------------------------------- DC SET
DC = 19
For r = 1 To 4
    Print #FFAR(r), "---------------------------------------------------------------------------"
    Print #FFAR(r), "SOURCE -- SOLUTION -- SOLVABILITY"
    Print #FFAR(r), "---------------------------------------------------------------------------"
    Print #FFAR(r), "PATH :-:" + TDIR
    Print #FFAR(r), "FILE :-:" + ORGTAGFILE(r)
    Print #FFAR(r), "---------------------------------------------------------------------------"
    Print #FFAR(r), "FOLDER DIRECTORY OBJECTIVE SCAN -- "
    Print #FFAR(r), "## " + AT$
    Print #FFAR(r), "---------------------------------------------------------------------------"
    
    For R1 = 0 To DC
        Print #FFAR(r), String(140, "#")
    Next
    Print #FFAR(r), "---------------------------------------------------------------------------"
Next

    PROCESSTIMEDIFF = Now

' ---------------------------------------------------------------
' ----------- EXECUTE BEGIN ------------------------
' ---------------------------------------------------------------

ScanPath.cmdScan_Click

For r = 1 To 4
    Close FFAR(r)
Next

Select Case F.DriveType
    Case 0: TDriveType = "Unknown"
    Case 1: TDriveType = "Removable"
    Case 2: TDriveType = "Fixed"
    Case 3: TDriveType = "Network"
    Case 4: TDriveType = "CD-ROM"
    Case 5: TDriveType = "RAM Disk"
End Select

'Set F1 = FS.GetDrive(Mid(AT$, 1)) '+ ":\")

Set F = FS.GetDrive(FS.GetDriveName(AT$))

i1 = Format(mDIRCount2, "000,000,000,000,000")
i2 = Format(mFILECount2, "000,000,000,000,000")
i3 = Format(COUNTPROCES2, "000,000,000,000,000")

Call NAUGHT_NULL(i1, 0, 0)
Call NAUGHT_NULL(i2, 0, 0)
Call NAUGHT_NULL(i3, 0, 0)


'DC IS FIXED ABOVE --- DC = 20
ReDim DAT1(40)
DC = -1
DC = DC + 1: DAT1(DC) = "DIR:                  " + i1
DC = DC + 1: DAT1(DC) = "FILE:                 " + i2
DC = DC + 1: DAT1(DC) = "PROCESS:              " + i3

DC = DC + 1: DAT1(DC) = "---------------------------------------------------------------------------"
DC = DC + 1: DAT1(DC) = "Volume Name:          " + F.VolumeName
DC = DC + 1: DAT1(DC) = "Drive Type:           " & UCase(F.DriveType)
DC = DC + 1: DAT1(DC) = "Drive Type:           " & UCase(TDriveType)  ' --- ABOVE
DC = DC + 1: DAT1(DC) = "File System:          " & F.FileSystem
DC = DC + 1: DAT1(DC) = "Serial Number:        " & F.serialnumber
i1 = Hex$(F.serialnumber)
i1 = Mid(i1, 1, 2) + "," + Mid(i1, 3, 2) + "," + Mid(i1, 5, 2) + "," + Mid(i1, 7, 2)
DC = DC + 1: DAT1(DC) = "Serial Number &H:     " & i1

i1 = F.ShareName
If i1 = "" Then i1 = "LOCAL SYSTEM DRIVE"
DC = DC + 1: DAT1(DC) = "Share Name:           " & i1

i1 = Format(F.TotalSize, "000,000,000,000,000")
Call NAUGHT_NULL(i1, F.TotalSize, i2)
DC = DC + 1: DAT1(DC) = "Total Size:           " & i1 + " - " + i2

i1 = Format(F.AvailableSpace, "000,000,000,000,000")
Call NAUGHT_NULL(i1, F.AvailableSpace, i2)
DC = DC + 1: DAT1(DC) = "Available:            " & i1 + " - " + i2

i1 = Format(F.FreeSpace, "000,000,000,000,000")
Call NAUGHT_NULL(i1, F.FreeSpace, i2)
DC = DC + 1: DAT1(DC) = "Free Space:           " & i1 + " - " + i2

INOW = Now
i = Format(DateDiff("s", PROCESSTIMEDIFF, INOW) / 60, "0.0##")
DC = DC + 1: DAT1(DC) = "Process Lenght Time:  " + i + " Minute Divide Clock"
i = Trim(Str(DateDiff("s", PROCESSTIMEDIFF, INOW)))
DC = DC + 1: DAT1(DC) = "Process Lenght Time:  " + i + " -- Seconds"
DC = DC + 1: DAT1(DC) = "BEGIN TIME:           " + Format(PROCESSTIMEDIFF, "DD MMM YYYY HH:MM:SS") + "h"
DC = DC + 1: DAT1(DC) = "AFTER TIME:           " + Format(INOW, "DD MMM YYYY HH:MM:SS") + "h" + " -" + Str(i) + " Sec"
DC = DC + 1: DAT1(DC) = "FORMAT BEGIN TIME:    " + Format(PROCESSTIMEDIFF, "DDD DD MMM YYYY HH:MM:SSa/p")
If OBJECTIVE_BRIEF = "" Then OBJECTIVE_BRIEF = "NORMAL"
DC = DC + 1: DAT1(DC) = "OBJECTIVE MODE BRIEF: " + OBJECTIVE_BRIEF





'MsgBox DC


ReDim Preserve DAT1(DC)

For R1 = 1 To 4

    Dim CHUNK As String
    FF01 = FreeFile
    Open FFILEAR(R1) For Binary As #FF01
        CHUNK = Space(8000)
        If CHUNK > Len(FF01) Then
            CHUNK = Space(Len(FF01) - 2)
        End If
        
        Get #FF01, 1, CHUNK
        For r = 0 To UBound(DAT1)
            STRING1 = DAT1(r)
            UU0 = String(140, "#")
            UU1 = String(140, "#")
            Mid(UU1, 1, Len(STRING1)) = STRING1
            
            UU1 = Replace(UU1, "#", " ")
            CHUNK = Replace(CHUNK, UU0, UU1, , 1)
        Next
        Put #FF01, 1, CHUNK
    Close #FF01
    
    'STILL NEED ALL READ IN AGAIN
    'TROUBLE TO APPEND A BLOCK TO A LARGE FILE
    FF02 = FreeFile
    Open FFILEAR(R1) + ".tmp" For Output As #FF02
    FF01 = FreeFile
    Open FFILEAR(R1) For Input As #FF01
    Do
        Line Input #FF01, line3
        Print #FF02, Trim(line3)
    Loop Until EOF(FF01)
    Close #FF01, #FF02
    On Error Resume Next
    Kill FFILEAR(R1)
    If Err.Number > 0 Then MsgBox "Error with Kill Before Temp Rename" + vbCrLf + FFILEAR(R1) + vbCrLf + Err.Description
    Err.Clear
    Name FFILEAR(R1) + ".tmp" As FFILEAR(R1)
    If Err.Number > 0 Then MsgBox "Error with Rename Temp File" + vbCrLf + FFILEAR(R1) + vbCrLf + Err.Description
    On Error GoTo 0

Next




FILE32 = "D:\0 00 LOGGERS TEXT"
FILE31 = "D:\0 00 LOGGERS TEXT\ARCHIVE OF DIRECTORY SCRIPT"
If Dir(FILE31, vbDirectory) = "" Then
    MkDir FILE32
    MkDir FILE31
End If


For r = 1 To 4
    'EVEN ARE ' " -- COMPLEX ---- "
    'ODD ARE ' " -- NORMAL ---- "
    Set F = FS.GETFILE(FFILEAR(r))
    F.Copy FILE31 + "\" + ORGTAGFILE(r)
    
    If r <= 2 Then
        Set F = FS.GETFILE(FFILEAR(r))
        F.Copy ScanPath.TxtPath.Text + "\" + ORGTAGFILE(r)
    End If

Next

'
'Shell "Explorer.exe /select, """ + FILE31 + "\" + ORGTAGFILE(2) + """", vbNormalFocus
'Shell "Explorer.exe /select, " + FFILEAR(2), vbNormalFocus


If OBJECTIVE_BRIEF <> "ALL DRIVE" And IRFAN_TO_DO_FROM_SEND_TO = False Then
    Shell "C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe """ + FFILEAR(4) + """", vbNormalFocus
    'End
End If

End Sub

Private Sub Label10_Click()
    'SEARCH
    'FROM SEND TO
    '---------------
    'Label10.Caption
    '----------------------------
    'REQUIRE RESTE AT EVERY BEGIN
    '----------------------------
    FILE_SCRIPT_COUNTER = 0
    IRFAN_TO_DO_FROM_SEND_TO = True
    Label5_Click

End Sub

Private Sub Label11_Click()

    Call MNU_IRFAN_08_EXPLORER_INI_Click

End Sub

Private Sub Label2_Click()
'Label2
ScanPath.Timer1.Enabled = True
Call Label1_Click
ScanPath.Timer1.Enabled = False
ScanPath.lblCount4 = "WORK COMPLETE"
End Sub

Private Sub Label3_Click()
    'Label3.CAPTION
    Me.WindowState = vbMinimized
    ScanPath.WindowState = vbNormal
    ScanPath.Show
    DoEvents
    ScanPath.Timer1.Enabled = True
    Call Label1_Click
    ScanPath.Timer1.Enabled = False
    ScanPath.lblCount4 = "WORK COMPLETE"
End Sub

Private Sub Label5_Click()
    
    'SCAN TO DO
    
    'Me.WindowState = vbMinimized

    'SEARCH
    '---------------
    'Label5.Caption
    '---------------
    MNU_IRFAN_01_Click
    '-----------------
    ScanPath.lblCount4 = "WORK COMPLETE"
    ScanPath.Timer1.Enabled = False

    Call XSCRIPT_CHUNK_DATE_UPDATE_SHOW_REQUIRE

End Sub

Private Sub Label6_Click()
    
    'SEARCH
    '---------------
    'Label6.Caption
    '---------------
    
    'On Error Resume Next
    
    If Label6.Caption = "IRFAN XSCRIPT PROCESS -- WORKING" Then Exit Sub
    
    QB = Label6.BackColor
    Label6.BackColor = QBColor(14)
    DoEvents
'    FILE13_SCRIPT_TEMP_FOLDER = "C:\TEMP\IRFAN_SLIDESHOW-" + GETCOMPUTERNAME + ".txt"
    
'    Call MNU_INI_BACK_NAUGHT_Click
    
'    Dim FF

    
'    ReDim HHAR01(30000)
'    HHC01 = 0
'    FF = FreeFile
'    Open "C:\TEMP\IRFAN_SLIDESHOW_X_SCRIPT.TXT" For Input As #FF
'        Do
'            Line Input #FF, LL1
'            If Trim(LL1) <> "" Then
'                HHC01 = HHC01 + 1
'                HHAR01(HHC01) = UCase(LL1)
'            End If
'        Loop Until EOF(FF)
'    Close #FF, FF1
'    ReDim Preserve HHAR01(HHC01)


Label6.Caption = "IRFAN XSCRIPT PROCESS -- WORKING"
Label6.Refresh
DoEvents


XSCRIPT_CHUNK_AND_PROCESS

Set F = FS.GETFILE(FILE13_SCRIPT_TEMP_FOLDER)

If BEFORE_XSCRIPT_COUNTER = AFTER_XSCRIPT_COUNTER Then STAT_SAME = " SAME - "
Label6.Caption = "IRFAN XSCRIPT PROCESS - FSIZE -" + Str(F.Size) + " - LINE COUNT -" + STAT_SAME + "BEFORE=" + Str(BEFORE_XSCRIPT_COUNTER) + " - AFTER=" + Str(AFTER_XSCRIPT_COUNTER) + " - TOTAL GONE=" + Str(BEFORE_XSCRIPT_COUNTER - AFTER_XSCRIPT_COUNTER)

'Sleep 200
Label6.BackColor = QB
   
    
End Sub

Private Sub Label7_Click()
'LABEL7.CAPTION
'Me.WindowState = vbMinimized
MNU_IRFAN_02_Click
End Sub

Private Sub Label8_Click()
'Me.WindowState = vbMinimized
MNU_IRFAN_03_NOTEPAD_Click
End Sub

Private Sub Label9_Click()
'LABEL9.CAPTION
'Me.WindowState = vbMinimized
Call MNU_INI_BACK_NAUGHT_Click
MNU_IRFAN_02_Click
End Sub

Private Sub MNU_000_SCAN_GIVE_Click()
MsgBox "NOT A OPTION HERE MENU UPDATE WITH MEDIA OBJECTIVE" + vbCrLf + "WHOLE DRIVE OR SUB DIRECTORY TO SCAN IS" + vbCrLf + AT$
End Sub



Private Sub MNU_ALL_DRIVE_SET_05_Click()
    On Error Resume Next
    For Each Control In LabDR
        OBJECTIVE_BRIEF = "ALL DRIVE"
        MCHAR = Mid(Control.Caption, 2, 1) + ":\"
        Err.Clear
        Set F = FS.GetDrive(MCHAR)
        If Err.Number = 0 Then
            If F.ISREADY = True Then
                ScanPath.TxtPath.Text = MCHAR
                AT$ = MCHAR
                ScanPath.Timer1.Enabled = True
                
                Call Label1_Click
                ScanPath.cmdScan_Click
            End If
        End If
    
    Next
    
    
    Shell "C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe """ + FFILEAR(4) + """", vbNormalFocus
    
    'End

    
End Sub

Private Sub MNU_IRFAN_03_NOTEPAD_Click()
'MNU_Irfan_03_NOTEPAD.Caption = "003 --- NOTEPAD -- C:\TEMP\IRFAN_SLIDESHOW.txt"
'FILE13_SCRIPT_TEMP_FOLDER = "C:\TEMP\IRFAN_SLIDESHOW-" + GETCOMPUTERNAME + ".txt"

'MNU_Irfan_03_NOTEPAD.Caption = "003 -- NOTEPAD C:\TEMP\IRFAN_SLIDESHOW.txt"
'Shell "C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe """ + FILE13_SCRIPT_TEMP_FOLDER + """", vbNormalNoFocus


FILE_PATH = "C:\PROGRAM FILES\NOTEPAD++\NOTEPAD++.EXE"
If Dir(FILE_PATH) <> "" Then FILE_PATH2 = FILE_PATH
FILE_PATH = "C:\PROGRAM FILES (X86)\NOTEPAD++\NOTEPAD++.EXE"
If Dir(FILE_PATH) <> "" Then FILE_PATH2 = FILE_PATH

Shell FILE_PATH2 + " " + FILE13_SCRIPT_TEMP_FOLDER, vbNormalNoFocus

End Sub

Private Sub MNU_IRFAN_04_Click()
If MNU_IRFAN_04.Checked = True Then MNU_IRFAN_04.Checked = False: Exit Sub
If MNU_IRFAN_04.Checked = False Then MNU_IRFAN_04.Checked = True
End Sub

Private Sub MNU_IRFAN_05_Click()
Shell "NOTEPAD ""C:\TEMP\IRFAN_SLIDESHOW_X_SCRIPT.TXT""", vbNormalFocus


End Sub

Private Sub MNU_IRFAN_08_EXPLORER_INI_Click()

Call MNU_INI_READ
' FILE13_SCRIPT_TEMP_FOLDER


'INI_StartIndex


'----------------------------------
'FIRST LOAD OF THIS VAR HAPPEN HERE
'----------------------------------
'FILE13_SCRIPT_TEMP_FOLDER = "C:\TEMP\IRFAN_SLIDESHOW-" + GetComputerName + ".txt"

If Dir(FILE13_SCRIPT_TEMP_FOLDER) <> "" Then
    Set F = FS.GETFILE(FILE13_SCRIPT_TEMP_FOLDER)

    '------------------------------------------------
    'DO GET LINE COUNT
    '------------------------------------------------
    'DO THIS IN CHUNK MODE AND COUNT VBCRLF - QUICKER
    '------------------------------------------------
    FF01 = FreeFile
    Open FILE13_SCRIPT_TEMP_FOLDER For Input As #FF01
    Do
        If Not (EOF(FF01)) Then Line Input #FF01, line3
        
        
        FILE_SCRIPT_COUNTER_V2 = FILE_SCRIPT_COUNTER_V2 + 1
        If Val(INI_StartIndex) = FILE_SCRIPT_COUNTER_V2 Then
            INI_FILE_AT_LINE = line3
            
            '---------------------------
            'THIS EXIT QUICK WHEN RESULT
            '---------------------------
            'NOT CHANGE RESULT HELD IN VAR -- FILE_SCRIPT_COUNTER
            'INSTEAD USE ALTERNATIVE VAR -- FILE_SCRIPT_COUNTER_V2
            '---------------------------
            Exit Do
        End If
        
    Loop Until EOF(FF01)
    Close #FF01

End If

Label11.Caption = "EXPLORER CURRENT INI FILE = " + INI_FILE_AT_LINE

Shell "EXPLORER /select, """ + INI_FILE_AT_LINE + """", vbNormalFocus


End Sub

Private Sub MNU_RERUN_Click()
Reset
UNLOAD_WITH_IRFAN_KILL = True

'SET ALL TIMERS IN ALL FORMS ENABLED=FALSE
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

'--------------------------------
'CARE TIMER'S WILL NEED RE-ENABLE
'--------------------------------
   
Dim Form As Form
For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next Form

DoEvents

Load Me


UNLOAD_WITH_IRFAN_KILL = False


End Sub

Private Sub MNU_VB_Click()
    
Dim FILESPEC ', TT As Long
FILESPEC = APP_PATH_VAR + "\" + App.EXEName + ".vbp"

'If IsIDE = False And Dir$(FILESPEC) <> "" Then
'    TT = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE  """ + FILESPEC + """", vbMinimizedNoFocus)
'
'    'RUN EXIT OPTION
'    'TT = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /runexit """ + FileSpec + """", vbMinimizedNoFocus)
'    End
'End If


'---------------------
VB_APP_PATH = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
If Dir(VB_APP_PATH) <> "" Then VB_APP_PATH2 = VB_APP_PATH
VB_APP_PATH = "C:\Program Files (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
If Dir(VB_APP_PATH) <> "" Then VB_APP_PATH2 = VB_APP_PATH


Shell VB_APP_PATH2 + " """ + FILESPEC + """", vbNormalFocus
End

End Sub

Public Sub A15()
'------------------------------------------
'FORM LAYOUT WITH THE DRIVES THAT ARE READY
'------------------------------------------
'COLOUR SET = Label16.BackColor + RGB(50, 50, 50)

On Error Resume Next

XC = 64
For Each Control In LabDR
    XC = XC + 1
    'Control.Visible = FALSE
    Control.AutoSize = False
    Control.Caption = "-" + Chr(XC) + "-"
    Control.AutoSize = True
    Control.Refresh
Next
DoEvents

For Each Control In LabDR
    If Control.Index Mod 2 = 0 Then
        Control.BackColor = &HC0FFC0
    End If
    MCHAR = Mid(Control.Caption, 2, 1) + ":\"
    Err.Clear
    Set F = FS.GetDrive(MCHAR)
            
    If F.ISREADY = False Or Err.Number > 0 Then
        Control.BackColor = Label16.BackColor + RGB(50, 50, 50)
    End If
    If Control.Width > XT Then XT = Control.Width
Next

XT = (Me.Width - 100) / 26
XD = 0
For Each Control In LabDR
    Control.Top = 0
    Control.Left = XD
    Control.Width = XT
    XD = XD + XT
    Control.Height = 400
    Control.Visible = True
Next


End Sub

Public Sub FORM_LAYOUT_WITH_DRIVES_READY()
'------------------------------------------
'FORM LAYOUT WITH THE DRIVES THAT ARE READY
'------------------------------------------
'COLOUR SET = Label4.BackColor


On Error Resume Next

XC = 64
For Each Control In LabDR
    XC = XC + 1
    'Control.Visible = FALSE
    Control.AutoSize = False
    Control.Caption = "-" + Chr(XC) + "-"
    Control.AutoSize = True
    Control.Refresh
Next
DoEvents

For Each Control In LabDR
    If Control.Index Mod 2 = 0 Then
        Control.BackColor = &HC0FFC0
    End If
    MCHAR = Mid(Control.Caption, 2, 1) + ":\"
    Err.Clear
    Set F = FS.GetDrive(MCHAR)
            
    If F.ISREADY = False Or Err.Number > 0 Then
        Control.BackColor = Label4.BackColor
    End If
    If Control.Width > XT Then XT = Control.Width
Next

XT = (Me.Width - 100) / 26
XD = 0
For Each Control In LabDR
    Control.Top = 0
    Control.Left = XD
    Control.Width = XT
    XD = XD + XT
    Control.Height = 400
    Control.Visible = True
Next


End Sub

Private Sub MNU_IRFAN_01_Click()
    
    ' --- ANOTHER MENU SET VAR SELECT DRIVE SET ADD THEN DO
    ' --- MNU_IRFAN_MODE_SET = True
    ' --- FILTER OUT VAR GATHERED
    
    
    '-------------------------------------------------
    'IF SCAN TO DO SAME AS LAST
    'DON'T RESET COUNT INI BACK NAUGHT
    '-------------------------------------------------
    'AND DON'T RESET OR DO ANYTHING IF ERROR READ INFO
    '-------------------------------------------------
    If DUPE_LAST_PATH2_ERROR_VAR = False Then
        If DUPE_LAST_PATH = False Then
            Call MNU_INI_BACK_NAUGHT_Click
        End If
    End If
    
    Dim FF

    On Error Resume Next
'    If IRFAN_TO_DO_FROM_SEND_TO     = False Then
    
'    FF = FreeFile
'    Open "C:\TEMP\IRFAN_SLIDESHOW_X_SCRIPT_LAST_RUN.TXT" For Input As #FF
'        Line Input #FF, LINE_TIME
'    Close #FF
'    LINE_TIME_VAR = DATEVAULE(LINE_TIME) + TimeValue(LINE_TIME)
'    IF
    
    
    '------------------------------------------------
    'LOAD THE XSCRIPT CHUNK FOR THE FILE SCAN PROCESS
    'FILTER FOR WHILE SCAN
    '------------------------------------------------
    Call X_SCRIPT_CHUNK

'    FF = FreeFile
'    Open "C:\TEMP\IRFAN_SLIDESHOW_X_SCRIPT_LAST_RUN.TXT" For Input As #FF
    
'    End If

    'ReDim Preserve HHAR01(HHC01)

    MNU_IRFAN_MODE_SET = True
    
    If Dir("D:\#0 IRFAN SHOW SCRIPT FILE SET", vbDirectory) = "" Then
        MkDir "D:\#0 IRFAN SHOW SCRIPT FILE SET"
    End If
    
    
    'Me.Hide
    ScanPath.Show
    ScanPath.WindowState = vbNormal
    
    i1 = ""
    i2 = ""
      
    ScanPath.cboMask.Text = "*.JPG;*.GIF"
    ScanPath.cboMask.Text = "*.JPG;*.GIF"
    ScanPath.cboMask.Text = "*.JPG;*.JPEG"
    'ScanPath.cboMask.Text = "*.JPG;*.GIF;*.MPG;*.AVI;*.MP4;*.3GP"
    'ScanPath.cboMask.Text = "*.MPG;*.AVI;*.MP4"
     
     
    i1 = i1 + "*.JPG|JPEG|JPE|GIF|BMP|DIB|TIF|TIFF|PNG|PCX|TGA|LBM|"
    If MNU_IRFAN_04.Checked = False Then
        i1 = i1 + "CLP|ANI|VOB|MPG|MP4|3GP|AVI|WMV|FLV"
    End If
'    i1 = "*.AVI|MP4|3GP"
    'i1 = i1 + "CLP|ANI|VOB|MPG|MP4|3GP|AVI|WMV"
    
    i1 = Replace(i1, "|", ";*.")
    ScanPath.cboMask.Text = i1

    'JPG|JPEG|JPE|GIF|BMP|DIB|TIF|TIFF|PNG|PCX|TGA|PCD|SUN|RAS|**ICO|AVI|IFF|PPM|PGM|PBM|LBM|WMF|EMF|WAV|PSD|MPG|MPE|MPEG|MID|RMI|MOV|CUR|ANI|DCX|EPS|CLP|CAM|G3|AIF|AU|SND|PSP|PSPIMAGE|ICL|SFW|KDC|RA|MP3|DCM|ACR|FPX|XBM|XPM|DJVU|SWF|FLV|IMG|IW44|WBMP|SGI|RGB|RLE|MED|RLE|SFF|NLM|NOL|NGG|GSM|JP2|JPC|J2K|JPF|FSH|CRW|B3D|WMA|WMV|TTF|SID|MNG|JNG|RAW|ECW|ASF|TXT|WAD|ICS|IDS|CDR|CMX|NEF|NRW|ERF|ORF|RAF|MRW|DCR|X3F|SRF|ARW|PEF|RW2|RWL|SRW|OGG|DWG|DXF|DDS|JPM|PLT|CR2|DNG|PS|PDF|AI|CGM|SVG|DCM|HDR|FITS|HDP|WDP|EXR|WSQ|JLS|WBZ|WBC|MP4|MKV|MTS|M2TS|TS|M2T|M4V|3GP|VOB|
    
    
    i1 = ""
    i1 = i1 + "*.JPG|JPEG|JPE|GIF|BMP|DIB|TIF|TIFF|TGA|PCD|SUN|"
    i1 = i1 + "RAS||PCX|PCDSUN|"
    i1 = i1 + "IFF|PPM|PGM|PBM|"
'    i1 = i1 + "WMF|EMF|WAV|PSD|"
    i1 = i1 + "MPE|"
    i1 = i1 + "CUR|DCX|EPS|"
    i1 = i1 + "G3|AIF|AU|SND|"
    i1 = i1 + "PSP|PSPIMAGE|ICL|SFW|KDC|"
    i1 = i1 + "DCM|ACR|FPX|XBM|XPM|DJVU|SWF|"
    i1 = i1 + "IMG|IW44|WBMP|SGI|RGB|RLE|MED|RLE|SFF|NLM|NOL|NGG|GSM|"
    i1 = i1 + "JP2|JPC|J2K|JPF|FSH|CRW|B3D|"
    i1 = i1 + "TTF|SID|MNG|JNG|"
    i1 = i1 + "X3F|SRF|ARW|PEF|RW2|RWL|SRW|"
    i1 = i1 + "DWG|DXF|DDS|JPM|PLT|"
    i1 = i1 + "PDF|AI|CGM|SVG|DCM|HDR|FITS|HDP|WDP|EXR|WSQ|"
    i1 = i1 + "JLS|WBZ|WBC"
    
'    i1 = i1 + ";" + Replace(i1, "|", ";*.")
    
'    ScanPath.cboMask.Text = i1
    '**ICO|
    'MID|RMI|
    'RA|MP3|
    'WMA|
    'RAW|
    'ECW|ASF|TXT|WAD|ICS|IDS|CDR|CMX|NEF|NRW|ERF|ORF|RAF|MRW|
    'DCR|
    'OGG|
    'CR2|DNG|PS|
    
    
    i2 = i2 + "AVI|RM|"
    i2 = i2 + "MPG|"
    i2 = i2 + "MPEG|"
    i2 = i2 + "MOV|"
    i2 = i2 + "CAM|"
    i2 = i2 + "WMV|FLV|FV|"
    i2 = i2 + "MP4|MKV|MTS|M2TS|TS|M2T|M4V|3GP|VOB"
    
    i2 = Replace(i1 + i2, "|", ";*.")
    
    
'    ScanPath.cboMask.Text = i2
    
    
    
    ScanPath.lblCount9.Caption = Replace(ScanPath.cboMask.Text, ";", "; ")

    

    VAR2 = Format(Now, "YYYY_MM_DD_HH_MM_SS") + ".txt"
    
    FILE11 = "D:\#0 IRFAN SHOW SCRIPT FILE SET\#00 IRFAN SCRIPT -- " + VAR2
    FILE12 = "D:\#0 IRFAN SHOW SCRIPT FILE SET\#00 IRFAN SCRIPT APPEND BUILD -- " + VAR2
    
    If Dir(FILE11) <> "" Then Kill FILE11
    If Dir(FILE12) <> "" Then Kill FILE12
    
        
    FF01 = FreeFile
    Open FILE11 For Output As #FF01

    FF02 = FreeFile ' LESS MODIFY
    Open FILE12 For Append As #FF02
    
    FF03 = -3
    FF04 = -4
    
    
    If IRFAN_TO_DO_FROM_SEND_TO = False Then
    
        MsgBox "SCAN FROM SEND TO HAS NOT A VALID PATH = SO ALL DRIVES WILL SCAN", vbMsgBoxSetForeground
        
    
        On Error Resume Next
        For Each Control In LabDR
        
            MCHAR = Mid(Control.Caption, 2, 1) + ":\"
            Err.Clear
            Set F = FS.GetDrive(MCHAR)
                    
            If Err.Number = 0 Then
                If F.ISREADY = True Then
                    ScanPath.TxtPath.Text = MCHAR
                    ScanPath.Timer1.Enabled = True
                    ScanPath.cmdScan_Click
                End If
            End If
        Next
    Else
        '--------------------------------------------------
        'THIS THE MAIN RUN IN - ALREADY GOT PATH ON SEND TO
        E$ = AT$
        'E$ = Command$
'        E$ = "C:\Program Files\# NO INSTALL REQUIRED\GOOGLEIMAGE-MEDIADOWNLOADER"
        If Mid$(E$, 1, 1) = """" Then E$ = Mid$(E$, 2)
    
        If Mid$(E$, Len(E$), 1) = """" Then E$ = Mid$(E$, 1, Len(E$) - 1)
        If FS.FileExists(E$) = True Then E$ = Mid$(E$, 1, InStrRev(E$, "\"))
        'ScanPath.Show
        DoEvents
        
        ScanPath.TxtPath.Text = E$
        ScanPath.Timer1.Enabled = True
        
        ScanPath.cmdScan_Click
    End If
    
    'Debug.Print ScanPath.ListView1.ListItems.Count
'    Print #FF01, "E:\MY WEBS\MATTHEWLAN.COM WEB\LOVEFOLDER\COOL\NHS\05 AGENT ORANGE.JPG"
'    Print #FF01, "E:\MY WEBS\MATTHEWLAN.COM WEB\LOVEFOLDER\COOL\NHS\05 AGENT ORANGE.JPG"
'    Print #FF02, "E:\MY WEBS\MATTHEWLAN.COM WEB\LOVEFOLDER\COOL\NHS\05 AGENT ORANGE.JPG"
'    Print #FF02, "E:\MY WEBS\MATTHEWLAN.COM WEB\LOVEFOLDER\COOL\NHS\05 AGENT ORANGE.JPG"
    
    Close FF01, FF02
    
'----------
'   EXAMPLE
'----------
'   FILE13_SCRIPT_TEMP_FOLDER = "C:\TEMP\IRFAN_SLIDESHOW-" + GETCOMPUTERNAME + ".txt"
'----------
'   EXAMPLE
'----------
'   FILE12 = "D:\#0 IRFAN SHOW SCRIPT FILE SET\#00 IRFAN SCRIPT APPEND BUILD -- " + VAR2
'----------
'   EXAMPLE
'----------
'   FILE14 = "C:\TEMP\IRFAN_SLIDESHOW-" + GetComputerName + "_BACKUP.txt"
'----------

    
    FS.COPYFILE FILE12, FILE13_SCRIPT_TEMP_FOLDER
    FILE14 = "C:\TEMP\IRFAN_SLIDESHOW-" + GetComputerName + "_BACKUP.txt"
    FS.COPYFILE FILE12, FILE14


'    If IRFAN_TO_DO_FROM_SEND_TO = False Then
'        'Shell "C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe """ + FILE13 + """", vbNormalFocus
'    End If
'    Call Label6_Click

    If DUPE_LAST_PATH2_ERROR_VAR = True Then

        Form_SEND_TO.Label10.Caption = Form_SEND_TO_Label10_VAR_HOLD + " - DONE - ERROR READ INFO X_SCRIPT IF TO DO OR NOT - GO IF WANT"
    
    Else
    
    
        If DUPE_LAST_PATH = True Then
            Form_SEND_TO.Label10.Caption = Form_SEND_TO_Label10_VAR_HOLD + " - DONE - CLICK X_SCRIPT AND THEN GO"
        Else
            
            Form_SEND_TO.Label10.Caption = Form_SEND_TO_Label10_VAR_HOLD + " - DONE - INI COUNT RESET WAS DONE"
        End If
        
    End If
    
'    Call LAUNCH_IRFANView

    '------------------------------------------------------------
    '-----------------------------
    'SET SOME RESULT VAR IN LABELS
    'AFTER THE SCAN
    '-----------------------------
    '------------------------
    'ONLY GET THE INI POINTER
    '------------------------
    Call MNU_INI_READ
    '------------------------------------------------------------
    'GET THE INI FILE AT POINTER -- AND UPDATE ALL LABELS ON FORM
    '------------------------------------------------------------
    Call Label0_Click
    '------------------------------------------------------------
    '------------------------------------------------------------

    
    ScanPath.Hide
    ScanPath.WindowState = vbNormal
    'Me.Visible = True
    'Me.WindowState = vbNormal


End Sub

Private Sub MNU_IRFAN_02_Click()
    
    'JOB CHECK AFTER XSCRIPT PROCESSED THAT COUNT IS UPDATE EVERYWHERE

'    PROBLEM
'    Call XSCRIPTCHUNK
'    Shell "C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe """ + FILE13_SCRIPT_TEMP_FOLDER + """", vbNormalFocus
    
    Call LAUNCH_IRFANView
End Sub

Sub GET_FOLDER_PATH_IRFANView()

    FOLDER_PATH_IRFANView = "C:\Program Files\IrfanView\i_view32.exe"
    If Dir(FOLDER_PATH_IRFANView) <> "" Then
        FOLDER_PATH_IRFANView2 = FOLDER_PATH_IRFANView
    End If
    FOLDER_PATH_IRFANView = "C:\Program Files (X86)\IrfanView\i_view32.exe"
    If Dir(FOLDER_PATH_IRFANView) <> "" Then
        FOLDER_PATH_IRFANView2 = FOLDER_PATH_IRFANView
    End If
    FOLDER_PATH_IRFANView = "C:\Program Files\IrfanView\i_view64.exe"
    If Dir(FOLDER_PATH_IRFANView) <> "" Then
        FOLDER_PATH_IRFANView2 = FOLDER_PATH_IRFANView
    End If
    

End Sub

Sub LAUNCH_IRFANView()
    
    'Me.WindowState = vbMinimized
    
    '-------------------------------------------------------------
    'IF FOLDER PATH HAS CHANGED - WANT TO RESET INI CONT TO NAUGHT
    'AND EXTRA CODE TO REMEMBER COUNT EACH FOLDER
    'GOT THE LAST FOLDER IS DUPE VAR SORTED
    'Dim DUPE_LAST_PATH
    '-------------------------------------------------------------

    Shell FOLDER_PATH_IRFANView + " /killmesoftly"
    Shell FOLDER_PATH_IRFANView + " /slideshow=""" + FILE13_SCRIPT_TEMP_FOLDER + """ /ONE /bf /silent /gray /RELOADONLOOP /QUITE", vbNormalFocus


    'JOB TO DO
    '--------------------------------------------------
    'GOT A CODE TO DO KILL ME SOFTLY AT UNLOD THIS FORM
    'MAYBE SOMETHING MORE
    'AND ESCAPE KEY EXIT TO THAT BETTER
    '--------------------------------------------------

    If IRFAN_TO_DO_FROM_SEND_TO = True Then
'        End
    End If


End Sub

Sub FLAG_TO_DO_XSCRIPT_02()
    
    FILE19X_SCRIPT_DATE_MODIFIED = "C:\TEMP\IRFAN_SLIDESHOW_X_SCRIPT_DATE_MODIFIED.TXT"
    FLAG_TO_DO_XSCRIPT = True

    If Dir(FILE13_SCRIPT_TEMP_FOLDER) <> "" Then
        Set F = FS.GETFILE(FILE13_SCRIPT_TEMP_FOLDER)
        ADATE = F.DateLastModified
        If Dir(FILE19X_SCRIPT_DATE_MODIFIED) <> "" Then
            FF = FreeFile
            Open FILE19X_SCRIPT_DATE_MODIFIED For Input As #FF
                Line Input #FF, LL1
            Close FF
            LL1_DATE_VAR = DateValue(LL1) + TimeValue(LL1)
            If LL1_DATE_VAR <> ADATE Then
                FLAG_TO_DO_XSCRIPT = True
            Else
                FLAG_TO_DO_XSCRIPT = False
            End If
        End If
    End If

End Sub

Sub XSCRIPT_CHUNK_DATE_UPDATE_SHOW_REQUIRE()
    
    Call FLAG_TO_DO_XSCRIPT_02

    If FLAG_TO_DO_XSCRIPT = False Then
        'SEARCH
        '---------------
        'Label6.Caption
        '---------------
        Label6.Caption = "IRFAN XSCRIPT PROCESS -- NOT REQUIRE -- BUT IF MANUAL EDIT -- YES MAYBE DO"
    Else
        Label6.Caption = "IRFAN XSCRIPT PROCESS -- REQUIRED MAYBE -- IS INCLUDED DURING SCAN ANYWAY - MAYBE IF MODIFIED BY EDIT"
        Label6.BackColor = QBColor(14)

    End If


End Sub
    



Sub XSCRIPT_CHUNK_AND_PROCESS()
    
    '-----------------------------------------------
    'JOB
    '-----------------------------------------------
    'WANT FILE DATE MODIFIED SO NOT DO TWICE IF EVER
    '-----------------------------------------------
    
    
    X_SCRIPT_CHUNK
    
    FILE17 = "C:\TEMP\IRFAN_SLIDESHOW-A-Z.TXT"
    FILE18 = FILE13_SCRIPT_TEMP_FOLDER
    
    Call XSCRIPT_CHUNK_DATE_UPDATE_SHOW_REQUIRE
    
    If Dir(FILE18) = "" Then
        MsgBox "NOT EXIST FILE" + vbCrLf + FILE18, vbMsgBoxSetForeground
        Exit Sub
    End If
    '---------------------
    'SUN 29 MAY 2K SIXTEEN
    '---------------------
    'MASSARGE BOX --
    'LUNCH BOX
    'JOCK STRAP
    'VIBRATOR IN HE BOX
    'BIG BOX LITTLE BOX
    
    FS.COPYFILE FILE18, FILE17
    
    FF01 = FreeFile
    Open FILE17 For Binary As #FF01
    FF02 = FreeFile
    Open FILE18 For Output As #FF02
    
    
    '------------------------------------------
    'HERE A XSCRIPT IS RUN
    'WHEN IS RUN
    'AND ALSO ONE FILTER DURING SCAN FROM CHUNK
    '------------------------------------------
    
    FILE_SCRIPT_COUNTER_VAR = FILE_SCRIPT_COUNTER
    
    BEFORE_XSCRIPT_COUNTER = 0
    AFTER_XSCRIPT_COUNTER = 0
    Do
        If MOD_VAR Mod 100 = 0 Or MOD_VAR = 0 Then
            Label6.Caption = "IRFAN XSCRIPT PROCESS -- WORKING --" + Str(MOD_VAR) + " OF" + Str(FILE_SCRIPT_COUNTER_VAR)
            Label6.Refresh
            DoEvents
        End If
        MOD_VAR = MOD_VAR + 1
        
        Line Input #FF01, FILESPEC
        STRING_SIZE = STRING_SIZE + Len(FILESPEC + vbCrLf)
        If FILESPEC <> OFILESPEC Then
            OFILESPEC = FILESPEC 'INTERLACE - REMOVE DUPE FILE
            
            BEFORE_XSCRIPT_COUNTER = BEFORE_XSCRIPT_COUNTER + 1
            
            Call XSCRIPT(FILESPEC) ' THROW IN FF02 - FREEFILE
        End If
    Loop Until EOF(FF01) Or STRING_SIZE >= LOF(FF01)
    Close FF01, FF02

    '--------------------------------------------------------------
    'FORGET THIS ONE AT MOMENT IT RESET INI COUNTER IF ANY CHANGING
    'DONE REMOVE BY X_SCRIPT
    '--------------------------------------------------------------
    'If WORKFLAG = True Then MNU_INI_BACK_NAUGHT_Click

    FILE_SCRIPT_COUNTER = AFTER_XSCRIPT_COUNTER

    If Dir(FILE13_SCRIPT_TEMP_FOLDER) <> "" Then
        Set F = FS.GETFILE(FILE13_SCRIPT_TEMP_FOLDER)
        ADATE = F.DateLastModified
        FF = FreeFile
        Open FILE19X_SCRIPT_DATE_MODIFIED For Output As #FF
            Print #FF, ADATE
        Close FF
        FLAG_TO_DO_XSCRIPT = False
    End If

 
    
    Label5.Caption = "IRFAN -- " + FILE13_SCRIPT_TEMP_FOLDER + " -- FSIZE -" + Str(F.Size) + " -- LINE COUNT =" + Str(FILE_SCRIPT_COUNTER)
    'Label6.Caption = "Irfan XSCRIPT PROCESS - FSIZE --" + Str(F.Size) + "LINE COUNT - BEFORE=" + Str(BEFORE_XSCRIPT_COUNTER) + " - AFTER=" + Str(AFTER_XSCRIPT_COUNTER)


End Sub
Sub X_SCRIPT_CHUNK()

    'PROBLEM
    
    'STRIPPER DONT WANT FROM EXTENTION
    'STRIPPER DO WANT FROM EXTENTION
    
    Dim FF

    On Error Resume Next
    
    ReDim HHAR01(30000)
    ReDim HHAR02(30000)
    HHC01 = 0
    HHC02 = -1
    FF = FreeFile
    Open "C:\TEMP\IRFAN_SLIDESHOW_X_SCRIPT.TXT" For Input As #FF
        Do
            Line Input #FF, LL1
            If Trim(LL1) <> "" Then
                
                'WHY SECOND STORE ONLY WANT ONE
                'SUPPOSE WE WANT TO BE GRABBING SOME FIRST
                If InStr(LL1, "### SET X DO INCLUDE") > 0 Or FLAG = "SET1" Then
                    FLAG = "SET1"
                    HHC02 = HHC02 + 1
                    If HHC02 > 0 Then
                        HHAR02(HHC02) = UCase(LL1)
                    End If
                End If
                
                If InStr(LL1, "### SET -- DONT INCLUDE IF THESE ARE ON THE SEND_TO COMMAND LINE") > 0 Or FLAG = "SET2" Then
                    FLAG = "SET2"
                    If InStr(AT$, LL1) > 0 Then
                        HHC02 = HHC02 + 1
                        If HHC02 > 0 Then
                            HHAR02(HHC02) = UCase(LL1)
                        End If
                    End If
                End If
            
                
                If Mid(FLAG, 1, 3) <> "SET" Then
                    HHC01 = HHC01 + 1
                    HHAR01(HHC01) = UCase(LL1)
                End If
            
            End If
        Loop Until EOF(FF)
    Close #FF, FF1
    
    ReDim Preserve HHAR01(HHC01)
    ReDim Preserve HHAR02(HHC02)
    
End Sub



Public Sub XSCRIPT(FILESPEC)
    '-----------------------------------
    'THE XSCRIPT IS RUN DURING SCAN ALSO
    'SO MAY NOT SEE RESULT UPON RUN
    '-----------------------------------


'   On Error GoTo EXITSUB
'   If Irfan_TO_DO_FROM_SEND_TO = True Then TAGX = False: GoTo WRITEANDEXIT
    TAGX = False
    
    For R5 = 1 To UBound(HHAR01)
        'Debug.Print HHAR01(R5)
        If InStr("*" + UCase(FILESPEC), HHAR01(R5)) > 0 Then
            TAGX = True
            WORKFLAG = True
            Exit For
        End If
    Next
    
    
    If TAGX = True Then
        For R5 = 1 To UBound(HHAR02)
            ' -- EXTENTION WANT CHECK
            
            ' -- ALL EXTENTION WANT CHECK
            
            If Mid(HHAR02(R5), Len(HHAR02(R5)), 1) = "\" Then
                 i5 = Mid(FILESPEC, 4, InStrRev(FILESPEC, "\") - 3)
                    
'                    If Len(i5) = Len(Mid(HHAR02(R5), 5 + 1)) Then Stop 'SWAP DIR SAME LENGHT
                    
                    'If Len(i5) = Len(Mid(HHAR02(R5), 5 + 1)) Then Stop ' NOT A CHECK WITH ROOT AT MOMENT
              
            End If
                
                
            'FOR THE IF ON THE SEND TO
            If InStr(UCase(FILESPEC), UCase(HHAR02(R5))) > 0 Then
            ' -- FOLDER WANT ---
                TAGX = False    '  --- ALLOW OBJECTIVE
                WORKFLAG = True
                Exit For
            End If
                
                
                
'Len(HHAR02(R5))

            If Mid(HHAR02(R5), 1, 4) = Mid(UCase(FILESPEC), Len(FILESPEC) - 3) _
             Or Mid(HHAR02(R5), 1, 1) = "*" Then
            
            ' -- FOLDER WANT ---
            If InStr("*" + UCase(FILESPEC), Mid(HHAR02(R5), 5 + 1)) > 0 Then
                    TAGX = False    '  --- ALLOW OBJECTIVE
                    WORKFLAG = True
                    Exit For
                End If
            End If
        Next
    End If
    
WRITEANDEXIT:
    If TAGX = False Then
        Print #FF02, FILESPEC
        AFTER_XSCRIPT_COUNTER = AFTER_XSCRIPT_COUNTER + 1
    End If
    
End Sub




Private Sub MNU_INI_BACK_NAUGHT_Click()

Call MNU_INI_READ

If IsNumeric(INI_StartIndex) Then
    If RC_INI <> "" Then
    
    FR = FreeFile
    Open FILEPATH_I_VIEW_INI For Output Lock Read Write As #FR
        Print #FR, RC_INI
    Close #FR
    
    End If
End If

If IsNumeric(INI_StartIndex) = False Then
    MsgBox "INI COUNT VALUE IS NOT IN IS_NUMERIC " + vbCrLf + "-- VALUE = -->" + INI_StartIndex + "<--"
End If

If RC_INI = "" Then
    MsgBox "INI IS NOT LOADED WITH STRING DATA TO DO A WRITE YET"
End If

End Sub


Private Sub MNU_INI_READ()

FILEPATH2 = ""
FILE_PATH = "C:\Program Files\IrfanView\i_view32.ini"
If Dir(FILE_PATH) <> "" Then FILEPATH2 = FILE_PATH
FILE_PATH = "C:\Program Files (X86)\IrfanView\i_view32.ini"
If Dir(FILE_PATH) <> "" Then FILEPATH2 = FILE_PATH
FILE_PATH = "C:\Program Files\IrfanView\i_view64.ini"
If Dir(FILE_PATH) <> "" Then FILEPATH2 = FILE_PATH

If FILEPATH2 = "" Then
    FILE_PATH1 = "C:\Program Files\IrfanView\i_view32.ini"
    FILE_PATH2 = "C:\Program Files (X86)\IrfanView\i_view32.ini"
    FILE_PATH3 = "C:\Program Files\IrfanView\i_view64.ini"
    MsgBox "NOT FIND " + vbCrLf + FILE_PATH1 + vbCrLf + " OR " + vbCrLf + FILE_PATH2 + vbCrLf + " OR " + vbCrLf + FILE_PATH3 + vbCrLf, vbMsgBoxSetForeground
    
    Exit Sub
End If

FR = FreeFile
Open FILEPATH2 For Input As #FR
    RC_INI = Input(LOF(FR), FR)
Close #FR

'Form2_ANY_SPECIAL_FOLDER.GetSpecialfolder_Debug

FILEPATH_I_VIEW_INI = FILEPATH2

'Debug.Print RC_INI

If InStr(RC_INI, "INI_Folder=%APPDATA%\IrfanView") > 0 Then
    X_SP_F = GetSpecialfolder(26) + "\IrfanView\"
    X_SP_F_i_view_ini = X_SP_F + Dir(X_SP_F + "\i_view*.ini")
    FR = FreeFile
    Open X_SP_F_i_view_ini For Input As #FR
        RC_INI = Input(LOF(FR), FR)
    Close #FR
    FILEPATH_I_VIEW_INI = X_SP_F_i_view_ini
End If

'Debug.Print RC_INI
    
INI_StartIndex1 = InStr(RC_INI, "StartIndex=") + Len("StartIndex=")
If INI_StartIndex1 = 0 Then
    INI_StartIndex = "INI COUNT INFO NOT FOUND"
    Exit Sub
End If
INI_StartIndex2 = InStr(INI_StartIndex1, RC_INI, vbCrLf)
INI_StartIndex = Mid(RC_INI, INI_StartIndex1, INI_StartIndex2 - INI_StartIndex1)

'AYE
'A1 = INI_StartIndex
'INIA1 = A1


RC_INI = Replace(RC_INI, "StartIndex=" + INI_StartIndex, "StartIndex=0")



'FR = FreeFile
'Open FILEPATH2 For Output Lock Read Write As #FR
'    Print #FR, RC
'Close #FR

End Sub

Private Sub MNU_VB_FOLDER_Click()
    
    Shell "Explorer.exe /e," + APP_PATH_VAR, vbNormalFocus

End Sub








