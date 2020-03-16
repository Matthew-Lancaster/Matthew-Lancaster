VERSION 5.00
Begin VB.Form Form_SEND_TO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008080&
   Caption         =   "SEND TO SCRIPT SUB FOLDER FILES SET"
   ClientHeight    =   8364
   ClientLeft      =   228
   ClientTop       =   888
   ClientWidth     =   12480
   Icon            =   "FORM_SEND_TO.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8364
   ScaleWidth      =   12480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7080
      Top             =   2292
   End
   Begin VB.Timer TIMER_SET_FORM_LAYOUT 
      Interval        =   1
      Left            =   7080
      Top             =   1884
   End
   Begin VB.Timer Timer_1_MINUTE 
      Interval        =   1000
      Left            =   7080
      Top             =   1548
   End
   Begin VB.ListBox List1 
      Height          =   240
      Left            =   9000
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "MPC MEDIA PLAYER CLASSIC VIDEO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   912
      Left            =   612
      TabIndex        =   38
      Top             =   4032
      Width           =   5652
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "EXPLORER CURRENT INI FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   600
      TabIndex        =   37
      Top             =   5760
      Width           =   5652
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAN SCAN SENDTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Left            =   600
      TabIndex        =   36
      Top             =   600
      Width           =   5652
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAN GO FROM INI NUMERIC ONE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   912
      Left            =   600
      TabIndex        =   35
      Top             =   3000
      Width           =   5652
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAN EDIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Left            =   600
      TabIndex        =   34
      Top             =   5160
      Width           =   5652
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAN GO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Left            =   600
      TabIndex        =   33
      Top             =   2400
      Width           =   5652
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAN XSCRIPT PROCESS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Left            =   600
      TabIndex        =   32
      Top             =   1800
      Width           =   5652
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAN SCAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   456
      Left            =   600
      TabIndex        =   31
      Top             =   1200
      Width           =   5652
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "Label4"
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   5280
      TabIndex        =   30
      Top             =   6480
      Visible         =   0   'False
      Width           =   612
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
      Height          =   336
      Index           =   0
      Left            =   3960
      TabIndex        =   29
      Top             =   6600
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   1
      Left            =   3480
      TabIndex        =   28
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   2
      Left            =   3480
      TabIndex        =   27
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   3
      Left            =   3720
      TabIndex        =   26
      Top             =   6468
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   4
      Left            =   3480
      TabIndex        =   25
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   5
      Left            =   3720
      TabIndex        =   24
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   6
      Left            =   3720
      TabIndex        =   23
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   7
      Left            =   3960
      TabIndex        =   22
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   8
      Left            =   3600
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   9
      Left            =   3840
      TabIndex        =   20
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   10
      Left            =   3840
      TabIndex        =   19
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   11
      Left            =   4080
      TabIndex        =   18
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   12
      Left            =   3840
      TabIndex        =   17
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   13
      Left            =   4080
      TabIndex        =   16
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   14
      Left            =   4080
      TabIndex        =   15
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   15
      Left            =   4440
      TabIndex        =   14
      Top             =   6600
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   16
      Left            =   3600
      TabIndex        =   13
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   17
      Left            =   3720
      TabIndex        =   12
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   18
      Left            =   3720
      TabIndex        =   11
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   19
      Left            =   3960
      TabIndex        =   10
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   20
      Left            =   3720
      TabIndex        =   9
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   21
      Left            =   3960
      TabIndex        =   8
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   22
      Left            =   3960
      TabIndex        =   7
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   23
      Left            =   4080
      TabIndex        =   6
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   24
      Left            =   3840
      TabIndex        =   5
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
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
      Height          =   336
      Index           =   25
      Left            =   4080
      TabIndex        =   4
      Top             =   6480
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "000 SCAN TO GIVE "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   600
      TabIndex        =   3
      Top             =   -12
      Width           =   5652
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "OBJECTIVE FILE ONLY DIR ONLY CRC OPTION "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   600
      TabIndex        =   2
      Top             =   7560
      Width           =   5700
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "BEGIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   1200
      TabIndex        =   1
      Top             =   6720
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER __"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu MNU_FOLDERING 
      Caption         =   "FOLDERING"
      NegotiatePosition=   1  'Left
      Begin VB.Menu MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC_2 
         Caption         =   "NOT USE XSCRIPT FOR MPC WORKER"
      End
      Begin VB.Menu MNU_FOLDERING_ARRAY 
         Caption         =   "C DOWNLOAD"
         Index           =   1
      End
      Begin VB.Menu MNU_FOLDERING_ARRAY 
         Caption         =   "D DOWNLOAD"
         Index           =   2
      End
      Begin VB.Menu MNU_FOLDERING_ARRAY 
         Caption         =   "DSC"
         Index           =   3
      End
      Begin VB.Menu MNU_FOLDERING_ARRAY 
         Caption         =   "DOWNLOAD AND GROUP MPC - NOT DSC"
         Index           =   4
      End
      Begin VB.Menu MNU_FOLDERING_ARRAY 
         Caption         =   "DOWNLOAD AND GROUP MPC"
         Index           =   5
      End
      Begin VB.Menu MNU_FOLDERING_ARRAY 
         Caption         =   "01 MM"
         Index           =   6
      End
      Begin VB.Menu MNU_FOLDERING_ARRAY 
         Caption         =   "01 MM NETWORK"
         Index           =   7
      End
      Begin VB.Menu MNU_FOLDERING_ARRAY 
         Caption         =   "02 MM 4TB"
         Index           =   8
      End
      Begin VB.Menu MNU_FOLDERING_ARRAY 
         Caption         =   "02 MM 4TB NETWORK"
         Index           =   9
      End
      Begin VB.Menu MNU_FOLDERING_ARRAY 
         Caption         =   "NOT ME"
         Index           =   10
      End
      Begin VB.Menu MNU_FOLDERING_ARRAY 
         Caption         =   "NOT ME NETWORK"
         Index           =   11
      End
      Begin VB.Menu MNU_FOLDERING_ARRAY 
         Caption         =   "SNAPSHOT HIKVISION GARDEN"
         Index           =   12
      End
      Begin VB.Menu MNU_FOLDERING_ARRAY 
         Caption         =   "SNAPSHOT HIKVISION GARDEN NETWORK SOURCE"
         Index           =   13
      End
      Begin VB.Menu MNU_SHOW_ALL_MENU_OFF 
         Caption         =   "SHOW ALL MENU OFF"
      End
   End
   Begin VB.Menu MNU_C_DRIVE_MPC 
      Caption         =   "DOWNLOAD AND GROUP MPC"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_SORT_ON_DATE_TRUE 
      Caption         =   "SORT ON DATE = TRUE"
   End
   Begin VB.Menu MNU_I_VIEW64_EXE 
      Caption         =   "I_VIEW.EXE"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu MNU_XNVIEW_EXE 
      Caption         =   "XNVIEW_EXE"
      NegotiatePosition=   1  'Left
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
      Begin VB.Menu MNU_SHOW_ALL_MENU 
         Caption         =   "008 --- SHOW ALL MENU"
      End
   End
   Begin VB.Menu MNU_00 
      Caption         =   "IRFAN"
      NegotiatePosition=   1  'Left
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
         Checked         =   -1  'True
      End
      Begin VB.Menu MNU_IRFAN_05 
         Caption         =   "005 --- EDIT XSCRIPT"
      End
      Begin VB.Menu MNU_IRFAN_07_NOT_USE_XSCRIPT 
         Caption         =   "007 --- NOT USE AUTO XSCRIPT"
      End
      Begin VB.Menu MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC 
         Caption         =   "008 --- NOT USE XSCRIPT FOR MPC WORKING"
      End
      Begin VB.Menu MNU_IRFAN_09_EXPLORER_INI 
         Caption         =   "009 --- EXPLORER FOLDER + SELECT FILE FROM CUR POS INI"
      End
   End
   Begin VB.Menu MNU_RERUN 
      Caption         =   "* RERUN *"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu MNU_DRIVE_VOL_NAME 
      Caption         =   "DRIVE_VOL_NAME"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu MNU_INFO_HELP 
      Caption         =   "INFO"
   End
   Begin VB.Menu MNU_ERROR_INFO 
      Caption         =   "ERROR INFO"
      Visible         =   0   'False
      Begin VB.Menu MNU_ERROR_INFO_1 
         Caption         =   "1.."
      End
      Begin VB.Menu MNU_ERROR_INFO_2 
         Caption         =   "2.."
      End
   End
End
Attribute VB_Name = "Form_SEND_TO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------------------------------------
' WORK DONE ON
' ----------------------------------------------------------------------------
' [ Sunday 10:32:30 Am_24 February 2019 ]
' ----------------------------------------------------------------------------
' REASON NEW CODE TO MAKE A BUTTON TO SWITCH
' FROM SHOW ALL MENU TO SHOW ALL MENU OFF
'
' A LOT MORE ON THAT ONE
'
' MAKE MENU FOR SHOW ALL MENU PROPERLY
' LINK MENU ABOUT XSCRIPT DUPLICATE INTO OTHER MENU
' MAKE CODE BETTER FOR SHOW ALL MENU UPTO A
' CERTAIN LEVEL NUMBER DETECT BY STRING RATHER THAN NUMBER INDEX
' MAKE MSI NOT START AS SHOW ALL MENU
' MAKE FORM STILL SHOW VBMAXIMIZED IF XSRIPT HAS REDUCER TO NONE FILE COUNT
' ADD SOME FOLDER TO XSCRIPT
' ADD SOME REM LINE INCLUDING IN XSCRIPT
' INFO TO FIND ABOUT HWO XSCRIPT DEAL WITH WILD CARD * _ DONE
' WORK TO DO WOULD LIKE IT IF TELL XSCRIPT IS GOING TO BE USED
' IN A QUICKER WAY FIRST MINI SCAN QUICKER _ LATER
' TIDY UP SOME MINOR CHANGE HAPPEN BUGGY HARDCORE
' USE MORE CHECKED MENU ITEM
' ASK QUESTION AT BEGIN IF HARDCODED OPTION TO USE WITH MAKE DRAW SCREEN FIRST
' SORT ON DATE TRUE FOR EVERYTHING
' ----------------------------------------------------------------------------


'------------------------------
'HARD CODE PUT A PATH IN TESTER
'------------------------------
'Label3.Caption = "SCAN TO GIVE -- " + AT$

Public MNU_SORT_ON_DATE_TRUE_REQUEST_MENU

Dim USE_ISIDE_HARDCODER_PATH_BEEN_SET_DONE


'----------------------------------------------------------
'SEARCH FIND THE EXTENSION SET USE FOR MEDIA PLAYER CLASSIC
'----------------------------------------------------------
Dim O_MhWndVAR
Dim P_COUNT
Dim Old_LINE_STRING
Dim FOLDER_PATH_IRFANView2
Dim OMe_WindowState
Dim OMe_WindowState_2

Dim SP_FOLDER_COUNT

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
Dim I
Dim i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

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



Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Const WM_CLOSE = &H10






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


Private Sub Form_Load()

'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
If App.PrevInstance = True Then End
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
Call VB_PROJECT_CHECKDATE
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
Dim reg_valuename, WShell, Cmd, CmdLine, I
GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv").EnumValues &H80000003, "S-1-5-19\Environment", reg_valuename
If IsArray(reg_valuename) <> 0 Then
    RequireAdmin = 1
End If

DISPLAY_MENU = False
If Dir(App.Path + "\" + GetComputerName + "-- SETTING PLUG.TXT") <> "" Then
    FR = FreeFile
    Open App.Path + "\" + GetComputerName + "-- SETTING PLUG.TXT" For Input As #FR
        Line Input #FR, LINE_VAR
        ' Print #FR, Format(Now, "HH:MM:SS DD/MM/YYYY")
    Close #FR
    
    DISPLAY_MENU = True
    If DateValue(LINE_VAR) + TimeValue(LINE_VAR) < Now Then
        Kill App.Path + "\" + GetComputerName + "-- SETTING PLUG.TXT"
        DISPLAY_MENU = False
    End If
End If

' FIRST MENU ITEM OF NAME THAT AFTER IS NOT DISPLAY WHEN SHOW ALL MENU OFF
' 01 MM
For r = 1 To MNU_FOLDERING_ARRAY.Count
    If MNU_FOLDERING_ARRAY.Item(r).Caption = "01 MM" Then
        VAR_MNU_FOLDERING_ARRAY_COUNT_IN = r
    End If
Next

For r = VAR_MNU_FOLDERING_ARRAY_COUNT_IN To MNU_FOLDERING_ARRAY.Count
    MNU_FOLDERING_ARRAY.Item(r).Visible = DISPLAY_MENU
Next

For r = 1 To MNU_FOLDERING_ARRAY.Count
    If InStr(MNU_FOLDERING_ARRAY.Item(r).Caption, "SNAPSHOT") > 0 Then
        MNU_FOLDERING_ARRAY.Item(r).Visible = True
    End If
Next

If MNU_FOLDERING_ARRAY.Item(VAR_MNU_FOLDERING_ARRAY_COUNT_IN).Visible = True Then
    MNU_SHOW_ALL_MENU.Checked = True
    MNU_SHOW_ALL_MENU_OFF.Checked = True
    MNU_SHOW_ALL_MENU_OFF.Visible = True
    MNU_SHOW_ALL_MENU_OFF.Caption = "SHOW ALL MENU ON"
Else
    MNU_SHOW_ALL_MENU.Checked = False
    MNU_SHOW_ALL_MENU_OFF.Checked = False
    MNU_SHOW_ALL_MENU_OFF.Visible = False
    MNU_SHOW_ALL_MENU_OFF.Caption = "SHOW ALL MENU OFF"
End If



If IsIDE = False Then
    If RequireAdmin = 0 Then
        I_TEXT = I_TEXT + "Set UAC = CreateObject(""SHELL.APPLICATION"")" + vbCrLf
        I_TEXT = I_TEXT + "UAC.ShellExecute """ + App.Path + "\" + App.EXEName + """, """", """", ""RUNAS"", 1" + vbCrLf
                
        VBSCRIPT_PATH = App.Path + "\RELOAD WITH ADMINISTRATOR RIGHTS.VBS"
        FR1 = FreeFile
        Open VBSCRIPT_PATH For Output As #FR1
            Print #FR1, I_TEXT
        Close #FR1
        
        Dim WSHShell
        Set WSHShell = CreateObject("WScript.Shell")
            WSHShell.Run """" + VBSCRIPT_PATH + """"
        Set WSHShell = Nothing
        'Unload Me
        End
    End If
End If
If IsIDE = True Then
    If RequireAdmin <> 1 Then
        MsgBox "Visual Basic IDE Not in Administrator Rights Mode"
        End
    End If
End If

IRFAN_TO_DO_FROM_SEND_TO = False

Set FS = CreateObject("Scripting.FileSystemObject")

ScanPath.Visible = False

Me.WindowState = vbMaximized

If GetComputerName = "4-ASUS-GL522VW" Then
    Call MNU_SHOW_ALL_MENU_OFF_CLICK
End If
If GetComputerName = "7-ASUS-GL522VW" Then
    MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC.Checked = True
    MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC_2.Checked = True
End If
If GetComputerName = "8-MSI-GP62M-7RD" Then
    Call MNU_SHOW_ALL_MENU_OFF_CLICK
End If
'If GetComputerName = "7-ASUS-GL522VW" Then
'    Call MNU_SHOW_ALL_MENU_OFF_CLICK
'End If


'THE SCAN TO BUTTON IRFANVIEW
'------------------------------------------------------------
Label10.BackColor = QBColor(14)

'01 OF 01
'FOLDER OF PROGRAM
Call GET_FOLDER_PATH_IRFANView

'----------------------------------
'FIRST LOAD OF THIS VAR HAPPEN HERE
'----------------------------------
FILE13_SCRIPT_TEMP_FOLDER = "C:\TEMP\IRFAN_SLIDESHOW-" + GetComputerName + ".TXT"

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
MNU_IRFAN_03_NOTEPAD.Caption = "003 --- NOTEPAD++ -- C:\TEMP\IRFAN_SLIDESHOW-" + GetComputerName + ".TXT"
Label8.Caption = "IRFAN EDIT -- NOTEPAD++ -- C:\TEMP\IRFAN_SLIDESHOW-" + GetComputerName + ".TXT"
'---------------------------

'------------------------
'ONLY GET THE INI POINTER
'------------------------
Call MNU_INI_READ


Me.Show
DoEvents


Call GIVE_DRIVE_COMMAND_STRING

'------------------------------------------------------------
'GET THE INI FILE AT POINTER -- AND UPDATE ALL LABELS ON FORM
'------------------------------------------------------------
Call Label0_Click


'------------------------------------------------------------
'MPC_MEDIA_PLAYER_SCAN
'HAS THE DATA SEARCH EXTENSION SET
'------------------------------------------------------------

'If Command$ <> "" Then MNU_IRFAN_02_Click
'Call Test4



Exit Sub
End




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
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------



'-------------------------------------------------
'-------------------------------------------------
'Private Sub Timer_1_MINUTE_Timer()
'
'Timer_1_MINUTE.Interval = 60000
'
'Call VB_PROJECT_CHECKDATE
'
'End Sub
'-------------------------------------------------
'-------------------------------------------------
Sub VB_PROJECT_CHECKDATE()

'TOP DELCLARE -------------
'DIM XVB_DATE
'-----------------------------------

' Exit Sub

If Mid(App.Path, 1, 1) <> "D" Then Exit Sub

PATH_FILE_NAME1 = App.Path + "\" + App.EXEName + ".EXE"
PATH_FILE_NAME2 = Replace(PATH_FILE_NAME1, "D:\VB6\", "D:\VB6-EXE\")

'---------------------------------------------------
'IF A NEW PROJECT NOT BEEN SYNC YET TO MIRROR FOLDER
'---------------------------------------------------
If Dir(PATH_FILE_NAME2) = "" Then
    On Error Resume Next
        If Dir(Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\")), vbDirectory) = "" Then
            CreateFolderTree Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\"))
        End If
        FS.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
    If Err.Number > 0 Then Exit Sub
End If

Set FS = CreateObject("Scripting.FileSystemObject")
Set F = FS.GetFile(PATH_FILE_NAME1)
XVB_DATE = F.DateLastModified
Set F = FS.GetFile(PATH_FILE_NAME2)
VB_DATE = F.DateLastModified

'----------------------------------
'WRITE INFO ABOUT DATE CHANGE NEWER
'----------------------------------

If XVB_DATE < VB_DATE And XVB_DATE > 0 Then

    '----------------------------------
    'Mon 10 April 2017 16:30:50--------
    '----------------------------------
    'PROJECT REFERANCE ----------------
    'wshom.ocx
    'IF HARD FIND DO BROWSER
    'IN SYSTEM32 AND SYSWOW64
    'DIR wshom.ocx /S
    '----------------------------------
    'WINDOWS SCRIPT HOST OBJECT MODEL
    'AFTER FIND ALSO HAVE TO SELECT HER
    '----------------------------------
    '----------------------------------
    'Mon 10 April 2017 15:45:11--------
    '----
    'VISUAL BASIC wshom.ocx NOT WORKING - Google Search
    'https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB721GB721&q=VISUAL+BASIC+wshom.ocx+NOT+WORKING&oq=VIUAL+BASIC+wshom.ocx+NOT+WORKING&gs_l=serp.3..30i10k1.1533.3987.0.4658.12.12.0.0.0.0.127.888.11j1.12.0....0...1c.1.64.serp..0.12.886...33i160k1j33i21k1.fYCPCpVPeVk
    '--------
    'WScript reference in VB
    'http://forums.codeguru.com/showthread.php?30458-WScript-reference-in-VB
    '----
    'Set objShell = WScript.CreateObject("WScript.Shell")
    '----------------------------------
    
    I_TEXT = "Dim FSO" + vbCrLf
    I_TEXT = I_TEXT + "Set FSO = CreateObject(""SCRIPTING.FILESYSTEMOBJECT"")" + vbCrLf
    I_TEXT = I_TEXT + "FSO.CopyFile """ + PATH_FILE_NAME2 + """,""" + Mid(PATH_FILE_NAME1, 1, InStrRev(PATH_FILE_NAME1, "\")) + """" + vbCrLf
    I_TEXT = I_TEXT + "Set UAC = CreateObject(""SHELL.APPLICATION"")" + vbCrLf
    I_TEXT = I_TEXT + "UAC.ShellExecute """ + App.Path + "\" + App.EXEName + """, """", """", ""RUNAS"", 1" + vbCrLf
            
    FR1 = FreeFile
    Open App.Path + "\RELOAD AND COPY.VBS" For Output As #FR1
    Print #FR1, I_TEXT
    Close #FR1
    
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + App.Path + "\RELOAD AND COPY.VBS" + """"
    Set WSHShell = Nothing
    'Unload Me
    End
End If

XVB_DATE = VB_DATE

End Sub

Function RequireAdmin2()
    Dim reg_valuename, WShell, Cmd, CmdLine, I

    GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv").EnumValues &H80000003, "S-1-5-19\Environment", reg_valuename
    If IsArray(reg_valuename) <> 0 Then
        RequireAdmin2 = 1
        Exit Function
    End If

    If IsIDE = True Then RequireAdmin2 = 0: Exit Function

    'Set Cmd = WScript.Arguments
    'For Each Arg In Split(Command$, " ")
    '    If Arg = "/admin" Then
    '        WScript.Echo "To script you must have administrator rights!"
    '        'RequireAdmin = 0
    '        'Exit Function
    '        WScript.Quit
    '    End If
    '    CmdLine = CmdLine & Chr(32) & Chr(34) & Cmd(I) & Chr(34)
    'Next
    'CmdLine = CmdLine & Chr(32) & Chr(34) & "/admin" & Chr(34)

    'Set WShell = WScript.CreateObject("WScript.Shell")
    'CreateObject("Shell.Application").ShellExecute WShell.ExpandEnvironmentStrings("%SystemRoot%\System32\WScript.exe"), Chr(34) & WScript.ScriptFullName & Chr(34) & CmdLine, "", "runas"
    'WScript.Quit
End Function


Sub Test2()

Dim FS, Idate As Date, Hdate As String, File1 As String, FILE2 As String
Set FS = CreateObject("Scripting.FileSystemObject")

Call GetFileDates

File1 = Mid$(List1.List(List1.ListCount - 1), 21)
For r = List1.ListCount - 2 To 0 Step -1
FILE2 = Mid$(List1.List(r), 21)
If Mid$(List1.List(List1.ListCount - 1), 1, 19) <> Mid$(List1.List(r), 1, 19) Then
FS.CopyFile File1, FILE2
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
FileSpec = Filespec2
Idate = DateSerial(1980, 1, 1)
If FS.FileExists(FileSpec) Then
Set F = FS.GetFile(FileSpec)
Idate = F.DateLastModified
End If
Hdate = Format$(Idate, "YYYY-MM-DD HH:MM:SS")


End Sub


Public Function FindFileSize(Filename As String) As Long
        On Error Resume Next
        Dim filedata As WIN32_FIND_DATA

        filedata = FindFile(Filename)
        
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
Dim I As Integer
i2 = 1
For I = lDataSize - 1 To 1 Step -2
'For i = 1 To lDataSize Step 2
    'bytearray(i2) = Val("&h" + Mid$(gg3, i, 2))
    ll = ll + Chr(Val("&h" + Mid$(gg3, I, 2)))
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

Private Sub MNU_C_DRIVE_MPC_Click()
Call MNU_FOLDERING_ARRAY_Click(4)
End Sub

Private Sub MNU_INFO_HELP_Click()
I = "While Infar Is Running" + vbCrLf + "And If Move Mouse To Top And Hold Left Button" + vbCrLf + "A Short While" + vbCrLf + "It Will Kill Irfan" + vbCrLf
I = I + vbCrLf
I = I + "It Does A Nice Kill User Close Becuase If KillMeSoftly By Own Program Irfan And A Picture Has Delay To Load Knocks Out The Task_Bar" + vbCrLf
I = I + "Detect Fullscreenclass And Then The Parent And Done"
MsgBox I, vbMsgBoxSetForeground
End Sub

Private Sub MNU_IRFAN_07_NOT_USE_XSCRIPT_Click()

If MNU_IRFAN_07_NOT_USE_XSCRIPT.Checked = False Then MNU_IRFAN_07_NOT_USE_XSCRIPT.Checked = True: Exit Sub
If MNU_IRFAN_07_NOT_USE_XSCRIPT.Checked = True Then MNU_IRFAN_07_NOT_USE_XSCRIPT.Checked = False


End Sub

Private Sub MNU_SHOW_ALL_MENU_OFF_CLICK()

' FIRST MENU ITEM OF NAME THAT AFTER IS NOT DISPLAY WHEN SHOW ALL MENU OFF
' 01 MM
For r = 1 To MNU_FOLDERING_ARRAY.Count
    If MNU_FOLDERING_ARRAY.Item(r).Caption = "01 MM" Then
        VAR_MNU_FOLDERING_ARRAY_COUNT_IN = r
    End If
Next

For r = VAR_MNU_FOLDERING_ARRAY_COUNT_IN To MNU_FOLDERING_ARRAY.Count
    MNU_FOLDERING_ARRAY.Item(r).Visible = False
Next

For r = 1 To MNU_FOLDERING_ARRAY.Count
    If InStr(MNU_FOLDERING_ARRAY.Item(r).Caption, "SNAPSHOT") > 0 Then
        MNU_FOLDERING_ARRAY.Item(r).Visible = True
    End If
Next

If MNU_FOLDERING_ARRAY.Item(VAR_MNU_FOLDERING_ARRAY_COUNT_IN).Visible = True Then
    MNU_SHOW_ALL_MENU.Checked = True
    MNU_SHOW_ALL_MENU_OFF.Checked = True
    MNU_SHOW_ALL_MENU_OFF.Visible = True
    MNU_SHOW_ALL_MENU_OFF.Caption = "SHOW ALL MENU ON"
Else
    MNU_SHOW_ALL_MENU.Checked = False
    MNU_SHOW_ALL_MENU_OFF.Checked = False
    MNU_SHOW_ALL_MENU_OFF.Visible = False
    MNU_SHOW_ALL_MENU_OFF.Caption = "SHOW ALL MENU OFF"
End If

MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC.Checked = False
MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC_2.Checked = False
MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC_2.Caption = "HardCORE"

If Dir(App.Path + "\" + GetComputerName + "-- SETTING PLUG.TXT") <> "" Then
    Kill App.Path + "\" + GetComputerName + "-- SETTING PLUG.TXT"
End If


End Sub

Private Sub MNU_SHOW_ALL_MENU_Click()


' FIRST MENU ITEM OF NAME THAT AFTER IS NOT DISPLAY WHEN SHOW ALL MENU OFF
' 01 MM
For r = 1 To MNU_FOLDERING_ARRAY.Count
    If MNU_FOLDERING_ARRAY.Item(r).Caption = "01 MM" Then
        VAR_MNU_FOLDERING_ARRAY_COUNT_IN = r
    End If
Next

For r = VAR_MNU_FOLDERING_ARRAY_COUNT_IN To MNU_FOLDERING_ARRAY.Count
    MNU_FOLDERING_ARRAY.Item(r).Visible = Not MNU_FOLDERING_ARRAY.Item(r).Visible
Next

For r = 1 To MNU_FOLDERING_ARRAY.Count
    If InStr(MNU_FOLDERING_ARRAY.Item(r).Caption, "SNAPSHOT") > 0 Then
        MNU_FOLDERING_ARRAY.Item(r).Visible = True
    End If
Next

If MNU_FOLDERING_ARRAY.Item(VAR_MNU_FOLDERING_ARRAY_COUNT_IN).Visible = True Then
    MNU_SHOW_ALL_MENU.Checked = True
    MNU_SHOW_ALL_MENU_OFF.Checked = True
    MNU_SHOW_ALL_MENU_OFF.Visible = True
    MNU_SHOW_ALL_MENU_OFF.Caption = "SHOW ALL MENU ON"
Else
    MNU_SHOW_ALL_MENU.Checked = False
    MNU_SHOW_ALL_MENU_OFF.Checked = False
    MNU_SHOW_ALL_MENU_OFF.Visible = False
    MNU_SHOW_ALL_MENU_OFF.Caption = "SHOW ALL MENU OFF"
End If



FR = FreeFile
Open App.Path + "\" + GetComputerName + "-- SETTING PLUG.TXT" For Output As #FR
    Print #FR, Format(Now + 1, "HH:MM:SS DD/MM/YYYY")
Close #FR

End Sub

Private Sub MNU_SORT_ON_DATE_TRUE_Click()

' IF THE REQUEST HAS BEEN CHANGD THEN IGNOR
' OKAY IGNOR THIS ONE
Form_SEND_TO.MNU_SORT_ON_DATE_TRUE_REQUEST_MENU = True

'FIND HERE 2
If MNU_SORT_ON_DATE_TRUE.Caption = "SORT ON DATE = TRUE" Then
    MNU_SORT_ON_DATE_TRUE.Caption = "SORT ON DATE = NOT"
    Exit Sub
End If
If MNU_SORT_ON_DATE_TRUE.Caption = "SORT ON DATE = NOT" Then
    MNU_SORT_ON_DATE_TRUE.Caption = "SORT ON DATE = TRUE"
    Exit Sub
End If

End Sub

Sub TIMER_SET_FORM_LAYOUT_TIMER()
TIMER_SET_FORM_LAYOUT.Enabled = False

On Error Resume Next
    If Me.WindowState = vbNormal Then
        Me.Width = Screen.Width / 2 '+ 2000 'Label2.Width - 2000
        Me.Height = Label2.Height + 4500
    End If
On Error GoTo 0


On Error Resume Next
    If Me.WindowState = vbNormal Then
        Me.Height = Label2.Top + Label2.Height + 1000
    End If
On Error GoTo 0

On Error Resume Next
    If Me.WindowState = vbNormal Then
        'Me.Top = 1000
        'Me.Left = 2400
        'Me.Width = 100
    End If
On Error GoTo 0



End Sub

Private Sub Form_Resize()

'Me.Top = 20
'Me.Left = 100


'-------------------
'FORM LAYOUT TO DO
'-------------------
'OMe_WindowState_2 = OMe_WindowState_2 + 1
If OMe_WindowState = Me.WindowState Then Exit Sub
'If OMe_WindowState <> Me.WindowState Then OMe_WindowState_2 = 0
OMe_WindowState = Me.WindowState

TIMER_SET_FORM_LAYOUT.Enabled = True


Exit Sub


X = 1
Y = 1
On Error Resume Next
For Each Control In Controls
        'CONTROL.NAME
    If Control.Enabled = True Then
        MNU_2 = 0
        If InStr(UCase(Control.Name), "MNU_") > 0 Then MNU = 1: MNU_2 = 1
        If MNU_2 = 0 Then
            If Control.Width + Control.Left > X Then X = Control.Width + Control.Left
            If Control.Height + Control.Top > Y Then Y = Control.Height + Control.Top
        End If
    End If
Next
On Error GoTo 0


ScanPath.Width = X + 110
If MNU = 1 Then
    pluso = 720
Else
    pluso = 450
End If
ScanPath.Height = Y + pluso

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

'If Dir(FOLDER_PATH_IRFANView) = "" Then MsgBox "EXE FILE NOT EXIST" + vbCrLf + vbCrLf + FOLDER_PATH_IRFANView

If UNLOAD_WITH_IRFAN_KILL = False Then

    Shell FOLDER_PATH_IRFANView2 + " /killmesoftly"

End If

Dim Control

'SET ALL TIMERS IN ALL FORMS ENABLED=FALSE
On Error Resume Next
    For I = 0 To Forms.Count - 1
        For Each Control In Forms(I).Controls
            If InStr(UCase(Control.Name), "TIMER") > 0 Then
                'Debug.Print Control.Name
                Control.Enabled = False
            End If
        Next
    Next I
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
    Set F = FS.GetFile(FILE13_SCRIPT_TEMP_FOLDER)

    '------------------------------------------------
    'DO GET LINE COUNT
    '------------------------------------------------
    'DO THIS IN CHUNK MODE AND COUNT VBCRLF - QUICKER
    '------------------------------------------------
    FILE_SCRIPT_COUNTER = 0
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

Call SORT_ON_DATE_CHECKER



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


Call SET_INI_NUMBER_IN_LABEL_TWO


'-------------------
'LABEL CAPTIONS DONE
'-------------------

'-------------------
'FORM LAYOUT TO DO
'-------------------

On Error Resume Next
    If Me.WindowState = vbNormal Then
        Me.Width = Label2.Width + 1000
        Me.Height = Label2.Height + 4500
    End If
On Error GoTo 0

DoEvents


For Each Control In Controls
    If Mid(Control.Name, 1, 5) = "Label" Then
        Control.Left = 0
        Control.Width = Me.Width
    End If
Next




ReDim RAF(70) 'AMPLE
I = 0
'i = i + 1: RAF(i) = "3"
    I = I + 1: RAF(I) = "5"
    I = I + 1: RAF(I) = "10"
    I = I + 1: RAF(I) = "6"
    I = I + 1: RAF(I) = "7"
    I = I + 1: RAF(I) = "9"
    I = I + 1: RAF(I) = "8"
    I = I + 1: RAF(I) = "12"
    I = I + 1: RAF(I) = "11"
    I = I + 1: RAF(I) = "2"
ReDim Preserve RAF(I)
'Label12 MEDIA PLAYER


Label3.Top = 550
HDIFF = 125
Label3.Height = 450
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
    If Me.WindowState = vbNormal Then
        Me.Height = Label2.Top + Label2.Height + 1000
    End If
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
    If Me.WindowState = vbNormal Then
        Me.Top = 1000
        Me.Left = 2400
    End If
On Error GoTo 0


'------------------------------------------
'FORM LAYOUT WITH THE DRIVES THAT ARE READY
'------------------------------------------
Call FORM_LAYOUT_WITH_DRIVES_READY

'--------------------------------
'SET SOME HIGHLIGHT COLOUR YELLOW
'--------------------------------
Label7.BackColor = QBColor(14)
Label11.BackColor = QBColor(14)
Label12.BackColor = QBColor(14)

End Sub

Sub SET_INI_NUMBER_IN_LABEL_TWO()

'REQUIRE TO CREAT AN INI FILE WITH IRFANVIEW FIRST
'BRING TEXT PATH STRING BACK IF NOT SETTER YET
'-------------------------------------------------
If IsNumeric(INI_StartIndex) = False Then Exit Sub
'-------------------------------------------------

If INI_StartIndex = "" Then INI_StartIndex = "0"
If FILE_SCRIPT_COUNTER = "" Then FILE_SCRIPT_COUNTER = "0"
If FLAG_TO_DO_XSCRIPT = True Then
    Label7.Caption = "IRFAN GO - FROM INI #" + Str(INI_StartIndex) + " Of" + Str(FILE_SCRIPT_COUNTER) + " -- INI COUNT WILL RESET TO #1 AT RUN - FOLDER CHANGE HAPPEN"
Else
    Label7.Caption = "IRFAN GO - FROM INI #" + Str(INI_StartIndex) + " Of" + Str(FILE_SCRIPT_COUNTER)
End If

Label9.Caption = "IRFAN GO - RESET INI TO NUMERIC #1 ONE -- FROM INI #" + Str(INI_StartIndex) + " Of" + Str(FILE_SCRIPT_COUNTER)

Label11.Caption = "EXPLORER CURRENT INI FILE = " + INI_FILE_AT_LINE

End Sub


Sub GIVE_DRIVE_COMMAND_STRING()

DUPE_LAST_PATH2_ERROR_VAR = False
'FILEPATH_APP_TMP1 = App.Path + "\Last Folder.tmp"
'If Dir(FILEPATH_APP_TMP1) <> "" Then FILEPATH_APP_TMP = FILEPATH_APP_TMP1

APP_PATH_VAR = App.Path + "\IRFANVIEW -- " + GetComputerName
If Dir(APP_PATH_VAR, vbDirectory) = "" Then
    MkDir APP_PATH_VAR
End If


FILEPATH_APP_TMP1 = APP_PATH_VAR + "\Last Folder --" + GetComputerName + "--.TXT"
FILEPATH_APP_TMP = FILEPATH_APP_TMP1

'If Dir(FILEPATH_APP_TMP, vbDirectory) = "" Then

'End If

'COMMAND$
'CLIPBOARD
'REMEMBER FILE

'THE SENDTO COMMAND$ WILL BE NOTHING IF ADMINISTRATOR IS NOT SET
'MsgBox Command$

AT$ = Command$
'AT$ = "D:\Videos\NOT\BUNKER"

'AT$ = "D:\DSC\2015-2018\2016 CyberShot HX60V\DCIM"



AT$ = Replace(AT$, """", "")

AT_LOADER = AT$

' TAKE THE FIRST ONE TO CHECK VALID PATH
If InStr(AT_LOADER, ";") > 0 Then
    AT_LOADER = Mid(AT$, 1, InStr(AT_LOADER, ";") - 1)
End If

'If IsIDE = True Then AT$ = "\\4-asus-gl522vw\4-asus-g_asus_g_d\DSC\2015 2016\2016 Sony CyberShot DSC-HX60V"
'If IsIDE = True Then AT$ = "\\4-asus-gl522vw\4-asus-g_asus_g_d\DSC\2015 2016"
ATISIDE$ = "D:\DSC"


If FS.FolderExists(AT_LOADER) = False Then
'    MsgBox "COMMAND$ FOLDER DONT EXIST" + vbCrLf + Command$
End If


'--------------------------
'FIND COMMANDLINE - SEND TO
'--------------------------
'If FS.FolderExists(AT$) = False Then
'    AT_TEMP$ = Clipboard.GetText
    If FS.FolderExists(AT_LOADER) = True Then
'        AT$ = AT_TEMP$
        If Form_SEND_TO.Label10.Caption = "IRFAN SCAN SENDTO" Then
            Form_SEND_TO.Label10.Caption = "IRFAN SCAN SENDTO - FROM -- SEND TO -- COMMAND_LINE"
        End If
    End If
'End If

'---------------------
'NONE SEND TO CMD LINE -- THEN
'IS CLIPBOARD ANY
'---------------------
If FS.FolderExists(AT_LOADER) = False Then
    AT_TEMP$ = Clipboard.GetText
    If FS.FolderExists(AT_TEMP$) = True Then
        Form_SEND_TO.Label10.Caption = "IRFAN SCAN SENDTO - FROM CLIPBOARD"
        AT$ = AT_TEMP$
    End If
End If
    
'--------------------------
'STILL NONE AFTER CLIPBOARD
'OR
'SEND TO
'USE REMEMBER FILE
'--------------------------
If FS.FolderExists(AT_LOADER) = False Then
    
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
If FS.FolderExists(AT_LOADER) = False Then
    If IsIDE = True Then
    '------------------------------------------------------------
        If USE_ISIDE_HARDCODER_PATH_BEEN_SET_DONE = False Then
            
            
            'GET THE INI FILE AT POINTER -- AND UPDATE ALL LABELS ON FORM
            '------------------------------------------------------------
            Call Label0_Click
    
            I = MsgBox("USE ISIDE HARDCODER PATH" + vbCrLf + ATISIDE$, vbMsgBoxSetForeground + vbYesNo)
            If I = vbYes Then
                Form_SEND_TO.Label10.Caption = "IRFAN SCAN SENDTO - FROM HARDCODE ISIDE"
                AT$ = ATISIDE$
                USE_ISIDE = True
            End If
            USE_ISIDE_HARDCODER_PATH_BEEN_SET_DONE = True
        End If
    End If
End If
    
    
'----------------------------------
If FS.FolderExists(AT_LOADER) = True Then

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
    
    FILEPATH_APP_TMP_COPY = Replace(FILEPATH_APP_TMP, "--.TXT", "--BACKCOPY.TXT")
    
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
            FS.CopyFile FILEPATH_APP_TMP, FILEPATH_APP_TMP_COPY
        End If
    End If
End If
    
If FS.FolderExists(AT_LOADER) = False Then
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

If Len(AT$) > 2 Then
    If Mid(AT$, 1, 2) = "\\" Then
        MNU_DRIVE_VOL_NAME.Caption = "DRIVE VOLUME NAME = __ " + Mid(AT$, 1, InStr(3, AT$, "\")) + " __ NETWORK"
    Else
        Set F = FS.GetDrive(Mid(AT$, 1, 3))
        MNU_DRIVE_VOL_NAME.Caption = "DRIVE VOLUME NAME = __ " + F.VolumeName
    End If
End If



'If IsIDE = True Then
    'AT$ = "M:\#\#D\00 Pen Drives MOBILES"
'    OBJECTIVE_BRIEF = "IDE MODE SET  " + AT$
    'AT$ = Clipboard.GetText

'End If

'If IsIDE = True Then AT$ = "C:\TEMP\" + Dir("C:\TEMP\*.*")

End Sub

Private Sub LabDR_Click(index As Integer)

If LabDR(index).BackColor = Label4.BackColor Then Exit Sub
AT$ = Chr(index + 65) + ":\"

Label3.Caption = "SCAN TO GIVE -- " + AT$

Set FS = CreateObject("Scripting.FileSystemObject")

Set F = FS.GetDrive(AT$)
MNU_DRIVE_VOL_NAME.Caption = "DRIVE VOLUME NAME = __ " + F.VolumeName


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

FILE11 = "#00 File List DIR & SUB HERE -- " + Format(Now, "YYYY_MM_DD_HH_MM_SS") + " -- COMPLEX -- " + FILE5TXT + ".TXT"
FILE12 = "#00 File List DIR & SUB HERE -- " + Format(Now, "YYYY_MM_DD_HH_MM_SS") + " -- NORMAL ---- " + FILE5TXT + ".TXT"
FILE13 = "#00 File List DIR & SUB HERE -- " + Format(Now, "YYYY_MM_DD_HH_MM_SS") + " -- COMPLEX -- " + FILE5TXT + "- DIRECTORY.TXT"
FILE14 = "#00 File List DIR & SUB HERE -- " + Format(Now, "YYYY_MM_DD_HH_MM_SS") + " -- NORMAL ---- " + FILE5TXT + "- DIRECTORY.TXT"

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
DC = DC + 1: DAT1(DC) = "Serial Number:        " & F.SerialNumber
i1 = Hex$(F.SerialNumber)
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
I = Format(DateDiff("s", PROCESSTIMEDIFF, INOW) / 60, "0.0##")
DC = DC + 1: DAT1(DC) = "Process Lenght Time:  " + I + " Minute Divide Clock"
I = Trim(Str(DateDiff("s", PROCESSTIMEDIFF, INOW)))
DC = DC + 1: DAT1(DC) = "Process Lenght Time:  " + I + " -- Seconds"
DC = DC + 1: DAT1(DC) = "BEGIN TIME:           " + Format(PROCESSTIMEDIFF, "DD MMM YYYY HH:MM:SS") + "h"
DC = DC + 1: DAT1(DC) = "AFTER TIME:           " + Format(INOW, "DD MMM YYYY HH:MM:SS") + "h" + " -" + Str(I) + " Sec"
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
    Open FFILEAR(R1) + ".TXT" For Output As #FF02
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
    Name FFILEAR(R1) + ".TXT" As FFILEAR(R1)
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
    Set F = FS.GetFile(FFILEAR(r))
    F.Copy FILE31 + "\" + ORGTAGFILE(r)
    
    If r <= 2 Then
        Set F = FS.GetFile(FFILEAR(r))
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

    Shell FOLDER_PATH_IRFANView2 + " /killmesoftly"
    
    Label11.BackColor = QBColor(14)
    Call MNU_IRFAN_09_EXPLORER_INI_Click

End Sub

Private Sub Label12_Click()

    If Label12.Caption = "MPC MEDIA PLAYER CLASSIC VIDEO ____ FILE SCANNED IS EMPTY NOT FOUND ANY VIDEO FILE __ PRESS AGAIN HERE TO END" Then
        End
    End If

    'SCAN TO DO
    
    'Me.WindowState = vbMinimized

    '---------------
    'SEARCH
    '---------------
    'MNU_IRFAN_01_Click
    '-----------------
    
    'If AT$ = "" OR Then
    
    'End If
        
    'Me.Hide
    'ScanPath.Show
    'ScanPath.WindowState = vbNormal
    'ScanPath.WindowState = vbMinimized
    
    ScanPath.Show
    ScanPath.WindowState = vbMaximized
    ScanPath.ListView1.Visible = True
    DoEvents
    
    On Error Resume Next
    ScanPath.Top = 0
    ScanPath.Left = 0
    On Error GoTo 0
    DoEvents

        
    PATH_PARSE = AT$
    Me.WindowState = vbMinimized
    
    Do
    
        PATH_PARSE_2 = ""
        If InStr(PATH_PARSE, ";") > 0 Then
            PATH_PARSE_2 = Mid(PATH_PARSE, 1, InStr(PATH_PARSE, ";") - 1)
            PATH_PARSE = Replace(PATH_PARSE, PATH_PARSE_2 + ";", "")
        Else
            PATH_PARSE_2 = Mid(PATH_PARSE, 1, Len(PATH_PARSE))
            PATH_PARSE = Replace(PATH_PARSE, PATH_PARSE_2, "")
        End If
        
        
        ' Debug.Print PATH_PARSE_2
        
        If PATH_PARSE_2 <> "" Then

            AT$ = PATH_PARSE_2
            
            If Dir(AT$, vbDirectory) = "" Then MsgBox "ERROR WITH GIVEN PATH" + vbCrLf + vbCrLf + AT$: End
            
            'AT$ = "D:\Videos\NOT\ME"
            
            'MsgBox "GIVEN PATH" + vbCrLf + vbCrLf + AT$: End
            
            
            
            RUN_FROM_MEDIA_PLAYER_CLASSIC = True
            
            'WHY SET CAUSER TO X-SCRIPT SOME AWAY _ NOT FOR MPC WANTING
            MNU_IRFAN_MODE_SET = False
            
            'Me.WindowState = vbNormal
        '    Me.Top = 200
        '    Me.Left = 100
            
            Call ScanPath.TIMER_SET_FORM_LAYOUT_2_TIMER
            
            Debug.Print Time$
            Debug.Print Time$
            Debug.Print Time$ + " START TIME ____"
            Debug.Print Time$
            Debug.Print Time$
            
            '------------------------------------------
            Call MPC_MEDIA_PLAYER_SCAN
            '------------------------------------------
        End If
    
    Loop Until PATH_PARSE = ""
    
    ScanPath.Hide
    ScanPath.WindowState = vbNormal
    'Me.Visible = True
    'Me.WindowState = vbNormal
    
    
    Debug.Print Time$
    Debug.Print Time$
    Debug.Print Time$ + " END TIME ____"
    Debug.Print Time$
    Debug.Print Time$
    
    
    ScanPath.lblCount4 = "WORK COMPLETE"
    ScanPath.Timer1.Enabled = False

    
    'BLACKLIST BLOCKER SCRIPT 01 of 02
    'BLACKLIST BLOCKER SCRIPT 02 OF 02
    
    'SEARCH
    
    OKAY_GO = 1
    If MNU_IRFAN_07_NOT_USE_XSCRIPT.Checked = True Then OKAY_GO = 0
    If MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC.Checked = True Then OKAY_GO = 0
    
    If OKAY_GO = 1 Then
        '-------------------------------------------------------------------------------
        Call XSCRIPT_CHUNK_AND_PROCESS_MPC
        Call Label6_Click
        '-------------------------------------------------------------------------------
        'Label6.Caption = "IRFAN XSCRIPT PROCESS -- WORKING"
        '-------------------------------------------------------------------------------
    
    End If
    
    

    NOT_USE_XSCRIPT = True
    'MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC.Checked = False
    'If NOT_USE_XSCRIPT = False Then
    If MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC.Checked = False Then

        'WRITE THE OUTPUT IF CHANGE HAPPEN FOR MPC AND OTHER
        '---------------------------------------------------
        Call FLAG_TO_DO_XSCRIPT_WRITE
        
        Call XSCRIPT_CHUNK_DATE_UPDATE_SHOW_REQUIRE

    End If

Set FS = CreateObject("Scripting.FileSystemObject")

'-------------------------------------------------------
'FILE13_SCRIPT_TEMP_FOLDER

RENAME_COPY_FILE = Replace(FILE13_SCRIPT_TEMP_FOLDER, ".TXT", ".M3U")

Reset

On Error Resume Next
If Dir(RENAME_COPY_FILE) <> "" Then Kill RENAME_COPY_FILE
Err.Clear
'-------------------------------------------------
'ON WINDOWS XP
'METHOD COPYFILE FAILED
'MAKE SURE BREAK ON UNHANDLED ERROR GOING PROPERLY
'-------------------------------------------------

FS.CopyFile FILE13_SCRIPT_TEMP_FOLDER, RENAME_COPY_FILE

If Err.Number > 0 Then
    On Error GoTo 0
    FR1 = FreeFile
    Open FILE13_SCRIPT_TEMP_FOLDER For Input As #FR1
    FR2 = FreeFile
    Open RENAME_COPY_FILE For Output As #FR2
    Do
        Line Input #FR1, LINE_1$
        Print #FR2, LINE_1$
    Loop Until EOF(FR1)
    Close FR1, FR2
End If

I = 0
If Dir(RENAME_COPY_FILE) <> "" Then
    FR1 = FreeFile
    Open RENAME_COPY_FILE For Input As #FR1
    If LOF(FR1) = 0 Then
        I = 1
    End If
    Close #FR1
End If

If I = 1 Then
    Label12.Caption = "MPC MEDIA PLAYER CLASSIC VIDEO ____ FILE SCANNED IS EMPTY NOT FOUND ANY VIDEO FILE __ PRESS AGAIN HERE TO END"
    Me.WindowState = vbMaximized
    Exit Sub
End If

'Shell "C:\Program Files (x86)\K-Lite Codec Pack\MPC-HC64\mpc-hc64_nvo.exe" + " Clear", vbMaximizedFocus
'Shell "C:\Program Files (x86)\K-Lite Codec Pack\MPC-HC64\mpc-hc64_nvo.exe" + " LOAD """ + RENAME_COPY_FILE + """", vbMaximizedFocus
Beep
Dim EXELOADER

If EXELOADER = "" And Dir("C:\Program Files (x86)\K-Lite Codec Pack\MPC-HC64\mpc-hc64.exe") <> "" Then
    EXELOADER = "C:\Program Files (x86)\K-Lite Codec Pack\MPC-HC64\mpc-hc64.exe"
End If

If EXELOADER = "" And Dir("C:\Program Files\MPC-HC\mpc-hc64.exe") <> "" Then
    EXELOADER = "C:\Program Files\MPC-HC\mpc-hc64.exe"
End If

If EXELOADER = "" And Dir("C:\Program Files (x86)\K-Lite Codec Pack\MPC-HC64\mpc-hc64_nvo.exe") <> "" Then
    EXELOADER = "C:\Program Files (x86)\K-Lite Codec Pack\MPC-HC64\mpc-hc64_nvo.exe"
End If

If EXELOADER = "" And Dir("C:\Program Files\K-Lite Codec Pack\MPC-HC64\mpc-hc64_nvo.exe") <> "" Then
    EXELOADER = "C:\Program Files\K-Lite Codec Pack\MPC-HC64\mpc-hc64_nvo.exe"
End If

If EXELOADER = "" And Dir("C:\Program Files\MPC-HC\mpc-hc64.exe") <> "" Then
    EXELOADER = "C:\Program Files\MPC-HC\mpc-hc64.exe"
End If

If EXELOADER = "" And Dir("C:\Program Files\MPC-HC\mpc-hc.exe") <> "" Then
    EXELOADER = "C:\Program Files\MPC-HC\mpc-hc.exe"
End If

If EXELOADER = "" And Dir("C:\Program Files (X86)\MPC-HC\mpc-hc.exe") <> "" Then
    EXELOADER = "C:\Program Files (X86)\MPC-HC\mpc-hc.exe"
End If

If NOTEPADEXELOADER = "" And Dir("C:\Program Files\Notepad++\notepad++.exe") <> "" Then
    NOTEPADEXELOADER = "C:\Program Files\Notepad++\notepad++.exe"
End If

If NOTEPADEXELOADER = "" And Dir("C:\Program Files (x86)\Notepad++\notepad++.exe") <> "" Then
    NOTEPADEXELOADER = "C:\Program Files (x86)\Notepad++\notepad++.exe"
End If


Beep
If EXELOADER <> "" Then
    If Dir(EXELOADER) <> "" Then
        'Shell EXELOADER
        On Error Resume Next
        Shell EXELOADER + " """ + RENAME_COPY_FILE + """", vbMaximizedFocus
        If Err.Number > 0 Then
            MsgBox "ERROR LOADING THE SHELL " + vbCrLf + VBCLRF + EXELOADER, vbMsgBoxSetForeground
    
        End If
    End If
End If

If NOTEPADEXELOADER <> "" Then
    If Dir(NOTEPADEXELOADER) <> "" Then
        'Shell EXELOADER
        On Error Resume Next
        ' ---------------------------------------------------------------------
        ' LOGG-A-HOLIC
        ' ---------------------------------------------------------------------
        ' TEMP REMOVE LINE GOT BORED OF LOAD IT IN NOTEPAD PRIOR TO WATCH VIDEO
        ' COULD HAVE OPTION ON
        ' OR FIND IN MENU EASIER ANYHOW
        ' BUT ORGINALLY THERE TO FIND FILM QUICKER
        ' ---------------------------------------------------------------------
        'Shell NOTEPADEXELOADER + " """ + RENAME_COPY_FILE + """", vbNormalNoFocus
        If Err.Number > 0 Then
            MsgBox "ERROR LOADING THE SHELL " + vbCrLf + VBCLRF + EXELOADER, vbMsgBoxSetForeground
        End If
    End If
End If


'Exit Sub
Beep
Unload Me

'Shell "C:\Program Files (x86)\K-Lite Codec Pack\MPC-HC64\mpc-hc64_nvo.exe" + " """ + RENAME_COPY_FILE + """ /PLAY", vbMaximizedFocus
''Shell "C:\Program Files (x86)\K-Lite Codec Pack\MPC-HC64\mpc-hc64_nvo.exe" + " """ + RENAME_COPY_FILE + """", vbMaximizedFocus
'Shell "C:\Program Files (x86)\K-Lite Codec Pack\MPC-HC64\mpc-hc64_nvo.exe" + " Play", vbMaximizedFocus
Beep

'mpc Clear
'mpc load <playlist_name>
'mpc Play
'----
'MEDIA PLAYER CLASSIC LOAD A PLAYLIST - Google Search
'https://www.google.co.uk/search?q=MEDIA+PLAYER+CLASSIC+LOAD+A+PLAYLIST&rlz=1C1CHBD_en-GBGB721GB721&oq=MEDIA+PLAYER+CLASSIC+LOAD+A+PLAYLIST&aqs=chrome..69i57j0j69i64.7640j0j7&sourceid=chrome&ie=UTF-8
'----
'----
'linux - How do I play a particular playlist with mpc? - Super User
'http://superuser.com/questions/444047/how-do-i-play-a-particular-playlist-with-mpc
'----
'End
'Unload Me

End Sub


Sub MPC_MEDIA_PLAYER_SCAN()

    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' --- ANOTHER MENU SET VAR SELECT DRIVE SET ADD THEN DO
    ' --- MNU_IRFAN_MODE_SET = True
    ' --- FILTER OUT VAR GATHERED
    
    
    '-------------------------------------------------
    'IF SCAN TO DO SAME AS LAST
    'DON'T RESET COUNT INI BACK NAUGHT
    '-------------------------------------------------
    'AND DON'T RESET OR DO ANYTHING IF ERROR READ INFO
    '-------------------------------------------------
'    If DUPE_LAST_PATH2_ERROR_VAR = False Then
'        If DUPE_LAST_PATH = False Then
'            Call MNU_INI_BACK_NAUGHT_Click
'        End If
'    End If
    
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
    'NOT_USE_XSCRIPT = True
    NOT_USE_XSCRIPT = False
    
    '----------------------------------------
    'C:\TEMP\IRFAN_SLIDESHOW_X_SCRIPT.TXT
    'MOST LIKLY NEEDED TO RUN WITH SETTING
    'BUT NOT MEDIA PLAYER CLASSICAL
    'FIND OUT LATER
    '----------------------------------------
    
    If NOT_USE_XSCRIPT = False Then
        Call X_SCRIPT_CHUNK
    End If

'    FF = FreeFile
'    Open "C:\TEMP\IRFAN_SLIDESHOW_X_SCRIPT_LAST_RUN.TXT" For Input As #FF
    
'    End If

    'ReDim Preserve HHAR01(HHC01)

    MNU_IRFAN_MODE_SET = True
    
    If Dir("D:\#0 IRFAN SHOW SCRIPT FILE SET", vbDirectory) = "" Then
        MkDir "D:\#0 IRFAN SHOW SCRIPT FILE SET"
    End If
    
    
    
    i1 = ""
    i2 = ""
      
    ScanPath.cboMask.Text = "*.JPG;*.GIF"
    ScanPath.cboMask.Text = "*.JPG;*.GIF"
    ScanPath.cboMask.Text = "*.JPG;*.JPEG"
    'ScanPath.cboMask.Text = "*.JPG;*.GIF;*.MPG;*.AVI;*.MP4;*.3GP"
    'ScanPath.cboMask.Text = "*.MPG;*.AVI;*.MP4"
     
     
    'i1 = i1 + "CLP|ANI|VOB|MPG|MP4|3GP|AVI|WMV"

    i1 = i1 + "*.JPG|JPEG|JPE|GIF|BMP|DIB|TIF|TIFF|PNG|PCX|TGA|LBM|"
    If MNU_IRFAN_04.Checked = False Then
        i1 = i1 + "CLP|ANI|VOB|MPG|MP4|3GP|AVI|WMV|FLV"
    End If
    
    '----------------------------------------------------------
    i1 = "*.AVI|MPG|MP4|VOB|FLV"
    '----------------------------------------------------------
    
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
    
    'i1 = i1 + "CLP|" '"ANI|VOB|MPG|MP4|3GP|AVI|WMV|FLV|ISO"
    
    i1 = ""
    i1 = i1 + "CLP|ANI|" 'VOB|MPG|MP4|3GP|AVI|WMV|FLV|ISO"
    
    i2 = i2 + "AVI|RM|"
    i2 = i2 + "MPG|"
    i2 = i2 + "MPEG|"
    i2 = i2 + "MOV|"
    i2 = i2 + "CAM|"
    i2 = i2 + "WMV|FLV|FV|"
    'i2 = i2 + "MP4|MKV|MTS|M2TS|TS|M2T|M4V|3GP|VOB|ISO"
    i2 = i2 + "MP4|MKV|MTS|M2TS|TS|M2T|M4V|VOB|ISO"
    
    i2 = Replace(i1 + i2, "|", ";*.")
    
    '----------------------------------------------------------
    'SEARCH FIND THE EXTENSION SET USE FOR MEDIA PLAYER CLASSIC
    '----------------------------------------------------------

    '----------------------------------------------------------
    i1 = "*.MPG|MP4|AVI|VOB|M4V|WMV|MKV|ASF|MOV|FLV|MPEG|IFO|ISO|QT|RM"
    '----------------------------------------------------------
    
    i1 = Replace(i1, "|", ";*.")
    ScanPath.cboMask.Text = i1
    
'    ScanPath.cboMask.Text = i2
    
    
    
    ScanPath.lblCount9.Caption = Replace(ScanPath.cboMask.Text, ";", "; ")

    VAR2 = Format(Now, "YYYY_MM_DD_HH_MM_SS") + ".TXT"
    
    FILE11 = "D:\#0 IRFAN SHOW SCRIPT FILE SET\#00 IRFAN SCRIPT -- " + VAR2
    FILE12 = "D:\#0 IRFAN SHOW SCRIPT FILE SET\#00 IRFAN SCRIPT APPEND BUILD -- " + VAR2
    
    'If MULTI_PATH_SCAN = True And MULTI_PATH_SCAN_COUNT = 1 Then
    If Dir(FILE11) <> "" Then Kill FILE11
    If Dir(FILE12) <> "" Then Kill FILE12
    'End If
    
        
    FF01 = FreeFile
    Open FILE11 For Output As #FF01

    FF02 = FreeFile ' LESS MODIFY
    Open FILE12 For Append As #FF02
    
    FF03 = -3
    FF04 = -4
    
    ScanPath.ListView1.View = lvwReport
    
    If IRFAN_TO_DO_FROM_SEND_TO = False And RUN_FROM_MEDIA_PLAYER_CLASSIC = False Then
    'If IRFAN_TO_DO_FROM_SEND_TO = False Then
    
        MsgBox "SCAN FROM SEND TO HAS NOT A VALID PATH = SO ALL DRIVES WILL SCAN", vbMsgBoxSetForeground
        
    
        On Error Resume Next
        For Each Control In LabDR
        
            MCHAR = Mid(Control.Caption, 2, 1) + ":\"
            Err.Clear
            Set F = FS.GetDrive(MCHAR)
                    
            If Err.Number = 0 Then
                If F.IsReady = True Then
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
    


'---------------------------------------------
' FF02 __ OUTOUT APPENDER FROM FILE12 VARIABLE
'---------------------------------------------
'LOAD ALLINTO A LISTVIEW FOR SORTER
'---------------------------------------------
    'ScanPath.ListView1.ListItems.Clear
    
'    With ScanPath.ListView1
'        .ColumnHeaders.Add , "FILE", "File", 4200, lvwColumnLeft
'        .View = lvwReport
'    End With
    
    If 1 = 2 Then
    
        ScanPath.ListView1.View = lvwReport
        ScanPath.ListView1.ListItems.Clear
        FF02 = FreeFile
        Open FILE12 For Input As #FF02
        If LOF(FF02) > 0 Then
            
            'With ScanPath.ListView1
            Do
            Line Input #FF02, LINE_STRING
            
            
            If Old_LINE_STRING <> LINE_STRING Then
            
                ScanPath.ListView1.ListItems.Add , , LINE_STRING
            
            End If
            
            Old_LINE_STRING = LINE_STRING
            'Set LV = .ListItems.Add(, , LINE_STRING)
            
            Loop Until EOF(FF02)
    
            'End With
        End If
        Close FF02
    
    End If

'---------------------------------------------

'        ListView1.Top = 1
'        ListView1.Top = 1
        
'        SORT_ON_DATE = True
        
        For WE = 1 To ScanPath.ListView1.ListItems.Count
        
'        If InStr(ScanPath.ListView1.ListItems.Item(WE), "C:\DOWNLOADS\# 00 VIDEO") > 0 Then
'            SORT_ON_DATE = True
'        End If
'
'        If InStr(ScanPath.ListView1.ListItems.Item(WE), "D:\VIDEO\NOT\X 00 NOT ME") > 0 Then
'            SORT_ON_DATE = False
'        End If
        
        
        'FIND HERE
'        D:\VIDEO\NOT\X 00 NOT ME
        
            If ScanPath.ListView1.ListItems.Item(WE).SubItems(4) = "" Then
                B1 = ScanPath.ListView1.ListItems.Item(WE)
                Set F = FSO.GetFile(B1)
                XVB_DATE_2 = Format(F.DateLastModified, "YYYY MM DD HH MM SS")
                
                ScanPath.ListView1.ListItems.Item(WE).SubItems(4) = XVB_DATE_2
            End If
        Next
        
        
        ScanPath.Refresh
        DoEvents
        
        AT$ = AT$
    
    SORT_ON_DATE = False
    If Form_SEND_TO.MNU_SORT_ON_DATE_TRUE.Caption = "SORT ON DATE = TRUE" Then
        SORT_ON_DATE = True
    End If
        
    'sort of PATH then FILE
    If SORT_ON_DATE = True Then
        ScanPath.ListView1.SortOrder = lvwDescending
        ScanPath.ListView1.SortKey = 4
    Else
        ScanPath.ListView1.SortOrder = lvwAscending
        ScanPath.ListView1.SortKey = 0
    End If
    
    ScanPath.ListView1.Sorted = True
    ScanPath.ListView1.Sorted = False
    
    FF02 = FreeFile
    Open FILE12 For Output As #FF02
        For WE = 1 To ScanPath.ListView1.ListItems.Count
            B1 = ScanPath.ListView1.ListItems.Item(WE)
            Print #FF02, B1
        Next
    Close #FF02




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

    
    FS.CopyFile FILE12, FILE13_SCRIPT_TEMP_FOLDER
    FILE14 = "C:\TEMP\IRFAN_SLIDESHOW-" + GetComputerName + "_BACKUP.TXT"
    FS.CopyFile FILE12, FILE14


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
    'Call MNU_INI_READ
    '------------------------------------------------------------
    'GET THE INI FILE AT POINTER -- AND UPDATE ALL LABELS ON FORM
    '------------------------------------------------------------
    'Call Label0_Click
    '------------------------------------------------------------
    '------------------------------------------------------------

    



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

    Call FLAG_TO_DO_XSCRIPT_WRITE
    
    Call XSCRIPT_CHUNK_DATE_UPDATE_SHOW_REQUIRE

    If MNU_IRFAN_07_NOT_USE_XSCRIPT.Checked = False Then
        '-------------------------------------------------------------------------------
        Call Label6_Click
        '-------------------------------------------------------------------------------
        'Label6.Caption = "IRFAN XSCRIPT PROCESS -- WORKING"
        '-------------------------------------------------------------------------------
    End If

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

'Label6_Click
Label6.Caption = "IRFAN XSCRIPT PROCESS -- WORKING"
Label6.Refresh
DoEvents


XSCRIPT_CHUNK_AND_PROCESS

Set F = FS.GetFile(FILE13_SCRIPT_TEMP_FOLDER)

If BEFORE_XSCRIPT_COUNTER = AFTER_XSCRIPT_COUNTER Then STAT_SAME = " SAME - "
Label6.Caption = "IRFAN XSCRIPT PROCESS - FSIZE -" + Str(F.Size) + " - LINE COUNT -" + STAT_SAME + "BEFORE=" + Str(BEFORE_XSCRIPT_COUNTER) + " - AFTER=" + Str(AFTER_XSCRIPT_COUNTER) + " - TOTAL GONE=" + Str(BEFORE_XSCRIPT_COUNTER - AFTER_XSCRIPT_COUNTER)

'Sleep 200
Label6.BackColor = QB
   
    
End Sub

Private Sub Label7_Click()
'LABEL7.CAPTION
'Me.WindowState = vbMinimized

Label7.BackColor = QBColor(14)
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
            If F.IsReady = True Then
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


Private Sub MNU_EXIT_Click()
End
End Sub

Sub SORT_ON_DATE_CHECKER()

'C:\DOWNLOADS\# 00 VIDEO"
If InStr(AT$, "\DOWNLOADS") > 0 Then
    MNU_SORT_ON_DATE_TRUE.Caption = "SORT ON DATE = TRUE"
End If

If InStr(AT$, "\DSC") > 0 Then
    MNU_SORT_ON_DATE_TRUE.Caption = "SORT ON DATE = TRUE"
End If

If InStr(AT$, "\DOWNLOADS\# 00 VIDEO") > 0 Then
    MNU_SORT_ON_DATE_TRUE.Caption = "SORT ON DATE = TRUE"
End If

' D:\VIDEO\NOT\X 00 NOT ME
If InStr(AT$, "\VIDEO\NOT\X 00 NOT ME") > 0 Then
    Beep
    MNU_SORT_ON_DATE_TRUE.Caption = "SORT ON DATE = NOT"
    MNU_SORT_ON_DATE_TRUE.Caption = "SORT ON DATE = TRUE"
End If


End Sub

Private Sub MNU_FOLDERING_ARRAY_Click(index As Integer)


If index = 1 Then PATH_FINDER = "C:\DOWNLOADS"
If index = 2 Then PATH_FINDER = "D:\DOWNLOADS"
If index = 3 Then PATH_FINDER = "D:\DSC"
' DOWNLOAD AND GROUP MPC ' USED TO BE INCLUDED -- D:\0 00 VIDEO 02 VOB\
' NOT DSC
If index = 4 Then PATH_FINDER = "C:\DOWNLOADS\# 00 VIDEO\;D:\0 00 VIDEO 03 DOWNLOAD\;D:\0 00 VIDEO 01\;D:\# MY DOCS\;D:\0 00 VIDEO CCSS\;D:\0 00 ART LOGGERS - WEBCAM\;D:\DL\"
' DOWNLOAD AND GROUP MPC ' USED TO BE INCLUDED -- D:\0 00 VIDEO 02 VOB\
If index = 5 Then PATH_FINDER = "C:\DOWNLOADS\# 00 VIDEO\;D:\0 00 VIDEO 03 DOWNLOAD\;D:\0 00 VIDEO 01\;D:\# MY DOCS\;D:\0 00 VIDEO CCSS\;D:\0 00 ART LOGGERS - WEBCAM\;D:\DL\;D:\DSC\"
' 01 MM
If index = 6 Then PATH_FINDER = "D:\VIDEO\NOT\X 01 ME;D:\VI_ DSC ME"
' 01 MM NETWORK
If index = 7 Then PATH_FINDER = "\\7-asus-gl522vw\7_asus_gl522vw_02_d_drive\VI_ DSC ME;\\7-asus-gl522vw\7_asus_gl522vw_02_d_drive\VIDEO\NOT\X 01 ME"
' 02 MM 4TB
If index = 8 Then PATH_FINDER = "D:\VIDEO\NOT\X 01 ME;D:\VI_ DSC ME;M:\Videos\0 00 VIDEO CAMERA - ME"
' 02 MM 4TB NETWORK
If index = 9 Then PATH_FINDER = "\\7-asus-gl522vw\7_asus_gl522vw_02_d_drive\VIDEO\NOT\X 01 ME;\\7-asus-gl522vw\7_asus_gl522vw_10_1_samsung_4tb_d\VI_ DSC ME;\\7-asus-gl522vw\7_asus_gl522vw_10_1_samsung_4tb_d\Videos\0 00 VIDEO CAMERA - ME"
' NOT ME
If index = 10 Then PATH_FINDER = "D:\VIDEO\NOT\X 00 NOT ME"
' NOT ME NETWORK
If index = 11 Then PATH_FINDER = "\\7-asus-gl522vw\7_asus_gl522vw_02_d_drive\VIDEO\NOT\X 00 NOT ME"
If index = 12 Then PATH_FINDER = "D:\0 00 VIDEO SNAPSHOT CCSE HIKVISION\snapshot"
If index = 13 Then PATH_FINDER = "\\8-msi-gp62m-7rd\8_msi_gp62m_7rd_02_d_drive\0 00 VIDEO SNAPSHOT CCSE HIKVISION\snapshot"

For r = 1 To 1
    MNU_FOLDERING_ARRAY(r).Checked = False
Next
MNU_FOLDERING_ARRAY(index).Checked = True

' TO DO LATER
' D:\0 00 VIDEO SNAPSHOT CCSE HIKVISION\;

'MNU_IRFAN_07_NOT_USE_XSCRIPT.Checked = True

AT$ = PATH_FINDER

'AT$ = "\\7-asus-gl522vw\7_asus_gl522vw_02_d_drive\0 00 VIDEO SNAPSHOT CCSE HIKVISION"

Label3.Caption = "SCAN TO GIVE -- " + AT$

Call SORT_ON_DATE_CHECKER

Set FS = CreateObject("Scripting.FileSystemObject")

If Len(AT$) > 2 Then
    If Mid(AT$, 1, 2) = "\\" Then
        MNU_DRIVE_VOL_NAME.Caption = "DRIVE VOLUME NAME = __ " + Mid(AT$, 1, InStr(3, AT$, "\")) + " __ NETWORK"
    Else
        Set F = FS.GetDrive(Mid(AT$, 1, 3))
        MNU_DRIVE_VOL_NAME.Caption = "DRIVE VOLUME NAME = __ " + F.VolumeName
    End If
End If

'If index > 0 Then
'    ' CALL MEDIA PLAYER CLASSIC _ AUTO
'    Call Label12_Click
'End If

End Sub

Private Sub MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC_Click()

If MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC.Checked = True Then MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC.Checked = False
If MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC_2.Checked = True Then MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC_2.Checked = False: Exit Sub

If MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC.Checked = False Then MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC.Checked = True
If MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC_2.Checked = False Then MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC_2.Checked = True

End Sub

Private Sub MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC_2_Click()

Call MNU_IRFAN_08_NOT_USE_XSCRIPT_MPC_Click

End Sub


Private Sub MNU_XNVIEW_EXE_CLICK()

'MNU_I_VIEW64_EXE

'C:\PSTART\PROGS\#_PORTABLEAPPS\PORTABLEAPPS\IRFANVIEWPORTABLE\IRFANVIEWPORTABLE.EXE


FILE_PATH = "C:\PSTART\PROGS\#_PORTABLEAPPS\PORTABLEAPPS\XNVIEWPORTABLE\XNVIEWPORTABLE.EXE"
If Dir(FILE_PATH) <> "" Then FILE_PATH2 = FILE_PATH

FILE_PATH = "C:\PROGRAM FILES (X86)\XNVIEW\XNVIEW.EXE"
If Dir(FILE_PATH) <> "" Then FILE_PATH2 = FILE_PATH

'FILE_PATH = "C:\PROGRAM FILES (X86)\IRFANVIEW\I_VIEW64.EXE"
'If Dir(FILE_PATH) <> "" Then FILE_PATH2 = FILE_PATH


'FILE_PATH = "C:\PROGRAM FILES (X86)\IRFANVIEW\I_VIEW64.EXE"
'IF DIR(FILE_PATH) <> "" THEN FILE_PATH2 = FILE_PATH


Shell FILE_PATH2 + " " + FILE13_SCRIPT_TEMP_FOLDER, vbMaximizedFocus


'C:\PROGRAM FILES (X86)\XNVIEW\XNVIEW.EXE
'C:\PSTART\PROGS\#_PORTABLEAPPS\PORTABLEAPPS\XNVIEWPORTABLE\XNVIEWPORTABLE.EXE
'C:\PSTART\PROGS\#_PORTABLEAPPS\PORTABLEAPPS\IRFANVIEWPORTABLE\IRFANVIEWPORTABLE.EXE


End Sub

Private Sub MNU_I_VIEW64_EXE_CLICK()

'MNU_I_VIEW64_EXE

FILE_PATH = "C:\Program Files\IrfanView\i_view64.exe"
If Dir(FILE_PATH) <> "" Then FILE_PATH2 = FILE_PATH
FILE_PATH = "C:\PROGRAM FILES (X86)\IrfanView\i_view64.exe"
If Dir(FILE_PATH) <> "" Then FILE_PATH2 = FILE_PATH

'FILE_PATH = "C:\PROGRAM FILES (X86)\IrfanView\i_view64.exe"
'If Dir(FILE_PATH) <> "" Then FILE_PATH2 = FILE_PATH

FILE_PATH = "C:\PStart\Progs\#_PortableApps\PortableApps\IrfanViewPortable\App\IrfanView\i_view32.exe"
If Dir(FILE_PATH) <> "" Then FILE_PATH2 = FILE_PATH


Shell FOLDER_PATH_IRFANView2, vbNormalNoFocus



'Shell FILE_PATH2 + " " + FILE13_SCRIPT_TEMP_FOLDER, vbNormalNoFocus

'C:\Program Files (x86)\XnView\xnview.exe
'C:\PStart\Progs\#_PortableApps\PortableApps\XnViewPortable\XnViewPortable.exe
'C:\PStart\Progs\#_PortableApps\PortableApps\IrfanViewPortable\IrfanViewPortable.exe


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
Shell "NOTEPAD """ + App.Path + "\IRFAN_SLIDESHOW_X_SCRIPT.TXT""", vbNormalFocus


End Sub

Sub INI_FILE()

' FILE13_SCRIPT_TEMP_FOLDER
' INI_StartIndex

'----------------------------------
'FIRST LOAD OF THIS VAR HAPPEN HERE
'----------------------------------
'FILE13_SCRIPT_TEMP_FOLDER = "C:\TEMP\IRFAN_SLIDESHOW-" + GetComputerName + ".txt"

If Dir(FILE13_SCRIPT_TEMP_FOLDER) <> "" Then
    Set F = FS.GetFile(FILE13_SCRIPT_TEMP_FOLDER)

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

End Sub

Private Sub MNU_IRFAN_09_EXPLORER_INI_Click()

Call MNU_INI_READ


Label11.Caption = "EXPLORER CURRENT INI FILE = " + INI_FILE_AT_LINE

Shell "EXPLORER /select, """ + INI_FILE_AT_LINE + """", vbNormalFocus


End Sub

Private Sub MNU_RERUN_Click()
Reset
UNLOAD_WITH_IRFAN_KILL = True

'SET ALL TIMERS IN ALL FORMS ENABLED=FALSE
On Error Resume Next
    For I = 0 To Forms.Count - 1
        For Each Control In Forms(I).Controls
            If InStr(UCase(Control.Name), "TIMER") > 0 Then
                'Debug.Print Control.Name
                Control.Enabled = False
            End If
        Next
    Next I
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

Private Sub Mnu_VB_ME_Click()

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


'Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
'End

'    Dim MSGQ, IR, CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2
    
    On Error Resume Next
    
    Beep
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
    'End

End Sub
'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)

    Dim FileSpec, FILE_SPEC_GO, I
    
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".vbp"
    I = CODER_VBP_FILE_NAME_2
    CODER_EXE_FILE_NAME_1 = I
    
    If InStr(I, " - 64 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(I, " - 64 BIT.vbp", ".vbp")
    If InStr(I, " - 32 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(I, " - 32 BIT.vbp", ".vbp")
    
    CODER_EXE_FILE_NAME_1 = Replace(CODER_EXE_FILE_NAME_1, ".vbp", ".exe")
    
    FileSpec = CODER_VBP_FILE_NAME_2
    FILE_SPEC_GO = FileSpec
    If Dir(FileSpec) = "" Then
        If LCase(GetSpecialfolder(&H29)) = LCase("C:\Windows\SysWOW64") Then
            FileSpec = Replace(FileSpec, ".vbp", " - 64 BIT.vbp")
        Else
            FileSpec = Replace(FileSpec, ".vbp", " - 32 BIT.vbp")
        End If
    End If
    
    If Dir(FileSpec) <> "" Then FILE_SPEC_GO = FileSpec
    
    CODER_VBP_FILE_NAME_2 = FILE_SPEC_GO
    
End Sub

Private Sub MNU_VB_Click()

End Sub

Private Sub MNU_VB_FOLDER_Click()
    Beep
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    Shell "EXPLORER /SELECT, """ + CODER_VBP_FILE_NAME_2 + """", vbNormalFocus
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
'---------------------------------------------------
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
    If Control.index Mod 2 = 0 Then
        Control.BackColor = &HC0FFC0
    End If
    MCHAR = Mid(Control.Caption, 2, 1) + ":\"
    Err.Clear
    Set F = FS.GetDrive(MCHAR)
            
    If F.IsReady = False Or Err.Number > 0 Then
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



'----
'FILESYSTEMOBJECT CHECK A DRIVE IS READY - Google Search
'https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB744GB744&q=FILESYSTEMOBJECT+CHECK+A+DRIVE+IS+READY&spell=1&sa=X&ved=0ahUKEwjeu8Cuo7vWAhXJJlAKHRw6CQAQvwUIJSgA&biw=1536&bih=682
'--------
'Microsoft Windows 2000 Scripting Guide - Ensuring That a Drive is Ready
'https://technet.microsoft.com/en-us/library/ee198736.aspx
'----
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set colDrives = objFSO.Drives
For Each objDrive In colDrives
    If objDrive.IsReady = True Then
        DRIVE_LETTERING = DRIVE_LETTERING + objDrive.DriveLetter
    End If
Next

For Each Control In LabDR
    If Control.index Mod 2 = 0 Then
        Control.BackColor = &HC0FFC0
    End If
    MCHAR = Mid(Control.Caption, 2, 1) '+ ":\"
    'Err.Clear
    
    
'    Set F = FS.GetDriveName(FS.GetAbsolutePathName(MCHAR))
'    Set F = FS.GetDrive(MCHAR)
'    Set F = FS.GetDrive(MCHAR)
'    If F.IsReady = False Or Err.Number > 0 Then
                  
            
    If InStr("*" + DRIVE_LETTERING, MCHAR) = 0 Then
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
    'i2 = i2 + "MP4|MKV|MTS|M2TS|TS|M2T|M4V|3GP|VOB"
    i2 = i2 + "MP4|MKV|MTS|M2TS|TS|M2T|M4V|VOB"
    
    i2 = Replace(i1 + i2, "|", ";*.")
    
    
'    ScanPath.cboMask.Text = i2
    
    
    
    ScanPath.lblCount9.Caption = Replace(ScanPath.cboMask.Text, ";", "; ")

    

    VAR2 = Format(Now, "YYYY_MM_DD_HH_MM_SS") + ".TXT"
    
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
                If F.IsReady = True Then
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

    
    FS.CopyFile FILE12, FILE13_SCRIPT_TEMP_FOLDER
    FILE14 = "C:\TEMP\IRFAN_SLIDESHOW-" + GetComputerName + "_BACKUP.TXT"
    FS.CopyFile FILE12, FILE14


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
    
    FOLDER_PATH_IRFANView = "C:\PStart\Progs\#_PortableApps\PortableApps\IrfanViewPortable\App\IrfanView\i_view32.exe"
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

    Shell FOLDER_PATH_IRFANView2 + " /killmesoftly"
    
    'DONT ESC QUIT EXIT LINE BELOW
    Shell FOLDER_PATH_IRFANView2 + " /slideshow=""" + FILE13_SCRIPT_TEMP_FOLDER + """ /ONE /bf /silent /gray /RELOADONLOOP /QUITE", vbNormalFocus

'    Shell FOLDER_PATH_IRFANView + " /slideshow=""" + FILE13_SCRIPT_TEMP_FOLDER + """ /BF /RELOADONLOOP /QUITE", vbNormalFocus



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


Sub FLAG_TO_DO_XSCRIPT_WRITE()
    
    If Dir(FILE13_SCRIPT_TEMP_FOLDER) <> "" Then
        Set F = FS.GetFile(FILE13_SCRIPT_TEMP_FOLDER)
        ADATE = F.DateLastModified
        FF = FreeFile
        Open FILE19X_SCRIPT_DATE_MODIFIED For Output As #FF
            Print #FF, ADATE
        Close FF
        FLAG_TO_DO_XSCRIPT = False
    End If

End Sub

Sub FLAG_TO_DO_XSCRIPT_READ_02()
    
    FILE19X_SCRIPT_DATE_MODIFIED = "C:\TEMP\IRFAN_SLIDESHOW_X_SCRIPT_DATE_MODIFIED.TXT"
    FLAG_TO_DO_XSCRIPT = True

    If Dir(FILE13_SCRIPT_TEMP_FOLDER) <> "" Then
        Set F = FS.GetFile(FILE13_SCRIPT_TEMP_FOLDER)
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
    
    Call FLAG_TO_DO_XSCRIPT_READ_02

    If FLAG_TO_DO_XSCRIPT = False Then
        'SEARCH
        '---------------
        'Label6.Caption
        '---------------
        Label6.Caption = "IRFAN XSCRIPT PROCESS -- NOT REQUIRE -- BUT IF MANUAL EDIT -- YES MAYBE DO"
    Else
        Label6.Caption = "IRFAN XSCRIPT PROCESS -- REQUIRED MAYBE -- IS INCLUDED DURING SCAN ALSO - MAYBE IF MODIFIED BY EDIT"
        Label6.BackColor = QBColor(14)

    End If

End Sub
    

Sub XSCRIPT_CHUNK_AND_PROCESS_MPC()
    
    '-----------------------------------------------
    'JOB
    '-----------------------------------------------
    'WANT FILE DATE MODIFIED SO NOT DO TWICE IF EVER
    '-----------------------------------------------
    
    'X_SCRIPT_CHUNK
    
    FILE17 = "C:\TEMP\IRFAN_SLIDESHOW-A-Z.TXT"
    FILE18 = FILE13_SCRIPT_TEMP_FOLDER
    
    'Call XSCRIPT_CHUNK_DATE_UPDATE_SHOW_REQUIRE
    
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
    
    FS.CopyFile FILE18, FILE17
    
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
        
        If EOF(FF01) Then
            Exit Do
        End If
        'FILE MIGHT BE EMPTY
        '-------------------
        On Error Resume Next
'        If EOF(FF01) Then
'            Close FF01
'            Close FF02
'            Exit Sub
'        End If
        Err.Clear
        Line Input #FF01, FileSpec
        If EOF(FF01) And Err.Number > 0 Then
            Exit Do
            'Close FF01
            'Close FF02
            'Exit Sub
        End If

        STRING_SIZE = STRING_SIZE + Len(FileSpec + vbCrLf)
        If FileSpec <> OFILESPEC Then
            OFILESPEC = FileSpec 'INTERLACE - REMOVE DUPE FILE
            
            BEFORE_XSCRIPT_COUNTER = BEFORE_XSCRIPT_COUNTER + 1
            
            'BLACKLIST BLOCKER SCRIPT 01 OF 02
            ESCAPE = 0
            
            'If InStr(FileSpec, "D:\VIDEO\") > 0 Then ESCAPE = 1
            'If InStr(FileSpec, "D:\V_DSC_ME\") > 0 Then ESCAPE = 1
            
            If ESCAPE = 0 Then
'                Print #FF02, FileSpec
                Call XSCRIPT(FileSpec) ' THROW IN FF02 - FREEFILE
                Else
                Print #FF02, FileSpec
                
            End If

            
            'Call XSCRIPT(FileSpec) ' THROW IN FF02 - FREEFILE
        End If
    Loop Until EOF(FF01) Or STRING_SIZE >= LOF(FF01)
    Close FF01, FF02

    '--------------------------------------------------------------
    'FORGET THIS ONE AT MOMENT IT RESET INI COUNTER IF ANY CHANGING
    'DONE REMOVE BY X_SCRIPT
    '--------------------------------------------------------------
    'If WORKFLAG = True Then MNU_INI_BACK_NAUGHT_Click

    FILE_SCRIPT_COUNTER = AFTER_XSCRIPT_COUNTER

    'Call FLAG_TO_DO_XSCRIPT_WRITE
    
    
    Set F = FS.GetFile(FILE18)
    Label5.Caption = "IRFAN -- " + FILE13_SCRIPT_TEMP_FOLDER + " -- FSIZE -" + Str(F.Size) + " -- LINE COUNT =" + Str(FILE_SCRIPT_COUNTER)
    'Label6.Caption = "Irfan XSCRIPT PROCESS - FSIZE --" + Str(F.Size) + "LINE COUNT - BEFORE=" + Str(BEFORE_XSCRIPT_COUNTER) + " - AFTER=" + Str(AFTER_XSCRIPT_COUNTER)

    FS.CopyFile FILE18, FILE17


End Sub





Sub XSCRIPT_CHUNK_AND_PROCESS()
    
    '-----------------------------------------------
    'JOB
    '-----------------------------------------------
    'WANT FILE DATE MODIFIED SO NOT DO TWICE IF EVER
    '-----------------------------------------------
    
    
    Call X_SCRIPT_CHUNK
    
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
    
    FS.CopyFile FILE18, FILE17
    
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
        
        If EOF(FF01) Then Exit Do
        On Error Resume Next
        Line Input #FF01, FileSpec
        If Err.Number > 0 Then
            On Error GoTo 0
        
            Exit Do
        
        End If
        
        STRING_SIZE = STRING_SIZE + Len(FileSpec + vbCrLf)
        If FileSpec <> OFILESPEC Then
            OFILESPEC = FileSpec 'INTERLACE - REMOVE DUPE FILE
            
            BEFORE_XSCRIPT_COUNTER = BEFORE_XSCRIPT_COUNTER + 1
            
            Call XSCRIPT(FileSpec) ' THROW IN FF02 - FREEFILE
        End If
    Loop Until EOF(FF01) Or STRING_SIZE >= LOF(FF01)
    Close FF01, FF02

    '--------------------------------------------------------------
    'FORGET THIS ONE AT MOMENT IT RESET INI COUNTER IF ANY CHANGING
    'DONE REMOVE BY X_SCRIPT
    '--------------------------------------------------------------
    'If WORKFLAG = True Then MNU_INI_BACK_NAUGHT_Click

    FILE_SCRIPT_COUNTER = AFTER_XSCRIPT_COUNTER

    Call FLAG_TO_DO_XSCRIPT_WRITE
 
    Set F = FS.GetFile(FILE18)

    Label5.Caption = "IRFAN -- " + FILE13_SCRIPT_TEMP_FOLDER + " -- FSIZE -" + Str(F.Size) + " -- LINE COUNT =" + Str(FILE_SCRIPT_COUNTER)
    'Label6.Caption = "Irfan XSCRIPT PROCESS - FSIZE --" + Str(F.Size) + "LINE COUNT - BEFORE=" + Str(BEFORE_XSCRIPT_COUNTER) + " - AFTER=" + Str(AFTER_XSCRIPT_COUNTER)

    FS.CopyFile FILE18, FILE17

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
        
    If Dir(App.Path + "\IRFAN_SLIDESHOW_X_SCRIPT.TXT") = "" Then
    
     MsgBox "MISSING " + vbCrLf + vbCrLf + App.Path + "\" + vbCrLf + "IRFAN_SLIDESHOW_X_SCRIPT.TXT" + vbCrLf + vbCrLf + "YOU WILL MOST LIKEY NEED THAT HARDCORE", vbMsgBoxSetForeground
     Exit Sub
    End If
    
    
    Open App.Path + "\IRFAN_SLIDESHOW_X_SCRIPT.TXT" For Input As #FF
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



Public Sub XSCRIPT(FileSpec)
    '-----------------------------------
    'THE XSCRIPT IS RUN DURING SCAN ALSO
    'SO MAY NOT SEE RESULT UPON RUN
    '-----------------------------------

    'If RUN_FROM_MEDIA_PLAYER_CLASSIC = True Then Exit Sub

'   On Error GoTo EXITSUB
'   If Irfan_TO_DO_FROM_SEND_TO = True Then TAGX = False: GoTo WRITEANDEXIT
    
    TAGX = False
    
    For R5 = 1 To UBound(HHAR01)
        'Debug.Print HHAR01(R5)
        If InStr("*" + UCase(FileSpec), HHAR01(R5)) > 0 Then
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
                 i5 = Mid(FileSpec, 4, InStrRev(FileSpec, "\") - 3)
                    
'                    If Len(i5) = Len(Mid(HHAR02(R5), 5 + 1)) Then Stop 'SWAP DIR SAME LENGHT
                    
                    'If Len(i5) = Len(Mid(HHAR02(R5), 5 + 1)) Then Stop ' NOT A CHECK WITH ROOT AT MOMENT
              
            End If
                
                
            'FOR THE IF ON THE SEND TO
            If InStr(UCase(FileSpec), UCase(HHAR02(R5))) > 0 Then
            ' -- FOLDER WANT ---
                TAGX = False    '  --- ALLOW OBJECTIVE
                WORKFLAG = True
                Exit For
            End If
                
            '### SET X DO INCLUDE  --- \ AT END MAKE SINGLE PATH NOT SUB
            '# HASH AND WILDCARD MEAN CERTAIN THING
            '# HASH IGNOR
            '# * IS ALLOW AND FILENAME START AT 5TH CHARATOR
            '# IS SIMUALR TO EXTENSTION AND THEN PART OF PATH
            '# IS * ALLOW THERE BE PATH _ 2 ACT LIKE DIFFERENT SET NOT AS EXTENSION 3 4 CHAR .AVI
            '# IN OTHER WORD ALLOW WILDCARD ANY EXTENSION FOR PATH
            '# SEE TOP FOR INSTRUCTION ALSO

            If Mid(HHAR02(R5), 1, 4) = Mid(UCase(FileSpec), Len(FileSpec) - 3) _
             Or Mid(HHAR02(R5), 1, 1) = "*" Then
            
            ' -- FOLDER WANT ---
            If InStr("*" + UCase(FileSpec), Mid(HHAR02(R5), 5 + 1)) > 0 Then
                    TAGX = False    '  --- ALLOW OBJECTIVE
                    WORKFLAG = True
                    Exit For
                End If
            End If
        Next
    End If
    
WRITEANDEXIT:

    'If RUN_FROM_MEDIA_PLAYER_CLASSIC = True Then TAGX = False

    If TAGX = False Then
        Print #FF02, FileSpec
        AFTER_XSCRIPT_COUNTER = AFTER_XSCRIPT_COUNTER + 1
    End If
    
End Sub




Private Sub MNU_INI_BACK_NAUGHT_Click()

Call MNU_INI_READ

If IsNumeric(INI_StartIndex) Then
    If RC_INI <> "" Then
    RC_INI = Replace(RC_INI, "StartIndex=" + INI_StartIndex, "StartIndex=0")

    FR = FreeFile
    Open FILEPATH_I_VIEW_INI For Output Lock Read Write As #FR
        Print #FR, StrConv(RC_INI, vbUnicode);
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
FILE_PATH = "C:\PStart\Progs\#_PortableApps\PortableApps\IrfanViewPortable\Data\IrfanView\i_view32.ini"
If Dir(FILE_PATH) <> "" Then FILEPATH2 = FILE_PATH

If InStr(FOLDER_PATH_IRFANView2, "C:\PStart\Progs\#_PortableApps\PortableApps\IrfanViewPortable") > 0 Then
    FILEPATH2 = Replace(UCase(FOLDER_PATH_IRFANView2), ".EXE", ".INI")
End If
'FILEPATH2 = Replace(UCase(FILEPATH2), ".EXE", ".INI")


'C:\PStart\Progs\#_PortableApps\PortableApps\IrfanViewPortable\App\IrfanView\i_view32.exe

If Dir(FILEPATH2) = "" Then
    FILE_PATH1 = "C:\Program Files\IrfanView\i_view32.ini"
    FILE_PATH2 = "C:\Program Files (X86)\IrfanView\i_view32.ini"
    FILE_PATH3 = "C:\Program Files\IrfanView\i_view64.ini"
    FILE_PATH4 = "C:\PStart\Progs\#_PortableApps\PortableApps\IrfanViewPortable\Data\IrfanView\i_view32.ini"
    ' C:\PStart\Progs\#_PortableApps\PortableApps\IrfanViewPortable\App\IrfanView\i_view32.exe
    
    MsgBox "NOT FIND " + vbCrLf + FILE_PATH1 + vbCrLf + " OR " + vbCrLf + FILE_PATH2 + vbCrLf + " OR " + vbCrLf + FILE_PATH3 + " OR " + vbCrLf + FILE_PATH4 + vbCrLf, vbMsgBoxSetForeground
    
    Exit Sub
End If

If Dir(FILEPATH2) <> "" Then
    FR = FreeFile
    Open FILEPATH2 For Binary As #FR
        'RC_INI = StrConv(Input(LOF(FR), FR), vbFromUnicode)
        RC_INI = Input(LOF(FR), FR)
    Close #FR
End If

'Form2_ANY_SPECIAL_FOLDER.GetSpecialfolder_Debug

FILEPATH_I_VIEW_INI = FILEPATH2

'Debug.Print RC_INI

PORTABLE_CHECK_VAR = False

If InStr(RC_INI, "INI_Folder=%APPDATA%\IrfanView") > 0 Then
    PORTABLE_CHECK_VAR = True

    X_SP_F = GetSpecialfolder(26) + "\IrfanView\"
    X_SP_F_i_view_ini = X_SP_F + Dir(X_SP_F + "\i_view*.ini")
    If Dir(X_SP_F_i_view_ini) <> "" Then
        FILEPATH_I_VIEW_INI = X_SP_F_i_view_ini
        On Error Resume Next
        FR = FreeFile
        Open X_SP_F_i_view_ini For Binary As #FR
        'Open X_SP_F_i_view_ini For Input As #FR
            If EOF(FR) = False Then
                RC_INI = Input(LOF(FR), FR)
            End If
        Close #FR
        On Error GoTo 0
    End If
End If

'PORTABLE
If PORTABLE_CHECK_VAR = True Then
    If InStr(RC_INI, StrConv("UNICODE", vbUnicode)) > 0 Then
        RC_INI = StrConv(RC_INI, vbFromUnicode)
    End If
    
    If InStr(RC_INI, "INI_Folder=") > 0 Then
            FR = FreeFile
            Open FILEPATH_I_VIEW_INI For Binary As #FR
                RC_INI = Input(LOF(FR), FR)
            Close #FR
            
            If InStr(RC_INI, StrConv("UNICODE", vbUnicode)) > 0 Then
                RC_INI = StrConv(RC_INI, vbFromUnicode)
            End If
    
            If InStr(RC_INI, "INI_Folder=") > 0 Then
                ABVAR_1 = InStr(RC_INI, "INI_Folder=") + Len("INI_Folder=")
                ABVAR_2 = InStr(InStr(RC_INI, "INI_Folder="), RC_INI, vbCrLf) - ABVAR_1
                
                FILEPATH_I_VIEW_INI = Mid(RC_INI, ABVAR_1, ABVAR_2) + "\i_view32.ini"
                
                If Dir(FILEPATH_I_VIEW_INI) <> "" Then
                    FR = FreeFile
                    Open FILEPATH_I_VIEW_INI For Binary As #FR
                        RC_INI = StrConv(Input(LOF(FR), FR), vbFromUnicode)
                    Close #FR
                End If
                PORTABLE_CHECK_VAR = False
            End If
    End If
    
End If



'Debug.Print RC_INI

' ------------------------------------------------------------------------
' IN PORTABLE MODE THE INDEX POINTER FOR SLIDE SHOW IS FOUND BY SEARCH FOR
' "StartIndex="
' WORKER
' ------------------------------------------------------------------------
On Error GoTo 0
INI_StartIndex1 = 0
INI_StartIndex1 = InStr(RC_INI, "StartIndex=") + Len("StartIndex=")
If InStr(RC_INI, "StartIndex=") = 0 Then
    INI_StartIndex = "INI COUNT INFO NOT FOUND"
    Exit Sub
End If
INI_StartIndex2 = InStr(INI_StartIndex1, RC_INI, vbCrLf)
On Error Resume Next
INI_StartIndex = Mid(RC_INI, INI_StartIndex1, INI_StartIndex2 - INI_StartIndex1)

'AYE
'A1 = INI_StartIndex
'INIA1 = A1


'RC_INI = Replace(RC_INI, "StartIndex=" + INI_StartIndex, "StartIndex=0")

'FR = FreeFile
'Open FILEPATH2 For Output Lock Read Write As #FR
'    Print #FR, RC
'Close #FR

Call INI_FILE

Call SET_INI_NUMBER_IN_LABEL_TWO

Exit Sub

'Private Sub MNU_VB_FOLDER_Click()
'
'    APP_PATH_VAR = App.Path + "\IRFANVIEW -- " + GetComputerName
'    If Dir(APP_PATH_VAR, vbDirectory) = "" Then
'        MkDir APP_PATH_VAR
'    End If
'    Shell "Explorer.exe /e," + APP_PATH_VAR, vbNormalFocus
'

End Sub






'Dim SP_FOLDER_COUNT
'
'Private Type RECT
'        Left As Long
'        Top As Long
'        Right As Long
'        Bottom As Long
'End Type
'
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
'    (ByVal lpClassName As String, _
'     ByVal lpWindowName As String) _
'    As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
'Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
'Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long


Function GetSpecialfolder_Script()
    
    Dim r As Long
    Dim IDL As ITEMIDLIST
    Dim SPECIAL_FOLDER_PATH As String
    SP_FOLDER_COUNT = 0
    
    'DEBUG
    For R2 = 0 To 255
        r = SHGetSpecialFolderLocation(100, R2, IDL)
        If r = NOERROR Then
            Path$ = Space$(512)
            r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
            SPECIAL_FOLDER_PATH = Left$(Path, InStr(Path, Chr$(0)) - 1)
            If SPECIAL_FOLDER_PATH <> "" Then
                'Debug.Print Str(R2) + " -- " + SPECIAL_FOLDER_PATH
                List1.AddItem Format(R2, "000") + " -- " + SPECIAL_FOLDER_PATH
                SP_FOLDER_COUNT = SP_FOLDER_COUNT + 1
                
                If InStr(DUPE_STRING, SPECIAL_FOLDER_PATH + " - ") > 0 Then
                    DUPE_COUNT = DUPE_COUNT + 1
                End If
                DUPE_STRING = DUPE_STRING + SPECIAL_FOLDER_PATH + " - "
            End If
        End If
    Next
    Label1.Caption = "Press -- Jump Any -- Special Folder *" + Str(SP_FOLDER_COUNT) + " -- Dupe Count *" + Str(DUPE_COUNT)
    DUPE_STRING = ""
End Function

Function GetSpecialfolder_Debug()
    
    Dim r As Long
    Dim IDL As ITEMIDLIST
    Dim SPECIAL_FOLDER_PATH As String
    SP_FOLDER_COUNT = 0
    
    'DEBUG
    For R2 = 0 To 255
        r = SHGetSpecialFolderLocation(100, R2, IDL)
        If r = NOERROR Then
            Path$ = Space$(512)
            r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
            SPECIAL_FOLDER_PATH = Left$(Path, InStr(Path, Chr$(0)) - 1)
            If SPECIAL_FOLDER_PATH <> "" Then
                Debug.Print Format(R2, "000") + " -- " + SPECIAL_FOLDER_PATH
                SP_FOLDER_COUNT = SP_FOLDER_COUNT + 1
                
                If InStr(DUPE_STRING, SPECIAL_FOLDER_PATH + " - ") > 0 Then
                    DUPE_COUNT = DUPE_COUNT + 1
                End If
                DUPE_STRING = DUPE_STRING + SPECIAL_FOLDER_PATH + " - "
            End If
        End If
    Next
    'Label1.Caption = "Press -- Jump Any -- Special Folder *" + Str(SP_FOLDER_COUNT) + " -- Dupe Count *" + Str(DUPE_COUNT)
    
    Debug.Print "Special Folder * " + Str(SP_FOLDER_COUNT) + " - -Dupe; Count * " + Str(DUPE_COUNT); ""
    DUPE_STRING = ""
End Function



Private Function GetSpecialfolder(CSIDL As Long) As String
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



'Private Sub List1_Click()
'
'    Shell "Explorer.exe /e," + Mid(List1.List(List1.ListIndex), 5), vbNormalFocus
'    Label1.Caption = Label1.Caption + " -- FOLDER LOADING"
'
'    Timer1.Enabled = True
'
'End Sub

'Sub SETUP_DIMENSIONS_NORM()
'
'    Dim RECTT1 As RECT
'    Dim RECTT2 As RECT
'    Dim RECTT4 As RECT
'
'    On Error Resume Next
'
'   ' If Me.WindowState <> vbMinimized Then
'
'        'DONT_RESIZE_WHILE_SETUP = True
'        Me.WindowState = vbNormal
'
'        Test1 = FindWindow("Shell_TrayWnd", vbNullString)
'        T1 = GetWindowRect(Test1, RECTT1) ' BOTTOM BAR
'        Test2 = FindWindow("MOM Class", vbNullString)
'        T1 = GetWindowRect(Test2, RECTT2) ' TOP BAR
'        'IRON BAR - DUMB BELLS
'        Test4 = FindWindowPart_BASEBAR("BaseBar")
'        T1 = GetWindowRect(Test4, RECTT4) ' LEFT BAR
'        A = RECTT4.Top
'        y1 = (RECTT1.Bottom - RECTT1.Top) * Screen.TwipsPerPixelY
'
'        If RECTT2.Bottom - RECTT2.Top < 33 Then
'
'            Y2 = (RECTT2.Bottom - RECTT2.Top) * Screen.TwipsPerPixelY
'
'        Else
'
'            Y2 = 0
'
'        End If
'
'        If Test4 > 0 Then
'            If RECTT4.Right = RECTT4.Left Then
'                X4 = (RECTT4.Right) * Screen.TwipsPerPixelX
'            Else
'                X4 = (RECTT4.Right - RECTT4.Left) * Screen.TwipsPerPixelX
'            End If
'        End If
'
'        SCREEN_WIDTH_SPACE = Screen.Width - X4
'        SCREEN_HEIGHT_SPACE = Screen.Height - y1 - Y2
'
'        Me.Width = SCREEN_WIDTH_SPACE - 1200
'        'Me.Height = SCREEN_HEIGHT_SPACE - 800
'        Me.Height = SCREEN_HEIGHT_SPACE - 1800
'        'THIS THE FORM HEIGHT
'        'THE BOX INSIDE IS ADJUST AFTER
'
'
'
'
'        'Form1.Left = (Screen.Width - Me.Width) / 2
'        Me.Left = X4 + (SCREEN_WIDTH_SPACE - Me.Width) / 2
'        Me.Top = Y2 + (SCREEN_HEIGHT_SPACE - Me.Height) / 2
'
''        DoEvents
'
''        Call SETUP_DIMENSIONS_INNER_FORM
'
''    End If
'
'End Sub

'Private Sub Timer1_Timer()
'    Exit Sub
'
'    Unload Me
'    Form1.WindowState = vbMinimized
'
'End Sub


Private Sub Timer1_Timer()
    Dim MhWndVAR As Long

    MhWndVAR = FindWindow("IrfanView", "IrfanView")
    If O_MhWndVAR <> MhWndVAR And MhWndVAR = 0 Then
        Call MNU_INI_READ
    End If
    O_MhWndVAR = MhWndVAR

    Dim tPA As POINTAPI
    
    ' Get cursor cordinates
    GetCursorPos tPA
    ' Set label caption to cursor cordinates
    'lblCordi.Caption = "X: " & tPA.X & "  Y: " & tPA.Y
    
    XI = GetAsyncKeyState(1) 'MOUSE LEFT BUTTON DOWN
    If tPA.Y < 100 And XI <> 0 Then

        P_COUNT = P_COUNT + 1
        Beep

        If P_COUNT > 20 Then
            'Shell FOLDER_PATH_IRFANView2 + " /killmesoftly"
            
            
            Do
                MhWndVAR = FindWindow("FullScreenClass", vbNullString)
                If MhWndVAR > 0 Then
                    If GetParentHwnd(MhWndVAR) > 0 Then
                        SendMessage GetParentHwnd(MhWndVAR), WM_CLOSE, 0, 0
                    
                    End If
                End If
            Loop Until MhWndVAR = 0
            
            
            Beep
        End If
    
    Else
        P_COUNT = 0
    End If


End Sub

Function GetParentHwnd(ByVal ReturnParent As Long) As String
   Dim I As Long
   Dim j As Long
   Dim k As Long
   I = ReturnParent
   If ReturnParent Then
      Do While I <> 0
         k = j
         j = I
         I = GetParent(I)
      Loop
    I = j
    End If
    GetParentHwnd = I
End Function

