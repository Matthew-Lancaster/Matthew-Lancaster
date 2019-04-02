VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetMatrix v1.7"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDLG 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Menu 
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4035
      ScaleWidth      =   2475
      TabIndex        =   145
      Top             =   960
      Width           =   2535
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   120
      ScaleHeight     =   690
      ScaleWidth      =   10545
      TabIndex        =   139
      Top             =   120
      Width           =   10545
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "frmMain.frx":0442
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "The net analyzer"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2520
         TabIndex        =   143
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "NetMatrix"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   142
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NetMatrix v1.7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6480
         TabIndex        =   141
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label lblNM 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "http://idrus.net"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6480
         TabIndex        =   140
         Top             =   360
         Width           =   3255
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   9720
         Picture         =   "frmMain.frx":0884
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   2760
      ScaleHeight     =   4035
      ScaleWidth      =   7875
      TabIndex        =   1
      Top             =   960
      Width           =   7935
      Begin VB.Frame fraHW 
         Caption         =   "Hardware scanner"
         Height          =   3735
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   7695
         Begin VB.TextBox txtHW 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   144
            Top             =   600
            Width           =   7455
         End
         Begin VB.CommandButton cmdSaveReports 
            Caption         =   "Sa&ve reports"
            Height          =   375
            Left            =   2760
            TabIndex        =   35
            Top             =   3240
            Width           =   2295
         End
         Begin VB.CommandButton cmdHardWareScan 
            Caption         =   "S&can hardware"
            Height          =   375
            Left            =   5160
            TabIndex        =   34
            Top             =   3240
            Width           =   2295
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Hardware scan result:"
            Height          =   255
            Left            =   120
            TabIndex        =   146
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame fraPing 
         Caption         =   "IP(s) ping-er"
         Height          =   3735
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   7695
         Begin VB.Frame Frame3_1 
            Caption         =   "Ping statistics"
            Height          =   2415
            Left            =   3000
            TabIndex        =   30
            Top             =   240
            Width           =   4575
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2055
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   31
               Top             =   240
               Width           =   4320
            End
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save (IPs)"
            Height          =   375
            Left            =   6120
            TabIndex        =   29
            Top             =   3120
            Width           =   1455
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "&Clear"
            Height          =   375
            Left            =   4560
            TabIndex        =   28
            Top             =   3120
            Width           =   1455
         End
         Begin VB.CommandButton cmdStart 
            Caption         =   "Start &ping"
            Height          =   375
            Left            =   3000
            TabIndex        =   27
            Top             =   3120
            Width           =   1455
         End
         Begin VB.Frame Frame3_2 
            Caption         =   "Ping this IP(s)"
            Height          =   3255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   2775
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   120
               ScaleHeight     =   315
               ScaleWidth      =   2475
               TabIndex        =   18
               Top             =   360
               Width           =   2535
               Begin VB.TextBox ip4 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   1920
                  MaxLength       =   3
                  TabIndex        =   22
                  Top             =   0
                  Width           =   495
               End
               Begin VB.TextBox ip3 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   1320
                  MaxLength       =   3
                  TabIndex        =   21
                  Top             =   0
                  Width           =   495
               End
               Begin VB.TextBox ip2 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   720
                  MaxLength       =   3
                  TabIndex        =   20
                  Top             =   0
                  Width           =   495
               End
               Begin VB.TextBox ip1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   19
                  Top             =   0
                  Width           =   495
               End
               Begin VB.Label Label3 
                  BackColor       =   &H80000009&
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   1800
                  TabIndex        =   25
                  Top             =   0
                  Width           =   135
               End
               Begin VB.Label Label3 
                  BackColor       =   &H80000009&
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   1200
                  TabIndex        =   24
                  Top             =   0
                  Width           =   135
               End
               Begin VB.Label Label3 
                  BackColor       =   &H80000009&
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   600
                  TabIndex        =   23
                  Top             =   0
                  Width           =   135
               End
            End
            Begin VB.CommandButton cmdRemove 
               Caption         =   "&Remove"
               Height          =   375
               Left            =   1440
               TabIndex        =   17
               Top             =   840
               Width           =   1215
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               Height          =   375
               Left            =   120
               TabIndex        =   16
               Top             =   840
               Width           =   1215
            End
            Begin VB.ListBox List1 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1230
               ItemData        =   "frmMain.frx":257E
               Left            =   120
               List            =   "frmMain.frx":2580
               TabIndex        =   15
               Top             =   1920
               Width           =   2520
            End
            Begin VB.CommandButton cmdClearList 
               Caption         =   "Clear &list"
               Height          =   375
               Left            =   120
               TabIndex        =   14
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Timer TimerTotalIP 
               Interval        =   75
               Left            =   120
               Top             =   120
            End
            Begin VB.Label lblTotalIP 
               BackStyle       =   0  'Transparent
               Caption         =   "IP loaded: 0"
               Height          =   255
               Left            =   1440
               TabIndex        =   26
               Top             =   1440
               Width           =   1215
            End
         End
         Begin MSComctlLib.ProgressBar Prog1 
            Height          =   255
            Left            =   3000
            TabIndex        =   32
            Top             =   2760
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
      End
      Begin VB.Frame fraLocal 
         Caption         =   "Local IP info"
         Height          =   3735
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   7695
         Begin VB.Frame Frame2_1 
            Caption         =   "Local host name"
            Height          =   855
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   7455
            Begin VB.CommandButton Copy1 
               Caption         =   "Copy"
               Height          =   375
               Left            =   6720
               TabIndex        =   11
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtLocalHost 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   10
               Text            =   "Local host name"
               Top             =   360
               Width           =   6495
            End
         End
         Begin VB.Frame Frame2_2 
            Caption         =   "Local IP address"
            Height          =   855
            Left            =   120
            TabIndex        =   6
            Top             =   1440
            Width           =   7455
            Begin VB.CommandButton Copy2 
               Caption         =   "Copy"
               Height          =   375
               Left            =   6720
               TabIndex        =   8
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtLocalIP 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   7
               Text            =   "Local IP address"
               Top             =   360
               Width           =   6495
            End
         End
         Begin VB.Frame Frame2_3 
            Caption         =   "Computer name"
            Height          =   855
            Left            =   120
            TabIndex        =   3
            Top             =   2520
            Width           =   7455
            Begin VB.TextBox txtUname 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   240
               Locked          =   -1  'True
               TabIndex        =   5
               Top             =   360
               Width           =   6375
            End
            Begin VB.CommandButton Copy3 
               Caption         =   "Copy"
               Height          =   375
               Left            =   6720
               TabIndex        =   4
               Top             =   360
               Width           =   615
            End
         End
      End
      Begin VB.Frame fraTCP 
         Caption         =   "TCP monitor"
         Height          =   3735
         Left            =   120
         TabIndex        =   134
         Top             =   120
         Width           =   7695
         Begin VB.Timer timAutoR 
            Interval        =   2500
            Left            =   120
            Top             =   240
         End
         Begin VB.CheckBox chkAuto 
            Caption         =   "&Auto refresh"
            Height          =   255
            Left            =   120
            TabIndex        =   137
            Top             =   3360
            Width           =   1335
         End
         Begin VB.CheckBox chkHideIP 
            Caption         =   "&Don't show rows for 0.0.0.0 and  127.0.0.1 IP addresses"
            Height          =   255
            Left            =   120
            TabIndex        =   136
            Top             =   3000
            Width           =   4335
         End
         Begin VB.CommandButton cmdGetTCPTable 
            Caption         =   "&Get TCP table"
            Height          =   375
            Left            =   4680
            TabIndex        =   135
            Top             =   3120
            Width           =   2895
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2535
            Left            =   120
            TabIndex        =   138
            Top             =   360
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   4471
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Local IP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Local Port"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Remote IP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Remote Port"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "State"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame fraScan 
         Caption         =   "Ports scanner"
         Height          =   3735
         Left            =   120
         TabIndex        =   106
         Top             =   120
         Visible         =   0   'False
         Width           =   7695
         Begin VB.CommandButton cmdScan 
            Caption         =   "&Scan"
            Height          =   375
            Left            =   120
            TabIndex        =   131
            Top             =   3240
            Width           =   2775
         End
         Begin VB.Timer TimerScan 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   7200
            Top             =   240
         End
         Begin VB.Frame Frame5_3 
            Caption         =   "Scan status"
            Height          =   1335
            Left            =   3000
            TabIndex        =   126
            Top             =   2160
            Width           =   4575
            Begin VB.Label lblScannedPorts 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               Height          =   255
               Left            =   2040
               TabIndex        =   130
               Top             =   360
               Width           =   2295
            End
            Begin VB.Label lblOpenPorts 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               Height          =   255
               Left            =   2040
               TabIndex        =   129
               Top             =   840
               Width           =   2295
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Now scanning on port:"
               Height          =   255
               Left            =   240
               TabIndex        =   128
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Open ports total:"
               Height          =   255
               Left            =   240
               TabIndex        =   127
               Top             =   840
               Width           =   1215
            End
         End
         Begin VB.Frame Frame5_2 
            Caption         =   "Open ports"
            Height          =   1815
            Left            =   3000
            TabIndex        =   123
            Top             =   240
            Width           =   4575
            Begin VB.Timer timerPercent 
               Enabled         =   0   'False
               Interval        =   1
               Left            =   3720
               Top             =   0
            End
            Begin VB.ListBox lstPort 
               Height          =   1035
               Left            =   120
               TabIndex        =   124
               TabStop         =   0   'False
               Top             =   360
               Width           =   4335
            End
            Begin VB.Label lblDef 
               BackStyle       =   0  'Transparent
               Caption         =   "Opened port definition..."
               Height          =   255
               Left            =   120
               TabIndex        =   125
               Top             =   1440
               Width           =   4335
            End
         End
         Begin VB.Frame Frame5_1 
            Caption         =   "Scan settings"
            Height          =   2295
            Left            =   120
            TabIndex        =   107
            Top             =   240
            Width           =   2775
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   120
               ScaleHeight     =   315
               ScaleWidth      =   2475
               TabIndex        =   112
               Top             =   600
               Width           =   2535
               Begin VB.TextBox TextIP4 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   1920
                  MaxLength       =   3
                  TabIndex        =   116
                  Text            =   "1"
                  Top             =   0
                  Width           =   495
               End
               Begin VB.TextBox TextIP3 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   1320
                  MaxLength       =   3
                  TabIndex        =   115
                  Text            =   "0"
                  Top             =   0
                  Width           =   495
               End
               Begin VB.TextBox TextIP2 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   720
                  MaxLength       =   3
                  TabIndex        =   114
                  Text            =   "0"
                  Top             =   0
                  Width           =   495
               End
               Begin VB.TextBox TextIP1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   113
                  Text            =   "127"
                  Top             =   0
                  Width           =   495
               End
               Begin VB.Label Label3 
                  BackColor       =   &H80000009&
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   5
                  Left            =   1800
                  TabIndex        =   119
                  Top             =   0
                  Width           =   135
               End
               Begin VB.Label Label3 
                  BackColor       =   &H80000009&
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   1200
                  TabIndex        =   118
                  Top             =   0
                  Width           =   135
               End
               Begin VB.Label Label3 
                  BackColor       =   &H80000009&
                  Caption         =   "."
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   600
                  TabIndex        =   117
                  Top             =   0
                  Width           =   135
               End
            End
            Begin VB.CheckBox chkBeep 
               Caption         =   "Beep for found open port(s)"
               Height          =   255
               Left            =   240
               TabIndex        =   111
               Top             =   1800
               Value           =   1  'Checked
               Width           =   2295
            End
            Begin VB.TextBox txtTimeout 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1440
               MaxLength       =   3
               TabIndex        =   110
               Text            =   "1"
               Top             =   1440
               Width           =   1215
            End
            Begin VB.TextBox txtLimit 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1440
               MaxLength       =   5
               TabIndex        =   109
               Text            =   "10000"
               Top             =   1080
               Width           =   1215
            End
            Begin VB.TextBox txtIP 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   108
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Host to scan:"
               Height          =   255
               Left            =   120
               TabIndex        =   122
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Timeout:"
               Height          =   195
               Left            =   240
               TabIndex        =   121
               Top             =   1440
               Width           =   570
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Stop scan on:"
               Height          =   195
               Left            =   240
               TabIndex        =   120
               Top             =   1080
               Width           =   1005
            End
         End
         Begin MSComctlLib.ProgressBar bar 
            Height          =   255
            Left            =   120
            TabIndex        =   132
            Top             =   2760
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label lblPercent 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            Height          =   255
            Left            =   2280
            TabIndex        =   133
            Top             =   2760
            Width           =   495
         End
      End
      Begin VB.Frame fraProcMon 
         Caption         =   "System process monitor"
         Enabled         =   0   'False
         Height          =   3735
         Left            =   120
         TabIndex        =   101
         Top             =   120
         Visible         =   0   'False
         Width           =   7695
         Begin VB.Timer TimerSysMon 
            Enabled         =   0   'False
            Interval        =   2500
            Left            =   1920
            Top             =   3360
         End
         Begin VB.CheckBox chkProc 
            Caption         =   "Auto refres&h"
            Height          =   255
            Left            =   240
            TabIndex        =   104
            Top             =   3240
            Width           =   1335
         End
         Begin VB.CommandButton cmdEnd 
            Caption         =   "&End processes"
            Height          =   375
            Left            =   5520
            TabIndex        =   103
            Top             =   3120
            Width           =   2055
         End
         Begin VB.CommandButton cmdView 
            Caption         =   "&View processes"
            Height          =   375
            Left            =   3360
            TabIndex        =   102
            Top             =   3120
            Width           =   2055
         End
         Begin MSComctlLib.ListView lvw 
            Height          =   2655
            Left            =   120
            TabIndex        =   105
            Top             =   360
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   4683
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Process"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Process ID"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame fraSweep 
         Caption         =   "Port sweeper"
         Enabled         =   0   'False
         Height          =   3735
         Left            =   120
         TabIndex        =   81
         Top             =   120
         Visible         =   0   'False
         Width           =   7695
         Begin VB.CommandButton cmdResetPS 
            Caption         =   "Rese&t all"
            Height          =   375
            Left            =   4440
            TabIndex        =   99
            Top             =   3240
            Width           =   3135
         End
         Begin VB.Frame Frame12 
            Caption         =   "Found servers"
            Height          =   2895
            Left            =   4440
            TabIndex        =   97
            Top             =   240
            Width           =   3135
            Begin VB.ListBox lstPS 
               Height          =   2400
               ItemData        =   "frmMain.frx":2582
               Left            =   120
               List            =   "frmMain.frx":2584
               TabIndex        =   98
               Top             =   360
               Width           =   2895
            End
         End
         Begin VB.Timer TimPS 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   7200
            Top             =   3240
         End
         Begin VB.Frame Frame11 
            Caption         =   "Settings"
            Height          =   1215
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   4215
            Begin VB.TextBox txtPSInterval 
               Height          =   285
               Left            =   3240
               MaxLength       =   4
               TabIndex        =   92
               Text            =   "200"
               Top             =   720
               Width           =   735
            End
            Begin VB.TextBox txtPSPort 
               Height          =   285
               Left            =   3240
               MaxLength       =   5
               TabIndex        =   91
               Text            =   "135"
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox txtToIP2 
               Height          =   285
               Left            =   1920
               MaxLength       =   3
               TabIndex        =   90
               Text            =   "255"
               ToolTipText     =   "Up to 254"
               Top             =   720
               Width           =   375
            End
            Begin VB.TextBox txtToIP1 
               Height          =   285
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   89
               Text            =   "127.0.0"
               Top             =   720
               Width           =   1215
            End
            Begin VB.TextBox txtFromIP 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   720
               TabIndex        =   88
               Text            =   "127.0.0.1"
               ToolTipText     =   "Format must be xxx.xxx.xxx.xxx"
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label10 
               Caption         =   "Timeout:"
               Height          =   255
               Left            =   2520
               TabIndex        =   96
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label25 
               BackStyle       =   0  'Transparent
               Caption         =   "Port:"
               Height          =   255
               Left            =   2520
               TabIndex        =   95
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "To:"
               Height          =   255
               Left            =   240
               TabIndex        =   94
               Top             =   720
               Width           =   255
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "From:"
               Height          =   255
               Left            =   240
               TabIndex        =   93
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Progress"
            Height          =   1575
            Left            =   120
            TabIndex        =   82
            Top             =   1560
            Width           =   4215
            Begin VB.TextBox txtPSIP 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   240
               Locked          =   -1  'True
               TabIndex        =   85
               Top             =   600
               Width           =   3735
            End
            Begin VB.CommandButton cmdPauseScan 
               Caption         =   "&Pause Scan"
               Enabled         =   0   'False
               Height          =   375
               Left            =   2160
               TabIndex        =   84
               Top             =   960
               Width           =   1815
            End
            Begin VB.CommandButton cmdStartScan 
               Caption         =   "Start S&can"
               Height          =   375
               Left            =   240
               TabIndex        =   83
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Current scanned IP:"
               Height          =   255
               Left            =   240
               TabIndex        =   86
               Top             =   360
               Width           =   1575
            End
         End
         Begin MSWinsockLib.Winsock SockPS 
            Left            =   6720
            Top             =   3240
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Sweep the hacking target!"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   100
            Top             =   3240
            Width           =   4215
         End
      End
      Begin VB.Frame fraMsg 
         Caption         =   "Messenger exploit"
         Height          =   3735
         Left            =   120
         TabIndex        =   57
         Top             =   120
         Width           =   7695
         Begin VB.Frame Frame10 
            Caption         =   "Message bomb"
            Height          =   2055
            Left            =   120
            TabIndex        =   67
            Top             =   1560
            Width           =   7455
            Begin VB.TextBox txtBombQuantity 
               Height          =   285
               Left            =   4800
               MaxLength       =   3
               TabIndex        =   73
               Text            =   "13"
               ToolTipText     =   "Be careful, setting quantity to its maximum value (999) could 'eat' so many resources of your computer"
               Top             =   1080
               Width           =   855
            End
            Begin VB.TextBox txtBombIP 
               Height          =   285
               Left            =   1320
               MaxLength       =   15
               TabIndex        =   72
               Text            =   "127.0.0.1"
               Top             =   360
               Width           =   1455
            End
            Begin VB.TextBox txtBombMsg 
               Height          =   285
               Left            =   1320
               MaxLength       =   125
               TabIndex        =   71
               Text            =   "Bomb...bomb...bomb..."
               Top             =   720
               Width           =   5895
            End
            Begin VB.TextBox txtBombLength 
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   70
               Text            =   "0 character of 125 allowed"
               Top             =   1080
               Width           =   2295
            End
            Begin VB.CommandButton cmdbomb 
               Caption         =   "Start &bombing"
               Height          =   375
               Left            =   240
               TabIndex        =   69
               Top             =   1560
               Width           =   6975
            End
            Begin VB.CommandButton cmdBombClear 
               Caption         =   "&Clear"
               Height          =   375
               Left            =   5760
               TabIndex        =   68
               Top             =   1080
               Width           =   1455
            End
            Begin MSComctlLib.ProgressBar barbomb 
               Height          =   255
               Left            =   3720
               TabIndex        =   74
               Top             =   360
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   1
               Scrolling       =   1
            End
            Begin VB.Label lblBarbomb 
               BackStyle       =   0  'Transparent
               Caption         =   "0%"
               Height          =   255
               Left            =   6720
               TabIndex        =   80
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label21 
               BackStyle       =   0  'Transparent
               Caption         =   "Progress:"
               Height          =   255
               Left            =   2880
               TabIndex        =   79
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "Quantity:"
               Height          =   255
               Left            =   3960
               TabIndex        =   78
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Receiver's IP:"
               Height          =   255
               Left            =   240
               TabIndex        =   77
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "Text to send:"
               Height          =   255
               Left            =   240
               TabIndex        =   76
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "Msg length:"
               Height          =   255
               Left            =   240
               TabIndex        =   75
               Top             =   1080
               Width           =   975
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Message send"
            Height          =   1215
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   7455
            Begin VB.CommandButton cmdNetSendClear 
               Caption         =   "&Clear"
               Height          =   375
               Left            =   6240
               TabIndex        =   63
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox txtNetSendLength 
               Height          =   285
               Left            =   3840
               Locked          =   -1  'True
               TabIndex        =   62
               Text            =   "0 character of 125 allowed"
               Top             =   360
               Width           =   2295
            End
            Begin VB.CommandButton cmdNetSendSendMsg 
               Caption         =   "Send mes&sage"
               Height          =   375
               Left            =   5400
               TabIndex        =   61
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox txtNetSendMsg 
               Height          =   285
               Left            =   1320
               MaxLength       =   125
               TabIndex        =   60
               Text            =   "Hello..."
               Top             =   720
               Width           =   3975
            End
            Begin VB.TextBox txtNetSendIP 
               Height          =   285
               Left            =   1320
               MaxLength       =   15
               TabIndex        =   59
               Text            =   "127.0.0.1"
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Msg length:"
               Height          =   255
               Left            =   2880
               TabIndex        =   66
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Text to send:"
               Height          =   255
               Left            =   240
               TabIndex        =   65
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Receiver's IP:"
               Height          =   255
               Left            =   240
               TabIndex        =   64
               Top             =   360
               Width           =   1095
            End
         End
      End
      Begin VB.Frame fraAdvScan 
         Caption         =   "Advanced port scanner"
         Enabled         =   0   'False
         Height          =   3735
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   7695
         Begin VB.Frame fraMaxConnections 
            Caption         =   "Max connections"
            Height          =   855
            Left            =   2640
            TabIndex        =   53
            Top             =   240
            Width           =   1575
            Begin VB.ComboBox Mc 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   54
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.CommandButton SS 
            Cancel          =   -1  'True
            Caption         =   "S&top"
            Height          =   375
            Left            =   2640
            TabIndex        =   52
            Top             =   1680
            Width           =   1575
         End
         Begin VB.CommandButton SN 
            Caption         =   "S&can"
            Default         =   -1  'True
            Height          =   375
            Left            =   2640
            TabIndex        =   51
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Frame fraOpenPorts 
            Caption         =   "Open ports"
            Height          =   3015
            Left            =   4320
            TabIndex        =   48
            Top             =   240
            Width           =   3255
            Begin VB.ListBox OP 
               Height          =   2205
               ItemData        =   "frmMain.frx":2586
               Left            =   120
               List            =   "frmMain.frx":2588
               TabIndex        =   50
               Top             =   240
               Width           =   3015
            End
            Begin VB.CommandButton CL 
               Caption         =   "&Clear list"
               Height          =   375
               Left            =   120
               TabIndex        =   49
               Top             =   2520
               Width           =   3015
            End
         End
         Begin VB.Frame fraScanPorts 
            Caption         =   "Port range"
            Height          =   855
            Left            =   120
            TabIndex        =   44
            Top             =   1200
            Width           =   2415
            Begin VB.TextBox P2 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   1320
               MaxLength       =   5
               TabIndex        =   46
               Text            =   "32676"
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox P1 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   120
               MaxLength       =   5
               TabIndex        =   45
               Text            =   "1"
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblTo 
               Alignment       =   2  'Center
               Caption         =   "to"
               Height          =   255
               Left            =   1080
               TabIndex        =   47
               Top             =   360
               Width           =   255
            End
         End
         Begin VB.Frame fraRemoteIP 
            Caption         =   "Remote IP"
            Height          =   855
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   2415
            Begin VB.TextBox IP 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   120
               MaxLength       =   15
               TabIndex        =   43
               Text            =   "127.0.0.1"
               Top             =   360
               Width           =   2175
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Statistics"
            Height          =   1335
            Left            =   120
            TabIndex        =   37
            Top             =   2160
            Width           =   4095
            Begin VB.Timer TimProgress 
               Enabled         =   0   'False
               Interval        =   1
               Left            =   3600
               Top             =   840
            End
            Begin VB.Label lblAdvPS 
               BackStyle       =   0  'Transparent
               Caption         =   "Now scanning on port: 0"
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
               Left            =   120
               TabIndex        =   41
               Top             =   240
               Width           =   3855
            End
            Begin VB.Label lblTotalPS 
               BackStyle       =   0  'Transparent
               Caption         =   "Total open port found: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   720
               Width           =   3855
            End
            Begin VB.Label lblIPPS 
               BackStyle       =   0  'Transparent
               Caption         =   "At IP: 127.0.0.1"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   960
               Width           =   3855
            End
            Begin VB.Label lblTPTS 
               BackStyle       =   0  'Transparent
               Caption         =   "Total port to scan: 0"
               Height          =   255
               Left            =   120
               TabIndex        =   38
               Top             =   480
               Width           =   3855
            End
         End
         Begin MSWinsockLib.Winsock S1 
            Index           =   0
            Left            =   240
            Top             =   1320
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin MSComctlLib.ProgressBar PB 
            Height          =   255
            Left            =   4320
            TabIndex        =   55
            Top             =   3360
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label lblPerc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0 %"
            Height          =   255
            Left            =   7080
            TabIndex        =   56
            Top             =   3360
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5460
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   450
      SimpleText      =   "NetMatrix"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8916
            Text            =   "Panel"
            TextSave        =   "Panel"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3440
            MinWidth        =   3440
            Text            =   "DateTime"
            TextSave        =   "DateTime"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   1323
            MinWidth        =   1323
            Text            =   "Num"
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
            Text            =   "Network status"
            TextSave        =   "Network status"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "NetMatrix"
            TextSave        =   "NetMatrix"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsknetworkstatus 
      Index           =   0
      Left            =   240
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer TimerStatus 
      Interval        =   75
      Left            =   600
      Top             =   960
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   960
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer timDateTime 
      Interval        =   10
      Left            =   1320
      Top             =   960
   End
   Begin VB.Timer timAnim 
      Interval        =   2500
      Left            =   1680
      Top             =   960
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Port sweeper
Dim Comparison
Dim Comparison2
Dim Comparison3
Dim Comparison4
Dim IPToScan As Integer
Dim IPToStopOn As Integer
Dim Paused
Dim TimeOut
Dim StartString

'Hardware scanner
Dim DeviceFound() As Variant
Dim DeviceList() As Variant
Dim DeviCecount As Integer
Dim ramas As Variant
Dim ramotipas As Variant
Dim PelesInt() As Variant
Dim PelesTipas() As Variant
Dim objWshNet, DeviceListLen, strServer, objService, isconnect, objDeviceSet, Device

'Message box
Dim msg As Byte

'Ports scanner
Dim port As Integer
Dim portnumber

'PC users
Private Declare Function GetProfilesDirectory Lib "userenv.dll" Alias "GetProfilesDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean
Private objFSO As FileSystemObject
Private objFolders As Folders, objFolder As Folder
Dim ECHO As ICMP_ECHO_REPLY
Dim tstr, i, a

'Get TCP table codes
Private Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type
'
Private Const ERROR_BUFFER_OVERFLOW = 111&
Private Const ERROR_INVALID_PARAMETER = 87
Private Const ERROR_NO_DATA = 232&
Private Const ERROR_NOT_SUPPORTED = 50&
Private Const ERROR_SUCCESS = 0&
'
Private Const MIB_TCP_STATE_CLOSED = 1
Private Const MIB_TCP_STATE_LISTEN = 2
Private Const MIB_TCP_STATE_SYN_SENT = 3
Private Const MIB_TCP_STATE_SYN_RCVD = 4
Private Const MIB_TCP_STATE_ESTAB = 5
Private Const MIB_TCP_STATE_FIN_WAIT1 = 6
Private Const MIB_TCP_STATE_FIN_WAIT2 = 7
Private Const MIB_TCP_STATE_CLOSE_WAIT = 8
Private Const MIB_TCP_STATE_CLOSING = 9
Private Const MIB_TCP_STATE_LAST_ACK = 10
Private Const MIB_TCP_STATE_TIME_WAIT = 11
Private Const MIB_TCP_STATE_DELETE_TCB = 12
'
Private Declare Function GetTcpTable Lib "iphlpapi.dll" (ByRef pTcpTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)
'
'Port definition
Private Declare Function OSGetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Const nBUFSIZEINI = 1024
Private Const nBUFSIZEINIALL = 4096

Private FilePathName As String
Private Const Packet = "QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!QQ-OWNS!!!!!!"

Private Sub Pause(interval)
'allows us to pause a loop or proceedure
Dim CurrentTime As Variant
  CurrentTime = Timer
    Do While Timer - CurrentTime < Val(interval)
      DoEvents:
    Loop
End Sub

Private Function GetPrivateProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String, ByVal szFileName As String) As String
   ' *** Get an entry in the inifile ***

   Dim szTmp                     As String
   Dim nRet                      As Long

   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetPrivateProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL, szFileName)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetPrivateProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI, szFileName)
   End If
   GetPrivateProfileString = Left$(szTmp, nRet)

End Function

Private Sub chkAuto_Click()
If chkAuto.Value = vbChecked Then
timAutoR.Enabled = True
Else
timAutoR.Enabled = False
End If
End Sub

Private Sub chkProc_Click()
If chkProc.Value = vbChecked Then
TimerSysMon.Enabled = True
Else
TimerSysMon.Enabled = False
End If
End Sub

Private Sub cmdAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub cmdBombClear_Click()
txtBombIP.Text = ""
txtBombMsg.Text = ""
End Sub

Private Sub cmdClearList_Click()
List1.Clear

'Total IP
lblTotalIP.Caption = "Total IP: " & List1.ListCount
End Sub

Private Sub cmdEnd_Click()
Dim process_terminate, hand
Dim i As Long
hand = OpenProcess(process_terminate, True, CLng(lvw.SelectedItem.SubItems(1))) 'get process handle
TerminateProcess hand, 0      'end define process
Call cmdView_Click
End Sub

Private Sub cmdExit_Click()
SockPS.Close
S1(0).Close
SocketsCleanup
End
End Sub

Private Sub cmdGetTCPTable_Click()
On Error GoTo doctor
    '
    Dim arrBuffer() As Byte
    Dim lngSize As Long
    Dim lngRetVal As Long
    Dim lngRows As Long
    Dim i As Long
    Dim TcpTableRow As MIB_TCPROW
    Dim lvItem As ListItem
    '
    ListView1.ListItems.Clear
    '
    lngSize = 0
    '
    'Call the GetTcpTable just to get the buffer size into the lngSize variable
    lngRetVal = GetTcpTable(ByVal 0&, lngSize, 0)
    '
    If lngRetVal = ERROR_NOT_SUPPORTED Then
        '
        'This API works only on Win 98//2000 and NT4 with SP4
        msg = MsgBox("IP Helper is not supported by this system.", vbCritical, "NetMatrix")
        Exit Sub
        '
    End If
    '
    'Prepare the buffer
    ReDim arrBuffer(0 To lngSize - 1) As Byte
    '
    'And call the function one more time
    lngRetVal = GetTcpTable(arrBuffer(0), lngSize, 0)
    '
    If lngRetVal = ERROR_SUCCESS Then
        '
        'The first 4 bytes contain the quantity of the table rows
        'Get that value to the lngRows variable
        CopyMemory lngRows, arrBuffer(0), 4
        '
        For i = 1 To lngRows
            '
            'Copy the table row data to the TcpTableRow structure
            CopyMemory TcpTableRow, arrBuffer(4 + (i - 1) * Len(TcpTableRow)), Len(TcpTableRow)
            '
            If Not ((chkHideIP.Value = vbChecked) And (GetIpFromLong(TcpTableRow.dwLocalAddr) = "0.0.0.0" Or GetIpFromLong(TcpTableRow.dwLocalAddr) = "127.0.0.1")) Then
                '
                'Add the data to the ListView control
                With TcpTableRow
                    Set lvItem = ListView1.ListItems.Add(, , GetIpFromLong(.dwLocalAddr))
                    lvItem.SubItems(1) = GetTcpPortNumber(.dwLocalPort)
                    lvItem.SubItems(2) = GetIpFromLong(.dwRemoteAddr)
                    lvItem.SubItems(3) = GetTcpPortNumber(.dwRemotePort)
                    lvItem.SubItems(4) = GetState(.dwState)
                End With
                '
            End If
            '
        Next i
        '
    End If
    '
Exit Sub
doctor:
msg = MsgBox("An error occured while executing commands." & vbCrLf & "Operation terminated!", vbCritical, "NetMatrix")
End Sub

Private Function GetIpFromLong(lngIPAddress As Long) As String
    '
    Dim arrIpParts(3) As Byte
    '
    CopyMemory arrIpParts(0), lngIPAddress, 4
    '
    GetIpFromLong = CStr(arrIpParts(0)) & "." & CStr(arrIpParts(1)) & "." & CStr(arrIpParts(2)) & "." & CStr(arrIpParts(3))
    '
End Function

Private Function GetState(lngState As Long) As String
    '
    Select Case lngState
        Case MIB_TCP_STATE_CLOSED: GetState = "CLOSED"
        Case MIB_TCP_STATE_LISTEN: GetState = "LISTEN"
        Case MIB_TCP_STATE_SYN_SENT: GetState = "SYN_SENT"
        Case MIB_TCP_STATE_SYN_RCVD: GetState = "SYN_RCVD"
        Case MIB_TCP_STATE_ESTAB: GetState = "ESTAB"
        Case MIB_TCP_STATE_FIN_WAIT1: GetState = "FIN_WAIT1"
        Case MIB_TCP_STATE_FIN_WAIT2: GetState = "FIN_WAIT2"
        Case MIB_TCP_STATE_CLOSE_WAIT: GetState = "CLOSE_WAIT"
        Case MIB_TCP_STATE_CLOSING: GetState = "CLOSING"
        Case MIB_TCP_STATE_LAST_ACK: GetState = "LAST_ACK"
        Case MIB_TCP_STATE_TIME_WAIT: GetState = "TIME_WAIT"
        Case MIB_TCP_STATE_DELETE_TCB: GetState = "DELETE_TCB"
    End Select
    '
End Function

Private Function GetTcpPortNumber(DWord As Long) As Long
    GetTcpPortNumber = DWord / 256 + (DWord Mod 256) * 256
End Function
'End of get TCP table codes

Private Sub cmdNetSendClear_Click()
txtNetSendIP.Text = ""
txtNetSendMsg.Text = ""
End Sub

Private Sub cmdNetSendSendMsg_Click()
On Error GoTo doctor
txtNetSendMsg.Enabled = False
txtNetSendIP.Enabled = False
Shell "net send " & txtNetSendIP.Text & " " & txtNetSendMsg.Text ', vbHide
msg = MsgBox("Message sent successfully.", vbInformation, "NetMatrix")
txtNetSendMsg.Enabled = True
txtNetSendIP.Enabled = True
Exit Sub
doctor:
msg = MsgBox("Unable to execute command, invalid operating system.", vbExclamation, "NetMatrix")
End Sub

Private Sub cmdPauseScan_Click()
SockPS.Close
txtFromIP.Enabled = True
txtToIP1.Enabled = True
txtFromIP.Text = txtToIP1.Text & "." & (IPToScan - 1)
IPToScan = IPToScan - 1
txtToIP2.Enabled = True
txtPSPort.Enabled = True
txtPSInterval.Enabled = True
cmdStartScan.Enabled = True
txtPSIP.Text = "Paused"
TimPS.Enabled = False
cmdPauseScan.Enabled = False
cmdResetPS.Enabled = True
Paused = 1
End Sub

Private Sub cmdResetPS_Click()
txtFromIP.Text = ""
txtToIP1.Text = ""
txtToIP2.Text = ""
txtPSPort.Text = ""
txtPSInterval.Text = ""
lstPS.Clear
txtPSIP.Text = "Reset done!"
End Sub

Private Sub cmdSaveReports_Click()
    Dim i As Long
    Dim sFile As String
With CDLG
    .CancelError = False
    .Filter = "Text document (*.txt)|*.txt|Log file (*.log) |*.log|All files (*.*) |*.*|"
    .ShowSave
    If .FileName = "" Then
    Exit Sub
    End If
End With
    sFile = CDLG.FileName
    CDLG.FileName = ""
    On Error GoTo error
    Open sFile For Output As #1
    Print #1, txtHW.Text
error:
    Close #1
End Sub

Private Sub cmdScan_Click()
On Error GoTo doctor
txtIP.Text = TextIP1.Text & "." & TextIP2.Text & "." & TextIP3.Text & "." & TextIP4.Text
If cmdScan.Caption = "&Stop" Then
TimerScan.Enabled = False
Picture2.Enabled = True
txtLimit.Enabled = True
txtTimeout.Enabled = True
chkBeep.Enabled = True
timerPercent.Enabled = False
cmdScan.Caption = "&Scan"
Exit Sub
ElseIf cmdScan.Caption = "&Scan" Then
port = 0
lblOpenPorts.Caption = "0"
lblScannedPorts.Caption = "0"
lstPort.Clear
TimerScan.interval = txtTimeout.Text
Picture2.Enabled = False
txtLimit.Enabled = False
txtTimeout.Enabled = False
chkBeep.Enabled = False
cmdScan.Caption = "&Stop"
TimerScan.Enabled = True
timerPercent.Enabled = True
End If
Exit Sub
doctor:
msg = MsgBox("An error occured, please make sure that you fill valid informations.", vbCritical, "NetMatrix")
End Sub

'PORT SWEEPER 2
'==============
Private Sub cmdStartScan_2_Click()
On Error GoTo doctor
SockPS.Close
SockPS.RemotePort = txtPSPort.Text
TimPS.Enabled = True
IPToStopOn = txtToIP2.Text

StartString = txtFromIP.Text
StartString = "192.168.1.1"
txtFromIP.Text = StartString
For R = 1 To 255
IPToScan = "192.168.1." + Trim(Str(R))


Comparison4 = txtFromIP.Text Like "*[.]*"
Do Until Comparison4 = False
txtFromIP.SelStart = 0
txtFromIP.SelLength = 1
txtFromIP.SelText = ""
Comparison4 = txtFromIP.Text Like "*[.]*"
Loop
If Paused <> 1 Then
IPToScan = txtFromIP.Text
End If
txtFromIP.Enabled = False
txtToIP1.Enabled = False
txtToIP2.Enabled = False
txtPSPort.Enabled = False
txtPSInterval.Enabled = False
cmdStartScan.Enabled = False
cmdPauseScan.Enabled = True
cmdResetPS.Enabled = False
If Paused <> 1 Then
lstPS.Clear
End If
If Paused <> 1 Then
txtPSIP.Text = "Starting..."
End If
Paused = 0
txtFromIP.Text = StartString
Exit Sub
doctor:
msg = MsgBox("Please fill in all required information first then try again.", vbExclamation, "NetMatrix")
End Sub


'PORT SWEEPER
'============
Private Sub cmdStartScan_Click()
On Error GoTo doctor
SockPS.Close
SockPS.RemotePort = txtPSPort.Text
TimPS.Enabled = True
IPToStopOn = txtToIP2.Text

StartString = txtFromIP.Text
StartString = "192.168.1.1"
txtFromIP.Text = StartString

Comparison4 = txtFromIP.Text Like "*[.]*"
Do Until Comparison4 = False
txtFromIP.SelStart = 0
txtFromIP.SelLength = 1
txtFromIP.SelText = ""
Comparison4 = txtFromIP.Text Like "*[.]*"
Loop
If Paused <> 1 Then
IPToScan = txtFromIP.Text
End If
txtFromIP.Enabled = False
txtToIP1.Enabled = False
txtToIP2.Enabled = False
txtPSPort.Enabled = False
txtPSInterval.Enabled = False
cmdStartScan.Enabled = False
cmdPauseScan.Enabled = True
cmdResetPS.Enabled = False
If Paused <> 1 Then
lstPS.Clear
End If
If Paused <> 1 Then
txtPSIP.Text = "Starting..."
End If
Paused = 0
txtFromIP.Text = StartString
Exit Sub
doctor:
msg = MsgBox("Please fill in all required information first then try again.", vbExclamation, "NetMatrix")
End Sub



Private Sub cmdView_Click()
Dim theloop, ret As String
Dim i As Long
Dim proc As PROCESSENTRY32
Dim snap As Long
Dim exename As String
lvw.ListItems.Clear 'clear listview contents
snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0) 'get snapshot handle
proc.dwSize = Len(proc)
theloop = ProcessFirst(snap, proc)       'first process and return value
i = 0
While theloop <> 0      'next process
exename = proc.szExeFile
ret = lvw.ListItems.Add(, "first" & CStr(i), exename)   'add process name to listview
lvw.ListItems("first" & CStr(i)).SubItems(1) = proc.th32ProcessID   'add process ID to listview
i = i + 1
theloop = ProcessNext(snap, proc)
Wend
CloseHandle snap       'close snapshot handle
End Sub

Private Sub cmdbomb_Click()
On Error GoTo doctor
cmdbomb.Enabled = False
txtBombIP.Enabled = False
txtBombMsg.Enabled = False
txtBombQuantity.Enabled = False
For i = 1 To txtBombQuantity.Text
DoEvents
Shell "net send " & txtBombIP.Text & " " & txtBombMsg.Text, vbHide
Dim BarV As Byte
BarV = (i / txtBombQuantity.Text) * 100
barbomb.Value = BarV
lblBarbomb.Caption = BarV & "%"
If barbomb.Value = barbomb.Max Then
msg = MsgBox(txtBombQuantity.Text & " bomb message sent successfully.", vbInformation, "NetMatrix")
cmdbomb.Enabled = True
txtBombIP.Enabled = True
txtBombMsg.Enabled = True
txtBombQuantity.Enabled = True
End If
Next i
Exit Sub
doctor:
msg = MsgBox("File not found (net.exe) or invalid operating system.", vbExclamation, "NetMatrix")
End Sub

Private Sub cmdHardwareScan_Click()
Call scanhard
End Sub

Private Sub Copy1_Click()
On Error Resume Next
Clipboard.SetText txtLocalHost.Text
End Sub

Private Sub Copy2_Click()
On Error Resume Next
Clipboard.SetText txtLocalIP.Text
End Sub

Private Sub Copy3_Click()
On Error Resume Next
Clipboard.SetText txtUname.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
SockPS.Close
SocketsCleanup
End
End Sub

Private Sub lstPort_Click()
Dim def As String
On Error GoTo doctor
Dim AppDir As String
AppDir = App.Path

FilePathName = AppDir + "\PortDef.ini"
def = GetPrivateProfileString("Definition", lstPort.Text, "", FilePathName)

DoEvents
lblDef.Caption = def
Exit Sub

doctor:
lblDef.Caption = "Unknown opened port"
End Sub

Private Sub Menu_MenuItemClick(MenuNumber As Long, MenuItem As Long)
If MenuNumber = 1 Then

  If MenuItem = 1 Then
    fraLocal.Visible = True
    fraLocal.Enabled = True
    status.Panels(1).Text = "Show local IP, local host name, and local computer name"
    Me.Caption = App.FileDescription & " - Local info"
    fraProcMon.Visible = False
    fraProcMon.Enabled = False
    fraHW.Visible = False
    fraHW.Enabled = False
    fraPing.Visible = False
    fraPing.Enabled = False
    fraScan.Visible = False
    fraScan.Enabled = False
    fraAdvScan.Visible = False
    fraAdvScan.Enabled = False
    fraSweep.Enabled = False
    fraSweep.Visible = False
    fraMsg.Enabled = False
    fraMsg.Visible = False
    fraTCP.Visible = False
    fraTCP.Enabled = False

  ElseIf MenuItem = 2 Then
    status.Panels(1).Text = "Show current processes on your computer"
    Me.Caption = App.FileDescription & " - System processes monitor"
    fraProcMon.Visible = True
    fraProcMon.Enabled = True
    fraLocal.Visible = False
    fraLocal.Enabled = False
    fraHW.Visible = False
    fraHW.Enabled = False
    fraPing.Visible = False
    fraPing.Enabled = False
    fraScan.Visible = False
    fraScan.Enabled = False
    fraAdvScan.Visible = False
    fraAdvScan.Enabled = False
    fraSweep.Enabled = False
    fraSweep.Visible = False
    fraMsg.Enabled = False
    fraMsg.Visible = False
    fraTCP.Visible = False
    fraTCP.Enabled = False

  ElseIf MenuItem = 3 Then
    status.Panels(1).Text = "Show local system's hardware informations"
    Me.Caption = App.FileDescription & " - Hardware scanner"
    fraHW.Visible = True
    fraHW.Enabled = True
    fraProcMon.Visible = False
    fraProcMon.Enabled = False
    fraLocal.Visible = False
    fraLocal.Enabled = False
    fraPing.Visible = False
    fraPing.Enabled = False
    fraScan.Visible = False
    fraScan.Enabled = False
    fraAdvScan.Visible = False
    fraAdvScan.Enabled = False
    fraSweep.Enabled = False
    fraSweep.Visible = False
    fraMsg.Enabled = False
    fraMsg.Visible = False
    fraTCP.Visible = False
    fraTCP.Enabled = False

  End If
  
ElseIf MenuNumber = 2 Then
  If MenuItem = 1 Then
    status.Panels(1).Text = "Ping specified IP's fast and easy"
    Me.Caption = App.FileDescription & " - IP(s) ping-er"
    fraPing.Visible = True
    fraPing.Enabled = True
    fraHW.Visible = False
    fraHW.Enabled = False
    fraProcMon.Visible = False
    fraProcMon.Enabled = False
    fraLocal.Visible = False
    fraLocal.Enabled = False
    fraScan.Visible = False
    fraScan.Enabled = False
    fraAdvScan.Visible = False
    fraAdvScan.Enabled = False
    fraSweep.Enabled = False
    fraSweep.Visible = False
    fraMsg.Enabled = False
    fraMsg.Visible = False
    fraTCP.Visible = False
    fraTCP.Enabled = False

  ElseIf MenuItem = 2 Then
    fraScan.Enabled = True
    fraScan.Visible = True
    status.Panels(1).Text = "Scans for open ports on local or remote computer"
    Me.Caption = App.FileDescription & " - Ports scanner"
    fraPing.Visible = False
    fraPing.Enabled = False
    fraHW.Visible = False
    fraHW.Enabled = False
    fraProcMon.Visible = False
    fraProcMon.Enabled = False
    fraLocal.Visible = False
    fraLocal.Enabled = False
    fraAdvScan.Visible = False
    fraAdvScan.Enabled = False
    fraSweep.Enabled = False
    fraSweep.Visible = False
    fraMsg.Enabled = False
    fraMsg.Visible = False
    fraTCP.Visible = False
    fraTCP.Enabled = False

  ElseIf MenuItem = 3 Then
    status.Panels(1).Text = "Scans range of ports on specified IP"
    Me.Caption = App.FileDescription & " - Advanced port scanner"
    fraAdvScan.Visible = True
    fraAdvScan.Enabled = True
    fraScan.Enabled = False
    fraScan.Visible = False
    fraPing.Visible = False
    fraPing.Enabled = False
    fraHW.Visible = False
    fraHW.Enabled = False
    fraProcMon.Visible = False
    fraProcMon.Enabled = False
    fraLocal.Visible = False
    fraLocal.Enabled = False
    fraSweep.Enabled = False
    fraSweep.Visible = False
    fraMsg.Enabled = False
    fraMsg.Visible = False
    fraTCP.Visible = False
    fraTCP.Enabled = False

  ElseIf MenuItem = 4 Then
    status.Panels(1).Text = "Scans given range of IP addresses for specified open port"
    Me.Caption = App.FileDescription & " - Port sweeper"
    fraSweep.Enabled = True
    fraSweep.Visible = True
    fraScan.Enabled = False
    fraScan.Visible = False
    fraPing.Visible = False
    fraPing.Enabled = False
    fraHW.Visible = False
    fraHW.Enabled = False
    fraProcMon.Visible = False
    fraProcMon.Enabled = False
    fraLocal.Visible = False
    fraLocal.Enabled = False
    fraAdvScan.Visible = False
    fraAdvScan.Enabled = False
    fraMsg.Enabled = False
    fraMsg.Visible = False
    fraTCP.Visible = False
    fraTCP.Enabled = False

  ElseIf MenuItem = 5 Then
    fraMsg.Enabled = True
    fraMsg.Visible = True
    status.Panels(1).Text = "Do varoius action using Net Send command easily"
    frmMain.Caption = App.FileDescription & " - Messenger exploit"
    fraSweep.Enabled = False
    fraSweep.Visible = False
    fraScan.Enabled = False
    fraScan.Visible = False
    fraPing.Visible = False
    fraPing.Enabled = False
    fraHW.Visible = False
    fraHW.Enabled = False
    fraProcMon.Visible = False
    fraProcMon.Enabled = False
    fraLocal.Visible = False
    fraLocal.Enabled = False
    fraAdvScan.Visible = False
    fraAdvScan.Enabled = False
    fraTCP.Visible = False
    fraTCP.Enabled = False

  ElseIf MenuItem = 6 Then
    fraTCP.Visible = True
    fraTCP.Enabled = True
    status.Panels(1).Text = "Watch and monitor TCP connections on your computer"
    frmMain.Caption = App.FileDescription & " - TCP monitor"
    fraMsg.Enabled = False
    fraMsg.Visible = False
    fraSweep.Enabled = False
    fraSweep.Visible = False
    fraScan.Enabled = False
    fraScan.Visible = False
    fraPing.Visible = False
    fraPing.Enabled = False
    fraHW.Visible = False
    fraHW.Enabled = False
    fraProcMon.Visible = False
    fraProcMon.Enabled = False
    fraLocal.Visible = False
    fraLocal.Enabled = False
    fraAdvScan.Visible = False
    fraAdvScan.Enabled = False

  End If

ElseIf MenuNumber = 3 Then
  If MenuItem = 1 Then
    frmAbout.Show vbModal
  
  ElseIf MenuItem = 2 Then
    End
    
End If
End If
End Sub

Private Sub P1_Change()
If P1.Text = "0" Then
MsgBox "Can't uses 0 (zero)", vbExclamation, "NetMatrix"
P1.Text = ""
End If
End Sub

Private Sub P2_Change()
If P2.Text = "0" Then
MsgBox "Can't uses 0 (zero)", vbExclamation, "NetMatrix"
P2.Text = ""
End If
End Sub

Private Sub timAnim_Timer()
If status.Panels(5).Text = "NetMatrix" Then
status.Panels(5).Text = "Version 1.7"
ElseIf status.Panels(5).Text = "Version 1.7" Then
status.Panels(5).Text = "Copyright"
ElseIf status.Panels(5).Text = "Copyright" Then
status.Panels(5).Text = "July 5th, 2005"
ElseIf status.Panels(5).Text = "July 5th, 2005" Then
status.Panels(5).Text = "by"
ElseIf status.Panels(5).Text = "by" Then
status.Panels(5).Text = "NeverHard!"
ElseIf status.Panels(5).Text = "NeverHard!" Then
status.Panels(5).Text = "NetMatrix"
End If
End Sub

Private Sub timAutoR_Timer()
On Error GoTo doctor
    '
    Dim arrBuffer() As Byte
    Dim lngSize As Long
    Dim lngRetVal As Long
    Dim lngRows As Long
    Dim i As Long
    Dim TcpTableRow As MIB_TCPROW
    Dim lvItem As ListItem
    '
    ListView1.ListItems.Clear
    '
    lngSize = 0
    '
    'Call the GetTcpTable just to get the buffer size into the lngSize variable
    lngRetVal = GetTcpTable(ByVal 0&, lngSize, 0)
    '
    If lngRetVal = ERROR_NOT_SUPPORTED Then
        '
        'This API works only on Win 98//2000 and NT4 with SP4
        msg = MsgBox("IP Helper is not supported by this system.", vbCritical, "NetMatrix")
        timAutoR.Enabled = False
        Exit Sub
        '
    End If
    '
    'Prepare the buffer
    ReDim arrBuffer(0 To lngSize - 1) As Byte
    '
    'And call the function one more time
    lngRetVal = GetTcpTable(arrBuffer(0), lngSize, 0)
    '
    If lngRetVal = ERROR_SUCCESS Then
        '
        'The first 4 bytes contain the quantity of the table rows
        'Get that value to the lngRows variable
        CopyMemory lngRows, arrBuffer(0), 4
        '
        For i = 1 To lngRows
            '
            'Copy the table row data to the TcpTableRow structure
            CopyMemory TcpTableRow, arrBuffer(4 + (i - 1) * Len(TcpTableRow)), Len(TcpTableRow)
            '
            If Not ((chkHideIP.Value = vbChecked) And (GetIpFromLong(TcpTableRow.dwLocalAddr) = "0.0.0.0" Or GetIpFromLong(TcpTableRow.dwLocalAddr) = "127.0.0.1")) Then
                '
                'Add the data to the ListView control
                With TcpTableRow
                    Set lvItem = ListView1.ListItems.Add(, , GetIpFromLong(.dwLocalAddr))
                    lvItem.SubItems(1) = GetTcpPortNumber(.dwLocalPort)
                    lvItem.SubItems(2) = GetIpFromLong(.dwRemoteAddr)
                    lvItem.SubItems(3) = GetTcpPortNumber(.dwRemotePort)
                    lvItem.SubItems(4) = GetState(.dwState)
                End With
                '
            End If
            '
        Next i
        '
    End If
    '
Exit Sub
doctor:
chkAuto.Value = vbUnchecked
Exit Sub
End Sub

Private Sub timDateTime_Timer()
 timDateTime.Enabled = False
On Error Resume Next
status.Panels(2).Text = Date & " " & Time
End Sub

Private Sub timerPercent_Timer()
On Error GoTo doctor
Dim a, b As String
Dim c As Byte
a = lblScannedPorts.Caption
b = txtLimit.Text
c = (lblScannedPorts.Caption / txtLimit.Text) * 100
lblPercent.Caption = c & " %"
bar.Value = c
Exit Sub
doctor:
msg = MsgBox("An error occured, please make sure that you fill valid informations.", vbCritical, "NetMatrix")
TimerScan.Enabled = False
timerPercent.Enabled = False
Picture2.Enabled = True
txtLimit.Enabled = True
txtTimeout.Enabled = True
chkBeep.Enabled = True
cmdScan.Caption = "&Scan"
Exit Sub
End Sub

Private Sub TimerStatus_Timer()
'Connectio status
   LoadNewWinsock
   On Error Resume Next
   wsknetworkstatus(1).SendData "TEST"
   If Err <> 0 Or wsknetworkstatus(1).LocalIP = "127.0.0.1" Then
      status.Panels(4).Text = "No connection"
   Else
      status.Panels(4).Text = "Connection present"
   End If
End Sub

Private Sub TimerSysMon_Timer()
Dim theloop, ret As String
Dim i As Long
Dim proc As PROCESSENTRY32
Dim snap As Long
Dim exename As String
lvw.ListItems.Clear 'clear listview contents
snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0) 'get snapshot handle
proc.dwSize = Len(proc)
theloop = ProcessFirst(snap, proc)       'first process and return value
i = 0
While theloop <> 0      'next process
exename = proc.szExeFile
ret = lvw.ListItems.Add(, "first" & CStr(i), exename)   'add process name to listview
lvw.ListItems("first" & CStr(i)).SubItems(1) = proc.th32ProcessID   'add process ID to listview
i = i + 1
theloop = ProcessNext(snap, proc)
Wend
CloseHandle snap       'close snapshot handle
End Sub

Private Sub TimerTotalIP_Timer()
'Total IP
lblTotalIP.Caption = "IP loaded: " & List1.ListCount
End Sub

Private Sub TimProgress_Timer()
lblAdvPS.Caption = "Now scanning on port: " & Str(RSContinue)
lblIPPS.Caption = "At IP: " & IP.Text
lblTotalPS.Caption = "Total openport found: " & OP.ListCount
Dim valA, valB, valC As Byte
valA = Str(RSContinue)
valB = Val(P2) - Val(P1)
valC = (valA / valB) * 100
lblTPTS.Caption = "Total port to scan: " & valB
On Error Resume Next
PB.Value = valC
lblPerc.Caption = PB.Value & " %"
If valA = P2.Text Then PB.Value = PB.Max
If Str(RSContinue) = P2.Text Then
SN.Enabled = True
SS.Enabled = False
P1.Enabled = True
IP.Enabled = True
P2.Enabled = True
Mc.Enabled = True
PB.Value = PB.Max
TimProgress.Enabled = False
End If
End Sub

Private Sub TimPS_Timer()
On Error GoTo doctor

txtPSIP.Text = txtToIP1.Text & "." & IPToScan

If TimeOut = 1 Then
 If SockPS.State = sckConnected Then
 SockPS.Close
 lstPS.AddItem txtToIP1.Text & "." & IPToScan
 'Beep
 
 TimeOut = 0
 IPToScan = IPToScan + 1
 Else
 SockPS.Close
 TimeOut = 0
 IPToScan = IPToScan + 1
 End If
End If

If TimeOut <> 1 Then
 If IPToScan <= IPToStopOn Then
 SockPS.RemoteHost = txtToIP1.Text & "." & IPToScan
 SockPS.Connect
 TimeOut = 1
 
Else
 txtFromIP.Enabled = True
 txtToIP1.Enabled = True
 txtToIP2.Enabled = True
 txtPSPort.Enabled = True
 txtPSInterval.Enabled = True
 cmdStartScan.Enabled = True
 cmdResetPS.Enabled = True
 txtPSIP.Text = "Done!"
 msg = MsgBox("Sweeping completed succesfully." & vbCrLf & lstPS.ListCount & " servers found!" & vbCrLf & "At port: " & txtPSPort.Text, vbInformation, "NetMatrix")
 cmdPauseScan.Enabled = False
 TimPS.Enabled = False
 End If
End If

Exit Sub
doctor:
Exit Sub
End Sub

Private Sub txtBombMsg_Change()
txtBombLength.Text = Len(txtBombMsg.Text) & " character of 175 allowed"
End Sub

Private Sub txtBombQuantity_Change()
If txtBombQuantity.Text = "0" Then
msg = MsgBox("Can't uses 0 (zero).", vbExclamation, "NetMatrix")
txtBombQuantity.Text = ""
Else
Exit Sub
End If
End Sub

Private Sub txtFromIP_Change()
If TimPS.Enabled = False Then
Comparison = txtFromIP.Text Like "*.*.*.*"
 If Comparison = False Then
 txtToIP1.Text = txtFromIP.Text
 End If
Comparison2 = txtToIP1.Text Like "*.*.*[!.]"
Comparison3 = txtToIP1.Text Like "*.*.*.*"
 Do While Comparison3 = True
  Do Until Comparison2 = True
  txtToIP1.SelStart = Len(txtToIP1.Text) - 1
  txtToIP1.SelLength = 1
  txtToIP1.SelText = ""
  Comparison2 = txtToIP1.Text Like "*.*.*[!.]"
  Loop
 Loop
End If
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdScan_Click
End Sub

Private Sub Form_Load()

'Menu
    fraLocal.Visible = True
    fraLocal.Enabled = True
    status.Panels(1).Text = "Show local IP, local host name, and local computer name"
    Me.Caption = App.FileDescription & " - Local info"
    fraProcMon.Visible = False
    fraProcMon.Enabled = False
    fraHW.Visible = False
    fraHW.Enabled = False
    fraPing.Visible = False
    fraPing.Enabled = False
    fraScan.Visible = False
    fraScan.Enabled = False
    fraAdvScan.Visible = False
    fraAdvScan.Enabled = False
    fraSweep.Enabled = False
    fraSweep.Visible = False
    fraMsg.Enabled = False
    fraMsg.Visible = False

'Max connection (Advanced Port Scanner)
Dim cval As Byte
For cval = 1 To 93
Mc.AddItem cval
Next

Mc.ListIndex = Mc.ListCount - Mc.ListCount


'Hardware scanner
    eilute = 0
    webUserName = ""
    webPassword = ""
    isClient = False
    isClienta = False
    klientoID = 0
    ramas = Array("Unknown", "Other", "SIP", "DIP", "ZIP", "SOJ", "Proprietary", _
                    "SIMM", "DIMM", "TSOP", "PGA", "RIMM", "SODIMM")
                
    ramotipas = Array("Unknown", "Other", "DRAM", "Synchronous DRAM", "Cache DRAM", _
                    "EDO", "EDRAM", "VRAM", "SRAM", "RAM", "ROM", "Flash", "EEPROM", _
                    "FEPROM", "EPROM", "CDRAM", "3DRAM", "SDRAM", "SGRAM")
                    
ReDim Preserve PelesInt(165)
PelesInt(1) = "Other"
PelesInt(2) = "Unknown"
PelesInt(3) = "Serial"
PelesInt(4) = "PS / 2"
PelesInt(5) = "Infrared"
PelesInt(6) = "HP - HIL"
PelesInt(7) = "Bus mouse"
PelesInt(8) = "ADB (Apple Desktop Bus)"
PelesInt(160) = "Bus mouse DB-9"
PelesInt(161) = "Bus mouse micro-DIN"
PelesInt(162) = "USB"

ReDim Preserve PelesTipas(10)
PelesTipas(1) = "Other"
PelesTipas(2) = "Unknown"
PelesTipas(3) = "Mouse"
PelesTipas(4) = "Track Ball"
PelesTipas(5) = "Track Point"
PelesTipas(6) = "Glide Point"
PelesTipas(7) = "Touch Pad"

Set objWshNet = CreateObject("Wscript.Network")
txtUname = objWshNet.ComputerName
'Call Scan_Click

'==================================================

'Net Sender character length
txtNetSendLength.Text = Len(txtNetSendMsg.Text) & " character of 125 allowed"
txtBombLength.Text = Len(txtBombMsg.Text) & " character of 125 allowed"

'System process monitor
Dim header As ColumnHeader
lvw.View = lvwReport
lvw.ColumnHeaders.Clear
Set header = lvw.ColumnHeaders.Add(, "first", "Process", 4200)  'set listview width
Set header = lvw.ColumnHeaders.Add(, "second", "ID", 1200)
lvw.Refresh

chkAuto.Value = vbChecked

'Connection status
   LoadNewWinsock
   On Error Resume Next
   wsknetworkstatus(1).SendData "TEST"
   If Err <> 0 Or wsknetworkstatus(1).LocalIP = "127.0.0.1" Then
      status.Panels(4).Text = "No connection   "
   Else
      status.Panels(4).Text = "Connection present   "
   End If

'Ports scanner
port = 0
txtIP = ws.LocalIP

'Local IP
    txtLocalHost.Text = ws.LocalHostName
    txtLocalIP.Text = ws.LocalIP
    
'IP(s) ping-er
On Error GoTo doctor
Dim s As String
s = App.Path & "\Ping.inf"
Dim a As String
Dim X As String
On Error GoTo doctor
Open s For Input As #1
Do Until EOF(1)
Input #1, a$
List1.AddItem a$
Loop
Close 1
s = ""


Call cmdStartScan_Click


Exit Sub

'Total IP
lblTotalIP.Caption = "Total IP: " & List1.ListCount

Exit Sub
doctor:
Dim msg As Byte
msg = MsgBox("File not found (Ping.inf). Ignoring file error, continuing application...", vbCritical, "NetMatrix - Error")
Exit Sub
End Sub

'IP(s) ping-er codes
Private Sub cmdStart_Click()
On Error GoTo doctor
Dim Cuts As Double
Prog1.Value = 0
Text1.Text = ""
Cuts = 100 / List1.ListCount
If List1.Text = "" Then
Else
    Call PrintInTextBox(List1.Text, ECHO)
    Prog1.Value = 100
    Exit Sub
End If

For i = 0 To List1.ListCount - 1
    Call PrintInTextBox(List1.List(i), ECHO)
    Prog1.Value = Cuts * (i + 1)
Next
Exit Sub
doctor:
Dim msg As Byte
msg = MsgBox("Internal error occured. No IP address tobe ping-ed", vbCritical, "NetMatrix - Error")
Exit Sub
End Sub

Private Sub cmdAdd_Click()
Dim StrIP As String
If ip1 = "" Or ip2 = "" Or ip3 = "" Or ip4 = "" Then
    MsgBox "IP Address must be 4 groups of 0 to 255 range of values.", vbExclamation, "NetMatrix"
    ip1.SetFocus
    Exit Sub
End If
If Val(ip1) > 255 Or Val(ip2) > 255 Or Val(ip3) > 255 Or Val(ip4) > 255 Then
    MsgBox "IP Address must be 4 groups of 0 to 255 range of values.", vbExclamation, "NetMatrix"
    ip1.SetFocus
    Exit Sub
Else
    StrIP = Trim(ip1) & "." & Trim(ip2) & "." & Trim(ip3) & "." & Trim(ip4)
    If List1.ListCount = 0 Then
        List1.AddItem StrIP
        Exit Sub
    End If
    If StrIP = "" Then
        MsgBox "Enter IP address in the textbox", vbExclamation, "NetMatrix"
        Exit Sub
    End If
    List1.Refresh
    For a = 0 To List1.ListCount - 1
        If List1.List(a) = StrIP Then
            MsgBox "IP Address already added to the List.", vbExclamation, "NetMatrix"
            Exit Sub
        Else
            List1.AddItem StrIP
            ip1.Text = ""
            ip2.Text = ""
            ip3.Text = ""
            ip4.Text = ""
            ip1.SetFocus
            Exit Sub
        End If
    Next
End If

'Total IP
lblTotalIP.Caption = "Total IP: " & List1.ListCount
End Sub

Private Sub cmdClear_Click()
Text1.Text = ""
End Sub

Private Sub cmdSave_Click()
On Error GoTo doctor
Dim sFile As String
sFile = App.Path & "\Ping.inf"
Open sFile For Output As #1
For i = 0 To 100
    List1.ListIndex = i
    Print #1, List1.Text
Next i
Exit Sub
doctor:
'msg = MsgBox("File not found (Ping.inf). Ignoring file error, continuing application...", vbCritical, "NetMatrix - Error")
Exit Sub
End Sub

Private Sub cmdRemove_Click()
If List1.SelCount = 0 Then
    MsgBox "Select an IP Address from list first", vbExclamation, "NetMatrix"
    Exit Sub
End If
List1.RemoveItem List1.ListIndex

'Total IP
lblTotalIP.Caption = "Total IP: " & List1.ListCount
End Sub

Private Sub PrintInTextBox(IPAddr As String, ECHO1 As ICMP_ECHO_REPLY)
        Call Ping(IPAddr, ECHO1)
    '"  " could be changed vbTab
        Text1.Text = Text1.Text & IPAddr & " : " & vbCrLf
        Text1.Text = Text1.Text & "  " & " Status : " & GetStatusCode(ECHO.status) & vbCrLf
        Text1.Text = Text1.Text & "  " & " Address : " & ECHO.Address & vbCrLf
        Text1.Text = Text1.Text & "  " & " DataSize : " & ECHO.DataSize & vbCrLf
        Text1.Text = Text1.Text & "  " & " Round Trip Time : " & ECHO.RoundTripTime & vbCrLf
End Sub

Private Sub ip1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or Len(ip1) = 2 Or KeyAscii = 32 Or KeyAscii = 46 Then ip2.SetFocus
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub ip2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or Len(ip2) = 2 Or KeyAscii = 32 Or KeyAscii = 46 Then ip3.SetFocus
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub ip3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or Len(ip3) = 2 Or KeyAscii = 32 Or KeyAscii = 46 Then ip4.SetFocus
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub ip4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or Len(ip4) = 2 Then cmdAdd.SetFocus
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 32 Or KeyAscii = 46 Then
Else
    KeyAscii = 0
End If
End Sub

'PC users
Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Long
    ZeroPos = InStr(1, sInput, Chr$(0))
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function

Private Sub TimerScan_Timer()
On Error Resume Next
If lblScannedPorts.Caption = Val(txtLimit.Text) - 1 Then
TimerScan.Enabled = False
bar.Value = bar.Max
lblPercent.Caption = bar.Value & " %"
msg = MsgBox("Scanning completed successfully." & vbCrLf & vbCrLf & "Scanned IP address: " & txtIP.Text & vbCrLf & "Total ports scanned: " & txtLimit.Text & vbCrLf & "Open ports total: " & lblOpenPorts.Caption, vbInformation, "NetMatrix")
bar.Value = bar.Max
lblPercent.Caption = bar.Value & " %"
timerPercent.Enabled = False
Picture2.Enabled = True
txtLimit.Enabled = True
txtTimeout.Enabled = True
chkBeep.Enabled = True
cmdScan.Caption = "&Scan"
End If
lblScannedPorts = port
If portnumber = Not txtLimit.Text Then
TimerScan.Enabled = False
timerPercent.Enabled = False
cmdScan.Caption = "&Scan"
Picture2.Enabled = False
txtLimit.Enabled = False
txtTimeout.Enabled = False
chkBeep.Enabled = False
Exit Sub
ElseIf ws.State = 7 Then
lstPort.AddItem ws.RemotePort
lblOpenPorts = lblOpenPorts + 1
If chkBeep = 1 Then Beep
ws.Close: port = port + 1
ws.Connect txtIP.Text, port
Exit Sub
ElseIf Not ws.State Then ws.Close
ws.Connect txtIP.Text, port
port = port + 1
End If
End Sub

Private Sub LoadNewWinsock()
   On Error Resume Next
   
   If wsknetworkstatus.Count > 1 Then
      If wsknetworkstatus(1).LocalIP <> "127.0.0.1" Then
         Exit Sub
      End If
   End If
   
   UnloadWinsock
   
   Load wsknetworkstatus(1)
   wsknetworkstatus(1).RemoteHost = wsknetworkstatus(1).LocalIP
   wsknetworkstatus(1).RemotePort = 5555
   wsknetworkstatus(1).LocalPort = 5555
   wsknetworkstatus(1).Protocol = sckUDPProtocol
   wsknetworkstatus(1).Bind
End Sub

Private Sub UnloadWinsock()
   If wsknetworkstatus.Count > 1 Then
      Unload wsknetworkstatus(1)
   End If
End Sub

Private Sub txtLimit_Change()
If txtLimit.Text = "0" Then
msg = MsgBox("Can't uses 0 (zero).", vbExclamation, "NetMatrix")
txtLimit.Text = ""
Else
Exit Sub
End If
End Sub

Private Sub txtNetSendMsg_Change()
txtNetSendLength.Text = Len(txtNetSendMsg.Text) & " character of 175 allowed"
End Sub

Private Sub txtPSInterval_Change()
If txtPSInterval.Text <> "" Then
TimPS.interval = txtPSInterval.Text
End If
End Sub

Private Sub txtPSInterval_Click()
txtPSInterval.SelStart = 0
txtPSInterval.SelLength = Len(txtPSInterval.Text)
End Sub

Private Sub txtTimeout_Change()
If txtTimeout.Text = "0" Then
msg = MsgBox("Can't uses 0 (zero).", vbExclamation, "NetMatrix")
txtTimeout.Text = ""
Else
Exit Sub
End If
End Sub

'Hardware scanner
'================
Private Sub GetSndDevInfo(objService, strWBEMClass)
  
    On Error Resume Next
        
    ReDim Preserve oDeviceType(100)
    ReDim Preserve oDeviceCaption(100)
    ReDim Preserve oDeviceParam(100)
    ReDim Preserve oDeviceInterf(100)
    
    Set objDeviceSet = objService.InstancesOf(strWBEMClass)
    'MsgBox strWBEMClass
    If objDeviceSet.Count <> 0 Then
        For Each Device In objDeviceSet
        
    Select Case strWBEMClass
' GARSAS----------------------------------
        Case "Win32_SoundDevice"
            txtHW.Text = txtHW.Text & "Sound device:" & vbCrLf & Device.Caption & vbCrLf & vbCrLf
            oDeviceType(eilute) = "Sound device"
            oDeviceCaption(eilute) = Device.Caption
            oDeviceParam(eilute) = ""
            oDeviceInterf(eilute) = ""
            eilute = eilute + 1
' VIDEO-----------------------------------
        Case "Win32_VideoController"
            txtHW.Text = txtHW.Text & "Video controller:" & vbCrLf & Device.Caption & vbCrLf & "Memory: " & Device.AdapterRAM / 1048576 & " MB" & vbCrLf & vbCrLf
            oDeviceType(eilute) = "Video controller"
            oDeviceCaption(eilute) = Device.Caption
            oDeviceParam(eilute) = Device.AdapterRAM / 1048576
            oDeviceInterf(eilute) = ""
            eilute = eilute + 1
' NETWORK----------------------------------
        Case "Win32_NetworkAdapter"
            If (Device.NetConnectionID = "Local Area Connection") And (Device.MACAddress <> "") Then
                txtHW.Text = txtHW.Text & "Network adapter:" & vbCrLf & Device.Caption & vbCrLf & Device.MACAddress & vbCrLf
                oDeviceType(eilute) = "Network adapter"
                oDeviceCaption(eilute) = Device.Caption
                oDeviceParam(eilute) = Device.MACAddress
                oDeviceInterf(eilute) = ""
                eilute = eilute + 1
            End If
' KEYBOARD---------------------------------
        Case "Win32_Keyboard"
            txtHW.Text = txtHW.Text & "Keyboard:" & vbCrLf & Device.Description & vbCrLf & vbCrLf
            oDeviceType(eilute) = "Keyboard"
            oDeviceCaption(eilute) = Device.Description
            oDeviceParam(eilute) = ""
            oDeviceInterf(eilute) = ""
            eilute = eilute + 1
' MOUSE---------------------------------
        Case "Win32_PointingDevice"
            txtHW.Text = txtHW.Text & "Pointing device/mouse:" & vbCrLf & Device.Caption & vbCrLf & PelesTipas(Device.PointingType) & vbCrLf & "Interface: " & PelesInt(Device.DeviceInterface) & vbCrLf & vbCrLf
            oDeviceType(eilute) = "Pointing device/mouse:"
            oDeviceCaption(eilute) = Device.Caption
            oDeviceParam(eilute) = PelesTipas(Device.PointingType)
            oDeviceInterf(eilute) = PelesInt(Device.DeviceInterface)
            eilute = eilute + 1
' DISK----------------------------------
        Case "Win32_DiskDrive"
            txtHW.Text = txtHW.Text & "Harddisk:" & vbCrLf & Device.Description & vbCrLf & Device.Caption & vbCrLf & "Capacity: " & Device.Size / 1000000000 & " GB" & vbCrLf & "Interface: " & Device.InterfaceType & vbCrLf & vbCrLf
            oDeviceType(eilute) = Device.Description
            oDeviceCaption(eilute) = Device.Caption
            oDeviceParam(eilute) = Device.Size
            oDeviceInterf(eilute) = Device.InterfaceType
            eilute = eilute + 1
' CD-ROM--------------------------------------
        Case "Win32_CDROMDrive"
            txtHW.Text = txtHW.Text & "CD ROM:" & vbCrLf & Device.Description & vbCrLf & "Drive letter: " & Device.Caption & vbCrLf & "Capacity: " & Device.Size / 1048576 & " MB" & vbCrLf & vbCrLf
            oDeviceType(eilute) = Device.Description
            oDeviceCaption(eilute) = Device.Caption
            oDeviceParam(eilute) = Device.Size
            oDeviceInterf(eilute) = ""
            eilute = eilute + 1
' SCSI------------------------------------------
        Case "Win32_SCSIController"
            txtHW.Text = txtHW.Text & "SCSI controller:" & vbCrLf & Device.Caption & vbCrLf & vbCrLf
            oDeviceType(eilute) = "SCSI Controller"
            oDeviceCaption(eilute) = Device.Caption
            oDeviceParam(eilute) = ""
            oDeviceInterf(eilute) = ""
            eilute = eilute + 1
' PROCESSOR-------------------------------------
        Case "Win32_Processor"
            txtHW.Text = txtHW.Text & "Processor/" & Device.Role & ":" & vbCrLf & Device.Name & vbCrLf & "Clock speed: " & Device.CurrentClockSpeed & vbCrLf & vbCrLf
            oDeviceType(eilute) = Device.Role
            oDeviceCaption(eilute) = Device.Name
            oDeviceParam(eilute) = Device.CurrentClockSpeed
            oDeviceInterf(eilute) = ""
            eilute = eilute + 1
' MEMORY-----------------------------------------
        Case "Win32_PhysicalMemory"
            txtHW.Text = txtHW.Text & Device.Description & ":" & vbCrLf & "Formfactor: " & ramas(Device.FormFactor) & vbCrLf & "Capacity: " & Device.Capacity / 1048576 & " MB" & vbCrLf & "Memory type: " & ramotipas(Device.MemoryType) & vbCrLf & vbCrLf
            oDeviceType(eilute) = Device.Description
            oDeviceCaption(eilute) = ramas(Device.FormFactor)
            oDeviceParam(eilute) = Device.Capacity / 1048576
            oDeviceInterf(eilute) = ramotipas(Device.MemoryType)
            eilute = eilute + 1
' FLOPYY--------------------------------------
        Case "Win32_FloppyDrive"
            txtHW.Text = txtHW.Text & Device.Description & vbCrLf & Device.Caption & vbCrLf & Device.MaxMediaSize & vbCrLf & vbCrLf
            oDeviceType(eilute) = Device.Description
            oDeviceCaption(eilute) = Device.Caption
            oDeviceParam(eilute) = Device.MaxMediaSize
            oDeviceInterf(eilute) = ""
            eilute = eilute + 1
' MODEM------------------------------------
        Case "Win32_POTSModem"
            txtHW.Text = txtHW.Text & "POTS modem:" & vbCrLf & Device.Caption & vbCrLf & "Max baud rate to phone: " & Device.MaxBaudRateToPhone & vbCrLf & Device.Description & vbCrLf & vbCrLf
            oDeviceType(eilute) = "POTS Modem"
            oDeviceCaption(eilute) = Device.Caption
            oDeviceParam(eilute) = Device.MaxBaudRateToPhone
            oDeviceInterf(eilute) = Device.Description
            eilute = eilute + 1
' INFRARED----------------------------------
        Case "Win32_InfraredDevice"
            txtHW.Text = txtHW.Text & "Infrared device:" & vbCrLf & Device.Caption & vbCrLf & vbCrLf
            oDeviceType(eilute) = "Infrared Device"
            oDeviceCaption(eilute) = Device.Caption
            oDeviceParam(eilute) = ""
            oDeviceInterf(eilute) = ""
            eilute = eilute + 1
' PCMCIA ----------------------------------
        Case "Win32_PCMCIAController"
            txtHW.Text = txtHW.Text & "PCMCIA controller:" & vbCrLf & Device.Caption & vbCrLf & vbCrLf
            oDeviceType(eilute) = "PCMCIA Controller"
            oDeviceCaption(eilute) = Device.Caption
            oDeviceParam(eilute) = ""
            oDeviceInterf(eilute) = ""
            eilute = eilute + 1
' TAPE -------------------------------------
        Case "Win32_TapeDrive"
            txtHW.Text = txtHW.Text & "Tape drive:" & vbCrLf & Device.Caption & vbCrLf & "Capacity: " & Device.MaxMediaSize & vbCrLf & Device.Description & vbCrLf & vbCrLf
            oDeviceType(eilute) = "Tape Drive"
            oDeviceCaption(eilute) = Device.Caption
            oDeviceParam(eilute) = Device.MaxMediaSize
            oDeviceInterf(eilute) = Device.Description
            eilute = eilute + 1
' BATTERY-----------------------------------
        Case "Win32_PortableBattery"
            txtHW.Text = txtHW.Text & "Portable battery:" & vbCrLf & Device.Caption & vbCrLf & Device.Chemistry & vbCrLf & vbCrLf
            oDeviceType(eilute) = "Portable Battery"
            oDeviceCaption(eilute) = Device.Caption
            oDeviceParam(eilute) = ""
            oDeviceInterf(eilute) = Device.Chemistry
            eilute = eilute + 1
    End Select
    Next
    End If

txtHW.Text = txtHW.Text & "========================================" & vbCrLf & vbCrLf
End Sub

Private Function ConnectTO(ByVal strNameSpace, _
                            ByVal strUserName, _
                            ByVal strPassword, _
                            ByRef strServer, _
                            ByRef objService)

    On Error Resume Next

    Dim objLocator, objWshNet

    ConnectTO = True     'There is no error.

    'Create Locator object to connect to remote CIM object manager
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    If Err.Number Then
        MsgBox "Error 0x" & CStr(Hex(Err.Number)) & _
                           " occurred in creating a locator object.", vbCritical, "NetMatrix"
        If Err.Description <> "" Then
            MsgBox "Error description: " & Err.Description & ".", vbCritical, "NetMatrix"
        End If
        Err.Clear
        ConnectTO = False     'An error occurred
        Exit Function
    End If

    'Connect to the namespace which is either local or remote
    Set objService = objLocator.ConnectServer(strServer, strNameSpace, _
       strUserName, strPassword)
    objService.Security_.impersonationlevel = 3
    If Err.Number Then
        MsgBox "Error 0x" & CStr(Hex(Err.Number)) & _
                           " occurred in connecting to server " _
           & strServer & ".", vbCritical, "NetMatrix"
        If Err.Description <> "" Then
            MsgBox "Error description: " & Err.Description & ".", vbCritical, "NetMatrix"
        End If
        Err.Clear
        ConnectTO = False     'An error occurred
    End If
End Function

Public Sub scanhard()
On Error GoTo doctor
Dim DeviceListLen As String
    frmPrompt.Show
    eilute = 0
    'MSFlexGrid1
    ReDim Preserve DeviceList(40)
    ReDim Preserve DeviceFound(40)
    DeviceListLen = 16
    DeviceList = Array("Win32_FloppyDrive", "Win32_DiskDrive", "Win32_CDROMDrive", _
                "Win32_Processor", _
                "Win32_PhysicalMemory", _
                "Win32_SoundDevice", "Win32_SCSIController", "Win32_VideoController", _
                "Win32_Keyboard", _
                "Win32_PointingDevice", _
                "Win32_NetworkAdapter", "Win32_POTSModem", _
                "Win32_InfraredDevice", _
                "Win32_PCMCIAController", _
                "Win32_TapeDrive", _
                "Win32_PortableBattery")

  
    strServer = txtUname
    isconnect = ConnectTO("root\cimv2", _
                   strUserName, _
                   strPassword, _
                   strServer, _
                   objService)
    If Not isconnect Then
        MsgBox "Please check the server name, " _
                        & "credentials and WBEM Core.", vbCritical, "NetMatrix"
    
    End If
    DeviCecount = 0
    For i = 0 To DeviceListLen - 1
        Set objDeviceSet = objService.InstancesOf(DeviceList(i))
        If objDeviceSet.Count <> 0 Then
            DeviceFound(DeviCecount) = DeviceList(i)
            DeviCecount = DeviCecount + 1
            Call GetSndDevInfo(objService, DeviceList(i))
            'MsgBox MSFlexGrid1.Rows
        End If
    Next
    frmPrompt.Hide
Exit Sub
doctor:
MsgBox "Failed to access hardware information. Operation terminated.", vbCritical, "NetMatrix"
frmPrompt.Hide
End Sub

'Advanced port scanner
'=====================
Private Sub CL_Click()
   Me.OP.Clear
End Sub

Private Sub SN_Click()
SN.Enabled = False
SS.Enabled = True
P1.Enabled = False
IP.Enabled = False
P2.Enabled = False
Mc.Enabled = False
TimProgress.Enabled = True
   Dim RS As Integer
   RSContinue = Val(Me.P1)
   For RS = 1 To Val(Me.Mc)
      Load Me.S1(RS)
      RSContinue = RSContinue + 1
      Me.S1(RS).Connect Me.IP, RSContinue
   Next RS
If Str(RSContinue) = P2.Text Then
SN.Enabled = True
SS.Enabled = False
P1.Enabled = True
IP.Enabled = True
P2.Enabled = True
Mc.Enabled = True
PB.Value = PB.Max
TimProgress.Enabled = False
End If
End Sub

Private Sub SS_Click()
SN.Enabled = True
SS.Enabled = False
P1.Enabled = True
IP.Enabled = True
P2.Enabled = True
Mc.Enabled = True
TimProgress.Enabled = False
   On Error GoTo Err:
   Dim RS As Integer
   For RS = 1 To Val(Me.Mc)
      Me.S1(RS).Close
      Unload Me.S1(RS)
   Next RS
Err:
   Exit Sub
End Sub

Private Sub S1_Connect(Index As Integer)
   Me.OP.AddItem "Found: port " + Str(Me.S1(Index).RemotePort)
   KeepGoing (Index)
End Sub

Private Sub S1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   KeepGoing (Index)
End Sub

