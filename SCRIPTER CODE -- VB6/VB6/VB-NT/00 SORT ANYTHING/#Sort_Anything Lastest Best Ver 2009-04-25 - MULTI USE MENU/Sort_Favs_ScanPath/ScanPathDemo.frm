VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ScanPath 
   BackColor       =   &H00808080&
   Caption         =   "ScanPath 2.0 - Sort  Anything -"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   15240
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   7095
      Top             =   3255
   End
   Begin VB.CommandButton Command17 
      Caption         =   "MOUSE MOMENT PAUSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   4020
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   1935
      Width           =   1155
   End
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   240
      Left            =   2088
      TabIndex        =   79
      Top             =   6780
      Width           =   5568
      _ExtentX        =   9816
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command16 
      Caption         =   "OPEN COPY LOGG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   10650
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   4500
      Width           =   1404
   End
   Begin VB.CommandButton Command15 
      Caption         =   "END SOON AS POSS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4056
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   1008
      Width           =   1404
   End
   Begin VB.CommandButton cmdScan_NO_LIST 
      Caption         =   "Scan Files NO LIST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2175
      Width           =   870
   End
   Begin VB.TextBox txtPath02 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1215
      TabIndex        =   50
      Text            =   "D:\"
      Top             =   660
      Width           =   4650
   End
   Begin VB.CommandButton Command12 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5925
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   705
      Width           =   540
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   6540
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   2475
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   6540
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   1875
      Width           =   1725
   End
   Begin VB.CheckBox Chk_SortOnDate 
      Caption         =   "Sort On Date"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   37
      Top             =   3300
      Width           =   1785
   End
   Begin VB.CommandButton CmdScanDir 
      Caption         =   "Scan Dir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1050
      Width           =   870
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help !"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   6420
      Picture         =   "ScanPathDemo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7845
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7725
      Top             =   3165
   End
   Begin VB.ComboBox cboMask 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   405
      TabIndex        =   1
      Text            =   "*.jpg"
      Top             =   330
      Width           =   6060
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5940
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   15
      Width           =   540
   End
   Begin VB.TextBox TxtPath 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   420
      TabIndex        =   0
      Text            =   "E:\My Music Zen"
      Top             =   30
      Width           =   5175
   End
   Begin VB.CheckBox chkSubFolders 
      Caption         =   "Search Sub-Folders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   2865
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox chkPatternMatching 
      Caption         =   "Pattern Matching"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox chkNormal 
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15
      TabIndex        =   2
      Top             =   1290
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkArchive 
      Caption         =   "Archive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15
      TabIndex        =   3
      Top             =   1515
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "Hidden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15
      TabIndex        =   4
      Top             =   1740
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Read Only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15
      TabIndex        =   5
      Top             =   1965
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkSystem 
      Caption         =   "System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15
      TabIndex        =   6
      Top             =   2175
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.ComboBox cboDate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2445
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1290
      Width           =   1545
   End
   Begin VB.ComboBox cboSizeType 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   3975
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2700
      Width           =   1365
   End
   Begin VB.ComboBox cboSize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2445
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2355
      Width           =   1545
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2430
      TabIndex        =   16
      Top             =   2700
      Width           =   1530
   End
   Begin VB.ComboBox cboSizeType 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3975
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3060
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2430
      TabIndex        =   18
      Top             =   3060
      Width           =   1530
   End
   Begin VB.CheckBox chkRefreshListView 
      Caption         =   "Refresh ListView (slower)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8250
      TabIndex        =   10
      Top             =   8325
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.CheckBox chkCopyMemory 
      Caption         =   "(display Date && Size)"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   11
      Top             =   3090
      Value           =   1  'Checked
      Width           =   1944
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2004
      Left            =   -12
      TabIndex        =   21
      Top             =   6276
      Width           =   13716
      _ExtentX        =   24183
      _ExtentY        =   3545
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CheckBox chkSort 
      Caption         =   "Sort Files(A-Z) while Scanning"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8160
      TabIndex        =   9
      Top             =   8055
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1605
      Width           =   870
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   0
      Left            =   2445
      TabIndex        =   13
      Top             =   1605
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   52101121
      CurrentDate     =   37299
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   1
      Left            =   2445
      TabIndex        =   14
      Top             =   1965
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   52101121
      CurrentDate     =   37296
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   2
      Left            =   4005
      TabIndex        =   44
      Top             =   1590
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      _Version        =   393216
      Format          =   52101122
      CurrentDate     =   37299
   End
   Begin VB.Label Label24 
      BackColor       =   &H0000FF00&
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1845
      TabIndex        =   85
      Top             =   2025
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label23 
      BackColor       =   &H008080FF&
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1860
      TabIndex        =   84
      Top             =   1665
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EXECUTE_ALL_LINK_FILE_FOLDER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9855
      TabIndex        =   82
      Top             =   3810
      Width           =   3390
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ARRIVE AT SOURCE DIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11070
      TabIndex        =   81
      Top             =   2475
      Width           =   3000
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   " ARRIVE AT DESTI DIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11070
      TabIndex        =   80
      Top             =   2190
      Width           =   2670
   End
   Begin VB.Label lblCount10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6510
      TabIndex        =   78
      Top             =   1830
      Width           =   9990
   End
   Begin VB.Label lblCount9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   324
      Left            =   24
      TabIndex        =   77
      Top             =   5928
      Width           =   13692
   End
   Begin VB.Label lblCount8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   36
      TabIndex        =   76
      Top             =   5616
      Width           =   13680
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MOVE ANYWHERE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   3984
      TabIndex        =   73
      Top             =   4188
      Width           =   1800
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COPY ANYWHERE TO ONE FOLDER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3975
      TabIndex        =   72
      Top             =   3600
      Width           =   3390
   End
   Begin VB.Label Command7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LIST FILES TO CLIP httrack"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7650
      TabIndex        =   71
      Top             =   4125
      Width           =   2670
      WordWrap        =   -1  'True
   End
   Begin VB.Label Command10 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USENET FROM OUTLOOK IN FOLDERS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   70
      Top             =   4200
      Width           =   3705
   End
   Begin VB.Label MOVE_IN_DATE_FOLDERS_JPGS 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MOVE IN DATE FOLDERS JPGS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   69
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Label Command9 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MOVE AnyWhere HardCode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -15
      TabIndex        =   68
      Top             =   5100
      Width           =   2685
   End
   Begin VB.Label Command13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CALL CRC DUPE MULTI MATRIX ON SMALLEST FILE LEAVE THE BIGGEST"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   8565
      TabIndex        =   67
      Top             =   2955
      Width           =   2700
      WordWrap        =   -1  'True
   End
   Begin VB.Label Command3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Find Replace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5508
      TabIndex        =   66
      Top             =   5124
      Width           =   1332
   End
   Begin VB.Label Command6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Grab String"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7650
      TabIndex        =   65
      Top             =   4425
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Command11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PATH TREE DIFF PLACE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7635
      TabIndex        =   64
      Top             =   5025
      Width           =   2415
      WordWrap        =   -1  'True
   End
   Begin VB.Label Command5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Files to Text List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7650
      TabIndex        =   63
      Top             =   3825
      Width           =   1620
      WordWrap        =   -1  'True
   End
   Begin VB.Label Command8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Find Replace .LNK'S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5508
      TabIndex        =   62
      Top             =   4824
      Width           =   2016
   End
   Begin VB.Label Command14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Del Emptys"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3972
      TabIndex        =   61
      Top             =   5124
      Width           =   1116
   End
   Begin VB.Label Command4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Del Boy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3972
      TabIndex        =   60
      Top             =   4824
      Width           =   816
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EXTRACT ALL WINRAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7650
      TabIndex        =   59
      Top             =   4710
      Width           =   2250
      WordWrap        =   -1  'True
   End
   Begin VB.Label Command1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CALL CRC DUPE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15
      TabIndex        =   58
      Top             =   3885
      Width           =   1620
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COPY ANYWHERE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3960
      TabIndex        =   57
      Top             =   3930
      Width           =   1770
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COPY CDROM ANYWHERE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3972
      TabIndex        =   56
      Top             =   4524
      Width           =   2556
   End
   Begin VB.Label lblCount7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   60
      TabIndex        =   55
      Top             =   5304
      Width           =   13668
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SEND DESIGNATED FOLDER TO ALL BLUETOOTHS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   8580
      TabIndex        =   53
      Top             =   2055
      Width           =   2460
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DONT MODIFY AUTO RUN NOW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   52
      Top             =   4500
      Width           =   3015
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "DESTI Path 02"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   51
      Top             =   705
      Width           =   1200
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FF00&
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1470
      TabIndex        =   46
      Top             =   2025
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CALL CRC DUPE Simulate Dont Del"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15
      TabIndex        =   45
      Top             =   3600
      Width           =   3435
   End
   Begin VB.Label lblCount1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6510
      TabIndex        =   43
      Top             =   0
      Width           =   9990
   End
   Begin VB.Label lblCount3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6510
      TabIndex        =   42
      Top             =   630
      Width           =   9990
   End
   Begin VB.Label lblCount4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6510
      TabIndex        =   41
      Top             =   930
      Width           =   9990
   End
   Begin VB.Label lblCount5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6510
      TabIndex        =   40
      Top             =   1230
      Width           =   9990
   End
   Begin VB.Label lblCount2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6510
      TabIndex        =   39
      Top             =   330
      Width           =   9990
   End
   Begin VB.Label lblCount6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6510
      TabIndex        =   38
      Top             =   1530
      Width           =   9990
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Working"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6975
      TabIndex        =   34
      Top             =   8025
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Date/Size Filter:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1830
      TabIndex        =   33
      Top             =   1095
      Width           =   1410
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Options:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15
      TabIndex        =   26
      Top             =   2430
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mask"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   15
      TabIndex        =   24
      Top             =   345
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   15
      TabIndex        =   22
      Top             =   30
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Attributes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   25
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Min"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1830
      TabIndex        =   30
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "File Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1845
      TabIndex        =   28
      Top             =   2355
      Width           =   585
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Max"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1830
      TabIndex        =   32
      Top             =   3120
      Width           =   300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1830
      TabIndex        =   29
      Top             =   1635
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "File Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1830
      TabIndex        =   27
      Top             =   1320
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1830
      TabIndex        =   31
      Top             =   1935
      Width           =   195
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB ME"
   End
   Begin VB.Menu Mnu_Cmd 
      Caption         =   "Commands"
      Begin VB.Menu Mnu_Com1 
         Caption         =   "Com1"
      End
      Begin VB.Menu Mnu_MatchBungDel 
         Caption         =   "MATCH BUNG AND DELETE FROM MATCH COMPARES"
      End
      Begin VB.Menu Mnu_MatchBungDateDesti 
         Caption         =   "MATCH BUNG DATE DESTINATION WAY OUT DATE CHK JOB"
      End
      Begin VB.Menu MNU_LIST_FILES_TO_CLIP_httrack 
         Caption         =   "LIST FILES TO CLIP httrack"
      End
   End
   Begin VB.Menu MNU_LOGG_COPY_PROPER 
      Caption         =   "LOGG FILE COPY PROPER"
   End
End
Attribute VB_Name = "ScanPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------Is This Sub Get on It at start      at       Bangers
Dim TX
Dim OX1, OY1
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public ASIZEMAX
Public DUPESOUCRE
Dim ITX()
Dim RTX() 'TX
Dim RRX() 'RX
Dim COPYSET_S()
Dim COPYSET_D()
Dim OPATH
Dim DDate2 As Date
Public Dog, OutPutPath, Xdate As Date, Tdate As Date, Zdate$, Ydate As Date, OLDFOLDER, OldSize

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

'Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long


'Private Type FILETIME
'    dwLowDateTime     As Long
'    dwHighDateTime    As Long
'End Type

'Private Type WIN32_FIND_DATA
'    dwFileAttributes  As Long
'    ftCreationTime    As FILETIME
'    ftLastAccessTime  As FILETIME
'    ftLastWriteTime   As FILETIME
'    nFileSizeHigh     As Long
'    nFileSizeLow      As Long
'    dwReserved0       As Long
'    dwReserved1       As Long
'    cFileName         As String * MAX_PATH
'    cAlternate        As String * 14
'End Type

'Private mWFD As WIN32_FIND_DATA

'Put This in a Bas Mobule
'Public m_CRC As clsCRC

'Public WxHex$

Public A1$, B1$, G1$, FF$

'Put in Sub Of Code
'Set m_CRC = New clsCRC
'm_CRC.Algorithm = CRC32
'WxHex$ = Hex(m_CRC.CalculateString(Form3.Text1.Text))
'WxHex$ = Hex(m_CRC.CalculateFile(Filename))


'Option Explicit

'##############################################################################################
'Purpose:   This project is a Demo for my cScanPath Class

'           This Class scans a specified path and returns the files it finds.
'           It has fairly comprehensive Filters. You can select files by:
'           Attributes (Normal, Hidden, Read Only, System etc)
'           File Size (>, <, Range)
'           File Date (From, To, Range)
'           File Extensions (multiple allowed i.e. *.txt;*.dat;*.tmp)

'           You can optionally scan sub-folders

'           To keep the demo simple I have only used attributes & Extensions
'           for Filter. For full example of this Class see WipeIt3 submission
       
'Author:    Richard Mewett ©2005
'##############################################################################################

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'We must declare with WithEvents to process the files returned


'Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer

Private Type POINTAPI
        x As Long
        y As Long
End Type




Private WithEvents SP As cScanPath
Attribute SP.VB_VarHelpID = -1


Private Sub SP_FileMatch(Filename As String, DFilename As String, Path As String)
    Dim LV As ListItem
    Dim uWIN32 As WIN32_FIND_DATA
    
    '####################################################################################################
    'This Event fires for each File found
    '####################################################################################################
    
    On Local Error GoTo GetFileError
    
    
'    If XX = 58 Then
        
        'lblCount1.Caption = mFileCount2
'        Print #1, "*" + Path + Filename
    
'    Exit Sub
'    End If
    
    
    
    
    
    With ListView1
        Set LV = .ListItems.Add(, , Filename)
        LV.SubItems(1) = Path
        'mod by matt l
        LV.SubItems(2) = DFilename
        
        
        
        
        
        
        
        
        
        
        If chkCopyMemory.Value Then
            'VB does not allows UDT's to be passed directly from a Class but we can access the structure
            'using pointers. We use CopyMemory to place the WIN32_FIND_DATA variable from the Class into
            'the uWIN32 declared in this Sub.
            
            'This is primarily a speed trick - we could retrieve the date & time information "ourselves"
            'using the standard VB functions or an API call but since we already have it...
            
            'CopyMemory ByVal uWIN32, ByVal SP.WFDPointer, ByVal Len(uWIN32)
            'LV.SubItems(3) = uWIN32.nFileSizeLow
            'LV.SubItems(4) = FormatFileDate(uWIN32.ftLastWriteTime)
        
                    'Remember Public FS
            If Filename <> "" Then
                Set F = Fs.getfile(Path + DFilename)
                adate1 = F.datelastmodified
                asize1 = F.Size
            Else
                Set F = Fs.GetFolder(Path)
                adate1 = F.datelastmodified
                'We May want this later Slow Down is Massive
                If OLDFOLDER <> Path Then
                asize1 = F.Size
                Else
                asize1 = OldSize
                End If
                OLDFOLDER = Path
                OldSize = asize1
            End If
'            LV.SubItems(2) = uWIN32.nFileSizeLow
'            LV.SubItems(3) = FormatFileDate(uWIN32.ftLastWriteTime)
            LV.SubItems(3) = asize1
            ASIZEMAX = ASIZEMAX + asize1
            
            LV.SubItems(4) = Format(adate1, "YYYY-MM-DD HH-MM-SS")




        
        End If
        
        

        
        
        'Refresh Count & Scroll the list every 10th file
        If (LV.Index Mod 50) = 0 Then
            If chkRefreshListView.Value Then
                .SelectedItem = LV
                .SelectedItem.EnsureVisible
'                DoEvents
            End If
            XXV = LV.Index
            'lblCount1.Caption = LV.Index
        End If
        
        
        
        
    End With
    If XXV < 1000 And XXV > 0 Then
'        Me.Refresh
        ListView1.Refresh
        Me.Refresh
        DoEvents
    End If

    
    
    Exit Sub
    
    
    

GetFileError:
    MsgBox Err.Description, vbCritical
End Sub


Sub DelJunkFrontPage()


ScanPath.ListView1.ListItems.Clear
ScanPath.TxtPath = "D:\MY WEBS\MatthewLan.Com Web\"
ScanPath.cboMask.Text = "*.*"
ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScanDir_Click
ScanPath.Show
Reset
For WE = ScanPath.ListView1.ListItems.COUNT To 1 Step -1
'    we = ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount1 = Trim(Str(WE))
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    A1$ = A1$ + ScanPath.ListView1.ListItems.Item(WE)
    outff = 0
    If InStr(A1$, "_private") > 0 Then outff = 1
    If InStr(A1$, "_vti") > 0 Then
        outff = 1
    End If
    
    If outff = 1 Then
        'ScanPath.ListView1.ListItems.Remove (we)
        On Error Resume Next
        Fs.DeleteFOLDER A1$, True
        If Err.Number > 0 Then Stop
        'err.description
        On Error GoTo 0

    End If
Next

End Sub

Sub INetContent()

If IsIDE = False Then
ScanPath.WindowState = 1
DoEvents
End If
DoEvents
ScanPath.Visible = True

OutPutPath = "M:\0 00 Art\Temp Inet Files JPGs\"

On Error Resume Next
MkDir "M:\0 00 Art\Temp Inet Files JPGs\#Gif"
MkDir "M:\0 00 Art\Temp Inet Files JPGs\#Png"
MkDir "M:\0 00 Art\Temp Inet Files JPGs\#Ico"
On Error GoTo 0


Open App.Path + "\Xdate.txt" For Input As #1
Line Input #1, Zdate$
Close #1
Ydate = DateValue(Zdate$) + TimeValue(Zdate$)


Label17 = "Process INet Content JPGs"

TxtPath.Text = "C:\Documents and Settings\Matt3\Local Settings\Temporary Internet Files\Content.IE5\"
TxtPath.Text = "D:\MY WEBS\MatthewLan.Com Web\"


cboDate.ListIndex = 0
DTPicker1(0) = Int(Ydate)

cboMask.Text = "*.jpg"
'cboSize.ListIndex = 1 'Less than
cboSize.ListIndex = 2 'Bigger Than
'cboSizeType(lCount).ListIndex = 0 'Byte
'cboSizeType(lCount).ListIndex = 2 'Meg
cboSizeType(lCount).ListIndex = 1 'K
txtSize(0) = 3
Call cmdScan_Click
TxtPath.Text = "C:\Documents and Settings\Matt4\Local Settings\Temporary Internet Files\Content.IE5\"
TxtPath.Text = "D:\MY WEBS\MatthewLan.Com Web\"
'Call cmdScan_Click

For WE = ListView1.ListItems.COUNT To 1 Step -1
    
    Label13 = ListView1.ListItems.COUNT
    DoEvents
    
    G1$ = ListView1.ListItems.Item(WE).SubItems(2)
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    
    Set F = Fs.getfile(A1$ + G1$)
    Tdate = F.datelastmodified
    If Tdate > Xdate Then Xdate = Tdate
    
    remo = 0
    If Tdate <= Ydate Then
        ListView1.ListItems.Remove (WE)
        remo = 1
    End If
    If remo = 0 Then
        hole = "Inet " + Format$(Now, "YYYY-MM-DD") + "\"
        GeText = LCase(Mid$(B1$, Len(B1$) - 2, 3)) + "\"
        If GeText = "jpg\" Then GeText = ""
        secext = "M:\0 00 Art\Temp Inet Files JPGs\" + hole + GeText + B1$
        If Fs.FileExists(secext) = True Then
            ListView1.ListItems.Remove (WE)
        End If
        secext = OutPutPath + hole + GeText + Mid$(B1$, 1, Len(B1$) - 4) + ".zip"
        If Fs.FileExists(secext) = True Then
            ListView1.ListItems.Remove (WE)
        End If
        secext = OutPutPath + hole + GeText + Mid$(B1$, 1, Len(B1$) - 4) + ".rar"
        If Fs.FileExists(secext) = True Then
            ListView1.ListItems.Remove (WE)
        End If
    End If
Next

Open App.Path + "\Xdate.txt" For Output As #1
Print #1, Xdate
Close #1
que = 0
    
hole = "INet " + Format$(Now, "YYYY-MM-DD") + "\"
On Error Resume Next
If que = 0 Then MkDir OutPutPath + hole: que = 1
On Error GoTo 0

For WE = ListView1.ListItems.COUNT To 1 Step -1
    
    Label13 = ListView1.ListItems.COUNT
    DoEvents
    
    G1$ = ListView1.ListItems.Item(WE).SubItems(2)
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    
    Set F = Fs.getfile(A1$ + G1$)
    
    
    GeText = LCase(Mid$(B1$, Len(B1$) - 2, 3)) + "\"
    If GeText = "jpg\" Then GeText = ""
    
    On Error Resume Next
    MkDir OutPutPath + hole + GeText
    On Error GoTo 0
    
    F.Copy OutPutPath + hole + GeText

Next

'-------------------------------------
'-------------------------------------
'DTPicker1(0).Value = vbNullString
ListView1.ListItems.Clear
TxtPath.Text = "C:\Documents and Settings\Matt3\Local Settings\Temporary Internet Files\Content.IE5\"
cboSize.ListIndex = 0 ' Not Used
cboMask.Text = "*.png;*.ico;*.gif"
Call cmdScan_Click

Exit Sub

For WE = ListView1.ListItems.COUNT To 1 Step -1
    
    Label13 = ListView1.ListItems.COUNT
    DoEvents
    
    G1$ = ListView1.ListItems.Item(WE).SubItems(2)
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    Set F = Fs.getfile(A1$ + G1$)
    Tdate = F.datelastmodified
    If Tdate > Xdate Then Xdate = Tdate
    remo = 0
    If Tdate <= Ydate Then
        ListView1.ListItems.Remove (WE)
        remo = 1
    End If
    If remo = 0 Then
        GeText = "#" + LCase(Mid$(B1$, Len(B1$) - 2, 3)) + "\"
        If Fs.FileExists(OutPutPath + GeText + B1$) = True Then
        ListView1.ListItems.Remove (WE)
        End If
    End If
Next

Open App.Path + "\Xdate.txt" For Output As #1
Print #1, Xdate
Close #1


For WE = ListView1.ListItems.COUNT To 1 Step -1
    
    Label13 = ListView1.ListItems.COUNT
    DoEvents
    
    G1$ = ListView1.ListItems.Item(WE).SubItems(2)
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    
    Set F = Fs.getfile(A1$ + G1$)

    GeText = "#" + LCase(Mid$(B1$, Len(B1$) - 2, 3)) + "\"
    F.Copy OutPutPath + GeText
Next

ListView1.ListItems.Clear
TxtPath.Text = "C:\Documents and Settings\Matt3\Local Settings\Temporary Internet Files\Content.IE5\"
cboSize.ListIndex = 0 ' Not Used
cboMask.Text = "*.png;*.ico;*.gif"
Call cmdScan_Click
Call cmdScanDir_Click

If IsIDE = False Then End

End Sub


Sub Bangers()

Label17 = "Process Banger Favs Quick"
'Call Banger_Favs
Label17 = "Working"

tpath1$ = "M:\04 Music ---\03 My Music Zen\"
tpath2$ = "M:\04 Music ---\04 My Music\"
tpath3$ = "M:\04 Music ---\Del\"
'MkDir tpath3$
TxtPath.Text = tpath1$

cboMask.Text = "*.mp3"

TxtPath.Text = "F:\04 Music ---\04 My Music"
TxtPath.Text = "F:\04 Music ---\04 My Music Talking Books"

Dim ra(30)

ra(1) = "M:\05 Media\MIDI"
ra(2) = "M:\04 Music ---\03 My Music Zen"
ra(3) = "M:\04 Music ---\04 My Music"
ra(4) = "M:\04 Music ---\05 Whole Lot"
ra(5) = "M:\04 Music ---\00 TBooks All"

ra(1) = "M:\04 Music ---"
ra(2) = "M:\Video\00 Vid Films"
ra(3) = "M:\Video\00 Vid XXX"
'ra(3) = "M:\04 Music ---\04 My Music Talking Books"

'ra(1) = "M:\04 Music ---\"
'ra(2) = "M:\04 Music ---\04 My Music Talking Books"
'ra(3) = "S:\Tbooks"
'ra(4) = "S:\My Music Talking Books"



For Ret = 1 To 3
DoEvents
If Right$(ra(Ret), 1) <> "\" Then ra(Ret) = ra(Ret) + "\"

TxtPath.Text = ra(Ret)

cboSize.ListIndex = 1 'Less than
txtSize(0) = 1

If Dog = True Then Exit Sub

Call cmdScan_Click

If Dog = True Then Exit Sub
DoEvents
If Dog = True Then Exit Sub

For WE = ListView1.ListItems.COUNT To 1 Step -1
    
    Label13 = WE
    DoEvents
    
    G1$ = ListView1.ListItems.Item(WE).SubItems(2)
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    w = DeleteFile(A1$ + G1$)

Next

Call cmdScanDir_Click

For WE = ListView1.ListItems.COUNT To 1 Step -1
  
    Label14 = WE
    DoEvents
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    A1$ = A1$ + B1$ + "\"
  
    
    Set F = Fs.GetFolder(A1$)
    
    If F.Size = 0 Then
        ListView1.ListItems.Remove (WE)
    End If
Next

A1$ = TxtPath.Text

Call Tagg2


For WE = 1 To ListView1.ListItems.COUNT
  
    Label16 = WE
    DoEvents
    
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)
    A1$ = A1$ + B1$ + "\"
    
    If InStr(A1$, " - 20") > 0 And InStr(LCase(A1$), "banging") > 0 Then
        Call TaggBang
    Else
        Call Tagg2
    End If

Next

Next

Label17 = "Done"



End

End Sub



Sub Tagg2()
 
 

FF$ = Mid$(A1$, Len(TxtPath.Text) + 1)
If FF$ = "" Then
FF$ = Mid$(A1$, 4)
End If
FF$ = Left$(FF$, Len(FF$) - 1)

Do
ttgx = InStr(FF$, "\")
If ttgx > 0 Then
FF$ = Mid$(FF$, 1, ttgx - 1) + " -- " + Mid$(FF$, ttgx + 1)
End If
Loop Until ttgx = 0

Call WriteTagg

End Sub

Sub TaggBang()
 
    
TY$ = ""
If InStr(A1$, "My Music\") Then TY$ = "My Music"
If InStr(A1$, "My Music Zen\") Then TY$ = "My Music Zen"
If TY$ = "My Music" And InStr(A1$, "Favs") Then TY$ = "My Music - Favs"
If TY$ = "My Music Zen" And InStr(A1$, "Favs") Then TY$ = "My Music Zen - Favs"
If TY$ = "My Music" And InStr(A1$, " Bang Sets") Then TY$ = "My Music - Sets"
If TY$ = "My Music Zen" And InStr(A1$, " Bang Sets") Then TY$ = "My Music Zen - Sets"
If TY$ = "My Music" And InStr(A1$, " Bang Fav Sets") Then TY$ = "My Music - Fav Sets"
If TY$ = "My Music Zen" And InStr(A1$, " Bang Fav Sets") Then TY$ = "My Music Zen - Fav Sets"
'03 Bang Sets
    
    
gg$ = Mid$(A1$, InStrRev(A1$, "\", Len(A1$) - 1) + 1)
TX = Left$(gg$, Len(gg$) - 1)

FF$ = "---------------[[(({{ " + TX + " - " + TY$ + " }}))]]---------------.mp3"


On Error Resume Next
Do
fg1 = FreeFile
Err.Clear
Open A1$ + FF$ For Output As fg1
Close fg1
If Err.Number > 0 Then FF$ = Left$(FF$, Len(FF$) - 5) + ".mp3"
Loop Until Err.Number = 0


'Call WriteTagg



End Sub



Sub WriteTagg()

On Local Error Resume Next
'Kill A1$ + "0#-" + FF$ + "-^.mp3"
'Kill A1$ + "01#-" + FF$ + "-^.mp3"
'Err.Clear
easy = Dir$(A1$ + "\*.txt") <> ""
 '
Do
    Err.Clear
'    If InStr(FF$, "Art And Music") > 0 Then Stop

    fg1 = FreeFile
    Open A1$ + "0#01--" + FF$ + "-^.mp3" For Output As fg1
    Close fg1
    FF$ = Left$(FF$, Len(FF$) - 1)
Loop Until Err.Number = 0
    
Do
    Err.Clear
    ek$ = String$(Len(FF$) + 20, "-")
    fg1 = FreeFile
    Open A1$ + "0#00--" + ek$ + "-^.mp3" For Output As fg1
    Close fg1
    If easy = True Then
        fg1 = FreeFile
        Open A1$ + "0#02----- Text File In Dir -----^.mp3" For Output As fg1
        Close fg1
    End If
    fg1 = FreeFile
    Open A1$ + "0#02--" + ek$ + "-^.mp3" For Output As fg1
    Close fg1
    FF$ = Left$(FF$, Len(FF$) - 1)
Loop Until Err.Number = 0
    
    
On Local Error GoTo 0
    


End Sub


Private Sub cboSize_Click()
    'cboSize.ListIndex = 2
    'txtSize(0).Text = 1
    
    'cboSizeType(0).Visible = True
    'cboSizeType(0).

    
    Select Case cboSize.ListIndex
    Case 0
        lblSize(0).Visible = False
        lblSize(1).Visible = False
        txtSize(0).Visible = False
        txtSize(1).Visible = False
        cboSizeType(0).Visible = False
        cboSizeType(1).Visible = False
    Case 1, 2
        lblSize(0).Caption = "Size"
        lblSize(0).Visible = True
        lblSize(1).Visible = False
        txtSize(0).Visible = True
        txtSize(1).Visible = False
        cboSizeType(0).Visible = True
        cboSizeType(1).Visible = False
    Case Else
        lblSize(0).Caption = "Min"
        lblSize(1).Caption = "Max"
        lblSize(1).Visible = True
        lblSize(0).Visible = True
        txtSize(0).Visible = True
        txtSize(1).Visible = True
        cboSizeType(0).Visible = True
        cboSizeType(1).Visible = True
    End Select
End Sub


Private Sub cboSizeType_Change(Index As Integer)
'cboSizeType = 0
End Sub

Private Sub chkCopyMemory_Click()
    
    chkSort.Enabled = (chkCopyMemory.Value = vbUnchecked)
    
End Sub


Private Sub chkSort_Click()
    chkCopyMemory.Enabled = (chkSort.Value = vbUnchecked)
End Sub

Private Sub cmdBrowse_Click()
    TxtPath.Text = GetFolder(Me.hWnd, "Scan Path:", TxtPath.Text)
    fg1 = FreeFile
    Open App.Path + "\Scan Path.txt" For Output As #fg1
    Print #fg1, TxtPath.Text
    Close #fg1

'    fg1 = FreeFile
'    Open App.Path + "\Scan Path.txt" For Input As #fg1
'    Line Input #fg1, ae$
'    Close #fg1
'    txtPath.Text = ae$



End Sub

Private Sub cmdHelp_Click()
    'txtHelp.Visible = Not txtHelp.Visible
End Sub

Public Sub cmdScan_Click()
    Dim lCount As Long
    Dim lSize(1) As Long
    
    'chkSort.Enabled = True
    
    
    If cmdScan.Caption = "Scan Files" Then
        '####################################################################################################
        'Process our Selection Criteria
        For lCount = 0 To 1
            lSize(lCount) = Val(txtSize(lCount).Text)
        
            Select Case cboSizeType(lCount).ListIndex
            Case 1: lSize(lCount) = lSize(lCount) * 1024
            Case 2: lSize(lCount) = (lSize(lCount) * 1024) * 1024
            End Select
        Next lCount
            
        If txtSize(1).Visible Then
            If lSize(0) = 0 Then
                MsgBox "Maximium Size value must exceed 0", vbExclamation
                txtSize(1).SetFocus
                Exit Sub
            ElseIf lSize(1) <= lSize(0) Then
                MsgBox "Maximium Size value must exceed Minimum Size", vbExclamation
                txtSize(1).SetFocus
                Exit Sub
            End If
        End If
        
        '####################################################################################################
        'Reset Form
        cmdScan.Caption = "Stop"
        
        lblCount1.Caption = "-"
 
        'Screen.MousePointer = vbHourglass
        'ListView1.ListItems.Clear
        
        If chkRefreshListView.Value = vbUnchecked Then
            LockWindowUpdate ListView1.hWnd
        End If
        
        '####################################################################################################
        'Create our Scan Object
        Set SP = New cScanPath
        
        With SP
            'Attributes
            .Archive = chkArchive.Value
            .Compressed = True
            .Hidden = chkHidden.Value
            .Normal = chkNormal.Value
            .ReadOnly = chkReadOnly.Value
            .System = chkSystem.Value
            
            'Date/Size (NOTE: you don't have to use these!)
            Select Case cboDate.ListIndex
            Case 0
                .DateType = Modified
            Case 1
                .DateType = Created
            Case 2
                .DateType = LastAccessed
            End Select
            
            If IsDate(DTPicker1(0).Value) Then
                .FromDate = DTPicker1(0).Value
            End If
            If IsDate(DTPicker1(1).Value) Then
                .ToDate = DTPicker1(1).Value
            End If
            
            Select Case cboSize.ListIndex
            Case 1
                .MaximumSize = lSize(0)
            Case 2
                .MinimumSize = lSize(0)
            Case 3
                .MinimumSize = lSize(0)
                .MaximumSize = lSize(1)
            End Select
            
            'Mask
            .Filter = cboMask.Text
            
            'Go - that was easy wasn't it!
            .StartScan TxtPath, chkSubFolders.Value, chkSort.Value And chkSort.Enabled, chkPatternMatching
            
            If Chk_SortOnDate.Value = vbChecked Then
                ListView1.SortOrder = lvwAscending
                ListView1.SortKey = 3
                ListView1.Sorted = True
                ListView1.Sorted = False
            Else
                ListView1.SortOrder = lvwAscending
                ListView1.SortKey = 0
                ListView1.Sorted = True
                ListView1.SortKey = 1
                ListView1.Sorted = True
                ListView1.Sorted = False
            End If
        
        End With
        
       With ListView1
            If .ListItems.COUNT > 0 Then
                .SelectedItem = .ListItems(1)
            End If
            
            lblCount1.Caption = .ListItems.COUNT
        End With
        
        LockWindowUpdate 0
        
        cmdScan.Caption = "Scan Files"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
        'End
    End If


cmdScan.Caption = "Scan Files"




'We always use direct scan as is better to sort after job then during thats all we need

Sorti = 0

Select Case Sorti
Case 0

Case 1
    'Sort On Date
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 4
    ListView1.Sorted = True
    ListView1.Sorted = False

Case 2
    'Sort On Dos File Name
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 2
    ListView1.Sorted = True
    ListView1.Sorted = False

Case 3
    'Sort On Size
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 3
    ListView1.Sorted = True
    ListView1.Sorted = False

Case 4
'sort of PATH then FILE
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 0
    ListView1.Sorted = True
    ListView1.SortKey = 1
    ListView1.Sorted = True
    ListView1.Sorted = False

Case 5
'sort of FILE then PATH
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 1
    ListView1.Sorted = True
    ListView1.SortKey = 0
    ListView1.Sorted = True
    ListView1.Sorted = False

End Select

End Sub



Public Sub cmdScanDir_Click()
    Dim lCount As Long
    Dim lSize(1) As Long
    
    If CmdScanDir.Caption = "Scan Dir" Then
        '####################################################################################################
        'Process our Selection Criteria
        For lCount = 0 To 1
            lSize(lCount) = Val(txtSize(lCount).Text)
        
            Select Case cboSizeType(lCount).ListIndex
            Case 1: lSize(lCount) = lSize(lCount) * 1024
            Case 2: lSize(lCount) = (lSize(lCount) * 1024) * 1024
            End Select
        Next lCount
            
        If txtSize(1).Visible Then
            If lSize(0) = 0 Then
                MsgBox "Maximium Size value must exceed 0", vbExclamation
                txtSize(1).SetFocus
                Exit Sub
            ElseIf lSize(1) <= lSize(0) Then
                MsgBox "Maximium Size value must exceed Minimum Size", vbExclamation
                txtSize(1).SetFocus
                Exit Sub
            End If
        End If
        
        '####################################################################################################
        'Reset Form
        cmdScan.Caption = "Stop"
        lblCount2.Caption = "-"
 
        Screen.MousePointer = vbHourglass
        ListView1.ListItems.Clear
        
        If chkRefreshListView.Value = vbUnchecked Then
            LockWindowUpdate ListView1.hWnd
        End If
        
        '####################################################################################################
        'Create our Scan Object
        Set SP = New cScanPath
        
        With SP
            'Attributes
            .Archive = chkArchive.Value
            .Compressed = True
            .Hidden = chkHidden.Value
            .Normal = chkNormal.Value
            .ReadOnly = chkReadOnly.Value
            .System = chkSystem.Value
            
            'Date/Size (NOTE: you don't have to use these!)
            Select Case cboDate.ListIndex
            Case 0
                .DateType = Modified
            Case 1
                .DateType = Created
            Case 2
                .DateType = LastAccessed
            End Select
            
            If IsDate(DTPicker1(0).Value) Then
                .FromDate = DTPicker1(0).Value
            End If
            If IsDate(DTPicker1(1).Value) Then
                .ToDate = DTPicker1(1).Value
            End If
            
            Select Case cboSize.ListIndex
            Case 1
                .MaximumSize = lSize(0)
            Case 2
                .MinimumSize = lSize(0)
            Case 3
                .MinimumSize = lSize(0)
                .MaximumSize = lSize(1)
            End Select
            
            'Mask
            .Filter = cboMask.Text
            
            'Go - that was easy wasn't it!
            .StartScanDir TxtPath, chkSubFolders.Value, chkSort.Value And chkSort.Enabled, chkPatternMatching
        
            
            'Correct Yes 1st one column sort keeps sort then another
            ListView1.SortOrder = lvwAscending
            ListView1.SortKey = 0
            ListView1.Sorted = True
            ListView1.SortKey = 1
            ListView1.Sorted = True
            ListView1.Sorted = False
        
            If Chk_SortOnDate.Value = vbChecked Then
                ListView1.SortOrder = lvwAscending
                ListView1.SortKey = 3
                ListView1.Sorted = True
                ListView1.Sorted = False
            End If
        
        End With
        
       With ListView1
            If .ListItems.COUNT > 0 Then
                .SelectedItem = .ListItems(1)
            End If
            
            lblCount1.Caption = .ListItems.COUNT
        End With
        
        LockWindowUpdate 0
        
        cmdScan.Caption = "Scan"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
   End If
'

lblCount1.Caption = ListView1.ListItems.COUNT
    

End Sub




Private Sub Command1_Click()
'Label5
'Command1_Click




Call CRCDUPE_X

End Sub

Private Sub Command10_Click()


'ScanPath.txtPath = "K:\00 HTTrack\Techno BeatPlexity.com MusicGateWay\http___www.beatplexity.com_musicgateway\www.beatplexity.com\musicgateway\"
'ScanPath.txtPath = "K:\00 HTTrack\Techno PoMac.Net Swarm\Http___pomac.netswarm.net_music.html_music\pomac.netswarm.net\music"

'ScanPath.txtPath = "X:\00 Blue Tooth Exchange Folder\IMAGES 2009\"

'ScanPath.txtPath = "M:\Temp\alt.support.schizophrenia JUST BORG"
'
'24hoursupport.helpdesk
'ScanPath.txtPath = "M:\Temp\24hoursupport.helpdesk"
ScanPath.TxtPath = "M:\Temp\alt.usenet.kooks"

If Mid(ScanPath.TxtPath, Len(ScanPath.TxtPath) - 1, 1) <> "\" Then ScanPath.TxtPath = ScanPath.TxtPath + "\"



'    'MOST HARD CODED
'    ScanPath.Text1 = "    Loop" + vbCrLf + "    " + vbCrLf + "    " + "FindClose lResult"
'    ScanPath.Text2 = "        FindClose lResult" + vbCrLf + "    " + vbCrLf + "    " + "Loop"


'ScanPath.Text2 = """>"
'ScanPath.Text1 = "http:"



ScanPath.ListView1.ListItems.Clear

ScanPath.cboMask.Text = "*.NWS"
ScanPath.chkSubFolders = vbUnchecked
Call ScanPath.cmdScan_Click
ScanPath.Show

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"
    
'B1 = ScanPath.ListView1.ListItems.Item(1)
'RXRIME = Format(DateValue(Mid(B1, 8, 10)), "YYYY-MM-MMM")
'FN = ScanPath.txtPath.Text + "HTTP's From Email's -- " + RXRIME + ".txt"

'If Dir(FN) <> "" Then Kill FN
'UU = FreeFile
'Open FN For Append As #UU



For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    
    
'    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " DEALT WITH FILES"
    
    DoEvents
    
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
    
    TT = TT + "-*/" + B1 + vbCrLf
    
    'Set F = Fs.getfile(A1 + B1)
    'ADATE1 = F.datelastmodified
    
    Open A1 + G1 For Binary As #1
    OO = Input(LOF(1), 1)
    Close #1
    
    
    'Clipboard.Clear
    'Clipboard.SetText oo
    
   
   adate1 = Mid(OO, InStr(OO, vbCrLf + "Date: ") + 2 + 6)
   adate1 = Mid(adate1, 1, InStr(adate1, vbCrLf))
   If InStr(adate1, ",") > 0 Then
        adate1 = Mid(adate1, InStr(adate1, ",") + 2)
   End If
   
   
   
            DDate = adate1
            'DDATE = Mid(DDATE, InStr(DDATE, ";") + 2)
                
            DDate = Replace(DDate, "GMT", "")
            DDate = Replace(DDate, "", "")
            DDate = Replace(DDate, "(BST)", "")
            UU = ""
            If InStr(DDate, "-") > 0 Then
                UU = DDate
                UU = Mid(DDate, 1, InStr(DDate, "-") - 2)
                UU1 = Mid(DDate, InStr(DDate, "-") + 1, 2)
                'UU = Mid(UU, 6)
                UU = DateValue(UU) + TimeValue(UU)
                uu2 = Val(UU1)
                UU = UU - TimeSerial(uu2, 0, 0)
                DDate2 = UU
            End If
            If InStr(DDate, "+") > 0 Then
                UU = DDate
                UU = Mid(DDate, 1, InStr(DDate, "+") - 2)
                UU1 = Mid(DDate, InStr(DDate, "+") + 1, 2)
                'UU = Mid(UU, 6)
                UU = DateValue(UU) + TimeValue(UU)
                uu2 = Val(UU1)
                UU = UU + TimeSerial(uu2, 0, 0)
                DDate2 = UU
            End If
   
            If UU <> "" Then DDate = Format(DDate2, "DD MM YYYY")
   
        'ADATE1 = F.datelastmodified
    
    
    
    'YY = Format(DDATE, "YYYY") + " 0" + Format(DDATE, "MM") + "(" + Format(DDATE, "MMM") + ") " + Format(DDATE, "DD DDD")
    YY = Format(DDate, "YYYY") + "-" + Format(DDate, "MM") + "-" + Format(DDate, "DD") + "  " + Format(DDate, "hh-mm-ss")
    
    
            
               SUBJ = B1
            SUBJ = Replace(SUBJ, "    ", "")
            SUBJ = Replace(SUBJ, vbCr, "")
            SUBJ = Replace(SUBJ, vbLf, "")
            SUBJ = Replace(SUBJ, Chr(9), "")
            SUBJ = Replace(SUBJ, """", "")
            SUBJ = Replace(SUBJ, "\", "")
            SUBJ = Replace(SUBJ, "?", "")
            SUBJ = Replace(SUBJ, ":", "-")
            SUBJ = Replace(SUBJ, "/", "-")
            SUBJ = Replace(SUBJ, "=", "-")
            SUBJ = Replace(SUBJ, "*", "-")
            SUBJ = Replace(SUBJ, "<", "")
            SUBJ = Replace(SUBJ, ">", "")
            SUBJ = Replace(SUBJ, "|", "")
            If SUBJ <> B1 Then B1 = SUBJ
    
    
    YY = YY + "#" + B1
    
    
    
    yy2 = Format(DDate, "YYYY") + " 0" + Format(DDate, "MM") + "(" + Format(DDate, "MMM") + ")"
    
    'Name A1 + G1 As A1 + YY
    
    'MOVE INTO FOLDER
    If Fs.FolderExists(ScanPath.TxtPath + yy2) = False Then MkDir ScanPath.TxtPath + yy2
    'If Fs.FileExists(ScanPath.txtPath + yy + "\" + yy2) = False Then
        
        LONGLOOP = 0
        Do
        LONGLOOP = LONGLOOP + 1
        If LONGLOOP > 50 Then Stop
        On Error Resume Next
        Set F = Fs.getfile(A1 + G1)
        Err.Clear
        If Fs.FileExists(ScanPath.TxtPath + "\" + yy2 + "\" + YY) = True Then Exit Do
        F.Move ScanPath.TxtPath + "\" + yy2 + "\" + YY
        
        'ERR.DESCRIPTION
        '>256 CHARS PATH
        If Err.Number = 76 Then
            YY = Mid(YY, 1, Len(YY) - 10) + ".nws"
        End If
        If Len(YY) < Len(ScanPath.TxtPath) And Err.Number > 0 Then
            Exit For
        End If
        Loop Until Err.Number = 0
        
        
    'End If
    
    'MOVE WHERE IS
    'If Fs.FileExists(A1$ + YY2) = False Then
        'F.Move A1$ + YY2
    'End If
    
    
    
    'F.Move A1 + YY + "\" + YY + " - " + B1
    
    
    
    
    
    
Next




End Sub

Private Sub Command11_Click()

Call PATH_TREE_ANOTHER_PLACE

End Sub

Private Sub Command13_Click()
Call CRCDUPE_MULTI_MATRIX_ON_SMALLEST_FILE
End Sub

Private Sub Command14_Click()
Call DelEMPTYS
End Sub

Public Sub cmdScan_NO_LIST_Click()
   
    XX = 58
    Dim lCount As Long
    Dim lSize(1) As Long
    mFileCount2 = 0

    With ListView1
        .ColumnHeaders.Item(1).Width = 1
    End With
    'chkSort.Enabled = True
    
    
    If cmdScan.Caption = "Scan Files" Then
        '####################################################################################################
        'Process our Selection Criteria
        For lCount = 0 To 1
            lSize(lCount) = Val(txtSize(lCount).Text)
        
            Select Case cboSizeType(lCount).ListIndex
            Case 1: lSize(lCount) = lSize(lCount) * 1024
            Case 2: lSize(lCount) = (lSize(lCount) * 1024) * 1024
            End Select
        Next lCount
            
        If txtSize(1).Visible Then
            If lSize(0) = 0 Then
                MsgBox "Maximium Size value must exceed 0", vbExclamation
                txtSize(1).SetFocus
                Exit Sub
            ElseIf lSize(1) <= lSize(0) Then
                MsgBox "Maximium Size value must exceed Minimum Size", vbExclamation
                txtSize(1).SetFocus
                Exit Sub
            End If
        End If
        
        '####################################################################################################
        'Reset Form
        cmdScan.Caption = "Stop"
        
        lblCount1.Caption = "-"
 
        'Screen.MousePointer = vbHourglass
        'ListView1.ListItems.Clear
        
        If chkRefreshListView.Value = vbUnchecked Then
            'LockWindowUpdate ListView1.hWnd
        End If
        
        '####################################################################################################
        'Create our Scan Object
        Set SP = New cScanPath
        
        With SP
            'Attributes
            .Archive = chkArchive.Value
            .Compressed = True
            .Hidden = chkHidden.Value
            .Normal = chkNormal.Value
            .ReadOnly = chkReadOnly.Value
            .System = chkSystem.Value
            
            'Date/Size (NOTE: you don't have to use these!)
            Select Case cboDate.ListIndex
            Case 0
                .DateType = Modified
            Case 1
                .DateType = Created
            Case 2
                .DateType = LastAccessed
            End Select
            
            If IsDate(DTPicker1(0).Value) Then
                .FromDate = DTPicker1(0).Value
            End If
            If IsDate(DTPicker1(1).Value) Then
                .ToDate = DTPicker1(1).Value
            End If
            
            Select Case cboSize.ListIndex
            Case 1
                .MaximumSize = lSize(0)
            Case 2
                .MinimumSize = lSize(0)
            Case 3
                .MinimumSize = lSize(0)
                .MaximumSize = lSize(1)
            End Select
            
            'Mask
            .Filter = cboMask.Text
            
            'Go - that was easy wasn't it!
            .StartScan_USE_NO_LIST TxtPath, chkSubFolders.Value, chkSort.Value And chkSort.Enabled, chkPatternMatching
            
            If Chk_SortOnDate.Value = vbChecked Then
                ListView1.SortOrder = lvwAscending
                ListView1.SortKey = 3
                ListView1.Sorted = True
                ListView1.Sorted = False
            Else
                ListView1.SortOrder = lvwAscending
                ListView1.SortKey = 0
                ListView1.Sorted = True
                ListView1.SortKey = 1
                ListView1.Sorted = True
                ListView1.Sorted = False
            End If
        
        End With
        
       With ListView1
            If .ListItems.COUNT > 0 Then
                .SelectedItem = .ListItems(1)
            End If
            
            lblCount1.Caption = .ListItems.COUNT
        End With
        
        LockWindowUpdate 0
        
        cmdScan.Caption = "Scan Files"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
        'End
    End If


cmdScan.Caption = "Scan Files"




'We always use direct scan as is better to sort after job then during thats all we need

Sorti = 0

Select Case Sorti
Case 0

Case 1
    'Sort On Date
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 4
    ListView1.Sorted = True
    ListView1.Sorted = False

Case 2
    'Sort On Dos File Name
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 2
    ListView1.Sorted = True
    ListView1.Sorted = False

Case 3
    'Sort On Size
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 3
    ListView1.Sorted = True
    ListView1.Sorted = False

Case 4
'sort of PATH then FILE
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 0
    ListView1.Sorted = True
    ListView1.SortKey = 1
    ListView1.Sorted = True
    ListView1.Sorted = False

Case 5
'sort of FILE then PATH
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 1
    ListView1.Sorted = True
    ListView1.SortKey = 0
    ListView1.Sorted = True
    ListView1.Sorted = False

End Select

End Sub

Private Sub Command15_Click()
EXITVAR = True
End Sub

Private Sub Command16_Click()

Shell "NOTEPAD D:\TEMP\LOGG MULTI COPY.TXT", vbMaximizedFocus

End Sub

Private Sub Command2_Click()
Call WINRAR_EXTRACT

End Sub

Private Sub Command3_Click()
    'MOST HARD CODED
    Call FIND_REPLACE





End Sub

Private Sub Command4_Click()
Call DelBoy




End Sub

Private Sub Command5_Click()

Call FilesToTextList

End Sub

Private Sub Command6_Click()
    'MOST HARD CODED
    Text1 = "Start to Search"
    Text2 = "How Does End"
    
    'Text1 = "http:"
    'Text2 = """"
    
    Call GRAB_STRING
    


End Sub

Private Sub Command7_Click()


ScanPath.TxtPath = "C:\HTTRACK\Techno PoMac.Net Swarm\Http___pomac.netswarm.net_music.html_music\pomac.netswarm.net\music\"

'ScanPath.txtPath = "C:\HTTRACK\Techno BeatPlexity.com MusicGateWay\http___www.beatplexity.com_musicgateway\www.beatplexity.com\musicgateway\"

'    'MOST HARD CODED
'    ScanPath.Text1 = "    Loop" + vbCrLf + "    " + vbCrLf + "    " + "FindClose lResult"
'    ScanPath.Text2 = "        FindClose lResult" + vbCrLf + "    " + vbCrLf + "    " + "Loop"


'ScanPath.Text2 = """>"
'ScanPath.Text1 = "http:"

'GREATER THAN ONE BYTE
cboSize.ListIndex = 2
txtSize(0).Text = 1

ScanPath.ListView1.ListItems.Clear

ScanPath.cboMask.Text = "*.MP3;*.ogg"
ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScan_Click
ScanPath.Show

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"
    
'B1 = ScanPath.ListView1.ListItems.Item(1)
'RXRIME = Format(DateValue(Mid(B1, 8, 10)), "YYYY-MM-MMM")
'FN = ScanPath.txtPath.Text + "HTTP's From Email's -- " + RXRIME + ".txt"

'If Dir(FN) <> "" Then Kill FN
'UU = FreeFile
'Open FN For Append As #UU

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    
    
'    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " DEALT WITH FILES"
    
    DoEvents
    
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    'TT = TT + "-*/" + B1 + vbCrLf
    'TT = TT + "-*/" + Mid(B1, 1, 4) + "*" + vbCrLf
    
    'GET WHOLE LINE FILE
    TT = TT + "-*/" + Mid(B1, 1, Len(B1) - 4) + "*" + vbCrLf
    
    
Next

Clipboard.Clear
Clipboard.SetText TT



End Sub

Private Sub Command8_Click()

ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask.Text = "*.LNK"
ScanPath.chkSubFolders = vbChecked

ScanPath.TxtPath = "E:\01 VB Shell Folders\00 Shell Notepad Loader"
ScanPath.TxtPath = "E:\01 VB Shell Folders\"

Call ScanPath.cmdScan_Click


'MY DOCS
sPersonalFolder = GetSpecialfolder(CSIDL_PERSONAL)

ScanPath.Show

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"
    

Dim RR As String, D1, TT, D4 As String, D5, S1, D3

Debug.Print vbCrLf

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    ScanPath.lblCount2 = Trim(Str(WE)) + " CHKED"
    
    DoEvents
    
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)




'Load File From Link
D1 = A1 + B1
Open A1 + B1 For Binary As #1
RR = Space$(LOF(1))
Get #1, 1, RR
Close #1
'If Mid$(rr$, &H568, 3) = ":" + Chr$(0) + "\" Then
'rr$ = Mid$(rr$, &H566)
'End If
S1 = InStr(105, RR, ":\") - 1
S2 = Len(RR)
If S1 > 0 Then

TT = Mid$(RR, S1)
D1 = Mid$(TT, 1, InStr(TT, Chr$(0)) - 1)


d2 = 0
d2 = InStr(D1, "My Documents")
If d2 > 0 Then
D3 = InStr(d2, D1, "\")
If D3 > 0 Then
D4 = "'%PERSONAL%" + Mid$(D1, D3) + "'"
D4 = "%PERSONAL%" + Mid$(D1, D3)
D4 = sPersonalFolder + Mid$(D1, D3)

'If Dir(A1 + B1) <> "" Then Kill A1 + B1

G1 = Mid(B1, 1, InStrRev(B1, ".") - 1)

D5 = Mid(D4, 1, InStrRev(D4, "\"))

If InStr(LCase(D4), "my web") > 0 Then
    'Stop
    D3 = InStr(LCase(D4), "my web")
    D4 = "D:\" + Mid$(D4, D3)
    D5 = Mid(D4, 1, InStrRev(D4, "\"))

End If

'If InStr(LCase(D4), "my web") = 0 Then

'Debug.Print D4
'Debug.Print A1 + B1
'Debug.Print G1
'Debug.Print D5

If Fs.FileExists(D4) = True Or Fs.FolderExists(D4) = True Then

        CNT3 = CNT3 + 1
        ScanPath.lblCount3 = Trim(Str(CNT3)) + " DEALT WITH"
        Debug.Print "CONVERT LINK "
        'Call Create_ShortCut(D4, A1 + B1, G1, D5)
    
    Else
    
        'MsgBox "FILE NOT EXIST" + VBCLRF + D4
        Debug.Print "FILE NOT EXIST ------ " + A1 + B1
        Debug.Print "FILE NOT EXIST ------ " + D4
        Debug.Print
    
    End If


'Else

    'Stop

'End If


'Stop
End If

End If 'S1
End If 'D3

'Exit Sub

Next

End Sub

Private Sub Label14_Click()

ScanPath.ListView1.ListItems.Clear

ScanPath.chkSubFolders = vbUnchecked

ScanPath.TxtPath = "D:\# MY DOCS\# 01 My Documents\00 Blue Tooth Exchange Folder\00 CONTACTS\## CONS BLUE SEND FOLDER"

Call ScanPath.cmdScan_Click

'On Error Resume Next
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"

If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit Sub
fp = App.Path + "\TOSEND.vcf"
If Dir(fp) <> "" Then Kill fp

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    ScanPath.lblCount2 = Trim(Str(WE)) + " ToT Process Files"
    DoEvents
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    SOURCE1 = A1 + B1
    Open SOURCE1 For Input As #1
        TT = Input(LOF(1), 1)
    Close #1
    
    Open fp For Append As #1
        Print #1, TT
    Close #1
Next


SOURCE1 = fp

'W595
ScanPath.lblCount2 = " W595 - PROCESS"
i = ExecCmdWAIT("C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=6c:0e:0d:30:fc:86 -BD_NAME=W595i -DEV_CLASS=00262746" + " """ + SOURCE1 + """")

'K700i
ScanPath.lblCount2 = " K700 - PROCESS"
i = ExecCmdWAIT("C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=00:fc:22:08:04:0e -BD_NAME=K700i -DEV_CLASS=00262738" + " """ + SOURCE1 + """")

'V600i
ScanPath.lblCount2 = " V600 - PROCESS"
i = ExecCmdWAIT("C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=00:12:ee:61:65:7e -BD_NAME=V600i -DEV_CLASS=00262738" + " """ + SOURCE1 + """")

'K800i
ScanPath.lblCount2 = " K800 - PROCESS"
i = ExecCmdWAIT("C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=00:1d:28:2b:90:9c -BD_NAME=K800i -DEV_CLASS=00262746" + " """ + SOURCE1 + """")

'W880i
ScanPath.lblCount2 = " W880 - PROCESS"
i = ExecCmdWAIT("C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=00:1d:28:36:2a:8d -BD_NAME=W80x2D07919180203 -DEV_CLASS=00262746" + " """ + SOURCE1 + """")

    
ScanPath.lblCount6 = "DONE"

End

End Sub

Private Sub Label15_Click()
Call COPY_CDROM_ANYWHERE
End Sub

Function FINDDRIVENAME(AD)

FINDDRIVENAME = ""
Set Fs = CreateObject("Scripting.FileSystemObject")

Set D = Fs.GetDrive(Mid(AD, 1, 1))
'Debug.Print D.VolumeName
FINDDRIVENAME = D.VolumeName


End Function
Function FINDDRIVE(AD)
FINDDRIVE = ""
Set Fs = CreateObject("Scripting.FileSystemObject")
'AD = "ScreenPlay Director"
For RX3 = 67 To 97
Err.Clear
On Error Resume Next ' FOR DEVICE NOT READY
Set D = Fs.GetDrive(Chr(RX3))
'Debug.Print Str(Val(r)) + " --- " + D.VolumeName
    If Err.Number = 0 Then
        If D.ISREADY Then
            XVN = Replace(UCase(D.VolumeName), " ", "")
            XDN = Replace(UCase(AD), " ", "")
            'Debug.Print XVN
            'Debug.Print XDN
            
            If XDN = XVN Then
                FINDDRIVE = Chr(RX3)
                Exit For
            End If
        End If
    End If
Next


End Function

Sub DIR_PARAMETERS_FOR_ARRAY()
VS = Replace(VS, " ", "_")
VD = Replace(VD, " ", "_")

If DIR_TD = "" Then DIR_TD = DIR_TS

i = i + 1
AD = FINDDRIVE(VS)
COPYSET_S(i) = AD + Mid(DIR_TS, InStr(DIR_TS, ":"))
AD = FINDDRIVE(VD)
COPYSET_D(i) = AD + Mid(DIR_TD, InStr(DIR_TD, ":"))
DIR_TD = ""

If Len(COPYSET_S(i)) < 3 Then COPYSET_S(i) = "": COPYSET_D(i) = ""
If Len(COPYSET_D(i)) < 3 Then COPYSET_S(i) = "": COPYSET_D(i) = ""
End Sub


Private Sub Label16_Click()

' ------------------ FIND
'Call COPY_ANYWHERE_PROPER


'WANT COPY THEN DEL SYNC CODE
'Call COPY_ANYWHERE_PROPER

Dim TCN, COUNT
Me.WindowState = 0
Me.Show
DoEvents
BLAST = Now


ScanPath.ListView1.Top = 3700
ScanPath.ListView1.Height = 3300
ScanPath.lblCount7.Top = ScanPath.ListView1.Top - ScanPath.lblCount7.Height
ScanPath.lblCount8.Top = ScanPath.lblCount7.Top - ScanPath.lblCount8.Height
ScanPath.lblCount9.Top = ScanPath.lblCount8.Top - ScanPath.lblCount9.Height
ScanPath.lblCount9.Visible = False

ScanPath.Height = ScanPath.ListView1.Height + ScanPath.ListView1.Top + 500

PBar1.Top = ScanPath.lblCount9.Top
PBar1.Height = ScanPath.lblCount9.Height
PBar1.Width = ScanPath.lblCount9.Width + 20
PBar1.Left = 0


cmdScan_NO_LIST.Visible = False
Text1.Visible = False
Text2.Visible = False
Command13.Visible = False
Label14.Visible = False

ScanPath.lblCount7.ZOrder (0)

Chk_SortOnDate.Visible = False
chkCopyMemory.Visible = False
lblSize(1).Visible = False
chkSubFolders.Visible = False
chkPatternMatching.Visible = False



ReDim COPYSET_S(1000)
ReDim COPYSET_D(1000)
ReDim ARRAY_DELETE_COPY(1000)
ReDim ARRAY_DELETE_ALL_EMPTY_FOLDER(1000)
i = 0




ReDim RTX(2)
RTX(1) = "U DISK"
RTX(2) = "SANDISK"

ReDim RRX(3)
RRX(1) = "E:\RAR"
RRX(2) = "D:\RAR*"
RRX(3) = "D:\# MY DOCS"

For R1 = 1 To UBound(RTX)
    For R2 = 1 To UBound(RRX)
        
        'ONLY FIRST OF DIR TO SANDDISK
        If R1 > 1 And R2 > 1 Then Exit For
        '--<><><>
        '-----------------<><><>
        VS = RTX(R1)
        VD = RRX(R1)
        DIR_TS = ":\"
        '-----------------<><><>
        Call DIR_PARAMETERS_FOR_ARRAY
        '-----------------<><><>
        '--<><><>
    Next
Next




'AUDIO RECORD
'--<><><>

ReDim RTX(3)
RTX(1) = "FUJI"
RTX(2) = "SONY_8GB"
RTX(3) = "WS_750M"
RTX(0) = ""

ReDim RRX(3)
RRX(1) = "M:\0 00 Art\0 MY DEVICES - ART\# FUJI - ALL"
RRX(2) = "M:\0 00 Art\0 MY DEVICES - ART\# SONY\# SONY - ALL"
RRX(3) = "M:\0 00 MEDIA\06 Olympus WS_750M"
RRX(0) = "M:\0 00 MEDIA\07 Olympus VN_8500PC"
RRX(0) = "M:\0 00 MEDIA\05 Olympus VN_6800PC"
RRX(0) = "M:\0 00 MEDIA\02 Olympus VN_2100PC"
RRX(0) = ""


For R1 = 1 To UBound(RTX)
    For R2 = 1 To UBound(RRX)
        '--<><><>
        '-----------------<><><>
        VS = RTX(R1)
        VD = "Elements"
        DIR_TS = ":\"
        DIR_TD = RRX(R2)
        '-----------------<><><>
        Call DIR_PARAMETERS_FOR_ARRAY
        ARRAY_DELETE_COPY(i) = True
        '-----------------<><><>
        '--<><><>
        '--<><><>
        '-----------------<><><>
        VS = RTX(R1)
        VD = "D_DRIVE"
        DIR_TS = ":\"
        DIR_TD = RRX(R2)
        '-----------------<><><>
        Call DIR_PARAMETERS_FOR_ARRAY
        ARRAY_DELETE_COPY(i) = True
        '-----------------<><><>
        '--<><><>
    Next
Next



'AUDIO RECORD
'--<><><>
'-----------------<><><>
VS = "WS_750M"
VD = "Elements"
DIR_TS = ":\"
DIR_TD = ":\0 00 MEDIA\06 Olympus WS_750M"
'-----------------<><><>
Call DIR_PARAMETERS_FOR_ARRAY
ARRAY_DELETE_COPY(i) = True
'-----------------<><><>
'--<><><> ' AND TO D DRIVE
'--<><><> ' AND TO D DRIVE
'-----------------<><><>
VD = "D DRIVE"
'-----------------<><><>
Call DIR_PARAMETERS_FOR_ARRAY
ARRAY_DELETE_COPY(i) = True
'-----------------<><><>



'ONE JOB WORK
'--<><><>
'-----------------<><><>
VS = "My_Book_SOUCE"
VD = "NAS"
DIR_TS = ":\#ACER"
'-----------------<><><>
Call DIR_PARAMETERS_FOR_ARRAY
ARRAY_DELETE_COPY(i) = True
'ARRAY_DELETE_ALL_EMPTY_FOLDER(i) = True
'-----------------<><><>
'--<><><>



'ALL C DRIVE
'--<><><>
'-----------------<><><>
VS = "C_DRIVE"
'-----------------<><><>
VD = "Elements"
DIR_TS = ":\"
DIR_TD = ":\#0-C-DR"
'-----------------<><><>
Call DIR_PARAMETERS_FOR_ARRAY
'-----------------<><><>
'--<><><>

'ALL D DRIVE
'--<><><>
'-----------------<><><>
VS = "D_DRIVE"
VD = "Elements"
DIR_TS = ":\"
DIR_TD = ":\#0-D-DR"
'-----------------<><><>
Call DIR_PARAMETERS_FOR_ARRAY
'-----------------<><><>
'--<><><>

'------- FROM D DRIVE TO BIGG DRIVE

ReDim RTX(5)
RTX(1) = "Elements"
RTX(2) = "NAS"
RTX(3) = "MY_BOOK_MC"
RTX(4) = "WD_PASS_1TB"
RTX(5) = "WD_PASS_4GB"

ReDim ITX(500)

L2 = L2 + 1: ITX(L2) = "# MAINTANCE TOOL'S"
L2 = L2 + 1: ITX(L2) = "# MY DOCS"
L2 = L2 + 1: ITX(L2) = "#0 INSTALLATIONS"
L2 = L2 + 1: ITX(L2) = "0 00 Art"
L2 = L2 + 1: ITX(L2) = "0 00 Art Loggers"
L2 = L2 + 1: ITX(L2) = "0 00 HTTrack"
L2 = L2 + 1: ITX(L2) = "0 00 Media"
L2 = L2 + 1: ITX(L2) = "Email Outlook 01"
L2 = L2 + 1: ITX(L2) = "Email Outlook 02"
L2 = L2 + 1: ITX(L2) = "Email_OE"
L2 = L2 + 1: ITX(L2) = "Feed -Demon3"
L2 = L2 + 1: ITX(L2) = "My Webs"
L2 = L2 + 1: ITX(L2) = "TEMP"
L2 = L2 + 1: ITX(L2) = "VB6"
L2 = L2 + 1: ITX(L2) = "VB6 Icons"
L2 = L2 + 1: ITX(L2) = "Wave's"
L2 = L2 + 1: ITX(L2) = "WINDOS98"
L2 = L2 + 1: ITX(L2) = "z#U"

For R = 1 To L2
    ITX(R) = "D:\" + ITX(R)
Next


For R1 = 1 To UBound(RTX)
    For R2 = 1 To L2
        '--<><><>
        '-----------------<><><>
        VS = "D_DRIVE"
        VD = RTX(R1)
        DIR_TS = ITX(R2)
        '-----------------<><><>
        Call DIR_PARAMETERS_FOR_ARRAY
        '-----------------<><><>
        '--<><><>
    Next
Next



ReDim ITX(5)
ITX(1) = "Elements"
ITX(2) = "NAS"
ITX(3) = "MY_BOOK_MC"
ITX(4) = "WD_PASS_1TB"
ITX(5) = "WD_PASS_4GB"


For RX = 1 To UBound(ITX)
'--<><><>
'-----------------<><><>
VS = "E_DRIVE"
'-----------------<><><>
VD = ITX(RX)
DIR_TS = ":\"
DIR_TD = ":\#0-E-DR"
'-----------------<><><>
Call DIR_PARAMETERS_FOR_ARRAY
'-----------------<><><>
'--<><><>
Next





'FROM ELEMENTS - TO
'NAS - ELEMENTS AND PASS 1TB- MOTHER

ReDim ITX(14)
ITX(1) = "#0 INSTALLATIONS"
ITX(2) = "0 00 Art"
ITX(3) = "0 00 Art Loggers"
ITX(4) = "0 00 ART WORK IN BOUND"
ITX(5) = "0 00 HTTrack"
ITX(6) = "0 00 Music ---"
ITX(7) = "0 00 Video"
ITX(8) = "My Webs"
ITX(9) = "#0 MAINTAINANCE TOOLS"
ITX(10) = "VB6 Icons"
ITX(11) = "Wave 's"
ITX(12) = "0 00 Video\01 Beavis & Butthead"
ITX(13) = "0 00 Video\00 Vid Films\FILM MP3's"
ITX(14) = "0 00 Video\00 Vid XXX\MP3"


VS = "Elements"
AD = FINDDRIVE(VS)

For R = 1 To UBound(ITX)
    ITX(R) = AD + ":\" + ITX(R)
Next


'FROM ELEMENTS - TO
'NAS - ELEMENTS AND PASS 1TB- MOTHER
For RX = 1 To UBound(ITX)
    
    '--<><><>
    '-----------------<><><>
    VD = "NAS"
    DIR_TS = ":\" + ITX(RX)
    '-----------------<><><>
    Call DIR_PARAMETERS_FOR_ARRAY
    '-----------------<><><>
    '--<><><>
    '--<><><>
    '-----------------<><><>
    VD = "Elements_MO"
    DIR_TS = ":\" + ITX(RX)
    '-----------------<><><>
    Call DIR_PARAMETERS_FOR_ARRAY
    '-----------------<><><>
    '--<><><>
    '--<><><>
    '-----------------<><><>
    VD = "WD_PASS_1TB"
    DIR_TS = ":\" + ITX(RX)
    '-----------------<><><>
    Call DIR_PARAMETERS_FOR_ARRAY
    '-----------------<><><>
    '--<><><>
Next



'FROM ELEMENTS TO
'ALL MEDIA CENTER HHD
ReDim ITX(8)
ITX(1) = "0 00 Art"
ITX(2) = "0 00 Art Loggers"
ITX(3) = "0 00 ART WORK IN BOUND"
ITX(4) = "0 00 HTTrack"
ITX(5) = "0 00 Music ---"
ITX(6) = "0 00 Video"
ITX(7) = "My Webs"
ITX(8) = "Wave 's"

VS = "Elements"
AD = FINDDRIVE(VS)
For R = 1 To UBound(ITX)
    ITX(R) = AD + ":\" + ITX(R)
Next

For RX = 1 To UBound(ITX)
    '--<><><>
    '-----------------<><><>
    VD = "ScreenPlay_Director"
    DIR_TS = ":\" + ITX(RX)
    '-----------------<><><>
    Call DIR_PARAMETERS_FOR_ARRAY
    '-----------------<><><>
    '--<><><>
    '--<><><>
    '-----------------<><><>
    VD = "MY_BOOK_MC"
    DIR_TS = ":\" + ITX(RX)
    '-----------------<><><>
    Call DIR_PARAMETERS_FOR_ARRAY
    '-----------------<><><>
    '--<><><>
Next

ReDim Preserve COPYSET_S(i)
ReDim Preserve COPYSET_D(i)
ReDim Preserve ARRAY_DELETE_COPY(1000)
ReDim Preserve ARRAY_DELETE_ALL_EMPTY_FOLDER(1000)



'--------------------------------------------
'BEGIN WORK ------------------------
For R = 1 To UBound(COPYSET_S)
    'If r = 4 Then Stop
    
    COPYSET_S(R) = Replace(COPYSET_S(R), "\\", "\")
    COPYSET_D(R) = Replace(COPYSET_D(R), "\\", "\")
    
    If Fs.DRIVEEXISTS(Mid(COPYSET_S(R), 1, 2)) = False Then
        COPYSET_S(R) = ""
        COPYSET_D(R) = ""
    End If
    If Fs.DRIVEEXISTS(Mid(COPYSET_D(R), 1, 2)) = False Then
        COPYSET_D(R) = ""
        COPYSET_S(R) = ""
    End If

    If COPYSET_S(R) <> "" Then
        ScanPath.ListView1.ListItems.Clear
        DoEvents
        DUPESOURCE = False
        ScanPath.TxtPath = COPYSET_S(R)
        If OSOURCE = ScanPath.TxtPath Then DUPESOURCE = True
        OSOURCE = ScanPath.TxtPath
        
        ScanPath.txtPath02 = COPYSET_D(R)
        TTN = Trim(Str(UBound(COPYSET_S))) + "+ # "
        ' DUPESOURCE = True
        COUNT = COUNT + 1
        AD1 = FINDDRIVENAME(COPYSET_S(R))
        AD2 = FINDDRIVENAME(COPYSET_D(R))
        TCN = Trim(Str(R)) + " of " + TTN
        
        COPYSET_COUNT = R
        
        lblCount5 = TCN + "   --- " + ScanPath.TxtPath + " - To - " + ScanPath.txtPath02
        lblCount8 = AD1 + " -To- " + AD2 + " ** " + lblCount5
        ScanPath.lblCount7 = lblCount5
        ScanPath.lblCount6 = "** WAIT FILE SYS *** " + AD1
        'If r >= 127 Or r = 1 Then Stop
        RCOUNT_CHUNK_MARK = 0
        BLAST = Now
        
        Label21.Caption = " " + COPYSET_S(R) + " ---- ARRIVE AT SOURCE DIRECTORY -- "
        Label20.Caption = " " + COPYSET_D(R) + " ---- ARRIVE AT DESTI DIRECTORY -- "
        Label20.Refresh
        Label21.Refresh
        Label20.Left = lblCount6.Left
        Label21.Left = lblCount6.Left
                
        DoEvents
        'If r >= 131 Or r < 3 Then
        Call COPY_ANYWHERE_PROPER
        'End If
        If EXITVAR = True Then Exit For
    End If
Next
        ScanPath.lblCount6 = "** /\/\/\ XXXxx WORK DONE xxXXX /\/\/\ *** "
End Sub

Private Sub Label19_Click()

Exit Sub

Me.WindowState = 0
Me.Show
DoEvents


'ScanPath.TxtPath = "N:\"
'ScanPath.txtPath02 = "M:\"
'Call MOVE_ANYWHERE_PROPER
'MsgBox "READY NEXT "
'Stop



ScanPath.TxtPath = "G:\0 00 Art\0 MY DEVICES - VIDEO"
ScanPath.txtPath02 = "M:\0 00 Art\0 MY DEVICES - VIDEO"
Call MOVE_ANYWHERE_PROPER
MsgBox "READY NEXT "
Stop




ScanPath.TxtPath = "J:\"
ScanPath.txtPath02 = "M:\"
Call MOVE_ANYWHERE_PROPER
MsgBox "READY NEXT "
Stop



'ScanPath.TxtPath = "D:\00 Pen Drives MOBILES\"
'ScanPath.txtPath02 = "M:\"
'Call MOVE_ANYWHERE_PROPER


End Sub

Private Sub Label20_Click()
    
    'DESTIFOLDERPATH = D1
    'SOURCEFOLDERPATH = A1
    
    'Shell "EXPLORER /n, /e, /select," + DESTIFOLDERPATH
    Shell "EXPLORER /n, /e," + DESTIFOLDERPATH

End Sub

Private Sub Label21_Click()
    
    'DESTIFOLDERPATH = D1
    'SOURCEFOLDERPATH = A1
    
    Shell "EXPLORER /n, /e," + SOURCEFOLDERPATH

End Sub

Private Sub Label22_Click()
Call EXECUTE_ALL_LINK_FILE_FOLDER
End Sub

Private Sub lblCount1_Change()
'Stop
'If OCNT > Val(lblCount1) Then Stop
'OCNT = Val(lblCount1)
End Sub

Private Sub MNU_LIST_FILES_TO_CLIP_httrack_Click()
Call Command7_Click
End Sub

Private Sub MNU_LOGG_COPY_PROPER_Click()
     Shell "NOTEPAD D:\TEMP\LOGG MULTI COPY_SMALL.TXT", vbMaximizedFocus
End Sub

Private Sub Mnu_MatchBungDateDesti_Click()


ScanPath.ListView1.ListItems.Clear

ScanPath.chkSubFolders = vbChecked


ScanPath.TxtPath = "F:\##ACER-D2"

Call ScanPath.cmdScan_Click


'destipath1 = "\\55-88-happy\MY BOOK (M)\##ASUS-D"
destipath1 = "\\55-88-happy\D"


'On Error Resume Next
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " Process Files"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    
    D1 = destipath1 + Mid(A1, Len(ScanPath.TxtPath) + 1)
    'D2 = destipath2 + Mid(A1, Len(ScanPath.txtPath) + 1)
    'D3 = "D:" + Mid(A1, Len(ScanPath.txtPath) + 1)
    
    
    
    If Fs.FileExists(A1 + B1) = True And Fs.FileExists(D1 + B1) = True Then
        
        Set F = Fs.getfile(D1 + B1)
        adate1 = F.datelastmodified
        If adate1 < DateSerial(2010, 2, 1) Then
        
        
        Err.Clear
        
        'DEL FROM OUR SWAP DRIVE
        Fs.DeleteFile A1 + B1, True
            
        TTD2 = TTD2 + 1
        ScanPath.lblCount3 = Trim(Str(TTD2)) + " Del-ed Files"
            
        End If
    End If
    
    'On Error Resume Next
    
    If Err.Number > 0 Then
        AG = AG + 1
    End If
    ScanPath.lblCount4 = Trim(Str(AG)) + " Del Errors"

Next


ScanPath.lblCount6 = "** DONE **"
ScanPath.WindowState = 0
    




End Sub

Private Sub Mnu_MatchBungDel_Click()

Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32

ScanPath.ListView1.ListItems.Clear

ScanPath.chkSubFolders = vbChecked


ScanPath.TxtPath = "F:\##ACER-D2"

Call ScanPath.cmdScan_Click


destipath1 = "\\55-88-happy\MY BOOK (M)\##ASUS-D"
destipath2 = "\\55-88-happy\D"


'On Error Resume Next
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
    If ScanPath.ListView1.ListItems.COUNT = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " Process Files"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    
    D1 = destipath1 + Mid(A1, Len(ScanPath.TxtPath) + 1)
    d2 = destipath2 + Mid(A1, Len(ScanPath.TxtPath) + 1)
    D3 = "D:" + Mid(A1, Len(ScanPath.TxtPath) + 1)
    
    
    
    If Fs.FileExists(A1 + B1) = True And Fs.FileExists(D1 + B1) = True Then
        
        If m_CRC.CalculateFile(A1 + B1) = m_CRC.CalculateFile(D1 + B1) Then
        
        Err.Clear
        'COPY TO OUR  -D-
        'Debug.Print A1
        'Debug.Print D3
        
        If Fs.FolderExists(D3) = False Then
            'MkDir Mid(D3, 1, InStrRev(D3, "\", Len(D3) - 1))
            MkDir D3
            Stop
            Err.Clear
        End If
        
        Fs.COPYFILE A1 + B1, D3 + B1, True
        
        'Err.Clear
        'DEL FROM OUR SWAP DRIVE
        If Err.Number = 0 Then Fs.DeleteFile A1 + B1, True
        
        
        On Error Resume Next
        'COPY TO DESTINATION'S -D-
        If Err.Number = 0 Then Fs.COPYFILE A1 + B1, d2 + B1
        
        'DEL FROM DESTINATION SWAP DRIVE
        If Err.Number = 0 Then Fs.DeleteFile D1 + B1, True
        On Error GoTo 0
        
            
        End If
    End If
    
    'On Error Resume Next
    
    If Err.Number > 0 Then
        AG = AG + 1
    End If
    ScanPath.lblCount3 = Trim(Str(AG)) + " Del Errors"

Next


ScanPath.lblCount6 = "** DONE **"
ScanPath.WindowState = 0
    



End Sub


Private Sub MOVE_IN_DATE_FOLDERS_M2CARD_Click()
'FIXED CODE DONT CHANGE


'If Command$ <> "" Then Me.WindowState = 0

'=NEWER THAN THIS DATE

'DTPicker1(0).CheckBox = True

'SOMETIMES WE USE

'cboDate.ListIndex = 0 ' MODIFIED
'DTPicker1(0) = Int(Now) - 15


ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask.Text = "*.MPG;*.3GP"
ScanPath.chkSubFolders = vbChecked

'ScanPath.txtPath = "D:\# MY DOCS\# 01 My Documents\00 Blue Tooth Exchange Folder\"
''Call ScanPath.cmdScan_Click

ScanPath.TxtPath = "D:\# W880i\VIDEO\"
Call ScanPath.cmdScan_Click

ScanPath.TxtPath = "M:\# W880i\VIDEO\"
Call ScanPath.cmdScan_Click




'## OF THIS NULL THEN DIR RELATIVE TO DIR OF PATH FILE
''## WELL TO FIRST DIR PATH

DXDIR = ScanPath.TxtPath

Set F = Fs.GetDrive("M")
If F.ISREADY = True Then
    Mid(DXDIR, 1, 1) = "M"
Else
    Mid(DXDIR, 1, 1) = "D"
End If

DTPicker1(0) = vbNullString





'Exit Sub



'    'MOST HARD CODED
'    ScanPath.Text1 = "    Loop" + vbCrLf + "    " + vbCrLf + "    " + "FindClose lResult"
'    ScanPath.Text2 = "        FindClose lResult" + vbCrLf + "    " + vbCrLf + "    " + "Loop"


'ScanPath.Text2 = """>"
'ScanPath.Text1 = "http:"


ScanPath.Show

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"
    
'B1 = ScanPath.ListView1.ListItems.Item(1)
'RXRIME = Format(DateValue(Mid(B1, 8, 10)), "YYYY-MM-MMM")
'FN = ScanPath.txtPath.Text + "HTTP's From Email's -- " + RXRIME + ".txt"

'If Dir(FN) <> "" Then Kill FN
'UU = FreeFile
'Open FN For Append As #UU

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
'    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " DEALT WITH FILES"
    
    DoEvents
    
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    TT = TT + "-*/" + B1 + vbCrLf
    
    Set F = Fs.getfile(A1 + B1)
    adate1 = F.datelastmodified
    
    
    
    L2 = InStr(OO, Chr(1) + Chr(0) + Chr(0) + Chr(0) + Chr(32))
    L2 = InStr(OO, Trim(Str(Year(Now))))
    If L2 = 0 Then
    L21 = InStr(OO, Trim(Str(Year(Now) - 1)))
    L23 = InStr(OO, Trim(Str(Year(Now) + 1)))
    If L21 > 0 Then L2 = L21
    If IL2 > 0 Then L2 = IL2
    If L23 > 0 Then L2 = IL2
    End If
    
    L2RIME = ""
    If L2 > 0 Then
        For R = 1 To 19
            L2RIME = L2RIME + Chr(Asc(Mid(OO, L2 + R - 1, 1)))
            '&HFA
        Next
    End If
    
    L2RIME2 = ""
    If L2RIME <> "" Then
        Mid(L2RIME, 5, 1) = "/"
        Mid(L2RIME, 8, 1) = "/"
        L2RIME2 = DateValue(L2RIME) + TimeValue(L2RIME)
    End If
    
    If L2RIME2 <> "" Then
        If L2RIME2 > DateSerial(2005, 1, 1) Then
            adate1 = L2RIME2
        Else
            Stop
        End If
    End If
    
    YY = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD")
    
'    If InStr(B1, "DSC") > 0 Then
'        yy2 = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD") + " " + Format(adate1, "hh-Mm-ss") + " - " + cam + " - " + Mid(B1, InStr(B1, "DSC"))
'    Else
'        yy2 = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD") + " " + Format(adate1, "hh-Mm-ss") + " - " + cam + " - " + B1
'    End If
    
    'yy2 = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD") + " " + Format(adate1, "hh-Mm-ss") + " - " + B1
    'yy2 = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD") + " - " + B1
    
    yy2 = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD") + " " + Format(adate1, "hh-Mm-ss") + " - " + Mid(B1, InStrRev(B1, "MOV"))
    
    If InStrRev(B1, "MOV") = 0 Then
        Me.WindowState = 0
        MsgBox "WE WANTED O KNOW IF ANY FILE NOT AND --.MOV -- FILE"
        Stop
    End If
    
    'If InStr(B1, "JFIF -") > 0 Then
    '    yy2 = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD") + " " + Format(adate1, "hh-Mm-ss") + " - " + Mid(B1, InStrRev(B1, "JFIF -") + 7)
    'End If
    
    'MOVE INTO FOLDER
    
    
        
    
    If Fs.FolderExists(DXDIR + YY) = False Then MkDir DXDIR + YY
    If Fs.FileExists(DXDIR + YY + "\" + yy2) = False Then
        Call ShellFileMove(A1 + B1, DXDIR + YY + "\" + yy2)
        'Name A1 + B1 As DXDIR + yy + "\" + yy2
    Else
        Debug.Print "File Exists Already SOURCE ="
        Debug.Print A1 + B1
    End If
    
    'MOVE WHERE IS
    If Fs.FileExists(A1$ + yy2) = False Then
        'F.Move A1$ + YY2
    End If
    
    
    
    'F.Move A1 + YY + "\" + YY + " - " + B1
    
    
    
    
    
    
Next


End Sub



Private Sub MNU_VB_Click()

Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
End
End Sub

Private Sub MOVE_IN_DATE_FOLDERS_JPGS_Click()

'PLYABLE CODE DO WHAT WANT WITH

'ScanPath.txtPath = "K:\00 HTTrack\Techno BeatPlexity.com MusicGateWay\http___www.beatplexity.com_musicgateway\www.beatplexity.com\musicgateway\"
'ScanPath.txtPath = "K:\00 HTTrack\Techno PoMac.Net Swarm\Http___pomac.netswarm.net_music.html_music\pomac.netswarm.net\music"

'ScanPath.txtPath = "X:\00 Blue Tooth Exchange Folder\IMAGES 2009\"
'ScanPath.txtPath = "M:\0 00 Art\00 My Pictures\01 All Phone\00 Mental\TEST"

If Command$ <> "" Then Me.WindowState = 0

'=NEWER THAN THIS DATE
cboDate.ListIndex = 0 ' MODIFIED
'DTPicker1(0).CheckBox = True
DTPicker1(0) = Int(Now) - 15


ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask.Text = "*.JPG;*.MPG;*.3GP"
ScanPath.chkSubFolders = vbChecked

ScanPath.TxtPath = "D:\# MY DOCS\# 01 My Documents\00 Blue Tooth Exchange Folder\"
Call ScanPath.cmdScan_Click

'## OF THIS NULL THEN DIR RELATIVE TO DIR OF PATH FILE
''## WELL TO FIRST DIR PATH
DXDIR = ScanPath.TxtPath
DTPicker1(0) = vbNullString

ScanPath.TxtPath = "M:\# W880i\VIDEO\"
Call ScanPath.cmdScan_Click



'Exit Sub



'    'MOST HARD CODED
'    ScanPath.Text1 = "    Loop" + vbCrLf + "    " + vbCrLf + "    " + "FindClose lResult"
'    ScanPath.Text2 = "        FindClose lResult" + vbCrLf + "    " + vbCrLf + "    " + "Loop"


'ScanPath.Text2 = """>"
'ScanPath.Text1 = "http:"


ScanPath.Show

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.COUNT)) + " TOTAL FILES"
    
'B1 = ScanPath.ListView1.ListItems.Item(1)
'RXRIME = Format(DateValue(Mid(B1, 8, 10)), "YYYY-MM-MMM")
'FN = ScanPath.txtPath.Text + "HTTP's From Email's -- " + RXRIME + ".txt"

'If Dir(FN) <> "" Then Kill FN
'UU = FreeFile
'Open FN For Append As #UU

For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    
'    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " DEALT WITH FILES"
    
    DoEvents
    
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    TT = TT + "-*/" + B1 + vbCrLf
    
    Set F = Fs.getfile(A1 + B1)
    adate1 = F.datelastmodified
    
    'C1 = YY + B1
    Open A1 + B1 For Binary As #1
    OO = UCase(Input(800, 1))
    Close #1
    cam = ""
    
    'Debug.Print OO
    'MsgBox OO
    
    If InStr(OO, "K800I") > 0 Then cam = "K800I"
    If InStr(OO, "W880I") > 0 Then cam = "W880I"
    If InStr(OO, "W595") > 0 Then cam = "W595"
    
    If cam = "" Then
        cam = "JFIF"
    End If
    
    L2 = InStr(OO, Chr(1) + Chr(0) + Chr(0) + Chr(0) + Chr(32))
    L2 = InStr(OO, Trim(Str(Year(Now))))
    If L2 = 0 Then
    L21 = InStr(OO, Trim(Str(Year(Now) - 1)))
    L23 = InStr(OO, Trim(Str(Year(Now) + 1)))
    If L21 > 0 Then L2 = L21
    If IL2 > 0 Then L2 = IL2
    If L23 > 0 Then L2 = IL2
    End If
    
    L2RIME = ""
    If L2 > 0 Then
        For R = 1 To 19
            L2RIME = L2RIME + Chr(Asc(Mid(OO, L2 + R - 1, 1)))
            '&HFA
        Next
    End If
    
    L2RIME2 = ""
    If L2RIME <> "" Then
        Mid(L2RIME, 5, 1) = "/"
        Mid(L2RIME, 8, 1) = "/"
        L2RIME2 = DateValue(L2RIME) + TimeValue(L2RIME)
    End If
    
    If L2RIME2 <> "" Then
        If L2RIME2 > DateSerial(2005, 1, 1) Then
            adate1 = L2RIME2
        Else
            Stop
        End If
    End If
    YY = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD") + " - " + cam
    
    If InStr(B1, "DSC") > 0 Then
        yy2 = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD") + " " + Format(adate1, "hh-Mm-ss") + " - " + cam + " - " + Mid(B1, InStr(B1, "DSC"))
    Else
        yy2 = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD") + " " + Format(adate1, "hh-Mm-ss") + " - " + cam + " - " + B1
    End If
    
    If InStr(B1, "JFIF -") > 0 Then
        yy2 = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD") + " " + Format(adate1, "hh-Mm-ss") + " - " + Mid(B1, InStrRev(B1, "JFIF -"))
    End If
    
    
    'MOVE INTO FOLDER
    
    If DXDIR = "" Then
        If Fs.FolderExists(ScanPath.TxtPath + YY) = False Then MkDir ScanPath.TxtPath + YY
        If Fs.FileExists(ScanPath.TxtPath + YY + "\" + yy2) = False Then
            'F.Move ScanPath.txtPath + YY + "\" + YY2
            Call ShellFileMove(A1 + B1, ScanPath.TxtPath + YY + "\" + yy2)
        End If
    Else
        If Fs.FolderExists(DXDIR + YY) = False Then MkDir DXDIR + YY
        If Fs.FileExists(DXDIR + YY + "\" + yy2) = False Then
            Call ShellFileMove(A1 + B1, DXDIR + YY + "\" + yy2)
        End If
    End If
        'If DXDIR <> "" Then F.Move DXDIR + yy + "\" + YY2
    
    
    'MOVE WHERE IS
    If Fs.FileExists(A1$ + yy2) = False Then
        'F.Move A1$ + YY2
    End If
    
    
    
    'F.Move A1 + YY + "\" + YY + " - " + B1
    
    
    
    
    
    
Next


End Sub



Private Sub Command9_Click()
    Call MOVE_ANY_WHERE
End Sub

Private Sub Form_Load()

End

If App.PrevInstance = True Then End
Debug.Print App.Path + "\" + App.EXEName + ".vbp"
If IsIDE = False Then
    'Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE "" /r """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
    'End
End If

Me.Caption = "ScanPath 2.0 - Sort  Anything - "

'x = FindWindow(vbNullString, "ScanPath 2.0 - Sort  Anything - Sort_AnyThing - MULTI USE MENU")
'ISANOTHERVBRUN = FindWindow("wndclass_desked_gsk", vbNullString)

'If x > 0 And ISANOTHERVBRUN = 0 Then End


'Shell "C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=00:1d:28:2b:90:9c -BD_NAME=K80x2D07919180203 -DEV_CLASS=00262746" + " """ + Source1 + """"
    
'Call ShellFileDelete("C:\C", True)
Dim i As Long
    
On Error Resume Next
For i = 1 To 100
    If InStr(GetSpecialfolder(i) + "--", "\01 Start Menu\Programs--") > 0 Then
    Debug.Print i, GetSpecialfolder(i)
    For R = 1 To 26
        MkDir GetSpecialfolder(i) + "\" + Chr(R + 64)
        MkDir GetSpecialfolder(i) + "\MICROSOFT"
        MkDir GetSpecialfolder(i) + "\# GAMES"
        MkDir GetSpecialfolder(i) + "\01"
    Next
    End If
    If InStr(GetSpecialfolder(i) + "--", "\01 Start Menu Common\Programs--") > 0 Then
    Debug.Print i, GetSpecialfolder(i)
    For R = 1 To 26
        MkDir GetSpecialfolder(i) + "\" + Chr(R + 64)
        MkDir GetSpecialfolder(i) + "\MICROSOFT"
        MkDir GetSpecialfolder(i) + "\# GAMES"
        MkDir GetSpecialfolder(i) + "\01"
    Next
    End If
    '
Next


    'Stop
'End

Me.BackColor = &HC0C0C0
    
Set Fs = CreateObject("Scripting.FileSystemObject")
    
'--------
'If you use ChkMem Size Files then Know G1$ Alternate Data uses same column(2) for its data
'-
'Seem to have a problem with some program an not other when use chkmem file size and date
'Work Around done that
'-
'The Date checker box does scan for time if the var contains time
'---------
    
ScanPath.Caption = ScanPath.Caption + App.EXEName
'scanpath.Caption = frmMain.Caption + App.EXEName
  
ListView1.Height = 1300
  
x = 1
y = 1
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True Then
        If Control.Width + Control.Left > x Then x = Control.Width + Control.Left
        If Control.Height + Control.Top > y Then y = Control.Height + Control.Top
        If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
    End If
Next
On Error GoTo 0


ScanPath.Width = x + 110
If mnu = 1 Then
    pluso = 720
Else
    pluso = 450
End If

'pluso = 0

'Me.Top = 100
'ScanPath.Height = y + pluso
    
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = ((Screen.Height - Me.Height) / 2) + 500
    
    
    Dim lCount As Long
    
    With cboMask
        .AddItem "*.*"
        .AddItem "*.dll;*.exe;*.ocx"
        .AddItem "*.doc;*.mdb;*.xls"
        .AddItem "*.bmp;*.gif;*.jpg;*.tif"
        .AddItem "*.bas;*.cls;*.ctl;*.frm;*.vbp"
        .ListIndex = 0
    End With
    
    DTPicker1(0).Value = DateSerial(Year(Now), Month(Now) - 3, Day(Now))
    DTPicker1(1).Value = Now
    
    DTPicker1(0).Value = Empty
    DTPicker1(1).Value = Empty
    
    With cboDate
        .AddItem "Modified"
        .AddItem "Created"
        .AddItem "Last Accessed"
        .ListIndex = 0
    End With
    
    With cboSize
        .AddItem "Any Size"
        .AddItem "Less Than"
        .AddItem "Greater Than"
        .AddItem "Between"
        .ListIndex = 0
    End With
        
    For lCount = 0 To 1
        With cboSizeType(lCount)
            .AddItem "Bytes"
            .AddItem "Kilobytes"
            .AddItem "Megabytes"
            .ListIndex = 0 ' set it to bytes here
        End With
    Next lCount
    
    With ListView1
        .ColumnHeaders.Add , "FILE", "File", 3000, lvwColumnLeft
        .ColumnHeaders.Add , "PATH", "Path", 4000, lvwColumnLeft
        .ColumnHeaders.Add , "DOS-8.3", "Dos Alternate", 2400, lvwColumnLeft
        .ColumnHeaders.Add , "SIZE", "Size", 1500, lvwColumnRight
        .ColumnHeaders.Add , "DATE", "Date", 3000, lvwColumnLeft
        .ColumnHeaders.Add , "CRC", "CRC", 1700, lvwColumnLeft
        
        
        .View = lvwReport
    End With

    Call chkCopyMemory_Click

ScanPath.Top = 500
ScanPath.ListView1.Top = Command9.Top + Command9.Height
ScanPath.ListView1.Height = 3300
ScanPath.lblCount7.Top = ScanPath.ListView1.Top - ScanPath.lblCount7.Height
ScanPath.lblCount8.Top = ScanPath.lblCount7.Top - ScanPath.lblCount8.Height
ScanPath.lblCount9.Top = ScanPath.lblCount8.Top - ScanPath.lblCount9.Height
ScanPath.lblCount7.Visible = False
ScanPath.lblCount8.Visible = False
ScanPath.lblCount9.Visible = False
ScanPath.lblCount10.Visible = False
ScanPath.PBar1.Visible = False

ScanPath.Height = ScanPath.ListView1.Height + ScanPath.ListView1.Top + 500



'Call DelJunkFrontPage

'ScanPath.Show

'Call INetContent

'Call Bangers

'txtPath.Text = "D:\MY WEBS\MatthewLan.Com Web\"

'Call CRCDUPE

'------------------------------
'center screen
'Me.Top = Screen.Height / 2 - Me.Height / 2
'Me.Left = Screen.Width / 2 - Me.Width / 2

'Me.Top = 400

Me.Show
DoEvents
chkCopyMemory.Value = vbUnchecked





'
'
'Copy
'Call Label16_Click
'
'Move
'Call Label19_Click
'
'Stop
'Exit Sub
'End

TxtPath = Clipboard.GetText
txtPath02 = Clipboard.GetText

ListView1.Width = Me.Width - 50

'
'CRC DUPE
'Label5_Click
'Command1_Click
'

'
''Shell "E:\01 VB Shell Folders\00 Shell VBasic 6 Loader\Sort_AnyThing - MULTI USE MENU.exe.lnk MOVE_IN_DATE_FOLDERS_JPGS", vbMinimizedNoFocus
'If Command$ = "MOVE_IN_DATE_FOLDERS_M2CARD" Then 'Or IsIDE = True Then
'    'MsgBox "HELLO"
'    ScanPath.Visible = False
'    Call MOVE_IN_DATE_FOLDERS_M2CARD_Click
'    End
'End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
End
'Cancel = True
End Sub

Private Sub List1_Click()

tpath3$ = "M:\04 Music ---\Del\"

B1$ = Mid$(List1.List(List1.ListIndex), 14)
C1$ = B1$ + "--Todel"

'tpath1$ = "M:\04 Music ---\"

C1$ = Mid$(B1$, InStrRev(B1$, "\") + 1)
List1.RemoveItem (List1.ListIndex)
'Name b1$ As c1$

'fs.copyFile b1$, tpath3$ + c1$
If Dir$(tpath3$ + C1$) <> "" Then Kill tpath3$ + C1$
Fs.MOVEFILE B1$, tpath3$ + C1$



End Sub

Private Sub Label18_Click()

Call COPY_ANYWHERE_ONE_FOLDER
Exit Sub


Dog = True

If cmdScan.Caption = "Stop" Then cmdScan_Click
Label17 = "Stopped before Complete"
End Sub


Private Sub Label5_Click()

    Label5.BackColor = Label6.BackColor
    Label5 = "CALL CRC DUPE GO MENTAL CRAZY"

End Sub


Private Sub SP_DirMatch(Directory As String, DDirectory As String, Path As String)
    '####################################################################################################
    'This Event fires for each Folder found
    '####################################################################################################
    
    
    
    Dim LV As ListItem

    With ListView1
    Set LV = .ListItems.Add(, , "")
        LV.SubItems(1) = Path + Directory
    End With
    'OPATH = Path

    TL = ScanPath.ListView1.ListItems.COUNT
    If TL < 1 Then Exit Sub
    ScanPath.ListView1.ListItems(TL).EnsureVisible
Exit Sub



'--------------



'sort of PATH then FILE
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 0
    ListView1.Sorted = True
    ListView1.SortKey = 1
    ListView1.Sorted = True
    ListView1.Sorted = False
    
    
For WE = 1 To ScanPath.ListView1.ListItems.COUNT
    A1$ = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(WE)
    YY$ = A1$ + B1$
    Print #1, YY$
Next


ScanPath.ListView1.ListItems.Clear


End Sub



Private Function FormatFileDate(CT As FILETIME) As String
    Const SHORT_DATE = "Short Date"
    
    Dim ST As SYSTEMTIME
    Dim ds1 As Single
    Dim ds2 As Single
       
    If FileTimeToSystemTime(CT, ST) Then
          ds1 = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
          ds2 = TimeSerial(ST.wHour, ST.wMinute, ST.wSecond)
          
          FormatFileDate = Format$(ds1 + ds2, "YYYY-MM-DD HH-MM-SS")
    End If
End Function

Private Sub Timer1_Timer()
'End
End Sub














Sub Jeepers()


On Local Error GoTo jeep
Drived2$ = Mid$(TxtPath.Text, 1, 3)
MkDir Drived2$ + "Temp\Anything"
v1$ = Mid$(TxtPath.Text, InStrRev(TxtPath.Text, "\"))

MkDir Drived2$ + "Temp\Anything" + v1$
On Local Error GoTo 0

List1.AddItem "Stage 1 of 2 : Make Dir's And Move Files to Temp"
List1.AddItem "------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For WE = 1 To ListView1.ListItems.COUNT
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(TxtPath.Text)
    C1$ = Drived2$ + "Temp\Anything" + v1$ + "\" + Mid$(A1$, ets1 + 2)

    etd1 = 13
    ets2 = 1
    Do
        'jc1$ = Mid$(c1$, etd1 + ets2-1)
        ets2 = InStr(etd1 + ets2, C1$, "\")
        etd1 = 1
        If ets2 = 0 Then Exit Do

        d2$ = Mid$(C1$, 1, ets2)

        If InStr(f1$, d2$) = 0 Then
            On Local Error GoTo jeep
            MkDir d2$
            If errs2 <> 75 And errs2 > 0 Then
                MsgBox ("Error Making Temp Directory")
                Stop
            End If
            On Local Error GoTo 0
        End If
    Loop Until ets2 = 0

    f1$ = d2$

    errs2 = 0
    On Local Error GoTo jeep
    Fs.MOVEFILE A1$ + B1$, C1$ + B1$
    On Local Error GoTo 0

    If errs2 <> 0 Then
        MsgBox ("Error Moving File to Temp")
        End
    End If

    List1.AddItem Format$(WE, "000 ") + "Move > " + A1$ + B1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'Fs.DeleteFolder comrad$ + "\" + WFD$

Next

Err.Clear
On Local Error Resume Next
Fs.DeleteFOLDER TxtPath.Text, True
If Err.Number > 0 Then Call HardDel
On Local Error GoTo 0


List1.AddItem "----------------------------------------"
List1.AddItem "Stage 2 of 2 : Move Files Back From Temp"
List1.AddItem "----------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

v2$ = Mid$(TxtPath.Text, 1, InStrRev(TxtPath.Text, "\"))

Err.Clear
On Local Error Resume Next
Fs.Movefolder Drived2$ + "Temp\Anything\*", v2$
If Err.Number > 0 Then Call HardMove
On Local Error GoTo 0

List1.AddItem "Stage 2 of 2 : Complete --------------"
List1.AddItem "--------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

'Timer1.Enabled = True

Exit Sub
jeep:
errs2 = Err.Number
errs3$ = Err.Description
Resume Next

End Sub



Sub HardDel()

Drived2$ = Mid$(TxtPath.Text, 1, 3)
v1$ = Mid$(TxtPath.Text, InStrRev(TxtPath.Text, "\"))


List1.AddItem "Stage Opp : Del Original Dir"
List1.AddItem "------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For WE = 1 To ListView1.ListItems.COUNT
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(TxtPath.Text)
    C1$ = Drived2$ + "My Music" + v1$ + "\" + Mid$(A1$, ets1 + 2)

    etd1 = 13
    ets2 = 1
    Do
        'jc1$ = Mid$(c1$, etd1 + ets2-1)
        ets2 = InStr(etd1 + ets2, C1$, "\")
        etd1 = 1
        If ets2 = 0 Then Exit Do

        d2$ = Mid$(C1$, 1, ets2 - 1)

        If InStr(f1$, d2$) = 0 Then
            Err.Clear
            On Local Error Resume Next
            'MkDir d2$
            Fs.DeleteFOLDER d2$, True
            If Err.Number <> 75 And Err.Number > 0 Then
                'MsgBox ("Error Making Temp Directory")
                '
            End If
            On Local Error GoTo 0
        End If
    Loop Until ets2 = 0

    f1$ = d2$

    errs2 = 0
    On Local Error Resume Next
'    fs.DeleteFile a1$ + b1$, True
    On Local Error GoTo 0

    'If errs2 <> 0 Then
    '    MsgBox ("Error Moving File to Temp")
    '    End
    'End If

    List1.AddItem Format$(WE, "000 ") + "Move > " + A1$ + B1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'fs.DeleteFolder comrad$ + "\" + WFD$

Next

A = A
jeep:
End Sub


Sub HardMove()

Drived2$ = Mid$(TxtPath.Text, 1, 3)
v1$ = Mid$(TxtPath.Text, InStrRev(TxtPath.Text, "\"))


List1.AddItem "Stage Opp : Del Original Dir"
List1.AddItem "------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For WE = 1 To ListView1.ListItems.COUNT
    A1$ = ListView1.ListItems.Item(WE).SubItems(1)
    B1$ = ListView1.ListItems.Item(WE)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(TxtPath.Text)
    C1$ = Drived2$ + "My Music" + v1$ + "\" + Mid$(A1$, ets1 + 2)
    c2$ = Drived2$ + "Temp\Anything" + v1$ + "\" + Mid$(A1$, ets1 + 2)

    etd1 = 13
    ets2 = 1
    Do
        'jc1$ = Mid$(c1$, etd1 + ets2-1)
        ets2 = InStr(etd1 + ets2, C1$, "\")
        etd1 = 1
        If ets2 = 0 Then Exit Do

        d2$ = Mid$(C1$, 1, ets2 - 1)

        If InStr(f1$, d2$) = 0 Then
            Err.Clear
            On Local Error Resume Next
            MkDir d2$
   '         fs.deletefolder d2$, True
            If Err.Number <> 75 And Err.Number > 0 Then
                'MsgBox ("Error Making Temp Directory")
                'T
            End If
            On Local Error GoTo 0
        End If
    Loop Until ets2 = 0

    f1$ = d2$
    Err.Clear
    errs2 = 0
    On Local Error Resume Next
    
    Fs.MOVEFILE c2$ + B1$, A1$ + B1$
    'err.number
    'err.description
    On Local Error GoTo 0

    'If errs2 <> 0 Then
    '    MsgBox ("Error Moving File to Temp")
    '    End
    'End If

    List1.AddItem Format$(WE, "000 ") + "Move > " + A1$ + B1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'fs.DeleteFolder comrad$ + "\" + WFD$

Next
    
    v1$ = Mid$(TxtPath.Text, 1, InStrRev(TxtPath.Text, "\") - 1)
    Err.Clear
    On Local Error Resume Next
    
    Fs.copyfolder Drived2$ + "Temp\Anything", v1$, True
    
    Fs.DeleteFOLDER Drived2$ + "Temp\Anything", True


jeep:
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


Private Sub TxtPath_Change()
Call txtPath_Click
End Sub

Private Sub txtPath_Click()

On Error Resume Next
If Mid(TxtPath, 2, 1) <> ":" Then
    If Err.Number = 0 Then
    If TxtPath <> Clipboard.GetText Then TxtPath = Clipboard.GetText
    If Mid(TxtPath, 2, 1) <> ":" Then Exit Sub
End If
End If


If ScanPath.TxtPath = "" Then Exit Sub

If Mid(ScanPath.TxtPath, Len(ScanPath.TxtPath), 1) <> "\" Then ScanPath.TxtPath = ScanPath.TxtPath + "\"

TxtPath.SelStart = 0
TxtPath.SelLength = Len(TxtPath)

End Sub

Private Sub txtPath02_Change()

Call txtPath02_CLICK

End Sub

Private Sub txtPath02_CLICK()
On Error Resume Next
'CLIP BOARD OPEN PROBLEM

DESTIPATH = txtPath02

If Mid(txtPath02, 2, 1) <> ":" Then
    If Mid(Clipboard.GetText, 2, 1) <> ":" Then Exit Sub
    If InStr(txtPath02, Clipboard.GetText) > 0 Then Exit Sub
    If txtPath02 <> Clipboard.GetText Then txtPath02 = Clipboard.GetText
End If

If Mid(TxtPath, 2, 1) <> ":" Then Exit Sub


If ScanPath.TxtPath = "" Then Exit Sub


If Mid(ScanPath.txtPath02, Len(ScanPath.txtPath02), 1) <> "\" Then ScanPath.txtPath02 = ScanPath.txtPath02 + "\"
txtPath02.SelStart = 0
txtPath02.SelLength = Len(txtPath02)

End Sub




Private Sub Timer3_Timer()


Dim x1, y1
'Dim Mouse, OX1, OY1

FindCursor x1, y1

'
'VPROG1 = FindWindow("#32770", "Output")
'If VPROG1 > 0 And OVPROG1 <> VPROG1 Then
'    VPROGID1 = True
'End If
'
'VPROG2 = FindWindow("Outlook Express Browser Class", vbNullString)
'If VPROG2 > 0 And OVPROG2 <> VPROG2 Then
'    VPROGID2 = True
'End If
'
'OVPROG1 = VPROG1
'OVPROG2 = VPROG2



'If VPROGID1 = True Then VPROGID1 = 0: OX1 = -1
'If VPROGID2 = True Then VPROGID2 = 0: OX1 = -1
If KCODE = True Then KCODE = 0: OX1 = -1
'If JCODE = True Then JCODE = 0: OX1 = -1


If OX1 <> x1 Or OY1 <> y1 Then
    MOUSE = Now + TimeSerial(0, 0, 5)
End If

If MOUSE > Now Then
    If XPAUSE <> True Then
        XPAUSE = True
       ' SUSPENDCODE
        'Debug.Print "YES"
    End If
Else
    If XPAUSE <> False Then
        XPAUSE = False
       ' SUSPENDCODE
    End If
End If

If XPAUSE = True Then Command17.BackColor = Label23.BackColor
If XPAUSE = 0 Then Command17.BackColor = Label6.BackColor


OX1 = x1
OY1 = y1
'Dim XPAUSE, OX1,OY1

End Sub

Public Sub FindCursor(x, y)

'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'FindCursor Tx, Ty
'Private Type POINTAPI
'        x As Long
'        Y As Long
'End Type

Dim P As POINTAPI

GetCursorPos P
'   return x and y co-ordinate
x = P.x ' / GetSystemMetrics(0) * Screen.Width
'   for current cursor position
y = P.y '/ GetSystemMetrics(1) * Screen.Height

End Sub


