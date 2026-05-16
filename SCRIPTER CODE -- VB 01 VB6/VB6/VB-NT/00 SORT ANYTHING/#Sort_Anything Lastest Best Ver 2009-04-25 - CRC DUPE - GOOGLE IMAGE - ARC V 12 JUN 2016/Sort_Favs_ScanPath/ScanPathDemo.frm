VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ScanPath 
   BackColor       =   &H00808080&
   Caption         =   "ScanPath 2.0 - Sort  Anything -"
   ClientHeight    =   7500
   ClientLeft      =   1440
   ClientTop       =   2520
   ClientWidth     =   15144
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   15144
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.ListBox List1 
      Height          =   1008
      Left            =   10000
      TabIndex        =   66
      Top             =   4440
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.CommandButton COMMAND_WINSTATE_RESIZE 
      BackColor       =   &H0000FF00&
      Caption         =   "Don't Use Winstate && Resize Until End"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.02
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7530
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   840
      Width           =   2445
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "Restore Column Sort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.02
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7530
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   465
      Width           =   2445
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.14
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   945
      TabIndex        =   61
      Text            =   "Not Given Yet"
      Top             =   840
      Width           =   6435
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "If 1st Byte is Continuous && Then Do Naught File Size && Keep Date 1 Day Before && Make a Note Text File Explaining"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.77
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2145
      TabIndex        =   58
      Top             =   2505
      Width           =   7815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.02
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7530
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   15
      Width           =   840
   End
   Begin VB.TextBox First_Folder_Path 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.14
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   945
      TabIndex        =   54
      Text            =   "Z:\OMNIPOTENT"
      Top             =   570
      Width           =   6435
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   150
      TabIndex        =   53
      Top             =   4245
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.02
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8370
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   15
      Width           =   885
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "READ ONLY PROTECTION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.77
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6540
      TabIndex        =   51
      Top             =   2230
      Value           =   1  'Checked
      Width           =   3420
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "1st 12928-HEX Bytes CRC Skip MP3 Tag Header- Use Only 1st Gig of File && Don't Use WildCard Exception Filters"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.77
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4050
      TabIndex        =   50
      Top             =   1740
      Value           =   1  'Checked
      Width           =   5910
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Dont Delete AnyThing 1st Folder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.77
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6540
      TabIndex        =   49
      Top             =   1485
      Value           =   1  'Checked
      Width           =   3420
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.02
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9255
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   15
      Width           =   720
   End
   Begin VB.CheckBox Chk_SortOnDate 
      Caption         =   "Sort On Date"
      Height          =   210
      Left            =   1395
      TabIndex        =   37
      Top             =   1860
      Width           =   1275
   End
   Begin VB.CommandButton CmdScanDir 
      Caption         =   "Scan Dir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.02
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7815
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3720
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help !"
      Height          =   630
      Left            =   1410
      Picture         =   "ScanPathDemo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2100
      Width           =   690
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3900
      Top             =   4155
   End
   Begin VB.ComboBox cboMask 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.02
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   945
      TabIndex        =   1
      Text            =   "*.TXT;*.DAT;*.TMP"
      Top             =   270
      Width           =   6435
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   7080
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   315
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.02
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   945
      TabIndex        =   0
      Text            =   "Z:\OMNIPOTENT"
      Top             =   -15
      Width           =   6120
   End
   Begin VB.CheckBox chkSubFolders 
      Caption         =   "Search Sub-Folders"
      Height          =   210
      Left            =   1395
      TabIndex        =   8
      Top             =   1620
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox chkPatternMatching 
      Caption         =   "Pattern Matching"
      Height          =   210
      Left            =   1395
      TabIndex        =   7
      Top             =   1395
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox chkNormal 
      Caption         =   "Normal"
      Height          =   210
      Left            =   0
      TabIndex        =   2
      Top             =   1395
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkArchive 
      Caption         =   "Archive"
      Height          =   210
      Left            =   0
      TabIndex        =   3
      Top             =   1620
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "Hidden"
      Height          =   210
      Left            =   0
      TabIndex        =   4
      Top             =   1845
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Read Only"
      Height          =   210
      Left            =   0
      TabIndex        =   5
      Top             =   2070
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkSystem 
      Caption         =   "System"
      Height          =   210
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   5085
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3540
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   0
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4950
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   5070
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4605
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   5055
      TabIndex        =   16
      Top             =   4575
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   1
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   5310
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   5055
      TabIndex        =   18
      Top             =   5310
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.CheckBox chkRefreshListView 
      Caption         =   "Refresh ListView (slower)"
      Height          =   195
      Left            =   150
      TabIndex        =   10
      Top             =   3510
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.CheckBox chkCopyMemory 
      Caption         =   "Use CopyMemory (display Date && Size)"
      Height          =   195
      Left            =   150
      TabIndex        =   11
      Top             =   3750
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CheckBox chkSort 
      Caption         =   "Sort Files(A-Z) while Scanning"
      Enabled         =   0   'False
      Height          =   195
      Left            =   135
      TabIndex        =   9
      Top             =   3990
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.02
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3690
      Visible         =   0   'False
      Width           =   795
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   0
      Left            =   5070
      TabIndex        =   13
      Top             =   3885
      Visible         =   0   'False
      Width           =   1545
      _ExtentX        =   2717
      _ExtentY        =   574
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   97189889
      CurrentDate     =   37299
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   1
      Left            =   5070
      TabIndex        =   14
      Top             =   4245
      Visible         =   0   'False
      Width           =   1545
      _ExtentX        =   2717
      _ExtentY        =   574
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   97189889
      CurrentDate     =   37296
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   2
      Left            =   6630
      TabIndex        =   44
      Top             =   3870
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1966
      _ExtentY        =   574
      _Version        =   393216
      Format          =   97189890
      CurrentDate     =   37299
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1650
      Left            =   13550
      TabIndex        =   47
      Top             =   6450
      Width           =   1260
      _ExtentX        =   2231
      _ExtentY        =   2915
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2240
      Left            =   13540
      TabIndex        =   46
      Top             =   4160
      Width           =   1280
      _ExtentX        =   2253
      _ExtentY        =   3931
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblCount12_MEMORY_OVERFLOW3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9980
      TabIndex        =   70
      Top             =   3030
      Width           =   5890
   End
   Begin VB.Label lblCount12_MEMORY_OVERFLOW2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9990
      TabIndex        =   69
      Top             =   2740
      Width           =   5890
   End
   Begin VB.Label lblCount11_MEMORY_OVERFLOW1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9980
      TabIndex        =   68
      Top             =   2450
      Width           =   5890
   End
   Begin VB.Label Label_COMMAND_WINSTATE_RESIZE_COLOR 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label_COMMAND_WINSTATE_RESIZE_COLOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.89
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4395
      TabIndex        =   67
      Top             =   1155
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Label lblCount_BIGGEST_WORKING_SIZE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9980
      TabIndex        =   65
      Top             =   2150
      Width           =   5890
   End
   Begin VB.Label Label7 
      Caption         =   "Logg File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.77
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   62
      Top             =   855
      Width           =   930
   End
   Begin VB.Label lblCount10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9990
      TabIndex        =   60
      Top             =   3330
      Width           =   5890
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label lblCount9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9980
      TabIndex        =   57
      Top             =   1850
      Width           =   5890
   End
   Begin VB.Label Label5 
      Caption         =   "1st Folder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.77
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   55
      Top             =   570
      Width           =   930
   End
   Begin VB.Label lblCount8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9990
      TabIndex        =   21
      Top             =   3930
      Width           =   5890
   End
   Begin VB.Label lblCount7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9990
      TabIndex        =   45
      Top             =   3630
      Width           =   5890
   End
   Begin VB.Label lblCount1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   320
      Left            =   9970
      TabIndex        =   43
      Top             =   20
      Width           =   5890
   End
   Begin VB.Label lblCount3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9980
      TabIndex        =   42
      Top             =   650
      Width           =   5890
   End
   Begin VB.Label lblCount4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9980
      TabIndex        =   41
      Top             =   950
      Width           =   5890
   End
   Begin VB.Label lblCount5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9980
      TabIndex        =   40
      Top             =   1250
      Width           =   5890
   End
   Begin VB.Label lblCount2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   320
      Left            =   9980
      TabIndex        =   39
      Top             =   330
      Width           =   5890
   End
   Begin VB.Label lblCount6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9980
      TabIndex        =   38
      Top             =   1550
      Width           =   5890
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Working"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8685
      TabIndex        =   34
      Top             =   5325
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Date/Size Filter:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.02
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5100
      TabIndex        =   33
      Top             =   3270
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Options:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.02
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1425
      TabIndex        =   26
      Top             =   1155
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Mask"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.77
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   24
      Top             =   270
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.77
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Attributes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.02
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   30
      TabIndex        =   25
      Top             =   1155
      Width           =   1065
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Min"
      Height          =   195
      Index           =   0
      Left            =   4455
      TabIndex        =   30
      Top             =   4635
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "File Size"
      Height          =   195
      Left            =   4470
      TabIndex        =   28
      Top             =   4260
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Max"
      Height          =   195
      Index           =   1
      Left            =   4455
      TabIndex        =   32
      Top             =   4995
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   195
      Left            =   4470
      TabIndex        =   29
      Top             =   3795
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "File Date"
      Height          =   195
      Left            =   4470
      TabIndex        =   27
      Top             =   3570
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   4455
      TabIndex        =   31
      Top             =   3840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Menu Mnu_Cmd 
      Caption         =   "Commands"
      Begin VB.Menu Mnu_Com1 
         Caption         =   "Com1"
      End
      Begin VB.Menu Mnu_Go_Note 
         Caption         =   "Show Note That Happens on Go Button"
      End
      Begin VB.Menu Mnu_Repeat_Questions_Toggle 
         Caption         =   "Repeat Questions Toggle Checked = False"
      End
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu Mnu_NotePad_Logg 
      Caption         =   "Open Logg File"
   End
   Begin VB.Menu Mnu_Explorer 
      Caption         =   "Explorer Logg"
   End
   Begin VB.Menu Mnu_Explorer_Column 
      Caption         =   "Explorer ListView"
      Begin VB.Menu Mnu_Explore_Path 
         Caption         =   "Explorer Path"
      End
      Begin VB.Menu Mnu_Explore_File_Select 
         Caption         =   "Explorer File Select"
      End
   End
   Begin VB.Menu Mnu_Options 
      Caption         =   "Options"
      Begin VB.Menu mnu_Dont_Del_From_First_Source_Folder 
         Caption         =   "Don't Del From First Source Folder - The Next Sorted Level Up From Source Path"
         Checked         =   -1  'True
      End
      Begin VB.Menu MNU_CRC_Check_After_First 
         Caption         =   "CRC Check After First 800 Bytes for MP3 Headers"
      End
      Begin VB.Menu Mnu_Simulate 
         Caption         =   "Simulate All Read Only"
      End
      Begin VB.Menu MNU_Check_FOR_NULL_STRING 
         Caption         =   "Check For All Null String And Or Same Byte File"
      End
   End
   Begin VB.Menu MNU_COMPARE 
      Caption         =   "WinMergeU.exe COMPARE"
   End
End
Attribute VB_Name = "ScanPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------
'TUE 17 mAY
'NEED NUMBERING SYSTEM FOR DUPE COUNT IN LOGG
'HERE
'WHOLE FILE IS A CRC32 Duplicate
'--------------------------------------------


'WORK DONE END SUN 15 MAY MIDNIGHT HOUR
'------------------------------------------
'NOW CAN DO CHUNK HUGE FILE COMPARE
'STILL READ IN 10 MEG CHUNK LESS WAS SLOWER
'------------------------------------------
'WORK TO DO MAKE LABEL DISPLAY LARGEST SIZE FILE IN THE WORK TO DO
'AT MOMENT HAS LARGEST FOUND IN SCAN
'------------------------------------------




'-------------Is This Sub Get on It at start      at       Bangers


Dim INWINRESIZE

'Public Option_Mnu_Any_Selected
Private Declare Function Putfocus _
        Lib "user32" _
        Alias "SetFocus" _
        (ByVal hWnd As Long) As Long

Dim COMPARE_STATE
Dim COMPARE_FILE1, COMPARE_FILE2

Public Go_RoutineGoing


Dim Checks_Dupe_Escape

Public ACTIVE_STATE
Public Dog, OutPutPath, Xdate As Date, Tdate As Date, Zdate$, Ydate As Date, OldFolder, OldSize

'Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



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

Private WithEvents SP As cScanPath
Attribute SP.VB_VarHelpID = -1


Sub Del_Emptys()

'Dim Fs22
Dim Fs
'Set Fs22 = CreateObject("Scripting.FileSystemObject")
Set Fs = CreateObject("Scripting.FileSystemObject")


'txtPath.Text = "M:\01 Installations\01 DVD Nfo Txt MP3 JPG"
''txtPath.Text = "M:\01 Installations\01 DVD Nfo Txt Log Diz"
'txtPath = "M:\01 Installations\01 DVD MP3 JPG GIF PDF"
'txtPath = "M:\00 Archive Work"
''txtPath = "M:\01 Archives"
'
'txtPath = "M:\01 Installations Icy"
'cmdScanDir_Click
'
'txtPath = "M:\01 Installations Lacie"
'cmdScanDir_Click
'
'txtPath = "M:\01 Installations WD200"
'cmdScanDir_Click
'
'
'
'
'
'
'T:\0 00 ART LOGGERS\# FEED STATION\My FeedStation Podcasts DATE ORDER\## ARCHIVE

ScanPath.ListView1.ListItems.Clear

If MENU_OPTION = 4 Then
    
    'If Fs.FolderExists(Clipboard.GetText) = True Then
    If Fs.FolderExists(Label20.Caption) = True Then
    
        'ScanPath.txtPath = Clipboard.GetText
        ScanPath.txtPath = Label20.Caption
        'Make Sure Backslash at End of Scan Path
        If Mid(ScanPath.txtPath, Len(ScanPath.txtPath)) <> "\" Then ScanPath.txtPath = ScanPath.txtPath + "\"
        
        'Dont Do Root's
        If Len(ScanPath.txtPath) < 4 Then End
        'Scan Folder Must Exist
        If Fs.FolderExists(ScanPath.txtPath) = False Then End
        Call cmdScanDir_Click
    End If
End If

If MENU_OPTION = 5 Then
    w$ = "E:\01 Start Menu\"
    w$ = "T:\0 00 ART LOGGERS"
    If Fs.FolderExists(w$) = True Then
    
        ScanPath.txtPath = w$
        'Make Sure Backslash at End of Scan Path
        If Mid(w$, Len(w$)) <> "\" Then ScanPath.txtPath = w$ + "\"
        
        'Dont Do Root's
        If Len(ScanPath.txtPath) < 4 Then End
        'Scan Folder Must Exist
        If Fs.FolderExists(ScanPath.txtPath) = False Then End
        Call cmdScanDir_Click
    End If
End If


'JUST IN CASE SORTED OPTION NOT SELECTED
'SORT ON FILE-PATH
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 0
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False
'SORT ON PATH-PATH
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 1
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False
          

    

Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(3))
Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(4))
    
    
    
Dim WFD As WIN32_FIND_DATA
Dim lResult As Long

For we = ListView1.ListItems.Count To 1 Step -1
    
    lblCount2 = Str(we)
    DoEvents
    
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)
    C1$ = A1$ + B1$
    
    
    ScanPath.ListView1.ListItems(we).Selected = True
    ScanPath.ListView1.SelectedItem.EnsureVisible
    DoEvents
    
    
    ERRZ = False
    On Error Resume Next
    Set F = Fs.GetFolder(C1$)
    If Err.Number > 0 Then
        ERRZ = True
        'MsgBox "ERROR GETFOLDER WITH THIS FILE" + vbCrLf + A1$ + B1$, vbMsgBoxSetForeground
        ScanPath.ListView1.ListItems.Item(we).SubItems(4) = "Get Folder Error"
        errhuge = errhuge + 1
        lblCount3 = Str(errhuge) + " Delete Errors"

    End If
    On Error GoTo 0
    
    lblCount = Trim(Str(we))
    'Can be Processor Heavy if Folders Contain Lot Files
    If ERRZ = False Then
            
        ScanPath.ListView1.ListItems.Item(we).SubItems(3) = Str(F.Size)
        'Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(4))
        If Val(ScanPath.ListView1.ListItems.Item(we).SubItems(3)) > 0 Then Stop
            
        If F.Size = 0 Then
            On Error Resume Next
            
            'THIS WILL DELETE WITH CONTENTS SO CHECK SIZE
            'FS.Deletefolder c1$, True 'True = Readonly deleted
            
            'MUST CHECK FOR NAUGHT SIZE FILES ALSO
            
            RmDir C1$
            If Err.Number > 0 Then
                'MsgBox "Cant Del " + vbCrLf + c1$
                errhuge = errhuge + 1
                ScanPath.ListView1.ListItems.Item(we).SubItems(4) = "Delete Error " + Err.Description
                Exit Sub
                lblCount3 = Str(errhuge) + " Delete Errors"
                
                Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(5))
            
            Else
                huge = huge + 1
                ScanPath.ListView1.ListItems.Item(we).SubItems(4) = "Deleted"
                DoEvents
                'Call LV_AutoSizeColumn(ScanPath.ListView1, ScanPath.ListView1.ColumnHeaders.Item(5))
            End If
        End If
        On Error GoTo 0
    End If
    DoEvents
Next

'MsgBox Str(huge) + " Folders Deleted" + vbCrLf + Str(errhuge) + " Folders Not Deleted"

Exit Sub

End




On Local Error GoTo jeep
Drived2$ = Mid$(txtPath.Text, 1, 3)
MkDir Drived2$ + "Temp\Anything"
v1$ = Mid$(txtPath.Text, InStrRev(txtPath.Text, "\"))

MkDir Drived2$ + "Temp\Anything" + v1$
On Local Error GoTo 0

List1.AddItem "Stage 1 of 2 : Make Dir's And Move Files to Temp"
List1.AddItem "------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For we = 1 To ListView1.ListItems.Count
    
    lblCount = Str(we)
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(txtPath.Text)
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
                i = MsgBox("Error Making Temp Directory", , ScanPath.Caption)
                Stop
            End If
            On Local Error GoTo 0
        End If
    Loop Until ets2 = 0

    f1$ = d2$

    errs2 = 0
    On Local Error GoTo jeep
    Fs22.moveFile A1$ + B1$, C1$ + B1$
    On Local Error GoTo 0

    If errs2 <> 0 Then
        i = MsgBox("Error Moving File to Temp", , ScanPath.Caption)
        End
    End If

    List1.AddItem Format$(we, "000 ") + "Move > " + A1$ + B1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'Fs22.DeleteFolder comrad$ + "\" + WFD$

Next

Err.Clear
On Local Error Resume Next
Set Fs22 = CreateObject("Scripting.FileSystemObject")
Fs22.Deletefolder txtPath.Text, True
If Err.Number > 0 Then Call HardDel
On Local Error GoTo 0


List1.AddItem "----------------------------------------"
List1.AddItem "Stage 2 of 2 : Move Files Back From Temp"
List1.AddItem "----------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

v2$ = Mid$(txtPath.Text, 1, InStrRev(txtPath.Text, "\"))

Err.Clear
On Local Error Resume Next
Set Fs22 = CreateObject("Scripting.FileSystemObject")
Fs22.Movefolder Drived2$ + "Temp\Anything\*", v2$
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


Sub DelJunkFrontPage()


ScanPath.ListView1.ListItems.Clear
ScanPath.txtPath = "D:\My Webs\MatthewLan.Com Web\"
ScanPath.cboMask.Text = "*.*"
ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScanDir_Click
ScanPath.Show
Reset
For we = ScanPath.ListView1.ListItems.Count To 1 Step -1
'    we = ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount1 = Trim(Str(we))
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    A1$ = A1$ + ScanPath.ListView1.ListItems.Item(we)
    outff = 0
    If InStr(A1$, "_private") > 0 Then outff = 1
    If InStr(A1$, "_vti") > 0 Then
        outff = 1
    End If
    
    If outff = 1 Then
        'ScanPath.ListView1.ListItems.Remove (we)
        On Error Resume Next
        Fs.Deletefolder A1$, True
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

txtPath.Text = "C:\Documents and Settings\Matt3\Local Settings\Temporary Internet Files\Content.IE5\"
txtPath.Text = "D:\My Webs\MatthewLan.Com Web\"


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
txtPath.Text = "C:\Documents and Settings\Matt4\Local Settings\Temporary Internet Files\Content.IE5\"
txtPath.Text = "D:\My Webs\MatthewLan.Com Web\"
'Call cmdScan_Click

For we = ListView1.ListItems.Count To 1 Step -1
    
    Label13 = ListView1.ListItems.Count
    DoEvents
    
    G1$ = ListView1.ListItems.Item(we).SubItems(2)
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)
    
    Set F = Fs.getfile(A1$ + G1$)
    Tdate = F.datelastmodified
    If Tdate > Xdate Then Xdate = Tdate
    
    remo = 0
    If Tdate <= Ydate Then
        ListView1.ListItems.Remove (we)
        remo = 1
    End If
    If remo = 0 Then
        hole = "Inet " + Format$(Now, "YYYY-MM-DD") + "\"
        GeText = LCase(Mid$(B1$, Len(B1$) - 2, 3)) + "\"
        If GeText = "jpg\" Then GeText = ""
        secext = "M:\0 00 Art\Temp Inet Files JPGs\" + hole + GeText + B1$
        If Fs.FileExists(secext) = True Then
            ListView1.ListItems.Remove (we)
        End If
        secext = OutPutPath + hole + GeText + Mid$(B1$, 1, Len(B1$) - 4) + ".zip"
        If Fs.FileExists(secext) = True Then
            ListView1.ListItems.Remove (we)
        End If
        secext = OutPutPath + hole + GeText + Mid$(B1$, 1, Len(B1$) - 4) + ".rar"
        If Fs.FileExists(secext) = True Then
            ListView1.ListItems.Remove (we)
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

For we = ListView1.ListItems.Count To 1 Step -1
    
    Label13 = ListView1.ListItems.Count
    DoEvents
    
    G1$ = ListView1.ListItems.Item(we).SubItems(2)
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)
    
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
txtPath.Text = "C:\Documents and Settings\Matt3\Local Settings\Temporary Internet Files\Content.IE5\"
cboSize.ListIndex = 0 ' Not Used
cboMask.Text = "*.png;*.ico;*.gif"
Call cmdScan_Click

Exit Sub

For we = ListView1.ListItems.Count To 1 Step -1
    
    Label13 = ListView1.ListItems.Count
    DoEvents
    
    G1$ = ListView1.ListItems.Item(we).SubItems(2)
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)
    Set F = Fs.getfile(A1$ + G1$)
    Tdate = F.datelastmodified
    If Tdate > Xdate Then Xdate = Tdate
    remo = 0
    If Tdate <= Ydate Then
        ListView1.ListItems.Remove (we)
        remo = 1
    End If
    If remo = 0 Then
        GeText = "#" + LCase(Mid$(B1$, Len(B1$) - 2, 3)) + "\"
        If Fs.FileExists(OutPutPath + GeText + B1$) = True Then
        ListView1.ListItems.Remove (we)
        End If
    End If
Next

Open App.Path + "\Xdate.txt" For Output As #1
Print #1, Xdate
Close #1


For we = ListView1.ListItems.Count To 1 Step -1
    
    Label13 = ListView1.ListItems.Count
    DoEvents
    
    G1$ = ListView1.ListItems.Item(we).SubItems(2)
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)
    
    Set F = Fs.getfile(A1$ + G1$)

    GeText = "#" + LCase(Mid$(B1$, Len(B1$) - 2, 3)) + "\"
    F.Copy OutPutPath + GeText
Next

ListView1.ListItems.Clear
txtPath.Text = "C:\Documents and Settings\Matt3\Local Settings\Temporary Internet Files\Content.IE5\"
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
txtPath.Text = tpath1$

cboMask.Text = "*.mp3"

txtPath.Text = "F:\04 Music ---\04 My Music"
txtPath.Text = "F:\04 Music ---\04 My Music Talking Books"

Dim ra(30)

ra(1) = "M:\05 Media\MIDI"
ra(2) = "M:\04 Music ---\03 My Music Zen"
ra(3) = "M:\04 Music ---\04 My Music"
ra(4) = "M:\04 Music ---\05 Whole Lot"
ra(5) = "M:\04 Music ---\00 TBooks All"

ra(1) = "M:\04 Music ---"
ra(2) = "M:\Videos\00 Vid Films"
ra(3) = "M:\Videos\00 Vid XXX"
'ra(3) = "M:\04 Music ---\04 My Music Talking Books"

'ra(1) = "M:\04 Music ---\"
'ra(2) = "M:\04 Music ---\04 My Music Talking Books"
'ra(3) = "S:\Tbooks"
'ra(4) = "S:\My Music Talking Books"



For ret = 1 To 3
DoEvents
If Right$(ra(ret), 1) <> "\" Then ra(ret) = ra(ret) + "\"

txtPath.Text = ra(ret)

cboSize.ListIndex = 1 'Less than
txtSize(0) = 1

If Dog = True Then Exit Sub

Call cmdScan_Click

If Dog = True Then Exit Sub
DoEvents
If Dog = True Then Exit Sub

For we = ListView1.ListItems.Count To 1 Step -1
    
    Label13 = we
    DoEvents
    
    G1$ = ListView1.ListItems.Item(we).SubItems(2)
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)
    w = DeleteFile(A1$ + G1$)

Next

Call cmdScanDir_Click

For we = ListView1.ListItems.Count To 1 Step -1
  
    Label14 = we
    DoEvents
    
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)
    A1$ = A1$ + B1$ + "\"
  
    
    Set F = Fs.GetFolder(A1$)
    
    If F.Size = 0 Then
        ListView1.ListItems.Remove (we)
    End If
Next

A1$ = txtPath.Text

Call Tagg2


For we = 1 To ListView1.ListItems.Count
  
    Label16 = we
    DoEvents
    
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)
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
 
 

FF$ = Mid$(A1$, Len(txtPath.Text) + 1)
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
 
    
ty$ = ""
If InStr(A1$, "My Music\") Then ty$ = "My Music"
If InStr(A1$, "My Music Zen\") Then ty$ = "My Music Zen"
If ty$ = "My Music" And InStr(A1$, "Favs") Then ty$ = "My Music - Favs"
If ty$ = "My Music Zen" And InStr(A1$, "Favs") Then ty$ = "My Music Zen - Favs"
If ty$ = "My Music" And InStr(A1$, " Bang Sets") Then ty$ = "My Music - Sets"
If ty$ = "My Music Zen" And InStr(A1$, " Bang Sets") Then ty$ = "My Music Zen - Sets"
If ty$ = "My Music" And InStr(A1$, " Bang Fav Sets") Then ty$ = "My Music - Fav Sets"
If ty$ = "My Music Zen" And InStr(A1$, " Bang Fav Sets") Then ty$ = "My Music Zen - Fav Sets"
'03 Bang Sets
    
    
gg$ = Mid$(A1$, InStrRev(A1$, "\", Len(A1$) - 1) + 1)
Tx$ = Left$(gg$, Len(gg$) - 1)

FF$ = "---------------[[(({{ " + Tx$ + " - " + ty$ + " }}))]]---------------.mp3"


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

Private Sub Check1_Click()
    Call Checks_Dupe2
    Option_Mnu_Any_Selected = True
End Sub
Private Sub Check2_Click()
    Call Checks_Dupe2
    Option_Mnu_Any_Selected = True
End Sub
Private Sub Check3_Click()
    Call Checks_Dupe2
    Option_Mnu_Any_Selected = True
End Sub

Private Sub Check4_Click()
MsgBox "This Menu Option Will Require Check All Files and Will Be Available Later"
End Sub

Private Sub chkCopyMemory_Click()
    
    chkSort.Enabled = (chkCopyMemory.Value = vbUnchecked)
    
End Sub


Private Sub chkSort_Click()
    chkCopyMemory.Enabled = (chkSort.Value = vbUnchecked)
End Sub

Private Sub cmdBrowse_Click()
    txtPath.Text = GetFolder(Me.hWnd, "Scan Path:", txtPath.Text)
    fg1 = FreeFile
    Open App.Path + "\Scan Path.txt" For Output As #fg1
    Print #fg1, txtPath.Text
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
    Scan_Individual = 0
    
    
    If cmdScan.Caption = "Scan" Then
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
                MsgBox "Maximium Size value must exceed 0", vbExclamation, ScanPath.Caption
                txtSize(1).SetFocus
                Exit Sub
            ElseIf lSize(1) <= lSize(0) Then
                MsgBox "Maximium Size value must exceed Minimum Size", vbExclamation, ScanPath.Caption
                txtSize(1).SetFocus
                Exit Sub
            End If
        End If
        
        '####################################################################################################
        'Reset Form
        cmdScan.Caption = "Stop"
        
        lblCount1.Caption = "-"
 
        Screen.MousePointer = vbHourglass
'        ListView1.ListItems.Clear
        
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
            
            RDClip = RDClip + txtPath + vbCrLf
            RDClip = RDClip + cboMask + vbCrLf
            
            'Go - that was easy wasn't it!
            .StartScan txtPath, chkSubFolders.Value, chkSort.Value And chkSort.Enabled, chkPatternMatching
            
            If mCancelScan = True Then Exit Sub: If ScanPath.Enabled = Visible Then SP.StopScan
            
            
            
            'we always use direct scan as is better to sort after job then during thats all we need
            If Chk_SortOnDate.Value = vbChecked Then
                ListView1.SortOrder = lvwAscending
                ListView1.SortKey = 3
                ListView1.Sorted = True
                ListView1.Sorted = False
            Else
                'ListView1.SortOrder = lvwAscending
'                ListView1.SortKey = 0
'                ListView1.Sorted = True
'                ListView1.SortKey = 1
'                ListView1.Sorted = True
'                ListView1.Sorted = False
            
                'JUST IN CASE SORTED OPTION NOT SELECTED
                'SORT ON FILE-PATH
                ScanPath.ListView1.SortOrder = lvwAscending
                ScanPath.ListView1.SortKey = 0
                ScanPath.ListView1.Sorted = True
                ScanPath.ListView1.Sorted = False
                'SORT ON PATH-PATH
                ScanPath.ListView1.SortOrder = lvwAscending
                ScanPath.ListView1.SortKey = 1
                ScanPath.ListView1.Sorted = True
                ScanPath.ListView1.Sorted = False
            
            End If
        
        End With
        
       With ListView1
            If .ListItems.Count > 0 Then
                .SelectedItem = .ListItems(1)
            End If
            
            lblCount1.Caption = .ListItems.Count
            RDClip = RDClip + Trim(Str(Scan_Individual)) + " Files This Individual Scan Folder" + vbCrLf
            RDClip = RDClip + lblCount1.Caption + " Files Total Scan" + vbCrLf
        End With
        
      '  LockWindowUpdate 0
        
        'cmdScan.Caption = "Scan"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
        'End
    End If


'frmMain.Refresh

'End
'If Command$ <> "" Then End
cmdScan.Caption = "Scan"

If Chk_SortOnDate = vbChecked Then

End If

End Sub

Public Sub LOCK_WINDOW_UPDATE_SUB(X)

If X = 1 Then LockWindowUpdate ListView1.hWnd
If X = 0 Then LockWindowUpdate 0

End Sub




Public Sub cmdScanDir_Click()
    Dim lCount As Long
    Dim lSize(1) As Long
    
    If cmdScan.Caption = "Scan" Then
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
                MsgBox "Maximium Size value must exceed 0", vbExclamation, ScanPath.Caption
                txtSize(1).SetFocus
                Exit Sub
            ElseIf lSize(1) <= lSize(0) Then
                MsgBox "Maximium Size value must exceed Minimum Size", vbExclamation, ScanPath.Caption
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
            .StartScanDir txtPath, chkSubFolders.Value, chkSort.Value And chkSort.Enabled, chkPatternMatching
        
            
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
            If .ListItems.Count > 0 Then
                .SelectedItem = .ListItems(1)
            End If
            
            lblCount1.Caption = .ListItems.Count
        End With
        
        LockWindowUpdate 0
        
        cmdScan.Caption = "Scan"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
    End If


lblCount1.Caption = ListView1.ListItems.Count
    

End Sub


Private Sub COMMAND_WINSTATE_RESIZE_Click()

'COMMAND_WINSTATE_RESIZE = Not COMMAND_WINSTATE_RESIZE

If COMMAND_WINSTATE_RESIZE.BackColor <> Label_COMMAND_WINSTATE_RESIZE_COLOR.BackColor Then

    COMMAND_WINSTATE_RESIZE.BackColor = Label_COMMAND_WINSTATE_RESIZE_COLOR.BackColor

Else

    COMMAND_WINSTATE_RESIZE.BackColor = &HFF00&

End If

End Sub

Private Sub Command1_Click()
    
    If Go_RoutineGoing = True Then
        i = MsgBox("Dont Want to Run Go Twice Continue Yes or Exit", vbYesNo, ScanPath.Caption)
        If i = vbYes Then Exit Sub
        CommandExit = True: Exit Sub
    End If

    
    If Dir(App.Path + "Menu Options 1.txt") = "" Then
        i = MsgBox("One HelpFul Note - This Process Sorts The Files On Folder And Then Files Which Keeps a Sort in Both of a Kind" + vbCrLf + "That Means Any Thing Alphabetically Later Will Have Less Priority And Be Deleted" + vbCrLf + "So Sort Your Folders Out" + vbCrLf + "Another Thing It Only Supports 2.5 Million Files in 3GB 32BIT of Memory" + vbCrLf + "Later I Want to Make a Edit File That Is Not Hardcoded Filters Easy Gota Array Already" + vbCrLf + "If I Knew Drag and Drop I Would Use That" + vbCrLf + "~" + vbCr + "Shine" + vbCrLf + "Select Yes If You Read This Message", vbYesNo, ScanPath.Caption)
        If i = vbYes Then
            FR1 = FreeFile
            Open App.Path + "Menu Options 1.txt" For Output As #FR1
                Print #FR1, "Opt 1"
            Close #FR1
        End If
    End If
    
    Go_RoutineGoing = True
    
    Call CRCDUPE

End Sub

Private Sub Command2_Click()

CommandPause = Not CommandPause

If CommandPause = True Then

    Command2.BackColor = vbRed

Else

    Command2.BackColor = &HFF00&

End If


End Sub

Private Sub Command3_Click()
CommandExit = True
mCancelScan = True
Unload Me
End Sub

Private Sub Command4_Click()
If WORK_DONE = False Then Exit Sub
'JUST IN CASE SORTED OPTION NOT SELECTED
'SORT ON FILE-PATH
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 0
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False
'SORT ON PATH-PATH
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 1
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False
'SORT ON SIZE
ScanPath.ListView1.SortOrder = lvwAscending
ScanPath.ListView1.SortKey = 3
ScanPath.ListView1.Sorted = True
ScanPath.ListView1.Sorted = False
End Sub

Private Sub Form_Activate()
    Call Form_Resize
End Sub

Private Sub Form_GotFocus()
    Call Form_Resize
End Sub

Private Sub Form_Load()
        
    If mCancelScan = True Then Exit Sub: ' vbMinimized
    
    'TEST
    Shell "C:\Program Files\ViceVersa Pro\ViceVersa.exe ""D:\VICE_VERSA SCRIPT FILE\TEST COPY BANGER IMAGES FOR CRC DUPE CHECKER.fsf"" /autoexec /autoclose", vbNormalFocus
    
    
    
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
        
    Me.Caption = "Sort_AnyThing - CRC DUPE - GOOGLE IMAGE" 'ScanPath.Caption + App.EXEName"
    'scanpath.Caption = frmMain.Caption + App.EXEName
      
    'ScanPath.Name = ScanPath.Caption
      
    ScanPath.txtPath = PathToLoad 'z_MENU_Form1.Label20.Caption
      
    X = 1
    y = 1
    On Error Resume Next
    For Each Control In Controls
        If Control.Enabled = True Then
            If Control.Width + Control.Left > X Then X = Control.Width + Control.Left
            If Control.Height + Control.Top > y Then y = Control.Height + Control.Top
            If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
        End If
    Next
    On Error GoTo 0
    
    
    ScanPath.Width = X + 1500
    If mnu = 1 Then
        pluso = 1000
    Else
        pluso = 450
    End If
    ScanPath.Height = y + pluso
        
    
    ListView1.Left = 0
    ListView2.Left = 0
    ListView1.Width = ScanPath.Width
    ListView2.Width = ScanPath.Width
    
    ListView1.Top = lblCount8.Top + lblCount8.Height
    ListView2.Top = ListView1.Top + ListView1.Height
    
    Me.Top = 800
    
    
    
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
        .ColumnHeaders.Add , "PATH", "Path", 5000, lvwColumnLeft
        .ColumnHeaders.Add , "DOS-8.3", "Dos", 1500, lvwColumnLeft
        .ColumnHeaders.Add , "SIZE", "Size", 2200, lvwColumnLeft
        .ColumnHeaders.Add , "DELETED", "DELETED", 2500, lvwColumnLeft
        
        .View = lvwReport
    End With
    
    With ListView2
        .ColumnHeaders.Add , "RESULT", "Status", 700 - 50, lvwColumnLeft
        .ColumnHeaders.Add , "FILE", "File", 4000, lvwColumnLeft
        .ColumnHeaders.Add , "PATH", "Path", 9500, lvwColumnLeft
        .View = lvwReport
    End With
  
    chkCopyMemory.Value = vbUnchecked
'    Call chkCopyMemory_Click

If IsIDE = True Then
    lblCount8 = "IsIDE = True"
Else
    lblCount8 = "IsIDE = False"
End If

MaxFileSizeCRCCheck = (1024 ^ 2) * 100
MP3_HEADER_TAG_SIZE = &H12948

ScanPath.Check2.Caption = "1st 12928-HEX Bytes CRC Skip MP3 Tag Header && Use Only 1st" + Str(Int(MaxFileSizeCRCCheck / 1024 ^ 2)) + " Meg of File && Don't Use WildCard Exception Filters"

'CRC Check After First 800 Bytes for MP3 Headers
MNU_CRC_Check_After_First.Caption = "CRC Check after 1st " + Hex$(MP3_HEADER_TAG_SIZE) + " Hex Bytes for Media Tag Heading Like MP3"


Call Mnu_Simulate_SET_CHECKED_Click

'Call DelJunkFrontPage

'ScanPath.Show

'Call INetContent

'Call Bangers

'txtPath.Text = "D:\My Webs\MatthewLan.Com Web\"

'Call CRCDUPE

'Exit Sub

Me.Hide
Load Z_Choice_Frm

'Load z_MENU_Form1
z_MENU_Form1.Show


Call Checks_Dupe1

If Z_Choice_Frm.Visible = True Then z_MENU_Form1.Visible = False
If Z_Choice_Frm.Visible = True Then Putfocus (Z_Choice_Frm.hWnd)

COMMAND_WINSTATE_RESIZE.BackColor = Label_COMMAND_WINSTATE_RESIZE_COLOR.BackColor

End Sub


Private Sub Form_LostFocus()
Call Form_Resize
End Sub

Private Sub Form_Resize()
'LockWindowUpdate ListView1.hWnd
''Exit Sub
'INWINRESIZE = False
If INWINRESIZE = True Then Exit Sub

INWINRESIZE = True

XGO = 1
If ACTIVE_STATE = "Gather Files" Then XGO = 0
If ACTIVE_STATE = "Put The Sizes" Then XGO = 0
XGO = 0

If 1 = 2 Or Me.WindowState <> vbNormal Then
    INWINRESIZE = False
    
    Exit Sub
'    Me.WindowState = vbNormal
End If

ListView2.Height = 1500

'VAR_BOTTOM_OBJECT = ListView2.Height + ListView2.Top
VAR_WIDTH_OBJECT = lblCount1.Left + lblCount1.Width ' + 250

'If ListView2.Width > VAR_WIDTH_OBJECT Then VAR_WIDTH_OBJECT = ListView2.Width


'On Error Resume Next
    Me.Width = VAR_WIDTH_OBJECT + 190
'    VARCENTER = False
'    Me.Height = (VAR_BOTTOM_OBJECT - Menu_Height * Screen.TwipsPerPixelY) + 1170 + (30 * 2)
'Me.Height = Label1.Height + 2500


Me.Top = Screen.Height / 2 - Me.Height / 2 '+400
Me.Left = Screen.Width / 2 - Me.Width / 2 '+400


If XGO = 1 Then INWINRESIZE = False: Exit Sub
If Me.Visible = False Then INWINRESIZE = False: Exit Sub
If Me.WindowState <> vbNormal Then INWINRESIZE = False: Exit Sub

Call ScanPath.LOCK_WINDOW_UPDATE_SUB(1)

'Me.Height = List1.Top + List1.Height
'Me.Width = List1.Left + List1.Width

ListView1.Left = 0
ListView1.Left = 0
'List1.Left = 0
ListView2.Width = lblCount1.Left + lblCount1.Width
ListView1.Width = lblCount1.Left + lblCount1.Width
ListView2.Height = 1550
DoEvents
ListView2.Refresh
'ListView2.Height = Me.Height - ((ListView1.Height + ListView1.Top)) '- 1550) 'END # HIGHER = LOWER
'List1.Height = Me.Height - ((ListView1.Height + ListView1.Top) - 1550) 'END # HIGHER = LOWER


Me.Height = ListView2.Top + ListView2.Height + 850 '- 700
'Me.Width = ListView2.Left + ListView2.Width + 120
Me.Width = lblCount1.Left + lblCount1.Width + 250

Me.Top = 500

INWINRESIZE = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

'If ScanPath.Enabled = True Then If ScanPath.Enabled = Visible Then SP.StopScan
CommandExit = True
mCancelScan = True
'Me.Hide

'z_MENU_Form1.Hide
'Unload z_MENU_Form1
'----------------------
' ABORT USING THIS
' Call zzSave_Checks
'----------------------
' ONLY SAVE IF CHANGES
'Call zzCheckTimer_Timer
'----------------------

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
   
DoEvents
   
Dim Form As Form
For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next Form

For r = 1 To 100
    DoEvents
Next

For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next Form

For r = 1 To 100
    DoEvents
Next

For r = 1 To 100
    DoEvents
Next

For r = 1 To 100
    DoEvents
Next

For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next Form

For r = 1 To 100
    DoEvents
Next

For Each Form In Forms
    Unload Form
'    DoEvents
    Set Form = Nothing
Next Form

Reset
Close


Unload z_MENU_Form1
Unload Z_Choice_Frm
Unload Me


'Crash End
'UUNABLE TO SORT AT MOMENT
End


'Stop

End Sub

Private Sub List1_Click()
Exit Sub
tpath3$ = "M:\04 Music ---\Del\"

B1$ = Mid$(List1.List(List1.ListIndex), 14)
C1$ = B1$ + "--Todel"

'tpath1$ = "M:\04 Music ---\"

C1$ = Mid$(B1$, InStrRev(B1$, "\") + 1)
List1.RemoveItem (List1.ListIndex)
'Name b1$ As c1$

'fs.copyFile b1$, tpath3$ + c1$
If Dir$(tpath3$ + C1$) <> "" Then Kill tpath3$ + C1$
Fs.moveFile B1$, tpath3$ + C1$



End Sub

Private Sub Label18_Click()
Dog = True

If cmdScan.Caption = "Stop" Then cmdScan_Click
Label17 = "Stopped before Complete"
End Sub

Private Sub lblCount1_Change()
    'Main.Command1.Caption = ScanPath.lblCount1

End Sub

Private Sub lblCount7_Change()
    ACTIVE_STATE = lblCount7
End Sub



Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

Exit Sub

End Sub

Private Sub ListView1_Click()
If COMPARE_STATE = True Then
    If COMPARE_FILE1 = "" Then
        
        DA1$ = ScanPath.ListView1.ListItems.Item(ListView1.SelectedItem.Index).SubItems(1)
        DB1$ = ScanPath.ListView1.ListItems.Item(ListView1.SelectedItem.Index)
        
        COMPARE_FILE1 = """" + DA1$ + DB1$ + """"
        MsgBox "1St File Selected" + vbCrLf + COMPARE_FILE1
        
        Exit Sub
    End If
    
    If COMPARE_FILE2 = "" Then
        DA1$ = ScanPath.ListView1.ListItems.Item(ListView1.SelectedItem.Index).SubItems(1)
        DB1$ = ScanPath.ListView1.ListItems.Item(ListView1.SelectedItem.Index)
        
        COMPARE_FILE2 = """" + DA1$ + DB1$ + """"
    
        
        'Exit Sub
    End If
    
    Shell "C:/Program Files/WinMerge/WinMergeU.exe " + COMPARE_FILE1 + " " + COMPARE_FILE2, vbNormalFocus
    COMPARE_STATE = False
    COMPARE_FILE1 = ""
    COMPARE_FILE2 = ""

End If

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

'ColumnHeader

If WORK_DONE = False Then Exit Sub

'.ColumnHeaders.Add , "FILE", "File", 3000, lvwColumnLeft
'.ColumnHeaders.Add , "PATH", "Path", 5000, lvwColumnLeft
'.ColumnHeaders.Add , "DOS-8.3", "Dos", 1500, lvwColumnLeft
'.ColumnHeaders.Add , "SIZE", "Size", 2200, lvwColumnLeft
'.ColumnHeaders.Add , "DELETED", "DELETED", 2500, lvwColumnLeft

'Debug.Print ColumnHeader

For i = 0 To ListView1.ColumnHeaders.Count - 1
    
    If ColumnHeader = ListView1.ColumnHeaders.Item(i + 1) Then
        If ScanPath.ListView1.SortOrder = lvwAscending Then
            ScanPath.ListView1.SortOrder = lvwDescending
        Else
            ScanPath.ListView1.SortOrder = lvwAscending
        End If
        ScanPath.ListView1.SortKey = i
        ScanPath.ListView1.Sorted = True
        ScanPath.ListView1.Sorted = False
        Exit Sub
    End If

Next



End Sub


Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim DA1$
    Dim DB1$
    
    If COMPARE_STATE = True Then Exit Sub
    
    DA1$ = ScanPath.ListView1.ListItems.Item(ListView1.SelectedItem.Index).SubItems(1)
    DB1$ = ScanPath.ListView1.ListItems.Item(ListView1.SelectedItem.Index)
    
    If Mnu_Explore_Path.Checked = True Then
        If DA1$ = "" Then Exit Sub
        Shell "explorer """ + DA1$ + """", vbNormalFocus
        Exit Sub
    End If

    If Mnu_Explore_File_Select.Checked = True Then
        If DA1$ = "" Then Exit Sub
        If DB1$ = "" Then Exit Sub
        Shell "explorer /select,""" + DA1$ + DB1$ + """", vbNormalFocus
        Exit Sub
    End If


End Sub

Private Sub MNU_COMPARE_Click()

MsgBox "SELECT TWO FILES FROM LISTVIEW1 AND THEN COMPARE MERGE - HAPPEN"

COMPARE_STATE = True

End Sub


Private Sub MNU_Check_FOR_NULL_STRING_Click()

MNU_Check_FOR_NULL_STRING.Checked = Not MNU_Check_FOR_NULL_STRING.Checked
Option_MNU_Check_FOR_NULL_STRING = True

Call Checks_Dupe1

Option_Mnu_Any_Selected = True

End Sub

Private Sub MNU_CRC_Check_After_First_Click()

MNU_CRC_Check_After_First.Checked = Not MNU_CRC_Check_After_First.Checked
Option_MNU_CRC_Check_After_First = True
Call Checks_Dupe1
Option_Mnu_Any_Selected = True

End Sub

Public Sub mnu_Dont_Del_From_First_Source_Folder_Click()

mnu_Dont_Del_From_First_Source_Folder.Checked = Not mnu_Dont_Del_From_First_Source_Folder.Checked
Call Checks_Dupe1
Option_Mnu_Dont_Del_From_First_Source_Folder1 = True
Option_Mnu_Any_Selected = True

End Sub

Sub Checks_Dupe1()

Checks_Dupe_Escape = 1
Check1.Value = mnu_Dont_Del_From_First_Source_Folder.Checked And 1
Check2.Value = MNU_CRC_Check_After_First.Checked And 1
Check3.Value = Mnu_Simulate.Checked And 1
Checks_Dupe_Escape = 0
Option_Mnu_Any_Selected = True


End Sub

Sub Checks_Dupe2()
Option_Mnu_Any_Selected = True
If Checks_Dupe_Escape = 1 Then Exit Sub

mnu_Dont_Del_From_First_Source_Folder.Checked = Check1.Value * -1
MNU_CRC_Check_After_First.Checked = Check2.Value * -1
Mnu_Simulate.Checked = Check3.Value * -1



End Sub

Private Sub Mnu_Explore_File_Select_Click()
Mnu_Explore_File_Select.Checked = Not Mnu_Explore_File_Select.Checked
Mnu_Explore_Path.Checked = False
End Sub

Private Sub Mnu_Explore_Path_Click()
Mnu_Explore_Path.Checked = Not Mnu_Explore_Path.Checked
Mnu_Explore_File_Select.Checked = False

End Sub

Private Sub Mnu_Explorer_Click()
    
    Shell "explorer """ + txtPath + """", vbNormalFocus

End Sub

Private Sub Mnu_Go_Note_Click()
 If Dir(App.Path + "Menu Options 1.txt") <> "" Then
 Kill App.Path + "Menu Options 1.txt"
 End If
End Sub

Public Sub Mnu_NotePad_Logg_Click()

vfile = App.Path + "\#_DATA\Visual Basic Code Loggs.txt"

vfile = Filename_Script

If vfile = "" Then
    vfile = ScanPath.txtPath
    If Dir(vfile, vbDirectory) <> "" Then
        OPTION2 = True
    End If
End If


If Dir(Filename_Script) <> "" Then
    
    'Me.WindowState = vbMinimized
    
    If Dir("C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe") <> "" Then
        Shell "C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe """ + vfile + """", vbNormalFocus
        Exit Sub
    End If
    
    If Dir("C:\Program Files\TextView\Textview.exe") <> "" Then
        Shell "C:\Program Files\TextView\Textview.exe """ + vfile + """", vbNormalFocus
    Else
        MsgBox "Look for TextView Here" + vbCrLf + "C:\Program Files\TextView\Textview.exe" + vbCrLf + "Or Else You Use Notepad - ShellExecute"
    End If
    
    ShellExecute Me.hWnd, "open", vfile, vbNullString, vbNullString, conSwNormal

End If

If OPTION2 = True Then
    Call Mnu_Explorer_Click
End If

End Sub

Private Sub Mnu_Repeat_Questions_Off_Click()

Var_Repeat_Questions_Off = True

End Sub

Private Sub Mnu_Repeat_Questions_Toggle_Click()

Mnu_Repeat_Questions_Toggle.Checked = Not Mnu_Repeat_Questions_Toggle.Checked

If Mnu_Repeat_Questions_Toggle.Checked Then
    Mnu_Repeat_Questions_Toggle.Caption = "Answer Question Message Boxes Toggle Checked = True"
Else
    Mnu_Repeat_Questions_Toggle.Caption = "Answer Question Message Boxes Toggle Checked = False"
End If


FR1 = FreeFile
Open App.Path + "\Repeat_Questions_Toggle_Flag.txt" For Output As #FR1
Print #FR1, Mnu_Repeat_Questions_Toggle.Checked
Close #FR1


End Sub

Private Sub Mnu_Simulate_SET_CHECKED_Click()
    'TEST MODE
    
    Mnu_Simulate.Checked = True
    Option_Mnu_Simulate = True
    Call Checks_Dupe1
    Option_Mnu_Any_Selected = True

End Sub


Private Sub Mnu_Simulate_Click()
    
    Mnu_Simulate.Checked = Not Mnu_Simulate.Checked
    Option_Mnu_Simulate = True
    Call Checks_Dupe1
    Option_Mnu_Any_Selected = True

End Sub

Private Sub MNU_VB_ME_Click()

Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
End

End Sub

Private Sub SP_DirMatch(Directory As String, DDirectory As String, Path As String)
    '####################################################################################################
    'This Event fires for each Folder found
    '####################################################################################################
End Sub

Private Sub SP_DirMatchxx(Directory As String, DDirectory As String, Path As String)
    '####################################################################################################
    'This Event fires for each Folder found
    '####################################################################################################



Dim LV As ListItem
    Dim uWIN32 As WIN32_FIND_DATA
    
    '####################################################################################################
    'This Event fires for each File found
    '####################################################################################################
    
    On Local Error GoTo GetFileError
    
    With ListView1
        Set LV = .ListItems.Add(, , Directory)
        LV.SubItems(1) = Path
        LV.SubItems(2) = DDirectory
        
        
        
        
        
        
        
        
        
        If chkCopyMemory.Value Then
            'VB does not allows UDT's to be passed directly from a Class but we can access the structure
            'using pointers. We use CopyMemory to place the WIN32_FIND_DATA variable from the Class into
            'the uWIN32 declared in this Sub.
            
            'This is primarily a speed trick - we could retrieve the date & time information "ourselves"
            'using the standard VB functions or an API call but since we already have it...
            
            'CopyMemory ByVal uWIN32, ByVal SP.WFDPointer, ByVal Len(uWIN32)
            'This is not working in EXE but does in IDE Revert back to Old Way
            'CopyMemory ByVal uWIN32, ByVal SP.WFDPointer, ByVal Len(uWIN32)
                
            'Remember Public FS
            If FileName <> "" Then
                Set F = Fs.getfile(Path + DFilename)
                adate1 = F.datelastmodified
                asize1 = F.Size
            Else
                Set F = Fs.GetFolder(Path)
                adate1 = F.datelastmodified
                'We May want this later Slow Down is Massive
                If OldFolder <> Path Then
                asize1 = F.Size
                Else
                asize1 = OldSize
                End If
                OldFolder = Path
                OldSize = asize1
            End If

'            LV.SubItems(2) = uWIN32.nFileSizeLow
'            LV.SubItems(3) = FormatFileDate(uWIN32.ftLastWriteTime)
            LV.SubItems(3) = asize1
            LV.SubItems(4) = Format(adate1, "YYYY-MM-DD HH-MM-SS")
        End If
        
        'Refresh Count & Scroll the list every 10th file
        If (LV.Index Mod 10) = 0 Then
            If chkRefreshListView.Value Then
                .SelectedItem = LV
                .SelectedItem.EnsureVisible
            End If
            
            lblCount1.Caption = LV.Index
        End If
    End With
    Exit Sub

GetFileError:
    MsgBox Err.Description, vbCritical, ScanPath.Caption
    Stop


End Sub


Private Sub SP_FileMatch(FileName As String, DFilename As String, Path As String, dFSize As Double)
    Dim LV As ListItem
    Dim uWIN32 As WIN32_FIND_DATA
    
    '####################################################################################################
    'This Event fires for each File found
    '####################################################################################################
    
    On Local Error GoTo GetFileError
    
    XAG = 1
    If USE_SCAN_REFINE = True Then
        For we = 1 To UBound(SCAN_REFINE1)
            If InStr(UCase(Path), SCAN_REFINE1(we)) Then Exit Sub
        Next
    End If
    
    '-------------------------------------------------------------
    'TEST REMOVE SMALL FILE SO CAN WORK LARGE MEMORY OVERFLOW
    '-------------------------------------------------------------
    'TAKE TO LONG THIS WAY
    'SEEM READ IN QUICKER THAN REMOVE
    '-------------------------------------------------------------
    'ANOTHER METHOD AS READ IN
    '-------------------------------------------------------------
    
'    If Test = Test + 10 Then
'        If dFSize < 310000000 Then Exit Sub '310,000,000
'    End If
    
    
    
    With ListView1
        Set LV = .ListItems.Add(, , FileName)
        LV.SubItems(1) = Path
        'mod by matt l
        LV.SubItems(2) = DFilename
        
        
        
        
        LV.SubItems(3) = dFSize
        
        
        
        
        
        
        
        
        
        
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
            If FileName <> "" Then
                Set F = Fs.getfile(Path + DFilename)
                adate1 = F.datelastmodified
                asize1 = F.Size
            Else
                Set F = Fs.GetFolder(Path)
                adate1 = F.datelastmodified
                'We May want this later Slow Down is Massive
                If OldFolder <> Path Then
                asize1 = F.Size
                Else
                asize1 = OldSize
                End If
                OldFolder = Path
                OldSize = asize1
            End If


        
        End If
        End With
        
        'Refresh Count & Scroll the list every 10th file
        If (LV.Index Mod 10) = 0 Then
            If chkRefreshListView.Value Then
                With ListView1
                    .SelectedItem = LV
                    .SelectedItem.EnsureVisible
                End With
            
            End If
            
            lblCount1.Caption = LV.Index
        End If
    
        If OLDlblCount6 <> BIGGEST_FILESIZE Then
            'lblCount6.Caption = Str(BIGGEST_FILESIZE) + " Biggest File"
            lblCount6.Caption = Format(BIGGEST_FILESIZE, "###,###,###,##0") + " Biggest Scanning FSize"
            OLDlblCount6 = BIGGEST_FILESIZE
            'lblCount6.Caption = Str(OLDlblCount6) + " Biggest File"
            ScanPath.lblCount_BIGGEST_WORKING_SIZE = Format(BIGGEST_FILESIZE, "###,###,###,##0") + " Biggest Worker FSize"
    
    
'            With ListView2
'            Set LV = .ListItems.Add(, , FileName)
'                LV.SubItems(1) = Path
'                .SelectedItem = LV
'                .SelectedItem.EnsureVisible
'            End With
        End If
    
    
    
    Exit Sub

GetFileError:
    MsgBox Err.Description, vbCritical, ScanPath.Caption
    Stop
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
Drived2$ = Mid$(txtPath.Text, 1, 3)
MkDir Drived2$ + "Temp\Anything"
v1$ = Mid$(txtPath.Text, InStrRev(txtPath.Text, "\"))

MkDir Drived2$ + "Temp\Anything" + v1$
On Local Error GoTo 0

List1.AddItem "Stage 1 of 2 : Make Dir's And Move Files to Temp"
List1.AddItem "------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For we = 1 To ListView1.ListItems.Count
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(txtPath.Text)
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
                i = MsgBox("Error Making Temp Directory", , ScanPath.Caption)
                Stop
            End If
            On Local Error GoTo 0
        End If
    Loop Until ets2 = 0

    f1$ = d2$

    errs2 = 0
    On Local Error GoTo jeep
    Fs.moveFile A1$ + B1$, C1$ + B1$
    On Local Error GoTo 0

    If errs2 <> 0 Then
        i = MsgBox("Error Moving File to Temp", , ScanPath.Caption)
        End
    End If

    List1.AddItem Format$(we, "000 ") + "Move > " + A1$ + B1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'Fs.DeleteFolder comrad$ + "\" + WFD$

Next

Err.Clear
On Local Error Resume Next
Fs.Deletefolder txtPath.Text, True
If Err.Number > 0 Then Call HardDel
On Local Error GoTo 0


List1.AddItem "----------------------------------------"
List1.AddItem "Stage 2 of 2 : Move Files Back From Temp"
List1.AddItem "----------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

v2$ = Mid$(txtPath.Text, 1, InStrRev(txtPath.Text, "\"))

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

Drived2$ = Mid$(txtPath.Text, 1, 3)
v1$ = Mid$(txtPath.Text, InStrRev(txtPath.Text, "\"))


List1.AddItem "Stage Opp : Del Original Dir"
List1.AddItem "------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For we = 1 To ListView1.ListItems.Count
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(txtPath.Text)
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
            Fs.Deletefolder d2$, True
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

    List1.AddItem Format$(we, "000 ") + "Move > " + A1$ + B1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'fs.DeleteFolder comrad$ + "\" + WFD$

Next

A = A
jeep:
End Sub


Sub HardMove()

Drived2$ = Mid$(txtPath.Text, 1, 3)
v1$ = Mid$(txtPath.Text, InStrRev(txtPath.Text, "\"))


List1.AddItem "Stage Opp : Del Original Dir"
List1.AddItem "------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For we = 1 To ListView1.ListItems.Count
    A1$ = ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ListView1.ListItems.Item(we)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(txtPath.Text)
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
    
    Fs.moveFile c2$ + B1$, A1$ + B1$
    'err.number
    'err.description
    On Local Error GoTo 0

    'If errs2 <> 0 Then
    '    MsgBox ("Error Moving File to Temp")
    '    End
    'End If

    List1.AddItem Format$(we, "000 ") + "Move > " + A1$ + B1$
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'fs.DeleteFolder comrad$ + "\" + WFD$

Next
    
    v1$ = Mid$(txtPath.Text, 1, InStrRev(txtPath.Text, "\") - 1)
    Err.Clear
    On Local Error Resume Next
    
    Fs.copyfolder Drived2$ + "Temp\Anything", v1$, True
    
    Fs.Deletefolder Drived2$ + "Temp\Anything", True


jeep:
End Sub



