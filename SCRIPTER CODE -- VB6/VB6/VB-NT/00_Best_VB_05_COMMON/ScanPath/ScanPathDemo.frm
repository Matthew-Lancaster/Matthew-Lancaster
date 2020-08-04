VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ScanPath 
   BackColor       =   &H00808080&
   Caption         =   "ScanPath 2.0 - Sort  Anything -"
   ClientHeight    =   8124
   ClientLeft      =   60
   ClientTop       =   36
   ClientWidth     =   12576
   Icon            =   "ScanPathDemo.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "ScanPathDemo.frx":08CA
   ScaleHeight     =   8124
   ScaleWidth      =   12576
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ListView ListView1 
      Height          =   2055
      Left            =   0
      TabIndex        =   13
      Top             =   4320
      Width           =   6375
      _ExtentX        =   11261
      _ExtentY        =   3641
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CheckBox chk_LIST_VIEW_SHORT_5 
      Caption         =   "LIST VIEW SHORT 5"
      Height          =   195
      Left            =   1920
      TabIndex        =   21
      Top             =   1200
      Width           =   2565
   End
   Begin VB.CommandButton cmdScan_NO_LIST 
      Caption         =   "Scan Files NO LIST"
      Height          =   645
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2250
      Width           =   870
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Del Emptys"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13260
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   2010
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Find Replace .LNK'S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10875
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   1515
      Width           =   1320
   End
   Begin VB.CommandButton Command13 
      Caption         =   "CALL CRC DUPE MULTI MATRIX ON SMALLEST FILE LEAVE THE BIGGEST"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   12615
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   3045
      Width           =   2205
   End
   Begin VB.TextBox txtPath02 
      Height          =   315
      Left            =   1215
      TabIndex        =   61
      Text            =   "D:\"
      Top             =   885
      Width           =   4650
   End
   Begin VB.CommandButton Command12 
      Caption         =   "..."
      Height          =   300
      Left            =   5925
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   930
      Width           =   540
   End
   Begin VB.CommandButton Command11 
      Caption         =   "PATH TREE DIFF PLACE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   11850
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   765
      Width           =   1395
   End
   Begin VB.CommandButton Command10 
      Caption         =   "USENET FROM OUTLOOK IN FOLDERS"
      Height          =   750
      Left            =   8760
      TabIndex        =   58
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "MOVE AnyWhere HardCode"
      Height          =   750
      Left            =   7800
      TabIndex        =   57
      Top             =   3000
      Width           =   960
   End
   Begin VB.CommandButton MOVE_IN_DATE_FOLDERS_JPGS 
      Caption         =   "MOVE IN DATE FOLDERS JPGS"
      Height          =   840
      Left            =   6780
      TabIndex        =   56
      Top             =   3000
      Width           =   960
   End
   Begin VB.CommandButton Command7 
      Caption         =   "LIST FILES TO CLIP httrack"
      Height          =   735
      Left            =   5610
      TabIndex        =   55
      Top             =   2955
      Width           =   1140
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Grab String"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   10875
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   30
      Width           =   930
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Files to Text List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12615
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   2520
      Width           =   2220
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Del Boy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13260
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   1500
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   585
      Left            =   10290
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   3300
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Height          =   585
      Left            =   10290
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   2700
      Width           =   1725
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Find Replace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   10890
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   765
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXTRACT ALL WINRAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   13260
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   765
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALL CRC DUPE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   13260
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   30
      Width           =   1575
   End
   Begin VB.CheckBox Chk_SortOnDate 
      Caption         =   "Sort On Date"
      Height          =   195
      Left            =   0
      TabIndex        =   37
      Top             =   3750
      Width           =   1785
   End
   Begin VB.CommandButton CmdScanDir 
      Caption         =   "Scan Dir"
      Height          =   480
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1275
      Width           =   870
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help !"
      Enabled         =   0   'False
      Height          =   825
      Left            =   8370
      Picture         =   "ScanPathDemo.frx":0E4C
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7845
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2880
      Top             =   3960
   End
   Begin VB.ComboBox cboMask 
      Height          =   288
      Left            =   405
      TabIndex        =   1
      Text            =   "*.jpg"
      Top             =   330
      Width           =   6060
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
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
         Size            =   8.4
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
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   3315
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox chkPatternMatching 
      Caption         =   "Pattern Matching"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   3090
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox chkNormal 
      Caption         =   "Normal"
      Height          =   195
      Left            =   15
      TabIndex        =   2
      Top             =   1740
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkArchive 
      Caption         =   "Archive"
      Height          =   195
      Left            =   15
      TabIndex        =   3
      Top             =   1965
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "Hidden"
      Height          =   195
      Left            =   15
      TabIndex        =   4
      Top             =   2190
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Read Only"
      Height          =   195
      Left            =   15
      TabIndex        =   5
      Top             =   2415
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CheckBox chkSystem 
      Caption         =   "System"
      Height          =   195
      Left            =   15
      TabIndex        =   6
      Top             =   2625
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1740
      Width           =   1545
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   0
      Left            =   4050
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3150
      Width           =   1365
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2805
      Width           =   1545
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   2505
      TabIndex        =   16
      Top             =   3150
      Width           =   1530
   End
   Begin VB.ComboBox cboSizeType 
      Height          =   315
      Index           =   1
      Left            =   4050
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3510
      Width           =   1365
   End
   Begin VB.TextBox txtSize 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   2505
      TabIndex        =   18
      Top             =   3510
      Width           =   1530
   End
   Begin VB.CheckBox chkRefreshListView 
      Caption         =   "Refresh ListView (slower)"
      Enabled         =   0   'False
      Height          =   195
      Left            =   10095
      TabIndex        =   10
      Top             =   7800
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.CheckBox chkCopyMemory 
      Caption         =   "(display Date && Size)"
      Height          =   195
      Left            =   0
      TabIndex        =   11
      Top             =   3540
      Value           =   1  'Checked
      Width           =   1785
   End
   Begin VB.CheckBox chkSort 
      Caption         =   "Sort Files(A-Z) while Scanning"
      Enabled         =   0   'False
      Height          =   195
      Left            =   10110
      TabIndex        =   9
      Top             =   8055
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan Files"
      Height          =   480
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1755
      Width           =   870
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   312
      Index           =   0
      Left            =   2520
      TabIndex        =   68
      Top             =   2064
      Width           =   1548
      _ExtentX        =   2709
      _ExtentY        =   572
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   152305665
      CurrentDate     =   37299
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   312
      Index           =   1
      Left            =   2520
      TabIndex        =   69
      Top             =   2424
      Width           =   1548
      _ExtentX        =   2709
      _ExtentY        =   572
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   152305665
      CurrentDate     =   37296
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   312
      Index           =   2
      Left            =   4080
      TabIndex        =   70
      Top             =   2052
      Width           =   1116
      _ExtentX        =   1969
      _ExtentY        =   572
      _Version        =   393216
      Format          =   152371202
      CurrentDate     =   37299
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
      Height          =   315
      Left            =   0
      TabIndex        =   14
      Top             =   3990
      Width           =   14955
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "SEND DESIGNATED FOLDER TO ALL BLUETOOTHS"
      Height          =   855
      Left            =   7800
      TabIndex        =   67
      Top             =   2010
      Width           =   1365
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "DONT MODIFY AUTO RUN NOW"
      Height          =   855
      Left            =   6750
      TabIndex        =   64
      Top             =   1995
      Width           =   960
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "DESTI Path 02"
      Height          =   270
      Left            =   0
      TabIndex        =   62
      Top             =   930
      Width           =   1200
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FF00&
      Caption         =   "Label6"
      Height          =   270
      Left            =   4635
      TabIndex        =   47
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Simulate Dont Del"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   11835
      TabIndex        =   46
      Top             =   30
      Width           =   1365
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCount1 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   6510
      TabIndex        =   43
      Top             =   15
      Width           =   4335
   End
   Begin VB.Label lblCount3 
      Alignment       =   2  'Center
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
      Height          =   300
      Left            =   6510
      TabIndex        =   42
      Top             =   645
      Width           =   4335
   End
   Begin VB.Label lblCount4 
      Alignment       =   2  'Center
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
      Height          =   300
      Left            =   6510
      TabIndex        =   41
      Top             =   945
      Width           =   4335
   End
   Begin VB.Label lblCount5 
      Alignment       =   2  'Center
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
      Height          =   300
      Left            =   6510
      TabIndex        =   40
      Top             =   1245
      Width           =   4335
   End
   Begin VB.Label lblCount2 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   6510
      TabIndex        =   39
      Top             =   330
      Width           =   4335
   End
   Begin VB.Label lblCount6 
      Alignment       =   2  'Center
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
      Height          =   300
      Left            =   6510
      TabIndex        =   38
      Top             =   1545
      Width           =   4335
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
      Left            =   8925
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
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1905
      TabIndex        =   33
      Top             =   1545
      Width           =   1410
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Options:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15
      TabIndex        =   26
      Top             =   2880
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mask"
      Height          =   270
      Left            =   15
      TabIndex        =   24
      Top             =   345
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Path"
      Height          =   276
      Left            =   12
      TabIndex        =   22
      Top             =   60
      Width           =   396
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Attributes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   25
      Top             =   1530
      Width           =   885
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Min"
      Height          =   195
      Index           =   0
      Left            =   1905
      TabIndex        =   30
      Top             =   3210
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "File Size"
      Height          =   195
      Left            =   1920
      TabIndex        =   28
      Top             =   2805
      Width           =   585
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Max"
      Height          =   195
      Index           =   1
      Left            =   1905
      TabIndex        =   32
      Top             =   3570
      Width           =   300
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   195
      Left            =   1905
      TabIndex        =   29
      Top             =   2085
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "File Date"
      Height          =   195
      Left            =   1905
      TabIndex        =   27
      Top             =   1770
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   1905
      TabIndex        =   31
      Top             =   2385
      Width           =   195
   End
   Begin VB.Menu Mnu_Cmd 
      Caption         =   "Commands"
      Begin VB.Menu MNU_VB 
         Caption         =   "VB ME"
      End
      Begin VB.Menu Mnu_Com1 
         Caption         =   "Com1"
      End
      Begin VB.Menu Mnu_MatchBungDel 
         Caption         =   "MATCH BUNG AND DELETE FROM MATCH COMPARES"
      End
      Begin VB.Menu Mnu_MatchBungDateDesti 
         Caption         =   "MATCH BUNG DATE DESTINATION WAY OUT DATE CHK JOB"
      End
   End
End
Attribute VB_Name = "ScanPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------Is This Sub Get on It at start      at       Bangers

Public EXIT_TRUE

Dim LAST_FILE_DATE_TIME1, LAST_FILE_DATE_TIME2


Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type



Dim OPATH
Dim DDate2 As Date
Public Dog, OutPutPath, Xdate As Date, Tdate As Date, Zdate$, Ydate As Date, OldFolder, OldSize

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

'Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long


Private Type FILETIME
    dwLowDateTime     As Long
    dwHighDateTime    As Long
End Type

Private Const MAX_PATH = 260

Private Type WIN32_FIND_DATA
    dwFileAttributes  As Long
    ftCreationTime    As FILETIME
    ftLastAccessTime  As FILETIME
    ftLastWriteTime   As FILETIME
    nFileSizeHigh     As Long
    nFileSizeLow      As Long
    dwReserved0       As Long
    dwReserved1       As Long
    cFileName         As String * MAX_PATH
    cAlternate        As String * 14
End Type

Private mWFD As WIN32_FIND_DATA

'Put This in a Bas Mobule
'Public m_CRC As clsCRC

'Public WxHex$

Public A1, B1, G1$, FF$

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


Private Sub SP_DirMatch(Directory As String, DDirectory As String, Path As String)
    '####################################################################################################
    'This Event fires for each Folder found
    '####################################################################################################
    
    If Path = "" Then Exit Sub
    
    Dim LV As ListItem

    With ListView1
    Set LV = .ListItems.Add(, , "")
        LV.SubItems(1) = Path + Directory
    End With
    
    
    'OPATH = Path

'    TL = ScanPath.ListView1.ListItems.Count
'    If TL < 1 Then Exit Sub
'    ScanPath.ListView1.ListItems(TL).EnsureVisible


Exit Sub



'--------------



'sort of PATH then FILE
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 0
    ListView1.Sorted = True
    ListView1.SortKey = 1
    ListView1.Sorted = True
    ListView1.Sorted = False
    
    
For WE = 1 To ScanPath.ListView1.ListItems.Count
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    YY$ = A1 + B1
    Print #1, YY$
Next


ScanPath.ListView1.ListItems.Clear


End Sub

Private Sub SP_DirMatchFAST(Directory As String, DDirectory As String, Path As String, DDate As Date)
    
'FileName As String, DFilename As String, Path As String, DDate As Date)
    '####################################################################################################
    'This Event fires for each Folder found
    '####################################################################################################
    
    If Path = "" Then Exit Sub
    
    If SCAN_PARTMASK <> "" And InStr(Path + Filename, SCAN_PARTMASK) = 0 Then Exit Sub
    
'    If FormatFileDate(DDate) >= LAST_FILE_DATE_TIME Then
    If DDate >= LAST_FILE_DATE_TIME Then
        LAST_FILE_DATE_TIME1 = DDate

        LAST_FILE_DATE_PATH = Path + Directory
        'If InStr(Path, "\Hot-Key") > 0 Then
        '    LAST_FILE_DATE_PATH_WEB_CAM
        'End If
        'LAST_FILE_DATE_PATH_HOT_KEY_SCREENSHOT = Path + Directory
    
    End If
    
'    If DDate >= LAST_FILE_DATE_TIME Then
'        LAST_FILE_DATE_TIME2 = DDate
'        'If InStr(Path, "\Hot-Key") > 0 Then
'            LAST_FILE_DATE_PATH_HOT_KEY_SCREENSHOT = Path + Directory
'        'End If
'    End If
    
    
    Exit Sub

    
    
    'Dim LV As ListItem

    'With ListView1
    'Set LV = .ListItems.Add(, , "")
    '    LV.SubItems(1) = Path + Directory
    'End With
    
    
    'OPATH = Path

'    TL = ScanPath.ListView1.ListItems.Count
'    If TL < 1 Then Exit Sub
'    ScanPath.ListView1.ListItems(TL).EnsureVisible


Exit Sub



'--------------



'sort of PATH then FILE
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 0
    ListView1.Sorted = True
    ListView1.SortKey = 1
    ListView1.Sorted = True
    ListView1.Sorted = False
    
    
For WE = 1 To ScanPath.ListView1.ListItems.Count
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    YY$ = A1 + B1
    Print #1, YY$
Next


ScanPath.ListView1.ListItems.Clear


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
    
    If Path = "" Then Exit Sub
    
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
            If Directory <> "" Then
                Set F = FSO.GetFolder(Path + DDirectory)
                ADATE1 = F.DateLastModified
'                ASIZE1 = Trim(Str(F.Size))
'            Else
'                Set F = FSO.GetFolder(Path)
'                If Len(Path) > 3 Then
'                    ADATE1 = F.datelastmodified
'                    'We May want this later Slow Down is Massive
'                    If OldFolder <> Path Then
'                        ASIZE1 = Trim(Str(F.Size))
'                    Else
'                        ASIZE1 = Trim(Str(OldSize))
'                    End If
'                    OldFolder = Path
'                    OldSize = ASIZE1
'                End If
            End If

'            LV.SubItems(2) = uWIN32.nFileSizeLow
'            LV.SubItems(3) = FormatFileDate(uWIN32.ftLastWriteTime)
'            LV.SubItems(3) = ASIZE1
            LV.SubItems(4) = Format(ADATE1, "YYYY-MM-DD HH-MM-SS")
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
    MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground


End Sub





Private Sub SP_FileMatch(Filename As String, DFilename As String, Path As String)
    Dim LV As ListItem
    Dim uWIN32 As WIN32_FIND_DATA
    Dim ASIZE1 As Long
    
    '####################################################################################################
    'This Event fires for each File found
    '####################################################################################################
    
    On Local Error GoTo GetFileError
    
    If Path = "" Then
        Exit Sub
    End If
    If Filename = "" Then
        Exit Sub
    End If

    

'    If XX = 58 Then
'
        'lblCount1.Caption = mFileCount2
'        Print #1, "*" + Path + Filename
'
'    Exit Sub
'    End If
    
    
    
    With ListView1
        Set LV = .ListItems.Add(, , Filename)
        LV.SubItems(1) = Path
        'mod by matt l
        LV.SubItems(2) = DFilename
        
        
        
        'lblCount1.Caption = LV.mFileCount2
        
        
        
        
        
        
        
        
        If chkCopyMemory.Value Then
            'VB does not allows UDT's to be passed directly from a Class but we can access the structure
            'using pointers. We use CopyMemory to place the WIN32_FIND_DATA variable from the Class into
            'the uWIN32 declared in this Sub.
            
            'This is primarily a speed trick - we could retrieve the date & time information "ourselves"
            'using the standard VB functions or an API call but since we already have it...
            
            'CopyMemory ByVal uWIN32, ByVal SP.WFDPointer, ByVal Len(uWIN32)
            'LV.SubItems(3) = uWIN32.nFileSizeLow
            'LV.SubItems(4) = FormatFileDate(uWIN32.ftLastWriteTime)
        'FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                    'Remember Public FS
            If Filename <> "" Then
                Set F = FSO.GetFile(Path + DFilename)
                ADATE1 = F.DateLastModified
                ASIZE1 = F.Size
'                If Val(ASIZE1) = 647732 Then Stop
            Else  ' IF FOLDER
                
                
                
'                Set F = FSO.GetFolder(Path)
'                ADATE1 = F.datelastmodified
'
'                'FOLDER SIZE BUT NOT IF REQUEST AND WHEN MULTI SAME SIZE
'
'
'                'We May want this later Slow Down is Massive
'                If OldFolder <> Path Then
'                    ASIZE1 = F.Size
'                Else
'                    ASIZE1 = OldSize
'                End If
'                OldFolder = Path
'                OldSize = ASIZE1
            End If
'            LV.SubItems(2) = uWIN32.nFileSizeLow
'            LV.SubItems(3) = FormatFileDate(uWIN32.ftLastWriteTime)
            LV.SubItems(3) = ASIZE1
            LV.SubItems(4) = Format(ADATE1, "YYYY-MM-DD HH-MM-SS")


        
        End If
        
        
        If Me.Caption <> "ScanPath 2.0 - Sort  Anything -ClipBoard Logger" Then
            lblCount1.Caption = ScanPath.ListView1.ListItems.Count
        
            'Refresh Count & Scroll the list every 10th file
            If (LV.Index Mod 10) = 0 Then
                If chkRefreshListView.Value Then
                    .SelectedItem = LV
                    .SelectedItem.EnsureVisible
                End If
                
                'lblCount1.Caption = LV.Index
            End If
        End If
    End With
    Exit Sub

GetFileError:
    MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground
End Sub

Private Sub SP_FileMatchFast(Filename As String, DFilename As String, Path As String, DDate As Date)
    
'    Dim LAST_FILE_DATE_TIME1, LAST_FILE_DATE_TIME2
    
    'Dim LV As ListItem
    'Dim uWIN32 As WIN32_FIND_DATA
    
    'Dim ASIZE1 As Long
    
    '####################################################################################################
    'This Event fires for each File found
    '####################################################################################################
    
    On Local Error GoTo GetFileError
    
    'If Path = "" Then Exit Sub
    'If fileName = "" Then Exit Sub
    
If SCAN_PARTMASK <> "" And InStr(Path + Filename, SCAN_PARTMASK) = 0 Then Exit Sub
    
'    If FormatFileDate(DDate) >= LAST_FILE_DATE_TIME Then
    If DDate >= LAST_FILE_DATE_TIME Then
        LAST_FILE_DATE_TIME1 = DDate

        LAST_FILE_DATE_PATH = Path + Filename
        'If InStr(Path, "\Hot-Key") > 0 Then
        '    LAST_FILE_DATE_PATH_WEB_CAM
        'End If
    End If
    
'    If DDate >= LAST_FILE_DATE_TIME Then
'        LAST_FILE_DATE_TIME2 = DDate
'        'If InStr(Path, "\Hot-Key") > 0 Then
'            LAST_FILE_DATE_PATH_HOT_KEY_SCREENSHOT = Path + FileName
'        'End If
'    End If
    
    
    Exit Sub

GetFileError:
    MsgBox Err.Description, vbCritical + vbMsgBoxSetForeground
End Sub



Private Function FormatFileDate(CT As FILETIME) As String
    
    Const SHORT_DATE = "Short Date"
    
    Dim ST As SYSTEMTIME
    Dim ds1 As Single
    Dim ds2 As Single
       
    'If FileTimeToSystemTime(CT, ST) Then
    '      ds1 = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
    '      ds2 = TimeSerial(ST.wHour, ST.wMinute, ST.wSecond)
          
    '      FormatFileDate = Format$(ds1 + ds2, "YYYY-MM-DD HH-MM-SS")
    'End If

End Function


Private Sub Timer1_Timer()
Timer1.Enabled = False
Exit Sub
'End
End Sub






'
'
'Sub CRC_SCRIPT_OF_FILES()
'
'Set FS = CreateObject("Scripting.FileSystemObject")
'
'Set m_CRC = New clsCRC
'm_CRC.Algorithm = CRC32
'
'ScanPath.ListView1.ListItems.Clear
'ScanPath.TxtPath = "G:\"
'ScanPath.cboMask.Text = "*.*"
'ScanPath.chkSubFolders = vbChecked
''ScanPath.chkCopyMemory.Value = vbUnchecked
'
''Call ScanPath.cmdScanDir_Click
'Call ScanPath.cmdScan_Click
'ScanPath.Show
'Reset
'On Error Resume Next
''Kill "D:\03 MICROSOFT\# VB 6\CRC CHECK LIST OF VS6.txt"
''Kill "D:\03 MICROSOFT\# VB 6\CRC CHECK LIST OF VS6 TOTAL DISK.txt"
'On Error GoTo 0
'FR1 = FreeFile
'Open "D:\03 MICROSOFT\# VB 6\CRC CHECK LIST OF VS6.txt" For Append As #FR1
'WE = ScanPath.ListView1.ListItems.Count
'ScanPath.lblCount1 = Trim(Str(WE))
'Print #FR1, "#----------------------------------"
'Print #FR1, "# FILE COUNT OF DISK ="; Str(WE)
'Print #FR1, "# COUNT--CRC32--SIZE--FILENAME"
'Print #FR1, "#----------------------------------"
'
'Dim A2 As String
'Dim H2 As String
'Dim VC As Long
'For WE = 1 To ScanPath.ListView1.ListItems.Count
''    we = ScanPath.ListView1.ListItems.Count
'    ScanPath.lblCount2 = Trim(Str(WE))
'    DoEvents
'
'    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
'    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
'    A1 = A1 + ScanPath.ListView1.ListItems.Item(WE)
'    A2 = A1
'    Set F = FSO.getfile(A2)
'    VC = VC + F.Size
'    a3 = Mid(A1, 3)
'    H1 = Hex$(m_CRC.CalculateFile(A2))
'    'If H1 = "0" Then Stop
'    LH = "00000000"
'    RSet LH = H1
'    LH = Replace(LH, " ", "0", , -1)
'    D1 = Format(WE, "00000")
'    D2 = Format(F.Size, "000,000,000")
'    Print #FR1, D1; ","; LH; ","; D2 + "," + a3
'    H2 = H2 + "-" + LH
'
'Next
'
'Close #FR1
'Set F = FSO.GETDrive(ScanPath.TxtPath)
'
'FR1 = FreeFile
'Open "D:\03 MICROSOFT\# VB 6\CRC CHECK LIST OF VS6 TOTAL DISK.txt" For Append As #FR1
'Print #FR1, "-------------------------------"
'Print #FR1, "VOLUME LABEL = " + F.VolumeName
'Print #FR1, "SCAN PATH = " + ScanPath.TxtPath
'Print #FR1, "TOTAL ADDED FILE SIZE = " + Format(VC, "0,000,000,000")
'Print #FR1, "FILE FILE COUNT OF DISK ="; Str(ScanPath.ListView1.ListItems.Count)
'LH = "00000000"
'RSet LH = Hex(m_CRC.CalculateString(H2))
'LH = Replace(LH, " ", "0", , -1)
'Print #FR1, "CRC OF ALL FILES ADDED IN A CRC STRING TO OBTAIN CRC32 = "; LH
''Print #FR1, "CRC OF ALL FILES ADDED IN CRC STRING SOURCE= "; H2
'Close #FR1
'End Sub
'
'



Sub DelJunkFrontPage()


ScanPath.ListView1.ListItems.Clear
ScanPath.TxtPath = "D:\MY WEBS\MatthewLan.Com Web\"
ScanPath.cboMask.Text = "*.*"
ScanPath.chkSubFolders = vbChecked
Call ScanPath.cmdScanDir_Click
ScanPath.Show
Reset
For WE = ScanPath.ListView1.ListItems.Count To 1 Step -1
'    we = ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount1 = Trim(Str(WE))
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    A1 = A1 + ScanPath.ListView1.ListItems.Item(WE)
    outff = 0
    If InStr(A1, "_private") > 0 Then outff = 1
    If InStr(A1, "_vti") > 0 Then
        outff = 1
    End If
    
    If outff = 1 Then
        'ScanPath.ListView1.ListItems.Remove (we)
        On Error Resume Next
        FSO.DeleteFolder A1, True
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

For WE = ListView1.ListItems.Count To 1 Step -1
    
    Label13 = ListView1.ListItems.Count
    DoEvents
    
    G1$ = ListView1.ListItems.Item(WE).SubItems(2)
    A1 = ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ListView1.ListItems.Item(WE)
    
    Set F = FSO.GetFile(A1 + G1$)
    Tdate = F.DateLastModified
    If Tdate > Xdate Then Xdate = Tdate
    
    remo = 0
    If Tdate <= Ydate Then
        ListView1.ListItems.Remove (WE)
        remo = 1
    End If
    If remo = 0 Then
        hole = "Inet " + Format$(Now, "YYYY-MM-DD") + "\"
        GeText = LCase(Mid$(B1, Len(B1) - 2, 3)) + "\"
        If GeText = "jpg\" Then GeText = ""
        secext = "M:\0 00 Art\Temp Inet Files JPGs\" + hole + GeText + B1
        If FSO.FileExists(secext) = True Then
            ListView1.ListItems.Remove (WE)
        End If
        secext = OutPutPath + hole + GeText + Mid$(B1, 1, Len(B1) - 4) + ".zip"
        If FSO.FileExists(secext) = True Then
            ListView1.ListItems.Remove (WE)
        End If
        secext = OutPutPath + hole + GeText + Mid$(B1, 1, Len(B1) - 4) + ".rar"
        If FSO.FileExists(secext) = True Then
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

For WE = ListView1.ListItems.Count To 1 Step -1
    
    Label13 = ListView1.ListItems.Count
    DoEvents
    
    G1$ = ListView1.ListItems.Item(WE).SubItems(2)
    A1 = ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ListView1.ListItems.Item(WE)
    
    Set F = FSO.GetFile(A1 + G1$)
    
    
    GeText = LCase(Mid$(B1, Len(B1) - 2, 3)) + "\"
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

For WE = ListView1.ListItems.Count To 1 Step -1
    
    Label13 = ListView1.ListItems.Count
    DoEvents
    
    G1$ = ListView1.ListItems.Item(WE).SubItems(2)
    A1 = ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ListView1.ListItems.Item(WE)
    Set F = FSO.GetFile(A1 + G1$)
    Tdate = F.DateLastModified
    If Tdate > Xdate Then Xdate = Tdate
    remo = 0
    If Tdate <= Ydate Then
        ListView1.ListItems.Remove (WE)
        remo = 1
    End If
    If remo = 0 Then
        GeText = "#" + LCase(Mid$(B1, Len(B1) - 2, 3)) + "\"
        If FSO.FileExists(OutPutPath + GeText + B1) = True Then
        ListView1.ListItems.Remove (WE)
        End If
    End If
Next

Open App.Path + "\Xdate.txt" For Output As #1
Print #1, Xdate
Close #1


For WE = ListView1.ListItems.Count To 1 Step -1
    
    Label13 = ListView1.ListItems.Count
    DoEvents
    
    G1$ = ListView1.ListItems.Item(WE).SubItems(2)
    A1 = ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ListView1.ListItems.Item(WE)
    
    Set F = FSO.GetFile(A1 + G1$)

    GeText = "#" + LCase(Mid$(B1, Len(B1) - 2, 3)) + "\"
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
ra(2) = "M:\Videos\00 Vid Films"
ra(3) = "M:\Videos\00 Vid XXX"
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

For WE = ListView1.ListItems.Count To 1 Step -1
    
    Label13 = WE
    DoEvents
    
    G1$ = ListView1.ListItems.Item(WE).SubItems(2)
    A1 = ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ListView1.ListItems.Item(WE)
    W = DeleteFile(A1 + G1$)

Next

Call cmdScanDir_Click

For WE = ListView1.ListItems.Count To 1 Step -1
  
    Label14 = WE
    DoEvents
    
    A1 = ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ListView1.ListItems.Item(WE)
    A1 = A1 + B1 + "\"
  
    
    Set F = FSO.GetFolder(A1)
    
    If F.Size = 0 Then
        ListView1.ListItems.Remove (WE)
    End If
Next

A1 = TxtPath.Text

Call Tagg2


For WE = 1 To ListView1.ListItems.Count
  
    Label16 = WE
    DoEvents
    
    A1 = ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ListView1.ListItems.Item(WE)
    A1 = A1 + B1 + "\"
    
    If InStr(A1, " - 20") > 0 And InStr(LCase(A1), "banging") > 0 Then
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
 
 

FF$ = Mid$(A1, Len(TxtPath.Text) + 1)
If FF$ = "" Then
FF$ = Mid$(A1, 4)
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
If InStr(A1, "My Music\") Then TY$ = "My Music"
If InStr(A1, "My Music Zen\") Then TY$ = "My Music Zen"
If TY$ = "My Music" And InStr(A1, "Favs") Then TY$ = "My Music - Favs"
If TY$ = "My Music Zen" And InStr(A1, "Favs") Then TY$ = "My Music Zen - Favs"
If TY$ = "My Music" And InStr(A1, " Bang Sets") Then TY$ = "My Music - Sets"
If TY$ = "My Music Zen" And InStr(A1, " Bang Sets") Then TY$ = "My Music Zen - Sets"
If TY$ = "My Music" And InStr(A1, " Bang Fav Sets") Then TY$ = "My Music - Fav Sets"
If TY$ = "My Music Zen" And InStr(A1, " Bang Fav Sets") Then TY$ = "My Music Zen - Fav Sets"
'03 Bang Sets
    
    
'GH$ = Mid$(A1, InStrRev(A1, "\", Len(A1) - 1) + 1)
Tx$ = Left$(GH$, Len(GH$) - 1)

FF$ = "---------------[[(({{ " + Tx$ + " - " + TY$ + " }}))]]---------------.mp3"


On Error Resume Next
Do
fg1 = FreeFile
Err.Clear
Open A1 + FF$ For Output As fg1
Close fg1
If Err.Number > 0 Then FF$ = Left$(FF$, Len(FF$) - 5) + ".mp3"
Loop Until Err.Number = 0


'Call WriteTagg



End Sub



Sub WriteTagg()

On Local Error Resume Next
'Kill A1 + "0#-" + FF$ + "-^.mp3"
'Kill A1 + "01#-" + FF$ + "-^.mp3"
'Err.Clear
easy = Dir$(A1 + "\*.txt") <> ""
 '
Do
    Err.Clear
'    If InStr(FF$, "Art And Music") > 0 Then Stop

    fg1 = FreeFile
    Open A1 + "0#01--" + FF$ + "-^.mp3" For Output As fg1
    Close fg1
    FF$ = Left$(FF$, Len(FF$) - 1)
Loop Until Err.Number = 0
    
Do
    Err.Clear
    ek$ = String$(Len(FF$) + 20, "-")
    fg1 = FreeFile
    Open A1 + "0#00--" + ek$ + "-^.mp3" For Output As fg1
    Close fg1
    If easy = True Then
        fg1 = FreeFile
        Open A1 + "0#02----- Text File In Dir -----^.mp3" For Output As fg1
        Close fg1
    End If
    fg1 = FreeFile
    Open A1 + "0#02--" + ek$ + "-^.mp3" For Output As fg1
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

End Sub

Private Sub chkCopyMemory_Click()
    
    chkSort.Enabled = (chkCopyMemory.Value = vbUnchecked)
    
End Sub


Private Sub chkSort_Click()
    chkCopyMemory.Enabled = (chkSort.Value = vbUnchecked)
End Sub

'Private Sub cmdBrowse_Click()
'    TxtPath.Text = GetFolder(Me.hWnd, "Scan Path:", TxtPath.Text)
'    fg1 = FreeFile
'    Open App.Path + "\Scan Path.txt" For Output As #fg1
'    Print #fg1, TxtPath.Text
'    Close #fg1
'
''    fg1 = FreeFile
''    Open App.Path + "\Scan Path.txt" For Input As #fg1
''    Line Input #fg1, ae$
''    Close #fg1
''    txtPath.Text = ae$
'
'
'
'End Sub

Private Sub cmdHelp_Click()
    'txtHelp.Visible = Not txtHelp.Visible
End Sub


Public Sub cmdScan_Click()
    Dim lCount As Long
    Dim lSize(1) As Long
    
    chkSort.Enabled = False
    
    
    If TxtPath = "" Then
    
        MSGBOX_OPTION = vbOKCancel
        If IsIDE = False Then MSGBOX_OPTION = vbOKOnly
        
        i = MsgBox("SCAN PATH" + vbCrLf + vbCrLf + "TxtPath = """" " + vbCrLf + vbCrLf + "REQUEST TO SCAN WAS EMPTY VARIABLE" + vbCrLf + vbCrLf + "EXIT RUN", MSGBOX_OPTION, vbMsgBoxSetForeground)
        If i = vbCancel Then
            Stop
        End If
        
        Exit Sub
    
    End If
    
    
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
                MsgBox "Maximium Size value must exceed 0", vbExclamation + vbMsgBoxSetForeground
                txtSize(1).SetFocus
                Exit Sub
            ElseIf lSize(1) <= lSize(0) Then
                MsgBox "Maximium Size value must exceed Minimum Size", vbExclamation + vbMsgBoxSetForeground
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
            'LockWindowUpdate (ListView1.hWnd)
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
                ListView1.Sorted = False
                ListView1.SortKey = 1
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
        
        'LockWindowUpdate (0)
        
        cmdScan.Caption = "Scan Files"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
        'End
    End If


cmdScan.Caption = "Scan Files"




'We always use direct scan as is better to sort after job then during thats all we need

'SORTI = 0



'SORTI = IS MADE COMMON PUBLIC BECUASE UNABLE TO EASY PASSED TO THIS PROCEDURE


Select Case SORTI
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
    ListView1.Sorted = False
    ListView1.SortKey = 1
    ListView1.Sorted = True
    ListView1.Sorted = False

Case 5
'sort of FILE then PATH
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = 1
    ListView1.Sorted = True
    ListView1.Sorted = False
    ListView1.SortKey = 0
    ListView1.Sorted = True
    ListView1.Sorted = False

End Select

End Sub





Public Sub cmdScanDir_FAST_Click()
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
                MsgBox "Maximium Size value must exceed 0", vbExclamation + vbMsgBoxSetForeground
                txtSize(1).SetFocus
                Exit Sub
            ElseIf lSize(1) <= lSize(0) Then
                MsgBox "Maximium Size value must exceed Minimum Size", vbExclamation + vbMsgBoxSetForeground
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
            .StartScanDir_FAST TxtPath, chkSubFolders.Value, chkSort.Value And chkSort.Enabled, chkPatternMatching
        
            
            'Correct Yes 1st one COLUMN sort keeps sort then another
'            ListView1.SortOrder = lvwAscending
'            ListView1.SortKey = 0
'            ListView1.Sorted = True
'            ListView1.SortKey = 1
'            ListView1.Sorted = True
'            ListView1.Sorted = False
'
'            If Chk_SortOnDate.Value = vbChecked Then
'                ListView1.SortOrder = lvwAscending
'                ListView1.SortKey = 3
'                ListView1.Sorted = True
'                ListView1.Sorted = False
'            End If
        
        End With
        
'        With ListView1
'            If .ListItems.Count > 0 Then
'                .SelectedItem = .ListItems(1)
'            End If
'
'            lblCount1.Caption = .ListItems.Count
'        End With
        
        'LockWindowUpdate 0
        
        cmdScan.Caption = "Scan"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
   End If
'

lblCount1.Caption = ListView1.ListItems.Count
    

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
                MsgBox "Maximium Size value must exceed 0", vbExclamation + vbMsgBoxSetForeground
                txtSize(1).SetFocus
                Exit Sub
            ElseIf lSize(1) <= lSize(0) Then
                MsgBox "Maximium Size value must exceed Minimum Size", vbExclamation + vbMsgBoxSetForeground
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
            .StartScanDir TxtPath, chkSubFolders.Value, chkSort.Value And chkSort.Enabled, chkPatternMatching
        
            
            'Correct Yes 1st one COLUMN sort keeps sort then another
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
        
        'LockWindowUpdate 0
        
        cmdScan.Caption = "Scan"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
   End If
'

lblCount1.Caption = ListView1.ListItems.Count
    

End Sub




Private Sub Command1_Click()

'Call CRCDUPE

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

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL FILES"
    
'B1 = ScanPath.ListView1.ListItems.Item(1)
'RXRIME = Format(DateValue(Mid(B1, 8, 10)), "YYYY-MM-MMM")
'FN = ScanPath.txtPath.Text + "HTTP's From Email's -- " + RXRIME + ".txt"

'If Dir(FN) <> "" Then Kill FN
'UU = FreeFile
'Open FN For Append As #UU



For WE = 1 To ScanPath.ListView1.ListItems.Count
    
    
    
'    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " DEALT WITH FILES"
    
    DoEvents
    
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
    
    TT = TT + "-*/" + B1 + vbCrLf
    
    'Set F = FsO.getfile(A1 + B1)
    'ADATE1 = F.datelastmodified
    
    Open A1 + G1 For Binary As #1
    OO = Input(LOF(1), 1)
    Close #1
    
    
    'Clipboard.Clear
    'Clipboard.SetText oo
    
   
   ADATE1 = Mid(OO, InStr(OO, vbCrLf + "Date: ") + 2 + 6)
   ADATE1 = Mid(ADATE1, 1, InStr(ADATE1, vbCrLf))
   If InStr(ADATE1, ",") > 0 Then
        ADATE1 = Mid(ADATE1, InStr(ADATE1, ",") + 2)
   End If
   
   
   
            DDate = ADATE1
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
                UU2 = Val(UU1)
                UU = UU - TimeSerial(UU2, 0, 0)
                DDate2 = UU
            End If
            If InStr(DDate, "+") > 0 Then
                UU = DDate
                UU = Mid(DDate, 1, InStr(DDate, "+") - 2)
                UU1 = Mid(DDate, InStr(DDate, "+") + 1, 2)
                'UU = Mid(UU, 6)
                UU = DateValue(UU) + TimeValue(UU)
                UU2 = Val(UU1)
                UU = UU + TimeSerial(UU2, 0, 0)
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
    If FSO.FolderExists(ScanPath.TxtPath + yy2) = False Then MkDir ScanPath.TxtPath + yy2
    'If FsO.FileExists(ScanPath.txtPath + yy + "\" + yy2) = False Then
        
        LONGLOOP = 0
        Do
        LONGLOOP = LONGLOOP + 1
        If LONGLOOP > 50 Then Stop
        On Error Resume Next
        Set F = FSO.GetFile(A1 + G1)
        Err.Clear
        If FSO.FileExists(ScanPath.TxtPath + "\" + yy2 + "\" + YY) = True Then Exit Do
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
    'If FsO.FileExists(A1 + YY2) = False Then
        'F.Move A1 + YY2
    'End If
    
    
    
    'F.Move A1 + YY + "\" + YY + " - " + B1
    
    
    
    
    
    
Next




End Sub

Private Sub Command13_Click()

'Call CRCDUPE_MULTI_MATRIX_ON_SMALLEST_FILE

End Sub

Private Sub Command14_Click()
'Call DelEMPTYS
End Sub

Public Sub CMDScan_NO_LIST_Click()
   
    'WHATS THIS FOR
    'XX = 58
    
    Dim lCount As Long
    Dim lSize(1) As Long

    With ListView1
        .ColumnHeaders.Item(1).Width = 1
    End With
    chkSort.Enabled = False
    
    
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
                MsgBox "Maximium Size value must exceed 0", vbExclamation + vbMsgBoxSetForeground
                txtSize(1).SetFocus
                Exit Sub
            ElseIf lSize(1) <= lSize(0) Then
                MsgBox "Maximium Size value must exceed Minimum Size", vbExclamation + vbMsgBoxSetForeground
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
            
            
            
            
            
            
            
            
'            If Chk_SortOnDate.Value = vbChecked Then
'                ListView1.SortOrder = lvwAscending
'                ListView1.SortKey = 3
'                ListView1.Sorted = True
'                ListView1.Sorted = False
'            Else
'                ListView1.SortOrder = lvwAscending
'                ListView1.SortKey = 0
'                ListView1.Sorted = True
'                ListView1.SortKey = 1
'                ListView1.Sorted = True
'                ListView1.Sorted = False
'            End If
        
        End With
        
'        With ListView1
'            If .ListItems.Count > 0 Then
'                .SelectedItem = .ListItems(1)
'            End If
'
'            lblCount1.Caption = .ListItems.Count
'        End With
        
        'LockWindowUpdate 0
        
        cmdScan.Caption = "Scan Files"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
        'End
    End If

Exit Sub


cmdScan.Caption = "Scan Files"



Exit Sub


'We always use direct scan as is better to sort after job then during thats all we need

SORTI = 0

Select Case SORTI
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


Public Sub CMDScan_NO_LIST_FAST_Click()
   
    'XX = 58
    
    Dim lCount As Long
    Dim lSize(1) As Long

    With ListView1
        .ColumnHeaders.Item(1).Width = 1
    End With
    chkSort.Enabled = False
    
    
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
                MsgBox "Maximium Size value must exceed 0", vbExclamation + vbMsgBoxSetForeground
                txtSize(1).SetFocus
                Exit Sub
            ElseIf lSize(1) <= lSize(0) Then
                MsgBox "Maximium Size value must exceed Minimum Size", vbExclamation + vbMsgBoxSetForeground
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
            .StartScan_USE_NO_LIST_FAST TxtPath, chkSubFolders.Value, chkSort.Value And chkSort.Enabled, chkPatternMatching
            
            
            
            
            
            
            
            
'            If Chk_SortOnDate.Value = vbChecked Then
'                ListView1.SortOrder = lvwAscending
'                ListView1.SortKey = 3
'                ListView1.Sorted = True
'                ListView1.Sorted = False
'            Else
'                ListView1.SortOrder = lvwAscending
'                ListView1.SortKey = 0
'                ListView1.Sorted = True
'                ListView1.SortKey = 1
'                ListView1.Sorted = True
'                ListView1.Sorted = False
'            End If
        
        End With
        
'        With ListView1
'            If .ListItems.Count > 0 Then
'                .SelectedItem = .ListItems(1)
'            End If
'
'            lblCount1.Caption = .ListItems.Count
'        End With
        
        'LockWindowUpdate 0
        
        cmdScan.Caption = "Scan Files"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
        'End
    End If


    'NEED THIS BACK ON
    cmdScan.Caption = "Scan Files"

Exit Sub


Exit Sub
'-------------------------------------------------------

'We always use direct scan as is better to sort after job then during thats all we need

SORTI = 0

Select Case SORTI
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

Public Sub CMDScan_NO_LIST_FAST_FORM_CAPTURE_Click()
   
    If TxtPath = "" Then Exit Sub
   
    Dim lCount As Long
    Dim lSize(1) As Long

    With ListView1
        .ColumnHeaders.Item(1).Width = 1
    End With
    chkSort.Enabled = False
    
    
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
                MsgBox "Maximium Size value must exceed 0", vbExclamation + vbMsgBoxSetForeground
                txtSize(1).SetFocus
                Exit Sub
            ElseIf lSize(1) <= lSize(0) Then
                MsgBox "Maximium Size value must exceed Minimum Size", vbExclamation + vbMsgBoxSetForeground
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
            .StartScan_USE_NO_LIST_FAST_FORM_CAPTURE TxtPath, chkSubFolders.Value, chkSort.Value And chkSort.Enabled, chkPatternMatching
        
        End With
        
        'LockWindowUpdate 0
        
        cmdScan.Caption = "Scan Files"
        Screen.MousePointer = vbDefault
    Else
        'User want's to stop current scan
        SP.StopScan
        'End
    End If


    'NEED THIS BACK ON
    cmdScan.Caption = "Scan Files"

End Sub



Private Sub Command2_Click()
'Call WINRAR_EXTRACT

End Sub

Private Sub Command3_Click()
    'MOST HARD CODED
    'Call FIND_REPLACE
    
    MsgBox "WHAT DO YOU WANT TO DO CLEAR TEXT FOR - NOTHING IS GOING TO HAPPEN", vbMsgBoxSetForeground




End Sub

Private Sub Command4_Click()
'Call DelBoy




End Sub

Private Sub Command5_Click()

'Call FilesToTextList

End Sub

Private Sub Command6_Click()
    'MOST HARD CODED
    Text1 = "Start to Search"
    Text2 = "How Does End"
    
    'Text1 = "http:"
    'Text2 = """"
    
    'Call GRAB_STRING
    


End Sub

Private Sub Command7_Click()
Exit Sub

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

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL FILES"
    
'B1 = ScanPath.ListView1.ListItems.Item(1)
'RXRIME = Format(DateValue(Mid(B1, 8, 10)), "YYYY-MM-MMM")
'FN = ScanPath.txtPath.Text + "HTTP's From Email's -- " + RXRIME + ".txt"

'If Dir(FN) <> "" Then Kill FN
'UU = FreeFile
'Open FN For Append As #UU

For WE = 1 To ScanPath.ListView1.ListItems.Count
    
    
    
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


'EXECUTE_TIMER_ENABLED = False

Clipboard.Clear
Clipboard.SetText TT

'EXECUTE_TIMER_ENABLED = True


End Sub

Private Sub Command8_Click()

ScanPath.ListView1.ListItems.Clear
ScanPath.cboMask.Text = "*.LNK"
ScanPath.chkSubFolders = vbChecked

ScanPath.TxtPath = "E:\01 VB Shell Folders\00 Shell Notepad Loader"
ScanPath.TxtPath = "E:\01 VB Shell Folders\"

Call ScanPath.cmdScan_Click


'MY DOCS
'sPersonalFolder = GetSpecialfolder(CSIDL_PERSONAL)

ScanPath.Show

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL FILES"
    

Dim RR As String, D1, TT, D4 As String, D5, S1, D3

'Debug.Print vbCrLf

For WE = 1 To ScanPath.ListView1.ListItems.Count
    
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


D2 = 0
D2 = InStr(D1, "My Documents")
If D2 > 0 Then
D3 = InStr(D2, D1, "\")
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

If FSO.FileExists(D4) = True Or FSO.FolderExists(D4) = True Then

        CNT3 = CNT3 + 1
        ScanPath.lblCount3 = Trim(Str(CNT3)) + " DEALT WITH"
        'Debug.Print "CONVERT LINK "
        'Call Create_ShortCut(D4, A1 + B1, G1, D5)
    
    Else
    
        'MsgBox "FILE NOT EXIST" + VBCLRF + D4
        'Debug.Print "FILE NOT EXIST ------ " + A1 + B1
        'Debug.Print "FILE NOT EXIST ------ " + D4
        'Debug.Print
    
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
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL FILES"

If ScanPath.ListView1.ListItems.Count = 0 Then Exit Sub
fp = App.Path + "\TOSEND.vcf"
If Dir(fp) <> "" Then Kill fp

For WE = 1 To ScanPath.ListView1.ListItems.Count
    ScanPath.lblCount2 = Trim(Str(WE)) + " Process Files"
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
'i = ExecCmdWAIT("C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=6c:0e:0d:30:fc:86 -BD_NAME=W595i -DEV_CLASS=00262746" + " """ + SOURCE1 + """")

'K700i
ScanPath.lblCount2 = " K700 - PROCESS"
'i = ExecCmdWAIT("C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=00:fc:22:08:04:0e -BD_NAME=K700i -DEV_CLASS=00262738" + " """ + SOURCE1 + """")

'V600i
ScanPath.lblCount2 = " V600 - PROCESS"
'i = ExecCmdWAIT("C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=00:12:ee:61:65:7e -BD_NAME=V600i -DEV_CLASS=00262738" + " """ + SOURCE1 + """")

'K800i
ScanPath.lblCount2 = " K800 - PROCESS"
'i = ExecCmdWAIT("C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=00:1d:28:2b:90:9c -BD_NAME=K800i -DEV_CLASS=00262746" + " """ + SOURCE1 + """")

'W880i
ScanPath.lblCount2 = " W880 - PROCESS"
'i = ExecCmdWAIT("C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=00:1d:28:36:2a:8d -BD_NAME=W80x2D07919180203 -DEV_CLASS=00262746" + " """ + SOURCE1 + """")

    
ScanPath.lblCount6 = "DONE"

End

End Sub

Private Sub Mnu_MatchBungDateDesti_Click()


ScanPath.ListView1.ListItems.Clear

ScanPath.chkSubFolders = vbChecked


ScanPath.TxtPath = "F:\##ACER-D2"

Call ScanPath.cmdScan_Click


'destipath1 = "\\55-88-happy\MY BOOK (M)\##ASUS-D"
destipath1 = "\\55-88-happy\D"


'On Error Resume Next
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL FILES"

For WE = 1 To ScanPath.ListView1.ListItems.Count
    
    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
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
    
    
    
    If FSO.FileExists(A1 + B1) = True And FSO.FileExists(D1 + B1) = True Then
        
        Set F = FSO.GetFile(D1 + B1)
        ADATE1 = F.DateLastModified
        If ADATE1 < DateSerial(2010, 2, 1) Then
        
        
        Err.Clear
        
        'DEL FROM OUR SWAP DRIVE
        FSO.DeleteFile A1 + B1, True
            
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

'Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32

ScanPath.ListView1.ListItems.Clear

ScanPath.chkSubFolders = vbChecked


ScanPath.TxtPath = "F:\##ACER-D2"

Call ScanPath.cmdScan_Click


destipath1 = "\\55-88-happy\MY BOOK (M)\##ASUS-D"
destipath2 = "\\55-88-happy\D"


'On Error Resume Next
ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL FILES"

For WE = 1 To ScanPath.ListView1.ListItems.Count
    
    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " Process Files"
    
    DoEvents
    
    'Command1.Caption = Trim(Str(ScanPath.ListView1.ListItems.Count))
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    
    D1 = destipath1 + Mid(A1, Len(ScanPath.TxtPath) + 1)
    D2 = destipath2 + Mid(A1, Len(ScanPath.TxtPath) + 1)
    D3 = "D:" + Mid(A1, Len(ScanPath.TxtPath) + 1)
    
    
    
    If FSO.FileExists(A1 + B1) = True And FSO.FileExists(D1 + B1) = True Then
        
        If m_CRC.CalculateFile(A1 + B1) = m_CRC.CalculateFile(D1 + B1) Then
        
        Err.Clear
        'COPY TO OUR  -D-
        'Debug.Print A1
        'Debug.Print D3
        
        If FSO.FolderExists(D3) = False Then
            'MkDir Mid(D3, 1, InStrRev(D3, "\", Len(D3) - 1))
            MkDir D3
            Stop
            Err.Clear
        End If
        
        FSO.CopyFile A1 + B1, D3 + B1, True
        
        'Err.Clear
        'DEL FROM OUR SWAP DRIVE
        If Err.Number = 0 Then FSO.DeleteFile A1 + B1, True
        
        
        On Error Resume Next
        'COPY TO DESTINATION'S -D-
        If Err.Number = 0 Then FSO.CopyFile A1 + B1, D2 + B1
        
        'DEL FROM DESTINATION SWAP DRIVE
        If Err.Number = 0 Then FSO.DeleteFile D1 + B1, True
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

Set F = FSO.GetDrive("M")
If F.IsReady = True Then
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

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL FILES"
    
'B1 = ScanPath.ListView1.ListItems.Item(1)
'RXRIME = Format(DateValue(Mid(B1, 8, 10)), "YYYY-MM-MMM")
'FN = ScanPath.txtPath.Text + "HTTP's From Email's -- " + RXRIME + ".txt"

'If Dir(FN) <> "" Then Kill FN
'UU = FreeFile
'Open FN For Append As #UU

For WE = 1 To ScanPath.ListView1.ListItems.Count
    
'    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " DEALT WITH FILES"
    
    DoEvents
    
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    TT = TT + "-*/" + B1 + vbCrLf
    
    Set F = FSO.GetFile(A1 + B1)
    ADATE1 = F.DateLastModified
    
    
    
    II = InStr(OO, Chr(1) + Chr(0) + Chr(0) + Chr(0) + Chr(32))
    II = InStr(OO, Trim(Str(Year(Now))))
    If II = 0 Then
    II1 = InStr(OO, Trim(Str(Year(Now) - 1)))
    II3 = InStr(OO, Trim(Str(Year(Now) + 1)))
    If II1 > 0 Then II = II1
    If II2 > 0 Then II = II2
    If II3 > 0 Then II = II2
    End If
    
    IIRIME = ""
    If II > 0 Then
        For r = 1 To 19
            IIRIME = IIRIME + Chr(Asc(Mid(OO, II + r - 1, 1)))
            '&HFA
        Next
    End If
    
    IIRIME2 = ""
    If IIRIME <> "" Then
        Mid(IIRIME, 5, 1) = "/"
        Mid(IIRIME, 8, 1) = "/"
        IIRIME2 = DateValue(IIRIME) + TimeValue(IIRIME)
    End If
    
    If IIRIME2 <> "" Then
        If IIRIME2 > DateSerial(2005, 1, 1) Then
            ADATE1 = IIRIME2
        Else
            Stop
        End If
    End If
    
    YY = Format(ADATE1, "YYYY") + " 0" + Format(ADATE1, "MM") + "(" + Format(ADATE1, "MMM") + ") " + Format(ADATE1, "DD DDD")
    
'    If InStr(B1, "DSC") > 0 Then
'        yy2 = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD") + " " + Format(adate1, "hh-Mm-ss") + " - " + cam + " - " + Mid(B1, InStr(B1, "DSC"))
'    Else
'        yy2 = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD") + " " + Format(adate1, "hh-Mm-ss") + " - " + cam + " - " + B1
'    End If
    
    'yy2 = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD") + " " + Format(adate1, "hh-Mm-ss") + " - " + B1
    'yy2 = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD") + " - " + B1
    
    yy2 = Format(ADATE1, "YYYY") + " 0" + Format(ADATE1, "MM") + "(" + Format(ADATE1, "MMM") + ") " + Format(ADATE1, "DD DDD") + " " + Format(ADATE1, "hh-Mm-ss") + " - " + Mid(B1, InStrRev(B1, "MOV"))
    
    If InStrRev(B1, "MOV") = 0 Then
        Me.WindowState = vbNormal
        DoEvents
        MsgBox "WE WANTED O KNOW IF ANY FILE NOT AND --.MOV -- FILE", vbMsgBoxSetForeground
        Stop
    End If
    
    'If InStr(B1, "JFIF -") > 0 Then
    '    yy2 = Format(adate1, "YYYY") + " 0" + Format(adate1, "MM") + "(" + Format(adate1, "MMM") + ") " + Format(adate1, "DD DDD") + " " + Format(adate1, "hh-Mm-ss") + " - " + Mid(B1, InStrRev(B1, "JFIF -") + 7)
    'End If
    
    'MOVE INTO FOLDER
    
    
        
    
    If FSO.FolderExists(DXDIR + YY) = False Then MkDir DXDIR + YY
    If FSO.FileExists(DXDIR + YY + "\" + yy2) = False Then
        'Call ShellFileMove(A1 + B1, DXDIR + yy + "\" + yy2)
        'Name A1 + B1 As DXDIR + yy + "\" + yy2
    Else
        'Debug.Print "File Exists Already SOURCE ="
        'Debug.Print A1 + B1
    End If
    
    'MOVE WHERE IS
    If FSO.FileExists(A1 + yy2) = False Then
        'F.Move A1 + YY2
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

ScanPath.lblCount1 = Trim(Str(ScanPath.ListView1.ListItems.Count)) + " TOTAL FILES"
    
'B1 = ScanPath.ListView1.ListItems.Item(1)
'RXRIME = Format(DateValue(Mid(B1, 8, 10)), "YYYY-MM-MMM")
'FN = ScanPath.txtPath.Text + "HTTP's From Email's -- " + RXRIME + ".txt"

'If Dir(FN) <> "" Then Kill FN
'UU = FreeFile
'Open FN For Append As #UU

For WE = 1 To ScanPath.ListView1.ListItems.Count
    
'    If ScanPath.ListView1.ListItems.Count = 0 Then Exit For
    
    TTD = TTD + 1
    ScanPath.lblCount2 = Trim(Str(TTD)) + " DEALT WITH FILES"
    
    DoEvents
    
    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ScanPath.ListView1.ListItems.Item(WE)
    G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
    
    TT = TT + "-*/" + B1 + vbCrLf
    
    Set F = FSO.GetFile(A1 + B1)
    ADATE1 = F.DateLastModified
    
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
    
    II = InStr(OO, Chr(1) + Chr(0) + Chr(0) + Chr(0) + Chr(32))
    II = InStr(OO, Trim(Str(Year(Now))))
    If II = 0 Then
    II1 = InStr(OO, Trim(Str(Year(Now) - 1)))
    II3 = InStr(OO, Trim(Str(Year(Now) + 1)))
    If II1 > 0 Then II = II1
    If II2 > 0 Then II = II2
    If II3 > 0 Then II = II2
    End If
    
    IIRIME = ""
    If II > 0 Then
        For r = 1 To 19
            IIRIME = IIRIME + Chr(Asc(Mid(OO, II + r - 1, 1)))
            '&HFA
        Next
    End If
    
    IIRIME2 = ""
    If IIRIME <> "" Then
        Mid(IIRIME, 5, 1) = "/"
        Mid(IIRIME, 8, 1) = "/"
        IIRIME2 = DateValue(IIRIME) + TimeValue(IIRIME)
    End If
    
    If IIRIME2 <> "" Then
        If IIRIME2 > DateSerial(2005, 1, 1) Then
            ADATE1 = IIRIME2
        Else
            Stop
        End If
    End If
    YY = Format(ADATE1, "YYYY") + " 0" + Format(ADATE1, "MM") + "(" + Format(ADATE1, "MMM") + ") " + Format(ADATE1, "DD DDD") + " - " + cam
    
    If InStr(B1, "DSC") > 0 Then
        yy2 = Format(ADATE1, "YYYY") + " 0" + Format(ADATE1, "MM") + "(" + Format(ADATE1, "MMM") + ") " + Format(ADATE1, "DD DDD") + " " + Format(ADATE1, "hh-Mm-ss") + " - " + cam + " - " + Mid(B1, InStr(B1, "DSC"))
    Else
        yy2 = Format(ADATE1, "YYYY") + " 0" + Format(ADATE1, "MM") + "(" + Format(ADATE1, "MMM") + ") " + Format(ADATE1, "DD DDD") + " " + Format(ADATE1, "hh-Mm-ss") + " - " + cam + " - " + B1
    End If
    
    If InStr(B1, "JFIF -") > 0 Then
        yy2 = Format(ADATE1, "YYYY") + " 0" + Format(ADATE1, "MM") + "(" + Format(ADATE1, "MMM") + ") " + Format(ADATE1, "DD DDD") + " " + Format(ADATE1, "hh-Mm-ss") + " - " + Mid(B1, InStrRev(B1, "JFIF -"))
    End If
    
    
    'MOVE INTO FOLDER
    
    If DXDIR = "" Then
        If FSO.FolderExists(ScanPath.TxtPath + YY) = False Then MkDir ScanPath.TxtPath + YY
        If FSO.FileExists(ScanPath.TxtPath + YY + "\" + yy2) = False Then
            'F.Move ScanPath.txtPath + YY + "\" + YY2
            'Call ShellFileMove(A1 + B1, ScanPath.TxtPath + yy + "\" + yy2)
        End If
    Else
        If FSO.FolderExists(DXDIR + YY) = False Then MkDir DXDIR + YY
        If FSO.FileExists(DXDIR + YY + "\" + yy2) = False Then
            'Call ShellFileMove(A1 + B1, DXDIR + yy + "\" + yy2)
        End If
    End If
        'If DXDIR <> "" Then F.Move DXDIR + yy + "\" + YY2
    
    
    'MOVE WHERE IS
    If FSO.FileExists(A1 + yy2) = False Then
        'F.Move A1 + YY2
    End If
    
    
    
    'F.Move A1 + YY + "\" + YY + " - " + B1
    
    
    
    
    
    
Next


End Sub



Private Sub Command9_Click()
    'Call MOVE_ANY_WHERE
End Sub

Public Sub SET_VAR_FS()

Set FS = CreateObject("Scripting.FileSystemObject")

End Sub

Public Sub Form_Load()
    
'Shell "C:\Program Files\WIDCOMM\Bluetooth Software\btsendto_explorer.exe -BD_ADDR=00:1d:28:2b:90:9c -BD_NAME=K80x2D07919180203 -DEV_CLASS=00262746" + " """ + Source1 + """"
    
'Call ShellFileDelete("C:\C", True)
    
Set FS = CreateObject("Scripting.FileSystemObject")

Me.BackColor = &HC0C0C0

'ScanPath.Show
ScanPath.Hide

'--------
'If you use ChkMem Size Files then Know G1$ Alternate Data uses same COLUMN(2) for its data
'-
'Seem to have a problem with some program an not other when use chkmem file size and date
'Work Around done that
'-
'The Date checker box does scan for time if the var contains time
'---------
    
ScanPath.Caption = ScanPath.Caption + App.EXEName
'scanpath.Caption = frmMain.Caption + App.EXEName
  
x = 1
y = 1
On Error Resume Next
For Each Control In Controls
        'CONTROL.NAME
    If Mid(LCase(Control.Name), 1, 3) <> "mnu" Then
        If Control.Enabled = True Then
            If Control.Width + Control.Left > x Then x = Control.Width + Control.Left
            If Control.Height + Control.Top > y Then y = Control.Height + Control.Top
            If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
        End If
    End If
Next
On Error GoTo 0


ScanPath.Width = x + 110
If mnu = 1 Then
    pluso = 720
Else
    pluso = 450
End If
ScanPath.Height = y + pluso
    
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
    .ColumnHeaders.Add , "FILE", "File", 4200, lvwColumnLeft
    .ColumnHeaders.Add , "PATH", "Path", 8100, lvwColumnLeft
    .ColumnHeaders.Add , "DOS-8.3", "Dos", 1400, lvwColumnLeft
    .ColumnHeaders.Add , "SIZE", "Size", 1100, lvwColumnRight
    .ColumnHeaders.Add , "DATE", "Date", 1700, lvwColumnLeft
    .ColumnHeaders.Add , "CRC", "CRC", 1700, lvwColumnLeft
    
    .View = lvwReport
End With

'Call chkCopyMemory_Click
chkCopyMemory.Value = False

ListView1.Width = Me.Width - 50


'Me.Show

'TxtPath = "D:\VB6\VB-NT\"
'cboMask = "*.EXE"
'ScanPath.chkSubFolders = vbChecked

'Call cmdScan_Click
'
'
'Open App.Path + "\File of EXE Programs in VB.txt" For Output As #1
'Print #1, Str(ScanPath.ListView1.ListItems.Count) + " ITEMS"
'For WE = 1 To ScanPath.ListView1.ListItems.Count
'    A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
'    B1 = ScanPath.ListView1.ListItems.Item(WE)
'    YY$ = A1 + B1
'    'Debug.Print yy
'    If YY <> "" Then
'        Print #1, A1 + B1
'    End If
'Next
'Close #1
    
    'Debug.Print Str(ScanPath.ListView1.ListItems.Count) + " ITEMS"
'Shell "notepad.exe """ + App.Path + "\File of EXE Programs in VB.txt"""

'Exit Sub


'Shell "E:\01 VB Shell Folders\00 Shell VBasic 6 Loader\Sort_AnyThing - MULTI USE MENU.exe.lnk MOVE_IN_DATE_FOLDERS_JPGS", vbMinimizedNoFocus
'If Command$ = "MOVE_IN_DATE_FOLDERS_M2CARD" Then 'Or IsIDE = True Then
'    'MsgBox "HELLO"
'    ScanPath.Visible = False
'    Call MOVE_IN_DATE_FOLDERS_M2CARD_Click
'    End
'End If


'Call GROUP_IN_500_JPG_MEDIA_PLAYER


End Sub


Private Sub Form_Unload(Cancel As Integer)
'End
End Sub

Private Sub List1_Click()

tpath3$ = "M:\04 Music ---\Del\"

B1 = Mid$(List1.List(List1.ListIndex), 14)
c1$ = B1 + "--Todel"

'tpath1$ = "M:\04 Music ---\"

c1$ = Mid$(B1, InStrRev(B1, "\") + 1)
List1.RemoveItem (List1.ListIndex)
'Name B1 As c1$

'fsO.copyFile B1, tpath3$ + c1$
If Dir$(tpath3$ + c1$) <> "" Then Kill tpath3$ + c1$
FSO.MoveFile B1, tpath3$ + c1$



End Sub

Private Sub Label18_Click()
Dog = True

If cmdScan.Caption = "Stop" Then cmdScan_Click
Label17 = "Stopped before Complete"
End Sub


Private Sub Label5_Click()

    Label5.BackColor = Label6.BackColor


End Sub





Sub Jeepers()


On Local Error GoTo jeep
Drived2$ = Mid$(TxtPath.Text, 1, 3)
MkDir Drived2$ + "Temp\Anything"
V1$ = Mid$(TxtPath.Text, InStrRev(TxtPath.Text, "\"))

MkDir Drived2$ + "Temp\Anything" + V1$
On Local Error GoTo 0

List1.AddItem "Stage 1 of 2 : Make Dir's And Move Files to Temp"
List1.AddItem "------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For WE = 1 To ListView1.ListItems.Count
    A1 = ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ListView1.ListItems.Item(WE)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(TxtPath.Text)
    c1$ = Drived2$ + "Temp\Anything" + V1$ + "\" + Mid$(A1, ets1 + 2)

    etd1 = 13
    ets2 = 1
    Do
        'jc1$ = Mid$(c1$, etd1 + ets2-1)
        ets2 = InStr(etd1 + ets2, c1$, "\")
        etd1 = 1
        If ets2 = 0 Then Exit Do

        D2$ = Mid$(c1$, 1, ets2)

        If InStr(F1$, D2$) = 0 Then
            On Local Error GoTo jeep
            MkDir D2$
            If errs2 <> 75 And errs2 > 0 Then
                MsgBox "Error Making Temp Directory", vbMsgBoxSetForeground
                Stop
            End If
            On Local Error GoTo 0
        End If
    Loop Until ets2 = 0

    F1$ = D2$

    errs2 = 0
    On Local Error GoTo jeep
    FSO.MoveFile A1 + B1, c1$ + B1
    On Local Error GoTo 0

    If errs2 <> 0 Then
        MsgBox "Error Moving File to Temp", vbMsgBoxSetForeground
        End
    End If

    List1.AddItem Format$(WE, "000 ") + "Move > " + A1 + B1
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'FsO.DeleteFolder comrad$ + "\" + WFD$

Next

Err.Clear
On Local Error Resume Next
FSO.DeleteFolder TxtPath.Text, True
If Err.Number > 0 Then Call HardDel
On Local Error GoTo 0


List1.AddItem "----------------------------------------"
List1.AddItem "Stage 2 of 2 : Move Files Back From Temp"
List1.AddItem "----------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

V2$ = Mid$(TxtPath.Text, 1, InStrRev(TxtPath.Text, "\"))

Err.Clear
On Local Error Resume Next
FSO.MoveFolder Drived2$ + "Temp\Anything\*", V2$
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
V1$ = Mid$(TxtPath.Text, InStrRev(TxtPath.Text, "\"))


List1.AddItem "Stage Opp : Del Original Dir"
List1.AddItem "------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For WE = 1 To ListView1.ListItems.Count
    A1 = ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ListView1.ListItems.Item(WE)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(TxtPath.Text)
    c1$ = Drived2$ + "My Music" + V1$ + "\" + Mid$(A1, ets1 + 2)

    etd1 = 13
    ets2 = 1
    Do
        'jc1$ = Mid$(c1$, etd1 + ets2-1)
        ets2 = InStr(etd1 + ets2, c1$, "\")
        etd1 = 1
        If ets2 = 0 Then Exit Do

        D2$ = Mid$(c1$, 1, ets2 - 1)

        If InStr(F1$, D2$) = 0 Then
            Err.Clear
            On Local Error Resume Next
            'MkDir d2$
            FSO.DeleteFolder D2$, True
            If Err.Number <> 75 And Err.Number > 0 Then
                'MsgBox ("Error Making Temp Directory",vbMsgBoxSetForeground)
                '
            End If
            On Local Error GoTo 0
        End If
    Loop Until ets2 = 0

    F1$ = D2$

    errs2 = 0
    On Local Error Resume Next
'    fsO.DeleteFile A1 + B1, True
    On Local Error GoTo 0

    'If errs2 <> 0 Then
    '    MsgBox ("Error Moving File to Temp",vbMsgBoxSetForeground)
    '    End
    'End If

    List1.AddItem Format$(WE, "000 ") + "Move > " + A1 + B1
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'fsO.DeleteFolder comrad$ + "\" + WFD$

Next

A = A
jeep:
End Sub


Sub HardMove()

Drived2$ = Mid$(TxtPath.Text, 1, 3)
V1$ = Mid$(TxtPath.Text, InStrRev(TxtPath.Text, "\"))


List1.AddItem "Stage Opp : Del Original Dir"
List1.AddItem "------------------------------------------------"
List1.ListIndex = List1.ListCount - 1
List1.Refresh

For WE = 1 To ListView1.ListItems.Count
    A1 = ListView1.ListItems.Item(WE).SubItems(1)
    B1 = ListView1.ListItems.Item(WE)

    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
    ets1 = Len(TxtPath.Text)
    c1$ = Drived2$ + "My Music" + V1$ + "\" + Mid$(A1, ets1 + 2)
    c2$ = Drived2$ + "Temp\Anything" + V1$ + "\" + Mid$(A1, ets1 + 2)

    etd1 = 13
    ets2 = 1
    Do
        'jc1$ = Mid$(c1$, etd1 + ets2-1)
        ets2 = InStr(etd1 + ets2, c1$, "\")
        etd1 = 1
        If ets2 = 0 Then Exit Do

        D2$ = Mid$(c1$, 1, ets2 - 1)

        If InStr(F1$, D2$) = 0 Then
            Err.Clear
            On Local Error Resume Next
            MkDir D2$
   '         fsO.deletefolder d2$, True
            If Err.Number <> 75 And Err.Number > 0 Then
                'MsgBox ("Error Making Temp Directory",vbMsgBoxSetForeground)
                'T
            End If
            On Local Error GoTo 0
        End If
    Loop Until ets2 = 0

    F1$ = D2$
    Err.Clear
    errs2 = 0
    On Local Error Resume Next
    
    FSO.MoveFile c2$ + B1, A1 + B1
    'err.number
    'err.description
    On Local Error GoTo 0

    'If errs2 <> 0 Then
    '    MsgBox ("Error Moving File to Temp",vbMsgBoxSetForeground)
    '    End
    'End If

    List1.AddItem Format$(WE, "000 ") + "Move > " + A1 + B1
    List1.ListIndex = List1.ListCount - 1
    List1.Refresh

    'fsO.DeleteFolder comrad$ + "\" + WFD$

Next
    
    V1$ = Mid$(TxtPath.Text, 1, InStrRev(TxtPath.Text, "\") - 1)
    Err.Clear
    On Local Error Resume Next
    
    FSO.CopyFolder Drived2$ + "Temp\Anything", V1$, True
    
    FSO.DeleteFolder Drived2$ + "Temp\Anything", True


jeep:
End Sub

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
  
  'TESTING
'  IsIDE = False
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************


Private Sub txtPath_Click()
If TxtPath <> Clipboard.GetText Then TxtPath = Clipboard.GetText

If Mid(ScanPath.TxtPath, Len(ScanPath.TxtPath) - 1, 1) <> "\" Then ScanPath.TxtPath = ScanPath.TxtPath + "\"

TxtPath.SelStart = 0
TxtPath.SelLength = Len(TxtPath)



End Sub

