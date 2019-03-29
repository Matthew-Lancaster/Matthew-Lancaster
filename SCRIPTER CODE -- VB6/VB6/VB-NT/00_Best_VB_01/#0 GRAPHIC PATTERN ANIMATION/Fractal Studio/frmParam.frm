VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmParam 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Parameters & Options"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   367
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   448
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CmD 
      Left            =   5640
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picPrg 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   1680
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   75
      Top             =   4920
      Width           =   4455
   End
   Begin VB.CommandButton btnRefreshParams 
      Height          =   495
      Left            =   6120
      Picture         =   "frmParam.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton btnCalc 
      Caption         =   "Calculate"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   1455
   End
   Begin TabDlg.SSTab TabOptions 
      Height          =   4815
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   529
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Calculation"
      TabPicture(0)   =   "frmParam.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgIcon_Calculation"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCalc"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblFractType"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraCalcParams"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraCalcOptions"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmbFractType"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Filters"
      TabPicture(1)   =   "frmParam.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "imgIcon_Filters"
      Tab(1).Control(1)=   "lblFilters"
      Tab(1).Control(2)=   "lblFiltersOut"
      Tab(1).Control(3)=   "lblFiltersIn"
      Tab(1).Control(4)=   "lblMapOut"
      Tab(1).Control(5)=   "lblMapIn"
      Tab(1).Control(6)=   "Line5"
      Tab(1).Control(7)=   "Line6"
      Tab(1).Control(8)=   "Line7"
      Tab(1).Control(9)=   "Line8"
      Tab(1).Control(10)=   "btnUpOut"
      Tab(1).Control(11)=   "btnDownOut"
      Tab(1).Control(12)=   "btnAddOut"
      Tab(1).Control(13)=   "cmbMapOut"
      Tab(1).Control(14)=   "cmbMapIn"
      Tab(1).Control(15)=   "btnRemOut"
      Tab(1).Control(16)=   "btnAddIn"
      Tab(1).Control(17)=   "btnUpIn"
      Tab(1).Control(18)=   "btnRemIn"
      Tab(1).Control(19)=   "btnDownIn"
      Tab(1).Control(20)=   "lstFiltersOut"
      Tab(1).Control(21)=   "lstFiltersIn"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "Zoom"
      TabPicture(2)   =   "frmParam.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstZoom"
      Tab(2).Control(1)=   "btnZoomClear"
      Tab(2).Control(2)=   "lblTrashVPHistory"
      Tab(2).Control(3)=   "imgIcon_Zoom"
      Tab(2).Control(4)=   "lblZoom"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Position"
      TabPicture(3)   =   "frmParam.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "imgIcon_Pos"
      Tab(3).Control(1)=   "lblPosition"
      Tab(3).Control(2)=   "lblPosSX"
      Tab(3).Control(3)=   "lblPosSY"
      Tab(3).Control(4)=   "lblPosEX"
      Tab(3).Control(5)=   "lblPosEY"
      Tab(3).Control(6)=   "fraPos"
      Tab(3).Control(7)=   "txtPosSX"
      Tab(3).Control(8)=   "txtPosSY"
      Tab(3).Control(9)=   "txtPosEX"
      Tab(3).Control(10)=   "txtPosEY"
      Tab(3).ControlCount=   11
      TabCaption(4)   =   "Palette"
      TabPicture(4)   =   "frmParam.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Line3"
      Tab(4).Control(1)=   "Line4"
      Tab(4).Control(2)=   "imgIcon_Palette"
      Tab(4).Control(3)=   "lblPalette"
      Tab(4).Control(4)=   "imgPal"
      Tab(4).Control(5)=   "scrPalImg"
      Tab(4).Control(6)=   "btnPalZoomIn"
      Tab(4).Control(7)=   "btnPalZoomOut"
      Tab(4).Control(8)=   "fraPalOptions"
      Tab(4).Control(9)=   "fraPalCtrl"
      Tab(4).ControlCount=   10
      TabCaption(5)   =   "Color Cycling"
      TabPicture(5)   =   "frmParam.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "imgIcon_Animation"
      Tab(5).Control(1)=   "lblPalanimSpeed"
      Tab(5).Control(2)=   "lblAnimation"
      Tab(5).Control(3)=   "lblNorm"
      Tab(5).Control(4)=   "btnPalAnim"
      Tab(5).Control(5)=   "txtPalSpeed"
      Tab(5).Control(6)=   "fraNormPAnim"
      Tab(5).Control(7)=   "fraInterpolatePAnim"
      Tab(5).Control(8)=   "UDPASpeed"
      Tab(5).ControlCount=   9
      TabCaption(6)   =   "States"
      TabPicture(6)   =   "frmParam.frx":04EA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "imgIcon_States"
      Tab(6).Control(1)=   "lblStates"
      Tab(6).Control(2)=   "Label1"
      Tab(6).Control(3)=   "lstStates"
      Tab(6).Control(4)=   "btnSaveState"
      Tab(6).Control(5)=   "btnRemState"
      Tab(6).Control(6)=   "btnSaveStateFile"
      Tab(6).ControlCount=   7
      Begin MSComCtl2.UpDown UDPASpeed 
         Height          =   285
         Left            =   -72960
         TabIndex        =   112
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPalSpeed"
         BuddyDispid     =   196625
         OrigLeft        =   2280
         OrigTop         =   1080
         OrigRight       =   2535
         OrigBottom      =   1335
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Frame fraInterpolatePAnim 
         Caption         =   "Interpolation"
         ForeColor       =   &H00000080&
         Height          =   1455
         Left            =   -71160
         TabIndex        =   108
         Top             =   960
         Width           =   2415
         Begin VB.ComboBox cmbPalAnimIPScale 
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "frmParam.frx":0506
            Left            =   240
            List            =   "frmParam.frx":0525
            TabIndex        =   111
            Text            =   "1 (None)"
            Top             =   960
            Width           =   1215
         End
         Begin VB.CheckBox chkPalAnimInterpolate 
            Caption         =   "Interpolate palette"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   255
            Left            =   240
            TabIndex        =   109
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblPalAnimIP_Scale 
            Caption         =   "Scale"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   255
            Left            =   240
            TabIndex        =   110
            Top             =   720
            Width           =   495
         End
      End
      Begin VB.Frame fraNormPAnim 
         Caption         =   "Normalization"
         ForeColor       =   &H00000080&
         Height          =   2055
         Left            =   -71160
         TabIndex        =   102
         Top             =   2520
         Width           =   2415
         Begin VB.TextBox txtNormDiff 
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1200
            TabIndex        =   107
            Text            =   "350"
            Top             =   1440
            Width           =   975
         End
         Begin VB.OptionButton optNormIntSpec 
            Caption         =   "Specify:"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   255
            Left            =   240
            TabIndex        =   106
            Top             =   1440
            Width           =   1095
         End
         Begin VB.OptionButton optNormIntDiff 
            Caption         =   "Set to max difference"
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   255
            Left            =   240
            TabIndex        =   105
            Top             =   1080
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.CheckBox chkNormalizeFrame 
            Caption         =   "Normalize Iterations"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   255
            Left            =   240
            TabIndex        =   103
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblNormDiff 
            Caption         =   "Interval Size"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   104
            Top             =   840
            Width           =   1215
         End
      End
      Begin MSComctlLib.ListView lstZoom 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   44
         Top             =   960
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   128
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Coordinates"
            Object.Width           =   8255
         EndProperty
      End
      Begin VB.CommandButton btnSaveStateFile 
         Caption         =   "Save to File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69840
         TabIndex        =   94
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton btnRemState 
         Caption         =   "Remove State"
         Height          =   375
         Left            =   -69840
         TabIndex        =   93
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton btnSaveState 
         Caption         =   "Save State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69840
         TabIndex        =   92
         Top             =   3240
         Width           =   1335
      End
      Begin MSComctlLib.ListView lstStates 
         Height          =   3255
         Left            =   -74805
         TabIndex        =   95
         Top             =   1320
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   5741
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtPalSpeed 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -73800
         TabIndex        =   86
         Text            =   "1"
         Top             =   1200
         Width           =   840
      End
      Begin VB.CommandButton btnPalAnim 
         Caption         =   "Start Cycling"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   87
         Top             =   4200
         Width           =   3015
      End
      Begin MSComctlLib.ListView lstFiltersIn 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   35
         Top             =   3260
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "On"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "P1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "P2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "P3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "P4"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "P5"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lstFiltersOut 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   28
         Top             =   1200
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   353
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "On"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "P1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "P2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "P3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "P4"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "P5"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton btnDownIn 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -69120
         Picture         =   "frmParam.frx":054F
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   3590
         Width           =   615
      End
      Begin VB.CommandButton btnRemIn 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -69120
         Picture         =   "frmParam.frx":0991
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4310
         Width           =   615
      End
      Begin VB.CommandButton btnUpIn 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -69120
         Picture         =   "frmParam.frx":0DD3
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton btnAddIn 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -69120
         Picture         =   "frmParam.frx":1215
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton btnRemOut 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -69120
         Picture         =   "frmParam.frx":1657
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2390
         Width           =   615
      End
      Begin VB.ComboBox cmbMapIn 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -73320
         TabIndex        =   34
         Text            =   "Inside Mapping Filter"
         Top             =   2900
         Width           =   3015
      End
      Begin VB.ComboBox cmbMapOut 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "frmParam.frx":1A99
         Left            =   -73320
         List            =   "frmParam.frx":1A9B
         TabIndex        =   27
         Text            =   "Outside Mapping Filter"
         Top             =   860
         Width           =   3015
      End
      Begin VB.ComboBox cmbFractType 
         ForeColor       =   &H00004000&
         Height          =   315
         ItemData        =   "frmParam.frx":1A9D
         Left            =   1440
         List            =   "frmParam.frx":1AD7
         TabIndex        =   2
         Text            =   "Mandelbrot"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtPosEY 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   -74160
         TabIndex        =   52
         Text            =   "2"
         Top             =   2180
         Width           =   1335
      End
      Begin VB.TextBox txtPosEX 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   -74160
         TabIndex        =   51
         Text            =   "2"
         Top             =   1820
         Width           =   1335
      End
      Begin VB.TextBox txtPosSY 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   -74160
         TabIndex        =   50
         Text            =   "-2"
         Top             =   1460
         Width           =   1335
      End
      Begin VB.TextBox txtPosSX 
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   -74160
         TabIndex        =   49
         Text            =   "-2"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Frame fraPos 
         Caption         =   "Position Controls"
         ForeColor       =   &H00000080&
         Height          =   1095
         Left            =   -74880
         TabIndex        =   56
         Top             =   2640
         Width           =   2055
         Begin VB.CommandButton btnPosDefault 
            Caption         =   "Default"
            Height          =   420
            Left            =   360
            TabIndex        =   53
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton btnZoomClear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -69000
         Picture         =   "frmParam.frx":1BD1
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Clear History"
         Top             =   480
         Width           =   495
      End
      Begin VB.Frame fraPalCtrl 
         Caption         =   "Palette Controls"
         ForeColor       =   &H00000080&
         Height          =   2775
         Left            =   -74880
         TabIndex        =   47
         Top             =   1800
         Width           =   3135
         Begin VB.CommandButton btnFirstBlack 
            Caption         =   "Set first color to black"
            Height          =   495
            Left            =   1560
            TabIndex        =   100
            Top             =   2160
            Width           =   1335
         End
         Begin VB.CommandButton btnPalInvert 
            Caption         =   "Invert Colors"
            Height          =   495
            Left            =   120
            TabIndex        =   62
            Top             =   2160
            Width           =   1335
         End
         Begin VB.CommandButton btnPalSave 
            Caption         =   "Save Palette"
            Height          =   495
            Left            =   1560
            TabIndex        =   66
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton btnPalOpen 
            Caption         =   "Open Palette"
            Height          =   495
            Left            =   1560
            TabIndex        =   64
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton btnPalRnd 
            Caption         =   "Create Random"
            Height          =   495
            Left            =   120
            TabIndex        =   61
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CommandButton btnPalMakeRnd 
            Caption         =   "Create New"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   60
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.Frame fraPalOptions 
         Caption         =   "New Palette Parameters"
         ForeColor       =   &H00000080&
         Height          =   2775
         Left            =   -71640
         TabIndex        =   41
         Top             =   1800
         Width           =   2895
         Begin VB.TextBox txtPalSize 
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   1920
            TabIndex        =   67
            Text            =   "2000"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtPalPeriod 
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   1920
            TabIndex        =   72
            Text            =   "20"
            Top             =   1800
            Width           =   735
         End
         Begin VB.VScrollBar scrPalDisp 
            Enabled         =   0   'False
            Height          =   255
            LargeChange     =   10
            Left            =   2400
            Max             =   100
            TabIndex        =   43
            Top             =   1320
            Width           =   255
         End
         Begin VB.TextBox txtPalDisp 
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   1920
            TabIndex        =   71
            Text            =   "100"
            Top             =   1320
            Width           =   495
         End
         Begin VB.CheckBox chkPalUseDisp 
            Caption         =   "Use Dispersion Control"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   960
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.Label lblPalSize 
            Caption         =   "Palette Size"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblPalPeriod 
            Caption         =   "Gradient Period"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   1835
            Width           =   1455
         End
         Begin VB.Label lblPalDispRate 
            Caption         =   "Dispersion Rate (%)"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.CommandButton btnPalZoomOut 
         Caption         =   "Zoom Out"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69960
         TabIndex        =   76
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton btnPalZoomIn 
         Caption         =   "Zoom In"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71160
         TabIndex        =   74
         Top             =   500
         Width           =   1095
      End
      Begin VB.HScrollBar scrPalImg 
         Height          =   255
         LargeChange     =   50
         Left            =   -74880
         SmallChange     =   5
         TabIndex        =   40
         Top             =   1320
         Width           =   6135
      End
      Begin VB.PictureBox imgPal 
         BackColor       =   &H00000000&
         Height          =   360
         Left            =   -74880
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   405
         TabIndex        =   32
         Top             =   980
         Width           =   6135
      End
      Begin VB.CommandButton btnAddOut 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -69120
         Picture         =   "frmParam.frx":1D5B
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2060
         Width           =   615
      End
      Begin VB.CommandButton btnDownOut 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -69120
         Picture         =   "frmParam.frx":219D
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1550
         Width           =   615
      End
      Begin VB.CommandButton btnUpOut 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -69120
         Picture         =   "frmParam.frx":25DF
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1200
         Width           =   615
      End
      Begin VB.Frame fraCalcOptions 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   3855
         Left            =   3360
         TabIndex        =   21
         Top             =   840
         Width           =   3255
         Begin VB.Frame fraOutputSize 
            Caption         =   " Output Size "
            ForeColor       =   &H00000080&
            Height          =   1335
            Left            =   120
            TabIndex        =   68
            Top             =   1320
            Width           =   2895
            Begin VB.CheckBox chkCustomOutputSize 
               Caption         =   "Check1"
               Height          =   255
               Left            =   1680
               TabIndex        =   114
               Top             =   360
               Value           =   1  'Checked
               Width           =   255
            End
            Begin VB.ComboBox cmbSize 
               ForeColor       =   &H00004000&
               Height          =   315
               ItemData        =   "frmParam.frx":2A21
               Left            =   240
               List            =   "frmParam.frx":2A4C
               TabIndex        =   12
               Text            =   "Custom"
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox txtH 
               ForeColor       =   &H00004000&
               Height          =   285
               Left            =   1680
               TabIndex        =   14
               Text            =   "540"
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox txtW 
               ForeColor       =   &H00004000&
               Height          =   285
               Left            =   240
               TabIndex        =   13
               Text            =   "800"
               Top             =   840
               Width           =   975
            End
            Begin VB.Label lblCustomSizeX 
               Caption         =   "x"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1395
               TabIndex        =   113
               Top             =   840
               Width           =   255
            End
            Begin VB.Label lblCalc2FileCustomSize 
               Caption         =   "Custom"
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   2040
               TabIndex        =   70
               Top             =   395
               Width           =   615
            End
         End
         Begin VB.CheckBox chkLinearCon 
            Caption         =   "Color Interpolation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.TextBox txtCalcParam3 
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   1800
            TabIndex        =   11
            Text            =   "15"
            Top             =   840
            Width           =   975
         End
         Begin VB.VScrollBar scrCalcParam3 
            Enabled         =   0   'False
            Height          =   255
            Left            =   2760
            TabIndex        =   23
            Top             =   855
            Width           =   255
         End
         Begin VB.Frame fraCalcOptionsFile 
            Caption         =   " Direct File Output "
            ForeColor       =   &H00000080&
            Height          =   975
            Left            =   120
            TabIndex        =   22
            Top             =   2760
            Width           =   2895
            Begin VB.TextBox txtCalc2File_Size 
               ForeColor       =   &H00004000&
               Height          =   285
               Left            =   840
               TabIndex        =   16
               Text            =   "2000"
               Top             =   400
               Width           =   855
            End
            Begin VB.CommandButton btnRenderToFile 
               Caption         =   "Render"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1920
               TabIndex        =   15
               Top             =   360
               Width           =   855
            End
            Begin VB.Label lblCalc2FileSize 
               Caption         =   "Size"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   240
               TabIndex        =   91
               Top             =   435
               Width           =   495
            End
         End
         Begin VB.Label lblCalcParamLines 
            Caption         =   "Lines In Chunk"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   865
            Width           =   1455
         End
      End
      Begin VB.Frame fraCalcParams 
         Caption         =   " Parameters "
         ForeColor       =   &H00000080&
         Height          =   3375
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   3135
         Begin VB.TextBox txtCalcP 
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   285
            Index           =   4
            Left            =   1800
            TabIndex        =   9
            Text            =   "0"
            Top             =   3000
            Width           =   1095
         End
         Begin VB.TextBox txtCalcP 
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   285
            Index           =   3
            Left            =   1800
            TabIndex        =   8
            Text            =   "0"
            Top             =   2640
            Width           =   1095
         End
         Begin VB.TextBox txtCalcP 
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   7
            Text            =   "0"
            Top             =   2280
            Width           =   1095
         End
         Begin VB.TextBox txtCalcP 
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   6
            Text            =   "0"
            Top             =   1920
            Width           =   1095
         End
         Begin MSComCtl2.UpDown UDBailout 
            Height          =   285
            Left            =   2640
            TabIndex        =   82
            Top             =   720
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtBailout"
            BuddyDispid     =   196683
            OrigLeft        =   2640
            OrigTop         =   720
            OrigRight       =   2880
            OrigBottom      =   975
            Max             =   200
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UDMaxI 
            Height          =   285
            Left            =   2640
            TabIndex        =   81
            Top             =   360
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtMaxIterations"
            BuddyDispid     =   196684
            OrigLeft        =   2640
            OrigTop         =   360
            OrigRight       =   2880
            OrigBottom      =   615
            Max             =   30000
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtCalcP 
            Enabled         =   0   'False
            ForeColor       =   &H00004000&
            Height          =   285
            Index           =   0
            Left            =   1800
            TabIndex        =   5
            Text            =   "0"
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox txtBailout 
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   1800
            TabIndex        =   4
            Text            =   "4"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtMaxIterations 
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   1800
            TabIndex        =   3
            Text            =   "180"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblCalcParam 
            Caption         =   "Param 01"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   89
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Label lblCalcParam 
            Caption         =   "Param 01"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   88
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label lblCalcParam 
            Caption         =   "Param 01"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   85
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblCalcParam 
            Caption         =   "Param 01"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   84
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000010&
            X1              =   240
            X2              =   2880
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label lblCalcParam 
            Caption         =   "Param 01"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label lblCalcOtherParams 
            Caption         =   "Fractal Parameters:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   255
            Left            =   720
            TabIndex        =   25
            Top             =   1155
            Width           =   1815
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            X1              =   240
            X2              =   2880
            Y1              =   1455
            Y2              =   1455
         End
         Begin VB.Label lblBailout 
            Caption         =   "Bailout Value"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   765
            Width           =   1335
         End
         Begin VB.Label lblMaxIterations 
            Caption         =   "Max Iterations"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   405
            Width           =   1335
         End
      End
      Begin VB.Label Label1 
         Caption         =   "This section is not yet implemented.."
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -72840
         TabIndex        =   115
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblTrashVPHistory 
         Caption         =   "Empty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   -69720
         TabIndex        =   101
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblNorm 
         BackColor       =   &H80000010&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmParam.frx":2AC4
         Height          =   1575
         Left            =   -74760
         TabIndex        =   99
         Top             =   2520
         Width           =   3015
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Realtime Color Cycling"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -74160
         TabIndex        =   98
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblPalette 
         Caption         =   "Palette Options / Creation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -74160
         TabIndex        =   97
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblStates 
         Caption         =   "States"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -74160
         TabIndex        =   96
         Top             =   480
         Width           =   975
      End
      Begin VB.Image imgIcon_States 
         Height          =   480
         Left            =   -74880
         Picture         =   "frmParam.frx":2BDB
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblPalanimSpeed 
         Caption         =   "Speed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   90
         Top             =   1200
         Width           =   615
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000014&
         X1              =   -69225
         X2              =   -69225
         Y1              =   3260
         Y2              =   4655
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000014&
         X1              =   -69225
         X2              =   -69225
         Y1              =   1220
         Y2              =   2820
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000010&
         X1              =   -69240
         X2              =   -69240
         Y1              =   3260
         Y2              =   4655
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000010&
         X1              =   -69240
         X2              =   -69240
         Y1              =   2820
         Y2              =   1220
      End
      Begin VB.Label lblMapIn 
         Caption         =   "<-- Mapping Filter"
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
         Left            =   -70080
         TabIndex        =   80
         Top             =   2935
         Width           =   1575
      End
      Begin VB.Label lblMapOut 
         Caption         =   "<-- Mapping Filter"
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
         Left            =   -70080
         TabIndex        =   79
         Top             =   920
         Width           =   1575
      End
      Begin VB.Label lblFiltersIn 
         Caption         =   "Inside Filters"
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
         Left            =   -74880
         TabIndex        =   78
         Top             =   2935
         Width           =   2775
      End
      Begin VB.Label lblFiltersOut 
         Caption         =   "Outside Filters"
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
         Left            =   -74880
         TabIndex        =   77
         Top             =   920
         Width           =   2655
      End
      Begin VB.Label lblFractType 
         Caption         =   "Fractal Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   160
         TabIndex        =   73
         Top             =   1000
         Width           =   1215
      End
      Begin VB.Label lblPosEY 
         Caption         =   "End Y"
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   65
         Top             =   2200
         Width           =   495
      End
      Begin VB.Label lblPosEX 
         Caption         =   "End X"
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   63
         Top             =   1815
         Width           =   495
      End
      Begin VB.Label lblPosSY 
         Caption         =   "Start Y"
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   59
         Top             =   1455
         Width           =   495
      End
      Begin VB.Label lblPosSX 
         Caption         =   "Start X"
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   58
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label lblPosition 
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -74160
         TabIndex        =   57
         Top             =   480
         Width           =   855
      End
      Begin VB.Image imgIcon_Animation 
         Height          =   480
         Left            =   -74880
         Picture         =   "frmParam.frx":301D
         Top             =   380
         Width           =   480
      End
      Begin VB.Image imgIcon_Pos 
         Height          =   480
         Left            =   -74880
         Picture         =   "frmParam.frx":3327
         Top             =   380
         Width           =   480
      End
      Begin VB.Image imgIcon_Palette 
         Height          =   480
         Left            =   -74880
         Picture         =   "frmParam.frx":3631
         Top             =   380
         Width           =   480
      End
      Begin VB.Label lblFilters 
         Caption         =   "Image Filters"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -74160
         TabIndex        =   55
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblCalc 
         Caption         =   "Calculation Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   840
         TabIndex        =   54
         Top             =   480
         Width           =   1695
      End
      Begin VB.Image imgIcon_Zoom 
         Height          =   480
         Left            =   -74880
         Picture         =   "frmParam.frx":3A73
         Top             =   380
         Width           =   480
      End
      Begin VB.Label lblZoom 
         Caption         =   "Viewpoint History"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -74160
         TabIndex        =   48
         Top             =   480
         Width           =   1815
      End
      Begin VB.Image imgIcon_Filters 
         Height          =   480
         Left            =   -74880
         Picture         =   "frmParam.frx":3EB5
         Top             =   380
         Width           =   480
      End
      Begin VB.Image imgIcon_Calculation 
         Height          =   480
         Left            =   120
         Picture         =   "frmParam.frx":42F7
         Top             =   380
         Width           =   480
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   -74880
         X2              =   -68760
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   -74880
         X2              =   -68760
         Y1              =   1665
         Y2              =   1665
      End
   End
   Begin VB.Menu MNU_LOGGF 
      Caption         =   "LOGG FOLDER"
   End
End
Attribute VB_Name = "frmParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAddIn_Click()

frmAddFilter.Tag = 1
frmAddFilter.Show vbModal
ShowParams Frac(aData.nActive)

End Sub

Private Sub btnAddOut_Click()

frmAddFilter.Tag = 0
frmAddFilter.Show vbModal
ShowParams Frac(aData.nActive)

End Sub

Private Sub btnCalc_Click()

Calculate aData.nActive, True

End Sub


Private Sub btnDownIn_Click()
Dim nIndex As Long
Dim Tmp As tFilterVars

If lstFiltersIn.ListItems.Count < 1 Then Exit Sub
nIndex = lstFiltersIn.SelectedItem.Index - 1

With Frac(aData.nActive)
    If nIndex < lstFiltersIn.ListItems.Count - 1 And .nFiltersIn > 1 Then
        'Swap this filter with the one below:
        Tmp = .FilterIn(nIndex + 1)
        .FilterIn(nIndex + 1) = .FilterIn(nIndex)
        .FilterIn(nIndex) = Tmp
        ShowParams Frac(aData.nActive)
        lstFiltersIn.ListItems(nIndex + 2).Selected = True
        Set lstFiltersIn.SelectedItem = lstFiltersIn.ListItems(nIndex + 2)
    End If
End With

End Sub

Private Sub btnDownOut_Click()
Dim nIndex As Long
Dim Tmp As tFilterVars

If lstFiltersOut.ListItems.Count < 1 Then Exit Sub
nIndex = lstFiltersOut.SelectedItem.Index - 1

With Frac(aData.nActive)
    If nIndex < lstFiltersOut.ListItems.Count - 1 And .nFiltersOut > 1 Then
        'Swap this filter with the one below:
        Tmp = .FilterOut(nIndex + 1)
        .FilterOut(nIndex + 1) = .FilterOut(nIndex)
        .FilterOut(nIndex) = Tmp
        ShowParams Frac(aData.nActive)
        lstFiltersOut.ListItems(nIndex + 2).Selected = True
        Set lstFiltersOut.SelectedItem = lstFiltersOut.ListItems(nIndex + 2)
    End If
End With

End Sub

Private Sub btnFirstBlack_Click()

Frac(aData.nActive).Pal(0).R = 0
Frac(aData.nActive).Pal(0).G = 0
Frac(aData.nActive).Pal(0).B = 0

End Sub

Private Sub btnPalAnim_Click()

If chkPalAnimInterpolate.Value = 1 Then
    DoPalAnim Frac(aData.nActive), 500, Val(txtPalSpeed.Text), Val(Left(cmbPalAnimIPScale.Text, 2))
Else
    DoPalAnim Frac(aData.nActive), 500, Val(txtPalSpeed.Text)
End If

End Sub

Private Sub btnPalInvert_Click()
Dim Cnt As Long

With Frac(aData.nActive)
    For Cnt = 0 To .nPalLen - 1
        .Pal(Cnt).R = 255 - .Pal(Cnt).R
        .Pal(Cnt).G = 255 - .Pal(Cnt).G
        .Pal(Cnt).B = 255 - .Pal(Cnt).B
    Next Cnt
End With
ShowParams Frac(aData.nActive)
DoEvents
Calculate aData.nActive, False

End Sub

Private Sub btnPalMakeRnd_Click()
Dim nLen As Long, nPeriod As Long

nLen = Val(txtPalSize.Text)

If nLen < Val(txtPalPeriod.Text) Then nLen = Val(txtPalPeriod.Text)

nPeriod = Val(txtPalPeriod.Text)
If nPeriod < 1 Then nPeriod = 1
txtPalPeriod.Text = Trim(Str(nPeriod))
Frac(aData.nActive).nColPeriod = nPeriod

MakePalette Frac(aData.nActive), nLen
txtPalSize.Text = Trim(Str(nLen))
DisplayPal Frac(aData.nActive)
Calculate aData.nActive, False

End Sub

Private Sub btnPalOpen_Click()
Dim szFile As String, fNum As Long
Dim Buf() As Byte, Cnt As Long

szFile = GetOpenFilePath("Load Palette", "Palette Files|*.pal|All Files|*.*")
If szFile = "" Then Exit Sub

fNum = FreeFile()
ReDim Buf(FileLen(szFile))
Open szFile For Binary As fNum
    Get fNum, 1, Buf
Close fNum

ReDim Frac(aData.nActive).Pal(UBound(Buf) \ 3)
For Cnt = 0 To (UBound(Buf) - 1) Step 3
    Frac(aData.nActive).Pal(Cnt \ 3).R = Buf(Cnt)
    Frac(aData.nActive).Pal(Cnt \ 3).G = Buf(Cnt + 1)
    Frac(aData.nActive).Pal(Cnt \ 3).B = Buf(Cnt + 2)
Next Cnt
Frac(aData.nActive).nPalLen = (UBound(Buf) \ 3)
ShowParams Frac(aData.nActive)
Calculate aData.nActive, False

End Sub

Private Sub btnPalRnd_Click()
Dim Cnt As Long

With Frac(aData.nActive)
    .nPalLen = Val(txtPalSize.Text)
    ReDim .Pal(0 To .nPalLen - 1)
    For Cnt = 0 To .nPalLen - 1
        .Pal(Cnt).R = Rnd * 255
        .Pal(Cnt).G = Rnd * 255
        .Pal(Cnt).B = Rnd * 255
    Next Cnt
End With
DisplayPal Frac(aData.nActive)
Calculate aData.nActive, False

End Sub

Private Sub btnPalSave_Click()
Dim szFile As String

szFile = GetSaveFilePath("Save Palette", "Palette Files|*.pal|All Files|*.*")
If szFile = "" Then Exit Sub
SavePalette Frac(aData.nActive), szFile

End Sub


Private Sub btnPosDefault_Click()

With Frac(aData.nActive)
    .SX = -2
    .EX = 2
    .SY = -2
    .EY = 2
End With

txtPosSX.Text = "-2"
txtPosSY.Text = "-2"
txtPosEX.Text = "2"
txtPosEY.Text = "2"

Calculate aData.nActive, True

End Sub

Private Sub btnRefreshParams_Click()

ShowParams Frac(aData.nActive)

End Sub

Private Sub btnRemIn_Click()
Dim nIndex As Integer
Dim Cnt As Integer

If lstFiltersIn.ListItems.Count < 1 Then Exit Sub
nIndex = lstFiltersIn.SelectedItem.Index - 1

For Cnt = nIndex To Frac(aData.nActive).nFiltersIn - 2
    Frac(aData.nActive).FilterIn(Cnt) = Frac(aData.nActive).FilterIn(Cnt + 1)
Next Cnt
If Frac(aData.nActive).nFiltersIn > 1 Then
    ReDim Preserve Frac(aData.nActive).FilterIn(0 To Frac(aData.nActive).nFiltersIn - 2)
    Frac(aData.nActive).nFiltersIn = Frac(aData.nActive).nFiltersIn - 1
Else
    ReDim Frac(aData.nActive).FilterIn(0)
    Frac(aData.nActive).nFiltersIn = 0
End If
ShowParams Frac(aData.nActive)

End Sub

Private Sub btnRemOut_Click()
Dim nIndex As Integer
Dim Cnt As Integer

If lstFiltersOut.ListItems.Count < 1 Then Exit Sub
nIndex = lstFiltersOut.SelectedItem.Index - 1

For Cnt = nIndex To Frac(aData.nActive).nFiltersOut - 2
    Frac(aData.nActive).FilterOut(Cnt) = Frac(aData.nActive).FilterOut(Cnt + 1)
Next Cnt
If Frac(aData.nActive).nFiltersOut > 1 Then
    ReDim Preserve Frac(aData.nActive).FilterOut(0 To Frac(aData.nActive).nFiltersOut - 2)
    Frac(aData.nActive).nFiltersOut = Frac(aData.nActive).nFiltersOut - 1
Else
    ReDim Frac(aData.nActive).FilterOut(0)
    Frac(aData.nActive).nFiltersOut = 0
End If
ShowParams Frac(aData.nActive)

End Sub

Private Sub btnRenderToFile_Click()

Dim szFile As String

szFile = GetSaveFilePath("Render 2 File - Specify File", "Windows Bitmaps|*.bmp")
If szFile <> "" Then
    Frac(aData.nActive).szPathCalcDirect = szFile
    Frac(aData.nActive).Calc2File = True
    Calculate aData.nActive, True
End If

End Sub

Private Sub btnUpIn_Click()
Dim nIndex As Long
Dim Tmp As tFilterVars

If lstFiltersIn.ListItems.Count < 1 Then Exit Sub
nIndex = lstFiltersIn.SelectedItem.Index - 1

With Frac(aData.nActive)
    If nIndex > 0 And .nFiltersIn > 1 Then
        'Swap this filter with the one below:
        Tmp = .FilterIn(nIndex - 1)
        .FilterIn(nIndex - 1) = .FilterIn(nIndex)
        .FilterIn(nIndex) = Tmp
        ShowParams Frac(aData.nActive)
        Set lstFiltersIn.SelectedItem = lstFiltersIn.ListItems(nIndex)
    End If
End With

End Sub

Private Sub btnUpOut_Click()
Dim nIndex As Long
Dim Tmp As tFilterVars

If lstFiltersOut.ListItems.Count < 1 Then Exit Sub
nIndex = lstFiltersOut.SelectedItem.Index - 1

With Frac(aData.nActive)
    If nIndex > 0 And .nFiltersOut > 1 Then
        'Swap this filter with the one above:
        Tmp = .FilterOut(nIndex - 1)
        .FilterOut(nIndex - 1) = .FilterOut(nIndex)
        .FilterOut(nIndex) = Tmp
        ShowParams Frac(aData.nActive)
        lstFiltersOut.ListItems(nIndex).Selected = True
        Set lstFiltersOut.SelectedItem = lstFiltersOut.ListItems(nIndex)
    End If
End With

End Sub

Private Sub btnZoomClear_Click()

Frac(aData.nActive).Zoom(1) = Frac(aData.nActive).Zoom(Frac(aData.nActive).nZooms - 1)
ReDim Preserve Frac(aData.nActive).Zoom(2)
Frac(aData.nActive).nZooms = 2
ShowParams Frac(aData.nActive)

End Sub

Private Sub chkCustomOutputSize_Click()
If chkCustomOutputSize.Value = 1 Then
    If Val(txtW.Text) < 100 Then txtW.Text = "100"
    If Val(txtH.Text) < 100 Then txtH.Text = "100"
    ResizeFractal Frac(aData.nActive), Val(txtW.Text), Val(txtH.Text), True
Else
    cmbSize_Click
End If
End Sub

Private Sub chkLinearCon_Click()

If chkLinearCon.Value = 1 Then
    Frac(aData.nActive).LinearContinuity = True
Else
    Frac(aData.nActive).LinearContinuity = False
End If

End Sub

Private Sub chkNormalizeFrame_Click()

If chkNormalizeFrame.Value = vbChecked Then
    optNormIntDiff.Enabled = True
    optNormIntSpec.Enabled = True
    txtNormDiff.Enabled = True
    Frac(aData.nActive).bPalAnimNormalize = True
Else
    optNormIntDiff.Enabled = False
    optNormIntSpec.Enabled = False
    txtNormDiff.Enabled = False
    Frac(aData.nActive).bPalAnimNormalize = False
End If

End Sub

Private Sub chkPalAnimInterpolate_Click()

If chkPalAnimInterpolate.Value = 1 Then
    lblPalAnimIP_Scale.Enabled = True
    cmbPalAnimIPScale.Enabled = True
Else
    lblPalAnimIP_Scale.Enabled = False
    cmbPalAnimIPScale.Enabled = False
End If

End Sub

Private Sub chkPalUseDisp_Click()

If chkPalUseDisp.Value = 1 Then
    Frac(aData.nActive).UseDispersionControl = True
Else
    Frac(aData.nActive).UseDispersionControl = False
End If

End Sub

Private Sub cmbFractType_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub lstFilters_Click()
Dim Cnt As Long

'For Cnt = 0 To lstFiltersOut.ListCount - 1
'    If lstFiltersOut.Selected(Cnt) Then
'        Frac(aData.nActive).Filter(Cnt).IsEnabled = True
'    Else
'        Frac(aData.nActive).Filter(Cnt).IsEnabled = False
'    End If
'Next Cnt

End Sub

Private Sub cmbMapIn_Click()
Dim nIndex As Integer, Map() As Integer, nMap As Integer, Cnt As Integer
Dim nFilter As Integer

nIndex = cmbMapIn.ListIndex
If nIndex < 0 Then Exit Sub

ReDim Map(0)
For Cnt = 0 To aData.nFilters - 1
    If aData.Filters(Cnt).nComp = FILTER_COMP_2DMAP Then
        'Add filter to list:
        Map(UBound(Map)) = Cnt
        ReDim Preserve Map(0 To UBound(Map) + 1)
    End If
Next Cnt

nFilter = Map(nIndex)

With Frac(aData.nActive).FilterMapIn
    .IsEnabled = True
    .nType = nFilter
    .nVars = 0
    Erase .Var
End With

Calculate aData.nActive, True

End Sub

Private Sub cmbMapIn_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbMapOut_Click()
Dim nIndex As Integer, Map() As Integer, nMap As Integer, Cnt As Integer
Dim nFilter As Integer

nIndex = cmbMapOut.ListIndex
If nIndex < 0 Then Exit Sub

ReDim Map(0)
For Cnt = 0 To aData.nFilters - 1
    If aData.Filters(Cnt).nComp = FILTER_COMP_2DMAP Then
        'Add filter to list:
        Map(UBound(Map)) = Cnt
        ReDim Preserve Map(0 To UBound(Map) + 1)
    End If
Next Cnt

nFilter = Map(nIndex)

With Frac(aData.nActive).FilterMapOut
    .IsEnabled = True
    .nType = nFilter
    .nVars = 0
    Erase .Var
End With

Calculate aData.nActive, True

End Sub

Private Sub cmbMapOut_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbPalAnimIPScale_KeyPress(KeyAscii As Integer)
'Prevent user from modifying:
KeyAscii = 0
End Sub

Private Sub cmbSize_Click()

Dim nSize As Long

Select Case cmbSize.ListIndex
Case -1
    Exit Sub
Case 0
    nSize = 150
Case 1
    nSize = 200
Case 2
    nSize = 250
Case 3
    nSize = 300
Case 4
    nSize = 350
Case 5
    nSize = 400
Case 6
    nSize = 500
Case 7
    nSize = 600
Case 8
    nSize = 650
Case 9
    nSize = 700
Case 10
    nSize = 800
Case 11
    nSize = 900
Case 12
    nSize = Val(txtW.Text)
End Select

If nSize < 100 Then nSize = 100
If cmbSize.ListIndex = 12 Then
    txtW.Text = Trim(Str(nSize))
End If
ResizeFractal Frac(aData.nActive), nSize, nSize, True

End Sub

Private Sub cmbSize_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
frmParent.mnuWindowShowparam.Visible = True
frmParent.mnuWindowSep01.Visible = True
Me.Visible = False
End Sub

Private Sub imgPal_Paint()
With Frac(aData.nActive).PalDisp
    BitBlt imgPal.hdc, 0, 0, .w, .H, .hdc, 0, 0, SRCCOPY
End With
End Sub

Private Sub lblCalc2FileCustomSize_Click()
If chkCustomOutputSize.Value = 1 Then
    chkCustomOutputSize.Value = 0
Else
    chkCustomOutputSize.Value = 1
End If
End Sub

Private Sub lstFiltersIn_DblClick()
Dim nIndex As Long

If lstFiltersIn.ListItems.Count < 1 Then Exit Sub
If Not lstFiltersIn.SelectedItem.Selected Then Exit Sub
nIndex = lstFiltersIn.SelectedItem.Index
EditFilter Frac(aData.nActive), False, False, lstFiltersIn.SelectedItem.Index - 1
ShowParams Frac(aData.nActive)
Set lstFiltersIn.SelectedItem = lstFiltersIn.ListItems(nIndex)

End Sub

Private Sub lstFiltersOut_DblClick()
Dim nIndex As Long

If lstFiltersOut.ListItems.Count < 1 Then Exit Sub
If Not lstFiltersOut.SelectedItem.Selected Then Exit Sub
nIndex = lstFiltersOut.SelectedItem.Index
EditFilter Frac(aData.nActive), False, True, lstFiltersOut.SelectedItem.Index - 1
ShowParams Frac(aData.nActive)
Set lstFiltersOut.SelectedItem = lstFiltersOut.ListItems(nIndex)

End Sub

Private Sub lstZooms_Click()

End Sub


Private Sub lstZoom_DblClick()
Dim nIndex As Long

If lstZoom.ListItems.Count < 1 Then Exit Sub
If Not lstZoom.SelectedItem.Selected Then Exit Sub
If lstZoom.SelectedItem.Index < 1 Then Exit Sub

nIndex = Frac(aData.nActive).nZooms - lstZoom.SelectedItem.Index

With Frac(aData.nActive)
    ReDim Preserve .Zoom(0 To nIndex + 1)
    .nZooms = nIndex + 1
    .SX = .Zoom(nIndex).SX
    .SY = .Zoom(nIndex).SY
    .EX = .Zoom(nIndex).EX
    .EY = .Zoom(nIndex).EY
End With
ShowParams Frac(aData.nActive)
Calculate aData.nActive, True

End Sub

Private Sub optNormIntDiff_Click()
If optNormIntDiff.Value = True Then
    Frac(aData.nActive).bPalAnimNormIntervalSpecified = False
    Frac(aData.nActive).nPalAnimNormInterval = Val(txtNormDiff.Text)
Else
    Frac(aData.nActive).bPalAnimNormIntervalSpecified = True
End If
End Sub

Private Sub optNormIntSpec_Click()
If optNormIntDiff.Value = True Then
    Frac(aData.nActive).bPalAnimNormIntervalSpecified = False
    Frac(aData.nActive).nPalAnimNormInterval = Val(txtNormDiff.Text)
Else
    Frac(aData.nActive).bPalAnimNormIntervalSpecified = True
End If
End Sub

Private Sub scrPalImg_Change()
DisplayPal Frac(aData.nActive)
End Sub

Private Sub scrPalImg_Scroll()
DisplayPal Frac(aData.nActive)
End Sub

Private Sub txtBailout_Change()
Frac(aData.nActive).BailOut = Val(txtBailout.Text)
End Sub

Private Sub txtCalc2File_Size_Change()
If Val(txtCalc2File_Size.Text) < 1 Then txtCalc2File_Size.Text = Trim(Str(Frac(aData.nActive).nW))
Frac(aData.nActive).nDirectFileSize = Val(txtCalc2File_Size.Text)
End Sub

Private Sub txtCalcP_Change(Index As Integer)
Frac(aData.nActive).Param(Index) = Val(txtCalcP(Index).Text)
End Sub

Private Sub txtMaxIterations_Change()
Frac(aData.nActive).MaxI = Val(txtMaxIterations.Text)
End Sub

Private Sub txtNormDiff_Change()
Frac(aData.nActive).nPalAnimNormInterval = Val(txtNormDiff.Text)
End Sub

Private Sub cmbFractType_Click()
Dim nFrac As Long

'Get the current fractal (number):
nFrac = aData.nActive

'Remove all filters:
ClearFilters (nFrac)

'Remove zoom/move history:
ReDim Preserve Frac(aData.nActive).Zoom(0)
Frac(aData.nActive).nZooms = 0

With Frac(nFrac)
    Select Case cmbFractType.ListIndex
    Case 0
        .SX = -2: .SY = -2: .EX = 2: .EY = 2
        .nType = TYPE_MANDELBROT
        .Param(0) = 0
        .Param(1) = 0
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_COS2SIN2, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_TRIG_SWIRL, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_COS_X0, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
        AddFilter Frac(nFrac), FILTER2D_MUL, False, 20, 0
    Case 1
        .SX = -2: .SY = -2: .EX = 2: .EY = 2
        .nType = TYPE_JULIA
        .Param(0) = 0.4
        .Param(1) = 0.3
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_COS2SIN2, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_TRIG_SWIRL, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_COS_X0, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
        AddFilter Frac(nFrac), FILTER2D_MUL, False, 20, 0
    Case 2
        .SX = -2: .SY = -2: .EX = 2: .EY = 2
        .nType = TYPE_MANDEL3
        .Param(0) = 0
        .Param(1) = 0
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_COS2SIN2, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_TRIG_SWIRL, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_COS_X0, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
        AddFilter Frac(nFrac), FILTER2D_MUL, False, 20, 0
    Case 3
        .SX = -2: .SY = -2: .EX = 2: .EY = 2
        .nType = TYPE_MANDEL4
        .Param(0) = 0
        .Param(1) = 0
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_COS2SIN2, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_TRIG_SWIRL, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_COS_X0, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
        AddFilter Frac(nFrac), FILTER2D_MUL, False, 20, 0
    Case 4
        .SX = -2: .SY = -2: .EX = 2: .EY = 2
        .nType = TYPE_MANDELRINGS
        .Param(0) = 0.02
        .Param(1) = 0
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 12
    Case 5
        .SX = -2: .SY = -2: .EX = 2: .EY = 2
        .nType = TYPE_BARNSLEYM1
        .Param(0) = 0
        .Param(1) = 0
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
        AddFilter Frac(nFrac), FILTER2D_MUL, False, 20, 0
    Case 6
        .SX = -2: .SY = -2: .EX = 2: .EY = 2
        .nType = TYPE_BARNSLEYM4
        .Param(0) = 0
        .Param(1) = 1
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
        AddFilter Frac(nFrac), FILTER2D_MUL, False, 20, 0
    Case 7
        .SX = -3.5: .SY = -3.5: .EX = 3.5: .EY = 3.5
        .nType = TYPE_BARNSLEYJ1
        .Param(0) = 1
        .Param(1) = 0.1
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_COS2SIN2, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_TRIG_SWIRL, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_COS_X0, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
        AddFilter Frac(nFrac), FILTER2D_MUL, False, 20, 0
    Case 8
        .SX = -3.5: .SY = -3.5: .EX = 3.5: .EY = 3.5
        .nType = TYPE_BARNSLEYJ2
        .Param(0) = 0.7
        .Param(1) = 0.73
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_COS2SIN2, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_TRIG_SWIRL, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_COS_X0, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
        AddFilter Frac(nFrac), FILTER2D_MUL, False, 20, 0
    Case 9
        .SX = -2: .SY = -2: .EX = 2: .EY = 2
        .nType = TYPE_BARNSLEYJ3
        .Param(0) = 0
        .Param(1) = 0
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_COS2SIN2, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_TRIG_SWIRL, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_COS_X0, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
        AddFilter Frac(nFrac), FILTER2D_MUL, False, 20, 0
    Case 10
        .SX = -2: .SY = -2: .EX = 2: .EY = 2
        .nType = TYPE_MANDELZPOW
        .Param(0) = 0
        .Param(1) = 0
        .Param(2) = 2
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_COS2SIN2, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_TRIG_SWIRL, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_COS_X0, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
        AddFilter Frac(nFrac), FILTER2D_MUL, False, 20, 0
    Case 11
        .SX = -2: .SY = -2: .EX = 2: .EY = 2
        .nType = TYPE_JULIAZPOW
        .Param(0) = 0.4
        .Param(1) = 0.3
        .Param(2) = 2
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_COS2SIN2, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_TRIG_SWIRL, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_COS_X0, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
        AddFilter Frac(nFrac), FILTER2D_MUL, False, 20, 0
    Case 12
        .SX = -2: .SY = -2: .EX = 2: .EY = 2
        .nType = TYPE_BARNSLEYMANJUL
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_COS2SIN2, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_TRIG_SWIRL, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_COS_X0, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
        AddFilter Frac(nFrac), FILTER2D_MUL, False, 20, 0
    Case 13
        .SX = -2: .SY = -2: .EX = 2: .EY = 2
        .nType = TYPE_SPIDER
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_COS2SIN2, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_TRIG_SWIRL, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_COS_X0, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
        AddFilter Frac(nFrac), FILTER2D_MUL, False, 20, 0
    Case 14
        .SX = -2: .SY = -3: .EX = 4: .EY = 3
        .nType = TYPE_MANDELLAMBDA
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_COS2SIN2, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_TRIG_SWIRL, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_COS_X0, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
        AddFilter Frac(nFrac), FILTER2D_MUL, False, 20, 0
    Case 15
        .SX = -2: .SY = -2: .EX = 2: .EY = 2
        .nType = TYPE_SIERPINSKI
        .BailOut = 4
        .FilterMapOut.nType = FILTER2DMAP_CONTINOUS_I
        AddFilter Frac(nFrac), FILTER2D_COS2SIN2, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_TRIG_SWIRL, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_COS_X0, True, 1, 1
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 8, 0
        AddFilter Frac(nFrac), FILTER2D_MUL, False, 20, 0
    Case 16
        .SX = -2: .SY = -2: .EX = 2: .EY = 2
        .Param(0) = 3
        .BailOut = 1E-20 '0.0000001
        .MaxI = 700
        .nType = TYPE_NEWTON
        .FilterMapOut.nType = FILTER2DMAP_DIRECT_I
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 5, 0
    Case 17
        .SX = -2: .SY = -2: .EX = 2: .EY = 2
        .Param(0) = 0
        .BailOut = 0.000001
        .MaxI = 80
        .nType = TYPE_NEWTONJULIANOVA
        .FilterMapOut.nType = FILTER2DMAP_DIRECT_I
        AddFilter Frac(nFrac), FILTER2D_MUL, True, 5, 0
    End Select
End With

If XXT = 1 Then Exit Sub


AddZoomCurrent aData.nActive, 0
'If Not aData.Initializing Then
    ShowParams Frac(aData.nActive)
    Calculate aData.nActive, True
'End If

End Sub
