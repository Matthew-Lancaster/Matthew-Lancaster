VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmPassLock 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "PassLock"
   ClientHeight    =   12000
   ClientLeft      =   0
   ClientTop       =   108
   ClientWidth     =   19200
   FillStyle       =   0  'Solid
   Icon            =   "frmPassLock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1000
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox Check7 
      BackColor       =   &H00FF0000&
      Caption         =   "Bezier ##"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1875
      TabIndex        =   51
      Top             =   4185
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer ControlsChangesTimer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   420
      Top             =   3885
   End
   Begin VB.Timer ControlsChangesTimer 
      Interval        =   1000
      Left            =   390
      Top             =   3435
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1875
      TabIndex        =   50
      Text            =   "Combo1"
      Top             =   5385
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00FF0000&
      Caption         =   "Sine Color Kelei"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1875
      TabIndex        =   47
      Top             =   5085
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   165
      Top             =   2355
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   195
      Top             =   1470
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   180
      Top             =   1050
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   180
      Top             =   630
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00FF0000&
      Caption         =   "Sine On Kelei"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1875
      TabIndex        =   22
      Top             =   4785
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF0000&
      Caption         =   "Exited Kelei"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1875
      TabIndex        =   21
      Top             =   4485
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1275
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Colour Sine"
      Top             =   2325
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1275
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Seed Colour"
      Top             =   2010
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00FF0000&
      Caption         =   "Line"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1875
      TabIndex        =   15
      Top             =   3885
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FF0000&
      Caption         =   "Circle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1875
      TabIndex        =   14
      Top             =   3585
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FF0000&
      Caption         =   "Dot"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1875
      TabIndex        =   13
      Top             =   3285
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1275
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Spin Speed"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1275
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Draw Width"
      Top             =   1365
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1275
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Zoom In Out"
      Top             =   1035
      Visible         =   0   'False
      Width           =   1200
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   315
      Left            =   2610
      TabIndex        =   9
      ToolTipText     =   "Slider3"
      Top             =   1380
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2180
      _ExtentY        =   550
      _Version        =   393216
      LargeChange     =   1
      Max             =   1000
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   315
      Left            =   2610
      TabIndex        =   8
      ToolTipText     =   "Slider2"
      Top             =   1050
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2180
      _ExtentY        =   550
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   100
      SelStart        =   10
      TickStyle       =   3
      Value           =   10
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   315
      Left            =   2595
      TabIndex        =   7
      ToolTipText     =   "Slider1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2180
      _ExtentY        =   550
      _Version        =   393216
      LargeChange     =   1
      Min             =   -2000
      Max             =   2000
      SelStart        =   9
      TickStyle       =   3
      Value           =   9
      TextPosition    =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5745
      Left            =   3930
      TabIndex        =   0
      Top             =   420
      Width           =   5955
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   624
         Left            =   30
         TabIndex        =   49
         Top             =   4275
         Width           =   3690
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   5250
         TabIndex        =   16
         Top             =   405
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1500
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "frmPassLock.frx":0442
         Top             =   5250
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "User"
         Top             =   5250
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Pass"
         Top             =   5250
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2895
         PasswordChar    =   "*"
         TabIndex        =   2
         Text            =   "***"
         Top             =   5250
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.TextBox txtUser 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   750
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Matt"
         Top             =   5250
         Visible         =   0   'False
         Width           =   750
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   210
         Left            =   15
         TabIndex        =   23
         Top             =   3525
         Width           =   3705
         _ExtentX        =   6541
         _ExtentY        =   360
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBarVol 
         Height          =   255
         Left            =   15
         TabIndex        =   43
         Top             =   1035
         Width           =   3225
         _ExtentX        =   5694
         _ExtentY        =   445
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBarLock 
         Height          =   255
         Left            =   15
         TabIndex        =   44
         Top             =   1290
         Width           =   3225
         _ExtentX        =   5694
         _ExtentY        =   445
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "System UpTime"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   30
         TabIndex        =   48
         ToolTipText     =   "Click to Speedup the Time Double Click Resets"
         Top             =   4035
         Width           =   3690
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00680202&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   46
         Top             =   1290
         Width           =   465
      End
      Begin VB.Label LblVol 
         Alignment       =   2  'Center
         BackColor       =   &H00680202&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   45
         Top             =   1035
         Width           =   465
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "12345"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   30
         TabIndex        =   42
         ToolTipText     =   "Click to Speedup the Time Double Click Resets"
         Top             =   780
         Width           =   3675
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "12345"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   30
         TabIndex        =   41
         ToolTipText     =   "Click to Speedup the Time Double Click Resets"
         Top             =   525
         Width           =   3675
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "12345"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   30
         TabIndex        =   40
         ToolTipText     =   "Click to Speedup the Time Double Click Resets"
         Top             =   270
         Width           =   3675
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Hit a Key"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   2535
         TabIndex        =   39
         ToolTipText     =   "F12 to Command Prompt - Disable"
         Top             =   4965
         Width           =   1185
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Keys"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1905
         TabIndex        =   38
         ToolTipText     =   "View keyboard and Mouse Log With Notepad"
         Top             =   4965
         Width           =   630
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Move"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   885
         TabIndex        =   37
         Top             =   4950
         Width           =   1005
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Mouse"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   45
         TabIndex        =   36
         ToolTipText     =   "Disable Hide Cursor"
         Top             =   4965
         Width           =   840
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Equinox"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   30
         TabIndex        =   35
         ToolTipText     =   "Click to Speedup the Time Double Click Resets"
         Top             =   2310
         Width           =   3675
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pay Day"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   30
         TabIndex        =   34
         ToolTipText     =   "Click to Speedup the Time Double Click Resets"
         Top             =   2055
         Width           =   3675
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Days This Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   30
         TabIndex        =   33
         ToolTipText     =   "Click to Speedup the Time Double Click Resets"
         Top             =   1800
         Width           =   3675
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Week No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   30
         TabIndex        =   32
         ToolTipText     =   "Click to Speedup the Time Double Click Resets"
         Top             =   1545
         Width           =   3675
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Next First Quarter "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   30
         TabIndex        =   31
         Top             =   3210
         Width           =   1965
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Moon Luminosity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   30
         TabIndex        =   30
         Top             =   2565
         Width           =   1965
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Now First Quarter "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   30
         TabIndex        =   29
         Top             =   2895
         Width           =   1965
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "12Jan hh:MM:SS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   2010
         TabIndex        =   28
         Top             =   3210
         Width           =   1710
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "000.00000%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   2010
         TabIndex        =   27
         Top             =   2565
         Width           =   1710
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "000.0000000%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1905
         TabIndex        =   26
         Top             =   3750
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "000.0000000%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   30
         TabIndex        =   25
         Top             =   3750
         Width           =   1860
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "12Jan hh:MM:SS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   2010
         TabIndex        =   24
         Top             =   2895
         Width           =   1710
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   5400
         TabIndex        =   3
         Top             =   1065
         Width           =   270
      End
   End
   Begin VB.Timer timPause 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   180
      Top             =   1890
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   195
      Top             =   225
   End
   Begin MSComctlLib.Slider Slider4 
      Height          =   315
      Left            =   2610
      TabIndex        =   17
      ToolTipText     =   "Slider4"
      Top             =   1710
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2180
      _ExtentY        =   550
      _Version        =   393216
      LargeChange     =   1
      Max             =   1000
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin MSComctlLib.Slider Slider5 
      Height          =   315
      Left            =   2610
      TabIndex        =   19
      ToolTipText     =   "Slider5"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2180
      _ExtentY        =   550
      _Version        =   393216
      LargeChange     =   1
      Max             =   1000
      SelStart        =   2
      TickStyle       =   3
      Value           =   2
   End
End
Attribute VB_Name = "frmPassLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const ProgramVersion = "1.0.1010"

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Dim OldControlValueChanges
Dim Th(200)
Dim OldCombo1ListIndex
'OldCombo1ListIndex
'Public MLeg

Public CIDonTop
Const ShowPass = True
'Const ShowPass = False

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Et, KKd, Yui, FirstPassLoad

Private inMessagebox

Private Sub DisplayMessage(ByVal msg As String)
    lblStatus.Caption = msg
    lblStatus.Left = Frame1.Width / 2 - lblStatus.Width / 2
    timPause.Enabled = True
End Sub



Sub SaveSlidersKKd()
    Exit Sub
End Sub




Sub DChcks()
    
'Call SubSwitches
Keleidoscope.HighRes.Enabled = False
Bezier.Timer1.Enabled = False
YinYang.Timer1.Enabled = False
Vector.Timer1.Enabled = False
Vector.Timer2.Enabled = False
frmPassLock.AutoRedraw = True
ExitVector = False

Text6.Text = "Zoom In Out"
Text7.Text = "Draw Width"
Text8.Text = "Spin Speed"
Text10.Text = "Colour Seed"
Text11.Text = "Colour Sine"

End Sub

Sub StartTheWhirl()

If App.title = "Tidal Information..." Then Exit Sub
'frmPassLock.Hide

yy$ = YinVectKeli$
'Keleidoscope.HighRes.Enabled = False
'YinYang.Timer1.Enabled = False
'Vector.Timer1.Enabled = False
'Vector.Timer2.Enabled = False
'frmPassLock.AutoRedraw = False
'Text8.Text = "Spin Speed"
'Text10.Text = "Seed Colour"
Call DChcks 'Falses all timers
Call SubSwitches 'Sets Timer To save
Call SubSwitches2 'Set timer to save an cls

If yy$ = "Bezier" Then
   Bezier.Timer1.Enabled = True
    frmPassLock.AutoRedraw = False
End If

If yy$ = "Keleidoscope" Then
    Keleidoscope.HighRes.Enabled = True
    frmPassLock.AutoRedraw = False
End If

If yy$ = "Yin Yang" Then
    YinYang.Timer1.Enabled = True
    Text8.Text = "Spin Speed 1"
    Text10.Text = "Spin Speed 2"
    frmPassLock.AutoRedraw = False
End If

If yy$ = "Vector Pattern" Then
    Vector.Timer1.Enabled = True
    Vector.Timer2.Enabled = True
'    frmPassLock.AutoRedraw = True
'    frmPassLock.Cls
    frmPassLock.AutoRedraw = False
End If
'frmPassLock.Show

End Sub


Private Sub Combo1_Change()
Call Combo1_Click
End Sub

Private Sub Combo1_Click()

If Combo1.Enabled = False Then Exit Sub
'frmPassLock.Hide

If App.title = "Tidal Information..." Then Exit Sub

'If OldCombo1ListIndex <> Combo1.ListIndex Then MsgBox = Str(Combo1.ListIndex)
If OldCombo1ListIndex = Combo1.ListIndex And OldCombo1ListIndex <> vbNullString Then Exit Sub

zzSave_Checks

yy$ = Combo1.List(Combo1.ListIndex)
YinVectKeli$ = Combo1.List(Combo1.ListIndex)

'zzSave_Checks

OldCombo1ListIndex = Combo1.ListIndex

Call zzLoadChecks 'Only way to set sliders at moment

Call StartTheWhirl
'frmPassLock.Show

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button <> 2 Or MLeg = 1 Then Exit Sub
'Call KnockerLogg
'MLeg = 1
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MLeg = 0
End Sub

Private Sub Label10_Click()
    Call ATidalX.Label25_Click
End Sub

Private Sub Label11_Click()
    Call ATidalX.Label11_Click
End Sub

Private Sub Label25_Click()
    Call ATidalX.Label25_Click
End Sub

Private Sub Label24_Click()
    Call ATidalX.Label24_Click
End Sub


Private Sub Label40_Click()
'frm PassLock.Visible = Not frmPassLock.Visible
Call ATidalX.Mnu_ELock_Off_Click
End Sub

Private Sub Label42_Click()
'If IsIDE = False Then Exit Sub
'frmPassLock.Visible = Not frmPassLock.Visible
Call ATidalX.Mnu_ELock_Off_Click
End Sub

Private Sub Option1_Click()

'this clears the control panel after bit time on use

If Option1.Value = False Or FirstPassLoad = True Then
    For Each Control In frmPassLock.Controls
        XX = 0
        If Control.Name = "Text1" Then XX = 1
        If InStr(Control.Name, "Option") Then XX = 1
        If InStr(Control.Name, "Frame1") Then XX = 1
        If InStr(Control.Name, "Timer") Then XX = 1
        If InStr(Control.Name, "timPause") Then XX = 1
        If InStr(Control.Name, "Label") Then XX = 1
        If InStr(Control.Name, "Progress") Then XX = 1
        If InStr(Control.Name, "LblVol") Then XX = 1
        If InStr(Control.Name, "List") Then XX = 1
        If XX = 0 Then Control.Visible = False
    Next
    Frame1.Height = (txtUser.Top) / Screen.TwipsPerPixelY + 3
    Exit Sub
End If

If Option1.Value = True Then
    For Each Control In frmPassLock.Controls
        XX = 0
        If Control.Name = "Text1" Then XX = 1
        If InStr(Control.Name, "Option") Then XX = 1
        If InStr(Control.Name, "Frame1") Then XX = 1
        If InStr(Control.Name, "Timer") Then XX = 1
        If InStr(Control.Name, "timPause") Then XX = 1
        If XX = 0 Then Control.Visible = True
    Next

    Frame1.Height = (txtUser.Top + txtUser.Height) / Screen.TwipsPerPixelY + 3
End If

End Sub

Private Sub Check1_Click()
Call SubSwitches
End Sub
Private Sub Check2_Click()
If Et = 1 Then Exit Sub
Call Checks
Et = 1: Check2.Value = vbChecked
Et = 0
End Sub

Private Sub Check3_Click()
If Et = 1 Then Exit Sub
Call Checks
Et = 1: Check3.Value = vbChecked
Et = 0
End Sub

Private Sub Check4_Click()
If Et = 1 Then Exit Sub
Call Checks
Et = 1: Check4.Value = vbChecked
Et = 0
End Sub

Private Sub Check7_Click()
If Et = 1 Then Exit Sub
Call Checks
Et = 1: Check7.Value = vbChecked
Et = 0
End Sub

Sub Checks()
    For Each Control In frmPassLock.Controls
        XX = 0
        If InStr(Control.Name, "Check") Then XX = 1
        If InStr(Control.Name, "Check1") Then XX = 0
        If InStr(Control.Name, "Check5") Then XX = 0
        If InStr(Control.Name, "Check6") Then XX = 0
        Et = 1
        If XX = 1 Then Control.Value = vbUnchecked
        Et = 0
    Next
Call SubSwitches
frmPassLock.Cls
End Sub



Sub SubSwitches()
    Yui = Now + TimeSerial(0, 0, 20)
End Sub

Sub SubSwitches2()
    Yui = Now + TimeSerial(0, 0, 20)
    Me.Cls
    frmPassLock.Cls
    Keleidoscope.Cls
    YinYang.Cls
    Vector.Cls
    Me.Refresh
    XPud = False
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SubSwitches2
End Sub

Private Sub Slider2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SubSwitches2
End Sub

Private Sub Slider3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SubSwitches2
End Sub

Private Sub Slider4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SubSwitches2
End Sub

Private Sub Slider5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SubSwitches2
End Sub

Private Sub Slider1_Change()
    Call SubSwitches
End Sub

Private Sub Slider2_Change()
    Call SubSwitches
End Sub

Private Sub Slider3_Change()
    Call SubSwitches
End Sub

Private Sub Slider4_Change()
    Call SubSwitches
End Sub

Private Sub Slider5_Change()
    Call SubSwitches
End Sub

'Private Sub Slider2_DragOver(Source As Control, x As Single, y As Single, State As Integer)
'frmPassLock.DrawWidth = Slider2.Value / 5
'End Sub


Private Sub Text5_Click()
'Frame1.Hide
'Unload frmPassLock
If IsIDE = False Then Exit Sub

If LockDontAllowHotKeyforPassword = True Then Exit Sub

'If IsIDE = False Then Exit Sub

'Call zzSave_Checks

'strName = String$(255, 0)
'lngBuffer = GetUserName(strName, Len(strName))
Keleidoscope.HighRes.Enabled = False
Bezier.Timer1.Enabled = False
YinYang.Timer1.Enabled = False
Vector.Timer1.Enabled = False
Vector.Timer2.Enabled = False
ExitVector = False
Unload frmPassLock
'Call zzSave_Checks this is done in the unload

DoEvents

YinVectKeli$ = ""


'Unload Keleidoscope
'Call Form_Unload

'MsgBox Format$(XxG + TimeSerial(0, 0, 5), "DD:MM:YY HH:MM:SS") + " --- " + Format$(Now, "DD:MM:YY HH:MM:SS") + " -- " + Str(XxG + TimeSerial(0, 0, 5) > Now)

GoTo jump
'IDE
If XxG + TimeSerial(0, 0, 10) > Now Then
    LockSSaverDelay = LockSaverDelay2
Else
    LockSSaverDelay = LockSaverDelay1
End If
'setlocktim

If LockDown = True Then LockSSaverDelay = LockSaverDelay3

LockSSaver = Now + LockSaverDelay

jump:

End Sub
Private Sub Label41_Click()
If IsIDE = True Then End
'Frame1.Hide
End Sub

Sub Labels_Tidal()


LblVol = ATidalX.LblVol

Label2.Caption = ATidalX.Label2.Caption
Label3.Caption = ATidalX.Label3.Caption


Label10.Caption = ATidalX.Label10.Caption
Label11.Caption = ATidalX.Label11.Caption
Label12.Caption = ATidalX.Label12.Caption
'Label20.Caption = ATidalX.Label20.Caption
Label21.Caption = ATidalX.Label21.Caption
'Label22.Caption = ATidalX.Label22.Caption
Label23.Caption = ATidalX.Label23.Caption
Label24.Caption = ATidalX.Label24.Caption
Label25.Caption = ATidalX.Label25.Caption
Label26.Caption = ATidalX.Label26.Caption
Label27.Caption = ATidalX.Label27.Caption
Label28.Caption = ATidalX.Label28.Caption
Label29.Caption = ATidalX.Label29.Caption
Label30.Caption = ATidalX.Label30.Caption
'If ATidalX.ProgressBar1.Value > 100 Then Exit Sub

'frmPassLock.ProgressBar1.Min = 0
'frmPassLock.ProgressBar1.Max = ATidalX.ProgressBar1.Max
'frmPassLock.ProgressBar1.Min = ATidalX.ProgressBar1.Min

On Local Error Resume Next
frmPassLock.ProgressBar1.Value = ATidalX.ProgressBar1.Value

End Sub

Private Sub Timer1_Timer()

If App.title = "Tidal Information..." Then
Unload Me
Exit Sub
End If

If FindWindow(vbNullString, "Windows Task Manager") > 0 Then

Dim Rect4 As RECT

hWnd2 = FindWindow(vbNullString, "Windows Task Manager")
hwnd3 = GetWindowRect(hWnd2, Rect4)

If Rect4.Top > 0 Then
If IsIDE = False Then ShowWindow hWnd2, SW_MINIMIZE
    ShowWindow hWnd2, SW_MINIMIZE
End If

End If


'This Will Happen to Mouse If User Is Logged off
'If nx = 0 And ny = 0 Then
'LockSSaver = Now + LockSaverDelay
'Call SetLockMax
'Call ProgressLock
'End If

Dim Nx As Long
Dim Ny As Long

ATidalX.FindCursor Nx, Ny

End Sub

Private Sub Timer2_Timer()

Label40 = "Time :-- " + Format$(Now, "DDD DD-MMM -- HH:MM:SSa/p")
Label41 = "Lock :-- " + Format$(XxG, "DDD DD-MMM -- HH:MM:SSa/p")
Label42 = "Lock T-Plus:- " + Format$(Now - XxG, "HH:MM:SS")

If Yui < Now Then
If Option1.Value = True Then Option1.Value = False
    Call Option1_Click
End If


End Sub



Private Function ReturnPassCLue()
Exit Function
   Dim PassClue
      PassClue = Left$(PasswordInMemory.strPassword, 1)
      For i = 2 To Len(PasswordInMemory.strPassword)
        PassClue = PassClue & "*"
      Next i
      ReturnPassCLue = PassClue
      
End Function



Sub ComboChange()

End Sub

Private Sub Timer4_Timer()

Call ComboChange

Call Labels_Tidal

'frmPassLock.ProgressBarVol1.Value = ATidalX.ProgressBarVol1.Value

frmPassLock.DrawWidth = Slider2.Value    ' Use wider brush.
End Sub

Private Sub Timer5_Timer()
'Call ATidalX.SlowKeyPress2
'Call ATidalX.Timer7WinAmpVolIDE_Timer
End Sub
Private Sub ControlsChangesTimer2_Timer()

zzSave_Checks
ControlsChangesTimer2.Enabled = False
End Sub

Private Sub ControlsChangesTimer_Timer()

'frmPassLock.Hide
ControlValueChanges = 0
On Error Resume Next
For Each Control In Me.Controls
    Err.Clear
    A = Control.Value
    If Err.Number = 0 Then
        RR = Control.Name
        If InStr(Control.Name, "ProgressBar") = 0 Then
            ControlValueChanges = ControlValueChanges + Control.Value
        End If
    End If
Next

If OldControlValueChanges <> ControlValueChanges And OldControlValueChanges > 0 Then
    'frmPassLock.Hide
    ControlsChangesTimer2.Enabled = False
    ControlsChangesTimer2.Interval = 5000
    ControlsChangesTimer2.Enabled = True
End If

OldControlValueChanges = ControlValueChanges

End Sub

Private Sub txtPassword_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
'   If IsPasswordCorrect(txtPassword.Text, txtUser.Text) <> yesitdoes Then
    'frmPassLock.Hide
   If GetUserName <> lpStrUser Or txtPassword.Text <> lpStrPassword Then
    
    'DisplayMessage "Username/Password is incorrect: (" & ReturnPassCLue() & ")"
'    Text1.Text = "Try the LSD or Fuck Off an Die Like Extremely Slowly: (" & ReturnPassCLue() & ")"
    Text1.Text = "Try the LSD or Fuck Off an Die Like Extremely Slowly: (" & Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS") & ")"
    Text1.Visible = True
      Else
        Unload frmPassLock
        DoEvents
        If LockDown2 < Now Then LockDownX = False
   End If
End If

End Sub

Private Sub txtUser_GotFocus()
'If txtUser.Locked = True Then txtPassword.SetFocus
'txtPassword.SetFocus
End Sub



Private Sub Form_Activate()

'Form_Load
If App.title = "Tidal Information..." Then Exit Sub
Yui = Now + TimeSerial(0, 0, 30)

Me.Cls

'frmPassLock.ProgressBarVol.Min = 0
'frmPassLock.ProgressBarVol.Max = 100

Timer4.Enabled = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 95 Then KeyCode = 0




End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

'If txtPassword.Visible = True Then txtPassword.SetFocus Else frmPassLock.Label27.SetFocus
On Local Error Resume Next
txtPassword.SetFocus
'frmPassLock.Label27.SetFocus
End Sub




Public Sub Form_Load()

If App.title = "Tidal Information..." Then
Unload Me
Exit Sub
End If

'frmPassLock.DrawWidth = dw    ' Use wider brush.


FirstPassLoad = True

Frame1.Visible = False

Call zzLoadChecks

lngI = SetFocuses(frmPassLock.hwnd)

Call ComboChange

If txtUser.Text = "" Then txtUser.Locked = False

End Sub

Private Sub Redraw_Form()
'ATidalX.SetFocus'Temp

If App.title = "Tidal Information..." Then Exit Sub
ATidalX.Refresh
DoEvents

'If frmPassLock.hide Then Exit Sub

If frmPassLock.Visible = True Then frmPassLock.SetFocus
frmPassLock.Refresh

DoEvents

If ShowPass = True Then
    Me.Width = Screen.Width
    Me.Height = Screen.Height ' - 2000
End If

'Picture1.Top = 0
'Picture1.Left = 0
'Picture1.Height = Screen.Height
'Picture1.Width = Screen.Width

Me.Top = 0
Me.Left = 0
'Frame1.Caption = App.Title '+ " Hutt Hutt"

ActivateForm
Frame1.Top = 0
Frame1.Width = (Label40.Left + Label40.Width) / Screen.TwipsPerPixelX + 2
Frame1.Left = Me.ScaleWidth - Frame1.Width
'Call Option1_Click

CidRunMe





'slider3
Text8.Top = 0
Text8.Left = 0
'slider4
Text10.Top = Text8.Top + Text8.Height
Text10.Left = 0
'Slider5
Text11.Top = Text10.Top + Text10.Height
Text11.Left = 0
'slider1
Text6.Top = Text11.Top + Text11.Height
Text6.Left = 0
'slider2
Text7.Top = Text6.Top + Text6.Height
Text7.Left = 0

Check2.Top = Text7.Top + Text7.Height
Check2.Left = 0
Check3.Top = Check2.Top + Check2.Height
Check3.Left = 0
Check4.Top = Check3.Top + Check3.Height
Check4.Left = 0
Check7.Top = Check4.Top + Check4.Height
Check7.Left = 0
Check1.Top = Check7.Top + Check7.Height
Check1.Left = 0
Check5.Top = Check1.Top + Check1.Height
Check5.Left = 0
Check6.Top = Check5.Top + Check5.Height
Check6.Left = 0

Combo1.Top = Check6.Top + Check6.Height
Combo1.Left = 0

Slider3.Top = Text8.Top
Slider3.Left = Text6.Width
Slider3.Width = (Screen.Width / Screen.TwipsPerPixelX) - Frame1.Width - Text6.Width

Slider4.Top = Text10.Top
Slider4.Left = Text10.Width
Slider4.Width = (Screen.Width / Screen.TwipsPerPixelX) - Frame1.Width - Text6.Width

Slider5.Top = Text11.Top
Slider5.Left = Text6.Width
Slider5.Width = (Screen.Width / Screen.TwipsPerPixelX) - Frame1.Width - Text6.Width

Slider1.Top = Text6.Top
Slider1.Left = Text6.Width
Slider1.Width = (Screen.Width / Screen.TwipsPerPixelX) - Frame1.Width - Text6.Width

Slider2.Top = Text7.Top
Slider2.Left = Text6.Width
Slider2.Width = (Screen.Width / Screen.TwipsPerPixelX) - Frame1.Width - Text6.Width


If Sld1 > 0 And 1 = 2 Then
    On Local Error Resume Next
    Slider1.Value = Sld1
    Slider2.Value = Sld2
    Slider3.Value = Sld3
    Slider4.Value = Sld4
    Slider5.Value = Sld5
    Check1.Value = Chked1
    Check2.Value = Chked2
    Check3.Value = Chked3
    Check4.Value = Chked4
    Check5.Value = Chked5
    Check6.Value = Chked6
    DChck5.Value = DChked5
    DChck6.Value = DChked6
    DChck7.Value = DChked7
    If Err.Number > 0 Then Kill App.Path + "\Sliders1-2-3.txt"
    
    On Local Error GoTo 0
End If

Frame1.Caption = "Extreme Lock Build: 2011"

FirstPassLoad = False
Frame1.Height = (txtUser.Top) / Screen.TwipsPerPixelY + 3
Frame1.Visible = True

End Sub

Private Sub Form_Resize()
    Redraw_Form

End Sub

Private Sub Form_Click()

If FirstPassLoad = True Then Exit Sub
Call CidRunMe

If Option1.Value = False Then Option1.Value = True: Exit Sub
If Option1.Value = True Then Option1.Value = False
Call Option1_Click
Me.Cls
XPud = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Yui = Now + TimeSerial(0, 0, 30)
End Sub



Private Sub Form_Unload(Cancel As Integer)

If App.title = "Tidal Information..." Then Exit Sub

frmPassLock.Hide
Call zzSave_Checks

App.title = "Tidal Information..."
'Me.Top = 100
'set CIDRunME on off Top

'got jump
CidRunMeOff

LockDontAllowHotKeyforPassword = False

If HittMan = False Or HittMan = 2 Then
    'HittMan = False
    'LockSSaver = Now + LockSaverDelay1
    'If HittMan2 = False Then Call ATidalX.SetLockMax
End If

If App.title <> "Tidal Information..." Then
    Keleidoscope.HighRes.Enabled = False
    Bezier.Timer1.Enabled = False
    YinYang.Timer1.Enabled = False
    Vector.Timer1.Enabled = False
    Vector.Timer2.Enabled = False

    Unload Bezier
    Unload Keleidoscope
    Unload YinYang
    Unload Vector
End If

ATidalX.Timer10.Enabled = False
ExitVector = False

YinVectKeli$ = ""

Test7 = 0

App.title = "Tidal Information..."


If HittMan2 = False Then
Call ATidalX.SetLockTime
Call ATidalX.SetLockMax
End If
LockSSaver = Now + LockSaverDelay1
Call ATidalX.SetLockMax


App.title = "Tidal Information..."
ATidalX.PassLockTimer.Enabled = True
End Sub

Sub CidRunMe()
Exit Sub
'WMI Demo - CPU Information

xxr = FindWindow(vbNullString, "WMI Demo - CPU Information")
rt = AlwaysOnTop(xxr)
xxr = FindWindow(vbNullString, "CID Run Me.")
rt = AlwaysOnTop(xxr)
CIDonTop = True


End Sub

Sub CidRunMeOff()
Exit Sub
If CIDonTop = False Then Exit Sub
CIDonTop = False
xxr = FindWindow(vbNullString, "WMI Demo - CPU Information")
rt = NotAlwaysOnTop(xxr)
ShowWindow xxr, SW_NORMAL
ShowWindow xxr, SW_RESTORE
xxr = FindWindow(vbNullString, "CID Run Me.")
rt = NotAlwaysOnTop(xxr)

End Sub



Sub zzLoadChecks()

dxx = 0

If Combo1.ListCount = 0 Then
Combo1.AddItem "Bezier"
Combo1.AddItem "Keleidoscope"
Combo1.AddItem "Yin Yang"
Combo1.AddItem "Vector Pattern"
End If

On Error Resume Next
If YinVectKeli$ = "" Then
    dxx = 1
    Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--\ChkSettings-YinVectKeli.txt" For Input As #1
        Line Input #1, QQ$
    Close #1
    If QQ$ <> "" Then
        Combo1.Enabled = False
        Combo1.ListIndex = Val(QQ$)
        Combo1.Enabled = True
    End If
End If

YinVectKeli$ = Combo1.List(Combo1.ListIndex)

'MsgBox "Load " + YinVectKeli$

i = 0
Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--\ChkSettings-" + YinVectKeli$ + ".txt" For Input As #1
Do
    Line Input #1, vv$
    Th(i) = vv$
    i = i + 1
    Line Input #1, vv$
    Th(i) = Val(vv$)
    i = i + 1
Loop Until EOF(1)
Close #1
    
tit = i
    
For Each Control In Me.Controls
    XX = 0
    If InStr(Control.Name, "Combo") Then XX = 1
'    If InStr(Control.name, "Check") Then xx = 1
    
    xxd = 0: gogo = 0
    For R = 0 To tit
        If Control.Name = Th(R) Then
            xxd = R: gogo = 1: Exit For
        End If
    Next
    
    If gogo = 1 And XX = 0 Then
        'If Th(xxd + 1) = 0 Then yy = vbUnchecked Else yy = vbChecked
'        T = Control.name
        On Error Resume Next
        Control.Value = Th(xxd + 1)
        On Error GoTo 0
    
    End If
Next

On Error GoTo 0

'Call Timer2_Timer

'Call Combo1_Click
'frmPassLock.hide

Call StartTheWhirl

End Sub

Sub zzSave_Checks()
'If frmPassLock.Hide Then Exit Sub
If App.title = "Tidal Information..." Then Exit Sub

'Shell "notepad " + App.Path + "\00_Text_Data\ChkSettings.txt"

'YinVectKeli$ = Combo1.List(Combo1.ListIndex)
'MsgBox "Save " + YinVectKeli$
'Me.Visible = False

If YinVectKeli$ = "" Then Exit Sub
'Me.Visible = True

Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--\ChkSettings-YinVectKeli.txt" For Output As #1
    'Print #1, OldCombo1ListIndex
    Print #1, Combo1.ListIndex
    'Print #1, YinVectKeli$
Close #1

On Error Resume Next
Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "--\ChkSettings-" + YinVectKeli$ + ".txt" For Output As #1
For Each Control In Me.Controls
    Err.Clear
    A = Control.Value
    If Err.Number = 0 Then
        Print #1, Control.Name
        Print #1, Control.Value
    End If
Next
Close #1


Exit Sub

'--------------Sweet New Code

End Sub


