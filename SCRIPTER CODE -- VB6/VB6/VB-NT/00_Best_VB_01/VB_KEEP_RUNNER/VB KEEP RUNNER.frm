VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   6936
   ClientLeft      =   192
   ClientTop       =   1140
   ClientWidth     =   12744
   Icon            =   "VB KEEP RUNNER.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6936
   ScaleWidth      =   12744
   Begin VB.Timer Timer_GET_KEY_ASYNC_STATE 
      Interval        =   5
      Left            =   9180
      Top             =   3996
   End
   Begin VB.Timer Timer_FOREGROUND_WINDOW_CHANGE_02 
      Interval        =   10
      Left            =   8832
      Top             =   4008
   End
   Begin VB.Timer Timer_EnumProcess 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8472
      Top             =   4008
   End
   Begin VB.ListBox lstProcess_3_SORTER 
      Height          =   276
      IntegralHeight  =   0   'False
      Left            =   5520
      Sorted          =   -1  'True
      TabIndex        =   78
      Top             =   3492
      Visible         =   0   'False
      Width           =   1728
   End
   Begin VB.ListBox lstProcess_2_ 
      Height          =   324
      IntegralHeight  =   0   'False
      Left            =   5520
      TabIndex        =   77
      Top             =   3144
      Visible         =   0   'False
      Width           =   1128
   End
   Begin VB.TextBox txtMhWnd 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2676
      TabIndex        =   74
      Top             =   6264
      Width           =   1320
   End
   Begin VB.PictureBox picCrossHair 
      BorderStyle     =   0  'None
      Height          =   384
      Left            =   5004
      Picture         =   "VB KEEP RUNNER.frx":0E42
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   59
      Top             =   1668
      Width           =   384
   End
   Begin VB.TextBox txthWnd 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1284
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   5052
      Width           =   1044
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1284
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   2436
      Width           =   2004
   End
   Begin VB.TextBox txtClass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1284
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   2700
      Width           =   2004
   End
   Begin VB.TextBox txtParent 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1284
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   3492
      Width           =   1044
   End
   Begin VB.TextBox txtStyle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1284
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   2964
      Width           =   2004
   End
   Begin VB.TextBox txtParentText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1284
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   3756
      Width           =   2004
   End
   Begin VB.TextBox txtParentClass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1284
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   4020
      Width           =   2004
   End
   Begin VB.TextBox txtRect 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1284
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   3228
      Width           =   2004
   End
   Begin VB.TextBox txtParentClassX 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1284
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   4800
      Width           =   2004
   End
   Begin VB.TextBox txtParentTextX 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1284
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   4548
      Width           =   2004
   End
   Begin VB.TextBox txtParentX 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1284
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   4284
      Width           =   2004
   End
   Begin VB.TextBox TxtEXE 
      BorderStyle     =   0  'None
      Height          =   228
      Left            =   3312
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   12
      Width           =   7488
   End
   Begin VB.Timer Timer_ALWAYS_ON_TOP_TO_START_WITH_ER 
      Interval        =   1000
      Left            =   8124
      Top             =   4008
   End
   Begin VB.TextBox txthWndHX 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2352
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   5052
      Width           =   936
   End
   Begin VB.TextBox txtParentHX 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2352
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   3492
      Width           =   924
   End
   Begin VB.TextBox TxtPID 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1284
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   2172
      Width           =   1044
   End
   Begin VB.TextBox txtAPhysicalMemory 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   14580
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5352
      Visible         =   0   'False
      Width           =   2688
   End
   Begin VB.TextBox txtPhysicalMemory 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   14580
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   5100
      Visible         =   0   'False
      Width           =   2688
   End
   Begin VB.TextBox txtAVirtualMemory 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   14580
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   5856
      Visible         =   0   'False
      Width           =   2688
   End
   Begin VB.TextBox txtVirtualMemory 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   14580
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5604
      Visible         =   0   'False
      Width           =   2688
   End
   Begin VB.TextBox txtPageFile 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   14580
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   6108
      Visible         =   0   'False
      Width           =   2688
   End
   Begin VB.TextBox txtAPageFile 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   14580
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   6360
      Visible         =   0   'False
      Width           =   2688
   End
   Begin VB.CheckBox chkOnTop 
      Appearance      =   0  'Flat
      Caption         =   "ME Always On Top _ Registry Setting Form Load Remember"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   36
      TabIndex        =   7
      ToolTipText     =   "STORE IN REISTRY FOR START UP OVERRIDE TIMER IF DO"
      Top             =   5640
      Width           =   5300
   End
   Begin VB.Timer Timer_1_SECOND 
      Interval        =   1000
      Left            =   5700
      Top             =   4008
   End
   Begin VB.FileListBox File3 
      Height          =   264
      Left            =   6336
      TabIndex        =   6
      Top             =   6696
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.FileListBox File2 
      Height          =   264
      Left            =   6336
      TabIndex        =   5
      Top             =   6420
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Timer Timer_DIR_FOR_C_DRIVE_ROOT_VICE_VERSA_VBSCRIPT 
      Interval        =   1000
      Left            =   7776
      Top             =   4008
   End
   Begin VB.Timer Timer_VB_PROJECT_CHECKDATE 
      Interval        =   1000
      Left            =   7428
      Top             =   4008
   End
   Begin VB.FileListBox File1 
      Height          =   264
      Left            =   6348
      TabIndex        =   4
      Top             =   6144
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.DirListBox Dir1 
      Height          =   1368
      Left            =   6348
      TabIndex        =   3
      Top             =   4752
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer TIMER_DO_IN_1_SECOND 
      Interval        =   1000
      Left            =   7080
      Top             =   4008
   End
   Begin VB.Timer ONE_MILLISECOND_Timer 
      Interval        =   1
      Left            =   6048
      Top             =   4008
   End
   Begin VB.ListBox lstProcess 
      Height          =   240
      Left            =   5508
      TabIndex        =   2
      Top             =   1740
      Visible         =   0   'False
      Width           =   1896
   End
   Begin MSComctlLib.ListView lstProcess_2_ListView 
      Height          =   576
      Left            =   5508
      TabIndex        =   0
      Top             =   1116
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   995
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Timer Timer_VB_MAXIMIZE 
      Interval        =   10
      Left            =   6420
      Top             =   4008
   End
   Begin VB.Timer Timer_FOREGROUND_WINDOW_CHANGE 
      Interval        =   1
      Left            =   6744
      Top             =   4008
   End
   Begin MSComctlLib.ListView lstProcess_3_SORTER_ListView 
      Height          =   576
      Left            =   8040
      TabIndex        =   1
      Top             =   1116
      Width           =   2544
      _ExtentX        =   4509
      _ExtentY        =   995
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label42 
      BackColor       =   &H00C0FFFF&
      Caption         =   "KILL AUTOHOTKEY.EXE"
      Height          =   240
      Left            =   3480
      TabIndex        =   105
      Top             =   270
      Width           =   1920
   End
   Begin VB.Label Label64 
      Caption         =   "40 Sec"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   4428
      TabIndex        =   104
      Top             =   3528
      Width           =   960
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TASKKILLER /F /IM * /T"
      Height          =   240
      Left            =   11832
      TabIndex        =   103
      Top             =   516
      Width           =   1764
   End
   Begin VB.Label Label17 
      Caption         =   "TASKKILLER /IM * /T"
      Height          =   240
      Left            =   11832
      TabIndex        =   102
      Top             =   792
      Width           =   1764
   End
   Begin VB.Label Label18 
      Caption         =   "TASKKILLER /IM *"
      Height          =   240
      Left            =   11832
      TabIndex        =   101
      Top             =   1332
      Width           =   1764
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFC0FF&
      Caption         =   "TASKKILLER /F /IM /T"
      Height          =   240
      Left            =   13620
      TabIndex        =   100
      Top             =   516
      Width           =   1848
   End
   Begin VB.Label Label20 
      Caption         =   "TASKKILLER /IM /T"
      Height          =   240
      Left            =   13620
      TabIndex        =   99
      Top             =   792
      Width           =   1860
   End
   Begin VB.Label Label21 
      Caption         =   "TASKKILLER /IM"
      Height          =   240
      Left            =   13620
      TabIndex        =   98
      Top             =   1068
      Width           =   1860
   End
   Begin VB.Label Label15 
      Caption         =   "TASKKILLER COMMAND LINE GENERATED "
      Height          =   240
      Left            =   36
      TabIndex        =   97
      ToolTipText     =   "PRESS HERE TO KILL BY PROCESS NAME"
      Top             =   792
      Width           =   5364
   End
   Begin VB.Label Label23 
      Caption         =   "HITT TO CONFIRM SELECTION KILL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   36
      TabIndex        =   96
      Top             =   1296
      Width           =   5364
   End
   Begin VB.Label Label25 
      Caption         =   "TASKLIST GO"
      Height          =   240
      Left            =   11832
      TabIndex        =   95
      Top             =   1920
      Width           =   5400
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Caption         =   "TASKKILLER /F /IM"
      Height          =   240
      Left            =   13620
      TabIndex        =   94
      Top             =   1332
      Width           =   1860
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0FFC0&
      Caption         =   "TASKKILLER /F /IM"
      Height          =   240
      Left            =   15504
      TabIndex        =   93
      Top             =   1068
      Width           =   1740
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFC0FF&
      Caption         =   "TASKKILLER /F /IM /T"
      Height          =   240
      Left            =   15504
      TabIndex        =   92
      Top             =   792
      Width           =   1740
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TASKKILLER /F /IM * /T"
      Height          =   240
      Left            =   15492
      TabIndex        =   91
      Top             =   516
      Width           =   1740
   End
   Begin VB.Label Label13 
      Caption         =   "TASKKILLER BY PROCESS NUMBER"
      Height          =   240
      Left            =   36
      TabIndex        =   90
      ToolTipText     =   "PRESS HERE TO KILL BY PROCESS NUMBER"
      Top             =   540
      Width           =   2880
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "KILLER With /F FORCE"
      Height          =   240
      Left            =   15492
      TabIndex        =   89
      Top             =   192
      Width           =   1740
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Without WildCard Butter"
      Height          =   240
      Left            =   13620
      TabIndex        =   88
      Top             =   192
      Width           =   1848
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "WILD CARD ****"
      Height          =   240
      Left            =   11832
      TabIndex        =   87
      Top             =   192
      Width           =   1764
   End
   Begin VB.Label Label34 
      Caption         =   "TASKKILLER BY HWND HANDLE _ POST MESSENGER"
      Height          =   240
      Left            =   11832
      TabIndex        =   86
      Top             =   2172
      Width           =   5400
   End
   Begin VB.Label Label43 
      Caption         =   "TASKKILLER BY HWND HANDLE _ PROCESS KILL FORCE"
      Height          =   240
      Left            =   11832
      TabIndex        =   85
      Top             =   2424
      Width           =   5400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B7FFB5&
      BorderWidth     =   5
      X1              =   11856
      X2              =   17196
      Y1              =   468
      Y2              =   468
   End
   Begin VB.Label Label56 
      Alignment       =   2  'Center
      Caption         =   "TASKKILLER PID_NUMERIC OPTION_INSTEAD"
      Height          =   240
      Left            =   13632
      TabIndex        =   84
      Top             =   1656
      Width           =   3600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   5
      X1              =   11856
      X2              =   17196
      Y1              =   1608
      Y2              =   1608
   End
   Begin VB.Label Label57 
      BackColor       =   &H00C0FFFF&
      Caption         =   "TASKKILLER COMMAND LINE EXECUTE STATUS"
      Height          =   240
      Left            =   36
      TabIndex        =   83
      Top             =   1044
      Width           =   5364
   End
   Begin VB.Label Label55 
      Alignment       =   2  'Center
      Caption         =   "TASKKILLER PID NAME"
      Height          =   240
      Left            =   2940
      TabIndex        =   82
      Top             =   540
      Width           =   2460
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00B7FFB5&
      BorderWidth     =   5
      X1              =   11856
      X2              =   17196
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00B7FFB5&
      BorderWidth     =   5
      X1              =   11856
      X2              =   17196
      Y1              =   3276
      Y2              =   3276
   End
   Begin VB.Label Label62 
      Caption         =   "TASKKILLER /F /IM *"
      Height          =   240
      Left            =   11832
      TabIndex        =   81
      Top             =   1068
      Width           =   1764
   End
   Begin VB.Label Label63 
      BackColor       =   &H00C0FFFF&
      Caption         =   "KILL WSCRIPT.EXE"
      Height          =   240
      Left            =   1875
      TabIndex        =   80
      Top             =   270
      Width           =   1575
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TASKKILLER /F /IM * /T"
      Height          =   240
      Left            =   36
      TabIndex        =   79
      Top             =   264
      Width           =   1740
   End
   Begin VB.Label lblCordi 
      Caption         =   "X: 1043  Y: 0032"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   36
      TabIndex        =   76
      Top             =   5928
      Width           =   5304
   End
   Begin VB.Label Label3 
      Caption         =   "Manage Window _ hWnd:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   36
      TabIndex        =   75
      Top             =   6276
      Width           =   2616
   End
   Begin VB.Label Label60 
      Caption         =   "Me on Top Entry Timer 20 Second Yes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   36
      TabIndex        =   73
      Top             =   5340
      Width           =   5304
   End
   Begin VB.Label Label_FORM_BACK_COLOUR 
      BackColor       =   &H00005959&
      Caption         =   "Label_FORM BACK_Form Ground COLOUR"
      Height          =   396
      Left            =   5532
      TabIndex        =   72
      Top             =   2652
      Visible         =   0   'False
      Width           =   1896
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0C0&
      Caption         =   "C1"
      Height          =   240
      Left            =   5580
      TabIndex        =   71
      Top             =   2196
      Visible         =   0   'False
      Width           =   216
   End
   Begin VB.Label Label45 
      BackColor       =   &H00FFC0FF&
      Caption         =   "C2"
      Height          =   240
      Left            =   5820
      TabIndex        =   70
      Top             =   2196
      Visible         =   0   'False
      Width           =   216
   End
   Begin VB.Label Label49 
      BackColor       =   &H00C0FFC0&
      Caption         =   "C3"
      Height          =   240
      Left            =   6072
      TabIndex        =   69
      Top             =   2196
      Visible         =   0   'False
      Width           =   216
   End
   Begin VB.Label Label47 
      BackColor       =   &H0080FFFF&
      Caption         =   "C4"
      Height          =   240
      Left            =   6300
      TabIndex        =   68
      Top             =   2196
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label58 
      Caption         =   "C5"
      Height          =   240
      Left            =   6564
      TabIndex        =   67
      Top             =   2196
      Visible         =   0   'False
      Width           =   228
   End
   Begin VB.Label Label59 
      BackColor       =   &H00C0FFFF&
      Caption         =   "C7"
      Height          =   240
      Left            =   6816
      TabIndex        =   66
      Top             =   2196
      Visible         =   0   'False
      Width           =   216
   End
   Begin VB.Label Label61 
      BackColor       =   &H0068F9CA&
      Caption         =   "C8"
      Height          =   240
      Left            =   7056
      TabIndex        =   65
      Top             =   2196
      Visible         =   0   'False
      Width           =   216
   End
   Begin VB.Label Label46 
      Caption         =   "Process Counter Is:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5520
      TabIndex        =   63
      Top             =   300
      Width           =   2436
   End
   Begin VB.Label Label50 
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7992
      TabIndex        =   62
      Top             =   300
      Width           =   3612
   End
   Begin VB.Label Label51 
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   16.8
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   8628
      TabIndex        =   61
      Top             =   1944
      Visible         =   0   'False
      Width           =   204
   End
   Begin VB.Label Label53 
      Caption         =   "Process Set Enumerate Event"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   5520
      TabIndex        =   60
      ToolTipText     =   "Pause Update for 20 Second"
      Top             =   720
      Width           =   6072
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   4548
      Picture         =   "VB KEEP RUNNER.frx":170C
      Top             =   1656
      Width           =   384
   End
   Begin VB.Label lblHwnd 
      Caption         =   "hWnd:"
      Height          =   240
      Left            =   36
      TabIndex        =   58
      Top             =   5052
      Width           =   1224
   End
   Begin VB.Label lblTitle 
      Caption         =   "Title:"
      Height          =   240
      Left            =   36
      TabIndex        =   57
      Top             =   2172
      Width           =   1224
   End
   Begin VB.Label Label1 
      Caption         =   "1.. Drag this icon over the window you want to spy/nMouse Move to Top and Form Will Shirnk_er Until Done to Make Room Under"
      Height          =   1224
      Left            =   3396
      TabIndex        =   56
      Top             =   2268
      Width           =   1992
   End
   Begin VB.Label lblClass 
      Caption         =   "Class:"
      Height          =   240
      Left            =   36
      TabIndex        =   55
      Top             =   2700
      Width           =   1224
   End
   Begin VB.Label lblParent 
      Caption         =   "Parent:"
      Height          =   240
      Left            =   36
      TabIndex        =   54
      Top             =   3492
      Width           =   1224
   End
   Begin VB.Label Label2 
      Caption         =   "Style:"
      Height          =   240
      Left            =   36
      TabIndex        =   53
      Top             =   2964
      Width           =   1224
   End
   Begin VB.Label lblParentText 
      Caption         =   "Parent Text:"
      Height          =   240
      Left            =   36
      TabIndex        =   52
      Top             =   3756
      Width           =   1224
   End
   Begin VB.Label Label5 
      Caption         =   "Parent Class:"
      Height          =   240
      Left            =   36
      TabIndex        =   51
      Top             =   4020
      Width           =   1224
   End
   Begin VB.Label Label6 
      Caption         =   "Rectangle:"
      Height          =   240
      Left            =   36
      TabIndex        =   50
      Top             =   3228
      Width           =   1224
   End
   Begin VB.Label Label7 
      Caption         =   "Parent X Class:"
      Height          =   240
      Left            =   36
      TabIndex        =   49
      Top             =   4800
      Width           =   1224
   End
   Begin VB.Label Label8 
      Caption         =   "Parent X Text:"
      Height          =   240
      Left            =   36
      TabIndex        =   48
      Top             =   4548
      Width           =   1224
   End
   Begin VB.Label Label9 
      Caption         =   "Parent X:"
      Height          =   240
      Left            =   36
      TabIndex        =   47
      Top             =   4284
      Width           =   1224
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Goto File Name"
      Height          =   228
      Left            =   2124
      TabIndex        =   46
      Top             =   12
      Width           =   1164
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Put Whole File on Clipboard"
      Height          =   228
      Left            =   12
      TabIndex        =   45
      Top             =   12
      Width           =   2088
   End
   Begin VB.Label Label24 
      Caption         =   "2.. Use 20 or 40 Second Timer and Hover Land On/nover the window you want to spy"
      Height          =   888
      Left            =   3384
      TabIndex        =   44
      Top             =   3900
      Width           =   1992
   End
   Begin VB.Label Label48 
      Caption         =   "20 Sec"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   3384
      TabIndex        =   43
      Top             =   3528
      Width           =   960
   End
   Begin VB.Label cmdCopy 
      Alignment       =   2  'Center
      Caption         =   "Copy All Content to Clipboard"
      Height          =   468
      Left            =   3384
      TabIndex        =   42
      Top             =   4812
      Width           =   1992
   End
   Begin VB.Label Label65 
      Caption         =   "Process ID:"
      Height          =   240
      Left            =   36
      TabIndex        =   41
      Top             =   2436
      Width           =   1224
   End
   Begin VB.Label Label66 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Kill PID"
      Height          =   240
      Left            =   2364
      TabIndex        =   40
      Top             =   2172
      Width           =   912
   End
   Begin VB.Label Label30 
      Caption         =   "TASKKILLER COMMAND LINE GENERATED "
      Height          =   240
      Left            =   48
      TabIndex        =   24
      ToolTipText     =   "PRESS HERE TO KILL BY PROCESS NAME"
      Top             =   1872
      Width           =   4404
   End
   Begin VB.Label Label22 
      Caption         =   "TASKKILLER BY PROCESS NUMBER"
      Height          =   240
      Left            =   60
      TabIndex        =   23
      ToolTipText     =   "PRESS HERE TO KILL BY PROCESS NUMBER"
      Top             =   1620
      Width           =   4392
   End
   Begin VB.Label Label41 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Memory Usage 1 Second Ticker"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   228
      Left            =   11868
      TabIndex        =   22
      Top             =   4596
      Visible         =   0   'False
      Width           =   3312
   End
   Begin VB.Label txtMemoryUsage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "txtMemoryUser"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   228
      Left            =   15192
      TabIndex        =   21
      Top             =   4596
      Visible         =   0   'False
      Width           =   2076
   End
   Begin VB.Label Label4 
      Caption         =   "Total Physical Memory My Computer MAXIMUM 32G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   11868
      TabIndex        =   20
      Top             =   4848
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.Label Label40 
      Caption         =   "Avaible Physical Memory:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   11868
      TabIndex        =   19
      Top             =   5352
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Label39 
      Caption         =   "Total Physical Memory _ K:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   11868
      TabIndex        =   18
      Top             =   5100
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Label38 
      Caption         =   "Avaible Virtual Memory:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   11868
      TabIndex        =   17
      Top             =   5856
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Label37 
      Caption         =   "Total Virtual Memory _ K:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   11868
      TabIndex        =   16
      Top             =   5604
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Label36 
      Caption         =   "Total Page File _ K:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   11868
      TabIndex        =   15
      Top             =   6108
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Label35 
      Caption         =   "Avaible Page File:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   11868
      TabIndex        =   14
      Top             =   6360
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Label52 
      Caption         =   "Label52"
      Height          =   348
      Left            =   8304
      TabIndex        =   64
      Top             =   1920
      Visible         =   0   'False
      Width           =   1056
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_ME_ON_TOP 
      Caption         =   "ME ON TOP"
   End
   Begin VB.Menu MNU_WINMERGE_ON_TOP_ALLTME 
      Caption         =   "WINMERGE_ON_TOP_ALLTME=YES"
   End
   Begin VB.Menu MNU_AUTOHOTKEYS_SET 
      Caption         =   "Run AutoHotKey Set"
   End
   Begin VB.Menu MNU_VERSION 
      Caption         =   "MNU_VERSION"
   End
   Begin VB.Menu MNU_EXE_01 
      Caption         =   "EXE_MNU_01"
      Begin VB.Menu MNU_VB_SYNCRONIZER 
         Caption         =   "VB SYNCRONIZER"
      End
      Begin VB.Menu MNU_LAUNCH_BATCH_COMPILER 
         Caption         =   "VB BATCH COMPILER"
      End
      Begin VB.Menu MNU_LAUNCH_Shell_VBasic_6_Loader 
         Caption         =   "VB Shell_VBasic_6_Loader"
      End
      Begin VB.Menu MNU_VB_LAUNCH_FAV_SET 
         Caption         =   "VB LAUNCH FAV SET"
      End
      Begin VB.Menu MNU_AUTOHOTKEY_STARTING 
         Caption         =   "AUTOHOTKEY STARTER"
      End
      Begin VB.Menu MNU_PIN_ITEM_BATCH_VBS 
         Caption         =   "VBS 12-PinItem BATCH.VBS"
      End
   End
   Begin VB.Menu MNU_GOODSYNC_COLLECTION_SCRIPT_RUN 
      Caption         =   "GOODSYNC COLLECTION SCRIPT RUN"
   End
   Begin VB.Menu MNU_MAXIMIZE_GOODSYNC 
      Caption         =   "MAXIMIZE GOODSYNC"
   End
   Begin VB.Menu MNU_CLOSE_GOODSYNC 
      Caption         =   "CLOSE GOODSYNC"
   End
   Begin VB.Menu MNU_GIVER_ME_UPTIME 
      Caption         =   "GIVE ME UPTIME"
   End
   Begin VB.Menu MNU_OS_RESTART 
      Caption         =   "OS RESTART"
   End
   Begin VB.Menu MNU_KILL_NOT_RESPOND_TOP 
      Caption         =   "KILL NOT RESPOND"
   End
   Begin VB.Menu MNU_TASK_KILLER_NOT_RESPONDER_FORCE 
      Caption         =   "KILL NOT RESPOND FORCE"
   End
   Begin VB.Menu MNU_TASK_KILLER_AUTOHOTKEYS 
      Caption         =   "KILL AUTOHOTKEY"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------
' WROTE SCRIPT FOR C DRIVE ROOT VICE VERSA WORKING VBSCRIPT GRUB4DOS
' CHECKS THE FOLDER FOR ANY CHANGE
' AND RUN THE SCRIPT
'-------------------------------------------------------------------
' FROM   ---- Tue 01-May-2018 14:02:31
' TO     ---- Tue 01-May-2018 14:32:54 30 MINUTE __ HALF AN HOUR
'-------------------------------------------------------------------

' SORTED THE UPDATER AND EXIT UNLOAD FROM MAIN FORM
' HEAD-ACHE TOOK LONG
' THIS UPDATER CODE IS IN THESE ONE
' 1. VB KEEP RUNNER __ HERE
' 2. ELITESPY
' 3. CLIPBOARD LOGGER
' 4. TIDAL
' 5. CPU %
' POSSIBLE SOME OTHER THERE
'-------------------------------------------------------------------
' FROM   ---- Mon 28-May-2018 17:26:15
' FROM   ---- Mon 28-May-2018 21:11:00 4 HOUR MINUS 15 MINUTE
'-------------------------------------------------------------------

' ------------------------------------------------------------------
' VARIABL DECLARE BLOCK FROM ELITEPSY
Public EXIT_TRUE

Option Explicit

Dim picCrossHair_MouseMove_Dragging_VAR

Dim O_Ret_Connected_To_The_Internet

'Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal Index As Long) As Long
'Private Const SM_CXSCREEN = 0
'Private Const SM_CYSCREEN = 1
'MsgBox GetSystemMetrics(SM_CXSCREEN) & "x" & GetSystemMetrics(SM_CYSCREEN)

Dim ESCAPE_ERROR_WINGDINGS_2, ESCAPE_ERROR_WINGDINGS_3
Dim TYPEMESSENGER

Dim WTrue_1, WTrue_2, WTrue_3, TW1, TW2, TW3
Dim FIRST_RUN_COLOUR_CYCLE
Dim Index_String_Camera


Dim IsIDE_TEST

Dim OHWnd_VB_EXE
Dim OHWnd_VB_CLIPPER_ERROR
Dim OHWnd_VB_LOADER
Dim OHWnd_TEAM_VIEWER
Dim O_mWnd_VB_VbaWindow_MAXIMIZE

Dim OcWnd
Dim IHWND, O_IHWND

Dim TO_SETTER

Dim Label51_Height

Dim O_GetForegroundWindow
Dim O_GetForegroundWindow_02
Dim O_STRING_COMPARE
Dim NOT_RESIZE_EVENTER_SET_HAPPEN_ME_XW
Dim NOT_RESIZE_EVENTER_SET_HAPPEN_ME_YH

Dim counter_ALWAYS_ON_TOP_TIMER
Dim TOP_HEIGHT

Dim FIRST_RUN_FOR_TOP_AND_LEFT
Dim NOT_RESIZE_EVENTER
Dim HY2, WX2, OFFSET_HY, OFFSET_WX


Dim OLD_i_VIRTUA_COP
Dim O_Me_WindowState
Dim RUN_ONCE_DEBUG_PRINT_HY_RESIZE
Dim RUN_ONCE_DEBUG_PRINT_HY
Dim TT, TTHY1, TTHY2, TTHY3
Dim RESIZE_LOOP_STOP
Dim RUN_ONCE_VB_DIRECTORY

Dim O_lstProcess_ListCount
Dim OHWnd_FINDER_1
Dim OHWnd_FINDER_2
Dim Result As Long
Dim OcWnd_GOOD_SYNC_CRASH, OcWnd_GOOD_SYNC
Dim I_Memmer, OcWnd_ICACLS_SETTER_PERMISSION_HWND


Dim Form As Form
Dim R_NEXT
Dim A_DATE_TIME As String
Dim A_DATE_TIME_PM

Dim COUNT_SHOW_THE_TIME

Dim PROGRAM_PATH_BAT
Dim DIR1_START_UP_DELAY_FOLDER

Dim FS

Dim XVB_DATE

'MATT.LAN@BTINTERNET.COM
'IMPROVEMENT
'DETAILS OF WINDOW IN TIMER AS IT CHAGES
'LACKING IN WINDOWSE ALSO
'D:\0 00 Art Loggers\# APP AND SCREEN\7-ASUS-GL522VW\FILE LOCATOR __ SavedCriteria.srf

Dim PROCESS_TO_KILLER_PID As String
Dim PROCESS_TO_KILLER As String
Public PROCESS_TO_KILLER_TO_GO

Dim EnumProcess_COUNTER

Dim DELAY_TICKER
Dim pid As Long, APP_NAME_EXE_PASS_FOR_CALL_BACK As String

Dim VAR_IN, VAR_OUT

Dim TIMER2_TIMER_BEGAN

Public MSGQ, IR, CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2

Dim WX, HY
Dim CMD_Process_STATE_TO_SET
Dim CMD_Process_STATE_TO_SET_1ST
Dim EXE_FILE_NAME
Dim COMMAND_EXE_MENU_NOTEPAD_PICK_UP_DATE As Date
Dim F1
Dim EXE_ITEM
Dim FILE_TEXT, FILE_EXE
Dim FR1
Dim NOTEPAD_PLUSPLUS_OR_NOTEPAD_NORMAL_RUN_WHEN_ERROR_LOAD_TIME_ONE_MSGBOX
Dim FILENAME_PATH_EXE_MENU


Dim hWnd_From_ListView
Dim From_ListView
Dim OLD_TxtEXE_Text
Dim lHwnd_Function_Button_Set_MIN_MAX
' ------------------------------------------------------------------
' VARIABL DECLARE BLOCK FROM ELITEPSY



'Dim PROCESS_TO_KILLER_PID As String
'Dim PROCESS_TO_KILLER As String
'Public PROCESS_TO_KILLER_TO_GO


' ----------------------------------------------------------
' MAJOR IMPROVEMENT AND GOT WORKING 1ST TIME RUNNING UPDATER
' THE UPDATED WHEN NEW COMPILED COMES ALONG EXIT AND RELAUNCH
' IN VBSCRIPT FROM A MIRROR FOLDER VB-EXE
' ----------------------------------------------------------
' Thu 03-May-2018 09:25:29
' Thu 03-May-2018 11:25:00 -- 3 HOUR
' ----------------------------------------------------------

Dim O_DAY_NOW_DIR_FOR_VIDEO_ME_YOU_TUBE
Dim O_DAY_NOW_DIR_FOR_VIDEO_KILLSOMETIME
Dim O_DAY_NOW_DIR_FOR_XXX_BUNKER_COM
Dim O_DAY_NOW_DIR_FOR_HARDWARE

Dim XVB_DATE_2
'Public EXIT_TRUE

Dim OV

Dim SIZE_SET

Dim FSO
Dim O_C_DRIVEROOT_V_VAR
Dim Err_COUNTER_Number

Dim O_Dir1_COUNTER_W_VAR
'Dim O_STRING_COMPARE
Dim FORM_LOAD_EnumProcess
'Dim O_GetForegroundWindow
'Dim O_mWnd_VB_VbaWindow_MAXIMIZE

Public ScreenTwipsX, ScreenTwipsY, ScreenWidthX, ScreenHeightY, Idle_Timer_Proc



Const E = 2.7182818284
Const pi = 3.141592648
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const shrdNoMRUString = &H2
Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400
Const LF_FACESIZE = 32
Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4
Const GRADIENT_FILL_RECT_H As Long = &H0
Const GRADIENT_FILL_RECT_V  As Long = &H1
Const GRADIENT_FILL_TRIANGLE As Long = &H2
Const WM_NCLBUTTONDOWN = &HA1
Const LP_HT_CAPTION = 2
Const GWL_WNDPROC = -4
Const SWP_HIDEWINDOW = &H80


Private Type POINTAPI
        X As Long
        Y As Long
End Type


Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWndParent As Long, _
     ByVal hWndChildAfter As Long, _
     ByVal lpszClass As String, _
     ByVal lpszTitle As String) _
    As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long  'MODULE 1141

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long

Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, _
    ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As _
    Long
Private Declare Function GetCaretBlinkTime Lib "user32" () As Long

Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long  'MODULE 1135
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction&, ByVal uParam&, ByRef lpvParam As Any, ByVal fuWinIni&) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const PROCESS_CREATE_PROCESS = &H80
Private Const PROCESS_CREATE_THREAD = &H2
Private Const PROCESS_DUP_HANDLE = &H40
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const PROCESS_SET_QUOTA = &H100
Private Const PROCESS_SET_INFORMATION = &H200
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_VM_OPERATION = &H8
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_VM_WRITE = &H20

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Private Type PROCESS_INFORMATION
  hProcess    As Long
  hThread     As Long
  dwProcessId As Long
  dwThreadId  As Long
End Type

Private Type STARTUPINFO
  cb              As Long
  lpReserved      As String
  lpDesktop       As String
  lpTitle         As String
  dwX             As Long
  dwY             As Long
  dwXSize         As Long
  dwYSize         As Long
  dwXCountChars   As Long
  dwYCountChars   As Long
  dwFillAttribute As Long
  dwFlags         As Long
  wShowWindow     As Integer
  cbReserved2     As Integer
  lpReserved2     As Long
  hStdInput       As Long
  hStdOutput      As Long
  hStdError       As Long
End Type

Private Enum Priorities
  p_RealTime = &H100
  p_Hight = &H80
  p_Normal = &H20
  p_Idle = &H40
End Enum

Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Private Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, hModule As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Any, ByVal lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpExitCode As Long) As Long
Private Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Boolean

Private Declare Function CloseHandle Lib "kernel32" _
        (ByVal hObject As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Const TH32CS_SNAPPROCESS = &H2&

Private Type MENUBARINFO
  cbSize As Long
  rcBar As RECT
  hMenu As Long
  hWndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type

Private MenuInfo As MENUBARINFO
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

' PRIVATE constants for ShowWindow API declaration
Private Const SW_HIDE = 0
Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_NORMAL = 1
Private Const SW_RESTORE = 9
Private Const SW_SHOW = 5

' PRIVATE constants for set window on top
'Private Const HWND_TOPMOST = -1
'Private Const HWND_NOTOPMOST = -2

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" _
      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Declare Function FindWindowDLL Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long

Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Const GW_HWNDNEXT = 2
Private Const WM_CLOSE = &H10


Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Type FILETIME
   LowDateTime          As Long
   HighDateTime         As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes     As Long
   ftCreationTime       As FILETIME
   ftLastAccessTime     As FILETIME
   ftLastWriteTime      As FILETIME
   nFileSizeHigh        As Long
   nFileSizeLow         As Long
   dwReserved0          As Long
   dwReserved1          As Long
   cFileName            As String * 260  'MUST be set to 260
   cAlternate           As String * 14
End Type

Private Declare Function MoveWindow _
        Lib "user32" _
        (ByVal hWnd As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal bRepaint As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function Putfocus _
        Lib "user32" _
        Alias "SetFocus" _
        (ByVal hWnd As Long) As Long

'Private Type FILETIME
'    dwLowDate  As Long
'    dwHighDate As Long
'End Type

' Private Type SYSTEMTIME
'     wYear      As Integer
'     wMonth     As Integer
'     wDayOfWeek As Integer
'     wDay       As Integer
'     wHour      As Integer
'     wMinute    As Integer
'     wSecond    As Integer
'     wMillisecs As Integer
' End Type

Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_WRITE = &H40000000
 
Private Declare Function CreateFile Lib "kernel32" Alias _
  "CreateFileA" (ByVal lpFileName As String, _
  ByVal dwDesiredAccess As Long, _
  ByVal dwShareMode As Long, _
  ByVal lpSecurityAttributes As Long, _
  ByVal dwCreationDisposition As Long, _
  ByVal dwFlagsAndAttributes As Long, _
  ByVal hTemplateFile As Long) _
  As Long

Private Declare Function LocalFileTimeToFileTime Lib _
    "kernel32" (lpLocalFileTime As FILETIME, _
     lpFileTime As FILETIME) As Long

Private Declare Function SetFileTime Lib "kernel32" _
  (ByVal hFile As Long, ByVal MullP As Long, _
   ByVal NullP2 As Long, lpLastWriteTime _
   As FILETIME) As Long

Private Declare Function SystemTimeToFileTime Lib _
   "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime _
   As FILETIME) As Long


'----
'VB6 CHANGE MODIFED DATE OF FOLDER - Google Search
'https://www.google.co.uk/search?q=VB6+CHANGE+MODIFED+DATE+OF+FOLDER&rlz=1C1CHBD_en-GBGB744GB744&oq=VB6+CHANGE+MODIFED+DATE+OF+FOLDER&aqs=chrome..69i57.14080j0j7&sourceid=chrome&ie=UTF-8
'--------
'FreeVBCode code snippet: Change a File's Last Modified Date Stamp
'http://www.freevbcode.com/ShowCode.asp?ID=1335
'----
'----
'Folder Creation Date
'http://forums.codeguru.com/showthread.php?24993-Folder-Creation-Date
'----
'Private Type FILETIME
'dwLowDateTime As Long
'dwHighDateTime As Long
'End Type
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

Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const GW_OWNER = 4
Private Const WS_EX_TOOLWINDOW = &H80
Private Const WS_EX_APPWINDOW = &H40000
Private Const GA_ROOT = 2

Private Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long

Private Const LVM_FIRST = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Private Const LVS_EX_FULLROWSELECT = &H20


Const BM_CLICK = &HF5&

Private Const WM_SETTEXT            As Long = &HC
Private Const WM_GETTEXT As Long = &HD
Private Const WM_GETTEXTLENGTH As Long = &HE

Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const INVALID_FILE_SIZE As Long = &HFFFFFFFF
Private Const INVALID_SET_FILE_POINTER As Long = -1

Private Const MAX_PROFILE_LEN As Long = 80

Private Const MAX_COMPUTERNAME_LENGTH As Long = 31

Private Const MAXLONG As Long = &H7FFFFFFF


'Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1


Private Const DELETE As Long = &H10000
Private Const READ_CONTROL As Long = &H20000
Private Const WRITE_DAC As Long = &H40000
Private Const WRITE_OWNER As Long = &H80000
Private Const SYNCHRONIZE As Long = &H100000
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const STANDARD_RIGHTS_READ As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_ALL As Long = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL As Long = &HFFFF
Private Const GENERIC_READ As Long = &H80000000
'Private Const GENERIC_WRITE As Long = &H40000000
Private Const GENERIC_EXECUTE As Long = &H20000000
Private Const GENERIC_ALL As Long = &H10000000

Private Const MAX_PATH As Long = 260


'Private Const PROCESS_TERMINATE As Long = &H1
'Private Const PROCESS_CREATE_THREAD As Long = &H2
Private Const PROCESS_SET_SESSIONID As Long = &H4
'Private Const PROCESS_VM_OPERATION As Long = &H8
'Private Const PROCESS_VM_READ As Long = &H10
'Private Const PROCESS_VM_WRITE As Long = &H20
'Private Const PROCESS_DUP_HANDLE As Long = &H40
'Private Const PROCESS_CREATE_PROCESS As Long = &H80
'Private Const PROCESS_SET_QUOTA As Long = &H100
'Private Const PROCESS_SET_INFORMATION As Long = &H200
'Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)

Private Const DEBUG_PROCESS As Long = &H1
Private Const DEBUG_ONLY_THIS_PROCESS As Long = &H2
Private Const CREATE_SUSPENDED As Long = &H4
Private Const DETACHED_PROCESS As Long = &H8
Private Const CREATE_NEW_CONSOLE As Long = &H10
Private Const NORMAL_PRIORITY_CLASS As Long = &H20
Private Const IDLE_PRIORITY_CLASS As Long = &H40
Private Const HIGH_PRIORITY_CLASS As Long = &H80
Private Const REALTIME_PRIORITY_CLASS As Long = &H100
Private Const CREATE_NEW_PROCESS_GROUP As Long = &H200
Private Const CREATE_UNICODE_ENVIRONMENT As Long = &H400
Private Const CREATE_SEPARATE_WOW_VDM As Long = &H800
Private Const CREATE_SHARED_WOW_VDM As Long = &H1000
Private Const CREATE_FORCEDOS As Long = &H2000
Private Const BELOW_NORMAL_PRIORITY_CLASS As Long = &H4000
Private Const ABOVE_NORMAL_PRIORITY_CLASS As Long = &H8000
Private Const CREATE_BREAKAWAY_FROM_JOB As Long = &H1000000

Private Const SE_PRIVILEGE_ENABLED_BY_DEFAULT As Long = &H1
Private Const SE_PRIVILEGE_ENABLED As Long = &H2
Private Const SE_PRIVILEGE_USED_FOR_ACCESS As Long = &H80000000

Private Const THREAD_TERMINATE As Long = &H1
Private Const THREAD_SUSPEND_RESUME As Long = &H2
Private Const THREAD_GET_CONTEXT As Long = &H8
Private Const THREAD_SET_CONTEXT As Long = &H10
Private Const THREAD_SET_INFORMATION As Long = &H20
Private Const THREAD_QUERY_INFORMATION As Long = &H40
Private Const THREAD_SET_THREAD_TOKEN As Long = &H80
Private Const THREAD_IMPERSONATE As Long = &H100
Private Const THREAD_DIRECT_IMPERSONATION As Long = &H200
Private Const THREAD_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF)

Private Const THREAD_BASE_PRIORITY_LOWRT As Long = 15
Private Const THREAD_BASE_PRIORITY_MAX As Long = 2
Private Const THREAD_BASE_PRIORITY_MIN As Long = -2
Private Const THREAD_BASE_PRIORITY_IDLE As Long = -15


'Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

'Function GetUserName() As String
'   Dim UserName As String * 255
'   Call GetUserNameA(UserName, 255)
'   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
'End Function
'
'Function GetComputerName() As String
'   Dim UserName As String * 255
'   Call GetComputerNameA(UserName, 255)
'   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
'End Function


' Public API declarations
'Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
'Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
'Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
'Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
'Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
'Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
'Public Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
'Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)



Public Function DirExist(sPath As String) As Boolean
    DirExist = (Dir(sPath, vbDirectory) <> "")
End Function

Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Function
Function NotAlwaysOnTop(ByVal hWnd As Long)
    Dim Flags
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags
End Function







Public Function GetShortName(ByVal sLongFileName As String) As String
' --> (All this modules code) Obtained from : -
' ---> Microsoft's/MSDN's code
' ->  http://support.microsoft.com/kb/q175512/
' ---> The original comments were by them :

       Dim lRetVal As Long, sShortPathName As String, iLen As Integer
       'Set up buffer area for API function call return
       sShortPathName = Space(255)
       iLen = Len(sShortPathName)

       'Call the function
       lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
       'Strip away unwanted characters.
       GetShortName = Left(sShortPathName, lRetVal)
End Function


Public Function GetLongName(ByVal sShortName As String) As String

' --> (All this modules code) Obtained from : -
' ---> Microsoft's/MSDN's code
' ->  http://support.microsoft.com/default.aspx?scid=kb;EN-US;154822
' ---> The original comments were by them :

     Dim sLongName As String
     Dim sTemp As String
     Dim iSlashPos As Integer

     'Add \ to short name to prevent Instr from failing
     
     sShortName = sShortName & "\"

     'Start from 4 to ignore the "[Drive Letter]:\" characters
     iSlashPos = InStr(4, sShortName, "\")

     'Pull out each string between \ character for conversion
     While iSlashPos
       sTemp = Dir(Left$(sShortName, iSlashPos - 1), vbNormal + vbHidden + vbSystem + vbDirectory)
       If sTemp = "" Then
         'Error 52 - Bad File Name or Number
         GetLongName = ""
         Exit Function
       End If
       sLongName = sLongName & "\" & sTemp
       iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
     Wend

     'Prefix with the drive letter
     GetLongName = Left$(sShortName, 2) & sLongName

End Function




'----------------------------------------------------
'I PUT THIS HERE FOR THE SHELL EXECUTE LOADER PROGRAM
'----------------------------------------------------
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
'----------------------------------------------------
'----------------------------------------------------


Private Sub Form_Activate()
    
    'Call Form3_CHROME_X_BUTTON_OFF.SetXState(Me.hWnd, False)
    'Call Form3_CHROME_X_BUTTON_OFF.SET_CHROME_WINDOW_X_BUTTON_CLOSE_TO_OFF
    
'    Call Form4_DISABLE_CLOSE_BUTTON.SET_BUTTON_CHROME

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If IsIDE = True Then
    If KeyCode = (27) Then End
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If IsIDE = True Then
    If KeyAscii = (27) Then End
End If
End Sub


Private Sub Form_Load()

'    HWND_GET = FindWindow("Sym_Common_Scan_Window", vbNullString)
'    MsgBox HWND_GET
'    End

'    YY = Now - DateDiff("d", DateSerial(2012, 1, 1), DateSerial(2013, 7, 1)) - DateDiff("d", DateSerial(2012, 1, 1), DateSerial(2012, 12, 31))
'    Debug.Print YY
'
'    Stop
'
'    YY = DateDiff("n", DateSerial(2018, 1, 1), DateSerial(2019, 1, 1))
'
'    YY = YY / 32000
'    XX = (425 / 100) * 60
'
'    Debug.Print XX
'    Stop

    If App.PrevInstance = True Then End
        

    
    Call Form2_Check_Project_Date.VB_PROJECT_CHECKDATE("FORM LOAD")
    
    Set FSO = CreateObject("Scripting.FileSystemObject")

    With lstProcess_2_ListView
        .ColumnHeaders.Add , "PID", "PID", 700 - 50, lvwColumnLeft
        .ColumnHeaders.Add , "EXE", "EXE", 9000, lvwColumnLeft
        .View = lvwReport
        .height = 7000
        .FullRowSelect = True
    End With
    With lstProcess_3_SORTER_ListView
        .ColumnHeaders.Add , "PID", "PID", 700 - 50, lvwColumnLeft
        .ColumnHeaders.Add , "EXE SORTED", "EXE SORTED", 9000, lvwColumnLeft
        .View = lvwReport
        .height = lstProcess_2_ListView.height
        .FullRowSelect = True
    End With

    Me.Visible = False
    Me.Caption = App.EXEName
    FORM_LOAD_EnumProcess = True

    MNU_VERSION.Caption = "Ver_2018_" + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision)) ' + " _ Matt.Lan@btinternet.com"

    If Form2_Check_Project_Date.GetWindowsVersion = 5.1 Then
        MNU_PIN_ITEM_BATCH_VBS.Visible = False
    End If

End Sub

Private Sub Form_Resize()
    
    Dim Me_Top, HwndTask, HwndResult
    
    If SIZE_SET = True Then Exit Sub
    SIZE_SET = True
    'Taskbar at the Bottom
    '---------------------
    Dim RectTaskbar As RECT
    HwndTask = FindWindow("Shell_TrayWnd", vbNullString)
    HwndResult = GetWindowRect(HwndTask, RectTaskbar)
    ' CENTER ON SCREEN COMPENSATE TASKBAR AT BOTTOM
    '----------------------------------------------
    If Screen.height - RectTaskbar.Top * Screen.TwipsPerPixelY < 2500 Then
        Me.Left = Screen.width / 2 - Me.width / 2
        Me_Top = (Screen.height - ((RectTaskbar.Bottom - RectTaskbar.Top) * Screen.TwipsPerPixelY)) / 2 - Me.height / 2
    Else
        Me.Left = Screen.width / 2 - Me.width / 2
        Me_Top = Screen.height / 2 - Me.height / 2
    End If
    If Me_Top < 0 Then Me_Top = 0
    Me.Top = Me_Top

    Me.width = lstProcess_3_SORTER_ListView.Left + lstProcess_3_SORTER_ListView.width + 300

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
For Each Form In Forms
Err.Clear
    If Form.EXIT_TRUE = True Then
        'Form.Name
        If Err.Number = 0 Then
            Me.EXIT_TRUE = True
        End If
    End If
Next Form
On Error GoTo 0

If IsIDE = False Then
    If Me.WindowState <> vbMinimized And Me.EXIT_TRUE = False Then
        Me.WindowState = vbMinimized
        Cancel = True
        Exit Sub
    End If
End If

If Me.EXIT_TRUE = True Then Cancel = False

For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next Form

'End

End Sub

Private Sub Label29_Click()
'PROCESS_TO_KILLER PID /F *
PROCESS_TO_KILLER_TO_GO = "/F /IM """ + Replace(UCase(PROCESS_TO_KILLER), ".EXE", "") + "*"" /T"

'PROCESS BUTTON_ER
'-----------------
'01 OF 10

Label22 = "TASKKILL " + "/F /IM """ + PROCESS_TO_KILLER_PID + """"
Label30 = "TASKKILL " + "/F /IM """ + Replace(UCase(PROCESS_TO_KILLER), ".EXE", "") + "*"" /T"
Label57.Caption = "COMMAND LINE STATUS__ " + Label30

'Label22 = "TASKKILL " + PROCESS_TO_KILLER_PID
'Label30 = "TASKKILL " + PROCESS_TO_KILLER_TO_GO

'Label57 = "TASKKILL " + PROCESS_TO_KILLER_TO_GO


'Label47.BackColor = GO SELECTED ACTIVE YELLOW
'Label23__________ = CONFIRM BOX

'Label22.BackColor = STARDARD GREY OUT WHEN NOT USER
'Label22.BackColor = STATUS BOX READY TO GO INFO

'Label49.BackColor = READY TO GO SELECTED BIG BOX AND NORMAL NOT ACTIVE
'ALSO FOR LB26 LB27 = PASTEL GREEN FROM LABEL11 BEFORE

'Label16 DUPE WITH Label29.BackColor = Label12.BackColor = BLUE COLOUR
'Label19 DUPE WITH Label28.BackColor = Label45.BackColor = BLUE RED
'Label26 DUPE WITH Label27.BackColor = Label49.BackColor = BLUE GREEN

'------------------------------------
'CONFIRM BOX
Label23.BackColor = Label49.BackColor
'------------------------------------

Label16.BackColor = Label12.BackColor
Label29.BackColor = Label12.BackColor

Label17.BackColor = Label22.BackColor
Label18.BackColor = Label22.BackColor
Label20.BackColor = Label22.BackColor
Label21.BackColor = Label22.BackColor

Label19.BackColor = Label45.BackColor
Label28.BackColor = Label45.BackColor

Label26.BackColor = Label49.BackColor
Label27.BackColor = Label49.BackColor
'------------------------------------

'ACTIVE TO GO
'------------------------------------
Label16.BackColor = Label47.BackColor
Label29.BackColor = Label47.BackColor

End Sub


Private Sub Label42_Click()

'PROCESS_TO_KILLER = lstProcess_3_SORTER_ListView.ListItems(lstProcess_3_SORTER_ListView.SelectedItem.Index).SubItems(1)
PROCESS_TO_KILLER = "AUTOHOTKEY.EXE"
'PROCESS_TO_KILLER_PID = lstProcess_3_SORTER_ListView.ListItems(lstProcess_3_SORTER_ListView.SelectedItem.Index)

'Label22.Caption = "TASKKILLER BY PID NUMBER __ # " + PROCESS_TO_KILLER_PID
Label30.Caption = "TASKKILLER NAME ___________ " + PROCESS_TO_KILLER


'PROCESS_TO_KILLER PID
Label29_Click

'PROCESS_TO_KILLER PID
'Label22_Click

'PROCESS_TO_KILLER /F * /T
Label57.Caption = "COMMAND LINE STATUS__ " + "TASKKILL /F /IM """ + Replace(UCase(PROCESS_TO_KILLER), ".EXE", "") + "*"" /T"

Label23_Click


End Sub

Private Sub Label23_Click()
Dim PROGRAM_PATH_BAT

If PROCESS_TO_KILLER_TO_GO = "" Then
    Beep
    Exit Sub
End If

Beep
'Me.WindowState = vbMinimized

'PROGRAM_PATH_BAT = "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT_03_PROCESS_KILLER.BAT"
PROGRAM_PATH_BAT = App.Path + "\BAT_03_PROCESS_KILLER.BAT"

'-----------------------------------------------
'----
'HOW DO PARAM IN START COMMAND LINE WITH QUOTE - Google Search
'https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB721GB721&q=HOW+DO+PARAM+IN+START+COMMAND+LINE+WITH+QUOTE&oq=HOW+DO+PARAM+IN+START+COMMAND+LINE+WITH+QUOTE&gs_l=serp.12..0i71k1l8.0.0.0.262638.0.0.0.0.0.0.0.0..0.0....0...1c..64.serp..0.0.0.z74guYcKqJc
'----
'----
'windows - Using the "start" command with parameters passed to the started program - Stack Overflow
'http://stackoverflow.com/questions/154075/using-the-start-command-with-parameters-passed-to-the-started-program
'----
'-----------------------------------------------
'MUST USE START COMMAND WITH SHORT-NAME FOR PATH
'BECAUSE ENLCOSE QUOTE WITH PARAM
'-----------------------------------------------

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'WELL ANSWER ACCUTALLY IS
'THE BATCH FILE HAS THE FORMAT %1 %2 %3 %4
'EACH ARE SEPERATE AND NOTHING ABLE DO ABOUT THAT
'LEAVE THE WORK TO BE DONE IN THE BATCH FILE ABOUT DOUBLE QUOTING
'IT IS GOOD AT PASS THE PARAMETER QUOTE IN PARAM SO OKAY ABOUT THAT
'OR ELSE HARD WORK AND THE %1 %2 %3 %4 ARE ABLE TO DO WITH FOR LOOP NESTER
'OR ELSE THE PARAM WILL ONLY GO UPTO 9 %9 __ %10 IS %1 BACK AGAIN
'-----------------------------------------------------------------------------------
'WITH START AND CMD AND BATCH PASS TO
'-----------------------------------------------------------------------------------
'AND 2ND THEN PATH TO BATCH COMMAND _NAME_.BAT DOES HAVE TO GET REDUCED TO SHORTNAME
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'Debug.Print PROGRAM_PATH_BAT
'PROGRAM_PATH_BAT = GetShortName(PROGRAM_PATH_BAT)

'Debug.Print PROCESS_TO_KILLER_TO_GO

'PROCESS_TO_KILLER_TO_GO = "/IM ""Kat MP3 Recorder.exe"""

PROCESS_TO_KILLER_TO_GO = Mid(Label57, InStr(Label57, "TASKKILL") + 9)

If Dir(PROGRAM_PATH_BAT) = "" Then

    MsgBox PROGRAM_PATH_BAT + vbCrLf + vbCrLf + "DON'T EXIST _ WILL CREATE ONE"
    
    Call CREATE_PROGRAM_PATH_BAT

End If



Shell "CMD /C START """" /REALTIME /MIN " + PROGRAM_PATH_BAT + " " + PROCESS_TO_KILLER_TO_GO, vbMinimizedNoFocus
' --------------------------------------------------------------------------
' Shell PROGRAM_PATH_BAT + " " + PROCESS_TO_KILLER_TO_GO, vbMinimizedNoFocus
' --------------------------------------------------------------------------
' THIS ABOVE IS A DIFFERCULT LINE AS THE BATCH IS STORE IN APP.PATH AND THAT MUST BE WITHOUT SPACE
' BEEN RENAME FROM 02 MAR 2019
' WAS TAKEN PASTE FROM ELITESPY
' IT OKAY IF USE LINE IN REM BUT START AND CMD TOGETHER MAKE HARD
' NOT MUCH INFO ON-LINE NONE FOUND FOR THIS WHOLE SUBJECT CMD WITH START
' --------------------------------------------------------------------------
' ----
' Problem with quotes around file names in Windows command shell - Super User
' https://superuser.com/questions/238810/problem-with-quotes-around-file-names-in-windows-command-shell
' ----
' HOW PUT QUOTE COMMANDLINE WITH START - Google Search
' https://www.google.com/search?q=HOW+PUT+QUOTE+COMMANDLINE+WITH+START
' ----
' shell - How do I use quotes in cmd start? - Stack Overflow
' https://stackoverflow.com/questions/27261692/how-do-i-use-quotes-in-cmd-start
' ----
' HOW PUT QUOTE COMMANDLINE WITH START - Google Search
' https://www.google.com/search?q=HOW+PUT+QUOTE+COMMANDLINE+WITH+START
' ----
' [ Saturday 01:52:50 Am_02 March 2019 ]
' --------------------------------------------------------------------------


Me.WindowState = vbMinimized

Timer_EnumProcess = True
Timer_EnumProcess.Interval = 1000
Label23.BackColor = Label22.BackColor
Label23.ForeColor = QBColor(0)

PROCESS_TO_KILLER_TO_GO = ""

End Sub

Sub CREATE_PROGRAM_PATH_BAT()

Dim i

i = ""
i = i + "@CD\" + vbCrLf
i = i + "@ECHO ------------------------------------------------------" + vbCrLf
i = i + "@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION" + vbCrLf
i = i + "@TITLE %~f0" + vbCrLf
i = i + "@IF NOT ""%1%""=="""" (" + vbCrLf
i = i + "" + vbCrLf
i = i + "    Rem ECHO %1 %2 %3 %4 %5 %6 %7 %8 %9" + vbCrLf
i = i + "    Rem PAUSE" + vbCrLf
i = i + "" + vbCrLf
i = i + "    @ECHO OFF" + vbCrLf
i = i + "    @ECHO -----------------------------" + vbCrLf
i = i + "    Rem FOR %%A IN (%*) DO (" + vbCrLf
i = i + "        Rem ECHO %%A" + vbCrLf
i = i + "    Rem )" + vbCrLf
i = i + "    Rem @ECHO -----------------------------" + vbCrLf
i = i + "" + vbCrLf
i = i + "    SET VAR_SET_TO_GO=%1 %2 %3 %4 %5 %6 %7 %8 %9" + vbCrLf
i = i + "    ECHO !VAR_SET_TO_GO!" + vbCrLf
i = i + "" + vbCrLf
i = i + "" + vbCrLf
i = i + "    @TITLE %~f0 __ %1 %2 %3 %4 %5 %6 %7 %8 %9 __" + vbCrLf
i = i + "    @ECHO -----------------------------" + vbCrLf
i = i + "    @ECHO PROCESS PROGRAM EXE TO KILLER" + vbCrLf
i = i + "    @ECHO -----------------------------" + vbCrLf
i = i + "    )" + vbCrLf
i = i + "@ECHO." + vbCrLf
i = i + "@ECHO OFF" + vbCrLf
i = i + "" + vbCrLf
i = i + "" + vbCrLf
i = i + "" + vbCrLf
i = i + "Rem ---- VAR SET TO GO NOT HERE AND THEN ASKER" + vbCrLf
i = i + "Rem ------------------------------------------" + vbCrLf
i = i + "" + vbCrLf
i = i + "IF %1.==. (" + vbCrLf
i = i + "" + vbCrLf
i = i + "    @ECHO ENTER THE KILL REQUIREMENT ADD A WILDCARD FOR YOUR OWN __ EXAMPLE __ TASKKILL /F /IM INPUTNAME /T" + vbCrLf
i = i + "    @SET INPUT=" + vbCrLf
i = i + "    @SET /P INPUT=Type input: %=%" + vbCrLf
i = i + "    ECHO." + vbCrLf
i = i + "    ECHO --------------------------------------------------" + vbCrLf
i = i + "    ECHO RUN COMMAND LINE ____ /F /IM !INPUT! /T" + vbCrLf
i = i + "    ECHO." + vbCrLf
i = i + "    TASKKILL /F /IM !INPUT! /T" + vbCrLf
i = i + ")" + vbCrLf
i = i + "" + vbCrLf
i = i + "Rem ---- VAR SET TO GO HAS A VALUE AND USER" + vbCrLf
i = i + "Rem ---- TAKEN WILD-CARD OUT TO USE THE VB CODE DO THAT THING" + vbCrLf
i = i + "Rem ---- INSTEAD WHEN PASS IT OVER ---- COMMAND LINE" + vbCrLf
i = i + "Rem ---------------------------------------" + vbCrLf
i = i + "IF NOT %1.==. (" + vbCrLf
i = i + "    @REM ------------------------------------------------------------------" + vbCrLf
i = i + "    @REM THE SOME @ECHO ON HERE IS IGNORE WITHIN THE CONDITION BLOCK" + vbCrLf
i = i + "    @REM AND THEN ADDITIONAL ECHO BY COMMAND ECHO TO SHOW EXECUTION COMMAND" + vbCrLf
i = i + "    @REM REDUNDANT COMMAND LEFT IN" + vbCrLf
i = i + "    @REM OUR OWN VB CODE WILL ADD /F /IM _*_ AND /T IF SELECTED" + vbCrLf
i = i + "    @REM WOULD LIKE FOR ABOVE ALO AS THIS A COMMAND INTENDING TO USE" + vbCrLf
i = i + "    @REM OR VB SCRIPTOR" + vbCrLf
i = i + "    @REM ------------------------------------------------------------------" + vbCrLf
i = i + "    @ECHO RUN HERE" + vbCrLf
i = i + "    @ECHO -----------------------------" + vbCrLf
i = i + "    @ECHO ON" + vbCrLf
i = i + "    @ECHO TASKKILL %VAR_SET_TO_GO%" + vbCrLf
i = i + "    @ECHO." + vbCrLf
i = i + "    TASKKILL %VAR_SET_TO_GO%" + vbCrLf
i = i + "    @ECHO OFF" + vbCrLf
i = i + ")" + vbCrLf
i = i + "" + vbCrLf
i = i + "@REM __ IF EXPLORER ASK TO BE KILL WAIT FEW SECOND" + vbCrLf
i = i + "@REM __ CHECKER AND RESUME" + vbCrLf
i = i + "" + vbCrLf
i = i + "@ECHO OFF" + vbCrLf
i = i + "" + vbCrLf
i = i + "ECHO %1 %2 %3 %4 %5 %6 %7 %8 %9 | findstr /I ""EXPLORER.EXE""" + vbCrLf
i = i + "if %errorlevel% == 0 SET EXPLORER_METHOD_TO_GO__=YES" + vbCrLf
i = i + "" + vbCrLf
i = i + "IF NOT ""%EXPLORER_METHOD_TO_GO__%""==""YES"" (" + vbCrLf
i = i + "@ECHO." + vbCrLf
i = i + "@ECHO --------------------------------" + vbCrLf
i = i + "@ECHO EXPLORER USE KILLER NOT DETECTOR" + vbCrLf
i = i + ")" + vbCrLf
i = i + "" + vbCrLf
i = i + "IF ""%EXPLORER_METHOD_TO_GO__%""==""YES"" (" + vbCrLf
i = i + "    @ECHO." + vbCrLf
i = i + "    @ECHO EXPLORER DETECTOR WAIT FEW MOMENT...." + vbCrLf
i = i + "    @ECHO." + vbCrLf
i = i + "" + vbCrLf
i = i + "    @SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION" + vbCrLf
i = i + "    @FOR /F %%# IN ('COPY /Z ""%~dpf0"" NUL') DO SET ""CR=%%#""" + vbCrLf
i = i + "    @FOR /L %%# IN (4,-1,1) DO (SET/P ""=Script will end in %%# seconds. !CR!""<NUL:" + vbCrLf
i = i + "        PING -n 2 127.0.0.1 >NUL:)" + vbCrLf
i = i + "    @ECHO." + vbCrLf
i = i + "" + vbCrLf
i = i + "    @REM ------------------------------------------" + vbCrLf
i = i + "    @REM ____ IF NOT ANY EXPLORER ARE RUNNING AFTER" + vbCrLf
i = i + "    @TASKLIST /FI ""IMAGENAME EQ EXPLORER.EXE""  | find ""INFO: No tasks are running"" > nul" + vbCrLf
i = i + "    @IF %errorlevel% EQU 0 GOTO YES" + vbCrLf
i = i + "    @GOTO NOT" + vbCrLf
i = i + ": YES" + vbCrLf
i = i + "    @ECHO." + vbCrLf
i = i + "    TASKLIST /FI ""IMAGENAME EQ EXPLORER.EXE""" + vbCrLf
i = i + "    @ECHO." + vbCrLf
i = i + "    @REM -----------------------------------" + vbCrLf
i = i + "    cmd /c EXPLORER" + vbCrLf
i = i + "    @REM -----------------------------------" + vbCrLf
i = i + "    @ECHO." + vbCrLf
i = i + "    TASKLIST /FI ""IMAGENAME EQ EXPLORER.EXE""" + vbCrLf
i = i + "    @ECHO." + vbCrLf
i = i + "    @ECHO ---- EXPLORER.EXE RESET AND RELOAD HAPPEN SUCCESSFULLY" + vbCrLf
i = i + "    :NOT" + vbCrLf
i = i + "    @ECHO OFF" + vbCrLf
i = i + "    @REM ------------------------------------------" + vbCrLf
i = i + ")" + vbCrLf
i = i + "@ECHO." + vbCrLf
i = i + "" + vbCrLf
i = i + "@REM AS NORMAL" + vbCrLf
i = i + "" + vbCrLf
i = i + "@ECHO OFF" + vbCrLf
i = i + "@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION" + vbCrLf
i = i + "@FOR /F %%# IN ('COPY /Z ""%~dpf0"" NUL') DO SET ""CR=%%#""" + vbCrLf
i = i + "@FOR /L %%# IN (20,-1,1) DO (SET/P ""=Script will end in %%# seconds. !CR!""<NUL:" + vbCrLf
i = i + "    PING -n 2 127.0.0.1 >NUL:)" + vbCrLf
i = i + "@ECHO." + vbCrLf
i = i + "" + vbCrLf
i = i + "Rem PAUSE" + vbCrLf
i = i + "Rem EXIT" + vbCrLf



PROGRAM_PATH_BAT = App.Path + "\BAT_03_PROCESS_KILLER.BAT"
FR1 = FreeFile
Open PROGRAM_PATH_BAT For Output As #FR1
    Print #FR1, i;
Close #FR1

End Sub


Private Sub Timer_EnumProcess_Timer()
'EnumProcess_COUNTER
'SUB ENUMPROCESS
Call EnumProcess
Timer_EnumProcess.Interval = 60000

'Label15.Caption = "Scan Processor For A Moment Same As Other Command Set Timer Second _ " + Trim(Str(EnumProcess_COUNTER))
'Label15.Caption = "Scan Processor For A Moment Timer Second _ " + Trim(Str(EnumProcess_COUNTER))
Label15.Caption = "Scan Processor For A Moment Timer Second _ " + Trim(Str(EnumProcess_COUNTER)) + " && Foregound Window" ' Change"

'Label23.BackColor = Label22.BackColor

If EnumProcess_COUNTER = 0 Then
    'Timer_EnumProcess = False
    Timer_EnumProcess.Interval = 60000
    EnumProcess_COUNTER = 10
End If

EnumProcess_COUNTER = EnumProcess_COUNTER - 1
'Label23.BackColor = Label22.BackColor

End Sub


'Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, _
'    ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
'Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As _
'    Long
'Declare Function GetCaretBlinkTime Lib "user32" () As Long

Sub ShowCustomCaret(ByVal width As Integer, ByVal height As Integer)
    On Error Resume Next
    With Screen.ActiveForm.ActiveControl
        CreateCaret .hWnd, 0, width, height
        ShowCaret .hWnd
    End With

'    mWnd_VB_VbaWindow_MAXIMIZE

End Sub

Private Sub Timer_FOREGROUND_WINDOW_CHANGE_02_Timer()

If IsIDE = True And IsIDE_TEST = True Then Timer_FOREGROUND_WINDOW_CHANGE_02.Interval = 10000

If GetForegroundWindow = O_GetForegroundWindow_02 Then
    Exit Sub
End If
O_GetForegroundWindow_02 = GetForegroundWindow


Call Timer_VB_MAXIMIZE_Timer

Dim I_HWnd
I_HWnd = FindWindow("MediaPlayerClassicW", vbNullString)
If I_HWnd > 0 Then
    SetWindowPos I_HWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End If

I_HWnd = FindWindow("Winamp v1.x", vbNullString)
If I_HWnd > 0 Then
    SetWindowPos I_HWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End If

End Sub




Private Sub Label63_Click()

PROCESS_TO_KILLER = lstProcess_3_SORTER_ListView.ListItems(lstProcess_3_SORTER_ListView.SelectedItem.Index).SubItems(1)
PROCESS_TO_KILLER_PID = lstProcess_3_SORTER_ListView.ListItems(lstProcess_3_SORTER_ListView.SelectedItem.Index)

Label22.Caption = "TASKKILLER BY PID NUMBER __ # " + PROCESS_TO_KILLER_PID
Label30.Caption = "TASKKILLER NAME ___________ " + PROCESS_TO_KILLER

PROCESS_TO_KILLER = "WSCRIPT.EXE"

'PROCESS_TO_KILLER PID
Label29_Click

'PROCESS_TO_KILLER PID
'Label22_Click

'PROCESS_TO_KILLER /F * /T
Label57.Caption = "COMMAND LINE STATUS__ " + "TASKKILL /F /IM """ + Replace(UCase(PROCESS_TO_KILLER), ".EXE", "") + "*"" /T"

Label23_Click


End Sub

Private Sub MNU_AUTOHOTKEY_STARTING_Click()
Dim objShell
Set objShell = CreateObject("Wscript.Shell")

objShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 21-AUTORUN.ahk""", 0, True
    
Set objShell = Nothing
End Sub

Private Sub MNU_AUTOHOTKEYS_SET_Click()

Me.WindowState = vbMinimized
Beep

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk" + """"
Set WSHShell = Nothing


End Sub

Private Sub MNU_CLOSE_GOODSYNC_Click()

Dim HWND_RESULT
HWND_RESULT = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F009}", vbNullString)
If HWND_RESULT > 0 Then
    Result = PostMessage(HWND_RESULT, WM_CLOSE, 0&, 0&)
End If

Me.WindowState = vbMinimized
Beep

End Sub

Private Sub MNU_GIVER_ME_UPTIME_Click()

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 58-GET SYSTEM UPTIME.BAT" + """"
Set WSHShell = Nothing

End Sub

Private Sub MNU_MAXIMIZE_GOODSYNC_Click()

Dim GOODSYNC_WINDOW_HWND
GOODSYNC_WINDOW_HWND = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F009}", vbNullString)

ShowWindow GOODSYNC_WINDOW_HWND, SW_MAXIMIZE
Beep
Me.WindowState = vbMinimized

End Sub

Private Sub MNU_OS_RESTART_Click()

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "C:\WINDOWS\system32\shutdown.exe" + """ -r", 0, True
Set WSHShell = Nothing
Beep
Me.WindowState = vbMinimized
End Sub

Private Sub MNU_PIN_ITEM_BATCH_VBS_Click()

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 12-PinItem BATCH.VBS" + """"
Set WSHShell = Nothing
Beep
Me.WindowState = vbMinimized

End Sub

Private Sub MNU_LAUNCH_BATCH_COMPILER_Click()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "D:\VB6\VB-NT\00_Best_VB_01\Batch_Compiler_Auto\BatchCompiler.exe" + """"
Set WSHShell = Nothing
Beep
Me.WindowState = vbMinimized
End Sub
Private Sub MNU_LAUNCH_Shell_VBasic_6_Loader_Click()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "D:\VB6\VB-NT\00_Best_VB_01\Shell VBasic 6 Loader\Shell VBasic 6 Loader.exe" + """"
Set WSHShell = Nothing
Beep
Me.WindowState = vbMinimized
End Sub


Private Sub MNU_VB_FOLDER_Click()
Shell "EXPLORER /SELECT, " + App.Path + "\" + App.EXEName + ".VBP", vbMaximizedFocus
Beep
Me.WindowState = vbMinimized
End Sub

Private Sub MNU_VB_LAUNCH_FAV_SET_Click()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "C:\SCRIPTER\SCRIPTER CODE -- AUTOKEY\Autokey -- 31-AUTORUN SET FAV VB & AUTOHOTKEY.ahk" + """"
Set WSHShell = Nothing
Beep
Me.WindowState = vbMinimized
End Sub

Private Sub MNU_GOODSYNC_COLLECTION_SCRIPT_RUN_Click()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 59-SCRIPT RUN GOOD OLE GOODSYNC SCRIPTOR FOLDER.BAT" + """"
Set WSHShell = Nothing

Beep
Me.WindowState = vbMinimized

End Sub

Private Sub MNU_VB_ME_Click()
    
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

    
    Beep
    On Error Resume Next
    Dim IGETTER
    IGETTER = GetSpecialfolder(38)
    
    If IGETTER = "" Then
        MsgBox "ERROR LOAD ADMIN"
        MsgBox "YOU HAVE TO SWITCH ADMIN MODE ON" + vbCrLf + vbCrLf + "THE COMMAND GetSpecialfolder(38) REQUIRE ADMINISTRATOR" + vbCrLf + vbCrLf + "RESULT RETURN NOTHING HERE GetSpecialfolder(38)__<" + GetSpecialfolder(38) + ">__ RESULT RETURNED NOTHING", vbMsgBoxSetForeground
        'MsgBox "YOU HAVE TO SWITCH ADMIN MODE ON" + vbCrLf + vbCrLf + "THE COMMAND GetSpecialfolder(38) REQUIRE ADMINISTRATOR" + vbCrLf + vbCrLf + "RESULT RETURN NOTHING HERE GetSpecialfolder(38)__< GetSpecialfolder(38) >__ RESULT RETURNED NOTHING", vbMsgBoxSetForeground
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
    
'    MsgBox """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + CODER_VBP_FILE_NAME_2 + """"
    
    Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + CODER_VBP_FILE_NAME_2 + """", vbNormalFocus
    
    Me.WindowState = vbMinimized 'to down and en-able exit
    
    EXIT_TRUE = True
    Unload Me
    'End

End Sub
'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    Dim FileSpec, FILE_SPEC_GO, i As String
    
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".vbp"
    i = CODER_VBP_FILE_NAME_2
    If InStr(i, " - 64 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(i, " - 64 BIT.vbp", ".vbp")
    If InStr(i, " - 32 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(i, " - 32 BIT.vbp", ".vbp")
    
    CODER_EXE_FILE_NAME_1 = Replace(CODER_EXE_FILE_NAME_1, ".vbp", ".exe")
    
    FileSpec = CODER_VBP_FILE_NAME_2
    If Dir(FileSpec) <> "" Then
        FILE_SPEC_GO = FileSpec
    Else
        If LCase(GetSpecialfolder(&H29)) = LCase("C:\Windows\SysWOW64") Then
            FileSpec = Replace(FileSpec, ".vbp", " - 64 BIT.vbp")
        Else
            FileSpec = Replace(FileSpec, ".vbp", " - 32 BIT.vbp")
        End If
    End If
    
    If Dir(FileSpec) <> "" Then FILE_SPEC_GO = FileSpec
    
    CODER_VBP_FILE_NAME_2 = FILE_SPEC_GO
End Sub

Public Function GetSpecialFolder_Show_Script_Debug(CSIDL As Long) As String

Dim R As Long
On Error Resume Next
For R = 0 To 120
    If Trim(GetSpecialfolder(R)) <> "" Then
        'Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
        'AAX = GetSpecialfolder(R)
    End If
Next
End
    
End Function

Public Function GetSpecialfolder(CSIDL As Long) As String
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



Private Sub MNU_VB_SYNCRONIZER_Click()
    Dim RUN_EXE
    Dim objShell
    Set objShell = CreateObject("Wscript.Shell")
    RUN_EXE = "D:\VB6\VB-NT\00_Best_VB_01\10 SYNCRONIZE\SYNCRONIZER.exe"
    objShell.Run """" + RUN_EXE + """", 1, False
    Set objShell = Nothing
End Sub

Private Sub ONE_MILLISECOND_Timer_Timer()

ONE_MILLISECOND_Timer.Enabled = False

Me.height = lstProcess_2_ListView.Top + lstProcess_2_ListView.height + (Menu_Height * Screen.TwipsPerPixelY) + 500 + 110

If IsIDE = True Then
    Me.WindowState = vbNormal
Else
    Me.WindowState = vbMinimized
End If

'Me.WindowState = vbMinimized
Me.Visible = True

End Sub


Private Sub lstProcess_2_ListView_KeyDown(KeyCode As Integer, Shift As Integer)
If IsIDE = True And KeyCode = (27) Then End
End Sub

Private Sub lstProcess_3_SORTER_ListView_BeforeLabelEdit(Cancel As Integer)
If IsIDE = True Then End
'If IsIDE = True And KeyCode = (27) Then End
End Sub

Private Function Menu_Height()
 
    MenuInfo.cbSize = Len(MenuInfo)
    
    If GetMenuBarInfo(Me.hWnd, OBJID_MENU, 0, MenuInfo) Then
   
        With MenuInfo.rcBar
       
'            Debug.Print "Left: " & CStr(.Left)
'            Debug.Print "Right: " & CStr(.Right)
'            Debug.Print "Top: " & CStr(.Top)
'            Debug.Print "Bottom: " & CStr(.Bottom)
            Menu_Height = CStr(.Bottom) - CStr(.Top)

        End With
       
    End If
   
End Function


Private Sub MNU_WINMERGE_ON_TOP_ALLTME_Click()
Beep
If MNU_WINMERGE_ON_TOP_ALLTME = "WINMERGE ON TOP ALLTIME=YES" Then
    MNU_WINMERGE_ON_TOP_ALLTME = "WINMERGE ON TOP ALLTIME=NOT"
Else
    MNU_WINMERGE_ON_TOP_ALLTME = "WINMERGE ON TOP ALLTIME=YES"
End If
End Sub

Private Sub MNU_EXIT_Click()
End
End Sub


Private Sub MNU_TASK_KILLER_NOT_RESPONDER_FORCE_Click()

'FORM1.MNU_BOOT_KILLER_Click
Beep
Me.WindowState = vbMinimized
'Shell "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-BOOT KILLER.BAT", vbMaximizedFocus
'Shell "CMD /C START """" /REALTIME /MAX ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-BOOT KILLER.BAT""", vbMaximizedFocus
Shell "CMD /C START """" /REALTIME /MAX ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-NOT RESPONDER KILLER FORCE.BAT""", vbMinimizedNoFocus
Beep

End Sub

Private Sub MNU_TASK_KILLER_NOT_RESPONDER_GENTAL_Click()

'FORM1.MNU_BOOT_KILLER_Click
Beep
Me.WindowState = vbMinimized
'Shell "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-BOOT KILLER.BAT", vbMaximizedFocus
'Shell "CMD /C START """" /REALTIME /MAX ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-BOOT KILLER.BAT""", vbMaximizedFocus
Shell "CMD /C START """" /REALTIME /MAX ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-NOT RESPONDER KILLER NOT FORCE.BAT""", vbMinimizedNoFocus
Beep
End Sub

Private Sub MNU_KILL_NOT_RESPOND_TOP_Click()
    Call MNU_TASK_KILLER_NOT_RESPONDER_GENTAL_Click
End Sub

Private Sub MNU_ME_ON_TOP_Click()
    
    Beep
    ' Put window on top of all others
    'SetWindowPos txtMhWnd.Text, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    Timer_ALWAYS_ON_TOP_TO_START_WITH_ER.Enabled = False
    Label60.BackColor = Label49.BackColor '49 58_59
    Label60.Caption = "Me on Top @ Loader EXE Timer 20 Second NOT _ DONE"
    MNU_ME_ON_TOP.Caption = "[__ ME_ON_TOP_=_YES __]"
    Me.chkOnTop.Value = 1
End Sub

Private Sub MNU_TASK_KILLER_AUTOHOTKEYS_Click()
    
PROCESS_TO_KILLER_TO_GO = "/F /IM AUTOHOTKEY* /T"
PROCESS_TO_KILLER = PROCESS_TO_KILLER_TO_GO

Me.WindowState = vbMinimized

PROGRAM_PATH_BAT = App.Path + "\BAT_03_PROCESS_KILLER.BAT"
If Dir(PROGRAM_PATH_BAT) = "" Then
    MsgBox PROGRAM_PATH_BAT + vbCrLf + vbCrLf + "DON'T EXIST _ WILL CREATE"
    Call CREATE_PROGRAM_PATH_BAT
End If

Shell "CMD /C START """" /REALTIME " + PROGRAM_PATH_BAT + " " + PROCESS_TO_KILLER, vbMaximizedFocus
' Shell "CMD /C START """" /REALTIME ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT_03_PROCESS_KILLER.BAT"" " + PROCESS_TO_KILLER, vbMaximizedFocus
Beep

End Sub


Private Sub Timer_DIR_FOR_VIDEO_ME_YOU_TUBE_Timer()

On Error Resume Next

Dim DIR_PATH

If O_DAY_NOW_DIR_FOR_VIDEO_ME_YOU_TUBE = Day(Now) Then Exit Sub

O_DAY_NOW_DIR_FOR_VIDEO_ME_YOU_TUBE = Day(Now)

DIR_PATH = "C:\DOWNLOADS\# 00 VIDEO\# VIDEO ME YOU_TUBE"
    
If Dir(DIR_PATH + "\" + Format(Now, "YYYY-MM-MMM"), vbDirectory) <> "" Then Exit Sub

MkDir DIR_PATH + "\" + UCase(Format(Now, "YYYY-MM-MMM"))

End Sub


Private Sub Timer_DIR_FOR_VIDEO_KILLSOMETIME_Timer()

On Error Resume Next

Dim DIR_PATH

If O_DAY_NOW_DIR_FOR_VIDEO_KILLSOMETIME = Day(Now) Then Exit Sub

O_DAY_NOW_DIR_FOR_VIDEO_KILLSOMETIME = Day(Now)

DIR_PATH = "C:\DOWNLOADS\# 00 VIDEO\# VIDEO KILLSOMETIME.COM"
    
If Dir(DIR_PATH + "\" + Format(Now, "YYYY-MM-MMM"), vbDirectory) <> "" Then Exit Sub

MkDir DIR_PATH + "\" + UCase(Format(Now, "YYYY-MM-MMM"))

End Sub




Private Sub Timer_DIR_FOR_XXX_BUNKER_COM_Timer()

On Error Resume Next

Dim DIR_PATH

If O_DAY_NOW_DIR_FOR_XXX_BUNKER_COM = Day(Now) Then Exit Sub

O_DAY_NOW_DIR_FOR_XXX_BUNKER_COM = Day(Now)

DIR_PATH = "D:\VIDEO\NOT\X 00 NOT ME\00 Vid XXX\BUNKER.COM"

'If FindWindowPart(DIR_PATH) = 0 Then
'    O_Dir1_COUNTER_W_VAR = ""
'    Exit Sub
'End If
 
    
If Dir(DIR_PATH + "\" + Format(Now, "YYYY-MM-MMM"), vbDirectory) <> "" Then Exit Sub

MkDir DIR_PATH + "\" + UCase(Format(Now, "YYYY-MM-MMM"))

End Sub


Private Sub Timer_1_SECOND_Timer()

Call Timer_DIR_FOR_XXX_BUNKER_COM_Timer
Call Timer_DIR_FOR_HARDWARE_Timer
Call Timer_DIR_FOR_VIDEO_KILLSOMETIME_Timer
Call Timer_DIR_FOR_VIDEO_ME_YOU_TUBE_Timer

End Sub

Private Sub Timer_DIR_FOR_HARDWARE_Timer()

On Error Resume Next

Dim DIR_PATH, ALLOW_GO, W_VAR, R
Dim SET_TRIGGER_1, SET_TRIGGER_2, R_Y

DIR_PATH = "E:\HARDWARE"

ALLOW_GO = False
If O_DAY_NOW_DIR_FOR_HARDWARE <> Day(Now) Then ALLOW_GO = True
O_DAY_NOW_DIR_FOR_HARDWARE = Day(Now)

If ALLOW_GO = False Then
    If FindWindowPart(DIR_PATH) = 0 Then
        O_Dir1_COUNTER_W_VAR = ""
        Exit Sub
    End If
End If

Dir1.Path = DIR_PATH

W_VAR = ""
Dir1.Refresh
For R = 0 To Dir1.ListCount
    W_VAR = W_VAR + Dir1.List(R)
    File1.Path = Dir1.List(R)
    If File1.ListCount = 0 Then W_VAR = Time$: Exit For
Next
If Err.Number > 0 Then Exit Sub

If O_Dir1_COUNTER_W_VAR = W_VAR Then Exit Sub

O_Dir1_COUNTER_W_VAR = W_VAR

'If File1.Pattern <> "*.*" Then File1.Pattern = "*.*"

If Err.Number > 0 Then O_Dir1_COUNTER_W_VAR = "": Exit Sub
SET_TRIGGER_1 = False
SET_TRIGGER_2 = False
For R = 0 To Dir1.ListCount
    File1.Path = Dir1.List(R)
    If File1.ListCount = 0 Then
        If Err.Number > 0 Then O_Dir1_COUNTER_W_VAR = "": Exit Sub
        FR1 = FreeFile
        Open File1.Path + "\" + Mid(Dir1.List(R), InStrRev(Dir1.List(R), "\") + 1) + "- " + Format(Now, "HH-MM-SS") + " - ITEM LEARN.txt" For Output As #FR1
            Print #FR1, Mid(Dir1.List(R), InStrRev(Dir1.List(R), "\") + 1) + " - " + Format(Now, "HH-MM-SS")
        Close #FR1
    End If
Next
    
If Dir(DIR_PATH + "\HARDWARE " + Format(Now, "YYYY-MM-DD -"), vbDirectory) = "" Then SET_TRIGGER_1 = True

For R_Y = 2004 To Year(Now)
    If Dir(DIR_PATH + "\HARDWARE " + Trim(Str(R_Y)) + "-12-40", vbDirectory) = "" Then
    
        SET_TRIGGER_2 = R_Y
        Exit For
    End If
Next

If SET_TRIGGER_1 = True Then
    MkDir DIR_PATH + "\HARDWARE " + Format(Now, "YYYY-MM-DD")
End If
If SET_TRIGGER_2 > 0 Then
    MkDir DIR_PATH + "\HARDWARE " + Trim(Str(R_Y)) + "-12-40"
    O_Dir1_COUNTER_W_VAR = ""
End If

End Sub


Private Sub Timer_DIR_FOR_C_DRIVE_ROOT_VICE_VERSA_VBSCRIPT_Timer()

If Timer_DIR_FOR_C_DRIVE_ROOT_VICE_VERSA_VBSCRIPT.Interval <> 10000 Then
    Timer_DIR_FOR_C_DRIVE_ROOT_VICE_VERSA_VBSCRIPT.Interval = 10000
End If

On Error Resume Next

Dim V_VAR, R, F

File2.Path = "C:\C DRIVE ROOT"
File2.Hidden = True
File2.Archive = True
File2.System = True

File3.Path = "C:\"
File3.Hidden = True
File3.Archive = True
File3.System = True


File2.Refresh
File3.Refresh
If File2.ListCount = 0 Then
    O_C_DRIVEROOT_V_VAR = ""
    Exit Sub
End If
 
V_VAR = ""
V_VAR = V_VAR + Trim(Str(File2.ListCount))
Err.Clear
'Err.Number
For R = 0 To File2.ListCount - 1
    
    Set F = FSO.GetFile(File2.Path + "\" + File2.List(R))
    V_VAR = V_VAR + File2.List(R)
    V_VAR = V_VAR + Format(F.DateLastModified, "YYYY-MM-DD HH-MM-SS")
    V_VAR = V_VAR + Trim(Str(F.Size))
    Set F = Nothing
    
    If FSO.FileExists("C:\" + File2.List(R)) Then
        Set F = FSO.GetFile("C:\" + File2.List(R))
        V_VAR = V_VAR + File2.List(R)
        V_VAR = V_VAR + Format(F.DateLastModified, "YYYY-MM-DD HH-MM-SS")
        V_VAR = V_VAR + Trim(Str(F.Size))
        Set F = Nothing
    End If

Next

V_VAR = V_VAR + Trim(Str(File3.ListCount))
For R = 0 To File3.ListCount - 1
    
    If InStr(UCase(File3.List(R)), ".SYS") = 0 Then
        Set F = FSO.GetFile(File3.Path + "\" + File3.List(R))
        V_VAR = V_VAR + File3.List(R)
        V_VAR = V_VAR + Format(F.DateLastModified, "YYYY-MM-DD HH-MM-SS")
        V_VAR = V_VAR + Trim(Str(F.Size))
        'If InStr(OV, V_VAR) = 0 And OV <> "" Then Stop
        'If InStr(File3.List(R), "page") > 0 Then Stop
        Set F = Nothing
    End If

Next

OV = V_VAR

If Err.Number = 0 Then
    Err_COUNTER_Number = 0
End If

If Err.Number > 0 Then
    'Debug.Print Err.descripton
    
    Err_COUNTER_Number = Err_COUNTER_Number + 1
    If Err_COUNTER_Number > 200 Then
        Me.WindowState = vbNormal
        MsgBox "ERROR COUNT BIGGER THAN > 200 FOR FOLDER " + vbCrLf + vbCrLf + File2.Path, vbMsgBoxSetForeground
    End If
    Exit Sub
End If

If O_C_DRIVEROOT_V_VAR <> V_VAR Then

    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
        WSHShell.Run """" + "C:\SCRIPTER\SCRIPTER CODE -- VBS\VBS 10-VICEVERSA _ SHELL FOLDERING__.AHK" + """ /RUN", 0, True
    Set WSHShell = Nothing

End If
O_C_DRIVEROOT_V_VAR = V_VAR

End Sub


Private Sub TIMER_DO_IN_1_SECOND_Timer()
O_GetForegroundWindow = 0
Call Timer_FOREGROUND_WINDOW_CHANGE_Timer
TIMER_DO_IN_1_SECOND.Enabled = False
End Sub

Private Sub Timer_VB_MAXIMIZE_Timer()

Dim Nx, Ny

'PURELY FOR VB ROUTINE IF ANOTHER ARE HERE

' ----------------------------------
' ALTERNATIVE LIKE HERE
' ----------------------------------
' Timer_01_VB_HELP_BOX_02_MSDN_Timer
' ---- OTHER -----------------------
' Timer_VB_MAXIMIZE
' ---- ALL THE VB ------------------
' ----------------------------------

'01 OF 02 DISABLE THE HELP BUTTON ACCIDENT PRESS
'FOCUS THE VB WINDOW ENTRY ON THE ESC KEY
'01 OF 02 THE MSDN BY CLOSE BY ESCAPE KEY


Dim mWnd_VB_VbaWindow_MAXIMIZE As Long
Dim GS_cWnd1 As Long, GS_cWnd2 As Long, GS_cWnd3 As Long, i3 As Long

'---------------------------------------
'01 OF 02 DISABLE THE HELP BUTTON ACCIDENT PRESS
'---------------------------------------
Dim MSDN_Lib, PWnd, cWnd
MSDN_Lib = FindWindow("#32770", "Microsoft Visual Basic")
If MSDN_Lib > 0 And MSDN_Lib <> OcWnd Then

    '---------------------------------------
    'RESULT = PostMessage(MSDN_Lib, WM_CLOSE, 0&, 0&)
    'Microsoft Visual Basic
    '#32770
    '---------------------------------------
    'INSTEAD OF METHOD 1 USE METHOD 2
    '---------------------------------------
    '1. DONT CLOSE THE HELP ALERT VB INFO
    '2. DISABLE THE HELP BUTTON ACIDENT FLICKER
    '---------------------------------------
    'HWND_2 = GET_CHILD_TEST(MSDN_Lib)
    '---------------------------------------
    
    PWnd = MSDN_Lib
    cWnd = FindWindowEx(PWnd, 0, "BUTTON", "HELP")
    EnableWindow cWnd, False
    
    '---------------------------------------
    '31 AUG 2K SIXTEEN
    'LEARN
    '---------------------------------------
    'Find Window Tutorial Part1 <A must see> by Paul Zaczkowski (from psc cd)
    'https://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37467&lngWId=1
    '---------------------------------------
End If

OcWnd = MSDN_Lib



'----------------------------------------
'FOCUS THE VB WINDOW ENTRY ON THE ESC KEY
'----------------------------------------

Dim CLASS_VAR

FindCursor Nx, Ny
mWnd_VB_VbaWindow_MAXIMIZE = WindowFromPoint(Nx, Ny)

CLASS_VAR = GetActiveWindowClass(mWnd_VB_VbaWindow_MAXIMIZE) 'Mwnd will get reset here if none
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Immediate" Then mWnd_VB_VbaWindow_MAXIMIZE = 0
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Locals" Then mWnd_VB_VbaWindow_MAXIMIZE = 0
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Watches" Then mWnd_VB_VbaWindow_MAXIMIZE = 0

'"VbaWindow" Or CLASS_VAR = "DesignerWindow" Then
If mWnd_VB_VbaWindow_MAXIMIZE > 0 Then
    If O_mWnd_VB_VbaWindow_MAXIMIZE <> mWnd_VB_VbaWindow_MAXIMIZE Then
        If CLASS_VAR = "VbaWindow" Or CLASS_VAR = "DesignerWindow" Then
            
            O_mWnd_VB_VbaWindow_MAXIMIZE = mWnd_VB_VbaWindow_MAXIMIZE
            'Beep
            ShowWindow mWnd_VB_VbaWindow_MAXIMIZE, SW_MAXIMIZE
            
            
            'NOT WORKER
            '----
            'Change caret size and blink rate
            'http://www.devx.com/vb2themax/Tip/18279
            '----
            '----------------------------------
            'Show a block cursor
            'ShowCustomCaret 6, 14
            '----------------------------------
            'CreateCaret mWnd_VB_VbaWindow_MAXIMIZE, 0, width, height
            CreateCaret mWnd_VB_VbaWindow_MAXIMIZE, 0, 6, 14
            ShowCaret mWnd_VB_VbaWindow_MAXIMIZE
            '----------------------------------
        End If
    End If
End If
'---------------------------------------



'------------------------------------------------------------------------------
' METHOD 3
' 1 ---- MICROSOFT VISUAL BASIC 01 OF 02 -- MSDN HELP ARRIVE ESC KEY GONE
' 1 ---- MICROSOFT VISUAL BASIC 02 OF 02 -- HELP ARRIVE MAKE BUTTON NOT ENABLED
' 2 ---- GOODSYNC ASK TO CLOSE AT END
' 3 ---- SECURITY BIT SET BUTTON PUSH ---- Error Applying Security
'------------------------------------------------------------------------------

'----------------------------------------------------
'TEMPROARY MISSING CODE WHEN CONFLICT COMPARE SYNC
'----------------------------------------------------
'Call Timer_01_VB_SOUNDBEEP_WHEN_BREAK_POINT_IDE_Timer
'----------------------------------------------------

'---------------------------------------
'01 OF 02 THE MSDN BY CLOSE BY ESCAPE KEY
'---------------------------------------
'CARE NOT MORE THAN ONE HELP WINDOW OPEN
'PRESS ESCAPE KEY CODE 27 TO CLOSE
'---------------------------------------
Dim MSDN_Library, Result
If GetAsyncKeyState(27) Then
    MSDN_Library = FindWindow(vbNullString, "MSDN Library Visual Studio 6.0")
    If MSDN_Library > 0 Then
        '----------------
        'CLOSE THE WINDOW
        '----------------
        Result = PostMessage(MSDN_Library, WM_CLOSE, 0&, 0&)
    End If
End If

'----------------------------------------
'FOCUS THE VB WINDOW ENTRY ON THE ESC KEY
'----------------------------------------

FindCursor Nx, Ny
mWnd_VB_VbaWindow_MAXIMIZE = WindowFromPoint(Nx, Ny)

CLASS_VAR = GetActiveWindowClass(mWnd_VB_VbaWindow_MAXIMIZE) 'Mwnd will get reset here if none
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Immediate" Then mWnd_VB_VbaWindow_MAXIMIZE = 0
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Locals" Then mWnd_VB_VbaWindow_MAXIMIZE = 0
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Watches" Then mWnd_VB_VbaWindow_MAXIMIZE = 0

'"VbaWindow" Or CLASS_VAR = "DesignerWindow" Then
If mWnd_VB_VbaWindow_MAXIMIZE > 0 Then
    If O_mWnd_VB_VbaWindow_MAXIMIZE <> mWnd_VB_VbaWindow_MAXIMIZE Then
        If CLASS_VAR = "VbaWindow" Or CLASS_VAR = "DesignerWindow" Then
            
            O_mWnd_VB_VbaWindow_MAXIMIZE = mWnd_VB_VbaWindow_MAXIMIZE
            Beep
            ShowWindow mWnd_VB_VbaWindow_MAXIMIZE, SW_MAXIMIZE
        End If
    End If
End If
'---------------------------------------


'---------------------------------------
'MAXIMIZE THE VB WINDOW ENTRY
'---------------------------------------
'MSDN_Lib = FindWindow("#32770", "Microsoft Visual Basic")
'IHWND = FindWindow(vbNullString, "wndclass_desked_gsk")
'I IHWND > 0
'    I_CH_HWND = FindWindowEx("VbaWindow", vbNullString)
'End If

'MAX WINDOWS WITHIN VB FOR ------------------------------------
'VbCode window To Max on Detect

FindCursor Nx, Ny
mWnd_VB_VbaWindow_MAXIMIZE = WindowFromPoint(Nx, Ny)

CLASS_VAR = GetActiveWindowClass(mWnd_VB_VbaWindow_MAXIMIZE) 'Mwnd will get reset here if none
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Immediate" Then mWnd_VB_VbaWindow_MAXIMIZE = 0
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Locals" Then mWnd_VB_VbaWindow_MAXIMIZE = 0
If GetWindowTitle(mWnd_VB_VbaWindow_MAXIMIZE) = "Watches" Then mWnd_VB_VbaWindow_MAXIMIZE = 0

If mWnd_VB_VbaWindow_MAXIMIZE > 0 Then
    If O_mWnd_VB_VbaWindow_MAXIMIZE <> mWnd_VB_VbaWindow_MAXIMIZE Then
        If CLASS_VAR = "VbaWindow" Or CLASS_VAR = "DesignerWindow" Then
            
            O_mWnd_VB_VbaWindow_MAXIMIZE = mWnd_VB_VbaWindow_MAXIMIZE
            Beep
            ShowWindow mWnd_VB_VbaWindow_MAXIMIZE, SW_MAXIMIZE
        End If
    End If
End If



'---------------------------------------
'02 OF 02 DISABLE THE HELP BUTTON ACCIDENT PRESS
'---------------------------------------
'2017_FEB_15 EARLY HOUR
'---------------------------------------
Dim TEAM_VIEWER
TEAM_VIEWER = FindWindow("#32770", "Sponsored session")
If TEAM_VIEWER > 0 And TEAM_VIEWER <> OHWnd_TEAM_VIEWER Then

    '---------------------------------------
    Result = PostMessage(TEAM_VIEWER, WM_CLOSE, 0&, 0&)
    'Microsoft Visual Basic
    '#32770
    '---------------------------------------
    'INSTEAD OF METHOD 1 USE METHOD 2
    '---------------------------------------
    '1. DONT CLOSE THE HELP ALERT VB INFO
    '2. DISABLE THE HELP BUTTON ACIDENT FLICKER
    '---------------------------------------
    'HWND_2 = GET_CHILD_TEST(MSDN_Lib)
    '---------------------------------------
    
'    pwnd = TEAM_VIEWER
'    cWnd = FindWindowEx(pwnd, 0, "BUTTON", "HELP")
'    EnableWindow cWnd, False
    
    PWnd = TEAM_VIEWER
    'GS_cWnd4 = 0
    'GS_cWnd1 = FindWindowEx(pwnd, 0, "Edit", vbNullString)
    
    'GS_cWnd2 = FindWindowEx(pwnd, 0, "Button4", "&OK")
    GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
    'GS_cWnd3 = FindWindowEx(pwnd, 0, "button", "&Cancel")
    
    
        
    If GS_cWnd2 > 0 Then
        '-------------------------------------------------
        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
        '1ST RUN DEBUG FOCUSED TO BUTTON BUT NOT HITT IT IN UNTIL 2ND
        'FOCUS PROBLEM FIRST
        '-------------------------------------------------
        Dim R_REPEAT
        For R_REPEAT = 1 To 3
            GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
            If GS_cWnd2 = 0 Then Exit For
            SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
    
    '---------------------------------------
    '31 AUG 2K SIXTEEN
    'LEARN
    '---------------------------------------
    'Find Window Tutorial Part1 <A must see> by Paul Zaczkowski (from psc cd)
    'https://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37467&lngWId=1
    '---------------------------------------
End If

OHWnd_TEAM_VIEWER = TEAM_VIEWER







'---------------------------------------
'03 OF 03 VISIAUL BASIC LOAD CRASH THE CLIPBOARD
'---------------------------------------
'2017_FEB_22 EARLY HOUR 01:43a
'---------------------------------------
'TYPE 01 OF 02
Dim VB_LOADER
VB_LOADER = FindWindow("#32770", "Visual Component Manager")
If VB_LOADER > 0 And VB_LOADER <> OHWnd_VB_LOADER Then
    PWnd = VB_LOADER
    GS_cWnd1 = FindWindowEx(PWnd, 0, vbNullString, "Can't open Clipboard")
    GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
    If GS_cWnd1 > 0 And GS_cWnd2 > 0 Then
        '-------------------------------------------------
        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
        '1ST RUN DEBUG FOCUSED TO BUTTON BUT NOT HITT IT IN UNTIL 2ND
        'FOCUS PROBLEM FIRST
        '-------------------------------------------------
        For R_REPEAT = 1 To 4
            GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
            If GS_cWnd2 = 0 Then Exit For
            SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
End If



'TYPE 02 OF 02
VB_LOADER = FindWindow("#32770", "Data View")
If VB_LOADER > 0 And VB_LOADER <> OHWnd_VB_LOADER Then
    PWnd = VB_LOADER
    GS_cWnd1 = FindWindowEx(PWnd, 0, vbNullString, "Method '~' of object '~' failed")
    GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
    If GS_cWnd1 > 0 And GS_cWnd2 > 0 Then
        For R_REPEAT = 1 To 4
                GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
                If GS_cWnd2 = 0 Then Exit For
                SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
End If
'TYPE 02 OF 02
'VB_LOADER = FindWindow("#32770", "Add-In Toolbar")
VB_LOADER = FindWindow("#32770", vbNullString)
If VB_LOADER > 0 And VB_LOADER <> OHWnd_VB_LOADER Then
    'OHWnd_VB_CLIPPER_ERROR = VB_CLIPPER_ERROR
    OHWnd_VB_LOADER = VB_LOADER
    PWnd = VB_LOADER
    GS_cWnd1 = FindWindowEx(PWnd, 0, vbNullString, "Method '~' of object '~' failed")
    GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
    If GS_cWnd1 > 0 And GS_cWnd2 > 0 Then
        For R_REPEAT = 1 To 4
                GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
                If GS_cWnd2 = 0 Then Exit For
                SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
End If

Dim VB_CLIPPER_ERROR
VB_CLIPPER_ERROR = FindWindow("#32770", vbNullString)
If VB_CLIPPER_ERROR > 0 And VB_CLIPPER_ERROR <> OHWnd_VB_CLIPPER_ERROR Then
    OHWnd_VB_CLIPPER_ERROR = VB_CLIPPER_ERROR
    PWnd = VB_CLIPPER_ERROR
    GS_cWnd1 = FindWindowEx(PWnd, 0, vbNullString, "Can't open Clipboard")
    GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
    If GS_cWnd1 > 0 And GS_cWnd2 > 0 Then
        For R_REPEAT = 1 To 4
                GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
                If GS_cWnd2 = 0 Then Exit For
                SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
End If




'KILL ITSELF IF RUNNING OR GETTER IN THE WAY OF LEARN
Dim VB_EXE
If IsIDE = True Then
    VB_EXE = FindWindow("ThunderRT6FormDC", "KAT MP3 RELOAD")
    If VB_EXE > 0 And VB_EXE <> OHWnd_VB_EXE Then
        Result = PostMessage(VB_EXE, WM_CLOSE, 0&, 0&)
    End If
End If


'TYPE 02 OF 02
VB_LOADER = FindWindow("#32770", "Add-In Toolbar")
If 1 = 2 And VB_LOADER > 0 And VB_LOADER <> OHWnd_VB_LOADER Then
'Add-In Toolbar

    '---------------------------------------
    'RESULT = PostMessage(VB_LOADER, WM_CLOSE, 0&, 0&)
    'Microsoft Visual Basic
    '#32770
    '---------------------------------------
    'INSTEAD OF METHOD 1 USE METHOD 2
    '---------------------------------------
    '1. DONT CLOSE THE HELP ALERT VB INFO
    '2. DISABLE THE HELP BUTTON ACIDENT FLICKER
    '---------------------------------------
    'HWND_2 = GET_CHILD_TEST(MSDN_Lib)
    '---------------------------------------
    
'    pwnd = VB_LOADER
'    cWnd = FindWindowEx(pwnd, 0, "BUTTON", "HELP")
'    EnableWindow cWnd, False
    
    PWnd = VB_LOADER
    'GS_cWnd4 = 0
    'GS_cWnd1 = FindWindowEx(pwnd, 0, "Edit", vbNullString)
    
    GS_cWnd1 = FindWindowEx(PWnd, 0, vbNullString, "Data View")
    GS_cWnd1 = FindWindowEx(PWnd, 0, vbNullString, "Can't open Clipboard")
    'Can't open Clipboard
    
    'GS_cWnd2 = FindWindowEx(pwnd, 0, "Button4", "&OK")
    GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
    'GS_cWnd3 = FindWindowEx(pwnd, 0, "button", "&Cancel")
    If GS_cWnd1 > 0 And GS_cWnd2 > 0 Then
    
        
        '-------------------------------------------------
        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
        '1ST RUN DEBUG FOCUSED TO BUTTON BUT NOT HITT IT IN UNTIL 2ND
        'FOCUS PROBLEM FIRST
        '-------------------------------------------------
        For R_REPEAT = 1 To 4
            GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
            If GS_cWnd2 = 0 Then Exit For
            SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
    
    '---------------------------------------
    '31 AUG 2K SIXTEEN
    'LEARN
    '---------------------------------------
    'Find Window Tutorial Part1 <A must see> by Paul Zaczkowski (from psc cd)
    'https://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37467&lngWId=1
    '---------------------------------------
End If


'FILE ALREADY open ____ LOADED _ MOVE ON
'-----------------------------
Dim TEXT1
VB_LOADER = FindWindow("#32770", "Microsoft Visual Basic")
If 1 = 2 And VB_LOADER > 0 And VB_LOADER <> OHWnd_VB_LOADER Then
    PWnd = VB_LOADER
    GS_cWnd1 = FindWindowEx(PWnd, 0, vbNullString, "OK")
    GS_cWnd1 = FindWindowEx(PWnd, 0, "Static", vbNullString)
    TEXT1 = WindowText(GS_cWnd1)
    
    GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
    If GS_cWnd1 > 0 And GS_cWnd2 > 0 Then
        '-------------------------------------------------
        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
        '1ST RUN DEBUG FOCUSED TO BUTTON BUT NOT HITT IT IN UNTIL 2ND
        'FOCUS PROBLEM FIRST
        '-------------------------------------------------
        For R_REPEAT = 1 To 4
            GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
            If GS_cWnd2 = 0 Then Exit For
            SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
End If












End Sub


Private Sub Timer_FOREGROUND_WINDOW_CHANGE_Timer()
    
If Timer_FOREGROUND_WINDOW_CHANGE.Interval <> 1000 Then
    Timer_FOREGROUND_WINDOW_CHANGE.Interval = 1000
End If

If GetForegroundWindow <> O_GetForegroundWindow Then
    If Me.WindowState = vbMinimized Then TIMER_DO_IN_1_SECOND.Enabled = True
End If
If FORM_LOAD_EnumProcess = False Then
    If Me.WindowState = vbMinimized Then Exit Sub
End If
 FORM_LOAD_EnumProcess = False
'If GetForegroundWindow <> Me.hWnd Then
'    If Me.WindowState = vbMinimized Then Exit Sub
'End If

If GetForegroundWindow <> O_GetForegroundWindow Then
    Call EnumProcess
End If


If InStr(MNU_WINMERGE_ON_TOP_ALLTME.Caption, "ALLTIME=YES") > 0 Then

    IHWND = FindWindow("WinMergeWindowClassW", vbNullString)
    If IHWND > 0 And O_IHWND <> IHWND Then
        
        SetWindowPos IHWND, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    End If
    O_IHWND = IHWND
End If


O_GetForegroundWindow = GetForegroundWindow

Dim R, A1, A2

For R = 2 To lstProcess_3_SORTER_ListView.ListItems.Count
    A1 = lstProcess_3_SORTER_ListView.ListItems.Item(R).SubItems(1)
    A2 = lstProcess_3_SORTER_ListView.ListItems.Item(R - 1).SubItems(1)
    If A1 = A2 Then
        If A1 = "VideoDownloaderUltimate.exe" Then
            Beep
            Process_Kill (Val(lstProcess_3_SORTER_ListView.ListItems.Item(R)))
        End If
    
    If A1 = "Windows10UpgraderApp.exe" Then
            Beep
            Process_Kill (Val(lstProcess_3_SORTER_ListView.ListItems.Item(R)))
    End If
    
    End If
Next



'Process_Kill (INETSource)


End Sub


' Enumerate open processes
Private Sub EnumProcess(Optional ByVal sEXEName As String = "")

    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As PROCESSENTRY32
    Dim TIME_LAB
    
    TIME_LAB = Format(Now, "YYYY MMM DD DDD HH:MM:SS am/pm")
    Mid(TIME_LAB, 24, 1) = "0"
    TIME_LAB = Mid(TIME_LAB, 1, 24) + "_._" + Mid(Format(Timer - Int(Timer), ".000"), 2) + "" + Mid(TIME_LAB, 25)
    'Label53_Here
    'Label53.Caption = "Process Set Enumerate Event: " + TIME_LAB + " _ Ticker"
    
    ' Create snapshot
    lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If lSnapShot <> -1 Then
        ' Clear list
        'LISTBOX
        lstProcess.Clear
        'lstProcess_SORTER.Clear
        ' Length of the structure
        tPE.dwSize = Len(tPE)
        
        ' Find first process
        lNextProcess = Process32First(lSnapShot, tPE)
        Do While lNextProcess
            ' Found specified process
            If sEXEName = Left$(tPE.szExeFile, Len(sEXEName)) And Len(sEXEName) > 0 Then
                Dim lProcess As Long
                Dim lExitCode As Long
                ' Open process
                lProcess = OpenProcess(0, False, tPE.th32ProcessID)
                ' Terminate process
                TerminateProcess lProcess, lExitCode
                ' Close handle
                CloseHandle lProcess
            Else
                ' Add exe to list
                lstProcess.AddItem Format(tPE.th32ProcessID, "000000 ") + tPE.szExeFile
'                lstProcess_2_.AddItem tPE.szExeFile
'                lstProcess_3_SORTER.AddItem tPE.szExeFile
            End If
            ' Get next process
            lNextProcess = Process32Next(lSnapShot, tPE)
        Loop
        
        ' Close handle
        CloseHandle (lSnapShot)
        
    Else
        lstProcess.AddItem "Cannot enumerate running process!"
    End If
    
    
    'If lstProcess.ListCount = 0 Then Exit Sub
    
    
    
'Dim REPROCESS, R_I, STRING_COMPARE
'REPROCESS = False
''GET AND EXTRA FIRST RUN WITH NOTHING AS RESULT
''OKAY NOW
''----------------------------------------------
'
'For R_I = 0 To lstProcess.ListCount - 1
'
'    STRING_COMPARE = STRING_COMPARE + lstProcess.List(R_I)
'
'    '                lstProcess_2_.AddItem tPE.szExeFile
'    '                lstProcess_3_SORTER.AddItem tPE.szExeFile
'Next
'If O_STRING_COMPARE <> STRING_COMPARE Then
'    REPROCESS = True
'End If
'O_STRING_COMPARE = STRING_COMPARE
'
'If REPROCESS = False Then Exit Sub
'
'lstProcess_2_.Clear
'lstProcess_3_SORTER.Clear
'
'Dim LV2 As ListItem ' NOT CONFUSE __ LISTVIEW WITH __ LISTITEM
'Dim LV3 As ListItem ' NOT CONFUSE __ LISTVIEW WITH __ LISTITEM
'Dim ITEM_ADD_10 As String
'Dim ITEM_ADD_21 As String
'Dim ITEM_ADD_22 As String
'
'lstProcess_2_ListView.ListItems.Clear
'lstProcess_3_SORTER_ListView.ListItems.Clear
'For R_I = 0 To lstProcess.ListCount - 1
'
'    ITEM_ADD_10 = Mid(lstProcess.List(R_I), 1, 7 - 1)
'    ITEM_ADD_21 = Mid(lstProcess.List(R_I), 8)
'    ITEM_ADD_22 = UCase(Mid(lstProcess.List(R_I), 8, 1)) + Mid(lstProcess.List(R_I), 9)
'
'    lstProcess_2_.AddItem lstProcess.List(R_I)
'    lstProcess_3_SORTER.AddItem ITEM_ADD_10 + " " + ITEM_ADD_22
'
'    With lstProcess_2_ListView
'        Set LV2 = .ListItems.Add(, , ITEM_ADD_10)
'        LV2.SubItems(1) = ITEM_ADD_21
'        '------------------------
'        'PAIN TO GET THE FORMULAR
'        'EVEN THEN FAR OUT AGAIN
'        '------------------------
'    End With
'
'    With lstProcess_3_SORTER_ListView
'        Set LV3 = .ListItems.Add(, , ITEM_ADD_10)
'        LV3.SubItems(1) = ITEM_ADD_22
'        '------------------------
'        'PAIN TO GET THE FORMULAR
'        'EVEN THEN FAR OUT AGAIN
'        '------------------------
'    End With
    
    
'
'
Dim REPROCESS, R_I, STRING_COMPARE
REPROCESS = False
'GET AND EXTRA FIRST RUN WITH NOTHING AS RESULT
'OKAY NOW
'----------------------------------------------

For R_I = 0 To lstProcess.ListCount - 1

    STRING_COMPARE = STRING_COMPARE + lstProcess.List(R_I)

    '                lstProcess_2_.AddItem tPE.szExeFile
    '                lstProcess_3_SORTER.AddItem tPE.szExeFile
Next
If O_STRING_COMPARE <> STRING_COMPARE Then
    REPROCESS = True
End If
O_STRING_COMPARE = STRING_COMPARE

If REPROCESS = False Then Exit Sub


Dim LV2 As ListItem ' NOT CONFUSE __ LISTVIEW WITH __ LISTITEM __ MUST BE __ ListItem
Dim LV3 As ListItem ' NOT CONFUSE __ LISTVIEW WITH __ LISTITEM
Dim ITEM_ADD_10 As String
Dim ITEM_ADD_21 As String
Dim ITEM_ADD_22 As String

'----------------------------------------------
'HERE IS ALREADY DONE IN FRMMAIN EXAMPLE SEARCH
'----------------------------------------------
'    With lstProcess_2_ListView
'        .ColumnHeaders.Add , "PID", "PID", 700 - 50, lvwColumnLeft
'        .ColumnHeaders.Add , "EXE", "EXE", 9000, lvwColumnLeft
'        .View = lvwReport
'    End With
'    With lstProcess_3_SORTER_ListView
'        .ColumnHeaders.Add , "PID", "PID", 700 - 50, lvwColumnLeft
'        .ColumnHeaders.Add , "EXE SORTED", "EXE SORTED", 9000, lvwColumnLeft
'        .View = lvwReport
'    End With


'If Timer_Pause_Update.Enabled = True Then Exit Sub


lstProcess_2_ListView.ListItems.Clear
lstProcess_3_SORTER_ListView.ListItems.Clear

For R_I = 0 To lstProcess.ListCount - 1

    ITEM_ADD_10 = Mid(lstProcess.List(R_I), 1, 7 - 1)
    ITEM_ADD_21 = Mid(lstProcess.List(R_I), 8)
    ITEM_ADD_22 = UCase(Mid(lstProcess.List(R_I), 8, 1)) + Mid(lstProcess.List(R_I), 9)

    'IF ERROR HERE ABOUT METHOD OR DATA MEMBER NOT FOUND
    'THEN MAYBE ALREADY LEARN THE REFERENCE LIST VIEW
    'IN ORDER TO USE Dim LV2 As ListItem
    'THE ListItem
    'MUST HAVE A SPECIAL REFERENCE FIND THAT ONE
    'HERE
    '----
    'WHAT IS THE VB6 REFERENCE TO USE LIST ITEM - Google Search
    'https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB744GB744&q=WHAT+IS+THE+VB6+REFERENCE+TO+USE+LIST+ITEM&spell=1&sa=X&ved=0ahUKEwjFltmnn4LWAhVDKlAKHfArAFwQvwUIJSgA&biw=1536&bih=694
    '--------
    'ListItem Object, ListItems Collection
    'https://msdn.microsoft.com/en-us/library/aa443480(v=vs.60).aspx
    '--------
    'Add Method (ListItems, ColumnHeaders), SubItems Property Example
    'https://msdn.microsoft.com/en-us/library/aa443210(v=vs.60).aspx
    '----
    '-------------------------------------------------------------
    'CORRECT AND AFTER HERE EXIT THE CODE AN BACK IN AGAIN
    '-------------------------------------------------------------
    'DON'T FORGET AFTER AND LOSE OF LISTVIEW CONTROL TO PICTUREBOX
    'REPLACE GUESS HERE ALSO
    '-------------------------------------------------------------
    
    
    '------------------------------------------------
'    FORM1.lstProcess_2_ListView.ListItems.Add , , lstProcess.List(R_I)
'    FORM1.lstProcess_3_SORTER_ListView.ListItems.Add , , ITEM_ADD_10 + " " + ITEM_ADD_22


    '----------------------------------------------------
    'IF ERROR HERE ABOUT TYPE MISMATCH _ AS ABOVE LEARNER
    'THEN MAYBE ALREADY LEARN THE REFERENCE LIST VIEW
    '----------------------------------------------------
    With lstProcess_2_ListView
        Set LV2 = .ListItems.Add(, , ITEM_ADD_10)
        LV2.SubItems(1) = ITEM_ADD_21
        '------------------------
        'PAIN TO GET THE FORMULAR
        'EVEN THEN FAR OUT AGAIN
        '------------------------
    End With

    With lstProcess_3_SORTER_ListView
        Set LV3 = .ListItems.Add(, , ITEM_ADD_10)
        LV3.SubItems(1) = ITEM_ADD_22
        '------------------------
        'PAIN TO GET THE FORMULAR
        'EVEN THEN FAR OUT AGAIN
        '------------------------
    End With
    
Next

Set LV2 = Nothing
Set LV3 = Nothing
Dim TAG_VAR
Dim TAG_VAR_2

'Label46.Caption = "Process Counter Is:" ' + Str(lstProcess.ListCount - 2)

If O_lstProcess_ListCount = -1 Then
    O_lstProcess_ListCount = lstProcess.ListCount
End If

If O_lstProcess_ListCount > lstProcess.ListCount Then TAG_VAR = " With_er Less"
If O_lstProcess_ListCount < lstProcess.ListCount Then TAG_VAR = " With_er Higher"
If O_lstProcess_ListCount = lstProcess.ListCount Then TAG_VAR = " Equal"

If O_lstProcess_ListCount > lstProcess.ListCount Then
    'UP
End If
If O_lstProcess_ListCount < lstProcess.ListCount Then
    'DOWN
    On Error Resume Next
    Err.Clear
End If
If O_lstProcess_ListCount = lstProcess.ListCount Then
    'EQUAL
    
    On Error Resume Next
    Err.Clear
End If
'----
'Wingdings 3 character set and equivalent Unicode characters
'http://www.alanwood.net/demos/wingdings-3.html
'----

Label50.Caption = Str(lstProcess.ListCount - 2) + TAG_VAR + " _ Timer One Minute"

O_lstProcess_ListCount = lstProcess.ListCount

Call LV_AutoSizeColumn(lstProcess_2_ListView, lstProcess_2_ListView.ColumnHeaders.Item(1))
Call LV_AutoSizeColumn(lstProcess_2_ListView, lstProcess_2_ListView.ColumnHeaders.Item(2))
Call LV_AutoSizeColumn(lstProcess_3_SORTER_ListView, lstProcess_3_SORTER_ListView.ColumnHeaders.Item(1))
Call LV_AutoSizeColumn(lstProcess_3_SORTER_ListView, lstProcess_3_SORTER_ListView.ColumnHeaders.Item(2))

'------------------------
'SORT ON COLUMN TWO LABEL
'------------------------
Dim GO_X
GO_X = 0
If lstProcess_3_SORTER_ListView.SortKey <> 1 Then GO_X = 1
If lstProcess_3_SORTER_ListView.SortOrder <> lvwAscending Then GO_X = 1
If lstProcess_3_SORTER_ListView.Sorted = True <> True Then GO_X = 1

If GO_X = 0 Then Exit Sub

lstProcess_3_SORTER_ListView.SortOrder = lvwAscending
lstProcess_3_SORTER_ListView.SortKey = 1
lstProcess_3_SORTER_ListView.Sorted = True
'lstProcess_3_SORTER_ListView.Sorted = False
    


End Sub

Private Sub LV_AutoSizeColumn(LV As ListView, Optional Column _
 As ColumnHeader = Nothing)

 Dim C As ColumnHeader
 If Column Is Nothing Then
  For Each C In LV.ColumnHeaders
   SendMessage LV.hWnd, LVM_FIRST + 30, C.Index - 1, -1
  Next
 Else
  SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
 End If
 LV.Refresh
End Sub



Private Function WindowText(ByVal window_hwnd As Long) As String

    Dim txtlen              As Long

    If window_hwnd = 0 Then
        Exit Function
    End If

    txtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, ByVal 0, ByVal 0)
    If txtlen = 0 Then
        Exit Function
    End If

    WindowText = String$(txtlen, vbNullChar)
    SendMessage window_hwnd, WM_GETTEXT, ByVal (txtlen + 1), ByVal StrPtr(WindowText)

End Function


Function GetActiveWindowTitle(ReturnParent As Long) As String
   Dim i As Long
   Dim j As Long
   i = GetForegroundWindow
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
   i = j
   ReturnParent = j
   GetActiveWindowTitle = GetWindowTitle(i)
End Function
Function GetActiveWindowClassParent(ReturnParent As Long) As String
   ReturnParent = 0
   Dim i As Long
   Dim j As Long
   i = GetForegroundWindow
    Do While i <> 0
        j = i
        i = GetParent(i)
    Loop
    i = j
    ReturnParent = j
    GetActiveWindowClassParent = GetWindowClass(i)
End Function
Function GetActiveWindowClass(ReturnParent As Long) As String
    GetActiveWindowClass = GetWindowClass(ReturnParent)
End Function


Function GetActiveWindow(ByVal ReturnParent As Boolean) As Long
   Dim i As Long
   Dim j As Long
   i = GetForegroundWindow
   If ReturnParent Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
      i = j
   End If
   GetActiveWindow = i
End Function


Function GetParentTOP(ByVal Handle As Long) As String
   Dim i As Long
   Dim j As Long, TTx1 As String
   i = Handle
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
      i = j
   '
   TTx1 = GetWindowTitle(i)
   GetParentTOP = i

End Function


Function GetParentTitle(ByVal Handle As Long) As String
   Dim i As Long
   Dim j As Long, TTx1 As String
   i = Handle
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
      i = j
   TTx1 = GetWindowTitle(i)
   GetParentTitle = TTx1

End Function

Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim L As Long
   Dim S As String
   L = GetWindowTextLength(hWnd)
   S = Space(L + 1)
   GetWindowText hWnd, S, L + 1
   GetWindowTitle = Left$(S, L)
End Function
Function GetWindowClass(ByVal hWnd As Long) As String
    Dim Ret As Long, sText As String
    sText = Space(255)
    Ret = GetClassName(hWnd, sText, 255)
    sText = Left$(sText, Ret)
   GetWindowClass = sText
End Function



Private Sub FindCursor(X, Y)

Dim P As POINTAPI

GetCursorPos P
'   return x and y co-ordinate
X = P.X ' / GetSystemMetrics(0) * Screen.Width
'   for current cursor position
Y = P.Y '/ GetSystemMetrics(1) * Screen.Height

End Sub


Function FindWindowPart(Test) As Long

FindWindowPart = False

Dim test_hwnd As Long

test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)
Do While test_hwnd <> 0
        
        If InStr(GetWindowTitle(test_hwnd), Test) > 0 Then
            FindWindowPart = test_hwnd
            Exit Do
        End If
        
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop


End Function



Private Function CreateFolderTree(ByVal sPath As String) As Boolean
    Dim nPos As Integer

    If Mid(sPath, Len(sPath), 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)

    On Error GoTo CreateFolderTreeError
    
    nPos = InStr(sPath, "\")
    While nPos > 0
        If Not FolderExists(Left$(sPath, nPos - 1)) Then
            MkDir Left$(sPath, nPos - 1)
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
    If Not FolderExists(sPath) Then MkDir sPath
    
    CreateFolderTree = True
    Exit Function

CreateFolderTreeError:
    Exit Function
End Function

Private Function FolderExists(sFolder As String) As Boolean
    '##############################################################################################
    'Returns True if the specified folder exists
    '##############################################################################################
    
    Dim WFD As WIN32_FIND_DATA
    Dim lResult As Long
    
    lResult = FindFirstFile(sFolder, WFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            FolderExists = True
        Else
            FolderExists = False
        End If
    End If
End Function


'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  'IsIDE = False
  'Exit Function
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************




'Public Sub DeleteDupe()
'
'Dim nMaxCount As Long
'Dim lpClassName As String
'Dim hand3 As Long
'Dim cch As Long
'Dim lpString As String
'
'Dim Handle2
'Dim Handle3
'Dim FileHWND2
'
'Handle2 = FindWindow(vbNullString, "Deleting...")
'Handle3 = FindWindow(vbNullString, "ImageDupe registered to EPSiLON")
'
''GetWindowState
'FileHWND2 = ""
'If Handle2 > 0 Then FileHWND2 = GetFileFromHwnd(Handle2)
'
'If FileHWND2 = "C:\Program Files\ImageDupe\ImageDupe.exe" And _
'    Check1.Value = vbChecked Then
''If handle2 > 0 Then
''If DupeCool = 0 Then
''handle3 = FindWindow(vbNullString, "ImageDupe registered to EPSiLON")
''MsgBox "Image Dupe Ready": DupeCool = 1
'Dim Rect5 As RECT
'Dim Rect7 As RECT
'hwnd9 = GetWindowRect(Handle2, Rect5)
'hWnd8 = GetWindowRect(Handle3, Rect7)
'
''CID_Run_Me.Top = (Rt2.Top * ScreenTwipsX) - CID_Run_Me.Height
'
'
'HX = 0 - (Rect5.Right - Rect5.Left) + 45
'HY = 40
'HW = Rect5.Right - Rect5.Left
'HH = Rect5.Bottom - Rect5.Top
'MoveWindow Handle2, HX, HY, HW, HH, True
'
'End If
'
''If handle2 = 0 And FindWindow(vbNullString, "ImageDupe registered to EPSiLON") = 0 Then DupeCo
'
'
'FileHWND2 = ""
'If Handle2 > 0 Then FileHWND2 = GetFileFromHwnd(Handle2)
'
'If FileHWND2 = "C:\Program Files\ImageDupe\ImageDupe.exe" And _
'    Check1.Value = vbChecked Then
'
''    kk = Int(Now) + Timer + 0.2
''    Do
''        DoEvents
''    Loop Until kk < Int(Now) + Timer
'
'    On Local Error Resume Next
'    Err.Clear
'    AppActivate "Deleting...", False
'
'    If DupeCool2 = 0 And Err.Number = 0 Then
'    ShowWindow Handle3, SW_NORMAL 'Restore
'    DupeCool2 = 1
'    End If
'
'    'Err.Description
'
'    If Err.Number > 0 Then DupeCool2 = 0: Exit Sub
'    On Local Error GoTo 0
'
''    kk = Int(Now) + Timer + 0.8
''    Do
''        DoEvents
''    Loop Until kk < Int(Now) + Timer
'    'Ew = Int(Now) + Timer + 1.5
'    Ew = Int(Now) + Timer + 1
'    ash$ = ""
'    Do
'        DoEvents
'        ash$ = GetActiveWindowTitle(False)
'    Loop Until Ew < Int(Now) + Timer Or ash$ = "Deleting..."
'
'
''    Ew = Now + TimeSerial(0, 0, 2)
'
''    Ew = Int(Now) + Timer + 1.5
'    Ew = Int(Now) + Timer + 1
'
'    ash$ = ""
'    Do
'        DoEvents
'        ash$ = GetActiveWindowTitle(False)
'    Loop Until Ew < Int(Now) + Timer Or ash$ = "Deleting..."
'
'
''    SendKeys "%D", True
'    If ash$ = "Deleting..." Then SendKeys "{ENTER}", True
'
'    'Ew = Int(Now) + Timer + 0.4
'    'Do
'    '    DoEvents
'    'Loop Until Ew < Int(Now) + Timer
'
'End If
'
'End Sub

'Public Sub DocsToGo()
'
'
'Handle2 = FindWindow(vbNullString, "DocumentsTo Go")
'If Handle2 = 0 Then Exit Sub
'FileHWND2 = ""
'If Handle2 > 0 Then FileHWND2 = GetFileFromHwnd(Handle2)
''CM$ = GetFileFromProc(CMediaSource)
'
'If FileHWND2 = "C:\Program Files\Documents To Go\DocsToGo.exe" Then
'
'
'    On Local Error Resume Next
'    Err.Clear
'    AppActivate "DocumentsTo Go", False
'    If Err.Number > 0 Then Exit Sub
'    On Local Error GoTo 0
'
'    ash$ = ""
'    Do
'        DoEvents
'        ash$ = GetActiveWindowTitle(False)
'    Loop Until Ew < Int(Now) + Timer Or ash$ = "DocumentsTo Go"
'
'
'    SendKeys "{ENTER}", True
'
'End If
'
'End Sub

'
'Private Sub Timer7_Timer()
'Timer7.Enabled = False
'Exit Sub
'
'Call DocsToGo

'Exit Sub

'Afx:10000000:0
'
'If FindWindow("Afx:10000000:0", vbNullString) > 0 Then
'    NeroDupe$ = ""
'    W = FindWindowPart(NeroDupe$)
'    If NeroDupe$ <> "" And NeroDupe$ <> NeroDupeTT$ Then
'        'MsgBox "Two Neros Of Same Running" + vbCrLf + NeroDupe$
'        Frm2Nero.Label1 = "Two Neros Of Same Running  " + vbCrLf + NeroDupe$ + "  "
'        Frm2Nero.Show
'        NeroDupeTT$ = NeroDupe$
'    End If
'End If
'If FindWindow("Afx:10000000:0", vbNullString) = 0 Then
'NeroDupeTT$ = ""
'End If
'
'
'
'
'If FindWindow(vbNullString, "Deleting...") > 0 Then
'If FindWindow(vbNullString, "ImageDupe registered to EPSiLON") > 0 Then
'    Call DeleteDupe
'End If
'End If
'
'
'
'
'
'GoTo skip1
'Xxr = FindWindow(vbNullString, "File has been modified")
'If Xxr > 0 Then
'    On Local Error Resume Next
'    Err.Clear
'    'AppActivate "File has been modified", False
'    SetForegroundWindow Xxr
'
'    If Err.Number > 0 Then Exit Sub
'    On Local Error GoTo 0
'    SendKeys "{ENTER}", True
'End If
'skip1:
'
'
'
'Exit Sub
'
'notepad2rt = FindWindow("#32770", "Notepad2")
'notepad3rt = FindWindow("Notepad2", vbNullString)
'xxrp = GetParent(notepad2rt)
'cuecue$ = GetWindowTitle(notepad3rt)
'If xxrp <> notepad3rt Then Exit Sub
'
'If notepad2rt = 0 Then Exit Sub
'
'If NotePad2RT2 = notepad3rt Then Exit Sub
'
'If GetForegroundWindow <> notepad2rt Then Exit Sub
'
'If InStr(cuecue$, "Untitled") > 0 Then Exit Sub
'
'If Mid$(cuecue$, 1, 1) <> "*" Then Exit Sub
'
''THIS MAKES ERROR KEEP SWAPING CAPS LOcK KEY
''SendKeys "{ENTER}", False
't = Timer + 0.3
'Do
'DoEvents
'If Timer < t - 20 Then Exit Do
'Loop Until t < Timer
'
'SendKeys "%Y", False
'
'If GetForegroundWindow = notepad2rt Then Exit Sub
'
'NotePad2RT2 = notepad3rt
'
'Exit Sub
'
'
'
'If FindWindow(vbNullString, "Enter Network Password") > 0 Then
'    On Local Error Resume Next
'    Err.Clear
'    AppActivate "Enter Network Password", False
'    If Err.Number > 0 Then Exit Sub
'    On Local Error GoTo 0
'    SendKeys "%s{ENTER}", True
'End If
'
'XX1 = FindWindow(vbNullString, "Microsoft Publisher")
'xx2 = FindWindow(vbNullString, "Microsoft Office Publisher")
'
'If XX1 Or xx2 > 0 Then
'    On Local Error Resume Next
'    Err.Clear
'    If Mnu_Publish.Checked = True Then
'        If XX1 > 0 Then AppActivate "Microsoft Publisher", False
'        If xx2 > 0 Then AppActivate "Microsoft Office Publisher", False
'    End If
'    If Err.Number > 0 Then Exit Sub
'    On Local Error GoTo 0
'    If Mnu_Publish.Checked = True Then SendKeys "{ENTER}", True
'End If
'
'
'
''Microsoft Office Outlook
'
'End Sub




'Sub SUB_001()
'
'
''INET = FindWindow("IEFrame", vbNullString)
''
''If INET3 <> INET And INET > 0 Then
''    INET3 = INET
''    TTT = cProcesses.Convert(INET, INETSource, cnFromhWnd Or cnToProcessID)
''    INETSource2 = Now + TimeSerial(0, 0, 10)
''End If
''
''INET2 = FindWindow("IEFrame", vbNullString)
''If INETSource > 0 And INET2 = 0 And INETSource2 < Now Then
''
''    Dim INETSourceTEST As Long
''    TTT = cProcesses.Convert(INET3, INETSourceTEST, cnFromhWnd Or cnToProcessID)
''
''    CM$ = GetFileFromProc(INETSource)
''    If CM$ <> "" And TTT = True Then
''        Process_Kill (INETSource)
''        MsgBox "INET PROCESS ID KILLED AS NOT IN USE", vbMsgBoxSetForeground
''        INETSource = 0
''    End If
''End If
'
'
'
'
'
'
'
'
'
'
'CMedia = FindWindow(vbNullString, "Creative MediaSource")
'
'If CMedia3 <> CMedia And CMedia > 0 Then
'    CMedia3 = CMedia
'    TTT = cProcesses.Convert(CMedia, CMediaSource, cnFromhWnd Or cnToProcessID)
'    CMediaSource2 = Now + TimeSerial(0, 0, 2)
'End If
'
'CMedia2 = FindWindow(vbNullString, "Creative MediaSource")
'If CMediaSource > 0 And CMedia2 = 0 And CMediaSource2 < Now Then
'    CM$ = GetFileFromProc(CMediaSource)
'    If CM$ <> "" Then Process_Kill (CMediaSource)
'    CMediaSource = 0
'End If
'
'
'CNote = "C:\WINDOZE\system32\NOTEPAD.EXE"
'
'ttt3 = cProcesses.Convert(CNote, CNoteSource, cnFromEXE Or cnToProcessID)
'If ttt3 = True Then
'    TTT2 = cProcesses.Convert(CNoteSource, CNoteSource2, cnFromProcessID Or cnTohWnd)
'    If CNoteSource4 = 0 And TTT2 = False Then CNoteSource4 = Now + TimeSerial(0, 0, 1)
'End If
'
'
'If CNoteSource > 0 And TTT2 = False And ttt3 = True And CNoteSource4 < Now Then
'    TTT2 = cProcesses.Convert(CNoteSource, CNoteSource2, cnFromProcessID Or cnTohWnd)
'    'If TTT2 = False Then Process_Kill (CNoteSource)
'    CNoteSource = 0
'    CNoteSource4 = 0
'End If
'
'
'
'
'End Sub


'
'Sub SUB_002()
'
'
'Dim Xxr As Long
'
'
'
'
'Rf = FindWindow(vbNullString, "WMI Demo - CPU Information")
'If CurProcHwnd = Me.hWnd And CurProcHwnd <> Rf And TiTBait <> CurProcHwnd And Rf > 0 Then
'
'    'Rf = FindWindow(vbNullString, "WMI Demo - CPU Information")
'
''    On Local Error Resume Next
''    AppActivate "WMI Demo - CPU Information"
''    AppActivate "CID Run Me"
''    On Local Error GoTo 0
'
'    'Rft = Putfocus(Rf)
'    'Rft = Putfocus(Me.hwnd)
'
'End If
'
'
'
'
''FOR OLD CALLER ID PROG BREAK AND PUT RUN -------------
'TiTBait = CurProcHwnd
'
''hwnd3 = FindWindowPart("Microsoft Visual Basic [break]")
'hwnd3 = FindWindow(vbNullString, "CallerID - Microsoft Visual Basic [break]")
'If hwnd3 > 0 Then
'On Local Error Resume Next
'If Mnu_CIDBreak.Checked = True Then
'AppActivate "CallerID - Microsoft Visual Basic [break]", False
''r = Int(Now) + Timer + 0.1
''Do
''DoEvents
'CurProcHwnd = GetForegroundWindow
''If CurProcHwnd = hwnd3 Then Exit Do
''Loop Until r < Int(Now) + Timer
'
'
'If CurProcHwnd = hwnd3 Then SendKeys "{f5}", False
'End If
'End If
'
'
'
'
'
''cphwnd = GetForegroundWindow
''If GetWindow("IEFrame", vbNullString) > 0 Then
''rrt$ = GetWindowTitle(CPHwnd)
'
''            Dim cText As String
'
''            cText = Space$(255)
'            'Ret = GetClassName(Peet2, sText, 255)
'
'
''TT = GetClassName(cphwnd, cText, 255)
''ctext = Mid$(ctext, 1, InStr(ctext, Chr(0)) - 1)
''If Mid$(cText, 1, 7) = "IEFrame" Then
''Xxr = cphwnd
''If Xxr > 0 And Xxr <> Xxr8 Then
'    'ShowWindow Xxr, SW_MAXIMIZE
''    ShowWindow Xxr, SW_MINIMIZE
''    Xxr8 = Xxr
''End If
''End If
'
'
''Xxr = FindWindow(vbNullString, "Recent Pageload Activity [Matthew Lan's Home Page] - Windows Internet Explorer")
''If Xxr > 0 And Xxr <> Xxr2 Then
'    'ShowWindow Xxr, SW_MAXIMIZE
''    ShowWindow Xxr, SW_MINIMIZE
''    Xxr2 = Xxr
''End If
'
''OutLook E Cool Code
''-Find Message
'
'
'
'Dim MeRyu4 As RECT
'Dim MeRyu5 As RECT
'
'' ---- MOVE BEGIN
'
'Xxr = FindWindow(vbNullString, "Find Message")
'wert1 = GetWindowRect(Xxr, MeRyu4)
'Xz = MeRyu4.Top - MeRyu4.Bottom
'If Xxr > 0 And (Xxr <> Xxr3 Or Xz <> Xxr3t) Then
'    'ShowWindow Xxr, SW_MINIMIZE
''    ShowWindow Xxr, SW_MAXIMIZE
'
'    Xxt1 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
'    If FindWindow(vbNullString, "Magnifier") > 0 Then
'        Xxt1 = FindWindow(vbNullString, "Magnifier")
'        wert1 = GetWindowRect(Xxt1, MeRyu4)
'        If MeRyu4.Top <> 32 Then
'            Xxt1 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
'        End If
'
'    End If
'
'    xxt2 = FindWindow("Shell_TrayWnd", vbNullString)
'    wert1 = GetWindowRect(Xxt1, MeRyu4)
'    wert2 = GetWindowRect(xxt2, MeRyu5)
'
'    'MoveWindow Xxr, 0, MeRyu4.Bottom, MeRyu4.Right, MeRyu5.Top - MeRyu4.Bottom, True
'
'        SW = Screen.Width / Screen.TwipsPerPixelX
'    HH = TELEPORT
'    MoveWindow Xxr, 0, HH, SW, MeRyu5.Top - HH, True
'
'
'
'    wert1 = GetWindowRect(Xxr, MeRyu4)
'
'    Xxr3 = Xxr
'    Xxr3t = MeRyu4.Top - MeRyu4.Bottom
'End If
'
'
'
'
''Our New NoteHard Code 2009 Aug
'
'Xxr = FindWindow("Notepad2", vbNullString)
'Flag2 = 0
'Flag = 0
''Simple Rem
'
''Del the ones that are no longer there
'For R = CID_Run_Me.Lst1.ListCount - 1 To 0 Step -1
'    If GetWindowState(Val(CID_Run_Me.Lst1.List(R))) = 0 Then
'        CID_Run_Me.Lst1.RemoveItem (R)
'        Flag = Flag + 1
'    End If
'
'    'no dont do this rember just dont include in resize then if come back in same posi
'    '
'    'Del the ones that have diff window state
'    'If GetWindowState(Val(CID_Run_Me.Lst1.List(R))) <> -1 Then
'    '    CID_Run_Me.Lst1.RemoveItem (R)
'    'End If
'
'    'detect normal ones
'    If GetWindowState(Val(CID_Run_Me.Lst1.List(R))) = -1 Then Flag = Flag + 1
'
'    'if all the hwnds in our list are present on screen then cool
'    If Val(CID_Run_Me.Lst1.List(R)) = Xxr Then Flag2 = 1
'
'    'XOR this rather add then some kind hashing system detect change
'    ii = ii Xor Val(CID_Run_Me.Lst1.List(R))
'Next
'
'If Flag2 = 0 Then Xxr5 = -7
'If Xxr = 0 Then Xxr5 = -5
'If Xxr > 0 And (Xxr5 <> ii Or Flag <> OFlag Or OXxR5 <> Xxr) Then
'    'TTT = cProcesses.Convert(Xxr, Otss, cnFromhWnd Or cnToProcessID)
'    ww = FindWinPartNotePad
'    OFlag = Flag
'    OXxR5 = Xxr
'End If
'Xxr5 = ii
'
'
'
''Xxr = FindWindow(vbNullString, "VVScheduler")
''If Xxr > 0 And Xxr <> Xxr6 Then
''    ShowWindow Xxr, SW_MAXIMIZE
''    ShowWindow Xxr, SW_MINIMIZE
''    Xxr6 = Xxr
''End If
'
'
'Xxr = FindWindow(vbNullString, "CallerID - Microsoft Visual Basic [run]")
''xxrt2 = FindWindow(vbNullString, "CallerID - Microsoft Visual Basic [break]")
''xxrt3 = FindWindow(vbNullString, "CallerID - Microsoft Visual Basic [design]")
'
'If Xxr > 0 And Xxr <> Xxr7 Then
''    ShowWindow Xxr, SW_MAXIMIZE
'    ShowWindow Xxr, SW_MINIMIZE
'    Xxr7 = Xxr
'End If
'
'
''xxr8 used
'
'
'
'Xxr = FindWindow(vbNullString, "Advanced Security for Outlook")
''xxrt2 = FindWindow(vbNullString, "CallerID - Microsoft Visual Basic [break]")
''xxrt3 = FindWindow(vbNullString, "CallerID - Microsoft Visual Basic [design]")
'
'If Xxr > 0 And Xxr10 < Now Then
'Xxr10 = Now + TimeSerial(0, 0, 5)
''    ShowWindow Xxr, SW_MAXIMIZE
'    'ShowWindow Xxr, SW_MINIMIZE
'    'AppActivate ("Advanced Security for Outlook"), True
'End If
'
'Xxr = FindWindow("#32770", "Backup Progress")
'
'If Xxr > 0 And Xxr11 <> Xxr Then
'
'    ShowWindow Xxr, SW_MINIMIZE
'    Xxr11 = Xxr
'
'    Xxr = FindWindow("CMainBackupApp", vbNullString)
'    ShowWindow Xxr, SW_MINIMIZE
'
'End If
'
'Xxr = FindWindow(vbNullString, "Logon - Matt Lan")
'
'If Xxr > 0 And Xxr12 < Now Then
'
'    Xxr12 = Now + TimeSerial(0, 0, 4)
'
'    AppActivate "Logon - Matt Lan", True
'
'    SendKeys "{enter}"
'
'End If
'
'
'Xxr = FindWindow("Outlook Express Browser Class", vbNullString)
'
'If Xxr14 <> GetWindowState(Xxr) Then
'    Xxr13 = 0
'End If
'Xxr14 = GetWindowState(Xxr)
'
'If Xxr > 0 And Xxr13 <> Xxr Then
'
'Xxt1 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
'    If FindWindow(vbNullString, "Magnifier") > 0 Then
'        Xxt1 = FindWindow(vbNullString, "Magnifier")
'        wert1 = GetWindowRect(Xxt1, MeRyu4)
'        If MeRyu4.Top <> 32 Then
'            Xxt1 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
'        End If
'
'End If
'xxt2 = FindWindow("Shell_TrayWnd", vbNullString)
''wert1 = GetWindowRect(Xxt1, MeRyu4)
'wert2 = GetWindowRect(xxt2, MeRyu5)
'Xxr13 = Xxr
''If MeRyu3.Top < 40 Then
'
'SW = Screen.Width / Screen.TwipsPerPixelX
'
'HH = TELEPORT
''MoveWindow Xxr, 0, MeRyu4.Bottom, MeRyu4.Right, MeRyu5.Top - MeRyu4.Bottom, True
'MoveWindow Xxr, 0, HH, SW, MeRyu5.Top - HH, True
'
'End If
'
'
'
'
'
'
'
'Xxr = FindWindow("ATH_Note", vbNullString)
'
'If Xxr > 0 And Xxr59 <> Xxr Then
'    Xxt1 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
'
'    If FindWindow(vbNullString, "Magnifier") > 0 Then
'        Xxt1 = FindWindow(vbNullString, "Magnifier")
'        wert1 = GetWindowRect(Xxt1, MeRyu4)
'        If MeRyu4.Top <> 32 Then
'            Xxt1 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
'        End If
'    End If
'
'
'    Xxr59 = Xxr
'
'    xxt2 = FindWindow("Shell_TrayWnd", vbNullString)
'    wert1 = GetWindowRect(Xxt1, MeRyu4)
'    wert2 = GetWindowRect(xxt2, MeRyu5)
'    'MoveWindow Xxr, 0, MeRyu4.Bottom, MeRyu4.Right, MeRyu5.Top - MeRyu4.Bottom, True
'
'    SW = Screen.Width / Screen.TwipsPerPixelX
'    HH = TELEPORT
'    MoveWindow Xxr, 0, HH, SW, MeRyu5.Top - HH, True
'
'
'End If
'
'
'
'
'
'
'
'
'
'
'
'
'
'Dim OTSS As Long
'
'
''Tidal Info."
'
''Xxr = FindWindow(vbNullString, "PassLock")
'Xxr = FindWindow(vbNullString, "Extreme Lock Build: 2011"): lol = 1
'If Xxr = 0 Then Xxr = FindWindow(vbNullString, "Tidal Information..."): lol = 0
'If Xxr > 0 And Xxr <> Xxr16 Then
'    'If lol = 1 Then rt = AlwaysOnTop(Me.hwnd)
'    'If lol = 0 Then rt = NotAlwaysOnTop(Me.hwnd)
'    'me.top
'    'me.left
'    Xxr16 = Xxr
'    'win2 = FindWindow("Winamp v1.x", vbNullString)
'    'win3 = FindWindow("Winamp PE", vbNullString)
'    'rt = AlwaysOnTop(win2)
'    'rt = AlwaysOnTop(win3)
'    'ShowWindow Me.hwnd, SW_NORMAL
'    'ShowWindow win2, SW_NORMAL
'
'    TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
'    Process_ABOVE_NORMAL_PRIORITY_CLASS (OTSS)
'
'End If
''End If
'
'
''D:\VB6\VB-NT\00_Best_VB_01\ClipBoard Logger\ClipBoard Logger.exe
'Xxr = FindWindow(vbNullString, "ClipBoard Logger")
'If Xxr > 0 And Xxr <> Xxr63 Then
'    Xxr63 = Xxr
'
'    TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
'    Process_ABOVE_NORMAL_PRIORITY_CLASS (OTSS)
'End If
'
'
'Dim Xxr71 As Long
'
''MEDIA CONVERTOR
'Xxr = FindWindow(vbNullString, "Mobile Media Converter")
'If Xxr > 0 Then
'
'    CNote = "C:\Program Files\MIKSOFT\Mobile Media Converter\ffmpeg.exe"
'    ttt3 = cProcesses.Convert(CNote, Xxr, cnFromEXE Or cnToProcessID)
'    If ttt3 = True And Xxr <> Xxr71 Then
'        Xxr71 = Xxr
'        'call Process_Low_BELOW_NORM (Xxr71)
'        Call Process_Low_Low(Xxr71)
'
'        Xxr = FindWindow(vbNullString, "Conversion status")
'        If Xxr > 0 Then
'            ShowWindow Xxr, SW_FORCEMINIMIZE 'SW_MINIMIZE
'        End If
'
'        'Process_ABOVE_NORMAL_PRIORITY_CLASS (Xxr71)
'
'    End If
'End If
'
'
'
''FEED DEMON
'Xxr = FindWindow("TfFDMain.UnicodeClass", vbNullString)
'If Xxr > 0 And Xxr <> Xxr70 Then
'    Xxr70 = Xxr
'
'    TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
'    'Process_Low_BELOW_NORM (Otss)
'    Process_ABOVE_NORMAL_PRIORITY_CLASS (OTSS)
'End If
'
'
'
'Xxr = FindWindow("Digital Wave Player", vbNullString)
'If Xxr > 0 And Xxr39 <> Xxr Then
'Xxr39 = Xxr
'TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
'Call Process_Low_Low(OTSS)
'End If
'
'
'
'Xxr = FindWindow(vbNullString, "Drive File Lister")
'If Xxr > 0 And Xxr60 <> Xxr Then
'    Xxr60 = Xxr
'    TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
'    Call Process_Low_Low(OTSS)
'End If
'
'Xxr = FindWindow(vbNullString, "WORD EXAMPLE")
'If Xxr > 0 And Xxr61 <> Xxr Then
'    Xxr61 = Xxr
'    TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
'    Call Process_Low_Low(OTSS)
'End If
'
'
''
'
'
'
'Xxr = FindWindow("CMainBackupApp", vbNullString)
'If Xxr > 0 And Xxr43 <> Xxr Then
'    Xxr43 = Xxr
'    TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
'    Call Process_Low_Low(OTSS)
'End If
'
'If Idle_Timer_Proc < Now And Idle_Timer_Proc > 0 Then Xxr44 = 0: Idle_Timer_Proc = 0
'
'Xxr = FindWindow(vbNullString, "Shedular For Vicea Verse")
'If Xxr > 0 Then
'    'Xxr = FindWindowPart("ViceVersa Pro - [", 1)
'    Xxr = FindWindowPart("ViceVersa Pro - [")
'    If Xxr > 0 And Xxr44 <> Xxr Then
'        Xxr44 = Xxr
'        'TTT = cProcesses.Convert(Xxr, Otss, cnFromhWnd Or cnToProcessID)
'        'Process_Low (Otss)
'    End If
'End If
'
'
'
'TF2 = FindWindow("Winamp v1.x", vbNullString)
'MsgResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)
'If MsgResult <> 1 And TF2 <> 0 Then
'
'    Xxr = FindWindow(vbNullString, "JPEG Encoder -- Shockingly Good Photo ART -- [-:¦:-•:*''' The One '''*:•-:¦:-]")
'    If Mnu_Dont_LowerART.Checked = True Then Xxr = 0
'
'    If Xxr > 0 And Xxr53 <> Xxr Then
'        Xxr53 = Xxr
'        TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
'        Process_Low_BELOW_NORM (OTSS)
'    End If
'
'
'
'    Xxr = FindWindow(vbNullString, "XNTP Control Test")
'    If Xxr > 0 And Xxr54 <> Xxr Then
'        Xxr54 = Xxr
'        TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
'        Process_Low_BELOW_NORM (OTSS)
'    End If
'End If
'
'
'Xxr = FindWindow(vbNullString, "RS232 Reed Contact Switch Door")
'
'If Xxr > 0 And Xxr63 <> Xxr Then
'    Xxr53 = Xxr
'    TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
'    Process_ABOVE_NORMAL_PRIORITY_CLASS (OTSS)
'End If
'
'
''
'
'
'Xxr = FindWindow(vbNullString, "Web Site Update Dates")
'If Xxr > 0 And Xxr55 <> Xxr Then
'    Xxr55 = Xxr
'    If InStr(GetFileFromHwnd(Xxr), "VB6.EXE") = 0 Then
'
'    TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
'    'Process_Low_BELOW_NORM (Otss)
'    Call Process_Low_Low(OTSS)
'    End If
'End If
'
'
'
''C:\Program Files\Ahead\Nero Wave Editor\WaveEdit.exe
'Xxr = FindWindow("Afx:10000000:0", vbNullString)
'If Xxr > 0 And Xxr73 <> Xxr Then
'    Xxr73 = Xxr
'    'If InStr(GetFileFromHwnd(Xxr), "WaveEdit.exe") = 0 Then
'
'    TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
'    Process_Low_BELOW_NORM (OTSS)
'    'Call Process_Low_Low(Otss)
'    'End If
'End If
'
'
'
'
'
'
''---------------------------------------
'XXRsCNT = 0
''---------------------------------------
'
'
'xxt2 = FindWindow("Shell_TrayWnd", vbNullString)
'wert2 = GetWindowRect(xxt2, MeRyu5)
'If MeRyu5.Top <> oShell_TrayWnd Then
'oShell_TrayWnd = MeRyu5.Top
'oShell_TrayWnd_GoX = True
'End If
'
'
'ReDim XET(800)
'XXT = 0
'XXT = XXT + 1: XET(XXT) = "wndclass_desked_gsk"
'XXT = XXT + 1: XET(XXT) = "MozillaUIWindowClass"
'XXT = XXT + 1: XET(XXT) = "WMPlayerApp"
'XXT = XXT + 1: XET(XXT) = "Master Commander"
'XXT = XXT + 1: XET(XXT) = "Chrome_WidgetWin_0"
'XXT = XXT + 1: XET(XXT) = "KLS Mail Backup 1.9.7.6"
'XXT = XXT + 1: XET(XXT) = "ClipBoard Logger"
'XXT = XXT + 1: XET(XXT) = "URL Logger"
'
''----WinHTTrack Website Copier - [New Project 1]
'XXT = XXT + 1: XET(XXT) = "Afx:00400000:8:00010011:00000000:01AA09EB"
'XXT = XXT + 1: XET(XXT) = "WinHTTrack Website Copier - [New Project 1]"
''----
'
'XXT = XXT + 1: XET(XXT) = "WordNet 2.1 Browser"
'XXT = XXT + 1: XET(XXT) = "WordPadClass"
'XXT = XXT + 1: XET(XXT) = "WinRarWindow"
'XXT = XXT + 1: XET(XXT) = "WIG PRO"
'XXT = XXT + 1: XET(XXT) = "Process Tamer Configuration GUI"
'XXT = XXT + 1: XET(XXT) = "PROCEXPL"
'XXT = XXT + 1: XET(XXT) = "PowerToy Calc"
''XXT = XXT + 1: XET(XXT) = "Mobile Media Converter"
'XXT = XXT + 1: XET(XXT) = "Google Image Downloader"
'XXT = XXT + 1: XET(XXT) = "Norton PartitionMagic 8.0"
'XXT = XXT + 1: XET(XXT) = "File Monitor - Sysinternals: www.sysinternals.com"
'XXT = XXT + 1: XET(XXT) = "FastStone Image Viewer 3.9"
'
''FeedDemon 3.1.0.30
'XXT = XXT + 1: XET(XXT) = "Ghost"
'
'XXT = XXT + 1: XET(XXT) = "Autoruns"
'
'XXT = XXT + 1: XET(XXT) = "MMCMainFrame"
'XXT = XXT + 1: XET(XXT) = "Multiple Rename"
'XXT = XXT + 1: XET(XXT) = "Microsoft Encarta Encyclopedia Deluxe 2001"
'XXT = XXT + 1: XET(XXT) = "Pro Photo Tools - "
'
''XXT = XXT + 1: XET(XXT) = "Ultralingua"
'
''AutoPix - V5.2.1 [Registered] news.btinternet.com
''XXT = XXT + 1: XET(XXT) = "ThunderRT6FormDC"
''REMOVE MAJOR LOOP
'
'
'XXT = XXT + 1: XET(XXT) = "Audacity"
'XXT = XXT + 1: XET(XXT) = "Digital Wave Player"
'XXT = XXT + 1: XET(XXT) = "Dictation"
'XXT = XXT + 1: XET(XXT) = "Nero Burning ROM Enterprise Edition"
'XXT = XXT + 1: XET(XXT) = "Nero Wave Editor"
'XXT = XXT + 1: XET(XXT) = "Notepad++"
'XXT = XXT + 1: XET(XXT) = "Meda MP3 Joiner 1.2"
'XXT = XXT + 1: XET(XXT) = "WindowsForms10.Window.8"
'XXT = XXT + 1: XET(XXT) = "Measurement Converter v1.0  -   Xemsoft"
'XXT = XXT + 1: XET(XXT) = "FrontPageExplorerWindow40"
'XXT = XXT + 1: XET(XXT) = "MSWinPub"
'XXT = XXT + 1: XET(XXT) = "Personal Folders - Microsoft Outlook"
'XXT = XXT + 1: XET(XXT) = "Document1 - Microsoft Word"
'XXT = XXT + 1: XET(XXT) = "XLMAIN"
'XXT = XXT + 1: XET(XXT) = "Microsoft Network Monitor 3.3"
'XXT = XXT + 1: XET(XXT) = "Multiple Rename"
'XXT = XXT + 1: XET(XXT) = "LopeEdit Lite - [Untitled]"
'
''ICEVIEW
'XXT = XXT + 1: XET(XXT) = "ThunderRT5Form"
'
'XXT = XXT + 1: XET(XXT) = "ImageDupe"
'XXT = XXT + 1: XET(XXT) = "TFormMain.UnicodeClass"
'XXT = XXT + 1: XET(XXT) = "Google Earth"
'XXT = XXT + 1: XET(XXT) = "FileZilla"
'XXT = XXT + 1: XET(XXT) = "TMainForm"
'XXT = XXT + 1: XET(XXT) = "Poikosoft Easy CD-DA Extractor - Audio File Encoder/Decoder/Converter"
'XXT = XXT + 1: XET(XXT) = "TfrmThumbnails"
'XXT = XXT + 1: XET(XXT) = "TMainForm.UnicodeClass"
'XXT = XXT + 1: XET(XXT) = "ThunderRT6TextBox"
'XXT = XXT + 1: XET(XXT) = "Tmainform"
'XXT = XXT + 1: XET(XXT) = "COOL2000SS"
'XXT = XXT + 1: XET(XXT) = "Theme_MenuBar_Class"
'XXT = XXT + 1: XET(XXT) = "Concise Oxford English Dictionary"
'XXT = XXT + 1: XET(XXT) = "WABBrowseView"
'XXT = XXT + 1: XET(XXT) = "Agent Ransack - [Search1]"
'XXT = XXT + 1: XET(XXT) = "Alcohol 120%"
'XXT = XXT + 1: XET(XXT) = "Amazon MP3 Downloader"
'XXT = XXT + 1: XET(XXT) = "Axialis IconWorkshop 5.0"
'XXT = XXT + 1: XET(XXT) = "TreeSize Free"
'XXT = XXT + 1: XET(XXT) = "ProShow Gold - Matt"
'XXT = XXT + 1: XET(XXT) = "System Commander"
'XXT = XXT + 1: XET(XXT) = "InfoRapid Search & Replace - [Search Results]"
'XXT = XXT + 1: XET(XXT) = "WebDrive Version 7.21"
'XXT = XXT + 1: XET(XXT) = "µTorrent4823DF041B09"
'XXT = XXT + 1: XET(XXT) = "DiskDigger"
'XXT = XXT + 1: XET(XXT) = "CMainBackupApp"
'XXT = XXT + 1: XET(XXT) = "MediaPlayerClassicW"
'XXT = XXT + 1: XET(XXT) = "metapad"
'XXT = XXT + 1: XET(XXT) = "mp3DCWndClass"
'XXT = XXT + 1: XET(XXT) = "untitled - KompoZer"
'XXT = XXT + 1: XET(XXT) = "AhaView"
'XXT = XXT + 1: XET(XXT) = "tednpad"
'XXT = XXT + 1: XET(XXT) = "Hyperbolic Tessellations"
'XXT = XXT + 1: XET(XXT) = "TextviewWndClass60"
'XXT = XXT + 1: XET(XXT) = "Revo Uninstaller"
'XXT = XXT + 1: XET(XXT) = "WinMerge"
'XXT = XXT + 1: XET(XXT) = "TXnewsFrame"
'XXT = XXT + 1: XET(XXT) = "TFormMain"
'XXT = XXT + 1: XET(XXT) = "SyncToy"
'
'XXT = XXT + 1: XET(XXT) = "CabinetWClass"
'
'XXT = XXT + 1: XET(XXT) = "ExploreWClass"
'XXT = XXT + 1: XET(XXT) = "cswThumbsPlusMain"
'XXT = XXT + 1: XET(XXT) = "MozillaUIWindowClass"
'XXT = XXT + 1: XET(XXT) = "IEFrame"
''TOGETHER
'XXT = XXT + 1: XET(XXT) = "Google Earth"
'XXT = XXT + 1: XET(XXT) = "QWidget"
'
'XXT = XXT + 1: XET(XXT) = "FileZilla Main Window"
'XXT = XXT + 1: XET(XXT) = "AdobeAcrobat"
'XXT = XXT + 1: XET(XXT) = "FrontPageExplorerWindow40"
''"ViceVersa Pro"
'
'
'ReDim Preserve XET(XXT)
'
'
'
''CHK NOT WANT
'For R = 1 To UBound(XET)
'
'XClass = XET(R)
'
'XXRsCNT = XXRsCNT + 1
'
'Xxr = 0
'If Xxr = 0 Then Xxr = FindWindow(XClass, vbNullString)
'If Xxr = 0 Then Xxr = FindWindow(vbNullString, XClass)
'
''IF-Y CODE ####
'
'Dim XXrXS(100)
'
'
'If oShell_TrayWnd_GoX = True Then XXrXS(XXRsCNT) = 0
'If Xxr = 0 Then XXrXS(XXRsCNT) = 0
'
'Teleport2 = 0
'
'If FindWindow(XClass, "WinAmp MP3% MINI") Then Xxr = 0
'If FindWindow(XClass, "Volume2") Then Xxr = 0
''CLOCK
'If FindWindow(XClass, "CLOCK") Then Xxr = 0
''If FindWindow(XClass, "") Then Xxr = 0
''If FindWindow(XClass, "") Then Xxr = 0
''If FindWindow(XClass, "") Then Xxr = 0
''If FindWindow(XClass, "") Then Xxr = 0
''If FindWindow(XClass, "") Then Xxr = 0
''If FindWindow(XClass, "") Then Xxr = 0
''If FindWindow(XClass, "") Then Xxr = 0
''If FindWindow(XClass, "") Then Xxr = 0
''If FindWindow(XClass, "") Then Xxr = 0
''ThunderRT6FormDC
'
'
'
'If Xxr = 0 Then XXrXS(XXRsCNT) = 0
'
'If Xxr > 0 And XXrXS(XXRsCNT) <> Xxr Then
'    XXrXS(XXRsCNT) = Xxr
'    Teleport2 = 0
'
'    If FindWindow(XClass, "Recycle Bin") Then Teleport2 = 80
'
'    Xxt1 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
'    wert1 = GetWindowRect(Xxt1, MeRyu4)
'
'    'FUTURE SEARCHES
'    'HwndCTLTask2 = FindWindow(vbNullString, "CTLTask")
'    xxt2 = FindWindow("Shell_TrayWnd", vbNullString)
'    wert2 = GetWindowRect(xxt2, MeRyu5)
'
'    XXrXS(XXRsCNT) = Xxr
'
'    xwid = Screen.Width / Screen.TwipsPerPixelX
'
''    MoveWindow Xxr, 0, MeRyu4.Bottom, xwid, MeRyu5.Top - MeRyu4.Bottom, True
'
'    SW = Screen.Width / Screen.TwipsPerPixelX
'    HH = TELEPORT + Teleport2
''    MoveWindow Xxr, 0, hh, SW, MeRyu5.Top - hh, True
'
'
'End If
'
'Next
'
'
''-------------------------------------------------------------------------
''END FOR LOOP TO SET WINDOW SIZES ARRAY
''-------------------------------------------------------------------------
'
'
''SHOW MESSAGE
'Xxr = FindWindow("#32770", "Outlook Express")
'
'If Xxr <> 0 Then
'    wert1 = GetWindowRect(Xxr, MeRyu4)
'    lk = MeRyu4.Right - MeRyu4.Left
'    If lk <> 488 Then Xxr = 0
'End If
'
'If Xxr41 > Now And WindowVisible(Xxr) = True And OlWinVis = False Then
'    Lock5 = 1
'End If
'
''PUT BACK IF WANT
''OlWinVis = WindowVisible(Xxr)
'
'If WindowVisible(Xxr) = False Then
'    Lock5 = 0
'    'Xxr41 = Now + TimeSerial(0, 0, 2)
'End If
'
'If Xxr <> 0 And Xxr41 < Now And Lock5 = 0 And WindowVisible(Xxr) = True Then
'
'Xxr40 = Xxr
'Xxr41 = Now + TimeSerial(0, 5, 0)
'
''PUT BACK IF WANT
''WindowVisible(Xxr) = False
'
'End If
'
'
'
'
''CLOSE PROGRAM BE GENTAL CODE -----------------------------
'
''the server could not be found
'Xxr = FindWindow("#32770", "Outlook Express")
'If Xxr > 0 And Xxr23 <> Xxr Then
'    Xxr23 = Xxr
'    wert1 = GetWindowRect(Xxr, MeRyu4)
'    If MeRyu4.Right - MeRyu4.Left = 429 Then
'        If MeRyu4.Bottom - MeRyu4.Top = 172 Then
'            MsgBox "What"
'            CloseProgramHwnd (Xxr)
'        End If
'    End If
'End If
'
'
''end DO send POST----------NOXX
'Xxr = FindWindow("#32770", "Outlook Express")
'If Xxr > 0 And Xxr24 <> Xxr Then
'    Xxr24 = Xxr
'    wert1 = GetWindowRect(Xxr, MeRyu4)
'   'If MeRyu4.Right - MeRyu4.Left = 462 Then
'        If MeRyu4.Bottom - MeRyu4.Top < 500 Then
'            'CloseProgramHwnd (Xxr)
'        End If
'    'End If
'End If
'
'
'
''REBOOT CLOSE
'Xxr = FindWindow("#32770", "Outlook Express")
'If Xxr > 0 And Xxr23 <> Xxr Then
'    Xxr23 = Xxr
'    wert1 = GetWindowRect(Xxr, MeRyu4)
'    If MeRyu4.Right - MeRyu4.Left = 429 Then
'        If MeRyu4.Bottom - MeRyu4.Top = 172 Then
'            MsgBox "What"
'            CloseProgramHwnd (Xxr)
'        End If
'    End If
'End If
'
'
'
'
''Time Out
''Outlook Express was unable to switch to the newsgroup 'alt.support.schizophrenia' on the server '01 Roids 01 Warrior'
'Xxr = FindWindow("#32770", "Outlook Express")
'If Xxr > 0 And Xxr49 <> Xxr Then
'    Xxr49 = Xxr
'    wert1 = GetWindowRect(Xxr, MeRyu4)
'    xxoe1 = MeRyu4.Right - MeRyu4.Left
'    xxoe2 = MeRyu4.Bottom - MeRyu4.Top
'    If xxoe1 = 429 Or xxoe1 = 514 Then
'        If xxoe2 = 429 Or xxoe2 = 136 Then
'            CloseProgramHwnd (Xxr)
'
'        End If
'    End If
'End If
'
'
''VAR TIMER CODE FOR WEB CAM ------------------------------
'
'
'Xxr = FindWindow(vbNullString, "WebCam Motion Capture")
'If Xxr = 0 Then
'    Xxp2 = FindWindow(vbNullString, "WebCam Motion Capture In IDE")
'    Xxr = Xxp2
'End If
''If Xxr > 0 Then Stop
'If Xxr = 0 And Xxr45 <> Xxr Then
'    Label38.BackColor = 0
'End If
'If Xxr > 0 And Xxr45 <> Xxr Then
'    Label38.BackColor = QBColor(12)
'    Xxr45 = Xxr
'    Xxr46 = Now + TimeSerial(0, 1, 0)
'    Timer14Var = Now + TimeSerial(0, 4, 0)
'End If
'
'
'
'
'
'If Xxr46 > 0 And Xxr = 0 Then
'    WebCamCount = 0
'End If
'
'If Xxr46 > 0 And Xxr46 < Now And Xxr > 0 Then
'Xxr46 = 0
'TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
'If Xxp2 = 0 Then Process_Kill (OTSS)
''Timer4.Enabled = True
''If WebCamTime > Now Then WebCamCount = WebCamCount + 1
''WebCamCount = WebCamCount + 1
''If WebCamCount >= 5 Then
''Timer4.Enabled = True
''WebCamTime = Now + TimeSerial(0, 1, 10)
'End If
''If WebCamTime < Now Then WebCamCount = 0
'
'
'
'
''If FindWindow(vbNullString, "Download") > 0 Then Call dss_download
'
'
'
'
'
''SEND KEY'S CODE -----------------------------
'
'
''If CurProcHwnd <> CurProcHwnd2 Or ReRun = 1 Or FirstRun2 = False Then
''start the stuff------------
'
'
''If FindWindow(vbNullString, "Security Information") > 0 Then
'If FindWindow(vbNullString, "Security Warning") > 0 Then
'    w21 = Timer + 0.1
'    w22 = Now + TimeSerial(0, 0, 2)
'    Do
'        DoEvents
'    Loop Until w22 < Now Or w21 < Timer
'    AppActivate "Security Warning", 0
'    SendKeys "{enter}", True
'End If
'
'
''OUTLOOK YOU HAVE MSGS TO SEND DO YOU WANT
'
'Xxr = FindWindow("#32770", "Outlook Express")
'If Xxr > 0 And Xxr72 <> Xxr Then
'
'    wert1 = GetWindowRect(Xxr, MeRyu4)
'    If MeRyu4.Bottom - MeRyu4.Top < 500 Then
'    'CloseProgramHwnd (Xxr)
'    Xxr72 = Xxr
'    If GetForegroundWindow <> Xxr Then SetForegroundWindow Xxr
'        SendKeys "N", True
'
'    End If
'End If
'
'
'
'Xxr = FindWindow("#32770", "Microsoft Office Genuine Advantage")
'If Xxr > 0 And Xxr73 <> Xxr Then
'    Xxr73 = Xxr
'    wert1 = GetWindowRect(Xxr, MeRyu4)
'    xxoe1 = MeRyu4.Right - MeRyu4.Left
'    xxoe2 = MeRyu4.Bottom - MeRyu4.Top
'    If xxoe1 = 410 Then
'        If xxoe2 = 171 Then
'
'            'DONT WORK
'            'CloseProgramHwnd (Xxr)
'
'
'            If GetForegroundWindow <> Xxr Then SetForegroundWindow Xxr
'            SendKeys "%R", True
'
'        End If
'    End If
'End If
'
'
'
'Xxr = FindWindow(vbNullString, "AutoComplete Passwords")
'If Xxr > 0 And Xxr24 <> Xxr Then
'            Xxr24 = Xxr
'            TTT = cProcesses.Convert(Me.hWnd, OTSS, cnFromhWnd Or cnToProcessID)
'            On Error Resume Next
''            AppActivate (Otss), False
''            SetForegroundWindow Me.hWnd
'            TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
''            AppActivate (Otss), True
''            SetForegroundWindow Xxr
'
'            If GetForegroundWindow = Xxr Then SendKeys "%(Y)", True
'            'If Err.Number = 0 Then Xxr24 = Xxr:msgbox
'            On Error GoTo 0
'End If
'
'
'
'Xxr = FindWindow("#32770", "Microsoft Visual Basic")
'If Xxr > 0 And XXR77 < Now Then
'
'    wert1 = GetWindowRect(Xxr, MeRyu4)
'    If MeRyu4.Right - MeRyu4.Left < 500 Then
'
'        XXR77 = Xxr
'        If GetForegroundWindow <> Xxr Then SetForegroundWindow Xxr
'        'AppActivate "Microsoft Visual Basic", True
'        SendKeys "{enter}"
'
'    End If
'End If
'
'
'
'
''Xxr = FindWinPart("Enter password for", True)
'Xxr = FindWindowPart("Enter password for")
'If Xxr = 0 Then XXR78 = 0
'If Xxr > 0 And XXR78 < Now Then
'    'wert1 = GetWindowRect(Xxr, MeRyu4)
''    If MeRyu4.Right - MeRyu4.Left < 500 Then
'        XXR78 = Xxr
'        Rc1 = 0
'        Do
'        If GetForegroundWindow <> Xxr Then SetForegroundWindow Xxr
'        If GetForegroundWindow <> Xxr Then Sleep 50
'        Loop Until Rc1 > 50 Or GetForegroundWindow = Xxr
'        SendKeys "bigparticals5{enter}", True
'A = A
''    End If
'End If
'
'
'Xxr = FindWindow("#32770", "Insert disk")
'If Xxr > 0 And XXR77 < Now Then
'
'    wert1 = GetWindowRect(Xxr, MeRyu4)
'    If MeRyu4.Right - MeRyu4.Left < 500 Then
'
'        XXR75 = Xxr
'        If GetForegroundWindow <> Xxr Then SetForegroundWindow Xxr
'        'AppActivate "Microsoft Visual Basic", True
'        SendKeys "{enter}"
'
'    End If
'End If
'
'
''
'
''InfoRapid Search & Replace
'
'Xxr = FindWindow("#32770", "InfoRapid Search & Replace")
'If Xxr > 0 And Xxr51 <> Xxr Then
'    Xxr51 = Xxr 'Now + TimeSerial(0, 0, 1)
'    SetForegroundWindow Xxr
'    'TTT = cProcesses.Convert(Xxr, Otss, cnFromhWnd Or cnToProcessID)
'    'On Local Error Resume Next
'    'Err.Clear
'    'AppActivate (Otss), False
'    'If Err.Number = 0 Then SendKeys "n", True
'    SendKeys "n", True
'    'On Local Error GoTo 0
'    'SendKeys "{esc}", True
'End If
'
'
'
''Error Copying File or Folder
'If FindWindow(vbNullString, "Error Copying File or Folder") > 0 Then
'   SendKeys "{esc}", False
'End If
''ENC2001.EXE - Application Error
'If FindWindow(vbNullString, "ENC2001.EXE - Application Error") > 0 Then
'   SendKeys "{esc}", False
'End If
'
'If FindWindow(vbNullString, "SHRUIL Tooltip Controller: ENC2001.EXE - Application Error") > 0 Then
'   SendKeys "{esc}", False
'End If
'
'If FindWindow("TfmAbout", "About Image Viewer") > 0 Then
'   SendKeys "{esc}", False
'End If
'
'If FindWindow("IEFrame", vbNullString) > 0 Then
'If FindWindow(vbNullString, "Internet Explorer") > 0 Then
''   SendKeys "%y", False
'   KKQ = 1
'End If
'End If
'
'If FindWindow("IEFrame", vbNullString) > 0 Then
'If FindWindow(vbNullString, "Confirm File Replace") > 0 Then
''   SendKeys "%y", False
'   KKQ = 1
'End If
'End If
'
'
'
'
'If KKQ <> 0 Then KKQ = KKQ + 1
'KKQ = 0 'SET NOT USE ------------------------
'If KKQ > 7 Then
'hwnd3 = FindWindow("IEFrame", vbNullString)
'If hwnd3 > 0 Then
'ash$ = GetWindowTitle(hwnd3)
'Err.Clear
'On Local Error Resume Next
'AppActivate ash$, False
'If Err.Number = 0 Then
'KKQ = 0
'End If
'End If
'End If
'
'
'
'
'
'
'
'
'
'
'Xxr = FindWindow("#32770", "Outlook Express")
'hWndChild2 = FindWindowEx(Xxr, 0&, "Button", "&Do Not Send")
'Xxrg = GetForegroundWindow
''If (Xxr > 0 And Xxr34 <> Xxr And Xxrg <> Xxr35) Or hWndChild2 > 0 Then
'If (Xxr > 0 And Xxr34 <> Xxr) Or hWndChild2 > 0 Then
'
'Xxr35 = Xxrg
''file2$ = GetFileFromHwnd(Xxr)
''    If InStr(file2$, "VPN_Auto-Dialer") Or InStr(file2$, "VB6.EXE") Or InStr(file2$, "SendEmail") Then
'    If 1 = 1 Then
'        test_hwnd = FindWindowDLL(0&, 0&)
'        hWndChild2 = FindWindowEx(Xxr, 0&, "Button", "&Do Not Send")
'        If hWndChild2 > 0 Then Xxr = hWndChild2: easyx = 1
'        If hWndChild2 = 0 Then
'            Do
'                hWndChild2 = FindWindowEx(test_hwnd, 0&, "Button", "&Do Not Send")
'                easyx = 0
'                If hWndChild2 > 0 Then
'                    easyx = 1
'                    Xxr = test_hwnd
'                End If
'            test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
'            Loop Until test_hwnd = 0 Or easyx = 1
'        End If
'
'        If easyx = 1 Then
'            If GetForegroundWindow <> Xxr Then
'                rt = AlwaysOnTop(Xxr)
'            End If
'            If GetForegroundWindow <> Xxr Then SetForegroundWindow Xxr
'            If GetForegroundWindow <> Xxr Then Rft = Putfocus(Xxr)
'
'            If Err.Number = 0 And GetForegroundWindow = Xxr Then
'                SendKeys "%S", True
'                Xxr34 = Xxr
'            End If
'        End If
'        On Local Error GoTo 0
'    End If
'End If
'
'
'
'
'
'
'
''SET CHILD WINDOW SIZE -----------------------------
''IEFRAME  ------------------ SET WINDOW IN INTERNET
'
'
'Xxr = FindWindow("IEFrame", vbNullString)
''Xxr = 0
'If Xxr > 0 And Xxr22 <> Xxr Then 'And GetForegroundWindow = Xxr Then
'
'Xxt1 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
'    If FindWindow(vbNullString, "Magnifier") > 0 Then
'        Xxt1 = FindWindow(vbNullString, "Magnifier")
'        wert1 = GetWindowRect(Xxt1, MeRyu4)
'        If MeRyu4.Top <> 32 Then
'            Xxt1 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
'        End If
'    End If
'
'
'Dim MeRect As RECT
'
'Xxr22 = Xxr
'xxt2 = FindWindow("Shell_TrayWnd", vbNullString)
'wert2 = GetWindowRect(xxt2, MeRect)
'
'SW = Screen.Width / Screen.TwipsPerPixelX
'HH = TELEPORT
'MoveWindow Xxr, 0, HH, SW, MeRect.Top - HH, True
'
'
'Xxr2 = FindWindowEx(Xxr, 0&, "Links Explorer", vbNullString)
'wert1 = GetWindowRect(Xxr2, MeRect)
'MoveWindow Xxr2, MeRect.Left, MeRect.Top - 57, 300, MeRect.Bottom - MeRect.Top, True
'
''Xxr = FindWindow("IEFrame", vbNullString)
''Xxr2 = FindWindowEx(Xxr, 0&, "Frame Tab", vbNullString)
''Xxr2 = FindWindowEx(Xxr2, 0&, "TabWindowClass", vbNullString)
''Xxr2 = FindWindowEx(Xxr2, 0&, "Shell DocObject View", vbNullString)
''Xxr2 = FindWindowEx(Xxr2, 0&, "Internet Explorer_Server", vbNullString)
''wert1 = GetWindowRect(Xxr2, MeRyu4)
''MoveWindow Xxr2, 0, 0, (MeRyu4.Right - MeRyu4.Left) - 200, MeRyu4.Bottom - MeRyu4.Top, True
'
'End If
'
'
'
''MAX WINDOWS WITHIN VB FOR ------------------------------------
''VbCode window To Max on Detect
'Dim mWnd As Long
'FindCursor Nx, Ny
'mWnd = WindowFromPoint(Nx, Ny)
'
'XXrXSt = GetActiveWindowClass(mWnd) 'Mwnd will get reset here if none
'If GetWindowTitle(mWnd) = "Immediate" Then mWnd = 0
'If GetWindowTitle(mWnd) = "Locals" Then mWnd = 0
'If GetWindowTitle(mWnd) = "Watches" Then mWnd = 0
'
'If XXrXSt = "VbaWindow" Or XXrXSt = "DesignerWindow" And mWnd > 0 And Xxr47 <> mWnd And mWnd > 0 Then
'    Xxr47 = mWnd
'    ShowWindow mWnd, SW_MAXIMIZE
'End If
'
'
'
'
'
''Shell Load 'S ---------------------------
'
'Xxr = FindWindow("CabinetWClass", "Recycle Bin")
'
'If Xxr > 0 And Xxr48 <> Xxr Then
'    Xxr48 = Xxr
'    Shell App.Path + "\DelReCycle\DelRecycle.exe", vbNormalFocus
'End If
'
'
'
'
''MOVE CREEPING UP WINDOWS BACK DOWN ---------------------
'
'
'Dim MeRyu3 As RECT
'
'If FindWindow(vbNullString, "Microsoft Excel") > 0 Then
'Dim HwndMicE As Long
'HwndMicE = FindWindow(vbNullString, "Microsoft Excel")
'
'wert = GetWindowRect(HwndMicE, MeRyu3)
'If MeRyu3.Top < 40 Then
'
''MoveWindow HwndMicE, MeRyu3.Left, 131, MeRyu3.Right - MeRyu3.Left, MeRyu3.Bottom - MeRyu3.Top, True
''only if its not already maxmized correct this
'
'End If
'End If
'
'
'
'
''Process explorer Properties
'CurProcHwnd = GetForegroundWindow
'Xxr = FindWindow("#32770", vbNullString)
'
'If Xxr > 0 And Xxr31 <> Xxr Then
'    ttr$ = GetWindowTitle(Xxr)
'    If InStr(ttr$, "Properties") > 0 Then
'        TT = GetParentTitle(Xxr)
'        If InStr(TT, "Process Explorer - Sysinternals") > 0 Then
'
'            wert1 = GetWindowRect(Xxr, MeRyu4)
'            'If MeRyu4.Top <> 263 And MeRyu4.Top < 40 Then
'                centery1 = ((Screen.Height / 2) / Screen.TwipsPerPixelX) - (MeRyu4.Bottom - MeRyu4.Top) / 2
'                centerx1 = (Screen.Width / 2) / Screen.TwipsPerPixelX - (MeRyu4.Right - MeRyu4.Left) / 2
'                MoveWindow Xxr, centerx1, centery1, MeRyu4.Right - MeRyu4.Left, MeRyu4.Top - MeRyu4.Bottom, True
'            'End If
'            Xxr31 = Xxr
'        End If
'    End If
'End If
'
'
'
'
'End Sub

Public Function MkDirNested(sFullPath As String) As Boolean
    On Error GoTo ErrHandler
    Dim LnNextSlash As Integer
    Dim LnStartPos As Integer
    Dim LsCurDir As String
    ' Set first char
    LnStartPos = 1
    ' validates path syntax
    If (Right$(sFullPath, 1) <> "\") Then sFullPath = (sFullPath & "\")
    Do
        LnNextSlash = InStr(LnStartPos, sFullPath, "\")
        If (LnNextSlash >= LnStartPos) Then
            LsCurDir = Left$(sFullPath, LnNextSlash)
            If (Not DirExist(LsCurDir)) Then
               ' Create the dir
               MkDir LsCurDir
            End If
             ' Check if it's the last char and exit if true
             LnStartPos = (LnNextSlash + 1)
             If (LnStartPos >= Len(sFullPath)) Then Exit Do
        End If
    Loop
    MkDirNested = True
    Exit Function
ErrHandler:
    MkDirNested = False
End Function



Public Function GetWindowState(ByVal lngHwnd As Long) As Integer
    If IsWindow(lngHwnd) = 1 Then
        GetWindowState = -1
    If IsIconic(lngHwnd) <> 0 Then
        GetWindowState = vbMinimized
    ElseIf IsZoomed(lngHwnd) <> 0 Then
        GetWindowState = vbMaximized
    End If
    End If
End Function



Public Function GetFileFromHwnd(lngHwnd) As String

'MsgBox getfilefromhwnd(Me.hwnd)

Dim lngProcess&, hProcess&, bla&, C&
Dim strFile As String
Dim X

strFile = String$(256, 0)
X = GetWindowThreadProcessId(lngHwnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
X = EnumProcessModules(hProcess, bla, 4&, C)
C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
GetFileFromHwnd = Left(strFile, C)

End Function


Public Function GetFileFromProc(lngProcess) As String

'MsgBox getfilefromhwnd(Me.hwnd)
'Dim lngProcess&, hProcess&, bla&, C&
Dim hProcess&, bla&, C&
Dim strFile As String
Dim X

strFile = String$(256, 0)
'x = GetWindowThreadProcessId(lngHwnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
X = EnumProcessModules(hProcess, bla, 4&, C)
C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
GetFileFromProc = Left(strFile, C)

End Function


Public Sub Process_Kill(P_ID As Long)
    '// Kill the wanted process
    
    Dim hProcess As Long
    Dim lExitCode As Long
    'NORMAL_PRIORITY_CLASS
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, P_ID)
    'If hProcess = 0 Then Call Err_Dll(Err.LastDllError, "OpenProcess failed", sLocation, "Kill_Process")
    
    If GetExitCodeProcess(hProcess, lExitCode) = False Then
    'Call Err_Dll(Err.LastDllError, "GetExitCodeProcess failed", sLocation, "Kill_Process")
    End If
    If TerminateProcess(hProcess, lExitCode) = False Then
    'Call Err_Dll(Err.LastDllError, "TerminateProcess failed", sLocation, "Kill_Process")
    End If
    If CloseHandle(hProcess) = False Then
    'Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Kill_Process")
    End If
End Sub

Public Sub Process_Idle(P_ID As Long)
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
    SetPriorityClass hProcess, NORMAL_PRIORITY_CLASS
    
End Sub

Public Sub Process_Low(P_ID As Long)
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
    If Idle_Timer_Proc < Now Then
        'SetPriorityClass hProcess, BELOW_NORMAL_PRIORITY_CLASS
        SetPriorityClass hProcess, IDLE_PRIORITY_CLASS
    Else
        SetPriorityClass hProcess, IDLE_PRIORITY_CLASS
    End If
End Sub

Public Sub Process_Low_Low(P_ID As Long)
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
    
    SetPriorityClass hProcess, IDLE_PRIORITY_CLASS

End Sub

Public Sub Process_Low_BELOW_NORM(P_ID As Long)
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
    
    SetPriorityClass hProcess, BELOW_NORMAL_PRIORITY_CLASS

End Sub

Public Sub Process_ABOVE_NORMAL_PRIORITY_CLASS(P_ID As Long)
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, P_ID)
    
    SetPriorityClass hProcess, ABOVE_NORMAL_PRIORITY_CLASS

End Sub

   

Public Sub closewindow(P_ID As Long)
    '// Kill the wanted process
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    '  hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, P_ID): If hProcess = 0 Then Call Err_Dll(Err.LastDllError, "OpenProcess failed", sLocation, "Kill_Process")
    
    '  If GetExitCodeProcess(hProcess, lExitCode) = False Then Call Err_Dll(Err.LastDllError, "GetExitCodeProcess failed", sLocation, "Kill_Process")
    '  If TerminateProcess(hProcess, lExitCode) = False Then Call Err_Dll(Err.LastDllError, "TerminateProcess failed", sLocation, "Kill_Process")
    
    ' If CloseHandle(hProcess) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Kill_Process")
End Sub

Public Sub Thread_Suspend(T_ID As Long)
    
    On Error GoTo VB_Error
    
        Dim hThread As Long
        Dim lSuspendCount As Long
        
        hThread = OpenThread(THREAD_SUSPEND_RESUME, False, T_ID)
        If hThread = 0 Then
            'Err_Dll Err.LastDllError, "OpenThread failed", sLocation, "Suspend_Thread"  'Call Error_API(Err.LastDllError, sLocation & "\cmdSuspend_Click", "OpenThread")
        End If
        
        lSuspendCount = SuspendThread(hThread)
        If lSuspendCount = -1 Then
            'Err_Dll Err.LastDllError, "Suspend failed", sLocation, "Suspend_Thread"
        End If
        If CloseHandle(hThread) = False Then
            'Err_Dll Err.LastDllError, "CloseThread failed", sLocation, "Suspend_Thread"
        End If
        
    Exit Sub

VB_Error:
    'Err_Vb Err.Number, Err.Description, sLocation, "Thread_Suspend"
    Resume Next
End Sub

Public Sub Thread_Resume(T_ID As Long)
    On Error GoTo VB_Error
    
        Dim hThread As Long
        Dim lSuspendCount As Long
        
        hThread = OpenThread(THREAD_SUSPEND_RESUME, False, T_ID)
        If hThread = 0 Then
            'Err_Dll Err.LastDllError, "OpenThread failed", sLocation, "Suspend_Thread"  'Call Error_API(Err.LastDllError, sLocation & "\cmdSuspend_Click", "OpenThread")
        End If
    
        lSuspendCount = ResumeThread(hThread)
        If lSuspendCount = -1 Then
            'Err_Dll Err.LastDllError, "OpenThread failed", sLocation, "Suspend_Thread"
        End If
        If CloseHandle(hThread) = False Then
            'Err_Dll Err.LastDllError, "CloseThread failed", sLocation, "Suspend_Thread"
        End If
        
    Exit Sub
VB_Error:
    'Err_Vb Err.Number, Err.Description, sLocation, "Thread_Resume"
    Resume Next
End Sub


Property Let WindowVisible(hWnd As Long, New_Value As Boolean)

    Call ShowWindow(hWnd, IIf(New_Value, SW_SHOW, SW_HIDE))

End Property

Property Get WindowVisible(hWnd As Long) As Boolean

    WindowVisible = (IsWindowVisible(hWnd) > 0)
  
End Property

Public Sub CloseProgram(ByVal Caption As String)
    
    Dim Hndle
    
    Hndle = FindWindow(vbNullString, Caption)
    If Hndle = 0 Then Exit Sub
    SendMessage Hndle, WM_CLOSE, 0&, 0&

End Sub

Public Sub CloseProgramHwnd(ByVal Handle As Long)

 If Handle = 0 Then Exit Sub
 SendMessage Handle, WM_CLOSE, 0&, 0&

End Sub




Private Sub chkOnTop_Click()
    
    If chkOnTop.Value = 1 Then
        ' Put window on top of all others
        
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        SaveSetting "EliteSpy+", "Settings", "AlwaysOnTop", "1"
        MNU_ME_ON_TOP.Caption = "[__ ME_ON_TOP_=_YES __]"
        Label60.BackColor = Label49.BackColor '49 58_59

        'FEEDBACK LOOP STACK OVERFLOW IF SETTTING HERE
        '---------------------------------------------
        'Me.chkOnTop.Value = 1
        'MNU_ME_ON_TOP_Click

    Else
        ' Remove window from top
        SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        SaveSetting "EliteSpy+", "Settings", "AlwaysOnTop", "0"
        MNU_ME_ON_TOP.Caption = "[__ ME_ON_TOP_=_NOT __]"
        'Label60.BackColor = Label59.BackColor '49 58_59
        Label60.BackColor = Label49.BackColor '49 58_59
        
        'Me.chkOnTop.Value = 0
    End If
    
    Timer_ALWAYS_ON_TOP_TO_START_WITH_ER.Enabled = False
    Label60.Caption = "Me on Top @ Loader EXE Timer 20 Second NOT DONE"
    'Label60.BackColor = Label59.BackColor '49 58_59

End Sub

'////////////////////////////////////////////////////////////////////
'//// BUTTON EVENTS
'////////////////////////////////////////////////////////////////////
Private Sub cmdMoveMax_Click()
    
    ' Make Sure there is a space between the Eye _ i _ I _ iResult I_Result I Borg Kim
    ' --------------------------------------------------------------------------------

    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd

    Dim Rect_Get As RECT
    Dim HX, HY, HW, HH, SSHW, SSHH
    Dim i_Result
    Dim ScreenSize As RECT
    
    i_Result = GetWindowRect(GetDesktopWindow(), ScreenSize)
    
    i_Result = GetWindowRect(txthWnd.Text, Rect_Get)
    
    HX = (Rect_Get.Right - Rect_Get.Left)
    HY = 40
    HW = Rect_Get.Right - Rect_Get.Left
    HH = Rect_Get.Bottom - Rect_Get.Top
    SSHW = ScreenSize.Right - ScreenSize.Left
    SSHH = ScreenSize.Bottom - ScreenSize.Top
    
    HX = 10
    HY = 10
    HW = SSHW - 20
    HH = SSHH - 240
    
    ShowWindow txtMhWnd.Text, SW_NORMAL
    
    MoveWindow txthWnd.Text, HX, HY, HW, HH, True
    lHwnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdMaximize_Click()
    ' Maximize window
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd
    
    ShowWindow txtMhWnd.Text, SW_MAXIMIZE

    lHwnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdMinimize_Click()
    
    ' Minimize window
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd
    
    ShowWindow txtMhWnd.Text, SW_MINIMIZE
    'ShowWindow txtMhWnd.Text, SW_NORMAL
    
    'ShowWindow txtParent.Text, SW_HIDE
    
    
    'ShowWindow txtParent.Text, SW_SHOW
    'ShowWindow txtParent.Text, SW_MINIMIZE
    'ShowWindow txtParent.Text, SW_NORMAL
    'ShowWindow txtParent.Text, SW_RESTORE
    'txtParent

    lHwnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdNormal_Click()
    ' Show window
    Beep
    
    Dim lhWndParentX
    
    If Val(txtMhWnd.Text) = 0 Then txtMhWnd.Text = Me.hWnd
   
    ' txtMhWnd.Text = GetParent(Val(txtMhWnd.Text))
    ' txtMhWnd.Text = GetParentHwnd(Val(txtMhWnd.Text))
    ' lhWndParentX = GetParentHwnd(Val(txtMhWnd.Text))
    ' txtMhWnd.Text = GetParentHwnd(Val(txtMhWnd.Text))

    ' txtMhWnd.Text = GetAncestor(Val(txtMhWnd.Text), GA_ROOT)

    ShowWindow txtMhWnd.Text, SW_NORMAL
    
    lHwnd_Function_Button_Set_MIN_MAX = Val(txtMhWnd.Text)
    Call ChunkCodeOnMouse
    
End Sub

Private Sub cmdFlash_Click()
    ' Flash window
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd
    
    FlashWindow txtMhWnd.Text, 3
    
    lHwnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdEnable_Click()
    ' Enable window
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd
    
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd
    EnableWindow txtMhWnd.Text, 1

    lHwnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdDisable_Click()
    ' Disable window
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd
    
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd: Exit Sub
    EnableWindow txtMhWnd.Text, 0

    lHwnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub

Private Sub cmdSetTitle_Click()
    Dim sTitle As String
    Dim i_Result
    
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd
    
    ' Ask user for new window title
    sTitle = InputBox("Enter new window title:", "EliteSpy +")
    ' Set new window title
    
    If sTitle = "" Then
        i_Result = MsgBox("You Going to Set Title Nothing Are Your Sure", vbYesNo + vbMsgBoxSetForeground)
        
        If i_Result = vbNo Then Beep: Exit Sub
        If i_Result <> vbYes Then Beep: Exit Sub
    End If
    Beep
    SetWindowText txtMhWnd.Text, sTitle
    
    lHwnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdOnTop_Click()
    Beep
    
    If txtMhWnd.Text = "" Then
        Call MNU_ME_ON_TOP_Click
        If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd
        Exit Sub
    End If
    
    ' Put window on top of all others
    SetWindowPos txtMhWnd.Text, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

    lHwnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdNotOnTop_Click()
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd

    ' Remove window from top
    SetWindowPos txtMhWnd.Text, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

    lHwnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub

Private Sub cmdHide_Click()
    ' Hide window
    Beep
    
    If txtMhWnd.Text = "" Then
        txtMhWnd.Text = Me.hWnd
        Call ChunkCodeOnMouse
        MsgBox "DON'T HIDE OWN WINDOW", vbMsgBoxSetForeground
        Exit Sub
    End If
    
    ShowWindow txtMhWnd.Text, SW_HIDE

    lHwnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdShow_Click()
    ' Show window
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd
    ShowWindow txtMhWnd.Text, SW_SHOW

    lHwnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdTerminate_Click()
    ' Close window
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd: Exit Sub
    SendMessage txtMhWnd.Text, WM_CLOSE, 0, 0
End Sub


Sub ChunkCodeOnMouse()
        
        Dim tRC2 As RECT
        Dim tPA As POINTAPI
        Dim lHwnd As Long
        Dim sTitle As String * 255
        Dim sClass As String * 255
        Dim sParentTitle As String * 255
        Dim sParentClass As String * 255
        Dim sParentTitleX As String * 255
        Dim sParentClassX As String * 255
        Dim lhWndParent As Long
        Dim lhWndParentX As Long
        Dim sStyle As String
        Dim lRetVal As Long
        Dim O_lhWndParent As Long
        Dim Set_Collect_More_Info
        Dim PID_MARK
        Dim Success_Result
        
        ' Get window rectCLEAR
        
        If From_ListView = False Then
            If hWnd_From_ListView = 0 And lHwnd_Function_Button_Set_MIN_MAX = 0 Then
                ' Get cursor position
                GetCursorPos tPA
                ' Get window handle from point
                lHwnd = WindowFromPoint(tPA.X, tPA.Y)
                ' Get window caption
            End If
        End If
        
        ' ---------------------------------------------------
        ' USED BY
        ' WHEN CLICK A PROCESS IN BOX THEN POLULATE THIS AREA
        ' RATHER THAN HOVER OVER WITH MOUSE
        ' Private Sub lstProcess_3_SORTER_ListView_Click()
        ' ---------------------------------------------------
        If From_ListView = True Then
            Set_Collect_More_Info = True
            If hWnd_From_ListView > 0 Then
                lHwnd = hWnd_From_ListView
                TxtPID.Text = Val(PROCESS_TO_KILLER_PID)
                hWnd_From_ListView = 0
                Set_Collect_More_Info = False
            End If
        End If
        If From_ListView = True And hWnd_From_ListView = 0 Then
            Set_Collect_More_Info = False
        End If
        
        
        If lHwnd_Function_Button_Set_MIN_MAX > 0 Then
            ' txthWnd.Text = lHwnd
            lHwnd = lHwnd_Function_Button_Set_MIN_MAX
            lHwnd_Function_Button_Set_MIN_MAX = 0
            ' Set_Collect_More_Info = false
        End If
        
        
        
        
        GetWindowRect lHwnd, tRC2
        
        lRetVal = GetWindowText(lHwnd, sTitle, 255)
        ' Get window class name
        lRetVal = GetClassName(lHwnd, sClass, 255)
        ' Get window style
        sStyle = GetWindowStyle(lHwnd)
        ' Get window parent
        
        O_lhWndParent = lHwnd
        lhWndParent = GetParent(lHwnd)
        If lhWndParent = 0 Then lhWndParent = O_lhWndParent
        lhWndParentX = GetParentHwnd(lHwnd)
        
        
        If Set_Collect_More_Info = True Then
            Success_Result = cProcesses.Get_PID_From_HWND(lhWndParentX, PID_MARK)
            
            TxtPID.Text = PID_MARK
        End If


        
        
        ' Get parent window caption
        lRetVal = GetWindowText(lhWndParent, sParentTitle, 255)
        ' Get parent window class name
        lRetVal = GetClassName(lhWndParent, sParentClass, 255)
        
        ' Get parentX window caption
        lRetVal = GetWindowText(lhWndParentX, sParentTitleX, 255)
        ' Get parentX window class name
        lRetVal = GetClassName(lhWndParentX, sParentClassX, 255)
        
        ' Set values to textboxes
        If From_ListView = False Then
        
            TxtEXE.Text = GetFileFromHwnd(lHwnd)
            
            
            PROCESS_TO_KILLER = Mid(TxtEXE.Text, InStrRev(TxtEXE.Text, "\") + 1)
            PROCESS_TO_KILLER_PID = TxtPID.Text
        
            Call MOUSE_HOOVER_SLECTION_CLICKER
        
        End If
        
        From_ListView = False
        
        
        txthWnd.Text = lHwnd
        txthWndHX.Text = Hex(lHwnd)
        txtTitle.Text = sTitle
        txtClass.Text = sClass
        txtStyle.Text = sStyle
        txtRect.Text = "(" & tRC2.Left & ", " & tRC2.Top & ") - (" & tRC2.Right & ", " & tRC2.Bottom & ")"
        
        'lhWndParent = GetParent(lhWndParent)
        'lhWndParentHX = Hex(GetParent(lhWndParent))
        If lHwnd <> lhWndParent Then
            txtParent.Text = lhWndParent
            txtParentHX.Text = Hex(lhWndParent)
            txtParentText.Text = sParentTitle
            txtParentClass.Text = sParentClass
        Else
            txtParent.Text = lhWndParent
            txtParentHX.Text = Hex(lhWndParent)
            txtParentText.Text = ""
            txtParentClass.Text = ""
        End If
        
        If lHwnd <> lhWndParentX And lhWndParentX <> lhWndParent Then
            txtParentX.Text = lhWndParentX
            txtParentTextX.Text = sParentTitleX
            txtParentClassX.Text = sParentClassX
        Else
            txtParentX.Text = lhWndParentX
            txtParentTextX.Text = ""
            txtParentClassX.Text = ""
        End If
        
        
        txtMhWnd.Text = lHwnd
        
        If TxtEXE.Text <> OLD_TxtEXE_Text Then
            Call TxtEXE_CLICK
        End If
        OLD_TxtEXE_Text = TxtEXE.Text


End Sub

'////////////////////////////////////////////////////////////////////
'//// PRIVATE FUNCTIONS
'////////////////////////////////////////////////////////////////////
' Get window styles
Private Function GetWindowStyle(ByVal lHwnd As Long) As String
    Dim lStyle As Long
        
    ' Get window styles
    lStyle = GetWindowLong(lHwnd, GWL_STYLE)
    
    ' Get window styles
    If lStyle And WS_BORDER Then GetWindowStyle = GetWindowStyle & "WS_BORDER "
    If lStyle And WS_CAPTION Then GetWindowStyle = GetWindowStyle & "WS_CAPTION "
    If lStyle And WS_CHILD Then GetWindowStyle = GetWindowStyle & "WS_CHILD "
    If lStyle And WS_CLIPCHILDREN Then GetWindowStyle = GetWindowStyle & "WS_CLIPCHILDREN "
    If lStyle And WS_CLIPSIBLINGS Then GetWindowStyle = GetWindowStyle & "WS_CLIPSIBLINGS "
    If lStyle And WS_DLGFRAME Then GetWindowStyle = GetWindowStyle & "WS_DLGFRAME "
    If lStyle And WS_GROUP Then GetWindowStyle = GetWindowStyle & "WS_GROUP "
    If lStyle And WS_HSCROLL Then GetWindowStyle = GetWindowStyle & "WS_HSCROLL "
    If lStyle And WS_MAXIMIZEBOX Then GetWindowStyle = GetWindowStyle & "WS_MAXIMIZEBOX "
    If lStyle And WS_MINIMIZEBOX Then GetWindowStyle = GetWindowStyle & "WS_MINIMIZEBOX "
    If lStyle And WS_SYSMENU Then GetWindowStyle = GetWindowStyle & "WS_SYSMENU "
    If lStyle And WS_POPUPWINDOW Then GetWindowStyle = GetWindowStyle & "WS_POPUPWINDOW "
    If lStyle And WS_TABSTOP Then GetWindowStyle = GetWindowStyle & "WS_TABSTOP "
    If lStyle And WS_THICKFRAME Then GetWindowStyle = GetWindowStyle & "WS_THICKFRAME "
    If lStyle And WS_VISIBLE Then GetWindowStyle = GetWindowStyle & "WS_VISIBLE "
    If lStyle And WS_VSCROLL Then GetWindowStyle = GetWindowStyle & "WS_VSCROLL "

End Function

Function GetParentHwnd(ByVal ReturnParent As Long) As String
   Dim i As Long
   Dim j As Long
   Dim k As Long
   i = ReturnParent
   If ReturnParent Then
      Do While i <> 0
         k = j
         j = i
         i = GetParent(i)
      Loop
    i = j
    End If
    GetParentHwnd = i
End Function




Private Sub lstProcess_2__Click()

Exit Sub

PROCESS_TO_KILLER = lstProcess_2_.List(lstProcess_2_.ListIndex)

'PROCESS_TO_KILLER = Replace(UCase(PROCESS_TO_KILLER), ".EXE", "")
'If InStr(UCase(lstProcess_2_.List(lstProcess_2_.ListIndex)), ".EXE") = 0 Then
'    Exit Sub
'End If

Label22 = "TASKKILLER __ " + Mid(PROCESS_TO_KILLER, 8)
'Label30 = ""

Beep
'Me.WindowState = vbMinimized

'Shell "CMD /C START """" /REALTIME /MAX ""TASKLIST"", vbMaximizedFocus"
'Shell "CMD /C START """" /REALTIME /MAX ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT_03_PROCESS_KILLER.BAT"" " + PROCESS_TO_KILLER, vbMaximizedFocus

'Timer_EnumProcess = True

End Sub


Private Sub lstProcess_2_ListView_Click()
'lstProcess_2_ListView_HERE

'lstProcess_2_ListView.Height
'-----------------------------------------
'if debugger in here it will crash the ide
'so try an make it good
'-----------------------------------------

'Do While Timer_EnumProcess.Enabled = True
'    EnumProcess_COUNTER = 0
'    DoEvents
'    'SLEEP 20
'Loop

'Stop

PROCESS_TO_KILLER = lstProcess_2_ListView.ListItems(lstProcess_2_ListView.SelectedItem.Index).SubItems(1)
PROCESS_TO_KILLER_PID = lstProcess_2_ListView.ListItems(lstProcess_2_ListView.SelectedItem.Index)

'PROCESS_TO_KILLER_PID = PROCESS_TO_KILLER_PID
'Stop
'PROCESS_TO_KILLER_PID = lstProcess_2_ListView.ListItems.Item(1).Text
'PROCESS_TO_KILLER = lstProcess_2_ListView.ListItems.Item(1).SubItems(1)

'Stop

'PROCESS_TO_KILLER_PID = lstProcess_2_ListView1.ListItems.Item(lstProcess_2_ListView1.SelectedItem.Index)


'PROCESS_TO_KILLER = lstProcess_2_ListView.SelectedItem.SubItems(1)


'DA1$ = ScanPath.ListView1.ListItems.Item(ListView1.SelectedItem.Index).SubItems(1)
'DB1$ = ScanPath.ListView1.ListItems.Item(ListView1.SelectedItem.Index)
'Dim PROCESS_TO_KILLER_PID_2
'For R_NEXT = 1 To 8
'    If Mid(PROCESS_TO_KILLER_PID, 1, 1) = "0" Then
'        PROCESS_TO_KILLER_PID_2 = Mid(PROCESS_TO_KILLER_PID, 2)
'    End If
'Next
'
'If PROCESS_TO_KILLER_PID = "" Then
''    PROCESS_TO_KILLER_PID = lstProcess_2_ListView.SelectedItem.Text
''    Exit Sub
'End If



Call LISTVIEW_CLICKER

End Sub

Private Sub lstProcess_3_SORTER_ListView_Click()

'-----------------------------------------
'if debugger in here it will crash the ide
'so try an make it good
'-----------------------------------------

'Do While Timer_EnumProcess.Enabled = True
'    EnumProcess_COUNTER = 0
'    DoEvents
'    'SLEEP 20
'Loop

'Stop

PROCESS_TO_KILLER = lstProcess_3_SORTER_ListView.ListItems(lstProcess_3_SORTER_ListView.SelectedItem.Index).SubItems(1)
PROCESS_TO_KILLER_PID = lstProcess_3_SORTER_ListView.ListItems(lstProcess_3_SORTER_ListView.SelectedItem.Index)

'PROCESS_TO_KILLER_PID = PROCESS_TO_KILLER_PID
'Stop
'PROCESS_TO_KILLER_PID = lstProcess_2_ListView.ListItems.Item(1).Text
'PROCESS_TO_KILLER = lstProcess_2_ListView.ListItems.Item(1).SubItems(1)

'Stop

'PROCESS_TO_KILLER_PID = lstProcess_2_ListView1.ListItems.Item(lstProcess_2_ListView1.SelectedItem.Index)


'PROCESS_TO_KILLER = lstProcess_2_ListView.SelectedItem.SubItems(1)


'DA1$ = ScanPath.ListView1.ListItems.Item(ListView1.SelectedItem.Index).SubItems(1)
'DB1$ = ScanPath.ListView1.ListItems.Item(ListView1.SelectedItem.Index)
'Dim PROCESS_TO_KILLER_PID_2
'For R_NEXT = 1 To 8
'    If Mid(PROCESS_TO_KILLER_PID, 1, 1) = "0" Then
'        PROCESS_TO_KILLER_PID_2 = Mid(PROCESS_TO_KILLER_PID, 2)
'    End If
'Next
'
'If PROCESS_TO_KILLER_PID = "" Then
''    PROCESS_TO_KILLER_PID = lstProcess_2_ListView.SelectedItem.Text
''    Exit Sub
'End If

Call LISTVIEW_CLICKER

End Sub

Sub LISTVIEW_CLICKER()

Dim Success_Result
Dim Hwnd_Var
Dim PID_MARK
Dim PID_INPUT As Long
Dim NAME_EXE As String
' ----------------------------------------------------------------
' CAP 1ST LETTER
' ----------------------------------------------------------------
Mid(PROCESS_TO_KILLER, 1, 1) = UCase(Mid(PROCESS_TO_KILLER, 1, 1))

Label22.Caption = "TASKKILLER BY PID NUMBER __ # " + PROCESS_TO_KILLER_PID
Label30.Caption = "TASKKILLER NAME ___________ " + PROCESS_TO_KILLER

TxtEXE.Text = GetFileFromProc(Val(PROCESS_TO_KILLER_PID))
TxtPID.Text = Val(PROCESS_TO_KILLER_PID)
If TxtEXE.Text = "" Then
    PID_INPUT = Val(PROCESS_TO_KILLER_PID)
    Hwnd_Var = cProcesses.GetHwnd_From_PID(PID_INPUT)
    If Hwnd_Var > 0 Then
        Success_Result = cProcesses.Get_PID_From_HWND(Hwnd_Var, PID_MARK)
        TxtEXE.Text = GetFileFromProc(Val(PID_MARK))
    End If
    If TxtEXE.Text = "" Then
        'i = cProcesses.GetEXEID(PID_INPUT, NAME_EXE)
        'TxtEXE.Text = NAME_EXE
        TxtEXE.Text = cProcesses.GetEXE_Path_From_ProcessID(PID_INPUT)
    End If
    If TxtEXE.Text = "" Then
        TxtEXE.Text = PROCESS_TO_KILLER
    End If
    
End If
Dim hWnd_Parent As Long
PID_INPUT = Val(PROCESS_TO_KILLER_PID)

hWnd_Parent = cProcesses.GetHwnd_From_PID(PID_INPUT)
' txthWnd.Text = hWnd_Parent
If hWnd_Parent > 0 Then
    hWnd_From_ListView = hWnd_Parent
End If

From_ListView = True
Call ChunkCodeOnMouse

End Sub

Sub MOUSE_HOOVER_SLECTION_CLICKER()

'MOUSE_HOOVER_SLECTION_CLICKER
'PROCESS_TO_KILLER
'PROCESS_TO_KILLER_PID

' ----------------------------------------------------------------
' CAP 1ST LETTER
' ----------------------------------------------------------------
Mid(PROCESS_TO_KILLER, 1, 1) = UCase(Mid(PROCESS_TO_KILLER, 1, 1))

Label22.Caption = "TASKKILLER BY PID NUMBER __ # " + PROCESS_TO_KILLER_PID
Label30.Caption = "TASKKILLER NAME ___________ " + PROCESS_TO_KILLER

'TxtEXE.Text = GetFileFromProc(Val(PROCESS_TO_KILLER_PID))
'TxtPID.Text = Val(PROCESS_TO_KILLER_PID)

' Dim hWnd_Parent As Long
' Dim PID_INPUT As Long
' PID_INPUT = Val(PROCESS_TO_KILLER_PID)

' hWnd_Parent = cProcesses.GetHwnd_From_PID(PID_INPUT)
' txthWnd.Text = hWnd_Parent

' hWnd_From_ListView = hWnd_Parent
' Call ChunkCodeOnMouse

End Sub



Private Sub lstProcess_Click()
Exit Sub

'PROCESS_TO_KILLER = lstProcess.List(lstProcess.ListIndex)
'PROCESS_TO_KILLER = Replace(UCase(PROCESS_TO_KILLER), ".EXE", "")
'If InStr(UCase(lstProcess.List(lstProcess.ListIndex)), ".EXE") = 0 Then
'    Exit Sub
'End If
'
'Me.WindowState = vbMinimized
'
''Shell "CMD /C START """" /REALTIME /MAX ""TASKLIST"", vbMaximizedFocus"
''Shell "CMD /C START """" /REALTIME /MAX ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT_03_PROCESS_KILLER.BAT"" " + PROCESS_TO_KILLER, vbMaximizedFocus
'
'
'Timer_EnumProcess = True

End Sub

Private Sub lstProcess_3_SORTER_Click()

PROCESS_TO_KILLER = lstProcess_3_SORTER.List(lstProcess_3_SORTER.ListIndex)
'PROCESS_TO_KILLER = Replace(UCase(PROCESS_TO_KILLER), ".EXE", "")
'If InStr(UCase(lstProcess_3_SORTER.List(lstProcess_3_SORTER.ListIndex)), ".EXE") = 0 Then
'    Exit Sub
'End If

'Label22 = "TASKKILLER __ " + Mid(PROCESS_TO_KILLER, 8)
'Label30 = ""

Beep
'Me.WindowState = vbMinimized
'Label22 = "TASKKILLER " + PROCESS_TO_KILLER + " __ SELECT AN OPTION ABOVE"

'Shell "CMD /C START """" /REALTIME /MAX ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT_03_PROCESS_KILLER.BAT"" " + PROCESS_TO_KILLER, vbMaximizedFocus

'Timer_EnumProcess = True

End Sub





Private Sub Timer_GET_KEY_ASYNC_STATE_Timer()

If IsIDE = True Then Timer_GET_KEY_ASYNC_STATE.Interval = 1000

Dim tPA As POINTAPI, lHwnd As Long, O_lhWndParent, lhWndParent, lhWndParentX
GetCursorPos tPA
lHwnd = WindowFromPoint(tPA.X, tPA.Y)
O_lhWndParent = lHwnd
lhWndParent = GetParent(lHwnd)
If lhWndParent = 0 Then lhWndParent = O_lhWndParent
lhWndParentX = GetParentHwnd(lHwnd)

If GetAsyncKeyState(27) < 0 Then
    If IsIDE = True Then
        If GetForegroundWindow = Me.hWnd Or lhWndParent = Me.hWnd Or lhWndParentX = Me.hWnd Then
        Unload Me
        End If
    Else
        If GetForegroundWindow = Me.hWnd Or lhWndParent = Me.hWnd Or lhWndParentX = Me.hWnd Then
            Me.WindowState = vbMinimized
        End If
    End If
End If

End Sub

'Private Sub MNU_BRing_Front_Click()
'' MNU_BRing_Front
'' Mnu_Menu_Item_Count
'' Mnu_Form_Count
''Call FindWinPartFront
''MNU_BRing_Front.Caption = "Bring All Front -- " + Format(Now, "DD-MMM-YYYY HH:MM:SS")
'
''GO_SUB = False
''For i = 1 To 255
''    bdf = GetAsyncKeyState(i)
''    If bdf < 0 Then
''        If i = 39 Then i = 0
''        If i = 116 Then i = 0 'F5 TEST ISIDE
''        If i = 16 Then GO_SUB = True 'LEFT SHIFT KEY
''        If GO_SUB = True Then
''            'Call KeyOrMouse
''            MNU_BRing_Front.Caption = "Bring All Front"
''            Exit Sub
''        End If
''    End If
''Next
'
'
''TEST=
''ExplorerGone = True
''ExplorerGone_TEST = True
'Timer_EXPLORER_GONE.Enabled = True
''getkeyasync
'End Sub

'Private Sub MNU_ME_ON_TOP_Click()
'
'    Beep
'    ' Put window on top of all others
'    'SetWindowPos txtMhWnd.Text, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'
'    Timer_ALWAYS_ON_TOP_TO_START_WITH_ER.Enabled = False
'    Label60.BackColor = Label49.BackColor '49 58_59
'    Label60.Caption = "Me on Top @ Loader EXE Timer 20 Second NOT _ DONE"
'    MNU_ME_ON_TOP.Caption = "[__ ME_ON_TOP_=_YES __]"
'    Me.chkOnTop.Value = 1
'End Sub




Private Sub txtClass_CLICK()

On Error GoTo ENDER
Clipboard.Clear
Clipboard.SetText txtClass

Exit Sub
ENDER:
DoEvents
Resume

End Sub

Private Sub TxtEXE_CLICK()

If TIMER2_TIMER_BEGAN > 0 Then Exit Sub

On Error GoTo ENDER
Clipboard.Clear
Clipboard.SetText TxtEXE

Exit Sub
ENDER:
DoEvents
Resume

End Sub


