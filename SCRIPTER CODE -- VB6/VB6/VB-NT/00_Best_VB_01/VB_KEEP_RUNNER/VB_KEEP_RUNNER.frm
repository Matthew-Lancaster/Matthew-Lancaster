VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   Caption         =   "VB_KEEP_RUNNER"
   ClientHeight    =   10116
   ClientLeft      =   48
   ClientTop       =   912
   ClientWidth     =   12864
   Icon            =   "VB_KEEP_RUNNER.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10116
   ScaleWidth      =   12864
   Begin VB.Timer FOREGROUND_WINDOW_CHANGE_DELAY_1_EXTRA_TO_DO 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11112
      Top             =   2040
   End
   Begin VB.FileListBox File_GOODSYNC 
      Height          =   264
      Left            =   11505
      TabIndex        =   136
      Top             =   3804
      Visible         =   0   'False
      Width           =   1236
   End
   Begin VB.Timer ONE_MILLISECOND_2_Timer 
      Interval        =   1
      Left            =   9990
      Top             =   1584
   End
   Begin VB.TextBox Text_OS_INSTALL_DATE_02 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4332
      Locked          =   -1  'True
      TabIndex        =   131
      Text            =   "OS INSTALL DATE"
      Top             =   4968
      Width           =   1905
   End
   Begin VB.TextBox Text_SYSTEM_START_TIME_02 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4332
      Locked          =   -1  'True
      TabIndex        =   130
      Text            =   "SYSTEM START"
      Top             =   4716
      Width           =   1905
   End
   Begin VB.TextBox Text_OS_INSTALL_DATE_01 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   516
      Locked          =   -1  'True
      TabIndex        =   129
      Text            =   "OS INSTALL DATE"
      Top             =   4968
      Width           =   3780
   End
   Begin VB.TextBox Text_SYSTEM_START_TIME_01 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   516
      Locked          =   -1  'True
      TabIndex        =   128
      Text            =   "SYSTEM START TIME"
      Top             =   4716
      Width           =   3780
   End
   Begin VB.PictureBox picCrossHair 
      BorderStyle     =   0  'None
      Height          =   384
      Left            =   5664
      Picture         =   "VB_KEEP_RUNNER.frx":0E42
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   115
      Top             =   336
      Width           =   384
   End
   Begin VB.Timer TIMER_OS_REBOOT 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   15180
      Top             =   1584
   End
   Begin VB.Timer TIMER_TO_RESIZE 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   14835
      Top             =   1584
   End
   Begin VB.Timer Timer_MOUSE_CORD 
      Interval        =   10
      Left            =   14490
      Top             =   1536
   End
   Begin VB.ListBox List_SORT_FOR_AHK_LIMITER 
      Height          =   240
      Left            =   9285
      Sorted          =   -1  'True
      TabIndex        =   105
      Top             =   2940
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Timer Timer_Pause_Update 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   14145
      Top             =   1584
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13815
      Top             =   1584
   End
   Begin VB.Timer Timer_SHOW_THE_TIME 
      Interval        =   1000
      Left            =   13470
      Top             =   1584
   End
   Begin VB.Timer Timer_GET_KEY_ASYNC_STATE 
      Interval        =   5
      Left            =   13125
      Top             =   1536
   End
   Begin VB.Timer Timer_FOREGROUND_WINDOW_CHANGE_02 
      Interval        =   10
      Left            =   12780
      Top             =   1584
   End
   Begin VB.Timer Timer_EnumProcess 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12420
      Top             =   1584
   End
   Begin VB.ListBox lstProcess_3_SORTER 
      Height          =   276
      IntegralHeight  =   0   'False
      Left            =   9285
      Sorted          =   -1  'True
      TabIndex        =   76
      Top             =   5196
      Visible         =   0   'False
      Width           =   1728
   End
   Begin VB.ListBox lstProcess_2_ 
      Height          =   324
      IntegralHeight  =   0   'False
      Left            =   9285
      TabIndex        =   75
      Top             =   4848
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
      Left            =   11730
      TabIndex        =   72
      Top             =   5508
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txthWnd 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1764
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   1044
      Width           =   1044
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1764
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   1296
      Width           =   1995
   End
   Begin VB.TextBox txtClass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1764
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   1560
      Width           =   1995
   End
   Begin VB.TextBox txtParent 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1764
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   2316
      Width           =   1044
   End
   Begin VB.TextBox txtStyle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1764
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   1812
      Width           =   1995
   End
   Begin VB.TextBox txtParentText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1764
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   2580
      Width           =   1995
   End
   Begin VB.TextBox txtParentClass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1764
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   2832
      Width           =   1995
   End
   Begin VB.TextBox txtRect 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1764
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   2064
      Width           =   1995
   End
   Begin VB.TextBox txtParentClassX 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1764
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   3600
      Width           =   1995
   End
   Begin VB.TextBox txtParentTextX 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1764
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3336
      Width           =   1995
   End
   Begin VB.TextBox txtParentX 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1764
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   3084
      Width           =   1995
   End
   Begin VB.TextBox TxtEXE 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1710
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   24
      Width           =   7452
   End
   Begin VB.Timer Timer_ALWAYS_ON_TOP_TO_START_WITH_ER 
      Interval        =   1000
      Left            =   12075
      Top             =   1584
   End
   Begin VB.TextBox txthWndHX 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2832
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1044
      Width           =   936
   End
   Begin VB.TextBox txtParentHX 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2832
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2316
      Width           =   930
   End
   Begin VB.TextBox TxtPID 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1764
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   792
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
      Left            =   16050
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5664
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
      Left            =   16050
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   5412
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
      Left            =   16050
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   6168
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
      Left            =   16050
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5928
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
      Left            =   16050
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   6420
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
      Left            =   16050
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   6672
      Visible         =   0   'False
      Width           =   2688
   End
   Begin VB.CheckBox chkOnTop 
      Appearance      =   0  'Flat
      Caption         =   "ME On Top _ Registry"
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
      Height          =   240
      Left            =   3816
      TabIndex        =   7
      ToolTipText     =   "STORE IN REISTRY FOR START UP OVERRIDE TIMER IF DO"
      Top             =   3216
      Width           =   2430
   End
   Begin VB.Timer Timer_1_SECOND 
      Interval        =   1000
      Left            =   9285
      Top             =   1584
   End
   Begin VB.FileListBox File3 
      Height          =   264
      Left            =   9270
      TabIndex        =   6
      Top             =   4164
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.FileListBox File2 
      Height          =   264
      Left            =   9270
      TabIndex        =   5
      Top             =   3900
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Timer Timer_DIR_FOR_C_DRIVE_ROOT_VICE_VERSA_VBSCRIPT 
      Interval        =   1000
      Left            =   11715
      Top             =   1584
   End
   Begin VB.Timer Timer_VB_PROJECT_CHECKDATE 
      Interval        =   1000
      Left            =   11370
      Top             =   1584
   End
   Begin VB.FileListBox File1 
      Height          =   264
      Left            =   9270
      TabIndex        =   4
      Top             =   3636
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   11505
      TabIndex        =   3
      Top             =   4284
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer TIMER_DO_IN_1_SECOND 
      Interval        =   1000
      Left            =   11025
      Top             =   1584
   End
   Begin VB.Timer ONE_MILLISECOND_Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9645
      Top             =   1584
   End
   Begin VB.ListBox lstProcess 
      Height          =   240
      Left            =   9285
      TabIndex        =   2
      Top             =   3384
      Visible         =   0   'False
      Width           =   1896
   End
   Begin MSComctlLib.ListView lstProcess_2_ListView 
      Height          =   570
      Left            =   9285
      TabIndex        =   0
      Top             =   870
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   995
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Timer Timer_VB_MAXIMIZE 
      Interval        =   10
      Left            =   10350
      Top             =   1584
   End
   Begin VB.Timer Timer_FOREGROUND_WINDOW_CHANGE 
      Interval        =   1
      Left            =   10695
      Top             =   1548
   End
   Begin MSComctlLib.ListView lstProcess_3_SORTER_ListView 
      Height          =   570
      Left            =   11805
      TabIndex        =   1
      Top             =   870
      Width           =   2550
      _ExtentX        =   4509
      _ExtentY        =   995
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView_CPU_INFO 
      Height          =   1488
      Left            =   516
      TabIndex        =   132
      Top             =   5220
      Width           =   5724
      _ExtentX        =   10097
      _ExtentY        =   2625
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView_VB_MODIFIED_ERROR 
      Height          =   1836
      Left            =   492
      TabIndex        =   140
      Top             =   9036
      Visible         =   0   'False
      Width           =   11736
      _ExtentX        =   20680
      _ExtentY        =   3239
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label_DE_DUPE_EXPLORER 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00DFFFFF&
      Caption         =   "DE-DUPE EXPLORER"
      Height          =   216
      Left            =   7512
      TabIndex        =   155
      Top             =   3132
      Width           =   1644
   End
   Begin VB.Label Label_CLOSE_EXPLORER 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00DFFFFF&
      Caption         =   "CLOSE EXPLORER"
      Height          =   216
      Left            =   7680
      TabIndex        =   154
      Top             =   2880
      Width           =   1476
   End
   Begin VB.Label Label55 
      BackColor       =   &H00DFFFFF&
      Caption         =   " MAX G-SYNC ALL"
      Height          =   216
      Left            =   7800
      TabIndex        =   153
      Top             =   2184
      Width           =   1368
   End
   Begin VB.Label Label24 
      BackColor       =   &H00DFFFFF&
      Caption         =   " END G-SYNC ALL"
      Height          =   216
      Left            =   7800
      TabIndex        =   152
      Top             =   2424
      Width           =   1356
   End
   Begin VB.Label Label_CHROME_PAGE_AUTO_ON 
      BackColor       =   &H00DFFFFF&
      Caption         =   " CHROME PAGE"
      Height          =   216
      Left            =   6276
      TabIndex        =   151
      Top             =   4548
      Width           =   1284
   End
   Begin VB.Label Label_CLOSE_OUTLOOK 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00DFFFFF&
      Caption         =   "CLOSE OUTLOOK"
      Height          =   216
      Left            =   7764
      TabIndex        =   150
      Top             =   3588
      Width           =   1404
   End
   Begin VB.Label Label_KILL_AND_RUN_ANOTHER_AUTOHOTKEY 
      Alignment       =   2  'Center
      BackColor       =   &H00DFFFFF&
      Caption         =   "KILL && RUN ANOTHER AHK"
      Height          =   216
      Left            =   6288
      TabIndex        =   149
      Top             =   1020
      Width           =   2160
   End
   Begin VB.Label Label_MAXIMIZE_CLIPBOARD_LOGGER 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00DFFFFF&
      Caption         =   "MAX CLIPPER"
      Height          =   216
      Left            =   8076
      TabIndex        =   148
      Top             =   1728
      Width           =   1092
   End
   Begin VB.Label Label_CLOSE_GOODSYNC2GO 
      BackColor       =   &H00DFFFFF&
      Caption         =   " END GS 2 GO"
      Height          =   216
      Left            =   6552
      TabIndex        =   147
      Top             =   2868
      Width           =   1092
   End
   Begin VB.Label Label_MAXIMIZE_GOODSYNC2GO 
      BackColor       =   &H00DFFFFF&
      Caption         =   " MAX GS 2 GO"
      Height          =   216
      Left            =   6576
      TabIndex        =   146
      Top             =   2412
      Width           =   1068
   End
   Begin VB.Label Label_MAXIMIZE_ELITE_SPY 
      BackColor       =   &H00DFFFFF&
      Caption         =   " MAX ELITE SPY"
      Height          =   216
      Left            =   6276
      TabIndex        =   145
      Top             =   1728
      Width           =   1296
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00DFFFFF&
      Caption         =   "NORMAL"
      Height          =   216
      Left            =   8328
      TabIndex        =   144
      Top             =   540
      Width           =   840
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "HWND"
      Height          =   216
      Left            =   6276
      TabIndex        =   143
      Top             =   540
      Width           =   576
   End
   Begin VB.Label Label_HWND_MAXIMIZE 
      Alignment       =   2  'Center
      BackColor       =   &H00DFFFFF&
      Caption         =   "MAXIMIZE"
      Height          =   216
      Left            =   7488
      TabIndex        =   142
      Top             =   540
      Width           =   828
   End
   Begin VB.Label Label_VB_MODIFIED_ACTIVE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   7272
      TabIndex        =   141
      ToolTipText     =   "HIIT ME SHOW NETWORK SYNC-ER"
      Top             =   6420
      Width           =   1032
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "VB Modify"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   6276
      TabIndex        =   139
      Top             =   6420
      Width           =   972
   End
   Begin VB.Label Label_VB_MODIFIED_TIME 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "GS HR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   8316
      TabIndex        =   138
      ToolTipText     =   "PUSH ME TO SHOW LISTVIEW INFO"
      Top             =   6420
      Width           =   840
   End
   Begin VB.Label Label_GOODSYNC_04_HOUR 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "GS HR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   5376
      TabIndex        =   137
      Top             =   3852
      Width           =   876
   End
   Begin VB.Label Label_GOODSYNC_02_HOUR 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "GS HR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   4476
      TabIndex        =   135
      Top             =   3852
      Width           =   888
   End
   Begin VB.Label Label_GOODSYNC_01 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "GOODSYNC LOGG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   516
      TabIndex        =   134
      Top             =   3852
      Width           =   3936
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00E2FEEA&
      Caption         =   "OPERATE ON REMOTE COMPUTER"
      Height          =   240
      Left            =   6276
      TabIndex        =   133
      Top             =   5040
      Width           =   2880
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      Caption         =   "HWND COUNT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6276
      TabIndex        =   127
      Top             =   5292
      Width           =   2880
   End
   Begin VB.Label Label_KILL_CMD_AND_AHK 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00DFFFFF&
      Caption         =   "KILL CMD* && AH&K*"
      Height          =   216
      Left            =   7764
      TabIndex        =   126
      Top             =   3840
      Width           =   1392
   End
   Begin VB.Label Label_KILL_CMD 
      BackColor       =   &H00DFFFFF&
      Caption         =   " KILL &CMD*"
      Height          =   216
      Left            =   6276
      TabIndex        =   125
      Top             =   3600
      Width           =   876
   End
   Begin VB.Label Label_8M 
      Alignment       =   2  'Center
      BackColor       =   &H00E2FEEA&
      Caption         =   "8M"
      Height          =   240
      Left            =   8640
      TabIndex        =   124
      Top             =   4788
      Width           =   516
   End
   Begin VB.Label Label_7G 
      Alignment       =   2  'Center
      BackColor       =   &H00E2FEEA&
      Caption         =   "7G"
      Height          =   240
      Left            =   8250
      TabIndex        =   123
      Top             =   4785
      Width           =   375
   End
   Begin VB.Label Label_5P 
      Alignment       =   2  'Center
      BackColor       =   &H00E2FEEA&
      Caption         =   "5P"
      Height          =   240
      Left            =   7845
      TabIndex        =   122
      Top             =   4785
      Width           =   390
   End
   Begin VB.Label Label_4G 
      Alignment       =   2  'Center
      BackColor       =   &H00E2FEEA&
      Caption         =   "4G"
      Height          =   240
      Left            =   7455
      TabIndex        =   121
      Top             =   4785
      Width           =   375
   End
   Begin VB.Label Label_3L 
      Alignment       =   2  'Center
      BackColor       =   &H00E2FEEA&
      Caption         =   "3L"
      Height          =   240
      Left            =   7065
      TabIndex        =   120
      Top             =   4785
      Width           =   375
   End
   Begin VB.Label Label_2E 
      Alignment       =   2  'Center
      BackColor       =   &H00E2FEEA&
      Caption         =   "2E"
      Height          =   240
      Left            =   6660
      TabIndex        =   119
      Top             =   4785
      Width           =   390
   End
   Begin VB.Label Label_1X 
      Alignment       =   2  'Center
      BackColor       =   &H00E2FEEA&
      Caption         =   "1X"
      Height          =   240
      Left            =   6276
      TabIndex        =   118
      Top             =   4788
      Width           =   372
   End
   Begin VB.Label Label_KILL_CHROME 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00DFFFFF&
      Caption         =   "KILL CHROME"
      Height          =   216
      Left            =   8052
      TabIndex        =   117
      Top             =   4548
      Width           =   1104
   End
   Begin VB.Label Label_RUN_AUTOHOTKEY_SET_NETWORK 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "RUN AUTOHOTKEY NETWORK"
      ForeColor       =   &H00C0FFFF&
      Height          =   216
      Left            =   6276
      TabIndex        =   116
      Top             =   1272
      Width           =   2568
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   5196
      Picture         =   "VB_KEEP_RUNNER.frx":170C
      Top             =   336
      Width           =   384
   End
   Begin VB.Image imgCursor 
      Height          =   432
      Left            =   5196
      MouseIcon       =   "VB_KEEP_RUNNER.frx":1B4E
      Top             =   336
      Width           =   600
   End
   Begin VB.Label LABEL_KILL_NOT_RESPOND_FORCE 
      BackColor       =   &H00DFFFFF&
      Caption         =   " TASKKILL NOT RESPOND FORCE"
      Height          =   216
      Left            =   6276
      TabIndex        =   114
      Top             =   4308
      Width           =   2616
   End
   Begin VB.Label LABEL_KILL_NOT_RESPOND 
      BackColor       =   &H00DFFFFF&
      Caption         =   " TASKKILL NOT RESPOND"
      Height          =   216
      Left            =   6276
      TabIndex        =   113
      Top             =   4080
      Width           =   2040
   End
   Begin VB.Label Label_TASK_KILLER_CMD 
      BackColor       =   &H00DFFFFF&
      Caption         =   " TASKKILL CMD*"
      Height          =   216
      Left            =   6276
      TabIndex        =   112
      Top             =   3840
      Width           =   1284
   End
   Begin VB.Label Label_CLOSE_HWND 
      Alignment       =   2  'Center
      BackColor       =   &H00DFFFFF&
      Caption         =   "CLOSE"
      Height          =   216
      Left            =   6864
      TabIndex        =   111
      Top             =   540
      Width           =   612
   End
   Begin VB.Label Lab_KILL_EXPLORER 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00DFFFFF&
      Caption         =   "KILL EXPLORER"
      Height          =   192
      Left            =   7896
      TabIndex        =   110
      Top             =   2664
      Width           =   1260
   End
   Begin VB.Label Label_KILL_HUBIC 
      BackColor       =   &H00DFFFFF&
      Caption         =   " KILL HUBIC"
      Height          =   192
      Left            =   6276
      TabIndex        =   109
      Top             =   3384
      Width           =   948
   End
   Begin VB.Label Label_GOODSYNC_COLLECTION_SCRIPT_RUN 
      BackColor       =   &H00DFFFFF&
      Caption         =   " RUN GOODSYNC SCRIPT SET"
      Height          =   216
      Left            =   6276
      TabIndex        =   108
      Top             =   1392
      Width           =   2412
   End
   Begin VB.Label Label_CLOSE_GOODSYNC 
      BackColor       =   &H00DFFFFF&
      Caption         =   " END GOODSYNC"
      Height          =   216
      Left            =   6276
      TabIndex        =   107
      Top             =   2640
      Width           =   1368
   End
   Begin VB.Label Command_Screen_Shot_Auto_ClipBoard_er 
      Caption         =   "Screen Shot Auto ClipBoard_er when Spy_er && Archive Mode _OFF_ Hitt Button Here to Change"
      Height          =   405
      Left            =   9285
      TabIndex        =   106
      Top             =   5910
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.Label Label54 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "INTERNET CONNECTED"
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
      Height          =   468
      Left            =   6276
      TabIndex        =   104
      Top             =   5928
      Width           =   2880
   End
   Begin VB.Label Label_MINIMIZE_GOODSYNC 
      BackColor       =   &H00DFFFFF&
      Caption         =   " MINIMIZE GOODSYNC"
      Height          =   216
      Left            =   6276
      TabIndex        =   103
      Top             =   1956
      Width           =   1740
   End
   Begin VB.Label Label_MAXIMIZE_GOODSYNC 
      BackColor       =   &H00DFFFFF&
      Caption         =   " MAX GOODSYNC"
      Height          =   216
      Left            =   6276
      TabIndex        =   102
      Top             =   2184
      Width           =   1380
   End
   Begin VB.Label Label_RUN_AUTOHOTKEY_SET 
      Alignment       =   2  'Center
      BackColor       =   &H00DFFFFF&
      Caption         =   "RUN AHK"
      Height          =   216
      Left            =   8232
      TabIndex        =   101
      Top             =   792
      Width           =   936
   End
   Begin VB.Label Label_KILL_AUTOHOTKEY 
      Alignment       =   2  'Center
      BackColor       =   &H00DFFFFF&
      Caption         =   "KILL &AUTOHOTKEY.EXE"
      Height          =   216
      Left            =   6288
      TabIndex        =   100
      Top             =   792
      Width           =   1920
   End
   Begin VB.Label Label64 
      Alignment       =   2  'Center
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
      Height          =   348
      Left            =   4968
      TabIndex        =   99
      Top             =   1044
      Width           =   1272
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TASKKILLER /F /IM * /T"
      Height          =   240
      Left            =   13305
      TabIndex        =   98
      Top             =   2505
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label Label17 
      Caption         =   "TASKKILLER /IM * /T"
      Height          =   240
      Left            =   13305
      TabIndex        =   97
      Top             =   2790
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label Label18 
      Caption         =   "TASKKILLER /IM *"
      Height          =   240
      Left            =   13305
      TabIndex        =   96
      Top             =   3330
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFC0FF&
      Caption         =   "TASKKILLER /F /IM /T"
      Height          =   240
      Left            =   15090
      TabIndex        =   95
      Top             =   2505
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Label Label20 
      Caption         =   "TASKKILLER /IM /T"
      Height          =   240
      Left            =   15090
      TabIndex        =   94
      Top             =   2790
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label21 
      Caption         =   "TASKKILLER /IM"
      Height          =   240
      Left            =   15090
      TabIndex        =   93
      Top             =   3060
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label23 
      Caption         =   "HITT TO KILL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   516
      TabIndex        =   92
      Top             =   4416
      Width           =   5736
   End
   Begin VB.Label Label25 
      Caption         =   "TASKLIST GO"
      Height          =   240
      Left            =   13305
      TabIndex        =   91
      Top             =   3915
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Caption         =   "TASKKILLER /F /IM"
      Height          =   240
      Left            =   15090
      TabIndex        =   90
      Top             =   3330
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0FFC0&
      Caption         =   "TASKKILLER /F /IM"
      Height          =   240
      Left            =   16980
      TabIndex        =   89
      Top             =   3060
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFC0FF&
      Caption         =   "TASKKILLER /F /IM /T"
      Height          =   240
      Left            =   16980
      TabIndex        =   88
      Top             =   2790
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TASKKILLER /F /IM * /T"
      Height          =   240
      Left            =   16965
      TabIndex        =   87
      Top             =   2505
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "KILLER With /F FORCE"
      Height          =   240
      Left            =   16965
      TabIndex        =   86
      Top             =   2190
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Without WildCard Butter"
      Height          =   240
      Left            =   15090
      TabIndex        =   85
      Top             =   2190
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "WILD CARD ****"
      Height          =   240
      Left            =   13305
      TabIndex        =   84
      Top             =   2190
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label Label34 
      Caption         =   "TASKKILLER BY HWND HANDLE _ POST MESSENGER"
      Height          =   240
      Left            =   13305
      TabIndex        =   83
      Top             =   4170
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.Label Label43 
      Caption         =   "TASKKILLER BY HWND HANDLE _ PROCESS KILL FORCE"
      Height          =   240
      Left            =   13305
      TabIndex        =   82
      Top             =   4425
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B7FFB5&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   13320
      X2              =   18660
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Label Label56 
      Alignment       =   2  'Center
      Caption         =   "TASKKILLER PID_NUMERIC OPTION_INSTEAD"
      Height          =   240
      Left            =   15105
      TabIndex        =   81
      Top             =   3645
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   13320
      X2              =   18660
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label57 
      BackColor       =   &H00DFFFFF&
      Caption         =   "TASKKILLER COMMAND LINE EXECUTE STATUS"
      Height          =   240
      Left            =   516
      TabIndex        =   80
      Top             =   4152
      Width           =   5736
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00B7FFB5&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   13320
      X2              =   18660
      Y1              =   4695
      Y2              =   4695
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00B7FFB5&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   13320
      X2              =   18660
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label62 
      Caption         =   "TASKKILLER /F /IM *"
      Height          =   240
      Left            =   13305
      TabIndex        =   79
      Top             =   3060
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label Label_KILL_WSCRIPT 
      BackColor       =   &H00DFFFFF&
      Caption         =   " KILL WSCRIPT"
      Height          =   216
      Left            =   6276
      TabIndex        =   78
      Top             =   3132
      Width           =   1212
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "TASKKILLER /F /IM * /T"
      Height          =   216
      Left            =   6276
      TabIndex        =   77
      Top             =   288
      Width           =   2892
   End
   Begin VB.Label lblCordi 
      Alignment       =   2  'Center
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
      Height          =   300
      Left            =   6276
      TabIndex        =   74
      Top             =   5616
      Width           =   2880
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
      Height          =   345
      Left            =   9285
      TabIndex        =   73
      Top             =   5520
      Visible         =   0   'False
      Width           =   2610
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
      Height          =   276
      Left            =   3816
      TabIndex        =   71
      Top             =   3492
      Width           =   2436
   End
   Begin VB.Label Label_FORM_BACK_COLOUR 
      BackColor       =   &H00005959&
      Caption         =   "Label_FORM BACK_Form Ground COLOUR"
      Height          =   390
      Left            =   9285
      TabIndex        =   70
      Top             =   4455
      Visible         =   0   'False
      Width           =   1890
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0C0&
      Caption         =   "C1"
      Height          =   240
      Left            =   9315
      TabIndex        =   69
      Top             =   2100
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label45 
      BackColor       =   &H00FFC0FF&
      Caption         =   "C2"
      Height          =   240
      Left            =   9555
      TabIndex        =   68
      Top             =   2100
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label49 
      BackColor       =   &H00C0FFC0&
      Caption         =   "C3"
      Height          =   240
      Left            =   9810
      TabIndex        =   67
      Top             =   2100
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label47 
      BackColor       =   &H0080FFFF&
      Caption         =   "C4"
      Height          =   240
      Left            =   10035
      TabIndex        =   66
      Top             =   2100
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label58 
      Caption         =   "C5"
      Height          =   240
      Left            =   10305
      TabIndex        =   65
      Top             =   2100
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label59 
      BackColor       =   &H00DFFFFF&
      Caption         =   "C7"
      Height          =   240
      Left            =   10545
      TabIndex        =   64
      Top             =   2100
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label61 
      BackColor       =   &H0068F9CA&
      Caption         =   "C8"
      Height          =   240
      Left            =   10785
      TabIndex        =   63
      Top             =   2100
      Visible         =   0   'False
      Width           =   210
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
      Height          =   375
      Left            =   9285
      TabIndex        =   61
      Top             =   60
      Width           =   2430
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
      Height          =   375
      Left            =   12885
      TabIndex        =   60
      Top             =   60
      Width           =   2445
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
      Height          =   345
      Left            =   12090
      TabIndex        =   59
      Top             =   90
      Width           =   210
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
      Height          =   315
      Left            =   9285
      TabIndex        =   58
      ToolTipText     =   "Pause Update for 20 Second"
      Top             =   480
      Width           =   6045
   End
   Begin VB.Label lblHwnd 
      Caption         =   "hWnd:"
      Height          =   240
      Left            =   516
      TabIndex        =   57
      Top             =   1044
      Width           =   1236
   End
   Begin VB.Label lblTitle 
      Caption         =   "Title:"
      Height          =   240
      Left            =   516
      TabIndex        =   56
      Top             =   1296
      Width           =   1236
   End
   Begin VB.Label Label1 
      Caption         =   $"VB_KEEP_RUNNER.frx":2418
      Height          =   1776
      Left            =   3816
      TabIndex        =   55
      Top             =   1416
      Width           =   2424
   End
   Begin VB.Label lblClass 
      Caption         =   "Class:"
      Height          =   240
      Left            =   516
      TabIndex        =   54
      Top             =   1560
      Width           =   1236
   End
   Begin VB.Label lblParent 
      Caption         =   "Parent:"
      Height          =   240
      Left            =   516
      TabIndex        =   53
      Top             =   2316
      Width           =   1236
   End
   Begin VB.Label Label2 
      Caption         =   "Style:"
      Height          =   240
      Left            =   516
      TabIndex        =   52
      Top             =   1812
      Width           =   1236
   End
   Begin VB.Label lblParentText 
      Caption         =   "Parent Text:"
      Height          =   240
      Left            =   516
      TabIndex        =   51
      Top             =   2580
      Width           =   1236
   End
   Begin VB.Label Label5 
      Caption         =   "Parent Class:"
      Height          =   240
      Left            =   516
      TabIndex        =   50
      Top             =   2832
      Width           =   1236
   End
   Begin VB.Label Label6 
      Caption         =   "Rectangle:"
      Height          =   240
      Left            =   516
      TabIndex        =   49
      Top             =   2064
      Width           =   1236
   End
   Begin VB.Label Label7 
      Caption         =   "Parent X Class:"
      Height          =   240
      Left            =   516
      TabIndex        =   48
      Top             =   3600
      Width           =   1236
   End
   Begin VB.Label Label8 
      Caption         =   "Parent X Text:"
      Height          =   240
      Left            =   516
      TabIndex        =   47
      Top             =   3336
      Width           =   1236
   End
   Begin VB.Label Label9 
      Caption         =   "Parent X:"
      Height          =   240
      Left            =   516
      TabIndex        =   46
      Top             =   3084
      Width           =   1236
   End
   Begin VB.Label Label_Goto_File_Name 
      BackColor       =   &H00DFFFFF&
      Caption         =   " Goto File Name"
      Height          =   240
      Left            =   516
      TabIndex        =   45
      Top             =   24
      Width           =   1176
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Put Whole File on Clipboard"
      Height          =   225
      Left            =   9285
      TabIndex        =   44
      Top             =   2670
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
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
      Height          =   348
      Left            =   3816
      TabIndex        =   43
      Top             =   1044
      Width           =   1128
   End
   Begin VB.Label cmdCopy 
      Alignment       =   2  'Center
      Caption         =   "Copy All Clipboard"
      Height          =   228
      Left            =   3816
      TabIndex        =   42
      Top             =   792
      Width           =   2436
   End
   Begin VB.Label Label65 
      Caption         =   "Process ID:"
      Height          =   240
      Left            =   516
      TabIndex        =   41
      Top             =   792
      Width           =   1236
   End
   Begin VB.Label Label66 
      Alignment       =   2  'Center
      BackColor       =   &H00E2FEEA&
      Caption         =   "Kill PID"
      Height          =   240
      Left            =   2832
      TabIndex        =   40
      Top             =   792
      Width           =   936
   End
   Begin VB.Label Label30 
      Caption         =   "TASKKILLER COMMAND LINE GENERATED "
      Height          =   240
      Left            =   516
      TabIndex        =   24
      ToolTipText     =   "PRESS HERE TO KILL BY PROCESS NAME"
      Top             =   540
      Width           =   4596
   End
   Begin VB.Label Label22 
      Caption         =   "TASKKILLER BY PROCESS NUMBER"
      Height          =   240
      Left            =   516
      TabIndex        =   23
      ToolTipText     =   "PRESS HERE TO KILL BY PROCESS NUMBER"
      Top             =   276
      Width           =   4596
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
      Height          =   225
      Left            =   13320
      TabIndex        =   22
      Top             =   4905
      Visible         =   0   'False
      Width           =   3315
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
      Height          =   225
      Left            =   16665
      TabIndex        =   21
      Top             =   4905
      Visible         =   0   'False
      Width           =   2070
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
      Height          =   225
      Left            =   13320
      TabIndex        =   20
      Top             =   5160
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
      Height          =   225
      Left            =   13320
      TabIndex        =   19
      Top             =   5670
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
      Height          =   225
      Left            =   13320
      TabIndex        =   18
      Top             =   5415
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
      Height          =   225
      Left            =   13320
      TabIndex        =   17
      Top             =   6165
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
      Height          =   225
      Left            =   13320
      TabIndex        =   16
      Top             =   5925
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
      Height          =   225
      Left            =   13320
      TabIndex        =   15
      Top             =   6420
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
      Height          =   210
      Left            =   13320
      TabIndex        =   14
      Top             =   6675
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label Label52 
      Caption         =   "Label52"
      Height          =   345
      Left            =   11775
      TabIndex        =   62
      Top             =   75
      Width           =   1050
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
   Begin VB.Menu MNU_VERSION 
      Caption         =   "MNU_VERSION"
   End
   Begin VB.Menu MNU_ME_ON_TOP 
      Caption         =   "ME ON TOP"
   End
   Begin VB.Menu MNU_TASK_KILLER_AUTOHOTKEYS 
      Caption         =   "KILL AUTOHOTKEY"
   End
   Begin VB.Menu MNU_AUTOHOTKEYS_SET 
      Caption         =   "RUN AUTOHOTKEY SET"
   End
   Begin VB.Menu MNU_SCREEN_SHOT_ME 
      Caption         =   "SCREENSHOT ME"
   End
   Begin VB.Menu MNU_MAXIMIZE_GOODSYNC 
      Caption         =   "MAXIMIZE GOODSYNC"
   End
   Begin VB.Menu MNU_CLOSE_GOODSYNC 
      Caption         =   "CLOSE GOODSYNC"
   End
   Begin VB.Menu MNU_GIVE_ME_TIME 
      Caption         =   "GIVE ME TIME"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_SHOW_THE_TIME 
      Caption         =   "TIME SHOW"
   End
   Begin VB.Menu MNU_GIVE_ME_TIME_WITHER_UTC 
      Caption         =   "GIVE ME TIME AND UNI_ UTC Time Toggle = YES"
   End
   Begin VB.Menu MNU_HOOVER_20_SECOND 
      Caption         =   "USE 20 SECOND HOOVER"
   End
   Begin VB.Menu MNU_WINMERGE_ON_TOP_ALLTME 
      Caption         =   "WINMERGE_ON_TOP_ALLTME=YES"
   End
   Begin VB.Menu MNU_KILL_MAX_AHK 
      Caption         =   "AHK KILL MAX"
   End
   Begin VB.Menu MNU_CMD_KILL_MAX 
      Caption         =   "CMD KILL MAX"
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
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_GIVER_ME_UPTIME 
      Caption         =   "GIVE ME UPTIME"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_LAUNCH_AUTORUNS_SET_BOOT 
      Caption         =   "AUTOHOTKEY BOOT"
   End
   Begin VB.Menu MNU_CLIPBOARDER_REPLACE_ER_AND 
      Caption         =   "CLIPBOARD REPLACE ""AND"""
   End
   Begin VB.Menu MNU_STOP_GOODSYNC_SCRIPTOR 
      Caption         =   "STOP GOODSYNC SCRIPTOR"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_GOOGLE_SYNC 
      Caption         =   "GOOGLE SYNC"
   End
   Begin VB.Menu MNU_OS_RESTART 
      Caption         =   "OS RESTART"
   End
   Begin VB.Menu Mnu_Menu_Item_Count 
      Caption         =   "Menu Item Count"
   End
   Begin VB.Menu Mnu_Control_Item_Count 
      Caption         =   "Control Item Count"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'  =============================================================
'# __ D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\VB_KEEP_RUNNER.vbp
'# __
'# __ VB_KEEP_RUNNER.vbp
'# __ VB_KEEP_RUNNER.exe
'# __
'# BY Matthew __ Matt.Lan@Btinternet.com __
'# __
'# START 01  TIME [ Tue 01-May-2018 14:02:31 ]
'# START 02  TIME [ Tue 01-May-2018 14:32:54 ]
'# LAST EDIT TIME [ Wed 20-Mar-2019 03:02:22 ]
'# __
'  =============================================================

' ------------------------------------------------------------------
' Location Internet ---- GTUHUB
' ------------------------------------------------------------------
' Matthew-Lancaster/Autokey -- 19-SCRIPT_TIMER_UTIL.ahk at master · Matthew-Lancaster/Matthew-Lancaster
' https://github.com/Matthew-Lancaster/Matthew-Lancaster/blob/master/SCRIPTER%20CODE%20--%20AUTOKEY/Autokey%20--%2019-SCRIPT_TIMER_UTIL.ahk
' ------------------------------------------------------------------
' ------------------------------------------------------------------
' Location Internet ---- GOOGLE
' ------------------------------------------------------------------
' ------------------------------------------------------------------
' VB_KEEP_RUNNER - Google Drive
' https://drive.google.com/drive/folders/1rkhnn_9zSzgr6nvzpvmI8o1mNR0t6-QE
' ------------------------------------------------------------------

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
' 1. VB_KEEP_RUNNER __ HERE
' 2. ELITESPY
' 3. CLIPBOARD LOGGER
' 4. TIDAL
' 5. CPU %
' POSSIBLE SOME OTHER THERE
'-------------------------------------------------------------------
' FROM   ---- Mon 28-May-2018 17:26:15
' FROM   ---- Mon 28-May-2018 21:11:00 4 HOUR MINUS 15 MINUTE
'-------------------------------------------------------------------


' [ Wednesday 15:36:40 Pm_Sixer March 2019 ]
' frmListMenu
' MIGHT DO WITHA SYNCER PROGRAM
' CAME FROM ELITESPY

' -------------------------------------------------------------------------------
' SESSION AT [__ Ver_2019_1.0.279 __] ---- Wed 20-Mar-2019
' FOR ROUTINE ---- Private Sub Label_GOODSYNC_01_Click()
' -------------------------------------------------------------------------------
' YOU MUST SET YOUR GOOD_SYNC LOGG TO GO WITH THE CUSTOM DIRECTORY
' THEY ARE STILL KEPT IN USUAL PLACE WITH THE SYNC FOLDER IF UNLESS SELECT NOT TO
' WHEN THEY ARE COMMON HERE
' ABLE TO TELL IF SYNC IS RUN AS GOOD SCHEDULE
' YOU ABEL TO SET A CUSTOM SCRIPT AT END OF ALL JOB TO CHECK THING OVER
' IF WHEN THAT RUN ABLE TO CHECK THING OVER
' IF SOMETHING LIKE HASN'T GROUND TO A HALT
' IF HAS GROUND TO HALT
' WHICH HARD TO TELL
' THE  MOST TIME LOGG ARE UPDATE AND RUN
' BUT THE IS NOT SHOW ANYMORE
' IF THAT WERE CASE
' GET ON TO IT AND WHEN END OF WHOLE SET JOB FINISHER
' AND CHECK THEN
' IF FIND THING NOT GOOD THEN ISSUE A CLOSE TO GOOD_SYNC AND RELOAD
' CURRENTLY I HAVE ONE AT END EVERY FINISH ALL JOB IN SCRIPT AT
' AFTER THEM ALL RUN AT BOTTOM SCRIPT
' --------------------------------------------------------------------------------
' [ Wednesday 02:42:30 Am_20 March 2019 ] WRITING
' --------------------------------------------------------------------------------
' CLIPBOARD COUNTER ---- 98 * CLIPBOARD'S COUNT
' Count = 1123 -- Wed 20-Mar-2019 01:54:15
' Count = 1221 -- Wed 20-Mar-2019 03:08:18
' -------------------------------------------------------------------------------

' ------------------------------------------------------------------
' ------------------------------------------------------------------
' ------------------------------------------------------------------
' ------------------------------------------------------------------

' ------------------------------------------------------------------
' VARIABL DECLARE BLOCK FROM ELITEPSY
' VARIABL DECLARE BLOCK FROM VB_KEEP_RUNNER
' ------------------------------------------------------------------

' UNABLE USE ShowWindow AS ANOTHER FUNCTION USER
Const ShowWindow_2 = 1, DontShowWindow = 0, DontWaitUntilFinished = False, WaitUntilFinished = True

Dim FROM_picCrossHair_MouseUP

Dim ENUMPROCESS_MUST_RUNNER

Dim ENUMPROCESS_NOT_RUN_YET

Dim OLD_ME_POS

Dim X_ONE_SECOND

Dim FOREGROUND_WINDOW_CHANGE_DELAY_1
Dim FOREGROUND_WINDOW_CHANGE_DELAY_2

Dim DO_ONCE_BOOTER

Dim OLD_hWnd_WINAMP_GetWindowState
Dim OLD_hWnd_MediaPlayerClassicW_GetWindowState

Dim RIPER

Dim LOCKED_IN_LOOP_EVENT_LEARN_KEEP_AWAY

Dim SET_NET_COPY_CHANGE_HAPPENER_LATER_WHEN_GO
Dim TOTAL_WORK_COUNT_01
Dim TOTAL_WORK_COUNT_02

Dim VB_EXE_SYNC_RUN_INSTANTLY
Dim OLD_ListView_VB_MODIFIED_ERROR_ListItems_Count

Dim Copier_Done_COUNTER_01
Dim Copier_Done_COUNTER_02
Dim Form_Resize_VB
Dim TIME_RETRY_REDO_NETWORK_ARRAY_VB_UPDATE

Dim FIRST_RUN_NETWORK_CHECKER_01
Dim FIRST_RUN_NETWORK_CHECKER_02

Dim OLD_VB_COUNT_ARRAY_SIZE

Dim TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_01
Dim TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_02
Dim SET_VB_EXE_ARRAY_FIRST_TIME
Dim NET_COPY_CHANGE_HAPPENER

Dim VB_EXE_SYNC_TIMER
Dim VB_EXE_ARRAY()

Dim OLD_V_TIME_01, V_TIME_02, SET_GOODSYNC_DONE_ONCE

Dim LISTVIEW_2_OR_3_HITT

Dim QUICK_KEY_ESCPAPE_AT_LOAD_20_SECOND


Dim TIMER_GO_COMPUTER_START

Public EXIT_TRUE

Dim FORM_LOAD_TRUE

Dim picCrossHair_MouseMove_Dragging_VAR

Dim O_Ret_Connected_To_The_Internet

Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal Index As Long) As Long
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
'MsgBox GetSystemMetrics(SM_CXSCREEN) & "x" & GetSystemMetrics(SM_CYSCREEN)

Dim ESCAPE_ERROR_WINGDINGS_2, ESCAPE_ERROR_WINGDINGS_3
Dim TYPEMESSENGER

Dim WTrue_1, WTrue_2, WTrue_3, TW1, TW2, TW3
Dim FIRST_RUN_COLOUR_CYCLE
Dim Index_String_Camera


Dim IsIDE_TEST

Dim OhWnd_VB_EXE
Dim OhWnd_VB_CLIPPER_ERROR
Dim OhWnd_VB_LOADER
Dim OhWnd_TEAM_VIEWER
Dim O_mWnd_VB_VbaWindow_MAXIMIZE

Dim OcWnd
Dim ihWnd, O_IhWnd

Dim TO_SETTER

Dim Label51_Height

Dim O_GetForegroundWindow
Dim O_GetForegroundWindow_02
Dim O_STRING_COMPARE
Dim NOT_RESIZE_EVENTER_SET_HAPPEN_ME_XW
Dim NOT_RESIZE_EVENTER_SET_HAPPEN_ME_YH

Dim Counter_ALWAYS_ON_TOP_TIMER
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
Dim OhWnd_FINDER_1
Dim OhWnd_FINDER_2
Dim Result As Long
Dim OcWnd_GOODSYNC_CRASH, OcWnd_GOOD_SYNC
Dim I_Memmer, OcWnd_ICACLS_SETTER_PERMISSION_hWnd


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
Dim lhWnd_Function_Button_Set_MIN_MAX
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

Dim O_DAY_NOW_DIR_FOR_VIDEO
Dim O_DAY_NOW_DIR_FOR_VIDEO_ME_YOU_TUBE
Dim O_DAY_NOW_DIR_FOR_VIDEO_KILLSOMETIME
Dim O_DAY_NOW_DIR_FOR_XXX_BUNKER_COM
Dim O_DAY_NOW_DIR_FOR_HARDWARE
Dim O_DAY_NOW_DIR_FOR_ARGUS_VIDEO

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

Dim ARCHIVE_Menu_Height

Dim FileName_O_VB
Dim FileName_O_AHK
Dim FileName_VB
Dim FileName_AHK
Dim OPERATION_CREATE_PATH_SET_NETWORK
Dim FileName_2
Dim FileName_4
Dim Array_FileName()
Dim Array_NETPATH_01()
Dim Array_NETPATH_02()
Dim ArrayCount
Dim NET_PATH
Dim FN_VAR_1
Dim FN_VAR_2
Dim ELEMENT1
Dim ELEMENT2
Dim ELEMENT3
Dim ELEMENT4
Dim ELEMENT5
Dim ELEMENT7

Public ScreenTwipsX, ScreenTwipsY, ScreenWidthX, ScreenHeightY, Idle_Timer_Proc

Const E = 2.7182818284
'Const pi = 3.141592648
Const hWnd_TOPMOST = -1
Const hWnd_NOTOPMOST = -2
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

Dim SET_COMPUTER_TO_RUN
Dim SET_COMPUTER_TO_RUN_PID_EXE
Dim SET_COMPUTER_TO_RUN_02
Dim FIND_COMPUTER_TO_RUN_VAR

Dim VAR_LAB_TEXT As String

Dim OLD_FindHandle_hWnd_COUNT
Dim SET_GO_CONTROL_LEFT_F1

Dim Old_lhWnd

Public VAR_FORM1_EXIST

Private Type POINTAPI
        x As Long
        y As Long
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
'Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

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

Private Declare Function Process32First Lib "Kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "Kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Private Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Function TerminateProcess Lib "Kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, hModule As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Any, ByVal lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32.dll" (ByVal hThread As Long, ByRef lpExitCode As Long) As Long
Private Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Boolean

Private Declare Function CloseHandle Lib "Kernel32" _
        (ByVal hObject As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "Kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
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
'Private Const hWnd_TOPMOST = -1
'Private Const hWnd_NOTOPMOST = -2

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "Kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

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

Private Declare Function GetShortPathName Lib "Kernel32" _
      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Declare Function FindWindowDLL Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long

Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Const GW_hWndNEXT = 2
Private Const WM_CLOSE = &H10


Private Declare Function FindFirstFile Lib "Kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "Kernel32" (ByVal hFindFile As Long) As Long

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
         ByVal x As Long, _
         ByVal y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal bRepaint As Long) As Long

Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

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
 
Private Declare Function CreateFile Lib "Kernel32" Alias _
  "CreateFileA" (ByVal lpFileName As String, _
  ByVal dwDesiredAccess As Long, _
  ByVal dwShareMode As Long, _
  ByVal lpSecurityAttributes As Long, _
  ByVal dwCreationDisposition As Long, _
  ByVal dwFlagsAndAttributes As Long, _
  ByVal hTemplateFile As Long) _
  As Long

Private Declare Function LocalFileTimeToFileTime Lib _
    "Kernel32" (lpLocalFileTime As FILETIME, _
     lpFileTime As FILETIME) As Long

Private Declare Function SetFileTime Lib "Kernel32" _
  (ByVal hFile As Long, ByVal MullP As Long, _
   ByVal NullP2 As Long, lpLastWriteTime _
   As FILETIME) As Long

Private Declare Function SystemTimeToFileTime Lib _
   "Kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime _
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
'Private Type SYSTEMTIME
'    wYear As Integer
'    wMonth As Integer
'    wDayOfWeek As Integer
'    wDay As Integer
'    wHour As Integer
'    wMinute As Integer
'    wSecond As Integer
'    wMilliseconds As Integer
'End Type

Private Declare Function CreateDirectory Lib "Kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Long) As Long
Private Declare Function GetFileTime Lib "Kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "Kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "Kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function GetFileAttributes Lib "Kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
'Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long

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

Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Dim sConnType As String * 255

'HDC SET
Private Declare Function Escape Lib "gdi32" (ByVal HDC As Long, ByVal nEscape As Long, ByVal nCount As Long, ByVal lpInData As String, lpOutData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
     
'HDC SET 2
'Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DPtoLP Lib "gdi32" (ByVal HDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
'Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal HDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal HDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function SelectClipPath Lib "gdi32" (ByVal HDC As Long, ByVal iMode As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long) As Long
     
'HDC SET 3
'Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal HDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal HDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal HDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
 
'HCD SET 4
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
'Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal HDC As Long, ByVal iCapabilitiy As Long) As Long
'Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal HDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
'Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
'Private Declare Function SelectPalette Lib "gdi32" (ByVal HDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
'Private Declare Function RealizePalette Lib "gdi32" (ByVal HDC As Long) As Long


Private Const WM_PAINT = &HF
Private Const WM_PRINT = &H317
Private Const PRF_CLIENT = &H4&    ' Draw the window's client area
Private Const PRF_CHILDREN = &H10& ' Draw all visible child
Private Const PRF_OWNED = &H20&    ' Draw all owned windows



'-----------------------
'THE UNIVERSAL TIME DOWN
'-----------------------
Private Declare Function GetTimeZoneInformation Lib "Kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public MoonPhaseDate

Private Const TIME_ZONE_ID_INVALID = -1
Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_DAYLIGHT = 2

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

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName As String * 64
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName As String * 64
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private Const A_SECOND = 0.00001158 ' one second as a fraction of a day
Private Const LPERIOD = 29.530589 ' average days between lunations
Private Const EPOCH = 8388.51399305556 ' days from 01/01/1900 til 12/18/1922 12:20:09 UT, lunation 0
Private Const pi = 3.14159265359
'-----------------------
'THE UNIVERSAL TIME UP
'-----------------------

' Dragging window
Private m_bDragging As Boolean
'-----------------------------
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PicBmp
    size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Const NCBASTAT As Long = &H33
Private Const NCBNAMSZ As Long = 16
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const HEAP_GENERATE_EXCEPTIONS As Long = &H4
Private Const NCBRESET As Long = &H32
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Const RASTERCAPS As Long = 38

Private Declare Function MessageBoxTimeout Lib "user32.dll" Alias "MessageBoxTimeoutA" ( _
ByVal hWnd As Long, _
ByVal lpText As String, _
ByVal lpCaption As String, _
ByVal uType As Long, _
ByVal wLanguageId As Long, _
ByVal dwMilliseconds As Long _
) As Long

Private Const IDYES& = 6&
Private Const IDNO& = 7&
Private Const MB_SETFOREGROUND& = &H10000
Private Const MB_YESNO& = &H4&
Private Const MB_ICONASTERISK& = &H40&
Private Const MB_TIMEDOUT& = &H7D00&


Private Const GW_hWndFIRST = 0
'Private Const GW_hWndNEXT = 2
Private Const GW_CHILD = 5

'Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
'Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long


'Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetComputerNameA Lib "Kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

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
'Private Declare Function CreateToolhelp32Snapshot Lib "Kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
'Private Declare Function Process32First Lib "Kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
'Public Declare Function Process32Next Lib "Kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
'Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
'Public Declare Function TerminateProcess Lib "Kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'Public Declare Sub GlobalMemoryStatus Lib "Kernel32" (lpBuffer As MEMORYSTATUS)
'Private Declare Sub CloseHandle Lib "Kernel32" (ByVal hPass As Long)




'----------------------------------------------------
'I PUT THIS HERE FOR THE SHELL EXECUTE LOADER PROGRAM
'----------------------------------------------------
'Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetComputerNameA Lib "Kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
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



Private Sub FOREGROUND_WINDOW_CHANGE_DELAY_1_EXTRA_TO_DO_Timer()
If FOREGROUND_WINDOW_CHANGE_DELAY_1 < Now Then
If FOREGROUND_WINDOW_CHANGE_DELAY_2 < Now Then
    Call Timer_FOREGROUND_WINDOW_CHANGE_Timer
    Call Timer_FOREGROUND_WINDOW_CHANGE_02_Timer
    FOREGROUND_WINDOW_CHANGE_DELAY_1_EXTRA_TO_DO.Enabled = False
End If
End If
End Sub

Private Sub Form_Load()
    
    ' CODE MOD
    ' COUNT BACKSLASHER
    ' --------------------------------------------------------------------
'    X1 = "D:\0 CLOUD\GD-INSYNC\rub.rim@gmail.com\snapshot (3)"
'    X2 = "D:\0 CLOUD\GD-INSYNC\rub.rim@gmail.com\snapshot (3)\20190611"
'    X3 = "D:\0 CLOUD\GD-INSYNC\rub.rim@gmail.com\snapshot (3)\20190611\A3"
'    X4 = Len(X1) - Len(Replace(X1, "\", ""))
'    X5 = Len(X2) - Len(Replace(X2, "\", ""))
'    a = a
'
'    End
    
' ------------------------------------------------------------------------------
    
'    Dim R, FR
'    ' Clipboard.GetText
'    Dim VAR_STRING As String
'
'    FR = FreeFile
'    Open "D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\new 1.txt" For Binary As FR
'        VAR_STRING = Space(LOF(FR))
'        Get #FR, , VAR_STRING
'    Close FR
'
'    Dim VAR_STRING_2
'    Dim XA
'
'    VAR_STRING_2 = Space(Len(VAR_STRING))
'    XA = 1
'
'    For R = 2 To Len(VAR_STRING) Step 2
'        Mid(VAR_STRING_2, XA, 1) = Mid(VAR_STRING, R, 1)
'        XA = XA + 1
'    Next
'
'    If Dir("D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\new 2.txt") <> "" Then
'        Kill "D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\new 2.txt"
'    End If
'
'    FR = FreeFile
'    Open "D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\new 2.txt" For Binary As FR
'        Put #FR, , VAR_STRING_2
'    Close FR
'
''    Clipboard.Clear
''    Clipboard.SetText VAR_STRING
'
'    End

' ------------------------------------------------------------------------------


    ' Call MNU_CLIPBOARDER_REPLACE_ER_AND_Click
    
    Dim i As String

    FORM_LOAD_TRUE = False
    SET_VB_EXE_ARRAY_FIRST_TIME = True
    QUICK_KEY_ESCPAPE_AT_LOAD_20_SECOND = Now + TimeSerial(0, 0, 20)
    VB_EXE_SYNC_TIMER = -1
    Form1.VAR_FORM1_EXIST = True

'    YY = Now - DateDiff("d", DateSerial(2012, 1, 1), DateSerial(2013, 7, 1)) - DateDiff("d", DateSerial(2012, 1, 1), DateSerial(2012, 12, 31))
'    Debug.Print YY
'    YY = DateDiff("n", DateSerial(2018, 1, 1), DateSerial(2019, 1, 1))
'    YY = YY / 32000
'    XX = (425 / 100) * 60
'    Debug.Print XX

    App.title = "VB_KEEP_RUNNER"
    Me.Caption = "VB_KEEP_RUNNER"

    If App.PrevInstance = True And IsIDE = False Then
        i = FindWinPart_SEARCHER_NOT_ME("VB_KEEP_RUNNER")
        ' MsgBox i
        If i > 0 Then
            ShowWindow i, SW_SHOW
            ShowWindow i, SW_RESTORE
            ShowWindow i, SW_NORMAL
            SetWindowPos i, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        End If
        End
    End If
    
    If App.PrevInstance = True And IsIDE = True Then
        'i = FindWindow(vbNullString, Me.Caption)
        i = FindWinPart_SEARCHER_NOT_ME("VB_KEEP_RUNNER")
        ShowWindow i, SW_NORMAL
    '        On Error Resume Next
        SetWindowPos i, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        End
    End If

    'KILL ITSELF IN __.EXE KILL SOFTLY
    'WHILE ISIDE LEARN
    '---------------------------------
    Dim VAR
    If IsIDE = True Then
        pid = -1
        VAR = cProcesses.GetEXEID(pid, App.Path + "\" + App.EXEName + ".exe")
        If pid <> -1 Then
            'Call Process_HIGH_PRIORITY_CLASS(PID)
            VAR = cProcesses.Process_Kill(pid)
            Beep
            End
        End If
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Call Form2_Check_Project_Date.VB_PROJECT_CHECKDATE("FORM LOAD")
    
    Dim XX
    Dim Control As Control
    XX = Label_Goto_File_Name.Left
    On Error Resume Next
    For Each Control In Me.Controls
    Control.Left = Control.Left - XX + 20
    Next
    On Error GoTo 0

    Me.BackColor = RGB(&HFE - 90, &HFF - 90, &HE1 - 90)
    Me.BackColor = RGB(40, 40, 100)
    Label60.Caption = "Me on Top 20 Sec"

    Call GO_DO_IT_WITH_LOW_END_COMPUTER


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
    With ListView_CPU_INFO
        .ColumnHeaders.Add , "CPU", "CPU", 1450, lvwColumnLeft
        .ColumnHeaders.Add , "INFO", "INFO", 4000, lvwColumnLeft
        .View = lvwReport
        .FullRowSelect = True
    End With
    With ListView_VB_MODIFIED_ERROR
        .ColumnHeaders.Add , "STATUS", "STATUS", 1000, lvwColumnLeft
        .ColumnHeaders.Add , "COMPUTER_NAME", "COMPUTER_NAME", 1000, lvwColumnLeft
        .ColumnHeaders.Add , "PATH_AND_FILE", "PATH_AND_FILE", 1000, lvwColumnLeft ' ListView_VB_MODIFIED_ERROR.width - .ColumnHeaders.Item(1).width - .ColumnHeaders.Item(2).width, lvwColumnLeft
        .ColumnHeaders.Add , "ERROR INFO", "ERROR INFO", 1000, lvwColumnLeft
        'ListView_VB_MODIFIED_ERROR.ColumnHeaders.Add , "PATH_AND_FILE", "PATH_AND_FILE", ListView_VB_MODIFIED_ERROR.width - ListView_VB_MODIFIED_ERROR.ColumnHeaders.Item(1).width - ListView_VB_MODIFIED_ERROR.ColumnHeaders.Item(2).width, lvwColumnLeft
        .View = lvwReport
        .FullRowSelect = True
    End With
    
    ListView_VB_MODIFIED_ERROR.Font.size = 7
    ListView_VB_MODIFIED_ERROR.GridLines = True
    ListView_VB_MODIFIED_ERROR.Appearance = ccFlat
    ListView_VB_MODIFIED_ERROR.HideColumnHeaders = True

    
    O_lstProcess_ListCount = -1

'    Me.Caption = App.EXEName
    
    FORM_LOAD_EnumProcess = True

    MNU_VERSION.Caption = "Ver_2019_" + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision)) ' + " _ Matt.Lan@btinternet.com"

    If Form2_Check_Project_Date.GetWindowsVersion = 5.1 Then
        MNU_PIN_ITEM_BATCH_VBS.Visible = False
    End If

    Label54.FontSize = 10
    
    Call READ_ALL_NETWORK_COMPUTER_NAME_PATH_INTO_ARRAY
    
    FIRST_RUN_FOR_TOP_AND_LEFT = 6
    Counter_ALWAYS_ON_TOP_TIMER = 20
    FORM_LOAD_TRUE = True
    
'    Exit Sub
    
    ' Me.Visible = False
    Call RUN_BLOCK_OF_SETUP_CONTROLS_POSTION
    
    
    
    Call WIDTH_AND_HEIGHT(WX, HY)

    ' ONE_MILLISECOND_Timer.Enabled = True

    TIMER_GO_COMPUTER_START = 4
    
'    MsgBox Me.WindowState
    'MsgBox vbNormal
    
    
    'If IsIDE = True Then Me.WindowState = vbNormal
'    If IsIDE = False Then Me.WindowState = vbMinimized

    ' Me.Hide
'     Me.WindowState = vbMinimized
'    Me.WindowState = vbNormal
    'Me.Hide
    ONE_MILLISECOND_Timer.Enabled = False
    'ONE_MILLISECOND_Timer.Interval = 0
    'ONE_MILLISECOND_Timer.Interval = 100
    'ONE_MILLISECOND_Timer.Enabled = True


Dim OR_LOGIC
OR_LOGIC = 0
If InStr(Command$, "MINIMAL") > 0 Then OR_LOGIC = 2
If InStr(Command$, "MIN") > 0 Then OR_LOGIC = 2
If InStr(Command$, "MAXIMUM") > 0 Then OR_LOGIC = 1
If InStr(Command$, "TASKBAR") > 0 Then OR_LOGIC = 1
If InStr(Command$, "T") > 0 Then OR_LOGIC = 1
If IsIDE = True Then OR_LOGIC = 1
' Me.Visible = True
' DoEvents
If OR_LOGIC = 1 Then
'    Me.WindowState = vbNormal
    
    ' MsgBox "HH"
    
    ' Exit Sub
End If

If OR_LOGIC = 1 Then
    ' Me.WindowState = vbNormal
    'Exit Sub
End If

If OR_LOGIC = 2 Then
     Me.WindowState = vbMinimized
End If


Timer_Pause_Update.Interval = 60000
Label53.ToolTipText = "Pause Update for 1 Minute"

Label_RUN_AUTOHOTKEY_SET_NETWORK.Visible = False

End Sub

Private Sub Form_Activate_2()
    'Call Form3_CHROME_X_BUTTON_OFF.SetXState(Me.hWnd, False)
    'Call Form3_CHROME_X_BUTTON_OFF.SET_CHROME_WINDOW_X_BUTTON_CLOSE_TO_OFF
    'Call Form4_DISABLE_CLOSE_BUTTON.SET_BUTTON_CHROME
End Sub


Private Sub Form_Resize()

' Call Form_Activate_3

Timer_Pause_Update.Interval = 4000
Timer_Pause_Update.Enabled = True
Label53.BackColor = Label59.BackColor

If NOT_RESIZE_EVENTER = True Then Exit Sub

If O_Me_WindowState <> Me.WindowState Then
    Form_Resize_VB = True
End If

If O_Me_WindowState = vbMaximized And Me.WindowState = vbNormal Then Exit Sub


If Me.WindowState = vbMinimized Then
    ListView_VB_MODIFIED_ERROR.Visible = False
End If

Timer_ALWAYS_ON_TOP_TO_START_WITH_ER.Enabled = True
Counter_ALWAYS_ON_TOP_TIMER = 20
Label60.BackColor = RGB(255, 255, 255)

If RESIZE_LOOP_STOP = True Then
    RESIZE_LOOP_STOP = False
    Exit Sub
End If

'--------SUB __ frmMain.EnumProcess IT IS IN HERE
If Me.WindowState <> vbMinimized Then
'    Call EnumProcess
    Timer_EnumProcess.Interval = 1000
End If

'RESIZE
'If CMD_Process_STATE_TO_SET_1ST = True Then
'    Call cmdProcess_Click
'End If
'------------------------------------------------
Call WIDTH_AND_HEIGHT(WX, HY)

If RUN_ONCE_DEBUG_PRINT_HY_RESIZE = False Then
    RUN_ONCE_DEBUG_PRINT_HY_RESIZE = True
    Debug.Print vbCrLf & Time & "-------"
    Debug.Print "Control.Name __ " + TT
    Debug.Print "Control.Top + Control.Height __ " + TTHY1
    Debug.Print "Control.Top __ " + TTHY2
    Debug.Print "Control.Height __ " + TTHY3
End If

If Me.height + Me.width = HY2 + WX2 Then Exit Sub
If Me.WindowState = vbMinimized Then Exit Sub

If O_Me_WindowState <> vbMaximized And Me.WindowState = vbMaximized Then Exit Sub
O_Me_WindowState = Me.WindowState
'VBNORMAL    = 0
'VBMAXIMIZED = 2

On Error Resume Next
'CHANGE HEIGHT DONT RELOOP RESIZE
'--------------------------------

RESIZE_LOOP_STOP = True
Me.width = WX2

'TOP AND LEFT ONLY FIRST RUN
'DO HERE LAST OF ALL
'If Me.WindowState <> 1 Then Exit Sub
If FIRST_RUN_FOR_TOP_AND_LEFT > 0 Then
    FIRST_RUN_FOR_TOP_AND_LEFT = FIRST_RUN_FOR_TOP_AND_LEFT - 1
    Debug.Print FIRST_RUN_FOR_TOP_AND_LEFT
    '6 RUN HAPPEN BEFORE LAST SET IS PROPER AT FORM_LOAD
    
    RESIZE_LOOP_STOP = True
    Me.Top = 0 '20 'TOP_HEIGHT 'Screen.Height / 2 - Me.Height / 2 - 800
    
    'MsgBox GetSystemMetrics(SM_CXSCREEN)*Screen.TwipsPerPixelX & "x" & GetSystemMetrics(SM_CYSCREEN)
    
    RESIZE_LOOP_STOP = True
    Me.Left = GetSystemMetrics(SM_CXSCREEN) * Screen.TwipsPerPixelX / 2 - Me.width / 2
End If

'-------------------------------------------------------------------------
'THE MENU HAS TO BE LOADED WITH HEIGHT AND THEN REDONE AGAIN TWICE FOR X Y
Call WIDTH_AND_HEIGHT(WX, HY)

If Me.height + Me.width = HY2 + WX2 Then Exit Sub

On Error Resume Next
'CHANGE HEIGHT DONT RELOOP RESIZE
'--------------------------------
RESIZE_LOOP_STOP = True

''FORM_RESIZE
'Call WIDTH_AND_HEIGHT(WX, HY)
'Me.height = HY2 - 100 '+ 350

RESIZE_LOOP_STOP = True
Me.width = WX2 + 80


Exit Sub

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

Sub COLOUR_BOX_SELECTOR_RESTORE_DEFAULT()
'CALL COLOUR_BOX_SELECTOR_RESTORE_DEFAULT

Dim ARRAY_CB()
ReDim ARRAY_CB(100)
Dim LDAC, R_COUNTER
Dim COLOUR_VAR
Dim Test

' --------------------------------------------------
' SOME AFTER HERE MIXED FROM ANOTHER APP
' VB_KEEP_RUNNER AND ELITESPY
' PASTED IN
' FLOUR -- COLOUR -- POWDER --
' --------------------------------------------------
LDAC = 0
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_GOODSYNC_COLLECTION_SCRIPT_RUN"  ' -- VB_KEEP_RUNNER ONLY
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_KILL_AUTOHOTKEY"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_RUN_AUTOHOTKEY_SET"              ' -- VB_KEEP_RUNNER ONLY
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_MAXIMIZE_GOODSYNC"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_MINIMIZE_GOODSYNC"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_KILL_WSCRIPT"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_KILL_HUBIC"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Lab_KILL_AHK"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Lab_KILL_EXPLORER"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_CLOSE_GOODSYNC"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_CLOSE_hWnd"                      ' -- ELITESPY ONLY
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_TASK_KILLER_CMD"                 ' -- VB_KEEP_RUNNER ONLY
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "LABEL_KILL_NOT_RESPOND"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "LABEL_KILL_NOT_RESPOND_FORCE"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_RUN_AUTOHOTKEY_SET_NETWORK"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_KILL_CHROME"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_1X"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_2E"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_3L"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_4G"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_5P"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_7G"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_8M"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_KILL_CMD"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_KILL_CMD_AND_AHK"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_HWND_MAXIMIZE"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_MAXIMIZE_ELITE_SPY"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_MAXIMIZE_GOODSYNC2GO"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_CLOSE_GOODSYNC2GO"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_MAXIMIZE_CLIPBOARD_LOGGER"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_KILL_AND_RUN_ANOTHER_AUTOHOTKEY"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_CLOSE_OUTLOOK"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_CHROME_PAGE_AUTO_ON"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_CLOSE_EXPLORER"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "Label_DE_DUPE_EXPLORER"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = ""
LDAC = LDAC + 1
ARRAY_CB(LDAC) = ""
LDAC = LDAC + 1
ARRAY_CB(LDAC) = ""
LDAC = LDAC + 1
ARRAY_CB(LDAC) = ""
LDAC = LDAC + 1
ARRAY_CB(LDAC) = ""

Test = "_1X_2E_3L_4G_5P_7G_8M"
Test = Replace(Test, "_", "Label_")

For R_COUNTER = 1 To LDAC
    If ARRAY_CB(R_COUNTER) = "" Then Exit For
Next

ReDim Preserve ARRAY_CB(R_COUNTER - 1)

Dim Control As Control
For Each Control In Me.Controls
    For R_COUNTER = 1 To UBound(ARRAY_CB)
        If Control.Name = ARRAY_CB(R_COUNTER) Then
            If InStr(Test, Control.Name) > 0 Then
                COLOUR_VAR = &HE2FEEA
            Else
                COLOUR_VAR = Label59.BackColor
            End If
            Control.BackColor = COLOUR_VAR
            Control.ForeColor = RGB(0, 0, 0)
        End If
    Next
Next

Label_RUN_AUTOHOTKEY_SET_NETWORK.BackColor = &H808080

End Sub


Sub GO_DO_IT_WITH_LOW_END_COMPUTER()
RIPER = False
If GetComputerName = "1-ASUS-X5DIJ" Then RIPER = True
If GetComputerName = "2-ASUS-EEE" Then RIPER = True
If GetComputerName = "3-LINDA-PC" Then RIPER = True
If GetComputerName = "5-ASUS-P2520LA" Then RIPER = True
If RIPER = True Then
    Label_VB_MODIFIED_TIME.Caption = "Not GO"
    ListView_VB_MODIFIED_ERROR.Visible = False
    Label_VB_MODIFIED_TIME.Visible = False
    Label_VB_MODIFIED_ACTIVE.Visible = False
    Label44.Visible = False
    Exit Sub
End If

End Sub


Sub VB_EXE_SYNC()

Exit Sub

Dim R, A1, A2
Dim TxtEXE_INFO, X1_F
Dim F1, F2
Dim VAR_DATE_1, VAR_DATE_2
Dim TT_0, TT_1, TT_2
Dim NET_NAME_1
Dim NET_NAME_2
Dim NET_NAME_12
Dim NET_NAME_22
Dim VB_EXE_ARRAY_COUNTER
Dim VB_EXE_ARRAY_COMPARE
Dim Copier_Done_Var
Dim Copier_Done_COUNTER_01
Dim VB_COUNT_ARRAY_SIZE
Dim TT_VAR_0
Dim TT_VAR_1
Dim TT_VAR_2
Dim TT_VAR_3
Dim APP_PATH_01
Dim APP_PATH_02
Dim VERIFY_ERROR_01
Dim VERIFY_ERROR_02
Dim RQUEEN
Dim Err_Description

If LOCKED_IN_LOOP_EVENT_LEARN_KEEP_AWAY = True Then Exit Sub
DoEvents

If VB_EXE_SYNC_RUN_INSTANTLY = True Then
    ' ----------------------------------
    Call TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_SUB
    VB_EXE_SYNC_RUN_INSTANTLY = False
    NET_COPY_CHANGE_HAPPENER = True
    VB_EXE_SYNC_TIMER = 0
    ' ----------------------------------
Else
    ' ----------------------------------
    DoEvents
    
    ' Call GO_DO_IT_WITH_LOW_END_COMPUTER
    If Label_VB_MODIFIED_TIME.Visible = False Then
        Exit Sub
    End If
    
    If SET_VB_EXE_ARRAY_FIRST_TIME = True Then
        ReDim Preserve VB_EXE_ARRAY(1)
        SET_VB_EXE_ARRAY_FIRST_TIME = False
    End If
    
    If TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_01 > VB_EXE_SYNC_TIMER Then
        TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_01 = 0
        VB_EXE_SYNC_TIMER = TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_01 + VB_EXE_SYNC_TIMER
    End If
    
    If TIME_RETRY_REDO_NETWORK_ARRAY_VB_UPDATE > 0 Then
        If RIPER = False Then
            SET_NET_COPY_CHANGE_HAPPENER_LATER_WHEN_GO = True
        End If
    End If
    
    If TIME_RETRY_REDO_NETWORK_ARRAY_VB_UPDATE > VB_EXE_SYNC_TIMER Then
        TIME_RETRY_REDO_NETWORK_ARRAY_VB_UPDATE = 0
        VB_EXE_SYNC_TIMER = TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_01 + VB_EXE_SYNC_TIMER
    End If
    
    If VB_EXE_SYNC_TIMER = -1 Then
        VB_EXE_SYNC_TIMER = Now + TimeSerial(0, 0, 5)
    End If
    If VB_EXE_SYNC_TIMER = 0 Then
        VB_EXE_SYNC_TIMER = Now + TimeSerial(0, 0, 30)
        Label_VB_MODIFIED_ACTIVE = "IDLE"
    End If
    If VB_EXE_SYNC_TIMER > Now + TimeSerial(0, 0, 20) Then
        If FIRST_RUN_NETWORK_CHECKER_02 = False Then
            FIRST_RUN_NETWORK_CHECKER_01 = True
            ' HERE TO SET VAR
            ' NET_COPY_CHANGE_HAPPENER = True
        End If
    End If
    
    Call TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_SUB
    
    Call CHECK_NET_ARRAY
    
    If NET_COPY_CHANGE_HAPPENER = True Then
        If VB_EXE_SYNC_TIMER > Now + TimeSerial(0, 0, 5) Then
            VB_EXE_SYNC_TIMER = Now + TimeSerial(0, 0, 5)
        End If
    End If
    
    If VB_EXE_SYNC_TIMER > Now Then
        Exit Sub
    End If
    
    ' ALL SET GO
    If VB_EXE_SYNC_TIMER <= Now Then
        VB_EXE_SYNC_TIMER = 0
        If SET_NET_COPY_CHANGE_HAPPENER_LATER_WHEN_GO = True Then
            TIME_RETRY_REDO_NETWORK_ARRAY_VB_UPDATE = 0
            NET_COPY_CHANGE_HAPPENER = True
        End If
    End If
    ' ----------------------------------
End If


LOCKED_IN_LOOP_EVENT_LEARN_KEEP_AWAY = True


VB_EXE_ARRAY_COUNTER = UBound(VB_EXE_ARRAY) - 1

VB_EXE_ARRAY_COMPARE = ""
For R = 1 To UBound(VB_EXE_ARRAY)
    VB_EXE_ARRAY_COMPARE = VB_EXE_ARRAY_COMPARE + "----" + UCase(VB_EXE_ARRAY(R))
Next

VB_COUNT_ARRAY_SIZE = 0
For R = 1 To lstProcess_3_SORTER_ListView.ListItems.Count
    A1 = lstProcess_3_SORTER_ListView.ListItems.Item(R).SubItems(1)
    A2 = lstProcess_3_SORTER_ListView.ListItems.Item(R)
    If Val(A2) > 0 Then
        TxtEXE_INFO = GetFileFromProc(Val(A2))
        If InStr(TxtEXE_INFO, "D:\VB6\") > 0 Then
            ' GET COUNT SIZE OF ALL PROCESS TO DO WITH VB IN PROCESS EXPLORER
            ' WANT MULTIPLE OF CHANGE TRIGGER EVENT CHECKER
            ' ---------------------------------------------------------------
            VB_COUNT_ARRAY_SIZE = VB_COUNT_ARRAY_SIZE + 1
            If InStr(VB_EXE_ARRAY_COMPARE, UCase(TxtEXE_INFO) + "----") = 0 Then
                VB_EXE_ARRAY_COMPARE = VB_EXE_ARRAY_COMPARE + "----" + UCase(TxtEXE_INFO)
                VB_EXE_ARRAY_COUNTER = VB_EXE_ARRAY_COUNTER + 1
                ReDim Preserve VB_EXE_ARRAY(VB_EXE_ARRAY_COUNTER)
                VB_EXE_ARRAY(VB_EXE_ARRAY_COUNTER) = TxtEXE_INFO
            End If
        End If
    End If
Next
' PUT OUR OWN EXE IN JUST IN CASE NOT RUNNER WHILE IDE EVIROMENT
' AND A FEW MORE
' MINE CHECKER _ HARD CODE AT MOMENT
' --------------------------------------------------------------
Dim ARRAY_SOD(11)
ARRAY_SOD(1) = App.Path + "\" + App.EXEName + ".EXE"
ARRAY_SOD(2) = "D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\#0 Send To Multi Menu.exe"
ARRAY_SOD(3) = "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.exe"
ARRAY_SOD(4) = "D:\VB6\VB-NT\00_Best_VB_01\URL Logger\URL Logger.exe"
ARRAY_SOD(5) = "D:\VB6\VB-NT\00_Best_VB_01\CPU % OF A PROGRAM\CPU % INDIVIDUAL PROCESS.exe"
ARRAY_SOD(6) = "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\Tidal.exe"
ARRAY_SOD(7) = "D:\VB6\VB-NT\00_Best_VB_01\CLIPBOARD_VIEWER\ClipBoard Viewer.exe"
ARRAY_SOD(8) = "D:\VB6\VB-NT\00_Best_VB_01\Shell Explorer Loader\Shell Explorer Loader.exe"
ARRAY_SOD(9) = "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\Cid-RunMe.exe"
ARRAY_SOD(10) = "D:\VB6\VB-NT\00_Best_VB_01\Shell VBasic 6 Loader\Shell VBasic 6 Loader.exe"
ARRAY_SOD(11) = "D:\VB6\VB-NT\00_Best_VB_01\Batch_Compiler_Auto\BatchCompiler.exe"

For R = 1 To UBound(ARRAY_SOD)
    TxtEXE_INFO = ARRAY_SOD(R)
    If InStr(VB_EXE_ARRAY_COMPARE, UCase(TxtEXE_INFO) + "----") = 0 Then
        VB_EXE_ARRAY_COMPARE = VB_EXE_ARRAY_COMPARE + "----" + UCase(TxtEXE_INFO)
        VB_EXE_ARRAY_COUNTER = VB_EXE_ARRAY_COUNTER + 1
        ReDim Preserve VB_EXE_ARRAY(VB_EXE_ARRAY_COUNTER)
        VB_EXE_ARRAY(VB_EXE_ARRAY_COUNTER) = TxtEXE_INFO
    End If
Next
' NET_COPY_CHANGE_HAPPENER = False
' ----------------------------------------------------------------------
' IF A PROCESS HAS LOADED AND GONE THEN DO ALL THE NETWORK CHECK ANOTHER
' BEST WAY TO CATCH RATHER THAN HASHER PID NUMBER WITH FILE PATH
' ----------------------------------------------------------------------
If OLD_VB_COUNT_ARRAY_SIZE > 0 And OLD_VB_COUNT_ARRAY_SIZE <> VB_COUNT_ARRAY_SIZE Then
    NET_COPY_CHANGE_HAPPENER = True
End If

OLD_VB_COUNT_ARRAY_SIZE = VB_COUNT_ARRAY_SIZE

For R = 1 To UBound(VB_EXE_ARRAY)
    TxtEXE_INFO = VB_EXE_ARRAY(R)
    If InStr(TxtEXE_INFO, "D:\VB6\") > 0 Then
        X1_F = Replace(TxtEXE_INFO, "\VB6\", "\VB6-EXE\")
        Err.Clear
        NET_NAME_1 = TxtEXE_INFO
        NET_NAME_2 = X1_F
        NET_NAME_12 = NET_NAME_1
        NET_NAME_22 = NET_NAME_2
        On Error Resume Next
        Set F1 = FSO.GetFile(NET_NAME_1)
        Set F2 = FSO.GetFile(NET_NAME_2)
        VAR_DATE_1 = 0
        VAR_DATE_2 = 0
        VAR_DATE_1 = F1.DateLastModified
        VAR_DATE_2 = F2.DateLastModified
        If Err.Number > 0 Then
            VAR_DATE_1 = 0
            VAR_DATE_2 = 0
        End If
'            Debug.Print TxtEXE_INFO
'            Debug.Print VAR_DATE_1
'            Debug.Print VAR_DATE_2
        ' Debug.Print X1_F
        If VAR_DATE_1 > 0 Then
            If VAR_DATE_1 > VAR_DATE_2 Then
                FSO.CopyFile NET_NAME_1, NET_NAME_2
                NET_COPY_CHANGE_HAPPENER = True
                TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_02 = Now
            End If
            If VAR_DATE_1 < VAR_DATE_2 Then
                FSO.CopyFile NET_NAME_2, NET_NAME_1
                NET_COPY_CHANGE_HAPPENER = True
                TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_02 = Now
            End If
        End If
        Set F1 = Nothing
        Set F2 = Nothing
    End If
DoEvents
Next
               
                
If FIRST_RUN_NETWORK_CHECKER_01 = True Then
    FIRST_RUN_NETWORK_CHECKER_02 = True
    FIRST_RUN_NETWORK_CHECKER_01 = False
    NET_COPY_CHANGE_HAPPENER = True
End If
            
            
If NET_COPY_CHANGE_HAPPENER = True Then
    TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_01 = 0
    TIME_RETRY_REDO_NETWORK_ARRAY_VB_UPDATE = 0
    VERIFY_ERROR_02 = False
    
    If Array_FileName(1) = "" Then
        Call READ_ALL_NETWORK_COMPUTER_NAME_PATH_INTO_ARRAY
    End If
    
    ListView_VB_MODIFIED_ERROR.ListItems.Clear
    
    ' TWICE FOR GOOD FUN
    ' ------------------
    Copier_Done_COUNTER_02 = 0
    TOTAL_WORK_COUNT_01 = 0
    TOTAL_WORK_COUNT_02 = 0
    ' ------------------
    Do
    ' ------------------
    
    Copier_Done_COUNTER_01 = 0
    Copier_Done_Var = False  ' TO DO WITH LOOP
    
    ' GET A COUNT OF WORK FOR DISPLAY 01 OF 10 EXAMPLE
    If Copier_Done_COUNTER_02 = 0 Then
        For TT_0 = 2 To 1 Step -1
        For TT_1 = UBound(Array_FileName) To 1 Step -1
        For TT_2 = UBound(Array_FileName) To 1 Step -1
            TxtEXE_INFO = VB_EXE_ARRAY(TT_1)
            NET_NAME_1 = TxtEXE_INFO
            NET_NAME_2 = Replace(TxtEXE_INFO, "\VB6\", "\VB6-EXE\")
            NET_NAME_12 = NET_NAME_1
            NET_NAME_22 = NET_NAME_2
            
            TT_VAR_3 = Format(Copier_Done_COUNTER_02, "0")
            TT_VAR_0 = Format(2 - TT_0, "0")
            TT_VAR_1 = Format(TT_1, "0")
            TT_VAR_2 = Format(TT_2, "0")
            Label_VB_MODIFIED_ACTIVE.Caption = "GO " + TT_VAR_3 + "," + TT_VAR_0 + "," + TT_VAR_1 + "," + TT_VAR_2
            Label_VB_MODIFIED_ACTIVE.FontSize = 9.2
            If UCase(Array_NETPATH_01(TT_1)) <> UCase(Array_NETPATH_01(TT_2)) Then
                TOTAL_WORK_COUNT_01 = TOTAL_WORK_COUNT_01 + 1
            End If
        Next
        Next
        Next
    End If
    
    ' -------------------------------------------------------------------
    ' FOR HIGER SPEED WE ARE WORKING BACKWARD
    ' AND THEN THE SORT ORDER OF MORE HIGHER SPEED COMPUTER IS COME FIRST
    ' -------------------------------------------------------------------
    For TT_0 = 2 To 1 Step -1
    For TT_1 = UBound(Array_FileName) To 1 Step -1
    For TT_2 = UBound(Array_FileName) To 1 Step -1
        DoEvents
        TxtEXE_INFO = VB_EXE_ARRAY(TT_1)
        NET_NAME_1 = TxtEXE_INFO
        NET_NAME_2 = Replace(TxtEXE_INFO, "\VB6\", "\VB6-EXE\")
        NET_NAME_12 = NET_NAME_1
        NET_NAME_22 = NET_NAME_2
        
        TT_VAR_3 = Format(Copier_Done_COUNTER_02, "0")
        TT_VAR_0 = Format(2 - TT_0, "0")
        TT_VAR_1 = Format(TT_1, "0")
        TT_VAR_2 = Format(TT_2, "0")
        Label_VB_MODIFIED_ACTIVE.Caption = "GO " + TT_VAR_3 + "," + TT_VAR_0 + "," + TT_VAR_1 + "," + TT_VAR_2
        Label_VB_MODIFIED_ACTIVE.FontSize = 9.2
        If UCase(Array_NETPATH_01(TT_1)) <> UCase(Array_NETPATH_01(TT_2)) Then
            TOTAL_WORK_COUNT_02 = TOTAL_WORK_COUNT_02 + 1
        End If
        
        If UCase(Array_NETPATH_01(TT_1)) <> UCase(Array_NETPATH_01(TT_2)) Then
            'DoEvents
            If Form_Resize_VB = True Then
                Form_Resize_VB = False
                TIME_RETRY_REDO_NETWORK_ARRAY_VB_UPDATE = Now + TimeSerial(0, 0, 30)
                LOCKED_IN_LOOP_EVENT_LEARN_KEEP_AWAY = False
                Exit Sub
            End If
            
            NET_NAME_1 = NET_NAME_12
            NET_NAME_2 = NET_NAME_22
            NET_NAME_1 = "\\" + Array_NETPATH_01(TT_1) + "\" + "" + Array_NETPATH_02(TT_1) + "_02_d_drive\" + Mid(NET_NAME_1, 4)
            NET_NAME_2 = "\\" + Array_NETPATH_01(TT_2) + "\" + "" + Array_NETPATH_02(TT_2) + "_02_d_drive\" + Mid(NET_NAME_2, 4)
            
            APP_PATH_01 = UCase("\\" + GetComputerName + "\" + "" + Replace(GetComputerName, "-", "_") + "_02_d_drive\" + Mid(NET_NAME_1, 4))
            APP_PATH_02 = UCase("\\" + Array_NETPATH_01(TT_2) + "\" + "" + Array_NETPATH_02(TT_2) + "_02_d_drive\" + Mid(NET_NAME_2, 4))
                                                
            If TT_0 = 2 Then
                NET_NAME_2 = NET_NAME_12
                NET_NAME_2 = "\\" + Array_NETPATH_01(TT_2) + "\" + "" + Array_NETPATH_02(TT_2) + "_02_d_drive\" + Mid(NET_NAME_2, 4)
            End If
                            
            ' ------------------------------------------------------
            ' DON'T COPIER ANYTHING TO DESTINATION YOUR OWN COMPUTER
            ' OR CRASH THE PROGRAM
            ' ------------------------------------------------------
            If UCase(NET_NAME_2) <> APP_PATH_02 Then
                Err.Clear
                On Error Resume Next
                DoEvents
                Set F1 = FSO.GetFile(NET_NAME_1)
                DoEvents
                Set F2 = FSO.GetFile(NET_NAME_2)
                DoEvents
                VAR_DATE_1 = 0
                VAR_DATE_2 = 0
                DoEvents
                VAR_DATE_1 = F1.DateLastModified
                DoEvents
                VAR_DATE_2 = F2.DateLastModified
                DoEvents
                
                If Err.Number > 0 Then
                    VAR_DATE_1 = 0
                    VAR_DATE_2 = 0
                End If
                If VAR_DATE_1 > 0 Then
                Call ADD_TO_ListView_VB_MODIFIED_ERROR(NET_NAME_1, NET_NAME_2, 1, Err.Number, Err.Description)
                TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_02 = Now
                Call TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_SUB
                
                    If VAR_DATE_1 > VAR_DATE_2 Then
                        VERIFY_ERROR_01 = False
                        For RQUEEN = 1 To 1
                            DoEvents
                            Err.Clear
                            FSO.CopyFile NET_NAME_1, NET_NAME_2
                            DoEvents
                            '----------------------------------
                            ' WHY THIS ERROR
                            ' "Index out of bounds"
                            ' DEBUG TRACE PROPER MESSENGER IS
                            ' "Permission denied"
                            ' GOT TO RUN TWICE TO GET PROPER IT
                            '----------------------------------
                            Err_Description = Err.Description
                            If Err.Number = 70 Then Err_Description = "Permission denied"
                            If Err_Description <> "Index out of bounds" Then Exit For
                            If Err.Number = 0 Then Exit For
                        Next
                        If Err.Number > 0 Then
                            Call ADD_TO_ListView_VB_MODIFIED_ERROR(NET_NAME_2, NET_NAME_1, 0, Err.Number, Err_Description)
                        End If
                        ' ------------------
                        ' AFTER COPIER CHECK
                        ' ------------------
                        If Err.Number = 0 Then
                            DoEvents
                            Set F1 = FSO.GetFile(NET_NAME_1)
                            DoEvents
                            Set F2 = FSO.GetFile(NET_NAME_2)
                            DoEvents
                            If F1.DateLastModified <> F2.DateLastModified Then VERIFY_ERROR_01 = True
                            DoEvents
                            If F1.size <> F2.size Then VERIFY_ERROR_01 = True
                            DoEvents
                            If VERIFY_ERROR_01 = True Then VERIFY_ERROR_02 = True
                            If VERIFY_ERROR_01 = True Then
                                Call ADD_TO_ListView_VB_MODIFIED_ERROR(NET_NAME_2, NET_NAME_1, 3, Err.Number, Err.Description)
                            End If
                        End If
                        If Err.Number = 0 And VERIFY_ERROR_01 = False Then
                            Copier_Done_Var = True
                            ' ---------------------------------------------------------------
                            ' LOOK AT THESE TWO LINE ONE BELOW A BIT
                            ' IT SEEMS COMPARE ERROR
                            ' BUT WHAT HAPPENING IS SOME FILE GET LOCKED UP TWICE
                            ' TO GET UPDATE DONE
                            ' MAYBE ONE IS UPDATED AND THE TIME OUT OF SYNC AGAIN
                            ' AND DONE AGAIN IN ANOTHER DIRECTION
                            ' ---------------------------------------------------------------
                            Call ADD_TO_ListView_VB_MODIFIED_ERROR(NET_NAME_2, NET_NAME_1, 2, Err.Number, Err.Description)
                        End If
                    End If
                End If
            End If
            'DoEvents
            
            If Form_Resize_VB = True Then
                Form_Resize_VB = False
                TIME_RETRY_REDO_NETWORK_ARRAY_VB_UPDATE = Now + TimeSerial(0, 0, 30)
                LOCKED_IN_LOOP_EVENT_LEARN_KEEP_AWAY = False
                Exit Sub
            End If

            ' ------------------------------------------------------
            ' DON'T COPIER ANYTHING TO DESTINATION YOUR OWN COMPUTER
            ' OR CRASH THE PROGRAM
            ' ------------------------------------------------------
            If UCase(NET_NAME_1) <> APP_PATH_01 Then
                If VAR_DATE_2 > 0 Then
                    Call ADD_TO_ListView_VB_MODIFIED_ERROR(NET_NAME_2, NET_NAME_1, 1, Err.Number, Err.Description)
                    TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_02 = Now
                    Call TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_SUB
                    
                    If VAR_DATE_1 < VAR_DATE_2 Then
                        VERIFY_ERROR_01 = False
                        For RQUEEN = 1 To 1
                            Err.Clear
                            DoEvents
                            FSO.CopyFile NET_NAME_2, NET_NAME_1
                            DoEvents
                            '----------------------------------
                            ' WHY THIS ERROR
                            ' "Index out of bounds"
                            ' DEBUG TRACE PROPER MESSENGER IS
                            ' "Permission denied"
                            ' GOT TO RUN TWICE TO GET PROPER IT
                            ' BUT QUICKER ERR.NUMBER IS OKAY ERROR 70 "Permission denied"
                            '----------------------------------
                            Err_Description = Err.Description
                            If Err.Number = 70 Then Err_Description = "Permission denied"
                            If Err_Description <> "Index out of bounds" Then Exit For
                            If Err.Number = 0 Then Exit For
                        Next
                        
                        If Err.Number > 0 Then
                            Call ADD_TO_ListView_VB_MODIFIED_ERROR(NET_NAME_1, NET_NAME_2, 0, Err.Number, Err_Description)
                        End If
                        ' ------------------
                        ' AFTER COPIER CHECK
                        ' ------------------
                        If Err.Number = 0 Then
                            DoEvents
                            Set F1 = FSO.GetFile(NET_NAME_1)
                            DoEvents
                            Set F2 = FSO.GetFile(NET_NAME_2)
                            DoEvents
                            If F1.DateLastModified <> F2.DateLastModified Then VERIFY_ERROR_01 = True
                            DoEvents
                            If VERIFY_ERROR_01 = True Then VERIFY_ERROR_02 = True
                            If F1.size <> F2.size Then VERIFY_ERROR_01 = True
                            DoEvents
                            If VERIFY_ERROR_01 = True Then
                                Call ADD_TO_ListView_VB_MODIFIED_ERROR(NET_NAME_1, NET_NAME_2, 3, Err.Number, Err.Description)
                            End If
                        End If
                        If Err.Number = 0 And VERIFY_ERROR_01 = False Then
                            Copier_Done_Var = True
                            ' ---------------------------------------------------------------
                            ' LOOK AT THESE TWO LINE ONE UP A BIT
                            ' ---------------------------------------------------------------
                            Call ADD_TO_ListView_VB_MODIFIED_ERROR(NET_NAME_1, NET_NAME_2, 2, Err.Number, Err.Description)
                        End If
                    End If
                End If
            End If
            Set F1 = Nothing
            Set F2 = Nothing
        End If
    Next
    Next
    Next
    
    If Copier_Done_Var = True Then
        Copier_Done_COUNTER_01 = Copier_Done_COUNTER_01 + 1
        Copier_Done_COUNTER_02 = Copier_Done_COUNTER_02 + 1
    End If
    
    Loop Until Copier_Done_COUNTER_02 > 4 Or Copier_Done_Var = False
    
    'REMOVE ListView
    'If ListView_VB_MODIFIED_ERROR.ListItems.Count > 1 Then
    'ListView_VB_MODIFIED_ERROR.ListItems.Remove (1)
    ListView_VB_MODIFIED_ERROR.ListItems(1).Text = Replace(ListView_VB_MODIFIED_ERROR.ListItems(1).Text, " _ ", " ")
    ListView_VB_MODIFIED_ERROR.ListItems(1).SubItems(1) = Format(Now, "DDD DD-MMM-YYYY HH:MM:SS AM/PM")
    ListView_VB_MODIFIED_ERROR.ListItems(1).SubItems(2) = ""

    'End If
    NET_COPY_CHANGE_HAPPENER = False
    If VERIFY_ERROR_01 = True Then
        'NET_COPY_CHANGE_HAPPENER = True
        TIME_RETRY_REDO_NETWORK_ARRAY_VB_UPDATE = Now + TimeSerial(0, 0, 30)
    End If
    
    Call LV_AutoSizeColumn(ListView_VB_MODIFIED_ERROR, Nothing)

End If

VB_EXE_SYNC_TIMER = 0

Call MARK_LABEL_RED_ERROR


LOCKED_IN_LOOP_EVENT_LEARN_KEEP_AWAY = False
If TIME_RETRY_REDO_NETWORK_ARRAY_VB_UPDATE > 0 Then
    TIME_RETRY_REDO_NETWORK_ARRAY_VB_UPDATE = Now + TimeSerial(0, 0, 30)
End If




End Sub

Sub CHECK_NET_ARRAY()

' Exit Sub ' FOR NOW

Dim R, TxtEXE_INFO, X1_F, NET_NAME_1, NET_NAME_2, NET_NAME_12, NET_NAME_22
Dim F2, VAR_DATE_1, VAR_DATE_2
For R = 1 To UBound(VB_EXE_ARRAY)
    TxtEXE_INFO = VB_EXE_ARRAY(R)
    If InStr(TxtEXE_INFO, "D:\VB6\") > 0 Then
        X1_F = Replace(TxtEXE_INFO, "\VB6\", "\VB6-EXE\")
        Err.Clear
        NET_NAME_1 = TxtEXE_INFO
        NET_NAME_2 = X1_F
        NET_NAME_12 = NET_NAME_1
        NET_NAME_22 = NET_NAME_2
        On Error Resume Next
        Set F1 = FSO.GetFile(NET_NAME_1)
        Set F2 = FSO.GetFile(NET_NAME_2)
        VAR_DATE_1 = 0
        VAR_DATE_2 = 0
        VAR_DATE_1 = F1.DateLastModified
        VAR_DATE_2 = F2.DateLastModified
        If Err.Number > 0 Then
            VAR_DATE_1 = 0
            VAR_DATE_2 = 0
        End If
        If VAR_DATE_1 > 0 Then
            If VAR_DATE_1 > VAR_DATE_2 Then
                NET_COPY_CHANGE_HAPPENER = True
                TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_02 = Now
            End If
            If VAR_DATE_1 < VAR_DATE_2 Then
                NET_COPY_CHANGE_HAPPENER = True
                TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_02 = Now
            End If
        End If
        Set F1 = Nothing
        Set F2 = Nothing
    End If
Next
End Sub


Sub TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_SUB()

Dim V_VAR, V_TIME_01
Dim TIME_MARKER

' VB_EXE_SYNC_TIMER = VB_EXE_SYNC_TIMER + 1

V_VAR = TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_02
V_TIME_01 = Trim(Str(DateDiff("s", V_VAR, Now)))
If Val(V_TIME_01) > -1 Then
    V_TIME_01 = V_TIME_01
    TIME_MARKER = " Sec"
End If
If Val(V_TIME_01) > 300 Then
    V_TIME_01 = Trim(Str(DateDiff("n", V_VAR, Now)))
    TIME_MARKER = " Min"
End If
If Val(V_TIME_01) > 60 And TIME_MARKER = " Min" Then
    V_TIME_01 = Trim(Str(DateDiff("h", V_VAR, Now)))
    TIME_MARKER = " Hour"
End If
If V_VAR < 1 Then
    V_TIME_01 = "0"
    TIME_MARKER = " Sec"
End If
If V_VAR < 1 Then
    If Label_VB_MODIFIED_TIME.Caption = "GS HR" _
    Or Label_VB_MODIFIED_TIME.Caption = "Not YET" Then
        V_TIME_01 = ""
        TIME_MARKER = "Not YET"
    End If
End If

Label_VB_MODIFIED_TIME.Caption = V_TIME_01 + TIME_MARKER
Label_VB_MODIFIED_TIME.FontSize = 9

If VB_EXE_SYNC_TIMER > 0 Then
    V_TIME_01 = Trim(Str(DateDiff("s", Now, VB_EXE_SYNC_TIMER)))
    If InStr(Label_VB_MODIFIED_ACTIVE.Caption, "GO ") = 0 Then
        Label_VB_MODIFIED_ACTIVE.Caption = "IDLE " + V_TIME_01
    End If
End If

End Sub

Sub ADD_TO_ListView_VB_MODIFIED_ERROR(VAR_STRING_10, VAR_STRING_20, ID_VAR, Err_Number, Err_Description)
    
    Dim WANT_HERE, RW, LV3
    Dim VAR_STRING_01, VAR_STRING_02
    Dim COMPARE_01, COMPARE_02
    Dim LABLE_TEXT
    Dim TESTER_STRING
    Dim SET_GO
    Dim COUNTER_CODE
    
    ListView_VB_MODIFIED_ERROR.Refresh
    
    If ID_VAR = 1 Then
        COUNTER_CODE = Format(TOTAL_WORK_COUNT_02, "00") + " of " + Format(TOTAL_WORK_COUNT_01 * (Copier_Done_COUNTER_02 + 1), "00")
        VAR_STRING_01 = Mid(VAR_STRING_10, 3, InStr(3, VAR_STRING_10, "\") - 3)
        VAR_STRING_02 = Mid(VAR_STRING_20, 3, InStr(3, VAR_STRING_20, "\") - 3)
        VAR_STRING_02 = "-->> " + VAR_STRING_02 + " -- " + Mid(VAR_STRING_20, InStr(3, VAR_STRING_20, "drive\") + Len("drive\") - 1)
        
        If ListView_VB_MODIFIED_ERROR.ListItems.Count > 0 Then
            ListView_VB_MODIFIED_ERROR.ListItems(1).SubItems(1) = VAR_STRING_01
            ListView_VB_MODIFIED_ERROR.ListItems(1).SubItems(2) = VAR_STRING_02
            ListView_VB_MODIFIED_ERROR.ListItems(1).Text = "COPIER _ " + COUNTER_CODE
        End If
        
        If ListView_VB_MODIFIED_ERROR.ListItems.Count = 0 Then
            With ListView_VB_MODIFIED_ERROR
                Set LV3 = .ListItems.Add(, , "COPIER _ " + COUNTER_CODE)
                LV3.SubItems(1) = VAR_STRING_01
                LV3.SubItems(2) = VAR_STRING_02
            End With
        End If
        If ListView_VB_MODIFIED_ERROR.ListItems.Count <> OLD_ListView_VB_MODIFIED_ERROR_ListItems_Count Then
            Call LV_AutoSizeColumn(ListView_VB_MODIFIED_ERROR, Nothing)
            'Call LV_AutoSizeColumn(ListView_VB_MODIFIED_ERROR, ListView_VB_MODIFIED_ERROR.ColumnHeaders.Item(1))
            'Call LV_AutoSizeColumn(ListView_VB_MODIFIED_ERROR, ListView_VB_MODIFIED_ERROR.ColumnHeaders.Item(2))
            'ListView_VB_MODIFIED_ERROR.ColumnHeaders.Add , "PATH_AND_FILE", "PATH_AND_FILE", ListView_VB_MODIFIED_ERROR.width - ListView_VB_MODIFIED_ERROR.ColumnHeaders.Item(1).width - ListView_VB_MODIFIED_ERROR.ColumnHeaders.Item(2).width, lvwColumnLeft
        End If
        ListView_VB_MODIFIED_ERROR.Refresh
        
        Exit Sub
    End If
    
    If ID_VAR = 0 Then
        Label_VB_MODIFIED_TIME.BackColor = RGB(200, 127, 127)
        LABLE_TEXT = Label_VB_MODIFIED_TIME.Caption
        LABLE_TEXT = Mid(LABLE_TEXT, 1, InStr(LABLE_TEXT, " "))
        If IsNumeric(Mid(LABLE_TEXT, 1, InStr(LABLE_TEXT, " "))) = False Or LABLE_TEXT = "0 " Then
            Label_VB_MODIFIED_TIME.Caption = "Net Err"
        End If
    End If
    
    ' TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_01 = Now + TimeSerial(0, 1, 0)
    
    If ID_VAR = 0 Then TESTER_STRING = "ERROR"
    If ID_VAR = 2 Then TESTER_STRING = "GOOD SYNC"
    If ID_VAR = 3 Then TESTER_STRING = "VERIFY ERR"
                
    If ID_VAR = 0 Or ID_VAR = 3 Then
        TIME_RETRY_REDO_NETWORK_ARRAY_VB_UPDATE = Now + TimeSerial(0, 0, 30)
    End If
    
    VAR_STRING_01 = Mid(VAR_STRING_10, 1, InStr(3, VAR_STRING_10, "\") - 1)
    VAR_STRING_02 = Mid(VAR_STRING_10, InStr(3, VAR_STRING_10, "drive\") + Len("drive\") - 1)
    COMPARE_01 = VAR_STRING_01 + VAR_STRING_02
    WANT_HERE = True
    If ListView_VB_MODIFIED_ERROR.ListItems.Count > 1 Then
        For RW = 2 To ListView_VB_MODIFIED_ERROR.ListItems.Count
            COMPARE_02 = ListView_VB_MODIFIED_ERROR.ListItems(RW).SubItems(1) + ListView_VB_MODIFIED_ERROR.ListItems(RW).SubItems(2)
            
            ' ------------------------
            ' THE COMPARE NEVER WARN
            ' ------------------------
            ' IT DO SOMETIME
            ' ESPECIALLY WITH OWN CODE
            ' ------------------------
            ' IT DO WHEN Copier_Done_COUNTER_02 > 0
            ' ------------------------
            
            If COMPARE_01 = COMPARE_02 Then
                WANT_HERE = False
                If Copier_Done_COUNTER_02 > 0 And ID_VAR = 2 Then
                    TESTER_STRING = TESTER_STRING + "+" + Trim(Str(Copier_Done_COUNTER_02 + 1))
                End If
                If TESTER_STRING = "ERROR" Then
                    TESTER_STRING = TESTER_STRING + Str(Err_Number)
                End If
                ListView_VB_MODIFIED_ERROR.ListItems(RW).SubItems(1) = VAR_STRING_01
                ListView_VB_MODIFIED_ERROR.ListItems(RW).SubItems(2) = VAR_STRING_02
                ListView_VB_MODIFIED_ERROR.ListItems(RW).SubItems(3) = Err_Description
                ListView_VB_MODIFIED_ERROR.ListItems(RW).Text = TESTER_STRING

            End If
        Next
    End If

    If WANT_HERE = True Then
        With ListView_VB_MODIFIED_ERROR
            If Copier_Done_COUNTER_02 > 0 And ID_VAR = 2 Then
                TESTER_STRING = TESTER_STRING + "+" + Trim(Str(Copier_Done_COUNTER_02 + 1))
            End If
            If TESTER_STRING = "ERROR" Then
                TESTER_STRING = TESTER_STRING + Str(Err_Number)
            End If
            Set LV3 = .ListItems.Add(, , TESTER_STRING)
            LV3.SubItems(1) = VAR_STRING_01
            LV3.SubItems(2) = VAR_STRING_02
            LV3.SubItems(3) = Err_Description
        End With
        ListView_VB_MODIFIED_ERROR.Visible = True
        Call Form_Resize
    End If
    
    If Copier_Done_COUNTER_02 = 1 Then
        For RW = 1 To ListView_VB_MODIFIED_ERROR.ListItems.Count
            If ListView_VB_MODIFIED_ERROR.ListItems(RW).Text = "GOOD SYNC" Then
                ListView_VB_MODIFIED_ERROR.ListItems(RW).Text = "GOOD SYNC+1"
            End If
        Next
    End If
    
    If ListView_VB_MODIFIED_ERROR.ListItems.Count <> OLD_ListView_VB_MODIFIED_ERROR_ListItems_Count Then
        Call LV_AutoSizeColumn(ListView_VB_MODIFIED_ERROR, ListView_VB_MODIFIED_ERROR.ColumnHeaders.Item(1))
        Call LV_AutoSizeColumn(ListView_VB_MODIFIED_ERROR, ListView_VB_MODIFIED_ERROR.ColumnHeaders.Item(2))
        ListView_VB_MODIFIED_ERROR.ColumnHeaders.Add , "PATH_AND_FILE", "PATH_AND_FILE", ListView_VB_MODIFIED_ERROR.width - ListView_VB_MODIFIED_ERROR.ColumnHeaders.Item(1).width - ListView_VB_MODIFIED_ERROR.ColumnHeaders.Item(2).width, lvwColumnLeft
'        ListView_VB_MODIFIED_ERROR.Refresh
    End If
    OLD_ListView_VB_MODIFIED_ERROR_ListItems_Count = ListView_VB_MODIFIED_ERROR.ListItems.Count
    
    Call MARK_LABEL_RED_ERROR
    ListView_VB_MODIFIED_ERROR.Refresh
    
End Sub

Sub MARK_LABEL_RED_ERROR()
    ' TIME_RETRY_ERROR_NETWORK_ARRAY_VB_UPDATE_01 > 0 Then
    
    Dim SET_GO, RW
        
    SET_GO = True
    For RW = 1 To ListView_VB_MODIFIED_ERROR.ListItems.Count
        If InStr(ListView_VB_MODIFIED_ERROR.ListItems(RW).Text, "ERROR") > 0 Then
            Label_VB_MODIFIED_TIME.BackColor = RGB(200, 127, 127) ' RED ERROR
            SET_GO = False
        End If
        If ListView_VB_MODIFIED_ERROR.ListItems(RW).Text = "VERIFY ERR" Then
            Label_VB_MODIFIED_TIME.BackColor = RGB(200, 127, 127) 'YELLOW WARNING
            SET_GO = False
        End If
    Next
    If SET_GO = True Then
        Label_VB_MODIFIED_TIME.BackColor = RGB(255, 255, 255)
    End If
End Sub


Public Function DirExist(sPath As String) As Boolean
    DirExist = (Dir(sPath, vbDirectory) <> "")
End Function

Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Function
Function NotAlwaysOnTop(ByVal hWnd As Long)
    Dim Flags
    SetWindowPos hWnd, hWnd_NOTOPMOST, 0, 0, 0, 0, Flags
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







Sub RUN_BLOCK_OF_SETUP_CONTROLS_POSTION()


Dim SETTING_WIDTH_LISTVIEW
Dim Label50_Width_VAR

TOP_HEIGHT = 20

'MATCH WING DING OUTER BIGGER
'OUTER BIGGER
'----------------------------
Label46.Top = TOP_HEIGHT
Label50.Top = TOP_HEIGHT
Label52.Top = TOP_HEIGHT

Label52.height = 400 'Label51_Height '+ 10 '350 + 10 'Label51.Height + 50
'Label52.width = Label51_Height
'INNER SMALLER
Dim ADJUST_LEFT_OFFSET

' ----------------------------------
' SCREENSIZE
' GetComputerName = "1-ASUS-X5DIJ"
' 1024 X 768
' GetComputerName = "2-ASUS-EEE"
' 1024 X 600
' GetComputerName = "4-ASUS-GL522VW"
' 1920 X 1080
' GetComputerName = "7-ASUS-GL522VW"
' 1920 X 1080
' ----------------------------------


SETTING_WIDTH_LISTVIEW = False
If GetComputerName = "1-ASUS-X5DIJ" Then
' ----------------------------------
    ' GetComputerName = "1-ASUS-X5DIJ"
    ' 1024 X 768
    ' SCREENSIZE
' ----------------------------------
    SETTING_WIDTH_LISTVIEW = True
    ' LABEL46.CAPTION = Process Counter Is:
    Label46.width = 1700 - 140
    Label53.FontSize = 9
    Label46.FontSize = Label53.FontSize
    Label50.FontSize = Label53.FontSize
    'INNER WINGDING
    'Label51.FontSize = Label50.FontSize
    Label46.FontSize = Label53.FontSize
End If

If GetComputerName = "2-ASUS-EEE" Then
' ----------------------------------
    ' GetComputerName = "2-ASUS-EEE"
    ' 1024 X 600
    ' SCREENSIZE
' ----------------------------------
    SETTING_WIDTH_LISTVIEW = True
    ' LABEL46.CAPTION = Process Counter Is:
    Label46.width = 1700 - 140

    Label53.FontSize = 9
    Label46.FontSize = Label53.FontSize
    Label50.FontSize = Label53.FontSize
    'INNER WINGDING
    'Label51.FontSize = Label50.FontSize
    Label46.FontSize = Label53.FontSize

End If

If SETTING_WIDTH_LISTVIEW = False Then
    SETTING_WIDTH_LISTVIEW = True
    ' LABEL46.CAPTION = Process Counter Is:
    Label46.width = 2700
    'Process Set Enumerate Event
    Label53.FontSize = 11
    Label46.FontSize = 15
    Label50.FontSize = Label46.FontSize
    'INNER WINGDING
    'Label51.FontSize = Label50.FontSize
    Label46.FontSize = Label46.FontSize
    '------------------------------------
End If

ADJUST_LEFT_OFFSET = 50

' LABEL51.CAPTION = IS THE WINGDING FRAME
' Label51_Here
Label51.Left = (Label46.Left + Label46.width) + ADJUST_LEFT_OFFSET

'OUTER BIGGER
'Label52.Left = (Label46.Left + Label46.width) + 40
'Label52.width = 500 'Label51.width + ADJUST_LEFT_OFFSET + 20
'WINGDING FONT SIZE
'----------------------------

'COUNTER
'Label50.Left = (Label52.Left + Label52.width) + 20



TxtEXE.Top = TOP_HEIGHT
Label_Goto_File_Name.Top = TOP_HEIGHT

Label46.Top = TOP_HEIGHT
Label46.ToolTipText = "Hitt Me To Process Enumerate"
Label50.ToolTipText = "Hitt Me To Process Enumerate"
Label51.ToolTipText = "Hitt Me To Process Enumerate"
'Label53.ToolTipText = "Hitt Me To Process Enumerate"

'PROCESS COUNTING
'Label46_HERE
'Label46.Width _ DO IT UP HIGHER AT Label51.Left

'Label46.height = Label10.height

Label46.height = Label52.height
'Label50_here
Label50.height = Label46.height

'PROCESS COUNTER LABEL INFO NAME
'-------------------------------
Label46.Left = TxtEXE.Left + TxtEXE.width + 50 + 10
'LINK
Label53.Left = Label46.Left
Label52.width = 500 'Label51.width + ADJUST_LEFT_OFFSET + 20
Label52.Left = Label46.Left + Label46.width + 20
Label50.Left = Label52.Left + Label52.width + 40



Label46.Top = TOP_HEIGHT
'Label46.Height = Label52.Height
'-------------------------------
'PROCESS COUNTER LABEL INFO % NUMERIC
'------------------------------------
Label50.Top = TOP_HEIGHT
Label50.height = Label46.height
'------------------------------------


'Label46.Height = Label50.Height
Label46.Top = TOP_HEIGHT
Label52.Top = TOP_HEIGHT

'Process Enumerate Time Happen
Label53.Top = Label52.Top + Label52.height + 40 ' + 20
Label53.height = 290

Label46.Top = TOP_HEIGHT
Label50.Top = TOP_HEIGHT
Label52.Top = TOP_HEIGHT

Label51.Top = TOP_HEIGHT + 20


'Label51.Height = Label50.Height
'Label52.Height = Label51.Height
Label52 = ""

' LAB_MAXIMIZE_GOOD_SYNC.Left = Label53.Left - 20

lstProcess_2_ListView.Left = Label53.Left - 20 '+ 1 ' + 1

SETTING_WIDTH_LISTVIEW = False

If GetComputerName = "1-ASUS-X5DIJ" Then
    SETTING_WIDTH_LISTVIEW = True
    ' lstProcess_2_ListView.Font.Size = Default = 8.4
    lstProcess_2_ListView.Font.size = 7.4
    lstProcess_3_SORTER_ListView.Font.size = lstProcess_2_ListView.Font.size
    'lstProcess_2_ListView_HERE
    lstProcess_2_ListView.width = 2400

End If

If GetComputerName = "2-ASUS-EEE" Then
    SETTING_WIDTH_LISTVIEW = True
    ' lstProcess_2_ListView.Font.Size = Default = 8.4
    lstProcess_2_ListView.Font.size = 7.4
    lstProcess_3_SORTER_ListView.Font.size = lstProcess_2_ListView.Font.size
    'lstProcess_2_ListView_HERE
    lstProcess_2_ListView.width = 2280

End If

If SETTING_WIDTH_LISTVIEW = False Then
    SETTING_WIDTH_LISTVIEW = True
    ' lstProcess_2_ListView.Font.Size = Default = 8.4
    lstProcess_2_ListView.Font.size = 8.4
    lstProcess_3_SORTER_ListView.Font.size = lstProcess_2_ListView.Font.size
    'lstProcess_2_ListView_HERE
    lstProcess_2_ListView.width = 3500
End If

lstProcess_3_SORTER_ListView.width = lstProcess_2_ListView.width
lstProcess_3_SORTER_ListView.Left = lstProcess_2_ListView.Left + lstProcess_2_ListView.width + 10

'-----------------------------------------
'lstProcess_2_ListView_HERE
'FINAL HEIGHT IS LATER DOWN A BIT
'-------------------------------------------------------
Call WIDTH_AND_HEIGHT(WX, HY)

lstProcess_2_ListView.height = HY - lstProcess_2_ListView.Top + 80
lstProcess_3_SORTER_ListView.height = lstProcess_2_ListView.height


'Label53.Caption = "Process Set Enumerate Event"


'Label44.Visible = False
'Label44.Left = Label46.Left + 5
'Label44.width = lstProcess_2_AND_lstProcess_3_WIDTH_POSTION_WITH_ER_OFFSET_LEFT - 10
'
'Label44.FontSize = 7
''Label44_HERE
'Label44.height = 580 '+ 10 '+ 10
'Label44.Caption = "LISTVIEW && CLICK_ER DONT LKE RUNNER IN THE IDE CRASH &WITH_ER  FOR NOW _ CODE PERFECT PROOF READ_ER __ Feeling to Do With ListView.OCX Version Fault Investigate _ How Annoying _ Code Inside IDE _ Check Error Outside IDE Compiled EXE"

'Label53_Here
'Label53.width = lstProcess_2_AND_lstProcess_3_WIDTH_POSTION_WITH_ER_OFFSET_LEFT - 10
'Label53.Caption = "Process Set Enumerate Event"

'PROCESS COUNTING
'Label50_here
'Label50_Width_VAR = Label53.width - (Label46.width + Label52.width) - 30
'If Label50_Width_VAR < 100 Then Label50_Width_VAR = 500

'Label50.width = Label50_Width_VAR
' SET THIS HIGHER WHEN LISTVIEW BOX SETUP AS NOT YET

'Label50.width = 3000
'----------------------
'----------------------
'----------------------
'----------------------
'END OF FORM LOAD BLOCK
'----------------------
'----------------------
'----------------------
'----------------------

'---------------------------------------------------
' Call MNU_ADMINISTRATOR_Click

' Call MNU_SHOW_THE_TIME_FORCE
' Call MNU_SHOW_THE_TIME_Click
' Call PUT_GIVE_ME_TIME_IN_MENU

'Me.Left = 800


'TOP
'lstProcess_2_ListView_HERE
lstProcess_2_ListView.Top = Label53.Top + Label53.height + 10
'TOP
lstProcess_3_SORTER_ListView.Top = lstProcess_2_ListView.Top

Dim lstProcess_2_AND_lstProcess_3_WIDTH_POSTION_WITH_ER_OFFSET_LEFT
lstProcess_2_AND_lstProcess_3_WIDTH_POSTION_WITH_ER_OFFSET_LEFT = lstProcess_2_ListView.width + lstProcess_3_SORTER_ListView.width + (10 - 20) - 10
lstProcess_2_AND_lstProcess_3_WIDTH_POSTION_WITH_ER_OFFSET_LEFT = (lstProcess_3_SORTER_ListView.Left + lstProcess_3_SORTER_ListView.width) - lstProcess_2_ListView.Left

'Label53_Here
Label53.width = lstProcess_2_AND_lstProcess_3_WIDTH_POSTION_WITH_ER_OFFSET_LEFT - 10


Dim lstProcess_3_WIDTH_POSTION_LEFT_AND_WIDTH
lstProcess_3_WIDTH_POSTION_LEFT_AND_WIDTH = lstProcess_3_SORTER_ListView.Left + lstProcess_3_SORTER_ListView.width
Label50_Width_VAR = lstProcess_3_WIDTH_POSTION_LEFT_AND_WIDTH - (Label52.Left + Label52.width) - 40
Label50.width = Label50_Width_VAR

'ListView_VB_MODIFIED_ERROR.Top = Label_GOODSYNC_01.Top + Label_GOODSYNC_01.height + 20
'ListView_VB_MODIFIED_ERROR.Top = ListView_CPU_INFO.Top + ListView_CPU_INFO.height + 20
ListView_VB_MODIFIED_ERROR.Top = Label_VB_MODIFIED_ACTIVE.Top + Label_VB_MODIFIED_ACTIVE.height + 20
ListView_VB_MODIFIED_ERROR.Left = ListView_CPU_INFO.Left
'FORM_LOAD
'Me.Height = HY2
'Me.Width = WX2

'Me.Top = Screen.Height / 2 - Me.Height / 2 - 800
'Me.Left = Screen.Width / 2 - Me.Width / 2 '- 400

'--------------------------------------------------------
'END OF FORM LOAD
'TRY SHOW
'Me.Show
'DoEvents
'--------------------------------------------------------
'NOT CALL THIS PART IN FORM LOAD UNLESS BEEN SHOW SOMEWAY
'OR NOT ALL CONTOL ARE VISIBLE THAT REQUIRE
'--------------------------------------------------------
'Call WIDTH_AND_HEIGHT(WX, HY)
'RESIZE HAPPEN

'Timer_EnumProcess.Enabled = True
'Timer_EnumProcess.Interval = 1000

' If Command$ <> "" Then Me.WindowState = vbMinimized

Call IsInternetConnected

MNU_KILL_MAX_AHK.Caption = "KILL AHK PID LIMITER > 100=TRUE"
MNU_CMD_KILL_MAX.Caption = "KILL CMD PID LIMITER > 20=TRUE"

Call MNU_GIVE_ME_TIME_WITHER_UTC_Click

Call SET_CONTROLS_COUNT
Call SET_MENU_PADD_WORK


Call SET_LABEL_PADD_WORK
'----------------------
'----------------------
'----------------------
'----------------------
'END OF FORM LOAD BLOCK
'----------------------
'----------------------
'----------------------
'----------------------
'If IsIDE = True Then Me.WindowState = vbNormal
'If IsIDE = False Then Me.WindowState = vbMinimized
    
'__ Sub Timer_ALWAYS_ON_TOP_TO_START_WITH_ER_Timer()
SetWindowPos Me.hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

' Command_Screen_Shot_Auto_ClipBoard_er.Caption = "Screen Shot Auto ClipBoard_er when Spy_er && Archive Mode _OFF" + vbCrLf + "Hitt Button Here to Change"

'Call CAMERA_FOLDER_TO_MATCH_WITH_DATE_FOR_WIFI_SDCF


'Dim R As Long, ATEST
'For R = 1 To 255
'If GetSpecialfolder(R) <> "" Then
'ATEST = ATEST + 1
'frmAbout.List1.AddItem Format(ATEST, "000") + "  " + Format(R, "000") + "__" + GetSpecialfolder(R)
'End If
'Next
'
'frmAbout.Show

'---------------------------------
'Label15.Caption = "Scan Processor Quicker For Some Moment and Also ForeGound Window hWnd Handle Change Massive esponce Speed Up That Method __ __ __ __ Soon Next Process Logger With Full Path Name"
'Label15.Caption = "Scan Processor For A Moment Same As Other Command Set _ " + Trim(Str(EnumProcess_COUNTER))
'Label15.Caption = "Scan Processor For A Moment Timer Second _ " + Trim(Str(EnumProcess_COUNTER)) + " Plus ForeGound Window Change"
Call Timer_EnumProcess_Timer
'---------------------------------




End Sub

Sub SET_MENU_PADD_WORK()

Dim i_Menu_Count, i_Form_Counter
Dim i_Menu_Not_Visa_Count

Dim Control As Control, Label_44, LABEL_48

Dim R_NEXT

Dim Text_Checker_Form_Menu As String

Dim MENU_ITEM_VAR
Dim i

For i = 0 To Forms.Count - 1
    
    For Each Control In Forms(i).Controls
        If InStr(UCase(Control.Name), "MNU_") > 0 Then
            If Control.Visible = True Then
                i_Menu_Count = i_Menu_Count + 1
            End If
            i_Menu_Not_Visa_Count = i_Menu_Not_Visa_Count + 1

        End If
    Next
Next

i_Menu_Not_Visa_Count = i_Menu_Not_Visa_Count - i_Menu_Count

Mnu_Menu_Item_Count.Caption = "Menu Item Count = " + Str(i_Menu_Count) + " &&" + Str(i_Menu_Not_Visa_Count) + " Not Visible"
'Mnu_Form_Count.Caption = "Form Counter = " + Str(Forms.Count - 1) '  + " Really 7"
'Mnu_Form_Count.Visible = False

i_Menu_Count = 0



For i = 0 To Forms.Count - 1
    Text_Checker_Form_Menu = ""
    frmListMenu.GetMenuInfo_Not_Indented GetMenu(Forms(i).hWnd), 0, "", Text_Checker_Form_Menu
    Text_Checker_Form_Menu = UCase(Text_Checker_Form_Menu)
    For Each Control In Forms(i).Controls
        If InStr(UCase(Control.Name), "MNU_") > 0 Then
            MENU_ITEM_VAR = Replace(Control.Caption, "[__ ", "")
            MENU_ITEM_VAR = Replace(MENU_ITEM_VAR, " __]", "")
            MENU_ITEM_VAR = UCase(Trim(MENU_ITEM_VAR))
            If InStr(Text_Checker_Form_Menu, "SUB MENU ----" + MENU_ITEM_VAR) = 0 Then
                
                'i_Menu_Count = i_Menu_Count + 1
                If InStr(Trim(Control.Caption), "[__ ") = 0 Then
                    Label_44 = Trim(Control.Caption)
                    'LABEL_48 = Replace(LABEL_44, " ", "_")
                    LABEL_48 = Label_44
                    LABEL_48 = Replace(LABEL_48, "___", "__")
                    LABEL_48 = "[__ " + LABEL_48 + " __]"
                    LABEL_48 = Replace(LABEL_48, "[__ [__ ", "[__ ")
                    LABEL_48 = Replace(LABEL_48, " __] __]", " __]")
                    If LABEL_48 <> Label_44 Then
                        Control.Caption = LABEL_48
                    End If
                End If
            End If
        End If
    Next
Next

'Stop

''MNU_BRing_Front
''---------------
'i_Form_Counter = Forms.Count
'i_Form_Counter = 0
''for each f
''For i = 0 To Forms.Count - 1
''    Load Forms(i)
''    Show Forms(i)
''Next
'
'Dim Form As Form
'i_Form_Counter = 0
'For Each Form In Forms
'    i_Form_Counter = i_Form_Counter + 1
'    Load Form
'    Form.Show
'    Show Form
'    'Set Form = Nothing
'Next Form
'
'i_Form_Counter = 0
'For Each Form In Forms
'    i_Form_Counter = i_Form_Counter + 1
'    Load Form
'    Form.Show
'    'Set Form = Nothing
'Next Form

i_Form_Counter = Forms.Count - 1
Me.Refresh
End Sub


Sub SET_CONTROLS_COUNT()

Dim Control As Control, Label_44, LABEL_48

Dim R_NEXT

Dim Text_Checker_Form_Menu As String

Dim i

Dim CONTROL_THIS_FORM_V_
Dim CONTROL_THIS_FORM_V_PLUS
Dim CONTROL_THIS_FORM_VN
Dim CONTROL_THIS_FORM_TOT_V_
Dim CONTROL_THIS_FORM_TOT_VN
Dim CONTROL_THIS_FORM_TOT_V__PLUS

CONTROL_THIS_FORM_V_ = 0
CONTROL_THIS_FORM_V_PLUS = 0
CONTROL_THIS_FORM_VN = 0
CONTROL_THIS_FORM_TOT_V_ = 0
CONTROL_THIS_FORM_TOT_VN = 0
CONTROL_THIS_FORM_TOT_V__PLUS = 0

Dim XXER

On Error Resume Next

For Each Form In Forms
    For Each Control In Form.Controls
        Err.Clear
        XXER = Control.Visible
        If Err.Number > 0 Then XXER = False
        If XXER = True Then
            If Form.hWnd = Me.hWnd Then
            CONTROL_THIS_FORM_V_ = CONTROL_THIS_FORM_V_ + 1
            End If
            CONTROL_THIS_FORM_TOT_V_ = CONTROL_THIS_FORM_TOT_V_ + 1
        Else
            If Form.hWnd = Me.hWnd Then
                CONTROL_THIS_FORM_VN = CONTROL_THIS_FORM_VN + 1
            End If
                CONTROL_THIS_FORM_TOT_VN = CONTROL_THIS_FORM_TOT_VN + 1
        End If
    Next
Next
On Error GoTo 0

CONTROL_THIS_FORM_TOT_V__PLUS = CONTROL_THIS_FORM_TOT_V_ + CONTROL_THIS_FORM_TOT_VN
CONTROL_THIS_FORM_V_PLUS = CONTROL_THIS_FORM_V_ + CONTROL_THIS_FORM_VN

' Mnu_Control_Item_Count.Caption = "Control Item Count All Form*" + Trim(Str(Forms.Count)) + " = " + Str(CONTROL_THIS_FORM_TOT_V__PLUS) + " &&" + Str(CONTROL_THIS_FORM_TOT_VN) + " Not Visible && Main Form" + Str(CONTROL_THIS_FORM_V_PLUS) + " &&" + Str(CONTROL_THIS_FORM_VN) + " Not Visible"
Mnu_Control_Item_Count.Caption = "Control Item Count All Form*" + Trim(Str(Forms.Count)) + " =" + Str(CONTROL_THIS_FORM_TOT_V__PLUS) + " &&" + Str(CONTROL_THIS_FORM_TOT_VN) + " Not Visible &&" + Str(CONTROL_THIS_FORM_V_PLUS) + " This Form" '+ " &&" + Str(CONTROL_THIS_FORM_VN) + " Not Visible"
End Sub



Sub SET_LABEL_PADD_WORK()

Dim Control As Control
Dim i

On Error Resume Next

For i = 0 To Forms.Count - 1
    For Each Control In Forms(i).Controls
        If InStr(Trim(Control.Caption), "/n") > 0 Then
            Control.Caption = Replace(Control.Caption, "/n", vbCrLf)
        End If
    Next
Next

Me.Refresh
End Sub



Private Sub Lab_KILL_EXPLORER_Click()

Call MNU_TASK_KILLER_EXPLORER_CLIPBOARD_Click

End Sub

Private Sub MNU_TASK_KILLER_EXPLORER_CLIPBOARD_Click()
    
'PASTE THE CURRENT SESSION TO CLIPBOARD __ QUITELY WITHOUT MSGBOX REPSONCE REQUIRING
'----------------------------------------------------
Call FindWindow_Get_All_Explorer("QUITE MSGBOX=TRUE")
'----------------------------------------------------

'PROCESS_TO_KILLER_TO_GO = "/F /IM EXPLORER* /T"
'PROCESS_TO_KILLER = PROCESS_TO_KILLER_TO_GO

Label23.BackColor = Label11.BackColor
    
Beep
Me.WindowState = vbMinimized
'Shell "CMD /C START """" /REALTIME ""C:\SCRIPTER\SCRIPTER CODE -- BAT\ELITESPY - TASKKILL BAT\TASKKILLER_EXPLORER.BAT"" " + PROCESS_TO_KILLER, vbMaximizedFocus
Beep
Dim objShell
Set objShell = CreateObject("Wscript.Shell")

'Shell "TASKKILL /F /IM ""EXPLORER*"""
objShell.Run "TASKKILL /F /IM ""EXPLORER*""", 0, True
    
Dim XNOW, FW_P
XNOW = Now + TimeSerial(0, 0, 4)
Do
    FW_P = FindWindow("Progman", "Program Manager")
Loop Until FW_P = 0 Or XNOW < Now

objShell.Run "C:WINDOWS\EXPLORER.EXE", 0, False
Set objShell = Nothing
    
Me.WindowState = vbMinimized
    
End Sub


Function FindWindow_Get_All_Explorer(QUITE_MSGBOX) As String

FindWindow_Get_All_Explorer = ""

Dim Huge, VAR_STRING
Dim WINDOW_TITLE
Dim test_hWnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String
Huge = 0

test_hWnd = FindWindow2(ByVal 0&, ByVal 0&)
VAR_STRING = ""
Do While test_hWnd <> 0
    
    WINDOW_TITLE = GetWindowTitle(test_hWnd)
    
    '--------------------------------------------------
    'C:\Windows\explorer.exe
    '--------------------------------------------------
    If GetWindowClass(test_hWnd) = "CabinetWClass" Then
        Huge = Huge + 1
        VAR_STRING = VAR_STRING + WINDOW_TITLE + vbCrLf + vbCrLf
    End If
        
    'retrieve the next window
    test_hWnd = GetWindow(test_hWnd, GW_hWndNEXT)

Loop

Dim I_1 As String

I_1 = Trim(Str(Huge)) + " __ COUNT __" + vbCrLf
I_1 = I_1 + "C:\Windows\explorer.exe" + vbCrLf
I_1 = I_1 + "FIND ON CLASS NAME __CabinetWClass__" + vbCrLf + vbCrLf
I_1 = I_1 + "EXPLORER SESSION CURRENTLY WORKING ON MAYBE REBOOT OR TASKKILL" + vbCrLf
I_1 = I_1 + "EXPLORER Window SET FOUND PASTE CLIPBOARD" + vbCrLf
I_1 = I_1 + "-----------------------------------------" + vbCrLf + vbCrLf
I_1 = I_1 + VAR_STRING

Clipboard.Clear
Clipboard.SetText VAR_STRING
FindWindow_Get_All_Explorer = I_1

If QUITE_MSGBOX = "QUITE MSGBOX=FALSE" Then MsgBox I_1

End Function




Private Sub MNU_TASK_KILLER_CMD_Click()
    
PROCESS_TO_KILLER_TO_GO = "/F /IM CMD* /T"
PROCESS_TO_KILLER = PROCESS_TO_KILLER_TO_GO

' Label23.BackColor = Label11.BackColor
    
Beep
Me.WindowState = vbMinimized
Shell "CMD /C START """" /REALTIME ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT_03_PROCESS_KILLER.BAT"" " + PROCESS_TO_KILLER, vbMaximizedFocus
Beep
    
End Sub


Private Sub Label_1X_Click()

VAR_LAB_TEXT = Label_1X.Caption

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_1X.BackColor = RGB(255, 255, 255)

SET_COMPUTER_TO_RUN = FIND_COMPUTER_TO_RUN(VAR_LAB_TEXT)

'Call CREATE_PATH_ARRAY_SET_NETWORK_ALL_SPEICAL_REQUEST

End Sub

Private Sub Label_2E_Click()

VAR_LAB_TEXT = Label_2E.Caption

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_2E.BackColor = RGB(255, 255, 255)

SET_COMPUTER_TO_RUN = FIND_COMPUTER_TO_RUN(VAR_LAB_TEXT)

'Call CREATE_PATH_ARRAY_SET_NETWORK_ALL_SPEICAL_REQUEST

End Sub

Private Sub Label_3L_Click()

VAR_LAB_TEXT = Label_3L.Caption

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_3L.BackColor = RGB(255, 255, 255)

SET_COMPUTER_TO_RUN = FIND_COMPUTER_TO_RUN(VAR_LAB_TEXT)

'Call CREATE_PATH_ARRAY_SET_NETWORK_ALL_SPEICAL_REQUEST

End Sub

Private Sub Label_4G_Click()

VAR_LAB_TEXT = Label_4G.Caption

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_4G.BackColor = RGB(255, 255, 255)

SET_COMPUTER_TO_RUN = FIND_COMPUTER_TO_RUN(VAR_LAB_TEXT)

'Call CREATE_PATH_ARRAY_SET_NETWORK_ALL_SPEICAL_REQUEST

End Sub

Private Sub Label_5P_Click()

VAR_LAB_TEXT = Label_5P.Caption

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_5P.BackColor = RGB(255, 255, 255)

SET_COMPUTER_TO_RUN = FIND_COMPUTER_TO_RUN(VAR_LAB_TEXT)

'Call CREATE_PATH_ARRAY_SET_NETWORK_ALL_SPEICAL_REQUEST

End Sub

Private Sub Label_7G_Click()

VAR_LAB_TEXT = Label_7G.Caption

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_7G.BackColor = RGB(255, 255, 255)

SET_COMPUTER_TO_RUN = FIND_COMPUTER_TO_RUN(VAR_LAB_TEXT)

'Call CREATE_PATH_ARRAY_SET_NETWORK_ALL_SPEICAL_REQUEST

End Sub

Private Sub Label_8M_Click()

VAR_LAB_TEXT = Label_8M.Caption

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_8M.BackColor = RGB(255, 255, 255)

SET_COMPUTER_TO_RUN = FIND_COMPUTER_TO_RUN(VAR_LAB_TEXT)

'Call CREATE_PATH_ARRAY_SET_NETWORK_ALL_SPEICAL_REQUEST

End Sub

Private Sub Label_CHROME_PAGE_AUTO_ON_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_CHROME_PAGE_AUTO_ON.BackColor = RGB(255, 255, 255)

' Const ShowWindow_2 = 1, DontShowWindow = 0, DontWaitUntilFinished = False, WaitUntilFinished = True
Dim objShell
Set objShell = CreateObject("Wscript.Shell")
objShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 79-LOAD URL AT BOOT CHROME.ahk""", DontShowWindow, DontWaitUntilFinished
Set objShell = Nothing

Me.WindowState = vbMinimized

End Sub

Private Sub Label_CLOSE_GOODSYNC2GO_Click()

Call Label_MAXIMIZE_GOODSYNC2GO_Click

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_CLOSE_GOODSYNC2GO.BackColor = RGB(255, 255, 255)

Call MNU_CLOSE_GOODSYNC_2_GO_Click

Me.WindowState = vbMinimized

End Sub

Private Sub Label_CLOSE_OUTLOOK_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_CLOSE_OUTLOOK.BackColor = RGB(255, 255, 255)

Dim hWnd_RESULT
hWnd_RESULT = FindWindow("rctrl_renwnd32", vbNullString)
If hWnd_RESULT > 0 Then
    Result = PostMessage(hWnd_RESULT, WM_CLOSE, 0&, 0&)
End If

Me.WindowState = vbMinimized
Beep

End Sub

Private Sub Label_GOODSYNC_01_Click()
' -------------------------------------------------------------------------------
' SESSION AT [__ Ver_2019_1.0.279 __]
' FOR ROUTINE ---- Private Sub Label_GOODSYNC_01_Click()
' -------------------------------------------------------------------------------
' YOU MUST SET YOUR GOOD_SYNC LOGG TO GO WITH THE CUSTOM DIRECTORY
' THEY ARE STILL KEPT IN USUAL PLACE WITH THE SYNC FOLDER IF UNLESS SELECT NOT TO
' WHEN THEY ARE COMMON HERE
' ABLE TO TELL IF SYNC IS RUN AS GOOD SCHEDULE
' YOU ABEL TO SET A CUSTOM SCRIPT AT END OF ALL JOB TO CHECK THING OVER
' IF WHEN THAT RUN ABLE TO CHECK THING OVER
' IF SOMETHING LIKE HASN'T GROUND TO A HALT
' IF HAS GROUND TO HALT
' WHICH HARD TO TELL
' THE  MOST TIME LOGG ARE UPDATE AND RUN
' BUT THE IS NOT SHOW ANYMORE
' IF THAT WERE CASE
' GET ON TO IT AND WHEN END OF WHOLE SET JOB FINISHER
' AND CHECK THEN
' IF FIND THING NOT GOOD THEN ISSUE A CLOSE TO GOOD_SYNC AND RELOAD
' CURRENTLY I HAVE ONE AT END EVERY FINISH ALL JOB IN SCRIPT AT
' AFTER THEM ALL RUN AT BOTTOM SCRIPT
' --------------------------------------------------------------------------------
' [ Wednesday 02:42:30 Am_20 March 2019 ] WRITING
' --------------------------------------------------------------------------------
' CLIPBOARD COUNTER ---- 98 * CLIPBOARD'S COUNT
' Count = 1123 -- Wed 20-Mar-2019 01:54:15
' Count = 1221 -- Wed 20-Mar-2019 03:08:18
' -------------------------------------------------------------------------------

Dim GS_NAME_01
Dim GS_NAME_02
Dim IRESULT
Dim F
Dim V_VAR
Dim GOOD_SYNC
Dim V_TIME_01
GS_NAME_01 = "D:\0 00 LOGGERS TEXT GOODSYNC\" + GetComputerName + "\"
GS_NAME_02 = "GOODSYNC-*.LOG"
If FSO.FolderExists(GS_NAME_01) = False Then
    IRESULT = CreateFolderTree(GS_NAME_01)
End If
If FSO.FolderExists(GS_NAME_01) = False Then Exit Sub
File_GOODSYNC.Path = GS_NAME_01
File_GOODSYNC.Pattern = GS_NAME_02
File_GOODSYNC.Refresh
If File_GOODSYNC.ListCount > 0 Then
    GOOD_SYNC = File_GOODSYNC.List(File_GOODSYNC.ListCount - 1)
    Label_GOODSYNC_01.Caption = GOOD_SYNC

    If FSO.FileExists(File_GOODSYNC.Path + "\" + GOOD_SYNC) Then
        Set F = FSO.GetFile(File_GOODSYNC.Path + "\" + GOOD_SYNC)
        'V_VAR = FORMAT(F.DateLastModified, "YYYY-MM-DD HH-MM-SS")
        V_VAR = F.DateLastModified
        V_TIME_01 = Trim(Str(DateDiff("n", V_VAR, Now)))
        If Val(V_TIME_02) = 0 Or SET_GOODSYNC_DONE_ONCE = False Then
            V_TIME_02 = V_TIME_01
        End If
        If Val(V_TIME_01) < Val(OLD_V_TIME_01) Or Val(V_TIME_02) = 0 Then
            V_TIME_02 = OLD_V_TIME_01
            SET_GOODSYNC_DONE_ONCE = True
        End If
        OLD_V_TIME_01 = V_TIME_01
        Label_GOODSYNC_02_HOUR.Caption = V_TIME_01 + " Min"
        Label_GOODSYNC_04_HOUR.Caption = V_TIME_02 + " Min"
    End If
End If

End Sub

Private Sub Label_GOODSYNC_02_HOUR_Click()
'Label_GOODSYNC_02_HOUR
End Sub

Private Sub Label_GOODSYNC_04_HOUR_Click()
'Label_GOODSYNC_04_HOUR.caption
End Sub

Private Sub Label_Goto_File_Name_Click()
'Beep

Shell "EXPLORER /e, /SELECT, " + TxtEXE, vbNormalFocus

Me.WindowState = vbMinimized


'txtFile

End Sub

Private Sub Label_HWND_MAXIMIZE_Click()
Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_HWND_MAXIMIZE.BackColor = RGB(255, 255, 255)

Dim hWnd_RESULT
hWnd_RESULT = txthWnd
If hWnd_RESULT > 0 Then
    ShowWindow hWnd_RESULT, SW_MAXIMIZE
    
    Me.WindowState = vbMinimized
End If

End Sub

Private Sub Label_KILL_AND_RUN_ANOTHER_AUTOHOTKEY_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_KILL_AND_RUN_ANOTHER_AUTOHOTKEY.BackColor = RGB(255, 255, 255)

Call KILL_AUTOHOTKEY_GLOBAL

Sleep 1000

Call MNU_AUTOHOTKEYS_SET_Click

End Sub

Private Sub Label_KILL_CMD_AND_AHK_Click()
Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_KILL_CMD_AND_AHK.BackColor = RGB(255, 255, 255)

SET_COMPUTER_TO_RUN_PID_EXE = "Cmd.exe"
If SET_COMPUTER_TO_RUN <> "" Then
    Call CREATE_PATH_ARRAY_SET_NETWORK_ALL_SPEICAL_REQUEST
    SET_COMPUTER_TO_RUN = ""
    Exit Sub
End If

Dim R, A1, A2
Dim ALL_DONE
Dim NAME_EXE As String
Dim PID_INPUT As Long
Dim i
Do
    ALL_DONE = True
    SET_COMPUTER_TO_RUN_PID_EXE = "Cmd.exe"
    For R = 1 To lstProcess_3_SORTER_ListView.ListItems.Count
        A1 = lstProcess_3_SORTER_ListView.ListItems.Item(R).SubItems(1)
        If InStr(A1, SET_COMPUTER_TO_RUN_PID_EXE) > 0 Then
            pid = Val(lstProcess_3_SORTER_ListView.ListItems.Item(R))
            i = cProcesses.Process_Kill(pid)
            ALL_DONE = False
        End If
    Next
    
    SET_COMPUTER_TO_RUN_PID_EXE = "Conhost.exe"
    For R = 1 To lstProcess_3_SORTER_ListView.ListItems.Count
        A1 = lstProcess_3_SORTER_ListView.ListItems.Item(R).SubItems(1)
        If InStr(A1, SET_COMPUTER_TO_RUN_PID_EXE) > 0 Then
            pid = Val(lstProcess_3_SORTER_ListView.ListItems.Item(R))
            i = cProcesses.Process_Kill(pid)
            ALL_DONE = False
        End If
    Next
    
    SET_COMPUTER_TO_RUN_PID_EXE = "AutoHotkey.exe"
    For R = 1 To lstProcess_3_SORTER_ListView.ListItems.Count
        A1 = lstProcess_3_SORTER_ListView.ListItems.Item(R).SubItems(1)
        If InStr(A1, SET_COMPUTER_TO_RUN_PID_EXE) > 0 Then
            pid = Val(lstProcess_3_SORTER_ListView.ListItems.Item(R))
            i = cProcesses.Process_Kill(pid)
            ALL_DONE = False
        End If
    Next
    Call EnumProcess
Loop Until ALL_DONE = True
End Sub

Private Sub Label_KILL_CMD_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_KILL_CMD.BackColor = RGB(255, 255, 255)

SET_COMPUTER_TO_RUN_PID_EXE = "Cmd.exe"
If SET_COMPUTER_TO_RUN <> "" Then
    Call CREATE_PATH_ARRAY_SET_NETWORK_ALL_SPEICAL_REQUEST
    SET_COMPUTER_TO_RUN = ""
    Exit Sub
End If

Dim R, A1, A2

For R = 1 To lstProcess_3_SORTER_ListView.ListItems.Count
    A1 = lstProcess_3_SORTER_ListView.ListItems.Item(R).SubItems(1)
    If InStr(A1, SET_COMPUTER_TO_RUN_PID_EXE) > 0 Then
        pid = Val(lstProcess_3_SORTER_ListView.ListItems.Item(R))
        cProcesses.Process_Kill (pid)
    End If
Next

SET_COMPUTER_TO_RUN_PID_EXE = "Conhost.exe"
For R = 1 To lstProcess_3_SORTER_ListView.ListItems.Count
    A1 = lstProcess_3_SORTER_ListView.ListItems.Item(R).SubItems(1)
    If InStr(A1, SET_COMPUTER_TO_RUN_PID_EXE) > 0 Then
        pid = Val(lstProcess_3_SORTER_ListView.ListItems.Item(R))
        cProcesses.Process_Kill (pid)
    End If
Next

Me.WindowState = vbMinimized

End Sub

Private Sub Label_MAXIMIZE_CLIPBOARD_LOGGER_Click()
Dim WINDOW_hWnd
WINDOW_hWnd = FindWindow("ThunderRT6FormDC", "ClipBoard Logger")
'WINDOW_hWnd = FindWindowPart("EliteSpy+ 2001 __ www.PlanetSourceCode.com __ Version")
Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_MAXIMIZE_CLIPBOARD_LOGGER.BackColor = RGB(255, 255, 255)

ShowWindow WINDOW_hWnd, SW_NORMAL
' BEep
Me.WindowState = vbMinimized
End Sub

Private Sub MAX_FORM_CHECKER()

' --------------------------------------------------------------------
' SOME FORM GET PUSH DOWN OF SCREEN INTO MINUS AREA ON ALL CORDINATE
' SO CHECKER AUTO ANY PUSH LIKE THAT THING BRING THEM CORD INTO SCREEN
' THAT ENABLE WHEN MINIMIZED THE PROGRAM ACTUALLY DO POP UP WHEN WANT
' --------------------------------------------------------------------
' FROM -- TO
' [ Wednesday 08:47:20 Am_12 June 2019 ]
' [ Wednesday 09:30:00 Am_12 June 2019 ]
' --------------------------------------------------------------------

' --------------------------------------------------------------------
' THIS CODE IS RUN BY -- WHICH IS QUITE QUICKER AT 10 MILLISECOND RUN
' ONLY IF FOREGROUND WINDOW CHANGE AND DELAY TIMER NOT TOO QUICK
' --------------------------------------------------------------------
'    Call Timer_FOREGROUND_WINDOW_CHANGE_Timer
'    Call Timer_FOREGROUND_WINDOW_CHANGE_02_Timer
' --------------------------------------------------------------------
' IT IS RUN BY PROJECT ELITESPY AND VB_KEEP_RUNNER
' --------------------------------------------------------------------

Dim WINDOW_hWnd(10)
Dim hWndResult
Dim RectFINDER As RECT
Dim SET_GO, RT, RB, RR, RL
Dim RR_EYE
Dim I_R

WINDOW_hWnd(1) = FindWindowPart("EliteSpy+ 2001 __ www.PlanetSourceCode.com __ Version")
WINDOW_hWnd(2) = FindWindow(vbNullString, "VB_KEEP_RUNNER")
WINDOW_hWnd(3) = FindWindow(vbNullString, "ClipBoard Logger")
WINDOW_hWnd(4) = FindWindow(vbNullString, "URL Logger")
WINDOW_hWnd(5) = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F009}", vbNullString) ' GOODSYNC
WINDOW_hWnd(6) = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F00A}", vbNullString) ' GOODSYNC2GO
WINDOW_hWnd(7) = FindWindow("rctrl_renwnd32", vbNullString) ' C:\Program Files (x86)\Microsoft Office\Office12\OUTLOOK.EXE

For RR_EYE = 1 To 10
    If WINDOW_hWnd(RR_EYE) > 0 Then
        I_R = GetWindowState(WINDOW_hWnd(RR_EYE))
        If I_R = -1 Then
            hWndResult = GetWindowRect(WINDOW_hWnd(RR_EYE), RectFINDER)
            
            SET_GO = True
            RT = RectFINDER.Top
            RB = RectFINDER.Bottom
            RR = RectFINDER.Right
            RL = RectFINDER.Left
            
            If RT < 0 And RB < 0 And RR < 0 And RL < 0 Then
                Call cmdMoveMax_SIMPLE(WINDOW_hWnd(RR_EYE))
            
            End If
        End If
    End If
Next

End Sub

Private Sub Label_MAXIMIZE_ELITE_SPY_Click()
    
    Dim WINDOW_hWnd
    ' window_hWnd = FindWindow("ThunderRT6FormDC", "EliteSpy+ 2001 __ www.PlanetSourceCode.com __ Version 1.0.510)")
    WINDOW_hWnd = FindWindowPart("EliteSpy+ 2001 __ www.PlanetSourceCode.com __ Version")
    Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
    Label_MAXIMIZE_ELITE_SPY.BackColor = RGB(255, 255, 255)
    
    Call MAX_FORM_CHECKER
    
    ShowWindow WINDOW_hWnd, SW_NORMAL
    
    Me.WindowState = vbMinimized

Exit Sub

Dim hWndResult
Dim RectFINDER As RECT
hWndResult = GetWindowRect(WINDOW_hWnd, RectFINDER)

Dim SET_GO, RT, RB, RR, RL

SET_GO = True
RT = RectFINDER.Top
RB = RectFINDER.Bottom
RR = RectFINDER.Right
RL = RectFINDER.Left

If RT < 0 And RB < 0 And RR < 0 And RL < 0 Then

End If

'' CENTER ON SCREEN COMPENSATE TASKBAR AT BOTTOM
''----------------------------------------------
'If Screen.height - RectTaskbar.Top * Screen.TwipsPerPixelY < 2500 Then
'    Me.Left = Screen.width / 2 - Me.width / 2
'    Me_Top = (Screen.height - ((RectTaskbar.Bottom - RectTaskbar.Top) * Screen.TwipsPerPixelY)) / 2 - Me.height / 2
'Else
'    Me.Left = Screen.width / 2 - Me.width / 2




' txtMhWnd.Text = WINDOW_hWnd



Call cmdMoveMax_Click

GoTo ENDER

Dim XNOW, I_R, XX_hWnd, VAR, FN_VAR
XNOW = Now + TimeSerial(0, 1, 5)
Do
    I_R = GetWindowState(WINDOW_hWnd)
    If I_R > 0 Then Exit Do
    Sleep 10
Loop Until XNOW < Now

XX_hWnd = WINDOW_hWnd



If I_R = 0 Then

    If XX_hWnd > 0 Then
        
        'WinGet, PID_01, PID, %FN_VAR_2% ahk_class AutoHotkey
        pid = -1
        VAR = cProcesses.Convert(XX_hWnd, pid, cnFromhWnd Or cnToProcessID)
        If pid > 0 Then
            Result = cProcesses.Process_Kill(pid)
            Exit Sub
        End If
    End If
    
    
    FN_VAR = "D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe"
    ' Const ShowWindow_2 = 1, DontShowWindow = 0, DontWaitUntilFinished = False, WaitUntilFinished = True
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
        WSHShell.Run """" + FN_VAR_1 + """", DontShowWindow, DontWaitUntilFinished
    Set WSHShell = Nothing



End If

ENDER:

Beep
Me.WindowState = vbMinimized
End Sub

Private Sub Label_MAXIMIZE_GOODSYNC2GO_Click()

Dim GOODSYNC_WINDOW_hWnd
GOODSYNC_WINDOW_hWnd = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F00A}", vbNullString)

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_MAXIMIZE_GOODSYNC2GO.BackColor = RGB(255, 255, 255)

ShowWindow GOODSYNC_WINDOW_hWnd, SW_MAXIMIZE
Beep
Me.WindowState = vbMinimized

End Sub

Private Sub Label_RUN_AUTOHOTKEY_SET_NETWORK_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT

Label_RUN_AUTOHOTKEY_SET_NETWORK.BackColor = RGB(255, 255, 255)
Label_RUN_AUTOHOTKEY_SET_NETWORK.BackColor = &H808080

' Call RELOAD_OR_KILL_PATH_ARRAY_SET_NETWORK_ALL_CODE_03_OF_04

MsgBox "TRY NOT TO RUN AT MOMENT WORK IN PROGRESS _ NOT RUN" + vbCrLf + vbCrLf + "Label_RUN_AUTOHOTKEY_SET_NETWORK"


End Sub

' 03 OF 04 ---- RESTORE CREATE FILE DELETE AND RELOAD
Sub RELOAD_OR_KILL_PATH_ARRAY_SET_NETWORK_ALL_CODE_03_OF_04()

    OPERATION_CREATE_PATH_SET_NETWORK = "AHK"
    Call RESTORE_THE_CREATE_PATH_ARRAY_SET_NETWORK_ALL_CODE

    Call TIMER_SUB_AUTOHOTKEY_RELOAD

    Call RETURN_FILENAME_FORMAT_LOCAL_LEVEL_VB_AND_AHK
    If Dir(FileName_AHK) <> "" Then
        Kill FileName_AHK
    End If
    
End Sub

Sub RETURN_FILENAME_FORMAT_LOCAL_LEVEL_VB_AND_AHK()

    FileName_VB = "_01_c_drive\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NET_VB_CODE_EXE_" + GetComputerName + ".TXT"
    FileName_AHK = "_01_c_drive\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NET_AUTOHOTKEY_CODE_EXE_" + GetComputerName + ".TXT"

    Dim ELEMENT1

    NET_PATH = GetComputerName
    ELEMENT1 = "\\"
    ELEMENT2 = NET_PATH
    NET_PATH = Replace(NET_PATH, "-", "_")
    ELEMENT3 = "\"
    ELEMENT4 = NET_PATH
    ELEMENT5 = FileName_VB
    FileName_VB = ELEMENT1 + ELEMENT2 + ELEMENT3 + ELEMENT4 + ELEMENT5
    ELEMENT5 = FileName_AHK
    FileName_AHK = ELEMENT1 + ELEMENT2 + ELEMENT3 + ELEMENT4 + ELEMENT5
    
End Sub

Sub TIMER_POLL_PATH_ARRAY_SET_NETWORK_ALL_SPEICAL_REQUEST_Timer()
    ' RUNNER FROM THIS TIMER
    ' Timer_1_SECOND_Timer
    
    Dim R_COUNTER, FILE_NAME, FileName
    
    FILE_NAME = "# DATA\KILL_RELOAD_ALL_NET_SPECIAL_REQUEST_" + GetComputerName + "_.TXT"
    FileName_4 = "D:\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\" + FILE_NAME
    
    On Error Resume Next
    SET_COMPUTER_TO_RUN_PID_EXE = ""
    FR1 = FreeFile
    Open FileName_4 For Input As #FR1
        Line Input #FR1, SET_COMPUTER_TO_RUN_PID_EXE
    Close #FR1
    
    If SET_COMPUTER_TO_RUN_PID_EXE = "Chrome.exe" Then
        Kill FileName_4
        Call Label_KILL_CHROME_Click
    End If

End Sub


Sub CREATE_PATH_ARRAY_SET_NETWORK_ALL_SPEICAL_REQUEST()
    ' RUNNER FROM THIS TIMER
    ' Timer_1_SECOND_Timer
    
    Dim R_COUNTER, FILE_NAME, FileName
    Dim FileName_O_SP_RQ
    
    '\\7-ASUS-GL522VW\7_ASUS_GL522VW_02_d_drive\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER
    SET_COMPUTER_TO_RUN_02 = Replace(SET_COMPUTER_TO_RUN, "-", "_")
    FILE_NAME = "# DATA\KILL_RELOAD_ALL_NET_SPECIAL_REQUEST_" + SET_COMPUTER_TO_RUN + "_.TXT"
    FileName_O_SP_RQ = "\\" + SET_COMPUTER_TO_RUN + "\" + SET_COMPUTER_TO_RUN_02 + "_02_d_drive\VB6\VB-NT\00_Best_VB_01\VB_KEEP_RUNNER\" + FILE_NAME
    
    R_COUNTER = SET_COMPUTER_TO_RUN

    FileName = FileName_O_SP_RQ
    FR1 = FreeFile
    On Error Resume Next
    
    ' Path/File access error
    Open FileName For Output As #FR1
        If Err.Description <> "" Then
            MsgBox Err.Description + vbCrLf + vbCrLf + FileName, vbMsgBoxSetForeground
            Exit Sub
        End If
        
        Print #FR1, SET_COMPUTER_TO_RUN_PID_EXE
    Close #FR1

End Sub



Sub CREATE_PATH_ARRAY_SET_NETWORK_ALL_CODE()
    ' RUNNER FROM THIS TIMER
    ' Timer_1_SECOND_Timer
    
    FileName_O_VB = "_01_c_drive\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NET_VB_CODE_EXE"
    FileName_O_AHK = "_01_c_drive\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NET_AUTOHOTKEY_CODE_EXE"

    OPERATION_CREATE_PATH_SET_NETWORK = "AHK"
    
    If OPERATION_CREATE_PATH_SET_NETWORK = "VB" Then
        FileName_2 = FileName_O_VB
        FileName_4 = "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NET_VB_CODE_EXE_" + GetComputerName + ".TXT"
    End If
    If OPERATION_CREATE_PATH_SET_NETWORK = "AHK" Then
        FileName_2 = FileName_O_AHK
        FileName_4 = "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 19-SCRIPT_TIMER_UTIL__KILL_RELOAD_ALL_NET_AUTOHOTKEY_CODE_EXE_" + GetComputerName + ".TXT"
    End If
    
    Call READ_ALL_NETWORK_COMPUTER_NAME_PATH_INTO_ARRAY

End Sub


Sub RESTORE_THE_CREATE_PATH_ARRAY_SET_NETWORK_ALL_CODE()

    Call CREATE_PATH_ARRAY_SET_NETWORK_ALL_CODE
    Call CREATE_FILENAME_FORMAT_ALL_NETWORK_LEVEL_FROM_ARRAY

End Sub


Sub CREATE_FILENAME_FORMAT_ALL_NETWORK_LEVEL_FROM_ARRAY()
    
    Dim R_COUNTER, FileName
    
    For R_COUNTER = 1 To UBound(Array_FileName)
        FileName = Array_FileName(R_COUNTER) + ".TXT"
        FR1 = FreeFile
        Open FileName For Output As #FR1
        Print #FR1, "This is a test string From VB"
        Close #FR1
    Next
    
End Sub

Sub TIMER_SUB_AUTOHOTKEY_RELOAD()

    ' DetectHiddenWindows, ON
    ' SetTitleMatchMode 3  ; EXACTLY
    
    Dim XX_hWnd, AHK_TERMINATOR_VERSION, VAR
    Dim TEMP_VAR_1
    Dim TEMP_VAR_2
    Dim TEMP_VAR_3
    
    
    FN_VAR_1 = "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk"
    AHK_TERMINATOR_VERSION = " - AutoHotkey v1.1.30.00" ' --AutoHotkey v"A_AhkVersion"
    TEMP_VAR_1 = FN_VAR_1
    TEMP_VAR_2 = """" + AHK_TERMINATOR_VERSION + """"
    TEMP_VAR_3 = TEMP_VAR_1 + TEMP_VAR_2
    TEMP_VAR_3 = Replace(TEMP_VAR_3, """", "")
    FN_VAR_2 = TEMP_VAR_3
    XX_hWnd = FindWindow(FN_VAR_2, "AutoHotkey")
    
    If XX_hWnd > 0 Then
        
        'WinGet, PID_01, PID, %FN_VAR_2% ahk_class AutoHotkey
        pid = -1
        VAR = cProcesses.Convert(XX_hWnd, pid, cnFromhWnd Or cnToProcessID)
        If pid > 0 Then
            Result = cProcesses.Process_Kill(pid)
            Exit Sub
        End If
    End If
    
    If XX_hWnd = 0 Then
        ' Const ShowWindow_2 = 1, DontShowWindow = 0, DontWaitUntilFinished = False, WaitUntilFinished = True
        Dim WSHShell
        Set WSHShell = CreateObject("WScript.Shell")
            WSHShell.Run """" + FN_VAR_1 + """", DontShowWindow, DontWaitUntilFinished
        Set WSHShell = Nothing
    End If
    ' "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk"
End Sub


Sub READ_ALL_NETWORK_COMPUTER_NAME_PATH_INTO_ARRAY()
    
    ReDim Preserve Array_FileName(50)
    ReDim Preserve Array_NETPATH_01(50)
    ReDim Preserve Array_NETPATH_02(50)
    Dim SET_GO
    
    ArrayCount = 0
    FR1 = FreeFile
    Open "C:\NETWORK_COMPUTER_NAME.txt" For Input As #FR1
    Do
        Line Input #FR1, NET_PATH
        
        SET_GO = True
        If InStr(NET_PATH, "BTHUB") > 0 Then SET_GO = False
        If InStr(NET_PATH, "NAS-QNAP-ML") > 0 Then SET_GO = False
        If SET_GO = True Then
            ArrayCount = ArrayCount + 1
            Array_NETPATH_01(ArrayCount) = NET_PATH
            Array_NETPATH_02(ArrayCount) = Replace(NET_PATH, "-", "_")
            ELEMENT1 = "\\"
            ELEMENT2 = Array_NETPATH_01(ArrayCount)
            ELEMENT3 = "\"
            ELEMENT4 = Array_NETPATH_02(ArrayCount)
            ELEMENT5 = FileName_2
            ELEMENT7 = "_" + NET_PATH

            Array_FileName(ArrayCount) = ELEMENT1 + ELEMENT2 + ELEMENT3 + ELEMENT4 + ELEMENT5 + ELEMENT7
        End If
    Loop Until EOF(FR1)
    Close FR1
    
    ReDim Preserve Array_FileName(ArrayCount)
    ReDim Preserve Array_NETPATH_01(ArrayCount)
    ReDim Preserve Array_NETPATH_02(ArrayCount)

End Sub



Private Sub Label_TASK_KILLER_CMD_Click()
Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_TASK_KILLER_CMD.BackColor = RGB(255, 255, 255)
Call MNU_TASK_KILLER_CMD_Click

Me.WindowState = vbMinimized

End Sub


Private Sub Label_CLOSE_hWnd_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_CLOSE_HWND.BackColor = RGB(255, 255, 255)

If Val(TxtPID) > 0 Then
    Thread_Resume TxtPID
End If

Dim hWnd_RESULT
hWnd_RESULT = txthWnd
If hWnd_RESULT > 0 Then
    Result = PostMessage(hWnd_RESULT, WM_CLOSE, 0&, 0&)
End If
hWnd_RESULT = txtParent
If hWnd_RESULT > 0 Then
    Result = PostMessage(hWnd_RESULT, WM_CLOSE, 0&, 0&)
End If
hWnd_RESULT = txtParentX
If hWnd_RESULT > 0 Then
    Result = PostMessage(hWnd_RESULT, WM_CLOSE, 0&, 0&)
End If

Me.WindowState = vbMinimized
' Beep

End Sub


Private Sub LABEL_KILL_NOT_RESPOND_Click()
Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
LABEL_KILL_NOT_RESPOND.BackColor = RGB(255, 255, 255)
Call MNU_KILL_NOT_RESPOND_TOP_Click

Me.WindowState = vbMinimized

End Sub


Private Sub LABEL_KILL_NOT_RESPOND_FORCE_Click()
Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
LABEL_KILL_NOT_RESPOND_FORCE.BackColor = RGB(255, 255, 255)

Call MNU_TASK_KILLER_NOT_RESPONDER_FORCE_Click

Me.WindowState = vbMinimized

End Sub


Function FIND_COMPUTER_TO_RUN(FIND_COMPUTER_TO_RUN_VAR As String)

Dim HH_1, HH_2, HH_3, HH_4, R_COUNTER
HH_1 = Mid(FIND_COMPUTER_TO_RUN_VAR, 1, 1)
HH_2 = Mid(FIND_COMPUTER_TO_RUN_VAR, 2, 1)
For R_COUNTER = 1 To UBound(Array_FileName)
    HH_3 = Mid(Array_NETPATH_01(R_COUNTER), 1, 1)
    HH_4 = Mid(Array_NETPATH_01(R_COUNTER), InStr(3, Array_NETPATH_01(R_COUNTER), "-") + 1, 1)
    If InStr(Array_NETPATH_01(R_COUNTER), "LINDA") Then
        HH_4 = Mid(Array_NETPATH_01(R_COUNTER), 3, 1)
    End If
    If InStr(Array_NETPATH_01(R_COUNTER), "MSI") Then
        HH_4 = Mid(Array_NETPATH_01(R_COUNTER), 3, 1)
    End If
    If HH_1 = HH_3 And HH_2 = HH_4 Then Exit For
Next
FIND_COMPUTER_TO_RUN = Array_NETPATH_01(R_COUNTER)

End Function

Private Sub Label_KILL_CHROME_Click()
 
Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_KILL_CHROME.BackColor = RGB(255, 255, 255)

SET_COMPUTER_TO_RUN_PID_EXE = "Chrome.exe"
If SET_COMPUTER_TO_RUN <> "" Then
    Call CREATE_PATH_ARRAY_SET_NETWORK_ALL_SPEICAL_REQUEST
    SET_COMPUTER_TO_RUN = ""
    Exit Sub
End If

Dim R, A1, A2

For R = 1 To lstProcess_3_SORTER_ListView.ListItems.Count
    A1 = lstProcess_3_SORTER_ListView.ListItems.Item(R).SubItems(1)
    If InStr(A1, SET_COMPUTER_TO_RUN_PID_EXE) > 0 Then
        pid = Val(lstProcess_3_SORTER_ListView.ListItems.Item(R))
        cProcesses.Process_Kill (pid)
    End If
Next

Me.WindowState = vbMinimized

End Sub






Private Sub Label_VB_MODIFED_TIME_Click()

End Sub

Private Sub Label_VB_MODIFIED_ERROR_Click()

End Sub

Private Sub Label_VB_MODIFIED_ACTIVE_Change()
'Stop

'Debug.Print Label_VB_MODIFIED_ACTIVE.Caption

End Sub

Private Sub Label_VB_MODIFIED_ACTIVE_Click()

NET_COPY_CHANGE_HAPPENER = True

End Sub

Private Sub Label_VB_MODIFIED_TIME_Click()
'Label_VB_MODIFIED_TIME.CAP
If ListView_VB_MODIFIED_ERROR.Visible = False Then
    ListView_VB_MODIFIED_ERROR.Visible = True
Else
    ListView_VB_MODIFIED_ERROR.Visible = False
End If
Call Form_Resize
Call Form_Resize

End Sub

Private Sub Label22_Click()
'Label22.CAPTION
'Clipboard.Clear
'Clipboard.SetText Label22

Label57.Caption = "COMMAND LINE STATUS__ " + Label22

'Label56.BackColor = Label59.BackColor
'Label55.BackColor = Label58.BackColor
End Sub

Private Sub Label24_Click()

Call Label_MAXIMIZE_GOODSYNC_Click
Call Label_MAXIMIZE_GOODSYNC2GO_Click
 
Call Label_CLOSE_GOODSYNC_Click
Call Label_CLOSE_GOODSYNC2GO_Click
 
End Sub

Private Sub Label30_Click()

Label57.Caption = "COMMAND LINE STATUS__ " + Label30

'Label55.BackColor = Label59.BackColor
'Label56.BackColor = Label58.BackColor

End Sub

Private Sub Label42_Click()
'Label42.caption
End Sub

Private Sub Label44_Click()
VB_EXE_SYNC_RUN_INSTANTLY = True
Call VB_EXE_SYNC
End Sub

Private Sub Label46_Click()
'Label46.LEFT
'Label46.FONT
'Label46.CAP
End Sub

Private Sub Label48_Click()

TIMER2_TIMER_BEGAN = Now + TimeSerial(0, 0, 20)

End Sub

Private Sub Label50_Click()
'Label50.width
'Label50.CAP
Call Timer_EnumProcess_Timer

End Sub

Private Sub Label_CLOSE_GOODSYNC_Click()

Call Label_MAXIMIZE_GOODSYNC_Click

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_CLOSE_GOODSYNC.BackColor = RGB(255, 255, 255)

Call MNU_CLOSE_GOODSYNC_Click

Me.WindowState = vbMinimized

End Sub

Private Sub Label72_Click()

End Sub

Private Sub Label51_Click()
'Label51.TOP
'Label51.FontSize
End Sub

Private Sub Label52_Click()
' Label52.LEFT
End Sub

Private Sub Label53_Click()
'Label53.width
'Label53.height
'Label53.left
'Label53.caption
'Label53.font

'Label53_Here
Call EnumProcess

Timer_Pause_Update.Interval = 60000
Timer_Pause_Update.Enabled = True

Label53.BackColor = Label59.BackColor

End Sub

Private Sub Label54_Click()
'Label54.top
'Label54.FONT -- NOT CODED

End Sub

Private Sub Label55_Click()

Call Label_MAXIMIZE_GOODSYNC_Click
Call Label_MAXIMIZE_GOODSYNC2GO_Click



End Sub

Private Sub Label60_Click()
'Label60.CAP
End Sub

Private Sub Label63_Click()

End Sub

Private Sub Label64_Click()

TIMER2_TIMER_BEGAN = Now + TimeSerial(0, 0, 40)

End Sub

Private Sub Label66_Click()
cProcesses.Process_Kill (TxtPID)
Beep

End Sub

Private Sub ListView_CPU_INFO_BeforeLabelEdit(Cancel As Integer)
'ListView_CPU_INFO.
End Sub

Private Sub lstProcess_2_ListView_DblClick()


PROCESS_TO_KILLER = lstProcess_2_ListView.ListItems(lstProcess_2_ListView.SelectedItem.Index).SubItems(1)
PROCESS_TO_KILLER_PID = lstProcess_2_ListView.ListItems(lstProcess_2_ListView.SelectedItem.Index)

Label22.Caption = "TASKKILLER BY PID NUMBER __ # " + PROCESS_TO_KILLER_PID
Label30.Caption = "TASKKILLER NAME ___________ " + PROCESS_TO_KILLER

'PROCESS_TO_KILLER PID
Label29_Click

'PROCESS_TO_KILLER PID
'Label22_Click

'PROCESS_TO_KILLER /F * /T
Label57.Caption = "COMMAND LINE STATUS__ " + "TASKKILL /F /IM """ + Replace(UCase(PROCESS_TO_KILLER), ".EXE", "") + "*"" /T"

Label23_Click

End Sub

Private Sub lstProcess_2_ListView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
LISTVIEW_2_OR_3_HITT = 2
End Sub

Private Sub lstProcess_2_ListView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Timer_Pause_Update.Interval = 4000
Timer_Pause_Update.Enabled = True
Label53.BackColor = Label59.BackColor

End Sub

Private Sub lstProcess_3_SORTER_ListView_DblClick()
PROCESS_TO_KILLER = lstProcess_3_SORTER_ListView.ListItems(lstProcess_3_SORTER_ListView.SelectedItem.Index).SubItems(1)
PROCESS_TO_KILLER_PID = lstProcess_3_SORTER_ListView.ListItems(lstProcess_3_SORTER_ListView.SelectedItem.Index)

Label22.Caption = "TASKKILLER BY PID NUMBER __ # " + PROCESS_TO_KILLER_PID
Label30.Caption = "TASKKILLER NAME ___________ " + PROCESS_TO_KILLER

'PROCESS_TO_KILLER PID
Label29_Click

'PROCESS_TO_KILLER PID
'Label22_Click

'PROCESS_TO_KILLER /F * /T
Label57.Caption = "COMMAND LINE STATUS__ " + "TASKKILL /F /IM """ + Replace(UCase(PROCESS_TO_KILLER), ".EXE", "") + "*"" /T"

Label23_Click

End Sub

Private Sub lstProcess_3_SORTER_ListView_KeyDown(KeyCode As Integer, Shift As Integer)
LISTVIEW_2_OR_3_HITT = 3
End Sub

Private Sub lstProcess_3_SORTER_ListView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
LISTVIEW_2_OR_3_HITT = 3
End Sub

Private Sub lstProcess_3_SORTER_ListView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Timer_Pause_Update.Interval = 4000
Timer_Pause_Update.Enabled = True
Label53.BackColor = Label59.BackColor

End Sub

Private Sub MNU_CLIPBOARDER_REPLACE_ER_AND_Click()
' CLIPBOARD REPLACE "AND"


Dim VAR_ST_1
Dim VAR_ST_2
Dim VAR_ST_4
Dim CHANGE_VAR_1
Dim CHANGE_VAR_2
Dim CH_1
Dim CH_2
Dim CH_3
Dim CMOS_AND_TTL
Dim R1
Dim R2
Dim x

'Debug.Print "----"
VAR_ST_1 = Clipboard.GetText
' VAR_ST_1 = "and ff and hh and yy and 123456789 kk"
CHANGE_VAR_1 = LCase(" AND ")
x = Len(CHANGE_VAR_1) + 1
For R1 = 0 To Len(CHANGE_VAR_1) + 1
    For R2 = 1 To Len(CHANGE_VAR_1)
        
        If R1 > 0 And R1 < x Then
            CHANGE_VAR_2 = CHANGE_VAR_1
            CH_2 = UCase(Mid(CHANGE_VAR_2, R2, 1))
            CH_3 = UCase(Mid(CHANGE_VAR_2, R1, 1))
            Mid(CHANGE_VAR_2, R2, 1) = CH_2
            Mid(CHANGE_VAR_2, R1, 1) = CH_3
        End If
        If R1 = 0 Then CHANGE_VAR_2 = CHANGE_VAR_1
        If R1 = x Then CHANGE_VAR_2 = UCase(CHANGE_VAR_1)
        If InStr(CMOS_AND_TTL, CHANGE_VAR_2) = 0 Then
            ' Debug.Print CHANGE_VAR_2
            VAR_ST_1 = Replace(VAR_ST_1, CHANGE_VAR_2, " & ")
            VAR_ST_1 = Replace(VAR_ST_1, vbCr + Mid(CHANGE_VAR_2, 2), vbCr + "& ")
            VAR_ST_1 = Replace(VAR_ST_1, vbLf + Mid(CHANGE_VAR_2, 2), vbLf + "& ")
            VAR_ST_1 = Replace(VAR_ST_1, Mid(CHANGE_VAR_2, 1, Len(CHANGE_VAR_2) - 1) + vbCr, " &" + vbCr)
            VAR_ST_1 = Replace(VAR_ST_1, Mid(CHANGE_VAR_2, 1, Len(CHANGE_VAR_2) - 1) + vbLf, " &" + vbLf)
        End If
        
        CMOS_AND_TTL = CMOS_AND_TTL + "--" + CHANGE_VAR_2 + "__"
    Next
Next

If VAR_ST_1 <> VAR_ST_4 Then

    Do
        On Error Resume Next
        Clipboard.Clear
        VAR_ST_2 = Clipboard.GetText
        On Error GoTo 0
        If VAR_ST_2 <> "" Then Sleep 500
    Loop Until VAR_ST_2 = ""
    Sleep 500
    Do
        On Error Resume Next
        Clipboard.SetText VAR_ST_1, vbCFText
        Sleep 500
        VAR_ST_2 = Clipboard.GetText
        On Error GoTo 0
        If VAR_ST_2 <> VAR_ST_1 Then Sleep 800
    Loop Until VAR_ST_2 = VAR_ST_1

End If

Me.WindowState = vbMinimized
Beep

End Sub

Private Sub MNU_GOOGLE_SYNC_Click()
'C:\Program Files\Google\Drive\googledrivesync.exe

'
'Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
'Label_CHROME_PAGE_AUTO_ON.BackColor = RGB(255, 255, 255)

' Const ShowWindow_2 = 1, DontShowWindow = 0, DontWaitUntilFinished = False, WaitUntilFinished = True
Dim objShell
Set objShell = CreateObject("Wscript.Shell")
objShell.Run """C:\Program Files\Google\Drive\googledrivesync.exe""", DontShowWindow, DontWaitUntilFinished
Set objShell = Nothing

Me.WindowState = vbMinimized
End Sub

Private Sub MNU_STOP_GOODSYNC_SCRIPTOR_Click()
' Const ShowWindow_2 = 1, DontShowWindow = 0, DontWaitUntilFinished = False, WaitUntilFinished = True
Dim objShell
Set objShell = CreateObject("Wscript.Shell")
objShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 74-GOODSYNC STOP JOB WORKER.ahk""", ShowWindow_2, DontWaitUntilFinished
Set objShell = Nothing
Me.WindowState = vbMinimized
End Sub

Private Sub Timer_ALWAYS_ON_TOP_TO_START_WITH_ER_Timer()
    
'    If IsIDE = True And IsIDE_TEST = True Then Timer_ALWAYS_ON_TOP_TO_START_WITH_ER.Interval = 10000
'
    'Label60_HERE
    'Label60.Caption = "Me on Top at Loader" + Str(Counter_ALWAYS_ON_TOP_TIMER) + "_20 Second YES"
    
    Label60.Caption = "Me on Top " + Str(Counter_ALWAYS_ON_TOP_TIMER) + "_20 Sec"
        
    ' ----------------------------------------------
    ' DON'T REINFOCE THE ONTOP GO-ER WHEN IN THE IDE
    ' ----------------------------------------------
    If IsIDE = True And Me.hWnd = GetForegroundWindow Then
        Counter_ALWAYS_ON_TOP_TIMER = Counter_ALWAYS_ON_TOP_TIMER - 1
        SetWindowPos Me.hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
    
    If Counter_ALWAYS_ON_TOP_TIMER > -1 Then Exit Sub
    
    ' Put window on top of all others
    'SetWindowPos txtMhWnd.Text, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    'SetWindowPos Me.hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    ' Put window on top of all others
    'SetWindowPos txtMhWnd.Text, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    'SetWindowPos Me.hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    ' Remove window from top
    SetWindowPos Me.hWnd, hWnd_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'    Label60.Caption = "Me on Top @ Loader EXE Timer 20 Second NOT DONE"
    Label60.Caption = "Me on Top 20 Sec"
    Label60.BackColor = Label59.BackColor '49 58_59
    Timer_ALWAYS_ON_TOP_TO_START_WITH_ER.Enabled = False
    'Label60.Caption = "Me on Top @ Loader EXE Timer 20 Second Already ON"

End Sub

Private Sub Timer_CONNECTED_TO_THE_INTERNET_Timer()

Call IsInternetConnected

End Sub

Private Sub MNU_CLOSE_GOODSYNC_Click()

Dim hWnd_RESULT
hWnd_RESULT = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F009}", vbNullString)
If hWnd_RESULT > 0 Then
    Result = PostMessage(hWnd_RESULT, WM_CLOSE, 0&, 0&)
End If

Me.WindowState = vbMinimized
Beep

End Sub

Private Sub MNU_CLOSE_GOODSYNC_2_GO_Click()

Dim hWnd_RESULT
hWnd_RESULT = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F00A}", vbNullString)
If hWnd_RESULT > 0 Then
    Result = PostMessage(hWnd_RESULT, WM_CLOSE, 0&, 0&)
End If

Me.WindowState = vbMinimized
Beep

End Sub


Private Sub Label_GOODSYNC_COLLECTION_SCRIPT_RUN_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_GOODSYNC_COLLECTION_SCRIPT_RUN.BackColor = RGB(255, 255, 255)

Call MNU_GOODSYNC_COLLECTION_SCRIPT_RUN_Click

Me.WindowState = vbMinimized

End Sub

Private Sub Label_KILL_HUBIC_Click()
'----
'How to select an external program's tray menu item? - Ask for Help - AutoHotkey Community
'https://autohotkey.com/board/topic/80327-how-to-select-an-external-programs-tray-menu-item/
'----
Dim hWnd_RESULT
hWnd_RESULT = FindWindow("hWndWrapper[hubiC.exe;;d22553e4-d0db-4e73-aa9f-48b38951eef0]", vbNullString)
If hWnd_RESULT > 0 Then
    Result = PostMessage(hWnd_RESULT, WM_CLOSE, 0&, 0&)
End If
' ---------------------------------------------------------------------
' CAN'T FIND UNLESS OPEN ON SCREEN
' AND CLOSE ONLY CLOSE THAT FORM NOT WHOLE THING
' UNLESS BY SLECT MENU OPTION
' IF USE AUTOHOTKEY - BUT STILL PROBLEM HAS OPEN MENU AND MOUSE CLICKER
' NOT GOOD IF TRAY ITEMS ARE BLOCK FROM APPEAR LIKE WHOLE TRAY FROZEN
' ---------------------------------------------------------------------

Dim VAR, EXE_STRING As String
pid = -1
VAR = cProcesses.Convert(hWnd_RESULT, pid, cnFromhWnd Or cnToProcessID)
VAR = cProcesses.Convert(pid, EXE_STRING, cnFromProcessID Or cnToEXE)


' -------------------------------------------------------------------
' BOTH ABLE TO USE .Convert HAS USE .GetEXEID ANYWAY AS PART FUNCTION
' DEBUG IT REMEMBER TO GET STRING INPUT AOUND RIGHT WAY WHEN USE
' .Convert(EXE_STRING, PID, cnFromEXE Or cnToProcessID)
' -------------------------------------------------------------------
'PID = -1
'VAR = cProcesses.GetEXEID(PID, EXE_STRING)
EXE_STRING = "HubiC.exe"
pid = -1
VAR = cProcesses.Convert(EXE_STRING, pid, cnFromEXE Or cnToProcessID)

PROCESS_TO_KILLER = EXE_STRING
PROCESS_TO_KILLER_PID = pid
Call LISTVIEW_CLICKER

'PROCESS_TO_KILLER PID
Label29_Click

'PROCESS_TO_KILLER PID
'Label22_Click

'PROCESS_TO_KILLER /F * /T
Label57.Caption = "COMMAND LINE STATUS__ " + "TASKKILL /F /IM """ + Replace(UCase(PROCESS_TO_KILLER), ".EXE", "") + "*"" /T"

Label23_Click



Me.WindowState = vbMinimized
Beep


End Sub

Private Sub MNU_CMD_KILL_MAX_Click()
Beep

If InStr(MNU_CMD_KILL_MAX.Caption, "KILL CMD PID LIMITER > 20=TRUE") > 0 Then
    MNU_CMD_KILL_MAX.Caption = "KILL CMD PID LIMITER > 20=FALSE"
Else
    MNU_CMD_KILL_MAX.Caption = "KILL CMD PID LIMITER > 20=TRUE"
End If

End Sub

Private Sub MNU_KILL_MAX_AHK_Click()
Beep

If InStr(MNU_KILL_MAX_AHK.Caption, "KILL AHK PID LIMITER > 100=TRUE") > 0 Then
    MNU_KILL_MAX_AHK.Caption = "KILL AHK PID LIMITER > 100=FALSE"
Else
    MNU_KILL_MAX_AHK.Caption = "KILL AHK PID LIMITER > 100=TRUE"
End If

End Sub

Private Sub MNU_GIVE_ME_TIME_WITHER_UTC_Click()

Dim i

i = MNU_GIVE_ME_TIME_WITHER_UTC.Caption

' DEFAULT AT STARTOR
'----------------------------------------------------
If TO_SETTER = False Then
    TO_SETTER = True
    i = "GIVE ME TIME AND UNI_ UTC Time Toggle = YES"
    i = "GIVE ME TIME AND UNI_ UTC Time Toggle = NOT"
    MNU_GIVE_ME_TIME_WITHER_UTC.Caption = i
    Call SET_MENU_PADD_WORK
    Exit Sub
End If

If InStr(i, "NOT") Then
    MNU_GIVE_ME_TIME_WITHER_UTC.Caption = Replace(i, "NOT", "YES")
    Exit Sub
End If
If InStr(i, "YES") Then
    MNU_GIVE_ME_TIME_WITHER_UTC.Caption = Replace(i, "YES", "NOT")
End If


End Sub


Private Sub MNU_RESET_TIMERING_FOR_AUTO_RELOADER_Click()
    Dim FOR_NEXT_R
    For FOR_NEXT_R = 0 To UBound(BLOCK_RUN_1)
        If BLOCK_RUN_1(FOR_NEXT_R) > Now Then BLOCK_RUN_1(FOR_NEXT_R) = 0
    Next
    Beep
End Sub


Private Sub MNU_SCREEN_SHOT_ME_Click()

    'FIND ADJUST
    'PrintCurrentFormOntoForm
    
    Dim TT, TS, FOLDERNAME_AUTO

    Me.SetFocus
    
    FOLDERNAME_AUTO = App.Path + "\# DATA\" + GetComputerName + "\"
    If Dir(FOLDERNAME_AUTO, vbDirectory) = "" Then
        On Error Resume Next
            MkDir FOLDERNAME_AUTO = App.Path + "\# DATA\"
            MkDir FOLDERNAME_AUTO
    End If

    TT = PrintCurrentFormOntoForm(SCREEN_CAP)
    TS = FOLDERNAME_AUTO + "FormCapture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + ".jpg"
    'MsgBox TS
    On Error Resume Next
    SavePicture SCREEN_CAP.Picture, TS
    Beep
    'ERR.NUMBER
    
    Clipboard.Clear
    Clipboard.SetData SCREEN_CAP.Picture, vbCFBitmap
    
    '= Clipboard.GetData(vbCFBitmap)

'----'[RESOLVED]
'Capture Picture Of A Form When A Portion Of It Is Off The Screen...-VBForums
'http://www.vbforums.com/showthread.php?734947-RESOLVED-Capture-Picture-Of-A-Form-When-A-Portion-Of-It-Is-Off-The-Screen
'----

     Call MNU_SCREENSHOT_IMAGE_FOLDER_Click

End Sub

Private Sub MNU_SCREENSHOT_IMAGE_FOLDER_Click()
    
    Dim TT, TS, FOLDERNAME_AUTO

    FOLDERNAME_AUTO = App.Path + "\# DATA\" + GetComputerName + "\"
    If Dir(FOLDERNAME_AUTO, vbDirectory) = "" Then
        On Error Resume Next
            MkDir FOLDERNAME_AUTO = App.Path + "\# DATA\"
            MkDir FOLDERNAME_AUTO
    End If
    
    Shell "EXPLORER /SELECT, """ + FIND_RECENT_FILE(FOLDERNAME_AUTO) + """", vbMaximizedFocus
    
    Me.WindowState = vbMinimized

End Sub

Function FIND_RECENT_FILE(INPUT_PATH) As String

Dim FSO, File, recentDate, objStartFolder, rootFolder, fol, fil, recentFile
' create a global copy of the filesystem object
Set FSO = CreateObject("Scripting.FileSystemObject")
Set recentFile = Nothing

objStartFolder = INPUT_PATH '"D:\0 00 Art Loggers\# APP AND SCREEN\7-ASUS-GL522VW\HOT_KEY_SCREEN_SHOT"

'Set fso = CreateObject("Scripting.FileSystemObject")
'Set rootFolder = fso.GetFolder(objStartFolder)

' get current folder
Set fol = FSO.GetFolder(objStartFolder)

' go thru each files in the folder
For Each File In fol.Files
    If (recentFile Is Nothing) Then
        Set recentFile = File
    ElseIf (File.DateLastModified > recentFile.DateLastModified) Then
        If InStr(UCase(File), "THUMBS.DB") = 0 Then
            Set recentFile = File
        End If
    End If
Next

FIND_RECENT_FILE = recentFile

End Function


Private Sub MNU_VERSION_Click()
' MNU_VERSION

Clipboard.Clear
Clipboard.SetText (MNU_VERSION.Caption)
Me.WindowState = vbMinimized

End Sub

Private Sub Text_SYSTEM_START_TIME_02_Change()
'Text_SYSTEM_START_TIME_02.text
End Sub

Private Sub Timer_SHOW_THE_TIME_Timer()

Dim A_DATE_TIME As String ', PATH_VOLUME_NAME, R_FOR
'Dim A_DATE_TIME_TEMP As String
Dim XI

'LCASE THE SECOND PM AM DIGIT
'BETTER

'A_DATE_TIME = "[__ " + Format(Now, "DDDD YYYY-MMM-DD-HH-MM-SS ") + A_DATE_TIME_PM + " __]"
'If Hour(Now) = 16 Then
'A_DATE_TIME = "[__ " + Format(Now, "DDDD YYYY-MMM-DD-HH-MM-SS AM/PM") + " __]"
'End If
'MsgBox Format(Now, "YYYY-MMM-DD -- HH:MM:SS - HH:MM:SS-AM/PM DDD")


A_DATE_TIME_PM = Mid(Format(Now, "HH-MM-SS Am/Pm"), 10)

A_DATE_TIME = "[__ " + Format(Now, "DDDD HH:MM:SS") + " " + A_DATE_TIME_PM + "__" + Format(Now, "DD MMMM YYYY") + " __]"
If Hour(Now) = 16 Then
    A_DATE_TIME = "[__ " + Format(Now, "DDDD HH:MM:SS AM/PM") + "__" + Format(Now, "DD MMMM YYYY") + " __]"
End If
A_DATE_TIME = Replace(A_DATE_TIME, "PM", "Pm")
A_DATE_TIME = Replace(A_DATE_TIME, "AM", "Am")
'------------------------------------------------------
'REPLACE LAST DIGIT AS 0
'----------------------------
XI = InStr(A_DATE_TIME, "m__")
Mid(A_DATE_TIME, XI - 3, 1) = "0"
'----------------------------



If MNU_SHOW_THE_TIME.Caption <> A_DATE_TIME Then
    MNU_SHOW_THE_TIME.Caption = A_DATE_TIME
End If

'If Val(Second(Now)) Mod 10 <> 0 Then Exit Sub

'MNU_SHOW_THE_TIME.Caption = A_DATE_TIME

'MNU_SHOW_THE_TIME.Enabled = False
'MNU_SHOW_THE_TIME.Enabled = True

End Sub


' Get current mouse cordinates
Private Sub Timer_MOUSE_CORD_Timer()
    'Timer_MOUSE_CORD.
    'Timer_MOUSE_CORD.Enabled = False
    
    If IsIDE = True And IsIDE_TEST = True Then Timer_MOUSE_CORD.Interval = 10000
    'IsIDE_TEST = True
    
    Dim tPA As POINTAPI
    Dim mWnd
    Dim i_string
    Dim LhWnd
    
    ' Get cursor cordinates
    GetCursorPos tPA
    ' Set label caption to cursor cordinates
    lblCordi.Caption = "X: " & tPA.x & "  Y: " & tPA.y
    
    LhWnd = WindowFromPoint(tPA.x, tPA.y)
    If Old_lhWnd = LhWnd Then
        If FindHandle_hWnd_COUNT_CHANGE = True Then
            Call EnumProcess
        End If
    End If
    
    Old_lhWnd = LhWnd
    
    If TIMER2_TIMER_BEGAN > Now Then
    
        Label48.Caption = Format(DateDiff("s", Now, TIMER2_TIMER_BEGAN), "00") + " Second"
        Label48.FontBold = True
        Label48.FontSize = 15
        
        Label64.Caption = Label48.Caption
        Label64.FontBold = Label48.FontBold
        Label64.FontSize = Label48.FontSize
        
        i_string = "USE " + Format(DateDiff("s", Now, TIMER2_TIMER_BEGAN, "00")) + " SECOND HOOVER"
        If i_string <> MNU_HOOVER_20_SECOND.Caption Then MNU_HOOVER_20_SECOND.Caption = i_string
            mWnd = WindowFromPoint(tPA.x, tPA.y)
            Call ChunkCodeOnMouse
        Else
            
            If TIMER2_TIMER_BEGAN <> 0 Then
                TIMER2_TIMER_BEGAN = 0
                Label48.Caption = "20 Sec"
                Label64.Caption = "40 Sec"
            End If
    End If
End Sub





Function IsInternetConnected()
    Dim ret As Long
    ret = InternetGetConnectedStateEx(ret, sConnType, 254, 0)
    If ret = 1 Then
    
        Dim CONNECT_TYPE
        CONNECT_TYPE = Replace(sConnType, " Connection", "")
        CONNECT_TYPE = Left$(CONNECT_TYPE, InStr(CONNECT_TYPE, Chr(0)) - 1)
        CONNECT_TYPE = CONNECT_TYPE + " HardCore"
        
        Label54.Caption = "Connected to the Internet by " & CONNECT_TYPE
        Label54.Caption = "Internet Connected" + vbCrLf + "By " & CONNECT_TYPE
        'MsgBox "You are connected to Internet via a " & sConnType, vbInformation
        O_Ret_Connected_To_The_Internet = False
    Else
         Label54.Caption = "Not connected to Internet"
         Label54.Caption = "Internet Not" + vbCrLf + "Connected"
'         Me.WindowState = vbMaximized
        If Me.Visible = True Then
         Me.WindowState = vbNormal
         'MsgBox "Not connected to internet", vbInformation + vbMsgBoxSetForeground
         If O_Ret_Connected_To_The_Internet = False Then
            O_Ret_Connected_To_The_Internet = True
            Call IsInternetConnected
         End If
        End If
    End If
End Function
     


Sub WIDTH_AND_HEIGHT(WX, HY)

On Error GoTo ENDER

WX = 0: HY = 0
Dim II22, II23
Dim XYZ
On Local Error Resume Next

Dim Control As Control
For Each Control In Me.Controls
    Err.Clear
    II22 = False
    II23 = Control.width
    'If Control.Name = "chkOnTop" Then Stop
    
    'Debug.Print Err.Description
    '----------------------------
    'LISTVIEW -- DONT USE ENABLED
    'OR HERE INSTEAD ONLY
    'CONTROL.Visible
    '----------------------------
    'II22 = CONTROL.Enabled
    If Err.Number = 0 Then
        II22 = Control.Visible
    End If
    
    If Err.Number = 0 And II22 = False Then
        If Control.Name = "chkOnTop" Then II22 = Control.Enabled
    End If
    'Control.Name
    
    If (II22 = True) And Err.Number = 0 Then
'        Debug.Print Control.Caption
    
        If Control.Left + Control.width > WX Then
            WX = Control.Left + Control.width
        End If
    
        If Control.Top + Control.height > HY Then
            HY = Control.Top + Control.height
'            Debug.Print Control.Name + " -- " + Str(HY)
            TT = Control.Name
            TTHY1 = Str(Val(Control.Top + Control.height))
            TTHY2 = Str(Val(Control.Top))
            TTHY3 = Str(Val(Control.height))
        End If
    End If
Next

If RUN_ONCE_DEBUG_PRINT_HY = False Then
    RUN_ONCE_DEBUG_PRINT_HY = True
    Debug.Print vbCrLf & Time & "-------"
    Debug.Print "Control.Name __ " + TT
    Debug.Print "Control.Top + Control.Height __ " + TTHY1
    Debug.Print "Control.Top __ " + TTHY2
    Debug.Print "Control.Height __ " + TTHY3
End If


'HEIGHT ADJUSTING _______________________________
OFFSET_HY = 500 + 240
OFFSET_WX = 250
If IsIDE = True Then
    OFFSET_HY = 500 + 100
    OFFSET_WX = OFFSET_WX - 150

End If
Dim SETTING_WIDTH_LISTVIEW
SETTING_WIDTH_LISTVIEW = False
If GetComputerName = "1-ASUS-X5DIJ" Then
    OFFSET_WX = 80
    OFFSET_HY = 500 + 240
End If
If GetComputerName = "2-ASUS-EEE" Then
    OFFSET_WX = 80
    OFFSET_HY = 500 + 240
End If

HY2 = HY + ((Menu_Height * Screen.TwipsPerPixelY) - Menu_Height) + OFFSET_HY
WX2 = WX + (Screen.TwipsPerPixelX) + OFFSET_WX

If FORM_LOAD_TRUE = False Then NOT_RESIZE_EVENTER = True
If HY2 <> Me.height Then
    Me.height = HY2
End If

ENDER:

End Sub





Private Sub Form_Resize_OLD()
    
    Dim Me_Top, hWndTask, hWndResult
    
    If SIZE_SET = True Then Exit Sub
    SIZE_SET = True
    'Taskbar at the Bottom
    '---------------------
    Dim RectTaskbar As RECT
    hWndTask = FindWindow("Shell_TrayWnd", vbNullString)
    hWndResult = GetWindowRect(hWndTask, RectTaskbar)
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
        If Err.Number = 0 Then
            EXIT_TRUE = True
            Exit For
        End If
    End If
Next Form
For Each Form In Forms
    Form.EXIT_TRUE = EXIT_TRUE
Next Form

If IsIDE = False Then
    If Me.WindowState <> vbMinimized And Me.EXIT_TRUE = False Then
        Me.WindowState = vbMinimized
        Cancel = True
        Exit Sub
    End If
End If

Dim i, Control

'SET ALL TIMERS IN ALL FORMS ENABLED TO =FALSE
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

If Me.EXIT_TRUE = True Then Cancel = False

For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next Form

'End

End Sub

Private Sub Label29_Click()
'PROCESS_TO_KILLER PID /F *

'PROCESS_TO_KILLER

If LISTVIEW_2_OR_3_HITT = 3 Then
    If lstProcess_3_SORTER_ListView.ListItems.Count > 0 Then
        If lstProcess_3_SORTER_ListView.SelectedItem.Index <> 0 Then
        PROCESS_TO_KILLER_PID = lstProcess_3_SORTER_ListView.ListItems(lstProcess_3_SORTER_ListView.SelectedItem.Index)
        PROCESS_TO_KILLER = lstProcess_3_SORTER_ListView.ListItems(lstProcess_3_SORTER_ListView.SelectedItem.Index).SubItems(1)
        End If
    End If
End If
If LISTVIEW_2_OR_3_HITT = 2 Then
    If lstProcess_2_ListView.ListItems.Count > 0 Then
        If lstProcess_2_ListView.SelectedItem.Index <> 0 Then
        PROCESS_TO_KILLER_PID = lstProcess_2_ListView.ListItems(lstProcess_2_ListView.SelectedItem.Index)
        PROCESS_TO_KILLER = lstProcess_2_ListView.ListItems(lstProcess_2_ListView.SelectedItem.Index).SubItems(1)
        End If
    End If
End If


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


Private Sub Label_KILL_AUTOHOTKEY_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_KILL_AUTOHOTKEY.BackColor = RGB(255, 255, 255)

Call KILL_AUTOHOTKEY_GLOBAL

Me.WindowState = vbMinimized

End Sub

Sub KILL_AUTOHOTKEY_GLOBAL()

SET_COMPUTER_TO_RUN_PID_EXE = "AutoHotkey.exe"
If SET_COMPUTER_TO_RUN <> "" Then
    Call CREATE_PATH_ARRAY_SET_NETWORK_ALL_SPEICAL_REQUEST
    SET_COMPUTER_TO_RUN = ""
    Exit Sub
End If

Dim R, A1, A2

' DO 1ST FOR SPEEDER

For R = 1 To lstProcess_3_SORTER_ListView.ListItems.Count
    A1 = lstProcess_3_SORTER_ListView.ListItems.Item(R).SubItems(1)
    If InStr(A1, SET_COMPUTER_TO_RUN_PID_EXE) > 0 Then
        pid = Val(lstProcess_3_SORTER_ListView.ListItems.Item(R))
        cProcesses.Process_Kill (pid)
    End If
Next

Dim EXECUTE_KILL_1
Dim EXECUTE_KILL_2
Dim EXECUTE_KILL_COUNTER
EXECUTE_KILL_COUNTER = 0
Do
    ' DO EXTRA FOR A GOOD MESSURE
    ENUMPROCESS_MUST_RUNNER = True
    Call EnumProcess
    EXECUTE_KILL_1 = False
    EXECUTE_KILL_COUNTER = EXECUTE_KILL_COUNTER + 1
    For R = 1 To lstProcess_3_SORTER_ListView.ListItems.Count
        A1 = lstProcess_3_SORTER_ListView.ListItems.Item(R).SubItems(1)
        If InStr(A1, SET_COMPUTER_TO_RUN_PID_EXE) > 0 Then
            pid = Val(lstProcess_3_SORTER_ListView.ListItems.Item(R))
            cProcesses.Process_Kill (pid)
            EXECUTE_KILL_1 = True
            EXECUTE_KILL_2 = True
        End If
    Next

Loop Until EXECUTE_KILL_1 = False Or EXECUTE_KILL_COUNTER > 400

If EXECUTE_KILL_COUNTER > 400 Then
    MsgBox "TRY TO KILL-AH ALL " + vbCrLf + SET_COMPUTER_TO_RUN_PID_EXE + vbCrLf + "BUT WAS PROBLEM SOME EXIST AFTER" + Str(EXECUTE_KILL_COUNTER) + " RETRY", vbMsgBoxSetForeground
End If

If EXECUTE_KILL_2 = True Or 1 = 1 Then
    ' Const ShowWindow_2 = 1, DontShowWindow = 0, DontWaitUntilFinished = False, WaitUntilFinished = True
    Dim objShell
    Set objShell = CreateObject("Wscript.Shell")
    objShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 78-TRAY ICON CLEANER - RUN_ONCE.ahk""", DontShowWindow, DontWaitUntilFinished
    Set objShell = Nothing
End If


' Me.WindowState = vbMinimized


End Sub

Sub KILL_WSCRIPT_GLOBAL()

SET_COMPUTER_TO_RUN_PID_EXE = "WSCRIPT.exe"
If SET_COMPUTER_TO_RUN <> "" Then
    Call CREATE_PATH_ARRAY_SET_NETWORK_ALL_SPEICAL_REQUEST
    SET_COMPUTER_TO_RUN = ""
    Exit Sub
End If

Dim R, A1, A2

' DO 1ST FOR SPEEDER

For R = 1 To lstProcess_3_SORTER_ListView.ListItems.Count
    A1 = lstProcess_3_SORTER_ListView.ListItems.Item(R).SubItems(1)
    If InStr(UCase(A1), UCase(SET_COMPUTER_TO_RUN_PID_EXE)) > 0 Then
        pid = Val(lstProcess_3_SORTER_ListView.ListItems.Item(R))
        cProcesses.Process_Kill (pid)
    End If
Next

Dim EXECUTE_KILL_1
Dim EXECUTE_KILL_2
Dim EXECUTE_KILL_COUNTER
Do
    ' DO EXTRA FOR A GOOD MESSURE
    Call EnumProcess
    EXECUTE_KILL_1 = False
    EXECUTE_KILL_COUNTER = EXECUTE_KILL_COUNTER + 1
    For R = 1 To lstProcess_3_SORTER_ListView.ListItems.Count
        A1 = lstProcess_3_SORTER_ListView.ListItems.Item(R).SubItems(1)
        If InStr(UCase(A1), UCase(SET_COMPUTER_TO_RUN_PID_EXE)) > 0 Then
            pid = Val(lstProcess_3_SORTER_ListView.ListItems.Item(R))
            cProcesses.Process_Kill (pid)
            EXECUTE_KILL_1 = True
            EXECUTE_KILL_2 = True
        End If
    Next

Loop Until EXECUTE_KILL_1 = False Or EXECUTE_KILL_COUNTER > 100

If EXECUTE_KILL_COUNTER > 100 Then
    MsgBox "TRY TO KILL-AH ALL " + vbCrLf + SET_COMPUTER_TO_RUN_PID_EXE + vbCrLf + "BUT WAS PROBLEM SOME EXIST AFTER 1-0 RETRY", vbMsgBoxSetForeground
End If

'If EXECUTE_KILL_2 = True Or 1 = 1 Then
'    ' Const ShowWindow_2 = 1, DontShowWindow = 0, DontWaitUntilFinished = False, WaitUntilFinished = True
'    Dim objShell
'    Set objShell = CreateObject("Wscript.Shell")
'    objShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 78-TRAY ICON CLEANER - RUN_ONCE.ahk""", DontShowWindow, DontWaitUntilFinished
'    Set objShell = Nothing
'End If


' Me.WindowState = vbMinimized


End Sub



Private Sub Label23_Click()

'Label23.Caption = "HITT TO CONFIRM SELECTION KILL"

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


Private Sub Label_RUN_AUTOHOTKEY_SET_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_RUN_AUTOHOTKEY_SET.BackColor = RGB(255, 255, 255)

Call MNU_AUTOHOTKEYS_SET_Click

Me.WindowState = vbMinimized

End Sub




Private Sub Label_MAXIMIZE_GOODSYNC_Click()

Dim GOODSYNC_WINDOW_hWnd
GOODSYNC_WINDOW_hWnd = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F009}", vbNullString)

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_MAXIMIZE_GOODSYNC.BackColor = RGB(255, 255, 255)

ShowWindow GOODSYNC_WINDOW_hWnd, SW_MAXIMIZE
Beep
Me.WindowState = vbMinimized

End Sub

Private Sub Label_MINIMIZE_GOODSYNC_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_MINIMIZE_GOODSYNC.BackColor = RGB(255, 255, 255)

Dim GOODSYNC_WINDOW_hWnd
GOODSYNC_WINDOW_hWnd = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F009}", vbNullString)

ShowWindow GOODSYNC_WINDOW_hWnd, SW_MINIMIZE
Beep
Me.WindowState = vbMinimized

End Sub

Private Sub MNU_SHOW_THE_TIME_Click()

'Timer_SHOW_THE_TIME_Timer()
'---------------------------
'MNU_GIVE_ME_TIME_Click()
'---------------------------

Call MNU_GIVE_ME_TIME_Click

End Sub


Private Sub Mnu_Form_Count_Click()
' Mnu_Menu_Item_Count
' Mnu_Form_Count
End Sub

Private Sub MNU_GIVE_ME_TIME_Click()

'Timer_SHOW_THE_TIME_Timer()
'---------------------------
'---------------------------

Dim XI
Dim A_DATE_TIME_UTC
Dim A_DATE_TIME As String

'--------------------------------
'get 24 hour format and add am pm
'--------------------------------
'or put back as 12 hour format
'WATCH CLOCK GO BACK PROBLEM
'--------------------------------
A_DATE_TIME = "[ " + Format(Now, "DDDD HH:MM:SS") + " " + A_DATE_TIME_PM + "_" + Format(Now, "DD MMMM YYYY") + " ]"
If Hour(Now) = 16 Then
    A_DATE_TIME = "[ " + Format(Now, "DDDD HH:MM:SS AM/PM") + "_" + Format(Now, "DD MMMM YYYY") + " ]"
End If
A_DATE_TIME = Replace(A_DATE_TIME, "PM", "Pm")
A_DATE_TIME = Replace(A_DATE_TIME, "AM", "Am")
'------------------------------------------------------
'REPLACE LAST DIGIT AS 0
'----------------------------
XI = InStr(A_DATE_TIME, "m_")
Mid(A_DATE_TIME, XI - 3, 1) = "0"
'----------------------------
'------------------------------------------------------
'------------------------------------------------------
Dim GMTBST
'Convert_To_UT
'-------------
'-----------------------------------------------------------------------------------------------
'ONE MINUTE ADJUST OFFSET IF DELAY IN TWO TIME WHEN SAME CALCULATOR MESSUREMENT THEN MAYBE ERROR
'OFFSET BY HALF 30 MINUTE
'-----------------------------------------------------------------------------------------------
If ConvertToUT(Now) + TimeSerial(0, 30, 0) < Now Then
    GMTBST = "BST BDT DST DLS 1 Hour Ahead UTC+1 CET"
Else
    GMTBST = "GMT UTC+0 WET"
End If

'--------------------------------
'get 24 hour format and add am pm
'--------------------------------
'or put back as 12 hour format
'--------------------------------
A_DATE_TIME_UTC = "[ " + Format(ConvertToUT(Now), "DDDD HH:MM:SS") + " " + A_DATE_TIME_PM + "_" + Format(ConvertToUT(Now), "DD MMMM YYYY") + " ]"
If Hour(Now) = 16 Then
    A_DATE_TIME_UTC = "[ " + Format(ConvertToUT(Now), "DDDD HH:MM:SS AM/PM") + "_" + Format(ConvertToUT(Now), "DD MMMM YYYY") + " ]"
End If
A_DATE_TIME_UTC = Replace(A_DATE_TIME_UTC, "PM", "Pm")
A_DATE_TIME_UTC = Replace(A_DATE_TIME_UTC, "AM", "Am")
'------------------------------------------------------
'REPLACE LAST DIGIT AS 0
'----------------------------
XI = InStr(A_DATE_TIME_UTC, "m_")
Mid(A_DATE_TIME_UTC, XI - 3, 1) = "0"
'----------------------------
If Day(ConvertToUT(Now)) = 7 - 1 Then A_DATE_TIME_UTC = Replace(A_DATE_TIME_UTC, Format(ConvertToUT(Now), "DD MMMM YYYY"), "Sixer " + Format(ConvertToUT(Now), "MMMM YYYY"))
If Day(Now) = 7 - 1 Then A_DATE_TIME = Replace(A_DATE_TIME, Format(Now, "DD MMMM YYYY"), "Sixer " + Format(Now, "MMMM YYYY"))


If InStr(MNU_GIVE_ME_TIME_WITHER_UTC.Caption, "=_YES") Then
    A_DATE_TIME = A_DATE_TIME + " The UK" + vbCrLf + A_DATE_TIME_UTC + " " + GMTBST
Else
    A_DATE_TIME_UTC = ""
End If
'------------------------------------------------------

Clipboard.Clear
Clipboard.SetText A_DATE_TIME

Beep
Me.WindowState = vbMinimized

End Sub

Public Function ConvertToUT(Time As Date) As Date
' convert system time to universal time and adjust for DST
Dim TZ As TIME_ZONE_INFORMATION

'01 of 02 -- SEARCH SAME REM LINE
'------------------------------------------------------------------------------------
'ERROR OF ORGINAL SOURCE -- SUPPOSED TO BE TEMP AS RETURN STRING IN FORM OF --- (GMT)
'OR OPTION HAS THAT IN TEXT STRING AND EXTRA TIME INFO
'BUT TEMP AS DATE -- IS USED THROUGH OUT THIS CODE
'------------------------------------------------------------------------------------
'ALSO MY COMPUTER IS UNABLE TO RETURN INFO ABOUT TIME ZONE -- WIN 10
'------------------------------------------------------------------------------------
'WORK AROUND ADD A ISNUMERIC DETECT

'TIMEZONE IS RETURN AS STRING AND NUMERIC IS ADJUST IS MADE AT GMT
'BUT NOT CORRECT TOWARD THE ORGINAL SOURCE CODE METHOD

'------------------------------------------------------------------------------------
'METHOD WORKED OUT
'DON'T DO STRING CONVERT METHOD TO NUMERIC
'AND MEAN ADD ANOTHER FUNCTION HERE
'------------------------------------------------------------------------------------

Dim temp As Date
Dim temp_str As String
Dim TimeZone_Numeric

Dim Rtn As Long
Rtn = GetTimeZoneInformation(TZ)
If Rtn > TIME_ZONE_ID_UNKNOWN Then

    If Rtn = TIME_ZONE_ID_STANDARD Then
    temp = DateAdd("n", TZ.Bias, Time)
    Else
    temp = DateAdd("n", (TZ.Bias + TZ.DaylightBias), Time)
    End If

End If
ConvertToUT = temp
'Temp = TimeZone
temp = TimeZone_Numeric

'Debug.Print IsNumeric(Mid("00:00:00", 1, 2))
'----------------------------------------
'Debug.Print IsNumeric(TimeZone_Numeric)
'----------------------------------------
'A DATE VAR IS RETRUNED AS NUMERIC = TRUE
'----------------------------------------
'----------------------------------------
'Debug.Print IsNumeric(TimeZone)

'If IsNumeric(IsNumeric(Mid(TimeZone, 1, 2))) Then
'    temp = TimeZone
'Else
'    temp_str = TimeZone
'End If

End Function

Private Sub MNU_HOOVER_20_SECOND_Click()
TIMER2_TIMER_BEGAN = Now
End Sub


Private Sub Timer_EnumProcess_Timer()
'EnumProcess_COUNTER
'SUB ENUMPROCESS
Call EnumProcess
Timer_EnumProcess.Interval = 60000

'Label15.Caption = "Scan Processor For A Moment Same As Other Command Set Timer Second _ " + Trim(Str(EnumProcess_COUNTER))
'Label15.Caption = "Scan Processor For A Moment Timer Second _ " + Trim(Str(EnumProcess_COUNTER))
'Label15.Caption = "Scan Processor For A Moment Timer Second _ " + Trim(Str(EnumProcess_COUNTER)) + " && Foregound Window" ' Change"

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

If FOREGROUND_WINDOW_CHANGE_DELAY_1 > Now Then
    FOREGROUND_WINDOW_CHANGE_DELAY_1_EXTRA_TO_DO.Enabled = True
    Exit Sub
End If
FOREGROUND_WINDOW_CHANGE_DELAY_1 = Now + TimeSerial(0, 0, 2)

Call MAX_FORM_CHECKER
' [ Wednesday 08:47:20 Am_12 June 2019 ]

Call Timer_VB_MAXIMIZE_Timer

Dim I_hWnd
Dim hWnd_MediaPlayerClassicW_GetWindowState
Dim hWnd_WINAMP_GetWindowState


I_hWnd = FindWindow("MediaPlayerClassicW", vbNullString)
If OLD_hWnd_MediaPlayerClassicW_GetWindowState <> hWnd_MediaPlayerClassicW_GetWindowState Then
    If I_hWnd > 0 Then
        SetWindowPos I_hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End If
OLD_hWnd_MediaPlayerClassicW_GetWindowState = hWnd_MediaPlayerClassicW_GetWindowState

I_hWnd = FindWindow("Winamp v1.x", vbNullString)

hWnd_WINAMP_GetWindowState = GetWindowState(I_hWnd)
If OLD_hWnd_WINAMP_GetWindowState <> hWnd_WINAMP_GetWindowState Then
    If I_hWnd > 0 Then
        SetWindowPos I_hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End If
OLD_hWnd_WINAMP_GetWindowState = hWnd_WINAMP_GetWindowState
End Sub




Private Sub Label_KILL_WSCRIPT_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_KILL_WSCRIPT.BackColor = RGB(255, 255, 255)

Dim R, A1, A2
Dim ALL_DONE
Dim NAME_EXE As String
Dim PID_INPUT As Long
Dim i
Do
    ALL_DONE = True
    SET_COMPUTER_TO_RUN_PID_EXE = UCase("WSCRIPT.EXE")
    For R = lstProcess_3_SORTER_ListView.ListItems.Count To 1 Step -1
        A1 = lstProcess_3_SORTER_ListView.ListItems.Item(R).SubItems(1)
        If InStr(UCase(A1), SET_COMPUTER_TO_RUN_PID_EXE) > 0 Then
            pid = Val(lstProcess_3_SORTER_ListView.ListItems.Item(R))
            i = cProcesses.Process_Kill(pid)
            ALL_DONE = False
        End If
    Next
    
    ENUMPROCESS_MUST_RUNNER = True
    Call EnumProcess
Loop Until ALL_DONE = True

Me.WindowState = vbMinimized

End Sub

Private Sub Label_CLOSE_EXPLORER_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_CLOSE_EXPLORER.BackColor = RGB(255, 255, 255)

Dim R, A1, A2
Dim ALL_DONE
Dim NAME_EXE As String
Dim PID_INPUT As Long
Dim i
Dim pid As Long
Dim hWnd_LINE As Long
Dim HWND_STR As String
Dim strarray As Variant

'PASTE THE CURRENT SESSION TO CLIPBOARD __ QUITELY WITHOUT MSGBOX REPSONCE REQUIRING
'----------------------------------------------------
Call FindWindow_Get_All_Explorer("QUITE MSGBOX=TRUE")
'----------------------------------------------------

HWND_STR = FindWinPart_EXPLORER
If HWND_STR <> "" Then
    strarray = Split(HWND_STR, vbCrLf)
    For R = 0 To UBound(strarray)
        hWnd_LINE = Val(strarray(R))
        If hWnd_LINE > 0 Then
            Result = PostMessage(hWnd_LINE, WM_CLOSE, 0&, 0&)
        End If
    Next
End If

Beep
Me.WindowState = vbMinimized

End Sub

Private Sub Label_DE_DUPE_EXPLORER_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_DE_DUPE_EXPLORER.BackColor = RGB(255, 255, 255)

Dim R, A1, A2
Dim ALL_DONE
Dim NAME_EXE As String
Dim PID_INPUT As Long
Dim i
Dim pid As Long
Dim hWnd_LINE As Long
Dim HWND_STR As String
Dim strarray As Variant
Dim TITLE_COMPARE_1
Dim TITLE_COMPARE_2

'PASTE THE CURRENT SESSION TO CLIPBOARD __ QUITELY WITHOUT MSGBOX REPSONCE REQUIRING
'----------------------------------------------------
Call FindWindow_Get_All_Explorer("QUITE MSGBOX=TRUE")
'----------------------------------------------------

HWND_STR = FindWinPart_EXPLORER
If HWND_STR <> "" Then
    strarray = Split(HWND_STR, vbCrLf)
    For R = 0 To UBound(strarray)
        hWnd_LINE = Val(strarray(R))
        TITLE_COMPARE_1 = "----" + UCase(GetWindowTitle(hWnd_LINE)) + "++++"
        If InStr(TITLE_COMPARE_2, TITLE_COMPARE_1) > 0 Then
            If hWnd_LINE > 0 Then
                Result = PostMessage(hWnd_LINE, WM_CLOSE, 0&, 0&)
            End If
        End If
        TITLE_COMPARE_2 = TITLE_COMPARE_2 + "----" + UCase(GetWindowTitle(hWnd_LINE)) + "++++"
    Next
End If

Beep
Me.WindowState = vbMinimized

End Sub
Private Sub MNU_AUTOHOTKEY_STARTING_Click()
' Const ShowWindow_2 = 1, DontShowWindow = 0, DontWaitUntilFinished = False, WaitUntilFinished = True
Dim objShell
Set objShell = CreateObject("Wscript.Shell")
objShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 21-AUTORUN.ahk""", ShowWindow_2, DontWaitUntilFinished
Set objShell = Nothing
Me.WindowState = vbMinimized
End Sub

Private Sub MNU_AUTOHOTKEYS_SET_Click()

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-AUTOHOTKEYS SET RELAUNCH CODE.ahk" + """"
Set WSHShell = Nothing
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

Dim GOODSYNC_WINDOW_hWnd
GOODSYNC_WINDOW_hWnd = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F009}", vbNullString)

ShowWindow GOODSYNC_WINDOW_hWnd, SW_MAXIMIZE
Beep
Me.WindowState = vbMinimized

End Sub


'----
'Msgbox With Timer - VB6 | Dream.In.Code
'https://www.dreamincode.net/forums/topic/59064-msgbox-with-timer/
'----

'Private Declare Function MessageBoxTimeout Lib "user32.dll" Alias "MessageBoxTimeoutA" ( _
'ByVal hWnd As Long, _
'ByVal lpText As String, _
'ByVal lpCaption As String, _
'ByVal uType As Long, _
'ByVal wLanguageId As Long, _
'ByVal dwMilliseconds As Long _
') As Long
'
'Private Const IDYES& = 6&
'Private Const IDNO& = 7&
'Private Const MB_SETFOREGROUND& = &H10000
'Private Const MB_YESNO& = &H4&
'Private Const MB_ICONASTERISK& = &H40&
'Private Const MB_TIMEDOUT& = &H7D00&

Private Sub Form_Load_555()
Dim lRet As Long
lRet = MessageBoxTimeout(0&, "your message", _
"Timeout MessageBox Test", MB_SETFOREGROUND Or MB_YESNO Or MB_ICONASTERISK, _
0&, 90000)

Select Case lRet
Case IDYES:
'// your code
Case IDNO:
'// your code

Case MB_TIMEDOUT:
'// your code

End Select
End Sub

Public Sub TIMER_OS_REBOOT_Timer()

' -------------------------------------
' TRY TO DO A COUNTDOWN TIMER IN MSGBOX
' BUT FAIL
' IF CLICK MNU ITEM
' OTHER TIMER ARE NOT RUN
' USE AHK TO VB6 INSTEAD
' -------------------------------------


Dim VB_LOADER, PWnd, GS_cWnd1, GS_cWnd2, R_REPEAT
VB_LOADER = FindWindow("#32770", "VB_KEEP_RUNNER")
If VB_LOADER > 0 Then 'And VB_LOADER <> OhWnd_VB_LOADER Then
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

End Sub

Private Sub MNU_OS_RESTART_Click()
Dim ihWnd
ihWnd = FindWindow("wndclass_desked_gsk", vbNullString)
ShowWindow ihWnd, SW_MINIMIZE
Me.WindowState = vbMinimized

Dim lRet
lRet = MsgBox("READY FOR OS RESTART" + vbCrLf + "YES COUNTDOWN" + vbCrLf + "With a default of 30" + vbCrLf + "If the timeout period is" + vbCrLf + "Set greater than 0 The /f Force parameter is implied.", vbYesNo + vbMsgBoxSetForeground)
If lRet = vbNo Then
    Beep
    Exit Sub
End If

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "C:\WINDOWS\system32\shutdown.exe" + """ -r", 0, True
Set WSHShell = Nothing
Beep
Me.WindowState = vbMinimized
End

' lRet = MessageBoxTimeout(0&, "READY FOR OS RESTART", _
"Timeout MessageBox Test", MB_SETFOREGROUND Or MB_YESNO Or MB_ICONASTERISK, _
0&, 90000)


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
    WSHShell.Run """" + "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 31-AUTORUN SET FAV VB & AUTOHOTKEY.ahk" + """"
Set WSHShell = Nothing
Beep
Me.WindowState = vbMinimized
End Sub

Private Sub MNU_GOODSYNC_COLLECTION_SCRIPT_RUN_Click()

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 59-RUN GOOD_SYNC SET SCRIPTOR.BAT" + """"
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
    ' Const ShowWindow_2 = 1, DontShowWindow = 0, DontWaitUntilFinished = False, WaitUntilFinished = True
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
    ' Const ShowWindow_2 = 1, DontShowWindow = 0, DontWaitUntilFinished = False, WaitUntilFinished = True
    Dim objShell
    Set objShell = CreateObject("Wscript.Shell")
    RUN_EXE = "D:\VB6\VB-NT\00_Best_VB_01\10 SYNCRONIZE\SYNCRONIZER.exe"
    objShell.Run """" + RUN_EXE + """", 1, False
    Set objShell = Nothing
End Sub

Private Sub MNU_LAUNCH_AUTORUNS_SET_BOOT_Click()
' Const ShowWindow_2 = 1, DontShowWindow = 0, DontWaitUntilFinished = False, WaitUntilFinished = True
Dim objShell
Set objShell = CreateObject("Wscript.Shell")

objShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 21-AUTORUN.ahk""", ShowWindow_2, DontWaitUntilFinished
    
Set objShell = Nothing
Me.WindowState = vbMinimized
End Sub


Private Sub ONE_MILLISECOND_Timer_Timer()

ONE_MILLISECOND_Timer.Enabled = False

' ----------------------------------------------------------------------------------------
' APP THAT START IN WINDOWS 10
' AND THEY BEGIN MINIMIZED THAT IS GOOD FOR TASKBAR
' IF EVER NOTICE SOMETIME HITT THE TASKBAR BUTTON AND THE TASKBAR ITEM MOVE SOMEWHERE ELSE
' NOT IF MINIMIZED
' BUT ONLY IF START UP SHOWER
' AND TASKBAR ITEM THAT WHAT WANT
' AND THEN FIND A WAY
' ----------------------------------------------------------------------------------------
' [ Monday 10:04:30 Am_27 May 2019 ]
' ----------------------------------------------------------------------------------------
' WELL STILL HAVE A PROBLME LIKE HERE
' IT SEEM THEN THAT IS TASKBAR PROPERTIE HAS AND EXTRA PARAMETER
' THAT MESSES IT UP
' ----------------------------------------------------------------------------------------
' STILL SAME PROBLEM
' AND THEN IT SEEM
' ONCE YOU MODIFY THAT TASKBAR PROPERTY TO ADD PARAMERTER IT MESSES WITH THAT ONE
' ----------------------------------------------------------------------------------------



If DO_ONCE_BOOTER = True Then Exit Sub
DO_ONCE_BOOTER = True


Dim OR_LOGIC
OR_LOGIC = 0
If InStr(Command$, "MINIMAL") > 0 Then OR_LOGIC = 2
If InStr(Command$, "MAXIMUM") > 0 Then OR_LOGIC = 1
If InStr(Command$, "TASKBAR") > 0 Then OR_LOGIC = 1
If InStr(Command$, "T") > 0 Then OR_LOGIC = 1
If IsIDE = True Then OR_LOGIC = 1
' Me.Visible = True
' DoEvents
If OR_LOGIC = 1 Then
    Me.WindowState = vbNormal
    
    ' MsgBox "HH"
    
    ' Exit Sub
End If

If OR_LOGIC = 1 Then
    ' Me.WindowState = vbNormal
    Exit Sub
End If



'Exit Sub
'
'Call WIDTH_AND_HEIGHT(WX, HY)
'Me.height = HY2 + 350

'If IsIDE = True Then
'    Me.WindowState = vbNormal
'Else
'    Me.WindowState = vbMinimized
'End If

'Me.WindowState = vbMinimized
'Me.Visible = True

End Sub

Private Sub ONE_MILLISECOND_2_Timer_Timer()

If QUICK_KEY_ESCPAPE_AT_LOAD_20_SECOND > Now Then
    If GetAsyncKeyState(27) Then End 'ESCAPE KEY
End If
If QUICK_KEY_ESCPAPE_AT_LOAD_20_SECOND < Now Then
    ONE_MILLISECOND_2_Timer.Enabled = False
End If

End Sub


Private Sub lstProcess_2_ListView_KeyDown(KeyCode As Integer, Shift As Integer)
LISTVIEW_2_OR_3_HITT = 2
If IsIDE = True And KeyCode = (27) Then End
End Sub

Private Sub lstProcess_3_SORTER_ListView_BeforeLabelEdit_NULL(Cancel As Integer)
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
    
    If Menu_Height <> ARCHIVE_Menu_Height Then
        TIMER_TO_RESIZE.Enabled = True
    End If
    
    ARCHIVE_Menu_Height = Menu_Height
   
End Function

Private Sub TIMER_TO_RESIZE_Timer()
    TIMER_TO_RESIZE.Enabled = False
    'NOT_RESIZE_EVENTER = False
    Call Form_Resize
    'NOT_RESIZE_EVENTER = True

End Sub
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
    'SetWindowPos txtMhWnd.Text, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    SetWindowPos Me.hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    Timer_ALWAYS_ON_TOP_TO_START_WITH_ER.Enabled = False
    Label60.BackColor = Label49.BackColor '49 58_59
    Label60.Caption = "Me on Top 20 Sec NOT"
    MNU_ME_ON_TOP.Caption = "[__ ME ON TOP = YES __]"
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



'////////////////////////////////////////////////////////////////////
'//// CROSSHAIR EVENTS
'////////////////////////////////////////////////////////////////////
Private Sub picCrossHair_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If user pressed left mouse button and we are not dragging
    If Button = vbLeftButton And Not m_bDragging Then
        picCrossHair_MouseMove_Dragging_VAR = True

        ' Set dragging flag to true
        m_bDragging = True
        ' Set mouse pointer
        Me.MouseIcon = imgCursor.MouseIcon
        Me.MousePointer = 99
        ' Erase picture from picCrossHair
        picCrossHair.Picture = Nothing
    End If
    
    If Button = vbLeftButton And m_bDragging Then
        picCrossHair_MouseMove_Dragging_VAR = True
        picCrossHair_MouseMove_02
    End If
    
End Sub


Private Sub picCrossHair_MouseMove_02()

    If picCrossHair_MouseMove_Dragging_VAR = False Then Exit Sub

    Dim tPA As POINTAPI
    
    ' Get cursor cordinates
    GetCursorPos tPA
    ' Set label caption to cursor cordinates
    'lblCordi.Caption = "X: " & tPA.X & "  Y: " & tPA.Y
    
    If tPA.y = 0 Or tPA.y < (Me.Top / Screen.TwipsPerPixelY) Then
        NOT_RESIZE_EVENTER = True
        Me.WindowState = vbNormal
        'Me.Hide
        'DoEvents
        
        If Me.height <> 1500 And Me.width <> 1500 Then
            NOT_RESIZE_EVENTER_SET_HAPPEN_ME_XW = Me.height
            NOT_RESIZE_EVENTER_SET_HAPPEN_ME_YH = Me.width
            Me.height = 1500
            Me.width = 1500
        End If
        
        
        NOT_RESIZE_EVENTER = False
    End If
    
    'Exit Sub
    ' If user pressed left mouse button and we are dragging
    
'    If Button = vbLeftButton And m_bDragging Then
'    End If

End Sub


Private Sub picCrossHair_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    picCrossHair_MouseMove_02
    
    'Exit Sub
    ' If user pressed left mouse button and we are dragging
    
'    If Button = vbLeftButton And m_bDragging Then
'    End If

End Sub

Private Sub picCrossHair_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If user pressed left mouse button and we are dragging
    If Button = vbLeftButton And m_bDragging Then
        ' Set dragging flag to true
        m_bDragging = False
        picCrossHair_MouseMove_Dragging_VAR = False

        ' Restore mouse pointer to normal (arrow)
        Me.MousePointer = vbNormal
        ' Load picture into picCrossHair
        picCrossHair.Picture = imgCursor.MouseIcon
        FROM_picCrossHair_MouseUP = True
        Call ChunkCodeOnMouse
        
        'If Me.WindowState <> vbMaximized Then
        '    Me.WindowState = vbMaximized
        'End If
        'If Me.Visible = False Then
        '    Me.Visible = True
        'End If
        
        'WIDTH MAYBE CHANGE BY SOMETHING THAN 1500
        If Me.height = 1500 And Me.width < 2000 Then
            Me.height = NOT_RESIZE_EVENTER_SET_HAPPEN_ME_XW
            Me.width = NOT_RESIZE_EVENTER_SET_HAPPEN_ME_YH
        End If
        
        If NOT_RESIZE_EVENTER_SET_HAPPEN_ME_XW > 0 Then
            Me.height = NOT_RESIZE_EVENTER_SET_HAPPEN_ME_XW
            Me.width = NOT_RESIZE_EVENTER_SET_HAPPEN_ME_YH
            NOT_RESIZE_EVENTER_SET_HAPPEN_ME_XW = 0
            NOT_RESIZE_EVENTER_SET_HAPPEN_ME_YH = 0
        End If
        
        If InStr(Command_Screen_Shot_Auto_ClipBoard_er.Caption, "Archive Mode _ON") > 0 Then
            'FORM
            Call SCREEN_SHOT_HERE(txthWnd.Text)
        End If

        'FORM WHEN BEHIND ANOTHER
        'WORK TO DO
        'Call SCREEN_SHOT_HERE_2(txthWnd.Text)
        
        
    End If
End Sub

Sub ChunkCodeOnMouse()
        
        Dim tRC2 As RECT
        Dim tPA As POINTAPI
        Dim LhWnd As Long
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
            If hWnd_From_ListView = 0 And lhWnd_Function_Button_Set_MIN_MAX = 0 Then
                ' Get cursor position
                GetCursorPos tPA
                ' Get window handle from point
                LhWnd = WindowFromPoint(tPA.x, tPA.y)
                ' Get window caption
            End If
        End If
        If FROM_picCrossHair_MouseUP = True Then
            FROM_picCrossHair_MouseUP = False
            If LhWnd > 0 Then
                O_lhWndParent = LhWnd
                lhWndParent = GetParent(LhWnd)
                If lhWndParent = 0 Then lhWndParent = O_lhWndParent
                lhWndParentX = GetParenthWnd(LhWnd)

                Success_Result = cProcesses.Get_PID_From_hWnd(lhWndParentX, PID_MARK)
                TxtPID.Text = PID_MARK
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
                LhWnd = hWnd_From_ListView
                TxtPID.Text = Val(PROCESS_TO_KILLER_PID)
                hWnd_From_ListView = 0
                Set_Collect_More_Info = False
            End If
        End If
        If From_ListView = True And hWnd_From_ListView = 0 Then
            Set_Collect_More_Info = False
        End If
        
        
        If lhWnd_Function_Button_Set_MIN_MAX > 0 Then
            ' txthWnd.Text = lhWnd
            LhWnd = lhWnd_Function_Button_Set_MIN_MAX
            lhWnd_Function_Button_Set_MIN_MAX = 0
            ' Set_Collect_More_Info = false
        End If
        
        
        
        
        
        
        GetWindowRect LhWnd, tRC2
        
        lRetVal = GetWindowText(LhWnd, sTitle, 255)
        ' Get window class name
        lRetVal = GetClassName(LhWnd, sClass, 255)
        ' Get window style
        sStyle = GetWindowStyle(LhWnd)
        ' Get window parent
        
        O_lhWndParent = LhWnd
        lhWndParent = GetParent(LhWnd)
        If lhWndParent = 0 Then lhWndParent = O_lhWndParent
        lhWndParentX = GetParenthWnd(LhWnd)
        
        
        If Set_Collect_More_Info = True Then
            Success_Result = cProcesses.Get_PID_From_hWnd(lhWndParentX, PID_MARK)
            
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
        
            TxtEXE.Text = GetFileFromhWnd(LhWnd)
            
            
            PROCESS_TO_KILLER = Mid(TxtEXE.Text, InStrRev(TxtEXE.Text, "\") + 1)
            PROCESS_TO_KILLER_PID = TxtPID.Text
        
            Call MOUSE_HOOVER_SLECTION_CLICKER
        
        End If
        
        From_ListView = False
        
        
        txthWnd.Text = LhWnd
        txthWndHX.Text = Hex(LhWnd)
        txtTitle.Text = sTitle
        txtClass.Text = sClass
        txtStyle.Text = sStyle
        txtRect.Text = "(" & tRC2.Left & ", " & tRC2.Top & ") - (" & tRC2.Right & ", " & tRC2.Bottom & ")"
        
        'lhWndParent = GetParent(lhWndParent)
        'lhWndParentHX = Hex(GetParent(lhWndParent))
        If LhWnd <> lhWndParent Then
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
        
        If LhWnd <> lhWndParentX And lhWndParentX <> lhWndParent Then
            txtParentX.Text = lhWndParentX
            txtParentTextX.Text = sParentTitleX
            txtParentClassX.Text = sParentClassX
        Else
            txtParentX.Text = lhWndParentX
            txtParentTextX.Text = ""
            txtParentClassX.Text = ""
        End If
        
        
        txtMhWnd.Text = LhWnd
        
        If TxtEXE.Text <> OLD_TxtEXE_Text Then
            Call TxtEXE_CLICK
        End If
        OLD_TxtEXE_Text = TxtEXE.Text


End Sub

Private Sub SCREEN_SHOT_HERE_2(hWnd_NUMBER)
    
    If InStr(Command_Screen_Shot_Auto_ClipBoard_er.Caption, "Archive Mode _ON") = 0 Then
        Exit Sub
    End If
    
    Dim TT, TS, FOLDERNAME_AUTO
    Dim hWnd_LONG As Long
    
    FOLDERNAME_AUTO = App.Path + "\# DATA\" + GetComputerName + "\"
    
    If RUN_ONCE_VB_DIRECTORY = False Then
        RUN_ONCE_VB_DIRECTORY = True
        If Dir(FOLDERNAME_AUTO, vbDirectory) = "" Then
            On Error Resume Next
                MkDir App.Path + "\# DATA\"
                MkDir FOLDERNAME_AUTO
        End If
    End If
    
    'TT = PrintCurrentFormOntoForm(SCREEN_CAP)
    
    hWnd_LONG = Val(hWnd_NUMBER)
    
    ' --------------------
    ' SCREEN_CAP IS A FORM
    ' --------------------
    
    'TT = Print_hWnd_FORM_ontoForm(SCREEN_CAP, Val(hWnd_NUMBER))
    TT = Print_hWnd_FORM_ontoForm_EVEN_WHEN_BEHIND_ANOTHER(SCREEN_CAP, hWnd_LONG)
    TS = FOLDERNAME_AUTO + "FormCapture- " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + ".jpg"
    'MsgBox TS
    On Error Resume Next
    'SavePicture SCREEN_CAP.Picture, TS
    'ERR.NUMBER
    On Error GoTo 0
    
    Clipboard.Clear
    Sleep 800
    DoEvents
    
    Clipboard.SetData SCREEN_CAP.Picture, vbCFBitmap
    Beep
    SCREEN_CAP.Show
    
    '= Clipboard.GetData(vbCFBitmap)

'----'[RESOLVED]
'Capture Picture Of A Form When A Portion Of It Is Off The Screen...-VBForums
'http://www.vbforums.com/showthread.php?734947-RESOLVED-Capture-Picture-Of-A-Form-When-A-Portion-Of-It-Is-Off-The-Screen
'----


End Sub

Private Sub SCREEN_SHOT_HERE(hWnd_NUMBER)
    Dim TT, TS, FOLDERNAME_AUTO
    
    If InStr(Command_Screen_Shot_Auto_ClipBoard_er.Caption, "Archive Mode _ON") = 0 Then
        Exit Sub
    End If
    
    FOLDERNAME_AUTO = App.Path + "\# DATA\" + GetComputerName + "\"
    
    If RUN_ONCE_VB_DIRECTORY = False Then
        RUN_ONCE_VB_DIRECTORY = True
        If Dir(FOLDERNAME_AUTO, vbDirectory) = "" Then
            On Error Resume Next
                MkDir App.Path + "\# DATA\"
                MkDir FOLDERNAME_AUTO
        End If
    End If
    
    'TT = PrintCurrentFormOntoForm(SCREEN_CAP)
    
    
    TT = Print_hWnd_FORM_ontoForm(SCREEN_CAP, Val(hWnd_NUMBER))
    TS = FOLDERNAME_AUTO + "EliteSPY_FormCapture_ " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + ".jpg"
    'MsgBox TS
    On Error Resume Next
    SavePicture SCREEN_CAP.Picture, TS
    'ERR.NUMBER
    On Error GoTo 0
    
    Clipboard.Clear
    Sleep 800
    DoEvents
    
    Clipboard.SetData SCREEN_CAP.Picture, vbCFBitmap
    Beep
    SCREEN_CAP.Show
    
    '= Clipboard.GetData(vbCFBitmap)

'----'[RESOLVED]
'Capture Picture Of A Form When A Portion Of It Is Off The Screen...-VBForums
'http://www.vbforums.com/showthread.php?734947-RESOLVED-Capture-Picture-Of-A-Form-When-A-Portion-Of-It-Is-Off-The-Screen
'----


End Sub

Private Sub Command_Screen_Shot_Auto_ClipBoard_er_Click()

'Screen_Shot_Auto_ClipBoard_er

If InStr(Command_Screen_Shot_Auto_ClipBoard_er.Caption, "Archive Mode") = 0 Then
    Command_Screen_Shot_Auto_ClipBoard_er.Caption = "Screen Shot Auto ClipBoard_er when Spy_er && Archive Mode _OFF_ Hitt Button Here to Change"
    Exit Sub
End If

If InStr(Command_Screen_Shot_Auto_ClipBoard_er.Caption, "Archive Mode _ON") > 0 Then
    Command_Screen_Shot_Auto_ClipBoard_er.Caption = "Screen Shot Auto ClipBoard_er when Spy_er && Archive Mode _OFF_ Hitt Button Here to Change"
Else
    Command_Screen_Shot_Auto_ClipBoard_er.Caption = "Screen Shot Auto ClipBoard_er when Spy_er && Archive Mode _ON_ Hitt Button Here to Change"
End If

End Sub

Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim R As Long, Pic As PicBmp, IPic As IPicture, IID_IDispatch As GUID
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    With Pic
        .size = Len(Pic)
        .Type = vbPicTypeBitmap
        .hBmp = hBmp
        .hPal = hPal
    End With
    R = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
    Set CreateBitmapPicture = IPic
End Function

Function hDCToPicture(ByVal hDCSrc As Long, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long, R As Long
    Dim hPal As Long, hPalPrev As Long, RasterCapsScrn As Long, HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long, LogPal As LOGPALETTE

    hDCMemory = CreateCompatibleDC(hDCSrc)
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        R = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        R = RealizePalette(hDCMemory)
    End If
    R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    R = DeleteDC(hDCMemory)
    Set hDCToPicture = CreateBitmapPicture(hBmp, hPal)
End Function
Function PrintScreenOntoForm(ByVal Form As Form)
    Set Form.Picture = hDCToPicture(GetDC(0), 0, 0, Screen.width / Screen.TwipsPerPixelX, Screen.height / Screen.TwipsPerPixelY)
End Function

Function PrintCurrentFormOntoForm(ByVal Form As Form)
    Dim R As RECT
    Dim hWndx, LEFT_RIGHT_INSET___________, i1, i2
    hWndx = GetForegroundWindow
    
    GetWindowRect hWndx, R
    
    'WINDOW 10 OR ADJUST
    Dim i, LEFT_RIGHT_INSET_AND_OFFSET
    
    LEFT_RIGHT_INSET_AND_OFFSET = 1
    If IsIDE = False Then
        LEFT_RIGHT_INSET_AND_OFFSET = 8
        LEFT_RIGHT_INSET___________ = LEFT_RIGHT_INSET_AND_OFFSET * 2
    End If
    i1 = LEFT_RIGHT_INSET_AND_OFFSET
    i2 = LEFT_RIGHT_INSET___________
    
    Set Form.Picture = hDCToPicture(GetDC(0), R.Left + i1, R.Top, (R.Right - R.Left) - i2, R.Bottom - R.Top)
End Function



Function Print_hWnd_FORM_ontoForm(ByVal Form As Form, hWnd_NUMBER As Long)
    Dim R As RECT
    Dim hWndx
    
    hWndx = hWnd_NUMBER
    
    'hWndx = GetForegroundWindow
        
    GetWindowRect hWndx, R
    
    Set Form.Picture = hDCToPicture(GetDC(0), R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top)
End Function

Function Print_hWnd_FORM_ontoForm_2(ByVal Form As Form, hWnd_NUMBER As Long)
    'NOT OVERLAP
    Dim R As RECT
    Dim hWndx
    
    hWndx = hWnd_NUMBER
    
    'hWndx = GetForegroundWindow
        
    GetWindowRect hWndx, R
    
    Set Form.Picture = hDCToPicture(GetDC(0), R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top)

    'rv = SendMessage(Me.hWnd, WM_PRINT, SCREEN_CAP_PICTURE.Picture1.HDC, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
    'rv = SendMessage(Me.hWnd, WM_PRINT, SCREEN_CAP_PICTURE.Picture1.HDC, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
    
    'rv = SendMessage(hWnd_NUMBERING, WM_PAINT, SCREEN_CAP.HDC, 0)
    'rv = SendMessage(hWnd_NUMBERING, WM_PRINT, SCREEN_CAP.HDC, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)

End Function

Private Function Print_hWnd_FORM_ontoForm_EVEN_WHEN_BEHIND_ANOTHER(ByVal Form As Form, hWnd_NUMBER As Long)
Dim rv As Long
Dim hWnd_NUMBERING As Long

    hWnd_NUMBERING = Val(hWnd_NUMBER)
    
    'Clipboard.Clear
    
'    Picture1.Width = Me.Width
'    Picture1.Height = Me.Height
   
    'Picture1.AutoRedraw = True
    
    'rv = SendMessage(hWnd_NUMBERING, WM_PAINT, SCREEN_CAP.HDC, 0)
    'rv = SendMessage(hWnd_NUMBERING, WM_PRINT, SCREEN_CAP.HDC, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
    
'    rv = SendMessage(hWnd_NUMBERING, WM_PAINT, SCREEN_CAP_PICTURE.Picture1.HDC, 0)
'    rv = SendMessage(hWnd_NUMBERING, WM_PAINT, SCREEN_CAP_PICTURE.Picture1.HDC, 0)
    
    rv = SendMessage(Me.hWnd, WM_PRINT, SCREEN_CAP_PICTURE.Picture1.HDC, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
    rv = SendMessage(Me.hWnd, WM_PRINT, SCREEN_CAP_PICTURE.Picture1.HDC, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
    
    SCREEN_CAP_PICTURE.Show
    
    
    'Set SCREEN_CAP_PICTURE.Picture1.Picture = SCREEN_CAP_PICTURE.Picture1.Image
    'Set SCREEN_CAP.Picture1.Picture = SCREEN_CAP_PICTURE.Picture1.Image
    
    'Set Form.Picture = SCREEN_CAP_PICTURE.Picture1.Image
    
    'Clipboard.SetData Picture1.Picture, vbCFBitmap

End Function



Sub RUNAS_ME()
        
'----
'Batch Script to Run as Administrator - Stack Overflow
'http://stackoverflow.com/questions/14639743/batch-script-to-run-as-administrator
'----
'Answered Apr 24 '14 at 12:10
'Nika Gerson Lohman MALE
'set ELEVATE_APP=Full command line without parameters for the app to run
'set ELEVATE_PARMS=The actual parameters for the app
'echo Set objShell = CreateObject("Shell.Application") >elevatedapp.vbs
'echo Set objWshShell = WScript.CreateObject("WScript.Shell") >>elevatedapp.vbs
'echo Set objWshProcessEnv = objWshShell.Environment("PROCESS") >>elevatedapp.vbs
'echo objShell.ShellExecute "%ELEVATE_APP%", "%ELEVATE_PARMS%", "", "runas" >>elevatedapp.vbs
'DEL elevatedapp.vbs
'Set objWshShell = WScript.CreateObject("WScript.Shell")
        
        
'IN THE __.VBP FILE PROJECT
'Reference=*\G{F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}#1.0#0#C:\Windows\SysWOW64\wshom.ocx#Windows Script Host Object Model
Dim objShell, objWshShell, objWshProcessEnv
Dim ELEVATE_APP, ELEVATE_PARMS
ELEVATE_APP = App.Path + "\" + App.EXEName + ".EXE"
Set objShell = CreateObject("Shell.Application")
Set objWshShell = CreateObject("WScript.Shell")
Set objWshProcessEnv = objWshShell.Environment("PROCESS")

objShell.ShellExecute ELEVATE_APP, ELEVATE_PARMS, "", "runas"

Set objShell = Nothing
Set objWshShell = Nothing
Set objWshProcessEnv = Nothing

Exit Sub

'-----------------------------
'IN THE __.VBP FILE PROJECT
'Reference=*\G{565783C6-CB41-11D1-8B02-00600806D9B6}#1.2#0#C:\Windows\SysWOW64\wbem\wbemdisp.TLB#Microsoft WMI Scripting V1.2 Library
Dim oRunas
Set oRunas = CreateObject("runas.clsrunas", GetComputerName) 'SERVER NAME
With oRunas
        .sDomain = "WORKGROUP"
        .sUserName = GetUserName
        .sPassword = " "
        .sCommand = App.Path + "\" + App.EXEName + ".EXE"
        .RunAs 'Call the Run As method
End With
Set oRunas = Nothing

End Sub



'////////////////////////////////////////////////////////////////////
'//// PRIVATE FUNCTIONS
'////////////////////////////////////////////////////////////////////
' Get window styles
Private Function GetWindowStyle(ByVal LhWnd As Long) As String
    Dim lStyle As Long
        
    ' Get window styles
    lStyle = GetWindowLong(LhWnd, GWL_STYLE)
    
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

' Make textboxes flat
Private Sub MakeFlat(LhWnd As Long)
    Dim lStyle As Long
    
    ' Get window style
    lStyle = GetWindowLong(LhWnd, GWL_EXSTYLE)
    ' Setup window styles
    lStyle = lStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    ' Set window style
    SetWindowLong LhWnd, GWL_EXSTYLE, lStyle
    RemoveBorder LhWnd
End Sub
Private Sub RemoveBorder(LhWnd As Long)
    Dim lStyle As Long
    
    ' Get window style
    lStyle = GetWindowLong(LhWnd, GWL_STYLE)
    ' Setup window styles
    lStyle = lStyle And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION Or WS_BORDER Or WS_SIZEBOX Or WS_THICKFRAME)
    ' Set window style
    SetWindowLong LhWnd, GWL_STYLE, lStyle
    ' Update window
    SetWindowPos LhWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

Sub PROCESS_RELOADER_WATCHER_VAR_SET_CHECK_CONDICTION(VAR_IN, VAR_OUT)


    Dim FLAG_ER, FOR_NEXT_R
    Dim APP_NAME_RELOAD_IT_ER_EXE_2
    
    VAR_IN = APP_NAME_RELOAD_IT_ER_EXE_2
    VAR_OUT = True
    '-----------------------------------
    FLAG_ER = True
    For FOR_NEXT_R = 0 To VAR_ARRAY
        If BLOCK_RUN_1(FOR_NEXT_R) > Now Then FLAG_ER = False
    Next
    If FLAG_ER = True Then VAR_ARRAY = 0
    '-----------------------------------
    For FOR_NEXT_R = 0 To VAR_ARRAY
        If BLOCK_RUN_1(FOR_NEXT_R) > Now And BLOCK_RUN_2(FOR_NEXT_R) = APP_NAME_RELOAD_IT_ER_EXE_2 Then
            VAR_OUT = False
            Exit Sub
        End If
    Next
    For FOR_NEXT_R = 0 To VAR_ARRAY
        If BLOCK_RUN_3(FOR_NEXT_R) = True Then
            VAR_OUT = False
            Exit Sub
        End If
    Next

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

Private Sub Timer_DIR_FOR_FACEBOOK_VIDEO_Timer()
On Error Resume Next

Dim DIR_PATH

DIR_PATH = "C:\DOWNLOADS\# 00 VIDEO\# VIDEO FACEBOOK"
    
If Dir(DIR_PATH + "\" + Format(Now, "YYYY-MM-MMM"), vbDirectory) <> "" Then Exit Sub

MkDir DIR_PATH + "\" + UCase(Format(Now, "YYYY-MM-MMM"))


End Sub

Private Sub Timer_DIR_FOR_ARGUS_VIDEO_Timer()

On Error Resume Next

Dim DIR_PATH

If O_DAY_NOW_DIR_FOR_ARGUS_VIDEO = Day(Now) Then Exit Sub

O_DAY_NOW_DIR_FOR_ARGUS_VIDEO = Day(Now)

DIR_PATH = "C:\DOWNLOADS\# 00 VIDEO\# VIDEO ARGUS"
    
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
Call Timer_DIR_FOR_ARGUS_VIDEO_Timer
If O_DAY_NOW_DIR_FOR_VIDEO <> Day(Now) Then
    O_DAY_NOW_DIR_FOR_VIDEO = Day(Now)
    Call Timer_DIR_FOR_FACEBOOK_VIDEO_Timer
End If

Call TIMER_POLL_PATH_ARRAY_SET_NETWORK_ALL_SPEICAL_REQUEST_Timer

' Call Timer_RELOAD_AUTOHOTKEY_APP_Timer

If FindHandle_hWnd_COUNT_CHANGE = True Then
    Call EnumProcess
End If
If MDIProcServ.TT_1VDT > 0 Or MDIProcServ.TT_2VDT > 0 Then
    Dim XI_STRING
    XI_STRING = ""
    XI_STRING = XI_STRING + Str(DateDiff("h", MDIProcServ.TT_1VDT, Now)) + " hr -- "
    XI_STRING = XI_STRING + Format(Int((DateDiff("h", MDIProcServ.TT_1VDT, Now) / 24)), "0") + " d "
    XI_STRING = XI_STRING + Format((DateDiff("h", MDIProcServ.TT_1VDT, Now) Mod 24), "0") + " hr "
    Form1.Text_SYSTEM_START_TIME_02.Text = XI_STRING
    Form1.Text_SYSTEM_START_TIME_02.FontSize = 10
    Form1.Text_OS_INSTALL_DATE_02.Text = Str(DateDiff("d", MDIProcServ.TT_2VDT, Now)) + " DAY"
    Form1.Text_OS_INSTALL_DATE_02.FontSize = 10
Else
    Form1.Text_SYSTEM_START_TIME_02.Text = "SYSTEM START _ " + Str(TIMER_GO_COMPUTER_START)
End If



'For Each Form In Forms
'    If Form.EXIT_TRUE = True Then
'        EXIT_TRUE = True
'        Exit For
'    End If
'Next Form
'
'If EXIT_TRUE = True Then
'    Exit Sub
'End If

On Error Resume Next
Err.Clear
TIMER_GO_COMPUTER_START = TIMER_GO_COMPUTER_START - 1
If TIMER_GO_COMPUTER_START = 0 Then
    Load MDIProcServ
End If
If TIMER_GO_COMPUTER_START < -10 Then TIMER_GO_COMPUTER_START = -10

' ----------------------------------------
' GOT PROBLEM WHEN UNLOADER
' ----------------------------------------
' CAN'T USE THIS NOW
' THE LINE ABOVE NEAR ALWAYS GOT ERROR
' WITH
' OBJECT VARTIABLE OR WIDTH BLOCK NOT SET
' MOSTLY WHEN GET TO
' ----------------------------------------
' LOOK HERE
' ----------------------------------------
' MORE TO DO IF ERROR PERSIST
' ----------------------------------------
' Call ComputerSystem_GET_INFO
' Call OperatingSystem_GET_INFO
' Call Processor_GET_INFO
' ----------------------------------------

'ERR.DESCRIPTION

X_ONE_SECOND = X_ONE_SECOND + 1
' EVERY 4 SECOND
If X_ONE_SECOND Mod 4 = 0 Then
    X_ONE_SECOND = 0
    If InStr(UCase(GetWindowTitle(Me.hWnd)), "NOT RESPONDING") > 0 Then
        Call Label53_Click
        Call Label_CLOSE_GOODSYNC_Click
        Call KILL_WSCRIPT_GLOBAL
        If Me.WindowState <> vbNormal Then
            Me.WindowState = vbNormal
        End If
    End If
End If



If Err.Number > 0 Then
    ' Timer_1_SECOND.Enabled = False
    Exit Sub
End If


Call Label_GOODSYNC_01_Click

Call VB_EXE_SYNC

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

Dim TIMER_VAR

TIMER_VAR = 20000
If IsIDE = True Then TIMER_VAR = 60000
If Timer_DIR_FOR_C_DRIVE_ROOT_VICE_VERSA_VBSCRIPT.Interval <> TIMER_VAR Then
    Timer_DIR_FOR_C_DRIVE_ROOT_VICE_VERSA_VBSCRIPT.Interval = TIMER_VAR
    Exit Sub
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
    V_VAR = V_VAR + Trim(Str(F.size))
    Set F = Nothing
    
    If FSO.FileExists("C:\" + File2.List(R)) Then
        Set F = FSO.GetFile("C:\" + File2.List(R))
        V_VAR = V_VAR + File2.List(R)
        V_VAR = V_VAR + Format(F.DateLastModified, "YYYY-MM-DD HH-MM-SS")
        V_VAR = V_VAR + Trim(Str(F.size))
        Set F = Nothing
    End If

Next

V_VAR = V_VAR + Trim(Str(File3.ListCount))
For R = 0 To File3.ListCount - 1
    
    If InStr(UCase(File3.List(R)), ".SYS") = 0 Then
        Set F = FSO.GetFile(File3.Path + "\" + File3.List(R))
        V_VAR = V_VAR + File3.List(R)
        V_VAR = V_VAR + Format(F.DateLastModified, "YYYY-MM-DD HH-MM-SS")
        V_VAR = V_VAR + Trim(Str(F.size))
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


If Me.Top + Me.Left + Me.width + Me.height <> OLD_ME_POS Then
    Timer_Pause_Update.Interval = 4000
    Timer_Pause_Update.Enabled = True
    Label53.BackColor = Label59.BackColor
End If

OLD_ME_POS = Me.Top + Me.Left + Me.width + Me.height

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
    'hWnd_2 = GET_CHILD_TEST(MSDN_Lib)
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
' 2 ---- GOOD_SYNC ASK TO CLOSE AT END
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
If GetAsyncKeyState(27) Then 'ESCAPE KEY
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
'IhWnd = FindWindow(vbNullString, "wndclass_desked_gsk")
'I IhWnd > 0
'    I_CH_hWnd = FindWindowEx("VbaWindow", vbNullString)
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
If TEAM_VIEWER > 0 And TEAM_VIEWER <> OhWnd_TEAM_VIEWER Then

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
    'hWnd_2 = GET_CHILD_TEST(MSDN_Lib)
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

OhWnd_TEAM_VIEWER = TEAM_VIEWER







'---------------------------------------
'03 OF 03 VISIAUL BASIC LOAD CRASH THE CLIPBOARD
'---------------------------------------
'2017_FEB_22 EARLY HOUR 01:43a
'---------------------------------------
'TYPE 01 OF 02
Dim VB_LOADER
VB_LOADER = FindWindow("#32770", "Visual Component Manager")
If VB_LOADER > 0 And VB_LOADER <> OhWnd_VB_LOADER Then
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
If VB_LOADER > 0 And VB_LOADER <> OhWnd_VB_LOADER Then
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
If VB_LOADER > 0 And VB_LOADER <> OhWnd_VB_LOADER Then
    'OhWnd_VB_CLIPPER_ERROR = VB_CLIPPER_ERROR
    OhWnd_VB_LOADER = VB_LOADER
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
If VB_CLIPPER_ERROR > 0 And VB_CLIPPER_ERROR <> OhWnd_VB_CLIPPER_ERROR Then
    OhWnd_VB_CLIPPER_ERROR = VB_CLIPPER_ERROR
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
    If VB_EXE > 0 And VB_EXE <> OhWnd_VB_EXE Then
        Result = PostMessage(VB_EXE, WM_CLOSE, 0&, 0&)
    End If
End If


'TYPE 02 OF 02
VB_LOADER = FindWindow("#32770", "Add-In Toolbar")
If 1 = 2 And VB_LOADER > 0 And VB_LOADER <> OhWnd_VB_LOADER Then
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
    'hWnd_2 = GET_CHILD_TEST(MSDN_Lib)
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
Dim Text1
VB_LOADER = FindWindow("#32770", "Microsoft Visual Basic")
If 1 = 2 And VB_LOADER > 0 And VB_LOADER <> OhWnd_VB_LOADER Then
    PWnd = VB_LOADER
    GS_cWnd1 = FindWindowEx(PWnd, 0, vbNullString, "OK")
    GS_cWnd1 = FindWindowEx(PWnd, 0, "Static", vbNullString)
    Text1 = WindowText(GS_cWnd1)
    
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
    If FOREGROUND_WINDOW_CHANGE_DELAY_2 > Now Then
        FOREGROUND_WINDOW_CHANGE_DELAY_1_EXTRA_TO_DO.Enabled = True
        Exit Sub
    End If
    FOREGROUND_WINDOW_CHANGE_DELAY_2 = Now + TimeSerial(0, 0, 2)
    
    Call EnumProcess
End If


If InStr(MNU_WINMERGE_ON_TOP_ALLTME.Caption, "ALLTIME=YES") > 0 Then

    ihWnd = FindWindow("WinMergeWindowClassW", vbNullString)
    If ihWnd > 0 And O_IhWnd <> ihWnd Then
        
        SetWindowPos ihWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    End If
    O_IhWnd = ihWnd
End If


O_GetForegroundWindow = GetForegroundWindow

Dim R, A1, A2

For R = 2 To lstProcess_3_SORTER_ListView.ListItems.Count
    A1 = lstProcess_3_SORTER_ListView.ListItems.Item(R).SubItems(1)
    A2 = lstProcess_3_SORTER_ListView.ListItems.Item(R - 1).SubItems(1)
    If A1 = A2 Then
        ' IF TWO ARE RUNNER
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
    Label53.Caption = "Enumerate Event " + TIME_LAB + " _ Ticker"
    
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

' --------------------------------------------
' WHEN KILL SOME THING WANT ENUM PROCESS AFTER
' WILL = TRUE IN THAT INSTANCE
' --------------------------------------------
' ENUMPROCESS_NOT_RUN_YET = TRUE FALSE
' WILL = FALSE IF NEVER HAD A RUN YET
' --------------------------------------------
If ENUMPROCESS_MUST_RUNNER = False And ENUMPROCESS_NOT_RUN_YET = True Then
    If Timer_Pause_Update.Enabled = True Then Exit Sub
End If
    
ENUMPROCESS_NOT_RUN_YET = True
ENUMPROCESS_MUST_RUNNER = False

Dim PROCESS_PID_STORE_LST_01, PROCESS_PID_STORE_LST_02

If LISTVIEW_2_OR_3_HITT = 3 Then
    If lstProcess_3_SORTER_ListView.ListItems.Count > 0 Then
        If lstProcess_3_SORTER_ListView.SelectedItem.Index <> 0 Then
            PROCESS_PID_STORE_LST_01 = lstProcess_3_SORTER_ListView.ListItems(lstProcess_3_SORTER_ListView.SelectedItem.Index)
        End If
    End If
End If
If LISTVIEW_2_OR_3_HITT = 2 Then
    If lstProcess_2_ListView.ListItems.Count > 0 Then
        If lstProcess_2_ListView.SelectedItem.Index <> 0 Then
            PROCESS_PID_STORE_LST_02 = lstProcess_2_ListView.ListItems(lstProcess_2_ListView.SelectedItem.Index)
        End If
    End If
End If


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
'    frmMain.lstProcess_2_ListView.ListItems.Add , , lstProcess.List(R_I)
'    frmMain.lstProcess_3_SORTER_ListView.ListItems.Add , , ITEM_ADD_10 + " " + ITEM_ADD_22


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

    'ADDITEM
    'lstProcess_3_SORTER_ListView.ListItems.Add
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


Label46.Caption = "Process Counter" ' + Str(lstProcess.ListCount - 2)

If O_lstProcess_ListCount = -1 Then
    O_lstProcess_ListCount = lstProcess.ListCount
End If

If O_lstProcess_ListCount > lstProcess.ListCount Then TAG_VAR = " Less"
If O_lstProcess_ListCount < lstProcess.ListCount Then TAG_VAR = " Higher"
If O_lstProcess_ListCount = lstProcess.ListCount Then TAG_VAR = " Equal"

Label52.Top = Label52.Top
Label51.FontSize = 18

Label51.width = 400
If O_lstProcess_ListCount > lstProcess.ListCount Or 1 = 1 Then
    'DOWN
    TAG_VAR_2 = Chr(75)
    'Label51_Here
    On Error Resume Next
    Err.Clear
    Label51.FontName = "Wingdings 3"
    If Err.Number > 0 And ESCAPE_ERROR_WINGDINGS_3 = False Then
        ESCAPE_ERROR_WINGDINGS_3 = True
        TYPEMESSENGER = ""
        TYPEMESSENGER = TYPEMESSENGER + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + "FROM HERE ____"
        TYPEMESSENGER = TYPEMESSENGER + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + "Download Wingdings 2 Font - Free Font Download" + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + "http://www.fontpalace.com/font-download/Wingdings+2/" + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + "Download Wingdings 3 Font - Free Font Download" + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + "http://www.fontpalace.com/font-download/Wingdings+3/" + vbCrLf
        MsgBox "ERROR BUT CONTINUE ANYWAY _ YOU GOT TO INSTALL FONT WINGDINGS 3" + TYPEMESSENGER, vbMsgBoxSetForeground
    End If

    Label51.Left = Label52.Left + 80
    Label51.Caption = TAG_VAR_2
    Label51.Top = Label52.Top '- 20
    Label51.height = Label52.height - (Label52.height - Label51.height) ' Label51_Height
    'There is ANOTHER LABEL HIDE IN BACKGROUND TO BLEND IN THE SIZE IS SMALLER
    'ON THE FORM IN GREY
End If
If O_lstProcess_ListCount < lstProcess.ListCount Then
    'UP
    TAG_VAR_2 = Chr(74)
    On Error Resume Next
    Err.Clear
    Label51.FontName = "Wingdings 3"
    If Err.Number > 0 And ESCAPE_ERROR_WINGDINGS_3 = False Then
        ESCAPE_ERROR_WINGDINGS_3 = True
        TYPEMESSENGER = ""
        TYPEMESSENGER = TYPEMESSENGER + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + "FROM HERE ____"
        TYPEMESSENGER = TYPEMESSENGER + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + "Download Wingdings 2 Font - Free Font Download" + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + "http://www.fontpalace.com/font-download/Wingdings+2/" + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + "Download Wingdings 3 Font - Free Font Download" + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + "http://www.fontpalace.com/font-download/Wingdings+3/" + vbCrLf
        MsgBox "ERROR BUT CONTINUE ANYWAY _ YOU GOT TO INSTALL FONT WINGDINGS 3" + TYPEMESSENGER, vbMsgBoxSetForeground
    End If
    
    Label51.Left = Label52.Left + 80
    Label51.Caption = TAG_VAR_2
    Label51.Top = Label52.Top '- 20
    Label51.height = Label52.height - (Label52.height - Label51.height) ' Label51_Height
End If
If O_lstProcess_ListCount = lstProcess.ListCount Then
    'EQUAL
    TAG_VAR_2 = Chr(240)
    
    On Error Resume Next
    Err.Clear
    Label51.FontName = "Wingdings 2"
    If Err.Number > 0 And ESCAPE_ERROR_WINGDINGS_2 = False Then
        ESCAPE_ERROR_WINGDINGS_2 = True
        TYPEMESSENGER = ""
        TYPEMESSENGER = TYPEMESSENGER + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + "FROM HERE ____"
        TYPEMESSENGER = TYPEMESSENGER + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + "Download Wingdings 2 Font - Free Font Download" + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + "http://www.fontpalace.com/font-download/Wingdings+2/" + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + "Download Wingdings 3 Font - Free Font Download" + vbCrLf
        TYPEMESSENGER = TYPEMESSENGER + "http://www.fontpalace.com/font-download/Wingdings+3/" + vbCrLf
        MsgBox "ERROR BUT CONTINUE ANYWAY _ YOU GOT TO INSTALL FONT WINGDINGS 2" + TYPEMESSENGER, vbMsgBoxSetForeground
    End If
    
    Label51.Left = Label52.Left + 80
    Label51.Caption = TAG_VAR_2
    Label51.Top = Label52.Top + 20
    Label51.height = Label52.height - (Label52.height - Label51.height) ' Label51_Height

End If
'----
'Wingdings 3 character set and equivalent Unicode characters
'http://www.alanwood.net/demos/wingdings-3.html
'----

Label50.Caption = Str(lstProcess.ListCount - 2) + TAG_VAR ' + "_Timer 1 Min"

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

If GO_X <> 0 Then
    lstProcess_3_SORTER_ListView.SortOrder = lvwAscending
    lstProcess_3_SORTER_ListView.SortKey = 1
    lstProcess_3_SORTER_ListView.Sorted = True
    'lstProcess_3_SORTER_ListView.Sorted = False
End If

Call KILL_ON_MAXIMUM_PROCESS_LIMIT_COUNT_AUTOHOTKEY
Call KILL_ON_MAXIMUM_PROCESS_LIMIT_COUNT_CMD

Dim SET_DOWN
If PROCESS_PID_STORE_LST_01 <> 0 Then
    For R_I = 1 To lstProcess_3_SORTER_ListView.ListItems.Count - 1
        If lstProcess_3_SORTER_ListView.ListItems(R_I) = PROCESS_PID_STORE_LST_01 Then
            lstProcess_3_SORTER_ListView.ListItems.Item(lstProcess_3_SORTER_ListView.ListItems.Count - 1).EnsureVisible
            lstProcess_3_SORTER_ListView.ListItems.Item(R_I).EnsureVisible
            SET_DOWN = R_I - 4
            If SET_DOWN < 1 Then SET_DOWN = 1
            lstProcess_3_SORTER_ListView.ListItems.Item(SET_DOWN).EnsureVisible
            lstProcess_3_SORTER_ListView.ListItems.Item(R_I).EnsureVisible
            lstProcess_3_SORTER_ListView.ListItems.Item(R_I).Selected = True
            ' lstProcess_3_SORTER_ListView.SetFocus
            Exit For
        End If
    Next
End If
If PROCESS_PID_STORE_LST_02 <> 0 Then
    For R_I = 1 To lstProcess_2_ListView.ListItems.Count - 1
        If lstProcess_2_ListView.ListItems(R_I) = PROCESS_PID_STORE_LST_02 Then
            lstProcess_2_ListView.ListItems.Item(lstProcess_2_ListView.ListItems.Count - 1).EnsureVisible
            lstProcess_2_ListView.ListItems.Item(R_I).EnsureVisible
            SET_DOWN = R_I - 4
            If SET_DOWN < 1 Then SET_DOWN = 1
            lstProcess_2_ListView.ListItems.Item(SET_DOWN).EnsureVisible
            lstProcess_2_ListView.ListItems.Item(R_I).EnsureVisible
            lstProcess_2_ListView.ListItems.Item(R_I).Selected = True
            ' lstProcess_2_ListView.SetFocus
            Exit For
        End If
    Next
End If
    
End Sub

Private Sub Timer_Pause_Update_Timer()
Timer_Pause_Update.Enabled = False
'Label53.width
Label53.BackColor = Label58.BackColor

End Sub

Sub KILL_ON_MAXIMUM_PROCESS_LIMIT_COUNT_CMD()

Dim T_COUNTER, R_I, R

If MNU_CMD_KILL_MAX.Caption = "KILL CMD PID LIMITER > 20=FALSE" Then
    Exit Sub
End If

PROCESS_TO_KILLER = lstProcess_2_ListView.ListItems(1).SubItems(1)
T_COUNTER = 1
For R_I = 0 To lstProcess_2_ListView.ListItems.Count - 1
    If R_I > 0 Then
        PROCESS_TO_KILLER = lstProcess_2_ListView.ListItems(R_I).SubItems(1)
        If InStr(PROCESS_TO_KILLER, "cmd.exe") > 0 Then
            T_COUNTER = T_COUNTER + 1
        End If
        If T_COUNTER > 20 Then
            pid = Val(lstProcess_2_ListView.ListItems.Item(R_I))
            cProcesses.Process_Kill (pid)
        End If
    End If
Next

End Sub

Sub KILL_ON_MAXIMUM_PROCESS_LIMIT_COUNT_AUTOHOTKEY()

Dim T_COMPARE, T_COUNTER, R_I

If MNU_KILL_MAX_AHK.Caption = "KILL AHK PID LIMITER > 100=FALSE" Then
    Exit Sub
End If

Call FindHandle_AutoHotkey
' ----------------------------------------------------------
' THIS ROUTINE OUGHT TO SORT IT BETTER THAN AFTER NOW IS TWO
' DONE
' NOW WRITE IT IN AUTOHOTKEYS EASIER
' [ Saturday 17:45:00 Pm_02 March 2019 ]
' ----------------------------------------------------------

PROCESS_TO_KILLER = lstProcess_3_SORTER_ListView.ListItems(1).SubItems(1)
T_COMPARE = PROCESS_TO_KILLER
T_COUNTER = 1
For R_I = 0 To lstProcess_3_SORTER_ListView.ListItems.Count - 1
    If R_I > 0 Then
        PROCESS_TO_KILLER = lstProcess_3_SORTER_ListView.ListItems(R_I).SubItems(1)
        If PROCESS_TO_KILLER = T_COMPARE Then
        If InStr(PROCESS_TO_KILLER, "AutoHotkey.exe") > 0 Then
            T_COUNTER = T_COUNTER + 1
        End If
        End If
        If T_COUNTER > 100 Then
            
            Call MNU_TASK_KILLER_AUTOHOTKEYS_Click
            Exit Sub
        
        End If
        
        T_COMPARE = PROCESS_TO_KILLER
    End If
    
Next

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
    



End Sub


Public Function FindHandle_AutoHotkey() As Long

Dim test_hWnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String
Dim WText As String
Dim HighTest_hWnd As Long

Dim T_COMPARE, T_COUNTER
Dim RI
Dim FINDER_LINE, FINDER_LINE_1
Dim FINDER_LINE_hWnd As Long
Dim VAR As Long

Dim T_COUNTER_FLAG_VAR

Do
    List_SORT_FOR_AHK_LIMITER.Clear
    
    'Find the first window
    test_hWnd = FindWindowDLL(ByVal 0&, ByVal 0&)
    T_COMPARE = ""
    T_COUNTER = 1
    
    Do While test_hWnd <> 0
        
        WText = GetWindowTitle(test_hWnd)
    
        ' Autokey -- 58-Auto Repeat Browser Function Set.ahk
        ' ahk_class #32770
        '".ahk - AutoHotkey v"
        
    '    If InStr(wText, "Autokey -- 58-Auto Repeat Browser Function Set.ahk") > 0 Then
    '        Stop
    '    End If
        
        If InStr(WText, ".ahk") > 0 Then
            List_SORT_FOR_AHK_LIMITER.AddItem Trim(WText) + " ---------X " + Str(test_hWnd)
        End If
        
        'retrieve the next window
        test_hWnd = GetWindow(test_hWnd, GW_hWndNEXT)
    
    Loop
        
    ' ---------------------------------------------
    ' ONLY IF THIS ONE LOADEDER CHECK IF WNAT TO MAYBE
    ' BUT SOME AHK BRING MSGBOX UP ON OWN WHILE RUN
    ' THIS IS TO STOP FAULT OF THE AUTORELOADED
    ' LOADING ONE REPEAT WHEN HAS SYNTAX ERROR
    ' ---------------------------------------------
    ' Autokey -- 28-AUTOHOTKEYS SET RELOADER.ahk
    ' ---------------------------------------------
        
    T_COMPARE = ""
    T_COUNTER_FLAG_VAR = False
    For RI = 0 To List_SORT_FOR_AHK_LIMITER.ListCount - 1
        FINDER_LINE = List_SORT_FOR_AHK_LIMITER.List(RI)
        FINDER_LINE_1 = Mid(FINDER_LINE, 1, InStr(FINDER_LINE, " ---------X ") - 1)
        FINDER_LINE_hWnd = Val(Mid(FINDER_LINE, InStr(FINDER_LINE, " ---------X  ") + Len(" ---------X  ")))
        If FINDER_LINE_1 = T_COMPARE Then
            T_COUNTER = T_COUNTER + 1
            If T_COUNTER > 15 Then
                T_COUNTER_FLAG_VAR = True
                
                
                
                ' NOT WORK FOR WHAT WE WANTER
                ' closewindow (FINDER_LINE_hWnd)
                
                ' FINDER_LINE_hWnd_2 = GetParent(FINDER_LINE_hWnd)
                ' closewindow (FINDER_LINE_hWnd_2)
                
                If FINDER_LINE_hWnd > 0 Then
                    pid = -1
                    VAR = cProcesses.Convert(FINDER_LINE_hWnd, pid, cnFromhWnd Or cnToProcessID)
                End If
                If pid > 0 Then
                    Result = cProcesses.Process_Kill(pid)
                End If
            
            End If
        Else
            T_COUNTER = 1
        End If
        
        T_COMPARE = FINDER_LINE_1
        
    Next
    
Loop Until T_COUNTER_FLAG_VAR = False
    


End Function


Public Function FindHandle_hWnd_COUNT_CHANGE() As Long

Dim test_hWnd As Long
Dim T_COUNTER
Dim FindHandle_hWnd_COUNT
'Find the first window
test_hWnd = FindWindowDLL(ByVal 0&, ByVal 0&)
T_COUNTER = 1
Do While test_hWnd <> 0
    'retrieve the next window
    test_hWnd = GetWindow(test_hWnd, GW_hWndNEXT)
    T_COUNTER = T_COUNTER + 1
Loop

FindHandle_hWnd_COUNT = T_COUNTER
If FindHandle_hWnd_COUNT <> OLD_FindHandle_hWnd_COUNT Then
    FindHandle_hWnd_COUNT_CHANGE = True
    Label42.Caption = "hWnd Count:" + Str(FindHandle_hWnd_COUNT)
End If
OLD_FindHandle_hWnd_COUNT = FindHandle_hWnd_COUNT

End Function





Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column As ColumnHeader = Nothing)
    Dim c As ColumnHeader
    If Column Is Nothing Then
    For Each c In LV.ColumnHeaders
        SendMessage LV.hWnd, LVM_FIRST + 30, c.Index - 1, -1
    Next
    Else
        SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
    End If
    LV.Refresh
End Sub



Private Function WindowText(ByVal WINDOW_hWnd As Long) As String

    Dim txtlen              As Long

    If WINDOW_hWnd = 0 Then
        Exit Function
    End If

    txtlen = SendMessage(WINDOW_hWnd, WM_GETTEXTLENGTH, ByVal 0, ByVal 0)
    If txtlen = 0 Then
        Exit Function
    End If

    WindowText = String$(txtlen, vbNullChar)
    SendMessage WINDOW_hWnd, WM_GETTEXT, ByVal (txtlen + 1), ByVal StrPtr(WindowText)

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
   Dim l As Long
   Dim s As String
   l = GetWindowTextLength(hWnd)
   s = Space(l + 1)
   GetWindowText hWnd, s, l + 1
   GetWindowTitle = Left$(s, l)
End Function
Function GetWindowClass(ByVal hWnd As Long) As String
    Dim ret As Long, sText As String
    sText = Space(255)
    ret = GetClassName(hWnd, sText, 255)
    sText = Left$(sText, ret)
   GetWindowClass = sText
End Function



Private Sub FindCursor(x, y)

Dim P As POINTAPI

GetCursorPos P
'   return x and y co-ordinate
x = P.x ' / GetSystemMetrics(0) * Screen.Width
'   for current cursor position
y = P.y '/ GetSystemMetrics(1) * Screen.Height

End Sub


Function FindWindowPart(Test) As Long

FindWindowPart = False

Dim test_hWnd As Long

test_hWnd = FindWindowDLL(ByVal 0&, ByVal 0&)
Do While test_hWnd <> 0
        
        If InStr(GetWindowTitle(test_hWnd), Test) > 0 Then
            FindWindowPart = test_hWnd
            Exit Do
        End If
        
    'retrieve the next window
    test_hWnd = GetWindow(test_hWnd, GW_hWndNEXT)

Loop


End Function


' --------------------------------------------------------------------
' ALL IN ONE FUNCTION LESS API LESS DEMAN JOB WORKER
' --------------------------------------------------------------------
Private Function CreateFolderTree(ByVal sPath As String) As Boolean
    Dim nPos As Integer

    If Mid(sPath, Len(sPath), 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)

    On Error GoTo CreateFolderTreeError
    
    nPos = InStr(sPath, "\")
    While nPos > 0
        If nPos - 1 > 3 Then
            If Dir(Left$(sPath, nPos - 1), vbDirectory) = "" Then
                MkDir Left$(sPath, nPos - 1)
            End If
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
    If Dir(sPath, vbDirectory) = "" Then MkDir sPath
    
    CreateFolderTree = True
    If Dir(sPath, vbDirectory) = "" Then
        CreateFolderTree = False
    End If

    Exit Function

CreateFolderTreeError:
    Exit Function
End Function


' --------------------------------------------------------------------
' WORK BUT EASIER ALL IN ONE FUNCTION NOT WITH API THOUGH
' MORE TRANSPORTBALE TO OTHER CODE MAYBE IN ONE FUNCTION LESS DELCLARE
' --------------------------------------------------------------------
Private Function CreateFolderTree_API(ByVal sPath As String) As Boolean
    Dim nPos As Integer

    If Mid(sPath, Len(sPath), 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)

    On Error GoTo CreateFolderTreeError
    
    nPos = InStr(sPath, "\")
    While nPos > 0
        If nPos - 1 > 3 Then ' NOT CHECK ROOT
            If Not FolderExists(Left$(sPath, nPos - 1)) Then
                MkDir Left$(sPath, nPos - 1)
            End If
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
    If Not FolderExists(sPath) Then MkDir sPath
    
    If Not FolderExists(sPath) Then
        CreateFolderTree_API = False
        Exit Function
    End If
    CreateFolderTree_API = True
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
'Dim FilehWnd2
'
'Handle2 = FindWindow(vbNullString, "Deleting...")
'Handle3 = FindWindow(vbNullString, "ImageDupe registered to EPSiLON")
'
''GetWindowState
'FilehWnd2 = ""
'If Handle2 > 0 Then FilehWnd2 = GetFileFromhWnd(Handle2)
'
'If FilehWnd2 = "C:\Program Files\ImageDupe\ImageDupe.exe" And _
'    Check1.Value = vbChecked Then
''If handle2 > 0 Then
''If DupeCool = 0 Then
''handle3 = FindWindow(vbNullString, "ImageDupe registered to EPSiLON")
''MsgBox "Image Dupe Ready": DupeCool = 1
'Dim Rect5 As RECT
'Dim Rect7 As RECT
'hWnd9 = GetWindowRect(Handle2, Rect5)
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
'FilehWnd2 = ""
'If Handle2 > 0 Then FilehWnd2 = GetFileFromhWnd(Handle2)
'
'If FilehWnd2 = "C:\Program Files\ImageDupe\ImageDupe.exe" And _
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
'FilehWnd2 = ""
'If Handle2 > 0 Then FilehWnd2 = GetFileFromhWnd(Handle2)
''CM$ = GetFileFromProc(CMediaSource)
'
'If FilehWnd2 = "C:\Program Files\Documents To Go\DocsToGo.exe" Then
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
'If CurProchWnd = Me.hWnd And CurProchWnd <> Rf And TiTBait <> CurProchWnd And Rf > 0 Then
'
'    'Rf = FindWindow(vbNullString, "WMI Demo - CPU Information")
'
''    On Local Error Resume Next
''    AppActivate "WMI Demo - CPU Information"
''    AppActivate "CID Run Me"
''    On Local Error GoTo 0
'
'    'Rft = Putfocus(Rf)
'    'Rft = Putfocus(Me.hWnd)
'
'End If
'
'
'
'
''FOR OLD CALLER ID PROG BREAK AND PUT RUN -------------
'TiTBait = CurProchWnd
'
''hWnd3 = FindWindowPart("Microsoft Visual Basic [break]")
'hWnd3 = FindWindow(vbNullString, "CallerID - Microsoft Visual Basic [break]")
'If hWnd3 > 0 Then
'On Local Error Resume Next
'If Mnu_CIDBreak.Checked = True Then
'AppActivate "CallerID - Microsoft Visual Basic [break]", False
''r = Int(Now) + Timer + 0.1
''Do
''DoEvents
'CurProchWnd = GetForegroundWindow
''If CurProchWnd = hWnd3 Then Exit Do
''Loop Until r < Int(Now) + Timer
'
'
'If CurProchWnd = hWnd3 Then SendKeys "{f5}", False
'End If
'End If
'
'
'
'
'
''cphWnd = GetForegroundWindow
''If GetWindow("IEFrame", vbNullString) > 0 Then
''rrt$ = GetWindowTitle(CPhWnd)
'
''            Dim cText As String
'
''            cText = Space$(255)
'            'Ret = GetClassName(Peet2, sText, 255)
'
'
''TT = GetClassName(cphWnd, cText, 255)
''ctext = Mid$(ctext, 1, InStr(ctext, Chr(0)) - 1)
''If Mid$(cText, 1, 7) = "IEFrame" Then
''Xxr = cphWnd
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
'    'if all the hWnds in our list are present on screen then cool
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
'    'If lol = 1 Then rt = AlwaysOnTop(Me.hWnd)
'    'If lol = 0 Then rt = NotAlwaysOnTop(Me.hWnd)
'    'me.top
'    'me.left
'    Xxr16 = Xxr
'    'win2 = FindWindow("Winamp v1.x", vbNullString)
'    'win3 = FindWindow("Winamp PE", vbNullString)
'    'rt = AlwaysOnTop(win2)
'    'rt = AlwaysOnTop(win3)
'    'ShowWindow Me.hWnd, SW_NORMAL
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
'    If InStr(GetFileFromhWnd(Xxr), "VB6.EXE") = 0 Then
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
'    'If InStr(GetFileFromhWnd(Xxr), "WaveEdit.exe") = 0 Then
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
'    'hWndCTLTask2 = FindWindow(vbNullString, "CTLTask")
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
'            CloseProgramhWnd (Xxr)
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
'            'CloseProgramhWnd (Xxr)
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
'            CloseProgramhWnd (Xxr)
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
'            CloseProgramhWnd (Xxr)
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
''If CurProchWnd <> CurProchWnd2 Or ReRun = 1 Or FirstRun2 = False Then
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
'    'CloseProgramhWnd (Xxr)
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
'            'CloseProgramhWnd (Xxr)
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
'hWnd3 = FindWindow("IEFrame", vbNullString)
'If hWnd3 > 0 Then
'ash$ = GetWindowTitle(hWnd3)
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
''file2$ = GetFileFromhWnd(Xxr)
''    If InStr(file2$, "VPN_Auto-Dialer") Or InStr(file2$, "VB6.EXE") Or InStr(file2$, "SendEmail") Then
'    If 1 = 1 Then
'        test_hWnd = FindWindowDLL(0&, 0&)
'        hWndChild2 = FindWindowEx(Xxr, 0&, "Button", "&Do Not Send")
'        If hWndChild2 > 0 Then Xxr = hWndChild2: easyx = 1
'        If hWndChild2 = 0 Then
'            Do
'                hWndChild2 = FindWindowEx(test_hWnd, 0&, "Button", "&Do Not Send")
'                easyx = 0
'                If hWndChild2 > 0 Then
'                    easyx = 1
'                    Xxr = test_hWnd
'                End If
'            test_hWnd = GetWindow(test_hWnd, GW_hWndNEXT)
'            Loop Until test_hWnd = 0 Or easyx = 1
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
'Dim hWndMicE As Long
'hWndMicE = FindWindow(vbNullString, "Microsoft Excel")
'
'wert = GetWindowRect(hWndMicE, MeRyu3)
'If MeRyu3.Top < 40 Then
'
''MoveWindow hWndMicE, MeRyu3.Left, 131, MeRyu3.Right - MeRyu3.Left, MeRyu3.Bottom - MeRyu3.Top, True
''only if its not already maxmized correct this
'
'End If
'End If
'
'
'
'
''Process explorer Properties
'CurProchWnd = GetForegroundWindow
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



Public Function GetWindowState(ByVal lnghWnd As Long) As Integer
    If IsWindow(lnghWnd) = 1 Then
        GetWindowState = -1
    If IsIconic(lnghWnd) <> 0 Then
        GetWindowState = vbMinimized
    ElseIf IsZoomed(lnghWnd) <> 0 Then
        GetWindowState = vbMaximized
    End If
    End If
End Function



Public Function GetFileFromhWnd(lnghWnd) As String

'MsgBox getfilefromhWnd(Me.hWnd)

Dim lngProcess&, hProcess&, bla&, c&
Dim strFile As String
Dim x

strFile = String$(256, 0)
x = GetWindowThreadProcessId(lnghWnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
x = EnumProcessModules(hProcess, bla, 4&, c)
c = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
GetFileFromhWnd = Left(strFile, c)

End Function


Public Function GetFileFromProc(lngProcess) As String

'MsgBox getfilefromhWnd(Me.hWnd)
'Dim lngProcess&, hProcess&, bla&, C&
Dim hProcess&, bla&, c&
Dim strFile As String
Dim x

strFile = String$(256, 0)
'x = GetWindowThreadProcessId(lnghWnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
x = EnumProcessModules(hProcess, bla, 4&, c)
c = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
GetFileFromProc = Left(strFile, c)

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

Public Sub CloseProgramhWnd(ByVal Handle As Long)

 If Handle = 0 Then Exit Sub
 SendMessage Handle, WM_CLOSE, 0&, 0&

End Sub




Private Sub chkOnTop_Click()
    
    If chkOnTop.Value = 1 Then
        ' Put window on top of all others
        
        SetWindowPos Me.hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        SaveSetting "EliteSpy+", "Settings", "AlwaysOnTop", "1"
        MNU_ME_ON_TOP.Caption = "[__ ME ON TOP = YES __]"
        Label60.BackColor = Label49.BackColor '49 58_59

        'FEEDBACK LOOP STACK OVERFLOW IF SETTTING HERE
        '---------------------------------------------
        'Me.chkOnTop.Value = 1
        'MNU_ME_ON_TOP_Click

    Else
        ' Remove window from top
        SetWindowPos Me.hWnd, hWnd_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        SaveSetting "EliteSpy+", "Settings", "AlwaysOnTop", "0"
        MNU_ME_ON_TOP.Caption = "[__ ME ON tOP = NOT __]"
        'Label60.BackColor = Label59.BackColor '49 58_59
        Label60.BackColor = Label49.BackColor '49 58_59
        
        'Me.chkOnTop.Value = 0
    End If
    
    Timer_ALWAYS_ON_TOP_TO_START_WITH_ER.Enabled = False
    Label60.Caption = "Me on Top 20 Sec NOT"
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
    lhWnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub

Private Sub cmdMoveMax_SIMPLE(txtMhWnd_SIMPLE)
    
    ' Make Sure there is a space between the Eye _ i _ I _ iResult I_Result I Borg Kim
    ' --------------------------------------------------------------------------------

    If txtMhWnd_SIMPLE = "" Then txtMhWnd_SIMPLE = Me.hWnd

    Dim Rect_Get As RECT
    Dim HX, HY, HW, HH, SSHW, SSHH
    Dim i_Result
    Dim ScreenSize As RECT
    
    i_Result = GetWindowRect(GetDesktopWindow(), ScreenSize)
    
    i_Result = GetWindowRect(txtMhWnd_SIMPLE, Rect_Get)
    
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
    
    ' ShowWindow txtMhWnd_SIMPLE, SW_NORMAL
    
    MoveWindow txtMhWnd_SIMPLE, HX, HY, HW, HH, True
    
    ' lhWnd_Function_Button_Set_MIN_MAX = Val(txtMhWnd_SIMPLE)
    ' Call ChunkCodeOnMouse

End Sub

Private Sub cmdMaximize_Click()
    ' Maximize window
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd
    
    ShowWindow txtMhWnd.Text, SW_MAXIMIZE

    lhWnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
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

    lhWnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdNormal_Click()
    ' Show window
    Beep
    
    Dim lhWndParentX
    
    If Val(txtMhWnd.Text) = 0 Then txtMhWnd.Text = Me.hWnd
   
    ' txtMhWnd.Text = GetParent(Val(txtMhWnd.Text))
    ' txtMhWnd.Text = GetParenthWnd(Val(txtMhWnd.Text))
    ' lhWndParentX = GetParenthWnd(Val(txtMhWnd.Text))
    ' txtMhWnd.Text = GetParenthWnd(Val(txtMhWnd.Text))

    ' txtMhWnd.Text = GetAncestor(Val(txtMhWnd.Text), GA_ROOT)

    ShowWindow txtMhWnd.Text, SW_NORMAL
    
    lhWnd_Function_Button_Set_MIN_MAX = Val(txtMhWnd.Text)
    Call ChunkCodeOnMouse
    
End Sub

Private Sub cmdFlash_Click()
    ' Flash window
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd
    
    FlashWindow txtMhWnd.Text, 3
    
    lhWnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdEnable_Click()
    ' Enable window
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd
    
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd
    EnableWindow txtMhWnd.Text, 1

    lhWnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdDisable_Click()
    ' Disable window
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd
    
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd: Exit Sub
    EnableWindow txtMhWnd.Text, 0

    lhWnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
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
    
    lhWnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
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
    SetWindowPos txtMhWnd.Text, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

    lhWnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdNotOnTop_Click()
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd

    ' Remove window from top
    SetWindowPos txtMhWnd.Text, hWnd_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

    lhWnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
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

    lhWnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdShow_Click()
    ' Show window
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd
    ShowWindow txtMhWnd.Text, SW_SHOW

    lhWnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdTerminate_Click()
    ' Close window
    Beep
    If txtMhWnd.Text = "" Then txtMhWnd.Text = Me.hWnd: Exit Sub
    SendMessage txtMhWnd.Text, WM_CLOSE, 0, 0
End Sub


'Sub ChunkCodeOnMouse()
'
'        Dim tRC2 As RECT
'        Dim tPA As POINTAPI
'        Dim lhWnd As Long
'        Dim sTitle As String * 255
'        Dim sClass As String * 255
'        Dim sParentTitle As String * 255
'        Dim sParentClass As String * 255
'        Dim sParentTitleX As String * 255
'        Dim sParentClassX As String * 255
'        Dim lhWndParent As Long
'        Dim lhWndParentX As Long
'        Dim sStyle As String
'        Dim lRetVal As Long
'        Dim O_lhWndParent As Long
'        Dim Set_Collect_More_Info
'        Dim PID_MARK
'        Dim Success_Result
'
'        ' Get window rectCLEAR
'
'        If From_ListView = False Then
'            If hWnd_From_ListView = 0 And lhWnd_Function_Button_Set_MIN_MAX = 0 Then
'                ' Get cursor position
'                GetCursorPos tPA
'                ' Get window handle from point
'                lhWnd = WindowFromPoint(tPA.X, tPA.Y)
'                ' Get window caption
'            End If
'        End If
'
'        ' ---------------------------------------------------
'        ' USED BY
'        ' WHEN CLICK A PROCESS IN BOX THEN POLULATE THIS AREA
'        ' RATHER THAN HOVER OVER WITH MOUSE
'        ' Private Sub lstProcess_3_SORTER_ListView_Click()
'        ' ---------------------------------------------------
'        If From_ListView = True Then
'            Set_Collect_More_Info = True
'            If hWnd_From_ListView > 0 Then
'                lhWnd = hWnd_From_ListView
'                TxtPID.Text = Val(PROCESS_TO_KILLER_PID)
'                hWnd_From_ListView = 0
'                Set_Collect_More_Info = False
'            End If
'        End If
'        If From_ListView = True And hWnd_From_ListView = 0 Then
'            Set_Collect_More_Info = False
'        End If
'
'
'        If lhWnd_Function_Button_Set_MIN_MAX > 0 Then
'            ' txthWnd.Text = lhWnd
'            lhWnd = lhWnd_Function_Button_Set_MIN_MAX
'            lhWnd_Function_Button_Set_MIN_MAX = 0
'            ' Set_Collect_More_Info = false
'        End If
'
'
'
'
'        GetWindowRect lhWnd, tRC2
'
'        lRetVal = GetWindowText(lhWnd, sTitle, 255)
'        ' Get window class name
'        lRetVal = GetClassName(lhWnd, sClass, 255)
'        ' Get window style
'        sStyle = GetWindowStyle(lhWnd)
'        ' Get window parent
'
'        O_lhWndParent = lhWnd
'        lhWndParent = GetParent(lhWnd)
'        If lhWndParent = 0 Then lhWndParent = O_lhWndParent
'        lhWndParentX = GetParenthWnd(lhWnd)
'
'
'        If Set_Collect_More_Info = True Then
'            Success_Result = cProcesses.Get_PID_From_hWnd(lhWndParentX, PID_MARK)
'
'            TxtPID.Text = PID_MARK
'        End If
'
'
'
'
'        ' Get parent window caption
'        lRetVal = GetWindowText(lhWndParent, sParentTitle, 255)
'        ' Get parent window class name
'        lRetVal = GetClassName(lhWndParent, sParentClass, 255)
'
'        ' Get parentX window caption
'        lRetVal = GetWindowText(lhWndParentX, sParentTitleX, 255)
'        ' Get parentX window class name
'        lRetVal = GetClassName(lhWndParentX, sParentClassX, 255)
'
'        ' Set values to textboxes
'        If From_ListView = False Then
'
'            TxtEXE.Text = GetFileFromhWnd(lhWnd)
'
'
'            PROCESS_TO_KILLER = Mid(TxtEXE.Text, InStrRev(TxtEXE.Text, "\") + 1)
'            PROCESS_TO_KILLER_PID = TxtPID.Text
'
'            Call MOUSE_HOOVER_SLECTION_CLICKER
'
'        End If
'
'        From_ListView = False
'
'
'        txthWnd.Text = lhWnd
'        txthWndHX.Text = Hex(lhWnd)
'        txtTitle.Text = sTitle
'        txtClass.Text = sClass
'        txtStyle.Text = sStyle
'        txtRect.Text = "(" & tRC2.Left & ", " & tRC2.Top & ") - (" & tRC2.Right & ", " & tRC2.Bottom & ")"
'
'        'lhWndParent = GetParent(lhWndParent)
'        'lhWndParentHX = Hex(GetParent(lhWndParent))
'        If lhWnd <> lhWndParent Then
'            txtParent.Text = lhWndParent
'            txtParentHX.Text = Hex(lhWndParent)
'            txtParentText.Text = sParentTitle
'            txtParentClass.Text = sParentClass
'        Else
'            txtParent.Text = lhWndParent
'            txtParentHX.Text = Hex(lhWndParent)
'            txtParentText.Text = ""
'            txtParentClass.Text = ""
'        End If
'
'        If lhWnd <> lhWndParentX And lhWndParentX <> lhWndParent Then
'            txtParentX.Text = lhWndParentX
'            txtParentTextX.Text = sParentTitleX
'            txtParentClassX.Text = sParentClassX
'        Else
'            txtParentX.Text = lhWndParentX
'            txtParentTextX.Text = ""
'            txtParentClassX.Text = ""
'        End If
'
'
'        txtMhWnd.Text = lhWnd
'
'        If TxtEXE.Text <> OLD_TxtEXE_Text Then
'            Call TxtEXE_CLICK
'        End If
'        OLD_TxtEXE_Text = TxtEXE.Text
'
'
'End Sub
'
'////////////////////////////////////////////////////////////////////
'//// PRIVATE FUNCTIONS
'////////////////////////////////////////////////////////////////////
' Get window styles
'Private Function GetWindowStyle(ByVal lhWnd As Long) As String
'    Dim lStyle As Long
'
'    ' Get window styles
'    lStyle = GetWindowLong(lhWnd, GWL_STYLE)
'
'    ' Get window styles
'    If lStyle And WS_BORDER Then GetWindowStyle = GetWindowStyle & "WS_BORDER "
'    If lStyle And WS_CAPTION Then GetWindowStyle = GetWindowStyle & "WS_CAPTION "
'    If lStyle And WS_CHILD Then GetWindowStyle = GetWindowStyle & "WS_CHILD "
'    If lStyle And WS_CLIPCHILDREN Then GetWindowStyle = GetWindowStyle & "WS_CLIPCHILDREN "
'    If lStyle And WS_CLIPSIBLINGS Then GetWindowStyle = GetWindowStyle & "WS_CLIPSIBLINGS "
'    If lStyle And WS_DLGFRAME Then GetWindowStyle = GetWindowStyle & "WS_DLGFRAME "
'    If lStyle And WS_GROUP Then GetWindowStyle = GetWindowStyle & "WS_GROUP "
'    If lStyle And WS_HSCROLL Then GetWindowStyle = GetWindowStyle & "WS_HSCROLL "
'    If lStyle And WS_MAXIMIZEBOX Then GetWindowStyle = GetWindowStyle & "WS_MAXIMIZEBOX "
'    If lStyle And WS_MINIMIZEBOX Then GetWindowStyle = GetWindowStyle & "WS_MINIMIZEBOX "
'    If lStyle And WS_SYSMENU Then GetWindowStyle = GetWindowStyle & "WS_SYSMENU "
'    If lStyle And WS_POPUPWINDOW Then GetWindowStyle = GetWindowStyle & "WS_POPUPWINDOW "
'    If lStyle And WS_TABSTOP Then GetWindowStyle = GetWindowStyle & "WS_TABSTOP "
'    If lStyle And WS_THICKFRAME Then GetWindowStyle = GetWindowStyle & "WS_THICKFRAME "
'    If lStyle And WS_VISIBLE Then GetWindowStyle = GetWindowStyle & "WS_VISIBLE "
'    If lStyle And WS_VSCROLL Then GetWindowStyle = GetWindowStyle & "WS_VSCROLL "
'
'End Function

Function GetParenthWnd(ByVal ReturnParent As Long) As String
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
    GetParenthWnd = i
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
'lstProcess_2_ListView.left
'lstProcess_2_ListView.Height
'lstProcess_2_ListView.top
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
LISTVIEW_2_OR_3_HITT = 2

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
LISTVIEW_2_OR_3_HITT = 3

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
Dim hWnd_VAR
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
    hWnd_VAR = cProcesses.GethWnd_From_PID(PID_INPUT)
    If hWnd_VAR > 0 Then
        Success_Result = cProcesses.Get_PID_From_hWnd(hWnd_VAR, PID_MARK)
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
    
    'Mod1_EnumChildProc.EnumChildWindow_Routine (hWndForm)
    
End If

Dim hWnd_Parent As Long
PID_INPUT = Val(PROCESS_TO_KILLER_PID)

hWnd_Parent = cProcesses.GethWnd_From_PID_VISIBLE(PID_INPUT)
If hWnd_Parent = 0 Then
    ' hWnd_Parent = cProcesses.GethWnd_From_PID(PID_INPUT)
End If
Dim IRESULT
If hWnd_Parent = 0 Then
    ' IRESULT = cProcesses.Convert(PID_INPUT, hWnd_Parent, cnFromProcessID Or cnTohWnd)
End If
If hWnd_Parent = 0 Then
      ' hWnd_Parent = cProcesses.GethWnd_From_PID_123(PID_INPUT)
      ' hWnd_Parent = GethWndFromProcess(PID_INPUT)
      
      ' ------------------------------------------------------------------------------
      ' PLAYING AORUND AND CHROME NOT GIVE hWnd FOR PROCESS
      ' EXAMPLE GOT TWO WINDOW OPEN
      ' BOTH GOT A TITLE
      ' OUGHT TO BE EASY FIND BOTH
      ' BUT CHROME ONLY GIVE FOR ONE IN FOCUS
      ' USE HERE AND YOUR SEE
      ' Process Explorer - sysinternals: www.sysinternals.com [7-ASUS-GL522VW\MATT 04]
      ' ahk_class PROCEXPL
      ' ahk_exe procexp64.EXE
      ' NEAR MESSED UP ALL MY CODE
      ' ------------------------------------------------------------------------------
      ' CAN'T GET THIS CHILD UNTIL GOT hWnd FROM PID
      ' Mod1_EnumChildProc.EnumChildWindow_Routine (hWndForm)
      ' ------------------------------------------------------------------------------

End If

' txthWnd.Text = hWnd_Parent
If hWnd_Parent > 0 Then
    hWnd_From_ListView = hWnd_Parent
End If

Call Label29_Click

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
If PROCESS_TO_KILLER <> "" Then
Mid(PROCESS_TO_KILLER, 1, 1) = UCase(Mid(PROCESS_TO_KILLER, 1, 1))
End If

Label22.Caption = "TASKKILLER BY PID NUMBER __ # " + PROCESS_TO_KILLER_PID
Label30.Caption = "TASKKILLER NAME ___________ " + PROCESS_TO_KILLER

'TxtEXE.Text = GetFileFromProc(Val(PROCESS_TO_KILLER_PID))
'TxtPID.Text = Val(PROCESS_TO_KILLER_PID)

' Dim hWnd_Parent As Long
' Dim PID_INPUT As Long
' PID_INPUT = Val(PROCESS_TO_KILLER_PID)

' hWnd_Parent = cProcesses.GethWnd_From_PID(PID_INPUT)
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

Dim tPA As POINTAPI, LhWnd As Long, O_lhWndParent, lhWndParent, lhWndParentX
GetCursorPos tPA
LhWnd = WindowFromPoint(tPA.x, tPA.y)
O_lhWndParent = LhWnd
lhWndParent = GetParent(LhWnd)
If lhWndParent = 0 Then lhWndParent = O_lhWndParent
lhWndParentX = GetParenthWnd(LhWnd)

If GetAsyncKeyState(27) < 0 Then
    If GetForegroundWindow = Me.hWnd Or lhWndParent = Me.hWnd Or lhWndParentX = Me.hWnd Then
        If IsIDE = True Then
            Unload Me
        Else
            Me.WindowState = vbMinimized
        End If
    End If
End If

If IsIDE = True Then
    Dim i, KEY_V, SETG
    For i = 1 To 255
        KEY_V = GetAsyncKeyState(i)
        SETG = 1
        If i = 116 Then SETG = 0       'F5 TEST ISIDE
        If i = 16 Then SETG = 0        'LEFT    SHIFT KEY
        If i = 17 Then SETG = 0        'CONTROL SHIFT KEY
        If i = 162 Then SETG = 0       'CONTROL SHIFT KEY LEFT
        If i = 1 Then SETG = 0         'LEFT MOUSE CLICK
'        If KEY_V < 0 And SETG = 1 Then MsgBox Str(KEY_V) + " -- " + Str(i)
        ' If KEY_V = 67 Then           'C
    Next
End If

' CONTROL LEFT AND F1
If GetAsyncKeyState(162) < 0 Then
SET_GO_CONTROL_LEFT_F1 = True
End If
If GetAsyncKeyState(112) < 0 And SET_GO_CONTROL_LEFT_F1 = True Then
    If Me.WindowState = vbMinimized Then
        ' Me.WindowState = vbNormal ' vbMaximized
        SET_GO_CONTROL_LEFT_F1 = False
    End If
End If


'C
If GetAsyncKeyState(67) < 0 Then
    If GetForegroundWindow = Me.hWnd Or lhWndParent = Me.hWnd Or lhWndParentX = Me.hWnd Then
        Call Label_KILL_CMD_Click
    End If
End If

'A
If GetAsyncKeyState(65) < 0 Then
    If GetForegroundWindow = Me.hWnd Or lhWndParent = Me.hWnd Or lhWndParentX = Me.hWnd Then
        Call Label_KILL_AUTOHOTKEY_Click
    End If
End If

'K
If GetAsyncKeyState(75) < 0 Then
    If GetForegroundWindow = Me.hWnd Or lhWndParent = Me.hWnd Or lhWndParentX = Me.hWnd Then
        Call Label_KILL_CMD_AND_AHK_Click
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
'    'SetWindowPos txtMhWnd.Text, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'    SetWindowPos Me.hWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'
'    Timer_ALWAYS_ON_TOP_TO_START_WITH_ER.Enabled = False
'    Label60.BackColor = Label49.BackColor '49 58_59
'    Label60.Caption = "Me on Top @ Loader EXE Timer 20 Second NOT _ DONE"
'    MNU_ME_ON_TOP.Caption = "[__ ME_ON_TOP_=_YES __]"
'    Me.chkOnTop.Value = 1
'End Sub






Function FindWinPart_SEARCHER_hWnd_TO_EXE(SEARCH_STRING) As Long
    
    Dim Text_TEMP_ER As String
    Dim test_hWnd As Long, _
        test_pid As Long, _
        test_thread_id As Long
    
    Dim cText As String
    Dim Huge
    
    'Find the first window
    test_hWnd = FindWindow2(ByVal 0&, ByVal 0&)
    Huge = 0
    Do While test_hWnd <> 0
        
        Text_TEMP_ER = GetFileFromhWnd(test_hWnd)
        'If InStr(GetWindowTitle(test_hWnd), SEARCH_STRING) > 0 Then
        If InStr(UCase(Text_TEMP_ER), UCase(SEARCH_STRING)) > 0 Then
            Huge = Huge + 1
            
            'Debug.Print Str(Huge) + "__" + Str(test_hWnd) + "__" + Text_TEMP_ER
            'Debug.Print Str(Huge) + "__" + Str(test_hWnd) + "__" + GetWindowTitle(test_hWnd)
            'Debug.Print Str(Huge) + "__" + Str(test_hWnd) + "__" + GetWindowClass(test_hWnd)
            'Debug.Print vbCrLf
            'EXIT SUB
            Huge = Huge
        End If
            
    'retrieve the next window
    test_hWnd = GetWindow(test_hWnd, GW_hWndNEXT)
    
    Loop
End Function


Function FindWinPart_SEARCHER_NOT_ME(SEARCH_STRING) As Long
    FindWinPart_SEARCHER_NOT_ME = False
    
    Dim test_hWnd As Long, _
        test_pid As Long, _
        test_thread_id As Long
    
    Dim cText As String
    Dim CLASS_NAME As String
    Dim XGO
    Dim CLASS_NAME_______________

    'Find the first window
    test_hWnd = FindWindow2(ByVal 0&, ByVal 0&)
    Do While test_hWnd <> 0
        
        CLASS_NAME_______________ = GetWindowClass(test_hWnd)
        ' If InStr(UCase(CLASS_NAME_______________), UCase(SEARCH_STRING)) > 0 Then XGO = True
        ' If InStr(UCase(GetWindowTitle(test_hWnd)), UCase(SEARCH_STRING)) > 0 Then XGO = True
        If GetWindowTitle(test_hWnd) = SEARCH_STRING Then XGO = True
        If CLASS_NAME_______________ = "ThunderRT6FormDC" Then
        If XGO = True And test_hWnd <> Form1.hWnd Then
            FindWinPart_SEARCHER_NOT_ME = test_hWnd: Exit Function
        End If
        End If
        XGO = False
            
    'retrieve the next window
    test_hWnd = GetWindow(test_hWnd, GW_hWndNEXT)
    
    Loop
End Function



Function FindWinPart_SEARCHER(SEARCH_STRING) As Long
    FindWinPart_SEARCHER = False
    
    Dim test_hWnd As Long, _
        test_pid As Long, _
        test_thread_id As Long
    
    Dim cText As String
    Dim CLASS_NAME As String
    Dim XGO
    Dim CLASS_NAME_______________

    'Find the first window
    test_hWnd = FindWindow2(ByVal 0&, ByVal 0&)
    Do While test_hWnd <> 0
        
        CLASS_NAME_______________ = GetWindowClass(test_hWnd)
        If InStr(UCase(CLASS_NAME_______________), UCase(SEARCH_STRING)) > 0 Then XGO = True
        If InStr(UCase(GetWindowTitle(test_hWnd)), UCase(SEARCH_STRING)) > 0 Then XGO = True
        If XGO = True Then
            FindWinPart_SEARCHER = test_hWnd: Exit Function
        End If
            
    'retrieve the next window
    test_hWnd = GetWindow(test_hWnd, GW_hWndNEXT)
    
    Loop
End Function



Function FindWinPart_EXPLORER() As String
    
    FindWinPart_EXPLORER = ""
    
    Dim test_hWnd As Long, _
        test_pid As Long, _
        test_thread_id As Long
    
    Dim cText As String
    Dim CLASS_NAME As String
    Dim XGO
    Dim CLASS_NAME_______________


    ' SET ONE DUMMY OR SPLIT NOT HAPPEN LATR IF ONLY 1
    FindWinPart_EXPLORER = "0000" + vbCrLf
    
    'Find the first window
    test_hWnd = FindWindow2(ByVal 0&, ByVal 0&)
    Do While test_hWnd <> 0
        XGO = False
        CLASS_NAME_______________ = GetWindowClass(test_hWnd)
        If InStr(CLASS_NAME_______________, "CabinetWClass") > 0 Then XGO = True
        ' If InStr(UCase(GetWindowTitle(test_hWnd)), UCase(SEARCH_STRING)) > 0 Then XGO = True
        If XGO = True Then
            FindWinPart_EXPLORER = FindWinPart_EXPLORER + Trim(Str(test_hWnd)) + vbCrLf
        End If
            
    'retrieve the next window
    test_hWnd = GetWindow(test_hWnd, GW_hWndNEXT)
    
    Loop
End Function






Public Function GethWndFromProcess(p_lngProcessId As Long) As Long
    Dim lngDesktop As Long
    Dim lngChild As Long
    Dim lngChildProcessID As Long
    
    On Error Resume Next
    
    'get the hWnd of the desktop
    lngDesktop = GetDesktopWindow()
    lngDesktop = FindWindowDLL(ByVal 0&, ByVal 0&)

    lngChild = GetWindow(lngDesktop, GW_CHILD)
    
    Do While lngChild <> 0
    
        'get the ThreadProcessID of the window
        Call GetWindowThreadProcessId(lngChild, lngChildProcessID)
        
        'if ProcessId matches the parameter passed
        'then return that value
        If lngChildProcessID = p_lngProcessId Then
            GethWndFromProcess = lngChild
            Exit Do
        End If
        
        'not found, continue enumeration
        lngChild = GetWindow(lngChild, GW_hWndNEXT)
    Loop
End Function



Private Sub txtRect_CLICK()
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText txtRect
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub

Private Sub txtStyle_CLICK()
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText txtStyle
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub

Private Sub txtParentX_CLICK()
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText txtParentX
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub

Private Sub txtParentHX_CLICK()
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText txtParentHX
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub

Private Sub txtParent_CLICK()
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText txtParent
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub

Private Sub TxtPID_CLICK()
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText TxtPID
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub

Private Sub txthWndHX_CLICK()
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText txthWndHX
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub
Private Sub txthWnd_CLICK()
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText txthWnd
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub
Private Sub txtClass_CLICK()
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText txtClass
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub

Private Sub TxtEXE_CLICK()
    ' MNU_HOOVER_20_SECOND
    If TIMER2_TIMER_BEGAN > 0 Then Exit Sub
    
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText TxtEXE
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub

Private Sub txtParentClass_CLICK()
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText txtParentClass
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub
Private Sub txtParentClassX_CLICK()
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText txtParentClassX
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub
Private Sub txtParentTextX_CLICK()
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText txtParentTextX
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub
Private Sub txtParentText_CLICK()
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText txtParentText
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub
Private Sub txtMemoryUsage_Click()
'txtMemoryUsage_HERE
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText txtMemoryUsage
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub

Private Sub txtTitle_CLICK()
    On Error GoTo ENDER
    Clipboard.Clear
    Clipboard.SetText txtTitle
    Exit Sub
ENDER:
    DoEvents
    Sleep 100
    Resume
End Sub


