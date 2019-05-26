VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000015&
   Caption         =   "EliteSpy+ by Andrea Batina 2001 & Also Matthew Lancaster __ Contact ME __ BT __ 07722224555"
   ClientHeight    =   7116
   ClientLeft      =   60
   ClientTop       =   936
   ClientWidth     =   15972
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7116
   ScaleWidth      =   15972
   Begin VB.Timer TIMER_TO_RESIZE 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   15690
      Top             =   3825
   End
   Begin VB.ListBox List_SORT_FOR_AHK_LIMITER 
      Height          =   456
      ItemData        =   "frmMain.frx":0442
      Left            =   11115
      List            =   "frmMain.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   121
      Top             =   5010
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.TextBox TxtPID 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1296
      Locked          =   -1  'True
      TabIndex        =   118
      Top             =   432
      Width           =   1044
   End
   Begin VB.CommandButton cmdMoveMax 
      Caption         =   "Move Window Maximum"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   924
      Left            =   1128
      TabIndex        =   117
      Top             =   5844
      Width           =   948
   End
   Begin VB.TextBox txtParentHX 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2364
      Locked          =   -1  'True
      TabIndex        =   115
      Top             =   2016
      Width           =   924
   End
   Begin VB.TextBox txthWndHX 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2364
      Locked          =   -1  'True
      TabIndex        =   114
      Top             =   696
      Width           =   924
   End
   Begin VB.Timer Timer_CONNECTED_TO_THE_INTERNET 
      Interval        =   10000
      Left            =   14196
      Top             =   3468
   End
   Begin VB.Timer Timer_Pause_Update 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   14568
      Top             =   3468
   End
   Begin VB.PictureBox picCrossHair 
      BorderStyle     =   0  'None
      Height          =   396
      Left            =   4956
      Picture         =   "frmMain.frx":0446
      ScaleHeight     =   396
      ScaleWidth      =   384
      TabIndex        =   112
      Top             =   324
      Width           =   384
   End
   Begin VB.Timer TIMER_ColorCycle 
      Interval        =   50
      Left            =   14916
      Top             =   3468
   End
   Begin ComctlLib.ListView lstProcess_3_SORTER_ListView 
      Height          =   1332
      Left            =   13956
      TabIndex        =   111
      Top             =   924
      Width           =   2004
      _ExtentX        =   3535
      _ExtentY        =   2350
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ListView lstProcess_2_ListView 
      Height          =   1332
      Left            =   11844
      TabIndex        =   110
      Top             =   924
      Width           =   2004
      _ExtentX        =   3535
      _ExtentY        =   2350
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
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
      TabIndex        =   108
      Top             =   5460
      Width           =   1320
   End
   Begin VB.CommandButton cmdNormal 
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   48
      TabIndex        =   107
      Top             =   5844
      Width           =   1068
   End
   Begin VB.CommandButton cmdMinimize 
      Caption         =   "Minimize"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   48
      TabIndex        =   106
      Top             =   6156
      Width           =   1068
   End
   Begin VB.CommandButton cmdMaximize 
      Caption         =   "Maximize"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   48
      TabIndex        =   105
      Top             =   6456
      Width           =   1068
   End
   Begin VB.CommandButton cmdFlash 
      Caption         =   "Flash"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2088
      TabIndex        =   104
      Top             =   5844
      Width           =   1068
   End
   Begin VB.CommandButton cmdEnable 
      Caption         =   "Enable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2088
      TabIndex        =   103
      Top             =   6156
      Width           =   1068
   End
   Begin VB.CommandButton cmdDisable 
      Caption         =   "Disable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2088
      TabIndex        =   102
      Top             =   6456
      Width           =   1068
   End
   Begin VB.CommandButton cmdSetTitle 
      Caption         =   "Set Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3168
      TabIndex        =   101
      Top             =   5844
      Width           =   1068
   End
   Begin VB.CommandButton cmdOnTop 
      Caption         =   "OnTop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3168
      TabIndex        =   100
      Top             =   6156
      Width           =   1068
   End
   Begin VB.CommandButton cmdNotOnTop 
      Caption         =   "Not OnTop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3168
      TabIndex        =   99
      Top             =   6456
      Width           =   1068
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4248
      TabIndex        =   98
      Top             =   5844
      Width           =   1104
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4248
      TabIndex        =   97
      Top             =   6156
      Width           =   1104
   End
   Begin VB.CommandButton cmdTerminate 
      Caption         =   "Terminate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4248
      TabIndex        =   96
      Top             =   6456
      Width           =   1104
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
      Left            =   8148
      Locked          =   -1  'True
      TabIndex        =   93
      Top             =   6864
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
      Left            =   8148
      Locked          =   -1  'True
      TabIndex        =   92
      Top             =   6612
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
      Left            =   8148
      Locked          =   -1  'True
      TabIndex        =   89
      Top             =   6108
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
      Left            =   8148
      Locked          =   -1  'True
      TabIndex        =   88
      Top             =   6360
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
      Left            =   8148
      Locked          =   -1  'True
      TabIndex        =   84
      Top             =   5604
      Width           =   2688
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
      Left            =   8148
      Locked          =   -1  'True
      TabIndex        =   83
      Top             =   5856
      Width           =   2688
   End
   Begin VB.DirListBox Dir2 
      Height          =   756
      Left            =   13284
      TabIndex        =   82
      Top             =   2280
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Timer Timer_VB_MAXIMIZE 
      Interval        =   1000
      Left            =   13848
      Top             =   3468
   End
   Begin VB.Timer Timer_FOREGROUND_WINDOW_CHANGE_02 
      Interval        =   10
      Left            =   13500
      Top             =   3468
   End
   Begin VB.Timer Timer_FOREGROUND_WINDOW_CHANGE 
      Interval        =   100
      Left            =   11100
      Top             =   3468
   End
   Begin VB.Timer Timer_ALWAYS_ON_TOP_TO_START_WITH_ER 
      Interval        =   1000
      Left            =   12816
      Top             =   3468
   End
   Begin VB.Timer Timer_IS_WINAMP_RUNNING_DO_CID 
      Interval        =   1000
      Left            =   13152
      Top             =   3468
   End
   Begin VB.Timer Timer_VIRCOP 
      Enabled         =   0   'False
      Left            =   12456
      Top             =   3480
   End
   Begin VB.DirListBox Dir1 
      Height          =   528
      Left            =   14640
      TabIndex        =   67
      Top             =   2784
      Visible         =   0   'False
      Width           =   1272
   End
   Begin VB.Timer Timer_GET_KEY_ASYNC_STATE 
      Interval        =   10
      Left            =   12120
      Top             =   3468
   End
   Begin VB.Timer Timer_01_VB_HELP_BOX_02_MSDN 
      Interval        =   200
      Left            =   11784
      Top             =   3468
   End
   Begin VB.Timer Timer_SET_FOCUS 
      Interval        =   1000
      Left            =   11448
      Top             =   3468
   End
   Begin VB.Timer Timer_SHOW_THE_TIME 
      Interval        =   1000
      Left            =   11100
      Top             =   3828
   End
   Begin VB.Timer Timer_EXPLORER_GONE 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   14196
      Top             =   3828
   End
   Begin VB.Timer SHOW_THE_TIME 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13848
      Top             =   3828
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   13980
      TabIndex        =   45
      Top             =   4305
      Visible         =   0   'False
      Width           =   996
   End
   Begin VB.ListBox lstProcess_2_ 
      Height          =   324
      IntegralHeight  =   0   'False
      Left            =   11115
      TabIndex        =   43
      Top             =   4320
      Visible         =   0   'False
      Width           =   1128
   End
   Begin VB.Timer Timer_1_SECOND 
      Interval        =   1000
      Left            =   10860
      Top             =   3348
   End
   Begin VB.ListBox lstProcess_3_SORTER 
      Height          =   276
      IntegralHeight  =   0   'False
      Left            =   11100
      Sorted          =   -1  'True
      TabIndex        =   30
      Top             =   4665
      Visible         =   0   'False
      Width           =   1728
   End
   Begin VB.Timer Timer_EnumProcess 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11448
      Top             =   3828
   End
   Begin VB.Timer Timer_PROCESS_RELOADER_WATCHER_07 
      Interval        =   1000
      Left            =   15252
      Top             =   3828
   End
   Begin VB.Timer Timer_PROCESS_RELOADER_WATCHER_05 
      Interval        =   1000
      Left            =   14904
      Top             =   3828
   End
   Begin VB.Timer Timer_PROCESS_RELOADER_WATCHER_04 
      Interval        =   1000
      Left            =   14556
      Top             =   3828
   End
   Begin VB.Timer Timer_PROCESS_RELOADER_WATCHER_03 
      Interval        =   1000
      Left            =   13500
      Top             =   3828
   End
   Begin VB.Timer Timer_PROCESS_RELOADER_WATCHER_02 
      Interval        =   1000
      Left            =   12456
      Top             =   3828
   End
   Begin VB.Timer Timer_PROCESS_RELOADER_WATCHER_01 
      Interval        =   1000
      Left            =   12120
      Top             =   3828
   End
   Begin VB.Timer Timer_COMMAND_EXE_MENU_NOTEPAD_PICK_UP 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13152
      Top             =   3828
   End
   Begin VB.TextBox TxtEXE 
      BorderStyle     =   0  'None
      Height          =   228
      Left            =   3348
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   34
      Width           =   7488
   End
   Begin VB.TextBox txtParentX 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1300
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2808
      Width           =   2004
   End
   Begin VB.TextBox txtParentTextX 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1300
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3072
      Width           =   2004
   End
   Begin VB.TextBox txtParentClassX 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1300
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3336
      Width           =   2004
   End
   Begin VB.TextBox txtRect 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1300
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1752
      Width           =   2004
   End
   Begin VB.TextBox txtParentClass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1300
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2544
      Width           =   2004
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
      Left            =   48
      TabIndex        =   15
      ToolTipText     =   "STORE IN REISTRY FOR START UP OVERRIDE TIMER IF DO"
      Top             =   4824
      Width           =   5304
   End
   Begin VB.TextBox txtParentText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1300
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2280
      Width           =   2004
   End
   Begin VB.CommandButton cmdTerminateProcess 
      Caption         =   "Terminate"
      Height          =   288
      Left            =   15000
      TabIndex        =   12
      Top             =   4620
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.ListBox lstProcess 
      Height          =   444
      IntegralHeight  =   0   'False
      Left            =   14640
      TabIndex        =   11
      Top             =   2268
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Timer Timer_MOUSE_CORD 
      Interval        =   10
      Left            =   11784
      Top             =   3828
   End
   Begin VB.TextBox txtStyle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1300
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1488
      Width           =   2004
   End
   Begin VB.TextBox txtParent 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1300
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2016
      Width           =   1044
   End
   Begin VB.TextBox txtClass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1300
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1224
      Width           =   2004
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1300
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   2004
   End
   Begin VB.TextBox txthWnd 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1300
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   696
      Width           =   1044
   End
   Begin VB.Label LAB_MAXIMIZE_VB_KEEP_RUNNER 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "MAX VB-KR"
      Height          =   216
      Left            =   2316
      TabIndex        =   130
      Top             =   3600
      Width           =   876
   End
   Begin VB.Label LAB_MAXIMIZE_GOODSYNC2GO 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "MAX GOODSYNC2GO"
      Height          =   216
      Left            =   48
      TabIndex        =   129
      Top             =   3840
      Width           =   1632
   End
   Begin VB.Label LAB_MAXIMIZE_HUBIC 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "MAX HUBIC"
      Height          =   216
      Left            =   1416
      TabIndex        =   128
      Top             =   3600
      Width           =   876
   End
   Begin VB.Label Lab_KILL_EXPLORER 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "KILL EXPLORER"
      Height          =   216
      Left            =   4104
      TabIndex        =   127
      Top             =   3600
      Width           =   1236
   End
   Begin VB.Label Label_CLOSE_GOODSYNC 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "CLOSE GOODSYNC"
      Height          =   228
      Left            =   3828
      TabIndex        =   126
      Top             =   3840
      Width           =   1512
   End
   Begin VB.Label Label_KILL_HUBIC 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "KILL HUBIC"
      Height          =   228
      Left            =   2904
      TabIndex        =   125
      Top             =   3840
      Width           =   900
   End
   Begin VB.Label Label_CLOSE_HWND 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "CLOSE HWND"
      Height          =   240
      Left            =   9108
      TabIndex        =   124
      Top             =   1992
      Width           =   1728
   End
   Begin VB.Label Lab_KILL_AHK 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "KILL AHK"
      Height          =   216
      Left            =   3348
      TabIndex        =   123
      Top             =   3600
      Width           =   732
   End
   Begin VB.Label LAB_MAXIMIZE_GOODSYNC 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "MAX GOODSYNC"
      Height          =   216
      Left            =   48
      TabIndex        =   122
      Top             =   3600
      Width           =   1344
   End
   Begin VB.Label Label66 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Kill PID"
      Height          =   240
      Left            =   2376
      TabIndex        =   120
      Top             =   432
      Width           =   912
   End
   Begin VB.Label HWND_BLOCK 
      Caption         =   "Process ID:"
      Height          =   240
      Index           =   1
      Left            =   48
      TabIndex        =   119
      Top             =   432
      Width           =   1224
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
      Left            =   4392
      TabIndex        =   116
      Top             =   1992
      Width           =   960
   End
   Begin VB.Label Label_KILL_WSCRIPT 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "KILL WSCRIPT"
      Height          =   228
      Left            =   1716
      TabIndex        =   113
      Top             =   3840
      Width           =   1164
   End
   Begin VB.Image imgCursor 
      Height          =   276
      Left            =   5280
      MouseIcon       =   "frmMain.frx":0D10
      Top             =   372
      Width           =   276
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   4500
      Picture         =   "frmMain.frx":15DA
      Top             =   276
      Width           =   384
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
      Left            =   48
      TabIndex        =   109
      Top             =   5472
      Width           =   2616
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
      Left            =   5448
      TabIndex        =   95
      Top             =   6864
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
      Left            =   5436
      TabIndex        =   94
      Top             =   6612
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
      Left            =   5436
      TabIndex        =   91
      Top             =   6108
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
      Left            =   5436
      TabIndex        =   90
      Top             =   6360
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
      Left            =   5436
      TabIndex        =   87
      Top             =   5604
      Width           =   2700
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
      Left            =   5436
      TabIndex        =   86
      Top             =   5856
      Width           =   2700
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
      Left            =   5436
      TabIndex        =   85
      Top             =   5352
      Width           =   5400
   End
   Begin VB.Label Label62 
      Caption         =   "TASKKILLER /F /IM *"
      Height          =   240
      Left            =   5436
      TabIndex        =   81
      Top             =   1740
      Width           =   1764
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00B7FFB5&
      BorderWidth     =   5
      X1              =   5460
      X2              =   10800
      Y1              =   3948
      Y2              =   3948
   End
   Begin VB.Label Label61 
      BackColor       =   &H0068F9CA&
      Caption         =   "C8"
      Height          =   240
      Left            =   12576
      TabIndex        =   80
      Top             =   2400
      Visible         =   0   'False
      Width           =   216
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
      Left            =   48
      TabIndex        =   79
      Top             =   4512
      Width           =   5304
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00B7FFB5&
      BorderWidth     =   5
      X1              =   5460
      X2              =   10800
      Y1              =   3372
      Y2              =   3372
   End
   Begin VB.Label Label59 
      BackColor       =   &H00C0FFFF&
      Caption         =   "C7"
      Height          =   240
      Left            =   12336
      TabIndex        =   78
      Top             =   2400
      Visible         =   0   'False
      Width           =   216
   End
   Begin VB.Label Label58 
      Caption         =   "C5"
      Height          =   240
      Left            =   12084
      TabIndex        =   77
      Top             =   2400
      Visible         =   0   'False
      Width           =   228
   End
   Begin VB.Label Label55 
      Alignment       =   2  'Center
      Caption         =   "TASKKILLER PID NAME"
      Height          =   240
      Left            =   5436
      TabIndex        =   76
      Top             =   2328
      Width           =   1776
   End
   Begin VB.Label Label57 
      BackColor       =   &H00C0FFFF&
      Caption         =   "TASKKILLER COMMAND LINE EXECUTE STATUS LINE"
      Height          =   240
      Left            =   5436
      TabIndex        =   75
      Top             =   3996
      Width           =   5400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   5
      X1              =   5460
      X2              =   10800
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label56 
      Alignment       =   2  'Center
      Caption         =   "TASKKILLER PID_NUMERIC OPTION_INSTEAD"
      Height          =   240
      Left            =   7236
      TabIndex        =   74
      Top             =   2328
      Width           =   3600
   End
   Begin VB.Label Command_Screen_Shot_Auto_ClipBoard_er 
      Caption         =   "Screen Shot Auto ClipBoard_er when Spy_er && Archive Mode _OFF_ Hitt Button Here to Change"
      Height          =   408
      Left            =   48
      TabIndex        =   73
      Top             =   4092
      Width           =   5304
   End
   Begin VB.Label cmdCopy 
      Alignment       =   2  'Center
      Caption         =   "Copy All to Clipboard"
      Height          =   252
      Left            =   3360
      TabIndex        =   72
      Top             =   3276
      Width           =   1992
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B7FFB5&
      BorderWidth     =   5
      X1              =   5460
      X2              =   10800
      Y1              =   1140
      Y2              =   1140
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
      Left            =   8760
      TabIndex        =   71
      Top             =   5100
      Width           =   2076
   End
   Begin VB.Label Label54 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CONNECT TO THE INTERNET"
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
      Height          =   300
      Left            =   24
      TabIndex        =   70
      Top             =   6768
      Width           =   5292
   End
   Begin VB.Label Label53 
      Caption         =   "Enumerate Event"
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
      Left            =   11064
      TabIndex        =   69
      ToolTipText     =   "Pause Update for 20 Second"
      Top             =   456
      Width           =   6072
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
      Left            =   14160
      TabIndex        =   66
      Top             =   60
      Width           =   336
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
      Left            =   15132
      TabIndex        =   65
      Top             =   34
      Width           =   2028
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
      Left            =   3360
      TabIndex        =   64
      Top             =   1992
      Width           =   960
   End
   Begin VB.Label Label46 
      Caption         =   "Process Counter:"
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
      Left            =   11064
      TabIndex        =   63
      Top             =   34
      Width           =   2436
   End
   Begin VB.Label Label47 
      BackColor       =   &H0080FFFF&
      Caption         =   "C4"
      Height          =   240
      Left            =   11820
      TabIndex        =   62
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label49 
      BackColor       =   &H00C0FFC0&
      Caption         =   "C3"
      Height          =   240
      Left            =   11592
      TabIndex        =   61
      Top             =   2400
      Visible         =   0   'False
      Width           =   216
   End
   Begin VB.Label Label45 
      BackColor       =   &H00FFC0FF&
      Caption         =   "C2"
      Height          =   240
      Left            =   11340
      TabIndex        =   60
      Top             =   2400
      Visible         =   0   'False
      Width           =   216
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0C0&
      Caption         =   "C1"
      Height          =   240
      Left            =   11100
      TabIndex        =   59
      Top             =   2400
      Visible         =   0   'False
      Width           =   216
   End
   Begin VB.Label Label44 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   768
      Left            =   11124
      TabIndex        =   58
      Top             =   5520
      Visible         =   0   'False
      Width           =   4776
   End
   Begin VB.Label Label43 
      Caption         =   "TASKKILLER BY HWND HANDLE _ PROCESS KILL FORCE"
      Height          =   240
      Left            =   5436
      TabIndex        =   57
      Top             =   3096
      Width           =   5400
   End
   Begin VB.Label Label42 
      BackColor       =   &H00FFE6E7&
      Caption         =   "Label42"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2652
      Left            =   17892
      TabIndex        =   56
      Top             =   180
      Width           =   2952
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
      Left            =   5436
      TabIndex        =   55
      Top             =   5100
      Width           =   3312
   End
   Begin VB.Label Label34 
      Caption         =   "TASKKILLER BY HWND HANDLE _ POST MESSENGER"
      Height          =   240
      Left            =   5436
      TabIndex        =   54
      Top             =   2844
      Width           =   5400
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "WILD CARD ****"
      Height          =   240
      Left            =   5436
      TabIndex        =   53
      Top             =   864
      Width           =   1764
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Without WildCard Butter"
      Height          =   240
      Left            =   7224
      TabIndex        =   52
      Top             =   864
      Width           =   1848
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "KILLER With /F FORCE"
      Height          =   240
      Left            =   9096
      TabIndex        =   51
      Top             =   864
      Width           =   1740
   End
   Begin VB.Label Label22 
      Caption         =   "TASKKILLER BY PROCESS NUMBER"
      Height          =   240
      Left            =   5436
      TabIndex        =   50
      ToolTipText     =   "PRESS HERE TO KILL BY PROCESS NUMBER"
      Top             =   3420
      Width           =   5400
   End
   Begin VB.Label Label_FORM_BACK_COLOUR 
      BackColor       =   &H00005959&
      Caption         =   "Label_FORM BACK_Form Ground COLOUR"
      Height          =   468
      Left            =   11100
      TabIndex        =   49
      Top             =   2856
      Visible         =   0   'False
      Width           =   1896
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TASKKILLER /F /IM * /T"
      Height          =   240
      Left            =   9096
      TabIndex        =   48
      Top             =   1188
      Width           =   1740
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFC0FF&
      Caption         =   "TASKKILLER /F /IM /T"
      Height          =   240
      Left            =   9108
      TabIndex        =   47
      Top             =   1464
      Width           =   1740
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0FFC0&
      Caption         =   "TASKKILLER /F /IM"
      Height          =   240
      Left            =   9108
      TabIndex        =   46
      Top             =   1740
      Width           =   1740
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0FFC0&
      Caption         =   "TASKKILLER /F /IM"
      Height          =   240
      Left            =   7236
      TabIndex        =   44
      Top             =   1992
      Width           =   1848
   End
   Begin VB.Label Label25 
      Caption         =   "TASKLIST GO"
      Height          =   240
      Left            =   5436
      TabIndex        =   42
      Top             =   2592
      Width           =   5400
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
      Left            =   48
      TabIndex        =   41
      Top             =   5112
      Width           =   5304
   End
   Begin VB.Label Label24 
      Caption         =   "2.. Use 20 or 40 Second Timer and Hover Land On/nover the window you want to spy"
      Height          =   888
      Left            =   3360
      TabIndex        =   40
      Top             =   2352
      Width           =   1992
   End
   Begin VB.Label Label23 
      Caption         =   "HITT BUTTON TO CONFIRM SELECTION SUBMIT COMPLY _AUTHORITY _ABUSE _AFFIRMATIVE APPLY REQUIRED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   828
      Left            =   5436
      TabIndex        =   39
      Top             =   4248
      Width           =   5400
   End
   Begin VB.Label Label30 
      Caption         =   "TASKKILLER COMMAND LINE GENERATED __ OF SELECTOR"
      Height          =   240
      Left            =   5436
      TabIndex        =   38
      ToolTipText     =   "PRESS HERE TO KILL BY PROCESS NAME"
      Top             =   3672
      Width           =   5400
   End
   Begin VB.Label Label21 
      Caption         =   "TASKKILLER /IM"
      Height          =   240
      Left            =   7224
      TabIndex        =   37
      Top             =   1740
      Width           =   1860
   End
   Begin VB.Label Label20 
      Caption         =   "TASKKILLER /IM /T"
      Height          =   240
      Left            =   7224
      TabIndex        =   36
      Top             =   1464
      Width           =   1860
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFC0FF&
      Caption         =   "TASKKILLER /F /IM /T"
      Height          =   240
      Left            =   7224
      TabIndex        =   35
      Top             =   1188
      Width           =   1848
   End
   Begin VB.Label Label18 
      Caption         =   "TASKKILLER /IM *"
      Height          =   240
      Left            =   5436
      TabIndex        =   34
      Top             =   2004
      Width           =   1764
   End
   Begin VB.Label Label17 
      Caption         =   "TASKKILLER /IM * /T"
      Height          =   240
      Left            =   5436
      TabIndex        =   33
      Top             =   1464
      Width           =   1764
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFC0C0&
      Caption         =   "TASKKILLER /F /IM * /T"
      Height          =   240
      Left            =   5436
      TabIndex        =   32
      Top             =   1188
      Width           =   1764
   End
   Begin VB.Label Label13 
      Caption         =   "Select Process And Then Button And Then Confirm_er To Killer"
      Height          =   240
      Left            =   5436
      TabIndex        =   31
      Top             =   600
      Width           =   5400
   End
   Begin VB.Label Label15 
      Caption         =   "Scan Processor For A Moment Timer Second _ "
      Height          =   240
      Left            =   5436
      TabIndex        =   29
      Top             =   348
      Width           =   5400
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Put Whole File on Clipboard"
      Height          =   228
      Left            =   48
      TabIndex        =   28
      Top             =   36
      Width           =   2088
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Goto File Name"
      Height          =   228
      Left            =   2160
      TabIndex        =   27
      Top             =   36
      Width           =   1164
   End
   Begin VB.Label HWND_BLOCK 
      Caption         =   "Parent X:"
      Height          =   240
      Index           =   9
      Left            =   48
      TabIndex        =   25
      Top             =   2808
      Width           =   1224
   End
   Begin VB.Label HWND_BLOCK 
      Caption         =   "Parent X Text:"
      Height          =   240
      Index           =   10
      Left            =   48
      TabIndex        =   24
      Top             =   3072
      Width           =   1224
   End
   Begin VB.Label HWND_BLOCK 
      Caption         =   "Parent X Class:"
      Height          =   240
      Index           =   11
      Left            =   48
      TabIndex        =   23
      Top             =   3336
      Width           =   1224
   End
   Begin VB.Label HWND_BLOCK 
      Caption         =   "Rectangle:"
      Height          =   240
      Index           =   5
      Left            =   48
      TabIndex        =   19
      Top             =   1752
      Width           =   1224
   End
   Begin VB.Label HWND_BLOCK 
      Caption         =   "Parent Class:"
      Height          =   240
      Index           =   8
      Left            =   48
      TabIndex        =   17
      Top             =   2544
      Width           =   1224
   End
   Begin VB.Label HWND_BLOCK 
      Caption         =   "Parent Text:"
      Height          =   240
      Index           =   7
      Left            =   48
      TabIndex        =   14
      Top             =   2280
      Width           =   1224
   End
   Begin VB.Label HWND_BLOCK 
      Caption         =   "Style:"
      Height          =   240
      Index           =   4
      Left            =   48
      TabIndex        =   9
      Top             =   1488
      Width           =   1224
   End
   Begin VB.Label HWND_BLOCK 
      Caption         =   "Parent:"
      Height          =   240
      Index           =   6
      Left            =   48
      TabIndex        =   7
      Top             =   2016
      Width           =   1224
   End
   Begin VB.Label HWND_BLOCK 
      Caption         =   "Class:"
      Height          =   240
      Index           =   3
      Left            =   48
      TabIndex        =   5
      Top             =   1236
      Width           =   1224
   End
   Begin VB.Label Label1 
      Caption         =   "1.. Drag this icon over the window you want to spy/nMouse Move to Top and Form Will Shirnk_er Until Done to Make Room Under"
      Height          =   1224
      Left            =   3360
      TabIndex        =   4
      Top             =   744
      Width           =   1992
   End
   Begin VB.Label HWND_BLOCK 
      Caption         =   "Title:"
      Height          =   240
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   960
      Width           =   1224
   End
   Begin VB.Label HWND_BLOCK 
      Caption         =   "hWnd:"
      Height          =   240
      Index           =   0
      Left            =   48
      TabIndex        =   0
      Top             =   696
      Width           =   1224
   End
   Begin VB.Label Label52 
      Caption         =   "Label52"
      Height          =   348
      Left            =   13848
      TabIndex        =   68
      Top             =   36
      Width           =   1188
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
   Begin VB.Menu MNU_ADMINISTRATOR 
      Caption         =   "ADMIN"
   End
   Begin VB.Menu MNU_ME_ON_TOP 
      Caption         =   "ME ON TOP"
   End
   Begin VB.Menu MNU_TASK_KILLER_EXPLORER_CLIPBOARD_02 
      Caption         =   "T.KILL EXPLORER CLIPPER"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_TASK_KILLER_AUTOHOTKEYS_2 
      Caption         =   "T.KILL AUTOHOTKEYS"
   End
   Begin VB.Menu MNU_KILL_MAX_AHK 
      Caption         =   "KILL AHK PID LIMITER > 100"
   End
   Begin VB.Menu MNU_WINMERGE_ON_TOP_ALLTME 
      Caption         =   "WINMERGE ON TOP ALLTIME=YES"
   End
   Begin VB.Menu MNU_SCREEN_SHOT_ME 
      Caption         =   "SCREENSHOT ME"
   End
   Begin VB.Menu MNU_SCREENSHOT_IMAGE_FOLDER 
      Caption         =   "IMAGE FOLDER"
   End
   Begin VB.Menu MNU_STARTER_FOLDER 
      Caption         =   "DELAY START_ER_FOLDER"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_DELAY_FOLDER_TOP 
      Caption         =   "DELAY FOLDER"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_KILL_EXPLORER_TOP 
      Caption         =   "KILL EXPLORER"
   End
   Begin VB.Menu MNU_KILL_NOT_RESPOND_TOP 
      Caption         =   "KILL NOT RESPOND"
   End
   Begin VB.Menu MNU_ADMIN_AUTO_START_RUN_AS 
      Caption         =   "ADMIN AUTO START RUNAS COMMAND && TEST=ON"
   End
   Begin VB.Menu MNU_SHOW_STATUS_FORM 
      Caption         =   "SHOW_STATUS_FORM"
   End
   Begin VB.Menu MNU_HOOVER_20_SECOND 
      Caption         =   "USE 20 SECOND HOOVER"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_BRing_Front 
      Caption         =   "Bring All Front"
   End
   Begin VB.Menu MNU_KILLER_A_VB_PROJECT 
      Caption         =   "KILLER_A_VB_PROJECT_AND_RELOAD_FOR_DEBUGGING_REPEATER FREEZE_ER"
   End
   Begin VB.Menu MNU_GIVE_ME_TIME 
      Caption         =   "GIVE ME TIME"
   End
   Begin VB.Menu MNU_SHOW_THE_TIME 
      Caption         =   "TIME SHOW"
   End
   Begin VB.Menu MNU_GIVE_ME_TIME_WITHER_UTC 
      Caption         =   "GIVE ME TIME AND UNI_ UTC Time Toggle = YES"
   End
   Begin VB.Menu MNU_SET_TIME_ON_DEVICE_SET_THAT_WANT_IT_HardCoder 
      Caption         =   "SET_TIME_ON_DEVICE_SET_THAT_WANT_IT_HardCoder"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_EXE_NOETPAD 
      Caption         =   "EDIT MENU EXE NOTEPAD++"
   End
   Begin VB.Menu MNU_NIRSOFT 
      Caption         =   "EXE MENU 01"
      Begin VB.Menu MNU_EXE_WINDOWSE 
         Caption         =   "WINDOWSE"
      End
      Begin VB.Menu MNU_EXE_AUTO_HOT_KEYS_SPY 
         Caption         =   "AUTO_HOT_KEYS_SPY"
      End
      Begin VB.Menu MNU_EXE_PROCESS_EXPLORER_SYS_INT_PSTART 
         Caption         =   "PROCESS_EXPLORER SYS_INT"
      End
      Begin VB.Menu MNU_EXE_NIRSOFT 
         Caption         =   "NIRSOFT"
      End
      Begin VB.Menu MNU_EXE_BOOT_KILLER_NOTEPAD 
         Caption         =   "BOOT_KILLER __NOTEPAD"
      End
   End
   Begin VB.Menu MNU_EXE_02 
      Caption         =   "EXE MENU 02"
      Visible         =   0   'False
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   1
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   2
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   3
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   4
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   5
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   6
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   7
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   8
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   9
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   10
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   11
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   12
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   13
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   14
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   15
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   16
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   17
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   18
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   19
      End
      Begin VB.Menu MNU_EXE_MENU_SCRIPT_FILL_02 
         Caption         =   "--01"
         Index           =   20
      End
   End
   Begin VB.Menu MNU_LAUNCH_SET________ 
      Caption         =   "EXE_ LAUNCH SET ________________"
      Begin VB.Menu MNU_LAUNCH_EXPLORER 
         Caption         =   "EXPLORER"
      End
      Begin VB.Menu MNU_LAUNCH_CMD 
         Caption         =   "CMD"
      End
      Begin VB.Menu MNU_LAUNCH_WINDOSE 
         Caption         =   "WINDOSE"
      End
      Begin VB.Menu MNU_LAUNCH_AUTORUNS 
         Caption         =   "AUTORUNS SYSINTERNALS"
      End
      Begin VB.Menu MNU_LAUNCH_AUTORUNS_SET_BOOT 
         Caption         =   "AUTOHOTKEY BOOT"
      End
      Begin VB.Menu MNU_LAUNCH_AUTORUNS_SET 
         Caption         =   "AUTOHOTKEY ICON SET"
      End
      Begin VB.Menu MNU_LAUNCH_PROCESS_EXPLORER 
         Caption         =   "PROCESS EXPLORER SYMANTECH"
      End
      Begin VB.Menu MNU_EXE_TASK_MANAGER 
         Caption         =   "TASK MANAGER _ SYSTEM ONE"
      End
      Begin VB.Menu MNU_LAUNCH_NIRSOFT 
         Caption         =   "NIRSOFT"
      End
      Begin VB.Menu MNU_LAUNCH_HUBIC 
         Caption         =   "HUBIC"
      End
      Begin VB.Menu MNU_LAUNCH_GRAMMARLY 
         Caption         =   "GRAMMARLY"
      End
      Begin VB.Menu MNU_LAUNCH_VB 
         Caption         =   "VB"
      End
      Begin VB.Menu MNU_LAUNCH_VB_SYNCRONIZER 
         Caption         =   "VB SYNCRONIZER"
      End
      Begin VB.Menu MNU_LAUNCH_BATCH_COMPILER 
         Caption         =   "VB BATCH COMPILER"
      End
      Begin VB.Menu MNU_LAUNCH_Shell_VBasic_6_Loader 
         Caption         =   "VB Shell_VBasic_6_Loader"
      End
      Begin VB.Menu MNU_LAUNCH_CLIPBOARD_LOGGER 
         Caption         =   "CLIPBOARD LOGGER"
      End
      Begin VB.Menu MNU_LAUNCH_KAT_MP3_RELOADER 
         Caption         =   "KAT MP3 RELOADER"
      End
      Begin VB.Menu MNU_MULTI_MENU 
         Caption         =   "MAXIMUM HIGH M MENU"
      End
   End
   Begin VB.Menu MNU_GET_EXPLORER_CURRENT_SESSION_CLIPBOADIN 
      Caption         =   "GET_EXPLORER__CLIPBOADER_SESSION"
   End
   Begin VB.Menu MNU_EXPLORER_SET________ 
      Caption         =   "EXPLORE SET _______________"
      Begin VB.Menu MNU_LAUNCH_EXPLORER_SCRIPTOR_FOLDER 
         Caption         =   "EXPLORER 01_ IN SCRIPTOR FOLDER"
      End
      Begin VB.Menu MNU_LAUNCH_EXPLORER_STARTUP_DELAY_FOLDER 
         Caption         =   "EXPLORER 02_ IN STARTUP_DELAY_FOLDER"
      End
      Begin VB.Menu MNU_LAUNCH_EXPLORER_D__DSC_2015_2017_2017_Sony_CyberShot_HX60V 
         Caption         =   "EXPLORER 03_ \DSC\2015 2017\2017 Sony CyberShot HX60V _ VIDEO"
      End
      Begin VB.Menu mnu_explorer 
         Caption         =   "EXPLORER 04_ OF MENU BE __ SELECTION"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MNU_TASK_KILLER_SET________ 
      Caption         =   "TASK KILLER SET ________________"
      Begin VB.Menu MNU_TASK_KILLER_NOT_RESPONDER_GENTAL 
         Caption         =   "TASK_KILLER_ NOT_RESPONDER"
      End
      Begin VB.Menu MNU_TASK_KILLER_NOT_RESPONDER_FORCE 
         Caption         =   "TASK_KILLER_ NOT_RESPONDER_FORCE"
      End
      Begin VB.Menu MNU_TASK_KILLER_EXPLORER_CLIPBOARD 
         Caption         =   "TASK_KILLER_ EXPLORER* AND CLIPBOARD INFO"
      End
      Begin VB.Menu MNU_TASK_KILLER 
         Caption         =   "TASK_KILLER_ PROCESS*"
      End
      Begin VB.Menu MNU_EXE_BOOT_KILLER 
         Caption         =   "TASK_KILLER_ BOOT_KILLER"
      End
      Begin VB.Menu MNU_TASK_KILLER_AUTOHOTKEYS 
         Caption         =   "TASK_KILLER_ AUTOHOTKEYS"
      End
      Begin VB.Menu MNU_TASK_KILLER_CMD 
         Caption         =   "TASK_KILLER_ CMD*"
      End
      Begin VB.Menu MNU_TASK_KILLER_CHROME 
         Caption         =   "TASK_KILLER_ CHROME*"
      End
   End
   Begin VB.Menu MNU_FILE_FOLDER_SCANNER_SET_______ 
      Caption         =   "FILE FOLDER SCANNER SET ________________"
      Begin VB.Menu MNU_FILE_LOCATOR_PRO_EXE 
         Caption         =   "00_ FILELOCATOR PRO"
      End
      Begin VB.Menu MNU_FILELOCATOR_CLIPBOARD_IMAGER 
         Caption         =   "01_ FLIE_LOCATOR_01 _CLIPBOARD_IMAGER"
      End
      Begin VB.Menu MNU_LAUNCH_INFO_RAPID_CLIPBOARD_LOGGER_TEXT 
         Caption         =   "02_ INFORAPID CLIPBOARD_LOGGER_TEXT"
      End
   End
   Begin VB.Menu MNU_RESTART_LABEL 
      Caption         =   "RESTART ________________"
   End
   Begin VB.Menu MNU_REBOOT_OPTION 
      Caption         =   "RESTART LATCHER LOCK DOUBLE"
   End
   Begin VB.Menu MNU_RESTART 
      Caption         =   "RESTART NOT FORCE AND INSTANT"
   End
   Begin VB.Menu MNU_NUKER 
      Caption         =   ""
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_NEXT_WORK_TALK_KILLER_NETWORK 
      Caption         =   "NEXT WORK TASK KILLER NETWORK OPTION"
   End
   Begin VB.Menu MNU_FULL_SCREEN_GAME_CONTROL_RETAIN 
      Caption         =   "FULL SCREEN GAME CONTROL RETAIN________________"
   End
   Begin VB.Menu MNU_PAUSE_VIRTUA_COP_2 
      Caption         =   "PAUSE_VIRTUA_COP_2"
   End
   Begin VB.Menu MNU_VIRTUA_COP_2_PAUSED_STATE 
      Caption         =   "VIRTUA_COP_2_PAUSED_STATE"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_MAXIMIZE_WINDOW_GOODSYNC 
      Caption         =   "MAXIMIZE WINDOW GOODSYNC"
   End
   Begin VB.Menu mnu_Contact_Me 
      Caption         =   "Contact ME For HardCore"
   End
   Begin VB.Menu mnu_Telephone_BT_07722224555 
      Caption         =   "Telephone BT 07722224555"
   End
   Begin VB.Menu MNU_MATT_LAN_BTINTERNET_COM 
      Caption         =   "MATT.LAN@BTINTERNET.COM"
   End
   Begin VB.Menu MNU_VERSION 
      Caption         =   "VERSION"
   End
   Begin VB.Menu Mnu_Form_Count 
      Caption         =   "Form_Count_er"
   End
   Begin VB.Menu MNU_RESET_TIMERING_FOR_AUTO_RELOADER 
      Caption         =   "RESET TIMERING DELAY FOR AUTO_RELOADER HardCoder"
   End
   Begin VB.Menu MNU_LOGG_OFF_CURRENT_USER 
      Caption         =   "LOGOFF CURRENT USER"
   End
   Begin VB.Menu MNU_LOG_BACK_IN_USER_ID_1 
      Caption         =   "LOG BACK IN USER ID _1"
   End
   Begin VB.Menu MNU_LOG_BACK_IN_USER_ID_2 
      Caption         =   "LOG BACK IN USER ID _2"
   End
   Begin VB.Menu MNU_CAMERA_DATA_ON_SD_CARD 
      Caption         =   "CAMERA DATA ON SD CARD"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_Menu_Item_Count 
      Caption         =   "Menu Item Count"
   End
   Begin VB.Menu Mnu_About 
      Caption         =   "About Info"
      Begin VB.Menu Mnu_About_Me 
         Caption         =   "About Me"
      End
      Begin VB.Menu Mnu_About_Helper 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----
'Andrea Batina - Google Search
'https://www.google.co.uk/search?q=Andrea+Batina&num=50&safe=active&source=lnms&tbm=isch&sa=X&ved=0ahUKEwjinr3s0pzeAhVpJsAKHXzUAs8Q_AUIDigB&biw=1536&bih=600#imgrc=YmIjX_nZl_QuEM:
'--------
'EliteSpy+ (with Code Generator) by Andrea Batina[Revelatek] (from psc cd)
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=28563&lngWId=1
'--------
'batinanet (Andrea Batina)
'https://github.com/batinanet
'--------
'Andrea Batina | LinkedIn
'https://www.linkedin.com/in/andrea-batina-08a28b44/?originalSubdomain=rs
'----


' -------------------------------------------------------------------------------------
' SESSION AGAIN
' FEW IMPROVEMNT ON
' IMPROVEMNTS ARE
' 1..
' CLICK IN LISTIEW BOX AND PROVIDE THE HWND INFO FILLED IN _ NEW CODE WORKING NOT DONE
' BEFORE HWND FROM PID
' 2..
' GOT NORM MAX MIN __ NOW ALSO GOT SET CORDINATE AND MOVE WINDOW CENTER TO FILL SCREEN
' JUST IN CASE BEEN MOVED OFF SCREEN OR MINITURE _ HELP THERE
' 3..
' MAKE QUICK KILL WSCRIPT.EXE BUTTON ON SCREEN
' 4..
' EXTRA THING TO NORM MIN MAX BUTTON WHEN UPDATE CHUCK DISPLAY HWND INFO EACH USE
' 5..
' 20 SEC TIMER WAS DISPLAY WRONG - FIXER
' 6..
' FINALLY MOVE MENU BUTTON AROUND SOME REMOVED DISABLED OLDER THING
' -------------------------------------------------------------------------------------
' Tue 23-Oct-2018 10:52:04
' Tue 23-Oct-2018 14:40:00 _ 4 HOUR WORK
' -------------------------------------------------------------------------------------

Option Explicit


Dim LISTVIEW_2_OR_3_HITT

'JOB TO CORRECT HERE _ __ 'NOT DOING MY RUNAS PROPER

Public EXIT_TRUE


Dim OLD_FD As String

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

Dim OHWnd_VB_EXE
Dim OHWnd_VB_CLIPPER_ERROR
Dim OHWnd_VB_LOADER
Dim OHWnd_TEAM_VIEWER
Dim O_mWnd_VB_VbaWindow_MAXIMIZE

Dim OcWnd
Dim IHWND, O_IHWND

Dim TO_SETTER

Dim Label_ProcessSetBar_Height

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
Dim PID As Long, APP_NAME_EXE_PASS_FOR_CALL_BACK As String

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

Dim ARCHIVE_Menu_Height


'--------------------------------------------------------------------------------
'    Component  : frmMain
'    Project    : EliteSpy
'
'    Description: Main form
'
'    Author     : Andrea Batina
'    Modified   : 31/10/2001
'--------------------------------------------------------------------------------
'Option Explicit

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
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Const GW_HWNDNEXT = 2
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Const TH32CS_SNAPPROCESS = &H2&
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

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' Dragging window?
Private m_bDragging As Boolean

Private Declare Function CloseHandle Lib "kernel32" _
        (ByVal hObject As Long) As Long

'-----------------------------------------------
'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
Private MenuInfo As MENUBARINFO
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

Private Type MENUBARINFO
  cbSize As Long
  rcBar As RECT
  hMenu As Long
  hWndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type
'-----------------------------------------------

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

'-----------------------------------------------
'AUTOSIZE LISTVIEW COLUMN
'Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

'Private Const LVM_FIRST = &H1000
'-----------------------------------------------
'Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column _
' As ColumnHeader = Nothing)
'
' Dim C As ColumnHeader
' If Column Is Nothing Then
'  For Each C In LV.ColumnHeaders
'   SendMessage LV.hWnd, LVM_FIRST + 30, C.Index - 1, -1
'  Next
' Else
'  SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
' End If
' LV.Refresh
'End Sub
'-----------------------------------------------
Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 
Private Const LVM_FIRST = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Private Const LVS_EX_FULLROWSELECT = &H20
     
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

Const BM_CLICK = &HF5&

Private Const WM_SETTEXT            As Long = &HC
Private Const WM_GETTEXT As Long = &HD
Private Const WM_GETTEXTLENGTH As Long = &HE

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWndParent As Long, _
     ByVal hWndChildAfter As Long, _
     ByVal lpszClass As String, _
     ByVal lpszTitle As String) _
    As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long  'MODULE 1141

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByRef lParam As Any) _
    As Long
    
'Private Const WM_SETTEXT            As Long = &HC
'Private Const WM_GETTEXT = &HD

Private Const NCBASTAT As Long = &H33
Private Const NCBNAMSZ As Long = 16
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const HEAP_GENERATE_EXCEPTIONS As Long = &H4
Private Const NCBRESET As Long = &H32
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Const RASTERCAPS As Long = 38


'Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142

'HDC SET
Private Declare Function Escape Lib "gdi32" (ByVal HDC As Long, ByVal nEscape As Long, ByVal nCount As Long, ByVal lpInData As String, lpOutData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
     
'HDC SET 2
'Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DPtoLP Lib "gdi32" (ByVal HDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
'Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal HDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal HDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function SelectClipPath Lib "gdi32" (ByVal HDC As Long, ByVal iMode As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
     
'HDC SET 3
'Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal HDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal HDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal HDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
 
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

     
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Dim sConnType As String * 255

'-----------------------
'THE UNIVERSAL TIME DOWN
'-----------------------
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
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

Private Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, _
    ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As _
    Long
Private Declare Function GetCaretBlinkTime Lib "user32" () As Long

Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long  'MODULE 1135
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long


'--------------------------------------
'--------------------------------------
'--------------------------------------
'--------------------------------------
'Excellent Code Set the File Time To ANything you want touch date Pro Keep you file time after rotate JPg's

'THIS TOGGLES A SMALL PASSWORD FEATURE THAT I EMBEDDED
Const MPassword = False

'THIS IS THE CODE THAT MUST BE ENTERED (FINISHED BY PRESSING HASH)
'WHEN THE PASSWORD TOGGLE IS TRUE
Const MSequence = 22515

Dim MEntered As Variant

Private Type FILETIME
    dwLowDate  As Long
    dwHighDate As Long
End Type

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
'Private Type SYSTEMTIME
'wYear As Integer
'wMonth As Integer
'wDayOfWeek As Integer
'wDay As Integer
'wHour As Integer
'wMinute As Integer
'wSecond As Integer
'wMilliseconds As Integer
'End Type
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

Private Declare Function FindWindowDLL _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long

Private Declare Function closewindow Lib "user32" Alias "CloseWindow" (ByVal hWnd As Long) As Long


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

'Function HideTaskBar()
'    TaskBarhWnd = FindWindow("Shell_traywnd", "")
'    If TaskBarhWnd <> 0 Then
'       Call SetWindowPos(TaskBarhWnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
'    End If
'End Function
'Function ShowTaskBar()
'    If TaskBarhWnd <> 0 Then
'       Call SetWindowPos(TaskBarhWnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
'    End If
'End Function

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
   
   ' NOT REQUIRE HERE HAVING MUCK AROUND
   ' STILL GOT ROUTINE IN CODE
   ' Strip away unwanted characters.
   ' GetWindowTitle = StripNulls(GetWindowTitle)
End Function

Public Function StripNulls(OriginalStr As String) As String
        ' This removes the extra Nulls so String comparisons will work
        If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
        End If
        StripNulls = OriginalStr
End Function

Function GetWindowClass(ByVal hWnd As Long) As String
    Dim Ret As Long, sText As String
    sText = Space(255)
    Ret = GetClassName(hWnd, sText, 255)
    sText = Left$(sText, Ret)
   GetWindowClass = sText
End Function




Function IsInternetConnected()
    Dim Ret As Long
    Ret = InternetGetConnectedStateEx(Ret, sConnType, 254, 0)
    If Ret = 1 Then
    
        Dim CONNECT_TYPE
        CONNECT_TYPE = Replace(sConnType, " Connection", "")
        CONNECT_TYPE = Left$(CONNECT_TYPE, InStr(CONNECT_TYPE, Chr(0)) - 1)
        CONNECT_TYPE = CONNECT_TYPE + " HardCore"
        
        Label54.Caption = "Connected to the Internet by " & CONNECT_TYPE
        'MsgBox "You are connected to Internet via a " & sConnType, vbInformation
        O_Ret_Connected_To_The_Internet = False
    Else
         Label54.Caption = "Not connected to Internet"
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
     
     
Private Function GetText(pHandle As Long) As String
       
    Dim Buffer As String
    Dim TextLength As Long
   
    TextLength = SendMessageAny(pHandle, WM_GETTEXTLENGTH, 0&, 0&)

    Buffer = String$(TextLength, 0&)
    SendMessageAny pHandle, WM_GETTEXT, TextLength + 1, Buffer
    GetText = Buffer
End Function
     
     
Private Sub SetFullRowSelection(ByVal hWndListView, ByVal bFullRow As Boolean)
   SendMessage hWndListView, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, ByVal CLng(bFullRow)
End Sub

Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column _
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



Private Sub Form_Activate_2222()
    Call WIDTH_AND_HEIGHT(WX, HY)
    Me.height = Me.height + 1
    Me.height = Me.height - 1
End Sub

'////////////////////////////////////////////////////////////////////
'//// FORM EVENTS
'////////////////////////////////////////////////////////////////////
Private Sub Form_Load()

Dim Label50_Width_VAR
Dim SETTING_WIDTH_LISTVIEW

TOP_HEIGHT = 20

MNU_NEXT_WORK_TALK_KILLER_NETWORK.Visible = False

Dim i As String
If App.PrevInstance = True And IsIDE = True Then
    'i = FindWindow(vbNullString, Me.Caption)
    ' ------------------------------------------------------------------------------------
    ' AT LOAD TIME THIS IS CORRECT NAME PARTIAL
    ' AFTER FORM LOAD CHANGE TO HERE
    ' "EliteSpy+ 2001 __ www.PlanetSourceCode.com" + " __ Version " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
    ' ------------------------------------------------------------------------------------
    i = FindWinPart_SEARCHER("EliteSpy+ by")
    ShowWindow i, SW_MAXIMIZE
'        On Error Resume Next
    SetWindowPos i, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End
End If

If App.PrevInstance = True Then
End
End If


'KILL ITSELF IN __.EXE KILL SOFTLY
'WHILE ISIDE LEARN
'---------------------------------
Dim VAR
If IsIDE = True Then
    PID = -1
    VAR = cProcesses.GetEXEID(PID, App.Path + "\" + App.EXEName + ".exe")
    If PID <> -1 Then
        'Call Process_HIGH_PRIORITY_CLASS(PID)
        VAR = cProcesses.Process_Kill(PID)
        Beep
        End
    End If
End If



Call Form2_Check_Project_Date.VB_PROJECT_CHECKDATE("FORM LOAD")


'THIS FEATURE WINDOWS 10 PRO
'HAS ACHIVED WORK OF MAKE EVERYTHING ADMIN
'FROM NOW
'----------
'secpol.msc
'----------
'MAYBE OTHER SETTING REQUIRE AND UNABLE TO SIMUMLATE NOT IN ADMIN UNLESS GO BACK HERE

'    Me.Hide
'    DoEvents

FIRST_RUN_FOR_TOP_AND_LEFT = 6
counter_ALWAYS_ON_TOP_TIMER = 20

Label_ProcessSetBar_Height = 400

'-----------------------------------------------------------------------
'MY CAREER AS THE SKILL TO KEEP IT GOING
'COMPARED TO YOUR PSYCHAIATRY NURSE TALK OF HAVING ABILITY TO EVEN EXIST
'2017 APRIL 24 SUNDAY AND MONDAY EARLY HOUR
'AND NOT ANY RESTRICTION ON SUBJECT OF TYPE ABOUT CAREER OF NOT ALLOWED
'TO USER
'-----------------------------------------------------------------------

'MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR ON"
'APP.PATH+"\"+App.EXENamE+".EXE"
If IsIDE = False And 1 = 2 Then
    If GetSpecialfolder(38) = "" Then
    Call RUNAS_ME
    'Me.Show
    'DoEvents
    'MsgBox "RUNAS ME ADMIN", vbMsgBoxSetForeground
    'Sleep 4000
    End
    
    End If
End If

    
'---------------------------------------------------------
'App.Caption = "EliteSpy+ by Andrea Batina 2001 __ ORIGIN UNKNOWN MALE THINKER & Also Matthew Lancaster __ Contact ME __ BT __ 07722224555"
App.title = "EliteSpy+" ' by Andrea B 2001 __ ORIGIN_UNKNOWN _www.PlanetSourceCode.com_ & Also Major Working By Matthew Lancaster __ Contact ME __ BT __ 07722224555"

Dim INFO_NOTE, INFO_NOTE_1, INFO_NOTE_2
'--------------------------------------
'INFO_NOTE_1 = "EliteSpy+ by Andrea B 2001 __ Origin_Unknown _www.PlanetSourceCode.com_ & Also Major Working By Matthew Lancaster __ Contact ME __ BT __ 07722224555"
INFO_NOTE_1 = "EliteSpy+ by Andrea Batina 2001 __ www.PlanetSourceCode.com_ & Major Worker By Matthew Lancaster __ Contact ME __ BT __ 07722224555"

INFO_NOTE_1 = "EliteSpy+ by Andrea B 2001 __ www.PlanetSourceCode.com_ & Big Timer Worker By Matthew Lancaster __ 07722224555"
INFO_NOTE_1 = "EliteSpy+ 2001 __ www.PlanetSourceCode.com" '_ & Big Timer Worker By Me"

INFO_NOTE_2 = " __ Version " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
INFO_NOTE = INFO_NOTE_1 + INFO_NOTE_2

Me.Caption = INFO_NOTE

MNU_VERSION.Caption = "EliteSpy+ Version " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision)) + " Version Mach Before 2017-Apr-22__Ver_1.0.144"

'Command1.Caption = "Screen Shot Auto ClipBoard_er when Spy_er && Archive Mode On Hitt Button Here to Change Flip Flopper"

If IsIDE = True Then
    'If 1 = 2 Then
    Debug.Print vbCrLf + "Ver. Mark Before 2017-Apr-22__"; Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
    
'    MsgBox INFO_NOTE
'    Stop
'    End
End If
'---------------------------------------------------------
    



'NOT DOING MY RUNAS PROPER
If 1 = 2 And FindWinPart_SEARCHER("GoodSync -") = 0 Then

    'HERE IS THE RUNAS CLASS
    'INCLUDE THE CLASS FOLLOW THE REM READ ME INFO AT BOTTOM LOCK
    '-----------------------------
    'IN THE __.VBP FILE PROJECT
    'Reference=*\G{565783C6-CB41-11D1-8B02-00600806D9B6}#1.2#0#C:\Windows\SysWOW64\wbem\wbemdisp.TLB#Microsoft WMI Scripting V1.2 Library
    Dim oRunas
    Set oRunas = CreateObject("runas.clsrunas", GetComputerName) 'SERVER NAME
    With oRunas
            .sDomain = "WORKGROUP"
            .sUserName = GetUserName
            .sPassword = " "
            .sCommand = "C:\Program Files\Siber Systems\GoodSync\GoodSync-v10.exe"
            .RunAs 'Call the Run As method
    End With
    Set oRunas = Nothing
    '--------------------------------------------------------------------------------------------
    'Title: Windows 2000/XP Run As without having to Type a password class
    'Description: This class enhances the brillant Run as command which ships
    'with Windows 2000/XP Run As allows you to Run a program as any user
    'without the need to log on and off.
    'My RunAs class allows Network Admins to use Run as without having to type in a password this
    'means you can run a batch file using run as with no user input.
    'This class uses the API code from my forthcoming RunApp product the only different
    'is my final product will allow admins to create encrypted files which can be passed into the
    'excuteable so the username and password won't be exposed in a batch or VBScript file.
    'To use this class compile it on the server or client you wish to use it
    'on. Register it using regsvr32 "Path to compiled runas.dll" at the command prompt.
    'This file came from Planet-Source-Code.com...the home millions of lines of source code
    'You can view comments on this code/and or vote on it at: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=33158&lngWId=1
    '
    'The author may have retained certain copyrights to this code...please observe their request and the law by reviewing all copyright conditions at the above URL.
    '--------------------------------------------------------------------------------------------
End If


'Me.BackColor = &H8000000F
Me.BackColor = Label_FORM_BACK_COLOUR.BackColor
EnumProcess_COUNTER = 20


'Sort the toggle label out
'Call Command_Screen_Shot_Auto_ClipBoard_er_Click




Call Timer_SHOW_THE_TIME_Timer

Set FS = CreateObject("Scripting.FileSystemObject")
O_lstProcess_ListCount = -1

With lstProcess_2_ListView
    .ColumnHeaders.Add , "PID", "PID", 700 - 50, lvwColumnLeft
    .ColumnHeaders.Add , "EXE", "EXE", 9000, lvwColumnLeft
    .View = lvwReport
End With
With lstProcess_3_SORTER_ListView
    .ColumnHeaders.Add , "PID", "PID", 700 - 50, lvwColumnLeft
    .ColumnHeaders.Add , "EXE SORTED", "EXE SORTED", 9000, lvwColumnLeft
    .View = lvwReport
End With
   
'Label42_Anchor
Label42.Visible = False
If 1 = 2 And GetComputerName = "7-ASUS-GL522VW" Then
    Label42.Visible = True
End If
   

' Set form width to default width (without process)
Me.width = 5580 + 120

' Make textboxes flat
MakeFlat txthWnd.hWnd
MakeFlat txthWndHX.hWnd
MakeFlat txtTitle.hWnd
MakeFlat txtClass.hWnd
MakeFlat txtRect.hWnd
MakeFlat txtParent.hWnd
MakeFlat txtParentHX.hWnd
MakeFlat txtParentText.hWnd
MakeFlat txtParentClass.hWnd
MakeFlat txtStyle.hWnd
MakeFlat txtMhWnd.hWnd

' Get value from registry
If GetSetting("EliteSpy+", "Settings", "AlwaysOnTop", "0") = "1" Then
    ' Put window on top of all others
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Me.chkOnTop.Value = 1
    MNU_ME_ON_TOP.Caption = "ME ON TOP = YES"
    Timer_ALWAYS_ON_TOP_TO_START_WITH_ER.Enabled = False
    Label60.BackColor = Label61.BackColor '61 LIME 45 RED 12 BLUE 49 58_59
    Label60.Caption = "Me on Top at Loader EXE Timer 20 Second Already ON"
    'Label60_HERE
Else
    MNU_ME_ON_TOP.Caption = "ME ON TOP = NOT"
    Label60.BackColor = Label59.BackColor '49 58_59
End If
    
' Enumerate open processes
EnumProcess

If IsIDE = True Then
    MNU_VB_ME.Caption = "VB IsIDE=True"
    'MNU_VB.Enabled = False
End If

'Call READ_EXE_MENU_FILE_AND_POPULATE_MENU
Call CHECK_EXE_MENU_FILE_EXIST_AND_CREATE_IF_ISNT_AND_READER
NOTEPAD_PLUSPLUS_OR_NOTEPAD_NORMAL_RUN_WHEN_ERROR_LOAD_TIME_ONE_MSGBOX = True

CMD_Process_STATE_TO_SET = False
CMD_Process_STATE_TO_SET_1ST = True
'Call cmdProcess_Click
    
ReDim Preserve BLOCK_RUN_1(20)
ReDim Preserve BLOCK_RUN_2(20)
ReDim Preserve BLOCK_RUN_3(20)


'    '---------------------------------------------------
'    A_DATE_TIME_PM = Mid(Format(Now, "HH-MM-SS AM/PM"), 10)
'    A_DATE_TIME = "[__ " + Format(Now, "YYYY-MM-DD-HH-MM-SS ") + A_DATE_TIME_PM + " __]"
'
'    'SHOW THE TIME A MOMENT OR MENU FLICKER EVERY SECOND
'    'MNU_SHOW_THE_TIME.Caption = "SHOW THE TIME __ " + A_DATE_TIME + " __ A MOMENT OR MENU FLICKER EVERY SECOND"
'    MNU_SHOW_THE_TIME.Caption = A_DATE_TIME
'    'AT FORM LOAD


'This is an example of how you can use this routine:
'enable full row selecting
SetFullRowSelection lstProcess_2_ListView.hWnd, True
SetFullRowSelection lstProcess_3_SORTER_ListView.hWnd, True
'disable full row selecting
'SetFullRowSelection ListView1.hwnd, False


Label23 = "HITT BUTTON TO CONFIRM SELECTION" + vbCrLf: Label23 = Label23 + "SUBMIT COMPLY _AUTHORITY _ABUSE _AFFIRMATIVE APPLY REQUIRED"
'+ vbCrLf + "ABOVE ALL DO NOT HARM ANY DR LOREN MOSHER"
Label23.height = 900
'Call cmdMemInfo_Click
    
Label42 = ""
i = ""
i = i + "REPORT BY SPECCY PAY FOR _________" + vbCrLf + vbCrLf
i = i + "COMPUTER NAME_ 7-ASUS-GL522VW" + vbCrLf + vbCrLf
i = i + "BIOS DATE __________________________ " + vbCrLf
i = i + "SIXTEEN_MARCH_2K_SIXTEENTH" + vbCrLf + vbCrLf
i = i + "Operating System ____________________" + vbCrLf
i = i + "Windows 10 Pro 64-bit" + vbCrLf + vbCrLf
i = i + "CPU ________________________________" + vbCrLf
i = i + "Intel Core i7 6700HQ @ 2.60GHz 57 °C" + vbCrLf
i = i + "Skylake 14nm Technology" + vbCrLf + vbCrLf
i = i + "RAM ______________________________" + vbCrLf
'RAM
i = i + "32.0GB Dual-Channel Unknown @ 1330MHz DDR4" + vbCrLf
i = i + "RAM THE DAY BEFORE" + vbCrLf
i = i + "16.0GB Dual-Channel Unknown @ 1063MHz DDR4" + vbCrLf
i = i + "MAXIMUM IMPROVENT OVER THOUGHT ERROR SSD DRIVE WAS RESULT OF BETTER" + vbCrLf + vbCrLf
i = i + "Motherboard ______________________" + vbCrLf
i = i + "ASUSTeK COMPUTER INC. GL552VW (U3E1) 65 °C" + vbCrLf + vbCrLf
i = i + "Graphics _________________________" + vbCrLf
i = i + "3DTV (1920x1080@60Hz)" + vbCrLf
i = i + "Generic PnP Monitor (1920x1080@60Hz)" + vbCrLf
i = i + "Intel HD Graphics 530 (ASUStek Computer Inc)" + vbCrLf
i = i + "2047MB NVIDIA GeForce GTX 960M (ASUStek Computer Inc) 52 °C" + vbCrLf
i = i + "ForceWare version: 359.46" + vbCrLf
i = i + "SLI Disabled" + vbCrLf + vbCrLf
i = i + "Storage __________________________" + vbCrLf
i = i + "1863GB Seagate ST2000LM007-1R8174 (SATA) 44 °C" + vbCrLf
'i = i + "3726GB ASMT 2115 SCSI Disk Device (USB (SATA))" + vbCrLf
'i = i + "1863GB ADATA HD710 USB Device (USB (SATA)) 44 °C" + vbCrLf
i = i + "119GB Samsung Flash Drive USB Device (USB)" + vbCrLf
i = i + "119GB Samsung Flash Drive USB Device (USB)" + vbCrLf + vbCrLf
i = i + "Optical Drives _____________________" + vbCrLf
i = i + "TSSTcorp CDDVDW SU-228HB" + vbCrLf + vbCrLf
i = i + "Audio ____________________________" + vbCrLf
i = i + "Conexant SmartAudio HD" + vbCrLf
Label42 = i

'COUNTER
'Label46_HERE
Label46.Left = Label23.Left + Label23.width + 100
Label53.Left = Label46.Left

'WINGDING3 UP ARROW
'Label51_Here
'INNER SMALLER
Label51.width = 340

'----------------------------
'MATCH WING DING INNER SMALLER
'__ CODED ADJUSTABLE .TOP TO CENTER THE SYMBOL WINGDING FONT SIZING
'INNER SMALLER
'----------------------------
'Label51_Here
'UP
'DOWN
'EQUAL

' PROCESS COUNTER NUMERIC VALUE
' SET LOWER AFTER LISTVIEW WIDTH SET
'Label53_Here
' Label53.width = lstProcess_2_AND_lstProcess_3_WIDTH_POSTION_WITH_ER_OFFSET_LEFT


'----------------------------
'SET HIGHER UP A BIT
'Label_ProcessSetBar_Height = 380
'----------------------------
'MAXIMUM VALUE SET FIRST
'OKAY USE VARIABLE INSTEAD WHEN WANT FURTHER
'----------------------------
'Label51.Height = Label_ProcessSetBar_Height
'----------------------------

'MATCH WING DING OUTER BIGGER
'OUTER BIGGER
'----------------------------
Label52.Top = TOP_HEIGHT
Label52.height = Label_ProcessSetBar_Height '+ 10 '350 + 10 'Label51.Height + 50
Label52.width = Label_ProcessSetBar_Height
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
    Label46.width = 2200
    Label53.FontSize = 12
    Label46.FontSize = Label53.FontSize
    Label50.FontSize = Label53.FontSize
    'INNER WINGDING
    'Label51.FontSize = Label50.FontSize
    Label46.FontSize = Label50.FontSize
    '------------------------------------
End If

ADJUST_LEFT_OFFSET = 50
' LABEL51.CAPTION = IS THE WINGDING FRAME
'Label51_Here
Label51.Left = (Label46.Left + Label46.width) + ADJUST_LEFT_OFFSET

'OUTER BIGGER
Label52.Left = (Label46.Left + Label46.width) + 20
' THE BOX AROUND WINGDING
Label52.width = 420 'Label51.width + ADJUST_LEFT_OFFSET + 20
'WINGDING FONT SIZE
'----------------------------

'COUNTER
Label50.Left = (Label52.Left + Label52.width) + 20


TxtEXE.Top = TOP_HEIGHT
Label10.Top = TOP_HEIGHT

Label46.Top = TOP_HEIGHT
Label46.ToolTipText = "Hitt Me To Process Enumerate"
Label50.ToolTipText = "Hitt Me To Process Enumerate"
Label51.ToolTipText = "Hitt Me To Process Enumerate"
'Label53.ToolTipText = "Hitt Me To Process Enumerate"

'PROCESS COUNTING
'Label46_HERE
'Label46.Width _ DO IT UP HIGHER AT Label51.Left

Label46.height = Label52.height
'Label50_here
Label50.height = Label46.height

'PROCESS COUNTER LABEL INFO NAME
'-------------------------------
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
Label53.Top = Label50.Top + Label50.height + 20 ' + 20

'Label53_Here
Label53.height = 300

'Label51.Height = Label50.Height
'Label52.Height = Label51.Height
Label52 = ""

' LAB_MAXIMIZE_GOODSYNC.Left = Label53.Left - 20

lstProcess_2_ListView.Left = Label53.Left - 20 '+ 1 ' + 1

SETTING_WIDTH_LISTVIEW = False
If GetComputerName = "1-ASUS-X5DIJ" Then
    SETTING_WIDTH_LISTVIEW = True
    ' lstProcess_2_ListView.Font.Size = Default = 8.4
    lstProcess_2_ListView.Font.Size = 7.4
    lstProcess_3_SORTER_ListView.Font.Size = lstProcess_2_ListView.Font.Size
    'lstProcess_2_ListView_HERE
    lstProcess_2_ListView.width = 2100

End If

If GetComputerName = "2-ASUS-EEE" Then
    SETTING_WIDTH_LISTVIEW = True
    ' lstProcess_2_ListView.Font.Size = Default = 8.4
    lstProcess_2_ListView.Font.Size = 7.4
    lstProcess_3_SORTER_ListView.Font.Size = lstProcess_2_ListView.Font.Size
    'lstProcess_2_ListView_HERE
    lstProcess_2_ListView.width = 2100

End If

If SETTING_WIDTH_LISTVIEW = False Then
    SETTING_WIDTH_LISTVIEW = True
    ' lstProcess_2_ListView.Font.Size = Default = 8.4
    lstProcess_2_ListView.Font.Size = 8.4
    lstProcess_3_SORTER_ListView.Font.Size = lstProcess_2_ListView.Font.Size
    'lstProcess_2_ListView_HERE
    lstProcess_2_ListView.width = 3500
End If

lstProcess_3_SORTER_ListView.width = lstProcess_2_ListView.width
lstProcess_3_SORTER_ListView.Left = lstProcess_2_ListView.Left + lstProcess_2_ListView.width + 10

'-----------------------------------------
'lstProcess_2_ListView_HERE
'FINAL HEIGHT IS LATER DOWN A BIT
'-------------------------------------------------------
lstProcess_2_ListView.height = Label54.height + Label54.Top - 2000 '7040 - 810 '7040 '- 410
'lstProcess_2_ListView.Height = '(Label44.Top - lstProcess_2_ListView.Top)
lstProcess_3_SORTER_ListView.height = lstProcess_2_ListView.height

'TOP
'lstProcess_2_ListView_HERE
lstProcess_2_ListView.Top = Label53.Top + Label53.height + 10
'TOP
lstProcess_3_SORTER_ListView.Top = lstProcess_2_ListView.Top

'COMPUTER SPEC
'-------------
Label42.Top = TOP_HEIGHT
Label42.Left = lstProcess_3_SORTER_ListView.Left + lstProcess_3_SORTER_ListView.width + 50
Label42.width = 3500
Label42.height = HY 'xy 'lstProcess_3_SORTER_ListView.Height + 1000
Label42.FontSize = 7
'Label42_Anchor

'BOX WITH INFO ABOUT
Label44.Top = lstProcess_2_ListView.Top + lstProcess_2_ListView.height + 10
'Label44_HERE
Dim lstProcess_2_AND_lstProcess_3_WIDTH_POSTION_WITH_ER_OFFSET_LEFT
lstProcess_2_AND_lstProcess_3_WIDTH_POSTION_WITH_ER_OFFSET_LEFT = lstProcess_2_ListView.width + lstProcess_3_SORTER_ListView.width + (10 - 20) - 10

Label44.Left = Label46.Left + 5
Label44.width = lstProcess_2_AND_lstProcess_3_WIDTH_POSTION_WITH_ER_OFFSET_LEFT - 10

Label44.FontSize = 7
'Label44_HERE
Label44.height = 580 '+ 10 '+ 10
Label44.Caption = "LISTVIEW && CLICK_ER DONT LKE RUNNER IN THE IDE CRASH &WITH_ER  FOR NOW _ CODE PERFECT PROOF READ_ER __ Feeling to Do With ListView.OCX Version Fault Investigate _ How Annoying _ Code Inside IDE _ Check Error Outside IDE Compiled EXE"

'Label53_Here
Label53.width = lstProcess_2_AND_lstProcess_3_WIDTH_POSTION_WITH_ER_OFFSET_LEFT '- 10
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
Call MNU_ADMINISTRATOR_Click

'Call MNU_SHOW_THE_TIME_FORCE
'Call MNU_SHOW_THE_TIME_Click
Call PUT_GIVE_ME_TIME_IN_MENU



'Me.Left = 800
Call WIDTH_AND_HEIGHT(WX, HY)

'-------------------------------------------------------
'FINAL HEIGHT
'-------------------------------------------------------
lstProcess_2_ListView.height = HY - 700 - 100 '- Label44.height
lstProcess_3_SORTER_ListView.height = lstProcess_2_ListView.height
Label44.Top = lstProcess_2_ListView.Top + lstProcess_2_ListView.height + 10

Dim lstProcess_3_WIDTH_POSTION_LEFT_AND_WIDTH
lstProcess_3_WIDTH_POSTION_LEFT_AND_WIDTH = lstProcess_3_SORTER_ListView.Left + lstProcess_3_SORTER_ListView.width
Label50_Width_VAR = lstProcess_3_WIDTH_POSTION_LEFT_AND_WIDTH - (Label52.Left + Label52.width) - 30
Label50.width = Label50_Width_VAR



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

If Command$ <> "" Then Me.WindowState = vbMinimized

Call IsInternetConnected

MNU_KILL_MAX_AHK.Caption = "KILL AHK PID LIMITER > 100=TRUE"

Call MNU_GIVE_ME_TIME_WITHER_UTC_Click

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
If IsIDE = True Then Me.WindowState = vbNormal
If IsIDE = False Then Me.WindowState = vbMinimized
    
'__ Sub Timer_ALWAYS_ON_TOP_TO_START_WITH_ER_Timer()
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

Command_Screen_Shot_Auto_ClipBoard_er.Caption = "Screen Shot Auto ClipBoard_er when Spy_er && Archive Mode _OFF" + vbCrLf + "Hitt Button Here to Change"

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
'Label15.Caption = "Scan Processor Quicker For Some Moment and Also ForeGound Window HWND Handle Change Massive esponce Speed Up That Method __ __ __ __ Soon Next Process Logger With Full Path Name"
'Label15.Caption = "Scan Processor For A Moment Same As Other Command Set _ " + Trim(Str(EnumProcess_COUNTER))
'Label15.Caption = "Scan Processor For A Moment Timer Second _ " + Trim(Str(EnumProcess_COUNTER)) + " Plus ForeGound Window Change"
Call Timer_EnumProcess_Timer
'---------------------------------




End Sub


Private Sub Form_Click()
If IsIDE = True Then
     End
End If
End Sub

'IF AUTOREDRAW MAIN FORM _NOT ON_ WILL BE FAULT NOT HAPPENING HERE
'-----------------------------------------------------------------
'NOT QUITE
'-----------------------------------------------------------------
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

Private Sub Form_LostFocus()
'Me.AutoRedraw = False
'Me.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.AutoRedraw = False
'Me.SetFocus
End Sub

'This is an example of how you can use this routine:
'enable full row selecting
'SetFullRowSelection ListView1.hwnd, True
'disable full row selecting
'SetFullRowSelection ListView1.hwnd, False
    
    

Private Sub Form_Resize()

If NOT_RESIZE_EVENTER = True Then Exit Sub

If O_Me_WindowState = vbMaximized And Me.WindowState = vbNormal Then Exit Sub

If RESIZE_LOOP_STOP = True Then
    RESIZE_LOOP_STOP = False
    Exit Sub
End If

'--------SUB __ frmMain.EnumProcess IT IS IN HERE
If Me.WindowState <> vbMinimized Then
    Call EnumProcess
    Timer_EnumProcess.Interval = 1000
End If

'RESIZE
If CMD_Process_STATE_TO_SET_1ST = True Then
    'Call cmdProcess_Click
End If
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


'HEIGHT ADJUSTING _______________________________
OFFSET_HY = 500 + 340
'OFFSET_HY = 1000  '-2000
OFFSET_WX = 180

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

'------------------------------------------------

'WIN 10
If IsIDE = False Then
'    OFFSET_HY = OFFSET_HY + 140
'    OFFSET_WX = OFFSET_WX + 130
End If

'RESIZE HAS USE OF Menu_Height FUNCTION
'--------------------------------------
HY2 = HY + ((Menu_Height * Screen.TwipsPerPixelY) - Menu_Height) + OFFSET_HY
WX2 = WX + (Screen.TwipsPerPixelX) + OFFSET_WX

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
Dim ME_HEIGHT_EXTRA
ME_HEIGHT_EXTRA = 20 '50 + 10
Me.height = HY2 + ME_HEIGHT_EXTRA
RESIZE_LOOP_STOP = True
Me.width = WX2

'-------------------------------------------------------------------------
'THE MENU HAS TO BE LOADED WITH HEIGHT AND THEN REDONE AGAIN TWICE FOR X Y
Call WIDTH_AND_HEIGHT(WX, HY)
HY2 = HY + ((Menu_Height * Screen.TwipsPerPixelY) - Menu_Height) + OFFSET_HY
WX2 = WX + (Screen.TwipsPerPixelX) + OFFSET_WX

If Me.height + Me.width = HY2 + WX2 Then Exit Sub

On Error Resume Next
'CHANGE HEIGHT DONT RELOOP RESIZE
'--------------------------------
RESIZE_LOOP_STOP = True
'FORM_RESIZE
Me.height = HY2 + ME_HEIGHT_EXTRA

RESIZE_LOOP_STOP = True
Me.width = WX2 + 80

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

Exit Sub






Dim O_WindowState
If O_WindowState <> Me.WindowState Then
    If Me.WindowState = vbMinimized Then
        'MNU_CID_RELOADER.Visible = False

        Exit Sub
    End If
'    REM OKAY THIS CODE FOR HERE MAX AS CODE IS ONLY INNER FORM
'    If Me.WindowState = vbMaximized Then Exit Sub
End If

'MNU_CID_RELOADER.Visible = True
'TIMER_MNU_CID_RELOADER_DISAPPEAR.Enabled = True

O_WindowState = Me.WindowState

'List1.Top = FILE1.Top + FILE1.Height
'REMEMBER FROM FIRST TIME SET
'----------------------------
Dim VAR_BOTTOM_OBJECT
Dim Label_Tune__________
If VAR_BOTTOM_OBJECT = 0 Then
    VAR_BOTTOM_OBJECT = Label_Tune__________.height + Label_Tune__________.Top
End If
'
'
Dim VAR_WIDTH_OBJECT
VAR_WIDTH_OBJECT = Label_Tune__________.Left + Label_Tune__________.width
'If List1.Width > VAR_WIDTH_OBJECT Then VAR_WIDTH_OBJECT = List1.Width

'------------------------------------------------
'LOAD THE FORM TO SIZE OF INNER FORM AND MENU BAR
'------------------------------------------------
'SOMETIME DO OTHER WAY AROUND
'LOAD THE INNER CONTROLS TO SIZE OUTER FORM
'------------------------------------------------

Me.width = VAR_WIDTH_OBJECT + 100

'Menu_Height2 = Menu_Height
'Debug.Print Menu_Height
'Menu_Height2 = 21

'Me.Height = ((VAR_BOTTOM_OBJECT + (Menu_Height2 * Screen.TwipsPerPixelY)) + 410)
'
'If VARCENTER = True Then Exit Sub
'    'Me.Top = Screen.Height / 2 - Me.Height / 2 '+400
'    'Me.Left = Screen.Width / 2 - Me.Width / 2 '+400
'
'    VARCENTER = True
'End Sub
'
'---------------------------------------------------------------
'http://www.vbforums.com/showthread.php?673134-RESOLVED-Minimum-height-for-menu-bar-to-be-visible
'-------------- MENU SIZE DECLARE
'Private Type rect
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'Private Type MENUBARINFO
'  cbSize As Long
'  rcBar As rect
'  hMenu As Long
'  hwndMenu As Long
'  fBarFocused As Boolean
'  fFocused As Boolean
'End Type
'Private MenuInfo As MENUBARINFO
'Private Const OBJID_MENU As Long = &HFFFFFFFD
'Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
'Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hwnd As Long, _
'ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
'Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

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

Private Sub Label14_Click()

Call MNU_TASK_KILLER_EXPLORER_CLIPBOARD_Click

End Sub

Private Sub Lab_KILL_EXPLORER_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Lab_KILL_EXPLORER.BackColor = RGB(255, 255, 255)

Call MNU_TASK_KILLER_EXPLORER_CLIPBOARD_Click
End Sub

Private Sub LAB_MAXIMIZE_GOODSYNC2GO_Click()
' ---------------------
' ALIGN BY
' lstProcess_2_ListView.LEFT
' LAB_MAXIMIZE_GOODSYNC.LEFT
' ---------------------

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
LAB_MAXIMIZE_GOODSYNC2GO.BackColor = RGB(255, 255, 255)

Dim GOODSYNC_WINDOW_HWND
GOODSYNC_WINDOW_HWND = FindWindow("ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F00A}", vbNullString)

ShowWindow GOODSYNC_WINDOW_HWND, SW_MAXIMIZE
Beep
Me.WindowState = vbMinimized

Dim VAR, EXE_STRING
PID = -1
VAR = cProcesses.Convert(GOODSYNC_WINDOW_HWND, PID, cnFromhWnd Or cnToProcessID)
VAR = cProcesses.Convert(PID, EXE_STRING, cnFromProcessID Or cnToEXE)

PROCESS_TO_KILLER = EXE_STRING
PROCESS_TO_KILLER_PID = PID
Call LISTVIEW_CLICKER

End Sub

Private Sub LAB_MAXIMIZE_HUBIC_Click()
' ---------------------
' ALIGN BY
' lstProcess_2_ListView.LEFT
' LAB_MAXIMIZE_HUBIC.LEFT
' ---------------------

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
LAB_MAXIMIZE_HUBIC.BackColor = RGB(255, 255, 255)

Dim WINDOW_HWND
WINDOW_HWND = FindWindow("Class_PLMain", "Process Lasso Pro")

ShowWindow WINDOW_HWND, SW_MAXIMIZE
Beep
Me.WindowState = vbMinimized

Dim VAR, EXE_STRING
PID = -1
VAR = cProcesses.Convert(WINDOW_HWND, PID, cnFromhWnd Or cnToProcessID)
VAR = cProcesses.Convert(PID, EXE_STRING, cnFromProcessID Or cnToEXE)

PROCESS_TO_KILLER = EXE_STRING
PROCESS_TO_KILLER_PID = PID
Call LISTVIEW_CLICKER

End Sub

Private Sub LAB_MAXIMIZE_VB_KEEP_RUNNER_Click()
' ---------------------
' ALIGN BY
' lstProcess_2_ListView.LEFT
' LAB_MAXIMIZE_HUBIC.LEFT
' ---------------------

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
LAB_MAXIMIZE_VB_KEEP_RUNNER.BackColor = RGB(255, 255, 255)

Dim WINDOW_HWND
WINDOW_HWND = FindWindow("ThunderRT6FormDC", "VB KEEP RUNNER")
txthWnd.Text = Trim(Str(WINDOW_HWND))

Call cmdNormal_Click
' Call cmdMoveMax_Click

Beep
Me.WindowState = vbMinimized

Dim VAR, EXE_STRING
PID = -1
VAR = cProcesses.Convert(WINDOW_HWND, PID, cnFromhWnd Or cnToProcessID)
VAR = cProcesses.Convert(PID, EXE_STRING, cnFromProcessID Or cnToEXE)

PROCESS_TO_KILLER = EXE_STRING
PROCESS_TO_KILLER_PID = PID
Call LISTVIEW_CLICKER
End Sub

Private Sub Label_CLOSE_GOODSYNC_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_CLOSE_GOODSYNC.BackColor = RGB(255, 255, 255)

Call MNU_CLOSE_GOODSYNC_Click

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


Private Sub Label_CLOSE_HWND_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_CLOSE_HWND.BackColor = RGB(255, 255, 255)

Dim HWND_RESULT
HWND_RESULT = txthWnd
If HWND_RESULT > 0 Then
    Result = PostMessage(HWND_RESULT, WM_CLOSE, 0&, 0&)
End If
HWND_RESULT = txtParent
If HWND_RESULT > 0 Then
    Result = PostMessage(HWND_RESULT, WM_CLOSE, 0&, 0&)
End If
HWND_RESULT = txtParentX
If HWND_RESULT > 0 Then
    Result = PostMessage(HWND_RESULT, WM_CLOSE, 0&, 0&)
End If

Me.WindowState = vbMinimized
Beep

End Sub

Private Sub Label_KILL_HUBIC_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_KILL_HUBIC.BackColor = RGB(255, 255, 255)

'----
'How to select an external program's tray menu item? - Ask for Help - AutoHotkey Community
'https://autohotkey.com/board/topic/80327-how-to-select-an-external-programs-tray-menu-item/
'----
Dim HWND_RESULT
HWND_RESULT = FindWindow("HwndWrapper[hubiC.exe;;d22553e4-d0db-4e73-aa9f-48b38951eef0]", vbNullString)
If HWND_RESULT > 0 Then
    Result = PostMessage(HWND_RESULT, WM_CLOSE, 0&, 0&)
End If
' ---------------------------------------------------------------------
' CAN'T FIND UNLESS OPEN ON SCREEN
' AND CLOSE ONLY CLOSE THAT FORM NOT WHOLE THING
' UNLESS BY SLECT MENU OPTION
' IF USE AUTOHOTKEY - BUT STILL PROBLEM HAS OPEN MENU AND MOUSE CLICKER
' NOT GOOD IF TRAY ITEMS ARE BLOCK FROM APPEAR LIKE WHOLE TRAY FROZEN
' ---------------------------------------------------------------------

Dim VAR, EXE_STRING As String
PID = -1
VAR = cProcesses.Convert(HWND_RESULT, PID, cnFromhWnd Or cnToProcessID)
VAR = cProcesses.Convert(PID, EXE_STRING, cnFromProcessID Or cnToEXE)


' -------------------------------------------------------------------
' BOTH ABLE TO USE .Convert HAS USE .GetEXEID ANYWAY AS PART FUNCTION
' DEBUG IT REMEMBER TO GET STRING INPUT AOUND RIGHT WAY WHEN USE
' .Convert(EXE_STRING, PID, cnFromEXE Or cnToProcessID)
' -------------------------------------------------------------------
'PID = -1
'VAR = cProcesses.GetEXEID(PID, EXE_STRING)
EXE_STRING = "HubiC.exe"
PID = -1
VAR = cProcesses.Convert(EXE_STRING, PID, cnFromEXE Or cnToProcessID)

PROCESS_TO_KILLER = EXE_STRING
PROCESS_TO_KILLER_PID = PID
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

Sub COLOUR_BOX_SELECTOR_RESTORE_DEFAULT()
' CALL COLOUR_BOX_SELECTOR_RESTORE_DEFAULT

Dim ARRAY_CB()
ReDim ARRAY_CB(100)
Dim LDAC, R_COUNTER

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
ARRAY_CB(LDAC) = "Label_CLOSE_HWND"                      ' -- ELITESPY ONLY
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "LAB_MAXIMIZE_HUBIC"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "LAB_MAXIMIZE_GOODSYNC2GO"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = "LAB_MAXIMIZE_VB_KEEP_RUNNER"
LDAC = LDAC + 1
ARRAY_CB(LDAC) = ""
LDAC = LDAC + 1
ARRAY_CB(LDAC) = ""
LDAC = LDAC + 1
ARRAY_CB(LDAC) = ""

For R_COUNTER = 1 To LDAC
    If ARRAY_CB(R_COUNTER) = "" Then Exit For
Next

ReDim Preserve ARRAY_CB(R_COUNTER - 1)

Dim Control As Control
For Each Control In Me.Controls
    For R_COUNTER = 1 To UBound(ARRAY_CB)
        If Control.Name = ARRAY_CB(R_COUNTER) Then
            Control.BackColor = Label59.BackColor
            Control.ForeColor = RGB(0, 0, 0)
        End If
    Next
Next

Lab_KILL_EXPLORER.BackColor = RGB(255, 255, 255)

Dim HWND_RESULT
HWND_RESULT = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F009}", vbNullString)
If HWND_RESULT > 0 Then
    Result = PostMessage(HWND_RESULT, WM_CLOSE, 0&, 0&)
End If

End Sub


Private Sub Label65_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub lstProcess_2_ListView_KeyDown(KeyCode As Integer, Shift As Integer)
LISTVIEW_2_OR_3_HITT = 2
End Sub

Private Sub lstProcess_2_ListView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
LISTVIEW_2_OR_3_HITT = 2
End Sub

Private Sub lstProcess_3_SORTER_ListView_KeyDown(KeyCode As Integer, Shift As Integer)
LISTVIEW_2_OR_3_HITT = 3

End Sub

Private Sub lstProcess_3_SORTER_ListView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
LISTVIEW_2_OR_3_HITT = 3

End Sub

Private Sub mnu_Contact_Me_Click()
mnu_Telephone_BT_07722224555.Visible = True
MNU_MATT_LAN_BTINTERNET_COM.Visible = True
End Sub

Private Sub MNU_MATT_LAN_BTINTERNET_COM_Click()
MNU_MATT_LAN_BTINTERNET_COM.Visible = False
End Sub

Private Sub mnu_Telephone_BT_07722224555_Click()
mnu_Telephone_BT_07722224555.Visible = False
MNU_MATT_LAN_BTINTERNET_COM.Visible = False
End Sub



Private Sub TIMER_TO_RESIZE_Timer()
    TIMER_TO_RESIZE.Enabled = False
    'NOT_RESIZE_EVENTER = False
    Call Form_Resize
    'NOT_RESIZE_EVENTER = True

End Sub


Sub SET_MENU_PADD_WORK()

Dim i_Menu_Count, i_Form_Counter
Dim i_Menu_Not_Visa_Count

Dim Control As Control, LABEL_44, LABEL_48

Dim R_NEXT

Dim Text_Checker_Form_Menu As String

Dim MENU_ITEM_VAR

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
                    LABEL_44 = Trim(Control.Caption)
                    'LABEL_48 = Replace(LABEL_44, " ", "_")
                    LABEL_48 = LABEL_44
                    LABEL_48 = Replace(LABEL_48, "___", "__")
                    LABEL_48 = "[__ " + LABEL_48 + " __]"
                    LABEL_48 = Replace(LABEL_48, "[__ [__ ", "[__ ")
                    LABEL_48 = Replace(LABEL_48, " __] __]", " __]")
                    If LABEL_48 <> LABEL_44 Then
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
    
    If Val(txthWnd.Text) = 0 Then
        Dim GOODSYNC_WINDOW_HWND
        GOODSYNC_WINDOW_HWND = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F009}", vbNullString)
        txthWnd.Text = GOODSYNC_WINDOW_HWND
        If GOODSYNC_WINDOW_HWND = 0 Then
            ' MsgBox "GIVE txthWnd.Text SOME INPUT IS EMPTY"
            MsgBox "TxthWnd.Text IS EMPTY" + vbCrLf + "COMPUTER WILL GIVE IT ME.HWND " + Str(Me.hWnd)
            txtMhWnd.Text = Me.hWnd

        Else
            If Val(txthWnd.Text) = 0 Then
                MsgBox "TxthWnd.Text IS EMPTY" + vbCrLf + "COMPUTER WILL GIVE IT GOODSYNC " + txthWnd.Text
            End If
        End If
    End If
    
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
    
    ShowWindow txthWnd.Text, SW_NORMAL
    
    MoveWindow txthWnd.Text, HX, HY, HW, HH, True
    lHwnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdMaximize_Click()
    ' Maximize window
    Beep
    If Val(txthWnd.Text) = 0 Then
        Dim GOODSYNC_WINDOW_HWND
        GOODSYNC_WINDOW_HWND = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F009}", vbNullString)
        txthWnd.Text = GOODSYNC_WINDOW_HWND
        If GOODSYNC_WINDOW_HWND = 0 Then
            ' MsgBox "GIVE txthWnd.Text SOME INPUT IS EMPTY"
            MsgBox "TxthWnd.Text IS EMPTY" + vbCrLf + "COMPUTER WILL GIVE IT ME.HWND " + Str(Me.hWnd)
            txtMhWnd.Text = Me.hWnd
        Else
            MsgBox "TxthWnd.Text IS EMPTY" + vbCrLf + "COMPUTER WILL GIVE IT GOODSYNC " + txthWnd.Text
        End If
    End If
    
    ShowWindow txtMhWnd.Text, SW_MAXIMIZE

    lHwnd_Function_Button_Set_MIN_MAX = Val(txthWnd.Text)
    Call ChunkCodeOnMouse

End Sub
Private Sub cmdMinimize_Click()
    
    ' Minimize window
    Beep
    If Val(txthWnd.Text) = 0 Then
        Dim GOODSYNC_WINDOW_HWND
        GOODSYNC_WINDOW_HWND = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F009}", vbNullString)
        txthWnd.Text = GOODSYNC_WINDOW_HWND
        If GOODSYNC_WINDOW_HWND = 0 Then
            ' MsgBox "GIVE txthWnd.Text SOME INPUT IS EMPTY"
            MsgBox "TxthWnd.Text IS EMPTY" + vbCrLf + "COMPUTER WILL GIVE IT ME.HWND " + Str(Me.hWnd)
            txtMhWnd.Text = Me.hWnd
        Else
            MsgBox "TxthWnd.Text IS EMPTY" + vbCrLf + "COMPUTER WILL GIVE IT GOODSYNC " + txthWnd.Text
        End If
    End If
    
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
    
    If Val(txthWnd.Text) = 0 Then
        Dim GOODSYNC_WINDOW_HWND
        GOODSYNC_WINDOW_HWND = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F009}", vbNullString)
        txthWnd.Text = GOODSYNC_WINDOW_HWND
        If GOODSYNC_WINDOW_HWND = 0 Then
            ' MsgBox "GIVE txthWnd.Text SOME INPUT IS EMPTY"
            MsgBox "TxthWnd.Text IS EMPTY" + vbCrLf + "COMPUTER WILL GIVE IT ME.HWND " + Str(Me.hWnd)
            txtMhWnd.Text = Me.hWnd
        Else
            MsgBox "TxthWnd.Text IS EMPTY" + vbCrLf + "COMPUTER WILL GIVE IT GOODSYNC " + txthWnd.Text
        End If
    End If
   
   
    ' txtMhWnd.Text = GetParent(Val(txtMhWnd.Text))
    ' txtMhWnd.Text = GetParentHwnd(Val(txtMhWnd.Text))
    ' lhWndParentX = GetParentHwnd(Val(txtMhWnd.Text))
    ' txtMhWnd.Text = GetParentHwnd(Val(txtMhWnd.Text))

    ' txtMhWnd.Text = GetAncestor(Val(txtMhWnd.Text), GA_ROOT)

    Dim SET_HWND As Long

    If Val(txtMhWnd.Text) > 0 Then SET_HWND = Val(txtMhWnd.Text)
    If Val(txthWnd.Text) > 0 Then SET_HWND = Val(txthWnd.Text)
    
    ShowWindow SET_HWND, SW_NORMAL
        
    lHwnd_Function_Button_Set_MIN_MAX = SET_HWND
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

Private Sub cmdTerminateProcess_Click()
    ' Terminate process
    EnumProcess lstProcess.List(lstProcess.ListIndex)
End Sub

Private Sub cmdCopy_Click()
    Dim sText As String
    
    ' Setup window information's
    sText = sText & "Window Handle:     " & txthWnd.Text & vbCrLf
    sText = sText & "Window Handle HX:  " & txthWndHX.Text & vbCrLf
    sText = sText & "Window Caption:    " & txtTitle.Text & vbCrLf
    sText = sText & "Window Class:      " & txtClass.Text & vbCrLf
    sText = sText & "Window Style:      " & txtStyle.Text & vbCrLf
    sText = sText & "Rectangle:         " & txtRect.Text & vbCrLf
    sText = sText & "Parent Handle:     " & txtParent.Text & vbCrLf
    sText = sText & "Parent Handle HX:  " & txtParentHX.Text & vbCrLf
    sText = sText & "Parent Caption:    " & txtParentText.Text & vbCrLf
    sText = sText & "Parent Class:      " & txtParentClass.Text & vbCrLf
    
    Clipboard.Clear
    ' Copy text to clipboard
    Clipboard.SetText sText
End Sub

'Private Sub cmdMemInfo_Click()
'    ' Remove window from top
'    'SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'    'chkOnTop.Value = 0
'
'    'frmMemInfo.Show , Me
'    Load frmMemInfo
'End Sub

'Private Sub cmdAbout_Click()
'    ' Remove window from top
'    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
'    chkOnTop.Value = 0
'    ' Show about box
'    frmAbout.Show vbModal
'End Sub
Private Sub cmdClose_Click()
    ' Close program
    Unload Me
End Sub

Private Sub cmdCodeGeneration_Click()
    
    ' CMD LABEL = Show Code Generation Wizard
    ' NOW GONE LIMIT OF CONTROL COUNTER REACHED
    ' -----------------------------------------
    
    ' Remove window from top
    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    chkOnTop.Value = 0
    frmCGWizard.Show , Me
End Sub


Sub SET_LABEL_PADD_WORK()

Dim Control As Control

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


Sub WIDTH_AND_HEIGHT(WX, HY)

On Error GoTo ENDER

WX = 0: HY = 0
Dim II22, II23
Dim XYZ
On Local Error Resume Next

Dim Control As Control
For Each Control In frmMain.Controls
    Err.Clear
    II22 = True
    II23 = Control.width
'    If CONTROL.Name = "lstProcess" Then Stop
    
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
    
    'If Err.Number > 0 Then II22 = False
    
    If (II22 = True) And Err.Number = 0 Then
'        Debug.Print Control.Caption
    
        If Control.Left + Control.width > WX Then
            WX = Control.Left + Control.width
        End If
    
        If Control.Top + Control.height > HY Then
            HY = Control.Top + Control.height
            'Debug.Print Control.Name + " -- " + str(HY)
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

ENDER:

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

'Exit Sub
'
'COMMENT OUT HERE OR UNLOAD FROM HERE WILL UNLOAD THE CLASS CODE WHICH HAS AND UNLOAD TERMINATE
'    ----------------
'#########################################################
'# Class destructor to kill private variables
'#########################################################
'Private Sub Class_Terminate()

EXIT_TRUE = True

'End

End Sub

Private Sub MNU_LAUNCH_GRAMARLY_Click()
'C:\Users\MATT 01\AppData\Local\GrammarlyForWindows\app-1.5.25\GrammarlyForWindows.exe
End Sub


Private Sub Label30_Click()

Label57.Caption = "COMMAND LINE STATUS__ " + Label30

Label55.BackColor = Label59.BackColor
Label56.BackColor = Label58.BackColor

End Sub

Private Sub Label41_Click()
'Label41.
End Sub

Private Sub Label42_Click()
'Label42_Anchor
End Sub

Private Sub Label44_Click()
'Label44_HERE
End Sub

Private Sub Label46_Click()
'Label46_HERE
Call EnumProcess
End Sub

Private Sub Label48_Click()

TIMER2_TIMER_BEGAN = Now

End Sub

Private Sub Label50_Click()
'Label50_here
Call EnumProcess
End Sub

Private Sub Label51_Click()
'Label51_Here
Call EnumProcess
End Sub

Private Sub Label53_Click()
'Label53_Here
Call EnumProcess

Timer_Pause_Update.Enabled = True

Label53.BackColor = Label59.BackColor

End Sub

Private Sub Label54_Click()
'Label54.CAPTION
End Sub

Private Sub Label55_Click()

Label57.Caption = "COMMAND LINE STATUS__ " + Label30

Label55.BackColor = Label59.BackColor
Label56.BackColor = Label58.BackColor

End Sub

Private Sub Label56_Click()

Label57.Caption = "COMMAND LINE STATUS__ " + Label22

Label56.BackColor = Label59.BackColor
Label55.BackColor = Label58.BackColor

End Sub

Private Sub Label57_Click()
'Label57.CAPTION

End Sub

Private Sub Label60_Click()
'Label60_HERE
End Sub

Private Sub Label62_Click()
'NEW
Beep
MsgBox "NOT IN YET"
End Sub

Private Sub Label_KILL_WSCRIPT_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Label_KILL_WSCRIPT.BackColor = RGB(255, 255, 255)

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

Private Sub Label64_Click()

TIMER2_TIMER_BEGAN = Now + TimeSerial(0, 0, 20)

End Sub

Private Sub Label66_Click()
'SELECT FORCE MOST
Call Label29_Click

'SELECT PID MODE
Call Label22_Click

'SELECT EXECUTE KILLER
Call Label23_Click


End Sub

Private Sub LAB_MAXIMIZE_GOODSYNC_Click()
' ---------------------
' ALIGN BY
' lstProcess_2_ListView.LEFT
' LAB_MAXIMIZE_GOODSYNC.LEFT
' ---------------------

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
LAB_MAXIMIZE_GOODSYNC.BackColor = RGB(255, 255, 255)

Dim GOODSYNC_WINDOW_HWND
GOODSYNC_WINDOW_HWND = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F009}", vbNullString)

ShowWindow GOODSYNC_WINDOW_HWND, SW_MAXIMIZE
Beep
Me.WindowState = vbMinimized

Dim VAR, EXE_STRING
PID = -1
VAR = cProcesses.Convert(GOODSYNC_WINDOW_HWND, PID, cnFromhWnd Or cnToProcessID)
VAR = cProcesses.Convert(PID, EXE_STRING, cnFromProcessID Or cnToEXE)

PROCESS_TO_KILLER = EXE_STRING
PROCESS_TO_KILLER_PID = PID
Call LISTVIEW_CLICKER


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

Private Sub MNU_ADMIN_AUTO_START_RUN_AS_Click()
    'MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR ON"
    'APP.PATH+"\"+App.EXENamE+".EXE"
    Beep
    'If IsIDE = False Then
        If GetSpecialfolder(38) <> "" Then
            i = MsgBox("ALREADY IN ADMIN MODE __ RUN TEST ANYWAY", vbYesNoCancel + vbMsgBoxSetForeground)
            If i = vbNo Then Exit Sub
            
            If i = vbCancel And IsIDE = True Then
                Stop
            Else
                If i = vbCancel Then Exit Sub
            End If
        End If
        If GetSpecialfolder(38) = "" Or i = vbYes Then
        Call RUNAS_ME
'        Me.Show
'        DoEvents
'        MsgBox "RUNAS ME ADMIN", vbMsgBoxSetForeground
'        Sleep 4000
        End
        
        End If
    'End If
    
End Sub

Private Sub MNU_CAMERA_DATA_ON_SD_CARD_Click()
Beep
MNU_CAMERA_DATA_ON_SD_CARD.Caption = Replace(MNU_CAMERA_DATA_ON_SD_CARD.Caption, " CARD __", " CARD _ WORKER __")
Dim O_MeWindowState
O_MeWindowState = Me.WindowState
    
    If IsIDE = True Then Me.WindowState = vbMinimized
        Call CAMERA_FOLDER_TO_MATCH_WITH_DATE_FOR_WIFI_SDCF
    If IsIDE = True Then Me.WindowState = O_MeWindowState

MNU_CAMERA_DATA_ON_SD_CARD.Caption = Replace(MNU_CAMERA_DATA_ON_SD_CARD.Caption, " CARD _ WORKER __", " CARD __")
Beep
End Sub

Private Sub MNU_DELAY_FOLDER_TOP_Click()

Call MNU_LAUNCH_EXPLORER_STARTUP_DELAY_FOLDER_Click

End Sub

Private Sub MNU_FILE_LOCATOR_PRO_EXE_Click()


Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
Dim EXECUTE_FILE_NAME

'EXECUTE_FILE_NAME = "C:WINDOWS\EXPLORER.EXE " + EXECUTE_PARAM
EXECUTE_FILE_NAME = "C:\Program Files\Mythicsoft\FileLocator Pro\FileLocatorPro.exe"

WSHShell.Run """" + EXECUTE_FILE_NAME + """"

Set WSHShell = Nothing

End Sub

Private Sub MNU_FILELOCATOR_CLIPBOARD_IMAGER_Click()


'SHELL" D:\0 00 Art Loggers\# APP AND SCREEN\7-ASUS-GL522VW\FILE LOCATOR __ SavedCriteria.srf


Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
Dim EXECUTE_FILE_NAME

'EXECUTE_FILE_NAME = "C:WINDOWS\EXPLORER.EXE " + EXECUTE_PARAM
EXECUTE_FILE_NAME = "D:\0 00 Art Loggers\# APP AND SCREEN\7-ASUS-GL522VW\FILE LOCATOR __ SavedCriteria.srf"

WSHShell.Run """" + EXECUTE_FILE_NAME + """"

Set WSHShell = Nothing

Rem ---------------------------------------------------------
Rem EXECUTE_FILE_NAME = "C:WINDOWS\EXPLORER.EXE /SELECT, C:\"
Rem VBS CODE WILL NOT RUN FROM NOTEPAD++
Rem ---------------------------------------------------------

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

If InStr(i, "= NOT") Then
    MNU_GIVE_ME_TIME_WITHER_UTC.Caption = Replace(i, " NOT", " YES")
    Exit Sub
End If
If InStr(i, "= YES") Then
    MNU_GIVE_ME_TIME_WITHER_UTC.Caption = Replace(i, " YES", " NOT")
End If


End Sub

Private Sub MNU_KILL_EXPLORER_TOP_Click()

Call MNU_TASK_KILLER_EXPLORER_CLIPBOARD_Click

End Sub

Private Sub MNU_KILL_MAX_AHK_Click()
Beep

If InStr(MNU_KILL_MAX_AHK.Caption, "KILL AHK PID LIMITER > 100=TRUE") > 0 Then
    MNU_KILL_MAX_AHK.Caption = "KILL AHK PID LIMITER > 100=FALSE"
Else
    MNU_KILL_MAX_AHK.Caption = "KILL AHK PID LIMITER > 100=TRUE"
End If

End Sub

Private Sub MNU_KILL_NOT_RESPOND_TOP_Click()
    Call MNU_TASK_KILLER_NOT_RESPONDER_GENTAL_Click
End Sub

Private Sub MNU_LAUNCH_AUTORUNS_Click()
    Beep
    Me.WindowState = vbMinimized
    Dim PARAM, FILE_EXE_HERE
    FILE_EXE_HERE = "C:\PStart\Progs\#_PortableApps\PortableApps\AutorunsPortable\App\Autoruns\Autoruns64.exe"
    PARAM = ""
    Shell "CMD /C START """" /REALTIME /MAX """ + FILE_EXE_HERE + """", vbMaximizedFocus
    Me.WindowState = vbMinimized
End Sub


Private Sub MNU_LAUNCH_AUTORUNS_SET_BOOT_Click()
Dim objShell
Set objShell = CreateObject("Wscript.Shell")

objShell.Run """C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 21-AUTORUN.ahk""", 0, True
    
Set objShell = Nothing
Me.WindowState = vbMinimized
End Sub

Private Sub MNU_LAUNCH_AUTORUNS_SET_Click()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 28-AUTOHOTKEYS SET RELAUNCH CODE.ahk" + """"
Set WSHShell = Nothing
Me.WindowState = vbMinimized
End Sub

Private Sub MNU_LAUNCH_VB_SYNCRONIZER_Click()
    Dim RUN_EXE
    Dim objShell
    Set objShell = CreateObject("Wscript.Shell")
    RUN_EXE = "D:\VB6\VB-NT\00_Best_VB_01\10 SYNCRONIZE\SYNCRONIZER.exe"
    objShell.Run """" + RUN_EXE + """", 1, False
    Set objShell = Nothing
    Me.WindowState = vbMinimized
End Sub

Private Sub MNU_LAUNCH_BATCH_COMPILER_Click()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "D:\VB6\VB-NT\00_Best_VB_01\Batch_Compiler_Auto\BatchCompiler.exe" + """"
Set WSHShell = Nothing
Me.WindowState = vbMinimized
End Sub

Private Sub MNU_LAUNCH_Shell_VBasic_6_Loader_Click()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.Run """" + "D:\VB6\VB-NT\00_Best_VB_01\Shell VBasic 6 Loader\Shell VBasic 6 Loader.exe" + """"
Set WSHShell = Nothing
Me.WindowState = vbMinimized
End Sub

Private Sub MNU_LAUNCH_GRAMMARLY_Click()

Dim PARAM, FILE_EXE_HERE
Dim PART_NAME, OKAY_TO_GO

FILE_EXE_HERE = "C:\Users\MATT 01\AppData\Local\GrammarlyForWindows\app-1.5.25\GrammarlyForWindows.exe"
PART_NAME = "\AppData\Local\GrammarlyForWindows\app-1.5.25\GrammarlyForWindows.exe"
Beep

If Dir("C:\Users\" + GetUserName + "\" + PART_NAME) <> "" Then
    FILE_EXE_HERE = "C:\Users\" + GetUserName + "\" + PART_NAME
    Shell "CMD /C START """" /REALTIME /MAX """ + FILE_EXE_HERE + """", vbMaximizedFocus
    Beep
    Me.WindowState = vbMinimized
End If

Exit Sub

'MEHTOD ONE AND METHOD 02
Beep
Me.WindowState = vbMinimized
'Dim PARAM, FILE_EXE_HERE
'Dim PART_NAME, OKAY_TO_GO

Dir1.Path = "C:\Users"

FILE_EXE_HERE = "C:\Users\MATT 01\AppData\Local\GrammarlyForWindows\app-1.5.25\GrammarlyForWindows.exe"
PART_NAME = "\AppData\Local\GrammarlyForWindows\app-1.5.25\GrammarlyForWindows.exe"

For R_NEXT = 1 To Dir1.ListCount - 1

    If Dir(Dir1.List(R_NEXT) + PART_NAME) <> "" Then
        
        FILE_EXE_HERE = "C:\Users\" + Dir1.List(R_NEXT) + PART_NAME
        OKAY_TO_GO = True
    End If

Next R_NEXT

'GETUSERNAME

Beep

If OKAY_TO_GO = False Then Exit Sub

Shell "CMD /C START """" /REALTIME /MAX """ + FILE_EXE_HERE + """", vbMaximizedFocus

End Sub


Private Sub MNU_LAUNCH_CLIPBOARD_LOGGER_Click()
    Beep
    Me.WindowState = vbMinimized
    Dim PARAM, FILE_EXE_HERE
    FILE_EXE_HERE = "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.exe"
    PARAM = ""
    Shell "CMD /C START """" /D C:\ /REALTIME /MAX """ + FILE_EXE_HERE + """", vbMinimizedNoFocus
End Sub

Private Sub MNU_LAUNCH_CMD_Click()
    Beep
    Me.WindowState = vbMinimized
    Dim PARAM, FILE_EXE_HERE
    FILE_EXE_HERE = "CMD"
    PARAM = ""
    Shell "CMD /C START """" /D C:\ /REALTIME /MAX """ + FILE_EXE_HERE + """", vbMinimizedNoFocus
End Sub

Private Sub MNU_LAUNCH_EXPLORER_Click()
    Beep
    Me.WindowState = vbMinimized
    Dim PARAM, FILE_EXE_HERE
    FILE_EXE_HERE = "EXPLORER"
    PARAM = ""
    Shell "CMD /C START """" /REALTIME /MAX """ + FILE_EXE_HERE + """", vbMinimizedNoFocus
    Timer_EnumProcess = True
    Timer_EnumProcess.Interval = 1000

End Sub

Private Sub MNU_LAUNCH_EXPLORER_D__DSC_2015_2017_2017_Sony_CyberShot_HX60V_Click()
    Dim FOLDER_NAME_
    Beep
    Me.WindowState = vbMinimized
    Dim PARAM, FILE_EXE_HERE
    FILE_EXE_HERE = "EXPLORER"
    FOLDER_NAME_ = "D:\DSC\2015 2017\2017 Sony CyberShot HX60V\SavedCriteria____MP4.srf"
    PARAM = " /SELECT, """ + FOLDER_NAME_ + """"
    Shell "CMD /C START """" /REALTIME /MAX """ + FILE_EXE_HERE + """ " + PARAM, vbMinimizedNoFocus

End Sub

Private Sub MNU_LAUNCH_EXPLORER_SCRIPTOR_FOLDER_Click()
    Beep
    Me.WindowState = vbMinimized
    Dim PARAM, FILE_EXE_HERE
    FILE_EXE_HERE = "EXPLORER"
    PARAM = " /SELECT, ""C:\SCRIPTER\SCRIPTER CODE -- BAT"""
    Shell "CMD /C START """" /REALTIME /MAX """ + FILE_EXE_HERE + """ " + PARAM, vbMinimizedNoFocus

End Sub

Private Sub MNU_LAUNCH_EXPLORER_STARTUP_DELAY_FOLDER_Click()
    Beep
    Me.WindowState = vbMinimized
    Dim PARAM, FILE_EXE_HERE
    Dim DIR1_START_UP_DELAY_FOLDER_X
    FILE_EXE_HERE = "EXPLORER"
    
    Call STARTUP_DELAY_FOLDER_NAME
    
    File1.Path = DIR1_START_UP_DELAY_FOLDER
        
'    DIR1_START_UP_DELAY_FOLDER_X = Dir(DIR1_START_UP_DELAY_FOLDER + "\*.*", vbDirectory)
'    If DIR1_START_UP_DELAY_FOLDER_X = "" Then
'        DIR1_START_UP_DELAY_FOLDER_X = DIR1_START_UP_DELAY_FOLDER
'    Else
'        DIR1_START_UP_DELAY_FOLDER_X = DIR1_START_UP_DELAY_FOLDER + "\" + DIR1_START_UP_DELAY_FOLDER_X
'    End If

    If File1.ListCount > 0 Then
        PARAM = " /SELECT, """ + File1.Path + "\" + File1.List(0) + """"
    Else
        PARAM = " /SELECT, """ + File1.Path + """"
    End If
    
    Shell "CMD /C START """" /REALTIME /MAX """ + FILE_EXE_HERE + """ " + PARAM, vbMinimizedNoFocus

End Sub

Private Sub MNU_LAUNCH_INFO_RAPID_CLIPBOARD_LOGGER_TEXT_Click()

Beep
Me.WindowState = vbMinimized
Dim PARAM, FILE_EXE_HERE, EXECUTE_PARAM
Dim EXECUTE_FILE_NAME
'FILE_EXE_HERE = "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\7-ASUS-GL522VW\Data\INFO-RAPID----.BAT"
'Shell "CMD /C START """" /REALTIME /MAX """ + FILE_EXE_HERE + """ " + PARAM, vbMinimizedNoFocus


'D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\INFO RAPID Search Result INFO.se
'SHELL" D:\0 00 Art Loggers\# APP AND SCREEN\7-ASUS-GL522VW\FILE LOCATOR __ SavedCriteria.srf

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")

EXECUTE_FILE_NAME = "C:WINDOWS\EXPLORER.EXE " + EXECUTE_PARAM
EXECUTE_FILE_NAME = "D:\0 00 Art Loggers\# APP AND SCREEN\7-ASUS-GL522VW\FILE LOCATOR __ SavedCriteria.srf"

EXECUTE_FILE_NAME = "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\# DATA\INFO RAPID Search Result INFO.se"

WSHShell.Run """" + EXECUTE_FILE_NAME + """"

Set WSHShell = Nothing

Rem ---------------------------------------------------------
Rem EXECUTE_FILE_NAME = "C:WINDOWS\EXPLORER.EXE /SELECT, C:\"
Rem VBS CODE WILL NOT RUN FROM NOTEPAD++
Rem ---------------------------------------------------------

End Sub

Private Sub MNU_LAUNCH_KAT_MP3_RELOADER_Click()
    
    Beep
    Me.WindowState = vbMinimized
    Dim PARAM, FILE_EXE_HERE
    FILE_EXE_HERE = "D:\VB6\VB-NT\00_Best_VB_01\KAT MP3 RELOAD\KAT MP3 RELOAD.exe"
    Shell "CMD /C START """" /REALTIME /MAX """ + FILE_EXE_HERE + """ " + PARAM, vbMinimizedNoFocus

'
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


Private Sub MNU_LAUNCH_WINDOSE_Click()
    Beep
    Me.WindowState = vbMinimized
    Dim PARAM, FILE_EXE_HERE, FILE_EXE_HERE1, FILE_EXE_HERE2, EXE_STRING
    FILE_EXE_HERE = "C:\Program Files (x86)\Greatis\WinDowse\WinDowse.exe"
    
    If Dir(FILE_EXE_HERE) = "" Then FILE_EXE_HERE = ""
    
    If FILE_EXE_HERE = "" Then
        FILE_EXE_HERE = "C:\Program Files\Greatis\WinDowse\WinDowse.exe"
    End If
    
    If Dir(FILE_EXE_HERE) = "" Then
    FILE_EXE_HERE = ""
        FILE_EXE_HERE1 = "C:\Program Files (x86)\Greatis\WinDowse\WinDowse.exe"
        FILE_EXE_HERE2 = "C:\Program Files\Greatis\WinDowse\WinDowse.exe"
        EXE_STRING = vbCrLf + vbCrLf + FILE_EXE_HERE1 + vbCrLf + vbCrLf + FILE_EXE_HERE2
        MsgBox "FILE TO LAUNCH EXE __ DOES NOT EXIST HERE" + EXE_STRING, vbMsgBoxSetForeground
        Exit Sub
    End If
    
    PARAM = " /SELECT, ""C:\SCRIPTER\SCRIPTER CODE -- BAT"""
    Shell "CMD /C START """" /REALTIME """ + FILE_EXE_HERE + """ " + PARAM, vbMinimizedNoFocus

End Sub



Private Sub MNU_LOGG_OFF_CURRENT_USER_Click()
Dim FILE_EXE_HERE, PARAM
Me.WindowState = vbMinimized

FILE_EXE_HERE = "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 10-LOGG OFF __ TSDISCON.EXE -V.BAT"
Shell FILE_EXE_HERE, vbMinimizedNoFocus

'----
'Windows 7 & 8 Quick Tip: How To Switch User from a Command Line | Next of Windows
'http://www.nextofwindows.com/windows-7-8-quick-tip-how-to-switch-user-from-a-command-line
'----

Exit Sub
'--------------------------------------

'FILE_EXE_HERE = " ""C:\WINDOWS\SYSTEM32\tsdiscon.exe"" /V"
Shell "CMD /C START """" /REALTIME """ + FILE_EXE_HERE, vbMinimizedNoFocus
Shell "CMD /C START """" /REALTIME" + FILE_EXE_HERE, vbMinimizedNoFocus
'Shell "CMD /C START """" /REALTIME """ + FILE_EXE_HERE + """ " + PARAM, vbMinimizedNoFocus


'--------------------
'ALSO INFO ABOUT HERE
'--------------------
'C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\SysInternals\PsTools>psshutdown /?
'
'PsShutdown v2.52 - Shutdown, logoff and power manage local and remote systems
'Copyright (C) 1999-2006 Mark Russinovich
'sysinternals -www.sysinternals.com
'
'usage:
'psshutdown -s|-r|-h|-d|-k|-a|-l|-o [-f] [-c] [-t [nn|h:m]] [-v nn] [-e [u|p]:xx:yy] [-m "message"] [-u Username [-p password]]
'n s] [\\computer[,computer[,...]|@file]
'--------------------

End Sub

Private Sub MNU_LOG_BACK_IN_USER_ID_1_Click()
Dim FILE_EXE_HERE, PARAM
Me.WindowState = vbMinimized
'FILE_EXE_HERE = "C:\WINDOWS\SYSTEM32\tscon.exe 1 /DEST:Console /PASSWORD:"" "" /V"

FILE_EXE_HERE = "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 10-LOGG ON CONSOLE 1 __ TSCON.EXE.BAT"

Shell FILE_EXE_HERE, vbMinimizedNoFocus
Exit Sub
'--------------------------------------

Shell "CMD /C START """" /REALTIME """ + FILE_EXE_HERE, vbMinimizedNoFocus

End Sub

Private Sub MNU_LOG_BACK_IN_USER_ID_2_Click()
Dim FILE_EXE_HERE, PARAM
Me.WindowState = vbMinimized
'FILE_EXE_HERE = "C:\WINDOWS\SYSTEM32\tscon.exe 2 /DEST:Console /PASSWORD:"" "" /V"

FILE_EXE_HERE = "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 10-LOGG ON CONSOLE 2 __ TSCON.EXE.BAT"
Shell FILE_EXE_HERE, vbMinimizedNoFocus
Exit Sub
'--------------------------------------

Shell "CMD /C START """" /REALTIME """ + FILE_EXE_HERE, vbMinimizedNoFocus

End Sub

Private Sub MNU_MAXIMIZE_WINDOW_GOODSYNC_Click()
'GoodSync - C SCRIPTOR __ G7 TO 3L_0
'ahk_class {B26B00DA-2E5D-4CF2-83C5-911198C0F009}
'ahk_exe GoodSync - v10.EXE

Dim HWND_ID

HWND_ID = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F009}", vbNullString)

ShowWindow HWND_ID, SW_MAXIMIZE

End Sub

Private Sub MNU_MULTI_MENU_Click()
Dim FILE_EXE_HERE, PARAM
Me.WindowState = vbMinimized
FILE_EXE_HERE = "D:\VB6\VB-NT\00_Send_To\Send To Multi Menu\#0 Send To Multi Menu.exe"
Shell "CMD /C START """" /REALTIME """ + FILE_EXE_HERE + """ " + PARAM, vbMinimizedNoFocus

End Sub

Private Sub MNU_NUKER_Click()
Beep
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

Private Sub MNU_SHOW_STATUS_FORM_Click()

Messenger_Box.Show
Messenger_Box.WindowState = vbNormal
End Sub




Private Sub MNU_STARTER_FOLDER_Click()

Call MNU_LAUNCH_EXPLORER_STARTUP_DELAY_FOLDER_Click

End Sub


Private Sub MNU_TASK_KILLER_EXPLORER_CLIPBOARD_02_Click()
Call MNU_TASK_KILLER_EXPLORER_CLIPBOARD_Click
End Sub

Private Sub MNU_WINMERGE_ON_TOP_ALLTME_Click()
Beep

If InStr(MNU_WINMERGE_ON_TOP_ALLTME.Caption, "WINMERGE ON TOP ALLTIME=YES") > 0 Then
    MNU_WINMERGE_ON_TOP_ALLTME.Caption = "WINMERGE ON TOP ALLTIME=NOT"
Else
    MNU_WINMERGE_ON_TOP_ALLTME.Caption = "WINMERGE ON TOP ALLTIME=YES"
End If


End Sub


Private Sub Timer_01_VB_HELP_BOX_02_MSDN_Timer()

' ----------------------------------
' ALTERNATIVE LIKE HERE
' ----------------------------------
' Timer_01_VB_HELP_BOX_02_MSDN_Timer
' ---- OTHER -----------------------
' Timer_VB_MAXIMIZE
' ---- ALL THE VB ------------------
' ----------------------------------

'--------------------------------------------------------------------------------
'1..
'FIREFOX ADD AND SERACH ENGINE WITHOUT QUESTION ADD

'2..
'GOODSYNC
'GOOD_SYNC_MSGBOX_AT_END_IF_JOB_S_RUNNING_HWND = FindWindow("#32770", "GoodSync")

'3..
'ICACLS -- Error Applying Security
'--------------------------------------------------------------------------------

Dim FINDER, PWnd, GS_cWnd2, R_REPEAT

'FIREFOX ADD AND SERACH ENGINE WITHOUT QUESTION ADD
'--------------------------------------------------
FINDER = FindWindow("MozillaDialogClass", "Add Search Engine")
If FINDER > 0 And FINDER <> OHWnd_FINDER_1 Then

    '---------------------------------------
    'RESULT = PostMessage(FINDER, WM_CLOSE, 0&, 0&)
    '---------------------------------------
    'HWND_2 = GET_CHILD_TEST(MSDN_Lib)
    '---------------------------------------
    
'    cWnd = FindWindowEx(pwnd, 0, "BUTTON", "HELP")
'    EnableWindow cWnd, False

    PWnd = FINDER
    'GS_cWnd4 = 0
    'GS_cWnd1 = FindWindowEx(pwnd, 0, "Edit", vbNullString)
    'GS_cWnd2 = FindWindowEx(pwnd, 0, "Button4", "&OK")
    'GS_cWnd3 = FindWindowEx(pwnd, 0, "button", "&Cancel")
    
    GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
    
        
    If GS_cWnd2 > 0 Then
        '-------------------------------------------------
        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
        '-------------------------------------------------
        For R_REPEAT = 1 To 10
            GS_cWnd2 = FindWindowEx(PWnd, 0, vbNullString, "OK")
            If GS_cWnd2 = 0 Then Exit For
            SendMessage GS_cWnd2, BM_CLICK, 0, 0
        Next
    End If
    
End If

OHWnd_FINDER_1 = FINDER



FINDER = FindWindow(vbNullString, "RoboForm Upgrade")
If FINDER > 0 And FINDER <> OHWnd_FINDER_2 Then
    OHWnd_FINDER_2 = FINDER
    '---------------------------------------
    Result = PostMessage(FINDER, WM_CLOSE, 0&, 0&)
    '---------------------------------------
    'HWND_2 = GET_CHILD_TEST(MSDN_Lib)
    '---------------------------------------

End If




'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'GOODSYNC
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

Dim GOOD_SYNC_MSGBOX_CRASH_HWND As Long, I3, EXE_STRING, VAR

'GOODSYNC CRASH kill task not create report confidential and too large zip create
'and also hold up not continue happen restarter
'--------------------------------------------------------------------------------
'TESTER
'------
'GOOD_SYNC_MSGBOX_CRASH_HWND = FindWinPart("GoodSync2Go")
'--------------------------------------------------------------------------------
GOOD_SYNC_MSGBOX_CRASH_HWND = FindWindow("#32770", "Preparing Crash Report - GoodSync")
I3 = GOOD_SYNC_MSGBOX_CRASH_HWND
If I3 = 0 Then GOOD_SYNC_MSGBOX_CRASH_HWND = FindWindow("#32770", "GoodSync Crash")
I3 = GOOD_SYNC_MSGBOX_CRASH_HWND
If I3 > 0 And I3 <> OcWnd_GOOD_SYNC_CRASH Then
    
    'PID = -1
    'Var = cProcesses.GetEXEID(PID, "C:\Program Files\WinMerge\WinMergeU.exe")
    'Var = cProcesses.Convert(PID, OTSS, cnFromProcessID Or cnTohWnd)
    'Var = cProcesses.Convert(i3, PID, cnFromhWnd Or cnToProcessID)
    'Var = cProcesses.Convert(i3, PID, cnFromhWnd Or cnToProcessID)
    'Var = cProcesses.Convert(i3, PID, cnFromhWnd Or cnToProcessIDcnToEXE)
    'PID = -1
    'Var = cProcesses.Convert(i3, EXE_STRING, cnFromhWnd Or CNToEXE)
    
    '---------------------------------------------------------------
    '2017
    'LEACHED FROM ELITESPY TO GET PROPER FULL EXE NAME
    'INCLUDED IN CLASS ROUTINE TOGETHER
    'THIS IS NOT ENOUGH WHEN BE SHORT EXE NAME -- cnToEXE
    'MAYBE UPDATE WITH OTHER CODE NEARBY
    'NORMAL ROUTINE BUT ADD THE cProcesses. BECUASE IT IS IN A CLASS
    'AND CLASS MUST BE INITALISED -- AS IT IS DONE
    '---------------------------------------------------------------
    EXE_STRING = cProcesses.GetFileFromHwnd(I3)
    'Stop
    PID = -1
    VAR = cProcesses.Convert(I3, PID, cnFromhWnd Or cnToProcessID)
    
    'Stop
    
    '---------------------------------------------------------------------
    'WORKING BUT NOT FULL EXE NAME PATH
    'Var = cProcesses.Convert(PID, EXE_STRING, cnFromProcessID Or cnToEXE)
    'Var = cProcesses.Convert(i3, PID, cnFromhWnd Or cnToEXE)
    '---------------------------------------------------------------------
    'VAR = GetEXEID(PID, ByRef sRunningEXE As String) As Boolean
    '-----------------------------------------------------------
    
    
    VAR = cProcesses.Process_Kill(PID)
    
    Sleep 1000
    
    Shell EXE_STRING, vbMinimizedNoFocus
    
    MsgBox_11 = "GOODSYNC RESTARTED OF ELITE SPY CONTROL" + vbCrLf + vbCrLf + EXE_STRING ', vbMsgBoxSetForeground
    
    'Load Messenger_Box
    Messenger_Box.Show
    Messenger_Box.WindowState = vbNormal
    
    
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
    'TEXT1 = GetWindowText GS_cWnd1, S, l + 1
    'TEXT1 = GetWindowTitle(GS_cWnd1)
    
    '----------------------------------------------------------------
    'TO GET THE TEXT OF A BUTTON OR MSGBOX ORGINAL SOURCE CREDIT HERE
    '----
    'GET THE TEXT OF A BUTTON WITH HWND IN VB6 - Google Search
    'https://www.google.co.uk/search?q=GET+THE+TEXT+OF+A+BUTTON+WITH+HWND+IN+VB6&oq=GET+THE+TEXT+OF+A+BUTTON+WITH+HWND+IN+VB6+&aqs=chrome..69i57.9997j0j7&sourceid=chrome&ie=UTF-8
    '--------
    'How you get the hwnd's of a text box or button of another window?-VBForums
    'http://www.vbforums.com/showthread.php?576117-How-you-get-the-hwnd-s-of-a-text-box-or-button-of-another-window
    '----
    '----------------------------------------------------------------
    
'    Dim GS_cWnd1 As Long
'
'    pwnd = i3
'    GS_cWnd4 = 0
'    GS_cWnd1 = FindWindowEx(pwnd, 0, "Edit", vbNullString) '"One or more jobs are running now:")
'    TEXT1 = GetText(GS_cWnd1)
'    If InStr(TEXT1, "One or more jobs are running now:") > 0 Then GS_cWnd4 = 1
'
'    GS_cWnd2 = FindWindowEx(pwnd, 0, "Button", "&OK")
'    GS_cWnd3 = FindWindowEx(pwnd, 0, "button", "&Cancel")
    
    
    'ENABLE BUTTON = FALSE -- LEARN BEOFRE
    'EnableWindow cWnd, False
    
    
    '---------------------------------------------------------
    'Const BM_CLICK = &HF5&
    
    'CREDIT TO LEARN THE PUSH BUTTON WITH THESE 2 OR 3 COMMAND
    
    '----
    'POSTMESSAGE TO PRESS BUTTON ON A MSGBOX VB6 - Google Search
    'https://www.google.co.uk/search?q=POSTMESSAGE+TO+PRESS+BUTTON+ON+A+MSGBOX+VB6&oq=POSTMESSAGE+TO+PRESS+BUTTON+ON+A+MSGBOX+VB6&aqs=chrome..69i57.13373j0j7&sourceid=chrome&ie=UTF-8
    '--------
    'VBA/VB.Net/VB6–Click Open/Save/Cancel Button on IE Download window – PART I
    'http://www.siddharthrout.com/2011/10/23/vbavb-netvb6click-opensavecancel-button-on-ie-download-window/
    '----
    '---------------------------------------------------------
    
'    If GS_cWnd2 > 0 And GS_cWnd3 > 0 And GS_cWnd4 = 1 Then
'
'        '-------------------------------------------------
'        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
'        '1ST RUN DEBUG FOCUSED TO BUTTON BUT NOT HITT IT IN UNTIL 2ND
'        'FOCUS PROBLEM FIRST
'        '-------------------------------------------------
'        For R_REPEAT = 1 To 3
'            SendMessage GS_cWnd2, BM_CLICK, 0, 0
'        Next
'    End If
    
    '---------------------------------------
    '31 AUG 2K SIXTEEN
    'LEARN
    '---------------------------------------
    'Find Window Tutorial Part1 <A must see> by Paul Zaczkowski (from psc cd)
    'https://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37467&lngWId=1
    '---------------------------------------
End If

OcWnd_GOOD_SYNC_CRASH = I3


Dim GOOD_SYNC_MSGBOX_AT_END_IF_JOB_S_RUNNING_HWND, GS_cWnd4, Text1, GS_cWnd3

'GOODSYNC
GOOD_SYNC_MSGBOX_AT_END_IF_JOB_S_RUNNING_HWND = FindWindow("#32770", "GoodSync")
I3 = GOOD_SYNC_MSGBOX_AT_END_IF_JOB_S_RUNNING_HWND
If I3 >= 0 And I3 <> OcWnd_GOOD_SYNC Then

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
    'TEXT1 = GetWindowText GS_cWnd1, S, l + 1
    'TEXT1 = GetWindowTitle(GS_cWnd1)
    
    '----------------------------------------------------------------
    'TO GET THE TEXT OF A BUTTON OR MSGBOX ORGINAL SOURCE CREDIT HERE
    '----
    'GET THE TEXT OF A BUTTON WITH HWND IN VB6 - Google Search
    'https://www.google.co.uk/search?q=GET+THE+TEXT+OF+A+BUTTON+WITH+HWND+IN+VB6&oq=GET+THE+TEXT+OF+A+BUTTON+WITH+HWND+IN+VB6+&aqs=chrome..69i57.9997j0j7&sourceid=chrome&ie=UTF-8
    '--------
    'How you get the hwnd's of a text box or button of another window?-VBForums
    'http://www.vbforums.com/showthread.php?576117-How-you-get-the-hwnd-s-of-a-text-box-or-button-of-another-window
    '----
    '----------------------------------------------------------------
    
    Dim GS_cWnd1 As Long
    
    PWnd = I3
    GS_cWnd4 = 0
    GS_cWnd1 = FindWindowEx(PWnd, 0, "Edit", vbNullString) '"One or more jobs are running now:")
    Text1 = GetText(GS_cWnd1)
    If InStr(Text1, "One or more jobs are running now:") > 0 Then GS_cWnd4 = 1
    
    GS_cWnd2 = FindWindowEx(PWnd, 0, "Button", "&OK")
    GS_cWnd3 = FindWindowEx(PWnd, 0, "button", "&Cancel")
    
    
    'ENABLE BUTTON = FALSE -- LEARN BEOFRE
    'EnableWindow cWnd, False
    
    
    '---------------------------------------------------------
    'Const BM_CLICK = &HF5&
    
    'CREDIT TO LEARN THE PUSH BUTTON WITH THESE 2 OR 3 COMMAND
    
    '----
    'POSTMESSAGE TO PRESS BUTTON ON A MSGBOX VB6 - Google Search
    'https://www.google.co.uk/search?q=POSTMESSAGE+TO+PRESS+BUTTON+ON+A+MSGBOX+VB6&oq=POSTMESSAGE+TO+PRESS+BUTTON+ON+A+MSGBOX+VB6&aqs=chrome..69i57.13373j0j7&sourceid=chrome&ie=UTF-8
    '--------
    'VBA/VB.Net/VB6–Click Open/Save/Cancel Button on IE Download window – PART I
    'http://www.siddharthrout.com/2011/10/23/vbavb-netvb6click-opensavecancel-button-on-ie-download-window/
    '----
    '---------------------------------------------------------
    
    If GS_cWnd2 > 0 And GS_cWnd3 > 0 And GS_cWnd4 = 1 Then
        
        '-------------------------------------------------
        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
        '1ST RUN DEBUG FOCUSED TO BUTTON BUT NOT HITT IT IN UNTIL 2ND
        'FOCUS PROBLEM FIRST
        '-------------------------------------------------
        For R_REPEAT = 1 To 3
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

OcWnd_GOOD_SYNC = I3


Dim ICACLS_SETTER_PERMISSION_HWND


'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
' ICACLS -- Error Applying Security
'-------------------------------------------------------------------------------

Dim ICACLS_cWnd1 As Long, ICACLS_cWnd2 As Long, ICACLS_cWnd3 As Long, i3_2 As Long
ICACLS_SETTER_PERMISSION_HWND = FindWindow("#32770", "Error Applying Security")
i3_2 = ICACLS_SETTER_PERMISSION_HWND
If i3_2 >= 0 Then 'And i3_2 <> OcWnd_ICACLS_SETTER_PERMISSION_HWND Then
    'Sleep 2
    
    PWnd = i3_2
    'ICACLS_cWnd4 = 0
    'ICACLS_cWnd2 = FindWindowEx(pWnd, 0, "button", "&Continue")
    
'    TEXT1 = GetText(ICACLS_cWnd2)
'    ICACLS_cWnd2 = FindWindowEx(ICACLS_cWnd2, 0, "static", vbNullString) '"One or more jobs are running now:")
'    TEXT1 = GetText(ICACLS_cWnd2)
'    ICACLS_cWnd2 = FindWindowEx(ICACLS_cWnd2, 0, "static", vbNullString) '"One or more jobs are running now:")
'    TEXT1 = GetText(ICACLS_cWnd2)
'    ICACLS_cWnd2 = FindWindowEx(ICACLS_cWnd1, 0, "Static", vbNullString) '"One or more jobs are running now:")
'    TEXT1 = GetText(ICACLS_cWnd1)
'    ICACLS_cWnd2 = FindWindowEx(ICACLS_cWnd2, 0, "Static", vbNullString) '"One or more jobs are running now:")
'    TEXT1 = GetText(ICACLS_cWnd1)
    
    'If InStr(TEXT1, "One or more jobs are running now:") > 0 Then ICACLS_cWnd4 = 1
    
    'ICACLS_cWnd2 = FindWindowEx(pWnd, 0, "Button", "&OK")
    'ICACLS_cWnd3 = FindWindowEx(pWnd, 0, "button", "&Cancel")
    ICACLS_cWnd2 = FindWindowEx(PWnd, 0, "button", "&Continue")
    
    If ICACLS_cWnd2 > 0 Then
        
        '-------------------------------------------------
        'RESULT = PostMessage(i3, WM_CLOSE, 0&, 0&)
        '1ST RUN DEBUG FOCUSED TO BUTTON BUT NOT HITT IT IN UNTIL 2ND
        'FOCUS PROBLEM FIRST
        '-------------------------------------------------
        For R_REPEAT = 1 To 2
            ICACLS_cWnd2 = FindWindowEx(PWnd, 0, "button", "&Continue")
            If ICACLS_cWnd2 > 0 Then
                
                i = SendMessage(ICACLS_cWnd2, BM_CLICK, 0, 0)
                'I = SendMessageany(ICACLS_cWnd2, BM_CLICK, 0, 0)
                'I = PostMessage(ICACLS_cWnd2, BM_CLICK, 0&, 0&)
                'I = SetForegroundWindow(ICACLS_SETTER_PERMISSION_HWND)
                'If I > 0 Then
                    'Sleep 100
                'End If
                
                If I_Memmer < Now Then
                    Beep
                    I_Memmer = Now + TimeSerial(0, 0, 3)
                            
                End If
            DoEvents
            End If
            Next
'            ICACLS_cWnd2 = FindWindowEx(pwnd, 0, "button", "&Continue")
'            If ICACLS_cWnd2 > 0 Then
                'AppActivate "Error Applying Security", True
                'Sleep 3

            
                'cWnd = FindWindowEx(pwnd, 0, "button", "Cancel")
                'EnableWindow cWnd, False
                'EnableWindow ICACLS_cWnd2, False
                

                
            
'            End If

    End If
    
    '---------------------------------------
    '31 AUG 2K SIXTEEN
    'LEARN
    '---------------------------------------
    'Find Window Tutorial Part1 <A must see> by Paul Zaczkowski (from psc cd)
    'https://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37467&lngWId=1
    '---------------------------------------
End If

OcWnd_ICACLS_SETTER_PERMISSION_HWND = i3_2

End Sub


Public Function WindowText(ByVal WINDOW_HWND As Long) As String

    Dim txtlen              As Long

    If WINDOW_HWND = 0 Then
        Exit Function
    End If

    txtlen = SendMessage(WINDOW_HWND, WM_GETTEXTLENGTH, ByVal 0, ByVal 0)
    If txtlen = 0 Then
        Exit Function
    End If

    WindowText = String$(txtlen, vbNullChar)
    SendMessage WINDOW_HWND, WM_GETTEXT, ByVal (txtlen + 1), ByVal StrPtr(WindowText)

End Function



Private Sub Label10_Click()
Beep

Shell "EXPLORER /e, /SELECT, " + TxtEXE, vbNormalFocus

Me.WindowState = vbMinimized


'txtFile


End Sub


'Dim DataObject As New DataObject
'Dim file As String = "C:\Path"
'DataObject.SetData(DataFormats.FileDrop, True, file)
'Clipboard.SetDataObject (DataObject)
'
'
'
'End Sub
'
'Private Sub btnCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopy.Click
'    ' Make a DataObject.
'    Dim data_object As New DataObject
'
'    ' Add the data in various formats.
'    data_object.SetData(DataFormats.Rtf, rchSource.Rtf)
'    data_object.SetData(DataFormats.Text, rchSource.Text)
'
'    Dim html_text As String
'    html_text = "<HTML>" & vbCrLf
'    html_text &= "  <HEAD>The Quick Brown Fox</HEAD>" & _
'        vbCrLf
'    html_text &= "  <BODY>" & vbCrLf
'    html_text &= rchSource.Text & vbCrLf
'    html_text &= "  </BODY>" & vbCrLf & "</HTML>"
'    data_object.SetData(DataFormats.Html, html_text)
'
'    ' Place the data in the Clipboard.
'    Clipboard.SetDataObject (data_object)
'End Sub


Sub EXE_LOADER(AA)
    Beep
    If Dir(EXE_FILE_NAME) = "" Then
        If EXE_FILE_NAME = "" Then
            MsgBox "NOT FOUND __ PASSED STRING IS NOTHING EMPTY""""", vbMsgBoxSetForeground
            Exit Sub
        End If
        
        MsgBox "NOT FOUND" + vbCrLf + vbCrLf + EXE_FILE_NAME, vbMsgBoxSetForeground
        Exit Sub
    End If
    Shell EXE_FILE_NAME, vbMaximizedFocus
End Sub

Private Sub Label11_Click()

frmCopyFile2ClipBoard.txtFile = TxtEXE.Text

frmCopyFile2ClipBoard.cmdCopyToClipboard_Click

End Sub

Private Sub Label15_Click()

Timer_EnumProcess = True
Timer_EnumProcess.Interval = 1000

End Sub

Private Sub Label16_Click()

'PROCESS BUTTON_ER
'-----------------
'01 OF 09

'PROCESS_TO_KILLER
PROCESS_TO_KILLER_TO_GO = "/F /IM """ + Replace(UCase(PROCESS_TO_KILLER), ".EXE", "") + "*"""

Label22 = "TASKKILL " + "/F /IM """ + PROCESS_TO_KILLER_PID + """"
Label30 = "TASKKILL " + "/F /IM """ + Replace(UCase(PROCESS_TO_KILLER), ".EXE", "") + "*"""
Label57.Caption = "COMMAND LINE STATUS__ " + Label30



'Label22 = "TASKKILL " + "/F /IM """ + PROCESS_TO_KILLER_PID + """ /T"
'Label30 = "TASKKILL " + "/F /IM """ + PROCESS_TO_KILLER + """ /T"

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


Private Sub Label17_Click()
'PROCESS_TO_KILLER * /T

'PROCESS BUTTON_ER
'-----------------
'02 OF 09


PROCESS_TO_KILLER_TO_GO = "/IM """ + Replace(UCase(PROCESS_TO_KILLER), ".EXE", "") + "*"" /T"

Label22 = "TASKKILL " + "/IM """ + PROCESS_TO_KILLER_PID + """ /T"
Label30 = "TASKKILL " + "/IM """ + Replace(UCase(PROCESS_TO_KILLER), ".EXE", "") + "*"""
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
Label17.BackColor = Label47.BackColor

End Sub

Private Sub Label18_Click()

'PROCESS_TO_KILLER *
'-----------------
'03 OF 09


'PROCESS_TO_KILLER

'PROCESS_TO_KILLER = Replace(UCase(PROCESS_TO_KILLER), ".EXE", "")
PROCESS_TO_KILLER_TO_GO = "/IM """ + Replace(UCase(PROCESS_TO_KILLER), ".EXE", "") + "*"""

Label22 = "TASKKILL " + "/IM """ + PROCESS_TO_KILLER_PID + """ /T"
Label30 = "TASKKILL " + "/IM """ + Replace(UCase(PROCESS_TO_KILLER), ".EXE", "") + "*"""
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
Label18.BackColor = Label47.BackColor

End Sub

Private Sub Label19_Click()

'PROCESS_TO_KILLER /F /T
'-----------------
'04 OF 09

'PROCESS_TO_KILLER

PROCESS_TO_KILLER_TO_GO = "/F /IM """ + PROCESS_TO_KILLER + """ /T"

Label22 = "TASKKILL " + "/F /IM """ + PROCESS_TO_KILLER_PID + """ /T"
Label30 = "TASKKILL " + "/F /IM """ + PROCESS_TO_KILLER + """ /T"
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
Label28.BackColor = Label47.BackColor
Label19.BackColor = Label47.BackColor

End Sub

Private Sub Label28_Click()
'PROCESS_TO_KILLER /F *

'PROCESS BUTTON_ER
'-----------------
'05 OF 09

PROCESS_TO_KILLER_TO_GO = "/F /IM """ + PROCESS_TO_KILLER + """ /T"

Label22 = "TASKKILL " + "/F /IM """ + PROCESS_TO_KILLER_PID + """ /T"
Label30 = "TASKKILL " + "/F /IM """ + PROCESS_TO_KILLER + "*"""
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
Label28.BackColor = Label47.BackColor
Label19.BackColor = Label47.BackColor

End Sub

Private Sub Label20_Click()

'PROCESS_TO_KILLER /T

'PROCESS BUTTON_ER
'-----------------
'06 OF 09

PROCESS_TO_KILLER_TO_GO = "/IM """ + PROCESS_TO_KILLER + """ /T"

Label22 = "TASKKILL " + "/IM """ + PROCESS_TO_KILLER_PID + """ /T"
Label30 = "TASKKILL " + "/IM """ + PROCESS_TO_KILLER + """ /T"
Label57.Caption = "COMMAND LINE STATUS__ " + Label30

'Label22 = "TASKKILL " + "/IM """ + PROCESS_TO_KILLER_PID + """ /T"
'Label30 = "TASKKILL " + "/IM """ + PROCESS_TO_KILLER + """ /T"

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
Label20.BackColor = Label47.BackColor

End Sub

Private Sub Label21_Click()

'PROCESS_TO_KILLER _ EMPTY

'PROCESS BUTTON_ER
'-----------------
'07 OF 09

PROCESS_TO_KILLER_TO_GO = "/IM """ + PROCESS_TO_KILLER + """"

Label22 = "TASKKILL " + "/IM """ + PROCESS_TO_KILLER_PID + """"
Label30 = "TASKKILL " + "/IM """ + PROCESS_TO_KILLER + """"
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
Label21.BackColor = Label47.BackColor

End Sub

Private Sub Label26_Click()

'PROCESS_TO_KILLER /F
'-----------------
'08 OF 09

PROCESS_TO_KILLER_TO_GO = "/F /IM """ + PROCESS_TO_KILLER + """"

Label22 = "TASKKILL " + "/F /IM """ + PROCESS_TO_KILLER_PID + """"
Label30 = "TASKKILL " + "/F /IM """ + PROCESS_TO_KILLER + """"
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
Label26.BackColor = Label47.BackColor
Label27.BackColor = Label47.BackColor

End Sub

Private Sub Label27_Click()

'PROCESS_TO_KILLER /F
'PROCESS BUTTON_ER
'-----------------
'09 OF 09

PROCESS_TO_KILLER_TO_GO = "/F /IM """ + PROCESS_TO_KILLER + """"

Label22 = "TASKKILL " + "/F /IM """ + PROCESS_TO_KILLER_PID + """"
Label30 = "TASKKILL " + "/F /IM """ + PROCESS_TO_KILLER + """"
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
Label26.BackColor = Label47.BackColor
Label27.BackColor = Label47.BackColor

End Sub


Private Sub Label22_Click()

'Label22.CAPTION
'Clipboard.Clear
'Clipboard.SetText Label22

Label57.Caption = "COMMAND LINE STATUS__ " + Label22

Label56.BackColor = Label59.BackColor
Label55.BackColor = Label58.BackColor

End Sub



Private Sub TIMER_ColorCycle_TIMER()

'Dim WTrue_1, WTrue_2, WTrue_3, TW1, TW2, TW3
'Dim FIRST_RUN_COLOUR_CYCLE
Dim TIMER_INT
If Label23.BackColor <> &H8000000F Then TIMER_INT = 50 Else TIMER_INT = 10
If TIMER_ColorCycle.Interval <> TIMER_INT Then TIMER_ColorCycle.Interval = TIMER_INT

WTrue_1 = WTrue_1 + TW1

If FIRST_RUN_COLOUR_CYCLE = True Then
    FIRST_RUN_COLOUR_CYCLE = False
    If TW1 = 0 Then TW1 = 6
    If TW2 = 0 Then TW2 = 7
    If TW3 = 0 Then TW3 = 8
End If

If WTrue_1 > 255 Then TW1 = -6: WTrue_1 = WTrue_1 + TW1
If WTrue_1 < 1 Then TW1 = 6: WTrue_1 = WTrue_1 + TW1

WTrue_2 = WTrue_2 + TW2

If WTrue_2 > 255 Then TW2 = -7: WTrue_2 = WTrue_2 + TW2
If WTrue_2 < 1 Then TW2 = 7: WTrue_2 = WTrue_2 + TW2

WTrue_3 = WTrue_3 + TW3

If WTrue_3 > 255 Then TW3 = -8: WTrue_3 = WTrue_3 + TW3
If WTrue_3 < 1 Then TW3 = 8: WTrue_3 = WTrue_3 + TW3
   
   
Line1.BorderColor = RGB(WTrue_3, WTrue_2, WTrue_1)    ' Set drawing color.
Line2.BorderColor = RGB(WTrue_3, WTrue_2, WTrue_1)   ' Set drawing color.
Line3.BorderColor = RGB(255 - WTrue_3, 255 - WTrue_2, 255 - WTrue_1)
Line4.BorderColor = RGB(255 - WTrue_3, 255 - WTrue_2, 255 - WTrue_1)
   
   
If Label23.BackColor <> &H8000000F Then
    Label23.BackColor = RGB(WTrue_3, WTrue_2, WTrue_1)   ' Set drawing color.
    Label23.ForeColor = RGB(255 - WTrue_3, 255 - WTrue_2, 255 - WTrue_1)
End If

'frmPassLock.Label4.BackColor = RGB(WTrue_3, WTrue_2, WTrue_1)   ' Set drawing color.
'frmPassLock.Label4.ForeColor = RGB(255 - WTrue_3, 255 - WTrue_2, 255 - WTrue_1)
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

Private Sub Label24_Click()

'Private Sub MNU_HOOVER_20_SECOND_Click()
TIMER2_TIMER_BEGAN = Now

End Sub

Private Sub Label25_Click()

Dim PATH_NAME_AND_FILE_1, PATH_NAME_AND_FILE_2, PATH_NAME_AND_FILE

Beep
'TASKLIST
'Me.WindowState = vbMinimized

'FR1 = FreeFile
PATH_NAME_AND_FILE_1 = "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 04 TASKLIST.BAT"
'Open PATH_NAME_AND_FILE_1 For Output As #FR1
'    Print #FR1, "TASKLIST"
'Close #FR1
PATH_NAME_AND_FILE = PATH_NAME_AND_FILE_1

If Dir(PATH_NAME_AND_FILE_1) = "" Then
    PATH_NAME_AND_FILE_2 = App.Path + "BAT 04 TASKLIST.BAT"
    Open App.Path + "BAT 04 TASKLIST.BAT" For Output As #FR1
        Print #FR1, "TASKLIST"
    Close #FR1
    PATH_NAME_AND_FILE = PATH_NAME_AND_FILE_2
End If

If Dir(PATH_NAME_AND_FILE) <> "" Then
    Shell "CMD /C START """" /REALTIME /MAX """ + PATH_NAME_AND_FILE + """", vbMinimizedNoFocus
    Exit Sub
End If
If Dir(PATH_NAME_AND_FILE) = "" Then
    MsgBox "BATCH FILE CREATED DID NOT HAPPEN" + vbCrLf + "2 ATTEMPT" + vbCrLf + "1.." + vbCrLf + PATH_NAME_AND_FILE_2 + vbCrLf + vbCrLf + "AND HERE" + vbCrLf + "2.. " + vbCrLf + PATH_NAME_AND_FILE_2
End If

End Sub


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

'lstProcess_2_ListView.width
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

'PROCESS_TO_KILLER = lstProcess_3_SORTER_ListView.ListItems(lstProcess_3_SORTER_ListView.SelectedItem.Index).SubItems(1)
'PROCESS_TO_KILLER_PID = lstProcess_3_SORTER_ListView.ListItems(lstProcess_3_SORTER_ListView.SelectedItem.Index)
Call LISTVIEW_CLICKER

End Sub

Sub LISTVIEW_CLICKER()

Dim Success_Result
Dim Hwnd_Var
Dim PID_MARK
Dim PID_INPUT As Long
Dim NAME_EXE As String

If PROCESS_TO_KILLER = "" Then
    MsgBox "NOT ANY PROCESS TO KILLER NAME GIVEN -- EMPTY"
    Exit Sub
End If
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





Private Sub MNU_BRing_Front_Click()
' MNU_BRing_Front
' Mnu_Menu_Item_Count
' Mnu_Form_Count
'Call FindWinPartFront
'MNU_BRing_Front.Caption = "Bring All Front -- " + Format(Now, "DD-MMM-YYYY HH:MM:SS")

'GO_SUB = False
'For i = 1 To 255
'    bdf = GetAsyncKeyState(i)
'    If bdf < 0 Then
'        If i = 39 Then i = 0
'        If i = 116 Then i = 0 'F5 TEST ISIDE
'        If i = 16 Then GO_SUB = True 'LEFT SHIFT KEY
'        If GO_SUB = True Then
'            'Call KeyOrMouse
'            MNU_BRing_Front.Caption = "Bring All Front"
'            Exit Sub
'        End If
'    End If
'Next


'TEST=
'ExplorerGone = True
'ExplorerGone_TEST = True
Timer_EXPLORER_GONE.Enabled = True
'getkeyasync
End Sub



Private Sub Timer_ALWAYS_ON_TOP_TO_START_WITH_ER_Timer()
    
    If IsIDE = True And IsIDE_TEST = True Then Timer_ALWAYS_ON_TOP_TO_START_WITH_ER.Interval = 10000
    
    'Label60_HERE
    Label60.Caption = "Me on Top at Loader EXE Timer" + Str(counter_ALWAYS_ON_TOP_TIMER) + "_20 Second YES"
    
    counter_ALWAYS_ON_TOP_TIMER = counter_ALWAYS_ON_TOP_TIMER - 1
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    If counter_ALWAYS_ON_TOP_TIMER > 1 Then Exit Sub
    
    ' Put window on top of all others
    'SetWindowPos txtMhWnd.Text, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    'SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    ' Put window on top of all others
    'SetWindowPos txtMhWnd.Text, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    'SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    ' Remove window from top
    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Label60.Caption = "Me on Top @ Loader EXE Timer 20 Second NOT DONE"
    Label60.BackColor = Label59.BackColor '49 58_59
    Timer_ALWAYS_ON_TOP_TO_START_WITH_ER.Enabled = False
    'Label60.Caption = "Me on Top @ Loader EXE Timer 20 Second Already ON"

End Sub

Private Sub Timer_CONNECTED_TO_THE_INTERNET_Timer()

Call IsInternetConnected

End Sub

Private Sub Timer_EXPLORER_GONE_Timer()


'BRING WINDOWS FRONT
'-----------------------------------------------
i = FindWinPartExplorerGone(True) ' (True)(False) = Quite Mode Don't Display  Result
'-----------------------------------------------
'Debug.Print str(i) + " Windows Brought Forward"

'Text_Info = ""
'Text_Info = Text_Info + " Forward Form Explorer Crash@" + Format(Now, "DD-MMM HH:MM:SS")
'Text_Info = Text_Info + " Left Shift Reset Menu Opt"

MNU_BRing_Front.Caption = "Bring All Front = " + Str(i)

Timer_EXPLORER_GONE.Enabled = False
'ExplorerGone = False

Beep

'----------------------------------
Exit Sub
'----------------------------------



'If GETWinNT_Ver <> "WIN XP" Then
'    Timer_EXPLORER_GONE.Enabled = False
'    Exit Sub
'End If


'1st FIT IN ANOTHER SUBROUTINE
'Call MNU_CLIPBOARD_API_PUBLIC_VAR_HOOK_Click


'-------------------
'If Explorer Crashes
'-------------------

'----------------
'EXPLORER DESKTOP
'----------------
If FindWindow("Progman", "Program Manager") = 0 Then

'    ExplorerGone = True
    
    Me.WindowState = Normal
    DoEvents
    Me.Refresh
    DoEvents
    Me.SetFocus
    DoEvents
    Me.Refresh
    DoEvents
    
    

'    Me.WindowState = Normal

'    If i = vbYes Then
        'Shell "c:\windows\Explorer.exe", vbNormalFocus
        
        'cmdLine$ = "c:\windows\Explorer.exe"
        'i = ExecCmd(cmdLine$)
    
        Shell "c:\windows\Explorer.exe", vbNormalFocus
    
'        Do
'            i2 = FindWindow("Progman", "Program Manager")
'            DoEvents
'            'Sleep 500
'        Loop Until i2 <> 0
'        A = A
    'End If
    
    'ONLY REQUIRE WIN XP
    'FORM_STAY_UP_WITH_MSGBOX = True
    
    'i = MsgBox("Reload Explorer, Crash -- HAPPENED", vbMsgBoxSetForeground)
    
    
    'FORM_STAY_UP_WITH_MSGBOX = False
    
    Exit Sub

End If

'If FindWindow("Progman", "Program Manager") <> 0 And ExplorerGone = True Then

    'BRING WINDOWS FRONT
    'i = FindWinPartExplorerGone(False) ' True = Quite Mode Don't Display  Result
    i = FindWinPartExplorerGone(True) ' True = Quite Mode Don't Display  Result
'    Debug.Print str(i) + " Windows Brought Forward"
 
    
    'If ExplorerGone_TEST = True Then
    '    ExplorerGone_TEST = False
MNU_BRing_Front.Caption = "Bring " + Str(i) + " Forward Form Explorer Crash@" + Format(Now, "DD-MMM HH:MM:SS") + " Left Shift Reset Menu Opt"
    
    'Else
        'MNU_BRing_Front.Caption = "Bring " + str(i) + " Forward -- Explorer Crash/Terminated @ " + Format(Now, "DD-MMM-YYYY HH:MM:SS") + " Left Shift Menu Reset Button Menu"
    
    'End If
    Timer_EXPLORER_GONE.Enabled = False
    
'    ExplorerGone = False
    Beep
    
'End If



End Sub



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

Function FindWinPartExplorerGone(Q) As Long
'AND BRING ALL WINDOWS FORWARD
'ONLY VB SUFFERS FROM THIS

FindWinPartExplorerGone = False

Dim Rect8 As RECT

Dim ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one

Dim Huge

test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
'test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)
Dim BR$, HWnd9, GWS, EF, ECUTE

BR$ = ""
Do While test_hwnd <> 0

    DoEvents
    HWnd9 = GetWindowRect(test_hwnd, Rect8)
        ash$ = GetWindowTitle(test_hwnd)
'        If InStr(Ash$, "Double") > 0 Then Stop
        GWS = GetWindowState(test_hwnd)
'        If gws = -1 Then gws2$ = "Window State Normal"
'        If gws = vbMinimized Then gws2$ = "Window State Minimized"
'        If gws = vbMaximized Then gws2$ = "Window State Maximized"
'
'
    If (Rect8.Top > 0 And Rect8.Left > 0) Or GWS = vbMinimized Then
        ash$ = GetWindowTitle(test_hwnd)
        If ash$ <> "" And InStr(BR$, "-- " + ash$ + " -- ") = 0 Then
            BR$ = BR$ + "-- " + ash$ + " -- "
'            On Error Resume Next
'            AppActivate Ash$, 0
'            On Error GoTo 0
            Huge = Huge + 1
            EF = SetForegroundWindow(HWnd9)
            'ef = Putfocus(hwnd9)
            ECUTE = Timer + 0.1
         
            Do
            
                DoEvents
            
                'Safety IN CASE TME RESETS BACK TO ZERO
                If Timer < ECUTE - 30 Then Exit Do
                
            Loop Until Timer > ECUTE
        
        End If
    End If
        
'retrieve the next window
test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop

If Q = False Then MsgBox Str(Huge) + " Windows Brought To Front", vbMsgBoxSetForeground

FindWinPartExplorerGone = Huge

End Function



Private Sub MNU_EXE_AUTO_HOT_KEYS_SPY_Click()
    FILE_EXE = "C:\Program Files\AutoHotkey\AU3_Spy.exe"
    Call EXE_LOADER(FILE_EXE)
End Sub


Private Sub MNU_EXE_TASK_MANAGER_Click()

Shell "C:\Windows\System32\Taskmgr.exe", vbMaximizedFocus
Me.WindowState = vbMinimized

End Sub

Public Sub MNU_EXE_BOOT_KILLER_Click()

'FORM1.MNU_BOOT_KILLER_Click
Beep
Me.WindowState = vbMinimized
'Shell "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-BOOT KILLER.BAT", vbMaximizedFocus
Shell "CMD /C START """" /REALTIME /MAX ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-BOOT KILLER.BAT""", vbMinimizedNoFocus
Beep

End Sub

Private Sub MNU_EXE_BOOT_KILLER_NOTEPAD_Click()
FILE_EXE = "C:\Program Files (x86)\Notepad++\notepad++.exe ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT 01-BOOT KILLER.BAT"""
Call EXE_LOADER(FILE_EXE)
End Sub


Sub EXE_LAUNCHER_(i, EXE_ITEM)

If Dir(EXE_ITEM) = "" Then
    Beep
    MsgBox "ITEM NOT AVAILABLE ANYMORE" + vbCrLf + vbCrLf + EXE_ITEM, vbMsgBoxSetForeground
    Exit Sub
End If

Shell EXE_ITEM, vbMaximizedFocus
   
End Sub

Private Sub MNU_EXE_MENU_SCRIPT_FILL_02_Click(Index As Integer)
    EXE_ITEM = MNU_EXE_MENU_SCRIPT_FILL_02.Item(Index).Caption
    If Dir(EXE_ITEM) = "" Then
        MNU_EXE_MENU_SCRIPT_FILL_02.Item(Index).Caption = EXE_ITEM + "__NOT EXIST"
    End If
    Call EXE_LAUNCHER_(Index, EXE_ITEM)
End Sub

'Private Sub MNU_EXE_MENU_SCRIPT_FILL_03_Click(Index As Integer)
'    EXE_ITEM = MNU_EXE_MENU_SCRIPT_FILL_03.Item(Index).Caption
'    If Dir(EXE_ITEM) = "" Then
'        MNU_EXE_MENU_SCRIPT_FILL_02.Item(Index).Caption = EXE_ITEM + "__NOT EXIST"
'    End If
'    Call EXE_LAUNCHER_(Index, EXE_ITEM)
'End Sub

Private Sub MNU_EXE_NIRSOFT_Click()
    FILE_EXE = "C:\PStart\Progs\0_Nirsoft_Package\nirlauncher.exe"
    Call EXE_LOADER(FILE_EXE)
End Sub


Sub CHECK_EXE_MENU_FILE_EXIST_AND_CREATE_IF_ISNT_AND_READER()

FILE_TEXT = App.Path + "\# DATA\EXE_NOTEPADING_SCRIPT_MENU__" + GetComputerName + ".TXT"
On Error Resume Next

If Dir(App.Path + "\# DATA\", vbDirectory) = "" Then MkDir App.Path + "\# DATA\"

If Err.Number > 0 Then
    MsgBox "TRY AGAIN " + vbCrLf + vbCrLf + "MkDir " + App.Path + """\# DATA\""" + vbCrLf + vbCrLf + "MkDir App.Path + ""\# DATA\""" + vbCrLf + vbCrLf + "COMMAND FAIL WITH ERROR" + vbCrLf + Err.Description, vbMsgBoxSetForeground
    Beep
    Exit Sub
End If

FILENAME_PATH_EXE_MENU = FILE_TEXT
If IsIDE = True Then Kill FILENAME_PATH_EXE_MENU
'WRITE IN A COMMON AREA AND THEN GETCOMPUTERNAME
'WORK TO DO
'-----------------------------------------------

If Dir(FILENAME_PATH_EXE_MENU) = "" Then
    
    FR1 = FreeFile
    Open FILENAME_PATH_EXE_MENU For Output As #FR1
        Print #FR1, "; -------------------------------------------------------------"
        Print #FR1, "; TUE 14 FEB 2017 02:40 AM + 2 DAY ____ MATT.LAN@BTINTERNET.COM"
        Print #FR1, "; MENU SYSTEM LAUNCHER CREATION TIME 04:30 AM NOW"
        Print #FR1, "; -------------------------------------------------------------"
        Print #FR1, "; REM WITH __ ; __ THIS CHARACTER"
        Print #FR1, "; REM FOR ALLOW NOTE-ING"
        Print #FR1, "; MAXIMUM 20 LINE OF EXE FOR MENU CUSTOM ALLOW 1ST 20 READ "
        Print #FR1, "; REMMER WHEN NOT WANTING TEMPORALLY"
        Print #FR1, "; -------------------------------------------------------------"
        Print #FR1, "; EDIT FILE AND TIMER WILL PICK UP CHANGER AND THEN"
        Print #FR1, "; STOP CHECKING AFTER UNTIL LOAD TIME"
        Print #FR1, "; IF FILE HERE DELETER IS PUT BACK WITH DEFAULT"
        Print #FR1, "; MENU 01 IS A HARD CODER MENU"
        Print #FR1, "; -------------------------------------------------------------"
        Print #FR1, "; 01 TO 20 OF MENU 02 __ LEAVE THESE LINE MARKER"
        Print #FR1, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        Print #FR1, "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
        Print #FR1, "C:\Program Files (X86)\Kat MP3 Recorder\Kat MP3 Recorder.EXE"
        Print #FR1, "; -------------------------------------------------------------"
        Print #FR1, "; 01 TO 20 OF MENU 03 __ LEAVE THESE LINE MARKER"
        Print #FR1, "C:\Program Files (x86)\Siber Systems\AI RoboForm\Identities.exe"
        Print #FR1, "C:\Program Files\Siber Systems\GoodSync\GoodSync-v10.exe"
        Print #FR1, "C:\GoodSync\GoodSync2Go-v10.exe"
        Print #FR1, "; -------------------------------------------------------------"
        Close #FR1
    End If

    Call READ_EXE_MENU_FILE_AND_POPULATE_MENU

End Sub

Private Sub MNU_EXE_NOETPAD_Click()

Dim NOTEPAD_FILE_NAME
Dim NOTEPAD_OKAY

Call CHECK_EXE_MENU_FILE_EXIST_AND_CREATE_IF_ISNT_AND_READER
'-----------------------------------------------------------
If Dir(FILENAME_PATH_EXE_MENU) = "" Then Beep: Exit Sub

If NOTEPAD_PLUSPLUS_OR_NOTEPAD_NORMAL_RUN_WHEN_ERROR_LOAD_TIME_ONE_MSGBOX = True Then
    '--------------------------------
    'DIS-ENGAGE THIS ONE REST OF TIME
    '--------------------------------
    NOTEPAD_PLUSPLUS_OR_NOTEPAD_NORMAL_RUN_WHEN_ERROR_LOAD_TIME_ONE_MSGBOX = False
    
    NOTEPAD_FILE_NAME = "C:\Program Files (x86)\Notepad++\notepad++.exe"
    If Dir(NOTEPAD_FILE_NAME) = "" Then
        Beep
        MsgBox "NOT EXIST NOETPAD++ __ USER NORMAL NOTEPAD INSTEAD ___ MESSENGER ONLY ONE WHILE LOAD " + vbCrLf + vbCrLf + NOTEPAD_FILE_NAME, vbMsgBoxSetForeground
    Else
        NOTEPAD_OKAY = True
    End If

    If NOTEPAD_OKAY = False Then
        NOTEPAD_FILE_NAME = "C:\Windows\System32\notepad.exe"
        If Dir(NOTEPAD_FILE_NAME) = "" Then
            Beep
            MsgBox "NOT EXIST NORMAL NOTEPAD___ " + vbCrLf + vbCrLf + NOTEPAD_FILE_NAME, vbMsgBoxSetForeground
        End If
    End If
End If

If Dir(FILENAME_PATH_EXE_MENU) = "" Then
    Beep
    MsgBox "ELITESPY __ FILENAME __ HAD AN ERROR CREATING " + FILENAME_PATH_EXE_MENU, vbMsgBoxSetForeground
    NOTEPAD_PLUSPLUS_OR_NOTEPAD_NORMAL_RUN_WHEN_ERROR_LOAD_TIME_ONE_MSGBOX = False
    Exit Sub
End If

Beep
Shell NOTEPAD_FILE_NAME + " " + FILENAME_PATH_EXE_MENU, vbMaximizedFocus
Beep

Set F1 = FS.GetFile(FILENAME_PATH_EXE_MENU)
COMMAND_EXE_MENU_NOTEPAD_PICK_UP_DATE = F1.DateLastModified

Timer_COMMAND_EXE_MENU_NOTEPAD_PICK_UP.Enabled = True
End Sub

Private Sub MNU_EXE_PROCESS_EXPLORER_SYS_INT_PSTART_Click()
    EXE_FILE_NAME = "C:\PSTART\PROGS\#_PORTABLEAPPS\PORTABLEAPPS\PROCESSEXPLORERPORTABLE\APP\PROCESSEXPLORER\PROCEXP.EXE /t"
    Call EXE_LOADER(EXE_FILE_NAME)
End Sub

Private Sub MNU_EXE_WINDOWSE_Click()
Beep
Me.WindowState = vbMinimized
EXE_FILE_NAME = "C:\Program Files (x86)\Greatis\WinDowse\WinDowse.exe"
If Dir(EXE_FILE_NAME) = "" Then
    MsgBox "NOT FOUND" + vbCrLf + vbCrLf + EXE_FILE_NAME, vbMsgBoxSetForeground
    Exit Sub
End If
Shell EXE_FILE_NAME, vbMaximizedFocus
End Sub

Private Sub MNU_EXIT_Click()
EXIT_TRUE = True
Unload Me
'End
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




Private Sub MNU_HOOVER_20_SECOND_Click()
TIMER2_TIMER_BEGAN = Now
End Sub

Private Sub MNU_KILLER_A_VB_PROJECT_Click()

PROCESS_TO_KILLER_TO_GO = "/F /IM VB6.EXE"

'Label22 = "TASKKILL " + "/F /IM """ + PROCESS_TO_KILLER_PID + """ /T"
'Label30 = "TASKKILL " + "/F /IM """ + PROCESS_TO_KILLER + """ /T"

'Label22 = "TASKKILL " + PROCESS_TO_KILLER_PID
'Label30 = "TASKKILL " + PROCESS_TO_KILLER_TO_GO

'Label57 = "TASKKILL " + PROCESS_TO_KILLER_TO_GO



Label16.BackColor = Label22.BackColor
Label17.BackColor = Label22.BackColor
Label18.BackColor = Label22.BackColor
Label19.BackColor = Label22.BackColor
Label20.BackColor = Label22.BackColor
Label21.BackColor = Label22.BackColor

Label26.BackColor = Label11.BackColor

Label23.BackColor = Label11.BackColor

PROGRAM_PATH_BAT = "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT_03_PROCESS_KILLER.BAT"
'KILLER
Beep
'SHELL AND PARAM
'-----------------------------------------------------------------------
Shell PROGRAM_PATH_BAT + " " + PROCESS_TO_KILLER_TO_GO, vbMaximizedFocus

'RUNNER
Beep
If GetSpecialfolder(38) = "" Then
    MsgBox "YOU HAVE TO SWITCH ADMIN MODE ON" + vbCrLf + vbCrLf + "THE COMMAND GetSpecialfolder(38) REQUIRE ADMINISTRATOR" + vbCrLf + vbCrLf + "RESULT RETURN NOTHING HERE GetSpecialfolder(38)__<" + GetSpecialfolder(38) + ">__ RESULT RETURNED NOTHING", vbMsgBoxSetForeground
    Exit Sub
End If

CODER_VBP_FILE_NAME_2 = "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.vbp"

Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + CODER_VBP_FILE_NAME_2 + """", vbNormalFocus

End Sub


Sub STARTUP_DELAY_FOLDER_NAME()

    Dim SPF_i

    SPF_i = DIR1_START_UP_DELAY_FOLDER
    '------------------------------------------------------
    'DIR1 = "E:\01 Start Menu\Programs\Startup-01-Delayed\"
    'USER FOLDER
    '07 START MENU PROGRAMS STARTUP
    '------------------------------------------------------
    SPF_i = GetSpecialfolder(7) + "-01-Delayed"
        '------------------------------------------------------
    'CHANGED TO
    'PROGRAM DATA
    '24 COMMON
    '------------------------------------------------------

    If InStr(UCase(SPF_i), UCase("\PROGRAMS\START")) = 0 Then
        'boot in too quick from task schedular
        'may restart few time"
        '-------------------------------
        'ERROR NOT WORKING AT THE MOMENT
        'Timer_Restart.Enabled = True
        'Exit Sub
        '-------------------------------
    End If

    On Error Resume Next
    If Dir(SPF_i, vbDirectory) = "" Then
        ChDrive SPF_i
        CreateFolderTree SPF_i
    End If
    
    If Dir(SPF_i, vbDirectory) = "" Then
        MsgBox "ERROR MAKE DIR " + vbCrLf + SPF_i: End
    End If
    
    SPF_i = SPF_i + "\" + GetComputerName + "--" + GetUserName + ""
    
    If Dir(SPF_i, vbDirectory) = "" Then
        ChDrive SPF_i
        CreateFolderTree SPF_i
    End If
    '------------------------------------------------------
    '------------------------------------------------------

    DIR1_START_UP_DELAY_FOLDER = SPF_i

End Sub

Private Sub MNU_LAUNCH_HUBIC_Click()
    Beep
    Me.WindowState = vbMinimized
    Dim PARAM, FILE_EXE_HERE
    FILE_EXE_HERE = "C:\Program Files\OVH\hubiC\hubiC.exe"
    PARAM = "run --showsync"
    Shell "CMD /C START """" /BELOWNORMAL """ + FILE_EXE_HERE + """", vbMinimizedNoFocus

End Sub

Private Sub MNU_LAUNCH_NIRSOFT_Click()
    Beep
    Me.WindowState = vbMinimized
    Dim PARAM, FILE_EXE_HERE
    FILE_EXE_HERE = "C:\PStart\Progs\0_Nirsoft_Package\nirlauncher.exe"
    PARAM = ""
    Shell "CMD /C START """" /REALTIME /MAX """ + FILE_EXE_HERE + """", vbMinimizedNoFocus
End Sub

Private Sub MNU_LAUNCH_PROCESS_EXPLORER_Click()
    Beep
    Me.WindowState = vbMinimized
    Dim PARAM, FILE_EXE_HERE
    FILE_EXE_HERE = "C:\PStart\Progs\#_PortableApps\PortableApps\ProcessExplorerPortable\App\ProcessExplorer\procexp.exe"
    PARAM = ""
    Shell "CMD /C START """" /REALTIME /MAX """ + FILE_EXE_HERE + """", vbMinimizedNoFocus
End Sub


Private Sub MNU_LAUNCH_VB_Click()
    Beep
    Me.WindowState = vbMinimized
    Dim PARAM, FILE_EXE_HERE
    PARAM = ""
'    Shell "CMD /C START """" /D C:\ /REALTIME /MAX """ + FILE_EXE_HERE + """", vbMaximizedFocus

    Beep
    If GetSpecialfolder(38) = "" Then
        MsgBox "YOU HAVE TO SWITCH ADMIN MODE ON" + vbCrLf + vbCrLf + "THE COMMAND GetSpecialfolder(38) REQUIRE ADMINISTRATOR" + vbCrLf + vbCrLf + "RESULT RETURN NOTHING HERE GetSpecialfolder(38)__<" + GetSpecialfolder(38) + ">__ RESULT RETURNED NOTHING", vbMsgBoxSetForeground
        Exit Sub
    End If
    
    'CODER_VBP_FILE_NAME_2 = "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.vbp"
    
    Dim PROGRAM_NAME
    PROGRAM_NAME = GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"
    
    Shell "CMD /C START """" /REALTIME /MAX """ + PROGRAM_NAME + """", vbMinimizedNoFocus

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


Private Sub Mnu_Menu_Item_Count_Click()
'Mnu_Menu_Item_Count

End Sub

Private Sub MNU_REBOOT_OPTION_Click()

If MNU_REBOOT_OPTION.Caption = "RESTART_NOT_OPTION_ENABLED" Then
    MNU_REBOOT_OPTION.Caption = "RESTART_YES_OPTION_ENABLED"
Else
    MNU_REBOOT_OPTION.Caption = "RESTART_NOT_OPTION_ENABLED"
End If

End Sub

Private Sub MNU_RESTART_Click()
If MNU_REBOOT_OPTION.Caption <> "RESTART_YES_OPTION_ENABLED" Then Exit Sub
    
Shell "CMD /C START """" /REALTIME /MAX ""SHUTDOWN"" -R -T 0""", vbNormalFocus

'    /t xxx     Set the time-out period before shutdown to xxx seconds.
'               The valid range is 0-315360000 (10 years), with a default of 30.
'               If the timeout period is greater than 0, the /f parameter is
'               implied.

End Sub

Private Sub MNU_SET_TIME_ON_DEVICE_SET_THAT_WANT_IT_HardCoder_Click()

Dim A_DATE_TIME As String, PATH_VOLUME_NAME, R_FOR
Dim A_DATE_TIME_TEMP As String

A_DATE_TIME = Format(Now, "YYYY-MM-DD-HH-MM-SS")

For R_FOR = 5 To 26
    PATH_VOLUME_NAME = ""
    PATH_VOLUME_NAME = Dir(Chr(R_FOR + 64) + ":", vbVolume)
    If PATH_VOLUME_NAME = "USB_MP3" Then Exit For
    If PATH_VOLUME_NAME = "NO NAME" Then Exit For
Next
PATH_VOLUME_NAME = Chr(R_FOR + 64) + ":"
If Dir(PATH_VOLUME_NAME + "\SetTime.txt") <> "" Then
    Reset
    FR1 = FreeFile
    Open Dir(PATH_VOLUME_NAME + "\SetTime.txt") For Binary As #FR1
        Put #FR1, , A_DATE_TIME
    Close #FR1
    
    FR1 = FreeFile
    Open Dir(PATH_VOLUME_NAME + "\SetTime.txt") For Input As #FR1
        Line Input #FR1, A_DATE_TIME_TEMP
    Close #FR1
    
    If A_DATE_TIME_TEMP <> A_DATE_TIME Then
    MsgBox "FILE NOT THE SAME"
    Stop
    End If
End If

Beep
Me.WindowState = vbMinimized

End Sub

Sub MNU_SHOW_THE_TIME_FORCE()

Exit Sub

Dim PATH_VOLUME_NAME, R_FOR
Dim A_DATE_TIME_TEMP As String
Dim A_DATE_TIME_PM

A_DATE_TIME_PM = Mid(Format(Now, "HH-MM-SS AM/PM"), 10)
A_DATE_TIME = "[__ " + Format(Now, "DDD MMM YYYY-MM-DD-HH-MM-SS ") + A_DATE_TIME_PM + " The United Kingdom __]"

MNU_SHOW_THE_TIME.Caption = A_DATE_TIME

End Sub

Sub PUT_GIVE_ME_TIME_IN_MENU()

Exit Sub

Dim PATH_VOLUME_NAME, R_FOR
Dim A_DATE_TIME_TEMP As String
Dim A_DATE_TIME_PM

A_DATE_TIME_PM = Mid(Format(Now, "HH-MM-SS Am/Pm"), 10)
'A_DATE_TIME = "[__ " + Format(Now, "YYYY-MM-DD-HH-MM-SS ") + A_DATE_TIME_PM + " __]"
A_DATE_TIME = "[__ " + Format(Now, "DDD MMM YYYY-MM-DD-HH-MM-SS ") + A_DATE_TIME_PM + " __]"

'SHOW THE TIME A MOMENT OR MENU FLICKER EVERY SECOND
'MNU_SHOW_THE_TIME.Caption = "SHOW THE TIME __ " + A_DATE_TIME + " __ A MOMENT OR MENU FLICKER EVERY SECOND"
'AT FORM LOAD
'------------

MNU_SHOW_THE_TIME.Caption = A_DATE_TIME

'Call SET_MENU_PADD_WORK

'SHOW_THE_TIME.Enabled = True

End Sub

Private Sub MNU_SHOW_THE_TIME_Click()

'Timer_SHOW_THE_TIME_Timer()
'---------------------------
'MNU_GIVE_ME_TIME_Click()
'---------------------------

Call MNU_GIVE_ME_TIME_Click

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

Private Sub MNU_TASK_KILLER_CHROME_Click()

Dim PROGRAM_PATH_BAT
PROGRAM_PATH_BAT = App.Path + "\BAT_03_PROCESS_KILLER.BAT"

PROCESS_TO_KILLER_TO_GO = "/F /IM CHROME* /T"
PROCESS_TO_KILLER = PROCESS_TO_KILLER_TO_GO

Label23.BackColor = Label11.BackColor

Beep
Me.WindowState = vbMinimized
Call CREATE_PROGRAM_PATH_BAT

Shell "CMD /C START """" /REALTIME /MIN " + PROGRAM_PATH_BAT + " " + PROCESS_TO_KILLER_TO_GO, vbMinimizedNoFocus


End Sub

Private Sub MNU_TASK_KILLER_CMD_Click()
    
PROCESS_TO_KILLER_TO_GO = "/F /IM CMD* /T"
PROCESS_TO_KILLER = PROCESS_TO_KILLER_TO_GO

Label23.BackColor = Label11.BackColor
    
Beep
Me.WindowState = vbMinimized
Shell "CMD /C START """" /REALTIME ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT_03_PROCESS_KILLER.BAT"" " + PROCESS_TO_KILLER, vbMaximizedFocus
Beep
    
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

Private Sub MNU_TASK_KILLER_NOT_RESPONDER_Click()

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
Private Sub MNU_VB_FOLDER_Click()
    Beep
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + CODER_EXE_FILE_NAME_1 + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
End Sub
Private Sub MNU_ADMINISTRATOR_Click()
    'Call MNU_ADMINISTRATOR_Click
    MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "TOR ON", "TOR OFF")
    End If
    MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "ADMINISTRATOR", "ADMINISTER Priviledge Elevation")

End Sub
'-------------------------------------------
'-------------------------------------------
'Private Type SHITEMID
'    cb As Long
'    abID As Byte
'End Type
'Private Type ITEMIDLIST
'    mkid As SHITEMID
'End Type
'Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
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


'////////////////////////////////////////////////////////////////////
'//// CROSSHAIR EVENTS
'////////////////////////////////////////////////////////////////////
Private Sub picCrossHair_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    
    If tPA.Y = 0 Or tPA.Y < (Me.Top / Screen.TwipsPerPixelY) Then
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


Private Sub picCrossHair_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    picCrossHair_MouseMove_02
    
    'Exit Sub
    ' If user pressed left mouse button and we are dragging
    
'    If Button = vbLeftButton And m_bDragging Then
'    End If

End Sub

Private Sub picCrossHair_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If user pressed left mouse button and we are dragging
    If Button = vbLeftButton And m_bDragging Then
        ' Set dragging flag to true
        m_bDragging = False
        picCrossHair_MouseMove_Dragging_VAR = False

        ' Restore mouse pointer to normal (arrow)
        Me.MousePointer = vbNormal
        ' Load picture into picCrossHair
        picCrossHair.Picture = imgCursor.MouseIcon
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

' Make textboxes flat
Private Sub MakeFlat(lHwnd As Long)
    Dim lStyle As Long
    
    ' Get window style
    lStyle = GetWindowLong(lHwnd, GWL_EXSTYLE)
    ' Setup window styles
    lStyle = lStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    ' Set window style
    SetWindowLong lHwnd, GWL_EXSTYLE, lStyle
    RemoveBorder lHwnd
End Sub
Private Sub RemoveBorder(lHwnd As Long)
    Dim lStyle As Long
    
    ' Get window style
    lStyle = GetWindowLong(lHwnd, GWL_STYLE)
    ' Setup window styles
    lStyle = lStyle And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION Or WS_BORDER Or WS_SIZEBOX Or WS_THICKFRAME)
    ' Set window style
    SetWindowLong lHwnd, GWL_STYLE, lStyle
    ' Update window
    SetWindowPos lHwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
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


Private Sub SHOW_THE_TIME_Timer()

'COUNT_SHOW_THE_TIME = COUNT_SHOW_THE_TIME + 1
'If COUNT_SHOW_THE_TIME > 40 Then
'    SHOW_THE_TIME.Enabled = False
'    COUNT_SHOW_THE_TIME = 0
'    Exit Sub
'End If
'
'MNU_SHOW_THE_TIME_Click

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

Public Sub FindCursor(X, Y)

Dim P As POINTAPI

GetCursorPos P
'   return x and y co-ordinate
X = P.X ' / GetSystemMetrics(0) * Screen.Width
'   for current cursor position
Y = P.Y '/ GetSystemMetrics(1) * Screen.Height

End Sub


Private Sub Timer_Pause_Update_Timer()
Timer_Pause_Update.Enabled = False
'Label53.width
Label53.BackColor = Label58.BackColor

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
Dim GS_cWnd1 As Long, GS_cWnd2 As Long, GS_cWnd3 As Long, I3 As Long

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
Dim Text1
VB_LOADER = FindWindow("#32770", "Microsoft Visual Basic")
If 1 = 2 And VB_LOADER > 0 And VB_LOADER <> OHWnd_VB_LOADER Then
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

End Sub

Private Sub Timer_GET_KEY_ASYNC_STATE_Timer()

If IsIDE = True And IsIDE_TEST = True Then Timer_GET_KEY_ASYNC_STATE.Interval = 1000

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

Exit Sub

Dim GKA, GO_SUB
'Exit Sub
For i = 1 To 255
    GKA = GetAsyncKeyState(i)
    If GKA < 0 Then
        'If i = 39 Then i = 0
        If i = 116 Then i = 0 'F5 TEST ISIDE
        If i = 27 Then Unload Me 'EXIT SUB'ESCAPE KEY
        
        If i = 0 Then Exit Sub
        
        If i = 16 Then GO_SUB = True 'LEFT SHIFT KEY
        
        If GO_SUB = True Then
            'Call KeyOrMouse
            'Exit Sub
        End If
    End If
Next

End Sub

Private Sub Timer_IS_WINAMP_RUNNING_DO_CID_Timer()
    
    Exit Sub
    
    If FindWindow(vbNullString, "CID Run Me.") > 0 Then Exit Sub
    
    If FindWindow("Winamp v1.x", vbNullString) > 0 Then
        Shell "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\Cid-RunMe.exe", vbNormalFocus
    End If

    
    If FindWindow("Winamp v1.x", vbNullString) = 0 Then
        'Fr1 = FreeFile
        'Open "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\0TextData\WinAmpEXEFileActive.txt" For Input As #Fr1
        '    Line Input #Fr1, TT$
        'Close #Fr1

        
        Dim TT$
        'TT$ = "C:\Program Files\00 WinAmp's\Winamp 556 & Remote\winamp.exe"
        If Dir("C:\Program Files\WinAmp\winamp.exe") = "" Then
            TT$ = "C:\Program Files (X86)\WinAmp\winamp.exe"
        End If
        Shell "C:\Program Files\WinAmp\winamp.exe", vbNormalNoFocus
        Exit Sub
    End If
End Sub

Private Sub Timer_PROCESS_RELOADER_WATCHER_01_Timer()
    
    'Process Lasso Pro
    Dim i
    Dim Control As Control

    On Error Resume Next
    For i = 0 To Forms.Count - 1
        For Each Control In Forms(i).Controls
            If InStr(UCase(Control.Name), "TIMER_PROCESS_RELOADER_WATCHER") > 0 Then
                    Control.Interval = 50000
                    If IsIDE = True Then
                    Control.Enabled = False
                    End If
            End If
        Next
    Next
    If IsIDE = True Then
        Exit Sub
    End If
    
    
    'Timer_PROCESS_RELOADER_WATCHER_01.Interval = 1000
    
    
    If FindWindow(vbNullString, "Process Lasso Pro") <> 0 Then
        'APP_NAME_RELOAD_IT_ER____ = ""
        Exit Sub
    End If
    If APP_NAME_RELOAD_IT_ER____ <> "" Then Exit Sub
    
    '--------------------------
    APP_NAME_RELOAD_IT_ER_EXE = "C:\Program Files\Process Lasso\ProcessLasso.exe"
    
    VAR_IN = APP_NAME_RELOAD_IT_ER_EXE
    Call PROCESS_RELOADER_WATCHER_VAR_SET_CHECK_CONDICTION(VAR_IN, VAR_OUT)
    If VAR_OUT = False Then Exit Sub
    
    '--------------------------
    APP_NAME_RELOAD_IT_ER____ = "PROCESS LASSO ____"
    
    '-----------------------------------
    VAR_ARRAY = VAR_ARRAY + 1
    BLOCK_RUN_3(VAR_ARRAY) = True
    
    Load FrmRUN_PROCESS_RELOAD
End Sub

Private Sub Timer_PROCESS_RELOADER_WATCHER_02_Timer()
    'EXPLORER
    
'    If FindWindow("CabinetWClass", vbNullString) <> 0 Then
'        'APP_NAME_RELOAD_IT_ER____ = ""
'        Exit Sub
'    End If
    If FindWindow(vbNullString, "Program Manager") <> 0 Then
        'APP_NAME_RELOAD_IT_ER____ = ""
        Exit Sub
    Else
        'If Me.WindowState = vbMinimized Then Me.WindowState = vbMinimized
    End If
    
    If APP_NAME_RELOAD_IT_ER____ <> "" Then Exit Sub
    
    DELAY_TICKER = DELAY_TICKER + 1
'    If DELAY_TICKER Mod 10 > 0 Then Exit Sub
    
    '--------------------------
    'MIGHT BE A PROCESS TIME HERE
    'TRY FIND CLASS NAME OF ONE SINGLE EXPLORER
    'PID TO HWND AND PARENT AND SCAN

'    i = FindWinPart_SEARCHER_HWND_TO_EXE("EXPLORER")
    
    '--------------------------
    APP_NAME_RELOAD_IT_ER_EXE = "C:\Windows\explorer.exe"
    PID = -1
    APP_NAME_EXE_PASS_FOR_CALL_BACK = APP_NAME_RELOAD_IT_ER_EXE
'    i = cProcesses.GetEXEID(PID, APP_NAME_EXE_PASS_FOR_CALL_BACK)
    
    If PID > 0 Then
'        APP_NAME_RELOAD_IT_ER____ = ""
'        Exit Sub
    End If
    
    '--------------------------
    'APP_NAME_RELOAD_IT_ER_EXE = "C:\Windows\explorer.exe"
    
    VAR_IN = APP_NAME_RELOAD_IT_ER_EXE
    Call PROCESS_RELOADER_WATCHER_VAR_SET_CHECK_CONDICTION(VAR_IN, VAR_OUT)
    If VAR_OUT = False Then Exit Sub
    
    '--------------------------
    APP_NAME_RELOAD_IT_ER____ = "C:\Windows\Explorer.exe ____"
    
    '-----------------------------------
    VAR_ARRAY = VAR_ARRAY + 1
    BLOCK_RUN_3(VAR_ARRAY) = True
    
    Load FrmRUN_PROCESS_RELOAD
End Sub


Private Sub Timer_PROCESS_RELOADER_WATCHER_03_Timer()
    'ClipBoard Logger
    
    Exit Sub
    
    If FindWindow(vbNullString, "ClipBoard Logger") <> 0 Then
        'APP_NAME_RELOAD_IT_ER____ = ""
        Exit Sub
    End If
    If APP_NAME_RELOAD_IT_ER____ <> "" Then Exit Sub
    
    '--------------------------
    APP_NAME_RELOAD_IT_ER_EXE = "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.exe"
    
    VAR_IN = APP_NAME_RELOAD_IT_ER_EXE
    Call PROCESS_RELOADER_WATCHER_VAR_SET_CHECK_CONDICTION(VAR_IN, VAR_OUT)
    If VAR_OUT = False Then Exit Sub
    
    '--------------------------
    APP_NAME_RELOAD_IT_ER____ = "ClipBoard Logger ____"
    
    '-----------------------------------
    VAR_ARRAY = VAR_ARRAY + 1
    BLOCK_RUN_3(VAR_ARRAY) = True
    
    Load FrmRUN_PROCESS_RELOAD
End Sub


Private Sub Timer_PROCESS_RELOADER_WATCHER_04_Timer()
    'GoodSync2Go
    
    'TxtEXE.Text = GetFileFromHwnd(lhWnd)
    If FindWinPart_SEARCHER("GoodSync2Go -") > 0 Then
        'APP_NAME_RELOAD_IT_ER____ = ""
        Exit Sub
    End If
    If APP_NAME_RELOAD_IT_ER____ <> "" Then Exit Sub
    
    '--------------------------
    APP_NAME_RELOAD_IT_ER_EXE = "C:\GoodSync\GoodSync2Go-v10.exe"
    
    VAR_IN = APP_NAME_RELOAD_IT_ER_EXE
    Call PROCESS_RELOADER_WATCHER_VAR_SET_CHECK_CONDICTION(VAR_IN, VAR_OUT)
    If VAR_OUT = False Then Exit Sub
    
    '--------------------------
    APP_NAME_RELOAD_IT_ER____ = "GoodSync2Go ____"
    
    '-----------------------------------
    VAR_ARRAY = VAR_ARRAY + 1
    BLOCK_RUN_3(VAR_ARRAY) = True
    
    Load FrmRUN_PROCESS_RELOAD
End Sub
Private Sub Timer_PROCESS_RELOADER_WATCHER_05_Timer()
    'GoodSync
    
    If FindWinPart_SEARCHER("GoodSync -") > 0 Then
        'APP_NAME_RELOAD_IT_ER____ = ""
        Exit Sub
    End If
    If APP_NAME_RELOAD_IT_ER____ <> "" Then Exit Sub
    
    '--------------------------
    APP_NAME_RELOAD_IT_ER_EXE = "C:\Program Files\Siber Systems\GoodSync\GoodSync-v10.exe"
    
    VAR_IN = APP_NAME_RELOAD_IT_ER_EXE
    Call PROCESS_RELOADER_WATCHER_VAR_SET_CHECK_CONDICTION(VAR_IN, VAR_OUT)
    If VAR_OUT = False Then Exit Sub
    
    '--------------------------
    APP_NAME_RELOAD_IT_ER____ = "GoodSync DESKTOP ____"
    
    '-----------------------------------
    VAR_ARRAY = VAR_ARRAY + 1
    BLOCK_RUN_3(VAR_ARRAY) = True
    
    Load FrmRUN_PROCESS_RELOAD
End Sub



Private Sub Timer_PROCESS_RELOADER_WATCHER_07_Timer()
    'GoodSync
    
    If FindWinPart_SEARCHER("HwndWrapper[hubiC.exe") > 0 Then
        'APP_NAME_RELOAD_IT_ER____ = ""
        Exit Sub
    End If
    If APP_NAME_RELOAD_IT_ER____ <> "" Then Exit Sub
    
    '--------------------------
    APP_NAME_RELOAD_IT_ER_EXE = "C:\Program Files\OVH\hubiC\hubiC.exe" ' run --showsync"
    
    VAR_IN = APP_NAME_RELOAD_IT_ER_EXE
    Call PROCESS_RELOADER_WATCHER_VAR_SET_CHECK_CONDICTION(VAR_IN, VAR_OUT)
    If VAR_OUT = False Then Exit Sub
    
    '--------------------------
    APP_NAME_RELOAD_IT_ER____ = "HUBIC ____"
    
    '-----------------------------------
    VAR_ARRAY = VAR_ARRAY + 1
    BLOCK_RUN_3(VAR_ARRAY) = True
    
    Load FrmRUN_PROCESS_RELOAD
End Sub

Private Sub Timer_SET_FOCUS_Timer()

'frmMain.SetFocus

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
    
    ' Get cursor cordinates
    GetCursorPos tPA
    ' Set label caption to cursor cordinates
    lblCordi.Caption = "X: " & tPA.X & "  Y: " & tPA.Y
    
    If TIMER2_TIMER_BEGAN + TimeSerial(0, 0, 20) > Now Then
    
        Label48.Caption = Format(20 - DateDiff("s", TIMER2_TIMER_BEGAN, Now), "00") + " Second"
        Label48.FontBold = True
        Label48.FontSize = 15
        
        i_string = "USE " + Format(DateDiff("s", Now, TIMER2_TIMER_BEGAN + TimeSerial(0, 0, 20), "00")) + " SECOND HOOVER"
        If i_string <> MNU_HOOVER_20_SECOND.Caption Then MNU_HOOVER_20_SECOND.Caption = i_string
            mWnd = WindowFromPoint(tPA.X, tPA.Y)
            Call ChunkCodeOnMouse
        Else
            
            If TIMER2_TIMER_BEGAN <> 0 Then
                TIMER2_TIMER_BEGAN = 0
                Label48.Caption = "20 Sec"
                ' Label48.FontBold = False
                ' Label48.FontSize = 10
            End If
    End If
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

If Timer_Pause_Update.Enabled = True Then Exit Sub

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

Label46.Caption = "Process Counter:" ' + Str(lstProcess.ListCount - 2)

If O_lstProcess_ListCount = -1 Then
    O_lstProcess_ListCount = lstProcess.ListCount
End If

If O_lstProcess_ListCount > lstProcess.ListCount Then TAG_VAR = " && Less"
If O_lstProcess_ListCount < lstProcess.ListCount Then TAG_VAR = " && Higher"
If O_lstProcess_ListCount = lstProcess.ListCount Then TAG_VAR = " && Equal"


If O_lstProcess_ListCount > lstProcess.ListCount Then
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

    Label51.Left = Label52.Left + 50
    'THE WINDDING LABLE MUST BE SMALLER THAN Label52 AS IS OFFSET IN
    Label51.width = ((Label52.Left + Label52.width) - Label51.Left) - 20
    Label51.Caption = TAG_VAR_2
    Label50.Left = (Label52.Left + Label52.width) + 20
    Label51.Top = Label52.Top
    Label51.height = Label52.height - (Label52.height - Label51.height) ' Label_ProcessSetBar_Height
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
    
    Label51.Left = Label52.Left + 50
    'THE WINDDING LABLE MUST BE SMALLER THAN Label52 AS IS OFFSET IN
    Label51.width = ((Label52.Left + Label52.width) - Label51.Left) - 20
    Label50.Left = (Label52.Left + Label52.width) + 20
    Label51.Caption = TAG_VAR_2
    Label51.Top = Label52.Top
    Label51.height = Label52.height - (Label52.height - Label51.height) ' Label_ProcessSetBar_Height
End If
If O_lstProcess_ListCount = lstProcess.ListCount Then
    'EQUAL
    TAG_VAR_2 = Chr(240)
    Label51.FontSize = 18
    
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
    
    Label51.Left = Label52.Left + 50
    'THE WINDDING LABLE MUST BE SMALLER THAN Label52 AS IS OFFSET IN
    Label51.width = ((Label52.Left + Label52.width) - Label51.Left) - 20
    Label50.Left = (Label52.Left + Label52.width) + 20
    Label51.Caption = TAG_VAR_2
    Label51.Top = Label52.Top
    Label51.height = Label52.height - (Label52.height - Label51.height) ' Label_ProcessSetBar_Height

End If



'----
'Wingdings 3 character set and equivalent Unicode characters
'http://www.alanwood.net/demos/wingdings-3.html
'----

Label50.Caption = Str(lstProcess.ListCount - 2) + TAG_VAR + "_Timer 1 Min"

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

Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String
Dim wText As String
Dim HighTest_hWnd As Long

Dim T_COMPARE, T_COUNTER
Dim RI
Dim FINDER_LINE, FINDER_LINE_1
Dim FINDER_LINE_HWND As Long
Dim VAR As Long

Dim T_COUNTER_FLAG_VAR

Do
    List_SORT_FOR_AHK_LIMITER.Clear
    
    'Find the first window
    test_hwnd = FindWindowDLL(ByVal 0&, ByVal 0&)
    T_COMPARE = ""
    T_COUNTER = 1
    
    Do While test_hwnd <> 0
        
        wText = GetWindowTitle(test_hwnd)
    
        ' Autokey -- 58-Auto Repeat Browser Function Set.ahk
        ' ahk_class #32770
        '".ahk - AutoHotkey v"
        
    '    If InStr(wText, "Autokey -- 58-Auto Repeat Browser Function Set.ahk") > 0 Then
    '        Stop
    '    End If
        
        If InStr(wText, ".ahk") > 0 Then
            List_SORT_FOR_AHK_LIMITER.AddItem Trim(wText) + " ---------X " + Str(test_hwnd)
        End If
        
        'retrieve the next window
        test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    
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
        FINDER_LINE_HWND = Val(Mid(FINDER_LINE, InStr(FINDER_LINE, " ---------X  ") + Len(" ---------X  ")))
        If FINDER_LINE_1 = T_COMPARE Then
            T_COUNTER = T_COUNTER + 1
            If T_COUNTER > 15 Then
                T_COUNTER_FLAG_VAR = True
                
                
                
                ' NOT WORK FOR WHAT WE WANTER
                ' closewindow (FINDER_LINE_HWND)
                
                ' FINDER_LINE_HWND_2 = GetParent(FINDER_LINE_HWND)
                ' closewindow (FINDER_LINE_HWND_2)
                
                If FINDER_LINE_HWND > 0 Then
                    PID = -1
                    VAR = cProcesses.Convert(FINDER_LINE_HWND, PID, cnFromhWnd Or cnToProcessID)
                End If
                If PID > 0 Then
                    Result = cProcesses.Process_Kill(PID)
                End If
            
            End If
        Else
            T_COUNTER = 1
        End If
        
        T_COMPARE = FINDER_LINE_1
        
    Next
    
Loop Until T_COUNTER_FLAG_VAR = False
    


End Function


Private Sub MNU_TASK_KILLER_AUTOHOTKEYS_2_Click()
    
Call MNU_TASK_KILLER_AUTOHOTKEYS_Click
    
Me.WindowState = vbMinimized
    
End Sub

Private Sub Lab_KILL_AHK_Click()

Call COLOUR_BOX_SELECTOR_RESTORE_DEFAULT
Lab_KILL_AHK.BackColor = RGB(255, 255, 255)

Call MNU_TASK_KILLER_AUTOHOTKEYS_Click

Me.WindowState = vbMinimized

End Sub

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


Private Sub SCREEN_SHOT_HERE_2(HWND_NUMBER)
    
    If InStr(Command_Screen_Shot_Auto_ClipBoard_er.Caption, "Archive Mode _ON") = 0 Then
        Exit Sub
    End If
    
    Dim TT, TS, FOLDERNAME_AUTO
    Dim HWND_LONG As Long
    
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
    
    HWND_LONG = Val(HWND_NUMBER)
    
    'TT = Print_HWND_FORM_ontoForm(SCREEN_CAP, Val(HWND_NUMBER))
    TT = Print_HWND_FORM_ontoForm_EVEN_WHEN_BEHIND_ANOTHER(SCREEN_CAP, HWND_LONG)
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

Private Sub SCREEN_SHOT_HERE(HWND_NUMBER)
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
    
    
    TT = Print_HWND_FORM_ontoForm(SCREEN_CAP, Val(HWND_NUMBER))
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



Private Sub txtMemoryUsage_Click()
'txtMemoryUsage_HERE
End Sub

Private Sub txtTitle_CLICK()
Clipboard.Clear
Clipboard.SetText txtTitle
End Sub

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

Private Sub Timer_COMMAND_EXE_MENU_NOTEPAD_PICK_UP_Timer()

On Error Resume Next

If Dir(FILENAME_PATH_EXE_MENU) = "" Then
    Beep
    MsgBox "ELITE SPY __ FILENAME WAIT RETURN MODIFIED __ HAD AN ERROR CREATING AND OR EITHER DISAPPEARED " + FILENAME_PATH_EXE_MENU, vbMsgBoxSetForeground
    Timer_COMMAND_EXE_MENU_NOTEPAD_PICK_UP.Enabled = False
    Exit Sub
End If

If Err.Number > 0 Then Exit Sub

Set F1 = FS.GetFile(FILENAME_PATH_EXE_MENU)
If COMMAND_EXE_MENU_NOTEPAD_PICK_UP_DATE <> F1.DateLastModified Then
    COMMAND_EXE_MENU_NOTEPAD_PICK_UP_DATE = F1.DateLastModified
    Timer_COMMAND_EXE_MENU_NOTEPAD_PICK_UP.Enabled = False
    
    If Err.Number > 0 Then Exit Sub

    Call READ_EXE_MENU_FILE_AND_POPULATE_MENU
End If

End Sub

Sub READ_EXE_MENU_FILE_AND_POPULATE_MENU()

    Dim LINE_GET
    Dim MENU_02_ACTIVE
    Dim MENU_03_ACTIVE
    Dim MENU_02_COUNTER, MENU_03_COUNTER

    If FILENAME_PATH_EXE_MENU <> "" Then
        If Dir(FILENAME_PATH_EXE_MENU) = "" Then Exit Sub
    Else
        Exit Sub
    End If

    Err.Clear
    FR1 = FreeFile
    Open FILENAME_PATH_EXE_MENU For Input As #FR1
        If Err.Number > 0 Then Exit Sub
        Do
            If Err.Number > 0 Then Exit Sub
            
            Line Input #FR1, LINE_GET
            If InStr(LINE_GET, "; 01 TO 20 OF MENU 02 __ LEAVE THESE LINE MARKER") > 0 Then
                MENU_02_ACTIVE = True
                MENU_03_ACTIVE = False
            End If
            If InStr(LINE_GET, "; 01 TO 20 OF MENU 03 __ LEAVE THESE LINE MARKER") > 0 Then
                MENU_02_ACTIVE = False
                MENU_03_ACTIVE = True
            End If
            
            If Mid(LINE_GET, 1, 1) <> ";" Then
                If MENU_02_ACTIVE = True And MENU_02_COUNTER <= 20 Then
                    MENU_02_COUNTER = MENU_02_COUNTER + 1
                    MNU_EXE_MENU_SCRIPT_FILL_02.Item(MENU_02_COUNTER).Visible = True
                    MNU_EXE_MENU_SCRIPT_FILL_02.Item(MENU_02_COUNTER).Caption = LINE_GET
                End If
                If MENU_03_ACTIVE = True And MENU_03_COUNTER <= 20 Then
                    MENU_03_COUNTER = MENU_03_COUNTER + 1
                    MNU_EXE_MENU_SCRIPT_FILL_02.Item(MENU_03_COUNTER).Visible = True
'                    MNU_EXE_MENU_SCRIPT_FILL_03.Item(MENU_03_COUNTER).Caption = LINE_GET
                End If
                
            End If
        Loop Until EOF(FR1)
    Close #FR1
    
    If MENU_02_COUNTER = 0 Then
        MNU_EXE_MENU_SCRIPT_FILL_02.Item(1).Caption = "NONE ITEM POLULATED"
    End If
    If MENU_03_COUNTER = 0 Then
'        MNU_EXE_MENU_SCRIPT_FILL_03.Item(1).Caption = "NONE ITEM POLULATED"
    End If

'Dim i
'On Error Resume Next
'    For i = 0 To Forms.Count - 1
'        For Each Control In Forms(i).Controls
'            If InStr(Control.Name, "MNU_EXE_MENU_SCRIPT_FILL_") > 0 Then
'                If Control.Caption = "--01" Then
'                    Control.Visible = False
'                End If
'            End If
'        Next
'    Next
'
'    For i = 0 To Forms.Count - 1
'        For Each Control In Forms(i).Controls
'            If InStr(Control.Name, "MNU_EXE_MENU_SCRIPT_FILL_") > 0 Then
'                If InStr(Control.Caption, "__NOT EXIST") = 0 Then
'                    If Dir(Control.Caption) = "" Then
'                        Control.Caption = Control.Caption + "__NOT EXIST"
'                    End If
'                End If
'            End If
'        Next
'    Next
End Sub



Private Sub MNU_GET_EXPLORER_CURRENT_SESSION_CLIPBOADIN_Click()

Call FindWindow_Get_All_Explorer("QUITE MSGBOX=FALSE")

End Sub


'Private Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long

Function FindWindow_Get_All_Explorer(QUITE_MSGBOX) As String

FindWindow_Get_All_Explorer = ""

Dim Huge, VAR_STRING
Dim WINDOW_TITLE
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String
Huge = 0

test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
VAR_STRING = ""
Do While test_hwnd <> 0
    
    WINDOW_TITLE = GetWindowTitle(test_hwnd)
    
    '--------------------------------------------------
    'C:\Windows\explorer.exe
    '--------------------------------------------------
    If GetWindowClass(test_hwnd) = "CabinetWClass" Then
        Huge = Huge + 1
        VAR_STRING = VAR_STRING + WINDOW_TITLE + vbCrLf + vbCrLf
    End If
        
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

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




Public Sub MNU_TASK_KILLER_Click()

Dim RUN_FROM_AUTO_TASK_KILLER

'FORM1.MNU_TASK_KILLER_Click
Beep
If RUN_FROM_AUTO_TASK_KILLER = True Then
    RUN_FROM_AUTO_TASK_KILLER = False
    i = FindWindow(vbNullString, "Administrator:  C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT_03_PROCESS_KILLER.BAT")
    If i > 0 Then Exit Sub
Else
'Me.WindowState = vbMinimized

End If

Me.WindowState = vbMinimized
'Shell "C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT_03_PROCESS_KILLER.BAT", vbMaximizedFocus

Shell "CMD /C START """" /REALTIME /MAX ""C:\SCRIPTER\SCRIPTER CODE -- BAT\BAT_03_PROCESS_KILLER.BAT""", vbMinimizedNoFocus
Beep
End Sub


Private Sub MNU_MAXIMIZE_GOODSYNC_Click()

Dim GOODSYNC_WINDOW_HWND
GOODSYNC_WINDOW_HWND = FindWindow("{B26B00DA-2E5D-4CF2-83C5-911198C0F009}", vbNullString)

ShowWindow GOODSYNC_WINDOW_HWND, SW_MAXIMIZE
Beep
Me.WindowState = vbMinimized

End Sub



Private Sub Timer_1_SECOND_Timer()

Dim FN1, FN2, F1, F2, FD1, FD2, DATESET As Date, XX

FN1 = "C:\Users\" + GetUserName + "\AppData\Roaming\Logitech\SetPoint\user.xml"
FN2 = "C:\SCRIPTER\SYNC_FOLDER_1\##_0005\user.xml"
If FS.FileExists(FN1) = True And FS.FileExists(FN2) = True Then
    Set F1 = FS.GetFile(FN1)
    Set F2 = FS.GetFile(FN2)
    FD1 = F1.DateLastModified
    If OLD_FD <> "" Then
        If InStr(OLD_FD, Format(FD1, "YYY-MM-DD HH-MM-SS")) = 0 Then
            FD1 = F1.DateLastModified - TimeSerial(0, 1, 0)
            DATESET = FD1
            TT = SetFileDateTime(FN1, DATESET)
            DATESET = F2.DateLastModified - TimeSerial(0, 2, 0)
            TT = SetFileDateTime(FN2, DATESET)
        End If
    End If
    If InStr(OLD_FD, Format(FD1, "YYY-MM-DD HH-MM-SS")) = 0 Then
        OLD_FD = OLD_FD + " ---- " + Format(FD1, "YYY-MM-DD HH-MM-SS")
    End If
    If Len(OLD_FD) > 800 Then
        XX = Len(OLD_FD) - 800
        If XX < 1 Then XX = 1
        OLD_FD = Mid(OLD_FD, XX)
    End If
End If

'
'If Timer_1_MINUTE.Interval <> 60000 Then
'    Timer_1_MINUTE.Interval = 60000
'End If

'Call RS232_LOGGER

End Sub

Sub RS232_LOGGER()

' ------------------------------------------------------------------
' NEW RS232 LOGGER
' OUTPUT FROM HERE WILL BE FILE BLOW REM LINE
' RSR232 DETECT A PIR
' A FILE IS CREATED WHEN THE PIR FEEL SOMETHING
' AND WHEN ASK IF NONTHING FILE WILL BE DELETER
' ------------------------------------------------------------------
On Error Resume Next

Shell "D:\VB6\VB-NT\00_Best_VB_01\RS232 LOGGER PIR\RS232 LOGGER.exe", vbHide

'Kill "C:\SCRIPTOR DATA\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 14-Brightness With Dimmer_" + GetComputerName + ".txt"

End Sub



Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim R As Long, Pic As PicBmp, IPic As IPicture, IID_IDispatch As GUID
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    With Pic
        .Size = Len(Pic)
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
    Dim HWndx, LEFT_RIGHT_INSET___________, i1, i2
    HWndx = GetForegroundWindow
    
    GetWindowRect HWndx, R
    
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




Function Print_HWND_FORM_ontoForm(ByVal Form As Form, HWND_NUMBER As Long)
    Dim R As RECT
    Dim HWndx
    
    HWndx = HWND_NUMBER
    
    'HWndx = GetForegroundWindow
        
    GetWindowRect HWndx, R
    
    Set Form.Picture = hDCToPicture(GetDC(0), R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top)
End Function

Function Print_HWND_FORM_ontoForm_2(ByVal Form As Form, HWND_NUMBER As Long)
    'NOT OVERLAP
    Dim R As RECT
    Dim HWndx
    
    HWndx = HWND_NUMBER
    
    'HWndx = GetForegroundWindow
        
    GetWindowRect HWndx, R
    
    Set Form.Picture = hDCToPicture(GetDC(0), R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top)

    'rv = SendMessage(Me.hwnd, WM_PRINT, SCREEN_CAP_PICTURE.Picture1.HDC, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
    'rv = SendMessage(Me.hwnd, WM_PRINT, SCREEN_CAP_PICTURE.Picture1.HDC, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
    
    'rv = SendMessage(HWND_NUMBERING, WM_PAINT, SCREEN_CAP.HDC, 0)
    'rv = SendMessage(HWND_NUMBERING, WM_PRINT, SCREEN_CAP.HDC, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)

End Function


Private Function Print_HWND_FORM_ontoForm_EVEN_WHEN_BEHIND_ANOTHER(ByVal Form As Form, HWND_NUMBER As Long)
Dim rv As Long
Dim HWND_NUMBERING As Long

    HWND_NUMBERING = Val(HWND_NUMBER)
    
    'Clipboard.Clear
    
'    Picture1.Width = Me.Width
'    Picture1.Height = Me.Height
   
    'Picture1.AutoRedraw = True
    
    'rv = SendMessage(HWND_NUMBERING, WM_PAINT, SCREEN_CAP.HDC, 0)
    'rv = SendMessage(HWND_NUMBERING, WM_PRINT, SCREEN_CAP.HDC, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
    
'    rv = SendMessage(HWND_NUMBERING, WM_PAINT, SCREEN_CAP_PICTURE.Picture1.HDC, 0)
'    rv = SendMessage(HWND_NUMBERING, WM_PAINT, SCREEN_CAP_PICTURE.Picture1.HDC, 0)
    
    rv = SendMessage(Me.hWnd, WM_PRINT, SCREEN_CAP_PICTURE.Picture1.HDC, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
    rv = SendMessage(Me.hWnd, WM_PRINT, SCREEN_CAP_PICTURE.Picture1.HDC, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
    
    SCREEN_CAP_PICTURE.Show
    
    
    'Set SCREEN_CAP_PICTURE.Picture1.Picture = SCREEN_CAP_PICTURE.Picture1.Image
    'Set SCREEN_CAP.Picture1.Picture = SCREEN_CAP_PICTURE.Picture1.Image
    
    'Set Form.Picture = SCREEN_CAP_PICTURE.Picture1.Image
    
    'Clipboard.SetData Picture1.Picture, vbCFBitmap

End Function


Private Sub Timer_VIRCOP_Timer()

If MNU_PAUSE_VIRTUA_COP_2.Checked = True Then Exit Sub

If Timer_VIRCOP.Interval <> 10000 Then Timer_VIRCOP.Interval = 10000

'Dim lhWndParent
Dim Hwnd_Var As Long, i As Long

If FindWindow(vbNullString, "VirtuaCop 2") > 0 Then

'    SetForegroundWindow FindWindow(vbNullString, "VirtuaCop 2")

'    lngI = SetFocuses(PopBack)
'    SetForegroundWindow PopBack
    
'    Putfocus FindWindow(vbNullString, "VirtuaCop 2")
'    --
'    lhWndParent = GetParent(FindWindow(vbNullString, "VirtuaCop 2"))
'    Putfocus lhWndParent

    Hwnd_Var = FindWindow(vbNullString, "VirtuaCop 2")
    If Hwnd_Var > 0 Then
        i = GetForegroundWindow()
    Else
        'Exit Sub
    End If

    If i <> OLD_i_VIRTUA_COP And OLD_i_VIRTUA_COP > 0 Then 'and Hwnd_Var<>i Then
        
        If Hwnd_Var = 0 Then
            MNU_VIRTUA_COP_2_PAUSED_STATE.Caption = "____ VIRTUA_COP_2 NOT LOADED"
            Exit Sub
        Else
            '-----------------------------------------------------------------------
            'THE FOREGROUND HAS CHANGED AND REQUIRE
            'I THINKER A RESET TO SOME TIMER OR SOMETHING TO NOT IMEDIATE LOAD AGAIN
            '-----------------------------------------------------------------------
            'Call MNU_PAUSE_VIRTUA_COP_2_Click
            'NOT RESET PAUSE BACK ON
            'BUT INSTEAD TIMER
            '-----------------------------------------------------------------------
            Timer_VIRCOP.Enabled = False
            Timer_VIRCOP.Interval = 10000
            Timer_VIRCOP.Enabled = True
        
        End If
        
        Timer_VIRCOP.Enabled = False
        Timer_VIRCOP.Interval = 10000
        Timer_VIRCOP.Enabled = True
        OLD_i_VIRTUA_COP = i
        Exit Sub
    End If
    OLD_i_VIRTUA_COP = i

    'MISSING CODE FROM MERGE NOT FOUND AT TIME
    'HURRY TO THE BACKUP WINMERGE AND GOODSYNC
    '-----------------------------------------
    'ChunkCodeOn_HWND (Hwnd_Var)
    '-----------------------------------------
    
    If i <> Hwnd_Var Then
        'Stop
        
        Beep
        'Debug.Print "Yes ---- " + Str(Time)
        cmdMaximize_Click
        i = SetForegroundWindow(Hwnd_Var)
    Else
        'Debug.Print "Not ---- " + Str(Time)
    End If
    
    'cmdMaximize_Click
    
End If

Exit Sub

If lblCordi.Caption = "X: 0  Y: 0" Then
    
    '--GetForegroundWindow Lib "user32" () As Long
    'Mwnd = WindowFromPoint(0, 0)
    
    'hWnd2 = SetForegroundWindow(Mwnd)
End If

End Sub

Private Sub Timer3_Timer()

If FindWindow(vbNullString, "VirtuaCop 2") = 0 Then
    Exit Sub
End If

'If GetAsyncKeyState(27) Then Unload Me

End Sub

'Private Sub txtClass_CLICK()
'Clipboard.Clear
'Clipboard.SetText txtClass
'End Sub

'Private Sub TxtEXE_CLICK()
'Clipboard.Clear
'Clipboard.SetText TxtEXE
'End Sub

'Private Sub txtTitle_CLICK()
'Clipboard.Clear
'Clipboard.SetText txtTitle
'End Sub



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


'------------------------------------------------------
'------------------------------------------------------
'------------------------------------------------------
'------------------------------------------------------

Sub TIMER_SET_CAMERA_CHECK_TIMER()

'XDrive1_ListCount = Drive1.ListCount
'If XDrive1_ListCount = O_DRIVE1_ListCount Then Exit Sub
'
'If XDrive1_ListCount <> O_DRIVE1_ListCount Then
'    If VAR_MC__HX60V_SET = False Then
'        O_DRIVE1_ListCount = XDrive1_ListCount
'    End If
'End If
'
'If Dir("MC - HX60V", vbVolume) = "" Then
'    TIMER_SET_CAMERA_CHECK.Interval = 10000
'    If VAR_MC__HX60V_SET = True Then
'        VAR_MC__HX60V_SET = False
'    End If
'End If
'
'If Dir("MC - HX60V", vbVolume) <> "" Then
'    If VAR_MC__HX60V_SET = False Then
'        VAR_MC__HX60V_SET = True
'        TIMER_SET_CAMERA_WITH_SECURITY_INFO.Enabled = True
'        TIMER_SET_CAMERA_CHECK.Interval = 4000
'    End If
'End If

End Sub



Sub TIMER_SET_CAMERA_WITH_SECURITY_INFO_TIMER()
''Exit Sub ' here ----
'PASS_DRIVE_AT = ""
'TO_GO = False
'For R = 69 To 89
'    On Error Resume Next
'    If Dir(Chr(R) + ":\DCIM", vbDirectory) <> "" Then TO_GO = True
'    If Err.Number > 0 Then
'        TIMER_SET_CAMERA_WITH_SECURITY_INFO.Enabled = False
'        Exit Sub
'    End If
'    On Error GoTo 0
'    If Dir(Chr(R) + ":\MP_ROOT", vbDirectory) <> "" Then TO_GO = True
'    If Dir(Chr(R) + ":\", vbDirectory) <> "" Then
'        If Dir(Chr(R) + ":\", vbVolume) = "MC - HX60V" Then
'            TO_GO = True
'        End If
'    End If
'
'
'    If TO_GO = True Then
'        PASS_DRIVE_AT = Chr(R)
'        '-----------------------------------
'        Call COPY_MY_SECURITY_INFO_TO_CAMERA(PASS_DRIVE_AT)
'        '-----------------------------------
'        Exit For
'
'    End If
'Next
'
'TIMER_SET_CAMERA_WITH_SECURITY_INFO.Enabled = False
'Exit Sub
'
''--------------------------------------------------------------
''WAS GOING TO CHECK VOLUME NAME EXIST WITHOUT CHECK FOLDER PAIR
''BUT TAKEN CARE OF ABOVE
''ANOTHER WAY OF WORK AROUND
''--------------------------------------------------------------
''ENDER CODE Tue 04 April 2017 10:39:48----------
''--------------------------------------------------------------
'
'
'If 1 = 2 And PASS_DRIVE_AT = "" Then
'    If Dir("MC - HX60V", vbVolume) <> "" Then
'        TO_GO = False
'        For R = 69 To 89
'            If Dir(Chr(R) + ":\", vbDirectory) <> "" Then
'                If Dir(Chr(R) + ":\", vbVolume) = "MC - HX60V" Then
'                    TO_GO = True
'                End If
'            End If
'
'
'            If TO_GO = True Then
'                PASS_DRIVE_AT = Chr(R)
'                '-----------------------------------
'                Call COPY_MY_SECURITY_INFO_TO_CAMERA(PASS_DRIVE_AT)
'                '-----------------------------------
'                Exit For
'            End If
'        Next
'    End If
'End If
'
'
'TIMER_SET_CAMERA_WITH_SECURITY_INFO.Enabled = False

End Sub


Private Sub TIMER_SET_CAMERA_UP_TO_MARK_FOLDER_Timer()

'If Timer_VICE_VERSA_TIMER_ON.Enabled = False Then Exit Sub
'
'On Error Resume Next
'Dir2.Path = "W:\DCIM"
'
'If Err.Number > 0 Then
'    ERR_COUNT1 = ERR_COUNT1 + 1
'    If ERR_COUNT1 > 500 Then
'        TIMER_SET_CAMERA_UP_TO_MARK_FOLDER.Enabled = False
'    End If
'    If Err.Number > 0 Then
'        Exit Sub
'    End If
'End If
'On Error GoTo 0
'
'
'TIMER_SET_CAMERA_UP_TO_MARK_FOLDER.Enabled = False
'
'Call LAST_FOLDER_CAMERA("W:\DCIM")
'Call LAST_FOLDER_CAMERA("W:\MP_ROOT")
'
''-----------------------------------
'PASS_DRIVE_AT = "W"
'Call COPY_MY_SECURITY_INFO_TO_CAMERA(PASS_DRIVE_AT)
''-----------------------------------
'
''WRITE_INFO_TXT_IN_EMPTY_FOLDER
'Call WRITE_INFO_TXT_IN_EMPTY_FOLDER("W:\DCIM")
'Call WRITE_INFO_TXT_IN_EMPTY_FOLDER("W:\MP_ROOT")

End Sub

Sub COPY_MY_SECURITY_INFO_TO_CAMERA(PASS_DRIVE_AT)
    
'    On Error Resume Next
'    i = App.Path + "\CAMERA DATA"
'
'    File2.Path = i
'    If Err.Number > 0 Then Exit Sub
'    File2.FileName = "*.*"
'
'    For R1 = 0 To File2.ListCount - 1
'
'        FN1 = File2.Path + "\" + File2.List(R1)
'        If FS.FileExists(FN1) = True Then
'            Set F1 = FS.GetFile(FN1)
'        End If
'
'        'FN2 = "W:\" + File2.List(R1)
'        FN2 = PASS_DRIVE_AT + ":\" + File2.List(R1)
'        If FS.FileExists(FN2) = True Then
'            Set F2 = FS.GetFile(FN2)
'        End If
'
'        NOT_COPY = False
'        If FS.FileExists(FN1) = True Then
'            If FS.FileExists(FN2) = True Then
'                If F1.DateLastModified = F2.DateLastModified Then
'                    NOT_COPY = True
'                End If
'            End If
'        End If
'
'        If NOT_COPY = False Then
'            FS.CopyFile FN1, FN2
'        End If
'    Next

End Sub


Sub WRITE_INFO_CAMERA(R2_INDEX, INFO_TEXT, NEW_DATE_LEARN)

    Dim i_2, PHOTO_DATE
    Dim STRIP_FILE_PATH_LINE_ON_FIRST_BACK_SLASHER
    Dim FOLDER_NAME
    Dim String_under_line_Space
    Dim Folder_Part_Path
    
    Set F = FS.GetFolder(Dir2.List(R2_INDEX))
    PHOTO_DATE = F.DateLastModified
    i_2 = Mid(Dir2.List(R2_INDEX), InStrRev(Dir2.List(R2_INDEX), "\") + 1)
    
    i_2 = i_2 + "_" + Format(PHOTO_DATE, "YY_MMM_DD_DDD")
    
    String_under_line_Space = String(1, "__")
    If UCase(Format(PHOTO_DATE, "MMM")) = "APR" Then String_under_line_Space = String(2, "_")
    If UCase(Format(PHOTO_DATE, "MMM")) = "MAY" Then String_under_line_Space = String(1, "_")
    
    STRIP_FILE_PATH_LINE_ON_FIRST_BACK_SLASHER = i_2 + String_under_line_Space + Index_String_Camera + INFO_TEXT
    FOLDER_NAME = Dir2.Path + "\" + STRIP_FILE_PATH_LINE_ON_FIRST_BACK_SLASHER
    
    MkDir FOLDER_NAME
    'i = SetFileDateTime(FOLDER_NAME, F.DateLastModified)
'    i = SetFOLDERDateTime(FOLDER_NAME, F.DateLastModified)
    
    
    Folder_Part_Path = Mid(FOLDER_NAME, InStrRev(FOLDER_NAME, "\") + 1)
    
    
    i = SetFOLDERDateTime(Dir2.Path, Folder_Part_Path, NEW_DATE_LEARN)

    
    
'    FR1 = FreeFile
'    Open FOLDER_NAME + "\TEXT_INFO_CAMERA_PHOTO_FOLDER_DATE.TXT" For Output As #FR1
'        Print #FR1, PHOTO_DATE
'        Print #FR1, "THE DATE THIS FOLDER CREATOR AND IMAGE ONE NEXT TO HER"
'        Print #FR1, FOLDER_NAME
'    Close #FR1
'
'    FR1 = FreeFile
'    Open FOLDER_NAME + "\TEXT_INFO_CAMERA_PHOTO_FOLDER_DATE JPG NULL.JPG" For Output As #FR1
'        Print #FR1, PHOTO_DATE
'    Close #FR1

End Sub

Sub CAMERA_FOLDER_TO_MATCH_WITH_DATE_FOR_WIFI_SDCF()
'__ CAMERA_FOLDER_TO_MATCH_WITH_DATE_FOR_WIFI@SDCF()

Dim PASS_DRIVE_AT
Dim TO_GO, R1, R2, R3, i_2, PHOTO_DATE
Dim STRIP_FILE_PATH_LINE_ON_FIRST_BACK_SLASHER
Dim FOLDER_NAME, R2_INDEX, NEW_PHOTO_DATE, R2_INDEX_ON

Dim R2_INDEX_ON_MOST_RECENT
Dim R2_INDEX_ON_EARLYEST
Dim R2_INDEX_ON_MONTHLY
Dim EARLYEST_PHOTO_DATE
Dim CHECK_PHOTO_DATE
Dim MONTHLY_PHOTO_DATE
Dim Search_Folder_string


Index_String_Camera = "_Info_"

PASS_DRIVE_AT = ""
TO_GO = False

For R3 = 1 To 2
If R3 = 1 Then Search_Folder_string = "DCIM"
If R3 = 2 Then Search_Folder_string = "MP_ROOT"


For R1 = 69 To 89
    On Error Resume Next
    PASS_DRIVE_AT = Chr(R1)
    Err.Clear
    TO_GO = False
    If FS.FolderExists(Chr(R1) + ":\" + Search_Folder_string) = True Then TO_GO = True

    If TO_GO = True And Err.Number = 0 Then
        PASS_DRIVE_AT = Chr(R1)
        Dir2.Path = Chr(R1) + ":\" + Search_Folder_string
        
        'DELETE FIRST
        '------------
        For R2_INDEX = 0 To Dir2.ListCount - 1
            If InStr(Dir2.List(R2_INDEX), Index_String_Camera) > 0 Then
                Kill Dir2.List(R2_INDEX)
                FS.DeleteFolder Dir2.List(R2_INDEX), True
            End If
        Next
        
        'earlyer
        '-------
        CHECK_PHOTO_DATE = DateSerial(2999, 1, 1)
        For R2_INDEX = 0 To Dir2.ListCount - 1
            i_2 = Mid(Dir2.List(R2_INDEX), InStrRev(Dir2.List(R2_INDEX), "\") + 1)
            If Len(i_2) = 8 Then
            'If InStr(Dir2.List(R2_INDEX), " _ INFO MARKER") = 0 Then
                Set F = FS.GetFolder(Dir2.List(R2_INDEX))
                PHOTO_DATE = F.DateLastModified
                If PHOTO_DATE < CHECK_PHOTO_DATE Then
                    CHECK_PHOTO_DATE = PHOTO_DATE
                    EARLYEST_PHOTO_DATE = PHOTO_DATE
                    R2_INDEX_ON_EARLYEST = R2_INDEX
                End If
            'End If
            End If
        Next
        Call WRITE_INFO_CAMERA(R2_INDEX_ON_EARLYEST, "Earlyest", EARLYEST_PHOTO_DATE)
        
        'today most recent
        CHECK_PHOTO_DATE = 0
        For R2_INDEX = 0 To Dir2.ListCount - 1
            i_2 = Mid(Dir2.List(R2_INDEX), InStrRev(Dir2.List(R2_INDEX), "\") + 1)
            If Len(i_2) = 8 Then
            'If InStr(Dir2.List(R2_INDEX), " _ INFO MARKER") = 0 Then
                Set F = FS.GetFolder(Dir2.List(R2_INDEX))
                PHOTO_DATE = F.DateLastModified
                If CHECK_PHOTO_DATE < PHOTO_DATE Then
                    CHECK_PHOTO_DATE = PHOTO_DATE
                    NEW_PHOTO_DATE = PHOTO_DATE
                    R2_INDEX_ON_MOST_RECENT = R2_INDEX
                End If
            'End If
            End If
        Next
        
        CHECK_PHOTO_DATE = -10
        'monthly
'        If Search_Folder_string <> "MP_ROOT" Then
        For R2_INDEX = 0 To Dir2.ListCount - 1
            i_2 = Mid(Dir2.List(R2_INDEX), InStrRev(Dir2.List(R2_INDEX), "\") + 1)
            If Len(i_2) = 8 Then
            If InStr(Dir2.List(R2_INDEX), Index_String_Camera) = 0 Then
                Set F = FS.GetFolder(Dir2.List(R2_INDEX))
                PHOTO_DATE = F.DateLastModified
                If CHECK_PHOTO_DATE <> Month(PHOTO_DATE) Then
                    CHECK_PHOTO_DATE = Month(PHOTO_DATE)

                    MONTHLY_PHOTO_DATE = Month(PHOTO_DATE)
                    R2_INDEX_ON_MONTHLY = R2_INDEX
                    Call WRITE_INFO_CAMERA(R2_INDEX_ON_MONTHLY, "Monthly_" + Format(PHOTO_DATE, "MMMM"), PHOTO_DATE)
                End If
            End If
            End If
        Next
'            End If
        
        
        'WRITE MOST RECENT
        'Call WRITE_INFO_CAMERA(R2_INDEX_ON_MOST_RECENT, "Most_Recent")
        Call WRITE_INFO_CAMERA(R2_INDEX_ON_MOST_RECENT, "Today", NEW_PHOTO_DATE)
    End If
Next
Next
    
    
'For R1 = 69 To 89
'    On Error GoTo 0
'    If Dir(Chr(R1) + ":\MP_ROOT", vbDirectory) <> "" Then TO_GO = True
'    If Dir(Chr(R1) + ":\", vbDirectory) <> "" Then
'        If Dir(Chr(R1) + ":\", vbVolume) = "MC - HX60V" Then
'            TO_GO = True
'        End If
'    End If
'
'
'    If TO_GO = True Then
'        PASS_DRIVE_AT = Chr(R1)
'        'Dir2.Path = FOLDER
'
'        '-----------------------------------
'        Call COPY_MY_SECURITY_INFO_TO_CAMERA(PASS_DRIVE_AT)
'        '-----------------------------------
'        Exit For
'
'    End If
'Next
'
'
'
'


End Sub


Function LATEST_DATE_FILE_ON_CAMERA_CHANGED()

'If Timer_VICE_VERSA_TIMER_ON.Enabled = False Then Exit Function
'
'
''GET DATE FOR CHANGES IF VALUE IS NONE
'If OPHOTO_DATE_VAR = 0 Then
'
'    FOLDER_FILE_NAME = App.Path + "\DATE LAST FILE CHANGE ON CAMERA.TXT"
'    If Dir(FOLDER_FILE_NAME) <> "" Then
'        FR1 = FreeFile
'        Open FOLDER_FILE_NAME For Input As #FR1
'            Line Input #FR1, LINE_VAR_INFO_NOW
'        Close #FR1
'        OPHOTO_DATE_VAR = DateValue(LINE_VAR_INFO_NOW) + TimeValue(LINE_VAR_INFO_NOW)
'    End If
'End If
'
'For RSEL = 1 To 2
'    If RSEL = 1 Then FOLDER = "DCIM"
'    If RSEL = 2 Then FOLDER = "MP_ROOT"
'
''    If InStr(FOLDER, "DCIM") > 0 Then
''        VAR_SIZE_LEN = 16
''    End If
''    If InStr(FOLDER, "MP_ROOT") > 0 Then
''        VAR_SIZE_LEN = 19
''    End If
'
'    On Error Resume Next
'    Dir2.Path = FOLDER
'
'    For R = 0 To Dir2.ListCount - 1
'        'Debug.Print Dir2.List(R)
'        FLAG_UP_FILE_EXIST = -2
'        File2.Path = Dir2.List(R)
'        For R2 = 0 To File2.ListCount - 1
'            Set F = FS.GetFile(File2.Path + "\" + File2.List(R2))
'            PHOTO_DATE = F.DateLastModified
'            If PHOTO_DATE > PHOTO_DATE_VAR Then
'              PHOTO_DATE_VAR = PHOTO_DATE
'            End If
'        Next
'    Next
'Next
'
''STORE THE CHANGED DATE IF
'If OPHOTO_DATE_VAR <> PHOTO_DATE_VAR Then
'
'    FOLDER_NAME = Dir2.List(R)
'    FOLDER_FILE_NAME = App.Path + "\DATE LAST FILE CHANGE ON CAMERA.TXT"
'
'    FR1 = FreeFile
'    Open FOLDER_FILE_NAME For Output As #FR1
'        Print #FR1, Now
'    Close #FR1
'
'    LATEST_DATE_FILE_ON_CAMERA_CHANGED = True
'
'End If
'
'OPHOTO_DATE_VAR = PHOTO_DATE_VAR

End Function



Sub WRITE_INFO_TXT_IN_EMPTY_FOLDER(FOLDER)

'If InStr(FOLDER, "DCIM") > 0 Then
'    VAR_SIZE_LEN = 16
'End If
'If InStr(FOLDER, "MP_ROOT") > 0 Then
'    VAR_SIZE_LEN = 19
'End If
'
'On Error Resume Next
'Dir2.Path = FOLDER
'
'For R = 0 To Dir2.ListCount - 1
'    'Debug.Print Dir2.List(R)
'
'    FLAG_UP_FILE_EXIST = -2
'
'    File2.Path = Dir2.List(R)
'    For R2 = 0 To File2.ListCount - 1
'        If Len(Dir2.List(R)) > VAR_SIZE_LEN Then
'            If FLAG_UP_FILE_EXIST = -2 Then
'                FLAG_UP_FILE_EXIST = False
'            End If
'            If UCase(File2.List(R2)) = "INFO ----------.TXT" Then
'                FLAG_UP_FILE_EXIST = True
'            End If
'        End If
'    Next
'
'    If File2.ListCount = 0 And Len(Dir2.List(R)) > VAR_SIZE_LEN Then
'        FLAG_UP_FILE_EXIST = False
'    End If
'
'    If FLAG_UP_FILE_EXIST = False Then
'
'        FOLDER_NAME = Dir2.List(R)
'        FR1 = FreeFile
'        Open FOLDER_NAME + "\INFO ----------.TXT" For Output As #FR1
'            Print #FR1, "MUST MAKE FOLDER NOT EMPTY VICE VERSA"
'            Print #FR1, FOLDER_NAME
'            Print #FR1, Now
'        Close #FR1
'
'    End If
'Next


End Sub

Sub LAST_FOLDER_CAMERA(FOLDER)

On Error Resume Next
'
'If InStr(FOLDER, "DCIM") > 0 Then
'    VAR_SIZE_LEN = 16
'End If
'If InStr(FOLDER, "MP_ROOT") > 0 Then
'    VAR_SIZE_LEN = 19
'End If
'
'Dir2.Path = FOLDER
'
'LAST_FOLDER = ""
'LAST_FOLDER_MARK = True
'For R = Dir2.ListCount - 1 To 0 Step -1
'    'Debug.Print Dir2.List(R)
'    File2.Path = Dir2.List(R)
'    If File2.ListCount > 0 Then
'        If Len(Dir2.List(R)) = VAR_SIZE_LEN Then
'            LAST_FOLDER = Dir2.List(R)
'            Exit For
'        End If
'    End If
'
'    If Len(Dir2.List(R)) > VAR_SIZE_LEN Then
'        LAST_FOLDER_MARK = False
'    End If
'Next
'
'On Error Resume Next
'
'If LAST_FOLDER_MARK = True Then
'    FOLDER_NAME = LAST_FOLDER + " -- GOT TO HERE -- " + Format(Now, "YYYY-MM-DD -- HH-MM-SS")
'    MkDir FOLDER_NAME
'    If Err.Number > 0 Then
'        MsgBox "ERROR MAKE MARKER FOLDER FOR CAMERA" + vbCrLf + FOLDER_NAME, vbMsgBoxSetForeground
'        Exit Sub
'    End If
'    FR1 = FreeFile
'    Open FOLDER_NAME + "\INFO ----------.TXT" For Output As #FR1
'        Print #FR1, "MUST MAKE FOLDER NOT EMPTY VICE VERSA"
'        Print #FR1, Now
'    Close #FR1
'    If Err.Number > 0 Then
'        MsgBox "ERROR MAKE FILE IN FOLDER FOR CAMERA" + vbCrLf + FOLDER_NAME + "\INFO ----------.TXT", vbMsgBoxSetForeground
'        Exit Sub
'    End If
'End If
'
'
''DELETE PREVIOUS MARKER ADD NEW
''IF NOTHING ON LAST ONE
'For R = Dir2.ListCount - 2 To 0 Step -1
'    'Debug.Print Dir2.List(R)
'    If Len(Dir2.List(R)) > VAR_SIZE_LEN Then
'        RmDir Dir2.List(R)
'    End If
'    If Len(Dir2.List(R)) = VAR_SIZE_LEN Then
'        Exit For
'    End If
'Next



End Sub

'------------------------------------------------------
'------------------------------------------------------
'------------------------------------------------------
'------------------------------------------------------




''--------------------------------------
''--------------------------------------
''--------------------------------------
''--------------------------------------
''Excellent Code Set the File Time To ANything you want touch date Pro Keep you file time after rotate JPg's
'
''THIS TOGGLES A SMALL PASSWORD FEATURE THAT I EMBEDDED
'Const MPassword = False
'
''THIS IS THE CODE THAT MUST BE ENTERED (FINISHED BY PRESSING HASH)
''WHEN THE PASSWORD TOGGLE IS TRUE
'Const MSequence = 22515
'
'Dim MEntered As Variant
'
' Private Type FILETIME
'     dwLowDate  As Long
'     dwHighDate As Long
' End Type
'
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
'
'Private Const OPEN_EXISTING = 3
'Private Const FILE_SHARE_READ = &H1
'Private Const FILE_SHARE_WRITE = &H2
'Private Const GENERIC_WRITE = &H40000000
'
'Private Declare Function CreateFile Lib "kernel32" Alias _
'   "CreateFileA" (ByVal lpFileName As String, _
'   ByVal dwDesiredAccess As Long, _
'   ByVal dwShareMode As Long, _
'   ByVal lpSecurityAttributes As Long, _
'   ByVal dwCreationDisposition As Long, _
'   ByVal dwFlagsAndAttributes As Long, _
'   ByVal hTemplateFile As Long) _
'   As Long
'
'Private Declare Function LocalFileTimeToFileTime Lib _
'     "kernel32" (lpLocalFileTime As FILETIME, _
'      lpFileTime As FILETIME) As Long
'
'Private Declare Function SetFileTime Lib "kernel32" _
'   (ByVal hFile As Long, ByVal MullP As Long, _
'    ByVal NullP2 As Long, lpLastWriteTime _
'    As FILETIME) As Long
'
'Private Declare Function SystemTimeToFileTime Lib _
'    "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime _
'    As FILETIME) As Long
'
'Private Declare Function CloseHandle Lib "kernel32" _
'   (ByVal hObject As Long) As Long

Public Function SetFileDateTime(ByVal FileName As String, _
  ByVal TheDate As String) As Boolean
'************************************************
'PURPOSE:    Set File Date (and optionally time)
'            for a given file)
         
'PARAMETERS: TheDate -- Date to Set File's Modified Date/Time
'            FileName -- The File Name

'Returns:    True if successful, false otherwise
'************************************************
If Dir(FileName) = "" Then Exit Function
If Not IsDate(TheDate) Then Exit Function

Dim lFileHnd As Long
Dim lRet As Long

Dim typFileTime As FILETIME
Dim typLocalTime As FILETIME
Dim typSystemTime As SYSTEMTIME

With typSystemTime
    .wYear = Year(TheDate)
    .wMonth = Month(TheDate)
    .wDay = Day(TheDate)
    .wDayOfWeek = Weekday(TheDate) - 1
    .wHour = Hour(TheDate)
    .wMinute = Minute(TheDate)
    .wSecond = Second(TheDate)
End With

lRet = SystemTimeToFileTime(typSystemTime, typLocalTime)
lRet = LocalFileTimeToFileTime(typLocalTime, typFileTime)

lFileHnd = CreateFile(FileName, GENERIC_WRITE, _
    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, _
    OPEN_EXISTING, 0, 0)
    
lRet = SetFileTime(lFileHnd, ByVal 0&, ByVal 0&, _
         typFileTime)

CloseHandle lFileHnd
SetFileDateTime = lRet > 0

End Function
'--------------------------------------
'--------------------------------------
'--------------------------------------
'--------------------------------------



Public Function SetFOLDERDateTime(ByVal FileName As String, Folder_Part_Path, _
  ByVal TheDate As String) As Boolean

    
'From here:
'
'http://www.tek-tips.com/viewthread.cfm?qid=1372273

'ModFileDT "c:\rootdir", "folder", "1/01/2007 4:18:02 PM"

'If Mid(Filename, Len(Filename)) <> "\" Then Filename = Filename + "\"

'ModFileDT FileName, Folder_Part_Path, TheDate '"1/05/2017 1:0:00"


ModFileDT FileName, Folder_Part_Path, TheDate


End Function

Function ModFileDT(strDir, strFileName, DateTime)

    Dim objShell, objFolder
    Dim colItems
    Dim objItem
    
    Set objShell = CreateObject("Shell.Application")
'    Set objFolder = objShell.Namespace(strDir + "\" + strFileName)
    Set objFolder = objShell.Namespace(strDir + "\") ' + strFileName)
'    Set objFolderItem = objFolder.Self
'    objFolder.Items.Item(strFileName).ModifyDate = DateTime
'    objFolder.Items.Item(strFileName).ModifyDate = DateTime

'    Me.WindowState = vbNormal
    

'    MsgBox objFolder.Name
''    objFolder.Items.Item(1).name
''    objFolder.Item.name
''    objFolder.namespace.item(1)
''    objFolder.Item

'Debug.Print objFolderItem.Path

'DateTime = "02/04/2017 20:08:44"
'objFolder.Items.Item(strFileName).ModifyDate = DateTime
'objFolder.Items.Item(strFileName).ModifyDate = Now + 1
'Debug.Print objFolder.Items.Item(strFileName).ModifyDate
'objFolderItem.Path.ModifyDate = DateTime

'Debug.Print objFolder.Item.Path

DateTime = Format((DateTime), "DD/MM/YYYY 00:00:00")

Set colItems = objFolder.Items
For Each objItem In colItems
    'Debug.Print objItem.Name
    If objItem.Name = strFileName Then
        'Stop
        objItem.ModifyDate = DateTime
        'objItem.Name.item.ModifyDate = DateTime
        'objItem.Name.ModifyDate = DateTime
        'objItem.Name.ModifyDate = DateTime
        'objItem.Name.ModifyDate = DateTime
        'objItem.Name.ModifyDate = DateTime
        'Stop
    End If
Next


End Function


Public Function Set2244FOLDERDateTime(ByVal FileName As String, _
  ByVal TheDate As String) As Boolean
'************************************************
'PURPOSE:    Set File Date (and optionally time)
'            for a given file)
         
'PARAMETERS: TheDate -- Date to Set File's Modified Date/Time
'            FileName -- The File Name

'Returns:    True if successful, false otherwise
'************************************************

Dim FT1, FT2

If Dir(FileName, vbDirectory) = "" Then Exit Function
If Not IsDate(TheDate) Then Exit Function

Dim lFileHnd As Long
Dim lRet As Long

Dim typFileTime As FILETIME
Dim typLocalTime As FILETIME
Dim typSystemTime As SYSTEMTIME

With typSystemTime
    .wYear = Year(TheDate)
    .wMonth = Month(TheDate)
    .wDay = Day(TheDate)
    .wDayOfWeek = Weekday(TheDate) - 1
    .wHour = Hour(TheDate)
    .wMinute = Minute(TheDate)
    .wSecond = Second(TheDate)
End With

lRet = SystemTimeToFileTime(typSystemTime, typLocalTime)
lRet = LocalFileTimeToFileTime(typLocalTime, typFileTime)

'CreateDirectory "C:\Test", ByVal &H0

lFileHnd = CreateDirectory(FileName, ByVal &H0)

'lFileHnd = CreateDirectory(FileName, GENERIC_WRITE, _
    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, _
    OPEN_EXISTING, 0, 0)

'lFileHnd = CreateFile(Filename, GENERIC_WRITE, _
'    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, _
'    OPEN_EXISTING, 0, 0)

'GetFileTime lFileHnd, FT1, FT1, FT2
lFileHnd = SetFileTime(lFileHnd, ByVal 0&, ByVal 0&, _
         typFileTime)

CloseHandle lFileHnd
'SetFOLDERDateTime = lRet > 0

End Function
'--------------------------------------
'--------------------------------------
'--------------------------------------
'--------------------------------------


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
'wYear As Integer
'wMonth As Integer
'wDayOfWeek As Integer
'wDay As Integer
'wHour As Integer
'wMinute As Integer
'wSecond As Integer
'wMilliseconds As Integer
'End Type
'Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Long) As Long
'Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
'Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
'Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long


Private Function Set2222FOLDERDateTime(ByVal FileName As String, ByVal TheDate As String) As Boolean
Dim lngHandle As Long
Dim FT1 As FILETIME, FT2 As FILETIME, SysTime As SYSTEMTIME

Dim typFileTime As FILETIME

Dim typSystemTime As SYSTEMTIME

With typSystemTime
    .wYear = Year(TheDate)
    .wMonth = Month(TheDate)
    .wDay = Day(TheDate)
    .wDayOfWeek = Weekday(TheDate) - 1
    .wHour = Hour(TheDate)
    .wMinute = Minute(TheDate)
    .wSecond = Second(TheDate)
End With

CreateDirectory FileName, ByVal &H0

lngHandle = GetFileAttributes(FileName)
'MsgBox lngHandle
GetFileTime lngHandle, FT1, FT1, FT2

'Debug.Print typSystemTime.wMonth



FileTimeToLocalFileTime FT2, FT1
'FileTimeToSystemTime FT1, typFileTime 'SysTime

'Dim typFileTime As FILETIME

'lRet = SetFileTime(lngHandle, ByVal 0&, ByVal 0&, _
         typFileTime)

'MsgBox "The selected file was created on" + Str$(SysTime.wMonth) + "/" + LTrim(Str$(SysTime.wDay)) + "/" + LTrim(Str$(SysTime.wYear))
CloseHandle lngHandle
End Function

'---------------------------------------------------------------------
'END OF DAY TRY CHANGE FOLDER TIME DID NOT WORK HERE ARE ALL THE URL
'MAYBE IF DONE WHEN CREATE FOLDER MIGHT BE AN IDEA
'BUT THERE IS A PROGRAM FOR C CODE THAT DOES THEM ALL IN EXE OR SOURCE
'---------------------------------------------------------------------
'HERE URL SCRIPTOR
'---------------------------------------------------------------------
'MORNING UNTIL NIGHT EVENING 5:39 PM STARVING WITHOUT DINNER TO EAT
'---------------------------------------------------------------------
'----
'VB6 CHANGE MODIFED DATE OF FOLDER - Google Search
'https://www.google.co.uk/search?q=VB6+CHANGE+MODIFED+DATE+OF+FOLDER&rlz=1C1CHBD_en-GBGB744GB744&oq=VB6+CHANGE+MODIFED+DATE+OF+FOLDER&aqs=chrome..69i57.14080j0j7&sourceid=chrome&ie=UTF-8
'--------
'Files and Directories Free code snippets Page 2
'http://www.freevbcode.com/files-and-directories?page=2
'--------
'Folder Creation Date
'http://forums.codeguru.com/showthread.php?24993-Folder-Creation-Date
'--------
'Changes file & folder dates - CodeProject
'https://www.codeproject.com/Articles/9033/Changes-file-folder-dates
'--------
'VB Helper: HowTo: Set a file's creation, last access, and last modified times
'http://www.vb-helper.com/howto_set_file_times.html
'--------
'Any way for VBA to change file modified property
'https://www.mrexcel.com/forum/excel-questions/86056-any-way-visual-basic-applications-change-file-modified-property.html
'--------
'vbscript - VBS how can I change the DateLastModified property on a Folder - Stack Overflow
'https://stackoverflow.com/questions/10015602/vbs-how-can-i-change-the-datelastmodified-property-on-a-folder
'--------
'How to touch file - VBScript - Tek-Tips
'http://www.tek-tips.com/viewthread.cfm?qid=1372273
'--------
'msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/reference/objects/folderitem/modifydate.asp
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/reference/objects/folderitem/modifydate.asp
'--------
'List Items in the My Computer Folder using VBScript (WSH) | Computing & Technology
'https://helloacm.com/list-items-in-the-my-computer-folder-using-vbscript-wsh/
'----
'
'
