VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form ATidalX 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tidal Information..."
   ClientHeight    =   8940
   ClientLeft      =   48
   ClientTop       =   1452
   ClientWidth     =   13740
   FillStyle       =   0  'Solid
   Icon            =   "Tidal-info.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   13740
   Visible         =   0   'False
   Begin VB.Timer TIMER_SECONDS20_AFTER_HOUR_TIMER_DOOR_TALKER 
      Enabled         =   0   'False
      Left            =   6036
      Top             =   84
   End
   Begin MCI.MMControl MMControl10 
      Height          =   288
      Left            =   9492
      TabIndex        =   70
      Top             =   2784
      Visible         =   0   'False
      Width           =   2832
      _ExtentX        =   4995
      _ExtentY        =   508
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer ME_PLAY_TEXT_TO_SPEECH_TIMER 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8640
      Top             =   2772
   End
   Begin VB.Timer TIMER_UNLOAD_FORCE 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3912
      Top             =   252
   End
   Begin VB.Timer Timer_1_MINUTE 
      Interval        =   1000
      Left            =   4092
      Top             =   1176
   End
   Begin VB.ListBox List2 
      Height          =   816
      Left            =   9912
      TabIndex        =   69
      Top             =   5616
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer Timer_REINTRODUCE_DSkeybd_m 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4848
      Top             =   1080
   End
   Begin VB.Timer TIMER_WinAmpNextTrack 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5868
      Top             =   1125
   End
   Begin VB.Timer TIMER_WinAmpPrevTrack 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5544
      Top             =   1110
   End
   Begin VB.Timer TIMER_WinAmpPause 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5220
      Top             =   1125
   End
   Begin VB.Timer Timer17 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6468
      Top             =   1635
   End
   Begin VB.Timer Timer_ART_TRIP_WIRE 
      Interval        =   500
      Left            =   7224
      Top             =   1635
   End
   Begin VB.Timer TimerONCEADAY 
      Interval        =   1000
      Left            =   5700
      Top             =   510
   End
   Begin VB.Timer TimerVOLCHANGETALK2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   7200
      Top             =   540
   End
   Begin VB.Timer VOICELISTTimer 
      Interval        =   50
      Left            =   6744
      Top             =   75
   End
   Begin VB.ListBox ListVOICE 
      Height          =   624
      Left            =   11196
      Sorted          =   -1  'True
      TabIndex        =   68
      Top             =   4260
      Visible         =   0   'False
      Width           =   1188
   End
   Begin VB.Timer TimerVOLCHANGETALK 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7224
      Top             =   105
   End
   Begin VB.Timer TIMER_VoiceStreamActiveMAKESUREOFF 
      Interval        =   500
      Left            =   4476
      Top             =   708
   End
   Begin VB.Timer TIMER_SECONDS20_AFTER_HOUR_TIMER 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5700
      Top             =   84
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   1122
      Left            =   11208
      TabIndex        =   66
      Text            =   "Text2"
      Top             =   4932
      Visible         =   0   'False
      Width           =   1343
   End
   Begin VB.Timer VoiceMP3SayTimer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6864
      Top             =   2070
   End
   Begin VB.Timer TimerSimulateStringVoice 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4776
      Top             =   60
   End
   Begin VB.Timer Timer14 
      Interval        =   100
      Left            =   7248
      Top             =   2070
   End
   Begin VB.Timer TimerSunRiseSetPercent 
      Interval        =   500
      Left            =   4824
      Top             =   525
   End
   Begin VB.Timer Timer4 
      Interval        =   10000
      Left            =   6840
      Top             =   1665
   End
   Begin VB.Timer WinampFastForwardTimer 
      Interval        =   1
      Left            =   6240
      Top             =   540
   End
   Begin VB.Timer TimerKeyInterceptSlow 
      Interval        =   10
      Left            =   6048
      Top             =   1635
   End
   Begin VB.Timer Timer16 
      Interval        =   1000
      Left            =   6144
      Top             =   2445
   End
   Begin VB.Timer PassLockTimer 
      Interval        =   500
      Left            =   4656
      Top             =   1635
   End
   Begin VB.Timer TimerKeyIntercept 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6756
      Top             =   525
   End
   Begin VB.Timer ShellWinampLoggLoader 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5304
      Top             =   570
   End
   Begin MCI.MMControl MMControl9 
      Height          =   336
      Left            =   9480
      TabIndex        =   63
      Top             =   2436
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   593
      _Version        =   393216
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer WinAmpfadeout_Timer 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   5268
      Top             =   105
   End
   Begin VB.Timer VCode5_TIMER 
      Interval        =   100
      Left            =   8640
      Top             =   2340
   End
   Begin VB.Timer TimerVideoKill 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7296
      Top             =   2490
   End
   Begin VB.Timer TimerNotePadYes 
      Interval        =   100
      Left            =   6564
      Top             =   2430
   End
   Begin VB.Timer TimerDoubleVolume 
      Interval        =   10
      Left            =   8160
      Top             =   1092
   End
   Begin MCI.MMControl MMControl7 
      Height          =   336
      Left            =   9480
      TabIndex        =   58
      Top             =   1836
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   593
      _Version        =   393216
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer ShellKeyLoad 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6408
      Top             =   1170
   End
   Begin VB.Timer FileSaveFileExit 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6804
      Top             =   1245
   End
   Begin VB.Timer Timer15 
      Interval        =   1000
      Left            =   8196
      Top             =   660
   End
   Begin VB.Timer TimerIceViewKill 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6936
      Top             =   2505
   End
   Begin VB.Timer Timer13 
      Interval        =   500
      Left            =   8664
      Top             =   1488
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   5700
      Top             =   2460
   End
   Begin VB.Timer SliderTimer 
      Interval        =   100
      Left            =   8664
      Top             =   1884
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   192
      Left            =   0
      TabIndex        =   53
      Top             =   552
      Visible         =   0   'False
      Width           =   3144
      _ExtentX        =   5546
      _ExtentY        =   339
      _Version        =   393216
      Max             =   100
      TickFrequency   =   10
   End
   Begin VB.Timer VolGetTimer 
      Interval        =   400
      Left            =   8676
      Top             =   1008
   End
   Begin VB.Timer TimerVolLab 
      Interval        =   100
      Left            =   8664
      Top             =   564
   End
   Begin MCI.MMControl MMControl6 
      Height          =   336
      Left            =   9480
      TabIndex        =   52
      Top             =   1536
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   593
      _Version        =   393216
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer VoiceMP3Start_TIMER 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8676
      Top             =   108
   End
   Begin VB.Timer Timer_CountD1Text 
      Interval        =   200
      Left            =   5256
      Top             =   2490
   End
   Begin VB.Timer Timer9 
      Interval        =   10000
      Left            =   5628
      Top             =   2085
   End
   Begin VB.Timer Timer8 
      Interval        =   10
      Left            =   7776
      Top             =   2376
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6060
      Top             =   2040
   End
   Begin VB.Timer Timer13PauseVoice 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5328
      Top             =   1650
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   432
      Left            =   4728
      TabIndex        =   45
      Top             =   5832
      Width           =   3150
   End
   Begin VB.Timer Timer12 
      Interval        =   100
      Left            =   8208
      Top             =   264
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Tommorow."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2148
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4260
      Width           =   1920
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Today."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4260
      Width           =   2076
   End
   Begin VB.Timer Timer11CPU 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5268
      Top             =   2130
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7224
      Top             =   1215
   End
   Begin VB.Timer Timer9ColorCycle 
      Interval        =   100
      Left            =   7812
      Top             =   3180
   End
   Begin VB.Timer Timer8WinAmpVol 
      Interval        =   50
      Left            =   7788
      Top             =   2784
   End
   Begin VB.Timer Timer7WinAmpVolIDE 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5604
      Top             =   1650
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11328
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "Tidal-info.frx":0E42
      Top             =   6204
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7764
      Top             =   1908
   End
   Begin VB.Timer Timer5 
      Interval        =   50
      Left            =   7764
      Top             =   1476
   End
   Begin VB.Timer ASYNC_Key_Press_Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7764
      Top             =   1080
   End
   Begin VB.Timer Timer3 
      Interval        =   50
      Left            =   7764
      Top             =   660
   End
   Begin VB.PictureBox TextToSpeech1 
      Enabled         =   0   'False
      Height          =   1155
      Left            =   9960
      ScaleHeight     =   1104
      ScaleWidth      =   1092
      TabIndex        =   13
      Top             =   4416
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MCI.MMControl MMControl4 
      Height          =   372
      Left            =   9480
      TabIndex        =   6
      Top             =   624
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   656
      _Version        =   393216
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MCI.MMControl MMControl3 
      Height          =   336
      Left            =   9480
      TabIndex        =   5
      Top             =   24
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   593
      _Version        =   393216
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MCI.MMControl MMControl2 
      Height          =   336
      Left            =   9480
      TabIndex        =   4
      Top             =   972
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   593
      _Version        =   393216
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   7764
      Top             =   216
   End
   Begin VB.DirListBox Dir2 
      Enabled         =   0   'False
      Height          =   504
      Left            =   8316
      TabIndex        =   3
      Top             =   5460
      Visible         =   0   'False
      Width           =   1512
   End
   Begin MCI.MMControl MMControl1 
      Height          =   336
      Left            =   9480
      TabIndex        =   2
      Top             =   336
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   593
      _Version        =   393216
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.PictureBox DirectSS1 
      Enabled         =   0   'False
      Height          =   240
      Left            =   4476
      ScaleHeight     =   192
      ScaleWidth      =   288
      TabIndex        =   1
      Top             =   3324
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.DirListBox Dir1 
      Enabled         =   0   'False
      Height          =   720
      Left            =   8316
      TabIndex        =   0
      Top             =   4704
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4452
      Top             =   276
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   168
      Left            =   -36
      TabIndex        =   17
      Top             =   4092
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   296
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBarVol1 
      Height          =   156
      Left            =   0
      TabIndex        =   41
      Top             =   252
      Width           =   3288
      _ExtentX        =   5800
      _ExtentY        =   296
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBarLock 
      Height          =   252
      Left            =   -12
      TabIndex        =   43
      Top             =   528
      Width           =   3288
      _ExtentX        =   5800
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBarCPU 
      Height          =   252
      Left            =   4464
      TabIndex        =   46
      Top             =   3000
      Width           =   3228
      _ExtentX        =   5694
      _ExtentY        =   466
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   168
      Left            =   -36
      TabIndex        =   49
      Top             =   3072
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   296
      _Version        =   393216
      Appearance      =   0
      Max             =   1000
      Scrolling       =   1
   End
   Begin MCI.MMControl MMControl5 
      Height          =   336
      Left            =   9480
      TabIndex        =   50
      Top             =   1260
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   593
      _Version        =   393216
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MCI.MMControl MMControl8 
      Height          =   336
      Left            =   9480
      TabIndex        =   59
      Top             =   2136
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   593
      _Version        =   393216
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MSComctlLib.ProgressBar ProgressBarVol2 
      Height          =   180
      Left            =   0
      TabIndex        =   67
      Top             =   408
      Width           =   3288
      _ExtentX        =   5800
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MCI.MMControl MMControl11 
      Height          =   288
      Left            =   9492
      TabIndex        =   71
      Top             =   3036
      Visible         =   0   'False
      Width           =   2832
      _ExtentX        =   4995
      _ExtentY        =   508
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MCI.MMControl MMControl12 
      Height          =   288
      Left            =   9492
      TabIndex        =   72
      Top             =   3276
      Visible         =   0   'False
      Width           =   2832
      _ExtentX        =   4995
      _ExtentY        =   508
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MCI.MMControl MMControl13 
      Height          =   288
      Left            =   9492
      TabIndex        =   73
      Top             =   3540
      Visible         =   0   'False
      Width           =   2832
      _ExtentX        =   4995
      _ExtentY        =   508
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MCI.MMControl MMControl14 
      Height          =   288
      Left            =   9492
      TabIndex        =   74
      Top             =   3804
      Visible         =   0   'False
      Width           =   2832
      _ExtentX        =   4995
      _ExtentY        =   508
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "000.0000000%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   276
      Left            =   0
      TabIndex        =   65
      ToolTipText     =   "Hitt Button on Left Side and Right Side --  Talker"
      Top             =   5004
      Width           =   4080
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   456
      Left            =   3120
      TabIndex        =   64
      Top             =   1560
      Width           =   636
   End
   Begin VB.Label Label39 
      BackColor       =   &H000000FF&
      Caption         =   "Label39"
      Enabled         =   0   'False
      Height          =   312
      Left            =   4488
      TabIndex        =   62
      Top             =   2160
      Width           =   624
   End
   Begin VB.Label Label38 
      BackColor       =   &H0000C000&
      Caption         =   "Label38"
      Enabled         =   0   'False
      Height          =   312
      Left            =   4512
      TabIndex        =   61
      Top             =   2604
      Width           =   684
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   468
      Left            =   3120
      TabIndex        =   60
      Top             =   1080
      Width           =   636
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Idle"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   57
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   5988
      Visible         =   0   'False
      Width           =   576
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "00:00:00"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   708
      TabIndex        =   56
      Top             =   5976
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Active"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   276
      Left            =   1944
      TabIndex        =   55
      ToolTipText     =   "View keyboard and Mouse Log With Notepad"
      Top             =   5976
      Visible         =   0   'False
      Width           =   864
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "00:00:00"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   276
      Left            =   2748
      TabIndex        =   54
      ToolTipText     =   "F12 to Command Prompt - Disable"
      Top             =   5976
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pills"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   216
      Left            =   8004
      TabIndex        =   51
      ToolTipText     =   "Click to Speedup the Time Double Click Resets"
      Top             =   6300
      Visible         =   0   'False
      Width           =   1788
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "System UpTime"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   216
      Left            =   7980
      TabIndex        =   48
      ToolTipText     =   "Click to Speedup the Time Double Click Resets"
      Top             =   6060
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label LabelCPU 
      Alignment       =   2  'Center
      BackColor       =   &H00680202&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   4968
      TabIndex        =   47
      Top             =   3324
      Width           =   468
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00680202&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   3264
      TabIndex        =   44
      Top             =   528
      Width           =   564
   End
   Begin VB.Label LblVol 
      Alignment       =   2  'Center
      BackColor       =   &H00680202&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   3264
      TabIndex        =   42
      Top             =   276
      Width           =   564
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Equinox"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   12
      TabIndex        =   40
      ToolTipText     =   "Click to Speedup the Time Double Click Resets"
      Top             =   2820
      Width           =   3948
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pay Day"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   0
      TabIndex        =   39
      ToolTipText     =   "Click to Speedup the Time Double Click Resets"
      Top             =   2556
      Width           =   3984
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Days This Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   -12
      TabIndex        =   38
      ToolTipText     =   "Click to Speedup the Time Double Click Resets"
      Top             =   2292
      Width           =   3984
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Week No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   -12
      TabIndex        =   37
      ToolTipText     =   "Click to Speedup the Time Double Click Resets"
      Top             =   2028
      Width           =   3984
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "12Jan hh:MM:SS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   2112
      TabIndex        =   36
      Top             =   3816
      Width           =   1896
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next First Quarter"
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
      Height          =   252
      Left            =   -12
      TabIndex        =   35
      Top             =   3816
      Width           =   2028
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Moon Luminosity"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   -12
      TabIndex        =   34
      Top             =   3252
      Width           =   2064
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   2856
      TabIndex        =   33
      ToolTipText     =   "F12 to Command Prompt - Disable"
      Top             =   5604
      Width           =   1236
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   2076
      TabIndex        =   32
      ToolTipText     =   "View keyboard and Mouse Log With Notepad"
      Top             =   5604
      Width           =   804
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Mouse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   996
      TabIndex        =   31
      Top             =   5604
      Width           =   1068
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Mouse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   -12
      TabIndex        =   30
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   5604
      Width           =   1008
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "DOD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   288
      Left            =   -12
      TabIndex        =   29
      ToolTipText     =   "Talker"
      Top             =   5292
      Width           =   4080
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "000.0000000%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   288
      Left            =   2112
      TabIndex        =   28
      ToolTipText     =   "Click Hitt Button For Clipboard Info Chuck and More Info when Shift Key With Button"
      Top             =   3252
      Width           =   1932
   End
   Begin VB.Label LABEL_TIME 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "12:12:12PM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   468
      Left            =   -12
      TabIndex        =   27
      ToolTipText     =   "Click to Speedup the Time Double Click Resets"
      Top             =   1560
      Width           =   3120
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "MON 19 OCT04"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   468
      Left            =   -12
      TabIndex        =   26
      Top             =   1092
      Width           =   3120
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "000.0000000%"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   216
      Left            =   4464
      TabIndex        =   25
      Top             =   5520
      Visible         =   0   'False
      Width           =   3720
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "000.0000000%"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   228
      Left            =   4464
      TabIndex        =   24
      Top             =   5268
      Visible         =   0   'False
      Width           =   3720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "000.0000000% SUNSET"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   -12
      TabIndex        =   23
      ToolTipText     =   "Hitt Button on Left Side and Right Side --  Talker"
      Top             =   4740
      Width           =   4080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "000.0000000%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   -12
      TabIndex        =   22
      ToolTipText     =   "Hitt Button on Left Side and Right Side --  Talker"
      Top             =   4476
      Width           =   4116
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "*-- Matta Ratta's --*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   276
      Left            =   -36
      TabIndex        =   21
      Top             =   0
      Width           =   3936
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "000.0000000%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   -12
      TabIndex        =   20
      Top             =   780
      Width           =   3984
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "12Jan hh:MM:SS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   2124
      TabIndex        =   19
      Top             =   3552
      Width           =   1896
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Now First Quarter"
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
      Height          =   252
      Left            =   -12
      TabIndex        =   18
      Top             =   3552
      Width           =   2028
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Enabled         =   0   'False
      Height          =   288
      Left            =   6504
      TabIndex        =   12
      Top             =   4068
      Width           =   828
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Enabled         =   0   'False
      Height          =   336
      Left            =   6504
      TabIndex        =   11
      Top             =   3348
      Width           =   828
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      Enabled         =   0   'False
      Height          =   336
      Left            =   6504
      TabIndex        =   10
      Top             =   3708
      Width           =   828
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      Enabled         =   0   'False
      Height          =   288
      Left            =   4476
      TabIndex        =   9
      Top             =   4908
      Width           =   1812
   End
   Begin VB.Label Label15 
      Caption         =   "Label15"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4452
      TabIndex        =   8
      Top             =   3696
      Width           =   1980
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Enabled         =   0   'False
      Height          =   420
      Left            =   4464
      TabIndex        =   7
      Top             =   4452
      Width           =   2748
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "Exit"
   End
   Begin VB.Menu Mnu_VB_ME_Run 
      Caption         =   "VB ME && RUN"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_ADMINISTRATOR 
      Caption         =   "ADMIN"
   End
   Begin VB.Menu Mnu_MINIMIZE 
      Caption         =   "Minimize"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_ 
      Caption         =   "Menu"
      Begin VB.Menu Mnu_Explore 
         Caption         =   "Explore Tidal Folder"
      End
      Begin VB.Menu mnu_Switches 
         Caption         =   "Switches"
      End
      Begin VB.Menu Mnu_SW_Visits 
         Caption         =   "SW Visits"
      End
      Begin VB.Menu Mnu_SigLine 
         Caption         =   "NotePad SigLine"
      End
      Begin VB.Menu Mnu_TimeLogg 
         Caption         =   "NOTEPAD TIME KEY PRESSED LOGG"
      End
      Begin VB.Menu Mnu_Explorer_Logg 
         Caption         =   "Explorer Logg Folder"
      End
      Begin VB.Menu Mnu_FavProgs 
         Caption         =   "Run Fav Progs ..."
      End
      Begin VB.Menu Mnu_KeyLog 
         Caption         =   "View KeyLogger"
      End
      Begin VB.Menu mnu_Excute 
         Caption         =   "Executie"
      End
      Begin VB.Menu Mnu_WebPage 
         Caption         =   "WebPage Edit"
      End
      Begin VB.Menu Mnu_Norton_Test 
         Caption         =   "Norton Sig Test"
      End
      Begin VB.Menu MNU_WinampFadeOut 
         Caption         =   "Winamp Fade an out"
      End
      Begin VB.Menu Mnu_ShowURL 
         Caption         =   "Show Url Page"
      End
      Begin VB.Menu Mnu_HideURL 
         Caption         =   "Hide Url Page"
      End
      Begin VB.Menu Mnu_Note_Url_Big 
         Caption         =   "NoteBook Url Big"
      End
      Begin VB.Menu Mnu_NotebookURL 
         Caption         =   "NoteBook URL"
      End
      Begin VB.Menu Mnu_ClipURL 
         Caption         =   "Clip URL"
      End
      Begin VB.Menu Mnu_VolMax 
         Caption         =   "Vol Max"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_VolMin 
         Caption         =   "Vol Min 2hrs"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_VolMinMusic 
         Caption         =   "Vol Min% Music"
      End
      Begin VB.Menu Mnu_Recycle 
         Caption         =   "Recycle Bin"
      End
      Begin VB.Menu Mnu_Knocker1 
         Caption         =   "NotePad Knocker Boy1"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_Knocker2 
         Caption         =   "NotePad Knocker Boy2"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_Knocker3 
         Caption         =   "NotePad Knocker Boy3"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_Knocker4 
         Caption         =   "NotePad Knocker Boy4"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_ExtremeLockLogg 
         Caption         =   "Extreme Lock Logg"
      End
      Begin VB.Menu Mnu_SysUpTime 
         Caption         =   "Sys UPTime"
      End
      Begin VB.Menu Mnu_CountD1 
         Caption         =   "New Year Time count Down"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_Idel 
         Caption         =   "Idle Active Logg Day"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_Idle_Now 
         Caption         =   "Idle Active Logg Now"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_EditMallSound 
         Caption         =   "Edit Sound Effect Mall"
      End
      Begin VB.Menu Mnu_SetTime3PM 
         Caption         =   "SET TIME BEFORE 3PM"
      End
      Begin VB.Menu Mnu_SetTimesMIDNIGHT 
         Caption         =   "SET TIME BEFORE MIDNIGHT"
      End
      Begin VB.Menu Mnu_SetTimesSUNNOON 
         Caption         =   "SET TIME BEFORE NOON"
      End
      Begin VB.Menu Mnu_SetTimesSUNSET 
         Caption         =   "SET TIME BEFORE PREVIOUS SUNSET"
      End
      Begin VB.Menu Mnu_SetTime 
         Caption         =   "SET TIME JUST BEFORE THIS HOUR"
      End
      Begin VB.Menu Mnu_SetTimeHOURFORWARD 
         Caption         =   "SET TIME HOUR FORWARD"
      End
      Begin VB.Menu Mnu_Time_Syne 
         Caption         =   "Time_Sync"
      End
      Begin VB.Menu Mnu_SimulateStringVoice 
         Caption         =   "SIMULATE STRING VOICE"
      End
   End
   Begin VB.Menu Mnu_Web 
      Caption         =   "Web"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_QuickWeb 
      Caption         =   "WebPages"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_ShowWeb 
      Caption         =   "Show Webs"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_Terminate 
      Caption         =   "Terminate Webs"
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_N 
      Caption         =   "News"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_NewsClipText 
      Caption         =   "Clip To Text"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_Tidal 
      Caption         =   "Tidal"
      Begin VB.Menu Mnu_VB6 
         Caption         =   "VB6 Shortcuts"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_VBISIDE 
         Caption         =   "Tidal Run In IDE"
         Checked         =   -1  'True
      End
      Begin VB.Menu Mnu_ELock 
         Caption         =   "Exteme Lock"
      End
      Begin VB.Menu Mnu_ELock_Off 
         Caption         =   "Exteme Lock Off"
      End
      Begin VB.Menu Mnu_Equinox 
         Caption         =   "Equinox"
      End
      Begin VB.Menu MNU_SPACE21 
         Caption         =   "----------"
      End
      Begin VB.Menu MNU_TALKER_EVERY_MINUTE_FOR_AN_HOUR 
         Caption         =   "TALKER EVERY MINUTE FOR AN HOUR"
      End
      Begin VB.Menu Mnu_Say_Time 
         Caption         =   "Say Time"
      End
      Begin VB.Menu Mnu_Say_Sunrise 
         Caption         =   "Say Sun Rise"
      End
      Begin VB.Menu Mnu_Say_SunSet 
         Caption         =   "Say Sun Set"
      End
      Begin VB.Menu Mnu_Say_SolarNoon 
         Caption         =   "Say Solar Noon"
      End
      Begin VB.Menu MNU_SAYMOON 
         Caption         =   "Say Moon 8 Phase Name"
      End
      Begin VB.Menu MNU_SAY_MOON 
         Caption         =   "Say Moon %"
      End
      Begin VB.Menu MNU_SAY_DAYLIGHT 
         Caption         =   "Say DayLight --And Or-- DarkNess %"
      End
      Begin VB.Menu MNU_VOX_LIST 
         Caption         =   "SHOW VOXLIST"
      End
      Begin VB.Menu Mnu_ClipBoard_Info_Block_Of_Moon 
         Caption         =   "Clipboard info Block Of Moon"
      End
   End
   Begin VB.Menu Mnu_Key_Hook 
      Caption         =   "Key Hook && Cmd"
      Begin VB.Menu Mnu_KeyCleaning 
         Caption         =   "Block all Keys Toggle KeyBoard Cleaning"
      End
      Begin VB.Menu MNU_SPACE1 
         Caption         =   "----------"
      End
      Begin VB.Menu Mnu_KHook 
         Caption         =   "Show Key Hook"
      End
      Begin VB.Menu Mnu_KeyHook 
         Caption         =   "Load an Use Key Hook"
      End
      Begin VB.Menu MNU_SPACE2 
         Caption         =   "----------"
         Index           =   1
      End
      Begin VB.Menu Mnu_KeyLoggFolder 
         Caption         =   "KeyLoggFolder"
      End
      Begin VB.Menu Mnu_KeyLogText 
         Caption         =   "KeyLogText"
      End
      Begin VB.Menu MNU_SPACE3 
         Caption         =   "----------"
      End
      Begin VB.Menu Mnu_Expand_W 
         Caption         =   "Expand Width"
      End
      Begin VB.Menu Mnu_BackLash 
         Caption         =   "Control / = BackSlash"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_CountD 
         Caption         =   "Stop"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_LoadVBanKill 
         Caption         =   "Load Vb and Kill Process"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MNU_Menu22 
      Caption         =   "WinAmp Vol Opt"
      Begin VB.Menu MNU_SPEECH_OFF 
         Caption         =   "SPEECH OFF"
         Checked         =   -1  'True
      End
      Begin VB.Menu MNU_NO_MATCH_WINAMP_VOL 
         Caption         =   "NO MATCH WINAMP VOL"
         Checked         =   -1  'True
      End
      Begin VB.Menu Mnu_NoJumpUpVol 
         Caption         =   "NO JUMP UP VOL -- LEAVE AS IS MASTER VOL"
      End
      Begin VB.Menu MNU_MAXVOL_VOICE 
         Caption         =   "MAX ALL VOICE VOL"
      End
      Begin VB.Menu MNU_DONT_ADJUST_VOL_ON_VOICE 
         Caption         =   "DONT ADJUST VOL ON VOICE"
         Checked         =   -1  'True
      End
      Begin VB.Menu MNU_NO_SET_VOL_NONE_SET_AS_1 
         Caption         =   "Dont Vol To None Set Min to 1"
         Checked         =   -1  'True
      End
      Begin VB.Menu MNU_LATENIGHTOVERRIDE 
         Caption         =   "OVERRIDE LATE NIGHT COOL OFF"
      End
      Begin VB.Menu MNU_HOT_KEY_LIST 
         Caption         =   "HOT KEY LIST"
      End
      Begin VB.Menu MNU_BOTH_VOL_SAME 
         Caption         =   "BOTH VOLS SAME"
      End
      Begin VB.Menu MNU_NO_PAUSE 
         Caption         =   "NOT PAUSE WINAMP"
      End
      Begin VB.Menu MNU_STOP_WINAMP_ON_THE_HOUR 
         Caption         =   "STOP WINAMP ON THE HOUR"
      End
      Begin VB.Menu MNU_NOT_DOOR_TALK 
         Caption         =   "NOT DOOR TALK"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu MNU_VOX_LIST2 
      Caption         =   "VoX Text"
   End
   Begin VB.Menu Mnu_Voice 
      Caption         =   "Voice Control"
   End
   Begin VB.Menu MNU_JOYPAD_SOFTWARE_STATUS 
      Caption         =   "** Joypad Software Not Installed **"
   End
   Begin VB.Menu MNU_JOY_PADDER 
      Caption         =   "JOY-PAD-DX"
   End
   Begin VB.Menu MNU_JOY_PADDER_1 
      Caption         =   "JOY-PAD-1"
   End
   Begin VB.Menu MNU_VOL_100 
      Caption         =   "VOL 100%"
   End
   Begin VB.Menu Mnu_About 
      Caption         =   "About"
      Begin VB.Menu MNU_ENDER_MNU 
         Caption         =   "EXIT WITH END ABRUPTLY __ RELOAD PROBLEM OVER"
      End
   End
   Begin VB.Menu MNU_TEST_FORM 
      Caption         =   "T"
      Visible         =   0   'False
      Begin VB.Menu MNU_WINAMP 
         Caption         =   "WIN"
         Index           =   1
      End
      Begin VB.Menu MNU_WINAMP 
         Caption         =   "WIN"
         Index           =   2
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MNU_FILE_STUCK_IN_USE 
      Caption         =   "FILE STUCK IN USE"
   End
   Begin VB.Menu MNU_VERSION 
      Caption         =   "Version"
   End
End
Attribute VB_Name = "ATidalX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim START_DOOR_TALKER
Dim FORM_HAS_GOT_TO_UNLOAD
Dim CHECK_TIME As Date
Dim RS232_LOGGER_FILE_TIME_V2_HOUR
Dim RS232_LOGGER_FILE_TIME_V2_MINUTE
Dim RS232_LOGGER_FILE_TIME_V2_SECOND
Dim RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_1 As Date
Dim RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_2 As Date

Dim RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD_RETRIG_SOUND_2

Dim SET_MMControl9_TO_GO_FLAG
Dim OLD_RS232_LOGGER_FILE_TIME_V2_MINUTE
Dim RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD_RETRIG_SOUND

Dim VAR_TALKER_EVERY_MINUTE_FOR_AN_HOUR As Date
Dim OLD_X9
Dim O_Minute_Now '' FOR TALK TIME EVERY MINUTE SWITCH
Dim OLD_MoonLabel

Dim RS232_LOGGER_FILE_TIME_V1_OLD
Dim TIME_TALKER_1_HOUR_3
Dim TIME_TALKER_2_MIN__3
Dim TIME_TALKER_3_SEC__3
Dim RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_1_OLD
Dim OLD_SPEAK_DOOR
Dim OLD_RS232_LOGGER_FILE_TIME_SECOND
Dim OLD_RS232_LOGGER_FILE_TIME_MINUTE
Dim OLD_RS232_LOGGER_FILE_TIME_HOUR
Dim OLD_RS232_LOGGER_FILE_TIME_CLOSE

Dim OLD_RS232_LOGGER_FILE_TIME
Dim RS232_LOGGER_FILE_TIME_CLOSE As Date
Dim RS232_LOGGER_FILE_TIME_OPEN  As Date
Dim RS232_LOGGER_FILE_TIME As Date
Dim OLD_RS232_LOGGER_FILE_VALUE As Date
Dim RS232_LOGGER_FILE_TIME_V1
Dim RS232_LOGGER_FILE_TIME_V2
Dim Form_Load_Done_For_Timer14

Dim EQUINOX_PERCENT

Public ALL_FORM_REQUEST_TO_END

Dim R2

Dim O_BONG_DAY_NOW
Dim X_COUNT_ME_SOUND
Dim SOUND_FILE_ME_SET_ONCE

' ---------------------------------------------------
' PROGRAMMER THE JOYPAD TALK TIME AND DATE ON BUTTONS WORKING AND
' NEW SOUND EFFECT WORD TEXT TO SPEECH DONE CALLED "ME"
' USING 5 CHANNELS OF HOLDING SAME SOUND BUT HAS TO BE IN 5 COPY OF FILE
' GOT IDEA FROM MONDAY LOTTERY LONG BEFORE HAD GOOD SOUND EFFECT WORD ME
' WHEN SELECTING A BALL BUT UNABLE TO FIND THAT PRECISE ONE
' FOUND AND ENGINE ONLINE TO DO IT BUT TEXT TO SPEECH EVEN ON LINE DOESN'T
' GO TO SOUND CARD GOOD SO USE DICTAPHONE RECORDER THAT PART AND MAKING WAVE
' FILE SO SMALL NOT REALLY THAT BE BUT COMPLAIN MY FILE NAME TOO LONG HAD TO
' MIX A BIT FROM ANOTHER SOUND ON QUILEY TO GET IT DONE
' HAVING FIVE SOUNDS MULTI TASKING OVERLAP EACH OTHER OR ELSE IT IS JUST
' PLAY ONE AFTER ANOTHER MAKE A SUPRISING GOOD SOUND EFFECT
' ---------------------------------------------------
' Thu 19-Apr-2018 09:44:40 BEGGINER
' Thu 19-Apr-2018 13:44:40 4 HOUR+
' ---------------------------------------------------


' RESPONCES ARE EVENT DRIVE SO A LITTLE MORE QUICK
' ---------------------------------------------------
' ALSO RE-WROTE CODE FOR APPEND FILE_IN_USE_DELAY FOR ALL FILE OPERATIONS
' ---------------------------------------------------
' Thu 19-Apr-2018 01:20:25 BEGINNER
' Thu 19-Apr-2018 07:58:44 7 HOUR -
' ---------------------------------------------------


'---------------------------------------------------
' PROGRAMMER THE JOYPAD UP ANOTHER LEVEL THE F710 BUTTONS ARE SWAPPED FROM VERSION BEFORE
' VOLUME CONTROL THE BOTTOM FRONT ROW ARE NOT FOUND ANYMORE ON THE OLD JOYPAD CONTROLLER
' RESPONCES ARE EVENT DRIVE SO A LITTLE MORE QUICK
'---------------------------------------------------
' ALSO RE-WROTE CODE FOR APPEND FILE_IN_USE_DELAY FOR ALL FILE OPERATIONS
'---------------------------------------------------
' Thu 19-Apr-2018 01:20:25 BEGINNER
' Thu 19-Apr-2018 07:58:44 7 HOUR -
'---------------------------------------------------

'---------------------------------------------------
' PROGRAMMER THE MOON AND SUN PERCENT TIME TALKER
' GIVING FREQUENCEY OF TIME TALK NOT TOO FREQUENT USING MOD
' AND TIDYIED UP THE LAYOUT OF FORM
' AND SORTED PROBLEM OF MAKE PUBLIC VARIABLE AND REF TO THE FORM WHEN UNLOADING EXIT
' NOT BOUNCING BACK IN BY NOT UNLOAD PROPERLY
' WHEN REFEANCE A VISUAL OBJECT OBSERVE
' MORE WORK ON TIDY FORM LAYOUT + ADD EUQINOX DISPLAY PERCENT
'---------------------------------------------------
' Fri 04-May-2018 14:35:50
' Fri 04-May-2018 17:32:48 __ 2 HOURS 57 MINUTES
' Fri 04-May-2018 18:54:48 __ 4 HOURS 19 MINUTES EXTENDED
'---------------------------------------------------

'---------------------------------------------------
' PROGRAMMER WORK ON VB PROJECT UPDATER
' IN 5 PROJECT ALL WITH ONE FORM
' AND ONLY REQUIRE ONE LINE IN MAIN STARTING PROJECT
' TO GET IT STARTED
' AND THE SYNCRONIZER PROGRAM I MADE NOW HAS 2 PROJECTS
' AND KEEP THEM SYNCRONIZED THAT WAY
'---------------------------------------------------
' Sat 05-May-2018 14:42:13
' Sat 05-May-2018 17:00:00 __ 2 HOURS 18 MINUTES
'---------------------------------------------------


Public O_TALKTIME
Public EXIT_TRUE

Dim FILENAME_IN_USE_CHECK As String

Dim OLD_SECOND_NOW
Dim XZ_TEST_VOICE_COUNTDOWN
Dim OLD_GETTING_REPEAT_IF_ANOTHER_SPEAK
Dim TIMER_TO_HALT_RERUN_ ' FAULT WORK AROUND FIXER

Dim GMTBST
'---------------------------------------
'C:\Windows\System32\MSCOMM32.OCX
'DECIDE KEEPER THIS CONTROL IS FOR LISTVIEW AND OTHER AND THE ONE HAS THE UPDATE FAULT HERE
'----
'Microsoft ActiveX Common Control MSCOMCTL.OCX Security Update Problem with Registration Affects Treeview and ListBox in Microsoft Access, Office, VB6
'http://www.fmsinc.com/microsoftaccess/controls/mscomctl/
'----
'E:\01 Start Menu\#_4-ASUS-GL522VW\Programs\Startup-01-Delayed\4-ASUS-GL522VW--MATT 01\BAT 03 REGISTRY AT BOOTER.BAT
'
'MICROSOFT COMM CONTROL 6.0
'REMOVER SAVE SPACE
'Thu 16 February 2017 05:53:06----------
'---------------------------------------
'C:\Windows\System32\MSFLXGRD.OCX
'MICROSOFT FLEX GRID CONTROL 6.0
'REMOVER SAVE SPACE
'Thu 16 February 2017 05:53:06----------
'---------------------------------------
'C:\Windows\System32\MSFLXGRD.OCX
'MICROSOFT FLEX GRID CONTROL 6.0
'REMOVER SAVE SPACE
'Thu 16 February 2017 05:53:06----------
'---------------------------------------
'C:\Windows\SysWOW64\ieframe.dll
'C:\Windows\SysWOW64\ieframe.dll
'MICROSOFT INTERNET CONTROLS
'REMOVER SAVE SPACE __ .DLL
'Thu 16 February 2017 05:53:06----------
'---------------------------------------
Dim TIME_ONE_SECOND_COUNTER

Dim O_ABH
Dim FIRST_RUN_SUN_PERCENT_TALKER
Dim VCodeTT1, VCodeTT2

'TOP DELCLARE
Dim XVB_DATE

'BOOKMARK FLAG POSITION TO WATCH
Dim FIRST_RUN_SUN_PERCENT_TALKER_LATCH_UP
Dim REQUEST_NOT_HAVE_Angle_Of_Illumination

'----------------------------
'MINIMIZE VB IDE AT RUN
'CAN BE ENABLED AGAIN IF WANT
'----------------------------
Dim ShiftPressed, Mnu_Solar_System
Dim SunResult_Reverse
Dim SunRise_Not_Talk_But_Day_Light_Percent_DblClick
Dim SunRise_Not_Talk_But_Day_Light_Percent_User_Button

'Dim SunRise_Bong_DblClick_Var
Dim O_DARK_OR_LIGHT
Dim RunOnce_Equinox

Dim USER_REQUEST_TEST_MOON_CHANGE_STATE_AND_TALK

Dim FLAG_TO_MOON_TALK
Dim VAR_REINTRODUCE_DSkeybd_M

Dim DAY_NOW_OLD

Dim ME_POSITION_BEEN_SET_ONCE
Dim MOON_MOD_COUNTER
Dim DARK_OR_MOD_COUNTER
Dim DarkLightOLD



'----------------------------------
'Sub TimerSunRiseSetPercent_Timer()

Dim TEST3
Dim DARK_OR_LIGHT, DARK_OR_LIGHT_OLD

'----------------------------------

'WHAT TO DO ABOUT THIS
'MENU.Timer2_Timer



Dim oButton
Dim MHDR, MHDR_ODD_MIN, ORT2, RTV1, RTV2

Dim MHDR_NOW_DATE
Dim OLDi, iTRIP, XHH
Dim TALKEQU1, TALKEQU2, YYFLAG1, YYFLAG2
Dim FLAGMOONSAYMORE


'## WORK NOW
'Open "D:\TEMP\SEASONS.txt" For Input As #ff
Dim Idate As Date

Dim OVOLTAG2
Dim ISSUECOMMTOSTOPWINAMP
Dim VOICECOMMANDSTART
Dim XX1, XX2, XX3, XX4


Dim XX20, XX21, OVOICE

'YES VOL2 IS NOT MUSIC
'If VolumeSet2 > VolumeSet1 Then Exit Sub
Dim OVOLTAG, ForceSayVol
Dim EQISP, oEQUIDATE1, oRXT, RXT
'If WinAmpStoppedPlayTime > 0 Then
Dim DOGBARKCOMMO, OVOLLAB, TEST1, NOVOLSAYASSTART
Dim oWinAmpChkIsPlaying
Public KATTIMER

Dim oVoiceToSpeak
'WORKING
'Private Sub TimerVolLab_Timer()

Dim VUMPATH1, VUMPATH2
Dim OHOUR5
Dim Tool4B, Tool4A
Dim ONHOURCOUNTDOWN
'
'USEFUL
'Public Sub TimerKeyIntercept_Timer()
Dim SUNSAYSTRING

'OFF TIMES
Dim OXAZY()
Dim XAZY()
Dim XZAP()

'Dim WINAMPBackOnAfterVoiceRequired
Dim Rect4 As RECT


'JUST REMOVED THIS CLEAN IT OUT
'''Call VoiceWait



'CHK HERE BEEN MODIFIED OUT
'Private Sub WinampFastForwardTimer_Timer()



'If VolumeSet2 < 8 Then VolumeSet2 = 8

'SORT HERE
'### CODE FINDS
'Sub Percent_EQI_1_100(EQI_PER_T, EQISOL, EQISOL2)
'Sub Percent_Moon_1_100(MOON_PER_T)
'Sub Percent_SUN_1_100(SUN_PER_T, DarkLight)
Dim Vol_FlagSet, DEMAND_MP3_OFF

Dim OldDarkSpeakin, OldNow2
Dim SunSetFor0Percent
Dim DayOrMidNightVol

Dim StartedVoiceLatch
Dim SimuX
Dim WinampISPlayingAndSetToStopFlagAfterVoiceString
Dim iiNow As Double
Dim xLATCHTIMEROverRide

Dim SayTimeVolAdjustLateNite, HasWinAmpBeenPlayingRecent
Dim WinampISPlayingAndSetToStopFlagAfterVoice
Dim LATCHTIMER As Double, HALO5, SPEAKOUT
Dim LATCHTIMEROverRide

Dim DemandSun, TTA3, XABG, EQI_XABG
Dim EQI_PER_i, EQI_PER_i2, EQI_PER_OT, EQI_PER_x, EQI_PER_x1, EQI_PER_MULTI, EQI_LIGHT, EQI_Run1ST, EQI_Run2nd
Dim Sun_PER_i, Sun_PER_i2, Sun_PER_OT, Sun_PER_x, Sun_PER_x1, Sun_PER_MULTI, Sun_LIGHT, Run2ndSun
Dim PER_i, PER_i2, MOON_PER_OT, PER_x, PER_x1, PER_MULTI, MOON_LIGHT, Run1STMoon
Public TSSP$, Latch
Public F1, F2, F3, F4, DexFORMAT1, DexFORMAT2, DexFORMAT3, ToolH, tMm
Dim LastSleep, AwakeSince, LastAwake, LastAwake2, QuickNap

Public IDLESCHEDULE, MdateMOD


Const MinVol2Set = 8


'Sig Line Stuff
'Timer_CountD1Text_timer()
'Lab10Sub()
Public PillAmountCurrentlyTaking, MoonH, Tool4, SunResult3, SunResult4, Tool4D

Dim STANDVAR, STANDVAR2, XVOL1, XVOL2, TooL46$, DOL1, DOL2

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type


Dim XI, XPDR, XXXC, XTT, XSW, XAOT, XCPN, XOT, XSTR, XPDRF
Dim WinAmpStoppedPlayTime
Dim SysOldUpTime, Tool77$

Dim Dix, Bullet, Change2
Dim CountD1, OLTt2, TimeWait2, OldKey2, OldVolSet

Dim ThisSysUptime2$, PillsYes, PillsNo, MsgResultStore, WinAmpPowerHwndStore

Dim INPUTDATE, TestDate, Year5, MonthStart, DayCount1, DayCountMonth, WholeYear1, DayCountT
Dim Month5, WeeksSinceYear, WeeksSinceInput As Long, ResultYearDate, Month7, WeeksPlusDays

Dim DaysHossTotalVar
Dim DaysHossInpaitentVar

Dim OldSecond
Dim TickGo, VCodeTTB1$, VCodeTTB2$, TimeSetIce, TimeSetVid, XSlider1

'Vcode Keylogger
'Timer14_Timer()
'the other one slow one
'Timer4 .Enabled = True

Public WinAmpStore, OldLabel23, KeyIdleTime2, NotKnot, KeyIdleTimeVideo
Public XmarkSpot, WinVol5, OlVcodeTT$, OldActiveWindow$
Public WinampHwndVM1, WinampHwndVM2, VideoHwnd
Public TestRun, Test5, Test7
'tt8 = FileInUse("D:\Temp\KeyHit-Tidal.txt")

'TestKeyLoggOff = True

Public MSGRESULT1, MsgResult2

'Timer13_Timer() ------- Look for this one for sendkeys on progs
'Private Sub Vcode_timer() --Look this save Key Loggs

Dim DisallowVideo

Public Ff1
'Key Searches--------
'--------------
'AtidalZ
'LockSaverDelay1
'Call WinonPoint
'Call Timer_CountD1Text_timer
'Call KeyOrMouse
'LastPressAY

'Call PassModule.Main
'Call SetLockMax
'Call KeyWindowCheck
'Sub Timer7WinAmpVolIDE_Timer
'Sub Timer8WinAmpVol_Timer()
'Sub MouseDetectMove()
'Sub SlowKeyPress()
'Call Equinox3()
'Call Form_Load
'Call Time_Code
'Call DateMidnight_CountDown
'Call SunRise_Bong
'Call SunRise_SunSet
'Call Tidal_Code
'Call Moon_Code
'----------------------------

Dim WinArray() As Long
Public WINAMP2, OWinAmp2, WinCap

'A = TimeSerial(6, 0, 0) ' Winamp Stop Pause 6 AM

Private Declare Function GetWindowLong _
    Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, _
    ByVal nIndex As Long) _
    As Long

Private Const GWL_STYLE = (-16)
Private Const WS_VISIBLE = &H10000000

Public NotePad2RT2
Public VolumeSet1, VolumeSet2, OldMsgResult, TimeDelay, OldVolumeSet
Public IdleOnNday1$, IdleOnNday2$
Public Timer11Var, AppTitle, EasyNow

Public QQX1 As Long, QQD1 As Long, QQN1 As Long
Public QQX2 As Long, QQD2 As Long, QQN2 As Long
Public QQ91 As Date, QQT1 As Date, QQZ1 As Date
Public QQ92 As Date, QQT2 As Date, QQZ2 As Date
Public OldDayNow
Public IdleMin1 As Date, IdleMin2 As Date, TTQ, SetFlagZero5, SetFlagZeroDay
Public WExpand, MMove2, MMove2Time, Mouse10CLicks, MouseClicks, MainIdleTime, IsOldIdleTime
Public IdleOff As Date, IdleOn As Date, OldMark, OldMark2
Public IdleOffT As Date, IdleOnT As Date
Public IdleOnN1$, IdleOnN2$, Fr0$, IdleOnN3$, IdleOnN4$
Public IdleOnN5$, IdleOnN6$
Public ThisSysUptime3 As Date
Public IdleT As Date, IdleS As Date, IdleMin As Date, IdleT2 As Date, IdleT2Flag
', , OldIdleMin As Date, IdleMinFlag

Public SliderC

Public LastSysUptime$
Public ThisSysUptime$

Public MMove
Public DownLoadFinished2, DownLoadFinished3, DownLoadFinished4



'Const PillsMinus = -5
Const PillsMinus = -15 '--13-4-2008


'***********************************************************\
'NOTES: Processor Usage Timer5

'YOU MUST HAVE WMI SDK INSTALLED.  YOU CAN GET IT AT
'http://msdn.microsoft.com/downloads/sdks/wmi/default.asp
'***********************************************************************

'Private asCpuPaths() As String
Private m_objCPUSet As SWbemObjectSet
Private m_objWMINameSpace As SWbemServices
'***********************************************************************

Public TTx$

Public Wd$

Public OutlookE, NeroE, NeroEX

Public KBPress3

Public Lab2$

Public KeyMP3Nx, KeyMP3Nv1, KeyMP3Nv2

Public YYt, YYx, Tim

Public FirstTime, ClearNextStageMp3

Public FiveMinTime As Date
Public DelayOn As Date
Public DelayOn2 As Date

Const MYDocU$ = "\# 01 My Documents\"

Public TtYy 'Day change for keylog day in file
Public TTTime
Public Yeast
Public TTVol
Public SigPath$

Const IDERun = True
'Const IDERun = False

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
'---In Module---'
Private Type TimeConv
    Days        As Long
    Hours       As Long
    Minutes     As Long
    Seconds     As Long
    MSeconds    As Long
End Type

Private Declare Function sndPlaySound Lib "Winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long  'MODULE 1126

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long


Public GetForegroundWindowOld

Public Lop

Private cProcesses As New clsCnProc

Dim test_hwnd As Long

'Note - Applications that require handling of SAPI events should declair the
'SpVoice as follows:
'Dim WithEvents Voice As SpVoice

Public Termin As Integer
Public A As Double
Public B As Double
Public d As Double
Public pi As Double
Public Jh As Double
Public Hg As Double
Public Ew As Double
Public Wq2 As Double
Public Zzx As Double
Public Pulp As Double
Public Bv$
Public Are As Integer
Public Are2 As Integer
Public X22
Public Y22

Public Xxx2
Public Yyy2

Public Ar2$
Public Changeday As Double
Public Lpsauce As Double
Public Pig As Double
Public Tig As Double
Public W$
Public extra As Integer
Public Splash2 As Double
Public scrwidth_old As Double
Public Colapse As Date

Public WTrue As Double
Public HWTrue As Double
Public KWTrue As Double
Public TW1 As Double
Public TW2 As Double
Public TW3 As Double

Public H321$
Public Est As Integer
Public Est2 As Integer
Public Est3 As Integer

Public Speak$
Public SegM As Date

Public DeaDab$

Public Time10

Public Tool1 As Date
Public Tool2 As Date
Public Tool3 As Date
Public Yellow As Date
Public Empire As Integer
Public Empire2 As Integer

Public BeSet1$
Public BeSet2$

Public Xcag
Public Segm9 As Date

Dim XD(51) As Double
Public Climb As Integer
Public Upset As Integer
Public Apple As Integer
Public Ttimechop As Date
Public Flogship As Integer
Public Flogship2 As Date
Public Aggy As Date
Public Bvs$
Public XxYy As Date
Public Wtime As Date
Public x1 As Date
Public x2 As Date
Public X3 As Date
Public X4 As Date
Public x5 As Date
Public X6 As Date
Public X7 As Date
Public X8 As Date
Public X9 As Date

Public X10 As Date
Public X11 As Date
Public X33 As Date
Public X55 As Date
Public X77 As Date

Public Tog1 As Integer
Public Tog2 As Date
Public Tog4 As Integer
Public Tog6 As Integer
Public Tog7 As Date
Public MoonLabel$
Public Asd As Integer

Public Wqe1$
Public Wqe2$
Public Wqe3$
Public Wqe4$

Public Dag As Date
Public UpDate2 As Integer
Public Previous1 As Integer

Public TooLNewSunRise As Date
Public TooLNewSunSet As Date
Public TooLNewNoon As Date

Public TooLLastSunRise As Date
Public TooLLastSunSet As Date
Public TooLLastNoon As Date

Public TooLTodaySunRise As Date
Public TooLDayBeforeSunSet As Date
Public TooLTodaySunSet As Date
Public TooLTodayNoon As Date

Public Dag5 As Integer

Public Wtool As Integer
Public Wtool2 As Long

Public R2D2 As Integer

Public Tooltipx As Single
Public Tooltipy As Single
Public Nx!, Ny!
Public Tx!, Ty!

Public Wtoolcol1 As Long
Public Wtoolcol2 As Long
Public Wtoolcol8 As Long
Public Wtoolcol11 As Long
Public Wtoolcol12 As Long
Public Wtoolcol13 As Long
Public Wtoolcol14 As Long
Public Wtoolcol20 As Long
Public Wtoolcol22 As Long
Public Wtoolcol23 As Long
Public Wtoolcol26 As Long
Public Trgb As Long
Public Trgb2 As Long
Public Txf As Integer


Public Easyride As Date

Public Xmp5 As Long
Public Ymp5 As Long
Public Xmp6 As Long
Public Ymp6 As Long
Public Startuptime As Date
Public Hidecursor2 As Integer
Public Startmouse As Integer
Public Wq5 As Long
Public Now4 As Date
Public Rty2 As Integer
Public Rty3 As Integer


Public Easyzap As Integer
Public Quake As Integer

Public Secstep As Long
Public pCount As Long
Public Pcount2 As Long
Public Pcount3 As Long
Public Hitinit As Integer
Public NoF12 As Integer

Public Curprocid As Long
Public Curprocid2 As Long
Public Tagbok As Integer        ' to do with sun rise/set label colors

'Public monthnow As Date         ' speak the month when it arrives
'Public monthnow2 As Date
Public Xdf2 As Integer
Public Xdf3 As Integer


Public Mousemove As Long

Public Mousetimersend1 As Date
Public Mousetimersend2 As Integer
Public MouseFirstRun As Integer

Public BdFc As Long

Public FirstRun2 As Integer

Public HwndCTLTask2 As Long
Public HwndCTLTask3 As Long
Public HwndCTLTask4 As Long
Public OldHwndCTLTask3 As Long
Public LockXY As Integer
Public LockXYForTime As Date
Public TopToolBar1 As Long
Public TopToolBar2 As Long

Public Pleet1 As Long
Public OldPleet2 As Long

Public LockMove As Double
Public LockMoveLatch As Long
Public lock44 As Long
Public RunTime2 As Date
Public DelayXYUpdates22 As Date
Public TidalMoveDelay As Long
'Public Delay1 As Date
Public Cosh As Long

Public LockNorts As Date

Public NortonAdate1 As Date
Public NortonAdate2 As Date

Public FS        ' Set = CreateObject("Scripting.FileSystemObject")

Public TidalTopPos As Integer
Public OldTopTidalPos As Integer

Public ScreenWidth As Long
Public ScreenHeight As Long

Public LockXY2 As Long

Public Que2 As Integer

Public PreRect3

Public TestTrig

Public QuickPage
Public TimeBack
Public RT2

Public Rtt

Public Eta2 As Date

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

'-----------------------------------------------
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
  hwndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type
'-----------------------------------------------


'---------------------------------------
'---------------------------------------
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
'---------------------------------------
'---------------------------------------
Private Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
'---------------------------------------
Private Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
'---------------------------------------
'---------------------------------------



Private Function Menu_Height()
 
    MenuInfo.cbSize = Len(MenuInfo)
    
    If GetMenuBarInfo(Me.hWnd, OBJID_MENU, 0, MenuInfo) Then
   
        With MenuInfo.rcBar
       
'            Debug.Print "Left: " & CStr(.Left)
'            Debug.Print "Right: " & CStr(.Right)
'            Debug.Print "Top: " & CStr(.Top)
'            Debug.Print "Bottom: " & CStr(.Bottom)
            'Menu_Height = CStr(.Top) + CStr(.Bottom)
            Menu_Height = CStr(.Bottom) '- CStr(.Top)
        End With
       
    End If
   
End Function
'------------------------------------------


Function PlayWav(ByVal u As String)  'MODULE 1114
soundfile1 = u
sndPlaySound u, 1
End Function

Sub SetLockMax()

If LockSaverDelay = 0 Then Exit Sub
ATidalX.ProgressBarLock.Max = DateDiff("s", 0, LockSaverDelay)


'Call ProgressLock
On Local Error Resume Next
ProgressBarLock.Value = DateDiff("s", Now, LockSSaver)
Label9.Caption = Mid$(Format$(TimeSerial(0, 0, ProgressBarLock.Value), "HH:MM:SS"), 4)

If App.title <> "Tidal Information..." Then
    frmPassLock.ProgressBarLock.Max = DateDiff("s", 0, LockSaverDelay)
    frmPassLock.ProgressBarLock.Value = ProgressBarLock.Value
    frmPassLock.Label9 = Label9
End If


'If frmPassLock.hide Then Exit Sub


End Sub



Private Sub Command3_Click()

End Sub

Private Sub FileSaveFileExit_timer()
        Exit Sub
        FileSaveFileExit.Enabled = False
        Call SendKey(70, 4) '#
        Call SendKey(83, 0) '#
        Call SendKey(70, 4) '#
        Call SendKey(88, 0) '#

End Sub




Private Sub Form_Load()

If Command$ = "" Then If App.PrevInstance = True Then End

Call Project_Check_Date.VB_PROJECT_CHECKDATE("ATIDALX")

OLD_RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD = -2

O_BONG_DAY_NOW = Day(Now)

MHDR_NOW_DATE = -1

FIRST_RUN_SUN_PERCENT_TALKER_LATCH_UP = True
DARK_OR_LIGHT_OLD = -2
'AWINAMP_CONTROL.Show
FLAG_TO_MOON_TALK = False
O_ABH = 0
DarkLightOLD = -2
'Exit Sub
FIRST_RUN_SUN_PERCENT_TALKER = False

Set FS = CreateObject("Scripting.FileSystemObject")

'Label10.ToolTipText = "Clipboard A Set of Info About Moon"
'Hardcoded

Call MNU_ADMINISTRATOR_Click

OLD_RS232_LOGGER_FILE_TIME_SECOND = -1
OLD_RS232_LOGGER_FILE_TIME_MINUTE = -1
OLD_RS232_LOGGER_FILE_TIME_HOUR = -1
START_DOOR_TALKER = True


Label26.ToolTipText = "MOON TALK %"
Label12.ToolTipText = "MOON 8 PHASE NAME TALK"
Label22.ToolTipText = "TIDAL KEY AND MOUSE LOG -- NOTEPAD"
Label13.ToolTipText = "Equinox Tidal Info Form"
Label25.ToolTipText = "TALK"
Label24.ToolTipText = "TALK NOW MOON PHASE NAME TALK"
Label11.ToolTipText = "TALK NEXT MOON PHASE NAME TALK"
Label7.ToolTipText = "TALK DATE"

Label2.ToolTipText = "TALK SUN RISE"
Label3.ToolTipText = "TALK SUN SET"
Label41.ToolTipText = "TALK SOLAR NOON"
LABEL_TIME.ToolTipText = "TALK TIME DATE VOLUME UP 2"
Label37.ToolTipText = "SOUND MALICIOUS RUTLAND"
Label40.ToolTipText = "TALK TIME AND Dog Bark Commando"

Label4.ToolTipText = "ENTER EXTREMER LOCK"


VOICELISTTimer.Enabled = False
VolGetTimer.Enabled = False
TimerDoubleVolume.Enabled = False
OVOLTAG = -1
OVOLTAG2 = -1
'-------------
'Mnu_NoJumpUpVol.Checked = False

ReDim OXAZY(20)
ReDim XAZY(20)
ReDim XZAP(20)

SimuX = -1

'StringVoiceToPLayActive = True

MEHwnD = Me.hWnd

'Run1STMoon = -2


App.title = "Tidal Information..."

If Dir(App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\", vbDirectory) = "" Then
    On Error Resume Next
        MkDir App.Path + "\00_Text_Data\"
        MkDir App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\"
    On Error GoTo 0
End If


If Dir(App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\ISIDE.txt") <> "" Then
    On Error Resume Next
        FR1 = FreeFile
        Fd$ = "o"
        Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\ISIDE.txt" For Input As #FR1
            Line Input #FR1, Fd$
        Close #FR1
    On Error GoTo 0
End If

If Fd$ = "I" Then Mnu_VBISIDE.Checked = True

If IsIDE = False And Mnu_VBISIDE.Checked Then
'    Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
'    Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" /r """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
'    End
End If

'or /runexit

'Call SolarTimeList

If 1 = 2 Then

    Set FS = CreateObject("Scripting.FileSystemObject")
    Dim d
    On Local Error Resume Next
    Set dc = FS.Drives
    For Each d In dc
        DR = d.DriveLetter
            If d.IsReady = True Then
                n = d.VolumeName
                If InStr(n, "RAM") > 0 Then
                    DRX = True
                    Exit For
                End If
            End If
    Next
    On Local Error GoTo 0
    RamDrive = DR
    If DRX = False Then RamDrive = Mid(App.Path, 1, 1)

End If


Me.BackColor = 0

KeyIdleTime = Now + TimeSerial(0, 0, 30)
KeyIdleTime2 = Now + TimeSerial(0, 1, 0)
KeyIdleTimeVideo = Now + TimeSerial(0, 30, 0)

ReDim WinArray(0)

'SHOULD USE MESSY TIMERS HERE
'MESSAGE TIMERS -- OH YEAH

Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer5.Enabled = False
Timer7.Enabled = False

Load ScanPath

'For notepad font bigger timer
RKeyPlus = 99

If Command$ = "" Then If App.PrevInstance = True Then End

If Dir(App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\KeyLoggDate.txt") <> "" Then
    On Error Resume Next
        FR = FreeFile
        Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\KeyLoggDate.txt" For Input As #FR
            Line Input #FR, A1$
        Close #FR
        If Err.Number = 0 Then
        On Error GoTo 0
        
        KeyLoggDate = DateValue(A1$) + TimeValue(A1$)
        If Int(Now) <> Int(KeyLoggDate) Then
            KeyLoggDate = Now
            FR = FreeFile
            Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\KeyLoggDate.txt" For Output As #FR
            Print #FR, KeyLoggDate
            Close #FR
        End If
    End If
End If

If Dir(App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\KeyLoggDate.txt") = "" Then
    KeyLoggDate = Now
    FR = FreeFile
    Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\KeyLoggDate.txt" For Output As #FR
    Print #FR, KeyLoggDate
    Close #FR
End If

    
    
    
DeadLock = Now + TimeSerial(0, 30, 0)

On Local Error Resume Next
Set m_objWMINameSpace = GetObject("winmgmts:")
On Local Error GoTo 0

If HOUR(Now) >= 19 Or HOUR(Now) < 8 Then cute = 20 Else cute = 10

LockSaverDelay1 = TimeSerial(0, cute, 0)
LockSaverDelay2 = TimeSerial(0, cute, 0)
LockSaverDelay3 = TimeSerial(0, 0, 10)
If IsIDE Then LockSaverDelay4 = TimeSerial(0, 0, 5) Else LockSaverDelay4 = TimeSerial(0, 1, 0)
LockSaverDelay = LockSaverDelay1

LockSSaver = Now + LockSaverDelay1

ProgressBarLock.Min = 0

Call SetLockMax

Call ProgressLock

ProgressBarCPU.Min = 0
ProgressBarCPU.Max = 100

'RamDrive

FirstTime = True


If 1 = 2 Then
    On Error Resume Next
        MkDir RamDrive + ":\Temp"
    On Error GoTo 0
    Open RamDrive + ":\temp\KeyHit-Tidal.txt" For Output As #1
        Print #1, "0"
    Close #1
    'Open "Q:\temp\KeyHit-Tidal2.txt" For Output As #1
    'Print #1, "0"
    'Close #1
End If

WExpand = -1



'------------------------------
'VolumeSet1

MsgResult = WinAmpChkIsPlaying

OldMsgResult = MsgResult

Call LoadLoggs

'WinampHwnd = FindWindow("Winamp v1.x", vbNullString)
'For R = 1 To 100
'MsgResult = SendMessage(WinampHwnd, WM_COMMAND, Raisevolumeby1percent, ByVal XcNul)
'Next


Eta2 = 0

'*************************************************************************************************
'*************************************************************************************************
'*************************************************************************************************
'*************************************************************************************************

'Call frmOperatingSystem.ShowForm
Load Menu

'Load MDIProcServ


Call FirstRun
  
'*************************************************************************************************
'*************************************************************************************************
'*************************************************************************************************
  

'Load frmIECtrl
'frmIECtrl.Show


'Timer1.Interval = 500
'Timer2.Interval = 500
'Timer3.Interval = 50
'Timer5.Interval = 500
'Timer6.Interval = 500

'Menu.Show

Menu.Timer1.Interval = 500
Menu.Timer2.Interval = 500

'If IsIDE() = False Then
'ATidalX.Timer7WinAmpVolIDE.Enabled = False
'End If
'If IsIDE() = True Then
'ATidalX.Timer8WinAmpVol.Enabled = False
'End If



Label2.BackColor = QBColor(15)
Label2.ForeColor = QBColor(0)
Label3.BackColor = QBColor(0)
Label3.ForeColor = QBColor(15)

If App.title <> "Tidal Information..." Then
    frmPassLock.Label2.BackColor = QBColor(15)
    frmPassLock.Label2.ForeColor = QBColor(0)
    frmPassLock.Label3.BackColor = QBColor(0)
    frmPassLock.Label3.ForeColor = QBColor(15)
End If

'Call AntiVirusMatt


'Load Volume2
'Volume2.Show

'Exit Sub
'Timer1.Enabled = False
'Timer2.Enabled = False
'Timer3.Enabled = False
'Timer4.Enabled = False
'Timer5.Enabled = False
'Timer6.Enabled = False

'Menu.Timer1.Enabled = False
'Menu.Timer2.Enabled = False

'D:\VB6\VB-NT\00_Best_VB_01\Wmi_CPU\WMICPU2 MINI.exe



'Shell "D:\VB6\VB-NT\00_Best_VB_01\Run Fav Progs\Run Fav Programs.exe", vbHide


If 1 = 2 And IsIDE = False Then
    Rf = FindWindow(vbNullString, "#00_Schedule_Timer")
    If Rf = 0 Then
        'Shell "D:\VB6\VB-NT\00_Best_VB_01\#00_Schedule_Timer\#00_Schedule_Timer-ACER.exe", vbNormalNoFocus
    End If
    'i = FindWindow(vbNullString, "Windows Task Manager")
    'A = DisableClose(i)
End If


'CID Run Me.

'Load FrmJoy
'If JOYPAD_SOCKETED_IN = True Then

FILE_NAME_DAY_INFO_TXT = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\DAY_INFO.TXT"
Dim DAY_GET
FR1 = FreeFile
If Dir(FILE_NAME_DAY_INFO_TXT) <> "" Then
    Open FILE_NAME_DAY_INFO_TXT For Input As #FR1
        Line Input #FR1, DAY_GET
    Close #FR1
End If
If DAY_GET <> Format(Now, "DDDD") Then 'Or 1 = 1 Then
    FR1 = FreeFile
    Open FILE_NAME_DAY_INFO_TXT For Output As #FR1
        Print #FR1, Format(Now, "DDDD")
    Close #FR1
    
    ProcessID_2 = Shell("regsvr32 /S " + App.Path + "\dx8vb.dll")
    Beep
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, ProcessID_2)
    TIMER_10 = Now + TimeSerial(0, 0, 10)
    Do
        Call GetExitCodeProcess(hProcess, ExitCode)
'        DoEvents
        'STILL RUNNER GIVE MORE TIME
        
        'SOMETIME TRIGGER HAPPEN ON NEW DATE PUT FORWARD CLOCK
        'ONLY IN EXE NOT IDE
'------------------------------------------------------------------
'        If ExitCode > 0 Then
'            VBLINE = "Call GetExitCodeProcess(hProcess, ExitCode)"
'            MsgBox "STOP IS HERE " + vbCrLf + VBLINE
'            Stop
'        End If
        If ExitCode = 259 Then
            TIMER_10 = Now + TimeSerial(0, 0, 10)
        End If
        If ExitCode < 259 Then Exit Do
        If TIMER_10 < Now Then
            VBLINE = "Call GetExitCodeProcess(hProcess, ExitCode)"
            MsgBox "TIME OUT OPTION OF EXIT __ HERE " + vbCrLf + VBLINE
            Exit Do
        End If
    Loop While 1 = 1
    
    CloseHandle hProcess
End If
'reg query HKLM\SOFTWARE\Classes /s /f dx8vb.dll
'MsgBox "HELP"


'MsgBox "HELP2"

Load FrmJoy
Load frmDXGAMEJOY

'MsgBox "HELP3"

'MsgBox "HELP3"

'End If

'ATidalZ.Show
'Drives.Show
'Counter1.Show
'Counter2.Show
'Counter3.Show
'Counter4.Show
 
 
'Call WorkVisits

Call OpenMixer(0)

tf = SetVolume(SYNTHESIZER, 50) '  --- MIDI GOES LOWER COZ ALWAYS SO LOUD
tf = SetVolume(COMPACTDISC, 100)
tf = SetVolume(WAVEOUT, 100)
tf = SetVolume(LINEIN, 50)
'tf = SetVolume(SPEAKER, ATidalX.LblVol)
'tf = SetVolume(WAVEIN, 100)

On Local Error GoTo 0
If Dir$(App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Volume.txt") <> "" Then
    Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Volume.txt" For Input As #1
    Line Input #1, ll1$
    Line Input #1, ll2$
    Close #1
End If

ATidalX.LblVol = str(GetVolume(SPEAKER))
VolumeSet1 = Val(ll1$)
VolumeSet2 = Val(ll2$)

'Debug.Print "VOLUME 1 ON LOAD = " + str(VolumeSet1)
'Debug.Print "VOLUME 2 ON LOAD = " + str(VolumeSet2)
'VOL 1 =  MUSIC
ATidalX.ProgressBarVol1.Value = VolumeSet1 'ATidalX.LblVol
If VolumeSet2 > 100 Then VolumeSet2 = 100
ATidalX.ProgressBarVol2.Value = VolumeSet2 'ATidalX.LblVol
'Call TimerDoubleVolume_Timer



If WinAmpChkIsPlaying = 1 Then
    ATidalX.LblVol = str(VolumeSet1)
Else
    ATidalX.LblVol = str(VolumeSet2)
End If
Call ATidalX.LblVol_Change

'VolumeSet2 = MinVol2Set
'tf = SetVolume(SPEAKER, ATidalX.LblVol)



VolGetTimer.Enabled = True
TimerDoubleVolume.Enabled = True


'FORM TO SCREEN POSITIONING
Call Mnu_Expand_W_Click




Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer5.Enabled = True
Timer7.Enabled = True
VOICELISTTimer.Enabled = True
'OVOLTAG = 0

NOVOLSAYASSTART = True
ATidalX.Show

'MINIMIZE VB IDE AT RUN
Timer17.Enabled = True


ATidalX.WindowState = vbNormal

DARK_OR_MOD_COUNTER = 0
MOON_MOD_COUNTER = -1
MOON_MOD_COUNTER = 0

'Call MNU_VOX_LIST_Click

'------------------------------------------------------------------


'KEYBOARD HOOK LAST OF ALL
'-------------------------

TestKeyLoggOff = False

'Use this one testing Auto Oppo When In IDE No Keylogg

'------------------------------------------
'If IsIDE = True Then TestKeyLoggOff = True
'PROBLEM
'GOT NEW CODE TO DETECT IDE NOW
'-------------------------------------------


'Use Force Set DO Use Keyogg
'TestKeyLoggOff = False

'we got a menu command to activate keylogging in VB it sets command to ignor to see if vb running
'then that will not block keys on vb
'**** here ****** >> Mnu_KeyHook_Click

Dim RunningKeyLogg

If 1 = 2 Then
    If TestKeyLoggOff = False Then
        Load DSkeybd_F
        RunningKeyLogg = True
    End If
    
    If IsIDE = False And Mnu_VBISIDE.Checked Then
        Load DSkeybd_F
        RunningKeyLogg = True
    End If

End If

'Don't get caught in the IDE Design Code
If IsVBrunningForCode = True Then
    
    'GO CODE WRONG THIS NEED RUN ALL THE TIME
    'NOT JUST ONLY AT FIRST KEY PRESS
    'WILL RUN ALWAYS WHEN A KEY PRESS
    
    ATidalX.Timer_REINTRODUCE_DSkeybd_m.Enabled = True
    
    'WRONG DON'T WANT CALLBACK VAR = 2 BECAUSE THAT BLOCK KEY
    'Callback = 2:
    'TRY AND UNDERSTAND WHY WAS OKAY IN IS-IDE
    
    'Exit Function
    
    Else
    
    
    Call Mnu_KeyHook_Click

End If


'Load DSkeybd_F
'RunningKeyLogg = True

'If RunningKeyLogg = true Then
'    ASYNC_Key_Press_Timer.Enabled = False
'Else
'    ASYNC_Key_Press_Timer.Enabled = True
'End If

'Call HookVB&(True)

'------------------------------------------------------------------
'COMMENTS BUNCH
'------------------------------------------------------------------


MNU_JOYPAD_SOFTWARE_STATUS.Visible = False

MNU_LATENIGHTOVERRIDE.Checked = True

MNU_SPEECH_OFF.Checked = True
MNU_SPEECH_OFF.Checked = False
Call SPEECH_OFF_SETTER
Call MNU_VERSION_Click

Form_Load_Done_For_Timer14 = True
End Sub

Sub WorkVisits()


'WORK TO DO CALC TIME SINCE LAST SEEN VISITORS ADD SAME FILE THE INFO
Err.Clear
On Error Resume Next
Open GetSpecialfolder(CSIDL_PERSONAL) + "\CallerID\#Since SW Visits.txt" For Input As #1
Open GetSpecialfolder(CSIDL_PERSONAL) + "\CallerID\#Since SW Visits!.txt" For Output As #2
'SORT HERE

'JOHN
'FRANK
'LIZ
'KYLIE
'SAM
'SARAN
'DALE
'CLAIR

'VISIT PERSON
'Line Input #1, STANDVAR: Call STANDVARSVISITS

'VISIT DOOR

'VISIT VOICE PHONE

'VISIT TEXT PHONE

Do

Line Input #1, STANDVAR


    If Mid$(STANDVAR, 3, 1) = "-" Then
        Call STANDVARSVISITS
    
    Else
        
        Print #2, STANDVAR

    End If
    

Loop Until EOF(1)

Close #1, #2

If Err.Number = 0 Then
    Kill GetSpecialfolder(CSIDL_PERSONAL) + "\CallerID\#Since SW Visits.txt"
    Name GetSpecialfolder(CSIDL_PERSONAL) + "\CallerID\#Since SW Visits!.txt" As GetSpecialfolder(CSIDL_PERSONAL) + "CallerID\#Since SW Visits.txt"
End If

End Sub

Sub STANDVARSVISITS()

If Mid$(STANDVAR, 1, 2) = "00" Then
    Print #2, STANDVAR
    Exit Sub
End If

INPUTDATE = DateValue(Mid$(STANDVAR, 1, 10))
TestDate = Now
Call FindTimeInfo

TT = " -- " + ResultYearDate + " Years" + str(Month7) + "M" + str(DayCountMonth) + "D" + str(WeeksSinceInput) + " Weeks &" + str(WeeksPlusDays) + " Days"


RR = InStr(STANDVAR, "##-")
If RR = 0 Then TT = " ##-" + TT
If RR = 0 Then RR = Len(STANDVAR)
STANDVAR = Trim(Mid$(STANDVAR, 1, RR + 2)) + TT
Print #2, STANDVAR

End Sub

Sub LoadLoggs()
'Exit Sub





'Counter1.Show
'Counter2.Show
'Counter3.Show
'Counter4.Show

If 1 = 2 Then
ScanPath.cboMask.Text = "*.txt"
ScanPath.txtPath.Text = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\"
Call ScanPath.cmdScan_Click

For we = 1 To ScanPath.ListView1.ListItems.Count
    A1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(we)
    'fs.copyFile a1$, a2$
    Set F = FS.GetFile(A1$ + B1$)
    DD = F.Size
    If DD > 1000! ^ 2 Then
    G$ = A1$ + "archive TEXT\" + Mid$(B1$, 1, Len(B1$) - 3) + Format$(Now, "YYYY-MM-DD HH-MM-SS") + ".txt"
    FS.MoveFile A1$ + B1$, G$
    
    End If
Next
End If



PATH1_VAR = App.Path + "\00_Text_Data\"
PATH2_VAR = GetComputerName + "-" + GetUserName + "\"
If Dir(PATH1_VAR + PATH2_VAR, vbDirectory) = "" Then
    
    On Error Resume Next
    MkDir PATH1_VAR
    MkDir PATH1_VAR + PATH2_VAR
    If Err.Number > 0 Then
        MsgBox "ERROR MAKE THIS FOLDER" + vbCrLf + PATH1_VAR + PATH2_VAR + vbCrLf + Err.DESCRIPTON
        Reset
        Unload Me
    End If
End If
On Error GoTo 0


FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Key Logger Text.txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4

Dim FR1
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
Print #FR1,
Print #FR1, "----- Start -----"
Print #FR1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS")
Close #FR1

FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Key Logger Text-Stripe.txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4

FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
Print #FR1,
Print #FR1, "----- Start -----"
Print #FR1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS")
Close #FR1

FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Key Logger Text " + Format$(KeyLoggDate, "YYYY-mm-dd HH-MM-SS") + ".txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4

FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
Print #FR1,
Print #FR1, "----- Start -----"
Print #FR1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS")
Close #FR1

FILENAME_USE = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\ISIDE.txt"
If Dir$(FILENAME_USE) <> "" Then
    FR1 = FreeFile
    Open FILENAME_USE For Input As #FR1
    Line Input #FR1, f8$
    Close #FR1
End If

Mnu_VBISIDE.Checked = False
If f8$ = "I" Then
    Mnu_VBISIDE.Checked = True
End If

GoTo skip

If 1 = 2 Then

    If IsIDE = False And f8$ = "I" Then
        'Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /make " + App.Path + " tidal.exe Tidal_info2"
        MsgBox "I"
        Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /r " + App.Path + "\Tidal.vbp"
        
        End
    End If
    
    If IsIDE = True And f8$ = "o" And IDERun = False Then
        'Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /make " + App.Path + " tidal.exe Tidal_info2"
        'good
        'Shell "VB6.EXE /make " + App.Path + "\Tidal.vbp " + App.Path + "\tidal.exe"
        'Shell App.Path + "\tidal_Launch.exe"
        'MsgBox "i"
        'Shell App.Path + "\tidal.exe"
        'End
    End If
    
    
    End If

skip:



On Error Resume Next
If TtYy <> Day(Now) Then
    If Dir(GetSpecialfolder(CSIDL_PERSONAL) + "\CallerID\01 KeyStroke_Logger_" + GetComputerName + "-" + GetUserName + ".txt") = "" Then
        MkDir GetSpecialfolder(CSIDL_PERSONAL) + "\CallerID"
    End If
    
    FILENAME_IN_USE_CHECK = GetSpecialfolder(CSIDL_PERSONAL) + "\CallerID\01 KeyStroke_Logger_" + GetComputerName + "-" + GetUserName + ".txt"
    FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4

    Fr5 = FreeFile
    Open FILENAME_IN_USE_CHECK For Append As #Fr5
        ttyu$ = Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS a/p")
        Print #Fr5, "Program Start.."
        Print #Fr5, ttyu$; " -------"
        TtYy = Day(Now)
    Close #Fr5
End If
On Error GoTo 0




End Sub

Private Sub Form_Unload(Cancel As Integer)

'If frmPassLock.show Then Cancel = True: Exit Sub

'----------------------
'PROBLEM
If App.title <> "Tidal Information..." And IsIDE = False Then Cancel = True: Exit Sub
'----------------------


Close


'qq8 = GetIdleM(3)
Close #Fr5

ATidalX.MMControl1.Command = "Close"
ATidalX.MMControl2.Command = "Close"
ATidalX.MMControl3.Command = "Close"
ATidalX.MMControl4.Command = "Close"
ATidalX.MMControl5.Command = "Close"
ATidalX.MMControl6.Command = "Close"
ATidalX.MMControl7.Command = "Close"
ATidalX.MMControl8.Command = "Close"
ATidalX.MMControl9.Command = "Close"

If Dir("C:\TEMP\Pause All For WinAMP-PLAY.txt") <> "" Then
    Kill "C:\TEMP\Pause All For WinAMP-PLAY.txt"
End If
If Dir("C:\TEMP\Pause All For WinAMP-STOP.txt") <> "" Then
    Kill "C:\TEMP\Pause All For WinAMP-STOP.txt"
End If

On Local Error Resume Next
FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Tidal Key and Mouse Log.txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4

gug = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #gug
    Print #gug, "Start Time "; Format$(Startuptime, "DDD DD/MM/YYYY HH:MM:SS AM/PM")
    Print #gug, "End   Time "; Format$(Now, "DDD DD/MM/YYYY HH:MM:SS AM/PM")
    Print #gug, "Mouse    ="; Mousey
    Print #gug, "Keyboard ="; Keyy
Close #gug

FR4 = FreeFile
Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Volume.txt" For Output As #FR4
    Print #FR4, str(VolumeSet1)
    Print #FR4, str(VolumeSet2)
Close #FR4

'Debug.Print "VOLUME 1 ON UNLOAD = " + str(VolumeSet1)
'Debug.Print "VOLUME 2 ON UNLOAD = " + str(VolumeSet2)

freef5 = FreeFile
Open Bv$ + Bvs$ For Output As #freef5
    Print #freef5, str$(X22)
    Print #freef5, str$(Y22)
Close #freef5

'freef5 = FreeFile
'Open App.Path + "\00_Text_Data\System UpTime Logg.txt" For Append As #freef5
'Print #freef5, Format$(Now, "DDD-DD-MMM-YYYY HH:MM:SS ") + Label31.Caption + " --- " + Str(GetTickCount)
'Close #freef5

If DSkeybd_F.IsHook = True Then
    Call DSkeybd_F.UnHookCommand_Click
End If


Close
Reset

'Call TerminateForms

QuitForms = True
aa_MainCode.TrueTerminate = True
aa_MainCode.ALL_FORM_REQUEST_TO_END = True

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


On Error Resume Next
Dim Form As Form
For Each Form In Forms
    If Form.Caption <> Me.Caption Then
        Unload Form
    End If
Next Form
Set Form = Nothing

DoEvents

Unload Me

If FORM_HAS_GOT_TO_UNLOAD = True Then End

FORM_HAS_GOT_TO_UNLOAD = True

End Sub


Function PreviousInstance() As Long
    
    Dim test_hwnd As Long
       
    PreviousInstance = False
    
    test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
    Do While test_hwnd <> 0
       'Check if the window isn't a child
       If GetParent(test_hwnd) = 0 Then
       
       ash$ = GetWindowTitle(test_hwnd)
       
       If ash$ = "Tidal Information..." Then boot = boot + 1
       If ash$ = ATidalX.Caption Then boot = boot + 1
       
       If boot = 2 Then PreviousInstance = True: Exit Do
       End If
       
       'retrieve the next window
       test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    Loop
End Function


Private Sub FirstRun()

'MsgBox "I"

ScreenWidth = (Screen.Width / Screen.TwipsPerPixelX) - 10
ScreenHeight = (Screen.Height / Screen.TwipsPerPixelY) - 10

scrwidth_old = Screen.Width

DelayXYUpdates22 = Now + TimeSerial(0, 0, 5)

MouseFirstRun = 0

RunTime2 = Now + TimeSerial(0, 0, 4)

CompName$ = GetComputerName

drive2$ = ""

If CompName$ = "MAGICRAM2EODUR" Then drive2$ = "H"
If CompName$ = "EDDIESCIENTIST" Then drive2$ = "D"
If CompName$ = "NASA1234567890" Then drive2$ = "C"
If CompName$ = "MEACHELLE12345" Then drive2$ = "C"
If CompName$ = "555ROIDSGOLDRIM" Then drive2$ = "C": drive3$ = "W"
If CompName$ = "MATT-555ROIDS" Then drive2$ = "C": drive3$ = "W"


If Command$ <> "" Then
    'FOR ABOUT INFO AND IF UPDATE HAPPEN
    Call UpDate5(1)

End If

'Call UpDate66

Call moonnext

Call ColourSet

Tagbok = -1

LockNorts = Now - 40

'----------
'C:\Program Files\Common Files\Microsoft Shared\Speech\sapi.dll
'Goto Project>Referances
'CheckBox MicroSoft Speech Object Libary
'Should Be Done
'
'But first you must also have installed microsoft Voice Program
'---------
'And Also dont forget this
'Public Voice As SpVoice
'
'Goes in Declarations
'

'Initialize the voice object
'Call Mnu_Voices_Click
Load TTSAppMain
'TTSAppMain.Visible = False

Startuptime = Now
Trgb = RGB(100, 100, 100)
Trgb2 = RGB(160, 160, 160)
Tog4 = 1
Tog1 = 1
Randomize

'longest day 2005
'speedtime = ((DateSerial(2005, 6, 21) - TimeSerial(0, 0, 10)) - Now)

'shortest day 2005
'speedtime = ((DateSerial(2005, 12, 22) - TimeSerial(0, 0, 10)) - Now)

'Equinox 2005
'speedtime = ((DateSerial(2005, 9, 25) - TimeSerial(0, 0, 10)) - Now)

'wtime = Now + TimeSerial(0, 0, 10)











'hourmate = -1
Flogship2 = Now + TimeSerial(0, 0, 10)

A = DateSerial(2003, 12, 31) + TimeSerial(23, 48, 0)
A = DateSerial(2004, 10, 18) + TimeSerial(14, 34, 0)
Jh = A
pi = (4 * Atn(1)) '+ 0.00000999999
XxYy = Now + TimeSerial(0, 0, 5)
'pig = Now

'hg = 1 + TimeSerial(0, 48, 28)
'ew = 0.9

   

If 1 = 2 Then
    If drive2$ <> "" Then
        On Local Error Resume Next
        ChDir drive3$ + ":" + MYDocU$ + "email bits\email settings\"
        er2 = Err.Number
    
        If er2 > 0 Then
        MkDir drive2$ + ":" + MYDocU$
        MkDir drive3$ + ":" + MYDocU$ + "email bits\email Settings"
        End If
        On Local Error GoTo 0
    End If
    
    
    SigPath$ = ":" + MYDocU$ + "01 Email Settings\Signatures\"
    
    GoTo JUMP10
    
    If 22 = 44 And Dir$(drive3$ + SigPath$ + "signature2.txt") <> "" Then
    
        Dir2.Path = drive2$ + ":\Program Files\Common Files\Symantec Shared\VirusDefs"
    
    
        If Dir$(drive3$ + SigPath$ + "signature6.html") <> "" Then
    
            Open drive3$ + SigPath$ + "Signature1.html" For Input As #1
            Open drive3$ + SigPath$ + "Signature6.html" For Input As #3
            Do
                On Local Error Resume Next
                Line Input #1, one$
                Line Input #3, one1$
                On Local Error GoTo 0
                If one$ <> one1$ Then
                    Kill drive3$ + SigPath$ + "Signature2.txt": Exit Do
                End If
            Loop Until EOF(1)
        
            Close #1, #3
      
        Else
            Kill drive3$ + SigPath$ + "Signature2.txt"
        End If
    End If
    
    If 22 = 44 And Dir$(drive3$ + SigPath$ + "signature2.txt") <> "" Then
        Open drive3$ + SigPath$ + "signature2.txt" For Input As #1
        On Local Error Resume Next
        Do
            Line Input #1, aw$
            If InStr(aw$, "Virus Definitions Date: ") Then
                BeSet1$ = aw$
            End If
            If InStr(aw$, "Definitions Revision : ") Then
                BeSet2$ = aw$
            End If
            On Local Error GoTo 0
        Loop Until EOF(1)
        Close #1
    End If
    
    If 22 = 44 And Dir$(drive3$ + SigPath$ + "signature3.txt") = "" Then
        BeSet1$ = ""
    End If
End If
    

JUMP10:

If 1 = 2 Then

    
    On Local Error GoTo 0
    
    
    Dim NewReboot
    freef5 = FreeFile
    Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\System UpTime Logg StartTime.txt" For Input Lock Write As #freef5
    Line Input #freef5, SysStartTimeOld
    Close #freef5
    
    SysStartTimeOldN = DateValue(SysStartTimeOld) + TimeValue(SysStartTimeOld)
    
    
    On Error Resume Next
    freef5 = FreeFile
    Open "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\00_Text_Data\System UpTime Logg Descrip Old.txt" For Input Lock Write As #freef5
        Line Input #freef5, SysOldUpTime 'Dum
        Line Input #freef5, SysOldUpTime 'Dum
        Line Input #freef5, SysOldUpTime 'Dum
        Line Input #freef5, SysOldUpTime 'Want
    Close #freef5
    On Error GoTo 0
    
    
    TTz1 = Mid$(TTz1, 1, 4) + "-" + Mid$(TTz1, 5, 2) + "-" + Mid$(TTz1, 7, 2) + " " + Mid$(TTz1, 9, 2) + ":" + Mid$(TTz1, 11, 2) + ":" + Mid$(TTz1, 13, 2)
    SysStartTimeNewN = DateValue(TTz1) + TimeValue(TTz1)
    
    'test reboot set time forward
    'SysStartTimeNewN = SysStartTimeNewN + TimeSerial(0, 0, 10)
    
    If SysStartTimeNewN <> SysStartTimeOldN Then
        NewReboot = True
        freef5 = FreeFile
        Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\System UpTime Logg StartTime.txt" For Output Lock Write As #freef5
        Print #freef5, SysStartTimeNewN
        Close #freef5
    End If
    
    
    If NewReboot = True Then
    
    On Error Resume Next
    'Do
    'hag = hag + 1
    'Err.Clear
    'FS.FileCopy App.Path + "\00_Text_Data\System UpTime Logg Descrip.txt", "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\00_Text_Data\System UpTime Logg Descrip Old.txt"
    ''Err.Description
    'Call Sleep(2)
    'If hag + 1000 Then MsgBox "Error Delay Here Copy File": Stop
    'Loop Until Err.Number = 0
    
    'Close #freef5
    
    Err.Clear
    On Error Resume Next
    freef5 = FreeFile
    Tx2$ = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\System UpTime Logg Descrip.txt"
    Open Tx2$ For Input Lock Write As #freef5
    yyg = Input(LOF(freef5) - 1, freef5)
    Close #freef5
    'err.description
    If Err.Number > 0 Then
        freef5 = FreeFile
        Tx2$ = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\System UpTime Logg Descrip.bak.txt"
        Open Tx2$ For Input Lock Write As #freef5
        yyg = Input(LOF(freef5), freef5)
        Close #freef5
    End If
    
    FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\System UpTime Logg.txt"
    FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4

    freef5 = FreeFile
    Open FILENAME_IN_USE_CHECK For Append Lock Write As #freef5
    Print #freef5, "*******************************************************************"
    Print #freef5, yyg;
    Print #freef5, "*******************************************************************"
    Close #freef5
    
    freef5 = FreeFile
    Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\System UpTime Logg Descrip Old.txt" For Output Lock Write As #freef5
    Print #freef5, "*******************************************************************"
    Print #freef5, yyg;
    Print #freef5, "*******************************************************************"
    Close #freef5
    
    End If

End If





TickGo = True


Call moonnext

'Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

HwndCTLTask2 = FindWindow(vbNullString, "CTLTask")
HwndCTLTask3 = FindWindow(vbNullString, "Caller ID")
HwndCTLTask4 = FindWindow(vbNullString, "WindowsFormsParkingWindow")



'------------------------
' SUNRISE BONG
'------------------------
'Bv$ = "D:\0 00 MUSIC ---\09 SOUND EFFECT\VB Wave's\362 Sound Effects"
Bv$ = "D:\VB6\VB Wave's\362 Sound Effects"

drive2$ = Bv$
FILENAME2 = Bv$ + "\wow.wav"
FILENAME2 = App.Path + "\#Wave Sounds\wow.wav"
If Dir$(FILENAME2) <> "" Then
    Wqe1$ = FILENAME2
    ' Set properties needed by MCI to open.
    ATidalX.MMControl1.Notify = True
    ATidalX.MMControl1.Wait = True
    ATidalX.MMControl1.Shareable = False
    ATidalX.MMControl1.DeviceType = "WaveAudio"
    ATidalX.MMControl1.Filename = Wqe1$
    ' Open the MCI WaveAudio device.
    ATidalX.MMControl1.Command = "Open"
'    ATidalX.MMControl1.Command = "Play"
End If
   
'SUNRISE BONG
'FILENAME2 = Bv$ + "\ZTUTTI.WAV"
FILENAME2 = App.Path + "\#Wave Sounds\ZTUTTI.WAV"
If Dir$(FILENAME2) <> "" Then
    Wqe2$ = FILENAME2
    ' Set properties needed by MCI to open.
    ATidalX.MMControl2.Notify = True
    ATidalX.MMControl2.Wait = True
    ATidalX.MMControl2.Shareable = False
    ATidalX.MMControl2.DeviceType = "WaveAudio"
    ATidalX.MMControl2.Filename = Wqe2$
    'Open the MCI WaveAudio device.
    ATidalX.MMControl2.Command = "Open"
'    ATidalX.MMControl2.Command = "Play"
End If

'Moon
FILENAME2 = Bv$ + "\SNRoll-1.WAV"
FILENAME2 = App.Path + "\#Wave Sounds\SNRoll-1.WAV"

Wqe3$ = Trim$(Dir$(FILENAME2))
If Wqe3$ <> "" Then
    Wqe3$ = FILENAME2
    ATidalX.MMControl3.Notify = True
    ATidalX.MMControl3.Wait = True
    ATidalX.MMControl3.Shareable = False
    ATidalX.MMControl3.DeviceType = "WaveAudio"
    ATidalX.MMControl3.Filename = Wqe3$
    ' Open the MCI WaveAudio device.
    ATidalX.MMControl3.Command = "Open"
    'ATidalX.MMControl3.Command = "Play"
End If
   
Bv$ = "D:\0 00 MUSIC ---\09 SOUND EFFECT\VB Wave's\WINDOWS"
FILENAME2 = Bv$ + "\start.WAV"
If Dir$(FILENAME2) <> "" Then
    Wqe4$ = FILENAME2
    
    ATidalX.MMControl4.Notify = True
    ATidalX.MMControl4.Wait = True
    ATidalX.MMControl4.Shareable = False
    ATidalX.MMControl4.DeviceType = "WaveAudio"
    ATidalX.MMControl4.Filename = Wqe4$
    ' Open the MCI WaveAudio device.
    ATidalX.MMControl4.Command = "Open"
    'ATidalX.MMControl4.Command = "Play"
 
End If

'PAUSE MP3 WINAMP SOUND BUT NOT SURE IF IN USE MAYBE IN ANOTHER AREA
If Dir(App.Path + "\#Wave Sounds\01 Woarble Tone 5 Mins.wav") <> "" Then
    ATidalX.MMControl6.Notify = True
    ATidalX.MMControl6.Wait = True
    ATidalX.MMControl6.Shareable = False
    ATidalX.MMControl6.DeviceType = "waveaudio"
    'ATidalX.MMControl6.Filename = "D:\Wave's\Camera1a_2_8kHz.wav"
    ATidalX.MMControl6.Filename = App.Path + "\#Wave Sounds\01 Woarble Tone 5 Mins.wav"
    ' Open the MCI WaveAudio device.
    ATidalX.MMControl6.Command = "Open"
    ' ATidalX.MMControl6.Command = "Play"
End If




'MALICIOUS KNOCKER BOY SOUND NOT REQUIRED ANYMORE WORKS THOUGH
If 1 = 1 Then

    If Dir(App.Path + "\#Wave Sounds\DobFig4 01.WAV") <> "" Then
    
    
        ATidalX.MMControl7.Notify = True
        ATidalX.MMControl7.Wait = True
        ATidalX.MMControl7.Shareable = False
        ATidalX.MMControl7.DeviceType = "waveaudio"
        'ATidalX.MMControl7.FileName = App.Path + "\SG_02_04.WAV"
        ATidalX.MMControl7.Filename = App.Path + "\#Wave Sounds\DobFig4 01.WAV"
        ATidalX.MMControl7.Command = "Open"
        ' ATidalX.MMControl7.Command = "Play"
    End If
    
        
    If Dir(App.Path + "\#Wave Sounds\DobFig22 01.WAV") <> "" Then
        
        ATidalX.MMControl9.Notify = True
        ATidalX.MMControl9.Wait = True
        ATidalX.MMControl9.Shareable = False
        ATidalX.MMControl9.DeviceType = "waveaudio"
        'ATidalX.MMControl9.FileName = App.Path + "\SG_02_04.WAV"
        ATidalX.MMControl9.Filename = App.Path + "\#Wave Sounds\DobFig22 01.WAV"
        ATidalX.MMControl9.Command = "Open"
        ATidalX.MMControl9.Command = "prev"
        ' ATidalX.MMControl9.Command = "Play"
    End If
    
    'HOT KEY SOUND DON'T REQUIRE
    If Dir(App.Path + "\#Wave Sounds\Woarble Tone 5 Mins.wav") <> "" Then
    
        
        ' Open the MCI WaveAudio device.
        ATidalX.MMControl8.Command = "Open"
        ATidalX.MMControl8.Notify = True
        ATidalX.MMControl8.Wait = True
        ATidalX.MMControl8.Shareable = False
        ATidalX.MMControl8.DeviceType = "waveaudio"
        'ATidalX.MMControl8.FileName = App.Path + "\#Wave Sounds\HandsFree.WAV"
        ATidalX.MMControl8.Filename = App.Path + "\#Wave Sounds\Woarble Tone 5 Mins.wav"
    
        ' Open the MCI WaveAudio device.
        ATidalX.MMControl8.Command = "Open"
        'ATidalX.MMControl8.Command = "Play"
        'ATidalX.MMControl8.Command = "Pause"
    
    End If
End If


'ATidalX.DirectSS1.Pitch = 150
'ATidalX.DirectSS1.Speed = 200
'ATidalX.DirectSS1.VolumeLeft = 65535
'ATidalX.DirectSS1.VolumeRight = 65535

'form_menu.Check1.Value = vbChecked
'speedtime = -20

'Bv$ = App.Path
'Bvs$ = "\00_Text_Data\Tidal-X-Y.txt"

'Call Equinox3

FirstRun2 = 1

'If IsIDE = False Then Est = 1

ChDir App.Path

Call VolNibUP

'tool1 = 0: tool2 = 0
'speedtime = TimeSerial(11, 30, 0) * 50

'Load frmReceiver
'Load frmSender

'frmReceiver.Show
'frmSender_CallID.Show

'frmSender.txtMsg.Text = "Reset Tidal " + Time$
'Call frmSender.cmdSend_Click

'MsgBox "I"

End Sub



Sub ColourSet()
Exit Sub

Wtoolcol1 = &HC000&
Wtoolcol2 = &HC000&
Wtoolcol8 = &HFFFFFF
Wtoolcol11 = &HFF00&
Wtoolcol12 = &HFF00&
Wtoolcol13 = &HFFFF&
Wtoolcol20 = &H404040
Wtoolcol22 = &H404040
Wtoolcol23 = &HFF00&
Wtoolcol26 = &HFFFFFF

Command1.BackColor = Wtoolcol1
Command2.BackColor = Wtoolcol2
LABEL_TIME.BackColor = Wtoolcol8
Label11.BackColor = Wtoolcol11
Label12.BackColor = Wtoolcol12
Label13.BackColor = Wtoolcol13
Label20.BackColor = Wtoolcol20
Label22.BackColor = Wtoolcol22
Label23.BackColor = Wtoolcol23
Label26.BackColor = Wtoolcol26


End Sub



Public Sub Init()

Tidal_Mod.List_ActiveProcess 'tProcess
Tidal_Mod.List_ActiveModules 'tProcess

End Sub


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Exit Sub


FindCursor Tx, Ty

If Wtool = 1 Then Exit Sub

Wtool = 1

If MMControl4.mode = 525 And Soff = False Then
    ATidalX.MMControl4.Command = "prev"
    ATidalX.MMControl4.Command = "Play"
End If


End Sub


















Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub
FindCursor Tx, Ty

If Wtool = 2 Then Exit Sub

Wtool = 2

If MMControl4.mode = 525 And Soff = False Then
ATidalX.MMControl4.Command = "prev"
ATidalX.MMControl4.Command = "Play"
End If


End Sub

Private Sub Command1_Click()
extra = 0: Changeday = 0
Tog7 = Now + TimeSerial(0, 0, 5)

End Sub
Private Sub Command2_Click()
extra = 1: Changeday = 0
Tog7 = Now + TimeSerial(0, 0, 5)
End Sub

Private Sub Label1_Click()
'LockDontAllowHotKeyforPassword = True
'Sluty3 = 2
'Call KeyWindowCheck

'PUT BACK WHEN DONE
'Call Mnu_VolMinMusic_Click

Call Mnu_TimeLogg_Click

End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub

FindCursor Tx, Ty

If Wtool = 11 Then Exit Sub

Wtool = 11

If MMControl4.mode = 525 And Soff = False Then
    ATidalX.MMControl4.Command = "prev"
    ATidalX.MMControl4.Command = "Play"
End If


End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub

FindCursor Tx, Ty

If Wtool = 12 Then Exit Sub

Wtool = 12

If MMControl4.mode = 525 And Soff = False Then
    ATidalX.MMControl4.Command = "prev"
    ATidalX.MMControl4.Command = "Play"
End If

End Sub




Private Sub Label21_Change()
'Mouse
Vcodez = 0
Call KeyOrMouse(1)

End Sub

Private Sub Label23_Change()
'Key
KeyIdleTime = Now + TimeSerial(0, 0, 30)

'If Dir$("D:\Temp\KeyHit-Tidal-VCodes.txt") = "" Then popk = 1 Else popk = 0

Call KeyOrMouse(2)

If XmarkSpot = True Then
    XmarkSpot = False
    WinampHwnd = FindWindow("Winamp v1.x", vbNullString)

    For r = 0 To 100
        MsgResult = SendMessage(WinampHwnd, WM_COMMAND, Raisevolumeby1percent, ByVal XcNul)
    Next
    WinVol5 = 0
End If

End Sub

Sub KeyOrMouse(KeyMouse)

If Lab2$ = Label21.Caption + Label23.Caption Then Exit Sub
Lab2$ = Label21.Caption + Label23.Caption

KeyIdleTimeVideo = Now + TimeSerial(0, 30, 0)

If HittMan = True Then HittMan = 2

'If MMove2Time < Now Then
'MMove2 = Val(Label21) + 50 'reset the mouse moves every ten secs
'MMove2Time = Now + TimeSerial(0, 0, 5)
'End If

'Mouse10CLicks = False
'If Val(Label21) > MMove2 Then
'MMove2Time = Now + TimeSerial(0, 0, 5)
'Mouse10CLicks = True
'MMove2 = Val(Label21) + 50 ' ------- 200 clicks
'End If




If KeyMouse = 1 Then
'    ATidalZ.Label13 = Int(MouseClicks)
End If

gogo = False
If KeyMouse = 1 And MouseClicks > 3 Then gogo = True
If KeyMouse = 2 Then gogo = True


If MainIdleTime = 0 Then MainIdleTime = Now
IDLESCHEDULE = Now + TimeSerial(0, 10, 0)




RS$ = "  188 111 107 109 45 50 49 119 118 121 120 27 106 "
RS$ = RS$ + "122 109 191 123 44 144 117 115 145 115  119"
rs3$ = " " + Trim(str(Vcodez)) + " "
    
easy = True
If InStr(RS$, rs3$) > 0 Then easy = 0

If App.title <> "Tidal Information..." Then
    'Call frmPassLock.Labels_Tidal
    'If Test7 > Now Then
    '    Test7 = 0
    '    HittMan = False
'
'    End If

End If
'DeadLock = Now + TimeSerial(0, 30, 0)

If ProgressBarLock.Value = 0 And App.title = "Tidal Information..." Then
    If easy = True Then
        XxG = Now
        
        'Temp Disable put back if want
        'Call PassModule.Main
        
        HittMan = 2

    End If
End If

If gogo = True Then
    
    LockSSaver = Now + LockSaverDelay
    Call SetLockTime
    SetLockMax

    'never equal zero except start
    'if it does it put back up before get here again
    'If KeyMouse = 2 Then
    'qq8 = GetIdleM(2)
    'End If
    MainIdleTime = Now
    easy = True
    
    Test7 = 0
    Call SetLockMax
    Call ProgressLock
End If
End Sub

Function IsIdleTime(iQ2)
IsIdleTime = False
If MainIdleTime = 0 Then Exit Function
If DateDiff("s", MainIdleTime, Now) > iQ2 Then IsIdleTime = True

End Function



Sub Label26_CLICK()

MNU_SPEECH_OFF.Checked = False
Call SPEECH_OFF_SETTER

Call MNU_VOX_LIST_Click

VOXLIST_NOT_ACTIVE = 1000

MNU_SAY_MOON2_Click

End Sub

Private Sub Label26_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Beep

Exit Sub

'MAKE SOUND EFFECT RUTLAND

FindCursor Tx, Ty

If Wtool = 26 Then Exit Sub

Wtool = 26

If MMControl4.mode = 525 And Soff = False Then
    ATidalX.MMControl4.Command = "prev"
    ATidalX.MMControl4.Command = "Play"
End If

End Sub



Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub

FindCursor Tx, Ty

If Wtool = 20 Then Exit Sub

Wtool = 20

If MMControl4.mode = 525 And Soff = False Then
    ATidalX.MMControl4.Command = "prev"
    ATidalX.MMControl4.Command = "Play"
End If

End Sub


Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub


FindCursor Tx, Ty

If Wtool = 22 Then Exit Sub

Wtool = 22

If MMControl4.mode = 525 And Soff = False Then
    ATidalX.MMControl4.Command = "prev"
    ATidalX.MMControl4.Command = "Play"
End If

End Sub



Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Label13.ToolTipText = "Talk The Info Displayed Here Happen -- Press Shift While Click Hitt Button for Clipboard Info Extra Given About URL Eqinox"

'Equinox to Precision to Time in Day
'----
'Equinox -Wikipedia
'https://en.wikipedia.org/wiki/Equinox
'----

Exit Sub

FindCursor Tx, Ty

If Wtool = 13 Then Exit Sub

Wtool = 13

If MMControl4.mode = 525 And Soff = False Then
    ATidalX.MMControl4.Command = "prev"
    ATidalX.MMControl4.Command = "Play"
End If

End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If GetAsyncKeyState(16) < -300 Then ShiftPressed = True
If GetAsyncKeyState(161) < -300 Then ShiftPressed = True

If ShiftPressed = False Then Exit Sub
    
ShiftPressed = False

INF = "" + vbCrLf
INF = INF + "Well I Tuned Up My Mathematical And Have To Express It Accuracy Is Slight Inaccurate Due" + vbCrLf
INF = INF + "To A Limitation I Known For A Long Time" + vbCrLf
INF = "" + vbCrLf
INF = INF + "Explain Info Example" + vbCrLf
INF = INF + "This URL Is a Professional Even As My Found Sun Rise Set Set And Sun Noon Code Is Quiet Professional" + vbCrLf
INF = INF + "That Won't Give The Ability To Tell You What Equinox Is Of A Time In The Day" + vbCrLf
INF = INF + "Rather Only Rounded To The One Day" + vbCrLf
INF = INF + "As The Story Goes This Was All Hard Work Once Upon A Time" + vbCrLf
INF = INF + "Professional Sunrise Is Quite Extensive Mathematical And Timezone GPS Adjustable" + vbCrLf
INF = INF + "Insert A Info Once I Was Able To Calc Guess Point A Reference To Where The Time Difference Of Day Was" + vbCrLf
INF = INF + "Exactly Six Am To Six Pm 12 Hour Split Somewhere Out To The East Of The English Channel" + vbCrLf
INF = INF + "I Got An Image Of Map Where I Place The Point Somewhere" + vbCrLf
INF = INF + "----" + vbCrLf
INF = INF + "Getting Back" + vbCrLf
INF = INF + "And Extra Info -- My Calc Around Equinox Was Only Able To Tell The Winter Month Season" + vbCrLf
INF = INF + "Had A Different Day 21st And 22nd ----" + vbCrLf
INF = INF + "I Verified Them Again On-Line Source Working Told Military Navy At The Time Of Working" + vbCrLf
INF = INF + "And Quite Accurate" + vbCrLf
INF = INF + "But Limit Was Day And Not Time In Day Obviously Require A Lot More Mathematical" + vbCrLf
INF = INF + "Being More To Do With Equinox Than Sun Rise Sun Set And Sun Noon ----" + vbCrLf
INF = INF + "While Winter Shoed A Difference The Other Four Season Stay On Fixed Day" + vbCrLf
INF = INF + "----" + vbCrLf
INF = INF + "Next Part Tricky Bit ----" + vbCrLf
INF = INF + "Equinox In Winter Is The Shortest Day" + vbCrLf
INF = INF + "And Draw A Math Curve Of Difference Of Difference Around The Shortest Day" + vbCrLf
INF = INF + "Was Inaccurate To Find Where The Tide Had Changed And Start Going The Other Way" + vbCrLf
INF = INF + "Where The Curve Would Bounce Up Down And Not Ona One Day" + vbCrLf
INF = INF + "My Object Was The Find The Equinox So Then I Knew The System Of +93 Day To The" + vbCrLf
INF = INF + "Other Four Season In One Given Year" + vbCrLf
INF = INF + "Finally The Answer Was To Take It From The Solstice In Summer Longest Day Instead" + vbCrLf
INF = INF + "More Variation And Hitt On The Day Change Over" + vbCrLf
INF = INF + "Interesting That One" + vbCrLf
INF = INF + "----" + vbCrLf
INF = INF + "Equinox to Precision to Time in Day" + vbCrLf
INF = INF + "----" + vbCrLf
INF = INF + "Equinox -Wikipedia" + vbCrLf
INF = INF + "https://en.wikipedia.org/wiki/Equinox" + vbCrLf
INF = INF + "----" + vbCrLf

INF = INF + "Here Is Where My Sun-Rise Sun-Set Source Code" + vbCrLf
INF = INF + "Orginal Source Came From Top On The Vb6 Internet" + vbCrLf
INF = INF + "----" + vbCrLf
INF = INF + "Scott Seligman" + vbCrLf
INF = INF + "http://www.scottandmichelle.net/scott/" + vbCrLf
INF = INF + "Based off of" + vbCrLf
INF = INF + "http://www.srrb.noaa.gov/highlights/sunrise/gen.html" + vbCrLf
INF = INF + "----" + vbCrLf
INF = INF + "These are the Person I Talk Before Here On-Line WithOut Referance to Name" + vbCrLf
INF = INF + "They Talk About Withholding a KeyHooker of Design Feature Wanting for Strange Reason About Hacking----" + vbCrLf
INF = INF + "And About They Do Work in Fractal But in Movie For NAitmation Rather Than Mathatical Them" + vbCrLf
INF = INF + "Good Morning Monday 17 Oct 2k Sixteen" + vbCrLf

Beep
Clipboard.Clear
Clipboard.SetText INF

MsgBox "ClipBoarded Here ---" + vbCrLf + INF

'Want more info

End Sub

Private Sub Label13_Click()

MNU_SPEECH_OFF.Checked = False
Call SPEECH_OFF_SETTER

Call MNU_VOX_LIST_Click

VOXLIST_NOT_ACTIVE = 1000

If GetAsyncKeyState(16) < -300 Then Beep: Exit Sub 'ShiftPressed = True
If GetAsyncKeyState(161) < -300 Then Beep: Exit Sub  'ShiftPressed = True
'ShiftPressed = True
'Want more info
Beep

If 1 = 2 And RunOnce_Equinox = False Then
    RunOnce_Equinox = True
    If Equinox.Visible = True Then Unload Equinox: Exit Sub
    Unload Equinox: Equinox.Show
    'Equinox Tidal Info Form
End If


Beep
itext = Label13.Caption
'itext = Replace(itext, "DoD=", "Difference of Difference Between DayLight From_ Day Before, = ")
itext = Replace(itext, "DoD=", "Difference of Light From_ Day Before, = ")
itext = Replace(itext, "Darker --", "_ Darker __ ")
itext = Replace(itext, "Lighter --", "_ Lighter __ ")
itext = Replace(itext, "Equal --", "_ Equal __ ")
Min_var_text = " Minutes"
Sec_var_text = " Seconds "
If InStr(Label13.Caption, "1m") Then Min_var_text = " Minute"
If InStr(Label13.Caption, "0m") Then Min_var_text = " Minute"
If InStr(Label13.Caption, "1s") Then Sec_var_text = " Second "
If InStr(Label13.Caption, "0s") Then Sec_var_text = " Second "
itext = Replace(itext, "m ", Min_var_text + "_")
itext = Replace(itext, "s ", Sec_var_text)
'20 Mar 22 Sep Vernal/Autumnal Equinox
itext = Replace(itext, "__", " Difference Between Vernal(Spring) & Autumnal Equinox Half-a-Year Toward Shortest Day")
'20 Mar 22 Sep
itext = Replace(itext, Min_var_text, Min_var_text + " & ")
itext = Replace(itext, "_", " ")
'itext = Replace(itext, "!", "")
For r = 1 To 10
    ro1 = Len(itext)
    itext = Replace(itext, "  ", " ")
    If Len(itext) = ro1 Then Exit For
Next

itext1 = Mid(itext, 1, InStr(itext, " Difference Between V"))
itext2 = Mid(itext, InStr(itext, " Difference Between V") + 1)

Call SpeakVoice(Replace(itext1, "_", ""))

Call SpeakVoice(Replace(itext2, "_", ""))

End Sub

Private Sub Label20_Click()
Exit Sub

Hidecursor2 = Not Hidecursor2
If Hidecursor2 = True Then MsgBox "Hide Cursor Disabled": Label20.ToolTipText = "Enable Hide Cursor"
If Hidecursor2 = False Then MsgBox "Hide Cursor Enabled": Label20.ToolTipText = "Disable Hide Cursor"


End Sub

Private Sub Label22_Click()

'Shell "notepad " + App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\tidal key and mouse log.txt", vbNormalFocus

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
FILE_NAME = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\tidal key and mouse log.txt"
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing



End Sub

Private Sub Label23_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub

FindCursor Tx, Ty

If Wtool = 23 Then Exit Sub

Wtool = 23

If MMControl4.mode = 525 And Soff = False Then
ATidalX.MMControl4.Command = "prev"
ATidalX.MMControl4.Command = "Play"
End If


End Sub


Private Sub Label10_Click()

MNU_SPEECH_OFF.Checked = False
Call SPEECH_OFF_SETTER

'Label10.ToolTipText = "Clipboard Set Info About Moon"

Call MNU_VOX_LIST_Click

If GetAsyncKeyState(16) < -300 Then ShiftPressed = True
If GetAsyncKeyState(161) < -300 Then ShiftPressed = True
'ShiftPressed = True
'Want more info

REQUEST_NOT_HAVE_Angle_Of_Illumination = True
Call Lab10Sub

'If ShiftPressed = True Then ShiftPressed = False: Exit Sub
'----------------------
'If ShiftPressed = True -- extra info will be Appended

'SEPERATE SUB THIS DOES A CLIPBOARD ---- AND ONE OTHER DEPENDANT

'Call VOXLIST.MNU_SELECTION_Click

DoEvents

'Wd$ = Clipboard.GetText + vbCrLf + Wd$


Clipboard.Clear
Debug.Print Wd$
Wdmod$ = Replace(Wd$, "Now ", "Moment Here_ ")
Wdmod$ = Replace(Wdmod$, " @ ", "@ ")

Wdmod$ = Replace(Wdmod$, "First Quarter@", "First Quarter_@")
Wdmod$ = Replace(Wdmod$, "Next ", "Next Happen_ ")
'II11 = InStr(Wdmod$, "@")
'II13 = InStr(Wdmod$, vbCrLf)
'II14 = InStr(II13, Wdmod$, "@") 'Next Happen_
'If II11 > II14 Then
'    STRIN_LINE = String(II11, "_")
'End If
'If II14 > II11 Then
'    STRIN_LINE = String(II14 - II13, "_")
'    Wdmod$ = Replace(Wdmod$, "@ ", "@" + STRIN_LINE)
'End If



STR_LEN = II2

'Wdmod$ = Replace(Wdmod$, "in ", "Of ")
Wdmod$ = Replace(Wdmod$, "H-M-S", "HMS") 'Time Lenght")
Wdmod$ = Replace(Wdmod$, "Day ", "& Day ")
'Wdmod$ = Replace(Wdmod$, "Luminosity @The Moment_ ", "Luminousness ")
Wdmod$ = Replace(Wdmod$, "Luminosity Moment Here_ ", "Luminousness_ ")

Wdmod$ = Replace(Wdmod$, "Synodic _ 29.530588853 & Day", "Synodic _ 29.530588853 Day")


Wd$ = Wdmod$

Debug.Print vbCrLf
Debug.Print Wdmod$
Clipboard.SetText Wd$


If GetComputerName = "4-ASUS-GL522VW" = True Then Mnu_Solar_System = True

If Mnu_Solar_System = False Then
    
    Mnu_Solar_System = True
    
    MsgBox "Msg Box Here From Request Only Once at Start" + vbCrLf + vbCrLf + "ClipBoarded Here ---" + vbCrLf + Wd$

End If

Beep
'MsgBox "ClipBoarded Here ---" + vbCrLf + Wd$

End Sub

Sub Lab10Sub()
'Label10.ToolTipText = "Clipboard Set iNF_2o About Moon"

'On Error GoTo 0
'On Local Error GoTo 0


Wd$ = ""
Dim WdXx As String

WdXx = WdXx + Label24.Caption + " @ " + Format(X10, "DDD dd mmm hh:mm:ss") + vbCrLf
WdXx = WdXx + Label11.Caption + " @ " + Format(X9, "DDD dd mmm hh:mm:ss") + vbCrLf
'x10 when its a countdown
'WdXx = WdXx + "in     : " + Format$((x10 - Int(x10)) - (Now - Int(Now)), "HH:MM:SS") + " Mins/Secs" + vbCrLf
RR = DateDiff("d", ConvertToUT(Now), X9) - 1
If RR = 0 Then WdXx = WdXx + "in " + Format$((X9 - Int(9)) - (Now - Int(Now)), "HH:MM:SS") + " H-M-S" '+ vbCrLf
If RR = 1 Then WdXx = WdXx + "in" + str(RR) + " Day " + Format$((X9 - Int(9)) - (Now - Int(Now)), "HH:MM:SS") + " H-M-S" '+ vbCrLf
If RR > 1 Then WdXx = WdXx + "in" + str(RR) + " Days " + Format$((X9 - Int(9)) - (Now - Int(Now)), "HH:MM:SS") + " H-M-S" ' + vbCrLf
WdXx = WdXx + vbCrLf


MOON_LIGHT = Illum(ConvertToUT(Now))

'WdXx = WdXx + "Luminosity Now " + Format$(MOON_LIGHT, "###.00000") + "% Day Before " + Format$(Illum(ConvertToUT(Now) - 1), "###.00000") + "%" + vbCrLf
WdXx = WdXx + "Luminosity Now " + " " + MOON_STATE_LIGHTER + " " + Format$(MOON_LIGHT, "###.000") + "% Day Before " + Format$(Illum(ConvertToUT(Now) - 1), "###.0") + "%" + vbCrLf
'WdXx = WdXx + "Luminosity Now x0x__%" + vbCrLf

'WdXx = WdXx + "Luminosity Yester  " + Format$(Illum(ConvertToUT(Now) - 1), "###.00000") + "%" + vbCrLf
'wsf = Age(ConvertToUT(Now))

'If Int(wsf) >= 10 Then fullmoon$ = Format$(Int(wsf), " 00")
'If Int(wsf) < 10 Then fullmoon$ = Format$(Int(wsf), "  0")


'WdXx = WdXx + "New Moon Age : " + fullmoon$
'If ConvertToUT(Now) < 1 Then plur = "" Else plur = "'s"
'If ConvertToUT(Now) < 1 Then plur = "" Else plur = "_Plural"
If ConvertToUT(Now) < 1 Then plur = "" Else plur = "_P"

If ConvertToUT(Now) < 1 Then plur = "" Else plur = ""

plur = ""
plur = plur + vbCrLf
plur = plur + "-----------------------------------------" + vbCrLf
plur = plur + "Wikipedia __ Full_Moon_Cycle New Moon Age" + vbCrLf
plur = plur + "-----------------------------------------" + vbCrLf
plur = plur + "This The Recognise One 29 Day" + vbCrLf
plur = plur + "Synodic Month has an Average Duration" + vbCrLf
plur = plur + "SM = 29.530588853 Day" + vbCrLf
plur = plur + "-----------------------------------------" + vbCrLf
plur = plur + "Average Duration of the Anomalistic Month" + vbCrLf
plur = plur + "AM = 27.55454988 Day" + vbCrLf
plur = plur + "-----------------------------------------"
plur = ""


'2017 mar
'Plural in the Rural Royalty
'---------------------------

'Wikipedia __ Full_moon_cycle
'Average Duration of the Anomalistic Month
'AM = 27.55454988 Day
'Synodic Month has an Average Duration
'SM = 29.530588853 Day


'----
'Full moon cycle - Wikipedia
'https://en.wikipedia.org/wiki/Full_moon_cycle
'----
'The average duration of the anomalistic month is:
'AM = 27.55454988 days[3]
'The synodic month has an average duration of:
'SM = 29.530588853 days[4]


REQUEST_NOT_HAVE_Angle_Of_Illumination = False
If REQUEST_NOT_HAVE_Angle_Of_Illumination = True Then
    REQUEST_NOT_HAVE_Angle_Of_Illumination = False
    WdXx = WdXx + vbCrLf
Else
'-----------------------------------------------------------------------------------------------
'WdXx = WdXx + "New Moon Age " + Format$(Age(ConvertToUT(Now)), "0.00") + " Day" + plur + vbCrLf
'-----------------------------------------------------------------------------------------------
WdXx = WdXx + "New Moon Age Synodic _ 29.530588853 Day _ " + Format$(Age(ConvertToUT(Now)), "0.00") + plur + vbCrLf

WdXx = WdXx + "Angle Of Illumination " + Format$(Angle(ConvertToUT(Now)), "0.00") + vbCrLf + vbCrLf
'WdXx = WdXx + "Angle Of Illumination x0x" + vbCrLf + vbCrLf
End If

'-----------------------------------------
'REPLACE SET ARE HERE
'SOME ARE
'-----------------------------------------
'Label10_Click()
'-----------------------------------------



'Angle(AnyDate As Date) As Single

'WdXx = WdXx + vbCrLf

'If ConvertToUT(Now) = Now Then gmtbst = "GMT" Else gmtbst = "BST"
If ConvertToUT(Now) = Now Then GMTBST = "GMT UTC+0 WET" Else GMTBST = "BST_BDT_DST_DLS 1 Hour Ahead UTC+1 CET"

TimeZoneiNF_2o = "Uni-Time (UT GMT Solar Atomic) " + Format(ConvertToUT(Now), "DDD DD-MMM-YY hh:mm:ss")
'UKTimeZoneiNF_2o = "Uk TimeZone " + gmtbst + " ¦ 30 Oct Rewind 1 Hour GMT - Chunk of Time Disappeared - OverLapped"
UKTimeZoneiNF_2o = "The UK TimeZone " + GMTBST '+ " .. 25 Oct Rewind 1 Hour GMT - OverLap"

WdXx = WdXx + TimeZoneiNF_2o + vbCrLf
WdXx = WdXx + UKTimeZoneiNF_2o + vbCrLf

'From BST to GMT. ...

If ShiftPressed = True Then
Dim INF_2 As String

'More iNF_2o and URL iNF_2o Ahead in Sub
'-----------------------------------
INF_2 = "" + vbCrLf
INF_2 = INF_2 + "BST – British Summer Time / British Daylight Time -Daylight Saving Time) ... British Summer Time -Bst) Is 1 Hour Ahead Of Coordinated Universal Time." + vbCrLf
INF_2 = INF_2 + "BST_BDT_DST" + vbCrLf
INF_2 = INF_2 + "British Summer Time" + vbCrLf
INF_2 = INF_2 + "British Daylight Time" + vbCrLf
INF_2 = INF_2 + "Daylight Saving Time" + vbCrLf
INF_2 = INF_2 + "" + vbCrLf
INF_2 = INF_2 + "UTC+0 WET -- Western European Timezone" + vbCrLf
INF_2 = INF_2 + "----" + vbCrLf
INF_2 = INF_2 + "UTC" + vbCrLf
INF_2 = INF_2 + "Iceland" + vbCrLf
INF_2 = INF_2 + "Portugal - Azores" + vbCrLf
INF_2 = INF_2 + "2 Counter ----" + vbCrLf
INF_2 = INF_2 + "" + vbCrLf
INF_2 = INF_2 + "UTC+1 CET -- Central European Timezone" + vbCrLf
INF_2 = INF_2 + "----" + vbCrLf
INF_2 = INF_2 + "Canary Islands" + vbCrLf
INF_2 = INF_2 + "Channel Islands" + vbCrLf
INF_2 = INF_2 + "England" + vbCrLf
INF_2 = INF_2 + "Faroe Islands" + vbCrLf
INF_2 = INF_2 + "Ireland" + vbCrLf
INF_2 = INF_2 + "Isle Of Man" + vbCrLf
INF_2 = INF_2 + "Jersey" + vbCrLf
INF_2 = INF_2 + "Northern Ireland" + vbCrLf
INF_2 = INF_2 + "Portugal" + vbCrLf
INF_2 = INF_2 + "Scotland" + vbCrLf
INF_2 = INF_2 + "United Kingdom" + vbCrLf
INF_2 = INF_2 + "Wales" + vbCrLf
INF_2 = INF_2 + "12 Counter ----" + vbCrLf
INF_2 = INF_2 + "----" + vbCrLf
INF_2 = INF_2 + "Europe Time Zone - Europe Current Time" + vbCrLf
INF_2 = INF_2 + "http://www.timetemperature.com/europe/europe_time_zones.shtml" + vbCrLf
INF_2 = INF_2 + "----" + vbCrLf
End If

'--------------------------------------------------
'ORDER OF SERACH TO ENDING MOST USEFUL RESULT FOUND
'--------------------------------------------------
'----
'GMT Is The Standard Time Of Which Country - Google Search
'https://www.google.co.uk/search?num=50&q=GMT+Is+The+Standard+Time+Of+Which+Country&oq=GMT+Is+The+Standard+Time+Of+Which+Country&gs_l=serp.3..35i39k1.1581.2329.0.2815.4.4.0.0.0.0.149.467.1j3.4.0....0...1c.1.64.serp..0.1.102.4cquI5T4lMA
'----
'----
'Greenwich Mean Time - Wikipedia
'https://en.wikipedia.org/wiki/Greenwich_Mean_Time
'----
'----
'UTC+1 - Google Search
'https://www.google.co.uk/search?num=50&q=UTC%2B1&oq=UTC%2B1&gs_l=serp.3..35i39k1j0i20k1j0l8.8641.11389.0.12134.7.7.0.0.0.0.120.707.3j4.7.0....0...1c.1.64.serp..0.7.705...0i7i30k1.549hxz1UUYw
'----
'----
'Europe Time Zone - Europe Current Time
'http://www.timetemperature.com/europe/europe_time_zones.shtml
'----




'EXTRA iNF_2O WANTED ON BUTTON PRESSED
If ShiftPressed = True Then
    ShiftPressed = False
    
    WdXx = WdXx + INF_2 + vbCrLf

End If


Wd$ = WdXx


'MsgBox WdXx

'End 'endhere
'Stop
'End

End Sub

Public Sub Label25_Click()
'Call Mnu_ClipURL_Click
    
MNU_SPEECH_OFF.Checked = False
Call SPEECH_OFF_SETTER
    
wd2$ = "Lumin_oss_ity  : " + Format$(Illum(ConvertToUT(Now)), "#00") + "%" + vbCrLf
    
'Call VoiceWait
Call VolumeFixedSet(10, True)
'Call VoiceMP3Pause
Call SpeakVoice(wd2$)
'If xLATCHTIMEROverRide = True Then LATCHTIMEROverRide = True




End Sub


Public Sub Label27_Click()
    
    'Call Label10_Click
    MHDR_NOW_DATE = -1

    Beep
    If IsIDE = True Then End

End Sub

Private Sub Label27_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 And oButton <> Button Then
    'Shell "NOTEPAD " + App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Voice_Talk_Text.txt", vbMaximizedFocus
    FILE_NAME = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Voice_Talk_Text.txt"
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.RUN """" + FILE_NAME + """"
    Set WSHShell = Nothing

End If

oButton = Button
End Sub

Private Sub Label27_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
oButton = -2

End Sub

Private Sub Label30_Click()
'Label30.
End Sub

Private Sub Label31_Click()
'label31

If SystemUptime.Days > 1 Then ii$ = " Days " Else ii$ = " Day "
MsgBox "SysUptime " & SystemUptime.Days & ii$ & _
        SystemUptime.Hours Mod 24 & "h " & _
        SystemUptime.Minutes & "m " & _
        SystemUptime.Seconds & "s " & _
        Format$(SystemUptime.MSeconds, "000") & " mil"


End Sub

Private Sub Label32_Click()
'Label32
End Sub

Private Sub Label33_Click()
'Label33
End Sub

Private Sub Label37_Click()
'    Label37.ToolTipText = "SOUND MALICIOUS RUTLAND"
    
    If Menu.SoundMall = vbChecked Then
        Menu.SoundMall = vbUnchecked
        ''Call ATidalX.SoundOffOn("OFF")
        ATidalX.Label37.BackColor = ATidalX.Label38.BackColor
    Else
        Menu.SoundMall = vbChecked
        'Call ATidalX.SoundOffOn("ON")
        ATidalX.Label37.BackColor = ATidalX.Label39.BackColor
    End If

End Sub

Private Sub Label4_Click()
'Label4.ToolTipText = "ENTER EXTREMER LOCK"
Call Mnu_ELock_Click
Exit Sub

'Call UpDate5(0)

tr$ = ""

SayTime = True 'make say time?

List1.AddItem Format(Now, "DD-MMM-YY HH:MM:SSa/p")

FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\SAY Time Key Press Logg.txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4

Fr8 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #Fr8
    Print #Fr8, Format(Now, "DD-MMM-YY -- HH:MM:SSa/p")
Close #Fr8


If UpDateText.Visible = True Then UpDateText.Hide: Exit Sub

'UpDateText.Show


End Sub



Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Exit Sub

FindCursor Tx, Ty

'If wtool = 4 Then Exit Sub

Wtool = 4

If MMControl4.mode = 525 And Soff = False Then
ATidalX.MMControl4.Command = "prev"
ATidalX.MMControl4.Command = "Play"
End If


End Sub



 
Private Sub Label9_Click()
Call Label9_DblClick
Exit Sub

'LockDown = True
'LockSaverDelay = LockSaverDelay3
'LockSSaver = Now + LockSaverDelay
End Sub

Private Sub Label9_DblClick()
If LockDown2 > 0 Then
    LockDown2 = 0: LockDown = False
    LockSaverDelay = LockSaverDelay2
    LockDownX = 0
    Exit Sub
End If
LockDown2 = Now + TimeSerial(0, 20, 0)
LockDownX = True
LockSaverDelay = LockSaverDelay4 '--- Short delay even shorter for IDE

'chk here Timer1_Timer
End Sub

Public Sub LblVol_Change()
    
    'Exit Sub
    
    On Local Error Resume Next
    'If TEST1 = LblVol Then Debug.Print LblVol
    'TEST1 = LblVol
    
    
    'If Val(ATidalX.LblVol) = 100 Then Stop
    
    'Debug.Print LblVol
    
    Call TimerVOLCHANGETALK_TRIG_START


    If Val(ATidalX.LblVol) <> 100 Then
    
        LblVol.BackColor = RGB(255, 80, 80)
        
    Else
    
        LblVol.BackColor = &H680202
    
    End If
    

    
    
    

    
    If Val(ATidalX.LblVol) < 0 Then ATidalX.LblVol = "0"
    If Val(ATidalX.LblVol) > 100 Then ATidalX.LblVol = "100"

    If WinAmpChkIsPlaying = 1 Then
            VolumeSet1 = Val(ATidalX.LblVol)
            ATidalX.ProgressBarVol1.Value = VolumeSet1 'ATidalX.LblVol
        Else
            VolumeSet2 = Val(ATidalX.LblVol)
            ATidalX.ProgressBarVol2.Value = VolumeSet2 'ATidalX.LblVol
            ' FIRST SET BY FORM
            ' -----------------
            OVOLTAG = VolumeSet2
    End If
    
    If MNU_BOTH_VOL_SAME.Checked = True Then
            VolumeSet1 = Val(ATidalX.LblVol)
            ATidalX.ProgressBarVol1.Value = VolumeSet1 'ATidalX.LblVol
            VolumeSet2 = Val(ATidalX.LblVol)
            ATidalX.ProgressBarVol2.Value = VolumeSet2 'ATidalX.LblVol
    End If
                
                
    If MNU_NO_SET_VOL_NONE_SET_AS_1.Checked = True Then
        If Val(ATidalX.LblVol) = 0 Then ATidalX.LblVol = "1"
    End If
                
    'CHK HERE -- MP3_OFF__ON_START_VOICE()
                
                
    'WHY HAD THIS DUPE BELOW
    If OVOLTAG2 <> VolumeSet2 And OVOLTAG2 > 0 And Me.Visible = True Then
        If WinAmpChkIsPlaying <> 1 Then
            'TRY BACK IN WAS FAULT SAY VOL IN MP3
            ForceSayVol = True
        End If
                
    End If
    OVOLTAG2 = VolumeSet2
                
                
    'XX20 = VoiceStreamActive = False
    'XX21 = StringVoiceToPLayActive = False
    
    XX21 = WinAmpChkIsPlaying <> 1
    XX22 = ISSUECOMMTOSTOPWINAMP = False
    XX20 = XX21 And XX22
    

    If TimerVOLCHANGETALK2.Enabled = False Then
        If OVOLTAG >= 0 And OVOLTAG = 0 And (XX20 Or ForceSayVol = True) Then
            ForceSayVol = False
            Call TimerVOLCHANGETALK_TRIG_START
            
            'Debug.Print "Old Vol Set VolumeSet2= " + str(VolumeSet2)
            'OVOLTAG = VolumeSet2 'Val(ATidalX.LblVol)
        End If
    End If
    If OVOLTAG = -1 Then OVOLTAG = 0
    ISSUECOMMTOSTOPWINAMP = False
        
    'ISSUECOMMTOSTOPSTARTWINAMP = False
                
    tf = SetVolume(SPEAKER, ATidalX.LblVol)
    Slider1.Value = Val(ATidalX.LblVol)
    
    'Call TimerDoubleVolume_Timer

End Sub

Private Sub LblVol_Click()
'Shell "notepad2 """+APP.PATH+"\" + GetComputerName + "-" + GetUserName + "\Vol Change Logg.txt""", vbNormalFocus
'VolumeSet1 = 1
'FILE_NAME = ""
'WSHShell.Run """" + FILE_NAME + """"
'Set WSHShell = Nothing

If Val(ATidalX.LblVol) <> 100 Then ATidalX.LblVol = "100"

End Sub


Private Sub Mnu_About_Click()
Call UpDate5(1)
End Sub

Private Sub MNU_BOTH_VOL_SAME_Click()
    MNU_BOTH_VOL_SAME.Checked = Not MNU_BOTH_VOL_SAME.Checked
End Sub

Private Sub Mnu_ClipBoard_Info_Block_Of_Moon_Click()
Call Label10_Click
End Sub

Private Sub MNU_ENDER_MNU_Click()
End
End Sub

Private Sub Mnu_Exit_Click()

EXIT_FORM_SET = True

'FORM_UNLOAD
Unload Me

End Sub

Private Sub MNU_JOY_PADDER_1_Click()
If frmDXGAMEJOY.JOYPAD_DX_LOADED = False Then
    Load frmDXGAMEJOY
End If
frmDXGAMEJOY.Show

Load FrmJoy
FrmJoy.Show

End Sub

Private Sub MNU_JOY_PADDER_Click()

Load frmDXGAMEJOY
frmDXGAMEJOY.Show
Load FrmJoy
'If IsIDE = True Then FrmJoy.Show

End Sub

Private Sub MNU_JOYPAD_SOFTWARE_STATUS_Click()
Beep
Unload frmDXGAMEJOY
Load frmDXGAMEJOY
End Sub

Private Sub MNU_LATENIGHTOVERRIDE_Click()
    MNU_LATENIGHTOVERRIDE.Checked = Not MNU_LATENIGHTOVERRIDE.Checked
End Sub

Private Sub Mnu_MINIMIZE_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub MNU_NO_PAUSE_Click()
    MNU_NO_PAUSE.Checked = Not MNU_NO_PAUSE.Checked
End Sub

Private Sub MNU_SAY_MOON2_Click()

'TALK MOON %
Call MNU_VOX_LIST_Click

VOXLIST_NOT_ACTIVE = 1000

Call Moon_Code

FLAG_TO_MOON_TALK = True

Call Percent_Moon_1_100(MOON_LIGHT)

FLAG_TO_MOON_TALK = False

End Sub

Private Sub MNU_NOT_DOOR_TALK_Click()
    MNU_NOT_DOOR_TALK.Checked = Not MNU_NOT_DOOR_TALK.Checked
End Sub

Private Sub MNU_SAY_MOON_Click()
'----------



End Sub

Public Sub MNU_SPEECH_OFF_Click()
    MNU_SPEECH_OFF.Checked = Not MNU_SPEECH_OFF.Checked
    Call SPEECH_OFF_SETTER
End Sub
Public Sub SPEECH_OFF_SETTER()
    
    If MNU_SPEECH_OFF.Checked = True Then
        MNU_NO_PAUSE.Checked = True
        MNU_STOP_WINAMP_ON_THE_HOUR.Checked = False
    Else
        MNU_NO_PAUSE.Checked = False
        MNU_STOP_WINAMP_ON_THE_HOUR.Checked = True
    End If

End Sub

Private Sub MNU_STOP_WINAMP_ON_THE_HOUR_Click()
    MNU_STOP_WINAMP_ON_THE_HOUR.Checked = Not MNU_STOP_WINAMP_ON_THE_HOUR.Checked
End Sub

Private Sub MNU_TALKER_EVERY_MINUTE_FOR_AN_HOUR_Click()

VAR_TALKER_EVERY_MINUTE_FOR_AN_HOUR = Now + TimeSerial(1, 0, 0)
I_R = "Talker Every Minute For An Hour -- " + Format(VAR_TALKER_EVERY_MINUTE_FOR_AN_HOUR, "HH-MM-SS AMPM")
MNU_TALKER_EVERY_MINUTE_FOR_AN_HOUR.Caption = I_R
Beep


End Sub

Private Sub MNU_VERSION_Click()

MNU_VERSION.Caption = "Ver. " + Trim$(str$(App.Major)) + "." + Trim$(str$(App.Minor)) + "." + Trim$(str$(App.Revision))

End Sub

Private Sub MNU_VOL_100_Click()

If Val(ATidalX.LblVol) <> 100 Then ATidalX.LblVol = "100"

End Sub

'Private Sub MNU_VB_FOLDER_Click()
'
'Shell "Explorer.exe /SELECT, """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
'Unload Me
'
'End Sub

Private Sub MNU_VOX_LIST_Click()


'Call MNU_VOX_LIST_Click

VOXLIST.Show
VOXLIST.WindowState = vbNormal
VOXLIST.Timer1 = True

For r = 1 To 10
    If VOXLIST.WindowState <> vbNormal Then
    
        VOXLIST.Show
        VOXLIST.WindowState = vbNormal
        DoEvents
    End If
Next


End Sub

Private Sub MNU_VOX_LIST2_Click()

Call MNU_VOX_LIST_Click

'VOXLIST.Timer2.Enabled = False

End Sub

Private Sub ProgressBar2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'ProgressBar2
End Sub

Public Sub Timer_REINTRODUCE_DSkeybd_M_Timer()


XXIsVBrunningForCode = IsVBrunningForCode


'Debug.Print "XXIsVBrunningForCode " + str(XXIsVBrunningForCode)
'List2.AddItem "XXIsVBrunningForCode " + str(XXIsVBrunningForCode)
'If XXIsVBrunningForCode = False Then
'    Exit Sub
'End If

'List2.ListIndex = List2.ListCount - 1

'If VAR_REINTRODUCE_DSkeybd_M = False Then
'    VAR_REINTRODUCE_DSkeybd_M = True
    
    'Unload DSkeybd_F
If XXIsVBrunningForCode = True Then
    
    If DSkeybd_F.IsHook = True Then
        Call DSkeybd_F.UnHookCommand_Click
        ASYNC_Key_Press_Timer.Enabled = True

        Exit Sub
    
    End If

End If

If XXIsVBrunningForCode = False Then
    'ATidalX.Timer_REINTRODUCE_DSkeybd_m.Enabled = False
    'Call Mnu_KeyHook_Click
    'Load DSkeybd_F
    

    If DSkeybd_F.IsHook = False Then
        '-----------------------------------
        'PROBLEM FIND OUT WHAT THESE ARE FOR
        'TestKeyLoggOff = False
        'TestHookinVB = True
        '-----------------------------------
    
        
    
        ASYNC_Key_Press_Timer.Enabled = False
        Load DSkeybd_F
        Call DSkeybd_F.Command1_Click
    
    End If
End If


End Sub

Private Sub TIMER_SECONDS20_AFTER_HOUR_TIMER_DOOR_TALKER_Timer()

TIMER_SECONDS20_AFTER_HOUR_TIMER_DOOR_TALKER.Enabled = False

DayOrMidNightVol = 20
Call VolumeFixedSet(DayOrMidNightVol, True)

Call Label40_Click

End Sub

Public Sub TIMER_WinAmpPause_Timer()

Call aa_MainCode.WinAmpPause

End Sub

Public Sub TIMER_WinAmpPrevTrack_Timer()

Call aa_MainCode.WinAmpPrevTrack

End Sub

Public Sub TIMER_WinAmpNextTrack_Timer()

Call aa_MainCode.WinAmpNextTrack

End Sub


Private Sub Timer17_Timer()
'MINIMIZE VB IDE AT RUN

Timer17.Enabled = False
Exit Sub

'Me.WindowState = vbMinimized
Tiger = FindWinPart("Tidal_Information - Microsoft Visual Basic")
    
'ShowWindow Tiger, SW_SHOWMINNOACTIVE
ShowWindow Tiger, SW_MINIMIZE
    
    
    
'lngI = SetFocuses(PopBack)
'SetForegroundWindow PopBack
'ingl = Putfocus(PopBack)



End Sub

Private Sub Timer18_Timer()

End Sub

Private Sub TimerONCEADAY_Timer()
Exit Sub
Dim PIO
PIO = False

If MHDR_NOW_DATE = -1 Then PIO = True
If RTV1 < Now Then PIO = True
'If RTV2 < Now Then PIO = True

If PIO = True Then
    
    If RTV2 > 45 Then RTV2 = 0
    For r = 0 To 59 Step 2
        RTV2 = r
        If r > Second(Now) Then Exit For
    Next
    RTV1 = Int(Now) + TimeSerial(HOUR(Now), Minute(Now), RTV2)
    
    'If RTV2 < Now Then
    '    RTV2 = Int(Now) + TimeSerial(HOUR(Now), Minute(Now) + 1, 30)
    'End If

    MHDR_NOW_DATE = 0
    MHDR_ODD_MIN = MHDR_ODD_MIN + 1
    'If MHDR_ODD_MIN = 1 Then MHDR = "Male! Has Do, RiPer!"
    'XM = XM + 1
    'If MHDR_ODD_MIN = 2 Then MHDR = "Three! RIP! RiPer!"
    'XM = XM + 1
    'If MHDR_ODD_MIN = 3 Then MHDR = "Male Has Not Proper With Trace Saturate JiTer!"
    'XM = XM + 1
    'If MHDR_ODD_MIN = 4 Then MHDR = "Beyond Learn With Female Riper!"
    'XM = XM + 1
    'If MHDR_ODD_MIN = 5 Then MHDR = ""
    'XM = XM + 1
    
    Dim XTELL()
    ReDim XTELL(4000)
    XTELL(1) = "RiPer!"
    XTELL(2) = "Voice! Pattern!"
    XTELL(3) = "JiTer!"
    XTELL(4) = "Muling!"
    
    XTELL(5) = "NORTH SIDE! , EAST! , SOUTH! , WEST SIDE!"
    'XTELL(6) = "NORTH EAST SIDE,, SOUTH EAST,, SOUTH WEST,, NORTH WEST SIDE!"
    xt = 5
    
    
    XTELL(xt) = "Splendid!"
    xt = xt + 1:    XTELL(xt) = "Cybernetics!"
    xt = xt + 1:    XTELL(xt) = "NetOMatic!"
    xt = xt + 1:    XTELL(xt) = "Neodymium!"
    xt = xt + 1:    XTELL(xt) = "Precocious!"
    xt = xt + 1:    XTELL(xt) = "Hot Potato!"
    xt = xt + 1:    XTELL(xt) = "Basket Ball!"
    xt = xt + 1:    XTELL(xt) = "Volley Ball!"
    xt = xt + 1:    XTELL(xt) = "Pancake Race!"
    xt = xt + 1:    XTELL(xt) = "Cartridge Load!"
    xt = xt + 1:    XTELL(xt) = "Three! RIP! RiPer!"
    'xt = xt + 1:    XTELL(xt) = ""
    'xt = xt + 1:    XTELL(xt) = ""
    'xt = xt + 1:    XTELL(xt) = ""
    'xt = xt + 1:    XTELL(xt) = ""
    'xt = xt + 1:    XTELL(xt) = ""
    
    
    F1 = FreeFile
    Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Voice_Talk_Text.txt" For Input As #F1
        Do
            Line Input #F1, io
            X22 = 0
            If Trim(io) <> "" Then X22 = 1
            If Mid(io, 1, 1) = "#" Then X22 = 0
            If X22 = 1 Then
                xt = xt + 1:    XTELL(xt) = io
            End If
            
        Loop Until EOF(F1)
    Close #F1
    
    For xt = 1 To UBound(XTELL)
        If XTELL(xt) = "" Then xt = xt - 1: Exit For
    Next
    ReDim Preserve XTELL(xt)

    
    If MHDR_ODD_MIN >= UBound(XTELL) Then MHDR_ODD_MIN = 0
        MHDR = XTELL(MHDR_ODD_MIN)
    MHDR_ODD_MIN = Int(UBound(XTELL) * Rnd) + 1
    
    Call SpeakVoice(MHDR)
End If


If TALKEQU1 <> "" Then
If Minute(Now) = 3 And Second(Now) = 11 And YYFLAG1 = 0 Then
    YYFLAG1 = 1
    Call SpeakVoice(TALKEQU1)
    Debug.Print "ENTER HERE -- " + TALKEQU1
End If

If Minute(Now) = 3 And Second(Now) = 20 And YYFLAG2 = 0 Then
    YYFLAG2 = 1
    Call SpeakVoice(TALKEQU2)
    Debug.Print "ENTER HERE -- " + TALKEQU2
End If
If Minute(Now) > 3 Then YYFLAG1 = 0: YYFLAG2 = 0
End If


'PUT ON BIT FREQ COZ HAVE PROB THIS FROM OTHER PROG
If Minute(Now) Mod 10 = 0 Then ODAYONCEADAY = -1

Exit Sub

If ODAYONCEADAY = Day(Now) Then Exit Sub



ODAYONCEADAY = Day(Now)
Call FileInUseDelay("D:\TEMP\Equinox.txt", True)
On Error Resume Next
FR = FreeFile
Open "D:\TEMP\Equinox.txt" For Input As #FR
    Line Input #FR, inp1
    Line Input #FR, inp2
Close #FR

'20-Mar-10 - Vernal Equinox
'21-Jun-10 - Summer Solstice

'THE Vernal Equinox, WAS 20TH MAR
'EQU1=

XX1 = Mid(inp1, 1, 6)
XX2 = Mid(inp2, 1, 6)

'Summer Solstice IS 21th Jun Called The Longest Day, And Made Famous By William Shakespear!"

dg1 = "The Summer! Solstice! Will Be On 21st Of June, Known as The Longest! Day And Made Famous! By William Shakespear!"
dg2 = "The Vernal! Equinox! Now! Was On The 20th Of March!"

TALKEQU1 = ""
TALKEQU2 = ""
If XX1 = "21-Jun" Then
    TALKEQU1 = dg1
End If
If XX2 = "21-Jun" Then
    TALKEQU2 = dg1
End If

If XX1 = "20-Mar" Then
    TALKEQU1 = dg2
End If
If XX2 = "20-Mar" Then
    TALKEQU2 = dg2
End If


'TEMP TAKE OUT WE KNOW ERRORS AT MOMENT
'If TALKEQU2 = "" Or TALKEQU1 = "" Then msgbox "Stop Not Correct Date": Stop



'15-Jan-10 07:12:00a GMT - Last # Solar Eclipse -Annular
'26-Jun-10 11:30:00a GMT - Next # Lunar Eclipse -Partial(Umbral)

'D:\TEMP\Eclipse.txt

Call FileInUseDelay("D:\TEMP\Eclipse.txt", True)
FR = FreeFile
Open "D:\TEMP\Eclipse.txt" For Input As #FR
    Line Input #1, inp1
    Line Input #1, inp2
Close #1

XX1 = Mid(inp1, 1, 9)
XX2 = Mid(inp2, 1, 9)
XX4 = DateValue(XX1)
XX5 = DateValue(XX2)


IDATESTR = ENGLISH_DATE_FOR_TALK(XX4, False)





End Sub


Function ENGLISH_DATE_FOR_TALK(INPUTDATE, LONGDAY As Boolean)
'JUST DAY AND MONTH FOR NOW
'FORMAT LIKE
'11TH OF JUNE


Dim FFT1, FFT2
FFT = Format$(INPUTDATE, "DD")
XX1 = Format$(INPUTDATE, "D")
Tx2 = "TH"
XX = Val(Mid(FFT, 2, 1))
If FFT = 11 Or FFT = 12 Or FFT = 13 Then XX = 0
If XX = 1 Then Tx2 = "ST"
If XX = 2 Then Tx2 = "ND"
If XX = 3 Then Tx2 = "RD"
FFT = Format$(INPUTDATE, " - MMMM")
If LONGDAY = True Then FFT = Format$(Now, "DDDD - MMMM")

FFT1 = Replace(FFT, "-", XX1 + LCase(Tx2) + " of")

ENGLISH_DATE_FOR_TALK = FFT1

End Function



Sub TimerVOLCHANGETALK_TRIG_START()
    'Exit Sub
    
    If ISSUECOMMTOSTOPWINAMP = True Then Exit Sub
    If WinAmpChkIsPlaying = 1 And ISSUECOMMTOSTOPWINAMP = False Then Exit Sub
    'If WinAmpChkIsPlaying = 1 And ForceSayVol = False Then Exit Sub
    'ForceSayVol = False

    If TimerVOLCHANGETALK2.Enabled = True Then Exit Sub
    TimerVOLCHANGETALK.Enabled = False
    TimerVOLCHANGETALK.Interval = 0
    TimerVOLCHANGETALK.Interval = 4000
    TimerVOLCHANGETALK.Enabled = True

End Sub


Private Sub TimerVOLCHANGETALK_Timer()

'TO LOOK LBLVOL_CHANGE

TimerVOLCHANGETALK.Enabled = False

If NOVOLSAYASSTART = False Then Exit Sub


'if isplay
'TT = Val(ATidalX.LblVol) - OVOLTAG
        
'Debug.Print "NEW Vol Set VolumeSet2= " + str(VolumeSet2)
'Debug.Print "NEW Vol Set LBLVOL= " + LblVol

If OVOLTAG = VolumeSet2 Then Exit Sub
TT = VolumeSet2 - OVOLTAG

'Debug.Print "NEW Vol CHANGE = " + str(TT)

If TT > 0 Then
    Speak = "Volume Higher! By" + str(Abs(TT)) + ", To! , " + str(VolumeSet2)
End If
If TT < 0 Then
    Speak = "Volume DROP! By" + str(Abs(TT)) + ", To! , " + str(VolumeSet2)
End If
'If TT < 0 Then Speak = "Volume Lower By" + str(Abs(TT))

' OVOLTAG = 0

OVOLTAG = VolumeSet2 'Val(ATidalX.LblVol)

'Debug.Print "FORCESAYVOL   " + str(ForceSayVol)
ForceSayVol = False

If TT = 0 Then Exit Sub



'Debug.Print Speak

'-------------------
'VOLUME HIGHER LOWER
'-------------------
Call SpeakVoice("TIME #" + Speak)


'--------------------------------

Exit Sub

'--------------------------------

Call FileInUseDelay(App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Volume.txt", False)
FR4 = FreeFile
Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Volume.txt" For Output As #FR4
Print #FR4, str(VolumeSet1)
Print #FR4, str(VolumeSet2)
Close #FR4


End Sub



Sub WaitVOICEFinish()

    '--------------------------------------------
    Exit Sub
    '--------------------------------------------

'    Do
'        XX20 = VoiceStreamActive
'        'XX21 = StringVoiceToPLayActive
'        'If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then Exit Do
'        If XX20 = False And XX21 = False Then Exit Do
'        Sleep 20
'        DoEvents
'    Loop Until False

End Sub


Sub TimerVOLCHANGETALK2_TRIG_START()

    TimerVOLCHANGETALK2.Enabled = False
    TimerVOLCHANGETALK2.Interval = 0
    TimerVOLCHANGETALK2.Interval = 4000
    TimerVOLCHANGETALK2.Enabled = True

End Sub



Private Sub TimerVOLCHANGETALK2_Timer()

TimerVOLCHANGETALK2.Enabled = False

Call SpeakVoice("Volume Higher By 2")

End Sub

Function FILTER_SPEAK_LIST(Tx, VoiceToSpeak)
'PROBLEM
'DONT KNOW WHY ANYMORE
'RUBBISH
TRIGV = False
FILTER_SPEAK_LIST = False
If Tx = 1 Then

    For r = 0 To ListVOICE.ListCount - 1
        If InStr(ListVOICE.List(r), MHDR) > 0 Then FILTER_SPEAK_LIST = True
        If InStr(ListVOICE.List(r), "TIME #") > 0 Then
            FILTER_SPEAK_LIST = True
        End If
    Next
    If FILTER_SPEAK_LIST = False Then Exit Function
End If

If Tx = 2 Then
    VoiceToSpeak = ""

    Do
        VoiceToSpeak = ListVOICE.List(0)
        ListVOICE.RemoveItem (0)
        If InStr(VoiceToSpeak, MHDR) > 0 Then Exit Do
        If InStr(VoiceToSpeak, "VOLUME") > 0 Then Exit Do
        If InStr(VoiceToSpeak, "TIME #") > 0 Then
        VoiceToSpeak = Mid(VoiceToSpeak, 7)
        FILTER_SPEAK_LIST = True
        End If
        If ListVOICE.ListCount = 0 Then Exit Do
    Loop Until 1 = 2

    If VoiceToSpeak <> "" Then
          FILTER_SPEAK_LIST = True
    End If

End If

End Function

Sub n_TIMER()

'NOT REAL

Exit Sub



'WHY FILTER
'TRIGV = FILTER_SPEAK_LIST(1, VoiceToSpeak)


If ATidalX.MNU_SPEECH_OFF.Checked = True And TRIGV = False Then Exit Sub


'XX1 = WINAMPBackOnAfterVoiceRequired = True
'XX2 = VoiceStreamEnds = True
'XX3 = StringVoiceToPLayActive = False
'XX4 = ListVOICE.ListCount = 0

'----------------------------------------------------
'MODIFY TO USE IN EVENT

If XX1 And XX2 And XX3 And XX4 Then
    
    'WINAMPBackOnAfterVoiceRequired = False
    'VoiceStreamEnds = False
    'Call MP3_ON_AFTER_VOICE_FINISHES

End If
    
'-----------------------------------------------------------

'STILL SOME PLAYING PROBLEM EVEN IF QUE EMPTY
If VoiceStreamActive = True Then Exit Sub
If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then Exit Sub

'EMPTIED QUE DONT REQUIRE ANY MORE
If ListVOICE.ListCount = 0 Then Exit Sub

Call WaitWavFinish

'GRAB NEXT IN QUE
VoiceToSpeak = ListVOICE.List(0)
ListVOICE.RemoveItem (0)




'TRIGV = FILTER_SPEAK_LIST(2, VoiceToSpeak)

'If TRIGV = False Then Exit Sub



'HERE GO ABOUT TO START VOICE
If WinAmpChkIsPlaying = 1 Then
'    TTSAppMain.RateSldr.Value = 2
Else
'    TTSAppMain.RateSldr.Value = 2
End If

'Call TTSAppMain.RateSldr_Scroll


'-----------------------------------  CHUNK CODE

'MODIFY TO USE IN EVENT RELOCATED
'If WinAmpChkIsPlaying = 1 Then
'    WINAMPBackOnAfterVoiceRequired = True
'    Call ATidalX.MP3_OFF__ON_START_VOICE
'End If



On Error Resume Next

'VOICECOMMANDSTART = True
'VoiceStreamEnds = False

'-------------------------------------------------


If VoiceToSpeak = "" Then MsgBox "WARN - VOICE TO SPEAK NONE : ERROR"

Call PAUSE_WINAMP_FOR_VOICE

If InStr(VoiceToSpeak, "Dark") Then
    VBLINE = "If InStr(VoiceToSpeak, ""Dark"") Then"
    MsgBox "STOP IS HERE " + vbCrLf + VBLINE
    Stop
End If

TTSAppMain.Voice.Speak VoiceToSpeak, SVSFlagsAsync


'PUT HERE BUT NOT GETTING USED
Debug.Print "01 -- " + Format(Now, "HH MM SS  ") + "VOICE SPEAK RESULT " + VoiceToSpeak



'-------------------
On Error Resume Next
      
'Call FileInUseDelay(TIDALPATHVOICE, False)
FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\VOICE SPEAK LOGG.TXT"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
TIDALPATHVOICE = FILENAME_IN_USE_CHECK
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append Lock Write As #FR1
    Print #FR1, "# VOICE SPEAK " + Format(Now, "DD-MM-YYYY - HH:MM:SS") + " - " + VoiceToSpeak
Close #FR1

If Err.Number > 0 Then MsgBox "ERROR APPEND FILE " + vbCrLf + TIDALPATHVOICE + vbCrLf + Err.Description

On Error GoTo 0
'-------------------

Exit Sub


' ----------- LOGG THE VOICE SPOKE

If WinAmpChkIsPlaying = 1 Or 1 = 1 Then

        Err.Clear
ResumeHere1:
        If Err.Number > 0 Then Sleep 50
        On Error GoTo ResumeHere1

        If VUMPATH1 = "" Then
            If FindWindow(vbNullString, "VU METER LOGGER") > 0 Then
                Open "D:\temp\VU LOGGER PATHS.txt" For Input As #1
                 Line Input #1, VUMPATH1
                Line Input #1, VUMPATH2
            End If
        End If

        If VUMPATH1 <> "" Then
            If FindWindow(vbNullString, "VU METER LOGGER") > 0 Then
    
                Err.Clear
ResumeHere2:
                If Err.Number > 0 Then Sleep 50
                On Error GoTo ResumeHere2
    
                'Call FileInUseDelay(VUMPATH2, False)
                FILENAME_IN_USE_CHECK = VUMPATH2
                FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
                DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
                FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4

                FR1 = FreeFile
                Open FILENAME_IN_USE_CHECK For Append Lock Write As #FR1
                    Print #FR1, "# VOICE SPEAK " + Format(Now, "HH:MM:SS") + " - " + VoiceToSpeak
                Close #FR1

    
                Err.Clear
ResumeHere4:
                If Err.Number > 0 Then Sleep 50
                On Error GoTo ResumeHere4

                'Call FileInUseDelay(TIDALPATHVOICE, False)
                FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\VOICE SPEAK LOGG_2.TXT"
                FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
                DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
                FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
                TIDALPATHVOICE = FILENAME_IN_USE_CHECK
                FR1 = FreeFile
                Open FILENAME_IN_USE_CHECK For Append Lock Write As #FR1
                    Print #FR1, "# VOICE SPEAK " + Format(Now, "HH:MM:SS") + " - " + VoiceToSpeak
                Close #FR1

            End If '  If FindWindow(vbNullString, "VU METER LOGGER") > 0 Then
        End If ' VUMPATH AVIAL
End If '  If WINAMP YES Then

Err.Clear
ResumeHere3:
    
If Err.Number > 0 Then Sleep 50
On Error GoTo ResumeHere3
      
'Call FileInUseDelay(TIDALPATHVOICE, False)

FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\VOICE SPEAK LOGG.TXT"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
TIDALPATHVOICE = FILENAME_IN_USE_CHECK
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append Lock Write As #FR1
    Print #FR1, "# VOICE SPEAK " + Format(Now, "DD-MM-YYYY - HH:MM:SS") + " - " + VoiceToSpeak
Close #FR1

On Error GoTo 0

'If ListVOICE.ListCount > 0 Then Call VOICELISTTIMER_TIMER

End Sub


Sub PAUSE_WINAMP_FOR_VOICE()
        'THIS IS DEPENDENT HERE #
        
        '# '# Sub VOICELISTTIMER_TIMER()
        
        '#
        'FORM
        'This VB TTS App sample demonstrates most of the TTS functionalities
        'Private Sub ShowEvent(ParamArray strArray())
        '#

        MsgResult = ATidalX.IsWinAmpPlayingBeforeVoiceStartAndStop
        WINAMPBackOnAfterVoiceRequired = False
        If MsgResult = 1 Then
            
            Call ATidalX.MP3_OFF__ON_START_VOICE
            Call ATidalX.TimerDoubleVolume_Timer
            WINAMPBackOnAfterVoiceRequired = True
        
        End If

End Sub



Sub SpeakVoice(VoiceToSpeak2)

If ATidalX.MNU_SPEECH_OFF.Checked = True Then
    VOICELISTTimer.Enabled = False
    Exit Sub
End If

'THIS WILL QUE THEM UP TO VOICE
O_VoiceToSpeak2 = VoiceToSpeak2
Do
    VoiceToSpeak2 = Replace(VoiceToSpeak2, "  ", " ")
    If O_VoiceToSpeak2 = VoiceToSpeak2 Then Exit Do
    O_VoiceToSpeak2 = VoiceToSpeak2
Loop Until True = False

VoiceToSpeak2 = Replace(VoiceToSpeak2, " .", ".")
VoiceToSpeak2 = Trim(VoiceToSpeak2)

If TIMER_SECONDS20_AFTER_HOUR_TIMER.Enabled = True Then
    ' IF ENABLED AND THEN CALL AGAIN RESET ADDER DELAY
    Call TIMER_SECONDS20_AFTER_HOUR_TIMER_TRIG_START
End If

If TimerVOLCHANGETALK2.Enabled = True Then
    Call TimerVOLCHANGETALK2_TRIG_START
End If

If TimerVOLCHANGETALK.Enabled = True Then
    Call TimerVOLCHANGETALK_TRIG_START
End If

'NEVER A DUPE
If IsNumeric(VoiceToSpeak2) = False Then
    For r = 0 To ListVOICE.ListCount - 1
        If InStr(ListVOICE.List(r) + "*", VoiceToSpeak2 + "*") > 0 Then
            Exit Sub
        End If
    Next
    If FORCE_TALKER_GO_WITHOUT_CHECKER_DUPE_IGNOR = False Then
        If OVOICE = VoiceToSpeak2 Then Exit Sub
        OVOICE = VoiceToSpeak2
    End If
    FORCE_TALKER_GO_WITHOUT_CHECKER_DUPE_IGNOR = False
End If

'Debug.Print "SPOKE VOICE == " + VoiceToSpeak2

'THIS WILL QUE THEM UP TO VOICE

ListVOICE.AddItem VoiceToSpeak2

VOXLIST.List1.AddItem Time$ + " - " + VoiceToSpeak2

'400 SECOND ACTIVE
If VOXLIST_NOT_ACTIVE > 400 Then
    VOXLIST.List1.ListIndex = VOXLIST.List1.ListCount - 1
End If

If VOXLIST.Check1.Value = 1 Then
    VOXLIST.List1.ListIndex = VOXLIST.List1.ListCount - 1
End If


If VOXLIST.List1.ListCount > 3000 Then
  VOXLIST.List1.RemoveItem (1)
End If

'FIND HERE
VOICELISTTimer.Enabled = True

End Sub


Public Sub Label7_Click()
    'ATidalX.Label7
    'THIS USED TO CLICK DATE THAT HAPPEN YES ALWAYS SOON AS
    
    MNU_SPEECH_OFF.Checked = False
    Call SPEECH_OFF_SETTER

    FORCE_TALKER_GO_WITHOUT_CHECKER_DUPE_IGNOR = True
    'TALK DATE INSTANT
    Call MNU_VOX_LIST_Click

    VOXLIST_NOT_ACTIVE = 1000
    Beep
    Call TIMER_SECONDS20_AFTER_HOUR_TIMER_Timer
End Sub

Private Sub Label40_Click()

    'Label40.ToolTipText = "TALK TIME AND Dog Bark Commando"
'    Beep
'    DOGBARKCOMMO = True
'    Call FORCESAYTIME
    
'    OLD_RS232_LOGGER_FILE_TIME_SECOND = -1
'    OLD_RS232_LOGGER_FILE_TIME_MINUTE = -1
'    OLD_RS232_LOGGER_FILE_TIME_HOUR = -1
    START_DOOR_TALKER = True

    FORCE_TALKER_GO_WITHOUT_CHECKER_DUPE_IGNOR = True
    Call RSR232_FRONT_DOOR_LOGGER_TALK

End Sub

Public Sub LABEL_TIME_Click()
    'ATIDALX.LABEL_TIME_Click
    
    MNU_SPEECH_OFF.Checked = False
    Call SPEECH_OFF_SETTER

    Call MNU_VOX_LIST_Click
    
    VOXLIST_NOT_ACTIVE = 1000

    'Beep
    FORCE_TALKER_GO_WITHOUT_CHECKER_DUPE_IGNOR = True
    Call FORCESAYTIME

End Sub


Private Sub Menu_Click()
'Menu.Show
End Sub

Private Sub Menu2_Click()
Menu.Show

End Sub

Private Sub Mnu_ClipURL_Click()

'dw = frmIECtrl.LV1.ListItems.Count
'lv.ListItems.item(dw).Text

'    Print #freefile1, ListView1.ListItems.item(et) + ";";
'    Print #freefile1, ListView1.ListItems.item(et).SubItems(1) + ";";


'URL$ = frmIECtrl.LV1.ListItems.item(dw)
'URLLoc$ = frmIECtrl.LV1.ListItems.item(dw).SubItems(1)

'Clipboard.Clear
'Clipboard.SetText Text1.Text + vbCrLf


End Sub

Private Sub Mnu_CountD1_Click()

TTx$ = ""
Call Timer_CountD1Text_timer
Clipboard.Clear
Clipboard.SetText TTx$


End Sub

Sub FindTimeInfo()

'InPutDate = DateValue("01-12-2008")
'TestDate = DateValue("31-05-2009")

DayCountT = DateDiff("d", INPUTDATE, TestDate)

Year5 = 0
For r5 = Year(INPUTDATE) + 1 To Year(TestDate) + 1
    If DateSerial(r5, Month(INPUTDATE), Day(INPUTDATE)) < Now Then Year5 = Year5 + 1
Next
For r5 = 1 To -2 Step -1
    XX = DateSerial(Year(TestDate), Month(TestDate) + r5, Day(INPUTDATE))
    MonthStart = XX
    If XX <= TestDate Then Exit For
Next
Month5 = 0
Month7 = 0
For r5 = 1 To 10000
    XX = DateSerial(Year(INPUTDATE), Month(INPUTDATE) + r5, Day(INPUTDATE))
    If Year(XX) <> oxx Then Month7 = 0
    oxx = Year(XX)
    If XX <= TestDate Then Month5 = Month5 + 1: Month7 = Month7 + 1
    If XX >= TestDate Then Exit For
Next

For r5 = Year(TestDate) To 1 Step -1
    If DateSerial(r5, Month(INPUTDATE), Day(INPUTDATE)) < Now Then
        DayCount1 = DateDiff("d", DateSerial(r5, Month(INPUTDATE), Day(INPUTDATE)), TestDate)
        DayCountMonth = DateDiff("d", MonthStart, TestDate)
        WholeYear1 = DateDiff("d", DateSerial(r5, Month(INPUTDATE), Day(INPUTDATE)), DateSerial(r5 + 1, Month(INPUTDATE), Day(INPUTDATE)))
        'Month5 = DateDiff("m", DateSerial(r5, Month(InPutDate), Day(InPutDate)), TestDate)
        WeeksSinceYear = DateDiff("w", DateSerial(r5, Month(INPUTDATE), Day(INPUTDATE)), TestDate) - 1
        WeeksSinceInput = Int((DateDiff("D", INPUTDATE, TestDate) - 1) / 7)
        WeeksPlusDays = (DateDiff("d", INPUTDATE, TestDate) - 1) - (WeeksSinceInput * 7)
                        
                
        ResultYearDate = Format$(Year5 + DayCount1 / WholeYear1, ".000")
        Exit For
    End If
Next


DexFORMAT1 = 0
DexFORMAT2 = 0
DexFORMAT3 = 0

F1 = 0: F2 = 0: F3 = 0: F4 = 0

'DO SECONDS IF MASSIVE BUT ONLY TO THE LAST REMAINING DAY

'tMm = DateDiff("s", Int(TestDate), TestDate)
'F1 = Int((tMm / 60 ^ 2) / 24)
'If F1 > 0 Then tMm = tMm - (DateDiff("s", Now, Now + 1) * F1)
'F2 = Int((tMm / 60 ^ 2))
'F3 = Int(tMm / 60)
'F4 = tMm Mod 60


'DONT DO SECONDS IF MASSIVE
If Abs(DateDiff("D", INPUTDATE, TestDate)) > 366 Then Exit Sub

tMm = DateDiff("s", INPUTDATE, TestDate)
F1 = Int((tMm / 60 ^ 2) / 24)
If F1 > 0 Then tMm = tMm - (DateDiff("s", Now, Now + 1) * F1)
F2 = Int((tMm / 60 ^ 2))
F3 = Int(tMm / 60)
F4 = tMm Mod 60

'THIS MAY NOT BE CORRECT TEST
'If F1 < 0 Then F1 = 0

'If F1 < 0 Then Stop

'EXAMPLE
ex = Format$(TimeSerial(0, F3, F4), "H:MM:SS")
ex1 = Format$(TimeSerial(0, F3, F4), "hH:MM:SS")
ex2 = Trim(str(Val(Mid$(ex1, 1, 2)))) + "h" + str(Val(Mid$(ex1, 4, 2))) + "m" + str(Val(Mid$(ex1, 7, 2))) + "s"
ex3 = Trim(str(Val(Mid$(ex1, 1, 2)))) + "h" + str(Val(Mid$(ex1, 4, 2))) + "m"

DexFORMAT1 = Trim(str(F1)) + "d " + ex
DexFORMAT2 = Trim(str(F1)) + "d " + ex2
DexFORMAT3 = Trim(str(F1)) + "d " + ex3

If F1 = 1 Then
    DexFORMAT1 = Trim(str(F1)) + "d " + ex
    DexFORMAT2 = Trim(str(F1)) + "d " + ex2
    DexFORMAT3 = Trim(str(F1)) + "d " + ex3
End If

If F1 = 0 Then
    DexFORMAT1 = ex
    DexFORMAT2 = ex2
    DexFORMAT3 = ex3
End If


End Sub



Private Sub Mnu_FavProgs_Click()

TxZ = "MAster BATch VB6 Compiler"
Do While FindWindow(vbNullString, TxZ) > 0
Sleep 100
DoEvents
Loop


Shell "D:\VB6\VB-NT\00_Best_VB_01\Run Fav Progs\Run Fav Programs.exe", vbNormalNoFocus

End Sub

Private Sub MNU_HOT_KEY_LIST_Click()
    MsgBox "CONTROL M = MINI WINDOW ON CURSOR"
End Sub

Private Sub MNU_DONT_ADJUST_VOL_ON_VOICE_Click()
MNU_DONT_ADJUST_VOL_ON_VOICE.Checked = Not MNU_DONT_ADJUST_VOL_ON_VOICE.Checked
If MNU_DONT_ADJUST_VOL_ON_VOICE.Checked = False Then Exit Sub
MNU_MAXVOL_VOICE.Checked = False
Mnu_NoJumpUpVol.Checked = False
MNU_NO_MATCH_WINAMP_VOL.Checked = False
End Sub

Private Sub MNU_MAXVOL_VOICE_Click()
MNU_MAXVOL_VOICE.Checked = Not MNU_MAXVOL_VOICE.Checked
If MNU_MAXVOL_VOICE.Checked = False Then Exit Sub
Mnu_NoJumpUpVol.Checked = False
MNU_NO_MATCH_WINAMP_VOL.Checked = False
End Sub

Private Sub MNU_NO_MATCH_WINAMP_VOL_Click()
MNU_NO_MATCH_WINAMP_VOL.Checked = Not MNU_NO_MATCH_WINAMP_VOL.Checked
End Sub

Private Sub MNU_NO_SET_VOL_NONE_SET_AS_1_Click()
MNU_NO_SET_VOL_NONE_SET_AS_1.Checked = Not MNU_NO_SET_VOL_NONE_SET_AS_1.Checked

End Sub

Private Sub Mnu_NoJumpUpVol_Click()
Mnu_NoJumpUpVol.Checked = Not Mnu_NoJumpUpVol.Checked
End Sub

Private Sub Mnu_SetTime_Click()
'Time = Format(HOUR(Now) - 1, "00:") + Format(59, "00:") + Format(40, "00")
NOW2 = DateSerial(Year(Now), Month(Now), Day(Now)) + TimeSerial(HOUR(Now) - 1, 59, 38)
Date = Format(NOW2, "DD-MM-YYYY")
Time = Format(NOW2, "HH:MM:SS")
End Sub


Private Sub Mnu_SetTime3PM_Click()

Time = Format(14, "00:") + Format(59, "00:") + Format(52, "00")

End Sub

Private Sub Mnu_SetTimeHOURFORWARD_Click()

Time = Format(HOUR(Now), "00:") + Format(59, "00:") + Format(45, "00")

End Sub

Private Sub Mnu_SetTimesMIDNIGHT_Click()
'Date = "23-01-2010"
Time = Format(23, "00:") + Format(59, "00:") + Format(30, "00")
End Sub

Private Sub Mnu_SetTimesSUNNOON_Click()

If Tool3 = 0 Then Exit Sub
Time = Format(Tool3 - TimeSerial(0, 0, 30), "HH:MM:SS")

End Sub

Private Sub Mnu_SetTimesSUNSET_Click()
Tool2 = LASTSunSet
If Tool2 = 0 Then Exit Sub
Time = Format(Tool2 - TimeSerial(0, 0, 30), "HH:MM:SS")
Date = Format(Tool2 - TimeSerial(0, 0, 30), "DD-MM-YYYY")

End Sub

Private Sub Mnu_SimulateStringVoice_Click()

'---------------------
'Exit Sub
'---------------------

'If StringVoiceToPLayActive = True Then Exit Sub

If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then Exit Sub

'Call VoiceWait
'Call VoiceMP3Pause
Call SpeakVoice("THE ONE,, THE FORCE")

'DoEvents

SimuX = 9
'StringVoiceToPLayActive = True
TimerSimulateStringVoice.Enabled = True

'NOT WANTED

If STRING_OF_VOICE_IN_QUE_PLAY = 0 Then
    STRING_OF_VOICE_IN_QUE_PLAY = Now
End If

End Sub

Public Sub Mnu_TimeLogg_Click()
    
'Shell "NOTEPAD.EXE " + App.Path + "\" + GetComputerName + "-" + GetUserName + "\SAY Time Key Press Logg.txt", vbNormalFocus
FILE_NAME = App.Path + "\" + GetComputerName + "-" + GetUserName + "\SAY Time Key Press Logg.txt"

If IsIDE = True Then
EXT_PATH = ".VBP"
Else
EXT_PATH = ".EXE"
End If

If Dir(FILE_NAME) = "" Then
    MsgBox "FILE NOT EXIST" + vbCrLf + vbCrLf + FILE_NAME + vbCrLf + vbCrLf + "MESSENGER BOX FROM" + vbCrLf + vbCrLf + App.Path + App.EXEName + EXT_PATH
    Exit Sub
End If


Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing

End Sub


Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'If Slider1.Visible = False Then Exit Sub

If OVOLLAB <> Slider1.Value Then
    SliderC = Now + TimeSerial(0, 0, 4)
Else
    Exit Sub
End If
OVOLLAB = Slider1.Value
        
MsgResult = WinAmpChkIsPlaying

If MsgResult = 1 Then
    VolumeSet1 = Slider1.Value
    ATidalX.LblVol = str(VolumeSet1)
Else
    VolumeSet2 = Slider1.Value
    ATidalX.LblVol = str(VolumeSet2)
End If
        
                
'tf = SetVolume(SPEAKER, ATidalX.LblVol)

End Sub

Sub TIMER_SECONDS20_AFTER_HOUR_TIMER_TRIG_START()
    
    TIMER_SECONDS20_AFTER_HOUR_TIMER.Enabled = False
    TIMER_SECONDS20_AFTER_HOUR_TIMER.Interval = 0
    TIMER_SECONDS20_AFTER_HOUR_TIMER.Interval = 4000
    TIMER_SECONDS20_AFTER_HOUR_TIMER.Enabled = True

    TIMER_SECONDS20_AFTER_HOUR_TIMER_DOOR_TALKER.Enabled = False
    TIMER_SECONDS20_AFTER_HOUR_TIMER_DOOR_TALKER.Interval = 0
    TIMER_SECONDS20_AFTER_HOUR_TIMER_DOOR_TALKER.Interval = 8000
    TIMER_SECONDS20_AFTER_HOUR_TIMER_DOOR_TALKER.Enabled = True

End Sub


Public Sub TIMER_SECONDS20_AFTER_HOUR_TIMER_Timer()

TIMER_SECONDS20_AFTER_HOUR_TIMER.Enabled = False

FFT = ENGLISH_DATE_FOR_TALK(Now, True)

DayOrMidNightVol = 20
Call VolumeFixedSet(DayOrMidNightVol, True)

Call SpeakVoice(FFT)


End Sub

Private Sub Timer14_Timer()

Timer14.Enabled = False

Exit Sub

Timer14.Interval = 10000

If Form_Load_Done_For_Timer14 = False Then Exit Sub



Call SendKey(Asc("Z"), 4) '# send Alt

'Call SendKey(Asc("Z"), 0) '# send Alt
'sub sendkey(


''Call VoiceWait
'Call Time_Code

End Sub

Private Sub TimerSimulateStringVoice_Timer()

'Exit Sub

SimuX = SimuX - 1

If SimuX = 0 Then
    
    'Call VoiceMP3Pause
    
    'StringVoiceToPLayActive = False
    
    Call SpeakVoice("ZERO X")
    
    STRING_OF_VOICE_IN_QUE_PLAY = 0
   
   Exit Sub
   
End If

If SimuX > 0 Then
    
    'Call VoiceMP3Pause
        
    'StringVoiceToPLayActive = True
    
    If STRING_OF_VOICE_IN_QUE_PLAY = 0 Then
        STRING_OF_VOICE_IN_QUE_PLAY = Now
    End If
    
    Call SpeakVoice(str(SimuX))

    Exit Sub
    
End If


If SimuX = -1 Then TimerSimulateStringVoice.Enabled = False


End Sub

Sub TimerSunRiseSetPercent_Timer()

'SET A DELAY TO EVERY SECOND NOT WORTH SPEED CHECKER

TimerSunRiseSetPercent.Interval = 1000

'DISABLE THIS AN CALL MANULE FOR TEST ORDER
'TimerSunRiseSetPercent.Enabled = True
'TimerSunRiseSetPercent.Enabled = FALSE

'If StringVoiceToPLayActive = True Then Exit Sub

'If ONHOURCOUNTDOWN = True Then Exit Sub


'---------------------------
'WHY =
'HAPPENS AT BOOT UP
'OR MORE TEST AT CHANGE OVER
'PROBLEM TO FIX
If TooLTodaySunSet = 0 Then
    If TEST3 = True Then
        MsgBox "LOOK HERE STOP TooLTodaySunSet = 0"
    End If
    TEST1 = True
    Exit Sub
End If
'---------------------------

'--------------------------------------------------------

If Now > TooLTodaySunSet Then
    Call NewSunRise
    Call NewSunSet
    Call NewNoon
End If

'If Now >= TooLTodaySunRise And Now < TooLTodaySunSet Then
'    DARK_OR_LIGHT = True
'Else
'    DARK_OR_LIGHT = False
'End If

'If DARK_OR_LIGHT <> O_DARK_OR_LIGHT Then
'    Call NewSunRise
'    Call NewSunSet
'    Call NewNoon
'End If

'01 OF 03
If Now >= TooLTodaySunRise And Now < TooLTodaySunSet Then
    ' = DAY LIGHT MODE
    DARK_OR_LIGHT = True
    'DAY LIGHT
    PC1 = DateDiff("s", TooLTodaySunRise, TooLTodaySunSet)
    PC2 = DateDiff("s", TooLTodaySunRise, Now)
'    SunResult4 = (PC2 / PC1) * 100
    SunResult4 = Round((PC2 / PC1) * 100, 3)
    SunResult_Reverse = SunResult4

'    Label2.Caption = srt1$ + "Sun Rise = " + Mid$(Format$(Tool1, "HH:MM:SSampm"), 1, 8) + srt2$ + " -- " + Format(100 - SunResult4, "0.000") + "% " + Format(SunResult4, "0.000") + "%"
    Label2.Caption = srt1$ + "Sun Rise = " + Mid$(Format$(Tool1, "HH:MM:SSampm"), 1, 8) + srt2$ + "_ " + Format(SunResult4, "0.000") + "% " + Format(100 - SunResult4, "0.000") + "%"
    Call Percent_SUN_1_100(SunResult4, "LIGHT")

Else
    '= DARKNESS MODE
    DARK_OR_LIGHT = False

    'THE SUNSET BEFORE SUNRISE IS WHAT WE WANT
    If TooLTodaySunSet > TooLNewSunRise Then
        DARKNESS_PERCENT = TooLDayBeforeSunSet
        Else
        DARKNESS_PERCENT = TooLTodaySunSet
        'PROBLEM TEST IF EVER SET
    End If
    
    
    PC1 = DateDiff("s", DARKNESS_PERCENT, TooLNewSunRise)
    PC2 = DateDiff("s", DARKNESS_PERCENT, Now)
    
    'SunResult3 = (PC2 / PC1) * 100
    'SunResult3 = Round((PC2 / PC1) * 100, 3)
    SunResult3 = Val(Format((PC2 / PC1) * 100, "0.000"))
    SunResult_Reverse = SunResult3
    
    Label3.Caption = sst1$ + "Sun Set = " + Mid$(Format$(Tool2, "HH:MM:SS a/p"), 1, 8) + sst2$ + " _ " + Format(SunResult3, "0.000") + "% " + Format(100 - SunResult3, "0.000") + "%"

    Call Percent_SUN_1_100(SunResult3, "DARK")

End If
    
'02 OF 03
If Now >= TooLTodaySunRise And Now < TooLTodaySunSet Then
    'DAY LIGHT
    
    PC1 = DateDiff("s", TooLTodaySunRise, TooLTodaySunSet)
    PC2 = DateDiff("s", TooLTodaySunRise, Now)

'    SunResult4 = (PC2 / PC1) * 100
    SunResult4 = Round((PC2 / PC1) * 100, 3)
    SunResult_Reverse = SunResult4

    Label2.Caption = srt1$ + "Sun Rise = " + Mid$(Format$(Tool1, "HH:MM:SSampm"), 1, 8) + srt2$ + " _ " + Format(SunResult4, "0.000") + "% " + Format(100 - SunResult4, "0.000") + "%"
    Call Percent_SUN_1_100(SunResult4, "LIGHT")
    
    'icy2 = InStrRev(Label3, ":")
    'Label3.Caption = Mid$(Label3, 1, icy2 + 2)
    'icy2 = InStrRev(Label3, ":")
    'Label2.Caption = Mid$(Label2, 1, icy2 + 2)


    GoTo Ende_sub
    'Exit Sub

End If

'03 OF 03
If Now >= TooLTodaySunSet Then
    'DARKNESS
    
    DARK_OR_LIGHT = False
    
    PC1 = DateDiff("s", TooLTodaySunSet, TooLNewSunRise)
    PC2 = DateDiff("s", TooLTodaySunSet, Now)
    
    'SunResult3 = (PC2 / PC1) * 100
    SunResult3 = Round((PC2 / PC1) * 100, 3)
    SunResult_Reverse = SunResult3
    
    Label3.Caption = sst1$ + "Sun Set = " + Mid$(Format$(Tool2, "HH:MM:SSampm"), 1, 8) + sst2$ + " _ " + Format(SunResult3, "0.000") + "% " + Format(100 - SunResult3, "0.000") + "%"

    Call Percent_SUN_1_100(SunResult3, "DARK")

    'icy2 = InStrRev(Label3, ":")
    'Label3.Caption = Mid$(Label3, 1, icy2 + 2)
    'icy2 = InStrRev(Label3, ":")
    'Label2.Caption = Mid$(Label2, 1, icy2 + 2)

End If


Ende_sub:

Label2.Caption = Replace(Label2.Caption, " _ ", " ")
Label3.Caption = Replace(Label3.Caption, " _ ", " ")

O_DARK_OR_LIGHT = DARK_OR_LIGHT

End Sub

Sub SEASONS()


Dim E As Date, d As Date, ShallGo, Fi

'E = DateSerial(2008, 12, 31) + (Now - Int(Now))
E = Now
d = DateSerial(2009, 1, 1)
d2 = DateSerial(2009, 7, 1)

If E > d Then ShallGo = False: Fi = 1
If E >= d2 Then ShallGo = True: Fi = 2
'If E >= D Then ShallGo = False: Fi = 3

If Fi = 1 Then: d = DateSerial(Year(d) - 1, 1, 1)
If Fi = 2 Then: d = DateSerial(Year(d) + 1, 1, 1)
'If fi = 3 Then: D = DateSerial(Year(D) - 1, 1, 1)


d = DateSerial(2010, 1, 1)
E = Now


ShallGo = False

If ShallGo = True Then
    F1 = DateDiff("d", E, d) - 1
    F2 = DateDiff("h", E, d) - 1
    F3 = DateDiff("n", E, d) - 1
    F4 = DateDiff("s", E, d)
    f5 = Format$(SystemUptime.MSeconds, "000")
    f5 = 1000 - f5
Else
    F1 = DateDiff("d", d, E)
    F2 = DateDiff("h", d, E)
    F3 = DateDiff("n", d, E)
    F4 = DateDiff("s", d, E)
    f5 = Format$(SystemUptime.MSeconds, "000")
End If

f6 = F2 Mod 24
f7 = F3 Mod 60
f8 = F4 Mod 60
rt = Now + F1
rt = rt + TimeSerial(f6, f7, f8)

TT$ = ""

TT$ = TT$ + Format$(Now, "ddd dd-mmm-yyyy hh:mm:ssa/p UK") + vbCrLf + vbCrLf

INPUTDATE = d
TestDate = Now
Call FindTimeInfo
'TT$ = TT$ + "Tott Time In & Out Hoss -- " + ResultYearDate + " Years" + str(Month7) + "M" + str(DayCountMonth) + "D" + str(WeeksSinceInput) + " Weeks" + vbCrLf

Dim TT2

TT2 = " "
'TT2 = TT2 + Trim(str(F1)) + " Days "
TT2 = TT2 + Format$(f6, "0") + "h "
TT2 = TT2 + Format$(f7, "0") + "m "
If ShallGo = True Then
    TT2 = TT2 + Format$(f8, "0") + "s"
Else
    TT2 = TT2 + Format$(f8, "0") + "s"
End If


If ShallGo = True Then
    TT$ = TT$ + "Count Down of New Year" + str(Year(d)) + " Friday" + vbCrLf
    If WeeksPlusDays > 1 Then QQ = " Days" Else QQ = " Day"
    TT$ = TT$ + "T-" + str(F1) + " Days" + str(WeeksSinceInput) + " Weeks &" + str(WeeksPlusDays) + QQ + TT2 + vbCrLf
    TT$ = TT$ + "T-" + str(F2) + " Hours" + vbCrLf
    TT$ = TT$ + "T-" + str(F3) + " Mins" + vbCrLf
    TT$ = TT$ + "T-" + str(F4) + " Secs" + vbCrLf
    TT$ = TT$ + "T- ." + Format$(f5, "000") + " Milli Secs" + vbCrLf
Else
    TT$ = TT$ + "Counting Out of New Year" + str(Year(d)) + " Friday" + vbCrLf
    If WeeksPlusDays > 1 Then QQ = " Days" Else QQ = " Day"
    TT$ = TT$ + "T+" + str(F1) + " Days" + str(WeeksSinceInput) + " Weeks &" + str(WeeksPlusDays) + QQ + TT2 + vbCrLf
    TT$ = TT$ + "T+" + str(F2) + " Hours" + vbCrLf
    TT$ = TT$ + "T+" + str(F3) + " Mins" + vbCrLf
    TT$ = TT$ + "T+" + str(F4) + " Secs" + vbCrLf
    TT$ = TT$ + "T+ ." + Format$(f5, "000") + " Milli Secs" + vbCrLf
End If

'tt$ = tt$ + "T-Minus :- ." + Trim(Str(f5)) + " More Milli Seconds Accordingly Here Via Compu GetTickCount" + vbCrLf

'MsgBox tt$


'TT$ = TT$ + vbCrLf

'On Error GoTo 0
'Call Lab10Sub
On Error Resume Next

'Count in new year
TT$ = TT$ + vbCrLf

'MsgBox tt$
'moon stuff
TT$ = TT$ + Wd$ + vbCrLf


'TT$ = TT$ + "GPS = Hove -- Center Of My Ceiling - Height Dont Count" + vbCrLf
'TT$ = TT$ + "GPS = Hove - My Room" + vbCrLf
TT$ = TT$ + "GPS  Hove" + vbCrLf
'sunrise sunset
If Now < Tool1 Or Now > Tool2 Then
    
'    PC1 = DateDiff("s", Tool4, Tool1)
'    PC2 = DateDiff("s", Tool4, Now)
'    SunResult3 = (PC2 / PC1) * 100

    TT$ = TT$ + "Sun Rise " + Mid$(Label2, 12) + vbCrLf
    TT$ = TT$ + "Sun Set " + Mid$(Label3, 11, 8) + " ¦ " + Format(100 - SunResult3, "0.00") + "% <> " + Format(SunResult3, "0.000") + "% Da-Da-Dra Darkness" + vbCrLf

Else
    
'    PC1 = DateDiff("s", Tool1, Tool2)
'    PC2 = DateDiff("s", Tool1, Now)
'    SunResult4 = (PC2 / PC1) * 100

    TT$ = TT$ + "Sun Rise " + Mid$(Label2, 12, 8) + " ¦ " + Format(100 - SunResult4, "0.00") + "% <> " + Format(SunResult4, "0.000") + "% Day Light" + vbCrLf
    TT$ = TT$ + "Sun Set " + Mid$(Label3, 11, 8) + vbCrLf


End If



'Solar Noon
TT$ = TT$ + "Solar Noon " + Mid$(Label41, 14) + vbCrLf

Dim srs As New clsSunRiseSet
srs.City = "London, England"

srs.DateDay = Now + 1
srs.DaySavings = DaySavingsSub(srs.DateDay)
srs.CalculateSun
tool5 = srs.SolarNoon
TT$ = TT$ + "Solar Noon " + Mid$(Format$(tool5, "HH:MM:SSampm"), 1, 8) + " 2Moro" + vbCrLf


TT$ = TT$ + vbCrLf

'TT$ = TT$ + "Difference of Light of Day Before" + vbCrLf


TT$ = TT$ + "Length of Day  " + Mid$(TooL46$, 1, 2) + "h " + Mid$(TooL46$, 4, 2) + "m" + Mid$(TooL46$, 9) + vbCrLf


srs.DateDay = Int(Now)
srs.CalculateSun
PC1 = DateDiff("s", srs.Sunrise, srs.Sunset)

srs.DateDay = Int(Now) + 1
srs.CalculateSun
PC2 = Abs(PC1 - DateDiff("s", srs.Sunrise, srs.Sunset))



TT$ = TT$ + "Difference of Light" + str(PC2) + " Secs" + vbCrLf
TT$ = TT$ + "Today " + DOL1 + vbCrLf
TT$ = TT$ + "2Moro " + DOL2 + vbCrLf

tj = InStrRev(Label13, "-- ")
TT$ = TT$ + vbCrLf
TT$ = TT$ + "Equinox / Solstice " + Trim(Mid$(Label13, tj + 3)) + vbCrLf

'Dim iX2

'iX2 = Val(Replace(Mid$(Label13, tj + 3), "%", ""))
'Call Percent_EQI_1_100(iX2, "")


If 1 = 2 Then
    Open "D:\Temp\Cross Quarter.txt" For Input As 1
        Line Input #1, HALO3
        Line Input #1, HALO5
        HALO3 = Mid$(HALO3, 1, Len(HALO3) - 7)
        HALO5 = Mid$(HALO5, 1, Len(HALO5) - 7)
    
    Close #1
    Open "D:\Temp\Equinox.txt" For Input As 1
        Line Input #1, HALO7
        Line Input #1, HALO4
    Close #1

End If


'TEST IS OKAY
'--------------------------------------------
If HALO3 = "" And 1 = 2 Then
'--------------------------------------------
    'SEASONS
    XX5 = InStr(HALO3, " -")
    If XX5 > 0 Then HALO3 = Mid$(HALO3, 1, XX5) + Mid$(HALO3, XX5 + 3)
    EQIHALO = DateValue(Mid$(HALO3, 1, XX5))
    LASTSAYSEASON = Mid$(HALO3, XX5)
    LASTSAYSEASON = Trim(Mid$(LASTSAYSEASON, 1, InStr(LASTSAYSEASON, "-") - 1))
    
    
    XX5 = InStr(HALO5, " -")
    If XX5 > 0 Then HALO5 = Mid$(HALO5, 1, XX5) + Mid$(HALO5, XX5 + 3)
    EQI1STHALO = DateValue(Mid$(HALO3, 1, XX5))
    EQIHALO = DateValue(Mid$(HALO5, 1, XX5))
    SAYSEASON = Mid$(HALO5, XX5)
    SAYSEASON = Trim(Mid$(SAYSEASON, 1, InStr(SAYSEASON, "-") - 1))
    EQI2NDHALO = EQIHALO
    EQIPCMAX = DateDiff("s", EQI1STHALO, EQI2NDHALO)
    EQIPC1 = DateDiff("s", EQI1STHALO, Now)
    EQIPC2 = Round(100 - ((EQIPC1 / (EQIPCMAX)) * 100), 4)
           
        
    'EQINOX'S
    XX5 = InStr(HALO7, " -")
    If XX5 > 0 Then HALO7 = Mid$(HALO7, 1, XX5) + Mid$(HALO7, XX5 + 3)
    XX5 = InStr(HALO4, " -")
    If XX5 > 0 Then HALO4 = Mid$(HALO4, 1, XX5) + Mid$(HALO4, XX5 + 3)
    EQIHALOTEST = DateValue(Mid$(HALO4, 1, XX5))
    If EQIHALOTEST < EQIHALO Then
    
        EQIHALO = DateValue(Mid$(HALO7, 1, XX5))
        LASTSAYSEASON = Mid$(HALO7, XX5)
        LASTSAYSEASON = Trim(Mid$(LASTSAYSEASON, 1, InStr(LASTSAYSEASON, "-") - 1))
        
        EQI1STHALO = DateValue(Mid$(HALO7, 1, XX5))
        EQIHALO = DateValue(Mid$(HALO4, 1, XX5))
        SAYSEASON = Mid$(HALO4, XX5)
        SAYSEASON = Trim(Mid$(SAYSEASON, 1, InStr(SAYSEASON, "-") - 1))
        EQI2NDHALO = EQIHALO
        EQIPCMAX = DateDiff("s", EQI1STHALO, EQI2NDHALO)
        EQIPC1 = DateDiff("s", EQI1STHALO, Now)
        EQIPC2 = Round(100 - ((EQIPC1 / (EQIPCMAX)) * 100), 4)
    
    
    End If


End If

If HALO3 <> "" Then
    TT$ = TT$ + HALO3 + " (+93 Days Equi/Sol)" + vbCrLf
    TT$ = TT$ + HALO5 + " (+93)" + vbCrLf '" (+93 Days Equi/Sol)" + vbCrLf
    
    Label30.Caption = SAYSEASON + Format$(EQIHALO, " DDD DD/MMM/YYYY") + str(EQIPC2) + "%"
    
    EQUIDATE1 = DateValue(Mid(HALO3, 1, 10))
    EQUIDATE2 = DateValue(Mid(HALO5, 1, 10))
    
    Call Percent_EQI_1_100(EQUIDATE1, EQUIDATE2, EQIPC2, SAYSEASON, LASTSAYSEASON)
End If

'Label30.Caption = "SEASON AND EQINOX % NOT USE YET" 'SAYSEASON + Format$(EQIHALO, " DDD DD/MMM/YYYY") + str(EQIPC2) + "%"
'Call SunRise_SunSet
'Label30.Caption = "SEASON AND EQINOX  " + Format$(EQUINOX_PERCENT, "###0.00") + "%"


TT$ = TT$ + vbCrLf

TT$ = TT$ + HALO7 + vbCrLf
TT$ = TT$ + HALO4 + vbCrLf

TT$ = TT$ + vbCrLf

If 1 = 2 Then

    Open "D:\Temp\Eclipse.txt" For Input As 1
        Line Input #1, HALO7
        TT$ = TT$ + HALO7 + vbCrLf
        Line Input #1, HALO7
        TT$ = TT$ + HALO7 + vbCrLf
    Close #1
    
    TT$ = TT$ + vbCrLf

End If


GoTo JUMPMOON

GG = 0
Open "D:\Temp\BlueMoon.txt" For Input As 1
    Do
        Line Input #1, HALO7
        
        XX5 = InStr(HALO7, "Moon -") + 5
        
        If InStr(HALO7, "Almanac Blue") > 0 Then
            xx7 = InStr(HALO7, " ")
            HALO7 = Mid$(HALO7, 1, xx7 - 1) + " " + Trim(Mid$(HALO7, xx7 + 3)) + " 4 In Season"
        End If
        
        If InStr(HALO7, "Calendar") > 0 Then
            HALO7 = Mid$(HALO7, 1, XX5) + " 2 in Month"
            
        
        Else
            
            'yes does anything run this one - dont think do
            '### botch code thinks
            If InStr(HALO7, "Harvest") = 0 And InStr(HALO7, "New Moon") = 0 And InStr(HALO7, "Full Moon") = 0 And InStr(HALO7, "Almanac Blue") = 0 Then
                HALO7 = Mid$(HALO7, 1, XX5) + " " + Trim(Mid$(HALO7, InStr(XX5, HALO7, " "), 3)) + " " + Trim(Mid$(HALO7, InStr(HALO7, "Full Moons") + 10))
            End If
            
            If InStr(HALO7, "in Month") > 0 Or InStr(HALO7, "in Season") > 0 Or InStr(HALO7, "in Year") > 0 Then
                HALO7 = Mid$(HALO7, 1, XX5) + " " + Trim(Mid$(HALO7, InStr(XX5, HALO7, " "), 3)) + " " + Trim(Mid$(HALO7, InStr(HALO7, "Full Moons") + 10))
            End If
        
        End If
        
        Do
            XX5 = InStr(HALO7, " -")
            If XX5 > 0 Then
                HALO7 = Mid$(HALO7, 1, XX5) + Mid$(HALO7, XX5 + 3)
                If InStr(HALO7, "Full Moon") = 11 Or InStr(HALO7, "New Moon") = 11 Then GG = GG + 1
                If GG = 2 Then
                    
                    GG = 3
                    
                    HALO7 = HALO7 + vbCrLf + Format(X9, "dd-mmm-yy HH:MMa/p ") + Format(MoonH + 1, "0") + "-of-8 " + Mid$(Label11.Caption, 6)
                        If Mid$(Label11.Caption, 6) <> "New Moon" Or Mid$(Label11.Caption, 6) <> "Full Moon" Then
                        XX1 = NextNewMoon(Now)
                        XX2 = NextFullMoon(Now)
                        If XX2 < XX1 Then moontext = "Full Moon": XX1 = XX2 Else moontext = "New Moon"
                    
                        HALO7 = HALO7 + vbCrLf + Format(XX1, "dd-mmm-yy ") + moontext
                    End If
                
                
                End If
            
            End If
        Loop Until XX5 = 0
        
        
        If Year(DateValue(Mid(HALO7, 1, 9))) <= Year(Now) + 1 Then
            TT$ = TT$ + HALO7 + vbCrLf
        End If
    Loop Until EOF(1)
Close #1



JUMPMOON:



End Sub


Sub Timer_CountD1Text_timer()

'    Timer_CountD1Text.Enabled = False

If Day(Now) <> DAY_NOW_OLD Then
    DAY_NOW_OLD = Day(Now)
Else
    Exit Sub
End If


Call SEASONS


'-------------------------------------------------
Exit Sub
'-------------------------------------------------

If IsIDE = True Then TEARGAS = 5 Else TEARGAS = 20

If (Second(Now) Mod TEARGAS <> 0 And TTx$ <> "" And OldNow2 <> Now) Or Label27.Caption = "Week No" Then
    Exit Sub
End If

OldNow2 = Now


If TickGo = True Then TickGo = False: Exit Sub


'On Local Error GoTo 0
On Error Resume Next

Tx2$ = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\System UpTime Logg Length.txt"
Tx3$ = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\System UpTime Logg Length.bak.txt"
FS.CopyFile Tx2$, Tx3$
freef5 = FreeFile

Open Tx2$ For Output Lock Write As #freef5
' "Sys Reboot Length "
Print #freef5, Format$(Now, "; DD-MM-YYYY; HH:MM:SS")
Close freef5

Tx2$ = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\System UpTime Logg Descrip.txt"
Tx3$ = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\System UpTime Logg Descrip.bak.txt"
FS.CopyFile Tx2$, Tx3$

freef5 = FreeFile
Open Tx2$ For Output Lock Write As #freef5
Print #freef5, "Sys Reboot Time ---- " + Format$(SysStartTimeOldN, "DD-MM-YYYY HH:MM:SS ")
Print #freef5, "Sys Reboot Length - " + Format$(Now, "DD-MM-YYYY HH:MM:SS")
Print #freef5, Label31.Caption
Print #freef5, "Tick Count - " + str(GetTickCount)
Close #freef5

'On Local Error GoTo ErrorExit



Call SEASONS



JUMPMOON:

'tt$ = tt$ + "Days in Month" + vbCrLf
TT$ = TT$ + vbCrLf

TT$ = TT$ + Label28 + vbCrLf
If Day(DateSerial(Year(Now), 2, 29)) = 29 Then pk = "Yes" Else pk = "No"
If pk = "Yes" Then
    RR8$ = "In a Leap Year " + pk
TT$ = TT$ + RR8$ + vbCrLf
End If
For r = Year(Now) + 1 To Year(Now) + 10
If Day(DateSerial(r, 2, 29)) = 29 Then Exit For
Next
RR8$ = "Leap Year" + str(r)

rr9$ = "Month Start #" + Format(DateSerial(Year(Now), Month(Now), 1), "DDD") + " - End #" + Format(DateSerial(Year(Now), Month(Now) + 1, 0), "DDD")
TT$ = TT$ + RR8$ + vbCrLf

TT$ = TT$ + vbCrLf

TT$ = TT$ + "Week# Starts" + vbCrLf
RR = InStr(5, Label27, "#")
tj = InStr(Label27, "-- Nx")
TT$ = TT$ + Mid$(Label27, 1, tj - 2) + vbCrLf
TT$ = TT$ + rr9$ + vbCrLf


TT$ = TT$ + vbCrLf
TT$ = TT$ + "SysUpTime  " + Format$(SysStartTimeNewN, "DDD DD-MMM-YY HH:MM:SSa/p") + vbCrLf
TT$ = TT$ + "This  " + Label31 + vbCrLf
TT$ = TT$ + "Last  " + SysOldUpTime + vbCrLf
'TT$ = TT$ + "Last - " + LastSysUptime$ + vbCrLf
TT$ = TT$ + vbCrLf

'tt$ = tt$ + "Last Ward Round" + vbCrLf
'RR = InStr(5, Label27, "#")
'tt$ = tt$ + Mid$(Label27, tj + 3) + "-08" + vbCrLf

'Dim InPutDate, TestDate, Year5, MonthStart, DayCount1, DayCountMonth, WholeYear1
'Dim Month5, WeeksSinceYear,WeeksSinceInput, ResultYearDate


INPUTDATE = DateValue("06-01-1997")
TestDate = Now
Call FindTimeInfo

DAYHOSSDIFF = DaysHossInpaitentVar - DaysHossTotalVar

TT$ = TT$ + "ToTT In & Out Hosso" + str(DayCountT) + " Days " + ResultYearDate + " Years" + str(Month7) + "M" + str(DayCountMonth) + "D" + str(WeeksSinceInput) + " Wks - Active" + vbCrLf
'TT$ = TT$ + "ToTT In & Out Hoss -" + str(DayCountT) + " Days " + ResultYearDate + " Years" + str(Month7) + "M" + str(DayCountMonth) + "D" + str(WeeksSinceInput) + " Wks - Active" + vbCrLf


oResultYearDate = ResultYearDate
XDayCountT = DayCountT
INPUTDATE = DateValue("01-01-2000")
TestDate = INPUTDATE + DaysHossTotalVar
Call FindTimeInfo

TT$ = TT$ + "ToTT Inside Hosso" + str(DaysHossTotalVar) + " Days " + ResultYearDate + " Years" + str(Month7) + "M" + str(WeeksSinceInput) + "Wks" + vbCrLf

'outhoss = Format(((XDayCountT - DaysHossTotalVar) / (XDayCountT)) * 100, "0.00")




'TT$ = TT$ + Format((DaysHossTotalVar / (XDayCountT - DaysHossTotalVar)) * 100, "0.00") + "% *IN* -- *OUT*" + str(Val(oResultYearDate) - Val(ResultYearDate)) + " Years" + vbCrLf
TT$ = TT$ + Format((DaysHossTotalVar / (XDayCountT - DaysHossTotalVar)) * 100, "0.00") + "% *IN*" + vbCrLf

TT$ = TT$ + "£1K a Week Room = £" + Format(((WeeksSinceInput) * 1000) / 1000, "###,##0") + "K" + vbCrLf


TT$ = TT$ + "As InPaitent +" + str(DAYHOSSDIFF) + " Days " + vbCrLf

'#SOLICITOR PUT PERCENT



'TT$ = TT$ + "ToTT As InPaitent +" + str(DaysHossInpaitentVar) + " Days " + vbCrLf

'TT$ = TT$ + "ToTT As InPaitent +" + str(DaysHossInpaitentVar) + " - " + ResultYearDate + " Years" + str(Month7) + "M" + str(WeeksSinceInput) + "Wks" + vbCrLf
'TT$ = TT$ + "As InPaitent +" + str(DAYHOSSDIFF) + " Days" + vbCrLf


INPUTDATE = DateValue("01-06-2007")
TestDate = DateValue("01-05-2009")
Call FindTimeInfo

TT$ = TT$ + "Last InPaitent " + str(DayCountT) + " Days "
TT$ = TT$ + ResultYearDate + " Years" + str(Month7) + "M" + str(DayCountMonth) + "D" + str(WeeksSinceInput) + " Wks" + vbCrLf


vvx2$ = str(DateDiff("d", DateValue("30-11-2008"), DateValue("30-03-2009")))

TT$ = TT$ + "Time Last 7th Hosso Building 2008/9 " + vvx2$ + " Days  4 Months " + vbCrLf

xyt2 = DateDiff("m", DateValue("01-02-2008"), DateValue("01-12-2008"))
TT$ = TT$ + "Time At Hostel 2007/8 " + str(xyt2) + "M" + vbCrLf


INPUTDATE = DateValue("01-12-2008")
TestDate = DateValue("01-05-2009")
Call FindTimeInfo
TT$ = TT$ + "Time As 7th Hosso Paitent " + ResultYearDate + " Years" + str(Month7) + "M" + str(DayCountMonth) + "D" + str(WeeksSinceInput) + " Wks" + vbCrLf


INPUTDATE = DateValue("23-03-2009")
TestDate = Now
Call FindTimeInfo

TT$ = TT$ + "Total My Independance " + ResultYearDate + " Years" + str(Month7) + "M" + str(DayCountMonth) + "D" + str(WeeksSinceInput) + " Wks" + vbCrLf
TT$ = TT$ + vbCrLf

TT$ = TT$ + "9 Types of Drugs - Never With Concent" + vbCrLf
'TT$ = TT$ + "Which Results in Force using Vilolence and then Cohesion" + vbCrLf
'TT$ = TT$ + "Legal Drug Dealing - Using the Printed Precription Requisition = Rape" + vbCrLf

'TT$ = TT$ + "LMAORAONTFISPMU."
'Laughing my ass off rolling about on the floor in stiches peeing myself

TT$ = TT$ + "LMAORAOTFIPOPWMUIS." + vbCrLf
'Laughing my ass off rolling about on the floor in puddles of pee wettting myself in stitches
TT$ = TT$ + vbCrLf

PILLSY = PillsYes
tr = Int(Now)
Do
    tr = tr + 1
    PILLSY = PILLSY + 4
Loop Until PILLSY >= 10000


TT$ = TT$ + "Pills Since 2003 =" + str(PillsYes) + " Mg's of *R*" + "  ^_^  " + Format(tr, "DDD DD-MMM-YY") + " = 10K Pills" + vbCrLf
'TT$ = TT$ + "Pills Really Taken =" + str(PillsNo) + vbCrLf
TT$ = TT$ + "Pills Really Had Hitt Rate = " + Format((PillsNo / PillsYes) * 100, "0.0000") + "%" + vbCrLf ' and Includes Injections" + vbCrLf

tr = Int(Now)
Do
    tr = tr + 1
    PillsYes = PillsYes + PillAmountCurrentlyTaking
    ease = (PillsNo / PillsYes) * 100
Loop Until ease < 50

'TT$ = TT$ + "That Will Take Till -- " + Format(tr, "DDD DD-MMM-YYYY") + " -- To Get Below 50% if Stopped Taking These Entirly Today" + vbCrLf
TT$ = TT$ + "Be Until. " + Format(tr, "DDD DD-MMM-YYYY") + " To Be Below 50% if Stopped These Today" + vbCrLf
TT$ = TT$ + "Ha Ha." + vbCrLf

Dim Nurses



Open "D:\Temp\Excell Nurses.txt" For Input As #1
Line Input #1, XPDRM 'Pscyhi Male
Line Input #1, XPDRF 'Pscyhi feme
Line Input #1, XPDR 'Pscyhi
Line Input #1, XXXC 'Nursres
Line Input #1, DRF 'Doctors
Line Input #1, DRM 'Doctors
Line Input #1, DR 'Doctors
Line Input #1, XSW
Line Input #1, XAOT
Line Input #1, XCPN
Line Input #1, XOT
Line Input #1, XSTR
Line Input #1, PAITENT2
Close #1
Nurses = str(XXXC)
PAITENTS = str(PAITENT2)




TT$ = TT$ + vbCrLf
TT$ = TT$ + Trim(str(XSW)) + " SW's -" + str(XCPN) + " CPN's -" + str(XOT) + " OT's -" + str(XAOT) + " AOT's -" + str(XSTR) + " STR" + vbCrLf
TT$ = TT$ + "Paitents ------------ =" + str(PAITENTS) + vbCrLf
TT$ = TT$ + "Nurses an Staff -- =" + Nurses + vbCrLf
TT$ = TT$ + "Doctors   ----------- =" + str(DR) + " [ *New>5* ]" + str(DRM) + " M " + str(DRF) + " Feme" + vbCrLf
TT$ = TT$ + "Psychiatrists ------ =" + str(XPDR) + " [ *New+2* ]" + str(XPDRM) + " M " + str(XPDRF) + " Feme" + vbCrLf
TT$ = TT$ + "Total Doctors ----- =" + str(Val(DR) + Val(XPDR)) + " Next ¦ M/F " + Format(((Val(DRF) + Val(XPDRF)) / (Val(DRM) + Val(XPDRM))) * 100, "#00.0") + "%" + vbCrLf
TT$ = TT$ + vbCrLf
TT$ = TT$ + "4 Hosso Buildings" + vbCrLf
TT$ = TT$ + "9 Wards" + vbCrLf
TT$ = TT$ + "20 Ward Transfers" + vbCrLf
TT$ = TT$ + "10 Building Transfers" + vbCrLf
'TT$ = TT$ + vbCrLf
'TT$ = TT$ + "## Meetings -- Soon" + vbCrLf

FR1 = FreeFile
Open "D:\Temp\VBArea\Awesome.txt" For Input As #FR1
Awesome = Val(Input$(LOF(FR1), FR1))
Close #FR1

'TT$ = TT$ + vbCrLf + "Total Photo Art Image's Awesome # < " + Format(Awesome, "###,###,##0") + " > - No Dupe Count" + vbCrLf
'TT$ = TT$ + vbCrLf + "Total Photo Art Image's Awesome # < " + Format(Awesome, "###,###,##0") + " >" + vbCrLf
TT$ = TT$ + vbCrLf + "Total Art Image's Awesome " + Format(Awesome, "###,###,##0") + vbCrLf

FR1 = FreeFile
Open "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_02_URL_and_TEXT.txt" For Input As #FR1
Line Input #1, pp
Line Input #1, pp
Line Input #1, pp
Line Input #1, pp
Close #FR1

TT$ = TT$ + vbCrLf + "URL Fav Link's # " + Mid(pp, 1, InStr(pp, " ") - 1) + vbCrLf


'The end of filling this string
TTx$ = TT$



GoTo jump5:
TT$ = TT$ + "Roids ReBound - " + Format$(Now, "ddd dd mmmm yyyy hh:mm:ssa UK") + vbCrLf
TT$ = TT$ + "T-Minus - T-Plus - Take you Fancy" + vbCrLf
jump5:


'On Local Error Resume Next

FR1 = FreeFile
Open GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-MP3.txt" For Input As #FR1
ee1$ = Input$(LOF(FR1), FR1)
Close #FR1

Do
tight = InStr(ee1$, "  ")
If tight > 0 Then
    ee1$ = Mid(ee1, 1, tight - 1) + Mid(ee1, tight + 1)
End If
Loop Until tight = 0


FR1 = FreeFile
Open GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-KEYS.txt" For Input As #FR1
ee10$ = Input$(LOF(FR1), FR1)
Close #FR1



FR1 = FreeFile
Open GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-URL's.txt" For Input As #FR1
ee7$ = Input$(LOF(FR1), FR1)
Close #FR1

FR1 = FreeFile
Open "D:\VB6\VB-NT\00_Best_VB_01\Batch_Compiler_Auto\VBDataNoArchive\#BatchComplierLoggerCount2.txt" For Input As #FR1
EE9$ = Input$(LOF(FR1), FR1)
Close #FR1

FR1 = FreeFile
Open GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-DW of the Day.txt" For Input As #FR1
Line Input #FR1, yy
EE8$ = EE8$ + yy + vbCrLf
Do
Line Input #FR1, yy
EE8$ = EE8$ + "-" + yy + vbCrLf

Loop Until EOF(FR1)
Close #FR1


FR1 = FreeFile
Open GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-CPU.txt" For Input As #FR1
ee2$ = Input$(LOF(FR1), FR1)
Close #FR1

'Fr1 = FreeFile
'Open GetSpecialfolder(CSIDL_PERSONAL) +"\01 Email Settings\Signatures\Signature-Weather.txt" For Input As #Fr1
'ee3$ = Input$(LOF(Fr1), Fr1)
'Close #Fr1

FR1 = FreeFile
Open GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-On This Day.txt" For Input As #FR1
ee4$ = Input$(LOF(FR1), FR1)
'MsgBox ee4$
Close #FR1
ffg = InStr(ee4$, "On This Day")
If Len(ee4$) - ffg < 34 Then ee4$ = Mid$(ee4$, 1, ffg - 3) + vbCrLf
'MsgBox Mid$(ee4$, 1, ffg - 3) + vbCrLf
For r = Len(ee4$) - 1 To Len(ee4$) - 10 Step -2
    If Mid$(ee4$, r, 2) = vbCrLf Then ee4$ = Mid(ee4$, 1, r - 2)
    If Mid$(ee4$, r, 2) <> vbCrLf Then Exit For
Next

INPUTDATE = TTz2
TestDate = Now
Call FindTimeInfo
Call DrivesGB
'ttegg$ = "#1" + vbCrLf + "Laptop " + TTz4 + TTz5 + vbCrLf
ttegg$ = "#1" + vbCrLf + Mid(TTz4, 1, 4) + " " + Mid(TTz4, 12, 11) + " " + TTz5 + vbCrLf
ttegg$ = ttegg$ + "Sys Ins " + Format(TTz2, "DDD DD-MMM-YY HH:MMa/p") + "  " + ResultYearDate + " Yrs" + str(Month7) + "M" + str(DayCountMonth) + "D" + vbCrLf '+ str(WeeksSinceInput) + " Wks" + vbCrLf
'ttegg$ = ttegg$ + Mid$(TTz3, 11) + vbCrLf
ttegg$ = ttegg$ + TTz3 + vbCrLf

'ttegg$ = ttegg$ + "# 02" + vbCrLf + "Asus Eee PC1008HA Seashell Atom N280 1.66GHz (2 CPU)" + vbCrLf + "Sys Ins  27-Oct-09 12:01:48p" + vbCrLf + vbCrLf
'ttegg$ = ttegg$ + vbCrLf + "# 02" + vbCrLf + "Asus Eee PC1008HA Seashell Atom N280 1.66GHz (2 CPU)" + vbCrLf + "Sys Ins 27-Oct-09 12:01p" + vbCrLf + vbCrLf
'ttegg$ = ttegg$ + "#2" + vbCrLf + "Asus Eee PC1008HA Seashell Atom N280 1.66GHz (2 CPU)" + vbCrLf + vbCrLf
ttegg$ = ttegg$ + "#2" + vbCrLf + "Asus Eee PC1008HA Seashell Atom N280 1.66GHz" + vbCrLf + vbCrLf

'Drive Space
ttegg$ = ttegg$ + TTz6 + vbCrLf

'ttegg$ = ttegg$ + "Jobs - (1).Put Max MEM In (2).Replaced a Fan" + vbCrLf
'----------

frx = FreeFile
Open GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-Nortons2Small.txt" For Output Lock Read Write As #frx

FR1 = FreeFile
Open GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-CountDown.txt" For Output Lock Read Write As #FR1
Print #FR1, "^_^"
Print #frx, "^_^"

Print #FR1, TTx$ 'Counting out the New Year
Print #frx, TTx$ 'Counting out the New Year

If XVOL1 = "Music Vol" + str(0) + "%" + vbCrLf Then XVOL1 = ""
If VolumeSet1 > 1 Or XVOL1 = "" Then XVOL1 = "Music Vol" + str(VolumeSet1) + "%" + vbCrLf
XVOL2 = "Master Vol" + str(VolumeSet2) + "%"

Print #FR1, XVOL2
Print #frx, XVOL2
Print #FR1, XVOL1
Print #frx, XVOL1


'ee1$ = ee1$ + vbCrLf

Print #FR1, ee1$  'Media Now Playing Winamp
Print #frx, ee1$  'Media Now Playing

Print #FR1, ee10$  'KEYS MOUSE
Print #frx, ee10$  'KEYS MOUSE

Print #FR1, ee7$ 'Urls
Print #frx, ee7$ 'Urls

iio = InStr(EE9$, "Compiled ")
timein2 = DateValue(Mid(EE9, iio + 13, 21)) + TimeValue(Mid(EE9, iio + 13, 21))

INPUTDATE = timein2
TestDate = Now
Call FindTimeInfo


'dex = Trim(str(F1)) + " Days " +



If F4 > 0 Then PIO = Trim(str(F4)) + "sec" ' Ago"
If F3 > 0 Then PIO = Trim(str(F3)) + "min " + Trim(str(F4)) + "sec"
If F2 > 0 Then PIO = Trim(str(F2)) + "hr " + Trim(str(F3)) + "min"
If F2 > 1 Then PIO = Trim(str(F2)) + "hrs " + Trim(str(F3)) + "min"
If F1 > 0 Then PIO = "1 Day Ago + " + Trim(str(F2)) + "hrs" ' Ago"
If F1 > 1 Then PIO = Trim(str(F1)) + " Days " + Trim(str(F2)) + "hrs" ' Ago"

iio = InStr(EE9$, "Total --")


If tMm < 80 Then
    EE9 = Mid(EE9, 1, iio - 3) + "  " + DexFORMAT2 + vbCrLf + Mid(EE9, iio)
Else
    EE9 = Mid(EE9, 1, iio - 3) + "  " + DexFORMAT3 + vbCrLf + Mid(EE9, iio)
End If

EE9$ = Replace(EE9$, "  ", " ")

Print #FR1, EE9$; ' VB Projects
Print #frx, EE9$; ' VB Projects

TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\LastLink48HourCount.txt"
Fr8 = FreeFile
Open TS For Input As #Fr8
Line Input #Fr8, linker
Close Fr8

TS = "D:\VB6\VB-NT\00_Best_VB_01\Linker2\DATA\Link-Command-Line--OBJECTS-48Hr-COUNT.txt"
Fr8 = FreeFile
Open TS For Input As #Fr8
Line Input #Fr8, OBJECTSVB
Close Fr8



Print #FR1, "VB Complier Linker # 48Hrs *" + Trim(linker) + " Projects *" + Trim(OBJECTSVB) + " Objects" + vbCrLf
Print #frx, "VB Complier Linker # 48Hrs *" + Trim(linker) + " Projects" + Trim(OBJECTSVB) + " Objects" + vbCrLf



Print #FR1, ttegg$ ' Sys Install Time
Print #frx, ttegg$ ' Sys Install Time


EE8$ = Replace(EE8$, "  ", " ")
Print #frx, EE8$  'Double Words
Print #FR1, EE8$ 'Double Words



TS = "D:\VB6\VB-NT\00_Best_VB_01\Active Idle\00_Text_Data\Idle and Active Logger Idle Over Time.txt"
Set F = FS.GetFile(TS)
MdateMOD = 0
If F.DateLastModified <> MdateMOD Then
    MdateMOD = F.DateLastModified
    Fr5 = FreeFile
    Open TS For Input As #Fr5
    '8000 exact match only
    Seek Fr5, LOF(Fr5) - 20000
    stringgrab = UCase(Input(20000, #Fr5))
    Close #Fr5

    Dim DXTime(12)
    Dim DXooo(12)

    oj = Len(stringgrab)
    For r = 12 To 1 Step -1
        oj = InStrRev(stringgrab, "AWAKE", oj - 1)
        'MAKE SURE DONT WRITE TWO AWAKES ONE
        
        If oj > 0 Then DXTime(r) = oj
        If oj = 0 Then Exit For
    Next

    px = r + 1
    oj1 = Len(stringgrab)
    For r = px To 12
        If DXTime(r) > 0 Then
            XX1 = InStr(DXTime(r), stringgrab, vbCrLf)
            If XX1 = 0 Then XX1 = Len(stringgrab)
            oj1 = XX1
            XX2 = InStrRev(stringgrab, vbCrLf, DXTime(r))
            
            YY5 = Mid$(stringgrab, XX2 + 2, XX1 - XX2)
            DXooo(r) = YY5
            
            'If yy5 <> yy8 Then DXooo(R) = yy5
            'yy8 = yy5
                
        End If
    Next
    'If DXooo(12) = DXooo(11) Then DXooo(12) = ""
    
    
    
    oj3 = 0
    'QUICK NAP
    Do
        oj = InStr(oj + 1, stringgrab, "QUICK NAP")
        If oj > 0 Then oj3 = oj
    Loop Until oj = 0
    
    XX1 = InStr(oj3, stringgrab, vbCrLf)
    If XX1 = 0 Then XX1 = Len(stringgrab)
    
    XX2 = InStrRev(stringgrab, vbCrLf, oj3)
    oo3 = Mid$(stringgrab, XX2 + 2, XX1 + XX2)

    Dim OutSleepStat(20)
    OSs = 0

    For r = px To 12 - 1
            
            
            'ASLEEP'S
            If r < 12 Then
                OSs = OSs + 1
                INPUTDATE = DateValue(Mid(DXooo(r), 5, 17)) + TimeValue(Mid(DXooo(r), 5, 17))
                TestDate = DateValue(Mid(DXooo(r), 28, 17)) + TimeValue(Mid(DXooo(r), 28, 17))
                Call FindTimeInfo
                If F1 < 0 Then
                    VBLINE = "If F1 < 0 Then"
                    MsgBox "STOP IS HERE " + vbCrLf + VBLINE
                    Stop
                End If
                
                'THIS SHOW WHEN FINISH SLEEP ON TESTDATE UNLIKE ABOVE
                OutSleepStat(OSs) = Format(TestDate, "DDMMM") + " ASleep " + DexFORMAT3
            End If
            
            
            
            
            OSs = OSs + 1
            'AWAKE'S
            'If R = 9 Then Stop
            XI = 0
            'If R = 11 Then Stop
            'If InStr(DXooo(R), "05:38:58") > 0 Then Stop
            Do
            
            If XI = 0 Then INPUTDATE = DateValue(Mid(DXooo(r), 28, 17)) + TimeValue(Mid(DXooo(r), 28, 17))
                XI = 1
            
                If DXooo(r + 1) <> "" Then
                    TestDate = DateValue(Mid(DXooo(r + 1), 5, 17)) + TimeValue(Mid(DXooo(r + 1), 5, 17))
                Else
                    TestDate = Now: XInPutDate = INPUTDATE
                End If
                
                Call FindTimeInfo
                'TEMP OUT NEED REASON WHY PUT HERE JOB LATER
                'If F1 < 0 Then Stop
                
                If r = 11 Then
                    OutSleepStat(OSs) = Format(INPUTDATE, "DDMMM") + " Awake " + DexFORMAT2
                End If
            If OutSleepStat(OSs) = "" Then OutSleepStat(OSs) = Format(INPUTDATE, "DDMMM") + " Awake " + DexFORMAT3
            XI2 = 0
            If r < 12 Then
            If InStr(LCase(DXooo(r + 1)), "awake+") Then
                r = r + 1: XI2 = 1
            End If
            End If
            Loop Until XI2 = 0 Or r > 12
            
            
    Next
    
    
    
    
    
    QuickNap = ""
    If XInPutDate < DateValue(Mid(oo3, 5, 17)) + TimeValue(Mid(oo3, 5, 17)) Then
    'Quick Nap
    If oj3 > 0 Then
        INPUTDATE = DateValue(Mid(oo3, 5, 17)) + TimeValue(Mid(oo3, 5, 17))
        TestDate = DateValue(Mid(oo3, 28, 17)) + TimeValue(Mid(oo3, 28, 17))
        Call FindTimeInfo

        ex = Format$(TimeSerial(0, F3, F4), "H:MM:SS")
        ex1 = Format$(TimeSerial(0, F3, F4), "hH:MM:SS")
        ex2 = Trim(str(Val(Mid$(ex1, 1, 2)))) + "h" + str(Val(Mid$(ex1, 4, 2))) + "m" + str(Val(Mid$(ex1, 7, 2))) + "s"
        ex3 = Trim(str(Val(Mid$(ex1, 1, 2)))) + "h" + str(Val(Mid$(ex1, 4, 2))) + "m"
    
        If Day(INPUTDATE) = 1 Then tsf = "st"
        If Day(INPUTDATE) = 0 Then tsf = "nd"
        If Day(INPUTDATE) = 0 Then tsf = "rd"
        If Day(INPUTDATE) > 3 Then tsf = "th"
        If Day(INPUTDATE) = 21 Then tsf = "st"
        If Day(INPUTDATE) = 22 Then tsf = "nd"
        If Day(INPUTDATE) = 31 Then tsf = "st"
        QuickNap = "Quick Quack Nap @ " + Format(INPUTDATE, "Ham/pm") + " " + Format(INPUTDATE, "DDD D") + tsf + " For " + ex3
        QuickNap = vbCrLf + "& A " + QuickNap
    End If
    End If
    
End If
'MdateMOD = 0

KILLY = ""
For r = 20 To 1 Step -1
    
    If OutSleepStat(r) <> "" Then
        KILLY = OutSleepStat(r) + vbCrLf + KILLY
        TTH = TTH + 1: If TTH = 5 Then Exit For
    End If
Next

KILLY = KILLY + QuickNap


'SLEEP OUTPUT
'msgbox KILLY

'Print #Fr1, KILLY
'Print #frx, KILLY





TS = "D:\VB6\VB-NT\00_Best_VB_04\RS232 LOGGER\RS232 DOOR REED DATA\LAST DOOR CLOSE.txt"
Fr8 = FreeFile
Open TS For Input As #Fr8
Do
    Line Input #Fr8, pio2
Loop Until EOF(Fr8)
Close Fr8

INPUTDATE = DateValue(pio2) + TimeValue(pio2)
TestDate = Now
Call FindTimeInfo

PIO = "Been InDoors " + DexFORMAT3 + vbCrLf

Print #FR1, PIO
Print #frx, PIO

iit = QEmptyRecycleBinSize
Print #FR1, iit
Print #frx, iit




'freemem = "Psychical Hard Mem Use " + frmMemInfo.txtMemoryUsagePsy.Text
freemem = "Ram Use " + frmMemInfo.txtMemoryUsagePsy.Text

Print #FR1, freemem 'Free Mem%
Print #frx, freemem 'Free Mem%


ee2$ = Mid(ee2$, 11)

Print #FR1, "CPU Use " + ee2$;
Print #frx, "CPU Use " + ee2$;







Print #FR1, "SigLine ";
Print #frx, "SigLine ";

'Print #Fr1, "Ya You'd EFN Like To Know Me..."
'Print #Fr1, ee3$ 'Weather
'Print #Fr1, ee4$; 'Last Next

'fr1'frx
Close #FR1, #frx



A1$ = GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-CountDown.txt"

frx = FreeFile
Open A1$ For Input Lock Read Write As #frx
YYO = Input(LOF(frx), frx)
Close #1
'LEN(YYO)'3056
YYO = Replace(YYO, vbCrLf + vbCrLf + vbCrLf, vbCrLf + vbCrLf)
'LEN(YYO)'3048

frx = FreeFile
Open A1$ For Output Lock Read Write As #frx
Print #frx, YYO;
Close #1




A1$ = GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-CountDown.txt"
Set F = FS.GetFile(A1$)
fSize = F.Size + 6


FILENAME_IN_USE_CHECK = GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-Nortons2Small.txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4

frx = FreeFile
Open FILENAME_IN_USE_CHECK For Append Lock Write As #frx

FILENAME_IN_USE_CHECK = GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-CountDown.txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4

FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append Lock Write As #FR1

Print #FR1, Format(fSize / 1024, "0.000") + "K";
Print #frx, Format(fSize / 1024, "0.000") + "K";
Close #FR1, #frx

On Error Resume Next

A1$ = GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-CountDown.txt"
a3$ = GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-Nortons2.txt"
Kill a3$
FS.MoveFile A1$, a3$

A1$ = GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-Nortons2.txt"
A2$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\My Email Signature Line.txt.tmp"
a3$ = "D:\MY WEBS\MatthewLan.Com Web\LoveFolder\Cool\My Email Signature Line.txt"
FS.CopyFile A1$, A2$
Kill a3$
Name A2$ As a3$


Mnu_SigLine.Caption = "NotePad SigLine " + Format(fSize / 1024, "0.000") + "K"


'll$ = "This System Check's with DR *Symantec's Norton 360"

'Main sig line docu is this
'Signature-Nortons2.txt


CountD1 = False
Exit Sub
ErrorExit:
Exit Sub

End Sub


Private Sub Mnu_EditMallSound_Click()
' Shell """C:\Program Files\Cool2000\cool2000.exe"" ""D:\VB6\VB-NT\00_BES~1\TIDAL_~1\WAVESO~1\DOBFIG~4.WAV""", vbNormalFocus
End Sub

Private Sub Mnu_Explore_Click()
Shell "Explorer.exe /e, " + App.Path, vbNormalFocus

End Sub

Private Sub Mnu_KeyHook_Click()

'If IsIDE = True Then TestKeyLoggOff = -2 'dont even run in IDE

TestKeyLoggOff = False
TestHookinVB = True

If DSkeybd_F.IsHook = False Then

    ASYNC_Key_Press_Timer.Enabled = False
    Load DSkeybd_F
    Call DSkeybd_F.Command1_Click

End If

End Sub

'Private Sub Mnu_KeyLoggDay_Click()
''Shell "NotePad " + App.Path + "\Logg\" + GetComputerName + "-" + GetUserName + "\Key Logger Text " + Format$(KeyLoggDate, "YYYY-mm-dd HH-MM-SS") + ".txt", vbNormalFocus
'FILE_NAME = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Key Logger Text " + Format$(KeyLoggDate, "YYYY-mm-dd HH-MM-SS") + ".txt"
'WSHShell.Run """" + FILE_NAME + """"
'Set WSHShell = Nothing
'End Sub


Private Sub Mnu_KeyLoggFolder_Click()
    Shell "Explorer.exe /e, " + App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\", vbNormalFocus
End Sub

Private Sub Mnu_KeyLogText_Click()

'Shell "NotePad " + App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\+Key Logger Text.txt", vbNormalFocus
FILE_NAME = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Key Logger Text.txt"

If Dir(FILE_NAME) <> "" Then
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing

Else

MsgBox "FILE NOT EXIST" + vbCrLf + FILE_NAME

End If
End Sub

Private Sub Mnu_ELock_Click()

LockDontAllowHotKeyforPassword = True
XxG = Now
'Call Bollocks
Call PassModule.Main

End Sub

Public Sub Mnu_ELock_Off_Click()
        
    Call SetLockTime
    Call SetLockMax
    Unload frmPassLock
    DoEvents
    App.title = "Tidal Information..."

End Sub


Private Sub Mnu_Equinox_Click()
If Equinox.Visible = True Then Unload Equinox: Exit Sub
Unload Equinox: Equinox.Show


End Sub

Private Sub mnu_Excute_Click()
Execute.Show
End Sub

Private Sub Mnu_News_Click()
'Open "C:" + MYDocU$ + "My Webs\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_Text2.html" For Input As #1
End Sub

Private Sub Mnu_Expand_W_Click()

WExpand = Not WExpand

On Local Error Resume Next


'----------------------
'UP AND DOWN AND HEIGHT
'----------------------

Call WCHC(WC, HC)
On Error GoTo 0

'ATidalX.Width = WC

Me.Width = WC
Me.Height = HC + (Menu_Height * Screen.TwipsPerPixelY) - 420



HwndCTLTask4 = FindWindow(vbNullString, "WindowsFormsParkingWindow")

'OFFICE TOOLBAR
HwndCTLTask4 = FindWindow("MOM Class", vbNullString)

Hwnd24 = GetWindowRect(HwndCTLTask4, Rect4)

'HERE ON START 1ST
If ME_POSITION_BEEN_SET_ONCE = False Then
If Hwnd24 > 0 Then
    If Rect4.Bottom - Rect4.Top <= 32 Then
        ATidalX.Top = ((Rect4.Bottom - Rect4.Top) * Screen.TwipsPerPixelY)
        ATidalX.Left = Screen.Width - ATidalX.Width
    End If
End If
If Hwnd24 = 0 Then
        ATidalX.Top = 0
        ATidalX.Left = Screen.Width - ATidalX.Width
    End If
End If
OldTopTidalPos = -1

ATidalX.WindowState = 0
ATidalX.Visible = True

On Error Resume Next
For Each Control In Controls
    TEST_VAR1 = "**LABEL COMMAND PROGRESS LBLVOL"
    TEST_VAR2 = "**COMMAND"
    If InStr(TEST_VAR1, Mid(UCase(Control.Name), 1, 4)) > 0 Then
        'Debug.Print str(Control.Left) ', Control.Name
        If Control.Left < 200 Then
            Control.Left = 0
        End If
    End If
    If InStr(TEST_VAR2, Mid(UCase(Control.Name), 1, 4)) > 0 Then
        If Control.Left < 200 Then
            Control.Left = -15
        End If
    End If
    
    SPY = Control.Name
    If InStr(TEST_VAR1, Mid(UCase(Control.Name), 1, 4)) > 0 Then
        If Control.Height + Control.Top > ME_HEIGHT_MAX Then
            ME_HEIGHT_MAX = Control.Height + Control.Top
        End If
        RESULT1 = Control.Left <= 0 And Control.Width > Me.Width - 300
        'RESULT2 = Control.Left < Me.Width + 500
        RESULT2 = True
        If RESULT1 And RESULT2 Then
            Control.Width = Me.Width
        End If
        RESULT1 = Control.Left + Control.Width > Me.Width - 400
        RESULT2 = Control.Left < Me.Width + 200
        If RESULT1 And RESULT2 Then
            'Control.Width Me.Width / 2
            Control.Left = Me.Width - Control.Width
        End If
    End If
Next

Label7.Width = Me.Width - Label37.Width
LABEL_TIME.Width = Me.Width - Label40.Width

ProgressBarLock.Width = Me.Width - LblVol.Width
ProgressBarVol1.Width = Me.Width - LblVol.Width
ProgressBarVol2.Width = Me.Width - LblVol.Width
ProgressBarLock.Width = Me.Width - LblVol.Width
Slider1.Width = ProgressBarLock.Width


'NOT USED MOMENT
Call SeXxYy


If ME_POSITION_BEEN_SET_ONCE = False Then
    ME_POSITION_BEEN_SET_ONCE = True
End If


End Sub


Function GetMenuLowXCord()





Exit Function


Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
Dim WinClass As String, WinTitle As String

Windows = ChildWindows(Me.hWnd)
Dim XEasy

For r = 1 To UBound(Windows)
    Retval = GetClassName(Windows(r), WinClassBuf, 255)
    WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
    
    If WinClass = "Button" Or WinClass = "mciWndClassXXX" Then
        
        WT = GetWindowTitle(Windows(r))
        'Debug.Print WT
        Hwnd24 = GetWindowRect(Windows(r), Rect4)
        If XEasy < Rect4.Bottom Then XEasy = Rect4.Bottom
        'Debug.Print WinClass + " -- " + str(XEasy) + " -- " + str(Rect4.Bottom) + " -- " + str(Rect4.Top)
    End If
    

Next

GetMenuLowXCord = -1

If XEasy = 0 Then Exit Function

GetMenuLowXCord = XEasy '* Screen.TwipsPerPixelY
Debug.Print GetMenuLowXCord
End Function

Sub WCHC(WC, HC)

'Dim WC, HC
WC = 0: HC = 0
On Local Error Resume Next
For Each Control In ATidalX
    TTH2 = 0
    TTH = LCase(Control.Name)
    'If Control.Name = "Label2" Then Stop
    CHECK_STRING = "**MNU_ DIR LIST TEXT LIST TIMER"
    If InStr(CHECK_STRING, Mid(UCase(Control.Name), 1, 4)) > 0 Then
        TTH2 = 1
    End If
    If InStr(CHECK_STRING, Mid(UCase(Control.Name), Len(Control.Name) - 3, 4)) > 0 Then
        TTH2 = 1
    End If
    CTT1 = UCase("**shellwinamploggloader shellkeyload filesavefileexit")
    CTT2 = UCase("filesavefileexit")
    CHECK_STRING = UCase(CTT1 + CCT2)
    If InStr(CHECK_STRING, Mid(UCase(Control.Name), 1, 10)) > 0 Then
        TTH2 = 1
    End If

'    Err.Clear
    'TTH3 = Control.Interval
    'If Err.Number = 0 Then TTH2 = 1
    
    If TTH2 = 0 Then
        
        Err.Clear
        ii22 = Control.Visible
        II23 = Control.Width
        If Err.Number > 0 Then ii22 = False
        
        If (Control.Enabled = True And ii22 = True) Or WExpand = True Then
            'Debug.Print Control.Caption
        
        
            If Control.Left + Control.Width > WC Then
                WC = Control.Left + Control.Width
            End If
        
            If Control.Top + Control.Height > HC Then
                HC = Control.Top + Control.Height
                'Debug.Print Control.Name + " -- " + str(HC)
            End If

        End If
    End If
Next


'WC = Label23.Left + Label23.Width
'If WExpand = True Then WC = WC + 400
'HC = Label23.Top + Label23.Height
End Sub


Private Sub MNU_EXPLORER_LOGG_Click()

Shell "Explorer.exe /e, " + App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\", vbNormalFocus

End Sub

Private Sub Mnu_ExtremeLockLogg_Click()
'Shell "notepad " + App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Extreme Lock.txt", vbNormalFocus
FILE_NAME = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Extreme Lock.txt"
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing
End Sub

Public Sub Mnu_Idel_Click()
'Shell "notepad " + App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Idle and Active Logger Day.txt", vbNormalFocus
FILE_NAME = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Idle and Active Logger Day.txt"
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing
End Sub

Public Sub Mnu_Idle_Now_Click()
'Shell "notepad " + App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Idle and Active Logger Now.txt", vbNormalFocus
FILE_NAME = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Idle and Active Logger Now.txt"
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing
End Sub

Private Sub Mnu_KeyCleaning_Click()
KeyClean = Not KeyClean
'If KeyClean = True Then MsgBox "All Keys Will be Blocked For Keyboard Cleaning"
'If KeyClean = False Then MsgBox "All Keys Will be UnBlocked Nice an Shine"

Mnu_KeyCleaning.Checked = KeyClean

End Sub

Private Sub Mnu_KeyLog_Click()

'Shell "notepad """ + GetSpecialfolder(CSIDL_PERSONAL) + "\CallerID\01 KeyStroke_Logger_" + GetComputerName + "-" + GetUserName + ".txt""", vbNormalFocus

FILE_NAME = GetSpecialfolder(CSIDL_PERSONAL) + "\CallerID\01 KeyStroke_Logger_" + GetComputerName + "-" + GetUserName + ".txt"

' FILE_PATH = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\"

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing

' Shell "Explorer.exe /e, " + App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\", vbNormalFocus

Shell "Explorer.exe /e, " + FILE_NAME, vbNormalFocus
End Sub

Private Sub Mnu_KHook_Click()
DSkeybd_F.Show
End Sub

Public Sub Mnu_Knocker1_Click()
'Shell "notepad " + App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy1.txt", vbNormalFocus
FILE_NAME = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy1.txt"
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing
End Sub

Public Sub Mnu_Knocker2_Click()
'Shell "notepad " + App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy2.txt", vbNormalFocus
FILE_NAME = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy2.txt"
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing
End Sub

Public Sub Mnu_Knocker3_Click()
'Shell "notepad " + App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy3.txt", vbNormalFocus
FILE_NAME = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy3.txt"
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing
End Sub

Public Sub Mnu_Knocker4_Click()
'Shell "notepad " + App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy4.txt", vbNormalFocus
FILE_NAME = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy4.txt"
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing
End Sub

Private Sub Mnu_LoadVBanKill_Click()
'Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE " + App.Path + App.EXEName + ".vbp"
'End
End Sub

Private Sub Mnu_NewsClipText_Click()
Exit Sub
Fref5 = FreeFile
Open "C:" + MYDocU$ + "My Webs\MatthewLan.Com Web\LoveFolder\Cool\LinkPage_Text3.html" For Input As Fref5

Ar$ = Input$(LOF(Fref5), Fref5)
Close #Fref5
Clipboard.Clear
Clipboard.SetText Ar$, vbCFText


End Sub

Private Sub Mnu_Norton_Test_Click()
'TestTrig = True
'Call AntiVirusMatt
'TestTrig = False

End Sub

Private Sub Mnu_Note_Url_Big_Click()
'Shell "notepad " + GetSpecialfolder(CSIDL_PERSONAL) + "\01 Palm HotSync\UrlLogg-Long-" + GetComputerName + GetUserName + ".txt", vbNormalFocus
FILE_NAME = GetSpecialfolder(CSIDL_PERSONAL) + "\01 Palm HotSync\UrlLogg-Long-" + GetComputerName + GetUserName + ".txt"
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing
End Sub

Private Sub Mnu_NotebookURL_Click()
'Shell "notepad " + GetSpecialfolder(CSIDL_PERSONAL) + "\01 Palm HotSync\UrlLogg1-" + GetComputerName + GetUserName + ".txt", vbNormalFocus
FILE_NAME = GetSpecialfolder(CSIDL_PERSONAL) + "\01 Palm HotSync\UrlLogg1-" + GetComputerName + GetUserName + ".txt"
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing
End Sub

Private Sub Mnu_QuickWeb_Click()
    QuickPage = True
End Sub

Private Sub Mnu_Recycle_Click()
    Shell "D:\VB6\VB-NT\00 Sort Anything\Sort_Anything_Scan_Recycle_Bin\Sort Anything Recycle Bin.exe", vbNormalFocus
End Sub

Private Sub Mnu_Say_SolarNoon_Click()

Empire = 3 'SolarNoon
Call SunRise_Bong

End Sub

Private Sub Mnu_Say_Sunrise_Click()

Empire = 1 'SunRise
Call SunRise_Bong

End Sub

Private Sub Mnu_Say_SunSet_Click()

Empire = 2 'SunRise
Call SunRise_Bong

End Sub

Private Sub Mnu_Say_Time_Click()

Beep

'Est = 1 'make say time?

'SayTime3 = True 'make say time? -- FOR THE MALL LOGG ALSO
Call Time_Code

End Sub

Private Sub Mnu_SayMoon_Click()
Beep

Xcag = 2
Segm9 = 0

USER_REQUEST_TEST_MOON_CHANGE_STATE_AND_TALK = True

End Sub

Private Sub Mnu_ShowURL_Click()
frmIECtrl.Show
End Sub
Private Sub Mnu_HideURL_Click()
frmIECtrl.Hide

End Sub


Private Sub Mnu_ShowWeb_Click()
Call ShowWeb
End Sub

Private Sub Mnu_SigLine_Click()

TTx$ = ""
Call Timer_CountD1Text_timer

'Shell "Notepad "+GetSpecialfolder(CSIDL_PERSONAL)+"\01 Email Settings\Signatures\Signature-Nortons2.txt", vbNormalFocus
'Shell "Notepad "+GetSpecialfolder(CSIDL_PERSONAL)+"\01 Email Settings\Signatures\Signature-Nortons2Small.txt", vbNormalFocus
'If FS.FileExists(GetSpecialfolder(CSIDL_PERSONAL) +"\01 Email Settings\Signatures\Signature-Nortons2.txt") = False Then Stop

'Shell "Notepad.exe " + GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-Nortons2.txt", vbNormalFocus
FILE_NAME = GetSpecialfolder(CSIDL_PERSONAL) + "\01 Email Settings\Signatures\Signature-Nortons2.txt"
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing

End Sub

Private Sub Mnu_SW_Visits_Click()

'Call WorkVisits
    
'Shell "Notepad " + GetSpecialfolder(CSIDL_PERSONAL) + "\CallerID\#Since SW Visits.txt", vbNormalFocus
FILE_NAME = GetSpecialfolder(CSIDL_PERSONAL) + "\CallerID\#Since SW Visits.txt"
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing

End Sub

Private Sub mnu_Switches_Click()
Menu.Show
End Sub

Private Sub Mnu_SysUpTime_Click()
'Shell "Notepad " + App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\System UpTime Logg.txt", vbNormalFocus
FILE_NAME = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\System UpTime Logg.txt"
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + FILE_NAME + """"
Set WSHShell = Nothing
End Sub

Private Sub Mnu_Terminate_Click()
'Process_Kill
t = 0
For r = 1 To ExecuteWeb.List1.ListCount - 1
    If Nb(r) <> 0 Then
        Process_Kill (Nb(r))
    End If
Next
End Sub

Private Sub Mnu_Time_Syne_Click()

Shell "D:\VB6\VB-NT\00_Best_VB_01\Time_Sync\Time Sync.exe", vbNormalNoFocus

End Sub

Private Sub Mnu_VB6_Click()
Shell "Explorer.exe /e, D:\VB6\VB-NT\00 Vb Shortcuts", vbNormalFocus

End Sub

Private Sub Mnu_VB_ME_Run_Click()
'LOAD AND RUN


Exit Sub

If IsIDE = True Then Stop: End
ATidalX.Visible = False
ATidalX.Enabled = False
Me.WindowState = 1
'Unload Me
Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\Tidal.vbp""", vbNormalFocus
Unload Me
End

End Sub

'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
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
    objShell.RUN """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
    Set objShell = Nothing
    
    EXIT_TRUE = True
    Unload Me
    Exit Sub

    
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
Private Sub MNU_VB_FOLDER_Click()
    Beep
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    Shell "Explorer.exe /SELECT, """ + App.Path + "\" + CODER_EXE_FILE_NAME_1 + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
End Sub

Private Sub MNU_ADMINISTRATOR_Click()
    
    MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR MODE ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "MODE ON", "MODE OFF")
    End If
    If MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR MODE ON" Then
        MNU_ADMINISTRATOR.Caption = "ADMIN ON"
        MNU_ADMINISTRATOR.Visible = True
    End If
    
End Sub
'-------------------------------------------
'-------------------------------------------
'---------------------------------------------------
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



Private Sub Mnu_VBISIDE_Click()
    Mnu_VBISIDE.Checked = Not Mnu_VBISIDE.Checked
    
    FR1 = FreeFile
    Fd$ = "o"
    Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\ISIDE.txt" For Output As #FR1
        If Mnu_VBISIDE.Checked = True Then Fd$ = "I"
        Print #FR1, Fd$
    Close #FR1
End Sub


Private Sub Mnu_Voices_Click()

Call Mnu_Voice_Click

End Sub

Private Sub Mnu_Voice_Click()
    
    Load TTSAppMain
    TTSAppMain.Show

End Sub

Private Sub Mnu_VolMax_Click()
'Volume.SetVolume (100)
End Sub

Private Sub Mnu_VolMin_Click()
'Volume.SetVolume (0)
TTVol = 120 * 60 '2 hours

End Sub

Private Sub Mnu_VolMinMusic_Click()
    'Volume.SetVolume (100)
    VolumeSet1 = 1

End Sub

Private Sub Mnu_WebPage_Click()
ExecuteWeb.Show
End Sub

Public Function GetActiveWindowTitle(ByVal ReturnParent As Boolean) As String
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
   GetActiveWindowTitle = GetWindowTitle(i)
End Function


Public Sub PassLockTimer_Timer()
    Exit Sub
    'PROBLEM TIMERS LOAD FORM
    On Local Error Resume Next

    Keleidoscope.HighRes.Enabled = False
    Bezier.Timer1.Enabled = False
    YinYang.Timer1.Enabled = False
    Vector.Timer1.Enabled = False
    Vector.Timer2.Enabled = False

    Unload Bezier
    Unload Keleidoscope
    Unload YinYang
    Unload Vector
    Unload frmPassLock
    PassLockTimer.Enabled = False

End Sub

Private Sub ProgressBarVol1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'ProgressBarVol
Slider1.Visible = True
SliderTimer.Enabled = True
SliderC = Now + TimeSerial(0, 0, 5)
End Sub

Private Sub ProgressBarVol2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'ProgressBarVol
Slider1.Visible = True
SliderTimer.Enabled = True
SliderC = Now + TimeSerial(0, 0, 5)
End Sub


Public Sub ShellKeyLoad_Timer()
Exit Sub
ShellKeyLoad.Enabled = False
If FindWindow(vbNullString, "KeyBoard Auto Loader") > 0 Then Exit Sub

Shell Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\Shell KeyBoard Loader\Shell KeyBoard Loader.exe", vbNormalFocus

End Sub

Private Sub ShellWinampLoggLoader_Timer()

ShellWinampLoggLoader.Enabled = False
Exit Sub
'------------------------------------

ShellWinampLoggLoader.Enabled = False
If FindWindow(vbNullString, "D:\VB6\VB-NT\00_Best_VB_01\WinAmp Logger\00 Music Info Logger.txt - Notepad2") > 0 Then Exit Sub
If FindWindow(vbNullString, "*D:\VB6\VB-NT\00_Best_VB_01\WinAmp Logger\00 Music Info Logger.txt - Notepad2") > 0 Then Exit Sub
Shell Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\WinAmp Logger\WinAmp Logger.exe", vbNormalFocus
End Sub

Private Sub Slider1_Click()
'Slider1
End Sub


Private Sub SliderTimer_Timer()
        
        If SliderC < Now And SliderC > 0 Then
            Slider1.Visible = False
            SliderC = 0
            SliderTimer.Enabled = False
        End If
        
End Sub


Private Sub Text1_Click()
yy$ = GetActiveWindowTitle(i)

Clipboard.Clear
Clipboard.SetText yy$ + Text1.Text + vbCrLf

End Sub


Private Sub Timer1_Timer()
'TIMER1.ENABLED = FALSE
'Timer1.Interval = 10

If Day(Now) <> O_BONG_DAY_NOW Then
    If Wqe2$ <> "" Then
        ATidalX.MMControl2.Command = "prev"
        ATidalX.MMControl2.Command = "Play"
    Else
        ATidalX.MMControl1.Command = "prev"
        ATidalX.MMControl1.Command = "Play"
    End If

    Call WaitWavFinish
End If
O_BONG_DAY_NOW = Day(Now)

GoTo JUMPSET1


If KATTIMER = 0 Then KATTIMER = Now + TimeSerial(0, 2, 0)
If KATTIMER < Now Then
    XXHOOK = DSkeybd_F.IsHook
    If DSkeybd_F.IsHook = False Then
        Call Mnu_KeyHook_Click
    End If
    
    KATTIMER = 0
    
    'CTRL ALT EQUAL SIGN
    Call SendKey(187, 6)
    'CTRL ALT SPACE
    DoEvents
    Call SendKey(&H20, 6)
    If XXHOOK = False And DSkeybd_F.IsHook = True Then
        Unload DSkeybd_F

    End If
End If

JUMPSET1:

'If OHOUR5 <> HOUR(Now) Then
'    OHOUR5 = HOUR(Now)
'    tf = SetVolume(WAVEOUT, 100)
'End If


If MNU_LATENIGHTOVERRIDE.Checked = False Then
If HOUR(Now) >= 1 And HOUR(Now) < 7 Then
    SayTimeVolAdjustLateNite = True
    RXT = "NIGHT TIME VOLUME'S NOW AT 1 OCLOCK AND LESS TALK TIME ACTIVITY"
    RXT = "NIGHT VOLUME'S AND ACTIVITY"
Else
    SayTimeVolAdjustLateNite = False
    RXT = "DAY TIME VOLUME'S NOW AT 7 OCLOCK WITH TALK TIME ACTIVITY HIGHER"
    RXT = "DAY VOLUME'S AND ACTIVITY"
End If
End If

If SayTimeVolAdjustLateNite = True And MNU_LATENIGHTOVERRIDE.Checked = True Then
    RXT = "Night Time Volumes and Activity have Been Manually OverRiden To Day Values"
    SayTimeVolAdjustLateNite = False
End If


If oRXT <> RXT And oRXT <> "" Then
    'Call VolumeFixedSet(0, True)
    Call SpeakVoice(RXT)
End If
oRXT = RXT




WINAMP2 = FindWindow("Winamp v1.x", vbNullString)
If WINAMP2 > 0 Then
    MSGRESULT1 = SendMessage(WINAMP2, WM_USER, 0, ByVal ipc_isplaying)
    MsgResult2 = SendMessage(WinAmpChkIsPlayingHwnd, WM_USER, 0, ByVal ipc_isplaying)
    If MSGRESULT1 = 1 Or MsgResult2 = 1 Then
        WinAmpStoppedPlayTime = 0
        HasWinAmpBeenPlayingRecent = Now + TimeSerial(0, 0, 15)
    
    End If
End If

'WORKING ON HERE
If WinAmpStoppedPlayTime > 0 Then
    If DateDiff("n", WinAmpStoppedPlayTime, Now) > 120 Then
        'VolumeSet1 = 1
        'MsgBox "Yes"
    End If
End If

'WORKING ON HERE
If WinAmpStoppedPlayTime > 0 Then
    If DateDiff("n", WinAmpStoppedPlayTime, Now) > 30 And SayTimeVolAdjustLateNite = True Then
        'VolumeSet2 = MinVol2Set
        'MsgBox "Yes"
    End If
End If


Dim TimeSet As Date

If 1 = 2 Then
    
    TimeSet = TimeSerial(6, 58, 0)
    
    TestRun = False
    
    
    If Test5 = 0 Then
        'TimeSet = TimeValue(Now) + TimeSerial(0, 0, 20)
        Test5 = Int(Now) + TimeSet ': LockSaverDelay = LockSaverDelay4'--- Short delay even shorter for IDE
        
        'Temp ByPass
        Test5 = Int(Now) + 40
        
        If TestRun = True Then Test5 = Now + TimeSerial(0, 0, 10) 'Test
        
        If DateDiff("s", Now, Test5) < 0 Then
            Test5 = Int(Now) + 1 + TimeSet
        End If
    End If
    
    'If IsIDE Then LockSaverDelay4 = TimeSerial(0, 0, 5) Else LockSaverDelay4 = TimeSerial(0, 1, 0)
    
    
    If ProgressBarLock.Value > 0 And App.title = "Tidal Information..." Then
    If Now + TimeSerial(0, 0, 10) > Test5 Then
        If LockSaverDelay <> LockSaverDelay4 Then  'LockSaverDelay4 '--- Short delay even shorter for IDE
            LockSaverDelay = LockSaverDelay4
            LockSSaver = Now + LockSaverDelay
            Call SetLockMax
        End If
    End If
    End If
    
    'If Now > Test And Test > 0 Then
    '    Test = -2
    '    ATidalX.MNU_WinampFadeOut_Click
    'End If
    
    If ProgressBarLock <= 0 And LockDontAllowHotKeyforPassword = False Then
    
        If Now > Test5 And App.title = "Tidal Information..." Then
            HittMan = True
            XxG = Now
            'LockSaverDelay = TimeSerial(0, 10, 0)
            'SetLockTime
            Call PassModule.Main
    '        SetLockTime
    'LockSaverDelay = TimeSerial(0, 10, 0)
                
            Test5 = Int(Now) + 1 + TimeSet 'Next Day
            
            Test7 = Now + TimeSerial(1, 30, 0) ' stay in lock 1 half hour -- as long as no key press that is set by
            If TestRun = True Then Test7 = Now + TimeSerial(0, 0, 5) ' stay in lock 10 secs in test
        End If
    End If



    If Now > Test7 And Test7 > 0 And App.title <> "Tidal Information..." Then
        'frmPassLock.Hide
        HittMan2 = True
        Unload frmPassLock
        HittMan2 = False
        Test5 = Int(Now) + 1 + TimeSet
        If TestRun = True Then Test5 = Now + TimeSerial(1, 0, 30)   'Test
        Test7 = Test5 + TimeSerial(1, 0, 0)
         'SetLockTime
    
        'LockSaverDelay = LockSaverDelay4
        'LockSSaver = Now + LockSaverDelay
        'Call SetLockMax
    End If

End If '1=2


If LockDown2 > Now And LockSSaver < Now Then
    'XxG = Now
    'ATidalX.MNU_WinampFadeOut_Click
    'Call PassModule.Main
End If

'qq8 = GetIdleM(5)


If Rabbit < Now And Rabbit > 0 Then
    Rabbit = 0
    UpDateText.Visible = False
End If



'timer1.Interval
'Debug.Print Time

'--------------------
'--------------------
'DAYLIGHT TALK ARRIVE FROM THIS ONE -- Percent_SUN_1_100
'--------------------
'Call Percent_SUN_1_100 ' Sub Holds the Talk Day Light % Event
'Routine Called by Dependant of
'Call TimerSunRiseSetPercent_Timer
'--------------------
'--------------------


Call DateMidnight_CountDown

'LESS QUICKER
If Second(TIME_ONE_SECOND_COUNTER) <> Second(Now) Then
    Call Time_Code
    Call SunRise_Bong
    Call SunRise_SunSet
    Call Tidal_Code
    Call RSR232_FRONT_DOOR_LOGGER_TALK
End If

TIME_ONE_SECOND_COUNTER = Now

Call Moon_Code
Call Percent_Moon_1_100(MOON_LIGHT)
    
    
    wk1$ = DateDiff("ww", DateSerial(Year(Now), 1, 1), Now) + 1
    wk3 = DateDiff("ww", DateSerial(Year(Now), 6, 1), Now) + 1
    wk5 = DateDiff("ww", DateSerial(Year(Now), 1, 1), Now)
    wk7$ = Format$(DateSerial(Year(Now), 1, (wk5 * 7)) + 1, "DDD DD-MMM")
    
    'WR# shows update on day of WR
    'Offset
    daywr1 = 4 'Monday
    
    wk2$ = DateDiff("w", DateSerial(2007, 6, daywr1) - 5, Now) + 1
    
    daywr2 = DateSerial(2007, 6, daywr1)
    'daywr3 = DateSerial(2008, 1, 14)
    daywr3 = DateSerial(2008, 2, 20) 'Now
    wk8 = DateDiff("d", daywr2, daywr3)
    wk2$ = DateDiff("w", daywr2, daywr3) + 1
    wk4$ = Format$(daywr2 + (wk2 * 7) - 5, "DD MMM")
    
    dtm1 = DateSerial(Year(Now), Month(Now), 1)
    dtm2 = DateSerial(Year(Now), Month(Now) + 1, 1)
    dtmd$ = DateDiff("d", dtm1, dtm2)
    dtm1 = DateSerial(Year(Now), Month(Now) - 1, 1)
    dtm2 = DateSerial(Year(Now), Month(Now), 1)
    dtmL$ = DateDiff("d", dtm1, dtm2)
    dtm1 = DateSerial(Year(Now), Month(Now) + 1, 1)
    dtm2 = DateSerial(Year(Now), Month(Now) + 2, 1)
    dtmn$ = DateDiff("d", dtm1, dtm2)
    
    dtm1 = DateSerial(Year(Now), Month(Now), 1)
    dtm2 = DateSerial(Year(Now), Month(Now) + 1, 1)
    For r = dtm2 - 1 To dtm2 - 5 Step -1
    If Weekday(r) > 1 And Weekday(r) < 7 Then Exit For
    Next
    
    XX$ = ""
    If r = dtm2 - 1 Then XX$ = " Last Day"
    xy$ = DateDiff("d", dtm2, r) + 1
    
    'Label27.Caption = "Wk#" + wk1$ + "-" + wk7$ + " -- Nx-WR#" + wk2$ + "-" + wk4$
    Label27.Caption = "Wk#" + wk1$ + "-" + wk7$
    Label28.Caption = "Days This Month#" + dtmd$ + " LM#" + dtmL$ + " NM#" + dtmn$
    Label29.Caption = "Pay Day " + Format$(r, "DDD dd-mmm") + " " + xy$ + " Days"
    
    
    'PILLS WORK OUT
    GoTo JUMPSET2
    
    xy$ = DateDiff("d", DateSerial(2007, 6, 1), Now) + 1
    
    xy$ = DateDiff("d", DateSerial(2008, 11, 30), DateSerial(2009, 3, 23)) + 1
    
    
    '06 Jan 1997 - 12 Feb 1997
    XY1 = DateDiff("d", DateSerial(1997, 1, 6), DateSerial(1997, 2, 12)) + 1
    
    '28 Jan 1998 - 24 Mar 1998
    XY1 = XY1 + DateDiff("d", DateSerial(1998, 1, 28), DateSerial(1998, 3, 24)) + 1
    
    '19 Aug 1998 - 03 Nov 1998
    XY1 = XY1 + DateDiff("d", DateSerial(1998, 8, 19), DateSerial(1998, 11, 3)) + 1
    
    '11 Dec 2002 - 13 Feb 2003
    XY1 = XY1 + DateDiff("d", DateSerial(2002, 12, 11), DateSerial(2003, 2, 13)) + 1
    
    '19 Oct 2005 - 11 Jan 2006
    XY1 = XY1 + DateDiff("d", DateSerial(2005, 10, 19), DateSerial(2006, 1, 11)) + 1
    
    '01 Jun 2007 - 01 Feb 2008
    XY1 = XY1 + DateDiff("d", DateSerial(2007, 6, 1), DateSerial(2008, 2, 1)) + 1
    
    'Rutland
    '01 Feb 2008 - 30 Nov 2008
    XY1 = XY1 + DateDiff("d", DateSerial(2007, 2, 1) - 1, DateSerial(2008, 11, 30)) + 1
    
    'Hosso still Count for rutland
    
    'Hosso 7th
    '30 Nov 2008 - 30 Mar 2009
    XY1 = XY1 + DateDiff("d", DateSerial(2008, 11, 30) - 1, DateSerial(2009, 3, 30)) + 1
    
    Err.Clear
    On Error Resume Next
    Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Pills.txt" For Input As #1
    Line Input #1, dd1$
    Line Input #1, dd2$
    Line Input #1, dd3$
    Close #1
    If Err.Number > 0 Then Exit Sub
    On Error GoTo 0
    
    XY3 = DateDiff("d", DateValue(dd1$), Now)
    
    'Easy Just change the pill amount to wat you taking whenever you want an updated the excell file EASY!
    PillAmountCurrentlyTaking = 4
    
    PillsYes = Val(dd3$) + (XY3 * PillAmountCurrentlyTaking)
    PillsNo = Val(dd2$) + (XY3 * PillAmountCurrentlyTaking)
    
    XY3 = DateDiff("d", DateSerial(2007, 12, 21), Now)
    DaysHossTotalVar = XY1
    XY1 = XY1 + DateDiff("d", DateSerial(2009, 3, 30) + 1, DateSerial(2009, 5, 1)) + 1
    DaysHossInpaitentVar = XY1
    
    Label32.Caption = "Days Paitent" + str(XY1) + " Total" + XY2$ + " Pills" + str(PillsYes)

JUMPSET2:

End Sub

Sub MAKETALKTIMESTRING()
    
    TALKTIME$ = ""
    R2 = 0
    YYt = Minute(Now)
    YYx = Minute(Now)
    
    For R3 = 0 To 59 Step 30
        If YYt = R3 Then R2 = 1: Exit For
    Next
    For R3 = 0 To 59 Step 15
        If YYt = R3 Then R2 = 1: Exit For
    Next
    For R3 = 0 To 59 Step 20
        'If YYt = R3 Then R2 = 1: Exit For
    Next

    yyg = HOUR(Now)
    If yyg > 12 Then yyg = HOUR(Now) - 12
    If yyg = 0 Then yyg = 12
    
    'MINUTE -----------------
    '------------------------
    'IF THE MINUTE = THE HOUR
    If YYt = yyg Then R2 = 1
    '------------------------
    'If YYt = 3 Then R2 = 1
    'If YYt = 5 Then R2 = 1
    If YYt = 11 Then R2 = 1
    If YYt = 22 Then R2 = 1
    If YYt = 33 Then R2 = 1
    If YYt = 44 Then R2 = 1
    If YYt = 55 Then R2 = 1
    'If YYt = 59 Then R2 = 1
    
'    YYx = 59
'    yyg = 12
'    If yyg > 12 Then yyg = yyg - 12
'    If yyg = 0 Then yyg = 12
    
    '-------------------------------------------------------
    'IF THE MINUTE = THE HOUR HEADING TOWARD IT SUBSTRACTING
    '-------------------------------------------------------
    For r = 1 To 12
        If YYx = 60 - r Then
            If yyg + 1 = r Then R2 = 1
            If yyg = 12 And r = 1 Then R2 = 1
        End If
    Next
    
    If HOUR(Now) = 23 And Minute(Now) = 33 Then R2 = 1
    If HOUR(Now) = 23 And Minute(Now) = 38 Then R2 = 1
    If HOUR(Now) = 23 And Minute(Now) = 40 Then R2 = 1
    If HOUR(Now) = 23 And Minute(Now) = 48 Then R2 = 1
    If HOUR(Now) = 23 And Minute(Now) = 49 Then R2 = 1
    If HOUR(Now) = 23 And Minute(Now) = 50 Then R2 = 1
    If HOUR(Now) = 23 And Minute(Now) = 57 Then R2 = 1
    If HOUR(Now) = 21 And Minute(Now) = 57 Then R2 = 1
    If HOUR(Now) = 21 And Minute(Now) = 59 Then R2 = 1
    
    '--------------------------------
    'Another one further on ---------
    '--------------------------------
    'EVERY HALF AN HOUR
    If Minute(Now) Mod 30 = 0 Then R2 = 1
'    EVERY TEN MINS
'    If Minute(Now) Mod 10 = 0 Then R2 = 1
'    If Minute(Now) Mod 15 = 0 Then R2 = 1

    ' ---------------------------------------------------------
    ' EVERY 5 MIN
    If Minute(Now) Mod 5 = 0 Then R2 = 1
    ' ---------------------------------------------------------

'    1 MIN
'    If Minute(Now) Mod 1 = 0 And SayTimeVolAdjustLateNite = False Then R2 = 1
    
'    If HOUR(Now) = 0 And Minute(Now) = 0 Then 'Call VoiceMP3Pause3
'    If HOUR(Now) = 4 And Minute(Now) = 0 Then 'Call VoiceMP3Pause
'    If HOUR(Now) = 6 And Minute(Now) = 30 Then 'Call VoiceMP3Pause3
'    If HOUR(Now) = 12 And Minute(Now) = 0 Then 'Call VoiceMP3Pause3
'    If HOUR(Now) = 18 And Minute(Now) = 30 Then 'Call VoiceMP3Pause3
    'Call VoiceMP3QuickPause


    TALKTIME$ = Mid$(Format$(Now, "HH:MM:SSampm"), 1, 5)
    
    pp1 = HOUR(Now) + 1
    pp2 = HOUR(Now)
    If pp1 > 12 Then pp1 = pp1 - 12
    If pp2 > 12 Then pp2 = pp2 - 12
    
    
    O_TALKTIME = TALKTIME$
    If Minute(Now) = 0 Then
        If HOUR(Now) <> 0 Then TSSP$ = " TODAY"
        
        '### EVERYDAY STRING IS SET FROM COUNTDOWN SUB
        'If TSSP$ = "" And HOUR(Now) = 0 Then TSSP$ = " Every Day"
        TALKTIME$ = str$(Val(Mid$(TALKTIME$, 1, 2))) + " O'Clock" + "  " + TSSP$: TSSP$ = ""
        Call TIMER_SECONDS20_AFTER_HOUR_TIMER_TRIG_START
    End If
    If TALKTIME$ <> O_TALKTIME Then R2 = 1
    
    
    'THIS ONE GETS IT ON 3PM
    O_TALKTIME = TALKTIME$
    If (HOUR(Now) = 3 Or HOUR(Now) = 15) And Minute(Now) = 0 Then
        TALKTIME$ = "The Time is Free of Charge, TODAY"
    End If
    If TALKTIME$ <> O_TALKTIME Then R2 = 1

    O_TALKTIME = TALKTIME$
    If (HOUR(Now) = 2 Or HOUR(Now) = 14) And Minute(Now) = 30 Then
'        rand2 = Int((Rnd(1) * 6))
        Rand2 = (Now - 1) Mod 7
        If Rand2 = 0 Then TALKTIME$ = "Half Past 2"
        If Rand2 = 1 Then TALKTIME$ = "2 Thirty"
        If Rand2 = 2 Then TALKTIME$ = "Too Thirsty"
        If Rand2 = 3 Then TALKTIME$ = "Too Dirty"
        If Rand2 = 4 Then TALKTIME$ = "Tooth Hurty"
        If Rand2 = 5 Then TALKTIME$ = "Too Flirty"
        If Rand2 = 6 Then TALKTIME$ = "Dirty Old Bastard"
        'MsgBox talktime$
    End If
    If TALKTIME$ <> O_TALKTIME Then R2 = 1
    
    If Minute(Now) = 15 Then
        TALKTIME$ = "Quarter Past " + Trim(Mid(Format(Now, "H AMPM"), 1, 2))
    End If
    HOUR_ADD_ONE = HOUR(Now) + 1
    If HOUR_ADD_ONE > 23 Then Debug.Print "HELLO MIDNIGHT"
    If Minute(Now) = 45 Then
        TALKTIME$ = "Quarter To " + Trim(Mid(Format(Now + TimeSerial(1, 0, 0), "H AMPM"), 1, 2))
    End If
    
    ' EVERY MINUTE MODIFY
    ' ------------------------------------------
    If O_Minute_Now <> Minute(Now) Then
        If VAR_TALKER_EVERY_MINUTE_FOR_AN_HOUR > Now Then
            R2 = 1
        End If
    End If
    
    
    If HOUR(Now) = 0 Then
        O_TALKTIME = TALKTIME$
        If Minute(Now) = 5 Then TALKTIME$ = str$(pp2) + " O 5"
        If Minute(Now) = 5 And HOUR(Now) = 0 Then TALKTIME$ = "5 Past Midnight Hour, HAVE, A HAPPY NIGHT HOUR"
        If Minute(Now) = 10 And HOUR(Now) = 0 Then TALKTIME$ = "10 Past Midnight Hour"
        If Minute(Now) = 15 Then TALKTIME$ = "Quarter Pass " + str$(pp2)
        If Minute(Now) = 15 And HOUR(Now) = 0 Then TALKTIME$ = "Quarter Past MidNight Hour"
        If Minute(Now) = 20 Then TALKTIME$ = "20 Past Midnight"
        If Minute(Now) = 25 Then TALKTIME$ = "25 Pass" + str$(pp2)
        
        If Minute(Now) = 35 Then TALKTIME$ = "25 To" + str$(pp1)
        If Minute(Now) = 45 Then TALKTIME$ = "Quarter to " + str$(pp1)
        If Minute(Now) = 40 And HOUR(Now) = 23 Then TALKTIME$ = "20 to MidNight Hour"
        If Minute(Now) = 44 And HOUR(Now) = 23 Then TALKTIME$ = "11 44"
        If Minute(Now) = 45 And HOUR(Now) = 23 Then TALKTIME$ = "Quarter to MidNight Hour"
        If Minute(Now) = 55 And HOUR(Now) = 23 Then TALKTIME$ = "5 to MidNight Hour"
        If Minute(Now) = 59 And HOUR(Now) = 23 Then TALKTIME$ = "1 Minute to MidNight Hour, HAVE, A HAPPY NIGHT HOUR"
        If TALKTIME$ <> O_TALKTIME Then R2 = 1
    End If

'   If Minute(Now) = 55 Then talktime$ = "5 To" + Str$(pp1)

    If Mid$(TALKTIME$, 3, 1) = ":" Then
    
    O_TALKTIME = TALKTIME$
    If Minute(Now) < 3 Then
        TALKTIME$ = str(Val(Mid$(TALKTIME$, 1, 2))) + " o " + str(Minute(Now))
    End If
    If TALKTIME$ <> O_TALKTIME Then R2 = 1
        
    If Minute(Now) < 10 Then
        TALKTIME$ = str(Val(Mid$(TALKTIME$, 1, 2))) + " o " + str(Minute(Now))
    End If
        
        O_TALKTIME = TALKTIME$
        If HOUR(Now) = 0 Then
            If Minute(Now) < 30 Then
                TALKTIME$ = str(Minute(Now))
                If Minute(Now) = 1 Then ttm$ = " Minute" Else ttm$ = " Minutes"
                TTH$ = " Past Midnight"
                TALKTIME$ = TALKTIME$ + ttm$ + TTH$
            End If
        End If
        If TALKTIME$ <> O_TALKTIME Then R2 = 1
        
        O_TALKTIME = TALKTIME$
        If HOUR(Now) = 23 Then
            If Minute(Now) > 40 Then
                TALKTIME$ = str((60 - Val(Minute(Now))))
                If 60 - Val(Minute(Now)) = 1 Then ttm$ = " Minute" Else ttm$ = " Minutes"
                TTH$ = " To Midnight"
                TALKTIME$ = TALKTIME$ + TTH$
            End If
        End If
        If TALKTIME$ <> O_TALKTIME Then R2 = 1
    
        O_TALKTIME = TALKTIME$
        If HOUR(Now) = 12 Then
            If Minute(Now) < 30 Then
                TALKTIME$ = str(Minute(Now))
                TTH$ = " Minutes Past Mid Day"
                TALKTIME$ = TALKTIME$ + TTH$
            End If
        End If
        If TALKTIME$ <> O_TALKTIME Then R2 = 1
    End If
        
    ' ----------------------------------
    ' SOME SORT OF TEN SECOND COUNT DOWN
    ' ----------------------------------
    ' NOT IN USE YET
    ' ----------------------------------
    R3 = 0
    For r = 1 To 12
        If YYx = 60 - r Then
            If yyg + 1 = r Then R3 = 1
            If yyg = 12 And r = 1 Then R3 = 1
        End If
    Next
    yyy = yyg + 1
    If yyy > 12 Then yyy = yyy - 12
    If yyy = 0 Then yyy = 12
    If Minute(Now) = 59 And Second(Now) < 10 Then
        If R3 > 0 Then TALKTIME$ = str$(60 - YYx) + " to " + str$(yyy)
    End If
        
'-----------------------------
        
    TTH$ = " Minutes To Midnight"
    
    O_TALKTIME = TALKTIME$
    If HOUR(Now) = 12 And Minute(Now) = 1 Then TALKTIME$ = "1 Minute Past MidDay"
    If HOUR(Now) = 23 And Minute(Now) = 33 Then TALKTIME$ = "11 33"
    If HOUR(Now) = 23 And Minute(Now) = 38 Then TALKTIME$ = "20 two " + TTH$
    If HOUR(Now) = 23 And Minute(Now) = 40 Then TALKTIME$ = "20 " + TTH$
    If HOUR(Now) = 23 And Minute(Now) = 44 Then TALKTIME$ = "11 44"
    If HOUR(Now) = 23 And Minute(Now) = 48 Then TALKTIME$ = "12 " + TTH$
    If HOUR(Now) = 23 And Minute(Now) = 49 Then TALKTIME$ = "11  " + TTH$
    If HOUR(Now) = 23 And Minute(Now) = 50 Then TALKTIME$ = "10  " + TTH$
    If HOUR(Now) = 23 And Minute(Now) = 55 Then TALKTIME$ = "5 " + TTH$
    If HOUR(Now) = 23 And Minute(Now) = 57 Then TALKTIME$ = "3 " + TTH$
    If HOUR(Now) = 23 And Minute(Now) = 59 Then TALKTIME$ = "1 Minute To Midnight, HAVE, A HAPPY NIGHT HOUR"
    If TALKTIME$ <> O_TALKTIME Then R2 = 1
    
'    If Minute(Now) Mod 1 = 0 And SayTimeVolAdjustLateNite = False Then SayTime2 = True
'    If Minute(Now) Mod 5 = 0 Then SayTime2 = True
'    If Minute(Now) Mod 10 = 0 Then SayTime2 = True
'    If Minute(Now) Mod 15 = 0 Then SayTime2 = True

    '--------------------------
    'SHOULD BE NOT IF ON DEMAND
    '--------------------------
    'For R = 6 To 56 Step 10
    '    If Minute(Now) = R Then SayTime2 = True
    'Next
    '--------------------------
    For r = 6 To 56 Step 10
        If Minute(Now) = r Then rt = 0
    Next
    '--------------------------

    'Debug.Print TALKTIME$ + "__" + str(R2)

End Sub

Sub FORCESAYTIME()
    'CALL HERE FROM -- SHIFT SCROLL LOCK -- TIME
    'AND CALL GFROM PRESS 2ND DOWN O BUTTON

    Call MAKETALKTIMESTRING

    Call VolumeFixedSet(100, False)
        
    List1.AddItem Format(Now, "DD-MMM-YY HH:MM:SSa/p")

    'MAKE TIME EVENT LOG WITH TEXT BELOW
    
    FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\SAY Time Key Press Logg.txt"
    FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
    
    Fr8 = FreeFile
    Open FILENAME_IN_USE_CHECK For Append As #Fr8
        STR5 = Format(Now, "DD-MMM-YY -- HH:MM:SSa/p")
        If DOGBARKCOMMO = True Then STR5 = STR5 + " -- DOG BARK COMMANDO"
        Print #Fr8, STR5
    Close #Fr8


    If DOGBARKCOMMO = False Then
         TALKTIME$ = TALKTIME$ + " And " + str(Second(Now)) + " Seconds" '------,, OUCH"
    End If
        
    If DOGBARKCOMMO = True Then
        TALKTIME$ = TALKTIME$ + ",, Dog Bark Commando" '------",, Dog Bark Supreme Commando STAR"
    End If


    XX1 = TimerVOLCHANGETALK2.Enabled = False
    XX2 = DOGBARKCOMMO = False
    If XX1 And XX2 Then
        O_VolumeSet2 = VolumeSet2
        VolumeSet2 = VolumeSet2 + 2
        If VolumeSet2 <> O_VolumeSet2 And VolumeSet2 <= 100 Then
            Call TimerVOLCHANGETALK2_TRIG_START
        End If
        ATidalX.LblVol = str(Val(VolumeSet2))
        
    End If
    
    If TimerVOLCHANGETALK.Enabled = True Then
        TimerVOLCHANGETALK_TRIG_START
    End If
    
    DOGBARKCOMMO = False
    'Call MAKETALKTIMESTRING
    Call SpeakVoice(TALKTIME$)

    Call TIMER_SECONDS20_AFTER_HOUR_TIMER_TRIG_START

End Sub


Sub Time_Code()

'If VoiceStreamActive = True Then Exit Sub


'If Minute(Now) >= 1 And StringVoiceToPLayActive = True Then
'    MsgBox "THIS USED StringVoiceToPLayActive = False"
'    StringVoiceToPLayActive = False
'End If


SayTime2 = SayTime

YYt = Minute(Now)

If YYt = YYx Or YYx = 0 Then
    YYx = YYt
    '---------------------
    Exit Sub
    '---------------------
End If

YYx = YYt

Call MAKETALKTIMESTRING

'Call WaitWavFinish

'WE RUN EVERY SECOND HERE COZ -- MIGHT BE MANUAL TIME ASKED

'NOT ON ANY SIX
If InStr(Format(Minute(Now), "00"), "6") > 0 Then
    R2 = 0
End If

'PROBLEM INTERESTING TRY IT OUT
If 1 = 2 Then
    'If ONHOURCOUNTDOWN = True And StringVoiceToPLayActive = False Then
    If ONHOURCOUNTDOWN = True Then
        If STRING_OF_VOICE_IN_QUE_PLAY = 0 Then
            
            R2 = 1
            ONHOURCOUNTDOWN = False
            'ATidalX.Timer1.Enabled = True
            ATidalX.Timer_CountD1Text.Enabled = True
        End If
    End If

End If

'-------------
'DONT WANT TIME AT MIDNIGHT IF COUNTDOWN MAYBE
'STRING_OF_VOICE_IN_QUE_PLAY = 0
'TRY IT OUT
'-------------

'Debug.Print TALKTIME$

'If R2 = 1 And TALKTIME$ <> "" Then SayTime = True

If STRING_OF_VOICE_IN_QUE_PLAY = 0 Then

    If TALKTIME$ <> "" And R2 = 1 Then SayTime = True

End If

'SAY TIME CAN BE A MANUAL BUTTON PRESS MAYBE
If (SayTime = True Or SayTime2 = True) Then
    SayTime = False
   
    'ONLY EVERY MINUTE SO WONT BE TWICE
    If Minute(Now) Mod 15 = 0 Then
        Call TIMER_SECONDS20_AFTER_HOUR_TIMER_TRIG_START
        FLAGMOONSAYMORE = True
    End If
    
    If Soff = False Then
        'If TimeTrigSay = True Then TALKTIME$ = Format$(Now, "DDDD DD-MMMM-YYYY ") + vbCrLf + TALKTIME$
        
        
        'Call VolumeFixedSet(1, True)
        'Call VoiceMP3Pause
        
        'THIS LAST OF STRING SOUND
        'AT END AFTER COUNT DOWN SO CAN RELEASE LATCH
        
        'If StringVoiceToPLayActive = True Then
        
        '    Call VolumeFixedSet(DayOrMidNightVol, True)

        'End If

        'StringVoiceToPLayActive = False
        
        
        Call SpeakVoice("TIME #" + TALKTIME$)

        TimeTrigSay = 0
        
        'SayTime3 = False
        'DOGBARKCOMMO = False

    End If

End If




End Sub

Sub RSR232_FRONT_DOOR_LOGGER_TALK()


    Exit Sub

    Dim TT_BEFORE


    PATH_FILE_NAME1 = "C:\SCRIPTOR DATA\VB6\VB-NT\00_Best_VB_01\Tidal_Info\RS232 FRONT DOOR OPEN.txt"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(PATH_FILE_NAME1) Then
        Set F = FSO.GetFile(PATH_FILE_NAME1)
        RS232_LOGGER_FILE_TIME_OPEN = 0
        RS232_LOGGER_FILE_TIME_OPEN = F.DateLastModified
    End If

    PATH_FILE_NAME1 = "C:\SCRIPTOR DATA\VB6\VB-NT\00_Best_VB_01\Tidal_Info\RS232 FRONT DOOR CLOSE.txt"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(PATH_FILE_NAME1) Then
        Set F = FSO.GetFile(PATH_FILE_NAME1)
        RS232_LOGGER_FILE_TIME_CLOSE = 0
        RS232_LOGGER_FILE_TIME_CLOSE = F.DateLastModified
    End If

    TALKER_VAR_1 = "DOOR!"
    TALKER_VAR_2 = "DOOR!"
    
    ' --------------------------------------------------
    ' FIRST IF BOTH SAME DATE
    ' --------------------------------------------------
    RS232_LOGGER_FILE_TIME = RS232_LOGGER_FILE_TIME_OPEN
    
    If RS232_LOGGER_FILE_TIME_OPEN > RS232_LOGGER_FILE_TIME_CLOSE Then
        RS232_LOGGER_FILE_TIME = RS232_LOGGER_FILE_TIME_OPEN
        TALKER_VAR_1 = "DOOR OPEN!"
        TALKER_VAR_2 = "DOOR OPEN!"
    End If
    If RS232_LOGGER_FILE_TIME_CLOSE > RS232_LOGGER_FILE_TIME_OPEN Then
        RS232_LOGGER_FILE_TIME = RS232_LOGGER_FILE_TIME_CLOSE
        TALKER_VAR_1 = "DOOR!"
        TALKER_VAR_2 = "DOOR CLOSE!"
    End If

    RS232_LOGGER_FILE_TIME_V1 = DateDiff("s", RS232_LOGGER_FILE_TIME, Now)
    RS232_LOGGER_FILE_TIME_V2 = RS232_LOGGER_FILE_TIME_V1
    
    f6 = Int((RS232_LOGGER_FILE_TIME_V1 / 60 ^ 2))
    f7 = Int(RS232_LOGGER_FILE_TIME_V1 / 60)
    ' On Error Resume Next
    ' OVER FLOW PROBLEM WHEN NUMBER BIG
    f8 = RS232_LOGGER_FILE_TIME_V1 ' Mod 60
    On Error GoTo 0
'    Debug.Print str(F6) + " -- " + str(F7) + " -- " + str(F8)
    RS232_LOGGER_FILE_TIME_V2_HOUR = f6
'    RS232_LOGGER_FILE_TIME_V2_HOUR = DateDiff("h", RS232_LOGGER_FILE_TIME, Now)
    RS232_LOGGER_FILE_TIME_V2_MINUTE = f7
    RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD = f7 Mod 60
    RS232_LOGGER_FILE_TIME_V2_SECOND = f8
    RS232_LOGGER_FILE_TIME_V2_SECOND_MOD = f8 Mod 60
    
    If RS232_LOGGER_FILE_TIME_V2_HOUR > 0 Then
        If OLD_RS232_LOGGER_FILE_TIME_HOUR <> RS232_LOGGER_FILE_TIME_V2_HOUR Then
            ' For R = 1 To 48 * 10
                ' If RS232_LOGGER_FILE_TIME_V2_HOUR = R Then
                    TIME_STEP___1_HOUR = "HOUR"
                    TIME_TALKER_1_HOUR = RS232_LOGGER_FILE_TIME_V2_HOUR
                    SPEECH_TO_GO = True
                    ' Exit For
                ' End If
            ' Next
        End If
    End If
    
'    If RS232_LOGGER_FILE_TIME_V2_HOUR > 0 Then
'        'SET_GO_2 = False
'        ' If OLD_RS232_LOGGER_FILE_TIME_MINUTE <> RS232_LOGGER_FILE_TIME_V2_MINUTE Then SET_GO_2 = True
'        'If SET_GO_2 = True Then
'            ' ----------------------------------------------------
'            ' 1 HOUR TO 5 HOUR CHECKER
'            ' AND AFTER 1 MINUTE ALSO CHECK SECONDS IN TEN'S VALUE
'            ' CLEAN ON A MINUTE ONLY MINUTE VALUE
'            ' OR ELSE ALSO SECOND VALUE
'            ' ----------------------------------------------------
'            For R = 1 To 48 * 10
'                If RS232_LOGGER_FILE_TIME_V2_HOUR = R Then
'                    SPEECH_TO_GO_HOUR_ER = True
'                End If
'            Next
'            If SPEECH_TO_GO_HOUR_ER = True Then
'                For R = 10 To 50 Step 10
'                    If RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD = R Then
'                        TIME_STEP___1_HOUR = "HOUR"
'                        TIME_TALKER_1_HOUR = RS232_LOGGER_FILE_TIME_V2_HOUR
'                        TIME_STEP___2_MIN_ = "MINUTE"
'                        TIME_TALKER_2_MIN_ = RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD
'                        SPEECH_TO_GO = True
'                    End If
'                Next
'            End If
'        'End If
'    End If
    
    If SPEECH_TO_GO = False Then
        'If RS232_LOGGER_FILE_TIME_V2 > 59 And RS232_LOGGER_FILE_TIME_V2 < 3600 * 2 Then
        If RS232_LOGGER_FILE_TIME_V2 > 59 Then
            'If OLD_RS232_LOGGER_FILE_TIME_MINUTE <> RS232_LOGGER_FILE_TIME_V2_MINUTE Then
                
            ' If RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD = R Then
            ' 1 MINUTE TO 20 MINUTE
            For r = 1 To 20
                If RS232_LOGGER_FILE_TIME_V2_MINUTE = r Then
                    TIME_STEP___2_MIN_ = "MINUTE"
                    TIME_TALKER_2_MIN__ = RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD
                    SPEECH_TO_GO = True
                    Exit For
                End If
            Next
            ' 20 MINUTE TO 1 HOUR
            For r = 20 To 60 Step 2
                If RS232_LOGGER_FILE_TIME_V2_MINUTE = r Then
                    TIME_STEP___2_MIN_ = "MINUTE"
                    TIME_TALKER_2_MIN_ = RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD
                    SPEECH_TO_GO = True
                    Exit For
                End If
            Next
            ' 1 HOUR TO 3 HOUR
            For r = 60 To 180 Step 5
                If RS232_LOGGER_FILE_TIME_V2_MINUTE = r Then
                    TIME_STEP___2_MIN_ = "MINUTE"
                    TIME_TALKER_2_MIN_ = RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD
                    SPEECH_TO_GO = True
                    Exit For
                End If
            Next
            ' 3 HOUR TO 10 HOUR
            For r = 180 To 1200 Step 10
                If RS232_LOGGER_FILE_TIME_V2_MINUTE = r Then
                    TIME_STEP___2_MIN_ = "MINUTE"
                    TIME_TALKER_2_MIN_ = RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD
                    SPEECH_TO_GO = True
                    Exit For
                End If
            Next
            
        'End If
        End If
    End If
    
    If RS232_LOGGER_FILE_TIME_V2 >= 3600 Then
        ' SET_GO_2 = False
        ' If OLD_RS232_LOGGER_FILE_TIME_SECOND <> RS232_LOGGER_FILE_TIME_V2_SECOND Then SET_GO_2 = True
        ' If OLD_RS232_LOGGER_FILE_TIME_MINUTE = 0 Then SET_GO_2 = True
        ' If SET_GO_2 = True Then
        
            If RS232_LOGGER_FILE_TIME_V2_HOUR > 0 Then
                TIME_STEP___1_HOUR = "HOUR"
                TIME_TALKER_1_HOUR = RS232_LOGGER_FILE_TIME_V2_HOUR
            End If

        
            ' ----------------------------------------------------
            ' 1 MINUTE TO 5 MINUTE CHECK
            ' AND AFTER 1 MINUTE ALSO CHECK SECONDS IN TEN'S VALUE
            ' CLEAN ON A MINUTE ONLY MINUTE VALUE
            ' OR ELSE ALSO SECOND VALUE
            ' ----------------------------------------------------
'            For R = 1 To 3
'                If RS232_LOGGER_FILE_TIME_V2_MINUTE = R Then
'                    SPEECH_TO_GO_MINUTE_ER = True
'                End If
'            Next
'            If SPEECH_TO_GO_MINUTE_ER = True Then
                
                
                For r = 10 To 50 Step 10
                    If RS232_LOGGER_FILE_TIME_V2_SECOND = r Then
                        TIME_STEP___2_MIN_ = "MINUTE"
                        TIME_TALKER_2_MIN_ = RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD
                        TIME_STEP___3_SEC_ = "SECOND"
                        TIME_TALKER_3_SEC_ = RS232_LOGGER_FILE_TIME_V2_SECOND
                        SPEECH_TO_GO = True
                    End If
                Next
            ' End If
        ' End If
    End If
    
    
    If RS232_LOGGER_FILE_TIME_V2_SECOND > 0 And RS232_LOGGER_FILE_TIME_V2 < 60 * 60 Then
        ' If OLD_RS232_LOGGER_FILE_TIME_SECOND <> RS232_LOGGER_FILE_TIME_V2 Then
            ' 1 SECOND TO 10 SECOND
            For r = 1 To 10
                If RS232_LOGGER_FILE_TIME_V2 = r Then
                    TIME_STEP___3_SEC_ = "SECOND"
                    TIME_TALKER_3_SEC_ = RS232_LOGGER_FILE_TIME_V2_SECOND
                    SPEECH_TO_GO = True
                    
                    SECOND_TALKER = True

                    Exit For
                End If
            Next
            ' 10 SECOND TO 20 SECOND
            For r = 10 To 20 Step 2
                If RS232_LOGGER_FILE_TIME_V2 = r Then
                    TIME_STEP___3_SEC_ = "SECOND"
                    TIME_TALKER_3_SEC_ = RS232_LOGGER_FILE_TIME_V2_SECOND
                    SPEECH_TO_GO = True
                    
                    SECOND_TALKER = True

                    Exit For
                End If
            Next
            ' 20 SECOND TO 1 MINUTE
            For r = 20 To 60 Step 5
                If RS232_LOGGER_FILE_TIME_V2 = r Then
                    TIME_STEP___3_SEC_ = "SECOND"
                    TIME_TALKER_3_SEC_ = RS232_LOGGER_FILE_TIME_V2_SECOND
                    SPEECH_TO_GO = True
                    
                                    SECOND_TALKER = True

                    Exit For
                End If
            Next
            '
'            For R = 58 To 59
'                If RS232_LOGGER_FILE_TIME_V2_SECOND = R Then
'                    TIME_STEP___3_SEC_ = "SECOND"
'                    TIME_TALKER_3_SEC_ = RS232_LOGGER_FILE_TIME_V2_SECOND
'                    SPEECH_TO_GO = True
'                    Exit For
'                End If
'            Next
            ' 1 MINUTE TO 3 MINUTE
            For r = 60 To 60 * 3 Step 10
                If RS232_LOGGER_FILE_TIME_V2 = r Then
                    TIME_STEP___2_MIN_ = "MINUTE"
                    TIME_TALKER_2_MIN_ = RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD
                    TIME_STEP___3_SEC_ = "SECOND"
                    TIME_TALKER_3_SEC_ = RS232_LOGGER_FILE_TIME_V2_SECOND_MOD
                    SPEECH_TO_GO = True
                    
                    SECOND_TALKER = True

                    Exit For
                End If
            Next
            '3 MINUTE TO 5 MINUTE
            For r = 60 * 3 To 60 * 5 Step 20
                If RS232_LOGGER_FILE_TIME_V2 = r Then
                    TIME_STEP___2_MIN_ = "MINUTE"
                    TIME_TALKER_2_MIN_ = RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD
                    TIME_STEP___3_SEC_ = "SECOND"
                    TIME_TALKER_3_SEC_ = RS232_LOGGER_FILE_TIME_V2_SECOND_MOD
                    SPEECH_TO_GO = True
                    
                    SECOND_TALKER = True

                    Exit For
                End If
            Next
            '5 MINUTE TO 10 MINUTE
            For r = 60 * 5 To 60 * 10 Step 20
                If RS232_LOGGER_FILE_TIME_V2 = r Then
                TIME_STEP___2_MIN_ = "MINUTE"
                TIME_TALKER_2_MIN_ = RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD
                TIME_STEP___3_SEC_ = "SECOND"
                TIME_TALKER_3_SEC_ = RS232_LOGGER_FILE_TIME_V2_SECOND_MOD
                SPEECH_TO_GO = True
                SECOND_TALKER = True
                Exit For
                End If
            Next
            '10 MINUTE TO 20 MINUTE
            For r = 60 * 10 To 60 * 20 Step 30
                If RS232_LOGGER_FILE_TIME_V2 = r Then
                TIME_STEP___2_MIN_ = "MINUTE"
                TIME_TALKER_2_MIN_ = RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD
                TIME_STEP___3_SEC_ = "SECOND"
                TIME_TALKER_3_SEC_ = RS232_LOGGER_FILE_TIME_V2_SECOND_MOD
                SPEECH_TO_GO = True
                SECOND_TALKER = True
                Exit For
                End If
            Next
            
            ' CLEAN MINUTE NONE SECOND
            For r = 60 To 60 * 10 Step 60
                If RS232_LOGGER_FILE_TIME_V2 = r Then
                    TIME_STEP___3_SEC_ = ""
                    TIME_TALKER_3_SEC_ = 0
                    SECOND_TALKER = False
                    Exit For
                End If
            Next
            ' CLEAR SECOND THAT TALK 60
            If Val(TIME_TALKER_3_SEC_) >= 60 Then
                    TIME_STEP___3_SEC_ = ""
                    TIME_TALKER_3_SEC_ = 0
                    SECOND_TALKER = False
            End If
            
        ' End If
    End If
    
    ' 10 MINUTE DOOR TALK REPEAT A FEW - FEW SECOND BEFORE AND AFTER
    ' --------------------------------
    For R_XY = 10 To 180 Step 10
        If R_XY = 10 Then R_XY_2 = 2 Else R_XY_2 = 2
        If RS232_LOGGER_FILE_TIME_V2 > 60 * R_XY - 1 And RS232_LOGGER_FILE_TIME_V2 < 60 * R_XY + R_XY_2 Then
            ' If OLD_RS232_LOGGER_FILE_TIME_SECOND <> RS232_LOGGER_FILE_TIME_V2 Then
            ' If RS232_LOGGER_FILE_TIME_V2 = R Then
                TIME_STEP___2_MIN_ = "MINUTE"
                TIME_TALKER_2_MIN_ = RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD
                TIME_STEP___3_SEC_ = "SECOND"
                TIME_TALKER_3_SEC_ = RS232_LOGGER_FILE_TIME_V2_SECOND_MOD
                SPEECH_TO_GO = True
                SECOND_TALKER = True
                
                
            'End If
            
            
            
                        ' CLEAR SECOND THAT TALK 60
            If Val(TIME_TALKER_3_SEC_) >= 60 Then
                    TIME_STEP___3_SEC_ = ""
                    TIME_TALKER_3_SEC_ = 0
                    SECOND_TALKER = False
            End If

            
            
        End If
    Next
    
    If RS232_LOGGER_FILE_TIME_V2_SECOND > 0 And RS232_LOGGER_FILE_TIME_V2 > 120 Then
        If SECOND_TALKER = False Then
            TIME_STEP___3_SEC_ = ""
            TIME_TALKER_3_SEC_ = 0
        End If
    End If
    
    If OLD_RS232_LOGGER_FILE_TIME_SECOND = -1 Or START_DOOR_TALKER = True Then
        If RS232_LOGGER_FILE_TIME_V2_HOUR > 0 Then
            TIME_STEP___1_HOUR = "HOUR"
            TIME_TALKER_1_HOUR = RS232_LOGGER_FILE_TIME_V2_HOUR
        End If
        If RS232_LOGGER_FILE_TIME_V2_MINUTE > 0 Then
            TIME_STEP___2_MIN = "MINUTE"
            TIME_TALKER_2_MIN_ = RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD
        End If
        If RS232_LOGGER_FILE_TIME_V2_SECOND > 0 Then
            TIME_STEP___3_SEC = "SECOND"
            TIME_TALKER_3_SEC_ = RS232_LOGGER_FILE_TIME_V2_SECOND_MOD
        End If
        
        If RS232_LOGGER_FILE_TIME_V2_SECOND > 0 And RS232_LOGGER_FILE_TIME_V2 > 60 * 6 Then
            TIME_STEP___3_SEC_ = ""
            TIME_TALKER_3_SEC_ = 0
        End If
   
        SPEECH_TO_GO = True
    End If
    
    ' If SPEECH_TO_GO = True Then Stop
    
    TIME_TALKER_1_HOUR_2 = TIME_TALKER_1_HOUR
    TIME_TALKER_2_MIN__2 = TIME_TALKER_2_MIN_
    TIME_TALKER_3_SEC__2 = TIME_TALKER_3_SEC_

    TIME_TALKER_1_HOUR = Trim(str(TIME_TALKER_1_HOUR))
    TIME_TALKER_2_MIN_ = Trim(str(TIME_TALKER_2_MIN_))
    TIME_TALKER_3_SEC_ = Trim(str(TIME_TALKER_3_SEC_))
    
    Dim ARR_HHMMSS(200)
    ARR_HHMMSS(6) = " SIXER"
    ARR_HHMMSS(16) = " SIXTHTEENTHAR"
    ARR_HHMMSS(17) = " SEVENTEENIETH"
    ARR_HHMMSS(18) = " EIGHTEENTIETH"
    ARR_HHMMSS(19) = " NINETEENIETH"
    ARR_HHMMSS(20) = " TWENTIETH"
    ARR_HHMMSS(30) = " THIRTIETH"
    ARR_HHMMSS(40) = " FORTIETH"
    ARR_HHMMSS(50) = " FIFTIETH"
    ARR_HHMMSS(60) = " SIXTIETH"
    ARR_HHMMSS(70) = " SEVENTIETH"
    ARR_HHMMSS(80) = " EIGHTIETH"
    ARR_HHMMSS(90) = " NINETIETH"
    ARR_HHMMSS(26) = " TWENTIETH SIXER"
    ARR_HHMMSS(36) = " THIRTIETH SIXER"
    ARR_HHMMSS(46) = " FORTIETH SIXER"
    ARR_HHMMSS(56) = " FIFTIETH SIXER"
    ARR_HHMMSS(66) = " SIXTIETH SIXER"
    ARR_HHMMSS(76) = " SEVENTIETH SIXER"
    ARR_HHMMSS(86) = " EIGHTIETH SIXER"
    ARR_HHMMSS(96) = " NINETIETH SIXER"
    ARR_HHMMSS(100) = " ONE HUNDREDTH"
    
    If Val(TIME_TALKER_1_HOUR) >= 100 And Val(TIME_TALKER_1_HOUR) < 999 Then
        TIME_TALKER_1_HOUR = ARR_HHMMSS(100) + " AND " + Mid(TIME_TALKER_1_HOUR, 2)
    End If
    
    
    If IsNumeric(TIME_TALKER_1_HOUR) Then RR = Trim(str(TIME_TALKER_1_HOUR))
    For r = 100 To 1 Step -1
        If ARR_HHMMSS(r) <> "" Then
            If InStr(Trim(str(RR)), Trim(str(r))) > 0 Then
                TIME_TALKER_1_HOUR = Replace(RR, RR, ARR_HHMMSS(r))
                Exit For
            End If
        End If
    Next
    RR = Trim(str(TIME_TALKER_2_MIN_))
    For r = 100 To 1 Step -1
        If ARR_HHMMSS(r) <> "" Then
            If InStr(Trim(str(RR)), Trim(str(r))) > 0 Then
                TIME_TALKER_2_MIN_ = Replace(RR, RR, ARR_HHMMSS(r))
                Exit For
            End If
        End If
    Next
    RR = Trim(str(TIME_TALKER_3_SEC_))
    For r = 100 To 1 Step -1
        If ARR_HHMMSS(r) <> "" Then
            If InStr(Trim(str(RR)), Trim(str(r))) > 0 Then
                TIME_TALKER_3_SEC_ = Replace(RR, RR, ARR_HHMMSS(r))
                Exit For
            End If
        End If
    Next
    
    If RS232_LOGGER_FILE_TIME_V2 = 0 Then
        Call SpeakVoice(TALKER_VAR_2)
    End If
    
    If SPEECH_TO_GO = True Then
    If TALKER_VAR_1 <> "" Then
        
        
        If RS232_LOGGER_FILE_TIME_V2 < 10 Then
            If RS232_LOGGER_FILE_TIME_V2 > 2 Then
                TALKER_VAR_2 = Replace(TALKER_VAR_2, " CLOSE!", "MAN!")
            End If
            If RS232_LOGGER_FILE_TIME_V2 > 5 Then
                TALKER_VAR_2 = Replace(TALKER_VAR_2, " OPEN!", "MAN!")
            End If
        End If
        
        If RS232_LOGGER_FILE_TIME_V2 > 2 Then
            TALKER_VAR_2 = Replace(TALKER_VAR_2, "CLOSE!", "")
        End If
        If RS232_LOGGER_FILE_TIME_V2 > 5 Then
            TALKER_VAR_2 = Replace(TALKER_VAR_2, "OPEN!", "")
        End If
        

    End If
    End If
    
    If RS232_LOGGER_FILE_TIME_V2 < 10 Then
        RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_1_OLD = 0
        RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_2 = 0
    End If
    
    ' ------------------------------------------------
    ' CHAIN DOG GUARD AS FALL BACK HAPPEN _WORK AROUND
    ' DON'T WANT BACKWARD TIME AFTER GIVEN
    ' ------------------------------------------------
'    TIME_TALKER_1_HOUR_2 = TIME_TALKER_1_HOUR
'    TIME_TALKER_2_MIN__2 = TIME_TALKER_2_MIN_
'    TIME_TALKER_3_SEC__2 = TIME_TALKER_3_SEC_
    TIME_TALKER_1_HOUR_3 = TIME_TALKER_1_HOUR_2
    TIME_TALKER_2_MIN__3 = TIME_TALKER_2_MIN__2
    TIME_TALKER_3_SEC__3 = TIME_TALKER_3_SEC__2
    RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_1 = TimeSerial(Val(TIME_TALKER_1_HOUR_2), Val(TIME_TALKER_2_MIN__2), Val(TIME_TALKER_3_SEC__2))
    RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_2 = TimeSerial(Val(TIME_TALKER_1_HOUR_3), Val(TIME_TALKER_2_MIN__3), Val(TIME_TALKER_3_SEC__3))
    If RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_1 < RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_2 Then
        SPEECH_TO_GO = False
    End If
    
    ' BOOT IN
    If OLD_RS232_LOGGER_FILE_TIME_SECOND = -1 Or START_DOOR_TALKER = True Then
        SPEECH_TO_GO = True
    End If
    
    If RS232_LOGGER_FILE_TIME_V1 < RS232_LOGGER_FILE_TIME_V1_OLD Then
        RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_1_OLD = 0
    End If
        
    If OLD_RS232_LOGGER_FILE_TIME_SECOND = -1 Then
    If OLD_RS232_LOGGER_FILE_TIME_SECOND = -1 Then
    If OLD_RS232_LOGGER_FILE_TIME_MINUTE = -1 Then
        OLD_RS232_LOGGER_FILE_TIME_CLOSE = RS232_LOGGER_FILE_TIME_CLOSE
        ' RS232_LOGGER_FILE_TIME_V1 = DateDiff("s", RS232_LOGGER_FILE_TIME, Now)
        
        On Error Resume Next
        f6 = Int((RS232_LOGGER_FILE_TIME_V1 / 60 ^ 2))
        f7 = Int(RS232_LOGGER_FILE_TIME_V1 / 60)
        ' On Error Resume Next
        ' OVER FLOW PROBLEM WHEN NUMBER BIG
        f8 = RS232_LOGGER_FILE_TIME_V1 ' Mod 60
        On Error GoTo 0
        OLD_RS232_LOGGER_FILE_TIME_HOUR = f6
        OLD_RS232_LOGGER_FILE_TIME_MINUTE = f7
        OLD_RS232_LOGGER_FILE_TIME_SECOND = f8
    End If
    End If
    End If
        
        
    If RS232_LOGGER_FILE_TIME_V1 < RS232_LOGGER_FILE_TIME_V1_OLD Or START_DOOR_TALKER = True Then
        If InStr(TALKER_VAR_2, "OPEN") > 0 Or START_DOOR_TALKER = True Then
            DOOR_OPEN_ANNOUNCE_HOW_LONG_BEEN = True
            
            TT_BEFORE = "DOOR TIME! "
            If OLD_RS232_LOGGER_FILE_TIME_HOUR > 0 Then
                i = OLD_RS232_LOGGER_FILE_TIME_HOUR
                If i < 2 Then HANDSTR = " Hour " Else HANDSTR = " Hours "
                TT_BEFORE = TT_BEFORE + Trim(str(i)) + HANDSTR
            End If
            If OLD_RS232_LOGGER_FILE_TIME_MINUTE > 0 Then
                i = OLD_RS232_LOGGER_FILE_TIME_MINUTE Mod 60
                If i < 2 Then HANDSTR = " Minute " Else HANDSTR = " Minutes "
                TT_BEFORE = TT_BEFORE + Trim(str(i)) + HANDSTR
            Else
                If OLD_RS232_LOGGER_FILE_TIME_HOUR > 0 Then
                    TT_BEFORE = TT_BEFORE + " And "
                End If
            End If
            If OLD_RS232_LOGGER_FILE_TIME_SECOND > 0 Then
                i = OLD_RS232_LOGGER_FILE_TIME_SECOND Mod 60
                If i < 2 Then HANDSTR = " Second " Else HANDSTR = " Seconds "
                TT_BEFORE = TT_BEFORE + Trim(str(i)) + HANDSTR
            End If
            
            TT_BEFORE = Replace(TT_BEFORE, "  ", " ")
            
            TT_BEFORE_AT = Format(OLD_RS232_LOGGER_FILE_TIME_CLOSE, "DDDD HH:MM AM/PM")
            TT_BEFORE = Trim(TT_BEFORE) + ", At! " + TT_BEFORE_AT
            
            TT_BEFORE = Replace(TT_BEFORE, "  ", " ")
            
        End If
        
    End If
    RS232_LOGGER_FILE_TIME_V1_OLD = RS232_LOGGER_FILE_TIME_V1
    
    If RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_1_OLD > RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_1 Then
        SPEECH_TO_GO = False
    End If
    If RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_1 > RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_1_OLD Then
        RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_1_OLD = RS232_LOGGER_FILE_TIME_V2_TIME_SERIAL_1
    End If
    
    Call Sound_REtrigger_MMCONTROL9("RETRIGGER")
    
    
    If SPEECH_TO_GO = True Then
        If RS232_LOGGER_FILE_TIME_V2 > 0 Then
        
            ' REPROCESSOR AS COPY ONTO OLD
            ' ----------------------------
            SpeakVoice_STRING = ""
            SpeakVoice_STRING = Trim(SpeakVoice_STRING + TIME_TALKER_1_HOUR) + " HOUR "
            SpeakVoice_STRING = Trim(SpeakVoice_STRING + TIME_TALKER_2_MIN_) + " MINUTE "
            SpeakVoice_STRING = Trim(SpeakVoice_STRING + TIME_TALKER_3_SEC_) + " SECOND "
            
            If TIME_TALKER_1_HOUR = 0 Then
                SpeakVoice_STRING = Trim(Replace(SpeakVoice_STRING, Trim(str(TIME_TALKER_1_HOUR)) + " HOUR", ""))
            End If
            If TIME_TALKER_2_MIN_ = 0 Then
                SpeakVoice_STRING = Trim(Replace(SpeakVoice_STRING, Trim(str(TIME_TALKER_2_MIN_)) + " MINUTE", ""))
            End If
            If TIME_TALKER_3_SEC_ = 0 Then
                SpeakVoice_STRING = Trim(Replace(SpeakVoice_STRING, Trim(str(TIME_TALKER_3_SEC_)) + " SECOND", ""))
            End If
            
            ' SpeakVoice_STRING = Trim(Replace(SpeakVoice_STRING, " SECOND", "SEC"))
            
            If TIME_TALKER_1_HOUR = 0 Then
                SpeakVoice_STRING = Trim(Replace(SpeakVoice_STRING, " SECOND", ""))
            End If
            
            ' 60 * 10 TEN MINUTE
            ' ------------------
            If MNU_NOT_DOOR_TALK.Checked = False And RS232_LOGGER_FILE_TIME_V1 < 60 * 10 Then
                SpeakVoice_STRING = SpeakVoice_STRING + " " + Trim(TALKER_VAR_2)
            End If
        
            If OLD_SPEAK_DOOR <> SpeakVoice_STRING Or OLD_RS232_LOGGER_FILE_TIME_SECOND < 0 Or START_DOOR_TALKER = True Then
                START_DOOR_TALKER = False
                SpeakVoice_STRING_HOLD = SpeakVoice_STRING
                ' -------------------------------------------------------
                ' GET SOME REPAT IF ANOTHER VOICE TALKER
                ' TEST OUT IF VARIABLE NOT PASS BACK THROUGH AS PARAMETER
                ' -------------------------------------------------------
                Call Sound_REtrigger_MMCONTROL9("SPEECH")

                If SET_MMControl9_TO_GO_FLAG = True And Trim(SpeakVoice_STRING) <> "" Then
                    ATidalX.MMControl9.Command = "prev"
                    ATidalX.MMControl9.Command = "Play"
                    
                    Debug.Print "BLIPPER_NEXT_TO_VOICE"

                End If
                SET_MMControl9_TO_GO_FLAG = False
                If TT_BEFORE <> "" Then
                    Call SpeakVoice(TT_BEFORE)
                    TT_BEFORE = ""
                End If
                
                Call SpeakVoice(SpeakVoice_STRING)

                OLD_SPEAK_DOOR = SpeakVoice_STRING_HOLD
            End If
        End If
    End If
    ' SHUT THAT DOOR
    
    OLD_RS232_LOGGER_FILE_TIME_SECOND = RS232_LOGGER_FILE_TIME_V2_SECOND
    OLD_RS232_LOGGER_FILE_TIME_MINUTE = RS232_LOGGER_FILE_TIME_V2_MINUTE
    OLD_RS232_LOGGER_FILE_TIME_HOUR = RS232_LOGGER_FILE_TIME_V2_HOUR
    OLD_RS232_LOGGER_FILE_TIME_CLOSE = RS232_LOGGER_FILE_TIME_CLOSE
End Sub

Sub Sound_REtrigger_MMCONTROL9(VALUE_ON)

    If VALUE_ON = "RETRIGGER" Then

        If RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD_RETRIG_SOUND > 1 Then
            RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD_RETRIG_SOUND = RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD_RETRIG_SOUND - 1
            
            'If RS232_LOGGER_FILE_TIME_V2_HOUR < 3 Then
                ATidalX.MMControl9.Command = "prev"
                ATidalX.MMControl9.Command = "Play"
                Debug.Print "BLIPPER_RETRIGGER"

                
            'Else
            '    SET_MMControl9_TO_GO_FLAG = True
            'End If
        End If
    End If
    
    If VALUE_ON = "SPEECH" Then
        If OLD_RS232_LOGGER_FILE_TIME_V2_MINUTE = -2 Then
            OLD_RS232_LOGGER_FILE_TIME_V2_MINUTE = RS232_LOGGER_FILE_TIME_V2_MINUTE
        End If
        
        If OLD_RS232_LOGGER_FILE_TIME_V2_MINUTE <> RS232_LOGGER_FILE_TIME_V2_MINUTE Then
            ' Debug.Print Now
            OLD_RS232_LOGGER_FILE_TIME_V2_MINUTE = RS232_LOGGER_FILE_TIME_V2_MINUTE
            
            RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD_RETRIG_SOUND = 1
            If RS232_LOGGER_FILE_TIME_V2_MINUTE = 10 Then
                RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD_RETRIG_SOUND = 3
            End If
            
            If RS232_LOGGER_FILE_TIME_V2_MINUTE Mod 10 = 0 Then
                RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD_RETRIG_SOUND = 2
            End If
            
            If RS232_LOGGER_FILE_TIME_V2_MINUTE < 10 Then
                RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD_RETRIG_SOUND = 1
            End If
            
            
            If RS232_LOGGER_FILE_TIME_V2_SECOND > 20 And RS232_LOGGER_FILE_TIME_V2 < 60 * 9 + 20 Then
                SET_MMControl9_TO_GO_FLAG = False
                
                RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD_RETRIG_SOUND = 0
            End If
            If RS232_LOGGER_FILE_TIME_V2_HOUR >= 3 Then
                SET_MMControl9_TO_GO_FLAG = False
                
                RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD_RETRIG_SOUND = 0
            End If
            
            If RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD_RETRIG_SOUND > 0 Then
                If RS232_LOGGER_FILE_TIME_V2_HOUR < 3 Then
                    ATidalX.MMControl9.Command = "prev"
                    ATidalX.MMControl9.Command = "Play"
                    
                    Debug.Print "BLIPPER"
                    RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD_RETRIG_SOUND_2 = True
                Else
                    SET_MMControl9_TO_GO_FLAG = True
                    ' RS232_LOGGER_FILE_TIME_V2_MINUTE_MOD_RETRIG_SOUND = 1
                End If
            End If
            A = A
        End If
        
    End If

End Sub


Sub DateMidnight_CountDown()
Dim YYSP2
'---------------------
'MIDNIGHT AND HOUR ONE
'---------------------

If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then
    
    If HOUR(Now) <> HOUR(STRING_OF_VOICE_IN_QUE_PLAY) Then
        XTAG = 0
        For RI1 = 0 To ListVOICE.ListCount - 1
            If InStr(ListVOICE.List(0), "TIME CD#") > 0 Then
                XTAG = 1
            End If
        Next
        If XTAG = 0 Then
            STRING_OF_VOICE_IN_QUE_PLAY = 0
        End If
    End If
End If



Dim Time10 As Date

Time10 = Now + speedtime
'Time10 = Int(Now) + TimeSerial(20, 59, 57)

EGG = 5
'TSSP$ = ""


XXV = HOUR(Now) + 1
If XXV > 12 Then XXV = XXV - 12
EGG = XXV
If EGG = 12 Then EGG = 20
If EGG = 6 Then EGG = 8
If EGG = 3 Then EGG = 3 + 2

If EGG = 1 Then EGG = 8
'If EGG = 2 Then EGG = 8




If Mid$(Date$, 1, 5) = "12-31" And HOUR(Time10) >= 18 Then EGG = 20
If Mid$(Date$, 1, 5) = "01-01" And HOUR(Time10) <= 5 Then EGG = 20



'-------------------
'MIDNIGHTS
'OF SPECIAL DAYS
'-------------------

t = ",, 2 THOUSAND AND 10"
If HOUR(Time10) = 23 Then
    'If 1 = 1 Then EGG = 20:
    TSSP$ = "EVERY DAY"
    
    If Mid$(Date$, 1, 5) = "12-31" Then
        EGG = 300 ': TSSP$ = "New Years Day,, Bang,, Bang,, Bang,," '9 MINS
        'TSSP$ = "DONG DONG ,, New Years Day OF 2 THOUSAND AND 10 ,, Bang,, Bang,, Bang,,HA HA HA"
        
        't = ",, 2 THOUSAND AND 10"
        TSR = ""
        For r = 1 To 10
            TSR = TSR + t
        Next
        TSR = TSR + ",, 10 TIMES"
        TSSP$ = "12 O'CLOCK ,, New Years Day," + TSR + ",, The Time is Free of Charge, TODAY" ',, Dirty Old bastard,, Tooth Hurty"
    End If
    
    If Mid$(Date$, 1, 5) = "12-30" Then EGG = 10 * 60: TSSP$ = "New Year's DAY,, FOR " + t + " FOR " + t + " FOR " + t
    If Mid$(Date$, 1, 5) = "12-30" Then EGG = 4 * 60: TSSP$ = "New Year's EVE,, FOR " + t + " FOR " + t + " FOR " + t
    
    If Mid$(Date$, 1, 5) = "12-23" Then EGG = 120: TSSP$ = "Chirstmas Day"
    
    If Mid$(Date$, 1, 5) = "12-24" Then EGG = 120: TSSP$ = "Chirstmas Eve"
    If Mid$(Date$, 1, 5) = "12-25" Then EGG = 120: TSSP$ = "Boxing Day"
    
    If Mid$(Date$, 1, 5) = "11-19" Then EGG = 120: TSSP$ = "My Birthday"
    If Mid$(Date$, 1, 5) = "09-09" Then EGG = 120: TSSP$ = "Andy's Birthday"
    If Mid$(Date$, 1, 5) = "09-08" Then EGG = 120: TSSP$ = "Edd's Birthday"
    If Mid$(Date$, 1, 5) = "12-04" Then EGG = 120: TSSP$ = "Dad's Birthday"
    If Mid$(Date$, 1, 5) = "08-15" Then EGG = 120: TSSP$ = "Mum's Birthday"
    If Mid$(Date$, 1, 5) = "05-18" Then EGG = 120: TSSP$ = "Grandma's Birthday"
End If

hd = 1



H2O = Time$

If H321 <> H2O Then
    'Debug.Print Time$ + "   FALSE"
    '-----------
    'StringVoiceToPLayActive = False
    '-----------
    'THIS SHOULD BE COOL BIT HARD DEBUG UNLESS USE ANOTHER VAR
    
End If


If OLD_SECOND_NOW = Second(Now) Then Exit Sub
OLD_SECOND_NOW = Second(Now)


Dim WNow As Date




'----------
'TENMINUTE-ER TEST ONE
'----------
TO_GO = False
'EGG = 240
TEST_TEN_MINUTE = True
For RS200 = 5 To 9
    For RS201 = 0 To 50 Step 10
        If H321 <> H2O And Minute(Time10) = RS200 + RS201 And Second(Time10) >= 60 - EGG Then TO_GO = True
    Next
Next

TO_GO = True
If XZ_TEST_VOICE_COUNTDOWN <= -5 Then TO_GO = False

TO_GO = False

If TEST_TEN_MINUTE = True And TO_GO = True Then


    WNow = Now - TimeSerial(0, 0, 1)
    
    YYSP = Format(TimeSerial(23, 59, 59) - (WNow - (Int(WNow))), "hh:mm:SS")
    YYSP = Mid(YYSP, 5)
    
    
 '   Debug.Print YYSP
'    If Mid(YYSP, 3, 2) = "01" Then Stop
    'If Mid(YYSP, 3, 2) = "00" Then Stop
    
    If Val(str(Val(Mid(YYSP, 1, 1)))) = 0 Then
        'yysp2 = str(Val(Mid(yysp, 4, 2))) '+ " sec"
        SET_VAR = True
    End If
    If SET_VAR = False Then
        If Val(str(Val(Mid(YYSP, 1, 1)))) = 1 Then
            YYSP2 = str(Val(Mid(YYSP, 1, 2))) + " MINUTE "
        SET_VAR = True
        End If
    End If
    If SET_VAR = False Then
        If Val(str(Val(Mid(YYSP, 1, 1)))) > 1 Then
            YYSP2 = str(Val(Mid(YYSP, 1, 2))) + " MINUTES "
        End If
    End If
     
    If Val(str(Val(Mid(YYSP, 3, 2)))) >= 1 Then
        YYSP2 = YYSP2 + str(Val(Mid(YYSP, 3, 2)))
    End If
    
    If Val(str(Val(Mid(YYSP, 4, 1)))) = 4 Then
        YYSP2 = YYSP2 + ", HARD CORE!"
    End If
    If Val(str(Val(Mid(YYSP, 4, 1)))) = 5 Then
         YYSP2 = YYSP2 + ", FIVE!"
    End If
    
    If Val(str(Val(Mid(YYSP, 1, 1)))) = 0 And Val(str(Val(Mid(YYSP, 3, 2)))) = 0 Then
        YYSP2 = YYSP2 + ", ALL!, DONE!, TOTAL!, BACKWARDS!, TEN MINUTES!, COUNTDOWN! TIME!, HIP, HIP, HOARY!, HAMMER!, HAMMER! HAMMER!, HAMMER!, HARD CORE!, HARD CORE!, HARD CORE!, HARD CORE!"
        XZ_TEST_VOICE_COUNTDOWN = XZ_TEST_VOICE_COUNTDOWN - 1
    End If
  '  Debug.Print YYSP
    If Val(str(Val(Mid(YYSP, 1, 1)))) <> 0 And Val(str(Val(Mid(YYSP, 3, 2)))) = 0 Then
        YYSP2 = YYSP2 + ", TOTAL!, BACKWARDS!, TEN MINUTES!, COUNTDOWN!, HARD CORE!"
    End If
    
    
    If Val(str(Val(Mid(YYSP, 4, 1)))) = 0 And Val(str(Val(Mid(YYSP, 3, 2)))) > 0 Then  'And Val(str(Val(Mid(yysp, 1, 2)))) = 0 Then
        YYSP2 = YYSP2 + ", TIME!"
    End If
    
     'And Val(str(Val(Mid(yysp, 4, 2)))) > 1
    TEST_TRIM = True
    If XZ_TEST_VOICE_COUNTDOWN = -1 Then
    TEST_TRIM = True
    End If
    If XZ_TEST_VOICE_COUNTDOWN = -2 Then
'        YYSP2 = Replace(YYSP2, ", ALL!, DONE!, TOTAL!, BACKWARDS!, TEN MINUTES!, COUNTDOWN! TIME!,, HIP, HIP, HORAY!, HAMMER!, HAMMER! HAMMER!, HAMMER!, HARD CORE!, HARD CORE!, HARD CORE!, HARD CORE!", "")
        YYSP2 = Replace(YYSP2, ", TOTAL!, BACKWARDS!, TEN MINUTES!, COUNTDOWN!, HARD CORE!", "")
        If InStr(YYSP2, ", ALL!, DONE!, TOTAL!, BACKWARDS!, TEN MINUTES!, COUNTDOWN! TIME!, HIP, HIP, HOARY!, HAMMER!, HAMMER! HAMMER!, HAMMER!, HARD CORE!, HARD CORE!, HARD CORE!, HARD CORE!") = 0 Then
            YYSP2 = Replace(YYSP2, ", HARD CORE!", " HARD CORE")
            YYSP2 = Replace(YYSP2, ", COUNTDOWN! HARD CORE", ", COUNTDOWN!, HARD CORE!")
            YYSP2 = Replace(YYSP2, ", TIME!", " TIME")
            YYSP2 = Replace(YYSP2, ", FIVE!", " FIVE")
        End If
    End If
    If XZ_TEST_VOICE_COUNTDOWN = -3 Then
        YYSP2 = Replace(YYSP2, ", TOTAL!, BACKWARDS!, TEN MINUTES!, COUNTDOWN!, HARD CORE!", "")
        If InStr(YYSP2, ", ALL!, DONE!, TOTAL!, BACKWARDS!, TEN MINUTES!, COUNTDOWN! TIME!, HIP, HIP, HOARY!, HAMMER!, HAMMER! HAMMER!, HAMMER!, HARD CORE!, HARD CORE!, HARD CORE!, HARD CORE!") = 0 Then
            YYSP2 = Replace(YYSP2, ", HARD CORE!", " HARD CORE")
            YYSP2 = Replace(YYSP2, ", TIME!", " TIME")
            YYSP2 = Replace(YYSP2, ", FIVE!", " FIVE")
            YYSP2 = Replace(YYSP2, " TIME", "")
            YYSP2 = Replace(YYSP2, " FIVE", "")
            YYSP2 = Replace(YYSP2, " MINUTE ", "")
            YYSP2 = Replace(YYSP2, " MINUTES ", "")
        End If
    End If
    If XZ_TEST_VOICE_COUNTDOWN = -4 Or XZ_TEST_VOICE_COUNTDOWN = 0 Then
'        YYSP2 = Replace(YYSP2, ", ALL!, DONE!, TOTAL!, BACKWARDS!, TEN MINUTES!, COUNTDOWN! TIME!,, HIP, HIP, HORAY!, HAMMER!, HAMMER! HAMMER!, HAMMER!, HARD CORE!, HARD CORE!, HARD CORE!, HARD CORE!", "")
        YYSP2 = Replace(YYSP2, ", TOTAL!, BACKWARDS!, TEN MINUTES!, COUNTDOWN!, HARD CORE!", "")
        If InStr(YYSP2, ", ALL!, DONE!, TOTAL!, BACKWARDS!, TEN MINUTES!, COUNTDOWN! TIME!, HIP, HIP, HOARY!, HAMMER!, HAMMER! HAMMER!, HAMMER!, HARD CORE!, HARD CORE!, HARD CORE!, HARD CORE!") = 0 Then
            YYSP2 = Replace(YYSP2, ", HARD CORE!", " HARD CORE")
            YYSP2 = Replace(YYSP2, " HARD CORE", "")
            YYSP2 = Replace(YYSP2, ", TIME!", " TIME")
            YYSP2 = Replace(YYSP2, ", FIVE!", " FIVE")
            YYSP2 = Replace(YYSP2, " TIME", "")
            YYSP2 = Replace(YYSP2, " FIVE", "")
            YYSP2 = Replace(YYSP2, " MINUTE ", "")
            YYSP2 = Replace(YYSP2, " MINUTES ", "")
        End If
    End If
    
    
    If STRING_OF_VOICE_IN_QUE_PLAY = 0 Then
        STRING_OF_VOICE_IN_QUE_PLAY = Now
    End If
    
    If OLD_GETTING_REPEAT_IF_ANOTHER_SPEAK <> YYSP2 Then
        YYSP2 = Replace(YYSP2, "  ", " ")
        YYSP2 = Trim(YYSP2)
        Call SpeakVoice(YYSP2)
    End If
    OLD_GETTING_REPEAT_IF_ANOTHER_SPEAK = YYSP

    Exit Sub

End If

'--------------------------
'THIS MIDNIGHT ONE
'--------------------------
TO_GO = False
If H321 <> Time$ And Now - Int(Now) > (TimeSerial(23, 59, 59) - TimeSerial(0, 0, EGG - 1)) Then TO_GO = True
If TO_GO = True Then
    
    H321 = Time$

    If Soff = False Then
        TSSP$ = "EVERY DAY"
        
    
   
        
        
        'If Minute(Time10) = EGG Then
        '    If TSSP$ <> "" Then Call SpeakVoice( TSSP$)
        'End If
        
        xxhe = Int(Now) + TimeSerial(HOUR(Time10) + 1, 0, 0)
        'WNow = DateValue(Now) + TimeSerial(23, 59, 59) - TimeSerial(0, 0, 1)
        WNow = Now - TimeSerial(0, 0, 1)
        
        YYSP = Format(TimeSerial(23, 59, 59) - (WNow - (Int(WNow))), "hh:mm:SS")
        YYSP = Mid(YYSP, 4)
        'YYSP = str(Val(Mid(YYSP, 1, 2))) + " MIN " + str(Val(Mid(YYSP, 4, 2))) + " sec TO 12 O clock"
        'YYSP = str(Val(Mid(YYSP, 1, 2))) + " MIN " + str(Val(Mid(YYSP, 4, 2))) + " sec TO 12 midnight, HOUR"
        'YYSP = str(Val(Mid(YYSP, 1, 2))) + " MIN " + str(Val(Mid(YYSP, 4, 2))) + " sec TO THE New Year ,, 2 Thousand and 10"
        'YYSP = str(Val(Mid(YYSP, 1, 2))) + " MIN " + str(Val(Mid(YYSP, 3))) + " sec TO Zero"
        
        If DateDiff("s", WNow, Int(Now) + 1) - 1 < 60 Then
            YYSP = str(Val(Mid(YYSP, 4, 2))) '+ " sec" ' TO THE New Year"
        Else
            If Val(Mid(YYSP, 4, 2)) > 0 Then
            
                YYSP = str(Val(Mid(YYSP, 1, 2))) + " MIN " + str(Val(Mid(YYSP, 4, 2))) '+ " sec" ' TO THE New Year"
            
            Else
            
                YYSP = str(Val(Mid(YYSP, 1, 2))) + " MIN "
            End If
        End If
        
        DayOrMidNightVol = 20
        Call VolumeFixedSet(DayOrMidNightVol, True)
        'Call VoiceMP3Pause
        
        'StringVoiceToPLayActive = True
        
        If STRING_OF_VOICE_IN_QUE_PLAY = 0 Then
            STRING_OF_VOICE_IN_QUE_PLAY = Now
        End If
        
        Call SpeakVoice("TIME CD#" + YYSP)


        'TSSP$ = "3,, 2,, 1,,  12 O'CLOCK ,, New Years Day, OF  2 THOUSAND AND 10 ,, HAA HAA"
        
'        Call SpeakVoice( TS)
        
        
        '10 DAYS AGO COPS
        
        'ECT CHARGE
        'TS = "The Time is Free of Charge"
        'Call SpeakVoice( TS)
        
        'Call SpeakVoice( str$(DateDiff("s", Time10, xxhe)))
    End If

    Exit Sub

End If



'----------
'HOUR ONE
'----------
TO_GO = False
If H321 <> H2O And Minute(Time10) = 59 And Second(Time10) >= 60 - EGG Then TO_GO = True
If TO_GO = True Then
    
    H321 = Time$

    WNow = Now
    YYSP = Format(TimeSerial(23, 59, 59) - (WNow - (Int(WNow))), "hh:mm:SS")
    YYSP = Mid(YYSP, 4)
    YYSP = str(Val(Mid(YYSP, 4, 2)) + 1)
    If HOUR(Now) = 14 Or HOUR(Now) = 2 Then If Val(YYSP) > 3 Then YYSP = " 3"
    
    
    DayOrMidNightVol = 20
    Call VolumeFixedSet(DayOrMidNightVol, False)
    'Call VoiceMP3Pause
    'StringVoiceToPLayActive = True
    'Debug.Print Time$ + "   TRUE"
    
    If STRING_OF_VOICE_IN_QUE_PLAY = 0 Then
        STRING_OF_VOICE_IN_QUE_PLAY = Now
    End If
    
    Call SpeakVoice("TIME CD#" + YYSP)
    
'    ONHOURCOUNTDOWN = True
    Exit Sub
    
End If





End Sub

Public Sub VolumeFixedSet(FixSet, UseSayTimeVolAdjustLateNite)

'SHOULDNT BE NEEDED REALLY TIMERS SET OFF SO NONE CLASH
'SAFE MESSURE

'WE HAVE DOUBLE MESSURE ON TIMER MAKE SU-RE OFF NOW

If VoiceStreamActive = True Then Exit Sub
If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then Exit Sub


If MNU_DONT_ADJUST_VOL_ON_VOICE.Checked = True Then
    'THIS ALWAYS USED UNLESS OTHER WAY NEEDS BE HERE
    'TEMP OUT TEST
    
    'oVolumeSet2 = ATidalX.LblVol
    oVolumeSet2 = VolumeSet2
        'ATidalX.LblVol = str(VolumeSet1)


End If
Exit Sub


'HEAVY USE OF SAYTIME 4
'SAYTIME 4 IS ONLY USED IN BRUTE LOUDEST TIME

'REALLY WE SHALL DETECT ANY TIME VOICE HAS ENDED ALL WE
'DO IS CHK oVolumeSet2 = ATidalX.LblVol
'IF NOT THEN DO THAT MAKE AS IS WAT WAS WHEN OLD


If Mnu_NoJumpUpVol.Checked = False Then
    oVolumeSet2 = ATidalX.LblVol
    VolumeSet2 = FixSet
        
    XXYW = 0
    If VolumeSet1 > VolumeSet2 Then
        If MNU_NO_MATCH_WINAMP_VOL.Checked = False Then
            VolumeSet2 = ATidalX.LblVol
            XXYW = 1
        End If
    End If
        
    ': XXYW = 0 '-- LATE VOL FLAG
    If MNU_MAXVOL_VOICE.Checked = True Then VolumeSet2 = 100 ': XXYW = 0
    
    '#### LATE NIGHT YES
    'VOL SET #2 IS BEING ADJUSTED BUT SHOULD NOT
    'MASTER VOL SHOULD KEPT AS IS THESE ARE ONLY TEMPORY
    'IDEA WAS WE REMEMBER MASTER VOL AFTER THESE
    'YES REMEMBERED HERE AT START ROUTINE - oVolumeSet2
    
    
    'IS NO NEED THIS IF THE ABOVE IF END IF HAS BEEN USED
    If UseSayTimeVolAdjustLateNite = True And XXYW = 0 Then
        '#### HERE NIGHT TIME VOL ADJUST
        If SayTimeVolAdjustLateNite = True Then
            
            '### DIVED DONW
            VolumeSet2 = Int(FixSet / 5)
            
            'DO OUR NIB UP NEVER HAVE ON SILENT
            If VolumeSet2 < 1 Then VolumeSet2 = 1
        End If
        
        ATidalX.LblVol = str(VolumeSet2)
    End If
    
End If


End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If x > Label2.Width / 2 Then
    Label2.ToolTipText = "Talk Day Light % or Move to Left of Button Talk Sun Rise"
Else
    Label2.ToolTipText = "Talk Sun Rise or Move to Left of Button Talk Day Light%"
End If

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If x > Label3.Width / 2 Then
    Label2.ToolTipText = "Talk Darkness % or Move to Left of Button Talk Sun Rise"
Else
    Label2.ToolTipText = "Talk Sun Rise or Move to Left of Button Talk Darkness%"
End If

End Sub

Private Sub Label41_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Label2.ToolTipText = "Talk Solar Noon Time"
Exit Sub

If x > Label41.Width / 2 Then
    Label2.ToolTipText = "Talk Day Light % or Move to Left of Button Talk Sun Rise"
Else
    Label2.ToolTipText = "Talk Sun Rise or Move to Left of Button Talk Day Light"
End If

End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

MNU_SPEECH_OFF.Checked = False
Call SPEECH_OFF_SETTER

MNU_VOX_LIST2_Click
Rem WHY COMPARE EDITOR


Call MNU_VOX_LIST_Click

VOXLIST_NOT_ACTIVE = 1000


'Call Percent_SUN_1_100 ' Sub Holds the Talk Day Light % Event
'Routine Called by Dependant of
'Call TimerSunRiseSetPercent_Timer

SunRise_Not_Talk_But_Day_Light_Percent_User_Button = True
If x > Label2.Width / 2 Then
    SunRise_Not_Talk_But_Day_Light_Percent_DblClick = True
    Call TimerSunRiseSetPercent_Timer
Else
    DarkLightOLD = -2
    SunRise_Not_Talk_But_Day_Light_Percent_DblClick = False
    Empire = 1 'SunRise
    Call SunRise_Bong
End If
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

MNU_SPEECH_OFF.Checked = False
Call SPEECH_OFF_SETTER

MNU_VOX_LIST2_Click
Rem -- WHT COMPARE EDITOR

Call MNU_VOX_LIST_Click

VOXLIST_NOT_ACTIVE = 1000

'Call Percent_SUN_1_100 ' Sub Holds the Talk Day Light % Event
'Routine Called by Dependant of
'Call TimerSunRiseSetPercent_Timer

SunRise_Not_Talk_But_Day_Light_Percent_User_Button = True
If x > Label3.Width / 2 Then
    SunRise_Not_Talk_But_Day_Light_Percent_DblClick = True
    Call TimerSunRiseSetPercent_Timer
Else
    SunRise_Not_Talk_But_Day_Light_Percent_DblClick = False
    Empire = 2 'SunSet
    Call SunRise_Bong
End If
End Sub

Private Sub Label41_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

MNU_SPEECH_OFF.Checked = False
Call SPEECH_OFF_SETTER

MNU_VOX_LIST2_Click
VOXLIST.WindowState = vbNormal
VOXLIST_NOT_ACTIVE = 1000

DARK_OR_MOD_COUNTER = 0

'Call Percent_SUN_1_100 ' Sub Holds the Talk Day Light % Event
'Routine Called by Dependant of
'Call TimerSunRiseSetPercent_Timer

SunRise_Not_Talk_But_Day_Light_Percent_DblClick = False
Empire = 3 'SolarNoon
Call SunRise_Bong

Exit Sub
'----------------

If x > Label41.Width / 2 Then
    SunRise_Not_Talk_But_Day_Light_Percent_DblClick = True
    Call TimerSunRiseSetPercent_Timer
Else
    SunRise_Not_Talk_But_Day_Light_Percent_DblClick = False
    Empire = 3 'SolarNoon
    Call SunRise_Bong
End If
End Sub

Sub Percent_EQI_1_100(EQUIDATE1, EQUIDATE2, EQI_PER_T, EQISOL, EQISOL2)
'OR SPRING WINTER TYPE

'Exit Sub

'NEED THIS SUNRISE SET TIME CODE

'If StringVoiceToPLayActive = True Then Exit Sub
'PLAY COUNT DOWNS IN ONE HITT - BIT OF A DECEPTION TO THE ORDER
'If STRING_OF_VOICE_IN_QUE_PLAY >= 0 Then Exit Sub
'If ONHOURCOUNTDOWN = True Then Exit Sub

Exit Sub

If oEQUIDATE1 <> EQUIDATE1 And 1 = XXXXXX Then

    oEQUIDATE1 = EQUIDATE1

    ff = FreeFile
    Open "D:\TEMP\SEASONS.txt" For Input As #ff
    Do
        
        Line Input #ff, line1
        line1 = Replace(line1, "/", "-")
        
        If Mid(line1, 3, 1) = "-" And Len(line1) > 8 Then
            linedate = DateValue(line1)
        End If
        
        If linedate < Now + 100 Then
            Debug.Print linedate
            Debug.Print EQUIDATE1
        End If
        
        If linedate = EQUIDATE1 Then
        
        Line Input #ff, line1
        Debug.Print line1
        VBLINE = "Debug.Print line1"
        MsgBox "STOP IS HERE " + vbCrLf + VBLINE
        Stop
        End If
        'Debug.Print line1
    Loop Until EOF(ff)
    Close #ff
End If

'Stop


If EQI_PER_T = 0 Then Exit Sub



EQI_PER_T = EQI_PER_T * 2


If DayOrMidNightVol = True Then
    XDX = 1
Else
    XDX = 2
End If

EQI_PER_TT = Round(100 - EQI_PER_T, XDX)

'EQI_PER_TT2 = Round(100 - EQI_PER_T, 2)



'If EQI_PER_T >= 95 Then EQI_PER_TT = Round(EQI_PER_T, 3)
'If (EQI_PER_T - 1) < 5 Then EQI_PER_TT = Round(EQI_PER_T, 3)



If Round(EQI_PER_TT / 2, XDX) = EQI_PER_OT Then Exit Sub
EQI_PER_OT = Round(EQI_PER_TT / 2, XDX)



SPEAKOUT = ""

'ii = Round(ABH, 3)

SPEAKIN = Trim(EQISOL + ",")

ENDSPEAK = ", AWAY"

If HasWinAmpBeenPlayingRecent >= Now Then ENDSPEAK = ""
    
ENDSPEAK = " TILL HERE ON "
' ON IN OR AT

'% OF SPING TILL SUMMER
'# GOOD NO DATE THAT

'## NOW PERCENT OF SPING CHANGE AROUND WE WANT

'## DONT HAVE TO SAY SINCE MAYBE TRY WHILE


    
'SPEAKIN =
    
'EQI_SPEAK = SPEAKIN + str(EQI_PER_TT) + "%" + ENDSPEAK + SPEAKOUT
'EQISOL2


'Spring
'Vernal Equinox Mar 20 2008 05:48 UT
'Cuoo1 = Format$((longest - 93), "DD-MM-yyyy")
If EQISOL2 = "Spring" Then Cuoo = " -93 Days  From Summer Solstice THAT IS LONGEST DAY"


'Summer Solstice
'  Summer Solstice Jun 20 2008 23:59 UT
'Cuoo2 = Format$((longest), "DD-MM-yyyy")
If EQISOL2 = "Summer" Then Cuoo = " -93 Days  From Summer Solstice THAT IS LONGEST DAY"

'autumn
'  Autumnal Equinox Sep 22 2008 15:44 UT
'Cuoo3 = Format$(longest + 93, "DD-MM-yyyy")

'Winter
'  Winter Solstice Dec 21 2008 12:04 UT
'Cuoo4 = Format$((shortest), "DD-MM-yyyy")


FFT1 = ENGLISH_DATE_FOR_TALK(EQUIDATE1, False)
'FFT1 = Replace(FFT, "-", XX1 + LCase(Tx2) + " Of")

'FFT2 = Replace(FFT, "-", XX1 + LCase(Tx2) + " Of")
FFT2 = ENGLISH_DATE_FOR_TALK(EQUIDATE2, False)



SPEAKOUT = ""
If EQISP Mod 1 = 0 Then
    'SPEAKOUT = " From " + FFT1 + " To " + FFT2 '+ " AND THEN WILL BE, DRUM ROLL,, WAIT FOR IT,,,,, " + EQISOL2 + "!"
End If

'If EQISP Mod 5 = 0 Then
    'SPEAKOUT = ", From " + Format(EQUIDATE1, "DD MMMM")
    'SPEAKOUT = SPEAKOUT + ", To " + Format(EQUIDATE2, "DD MMMM")
    'SPEAKOUT = SPEAKOUT + ", Art Is Known As, THE Cross Quarter Season, Being FOR " + EQISOL2 + " , Plus 93 Day's, From The Equinox AND Solistice"
'End If

EQISP = EQISP + 1

'EQI_SPEAK = EQISOL2 + "! is " + Trim(str(EQI_PER_TT)) + "%! From The " + FFT1 + "!, Until " + EQISOL + "! That Begin's, On The " + FFT2 + "!"

'TWO LINES NOT BAD -- WELL TEMP OUT
'EQI_SPEAK = EQISOL2 + "! is " + Trim(str(EQI_PER_TT)) + "%!  On The " + FFT1 + "!, Until!, " + EQISOL + "! On The " + FFT2 + "!"

'Debug.Print EQI_SPEAK
'Spring! is 20.1%!  From The 3rd of February!, Until Summer! Begins on The 5th of May!



'DATES
'SPEAKSun = "Sun " + str(99.8) + " % "


If EQI_SPEAK <> "" And EQI_Run1ST = 1 Then


    Call VolumeFixedSet(10, True)
    
    Call SpeakVoice(EQI_SPEAK)
 
    EQI_Run2nd = 2
End If


If EQI_SPEAK <> "" Then EQI_Run1ST = 1



End Sub


Sub Percent_Moon_1_100(MOON_PER_T)

If FLAG_TO_MOON_TALK = False Then
    If MOON_PER_T = 0 Then Exit Sub
    If MOON_PER_T = MOON_PER_OT Then Exit Sub
End If

'If StringVoiceToPLayActive = True Then Exit Sub
'If ONHOURCOUNTDOWN = True Then Exit Sub
'If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then Exit Sub

'# PUT IN VARS
'Dim PER_i, PER_i2, MOON_PER_OT, PER_x, PER_x1, PER_MULTI



'REM HERE FINAL IMPLIMENT CODE
'------------------------------------
'PER_x = -1
'PER_i2 = 0
'PER_OT = -1
'PER_T = 94
'PER_i = 0.01
'If PER_i > 0 Then PER_i2 = 1 Else PER_i2 = -1
'------------------------------------


'------------------------------------------------------------
'IF FIRST RUN VALUE NOT SET
'WAIT FOR MOON LUMINOSITY DIRECTION
'------------------------------------------------------------
'MOON SMALLER BIGGER
'USE THIS WITH IMPLIMENT AND TEST CODE
'---------------------------------------
'GET THE MOON DIRECTION
'---------------------------------------
MOON_PER_OT_DONE = False
If MOON_PER_OT = 0 Then
    MOON_PER_OT = MOON_PER_T
    Do
        MOON_LIGHT = Illum(ConvertToUT(Now))
        'WHEN MOON_LIGHT -- CALLS ILLUM() IT ALSO GIVES THE RESULT TO -- MOON_PER_T
        '--------------------------------------------------------------------------
        'AVOIDER CROSS OVER PROBLEM BY WAIT LONGER AT EXTREME
        '5 DIGIT AFTER DECIMAL
        'TAKE FRM 3RD DIGIT
        THIRD_DIGIT_OFFSET = 0.005
        '----------------------------------------------------
        If MOON_LIGHT > THIRD_DIGIT_OFFSET And MOON_LIGHT < 100 - THIRD_DIGIT_OFFSET Then
            If MOON_LIGHT <> MOON_PER_OT Then
                'OPPOITE DIRECTION WRONG OR CORRECT
                'MAYBE FIRST HITT CORRECT BUT LATER NOT SWAP AROUND
                'CAREFUL WHEN EXTREME MIGHT REQUIRE LOOK OVER
                '--------------------------------------------------
                If MOON_PER_OT < MOON_LIGHT Then PER_i2 = -1 Else PER_i2 = 1
                MOON_PER_OT_DONE = True
                Exit Do
            End If
        End If
        Sleep 100
        COUNT_TASK = COUNT_TASK + 1
        DoEvents
    Loop Until MOON_PER_OT_DONE = True
End If
MOON_PER_OT = MOON_PER_T

If MOON_PER_OT_DONE = True Then
    Debug.Print "----------------------------------------------------"
    Debug.Print Trim(str(COUNT_TASK)) + " OF SLEEP 100 MILLISECOND WAIT "
    Debug.Print "RESULT 100Ms WAIT TO 1000 MILLI-SECOND WAIT FOR THIS"
    Debug.Print "THIS MAYBE LONGER DURING THE CROSS OVER"
    Debug.Print "----------------------------------------------------"
    'EASY TO MAINTAIN THE CROSS OVER IN SOFTWARE
    'LETS GET IT DONE
    'LET'S GET IT DONE - DONE -- Tue 20 September 2016 13:10:01----------
    '-----------------------------------------------------------------
    'Sub Moon_Code()
    '-----------------------------------------------------------------
    'IN THIS SUB ROUTINE IS THE MOON STATECHANGE AND ANNOUNCE
    '-----------------------------------------------------------------
End If




'If MOON_PER_OT < MOON_PER_T Then PER_i2 = 1 Else PER_i2 = -1
'If MOON_PER_OT = -1 Then
'    If PER_i > 0 Then PER_i2 = 1 Else PER_i2 = -1
'End If
'MOON_PER_OT = MOON_PER_T
'------------------------------------------------------------




'REM HERE FINAL IMPLIMENT CODE
'------------------------------------
'Me.Show
'Do
'Sleep 200
'DoEvents
'------------------------------------

SPEAKMOON = ""
PER_MULTI = 1: IK = 1

PER_MULTI = 10: IK = 2 '    = 1 DECIMAL PLACES
PER_MULTI = 100: IK = 3 '    = 2 DECIMAL PLACES

'If MOON_PER_T >= 97 Then PER_MULTI = 100: IK = 3
'If MOON_PER_T >= 99 Then PER_MULTI = 100
'If (MOON_PER_T - 1) < 3 Then PER_MULTI = 100: IK = 3
'If (MOON_PER_T - 1) < 1 Then PER_MULTI = 100



'2 DECI PLACE
'PER_MULTI = 100: IK = 3
'IK_TXT = String(IK - 1, "0")
'PER_MULTI_TXT = "0." + IK_TXT


'3 DECI PLACE
'PER_MULTI = 1000: IK = 4



If 1 = 2 Then
    If SayTimeVolAdjustLateNite = True Then
        PER_MULTI = 10: IK = 2 '  = 2 DECIMAL PLACES
    End If
End If

If FLAG_TO_MOON_TALK = True Then PER_x = 0

ABH = -1
If (MOON_PER_OT > 0 And PER_i2 <> 0) Or FLAG_TO_MOON_TALK = True Then
'
'    If PER_i2 > 0 Then
'        '--------------------------------------------------
'        'FORMAT(NUMNER) ROUNDING IT UP - NOT WHAT WE WANT
'        'If Format(MOON_PER_T, PER_MULTI_TXT) <> PER_x Then
'        '    PER_x = Format(MOON_PER_T, PER_MULTI_TXT)
'        '--------------------------------------------------
'
'        If Int(MOON_PER_T * PER_MULTI) <> PER_x Then
'            PER_x = Int(MOON_PER_T * PER_MULTI)
'            'REM HERE FINAL IMPLIMENT CODE
'            'Label1 = Format(Int(MOON_PER_T * PER_MULTI) / PER_MULTI, "0.00")
'
'            ix = Round(Int(MOON_PER_T * PER_MULTI) / PER_MULTI, IK)
'            ABH = Format(ix, "0.#####")
'
'        End If
'    End If
'
'    If PER_i2 < 0 Then
'
'        If Int(MOON_PER_T * PER_MULTI) <> PER_x Then
'            PER_x = Int(MOON_PER_T * PER_MULTI)
'
'            'REM HERE FINAL IMPLIMENT CODE
'            'Label1 = Format((PER_x / PER_MULTI) + (1 / PER_MULTI), "0.00")
'            ix = Round(Int(MOON_PER_T * PER_MULTI) / PER_MULTI, IK)
'            ABH = Format(ix, "0.#####")
'
'        End If
'    End If

    'rem abort the above for a moment same result send out
    PER_x = Int(MOON_PER_T * PER_MULTI)
    ix = Round(Int(MOON_PER_T * PER_MULTI) / PER_MULTI, IK)
    ABH = Format(ix, "0.#####")
End If
'------------------------------------------------------------------------------------------------
'THIS GET THE DECIMAL PLACE BUTER AND THE SETTING OF HOW OFTEN TALKER IS REPLACE NOW MOD VARIABLE
'------------------------------------------------------------------------------------------------

'If ABH = 0 Then Exit Sub

If FLAG_TO_MOON_TALK = False Then
    If ABH = O_ABH Then Exit Sub
End If
O_ABH = ABH


'AFTER DIVEDE ABOVE ONLY EVERY OTHER # ALSO
'MIN WITH MOD EVERY OTHER ODD EVEN ONE AT MOD 1
'CAN'T DO MOD 1 OR IT BE ALL THE TIME TRY 2 AND UP
'--------------------------------------------------------
'THE CODE IS ALL POLLED TO QUICK
'SO CHANGE TO EVERY OTHER READING USE ABH >= 0

' __ FLAG_TO_MOON_TALK __ COME FROM CLICK EVENT PRESS_ER
' __ SO BY-PASS THIS BIT OF DELAY IF SETTER
'------------------------------------------------------------
'If FLAG_TO_MOON_TALK = False Then
'
'        If MOON_MOD_COUNTER = -1 Then
'            MOON_MOD_COUNTER = 1
'            Else
'
'            ' SET THE SPEED OF THE MOON TALK
'            If MOON_MOD_COUNTER Mod 2 = 0 Then
'                'NOTHING
'                MOON_MOD_COUNTER = 1
'            Else
'                MOON_MOD_COUNTER = MOON_MOD_COUNTER + 1
'                '--------------------------------------------------------------
'                'EXIT SUB HERE WITH THE DELAY OF COUNT A FEW AND THEN LET RIPER
'                '--------------------------------------------------------------
'                FLAG_TO_MOON_TALK = False
'                'Exit Sub
'                '--------------------------------------------------------------
'            End If
'        End If
''    End If
'End If

If ABH = 0 Or ABH = 100 Then MOON_MOD_COUNTER = 0

If FLAG_TO_MOON_TALK = False Then

    ' ------------------------------
    ' SET THE SPEED OF THE MOON TALK
    ' ------------------------------
    'Mod 1 __ EVERY ONE         = EVERY 1  MINUTE  55 SECOND __ 115 SECONDS
    'Mod 2 __ EVERY OTHER ONE 2 = EVERY 3  MINUTES 50 SECOND __ 230 SECONDS
    'Mod 3 __ EVERY OTHER ONE 3 = EVERY 4  MINUTES 55 SECOND __ 295 SECONDS
    'Mod 4 __ EVERY OTHER ONE 4 = EVERY    MINUTES    SECOND __ 460 SECONDS
    'Mod 8 __ EVERY OTHER ONE 8 = EVERY 15 MINUTES 20 SECOND __ 920 SECONDS
    ' ------------------------------
    If MOON_MOD_COUNTER Mod 8 = 0 Then
        TALK_MOON_FLAG_VAR = True
    Else
        TALK_MOON_FLAG_VAR = False
    End If
    
    MOON_MOD_COUNTER = MOON_MOD_COUNTER + 1
    
    If TALK_MOON_FLAG_VAR = False Then
        Exit Sub
    End If

End If

'------------------------------------------------------------------------------
'ADJUSTMENT --------------------
'HERE --------------------------
'THE
'Percent_SUN_1_100 -- DAYLIGHT %
'AND
'HERE --------------------------
'THE
'Percent_Moon_1_100 ----- MOON %
'------------------------------------------------------------------------------

'MOON_PER_OT =0 = FIRST RUN --- GOT TO EQUAL A VAULE FIRST
'If ABH >= 0 Or MOON_PER_OT = 0 Then
If ABH >= 0 Or FLAG_TO_MOON_TALK = True Then
    SPEAKMOON = str(ABH) + "%"
End If


'------------------------------------------------------------
'MOON SMALLER BIGGER
'USE THIS WITH IMPLIMENT AND TEST CODE
'---------------------------------------
'If MOON_PER_OT < MOON_PER_T Then PER_i2 = 1 Else PER_i2 = -1
'If MOON_PER_OT = -1 Then
'    If PER_i > 0 Then PER_i2 = 1 Else PER_i2 = -1
'End If
'MOON_PER_OT = MOON_PER_T
'------------------------------------------------------------

'---------------------------------------
'NEED SOMETHING AROUND HERE TO TRIGGER TALK
'WITH
'FLAG_TO_MOON_TALK = true



'FIRST RUN WON'T KNOW IF SMALLER BIGGER -- SO DO WITHOUT THAT BIT AND Run1STMoon IS USED

If SPEAKMOON <> "" Then 'And Run1STMoon = 1 Then

    'PROBLEM
    'Call VolumeFixedSet(10, True)
        
        SPEAKMOON = Trim(SPEAKMOON)
        'if FLAGMOONSAYMORE = True Then
        'swap here or later opposite work setter
        '---------------------------------------
        If PER_i2 = -1 Then
            MOON_STATE_LIGHTER = "Swell!"
        Else
            MOON_STATE_LIGHTER = "Smaller!"
        End If
        
        SPEAKMOON = "MOON! " + MOON_STATE_LIGHTER + " " + SPEAKMOON
        
        'If PER_i2 = -1 Then SpeakMoon = SpeakMoon + ", Smaller!"
        'If PER_i2 = 1 Then speakmoon = speakmoon + ", Greatness!" '--------", Greator!"
        'If PER_i2 = 1 Then SpeakMoon = SpeakMoon + ", Swell!"
        
        
        FLAGMOONSAYMORE = False
    'End If
    
    'Call VoiceMP3Pause
    Call SpeakVoice(SPEAKMOON)

End If

'If SpeakMoon <> "" And Run1STMoon = 0 Then
'
'    'MOON HAS NOT A VAULE YET BECAUSE IMPORTANT TO KNOW UP OR DOWN
'    'SO TAKE A VAULE THIS WAY WITH FIRST SELECTION I THINK ROUNDING UP METHOD
'    If ABH = -1 Then
'        'ABH
'        ix = Round(Int(MOON_PER_T * PER_MULTI) / PER_MULTI, IK)
'        ABH = Format(ix, "0.#####")
'    End If
'
'
'    'WHY
'    FLAGMOONSAYMORE = False
'    Call SpeakVoice(SpeakMoon)
'
'End If
'
'
'If SpeakMoon <> "" Then Run1STMoon = 1



End Sub


Sub Percent_SUN_1_100(SUN_PER_T, DarkLight)

'Exit Sub

'FORM_STARTING
If SUN_PER_T = -1 Then Exit Sub

'If StringVoiceToPLayActive = True Then Exit Sub
'If ONHOURCOUNTDOWN = True Then Exit Sub

'If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then Exit Sub

'SUN_PER_T = SUN_PER_T / 2

'MORE NORM
'SUN_PER_TT = Round(SUN_PER_T, 0)


'If SayTimeVolAdjustLateNite = False Or 1 = 1 Then
'    'BIT FREQ THIS ONE
'    SUN_PER_TT = Round(SUN_PER_T, 1)
'End If


'If Val(Format(SUN_PER_T, "0.000")) = DarkLightOLD Then
'If SUN_PER_T = DarkLightOLD Then
'    EQUAL_THEN_NOT_TALKER = True
'Else
'    EQUAL_THEN_NOT_TALKER = False
'End If

EQUAL_THEN_NOT_TALKER = True

WANT_SUN_TALKER = False
If Format(SUN_PER_T, "0.00") <> Format(DarkLightOLD, "0.00") Then
    WANT_SUN_TALKER = True
End If

WANT_TALKER_SUN_DEMAND = False
If FIRST_RUN_SUN_PERCENT_TALKER_LATCH_UP = True And True = False Then
    DARK_OR_MOD_COUNTER = 0
    WANT_TALKER_SUN_DEMAND = True
    FIRST_RUN_SUN_PERCENT_TALKER_LATCH_UP = False
End If
If SunRise_Not_Talk_But_Day_Light_Percent_User_Button = True Then
    DARK_OR_MOD_COUNTER = 0
    WANT_TALKER_SUN_DEMAND = True
    SunRise_Not_Talk_But_Day_Light_Percent_User_Button = False
End If
If SunRise_Not_Talk_But_Day_Light_Percent_DblClick = True Then
    DARK_OR_MOD_COUNTER = 0
    WANT_TALKER_SUN_DEMAND = True
    SunRise_Not_Talk_But_Day_Light_Percent_DblClick = False
End If

If WANT_SUN_TALKER = False And WANT_TALKER_SUN_DEMAND = False Then Exit Sub

'If WANT_SUN_TALKER = True Then
    'EQUAL_THEN_NOT_TALKER = False
    'DIVEDE 2 MAKE LESS FREQ
    'If Round(SUN_PER_TT / 2, 1) = Sun_PER_OT Then Exit Sub
    'Sun_PER_OT = Round(SUN_PER_TT / 2, 1)
    '-------------------------------------
    'If Round(SUN_PER_TT, 1) = Sun_PER_OT Then Exit Sub
    'Sun_PER_OT = Round(SUN_PER_TT, 1)
    '-------------------------------------
    'SUN_PER_TT = Round(SUN_PER_T, 1)
    'If Val(Format(SUN_PER_T, "0.0")) <> SUN_PER_TT Then Stop
    'TEST RESULT FORMAT = 75.1 AND Round(SUN_PER_T, 1) = 75 RESULT SAME OF TIME BEFORE
    

    'SHOULD CALL A DARK% EVERY ROUND TO MIDDLE .05 OR .#5
    
    'SUN_PER_TT = Val(Format(SUN_PER_T, "0.0"))
    'If SUN_PER_TT = Sun_PER_OT Then Exit Sub
    'Sun_PER_OT = SUN_PER_TT
    '-------------------------------------
    
    'AFTER DIVEDE ABOVE ONLY EVERY OTHER # ALSO
    'MIN WITH MOD EVERY OTHER ODD EVEN ONE AT MOD 1
    'CAN'T DO MOD 1 OR IT BE ALL THE TIME TRY 2 AND UP
    '--------------------------------------------------------
    'If DARK_OR_MOD_COUNTER = -1 Or Test = Test Then
    
If WANT_TALKER_SUN_DEMAND = False Then
    
    If Format(SUN_PER_T, "0.00") <> Format(DarkLightOLD, "0.00") Then
        If Format(SUN_PER_T, "0.00") = "0.00" Then DARK_OR_MOD_COUNTER = 0
        If Format(SUN_PER_T, "100.00") = "100.00" Then DARK_OR_MOD_COUNTER = 0
    End If
    '------------------------------------------------
    'SET THE TIME SPEAK SPEED
    '------------------------------------------------
    'MOD 100  __  52 SECONDS
    'MOD 1700 __ 884 SECONDS __ 14 MINUTES 44 SECONDS
    '------------------------------------------------
    
    If DARK_OR_MOD_COUNTER Mod 1700 = 0 Then
            SAY_SUN_TIME = True
        Else
            SAY_SUN_TIME = False
    End If
    
    DARK_OR_MOD_COUNTER = DARK_OR_MOD_COUNTER + 1
    
    If SAY_SUN_TIME = False Then
        Exit Sub
    End If
End If


'---------------------------------------------------
'3 DECIMAL PLACE
'MAYBE ROUNDED AND DECIMAL 2 OR LESS SHOWING
'MUST REDCUE ONE DECIMAL PLACE TO MAKE LESS FREQUENT
'---------------------------------------------------
DarkLightOLD = SUN_PER_T
'-----

'------ TO DO WITH SOUND STOP MP3 AT CROSS DAYLIGHT TIMES
'XAZY(1) = False
'If SUN_PER_T >= 98 Then XAZY(1) = True
'---
'If DarkLight = "DARK" Then XXV = 4 Else XXV = 2
'---
'If SUN_PER_T <= XXV Then XAZY(1) = True

'NOON
'If SUN_PER_T >= 49.5 And SUN_PER_T <= 50.5 Then XAZY(1) = True
'---


'If SUN_PER_T >= 95 Then SUN_PER_TT = Round(SUN_PER_T, 1)
'If (SUN_PER_T - 1) < 5 Then SUN_PER_TT = Round(SUN_PER_T, 1)

SPEAKOUT = ""

'SPEAKIN = EQISOL + ","
'ENDSPEAK = ", AWAY"

'If HasWinAmpBeenPlayingRecent >= Now Then ENDSPEAK = ""

If DarkLight = "DARK" Then SPEAKIN = "DarkNess! ": ixx = 100 - TTA3
If DarkLight = "LIGHT" Then SPEAKIN = "Day Light! ": ixx = TTA3
'If DarkLight = "DARK" Then SPEAKIN = "Da-Da-Dra Darkness ": ixx = 100 - TTA3


'IXX = Length OF DAY
'SPEAKOUT = ""
'If XABG Mod 4 = 0 Or Run2ndSun = 0 Then
'    SPEAKOUT = "OF " + Format(ixx, "0") + "%, IN DAY"
'End If
'If XABG Mod 8 = 0 Then
'     SPEAKOUT = "OF " + Format(ixx, "0") + "%, IN DAY " ', Length OF " + SPEAKIN
'    If XABG > 999999 Then XABG = 0
'End If
'XABG = XABG + 1
'
''If HasWinAmpBeenPlayingRecent >= Now Then SPEAKOUT = ""
'
'SPEAKSun = ""
'SPEAKOUT = ""

'SPEAKSun = SPEAKIN + Trim(str(SUN_PER_TT)) + "%" + COMMA + SPEAKOUT

SunResult_Reverse_100 = 100 - SunResult_Reverse

'----------------------------------------------------------------
'Don't Talk % Day Light & When Sun Rise If User Want Button
'or Only the Other
'----------------------------------------------------------------
'THIS BUTTON OPTION EITHER TRUE OR FALSE DEPEND IF TALK
'WANTED % OF DAY LIGHT OR OTHER LEFT RIGHT OF BUTTON
'LOOKER BELOW ONLY SET ONCE FOR THE BUTTON
'----------------------------------------------------------------
'SunRise_Not_Talk_But_Day_Light_Percent_User_Button = True
'----------------------------------------------------------------
'THIS ONE FOR EITHE SIDE LEFT RIGHT ONE ABOVE IF TRIGGER BY USER
'----------------------------------------------------------------
'SunRise_Not_Talk_But_Day_Light_Percent_DblClick = True
'----------------------------------------------------------------
'TWO TYPE 2 -- 1 DAYLIGHT SUNRISE TYPE THING OTHER DAYLIGHT %
'----------------------------------------------------------------
'SPEAKSun = SPEAKIN + Format(SunResult_Reverse, "0.0000") + "%, Thesaurus " + Format(SunResult_Reverse_100, "0") + "%" + COMMA + SPEAKOUT

SPEAKSun = SPEAKIN + Format(SunResult_Reverse, "0.000") + "%, Opposite " + Format(SunResult_Reverse_100, "0") + "%" + COMMA + SPEAKOUT

'----
'Antonym Synonyms, Antonym Antonyms | Thesaurus.com
'http://www.thesaurus.com/browse/antonym?s=t
'----

'Equinox to Precision to Time in Day
'----
'Equinox -Wikipedia
'https://en.wikipedia.org/wiki/Equinox
'----

If SPEAKSun <> "" Then
        Call VolumeFixedSet(10, True)
        Call SpeakVoice(SPEAKSun)
        Run2ndSun = 2
End If

End Sub


'### 01
Function NewNoon()
    PlusNow = 0
    Do
        Dim srs As New clsSunRiseSet
        srs.City = "London, England"
        srs.DateDay = Now + PlusNow
        srs.DaySavings = DaySavingsSub(srs.DateDay)
    
        srs.CalculateSun

        NewNoon = srs.SolarNoon
        PlusNow = PlusNow + 1
    Loop Until NewNoon > Now

    TooLNewNoon = NewNoon
    i = LASTNoon
    
    '## Todays
    srs.DateDay = Now
    srs.DaySavings = DaySavingsSub(srs.DateDay)
    srs.CalculateSun
    TooLTodayNoon = srs.SolarNoon


End Function

'### 02
Function NewSunRise()
    PlusNow = 0
    Do
        Dim srs As New clsSunRiseSet
        srs.City = "London, England"
        srs.DateDay = Now + PlusNow
        srs.DaySavings = DaySavingsSub(srs.DateDay)
    
        srs.CalculateSun

        NewSunRise = srs.Sunrise
        PlusNow = PlusNow + 1
    Loop Until NewSunRise > Now

    Tool1 = NewSunRise

    srs.DateDay = srs.DateDay - 1
    srs.DaySavings = DaySavingsSub(srs.DateDay)
    srs.CalculateSun
    Tool4A = srs.Sunrise
    
    TooLNewSunRise = NewSunRise
    i = LASTSunRise
    
    '## Todays
    srs.DateDay = Now
    srs.DaySavings = DaySavingsSub(srs.DateDay)
    srs.CalculateSun
    TooLTodaySunRise = srs.Sunrise


End Function

'### 03
Function NewSunSet()
    PlusNow = 0
    Do
        Dim srs As New clsSunRiseSet
        srs.City = "London, England"
        srs.DateDay = Now + PlusNow
        srs.DaySavings = DaySavingsSub(srs.DateDay)
    
        srs.CalculateSun

        NewSunSet = srs.Sunset
        PlusNow = PlusNow + 1
    Loop Until NewSunSet > Now
    
    Tool2 = NewSunSet
    
    srs.DateDay = srs.DateDay - 1
    srs.DaySavings = DaySavingsSub(srs.DateDay)
    srs.CalculateSun
    Tool4B = srs.Sunset

    TooLNewSunSet = NewSunSet
    i = LASTSunSet
    
    
    '## Day Before
    srs.DateDay = Now - 1
    srs.DaySavings = DaySavingsSub(srs.DateDay)
    srs.CalculateSun
    TooLDayBeforeSunSet = srs.Sunset
    '## Todays
    srs.DateDay = Now
    srs.DaySavings = DaySavingsSub(srs.DateDay)
    srs.CalculateSun
    TooLTodaySunSet = srs.Sunset
    
    
    
    
End Function


'----------
'### 01

Function LASTNoon()
    PlusNow = 1
    Do
        Dim srs As New clsSunRiseSet
        srs.City = "London, England"
        srs.DateDay = Now + PlusNow
        srs.DaySavings = DaySavingsSub(srs.DateDay)
    
        srs.CalculateSun

        LASTNoon = srs.SolarNoon
        PlusNow = PlusNow - 1
    Loop Until LASTNoon < Now

    TooLLastNoon = LASTNoon

End Function


'### 02

Function LASTSunRise()
    PlusNow = 1
    Do
        Dim srs As New clsSunRiseSet
        srs.City = "London, England"
        srs.DateDay = Now + PlusNow
        srs.DaySavings = DaySavingsSub(srs.DateDay)
    
        srs.CalculateSun

        LASTSunRise = srs.Sunrise
        PlusNow = PlusNow - 1
    Loop Until LASTSunRise < Now

    'Tool1 = LASTSunRise

    'srs.DateDay = srs.DateDay - 1
    'srs.DaySavings = DaySavingsSub(srs.DateDay)
    'srs.CalculateSun
    'Tool4A = srs.Sunrise

    TooLLastSunRise = LASTSunRise

    
End Function


'### 03

Function LASTSunSet()
    PlusNow = 1
    Do
        Dim srs As New clsSunRiseSet
        srs.City = "London, England"
        srs.DateDay = Now + PlusNow
        srs.DaySavings = DaySavingsSub(srs.DateDay)
    
        srs.CalculateSun

        LASTSunSet = srs.Sunset
        PlusNow = PlusNow - 1
    Loop Until LASTSunSet < Now
    
    'Tool2 = LASTSunSet
    '
    'srs.DateDay = srs.DateDay - 1
    'srs.DaySavings = DaySavingsSub(srs.DateDay)
    'srs.CalculateSun
    'Tool4B = srs.Sunset
    TooLLastSunSet = LASTSunSet

End Function







Sub SunRise_Bong()

'Don't Talk % Day Light & When Sun Rise If User Want Button
'or Only the Other
If SunRise_Not_Talk_But_Day_Light_Percent_User_Button = True Then
    '------------------------
    'Not to talk for Bong it when True
    If SunRise_Not_Talk_But_Day_Light_Percent_DblClick = True Then
        Exit Sub
    End If
End If
SunRise_Not_Talk_But_Day_Light_Percent_User_Button = False


'## ONLY AT BEGINING
'WE CAN DUMMY INTO iDate VAR AS VAR IN Q' IN IF STATEMENT
'IS GOT FROM FUINCTION

'WHENEVER CALLED AN ELSE WHERE ALL OTHER VARS GOT
' TODAY AND LAST SUNSET'S

If TooLNewSunRise = 0 Then Idate = NewSunRise
If TooLNewSunSet = 0 Then Idate = NewSunSet
If TooLNewNoon = 0 Then Idate = NewNoon



'SUNRISE
If Now >= TooLNewSunRise Then
    Idate = NewSunRise
    Empire = 1
    DarkLightOLD = -2
End If

'SUNSET
If Now >= TooLNewSunSet Then
    Idate = NewSunSet
    Empire = 2
    DarkLightOLD = -2
End If

'SolarNoon
If Now >= TooLNewNoon Then
    Idate = NewNoon
    Empire = 3
End If



If Empire = 1 Then
    SUNSAYSTRING = "Sun Rise .": TOOL1X = TooLTodaySunRise
End If
If Empire = 2 Then SUNSAYSTRING = "Sun Set .": TOOL1X = TooLTodaySunSet
If Empire = 3 Then SUNSAYSTRING = "Sun Noon .": TOOL1X = TooLTodayNoon

   
If Empire > 0 Then
    
    Call VolumeFixedSet(80, False)

    'Call VoiceMP3Pause
    DEMAND_MP3_OFF = True
    
    Call MP3_OFF__ON_START_VOICE
        
    If Wqe2$ <> "" Then
        ATidalX.MMControl2.Command = "prev"
        ATidalX.MMControl2.Command = "Play"
    Else
        ATidalX.MMControl1.Command = "prev"
        ATidalX.MMControl1.Command = "Play"
    End If

    Call WaitWavFinish

End If
    
    

    
'Call SaySunRiseSet

If Empire > 0 Then
    
    If Minute(TOOL1X) < 10 And Minute(TOOL1X) > 0 Then crud$ = "o" + str(Minute(TOOL1X)) Else crud$ = " " + str(Minute(TOOL1X))
    If Minute(TOOL1X) = 0 Then crud$ = "O'Clock"

    YY5 = HOUR(TOOL1X)
    If YY5 > 12 Then YY5 = YY5 - 12
    
    STRINGXX = ""
    
    If TOOL1X > Int(Now) + 1 Then STRINGXX = "Tommorrow's "
    
    
    Call SpeakVoice(STRINGXX + SUNSAYSTRING + str$(YY5) + " " + crud$ + " and " + str$(Second(TOOL1X)) + " Seconds")
    
    'IS THIS USEFUL FOR ANYTHING
    SunSetFor0Percent = True

    Call TimerSunRiseSetPercent_Timer

End If

Empire = 0

End Sub

Function DaySavingsSub(daysaving As Date)

'Daylight Saving, system of setting clocks 1 or 2 hours ahead so that both sunrise and sunset occur at a
'later hour, producing an additional period of daylight in the evening. In the North Temperate Zone clocks
'are usually set ahead 1 hour in the spring and set back to standard time in the fall. The idea of
'daylight saving was mentioned in a whimsical essay in 1784 by Benjamin Franklin; it was first
'advocated seriously by a British builder, William Willett, in the pamphlet Waste of Daylight (1907).
'Daylight saving has been used in the United States and in many European countries since World War I,
'when the system was adopted in order to conserve fuel needed to produce electric power.
'Some localities reverted to standard time after the war, but others retained daylight saving.
'During World War II the U.S. Congress passed a law putting the entire country on "war time,"
'which set clocks 1 hour ahead of standard time for the duration of the war. War time was also followed
'in Britain, where clocks were put ahead s
'till another hour during the summer. In the U.S. during peacetime, daylight saving was a subject of
'controversy. Farmers, who usually work schedules determined by sun time and are therefore
'inconvenienced when they must conduct business on a different time basis, registered strong
'opposition. Railroad, bus, and plane scheduling was hampered by time inconsistencies among
'various cities and states. The Uniform Time Act, enacted by the U.S. Congress in 1966, established
'a system of uniform (within each time zone) daylight saving time throughout the U.S. and its possessions,
'exempting only those states in which the state legislature voted to keep the entire state on standard time.
'Under legislation enacted in 1986, daylight saving time begins at 2 AM on the first Sunday of April and
'ends at 2 AM on the last Sunday of October.

'"Daylight Saving." Microsoft® Encarta® Encyclopedia 2001. © 1993-2000 Microsoft Corporation. All rights reserved.



   Dim A1 As Date
   Dim A2 As Date
   Dim R1 As Date
   Dim w1 As Integer
   'Dim daysaving As Long
   
   
    A1 = DateSerial(Year(daysaving), 3, 31) + TimeSerial(2, 0, 0)
    For R1 = A1 To A1 - 20 Step -1
        w1 = Weekday(R1)
        If w1 = 1 Then Exit For
    Next
    A1 = R1
    A2 = DateSerial(Year(daysaving), 10, 31) + TimeSerial(2, 0, 0)
    For R1 = A2 To A2 - 20 Step -1
     w1 = Weekday(R1)
     If w1 = 1 Then Exit For
    Next
    A2 = R1
    
    'seems to want a inverted result this
    If daysaving < A1 Then DaySavingsSub = False
    If daysaving > A1 Then DaySavingsSub = True
    If daysaving > A2 Then DaySavingsSub = False


End Function


Sub SunRise_SunSet()



'If Int(Timer) <> splash2 Then

'GoTo jump9
    
    If (Tool1 < Now + speedtime And Tool2 > Now + speedtime And Tagbok = 1) Or Tagbok = -1 Then
 '       Label2.BackColor = QBColor(15)
  '      Label2.ForeColor = QBColor(0)
   '     Label3.BackColor = QBColor(0)
    '    Label3.ForeColor = QBColor(15)
        Tagbok = 0
    '    srt1$ = "> ": srt2$ = " <"
    End If

    If (Tool1 < Now + speedtime And Tool2 < Now + speedtime And Tagbok = 0) Or Tagbok = -1 Then
      '  Label2.BackColor = QBColor(15)
       ' Label2.ForeColor = QBColor(0)
        'Label3.BackColor = QBColor(0)
        'Label3.ForeColor = QBColor(15)
        'sst$ = "-> "
        Tagbok = 1
     '   sst1$ = "> ": sst2$ = " <"
    End If
'jump9:
'End If

'If tagbok = 0 Then srt1$ = "> ": srt2$ = " <"
'If tagbok = 1 Then sst1$ = "> ": sst2$ = " <"


'splash2 = Int(Timer)


    Dim j41 As Date
    Dim j42 As Date
    Dim j11 As Date
    Dim j22 As Date
    Dim j1 As Date
    Dim j2 As Date


If Changeday <> Int(Now) Or DemandSun = True Then
    DemandSun = False
    
    For r = 1 To 0 Step -1
    
    Call Equinox3
  
    'If Changeday <> 0 Then Extra = 0
  
    Changeday = Int(Now)

    'If extra = 1 Then Label9.Caption = "â" Else Label9.Caption = "á"
     
    Dim srs As New clsSunRiseSet
    srs.City = "London, England"
    

    'srs.DateDay = Now - 1 + speedtime + Extra
    'srs.DateDay = Now - 1 + Extra
    'srs.DateDay = DateValue("29-03-2008")
    
    
    'seems to be want inverted result this
    'srs.DaySavings = DaySavingsSub(srs.DateDay)
    'srs.DaySavings = True
    
    'If DaySavingsSub(srs.DateDay + 1) = True And srs.DaySavings = False Then xdas = 1
    
    'srs.CalculateSun

    'TooLNewSunRise = 0
    'TooLNewSunSet = 0

    'tool41 = srs.Sunrise
    'tool42 = srs.Sunset
    'tool43 = srs.SolarNoon
    
    'If xdas = 1 Then tool41 = tool41 + TimeSerial(0, 1, 0)
    'If xdas = 1 Then tool42 = tool42 + TimeSerial(0, 1, 0)
    'If xdas = 1 Then tool43 = tool43 + TimeSerial(0, 1, 0)

    'tool46$ = Format$(tool42 - tool41, "HH:MM:SS")

    'work yesterday rise set for percent dark light
    srs.DateDay = Now - 1
    srs.DaySavings = DaySavingsSub(srs.DateDay)
    srs.CalculateSun
    'Tool4B = srs.Sunset ': tool4d = srs.Sunrise
    
    srs.DateDay = Now
    srs.DaySavings = DaySavingsSub(srs.DateDay)
    srs.CalculateSun
    'Tool4 = srs.Sunset
    'tool4db = srs.Sunrise
    
    '#THIS USED IN DARKNESS CODE
    'PC1 = DateDiff("s", Tool4, Tool4D)
    'PC2 = DateDiff("s", Tool4, Now)
    
    
    srs.DateDay = Now + 1
    srs.DaySavings = DaySavingsSub(srs.DateDay)
    srs.CalculateSun
    'Tool4D = srs.Sunrise
    
    'If Tool4 > Now Then Tool4 = Tool4B: Tool4D = tool4db
    
    'if tool4

    srs.DateDay = Now + extra + r
    srs.DaySavings = DaySavingsSub(srs.DateDay)
    
    srs.CalculateSun
    Tool1 = srs.Sunrise
    Tool2 = srs.Sunset
    Tool3 = srs.SolarNoon
    
    
    srt1$ = "": srt2$ = ""
    sst1$ = "": sst2$ = ""
    
    If Now > Tool1 And Now < Tool2 Then
        srt1$ = "> ": srt2$ = " <"
        Else
        sst1$ = "> ": sst2$ = " <"
    End If

    
    tool6$ = Format$(Tool2 - Tool1, "HH:MM:SS")

    If r = 0 Then
    

    
    Label2.Caption = srt1$ + "Sun Rise = " + Mid$(Format$(Tool1, "HH:MM:SSampm"), 1, 8) + srt2$
    Label3.Caption = sst1$ + "Sun Set = " + Mid$(Format$(Tool2, "HH:MM:SSampm"), 1, 8) + sst2$
    Label41.Caption = "Solar Noon = " + Mid$(Format$(Tool3, "HH:MM:SSampm"), 1, 8)

    'PROBLEM
    If 1 = 2 Then
        Call FileInUseDelay("D:\Temp\SunRiseSet.txt", True)
        FR = FreeFile
        Open "D:\Temp\SunRiseSet.txt" For Output Lock Read Write As #FR
        Print #FR, Format$(Tool1, "HH:MM:SS")
        Print #FR, Format$(Tool2, "HH:MM:SS")
        Close #FR
    
    End If
    
    'FR = FreeFile
    'Open "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\SunRiseSet.txt" For Output As #FR
    'Print #FR, Format$(Tool1, "HH:MM:SS")
    'Print #FR, Format$(Tool2, "HH:MM:SS")
    'Close #FR
    
    Label5.Caption = "Difference=" + tool6$
    Label6.Caption = "Diff Yesterday=" + TooL46$

    End If
    

    srs.DateDay = Now + extra + r - 1
    srs.CalculateSun
    Tool41 = srs.Sunrise
    Tool42 = srs.Sunset
    
    srs.DateDay = Now + extra + r
    srs.CalculateSun
    Tool1 = srs.Sunrise
    Tool2 = srs.Sunset
    
    darker1 = DateDiff("S", Tool1, Tool2)
    darker2 = DateDiff("S", Tool41, Tool42)
    If darker1 < darker2 Then darker3 = "Darker " Else darker3 = "Lighter"
    If darker1 = darker2 Then darker3 = "Equal  "

    TTA1 = DateDiff("s", Tool41, Tool42)
    TTA2 = DateDiff("s", Int(Tool41), Int(Tool41) + 1)
    TTA3 = Format((TTA1 / TTA2) * 100, "0.0000")
    
    TooL46$ = Format$(Tool42 - Tool41, "HH:MM:SS") + " - " + TTA3 + "%"



    j41 = Tool41
    j42 = Tool42
    j11 = Tool1
    j22 = Tool2

    j1 = j41 - j42
    j2 = j11 - j22

    tool55$ = Format$(((j41 - j42) - (j11 - j22)), "HH:MM:SS")

    TYU = ((j41 - j42) - (j11 - j22)) - TimeSerial(0, 0, 2)

    tyu2 = Abs(TYU / (TimeSerial(0, 3, 52)) * 100)

    
    TestDate = Int(Now + r)
    TestDate = TestDate - 1 'Have to do this offset
    ggh = Equinox5 - 93
    yy = 1
    ccnt = 0
    For rx = ggh To TestDate
        If Equinox5 = rx Then yy = Not yy
        If Equinox5 + 93 = rx Then yy = Not yy
        If Equinox6 = rx Then yy = Not yy
        If Equinox6 + 93 = rx Then yy = Not yy
        If yy = 1 Then
            ccnt = ccnt + 1
        Else
            ccnt = ccnt - 1
        End If
    Next
    
    tyu2 = (ccnt / 93) * 100
    
    If r = 0 Then DOL1 = Trim(str(Val(Mid$(tool55$, 4, 2)))) + "m" + str(Val(Mid$(tool55$, 7, 2))) + "s " + Trim(darker3)
    If r = 1 Then DOL2 = Trim(str(Val(Mid$(tool55$, 4, 2)))) + "m" + str(Val(Mid$(tool55$, 7, 2))) + "s " + Trim(darker3)
    
    'tool55$ = "DoD=" + Mid$(tool55$, 5) + " Mm:Ss -- " + darker3 + "-- " + Format$(tyu2, "###0.00") + "%"
    tool55$ = "DoD=" + DOL2 + " -- " + Format$(tyu2, "###0.00") + "%"
    '20 Mar 22 Sep Vernal/Autumnal Equinox
    
    EQUINOX_PERCENT = tyu2
    
    Label30.Caption = "SEASON AND EQINOX  " + Format$(EQUINOX_PERCENT, "###0.00") + "%"

    
    If r = 1 Then
        Tool77$ = tool55$
    Else
        Label13.Caption = tool55$
    End If

Next
End If


End Sub



Sub Moon_Code()



tttd$ = ""
'If secstep <> Int(Timer) Then

 '   If Month(monthnow) <> Month(Now) Then
  '      tttd$ = " And It's " + Format$(Now, "mmmm") + " The First"
   ' End If
    
    'monthnow = Now
   On Local Error Resume Next
    ProgressBar1.Value = Now + speedtime
   On Local Error GoTo 0

'End If

'speedtime = speedtime + TimeSerial(0, 20, 0)

If Aggy <> Now Or Tog6 = 1 Then
    Aggy = Now

    'TheDate = ConvertToUT(Now + speedtime)

    TheDate = (Now + speedtime)
    ProgressBar2.Value = (Illum(TheDate)) * 10
    Label10.Caption = Format$(Illum(TheDate), "###.00000") + "%"
    ProgressBar2.ToolTipText = "MOON - Illum -- " + Format$(Illum(TheDate), "###.00000") + "%"
    Call Lab10Sub

    
    If App.title <> "Tidal Information..." Then
        frmPassLock.Label10.Caption = Format$(Illum(TheDate), "###.00000") + "%"
    End If
    
    If X9 + 70 < Now Then X9 = Now + 10
 
    If OLD_X9 <> X9 And OLD_X9 > 0 Then Xcag = 0
    OLD_X9 = X9
    If Now + speedtime > X9 And Xcag <> 2 Then
        Xcag = 2
        Segm9 = 0
    End If
   
   
    If USER_REQUEST_TEST_MOON_CHANGE_STATE_AND_TALK = True Then
        Xcag = 2: Segm9 = 0
        OLD_MoonLabel = ""
    End If
   
    If Xcag = 2 And Segm9 = 0 Then
   
        '--------------------------------
        'THIS BE AROUND WHERE MOON CHANGE
        '--------------------------------
                
        '----------------------------------------------------------------
        'NOT TO RESET THIS VALUE UNLESS TRIGGERED BY HAPPEN NOT WHEN USER
        'FOR THE MOON STATE AND SOUND CHANGE
        '----------------------------------------------------------------
        If USER_REQUEST_TEST_MOON_CHANGE_STATE_AND_TALK = False Then
            'MOON_PER_OT = 0
            
            'ON IT WAY DOWN AND NAUGHT CROSSING
            'LESS THAN 50% THEN I WILL ONLY BE CHANGE TO GOING HIGHER UP
            
            If MOON_PER_OT < 50 Then PER_i2 = 1 Else PER_i2 = -1
        
        
        End If
        USER_REQUEST_TEST_MOON_CHANGE_STATE_AND_TALK = False
        'THIS WILL HELP THE CODE THAT KNOW THE DIRECTION MOON-LIGHT IS GOING
        '-------------------------------------------------------------------
        
        Call WaitWavFinish
            
        If Soff = False Then
            
            Call VolumeFixedSet(20, True)
            
            'Call VoiceMP3Pause
            
            NOT_RUN_YET_HALTER = False
            If TIMER_TO_HALT_RERUN_ > Now Then
                'NOT_RUN_YET_HALTER = True
            End If
            TIMER_TO_HALT_RERUN_ = Now + TimeSerial(1, 0, 0)
            
            If NOT_RUN_YET_HALTER = False Then
                If Wqe3$ <> "" Then
                
                    
                'SEEM A BUG HERE REPEATER FOR A WHILE LONG
                'NOW WORK AROUND QUICK FIXER LATCH NOT SOLVER PROBLEM UNLESS CATCHER
                'NOT TO RUN FOR FIVE TO TEN MINUTE UNLESS USER REQUESTER
                
                'FIXER NOW WONDERFUL
                
                    
                    
                
                    ATidalX.MMControl3.Command = "prev"
                    ATidalX.MMControl3.Command = "Play"
                Else
                    ATidalX.MMControl1.Command = "prev"
                    ATidalX.MMControl1.Command = "Play"
                End If
            End If
        End If
        Segm9 = 10 'Now '+ TimeSerial(0, 0, 1)
        Xcag = 2
    End If

    If Segm9 > 0 And Segm9 < Now Then
        
        Call WaitWavFinish
        
        Segm9 = -1
        'A_ATidalX.DirectSS1.speak Mid$(moonlabel$, 5)
        If Soff = False Then
            
            'GOT DUPE CODE HERE ALSO LINE SAME ABOVE WE KEEP THIS
            'Call VoiceMP3Pause
            'Call SpeakVoice(Mid$(MoonLabel$, 5 + 1))
            
            If OLD_MoonLabel <> MoonLabel$ Then
                Call SpeakVoice(Mid$(MoonLabel$, 1))
            End If
            OLD_MoonLabel = MoonLabel$
          
        End If
        Call moonnext
        Xcag = 2
    End If
End If



End Sub

Sub Tidal_Code()

Zzx = TimeSerial(0, 0, 1) '- 0.000002

'ttd$ = Format$(zzx, "####.00000000")

Do
    wq1 = Timer

    If Wq5 <> Int(wq1) Or Now4 <> Now Then Now4 = Now: wq1 = Timer
    'now4 = Now
    Wq5 = Int(wq1)
    'wq1 = Int(Timer)
    'If wq1 <> wq2 Then pulp = 0
    'wq2 = wq1

    'wq = (1 + (1 - (wq1 - Int(wq1))))
    'pulp = (zzx / (wq))

    wq = (wq1 - Int(wq1)) '+ 0.000001
    Pulp = ((Zzx * 10000000) * (wq)) / 10000000

    'Label14 = Format$(pulp, "####.00000000")
    'Label15 = Format$(zzx, "####.00000000")
    'Label16 = Format$(wq, "####.00000000")
    'If pulp >= 0.00001156 Then
    'wq = wq
    'End If
    'If pulp <= 0 Then
    'wq = wq
    ''pulp = 0.0001
    'End If
'    DoEvents
Loop Until Pulp > 0.0000001 Or 1 = 1



'pulp = pulp + zzx
'- TimeSerial(0, 56, 0)
'now4 = DateSerial(2004, 10, 18) + TimeSerial(14, 34, 0)
'now4 = DateSerial(2004, 10, 18) + TimeSerial(20, 54, 0)
'now4 = DateSerial(2004, 10, 19) + TimeSerial(2, 58, 0)
'now4 = DateSerial(2004, 10, 19) + TimeSerial(15, 19, 0)
'now4 = DateSerial(2004, 10, 20) + TimeSerial(16, 17, 0)
'now4 = DateSerial(2004, 10, 21) + TimeSerial(17, 43, 0)
'now4 = DateSerial(2004, 10, 22) + TimeSerial(19, 23, 0)
'now4 = DateSerial(2004, 10, 23) + TimeSerial(20, 49, 0)
'now4 = DateSerial(2004, 10, 25) + TimeSerial(21, 38, 0)
'now4 = DateSerial(2004, 10, 28) + TimeSerial(23, 36, 0)
'now4 = DateSerial(2004, 10, 31) + TimeSerial(0, 40, 0)
'now4 = DateSerial(2004, 11, 8) + TimeSerial(7, 42, 0)
'now4 = DateSerial(2004, 11, 30) + TimeSerial(0, 49, 0)
'now4 = DateSerial(2004, 12, 31) + TimeSerial(13, 43, 0)
'now4 = DateSerial(2005, 1, 31) + TimeSerial(14, 32, 0)
'now4 = DateSerial(2005, 2, 28) + TimeSerial(13, 33, 0)
'now4 = DateSerial(2005, 3, 31) + TimeSerial(15, 29, 0)
'now4 = DateSerial(2005, 4, 30) + TimeSerial(16, 27, 0)
'now4 = DateSerial(2005, 5, 31) + TimeSerial(18, 53, 0)
'now4 = DateSerial(2005, 6, 30) + TimeSerial(19, 14, 0)
'now4 = DateSerial(2005, 7, 31) + TimeSerial(20, 43, 0)
'now4 = DateSerial(2005, 8, 30) + TimeSerial(21, 39, 0)
'now4 = DateSerial(2005, 9, 30) + TimeSerial(21, 51, 0)
'now4 = DateSerial(2005, 10, 31) + TimeSerial(22, 13, 0)
'now4 = DateSerial(2005, 11, 30) + TimeSerial(22, 16, 0)
'now4 = DateSerial(2005, 12, 31) + TimeSerial(23, 31, 0)


'now4 = now4 + pig
'pig = pig + TimeSerial(0, 0, 1)

Label7.Caption = Format$(Now4 + speedtime, "DDD dd mmm yy")
LABEL_TIME.Caption = "" + LCase$(Format$(Now4 + speedtime, "HH:MM:SSa/p")) + ""



If Tog6 = 1 Then speedtime = speedtime + TimeSerial(0, 30, 0)

B = Now4 + Pulp + (speedtime / 10) '- TimeSerial(0, 17, 0)
'If pulp = 0 Then b = Now + hg
'c = DateDiff("s", a, b)
E = 1 + TimeSerial(0, 50, 28)
'b = a + hg
' a =
'tig = 0.002267
'tig = 0.000712
'tig = -0.000278
Tig = 0.0000386  ''''''''''tide calibration rem-er
d = ((1 + Cos((((B - A) / E) * (pi * (4 + Tig))))) / 2) * 100


'tig = tig - 0.000001''''''''''tide calibration rem-er
'tig = tig + 0.000001''''''''''tide calibration rem-er
W$ = Format$(Tig, "0.00000000")



'd = (b - a) / e
F = B
Label1.Caption = "Tide = " + Format$(d, "###.0000000") + "%  "

'Label10.Refresh

'D:\Wave's\362 Sound Effects
'Set properties needed by MCI to open.
   
 
If Flogship = 0 And Flogship2 < Now Then
    Flogship = 1
End If

If SegM < Now Then

    If Flogship = 0 And Est2 > 0 Then Est2 = 0: Flogship = 1
        
    If Est2 = 3 Then
        
        Call WaitWavFinish
        
        
        Est2 = 0
        If Dag < Now And Soff = False And Toff = False Then
            'ATidalX.DirectSS1.speak "High Tide"

            'Call VoiceMP3Pause
            Call SpeakVoice("High Tide")
         
        End If
    End If
End If

If Est2 = 2 Then
    
    
    Call WaitWavFinish
    
    Est2 = 0
    If Dag < Now And Soff = False And Toff = False Then
        'ATidalX.DirectSS1.speak "Low Tide"
        
        'Call VoiceMP3Pause
        Call SpeakVoice("Low Tide")
      
    End If

End If

If Est2 = 1 Then
    Call WaitWavFinish
    
    Est2 = 0
    If Dag < Now And Soff = False And Toff = False Then
        'ATidalX.DirectSS1.speak "Middle Tide"
        
        'Call VoiceMP3Pause
        Call SpeakVoice("Middle Tide")
   
    End If
End If
 
If (d > 99.99999 Or d < 0.00001) Or (d > 49.999 And d < 50.001) Then
    trigx = 0
    If Climb = 2 And d >= 50 And Upset = 0 Then
        'If apple = 1 Then trigx = 1
        Apple = 1
        Upset = 1
        Est3 = 1
    End If
    If Climb = 1 And d <= 50 And Upset = 0 Then
        'If apple = 1 Then trigx = 1
        Apple = 1
        Upset = 1
        Est3 = 1
    End If
    If Climb = 2 And d < XD(1) And XD(1) > 0 Then
        trigx = 1
        Upset = 0
        Est3 = 3 'high tide
    End If
    If Climb = 1 And d > XD(1) And XD(1) > 0 Then
        trigx = 1
        Upset = 0
        Est3 = 2
    End If
 
    If d < XD(1) And XD(1) > 0 Then Climb = 1
    If d > XD(1) And XD(1) > 0 Then Climb = 2
 
 
    'If timechop < Now Then
    'timechop = Now '+ TimeSerial(0, 0, 1)
    XD(1) = XD(2)
    XD(2) = d
    'End If

 
    XD(1) = d
    
    If trigx = 1 Then
        Est2 = Est3
        qtr = 0
 
        
            
        Call WaitWavFinish
        If Dag < Now And Est3 = 3 And Soff = False And Toff = False Then
            'Tide Wow
            ATidalX.MMControl1.Command = "prev"
            ATidalX.MMControl1.Command = "Play"
            Call WaitWavFinish
        End If
   
    End If
 
End If

End Sub


Private Sub Timer10_Timer()
'tf = SetVolume(SYNTHESIZER, 5)
End Sub

Private Sub Timer11_Timer()
Exit Sub
If Minute(Now) = Timer11Var Then Exit Sub
On Local Error Resume Next

'Open "M:\Temp\Stay Alive Drive.txt" For Output As #1
'Print #1, Now
'Close #1


'Open "N:\Temp\Stay Alive Drive.txt" For Output As #1
'Print #1, Now
'Close #1

Timer11Var = Minute(Now)

End Sub

Private Sub Timer11CPU_Timer()
Exit Sub

Set m_objWMINameSpace = GetObject("winmgmts:")

Dim oCpu As SWbemObject
Set m_objCPUSet = m_objWMINameSpace.InstancesOf("Win32_Processor")
Set oCpu = m_objCPUSet("\\MATT-555ROIDS\root\cimv2:Win32_Processor.DeviceID=""CPU0""")

'Label1 = Str(oCpu.LoadPercentage) + "% Cpu"
ProgressBarCPU.Value = oCpu.LoadPercentage
LabelCPU = str(oCpu.LoadPercentage) + "%"

End Sub


Private Sub Timer12_Timer()
        
Exit Sub
        
        'timer was 70
        'NotePad Font Bigger
        
        If KeyPlus = True Then
            RKeyPlus = 0
            KeyPlus = False
            SndKey = True
        End If
        
        KEYTIME = 21
        KEYTIME = 2
        If RKeyPlus < KEYTIME Then
            RKeyPlus = RKeyPlus + 1
            If RKeyPlus = KEYTIME Then SndKey = False
            Call SendKey(107, 2) 'Key + numpad
        End If
        
End Sub


Private Sub Timer13_Timer()
    
    If 1 = 2 Then
    
        Dim Tiger As Long
        
        ''Nortons 15 Days Reminder
        'Tiger = FindWindow("Sym_ccWebWindow_Class", "Symantec")
        'If GetForegroundWindow = Tiger And Tiger > 0 Then
        '    Call SendKey(Asc("N"), 4) '# Send keys
        '    Call SendKey(9, 0) '# Send keys
        '    Call SendKey(40, 0) '# Send keys
        '    Call SendKey(Asc("O"), 4) '# Send keys
        'End If
        
        Tiger = FindWindow("#32770", "Digital Wave Player")
        If GetForegroundWindow = Tiger And Tiger > 0 Then
            Call SendKey(Asc("Y"), 0) '# Send keys
        End If
        
        Tiger = FindWindow("Digital Wave Player", vbNullString)
        tiger2 = FindWindow(vbNullString, "Transferring the file")
        
        If Tiger > 0 And WindowVisible(Tiger) = False And tiger2 = 0 Then
            WindowVisible(Tiger) = True
        End If
        
       
        
        'Move to Bottom Of Text
        ff = FindWindow(vbNullString, "D:\VB6\VB-NT\00_Best_VB_01\WinAmp Logger\00 Music Info Logger.txt - Notepad2")
        If ff = 0 Then ff = FindWindow(vbNullString, GetSpecialfolder(CSIDL_PERSONAL) + "\03 NotePad Files\00 TOP\Clued Up.txt - Notepad2")
        
        If GetForegroundWindow = ff And ff > 0 And ff <> Ff1 Then
            Ff1 = ff
            Call SendKey(35, 2) '# Send keys
        End If
        
        
        '--Cant do this coz sometimes you have to enter filename
        'ff = FindWindow("#32770", "Make Project")
        'If GetForegroundWindow = ff And ff > 0 Then
        '    Call SendKey(Asc("Y"), 0) '# Send keys
        'End If
        
        ff = FindWindow("#32770", "Microsoft Visual Basic")
        If GetForegroundWindow = ff And ff > 0 Then
            Dim Rect4 As RECT
            
            Hwnd24 = GetWindowRect(ff, Rect4)
            
            If Rect4.Left = 299 And Rect4.Top = 194 And Rect4.Right = 706 And Rect4.Bottom = 484 Then
                Call SendKey(Asc("Y"), 0) '# Send keys
            End If
    
            
        End If
    
    End If '1=2
    
    '-------------------------------
    
    tt1 = FindWindow(vbNullString, "Extreme Lock Build: 2011")
    TT2 = FindWindow(vbNullString, "Tidal Information...")
    
    If OLTt2 <> TT2 Then
    If TestRun = True Then
    If App.title = "Tidal Information..." Then
    Label40.BackColor = Label38.BackColor
    Else
    Label40.BackColor = Label39.BackColor
    frmPassLock.Top = Screen.Height
    End If
    End If
    'If tt1 > 0 Then frmPassLock.Top = Screen.Height - 1000
    
    
    Tx = GetForegroundWindow
    If Tx <> tt1 And Tx <> TT2 Then PopBack = GetForegroundWindow
    If App.title = AppTitle Then Exit Sub
    If AppTitle = "Extreme Lock Build: 2011" Then
      
    ATidalX.SetFocus
    DoEvents
    lngI = SetFocuses(PopBack)
    SetForegroundWindow PopBack
    ingl = Putfocus(PopBack)
    'MsgBox Str(PopBack)
    End If
    
    'AppTitle = App.Title
    
    End If
    OLTt2 = TT2

End Sub




Private Sub Timer4_Timer()
Timer4.Enabled = False
Exit Sub
'If VolumeSet2 < MinVol2Set Then VolumeSet2 = MinVol2Set


If IDLESCHEDULE > Now Or IDLESCHEDULE = 0 Then Exit Sub
IDLESCHEDULE = Now + TimeSerial(0, 30, 0)

Rf = FindWindow(vbNullString, "#00_Schedule_Timer")
If Rf = 0 Then
    'Shell "D:\VB6\VB-NT\00_Best_VB_01\#00_Schedule_Timer\#00_Schedule_Timer-ACER.exe", vbNormalNoFocus
End If


End Sub

Public Sub TimerKeyInterceptSlow_Timer()
TimerKeyInterceptSlow.Enabled = False

'bdf1 = GetAsyncKeyState(115)
'Debug.Print bdf1
'If bdf1 <> 0 Or Vcodez = 115 Then
'If Vcodez = 115 Then
'    WAFastFF = True 'F4 = Fast F
'    Call aa_MainCode.WinAmpCommands(4)
'End If
End Sub

Public Sub TimerKeyIntercept_Timer()

'TestHookinVB = True
'Exit Sub
If 1 = 2 Then
    If ASYNC_Key_Press_Timer.Enabled = True Then
        ASYNC_Key_Press_Timer.Enabled = False
        ASYNC_Key_Press_Timer.Interval = 20
        ASYNC_Key_Press_Timer.Enabled = True
    End If
End If

If DSkeybd_F.List1.ListCount = 0 Then TimerKeyIntercept.Enabled = False: Flag22 = False: Exit Sub

Dim KeyVar

'Exit Sub

KeyVar = DSkeybd_F.List1.List(DSkeybd_F.List1.ListCount - 1)
DSkeybd_F.List1.RemoveItem (DSkeybd_F.List1.ListCount - 1)
DSkeybd_F.Label5 = DSkeybd_F.List1.ListCount

If KeyVar = "" Then Exit Sub

Action = Val(Mid$(KeyVar, 1, 3))
State = Val(Mid$(KeyVar, 5, 3))
vcode = Val(Mid$(KeyVar, 9, 3))
Scode = Val(Mid$(KeyVar, 13, 3))
Vcodez = vcode

Down = Action And 2: Up = Action And 4: Ext = Action And 1
Shift = State And 1


'Call KeyStateTest(Vcode)
'If CapsStat = False Then If (State And 1) = 1 Then Shift = False: State = 0

Ctrl = False: LeftAlt = False: RightAlt = False: AnyAlt = False: NoStateKey = False

If (State And 2) = 2 Then Ctrl = True
If (State And 3) = 3 Then Ctrl = True
If (State And 4) = 4 Then LeftAlt = True: AnyAlt = True
If (State And 5) = 5 Then LeftAlt = True: AnyAlt = True
If (State And 6) = 6 Then RightAlt = True: AnyAlt = True
If (State And 7) = 7 Then RightAlt = True: AnyAlt = True
If AnyAlt = False And Ctrl = False And Shift = False Then NoStateKey = True Else NoStateKey = False
  
ATidalX.Label15.Caption = Action & "-" & State & "-" & vcode & "-" & Scode
  
If Action = 4 Or Action = 5 Then If Bullet = False Then Sluty3 = 0 ': WAFastFF = 0
  
'If IsIDE = True Then
'    Debug.Print "STATE = " + str(State) + "  ---  "; "ACTION = " + str(Action) + "  ---  "; "CTRL = " + str(Ctrl) + "  ---  " + "VCODE = " + str(Vcodez)
'End If

If Action = 4 Or Action = 5 Then Exit Sub ' ---- Key Release after Repeat 5 when shift keys 4 normal
'DPrint action, state, vcode, scode

'Call KeyStateTest(Vcode)

Keyy = Keyy + 1
ATidalX.Label23.Caption = str$(Keyy)




'------------------------------------------
Exit Sub
'------------------------------------------




'------------------------------------------------------------------------------------------
Dim Test_Hwnd3 As Long
Dim Mwnd As Long
ATidalX.FindCursor Nx, Ny
Mwnd = WindowFromPoint(Nx, Ny)
FwnD = GetForegroundWindow



HH$ = GetWindowTitle(GetForegroundWindow)
XX = 0: XX2 = 0: PIP = 0
If InStr(HH$, "IceView (Unregistered)") > 0 Then XX = True: PIP = 1
If InStr(HH$, "Notepad2") > 0 Then XX = True
If InStr(HH$, "Microsoft Excel") > 0 Then XX = True: XX2 = True
If InStr(HH$, "Microsoft Word") > 0 Then XX = True: XX2 = True

'----------------------------------------

If vcode = VK_LSHIFT Or vcode = VK_RSHIFT Then
    'Call KeyStateTest(vcode)
    
    '# Make Caps Go off is shift pressed
    'If CapsStat = True Then Call SendKey(VK_CAPITAL, 0)
    
    '# Make numlock Go off is shift pressed
    'If NumsStat = True Then Call SendKey(VK_NUMLOCK, 0)

End If


'----------------------------------------

If AnyAlt = True And vcode = 119 Then
    'ATidalX.MNU_WinampFadeOut_Click 'F8 = Pause
End If

'----------------------------------------

If vcode = 223 And IsVBrunningForMusick = False Then 'esc
    Tiger = FindWindow("Notepad2", vbNullString)
    HH$ = GetWindowTitle(FindWindow("Notepad2", vbNullString))
    If InStr(HH$, "- Notepad2") > 0 Then XTiger = 1 Else XTiger = 0
        
    edge = 0
    If XTiger = 1 Then
        If GetWindowState(Tiger) = vbMinimized Then
            ShowWindow Tiger, SW_SHOWNORMAL
            edge = 1
        End If
        If GetWindowState(Tiger) = vbMaximized Or GetWindowState(Tiger) = -1 Then
            If edge = 0 Then ShowWindow Tiger, SW_MINIMIZE
        End If
    End If
End If
        
IsVBrunningForMusick2 = IsVBrunningForMusick
'IsVBrunningForMusick2 = 0

'If Vcode > 0 Then Stop

If IsVBrunningForMusick2 = 0 Then

    If vcode = 121 And Sluty3 <> 4 Then
        Sluty3 = 4 'F10 = 'Vol Up
        Bullet = True
        'VolumeLevelSet
        'Debug.Print "111"
    End If
    
    If vcode = 120 And Sluty3 <> 5 Then
        Sluty3 = 5 'F9 = 'Vol Down
        Bullet = True
        'Debug.Print "222"
        'VolumeLevelSet
    End If
    
    If vcode = 119 Then
        
        
        Call aa_MainCode.WinAmpCommands(2)  'F8 = Pause
        'Debug.Print "Hello" + Str(Now)
    End If
    
    If vcode = 118 Then Call aa_MainCode.WinAmpCommands(3)  'F7 = Next Track

End If

'on Mobile Remote
If Actionz = 2 And Statez = 6 And vcode = 50 Then Call aa_MainCode.WinAmpCommands(3)  'F7 = Next Track
If Actionz = 2 And Statez = 6 And vcode = 49 Then Call aa_MainCode.WinAmpCommands(2)  'F8 = Pause???? no 118=f8

'Control Key
If Ctrl = True And vcode = 118 Then Call aa_MainCode.WinAmpCommands(3)  'F7 = Next Track
If Ctrl = True And vcode = 119 Then Call aa_MainCode.WinAmpCommands(2)  'F8 = Pause
If Ctrl = True And vcode = 120 Then Sluty3 = 5    'F9 = 'Vol Down
If Ctrl = True And vcode = 121 Then Sluty3 = 4    'F10 = 'Vol Up

'----------------------------------------

'F12 Timer
If vcode = 123 And Sluty = 0 Then
   ATidalX.ShellKeyLoad.Enabled = True
End If

'Check ----------------- Timer8WinAmpVol_Timer
'If Sluty3 > 0 Then Call WinAmpCommands(Sluty3)
If Sluty > 0 Or Sluty4 > 0 Then Call KeyWindowCheck


'----------------------------------------
'Keycode Vcode stuff -------- LOCK
If (LockDownX = False Or Now - XxG > TimeSerial(2, 0, 0)) Or HittMan = True Then
    If LockDontAllowHotKeyforPassword = False Then
        If vcode = 44 And Ctrl = True Then Call ATidalX.Bollocks
    End If
End If

'Call MallKnocking

'----------------------------------------
'Ctrl R Dictor Up Down
If Ctrl = True And vcode = 82 Then
    Tiger = FindWindow("#32770", "Recording Window")
    If Tiger > 0 And WindowVisible(Tiger) = True Then
        WindowVisible(Tiger) = False
    Else
        WindowVisible(Tiger) = True
        wr = Putfocus(HottMini) '' Still dupious if need this one
        lngI = SetFocuses(HottMini)
        SetForegroundWindow Tiger
    End If
    RecordWindow = Now + TimeSerial(0, 3, 0)
    Exit Sub
End If
    

If Vcodez = 117 Then ' F6 Delay five mins then on not when outlook
    'Me.SetFocus
    'Me.Refresh
    'DoEvents
    
    'Public Sub Label27_Click()
    'Call Label10_Click
    MHDR_NOW_DATE = -1
    
    'Shell "D:\VB6\VB-NT\00_Best_VB_01\Clip_Type\Clip_Type.exe", vbNormalFocus
    
    'AppActivate "Clip Type Form", True
End If
'----------------------------------------

'Play Dialing tone = Pause Off On
If vcode = 117 And 1 = 2 Then
    Go8 = True
    ATidalX.MMControl8.Command = "Pause"
    'If ATidalX.MMControl8.mode = 526 Then
    '    VolumeSet1 = GetVolume(SPEAKER)
    '    tf = SetVolume(SPEAKER, 2)
    'Else
    '    tf = SetVolume(SPEAKER, VolumeSet1)
    'End If
    
    Exit Sub
End If

'----------------------------------------
'Play dialing tone = Reset an Play
If vcode = 117 And Ctrl = True And 1 = 2 Then
    Go8 = True
    VolumeSet1 = GetVolume(SPEAKER)
        
    tf = SetVolume(SPEAKER, 2)

    'If ATidalX.MMControl8.mode = 526 Then
    'ATidalX.MMControl8.Command = "Prev"
    'ATidalX.MMControl8.Command = "Play"
    Exit Sub
End If
'----------------------------------------


'Ctrl M Hott Mini Foreground Window
'----------------------------------------
If App.title = "Tidal Information..." Then
    If Ctrl = True And vcode = 77 And 1 = 2 Then
        Tiger = GetForegroundWindow
        If HottMini > 0 Then
            'ShowWindow HottMini, SW_SHOWNORMAL
            'wr = Putfocus(HottMini)
            'lngI = SetFocuses(HottMini)
            'SetForegroundWindow HottMini
            HottMini = 0
        Else
            'ShowWindow Tiger, SW_MINIMIZE
            HottMini = Tiger
        End If
        Exit Sub
    End If
End If


'Ctrl M Hott Mini Any Window at Cursor
'----------------------------------------
If App.title = "Tidal Information..." Then
    If Ctrl = True And vcode = 77 Then
        Test_Hwnd3 = Mwnd
        title$ = GetParentWindowTitle(Test_Hwnd3, True)
        'errors with this when caps is pressed thinks key is one wanted
        ShowWindow Test_Hwnd3, SW_MINIMIZE
    End If
End If

'Ctrl L Hott Mini All Windows Bar My Favs
'Ctrl L = 76
If App.title = "Tidal Information..." Then
    If Ctrl = True And vcode = 76 Then
        'Shell "D:\VB6\VB-NT\00_Best_VB_01\MiniAllButFavs\MiniAllButFavs.exe"
        Exit Sub
    End If
End If


'SHIFT SCROLL LOCK -- TIME
If App.title = "Tidal Information..." Then
    If State = 1 And vcode = 145 Then
        'SayTime4 = True 'make say time?
        'SayTime3 = True 'make say time? -- FOR THE MALL LOGG ALSO
        Call FORCESAYTIME
        'Call Time_Code
        
        'TIMER_SECONDS20_AFTER_HOUR_TIMER.Enabled = True
            
    
    End If
End If



End Sub

Sub MallKnocking()

Exit Sub

'------------------------------------------
'Mal Knock
'alt scroll lock ALT  makes change mal sound on off
'----------------------------------------
If vcode = 145 And AnyAlt = True Then
    If Menu.SoundMall = vbChecked Then
        Menu.SoundMall = vbUnchecked
        'Call ATidalX.SoundOffOn("OFF")
        ATidalX.Label37.BackColor = ATidalX.Label38.BackColor
    Else
        Menu.SoundMall = vbChecked
        'Call ATidalX.SoundOffOn("ON")
        ATidalX.Label37.BackColor = ATidalX.Label39.BackColor
    End If
End If

'Scroll lock or numb pad *
If NoStateKey = True And (vcode = 106 Or vcode = 145) Then
    If FindWindow("TfrmWinDowse", vbNullString) = 0 Then
        Call Counter1.WriteCounter
        If vcode = 145 And XX2 = 0 Then
            ''Call ATidalX.Sound
        End If
    End If
End If

'Open notepad with alt numpad *
If AnyAlt = True And (vcode = 106 Or vcode = 145) Then
    Call ATidalX.Mnu_Knocker1_Click
End If

'Scroll lock is Vcode=3 when ctrl is pressed
If Ctrl = True And vcode = 3 Then
    Call Counter1.WriteCounter
    'If XX2 = 0 Then 'Call ATidalX.Sound
End If

'Counter 2 - Cough
If NoStateKey = True And vcode = 109 Then
    Call Counter2.WriteCounter
    ''Call ATidalX.Sound
End If
If AnyAlt = True And vcode = 109 Then
    Call ATidalX.Mnu_Knocker2_Click
End If
If Ctrl = True And vcode = 109 Then
    Call Counter2.WriteCounter
    ''Call ATidalX.Sound
End If

'Counter 3 = Phone
If NoStateKey = True And vcode = 12 Then
    Call Counter3.WriteCounter
    'Call ATidalX.Sound
End If
'---
If AnyAlt = True And vcode = 12 Then
    Call ATidalX.Mnu_Knocker3_Click
End If
'---
If Ctrl = True And vcode = 12 Then
    Call Counter3.WriteCounter
    'Call ATidalX.Sound
End If

'Count 4 = Tapp Running key bottom keft next to Ctrl
If NoStateKey = True And vcode = 93 Then
    Call Counter4.WriteCounter
    'Call ATidalX.Sound
End If
If AnyAlt = True And vcode = 93 Then
    Call ATidalX.Mnu_Knocker4_Click
End If
If Ctrl = True And vcode = 93 Then
    Call Counter4.WriteCounter
    'Call ATidalX.Sound
End If

End Sub

Private Sub Timer15_Timer()

    Dim Tiger As Long

    Tiger = FindWindow("#32770", "Recording Window")
    If Tiger > 0 And Tiger <> Xxr5 Then
        RecordWindow = Now + TimeSerial(0, 3, 0)
        Xxr5 = Tiger
    End If
    
    If RecordWindow < Now And RecordWindow > 0 Then
        RecordWindow = 0
        Tiger = FindWindow("#32770", "Recording Window")
        WindowVisible(Tiger) = False
    End If

    Exit Sub

    filespec1 = App.Path + "\#Wave Sounds\DOBFIG4 01.WAV"
    Set F = FS.GetFile((filespec1))
    Adate1 = F.DateLastModified
    Filespec2 = App.Path + "\#Wave Sounds\DOBFIG4 02.WAV"
    Set F = FS.GetFile((Filespec2))
    Adate2 = F.DateLastModified
    If Adate2 > Adate1 Then
    
    ATidalX.MMControl7.Command = "Close"
    
    FS.CopyFile Filespec2, filespec1
    
    ATidalX.MMControl7.Notify = True
    ATidalX.MMControl7.Wait = True
    ATidalX.MMControl7.Shareable = False
    ATidalX.MMControl7.DeviceType = "waveaudio"
    'ATidalX.MMControl7.FileName = App.Path + "\SG_02_04.WAV"
    ATidalX.MMControl7.Filename = App.Path + "\#Wave Sounds\DobFig4 01.WAV"
    ATidalX.MMControl7.Command = "Open"
            
    End If
    
    
    filespec1 = App.Path + "\#Wave Sounds\DOBFIG22 01.WAV"
    Set F = FS.GetFile((filespec1))
    Adate1 = F.DateLastModified
    Filespec2 = App.Path + "\#Wave Sounds\DOBFIG22 02.WAV"
    Set F = FS.GetFile((Filespec2))
    Adate2 = F.DateLastModified
    If Adate2 > Adate1 Then
    
    ATidalX.MMControl9.Command = "Close"
    
    FS.CopyFile Filespec2, filespec1
    
    ATidalX.MMControl9.Notify = True
    ATidalX.MMControl9.Wait = True
    ATidalX.MMControl9.Shareable = False
    ATidalX.MMControl9.DeviceType = "waveaudio"
    ATidalX.MMControl9.Filename = App.Path + "\#Wave Sounds\DobFig22 01.WAV"
    ATidalX.MMControl9.Command = "Open"
            
    
    
    End If
   
   
   

End Sub


Private Sub Timer16_Timer()
    Timer16.Enabled = False
    Exit Sub
    
    ff = FindWindow("#32770", "Confirm File Delete")
    If GetForegroundWindow = ff And ff > 0 Then
        Call SendKey(Asc("Y"), 0) '# Send keys
    End If
    
    ff = FindWindow("#32770", "Confirm Multiple File Delete")
    If GetForegroundWindow = ff And ff > 0 Then
        Call SendKey(Asc("Y"), 0) '# Send keys
    End If
    

End Sub

Private Sub Timer2_Timer()

'i = XxG
            
If DelayOn < Now And DelayOn > 0 Then
    DelayOn = 0
    DelayOn2 = Now + TimeSerial(0, 8, 0)
    WinAmpPowerHwnd = FindWindow("Winamp v1.x", vbNullString)
    MsgResult = SendMessage(WinAmpPowerHwnd, WM_USER, 0, ByVal ipc_isplaying)
    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.
    If MsgResult = 0 Or MsgResult = 3 Then
        MsgResult = SendMessage(WinAmpPowerHwnd, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
    End If
    'start play if stopped
    'If MsgResult = 0 Then
    '    MsgResult = SendMessage(WinAmpPowerHwnd, WM_COMMAND, WINAMP_BUTTON2, ByVal XcNul)
    'End If
End If

If DelayOn2 < Now And DelayOn2 > 0 Then
    DelayOn = 0
    DelayOn2 = 0
    WinAmpPowerHwnd = FindWindow("Winamp v1.x", vbNullString)
    MsgResult = SendMessage(WinAmpPowerHwnd, WM_USER, 0, ByVal ipc_isplaying)
    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.
    If MsgResult = 1 Then
        MsgResult = SendMessage(WinAmpPowerHwnd, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
    End If
    'start play if stopped
    'If MsgResult = 0 Then
    '    MsgResult = SendMessage(WinAmpPowerHwnd, WM_COMMAND, WINAMP_BUTTON2, ByVal XcNul)
    'End If
End If



If FiveMinTime < Now Then
    'caps lock
    'Call SetKeys("CAPSLOCK_OFF")
    FiveMinTime = Now + TimeSerial(0, 5, 0)
End If


If TtYy <> Day(Now) Then
    
    FILENAME_IN_USE_CHECK = GetSpecialfolder(CSIDL_PERSONAL) + "\CallerID\01 KeyStroke_Logger_" + GetComputerName + "-" + GetUserName + ".txt"
    FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4

    Fr5 = FreeFile
    Open FILENAME_IN_USE_CHECK For Append As #Fr5
        ttyu$ = Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS a/p")
        Print #Fr5, "Program Start.."
        Print #Fr5, ttyu$; " -------"
        TtYy = Day(Now)
    Close #Fr5
End If

If TTVol > 0 Then
    TTVol = TTVol - 1
    If TTVol = 1 Then
        TTVol = 0
        Call Mnu_VolMax_Click
    End If
End If

'Eta2 = Eta2 + 1
'If Eta2 < Now Then
'    Eta2 = Now + TimeSerial(0, 20, 0)
'    'Shell "D:\VB6\VB-NT\Time_Sync\Time Sync.exe", vbNormalNoFocus
'End If


'-------------------------------------------------------
Exit Sub
'-------------------------------------------------------

If Rtt <> 802854 And Rtt > 0 Then
On Local Error Resume Next
Open "C:\Program Files\USB Key\SysLock.Bmp" For Input As #1
Rtt = LOF(1)
Close #1
If Rtt = 0 Then Rtt = -1
FS.CopyFile "C:\Program Files\SysLock AG_OR.bmp", "C:\Program Files\USB Key\SysLock.Bmp"
End If


On Local Error GoTo 0

'Call WebPages

'Call AntiVirusMatt
'Only on load now

If LABEL_TIME.BackColor = 0 Then
    Call ColourSet
End If

On Local Error Resume Next

If DelayXYUpdates22 < Now And DelayXYUpdates22 > 0 And Bv$ <> "" Then
    DelayXYUpdates22 = 0
    If X22 + Y22 <> Cosh Then
        freef5 = FreeFile
        Open Bv$ + Bvs$ For Output As #freef5
        Print #freef5, str$(X22)
        Print #freef5, str$(Y22)
        Close #freef5
    Cosh = X22 + Y22
    End If
End If

End Sub

Private Sub Timer3_Timer()


Call MouseDetectMove


Call ProgressLock



End Sub

Sub ProgressLock()

'Call Timer7_Timer

If DateDiff("s", Now, LockSSaver) > -1 Then Call SetLockMax
'MsgBox DateDiff("s", Now, LockSSaver)
    
If DateDiff("s", Now, LockSSaver) > -1 Then
    'time changes make fault
    On Local Error Resume Next
    ProgressBarLock.Value = DateDiff("s", Now, LockSSaver)
    On Local Error GoTo 0
    
    Label9.Caption = Mid$(Format$(TimeSerial(0, 0, ProgressBarLock.Value), "HH:MM:SS"), 4)
    
    If App.title <> "Tidal Information..." Then
        frmPassLock.Label9 = Label9
        frmPassLock.ProgressBarLock.Value = ProgressBarLock.Value
    End If
    
    ProgressBarLock.Refresh
    Label9.Refresh
    On Local Error GoTo 0
End If



'ProgressBarLock.max
        'If GetUserName <> "Matt" Then
        '    LockSSaver = Now + LockSaverDelay
        '    SetLockMax
       ' End If
End Sub


Private Sub ASYNC_Key_Press_Timer_Timer()
'NOT READY YET
'--------------
Exit Sub


ASYNC_Key_Press_Timer.Interval = 1

Dim bdf As Long

kbpress2 = 0
vcode = 0
For i = 5 To 255
    bdf = GetAsyncKeyState(i)
    If bdf < -300 Then vcode = i: Exit For
Next


'If Vcode > 0 And OldKey2 = 0 Then TimeWait2 = 0
'If Vcode > 0 And OldKey2 > 0 Then
'    If Vcode <> OldKey2 Then TimeWait2 = 0
'End If
        
OldKey2 = vcode

'If Vcode > 0 And TimeWait2 > Timer Then
'Vcode = 0
'End If

'If Vcode > 0 Then
'    TimeWait2 = Timer + 0.1
'End If

If vcode > 0 Then
'Debug.Print "BlowUp " + Str(Now)
End If


If vcode > 0 Then
    Label14 = vcode
    Label16 = bdf
    If bdf = -32767 Then Action = 4
    If bdf = -32768 Then Action = 2
    'xx = Callback&(Action, 0, vcode, 0)
    With DSkeybd_F.List1
        'If .ListCount > 500 Then .Clear
        '# Add the new entry at the first position and highlight it
        '.AddItem "   " & Action & Chr$(9) & State & Chr$(9) & Vcode & Chr$(9) & Scode, 0
        .AddItem Format$(Action, "000 ") & Format$(0, "000 ") & Format$(vcode, "000 ") & Format$(0, "000 ")
        .Selected(.NewIndex) = True
        DSkeybd_F.Label5 = .ListCount
        ATidalX.TimerKeyIntercept.Enabled = True
    End With

End If




End Sub


Private Sub Timer5_Timer()

Call SeXxYy
'Call XYLabelsColor

'count nero
'If GetAsyncKeyState(1) Then Count2Check = Val(Menu.Text1.Text)

End Sub

Private Sub XYLabelsColor()
Exit Sub

If Wtool = 0 Or Wtool2 <> Wtool Then
    If Txf = 1 Then
        Txf = 0
        Command1.BackColor = Wtoolcol1
        Command2.BackColor = Wtoolcol2
        LABEL_TIME.BackColor = Wtoolcol8
        Label11.BackColor = Wtoolcol11
        Label12.BackColor = Wtoolcol12
        Label13.BackColor = Wtoolcol13
        Label20.BackColor = Wtoolcol20
        Label22.BackColor = Wtoolcol22
        If Label23.BackColor = (Wtoolcol23 Xor Trgb) Or 16761024 Then
            Label23.BackColor = Wtoolcol23
        End If
        Label26.BackColor = Wtoolcol26
    End If
End If

If Wtool > 0 Then Txf = 1
If Wtool = 1 Then Command1.BackColor = Wtoolcol1 Xor Trgb2
If Wtool = 2 Then Command2.BackColor = Wtoolcol2 Xor Trgb2
If Wtool = 8 Then LABEL_TIME.BackColor = Wtoolcol8 Xor Trgb
If Wtool = 11 Then Label11.BackColor = Wtoolcol11 Xor Trgb
If Wtool = 12 Then Label12.BackColor = Wtoolcol12 Xor Trgb
If Wtool = 13 Then Label13.BackColor = Wtoolcol13 Xor Trgb
If Wtool = 20 Then Label20.BackColor = RGB(180, 180, 180)
If Wtool = 22 Then Label22.BackColor = RGB(180, 180, 180)
If Wtool = 23 Then
    Label23.BackColor = Wtoolcol23 Xor Trgb
End If
If Wtool = 26 Then Label26.BackColor = Wtoolcol26 Xor Trgb

If (Ty <> Ny Or Tx <> Nx) Then
    Wtool = 0
End If

Wtool2 = Wtool



End Sub


Private Sub ColorCycle()

WTrue = WTrue + TW1

If WTrue > 255 Then TW1 = -6: WTrue = WTrue + TW1
If WTrue < 1 Then TW1 = 6: WTrue = WTrue + TW1

HWTrue = HWTrue + TW2

If HWTrue > 255 Then TW2 = -7: HWTrue = HWTrue + TW2
If HWTrue < 1 Then TW2 = 7: HWTrue = HWTrue + TW2

KWTrue = KWTrue + TW3

If KWTrue > 255 Then TW3 = -8: KWTrue = KWTrue + TW3
If KWTrue < 1 Then TW3 = 8: KWTrue = KWTrue + TW3
   
Label4.BackColor = RGB(KWTrue, HWTrue, WTrue)   ' Set drawing color.
Label4.ForeColor = RGB(255 - KWTrue, 255 - HWTrue, 255 - WTrue)
'frmPassLock.Label4.BackColor = RGB(kwtrue, hwtrue, wtrue)   ' Set drawing color.
'frmPassLock.Label4.ForeColor = RGB(255 - kwtrue, 255 - hwtrue, 255 - wtrue)
End Sub


Private Sub WaitWavFinish()
Call WaitVOICEFinish

Do
    DoEvents
    If Wqe1$ <> "" Then wer1 = MMControl1.mode Else wer1 = 525
    If EXIT_FORM_SET = True Then Exit Sub
    If Wqe2$ <> "" Then wer2 = MMControl2.mode Else wer2 = 525
    If EXIT_FORM_SET = True Then Exit Sub
    If Wqe3$ <> "" Then wer3 = MMControl3.mode Else wer3 = 525
    If wer1 = 525 And wer2 = 525 And wer3 = 525 Then Exit Do
    If EXIT_FORM_SET = True Then Exit Sub
    Sleep 50
Loop Until ABC = BCA

End Sub



Private Sub SeXxYy()

Exit Sub

On Error Resume Next

Call SeXxYyPart2(Spiq)



If Spiq = 1 Then Exit Sub

If ATidalX.Left + ATidalX.Width > Screen.Width And LockMove < Timer Then
    ATidalX.Left = Screen.Width - ATidalX.Width
    DoEvents
    '%%%%
    Call ATidalZ.Resize_Code
    'Call Drives.Resize_Code
End If

If Err.Number > 0 Then Exit Sub

TidalMoveDelay2 = ATidalX.Left + ATidalX.Top

If TidalMoveDelay <> TidalMoveDelay2 Then
    DelayXYUpdates22 = Now + TimeSerial(0, 0, 2)
End If

TidalMoveDelay = TidalMoveDelay2

If ME_POSITION_BEEN_SET_ONCE = True Then Exit Sub
'ME_POSITION_BEEN_SET_ONCE = True


If Screen.Width = 19200 And Screen.Height = 15360 And scrwidth_old <> 19200 Then
    Dim Rect4 As RECT
    'ATidalX.Top = xxx2
    'ATidalX.Left = yyy2
    'HwndCTLTask4 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
    HwndCTLTask4 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
    
    Hwnd24 = GetWindowRect(HwndCTLTask4, Rect4)
        
    
    'SEEM NOT USED
    '2ND INSTANCE THIS FOUND 1 AT START
    ATidalX.Top = ((Rect4.Bottom - Rect4.Top) * Screen.TwipsPerPixelX)
    
    ATidalX.Left = Screen.Width - ATidalX.Width
    
End If

If Screen.Width = 19200 And Screen.Height = 15360 Then
    Xxx2 = X22
    Yyy2 = Y22
End If

scrwidth_old = Screen.Width

End Sub


Private Sub SeXxYyPart2(Spiq)

Exit Sub

If Screen.Width <> 19200 Then Exit Sub
    
Dim Abcd
Dim Abcd2
Dim Abcd3
Dim Abcd4
Dim Abcd6
    
Dim Rect2 As RECT
Dim Rect3 As RECT
Dim Rect4 As RECT
Dim Rect6 As RECT
Dim Hwnd22 As Long
Dim Hwnd23 As Long
Dim Hwnd24 As Long
Dim Hwnd26 As Long

Dim Rect22me As RECT
Dim Rect23 As RECT

Hwnd22 = GetWindowRect(HwndCTLTask1, Rect2)
    
If Hwnd22 = 0 Then
    HwndCTLTask2 = FindWindow(vbNullString, "CTLTask")
    Hwnd22 = GetWindowRect(HwndCTLTask2, Rect2)
End If
    
HwndCTLTask3 = FindWindowPart("Caller ID - Ver")
    
If HwndCTLTask3 > 0 Then
    Hwnd23 = GetWindowRect(HwndCTLTask3, Rect3)
End If
    
If Hwnd23 = 0 Then PreRect3 = 0
    
If Hwnd23 > 0 And Menu.CheckLock.Value = vbChecked And PreRect3 <> Rect3.Left Then
    Hwnd26 = GetWindowRect(Me.hWnd, Rect6)
    MoveWindow Me.hWnd, Rect3.Left - (Rect6.Right - Rect6.Left), Rect3.Top, Rect6.Right - Rect6.Left, Rect6.Bottom - Rect6.Top, True
    PreRect3 = Rect3.Left
    Exit Sub
End If
    
Hwnd24 = GetWindowRect(HwndCTLTask4, Rect4)
    
If Hwnd24 = 0 Then
    HwndCTLTask4 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
    Hwnd24 = GetWindowRect(HwndCTLTask4, Rect4)
    'rect5.top
End If
    
Hwnd26 = GetWindowRect(Me.hWnd, Rect6)
    
Spiq = 0

Dim OldRect2 As RECT
Dim OldRect3 As RECT
Dim OldRect4 As RECT
Dim OldRect5 As RECT
Dim OldRect6 As RECT


Abcd2 = MyTypesTheSame(Rect2, OldRect2)
Abcd3 = MyTypesTheSame(Rect3, OldRect3)
Abcd4 = MyTypesTheSame(Rect4, OldRect4)
Abcd6 = MyTypesTheSame(Rect6, OldRect6)

Abcd = Abcd2 And Abcd3 And Abcd4 And Abcd6


OldRect2 = Rect2
OldRect3 = Rect3
OldRect4 = Rect4
OldRect6 = Rect6


'If OldAbcd = abcd Then
'    Spiq = 1: Exit Sub
'End If


If Abcd = False Then
Abcd255 = Abcd2
Abcd355 = Abcd3
Abcd455 = Abcd4
Abcd655 = Abcd6
End If


If Abcd = False Or (GetAsyncKeyState(1) And GetForegroundWindow = Me.hWnd) Then
    LockMove = Timer + 0.1
    Spiq = 1: Exit Sub
End If

If LockMove > Timer Then
Spiq = 1: Exit Sub
End If
   
If LockMove + 0.5 < Timer And Abcd = True Then
    Spiq = 1: Exit Sub
End If

    
    
    
    
    
    
If HwndCTLTask2 > 0 Then 'CTLTask
    
    If Rect6.Top = Rect2.Bottom Then lock44 = 1
    
    If Rect2.Bottom <> TopToolBar1 And _
        Rect2.Bottom < 300 And lock44 = 1 Then
           Rect6.Top = Rect2.Bottom
    End If
    
    If Rect6.Top <> Rect6.Bottom Then lock44 = 0
    
    TopToolBar1 = Rect2.Bottom
    
End If

If 1 = 2 Then
    
        
    If HwndCTLTask3 > 0 Then 'callerID
        
        If Abcd655 = False And Abcd355 = True And Rect6.Right < Rect3.Left Then
        
        'And LockXYForTime < Now Then
        
            LockXY = 0
            
        End If
            
        
        'if inside caller id window then.......
        'ew = 0
        'If Rect6.Right > Rect3.Left Then ew = -1
            
            
        'if below caller id window then .......
        'If Rect6.Top > Rect3.Bottom Then
        '    ew = 1
        '    LockXY = 0
        'End If
        '---------------------
            
        
        If Rect6.Right > Rect3.Left And Rect6.Top < Rect3.Bottom Then
        
            LockXY = 1
                
            If Hwnd23 > 0 Then
                MoveWindow Me.hWnd, Rect3.Left - (Rect6.Right - Rect6.Left), Rect3.Top, Rect6.Right - Rect6.Left, Rect6.Bottom - Rect6.Top, True
            End If
        End If
    End If
    

    'If abcd6 = True And abcd3 = False And LockXY = 1 Then
    
    If LockXY = 1 Then
        
        SoggyPie = True
        LockXYForTime = Now + TimeSerial(0, 0, 2)
        Hwnd23 = GetWindowRect(HwndCTLTask3, Rect3) 'callerid
        If Abcd3 = False Then
            MoveWindow Me.hWnd, Rect3.Left - (Rect6.Right - Rect6.Left), Rect3.Top, Rect6.Right - Rect6.Left, Rect6.Bottom - Rect6.Top, True
        End If
    End If
    
    
    'CallerId
    If HwndCTLTask3 = 0 And OldHwndCTLTask3 > 0 And LockXY = 1 Then
        HwndCTLTask3 = FindWindow("_GD_Sidebar", vbNullString)
        If HwndCTLTask3 > 0 Then
            Hwnd23 = GetWindowRect(HwndCTLTask3, Rect23)
        End If
        If Rect23.Top = 0 Then Hwnd23 = 0
        
        If Hwnd23 > 0 Then
            Hwnd22me = GetWindowRect(Me.hWnd, Rect22me)
            MoveWindow Me.hWnd, Rect23.Left - (Rect22me.Right - Rect22me.Left) + 1, Rect22me.Top, Rect22me.Right - Rect22me.Left, Rect22me.Bottom - Rect22me.Top, True
            
        Else
            ATidalX.Left = Screen.Width - ATidalX.Width
        End If
    End If
        
    OldHwndCTLTask3 = HwndCTLTask3

End If

'Office Shortcut
'If HwndCTLTask4 > 0 And Rect4.Top < 2 Then
'    TidalTopPos = Rect4.Bottom '* Screen.TwipsPerPixelY
'Else
'    TidalTopPos = 0
'End If


'If Abcd655 = False Then


On Local Error Resume Next
If Rect6.Top < Rect4.Bottom Then
    If Rect6.Right < Rect4.Left And LockXY = 0 And Rect4.Top < 2 Then
        ATidalX.Top = 0: ATidalX.Left = 0
    End If

    If Rect6.Right > Rect4.Left And LockXY = 0 And Rect4.Top < 2 Then
        'MoveWindow Me.hWnd, Rect6.Left, Rect4.Bottom, Rect6.Right - Rect6.Left, Rect6.Bottom - Rect6.Top, True
        'atidalx.Top=
        
        HwndCTLTask4 = FindWindow(vbNullString, "WindowsFormsParkingWindow")

        Hwnd24 = GetWindowRect(HwndCTLTask4, Rect4)
        
        ATidalX.Top = (Rect4.Bottom * Screen.TwipsPerPixelX)
        ATidalX.Left = Screen.Width - ATidalX.Width
        Call ATidalZ.Resize_Code
        Call Drives.Resize_Code

    End If
End If


'
'hwnd24 ' Office     '  and rect4
If Hwnd24 = 0 Then ATidalX.Top = 0
Hwnd24 = GetWindowRect(Me.hWnd, Rect6)


If (Rect6.Top = LockXY2 Or ((Rect6.Top + Rect4.Bottom) < 20)) Then
        If Rect4.Bottom > 0 And Rect4.Top < 2 Then
            'If Rect4.Bottom > 0 Then rect6t = Rect4.Bottom Else rect6t = 0
            'MoveWindow Me.hwnd, Rect6.Left, rect6t, Rect6.Right - Rect6.Left, Rect6.Bottom - Rect6.Top, True
            
            If Rect4.Bottom > 0 Then
                Me.Top = Screen.TwipsPerPixelY * Rect4.Bottom
            End If
            
            If Rect4.Bottom <= 0 Then
                Me.Top = 0
            End If
            
            Hwnd24 = GetWindowRect(Me.hWnd, Rect6)
            LockXY2 = Rect6.Top
        End If
  

End If


If (Rect6.Top = LockXY2 And Rect4.Top > Rect6.Top) Or Rect6.Top < 0 Then
    If Rect4.Top > 50 Then
'        rect6t = 0
        'MoveWindow Me.hwnd, Rect6.Left, 0, Rect6.Right - Rect6.Left, Rect6.Bottom - Rect6.Top, True
        Me.Top = 0
        Hwnd24 = GetWindowRect(Me.hWnd, Rect6)
        LockXY2 = Rect6.Top
    End If
End If



'End If
    
'If Abcd455 = False And Abcd655 = True Then
'If Rect4.Top < 2 Then
'    MoveWindow Me.hWnd, Rect6.Left, Rect4.Bottom, Rect6.Right - Rect6.Left, Rect6.Bottom - Rect6.Top, True
'End If
'End If

    
    
    

X22 = ATidalX.Top
Y22 = ATidalX.Left




End Sub

' Return True if these structures contain the same data.
Private Function MyTypesTheSame(A As RECT, B As RECT) _
    As Boolean
'Robert Heinig points out that this will not work if the UDTs contain
'strings or object references. In that case, this method would
'compare only the addresses of the objects not the objects themselves

Dim i As Integer
Dim a_bytes() As Byte
Dim b_bytes() As Byte
Dim different As Boolean

    ' Copy the data into arrays of bytes.
    ReDim a_bytes(1 To Len(A))
    ReDim b_bytes(1 To Len(B))

    CopyMemory a_bytes(1), A, Len(A)
    CopyMemory b_bytes(1), B, Len(B)

    ' Compare the bytes.
    For i = 1 To Len(A)
        If a_bytes(i) <> b_bytes(i) Then
            MyTypesTheSame = False
            Exit Function
        End If
    Next i

    MyTypesTheSame = True
End Function




Private Sub NortonsFileSpec(Adate1 As Date, Adate2 As Date, Directory1$)
    
    

dea$ = ""

For r = 1 To Dir2.ListCount - 1

    Da$ = Dir2.List(r)
    W = 1
    
    For R2 = 1 To 10
        W = InStr(W + 1, Da$, "\")
        If W = 0 Then Exit For
        w2 = W
    Next

    Da$ = Mid$(Dir2.List(r), w2 + 1)

    
    bfc$ = Trim$(str$(Year(Now) - 1))
    bfc2$ = Trim$(str$(Year(Now)))
    
    If Mid$(Da$, 1, 4) = bfc$ Or Mid$(Da$, 1, 4) = bfc2$ Then
        dea$ = Da$
        LockNorts = NortonAdate2
    End If

Next
    
    
Directory1$ = dea$
    
Dim F, FileSpec
   
If Directory1$ <> "" Then
    FileSpec = drive2$ + ":\Program Files\Common Files\Symantec Shared\VirusDefs\" + Directory1$ + "\TECHNOTE.TXT"
    On Local Error Resume Next
'    Set F = fs.GetFile((FileSpec))
    FOLDERSPEC = drive2$ + ":\Program Files\Common Files\Symantec Shared\VirusDefs\" + Directory1$
    Set F = FS.GetFolder((FOLDERSPEC))
    Adate1 = F.DateLastModified
End If
'Temp take out when working again for nortons 360
Exit Sub

FileSpec = drive2$ + ":\Program Files\Common Files\Symantec Shared\VirusDefs\" + Directory1$ + "\VSCANmsx.DAT" 'definfo.dat"
'FileSpec = drive2$ + ":\Program Files\Common Files\Symantec Shared\VirusDefs\" + Directory1$ + "\VIRSCANT.DAT" 'definfo.dat"
Set F = FS.GetFile((FileSpec))
'Adate2 = F.DateLastModified
Adate2 = Adate1
End Sub




Private Sub AntiVirusMatt()
    
Exit Sub
    
'FindFileInfo("R:\PROGRA~1\COMMON~1\SYMANT~1\VIRUSD~1")

If TestTrig = True Then GoTo skip2

If drive2$ = "" Then Exit Sub

If DateDiff("h", LockNorts, Now) < 24 Then Exit Sub

Call NortonsFileSpec(NortonAdate1, NortonAdate2, dea$)

If LockNorts = NortonAdate2 Then Exit Sub

skip2:

'Call NortonsFileSpec(NortonAdate1, NortonAdate2, dea$)

Dir2.Refresh

Call NortonsFileSpec(NortonAdate1, NortonAdate2, dea$)

If NortonAdate1 = 0 Then Exit Sub

aw$ = BeSet1$

If aw$ <> "" Then
    DeaDab$ = Mid$(aw$, InStr(aw$, """>") + 2)
    w2 = InStr(DeaDab, ":")
    DeaDab$ = Mid$(DeaDab$, 1, w2 + 6)
End If


'dea2 = DateSerial(Val(Mid$(dea$, 1, 4)), Val(Mid$(dea$, 5, 2)), Val(Mid$(dea$, 7, 2)))

dea3$ = Format$(NortonAdate1, "DDD DD MMMM YYYY")
dea4$ = Format$(NortonAdate2, "DDD DD MMMM YYYY HH:MM:SSa/p")






aw$ = BeSet2$

w3 = InStr(aw$, """>")

If aw$ <> "" Then
    deadab2$ = Mid$(aw$, w3 + 2, 3)
End If

If dea3$ <> DeaDab$ Or Mid$(dea$, 10) <> deadab2$ Or Que2 = 0 Or TestTrig = True Then
    
    Que2 = 1
    
    
    'Open drive3$ + SigPath$ + "Signature2.txt" For Output As #1

    'Print #1, ""
    
    'If CompName$ = "MAGICRAM2EODUR" Then Print #1, "Matthew..."
    'If CompName$ = "EDDIESCIENTIST" Then Print #1, "Eddie..."
    'If CompName$ = "NASA1234567890" Then Print #1, "Marianne..."
    'If CompName$ = "MEACHELLE12345" Then Print #1, "Meachelle..."
    
    'Print #1, ""

    FEG = FreeFile
    On Local Error Resume Next
    Open drive2$ + ":\Program Files\Norton AntiVirus\readme.txt" For Input As #FEG
    If Err.Number > 0 Then
        Open drive2$ + ":\Program Files\Norton Internet Security\Norton AntiVirus\readme.txt" For Input As #FEG
    End If
    
    
    version2$ = ""
    Do
        Line Input #FEG, wsd$
    
        If InStr(wsd$, "2004") And version2$ = "" Then version2$ = "2004"
        If InStr(wsd$, "2005") And Val(version2$) < 2005 Then version2$ = "2005"
        If InStr(wsd$, "2006") And Val(version2$) < 2006 Then version2$ = "2006"
        If InStr(wsd$, "2007") And Val(version2$) < 2007 Then version2$ = "2007"
        If InStr(wsd$, "2008") And Val(version2$) < 2008 Then version2$ = "2008"
    
    Loop Until EOF(FEG)
    Close #FEG

    
    If version2$ = "" Then
        If CompName$ = "MAGICRAM2EODUR" Then
            version2$ = "2005"
        Else
            version2$ = "2004"
        End If
    End If

    cols1$ = "<font color=""#66FF66"">"
    cols2$ = "<font color=""#FFFFFF"">"
    cols3$ = "<font color=""#FFFF90"">"

    
    gbh1$ = "This System Check's with Nortons " + cols1$ + version2$ + "</font>"
    'gbh2$ = "--- Norton Anti-Virus " + cols$ + version2$ + " ---"
    dea3$ = ""
    gbh3$ = "Using Definitions Set Date: " ' + cols1$ + dea3$ + "</font> --"
    BeSet1$ = gbh3
    
    txt1$ = "This System Check's with Nortons " + version2$ + vbCrLf
    txt1$ = txt1$ + gbh3$ + dea4$
    
    '"-- Virus Definitions Date: " + cols$ + Dea3$ + "</font> --"
    'If InStr(dea4$, dea3$) > 0 Then
    '    Tdea4$ = cols3$ + "the Same Day at " + cols1$ + Mid$(dea4$, 24)
    'Else
        tdea4$ = cols1$ + dea4$ + "</font>" 'Mid$(dea4$, 1, 20) + " " + cols3$ + Mid$(dea4$, 24)
    'End If
    'beset2$ = "- Rev. " + cols1$ + Mid$(dea$, 10) + cols2$ + " Uploaded on " + cols1$ + Tdea4$ + "</font> -"
    gbh4$ = BeSet1$
    'gbh5$ = beset2$
    gbh5$ = tdea4$
    BeSet2$ = gbh5$
    'gbh7$ = "--- Uploaded At: ---"
    'gbh8$ = "--- " + cols$ + Dea4$ + "</font> ---"
    
    
    
    
    'Print #1, gbh6$
    'Print #1, gbh1$
    'Print #1,
    'Print #1, gbh7$
    'Print #1, gbh2$
    'Print #1, gbh3$
    'Print #1, gbh5$
    'Print #1, gbh7$
    'Print #1, gbh8$
    
    'Print #1, gbh6$
    'Print #1,

    'Close #1
    
    'Open drive3$ + sigPath$+"Signature2.txt" For Input As #1
    Open drive3$ + SigPath$ + "Signature-Nortons-txt.txt" For Output As #3
        Print #3, txt1$
    Close #3
    
    Open drive3$ + SigPath$ + "Signature-Nortons.html" For Input As #1
    Open drive3$ + SigPath$ + "Signature-Nortons-In.html" For Output As #2
    Do
    Line Input #1, ll$
    Print #2, ll$
    Loop Until EOF(1)
    Close #1, #2
    
    Open drive3$ + SigPath$ + "Signature-Nortons-In.html" For Input As #1
    Open drive3$ + SigPath$ + "Signature-Nortons.html" For Output As #2
    'Open drive3$ + sigPath$+"Signature6.html" For Output As #3
    
    'gbh3$ = "Using Virus Definitions Date: " + Dea3$
    
    Do
        Line Input #1, one$
        If one$ = "<!--End---->" Then CubeFace = 0
        If one$ = "<!--Start-->" Then
            Print #2, one$
            'If CompName$ = "MAGICRAM2EODUR" Then Print #2, "Matthew..."
            'If CompName$ = "EDDIESCIENTIST" Then Print #2, "Eddie..."
            'If CompName$ = "NASA1234567890" Then Print #2, "Marianne..."
            'If CompName$ = "MEACHELLE12345" Then Print #2, "Meachelle..."
            Print #2, "<center>"
            Print #2, gbh1$ + "<br>"
            'Print #2, gbh2$ + "<br>"
            Print #2, gbh3$ + "<br>"
            Print #2, gbh5$
            'Print #2, gbh7$ + "<br>"
            'Print #2, gbh8$
            CubeFace = 1
        End If
        
        If CubeFace = 0 Then Print #2, one$
    
    Loop Until EOF(1)
    Close #1, #2 ', #3
    Kill drive3$ + SigPath$ + "Signature-Nortons-In.html"
    
    'Open drive3$ + SigPath$ + "Signature1.html" For Input As #1
    'Open drive3$ + SigPath$ + "Signature6.html" For Output As #3
    'Do
    '    Line Input #1, one$
    '    Print #3, one$
    'Loop Until EOF(1)
   '
   ' Close #1, #3
   '
   '
   ' Open drive3$ + SigPath$ + "Signatures\Signature2.txt" For Input As #1
   ' Open drive3$ + SigPath$ + "Signatures\Signature3.txt" For Output As #2

   ' Do
   '     Line Input #1, adf$
   '     Print #2, adf$
   ' Loop Until EOF(1)
    
    
    '"MATT-555ROIDS"
    
    'If CompName$ = "MATT-555ROIDS" Then
    Print #2, "http://MatthewLan.Com"
    'If CompName$ = "555ROIDSGOLDRIM" Then Print #2, "http://MatthewLan.Com"
    'If CompName$ = "MAGICRAM2EODUR" Then Print #2, "http://MatthewLan.Com"
    'If CompName$ = "EDDIESCIENTIST" Then Print #2, "http://uk.geocities.com/elancaster@btinternet.com/"
    'If CompName$ = "NASA1234567890" Then Print #2, "http://uk.geocities.com/marianne.farley@btinternet.com/"
    'If CompName$ = "MEACHELLE12345" Then Print #2, "http://uk.geocities.com/meachelle.thwaites@btinternet.com/"

    Print #2,
    
    Close #1, #2

    'Outgoing mail is certified Virus Free.
    'Checked by AVG anti-virus system (http://www.grisoft.com).
    'Version: 6.0.786 / Virus Database: 532 - Release Date: 29/10/2004

End If


End Sub

Private Sub moonnext()

R2D2 = 0
Call moon2

X10 = X9

For R2D2 = 1 To 14

Call moon2

Label12.Caption = Format(X10, "ddmmm hh:mm:ss")
Label26.Caption = Format(X9, "ddmmm hh:mm:ss")

If App.title <> "Tidal Information..." Then
    frmPassLock.Label12.Caption = Format(X10, "ddmmm hh:mm:ss")
    frmPassLock.Label26.Caption = Format(X9, "ddmmm hh:mm:ss")
End If

'If tog1 = 1 Then Label12.Caption = Format(x9, "ddmmm hh:mm:ss")

'If tog1 = 0 Then Label12.Caption = Format$(Illum(x9), "000.00000") + "%"
'est5$ = Format$(Illum(x7), "000.00000") + "%"
If X10 <> X9 Then X10 = X9: Exit For

Next


R2D2 = 0
Call moon2
ProgressBar1.Min = 0
ProgressBar1.Max = X9

On Error Resume Next
' ERROR HERE IF DEBUG IDE WAIT LONG
ProgressBar1.Min = X10
' ---------------------------------


If App.title <> "Tidal Information..." Then
    frmPassLock.ProgressBar1.Min = 0
    frmPassLock.ProgressBar1.Max = X9
    frmPassLock.ProgressBar1.Min = X10
End If
End Sub

Sub moon2()

'-----------------------------------
'WHY R2D2 SET TO A VAULE OF 1 OFFSET
'REASON THE CODE THAT CALL HERE IS CHECK FOR NEXT MOON STATE AS WELL AS FIRST
'-----------------------------------

'TheDate = Now '+ speedtime - R2D2
'TheDate = Now - R2D2


TheDate = ConvertToUT(Now) - R2D2


'x11 = PreviousLastQuarter(TheDate)
x1 = NextNewMoon(TheDate)

'x33 = PreviousNewMoon(TheDate)
X3 = NextFirstQuarter(TheDate)

'x55 = PreviousFirstQuarter(TheDate)
x5 = NextFullMoon(TheDate)

'x77 = PreviousFullMoon(TheDate)
X7 = NextLastQuarter(TheDate)

x2 = DateAdd("s", -345931, x1)
X4 = DateAdd("s", -345931, X3)
X6 = DateAdd("s", -345931, x5)
X8 = DateAdd("s", -345931, X7)

'x4 = DateAdd("s", DateDiff("s", x1 + (345931), x1) / 2, x1)
'x6 = DateAdd("s", DateDiff("s", x1 + (345931), x1) / 2, x1)
'x8 = DateAdd("s", DateDiff("s", x1 + (345931), x1) / 2, x1)
'x4 = DateAdd("s", DateDiff("s", x5, x3) / 2, x5)
'x6 = DateAdd("s", DateDiff("s", x7, x5) / 2, x7)
'x8 = DateAdd("s", DateDiff("s", x1, x7) / 2, x1)

If x2 < TheDate Then x2 = TheDate + 50
If X4 < TheDate Then X4 = TheDate + 50
If X6 < TheDate Then X6 = TheDate + 50
If X8 < TheDate Then X8 = TheDate + 50

X9 = ConvertToUT(Now) + 400

If x1 < x2 Then X9 = x1: Asd = 1
If x2 < X3 And x2 < X9 Then X9 = x2: Asd = 2
If X3 < X4 And X3 < X9 Then X9 = X3: Asd = 3
If X4 < x5 And X4 < X9 Then X9 = X4: Asd = 4
If x5 < X6 And x5 < X9 Then X9 = x5: Asd = 5
If X6 < X7 And X6 < X9 Then X9 = X6: Asd = 6
If X7 < X8 And X7 < X9 Then X9 = X7: Asd = 7
If X8 < x1 And X8 < X9 Then X9 = X8: Asd = 8

'Label12.Caption = Format(x9, "ddmmm hh:mm:ss")

If Asd = 1 Then Label11.Caption = "Next New Moon": MoonH = 0
If Asd = 4 Then Label11.Caption = "Next Waxing Crescent": MoonH = 1
If Asd = 3 Then Label11.Caption = "Next First Quarter": MoonH = 2
If Asd = 6 Then Label11.Caption = "Next Waxing Gibbous": MoonH = 3
If Asd = 5 Then
    Label11.Caption = "Next Full Moon": MoonH = 4
End If
If Asd = 8 Then Label11.Caption = "Next Waning Gibbous": MoonH = 5
If Asd = 7 Then Label11.Caption = "Next Last Quarter": MoonH = 6
If Asd = 2 Then Label11.Caption = "Next Waning Crescent": MoonH = 7

MoonLabel$ = Label11.Caption

If Asd = 1 Then label12s$ = "Now Waning Crescent"
If Asd = 2 Then label12s$ = "Now Last Quarter"
If Asd = 3 Then label12s$ = "Now Waxing Crescent"
If Asd = 4 Then label12s$ = "Now New Moon"
If Asd = 5 Then label12s$ = "Now Waxing Gibbous"
If Asd = 6 Then label12s$ = "Now First Quarter"
If Asd = 7 Then label12s$ = "Now Waning Gibbous"
If Asd = 8 Then label12s$ = "Now Full Moon"

Label24.Caption = label12s$

End Sub

'Public Sub Label11_Click()

Public Sub Label24_Click()

MNU_SPEECH_OFF.Checked = False
Call SPEECH_OFF_SETTER

Call MNU_VOX_LIST_Click

VOXLIST_NOT_ACTIVE = 1000
            
'wd$ = wd$ + "Currently      : " + Label24.Caption + vbCrLf
'wd$ = wd$ + "Soon   : " + Label11.Caption + vbCrLf
Beep

XA = Format(X10, "dd")
If Val(Mid(XA, Len(XA))) = 0 Then TALK_1ST_2ND_3RD_4TH = "th"
If Val(Mid(XA, Len(XA))) = 1 Then TALK_1ST_2ND_3RD_4TH = "st"
If Val(Mid(XA, Len(XA))) = 2 Then TALK_1ST_2ND_3RD_4TH = "nd"
If Val(Mid(XA, Len(XA))) = 3 Then TALK_1ST_2ND_3RD_4TH = "rd"
If Val(Mid(XA, Len(XA))) > 3 Then TALK_1ST_2ND_3RD_4TH = "th"
XM = TALK_1ST_2ND_3RD_4TH

'wd1$ = Label24.Caption + " " + Format(X10, "dd mmmm hh:mm am/pm")
'wd2$ = Label11.Caption + " " + Format(x9, "dd mmmm hh:mm am/pm")

wd1$ = Label11.Caption + " " + Format(X10, "dd") + "" + XM + " of " + Format(X9, "mmmm") + ", " + Format(X9, "hh:mm am/pm")


Call SpeakVoice(wd1$)


End Sub


Public Sub Label11_Click()

MNU_SPEECH_OFF.Checked = False
Call SPEECH_OFF_SETTER

Call MNU_VOX_LIST_Click

VOXLIST_NOT_ACTIVE = 1000

XA = Format(X10, "dd")
If Val(Mid(XA, Len(XA))) = 0 Then TALK_1ST_2ND_3RD_4TH = "th"
If Val(Mid(XA, Len(XA))) = 1 Then TALK_1ST_2ND_3RD_4TH = "st"
If Val(Mid(XA, Len(XA))) = 2 Then TALK_1ST_2ND_3RD_4TH = "nd"
If Val(Mid(XA, Len(XA))) = 3 Then TALK_1ST_2ND_3RD_4TH = "rd"
If Val(Mid(XA, Len(XA))) > 3 Then TALK_1ST_2ND_3RD_4TH = "th"
XM = TALK_1ST_2ND_3RD_4TH

' wd1$ = Label24.Caption + " " + Format(X10, "dd mmmm hh:mm am/pm")
wd1$ = Label11.Caption + " " + Format(X10, "dd") + "" + XM + " of " + Format(X9, "mmmm") + ", " + Format(X9, "hh:mm am/pm")

XA = Format(X9, "dd")
If Val(Mid(XA, Len(XA))) = 0 Then TALK_1ST_2ND_3RD_4TH = "th"
If Val(Mid(XA, Len(XA))) = 1 Then TALK_1ST_2ND_3RD_4TH = "st"
If Val(Mid(XA, Len(XA))) = 2 Then TALK_1ST_2ND_3RD_4TH = "nd"
If Val(Mid(XA, Len(XA))) = 3 Then TALK_1ST_2ND_3RD_4TH = "rd"
If Val(Mid(XA, Len(XA))) > 3 Then TALK_1ST_2ND_3RD_4TH = "th"
XM = TALK_1ST_2ND_3RD_4TH

wd2$ = Label11.Caption + " " + Format(X9, "dd") + "" + XM + " of " + Format(X9, "mmmm") + ", " + Format(X9, "hh:mm am/pm")

'Call VoiceWait
'Call VoiceMP3Pause
FORCE_TALKER_GO_WITHOUT_CHECKER_DUPE_IGNOR = True
Call SpeakVoice(wd2$)

Exit Sub


Tog4 = Tog4 Xor 1

If Tog4 = 0 Then
If Asd = 1 Then label11s$ = "Now Waning Crescent"
If Asd = 2 Then label11s$ = "Now Last Quarter"
If Asd = 3 Then label11s$ = "Now Waxing Crescent"
If Asd = 4 Then label11s$ = "Now New Moon"
If Asd = 5 Then label11s$ = "Now Waxing Gibbous"
If Asd = 6 Then label11s$ = "Now First Quarter"
If Asd = 7 Then label11s$ = "Now Waning Gibbous"
If Asd = 8 Then label11s$ = "Now Full Moon"

Label11.Caption = label11s$
End If

If Tog4 = 1 Then Label11.Caption = MoonLabel$


'If tog1 = 1 Then
'If tog4 = 1 Then Label12.Caption = Format(x9, "ddmmm hh:mm:ss")
'If tog4 = 0 Then Label12.Caption = Format(x10, "ddmmm hh:mm:ss")
'End If

'If tog1 = 0 Then
'If tog4 = 1 Then Label12.Caption = Format$(Illum(x9), "###.00000") + "%"
'If tog4 = 0 Then Label12.Caption = Format$(Illum(x10), "###.00000") + "%"
'End If


Label12.Caption = Format(X9, "ddmmm hh:mm:ss")
frmPassLock.Label12.Caption = Format(X9, "ddmmm hh:mm:ss")
Label26.Caption = Format(X10, "ddmmm hh:mm:ss")
frmPassLock.Label26.Caption = Format(X10, "ddmmm hh:mm:ss")

End Sub


Private Sub Label12_Click()


Call MNU_VOX_LIST_Click

VOXLIST_NOT_ACTIVE = 1000

'MOON 8 PHASE NAME TALK

Call Mnu_SayMoon_Click

Exit Sub

Tog1 = Tog1 Xor 1


If Tog1 = 1 Then
If Tog4 = 1 Then Label12.Caption = Format(X9, "ddmmm hh:mm:ss")
If Tog4 = 0 Then Label12.Caption = Format(X10, "ddmmm hh:mm:ss")
If Tog4 = 1 Then frmPassLock.Label12.Caption = Format(X9, "ddmmm hh:mm:ss")
If Tog4 = 0 Then frmPassLock.Label12.Caption = Format(X10, "ddmmm hh:mm:ss")
End If

If Tog1 = 0 Then
If Tog4 = 1 Then Label12.Caption = Format$(Illum(X9), "###.00000") + "%"
If Tog4 = 0 Then Label12.Caption = Format$(Illum(X10), "###.00000") + "%"
If Tog4 = 1 Then frmPassLock.Label12.Caption = Format$(Illum(X9), "###.00000") + "%"
If Tog4 = 0 Then frmPassLock.Label12.Caption = Format$(Illum(X10), "###.00000") + "%"
End If




End Sub




Public Sub FindCursor(x, y)

Dim P As POINTAPI

GetCursorPos P
'   return x and y co-ordinate
x = P.x ' / GetSystemMetrics(0) * Screen.Width
'   for current cursor position
y = P.y '/ GetSystemMetrics(1) * Screen.Height

End Sub

Sub WinonPoint()

Exit Sub

'Get the window under the cursor
If IsIDE = False Then Exit Sub
FindCursor Nx, Ny
Mwnd = WindowFromPoint(Nx, Ny)
WW$ = GetWindowTitle(Mwnd)
XX = 0
If InStr(WW$, "Extreme Lock Build: 2011") > 0 Then XX = 1
If InStr(WW$, "Tidal_Info") > 0 Then XX = 1
If InStr(WW$, "Tidal Info") > 0 Then XX = 1
If XX = 1 And InStr(WW$, "(Code)") > 0 Then XX = 0

'if not in ide then load but not more than once
If TestKeyLoggOff = False And TestKeyLoggOff <> -2 Then
    Call DSkeybd_F.Command1_Click
    TestKeyLoggOff = -2
End If


'TestKeyLoggOff = True
'TestKeyLoggOff = False
If TestKeyLoggOff = True Then
    If XX = 1 And DSkeybd_F.IsHook = False Then
        Call DSkeybd_F.Command1_Click
    End If
    If XX = 0 And DSkeybd_F.IsHook = True Then
        Call DSkeybd_F.UnHookCommand_Click
    End If
End If

A = Me.hWnd

End Sub



Private Sub MouseDetectMove()

FindCursor Nx, Ny

'This Will Happen to Mouse If User Is Logged off
If Nx = 0 And Ny = 0 Then
    LockSSaver = Now + LockSaverDelay
    'Winamp Video
    Call SetLockMax
    Call ProgressLock
End If

If Nx = 0 Or Ny = 0 Then Exit Sub


If Quake = 0 Then
    If (Nx <> Xmp5 Or Ny <> Ymp5) And (Nx <> ScreenWidth And Ny <> ScreenHeight) Then
        
Call WinonPoint
        
        
        'If Sluty3 = 0 Then StandBy = Now + TimeSerial(0, 0, 3)

'        If Tim < Now Then
 '           List1.AddItem Format$(Now, "HH:MM:SS") + Str(nx) + " -- " + Str(ny)
  '          List1.ListIndex = List1.ListCount - 1
   '         Tim = Now + TimeSerial(0, 0, 1)
    '    End If

        Mousey = Mousey + Sqr(Abs(Xmp5 - Nx) ^ 2 + Abs(Ymp5 - Ny) ^ 2)
        
        'Mousey2 = Sqr(Abs(Xmp5 - Nx) ^ 2 + Abs(Ymp5 - Ny) ^ 2)
        'If Mousey2 > 15 Then Mouse10CLicks = True Else Mouse10CLicks = False
        
        MouseClicks = Sqr(Abs(Xmp5 - Nx) ^ 2 + Abs(Ymp5 - Ny) ^ 2)
                    
        If MouseFirstRun = 0 Then
            MouseFirstRun = 1
            Mousey = 0
        End If
        
        If MouseClicks > 3 Then
            Label21.Caption = str$(Mousey)
        End If
        
        If Easyride > Now Then Label23.BackColor = Wtoolcol23

        Startmouse = 1
        If DelayXYUpdates22 > Now Then DelayXYUpdates22 = Now - 1
    
        Xmp5 = Nx: Ymp5 = Ny
    
    End If
End If

If (Quake = 1 And (Nx <> ScreenWidth And Ny <> ScreenHeight)) Or (Easyride < Now And Quake = 1) Then
    If drive2$ <> "" And Hidecursor2 = False Then
        If CompName$ <> "MEACHELLE12345" Then
            SetCursorPos Xmp6, Ymp6
            Xmp5 = Xmp6: Ymp5 = Ymp6
            Quake = 0
            Easyride = Now - 1
        End If
    End If
End If

If Easyride < Now Then
    If Label22.BackColor = &HFFC0C0 Then Label22.BackColor = &HFF0000
    If Label23.BackColor = &HFFC0C0 Then
        Label23.BackColor = Wtoolcol23
    End If
End If

End Sub


'Private Sub MouseDetectMove()
Public Sub SlowKeyPress2()


End Sub



Public Sub Bollocks()
'think this just runs if hotkey make do it
'Try so all come through this --- the menu option
    
    If App.title = "Tidal Information..." Then
        XxG = Now
        'Load frmPassLock
        Call PassModule.Main
        Exit Sub
    End If
    
    If frmPassLock.Text4.Visible = True Or XxG + TimeSerial(0, 0, 10) > Now Then
        Call SetLockTime
        Call SetLockMax
        Unload frmPassLock
        DoEvents
    End If

End Sub



Sub SetLockTime()
            
            If HOUR(Now) >= 19 Or HOUR(Now) < 8 Then cute = 20 Else cute = 10
                LockSaverDelay1 = TimeSerial(0, cute, 0)
                LockSaverDelay2 = TimeSerial(0, cute, 0)
            
            If XxG + TimeSerial(0, 0, 10) > Now Then
                LockSaverDelay = LockSaverDelay2
            Else
                LockSaverDelay = LockSaverDelay1
            End If
            
            If LockDown = True Then LockSaverDelay = LockSaverDelay3
            If LockDown2 > Now Then LockSaverDelay = LockSaverDelay4 'LockSaverDelay4 '--- Short delay even shorter for IDE
            'SetLockMax
            'DeadLock = Now + TimeSerial(0, 30, 0)
            
End Sub



Sub Equinox3()

Dim srs As New clsSunRiseSet
srs.City = "London, England"

Rty2 = 0
For Rty3 = -2 To 5
iy = 0
toolnow = DateSerial(Year(Now) + Rty2, 6, Rty3 + 21)
toolx3 = toolnow
srs.DateDay = toolnow
srs.CalculateSun
toolx1 = srs.Sunrise
toolx2 = srs.Sunset
toolx5 = DateDiff("s", toolx1, toolx2)
toolnow = DateSerial(Year(Now) + Rty2, 6, Rty3 + 22)
srs.DateDay = toolnow
srs.CalculateSun
toolx1 = srs.Sunrise
toolx2 = srs.Sunset
toolx6 = DateDiff("s", toolx1, toolx2)
If toolx5 > toolx6 Then toolnow = toolx3: Exit For
Next

'mls$(R) = Format$(toolnow - 93, "DD-MM-yyyy")
'dizzy$ = "Vernal Equinox"
'mls$(R) = Format$(toolnow + 93, "DD-MM-yyyy")
'dizzy$ = "Autumnal Equinox"

 
Equinox5 = toolnow - 93
Equinox6 = toolnow + 93
'Call EquinoxSub

End Sub


Sub WebPages()
Exit Sub

'If TimeBack = 0 Then TimeBack = Now - 1

If TimeBack > Now And QuickPage = False Then Exit Sub

QuickPage = False

hj = 0
If TimeBack = 0 Then hj = 1

H = HOUR(Now + TimeSerial(1, 0, 0))
If H = 0 Then H = 24

TimeBack = Int(Now + 200) + TimeSerial(H, 0, 0)

If hj = 1 Then Exit Sub

'Exit Sub

'hwndw = FindWindow(vbNullString, "BBC - Weather Centre - UK Weather - Windows Internet Explorer")

'If hwndw > 0 Then Exit Sub

lhtmp = FindWindow("IEFrame", vbNullString)
If lhtmp > 0 Then Exit Sub

Timer6.Enabled = True
'Timer6.Interval = 3000

RT2 = 0

'hwndw = FindWindow(vbNullString, "Google News U.K. - Windows Internet Explorer")
'hwndw = FindWindow(vbNullString, "Mirror.co.uk - Windows Internet Explorer")
'Shell "Explorer.exe http://www.friendsreunited.co.uk/friendsreunited.asp"

End Sub

Private Sub Timer6_Timer()

'Exit Sub

If MinDelayPause > 0 And MinDelayPause < Now Then
    MinDelayPause = 0
    WinAmpPowerHwnd = FindWindow("Winamp v1.x", vbNullString)
    MsgResult = SendMessage(WinAmpPowerHwnd, WM_USER, 0, ByVal ipc_isplaying)
    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.
    If MsgResult = 1 Then
        MsgResult2 = SendMessage(WinAmpPowerHwnd, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
    End If
End If

Exit Sub



'ShowWindow wed, SW_MINIMIZE
'ShowWindow wed, SW_HIDE
'ShowWindow wed, SW_SHOWNORMAL
'ShowWindow wed, SW_NORMAL
'ShowWindow wed, SW_SHOWMINIMIZED
'ShowWindow wed, SW_SHOWMAXIMIZED
'ShowWindow wed, SW_MAXIMIZE
'ShowWindow wed, SW_SHOWNOACTIVATE
'ShowWindow wed, SW_SHOW
'ShowWindow wed, SW_MINIMIZE
'ShowWindow wed, SW_SHOWMINNOACTIVE
'ShowWindow wed, SW_SHOWNA
'ShowWindow wed, SW_RESTORE
'ShowWindow wed, SW_SHOWDEFAULT
'ShowWindow wed, SW_FORCEMINIMIZE
'ShowWindow wed, SW_MAX




Dim Ty As Long
Dim ab$(300)
ReDim Preserve Nb(200)

Ty = 0
cc$ = "C:\Program Files\Internet Explorer\IEXPLORE.EXE "
For Ty = 1 To ExecuteWeb.List1.ListCount - 1
ab$(Ty) = cc$ + ExecuteWeb.List1.List(Ty)
Next

'Dim Tr As Long

RT2 = RT2 + 1

'Shell "l", vbMinimizedNoFocus

Dim Hwnd4 As Long

'Hwnd4 = InstanceToWnd(Nb(RT2))
'If Hwnd4 > 0 Then


'Shell (ab$(Ty))
'Tr = Shell(ab$(RT2), vbMinimizedNoFocus)

tr = Shell(ab$(RT2), vbMinimizedNoFocus)

Nb(RT2) = tr
Ty = ExecuteWeb.List1.ListCount - 1
'If RT2 = 2 Then
'Timer6.Enabled = False: Exit Sub
'End If
If RT2 = Ty Then
Timer6.Enabled = False: Exit Sub
End If

Dim hWnd2 As Long
Dim stitle2 As String
Dim sclass2 As String

'Ty = cProcesses.GetHWnd(Tr, Hwnd2, Stitle2, Sclass2)
'ShowWindow Hwnd2, SW_MINIMIZE

End Sub

Sub ShowWeb()

Dim hWnd2 As Long
Dim stitle2 As String
Dim sclass2 As String

For r = 1 To ExecuteWeb.List1.ListCount - 1

If Nb(r) <> 0 Then

stitle2 = ""
sclass2 = ""
hWnd2 = 0
Ty = cProcesses.GetHWnd(Nb(r), hWnd2, stitle2, sclass2)
ShowWindow hWnd2, SW_NORMAL
GoTo kl
ShowWindow hWnd2, SW_MINIMIZE
ShowWindow hWnd2, SW_HIDE
ShowWindow hWnd2, SW_SHOWNORMAL
ShowWindow hWnd2, SW_NORMAL
ShowWindow hWnd2, SW_SHOWMINIMIZED
ShowWindow hWnd2, SW_SHOWMAXIMIZED
ShowWindow hWnd2, SW_MAXIMIZE
ShowWindow hWnd2, SW_SHOWNOACTIVATE
ShowWindow hWnd2, SW_SHOW
ShowWindow hWnd2, SW_MINIMIZE
ShowWindow hWnd2, SW_SHOWMINNOACTIVE
ShowWindow hWnd2, SW_SHOWNA
ShowWindow hWnd2, SW_RESTORE
ShowWindow hWnd2, SW_SHOWDEFAULT
ShowWindow hWnd2, SW_FORCEMINIMIZE
ShowWindow hWnd2, SW_MAX
kl:

End If
'ShowWindow Hwnd2, SW_SHOWDEFAULT
Next

'ShowWindow wed, SW_MINIMIZE
'ShowWindow wed, SW_HIDE
'ShowWindow wed, SW_SHOWNORMAL
'ShowWindow wed, SW_NORMAL
'ShowWindow wed, SW_SHOWMINIMIZED
'ShowWindow wed, SW_SHOWMAXIMIZED
'ShowWindow wed, SW_MAXIMIZE
'ShowWindow wed, SW_SHOWNOACTIVATE
'ShowWindow wed, SW_SHOW
'ShowWindow wed, SW_MINIMIZE
'ShowWindow wed, SW_SHOWMINNOACTIVE
'ShowWindow wed, SW_SHOWNA
'ShowWindow wed, SW_RESTORE
'ShowWindow wed, SW_SHOWDEFAULT
'ShowWindow wed, SW_FORCEMINIMIZE
'ShowWindow wed, SW_MAX

End Sub


Private Sub Timer7_Timer()

Exit Sub

OutlookE3 = GetForegroundWindow
OutlookE2 = FindWindow("#32770", "Outlook Express")

If OutlookE2 > 0 Then
    If OutlookE2 = OutlookE3 And OutlookE2 <> OutlookE Then
        
        ''Call VoiceMP3Pause
        'Call SpeakVoice( "OutLook Started")
        
        OutlookE = OutlookE3
    End If
End If


If OutlookE > 0 Then
    If OutlookE2 <> OutlookE3 Then
        OutlookE = 0
        
        ''Call VoiceMP3Pause
        'Call SpeakVoice( "OutLook Finished")
    
    End If
End If

neroe3 = GetForegroundWindow
neroe2 = FindWindow("#32770", vbNullString)

Dim Rectx As RECT
If neroe2 > 0 And neroe2 = neroe3 Then
    If neroe2 <> NeroEX Then
        Hwnd24 = GetWindowRect(neroe2, Rectx)
        'X= 418 Y= 343 W= 384 H= 151
        If Rectx.Right - Rectx.Left = 384 And Rectx.Bottom - Rectx.Top = 151 Then
            NeroEX = neroe2
            
            ''Call VoiceMP3Pause
            'Call SpeakVoice( "Nero 1 Started")
        
        End If
    End If
End If


If NeroEX > 0 Then
        Hwnd24 = GetWindowRect(NeroEX, Rectx)
        If Hwnd24 = 0 Then
            
            ''Call VoiceMP3Pause
            'Call SpeakVoice( "Nero 1 Finished")
        
        NeroEX = 0
    End If
End If


'Pop Up window at end
'neroe3 = GetForegroundWindow
neroe2 = FindWindow("#32770", "NeroVision Express 3")

If neroe2 > 0 Then
'    If neroE2 = neroe3 And neroE2 <> NeroE Then
    If neroe2 <> NeroE Then
        NeroE = neroe2
        
        ''Call VoiceMP3Pause
        'Call SpeakVoice( "Nero 2 Finished")
    
    End If
End If

DownLoadFinished2 = FindWindow("#32770", "Download complete")

If DownLoadFinished2 > 0 And DownLoadFinished3 <> DownLoadFinished2 Then
    DownLoadFinished3 = DownLoadFinished2
    'DownLoadFinished4 = DownLoadFinished2
    
    'Call VoiceWait
    'Call VoiceMP3Pause
    Call SpeakVoice("DownLoad Finished")

End If

If DownLoadFinished3 > 0 Then
    gw$ = GetWindowTitle(DownLoadFinished3)
    If gw$ = "" Then
        DownLoadFinished3 = 0
    End If

End If



'------------
Exit Sub


'----------- Batch of set lock max progs

If FindWindow("Winamp Video", vbNullString) = GetForegroundWindow Then
LockSSaver = Now + LockSaverDelay
Call SetLockMax
End If

'IceView
If FindWindow("ThunderRT5Form", vbNullString) Then
LockSSaver = Now + LockSaverDelay
Call SetLockMax
End If

If FindWindow("#32770", "Outlook Express") Then
LockSSaver = Now + LockSaverDelay
Call SetLockMax
End If
If FindWindow(vbNullString, "Drive File Lister") Then
LockSSaver = Now + LockSaverDelay
Call SetLockMax
End If
If FindWindow(vbNullString, "ViceVersa Pro - (Registered)") Then
LockSSaver = Now + LockSaverDelay
Call SetLockMax
End If

If FindWindow(vbNullString, "ImageDupe") Then
LockSSaver = Now + LockSaverDelay
Call SetLockMax
End If
    
If FindWindow(vbNullString, "AhaView") > 0 Then Exit Sub

'----------- End -- Batch of set lock max progs



If LockSSaver < Now And App.title = "Tidal Information..." Then
    XxG = Now
    Call PassModule.Main
End If


If App.title = "Tidal Information..." And DeadLock < Now Then
    DeadLock = Now + TimeSerial(0, 30, 0)
    XxG = Now
    Call PassModule.Main
End If

End Sub

Public Sub Timer8WinAmpVol_Timer()


If 1 = 2 Then
    If Actionz = 0 Or Actionz = 2 Or Actionz = 3 Then
        'LockSSaver = Now + LockSaverDelay
        ATidalX.SetLockMax
    End If
End If


'GOT A EXIT SUB IN HERE FOR MOMENT CONTROL VOLUME WITH KEY ASYNC
VolumeLevelSet


End Sub
    

Public Sub VolumeLevelSet()
    
    Exit Sub
    
    '2016
    'PROBLEM FIND OUT WHT KEYS THESE ARE CONTROL VOLUME
    
    
    
    If Sluty3 = 0 Then Exit Sub
    
    If Sluty3 < 10 And Bullet = False Then
        bdf1 = GetAsyncKeyState(121)
        bdf2 = GetAsyncKeyState(120)
        If bdf1 = 0 And Sluty3 = 4 Then
            Sluty3 = 0: Exit Sub
        End If
        If bdf2 = 0 And Sluty3 = 5 Then
            Sluty3 = 0: Exit Sub
        End If
    End If
    Bullet = False
    
    If Sluty3 = 10 Then Sluty3 = 4
    If Sluty3 = 11 Then Sluty3 = 5
    If Sluty3 = 4 Then ATidalX.LblVol = str(Val(ATidalX.LblVol) + 1) 'Raise Vol)
    If Sluty3 = 5 Then ATidalX.LblVol = str(Val(ATidalX.LblVol) - 1)  'Lower Vol)
    'If Sluty3 = 4 Or Sluty3 = 5 Then Call Timer1Set ': Debug.Print Str(Now)

    
'    If VAL(ATidalX.LblVol)<= 8 Then Timer8WinAmpVol.Enabled = False: Timer8WinAmpVol.Interval = 200: Timer8WinAmpVol.Enabled = True
'    If VAL(ATidalX.LblVol)> 8 Then Timer8WinAmpVol.Enabled = False: Timer8WinAmpVol.Interval = 30: Timer8WinAmpVol.Enabled = True
    
    If Val(ATidalX.LblVol) < 0 Then ATidalX.LblVol = str(0)
    If Val(ATidalX.LblVol) > 100 Then ATidalX.LblVol = str(100)
    
    
    MsgResult = WinAmpChkIsPlaying
    
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.
    
    If MsgResult = 1 Then
        VolumeSet1 = ATidalX.LblVol
    Else
        VolumeSet2 = ATidalX.LblVol
    End If


    DoEvents
        
    tf = SetVolume(SPEAKER, ATidalX.LblVol)
    
    On Local Error Resume Next

    'fr4 = FreeFile
    'Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Volume.txt" For Output As #fr4
    'Print #fr4, str(VolumeSet1)
    'Print #fr4, str(VolumeSet2)
    'Close #fr4

    'Sluty3 = 0


End Sub





Public Sub TimerDoubleVolume_Timer()
    
    'DEFAULT 10 - WE CHANGE LESS FREQENT
    
    If TimerDoubleVolume.Interval = 10 Then
        TimerDoubleVolume.Interval = 100
    End If
    
    
    'ONLY TO USE IN MANULE MODE
    'PROBLEM
    'If VoiceStreamActive = True Then Exit Sub



    'SAY TIME OR ANY SPEAK
    'If SayTime4 = True Then Exit Sub

    '## NEW FIX IGNOR ABOVE
    ' CODE TO CHANGE MASTER MUSIC VOL ON ANY PAUSE UNPAUSE
    ' WHEATHER MANUAL OR AUTO
    ' WE KNOW AUTO DOES ANYWAY SO THIS SHOULD ALLOW
    ' AUTO TO WORK AS WELL AS MANUAL

    ' EXAMPLE ENTER PAUSE STATE VOLS CHANGED IN AUTO
    ' HERE WAT HAPPEN -- WILL SEE PAUSE CHANGE AND WAT
    ' DECIDE TO CHANGE VOLS MASTER MUSIC
    ' SHOULD KNOW IF VOL FOR EACH IS AS SHOULD BE EASY REALLY
    ' ANY REMEMBERING OR IS THAT TAKEN CARE OF IN AUTOS FOR
    ' ONLY AUTOS OR WAT
    
    If WinAmpChkIsPlaying = 1 Then
        ATidalX.ProgressBarVol1.scrolling = ccScrollingSmooth
        ATidalX.ProgressBarVol2.scrolling = ccScrollingStandard
    Else
        ATidalX.ProgressBarVol1.scrolling = ccScrollingStandard
        ATidalX.ProgressBarVol2.scrolling = ccScrollingSmooth
    End If


    
    If WinAmpChkIsPlaying = 1 Then
        If Val(ATidalX.LblVol) = VolumeSet1 Then Exit Sub
        ATidalX.LblVol = str(VolumeSet1)  ' MUSIC VOL
    Else
        If Val(ATidalX.LblVol) = VolumeSet2 Then Exit Sub
        ATidalX.LblVol = str(VolumeSet2)  ' MASTER VOL
    End If
    
    tf = SetVolume(SPEAKER, Val(ATidalX.LblVol))
   
    
End Sub


Private Sub TimerNotePadYes_Timer()
TimerNotePadYes.Enabled = False
'----------------------------------
Exit Sub


'If MMControl8.mode = 525 And Go8 = True Then
'    ATidalX.MMControl8.Command = "prev"
'    ATidalX.MMControl8.Command = "Play"
'End If


notepad2rt = FindWindow("#32770", "Notepad2")
notepad3rt = FindWindow("Notepad2", vbNullString)
xxrp = GetParent(notepad2rt)
cuecue$ = GetWindowTitle(notepad3rt)

If xxrp <> notepad3rt Then Exit Sub

If notepad2rt = 0 Then Exit Sub

If NotePad2RT2 = notepad3rt Then Exit Sub

If GetForegroundWindow <> notepad2rt Then Exit Sub

If InStr(cuecue$, "Untitled") > 0 Then Exit Sub
If InStr(cuecue$, "(Read Only)") > 0 Then Exit Sub

If Mid$(cuecue$, 1, 1) <> "*" Then Exit Sub
    
'THIS MAKES ERROR KEEP SWAPING CAPS LOcK KEY
'SendKeys "{ENTER}", False
'T = Timer + 0.3
'Do
'DoEvents
'If Timer < T - 20 Then Exit Do
'Loop Until T < Timer

'SendKeys "%Y", False
Call SendKey(Asc("Y"), 0)  '# Y for Yes

If GetForegroundWindow = notepad2rt Then Exit Sub

NotePad2RT2 = notepad3rt

End Sub



Private Sub LABEL_TIME_DblClick()
'TooLNewSunRise = 1
'TooLNewSunSet = 1
Tog6 = 0
speedtime = 0
Dag = Now + TimeSerial(0, 0, 9)
Call moonnext
Tog7 = Now + TimeSerial(0, 0, 5)
End Sub



Private Sub TimerVolLab_Timer()

    If VoiceStreamActive = True Then Exit Sub
    If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then Exit Sub


    WinAmpPowerHwnd = FindWindow("Winamp v1.x", vbNullString)
    MsgResult = SendMessage(WinAmpPowerHwnd, WM_USER, 0, ByVal ipc_isplaying)

    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.
    Dim FTIME1 As Date
    Dim FTIME2 As Date
    FTIME1 = Now - Int(Now)
    FTIME2 = Int(Now)
    
    
    '--------------------
    XAZYTAG = 0
    
    
    '------------------------------------
    On Error Resume Next
        If UBound(XAZY) = 0 Then Exit Sub
        If Err.Number > 0 Then Exit Sub
    On Error GoTo 0
        
    For RS = XAZYTAG To UBound(XAZY) ' EXIT'S ON -2 THEN OKAY
    
        RS5 = RS5 + 1
    
        XAZYTAG = XAZYTAG + 1
        
        'XAZY(1)  IS NOT PART THIS SUB
        'AND USED TO BY SUNRISE SET CODE -- + NOON
        'SO WE HAVE > 1
        If XAZYTAG > 1 Then XAZY(XAZYTAG) = -2
        
        Select Case RS5
            Case 1
                '----------11:11:11 STOP
                XAZY(XAZYTAG) = FTIME1 > TimeSerial(11, 8, 0) And FTIME1 < TimeSerial(11, 20, 0)
            Case 2
                II3 = TooLTodaySunRise - TimeSerial(0, 0, 4)
                ii4 = TooLTodaySunRise + TimeSerial(0, 0, 4)
                XAZY(XAZYTAG) = (FTIME1 > II3) And FTIME1 < ii4
            Case 3
                II3 = TooLTodaySunSet - TimeSerial(0, 0, 4)
                ii4 = TooLTodaySunSet + TimeSerial(0, 0, 4)
                XAZY(XAZYTAG) = (FTIME1 > II3) And FTIME1 < ii4
            Case 4
                II3 = TooLTodayNoon - TimeSerial(0, 0, 4)
                ii4 = TooLTodayNoon + TimeSerial(0, 0, 4)
                XAZY(XAZYTAG) = (FTIME1 > II3) And FTIME1 < ii4
            Case 5
                '----------12AM STOP
                'XAZY(XAZYTAG) = FTIME1 > TimeSerial(23, 58, 0) And FTIME1 < TimeSerial(0, 5, 0)
            Case 6
                '---------- SUNDAY 4PM STOP  ------ 'SAT = DAY  CHK
                'XAZY(XAZYTAG) = (FTIME1 > TimeSerial(2, 58, 0) And FTIME1 < TimeSerial(3, 5, 0)) And Day(Now) = 6
            Case 7
                '----------3PM STOP
                'XAZY(XAZYTAG) = FTIME1 > TimeSerial(14, 55, 0) And FTIME1 < TimeSerial(15, 5, 0)
            Case 8
                '----------5.30PM STOP
                'XAZY(XAZYTAG) = FTIME1 > TimeSerial(16, 25, 0) And FTIME1 < TimeSerial(17, 35, 0)
            Case 9
                '----------10PM STOP
                'XAZY(XAZYTAG) = FTIME1 > TimeSerial(21, 55, 0) And FTIME1 < TimeSerial(22, 5, 0)
            Case 10
                '----------11PM STOP
                'XAZY(XAZYTAG) = FTIME1 > TimeSerial(22, 55, 0) And FTIME1 < TimeSerial(23, 5, 0)
            Case 11
            
        
        End Select
        
        If XAZY(XAZYTAG) = -2 Then Exit For
    
    
        XAZYTAG = XAZYTAG + 1
        For R1 = 0 To 23
            For R2 = 0 To 30 Step 30
                'EXIT FOR BELOW
                Tx = -1
                If R1 Mod 3 = 0 And R2 = 0 Then
                    rx = 1
                    Else: rx = 0
                End If
                TS1 = TimeSerial(R1, R2 - (2 + rx), 55)
                TS2 = TimeSerial(R1, R2 + (1 + rx), 1)
                XAZY(XAZYTAG) = False
                If FTIME2 + TS1 < Now And FTIME2 + TS2 > Now Then
                    XAZY(XAZYTAG) = True
                    Tx = 0
                    'Debug.Print "--------"
                    'Debug.Print Now
                    'Debug.Print Format(FTIME2 + TimeSerial(R1, R2 - (2 + rx), 58), "dd MM YYYY HH MM SS")
                    'Debug.Print Format(FTIME2 + TimeSerial(R1, R2 + (1 + rx), 1), "dd MM YYYY HH MM SS")
                    Tx = -3
                    Exit For
                
                End If
                If Now < TS2 + FTIME2 Or Tx = -3 Then
                    Exit For
                End If
            
            Exit For
            Next
            If Now < TS2 + FTIME2 Or Tx = -3 Then
                Exit For
            End If
        Next
    Next
        
        
    ReDim Preserve XAZY(XAZYTAG)
    
    For XAZYTAG = 1 To UBound(XAZY) ' EXIT'S ON -2 THEN OKAY
        'Debug.Print XAZYTAG
        'Debug.Print XAZY(XAZYTAG)
        'Debug.Print OXAZY(XAZYTAG)
        
        
        'MAY OF BEEN THE BIG PROBLEM OF PAUSE AND NOT RESUME
        'THIS FOR NEXT LOOP NEED CHECK IF PLAYING AT EACH STEP
        
        WinAmpPowerHwnd = FindWindow("Winamp v1.x", vbNullString)
        MsgResult = SendMessage(WinAmpPowerHwnd, WM_USER, 0, ByVal ipc_isplaying)

        If MsgResult = 1 And XAZY(XAZYTAG) = True And OXAZY(XAZYTAG) <> XAZY(XAZYTAG) Then
            XZAP(XAZYTAG) = True
            'If HOUR(Now) = 11 Then
            
            'BOOKMARK FLAG POSITION TO WATCH
            
            Debug.Print "---------------------------"
            Debug.Print Format(Now, "MMM DD HH MM SS") + " WINAMP COMMAND PAUSE"
            Debug.Print "IN SUB -- TimerVolLab_Timer"
            Debug.Print "---------------------------"
            If MNU_STOP_WINAMP_ON_THE_HOUR.Checked = True Then
                MsgResult2 = SendMessage(WinAmpPowerHwnd, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
            End If
            'Debug.Print "STOP ON" + str(RS5)
        End If
    
        '---------------------------------------------------------------
        'WHAT IS RESULT3 -- SHOULD NOT HAPPEN HERE IS BEEN ONLY IN PAUSE
        'YES CORRECT IS 3 WHEN PAUSE
        '---------------------------------------------------------------
        If MsgResult = 3 And XAZY(XAZYTAG) = False And OXAZY(XAZYTAG) <> XAZY(XAZYTAG) And XZAP(XAZYTAG) = True Then
            'If HOUR(Now) = 11 Then
            
            If MNU_STOP_WINAMP_ON_THE_HOUR.Checked = True Then
                MsgResult2 = SendMessage(WinAmpPowerHwnd, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
            End If
            Debug.Print "START ON" + str(RS5)
        End If

        OXAZY(XAZYTAG) = XAZY(XAZYTAG)
        If XAZY(XAZYTAG) = False And XZAP(XAZYTAG) = True Then XZAP(XAZYTAG) = False
    
    Next

    '---- ANY MORE STOPS WANTED
    '11:11:11 -- 6PM -- SUN SET RISE NOON
    '3PM
    '5.30
    '10PM
    '4PM ON SUNDAY
    '12 MIDDAY

End Sub

Sub VoiceMP3Pause()
    
    Call MP3_OFF__ON_START_VOICE
    
    Exit Sub
    
    
    
    'NEED IS WINAMP PLAY THEN TO STOP
    'IF NOT RESET ALL FLAGS TO DO WITH MP3 LATER
    'CHK NOT PUT IN STOP STATE BY US
    
    
    'NEED IS WINAMP IN STOP STATE AND NEEDS TO RESUME PLAY
        
    'NEED IF LITTLE VOICE AND STRING COMES TO START THEN
    'SET STRING MODE ACTIVE
        
    '--------------------------
    
    'IS STRING OF VOICE ASKED TO PLAY AND ALREAY PLAYING STRING THEN FLAGG IT
    'THIS = WINAMP IS PLAYING AND ALSO A STRING OF VOICE IS PLAYING
    If WinampISPlayingAndSetToStopFlagAfterVoiceString = True Then Exit Sub
    
    'If StringVoiceToPLayActive = True Then
    If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then
        If WinampISPlayingAndSetToStopFlagAfterVoice = True Then
            WinampISPlayingAndSetToStopFlagAfterVoiceString = True
        End If
    End If
    
    'THIS IS CHK IF IN STOP STATE BY US
    If WinampISPlayingAndSetToStopFlagAfterVoice = True Then Exit Sub
        
    '--------------------------
        
    'DO - IF WINAMP IN STOP STATE BY US AND TO DO STRING VOICE THEN
    'JUST SET STRING FLAG GOING NO OTHER CHK
    
        
        
    
    '#0 ASK MP3 TO STOP
    
    
    
    '#2 IF MP3 STOPPED COZ ASKED AND REQUEST AGAIN THEN
    '#2 IGNOR 2ND REQUEST
    
    '#3 IF STRING OF VOICE THEN HOLD SOUND
    
    '#4 IF STRING OF VOICE AND LITTLE ONE VOICE
    '#4 THAT WILL GET QUE-ED IN BUFFER
    '#4 AND WHEN STRING FINISHED
    '#4 THE THE LITTLE VOICE SHOULD ADD TO BUFFER TO COME NEXT
    '#4 AS USUAL AND MAKE SURE MP3 NOT ACTIVATED FOR
    '#4 FOR EVEN MILLI SECS
    '#4 IE THERE SHOULD BE KNOWN THAT THERE IS MORE VOICE IN QUE
    
    '#5 ALWAYS RESET LATCH THAT SAID WAS STRING VOICE
    
    '#7 WHEN IS STRING VOICE OR VOICE IS PLAYING IN ANYWAY
    '#7 THEN NEVER NEED TO DO CHK'S TO STOP MP3 %%
    
    '#8 SHALL WE HAVE COUNTER TO SHOW HOW MANY IN QUE
    '#8 NO
    
    '#9 IF LITTLE VOICE AND STRING SETS TO START THEN
    '#9 SHOULD BE AS NORM %%
    '#9 JUST SET FLAGUP FOR STRING IN MOTION
    
    
    'OLD SUPER CODE WAY OF DOING IT
    'WinAmpPowerHwnd = FindWindow("Winamp PE", vbNullString)
    'MsgResult = WinAmpChkIsPlaying
    'WinAmpPowerHwnd = WinAmpChkIsPlayingHwnd
    'WinAmpPowerHwndStore = WinAmpPowerHwnd
    
    WINAMP2 = FindWindow("Winamp v1.x", vbNullString)
    
    If WINAMP2 > 0 Then MsgResult = SendMessage(WINAMP2, WM_USER, 0, ByVal ipc_isplaying)

    If WINAMP2 = 0 Or MsgResult = 0 Then
        'CLEAR ALL FLAGS AS GOT THIS FAR AN NONE PLAYING
        
        'WHY = REMOVED 2016
        'StringVoiceToPLayActive = False
        
        'NEVER GET THIS FAR ANYHOW
        
        'TRY THIS
        WinampISPlayingAndSetToStopFlagAfterVoiceString = False
    End If


    If WINAMP2 > 0 And MsgResult = 1 Then
    
    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.
    
    If SayTime4 = False Then Call VolNibUP

    '# SWITCH MP3 TO PAUSE AS SAYS PLAYING THEN READY FOR VOICE
    MsgResult2 = SendMessage(WINAMP2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
    DoEvents
    'SLEEP H
    Sleep 200
    DoEvents
    
    Call TimerDoubleVolume_Timer
        
    WinampISPlayingAndSetToStopFlagAfterVoice = True
        
    'If StringVoiceToPLayActive = True Then
    If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then
        WinampISPlayingAndSetToStopFlagAfterVoiceString = True
    End If
    
    'NEXT THIS TIMER WILL START
    VoiceMP3Start_TIMER.Enabled = True
    
    Exit Sub
    
End If

'THIS ALWAYS STARTS ONLY IF SUB NOT EXIT SUB AS ABOVE
VoiceMP3SayTimer4.Enabled = True

End Sub



Public Sub VOICELISTTIMER_TIMER()

On Error GoTo Err_trap

XGO = 0
If WINAMPBackOnAfterVoiceRequired = True Then
    If VoiceStreamActive = False Then
        If STRING_OF_VOICE_IN_QUE_PLAY = 0 Then
            If ListVOICE.ListCount = 0 Then
                '----------------
                XGO = 1
                '----------------
            End If
        End If
    End If
End If


'------------------------------------------------------
'PROBLEM GO INTO A REPEAT QUICK LOOP EVERY OTHER MOMENT
'------------------------------------------------------
If ListVOICE.ListCount > 0 Then
    'Debug.Print "ListVOICE.ListCount -- " + str(ListVOICE.ListCount)
End If

If XGO = 1 Then

'-----------------------
'SOMETHING WRONG HERE ABOUT THE NEXT VAR -- WAS REM OUT BUT THINK IS REQUIRE
'CHECK 2ND NEXT VAR ALSO
'-----------------------
'SEEM FIX THAT PROBLEM WAS ONLY HAPPEN AT DARKNESS! TALK VOICE
'AND MAY STRING OF VOICE COUNTDOWN TO CHECK AT THE HOUR OCLOCK
'-------------------------------------------------------------
    
    WINAMPBackOnAfterVoiceRequired = True
'    VoiceStreamEnds = False
    Call MP3_ON_AFTER_VOICE_FINISHES
    WINAMPBackOnAfterVoiceRequired = False
    Call ATidalX.TimerDoubleVolume_Timer

End If

If ListVOICE.ListCount = 0 Then
    '-----------------------------
    'VOICELISTTimer.Enabled = False
    'Exit Sub
    '-----------------------------
End If

If VoiceStreamActive = True Then Exit Sub

'WANT TO LOAD ANOTHER IN CANT USE LINE BELOW
'If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then Exit Sub



Call WaitWavFinish
If ListVOICE.ListCount = 0 Then Exit Sub
VoiceToSpeak = ListVOICE.List(0)
ListVOICE.RemoveItem (0)

If Trim(VoiceToSpeak) = "" Then Exit Sub
'HERE GO ABOUT TO START VOICE
If WinAmpChkIsPlaying = 1 Then
'    TTSAppMain.RateSldr.Value = 2
Else
    'TTSAppMain.RateSldr.Value = 2
End If

'Call TTSAppMain.RateSldr_Scroll

'If WinAmpChkIsPlaying = 1 Then
'    WINAMPBackOnAfterVoiceRequired = True
'    Call ATidalX.MP3_OFF__ON_START_VOICE
'End If

WINAMPBackOnAfterVoiceRequired = False
MsgResult = ATidalX.IsWinAmpPlayingBeforeVoiceStartAndStop
If MsgResult = 1 Then
    WINAMPBackOnAfterVoiceRequired = True
    
    '------------------------------------------
    'FAULT HERE ABOUT NOT GOING BACK TO PLAYING
    'WHEN TIMING IS AT BREAK POINT STEP MODE
    '------------------------------------------
    
    
    Call ATidalX.MP3_OFF__ON_START_VOICE
    Call ATidalX.TimerDoubleVolume_Timer
 End If

'VOICECOMMANDSTART = True
'VoiceStreamEnds = False

VoiceToSpeak = Trim(Replace(VoiceToSpeak, "TIME #", ""))
VoiceToSpeak = Trim(Replace(VoiceToSpeak, "TIME CD#", ""))


TTSAppMain.Voice.Speak VoiceToSpeak, SVSFlagsAsync

'Debug.Print "VOICE SPEAK RESULT " + VoiceToSpeak


Debug.Print "02 -- " + Format(Now, "HH MM SS  ") + "VOICE SPEAK RESULT " + VoiceToSpeak



Exit Sub
Err_trap:
'MOST LIKELY A INCORECT VOICE DRIVER SELECTED THAT NOT PRESENT
If SAY_ONCE_MOST_LIKELY_A_INCORECT_VOICE_DRIVER = False Then
MsgBox "MOST LIKELY A INCORECT VOICE DRIVER SELECTED THAT NOT PRESENT IN THE SYSTEM", vbMsgBoxSetForeground
End If
SAY_ONCE_MOST_LIKELY_A_INCORECT_VOICE_DRIVER = True
'Stop
Resume Next




'-----------------
Exit Sub
'-----------------



' ----------- LOGG THE VOICE SPOKE

If WinAmpChkIsPlaying = 1 Or 1 = 1 Then

        Err.Clear
ResumeHere1:
        If Err.Number > 0 Then Sleep 50
        On Error GoTo ResumeHere1

        If VUMPATH1 = "" Then
            If FindWindow(vbNullString, "VU METER LOGGER") > 0 Then
                Open "D:\temp\VU LOGGER PATHS.txt" For Input As #1
                 Line Input #1, VUMPATH1
                Line Input #1, VUMPATH2
            End If
        End If

        If VUMPATH1 <> "" Then
            If FindWindow(vbNullString, "VU METER LOGGER") > 0 Then
    
                Err.Clear
ResumeHere2:
                If Err.Number > 0 Then Sleep 50
                On Error GoTo ResumeHere2
    
                'Call FileInUseDelay(VUMPATH2, False)
                
                FILENAME_IN_USE_CHECK = VUMPATH2
                FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
                DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
                FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
                
                FR1 = FreeFile
                Open FILENAME_IN_USE_CHECK For Append Lock Write As #FR1
                    Print #FR1, "# VOICE SPEAK " + Format(Now, "HH:MM:SS") + " - " + VoiceToSpeak
                Close #FR1

    
                Err.Clear
ResumeHere4:
                If Err.Number > 0 Then Sleep 50
                On Error GoTo ResumeHere4

                TIDALPATHVOICE = "D:\TEMP\VOICE SPEAK LOGG.TXT"
                'Call FileInUseDelay(TIDALPATHVOICE, False)
                
                FILENAME_IN_USE_CHECK = TIDALPATHVOICE
                FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
                DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
                FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
                
                FR1 = FreeFile
                Open FILENAME_IN_USE_CHECK For Append Lock Write As #FR1
                    Print #FR1, "# VOICE SPEAK " + Format(Now, "HH:MM:SS") + " - " + VoiceToSpeak
                Close #FR1

            End If '  If FindWindow(vbNullString, "VU METER LOGGER") > 0 Then
        End If ' VUMPATH AVIAL
End If '  If WINAMP YES Then

'-------------
'BACK TRACKING
'-------------
Err.Clear
ResumeHere3:
If Err.Number > 0 Then Sleep 50
On Error GoTo ResumeHere3
      
TIDALPATHVOICE = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\VOICE SPEAK LOGG.TXT"
'Call FileInUseDelay(TIDALPATHVOICE, False)
FILENAME_IN_USE_CHECK = TIDALPATHVOICE
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append Lock Write As #FR1
    Print #FR1, "# VOICE SPEAK " + Format(Now, "DD-MM-YYYY - HH:MM:SS") + " - " + VoiceToSpeak
Close #FR1

On Error GoTo 0


'If ListVOICE.ListCount > 0 Then Call VOICELISTTIMER_TIMER


End Sub



Sub VoiceMP3SayTimer4_timer()

    'NEW CODE REPLACES
    VoiceMP3SayTimer4.Enabled = False
    
    '-----------------
    Exit Sub
    '-----------------
    
    'StartedVoiceLatch = False

    'If VoiceStreamEnds = False Then Exit Sub
    'VoiceStreamEnds = False
    
    
    VoiceMP3SayTimer4.Enabled = False
    
    Call SayTimer4BackToLower

End Sub



Public Sub ME_PLAY_TEXT_TO_SPEECH_TIMER_TIMER()
'ATidalX.ME_PLAY_TEXT_TO_SPEECH_TIMER.Enabled = True
SPEED_TIMER = 40
If ME_PLAY_TEXT_TO_SPEECH_TIMER.Interval <> SPEED_TIMER Then ME_PLAY_TEXT_TO_SPEECH_TIMER.Interval = SPEED_TIMER

SOUND_FILE_ME_1 = App.Path + "\#Wave Sounds\180419_0012_2018-04-19--10-50__ME__WORD_TEXT_TO_SPEECH_SOUND_EFFECT__WWW_TEXTOSPEECH_INFO_1.WAV"

SOUND_FILE_ME_1 = App.Path + "\#Wave Sounds\ME__WORD_TEXT_TO_SPEECH_1.WAV"
SOUND_FILE_ME_2 = App.Path + "\#Wave Sounds\ME__WORD_TEXT_TO_SPEECH_2.WAV"
SOUND_FILE_ME_3 = App.Path + "\#Wave Sounds\ME__WORD_TEXT_TO_SPEECH_3.WAV"
SOUND_FILE_ME_4 = App.Path + "\#Wave Sounds\ME__WORD_TEXT_TO_SPEECH_4.WAV"
SOUND_FILE_ME_5 = App.Path + "\#Wave Sounds\ME__WORD_TEXT_TO_SPEECH_5.WAV"

'SOUND_FILE_ME_1 = App.Path + "\#Wave Sounds\ZTUTTI - Copy_1.WAV"
'SOUND_FILE_ME_2 = App.Path + "\#Wave Sounds\ZTUTTI - Copy_2.WAV"
'SOUND_FILE_ME_3 = App.Path + "\#Wave Sounds\ZTUTTI - Copy_3.WAV"
'SOUND_FILE_ME = App.Path + "\#Wave Sounds\ZTUTTI.WAV"

If Dir(SOUND_FILE_ME_1) <> "" And SOUND_FILE_ME_SET_ONCE = False Then
SOUND_FILE_ME_SET_ONCE = True

    ATidalX.MMControl10.Enabled = True
    ATidalX.MMControl10.Notify = True
    ATidalX.MMControl10.Wait = True
    ATidalX.MMControl10.Shareable = False
    ATidalX.MMControl10.DeviceType = "waveaudio"
    ATidalX.MMControl10.Filename = SOUND_FILE_ME_1
    ATidalX.MMControl10.Command = "Open"
    
    ATidalX.MMControl11.Enabled = True
    ATidalX.MMControl11.Notify = True
    ATidalX.MMControl11.Wait = True
    ATidalX.MMControl11.Shareable = False
    ATidalX.MMControl11.DeviceType = "waveaudio"
    ATidalX.MMControl11.Filename = SOUND_FILE_ME_2
    ATidalX.MMControl11.Command = "Open"
    
    ATidalX.MMControl12.Enabled = True
    ATidalX.MMControl12.Notify = True
    ATidalX.MMControl12.Wait = True
    ATidalX.MMControl12.Shareable = False
    ATidalX.MMControl12.DeviceType = "waveaudio"
    ATidalX.MMControl12.Filename = SOUND_FILE_ME_3
    ATidalX.MMControl12.Command = "Open"
    
    ATidalX.MMControl13.Enabled = True
    ATidalX.MMControl13.Notify = True
    ATidalX.MMControl13.Wait = True
    ATidalX.MMControl13.Shareable = False
    ATidalX.MMControl13.DeviceType = "waveaudio"
    ATidalX.MMControl13.Filename = SOUND_FILE_ME_4
    ATidalX.MMControl13.Command = "Open"
    
    ATidalX.MMControl14.Enabled = True
    ATidalX.MMControl14.Notify = True
    ATidalX.MMControl14.Wait = True
    ATidalX.MMControl14.Shareable = False
    ATidalX.MMControl14.DeviceType = "waveaudio"
    ATidalX.MMControl14.Filename = SOUND_FILE_ME_5
    ATidalX.MMControl14.Command = "Open"
    
    
End If
If X_COUNT_ME_SOUND > 5 Then X_COUNT_ME_SOUND = 0
X_COUNT_ME_SOUND = X_COUNT_ME_SOUND + 1
If X_COUNT_ME_SOUND = 1 Then
    ATidalX.MMControl10.Command = "Prev"
    ATidalX.MMControl10.Command = "Play"
End If
If X_COUNT_ME_SOUND = 2 Then
    ATidalX.MMControl11.Command = "Prev"
    ATidalX.MMControl11.Command = "Play"
End If
If X_COUNT_ME_SOUND = 3 Then
    ATidalX.MMControl12.Command = "Prev"
    ATidalX.MMControl12.Command = "Play"
End If
If X_COUNT_ME_SOUND = 4 Then
    ATidalX.MMControl13.Command = "Prev"
    ATidalX.MMControl13.Command = "Play"
End If
If X_COUNT_ME_SOUND = 5 Then
    ATidalX.MMControl14.Command = "Prev"
    ATidalX.MMControl14.Command = "Play"
End If

'XN = Int(Now) + Timer + 0.14
'Do
'    DoEvents
'    X_DEBUG = X_DEBUG + 1
'Loop Until ATidalX.MMControl10.mode <> 526 Or XN < Int(Now) + Timer
'If X_DEBUG > 1 Then Debug.Print X_DEBUG
'BEEN_GO = False
'If ATidalX.MMControl10.mode <> 526 Then BEEN_GO = True
'If ATidalX.MMControl11.mode <> 526 Then BEEN_GO = True
'If ATidalX.MMControl12.mode <> 526 Then BEEN_GO = True
'If ATidalX.MMControl13.mode <> 526 Then BEEN_GO = True
'If ATidalX.MMControl14.mode <> 526 Then BEEN_GO = True
'If BEEN_GO = False Then Debug.Print BEEN_GO


End Sub

Sub SayTimer4BackToLower()
    
    
      


End Sub


Sub MP3_IS_ABOUT_TO_OFF_ON_CHANGE_VOLS(OFF_MP3)


    'IF MASTER VOL LESS THAN MUSIC VOL
    'AND MUSIC PLAYING THEN DO NONE EXIT
    'DEFAULT THIS WONT GET CALLED UNLESS MUSIC PLAYYING

    If MNU_NO_MATCH_WINAMP_VOL.Checked = False Then
        If VolumeSet2 > VolumeSet1 Then Exit Sub
    End If

    Vol_FlagSet = Not OFF_MP3


    'TRUE = MP3 PLAY MODE IS SET TO GO ON SET MUSIC VOL ON
    'OR
    'FALSE THEN USE MASTER VOL AS VOICE MODE ON NOT MUSIC

    SayTime4 = False
    
    If OFF_MP3 = True Then
        If ATidalX.LblVol = str(VolumeSet1) Then Exit Sub
        ATidalX.LblVol = str(VolumeSet1)  ' MUSIC VOL
    Else
        If ATidalX.LblVol = str(VolumeSet2) Then Exit Sub
        ATidalX.LblVol = str(VolumeSet2)  ' MASTER VOL
    End If
    
    tf = SetVolume(SPEAKER, ATidalX.LblVol)

End Sub


Sub MP3_ON_AFTER_VOICE_FINISHES()
    
    ' MP3 WILL TURN BACK ON --
    ' BUT REPEAT ATTPEMPTS TO ON WILL NOT BE NEEDED
    ' IF IN STRING MODE
    
    If WINAMPBackOnAfterVoiceRequired = False Then
        Exit Sub
    End If
    
    If VoiceStreamActive = True Then
        Exit Sub
    End If
    If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then
        Exit Sub
    End If

    
    
    'YES VOL2 IS NOT MUSIC
    'CLEAN CORRECT
    'WHY -- WE THINK THIS CHANGES VOL AND DOUBLE VOL IS USED FOR MANUAL
    If oVolumeSet2 <> VolumeSet2 Then
        'VolumeSet2 = oVolumeSet2
        'tf = SetVolume(SPEAKER, VolumeSet2)
    End If
    
    
    'Call SayTimer4BackToLower
    
    'EASY JOB DONE KNOW WAT TO DO PUT WINAMP BACK ON
    'UNLESS BEEN MANUALE STARTED
    WINAMP2 = FindWindow("Winamp v1.x", vbNullString)
    MsgResult = SendMessage(WINAMP2, WM_USER, 0, ByVal ipc_isplaying)
    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.
    If MsgResult = 3 Then
        
        'VOLUME NORMALISING
        
        'VolumeSet2 = MASTER VOL
        If Vol_FlagSet = True Then
            
            Call MP3_IS_ABOUT_TO_OFF_ON_CHANGE_VOLS(True)
        End If
        
        
        'ON AFTER VOICE
        'ISSUECOMMTOSTOPSTARTWINAMP = True
        
        
        '-------------------------------
        'BOOKMARK FLAG POSITION TO WATCH
        '-------------------------------
        '-------------------------------------------------
        'THE PROBLEM NEED SOLVE OF PUT PLAYING OFF AND THEN NEXT TURN BACK
        'IS NEAR HERE
        'GUESS THIS IS THE TURN BACK ON IN NORMAL SERNERIO
        '-------------------------------------------------
        'MOST CASE THIS HAPEN AS NORMAL PAUSE IS RESUMED BACK TO PLAYING
        '---------------------------------------------------------------
        'DARKNESS VOICE HAPPEN MOMENT AFTER TIME IN BREAKPOINT STEP MODE
        'AND PAUSE PROBLEM AGAIN
        '-----------------------
        'BECAUSE TIME NEVER HAPPEN TOGETHER THIS IS PROBLEM WHERE DARKNESS!
        'WOULD FOLLOW EXACT SAME TIME AFTER
        '----------------------------------
        
        MsgResult2 = SendMessage(WINAMP2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
        Debug.Print "-- -- " + Format(Now, "HH MM SS") + "  SET WINAMP PLAYING -- HWND " + str(WINAMP2)
    End If
    

End Sub

Sub MP3_OFF__ON_START_VOICE()
        
    ' MP3 WILL TURN OFF --
    ' BUT REPEAT ATTPEMPTS TO OFF WILL NOT BE NEEDED
    ' IF IN STRING MODE
     
    '---------------------------------------------------------------
    'TEST SEE IF PUT HERE WILL STOP WINAMP PAUSE MODE AT WINAMP PLAY
    '1ST TEST CONDICTION
    If ATidalX.MNU_NO_PAUSE.Checked = True Then Exit Sub
    
    
   ' IS NOT PLAYING THEN ACTION THAT -- PRETTY MUCH DO NONE

    
    WINAMP2 = FindWindow("Winamp v1.x", vbNullString)
    If WINAMP2 = 0 Then Exit Sub
    MsgResult = SendMessage(WINAMP2, WM_USER, 0, ByVal ipc_isplaying)
    ' IT IS PLAYING = 1 THEN
    If MsgResult <> 1 Then
    
        Exit Sub
        
    End If
    
    'CHECK THE VOLUME IS UP FROM MUTE
    MsgResult = SendMessage(WINAMP2, WM_USER, 0, ByVal ipc_isplaying)
    ' IT IS PLAYING = 1 THEN STOP GOING WITH PAUSE
    If MsgResult = 1 Then
        '----------------------------------------------------------------------------
        '----
        'Winamp: set Volume to a given value - Ask for Help - AutoHotkey Community
        'https://autohotkey.com/board/topic/53975-winamp-set-volume-to-a-given-value/
        '----
        'GET THE MASTER VOLUME OF WINAMP 0 TO 255
        '----------------------------------------------------------------------------
        MsgResult = SendMessage(WINAMP2, WM_USER, -666, ByVal 122)
        '----------------------------------------------------------------------------
        'IF MUTE NAUGHT WE WANT ABORT USE A PUASE FOR SOUND DOWN WHILE VOICER
        '----------------------------------------------------------------------------
        Exit Sub
    End If
    
    If MsgResult = 1 Then
        
        'DEMAND IS USED WHEN SOUND EFFECT PRECEDES VOICE
        'THIS MEANS THIS ROUTINE WAS CALLED MANUAL
        'RATHER THAN BY THE EVENT DRIVEN WHICH IS MORE USUAL
        
        
        'VOLUME NORMALISING
        'VolumeSet2 = MASTER VOL
        
        '--------------
        'THIS KEEPS LEVEL AT MP3 LEVEL -- IF THIS SO THEN NONE TO DO
        '##   ---- BUT
        'WE MAY WANT TO OVERRIDE THIS AND HAVE AS LESS VOLUME WHEN SPEAKING
        'LIKE WHEN MP3 LISTERN IS LOW VOL AND HAVE TO TURN UP HIGHER
        '--------------
'        If VolumeSet2 > VolumeSet1 Then
            'MNU_NO_MATCH_WINAMP_VOL.Checked
            
'            Call MP3_IS_ABOUT_TO_OFF_ON_CHANGE_VOLS(False)
'        End If
        
        
        
        '-------------------------------------------
        'THIS TURN PLAYING OFF AND NEVER SET BACK ON
        '-------------------------------------------
        'PROBLEM
        '-------
        'BREAK POINT -- FIND OUT WHERE EXIT SUB TO NEXT
        '----------------------------------------------
        'NORMAL IS OKAY
        'BUT SEEM WHEN PLAY VOICE TWO TOGETHER IS PROBLEM
        'WATCH DEBUG.PRINT RESULT
        '-------------------------------------------------
        'SEEM ALWAYS ERROR PLAYING WHEN DARKNESS% VOICE
        '----------------------------------------------
        'SOLVED NOW
        '----------
        'NOT SOLVED YET
        'PROBLEM WHEN DARKNESS EXTACT AFTER A TIME TALK VOICE
        '-----------------------------------------------------
        'FLAG LATCHING NOT CORRECT YET -- TIMING IS OUT
        '----------------------------------------------
        
        
        ISSUECOMMTOSTOPWINAMP = True
        MsgResult2 = SendMessage(WINAMP2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
        'Call TimerDoubleVolume_Timer
        Debug.Print "-- -- " + Format(Now, "HH MM SS") + "  SET WINAMP PAUSED  -- HWND " + str(WINAMP2)
        
        
    
    
    End If

    ' Sub VoiceMP3Pause()

End Sub



Sub VoiceMP3Start_TIMER_timer()
    'THIS TIME SHOULD START EVEN IF NOT IN MP3 TO RESTORE VOLUME
    
    
    'THE ONLY THING RESET THIS IS AFTER CLOCK HOUR AFTER STRING
    'THAT IS DONE HERE - CHK
    
    'If WinampISPlayingAndSetToStopFlagAfterVoiceString = True Then Exit Sub
    
        
        
        
    'IF REQUEST TO VOICE IS FLAGGED THEN
    'MUST WAIT UNTIL A VOICE IS DETECTED FIRST
    
    
    
    
    'If StringVoiceToPLayActive = True Then Exit Sub
    
    'If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then Exit Sub
    'If VoiceStreamEnds = False Then Exit Sub
    
    
    If VoiceStreamActive = True Then Exit Sub
    If STRING_OF_VOICE_IN_QUE_PLAY > 0 Then Exit Sub
    
    
    'VoiceStreamEnds = False
    
    VoiceMP3Start_TIMER.Enabled = False
    
    
    Call SayTimer4BackToLower
    
    
    'MsgResult = WinAmpChkIsPlaying
    'WinAmpPowerHwnd = WinAmpChkIsPlayingHwnd
    
    'IS WINAMP
    
    If WinampISPlayingAndSetToStopFlagAfterVoice = True Then
        WinampISPlayingAndSetToStopFlagAfterVoice = False
        WinampISPlayingAndSetToStopFlagAfterVoiceString = False

    
        WINAMP2 = FindWindow("Winamp v1.x", vbNullString)
        
        MsgResult = SendMessage(WINAMP2, WM_USER, 0, ByVal ipc_isplaying)
    
    
        'OLD SUPER CODE USE
        'WinAmpPowerHwnd = WinAmpPowerHwndStore
        'MsgResult = SendMessage(WinAmpPowerHwnd, WM_USER, 0, ByVal ipc_isplaying)

        '// IPC_ISPLAYING returns the status of playback.
        '// If it returns 1, it is playing. if it returns 3, it is paused,
        '// if it returns 0, it is not playing.
    
        ' IT IS PAUSED = 3 THEN START GOING AGAIN AS WAS SET BEFORE IN
        ' Sub VoiceMP3Pause()
    
        If MsgResult = 3 Then
            MsgResult2 = SendMessage(WINAMP2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
            Call TimerDoubleVolume_Timer
        End If

    End If
    

    VoiceMP3Start_TIMER.Enabled = False

End Sub



Sub VolNibUP()
    
    Exit Sub
    
    'AND VOL SET IF MP3 PLAY
    
    tf = GetVolume(SPEAKER)
    If tf = 0 Then
        tf = SetVolume(SPEAKER, 1)
    Else
        Exit Sub
    End If
    
    MsgResult = WinAmpChkIsPlaying
    
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.
    
    If MsgResult = 1 Then
        VolumeSet1 = ATidalX.LblVol ' MUSIC VOL
    Else
        VolumeSet2 = ATidalX.LblVol ' MASTER VOL
    End If
        
    tf = SetVolume(SPEAKER, ATidalX.LblVol)
    
End Sub



'WaitWavFinish

Private Property Get SystemUptime() As TimeConv
SystemUptime = ConvertTime(GetTickCount)
End Property

Private Property Get TickTime(Tick2) As TimeConv
TickTime = ConvertTime(Tick2)
End Property
 

Private Function ConvertTime(ByVal Tick As Long) As TimeConv
 
ConvertTime.MSeconds = Tick Mod 1000
 
Tick = Tick \ 1000
ConvertTime.Days = ((Tick) \ (24 * (60 ^ 2)))
 
'If ConvertTime.Days > 0 Then Tick = (Tick - 24 * (60 ^ 2)) * ConvertTime.Days

ConvertTime.Hours = Tick \ (60 ^ 2) Mod 24
 
'If ConvertTime.Hours > 0 Then Tick = Tick - ((60 ^ 2) * ConvertTime.Hours)

ConvertTime.Minutes = (Tick \ 60) Mod 60
 
ConvertTime.Seconds = Tick Mod 60
 
End Function

Private Sub Timer8_Timer()
'THE EVENTS SEEM SOMETIMES TO BE STUCK IN A TIMER AND EVENT UNLOAD HAS TO BE CHECK HARDER
'----------------------------------------------------------------------------------------
If QuitForms = True Then
    UNLOAD_FORMER = True
End If
If aa_MainCode.TrueTerminate = True Then
    UNLOAD_FORMER = True
End If
If aa_MainCode.ALL_FORM_REQUEST_TO_END = True Then
    UNLOAD_FORMER = True
End If
If UNLOAD_FORMER = True Then
    Unload Me
    Exit Sub
End If



If SysStartTimeNewN = 0 Then
    Exit Sub
End If

'Label31 = "SysUptime " & SystemUptime.Days & ii$ & _
        SystemUptime.Hours Mod 24 & "h " & _
        SystemUptime.Minutes & "m " & _
        SystemUptime.Seconds & "s " & _
        Format$(SystemUptime.MSeconds, "000") & " mil"
'Label31 = "SysUptime"

If SystemUptime.Days > 1 Then ii$ = " Days " Else ii$ = " Day "
If Int(Now) - Int(SysStartTimeNewN) > 1 Then ii$ = " Days " Else ii$ = " Day "
Label31 = "Sys UpTime -" + str(Int(Now) - Int(SysStartTimeNewN)) + ii$
Label31 = Label31 + Format$(Now - SysStartTimeNewN, "h") + "h "
Label31 = Label31 + Format$(Now - SysStartTimeNewN, "m") + "m "
Label31 = Label31 + Format$(Now - SysStartTimeNewN, "s") + "s "
Label31 = Label31 + Format$(Abs(SystemUptime.MSeconds), "000") & " mil"

'TimeUpOld = TickTime(SysStartTimeOldN).Days

'If TickTime(SysStartTimeOldN).Days > 1 Then ii$ = " Days " Else ii$ = " Day "
'If Int(Now) - Int(SysStartTimeOldN) > 1 Then ii$ = " Days " Else ii$ = " Day "
'Label31 = "Sys UpTime -" + Str(Int(SysStartTimeNewN) - Int(SysStartTimeOldN)) + ii$
'Label31 = Label31 + Format$(Now - SysStartTimeOldN, "h") + "h "
'Label31 = Label31 + Format$(Now - SysStartTimeOldN, "m") + "m "
'Label31 = Label31 + Format$(Now - SysStartTimeOldN, "s") + "s "
'Label31 = Label31 + Format$(Abs(SystemUptime.MSeconds), "000") & " mil"




'Debug.Print Label31 + vbCrLf

If App.title = "Tidal Information..." Then Exit Sub
    
frmPassLock.Label31 = Label31

End Sub



Private Sub Timer9_Timer()
Timer9.Enabled = False
Exit Sub


If (Minute(Now)) Mod 10 = 0 Then
Call DrivesGB
End If

'If MMove > 0 Then
'    MMove = Val(Label21) + 100
'End If

Exit Sub

On Local Error Resume Next
'AhaView
If App.title = "Tidal Information..." Then Exit Sub

If FindWindow(vbNullString, "AhaView") > 0 And LockSSaver < Now Then
        AppActivate "AhaView"
End If

End Sub

Private Sub Timer9ColorCycle_Timer()

Call ColorCycle

End Sub

Public Function GetFileFromHwnd(lngHwnd) As String

'Timer12_Timer

'MsgBox getfilefromhwnd(Me.hwnd)

Dim lngProcess&, hProcess&, bla&, C&
Dim strFile As String
Dim x

strFile = String$(256, 0)
x = GetWindowThreadProcessId(lngHwnd, lngProcess)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lngProcess)
x = EnumProcessModules(hProcess, bla, 4&, C)
C = GetModuleFileNameEx(hProcess, bla, strFile, Len(strFile))
GetFileFromHwnd = Left(strFile, C)

End Function



Sub MuteVol()

Exit Sub

ATidalX.LblVol = str(0)
'Volume.SetVolume (VolumeSet)

VOLUME2.ProgressBar1.Value = ATidalX.LblVol
LblVol = ATidalX.LblVol

If App.title = "Tidal Information..." Then Exit Sub

'frmPassLock.
frmPassLock.LblVol = ATidalX.LblVol

End Sub


Public Sub TIMER_VoiceStreamActiveMAKESUREOFF_Timer()
    
If VoiceStreamActive = False Then Exit Sub
    
'1 AND 0 - NOT TRUE FALSE
'1 = THINK NOT RUNNING YES
If TTSAppMain.Voice.Status.RunningState = 1 Then
    VoiceStreamActive = False
    TIMER_VoiceStreamActiveMAKESUREOFF.Enabled = False
'Else: Stop
End If

    
Exit Sub
    
'If TTSAppMain.Voice.Status.RunningState = False Then
'
'    VoiceStreamActive = False
'    VoiceStreamActiveMAKESUREOFF_TIMER.Enabled = False
'
'End If
'
'    VoiceStreamActiveMAKESUREOFF_TIMER.Enabled = False
'    VoiceStreamEnds = True
'    VoiceStreamActive = False
'
'
'    ATidalX.TimerSimulateStringVoice.Enabled = True
'    ATidalX.Timer1.Enabled = True
'    ATidalX.Timer_CountD1Text.Enabled = True
'
'    Call ATidalX.MP3_ON_AFTER_VOICE_FINISHES
'
'    Debug.Print "VOICE ENDED BY USE OF VERIFY CHK MAKE SURE OFF CODE"
'
'
''End If
'
End Sub




Private Sub VolGetTimer_Timer()

ATidalX.LblVol = str(GetVolume(SPEAKER))

End Sub


Function FindWinPart(TT$) As Long
FindWinPart = False
'ViceVersa Pro -
Dim ash$
Dim test_hwnd As Long, _
    test_pid As Long, _
    test_thread_id As Long

Dim cText As String
'Dim Rect8 As RECT

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one
test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
br$ = ""
Do While test_hwnd <> 0
    
        'hwnd9 = GetWindowRect(test_hwnd, Rect8)
        ash$ = GetWindowTitle(test_hwnd)
'        If InStr(ash$, "Double") > 0 Then Stop
'        gws = GetWindowState(test_hwnd)
'        If gws = -1 Then gws2$ = "Window State Normal"
'        If gws = vbMinimized Then gws2$ = "Window State Minimized"
'        If gws = vbMaximized Then gws2$ = "Window State Maximized"

    
    'If (Rect8.Top > 0 And Rect8.Left > 0) Or gws = vbMinimized Then
        
        
        ash$ = GetWindowTitle(test_hwnd)
        
        XX = 0
        If TT$ = "ViceVersa Pro -" Then XX = 1
        If TT$ = "ViceVersa Pro -" And InStr(ash$, "%") > 0 Then
            XX = 0
        End If
        
        
        If ash$ <> "" And InStr(ash$, TT$) > 0 And XX = 0 Then
            FindWinPart = test_hwnd
            Exit Do
        End If
    'End If
        
'retrieve the next window
test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)

Loop
'MsgBox Str(huge) + " Windows Brought To Front"

End Function

Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim l As Long
   Dim S As String
   l = GetWindowTextLength(hWnd)
   S = Space(l + 1)
   GetWindowText hWnd, S, l + 1
   GetWindowTitle = Left$(S, l)

End Function



Private Sub TimerIceViewKill_Timer()
Exit Sub


yyg = FindWindow("ThunderRT5Form", vbNullString)

If yyg = 0 Then OYYG = 0: TimeSetIce = 0: Exit Sub

If OYYG = 0 And yyg > 0 Then TYYG = Now + TimeSerial(1, 0, 0)

OYYG = yyg

If TYYG < Now And TYYG > 0 Then TYYG = 0: GoTo jump9

Timeset2 = TimeSerial(6, 55, 0)
If TimeSetIce = 0 Then
    If Int(Now) + Timeset2 < Now Then
        TimeSetIce = Timeset2 + Int(Now) + 1
    Else
        TimeSetIce = Timeset2 + Int(Now)
    End If
End If

If Now < TimeSetIce Then Exit Sub
TimeSetIce = Timeset2 + Int(Now) + 1


jump9:

If FindWindow("ThunderRT5Form", vbNullString) = 0 Then Exit Sub

TT = cProcesses.Convert("C:\Program Files\IceView\IceView.exe", Otss, cnFromEXE Or cnToProcessID)

If TT = False Then Exit Sub

'Process_Kill (Otss)
OYYG = 0

End Sub

Private Sub TimerVideoKill_Timer()

Exit Sub

If KeyIdleTimeVideo > Now Then Exit Sub

hWnd2 = WinAmpVideoChkIsOpen

If hWnd2 = 0 Then
TimeSetVid = 0
LockVideoWinAmp = 0
OYYG2 = 0
Exit Sub
End If
'Exit Sub

LockVideoWinAmp = hWnd2

If LockVideoWinAmp = 0 Then OYYG2 = 0: Exit Sub

If OYYG2 = 0 And LockVideoWinAmp > 0 Then TYYG2 = Now + TimeSerial(0, 20, 0)

OYYG2 = LockVideoWinAmp

If TYYG2 < Now And TYYG2 > 0 Then TYYG2 = 0: GoTo JUMP10
    
Exit Sub
    
'Either Stops on 20 mins or at the time below
'if not rem out exit sub
    
Timeset2 = TimeSerial(6, 55, 0)
If TimeSetVid = 0 Then
    If Int(Now) + Timeset2 < Now Then
        TimeSetVid = Timeset2 + Int(Now) + 1
    Else
        TimeSetVid = Timeset2 + Int(Now)
    End If
End If

If Now < TimeSetVid Then Exit Sub
TimeSetVid = Timeset2 + Int(Now) + 1

JUMP10:

TT = cProcesses.Convert(LockVideoWinAmp, Otss, cnFromhWnd Or cnToProcessID)
If TT = False Then Exit Sub

Call CloseProgramHwnd(WinAmpVideoChkIsOpen)

'Process_Kill (Otss)
LockVideoWinAmp = 0
OYYG2 = 0

End Sub


Public Sub CloseProgram(ByVal Caption As String)
Dim Handle
 Handle = FindWindow(vbNullString, Caption)
 If Handle = 0 Then Exit Sub
 SendMessage Handle, WM_CLOSE, 0&, 0&

End Sub

Public Sub CloseProgramHwnd(ByVal Handle As Long)

 If Handle = 0 Then Exit Sub
 SendMessage Handle, WM_CLOSE, 0&, 0&

End Sub



Public Function IsWinAmpPlayingBeforeVoiceStartAndStop() As Long

Dim Retval As Long, ProcessID As Long, ThreadID As Long

Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
Dim WinClass As String, WinTitle As String


'NEW MOD CODE -- DONT BIG DEAL ON THIS FOR NOW
'WinAmpChkIsPlaying ' = False
IsWinAmpPlayingBeforeVoiceStartAndStop = False

WINAMP2 = FindWindow("Winamp v1.x", vbNullString)

If WINAMP2 = 0 Then Exit Function
    
MsgResult = SendMessage(WINAMP2, WM_USER, 0, ByVal ipc_isplaying)

IsWinAmpPlayingBeforeVoiceStartAndStop = MsgResult
'If IsWinAmpPlayingBeforeVoiceStartAndStop <> 1 Then IsWinAmpPlayingBeforeVoiceStartAndStop = False
'If IsWinAmpPlayingBeforeVoiceStartAndStop = 1 Then IsWinAmpPlayingBeforeVoiceStartAndStop = True

If oWinAmpChkIsPlaying <> MsgResult Then
    If MsgResult = 1 Then
'        TTSAppMain.RateSldr = 0
    Else
'        TTSAppMain.RateSldr = 2
    End If
End If

oWinAmpChkIsPlaying = MsgResult
Exit Function

End Function



Public Function WinAmpChkIsPlaying()

If ATidalX.MMControl8.mode = 526 Then
    MsgResult = 1
    GoTo jump
End If

Dim Retval As Long, ProcessID As Long, ThreadID As Long
Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
Dim WinClass As String, WinTitle As String
'Retval = GetClassName(WinArray(R), WinClassBuf, 255)
'WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces


'NEW MOD CODE -- DONT BIG DEAL ON THIS FOR NOW
WinAmpChkIsPlaying = False
WINAMP2 = FindWindow("Winamp v1.x", vbNullString)

If WINAMP2 = 0 Then Exit Function
'If WINAMP2 = 0 Then OWinAmp2 = 0: Exit Function
    
MsgResult = SendMessage(WINAMP2, WM_USER, 0, ByVal ipc_isplaying)

jump:
WinAmpChkIsPlaying = MsgResult
'If WinAmpChkIsPlaying <> 1 Then WinAmpChkIsPlaying = False
'If WinAmpChkIsPlaying = 1 Then WinAmpChkIsPlaying = True

If oWinAmpChkIsPlaying <> MsgResult Then
    If MsgResult = 1 Then
'        TTSAppMain.RateSldr = 0
    Else
'        TTSAppMain.RateSldr = 2
    End If
End If

oWinAmpChkIsPlaying = MsgResult

'Exit Function

End Function


Function WinAmpChkIsPlayingxxOLD()


Dim Retval As Long, ProcessID As Long, ThreadID As Long

Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
Dim WinClass As String, WinTitle As String

'### --------------------------



    WinAmpChkIsPlayingHwnd = 0
    WinAmpChkIsPlaying = 0
    
    MsgResult = 0

    WINAMP2 = FindWindow("Winamp v1.x", vbNullString)
    If WINAMP2 = 0 Then OWinAmp2 = 0: Exit Function
    
    If WINAMP2 = WinAmpStore Then
        MsgResult = SendMessage(WINAMP2, WM_USER, 0, ByVal ipc_isplaying)
        '// IPC_ISPLAYING returns the status of playback.
        '// If it returns 1, it is playing. if it returns 3, it is paused,
        '// if it returns 0, it is not playing.
        If MsgResult = 1 Then
            WinAmpChkIsPlayingHwnd = WINAMP2
            WinAmpChkIsPlaying = MsgResult
            Exit Function
        End If
    End If
    If WinAmpStore > 0 Then
        MsgResult = SendMessage(WinAmpStore, WM_USER, 0, ByVal ipc_isplaying)
        '// IPC_ISPLAYING returns the status of playback.
        '// If it returns 1, it is playing. if it returns 3, it is paused,
        '// if it returns 0, it is not playing.
        If MsgResult = 1 Then
            WinAmpChkIsPlayingHwnd = WinAmpStore
            WinAmpChkIsPlaying = MsgResult
            Exit Function
        End If
    End If
    
    
    
    
    
    If OWinAmp2 <> WINAMP2 And WINAMP2 > 0 Then
        cuk = 0
        For r = 0 To UBound(WinArray)
            If WinArray(r) = WINAMP2 Then cuk = 1: Exit For
        Next
        If cuk = 0 Then
            ReDim Preserve WinArray(UBound(WinArray) + 1)
            WinArray(UBound(WinArray)) = WINAMP2
        End If
            
        squash = False
        For r = 0 To UBound(WinArray)
            If WinArray(r) > 0 Then
                pot = GetWindowTitle(WinArray(r))
                If pot = "" Then
                    WinArray(r) = 0
                    squash = True
                End If
            End If
        Next
        
        If squash = True Then
            xc = 0
            For r = 0 To UBound(WinArray)
                If WinArray(r) > 0 Then
                    WinArray(xc) = WinArray(r)
                    xc = xc + 1
                End If
            Next
            xcd = xc
            If xcd = 0 Then xcd = 1
            ReDim Preserve WinArray(xcd - 1)
        End If
    End If
    
    OWinAmp2 = WINAMP2

    WinAmpChkIsPlaying = 0
                
    For r = 0 To UBound(WinArray)
        If WinArray(r) > 0 Then
            Retval = GetClassName(WinArray(r), WinClassBuf, 255)
            WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
            
        
            If WinClass = "Winamp v1.x" Then
                MsgResult = SendMessage(WinArray(r), WM_USER, 0, ByVal ipc_isplaying)
                '// IPC_ISPLAYING returns the status of playback.
                '// If it returns 1, it is playing. if it returns 3, it is paused,
                '// if it returns 0, it is not playing.
                
                If MsgResult = 1 Then 'Or MsgResult = 3 Then
                    WinAmpChkIsPlaying = True
                    WinAmpChkIsPlayingHwnd = WinArray(r)
                    
                    WinAmpStore = WinArray(r)
                    Exit For
                End If
                If MsgResult = 0 Or MsgResult = 3 Then 'Or MsgResult = 3 Then
                    WinAmpChkIsPlaying = False
                    WinAmpChkIsPlayingHwnd = WinArray(r)
                    WinAmpStore = WinArray(r)
                    Exit For
                End If
            
            
            End If
        End If
    Next

If MsgResult = 0 Or MsgResult = 3 Then
    WinAmpChkIsPlayingHwnd = WinAmpStore
    WinAmpChkIsPlaying = MsgResult
    
    If MsgResult = 1 Then WinAmpStoppedPlayTime = 0
    If MsgResult <> 1 And WinAmpStoppedPlayTime = 0 Then
        WinAmpStoppedPlayTime = Now
    End If
End If

End Function


Function WinAmpVideoChkIsOpen()

Dim Retval As Long, ProcessID As Long, ThreadID As Long
Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
Dim WinClass As String, WinTitle As String

WinAmpVideoChkIsOpen = 0

WINAMP2 = FindWindow("Winamp Video", vbNullString)

If WINAMP2 = 0 Then Exit Function

    If OWinAmp2 <> WINAMP2 And WINAMP2 > 0 Then
        cuk = 0
        For r = 0 To UBound(WinArray)
            If WinArray(r) = WINAMP2 Then cuk = 1: Exit For
        Next
        If cuk = 0 Then
            ReDim Preserve WinArray(UBound(WinArray) + 1)
            WinArray(UBound(WinArray)) = WINAMP2
        End If
    End If
    
    
        OWinAmp2 = WINAMP2
        For r = 0 To UBound(WinArray)
            If WinArray(r) > 0 Then

             Retval = GetClassName(WinArray(r), WinClassBuf, 255)
             WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
             'RetVal = GetWindowText(lhWnd, WinTitleBuf, 255)
             'WinTitle = StripNulls(WinTitleBuf)
                If WinClass = "Winamp Video" Then
                    'lngStyle = GetWindowLong(WinArray(R), GWL_STYLE)
                    'fIsDBCVisible = ((lngStyle And WS_VISIBLE) = WS_VISIBLE)
                    fIsDBCVisible = IsWindowVisible(WinArray(r))
                    If fIsDBCVisible = 1 Then
                        'MsgResult = SendMessage(WinArray(R), WM_USER, 0, ByVal ipc_isplaying)
                        
                        WinAmpVideoChkIsOpen = WinArray(r)
                        Exit For
                    End If
                End If
            End If
        Next
    
    

End Function



Private Sub VCode5_TIMER_TIMER()
'Call StripCodeVLogg1StRun

If 1 = 2 Then

    If OldLabel23 <> Label23.Caption Then
        On Local Error Resume Next
        FR = FreeFile
        Open RamDrive + ":\temp\KeyHit-Tidal.txt" For Output As #FR
            Print #FR, Label23.Caption;
        Close #FR
    '    FR = FreeFile
    '    Open "Q:\temp\KeyHit-Tidal2.txt" For Output As #FR
    '        Print #FR, Label23.Caption;
    '    Close #FR
        On Local Error GoTo 0
    End If
End If

OldLabel23 = Label23.Caption

If KeyIdleTime = 0 Then Exit Sub

If KeyIdleTime > Now And KeyIdleTime2 < Now Then
    'GoTo JumpHere
    'Exit Sub
End If

If KeyIdleTime > Now Then
    Exit Sub
End If

JumpHere:

'NotKnot = Not NotKnot
'If NotKnot = True Then Label40.BackColor = Label39.BackColor
'If NotKnot = False Then Label40.BackColor = Label38.BackColor
'Label40 = Trim(Str(qk))

KeyIdleTime = 0
KeyIdleTime2 = Now + TimeSerial(0, 1, 0)

If VCodeTT_ = "" Then Exit Sub

Dim FR1

'--------------If tt8 = 0 Then tt8 = FileInUse(tx6$)

'01 OF 03
i = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Key Logger Text-Stripe.txt"
FILE_NAME_KEY_DATA_LOGGER = i
If FileInUse(FILE_NAME_KEY_DATA_LOGGER) = True Then Exit Sub

'02 OF 03
i = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Key Logger Text.txt"
FILE_NAME_KEY_DATA_LOGGER = i
If FileInUse(FILE_NAME_KEY_DATA_LOGGER) = True Then Exit Sub

'03 OF 03
i = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Key Logger Text " + Format$(KeyLoggDate, "YYYY-mm-dd HH-MM-SS") + ".txt"
FILE_NAME_KEY_DATA_LOGGER = i
If FileInUse(FILE_NAME_KEY_DATA_LOGGER) = True Then Exit Sub

'01 OF 03
Call StripCodeVLogg

'tt98 = FileInUse(tx5$)
'tt99 = FileInUse(tx6$)
'If tt99 = True Or tt98 = True Then Exit Sub

On Local Error GoTo ExitSubNow

'02 OF 03
FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Key Logger Text.txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4

FR1 = FreeFile
tx5$ = FILENAME_IN_USE_CHECK
Open FILENAME_IN_USE_CHECK For Append As #FR1
    Print #FR1, VCodeTT_;
Close #FR1

If Err.Number > 0 Then
    VCodeTT1 = VCodeTT1 + VCodeTT
Else
    VCodeTT1 = ""
End If

Err.Clear
'03 OF 03
FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Key Logger Text " + Format$(KeyLoggDate, "YYYY-mm-dd HH-MM-SS") + ".txt"
FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
tx6$ = FILENAME_IN_USE_CHECK
FR1 = FreeFile
Open FILENAME_IN_USE_CHECK For Append As #FR1
    Print #FR1, VCodeTT_;
Close #FR1

If Err.Number > 0 Then
    VCodeTT2 = VCodeTT2 + VCodeTT_
Else
    VCodeTT2 = ""
End If

If Err.Number = 0 Then
    VCodeTT_ = ""
End If

'---------------------



'tx7$ = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Key Logger Text List.txt"
'Fr1 = FreeFile
'Open tx7$ For Append As #Fr1
'    Print #Fr1, KeyLog2;
'Close #Fr1

'KeyLog2 = ""


ExitSubNow:
End Sub


Sub StripCodeVLogg()
    
    On Local Error GoTo ErrTrap5
    
    If OlVcodeTT$ <> "" Then R3$ = Mid$(OlVcodeTT$, Len(OlVcodeTT$))
'    r4$ = OlVcodeTT$
    r5$ = VCodeTT$
    
    For r = 1 To Len(r5$)
        xc = 0
        If Mid$(r5$, r, 1) = "" Then xc = 1 'esc
        If Mid$(r5$, r, 1) = "" Then xc = 1 'FF
        'If Mid$(b1$, r - 1, 1) = "_" Then XC = 1
        If xc = 1 Then
            r5$ = Mid$(r5$, 1, r - 1) + Mid$(r5$, r + 1)
            r = r - 1
        End If
    Next
    
    If Mid$(r5$, 1, 1) = Chr(8) Then
        r5$ = R3$ + Mid$(r5$, 2)
    End If
    
    For r = 1 To Len(r5$)
        xc = 0
        If Mid$(r5$, r, 1) = Chr(8) Then xc = 1
        'If Mid$(b1$, r - 1, 1) = "_" Then XC = 1
        If xc = 1 Then
            If r = 1 Then r5$ = Mid$(r5$, 1, r - 1) + Mid$(r5$, r + 1): r = r - 1
            If r > 1 Then r5$ = Mid$(r5$, 1, r - 2) + Mid$(r5$, r + 1): r = r - 2
        End If
    Next
    
    'R5$ = "123" + Chr$(10) + "12345"
    For r = 1 To Len(r5$)
        xc = 0
        'If InStr(" _-()&{[]}!'@;#~><!""£$%^&*()_-+={}~@:?><[]#';/.,\`¬"",\/1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ" + vbCr, Mid$(r5$, R, 1)) > 0 Then xc = 1
        If InStr(" _-)[]}!'@;#~><""^*()_-+={}~@:?><[]#';/\`¬,\/1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ" + vbCr, Mid$(r5$, r, 1)) > 0 Then xc = 1
        'If InStr("'¡¢¾¤Ü¿ßYß½ ", Mid$(r5$, R, 1)) > 0 Then XC = 0
        If InStr("'¡¢¾¤Ü¿ßß½ ", Mid$(r5$, r, 1)) > 0 Then xc = 0
        'If Mid$(b1$, r - 1, 1) = "_" Then XC = 1
        If Mid$(r5$, r, 1) = " " Then xc = 1
        
        
        If xc = 0 Then
            r5$ = Mid$(r5$, 1, r - 1) + Mid$(r5$, r + 1)
        End If
    Next
    
    If R3$ <> "" Then r5$ = R3$ + LCase(r5$)
    
    For r = 1 To Len(r5$) - 1
        xc = 0
        If InStr(" _-()&{[]}!'@;#~><!""£$%^&*()_-+={}~@:?><[]#';/.,\`¬"",\/1234567890" + vbCr, Mid$(r5$, r, 1)) > 0 Then xc = 1
        'If Mid$(b1$, r - 1, 1) = "_" Then XC = 1
        If Mid$(r5$, r, 1) = " " Then xc = 0
        If xc = 1 Then Mid$(r5$, r + 1, 1) = UCase(Mid$(r5$, r + 1, 1))
    Next
    
    If Len(r5$) = 1 And R3$ <> "" Then r5$ = ""
    If R3$ <> "" And Len(r5$) > 1 Then r5$ = Mid$(r5$, 2)
    
    'If r5$ <> "" Then Mid$(r5$, 1, 1) = UCase(Mid$(r5$, 1, 1))
    
    VCodeTT2 = r5$
        
    If r5$ <> "" Then OlVcodeTT$ = VCodeTT2

    Dim HHS$
    If VCodeTT2 <> "" Then
        ash$ = GetActiveWindowTitle(True)
        
        
        If OldActiveWindow$ <> ash$ Then
            HHS$ = vbCrLf + Format$(Now, "DD-MM-YY HH:MM:SS ") + " -- " + ash + vbCrLf
            HHS$ = HHS$ + "-------------------------------------------------------" + vbCrLf
        End If
        
        HHS$ = HHS$ + VCodeTT2
        
        VCodeTTB2$ = VCodeTTB2$ + HHS$

        FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Key Logger Text-Stripe.txt"
        FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
        DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
        FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
        tx8$ = FILENAME_IN_USE_CHECK
        FR1 = FREFILE
        Open FILENAME_IN_USE_CHECK For Append As FR1
            Print #FR1, VCodeTTB2$
        Close #FR1
        VCodeTTB2$ = ""
        
        OldActiveWindow$ = ash$
    End If

ErrTrap5:
End Sub

Sub StripCodeVLogg1StRun()

'Exit Sub


A1$ = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\"
A2$ = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\"
FS.CopyFile A1$, A2$

Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Key Logger Text-Stripe.txt" For Binary As #1
r5$ = Space$(LOF(1))
Get #1, 1, r5$
Close #1
    
    'do that with instr's later
    
    For r = 1 To Len(r5$)
        xc = 0
        If Mid$(r5$, r, 1) = "" Then xc = 1 'esc
        If Mid$(r5$, r, 1) = "" Then xc = 1 'FF
        'If Mid$(b1$, r - 1, 1) = "_" Then XC = 1
        If xc = 1 Then
            r5$ = Mid$(r5$, 1, r - 1) + Mid$(r5$, r + 1)
            r = r - 1
        End If
    Next
    
    
    
    For r = 1 To Len(r5$)
        xc = 0
        If Mid$(r5$, r, 1) = Chr(8) Then xc = 1
        'If Mid$(b1$, r - 1, 1) = "_" Then XC = 1
        If xc = 1 Then
            r5$ = Mid$(r5$, 1, r - 2) + Mid$(r5$, r + 1)
            r = r - 2
        End If
    Next
    
    
    
    'R5$ = "123" + Chr$(10) + "12345"
    For r = 1 To Len(r5$)
        xc = 0
        If InStr(" _-()&{[]}!'@;#~><!""£$%^&*()_-+={}~@:?><[]#';/.,\`¬"",\/1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ" + vbCr, Mid$(r5$, r, 1)) > 0 Then xc = 1
        'If Mid$(b1$, r - 1, 1) = "_" Then XC = 1
        If xc = 0 Then
            r5$ = Mid$(r5$, 1, r - 1) + Mid$(r5$, r + 1)
        End If
    Next
    
    
    
    
    r5$ = LCase(r5$)
    
    For r = 2 To Len(r5$)
        xc = 0
        If InStr(" _-()&{[]}!'@;#~><,\/1234567890" + vbCr, Mid$(r5$, r - 1, 1)) > 0 Then xc = 1
        'If Mid$(b1$, r - 1, 1) = "_" Then XC = 1
        If xc = 1 Then Mid$(r5$, r, 1) = UCase(Mid$(r5$, r, 1))
    Next
    
    Mid$(r5$, 1, 1) = UCase(Mid$(r5$, 1, 1))
    
    
    
On Local Error Resume Next


Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Key Logger Text-Stripe.txt" For Output As #1
Print #1, r5$
Close #1


End Sub

Public Sub MNU_WinampFadeOut_Click()

XmarkSpot = True
WinVol5 = 100
MNU_WinampFadeOut.Enabled = True

End Sub

Private Sub WinAmpfadeout_Timer_Timer()

If XmarkSpot = False Then Exit Sub

If WinVol5 > 0 Then WinVol5 = WinVol5 - 1

If WinVol5 > 0 Then
    WinampHwnd = FindWindow("Winamp v1.x", vbNullString)
    MsgResult = SendMessage(WinampHwnd, WM_COMMAND, Lowervolumeby1percent, ByVal XcNul)
End If

WinampFadeOut.Enabled = False

Exit Sub

If WinVol5 < 1 Then
    WinampHwnd = FindWindow("Winamp v1.x", vbNullString)
    'pause winamp??
    MsgResult = SendMessage(WinampHwnd, WM_USER, 0, ByVal ipc_isplaying)
    
    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.
    If MsgResult = 1 Then
        MsgResult = SendMessage(WinampHwnd, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
    End If
   XmarkSpot = False
    For r = 0 To 100
        MsgResult = SendMessage(WinampHwnd, WM_COMMAND, Raisevolumeby1percent, ByVal XcNul)
    Next
End If

WinampFadeOut.Enabled = False

End Sub






Public Sub SoundOffOn(TTS$)

If InCode = True Then Exit Sub
If Menu.MuteAllMal = vbChecked Then Exit Sub
InCode = True



'this for pulse sound
'sometimes the sound on winamp volume is connected to main out put so this does not sound goes mute
    
'if sound effect mute then stop playback winamp hitt sound then sorted
'WinAmpChkIsPlayingHwnd
'WinampHwnd = FindWindow("Winamp v1.x", vbNullString)
MsgResult2 = WinAmpChkIsPlaying
WinampHwndVM2 = WinAmpChkIsPlayingHwnd
For r = 1 To 100
    MsgResult = SendMessage(WinampHwndVM2, WM_COMMAND, Lowervolumeby1percent, ByVal XcNul)
Next

    'WinampHwndVM2 = FindWindow("Winamp v1.x", vbNullString)
    'pause winamp??
    'MsgResult2 = SendMessage(WinampHwndVM2, WM_USER, 0, ByVal ipc_isplaying)
    
    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.
    If MsgResult2 = 1 Then
        MsgResult = SendMessage(WinampHwndVM2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
    End If


VolumeSet5 = GetVolume(SPEAKER)
    

If TTS$ = "OFF" Then
    tf = SetVolume(SPEAKER, 1)
Else
    tf = SetVolume(SPEAKER, 15)
End If


'Call VoiceWait
'Call VoiceMP3Pause
Call SpeakVoice(TTS$)
   
    
Do
    DoEvents
    If TTSAppMain.Voice.Status.RunningState = 1 Then Exit Do
Loop Until 1 = 2
    
    
    
tf = SetVolume(SPEAKER, VolumeSet5)

'WinampHwnd = FindWindow("Winamp v1.x", vbNullString)
'MsgResult2 = WinAmpChkIsPlaying
'WinampHwndVM2 = WinAmpChkIsPlayingHwnd

For r = 1 To 100
    MsgResult = SendMessage(WinampHwndVM2, WM_COMMAND, Raisevolumeby1percent, ByVal XcNul)
Next
    
    'start play if paused
    If MsgResult2 = 1 Then
        MsgResult = SendMessage(WinampHwndVM2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
    End If

InCode = False

End Sub

Public Function GetParentWindowTitle(Test_Hwnd5 As Long, ByVal ReturnParent As Boolean) As String
   Dim i As Long
   Dim j As Long
   i = Test_Hwnd5 'GetForegroundWindow
   If ReturnParent Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
   i = j
   End If
   GetParentWindowTitle = GetWindowTitle(i)
    Test_Hwnd5 = i
End Function

Sub SolarTimeList()

Exit Sub

Dim srs As New clsSunRiseSet
srs.City = "London, England"

'------------
For P = 1 To 2
oy = 0
TT2 = "": TT3 = ""
y5 = 0: y6 = 20
f5 = 0: f6 = 20
TT$ = ""
If P = 1 Then tt1$ = "Solar Noon As English Time GMT to BST -- Goes Jumpy in Results at Cross over Time" + vbCrLf + vbCrLf
If P = 2 Then tt1$ = TT$ + "Solar Noon As Uni-Time (UTC, UT, UT1, GMT, Solar, Atomic) More Smooth Idea" + vbCrLf + vbCrLf


For r = DateSerial(2009, 1, 1) To DateSerial(2016, 1, 1)
srs.DateDay = r
If P = 1 Then srs.DaySavings = DaySavingsSub(srs.DateDay)
If P = 2 Then srs.DaySavings = DaySavingsSub(DateSerial(2009, 1, 1))
If srs.DaySavings = False Then GG = " GMT"
If srs.DaySavings = True Then GG = " BST"
srs.CalculateSun
tool5 = srs.SolarNoon

i5 = 0: i6 = 0
s5 = 0: s6 = 0
tt5 = tool5 - Int(tool5)
If tt5 > y5 Then y5 = tt5: i5 = 1
If tt5 < y6 Then y6 = tt5: i6 = 1
If tt5 > f5 Then f5 = tt5: s5 = 1
If tt5 < f6 Then f6 = tt5: s6 = 1

If tt5 > tto Then
H = " > Higher"
Else
H = " < Lower"
End If
If tt5 = tto Then H = " = Equal"

tto = tt5
HH = "Solar Noon = " + Format$(tool5, "DD-MM-YYYY HH:MM:SSampm") + H + GG + vbCrLf
TT$ = TT$ + HH
If i5 = 1 Then ij1 = HH
If i6 = 1 Then ij2 = HH
If s5 = 1 Then iw1 = HH
If s6 = 1 Then iw2 = HH

If oy = 0 Then oy = Year(r)
If Year(r) <> oy Then
TT2 = TT2 + "Lowest" + str(Year(r) - 1) + " = " + ij2
TT2 = TT2 + "Highest" + str(Year(r) - 1) + " = " + ij1
y5 = 0: y6 = 20
End If
oy = Year(r)

Next

TT3 = TT3 + "All Time Lowest = " + iw2
TT3 = TT3 + "All Time Highest = " + iw1

If P = 1 Then Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Solar Noon List GMT & BST 2009 - 2015.txt" For Output As #1
If P = 2 Then Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Solar Noon List GMT 2009 - 2015.txt" For Output As #1
Print #1, "www.MatthewLan.com --- Matt.Lan@btinternet.com"
Print #1, "Text Produced :: "; Format$(Now, "DDD DD-MMM-YYYY HH:MM:SSa/p") + vbCrLf
Print #1, "UK -- Brighton & Hove - Hove Actually -- Exaclty On the Center of the Roof of My Room Latitude,Longitude = Time Zone " + vbCrLf
Print #1, "You Can see The Different Eqinoxes/Solistice Produce Diff Lower/Higher Times " + vbCrLf
Print #1, tt1$;
Print #1, TT2
Print #1, TT3
Print #1, TT$;
Close #1

Next


End

End Sub

Private Sub WinampFastForwardTimer_Timer()
'WAFastFF = True
ATidalX.WinampFastForwardTimer.Enabled = False
Exit Sub
Call aa_MainCode.WinAmpCommands(4)

End Sub




Private Sub Timer_ART_TRIP_WIRE_Timer()
Timer_ART_TRIP_WIRE.Enabled = False
Exit Sub
'i = LastModifiedToCurrent("K:\TEMP\ART_PROG_TRIP_WIRE.txt")

Dim HH As Long, Otss As Long

HH = FindWindow(vbNullString, "Jpeg Slides PJpeg ART.exe - Application Error")

If HH > 0 Then
    AppActivate "Jpeg Slides PJpeg ART.exe - Application Error", False
    SendKeys "{ENTER}", True
End If
HH = FindWindow(vbNullString, "Jpeg Encoder -- 2009 The One --: Jpeg Slides PJpeg ART.exe - Application Error")
If HH > 0 Then
    AppActivate "Jpeg Slides PJpeg ART.exe - Application Error", False
    SendKeys "{ENTER}", True
End If


HH = FindWindow(vbNullString, "JPEG Encoder -- Shockingly Good Photo ART -- [-:¦:-•:*''' The One '''*:•-:¦:-]")
If HH > 0 Then XHH = True
'If HH = 0 And XHH = True Then End
If HH = 0 And XHH = False Then Exit Sub


'If FS.FileExists("K:\TEMP\ART_PROG_TRIP_WIRE.txt") = False And XHH = True Then End
If FS.FileExists("K:\TEMP\ART_PROG_TRIP_WIRE.txt") = False Then Exit Sub

Set F = FS.GetFile("K:\TEMP\ART_PROG_TRIP_WIRE.txt")
i = F.DateLastModified
If OLDi <> i Then iTRIP = Now + TimeSerial(0, 0, 30)
OLDi = i
'Debug.Print i, OLDi

If iTRIP > Now Then Exit Sub
XHH = False

'TT = cProcesses.Convert(HH, OTSS, cnFromhWnd Or cnToProcessID)
Otss = -1
TT = cProcesses.GetEXEID(Otss, "D:\VB6\VB-NT\00_Best_VB_04\JPEG-Encoder-ART\Jpeg Slides PJpeg ART.exe")

If TT = True Then Process_Kill (Otss)

XX = 0
Do
    XX = XX + 1
    Otss = -1
    TT = cProcesses.GetEXEID(Otss, "D:\VB6\VB-NT\00_Best_VB_04\JPEG-Encoder-ART\Jpeg Slides PJpeg ART.exe")
    If TT = True Then Call Sleep(500)
    If XX > 100 Then End
    'Debug.Print XX
Loop Until TT = False

HH = FindWindow(vbNullString, "Jpeg Slides PJpeg ART.exe - Application Error")
If HH > 0 Then
    AppActivate "Jpeg Slides PJpeg ART.exe - Application Error", True
    SendKeys "{ENTER}", True
End If

Shell "D:\VB6\VB-NT\00_Best_VB_04\JPEG-Encoder-ART\Jpeg Slides PJpeg ART.exe", vbNormalFocus
XHH = False

End Sub


'-------------------------------------------------
'-------------------------------------------------
Private Sub Timer_1_MINUTE_Timer()

'Timer_1_MINUTE.Interval = 1000

End Sub

'-------------------------------------------------
'-------------------------------------------------
Sub ARCHIVE_OF_PROCEDURE_ABOVE_VB_PROJECT_CHECKDATE()

'MsgBox "VB PROJECT DATE IS CHANGED MAYBE NEWER", vbMsgBoxSetForeground

'------------------------------------------------
'DO THE WRITE INFO ABOUT RELOAD MIRROR EXE FOLDER
'------------------------------------------------
TIMER_RETRY_WRITE_INFO_UNTIL_DONE1.Enabled = True
TIMER_RETRY_WRITE_INFO_UNTIL_DONE2.Enabled = True

'--------
'02 OF 02
'----------------------------------------------------------------
'USE THIS CODE SNIP TO RELOAD ANY APPEND EXE IF WAITING FROM INFO
'----------------------------------------------------------------
'BUT NEVER RELOAD OUR OWN CODE -- LET ANOTHER PROGRAM
'----------------------------------------------------------------
    
'If Dir("D:\VB6-EXE'S\#00 RELOAD_APPEND_SCRIPT\") = "" Then Exit Sub
'
'File3.Path = "D:\VB6-EXE'S\#00 RELOAD_APPEND_SCRIPT\"
'File3.FileName = "*.*"
'File3.Refresh
'
'On Error Resume Next
'If File3.ListCount > 0 Then
'    For R = 0 To File3.ListCount
'        If InStr("-" + File3.List(R), "RELOAD ME -- ") > 0 Then
'
'            FR1 = FreeFile
'            Open File3.Path + "\" + File3.List(R) For Input As #FR1
'                Line Input #FR1, RELOAD_VAR_TO_DO
'            Close #FR1
'
'            If Err.Number > 0 Then Exit Sub
'            'Debug.Print RELOAD_VAR_TO_DO
'
'            Err.Clear
'
'            '------------------------------------------------------
'            'GREEN LIGHT GOT - AND THEN - ALL OVER FOR THIS PROGRAM
'            '------------------------------------------------------
'            'HANDED OVER TO EXE RELOADER SHELL
'            '------------------------------------------------------
'            If Dir(File3.Path + "\" + File3.List(R) + ".GREEN.TXT") = "" Then
'                If UCase(RELOAD_VAR_TO_DO) <> UCase(PATH_FILE_NAME1) Then
'                    FR1 = FreeFile
'                    Open File3.Path + "\" + File3.List(R) + ".GREEN.TXT" For Output As #FR1
'                    Close #FR1
'                    If Err.Number > 0 Then Exit Sub
'
'                    Shell "D:\VB6\VB-NT\00_Best_VB_01\RELOAD_NETWORK_SYNC_EXE\RELOAD_Network_Sync_EXE_VB_MIRROR.exe """ + RELOAD_VAR_TO_DO + """"
'                End If
'            End If
'        End If
'    Next
'End If
'
End Sub
'-------------------------------------------------
'-------------------------------------------------

