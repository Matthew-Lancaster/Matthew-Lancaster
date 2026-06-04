VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form CID_Run_Me 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CID Run Me"
   ClientHeight    =   10725
   ClientLeft      =   21540
   ClientTop       =   16950
   ClientWidth     =   13515
   Icon            =   "frmCentral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10725
   ScaleWidth      =   13515
   WindowState     =   1  'Minimized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   240
      Left            =   1428
      TabIndex        =   71
      Top             =   0
      Width           =   4512
      _ExtentX        =   7964
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer_1_MINUTE 
      Interval        =   1000
      Left            =   8592
      Top             =   216
   End
   Begin VB.Timer Timer_SAVE_AGRO_STORE2_WINAMP 
      Interval        =   2000
      Left            =   9888
      Top             =   84
   End
   Begin VB.Timer Timer_EXPLORER_GONE 
      Interval        =   1000
      Left            =   9384
      Top             =   156
   End
   Begin VB.Timer Timer_For_Menu_Gone 
      Interval        =   5000
      Left            =   8976
      Top             =   324
   End
   Begin VB.Timer Timer_KEY_ESCAPE 
      Interval        =   1
      Left            =   7140
      Top             =   1944
   End
   Begin VB.Timer TIMER_SHELL_PROGRAM_SET_EXE_AT_STARTER 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   12696
      Top             =   696
   End
   Begin VB.Timer Timer_FORM_UNLOAD 
      Interval        =   100
      Left            =   12696
      Top             =   168
   End
   Begin VB.Timer TimerSTOP_ON_FOLDER 
      Interval        =   500
      Left            =   7104
      Top             =   870
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H00716311&
      Caption         =   "Stop in 3 Hour"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3984
      TabIndex        =   57
      Top             =   3912
      Width           =   2868
   End
   Begin VB.Timer Timer16 
      Interval        =   2000
      Left            =   9804
      Top             =   4548
   End
   Begin VB.Timer zzCheckTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12840
      Top             =   4935
   End
   Begin VB.Timer Timer15 
      Interval        =   100
      Left            =   12210
      Top             =   4785
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11610
      Top             =   2205
   End
   Begin VB.Timer TimerDay 
      Interval        =   1000
      Left            =   10095
      Top             =   2190
   End
   Begin VB.ListBox Lst1 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6960
      TabIndex        =   39
      Top             =   3204
      Width           =   795
   End
   Begin VB.Timer Timer13 
      Interval        =   1000
      Left            =   11940
      Top             =   3090
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12675
      Top             =   4305
   End
   Begin VB.Timer TimerFileSave2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   12195
      Top             =   4245
   End
   Begin VB.Timer TimerFileSave 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11610
      Top             =   4275
   End
   Begin VB.Timer Timer12 
      Interval        =   20
      Left            =   10740
      Top             =   4005
   End
   Begin VB.ListBox List4 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7560
      TabIndex        =   37
      Top             =   4668
      Width           =   2025
   End
   Begin VB.Timer Timer11 
      Interval        =   500
      Left            =   10035
      Top             =   3990
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00716311&
      Caption         =   "Stop on Folder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3984
      TabIndex        =   33
      Top             =   4152
      Width           =   2868
   End
   Begin VB.Timer VocdeTimer 
      Interval        =   1000
      Left            =   12540
      Top             =   3585
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00716311&
      Caption         =   "Stop in 1 Hour"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3984
      TabIndex        =   32
      Top             =   3684
      Width           =   2868
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00716311&
      Caption         =   "Stop in 10 Sec"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3984
      TabIndex        =   31
      Top             =   3204
      Width           =   2868
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   12120
      Top             =   3585
   End
   Begin VB.Timer Timer9 
      Interval        =   500
      Left            =   11700
      Top             =   3585
   End
   Begin VB.Timer Timer8 
      Interval        =   500
      Left            =   11265
      Top             =   3585
   End
   Begin VB.Timer Timer_Idle 
      Interval        =   1000
      Left            =   10860
      Top             =   2715
   End
   Begin VB.Timer Timer7 
      Interval        =   500
      Left            =   10860
      Top             =   3570
   End
   Begin VB.Timer Timer6 
      Interval        =   500
      Left            =   10440
      Top             =   3570
   End
   Begin VB.Timer Timer5 
      Interval        =   5000
      Left            =   10020
      Top             =   3570
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   10035
      Top             =   3150
   End
   Begin VB.Timer LastPressTrigTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10425
      Top             =   3135
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   10440
      Top             =   2715
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   10080
      Top             =   2730
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00716311&
      Caption         =   "Stop at the Track End Buffer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3984
      TabIndex        =   27
      Top             =   4392
      Width           =   2868
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00716311&
      Caption         =   "Stop in 15 Min"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3984
      TabIndex        =   20
      Top             =   3444
      Width           =   2868
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H007D9C34&
      Caption         =   "Winamp OnTop"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   10575
      TabIndex        =   19
      Top             =   5190
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H007D9C34&
      Caption         =   "ImageDupe"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   10575
      TabIndex        =   18
      Top             =   4965
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.ListBox List3 
      Enabled         =   0   'False
      Height          =   840
      Left            =   9435
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   1395
      Width           =   525
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Arh So "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   7752
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2295
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.ListBox List2 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7404
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   3876
      Width           =   2040
   End
   Begin RichTextLib.RichTextBox RTB1 
      CausesValidation=   0   'False
      Height          =   1416
      Left            =   7068
      TabIndex        =   14
      Top             =   6792
      Width           =   5868
      _ExtentX        =   10345
      _ExtentY        =   2487
      _Version        =   393217
      BackColor       =   16711680
      Enabled         =   0   'False
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      OLEDropMode     =   0
      TextRTF         =   $"frmCentral.frx":0E42
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox List1 
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      Height          =   2205
      Left            =   -30
      TabIndex        =   13
      Top             =   6870
      Visible         =   0   'False
      Width           =   5925
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Monitor"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8385
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6285
      Width           =   960
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   9600
      TabIndex        =   7
      Top             =   5580
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Music."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8385
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5895
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   240
      Left            =   1428
      TabIndex        =   72
      Top             =   240
      Width           =   4512
      _ExtentX        =   7964
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar3 
      Height          =   255
      Left            =   7065
      TabIndex        =   73
      Top             =   555
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label46 
      BackColor       =   &H00000000&
      Caption         =   "0:00:00"
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
      Height          =   312
      Left            =   1260
      TabIndex        =   70
      Top             =   1620
      Width           =   984
   End
   Begin VB.Label Label45 
      BackColor       =   &H00000000&
      Caption         =   "Time Now"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   0
      TabIndex        =   69
      Top             =   1620
      Width           =   1248
   End
   Begin VB.Label Label43 
      BackColor       =   &H00C00000&
      Caption         =   "_2nd Stop Time During Play"
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
      Height          =   264
      Left            =   2280
      TabIndex        =   68
      Top             =   1452
      Width           =   2916
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H009C0D05&
      Caption         =   "# 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   264
      Left            =   5208
      TabIndex        =   67
      Top             =   1176
      Width           =   504
   End
   Begin VB.Label Label37_2 
      Alignment       =   2  'Center
      BackColor       =   &H009C0D05&
      Caption         =   "# 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   264
      Left            =   5208
      TabIndex        =   66
      Top             =   1452
      Width           =   504
   End
   Begin VB.Label Label37_STARTTIME_2_GAP 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00:00:00a"
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
      Height          =   264
      Left            =   5736
      TabIndex        =   65
      Top             =   1452
      Width           =   1104
   End
   Begin VB.Label Label42 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "&& Day Before"
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
      Height          =   312
      Left            =   4824
      TabIndex        =   64
      Top             =   852
      Width           =   1212
   End
   Begin VB.Label Label41 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "of 24Hr Today "
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
      Height          =   312
      Left            =   4836
      TabIndex        =   63
      Top             =   504
      Width           =   1212
   End
   Begin VB.Label Label_RADIO_STREAMER_OR_MP3_PLAYING 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Internet Hardwire Radio Playing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   636
      Left            =   12
      TabIndex        =   62
      Top             =   3984
      Width           =   3960
   End
   Begin VB.Label Label_WINAMP_VER_ON_FORM 
      BackColor       =   &H00FFFFFF&
      Caption         =   "WINAMP V."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   12
      TabIndex        =   61
      Top             =   3708
      Width           =   3960
   End
   Begin VB.Label Label_Tune__________ 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tunes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   0
      TabIndex        =   60
      Top             =   5280
      Width           =   6852
   End
   Begin VB.Label Label_Tune_File_Name 
      BackColor       =   &H00FFFFFF&
      Caption         =   "File Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   612
      Left            =   0
      TabIndex        =   59
      Top             =   4644
      Width           =   6852
   End
   Begin VB.Label LabelX 
      Caption         =   "Label41"
      Height          =   216
      Left            =   7632
      TabIndex        =   58
      Top             =   5484
      Visible         =   0   'False
      Width           =   1848
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "OOO"
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
      Height          =   276
      Left            =   7116
      TabIndex        =   56
      Top             =   1308
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label Label39 
      BackColor       =   &H00000000&
      Caption         =   "00:00"
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
      Height          =   300
      Left            =   12
      TabIndex        =   55
      ToolTipText     =   "NOT BUTTON ACTIVE"
      Top             =   2268
      Width           =   2220
   End
   Begin VB.Label Label38 
      BackColor       =   &H00000000&
      Caption         =   "Web Cam"
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
      Height          =   300
      Left            =   12
      TabIndex        =   54
      ToolTipText     =   "SHOW IMAGE ICEVIEW"
      Top             =   1956
      Width           =   2220
   End
   Begin VB.Label Label37________________ 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "00:00:00a"
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
      Height          =   264
      Left            =   5736
      TabIndex        =   53
      Top             =   1176
      Width           =   1104
   End
   Begin VB.Label Label36 
      BackColor       =   &H00C00000&
      Caption         =   "_Start Time Orginal 1st"
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
      Height          =   276
      Left            =   2280
      TabIndex        =   52
      Top             =   1164
      Width           =   2916
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Process Set Licked And Are Still on The Runner"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   12264
      TabIndex        =   51
      Top             =   3192
      Visible         =   0   'False
      Width           =   1332
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   4764
      X2              =   4764
      Y1              =   516
      Y2              =   1146
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   2484
      X2              =   2496
      Y1              =   2832
      Y2              =   3132
   End
   Begin VB.Label Label34 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "00.0%"
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
      Height          =   312
      Left            =   6072
      TabIndex        =   50
      Top             =   528
      Width           =   768
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0.0%"
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
      Height          =   312
      Left            =   6060
      TabIndex        =   49
      Top             =   852
      Width           =   780
   End
   Begin VB.Label Label32 
      BackColor       =   &H00000000&
      Caption         =   "Play Day B4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   2280
      TabIndex        =   48
      Top             =   852
      Width           =   1524
   End
   Begin VB.Label Label31 
      BackColor       =   &H00400000&
      Caption         =   $"frmCentral.frx":0ED5
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
      Height          =   1092
      Left            =   2280
      TabIndex        =   47
      Top             =   1716
      Width           =   3444
   End
   Begin VB.Label Label30 
      BackColor       =   &H00000000&
      Caption         =   "Play Today "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   2280
      TabIndex        =   46
      Top             =   516
      Width           =   1524
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Idle Stop "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   336
      Left            =   3996
      TabIndex        =   45
      Top             =   2856
      Width           =   1572
   End
   Begin VB.Label Label28 
      BackColor       =   &H00000000&
      Caption         =   "Tott Time "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   0
      TabIndex        =   44
      Top             =   516
      Width           =   1248
   End
   Begin VB.Label Label27 
      BackColor       =   &H00000000&
      Caption         =   "Time Left "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   432
      Left            =   12
      TabIndex        =   43
      Top             =   1176
      Width           =   1236
   End
   Begin VB.Label Label26 
      BackColor       =   &H00000000&
      Caption         =   "Play Time "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   324
      Left            =   0
      TabIndex        =   42
      Top             =   840
      Width           =   1248
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "MP3 1 Before"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   41
      Top             =   276
      Width           =   1392
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "MP3 Now"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   40
      Top             =   12
      Width           =   972
   End
   Begin VB.Label Label22 
      BackColor       =   &H00000000&
      Caption         =   "10000 Hoovered"
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
      Height          =   312
      Left            =   12
      TabIndex        =   38
      Top             =   2916
      Width           =   2220
   End
   Begin VB.Label Label20 
      BackColor       =   &H00000000&
      Caption         =   "0:00:00"
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
      Height          =   324
      Left            =   1260
      TabIndex        =   36
      Top             =   1176
      Width           =   984
   End
   Begin VB.Label Label19 
      BackColor       =   &H00000000&
      Caption         =   "1000 Clicked "
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
      Height          =   288
      Left            =   12
      TabIndex        =   35
      Top             =   2592
      Width           =   2220
   End
   Begin VB.Label Label18 
      BackColor       =   &H00000000&
      Caption         =   "00:00:00a"
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
      Height          =   780
      Left            =   5748
      TabIndex        =   34
      Top             =   1740
      Width           =   1092
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000000&
      Caption         =   "00:00:00"
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
      Height          =   312
      Left            =   3828
      TabIndex        =   30
      Top             =   852
      Width           =   948
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   228
      Left            =   5952
      TabIndex        =   29
      Top             =   264
      Width           =   900
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "0:00:00"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   432
      Left            =   7080
      TabIndex        =   28
      Top             =   108
      Width           =   1440
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "---- ** Title ** ----"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   6900
      TabIndex        =   26
      Top             =   4200
      Width           =   2496
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C00000&
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   336
      Left            =   5592
      TabIndex        =   25
      Top             =   2856
      Width           =   1248
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "00:00:00"
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
      Height          =   312
      Left            =   3828
      TabIndex        =   24
      Top             =   516
      Width           =   948
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   264
      Left            =   5952
      TabIndex        =   23
      Top             =   0
      Width           =   900
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "0:00:00"
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
      Height          =   312
      Left            =   1260
      TabIndex        =   22
      Top             =   516
      Width           =   984
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "0:00:00"
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
      Height          =   324
      Left            =   1260
      TabIndex        =   21
      Top             =   840
      Width           =   984
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H007D9C34&
      Caption         =   "111"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   7875
      TabIndex        =   8
      Top             =   3090
      Width           =   1605
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Processes"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   492
      Left            =   7704
      TabIndex        =   12
      Top             =   1668
      Width           =   1020
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "10@"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   348
      Left            =   12264
      TabIndex        =   11
      Top             =   2808
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H007B45FA&
      Caption         =   "Monitor On"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   7935
      TabIndex        =   10
      Top             =   825
      Width           =   1275
   End
   Begin VB.Label Label6 
      BackColor       =   &H00404000&
      Caption         =   "Key"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   9612
      TabIndex        =   6
      Top             =   3084
      Visible         =   0   'False
      Width           =   828
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404000&
      Caption         =   "Mouse "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   276
      Left            =   9612
      TabIndex        =   5
      Top             =   2796
      Visible         =   0   'False
      Width           =   828
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H00716311&
      Caption         =   "Key"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   10452
      TabIndex        =   4
      Top             =   3084
      Visible         =   0   'False
      Width           =   1788
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00716311&
      Caption         =   "Mouse"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   264
      Left            =   10452
      TabIndex        =   3
      Top             =   2808
      Visible         =   0   'False
      Width           =   1788
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "6/6/06 6:6:6P"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   10110
      TabIndex        =   2
      Top             =   1425
      Width           =   2760
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H007B45FA&
      Caption         =   " Music On"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   7935
      TabIndex        =   1
      Top             =   1185
      Width           =   1275
   End
   Begin VB.Menu Mnu_exit 
      Caption         =   "Exit"
   End
   Begin VB.Menu Mnu_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB_FOLDER"
   End
   Begin VB.Menu MNU_ADMINISTRATOR 
      Caption         =   "ADMIN_OFF"
   End
   Begin VB.Menu Menu_Menu 
      Caption         =   "&Menu"
      Begin VB.Menu Menu_KeyboardM 
         Caption         =   "Keyboard Menu"
      End
      Begin VB.Menu Menu_KnowProcess 
         Caption         =   "Known Processes I Like"
      End
      Begin VB.Menu Mnu_CIDBreak 
         Caption         =   "CID CODE Break Resume"
      End
      Begin VB.Menu Mnu_NotePad 
         Caption         =   "Notepad AutoSave"
      End
      Begin VB.Menu Mnu_Publish 
         Caption         =   "Publisher AutoFlow"
      End
      Begin VB.Menu Mnu_Voiceonoff 
         Caption         =   "Voice Off"
      End
      Begin VB.Menu Mnu_VB 
         Caption         =   "Move VB ME"
      End
      Begin VB.Menu MNU_MOVE_EXP 
         Caption         =   "MOVE EXPLORA"
      End
      Begin VB.Menu Mnu_Dont_LowerART 
         Caption         =   "Dont Lower My Art Prog CPU"
      End
      Begin VB.Menu MNU_CPU_PROG_DONT_USE_IT 
         Caption         =   "CPU PROG DON'T USE IT STARTER WHEN TICK BOX"
         Checked         =   -1  'True
      End
      Begin VB.Menu MNU_CPU_PROG 
         Caption         =   "CPU PROG LOAD IT"
      End
      Begin VB.Menu Mnu_Process 
         Caption         =   "Process"
         Begin VB.Menu Mnu_Process_W_Win 
            Caption         =   "Processes"
         End
         Begin VB.Menu Mnu_Proc 
            Caption         =   "Show Hide Process Box"
         End
      End
   End
   Begin VB.Menu Mnu_Logs 
      Caption         =   "Loggs"
      Begin VB.Menu MNU_Logg_Folder_WEBCAM 
         Caption         =   "WEBCAM Folder "
      End
      Begin VB.Menu MNU_CLIP_SONG_PLAYING 
         Caption         =   "CLIPBOARD SONG PLAYING"
      End
      Begin VB.Menu Mnu_LoggFolder 
         Caption         =   "Logg Folder"
      End
      Begin VB.Menu MNU_AUTO_HOT_KEY_MUSIC_ON_ALL_THE_TIME 
         Caption         =   "LOADER __ AUTO HOT KEY __ TUNE CHANGE AND PLAY ALL THE TIME"
      End
      Begin VB.Menu MNU_00_MUSIC_INFO_LOGGER_TXT 
         Caption         =   "\00 MUSIC INFO LOGGER.TXT"
      End
      Begin VB.Menu MNU_NOTEPAD_01_TOTAL_LIST_TXT 
         Caption         =   "\01 TOTAL LIST.TXT"
      End
      Begin VB.Menu MNU_SPACER1 
         Caption         =   "----"
      End
      Begin VB.Menu MNU_WINAMP_FNOOB 
         Caption         =   "LOADER __ WINAMP WITHER PLAYING FNOOB"
      End
      Begin VB.Menu Mnu_Musicday_Trim_TAIL_EXE 
         Caption         =   "LOADER __ Tail-4.2.12\Tail.exe __ WITHER __  Winamp Lenght Logg Trim"
      End
      Begin VB.Menu MNU_MUSICDAY_TRIM_TAIL_IN_BROWSER_CHROME 
         Caption         =   "LOADER __ Winamp Lenght Logg Trim __ IN_BROWSER_CHROME"
      End
      Begin VB.Menu Mnu_Musicday_Trim 
         Caption         =   "LOADER __ NOTEPAD++ __ Winamp Lenght Logg Trim "
      End
      Begin VB.Menu Mnu_Musicday_Trim_NOTEPAD 
         Caption         =   "LOADER __ NOTEPAD __ Winamp Lenght Logg Trim "
      End
      Begin VB.Menu MNU_SPACER2 
         Caption         =   "----"
      End
      Begin VB.Menu Mnu_TAIL_EXE_BLUETOOTH_LOGGER 
         Caption         =   "LOADER __ Tail-4.2.12\Tail.exe __ BLUETOOTH_LOGGER"
      End
      Begin VB.Menu Mnu_WinampLenghtLogg 
         Caption         =   "WinAmp Lenght Logg -- NOT IN USE AS LOGGER UPDATE FOR LONG"
      End
      Begin VB.Menu Mnu_Musicday 
         Caption         =   "Music Day Log -- Numeric Lenght Time Logger"
      End
      Begin VB.Menu Mnu_KeyLogText 
         Caption         =   "Key Logger Text"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_KeyLoggDay 
         Caption         =   "Key Logger Day Text"
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_KeyLoggFolder 
         Caption         =   "Key Logger Text Folder"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Mnu_Mini 
      Caption         =   "Mini"
   End
   Begin VB.Menu MNU_BRing_Front 
      Caption         =   "Bring All Front"
   End
   Begin VB.Menu MNU_GET_EXPLORER_CURRENT_SESSION_CLIPBOADIN 
      Caption         =   "Get_Explorer_Current_Session_Clipboadin__Taskkill"
      Visible         =   0   'False
   End
   Begin VB.Menu Menu_Help 
      Caption         =   "&Help"
      Visible         =   0   'False
      Begin VB.Menu Menu_HelpAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu MNU_MY_NAME 
      Caption         =   "Matthew Lancaster"
   End
   Begin VB.Menu Mnu_My_Email 
      Caption         =   "Matt.Lan@btinternet.com"
   End
   Begin VB.Menu Mnu_Contact_Me_For_HardCore 
      Caption         =   "Contact Me For HardCore __ 077722224555"
   End
End
Attribute VB_Name = "CID_Run_Me"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'RESIZE FIND HERE HAPPEN
'-----------------------
'Me.Height = (y + PLUSO)
'Me.Width = (x + 100)

Dim OLD_Label11

Public WinAmp_Music_Lenght_Logg_TRIM__7_ASUS_GL522VW

Public OLD_MENU_HEIGHT

Public I_N_TAIL
Public I_N_NOTEPAD
Public I_N_AUTOHOTKEY

Dim XVB_DATE

Dim WINAMP_HAS_BEEN_FOUNDFIRST_TIME
Dim ExplorerGone, ExplorerGone_TEST
Dim WINAMP_DONT_EXIST_ANYMORE_DONE
Dim WINAMP_DONT_EXIST_ANYMORE

Dim O_WINAMP_HAS_BEEN_FOUND

Dim O_VCODE
Dim O_WINAMP_HANDLE
Dim BEFORE_StartPlayTime_2
Dim BEFORE_CurrentSongx


'WE WANT TIME LAST MP3 FINISHED SHOWN
'START PLAY TIME
'AND HOW LONG ONE BATCH WENT FOR LIKE YEAH A LONG STOP TING
'AND STOP ON END QUE WINAMP A MUST
'DONT TELL YOU CAN GET GADGET AND BRING UP TAKES 3 YEARS OF MENU

Dim TIMER8_LEFT_THE_FORM_LOAD_ENTRY_YET
Dim BEFORE_MSGRESULT_MP3_PLAYING
Dim MESSENGER_BOX_ABOUT_GetSpecialfolder

Dim TAIL_MESSENGER_BOX_DONE
Dim Fdate1 As Date
Dim Fdate2 As Date

Dim MSGBOX1
Dim FILE_SPEC_GO

Dim TAIL_EXE_BLUETOOTH_LOGGER_CASE
Dim EXE_CASE

Public CID_RUN_ME_VAR_TO_EXIT_HAPPEN

Dim OLD_SunSet

Public PATH_FILE_AT_ENTRY_SUBROUTINE_RUN_ONCE
Public PATH_FILE_AT_ENTRY_SUBROUTINE

Public MNU_AUTO_HOT_KEY_MUSIC_ON_ALL_THE_TIME_AT_FOR_LOAD_WHEN_RELOAD_ON
Public MNU_WINAMP_FNOOB_AT_FORM_LOAD

Dim START_TIME_HANDLE_HIGH_FAULT As Date

Dim CurrentSongx1_DOUBLE_BEFORE_CHECKER_1
Dim CurrentSongx1_DOUBLE_BEFORE_CHECKER_2

Dim FORM_LOAD_CURRENT_SONG
Dim TAIL_AT_FORM_LOAD_1
Dim TAIL_AT_FORM_LOAD_2

Dim CurrentSongx2_COMPARE_RADIO_SONG

'---------------------------------------------------
'PROGRAM IS IN A FAULT HANDLE COUNT CLIMB
'NOT ANYMORE SUB ROUTINE IS NOT IN OPERATION ANYMORE
'LOCATE IF YOU WANT
'---------------------------------------------------

Public OHOURNOW
'EXIT PER HOUR

Dim XET(), oShell_TrayWnd

Dim INET, INET2, INET3, INETSource2, INETSource
Dim FFT, OXMP3
Dim CurrentSongx20
Dim Last50EXEListTXT
Dim Last50EXEListAPPEND
Dim AGROCURRENTOLD
Dim AGROSTOREEX, AGROCURRENT, HWND_HOV_TOTAL
Dim AGRONOW, AGROSTORE, OutPutMP3SigNOW
Dim SaveHWNDModNOW
Dim LastMixLoad
Dim InPutDate, TestDate, Year5, MonthStart, DayCount1, DayCountMonth, WholeYear1, DayCountT
Dim Month5, WeeksSinceYear, WeeksSinceInput, ResultYearDate, Month7, WeeksPlusDays
Dim DexFORMAT1, DexFORMAT2, DexFORMAT3
Dim F1, F2, F3, F4

Dim SunSet, SunRise, AK47, Xxp2, OCheckQ, OKM, AGRO, TF2 As Long, OLTF2, OTimeWinamp
Dim oWinamp, IDate
Dim TDX As String

Dim StartPlayTime_1, StartPlayTime_2


Dim TFX, OTFX, XSG, OUREXEWINAMP, OFlag, OXxR5, OCIDLeft, LockMove, LastWeb, NoRunYet, Timer14Var
Public VMM, LastPressTrig, WebCamTime, WebCamCount
Public Xx$, TT1$, Tt2$, Gg$
Public Nx!, Ny!
Dim Rt2 As RECT
Dim WinTitleBuf As String * 255
 
Private Declare Function WindowFromPoint _
                 Lib "user32" (ByVal xPoint As Long, _
                               ByVal yPoint As Long) As Long

Public MFile As New mclsFileProperties
'VbaWindow
'lastpress
'f8 keys are done in tidal
'WMI Demo - CPU Information

'Good Sub's
'-------------
'GetFileFromHwnd (handle2)
'GetFileFromProc (CMediaSource)

'Sub KeyOrMouse()

'On Local Error Resume Next

'Dim oWinamp As New WINAMPCOMLib.Application

'A = TimeSerial(6, 0, 0)

Public OCurProcHwnd, CurProcHwnd_Count, DayNow, CurProcHwnd_Count_Yester, Txz1$
Public OCurWinHovHwnd, CurWinHovHwnd_Count, CurWinHovHwnd_Count_Yester, DayNow2

Public OldStatmp3, OlWinVis, Lock5, WinampHWVM2

Dim WinArray() As Long
Public WinAmp2, OWinAmp2, WinCap, Mini1


Public CNote, CNote3, CNoteSource, CNoteSource2, CNoteSource4 As Date, CNote2, TTT, TTT2
Public LastVice, KeyHittsOld

Public OldTF2 ' Last WinAmp Running.txt
Public OldWinampFileName, WinampTimerFileName, OldWinampHwnd

'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wmessage As Long, ByVal wParam As Long, lParam As Any) As Long
'Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Const WM_CLOSE = &H10

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWndParent As Long, _
     ByVal hWndChildAfter As Long, _
     ByVal lpszClass As String, _
     ByVal lpszTitle As String) _
    As Long

Public TiTBait, YYTT, LastCurProc5, NeroMin, DvdtoMpegMin, XZA
Public Xtop5 As Date, Timer8Flag
Public WinStopTimer, WinStopTimerH, WinStopTimerTot, CurrentSong1$, CurrentSong2$, MsgResultMp3Old
Public CurrentSongx1, CurrentSongx2$, TwicePro2, ProgressBar1v, Label3v

Public Xtop, XTOP2, TotalMusic1, TotalMusic2, TotalMusicDay

Public Xxr As Long, Xxr1, Xxr2, Xxr3, Xxr3t, Xxr4, Xxr5, Xxr6, Xxr7, Xxr8, Xxr9, Xxr10, Xxr11
Public Xxr12, Xxr13, Xxr14, Xxr15, Xxr16, Xxr17, Xxr18, Xxr19, Xxr20, Xxr21, Xxr22
Public Xxr23, Xxr24, Xxr25, Xxr26, Xxr27, Xxr28, Xxr29, Xxr30, Xxr31, Xxr32, Xxr33
Public Xxr34, Xxr35, Xxr36, Xxr37, Xxr38, Xxr39, Xxr40, Xxr41, Xxr42, Xxr43, Xxr44
Public Xxr45, Xxr46, Xxr47, Xxr48, Xxr49, Xxr50, Xxr51, Xxr52, Xxr53, Xxr54, Xxr55
Public Xxr56, Xxr57, Xxr58, Xxr59, Xxr60, Xxr61, Xxr62, Xxr63, Xxr70, Xxr71 As Long
Public Xxr72, Xxr73, Xxr74, XXR75, XXR77, XXR78

Dim XXRsCNT
Dim XXrXS(100)

Public XxrTimer
Public NotePTime, NotePad4Time
Public WRNow
Public TTx$, TTg$

Public WinAmpT, WinAmpL

Public DupeCool, DupeCool2

Public CMediaSource As Long, CMediaSource2, CMedia As Long
Public CMedia3

Public WinampHWSource As Long, WinampHWSource2, WinampHW As Long
Public WinampHW3, WinampArrayCount

Public Voice As SpVoice



Enum WindowStates
  SW_HIDE = 0
  SW_SHOWNORMAL = 1
  SW_NORMAL = 1
  SW_SHOWMINIMIZED = 2
  SW_SHOWMAXIMIZED = 3
  SW_MAXIMIZE = 3
  SW_SHOWNOACTIVATE = 4
  SW_SHOW = 5
  SW_MINIMIZE = 6
  SW_SHOWMINNOACTIVE = 7
  SW_SHOWNA = 8
  SW_RESTORE = 9
  SW_SHOWDEFAULT = 10
  SW_FORCEMINIMIZE = 11
  SW_MAX = 11
End Enum



Public OutlookPressTimer As Date

Public NeroDupeTT$
Public NotePad2RT2

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd _
        As Long, lpPoint As POINTAPI)
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function RasEnumEntriesA Lib "rasapi32.dll" _
    (ByVal Reserved As String, ByVal lpszPhonebook As String, _
    lprasentryname As Any, lpcb As Long, lpcEntries As Long) _
    As Long



Private Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpcb As Long, lpcConnections As Long) As Long





Private Const sLocation As String = "CID"

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
'SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_ON
Private Const MONITOR_ON = -1&
Private Const MONITOR_OFF = 2&
Private Const SC_MONITORPOWER = &HF170&
Private Const WM_SYSCOMMAND = &H112
Private Const ipc_isplaying As Long = 104
Private Const IPC_GETPLAYLISTFILE  As Long = 211
Private Const IPC_GETOUTPUTTIME  As Long = 105
Private Const IPC_GETINFO As Long = 126
Private Const WINAMP_BUTTON2 = 40045
Private Const WINAMP_BUTTON3 = 40046
Private Const WM_CLOSE = &H10
Private Const WM_USER = &H400
Private Const WM_COMMAND = &H111
Private Const GW_HWNDNEXT = 2

Private Type RASENTRYNAME95
    dwSize As Long
    szEntryName(256) As Byte
End Type





Public Mousey As Long
Public Keyy As Long
Public Killtime As Long

Public Sluty As Long
Public Sluty2 As Long

Public Tibulate As Double

Public Startuptime As Double
Public Postimer As Date

Dim Music_Off_Timer As Date
Dim Monitor2_Off_Timer As Date
Public Skip2 As Long

Public ScrWidth_Old As Double

Public HeightOffSet As Long
Public Ew2 As Date

Dim MsgResult As Long
Dim XcNul As Long
Dim LhRet As Long

Public Atg$
Public CurProcHwnd2 As Long
Public TameBeast2 As Long

Public Chink55 As Long

Public ListBox55

Public Qwerty2 As Long

Public Mousey3 As Long
Public Mouse55 As Long
Public KeyBack

Public TogW As Long
Public TogW2 As Long

Public Xmp1 As Long
Public Ymp1 As Long

Public Comstr As Long
Public Wq5 As Long
Public Now4 As Date
Public Bdf3 As Long

Public Xmarks As Date

Public Sps As Long
Public Spd As Long

Public FormLoad1 As Long

Public AgesEmpire As Long

Public FirstRun As Long

Public MattsTimer5 As Date

Public AgPop As Date

Public Dbon As Double
Public DbonPop As Long

Public MajorOnTime As Date
Public MinorOnTime As Date

Public Rem5 As Long

'Public NoMonOff As long
'Public NoMusic As long
Public NoMonOffX As Long
Public NoMusicX As Long

Public Dbon2 As Long

Public DayCounter As Long
Public DayCheck As Date
Public Daymouse As Long
Public Daykeyy As Long

Public Gabh As Long
Public Gabh2 As Long
Public Gabh3 As Long
Public GabhTimer As Date
Public GabhTimer2 As Date

Public DayMouseTimer As Date

Public TotalDayMouse As Long
Public TotalDayKeyy As Long

Public AsusTimer20 As Date

Public MattsTimer5Old As Date
Public MattsTimer2Old As Date
Public PeaShoot As Double
Public XRad As Long
Public OldXRad As Long
Public OldWState As Long

Public WLFN1$
Public WLFN2$

Public Timer5Div As Date

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


Public FirstTime5Run As Long

Public LastCurProcHwnd As Long


Public FirstRun2 As Long


Public NoMonMusChange As Long

Public PlimBurn1 As Long
Public PlimBurn2 As Long

Public tagad2perm$
Public wedf$
Public wedg$
Public Chit22 As Long
Public Chit24 As Long

Public ReRun As Long

Public LockInToMp3 As Long
'-----------------------------------------------------
'-----------------------------------------------------
'-----------------------------------------------------

Public QueTheAmpTest As Date

Public AutoConVpn As Date

Public OldWindow_Title_String
Public OldAshPI As Integer

Public Agust As Date ' timer for sub bashing for winamp stopper if diff album
Public Agust2 As Date

Public XtX1
Public XtX2

Public Label10BackColor
Public Label21BackColor
Public Label23BackColor
Public InitHeight
Public EmSQ

Public KKQ

Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long  'MODULE 1141

'---------------------------------------------------------------
'http://www.vbforums.com/showthread.php?673134-RESOLVED-Minimum-height-for-menu-bar-to-be-visible
'-------------- MENU SIZE DECLARE
'Private Type rect
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
Private Type MENUBARINFO
  cbSize As Long
  rcBar As RECT
  hMenu As Long
  hwndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type
Private MenuInfo As MENUBARINFO
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function FindWindowDLL Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long




Public Sub Init()


NoMonOff2 = 0
NoMonOff = 0
NoMusic = 0
ProcessId24 = 0
ProcessId22 = 0
frontpagepid = 0
ProcessId25Ssam = 0

ProcLB.List1.Clear
mdlProcess.List_ActiveProcess
mdlProcess.List_ActiveModules

Call OutputProcFileList

End Sub


Sub CheckSave()
    Call zzSave_Checks
    Exit Sub
    
    'CheckSave = Cheese Cake
    Call FileInUseDelay(App.Path + "\Switches.txt", 1)
    GUG = FreeFile
    Open App.Path + "\Switches.txt" For Output As #GUG
    For Each Control In Controls
        If InStr(Control.Name, "Check") > 0 Then
            Print #GUG, Control.Value
        End If
    Next
    Close #GUG
End Sub


Private Sub Check1_Click()
    DupeCool = 0
    DupeCool2 = 0
    Call CheckSave
End Sub

Private Sub Check2_Click()

If Check2.Value = vbChecked Then
    win2 = FindWindow("Winamp v1.x", vbNullString)
    win3 = FindWindow("Winamp PE", vbNullString)
    rt = AlwaysOnTop(win2)
'    rt = AlwaysOnTop(win3)
    ShowWindow win2, SW_NORMAL
    'ShowWindow win2, SW_SHOW
    'ShowWindow win3, SW_NORMAL
    rt = AlwaysOnTop(win2)
Else
    win2 = FindWindow("Winamp v1.x", vbNullString)
    win3 = FindWindow("Winamp PE", vbNullString)
    rt = NotAlwaysOnTop(win2)
'    rt = NotAlwaysOnTop(win3)
    ShowWindow win2, SW_MINIMIZE
    'ShowWindow win2, SW_SHOWNOACTIVATE
    'ShowWindow win3, SW_SHOWNOACTIVATE
    'ShowWindow win2, SW_NORMAL

End If
Call CheckSave
End Sub

Private Sub Check3_Click()
If XZA = True Then Exit Sub
XZA = True
'Check3.Value = vbUnchecked
Rth2 = Check3.Value = vbChecked And Check5.Value = vbUnchecked And Check6.Value = vbUnchecked And Check8.Value = vbUnchecked
rt = Check5.Value = vbChecked Or Check6.Value = vbChecked Or Check8.Value = vbChecked
Check5.Value = vbUnchecked
Check6.Value = vbUnchecked
Check8.Value = vbUnchecked
DoEvents
XZA = True

If rt = True Or Rth2 = True Then
Check3.Value = vbChecked
WinStopTimerTot = TimeSerial(0, 15, 0)
'WinStopTimer = Now + WinStopTimerTot
DoEvents
Call KeyOrMouse

End If

XZA = False
Call CheckSave
End Sub

Private Sub Check4_Click()
'----
'Buffer stop - Wikipedia
'https://en.wikipedia.org/wiki/Buffer_stop
'----
Call CheckSave

End Sub

Private Sub Check5_Click()
If XZA = True Then Exit Sub
XZA = True
Rth2 = Check5.Value = vbChecked And Check3.Value = vbUnchecked And Check6.Value = vbUnchecked And Check8.Value = vbUnchecked
rt = Check3.Value = vbChecked Or Check6.Value = vbChecked Or Check8.Value = vbChecked
Check3.Value = vbUnchecked
Check6.Value = vbUnchecked
Check8.Value = vbUnchecked
DoEvents

If rt = True Or Rth2 = True Then
XZA = True
Check5.Value = vbChecked
WinStopTimerTot = TimeSerial(0, 0, 10)
DoEvents
'WinStopTimer = Now + WinStopTimerTot
Call KeyOrMouse
End If
XZA = False
Call CheckSave

End Sub

Private Sub Check6_Click()
If XZA = True Then Exit Sub
XZA = True
Rth2 = Check3.Value = vbUnchecked And Check5.Value = vbUnchecked And Check6.Value = vbChecked And Check8.Value = vbUnchecked
rt = Check3.Value = vbChecked Or Check5.Value = vbChecked Or Check8.Value = vbChecked
Check3.Value = vbUnchecked
Check5.Value = vbUnchecked
Check8.Value = vbUnchecked
DoEvents
If rt = True Or Rth2 = True Then
XZA = True
Check6.Value = vbChecked
DoEvents

WinStopTimerTot = TimeSerial(1, 0, 0)
'WinStopTimer = Now + WinStopTimerTot
Call KeyOrMouse

End If

XZA = False
Call CheckSave
End Sub

Private Sub Check7_Click()

Call CheckSave

End Sub

Private Sub Check8_Click()
If XZA = True Then Exit Sub
XZA = True
Rth2 = Check3.Value = vbUnchecked And Check5.Value = vbUnchecked And Check6.Value = vbUnchecked And Check8.Value = vbChecked
rt = Check3.Value = vbChecked Or Check5.Value = vbChecked Or Check6.Value = vbChecked
Check3.Value = vbUnchecked
Check5.Value = vbUnchecked
Check6.Value = vbUnchecked
DoEvents
If rt = True Or Rth2 = True Then
XZA = True
Check8.Value = vbChecked
DoEvents

WinStopTimerTot = TimeSerial(3, 0, 0)
'WinStopTimer = Now + WinStopTimerTot
Call KeyOrMouse

End If

XZA = False
Call CheckSave
End Sub


Private Sub Command2_Click()

If Command2.BackColor = QBColor(12) Then
    Command2.BackColor = QBColor(10)
    Menu_StopCurrent.Checked = True
    Exit Sub
End If

If Command2.BackColor = QBColor(10) Then
    Command2.BackColor = QBColor(12)
    Menu_StopCurrent.Checked = False
End If

End Sub





Private Sub Label11_Change()

    tt = Format$(Now, "HH:MM:SSa/p")
    tt = tt + " - " + Label11
    
    XMP3 = FindWindow(vbNullString, "MP3 PLAYED")
    If XMP3 > 0 And OXMP3 <> XMP3 Then Exit Sub
    OXMP3 = XMP3
    
    'Shell "D:\VB6\VB-NT\00_Best_VB_01\MP3 PLAYED\MP3 PLAYED.exe", vbNormalFocus
    
    'MP3_PLAYED.List1.AddItem tt
    
End Sub

Private Sub Label13_Change()
    
    If Len(Label13) = 8 And Mid$(Label13, 1, 1) = "0" Then Label13 = Mid$(Label13, 2)



    'Beep

End Sub

Private Sub Label13_Click()
'Label13
End Sub

Private Sub Label17_Click()
'Label17
End Sub

Private Sub Label18_Click()
'Label18
End Sub

Private Sub Label19_Click()
'Label19
End Sub

Private Sub Label20_Click()
'Label20
End Sub

Private Sub Label21_Change()
    Call Mouse_or_Key
End Sub

Sub Mouse_or_Key()
    Idle_Timer_Proc = Now + TimeSerial(0, 30, 0)
End Sub

Private Sub Label22_Click()
'Label22
End Sub

Private Sub Label23_Change()
    'Label23
    Call Mouse_or_Key
End Sub

Private Sub Label23_Click()
'Label23
End Sub

Private Sub Label28_Click()
Call MNU_CLIP_SONG_PLAYING_Click
End Sub

Private Sub Label32_Click()
'Label32
End Sub

Private Sub Label33_Click()
'Label33
End Sub

Private Sub Label35_Click()
'Unload Me
'
'End

'Process's Licked That Are Still Running

End Sub

Private Sub Label37_Click()
'Label37________________
'Label37________________ = Format(StartPlayTime, "H:MM:SSa/p")
'Label37_STARTTIME_2_GAP = Format(StartPlayTime, "H:MM:SSa/p")

End Sub

Private Sub Label38_Click()
'Label38

'Shell "C:\Program Files\IceView\iceview.exe ""D:\VB6\VB-NT\00_Best_VB_04\Webcam_Motion\VBDataNoArchive\Web_Cam Pic Always.jpg"""

End Sub

Private Sub Label39_Click()
'Label39
End Sub

Private Sub Label40_Click()
'Label40
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Call Label6Hit

End Sub

Sub Label6Hit()

If OptionsMenu.Top <> 0 Then
    OptionsMenu.WindowState = 0
    OptionsMenu.Refresh
End If

If OptionsMenu.WindowState <> 0 Then
    OptionsMenu.Width = CID_Run_Me.Width
    OptionsMenu.Top = CID_Run_Me.Top - OptionsMenu.Height
    OptionsMenu.Left = ScreenWidthX - OptionsMenu.Width - OffSetGoogle
End If

If OptionsMenu.Visible = False Then
    OptionsMenu.Show
    OptionsMenu.List1.SetFocus
    Exit Sub
End If

If OptionsMenu.Visible = True Then
    OptionsMenu.Visible = False
End If

End Sub

Private Sub Command1_Click()
If Music_Off_Timer > 0 Then Amon = 1
If Amon = 0 Then
    Music_Off_Timer = 0
    winamp_player_off
    Music_Off_Timer = Now + TimeSerial(12, 0, 0)
End If
If Amon = 1 Then Music_Off_Timer = 0: Amon = 0

End Sub

Private Sub Command3_Click()
If Monitor2_Off_Timer > 0 Then Amon = 1
Monitor2_Off_Timer = Now + TimeSerial(12, 0, 0)
If Amon = 1 Then Monitor2_Off_Timer = 0: Amon = 0

End Sub


Sub FileManipWinamp()

Exit Sub
'RUN ON START UP FOR MODDING LIST CLEAN
        
        
LF1 = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day Total.txt" For Input As #LF1
LF2 = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day Total2.txt" For Output As #LF2
Do
    Line Input #LF1, ll$
    oll = ll
    ex = 0
    If Trim(ll$) = "" Then ex = 1
    If InStr(ll$, "WinAmp Status") > 0 Then
        ex = 1
    End If
    If Len(ll$) <= Len("01-06-2008 17:08:33 Winamp Lenght = 04:37 Song Nam  ") Then
        'If InStr(LL$, "Song Name") Then Stop
        ex = 1
    End If
    
    RR = InStrRev(ll, " Winamp Lenght ")
    If RR > 30 Then
        ll = Mid$(ll$, RR - 22)
    End If
    
    
    RR = InStr(ll$, " --")
    If RR <> 20 Then
        ll$ = Mid$(ll$, 1, 19) + " -- " + Mid$(ll$, 21)
    End If
    
    RR = InStr(ll, "-- Song")
    RR1 = InStr(ll, "Song")
    If RR = 0 And RR1 > 0 Then
        ll$ = Mid$(ll$, 1, RR1 - 1) + "-- " + Mid$(ll$, RR1)
    End If
    
    
    RR = InStr(ll, "Winamp Lenght =")
    If RR > 0 Then
        ll$ = Mid$(ll$, 1, 36) + " Played = " + Mid$(ll$, 40)
    End If
    
    RR = InStr(ll, "Play=")
    If RR > 0 Then
        ll$ = Mid$(ll$, 1, RR - 1) + "Played = " + Mid$(ll$, 44)
        'LL$ = Mid$(LL$, 1, 50) + "Played = " + Mid$(LL$, 41)
    End If
    
    'Winamp Lenght =
    
    
    RR = InStr(ll$, "Song")
    
    If RR = 56 Then
    ll$ = Mid$(ll$, 1, 46) + "00:" + Mid$(ll$, 47)
    End If
    
    If RR = 58 Then
        ll$ = Mid$(ll$, 1, 46) + "0" + Mid$(ll$, 47)
    End If
    
    
    'Do
    RR = InStr(ll, "  -- Song")
    RR1 = InStr(ll, "Song")
    If RR > 0 Then
        ll$ = Mid$(ll$, 1, RR) + Mid$(ll$, RR + 1)
    End If
    'Loop Until rr = 0
    
    
    RR = InStr(ll$, "Song")
    If InStr(ll, "=  -- Song") > 0 Then ex = 1
    
    If RR <> 59 And ex = 0 Then
    Stop
    'oll
    End If
    
    
    If ex = 0 Then Print #LF2, ll$
Loop Until EOF(LF1)
Close #LF1, #LF2

End Sub




Public Sub Form_Load()

If App.PrevInstance = True Then End

WinAmp_Music_Lenght_Logg_TRIM__7_ASUS_GL522VW = "https://drive.google.com/open?id=1zPX4SMaEF7LOe05Ag8jNv5KVxvCfJVpk"

Dim i As Long
Dim R As Long

'i = IsInternetConnected
i = IsInternetConnected_Mute()

TAIL_AT_FORM_LOAD_1 = True
TAIL_AT_FORM_LOAD_2 = True

PATH_FILE_AT_ENTRY_SUBROUTINE_RUN_ONCE = True
START_TIME_HANDLE_HIGH_FAULT = 0
Reload_Me = False

'CurrentSongx1_DOUBLE_BEFORE_CHECKER = ""
'FORM_LOAD_CURRENT_SONG = ""
'TAIL_AT_FORM_LOAD = ""
'CurrentSongx2_COMPARE_RADIO_SONG = ""

ChDrive App.Path
ChDir App.Path

ReDim NoteHard1(0)
ReDim NoteHard2(0)

Dim Rect8 As RECT
Dim MeRyu3  As RECT
Dim MeRyu4  As RECT
Dim MeRyu5  As RECT

Label10BackColor = Label10.BackColor
Label21BackColor = Label21.BackColor
Label23BackColor = Label23.BackColor

Label11.Caption = "---- ** Title ** ----"

LastWeb = Now

Dim CTCOL()
ReDim CTCOL(Me.Controls.Count)
'-------------------------------------------
'TEMP PUT OUT
'-------------------------------------------

For Each Control In Controls
    If InStr(LCase(Control.Name), "timer") > 0 Then
        If Control.Enabled = True Then
            i = i + 1
            CTCOL(i) = Control.Name
            Control.Enabled = False
        End If
    End If
Next
ReDim Preserve CTCOL(i)

Set FS = CreateObject("Scripting.FileSystemObject")

If 1 = 2 Then
    On Local Error Resume Next
    Set dc = FS.Drives
    For Each d In dc
        dr = d.DriveLetter
        n = d.VolumeName
        If InStr(n, "RAM") > 0 Then Exit For
    Next
    On Local Error GoTo 0
    RamDrive = dr
End If
RamDrive = ""

Call LoadHwndCounts

'ef = MFile.FindFileInfo("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE", False)
'ef = MFile.LastWriteTime

'Exit Sub

If IsIDE = False Then
'    Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /r ""D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\Cid-RunMe.vbp""", vbMinimizedNoFocus
'    End
End If

Dim WinampArray(30)
Dim WinampArray2(30)
'oWinamp.AutoloadEnabled = False
    
ReDim WinArray(20)

O = FindHighestHandle



'Xxr = 22432
'TTT = cProcesses.Convert(Xxr, Otss, cnFromhWnd Or cnToProcessID)
'Process_Kill (Otss)

KeyLoggDate = Now
    
Dim Fr1

'Fr1 = FreeFile
'Open App.Path + "\Text_Data_KeyLogg\Key Logger Text.txt" For Append As #Fr1
'Print #Fr1,
'Print #Fr1, "----- Start -----"
'Print #Fr1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS")
'Close #Fr1

'Fr1 = FreeFile
'Open App.Path + "\Text_Data_KeyLogg\Key Logger Text " + Format$(KeyLoggDate, "YYYY-mm-dd HH-MM-SS") + ".txt" For Append As #Fr1
'Print #Fr1,
'Print #Fr1, "----- Start -----"
'Print #Fr1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS")
'Close #Fr1

WinStopTimerTot = TimeSerial(0, 15, 0)
WinStopTimer = Now + WinStopTimerTot

TF2 = FindWindow("Winamp v1.x", vbNullString)
MsgResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)

If MsgResult = 0 Or MsgResult = 3 Then
    f8 = DateDiff("s", Now, WinStopTimer) Mod 60
    f7 = Int(DateDiff("s", Now, WinStopTimer) / 60)
    WinStopTimerH = TimeSerial(0, f7, f8)
End If

MsgResultMp3Old = -1

'If IsIDE = False Then Load DSkeybd_F
'If IsIDE = False Then DSkeybd_F.Show

'Load DSkeybd_F
'show if  you want to show
'DSkeybd_F.Show
'   Initialize the voice object

Set Voice = New SpVoice
On Error Resume Next
Set Voice.Voice = Voice.GetVoices().Item(3)
If IsIDE() = True Then
    Set Voice.Voice = Voice.GetVoices().Item(2)
End If
On Error GoTo 0

'5=man
'4=lady

'swap 2 & 3 if when compiled & 4 & 5

'   Set Voice.Voice = Voice.GetVoices().Item(?) 'sample TTS
'   Set Voice.Voice = Voice.GetVoices().Item(3) 'microsoft Sam
'   Set Voice.Voice = Voice.GetVoices().Item(2) 'microsoft Mike
'   Set Voice.Voice = Voice.GetVoices().Item(0) 'microsoft mary

'Voice.Status.RunningState

'---------------------------------------------------------------------------
'2017
'---------------------------------------------------------------------------
'THERE IS A COMMAND OF GETSPECIALFOLDER
'WHICH IS A ADMIN ONLY MODE COMMAND TO RETURN A RESULT
'THAT TYPE THING MAKE CRASH WHEN VB-IDE IS IN ADMIN BUT EXE COMPLILER IS NOT
'---------------------------------------------------------------------------

'----------------------------------------------------------------
Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)

If CODER_EXE_FILE_NAME_1 <> "" Then
    Call Date_File(CODER_EXE_FILE_NAME_1, Fdate1)
    '------------------------------------------------------------
    Call Date_File(CODER_VBP_FILE_NAME_2, Fdate2)
    '------------------------------------------------------------
End If

If Dir("D:\Wave's\Cosmi\CANNON.WAV") <> "" Then
    MMControl1.Notify = True
    MMControl1.Wait = True
    MMControl1.Shareable = False
    MMControl1.DeviceType = "WaveAudio"
    'MMControl1.fileName = App.Path + "\005.Wav"
    MMControl1.fileName = "D:\Wave's\Cosmi\CANNON.WAV"
    MMControl1.Command = "Open"
    'MMControl1.Command = "Prev"
    'MMControl1.Command = "Play"
End If

Idle_Timer = Now

Monitor2_Off_Timer = Now + TimeSerial(24, 0, 0)
Music_Off_Timer = Now + TimeSerial(24, 0, 0)

Agust = Now

AsusTimer20 = Now + TimeSerial(0, 0, 30)

MattsTimer5 = 0

Load OptionsMenu

MajorOnTime = TimeSerial(0, 5, 0)
MinorOnTime = TimeSerial(0, 2, 0)

FirstRun = True
        
FormLoad1 = 1

ReDim HwndArray(200)
ReDim HwndArra2(200)
ReDim GetWinLen(200)

For R = 0 To 200
    HwndArray(R) = -2
    HwndArra2(R) = -2
Next

Path_Folder = App.Path + "\0TextData\" + GetComputerName

If (Not DirExist(Path_Folder)) Then
    MkDirNested Path_Folder
End If

'-------------------
'NOT RUNNER EXIT SUB
'-------------------
DUN_Services sArray

If Dir$(App.Path + "\0TextData\" + GetComputerName + "\AutoVPN.txt") <> "" Then
FF1 = FreeFile
    Open App.Path + "\0TextData\" + GetComputerName + "\AutoVPN.txt" For Input As #FF1
    Line Input #FF1, egg$
    Close #FF1
    AutoConVpn = DateValue(egg$) + TimeValue(egg$)
End If

CID_Run_Me.Caption = "CID Run Me."
'CID_Run_Me.Caption = "CID Run Me - Ver " + Trim$(str$(App.Major)) + "." + Trim$(str$(App.Minor)) + "." + Trim$(str$(App.Revision))

Startuptime = Now
Postimer = Now + TimeSerial(0, 1, 30)
CID_Run_Me.WindowState = 0
Tibulate = Now + TimeSerial(0, 1, 0)

HeightOffSet = 500

yyu = 42

'Code to Auto Size Form based on controls used Including Menu Bar Sketchy Style
x = 1
y = 1
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True And Control.Visible = True Then
        If Control.Width + Control.Left > x Then x = Control.Width + Control.Left
        If Control.Height + Control.Top > y Then y = Control.Height + Control.Top
        If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
    End If
Next

'-----------------------------------------------------------
'FORM NOT RESIZE WHEN MINIMISED OR MAXIMSED
'HAPPEN WHEN FORM RELOAD OPERATION
'BUT COULD SET AT RELOAD OR ANY LOAD THAT IS AT STATE NORMAL
'-----------------------------------------------------------
On Error Resume Next
'On Error GoTo 0


If mnu = 1 Then
    PLUSO = 720: PLUSO = 700  'Sometimes Different
Else
    PLUSO = 450
End If

PLUSO = 1000 '  2 MENU LINES

Me.Height = (y + PLUSO)
'Me.Width = (x + 120)
'----------------------------------
'
'ME.Width --- HAPPEN HERE SEARCH_ER
'Me.Width = VAR_WIDTH_OBJECT + 100
'----------------------------------
On Error GoTo 0

'Me.Left = (Screen.Width - Me.Width)
'Me.Top = (Screen.Height - Me.Height)

Line1.x1 = Label4.Left + Label4.Width
Line1.x2 = Line1.x1
Line1.y1 = Label4.Top
Line1.y2 = (Label22.Top + Label22.Height) - 8 '- 20 ' + 50

Line2.x1 = Label33.Left + 40
Line2.x2 = Line2.x1
Line2.y1 = Label34.Top
Line2.y2 = (Label33.Top + Label33.Height) + 50

'Me.Show

'Exit Sub

'MsgBox (str$(ScreenHeightY) + vbCr + str$(CID_Run_Me.Height))

CID_Run_Me.WindowState = 0

CID_Run_Me.Refresh

OptionsMenu.Width = CID_Run_Me.Width
OptionsMenu.Top = CID_Run_Me.Top - OptionsMenu.Height
OptionsMenu.Left = ScreenWidthX - OptionsMenu.Width - OffSetGoogle

If Dir$(App.Path + "\0TextData\" + GetComputerName + "\winamp-list.txt") <> "" Then
    freef5 = FreeFile
    Open App.Path + "\0TextData\" + GetComputerName + "\winamp-list.txt" For Input As #freef5
    Do
        Line Input #freef5, Atg$
    Loop Until EOF(freef5)
    Close #freef5
    Atg$ = Mid$(Atg$, 24)
End If

QueTheAmpTest = Now + TimeSerial(0, 0, 20)

'---------------------------------------------------------------------
'MAKE PROCESS LIST SCRIPT AND SEND OUT TO A FILE
'ALSO AVAILABLE FROMA MENU OPTION TO CLICK
'AS FAR AS I LEARN IT ONLY RUN AT THE FORM LOAD
'Call Init
'---------------------------------------------------------
'MUST BE A TIME DELAY WITH THIS ONE AT STARTER
'PROCESS INTESIVE FOR A PROCESS
'Wed 11 January 2017 23:30:42----------
'SOMETHING IS USING A LOT OF HANDLE HERE PROJECT HIGHER NORTON TALK IT
'AND AUTO EXIT ITSELF AFTER TIME AND ANOTHER RERUN LOADER
'THIS REM OUT MAY OF SOLVED IT
'DOUBT IT
'---------------------------------------------------------------------

Load frmReceiver
Load frmSender_CallID

If vbCompiled2 = True Then
    Load DSkeybd_F
    MajorOnTime = TimeSerial(0, 5, 0)
    MinorOnTime = TimeSerial(0, 2, 0)
Else
    'Timer4.Enabled = True
    MajorOnTime = TimeSerial(0, 1, 0)
    MinorOnTime = TimeSerial(0, 0, 20)
End If

'---------------------------
'---------------------------
'---------------------------
'---------------------------
Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\2cidrunme.txt"

If Dir(Path_And_FileName) <> "" Then
    free5 = FreeFile
    Open Path_And_FileName For Input As #free5
    Do
        Line Input #free5, aed$
        If InStr(aed$, "Last System ReBoot Time") Then
            wett = 1
            Qwerty2 = 0
            Mouse55 = 0
        End If
        
        If wett = 1 Then
            If InStr(aed$, "Keyboard =") Then
                Qwerty2 = Qwerty2 + Val(Mid$(aed$, InStr(aed$, "Keyboard =") + 10))
            End If
            If InStr(aed$, "Mouse =") Then
                Mouse55 = Mouse55 + Val(Mid$(aed$, InStr(aed$, "Mouse =") + 7, 11))
            End If
        End If
    Loop Until EOF(free5)
    Close free5
End If

If Dir$(App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCount.txt") <> "" Then
    free5 = FreeFile
    Open App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCount.txt" For Input As #free5
    Do
        If EOF(free5) Then Exit Do
        Line Input #free5, aed$
        If InStr(aed$, "Day Counter") And InStr(aed$, "End Of Day Counter") = 0 Then
            DayCheck = DateValue(Mid$(aed$, 15))
            Line Input #free5, aed$
            Daymouse = Val(Mid$(aed$, 15))
            Line Input #free5, aed$
            Daykeyy = Val(Mid$(aed$, 15))
        End If
    Loop Until EOF(free5)
End If

If DayCheck = 0 Then DayCheck = Now

Close #free5

If Dir$(App.Path + "\0TextData\" + GetComputerName + "\WinAmpLog\Winamp-list-CallerID.txt") <> "" Then
    freef5 = FreeFile
    Open App.Path + "\0TextData\" + GetComputerName + "\WinAmpLog\Winamp-list-CallerID.txt" For Input As #freef5
    Do
        Line Input #freef5, Atg$
    Loop Until EOF(freef5)
    Close #freef5
    Atg$ = Trim(Mid$(Atg$, 23))
End If

'frmSender_CallID.txtMsg.Text = "KeyCounter " + str$(Qwerty2 + Keyy)
'Call frmSender_CallID.cmdSend_Click

'frmSender_CallID.txtMsg.Text = "MouseCounter " + str$(Mouse55 + Mousey)
'Call frmSender_CallID.cmdSend_Click

FirstRun = False

Call proc24

frmSender_CallID.txtMsg.Text = "Request Play " + Time$
'Call frmSender_CallID.cmdSend_Click

If CheckT1$ = "2" Then
    Call ClingDing
End If

On Error Resume Next
GUG = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\2Total Music.txt" For Input As #GUG
    'R = Err.Number
    Line Input #GUG, Totalm1$
    Line Input #GUG, Totalm2$
    Line Input #GUG, TotalmD$
Close GUG
On Error GoTo 0

TotalMusic1 = Val(Totalm1$)
TotalMusic2 = Val(Totalm2$)
TotalMusicDay = Val(TotalmD$)

FirstRun2 = False

'Call Mnu_Cpu_Click

'---------------------------------
'CPU CONSUMPTION HERE
'---------------------------------
'CODE SUB ROUTINE HAPPEN IN TIMER1
'NOW TEMPORALLY REMOVED
'---------------------------------
If 1 = 2 Then
    Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\Last50EXEList.txt"
    If Dir(Path_And_FileName) <> "" Then
        hnd = FreeFile
        Open Path_And_FileName For Input As hnd
        Do
            If Not EOF(hnd) Then
                Line Input #hnd, ff$
                List4.AddItem ff$
            End If
        Loop Until EOF(hnd)
        Close #hnd
    End If
End If

Dim LF1, LF2

On Error Resume Next
LF1 = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Play Pause Log Total.txt" For Append As #LF1
PATH_FILENAME = App.Path + "\0TextData\" + GetComputerName + "\WinAmp Play Pause Log.txt"
If Dir(PATH_FILENAME) <> "" Then
    LF2 = FreeFile
    Open PATH_FILENAME For Input As #LF2
    Do
        Line Input #LF2, ll$
        If Trim(ii$) <> "" Then Print #LF1, ll$
        '----------------------------------------------------
    '    DoEvents
    '    If CID_RUN_ME_VAR_TO_EXIT_HAPPEN = True Then Unload CID_Run_Me:EXIT SUB
    Loop Until EOF(LF2)
    Close #LF1, #LF2
    Kill PATH_FILENAME
End If

'If FS.FileExists(App.Path + "\0TextData\"+Computername+"\WinAmp Music Lenght Logg.txt") = True Then
'    LF1 = FreeFile
'    Open App.Path + "\0TextData\"+Computername+"\WinAmp Music Lenght Logg Total.txt" For Append As #LF1
'    LF2 = FreeFile
'    Open App.Path + "\0TextData\"+Computername+"\WinAmp Music Lenght Logg.txt" For Input As #LF2
'    Do
'        Line Input #LF2, ll$
'        If Trim(ii$) <> "" Then Print #LF1, ll$
'    Loop Until EOF(LF2)
'    Close #LF1, #LF2
'    Kill App.Path + "\0TextData\"+Computername+"\WinAmp Music Lenght Logg.txt"
'End If

'Call FileManipWinamp
        
'If FS.fielexists(App.Path + "\0TextData\"+Computername+"\WinAmp Music Lenght Logg Day.txt") = True Then
'LF1 = FreeFile
'Open App.Path + "\0TextData\"+Computername+"\WinAmp Music Lenght Logg Day Total.txt" For Append As #LF1
'LF2 = FreeFile
'Open App.Path + "\0TextData\"+Computername+"\WinAmp Music Lenght Logg Day.txt" For Input As #LF2
'Do
'Line Input #LF2, TT$
'If Trim(TT$) <> "" Then Print #LF1, TT$
'Loop Until EOF(LF2)
'Close #LF1, #LF2
'Kill App.Path + "\0TextData\"+Computername+"\WinAmp Music Lenght Logg Day.txt"
'End If
        
LF2 = FreeFile
'Open App.Path + "\0TextData\"+Computername+"\WinAmp Music Lenght Logg Day Total.txt" For Binary As #LF2
Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg-TRIM__" + GetComputerName + ".txt" For Binary As #LF2
If LOF(LF2) > 800! * 1024! Then
    PIC1$ = Space$((800! * 1024!) + 1)
        Seek LF2, LOF(LF2) - (800! * 1024!)
        Get #LF2, , PIC1$
    Close #LF2

    ii = InStr(PIC1$, vbCrLf)
    PIC1$ = Mid$(PIC1$, ii + 2)
    
    '----------------------------------------------------
    DoEvents
    If CID_RUN_ME_VAR_TO_EXIT_HAPPEN = True Then Unload CID_Run_Me: Exit Sub
    
    
    Kill App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg-TRIM__" + GetComputerName + ".txt"
    LF2 = FreeFile
    Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg-TRIM__" + GetComputerName + ".txt" For Binary As #LF2
        Put #LF2, , PIC1$
    Close #LF2
    PIC1$ = ""
End If
Close #LF2
        
        
        
'----------------------------------------------------
DoEvents
If CID_RUN_ME_VAR_TO_EXIT_HAPPEN = True Then Unload CID_Run_Me: Exit Sub
        
LF1 = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day - Sig.txt.Tmp" For Append As #LF1
LF2 = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day - Sig.txt" For Input As #LF2
ILEN_VAR = LOF(LF2) - 50000
If ILEN_VAR < 1 Then ILEN_VAR = 1
Seek LF2, ILEN_VAR
Do
    Line Input #LF2, ll$
    If Trim(ii$) <> "" Then Print #LF1, ll$
    '----------------------------------------------------
    DoEvents
    If CID_RUN_ME_VAR_TO_EXIT_HAPPEN = True Then Unload CID_Run_Me: Exit Sub
Loop Until EOF(LF2)
Close #LF1, #LF2

'----------------------------------------------------
If CID_RUN_ME_VAR_TO_EXIT_HAPPEN = True Then Unload CID_Run_Me: Exit Sub

Kill App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day - Sig.txt"
Name App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day - Sig.txt.tmp" As App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day - Sig.txt"
        
        
'CurrentSongx2$ = CurrentSongx1
Fr1 = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Last One.txt" For Input As #Fr1
    Line Input #Fr1, CurrentSongx2$
Close #Fr1

I_22 = Trim(Replace(CurrentSongx2$, "[Stopped]", ""))
CurrentSongx2$ = I_22
        
        
        
FORM_LOAD_CURRENT_SONG = True
Xx = InStr(CurrentSongx2$, "Tune_")
'Xx = InStr(CurrentSongx2$, "Song Name")
'CurrentSongx2$ = Mid(CurrentSongx2$, Xx + 12) ' ERROR
CurrentSongx2$ = Mid(CurrentSongx2$, Xx + 6)

I_22 = Trim(Replace(CurrentSongx2$, "[Stopped]", ""))
CurrentSongx2$ = I_22
        
        
        
ScreenTwipsX = Screen.TwipsPerPixelX
ScreenTwipsY = Screen.TwipsPerPixelY
ScreenWidthX = Screen.Width
ScreenHeightY = Screen.Height
CID_Run_Me.BackColor = 0
'Me.Refresh



'XZA = True
'gug = FreeFile
'Open App.Path + "\Switches.txt" For Input As #gug
'For Each Control In Controls
'    If InStr(Control.Name, "Check") > 0 Then
'        Line Input #gug, cv
'        Control.Value = Val(cv)
'    End If
'Next
'Close #gug
'XZA = False


'----------------------------------------------------
DoEvents
If CID_RUN_ME_VAR_TO_EXIT_HAPPEN = True Then Unload CID_Run_Me: Exit Sub

'----------------------------------------------
'2017 NOT SURE WHY AT BOOT IN IS DELAY SOMETIME
'----------------------------------------------
'Call Timer8_Timer
'----------------------------------------------
'BETTER METHOD IF HAS TO RUN FIRST WITHINER
'1ST TIMER TO RUN LOAD WANTER
'----------------------------------------------

'----------------------------------------------
'Call OutPutMP3Sig

Me.Show
DoEvents

'Wat a Job Now Re-Enable any timers that were enabled on start SWEET Safe
For Each Control In Controls
    If InStr(LCase(Control.Name), "timer") > 0 Then
        For i = 1 To UBound(CTCOL)
            If CTCOL(i) = Control.Name Then Control.Enabled = True
        Next
    End If
Next

'----------------------------------------------------
DoEvents
If CID_RUN_ME_VAR_TO_EXIT_HAPPEN = True Then Unload CID_Run_Me: Exit Sub

Call zzLoad_Checks

'------------------------------------------
'2017 ERROR NOT FIND THIS SUB
'------------------------------------------
'MADE ERROR AT EXE LOAD ONLY NOT IDE EFFORT
'------------------------------------------
'Load MP3_PLAYED
'------------------------------------------

Me.Show
If IsIDE = False Then Me.WindowState = 1

OHOURNOW = Hour(Now)

'----------------------------------------------------
DoEvents
If CID_RUN_ME_VAR_TO_EXIT_HAPPEN = True Then Unload CID_Run_Me: Exit Sub

TIMER_SHELL_PROGRAM_SET_EXE_AT_STARTER.Enabled = True


TIMER8_LEFT_THE_FORM_LOAD_ENTRY_YET = True

Me.Show

Call MNU_ADMINISTRATOR_Click

Timer_For_Menu_Gone.Enabled = True

End Sub




Private Sub MNU_00_MUSIC_INFO_LOGGER_TXT_Click()

I_N_NOTEPAD = "C:\Program Files (x86)\Notepad++\notepad++.exe"
If Dir(I_N_NOTEPAD) = "" Then
    I_N_NOTEPAD = "C:\Program Files\Notepad++\notepad++.exe"
End If
Path_And_FileName = "D:\VB6\VB-NT\00_Best_VB_01\WinAmp Logger\00 Music Info Logger.txt"
If Dir(I_N_NOTEPAD) <> "" Then
    Shell I_N_NOTEPAD + " """ + Path_And_FileName + """", vbMaximizedFocus
    Exit Sub
End If

Shell "NotePad " + Path_And_FileName, vbMaximizedFocus
Beep

End Sub

Private Sub Mnu_Contact_Me_For_HardCore_Click()
If Mnu_My_Email.Visible = False Then
Mnu_My_Email.Visible = True
MNU_MY_NAME.Visible = True
MNU_GET_EXPLORER_CURRENT_SESSION_CLIPBOADIN.Visible = False
Mnu_Contact_Me_For_HardCore.Caption = "Contact Me For HardCore __ 077722224555"
Exit Sub
Else
Timer_For_Menu_Gone.Enabled = True
End If
End Sub

Private Sub MNU_GET_EXPLORER_CURRENT_SESSION_CLIPBOADIN_Click()
Call FindWindow_Get_All_Explorer
End Sub

Private Sub Mnu_Musicday_Trim_NOTEPAD_Click()

I_N_NOTEPAD = "C:\Program Files (x86)\Notepad++\notepad++.exe"
If Dir(I_N_NOTEPAD) = "" Then
    I_N_NOTEPAD = "C:\Program Files\Notepad++\notepad++.exe"
End If
Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg-TRIM__" + GetComputerName + ".txt"
If Dir(I_N_NOTEPAD) <> "" Then
    Shell I_N_NOTEPAD + " """ + Path_And_FileName + """", vbMaximizedFocus
    Exit Sub
End If

Shell "NotePad " + Path_And_FileName, vbMaximizedFocus
Beep
End Sub

Private Sub MNU_NOTEPAD_01_TOTAL_LIST_TXT_Click()
I_N_NOTEPAD = "C:\Program Files (x86)\Notepad++\notepad++.exe"
If Dir(I_N_NOTEPAD) = "" Then
    I_N_NOTEPAD = "C:\Program Files\Notepad++\notepad++.exe"
End If
Path_And_FileName = "D:\0 00 MUSIC ---\04 My Music\01 Total List.txt"
If Dir(I_N_NOTEPAD) <> "" Then
    Shell I_N_NOTEPAD + " """ + Path_And_FileName + """", vbMaximizedFocus
    Exit Sub
End If

Shell "NotePad " + Path_And_FileName, vbMaximizedFocus
Beep
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'----------------------------------------------------------------------
'IF PROGRESS BAR CHANGED TO PICTURE BOX AS See RIGHT THING
'THEN REPLACE HER
'WHEN AFTER THE OCX FILE HAS MAYBE HAD A NEW VERSION AT LOAD TIME
'AND AFTER NEW INSTALL VB OR SOMETHING UPGRADED ONCE BEFORE FIRST FIXER
'----------------------------------------------------------------------
'FORGET HOW MAKE A SMOOTH SCROOL BAR
'AT THE MOMENT
'----------------------------------------------------------------------
'THE DRIVE FOR SCROLL BAR IS
'THE OCX
'----------------------------------------------------------------------
'MICROSOFT WINDOWS COMMON CONTROLS 5.0 SP 2
'1..
'C:\WINDOWS\SYSWOW64\COMCTL32.OCX
'2..
'C:\WINDOWS\SYSTEM32\COMCTL32.OCX
'----------------------------------------------------------------------
'LISTVIEW ALSO GET A DRIVER UPPER PROBLEM
'IF THE LIST VIEW NOT ON SCREEN BUT USE AS THE VARIABLE TYPE
'THEN THAT OCX IS
'IN TOOLBOX RIGHT CLICK COMPONENT
'--------------------------------
'MICROSOFT WINDOWS COMMON CONTROLS-2 6.0 SP 6
'C:\WINDOWS\SYSWOW64\MSCOMCT2.OCX
'------------------------------------------
'21 FEB 2017 11:14 PM LONG DAY NOT ANY FOOD
'------------------------------------------


End Sub

Private Sub Timer_For_Menu_Gone_Timer()
Timer_For_Menu_Gone.Enabled = False
Mnu_My_Email.Visible = False
MNU_MY_NAME.Visible = False
MNU_GET_EXPLORER_CURRENT_SESSION_CLIPBOADIN.Visible = True
Mnu_Contact_Me_For_HardCore.Caption = "Contact Me"

End Sub

Private Sub Timer_KEY_ESCAPE_Timer()
If IsIDE = True Then Timer_KEY_ESCAPE.Enabled = False: Exit Sub
If IsIDE = False Then Timer_KEY_ESCAPE.Enabled = False: Exit Sub

Dim bdf As Long

KBPress2 = 0
vcode = 0
For i = 5 To 255
    bdf = GetAsyncKeyState(i)
    If bdf < -300 Then vcode = i: Exit For
Next

If vcode <> O_VCODE Then
    'ESCAPE
    If vcode = 27 Then Unload Me
End If
O_VCODE = vcode

End Sub

Private Sub Timer_SAVE_AGRO_STORE2_WINAMP_Timer()
Call AGROSAVE2
End Sub

Sub TIMER_SHELL_PROGRAM_SET_EXE_AT_STARTER_TIMER()

EXE_CASE = EXE_CASE + 1
If EXE_CASE > 8 Then EXE_CASE = 1

If EXE_CASE = 8 Then
    TIMER_SHELL_PROGRAM_SET_EXE_AT_STARTER.Enabled = False
    Exit Sub
End If

'----------------------------------------------------
DoEvents
If CID_RUN_ME_VAR_TO_EXIT_HAPPEN = True Then Unload CID_Run_Me: Exit Sub
'----------------------------------------------------

'------------------------------------------------
'CODE EXE PROGRAM TAKE A WHILE TO LOAD AT STARTER
'AND THEN BEST RUN BY HERE TIMER
'------------------------------------------------

If EXE_CASE = 1 Then
    'Call Mnu_TAIL_EXE_BLUETOOTH_LOGGER_Click
End If
If EXE_CASE = 2 Then
    'Call Mnu_TAIL_EXE_BLUETOOTH_LOGGER_Click
End If
If EXE_CASE = 3 Then
    'Call Mnu_TAIL_EXE_BLUETOOTH_LOGGER_Click
End If

'If EXE_CASE = 4 Then
'    Call Mnu_TAIL_EXE_BLUETOOTH_LOGGER_Click
'End If

If EXE_CASE = 5 Then
    Call Mnu_Musicday_Trim_TAIL_EXE_Click
End If

If EXE_CASE = 6 Then
    Call MNU_AUTO_HOT_KEY_MUSIC_ON_ALL_THE_TIME_Click
End If

If EXE_CASE = 7 Then
    'Call MNU_WINAMP_FNOOB_Click
End If

If EXE_CASE = 8 Then
    Rf = FindWindow(vbNullString, "WMICPU1 BIG")
    If MNU_CPU_PROG_DONT_USE_IT.Checked = False Then
        If Rf = 0 Then
            'Shell "D:\VB6\VB-NT\00_Best_VB_01\Wmi_CPU\WMICPU1 BIG.exe", vbNormalNoFocus
        End If
    End If
End If

End Sub

Sub LoadHwndCounts()


'------------------------------------------
'TAKEN OUT OF CODE 2017 JAN
'SPEED REDUCTION HERE WITH AT LOADTIME ALSO
'------------------------------------------

Exit Sub

TDX = App.Path + "\0TextData\" + GetComputerName + "\2HwndCount.txt"
If FS.FileExists(TDX) Then
    Set F = FS.GetFile(TDX)
    If F.Size = 0 Then TDX = App.Path + "\0TextData\" + GetComputerName + "\2HwndCount-BAK.txt"
Else
    TDX = App.Path + "\0TextData\" + GetComputerName + "\2HwndCount-BAK.txt"
    Exit Sub
End If
    
Call FileInUseDelay(TDX, 0)
GUG = FreeFile
Open TDX For Input As #GUG
    Line Input #GUG, t1$
    Line Input #GUG, t2$
    Line Input #GUG, t3$
    Line Input #GUG, t4$
Close #GUG



OCurProcHwnd = Val(t1$)
CurProcHwnd_Count = Val(t2$)
CurProcHwnd_Count_Yester = Val(t4$)

If DateValue(t3$) <> Int(Now) Then
    AGRONOW = 0
    Call AGROSTORESAVE
    
    CurProcHwnd_Count_Yester = CurProcHwnd_Count
    CurProcHwnd_Count = 0
End If
DayNow = Day(Now)


TDX = App.Path + "\0TextData\" + GetComputerName + "\2HwndWinHovCount.txt"
If FS.FileExists(TDX) Then
    Set F = FS.GetFile(TDX)
    If F.Size = 0 Then TDX = App.Path + "\0TextData\" + GetComputerName + "\2HwndWinHovCount-BAK.txt"
Else
    TDX = App.Path + "\0TextData\" + GetComputerName + "\2HwndWinHovCount-BAK.txt"
End If

Call FileInUseDelay(TDX, 0)

GUG = FreeFile
Open TDX For Input As #GUG
Line Input #GUG, t1$
Line Input #GUG, t2$
Line Input #GUG, t3$
Line Input #GUG, t4$
Close #GUG


OCurWinHovHwnd = Val(t1$)
CurWinHovHwnd_Count = Val(t2$)
CurWinHovHwnd_Count_Yester = Val(t4$)

If DateValue(t3$) <> Int(Now) Then

    CurWinHovHwnd_Count_Yester = CurWinHovHwnd_Count
    CurWinHovHwnd_Count = 0
End If
DayNow2 = Day(Now)


Call AGROSTORESAVE

End Sub


Sub SaveHwndCount()
    
'If SaveHWNDModNOW > Now Then Exit Sub
'SaveHWNDModNOW = Now + TimeSerial(0, 0, 20)

On Error Resume Next

If CurProcHwnd_Count = 0 Then Exit Sub

TDX = App.Path + "\0TextData\" + GetComputerName + "\2HwndCount.txt"
Call FileInUseDelay(TDX, 1)

GUG = FreeFile
Open TDX For Output As #GUG
    Print #GUG, str(OCurProcHwnd)
    Print #GUG, str(CurProcHwnd_Count)
    Print #GUG, Date
    Print #GUG, str(CurProcHwnd_Count_Yester)
Close #GUG

Set F = FS.GetFile(TDX)
If F.Size > 0 Then
    Call FileInUseDelay(TDX, 0)
    Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\2HwndCount-BAK.txt", 1)
    F.Copy App.Path + "\0TextData\" + GetComputerName + "\2HwndCount-BAK.txt"
End If

TDX = App.Path + "\0TextData\" + GetComputerName + "\2HwndWinHovCount.txt"
Call FileInUseDelay(TDX, 1)
GUG = FreeFile
Open TDX For Output As #GUG
Print #GUG, str(OCurWinHovHwnd)
Print #GUG, str(CurWinHovHwnd_Count)
Print #GUG, Date
Print #GUG, str(CurWinHovHwnd_Count_Yester)
Close #GUG

Set F = FS.GetFile(TDX)
If F.Size > 0 Then
    Call FileInUseDelay(TDX, 0)
    Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\2HwndWinHovCount-BAK.txt", 1)
    F.Copy App.Path + "\0TextData\" + GetComputerName + "\2HwndWinHovCount-BAK.txt"
End If

End Sub




Private Sub Form_Unload(Cancel As Integer)

'Unload DSkeybd_F
'Call SetKeyboardHook(0, 0, 0)      '# Release the hook when closing

'--------------------------------------------------------
'TEST WHAT IS UNLOADING EVERY HOUR
'HANDLE GETTING HIGH PROBLEM GUESSING
'--------------------------------------------------------
'Cancel = True
'--------------------------------------------------------
'FIND AND ADD BETTER CODE NOT EXACT AT HOUR
'START_TIME_HANDLE_HIGH_FAULT = Now + TimeSerial(3, 0, 0)
'--------------------------------------------------------
'Check Reload_Me __ TO Adddtion
'--------------------------------------------------------

CID_RUN_ME_VAR_TO_EXIT_HAPPEN = True

Call zzCheckTimer_Timer

Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\2Total Music.txt", 1)
GUG = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\2Total Music.txt" For Output As #GUG
    Print #GUG, TotalMusic1
    Print #GUG, TotalMusic2
    Print #GUG, TotalMusicDay
Close GUG

Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\2CidRunME.txt", 1)
GUG = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\2CidRunME.txt" For Append As #GUG
    Print #GUG, "Start = "; Format$(Startuptime, "DD/MM/YYYY HH:MM:SS"); " ";
    Print #GUG, "End = "; Format$(Now, "DD/MM/YYYY HH:MM:SS"); " ";
    wds$ = "         "
    LSet wds$ = Trim(str(Mousey))
    Print #GUG, "Mouse = "; wds$; " ";
    wds$ = Trim(str(Keyy))
    Print #GUG, "Keyboard = "; wds$
Close #GUG


AGRONOW = 0
Call AGROSTORESAVE

'----------

MustUnload = True

Dim Form1 As Form

For Each Form1 In Forms
    If Form1.Caption <> Me.Caption Then
        Unload Form1
        Set Form1 = Nothing
        DoEvents
    End If
Next Form1


If Reload_Me = True Then
    Reset
    Reload_Me = False
    
    'COMMAND HERE DOES NOT EXECUTEAS FORM MOST LIKEY ALREADY GOING
    Unload CID_Run_Me
    DoEvents
    Load CID_Run_Me
    Call Form_Load
    
    Cancel = True
    Exit Sub
End If

'----------
End
'----------

End Sub


Private Sub Label1_Click()
Call Command1_Click
End Sub

Private Sub Label14_Click()
Call Command3_Click
End Sub





Function vbCompiled(Optional hWndMain As Variant, Optional Buffer As String) As Boolean

Dim nRtn As Long

Buffer = Space$(256)

nRtn = GetModuleFileNameA(hWndMain, Buffer, Len(Buffer))
Buffer = UCase(Left(Buffer, nRtn))

End Function


Function vbCompiled2(Optional hWndMain As Variant, Optional Buffer As String) As Boolean

Dim nRtn As Long

Buffer = Space$(256)

nRtn = GetModuleFileNameA(0&, Buffer, Len(Buffer))

Buffer = UCase(Left(Buffer, nRtn))

If Right(Buffer, 8) = "\VB6.EXE" Then
    vbCompiled2 = False
Else
    vbCompiled2 = True
End If

End Function


Private Sub Label8_Click()
'Label8.CAPTION
End Sub

Private Sub LastPressTrigTimer_Timer()
'If LastPressTrig = True Then
'LastPressTrig = False

'THIS WORKS ON KEY PRESS TRIGS THIS ONCE

LastPressTrigTimer.Enabled = False
Call CID_Run_Me.Lastpress


End Sub


Private Sub List4_Click()
'List4.
End Sub

Private Sub Menu_Date1991_Click()
'Shell Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\Date1991\date1991.exe", vbNormalFocus
End Sub
Private Sub Menu_Pulse_Click()
'Shell "D:\VB6\VB-NT\pulse\pulse.exe", vbNormalFocus
End Sub
Private Sub Menu_Alcohol120_Click()
'Shell "R:\Program Files\Alcohol Soft\Alcohol 120\Alcohol.exe", vbNormalFocus
End Sub

Private Sub Menu_AttachmentExtractor_Click()
'Shell "D:\VB6\VB-NT\Attachment Extractor 672673312002\Attachment Extractor.exe", vbNormalFocus
End Sub

Private Sub Menu_CreativePlayCenter_Click()
'Shell "R:\Program Files\Creative\SBAudigy\PlayCenter2\CTPlay2.exe", vbNormalFocus
End Sub

Private Sub Menu_KeyboardM_Click()
Call Label6Hit
End Sub



Private Sub Menu_NoteCallWinamp_Click()
'Shell "explorer D:\VB6\VB-NT\Cid-Run-Me\0TextData\WinAmpLog\Winamp-list-CallerAmp.txt"
'Shell "notepad D:\VB6\VB-NT\Cid-Run-Me\0TextData\WinAmpLog\Winamp-list-CallerID.txt", vbNormalFocus
'Shell "notepad C:\Program Files\Winamp Caller\Current Play To Text File Append.txt", vbNormalFocus

'Shell "notepad C:\Program Files\00 WinAmp's\Winamp Caller\Current Play To Text File Append.txt"
End Sub

Private Sub Menu_NoteGoldWinamp_Click()
Shell "notepad " + App.Path + "\0TextData\" + GetComputerName + "\WinAmpLog\Winamp-list-GoldAmp.txt", vbNormalFocus
End Sub


Private Sub Menu_StopCurrent_Click()

Call Command2_Click

End Sub

Private Sub Menu_KnowProcess_Click()
Beep

End Sub

Private Sub MNU_AUTO_HOT_KEY_MUSIC_ON_ALL_THE_TIME_Click()

'TEMP SHUT IT DOWN ERROR IN SCRIPTOR AHK CODE
'--------------------------------------------

I_N_AUTOHOTKEY = "C:\Program Files\AutoHotkey\AutoHotkey.exe"
Path_And_FileName = "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 08-WINAMP VOLUME UP ON TRACK CHANGE IF PLAYING.AHK"
If Dir(I_N_AUTOHOTKEY) <> "" Then
    
    If MNU_AUTO_HOT_KEY_MUSIC_ON_ALL_THE_TIME_AT_FOR_LOAD_WHEN_RELOAD_ON = False Then
        Shell I_N_AUTOHOTKEY + " """ + Path_And_FileName + """", vbMinimizedFocus
        
        Beep
        MNU_AUTO_HOT_KEY_MUSIC_ON_ALL_THE_TIME_AT_FOR_LOAD_WHEN_RELOAD_ON = True
    End If
    Exit Sub
Else
    MsgBox "THE __ NOT EXIST" + vbCrLf + vbCrLf + I_N_AUTOHOTKEY
    Beep
End If

End Sub

Private Sub Mnu_TAIL_EXE_BLUETOOTH_LOGGER_Click()
'HAVE LOOK KEEP THE MUSIC ON

TAIL_EXE_BLUETOOTH_LOGGER_CASE = TAIL_EXE_BLUETOOTH_LOGGER_CASE + 1
If TAIL_EXE_BLUETOOTH_LOGGER_CASE > 4 Then TAIL_EXE_BLUETOOTH_LOGGER_CASE = 1

DoEvents
If CID_RUN_ME_VAR_TO_EXIT_HAPPEN = True Then Unload CID_Run_Me: Exit Sub

'I_N_TAIL = "C:\Program Files\# NO INSTALL REQUIRED\Tail-4.2.12\Tail.exe"
I_N_TAIL = "C:\PStart\# NOT INSTALL REQUIRED\Tail\Tail.exe"
If Dir(I_N_TAIL) = "" Then
    MsgBox "THE EXE FILE __ NOT EXIST" + vbCrLf + vbCrLf + I_N_TAIL
    Beep
    Exit Sub
End If

'TAIL_AT_FORM_LOAD_2 = False

'1-ASUS-EEE
If TAIL_EXE_BLUETOOTH_LOGGER_CASE = 1 Then
    Path_And_FileName = "C:\0 00 LOGGERS TEXT\BlueToothView Logger\BlueToothView_Logger -- 1-ASUS-EEE -- StarTech.txt"
    If FindWinPart_ANY_STRING("Tail for Win32 - [Non-Workspace Files - " + Path_And_FileName + "]") = 0 Then
            Shell I_N_TAIL + " """ + Path_And_FileName + """", vbMinimized
        Else
            'If IsIDE = True Then Stop
            'msgbox "MESSAGE BOX ONLY ONCE"
    End If
End If

'1-ASUS-X5DIJ.txt"
If TAIL_EXE_BLUETOOTH_LOGGER_CASE = 2 Then
    Path_And_FileName = "C:\0 00 LOGGERS TEXT\BlueToothView Logger\BlueToothView_Logger 1-ASUS-X5DIJ.txt"
    If FindWinPart_ANY_STRING("Tail for Win32 - [Non-Workspace Files - " + Path_And_FileName + "]") = 0 Then
            Shell I_N_TAIL + " """ + Path_And_FileName + """", vbMinimized
        Else
            'If IsIDE = True Then Stop
            'msgbox "MESSAGE BOX ONLY ONCE"
    End If
End If

'7-ASUS-GL522VW
If TAIL_EXE_BLUETOOTH_LOGGER_CASE = 3 Then
    Path_And_FileName = "C:\0 00 LOGGERS TEXT\BlueToothView Logger\BlueToothView_Logger -- 7-ASUS-GL522VW.txt"
    If FindWinPart_ANY_STRING("Tail for Win32 - [Non-Workspace Files - " + Path_And_FileName + "]") = 0 Then
            Shell I_N_TAIL + " """ + Path_And_FileName + """", vbMinimized
        Else
            'If IsIDE = True Then Stop
            'msgbox "MESSAGE BOX ONLY ONCE"
    End If
End If

End Sub


Private Sub Mnu_CIDBreak_Click()

If Mnu_CIDBreak.Checked = True Then Mnu_CIDBreak.Checked = False: Exit Sub
If Mnu_CIDBreak.Checked = False Then Mnu_CIDBreak.Checked = True


End Sub

Private Sub Mnu_Cpu_Click()
'Rf = FindWindow(vbNullString, "WMICPU1 BIG")
'If Rf = 0 Then
'Shell Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\Wmi_CPU\WMICPU1 BIG.exe", vbNormalFocus
'End If

'Rf = FindWindow(vbNullString, "WMICPU2 MINI")
'If Rf = 0 Then
'Shell Mid$(App.Path, 1, 3) + "VB6\VB-NT\00_Best_VB_01\Wmi_CPU\WMICPU2 MINI.exe", vbNormalFocus
'End If
End Sub

Private Sub LABLE_SONG_PLAYING()

Err.Clear

On Error Resume Next

TF2 = FindWindow("Winamp v1.x", vbNullString)

If TF2 = 0 Or WindowVisible(TF2) = False Then Exit Sub

Dim Rect7 As RECT

hWnd8 = GetWindowRect(TF2, Rect7)

'msgresult33 = SendMessage(TF2, WM_USER, 1, ByVal IPC_GETOUTPUTTIME)
'------------------------------------------------------------------
'WHEN AN INTENET DIGITAL BROADCAST THE OUTPUT TIME WILL BE NAUGHT
'SO WE WIL REQUIRE NAME OF SONG TO TELL TO IGNR THAT ONE
'------------------------------------------------------------------
CurrentSongx1_LONG = ""
For R2 = 1 To 4
    If TF2 > 0 Then
        CurrentSongx1 = GetWindowTitle(TF2)
        CurrentSongx1 = Trim(Replace(CurrentSongx1, "[Stopped]", ""))
        CurrentSongx1 = Trim(Replace(CurrentSongx1, "[Paused]", ""))
        ' If Len(CurrentSongx1) < Len(CurrentSongx1_LONG) Then
            ' IF STRING WAS SHORT AT FRONT
            ' COMPARE IF NEW STRING IS LONGER AND ALSO CONTAIN FIRST
            ' WHY ABORT
            ' STRING IS MOST LIKLEY TO BE SHORT FIRST REQUESTER
            ' -------------------------------------------------
            ' SO GET FEW IDEA
            
            'If InStr(CurrentSongx1_LONG, CurrentSongx1) Then
            '    CurrentSongx1 = CurrentSongx1_LONG
            'End If
        ' End If
        ' CurrentSongx1_LONG = CurrentSongx1
    End If
Next

Label_WINAMP_VER_ON_FORM = "WINAMP v5.552 (x86) - Apr 10 2009"

If CurrentSongx1 = BEFORE_CurrentSongx Then Exit Sub
BEFORE_CurrentSongx = CurrentSongx1

Test_RADIO = SendMessage(TF2, WM_USER, 1, ByVal IPC_GETOUTPUTTIME)
If Test_RADIO = -1 Then RADIO_SONG = True
If Test_RADIO > -1 Then RADIO_SONG = False


'Txz1$ = "NOT USE EXE FILENAME WHEN DIGITAL HARDWIRE RADIO"
'If RADIO_SONG = False Then
    Dim oWinamp
    Set oWinamp = CreateObject("WinampCOM.Application")
    Txz1$ = oWinamp.CurrentSongFileName
    Set oWinamp = Nothing
    
    If Err.Number > 0 Then
        Label_Tune_File_Name.Caption = "WinampCOM.Application __ ERROR GET"
        Beep
        Exit Sub
    End If
    
'End If

'Call FileInUseDelay(App.Path + "\00 Music Info Logger.txt", 1)

'gug = FreeFile
'Open App.Path + "\00 Music Info Logger.txt" For Append As #gug

If Test_RADIO = -1 Then Label_RADIO_STREAMER_OR_MP3_PLAYING = "Internet Hardwire Radio Playing is Not MP3 Numeric WinAmp WinRiper Stream"
If Test_RADIO > -1 Then Label_RADIO_STREAMER_OR_MP3_PLAYING = "MP3 Playing Numeric On"
If InStr(Txz1$, "http://") = 1 Then
    Label_RADIO_STREAMER_OR_MP3_PLAYING = Replace(Label_RADIO_STREAMER_OR_MP3_PLAYING, "MP3 Playing Numeric On", "HTTP:\\ MP3 Playing Numeric On")
End If

tt = String$(70, "-") + vbCrLf
tt = tt + Format$(Now, "DDD DD-MM-YYYY HH:MM:SS") + vbCrLf
tt = tt + String$(70, "-") + vbCrLf
tt = tt + "Play % Now = " + Label7 + vbCrLf
tt = tt + "Play Lenght = " + Label4 + vbCrLf
tt = tt + "Play Now = " + Label13 + vbCrLf
tt = tt + String$(70, "-") + vbCrLf
tt = tt + "File Name = " + Txz1$ + vbCrLf
tt = tt + "Tune      = " + Label11 + vbCrLf
tt = tt + String$(70, "-") + vbCrLf
tt = tt + "" + vbCrLf
tt = tt + String$(70, "-") + vbCrLf

'Clipboard.Clear
'Clipboard.SetText tt

STRIP_FIRST_DOT = Trim(Mid(Label11, InStr(Label11, ".") + 2))
'CHECK NURMERIC IF FRONT DOT
'LATER
'---------------------------
Label_Tune_File_Name.Caption = Txz1$
Label_Tune__________.Caption = STRIP_FIRST_DOT



'Print #gug, TT;
'Close #gug

'Shell "notepad " + App.Path + "\00 Music Info Logger.txt", vbNormalFocus


End Sub


Private Sub MNU_CLIP_SONG_PLAYING_Click()

Err.Clear

On Error Resume Next

TF2 = FindWindow("Winamp v1.x", vbNullString)

If TF2 = 0 Or WindowVisible(TF2) = False Then Exit Sub

Dim Rect7 As RECT

hWnd8 = GetWindowRect(TF2, Rect7)


'msgresult33 = SendMessage(TF2, WM_USER, 1, ByVal IPC_GETOUTPUTTIME)
'------------------------------------------------------------------
'WHEN AN INTENET DIGITAL BROADCAST THE OUTPUT TIME WILL BE NAUGHT
'SO WE WIL REQUIRE NAME OF SONG TO TELL TO IGNR THAT ONE
'------------------------------------------------------------------


If TF2 > 0 Then
    
    CurrentSongx1 = GetWindowTitle(TF2)
    CurrentSongx1 = Trim(Replace(CurrentSongx1, "[Stopped]", ""))
    CurrentSongx1 = Trim(Replace(CurrentSongx1, "[Paused]", ""))

End If

If CurrentSongx1 <> "" And InStr(CurrentSongx1, "(FNOOB TECHNO RADIO) - Winamp") = 0 Then
    'Exit Sub
Else
    RADIO_SONG = True
End If


Txz1$ = "NOT USE EXE FILENAME WHEN DIGITAL HARDWIRE RADIO"
If RADIO_SONG = False Then
    Dim oWinamp
    Set oWinamp = CreateObject("WinampCOM.Application")
    Txz1$ = oWinamp.CurrentSongFileName
    Set oWinamp = Nothing
    
    If Err.Number > 0 Then
        Beep
        Exit Sub
    End If
    
End If

'Call FileInUseDelay(App.Path + "\00 Music Info Logger.txt", 1)

'gug = FreeFile
'Open App.Path + "\00 Music Info Logger.txt" For Append As #gug

tt = String$(70, "-") + vbCrLf
tt = tt + Format$(Now, "DDD DD-MM-YYYY HH:MM:SS") + vbCrLf
tt = tt + String$(70, "-") + vbCrLf
tt = tt + "Play % Now = " + Label7 + vbCrLf
tt = tt + "Play Lenght = " + Label4 + vbCrLf
tt = tt + "Play Now = " + Label13 + vbCrLf
tt = tt + String$(70, "-") + vbCrLf
tt = tt + "File Name = " + Txz1$ + vbCrLf
tt = tt + "Tune      = " + Label11 + vbCrLf
tt = tt + String$(70, "-") + vbCrLf
tt = tt + "" + vbCrLf
tt = tt + String$(70, "-") + vbCrLf

Clipboard.Clear
Clipboard.SetText tt


Beep

'Print #gug, TT;
'Close #gug

'Shell "notepad " + App.Path + "\00 Music Info Logger.txt", vbNormalFocus

End Sub

Private Sub MNU_CPU_PROG_Click()
Rf = FindWindow(vbNullString, "WMICPU1 BIG")
If Rf = 0 Then
    Shell "D:\VB6\VB-NT\00_Best_VB_01\Wmi_CPU\WMICPU1 BIG.exe", vbNormalNoFocus
End If

End Sub

Private Sub MNU_CPU_PROG_DONT_USE_IT_Click()
If MNU_CPU_PROG_DONT_USE_IT.Checked = True Then MNU_CPU_PROG_DONT_USE_IT.Checked = False: Exit Sub
If MNU_CPU_PROG_DONT_USE_IT.Checked = False Then MNU_CPU_PROG_DONT_USE_IT.Checked = True
End Sub

Private Sub Mnu_Dont_LowerART_Click()

If Mnu_Dont_LowerART.Checked = True Then Mnu_Dont_LowerART.Checked = False: Exit Sub
If Mnu_Dont_LowerART.Checked = False Then Mnu_Dont_LowerART.Checked = True

End Sub

Private Sub Mnu_exit_Click()

'End
Unload Me

End Sub


Private Sub MNU_BRing_Front_Click()
'Call FindWinPartFront
'MNU_BRing_Front.Caption = "Bring All Front -- " + Format(Now, "DD-MMM-YYYY HH:MM:SS")

GO_SUB = False
For i = 1 To 255
    bdf = GetAsyncKeyState(i)
    If bdf < 0 Then
        If i = 39 Then i = 0
        If i = 116 Then i = 0 'F5 TEST ISIDE
        If i = 16 Then GO_SUB = True 'LEFT SHIFT KEY
        If GO_SUB = True Then
            'Call KeyOrMouse
            MNU_BRing_Front.Caption = "Bring All Front"
            Exit Sub
        End If
    End If
Next


'TEST=
ExplorerGone = True
ExplorerGone_TEST = True
Timer_EXPLORER_GONE.Enabled = True
'getkeyasync
End Sub

Private Sub Timer_EXPLORER_GONE_Timer()


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

    ExplorerGone = True
    
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

If FindWindow("Progman", "Program Manager") <> 0 And ExplorerGone = True Then

    'BRING WINDOWS FRONT
    'i = FindWinPartExplorerGone(False) ' True = Quite Mode Don't Display  Result
    i = FindWinPartExplorerGone(True) ' True = Quite Mode Don't Display  Result
'    Debug.Print str(i) + " Windows Brought Forward"
 
    
    'If ExplorerGone_TEST = True Then
    '    ExplorerGone_TEST = False
MNU_BRing_Front.Caption = "Bring " + str(i) + " Forward Form Explorer Crash@" + Format(Now, "DD-MMM HH:MM:SS") + " Left Shift Reset Menu Opt"
    
    'Else
        'MNU_BRing_Front.Caption = "Bring " + str(i) + " Forward -- Explorer Crash/Terminated @ " + Format(Now, "DD-MMM-YYYY HH:MM:SS") + " Left Shift Menu Reset Button Menu"
    
    'End If
    Timer_EXPLORER_GONE.Enabled = False
    
    ExplorerGone = False
    Beep
    
End If



End Sub

Function FindWinPartExplorerGone(Q) As Long
'AND BRING ALL WINDOWS FORWARD
'ONLY VB SUFFERS FROM THIS

FindWinPartExplorerGone = False

Dim Rect8 As RECT



Dim Window_Title_String
Dim Test_HWND As Long, _
    Test_pid As Long, _
    Test_thread_id As Long

Dim cText As String

'Find the first window

'need this is you gona use this procedure get from CIDRun ME an another one

Dim Huge

'Test_HWND = FindWindow2(ByVal 0&, ByVal 0&)
Test_HWND = FindWindowDLL(ByVal 0&, ByVal 0&)
BR$ = ""
Do While Test_HWND <> 0
    
    
    DoEvents
    hwnd9 = GetWindowRect(Test_HWND, Rect8)
        Window_Title_String = GetWindowTitle(Test_HWND)
'        If InStr(Window_Title_String, "Double") > 0 Then Stop
        gws = GetWindowState(Test_HWND)
'        If gws = -1 Then gws2$ = "Window State Normal"
'        If gws = vbMinimized Then gws2$ = "Window State Minimized"
'        If gws = vbMaximized Then gws2$ = "Window State Maximized"
'
'
    If (Rect8.Top > 0 And Rect8.Left > 0) Or gws = vbMinimized Then
        Window_Title_String = GetWindowTitle(Test_HWND)
        If Window_Title_String <> "" And InStr(BR$, "-- " + Window_Title_String + " -- ") = 0 Then
            BR$ = BR$ + "-- " + Window_Title_String + " -- "
'            On Error Resume Next
'            AppActivate Window_Title_String, 0
'            On Error GoTo 0
            Huge = Huge + 1
            ef = SetForegroundWindow(hwnd9)
            'ef = Putfocus(hwnd9)
            ecute = Timer + 0.1
         
            Do
            
                DoEvents
            
                'Safety IN CASE TME RESETS BACK TO ZERO
                If Timer < ecute - 30 Then Exit Do
                
            Loop Until Timer > ecute
        
        End If
    End If
        
'retrieve the next window
Test_HWND = GetWindow(Test_HWND, GW_HWNDNEXT)

Loop

If Q = False Then MsgBox str(Huge) + " Windows Brought To Front", vbMsgBoxSetForeground

FindWinPartExplorerGone = Huge

End Function





Private Sub Mnu_KeyLoggDay_Click()
    Shell "NotePad " + App.Path + "\Text_Data_KeyLogg\Key Logger Text " + Format$(KeyLoggDate, "YYYY-mm-dd HH-MM-SS") + ".txt", vbNormalFocus
End Sub

Private Sub Mnu_KeyLoggFolder_Click()
Shell "Explorer.exe /e, " + App.Path + "\Text_Data_KeyLogg\", vbNormalFocus
End Sub

Private Sub Mnu_KeyLogText_Click()
    Shell "NotePad " + App.Path + "\Text_Data_KeyLogg\Key Logger Text.txt", vbNormalFocus
End Sub

Private Sub MNU_Logg_Folder_WEBCAM_Click()
'Shell "C:\Program Files\IceView\iceview.exe ""D:\VB6\VB-NT\00_Best_VB_04\Webcam_Motion\VBDataNoArchive\Web_Cam Pic Always.jpg"""

Shell "Explorer.exe /e, " + """D:\0 00 Art Loggers - WEBCAM\""", vbNormalFocus
Shell "Explorer.exe /e, " + """M:\0 00 Art Loggers - WEBCAM\""", vbMinimizedNoFocus



End Sub

Private Sub Mnu_LoggFolder_Click()
Shell "Explorer.exe /e, " + App.Path + "\0TextData\" + GetComputerName + "", vbMaximizedFocus

End Sub

Private Sub Mnu_Mini_Click()
'GT = GetForegroundWindow
'GT = FindWindow("#32770", vbNullString)
'GT1 = FindWinPart("NeroVision Express 3")
'GT2 = FindWinPart("DVD-TO-MPEG", True)
'LastCurProc5


GT = FindWindow("Digital Wave Player", vbNullString)
If GT > 0 Then Mini1 = Not Mini1
'If Mini1 = 0 And GT > 0 Then
'    ShowWindow GT, SW_MINIMIZE
'End If

'If Mini1 = -1 And GT > 0 Then
If GT > 0 Then
    ShowWindow GT, SW_NORMAL
End If


Exit Sub

'
GT = FindWindow("Digital Wave Player", vbNullString)
If NeroMin = 0 Then
GT1 = FindWinPart("NeroVision Express 3", True)
'GT2 = FindWinPart("Writing G:\00 DVD", True)
'GT2 = FindWinPart("#1 DVD Ripper", True)
'NeroMin = 1
If GT1 > 0 Then ShowWindow GT1, SW_MINIMIZE: NeroMin = 1: Exit Sub

Exit Sub
End If


If NeroMin = 0 Then
GT1 = FindWinPart("NeroVision Express 3", True)
'GT2 = FindWinPart("Writing G:\00 DVD", True)
'GT2 = FindWinPart("#1 DVD Ripper", True)
'NeroMin = 1
If GT1 > 0 Then ShowWindow GT1, SW_MINIMIZE: NeroMin = 1: Exit Sub

Exit Sub
End If

If NeroMin = 1 Then
GT1 = FindWinPart("NeroVision Express 3", False)
'GT2 = FindWinPart("Writing G:\00 DVD", False)
'GT2 = FindWinPart("#1 DVD Ripper", False)
If GT1 > 0 Then ShowWindow GT1, SW_NORMAL: NeroMin = 0
'If GT2 > 0 Then ShowWindow GT2, SW_NORMAL: NeroMin = 0
NeroMin = 0
End If

'Digital Wave Player



End Sub

Private Sub MNU_MOVE_EXP_Click()

If MNU_MOVE_EXP.Checked = True Then MNU_MOVE_EXP.Checked = False: Exit Sub
If MNU_MOVE_EXP.Checked = False Then MNU_MOVE_EXP.Checked = True

End Sub

Private Sub Mnu_MusicDay_Click()

I_N_NOTEPAD = "C:\Program Files (x86)\Notepad++\notepad++.exe"
If Dir(I_N_NOTEPAD) = "" Then
    I_N_NOTEPAD = "C:\Program Files\Notepad++\notepad++.exe"
End If
If Dir(I_N_NOTEPAD) <> "" Then
    Shell I_N_NOTEPAD + " " + App.Path + "\0TextData\" + GetComputerName + "\2Total_Music_Day_Log.txt", vbMaximizedFocus
    Exit Sub
End If

Shell "NotePad " + App.Path + "\0TextData\" + GetComputerName + "\2Total_Music_Day_Log.txt", vbMaximizedFocus
Beep

End Sub

Private Sub Mnu_Musicday_Trim_Click()

I_N_NOTEPAD = "C:\Program Files (x86)\Notepad++\notepad++.exe"
If Dir(I_N_NOTEPAD) = "" Then
    I_N_NOTEPAD = "C:\Program Files\Notepad++\notepad++.exe"
End If
Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg-TRIM__" + GetComputerName + ".txt"
If Dir(I_N_NOTEPAD) <> "" Then
    Shell I_N_NOTEPAD + " """ + Path_And_FileName + """", vbMaximizedFocus
    Exit Sub
End If

Shell "NotePad """ + Path_And_FileName + """", vbMaximizedFocus
Beep

End Sub

Private Sub MNU_MUSICDAY_TRIM_TAIL_IN_BROWSER_CHROME_Click()

'-------------------------------------------------------------------------------------
'If FindWindow(vbNullString, "Tail.exe") <> 0 Then Beep: Exit Sub
'-------------------------------------------------------------------------------------
'If TAIL_AT_FORM_LOAD = True Then
'    Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg-TRIM__" + GetComputerName + ".txt"
'    TAIL_AT_FORM_LOAD_1 = False
'    If FindWinPart_ANY_STRING("Tail for Win32 - [") > 0 Then
'        If FindWinPart_ANY_STRING(Path_And_FileName) > 0 Then
'            TAIL_AT_FORM_LOAD_2 = True
'            'Beep
'            'Exit Sub
'        End If
'    End If
'End If
'-------------------------------------------------------------------------------------

'I_N_TAIL = "C:\Program Files\# NO INSTALL REQUIRED\Tail-4.2.12\Tail.exe"
'If Dir(I_N_TAIL) = "" Then
'    'If TAIL_AT_FORM_LOAD_1 = False Then
'     If TAIL_MESSENGER_BOX_DONE = False Then
'        MsgBox "THE TAIL PROGRAM FOR DOCUMENT LOGGER IS NOT EXISTER THIS IS WHERE EXPECTED" + vbCrLf + vbCrLf + I_N_TAIL
'        TAIL_MESSENGER_BOX_DONE = True
'    End If
'End If

'METAL LITTMAN
'Path_And_FileName = "https://drive.google.com/file/d/0B2hfTi1eLYL5dUxXY2NicGx6R2c/view"
'ROIDSRIM
Path_And_FileName = "https://drive.google.com/file/d/0BwoB_cPOibCPV0RnUXNCZjJzT3M/view"
Path_And_FileName = WinAmp_Music_Lenght_Logg_TRIM__7_ASUS_GL522VW
'---------------
'WinAmp Music Lenght Logg-TRIM__7-ASUS-GL522VW.txt - Google Drive
'https://drive.google.com/file/d/0BwoB_cPOibCPV0RnUXNCZjJzT3M/view
'---------------

'If FindWinPart_ANY_STRING("Tail for Win32 - [Non-Workspace Files - " + Path_And_FileName + "]") = 0 Then

'----
'WinAmp Music Lenght Logg-TRIM__7-ASUS-GL522VW.txt - Google Drive
'https://drive.google.com/file/d/0B2hfTi1eLYL5dUxXY2NicGx6R2c/view
'----
        
        
'    Else
        'If IsIDE = True Then Stop
        'msgbox "MESSAGE BOX ONLY ONCE"
'End If

'Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg-TRIM__" + GetComputerName + ".txt"
'If Dir(I_N_TAIL) <> "" Then
'    Shell I_N_TAIL + " """ + Path_And_FileName + """", vbMinimizedNoFocus
'    Beep
'    Exit Sub
'End If
'
'TAIL_AT_FORM_LOAD_1 = False


i = "----" + vbCrLf
i = i + "WinAmp Music Lenght Logg-TRIM__7-ASUS-GL522VW.txt - Google Drive" + vbCrLf
i = i + WinAmp_Music_Lenght_Logg_TRIM__7_ASUS_GL522VW + vbCrLf
i = i + "----" + vbCrLf
i = i + "----" + vbCrLf
i = i + "COMPUTER_CODING_VB - Google Drive" + vbCrLf
i = i + "https://drive.google.com/drive/folders/0BwoB_cPOibCPV0ZsSklwcE92X3c" + vbCrLf
i = i + "----" + vbCrLf

'---------------
'WinAmp Music Lenght Logg-TRIM__7-ASUS-GL522VW.txt - Google Drive
'https://drive.google.com/file/d/0BwoB_cPOibCPV0RnUXNCZjJzT3M/view
'---------------


Clipboard.Clear
Clipboard.SetText i

'-----------------------------------------------------
'/C TO QUITE PASS
'/K TO STOP IF ANY COMMAND SWITCH INSTRUCTION IS WRONG
'-----------------------------------------------------
Path_And_FileName = "https://drive.google.com/file/d/0BwoB_cPOibCPV0RnUXNCZjJzT3M/view"
Path_And_FileName = WinAmp_Music_Lenght_Logg_TRIM__7_ASUS_GL522VW

Shell "CMD /C START """" /HIGH /MAX ""C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"" """ + Path_And_FileName + """", vbMaximizedFocus



End Sub


Private Sub Mnu_Musicday_Trim_TAIL_EXE_Click()

'-------------------------------------------------------------------------------------
'If FindWindow(vbNullString, "Tail.exe") <> 0 Then Beep: Exit Sub
'-------------------------------------------------------------------------------------
'If TAIL_AT_FORM_LOAD = True Then
'    Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg-TRIM__" + GetComputerName + ".txt"
'    TAIL_AT_FORM_LOAD_1 = False
'    If FindWinPart_ANY_STRING("Tail for Win32 - [") > 0 Then
'        If FindWinPart_ANY_STRING(Path_And_FileName) > 0 Then
'            TAIL_AT_FORM_LOAD_2 = True
'            'Beep
'            'Exit Sub
'        End If
'    End If
'End If
'-------------------------------------------------------------------------------------

'I_N_TAIL = "C:\Program Files\# NO INSTALL REQUIRED\Tail-4.2.12\Tail.exe"
I_N_TAIL = "C:\PStart\# NOT INSTALL REQUIRED\Tail\Tail.exe"
If Dir(I_N_TAIL) = "" Then
    'If TAIL_AT_FORM_LOAD_1 = False Then
     If TAIL_MESSENGER_BOX_DONE = False Then
        MsgBox "THE TAIL PROGRAM FOR DOCUMENT LOGGER IS NOT EXISTER THIS IS WHERE EXPECTED" + vbCrLf + vbCrLf + I_N_TAIL
        TAIL_MESSENGER_BOX_DONE = True
    End If
End If

If Dir(I_N_TAIL) <> "" Then
Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg-TRIM__" + GetComputerName + ".txt"
If FindWinPart_ANY_STRING("Tail for Win32 - [Non-Workspace Files - " + Path_And_FileName + "]") = 0 Then
        Shell I_N_TAIL + " """ + Path_And_FileName + """", vbMinimized
    Else
        'If IsIDE = True Then Stop
        'msgbox "MESSAGE BOX ONLY ONCE"
End If
End If

'Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg-TRIM__" + GetComputerName + ".txt"
'If Dir(I_N_TAIL) <> "" Then
'    Shell I_N_TAIL + " """ + Path_And_FileName + """", vbMinimizedNoFocus
'    Beep
'    Exit Sub
'End If
'
'TAIL_AT_FORM_LOAD_1 = False

End Sub

Private Sub Mnu_NotePad_Click()
If Mnu_NotePad.Checked = True Then Mnu_NotePad.Checked = False: Exit Sub
If Mnu_NotePad.Checked = False Then Mnu_NotePad.Checked = True

End Sub


Private Sub Mnu_Proc_Click()
If ProcLB.Visible = True Then ProcLB.Visible = False: Exit Sub
If ProcLB.Visible = False Then ProcLB.Visible = True
End Sub

Private Sub Mnu_Process_W_Win_Click()

ProcLB.List1.Clear
mdlProcess.List_ActiveProcess
mdlProcess.List_ActiveModules

Call OutputProcFileList

I_N_NOTEPAD = "C:\Program Files (x86)\Notepad++\notepad++.exe"
If Dir(I_N_NOTEPAD) = "" Then
    I_N_NOTEPAD = "C:\Program Files\Notepad++\notepad++.exe"
End If
If Dir(I_N_NOTEPAD) <> "" Then
    Shell I_N_NOTEPAD + " " + App.Path + "\00 Processes\Processes " + Format$(Now, "yyyy-mm-dd HH-MM-SS") + ".txt", vbMaximizedFocus
    Exit Sub
End If

Shell "notepad " + App.Path + "\00 Processes\Processes " + Format$(Now, "yyyy-mm-dd HH-MM-SS") + ".txt", vbMaximizedFocus

ProcLB.Visible = True

End Sub

Sub OutputProcFileList()

Path_Folder = App.Path + "\0TextData\" + GetComputerName + "\00 Processes"

If (Not DirExist(Path_Folder)) Then
    MkDirNested Path_Folder
End If

Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\00 Processes\Processes " + Format$(Now, "yyyy-mm-dd HH-MM-SS") + ".txt"

'If GetUserName <> "Matt3" Then Exit Sub
Dim Fr1
Call FileInUseDelay(Path_And_FileName, 1)
Fr1 = FreeFile
Open Path_And_FileName For Output As #Fr1

    Print #Fr1, Format$(Now, "DDD DD-MMM-YYYY HH:MM:SS")
    Print #Fr1, "------------------------------------------------------------------------"
    
    For R = 0 To ProcLB.List1.ListCount - 1
        Print #Fr1, ProcLB.List1.List(R)
    Next
    
    Print #Fr1, "------------------------------------------------------------------------"
Close #Fr1
End Sub

Private Sub Mnu_ProWOut_Win_Click()

End Sub

Private Sub Mnu_Publish_Click()
If Mnu_Publish.Checked = True Then Mnu_Publish.Checked = False: Exit Sub
If Mnu_Publish.Checked = False Then Mnu_Publish.Checked = True

End Sub


Private Sub Mnu_VB_Click()

If Mnu_VB.Checked = True Then Mnu_VB.Checked = False: Exit Sub
If Mnu_VB.Checked = False Then Mnu_VB.Checked = True

End Sub

'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
    If GetSpecialfolder(&H29) = "" Then
        If MESSENGER_BOX_ABOUT_GetSpecialfolder = False Then
            MsgBox "AN ERROR HAPPEN BECUASE GET GetSpecialfolder(&H29) DOES NOT WORK UNLESS ADMIN MODE" + vbCrLf + "MEANT TO RETURN THE PATH OF 64 BIT SYSWOW64 BIT FOLDER", vbMsgBoxSetForeground
        End If
        MESSENGER_BOX_ABOUT_GetSpecialfolder = True
        Exit Sub
    End If
    
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

    Dim MSGQ, iR
    Dim CODER_EXE_FILE_NAME_1
    Dim CODER_VBP_FILE_NAME_2
    Dim EXIT_TRUE
    
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
        iR = MsgBox("NOT IN IDE ARE YOU SURE", MSGQ, vbMsgBoxSetForeground)
        If iR <> vbYes Then Exit Sub
    End If
    
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
    Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + CODER_VBP_FILE_NAME_2 + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
    
End Sub

Private Sub MNU_VB_FOLDER_Click()
    Beep
    Dim CODER_EXE_FILE_NAME_1
    Dim CODER_VBP_FILE_NAME_2
    Dim EXIT_TRUE
    
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + CODER_EXE_FILE_NAME_1 + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
    
End Sub
Private Sub MNU_ADMINISTRATOR_Click()
    'Call MNU_ADMINISTRATOR_Click
    ' Beep
    MNU_ADMINISTRATOR.Caption = "Administrator ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "Administrator ON", "Administrator OFF")
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


Private Sub Mnu_Voiceonoff_Click()
    If Mnu_Voiceonoff.Checked = True Then Mnu_Voiceonoff.Checked = False: Exit Sub
    If Mnu_Voiceonoff.Checked = False Then Mnu_Voiceonoff.Checked = True
End Sub

Private Sub MNU_WINAMP_FNOOB_Click()
If FindWinPart_ANY_STRING("(FNOOB TECHNO RADIO) - Winamp") > 0 Then
    If MNU_WINAMP_FNOOB_AT_FORM_LOAD = False Then
        MNU_WINAMP_FNOOB_AT_FORM_LOAD = True
    
        Beep
    End If
    Exit Sub
End If
Beep


'I1 = """D:\0 00 MUSIC ---\04 My Music\01 Banging Tunes\00 Special\Sp 01 Will De Cruize - Ol Mate\# 00 hifi www.stream1.nubreaks.com.m3u"""

'----------------------------------
'HERE NOT USE ANYMORE I USR DEFAULT
'----------------------------------
'I1 = """http://play.fnoobtechno.com:2199/tunein/fnoobtechno.mp3"""
'----------------------------------

'I2 = """C:\Program Files (x86)\K-Lite Codec Pack\MPC-HC64\mpc-hc64.exe"""

i2 = """C:\Program Files (x86)\Winamp\winamp.exe"""
Shell i2 + " " + I1, vbNormalFocus

End Sub

Private Sub Mnu_WinampLenghtLogg_Click()

Call AGROSAVE2

I_N_NOTEPAD = "C:\Program Files (x86)\Notepad++\notepad++.exe"
If Dir(I_N_NOTEPAD) = "" Then
    I_N_NOTEPAD = "C:\Program Files\Notepad++\notepad++.exe"
End If
If Dir(I_N_NOTEPAD) <> "" Then
    Shell I_N_NOTEPAD + " " + App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Total.txt", vbMaximizedFocus
    Exit Sub
End If


Shell "notepad " + App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Total.txt", vbMaximizedFocus

Exit Sub



If FS.FileExists(App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day.txt") = True Then
LF1 = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day Total.txt" For Append As #LF1
LF2 = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day.txt" For Input As #LF2
Do
    Line Input #LF2, ll$
    If Trim(ii$) <> "" Then Print #LF1, ll$
Loop Until EOF(LF2)
Close #LF1, #LF2
Kill App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day.txt"
End If

I_N_NOTEPAD = "C:\Program Files (x86)\Notepad++\notepad++.exe"
If Dir(I_N_NOTEPAD) = "" Then
    I_N_NOTEPAD = "C:\Program Files\Notepad++\notepad++.exe"
End If
If Dir(I_N_NOTEPAD) <> "" Then
    Shell I_N_NOTEPAD + " " + App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Total.txt", vbMaximizedFocus
    Exit Sub
End If


Shell "notepad " + App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Total.txt", vbMaximizedFocus

Exit Sub


If Dir(App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg.txt") <> "" Then
    Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Total.txt", 1)
    LF1 = FreeFile
    Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Total.txt" For Append As #LF1
    
    Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg.txt", 0)
    
    LF2 = FreeFile
    Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg.txt" For Input As #LF2
    Do
        Line Input #LF2, ll$
        If Trim(ii$) <> "" Then Print #LF1, ll$
    Loop Until EOF(LF2)
    Close #LF1, #LF2
    Kill App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg.txt"
End If



End Sub

Private Sub Timer_FORM_UNLOAD_Timer()
If IsIDE = True Then Timer_FORM_UNLOAD.Enabled = False
If CID_RUN_ME_VAR_TO_EXIT_HAPPEN = True Then Unload Me: Exit Sub

End Sub

Private Sub Timer1_Timer()

Timer1.Enabled = False
Exit Sub

If Timer1.Interval <> 10000 Then Timer1.Interval = 10000

If START_TIME_HANDLE_HIGH_FAULT = 0 Then
    START_TIME_HANDLE_HIGH_FAULT = 1000 + Now + TimeSerial(0, 3, 0)
End If

'If START_TIME_HANDLE_HIGH_FAULT <= Now Then Reload_Me = True: Load Form_UNLOAD_3_HOUR
If START_TIME_HANDLE_HIGH_FAULT <= Now Then Load Form_UNLOAD_3_HOUR


'---------------------------------------------------------------------------------------------
Exit Sub
'---------------------------------------------------------------------------------------------



If Timer1.Interval <> 500 Then Timer1.Interval = 500

    'If CID_Run_Me.Height < 500 Then Call HeightChk

    'LabelX = Val(LabelX) + 1


    'Call PROCESS_KILL_CODE
    Call WINAMP_PROCESS_LICKED_AND_MACHANICS_CODE
    'Call PROCESS_CPU_ADJUST_CODE
    'Call PROCESS_KILL_CODE


    'GRAB CORD'S FOR SYS BAR -----------------------

    'REMARK FROM TOP
    'TELEPORT = 55

    'ww = FindWinPartNotePad
  
    Xxt1 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
    If FindWindow(vbNullString, "Magnifier") > 0 Then
        Xxt1 = FindWindow(vbNullString, "Magnifier")
        wert1 = GetWindowRect(Xxt1, MeRyu4)
        If MeRyu4.Top <> 32 Then
            Xxt1 = FindWindow(vbNullString, "WindowsFormsParkingWindow")
        End If
    End If

    
    
    
    '#---- 01
    xxt2 = FindWindow("Shell_TrayWnd", vbNullString)
    wert1 = GetWindowRect(Xxt1, MeRyu4)
    
    '#---- 02
    Test_HWND = FindWinPartTOP("BaseBar")
    wert1 = GetWindowRect(Test_HWND, MeRect)
    
    
    
    
    'REMARK ------- FROM TOP -------------
    
    If MeRect.Top = 0 Then
        'WITH OFFICE = 0
        TELEPORT = MeRect.Bottom
        TELEPORT = 30
    
    Else
        TELEPORT = 60
    End If
    
    
'--------------------
'--------------------
'--------------------
'--------------------
    
    
    
    

'START --------------



'LAST 50 EXE'S --------------------
CurProcHwnd = GetForegroundWindow
If OCurProcHwnd <> CurProcHwnd Then
    CurProcHwnd_Count = CurProcHwnd_Count + 1
    'CALL AGROSTORESAVE
    Label19 = (str(CurProcHwnd_Count)) + " - Licked"
    
    filexxx$ = GetFileFromHwnd(CurProcHwnd)
    If filexxx$ <> List4.List(0) And Trim(filexxx$) <> "" Then
    If InStr(VMM, filexxx$) = 0 Then
        List4.AddItem filexxx$, 0
        Last50EXEListAPPEND = Last50EXEListAPPEND + vbCrLf + Format(Now, "DD-MM-YYYY HH:MM:SS ") + " -- " + filexxx$
    
    End If
    
    If List4.ListCount > 8000 Then
        List4.RemoveItem (8000)
    End If
    
    On Error Resume Next
    
    Last50EXEListTXT = ""
    
    VMM = ""
    For R = 0 To List4.ListCount - 1
        Last50EXEListTXT = Last50EXEListTXT + vbCrLf + List4.List(R)
        VMM = VMM + List4.List(R)
    Next
    End If
End If





'ONLY ONCE DAY CODE HERE
If DayNow <> Day(Now) Then

    CurProcHwnd_Count_Yester = CurProcHwnd_Count
    CurProcHwnd_Count = 0

    Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\2HwndCountTotal.txt", 1)
    GUG = FreeFile
    Open App.Path + "\0TextData\" + GetComputerName + "\2HwndCountTotal.txt" For Append As #GUG
        Print #GUG, Format$(Int(Now) - 1, "DDD dd-mm-yyyy") + " --" + str(CurProcHwnd_Count_Yester)
    Close #GUG

    'CHK CORRECT NOT SAVED
    'CALL AGROSTORESAVE
End If

OCurProcHwnd = CurProcHwnd
DayNow = Day(Now)








End Sub


Sub PROCESS_CPU_ADJUST_CODE()

Exit Sub

'PROCESS LOWER WITH SHELL PROGRAM

XXRsCNT = XXRsCNT + 1

If XXRsCNT > UBound(XXrXS) Then
    Debug.Print "ERROR MESSAGE HERE IN ROUTINE ---- PROCESS_CPU_ADJUST_CODE ---- ARRAY GOING OVER BOUND VALUE MAXIMUM"
    Exit Sub
End If

'Xxr = FindWindow(vbNullString, "Kat MP3 Recorder.exe - Application Error")
If Xxr > 0 And XXrXS(XXRsCNT) <> Xxr Then
    XXrXS(XXRsCNT) = Xxr
    CloseProgramHwnd (Xxr)
    'If GetForegroundWindow <> Xxr Then SetForegroundWindow Xxr
    'SendKeys "{ENTER}", True
        
    'TTT = cProcesses.Convert(Xxr, Otss, cnFromhWnd Or cnToProcessID)
    'Process_ABOVE_NORMAL_PRIORITY_CLASS (Otss)
    'Process_Low_BELOW_NORM (Otss)
    'Process_Low_Low(Otss)

    Rf = cProcesses.GetEXEID(-1, "C:\Program Files\Kat MP3 Recorder\Kat MP3 Recorder.exe")
    If Rf = 0 Then
        'Shell "C:\Program Files\Kat MP3 Recorder\Kat MP3 Recorder.exe", vbMinimizedNoFocus
    End If

End If



'DICTOR ------------------------

XXRsCNT = XXRsCNT + 1
Xxr = FindWindow(vbNullString, "Transferring the file")
If Xxr > 0 And XXrXS(XXRsCNT) <> Xxr Then
    XXrXS(XXRsCNT) = Xxr

    TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
    Process_ABOVE_NORMAL_PRIORITY_CLASS (OTSS)
    'Process_Low_BELOW_NORM (Otss)
    'Process_Low_Low(Otss)
End If





'PROCESS TO GET FOR SEND KEY ------------------

'FrontPageExplorerWindow40
Xxr = FindWindow("#32770", "Would you like to go offline?")
Xxr = 0
If Xxr > 0 And Xxr25 <> Xxr Then
        Xxr25 = Xxr
        TTT = cProcesses.Convert(Me.hWnd, OTSS, cnFromhWnd Or cnToProcessID)
        On Error Resume Next
        AppActivate (OTSS), False
        SetForegroundWindow Me.hWnd
        TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
        AppActivate (OTSS), True
        SetForegroundWindow Xxr
        If GetForegroundWindow = Xxr Then SendKeys "%(T)", True
        'If Err.Number = 0 Then Xxr25 = Xxr
        On Error GoTo 0
End If




End Sub



Sub PROCESS_KILL_CODE()


'PROCESS  CODE KILL -----------------------


If FFT = 0 Then FFT = Now + TimeSerial(0, 0, 1)

XXRsCNT = XXRsCNT + 1
Xxr = FindWindow(vbNullString, "Norton 360")
If Xxr = 0 Then Xxr = FindWindow("npc.ui.HelpAndSupport", vbNullString)

If Xxr > 0 And FFT < Now Then
    XXrXS(XXRsCNT) = Xxr
    FFT = 0
    
    'CloseProgramHwnd (Xxr)
    'TTT = cProcesses.Convert(Xxr, Otss, cnFromhWnd Or cnToProcessID)
    'Process_Kill (Otss)
    
    'Dim OU As Long
    'OU = -1
    'Rf = cProcesses.GetEXEID(OU, "HSLoader.exe")
    'Process_Kill (OU)
    'ShowWindow Xxr, SW_MINIMIZE
    
End If



'Vodafone Mobile Connect Lite
'Class Name:-ThunderRT6FormDC
Xxr = FindWindow("ThunderRT6FormDC", "Vodafone Mobile Connect Lite")

If Xxr > 0 And Xxr27 <> Xxr Then
Xxr27 = Xxr
Xxr28 = Now + TimeSerial(0, 3, 0)
End If

If Xxr28 > 0 And Xxr28 < Now Then
Xxr28 = 0
TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
Process_Kill (OTSS)
Xxr = FindWindow("PhoneConnectorTTPClient", "TTPClient_dlg_title")
TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
Process_Kill (OTSS)

End If




Xxr = FindWindow(vbNullString, "vbVidCap")
If Xxr = 0 Then Xxr58 = 0
If Xxr > 0 And Xxr57 <> Xxr Then
Xxr57 = Xxr
Xxr58 = Now + TimeSerial(0, 10, 0)
End If

If Xxr58 > 0 And Xxr58 < Now Then
Xxr58 = 0
TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
'TEMP OUT
'Process_Kill (Otss)
End If




'Vodafone Mobile Connect Lite
'Class Name:-ThunderRT6FormDC
Xxr = FindWindow(vbNullString, "Web Site Update Dates")

If Xxr > 0 And Xxr50 <> Xxr Then
Xxr50 = Xxr
Xxr52 = Now + TimeSerial(0, 10, 0)
Else
Xxr52 = 0
End If

If Xxr52 > 0 And Xxr52 < Now Then
    Xxr52 = 0
    Xxr = FindWindow(vbNullString, "Web Site Update Dates")
    TTT = cProcesses.Convert(Xxr, OTSS, cnFromhWnd Or cnToProcessID)
    Process_Kill (OTSS)

End If



End Sub

Sub WINAMP_PROCESS_LICKED_AND_MACHANICS_CODE()


'End of the XxR's

'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------



Dim pCount As Long
Dim wfile2 As String
Dim wfile As String

Call FindCursor(x!, y!)

'Detect Mouse buttons
For i = 1 To 5
    bdf = GetAsyncKeyState(i)
    If bdf <> 0 Then
        Call KeyOrMouse
    End If
Next



'Dim CurProcHwnd As Long
CurProcHwnd = GetForegroundWindow

If CurProcHwnd = 0 Then Exit Sub

If LockInToMp3 = True Then
    ReRun = 1
'    Chit22 = 1
End If

Dim Peet2 As Long
Dim Peet4 As Long
  
    
    
    
'MACHANICS CODE ----------------------
    

'If Tagad > 0 And ReRun = 0 And Curprochwnd = Me.hWnd Then CurProchwnd2 = Me.hWnd

Peet3$ = ""
Peet33$ = ""
If List2.ListCount > 0 Then
    For R = List2.ListCount - 1 To 0 Step -1
        peet1$ = List2.List(R)
        Peet2 = Val(Mid$(peet1$, 1, 9))
        Peet4 = Val(Mid$(peet1$, 11, 8))
        If GetWindow(Peet2, GW_HWNDFIRST) = 0 Then
            List2.RemoveItem R
            ReRun = 1
        Else
            Peet3$ = Peet3$ + str$(Peet2)
            Peet33$ = Peet33$ + str$(Peet4)
        End If
    Next
End If


  
        
If CurProcHwnd <> Me.hWnd And CurProcHwnd <> FindWindow(vbNullString, "WMI Demo - CPU Information") Then
    LastCurProcHwnd = CurProcHwnd
End If
        
        

    If InStr(Peet3$, str(CurProcHwnd)) = 0 Then
        totss = cProcesses.Convert(CurProcHwnd, OTSS, cnFromhWnd Or cnToProcessID)
        List2.AddItem Format$(CurProcHwnd, "000000000") + " " + Format$(OTSS, "0000000") + " "
        Peet3$ = Peet3$ + str$(CurProcHwnd)
        Peet33$ = Peet33$ + str$(OTSS)
    End If



    FlingGrater1$ = ""
    FlingGrater2$ = ""

    NoMonOff2 = 0
    NoMonOff = 0
    NoMusic = 0
    frontpagepid = 0
    ProcessId25Ssam = 0

    Dim lhTmp As Long
    
    Window_Title_String = ""
        
    List1.Clear
    'RTB1.Text = ""
    TTg$ = ""
        
    Chit24 = 0
        
    If LockInToMp3 = False Then Chit22 = 0
        
    If LastCurProcHwnd = 0 Then LastCurProcHwnd = Me.hWnd
    
    lhTmp = GetWindow(LastCurProcHwnd, GW_HWNDFIRST)
    Window_Title_String = GetWindowTitle(LastCurProcHwnd)
    wfile2 = GetFileFromHwnd(LastCurProcHwnd)
            
    Dim Ret As Long, sText As String
    sText = Space(255)
    Ret = GetClassName(LastCurProcHwnd, sText, 255)
    sText = Left$(sText, Ret)
    HwndDd = GetWindowRect(LastCurProcHwnd, Rt2)
    gws = GetWindowState(LastCurProcHwnd)
    If gws = -1 Then gws2$ = "Window State Normal"
    If gws = vbMinimized Then gws2$ = "Window State Minimized"
    If gws = vbMaximized Then gws2$ = "Window State Maximized"

    
    List1.AddItem "Last Current Process --------"
    List1.AddItem "--------------------------------------------"
    
    List1.AddItem "Title Name:-" + Window_Title_String
    List1.AddItem "Class Name:-" + sText
    List1.AddItem "Hwnd:--" + str(Peet2)
    LastCurProc5 = Peet2
    List1.AddItem "PID :--" + str(Peet4)
    List1.AddItem "X=" + str(Rt2.Left) + " Y=" + str(Rt2.Top) + " W=" + str(Rt2.Right - Rt2.Left) + " H=" + str(Rt2.Bottom - Rt2.Top)
    List1.AddItem gws2$
    List1.AddItem "File Name :--"
    List1.AddItem wfile2
    List1.AddItem "----------------------------------------"
                
     
    
    
    
    
    
    
    
    
    For R = 0 To List2.ListCount - 1
        
        peet1$ = List2.List(R)
        Peet2 = Val(Mid$(peet1$, 1, 9))
        Peet4 = Val(Mid$(peet1$, 11, 8))
            
        lhTmp = GetWindow(Peet2, GW_HWNDFIRST)
        Window_Title_String = GetWindowTitle(Peet2)
           
        If CurProcHwnd = Peet2 Then
            lsidx = R
        End If
            
            
            wfile2 = GetFileFromHwnd(Peet2)
            
            HwndDd = GetWindowRect(Peet2, Rt2)
            
            gws = GetWindowState(Peet2)
            If gws = -1 Then gws2$ = "Window State Normal"
            If gws = vbMinimized Then gws2$ = "Window State Minimized"
            If gws = vbMaximized Then gws2$ = "Window State Maximized"
            
            sText = Space(255)
            Ret = GetClassName(Peet2, sText, 255)
            sText = Left$(sText, Ret)
            List1.AddItem "Title Name:-" + Window_Title_String
            List1.AddItem "Class Name:-" + sText
            List1.AddItem "Hwnd:--" + str(Peet2)
            List1.AddItem "PID :--" + str(Peet4)
            List1.AddItem "X=" + str(Rt2.Left) + " Y=" + str(Rt2.Top) + " W=" + str(Rt2.Right - Rt2.Left) + " H=" + str(Rt2.Bottom - Rt2.Top)
            List1.AddItem gws2$
            List1.AddItem "File Name :--"
            List1.AddItem wfile2
            List1.AddItem "-------------------------"
                
            OTSS = Peet4
            
            Call Perfect(wfile2, 1)
            
            If NoMusic = 1 And OldNoMusic = 0 Then FlingGrater1$ = wfile2
            If NoMonOff = 1 And OldNoMonOff = 0 Then FlingGrater2$ = wfile2
            
            OldNoMusic = NoMusic
            OldNoMonOff = NoMonOff
            
            If InStr(LCase(wfile2), "amp caller\winamp.exe") Then
                ProcessId22 = OTSS
                Chit22 = 1
                LockInToMp3 = False
            End If
                
            If InStr(LCase(wfile2), "winamp.exe") Then
                ProcessId24 = OTSS
                Chit24 = 1
            End If

            
                
        Next
    
        
        If (NoMusic = 1 And Cmdv = 1) Then NoMusic = 0
    
        If Ebuy = 0 Then Ebuyer = 0
    
        NoMonOffX = NoMonOff
        NoMusicX = NoMusic

        Call proc24
        
        List1.AddItem "Special Processes........."
        
        If ProcessId22 > 0 Then
            Window_Title_String = GetWindowTitle(Winamp22Handle2)
            If Window_Title_String = "" Then ProcessId22 = 0
        End If
        
        If ProcessId22 > 0 Then
            If InStr(Peet3$, str(Winamp22Handle2)) = 0 Then
                List2.AddItem Format$(Winamp22Handle2, "000000000") + " " + Format$(ProcessId22, "0000000") + " "
                Peet3$ = Peet3$ + str(Winamp22Handle2)
                Peet33$ = Peet33$ + str$(ProcessId22)
            End If
            wfile2 = GetFileFromHwnd(Winamp22Handle2)
            profile22$ = wfile2
            sText = Space(255)
            Ret = GetClassName(Winamp22Handle2, sText, 255)
            sText = Left$(sText, Ret)
            List1.AddItem "WinAmp Caller ID......"
            List1.AddItem profile22$
            List1.AddItem "Title Name--" + Window_Title_String
            List1.AddItem "Class Name--" + sText
            List1.AddItem "Hwnd:--" + str(Winamp22Handle2)
            List1.AddItem "PID :--" + str(ProcessId22)
            List1.AddItem "-------------------------"
            LockInToMp3 = False
            Chit22 = 1
        End If
        
        If ProcessId24 > 0 Then
            Window_Title_String = GetWindowTitle(Winamp24Handle2)
            If Window_Title_String = "" Then ProcessId24 = 0
        End If

        If ProcessId24 > 0 Then
            If InStr(Peet3$, str(Winamp24Handle2)) = 0 Then
                List2.AddItem Format$(Winamp24Handle2, "000000000") + " " + Format$(ProcessId24, "0000000") + " "
                Peet3$ = Peet3$ + str(Winamp24Handle2)
                Peet33$ = Peet33$ + str$(ProcessId24)
            End If
            wfile2 = GetFileFromHwnd(Winamp24Handle2)

            profile24$ = wfile2
            sText = Space(255)
            Ret = GetClassName(Winamp24Handle2, sText, 255)
            sText = Left$(sText, Ret)
            List1.AddItem "WinAmp Gold Amp......."
            List1.AddItem profile24$
            List1.AddItem "Title Name--" + Window_Title_String
            List1.AddItem "Class Name--" + sText
            List1.AddItem "Hwnd:--" + str(Winamp24Handle2)
            List1.AddItem "PID :--" + str(ProcessId24)
            List1.AddItem "-------------------------"
            Chit24 = 1
        End If
        
        ReRun = 0
    
        For R1 = 0 To List1.ListCount - 1
            TTg$ = TTg$ + List1.List(R1) + vbCr
            'RTB1.Text = RTB1.Text + List1.List(R1) + vbCr
        Next
        
        If FlingGrater1$ <> "" Then
            TTg$ = TTg$ + "NoMusic  := " + FlingGrater1$ + vbCr
            'RTB1.Text = RTB1.Text + "NoMusic  := " + FlingGrater1$ + vbCr
        End If
        If FlingGrater2$ <> "" Then
            TTg$ = TTg$ + "NoMonOff := " + FlingGrater2$ + vbCr
            'RTB1.Text = RTB1.Text + "NoMonOff := " + FlingGrater2$ + vbCr
        End If

    CurProcHwnd2 = CurProcHwnd
    
    If TTg$ <> TTx$ Then
    RTB1.Text = TTg$
    RTB1.SelStart = 0
    RTB1.SelLength = Len(RTB1.Text)

    RTB1.SelColor = &H80FFFF
    RTB1.SelStart = 0
    RTB1.SelLength = 0
    
    
    TTx$ = TTg$
    
    'Call Label15_Click
    End If
'End If



If 1 = 2 Then
For iCtr = 0 To UBound(sArray)
    Skip2 = FindWindow(vbNullString, sArray(iCtr))
    If Skip2 > 0 Then Exit For
Next

'ANCIENT XCOPY DOS CODE ----------------
If Skip2 = 0 Then Skip2 = FindWindow(vbNullString, "Let The Backup Comence...")

If RASCount() Or Skip2 > 0 Then
    NoMusic = 1: NoMonOff = 1
    NoMusicX = 1: NoMonOffX = 1
End If
End If


FirstRun2 = True

If Chit22 = 0 Then
'    ProcessId22 = 0
End If

'If Chit24 = 0 Then ProcessId24 = 0

w14 = GetWindow(Winamp22Handle2, GW_HWNDFIRST)
w15 = GetWindow(Winamp24Handle2, GW_HWNDFIRST)
If w14 < 1 And Winamp22Handle2 > 0 Then
    Winamp22Handle2 = 0: ProcessId22 = 0: ReRun = 1
End If
If w15 < 1 And Winamp24Handle2 > 0 Then
    Winamp24Handle2 = 0: ProcessId24 = 0: ReRun = 1
End If

If ProcessId22 > 0 Then wedf$ = "+" Else wedf$ = ""
If ProcessId24 > 0 Then wedg$ = "@" Else wedg$ = ""

If Label15.Caption <> Trim(str(List2.ListCount)) + wedf$ + wedg$ Then
    Label15.Caption = (str(List2.ListCount)) + wedf$ + wedg$
End If

Call proc24



End Sub


Private Sub Timer10_Timer()
Timer10.Enabled = False
Exit Sub

On Local Error Resume Next
Xxr = FindWindow(vbNullString, "Error Connecting to 00 W880i")
If Xxr > 0 And Xxr21 < Now Then
    Xxr21 = Now + TimeSerial(0, 0, 4)
    AppActivate "Error Connecting to 00 W880i", False
    SendKeys "{ESC}", False
End If

End Sub

Private Sub Timer11_Timer()

'TIMER11.INTERVAL

If Timer11.Interval <> 1000 Then Timer11.Interval = 1000
If FindWindow(vbNullString, "Tidal Information...") = 0 Then TOG1 = 1 Else TOG1 = 0
If FindWindow(vbNullString, "Extreme Lock Build: 2011") = 0 And TOG1 = 1 Then Exit Sub

On Local Error GoTo Errend

If RamDrive <> "" Then
    frg = FreeFile
    Open RamDrive + ":\temp\KeyHit-Tidal.txt" For Input Lock Write As #frg
        Line Input #frg, kb$
    Close #frg
End If


KeyHitts = Val(kb$)
If KeyHittsOld = 0 Then
    KeyHittsOld = KeyHitts - 1
    If KeyHittsOld < 0 Then KeyHittsOld = 0
End If
If KeyHittsOld > KeyHitts Then KeyHittsOld = 0

If KeyHittsOld = KeyHitts Then Exit Sub

Keyy = Keyy + (KeyHitts - KeyHittsOld)
Daykeyy = Daykeyy + (KeyHitts - KeyHittsOld)

If Daykeyy + Daymouse <> OKM Then
Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCount.txt", 1)
GUG = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCount.txt" For Append As #GUG
Print #GUG, "Day Counter  = "; Now
Print #GUG, "Day Mouse    = "; Daymouse
Print #GUG, "Day Keyboard = "; Daykeyy
Close #GUG
End If
OKM = Daykeyy + Daymouse

KeyHittsOld = KeyHitts


Kbbdf = True
KBPress2 = vcode
LastPressTrig = True
LastPressTrigTimer.Enabled = True
'Call CID_Run_Me.Lastpress

Errend:
Close #frg
End Sub

Private Sub Timer12_Timer()

'TIMER12.INTERVAL = 20 ' DEFAULT
'HANDLE HOOVER GOING OFF
Timer12.Interval = 0
Timer12.Enabled = False
Exit Sub



FindCursor2 Nx, Ny
mWnd = WindowFromPoint(Nx, Ny)
If OCurWinHovHwnd = mWnd Then Exit Sub

If DayNow2 <> Day(Now) Then
    CurWinHovHwnd_Count_Yester = CurWinHovHwnd_Count
    CurWinHovHwnd_Count = 0
    
    
    
    HWND_HOV_TOTAL = HWND_HOV_TOTAL + vbCrLf + Format$(Int(Now), "DDD dd-mm-yyyy") + " --" + str(CurWinHovHwnd_Count_Yester)
    
    Call AGROSTORESAVE
End If

CurWinHovHwnd_Count = CurWinHovHwnd_Count + 1

Label22 = (str(CurWinHovHwnd_Count)) + " - Hoovered"

OCurWinHovHwnd = mWnd
DayNow2 = Day(Now)

'CALL AGROSTORESAVE

End Sub


Private Sub Timer13_Timer()
Timer13.Enabled = False
'--------------------------
Exit Sub

TFX = FindWindow("Winamp v1.x", vbNullString)
If TFX <> OTFX Then
    OUREXEWINAMP = GetFileFromHwnd(TFX)
    If InStr(LCase(OUREXEWINAMP), LCase("Winamp Gold 07 Video X\winamp.exe")) > 0 Then
        OUREXEWINAMP = True
    Else
        OUREXEWINAMP = False
    End If
End If
OTFX = TFX

End Sub



Private Sub Timer15_Timer()

If SunSet = 0 Then Timer15.Enabled = False: Exit Sub
'A = Timer15.Interval 100 DEFAULT
'INTERVAL TOO HIGH
If Timer15.Interval <> 50000 Then Timer15.Interval = 50000

If OLD_SunSet = SunSet Then Exit Sub

OLD_SunSet = SunSet

If (Now - Int(Now) < SunSet And Now - Int(Now) > SunRise) = False Then
    If AK47 = False Then
        Label38.Caption = "Web Cam <Dark>"
        Timer3.Enabled = False
        Timer14Var = Now + TimeSerial(0, 5, 0)
    End If
    AK47 = True: Timer14Var = Now + TimeSerial(0, 5, 0)
Exit Sub
End If



Timer3.Enabled = True
AK47 = False
Xxr = FindWindow(vbNullString, "WebCam Motion Capture")

If Xxr = 0 Then Xxp2 = FindWindow(vbNullString, "WebCam Motion Capture In IDE")

If Xxr > 0 Or Xxp2 > 0 Then
    LastWeb = Now: Timer14Var = Now + TimeSerial(0, 4, 0)
End If

If LastWeb = 0 Then LastWeb = Now

'If DateDiff("n", LastWeb, Now) >= 5 And Timer14.Enabled = False Then
'    Timer14.Enabled = True
'End If

End Sub

Private Sub Timer16_Timer()

'-----------------------------------------------------
'ONLY ONCE LATCH DO SOON AS POSS
'-----------------------------------------------------
'1ST TIMER TO RUN AWAY FORM FROM_LOAD AS DELAY ROUTINE
'-----------------------------------------------------
If TIMER8_LEFT_THE_FORM_LOAD_ENTRY_YET <> 2 Then
    Call Timer8_Timer
End If
'-----------------------------------------------------

'Schedule_Timer - Microsoft Visual Basic [break] - [Form1 (Code)]

'TIMER16.INTERVAL === 2000

Timer16.Interval = 0
RR = FindWinPartCodeBreak("Visual Basic [break]")
If RR <> "" Then
    i = MsgBox("BREAK IN VB CODE AT DETECTED BY ANOTHER CODE PROGRAM " + vbCrLf + Format(Now, "DDD DD-MMM HH:MM:SSa/p") + vbCrLf + RR + "CID RUN ME ALERTED THIS" + vbCrLf + "STOP AT CODE HERE INSPECT MSGBOX", vbYesNo + vbMsgBoxSetForeground)
    If i = vbYes Then
        If IsIDE = False Then
            MsgBox "STOP BREAK POINT ONLY HAPPEN IN ISIDE=TRUE __ CONTINUE ANYHOW"
        Else
            Stop
        End If
    End If
End If
'MsgBox "BREAK IN VB CODE AT " + vbCrLf + Format(Now, "DDD DD-MMM HH:MM:SSa/p") + vbCrLf + rr, vbCritical, "CID RUN ME ALERTED THIS"

End Sub

Private Sub Timer2_Timer()

'If IsIDE = True Then Timer2.Enabled = False

'-----------
'MOVING FORM
'-----------

'GoTo Skip5

'------------------------
'A = Timer2.Interval = 50
'DEFAULT BIT FREQ 50
'------------------------
If Timer2.Interval <> 1000 Then Timer2.Interval = 1000

If CID_Run_Me.Top + CID_Run_Me.Height + HeightOffSet > ScreenHeightY Or _
    OptionsMenu.Top + OptionsMenu.Height + HeightOffSet > ScreenHeightY Then
    If XRad <> CID_Run_Me.Top + OptionsMenu.Top Then
        If GetAsyncKeyState(1) < 0 Or 1 = 1 Then
            'PeaShoot = Timer + 0.4
            OCIDLeft = 0
        End If
    End If
End If


XRad = CID_Run_Me.Top + OptionsMenu.Top

If LockMove = 1 Then OCIDLeft = CID_Run_Me.Left

If CID_Run_Me.Left > ScreenWidthX - CID_Run_Me.Width Then
    LockMove = 0
    OCIDLeft = 0
End If

If GetAsyncKeyState(1) Then
    LockMove = 1
    OCIDLeft = CID_Run_Me.Left
End If


If (OCIDLeft <> CID_Run_Me.Left Or OldWState <> CID_Run_Me.WindowState) And LockMove = 0 Then

'If (PeaShoot < Timer And PeaShoot > 0) Or PeaShoot - 100 > Timer Or (OldWState <> CID_Run_Me.WindowState) And lockmove2 = 0 Then
    PeaShoot = 0
    
    If CID_Run_Me.WindowState = 0 Then
        'CID_Run_Me.Left = ScreenWidthX - CID_Run_Me.Width - OffSetGoogle
        Me.Left = (Screen.Width - Me.Width)
        xxt2 = FindWindow("Shell_TrayWnd", vbNullString)
        wert2 = GetWindowRect(xxt2, MeRyu5)
        Me.Top = (MeRyu5.Top * Screen.TwipsPerPixelY) - Me.Height '(Screen.Height - Me.Height) - 400



        'CID_Run_Me.Top = ScreenHeightY - (CID_Run_Me.Height + HeightOffSet)

        rt = FindWindow("Shell_TrayWnd", vbNullString)
        If rt > 0 Then
            HwndDd = GetWindowRect(rt, Rt2)
        End If
        'CID_Run_Me.Top = (Rt2.Top * ScreenTwipsY) - CID_Run_Me.Height




        CID_Run_Me.Refresh

        OptionsMenu.Width = CID_Run_Me.Width
        OptionsMenu.Top = CID_Run_Me.Top - OptionsMenu.Height
        OptionsMenu.Left = ScreenWidthX - OptionsMenu.Width - OffSetGoogle
    End If
    
    If OptionsMenu.WindowState = 0 And CID_Run_Me.WindowState <> 0 Then
        'OptionsMenu.Width = CID_Run_Me.Width
        OptionsMenu.Top = ScreenHeightY - (HeightOffSet) - OptionsMenu.Height
        OptionsMenu.Left = ScreenWidthX - OptionsMenu.Width - OffSetGoogle
    End If

OCIDLeft = CID_Run_Me.Left

End If


'OldXRad = XRad
OldWState = CID_Run_Me.WindowState




End Sub



Private Sub Timer3_Timer()
    
    If IsIDE = True Then Timer3.Enabled = False
    
    wd = DateDiff("s", LastWeb, Now)
    f6 = Int(wd / 3600)
    f7 = Int(wd / 60)
    f8 = wd Mod 60
    lab2 = Format$(TimeSerial(0, f7, f8), "HH:MM:SS")
    If Mid$(lab2, 1, 4) = "00:0" Then lab2 = (Mid(lab2, 4))
    If Mid$(lab2, 1, 3) = "00:" Then lab2 = (Mid(lab2, 4))
    If wd < 60 Then piol = " Sec"
    Label38 = "Web Cam " + lab2 + piol
    
    Exit Sub
    
    
    A1 = DateDiff("s", LastWeb, Now)
    Select Case A1
        Case Is < 60
            Label38 = "Web Cam " + Format(DateDiff("s", LastWeb, Now), "0") + " Secs Ago"
        Exit Sub
    End Select
    
    A1 = DateDiff("n", LastWeb, Now)
    Select Case A1
        Case Is < 60
            Label38 = "Web Cam " + Format(DateDiff("n", LastWeb, Now), "0") + " Mins Ago"
        Exit Sub
    End Select
  
    A1 = DateDiff("h", LastWeb, Now)
    Select Case A1
        Case Is < 999999
            Label38 = "Web Cam " + Format(DateDiff("h", LastWeb, Now), "0") + " Hrs Ago"
        Exit Sub
    End Select
  
End Sub

Private Sub Timer14_Timer()
    Timer14.Enabled = False
    Exit Sub

    If Timer14Var > Now Or Timer14Var = 0 Then Exit Sub
    
    MsgBox "Reminder: The Web Cam Has Not Ran For" + str(DateDiff("n", LastWeb, Now)) + " Minutes Now"
    Timer14Var = Now + TimeSerial(0, 8, 0)
End Sub

Private Sub Timer4_Timer()
Timer4.Enabled = False
Exit Sub


If Xxp2 = 0 Then
    'MsgBox ("Web Cam Was Killed after 1 Min")
Else
    'MsgBox ("The VB6 Run Mode Web Cam IS Running Still After 1 Min")
End If


For i = 5 To 255
    
    bdf = GetAsyncKeyState(i)
 
    If bdf <> 0 Then
        Kbbdf = True
        Call CID_Run_Me.Lastpress
    End If
Next

End Sub

Private Sub Timer6_Timer()
Timer6.Enabled = False
'-------------------------
Exit Sub

'Private Declare Function GetAsyncKeyState Lib "user32" _
        (ByVal vKey As Long) As Integer

GoTo Skip2

bdfh2 = GetAsyncKeyState(1)
If bdfh2 <> 0 Then
    CurProcHwnd3 = GetForegroundWindow
    If Winamp24Handle2PL = CurProcHwnd3 Then
        Agust2 = Now + TimeSerial(0, 0, 2)
    End If
End If

Skip2:

tf1 = FindWindow("Winamp PE", vbNullString)
TF2 = FindWindow("Winamp v1.x", vbNullString)
If tf1 = 0 Then Exit Sub
If tf1 > 0 Then

    wfile2 = GetFileFromHwnd(tf1)
End If
gg1 = InStrRev(wfile2, "\")
ggh$ = Mid$(wfile2, 1, gg1)

On Local Error GoTo errhand

Do
Err.Clear
xer = 0
Xx$ = ggh$
TT1$ = Dir$(ggh$ + "Current Play To Text File.txt")
Tt2$ = ggh$ + "Current Play To Text File Append.txt"
tty3$ = ggh$ + "Current Play To Text File -One.txt"

If TT1$ <> "" Then

'TimerFileSave.Enabled = True

    Call FileInUseDelay(Xx$ + TT1$, 0)
    F4 = FreeFile
    Open Xx$ + TT1$ For Input As #F4
    Call FileInUseDelay(Tt2$, 1)
    f5 = FreeFile
    Open Tt2$ For Append As #f5
    Print #f5,
    Print #f5, Format$(Now, "DDD DD-MMM-YY HH:MM:SS ap")
    Print #f5, "--- -- --- -- -- -- -- -"
    Do
        'Err.Clear
'        On Local Error Resume Next
        Line Input #F4, Gg$
        'xer = Err.Number
        Print #f5, Gg$
    Loop Until EOF(F4)
    Close #F4
    Print #f5,
    Close #f5



'MMControl1.Command = "Prev"
'MMControl1.Command = "Play"


'''"D:\Wave's\Cosmi\CANNON.WAV"
If InStr(Gg$, " - Winamp") > 0 Then
    Call Frm_WinAmpDisplay.DisplayWinamp
End If

On Local Error Resume Next
Kill Xx$ + TT1$


End If

Loop Until xer = 0

Exit Sub
errhand:
DoEvents

'Resume

End Sub


Public Sub proc22()
Exit Sub

'Caller Amp
If (ProcessId22 <> oldprocessid22 And ProcessId22 > 0) Or PlimBurn1 = False Then
    Winamp22Handle2 = hWndFromProcID(ProcessId22)
    ReRun = 1
End If

If ProcessId22 > 0 And Winamp22Handle2 < 1 Then
    PlimBurn1 = False
Else
    PlimBurn1 = True
End If

MsgResult = SendMessage(Winamp22Handle2, WM_USER, 0, ByVal ipc_isplaying)

'// IPC_ISPLAYING returns the status of playback.
'// If it returns 1, it is playing. if it returns 3, it is paused,
'// if it returns 0, it is not playing.

If MsgResult = 1 Then
    If XtX1 = -1 Then XtX1 = 0
Else
    XtX1 = -1
End If

oldprocessid22 = ProcessId22

End Sub



Public Sub proc24()

'Stop

'Gold Amp
Call proc22
    
'    winamp24handle2 = hWndFromProcID(ProcessId24)
If (ProcessId24 <> oldprocessid24 And ProcessId24 > 0) Or PlimBurn2 = False Then
    tagotd = 1
    Winamp24Handle2 = hWndFromProcID(ProcessId24)
    Winamp24Handle2PL = hWndFromProcIDPlayList(ProcessId24)
End If

If ProcessId24 > 0 And Winamp24Handle2 < 1 Then
    PlimBurn2 = False
Else
    PlimBurn2 = True
End If

oldprocessid24 = ProcessId24




'// IPC_ISPLAYING returns the status of playback.
'// If it returns 1, it is playing. if it returns 3, it is paused,
'// if it returns 0, it is not playing.

On Local Error GoTo trap1

If MsgResult = 1 And EmSQ = 1 Then
        
        
    Window_Title_String = GetWindowTitle(Winamp24Handle2)

    If Trim(Window_Title_String) <> "" Then
        pdq$ = "- Winamp"
        'If InStr(Window_Title_String, pdq$) = 0 Then pdq$ = "Winamp 5.094": Exit Sub
        Window_Title_String = Trim(Left(Window_Title_String, InStr(Window_Title_String, pdq$) - 1))
    End If
    EmSQ = 0
    
    Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\WinAmpLog\Winamp-list-GoldAmp.txt", 1)
    freef5 = FreeFile
    Open App.Path + "\0TextData\" + GetComputerName + "\WinAmpLog\Winamp-list-GoldAmp.txt" For Append As #freef5
    If Trim(Window_Title_String) <> "" Then Print #freef5, Format$(Now, "DD/MM/YYYY HH:MM:SS ampm ") + "[Play ] -- " + Window_Title_String
    Close #freef5


    
    'TimerFileSave2.Enabled = True
End If
On Local Error GoTo 0


If MsgResult = 1 Then
    NoMonOff = 1
    NoMusic = 1
    If XtX2 = -1 Then XtX2 = 0
Else
    EmSQ = 1
    NoMonOff = NoMonOffX
    NoMusic = NoMusicX
    Agust = Now - 4
    XtX2 = -1
End If

If Rem5 <> MsgResult Then
    Rem5 = MsgResult
    Call Timer2_Timer
End If

If tagotd = 1 Then ReRun = 1


Exit Sub

trap1:

DoEvents

'Resume

End Sub



Public Sub FindCursor2(x, y)
Dim P As POINTAPI
GetCursorPos P
x = P.x
y = P.y
End Sub

Public Sub FindCursor(ByRef x!, ByRef y!)

Dim P As POINTAPI

GetCursorPos P

x = P.x

y = P.y

If x <> 1 And y <> 1 And DbonPop = 1 And Dbon2 = 0 Then
    Dbon = Timer + 0.1
    Dbon2 = 1
End If

If x = 1 And y = 1 Then
    Dbon = Timer + 0.7: DbonPop = 1
End If

If Dbon - 100 > Timer Then
    Dbon = Timer
End If

If FirstRun = True Then
    Xmp1 = x
    Ymp1 = y
End If

If DbonPop = 1 And Dbon < Timer Then
    Dbon2 = 0
    DbonPop = 0
    Xmp1 = x
    Ymp1 = y
End If

If (x <> Xmp1 Or y <> Ymp1) And Dbon < Timer Then
    
    WRNow = Now + TimeSerial(0, 0, 3) 'winamp dont stop if mouse
    
    
    Idle_Timer = Now
    
    Mousey = Mousey + Sqr(Abs(Xmp1 - x) ^ 2 + Abs(Ymp1 - y) ^ 2)
    Daymouse = Daymouse + Sqr(Abs(Xmp1 - x) ^ 2 + Abs(Ymp1 - y) ^ 2)
    
    Xmarks = Now - 1
    'MattsTimer2 = Now + MinorOnTime
    
    Call Mystery
     
    If Mousey3 <> Now Then
        frmSender_CallID.txtMsg.Text = "MouseCounter " + str$(Mouse55 + Mousey)
        'Call frmSender_CallID.cmdSend_Click
        Mousey3 = Now
    End If
    
    If Mousey > 0 Then
        
        Call ClingDing
        
    End If

    Call KeyOrMouse

End If

Xmp1 = P.x

Ymp1 = P.y

End Sub

Sub KeyOrMouse()

MattsTimer2 = Now + MinorOnTime
WinampTimerFileName = Now

TF2 = FindWindow("Winamp v1.x", vbNullString)
MsgResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)

If MsgResult = 1 Then
    WinStopTimer = Now + WinStopTimerTot
    
    '---------------------------
    'in operation 2017
    '---------------------------
    'Debug.Print WinStopTimerTot
End If
End Sub

Public Sub ClingDing()

'this is our usual fav one
If CheckT1$ = "3" Then
    Label21.Caption = Trim(str$(Daymouse))
    Label23.Caption = Trim(str$(Daykeyy))
End If



If TotalDayMouse > 0 Then
    Label21.Caption = Trim(str$(TotalDayMouse))
    Label23.Caption = Trim(str$(TotalDayKeyy))
    Label21.BackColor = QBColor(12)
    Label23.BackColor = QBColor(12)
    Exit Sub
End If


If CheckT1$ = "1" Then
    Label21.Caption = Trim(str$(Mousey))
    Label23.Caption = Trim(str$(Keyy))
End If
        
If CheckT1$ = "2" Then
    Label21.Caption = Trim(str$(Mouse55 + Mousey))
    Label23.Caption = Trim(str$(Qwerty2 + Keyy))
End If
        




End Sub



Public Sub dss_download()

Dim nMaxCount As Long
Dim lpClassName As String
Dim hand3 As Long
Dim cch As Long
Dim lpString As String
  
handle2 = FindWindow(vbNullString, "Download")

filehwnd2$ = ""
If handle2 > 0 Then filehwnd2$ = GetFileFromHwnd(handle2)

If filehwnd2$ = "R:\Program Files\Olympus\DSSPlayer2002\DictWnd.exe" Then
    
    On Local Error Resume Next
    AppActivate "Download", False
    On Local Error GoTo 0

    Ew = Now + TimeSerial(0, 0, 2)
    Do
        DoEvents
        Window_Title_String = GetActiveWindowTitle(False)
    Loop Until Ew < Now Or Window_Title_String = "Download"


    SendKeys "{enter}", False

End If

End Sub




Public Sub Bashing()

If Command2.BackColor = QBColor(12) Then Exit Sub
If Menu_StopCurrent.Checked = False Then Exit Sub
If ProcessId24 > 0 Then
    
    MsgResult = SendMessage(Winamp24Handle2, WM_USER, 0, ByVal ipc_isplaying)

    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.

    'if playing stop plyaing it
    If MsgResult = 1 Then
     
        MsgResult = SendMessage(Winamp24Handle2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
        
    End If
End If
        
End Sub

Public Sub winamp_player_off()

'If Music_Off_Timer > 0 Then Exit Sub
    
If ProcessId22 > 0 Then
    
    MsgResult = SendMessage(Winamp22Handle2, WM_USER, 0, ByVal ipc_isplaying)

    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.

    If MsgResult = 1 Then
     
        MsgResult = SendMessage(Winamp22Handle2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
        
    End If
End If
        
End Sub

Function RASCount() As Integer
  Dim lpRasConn(0 To 1) As Long ' dummy buffer area
  Dim rc As Long ' return code
  Dim lpcb As Long ' buffer size
  Dim lpcConnections As Long ' connection count

  lpRasConn(0) = 32 ' each returned item is at least 32 bytes long
  lpcb = 0 ' set total number of usable bytes in the buffer to zero
  
  rc = RasEnumConnections(lpRasConn(0), lpcb, lpcConnections)
  RASCount = lpcConnections ' return connection count

End Function

Public Sub DUN_Services(DUN_Array() As String)
Exit Sub

'Pass in Empty array for DUN_Array
    Dim S As Long, Ln As Long, conname As String, i As Long
    Dim R(255) As RASENTRYNAME95
    R(0).dwSize = 264
    S = 256 * R(0).dwSize
    
    Call RasEnumEntriesA(vbNullString, vbNullString, R(0), S, Ln)
    Ln = Ln - 1
    ReDim DUN_Array(200)
    
    pigbum = Ln
    
    For i = 0 To Ln
        conname = StrConv(R(i).szEntryName(), vbUnicode)
        If InStr(conname, "Broadband") > 0 Or _
        InStr(conname, "Bluetooth") > 0 Then
            For i2 = i + 1 To Ln
                R(i2 - 1) = R(i2)
            Next
        pigbum = pigbum - 1
        End If
    Next
    
    Ln = pigbum
    
    For i = 0 To Ln
        conname = StrConv(R(i).szEntryName(), vbUnicode)
        DUN_Array(i) = Left$(conname, InStr(conname, _
          vbNullChar) - 1)
    Next i
    
    pt = i
    For i = 0 To Ln
        conname = StrConv(R(i).szEntryName(), vbUnicode)
        DUN_Array(i + pt - 1) = "Connect " + Left$(conname, InStr(conname, _
          vbNullChar) - 1)
    Next i
    
    pt = pt + i
    For i = 0 To Ln
        conname = StrConv(R(i).szEntryName(), vbUnicode)
        DUN_Array(i + pt - 1) = "Connecting " + Left$(conname, InStr(conname, _
          vbNullChar) - 1) + "..."
    Next i

    pt = pt + i
    For i = 0 To Ln
        conname = StrConv(R(i).szEntryName(), vbUnicode)
        DUN_Array(i + pt - 1) = "Connecting to " + Left$(conname, InStr(conname, _
          vbNullChar) - 1)
    Next i



'Skip2 = FindWindow(vbNullString, "Connect BTOW - Surftime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connect BTOW - Daytime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting BTOW - Daytime...")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting BTOW - Surftime...")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting to BTOW - Daytime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting to BTOW - Surftime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Dial-up Connection")
    
    
    
    DUN_Array(pt + i - 1) = "Dial-up Connection"


    ReDim Preserve DUN_Array(pt + i - 1)




End Sub


Public Sub Winamp_Player_On()

aw$ = globalhot$

   'ras listesini yuklesi baslar
   


If NoMusic <> 0 Or (Music_Off_Timer > Now And Music_Off_Timer > 0) Then Exit Sub
        


On Local Error Resume Next
    For iCtr = 0 To UBound(sArray)
        Skip2 = FindWindow(vbNullString, sArray(iCtr))
        If Skip2 > 0 Then Exit For
    Next
On Local Error GoTo 0


'Skip2 = FindWindow(vbNullString, "Connect BTOW - Surftime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connect BTOW - Daytime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting BTOW - Daytime...")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting BTOW - Surftime...")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting to BTOW - Daytime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Connecting to BTOW - Surftime")
'Skip2 = Skip2 + FindWindow(vbNullString, "Dial-up Connection")
     
     
     
     
     
     
     
If Skip2 > 0 Then NoMusic = 1

If ProcessId22 = 0 Then
    
    If LockInToMp3 = False Then
        ProcessId22 = Shell("C:\program files\winamp caller\winamp.exe", vbMinimizedNoFocus)
    End If
    Chit22 = 1
    ReRun = 1
    LockInToMp3 = True
    Call proc22
    
Else

    MsgResult = SendMessage(Winamp22Handle2, WM_USER, 0, ByVal ipc_isplaying)
    
    If MsgResult = 0 Or MsgResult = 3 Then
        
           MsgResult = SendMessage(Winamp22Handle2, WM_COMMAND, WINAMP_BUTTON2, ByVal XcNul)
    
    End If

End If

End Sub



Public Sub Mystery()


If Kbbdf = True Then
    Idle_Timer = Now
End If
    
    'Call Winamp_Player_On
    
If Kbbdf = True Then
'frmReceiver.txtTarget.Text = ""
    'Keyy = Keyy + 1
    'Daykeyy = Daykeyy + 1
    
    frmSender_CallID.txtMsg.Text = "KeyCounter " + str$(Keyy + Qwerty2)
    
    'Call frmSender_CallID.cmdSend_Click
    
    Xmarks = Now
    
    Call ClingDing
    
    If TotalDayMouse = 0 And TotalDayKeyy = 0 Then Label23.BackColor = Label23BackColor
    
    Call KeyOrMouse

End If


If Trim(globalhot$) = "" Then Exit Sub

If InStr(globalhot$, "Hit Request") Then
    frmSender_CallID.txtMsg.Text = "KeyCounter " + str$(Qwerty2 + Keyy)
    'Call frmSender_CallID.cmdSend_Click
    frmSender_CallID.txtMsg.Text = "MouseCounter " + str$(Mouse55 + Mousey)
    'Call frmSender_CallID.cmdSend_Click
End If

If InStr(globalhot$, "winamp play") Then
    Call Winamp_Player_On
    MattsTimer5 = Now + TimeSerial(0, 3, 0)
End If


If InStr(globalhot$, "winamp stop") Then
    Call winamp_player_off
    If MattsTimer5 > Now Then MattsTimer5 = Now + TimeSerial(0, 1, 30)
End If

If InStr(globalhot$, "Monitor To The Fucking A") Then
    'Call winamp_player_off
    MattsTimer5 = Now + TimeSerial(0, 3, 0)
End If

globalhot$ = ""

End Sub

Public Sub Lastpress()

Kbbdf = True



If Kbbdf = True Then Call Mystery

Kbbdf = False

End Sub

Private Sub Timer5_Timer()


'-----------------------------------------------------
'-----------------------------------------------------
'ONLY ONCE LATCH DO SOON AS POSS
'-----------------------------------------------------
'1ST TIMER TO RUN AWAY FORM FROM_LOAD AS DELAY ROUTINE
'-----------------------------------------------------
If TIMER8_LEFT_THE_FORM_LOAD_ENTRY_YET <> 2 Then
    Call Timer8_Timer
End If
'-----------------------------------------------------
'-----------------------------------------------------

Call LABLE_SONG_PLAYING

' ---------------------------------------------
' BELOW ONLY WRITE EXE PATH OF WINAMP TO LOGGER
' ---------------------------------------------
Dim WinampFileEXE$

WinampHWVM1 = FindWindow("Winamp v1.x", vbNullString)

If WinampHWVM1 = 0 Or WinampHWVM1 = WinampHWVM2 Then Exit Sub
'TTT = cProcesses.Convert(WinampHW, pid, cnFromhWnd Or cnToprocessid)
'TTT = cProcesses.Convert(pid, WinampFileEXE$, cnFromprocessid Or cnToexe)
WinampFileEXE$ = GetFileFromHwnd(WinampHWVM1)
WinampHWVM2 = WinampHWVM1

''D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\0TextData\0TextData\WinAmpEXEFileActive.txt"
Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\WinAmpEXEFileActive.txt", 1)

fre = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\WinAmpEXEFileActive.txt" For Output As #fre
    Print #fre, WinampFileEXE$
Close #fre

End Sub



Private Sub Timer_Idle_Timer()

If YYTT = 0 Then YYTT = Int(Now) + 1

If YYTT < Now Then
    YYTT = Int(Now) + 1
    On Error Resume Next
    'Kill App.Path + "\WinAmp Music Lenght Logg Day.txt"
    On Error GoTo 0
End If

Idle_Timer = 0

If Idle_Timer + TimeSerial(0, 1, 0) < Now And Idle_Timer > 0 Then
    Idle_Timer = Now - TimeSerial(0, 5, 0)
    Idle_Timer = 0
    On Local Error Resume Next
    'AppActivate "Winamp Playlist Editor", True
    On Local Error GoTo 0
End If


'-------------------
'NOT REQUIRE THINKER
'-------------------
'Exit Sub




If WinampArrayCount = 0 Then WinampArrayCount = 1

WinampHW = FindWindow("Winamp v1.x", vbNullString)

If WinampHW3 <> WinampHW And WinampHW > 0 Then
    WinampHW3 = WinampHW
    easyx = 0
    For R = 1 To WinampArrayCount
        If WinampHW = WinampArray(R) Then
            easyx = 1
        End If
    Next
    
    If easyx = 0 Then
    WinampArray(WinampArrayCount) = WinampHW
    'WinampHWSource2 = Now + TimeSerial(0, 0, 3)
    TTT = cProcesses.Convert(WinampHW, WinampHWSource, cnFromhWnd Or cnToProcessID)
    WinampArray2(WinampArrayCount) = WinampHWSource
    WinampArrayCount = WinampArrayCount + 1
    End If
End If



For R = 1 To WinampArrayCount - 1
'winamphw2 = FindWindow("Winamp v1.x", vbNullString)
If WinampArray(R) > 0 Then WinampHW = GetWindow(WinampArray(R), 0)
If WinampHW = 0 Then
    CM$ = GetFileFromProc(WinampArray2(R))
'    WinampFileExe$=
    If CM$ <> "" Then
        'Process_Kill (WinampArray2(R))
    End If
    WinampHWSource = 0
    WinampHW3 = 0
End If
Next

For R = 1 To WinampArrayCount
If WinampArray(R) > 0 Then WinampHW = GetWindow(WinampHW3, 0)
If WinampHW = 0 Then
    WinampArray(R) = 0
    WinampArrayCount = WinampArrayCount - 1
End If
Next
End Sub




Private Sub Date_File(Filespec_2, IDate As Date)

If Dir(Filespec_2) = "" Then
    If MSGBOX1 = False Then
        MSGBOX1 = True
        MsgBox "FILE NOT EXISTER " + vbCrLf + Filespec2$, vbMsgBoxSetForeground
    End If
End If

'MsgBox Filespec_2

PDATE = 0
FileSpec = Filespec_2
Set F = FS.GetFile(Filespec_2)
IDate = F.DateLastModified

'MsgBox IDate

End Sub


Sub AGROSTORESAVE()

    Exit Sub


    'If RADIO_SONG = False Then
        If AGRONOW > Now Then Exit Sub
        AGRONOW = Now + TimeSerial(0, 0, 20)
    'End If

    '---------------------
    'WHY CALL ITSELF ERROR
    '---------------------
'    Call AGROSTORESAVE


    If Last50EXEListTXT <> "" Then
        Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\Last50EXEList.txt", 1)
        hnd = FreeFile
        VMM = ""
        Open App.Path + "\0TextData\" + GetComputerName + "\Last50EXEList.txt" For Output Lock Write As #hnd
    
        Print #hnd, Last50EXEListTXT
        Close #hnd
        Last50EXEListTXT = ""
    End If
    
    If Last50EXEListAPPEND <> "" Then
        Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\Last50EXEList-AppendLogg.txt", 1)
        hnd = FreeFile
        Open App.Path + "\0TextData\" + GetComputerName + "\Last50EXEList-AppendLogg.txt" For Append Lock Write As #hnd
            Print #hnd, Last50EXEListAPPEND
        Close #hnd
        Last50EXEListAPPEND = ""
    End If

    Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\2HwndWinHovCountTotal.txt"
    
    If HWND_HOV_TOTAL <> "" Then
        Call FileInUseDelay(Path_And_FileName, 1)
        GUG = FreeFile
        Open Path_And_FileName For Append As #GUG
            Print #GUG, HWND_HOV_TOTAL
        Close #GUG
        HWND_HOV_TOTAL = ""
    End If


'Call AGROSAVE2

    
End Sub

Sub AGROSAVE2()

    If Trim(AGROCURRENT) <> "" And AGROCURRENT <> AGROCURRENTOLD Then
        Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Last One.txt"
        Call FileInUseDelay(Path_And_FileName, 1)
        Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Last One.txt"
        Fr1 = FreeFile
        Open Path_And_FileName For Output As #Fr1
            Print #Fr1, AGROCURRENT
        Close #Fr1
        AGROCURRENTOLD = AGROCURRENT
    
        Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day Total.txt", 1)
        Fr1 = FreeFile
        Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day Total.txt" For Append As #Fr1
            Print #Fr1, AGROCURRENT
        Close #Fr1
    
    
    End If

    If AGROSTORE <> "" Then

        Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\" + GetComputerName + "\WinAmp Music Lenght Logg.txt", 1)
        Fr1 = FreeFile
        Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg.txt" For Append As #Fr1
            Print #Fr1, AGROSTORE
        Close #Fr1
        
        Fr1 = FreeFile
            Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg-TRIM__" + GetComputerName + ".txt" For Append As #Fr1
        Close #Fr1
        ' Reset
        ' Stop
        Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg-TRIM__" + GetComputerName + ".txt", 1)
'        Debug.Print App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg-TRIM__" + GetComputerName + ".txt"
        Fr1 = FreeFile
        Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg-TRIM__" + GetComputerName + ".txt" For Append As #Fr1
            Print #Fr1, AGROSTORE
        Close #Fr1

        Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day.txt", 1)
        Fr1 = FreeFile
        Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day.txt" For Append As #Fr1
            Print #Fr1, AGROSTORE
        Close #Fr1
    
        AGROSTORE = ""

    End If

    If AGROSTOREEX <> "" Then
    
        Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day - Sig.txt", 1)
        Fr1 = FreeFile
        Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day - Sig.txt" For Append As #Fr1
            Print #Fr1, AGROSTOREEX
        Close #Fr1

        AGROSTOREEX = ""
    End If

End Sub

Private Sub Timer8_Timer()

Dim TIMER_FORMAT_ As String

Dim XTOP2 As Date
Dim RADIO_SONG


Label46.Caption = Format(Now, "HH:MM:SS")

If TIMER8_LEFT_THE_FORM_LOAD_ENTRY_YET = 0 Then Exit Sub
TIMER8_LEFT_THE_FORM_LOAD_ENTRY_YET = 2

'A = Timer8.Interval = 500 DEFAULT

'-----------------
'MAIN GET THE TUNE
'-----------------



Dim TEXT_BIT

'----
'2017
'----
'TIMER8.INTERVAL = 500
'2Total Music.txt
'----
'If TF2 = 0 Or WindowVisible(TF2) = False Then Exit Sub

TF2 = FindWindow("Winamp v1.x", vbNullString)

If TF2 = 0 Then
    If TF2 = 0 And WINAMP_DONT_EXIST_ANYMORE > 0 Then
        If WINAMP_DONT_EXIST_ANYMORE_DONE = False Then
            WINAMP_DONT_EXIST_ANYMORE_DONE = True
    AGRO = Format$(Now, "DD-MM-YYYY HH:MM:SS") + " __ WINAMP HAS LEFT THE WORKING _____________________________________"
            AGROSTORE = AGROSTORE + vbCrLf + AGRO
            Call AGROSAVE2
        End If
    End If
    Exit Sub
End If

WINAMP_HAS_BEEN_FOUND = TF2
If WINAMP_HAS_BEEN_FOUND > 0 And WINAMP_HAS_BEEN_FOUND <> O_WINAMP_HAS_BEEN_FOUND Then
    O_WINAMP_HAS_BEEN_FOUND = WINAMP_HAS_BEEN_FOUND
    
    AGRO = Format$(Now, "DD-MM-YYYY HH:MM:SS") + " __ WINAMP HAS ARRIVED ______________________________________________"
    AGROSTORE = AGROSTORE + vbCrLf + AGRO
    TIMER_FORMAT_ = Format$(Now, "DD-MM-YYYY HH:MM:SS")
    AGRO = ""
    AGRO = AGRO + TIMER_FORMAT_ + " __ WinAmp Music Lenght Logg-TRIM__7-ASUS-GL522VW.txt - Google Drive" + vbCrLf
    'AGRO = AGRO + TIMER_FORMAT_ + " __ " + Path_And_FileName = WinAmp_Music_Lenght_Logg_TRIM__7_ASUS_GL522VW + vbCrLf
    AGRO = AGRO + TIMER_FORMAT_ + " __ " + WinAmp_Music_Lenght_Logg_TRIM__7_ASUS_GL522VW + vbCrLf
    AGRO = AGRO + TIMER_FORMAT_ + " __ COMPUTER_CODING_VB - Google Drive" + vbCrLf
    AGRO = AGRO + TIMER_FORMAT_ + " __ https://drive.google.com/drive/folders/0BwoB_cPOibCPV0ZsSklwcE92X3c"
    
'i = "----" + vbCrLf
'i = i + "WinAmp Music Lenght Logg-TRIM__7-ASUS-GL522VW.txt - Google Drive" + vbCrLf
'i = i + "https://drive.google.com/file/d/0BwoB_cPOibCPV0RnUXNCZjJzT3M/view" + vbCrLf
'i = i + "----" + vbCrLf
'i = i + "----" + vbCrLf
'i = i + "COMPUTER_CODING_VB - Google Drive" + vbCrLf
'i = i + "https://drive.google.com/drive/folders/0BwoB_cPOibCPV0ZsSklwcE92X3c" + vbCrLf
'i = i + "----" + vbCrLf

    
    AGROSTORE = AGROSTORE + vbCrLf + AGRO
    Call AGROSAVE2
End If


WINAMP_HANDLE = TF2
If WINAMP_HANDLE <> O_WINAMP_HANDLE And O_WINAMP_HANDLE > 0 Then
    O_WINAMP_HANDLE = WINAMP_HANDLE
    AGRO = Format$(Now, "DD-MM-YYYY HH:MM:SS") + " __ WINAMP CHANGED IT HANDLE ________________________________________"
    AGROSTORE = AGROSTORE + vbCrLf + AGRO
    Call AGROSAVE2
'    TIMER_FORMAT_ = Format$(Now, "DD-MM-YYYY HH:MM:SS")
'    AGRO = ""
'    AGRO = AGRO + TIMER_FORMAT_ + " ----" + vbCrLf
'    AGRO = AGRO + TIMER_FORMAT_ + " WinAmp Music Lenght Logg-TRIM__7-ASUS-GL522VW.txt - Google Drive" + vbCrLf
'    AGRO = AGRO + TIMER_FORMAT_ + " https://drive.google.com/file/d/0B2hfTi1eLYL5dUxXY2NicGx6R2c/view" + vbCrLf
'    AGRO = AGRO + TIMER_FORMAT_ + " ----" + vbCrLf
End If



WINAMP_DONT_EXIST_ANYMORE = TF2
WINAMP_DONT_EXIST_ANYMORE_DONE = False




Dim Rect7 As RECT

hWnd8 = GetWindowRect(TF2, Rect7)

'If Rect7.Top < 0 Then Stop

'Debug.Print str(TF2) + " -- " + str(Rect7.Top)

'Debug.Print str(TF2) + " -- " + str(WindowVisible(TF2))

'Exit Sub

'If TF2 <> 0 And OTimeWinamp = 0 Then OTimeWinamp = Now + TimeSerial(0, 0, 3)
'If OTimeWinamp > Now Then Exit Sub
'TF2 = FindWindow("Winamp v1.x", vbNullString)
'If TF2 = 0 Then OTimeWinamp = 0

'----------------------------------------------------------------------

' --------------------------
' GET LABEL11 __ TRACK TITLE
' Call Timer9_Timer
' --------------------------

' For R2 = 1 To 4
    CurrentSongx1 = GetWindowTitle(TF2)
' Next
CurrentSongx1 = Trim(Replace(CurrentSongx1, "[Stopped]", ""))
CurrentSongx1 = Trim(Replace(CurrentSongx1, "[Paused]", ""))

If InStr(CurrentSongx1, "Buffer") > 0 Then
    XBR0 = InStr(CurrentSongx1, "Buffer")
    XBR1 = InStr(CurrentSongx1, "[")
    XBR2 = InStr(CurrentSongx1, "]")
    CurrentSongx1 = Mid(CurrentSongx1, 1, XBR1 - 2) + Mid(CurrentSongx1, XBR2 + 1)
End If

'If Len(CurrentSongx1) = 85 Then Stop
'If InStr(CurrentSongx1, "sylum Breaks - SlickWillyDee") = 4 Then Stop


'CurrentSongx1 = "[Buffer: 90%] Hypnotek Sessions - John Rowe -  - (FNOOB TECHNO RADIO)"
'If InStr(CurrentSongx1, "[Buffer: ") Then
'    '[Buffer: 0%]
'    If InStr(CurrentSongx1, "%]") Then
'        CurrentSongx1 = Mid(CurrentSongx1, InStr(CurrentSongx1, "%]") + 3)
'    End If
'End If



'----------------------------------------------------------------------------------------------
Test_RADIO = SendMessage(TF2, WM_USER, 1, ByVal IPC_GETOUTPUTTIME)
If Test_RADIO = -1 Then RADIO_SONG = True
If Test_RADIO > -1 Then RADIO_SONG = False


If CurrentSongx1_DOUBLE_BEFORE_CHECKER_2 <> CurrentSongx1 Then
    StartPlayTime_2 = Now
    'StartPlayTime_1 = Now
Else
'Exit Sub
    'StartPlayTime_1 = Now
    
End If

If CurrentSongx1_DOUBLE_BEFORE_CHECKER_2 <> CurrentSongx1 Then
    StartPlayTime_1 = Now
End If

CurrentSongx1_DOUBLE_BEFORE_CHECKER_2 = CurrentSongx1

'        If StartPlayTime_1 = 0 And StartPlayTime_2 > 0 Then
'            StartPlayTime_1 = StartPlayTime_2
'        Else
'            StartPlayTime_1 = Now
'        End If





'Debug.Print CurrentSongx1
'1. [HTTP/1.1 200 OK] http://play.fnoobtechno.com:2199/tunein/fnoobtechno.mp3 - Winamp

If RADIO_SONG = False Then

    'Maybe require here hen not digital hardwiring internet radio mode
    '-----------------------------------------------------------------
    If TF2 = 0 And Timer8Flag = True Then Exit Sub
    '-----------------------------------------------------------------

    msgresult33 = SendMessage(TF2, WM_USER, 1, ByVal IPC_GETOUTPUTTIME)
    '------------------------------------------------------------------
    'WHEN AN INTENET DIGITAL BROADCAST THE OUTPUT TIME WILL BE NAUGHT
    'SO WE WIL REQUIRE NAME OF SONG TO TELL TO IGNR THAT ONE
    '------------------------------------------------------------------

    '------------------------------------------------------------------
    'BYPASS SKIP OVER THIS BASTARD CUSTORD
    '------------------------------------------------------------------
'    If msgresult33 <= 0 Then
'
'        CurrentSongx1 = GetWindowTitle(TF2)
'
'        If InStr(CurrentSongx1, "(FNOOB TECHNO RADIO) - Winamp") = 0 Then
'            'Exit Sub
'        Else
'            Radio_Song = True
'        End If
'    End If


End If


If OldTF2 <> TF2 Then xxz = 1 Else xxz = 0

OldTF2 = TF2

Timer8Flag = True
'CurrentSongx1 = GetWindowTitle(TF2)

'CurrentSongx1 = Mid$(CurrentSongx1, 1, InStrRev(CurrentSongx1, " - Win"))
CurrentSongx1 = Trim(Replace(CurrentSongx1, "- Winamp", ""))

CurrentSongx1 = Trim(Replace(CurrentSongx1, "[Stopped]", ""))
CurrentSongx1 = Trim(Replace(CurrentSongx1, "[Paused]", ""))


ISET = 0
If RADIO_SONG = True Then
    'If InStr(CurrentSongx2$, CurrentSongx1) = 0 Or Label13 = "" Then ISET = 1
    CurrentSongx2$ = CurrentSongx1
    'If CurrentSongx2$ <> CurrentSongx1 Then ISET = 1
    ISET = 1
End If

CurrentSongx1 = Trim(Replace(CurrentSongx1, "[Stopped]", ""))
CurrentSongx1 = Trim(Replace(CurrentSongx1, "[Paused]", ""))

If RADIO_SONG = False Then
    'If CurrentSongx2$ <> CurrentSongx1 Or Label13 = "0:00:00" Or Label13 = "" Then ISET = 1
    If CurrentSongx2$ <> CurrentSongx1 Or Label13 = "" Then ISET = 1
End If
    

If BEFORE_StartPlayTime_2 <> StartPlayTime_2 Then
    BEFORE_StartPlayTime_2 = StartPlayTime_2
    ISET = 1
End If
    
    
If ISET = 1 Then
    If RADIO_SONG = False Then
        
        If ProgressBar1.Value > 5000 Then
            If ProgressBar1m > 0 Then ProgressBar2.Max = ProgressBar1m
            If ProgressBar1m = 0 Then ProgressBar2.Max = ProgressBar1.Max
        
            On Error Resume Next
            ProgressBar2.Value = ProgressBar1v
            On Error GoTo 0
        End If
        
        If Label13 = "0:00:00" And Label3v = "" And AGRO2 <> "" Then
            
            Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Last One.txt", 0)
            Fr1 = FreeFile
            Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Last One.txt" For Input As #Fr1
                Line Input #Fr1, AGRO2
            Close #Fr1
            StartPlayTime_1 = DateValue(Mid$(AGRO2, 1, 19)) + TimeValue(Mid$(AGRO2, 1, 19))
            StartPlayTime_2 = DateValue(Mid$(AGRO2, 1, 19)) + TimeValue(Mid$(AGRO2, 1, 19))
            tak = InStr(AGRO2, "Winamp Lenght Played =")
            tak = InStr(tak + 1, AGRO2, "=")
            tak2 = InStr(tak + 1, AGRO2, "--")
            Label3v = Trim(Mid$(AGRO2, tak + 2, tak2 - tak - 2))
                
        End If
        
        If Label3v <> "" Then
            Label13 = Label3v
        End If
        
        Label37________________ = Format(StartPlayTime_1, "H:MM:SSa/p")
        Label37_STARTTIME_2_GAP = Format(StartPlayTime_2, "H:MM:SSa/p")
                
        rap = Label13
        If Len(rap) = 5 Then rap = "00:" + Label13
        If Len(rap) = 7 Then rap = "0" + Label13
        
    End If
    
    STRIP_FIRST_DOT = Trim(Mid(CurrentSongx2$, InStr(CurrentSongx2$, ".") + 2))
    CurrentSongx2$ = STRIP_FIRST_DOT
    
    'AGRO = Format$(Now, "DD-MM-YYYY HH:MM:SS") + " -- Winamp Lenght Played = " + rap + " -- Song Name = " + CurrentSongx2$
    AGRO = Format$(Now, "DD-MM-YYYY HH:MM:SS") + " _Lenght Play_ " + rap + " _Tune_ " + CurrentSongx2$
    'Dim TEXT_BIT
    
    
    'If InStr(CurrentSongx2$, "__ FNOOB TECHNO RADIO") > 0 Then Stop
    
    
    If RADIO_SONG = True Then
        'i = InStr(CurrentSongx2$, "__")
        
        'If FORM_LOAD_CURRENT_SONG = True Then
            'FORM_LOAD_CURRENT_SONG = False
            If InStr(CurrentSongx2$, "__") > 0 Then
               CurrentSongx2$ = Mid(CurrentSongx2$, InStr(CurrentSongx2$, "__") + 3)
               CurrentSongx2$ = Trim(Replace(CurrentSongx2$, "[Stopped]", ""))
               CurrentSongx2$ = Trim(Replace(CurrentSongx2$, "[Paused]", ""))
               'STRIP_FIRST_DOT = Trim(Mid(CurrentSongx2$, InStr(CurrentSongx2$, ".") + 2))
               'CurrentSongx2$ = STRIP_FIRST_DOT
               'CHECK NURMERIC IF FRONT DOT
               'LATER
               '---------------------------
            End If
        'End If
        
        TEXT_BIT = "_Lenght Play_ " + rap + " _Tune_"
        STRING_LEN_TEXT_BIT = Len(TEXT_BIT)
        STRING_LEN_TEXT_BIT = 2
        AGRO = Replace(AGRO, TEXT_BIT, String(STRING_LEN_TEXT_BIT, "_"))
        AGRO = Replace(AGRO, " 1. ", " ")
        AGRO = Replace(AGRO, " 2. ", " ")
        '-----------------------
        'OUTPUT IS IN THIS STYLE
        '05-04-2017 01:21:38 __ Asylum Breaks Show with DJ SlickWillyDee and Janey the Asylum Nurse on NuBreaks.com Every Weds 7 til 9 pm GMT (2 til 4 pm EST) (Nubreaks.com Radio Asylum Breaks with DJ SlickWillyDee on NuBreaks.com)
        'SIXTH
        '-----------------------
    End If
            
'    If InStr(CurrentSongx2$, "__ FNOOB TECHNO RADIO") > 0 Then Stop
    
    'Find Not Used Now 25-10-09
    'Xx = InStr(CurrentSongx2$, "Song Name")
    'ii1 = Mid(CurrentSongx2$, Xx + 12)
    
    'Xx = InStrRev(AGRO, "Song Name")
    Xx = InStrRev(AGRO, "Tune_")
    ii2 = Trim(Mid(AGRO, Xx + 12))
    
    'If Label13 <> "0:00:00" And Label13 <> "" And Trim(ii1) <> "" And ii1 <> ii2 Then
        
    'CurrentSongx2$ NEW THIS WATCH FOR TESTING
    
    ALLOW_TO_STORE_WHEN_TUNES_MIXER_CHANGE = False
    If RADIO_SONG = True Then
        If CurrentSongx2$ <> CurrentSongx2_COMPARE_RADIO_SONG Then
            ALLOW_TO_STORE_WHEN_TUNES_MIXER_CHANGE = True
        End If
    End If
    
    CurrentSongx2_COMPARE_RADIO_SONG = CurrentSongx2$
    
    
    If ALLOW_TO_STORE_WHEN_TUNES_MIXER_CHANGE = True Then
        l = l
    End If
    
    
    '---------------------------------------------------------------------------------------------
    PASS_NEXT_IF_GATE_LOGIC_CONDITION = False
    '---------------------------------------------------------------------------------------------
    '01 OF 02
    '---------------------------------------------------------------------------------------------
    If Label13 <> "0:00:00" And Label13 <> "" And ii1 <> ii2 And Trim(CurrentSongx2$) <> "" Then
        PASS_NEXT_IF_GATE_LOGIC_CONDITION = True
    End If
    '---------------------------------------------------------------------------------------------
    
    '---------------------------------------------------------------------------------------------
    '02 OF 02
    '---------------------------------------------------------------------------------------------
    If ALLOW_TO_STORE_WHEN_TUNES_MIXER_CHANGE = True Then PASS_NEXT_IF_GATE_LOGIC_CONDITION = True
    '---------------------------------------------------------------------------------------------
    
    
    If PASS_NEXT_IF_GATE_LOGIC_CONDITION = True Then
        
        On Error GoTo ErrHand2
        
        If msgresult33 > 0 Or RADIO_SONG = True Then
            
            If RADIO_SONG = False Then
                rap = Label13
                If Len(rap) = 5 Then rap = "00:" + Label13
                If Len(rap) = 7 Then rap = "0" + Label13
                'AGRO = Format$(Now, "DD-MM-YYYY HH:MM:SS") + " -- Winamp Lenght Played = " + rap + " -- Song Name = " + CurrentSongx2$
                AGRO = Format$(Now, "DD-MM-YYYY HH:MM:SS") + " _Lenght Play_ " + rap + " _Tune_ " + CurrentSongx2$
            End If
            
            '------------------------------------------------------------
            'DON'T REQUIRE LENGHT NUMERIC INFO WHEN RADIO ALWAYS NAUGHTER
            '------------------------------------------------------------
            If RADIO_SONG = True Then
                'ALREADY DONE ABOVE CODER
                '------------------------
'                TEXT_BIT = "_Lenght Play_ " + rap + " _Tune_"
'                STRING_LEN_TEXT_BIT = Len(TEXT_BIT)
'                STRING_LEN_TEXT_BIT = 2
'                AGRO = Replace(AGRO, TEXT_BIT, String(STRING_LEN_TEXT_BIT, "_"))
'                AGRO = Replace(AGRO, " 1. ", " ")
            End If
            
            
            
            AGROSTORE = AGROSTORE + vbCrLf + AGRO
            
            
            '--------------------------------------------------------------------
            'SAVE THE LOGG OF OTHER THING EXE LIST AND SUCH NOT WANT REQUIRE HERE
            '--------------------------------------------------------------------
            'Call AGROSTORESAVE

            'SAVE THE MUSIC LOGG ONLY
            '------------------------
            Call AGROSAVE2
            'FILEAWAY = SAVE
            
        
            OUREXEWINAMP = False
            If RADIO_SONG = False Then
            
            If TF2 <> OLTF2 Then
                OLTF2 = TF2
                
                If CurrentSongx2$ <> CurrentSongx1 Or Txz1$ = "" Then
                
                    
                    Dim ExeS As String
                    ExeS = LCase(GetFileFromHwnd(TF2))
                    If InStr(ExeS, LCase("Winamp Gold 07 Video X")) Then
                        OUREXEWINAMP = True
                    Else
                        OUREXEWINAMP = False
                    End If
                    
'                    Otss = -1
'                    PIO = cProcesses.Convert(TF2, Otss, cnFromhWnd Or cnToProcessID)
'                    exen = cProcesses.GetEXEID(Otss, ExeS)

                
                    'Set oWinamp = CreateObject("WinampCOM.Application")
                    'Txz1$ = oWinamp.CurrentSongFileName
                    'Set oWinamp = Nothing
                End If
                
                'If InStr(Txz1$, "M:\Videos\00 Vid XXX") > 0 Then OUREXEWINAMP = True Else OUREXEWINAMP = False
            
            End If
    
    
    
            If RADIO_SONG = False Then
            If OUREXEWINAMP = False Then
                
                AGROEX = AGRO
                
            Else
                XSG = XSG + 1
                If XSG = 1 Then prvi = "Private -- (¯`'•.¸ -‹(•¿•)›- ¸.•'´¯)"
                If XSG = 2 Then prvi = "Private --  <º)))><¸..·´`·..¸><(((º>"
                If XSG = 3 Then prvi = "Private --  -:¦:-•:*'''''*:•-:¦:-"
                If XSG > 3 Then XSG = 0
                'Call FileInUseDelay(App.Path + "\0TextData\"+Computername+"\WinAmp Music Lenght Logg Day - Sig.txt", 1)
                'Fr1 = FreeFile
                'Open App.Path + "\0TextData\"+Computername+"\WinAmp Music Lenght Logg Day - Sig.txt" For Append As #Fr1
                    uu = InStr(AGRO, "Song")
                    If Mid(ii$, 10) <> Mid(LastMixLoad, 10) Then
                        AGROEX = Mid$(AGRO, 1, uu + 11) + prvi
                        'Print #Fr1, Mid$(AGRO, 1, uu + 11) + prvi
                    End If
                'Close #Fr1
            End If
        
            End If 'xxxx not xxxx If RADIO_SONG = False Then
            End If ' radio false
         
         
            AGROSTOREEX = AGROSTOREEX + vbCrLf + AGROEX
            AGROCURRENT = AGRO
                        
            '--------------------------------------------------------------------
            'SAVE THE LOGG OF OTHER THING EXE LIST AND SUCH NOT WANT REQUIRE HERE
            '--------------------------------------------------------------------
            'Call AGROSTORESAVE

            'SAVE THE MUSIC LOGG ONLY
            '------------------------
            Call AGROSAVE2
            'FILEAWAY = SAVE
            
'            Call AGROSTORESAVE
            
        End If
        
        On Error GoTo 0
    
    End If
End If

CurrentSongx2$ = CurrentSongx1
CurrentSongx2$ = Trim(Replace(CurrentSongx2$, "[Stopped]", ""))
CurrentSongx2$ = Trim(Replace(CurrentSongx2$, "[Paused]", ""))



'tf1 = FindWindow("Winamp PE", vbNullString)
'tf2 = FindWindow("Winamp v1.x", vbNullString)
'If tf1 > 0 Then

On Error Resume Next
'MsgResult = SendMessage(tf2, WM_USER, 0, ByVal ipc_isplaying)
MsgResult = ChkWinAmpPlay

If MsgResult = 3 Then MP3_PLAYING = "__" + "Paused ++"
If MsgResult = 0 Then MP3_PLAYING = "__" + "STOPPED ++"
If MsgResult = 1 Then MP3_PLAYING = "__" + "PLAYING": MP3_PLAYING = ""

If BEFORE_MSGRESULT_MP3_PLAYING <> MsgResult Then
    If BEFORE_MSGRESULT_MP3_PLAYING <> 1 Then
        StartPlayTime_2 = Now
        SEND_UPDATE_MATH = True
    End If
End If
BEFORE_MSGRESULT_MP3_PLAYING = MsgResult


'Lenght Song
MsgResult3 = SendMessage(TF2, WM_USER, 1, ByVal IPC_GETOUTPUTTIME)

'Posi in Song
MsgResult2 = SendMessage(TF2, WM_USER, 0, ByVal IPC_GETOUTPUTTIME)

If MsgResult <> 1 Then
    'WinStopTimer = -1 'Now + WinStopTimerTot
End If

'If radio_dong = True Then


If RADIO_SONG = False Or RADIO_SONG = True Then
If MsgResult > 0 And (MsgResult3 > 0 Or MsgResult2 > 0) Then
    
    If WinAmpL <> MsgResult3 And MsgResult3 > 0 Then
        ProgressBar1m = ProgressBar1.Max
        ProgressBar1.Max = MsgResult3 * 1000
        WinAmpL = MsgResult3
        F2 = DateDiff("h", 0, MsgResult3) - 1
        F3 = DateDiff("n", 0, MsgResult3) - 1
        F4 = DateDiff("s", 0, MsgResult3)
        f6 = Int(MsgResult3 / 3600)
        f7 = Int(MsgResult3 / 60)
        f8 = MsgResult3 Mod 60
        'Total
        lab$ = Format$(TimeSerial(0, f7, f8), "HH:MM:SS")
        If Mid$(lab$, 1, 1) = "0" Then lab$ = Mid$(lab$, 2)
        If Mid$(lab$, 1, 1) = "0" Then lab$ = Mid$(lab$, 3)
        Label4 = lab$
        
        Xtop = TimeSerial(0, f7, f8)
        Xtop5 = TimeSerial(0, f7, f8) - TimeSerial(0, 0, 1)
        Label4.FontSize = 12
    Else
        ProgressBar1.Max = 100
        ProgressBar1.Value = 100
        'Label4 = "NOT NUMERIC STREAM_ER"
        'Label4.Caption = "NOT NUMERIC STREAM_ER"
        'Label4.Caption = "NOT # INFO STREAM"
        Label4.Caption = "0x:xx:x0"
        'Label4.FontSize = 7
        Label4.FontSize = 12
        
    End If
    
    
    
    Hg = Int(MsgResult2 / 1000)
    f6 = Int((Hg / 60 ^ 2))
    f7 = Int(Hg / 60)
    f8 = Hg Mod 60
    XTOP2 = TimeSerial(0, f7, f8)
    
    'Position Now
    If XTOP2 > Xtop And MsgResult3 > 0 Then
        XTOP2 = Xtop
    End If
    
    lab$ = Format$(XTOP2, "HH:MM:SS")
    If Mid$(lab$, 1, 1) = "0" Then lab$ = Mid$(lab$, 2)
    If Mid$(lab$, 1, 1) = "0" Then lab$ = Mid$(lab$, 3)
    Label3v = Label3
    Label3 = lab$
    totter = DateDiff("s", XTOP2, Xtop)
    Hg = Int(totter)
    f6 = Int((Hg / 60 ^ 2))
    f7 = Int(Hg / 60)
    f8 = Hg Mod 60
    'totter = Now + TimeSerial(f6, f7, f8) %%
    totter = Now + TimeSerial(0, f7, f8)
    
    If RADIO_SONG = False Then
        Label18 = Format$(totter, "HH:MM:SSa/p") + MP3_PLAYING
    End If
    If RADIO_SONG = True Then
        Label18 = "24 Hour MidNight"
    End If
    
    
    Dim TiME_VAR As Double
    TiME_VAR = (Int(Now) + 1) '* 1000
    If MsgResult3 = -1 Then MsgResult3 = TiME_VAR 'DateDiff("s", MsgResult3, Int(Now) + 1) ' * 1000
    
    
    Hg = Int(MsgResult3) - Int(MsgResult2 / 1000)
    f6 = Int((Hg / 60 ^ 2))
    f7 = Int(Hg / 60)
    f8 = Hg Mod 60
    totter = TimeSerial(f6, f7, f8)
    
    'Format$(Totter, "HH:MM:SS")
    If totter >= TimeSerial(1, 0, 0) Then
        Label20 = Format$(totter, "H:MM:SS")
    Else
        'If totter < TimeSerial(0, 10, 0) Then
        Label20 = Mid$(Format$(totter, "H:MM:SS"), 3)
    End If
    
    If RADIO_SONG = True Then
        Label27.FontSize = 8
        Label27.Caption = "Time Left To on Next MidNight"
    Else
        Label27.FontSize = 12
        Label27.Caption = "Time Left"
    End If

    
    If Label3v = "" Then
            'Label3v = Label3
            Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Last One.txt", 0)
            Fr1 = FreeFile
            Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Last One.txt" For Input As #Fr1
                Line Input #Fr1, AGRO2
            Close #Fr1
            Label3v = str(AGRO2)
    End If
    
    If Check4.Value = vbChecked And XTOP2 >= Xtop5 And XTOP2 > 0 Then
        If MsgResult = 1 Then
            Call Any_Winamp_Player_Off_End
        End If
    End If
End If
End If


If MsgResult2 <= ProgressBar1.Max Then

ProgressBar1v = ProgressBar1.Value
ProgressBar1.Value = MsgResult2

'MsgResult2 = SendMessage(tf2, WM_USER, 0, ByVal IPC_GETOUTPUTTIME)

beeh = (MsgResult2 / ProgressBar1.Max) * 100
If Int(beeh) = 100 Then
    Label7 = Format$(beeh, "000") + "%"
Else
    Label7 = Format$(beeh, "0.0") + "%"
End If

'Open "u:\Temp\MP3%.txt" For Output Lock Write As #1
'Print #1, Format$(beeh, "0.00");
'Close #1

If Label3 = "00:00" Or beeh < 0 Then
    Label7 = "0%"
End If
End If

If TotalMusicDay <> Day(Now) Then

f6 = Int((TotalMusic2 / 60 ^ 2))
f7 = Int(TotalMusic2 / 60)
f8 = TotalMusic2 Mod 60
xtop3 = TimeSerial(0, f7, f8)

ttmd$ = Format$(xtop3, "HH:MM:SS")

On Local Error GoTo ErrHand2

Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\2Total_Music_Day_Log.txt", 1)
Fr1 = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\2Total_Music_Day_Log.txt" For Append As #Fr1
Print #Fr1, ttmd$ + " - Total Music Listern For Day Ending - " + Format$(Now, "DDD DD-MM-YYYY HH:MM:SS")
Close Fr1
If Mid$(ttmd$, 1, 1) = "0" Then ttmd$ = Mid$(ttmd$, 2)
Label17 = Format(TimeValue(ttmd$), "H:M:S")
TotalMusic2 = 0
TotalMusic1 = XTOP2

End If


TotalMusicDay = Day(Now)


If TotalMusic1 > 0 Then
    ttm = DateDiff("s", TotalMusic1, XTOP2)
    If ttm > 0 And ttm < 10 Then
    TotalMusic2 = TotalMusic2 + ttm
    Label12 = ttm
    End If
End If

TotalMusic1 = XTOP2

'Total Music

F1 = Int((TotalMusic2 / 60 ^ 2) / 24)
F2 = Int((TotalMusic2 / 60 ^ 2))
F3 = Int(TotalMusic2 / 60)
F4 = TotalMusic2 Mod 60
xtop3 = TimeSerial(0, F3, F4)
  
lab$ = Format$(xtop3, "H:MM:SS")

''second(TotalMusic2)
'If xtop3 >= TimeSerial(1, 0, 0) Then
'    lab$ = Format$(xtop3, "H:MM:SS")
'Else
'    lab$ = Format$(xtop3, "M:SS")
'End If


Label8 = lab$

Xxo = Label8.Caption
If Len(Label8) <= 5 Then Xxo = "0:" + Xxo
etphone = TimeValue(Xxo)
etphtot1 = TimeSerial(23, 59, 59)

result3 = (etphone / etphtot1) * 100
Label34 = Format(result3, "0.0") + "%"

On Error GoTo 0
On Local Error GoTo 0

Call OutPutMP3Sig



'MsgResult3 = SendMessage(Winamp24Handle2, WM_USER, 8, ByVal IPC_GETINFO)

'#define IPC_GETOUTPUTTIME 105
'/* int res = SendMessage(hwnd_winamp,WM_WA_IPC,mode,IPC_GETOUTPUTTIME);
'** returns the position in milliseconds of the current track (mode = 0),
'** or the track length, in seconds (mode = 1). Returns -1 if not playing or error.
'*/


'#define IPC_GETINFO 126
'/* (requires Winamp 2.05+)
'** int inf=SendMessage(hwnd_winamp,WM_WA_IPC,mode,IPC_GETINFO);
'** IPC_GETINFO returns info about the current playing song. The value
'** it returns depends on the value of 'mode'.
'** Mode      Meaning
'** ------------------
'** 0         Samplerate (i.e. 44100)
'** 1         Bitrate  (i.e. 128)
'** 2         Channels (i.e. 2)
'** 3 (5+)    Video LOWORD=w HIWORD=h
'** 4 (5+)    > 65536, string (video description)
'*/

Exit Sub
ErrHand2:
'err.Description
DoEvents
Reset
Fr1 = FreeFile
Resume
End Sub

Sub WinampFolderStop()
'most time this is off
If Check7.Value = vbUnchecked Then Exit Sub

'If tf2 = 0 Then
'    Set oWinamp = Nothing
'End If


'Dim oWinamp As New WINAMPCOMLib.Application

'ed$ = oWinamp.Status

TF2 = FindWindow("Winamp v1.x", vbNullString)
'MsgResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)
'If OldTF2 <> tf2 And tf2 > 0 Then

If TF2 <> 0 And OTimeWinamp = 0 Then OTimeWinamp = Now + TimeSerial(0, 0, 2)
If OTimeWinamp > Now Then Exit Sub
If TF2 = 0 Then OTimeWinamp = 0


'End If

'### SHOULD BE HERE REALLY NEW HOME BUT YEAH LOGGS WHAT
'### WINAMP.EXE CHANGES
xxz = 0
If OldTF2 <> TF2 And TF2 > 0 Then
xxz = 1
IFileName = GetFileFromHwnd(TF2)
On Local Error GoTo errhand3
Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\WinAmpLog\Last WinAmp Running.txt", 1)
Fr1 = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\WinAmpLog\Last WinAmp Running.txt" For Output As #Fr1
Print #Fr1, IFileName
Close #Fr1
End If
On Local Error GoTo 0
OldTF2 = TF2


'OldWinampHwnd = tf2

If TF2 = 0 Or WindowVisible(TF2) = False Then Exit Sub

If TF2 > 0 And MsgResult = 1 Then
Dim oWinamp 'As Object
'Dim oWinamp As New WINAMPCOMLib.Application
Set oWinamp = CreateObject("WinampCOM.Application")
Txz1$ = oWinamp.CurrentSongFileName
Set oWinamp = Nothing
End If

dxf1 = InStrRev(Txz1$, "\")
If dxf1 > 0 Then dxf2 = InStrRev(Txz1$, "\", dxf1 - 1)
If dxf2 > 0 Then txz2$ = Mid$(Txz1$, dxf2 + 1, dxf1 - dxf2)
If txz2$ <> OldWinampFileName And OldWinampFileName <> "" And dxf2 > 0 Then
If WinampTimerFileName + TimeSerial(0, 0, 2) < Now Then
    'stop
    Xeasy22 = Int(Now) + Timer + 1
    Do
    MsgResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)

    If MsgResult = 1 And xxz = 0 Then

        MsgResult2 = SendMessage(TF2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
    
'        Call Any_Winamp_Player_Off_End
    End If
    
    If Xeasy22 < Int(Now) + Timer Then
    Exit Do
    End If
    Loop Until MsgResult = 0
    
    

End If
End If

OldWinampFileName = txz2$
OldWinampHwnd = TF2

Exit Sub
errhand3:
DoEvents
Resume

End Sub


Sub OutPutMP3Sig()

If OutPutMP3SigNOW > Now Then Exit Sub
OutPutMP3SigNOW = Now + TimeSerial(0, 0, 10)



'For Saftey in case program crash
Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\2Total Music.txt", 1)
GUG = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\2Total Music.txt" For Output As #GUG
Print #GUG, TotalMusic1
Print #GUG, TotalMusic2
Print #GUG, TotalMusicDay
Close GUG

TF2 = FindWindow("Winamp v1.x", vbNullString)
MsgResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)

'// IPC_ISPLAYING returns the status of playback.
'// If it returns 1, it is playing. if it returns 3, it is paused,
'// if it returns 0, it is not playing.

If MsgResult = 1 Then statmp3 = "Playing."
If MsgResult = 3 Then statmp3 = "Paused."
If MsgResult = 0 Then statmp3 = "Not Playing."

If OldStatmp3 <> statmp3 Then

On Error Resume Next
Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\WinAmp Play Pause Log.txt", 1)
LF = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Play Pause Log.txt" For Append As #LF
Print #LF, Format$(Now, "DDD DD-MMM-YY HH:MM:SS") + " WinAmp Status = " + statmp3
Close #LF
On Error GoTo 0

End If

OldStatmp3 = statmp3
On Local Error GoTo 0
On Error GoTo 0

If (Second(Now) Mod 10 <> 0) And NoRunYet = True Then Exit Sub
NoRunYet = True
TF2 = FindWindow("Winamp v1.x", vbNullString)

MsgResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)

'// IPC_ISPLAYING returns the status of playback.
'// If it returns 1, it is playing. if it returns 3, it is paused,
'// if it returns 0, it is not playing.

If MsgResult = 1 Then statmp3 = "Playing"
If MsgResult = 3 Then statmp3 = "Paused"
If MsgResult = 0 Then statmp3 = "Not Playing"

te$ = "WinAmp v5.56 "
te$ = te$ + "#Status #" + statmp3 + vbCrLf


ll2$ = Label11
If ll2$ = "---- ** Title ** ----" Then
    ll2$ = "WinAmp v5.56 -- Is Not Loaded :)"
End If
    
ll2$ = Format(StartPlayTime_1, "DD-MM hh:mm") + " - " + ll2$

ll2_4$ = Format(StartPlayTime_2, "DD-MM hh:mm") + " - " + ll2$
'Finish ------------------------------------------------------


If OUREXEWINAMP = True Then
    XSG = XSG + 1
    If XSG = 1 Then prvi = "Private -- (¯`'•.¸ -‹(•¿•)›- ¸.•'´¯)"
    If XSG = 2 Then prvi = "Private --  <º)))><¸..·´`·..¸><(((º>"
    If XSG = 3 Then prvi = "Private --  -:¦:-•:*'''''*:•-:¦:-"
    If XSG > 3 Then XSG = 0
    te$ = te$ + prvi + vbCrLf
Else
    te$ = te$ + ll2$ + vbCrLf
End If


'-------------------------------------------------------------
'---------------------- HERE SORT LENGHT
'----------***************

Dim LENGTODO
LENGTODO = 250

'MsgBox ee5$



On Error Resume Next
Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day - Sig.txt", 0)
Fr1 = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\WinAmp Music Lenght Logg Day - Sig.txt" For Input As #Fr1
If LOF(Fr1) - 3000 > 0 Then Seek Fr1, LOF(Fr1) - 3000
Do
Line Input #Fr1, ll$

'cute = InStr(ll$, "Song Name =")
cute = InStr(ll$, "Tune_")
If cute > 0 Then
    If Len(ll$) > cute + 9 Then ll2$ = Mid$(ll$, 1, 5) + " " + Mid$(ll$, 12, 5) + " - " + Trim(Mid$(ll$, cute + 11))
    
    If ll2$ <> "" And Mid(ll2$, 13) <> Mid(oll2$, 13) Then
        ee5$ = ll2$ + vbCrLf + ee5$
        oll2$ = ll2$
    End If
End If
Loop Until EOF(Fr1)
Close #Fr1







If Len(ee5$) > LENGTODO Then

    'EE5$ = Mid$(EE5$, Len(EE5$) - LENGTODO)
    ee5$ = Mid$(ee5$, 1, LENGTODO)
'    MsgBox ee5$
    eet = InStrRev(ee5$, vbCrLf)
    If eet > 0 Then ee5$ = Mid$(ee5$, 1, eet)
End If

'ee5$ = vbCrLf + "Recent Media Played -" + vbCrLf + ee5$

'ee5$ = vbCrLf + ee5$


te$ = te$ + ee5$
'-----------------------------------------------------------------------------







te$ = te$ + vbCrLf

te$ = te$ + "Play Time " + Label4 + vbCrLf
te$ = te$ + "Media Time Poss " + Label3 + vbCrLf
te$ = te$ + "Played " + Label7 + vbCrLf
Xxr1 = Label8
If Len(Label8) <= 5 Then Xxr1 = "0:" + Xxr1: tss = "00:"
etphone = TimeValue(Xxr1)
etphtot1 = TimeSerial(23, 59, 59) 'WHOLE DAY
etphtot1 = TimeSerial(Hour(Now), Minute(Now), Second(Now)) 'TRUE DAY TO NOW

result3 = (etphone / etphtot1) * 100

Label34 = Format(result3, "0.0") + "%"

te$ = te$ + "Total Played Today " + tss + Label8 + "  " + Format(result3, "0.0") + "% of Day True to Now" + vbCrLf


On Error Resume Next
Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\2Total_Music_Day_Log.txt", 0)
Fr1 = FreeFile
Open App.Path + "\0TextData\" + GetComputerName + "\2Total_Music_Day_Log.txt" For Input As #Fr1
Seek Fr1, LOF(Fr1) - 140
Do
Line Input #Fr1, lr$
Loop Until EOF(Fr1)
Close #Fr1
If Err.Number > 0 Then Exit Sub
On Error GoTo 0

lr$ = Mid$(lr$, 1, 8)
If Mid$(lr$, 1, 1) = "0" Then lr$ = Mid$(lr$, 2)

etphone = TimeValue(lr$)
etphtot1 = TimeSerial(23, 59, 59)
result3 = (etphone / etphtot1) * 100

te$ = te$ + "Total Played Yester " + lr$ + "  " + Format(result3, "0.0") + "% of Day" + vbCrLf

Label17 = lr$
Label33 = Format(result3, "0.0") + "%"




'LastMixLoad = ll2$


'MsgBox EE5$


If Err.Number > 0 Then Exit Sub
On Error GoTo 0








'On Error GoTo Errorsub2

'te$ = te$ + vbCrLf
'te$ = te$ + "Mouse Today " + Label21 + vbCrLf
'te$ = te$ + "Keys    Today " + Label23 + vbCrLf
'te$ = te$ + vbCrLf
'te$ = te$ + "Mouse Yester " + str(M5) + vbCrLf
'te$ = te$ + "Keys    Yester " + str(K5) + vbCrLf
'te$ = te$ + vbCrLf
'te$ = te$ + "Windoze Licked Today " + str(CurProcHwnd_Count) + vbCrLf
'te$ = te$ + "Windoze Licked Yester " + str(CurProcHwnd_Count_Yester) + vbCrLf + vbCrLf
''te$ = te$ + vbCrLf
'te$ = te$ + "Windoze Hovered Today " + str(CurWinHovHwnd_Count) + vbCrLf
'te$ = te$ + "Windoze Hovered Yester " + str(CurWinHovHwnd_Count_Yester) + vbCrLf





On Error Resume Next

Path_Folder = "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures" + "__" + GetComputerName

If (Not DirExist(Path_Folder)) Then
    MkDirNested Path_Folder
End If

Path_And_FileName = "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures" + "__" + GetComputerName + "\Signature-MP3.txt"
Call FileInUseDelay(Path_And_FileName, 1)
'----------------------------------------
Fr1 = FreeFile
Open Path_And_FileName For Output As #Fr1
    Print #Fr1, te$;
Close #Fr1





'MsgBox te$

End Sub





Private Sub Timer9_Timer()


'-----------------------------------------------------
'ONLY ONCE LATCH DO SOON AS POSS
'-----------------------------------------------------
'1ST TIMER TO RUN AWAY FORM FROM_LOAD AS DELAY ROUTINE
'-----------------------------------------------------
'THIS SEEM MAIN ROUTINE HAVE WANT TO BE IN REASONER
'-----------------------------------------------------
If TIMER8_LEFT_THE_FORM_LOAD_ENTRY_YET <> 2 Then
    Call Timer8_Timer
End If
'-----------------------------------------------------


'----------------------------------
'Timer9.Interval === 500
'----------------------------------
'TIMER ON HALF A SECOND WHEN WORKER
'----------------------------------

'----------------------------------------------------
TF2 = FindWindow("Winamp v1.x", vbNullString)
If TF2 = 0 Then Exit Sub
'----------------------------------------------------
'DON'T REQUIRE TIME FOR STREAM DIGITAL HARDWIRE RADIO
'----------------------------------------------------
If InStr(GetWindowTitle(TF2), "(FNOOB TECHNO RADIO) - Winamp") > 0 Then Exit Sub


MsgResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)


If Label9 = "Zero" And MsgResult = 1 And (MsgResultMp3Old = 3 Or MsgResultMp3Old = 0) Then
    WinStopTimer = Now + WinStopTimerTot
    f8 = DateDiff("s", Now, WinStopTimer) Mod 60
    f7 = Int(DateDiff("s", Now, WinStopTimer) / 60)
    WinStopTimerH = TimeSerial(0, f7, f8)
End If



If MsgResult <> MsgResultMp3Old And MsgResultMp3Old <> -1 And MsgResult = 3 Then
'    If Dir$("D:\TEMP\PauseTimeWinampFlag2.txt") = "" Then
        'WinStopTimerH = WinStopTimer - Now
        If WinStopTimer > 0 Then
        f8 = DateDiff("s", Now, WinStopTimer) Mod 60
        f7 = Int(DateDiff("s", Now, WinStopTimer) / 60)
        WinStopTimerH = TimeSerial(0, f7, f8)
        End If
        'WinStopTimer = Now + WinStopTimerTot
'    End If
End If
If MsgResult <> MsgResultMp3Old And MsgResultMp3Old <> -1 And MsgResult = 1 Then
    WinStopTimer = Now + WinStopTimerH
    WinStopTimerH = 0
    Debug.Print WinStopTimerH
End If

MsgResultMp3Old = MsgResult

TF2 = FindWindow("Winamp v1.x", vbNullString)

Label9 = ttmd$

If TF2 = 0 Then Exit Sub

For R2 = 1 To 4
    ' WAIT FOR LONGER TEXT RESULT
    ' USUAL 1ST IS WRONG 1 CHAR MISSING
    ' ---------------------------
    GetWindowTitle_WINAMP = GetWindowTitle(TF2)
Next
rfh1 = InStrRev(GetWindowTitle_WINAMP, " - ")
If rfh1 > 0 Then GetWindowTitle_WINAMP = Mid$(GetWindowTitle_WINAMP, 1, rfh1 - 1)

If TF2 > 0 Then
    For R2 = 1 To 4
        Label11 = GetWindowTitle_WINAMP
        Label11 = GetWindowTitle_WINAMP
    Next
    If OLD_Label11 <> Label11 Then
        If InStr(Label11, "Buffer") > 0 Then
            XBR0 = InStr(Label11, "Buffer")
            XBR1 = InStr(Label11, "[")
            XBR2 = InStr(Label11, "]")
            Label11 = Mid(Label11, 1, XBR1 - 2) + Mid(Label11, XBR2 + 1)
        End If

        If Len(Label11) = 76 Then Stop
        If InStr(Label11, "sylum Breaks - SlickWillyDee") = 4 Then Stop

        If InStr(Label11, "uffer") > 0 Then Stop
    End If
    OLD_Label11 = Label11
End If

W = 0

'WinStopTimerTot = TimeSerial(0, 0, 15)

If WinStopTimer >= Now Then
    'f6 = Int((WinStopTimer / 60 ^ 2))
    f8 = DateDiff("s", Now, WinStopTimer) Mod 60
    f7 = Int(DateDiff("s", Now, WinStopTimer) / 60)
    xtop7 = TimeSerial(0, f7, f8)
End If

'If WinStopTimer <> -1 Then WinStopTimer = 0

ttmd$ = Mid$(Format$(xtop7, "HH:MM:SS"), 4)

If WinStopTimer < Now And WinStopTimer > 0 Then
    xtop7 = 0
    ttmd$ = Mid$(Format$(xtop7, "HH:MM:SS"), 4)
End If

If WinStopTimerH > 0 Then
    ttmd$ = "Zero"
    ttmd$ = Mid$(Format$(WinStopTimerH, "HH:MM:SS"), 4)
    'WinStopTimerH
End If

If WinStopTimer < 0 Then
    ttmd$ = "Zero"
End If

Label9 = ttmd$
rt = Check3.Value = vbChecked And Check5.Value = vbChecked And Check6.Value = vbChecked And Check8.Value = vbChecked

If WinStopTimer < Now And CurrentSong$ = "" And rt = True Then
'    Call Any_Winamp_Player_Off
    TF2 = FindWindow("Winamp v1.x", vbNullString)
    rrt$ = ""
    If TF2 > 0 Then rrt$ = GetWindowTitle(TF2)
    CurrentSong$ = rrt$
'    Check3.Value = vbUnchecked
End If


TF2 = FindWindow("Winamp v1.x", vbNullString)
CurrentSong1$ = GetWindowTitle(TF2)
CurrentSong1$ = Mid$(CurrentSong1$, 1, InStrRev(CurrentSong1$, " - Win"))
If WinStopTimer < Now And WinStopTimer > 0 Then
    Call Any_Winamp_Player_Off
End If
CurrentSong2$ = CurrentSong1$

End Sub

Public Sub Any_Winamp_Player_Off_End()
    
    TF2 = FindWindow("Winamp v1.x", vbNullString)

    MsgResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)
    
    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.

    If MsgResult = 1 Then
        MsgResult = SendMessage(TF2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
        WinStopTimer = -1
    End If
End Sub
Public Sub Any_Winamp_Player_Off()

rt = Check3.Value = vbUnchecked And Check5.Value = vbUnchecked And Check6.Value = vbUnchecked And Check8.Value = vbUnchecked
If rt = True Then Exit Sub

If CurrentSong1$ = CurrentSong2$ Then Exit Sub
TF2 = FindWindow("Winamp v1.x", vbNullString)
rrt$ = GetWindowTitle(TF2)

If TF2 <> OldTF2 Then Exit Sub
OldTF2 = TF2

If TF2 > 0 Then
    
    'MsgResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)
    
    '// IPC_ISPLAYING returns the status of playback.
    '// If it returns 1, it is playing. if it returns 3, it is paused,
    '// if it returns 0, it is not playing.

    
    Do
        MsgResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)
        DoEvents
        If MsgResult = 1 Then
            MsgResult = SendMessage(TF2, WM_COMMAND, WINAMP_BUTTON3, ByVal XcNul)
            WinStopTimer = -1
        End If
    Loop Until MsgResult <> 1

End If
        
End Sub




Sub FindTimeInfo()

'InPutDate = DateValue("01-12-2008")
'TestDate = DateValue("31-05-2009")
InPutDate = IDate
TestDate = Now

DayCountT = DateDiff("d", InPutDate, TestDate)

Year5 = 0
For r5 = Year(InPutDate) + 1 To Year(TestDate) + 1
    If DateSerial(r5, Month(InPutDate), Day(InPutDate)) < Now Then Year5 = Year5 + 1
Next
For r5 = 1 To -2 Step -1
    Xx = DateSerial(Year(TestDate), Month(TestDate) + r5, Day(InPutDate))
    MonthStart = Xx
    If Xx <= TestDate Then Exit For
Next
Month5 = 0
Month7 = 0
For r5 = 1 To 10000
    Xx = DateSerial(Year(InPutDate), Month(InPutDate) + r5, Day(InPutDate))
    If Year(Xx) <> oxx Then Month7 = 0
    oxx = Year(Xx)
    If Xx <= TestDate Then Month5 = Month5 + 1: Month7 = Month7 + 1
    If Xx >= TestDate Then Exit For
Next

For r5 = Year(TestDate) To 1 Step -1
    If DateSerial(r5, Month(InPutDate), Day(InPutDate)) < Now Then
        DayCount1 = DateDiff("d", DateSerial(r5, Month(InPutDate), Day(InPutDate)), TestDate)
        DayCountMonth = DateDiff("d", MonthStart, TestDate)
        WholeYear1 = DateDiff("d", DateSerial(r5, Month(InPutDate), Day(InPutDate)), DateSerial(r5 + 1, Month(InPutDate), Day(InPutDate)))
        'Month5 = DateDiff("m", DateSerial(r5, Month(InPutDate), Day(InPutDate)), TestDate)
        WeeksSinceYear = DateDiff("w", DateSerial(r5, Month(InPutDate), Day(InPutDate)), TestDate) - 1
        WeeksSinceInput = DateDiff("w", InPutDate, TestDate)
        WeeksPlusDays = DateDiff("d", InPutDate, TestDate) - (DateDiff("w", InPutDate, TestDate) * 7)
                        
                
        ResultYearDate = Format$(Year5 + DayCount1 / WholeYear1, ".000")
        Exit For
    End If
Next



tMm = DateDiff("s", InPutDate, TestDate)
F1 = Int((tMm / 60 ^ 2) / 24)
If F1 > 0 Then tMm = tMm - (DateDiff("s", Now, Now + 1) * F1)
F2 = Int((tMm / 60 ^ 2))
F3 = Int(tMm / 60)
F4 = tMm Mod 60

'EXAMPLE
ex = Format$(TimeSerial(0, F3, F4), "H:MM:SS")
ex1 = Format$(TimeSerial(0, F3, F4), "hH:MM:SS")
ex2 = str(Val(Mid$(ex1, 1, 2))) + "h" + str(Val(Mid$(ex1, 4, 2))) + "m" + str(Val(Mid$(ex1, 7, 2))) + "s"
ex3 = str(Val(Mid$(ex1, 1, 2))) + "h" + str(Val(Mid$(ex1, 4, 2))) + "m"

DexFORMAT1 = Trim(str(F1)) + " Days " + ex
DexFORMAT2 = Trim(str(F1)) + " Days " + ex2
DexFORMAT3 = Trim(str(F1)) + " Days " + ex3

DexFORMAT1 = Trim(str(F1)) + " Days " + ex
DexFORMAT2 = Trim(str(F1)) + " Days " + ex2
DexFORMAT3 = Trim(str(F1)) + " Days " + ex3

If F1 = 1 Then
    DexFORMAT1 = Trim(str(F1)) + " Day " + ex
    DexFORMAT2 = Trim(str(F1)) + " Day " + ex2
    DexFORMAT3 = Trim(str(F1)) + " Day " + ex3
End If

'If F1 = 1 Then
    DexFORMAT1 = Trim(str(F1)) + " D " + ex1
    DexFORMAT2 = Trim(str(F1)) + " D " + ex2
    DexFORMAT3 = Trim(str(F1)) + " D " + ex3
'End If


If F1 = 0 Then
    'DexFORMAT1 = ex1
    'DexFORMAT2 = ex2
    'DexFORMAT3 = ex3
End If



End Sub

Private Sub TimerDay_Timer()

If TimerDay.Interval <> 10000 Then TimerDay.Interval = 10000
'ONCE A DAY ROUTINE

DAY_NOW = Day(Now)
If OLD_DAY_NOW = DAY_NOW Then Exit Sub
OLD_DAY_NOW = DAY_NOW

'If Hour(Now) <> OHOURNOW Then Unload Me
OHOURNOW = Hour(Now)


te1$ = "" 'te1$ + vbCrLf
te1$ = te1$ + "Keys Today " + Label23.Caption + " - Yester " + str(K5) + vbCrLf
te1$ = te1$ + "Mouse Today " + Format(Val(Label21.Caption) / 1000, "0") + "K - Yester " + Format(Val(M5) / 1000, "0") + "K" + vbCrLf
te1$ = te1$ + vbCrLf
te1$ = te1$ + "Windoze Licked    Today " + str(CurProcHwnd_Count) + " - Yester " + str(CurProcHwnd_Count_Yester) + vbCrLf
te1$ = te1$ + "Windoze Hovered Today " + str(CurWinHovHwnd_Count) + " - Yester " + str(CurWinHovHwnd_Count_Yester) + vbCrLf

Path_Folder = "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures" + "__" + GetComputerName

If (Not DirExist(Path_Folder)) Then
    MkDirNested Path_Folder
End If

Path_And_FileName = "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures" + "__" + GetComputerName + "\Signature-KEYS.txt"

On Error Resume Next
'On Error GoTo 0
Call FileInUseDelay(Path_And_FileName, 1)

Path_And_FileName = "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures" + "__" + GetComputerName + "\Signature-KEYS.txt"
Fr1 = FreeFile
Open Path_And_FileName For Output As #Fr1
    Print #Fr1, te1$;
Close #Fr1


FILEX = "D:\VB6\VB-NT\00_Best_VB_04\Webcam_Motion\VBDataNoArchive\ArrayPic.jpg"
'If FS.FileExists(FILEX) Then File
    Set F = FS.GetFile(FILEX)
    IDate = F.DateLastModified
'End If
Call FindTimeInfo


If DateDiff("n", IDate, Now) > 15 Then
    Label39.BackColor = QBColor(12)
Else
    Label39.BackColor = 0
End If

Label39 = "WC " + DexFORMAT1

'If F1 = 0 Then Label39 = "WC" + str(F2) + "h" + str(F3) + "m" + str(F4) + "s"
'If F1 > 0 Then Label39 = "WC " + Format(F1, "0") + "d" + str(F2) + "h" + str(F3) + "m" + str(F4) + "s"

If Day(DayCheck) <> Day(Now) Or SunSet = 0 Then
    DayCheck = Now
    Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\SunRiseSet.txt"
    
    If Dir(Path_And_FileName) <> "" Then
        Call FileInUseDelay(Path_And_FileName, 0)
        
        Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\SunRiseSet.txt"
        Open Path_And_FileName For Input As #1
            Line Input #1, SR$
            Line Input #1, SS$
        Close #1
    
        If SR$ = "" Then SunSet = 1: Exit Sub
    
        SunRise = TimeValue(SR$) - TimeSerial(0, 40, 0)
        If TimeValue(SS$) < TimeSerial(22, 0, 0) Then
            SunSet = TimeValue(SS$) + TimeSerial(2, 0, 0)
        Else
            SunSet = TimeSerial(23, 59, 59)
        End If
        SunRise = SunRise + TimeSerial(0, 10, 0)
        SunSet = SunSet + TimeSerial(0, 5, 0)
    End If
End If

If Day(DayCheck) <> Day(Now) Then
    
    Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCount.txt"
    Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCount.txt", 1)
    Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCount.txt"
    GUG = FreeFile
    Open App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCount.txt" For Append As #GUG
    Print #GUG, "End Of Day Counter  = "; Now
    Print #GUG, "Day Mouse           = "; Daymouse
    Print #GUG, "Day Keyboard        = "; Daykeyy
    Close #GUG
    
    'On Local Error GoTo Errorsub
    
    Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCountLog.txt"
    Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCountLog.txt", 1)
    Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCountLog.txt"
    GUG = FreeFile
    Open App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCountLog.txt" For Append As #GUG
    
    d$ = "          "
    u$ = "          "
    RSet d$ = Trim(str$(Daymouse))
    RSet u$ = Trim(str$(Daykeyy))
    
    wer$ = Format$(Now, "DD/MM/YY HH:MM:SS") + " - " + "Mouse = " + d$ + " @ " + "Keyboard = " + u$
    Print #GUG, wer$
    Close #GUG
    
    
    Path_Folder = "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures" + "__" + GetComputerName

    If (Not DirExist(Path_Folder)) Then
        MkDirNested Path_Folder
    End If
    
    Path_And_FileName = "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures" + "__" + GetComputerName + "\Signature-MouseKey.txt"

    Call FileInUseDelay(Path_And_FileName, 1)
    Path_And_FileName = "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures" + "__" + GetComputerName + "\Signature-MouseKey.txt"
    GUG = FreeFile
    Open Path_And_FileName For Output As #GUG
        Print #GUG, "Mouse Yester = "; d$
        Print #GUG, "Keys Yester = "; u$
    Close #GUG
    M5 = Val(d$)
    K5 = Val(u$)
   
    
    OptionsMenu.List1.AddItem wer$
    
    DayMouseTimer = Now + TimeSerial(0, 0, 10)
    TotalDayMouse = Daymouse
    TotalDayKeyy = Daykeyy
    Daymouse = 0
    Daykeyy = 0

'If GetUserName = "Matt3" Then
    If chucky = 0 Then
        Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCount.txt"
        Call FileInUseDelay(App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCount.txt", 1)
        
        Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCount.txt"
        GUG = FreeFile
        Open App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCount.txt" For Output As #GUG
            Print #GUG, "Day Counter  = "; Now
            Print #GUG, "Day Mouse    = "; Daymouse
            Print #GUG, "Day Keyboard = "; Daykeyy
        Close #GUG
    End If
End If
'End If

On Local Error GoTo 0

'If AsusTimer20 > Now Then Call Proc25Asus


If DayMouseTimer < Now And DayMouseTimer > 0 And (Daykeyy > 0 Or Daymouse > 0) Then
    DayMouseTimer = 0
    TotalDayKeyy = 0
    TotalDayMouse = 0
    Label21.BackColor = Label21BackColor
    Label23.BackColor = Label23BackColor
End If


W = 0
If MattsTimer2 > Now Then W = DateDiff("s", Now, MattsTimer2)

If W > 0 Then
    If MattsTimer5 > Now Then
        If Label10.BackColor <> Label10.BackColor Then Label10.BackColor = Label10.BackColor
    Else
        If Label10.BackColor <> &HFF0000 Then Label10.BackColor = &HFF0000
    End If
End If

If W = 0 And MattsTimer5 > Now Then
    W = DateDiff("s", Now, MattsTimer5):
    If Label10.BackColor <> &H0 Then Label10.BackColor = &H0
End If

Label10.Caption = Trim$(str$(W))

Label2.Caption = Format$(Now, "DDD D MMM H:MM:SSa/p")

GoTo Skip3

If Xmarks < Now And Xmarks > 0 Then
    Xmarks = 0
    Label23.BackColor = Label23BackColor
End If

If NoMonMusChange <> (NoMonOff2 + 5) + (NoMonOff + 2) + (NoMusic + 4) Then
If NoMonOff = 0 Then Command3.BackColor = &HFFFF&
If NoMonOff = 1 Then Command3.BackColor = &HFF00&
If NoMonOff2 = 1 Then Command3.BackColor = &HFF80FF
If NoMusic = 0 Then Command1.BackColor = &HFFFF&
If NoMusic = 1 Then Command1.BackColor = &HFF00&
NoMonMusChange = (NoMonOff2 + 5) + (NoMonOff + 2) + (NoMusic + 4)
End If


If ScreenWidthX = 19200 And ScreenHeightY = 12000 And ScrWidth_Old <> 19200 Then

    Me.Left = (Screen.Width - Me.Width)
    Me.Top = (Screen.Height - Me.Height)

    'CID_Run_Me.Left = ScreenWidthX - CID_Run_Me.Width - OffSetGoogle
'    HeightOffSet = 400

    'CID_Run_Me.Top = ScreenHeightY - (CID_Run_Me.Height + HeightOffSet)

End If
Skip3:



ScrWidth_Old = ScreenWidthX

Call dss_download



If Day(DayCheck) = Day(Now) Then chucky = 1
DayCheck = Now

'---------------------------
Exit Sub


If Tibulate > 0 And Tibulate < Now Then
    Tibulate = 0 ': CID_Run_Me.WindowState = 0
End If
If FormLoad1 = 1 Then
       MattsTimer2 = Now + MinorOnTime + MinorOnTime
    FormLoad1 = 0
End If

If Mid$(Time$, 4) = "20:00" Then
    MattsTimer5 = Now + TimeSerial(0, 2, 0)
End If

If Mid$(Time$, 4) = "40:00" Then
    MattsTimer5 = Now + TimeSerial(0, 2, 0)
End If

If Mid$(Time$, 4) = "00:00" Then
    MattsTimer5 = Now + TimeSerial(0, 8, 0)
End If

Dim MattsLockMonOff7AM

Dim Hg As Date

Hg = TimeValue(str$(Now))

If Hg < TimeSerial(7, 0, 0) Then
    MattsLockMonOff7AM = True
Else
    MattsLockMonOff7AM = False
End If

'Label3.Caption = str(NoMonoff)

If MattsTimer5 < MattsTimer2 And MattsTimer2 > 0 Then MattsTimer5 = Now - 3
If MattsTimer2 < MattsTimer5 And MattsTimer5 > 0 Then MattsTimer2 = Now - 3

If (MattsTimer5 < Now And MattsTimer5 > 0) And _
    (MattsTimer2 < Now And MattsTimer2 > 0) Then
'        MsgBox ("Ping")
        
        
        
    If Monitor2_Off_Timer = 0 Then
 '       MsgBox ("Ping")
        If NoMonOff = 0 And NoMonOff2 = 0 Then
            MattsTimer5 = 0: MattsTimer2 = 0
            
            'SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_OFF
        End If
    End If
End If


'If MattsTimer5 < Now Then MattsTimer5 = 0
'If MattsTimer2 < Now Then MattsTimer2 = 0


If MattsTimer2 > Now And MattsTimer2Old <> MattsTimer2 Then
    MattsLockMonOff7AM = False
End If

If MattsTimer2 > Now And MattsTimer2Old < Now Or _
   MattsTimer5 > Now And MattsTimer5Old < Now Then
    If MattsLockMonOff7AM = False Then
        SendMessage Me.hWnd, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_ON
    End If
End If



MattsTimer2Old = MattsTimer2
MattsTimer5Old = MattsTimer5


'---------------------------

If Music_Off_Timer < Now Then Music_Off_Timer = 0
If f12_off_timer < Now Then f12_off_timer = 0
If Monitor2_Off_Timer < Now Then Monitor2_Off_Timer = 0

If Music_Off_Timer = 0 And Label1.Caption <> "Music On" Then Label1.Caption = "Music On": Command1.Caption = "Music"
If Monitor2_Off_Timer = 0 And Label14.Caption <> "Monitor On" Then Label14.Caption = "Monitor On": Command3.Caption = "Monitor"

If Music_Off_Timer <> 0 Then
    If Command1.Caption <> "Music Off" Then Command1.Caption = "Music Off"
    Label1.Caption = Format$(Music_Off_Timer - Now, "HH:MM:SS")
End If

If Monitor2_Off_Timer <> 0 Then
    If Command3.Caption <> "Monitor Off" Then Command3.Caption = "Monitor Off"
    Label14.Caption = Format$(Monitor2_Off_Timer - Now, "HH:MM:SS")
End If
    


Exit Sub

Errorsub:
DoEvents

End Sub

Private Sub TimerFileSave_Timer()
TimerFileSave.Enabled = False
Exit Sub
    
    
On Error GoTo Ende5

Ende5:

If Err.Number = 0 Then TimerFileSave.Enabled = False

End Sub

Private Sub TimerFileSave2_Timer()
TimerFileSave2.Enabled = False
Exit Sub

On Error GoTo Ende6
Ende6:

If Err.Number = 0 Then TimerFileSave2.Enabled = False

End Sub

Private Sub TimerSTOP_ON_FOLDER_Timer()


'NOT MEANT TO BE RUN UNLESS ASK
'------------------------------
TimerSTOP_ON_FOLDER.Enabled = False

Exit Sub


TF2 = FindWindow("Winamp v1.x", vbNullString)
MsgResult = SendMessage(TF2, WM_USER, 0, ByVal ipc_isplaying)
If MsgResult <> 1 Then Exit Sub
If WindowVisible(TF2) = False Then Exit Sub

'50% OF THE MP3 MUST PLAY -- FOR STOP ON FOLDER
'AY -- OF THE PREVIOUS
'ITS HASSEL USING WINAMP COMM ON START MP3'S
'ONLY WAY DETERMIN THAT NEXT SONG BEFORE PLAYED IS NEW FOLDER
'EASY

'Lenght Song
'MsgResult3 = SendMessage(TF2, WM_USER, 1, ByVal IPC_GETOUTPUTTIME)

'Posi in Song
'MsgResult2 = SendMessage(TF2, WM_USER, 0, ByVal IPC_GETOUTPUTTIME)

'DONT BOTHER NOW DO ASSEL WAY IMEDIATES
'XX1 = Round((MsgResult2 / (MsgResult3 * 1000) * 100), 4)
'If XX1 < 50 Then
'    Exit Sub
'End If

If GetWindowTitle(TF2) = CurrentSongx20 Then Exit Sub
CurrentSongx20 = GetWindowTitle(TF2)




'Timer8Flag = True
'CurrentSongx1 = GetWindowTitle(TF2)
'CurrentSongx1 = Mid$(CurrentSongx1, 1, InStrRev(CurrentSongx1, " - Win"))
Call WinampFolderStop

End Sub

Private Sub VocdeTimer_Timer()
VocdeTimer.Enabled = False
Exit Sub

On Local Error GoTo ErrHand5
    
'If Dir$("D:\TEMP\KeyHit-Tidal-VCodes.txt") <> "" Then
'    frg = FreeFile
'    Open "D:\TEMP\KeyHit-Tidal-VCodes.txt" For Binary As #frg
'    VcodeTT$ = Space$(LOF(frg))
'    Get #1, , VcodeTT$
'    Close #frg
'    Kill "D:\TEMP\KeyHit-Tidal-VCodes.txt"
'End If

If VcodeTT$ = "" Then Exit Sub
    
    
Path_Folder = App.Path + "\0TextData\" + GetComputerName + "\Text_Data_KeyLogg\"

If (Not DirExist(Path_Folder)) Then
    MkDirNested Path_Folder
End If

Dim Fr1

Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\Text_Data_KeyLogg\Key Logger Text.txt"
Call FileInUseDelay(Path_And_FileName, 1)

Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\Text_Data_KeyLogg\Key Logger Text.txt"
Fr1 = FreeFile
Open Path_And_FileName For Append As #Fr1
    Print #Fr1, VcodeTT$;
Close #Fr1

Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\Text_Data_KeyLogg\Key Logger Text " + Format$(KeyLoggDate, "YYYY-mm-dd HH-MM-SS") + ".txt"
Call FileInUseDelay(Path_And_FileName, 1)

Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\Text_Data_KeyLogg\Key Logger Text " + Format$(KeyLoggDate, "YYYY-mm-dd HH-MM-SS") + ".txt"
Fr1 = FreeFile
Open Path_And_FileName For Append As #Fr1
    Print #Fr1, VcodeTT$;
Close #Fr1

VcodeTT$ = ""
ErrHand5:
DoEvents
End Sub



Function ChkWinAmpPlay()

Dim Retval As Long, ProcessID As Long, ThreadID As Long
Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
Dim WinClass As String, WinTitle As String

'Dim WinArray() As Long
'Public WinAmp2, OWinAmp2, WinCap
'ReDim WinArray(20)

WinAmp2 = FindWindow("Winamp v1.x", vbNullString)
MsgResult = SendMessage(WinAmp2, WM_USER, 0, ByVal ipc_isplaying)
ChkWinAmpPlay = MsgResult

Exit Function

MsgResult = 0

    WinAmp2 = FindWindow("Winamp v1.x", vbNullString)
    
    If OWinAmp2 <> WinAmp2 And WinAmp2 > 0 Then
        cuk = 0
        For R = 1 To 20
            If WinArray(R) = WinAmp2 Then cuk = 1: Exit For
        Next
        If cuk = 0 Then
            For R = 1 To 20
                If WinArray(R) = 0 Then WinArray(R) = WinAmp2: Exit For
            Next
        End If
            
        For R = 1 To 20
            If WinArray(R) > 0 Then
                pot = GetWindowTitle(WinArray(R))
                If pot = "" Then
                    WinArray(R) = 0
                End If
            End If
        Next
    End If
    
    
    
    OWinAmp2 = WinAmp2
        For R = 1 To 20
            If WinArray(R) > 0 Then
                Retval = GetClassName(WinArray(R), WinClassBuf, 255)
                 WinClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
                 'RetVal = GetWindowText(lhWnd, WinTitleBuf, 255)
                 'WinTitle = StripNulls(WinTitleBuf)
                If WinClass = "Winamp v1.x" Then
                    MsgResult = SendMessage(WinArray(R), WM_USER, 0, ByVal ipc_isplaying)
                    If MsgResult = 1 Then Exit For
                End If
            End If
        Next
    
ChkWinAmpPlay = MsgResult
    

End Function

Property Let WindowVisible(hWnd As Long, New_Value As Boolean)

    Call ShowWindow(hWnd, IIf(New_Value, SW_SHOW, SW_HIDE))

End Property

Property Get WindowVisible(hWnd As Long) As Boolean

    WindowVisible = (IsWindowVisible(hWnd) > 0)
  
End Property

Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim l As Long
   Dim S As String
   l = GetWindowTextLength(hWnd)
   S = Space(l + 1)
   GetWindowText hWnd, S, l + 1
   GetWindowTitle = Left$(S, l)
End Function


Sub FileInUseDelay(Tx_FileInUseDelay, i)
        
    PATH_FILE_AT_ENTRY_SUBROUTINE = Tx_FileInUseDelay
    
    entry = 0
    QQ = Now + TimeSerial(0, 4, 0)
    QQ4_ = False
    Do
        'DoEvents
        ii = FileInUse(Tx_FileInUseDelay)
        If ii = True Then
            If entry = 0 Then
                CntFIU = CntFIU + 1
                Label40.Caption = Trim(str(CntFIU Mod 10))
                If CntFIU Mod 10 = 0 Then CntFIU = 0
            End If
            entry = 1
            
            If i = 1 Then
                Label40.BackColor = QBColor(12)
            Else
                Label40.BackColor = QBColor(10)
            End If
            Sleep (200)
        End If
        
        If PATH_FILE_AT_ENTRY_SUBROUTINE <> Tx_FileInUseDelay Then
            If Path_And_FileName <> Tx_FileInUseDelay Then
                'If IsIDE = True Then Stop
                If IsIDE = True Then
                    Shell "EXPLORER /SELECT, """ + Tx_FileInUseDelay + """", vbMaximizedFocus
                End If
            End If
        End If
        
        'If CID_RUN_ME_VAR_TO_EXIT_HAPPEN = True Then Unload Me: Exit Sub
        
        If QQ < Now Then QQ4_ = True: Exit Do
    Loop Until ii = False
    If ii = False Then Label40.BackColor = 0
    
    If ii = True Or QQ4_ = True Then
        
        'MINUTE
        Debug.Print str(DateDiff("n", QQ, Now)) + "MINUTE DELAY FILE IN USE"
        
        If Tx_FileInUseDelay = "2cidrunME-method.txt" Then
            Shell "EXPLORER /SELECT, """ + Tx_FileInUseDelay + """", vbMaximizedFocus
            MsgBox "Trouble File Stuck In Use freq error report MAYBE THIS FROM FORM LOAD WHEN ACTIVE TO STOP FILTER REQUIRE -- " + vbCrLf + Tx_FileInUseDelay + vbCrLf + "Longer than 4 Mins" + vbCrLf + "TAKE YOU THERE CHECK UNLOCKER STATE MAYBE REBOOT REQUIRING"
        End If
        
        MsgBox "Trouble File Stuck In Use " + vbCrLf + Tx_FileInUseDelay + vbCrLf + "Longer than 20 Mins"
        '2cidrunME-method.txt
        '2cidrunME-method.txt
        If IsIDE = True Then Stop
    End If


    
    'BECUASE OF THE DOEVENTS
    'THE PATH SHOULD BE CHECKED AT EXIT AS ANOTHER REQUEST TO HERE MAY BEEN DONE AS CODE CHECKER SHOWEN
    'IT IS POSIABLE TO PUT THE VARIABEL BACK
    'WHEN IT IS THE COMMON ENTRY ONE OUTSIDE OF THIS SUB ROUTINE
    'OR BE CARE AND SET IT TWICE OUTSIDE THE ROUTINE
    'EXAMPLE HOW DONE
    

    
    If PATH_FILE_AT_ENTRY_SUBROUTINE <> Tx_FileInUseDelay Then
        If Path_And_FileName <> Tx_FileInUseDelay Then
            
            Debug.Print PATH_FILE_AT_ENTRY_SUBROUTINE
            Debug.Print Path_And_FileName
            Debug.Print Tx_FileInUseDelay
            Debug.Print str(DateDiff("n", QQ - TimeSerial(0, 20, 0), Now)) + " -- MINUTE DELAY FILE IN USE"
            Debug.Print str(DateDiff("s", QQ - TimeSerial(0, 20, 0), Now)) + " -- SECOND DELAY FILE IN USE"
            
            'If IsIDE = True Then Stop
            
            Path_And_FileName = Tx_FileInUseDelay
            
            If PATH_FILE_AT_ENTRY_SUBROUTINE_RUN_ONCE = True Then
                PATH_FILE_AT_ENTRY_SUBROUTINE_RUN_ONCE = False
                'If IsIDE = True Then Stop
            End If
        End If
    End If
End Sub





Sub zzLoad_Checks()

'Exit Sub

'Put This In Declarations
'Dim OCheckQ
'Very Nice Code Will Save all your Check Boxes an Menu Checks and Values you can filter
'some out -- works as best I Know

'if you use menu checks you have to set them yourself on clicks
'If Mnu_VB.Checked = True Then Mnu_VB.Checked = False: Exit Sub
'If Mnu_VB.Checked = False Then Mnu_VB.Checked = True

'3 Parts
' ---
'1 Load
'2 Save
'3 Timer ' This Chks for Changes using XOR Hash
'   Best way I know with Timer Hardly CPU Usage Unfriendly

'zzCheckTimer.Enabled = True
'Dont Have Timer Enabled on Form Load

'Call Timer to Run On Form Unload ----- call zzCheckTimer_Timer
'Then you could set timer slow like 10 Secs - I Do

zzCheckTimer.Enabled = False

Dim Th()
ReDim Th(Me.Controls.Count * 3)

On Error Resume Next
i = 0
Open App.Path + "\0TextData\" + GetComputerName + "\ChkSettings.txt" For Input As #1
Do
    Line Input #1, vv$
    Th(i) = vv$
    i = i + 1
    Line Input #1, vv$
    Th(i) = Val(vv$)
    i = i + 1
    Line Input #1, vv$
    Th(i) = Val(vv$)
    i = i + 1
Loop Until EOF(1)
Close #1
    
tit = i
'Debug.Print "0--------------"
For Each Control In Me.Controls
    Xx = 1
    
    ppi = LCase(Control.Name)
    'Debug.Print ppi
    If InStr(ppi, LCase("Combo")) Then Xx = 0
    If InStr(ppi, LCase("Check")) Then Xx = 0
    If InStr(ppi, LCase("Mnu")) Then Xx = 0
    If InStr(ppi, LCase("Menu")) Then Xx = 0
    
    
    xxd = -1
    For R = 0 To tit
        If Control.Name = Th(R) Then
            xxd = R: Exit For
        End If
    Next
    
    If xxd >= 0 And Xx = 0 Then
        On Error Resume Next
        If Th(xxd + 1) <> 0 Then Control.Value = Th(xxd + 1)
        If Th(xxd + 2) <> 0 Then Control.Checked = Th(xxd + 2)
        'Dont Let People Play Around If you Want to Hard Code In Designer To Enable Chk This
        'This Lets Happen Eg If Th(xxd + 2) <> 0
        
        A1 = 0
        A1 = Control.Value
        B1 = 0
        B1 = Control.Checked
        OCheckQ = OCheckQ Xor (A1 + B1)
        On Error GoTo 0
    End If
Next
zzCheckTimer.Enabled = True
End Sub

Sub zzSave_Checks()

Dim A1, B1 As Integer
On Error Resume Next
Open App.Path + "\0TextData\" + GetComputerName + "\ChkSettings.txt" For Output As #1

For Each Control In Me.Controls
    Err.Clear
        A1 = 0
        A1 = Control.Value
        A2 = Err.Number
    Err.Clear
        B1 = 0
        B1 = Control.Checked
        B2 = Err.Number
    
    If A2 = 0 Or B2 = 0 Then
        'Debug.Print Control.Name
        Print #1, Control.Name
        Print #1, A1
        Print #1, B1
    End If
Next
Close #1

End Sub



Private Sub zzCheckTimer_Timer()

'A = zzCheckTimer.Interval = 1000
'DEFAULT BIT FREQ 1000



If zzCheckTimer.Interval <> 40000 Then zzCheckTimer.Interval = 40000

On Error Resume Next
checkq = 0
'Debug.Print "--------------------------"

For Each Control In Me.Controls
    If InStr(Control.Name, "Progress") = 0 Then
    'Debug.Print Control.Name
    Err.Clear
        A1 = 0
        A1 = Control.Value
        If Err.Number = 0 Then checkq = checkq Xor (A1)
    Err.Clear
        B1 = 0
        B1 = Control.Checked
        If Err.Number = 0 Then checkq = checkq Xor (B1)
    End If
Next

If checkq = OCheckQ Then Exit Sub
OCheckQ = checkq
Call zzSave_Checks

End Sub

'---------

Public Sub CloseProgram(ByVal Caption As String)

 Hndle = FindWindow(vbNullString, Caption)
 If Hndle = 0 Then Exit Sub
 SendMessage Hndle, WM_CLOSE, 0&, 0&

End Sub

Public Sub CloseProgramHwnd(ByVal Handle As Long)

 If Handle = 0 Then Exit Sub
 SendMessage Handle, WM_CLOSE, 0&, 0&

End Sub


Private Sub Form_Resize()

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
If VAR_BOTTOM_OBJECT = 0 Then
    VAR_BOTTOM_OBJECT = Label_Tune__________.Height + Label_Tune__________.Top
End If
'
'
'
VAR_WIDTH_OBJECT = Label_Tune__________.Left + Label_Tune__________.Width
'If List1.Width > VAR_WIDTH_OBJECT Then VAR_WIDTH_OBJECT = List1.Width

'------------------------------------------------
'LOAD THE FORM TO SIZE OF INNER FORM AND MENU BAR
'------------------------------------------------
'SOMETIME DO OTHER WAY AROUND
'LOAD THE INNER CONTROLS TO SIZE OUTER FORM
'------------------------------------------------

Me.Width = VAR_WIDTH_OBJECT + 100

Me.Height = ((VAR_BOTTOM_OBJECT + (Menu_Height * Screen.TwipsPerPixelY)) + 410)

If OLD_MENU_HEIGHT <> Menu_Height Then
    OldWState = -10
    OLD_MENU_HEIGHT = Menu_Height
End If

If VARCENTER = True Then Exit Sub

    'Me.Top = Screen.Height / 2 - Me.Height / 2 '+400
    'Me.Left = Screen.Width / 2 - Me.Width / 2 '+400
    
    VARCENTER = True

End Sub


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

Private Function Menu_Height()
 
    MenuInfo.cbSize = Len(MenuInfo)
    
    If GetMenuBarInfo(Me.hWnd, OBJID_MENU, 0, MenuInfo) Then
   
        With MenuInfo.rcBar
       
'            Debug.Print "Left: " & CStr(.Left)
'            Debug.Print "Right: " & CStr(.Right)
'            Debug.Print "Top: " & CStr(.Top)
'            Debug.Print "Bottom: " & CStr(.Bottom)
            'Menu_Height = CStr(.Top) + CStr(.Bottom)
            Menu_Height = CStr(.Bottom) - CStr(.Top)
            'Debug.Print Menu_Height

        End With
       
    End If
   
End Function
'------------------------------------------

'-------------------------------------------------
'-------------------------------------------------
Private Sub Timer_1_MINUTE_Timer()

Timer_1_MINUTE.Interval = 60000

Call VB_PROJECT_CHECKDATE

End Sub
'-------------------------------------------------
'-------------------------------------------------
Sub VB_PROJECT_CHECKDATE()

'TOP DELCLARE -------------
'DIM XVB_DATE
'--------------------------

If Mid(App.Path, 1, 1) <> "D" Then Exit Sub

PATH_FILE_NAME1 = App.Path + "\" + App.EXEName + ".EXE"
PATH_FILE_NAME2 = Replace(PATH_FILE_NAME1, "D:\VB6\", "D:\VB6-EXE'S\")

'---------------------------------------------------
'IF A NEW PROJECT NOT BEEN SYNC YET TO MIRROR FOLDER
'---------------------------------------------------
'MINOR WORK AROUND CREATE PATH
'---------------------------------------------------
If Dir(PATH_FILE_NAME2) = "" Then
    On Error Resume Next
        MkDir Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\"))
        FS.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
    If Err.Number > 0 Then Exit Sub
End If

Set F = FS.GetFile(PATH_FILE_NAME2)

VB_DATE = F.DateLastModified

'--------
'01 OF 02
'----------------------------------
'WRITE INFO ABOUT DATE CHANGE NEWER
'----------------------------------

If XVB_DATE < VB_DATE And XVB_DATE > 0 Then
    
    '-----------------------------------
    'RUN A VBS TO COPY OVER AND RELOADER
    '-----------------------------------
    
'    FR1 = FreeFile
'    Open App.Path + "\RELOAD AND COPY.VBS" For Output As #FR1


    '----------------------------------
    'Mon 10 April 2017 16:30:50--------
    '----------------------------------
    'PROJECT REFERANCE ----------------
    'wshom.ocx
    'IF HARD FIND DO BROWSER
    'IN SYSTEM32 AND SYSWOW64
    'DIR wshom.ocx /S
    '----------------------------------
    'WINDOWS SCRIPT HOST OBJECT MODEL
    'AFTER FIND ALSO HAVE TO SELECT HER
    '----------------------------------
    '----------------------------------
    'Mon 10 April 2017 15:45:11--------
    '----
    'VISUAL BASIC wshom.ocx NOT WORKING - Google Search
    'https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB721GB721&q=VISUAL+BASIC+wshom.ocx+NOT+WORKING&oq=VIUAL+BASIC+wshom.ocx+NOT+WORKING&gs_l=serp.3..30i10k1.1533.3987.0.4658.12.12.0.0.0.0.127.888.11j1.12.0....0...1c.1.64.serp..0.12.886...33i160k1j33i21k1.fYCPCpVPeVk
    '--------
    'WScript reference in VB
    'http://forums.codeguru.com/showthread.php?30458-WScript-reference-in-VB
    '----
    'Set objShell = WScript.CreateObject("WScript.Shell") - ------Not WORKER
    '----------------------------------
    '----------------------------------
    
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    
    WSHShell.Run """" + App.Path + "\RELOAD AND COPY.VBS" + """"
    Set WSHShell = Nothing
    Unload Me
    
End If

XVB_DATE = VB_DATE

End Sub


