VERSION 5.00
Object = "{DEEABA2C-9ED3-11D4-92CF-9E7BD6A8DC73}#1.0#0"; "MediaPlayList.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "mciPlayer"
   ClientHeight    =   19995
   ClientLeft      =   300
   ClientTop       =   0
   ClientWidth     =   7005
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFC0&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000040&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   19995
   ScaleWidth      =   7005
   Begin MediaPlaylist.ViewText vtKBPS 
      Height          =   90
      Left            =   2520
      TabIndex        =   75
      Top             =   2520
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   159
      Leds            =   3
      DoubleSize      =   0   'False
      Picture         =   "frmMain.frx":0442
   End
   Begin MediaPlaylist.ViewText vtHZ 
      Height          =   90
      Left            =   2880
      TabIndex        =   74
      Top             =   2520
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   159
      Leds            =   2
      DoubleSize      =   0   'False
      Picture         =   "frmMain.frx":257C
   End
   Begin MediaPlaylist.ViewText vtTimer 
      Height          =   90
      Left            =   3120
      TabIndex        =   73
      Top             =   2520
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   159
      Leds            =   5
      DoubleSize      =   0   'False
      Picture         =   "frmMain.frx":46B6
   End
   Begin MediaPlaylist.ViewText ScrollTitle1 
      Height          =   90
      Left            =   120
      TabIndex        =   72
      Top             =   2520
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   159
      Leds            =   31
      DoubleSize      =   0   'False
      Picture         =   "frmMain.frx":67F0
   End
   Begin MediaPlaylist.TimeDisplay TimeDisplay1 
      Height          =   195
      Left            =   2160
      TabIndex        =   71
      Top             =   2160
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   344
      Leds            =   4
      DoubleSize      =   0   'False
      Minimized       =   0   'False
      Twinkle         =   0   'False
      Picture         =   "frmMain.frx":892A
   End
   Begin MediaPlaylist.Balance Balance1 
      Height          =   195
      Left            =   1320
      TabIndex        =   70
      Top             =   2160
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   344
      DoubleSize      =   0   'False
      Picture         =   "frmMain.frx":98B8
   End
   Begin MediaPlaylist.Volume Volume1 
      Height          =   195
      Left            =   120
      TabIndex        =   69
      Top             =   2160
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   344
      DoubleSize      =   0   'False
      Picture         =   "frmMain.frx":1F216
   End
   Begin MediaPlaylist.TimeSlider TimeSlider1 
      Height          =   150
      Left            =   120
      TabIndex        =   68
      Top             =   1920
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   265
      DoubleFormat    =   0   'False
      Picture         =   "frmMain.frx":34B74
   End
   Begin VB.PictureBox PicSrcText 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4560
      Picture         =   "frmMain.frx":36FDE
      ScaleHeight     =   270
      ScaleWidth      =   2325
      TabIndex        =   67
      Top             =   9600
      Width           =   2325
   End
   Begin VB.PictureBox stdText 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4560
      Picture         =   "frmMain.frx":3910A
      ScaleHeight     =   270
      ScaleWidth      =   2325
      TabIndex        =   66
      Top             =   9960
      Width           =   2325
   End
   Begin VB.PictureBox stdVolume 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   240
      Picture         =   "frmMain.frx":3B234
      ScaleHeight     =   6495
      ScaleWidth      =   1020
      TabIndex        =   65
      Top             =   12240
      Width           =   1020
   End
   Begin VB.Timer ScrollTimer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3600
      Top             =   7320
   End
   Begin VB.PictureBox stdMain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   120
      Picture         =   "frmMain.frx":50B82
      ScaleHeight     =   1740
      ScaleWidth      =   4125
      TabIndex        =   64
      Top             =   8520
      Width           =   4125
   End
   Begin VB.PictureBox stdMonoSter 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   5640
      Picture         =   "frmMain.frx":682F4
      ScaleHeight     =   360
      ScaleWidth      =   870
      TabIndex        =   63
      Top             =   8280
      Width           =   870
   End
   Begin VB.PictureBox stdPosBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   120
      Picture         =   "frmMain.frx":693B6
      ScaleHeight     =   150
      ScaleWidth      =   4605
      TabIndex        =   62
      Top             =   8040
      Width           =   4605
   End
   Begin VB.PictureBox stdStatus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   240
      Picture         =   "frmMain.frx":6B810
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   61
      Top             =   7680
      Width           =   630
   End
   Begin VB.PictureBox stdTime 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   240
      Picture         =   "frmMain.frx":6BCD2
      ScaleHeight     =   195
      ScaleWidth      =   1485
      TabIndex        =   60
      Top             =   7440
      Width           =   1485
   End
   Begin VB.PictureBox stdShufRep 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   4800
      Picture         =   "frmMain.frx":6CC50
      ScaleHeight     =   1275
      ScaleWidth      =   1380
      TabIndex        =   59
      Top             =   5760
      Width           =   1380
   End
   Begin VB.PictureBox stdTitlebar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   120
      Picture         =   "frmMain.frx":72836
      ScaleHeight     =   1305
      ScaleWidth      =   5160
      TabIndex        =   58
      Top             =   4200
      Width           =   5160
   End
   Begin VB.PictureBox stdButton 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   2520
      Picture         =   "frmMain.frx":88730
      ScaleHeight     =   540
      ScaleWidth      =   2040
      TabIndex        =   57
      Top             =   3600
      Width           =   2040
   End
   Begin VB.PictureBox PicRecording 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   1200
      ScaleHeight     =   135
      ScaleWidth      =   45
      TabIndex        =   56
      Top             =   7680
      Width           =   45
   End
   Begin VB.PictureBox PicStatus 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   960
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   55
      Top             =   7680
      Width           =   135
   End
   Begin VB.PictureBox PicSrcStatus 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   240
      ScaleHeight     =   135
      ScaleWidth      =   630
      TabIndex        =   54
      Top             =   7680
      Width           =   630
   End
   Begin VB.PictureBox PicSrcTime 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   1485
      TabIndex        =   53
      Top             =   7440
      Width           =   1485
   End
   Begin VB.PictureBox PicStereo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   6120
      ScaleHeight     =   180
      ScaleWidth      =   420
      TabIndex        =   52
      Top             =   8760
      Width           =   420
   End
   Begin VB.PictureBox PicMono 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   5640
      ScaleHeight     =   180
      ScaleWidth      =   405
      TabIndex        =   51
      Top             =   8760
      Width           =   405
   End
   Begin VB.PictureBox PicSrcMonoSter 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   5640
      ScaleHeight     =   360
      ScaleWidth      =   870
      TabIndex        =   50
      Top             =   8280
      Width           =   870
   End
   Begin VB.Timer DisplayTimer 
      Enabled         =   0   'False
      Interval        =   490
      Left            =   4320
      Top             =   6360
   End
   Begin VB.PictureBox PicPlayList 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   3
      Left            =   6600
      ScaleHeight     =   180
      ScaleWidth      =   345
      TabIndex        =   49
      Top             =   7920
      Width           =   345
   End
   Begin VB.PictureBox PicPlayList 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   2
      Left            =   6600
      ScaleHeight     =   180
      ScaleWidth      =   345
      TabIndex        =   48
      Top             =   7680
      Width           =   345
   End
   Begin VB.PictureBox PicPlayList 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   6600
      ScaleHeight     =   180
      ScaleWidth      =   345
      TabIndex        =   47
      Top             =   7440
      Width           =   345
   End
   Begin VB.PictureBox PicPlayList 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   0
      Left            =   6600
      ScaleHeight     =   180
      ScaleWidth      =   345
      TabIndex        =   46
      Top             =   7200
      Width           =   345
   End
   Begin VB.PictureBox PicEQ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   3
      Left            =   6120
      ScaleHeight     =   180
      ScaleWidth      =   345
      TabIndex        =   45
      Top             =   7920
      Width           =   345
   End
   Begin VB.PictureBox PicEQ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   2
      Left            =   6120
      ScaleHeight     =   180
      ScaleWidth      =   345
      TabIndex        =   44
      Top             =   7680
      Width           =   345
   End
   Begin VB.PictureBox PicEQ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   6120
      ScaleHeight     =   180
      ScaleWidth      =   345
      TabIndex        =   43
      Top             =   7440
      Width           =   345
   End
   Begin VB.PictureBox PicEQ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   0
      Left            =   6120
      ScaleHeight     =   180
      ScaleWidth      =   345
      TabIndex        =   42
      Top             =   7200
      Width           =   345
   End
   Begin VB.PictureBox PicShuffle 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   3
      Left            =   5280
      ScaleHeight     =   225
      ScaleWidth      =   705
      TabIndex        =   41
      Top             =   7920
      Width           =   705
   End
   Begin VB.PictureBox PicShuffle 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   2
      Left            =   5280
      ScaleHeight     =   225
      ScaleWidth      =   705
      TabIndex        =   40
      Top             =   7680
      Width           =   705
   End
   Begin VB.PictureBox PicShuffle 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   1
      Left            =   5280
      ScaleHeight     =   225
      ScaleWidth      =   705
      TabIndex        =   39
      Top             =   7440
      Width           =   705
   End
   Begin VB.PictureBox PicShuffle 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   0
      Left            =   5280
      ScaleHeight     =   225
      ScaleWidth      =   705
      TabIndex        =   38
      Top             =   7200
      Width           =   705
   End
   Begin VB.PictureBox PicRepeat 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   3
      Left            =   4800
      ScaleHeight     =   225
      ScaleWidth      =   420
      TabIndex        =   37
      Top             =   7920
      Width           =   420
   End
   Begin VB.PictureBox PicRepeat 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   2
      Left            =   4800
      ScaleHeight     =   225
      ScaleWidth      =   420
      TabIndex        =   36
      Top             =   7680
      Width           =   420
   End
   Begin VB.PictureBox PicRepeat 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   1
      Left            =   4800
      ScaleHeight     =   225
      ScaleWidth      =   420
      TabIndex        =   35
      Top             =   7440
      Width           =   420
   End
   Begin VB.PictureBox PicRepeat 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   0
      Left            =   4800
      ScaleHeight     =   225
      ScaleWidth      =   420
      TabIndex        =   34
      Top             =   7200
      Width           =   420
   End
   Begin VB.PictureBox PicSrcShufRep 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   4800
      ScaleHeight     =   1275
      ScaleWidth      =   1380
      TabIndex        =   33
      Top             =   5760
      Width           =   1380
   End
   Begin VB.PictureBox PicMain 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   120
      ScaleHeight     =   1740
      ScaleWidth      =   4125
      TabIndex        =   32
      Top             =   10320
      Width           =   4125
   End
   Begin VB.PictureBox PicSrcMain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   120
      ScaleHeight     =   1740
      ScaleWidth      =   4125
      TabIndex        =   31
      Top             =   8520
      Width           =   4125
   End
   Begin VB.PictureBox PicSrcPosbar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   120
      ScaleHeight     =   150
      ScaleWidth      =   4605
      TabIndex        =   30
      Top             =   8040
      Width           =   4605
   End
   Begin VB.PictureBox PicTitleOptEmpty 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   6360
      ScaleHeight     =   600
      ScaleWidth      =   120
      TabIndex        =   29
      Top             =   4200
      Width           =   120
   End
   Begin VB.PictureBox PicTitleOpt 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   4
      Left            =   6480
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   28
      Top             =   5400
      Width           =   120
   End
   Begin VB.PictureBox PicTitleOpt 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   3
      Left            =   6360
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   27
      Top             =   5280
      Width           =   120
   End
   Begin VB.PictureBox PicTitleOpt 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   2
      Left            =   6240
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   26
      Top             =   5160
      Width           =   120
   End
   Begin VB.PictureBox PicTitleOpt 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   1
      Left            =   6120
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   25
      Top             =   5040
      Width           =   120
   End
   Begin VB.PictureBox PicTitleOpt 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   0
      Left            =   6000
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   24
      Top             =   4920
      Width           =   120
   End
   Begin VB.PictureBox PicTitleBut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   9
      Left            =   5640
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   23
      Top             =   4920
      Width           =   135
   End
   Begin VB.PictureBox PicTitleBut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   8
      Left            =   5400
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   22
      Top             =   4920
      Width           =   135
   End
   Begin VB.PictureBox PicTitleBut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   7
      Left            =   5640
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   21
      Top             =   4680
      Width           =   135
   End
   Begin VB.PictureBox PicTitleBut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   6
      Left            =   5400
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   20
      Top             =   4680
      Width           =   135
   End
   Begin VB.PictureBox PicTitleBut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   5
      Left            =   5880
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   19
      Top             =   4440
      Width           =   135
   End
   Begin VB.PictureBox PicTitleBut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   4
      Left            =   5640
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   18
      Top             =   4440
      Width           =   135
   End
   Begin VB.PictureBox PicTitleBut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   3
      Left            =   5400
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   17
      Top             =   4440
      Width           =   135
   End
   Begin VB.PictureBox PicTitleBut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   2
      Left            =   5880
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   16
      Top             =   4200
      Width           =   135
   End
   Begin VB.PictureBox PicTitleBut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   1
      Left            =   5640
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   15
      Top             =   4200
      Width           =   135
   End
   Begin VB.PictureBox PicTitleBut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   0
      Left            =   5400
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   14
      Top             =   4200
      Width           =   135
   End
   Begin VB.PictureBox PicTitleBar 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   5
      Left            =   120
      ScaleHeight     =   210
      ScaleWidth      =   4125
      TabIndex        =   13
      Top             =   6960
      Width           =   4125
   End
   Begin VB.PictureBox PicTitleBar 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   4
      Left            =   120
      ScaleHeight     =   210
      ScaleWidth      =   4125
      TabIndex        =   12
      Top             =   6720
      Width           =   4125
   End
   Begin VB.PictureBox PicTitleBar 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   3
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   4125
      TabIndex        =   11
      Top             =   6480
      Width           =   4125
   End
   Begin VB.PictureBox PicTitleBar 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   2
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   4125
      TabIndex        =   10
      Top             =   6240
      Width           =   4125
   End
   Begin VB.PictureBox PicTitleBar 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   1
      Left            =   120
      ScaleHeight     =   210
      ScaleWidth      =   4125
      TabIndex        =   9
      Top             =   6015
      Width           =   4125
   End
   Begin VB.PictureBox PicTitleBar 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   0
      Left            =   120
      ScaleHeight     =   210
      ScaleWidth      =   4125
      TabIndex        =   8
      Top             =   5790
      Width           =   4125
   End
   Begin VB.PictureBox PicSrcTitlebar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   120
      ScaleHeight     =   1305
      ScaleWidth      =   5160
      TabIndex        =   7
      Top             =   4200
      Width           =   5160
   End
   Begin VB.PictureBox PicEject 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1920
      ScaleHeight     =   240
      ScaleWidth      =   330
      TabIndex        =   6
      Top             =   3600
      Width           =   330
   End
   Begin VB.PictureBox PicNext 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1560
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   5
      Top             =   3600
      Width           =   330
   End
   Begin VB.PictureBox PicStop 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1200
      ScaleHeight     =   270
      ScaleWidth      =   345
      TabIndex        =   4
      Top             =   3600
      Width           =   345
   End
   Begin VB.PictureBox PicPause 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   840
      ScaleHeight     =   270
      ScaleWidth      =   345
      TabIndex        =   3
      Top             =   3600
      Width           =   345
   End
   Begin VB.PictureBox PicPlay 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   480
      ScaleHeight     =   270
      ScaleWidth      =   345
      TabIndex        =   2
      Top             =   3600
      Width           =   345
   End
   Begin VB.PictureBox PicPrevious 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   345
      TabIndex        =   1
      Top             =   3600
      Width           =   345
   End
   Begin VB.PictureBox PicSrcButton 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   2520
      ScaleHeight     =   540
      ScaleWidth      =   2040
      TabIndex        =   0
      Top             =   3600
      Width           =   2040
   End
   Begin VB.Timer PauseTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5640
      Top             =   9120
   End
   Begin VB.Image imgSEject 
      Enabled         =   0   'False
      Height          =   105
      Left            =   3180
      ToolTipText     =   "Open CD Tray"
      Top             =   45
      Width           =   135
   End
   Begin VB.Image imgSNext 
      Enabled         =   0   'False
      Height          =   105
      Left            =   3045
      ToolTipText     =   "Next Track"
      Top             =   45
      Width           =   135
   End
   Begin VB.Image imgSStop 
      Enabled         =   0   'False
      Height          =   105
      Left            =   2910
      ToolTipText     =   "Stop"
      Top             =   45
      Width           =   135
   End
   Begin VB.Image imgSPause 
      Enabled         =   0   'False
      Height          =   105
      Left            =   2775
      ToolTipText     =   "Pause"
      Top             =   45
      Width           =   135
   End
   Begin VB.Image imgSPlay 
      Enabled         =   0   'False
      Height          =   105
      Left            =   2640
      ToolTipText     =   "Play"
      Top             =   45
      Width           =   135
   End
   Begin VB.Image imgSPrev 
      Enabled         =   0   'False
      Height          =   105
      Left            =   2505
      ToolTipText     =   "Previous Track"
      Top             =   45
      Width           =   135
   End
   Begin VB.Image imgAllOnTop 
      Height          =   120
      Left            =   150
      ToolTipText     =   "Toggle Always On Top"
      Top             =   480
      Width           =   120
   End
   Begin VB.Image imgMono 
      Height          =   180
      Left            =   3180
      Top             =   615
      Width           =   405
   End
   Begin VB.Image imgStereo 
      Height          =   180
      Left            =   3585
      Top             =   615
      Width           =   420
   End
   Begin VB.Image imgVisual 
      Height          =   120
      Left            =   150
      ToolTipText     =   "Visualization Menu"
      Top             =   840
      Width           =   120
   End
   Begin VB.Image imgFileInfo 
      Height          =   120
      Left            =   150
      ToolTipText     =   "File Info Box"
      Top             =   600
      Width           =   120
   End
   Begin VB.Image imgOptions 
      Height          =   120
      Left            =   150
      ToolTipText     =   "Options"
      Top             =   360
      Width           =   120
   End
   Begin VB.Image imgPlayList 
      Height          =   180
      Left            =   3630
      ToolTipText     =   "Toggle Playlist Editor"
      Top             =   870
      Width           =   345
   End
   Begin VB.Image imgEQ 
      Height          =   180
      Left            =   3285
      ToolTipText     =   "Toggle Graphical Equalizer"
      Top             =   870
      Width           =   345
   End
   Begin VB.Image imgRepeat 
      Height          =   225
      Left            =   3150
      ToolTipText     =   "Toggle Repeat"
      Top             =   1335
      Width           =   420
   End
   Begin VB.Image imgShuffle 
      Height          =   225
      Left            =   2460
      ToolTipText     =   "Toggle Shuffle"
      Top             =   1335
      Width           =   705
   End
   Begin VB.Image imgEject 
      Height          =   255
      Left            =   2040
      ToolTipText     =   "Open CD Tray"
      Top             =   1335
      Width           =   345
   End
   Begin VB.Image imgDouble 
      Height          =   120
      Left            =   150
      ToolTipText     =   "Toggle Double Size Mode"
      Top             =   720
      Width           =   120
   End
   Begin VB.Image ImgMenu 
      Height          =   135
      Left            =   0
      ToolTipText     =   "Winamp Menu"
      Top             =   0
      Width           =   135
   End
   Begin VB.Image ImgTitleBarKnop3 
      Height          =   135
      Left            =   3855
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   135
   End
   Begin VB.Image ImgTitleBarKnop2 
      Height          =   135
      Left            =   3705
      ToolTipText     =   "Toggle Windowshade Mode"
      Top             =   0
      Width           =   135
   End
   Begin VB.Image ImgTitleBarKnop1 
      Height          =   135
      Left            =   3555
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   135
   End
   Begin VB.Image ImgTitleOptEmpty 
      Height          =   600
      Left            =   150
      Top             =   360
      Width           =   120
   End
   Begin VB.Image ImgTitleBar 
      Height          =   210
      Left            =   0
      Top             =   0
      Width           =   4125
   End
   Begin VB.Image imgNext 
      Height          =   255
      Left            =   1605
      ToolTipText     =   "Next Track"
      Top             =   1335
      Width           =   345
   End
   Begin VB.Image imgStop 
      Height          =   255
      Left            =   1260
      ToolTipText     =   "Stop"
      Top             =   1335
      Width           =   345
   End
   Begin VB.Image imgPause 
      Height          =   255
      Left            =   915
      ToolTipText     =   "Pause"
      Top             =   1335
      Width           =   345
   End
   Begin VB.Image imgPlay 
      Height          =   255
      Left            =   585
      ToolTipText     =   "Play"
      Top             =   1335
      Width           =   345
   End
   Begin VB.Image imgPrevious 
      Height          =   255
      Left            =   240
      ToolTipText     =   "Previous Track"
      Top             =   1335
      Width           =   345
   End
   Begin VB.Image ImgMain 
      Height          =   1740
      Left            =   0
      Top             =   0
      Width           =   4125
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public stndInfo, stndVisual As Boolean
Dim TimeHasPlayed As Long
Public Minimized As Boolean
Dim Recording As Boolean
Public BigForm As Boolean, stndOptions As Boolean
Dim MoveForm As Boolean, MoveFirst As Boolean
Public PlayByStart As Boolean
Public QuitmciPlayer As Boolean
Dim sOldTitle As String, stateTitle As String
Dim Message As String
Dim i As Integer
Dim introTimer, moveTimer
Dim FX, FY, AX, AY
Public inTitleBar As Integer
Public Playing As Boolean
Public EqualizerHangsL As Boolean
Public EqualizerHangsT As Boolean
Public PlayListHangsL As Boolean
Public PlayListHangsT As Boolean
Public stndPlayList As Integer
Public stndEQ As Integer
Public stndOnTop As Boolean
Public intTimeMode As Integer

Private Sub DisplayTimer_Timer()
    DoEvents
    If frmPlayList.Visible = True Then SetVT:
    Dim intSec As Integer, intMin As Integer, intHour As Integer
    If TimeSlider1.SetStartAt = False Then
        Select Case intTimeMode
            Case Is = 0: TimeHasPlayed = frmPlayList.Playlist1.TimeElapsed:
            Case Is = 1: TimeHasPlayed = frmPlayList.Playlist1.TimeRemaining:
            Case Is = 2: TimeHasPlayed = frmPlayList.Playlist1.TimeSong:
            Case Is = 3: TimeHasPlayed = frmPlayList.Playlist1.TimeTotalElapsed: intHour = Int(TimeHasPlayed / 3600):
            Case Is = 4: TimeHasPlayed = frmPlayList.Playlist1.TimeTotalRemaining: intHour = Int(TimeHasPlayed / 3600):
            Case Is = 5: TimeHasPlayed = frmPlayList.Playlist1.TimeTotal: intHour = Int(TimeHasPlayed / 3600):
        End Select
    Else
        TimeHasPlayed = TimeSlider1.StartAt
    End If
    If frmPlayList.Playlist1.Status = "Paused" Then
        TimeDisplay1.Twinkle = True
    Else
        TimeDisplay1.Twinkle = False
    End If
    TimeDisplay1.SetTime TimeHasPlayed
    TimeSlider1.TimeMaximum = frmPlayList.Playlist1.TimeSong
    TimeSlider1.TimeElapsed = frmPlayList.Playlist1.TimeElapsed
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim t#
    If Me.WindowState = 1 Then Exit Sub:
    If Minimized = True Then Exit Sub:
    Message = "mciPlayer v.1.01" & vbCrLf & vbCrLf & " Design By Patrick"
    Select Case KeyCode
        Case vbKeyA: If Shift = 2 Then imgAllOnTop_Click:
        Case vbKeyD: If Shift = 2 Then imgDouble_Click:
        Case vbKeyR: imgRepeat_MouseDown 1, 0, 210, 105: t# = Timer + 0.5: Do Until Timer >= t#: DoEvents: Loop: imgRepeat_MouseUp 1, 0, 210, 105:
        Case vbKeyS: imgShuffle_MouseDown 1, 0, 345, 165: t# = Timer + 0.5: Do Until Timer >= t#: DoEvents: Loop: imgShuffle_MouseUp 1, 0, 345, 165:
        Case vbKeyT: If Shift = 2 Then PicTime_Click 1:
        Case vbKeyF1: MsgBox "F1 = Under Construction !":
        Case vbKeyF2: MsgBox Message, vbOKOnly, " About":
        Case vbKeyF4: MsgBox "Made with Visual Basic", vbOKOnly, " About":
        Case vbKeyF5: MsgBox "Made with Visual Basic", vbOKOnly, " About":
        Case vbKeyF9: MsgBox Message, vbOKOnly, " About":
        Case vbKeyF10: ImgTitleBarKnop3_MouseDown 1, 0, 60, 45:
    End Select
End Sub

Private Sub Form_Load()
    If (App.PrevInstance = True) Then
        End
    End If
    Dim a
    Dim errNumber As Integer
    a = GetSetting(App.EXEName, "Mode", "StartPlay", False)
    PlayByStart = IIf(CBool(a), a, False)
    If frmPlayList.Playlist1.CheckSoundCard = False Then End:
    If GetSetting(App.EXEName, "Settings", "Shortcut", "False") = False Then frmPlayList.comFunc1.CreateShortcut App.EXEName, App.Path, stDesktop_Program_NewGroup:
    a = GetSetting(App.EXEName, "Mode", "Timer", 0)
    intTimeMode = IIf(IsNumeric(a) = True, a, 0)
    On Error GoTo err_Handle:
    errNumber = 1
    a = GetSetting(App.EXEName, "Visual", "Big", False)
    BigForm = IIf(CBool(a), a, False)
    errNumber = 2
    Minimized = False
    imgPositions
    frmPlayList.SetHoleView
    imgCorrections
    WijzigFormaat
    Playing = False
    PlayListHangsT = True
    EqualizerHangsT = True
    a = GetSetting(App.EXEName, "Visual", "StartPos", "300,300")
    For i = 1 To Len(a)
        If Mid(a, i, 1) = "," Then Exit For:
        If Mid(a, i, 1) = ";" Then Exit For:
    Next
    Me.Left = IIf(IsNumeric(Left$(a, i - 1)) = True, Left(a, i - 1), 300)
    errNumber = 3
    Me.Top = IIf(IsNumeric(Right$(a, Len(a) - i - 1)) = True, Right(a, Len(a) - i), 300)
    a = GetSetting(App.EXEName, "Visual", "Equalizer", 0)
    stndEQ = IIf(IsNumeric(a), a, 0)
    If stndEQ <> 0 And stndEQ <> 1 Then stndEQ = 0:
    If stndEQ = 1 Then imgEQ.Picture = PicEQ(2).Image: frmEqualizer.Show:
    frmEqualizer.Top = Me.Top + Me.Height
    frmEqualizer.Left = Me.Left
    a = GetSetting(App.EXEName, "Visual", "Playlist", 0)
    stndPlayList = IIf(IsNumeric(a), a, 0)
    If stndPlayList <> 0 And stndPlayList <> 1 Then stndPlayList = 0:
    If stndPlayList = 1 Then imgPlayList.Picture = PicPlayList(2).Image: frmPlayList.Show:
    frmPlayList.Top = Me.Top + Me.Height + IIf(stndEQ = 1, frmEqualizer.Height, 0)
    frmPlayList.Left = Me.Left
    KeyPreview = True
    TimeDisplay1.SetTime 0
    If PlayByStart = True Then imgPlay_MouseUp 1, 0, 90, 105:
    Exit Sub
err_Handle:
    If errNumber = 1 Then
        BigForm = False
    ElseIf errNumber = 2 Then
        'nothing
    ElseIf errNumber = 3 Then
        Me.Top = 300
    End If
    Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.EXEName, "Visual", "Big", BigForm
    SaveSetting App.EXEName, "Visual", "Playlist", stndPlayList
    SaveSetting App.EXEName, "Visual", "Equalizer", stndEQ
    SaveSetting App.EXEName, "Visual", "StartPos", Str(Me.Left) & "," & Str(Me.Top)
    SaveSetting App.EXEName, "Mode", "Timer", intTimeMode
    SaveSetting App.EXEName, "Mode", "StartPlay", PlayByStart
End Sub

Private Sub imgAllOnTop_Click()
    If stndOnTop = False Then
        imgAllOnTop.Picture = PicTitleOpt(1).Image
        stndOnTop = True
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
        SetWindowPos frmPlayList.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    Else
        Set imgAllOnTop.Picture = Nothing
        stndOnTop = False
        SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
        SetWindowPos frmPlayList.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    End If
End Sub

Private Sub imgDouble_Click()
    If BigForm = True Then
        BigForm = False
        Set imgDouble.Picture = Nothing
    Else
        BigForm = True
        imgDouble.Picture = PicTitleOpt(3).Image
    End If
    WijzigFormaat
    frmEqualizer.EQFormat
    If stndEQ = 1 Then
        frmEqualizer.Top = Me.Top + Me.Height
        frmEqualizer.Left = Me.Left
        If stndPlayList = 1 Then
            frmPlayList.Top = frmEqualizer.Top + frmEqualizer.Height
            frmPlayList.Left = Me.Left
        End If
    Else
        If stndPlayList = 1 Then
            frmPlayList.Top = Me.Top + Me.Height
            frmPlayList.Left = Me.Left
        End If
    End If
End Sub

Private Sub imgEject_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage PicEject, 0, 0, 22, 16, PicSrcButton, 114, 16
End Sub

Private Sub imgEject_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage PicEject, 0, 0, 22, 16, PicSrcButton, 114, 0
    frmPlayList.Playlist1.AddFile
End Sub

Private Sub imgEQ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If stndEQ = 0 Then
        imgEQ.Picture = PicEQ(3).Image
    Else
        imgEQ.Picture = PicEQ(1).Image
    End If
End Sub

Private Sub imgEQ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If stndEQ = 0 Then
        imgEQ.Picture = PicEQ(2).Image
        frmEqualizer.Top = Me.Top + Me.Height
        frmEqualizer.Left = Me.Left
        frmEqualizer.EQFormat
        frmEqualizer.Show
        If PlayListHangsT = True Then frmPlayList.Top = Me.Top + Me.Height + frmEqualizer.Height:
        stndEQ = 1
    Else
        imgEQ.Picture = PicEQ(0).Image
        frmEqualizer.Hide
        If PlayListHangsT = True Then frmPlayList.Top = Me.Top + Me.Height:
        stndEQ = 0
    End If
End Sub

Public Sub imgFileInfo_Click()
    If stndInfo = False Then
        stndInfo = True
        imgFileInfo.Picture = PicTitleOpt(2).Image
        frmPlayList.Playlist1.FileInfo
    Else
        stndInfo = False
        Set imgFileInfo.Picture = Nothing
    End If
End Sub

Private Sub ImgMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgMenu.Picture = PicTitleBut(3).Image
End Sub

Private Sub ImgMenu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgMenu.Picture = PicTitleBut(0).Image
    imgOptions_Click
End Sub

Private Sub imgNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage PicNext, 0, 0, 20, 16, PicSrcButton, 92, 18
End Sub

Private Sub imgNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage PicNext, 0, 0, 22, 19, PicSrcButton, 92, 0
    frmPlayList.Playlist1.EventNext
End Sub

Private Sub imgOptions_Click()
    If stndOptions = False Then
        imgOptions.Picture = PicTitleOpt(0).Image
        stndOptions = True
        frmOptions.Show
    Else
        Set imgOptions.Picture = Nothing
        stndOptions = False
        Unload frmOptions
    End If
End Sub

Private Sub imgPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage PicPause, 0, 0, 23, 19, PicSrcButton, 46, 18
End Sub

Private Sub imgPause_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage PicPause, 0, 0, 23, 19, PicSrcButton, 46, 0
    frmPlayList.Playlist1.EventPause
    ScrollTimer.Enabled = True
    DisplayTimer.Enabled = True
End Sub

Private Sub imgPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage PicPlay, 0, 0, 23, 19, PicSrcButton, 23, 18
End Sub

Private Sub imgPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage PicPlay, 0, 0, 23, 19, PicSrcButton, 23, 0
    frmPlayList.Playlist1.EventPlay
    ScrollTimer.Enabled = True
    DisplayTimer.Enabled = True
End Sub

Private Sub imgPlayList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If stndPlayList = 0 Then
        imgPlayList.Picture = PicPlayList(3).Image
    Else
        imgPlayList.Picture = PicPlayList(1).Image
    End If
End Sub

Private Sub imgPlayList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If frmPlayList.Playlist1.IsMovie(frmPlayList.Playlist1.CurrentFile) And Playing = True Then
        If stndPlayList = 0 Then
            stndPlayList = 1
            imgPlayList.Picture = PicPlayList(2).Image
        Else
            stndPlayList = 0
            imgPlayList.Picture = PicPlayList(0).Image
        End If
        Exit Sub
    End If
    If stndPlayList = 0 Then
        imgPlayList.Picture = PicPlayList(2).Image
        stndPlayList = 1
        frmPlayList.Top = Me.Top + Me.Height + IIf(stndEQ = 1, frmEqualizer.Height, 0)
        frmPlayList.Left = Me.Left
        frmPlayList.Show
    Else
        imgPlayList.Picture = PicPlayList(0).Image
        stndPlayList = 0
        frmPlayList.Hide
    End If
End Sub

Private Sub imgPrevious_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage PicPrevious, 0, 0, 23, 19, PicSrcButton, 0, 18
End Sub

Private Sub imgPrevious_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage PicPrevious, 0, 0, 23, 19, PicSrcButton, 0, 0
    frmPlayList.Playlist1.EventPrevious
End Sub

Private Sub imgRepeat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If frmPlayList.Playlist1.Repeat = 0 Then
        imgRepeat.Picture = PicRepeat(3).Image
        Message = "Repeat 1"
    ElseIf frmPlayList.Playlist1.Repeat = 1 Then
        imgRepeat.Picture = PicRepeat(3).Image
        Message = "Repeat All"
    Else
        imgRepeat.Picture = PicRepeat(1).Image
        Message = "Repeat Off"
    End If
    ScrollTimer.Enabled = False
    ScrollTitle1.ShowMessage Message
End Sub

Private Sub imgRepeat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If frmPlayList.Playlist1.Repeat = 0 Then
        imgRepeat.Picture = PicRepeat(2).Image
        frmPlayList.Playlist1.Repeat = 1
    ElseIf frmPlayList.Playlist1.Repeat = 1 Then
        imgRepeat.Picture = PicRepeat(2).Image
        frmPlayList.Playlist1.Repeat = 2
    Else
        imgRepeat.Picture = PicRepeat(0).Image
        frmPlayList.Playlist1.Repeat = 0
    End If
    ScrollTimer.Enabled = True
End Sub

Private Sub imgSEject_Click()
    frmPlayList.Playlist1.AddFile
End Sub

Public Sub imgShuffle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If frmPlayList.Playlist1.ShuffleMode = False Then
        imgShuffle.Picture = PicShuffle(3).Image
        Message = "random mode on"
    Else
        imgShuffle.Picture = PicShuffle(1).Image
        Message = "random mode off"
    End If
    ScrollTimer.Enabled = False
    ScrollTitle1.ShowMessage Message
End Sub

Private Sub imgShuffle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If frmPlayList.Playlist1.ShuffleMode = False Then
        imgShuffle.Picture = PicShuffle(2).Image
        frmPlayList.Playlist1.ShuffleMode = True
    Else
        imgShuffle.Picture = PicShuffle(0).Image
        frmPlayList.Playlist1.ShuffleMode = False
    End If
    ScrollTimer.Enabled = True
End Sub

Private Sub imgSNext_Click()
    frmPlayList.Playlist1.EventNext
End Sub

Private Sub imgSPause_Click()
    frmPlayList.Playlist1.EventPause
    ScrollTimer.Enabled = True
    DisplayTimer.Enabled = True
End Sub

Private Sub imgSPlay_Click()
    frmPlayList.Playlist1.EventPlay
    ScrollTimer.Enabled = True
    DisplayTimer.Enabled = True
End Sub

Private Sub imgSPrev_Click()
    frmPlayList.Playlist1.EventPrevious
End Sub

Private Sub imgSStop_Click()
    frmPlayList.Playlist1.EventStop
End Sub

Private Sub imgStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage PicStop, 0, 0, 23, 19, PicSrcButton, 69, 18
End Sub

Private Sub imgStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage PicStop, 0, 0, 23, 19, PicSrcButton, 69, 0
    frmPlayList.Playlist1.EventStop
End Sub

Private Sub ImgTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MoveForm = False Then
        FX = ImgTitleBar.Left + ImgTitleBar.Width / 2
        FY = ImgTitleBar.Top + ImgTitleBar.Height / 2
        AX = X
        AY = Y
        MoveForm = True
        MoveFirst = True
    End If
End Sub
    
Private Sub ImgTitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MoveForm = True Then
        Dim posW, posH
        posW = Int((FX + X - IIf(MoveFirst = True, X, AX)) / 15) * 15
        posH = Int((FY + Y - IIf(MoveFirst = True, Y, AY)) / 15) * 15
        FX = IIf(MoveFirst = True, FX, posW)
        FY = IIf(MoveFirst = True, FY, posH)
        If MoveFirst = False Then Me.Move FX, FY:
        MoveFirst = False
        Dim posLBC As Integer, posTBC As Integer
        Dim posLBC1 As Integer, posTBC1 As Integer
        posLBC = Me.Left + Me.Width
        posTBC = Me.Top + Me.Height
        posLBC1 = frmEqualizer.Left + frmEqualizer.Width
        posTBC1 = frmEqualizer.Top + frmEqualizer.Height
        If stndEQ = 1 Then
            If EqualizerHangsL = False Then
                If posLBC < frmEqualizer.Left + 120 And posLBC > frmEqualizer.Left - 120 Then
                    If Me.Top < frmEqualizer.Top + 120 And Me.Top > frmEqualizer.Top - 120 Then
                        frmEqualizer.Move (Me.Left + Me.Width), Me.Top
                        EqualizerHangsL = True
                        EqualizerHangsT = False
                    End If
                Else
                    EqualizerHangsL = False
                End If
            Else
                frmEqualizer.Move (Me.Left + Me.Width), Me.Top
            End If
            If EqualizerHangsT = False Then
                If posTBC < frmEqualizer.Top + 120 And posTBC > frmEqualizer.Top - 120 Then
                    If Me.Left < frmEqualizer.Left + 120 And Me.Left > frmEqualizer.Left - 120 Then
                        frmEqualizer.Move Me.Left, (Me.Top + Me.Height)
                        EqualizerHangsT = True
                        EqualizerHangsL = False
                    End If
                Else
                    EqualizerHangsT = False
                End If
            Else
                frmEqualizer.Move Me.Left, (Me.Top + Me.Height)
            End If
        Else
            EqualizerHangsL = False
            EqualizerHangsT = False
        End If
        If stndPlayList = 1 Then
            If EqualizerHangsL = True And stndEQ = 1 Then posLBC = posLBC1:
            If EqualizerHangsT = True And stndEQ = 1 Then posTBC = posTBC1:
            If PlayListHangsL = False Then
                If posLBC < frmPlayList.Left + 120 And posLBC > frmPlayList.Left - 120 Then
                    If posTBC < frmPlayList.Top + 120 And posTBC > frmPlayList.Top - 120 Then
                        frmPlayList.Move (Me.Left + Me.Width + IIf(EqualizerHangsL = True, frmEqualizer.Width, 0)), Me.Top
                        PlayListHangsL = True
                        PlayListHangsT = False
                    ElseIf EqualizerHangsL = False And stndEQ = 1 And Me.Top < frmPlayList.Top + 120 And Me.Top > frmPlayList.Top - 120 Then
                        frmPlayList.Move (Me.Left + Me.Width), Me.Top
                        PlayListHangsL = True
                        PlayListHangsT = False
                    Else
                        PlayListHangsL = False
                    End If
                Else
                    PlayListHangsL = False
                End If
            Else
                frmPlayList.Move (Me.Left + Me.Width + IIf(EqualizerHangsL = True, frmEqualizer.Width, 0)), Me.Top
            End If
            If PlayListHangsT = False Then
                If posTBC < frmPlayList.Top + 120 And posTBC > frmPlayList.Top - 120 Then
                    If posLBC < frmPlayList.Left + 120 And posLBC > frmPlayList.Left - 120 Then
                        frmPlayList.Move Me.Left, (Me.Top + Me.Height + IIf(EqualizerHangsT = True, frmEqualizer.Height, 0))
                        PlayListHangsT = True
                        PlayListHangsL = False
                    ElseIf Me.Left < frmPlayList.Left + 120 And Me.Left > frmPlayList.Left - 120 Then
                        frmPlayList.Move Me.Left, (Me.Top + Me.Height + IIf(EqualizerHangsT = True, frmEqualizer.Height, 0))
                        PlayListHangsT = True
                    Else
                        PlayListHangsT = False
                    End If
                Else
                    PlayListHangsT = False
                End If
            Else
                frmPlayList.Move Me.Left, (Me.Top + Me.Height + IIf(EqualizerHangsT = True, frmEqualizer.Height, 0))
            End If
        Else
            PlayListHangsL = False
            PlayListHangsT = False
        End If
    End If
End Sub

Private Sub ImgTitleBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm = False
End Sub

Private Sub ImgTitleBarKnop1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgTitleBarKnop1.Picture = PicTitleBut(4).Image
End Sub

Private Sub ImgTitleBarKnop1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgTitleBarKnop1.Picture = PicTitleBut(1).Image
    Me.WindowState = 1
    If stndPlayList = 1 Then imgPlayList_MouseUp 1, 0, 345, 300:
End Sub

Private Sub ImgTitleBarKnop2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgTitleBarKnop2.Picture = PicTitleBut(7).Image
End Sub

Private Sub ImgTitleBarKnop2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgTitleBarKnop2.Picture = PicTitleBut(6).Image
    If Minimized = True Then
        Minimized = False
    Else
        Minimized = True
    End If
    If Minimized = False Then
        imgSPrev.Enabled = False
        imgSPlay.Enabled = False
        imgSPause.Enabled = False
        imgSStop.Enabled = False
        imgSNext.Enabled = False
        imgSEject.Enabled = False
        If BigForm = True Then
            Me.Height = 3480
            ImgMenu.Top = 90
            ImgTitleBarKnop1.Top = 120
            ImgTitleBarKnop2.Top = 120
            ImgTitleBarKnop3.Top = 120
            vtTimer.Visible = False
        Else
            Me.Height = 1740
            ImgMenu.Top = 45
            ImgTitleBarKnop1.Top = 45
            ImgTitleBarKnop2.Top = 45
            ImgTitleBarKnop3.Top = 45
            vtTimer.Visible = False
        End If
        inTitleBar = 0
        ImgTitleBar.Picture = PicTitleBar(inTitleBar).Image
    Else
        imgSPrev.Enabled = True
        imgSPlay.Enabled = True
        imgSPause.Enabled = True
        imgSStop.Enabled = True
        imgSNext.Enabled = True
        imgSEject.Enabled = True
        If BigForm = True Then
            Me.Height = 390
            ImgMenu.Top = 60
            ImgTitleBarKnop1.Top = 90
            ImgTitleBarKnop2.Top = 90
            ImgTitleBarKnop3.Top = 90
            vtTimer.Move 3870, 90
            vtTimer.DoubleSize = True
            vtTimer.Visible = True
        Else
            Me.Height = 195
            ImgMenu.Top = 30
            ImgTitleBarKnop1.Top = 30
            ImgTitleBarKnop2.Top = 30
            ImgTitleBarKnop3.Top = 30
            vtTimer.Move 1935, 45
            vtTimer.DoubleSize = False
            vtTimer.Visible = True
        End If
        inTitleBar = 2
        ImgTitleBar.Picture = PicTitleBar(inTitleBar).Image
    End If
    If EqualizerHangsL = True Then
        frmEqualizer.Top = Me.Top + Me.Height
        frmEqualizer.Left = Me.Left
    End If
    If EqualizerHangsT = True Then
        frmEqualizer.Top = Me.Top + Me.Height
        frmEqualizer.Left = Me.Left
    End If
    If PlayListHangsL = True Then
        frmPlayList.Top = Me.Top + Me.Height + IIf(EqualizerHangsT = True And stndEQ = 1, frmEqualizer.Height, 0)
        frmPlayList.Left = Me.Left
    End If
    If PlayListHangsT = True Then
        frmPlayList.Top = Me.Top + Me.Height + IIf(EqualizerHangsT = True And stndEQ = 1, frmEqualizer.Height, 0)
        frmPlayList.Left = Me.Left
    End If
    If WinVerOk = True Then
        If frmPlayList.comFunc1.FileExist(transSkinPath) And frmMain.Minimized = False Then
            frmPlayList.comFunc1.TransRestore frmMain.hWnd, IIf(BigForm = False, 4125, 8250), IIf(BigForm = False, 1740, 3480)
            frmPlayList.comFunc1.TransWinamp frmMain.hWnd, frmEqualizer.hWnd, transSkinPath, BigForm
        Else
            frmPlayList.comFunc1.TransRestore frmMain.hWnd, IIf(BigForm = False, 4125, 8250), IIf(BigForm = False, 1740, 3480)
        End If
        If stndEQ = 0 Then frmEqualizer.Hide:
    End If
End Sub

Private Sub ImgTitleBarKnop3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgTitleBarKnop3.Picture = PicTitleBut(5).Image
    DisplayTimer.Enabled = False
    frmPlayList.Playlist1.EventStop
    frmPlayList.Playlist1.SaveList
    QuitmciPlayer = True
    Dim Form As Form
    For Each Form In Forms
        If Form.Name <> Me.Name Then
            Unload Form
            Set Form = Nothing
        End If
    Next Form
    Unload Me
    End
End Sub

Private Sub ImgTitleBarKnop3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgTitleBarKnop3.Picture = PicTitleBut(2).Image
End Sub

Private Sub mnuSkins_Click()
    frmOptions.Show
End Sub

Private Sub imgVisual_Click()
    If stndVisual = False Then
        imgVisual.Picture = PicTitleOpt(4).Image
        stndVisual = True
        MsgBox "Under Construction !"
    Else
        Set imgVisual.Picture = Nothing
        stndVisual = False
    End If
End Sub

Private Sub PicTime_Click(Index As Integer)
    intTimeMode = intTimeMode + 1
    If intTimeMode = 6 Then intTimeMode = 0
End Sub

Sub WijzigFormaat()
    Me.Visible = False
    On Error GoTo err_Handle
    Me.ScaleMode = 1
    Dim i As Integer, imgNR As Integer
    Dim variabele As Control
    If BigForm = True Then
        Me.Width = 8250
        Me.Height = 3480
        For Each variabele In Controls
            If TypeOf variabele Is Image And InStr(1, "Src", variabele.Name, 0) = 0 Then
                variabele.Stretch = True
                variabele.Move (variabele.Left * 2), (variabele.Top * 2), (variabele.Width * 2), (variabele.Height * 2)
            End If
        Next
        PlaceTimer
        TimeSlider1.DoubleFormat = True
        TimeSlider1.Move 450, 2160
        Volume1.DoubleSize = True
        Volume1.Move 3210, 1710
        Balance1.DoubleSize = True
        Balance1.Move 5310, 1710
        ScrollTitle1.DoubleSize = True
        ScrollTitle1.Move 3300, 810
        vtHZ.DoubleSize = True
        vtHZ.Move 4680, 1290
        vtKBPS.DoubleSize = True
        vtKBPS.Move 3330, 1290, 150, 180
        PicStatus.Move 780, 840, 270, 270
        PicRecording.Move 720, 840, 90, 270
        imgNR = 0
    Else
        Me.Width = 4125
        Me.Height = 1740
        For Each variabele In Controls
            If TypeOf variabele Is Image And InStr(1, "Src", variabele.Name, 0) = 0 Then
                variabele.Stretch = True
                variabele.Move (variabele.Left / 2), (variabele.Top / 2), (variabele.Width / 2), (variabele.Height / 2)
            End If
        Next
        PlaceTimer
        TimeSlider1.DoubleFormat = False
        TimeSlider1.Move 225, 1080
        Volume1.DoubleSize = False
        Volume1.Move 1605, 855
        Balance1.DoubleSize = False
        Balance1.Move 2655, 855
        ScrollTitle1.DoubleSize = False
        ScrollTitle1.Move 1650, 405
        vtHZ.DoubleSize = False
        vtHZ.Move 2340, 645
        vtKBPS.DoubleSize = False
        vtKBPS.Move 1665, 645
        PicStatus.Move 390, 420, 135, 135
        PicRecording.Move 360, 420, 45, 135
        imgNR = 0
    End If
    ViewStatus
    If WinVerOk = True Then
        If frmPlayList.comFunc1.FileExist(transSkinPath) And frmMain.Minimized = False Then
            frmPlayList.comFunc1.TransRestore frmMain.hWnd, IIf(BigForm = False, 4125, 8250), IIf(BigForm = False, 1740, 3480)
            frmPlayList.comFunc1.TransRestore frmEqualizer.hWnd, IIf(BigForm = False, 4125, 8250), IIf(BigForm = False, 1740, 3480)
            frmPlayList.comFunc1.TransWinamp frmMain.hWnd, frmEqualizer.hWnd, transSkinPath, BigForm
        Else
            frmPlayList.comFunc1.TransRestore frmMain.hWnd, IIf(BigForm = False, 4125, 8250), IIf(BigForm = False, 1740, 3480)
            frmPlayList.comFunc1.TransRestore frmEqualizer.hWnd, IIf(BigForm = False, 4125, 8250), IIf(BigForm = False, 1740, 3480)
        End If
        If stndEQ = 0 Then frmEqualizer.Hide:
    End If
err_Handle:
    Me.Visible = True
End Sub

Sub WijzigLayOut()
    SetImage PicMain, 0, 0, 275, 116, PicSrcMain, 0, 0
    SetImage PicPrevious, 0, 0, 23, 19, PicSrcButton, 0, 0
    SetImage PicPlay, 0, 0, 23, 19, PicSrcButton, 23, 0
    SetImage PicPause, 0, 0, 23, 19, PicSrcButton, 46, 0
    SetImage PicStop, 0, 0, 23, 19, PicSrcButton, 69, 0
    SetImage PicNext, 0, 0, 22, 19, PicSrcButton, 92, 0
    SetImage PicEject, 0, 0, 22, 16, PicSrcButton, 114, 0
    SetImage PicTitleBar(0), 0, 0, 275, 14, PicSrcTitlebar, 27, 0
    SetImage PicTitleBar(1), 0, 0, 275, 14, PicSrcTitlebar, 27, 15
    SetImage PicTitleBar(2), 0, 0, 275, 13, PicSrcTitlebar, 27, 30
    SetImage PicTitleBar(3), 0, 0, 275, 13, PicSrcTitlebar, 27, 43
    SetImage PicTitleBar(4), 0, 0, 275, 14, PicSrcTitlebar, 27, 56
    SetImage PicTitleBar(5), 0, 0, 275, 14, PicSrcTitlebar, 27, 71
    ImgTitleBar.Picture = PicTitleBar(inTitleBar).Image
    SetPicture PicTitleBut(0), 0, 0, 9, 9, PicSrcTitlebar, 0, 0
    SetPicture PicTitleBut(1), 0, 0, 9, 9, PicSrcTitlebar, 9, 0
    SetPicture PicTitleBut(2), 0, 0, 9, 9, PicSrcTitlebar, 18, 0
    SetPicture PicTitleBut(3), 0, 0, 9, 9, PicSrcTitlebar, 0, 9
    SetPicture PicTitleBut(4), 0, 0, 9, 9, PicSrcTitlebar, 9, 9
    SetPicture PicTitleBut(5), 0, 0, 9, 9, PicSrcTitlebar, 18, 9
    SetPicture PicTitleBut(6), 0, 0, 9, 9, PicSrcTitlebar, 0, 18
    SetPicture PicTitleBut(7), 0, 0, 9, 9, PicSrcTitlebar, 9, 18
    SetPicture PicTitleBut(8), 0, 0, 9, 9, PicSrcTitlebar, 0, 27
    SetPicture PicTitleBut(9), 0, 0, 9, 9, PicSrcTitlebar, 9, 27
    ImgMenu.Picture = PicTitleBut(0).Image
    ImgTitleBarKnop1.Picture = PicTitleBut(1).Image
    ImgTitleBarKnop2.Picture = PicTitleBut(6).Image
    ImgTitleBarKnop3.Picture = PicTitleBut(2).Image
    SetImage PicTitleOptEmpty, 0, 0, 8, 40, PicSrcTitlebar, 304, 2
    SetPicture PicTitleOpt(0), 0, 0, 8, 8, PicSrcTitlebar, 304, 46
    SetPicture PicTitleOpt(1), 0, 0, 8, 8, PicSrcTitlebar, 312, 54
    SetPicture PicTitleOpt(2), 0, 0, 8, 8, PicSrcTitlebar, 320, 62
    SetPicture PicTitleOpt(3), 0, 0, 8, 8, PicSrcTitlebar, 328, 70
    SetPicture PicTitleOpt(4), 0, 0, 8, 8, PicSrcTitlebar, 336, 78
    If stndOptions = True Then
        imgOptions.Picture = PicTitleOpt(0).Image
    Else
        Set imgOptions.Picture = Nothing
    End If
    If stndOnTop = True Then
        imgAllOnTop.Picture = PicTitleOpt(1).Image
    Else
        Set imgAllOnTop.Picture = Nothing
    End If
    If stndInfo = True Then
        imgFileInfo.Picture = PicTitleOpt(2).Image
    Else
        Set imgFileInfo.Picture = Nothing
    End If
    If BigForm = True Then
        imgDouble.Picture = PicTitleOpt(3).Image
    Else
        Set imgDouble.Picture = Nothing
    End If
    If stndVisual = True Then
        imgVisual.Picture = PicTitleOpt(4).Image
    Else
        Set imgVisual.Picture = Nothing
    End If
    SetImage PicRepeat(0), 0, 0, 28, 15, PicSrcShufRep, 0, 0
    SetImage PicRepeat(1), 0, 0, 28, 15, PicSrcShufRep, 0, 15
    SetImage PicRepeat(2), 0, 0, 28, 15, PicSrcShufRep, 0, 30
    SetImage PicRepeat(3), 0, 0, 28, 15, PicSrcShufRep, 0, 45
    If frmPlayList.Playlist1.Repeat = 0 Then
        imgRepeat.Picture = PicRepeat(0).Image
    Else
        imgRepeat.Picture = PicRepeat(2).Image
    End If
    SetImage PicShuffle(0), 0, 0, 47, 15, PicSrcShufRep, 28, 0
    SetImage PicShuffle(1), 0, 0, 47, 15, PicSrcShufRep, 28, 15
    SetImage PicShuffle(2), 0, 0, 47, 15, PicSrcShufRep, 28, 30
    SetImage PicShuffle(3), 0, 0, 47, 15, PicSrcShufRep, 28, 45
    If frmPlayList.Playlist1.ShuffleMode = False Then
        imgShuffle.Picture = PicShuffle(0).Image
    Else
        imgShuffle.Picture = PicShuffle(2).Image
    End If
    SetImage PicEQ(0), 0, 0, 23, 12, PicSrcShufRep, 0, 61
    SetImage PicEQ(1), 0, 0, 23, 12, PicSrcShufRep, 46, 61
    SetImage PicEQ(2), 0, 0, 23, 12, PicSrcShufRep, 0, 73
    SetImage PicEQ(3), 0, 0, 23, 12, PicSrcShufRep, 46, 73
    If stndEQ = 0 Then
        imgEQ.Picture = PicEQ(0).Image
    Else
        imgEQ.Picture = PicEQ(2).Image
    End If
    SetImage PicPlayList(0), 0, 0, 23, 12, PicSrcShufRep, 23, 61
    SetImage PicPlayList(1), 0, 0, 23, 12, PicSrcShufRep, 69, 61
    SetImage PicPlayList(2), 0, 0, 23, 12, PicSrcShufRep, 23, 73
    SetImage PicPlayList(3), 0, 0, 23, 12, PicSrcShufRep, 69, 73
    If stndPlayList = 0 Then
        imgPlayList.Picture = PicPlayList(0).Image
    Else
        imgPlayList.Picture = PicPlayList(2).Image
    End If
    If frmPlayList.Playlist1.CurrentStereo = False Then
        SetImage PicMono, 0, 0, 27, 12, PicSrcMonoSter, 29, 0
        SetImage PicStereo, 0, 0, 28, 12, PicSrcMonoSter, 0, 12
    Else
        SetImage PicMono, 0, 0, 27, 12, PicSrcMonoSter, 29, 12
        SetImage PicStereo, 0, 0, 28, 12, PicSrcMonoSter, 0, 0
    End If
    ViewStatus
End Sub

Sub SetImage(Destination As Control, X As Long, Y As Long, W As Long, H As Long, Source As Control, startX As Long, StartY As Long)
    On Error GoTo err_Handle
    Dim DestName As String
    BitBlt Destination.hdc, X, Y, W, H, Source.hdc, startX, StartY, SRCCOPY
    Destination.Refresh
    DestName = "img" & Mid$(Destination.Name, 4, Len(Destination.Name))
    Me(DestName).Picture = Destination.Image
err_Handle:
End Sub

Sub SetPicture(Destination As Control, X As Long, Y As Long, W As Long, H As Long, Source As Control, startX As Long, StartY As Long)
    Dim DestName As String
    BitBlt Destination.hdc, X, Y, W, H, Source.hdc, startX, StartY, SRCCOPY
    Destination.Refresh
End Sub

Private Sub ViewStatus()
    PicRecording.Visible = False
    If BigForm = False Then
        Select Case frmPlayList.Playlist1.Status
            Case Is = "Playing"
                BitBlt PicStatus.hdc, 0, 0, 9, 9, PicSrcStatus.hdc, 0, 0, SRCCOPY
                PicRecording.Visible = True
                If Recording = True Then
                    BitBlt PicRecording.hdc, 0, 0, 3, 9, PicSrcStatus.hdc, 39, 0, SRCCOPY
                Else
                    BitBlt PicRecording.hdc, 0, 0, 3, 9, PicSrcStatus.hdc, 36, 0, SRCCOPY
                End If
                PicRecording.Refresh
            Case Is = "Paused"
                BitBlt PicStatus.hdc, 0, 0, 9, 9, PicSrcStatus.hdc, 9, 0, SRCCOPY
            Case Is = "Stopped"
                BitBlt PicStatus.hdc, 0, 0, 9, 9, PicSrcStatus.hdc, 18, 0, SRCCOPY
            Case Else
                BitBlt PicStatus.hdc, 0, 0, 9, 9, PicSrcStatus.hdc, 27, 0, SRCCOPY
        End Select
    Else
        Select Case frmPlayList.Playlist1.Status
            Case Is = "Playing"
                StretchBlt PicStatus.hdc, 0, 0, 18, 18, PicSrcStatus.hdc, 0, 0, 9, 9, SRCCOPY
                PicRecording.Visible = True
                If Recording = True Then
                    StretchBlt PicRecording.hdc, 0, 0, 6, 18, PicSrcStatus.hdc, 39, 0, 3, 9, SRCCOPY
                Else
                    StretchBlt PicRecording.hdc, 0, 0, 6, 18, PicSrcStatus.hdc, 36, 0, 3, 9, SRCCOPY
                End If
                PicRecording.Refresh
            Case Is = "Paused"
                StretchBlt PicStatus.hdc, 0, 0, 18, 18, PicSrcStatus.hdc, 9, 0, 9, 9, SRCCOPY
            Case Is = "Stopped"
                StretchBlt PicStatus.hdc, 0, 0, 18, 18, PicSrcStatus.hdc, 18, 0, 9, 9, SRCCOPY
            Case Else
                StretchBlt PicStatus.hdc, 0, 0, 18, 18, PicSrcStatus.hdc, 27, 0, 9, 9, SRCCOPY
        End Select
    End If
    PicStatus.Refresh
End Sub

Private Sub imgPositions()
    ImgMain.Move 0, 0, 4125, 1740
    imgPrevious.Move 240, 1320, 330, 255
    imgPlay.Move 585, 1320, 330, 255
    imgPause.Move 930, 1320, 330, 255
    imgStop.Move 1275, 1320, 330, 255
    imgNext.Move 1620, 1320, 330, 255
    imgEject.Move 2040, 1335, 330, 240
    ImgMenu.Move 90, 45, 135, 135
    ImgTitleBar.Move 0, 0, 4125, 210
    ImgTitleBarKnop1.Move 3645, 45, 135, 120
    ImgTitleBarKnop2.Move 3795, 45, 135, 120
    ImgTitleBarKnop3.Move 3945, 45, 135, 120
    ImgTitleOptEmpty.Move 150, 360, 120, 600
    imgOptions.Move 150, 360, 120, 120
    imgAllOnTop.Move 150, 480, 120, 120
    imgDouble.Move 150, 720, 120, 120
    imgFileInfo.Move 150, 600, 120, 120
    imgVisual.Move 150, 840, 120, 120
    imgRepeat.Move 3150, 1335, 420, 225
    imgShuffle.Move 2460, 1335, 705, 225
    imgEQ.Move 3285, 870, 345, 180
    imgPlayList.Move 3630, 870, 345, 180
    imgMono.Move 3180, 615, 405, 180
    imgStereo.Move 3585, 615, 420, 180
    imgSPrev.Move 2505, 45, 135, 105
    imgSPlay.Move 2640, 45, 135, 105
    imgSPause.Move 2775, 45, 135, 105
    imgSStop.Move 2910, 45, 135, 105
    imgSNext.Move 3045, 45, 135, 105
    imgSEject.Move 3180, 45, 135, 105
End Sub

Private Sub imgCorrections()
    If BigForm = False Then
        Dim variabele As Control
        For Each variabele In Controls
            If TypeOf variabele Is Image And InStr(1, "Src", variabele.Name, 0) = 0 Then
                variabele.Stretch = False
                variabele.Move (variabele.Left * 2), (variabele.Top * 2), (variabele.Width * 2), (variabele.Height * 2)
            End If
        Next
    End If
End Sub

Private Sub ScrollTimer_Timer()
    On Error GoTo err_Handle
    Static sOldSkinDir As String
    Static sOldSkinName As String
    If sOldSkinDir <> frmPlayList.Playlist1.SkinDir Then
        sOldSkinDir = frmPlayList.Playlist1.SkinDir
    End If
    If sOldSkinName <> frmPlayList.Playlist1.SkinName Then
        sOldSkinName = frmPlayList.Playlist1.SkinName
        If sOldSkinDir = "" Or sOldSkinName = "" Then GoTo SCROLL:
        frmPlayList.Playlist1.WinAMP = True
        frmPlayList.SetHoleView
    End If
SCROLL:
    DoEvents
    Static sOldMed As String
    Static sCaption As String
    If sOldTitle <> frmPlayList.Playlist1.ScrollText Then
        sOldTitle = frmPlayList.Playlist1.ScrollText
        stateTitle = sOldTitle
        Dim iPoint As Integer
        iPoint = InStr(1, sOldTitle, ".")
        sCaption = Mid(sOldTitle, iPoint + 1, Len(sOldTitle) - iPoint + 1)
        ScrollTitle1.ScrollText = sOldTitle
        introTimer = Timer + 10
    End If
    If sOldMed <> frmPlayList.Playlist1.CurrentFile Then
        sOldMed = frmPlayList.Playlist1.CurrentFile
        Select Case frmPlayList.Playlist1.CurrentExt
            Case Is = "CDA": Volume1.VolumeDevice dvcCD:
            Case Is = "MP1", "MP2", "MP3", "WAV", "AVI", "MOV", "MPG", "MPEG", "MPE": Volume1.VolumeDevice dvcWav:
            Case Is = "MID": Volume1.VolumeDevice dvcMidi
            Case Else
                Volume1.VolumeDevice dvcMaster
        End Select
        vtKBPS.ScrollText = IIf(frmPlayList.Playlist1.CurrentkBps < 100, IIf(frmPlayList.Playlist1.CurrentkBps < 10, "  " & frmPlayList.Playlist1.CurrentkBps, " " & frmPlayList.Playlist1.CurrentkBps), frmPlayList.Playlist1.CurrentkBps)
        vtHZ.ScrollText = IIf(frmPlayList.Playlist1.CurrentHz < 10, " " & frmPlayList.Playlist1.CurrentHz, frmPlayList.Playlist1.CurrentHz)
        vtKBPS.ShowMessage IIf(frmPlayList.Playlist1.CurrentkBps < 100, IIf(frmPlayList.Playlist1.CurrentkBps < 10, "  " & frmPlayList.Playlist1.CurrentkBps, " " & frmPlayList.Playlist1.CurrentkBps), frmPlayList.Playlist1.CurrentkBps)
        vtHZ.ShowMessage IIf(frmPlayList.Playlist1.CurrentHz < 10, " " & frmPlayList.Playlist1.CurrentHz, frmPlayList.Playlist1.CurrentHz)
        If frmPlayList.Playlist1.CurrentStereo = False Then
            SetImage PicMono, 0, 0, 27, 12, PicSrcMonoSter, 29, 0
            SetImage PicStereo, 0, 0, 28, 12, PicSrcMonoSter, 0, 12
        Else
            SetImage PicMono, 0, 0, 27, 12, PicSrcMonoSter, 29, 12
            SetImage PicStereo, 0, 0, 28, 12, PicSrcMonoSter, 0, 0
        End If
        CheckWindowState
    End If
    If Minimized = False And WindowState = 0 Then
        If frmPlayList.Playlist1.Status <> "Playing" Then introTimer = Timer + 10:
        If Volume1.ChangeVolume = False And Balance1.ChangeBalance = False Then
            ScrollTitle1.ScrollNow
            Me.Caption = sCaption
        Else
            If Volume1.ChangeVolume = True Then ScrollTitle1.ShowMessage Volume1.Message:
            If Balance1.ChangeBalance = True Then ScrollTitle1.ShowMessage Balance1.Message:
        End If
    Else
        CheckWindowState
    End If
    ViewOrNot
    ViewStatus
err_Handle:
End Sub

Public Sub CheckWindowState()
    DoEvents
    Dim cHour, cMin, cSec, capTime
    cHour = Int(TimeHasPlayed / 3600)
    cMin = Int((TimeHasPlayed - (cHour * 3600)) / 60)
    cSec = Int((TimeHasPlayed - (cHour * 3600)) Mod 60)
    stateTitle = Mid(stateTitle, 2) & Left(stateTitle, 1)
    If WindowState = 1 Then
        If cHour <> 0 Then
            capTime = IIf(cHour < 10, "0" & cHour, cHour) & ":" & IIf(cMin < 10, "0" & cMin, cMin) & ":" & IIf(cSec < 10, "0" & cSec, cSec)
        Else
            capTime = IIf(cMin < 10, "0" & cMin, cMin) & ":" & IIf(cSec < 10, "0" & cSec, cSec)
        End If
        Me.Caption = capTime & "  " & stateTitle
    End If
    If Minimized = True And WindowState = 0 Then
        Me.Caption = stateTitle
        vtTimer.ShowMessage IIf(cMin < 10, "0" & cMin, cMin) & ":" & IIf(cSec < 10, "0" & cSec, cSec)
    End If
End Sub

Public Sub ViewOrNot()
    Dim blnMovieShow As Boolean
    blnMovieShow = frmPlayList.Playlist1.IsMovie(frmPlayList.Playlist1.CurrentFile)
    If blnMovieShow = True Then
        frmPlayList.picREM.Enabled = False
    Else
        frmPlayList.picREM.Enabled = True
    End If
    If stndPlayList = 0 Then
        Static stndShowing
        If blnMovieShow Then
            If stndShowing <> 1 Then
                frmPlayList.Show
            End If
            stndShowing = 1
        Else
            If stndShowing <> 0 And stndPlayList = 0 Then
                frmPlayList.Hide
            End If
            stndShowing = 0
        End If
    End If
End Sub

Private Sub TimeDisplay1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Select Case TimeDisplay1.Leds
            Case Is = 4: TimeDisplay1.Leds = 5:
            Case Is = 5: TimeDisplay1.Leds = 6:
            Case Is = 6: TimeDisplay1.Leds = 4:
        End Select
        PlaceTimer
    Else
        intTimeMode = intTimeMode + 1
        If intTimeMode = 6 Then intTimeMode = 0:
        If TimeDisplay1.Leds = 4 And intTimeMode > 2 Then intTimeMode = 0:
    End If
End Sub

Private Sub PlaceTimer()
    If BigForm = True Then
        If Minimized = True Then
            vtTimer.Move 3870, 90
            vtTimer.DoubleSize = True
            vtTimer.Visible = True
        Else
            vtTimer.Visible = False
            Select Case TimeDisplay1.Leds
                Case Is = 4
                    TimeDisplay1.Move 1440, 780
                Case Is = 5
                    TimeDisplay1.Move 900, 780
                Case Is = 6
                    TimeDisplay1.Move 540, 780
            End Select
            TimeDisplay1.DoubleSize = True
        End If
    Else
        If Minimized = True Then
            vtTimer.Move 1935, 45
            vtTimer.DoubleSize = False
            vtTimer.Visible = True
        Else
            vtTimer.Visible = False
            Select Case TimeDisplay1.Leds
                Case Is = 4
                    TimeDisplay1.Move 720, 390
                Case Is = 5
                    TimeDisplay1.Move 450, 390
                Case Is = 6
                    TimeDisplay1.Move 270, 390
            End Select
            TimeDisplay1.DoubleSize = False
        End If
    End If
End Sub

Private Sub vtTimer_Click()
    intTimeMode = intTimeMode + 1
    If intTimeMode > 2 Then intTimeMode = 0:
End Sub

Private Sub SetVT()
    Dim cHour, cMin, cSec, tTime, tTime1 As String
    tTime = frmPlayList.Playlist1.TimeTotal
    cHour = Int(tTime / 3600)
    cMin = Int((tTime - (cHour * 3600)) / 60)
    cSec = Int((tTime - (cHour * 3600)) Mod 60)
    tTime1 = "/" & IIf(cHour < 10, "0" & cHour, cHour) & ":" & IIf(cMin < 10, "0" & cMin, cMin) & ":" & IIf(cSec < 10, "0" & cSec, cSec)
    tTime = frmPlayList.Playlist1.TimeSong
    cMin = Int(tTime / 60)
    cSec = Int(tTime Mod 60)
    tTime1 = IIf(cMin < 10, "0" & cMin, cMin) & ":" & IIf(cSec < 10, "0" & cSec, cSec) & tTime1
    frmPlayList.vtTotaltime.ShowMessage tTime1
    tTime = frmPlayList.Playlist1.TimeElapsed
    cMin = Int(tTime / 60)
    cSec = Int(tTime Mod 60)
    frmPlayList.vtTime.ShowMessage " " & IIf(cMin < 10, "0" & cMin, cMin) & ":" & IIf(cSec < 10, "0" & cSec, cSec)
End Sub
