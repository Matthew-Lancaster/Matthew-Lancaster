VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   Caption         =   "Matt Run AS"
   ClientHeight    =   10068
   ClientLeft      =   60
   ClientTop       =   -1536
   ClientWidth     =   14220
   FillStyle       =   0  'Solid
   Icon            =   "Matt_RunAS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10068
   ScaleWidth      =   14220
   Begin VB.Timer Timer_ART_TRIP_WIRE 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8880
      Top             =   8340
   End
   Begin VB.ListBox List1 
      Height          =   240
      ItemData        =   "Matt_RunAS.frx":014A
      Left            =   2052
      List            =   "Matt_RunAS.frx":014C
      Sorted          =   -1  'True
      TabIndex        =   103
      Top             =   156
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "0 - URL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   3855
      TabIndex        =   115
      Top             =   3135
      Width           =   840
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB Sort_AnyThing - MULTI USE MENU COPY PROPER.exe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   0
      TabIndex        =   114
      Top             =   1305
      Width           =   6705
   End
   Begin VB.Label Label118 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   675
      TabIndex        =   113
      Top             =   6390
      Width           =   135
   End
   Begin VB.Label Label116 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "WORNET 2.1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   1380
      TabIndex        =   112
      Top             =   6405
      Width           =   1470
   End
   Begin VB.Label Label117 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Norton Ghost"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   9990
      TabIndex        =   111
      Top             =   1365
      Width           =   1545
   End
   Begin VB.Label Label115 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Norton Internet Security"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   10155
      TabIndex        =   110
      Top             =   1695
      Width           =   2730
   End
   Begin VB.Label Label114 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB DATE 1991"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   9915
      TabIndex        =   109
      Top             =   2895
      Width           =   1665
   End
   Begin VB.Label Label113 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "# WINDOWS MEDIA PLAYER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   8796
      TabIndex        =   108
      Top             =   3708
      Width           =   3228
   End
   Begin VB.Label Label112 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB DIR ON ANOTHER DRIVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   3630
      TabIndex        =   107
      Top             =   6360
      Width           =   3315
   End
   Begin VB.Label Label111 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "TXT VOICE LOGG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   9720
      TabIndex        =   106
      Top             =   3270
      Width           =   2025
   End
   Begin VB.Label Label110 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "WINAMP LOADER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   9585
      TabIndex        =   105
      Top             =   2175
      Width           =   2070
   End
   Begin VB.Label Label109 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB RUN FAVS WITH COMMAND"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   4170
      TabIndex        =   104
      Top             =   7455
      Width           =   3660
   End
   Begin VB.Label Label108 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Bluetooth File Transfer Wizard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   8055
      TabIndex        =   102
      Top             =   795
      Width           =   3465
   End
   Begin VB.Label Label106 
      AutoSize        =   -1  'True
      BackColor       =   &H00400000&
      Caption         =   "VB METRONOME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   4545
      TabIndex        =   100
      Top             =   5265
      Width           =   1995
   End
   Begin VB.Label Label105 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   6240
      TabIndex        =   99
      Top             =   1536
      Width           =   144
   End
   Begin VB.Label Label104 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB Kaleidoscope"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   9525
      TabIndex        =   98
      Top             =   5130
      Width           =   1980
   End
   Begin VB.Label Label103 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   6000
      TabIndex        =   97
      Top             =   1872
      Width           =   144
   End
   Begin VB.Label Label102 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "TASK MANAGER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   8490
      TabIndex        =   96
      Top             =   1920
      Width           =   1950
   End
   Begin VB.Label Label101 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   5772
      TabIndex        =   95
      Top             =   1488
      Width           =   144
   End
   Begin VB.Label Label100 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Google Downloader WIG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   8385
      TabIndex        =   94
      Top             =   7515
      Width           =   2805
   End
   Begin VB.Label Label99 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   5928
      TabIndex        =   93
      Top             =   1500
      Width           =   144
   End
   Begin VB.Label Label98 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   5772
      TabIndex        =   92
      Top             =   1476
      Width           =   144
   End
   Begin VB.Label Label97 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "TXT - Programming Snippets.txt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   9270
      TabIndex        =   91
      Top             =   7845
      Width           =   3645
   End
   Begin VB.Label Label96 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Bluetooth Devices"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   8556
      TabIndex        =   90
      Top             =   468
      Width           =   2064
   End
   Begin VB.Label Label95 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB SCREEN SAVER ON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   4005
      TabIndex        =   89
      Top             =   675
      Width           =   2760
   End
   Begin VB.Label Label94 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   5652
      TabIndex        =   88
      Top             =   1812
      Width           =   144
   End
   Begin VB.Label Label93 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "# Web Image Grab PC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   9600
      TabIndex        =   87
      Top             =   6045
      Width           =   2490
   End
   Begin VB.Label Label92 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "DARK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   4320
      TabIndex        =   86
      Top             =   1632
      Width           =   660
   End
   Begin VB.Label Label91 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "TXT - 2Double Trouble"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   9945
      TabIndex        =   85
      Top             =   6750
      Width           =   2565
   End
   Begin VB.Label Label90 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   5832
      TabIndex        =   84
      Top             =   1536
      Width           =   144
   End
   Begin VB.Label Label89 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Process Tamer Tray"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   2700
      TabIndex        =   83
      Top             =   4215
      Width           =   2295
   End
   Begin VB.Label Label88 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "TXT - 2PINNO'S.TXT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   10155
      TabIndex        =   82
      Top             =   6420
      Width           =   2310
   End
   Begin VB.Label Label87 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   5832
      TabIndex        =   81
      Top             =   1512
      Width           =   144
   End
   Begin VB.Label Label86 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "SHELL MENU VIEW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   150
      TabIndex        =   80
      Top             =   6075
      Width           =   2265
   End
   Begin VB.Label Label85 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   5760
      TabIndex        =   79
      Top             =   1668
      Width           =   144
   End
   Begin VB.Label Label84 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Nero ShowTime"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   2100
      TabIndex        =   78
      Top             =   456
      Width           =   1788
   End
   Begin VB.Label Label82 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "FIREFOX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   2055
      TabIndex        =   76
      Top             =   4575
      Width           =   1050
   End
   Begin VB.Label Label81 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB SCREEN SAVER OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   3960
      TabIndex        =   75
      Top             =   1170
      Width           =   2880
   End
   Begin VB.Label Label80 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "# mplayer2.exe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   2685
      TabIndex        =   74
      Top             =   3600
      Width           =   1710
   End
   Begin VB.Label Label79 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB ClipBoard Logger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   9012
      TabIndex        =   73
      Top             =   8160
      Width           =   2376
   End
   Begin VB.Label Label78 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB VU METER LOGGER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   3780
      TabIndex        =   72
      Top             =   2325
      Width           =   2760
   End
   Begin VB.Label Label77 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "SHELL EX VIEW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   150
      TabIndex        =   71
      Top             =   5775
      Width           =   1875
   End
   Begin VB.Label Label76 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB PLAY BreakOut"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   9930
      TabIndex        =   70
      Top             =   5745
      Width           =   2175
   End
   Begin VB.Label Label75 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB Sort_AnyThing - MULTI USE MENU.exe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   4905
      TabIndex        =   69
      Top             =   3360
      Width           =   4830
   End
   Begin VB.Label Label74 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Drives_Gig"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   2112
      TabIndex        =   68
      Top             =   2988
      Width           =   1248
   End
   Begin VB.Label Label73 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB TIMER CODE.exe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   6708
      TabIndex        =   67
      Top             =   2724
      Width           =   2364
   End
   Begin VB.Label Label72 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB CLOCK #M02 SCR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   8835
      TabIndex        =   66
      Top             =   4860
      Width           =   2520
   End
   Begin VB.Label Label71 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Process Tamer Configurator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   3120
      TabIndex        =   65
      Top             =   3876
      Width           =   3228
   End
   Begin VB.Label Label70 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB Video Scroller Credits"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   315
      TabIndex        =   64
      Top             =   7020
      Width           =   2895
   End
   Begin VB.Label Label69 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Fat 32 Formatter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   9735
      TabIndex        =   63
      Top             =   4095
      Width           =   1860
   End
   Begin VB.Label Label68 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB URL Logger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   9630
      TabIndex        =   62
      Top             =   8490
      Width           =   1800
   End
   Begin VB.Label Label67 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Mp3DirectCut.exe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   5490
      TabIndex        =   61
      Top             =   4890
      Width           =   2025
   End
   Begin VB.Label Label66 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB Tidal.exe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   1956
      TabIndex        =   60
      Top             =   2544
      Width           =   1416
   End
   Begin VB.Label Label65 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB Cid-RunMe.exe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   1512
      TabIndex        =   59
      Top             =   732
      Width           =   2112
   End
   Begin VB.Label Label64 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB WebDates"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   2835
      TabIndex        =   58
      Top             =   5505
      Width           =   1575
   End
   Begin VB.Label Label58 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   5532
      TabIndex        =   50
      Top             =   1608
      Width           =   144
   End
   Begin VB.Label Label57 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   5460
      TabIndex        =   49
      Top             =   1428
      Width           =   144
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "WINAMP LOGGER EXE"
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
      Height          =   285
      Left            =   15
      TabIndex        =   48
      Top             =   3615
      Width           =   2640
   End
   Begin VB.Label Label52 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB RUN FAVS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   4140
      TabIndex        =   45
      Top             =   6795
      Width           =   1650
   End
   Begin VB.Label LABEL51 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   1695
      TabIndex        =   44
      Top             =   5565
      Width           =   120
   End
   Begin VB.Label Label47 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "FILEMON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   5445
      TabIndex        =   42
      Top             =   4245
      Width           =   1080
   End
   Begin VB.Label Label45 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "SUPER COPIER CONFIG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   2892
      TabIndex        =   41
      Top             =   2736
      Width           =   2772
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "# FastStone Image Viewer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   2400
      TabIndex        =   39
      Top             =   4995
      Width           =   2955
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "TRUE CRYPT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   4290
      TabIndex        =   35
      Top             =   5460
      Width           =   1560
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "ADDRESS BOOK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   3645
      TabIndex        =   33
      Top             =   4530
      Width           =   1980
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "WINAMP LOGGER TXT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   15
      TabIndex        =   32
      Top             =   4065
      Width           =   2610
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "WordNet 2.1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   4515
      TabIndex        =   27
      Top             =   5925
      Width           =   1395
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "SOUND VOL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   1800
      TabIndex        =   22
      Top             =   2256
      Width           =   1416
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Winamp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   3228
      TabIndex        =   9
      Top             =   1872
      Width           =   924
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "COOL 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   2820
      TabIndex        =   6
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Encarta Encyclopedia 2001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   396
      TabIndex        =   2
      Top             =   1032
      Width           =   3084
   End
   Begin VB.Label Label107 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "KEYBOARD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   4995
      TabIndex        =   101
      Top             =   150
      Width           =   1380
   End
   Begin VB.Label Label83 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Run Fav Programs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   2700
      TabIndex        =   77
      Top             =   135
      Width           =   2145
   End
   Begin VB.Label Label49 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB - 01 TAGG - GEN - CODE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   120
      TabIndex        =   57
      Top             =   8325
      Width           =   3165
   End
   Begin VB.Label Label55 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB - 02 DUMMMY TAGG - GEN - CODE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   60
      TabIndex        =   56
      Top             =   7890
      Width           =   4320
   End
   Begin VB.Label Label56 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB - 03 PLAYLIST - GEN - CODE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   90
      TabIndex        =   55
      Top             =   7470
      Width           =   3645
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "# Vcap_Src_WebCam"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   9195
      TabIndex        =   54
      Top             =   60
      Width           =   2475
   End
   Begin VB.Label Label61 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "# CAM SPLITTER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   7932
      TabIndex        =   53
      Top             =   1116
      Width           =   1956
   End
   Begin VB.Label Label60 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Thesaurus Ultralingua"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   9315
      TabIndex        =   52
      Top             =   2550
      Width           =   2535
   End
   Begin VB.Label Label59 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "OUTLOOK E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   8100
      TabIndex        =   51
      Top             =   1572
      Width           =   1440
   End
   Begin VB.Label Label54 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB 01 TAGG - GEN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   4830
      TabIndex        =   47
      Top             =   8040
      Width           =   2130
   End
   Begin VB.Label Label53 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB 03 PLAYLIST - GEN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   4830
      TabIndex        =   46
      Top             =   8775
      Width           =   2610
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB 02 DUMMMY TAGG - GEN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   4830
      TabIndex        =   43
      Top             =   8415
      Width           =   3285
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "WINRAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   6750
      TabIndex        =   40
      Top             =   6465
      Width           =   990
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "# Mobile Media Converter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   6750
      TabIndex        =   38
      Top             =   5745
      Width           =   2880
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Easy CD-DA Extractor 5.1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   6750
      TabIndex        =   37
      Top             =   7170
      Width           =   2910
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "WORD PAD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   15
      TabIndex        =   36
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "NOTEPAD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   15
      TabIndex        =   34
      Top             =   4965
      Width           =   1185
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "GoogleImage Downloader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   6750
      TabIndex        =   31
      Top             =   6810
      Width           =   2970
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "POWER CALC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   6750
      TabIndex        =   30
      Top             =   6105
      Width           =   1650
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "SAFE REMOVE HARDWARE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   6750
      TabIndex        =   29
      Top             =   5400
      Width           =   3255
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "WinHTTRACK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   6750
      TabIndex        =   28
      Top             =   5055
      Width           =   1620
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "01 Total List.txt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   15
      TabIndex        =   26
      Top             =   4515
      Width           =   1725
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VICE VERSA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   6750
      TabIndex        =   25
      Top             =   4695
      Width           =   1470
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Process Explorer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   6720
      TabIndex        =   24
      Top             =   4290
      Width           =   1980
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "TWEAK XP Pro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   6756
      TabIndex        =   23
      Top             =   3972
      Width           =   1716
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "AutoRuns"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   15
      TabIndex        =   21
      Top             =   2250
      Width           =   1155
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "KeyBoard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   6732
      TabIndex        =   20
      Top             =   1512
      Width           =   1116
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "PartitionMagic 8.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   15
      TabIndex        =   19
      Top             =   3165
      Width           =   2025
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "SUPER COPIER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   15
      TabIndex        =   18
      Top             =   2715
      Width           =   1845
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "TWEAK UI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   6756
      TabIndex        =   17
      Top             =   3600
      Width           =   1164
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "WinDowse 01 RunAs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   6732
      TabIndex        =   16
      Top             =   2856
      Width           =   2340
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Magnifier"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   6735
      TabIndex        =   15
      Top             =   1020
      Width           =   1050
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "FB Photo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   6732
      TabIndex        =   14
      Top             =   528
      Width           =   1056
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Google Earth"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   6735
      TabIndex        =   13
      Top             =   30
      Width           =   1500
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Elite Spy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   6732
      TabIndex        =   12
      Top             =   2460
      Width           =   996
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "WinDowse 02 Shell"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   6732
      TabIndex        =   11
      Top             =   3228
      Width           =   2148
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Feed Demon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   288
      Left            =   6732
      TabIndex        =   10
      Top             =   2016
      Width           =   1416
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB Encrypt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   1365
      TabIndex        =   8
      Top             =   1875
      Width           =   1320
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "# 0 EXE's"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   45
      TabIndex        =   7
      Top             =   1875
      Width           =   1080
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "COOL 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   1515
      TabIndex        =   5
      Top             =   1500
      Width           =   915
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "COOL 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   195
      TabIndex        =   4
      Top             =   1500
      Width           =   915
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "OPUS2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   45
      Width           =   840
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Palm && Sync"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   435
      TabIndex        =   1
      Top             =   540
      Width           =   1470
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "OPUS1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   855
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu KILLMENU 
      Caption         =   "KILL MENU"
      Begin VB.Menu KILL_JEPG_ART 
         Caption         =   "KILL JEPG ART SLIDESHOW WHEN FROZEN"
      End
   End
   Begin VB.Menu MNU_VBEXE 
      Caption         =   "NOTEPAD VB EXE"
   End
   Begin VB.Menu MNU_VERSION 
      Caption         =   "VERSION"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'RUN AS PROGRAM
'
'
'Screen SAVER




'%SystemRoot%\system32\osk.exe
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






Dim OLDi, iTRIP, XHH

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Const SPI_GETSCREENSAVEACTIVE = 16
Private Const SPIF_SENDWININICHANGE = &H2
Private Const SPIF_UPDATEINIFILE = &H1

Private Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
    ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long





Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Function
Function NotAlwaysOnTop(ByVal hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags
End Function

'"C:\Program Files\Runtime Software\DriveImage XML\dixml.exe"

'WANT BLUE TOOTH ON HERE

Private Sub Form_Load()
'NEED BLUETOOTH CONFIG

Call MNU_VERSION_Click

'Load ScanPath
'Shell "notepad.exe """ + App.Path + "\File of EXE Programs in VB.txt"""

'Exit Sub



'Debug.Print Format(DateValue("13 aug 1998"), "DdD--mm-yy")
'Debug.Print Format(DateValue("13 aug 1999"), "DdD--mm-yy")
'Debug.Print Format(DateValue("13 aug 2000"), "DdD--mm-yy")
'Debug.Print Format(DateValue("13 aug 2001"), "DdD--mm-yy")
'Debug.Print Format(DateValue("13 aug 2002"), "DdD--mm-yy")
'Debug.Print Format(DateValue("13 aug 2003"), "DdD--mm-yy")
'Debug.Print Format(DateValue("13 aug 2004"), "DdD--mm-yy")
'Debug.Print Format(DateValue("13 aug 2005"), "DdD--mm-yy")
'Stop
'End
Set FS = CreateObject("Scripting.FileSystemObject")


'If Command$ = "ART" Or IsIDE = True Then
If Command$ = "ART" Then
    If App.PrevInstance = True Then End
    Me.WindowState = 1
    Timer_ART_TRIP_WIRE.enabled = True

End If


For Each Control In Controls
        
        On Error Resume Next
        Control.Caption = Replace(Control.Caption, "#", "")
        Control.Caption = Trim(Control.Caption)
        On Error GoTo 0


Next





For Each Control In Controls
        
    x = Control.Name = "Timer_ART_TRIP_WIRE"
    If x = False Then
        On Error Resume Next
        If Control.Caption = ".." Then Control.Visible = False
        Control.BackColor = &HFF0000
        On Error GoTo 0

    End If
Next






'ALWAYS




'Shell "C:\Program Files\Media Player Classic\mplayerc.exe", vbNormalFocus
'End





'Code to Auto Size Form based on controls used Including Menu Bar Sketchy Style
'Make sure form set to scale Twips
'put in your form load

Set FS = CreateObject("Scripting.FileSystemObject")
'Sys Object only uses 1 2 3 = 3 temp



' System Folder - Windows\System32
' SystemFolder = oFS.GetSpecialfolder(1)

' Windows Folder Path


sWindowsFolder = GetSpecialfolder(0)

sMyDocsFolder = GetSpecialfolder(5)


On Error Resume Next
For Each Control In Controls
    If InStr(UCase((Control.Name)), "LAB") > 0 Then
        Control.Visible = False
    End If
Next

Dim R As Long
On Error Resume Next
If IsIDE = True Then
'    For R = 0 To 120
'        Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
'    Next
    'End
End If


Me.Show



x = 1
y = 1
On Error Resume Next
For Each Control In Controls

    If Mid(Control.Caption, 1, 1) = "#" Then
        Control.BackColor = &H400000
    End If
    
    If Control.enabled = True And Control.Visible = True Then
        If Control.Width + Control.Left > x Then x = Control.Width + Control.Left
        If Control.Height + Control.Top > y Then y = Control.Height + Control.Top
        If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
    End If
Next
On Error GoTo 0

Me.Width = (x + 80)
If mnu = 1 Then
    pluso = 740 ': pluso = 1100 'Sometimes Different
Else
    pluso = 450
End If

Me.Height = (y + pluso)
Me.Refresh
DoEvents

'Me.Left = (Screen.Width - Me.Width) / 2
'Me.Top = ((Screen.Height - Me.Height) / 2) - 200




Me.Width = (Screen.Width - 500)
Me.Height = (Screen.Height - 3000)


Me.Left = (Screen.Width - Me.Width) / 2
'Me.Top = ((Screen.Height - Me.Height) / 2) - 200
Me.Top = 0


Call SubCode

'
'AlwaysOnTop (Me.hwnd)


End Sub



Sub SubCode()

'Me.Show
DoEvents
x = 0
L = 0


For Each Control In Controls
    If InStr(UCase((Control.Name)), "LAB") > 0 Then
        Control.AutoSize = True
    End If
Next


For Each Control In Controls
    If InStr(UCase((Control.Name)), "LAB") > 0 Then
    If InStr((Control.Caption), "..") = 0 Then
        List1.AddItem Control.Caption
    End If
    End If
Next



For R = 0 To List1.ListCount - 1
    For Each Control In Controls
    
    
    If InStr(UCase((Control.Name)), "LAB") > 0 Then
    If Trim(Control.Caption) <> "" And Control.Caption = List1.List(R) Then
    Debug.Print Control.Caption
    If InStr((Control.Caption), "..") = 0 Then
    If Control.Width > w1 Then w1 = Control.Width
    Control.Visible = True
    If x > Me.Height - 1100 Then x = 0: L = L + w1 + 50: w1 = 0
    
    Control.Top = x
    Control.Left = L
    
    x = x + Control.Height + 15
    
    'If x > Me.Height Then Exit Sub
    
    
    
   End If
   End If
   End If
   

Next
Next

End Sub










Private Sub KILL_JEPG_ART_Click()


Timer_ART_TRIP_WIRE.enabled = True
Exit Sub
Dim HH As Long, OTSS As Long
HH = FindWindow(vbNullString, "JPEG Encoder -- Shockingly Good Photo ART -- [-:¦:-•:*''' The One '''*:•-:¦:-]")
TT = cProcesses.Convert(HH, OTSS, cnFromhWnd Or cnToProcessID)
Process_Kill (OTSS)

End Sub

Private Sub Label1_Click()

'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Olympus\DSSPlayer2002\DictWnd.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing


End
End Sub

Private Sub Label10_Click()
MsgBox "DONT USE NOT ENABLED"
End
Shell "D:\VB6\VB-NT\00_Best_VB_01\Encrypt Email Send\Send Encrypt.exe", vbNormalFocus
End
End Sub

Private Sub Label100_Click()
ChDrive "M:\0 00 Art\Google Downloader"
ChDir "M:\0 00 Art\Google Downloader"
Shell "C:\Program Files\# NO INSTALL REQUIRED\WebImageGrab_pc\WebImageGrab_pro_pc\WIG PRO.exe", vbMaximizedFocus
End
End Sub

Private Sub Label102_Click()
Shell "C:\WINDOWS\system32\taskmgr.exe", vbNormalFocus
End
End Sub

Private Sub Label104_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\00 Graphics\Pattern\Kaleidoscope\Kaleidoscope.exe", vbNormalFocus
End
End Sub

Private Sub Label106_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\METRONOME\METRONOME.exe", vbNormalFocus
End
End Sub

Private Sub Label107_Click()

'KEYBOARD
Shell sWindowsFolder + "\system32\OSK.exe"
End

End Sub

Private Sub Label108_Click()

Shell sWindowsFolder + "\system32\fsquirt.exe -send", vbNormalFocus
End
End Sub

Private Sub Label109_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\Run Fav Progs\Run Fav Programs.exe HH", vbNormalFocus
End
End Sub

Private Sub Label11_Click()
Shell "C:\Program Files\Greatis\WinDowse\WinDowse.exe", vbNormalFocus

End

Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Greatis\WinDowse\WinDowse.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing

End

End Sub



Private Sub Label110_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\Shell Winamp Loader\Winamp Loader.exe", vbNormalFocus
End
End Sub

Private Sub Label111_Click()
    
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
Dim EXECUTE_FILE_NAME
EXECUTE_FILE_NAME = "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\MATT-555ROIDS\00_Text_Data_02\VOICE SPEAK LOGG.TXT"
WSHShell.Run """" + EXECUTE_FILE_NAME + """"
Set WSHShell = Nothing
End
End Sub

Private Sub Label112_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\#0 SMALL PROGS\Create on another drive\Create Dir Another Drive.exe", vbNormalFocus
End
End Sub

Private Sub Label113_Click()
Shell "C:\Program Files\Windows Media Player\wmplayer.exe", vbNormalFocus
End
End Sub

Private Sub Label116_Click()
ChDir "C:\Program Files\WordNet\2.1\bin"
Shell "C:\Program Files\WordNet\2.1\bin\wnb.exe", vbNormalFocus
End
End Sub

Private Sub Label114_Click()
TS = "D:\VB6\VB-NT\00_Best_VB_01\Date1991\Date1991.exe"
Shell TS, vbNormalFocus
End
End Sub

Private Sub Label115_Click()
'Shell "C:\Program Files\Norton 360\Engine\3.0.0.135\uiStub.exe ""/page {5ABC34AE-1037-4f5d-BF93-B2B74C80B5F7}""", vbNormalFocus
Shell "C:\Program Files\Norton Internet Security\Engine\18.6.0.29\uistub.exe", vbNormalFocus
End
End Sub

Private Sub Label117_Click()
Shell "C:\Program Files\Norton Ghost\Console\VProConsole_.exe", vbNormalFocus
End
End Sub

Private Sub Label118_Click()
Set oRunas = CreateObject("runas.clsrunas", "55-88-HAPPY")
With oRunas
    .sDomain = "WORKGROUP"
    .sUserName = "MATT_02"
    .sPassword = " "
    .sCommand = "C:\Program Files\Cyberdyne Systems\iCEBar.v3\"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End
End Sub

Private Sub Label12_Click()
Exit Sub
fr1 = FreeFile
Open "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\0TextData\WinAmpEXEFileActive.txt" For Input As #fr1
Line Input #fr1, TT$
Close #fr1

'TY = cProcesses.Convert(Pid, hwnd3, TT$)
'If FindWindow("Winamp v1.x", vbNullString) = 0 Then

'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt2"
    .sPassword = "matto28"
    .sCommand = TT$
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End
End Sub

Private Sub Label13_Click()
    
Shell "C:\Program Files\FeedDemon\FeedDemon.exe", vbMaximizedFocus
End
    
'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt3"
    .sPassword = " "
    .sCommand = "C:\Program Files\FeedDemon\FeedDemon.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing

End
    

End
End Sub

Private Sub Label14_Click()



Shell "D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.exe", vbNormalFocus
End

End Sub



Private Sub Label15_Click()

Shell "C:\Program Files\Google\Google Earth\client\googleearth.exe", vbMaximizedFocus
End

'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Google\Google Earth\client\googleearth.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End
End Sub

Private Sub Label16_Click()

'C:\Program Files\Adobe Photo UpLoader FaceBook\Adobe Photo Uploader for Facebook\Adobe Photo Uploader for Facebook.exe

Shell "C:\Program Files\Adobe Photo UpLoader FaceBook\Adobe Photo Uploader for Facebook\Adobe Photo Uploader for Facebook.exe", vbNormalFocus
End

'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Adobe Photo UpLoader FaceBook\Adobe Photo Uploader for Facebook\Adobe Photo Uploader for Facebook.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing

End

End Sub

Private Sub Label17_Click()
Shell sWindowsFolder + "\system32\magnify.exe", vbNormalFocus
End


'matt-555roids

Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = sWindowsFolder + "\system32\magnify.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End

End Sub

Private Sub Label18_Click()
'"C:\Program Files\Greatis\WinDowse\WinDowse.exe"


Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = " "
    .sCommand = "C:\Program Files\Greatis\WinDowse\WinDowse.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing

End

End Sub

Private Sub Label19_Click()

Shell sWindowsFolder + "\system32\TweakUI.exe", vbNormalFocus
End

End Sub

Private Sub Label2_Click()
'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Palm\Palm.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing


'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Palm\Hotsync.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing


MsgBox "Press Button in Sync USB Wire to Start"

End



End
End Sub

Private Sub Label20_Click()

Shell "C:\Program Files\SuperCopier2\SuperCopier2.exe", vbNormalFocus

End

End Sub

Private Sub Label21_Click()
Shell "C:\Program Files\Symantec\Norton PartitionMagic 8.0\PMagic.exe", vbNormalFocus
End
End Sub

Private Sub Label22_Click()
Shell sWindowsFolder + "\system32\osk.exe", vbNormalFocus
End

End Sub

Private Sub Label23_Click()
Shell "C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\Autoruns Microsoft\autoruns.exe", vbMaximizedFocus
End
End Sub

Private Sub Label24_Click()
If Dir(sWindowsFolder + "\system32\sndvol32.exe") <> "" Then
    Shell sWindowsFolder + "\system32\sndvol32.exe", vbNormalFocus
    End
End If
If Dir(sWindowsFolder + "\system32\sndvol.exe") <> "" Then
    Shell sWindowsFolder + "\system32\sndvol.exe", vbNormalFocus
    End
End If

End

End Sub

Private Sub Label25_Click()
Shell "C:\Program Files\Tweak-XP Pro\Tweak-xp.exe", vbNormalFocus
End

End Sub

Private Sub Label26_Click()
Shell "D:\VB6\VB-NT\00 Sort Anything\#Sort_Anything Lastest Best Ver 2011-04-17 - COPY - MULTI USE MENU  - AND COPY\Sort_AnyThing - COPY  - MULTI USE MENU - COPY.exe", vbNormalFocus
End

End Sub

Private Sub Label27_Click()
Shell "C:\Program Files\# NO INSTALL REQUIRED\01 www.System Internals.com\ProcessExplorer\procexp.exe", vbNormalFocus
End
End Sub

Private Sub Label28_Click()
Shell "C:\Program Files\ViceVersa Pro 2\ViceVersa.exe", vbNormalFocus
End
End Sub

Private Sub Label29_Click()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
Dim EXECUTE_FILE_NAME
EXECUTE_FILE_NAME = sMyDocsFolder + "\CD-ROM\01 Total List.txt"
WSHShell.Run """" + EXECUTE_FILE_NAME + """"
Set WSHShell = Nothing
End
End Sub

Private Sub Label3_Click()
'---------------------------
'ENC2001.EXE - Unable To Locate Component
'---------------------------
'This application has failed to start because MVCL53N.dll was not found. Re-installing the application may fix this problem.
'---------------------------
'OK
'---------------------------

ChDir "C:\Program Files\Microsoft Encarta\Encarta Encyclopedia 2001"
ChDir "C:\Program Files\Common Files\Microsoft Shared\Reference 2001"
'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt5"
    '.sPassword = "matto28"
    .sPassword = " "
    .sCommand = "C:\Program Files\Microsoft Encarta\Encarta Encyclopedia 2001\ENC2001.EXE"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing

End
End Sub

Private Sub Label30_Click()
Shell "C:\Program Files\WordNet\2.1\bin\", vbNormalFocus
End

End Sub

Private Sub Label31_Click()
Shell "C:\Program Files\WinHTTrack\WinHTTrack.exe", vbNormalFocus
End

End Sub

Private Sub Label32_Click()

'SAFE REMOVE HARDWARE
Shell sWindowsFolder + "\system32\RunDll32.exe shell32.dll,Control_RunDLL hotplug.dll"

End Sub

Private Sub Label33_Click()
Shell sWindowsFolder + "\system32\PowerCalc.exe", vbNormalFocus
End
End Sub

Private Sub Label34_Click()
ChDrive "M:\0 00 Art\Google IMAGE Downloader"
ChDir "M:\0 00 Art\Google IMAGE Downloader"
Shell "C:\Program Files\# NO INSTALL REQUIRED\GoogleImageDownloader-v1.1-win\GoogleImgDownloader.exe", vbMaximizedFocus
End

End Sub

Private Sub Label35_Click()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
Dim EXECUTE_FILE_NAME
EXECUTE_FILE_NAME = "D:\VB6\VB-NT\00_Best_VB_01\WinAmp Logger\00 Music Info Logger.txt"
WSHShell.Run """" + EXECUTE_FILE_NAME + """"
Set WSHShell = Nothing
End
End Sub

Private Sub Label36_Click()
Shell "C:\Program Files\Outlook Express\wab.exe", vbMaximizedFocus
End

End Sub

Private Sub Label37_Click()

'Shell sWindowsFolder + "\Notepad.exe", vbNormalFocus
'Shell sWindowsFolder + "\system32\NotePad\Notepad.exe", vbNormalFocus

VPATH = "C:\Program Files (x86)\Notepad++\notepad++.exe"
If Dir(VPATH) <> "" Then
    Shell VPATH, vbNormalFocus
Else
    VPATH = "C:\Program Files\Notepad++\notepad++.exe"
    Shell VPATH, vbNormalFocus
End If

End

End Sub

Private Sub Label38_Click()
Shell "C:\Program Files\TrueCrypt\TrueCrypt.exe", vbNormalFocus
End

End Sub

Private Sub Label39_Click()

Shell "C:\Program Files\Windows NT\Accessories\wordpad.exe", vbMaximizedFocus
End


End Sub

Private Sub Label4_Click()

'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Olympus\Digital Wave Player\DWP.exe"
'    .sCommand = "Program you want to run i.e. c:\winnt\notepad.exe remember"
'             you must explictly use the relevant program for example if you wanted
'             to open a text document you can't use c:\text.txt or "notepad c:\text.txt"
'             you would have to use
'             "c:\winnt\notepad.exe c:\text.txt"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing

End

End Sub

Private Sub Label40_Click()
Shell "C:\Program Files\Easy CD-DA Extractor 5.1\ezcddaxfc.exe", vbMaximizedFocus
End

End Sub

Private Sub Label41_Click()
Shell "C:\Program Files\MIKSOFT\Mobile Media Converter\mmc.exe", vbNormalFocus
End
End Sub

Private Sub Label42_Click()
Me.WindowState = 1
    'Shell "C:\Program Files\Norton 360\Engine\3.0.0.135\uiStub.exe ""/page {5ABC34AE-1037-4f5d-BF93-B2B74C80B5F7}""", vbNormalFocus
    'End
    CLOSEPROG = FindWindow("MozillaWindowClass", vbNullString)
    If CLOSEPROG = 0 Then Shell "C:\Program Files\Mozilla Firefox\firefox.exe ", vbNormalFocus
    
    Do
        Sleep 1000
        CLOSEPROG = FindWindow("MozillaWindowClass", vbNullString)

    Loop Until CLOSEPROG > 0
    Sleep 5000
    
    Shell "C:\Program Files\Mozilla Firefox\firefox.exe /n, /e, " + "http://www.nationwide.co.uk/default.htm"
    Shell "C:\Program Files\Mozilla Firefox\firefox.exe " + "http://www.vodafone.co.uk/cgi-bin/COUK/portal/ep/home.do"
    Shell "C:\Program Files\Mozilla Firefox\firefox.exe " + "http://www.co-operativebank.co.uk/servlet/Satellite/1193206375355,CFSweb/Page/Bank"
    Shell "C:\Program Files\Mozilla Firefox\firefox.exe " + "https://www.mybank.alliance-leicester.co.uk/index.asp"
    Shell "C:\Program Files\Mozilla Firefox\firefox.exe " + "https://www.bankcardservices.co.uk/NASApp/NetAccessXX/WelcomeScreen?country=uk&language=en&org=90"
    Shell "C:\Program Files\Mozilla Firefox\firefox.exe " + "http://uk.mc861.mail.yahoo.com/mc/welcome?.partner=bt-1&.rand=6fpal5m346kh2#_pg=showFolder&fid=Inbox&order=down&tt=11289&pSize=200&.rand=2083009526&hash=2c2823b2d0f4eff82f4acd733ca05bf7&.jsrand=5725228"
    Shell "C:\Program Files\Mozilla Firefox\firefox.exe https://market.android.com/"
    Shell "C:\Program Files\Mozilla Firefox\firefox.exe http://store.ovi.com/"
     Shell "C:\Program Files\Mozilla Firefox\firefox.exe http://192.168.1.65/"  '    AREO
'    Shell "C:\Program Files\Mozilla Firefox\firefox.exe "
'    Shell "C:\Program Files\Mozilla Firefox\firefox.exe "
'    Shell "C:\Program Files\Mozilla Firefox\firefox.exe "
'    Shell "C:\Program Files\Mozilla Firefox\firefox.exe "
'    Shell "C:\Program Files\Mozilla Firefox\firefox.exe "
'    Shell "C:\Program Files\Mozilla Firefox\firefox.exe "

     Shell "C:\Program Files\Mozilla Firefox\firefox.exe -extoff http://www.homemove.org.uk/Login.aspx"

    Shell "C:\Program Files\Google\Google Earth\client\googleearth.exe", vbNormalFocus
    Shell "D:\VB6\VB-NT\00_Best_VB_01\FileDownloader\FileDownloader.exe", vbNormalFocus
    
    Shell "C:\WINDOWS\system32\OSK.exe"
    
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE -extoff  http://www.rightmove.co.uk/rss/property-to-rent/find.html?locationIdentifier=OUTCODE%5E240&minPrice=500&maxPrice=700&maxDaysSinceAdded=3&radius=5.0"

End

End Sub

Private Sub Label43_Click()

Shell "C:\Program Files\FastStone Image Viewer\FSViewer.exe", vbMaximizedFocus
End
End Sub

Private Sub Label44_Click()
Shell "C:\Program Files\WinRAR\WinRAR.exe", vbNormalFocus
End

End Sub

Private Sub Label45_Click()
Shell "C:\Program Files\SuperCopier2\SC2Config.exe", vbNormalFocus
Me.WindowState = 1
MsgBox "PRESS WHEN DONE SUPER COPIER"
End
End Sub

Private Sub Label47_Click()
Shell "C:\Program Files\FileMon\FileMon.exe", vbNormalFocus

End

End Sub

Private Sub Label48_Click()
TS = "D:\VB6\VB-NT\00_Best_VB_01\Date1991\Date1991.exe"
Shell TS, vbNormalFocus
End
End Sub

Private Sub Label49_Click()
Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE """ + "D:\VB6\VB-NT\00 MP3 Sound\MP3 Tags File Title ID3v2\MP3 Tags File Title ID3v2.vbp""", vbNormalFocus
End
End Sub

Private Sub Label5_Click()
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
With oRunas
    .sDomain = "WORKGROUP"
    .sUserName = "Matt"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Cool2000\cool2000.exe"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End
End Sub

Private Sub Label50_Click()
Shell "D:\VB6\VB-NT\00 MP3 Sound\Sort_Anything_ScanPath_Mp3_Dummy_Id_Tag_For_Winamp_Playlist\Mp3_Dummy_Id_Tag_For_Winamp_Playlist.exe", vbNormalFocus
End
End Sub

Private Sub Label51_Click()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
Dim EXECUTE_FILE_NAME
EXECUTE_FILE_NAME = "D:\VB6\VB-NT\#1 Documents\Programming Snippets.txt"
WSHShell.Run """" + EXECUTE_FILE_NAME + """"
Set WSHShell = Nothing
End
End Sub

Private Sub Label52_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\Run Fav Progs\Run Fav Programs.exe", vbNormalFocus
End
End Sub

Private Sub Label53_Click()
Shell "D:\VB6\VB-NT\00 MP3 Sound\MP3 Winamp PlayList Generator ID3v2\MP3 Winamp PlayList Generator ID3V2.exe", vbNormalFocus
End

End Sub

Private Sub Label54_Click()

Shell "D:\VB6\VB-NT\00 MP3 Sound\MP3 Tags File Title ID3v2\MP3 Tags File Title ID3V2.exe", vbNormalFocus
End
End Sub

Private Sub Label55_Click()

Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE """ + "D:\VB6\VB-NT\00 MP3 Sound\Sort_Anything_ScanPath_Mp3_Dummy_Id_Tag_For_Winamp_Playlist\Mp3_Dummy_Id_Tag_For_Winamp_Playlist.vbp""", vbNormalFocus
End

End Sub

Private Sub Label56_Click()

Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE """ + "D:\VB6\VB-NT\00 MP3 Sound\MP3 Winamp PlayList Generator ID3v2\MP3 Winamp PlayList Generator ID3v2.vbp""", vbNormalFocus
End
End Sub

Private Sub Label59_Click()
Shell "C:\Program Files\OE-QuoteFix\OELaunch.exe", vbNormalFocus
End
End Sub

Private Sub Label6_Click()
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
With oRunas
    .sDomain = "WORKGROUP"
    .sUserName = "Admin"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Cool2000\cool2000.exe"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End

End Sub

Private Sub Label60_Click()


Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
With oRunas
    .sDomain = "WORKGROUP"
    .sUserName = "Matt5"
    .sPassword = " "
    .sCommand = "C:\Program Files\Ultralingua\Ultralingua 4\ultralingua.exe"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End

End
End Sub

Private Sub Label61_Click()
Shell "C:\Program Files\CamSplitter\camsplitter.exe /standalone", vbNormalFocus
End
End Sub

Private Sub Label62_Click()
    
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
Dim EXECUTE_FILE_NAME
EXECUTE_FILE_NAME = "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\MATT-555ROIDS\00_Text_Data_02\VOICE SPEAK LOGG.TXT"
WSHShell.Run """" + EXECUTE_FILE_NAME + """"
Set WSHShell = Nothing
End

End Sub

Private Sub Label63_Click()

Shell "D:\VB6\VB-NT\00_Best_VB_04\Vcap_Src_WebCam\vbVidCapWebCam.exe", vbNormalFocus

End
End Sub

Private Sub Label65_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\Cid-RunMe.exe", vbNormalFocus
End
End Sub

Private Sub Label66_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\Tidal.exe", vbNormalFocus
End
End Sub

Private Sub Label67_Click()
Shell "C:\Program Files\# NO INSTALL REQUIRED\mp3DirectCut\Mp3DirectCut.exe", vbNormalFocus

End
End Sub

Private Sub Label68_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\URL Logger\URL Logger.exe", vbNormalFocus
End
End Sub

Private Sub Label69_Click()
Shell "C:\Program Files\# NO INSTALL REQUIRED\FAT32\Fat32Formatter.bat", vbNormalFocus
End
End Sub

Private Sub Label7_Click()
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
With oRunas
    .sDomain = "WORKGROUP"
    .sUserName = "Administrator"
    .sPassword = "matto28"
    .sCommand = "C:\Program Files\Cool2000\cool2000.exe"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End
End Sub

Private Sub Label71_Click()
Shell "C:\Program Files\ProcessTamer\ProcessTamerConfigurator.exe", vbNormalFocus
End
End Sub

Private Sub Label72_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_04\CLOCK #M02 SCR\CLOCK #M02 SCR.exe", vbNormalFocus
End
End Sub

Private Sub Label73_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\TIMER CODE\TIMER CODE.exe", vbNormalFocus
End
End Sub

Private Sub Label74_Click()

Shell "D:\VB6\VB-NT\00_Best_VB_01\Drives Gig_MINIMIZE VERSION\Drives_Gig.exe", vbNormalFocus
End

End Sub

Private Sub Label75_Click()
Shell "D:\VB6\VB-NT\00 Sort Anything\#Sort_Anything Lastest Best Ver 2009-04-25 - MULTI USE MENU\Sort_AnyThing - MULTI USE MENU.exe", vbNormalFocus
End

End Sub

Private Sub Label76_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\00 Games\Breakout\BreakOut.exe", vbNormalFocus
End

End Sub


Private Sub Label77_Click()
Shell "C:\Program Files\# NO INSTALL REQUIRED\SHELL EX VIEW\shexview.exe", vbNormalFocus
End

End Sub

Private Sub Label78_Click()

Shell "D:\VB6\VB-NT\00_Best_VB_01\VU_METER_LOGGER\VU METER LOGGER.exe", vbMinimizedNoFocus

'Shell "D:\VB6\VB-NT\00_Best_VB_01\URL Logger\URL Logger.exe", vbNormalFocus
End
End Sub

Private Sub Label79_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\ClipBoard Logger\ClipBoard Logger.exe", vbMinimizedNoFocus
End
End Sub

Private Sub Label8_Click()
ChDrive App.Path
ChDir App.Path
Shell "D:\VB6\VB-NT\00_Best_VB_01\WinAmp Logger\WinAmp Logger.exe", vbNormalNoFocus
End
End Sub

Private Sub Label80_Click()
Shell "C:\Program Files\Windows Media Player\mplayer2.exe"
End
End Sub

Private Sub Label81_Click()
i = SetScreenSaverState(False, False)
End
End Sub

Private Sub Label82_Click()
    Shell "C:\Program Files\Mozilla Firefox\firefox.exe", vbNormalFocus
    End
End Sub

Private Sub Label83_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\Run Fav Progs\Run Fav Programs.exe", vbNormalNoFocus
End
End Sub

Private Sub Label84_Click()
Shell "C:\Program Files\Ahead\Nero ShowTime\ShowTime.exe", vbNormalFocus
End
End Sub

Private Sub Label86_Click()
Shell "C:\Program Files\# NO INSTALL REQUIRED\SHELL MENU View\shmnview.exe", vbNormalFocus
End
End Sub

Private Sub Label88_Click()
Shell "D:\# MY DOCS\# 01 My Documents\CallerID\2PINNO'S.TXT", vbNormalFocus
End
End Sub

Private Sub Label89_Click()
Shell "C:\Program Files\ProcessTamer\ProcessTamerTray.exe", vbNormalFocus
End
End Sub

'Shell "D:\VB6\VB-NT\00 MP3 Sound\DSS_Olympus\DSS_Olympus.exe"
'End


Private Sub Label9_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\#0 SMALL PROGS\50LastEXE's\50LastEXE.exe", vbNormalFocus
End
End Sub

Private Sub Label92_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\DARK\DARK.exe", vbNormalFocus
End
End Sub

Private Sub Label93_Click()
'Shell "C:\Program Files\# NO INSTALL REQUIRED\WebImageGrab_pc\WIG light.exe", vbNormalFocus
ChDrive "M:"
ChDir "M:\0 00 Art\Google Downloader\"
Shell "M:\0 00 Art\Google Downloader\WebImageGrab_pro_pc\WIG PRO.exe", vbNormalFocus
End
End Sub

Private Sub Label95_Click()
'Screen SAVER


i = SetScreenSaverState(True, True)
End
End Sub

Private Sub Label96_Click()
    Shell "rundll32.exe ""shell32.dll,Control_RunDLL """ + sWindowsFolder + "\system32\bthprops.cpl""", vbNormalFocus
    End
End Sub

Private Sub Label97_Click()
Shell "D:\VB6\VB-NT\#1 Documents\Programming Snippets.txt", vbNormalFocus
End
End Sub

Private Sub MNU_VB_Click()
VBPATH = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
If Dir(VBPATH) = "" Then
    VBPATH = "C:\Program Files (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
End If

Shell VBPATH + " """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
End
End Sub

Private Sub MNU_VB_FOLDER_Click()

Shell "EXPLORER /SELECT, " + App.Path + "\" + App.EXEName + ".VBP", vbMaximizedFocus

End Sub

Private Sub MNU_VBEXE_Click()

Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
Dim EXECUTE_FILE_NAME

'EXECUTE_FILE_NAME = "C:WINDOWS\EXPLORER.EXE " + EXECUTE_PARAM
EXECUTE_FILE_NAME = App.Path + "\File of EXE Programs in VB.txt"

WSHShell.Run """" + EXECUTE_FILE_NAME + """"
Set WSHShell = Nothing

End Sub

Private Sub MNU_VERSION_Click()
MNU_VERSION.Caption = "Ver. " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
End Sub

Private Sub Timer_ART_TRIP_WIRE_Timer()

'i = LastModifiedToCurrent("K:\TEMP\ART_PROG_TRIP_WIRE.txt")

Dim HH As Long, OTSS As Long
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
If HH = 0 And XHH = True Then End
If HH = 0 Then Exit Sub


If FS.FileExists("K:\TEMP\ART_PROG_TRIP_WIRE.txt") = False And XHH = True Then End

Set F = FS.GetFile("K:\TEMP\ART_PROG_TRIP_WIRE.txt")
i = F.DateLastModified
If OLDi <> i Then iTRIP = Now + TimeSerial(0, 0, 15)
OLDi = i
'Debug.Print i, OLDi

If iTRIP > Now Then Exit Sub

'TT = cProcesses.Convert(HH, OTSS, cnFromhWnd Or cnToProcessID)
OTSS = -1
TT = cProcesses.GetEXEID(OTSS, "D:\VB6\VB-NT\00_Best_VB_04\JPEG-Encoder-ART\Jpeg Slides PJpeg ART.exe")

If TT = True Then Process_Kill (OTSS)

XX = 0
Do
    XX = XX + 1
    OTSS = -1
    TT = cProcesses.GetEXEID(OTSS, "D:\VB6\VB-NT\00_Best_VB_04\JPEG-Encoder-ART\Jpeg Slides PJpeg ART.exe")
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

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************






' return the Enabled state of the screen saver

Function GetScreenSaverState() As Boolean
    Dim Result As Long
    SystemParametersInfo SPI_GETSCREENSAVEACTIVE, 0, Result, 0
    GetScreenSaverState = (Result <> 0)
End Function

' enable or disable the screen saver
'
' if second argument is true, it writes changes in user's profile
' returns True if the operation was successful, False otherwise

Function SetScreenSaverState(ByVal enabled As Boolean, _
    Optional ByVal PermanentChange As Boolean) As Boolean
    Dim fuWinIni As Long
    If PermanentChange Then
        fuWinIni = SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE
    End If
    SetScreenSaverState = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, enabled, _
        ByVal 0&, fuWinIni) <> 0
End Function


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






