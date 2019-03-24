VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   Caption         =   "Matt Run AS"
   ClientHeight    =   10656
   ClientLeft      =   60
   ClientTop       =   -1668
   ClientWidth     =   15240
   FillStyle       =   0  'Solid
   Icon            =   "Matt_RunAS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10656
   ScaleWidth      =   15240
   Begin VB.ListBox List1 
      Height          =   432
      Left            =   696
      Sorted          =   -1  'True
      TabIndex        =   141
      Top             =   60
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Timer Timer_1_MINUTE 
      Interval        =   1000
      Left            =   5004
      Top             =   4176
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1404
      Top             =   588
   End
   Begin VB.Timer Timer_ART_TRIP_WIRE 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   900
      Top             =   444
   End
   Begin VB.Label Label_800 
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
      Index           =   71
      Left            =   3432
      TabIndex        =   140
      Top             =   4776
      Width           =   144
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "FIREFOX USER 01"
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
      Index           =   142
      Left            =   12720
      TabIndex        =   118
      Top             =   8808
      Width           =   2076
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB ----- 1 FOLDER"
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
      Index           =   141
      Left            =   24
      TabIndex        =   117
      Top             =   6504
      Width           =   2040
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "NERO WAVE EDITOR"
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
      Index           =   140
      Left            =   4080
      TabIndex        =   116
      Top             =   2040
      Width           =   2424
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Grant NTFS Access"
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
      Index           =   139
      Left            =   48
      TabIndex        =   115
      Top             =   5544
      Width           =   2232
   End
   Begin VB.Label Label_800 
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
      Index           =   138
      Left            =   132
      TabIndex        =   114
      Top             =   108
      Width           =   144
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB ---- 1"
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
      Index           =   137
      Left            =   2208
      TabIndex        =   113
      Top             =   6516
      Width           =   900
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "LOGITECH UNIFYER"
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
      Index           =   136
      Left            =   72
      TabIndex        =   112
      Top             =   5220
      Width           =   2328
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB -- MASSIVE"
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
      Index           =   135
      Left            =   4776
      TabIndex        =   111
      Top             =   7560
      Width           =   1680
   End
   Begin VB.Label Label_800 
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
      Index           =   134
      Left            =   240
      TabIndex        =   110
      Top             =   120
      Width           =   144
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   133
      Left            =   1836
      TabIndex        =   109
      Top             =   60
      Width           =   840
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   132
      Left            =   8004
      TabIndex        =   108
      Top             =   468
      Width           =   6648
   End
   Begin VB.Label Label_800 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "MS OFFICE TOOLBAR ----"
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
      Index           =   131
      Left            =   12732
      TabIndex        =   107
      Top             =   8436
      Width           =   2952
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   130
      Left            =   3024
      TabIndex        =   106
      Top             =   7740
      Width           =   1476
   End
   Begin VB.Label Label_800 
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
      Index           =   129
      Left            =   9990
      TabIndex        =   105
      Top             =   1200
      Width           =   1545
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   128
      Left            =   10212
      TabIndex        =   104
      Top             =   1680
      Width           =   2736
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   127
      Left            =   4752
      TabIndex        =   103
      Top             =   10032
      Width           =   1632
   End
   Begin VB.Label Label_800 
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
      Index           =   126
      Left            =   12912
      TabIndex        =   102
      Top             =   4908
      Width           =   3228
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   125
      Left            =   96
      TabIndex        =   101
      Top             =   6168
      Width           =   3240
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   124
      Left            =   12792
      TabIndex        =   100
      Top             =   3480
      Width           =   2064
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   123
      Left            =   7680
      TabIndex        =   99
      Top             =   10716
      Width           =   2028
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   122
      Left            =   6720
      TabIndex        =   98
      Top             =   7560
      Width           =   3600
   End
   Begin VB.Label Label_800 
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
      Index           =   121
      Left            =   7920
      TabIndex        =   97
      Top             =   840
      Width           =   3465
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   120
      Left            =   4488
      TabIndex        =   96
      Top             =   60
      Width           =   1332
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   119
      Left            =   4728
      TabIndex        =   95
      Top             =   6228
      Width           =   1992
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "EXPLORER MAXIMIZED"
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
      Index           =   118
      Left            =   60
      TabIndex        =   94
      Top             =   5856
      Width           =   2664
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   117
      Left            =   4704
      TabIndex        =   93
      Top             =   5580
      Width           =   1920
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "SHELL - KeyBoard Loader"
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
      Index           =   116
      Left            =   7740
      TabIndex        =   92
      Top             =   9312
      Width           =   2976
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   115
      Left            =   8400
      TabIndex        =   91
      Top             =   2040
      Width           =   1908
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "P2ROCESS EXPLOR - TASK MANAGER"
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
      Index           =   114
      Left            =   24
      TabIndex        =   90
      Top             =   8292
      Width           =   4464
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "SHELL - Explorer Loader"
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
      Index           =   112
      Left            =   7716
      TabIndex        =   89
      Top             =   9672
      Width           =   2820
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB ----- BATCH COMPLIER"
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
      Index           =   111
      Left            =   11160
      TabIndex        =   88
      Top             =   120
      Width           =   3024
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   110
      Left            =   9240
      TabIndex        =   87
      Top             =   3120
      Width           =   3660
   End
   Begin VB.Label Label_800 
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
      Index           =   109
      Left            =   13080
      TabIndex        =   86
      Top             =   1680
      Width           =   2064
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   108
      Left            =   4368
      TabIndex        =   85
      Top             =   816
      Width           =   2664
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Notepad++"
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
      Index           =   107
      Left            =   96
      TabIndex        =   84
      Top             =   7176
      Width           =   1236
   End
   Begin VB.Label Label_800 
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
      Index           =   106
      Left            =   4380
      TabIndex        =   83
      Top             =   1548
      Width           =   660
   End
   Begin VB.Label Label_800 
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
      Index           =   105
      Left            =   11760
      TabIndex        =   82
      Top             =   2760
      Width           =   2565
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "DIGI BOX TERESTRIAL MEDIA"
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
      Index           =   104
      Left            =   7212
      TabIndex        =   81
      Top             =   7932
      Width           =   3480
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   103
      Left            =   2808
      TabIndex        =   80
      Top             =   3564
      Width           =   2304
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   102
      Left            =   13008
      TabIndex        =   79
      Top             =   2400
      Width           =   2328
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "KLS BACKUP MAIL"
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
      Index           =   101
      Left            =   11856
      TabIndex        =   78
      Top             =   4476
      Width           =   2160
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   100
      Left            =   8112
      TabIndex        =   27
      Top             =   60
      Width           =   2208
   End
   Begin VB.Label Label_800 
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
      Index           =   99
      Left            =   2880
      TabIndex        =   37
      Top             =   45
      Width           =   1470
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   98
      Left            =   8112
      TabIndex        =   38
      Top             =   60
      Width           =   2028
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "SHELL - Loader"
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
      Index           =   97
      Left            =   7692
      TabIndex        =   76
      Top             =   10008
      Width           =   1788
   End
   Begin VB.Label Label_800 
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
      Index           =   96
      Left            =   2748
      TabIndex        =   75
      Top             =   420
      Width           =   1788
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   95
      Left            =   1728
      TabIndex        =   39
      Top             =   1512
      Width           =   2112
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   94
      Left            =   12792
      TabIndex        =   73
      Top             =   9192
      Width           =   1020
   End
   Begin VB.Label Label_800 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "ICE BAR"
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
      Index           =   93
      Left            =   1584
      TabIndex        =   74
      Top             =   420
      Width           =   960
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   92
      Left            =   3960
      TabIndex        =   72
      Top             =   1176
      Width           =   2784
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   91
      Left            =   12864
      TabIndex        =   71
      Top             =   5976
      Width           =   1704
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB ClipBoard Logger && URL LOGGER"
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
      Index           =   90
      Left            =   4764
      TabIndex        =   70
      Top             =   7260
      Width           =   4296
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   89
      Left            =   2820
      TabIndex        =   69
      Top             =   2472
      Width           =   2748
   End
   Begin VB.Label Label_800 
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
      Index           =   88
      Left            =   3468
      TabIndex        =   77
      Top             =   4836
      Width           =   144
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   87
      Left            =   9948
      TabIndex        =   68
      Top             =   4476
      Width           =   1824
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   86
      Left            =   4704
      TabIndex        =   67
      Top             =   5952
      Width           =   2160
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   85
      Left            =   10452
      TabIndex        =   66
      Top             =   2040
      Width           =   4836
   End
   Begin VB.Label Label_800 
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
      Index           =   84
      Left            =   4608
      TabIndex        =   65
      Top             =   5244
      Width           =   1248
   End
   Begin VB.Label Label_800 
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
      Index           =   83
      Left            =   4728
      TabIndex        =   64
      Top             =   10740
      Width           =   2364
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   82
      Left            =   4740
      TabIndex        =   63
      Top             =   10380
      Width           =   2472
   End
   Begin VB.Label Label_800 
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
      Index           =   81
      Left            =   3276
      TabIndex        =   62
      Top             =   4764
      Width           =   144
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   80
      Left            =   4740
      TabIndex        =   61
      Top             =   9396
      Width           =   2868
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   79
      Left            =   10176
      TabIndex        =   60
      Top             =   6276
      Width           =   1884
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   78
      Left            =   4728
      TabIndex        =   59
      Top             =   6924
      Width           =   1752
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   77
      Left            =   12840
      TabIndex        =   58
      Top             =   6324
      Width           =   2016
   End
   Begin VB.Label Label_800 
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
      Index           =   76
      Left            =   4740
      TabIndex        =   57
      Top             =   9708
      Width           =   1416
   End
   Begin VB.Label Label_800 
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
      Index           =   75
      Left            =   1680
      TabIndex        =   56
      Top             =   1176
      Width           =   2112
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   74
      Left            =   4740
      TabIndex        =   55
      Top             =   9012
      Width           =   1548
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "EASY MPEG"
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
      Index           =   73
      Left            =   11112
      TabIndex        =   119
      Top             =   9936
      Width           =   1404
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VOICE TXT LOGG"
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
      Index           =   72
      Left            =   8136
      TabIndex        =   120
      Top             =   2388
      Width           =   2064
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   70
      Left            =   8400
      TabIndex        =   53
      Top             =   60
      Width           =   2424
   End
   Begin VB.Label Label_800 
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
      Index           =   69
      Left            =   11748
      TabIndex        =   52
      Top             =   804
      Width           =   1956
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   68
      Left            =   10368
      TabIndex        =   51
      Top             =   2400
      Width           =   2508
   End
   Begin VB.Label Label_800 
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
      Index           =   67
      Left            =   8544
      TabIndex        =   50
      Top             =   1572
      Width           =   1440
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "WINDOWS MEDIA PLAYER"
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
      Index           =   66
      Left            =   2340
      TabIndex        =   49
      Top             =   3240
      Width           =   3024
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "SHELL - Execute Loader"
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
      Index           =   65
      Left            =   7692
      TabIndex        =   48
      Top             =   10332
      Width           =   2772
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "BATTERY MONITOR"
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
      Index           =   64
      Left            =   9240
      TabIndex        =   47
      Top             =   2760
      Width           =   2325
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   348
      Index           =   63
      Left            =   8340
      TabIndex        =   121
      Top             =   6360
      Width           =   396
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   62
      Left            =   7392
      TabIndex        =   122
      Top             =   3612
      Width           =   1584
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "FREE COMMMANDER MULTI"
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
      Index           =   61
      Left            =   5160
      TabIndex        =   42
      Top             =   1680
      Width           =   3288
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Program Snippalists"
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
      Index           =   60
      Left            =   12
      TabIndex        =   123
      Top             =   7476
      Width           =   2292
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   348
      Index           =   59
      Left            =   8340
      TabIndex        =   54
      Top             =   6720
      Width           =   396
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   58
      Left            =   8328
      TabIndex        =   124
      Top             =   4884
      Width           =   3672
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   57
      Left            =   8316
      TabIndex        =   125
      Top             =   5244
      Width           =   4416
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   56
      Left            =   4788
      TabIndex        =   45
      Top             =   7896
      Width           =   2172
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   55
      Left            =   4788
      TabIndex        =   44
      Top             =   8628
      Width           =   2604
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   54
      Left            =   4788
      TabIndex        =   41
      Top             =   8268
      Width           =   3348
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   348
      Index           =   53
      Left            =   8340
      TabIndex        =   126
      Top             =   5988
      Width           =   396
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   52
      Left            =   8328
      TabIndex        =   127
      Top             =   5592
      Width           =   3240
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   51
      Left            =   4644
      TabIndex        =   46
      Top             =   4896
      Width           =   1632
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   50
      Left            =   3360
      TabIndex        =   128
      Top             =   5568
      Width           =   1068
   End
   Begin VB.Label Label_800 
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
      Index           =   49
      Left            =   5280
      TabIndex        =   129
      Top             =   840
      Width           =   144
   End
   Begin VB.Label Label_800 
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
      Index           =   48
      Left            =   2136
      TabIndex        =   43
      Top             =   2856
      Width           =   2772
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   47
      Left            =   11196
      TabIndex        =   130
      Top             =   9276
      Width           =   948
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "FastStone Image Viewer"
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
      Index           =   46
      Left            =   11172
      TabIndex        =   131
      Top             =   10668
      Width           =   2724
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Norton 360"
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
      Index           =   45
      Left            =   7392
      TabIndex        =   40
      Top             =   3252
      Width           =   1248
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   44
      Left            =   12864
      TabIndex        =   36
      Top             =   5616
      Width           =   2892
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   43
      Left            =   11196
      TabIndex        =   35
      Top             =   10332
      Width           =   2904
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   42
      Left            =   1356
      TabIndex        =   34
      Top             =   6912
      Width           =   1308
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   41
      Left            =   10992
      TabIndex        =   132
      Top             =   8460
      Width           =   1524
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   40
      Left            =   144
      TabIndex        =   133
      Top             =   6888
      Width           =   1152
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   39
      Left            =   12696
      TabIndex        =   134
      Top             =   9612
      Width           =   1920
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   38
      Left            =   12
      TabIndex        =   33
      Top             =   4068
      Width           =   2628
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   37
      Left            =   11544
      TabIndex        =   32
      Top             =   6804
      Width           =   2904
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   36
      Left            =   2808
      TabIndex        =   31
      Top             =   6912
      Width           =   1608
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   35
      Left            =   6708
      TabIndex        =   30
      Top             =   4476
      Width           =   3156
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   34
      Left            =   10920
      TabIndex        =   29
      Top             =   8892
      Width           =   1572
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   33
      Left            =   11124
      TabIndex        =   28
      Top             =   9588
      Width           =   1404
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "WINAMP MUSIC 01"
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
      Index           =   32
      Left            =   48
      TabIndex        =   135
      Top             =   4428
      Width           =   2136
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "WINAMP ONE"
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
      Index           =   31
      Left            =   8100
      TabIndex        =   136
      Top             =   840
      Width           =   1560
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   30
      Left            =   72
      TabIndex        =   26
      Top             =   4800
      Width           =   1776
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   29
      Left            =   2556
      TabIndex        =   25
      Top             =   7332
      Width           =   1404
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "P1ROCESS EXPLORER"
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
      Index           =   28
      Left            =   120
      TabIndex        =   24
      Top             =   7860
      Width           =   2652
   End
   Begin VB.Label Label_800 
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
      Index           =   27
      Left            =   8130
      TabIndex        =   137
      Top             =   1155
      Width           =   1545
   End
   Begin VB.Label Label_800 
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
      Index           =   26
      Left            =   2928
      TabIndex        =   23
      Top             =   8748
      Width           =   1716
   End
   Begin VB.Label Label_800 
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
      Index           =   25
      Left            =   1236
      TabIndex        =   22
      Top             =   2496
      Width           =   1416
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   23
      Left            =   12
      TabIndex        =   21
      Top             =   2256
      Width           =   1116
   End
   Begin VB.Label Label_800 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB CLIP TYPE"
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
      Index           =   22
      Left            =   4752
      TabIndex        =   20
      Top             =   6600
      Width           =   1620
   End
   Begin VB.Label Label_800 
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
      Index           =   21
      Left            =   2832
      TabIndex        =   138
      Top             =   2052
      Width           =   1116
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   20
      Left            =   12
      TabIndex        =   19
      Top             =   3168
      Width           =   2040
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   19
      Left            =   12
      TabIndex        =   18
      Top             =   2844
      Width           =   1776
   End
   Begin VB.Label Label_800 
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
      Index           =   18
      Left            =   5688
      TabIndex        =   17
      Top             =   3552
      Width           =   1164
   End
   Begin VB.Label Label_800 
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
      Index           =   17
      Left            =   5676
      TabIndex        =   16
      Top             =   2460
      Width           =   2340
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   16
      Left            =   6900
      TabIndex        =   15
      Top             =   1188
      Width           =   1056
   End
   Begin VB.Label Label_800 
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
      Index           =   15
      Left            =   6816
      TabIndex        =   14
      Top             =   468
      Width           =   1056
   End
   Begin VB.Label Label_800 
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
      Index           =   14
      Left            =   6090
      TabIndex        =   13
      Top             =   15
      Width           =   1500
   End
   Begin VB.Label Label_800 
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
      Index           =   13
      Left            =   5676
      TabIndex        =   12
      Top             =   3192
      Width           =   996
   End
   Begin VB.Label Label_800 
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
      Index           =   12
      Left            =   5676
      TabIndex        =   11
      Top             =   2832
      Width           =   2148
   End
   Begin VB.Label Label_800 
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
      Index           =   11
      Left            =   6660
      TabIndex        =   10
      Top             =   2028
      Width           =   1416
   End
   Begin VB.Label Label_800 
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
      Index           =   10
      Left            =   1776
      TabIndex        =   9
      Top             =   2160
      Width           =   924
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   9
      Left            =   1320
      TabIndex        =   8
      Top             =   1860
      Width           =   1296
   End
   Begin VB.Label Label_800 
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
      Index           =   8
      Left            =   45
      TabIndex        =   7
      Top             =   1875
      Width           =   1080
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   7
      Left            =   12
      TabIndex        =   139
      Top             =   3612
      Width           =   2628
   End
   Begin VB.Label Label_800 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Cool 3"
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
      Index           =   6
      Left            =   3372
      TabIndex        =   6
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   5
      Left            =   516
      TabIndex        =   5
      Top             =   1452
      Width           =   912
   End
   Begin VB.Label Label_800 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "COOL EDIT 1"
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
      Index           =   4
      Left            =   60
      TabIndex        =   4
      Top             =   1128
      Width           =   1512
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   3
      Left            =   4692
      TabIndex        =   3
      Top             =   444
      Width           =   816
   End
   Begin VB.Label Label_800 
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
      Index           =   2
      Left            =   84
      TabIndex        =   2
      Top             =   792
      Width           =   3084
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   1
      Left            =   13092
      TabIndex        =   1
      Top             =   3120
      Width           =   1452
   End
   Begin VB.Label Label_800 
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
      Height          =   288
      Index           =   0
      Left            =   5652
      TabIndex        =   0
      Top             =   456
      Width           =   828
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu Mnu_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDERING"
   End
   Begin VB.Menu MNU_ADMINISTRATOR 
      Caption         =   "ADMINISTRATOR"
   End
   Begin VB.Menu MNU_RUNAS_DOWSE 
      Caption         =   "TEST RUNAS ON WINDOSE"
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
Dim XVB_DATE
    
Dim FS
    
'---------------------
'SEARCH 20 FEB MON
'STUCK HERE WORKING ON
'---------------------
    
Dim X_COL_1(200), X_COL_2(200)
    
    
Dim SET_WIN_STATE
    
    
Dim I1, i2____, i3_Exe
Dim First_Run_Done
'--------------------------
    '----------------------
    'TEMP OUT WON'T COMPILE
    '----------------------
    'SystemParametersInfo SPI_GETSCREENSAVEACTIVE, 0, Result, 0
    
    '----------------------
    'TEMP OUT WON'T COMPILE
    '----------------------
    'SetScreenSaverState = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, enabled, _
        ByVal 0&, fuWinIni) <> 0
'--------------------------





Sub REM_LINE()

'"C:\Program Files\Runtime Software\DriveImage XML\dixml.exe"

'WANT BLUE TOOTH ON HERE

'ERROR AT LOAD FOR HERE

'Begin VB.Menu KILLMENU
'  Caption         =   "KILL MENU"
'  Begin VB.Menu KILL_JEPG_ART
'        Caption         =   "KILL JEPG ART SLIDESHOW WHEN FROZEN"
'  End
'End
'Begin VB.Menu MNU_VBEXE
'  Caption         =   "NOTEPAD VB EXE"
'End


End Sub








Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If IsIDE = True Then
    If KeyCode = 27 Then: Beep: End
    '----------------
    '27 IS ESCPAE KEY
    'CODER
    '----------------


End If

Beep
End Sub

Private Sub Form_Load()
        'NEED BLUETOOTH CONFIG


        '------------------------------------------------------------
        'APPEAR ERROR WITH COTROLS AFTER UPDATER TO SOLVE .OCX DRIVER
        'GONE WRONG BY REMOVE REISTERY KEY
        'AS MICROSOFT TALK
        '------------------------------------------------------------
        

        On Error GoTo 0









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


        sWindowsFolder = FS.GetSpecialfolder(0)

        'sMyDocsFolder = GetSpecialfolder(5)



        Dim R As Long
        On Error Resume Next
        If IsIDE = True Then
'                For R = 0 To 120
        '        Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
'                Next
                'End
        End If



'Exit Sub



'Me.Show
'DoEvents


For Each Control In Controls
        
        On Error Resume Next
        Control.Caption = Replace(Control.Caption, "#", "")
        Control.Caption = Trim(Control.Caption)
        On Error GoTo 0


Next



For Each Control In Controls
        
    X = Control.Name = "Timer_ART_TRIP_WIRE"
    If X = False Then
        On Error Resume Next
        If Control.Caption = ".." Then Control.Visible = False
        Control.BackColor = &HFF0000
        On Error GoTo 0

    End If
Next


X = 1
Y = 1
On Error Resume Next
For Each Control In Controls

    If Mid(Control.Caption, 1, 1) = "#" Then
        Control.BackColor = &H400000
    End If
    
    If Control.enabled = True And Control.Visible = True Then
        If Control.Width + Control.Left > X Then X = Control.Width + Control.Left
        If Control.Height + Control.Top > Y Then Y = Control.Height + Control.Top
        If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
    End If
Next
On Error GoTo 0

Me.Width = (X + 80)
If mnu = 1 Then
    pluso = 740 ': pluso = 1100 'Sometimes Different
Else
    pluso = 450
End If

Me.Height = (Y + pluso)
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

'SET THE DEADLINKS
'Timer1.enabled = True

Call MNU_ADMINISTRATOR_Click

Call MNU_VERSION_Click

End Sub


Sub SubCode()

'Me.Show
DoEvents
X = 0
L = 0


For Each Control In Controls
    If InStr(UCase((Control.Name)), "LAB") > 0 Then
        Control.AutoSize = True
    End If
Next

For Each Control In Controls
    If InStr(UCase((Control.Name)), "LAB") > 0 Then
        If InStr((Control.Caption), "..") > 0 Then
            Control.Visible = False
        End If
    End If
Next



For Each Control In Controls
    If InStr(UCase((Control.Name)), "LAB") > 0 Then
    If InStr((Control.Caption), "..") = 0 Then
        List1.AddItem Control.Caption
    End If
    End If
Next


Dim Control_index(2000)
For R = 0 To List1.ListCount - 1
    Index = 0
    For Each Control In Controls
        Index = Index + 1
        If InStr(UCase((Control.Name)), "LAB") > 0 Then
            If Trim(Control.Caption) <> "" And Control.Caption = List1.List(R) And Control_index(Index) = "" Then
                Debug.Print Control.Caption
                If InStr((Control.Caption), "..") = 0 Then
                    Control_index(Index) = "Got_1"
                    If Control.Width > w1 Then w1 = Control.Width
                    If X > Me.Height - 1100 Then X = 0: L = L + w1 + 50: w1 = 0
                    
'                    Control.BackColor = &HFF0000
                    Control.Visible = True
                    Control.Top = X
                    Control.Left = L
                    
        '            If Control.Caption = "Norton 360" Then Stop
                    
                    X = X + Control.Height + 15
                    
                    'If x > Me.Height Then Exit Sub
                    
                End If
            End If
        End If
    Next
Next

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

        Dim oRunas

        Set oRunas = CreateObject("runas.clsrunas", GetComputerName)
        'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
        With oRunas
                .sDomain = "WORKGROUP"
        '    .sUserName = "Username you want to Run the Program  "
                .sUserName = "MATT 01"
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
        Shell "C:\Program Files\Autoruns Microsoft\AutoRuns.exe", vbMaximizedFocus
        End
End Sub

Private Sub Label24_Click()
        Shell sWindowsFolder + "\system32\sndvol32.exe", vbNormalFocus
        End

End Sub

Private Sub Label25_Click()
        Shell "C:\Program Files\Tweak-XP Pro\Tweak-xp.exe", vbNormalFocus
        End

End Sub

Private Sub Label26_Click()
        Shell "C:\Program Files\Norton Ghost\Console\VProConsole_.exe", vbNormalFocus
        End
End Sub

Private Sub Label27_Click()
        Shell "C:\Program Files\ProcessExplorer\Procexp.exe", vbNormalFocus
        End
End Sub

Private Sub Label28_Click()
        Shell "C:\Program Files\ViceVersa Pro 2\ViceVersa.exe", vbNormalFocus
        End
End Sub

Private Sub Label29_Click()
        Shell "NOTEPAD """ + sMyDocsFolder + "\CD-ROM\01 Total List.txt""", vbMaximizedFocus
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
                .sUserName = "Matt"
                .sPassword = "matto28"
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

        If FS.FileExists("") = True Then
        Shell "C:\Program Files\# NO INSTALL REQUIRED\GoogleImageDownloader-v1.1-win\GoogleImgDownloader.exe", vbNormalFocus
        End
        End If

        Shell "C:\Program Files\GoogleImageDownloader-v1.1-win\GoogleImgDownloader.exe", vbMaximizedFocus
        End

End Sub

Private Sub Label35_Click()
        Shell "NOTEPad ""D:\VB6\VB-NT\00_Best_VB_01\WinAmp Logger\00 Music Info Logger.txt""", vbNormalNoFocus
        End
End Sub

Private Sub Label36_Click()
        Shell "C:\Program Files\Outlook Express\wab.exe", vbMaximizedFocus
        End

End Sub

Private Sub Label37_Click()

        Shell sWindowsFolder + "\Notepad.exe", vbNormalFocus
        'Shell sWindowsFolder + "\system32\NotePad\Notepad.exe", vbNormalFocus
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

Private Sub Label40_Click()
    Shell "C:\Program Files\Easy CD-DA Extractor 5.1\ezcddaxfc.exe", vbMaximizedFocus
    End
End Sub

Private Sub Label41_Click()
        Shell "C:\Program Files\MIKSOFT\Mobile Media Converter\mmc.exe", vbNormalFocus
        End
End Sub

Private Sub Label42_Click()
        Shell "C:\Program Files\Norton 360\Engine\3.0.0.135\uiStub.exe ""/page {5ABC34AE-1037-4f5d-BF93-B2B74C80B5F7}""", vbNormalFocus
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
        Shell "NOTEPad ""D:\VB6\VB-NT\#1 Documents\Programming Snippets.txt""", vbMaximizedFocus
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

Private Sub Label58_Click()
        Shell "C:\Program Files\Windows Media Player\wmplayer.exe", vbNormalFocus
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

Private Sub Label61_Click()

    Shell "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Dir Create on another drive\Create Dir Another Drive.exe", vbNormalFocus
        End
End Sub

Private Sub Label62_Click()
    
    Shell "NOTEPAD ""D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\MATT-555ROIDS\00_Text_Data_02\VOICE SPEAK LOGG.TXT""", vbMaximizedFocus
        End
End Sub

Private Sub Label63_Click()
        Shell "C:\Program Files\EasyMPEG Lite\EasyMPEG Lite\Lite.exe", vbNormalFocus
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

Private Sub Label77_Click()
        Shell "C:\Program Files\ProcessTamer\ProcessTamerConfigurator.exe", vbNormalFocus
        End
        End Sub

        Private Sub Label8_Click()
        ChDrive App.Path
        ChDir App.Path
        Shell "D:\VB6\VB-NT\00_Best_VB_01\WinAmp Logger\WinAmp Logger.exe", vbNormalNoFocus
        End
End Sub

Private Sub Label81_Click()

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

        Private Sub Label85_Click()
        Shell "D:\VB6\VB-NT\00_Best_VB_01\Shell Winamp Loader\Winamp Loader.exe", vbNormalFocus
        End
End Sub

Private Sub Label86_Click()
        ChDir "C:\Program Files\WordNet\2.1\bin"
        Shell "C:\Program Files\WordNet\2.1\bin\wnb.exe", vbNormalFocus
        End
        End Sub

        Private Sub Label89_Click()

        Shell "C:\Program Files\00 WINAMP'S\WINAMP XXX\winamp.exe", vbNormalFocus
        End


End Sub

'Shell "D:\VB6\VB-NT\00 MP3 Sound\DSS_Olympus\DSS_Olympus.exe"
'End


Private Sub Label9_Click()
        Shell "D:\VB6\VB-NT\00_Best_VB_01\50LastEXE's\50LastEXE.exe", vbNormalFocus
        End
        End Sub

        Private Sub Label90_Click()

        Shell "C:\Program Files\00 WINAMP'S\WINAMP MY MUSIC 01\winamp.exe", vbNormalFocus

End

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
    '----------------------
    'TEMP OUT WON'T COMPILE
    '----------------------
    'SystemParametersInfo SPI_GETSCREENSAVEACTIVE, 0, Result, 0
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
    '----------------------
    'TEMP OUT WON'T COMPILE
    '----------------------
'    SetScreenSaverState = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, enabled, _
        ByVal 0&, fuWinIni) <> 0
End Function






Private Sub MNU_VB_Click()

Me.Hide

'VB_EXE=

Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE", vbNormalFocus

End
End Sub


Private Sub MNU_RESTART_Click()
Shell "SHUTDOWN -R -T 0", vbHide
End Sub

Private Sub MNU_RESTART_FORCE_Click()
Shell "RESTART -R -F -T 0", vbHide
End Sub

Private Sub MNU_RUN_NOW_Click()
XX = 0
End Sub

Private Sub MNU_SHUTDOWN_Click()
Shell "SHUTDOWN -S -T 0", vbHide

End Sub

Private Sub Form_Resize()

If WIDTH_SCREEN_SIZE_NOT_SET_YET_MAXIMIZED = False Then Exit Sub
Call SET_SCREEN_SIZE

End Sub

'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    Dim FileSpec, FILE_SPEC_GO
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

Private Sub MNU_EXIT_Click()
End
End Sub

Private Sub MNU_RUNAS_DOWSE_Click()
'    Call Label14_Click

Beep
Set oRunas = CreateObject("runas.clsrunas", GetComputerName)
With oRunas
        .sDomain = "WORKGROUP"
'        .sUserName = GetUserName
        .sUserName = "MATT 02"
        .sPassword = " "
        .sCommand = "C:\Program Files\Siber Systems\GoodSync\GoodSync-v10.exe"
        .RunAs '____ Call the Run As method
End With
Set oRunas = Nothing
End

End Sub

Private Sub Mnu_VB_ME_Click()
VBPATH = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
If Dir(VBPATH) = "" Then
    VBPATH = "C:\Program Files (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
End If

Shell VBPATH + " """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
End
End Sub
Private Sub MNU_VB_FOLDER_Click()
    Beep
    Dim EXIT_TRUE
    
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + App.EXEName + """", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
End Sub
Private Sub MNU_ADMINISTRATOR_Click()
    'Call MNU_ADMINISTRATOR_Click
    Beep
    MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR MODE ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "MODE ON", "MODE OFF")
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



Function RUN_EXE(First_Run_Done, A3_Indexer, ITEST, i_32_64_BIT, i2____, i3_Exe, VI)

'---------------------------------------------
'CHECK A PROGRAM EXIST AND-ER 64 BIT METHOD-ER
'---------------------------------------------
i_TEST = ""
i_32_64_BIT = "32 BIT"
'FLIP FLOP
i4_Exe = i3_Exe
If Dir(i4_Exe) = "" Then
    i4_Exe = Replace(LCase(i3_Exe), "c:\program files\", "C:\Program Files (x86)\")
End If
If InStr(LCase(i4_Exe), LCase("c:\program files\")) > 0 Then
    i4_Exe = Replace(LCase(i3_Exe), "c:\program files\", "C:\Program Files (x86)\")
Else
    If InStr(LCase(i4_Exe), "C:\Program Files (x86)\") > 0 Then
        i4_Exe = Replace(LCase(i3_Exe), "c:\program files (x86)\", "C:\Program Files\")
    End If
End If


If Dir(i4_Exe) <> "" Then
    i3_Exe = i4_Exe
End If

If Dir(i4_Exe) = "" Then
    i_TEST = "NOT EXIST EXE"
Else
    i_32_64_BIT = "32 BIT"
End If
If Dir(i3_Exe) = "" Then
    i3_Exe = Replace(LCase(i3_Exe), "c:\program files\", "C:\Program Files (x86)\")
End If



If i_TEST = "NOT EXIST EXE" Then
    Label_800.Item(Index).BackColor = RGB(255, 0, 0)
End If

'------------------------
'STRIP A TEST BIT NOT WORKING
'------------------------
If i3 = "--" Then i3 = ""
'------------------------

'---
'RUN
'---

If First_Run_Done = True Then
    If i3_Exe <> "" And First_Run_Done = True Then
        If i_TEST = "" Then Me.Hide
        Beep
        PROCESS_ID_NUMBER = Shell(i3_Exe, VI)
        Beep
        RUN_EXE = PROCESS_ID_NUMBER
        
        If PROCESS_ID_NUMBER = 0 Then
            Beep
            Label_800.Item(Index).BackColor = RGB(0, 255, 255)
        End If
    End If
End If
    
    
End Function

Sub Label_800_TEST(A1_INDEX, A2_INDEX)


'STUCK HERE WORKING ON
'---------------------
A3_Indexer = Val(A1_INDEX)

I1 = Trim(Label_800.Item(Val(A3_Index)).Caption)

If First_Run_Done = True Then
    Debug.Print vbCrLf + "----" + vbCrLf
    Debug.Print I1
    
    If IsIDE = True Then
        A = Now + TimeSerial(0, 10, 0) 'TEST FOR A WHILE
    End If
    If A > Now And IsIDE = True Then
        Clipboard.Clear
        Sleep 100
        Clipboard.SetText I1
    End If
End If


VI_1 = vbNormalFocus
VI_2 = vbMaximizedFocus

i3_Exe = ""


'------------------------------------------------------------------------
'ALL THE EXE THAT ARE PROGRAMMED IN HARDCODED ARE CHECKED FOR AVAIL EXIST
'------------------------------------------------------------------------
Dim RESULT_EXIST
RESULT_EXIST = False
'----------
'--------------------------------------
i2____ = "AutoRuns"
i3_Exe = "C:\Program Files\Autoruns Microsoft\AutoRuns.exe"
VI = VI_2
I_RESULT = RUN_EXE(First_Run_Done, A3_Indexer, ITEST, i_32_64_BIT, i2____, i3_Exe, VI)
If I_RESULT > 0 Then RESULT_EXIST = True
'--------------------------------------
i2____ = "LOGITECH UNIFYER"
i3_Exe = "C:\Program Files\Common Files\Logishrd\Unifying\DJCUHost.exe"
VI = VI_1
I_RESULT = RUN_EXE(First_Run_Done, A3_Indexer, ITEST, i_32_64_BIT, i2____, i3_Exe, VI)
If I_RESULT > 0 Then RESULT_EXIST = True
'--------------------------------------
i2____ = "Grant NTFS Access"
i3_Exe = "C:\HBCD\Programs\Programs\NTFSAccess2.2.exe"
VI = VI_1
I_RESULT = RUN_EXE(First_Run_Done, A3_Indexer, ITEST, i_32_64_BIT, i2____, i3_Exe, VI)

If I_RESULT > 0 Then RESULT_EXIST = True
'--------------------------------------
i2____ = "WinDowse 02 Shell"
i3_Exe = "C:\Program Files\Greatis\WinDowse\WinDowse.exe"
VI = VI_1
I_RESULT = RUN_EXE(First_Run_Done, A3_Indexer, ITEST, i_32_64_BIT, i2____, i3_Exe, VI)
If I_RESULT > 0 Then RESULT_EXIST = True
'--------------------------------------
'--------------------------------------
'--------------------------------------
'--------------------------------------
i2____ = "EXPLORER MAXIMIZED": i3_Exe = "C:\Windows\explorer.exe": VI = VI_1
I_RESULT = RUN_EXE(First_Run_Done, A3_Indexer, ITEST, i_32_64_BIT, i2____, i3_Exe, VI)
If I_RESULT > 0 Then RESULT_EXIST = True
'--------------------------------------
'--------------------------------------
i2____ = "--": i3_Exe = "--": VI = VI_1
I_RESULT = RUN_EXE(First_Run_Done, A3_Indexer, ITEST, i_32_64_BIT, i2____, i3_Exe, VI)
If I_RESULT > 0 Then RESULT_EXIST = True
'--------------------------------------
'EXPLORER MAXIMIZED



If First_Run_Done = False Then
    If RESULT_EXIST = False Then
        Label_800.Item(Val(A3_Indexer)).BackColor = RGB(255, 0, 0)
        Label_800.Item(Val(A3_Indexer)).Refresh
    End If
End If

X_COL_1(A3_Index) = Label_800.Item(Val(A3_Indexer)).BackColor
X_COL_2(A3_Index) = Label_800.Item(Val(A3_Indexer)).ForeColor

Label_800.Item(Val(A3_Indexer)).BackColor = RGB(255, 255, 255) - 20
Label_800.Item(Val(A3_Indexer)).ForeColor = RGB(0, 0, 0) + 200

If First_Run_Done = True Then End


Exit Sub






If 1 = 2 And i = "WinDowse 02 Shell" Then
        LOAD_EXE = "C:\Program Files\Greatis\WinDowse\WinDowse.exe"
        If Dir(LOAD_EXE) = "" Then
            LOAD_EXE = "C:\Program Files (X86)\Greatis\WinDowse\WinDowse.exe"
        End If
        Shell LOAD_EXE, vbNormalFocus  ', vbMaximizedFocus
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
End If




MsgBox "PROGRAM FOR " + vbCrLf + i + vbCrLf + "NOT WORKING YET", vbMsgBoxSetForeground
'Exit Function



'Select Case number
    '----
    'VBA USE THE CASE - Google Search
    'https://www.google.co.uk/search?num=50&q=VBA+USE+THE+CASE&oq=VBA+USE+THE+CASE&gs_l=serp.3..0i8i30k1.18801.22385.0.22979.6.6.0.0.0.0.420.737.3-1j1.2.0....0...1c.1.64.serp..4.2.736...35i39k1j0i22i30k1.54j1L4U_61s
    '--------
    'Select...Case Statement (Visual Basic)
    'https://msdn.microsoft.com/en-us/library/cy37t14y.aspx
    '----
    '------------------------------------------------------------------
    'Case 1 To 5
    '    Debug.WriteLine("Between 1 and 5, inclusive")
        ' The following is the only Case clause that evaluates to True.
    'Case 6, 7, 8
    '    Debug.WriteLine("Between 6 and 8, inclusive")
    'Case 9 To 10
    '    Debug.WriteLine("Equal to 9 or 10")
    'Case Else
    '    Debug.WriteLine("Not between 1 and 10, inclusive")
    '------------------------------------------------------------------
'End Select

If Label_800.Item(Index).Caption = "AutoRuns" Then
End If





End Sub

Private Sub Label_800_Click(Index As Integer)

    Control_index = Index
    'A1 = Control.Name

    Call Label_800_TEST("", Label_800.Item(Index)) 'Label_800.Name

End Sub


Sub SET_SCREEN_SIZE()
        X = 1
        Y = 1
        On Error Resume Next
        For Each Control In Controls
                If Control.enabled = True And Control.Visible = True Then
                        If Control.Width + Control.Left > X Then X = Control.Width + Control.Left
                        If Control.Height + Control.Top > Y Then Y = Control.Height + Control.Top
                        If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
                End If
        Next
        'On Error GoTo 0

        Err.Clear
        Me.Width = (X + 180)
        If Err.Number > 0 Then
            WIDTH_SCREEN_SIZE_NOT_SET_YET_MAXIMIZED = True
        Else
            WIDTH_SCREEN_SIZE_NOT_SET_YET_MAXIMIZED = False
        End If
        
        If mnu = 1 Then
                pluso = 740 ': pluso = 1100 'Sometimes Different
        Else
                pluso = 450
        End If

        pluso = 900
        Me.Height = (Y + pluso)


        Me.Refresh
        DoEvents

        Me.Left = (Screen.Width - Me.Width) / 2
        Me.Top = 100 '((Screen.Height - Me.Height) / 2) - 800


End Sub


Private Sub MNU_VERSION_Click()
MNU_VERSION.Caption = "Ver. " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
End Sub

Private Sub Timer1_Timer()
        
    If SET_WIN_STATE = 0 Then

        Me.Show
        DoEvents
        Me.WindowState = vbMaximized
        DoEvents
        Me.Refresh
        DoEvents
        
        SET_WIN_STATE = 2
        Exit Sub
            
    End If
    
    If SET_WIN_STATE = 2 Then
        
        Call SET_SCREEN_SIZE
        SET_WIN_STATE = 3
        
        Exit Sub
            
    End If
    
    If SET_WIN_STATE = 3 Then
        
        On Error Resume Next
        For Each Control In Label_800
            If Control.Caption = ".." Then Control.Visible = False
        Next
            
            
        First_Run_Done = False
        Timer1.enabled = False
        'Label_800.Item.Index
        On Error Resume Next
        
        For RCOUNT = 0 To Label_800.Count
            i = Label_800.Item(RCOUNT).Caption
            'WORK HERE BREAK POINT SET WITH STOP
            'If IsIDE = True And i = "OPUS1" Then Stop
            Call Label_800_TEST(RCOUNT, i)
        Next
        
        First_Run_Done = True
        
        SET_WIN_STATE = 4
        Exit Sub
            
    End If

    If SET_WIN_STATE = 4 Then

        For R = 0 To UBound(X_COL_1)
            X_COL_1(R) = Label_800.Item(Val(R)).BackColor
            X_COL_2(R) = Label_800.Item(Val(R)).ForeColor

            Label_800.Item(X_COL_1(R)).BackColor = RGB(255, 255, 255) - 20
            Label_800.Item(X_COL_2(R)).ForeColor = RGB(0, 0, 0) + 200
        Next

        SET_WIN_STATE = 5
        Timer1.enabled = False
        Exit Sub
            
    End If



End Sub


'-------------------------------------------------
'-------------------------------------------------
Private Sub Timer_1_MINUTE_Timer()

Timer_1_MINUTE.enabled = False

Exit Sub

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
If Dir(PATH_FILE_NAME2) <> "" Then
    Set F = FS.GetFile(PATH_FILE_NAME2)
    VB_DATE = F.DateLastModified
End If

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



