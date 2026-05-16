VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   Caption         =   "Matt Run AS"
   ClientHeight    =   10650
   ClientLeft      =   60
   ClientTop       =   -1665
   ClientWidth     =   15240
   FillStyle       =   0  'Solid
   Icon            =   "Matt_RunAS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   Begin VB.Timer Timer_ART_TRIP_WIRE 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2040
      Top             =   5280
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   100
      Top             =   4920
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label34 
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
      Height          =   285
      Left            =   4200
      TabIndex        =   121
      Top             =   5640
      Width           =   2130
   End
   Begin VB.Label Label30 
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
      Height          =   285
      Left            =   1200
      TabIndex        =   120
      Top             =   10080
      Width           =   2040
   End
   Begin VB.Label Label122 
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
      Height          =   285
      Left            =   4080
      TabIndex        =   119
      Top             =   2040
      Width           =   2460
   End
   Begin VB.Label Label121 
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
      Left            =   840
      TabIndex        =   118
      Top             =   120
      Width           =   150
   End
   Begin VB.Label Label120 
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
      Left            =   600
      TabIndex        =   117
      Top             =   120
      Width           =   150
   End
   Begin VB.Label LABEL_VB 
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
      Height          =   285
      Left            =   1200
      TabIndex        =   116
      Top             =   9720
      Width           =   900
   End
   Begin VB.Label Label62 
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
      Left            =   360
      TabIndex        =   115
      Top             =   120
      Width           =   150
   End
   Begin VB.Label Label48 
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
      Height          =   285
      Left            =   4320
      TabIndex        =   114
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label46 
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
      Left            =   240
      TabIndex        =   113
      Top             =   120
      Width           =   150
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
      Left            =   1320
      TabIndex        =   112
      Top             =   120
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
      Left            =   8040
      TabIndex        =   111
      Top             =   480
      Width           =   6705
   End
   Begin VB.Label Label118 
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
      Height          =   285
      Left            =   4140
      TabIndex        =   110
      Top             =   5160
      Width           =   2925
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
      Left            =   2880
      TabIndex        =   109
      Top             =   5160
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
      TabIndex        =   108
      Top             =   1200
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
      TabIndex        =   107
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
      Left            =   11640
      TabIndex        =   106
      Top             =   6120
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
      Height          =   285
      Left            =   8880
      TabIndex        =   105
      Top             =   3480
      Width           =   3225
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
      Left            =   3000
      TabIndex        =   104
      Top             =   6120
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
      Left            =   12240
      TabIndex        =   103
      Top             =   3480
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
      Left            =   12600
      TabIndex        =   102
      Top             =   5640
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
      Left            =   6720
      TabIndex        =   101
      Top             =   7560
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
      Left            =   7920
      TabIndex        =   99
      Top             =   840
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
      Left            =   8880
      TabIndex        =   97
      Top             =   4800
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
      Height          =   285
      Left            =   360
      TabIndex        =   96
      Top             =   120
      Width           =   150
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
      Left            =   8640
      TabIndex        =   95
      Top             =   5130
      Width           =   1980
   End
   Begin VB.Label Label103 
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
      Height          =   285
      Left            =   11880
      TabIndex        =   94
      Top             =   4200
      Width           =   3030
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
      Left            =   8400
      TabIndex        =   93
      Top             =   2040
      Width           =   1950
   End
   Begin VB.Label Label101 
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
      Height          =   285
      Left            =   11160
      TabIndex        =   92
      Top             =   3840
      Width           =   4545
   End
   Begin VB.Label Label100 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "GoogleImage Downloader WIG And Orginal"
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
      TabIndex        =   91
      Top             =   9120
      Width           =   4905
   End
   Begin VB.Label Label99 
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
      Height          =   285
      Left            =   12000
      TabIndex        =   90
      Top             =   4560
      Width           =   2850
   End
   Begin VB.Label Label98 
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
      Height          =   285
      Left            =   11160
      TabIndex        =   89
      Top             =   120
      Width           =   3030
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
      Left            =   9240
      TabIndex        =   88
      Top             =   3120
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
      Height          =   285
      Left            =   13080
      TabIndex        =   87
      Top             =   1680
      Width           =   2070
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
      Left            =   3960
      TabIndex        =   86
      Top             =   840
      Width           =   2760
   End
   Begin VB.Label Label94 
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
      Height          =   285
      Left            =   7800
      TabIndex        =   85
      Top             =   2400
      Width           =   1245
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
      Height          =   285
      Left            =   4080
      TabIndex        =   84
      Top             =   1560
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
      Left            =   11760
      TabIndex        =   83
      Top             =   2760
      Width           =   2565
   End
   Begin VB.Label Label90 
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
      Height          =   285
      Left            =   7080
      TabIndex        =   82
      Top             =   9840
      Width           =   3480
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
      TabIndex        =   81
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
      Left            =   11760
      TabIndex        =   80
      Top             =   2400
      Width           =   2310
   End
   Begin VB.Label Label87 
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
      Height          =   285
      Left            =   4440
      TabIndex        =   79
      Top             =   3120
      Width           =   2235
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
      TabIndex        =   78
      Top             =   6360
      Width           =   2265
   End
   Begin VB.Label Label85 
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
      Height          =   285
      Left            =   12120
      TabIndex        =   77
      Top             =   4920
      Width           =   1815
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
      TabIndex        =   76
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
      Left            =   3000
      TabIndex        =   74
      Top             =   5520
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
      TabIndex        =   73
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
      Left            =   3960
      TabIndex        =   72
      Top             =   3480
      Width           =   1710
   End
   Begin VB.Label Label79 
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
      Height          =   285
      Left            =   10920
      TabIndex        =   71
      Top             =   7560
      Width           =   4020
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
      Left            =   3840
      TabIndex        =   70
      Top             =   2400
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
      Left            =   120
      TabIndex        =   69
      Top             =   6000
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
      TabIndex        =   68
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
      Left            =   10560
      TabIndex        =   67
      Top             =   2040
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
      Height          =   285
      Left            =   2400
      TabIndex        =   66
      Top             =   3000
      Width           =   1245
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
      Height          =   285
      Left            =   12360
      TabIndex        =   65
      Top             =   6960
      Width           =   2370
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
      Left            =   12240
      TabIndex        =   64
      Top             =   6480
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
      TabIndex        =   63
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
      Left            =   3480
      TabIndex        =   62
      Top             =   7200
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
      Left            =   9120
      TabIndex        =   61
      Top             =   3840
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
      Left            =   11040
      TabIndex        =   60
      Top             =   8280
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
      Left            =   9480
      TabIndex        =   59
      Top             =   4200
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
      Height          =   285
      Left            =   2040
      TabIndex        =   58
      Top             =   2640
      Width           =   1410
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
      TabIndex        =   57
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
      Left            =   3240
      TabIndex        =   56
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label58 
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
      Height          =   285
      Left            =   12360
      TabIndex        =   48
      Top             =   5280
      Width           =   2805
   End
   Begin VB.Label Label57 
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
      Left            =   9240
      TabIndex        =   47
      Top             =   2760
      Width           =   2325
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
      TabIndex        =   46
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
      Left            =   3360
      TabIndex        =   43
      Top             =   6840
      Width           =   1650
   End
   Begin VB.Label LABEL51 
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
      Height          =   285
      Left            =   5160
      TabIndex        =   42
      Top             =   1680
      Width           =   3330
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
      Left            =   5520
      TabIndex        =   40
      Top             =   4200
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
      Height          =   285
      Left            =   3840
      TabIndex        =   39
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "INFRAN VIEW && FastStone Image Viewer"
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
      Left            =   7080
      TabIndex        =   37
      Top             =   9480
      Width           =   4650
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
      Left            =   5040
      TabIndex        =   33
      Top             =   4920
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
      Left            =   3240
      TabIndex        =   31
      Top             =   4560
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
      TabIndex        =   30
      Top             =   4065
      Width           =   2610
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
      Height          =   285
      Left            =   2760
      TabIndex        =   9
      Top             =   1875
      Width           =   930
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
      TabIndex        =   98
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
      Left            =   2760
      TabIndex        =   75
      Top             =   120
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
      TabIndex        =   55
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
      TabIndex        =   54
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
      TabIndex        =   53
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
      Left            =   8400
      TabIndex        =   52
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
      Height          =   285
      Left            =   7935
      TabIndex        =   51
      Top             =   1230
      Width           =   1950
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
      Left            =   9120
      TabIndex        =   50
      Top             =   2400
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
      Height          =   285
      Left            =   8100
      TabIndex        =   49
      Top             =   1560
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
      TabIndex        =   45
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
      TabIndex        =   44
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
      TabIndex        =   41
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
      TabIndex        =   38
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
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   34
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
      Left            =   120
      TabIndex        =   32
      Top             =   4920
      Width           =   1185
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
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   27
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
      Height          =   285
      Left            =   6720
      TabIndex        =   24
      Top             =   4290
      Width           =   2745
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
      Height          =   285
      Left            =   9480
      TabIndex        =   20
      Top             =   4560
      Width           =   1650
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
      Top             =   720
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
      Height          =   285
      Left            =   6735
      TabIndex        =   14
      Top             =   360
      Width           =   1050
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
      Left            =   4680
      TabIndex        =   3
      Top             =   480
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
      Left            =   13080
      TabIndex        =   1
      Top             =   3120
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
      Left            =   5640
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Menu Mnu_VB 
      Caption         =   "VB ME"
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 
' = LastModifiedToCurrent("K:\TEMP\ART_PROG_TRIP_WIRE.txt")

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

Set F = FS.getfile("K:\TEMP\ART_PROG_TRIP_WIRE.txt")
i = F.datelastmodified
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






Private Sub VB_Click()
Me.Hide

Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE", vbNormalFocus

End
End Sub

