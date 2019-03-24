VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "Matt Run AS"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   -1920
   ClientWidth     =   11880
   Icon            =   "Matt_RunAS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   11880
   Begin VB.Label Label91 
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
      Left            =   5670
      TabIndex        =   90
      Top             =   990
      Width           =   120
   End
   Begin VB.Label Label90 
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
      Height          =   285
      Left            =   8100
      TabIndex        =   89
      Top             =   450
      Width           =   2160
   End
   Begin VB.Label Label89 
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
      Height          =   285
      Left            =   8100
      TabIndex        =   88
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label88 
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
      Left            =   5685
      TabIndex        =   87
      Top             =   945
      Width           =   120
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
      Height          =   285
      Left            =   5655
      TabIndex        =   86
      Top             =   990
      Width           =   120
   End
   Begin VB.Label Label86 
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
      TabIndex        =   85
      Top             =   45
      Width           =   1470
   End
   Begin VB.Label Label85 
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
      Left            =   8115
      TabIndex        =   84
      Top             =   60
      Width           =   2070
   End
   Begin VB.Label Label84 
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
      Left            =   5670
      TabIndex        =   83
      Top             =   945
      Width           =   120
   End
   Begin VB.Label Label83 
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
      Left            =   5535
      TabIndex        =   82
      Top             =   825
      Width           =   120
   End
   Begin VB.Label Label82 
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
      Left            =   5385
      TabIndex        =   81
      Top             =   795
      Width           =   120
   End
   Begin VB.Label Label81 
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
      Height          =   285
      Left            =   60
      TabIndex        =   80
      Top             =   75
      Width           =   1005
   End
   Begin VB.Label Label80 
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
      Left            =   5550
      TabIndex        =   79
      Top             =   780
      Width           =   120
   End
   Begin VB.Label Label79 
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
      Left            =   5520
      TabIndex        =   78
      Top             =   825
      Width           =   120
   End
   Begin VB.Label Label78 
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
      Left            =   5370
      TabIndex        =   77
      Top             =   795
      Width           =   120
   End
   Begin VB.Label Label77 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "PROCESS TAMER"
      BeginProperty Font 
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
      Left            =   8250
      TabIndex        =   76
      Top             =   4425
      Width           =   2115
   End
   Begin VB.Label Label76 
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
      Left            =   5535
      TabIndex        =   75
      Top             =   780
      Width           =   120
   End
   Begin VB.Label Label75 
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
      Left            =   5430
      TabIndex        =   74
      Top             =   1005
      Width           =   120
   End
   Begin VB.Label Label74 
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
      Left            =   5280
      TabIndex        =   73
      Top             =   975
      Width           =   120
   End
   Begin VB.Label Label73 
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
      Left            =   5205
      TabIndex        =   72
      Top             =   1125
      Width           =   120
   End
   Begin VB.Label Label72 
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
      Left            =   5445
      TabIndex        =   71
      Top             =   960
      Width           =   120
   End
   Begin VB.Label Label71 
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
      Left            =   5415
      TabIndex        =   70
      Top             =   1005
      Width           =   120
   End
   Begin VB.Label Label70 
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
      Left            =   5265
      TabIndex        =   69
      Top             =   975
      Width           =   120
   End
   Begin VB.Label Label69 
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
      Left            =   5190
      TabIndex        =   68
      Top             =   1125
      Width           =   120
   End
   Begin VB.Label Label68 
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
      Left            =   5430
      TabIndex        =   67
      Top             =   960
      Width           =   120
   End
   Begin VB.Label Label67 
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
      Left            =   5295
      TabIndex        =   66
      Top             =   840
      Width           =   120
   End
   Begin VB.Label Label66 
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
      Left            =   5145
      TabIndex        =   65
      Top             =   810
      Width           =   120
   End
   Begin VB.Label Label65 
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
      Left            =   5070
      TabIndex        =   64
      Top             =   960
      Width           =   120
   End
   Begin VB.Label Label64 
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
      Left            =   5310
      TabIndex        =   63
      Top             =   795
      Width           =   120
   End
   Begin VB.Label Label63 
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
      Height          =   285
      Left            =   2790
      TabIndex        =   62
      Top             =   3885
      Width           =   1410
   End
   Begin VB.Label Label62 
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
      Height          =   285
      Left            =   8130
      TabIndex        =   61
      Top             =   2010
      Width           =   2025
   End
   Begin VB.Label Label61 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "DIR ON ANOTHER DRIVE"
      BeginProperty Font 
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
      Left            =   3060
      TabIndex        =   60
      Top             =   6150
      Width           =   2910
   End
   Begin VB.Label Label60 
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
      Left            =   5295
      TabIndex        =   59
      Top             =   795
      Width           =   120
   End
   Begin VB.Label Label59 
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
      Left            =   5055
      TabIndex        =   58
      Top             =   960
      Width           =   120
   End
   Begin VB.Label Label58 
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
      Height          =   285
      Left            =   2340
      TabIndex        =   57
      Top             =   3240
      Width           =   3075
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
      Height          =   285
      Left            =   5130
      TabIndex        =   56
      Top             =   810
      Width           =   120
   End
   Begin VB.Label Label56 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   330
      Left            =   45
      TabIndex        =   55
      Top             =   6495
      Width           =   405
   End
   Begin VB.Label Label55 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   330
      Left            =   45
      TabIndex        =   54
      Top             =   6135
      Width           =   405
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
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
      TabIndex        =   53
      Top             =   3615
      Width           =   2640
   End
   Begin VB.Label Label54 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "TAGG - GEN"
      BeginProperty Font 
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
      Left            =   480
      TabIndex        =   52
      Top             =   5760
      Width           =   1395
   End
   Begin VB.Label Label53 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "PLAYLIST - GEN"
      BeginProperty Font 
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
      Left            =   480
      TabIndex        =   51
      Top             =   6495
      Width           =   1875
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
      Left            =   8235
      TabIndex        =   50
      Top             =   3630
      Width           =   1650
   End
   Begin VB.Label LABEL51 
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
      Height          =   285
      Left            =   3105
      TabIndex        =   49
      Top             =   6555
      Width           =   2325
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "DUMMMY TAGG - GEN"
      BeginProperty Font 
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
      Left            =   480
      TabIndex        =   48
      Top             =   6135
      Width           =   2550
   End
   Begin VB.Label Label49 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "VB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   330
      Left            =   45
      TabIndex        =   47
      Top             =   5760
      Width           =   405
   End
   Begin VB.Label Label48 
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
      Left            =   8130
      TabIndex        =   46
      Top             =   1590
      Width           =   1665
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
      Left            =   2805
      TabIndex        =   45
      Top             =   4275
      Width           =   1080
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
      Left            =   5280
      TabIndex        =   44
      Top             =   840
      Width           =   120
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
      Left            =   2265
      TabIndex        =   43
      Top             =   2865
      Width           =   2835
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
      Left            =   6120
      TabIndex        =   42
      Top             =   6165
      Width           =   990
   End
   Begin VB.Label Label43 
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
      Height          =   285
      Left            =   1725
      TabIndex        =   41
      Top             =   5070
      Width           =   2760
   End
   Begin VB.Label Label42 
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
      Height          =   285
      Left            =   8235
      TabIndex        =   40
      Top             =   3225
      Width           =   1260
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Mobile Media Converter"
      BeginProperty Font 
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
      Left            =   6135
      TabIndex        =   39
      Top             =   5505
      Width           =   2685
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
      Left            =   8220
      TabIndex        =   38
      Top             =   4050
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
      TabIndex        =   37
      Top             =   5400
      Width           =   1335
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
      Left            =   2640
      TabIndex        =   36
      Top             =   5415
      Width           =   1560
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
      TabIndex        =   35
      Top             =   4965
      Width           =   1185
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
      Left            =   2010
      TabIndex        =   34
      Top             =   4665
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
      TabIndex        =   33
      Top             =   4065
      Width           =   2610
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
      Left            =   6120
      TabIndex        =   32
      Top             =   6510
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
      Left            =   6120
      TabIndex        =   31
      Top             =   5805
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
      Left            =   6150
      TabIndex        =   30
      Top             =   5190
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
      Left            =   6120
      TabIndex        =   29
      Top             =   4860
      Width           =   1620
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
      Left            =   3135
      TabIndex        =   28
      Top             =   5790
      Width           =   1395
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
      TabIndex        =   27
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
      Left            =   6120
      TabIndex        =   26
      Top             =   4530
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
      Left            =   6105
      TabIndex        =   25
      Top             =   4215
      Width           =   1980
   End
   Begin VB.Label Label26 
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
      Left            =   8130
      TabIndex        =   24
      Top             =   1155
      Width           =   1545
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
      Height          =   255
      Left            =   6105
      TabIndex        =   23
      Top             =   3915
      Width           =   1740
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
      Height          =   285
      Left            =   1890
      TabIndex        =   22
      Top             =   2475
      Width           =   1440
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
      Height          =   285
      Left            =   6090
      TabIndex        =   20
      Top             =   1500
      Width           =   1155
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
      Height          =   285
      Left            =   6105
      TabIndex        =   17
      Top             =   3585
      Width           =   1185
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
      Height          =   285
      Left            =   6090
      TabIndex        =   16
      Top             =   2490
      Width           =   2400
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
      Left            =   6090
      TabIndex        =   15
      Top             =   1005
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
      Left            =   6090
      TabIndex        =   14
      Top             =   510
      Width           =   1080
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
      Left            =   6090
      TabIndex        =   13
      Top             =   15
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
      Height          =   285
      Left            =   6090
      TabIndex        =   12
      Top             =   3225
      Width           =   1005
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
      Height          =   285
      Left            =   6090
      TabIndex        =   11
      Top             =   2865
      Width           =   2190
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
      Height          =   285
      Left            =   6090
      TabIndex        =   10
      Top             =   1995
      Width           =   1455
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
      Left            =   2055
      TabIndex        =   9
      Top             =   1875
      Width           =   945
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Encrypt"
      BeginProperty Font 
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
      Left            =   1035
      TabIndex        =   8
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "EXE's"
      BeginProperty Font 
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
      Top             =   1890
      Width           =   690
   End
   Begin VB.Label Label7 
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
      Height          =   285
      Left            =   2610
      TabIndex        =   6
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Cool 2"
      BeginProperty Font 
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
      Left            =   1755
      TabIndex        =   5
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label5 
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
      Height          =   285
      Left            =   60
      TabIndex        =   4
      Top             =   1080
      Width           =   1515
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Opus"
      BeginProperty Font 
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
      Left            =   2145
      TabIndex        =   3
      Top             =   60
      Width           =   630
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
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   735
      Width           =   3135
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
      Left            =   45
      TabIndex        =   1
      Top             =   405
      Width           =   1470
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "Opus"
      BeginProperty Font 
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
      Left            =   1275
      TabIndex        =   0
      Top             =   60
      Width           =   645
   End
   Begin VB.Menu Mnu_VB 
      Caption         =   "VB ME"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'"C:\Program Files\Runtime Software\DriveImage XML\dixml.exe"

'WANT BLUE TOOTH ON HERE

Private Sub Form_Load()
'NEED BLUETOOTH CONFIG

For Each Control In Controls
    If Control.Caption = ".." Then Control.Visible = False
Next










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

sMyDocsFolder = GetSpecialfolder(5)



Dim R As Long
On Error Resume Next
If ISIDE = True Then
    For R = 0 To 120
        Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
    Next
    'End
End If

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
On Error GoTo 0

Me.Width = (x + 180)
If mnu = 1 Then
    pluso = 740 ': pluso = 1100 'Sometimes Different
Else
    pluso = 450
End If

pluso = 1000
Me.Height = (y + pluso)


Me.Refresh
DoEvents

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = ((Screen.Height - Me.Height) / 2) - 200





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
Line Input #fr1, tt$
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
    .sCommand = tt$
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

If FS.FILEEXISTS("") = True Then
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


Private Sub Mnu_VB_Click()
'If IsIDE = False Then
    Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
'    End
'End If
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

