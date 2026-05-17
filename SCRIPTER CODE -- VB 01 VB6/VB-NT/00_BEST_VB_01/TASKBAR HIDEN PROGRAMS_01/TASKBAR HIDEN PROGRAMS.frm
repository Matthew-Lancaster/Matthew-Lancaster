VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000008&
   Caption         =   "HIDDEN TASKBAR PROGRAMS"
   ClientHeight    =   10584
   ClientLeft      =   60
   ClientTop       =   636
   ClientWidth     =   11976
   FillStyle       =   0  'Solid
   Icon            =   "TASKBAR HIDEN PROGRAMS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10584
   ScaleWidth      =   11976
   Begin VB.Timer Timer_ART_TRIP_WIRE 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5805
      Top             =   3375
   End
   Begin VB.ListBox List1 
      Height          =   624
      Left            =   4080
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   2970
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H008080FF&
      Caption         =   "Label2"
      Height          =   405
      Left            =   3660
      TabIndex        =   129
      Top             =   1305
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   127
      Left            =   1920
      TabIndex        =   128
      Top             =   1140
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   126
      Left            =   1830
      TabIndex        =   127
      Top             =   1185
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   125
      Left            =   1920
      TabIndex        =   126
      Top             =   1125
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   124
      Left            =   1830
      TabIndex        =   125
      Top             =   1170
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   123
      Left            =   1845
      TabIndex        =   124
      Top             =   990
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   122
      Left            =   1755
      TabIndex        =   123
      Top             =   1035
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   121
      Left            =   1845
      TabIndex        =   122
      Top             =   975
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   120
      Left            =   1755
      TabIndex        =   121
      Top             =   1020
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   119
      Left            =   1905
      TabIndex        =   120
      Top             =   975
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   118
      Left            =   1815
      TabIndex        =   119
      Top             =   1020
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   117
      Left            =   1905
      TabIndex        =   118
      Top             =   960
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   116
      Left            =   1815
      TabIndex        =   117
      Top             =   1005
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   115
      Left            =   1830
      TabIndex        =   116
      Top             =   825
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   114
      Left            =   1740
      TabIndex        =   115
      Top             =   870
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   113
      Left            =   1830
      TabIndex        =   114
      Top             =   810
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   112
      Left            =   1740
      TabIndex        =   113
      Top             =   855
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   111
      Left            =   1785
      TabIndex        =   112
      Top             =   1110
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   110
      Left            =   1695
      TabIndex        =   111
      Top             =   1155
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   109
      Left            =   1785
      TabIndex        =   110
      Top             =   1095
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   108
      Left            =   1695
      TabIndex        =   109
      Top             =   1140
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   107
      Left            =   1710
      TabIndex        =   108
      Top             =   960
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   106
      Left            =   1620
      TabIndex        =   107
      Top             =   1005
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   105
      Left            =   1710
      TabIndex        =   106
      Top             =   945
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   104
      Left            =   1620
      TabIndex        =   105
      Top             =   990
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   103
      Left            =   1770
      TabIndex        =   104
      Top             =   945
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   102
      Left            =   1680
      TabIndex        =   103
      Top             =   990
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   101
      Left            =   1770
      TabIndex        =   102
      Top             =   930
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   100
      Left            =   1680
      TabIndex        =   101
      Top             =   975
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   99
      Left            =   1695
      TabIndex        =   100
      Top             =   795
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   98
      Left            =   1605
      TabIndex        =   99
      Top             =   840
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   97
      Left            =   1695
      TabIndex        =   98
      Top             =   780
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   96
      Left            =   1605
      TabIndex        =   97
      Top             =   825
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   95
      Left            =   1845
      TabIndex        =   96
      Top             =   1020
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   94
      Left            =   1755
      TabIndex        =   95
      Top             =   1065
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   93
      Left            =   1845
      TabIndex        =   94
      Top             =   1005
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   92
      Left            =   1755
      TabIndex        =   93
      Top             =   1050
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   91
      Left            =   1770
      TabIndex        =   92
      Top             =   870
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   90
      Left            =   1680
      TabIndex        =   91
      Top             =   915
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   89
      Left            =   1770
      TabIndex        =   90
      Top             =   855
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   88
      Left            =   1680
      TabIndex        =   89
      Top             =   900
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   87
      Left            =   1830
      TabIndex        =   88
      Top             =   855
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   86
      Left            =   1740
      TabIndex        =   87
      Top             =   900
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   85
      Left            =   1830
      TabIndex        =   86
      Top             =   840
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   84
      Left            =   1740
      TabIndex        =   85
      Top             =   885
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   83
      Left            =   1755
      TabIndex        =   84
      Top             =   705
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   82
      Left            =   1665
      TabIndex        =   83
      Top             =   750
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   81
      Left            =   1755
      TabIndex        =   82
      Top             =   690
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   80
      Left            =   1665
      TabIndex        =   81
      Top             =   735
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   79
      Left            =   1710
      TabIndex        =   80
      Top             =   990
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   78
      Left            =   1620
      TabIndex        =   79
      Top             =   1035
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   77
      Left            =   1710
      TabIndex        =   78
      Top             =   975
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   76
      Left            =   1620
      TabIndex        =   77
      Top             =   1020
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   75
      Left            =   1635
      TabIndex        =   76
      Top             =   840
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   74
      Left            =   1545
      TabIndex        =   75
      Top             =   885
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   73
      Left            =   1635
      TabIndex        =   74
      Top             =   825
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   72
      Left            =   1545
      TabIndex        =   73
      Top             =   870
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   71
      Left            =   1695
      TabIndex        =   72
      Top             =   825
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   70
      Left            =   1605
      TabIndex        =   71
      Top             =   870
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   69
      Left            =   1695
      TabIndex        =   70
      Top             =   810
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   68
      Left            =   1605
      TabIndex        =   69
      Top             =   855
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   67
      Left            =   1620
      TabIndex        =   68
      Top             =   675
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   66
      Left            =   1530
      TabIndex        =   67
      Top             =   720
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   65
      Left            =   1620
      TabIndex        =   66
      Top             =   660
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   64
      Left            =   1530
      TabIndex        =   65
      Top             =   705
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   63
      Left            =   1920
      TabIndex        =   64
      Top             =   1155
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   62
      Left            =   1830
      TabIndex        =   63
      Top             =   1200
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   61
      Left            =   1920
      TabIndex        =   62
      Top             =   1140
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   60
      Left            =   1830
      TabIndex        =   61
      Top             =   1185
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   59
      Left            =   1845
      TabIndex        =   60
      Top             =   1005
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   58
      Left            =   1755
      TabIndex        =   59
      Top             =   1050
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   57
      Left            =   1845
      TabIndex        =   58
      Top             =   990
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   56
      Left            =   1755
      TabIndex        =   57
      Top             =   1035
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   55
      Left            =   1905
      TabIndex        =   56
      Top             =   990
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   54
      Left            =   1815
      TabIndex        =   55
      Top             =   1035
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   53
      Left            =   1905
      TabIndex        =   54
      Top             =   975
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   52
      Left            =   1815
      TabIndex        =   53
      Top             =   1020
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   51
      Left            =   1830
      TabIndex        =   52
      Top             =   840
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   50
      Left            =   1740
      TabIndex        =   51
      Top             =   885
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   49
      Left            =   1830
      TabIndex        =   50
      Top             =   825
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   48
      Left            =   1740
      TabIndex        =   49
      Top             =   870
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   47
      Left            =   1785
      TabIndex        =   48
      Top             =   1125
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   46
      Left            =   1695
      TabIndex        =   47
      Top             =   1170
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   45
      Left            =   1785
      TabIndex        =   46
      Top             =   1110
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   44
      Left            =   1695
      TabIndex        =   45
      Top             =   1155
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   43
      Left            =   1710
      TabIndex        =   44
      Top             =   975
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   42
      Left            =   1620
      TabIndex        =   43
      Top             =   1020
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   41
      Left            =   1710
      TabIndex        =   42
      Top             =   960
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   40
      Left            =   1620
      TabIndex        =   41
      Top             =   1005
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   39
      Left            =   1770
      TabIndex        =   40
      Top             =   960
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   38
      Left            =   1680
      TabIndex        =   39
      Top             =   1005
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   37
      Left            =   1770
      TabIndex        =   38
      Top             =   945
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   36
      Left            =   1680
      TabIndex        =   37
      Top             =   990
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   35
      Left            =   1695
      TabIndex        =   36
      Top             =   810
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   34
      Left            =   1605
      TabIndex        =   35
      Top             =   855
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   33
      Left            =   1695
      TabIndex        =   34
      Top             =   795
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   32
      Left            =   1605
      TabIndex        =   33
      Top             =   840
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   31
      Left            =   1845
      TabIndex        =   32
      Top             =   1035
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   30
      Left            =   1755
      TabIndex        =   31
      Top             =   1080
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   29
      Left            =   1845
      TabIndex        =   30
      Top             =   1020
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   28
      Left            =   1755
      TabIndex        =   29
      Top             =   1065
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   27
      Left            =   1770
      TabIndex        =   28
      Top             =   885
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   26
      Left            =   1680
      TabIndex        =   27
      Top             =   930
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   25
      Left            =   1770
      TabIndex        =   26
      Top             =   870
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   24
      Left            =   1680
      TabIndex        =   25
      Top             =   915
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   23
      Left            =   1830
      TabIndex        =   24
      Top             =   870
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   22
      Left            =   1740
      TabIndex        =   23
      Top             =   915
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   21
      Left            =   1830
      TabIndex        =   22
      Top             =   855
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   20
      Left            =   1740
      TabIndex        =   21
      Top             =   900
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   19
      Left            =   1755
      TabIndex        =   20
      Top             =   720
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   18
      Left            =   1665
      TabIndex        =   19
      Top             =   765
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   17
      Left            =   1755
      TabIndex        =   18
      Top             =   705
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   16
      Left            =   1665
      TabIndex        =   17
      Top             =   750
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   15
      Left            =   1710
      TabIndex        =   16
      Top             =   1005
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   14
      Left            =   1620
      TabIndex        =   15
      Top             =   1050
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   13
      Left            =   1710
      TabIndex        =   14
      Top             =   990
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   12
      Left            =   1620
      TabIndex        =   13
      Top             =   1035
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   11
      Left            =   1635
      TabIndex        =   12
      Top             =   855
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   10
      Left            =   1545
      TabIndex        =   11
      Top             =   900
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   9
      Left            =   1635
      TabIndex        =   10
      Top             =   840
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   8
      Left            =   1545
      TabIndex        =   9
      Top             =   885
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   7
      Left            =   1695
      TabIndex        =   8
      Top             =   840
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   6
      Left            =   1605
      TabIndex        =   7
      Top             =   885
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   5
      Left            =   1695
      TabIndex        =   6
      Top             =   825
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   4
      Left            =   1605
      TabIndex        =   5
      Top             =   870
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   3
      Left            =   1620
      TabIndex        =   4
      Top             =   690
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   2
      Left            =   1530
      TabIndex        =   3
      Top             =   735
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   0
      Left            =   1620
      TabIndex        =   2
      Top             =   675
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   435
      Index           =   1
      Left            =   1530
      TabIndex        =   1
      Top             =   720
      Width           =   210
   End
   Begin VB.Menu Mnu_VB 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_SHOW 
      Caption         =   "SHOW"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SHOWME
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


Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long



Private Function GetSpecialfolder(CSIDL As Long) As String
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


Function AlwaysOnTop(ByVal hWnd As Long)  'Makes a form always on top
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Function
Function NotAlwaysOnTop(ByVal hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function

'"C:\Program Files\Runtime Software\DriveImage XML\dixml.exe"

'WANT BLUE TOOTH ON HERE

Sub OLD()


'If Command$ = "ART" Or IsIDE = True Then
'If Command$ = "ART" Then
'    If App.PrevInstance = True Then End
'    Me.WindowState = 1
'    Timer_ART_TRIP_WIRE.enabled = True
'
'End If


For Each Control In Controls
        
    x = Control.Name = "Timer_ART_TRIP_WIRE"
    If x = False Then
        On Error Resume Next
        If Control.Caption = ".." Then Control.Visible = False
        On Error GoTo 0

    End If
Next






'ALWAYS




'Shell "C:\Program Files\Media Player Classic\mplayerc.exe", vbNormalFocus
'End






End Sub


Private Sub Form_Load()
'NEED BLUETOOTH CONFIG


Set FS = CreateObject("Scripting.FileSystemObject")


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


On Error Resume Next
For Each Control In Controls
    If InStr(UCase((Control.Name)), "LAB") > 0 Then
        Control.Visible = False
    End If
Next

Dim r As Long
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

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = ((Screen.Height - Me.Height) / 2) - 200




Me.Width = (Screen.Width - 500)
Me.Height = (Screen.Height - 1500)


Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = ((Screen.Height - Me.Height) / 2) - 200


Call SubCode

'
'AlwaysOnTop (Me.hwnd)


End Sub



Sub SubCode()

Dim A1(20)

T = -1
T = T + 1: A1(T) = "Active Idle..."
T = T + 1: A1(T) = "WMICPU1"
T = T + 1: A1(T) = "WinAmp MP3% MINI"
T = T + 1: A1(T) = "VolumeBar MINI"
T = T + 1: A1(T) = "WMICPU2"
T = T + 1: A1(T) = "KAT RECORDER"
T = T + 1: A1(T) = "ClipBoard Logger"
T = T + 1: A1(T) = "00_Schedule_Timer"
T = T + 1: A1(T) = "URL Logger"
T = T + 1: A1(T) = "CID Run Me."
T = T + 1: A1(T) = "Tidal Information..."
T = T + 1: A1(T) = "RS232 Reed Contact Switch Door"
T = T + 1: A1(T) = "Clip Type Form"
T = T + 1: A1(T) = "SUSPEND PROGRAM"
T = T + 1: A1(T) = "Tidal Information..."
T = T + 1: A1(T) = "MY PLAYED"
'T = T + 1: A1(T) = "Winamp v1.x" 'CLASS



For r = 0 To T
    Label1(r) = A1(r)
    Debug.Print Label1(r)
    List1.AddItem Label1(r).Caption
Next

'Me.Show
DoEvents
x = 0
L = 0

For r = 0 To T
    Label1(r) = List1.List(r)
Next


For r = 0 To T
    
    If Label1(r).Width > w1 Then w1 = Label1(r).Width
   Label1(r).Visible = True
    If x > Me.Height - 1100 Then x = 0: L = L + w1 + 50: w1 = 0
    
    Label1(r).Top = x
    Label1(r).Left = L
    
    x = x + Label1(r).Height + 15
    
    'If x > Me.Height Then Exit Sub
    
    
    
   

Next

For r = 0 To T
    i = FindWindow(vbNullString, Label1(r))
    If i = 0 Then Label1(r).BackColor = Label2.BackColor

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

Private Sub Label_Click()

'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = " "
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

Private Sub Label104_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\00 Graphics\Pattern\Kaleidoscope\Kaleidoscope.exe", vbNormalFocus
End
End Sub

Private Sub Label107_Click()

Shell sWindowsFolder + "\system32\OSK.exe"
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
    .sPassword = " "
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
    .sPassword = " "
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
    .sPassword = " "
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
    .sPassword = " "
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

Private Sub Label1_Click(Index As Integer)

If SHOWME = True Then
    'SHOWWINDOW


End If

End Sub

Private Sub Label2_Click()
'matt-555roids
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
'Set oRunas = CreateObject("runas.clsrunas", "Your Server")
With oRunas
    .sDomain = "WORKGROUP"
'    .sUserName = "Username you want to Run the Program  "
    .sUserName = "Matt"
    .sPassword = " "
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
    .sPassword = " "
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
    .sPassword = " "
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
    .sPassword = " "
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

Private Sub Label59_Click()
Shell "C:\Program Files\OE-QuoteFix\OELaunch.exe", vbNormalFocus
End
End Sub

Private Sub Label6_Click()
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
With oRunas
    .sDomain = "WORKGROUP"
    .sUserName = "Admin"
    .sPassword = " "
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
    
    Shell "NOTEPAD ""D:\VB6\VB-NT\00_Best_VB_01\Tidal_Info\MATT-555ROIDS\00_Text_Data_02\VOICE SPEAK LOGG.TXT""", vbMaximizedFocus

End Sub

Private Sub Label63_Click()

Shell "D:\VB6\VB-NT\00_Best_VB_04\Vcap_Src_WebCam\vbVidCapWebCam.exe", vbNormalFocus

End
End Sub

Private Sub Label7_Click()
Set oRunas = CreateObject("runas.clsrunas", "matt-555roids")
With oRunas
    .sDomain = "WORKGROUP"
    .sUserName = "Administrator"
    .sPassword = " "
    .sCommand = "C:\Program Files\Cool2000\cool2000.exe"
    .RunAs 'Call the Run As method
End With
Set oRunas = Nothing
End
End Sub

Private Sub Label8_Click()
ChDrive App.Path
ChDir App.Path
Shell "D:\VB6\VB-NT\00_Best_VB_01\WinAmp Logger\WinAmp Logger.exe", vbNormalNoFocus
End
End Sub

Private Sub Label81_Click()
i = SetScreenSaverState(False, False)
End
End Sub

Private Sub Label83_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\Run Fav Progs\Run Fav Programs.exe", vbNormalNoFocus
End
End Sub

Private Sub Label89_Click()
Shell "C:\Program Files\ProcessTamer\ProcessTamerTray.exe", vbNormalFocus
End
End Sub

'Shell "D:\VB6\VB-NT\00 MP3 Sound\DSS_Olympus\DSS_Olympus.exe"
'End


Private Sub Label9_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\50LastEXE's\50LastEXE.exe", vbNormalFocus
End
End Sub

Private Sub Label92_Click()
Shell "D:\VB6\VB-NT\00_Best_VB_01\DARK\DARK.exe", vbNormalFocus
End
End Sub

Private Sub Label95_Click()
'Screen SAVER


i = SetScreenSaverState(True, True)
End
End Sub

Private Sub Label99_Click(Index As Integer)

End Sub

Private Sub MNU_SHOW_Click()
SHOWME = True
End Sub

Private Sub Mnu_VB_Click()
'If IsIDE = False Then
    Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
'    End
'End If
End
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


If FS.FILEEXISTS("K:\TEMP\ART_PROG_TRIP_WIRE.txt") = False And XHH = True Then End

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






