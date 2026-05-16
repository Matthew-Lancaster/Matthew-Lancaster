VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "--- IRFAN VIEW ---"
   ClientHeight    =   5610
   ClientLeft      =   225
   ClientTop       =   885
   ClientWidth     =   10350
   Icon            =   "#0 00 SEND TO IRFANVIEW.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   9000
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8280
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7320
      Top             =   480
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   8040
      TabIndex        =   43
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label15 
      BackColor       =   &H00800000&
      Caption         =   "Label15"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   48
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label14 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label14"
      Height          =   255
      Left            =   4440
      TabIndex        =   47
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LOOP WRAP ROUND"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5040
      TabIndex        =   46
      Top             =   3840
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Label Label12 
      BackColor       =   &H00008080&
      Caption         =   "Label12"
      Height          =   375
      Left            =   2040
      TabIndex        =   45
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "Label4"
      Height          =   375
      Left            =   2760
      TabIndex        =   44
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   1320
      TabIndex        =   41
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   1320
      TabIndex        =   40
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   1560
      TabIndex        =   39
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   1320
      TabIndex        =   38
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   1560
      TabIndex        =   37
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   1560
      TabIndex        =   36
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   7
      Left            =   1800
      TabIndex        =   35
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   8
      Left            =   1440
      TabIndex        =   34
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   9
      Left            =   1680
      TabIndex        =   33
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   10
      Left            =   1680
      TabIndex        =   32
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   11
      Left            =   1920
      TabIndex        =   31
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   12
      Left            =   1680
      TabIndex        =   30
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   13
      Left            =   1920
      TabIndex        =   29
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   14
      Left            =   1920
      TabIndex        =   28
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   15
      Left            =   2160
      TabIndex        =   27
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   16
      Left            =   480
      TabIndex        =   26
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   17
      Left            =   1560
      TabIndex        =   25
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   18
      Left            =   1560
      TabIndex        =   24
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   19
      Left            =   1800
      TabIndex        =   23
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   20
      Left            =   1560
      TabIndex        =   22
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   21
      Left            =   1800
      TabIndex        =   21
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   22
      Left            =   1800
      TabIndex        =   20
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   23
      Left            =   2040
      TabIndex        =   19
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   24
      Left            =   1680
      TabIndex        =   18
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   25
      Left            =   1920
      TabIndex        =   17
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   26
      Left            =   1920
      TabIndex        =   16
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   27
      Left            =   2160
      TabIndex        =   15
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   28
      Left            =   1920
      TabIndex        =   14
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   29
      Left            =   2160
      TabIndex        =   13
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   30
      Left            =   2160
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   31
      Left            =   2400
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BEGIN -- COMMAND"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   10
      Top             =   3720
      Width           =   4680
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IRFAN VIEW SLIDESHOW TEXT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000FF00&
      Caption         =   "Label10"
      Height          =   615
      Left            =   1080
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CLOSE AFTER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3780
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LOOP WRAP ROUND"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   3660
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BEGIN -- COMMAND --- IRFAN VIEW ---"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   8850
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6720
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6480
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "** FINSIHED HORAY **"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   5040
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Label1"
      Height          =   195
      Left            =   255
      TabIndex        =   0
      Top             =   165
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB"
   End
   Begin VB.Menu MNU_X_SCRIPT 
      Caption         =   "X_SCRIPT"
      NegotiatePosition=   1  'Left
      Begin VB.Menu MNU_EDIT 
         Caption         =   "EDIT"
      End
      Begin VB.Menu MNU_X_EDIT 
         Caption         =   "X_EDIT"
      End
      Begin VB.Menu MNU_X_PROCESS 
         Caption         =   "X_PROCESS"
      End
   End
   Begin VB.Menu MNU_AV 
      Caption         =   "AV_MENU"
      Begin VB.Menu MNU_AV_NAUGHT_OTHER 
         Caption         =   "AV_NAUGHT_OTHER"
      End
      Begin VB.Menu MNU_ALL_AV 
         Caption         =   "ALL AV"
      End
      Begin VB.Menu MNU_AV_AVI 
         Caption         =   "AV_AVI"
      End
      Begin VB.Menu MNU_AV_MPG 
         Caption         =   "AV_MPG"
      End
      Begin VB.Menu MNU_AV_MP4 
         Caption         =   "AV_MP4"
      End
   End
   Begin VB.Menu MNU_TRIP_A_TRON 
      Caption         =   "TRIP_A_TRON"
   End
   Begin VB.Menu MNU_RSR 
      Caption         =   "RUN_SET_RETAINED"
   End
   Begin VB.Menu MNU_RESET 
      Caption         =   "RESET"
   End
   Begin VB.Menu MNU_SPAR 
      Caption         =   "SPAR"
      Begin VB.Menu MNU_INI_BACK_BASE 
         Caption         =   "INI BACK TO BASE"
      End
      Begin VB.Menu MNU_INI_BACK_1000 
         Caption         =   "INI_BACK_1000"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#0 00 SEND TO IRFANVIEW.VBP

Dim A2 'TO ADD
Dim OFBSIZE
Dim MNU_TRIP_A_TRON_Click_VAR

Public DRD, DDATE, BUILD_LAYOUT_MOD

Dim HAVE_RUN_AT_BEGIN, INIA1
Dim x_begin_in_second, FIRSTRUN

Dim OPD2, OPD, OPT, OPT2, TT2, TT3

Dim Z1, X2, X3, W$
Public FS        ' Set = CreateObject("Scripting.FileSystemObject")
Public MTIME, STOPCLOCK
Public AV_AVI, AV_MPG, AV_MP4
Public AV_NAUGHT_OTHER
Public RSR
Public TRIPATRON
Dim DRIVE_PICK_DO


Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


' URL LOGGER VAR SET
' ------------------------------------------------------
' ------------------------------------------------------
' ------------------------------------------------------
' ------------------------------------------------------
' ------------------------------------------------------


'JOBS
'Make url logger less error prone APPEND FILES AN FILE USE WAIT

'Option Explicit

'Dim A1 As String, B1 As String

Public WinDay1, UrlDay1, Edge, HALO, NextStep
Public WinDay2, UrlDay2
Public Day1, Day2, OldDay As Date, OldDay2 As Long, LastURL, CuteURL
Public UrlOrExplor
Dim CountsArry(10)
Dim CountsArr2(20)
Public OldWinTitle, OldAsh, Ash$
Public TrueExit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1
Public vFileName As String

Private Const WM_GETTEXTLENGTH = &HE
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long  'MODULE 1141

Private Declare Function FindWindowUSER Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWndParent As Long, _
     ByVal hWndChildAfter As Long, _
     ByVal lpszClass As String, _
     ByVal lpszTitle As String) _
    As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByRef lParam As Any) _
    As Long

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
    
    
    
Private Const WM_SETTEXT            As Long = &HC
Private Const WM_GETTEXT = &HD

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long  'MODULE 1142

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const OF_READ = &H0
Const OF_WRITE = &H1
Const OF_READWRITE = &H2
Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40
Const OF_PARSE = &H100
Const OF_DELETE = &H200
Const OF_VERIFY = &H400
Const OF_CANCEL = &H800
Const OF_CREATE = &H1000
Const OF_PROMPT = &H2000
Const OF_EXIST = &H4000
Const OF_REOPEN = &H8000

Private Type BROWSEINFO
  hwndOwner      As Long
  pidlRoot       As Long
  pszDisplayName As String
  lpszTitle      As String
  ulFlags        As Long
  lpfnCALLBACK   As Long
  lParam         As Long
  iImage         As Long
End Type

Private Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

Private Type SECURITY_ATTRIBUTES
  nLength              As Long
  lpSecurityDescriptor As Long
  bInheritHandle       As Long
End Type

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

'On Error Resume Next
'MkDir App.Path + "\" + GetComputerName+"\" + GetComputerName
'On Error GoTo 0

Public Sub MNU_HEIGHT_BAR()
    
'    Dim hWndForm        As Long
'    Dim hWndTextBox     As Long
'    Dim RetVal As Long
'    Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
'    Dim WinClass As String, WinTitle As String, WinTitle2 As String

Call URLLogging

End Sub


Public Sub URLLogging()
    
'    Timer1.Interval = 100
    
    Dim hWndForm        As Long
    Dim hWndTextBox     As Long
    Dim RetVal As Long
    Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
    Dim WinClass As String, WinTitle As String, WinTitle2 As String

    Dim Ash$
    
'    ttg = 0

'    If hWndFormi > 0 Then Stop
    'If hWndFormi = GetForegroundWindow Then ttg = 1
'    ttg = 1
            ' --- UPDATE
        ' ALSO MAIN AND CLASS IS AT TOP LSD # DIGGER
'    hWndForme = FindWindow("ExploreWClass", vbNullString)
'    If hWndForme = 0 Then hWndForme = FindWindow("CabinetWClass", vbNullString)
'
'    hWndFormFIRE = FindWindow("MozillaUIWindowClass", vbNullString)
    
'    If hWndForme = GetForegroundWindow Then ttg = 2
'    If hWndFormFIRE = GetForegroundWindow Then ttg = 3
'
'    If ttg = 1 Then hWndForm = hWndFormI
'    If ttg = 2 Then hWndForm = hWndForme
'    If ttg = 3 Then hWndForm = hWndFormFIRE

    hWndFormI = FindWindow(vbNullString, "--- IRFAN VIEW ---")
'    hWndFormI2 = GetParent(hWndFormI)
    
    hWndForm = hWndFormI
    
'    If ttg = 0 Then
'        Exit Sub
'    End If
    
    Dim hWndIEChild As Long
    
'      If ttg = 1 Then
'        hWndIEChild = FindWindowEx(hWndForm, 0&, "WorkerW", vbNullString)
'        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ReBarWindow32", vbNullString)
'        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "Address Band Root", vbNullString)
'        'hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ComboBoxEx32", vbNullString)
'        'hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ComboBox", vbNullString)
'        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "Edit", vbNullString)
'    End If
    
    
    
'    If ttg = 1 Then
        '-#-MSD
        hWndIEChild = FindWindowEx(hWndForm, 0&, "Progman", vbNullString)
        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "MsoCommandBarDock", vbNullString)
        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "wndclass_desked_gskt", vbNullString)
        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "Internet Explorer_Server", vbNullString)
        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "SHELLDLL_DefView", vbNullString)
        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "Edit", vbNullString)
        ' #-LSD
        
'*** Digger *********************************************************************
'
'#    Handle   Class name                        Text
'
'001  2622780  ThunderRT6FormDC                  --- IRFAN VIEW ---
'002  983096   Internet Explorer_Server
'003  393312   SHELLDLL_DefView
'004  196732   Progman                           Program Manager
'
'
'*** Executable Information *****************************************************
        
' *** Digger *********************************************************************
'
'#    Handle  Class name                        Text
'
'001  721704  ThunderFormDC                     --- IRFAN VIEW ---
'002  722532  MsoCommandBar                     Menu Bar
'003  525916  MsoCommandBarDock                 MsoDockTop
'004  329048  wndclass_desked_gsk               IRFAN_VIEW - Microsoft Visual ...
'005  524344  Internet Explorer_Server
'006  393312  SHELLDLL_DefView
'007  196732  Progman                           Program Manager

' *** Digger *********************************************************************
'
'#    Handle   Class name                        Text
'
'001  2622780  ThunderRT6FormDC                  --- IRFAN VIEW ---
'002  525916   MsoCommandBarDock                 MsoDockTop
'003  329048   wndclass_desked_gsk               IRFAN_VIEW - Microsoft Visual...
'004  1247020  MsoCommandBarDock                 MsoDockTop
'005  1246944  wndclass_desked_gsk               URL_Logger - Microsoft Visual...
'006  983096   Internet Explorer_Server
'007  393312   SHELLDLL_DefView
'008  196732   Progman                           Program Manager
'
       
'      *** Digger *********************************************************************
'
'#    Handle   Class name                        Text
'
'001  1311528  ThunderFormDC                     --- IRFAN VIEW ---
'002  788080   MsoCommandBar                     Custom 1
'003  525916   MsoCommandBarDock                 MsoDockTop
'004  329048   wndclass_desked_gsk               IRFAN_VIEW - Microsoft Visual...
'005  68030    ToolbarWindow32
'006  68032    ReBarWindow32
'007  1509462  Notepad++                         E:\01 Desktop\WinDowse.txt - ...
'008  524344   Internet Explorer_Server
'009  393312   SHELLDLL_DefView
'010  196732   Progman                           Program Manager
'
'
       
        ' ---- EDIT AT TOP WINDOWSE
        ' ---- WORK ROM DOWSE DIGGER LSD # IP MSD
    
'    End If
    
'    If ttg = 2 Then
'        hWndIEChild = FindWindowEx(hWndForm, 0&, "WorkerW", vbNullString)
'        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ReBarWindow32", vbNullString)
'        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ComboBoxEx32", vbNullString)
'        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ComboBox", vbNullString)
'        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "Edit", vbNullString)
'    End If
'
'    If ttg = 3 And 50 = 100 Then
'        hWndIEChild = FindWindowEx(hWndForm, 0&, "MozillaWindowClass", vbNullString)
'        hWndIEChild = FindWindowEx(hWndIEChild, 0&, "MozillaUIWindowClass", vbNullString)
'    End If
'
    
    
    If hWndIEChild <> 0 Then
        
    window_hwnd = hWndIEChild
    txtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, 0, 0)
        
    txtlen = SendMessage(window_hwnd, WM_GETTEXT, 255, ByVal WinTitleBuf)
    WinTitle = StripNulls(WinTitleBuf)
    Ash$ = GetWindowTitle(hWndForm)
    
    isd = InStr(Ash$, " - Windows Internet Explorer")
    If isd > 0 Then
        Ash$ = Left$(Ash$, isd - 1)
        UrlOrExplor = True
    Else
        UrlOrExplor = False
    End If
        
'    If OldWinTitle <> WinTitle And OldAsh <> Ash$ Then
    
    If OldWinTitle <> WinTitle And Trim(WinTitle) <> "" And NextStep = 2 Then NextStep = 0
    
    If OldWinTitle <> WinTitle And Trim(WinTitle) <> "" And NextStep = 0 Then
        NextStep = 2
        'Debug.Print Time
        'Debug.Print OldWinTitle
        'Debug.Print WinTitle
        'Debug.Print OldAsh
        'Debug.Print Ash$
        
    '    ash2$ = Ash$
        WinTitle2 = WinTitle
        XX = 0
        
    '    If Trim(Ash$) = "" And Trim(WinTitle2) = "" Then xx = 1
    '    If Ash$ = "Windows Internet Explorer" And Trim(WinTitle2) = "" Then xx = 1
    '    If Ash$ = "Windows Explorer" And Trim(WinTitle2) = "" Then xx = 1
        
    '    If ash2$ <> "" And WinTitle2 = "" And Mid$(ash2$, 2, 1) = ":" Then WinTitle2 = ash2$
                
    '    If xx = 0 Then Call DisplayLogging("", ash2$, WinTitle2)
        OldWinTitle = WinTitle
    '    OldAsh = Ash$
        Ash$ = ""
    End If
        
    
    If OldAsh <> Ash$ And Ash$ <> "" And NextStep = 2 Then
    
    '    NextStep = 0
        'Debug.Print Time
        'Debug.Print OldWinTitle
        'Debug.Print WinTitle
        'Debug.Print OldAsh
        'Debug.Print Ash$
        
        ASH2$ = Ash$
        WinTitle2 = WinTitle
        XX = 0
        
        If Trim(Ash$) = "" Or Trim(WinTitle2) = "" Then XX = 1
    '    If Ash$ = "Windows Internet Explorer" And Trim(WinTitle2) = "" Then xx = 1
    '    If Ash$ = "Windows Explorer" And Trim(WinTitle2) = "" Then xx = 1
        
    '    If ash2$ <> "" And WinTitle2 = "" And Mid$(ash2$, 2, 1) = ":" Then WinTitle2 = ash2$
                
        If XX = 0 Then
'            Call DisplayLogging("", ASH2$, WinTitle2)
            If Mid$(ASH2$, 1, 5) <> "Http:" Then NextStep = 0
        End If
        
        OldWinTitle = WinTitle
        OldAsh = Ash$
    End If
        

End If
    
End Sub

Public Function StripNulls(OriginalStr As String) As String
        ' This removes the extra Nulls so String comparisons will work
        If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
        End If
        StripNulls = OriginalStr
End Function

Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim l As Long
   Dim s As String
   l = GetWindowTextLength(hWnd)
   s = Space(l + 1)
   GetWindowText hWnd, s, l + 1
   GetWindowTitle = Left$(s, l)

End Function


Sub Test()

    Dim hWndForm        As Long
    Dim hWndTextBox     As Long
    Dim RetVal As Long
    Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
    Dim WinClass As String, WinTitle As String, WinTitle2 As String

    Dim Ash$
    
    ttg = 0
    hWndForm = Me.hWnd 'FindWindow("IEFrame", vbNullString)
    
    'If hWndFormi = GetForegroundWindow Then ttg = 1
    
    Dim hWndIEChild As Long
    
    
    For R = 1 To 10
        
        'hWndIEChild = FindWindowEx(hWndForm, 0&, vbNullString, vbNullString)
        hWndIEChild = GetWindow(hWndForm, GW_CHILD)
        WinTitle = GetWindowTitle(hWndIEChild)
        'txtlen = SendMessage(hWndIEChild, WM_GETTEXTLENGTH, 0, 0)
        'txtlen = SendMessage(hWndIEChild, WM_GETTEXT, 255, ByVal WinTitleBuf)
    
        'WinTitle = StripNulls(WinTitleBuf)
        Debug.Print str(hWndIEChild) + " -- " + WinTitle
    
        'W = GetWindowTitle(hWndForm)
        hWndForm = hWndIEChild
    
    Next
    
    hWndIEChild = FindWindowEx(hWndIEChild, 0&, "ReBarWindow32", vbNullString)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
    If Dir("C:\TEMP\INFRA_WORK_COMPLETE.TXT") <> "" Then
        'KILL WITH UNLOAD ME
        Kill "C:\TEMP\INFRA_WORK_COMPLETE.TXT"
    End If

End Sub

Private Sub LabDR_Click(index As Integer)
Text1.Text = Chr(index + 65) + Mid(Text1.Text, 2)
LabDR(index).BackColor = Label4.BackColor
DRIVE_PICK_DO = True
End Sub

Private Sub Form_Load()

'SHELL "C:\Program Files\Greatis\WinDowse\WinDowse.exe",vbNormalFocus
'i = App.title
'Debug.Print i
'End
Set FS = CreateObject("Scripting.FileSystemObject")
MNU_TRIP_A_TRON_Click_VAR = False
'Me.Caption = "--- IRFAN VIEW ---VV"
Me.Caption = "--- IRFAN VIEW ---"
AV_ALL = False
AV_AVI = False
AV_MP4 = False
AV_MPG = False
TRIPATRON = 0
RSR = 0
FIRSTRUN = True
STOPCLOCK = 0

Label19 = " ---- ---- ----------- -- \/\/\/\/\/\/\/ --- IRFAN VIEW --- \/\/\/\/\/\/\/ -- ----------- ---- ---- "
Label19.BackColor = &HFF0000
'Label3 = "--- IRFAN VIEW ---"
'Label5 = "IRFAN VIEW SLIDESHOW TEXT"
'Label11 = "IRFAN VIEW SLIDESHOW TEXT"
'Label11 = "LOOP WRAP ROUND"
Label9 = "LOOP WRAP ROUND"
'Label9 = "CLOSE AFTER"
'Label11 = "CLOSE AFTER"
Label5 = "CLOSE AFTER" + vbCrLf + vbCrLf + vbCrLf
'Label19 = "BEGIN -- COMMAND"
Label3 = "BEGIN -- COMMAND"
Label3.BackColor = &H800000
Label8 = "** FINSIHED HORAY **"
Label8 = "Ready TO OPP"
MNU_TRIP_A_TRON_Click_VAR = True

If FindWindow(vbNullString, "INFAR_BATCH") > 0 Then
    If Dir("C:\TEMP\INFRA_WORK_DRIVE_SET.TXT") <> "" Then
        FF7 = FreeFile
        Open "C:\TEMP\INFRA_WORK_DRIVE_SET.TXT" For Input As #FF7
        DRD = Input(LOF(FF7), FF7)
        Close #FF7
    End If
End If

Call GETFILE_SLIDE_SHOW
Call LABELFILEINI
Call Timer3_Timer
BUILD_LAYOUT_MOD = -9
Call CHECK_CMD_IS_RUN_OR_LEFT

    

If FindWindow(vbNullString, "INFAR_BATCH") > 0 Then
        Call MNU_RSR_MODIFY
        HAVE_RUN_AT_BEGIN = False
    Else
        HAVE_RUN_AT_BEGIN = True
        Timer2.Enabled = True
End If


BUILD_LAYOUT_MOD = -9
Call CHECK_CMD_IS_RUN_OR_LEFT

'center screen
'Me.Left = (Screen.Width / 2 - Me.Width / 2) + 80
Me.Top = (Screen.Height / 2 - Me.Height / 2) - 2000
Me.Top = 400 '(Me.Height) +
Me.Left = Screen.Width - (Me.Width) - 300

'Call URLLogging


Me.Show
DoEvents


End Sub


Sub GETFILE_SLIDE_SHOW()

    'DDATE = "C:\#_IRFAN_VIEW_USB\TEMP_VBA\IRFAN_SLIDESHOW_" + DDATE + ".txt"
    DDAT2 = "C:\#_IRFAN_VIEW_USB\TEMP_VBA\*.txt"
    
    DDATE = Dir(DDAT2)
    Do While DDATE <> ""
        DDAT3 = DDATE
        DDATE = Dir()
    Loop
    DDATE = "C:\#_IRFAN_VIEW_USB\TEMP_VBA\" + DDAT3

End Sub


Sub BUILD_LAYOUT()
    

'Call MNU_HEIGHT_BAR
'Call URLLogging


On Error Resume Next
For Each Control In Controls
    If Control.Visible = True Then
        BUILD_LAYOUT_MODi = Len(Control.Caption) + BUILD_LAYOUT_MODi
    End If
Next

If BUILD_LAYOUT_MOD = BUILD_LAYOUT_MODi Then Exit Sub
BUILD_LAYOUT_MOD = BUILD_LAYOUT_MODi

'----------------------------------------------------------------------------------------------
    
    
    


Call LABDRIVECHR

Call LABS

Exit Sub


'FindWindow(vbNullString, Me.Caption)
'If FindWindow(vbNullString, Me.Caption) > 0 Then
'
'    Dim rctFrame As RECT
'    Dim lpTL As POINTAPI, lpBR As POINTAPI
'
'    GetWindowRect Me.hWnd, rctFrame
'
'    lpTL.x = rctFrame.Left
'    lpTL.y = rctFrame.Top
'    lpBR.x = rctFrame.Right
'    lpBR.y = rctFrame.Bottom
'
'    RECTH = rctFrame.Bottom - rctFrame.Top
'    RECTH = RECTH * Screen.TwipsPerPixelY

'SHORTOFHEIGHT = TOPHEIGHT - RECTH
'Debug.Print SHORTOFHEIGHT
'Debug.Print TOPHEIGHT
'Debug.Print TOPHEIGHT + SHORTOFHEIGHT
'Debug.Print Label3.Top + Label3.Height

'End If
'If SHORTOFHEIGHT > 0 Then TOPHEIGHT = TOPHEIGHT + SHORTOFHEIGHT


Me.Height = TOPHEIGHT


End Sub


Sub LABDRIVECHR()

On Error Resume Next

Me.Width = Screen.Width - 3000
Me.Left = Screen.Width - Me.Width



XT = (Me.Width + 100 - 16) / 26


XC = 64
For Each Control In LabDR
    XC = XC + 1
    If XC >= 91 Then
        Control.Visible = False
    End If
Next
    

For Each Control In LabDR
    Control.AutoSize = True
Next
DoEvents

XC = 64
For Each Control In LabDR
    XC = XC + 1
    If XC >= 91 Then Exit For
    Control.Visible = True
    Control.Caption = "-" + Chr(XC) + "-"
    Control.AutoSize = False

Next

'XC = 64
'CHH = 0
'For Each Control In LabDR
'    XC = XC + 1
'    If XC > 92 Then Exit For
'
''    Control.Height = Control.Height + 40
''    If Control.Height > CHH Then CHH = Control.Height
'Next
'CHH = CHH + 40


XC = 64
For Each Control In LabDR
    XC = XC + 1
    If XC > 92 Then Exit For
    Control.Height = Control.Height + 80
    Control.Top = 0
    Control.Width = XT 'XT = (XX + 100 - 16) / 26
    Control.Left = XT * (XC - 65)
Next




XC = 64
For Each Control In LabDR
    If XC > 92 Then Exit For
    XC = XC + 1
    
    
        If Control.BackColor <> Label4.BackColor And _
            Control.BackColor <> &HC0FFC0 And _
            Control.BackColor <> &HC0FFC0 Then
    
        
            If XC Mod 2 = 0 Then Control.BackColor = &HC0FFC0
            If XC Mod 2 = 1 Then Control.BackColor = &H80FFFF
            If InStr("*" + DRD, Chr(XC)) > 0 Then Control.BackColor = Label4.BackColor
    
            If DRIVE_ISREADY_AND_WANT(Chr(XC)) = False Then
                Control.BackColor = Label12.BackColor
            End If
        End If

Next

LabDR(0).Left = 0

End Sub



Sub LABS()

On Error Resume Next




Label19.Width = Me.Width
Label3.Width = Me.Width
Label5.Width = Me.Width

Label19.Top = LabDR(0).Height
Label5.Top = Label19.Height + Label19.Top
Label3.Top = Label5.Height + Label5.Top

Me.Height = Label3.Height + Label3.Top + 800

'Me.Width = (Screen.Width - Me.Width) - 2000


For Each Control In Controls
    If Control.Name <> "LabDR" Then
        If Control.Visible = True Then
            Control.AutoSize = True
            Control.WordWrap = True
            Control.FontSize = 17
            'If Control.Width > XX Then XX = Control.Width
            'Control.Left = 0
        End If
    End If
Next
'XC = 65
'XD = 0
'

DoEvents
For Each Control In Controls
    If Control.Name <> "LabDR" Then
        If Control.Visible = True Then
    Control.AutoSize = False
    Control.Width = Me.Width
    AutoSize = False
    
    Control.Left = 0
    'XD = XD + XT
        End If
    End If

Next

Label19.Width = Me.Width
Label3.Width = Me.Width
Label5.Width = Me.Width

Label19.Top = LabDR(0).Height
Label5.Top = Label19.Height + Label19.Top
Label3.Top = Label5.Height + Label5.Top

Me.Height = Label3.Height + Label3.Top + 800

'Label3.Top = LABDRHEIGHT
'Label5.Top = Label5.Height + Label5.Top

'YY = Label3.Top + Label3.Height
'XX = Me.Width

'Me.Width = Screen.Width - 2000
'Me.Height = YY + 800

'TOPHEIGHT = Form1.Height

End Sub


Function DRIVE_ISREADY_AND_WANT(CHRGIVE)
    
    DRIVE_ISREADY_AND_WANT = False
    
    On Error Resume Next
    drvpath = CHRGIVE
    Dim FS, d, s, t, ty
    Set FS = CreateObject("Scripting.FileSystemObject")
    Set d = FS.GetDrive(drvpath)
    If Err.Number > 0 Then Exit Function
    Select Case d.DriveType
        Case 0: t = "Unknown"
        Case 1: t = "Removable"
        Case 2: t = "Fixed"
        Case 3: t = "Network"
        Case 4: t = "CD-ROM"
        Case 5: t = "RAM Disk"
    End Select
    s = "Drive " & d.DriveLetter & ": - " & t
    
    ty = d.DriveType
'    If d.IsReady = True And (ty = 4 Or ty = 3) Then Stop
'    If d.IsReady = True And (ty = 4 And ty = 3) Then Stop
    
    If d.IsReady Then
        If d.VolumeName <> "RAMDISK" And _
        ty <> 4 And _
        ty <> 3 And _
        ty <> 5 Then

'    If d.IsReady And (ty <> 4 Or ty <> 3) Then
        s = s & vbCrLf & "Drive is Ready."
        DRIVE_ISREADY_AND_WANT = True
    End If
    End If


End Function

Private Sub Label3_Click()


'EXIT IF DONE
'EXIT IF AT PROGRESS

If Dir("C:\TEMP\INFRA_WORK_COMPLETE.TXT") <> "" Then Exit Sub

If DRIVE_PICK_DO = False Then
    MsgBox "CONVEY MESSAGE  -- PICK A DRIVE"
    Exit Sub
End If


If FindWindow(vbNullString, "INFAR_BATCH") > 0 Then
'    Call MNU_RSR_MODIFY
'    HAVE_RUN_AT_BEGIN = False
'    Else
'    HAVE_RUN_AT_BEGIN = True
'
    MsgBox "CONVEY MESSAGE  -- A COMMAND BATCH ALREADY IN PROGRESS"
    Exit Sub

End If
'If Label3.Caption = "BEGIN -- COMMAND" Then
         
    
    'On Error Resume Next
'    ChDrive W$
'    ChDir W$
    
    FSS = "C:\TEMP\INFAR_DIR_BUILD_TIME_END.TXT"
    If Dir(FSS) <> "" Then Kill FSS
    
    STOPCLOCK = 0
    
    FSS = "C:\TEMP\INFRA_WORK_COMPLETE.TXT"
    If Dir(FSS) <> "" Then Kill FSS
    
    If x1 > 0 Then V1 = " /reloadonloop /quite"
    
    W2$ = LCase(W$ + "\" + V1 + """")
    
    If Text1.Text <> "" Then W2$ = Text1.Text + ":\"
'
'    If X3 = 1 Then
'        W$ = "D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDESHOW.txt"
'        W2$ = LCase(W$ + "" + V1 + """")
'    End If
    
    FTAG = ""
    
    FTAG2 = FTAG2 + "@C:" + vbCrLf
    FTAG2 = FTAG2 + "@CD \" + vbCrLf
    FTAG2 = FTAG2 + "@ECHO OFF" + vbCrLf
    FTAG2 = FTAG2 + "TITLE INFAR_BATCH" + vbCrLf
    FTAG2 = FTAG2 + "@ECHO ON" + vbCrLf
    FTAG2 = FTAG2 + "@ECHO *********************************************************" + vbCrLf
    FTAG2 = FTAG2 + "@ECHO *********************************************************" + vbCrLf
    FTAG2 = FTAG2 + "@ECHO *********************************************************" + vbCrLf
    FTAG2 = FTAG2 + "@ECHO *********************************************************" + vbCrLf
    FTAG2 = FTAG2 + "@ECHO *********************************************************" + vbCrLf
    
    FTAG2 = FTAG2 + "@ECHO BEGIN PROCESS ACTIVE BUILD TREE INFAR" + vbCrLf
    
    
    FF7 = FreeFile
    Open "C:\TEMP\INFRA_WORK_DRIVE_SET.TXT" For Output As #FF7
    
    For Each Control In LabDR
        If Control.BackColor = Label4.BackColor Then
            
            W2$ = Mid(Control.Caption, 2, 1) + ":\"
            Print #FF7, Mid(Control.Caption, 2, 1);
            
            DTAG = DTAG + Mid(Control.Caption, 2, 1)
            
            FTAG = FTAG + "@ECHO TIME --- " + vbCrLf
            FTAG = FTAG + "DATE /T" + vbCrLf
            FTAG = FTAG + "TIME /T" + vbCrLf
            
            FTAG = FTAG + "@DATE /T >C:\TEMP\INFAR_DIR_BUILD_TIME_BEGIN.TXT" + vbCrLf
            FTAG = FTAG + "@TIME /T >>C:\TEMP\INFAR_DIR_BUILD_TIME_BEGIN.TXT" + vbCrLf
            
            If JPG = True Then
                '---BUILD MESSURE LENGHT TIME
                FTAG = FTAG + "@TIME /T >>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
                ' = FTAG + "@ECHO DIR " + W2$ + "*.JPG /OGN /B /S >>C:\TEMP\IRFAN_SLIDESHOW.tx t>>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
                'FTAG = FTAG + "@TIME /T >>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
                '---
            End If
            
            If AV_AVI = True Then
                '---BUILD MESSURE LENGHT TIME
                FTAG = FTAG + "@TIME /T >>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
                'FTAG = FTAG + "@ECHO DIR " + W2$ + "*.AVI /OGN /B /S >>C:\TEMP\IRFAN_SLIDESHOW.tx t>>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
                'FTAG = FTAG + "@TIME /T >>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
                '---
            End If
            
            If AV_MPG = True Then
                '---BUILD MESSURE LENGHT TIME
                FTAG = FTAG + "@TIME /T >>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
                'FTAG = FTAG + "@ECHO DIR " + W2$ + "*.MPG /OGN /B /S >>C:\TEMP\IRFAN_SLIDESHOW.tx t>>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
                'FTAG = FTAG + "@TIME /T >>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
                '---
            End If
            
            If AV_MP4 = True Then
                '---BUILD MESSURE LENGHT TIME
                FTAG = FTAG + "@TIME /T >>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
                'FTAG = FTAG + "@ECHO DIR " + W2$ + "*.MP4 /OGN /B /S PIPE --- C:\TEMP\IRFAN_SLIDESHOW.txT >>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
                'FTAG = FTAG + "@TIME /T >>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
                '---
            End If
            
            
            If AV_NAUGHT_OTHER = False Then JPG = True
            
            JPG = True
            AV_AVI = True
            AV_MPG = True
            AV_MP4 = True
            
            If JPG = True Then
                FTAG = FTAG + "TIME /T" + vbCrLf
                FTAG = FTAG + "@TIME /T >>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
                FTAG = FTAG + "DIR " + W2$ + "*.JPG /OGN /B /S >>C:\TEMP\IRFAN_SLIDESHOW.txt" + vbCrLf
                FTAG2 = FTAG2 + "@ECHO DIR " + W2$ + "*.JPG /OGN /B /S PIPE APPEND ---- C:\TEMP\IRFAN_SLIDESHOW.txt" + vbCrLf
                FTAG = FTAG + "TIME /T" + vbCrLf
                FTAG = FTAG + "@TIME /T >>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
            End If
            If AV_AVI = True Then
                FTAG = FTAG + "TIME /T" + vbCrLf
                FTAG = FTAG + "DIR " + W2$ + "*.AVI /OGN /B /S >>C:\TEMP\IRFAN_SLIDESHOW-AV.txt" + vbCrLf
                FTAG2 = FTAG2 + "@ECHO DIR " + W2$ + "*.AVI /OGN /B /S PIPE APPEND ---- C:\TEMP\IRFAN_SLIDESHOW.txt" + vbCrLf
                FTAG = FTAG + "TIME /T" + vbCrLf
                FTAG = FTAG + "@TIME /T >>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
            End If
            If AV_MPG = True Then
                FTAG = FTAG + "TIME /T" + vbCrLf
                FTAG = FTAG + "DIR " + W2$ + "*.MPG /OGN /B /S >>C:\TEMP\IRFAN_SLIDESHOW-AV.txt" + vbCrLf
                FTAG2 = FTAG2 + "@ECHO DIR " + W2$ + "*.MPG /OGN /B /S PIPE APPEND ---- C:\TEMP\IRFAN_SLIDESHOW.txt" + vbCrLf
                FTAG = FTAG + "TIME /T" + vbCrLf
                FTAG = FTAG + "@TIME /T >>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
            End If
            If AV_MP4 = True Then
                FTAG = FTAG + "TIME /T" + vbCrLf
                FTAG = FTAG + "DIR " + W2$ + "*.MP4 /OGN /B /S >>C:\TEMP\IRFAN_SLIDESHOW-AV.txt" + vbCrLf
                FTAG2 = FTAG2 + "@ECHO DIR " + W2$ + "*.MP4 /OGN /B /S PIPE APPEND ---- C:\TEMP\IRFAN_SLIDESHOW.txt" + vbCrLf
                FTAG = FTAG + "TIME /T" + vbCrLf
                FTAG = FTAG + "@TIME /T >>C:\TEMP\INFAR_DIR_BUILD_LENGHT_TIME.TXT" + vbCrLf
            End If
                
                
        End If
    
    Next
    
    FTAG = FTAG + "@ECHO CONVEY MESSAGE WORK DO >C:\TEMP\INFRA_WORK_COMPLETE.TXT" + vbCrLf
    FTAG = FTAG + "@DATE /T >C:\TEMP\INFAR_DIR_BUILD_TIME_END.TXT" + vbCrLf
    FTAG = FTAG + "@TIME /T >>C:\TEMP\INFAR_DIR_BUILD_TIME_END.TXT" + vbCrLf

    
    
    FTAG = FTAG + "@ECHO E:\My Webs\MatthewLan.Com Web\LoveFolder\Cool\NHS\05 Agent Orange.jpg >>C:\TEMP\IRFAN_SLIDESHOW.txt" + vbCrLf
    FTAG = FTAG + "@ECHO E:\My Webs\MatthewLan.Com Web\LoveFolder\Cool\NHS\05 Agent Orange.jpg >>C:\TEMP\IRFAN_SLIDESHOW.txt" + vbCrLf
    
    FTAG = FTAG + "@TYPE  C:\TEMP\IRFAN_SLIDESHOW.txt >C:\TEMP\IRFAN_SLIDESHOW-BACKUP.txt" + vbCrLf
    FTAG = FTAG + "@TYPE  C:\TEMP\IRFAN_SLIDESHOW-AV.txt >C:\TEMP\IRFAN_SLIDESHOW-BACKUP.txt" + vbCrLf
    
    FTAG = FTAG + "@EXIT" + vbCrLf
    
    DDATE = Format(Now, "YYYY_MM_DD_HH_MM_SS")
    
    If Dir("C:\#_IRFAN_VIEW_USB", vbDirectory) = "" Then
        MkDir "C:\#_IRFAN_VIEW_USB"
    End If
    
    If Dir("C:\#_IRFAN_VIEW_USB\TEMP_VBA", vbDirectory) = "" Then
        MkDir "C:\#_IRFAN_VIEW_USB\TEMP_VBA"
    End If
    
    DDATE = "C:\#_IRFAN_VIEW_USB\TEMP_VBA\IRFAN_SLIDESHOW_" + DDATE + ".txt"
    
    FTAG = FTAG + "@TYPE  C:\TEMP\IRFAN_SLIDESHOW.txt >" + DDATE + vbCrLf
    FTAG = FTAG + "@TYPE  C:\TEMP\IRFAN_SLIDESHOW-AV.txt >>" + DDATE + vbCrLf
    
    FTAG2 = FTAG2 + "@ECHO *********************************************************" + vbCrLf
    FTAG2 = FTAG2 + "@ECHO *********************************************************" + vbCrLf
    FTAG2 = FTAG2 + "@ECHO *********************************************************" + vbCrLf
    FTAG2 = FTAG2 + "@ECHO *********************************************************" + vbCrLf
    FTAG2 = FTAG2 + "@ECHO *********************************************************" + vbCrLf
    
    FTAG = FTAG2 + FTAG
    
    'Debug.Print FTAG
    
    Close #FF7
    
    FF = FreeFile
    Open App.Path + "\DATA\BATCHCMD.BAT" For Output As FF
        Print #FF, FTAG
    Close #FF
    
    FF = FreeFile
    Open "C:\TEMP\INFAR BATCHCMD.BAT" For Output As FF
        Print #FF, FTAG
    Close #FF
    
    FF = FreeFile
    Open App.Path + "\DATA\INFAR BATCHCMD -- " + Format(Now, "YYYY-MM-DD -HH-MM-SS") + "--.BAT.TXT" For Output As FF
        Print #FF, FTAG
    Close #FF
    
    
    ' -- NURSARY 02
    FF = FreeFile
    Open App.Path + "\DATA\INFAR BATCHCMD -- DATA-TIME.TXT" For Input As FF
        GI = Input$(LOF(FF), FF)
        GI = Format(DateValue(GI) + TimeValue(GI), "YYYY-MMM-DD HH:MM:SS")
        GIA = DateValue(GI) + TimeValue(GI)
    Close #FF
    
    
    
    
    ' ------------ \/\/\/\/\\/\/\
    
    
'    If Dir(DDATE) <> "" Then
'
'        FF = FreeFile
'        Open DDATE For Input As FF
'            GIH = Input$(LOF(FF), FF)
'        Close #FF
'
'        'Kill DDATE
'
'        FF = FreeFile
'        Open App.Path + "\IRFAN_SLIDESHOW ARCHIVE -- " + Format(GIA, "YYYY-MM-DD -HH-MM-SS") + "--.txt" For Output As FF
'            Print #FF, GIH
'        Close #FF
'
'    End If
    
    
    ' ------------ \/\/\/\/\\/\/\
    
    
    '
    '' -- TIME STAMP OUT
    'FF = FreeFile
    'Open App.Path + "\DATA\IRFAN_SLIDESHOW -- " + GI + " -- .txt" For Output As FF
    '    Print FF, GI;
    'Close #FF
    
    
    
    If Dir(App.Path + "\DATA", vbDirectory) = "" Then
        MkDir App.Path + "\DATA"
    End If
    ' -- NURSARY 01
    FF = FreeFile
    Open App.Path + "\DATA\INFAR BATCHCMD -- DATA-TIME.TXT" For Output As FF
        Print #FF, Format(Now, "DD-MMM-YYYY HH:MM:SS");
    Close #FF
    
    
    
    
    Shell "CMD /C """ + App.Path + "\DATA\BATCHCMD.BAT""", vbMinimizedNoFocus
    'Shell "CMD " + FTAG, vbNormalFocus
    'Shell FTAG, vbNormalFocus
    
'    Do
'        If FileInUse(DDATE) = True Then
'            'Label8 = "FILE IS PROCESS ACTIVE BUILD - " + "DIR " + W2$ + "*.JPG"
'            Label8 = "FILE IS ACTIVE BUILD - DIR " + DTAG '" '+ W2$ + "*.JPG"
'            'Label8.AutoSize = True
'            Call BUILD_LAYOUT(True)
'
'            Exit Do
'        End If
'        Sleep 1
'        DoEvents
'    Loop Until ENDFLAG = True

    MTIME = Now
    Label3.Caption = "BEGIN -- INFAR"
    'Call Timer1_Timer
    
'    'YELLOW
'    Label3.BackColor = Label14.BackColor
'    Label3.ForeColor = Label14.ForeColor
'
'    Exit Sub

'End If





'If InStr(Label3.Caption, "BEGIN -- INFAR") > 0 Then
'    Call Timer1_Timer

    
'    Do
'        If FileInUse(DDATE) = False _
'            Then
'                Exit Do
'            End If
'        Sleep 1
'        DoEvents
'
'    Loop Until ENDFLAG = True
    
    
    
    
    
    
    
    '
    '@TIME /T
    '
    '@IF EXIST "D:\VB6\VB-NT\00_Best_VB_01\#0 SMALL PROGS TO HAVE WITH BATCH FILE CODE\KILL A BATCH CMD CODE FILE\KILL A BATCH CMD CODE FILE_01_ENTRY.EXE"  "D:\VB6\VB-NT\00_Best_VB_01\#0 SMALL PROGS TO HAVE WITH BATCH FILE CODE\KILL A BATCH CMD CODE FILE\KILL A BATCH CMD CODE FILE_01_ENTRY.EXE"
    '
'    RSR
    
    
    
'
'    If Text1 <> "" Or RSR = True Then
'        Call MNU_X_PROCESS_Click
'            End
'    End If
    
'    'SLIDE OF DIR NOT SUB DEBTH
'    ZZ = "C:\Program Files\IrfanView\i_view32.exe ""/slideshow=" + W2$
'    Debug.Print ZZ
'    Shell ZZ, vbNormalFocus
'
    'If STOPCLOCK > 0 Then
    '    End
    'End If
    
'End If

'End If
End Sub

Private Sub Label4_Click()
'Label4
End Sub

Private Sub Label11_Click()
'Label11
'If X3 = 0 Then
'    Label11.BackColor = Label10.BackColor
'    Label11.ForeColor = 0
'Else
'    Label11.BackColor = Label3.BackColor
'    Label11.ForeColor = RGB(255, 255, 255)
'End If
'X3 = Not X3
End Sub

Private Sub Label8_Click()
'Label8
End Sub

Private Sub MNU_AV_Click()
AV = True
End Sub

Private Sub MNU_ALL_AV_Click()
AV_ALL = True
AV_AVI = True
AV_MP4 = True
AV_MPG = True
MsgBox "-- /\/\/\/\ --- ALL AUDIO VIDEO --- /\/\/\/\ --"

End Sub

Private Sub MNU_AV_AVI_Click()
AV_AVI = True
MsgBox "AV_AVI = TRUE"

End Sub

Private Sub MNU_AV_MP4_Click()
AV_MP4 = True
MsgBox "AV_MP4 = TRUE"

End Sub

Private Sub MNU_AV_MPG_Click()
AV_MPG = True
MsgBox "AV_MPG = TRUE"

End Sub

Private Sub MNU_AV_NAUGHT_OTHER_Click()
    
    AV_NAUGHT_OTHER = True
    ' OF JPG ---------
    If Not (AV_AVI = False And AV_MP4 = False And AV_MPG = False) Then
            MsgBox "SELECT NAUGHT OTHER -- JPG EXCLUDE"
        Else
            MsgBox "SELECT NAUGHT OTHER -- JPG EXCLUDE" + vbCrLf + "NOT A SELECTION MADE OF " + vbCrLf + "- ELSE -" + vbCrLf + "DECIDE PROCESS  TO" + vbCrLf + "-- /\/\/\/\ --- ALL AUDIO VIDEO --- /\/\/\/\ --"
            AV_ALL = True
            AV_AVI = True
            AV_MP4 = True
            AV_MPG = True
    End If

    
End Sub

Private Sub MNU_EDIT_Click()
    Shell "C:\PROGRAM FILES\NOTEPAD++\NOTEPAD++.EXE C:\TEMP\IRFAN_SLIDESHOW.TXT", vbNormalFocus
End Sub

Private Sub MNU_INI_BACK_1000_Click()
'Me.WindowState = vbMinimized

Open "C:\Program Files\IrfanView\i_view32.ini" For Input As #1
    Do
        Line Input #1, l
        If InStr("-" + l, "StartIndex=") > 0 Then
            A1 = Val(Mid(l, Len("StartIndex=") + 1))
        End If
        If EOF(1) Then A1 = "2"
    Loop Until A1 <> ""
Close #1


FR = FreeFile
Open "C:\Program Files\IrfanView\i_view32.ini" For Input As #FR
RC = Input(LOF(FR), FR)
Close #FR

A1 = Trim(str(A1))
B1 = Trim(str(A1 - 1000))
RC = Replace(RC, "StartIndex=" + A1, "StartIndex=" + B1)
'Label3 = "LOAD INFAR INI -- " + str(B1)
INIA1 = A1

FR = FreeFile
Open "C:\Program Files\IrfanView\i_view32.ini" For Output Lock Read Write As #FR
Print #FR, RC
Close #FR

Call LABELFILEINI


End Sub

Private Sub MNU_INI_BACK_BASE_Click()

'Me.WindowState = vbMinimized

Open "C:\Program Files\IrfanView\i_view32.ini" For Input As #1
    Do
        Line Input #1, l
        If InStr("-" + l, "StartIndex=") > 0 Then
            A1 = Val(Mid(l, Len("StartIndex=") + 1))
        End If
        If EOF(1) Then A1 = "2"
    
    'LOOP UNTILL A1 GOT SOMETHING OR EOF
    Loop Until A1 <> ""
Close #1

FR = FreeFile
Open "C:\Program Files\IrfanView\i_view32.ini" For Input As #FR
RC = Input(LOF(FR), FR)
Close #FR

A1 = Trim(str(A1))
B1 = Trim(str(1))
RC = Replace(RC, "StartIndex=" + A1, "StartIndex=" + B1)
Label3 = "LOAD INFAR INI -- " + str(B1)
INIA1 = A1

FR = FreeFile
Open "C:\Program Files\IrfanView\i_view32.ini" For Output Lock Read Write As #FR
Print #FR, RC
Close #FR

Call LABELFILEINI

'Label5 = "CLOSE AFTER"


End Sub

Private Sub MNU_RESET_Click()
    
    
    
    MNU_TRIP_A_TRON_Click_VAR = True
    Label3 = "BEGIN -- COMMAND"
    OPD2 = 0: OPD = 0: OPT = 0: OPT2 = 0
    DRD = ""
    BUILD_LAYOUT_MOD = -9
    
    For Each Control In LabDR
        Control.BackColor = &H80FFFF
    Next
    
    Call LABDRIVECHR
    If Dir("C:\TEMP\INFRA_WORK_DRIVE_SET.TXT") <> "" Then
    Kill "C:\TEMP\INFRA_WORK_DRIVE_SET.TXT"
    DRD = ""
    End If
    
    If Dir("C:\TEMP\INFRA_WORK_COMPLETE.TXT") <> "" Then
    Kill "C:\TEMP\INFRA_WORK_COMPLETE.TXT"
    End If
    
    'Call Form_Load
    BUILD_LAYOUT

End Sub

Private Sub MNU_RSR_Click()

    'RSR = True
    
    'GET ARCHIVE FILE
    If Dir(DDATE) = "" Then Exit Sub
        
    Call GETFILE_SLIDE_SHOW
    
    'Shell "C:\Program Files\IrfanView\i_view32.exe ""/slideshow=C:\TEMP\IRFAN_SLIDESHOW.txt /ONE /bf /silent /gray /RELOADONLOOP /QUITE""", vbMinimizedNoFocus

    FILE1 = "C:\TEMP\IRFAN_SLIDESHOW.TXT"
    
    Shell "C:\Program Files\IrfanView\i_view32.exe /killmesoftly"
    Shell "C:\Program Files\IrfanView\i_view32.exe ""/slideshow=" + FILE1 + " /ONE /bf /silent /gray /RELOADONLOOP /QUITE""", vbMinimizedNoFocus


End Sub

Private Sub MNU_TRIP_A_TRON_Click()
MNU_TRIP_A_TRON_Click_VAR = True
End Sub

Private Sub MNU_VB_Click()
Dim FileSpec, TT As Long

FileSpec = App.Path + "\" + App.EXEName + ".vbp"

If Dir(FileSpec) = "" Then
    FileSpec = App.Path + "\" + Dir(App.Path + "\" + "*.VBP")
End If


If Dir(FileSpec) = "" Then
    MsgBox "DONT FIND --" + vbCrLf + FileSpec
    Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE", vbNormalFocus
    End

End If

'If IsIDE = False Then
'Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE """ + FileSpec + """", vbNormalFocus
'End
'End If

VB_APP_PATH = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
If Dir(VB_APP_PATH) <> "" Then VB_APP_PATH2 = VB_APP_PATH
VB_APP_PATH = "C:\Program Files (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
If Dir(VB_APP_PATH) <> "" Then VB_APP_PATH2 = VB_APP_PATH


Shell VB_APP_PATH2 + " """ + FileSpec + """", vbNormalFocus
End


End Sub


Private Sub Label5_Click()
'Label5
Label9.BackColor = Label3.BackColor
Label9.ForeColor = RGB(255, 255, 255)
Label5.BackColor = Label10.BackColor
Label5.ForeColor = 0
x1 = 1
X2 = 0
End Sub

Private Sub Label9_Click()
Label5.BackColor = Label3.BackColor
Label5.ForeColor = RGB(255, 255, 255)
Label9.BackColor = Label10.BackColor
Label9.ForeColor = 0
X2 = 1
x1 = 0
End Sub


Private Sub MNU_X_EDIT_Click()

Shell "C:\PROGRAM FILES\NOTEPAD++\NOTEPAD++.EXE C:\TEMP\IRFAN_SLIDESHOW_X_SCRIPT.TXT", vbNormalFocus

End Sub

Private Sub MNU_X_PROCESS_Click()
    
    Dim FF
    
    Label3.BackColor = Label15.BackColor
    Label3.ForeColor = Label15.ForeColor
    
'    FS.CopyFile "C:\TEMP\IRFAN_SLIDESHOW-BACKUP.txt", "C:\TEMP\IRFAN_SLIDESHOW.txt"

    Dim HH(30000)
    HHC = 0
    On Error Resume Next
    FF = FreeFile
    Open "C:\TEMP\IRFAN_SLIDESHOW_X_SCRIPT.TXT" For Input As #FF
        Do
            Line Input #FF, LL1
            If Trim(LL1) <> "" Then
                HHC = HHC + 1
                
                HH(HHC) = LL1
            End If
        Loop Until EOF(FF)
    Close #FF, FF1
    
    
    
    FS.CopyFile "C:\TEMP\IRFAN_SLIDESHOW.txt", DDATE + "ALL BACKUP.TXT"
    FS.CopyFile "C:\TEMP\IRFAN_SLIDESHOW.txt", "C:\TEMP\IRFAN_SLIDESHOW-A-Z.txt"
    
    FF1 = FreeFile
    Open "C:\TEMP\IRFAN_SLIDESHOW-A-Z.txt" For Binary As #FF1

    FF2 = FreeFile
    Open "C:\TEMP\IRFAN_SLIDESHOW.txt" For Output As #FF2
    
    Do
        DoEvents
        Line Input #FF, LL1
        
        If Trim(LL1) <> "" Then
            TAGX = False
            For R = 1 To HHC
                If InStr("*" + LL1, HH(R)) > 0 Then TAGX = True
            Next
            
            If TAGX = False Then
                Print #FF2, LL1
                COU = COU + 1
                Else
                A = A
            End If
        End If
        
    Loop Until EOF(FF)
    
'    MsgBox COU
    
    Close #FF, FF1, FF2
    
'    Kill "C:\TEMP\IRFAN_SLIDESHOWA-Z.txt"
    Call MNU_RSR_MODIFY
    
    
    
    Label3.BackColor = Label14.BackColor
    Label3.ForeColor = Label14.ForeColor
    
    Kill "C:\TEMP\IRFAN_SLIDESHOW-A-Z.txt"
    FS.CopyFile "C:\TEMP\IRFAN_SLIDESHOW.txt", DDATE
    
    

End Sub



Sub LABELFILEINI()

'Me.WindowState = vbMinimized
    Dim A2, L2, U
    
If Dir(DDATE) = "" Then
    Label5 = "C:\TEMP\IRFAN_SLIDESHOW.TXT --  NOT FOUND"
    Exit Sub
End If


FF1 = FreeFile
Open "C:\Program Files\IrfanView\i_view32.ini" For Input As #FF1
    Do
        Line Input #FF1, l
        If InStr("-" + l, "StartIndex=") > 0 Then
            A1 = Val(Mid(l, Len("StartIndex=") + 1))
        End If
        If EOF(FF1) Then A1 = "2"
    Loop Until A1 <> ""
Close #FF1
INIA1 = A1
Reset

A2 = ""

FF2 = FreeFile
Open DDATE For Input As #FF2
    Do
        If Not EOF(FF2) Then Line Input #FF2, L2
        U = U + 1
        If U = A1 Then Exit Do 'a2 = l
    Loop Until EOF(FF2) Or A2 <> ""
Close #FF2
If A2 = "" Then A2 = L2

'If L2 = "" And Dir(Text1, vbDirectory) = "" Then
'If L2 = "" Then
'    Label5 = "POINTER IN FILE OF INI NOT FOUND =" + A1
'Else
    Label5 = Trim(str(A1) + vbCrLf + Replace(L2, "\", "\ \")) + vbCrLf + vbCrLf '+ vbCrLf
'End If

If L2 = "" Then Label5 = "FILE SCRIPT EMPTY"


Call BUILD_LAYOUT

End Sub

Sub CHECK_CMD_IS_RUN_OR_LEFT()
    
'    If Dir(DDATE) = "" Then Exit Sub
    
    
    
    If OPD = 0 Or OPT = 0 Then
        If Dir("C:\TEMP\INFAR_DIR_BUILD_TIME_BEGIN.TXT") <> "" Then
          On Error Resume Next
            FF = FreeFile
            Open "C:\TEMP\INFAR_DIR_BUILD_TIME_BEGIN.TXT" For Input As #FF
                Line Input #FF, OPD
                Line Input #FF, OPT
            Close #FF
        End If
    End If
            
    On Error Resume Next
    MTIME = DateValue(OPD) + TimeValue(OPT)
    On Error GoTo 0
    
    TT = DateDiff("S", MTIME, Now)
    TT2 = "000"
    III = Format((TT / 60), ".00")
    LSet TT2 = Mid(III, InStr(III, "."))
    TT2 = Mid(III, 1, InStr(III, ".") - 1) + TT2
    TT2 = Replace(TT2, " ", "0")

    
    If OPD2 = 0 Or OPT2 = 0 Then
        If Dir("C:\TEMP\INFAR_DIR_BUILD_TIME_END.TXT") <> "" Then
            FF = FreeFile
            Open "C:\TEMP\INFAR_DIR_BUILD_TIME_END.TXT" For Input As #FF
                Line Input #FF, OPD2
                Line Input #FF, OPT2
            Close #FF
    
        STOPCLOCK = DateValue(OPD2) + TimeValue(OPT2)
        TT = DateDiff("S", MTIME, STOPCLOCK)
        TT3 = "000"
        III = Format((TT / 60), ".00")
        LSet TT3 = Mid(III, InStr(III, "."))
        TT3 = Mid(III, 1, InStr(III, ".") - 1) + TT3
        TT3 = Replace(TT3, " ", "0")
    
        End If
    End If
                
    Set F = FS.GETFILE("C:\TEMP\IRFAN_SLIDESHOW.txt")
    FBSIZE = "FB_SIZE = " + Format(F.Size, "###,###,###,##0")
        
    If STOPCLOCK > 0 Then
        ZLAB = " - MIN = " + TT2 + " FIN = " + TT3 + " - " + FBSIZE
    Else
        ZLAB = " - MIN = " + TT2 + " - " + FBSIZE
    End If
    
    Label3.Caption = "BEGIN -- INFAR -- " + Format(Now, "DD-MMM-YY") + ZLAB
    
    
    
    Call BUILD_LAYOUT


    
End Sub

Private Sub MNU_RSR_MODIFY()
    
    RSR = True
    Label3.Caption = "BEGIN -- INFAR"
    
    MTIME = Now
    
    If Dir(DDATE) = "" Then Exit Sub
    
    Call CHECK_CMD_IS_RUN_OR_LEFT
    

End Sub



Private Sub Timer1_Timer()

'If InStr(Label3.Caption, "BEGIN -- INFAR") > 0 Then Exit Sub
If FindWindow(vbNullString, "INFAR_BATCH") > 0 Then Exit Sub

'    TT = Now
    Call CHECK_CMD_IS_RUN_OR_LEFT
    
    If Dir("C:\TEMP\INFRA_WORK_COMPLETE.TXT") <> "" Then
    
        'KILL WITH UNLOAD ME
'        Kill "C:\TEMP\INFRA_WORK_COMPLETE.TXT"
'        STOPCLOCK = TT
        
        'YELLOW
        Label3.BackColor = Label14.BackColor
        Label3.ForeColor = Label14.ForeColor
        
        Call MNU_X_PROCESS_Click
        
        Call Label3_Click
        

        Call BUILD_LAYOUT

        If MNU_TRIP_A_TRON_Click_VAR = True Then
            'KILL WITH UNLOAD ME
            Kill "C:\TEMP\INFRA_WORK_COMPLETE.TXT"
            Call MNU_RSR_Click
            
        
        End If
        

    End If
            

    
End Sub

Private Sub Timer2_Timer()
xbegininsecond = xbegininsecond + 1

If xbegininsecond > 120 Then Call Label3_Click

End Sub

Private Sub Timer3_Timer()

    Timer3.Interval = 500
    Dim FF5
    On Error Resume Next

    CHECK_CMD_IS_RUN_OR_LEFT

    Set F = FS.GETFILE("C:\TEMP\IRFAN_SLIDESHOW.txt")
    FBSIZE = "FB_SIZE = " + Format(F.Size, "###,###,###,##0")
    If OFBSIZE = FBSIZE Then Exit Sub
    OFBSIZE = FBSIZE

    FF5 = FreeFile
    Open "C:\TEMP\IRFAN_SLIDESHOW.TXT" For Binary As #FF5
        Seek FF5, LOF(FF5) - 1000
        Do
            If Not (EOF(FF5)) Then Line Input #FF5, A2
            
            If A1 = A2 Then Exit Do
            A1 = A2
    
            
        Loop Until EOF(FF5)
    Close #FF5
    
    Label5 = Trim(str(INIA1) + vbCrLf + Replace(A1, "\", "\ \")) + vbCrLf
       
    Call BUILD_LAYOUT

End Sub
