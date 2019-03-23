VERSION 5.00
Begin VB.Form frm_SEND_TO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008080&
   Caption         =   "SEND TO SCRIPT SUB FOLDER FILES SETCOUNT DIR"
   ClientHeight    =   6495
   ClientLeft      =   165
   ClientTop       =   540
   ClientWidth     =   13830
   Icon            =   "Test.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   13830
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   6600
      TabIndex        =   48
      Top             =   276
      Visible         =   0   'False
      Width           =   552
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   7920
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7440
      Top             =   600
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   9000
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "COUNT FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   12
      Left            =   8040
      TabIndex        =   47
      Top             =   5760
      Width           =   5790
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H005ED1EA&
      Caption         =   "Label2"
      Height          =   495
      Left            =   10080
      TabIndex        =   46
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "COUNT FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   11
      Left            =   8040
      TabIndex        =   45
      Top             =   4800
      Width           =   5790
      WordWrap        =   -1  'True
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "COUNT FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   10
      Left            =   8040
      TabIndex        =   44
      Top             =   5280
      Width           =   5790
      WordWrap        =   -1  'True
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "COUNT FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   9
      Left            =   6240
      TabIndex        =   43
      Top             =   5640
      Width           =   5790
      WordWrap        =   -1  'True
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "COUNT FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   8
      Left            =   6240
      TabIndex        =   42
      Top             =   5160
      Width           =   5790
      WordWrap        =   -1  'True
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "COUNT FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   7
      Left            =   600
      TabIndex        =   41
      Top             =   4920
      Width           =   5790
      WordWrap        =   -1  'True
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "COUNT FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   6
      Left            =   600
      TabIndex        =   40
      Top             =   4440
      Width           =   5790
      WordWrap        =   -1  'True
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "COUNT FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   5
      Left            =   600
      TabIndex        =   39
      Top             =   3960
      Width           =   5790
      WordWrap        =   -1  'True
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "COUNT FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   600
      TabIndex        =   38
      Top             =   3480
      Width           =   5790
      WordWrap        =   -1  'True
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "COUNT DIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   600
      TabIndex        =   37
      Top             =   3000
      Width           =   5790
      WordWrap        =   -1  'True
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "SCAN THESE DRIVES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   600
      TabIndex        =   36
      Top             =   600
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAR GO FROM INI NUMERIC ONE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   7320
      TabIndex        =   35
      Top             =   3120
      Visible         =   0   'False
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAR EDIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7320
      TabIndex        =   34
      Top             =   4200
      Visible         =   0   'False
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAR GO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7320
      TabIndex        =   33
      Top             =   2520
      Visible         =   0   'False
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAR XSCRIPT PROCESS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7320
      TabIndex        =   32
      Top             =   2040
      Visible         =   0   'False
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "IRFAR SCAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7320
      TabIndex        =   31
      Top             =   1320
      Visible         =   0   'False
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "Label4"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5280
      TabIndex        =   30
      Top             =   5385
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3960
      TabIndex        =   29
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3480
      TabIndex        =   28
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3480
      TabIndex        =   27
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3720
      TabIndex        =   26
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3480
      TabIndex        =   25
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3720
      TabIndex        =   24
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3720
      TabIndex        =   23
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3960
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3600
      TabIndex        =   21
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3840
      TabIndex        =   20
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3840
      TabIndex        =   19
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   4080
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3840
      TabIndex        =   17
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   4080
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   4080
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   4440
      TabIndex        =   14
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3600
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3720
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3720
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3960
      TabIndex        =   10
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3720
      TabIndex        =   9
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3960
      TabIndex        =   8
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3960
      TabIndex        =   7
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   4080
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   3840
      TabIndex        =   5
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabDR 
      Alignment       =   2  'Center
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
      Left            =   4080
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "LOGG FOLDER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      Caption         =   "OBJECTIVE FILE ONLY DIR ONLY CRC OPTION "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   5700
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "BEGIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB ME"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_EXPLORER 
      Caption         =   "OPEN LOGG FOLDER"
   End
   Begin VB.Menu MNU_PULL 
      Caption         =   "MENU PULL                                      "
      NegotiatePosition=   1  'Left
      Visible         =   0   'False
   End
   Begin VB.Menu MNU_TEXTLOAD 
      Caption         =   "TEXT LOAD LOGG'S"
      Begin VB.Menu MNU_01 
         Caption         =   "MNU_01"
      End
      Begin VB.Menu MNU_02 
         Caption         =   "MNU_02"
      End
      Begin VB.Menu MNU_03 
         Caption         =   "MNU_03"
      End
      Begin VB.Menu MNU_04 
         Caption         =   "MNU_04"
      End
   End
   Begin VB.Menu MNU_VOID 
      Caption         =   "OTHER MENU"
      Visible         =   0   'False
      Begin VB.Menu MNU_000_SCAN_GIVE 
         Caption         =   "000 --- SCAN SET TO GIVE"
      End
      Begin VB.Menu MNU_DIRECTORY_ONLY_01 
         Caption         =   "001 --- DIRECTORY ONLY"
      End
      Begin VB.Menu MNU_FILE_ONLY_02 
         Caption         =   "002 --- FILE ONLY"
      End
      Begin VB.Menu MNU_CRC_03 
         Caption         =   "003 --- CRC OPTION"
      End
      Begin VB.Menu MNU_CRC_NOT_PEND_DRIVE_04 
         Caption         =   "004 --- CRC -- DETECT NOT PEN DRIVE OPTION"
      End
      Begin VB.Menu MNU_CRC_NOT_PEND_DRIVE 
         Caption         =   "004 --- CRC -- DETECT BELOW TEN MEGABYTE"
      End
      Begin VB.Menu MNU_ALL_DRIVE_SET_05 
         Caption         =   "007 --- SWING ALL DRIVE SET FILING SYSTEM"
      End
   End
   Begin VB.Menu MNU_ESCAPE 
      Caption         =   "ESCAPE TO EXIT"
   End
End
Attribute VB_Name = "frm_SEND_TO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Routine_Set_Drive_Label
'Main_Routine

'----------------------------------------------------
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
'----------------------------------------------------
'----------------------------------------------------



Public VARCENTER
Dim FFILEAR(), WORKFLAG
Dim OBJECTIVE_BRIEF
Dim i
Dim i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16

Public LOGGFOLDER

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Dim OFILESPEC
Dim HHAR01()
Dim HHAR02()
Public FILE11, FILE12
Dim DAT1(), AT$
Dim QQ1 As Long
Private Declare Function TouchFileTimes Lib "imagehlp.dll" (ByVal FileHandle As Long, ByRef pSystemTime As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As BOOL
Enum BOOL
   FALSE_ = 0
   TRUE_ = 1
End Enum
'Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Private Declare Function TouchFileTimes Lib "imagehlp.dll" (ByVal FileHandle As Long, ByRef pSystemTime As Any) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
'Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'right order

'Public Type FILETIME
'    dwLowDateTime     As Long
'    dwHighDateTime    As Long
'End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes  As Long
    ftCreationTime    As FILETIME
    ftLastAccessTime  As FILETIME
    ftLastWriteTime   As FILETIME
    nFileSizeHigh     As Long
    nFileSizeLow      As Long
    dwReserved0       As Long
    dwReserved1       As Long
    cFileName         As String * MAX_PATH
    cAlternate        As String * 14
End Type

'----------------------------------------------------
'I PUT THIS HERE FOR THE SHELL EXECUTE LOADER PROGRAM
'----------------------------------------------------
'Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
'Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
'----------------------------------------------------
'----------------------------------------------------





Private Sub SET_VAR_FS()

Set FS = CreateObject("Scripting.FileSystemObject")

End Sub


Private Sub Main_Routine()
    
    Dim MCHAR
    
    On Error Resume Next
    
'    Dim DC
'    Set DC = FS.Drives
'    For Each D In DC
'        If D.DriveLetter <> "C" Then
'            s = s & D.DriveLetter
'
'            If D.DriveType = 3 Then
'                'Stop
'                'n = d.ShareName
'            ElseIf D.IsReady Then
'                N = N + D.DriveLetter 'd.VolumeName
'            End If
'                's = s & n & "<BR>"
'        End If
'    Next

    N = GET_DRIVES

'    If InStr("*" + LCase(n), LCase(Control.Caption)) > 0 Then
'        Control.BackColor = Label4.BackColor
'    End If
    
    
    For Each Control In LabDR
        OBJECTIVE_BRIEF = "ALL DRIVE"
        MCHAR = Mid(Control.Caption, 2, 1) + ":\"
        Err.Clear
        
        'If Control.BackColor <> 12648384 And Control.BackColor <> 8454143 Then
        
        N = GET_DRIVES
        If GetComputerName = "8-MSI-GP62M-7RD" Then
            N = Replace(N, "C", "") ' ---- HARDCODER -- DON'T DO C DRIVE AGAIN HERE
        End If
        If GetComputerName = "7-ASUS-GL522VW" Then
            N = Replace(N, "C", "") ' ---- HARDCODER -- DON'T DO C DRIVE AGAIN HERE
            N = Replace(N, "D", "") ' ---- HARDCODER -- DON'T DO C DRIVE AGAIN HERE
            N = Replace(N, "E", "") ' ---- HARDCODER -- DON'T DO C DRIVE AGAIN HERE
            N = Replace(N, "F", "") ' ---- HARDCODER -- DON'T DO C DRIVE AGAIN HERE
            N = Replace(N, "G", "") ' ---- HARDCODER -- DON'T DO C DRIVE AGAIN HERE
            N = Replace(N, "H", "") ' ---- HARDCODER -- DON'T DO C DRIVE AGAIN HERE
            N = Replace(N, "I", "") ' ---- HARDCODER -- DON'T DO C DRIVE AGAIN HERE
            N = Replace(N, "J", "") ' ---- HARDCODER -- DON'T DO C DRIVE AGAIN HERE
            N = Replace(N, "N", "") ' ---- HARDCODER -- DON'T DO C DRIVE AGAIN HERE
            N = Replace(N, "T", "") ' ---- HARDCODER -- DON'T DO C DRIVE AGAIN HERE
        End If
        
        If InStr("*" + LCase(N), LCase(Mid(Control.Caption, 2, 1))) = 0 And Control.BackColor <> 8421504 Then
            Control.BackColor = vbRed 'Label4.BackColor
        End If

        
        If InStr("*" + LCase(N), LCase(Mid(Control.Caption, 2, 1))) > 0 Then
            
            'Set F = FS.GetDrive(MCHAR)
            If Err.Number = 0 Then
                'If F.IsReady = True Then
                    'MCHAR = "E:\" 'TEST
                    ScanPath.TxtPath.Text = MCHAR
                    AT$ = MCHAR
                    'ScanPath.Timer1.Enabled = True
                    
                    If mCancelScan2 = True Then Exit For
                    
                    Call Label1_Click
                    
                    'i7 = DateDiff("n", PROCESSBEGIN, Now)
                    LastDriveTime = MCHAR + " -- " + Trim(Str(DateDiff("n", PROCESSBEGIN2, Now))) + " Min -- " + Trim(Str(DateDiff("s", PROCESSBEGIN2, Now) Mod 60)) + " Sec"
                    
                    Call Timer1_Timer
                    Call Timer2_Timer
                    If ScanPath.Enabled = True Then Call ScanPath.Timer1_Timer
                    
                    'ScanPath.cmdScan_Click
                    'Exit For
                    If mCancelScan2 = True Then Exit For
                'End If
            End If
        End If
    
    Next
    
    
    'Shell "C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe """ + FFILEAR(4) + """", vbNormalFocus
    
    'End

    
End Sub



Private Sub Form_Activate()



'File1.Path = "D:\0 00 LOGGERS TEXT\ARCHIVE OF DIRECTORY SCRIPT\SCRIPT LIST ALL DRIVES\2018-11-04 21-26-28 -- 8-MSI-GP62M-7RD"
'For R1 = 0 To File1.ListCount - 1
'
'    Filename = File1.Path + "\" + File1.List(R1)
'    FF01 = FreeFile
'    Open Filename + ".TMP" For Output As #FF01
'    FF02 = FreeFile
'    Open Filename For Input As #FF02
'    Do
'        Line Input #FF02, LINE_STRING
'        LINE_STRING = Trim(LINE_STRING)
'        Print #FF01, LINE_STRING
'    Loop Until EOF(FF02)
'    Close #FF01
'    Close #FF02
'    Kill Filename
'    Name Filename + ".TMP" As Filename
'Next
'
'End


Call SET_VAR_FS

'Call GIVE_DRIVE_COMMAND_STRING

'Me.Show
Me.Hide

LOGGFOLDER = "D:\0 00 LOGGERS TEXT\ARCHIVE OF DIRECTORY SCRIPT\SCRIPT LIST ALL DRIVES\" + Format(Now, "YYYY-MM-DD HH-MM-SS") + " -- " + GetComputerName + "\"

If Dir("D:\0 00 LOGGERS TEXT\", vbDirectory) = "" Then
    MkDir "D:\0 00 LOGGERS TEXT"
End If
If Dir("D:\0 00 LOGGERS TEXT\ARCHIVE OF DIRECTORY SCRIPT\", vbDirectory) = "" Then
    MkDir "D:\0 00 LOGGERS TEXT\ARCHIVE OF DIRECTORY SCRIPT\"
End If
If Dir("D:\0 00 LOGGERS TEXT\ARCHIVE OF DIRECTORY SCRIPT\SCRIPT LIST ALL DRIVES\", vbDirectory) = "" Then
    MkDir "D:\0 00 LOGGERS TEXT\ARCHIVE OF DIRECTORY SCRIPT\SCRIPT LIST ALL DRIVES\"
End If
If Dir(LOGGFOLDER, vbDirectory) = "" Then
    MkDir LOGGFOLDER
End If


LabelX(1).Caption = LOGGFOLDER


'FILE13 = "C:\TEMP\IRFAN_SLIDESHOW.txt"
'Set F = FS.GETFILE(FILE13)


'Label3.Caption = "SCAN TO GIVE -- " + AT$
'Label5.Caption = "IRFAR -- " + FILE13 + " -- " + Str(F.Size)


'Label2.FontSize = 12
'Label2.WordWrap = False
'Label2.AutoSize = True

LabelX(2) = "OBJECTIVE ---------- DRIVE SITEMAP -------" + vbCrLf
'LabelX(2) = LabelX(2) + "------------------" + vbCrLf
LabelX(2) = LabelX(2) + "GIVE -- FILE SCRIPT AS BOTH COMPLEX AND NORMAL AND BOTH FOLDER AND FILE SCRIPT" + vbCrLf
'LabelX(2) = LabelX(2) + "------- SITEMAP OF YOUR DRIVES" + vbCrLf
'Label2 = Label2 + "GIVE -- FILE ONLY OPTION" + vbCrLf
'Label2 = Label2 + "GIVE -- DIRECTORY ONLY OPTION" + vbCrLf
'Label2 = Label2 + "GIVE -- LATER -- CRC OPTION" + vbCrLf
'Label2 = Label2 + "GIVE 2012-OCT-21 -- SEND TO NOW HAS WORKING WITH IRFAR" + vbCrLf
'Label2 = Label2 + "       AND XSCRIPT DONT PROCESS WITH IT" + vbCrLf
'Label2 = Label2 + "OPTION TO GIVE DRIVE WHOLE SELECTOR TO ROOT DOWN BY FILE SEND-TO" + vbCrLf
'
'Label2 = Label2 + "------------------" + vbCrLf
'Label2 = Label2 + "PRESS BUTTON AIM HERE - BEGIN AS OPTION NUMBER ONE" + vbCrLf
LabelX(2) = LabelX(2) + "IF YOU NOTICE THE PATH AND FILE PATH ARE OUT OF SYNC IT IS BECAUSE THE SCAN IS IS NATRUAL ORDER - NOT SORTED" + vbCrLf
LabelX(2) = LabelX(2) + "WOULD HAVE TO LOAD INTO MEMORY TO SORT" + vbCrLf
LabelX(2) = LabelX(2) + "------------------"

'Label2.BackColor = QBColor(14)
'MsgBox Label2


'1 of 2 SET AS DRIVES

'Me.Top = Screen.Height / 2 - Me.Height / 2 '+400
'Me.Left = Screen.Width / 2 - Me.Width / 2 '+400



'XALL THE DRIVES HEIGHT WILL BE SET TO FIRST ONE
'LabDR(1).Height = 500

LabelX(1).BackColor = Label2.BackColor 'RGB(100, 200, 200)

LabelX(0).Caption = "SCAN ALL THESE DRIVES - HITT TO GO - AUTO GO" 'PRESS A KEY OR CLICK"
LabelX(0).FontSize = 20
LabelX(1).FontSize = 12
LabelX(2).FontSize = 12
'LabelX(9).Visible = False
'LabelX(9).FontSize = 15
LabelX(4).BackColor = Label2.BackColor
LabelX(3).BackColor = Label2.BackColor

LabelX(7).BackColor = Label2.BackColor
LabelX(8).BackColor = Label2.BackColor


LabelX(9).FontSize = 15
LabelX(10).FontSize = 15

LabelX(11).FontSize = 11
LabelX(12).FontSize = 11
LabelX(11).BackColor = Label2.BackColor 'RGB(100, 200, 200)
LabelX(12).BackColor = Label2.BackColor 'RGB(100, 200, 200)

'------------------------
'Set the Drives Positions
'They take the size of the form
'GOT TO SET THE HEIGHT AND MOVE TO THE TOP
Call A12
'------------------------

'For Each Control In Controls
'    If Mid(Control.Name, 1, 5) = "Label" Then
'        Control.AutoSize = False
'    End If
'Next
Dim XXWidth, YYHeight

For Each Control In Controls
    If Mid(Control.Name, 1, 6) = "LabelX" Then
        Control.Left = 0
        Control.AutoSize = True
        'Control.FontSize = 12
        Control.Refresh
        Control.WordWrap = False
'        Control.Width = Me.Width
        Control.Refresh
        Control.AutoSize = False
        If Control.Width > XXWidth Then XXWidth = Control.Width + Control.Left
    End If
Next

Me.Width = XXWidth

'------------------------
'Set the Drives Positions
'They take the size of the form
'Call A12
'------------------------

'For Each Control In Controls
'    If (Mid(Control.Name, 1, 5) = "Label" Or _
'        Control.Name = "LabDR") And _
'            Control.Visible Then
'                If Control.Width > XXWidth Then XXWidth = Control.Width + Control.Left
'    End If
'Next


'Me.Width = XXWidth + 120


'LabelX(1).Caption = ""
'LabelX(2).Caption = ""


HL = LabDR(1).Height + LabDR(1).Top
For Each Control In LabelX
    If Control.Visible = True Then
        Control.Top = HL
        HL = Control.Height + Control.Top + 40
            'Control.WordWrap = TRUE
    End If
Next



'Label10.Top = LabDR(1).Height + LabDR(1).Top
'Label3.Top = Label10.Height + Label10.Top + 20
'Label2.Top = Label3.Height + Label3.Top + 20



For Each Control In Controls
    If LCase(Mid(Control.Name, 1, 3)) <> "mnu" Then
        If LCase(Mid(Control.Name, 1, 5)) <> "timer" Then
            If Control.Visible = True Then
                If Control.Width > XXWidth Then XXWidth = Control.Width + Control.Left
                If Control.Height > YYHeight Then YYHeight = Control.Height + Control.Top
            End If
        End If
    End If
Next

For Each Control In Controls
    If Control.Name = "LabelX" Then
        Control.Width = XXWidth + 150
    End If
Next

'Me.Show

Me.Width = XXWidth + 250 '240
Me.Height = YYHeight + 880

'------------------------
'Set the Drives Positions
'They take the size of the form
Call A12
'------------------------

VARCENTER = False

'Me.Show
Me.Show
ScanPath.Timer1.Enabled = True

Call Main_Routine


If mCancelScan2 = True Then Unload ScanPath: Unload Me: Exit Sub

LabelX(0).Caption = "SCAN ALL THESE DRIVES DONE -- DRIVE SITEMAP FINISHED -- @ -- " + Format(Now, "DDD DD-MMM-YYYY HH:MM:SS")
LabelX(0).BackColor = RGB(120, 220, 120)
'LabelX(0).ForeColor = RGB(255, 255, 255)
LabelX(0).FontSize = LabelX(0).FontSize - 5

If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
Beep


If mCancelScan2 = True Then Unload ScanPath: Unload Me: Exit Sub

ScanPath.Timer1.Enabled = False
Timer1.Enabled = False
Timer2.Enabled = False


Exit Sub



'If Command$ <> "" Then MNU_IRFAR_02_Click
'Call Test4
Exit Sub


End
'Call Test2
'Exit Sub
Call Test1
End



Exit Sub


'1 pound = 0.0714285714 stone
'    More about calculator.

'1 pound = 0.0714285714 stone

stone = 112 * 0.0714285714
stone = 212 * 0.0714285714
Debug.Print stone
'= 7.9999999968
'= 15.1428571368
 
 
 'More than a year ago, in July 2007, International Space Station astronauts threw
 'an obsolete, refrigerator-sized ammonia reservoir overboard. Ever since,
 'the 1400-lb piece of space junk has been circling Earth in a decaying orbit--and now
 'it is about to reenter.
 
stone = 1400 * 0.0714285714
Debug.Print stone

 
 
 
' That 's cheating! You still don't know HOW it's done.

'256 ^ 0 = 1
'256 ^ 1 = 256
'256 ^ 2 = 65,536
'256 ^ 3 = 16,777,216

'So, 194.247.44.146 is ...

'146 * 256^0 = 146
'44 * 256^1 = 11,264
'247*256^2 = 16,187,392
'194*256^3 = 3,254,779,904
'
'Add those up, and ...

'146 + 11,264 + 16,187,392 + 3,254,779,904 = 3270978706
'
'3410960384  3410964479  UK  UNITED KINGDOM

Dim a0
Dim A1
a0 = 3270978706#
A1 = a0 / (256 ^ 3)
a2 = A1 - Int(A1) / 256 ^ 2
a3 = 3270978706# / 256 ^ 1
a4 = 3270978706# / 256 ^ 0

 
'203.79.32.0
'203.79.47.255
 
 
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_Load()

If App.PrevInstance = True Then End

'Me.Show

 
End Sub



Private Sub Form_Resize()

If VARCENTER = True Then Exit Sub
If Me.WindowState <> vbNormal And Me.WindowState <> vbmaximised Then Exit Sub

'Me.Top = Screen.Height / 2 - Me.Height / 2 '+400
Me.Top = 100
Me.Left = Screen.Width / 2 - Me.Width / 2 - 500


VARCENTER = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
'If ScanPath.Enabled = True Then SP.StopScan
'On Error GoTo 0

mCancelScan2 = True

Unload ScanPath

Dim Form As Form
For Each Form In Forms
    If Form.Enabled = True Then Unload Form
Next

'End
End Sub


Sub Test2()

Dim FS, Idate As Date, Hdate As String, File1 As String, FILE2 As String
Set FS = CreateObject("Scripting.FileSystemObject")

Call GetFileDates

File1 = Mid$(List1.List(List1.ListCount - 1), 21)
For R = List1.ListCount - 2 To 0 Step -1
FILE2 = Mid$(List1.List(R), 21)
If Mid$(List1.List(List1.ListCount - 1), 1, 19) <> Mid$(List1.List(R), 1, 19) Then
FS.CopyFile File1, FILE2
TT = LastModifiedToCurrent(FILE2)
End If
Next

Call GetFileDates

End Sub

Public Sub GetFileDates()
Dim FS, Idate As Date, Hdate As String, File1 As String, FILE2 As String
Set FS = CreateObject("Scripting.FileSystemObject")
List1.Clear
File1 = "D:\VB6\VB-NT\00_Best_VB_01\Banging_Tunes\BangList Total Logg.html"
Call Date_File(File1, Idate, Hdate)
'dd = Hdate + " " + file1
List1.AddItem Hdate + " " + File1

For R = 3 To 25
On Error Resume Next
Err.Clear
Set g = FS.GetDrive(FS.GetDriveName(FS.GetAbsolutePathName(Chr$(R + 64) + ":")))
On Error GoTo 0
If InStr(RR, g.VolumeName) = 0 And Err.Number = 0 Then
RR = RR + g.VolumeName + "-"
If Mid(g.VolumeName, Len(g.VolumeName) - 1, 2) = "GB" Then

FILE2 = Chr$(R + 64) + ":\Banging\BangList Total Logg.html"
Call Date_File(FILE2, Idate, Hdate)
List1.AddItem Hdate + " " + FILE2

'fs.copyFile file1, file2


End If
'Nuke4 = Nuke4 + g.totalsize / 1024 ^ 3
'Nuke5 = Nuke5 + g.freespace / 1024 ^ 3
End If
Next

End Sub

Public Sub Date_File(Filespec2, Idate As Date, Hdate As String)
'Call Date_File(filespec2$, Idate)

Dim FS, F
Set FS = CreateObject("Scripting.FileSystemObject")

pdate = 0
FILESPEC = Filespec2
Idate = DateSerial(1980, 1, 1)
If FS.FileExists(FILESPEC) Then
Set F = FS.GETFILE(FILESPEC)
Idate = F.DateLastModified
End If
Hdate = Format$(Idate, "YYYY-MM-DD HH:MM:SS")


End Sub


Public Function FindFileSize(Filename As String) As Long
        On Error Resume Next
        Dim filedata As WIN32_FIND_DATA

        filedata = Findfile(Filename)
        
        If filedata.nFileSizeHigh = 0 Then
            FindFileSize = Format$(filedata.nFileSizeLow, "###,###,###")
        Else
            FindFileSize = Format$(filedata.nFileSizeHigh, "###,###,###")
        End If
End Function



Function LastModifiedToCurrent(duFile As String)
    Dim lngHandle As Long
    lngHandle = CreateFile(duFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    If TouchFileTimes(lngHandle, ByVal 0&) = 0 Then
        'MsgBox "Error while changing file dates. Access Denied." + vbCrLf + duFile, vbCritical, "Modify Error"
    LastModifiedToCurrent = False
    Else
        'MsgBox "File date changed successfully.", vbInformation, "Modify Successful"
    LastModifiedToCurrent = True
    End If
    CloseHandle lngHandle
End Function

Sub Test1()

'restore header on wave fle
Open "m:\DWR 2008-11-27 00-36-03 -- 2008-11-28 13-02-14 -#- Janice Paul Short.wav" For Binary As #1
ll$ = "TTTT"
Get #1, 5, ll$
QQ1 = LOF(1)
Close #1
A1 = Asc(Mid$(ll, 1, 1))
a2 = Asc(Mid$(ll, 2, 1))
a3 = Asc(Mid$(ll, 3, 1))
a4 = Asc(Mid$(ll, 4, 1))
Mid(ll, 1, 1) = Chr(&H78)
Mid(ll, 2, 1) = Chr(&H10)
Mid(ll, 3, 1) = Chr(&H83)
Mid(ll, 4, 1) = Chr(&H23)

Mid(ll, 1, 1) = Chr(&H78)
Mid(ll, 2, 1) = Chr(&H94)
Mid(ll, 3, 1) = Chr(&H45)
Mid(ll, 4, 1) = Chr(&H0)

Open "M:\05 Media\Digital Wave Player M Drive\DR2008\DWR 2008-11-26 07-34-18 -- 2008-11-26 07-41-10 -#- New Anita Dingo Missed Steve On Linda On.wav" For Binary As #1
qq2 = LOF(1)
Close #1
Open "M:\05 Media\Digital Wave Player M Drive\DR2008\DWR 2008-11-26 09-39-09 -- 2008-11-27 00-35-50 -#- Steve Ibrag Linda Itching For Attack But Lets Clearly See Who Is Attacking First.wav" For Binary As #1
'qq2 = LOF(1)
jj$ = Space$(&H86)
Get #1, 1, jj$
Close #1

HH = &H78108323
HH = &H457894
gg1 = Hex$(QQ1 - 8)
gg2 = Hex$(qq2 - 8)
'4,560,000

Mid(ll, 1, 1) = Chr(&H78)
Mid(ll, 2, 1) = Chr(&H94)
Mid(ll, 3, 1) = Chr(&H45)
Mid(ll, 4, 1) = Chr(&H0)

Mid(ll, 1, 1) = Chr(&H78)
Mid(ll, 2, 1) = Chr(&H60)
Mid(ll, 3, 1) = Chr(&H8C)
Mid(ll, 4, 1) = Chr(&H56)

'QQ1 = QQ1 * 1.03025 '37.26,11 - correct from time files

'QQ1 = QQ1 * 1.032

QQ1 = QQ1 - 29000 '* 1.0000001 '36:20:11 - correct run lenght

gg3 = Hex$(Int(QQ1))
ll = ""
lDataSize = Len(gg3)
Dim i As Integer
i2 = 1
For i = lDataSize - 1 To 1 Step -2
'For i = 1 To lDataSize Step 2
    'bytearray(i2) = Val("&h" + Mid$(gg3, i, 2))
    ll = ll + Chr(Val("&h" + Mid$(gg3, i, 2)))
    i2 = i2 + 1
Next



'T1 = DateSerial(bytearray(1), bytearray(2), bytearray(3),bytearray(4))

t1 = DateDiff("h", DateSerial(2008, 11, 27) + TimeSerial(0, 36, 3), DateSerial(2008, 11, 28) + TimeSerial(13, 2, 14))
t2 = DateDiff("n", DateSerial(2008, 11, 27) + TimeSerial(0, 36, 3), DateSerial(2008, 11, 28) + TimeSerial(13, 2, 14)) Mod 60
t3 = DateDiff("s", DateSerial(2008, 11, 27) + TimeSerial(0, 36, 3), DateSerial(2008, 11, 28) + TimeSerial(13, 2, 14)) Mod 60

A1 = Asc(Mid$(ll, 1, 1))
'a2 = Asc(Mid$(ll, 2, 1))
'a3 = Asc(Mid$(ll, 3, 1))
'a4 = Asc(Mid$(ll, 4, 1))
getb$ = " "
Open "m:\DWR 2008-11-27 00-36-03 -- 2008-11-28 13-02-14 -#- Janice Paul Short.wav" For Binary As #1
'll$ = "TTTT"
Put #1, 1, jj$
Put #1, 5, ll
Put #1, &H7D, ll
'Get #1, LOF(1), getb$
'Put #1, LOF(1), Chr$(0)
Close #1

A = Hex$(Asc(getb$))
A = Len(ll)

Shell "C:\Program Files\Windows Media Player\mplayer2.exe ""m:\DWR 2008-11-27 00-36-03 -- 2008-11-28 13-02-14 -#- Janice Paul Short.wav""", vbNormalFocus

End
End Sub

Sub Test3()
'Form1.Hide
'ScanPath.Show
'ScanPath.TxtPath.Text = "M:\04 Music ---\04 My Music\01 Banging Tunes"
'ScanPath.TxtPath.Text = "M:\04 Music ---\04 My Music\01 Banging Tunes"
'ScanPath.cmdScanDir_Click
'
'For WE = 1 To ScanPath.ListView2.ListItems.Count
'    A1$ = ScanPath.ListView2.ListItems.Item(WE).SubItems(1)
'    B1$ = ScanPath.ListView2.ListItems.Item(WE)
'    A1$ = A1$ + B1$
'
'    'ets1 = Len("D:\# MY DOCS\# 01 My Documents\Favorites")
'
'    et = InStr(A1$, "\200")
'    If et > 0 Then
'    If et = Len(A1$) - 4 Then
'    ef = InStrRev(A1$, "\", et - 1)
'
'    'If Len(a1$) > Len(ScanPath.txtPath.Text) + 1 And InStr(Len(ScanPath.txtPath.Text) + 2, a1$, "\") > 0 Then
'
'    tg1 = Mid$(A1$, ef, et - ef)
'    tg2 = Mid$(A1$, et + 1, 4)
'    tg3 = Mid$(A1$, 1, ef) + Mid$(tg1, 2, 7) + tg2 + " - " + Mid$(tg1, 16)
'    'ef = InStr(et + 6, a1$, "\")
'    'If ef = 0 Then
'    TY = Mid$(A1$, 1, et - 1)
'    'tx = Mid$(a1$, 1, Len(ScanPath.txtPath.Text) + 12) + " " + tg1 + " - " + Mid$(a1$, Len(ScanPath.txtPath.Text) + 14)
'    'ty = a1$ + "\" + Mid$(a1$,et+5 , 5)
'    On Error GoTo 0
''    MkDir ty
''    Name ty As tg3
''    MkDir tg3
'    'End If
'    Set FS = CreateObject("Scripting.FileSystemObject")
'    'ScanPath.txtPath.Text = a1$
'    'ScanPath.cmdScan_Click
'
'    Set F = FS.GetFolder(A1$)
'        Set fc = F.Files
'        For Each F1 In fc
'            's = s & f1.Name
'            's = s & vbCrLf
'    '    Next
'
'
'    'For we2 = 1 To ScanPath.ListView1.ListItems.Count
'    '    a1$ = ScanPath.ListView1.ListItems.Item(we2).SubItems(1)
'    '    b1$ = ScanPath.ListView1.ListItems.Item(we2)
'
'
'    Set F1 = FS.GETFILE(A1$ + "\" + F1.Name)
'    F1.Move tg3 + "\" + F1.Name
'    Next
'    'F.Move tg3
'    Set fc = F.SubFolders
'        For Each F1 In fc
'            Set F1 = FS.GetFolder(A1$ + "\" + F1.Name)
'            F1.Move tg3 + "\" + F1.Name
'
'        Next
'
'    RmDir A1$
'
'    End If
'    End If
'Next

End Sub



Sub Test4()

Exit Sub
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




Sub Label0_Click()


End Sub


Sub GIVE_DRIVE_COMMAND_STRING()

Exit Sub


If AT$ = "" Then
        AT$ = Command$
    Else
        OBJECTIVE_BRIEF = "SINGLE DRIVE SELECT VIA BUTTON"
        Exit Sub
End If



If IsIDE = True Then
    AT$ = "M:\#\#D\00 Pen Drives MOBILES"
'    OBJECTIVE_BRIEF = "IDE MODE SET  " + AT$

End If
'If IsIDE = True Then AT$ = "C:\TEMP\" + Dir("C:\TEMP\*.*")

If AT$ <> "" And Mid$(AT$, 1, 1) = """" Then
    AT$ = Mid$(AT$, 2)
    AT$ = Mid$(AT$, 1, Len(AT$) - 1)
    OBJECTIVE_BRIEF = "SEND TO CONTEXT MENU MODE FOLDER"

End If


On Error Resume Next

If FS.FileExists(AT$) Then
    AT$ = Mid(AT, 1, 1) + ":\"
    OBJECTIVE_BRIEF = "SEND TO CONTEXT MENU MODE -- ROOT DRIVE FROM -- FILE SELECT"

End If

End Sub

Private Sub LabDR_Click(Index As Integer)
Exit Sub
If LabDR(Index).BackColor = Label4.BackColor Then Exit Sub
AT$ = Chr(Index + 65) + ":\"

Label3.Caption = "SCAN TO GIVE -- " + AT$



End Sub


Public Sub NAUGHT_NULL(VAR1, VAR2, VAR3)

'Call SEND_TO.NAUGHT_NULL(i1, F.TotalSize, i2)
VAR1 = Replace(VAR1, "000,", "___,")
VAR1 = Replace(VAR1, " 00,", "___,")
VAR1 = Replace(VAR1, "  0,", "___,")

VAR1 = Replace(VAR1, "_,0", "_,_")
VAR1 = Replace(VAR1, "_,00", "_,__")
VAR1 = Replace(VAR1, "_,_0", "_,__")
VAR1 = Replace(VAR1, ",__0", ",___")

If VAR3 = -1 Then Exit Sub

For R = 1 To 8
    RULER = 1024 ^ R
    If R = 1 Then RULETEXT = "KILO RANGE"
    If R = 2 Then RULETEXT = "MEGA RANGE"
    If R = 3 Then RULETEXT = "GIGA RANGE"
    If R = 4 Then RULETEXT = "TERA - TETRA - TERRORIZOR - RANGE"
    If VAR2 >= RULER Then VAR3 = RULETEXT
Next


End Sub

Private Sub Label1_Click()
' --- BEGIN


'Dim AT$

'Set FS = CreateObject("Scripting.FileSystemObject")

'Call GIVE_DRIVE_COMMAND_STRING

'Set F = FS.GetDrive(Mid(AT$, 1)) '+ ":\")

Set F = FS.GetDrive(FS.GetDriveName(AT$))

'Me.Hide

'ScanPath.Show
'DoEvents
'ScanPath.SetFocus
'DoEvents
'ScanPath.Show
'DoEvents


ScanPath.TxtPath.Text = AT$

FILE5TXT = ScanPath.TxtPath.Text
FILE5TXT = Replace(FILE5TXT, "\", "_")
FILE5TXT = Replace(FILE5TXT, ":", "_")
FILE5TXT = FILE5TXT + " VOL - " + F.VolumeName

'Set F = FS.GetFolder(folderspec)
's = F.DateCreated

's = s & "Created: " & F.DateCreated & vbCrLf
's = s & "Last Accessed: " & F.DateLastAccessed & vbCrLf
's = s & "Last Modified: " & F.DateLastModified
'MsgBox s, 0, "File Access Info"

'DAT1 = F.DateCreated
'DAT1 = F.DateLastModified
'DAT1 = F.DateLastAccessed

If Len(ScanPath.TxtPath.Text) = 3 Then
    'SUBTEXTFOLDER = "#0 ACHIVE - SCRIPT FILE FOLDER AND SUB\"
    SUBTEXTFOLDER = LOGGFOLDER
    'TDIR = ScanPath.TxtPath.Text + SUBTEXTFOLDER
    TDIR = SUBTEXTFOLDER
    
    If Dir(TDIR, vbDirectory) = "" Then
        MkDir TDIR
    End If
    
'    TDIR = SUBTEXTFOLDER
End If

FILE11 = "FILE & FOLDER SCRIPT -- " + Format(Now, "YYYY_MM_DD_HH_MM_SS") + " -- COMPLEX -- " + FILE5TXT + ".txt"
FILE12 = "FILE & FOLDER SCRIPT -- " + Format(Now, "YYYY_MM_DD_HH_MM_SS") + " -- NORMAL ---- " + FILE5TXT + ".txt"
FILE13 = "FOLDER SCRIPT -- " + Format(Now, "YYYY_MM_DD_HH_MM_SS") + " -- COMPLEX -- " + FILE5TXT + ".txt"
FILE14 = "FOLDER SCRIPT -- " + Format(Now, "YYYY_MM_DD_HH_MM_SS") + " -- NORMAL ---- " + FILE5TXT + ".txt"


Dim ORGTAGFILE(4)
ORGTAGFILE(1) = FILE11
ORGTAGFILE(2) = FILE12
ORGTAGFILE(3) = FILE13
ORGTAGFILE(4) = FILE14



FILE21 = TDIR + FILE11  '" -- COMPLEX --  "
FILE22 = TDIR + FILE12  '" -- NORMAL ---- "
FILE23 = TDIR + FILE13  '" -- COMPLEX --  DIR"
FILE24 = TDIR + FILE14  '" -- NORMAL ---- DIR"

'ABOVE
'WORK DIM THE ARRAY HERE
'BELOW

ReDim FFILEAR(4)
FFILEAR(1) = FILE21
FFILEAR(2) = FILE22
FFILEAR(3) = FILE23
FFILEAR(4) = FILE24

For Each Control In Controls
If Mid(Control.Name, 1, 5) = "MNU_0" Then
    RT = Val(Mid(Control.Name, 5))
        Control.Caption = Mid(FFILEAR(RT), InStrRev(FFILEAR(RT), "\") + 1)
    End If
Next

FF01 = FreeFile
Open FILE21 For Output As #FF01
FF02 = FreeFile
Open FILE22 For Output As #FF02
FF03 = FreeFile
Open FILE23 For Output As #FF03
FF04 = FreeFile
Open FILE24 For Output As #FF04

Dim FFAR(4)
FFAR(1) = FF01
FFAR(2) = FF02
FFAR(3) = FF03
FFAR(4) = FF04



' -------------------------------------- DC SET
' -------------------------------------- DC SET
' -------------------------------------- DC SET
DC = 19
For R = 1 To 4
    Print #FFAR(R), "---------------------------------------------------------------------------"
    Print #FFAR(R), "SOURCE -- SOLUTION -- SOLVABILITY -- Drive SiteMap"
    Print #FFAR(R), "---------------------------------------------------------------------------"
    Print #FFAR(R), "PATH :-:" + TDIR
    Print #FFAR(R), "FILE :-:" + ORGTAGFILE(R)
    Print #FFAR(R), "---------------------------------------------------------------------------"
    Print #FFAR(R), "FOLDER DIRECTORY OBJECTIVE SCAN -- "
    Print #FFAR(R), "## " + AT$
    Print #FFAR(R), "---------------------------------------------------------------------------"
    
    For R1 = 0 To DC
        Print #FFAR(R), String(140, "#")
    Next
    Print #FFAR(R), "---------------------------------------------------------------------------"
    Print #FFAR(R), "<-------------------->"

Next

    PROCESSTIMEDIFF = Now

' ---------------------------------------------------------------
' ----------- EXECUTE BEGIN ------------------------
' ---------------------------------------------------------------


'ScanPath.Show


ScanPath.cmdScan_Click


For R = 1 To 4
    Close FFAR(R)
Next

Select Case F.DriveType
    Case 0: TDriveType = "Unknown"
    Case 1: TDriveType = "Removable"
    Case 2: TDriveType = "Fixed"
    Case 3: TDriveType = "Network"
    Case 4: TDriveType = "CD-ROM"
    Case 5: TDriveType = "RAM Disk"
End Select

'Set F1 = FS.GetDrive(Mid(AT$, 1)) '+ ":\")

Set F = FS.GetDrive(FS.GetDriveName(AT$))

i1 = Format(mDIRCount2, "000,000,000,000,000")
i2 = Format(mFILECount2, "000,000,000,000,000")
i3 = Format(COUNTPROCES2, "000,000,000,000,000")

'THIS DOES FORMATING TO THE NUMBER VAR WITH 1000 SEPERATOR
'LIKE 000,000,000,000,023
'LIKE ___,___,___,___,023

Call NAUGHT_NULL(i1, 0, 0)
Call NAUGHT_NULL(i2, 0, 0)
Call NAUGHT_NULL(i3, 0, 0)


'  --  ___,___,___,___,__1  - 3-3-3-3 = 12 DIGITS = AROUND 20 GIGBYTE
'DC IS FIXED ABOVE --- DC = 20
ReDim DAT1(40)
DC = -1
DC = DC + 1: DAT1(DC) = "DIR:                  " + i1
DC = DC + 1: DAT1(DC) = "FILE:                 " + i2
DC = DC + 1: DAT1(DC) = "PROCESS:              " + i3

DC = DC + 1: DAT1(DC) = "---------------------------------------------------------------------------"
DC = DC + 1: DAT1(DC) = "Volume Name:          " + F.VolumeName
DC = DC + 1: DAT1(DC) = "Drive Type:           " & UCase(F.DriveType)
DC = DC + 1: DAT1(DC) = "Drive Type:           " & UCase(TDriveType)  ' --- ABOVE
DC = DC + 1: DAT1(DC) = "File System:          " & F.FileSystem
DC = DC + 1: DAT1(DC) = "Serial Number:        " & F.serialnumber
i1 = Hex$(F.serialnumber)
i1 = Mid(i1, 1, 2) + "," + Mid(i1, 3, 2) + "," + Mid(i1, 5, 2) + "," + Mid(i1, 7, 2)
DC = DC + 1: DAT1(DC) = "Serial Number &H:     " & i1

i1 = F.ShareName
If i1 = "" Then i1 = "LOCAL SYSTEM DRIVE"
DC = DC + 1: DAT1(DC) = "Share Name:           " & i1

i1 = Format(F.TotalSize, "000,000,000,000,000")
Call NAUGHT_NULL(i1, F.TotalSize, i2)
DC = DC + 1: DAT1(DC) = "Total Size:           " & i1 + " - " + i2

i1 = Format(F.AvailableSpace, "000,000,000,000,000")
Call NAUGHT_NULL(i1, F.AvailableSpace, i2)
DC = DC + 1: DAT1(DC) = "Available:            " & i1 + " - " + i2

i1 = Format(F.FreeSpace, "000,000,000,000,000")
Call NAUGHT_NULL(i1, F.FreeSpace, i2)
DC = DC + 1: DAT1(DC) = "Free Space:           " & i1 + " - " + i2

INOW = Now
i = Format(DateDiff("s", PROCESSTIMEDIFF, INOW) / 60, "0.0##")
DC = DC + 1: DAT1(DC) = "Process Lenght Time:  " + i + " Minute Divide Clock"
i = Trim(Str(DateDiff("s", PROCESSTIMEDIFF, INOW)))
DC = DC + 1: DAT1(DC) = "Process Lenght Time:  " + i + " -- Seconds"
DC = DC + 1: DAT1(DC) = "BEGIN TIME:           " + Format(PROCESSTIMEDIFF, "DD MMM YYYY HH:MM:SS") + "h"
DC = DC + 1: DAT1(DC) = "AFTER TIME:           " + Format(INOW, "DD MMM YYYY HH:MM:SS") + "h" + " -" + Str(i) + " Sec"
DC = DC + 1: DAT1(DC) = "FORMAT BEGIN TIME:    " + Format(PROCESSTIMEDIFF, "DDD DD MMM YYYY HH:MM:SSa/p")
If OBJECTIVE_BRIEF = "" Then OBJECTIVE_BRIEF = "NORMAL"
DC = DC + 1: DAT1(DC) = "OBJECTIVE MODE BRIEF: " + OBJECTIVE_BRIEF





'MsgBox DC


ReDim Preserve DAT1(DC)

For R1 = 1 To 4

    Dim CHUNK As String
    FF01 = FreeFile
    Open FFILEAR(R1) For Binary As #FF01
        CHUNK = Space(8000)
        If CHUNK > Len(FF01) Then
            CHUNK = Space(Len(FF01) - 2)
        End If
        
        Get #FF01, 1, CHUNK
        For R = 0 To UBound(DAT1)
            STRING1 = DAT1(R)
            UU0 = String$(140, "#")
            UU1 = String$(140, "#")
            Mid(UU1, 1, Len(STRING1)) = STRING1
            
            UU1 = Replace(UU1, "#", " ")
            CHUNK = Replace(CHUNK, UU0, UU1, , 1)
        Next
        Put #FF01, 1, CHUNK
    Close #FF01

Next



For R1 = 1 To 4
    FF01 = FreeFile
    Open FFILEAR(R1) + ".TMP" For Output As #FF01
    FF02 = FreeFile
    Open FFILEAR(R1) For Input As #FF02
    Do
        Line Input #FF02, LINE_STRING
        LINE_STRING = Trim(LINE_STRING)
        Print #FF01, LINE_STRING
    Loop Until EOF(FF02)
    Close #FF01
    Close #FF02
    Kill FFILEAR(R1)
    Name FFILEAR(R1) + ".TMP" As FFILEAR(R1)
Next




'FILE32 = "D:\0 00 LOGGERS TEXT"
'FILE31 = "D:\0 00 LOGGERS TEXT\ARCHIVE OF DIRECTORY SCRIPT"
'
'If Dir(FILE31, vbDirectory) = "" Then
'    MkDir FILE32
'    MkDir FILE31
'End If
'
'
'For R = 1 To 4
'    'EVEN ARE ' " -- COMPLEX ---- "
'    'ODD ARE ' " -- NORMAL ---- "
'    Set F = FS.GETFILE(FFILEAR(R))
'    F.Copy FILE31 + "\" + ORGTAGFILE(R)
'Next

'
'Shell "Explorer.exe /select, """ + FILE31 + "\" + ORGTAGFILE(2) + """", vbNormalFocus
'Shell "Explorer.exe /select, " + FFILEAR(2), vbNormalFocus


'If OBJECTIVE_BRIEF <> "ALL DRIVE" And IRFAR_TO_DO_FROM_SEND_TO = False Then
    
    If mCancelScan2 = True Then Exit Sub
    
    
    ' Shell "C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe """ + FFILEAR(4) + """", vbNormalFocus
    Me.WindowState = vbNormal
    Beep
    
    'End
'End If

End Sub

Private Sub Label10_Click()
IRFAR_TO_DO_FROM_SEND_TO = True
Label5_Click
End Sub

Private Sub Label2_Click()
'Label2
ScanPath.Timer1.Enabled = True
Call Label1_Click
ScanPath.Timer1.Enabled = False
ScanPath.lblCount4 = "WORK COMPLETE"
End Sub

Private Sub Label3_Click()

'Label3.Caption = "D:\0 00 LOGGERS TEXT\ARCHIVE OF DIRECTORY SCRIPT\SCRIPT LIST ALL DRIVES\"


'Me.WindowState = vbMinimized
'ScanPath.WindowState = vbNormal
'ScanPath.Show
'DoEvents
'ScanPath.Timer1.Enabled = True
'Call Label1_Click
'ScanPath.Timer1.Enabled = False
'ScanPath.lblCount4 = "WORK COMPLETE"
End Sub

Private Sub Label5_Click()
'Label5
'Me.WindowState = vbMinimized

MNU_IRFAR_01_Click
ScanPath.lblCount4 = "WORK COMPLETE"
ScanPath.Timer1.Enabled = False
End Sub

Private Sub Label6_Click()
    If Label6.BackColor = QBColor(14) Then Exit Sub
    
    QB = Label6.BackColor
    Label6.BackColor = QBColor(14)
    DoEvents
    FILE13 = "C:\TEMP\IRFAN_SLIDESHOW.txt"
    
'    Call MNU_INI_BACK_NAUGHT_Click
    
    Dim FF

    On Error Resume Next
    
'    ReDim HHAR01(30000)
'    HHC01 = 0
'    FF = FreeFile
'    Open "C:\TEMP\IRFAN_SLIDESHOW_X_SCRIPT.TXT" For Input As #FF
'        Do
'            Line Input #FF, LL1
'            If Trim(LL1) <> "" Then
'                HHC01 = HHC01 + 1
'                HHAR01(HHC01) = UCase(LL1)
'            End If
'        Loop Until EOF(FF)
'    Close #FF, FF1
'    ReDim Preserve HHAR01(HHC01)


XSCRIPT_CHUNK_AND_PROCESS
'XSCRIPTCHUNK
FILE13 = "C:\TEMP\IRFAN_SLIDESHOW.txt"
Set F = FS.GETFILE(FILE13)

Label5.Caption = "IRFAR -- " + FILE13 + " -- " + Str(F.Size)

Sleep 200
Label6.BackColor = QB
    
    
End Sub

Private Sub Label7_Click()
Me.WindowState = vbMinimized
MNU_IRFAR_02_Click
End Sub

Private Sub Label8_Click()
Me.WindowState = vbMinimized
MNU_IRFAR_03_NOTEPAD_Click
End Sub

Private Sub Label9_Click()
Me.WindowState = vbMinimized
Call MNU_INI_BACK_NAUGHT_Click
MNU_IRFAR_02_Click
End Sub

Private Sub MNU_000_SCAN_GIVE_Click()
MsgBox "NOT A OPTION HERE MENU UPDATE WITH MEDIA OBJECTIVE" + vbCrLf + "WHOLE DRIVE OR SUB DIRECTORY TO SCAN IS" + vbCrLf + AT$
End Sub



Private Sub MNU_01_Click()
    
Shell "C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe """ + FFILEAR(1) + """", vbNormalFocus

End Sub
Private Sub MNU_02_Click()

Shell "C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe """ + FFILEAR(2) + """", vbNormalFocus

End Sub
Private Sub MNU_03_Click()

Shell "C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe """ + FFILEAR(3) + """", vbNormalFocus

End Sub
Private Sub MNU_04_Click()

Shell "C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe """ + FFILEAR(4) + """", vbNormalFocus

End Sub


Private Sub MNU_ESCAPE_Click()
Beep
End Sub

Private Sub MNU_EXIT_Click()
Beep
Unload Me
End Sub

Private Sub MNU_EXPLORER_Click()
Beep
Shell "Explorer.exe " + LOGGFOLDER, vbNormalFocus

End Sub

Private Sub MNU_IRFAR_03_NOTEPAD_Click()

FILE13 = "C:\TEMP\IRFAN_SLIDESHOW.txt"

Shell "C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe """ + FILE13 + """", vbNormalNoFocus

Shell "C:\PROGRAM FILES\NOTEPAD++\NOTEPAD++.EXE " + FILE13, vbNormalNoFocus

End Sub

Private Sub MNU_IRFAR_04_Click()
If MNU_IRFAR_04.Checked = True Then MNU_IRFAR_04.Checked = False: Exit Sub
If MNU_IRFAR_04.Checked = False Then MNU_IRFAR_04.Checked = True
End Sub

Private Sub MNU_IRFAR_05_Click()
Shell "NOTEPAD ""C:\TEMP\IRFAN_SLIDESHOW_X_SCRIPT.TXT""", vbNormalFocus


End Sub

Private Sub MNU_VB_Click()
    
    Dim FILESPEC, TT As Long
    FILESPEC = App.Path + "\" + App.EXEName + ".vbp"

    If IsIDE = False And Dir(FILESPEC) <> "" Then
        On Error Resume Next
        If Dir("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE") <> "" Then
            TT = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE  """ + FILESPEC + """", vbMinimizedNoFocus)
        End If
        If Dir("C:\Program Files (X86)\Microsoft Visual Studio\VB98\VB6.EXE") <> "" Then
            TT = Shell("C:\Program Files (X86)\Microsoft Visual Studio\VB98\VB6.EXE  """ + FILESPEC + """", vbMinimizedNoFocus)
        End If
        
        If TT = 0 Then MsgBox "None Process PID Number Returned" + vbCrLf + "Must Mean Not in Admin Mode to Run Shell Command"
        
        'RUN EXIT OPTION
        'TT = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE /runexit """ + FileSpec + """", vbMinimizedNoFocus)
        End
    End If

End Sub

Function GET_DRIVES()

Dim DC
Set DC = FS.Drives
For Each D In DC
  s = s & D.DriveLetter
  If D.DriveType = 3 Then
    Stop
    'n = d.ShareName
  ElseIf D.IsReady Then
    N = N + D.DriveLetter 'd.VolumeName
  End If
  's = s & n & "<BR>"
Next

GET_DRIVES = N

End Function


Sub A12()

On Error Resume Next

XC = 64
For Each Control In LabDR
    XC = XC + 1
    'Control.Visible = FALSE
    Control.AutoSize = False
    Control.Caption = "-" + Chr(XC) + "-"
    Control.AutoSize = True
    Control.Refresh
    Control.AutoSize = False
Next

DoEvents



For Each Control In LabDR
    
    If Control.Index Mod 2 = 0 Then
        Control.BackColor = &HC0FFC0
    End If
    
    MCHAR = Mid(Control.Caption, 2, 1) + ":\"
    Err.Clear
    
    'Set F = FS.GetDrive(MCHAR)
            
    If InStr("*" + LCase(GET_DRIVES), LCase(Mid(Control.Caption, 2, 1))) = 0 Then
        Control.BackColor = Label4.BackColor
    End If



'    If F.IsReady = False Or Err.Number > 0 Then
'        Control.BackColor = Label4.BackColor
'    End If
    
    If Control.Width > XT Then XT = Control.Width
Next

XT = (Me.Width - 100) / 26
XD = 0
For Each Control In LabDR
    Control.Top = 0
    Control.Left = XD
    Control.Width = XT
    XD = XD + XT
    'Control.Height = 500
    Control.Visible = True
Next


End Sub

Private Sub MNU_IRFAR_01_Click()
    
    ' --- ANOTHER MENU SET VAR SELECT DRIVE SET ADD THEN DO
    ' --- MNU_IRFAR_MODE_SET = True
    ' --- FILTER OUT VAR GATHERED
    
    
    Call MNU_INI_BACK_NAUGHT_Click
    
    Dim FF

    On Error Resume Next
'    If IRFAR_TO_DO_FROM_SEND_TO = False Then
    XSCRIPTCHUNK
'    End If

    'ReDim Preserve HHAR01(HHC01)
    

    MNU_IRFAR_MODE_SET = True
    
    If Dir("D:\#0 IRFAR SHOW SCRIPT FILE SET", vbDirectory) = "" Then
        MkDir "D:\#0 IRFAR SHOW SCRIPT FILE SET"
    End If
    
    
    Me.Hide
    ScanPath.Show
    ScanPath.WindowState = vbNormal
    
    
i1 = ""
i2 = ""
  
  
ScanPath.cboMask.Text = "*.JPG;*.GIF"
ScanPath.cboMask.Text = "*.JPG;*.GIF"
ScanPath.cboMask.Text = "*.JPG;*.JPEG"
'ScanPath.cboMask.Text = "*.JPG;*.GIF;*.MPG;*.AVI;*.MP4;*.3GP"
'ScanPath.cboMask.Text = "*.MPG;*.AVI;*.MP4"
 
 
    i1 = i1 + "*.JPG|JPEG|JPE|GIF|BMP|DIB|TIF|TIFF|PNG|PCX|TGA|LBM|"
    If MNU_IRFAR_04.Checked = False Then
        i1 = i1 + "CLP|ANI|VOB|MPG|MP4|3GP|AVI"
    End If
'    i1 = "*.AVI|MP4|3GP"
    
    i1 = Replace(i1, "|", ";*.")
    ScanPath.cboMask.Text = i1

    'JPG|JPEG|JPE|GIF|BMP|DIB|TIF|TIFF|PNG|PCX|TGA|PCD|SUN|RAS|**ICO|AVI|IFF|PPM|PGM|PBM|LBM|WMF|EMF|WAV|PSD|MPG|MPE|MPEG|MID|RMI|MOV|CUR|ANI|DCX|EPS|CLP|CAM|G3|AIF|AU|SND|PSP|PSPIMAGE|ICL|SFW|KDC|RA|MP3|DCM|ACR|FPX|XBM|XPM|DJVU|SWF|FLV|IMG|IW44|WBMP|SGI|RGB|RLE|MED|RLE|SFF|NLM|NOL|NGG|GSM|JP2|JPC|J2K|JPF|FSH|CRW|B3D|WMA|WMV|TTF|SID|MNG|JNG|RAW|ECW|ASF|TXT|WAD|ICS|IDS|CDR|CMX|NEF|NRW|ERF|ORF|RAF|MRW|DCR|X3F|SRF|ARW|PEF|RW2|RWL|SRW|OGG|DWG|DXF|DDS|JPM|PLT|CR2|DNG|PS|PDF|AI|CGM|SVG|DCM|HDR|FITS|HDP|WDP|EXR|WSQ|JLS|WBZ|WBC|MP4|MKV|MTS|M2TS|TS|M2T|M4V|3GP|VOB|
    
    
    i1 = ""
    i1 = i1 + "*.JPG|JPEG|JPE|GIF|BMP|DIB|TIF|TIFF|TGA|PCD|SUN|"
    i1 = i1 + "RAS||PCX|PCDSUN|"
    i1 = i1 + "IFF|PPM|PGM|PBM|"
'    i1 = i1 + "WMF|EMF|WAV|PSD|"
    i1 = i1 + "MPE|"
    i1 = i1 + "CUR|DCX|EPS|"
    i1 = i1 + "G3|AIF|AU|SND|"
    i1 = i1 + "PSP|PSPIMAGE|ICL|SFW|KDC|"
    i1 = i1 + "DCM|ACR|FPX|XBM|XPM|DJVU|SWF|"
    i1 = i1 + "IMG|IW44|WBMP|SGI|RGB|RLE|MED|RLE|SFF|NLM|NOL|NGG|GSM|"
    i1 = i1 + "JP2|JPC|J2K|JPF|FSH|CRW|B3D|"
    i1 = i1 + "TTF|SID|MNG|JNG|"
    i1 = i1 + "X3F|SRF|ARW|PEF|RW2|RWL|SRW|"
    i1 = i1 + "DWG|DXF|DDS|JPM|PLT|"
    i1 = i1 + "PDF|AI|CGM|SVG|DCM|HDR|FITS|HDP|WDP|EXR|WSQ|"
    i1 = i1 + "JLS|WBZ|WBC"
    
'    i1 = i1 + ";" + Replace(i1, "|", ";*.")
    
'    ScanPath.cboMask.Text = i1
    '**ICO|
    'MID|RMI|
    'RA|MP3|
    'WMA|
    'RAW|
    'ECW|ASF|TXT|WAD|ICS|IDS|CDR|CMX|NEF|NRW|ERF|ORF|RAF|MRW|
    'DCR|
    'OGG|
    'CR2|DNG|PS|
    
    
    i2 = i2 + "AVI|RM|"
    i2 = i2 + "MPG|"
    i2 = i2 + "MPEG|"
    i2 = i2 + "MOV|"
    i2 = i2 + "CAM|"
    i2 = i2 + "WMV|FLV|FV|"
    i2 = i2 + "MP4|MKV|MTS|M2TS|TS|M2T|M4V|3GP|VOB"
    
    i2 = Replace(i1 + i2, "|", ";*.")
    
    
'    ScanPath.cboMask.Text = i2
    
    
    
    ScanPath.lblCount9.Caption = Replace(ScanPath.cboMask.Text, ";", "; ")

    

    VAR2 = Format(Now, "YYYY_MM_DD_HH_MM_SS") + ".txt"
    
    FILE11 = "D:\#0 IRFAR SHOW SCRIPT FILE SET\#00 IRFAR SCRIPT -- " + VAR2
    FILE12 = "D:\#0 IRFAR SHOW SCRIPT FILE SET\#00 IRFAR SCRIPT APPEND BUILD -- " + VAR2
    
    If Dir(FILE11) <> "" Then Kill FILE11
    If Dir(FILE12) <> "" Then Kill FILE12
    
        
    FF01 = FreeFile
    Open FILE11 For Output As #FF01

    FF02 = FreeFile ' LESS MODIFY
    Open FILE12 For Append As #FF02
    
    
    If IRFAR_TO_DO_FROM_SEND_TO = False Then
    
    On Error Resume Next
    For Each Control In LabDR
    
        MCHAR = Mid(Control.Caption, 2, 1) + ":\"
        Err.Clear
        Set F = FS.GetDrive(MCHAR)
                
        If Err.Number = 0 Then
            If F.IsReady = True Then
                ScanPath.TxtPath.Text = MCHAR
                ScanPath.Timer1.Enabled = True
                ScanPath.cmdScan_Click
            End If
        End If
    Next
    
    Else
        E$ = Command$
'        E$ = "C:\Program Files\# NO INSTALL REQUIRED\GOOGLEIMAGE-MEDIADOWNLOADER"
        If Mid$(E$, 1, 1) = """" Then E$ = Mid$(E$, 2)
    
        If Mid$(E$, Len(E$), 1) = """" Then E$ = Mid$(E$, 1, Len(E$) - 1)
        If FS.FileExists(E$) = True Then E$ = Mid$(E$, 1, InStrRev(E$, "\"))
            ScanPath.TxtPath.Text = E$
            ScanPath.Timer1.Enabled = True
            ScanPath.cmdScan_Click
    End If
    
    Print #FF01, "E:\MY WEBS\MATTHEWLAN.COM WEB\LOVEFOLDER\COOL\NHS\05 AGENT ORANGE.JPG"
    Print #FF01, "E:\MY WEBS\MATTHEWLAN.COM WEB\LOVEFOLDER\COOL\NHS\05 AGENT ORANGE.JPG"
    Print #FF02, "E:\MY WEBS\MATTHEWLAN.COM WEB\LOVEFOLDER\COOL\NHS\05 AGENT ORANGE.JPG"
    Print #FF02, "E:\MY WEBS\MATTHEWLAN.COM WEB\LOVEFOLDER\COOL\NHS\05 AGENT ORANGE.JPG"

    
        
        
    Close FF01, FF02
    
    FILE13 = "C:\TEMP\IRFAN_SLIDESHOW.txt"
    FS.CopyFile FILE12, FILE13
    FILE14 = "C:\TEMP\IRFAN_SLIDESHOW_BACKUP.txt"
    FS.CopyFile FILE12, FILE13

    If IRFAR_TO_DO_FROM_SEND_TO = False Then
        Shell "C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe """ + FILE13 + """", vbNormalFocus
    End If
    
'    Call Label6_Click
    
    Shell "C:\Program Files\IrfanView\i_view32.exe /killmesoftly"
    Shell "C:\Program Files\IrfanView\i_view32.exe ""/slideshow=" + FILE13 + " /ONE /bf /silent /gray /RELOADONLOOP /QUITE""", vbMinimizedNoFocus


    If IRFAR_TO_DO_FROM_SEND_TO = True Then
        End
    End If
End Sub


Sub XSCRIPT_CHUNK_AND_PROCESS()
    
    XSCRIPTCHUNK
    
    FILE17 = "C:\TEMP\IRFAN_SLIDESHOW-A-Z.TXT"
    FILE18 = "C:\TEMP\IRFAN_SLIDESHOW.TXT"
    
    FS.CopyFile FILE18, FILE17
    
    FF01 = FreeFile
    Open FILE17 For Binary As #FF01
    FF02 = FreeFile
    Open FILE18 For Output As #FF02
    
    Do
        Line Input #FF01, FILESPEC
        If FILESPEC <> OFILESPEC Then
            OFILESPEC = FILESPEC 'INTERLACE
            
            Call XSCRIPT(FILESPEC)  ' THROW IN FF02
        End If
    Loop Until EOF(FF01)
    Close FF01, FF02

    If WORKFLAG = True Then MNU_INI_BACK_NAUGHT_Click

End Sub
Sub XSCRIPTCHUNK()

    'PROBLEM
    
    'STRIPPER DONT WANT FROM EXTENTION
    'STRIPPER DO WANT FROM EXTENTION
    
    Dim FF

    On Error Resume Next
    
    ReDim HHAR01(30000)
    ReDim HHAR02(30000)
    HHC01 = 0
    HHC02 = -1
    FF = FreeFile
    Open "C:\TEMP\IRFAN_SLIDESHOW_X_SCRIPT.TXT" For Input As #FF
        Do
            Line Input #FF, LL1
            If Trim(LL1) <> "" Then
                If FLAG <> "SET" Then
                    HHC01 = HHC01 + 1
                    HHAR01(HHC01) = UCase(LL1)
                End If
                
                If Mid(LL1, 1, 3) = "###" Or FLAG = "SET" Then
                    FLAG = "SET"
                    HHC02 = HHC02 + 1
                    If HHC02 > 0 Then
                        HHAR02(HHC02) = UCase(LL1)
                    End If
                End If
            End If
        Loop Until EOF(FF)
    Close #FF, FF1
    
    ReDim Preserve HHAR01(HHC01)
    ReDim Preserve HHAR02(HHC02)
    
End Sub



Public Sub XSCRIPT(FILESPEC)
'   On Error GoTo EXITSUB
'   If IRFAR_TO_DO_FROM_SEND_TO = True Then TAGX = False: GoTo WRITEANDEXIT
    TAGX = False
    
    For R5 = 1 To UBound(HHAR01)
        If InStr("*" + UCase(FILESPEC), HHAR01(R5)) > 0 Then
            TAGX = True
            WORKFLAG = True
            Exit For
        End If
    Next
    
    
    If TAGX = True Then
        For R5 = 1 To UBound(HHAR02)
            ' -- EXTENTION WANT CHECK
            
            ' -- ALL EXTENTION WANT CHECK
            
            If Mid(HHAR02(R5), Len(HHAR02(R5)), 1) = "\" Then
                 i5 = Mid(FILESPEC, 4, InStrRev(FILESPEC, "\") - 3)
                    
'                    If Len(i5) = Len(Mid(HHAR02(R5), 5 + 1)) Then Stop 'SWAP DIR SAME LENGHT
                    
                    'If Len(i5) = Len(Mid(HHAR02(R5), 5 + 1)) Then Stop ' NOT A CHECK WITH ROOT AT MOMENT
              
            End If
                
                
'Len(HHAR02(R5))

            If Mid(HHAR02(R5), 1, 4) = Mid(UCase(FILESPEC), Len(FILESPEC) - 3) _
             Or Mid(HHAR02(R5), 1, 1) = "*" Then
            
            ' -- FOLDER WANT ---
            If InStr("*" + UCase(FILESPEC), Mid(HHAR02(R5), 5 + 1)) > 0 Then
                    TAGX = False    '  --- ALLOW OBJECTIVE
                    WORKFLAG = True
                    Exit For
                End If
            End If
        Next
    End If
    
WRITEANDEXIT:
    If TAGX = False Then
        Print #FF02, FILESPEC
    End If
    
End Sub


Private Sub MNU_IRFAR_02_Click()
    
    Me.WindowState = vbMinimized
    
    FILE13 = "C:\TEMP\IRFAN_SLIDESHOW.txt"


    
'PROBLEM
'    Call XSCRIPTCHUNK
    
    
'    Shell "C:\Program Files\# NO INSTALL REQUIRED\Text View\TextView\Textview.exe """ + FILE13 + """", vbNormalFocus
    
    Shell "C:\Program Files\IrfanView\i_view32.exe /killmesoftly"
    Shell "C:\Program Files\IrfanView\i_view32.exe ""/slideshow=" + FILE13 + " /ONE /bf /silent /RELOADONLOOP /QUITE""", vbMinimizedNoFocus
    ' /gray
End Sub


Private Sub MNU_INI_BACK_NAUGHT_Click()

FR = FreeFile
Open "C:\Program Files\IrfanView\i_view32.ini" For Input As #FR
    Do
        Line Input #FR, L
        If InStr("-" + L, "StartIndex=") > 0 Then
            A1 = Val(Mid(L, Len("StartIndex=") + 1))
        End If
        If EOF(FR) Then A1 = "2"
    Loop Until A1 <> ""
Close #FR

FR = FreeFile
Open "C:\Program Files\IrfanView\i_view32.ini" For Input As #FR
    RC = Input(LOF(FR), FR)
Close #FR

A1 = Trim(Str(A1))
B1 = Trim(Str(0))
RC = Replace(RC, "StartIndex=" + A1, "StartIndex=" + B1)
INIA1 = A1

FR = FreeFile
Open "C:\Program Files\IrfanView\i_view32.ini" For Output Lock Read Write As #FR
    Print #FR, RC
Close #FR

End Sub

Private Sub MNU_VB_FOLDER_Click()

Shell "EXPLORER /SELECT, " + App.Path + "\" + App.EXEName + ".VBP", vbMaximizedFocus

End Sub

Private Sub Timer1_Timer()

COUNTPROCES5 = COUNTPROCES4
COUNTPROCES4 = 0
If COUNTPROCES9 < COUNTPROCES8 Then COUNTPROCES9 = COUNTPROCES8

End Sub

Private Sub Timer2_Timer()

COUNTPROCES9 = COUNTPROCES8
COUNTPROCES8 = 0

End Sub
