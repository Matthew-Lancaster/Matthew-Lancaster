VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   BackColor       =   &H80000007&
   Caption         =   "XNTP Control Test"
   ClientHeight    =   10050
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   16380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NNTP-Test.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10050
   ScaleWidth      =   16380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00A6FC94&
      Caption         =   "&Ready To Go Presss Here"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   277
      Top             =   6585
      Width           =   1830
   End
   Begin VB.TextBox Text3FX 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   0
      TabIndex        =   276
      Text            =   "Date My Time "
      Top             =   5205
      Width           =   7800
   End
   Begin VB.TextBox Text3dx 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   0
      TabIndex        =   275
      Text            =   "Last Run -----"
      Top             =   6105
      Width           =   7800
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Open Header Logg - Choose Group To Left"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   11775
      Style           =   1  'Graphical
      TabIndex        =   274
      Top             =   7755
      Width           =   2070
   End
   Begin VB.TextBox Text3x 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   0
      TabIndex        =   170
      Text            =   "Date ------------"
      Top             =   5670
      Width           =   7800
   End
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   330
      Left            =   0
      TabIndex        =   41
      Top             =   4860
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      ForeColor       =   &H80000017&
      Height          =   2115
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2670
      Width           =   7800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&First"
      Height          =   375
      Left            =   5160
      TabIndex        =   23
      Top             =   8490
      Width           =   630
   End
   Begin VB.CheckBox chkDebug 
      Appearance      =   0  'Flat
      Caption         =   "Debug &mode"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4875
      TabIndex        =   22
      Top             =   8085
      Width           =   1215
   End
   Begin VB.TextBox txtNewsgroup 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   885
      TabIndex        =   8
      Text            =   "alt.test"
      Top             =   7395
      Width           =   6915
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   6645
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   7020
      Width           =   1155
   End
   Begin VB.TextBox txtUsername 
      Height          =   345
      Left            =   4725
      TabIndex        =   4
      Top             =   7020
      Width           =   1170
   End
   Begin VB.TextBox txtServer 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   885
      TabIndex        =   2
      Text            =   "news.internet.com"
      Top             =   7005
      Width           =   3060
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last"
      Height          =   375
      Left            =   4515
      TabIndex        =   19
      Top             =   8490
      Width           =   630
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   3885
      TabIndex        =   18
      Top             =   8490
      Width           =   630
   End
   Begin VB.CommandButton cmdStat 
      Caption         =   "S&tat"
      Height          =   375
      Left            =   2910
      TabIndex        =   17
      Top             =   8490
      Width           =   975
   End
   Begin VB.CommandButton cmdArticle 
      Caption         =   "&Article"
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   8490
      Width           =   975
   End
   Begin VB.CommandButton cmdHeader 
      Caption         =   "&Header"
      Height          =   375
      Left            =   975
      TabIndex        =   15
      Top             =   8490
      Width           =   975
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&End"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5805
      TabIndex        =   20
      Top             =   8490
      Width           =   975
   End
   Begin VB.CommandButton cmdBody 
      Caption         =   "&Body"
      Height          =   375
      Left            =   1935
      TabIndex        =   16
      Top             =   8490
      Width           =   975
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   375
      Left            =   2910
      TabIndex        =   12
      Top             =   8085
      Width           =   975
   End
   Begin VB.CommandButton cmdXover 
      Caption         =   "&Xover"
      Height          =   375
      Left            =   3885
      TabIndex        =   13
      Top             =   8085
      Width           =   975
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&List"
      Height          =   375
      Left            =   1950
      TabIndex        =   11
      Top             =   8085
      Width           =   960
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      ForeColor       =   &H80000017&
      Height          =   2835
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -15
      Width           =   7800
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "&Disconnect"
      Height          =   375
      Left            =   975
      TabIndex        =   10
      Top             =   8085
      Width           =   960
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   8085
      Width           =   975
   End
   Begin NNTPTEST.NNTP NNTP1 
      Left            =   7125
      Top             =   8430
      _ExtentX        =   741
      _ExtentY        =   741
      Debugger        =   0   'False
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   7
      Left            =   10575
      TabIndex        =   273
      Top             =   8355
      Width           =   1155
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   10575
      TabIndex        =   272
      Top             =   7755
      Width           =   1155
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   6
      Left            =   10575
      TabIndex        =   271
      Top             =   8055
      Width           =   1155
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   8
      Left            =   10575
      TabIndex        =   270
      Top             =   8655
      Width           =   1155
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   9
      Left            =   10575
      TabIndex        =   269
      Top             =   8955
      Width           =   1155
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9390
      TabIndex        =   268
      Top             =   7455
      Width           =   1155
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7860
      TabIndex        =   267
      Top             =   7455
      Width           =   1500
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   63
      Left            =   10455
      TabIndex        =   266
      Top             =   2865
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   62
      Left            =   10455
      TabIndex        =   265
      Top             =   2775
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   61
      Left            =   10530
      TabIndex        =   264
      Top             =   2835
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   60
      Left            =   10530
      TabIndex        =   263
      Top             =   2745
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   59
      Left            =   10440
      TabIndex        =   262
      Top             =   2805
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   58
      Left            =   10440
      TabIndex        =   261
      Top             =   2715
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   57
      Left            =   10515
      TabIndex        =   260
      Top             =   2775
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   56
      Left            =   10515
      TabIndex        =   259
      Top             =   2685
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   55
      Left            =   10425
      TabIndex        =   258
      Top             =   2805
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   54
      Left            =   10425
      TabIndex        =   257
      Top             =   2715
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   53
      Left            =   10500
      TabIndex        =   256
      Top             =   2775
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   52
      Left            =   10500
      TabIndex        =   255
      Top             =   2685
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   51
      Left            =   10410
      TabIndex        =   254
      Top             =   2745
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   50
      Left            =   10410
      TabIndex        =   253
      Top             =   2655
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   49
      Left            =   10485
      TabIndex        =   252
      Top             =   2715
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   48
      Left            =   10485
      TabIndex        =   251
      Top             =   2625
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   47
      Left            =   10470
      TabIndex        =   250
      Top             =   2805
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   46
      Left            =   10470
      TabIndex        =   249
      Top             =   2715
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   45
      Left            =   10545
      TabIndex        =   248
      Top             =   2775
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   44
      Left            =   10545
      TabIndex        =   247
      Top             =   2685
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   43
      Left            =   10455
      TabIndex        =   246
      Top             =   2745
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   42
      Left            =   10455
      TabIndex        =   245
      Top             =   2655
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   41
      Left            =   10530
      TabIndex        =   244
      Top             =   2715
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   40
      Left            =   10530
      TabIndex        =   243
      Top             =   2625
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   39
      Left            =   10440
      TabIndex        =   242
      Top             =   2745
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   38
      Left            =   10440
      TabIndex        =   241
      Top             =   2655
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   37
      Left            =   10515
      TabIndex        =   240
      Top             =   2715
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   36
      Left            =   10515
      TabIndex        =   239
      Top             =   2625
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   35
      Left            =   10425
      TabIndex        =   238
      Top             =   2685
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   34
      Left            =   10425
      TabIndex        =   237
      Top             =   2595
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   33
      Left            =   10500
      TabIndex        =   236
      Top             =   2655
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
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
      Index           =   32
      Left            =   10500
      TabIndex        =   235
      Top             =   2565
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   31
      Left            =   10455
      TabIndex        =   234
      Top             =   2775
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   30
      Left            =   10455
      TabIndex        =   233
      Top             =   2685
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   29
      Left            =   10530
      TabIndex        =   232
      Top             =   2745
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   28
      Left            =   10530
      TabIndex        =   231
      Top             =   2655
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   27
      Left            =   10440
      TabIndex        =   230
      Top             =   2715
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   26
      Left            =   10440
      TabIndex        =   229
      Top             =   2625
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   25
      Left            =   10515
      TabIndex        =   228
      Top             =   2685
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   24
      Left            =   10515
      TabIndex        =   227
      Top             =   2595
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   23
      Left            =   10425
      TabIndex        =   226
      Top             =   2715
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   22
      Left            =   10425
      TabIndex        =   225
      Top             =   2625
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   21
      Left            =   10500
      TabIndex        =   224
      Top             =   2685
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   20
      Left            =   10500
      TabIndex        =   223
      Top             =   2595
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   19
      Left            =   10410
      TabIndex        =   222
      Top             =   2655
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   18
      Left            =   10410
      TabIndex        =   221
      Top             =   2565
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   17
      Left            =   10485
      TabIndex        =   220
      Top             =   2625
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   16
      Left            =   10485
      TabIndex        =   219
      Top             =   2535
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   15
      Left            =   10470
      TabIndex        =   218
      Top             =   2715
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   14
      Left            =   10470
      TabIndex        =   217
      Top             =   2625
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   13
      Left            =   10545
      TabIndex        =   216
      Top             =   2685
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   12
      Left            =   10545
      TabIndex        =   215
      Top             =   2595
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   10455
      TabIndex        =   214
      Top             =   2655
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   10455
      TabIndex        =   213
      Top             =   2565
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   10530
      TabIndex        =   212
      Top             =   2625
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   10530
      TabIndex        =   211
      Top             =   2535
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   10440
      TabIndex        =   210
      Top             =   2655
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   10440
      TabIndex        =   209
      Top             =   2565
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   10515
      TabIndex        =   208
      Top             =   2625
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   10515
      TabIndex        =   207
      Top             =   2535
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   10425
      TabIndex        =   206
      Top             =   2595
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   10425
      TabIndex        =   205
      Top             =   2505
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   10500
      TabIndex        =   204
      Top             =   2565
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   10500
      TabIndex        =   203
      Top             =   2475
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10560
      TabIndex        =   202
      Top             =   4005
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10545
      TabIndex        =   201
      Top             =   3930
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10635
      TabIndex        =   200
      Top             =   3930
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10620
      TabIndex        =   199
      Top             =   3855
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10455
      TabIndex        =   198
      Top             =   3900
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10440
      TabIndex        =   197
      Top             =   3825
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10530
      TabIndex        =   196
      Top             =   3825
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10515
      TabIndex        =   195
      Top             =   3750
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10455
      TabIndex        =   194
      Top             =   3885
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10440
      TabIndex        =   193
      Top             =   3810
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10530
      TabIndex        =   192
      Top             =   3810
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10515
      TabIndex        =   191
      Top             =   3735
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10350
      TabIndex        =   190
      Top             =   3780
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10335
      TabIndex        =   189
      Top             =   3705
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10425
      TabIndex        =   188
      Top             =   3705
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10410
      TabIndex        =   187
      Top             =   3630
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10425
      TabIndex        =   186
      Top             =   3855
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10410
      TabIndex        =   185
      Top             =   3780
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10500
      TabIndex        =   184
      Top             =   3780
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10485
      TabIndex        =   183
      Top             =   3705
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10320
      TabIndex        =   182
      Top             =   3750
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10305
      TabIndex        =   181
      Top             =   3675
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10395
      TabIndex        =   180
      Top             =   3675
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10380
      TabIndex        =   179
      Top             =   3600
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10320
      TabIndex        =   178
      Top             =   3735
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10305
      TabIndex        =   177
      Top             =   3660
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10395
      TabIndex        =   176
      Top             =   3660
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10380
      TabIndex        =   175
      Top             =   3585
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10215
      TabIndex        =   174
      Top             =   3630
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10200
      TabIndex        =   173
      Top             =   3555
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10290
      TabIndex        =   172
      Top             =   3555
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "*"
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
      Left            =   10275
      TabIndex        =   171
      Top             =   3480
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   95
      Left            =   10590
      TabIndex        =   169
      Top             =   1890
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   94
      Left            =   10545
      TabIndex        =   168
      Top             =   1875
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   93
      Left            =   10575
      TabIndex        =   167
      Top             =   1785
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   92
      Left            =   10545
      TabIndex        =   166
      Top             =   1725
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   91
      Left            =   10500
      TabIndex        =   165
      Top             =   1710
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   90
      Left            =   10545
      TabIndex        =   164
      Top             =   1620
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   89
      Left            =   10455
      TabIndex        =   163
      Top             =   1605
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   88
      Left            =   10410
      TabIndex        =   162
      Top             =   1590
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   87
      Left            =   10440
      TabIndex        =   161
      Top             =   1500
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   86
      Left            =   10410
      TabIndex        =   160
      Top             =   1440
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   85
      Left            =   10365
      TabIndex        =   159
      Top             =   1425
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   84
      Left            =   10395
      TabIndex        =   158
      Top             =   1335
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   83
      Left            =   10485
      TabIndex        =   157
      Top             =   1875
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   82
      Left            =   10440
      TabIndex        =   156
      Top             =   1860
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   81
      Left            =   10470
      TabIndex        =   155
      Top             =   1770
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   80
      Left            =   10440
      TabIndex        =   154
      Top             =   1710
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   79
      Left            =   10395
      TabIndex        =   153
      Top             =   1695
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   78
      Left            =   10425
      TabIndex        =   152
      Top             =   1605
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   77
      Left            =   10350
      TabIndex        =   151
      Top             =   1590
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   76
      Left            =   10305
      TabIndex        =   150
      Top             =   1575
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   75
      Left            =   10335
      TabIndex        =   149
      Top             =   1485
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   74
      Left            =   10305
      TabIndex        =   148
      Top             =   1425
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   73
      Left            =   10260
      TabIndex        =   147
      Top             =   1410
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   72
      Left            =   10290
      TabIndex        =   146
      Top             =   1320
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   71
      Left            =   10590
      TabIndex        =   145
      Top             =   1815
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   70
      Left            =   10545
      TabIndex        =   144
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   69
      Left            =   10575
      TabIndex        =   143
      Top             =   1710
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   68
      Left            =   10545
      TabIndex        =   142
      Top             =   1650
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   67
      Left            =   10500
      TabIndex        =   141
      Top             =   1635
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   66
      Left            =   10530
      TabIndex        =   140
      Top             =   1545
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   65
      Left            =   10455
      TabIndex        =   139
      Top             =   1530
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   64
      Left            =   10410
      TabIndex        =   138
      Top             =   1515
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   63
      Left            =   10440
      TabIndex        =   137
      Top             =   1425
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   62
      Left            =   10410
      TabIndex        =   136
      Top             =   1365
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   61
      Left            =   10365
      TabIndex        =   135
      Top             =   1350
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   60
      Left            =   10395
      TabIndex        =   134
      Top             =   1260
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   59
      Left            =   10485
      TabIndex        =   133
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   58
      Left            =   10440
      TabIndex        =   132
      Top             =   1785
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   57
      Left            =   10470
      TabIndex        =   131
      Top             =   1695
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   56
      Left            =   10440
      TabIndex        =   130
      Top             =   1635
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   55
      Left            =   10395
      TabIndex        =   129
      Top             =   1620
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   54
      Left            =   10425
      TabIndex        =   128
      Top             =   1530
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   53
      Left            =   10350
      TabIndex        =   127
      Top             =   1515
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   52
      Left            =   10305
      TabIndex        =   126
      Top             =   1500
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   51
      Left            =   10335
      TabIndex        =   125
      Top             =   1410
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   50
      Left            =   10305
      TabIndex        =   124
      Top             =   1350
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   49
      Left            =   10260
      TabIndex        =   123
      Top             =   1335
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   48
      Left            =   10290
      TabIndex        =   122
      Top             =   1245
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   47
      Left            =   10530
      TabIndex        =   121
      Top             =   1860
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   46
      Left            =   10485
      TabIndex        =   120
      Top             =   1845
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   45
      Left            =   10515
      TabIndex        =   119
      Top             =   1755
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   44
      Left            =   10485
      TabIndex        =   118
      Top             =   1695
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   43
      Left            =   10440
      TabIndex        =   117
      Top             =   1680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   42
      Left            =   10470
      TabIndex        =   116
      Top             =   1590
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   41
      Left            =   10395
      TabIndex        =   115
      Top             =   1575
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   40
      Left            =   10350
      TabIndex        =   114
      Top             =   1560
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   39
      Left            =   10380
      TabIndex        =   113
      Top             =   1470
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   38
      Left            =   10350
      TabIndex        =   112
      Top             =   1410
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   37
      Left            =   10305
      TabIndex        =   111
      Top             =   1395
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   36
      Left            =   10335
      TabIndex        =   110
      Top             =   1305
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   35
      Left            =   10425
      TabIndex        =   109
      Top             =   1845
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   34
      Left            =   10380
      TabIndex        =   108
      Top             =   1830
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   33
      Left            =   10410
      TabIndex        =   107
      Top             =   1740
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Index           =   32
      Left            =   10380
      TabIndex        =   106
      Top             =   1680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10335
      TabIndex        =   105
      Top             =   1665
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10365
      TabIndex        =   104
      Top             =   1575
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10290
      TabIndex        =   103
      Top             =   1560
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10245
      TabIndex        =   102
      Top             =   1545
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10275
      TabIndex        =   101
      Top             =   1455
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10245
      TabIndex        =   100
      Top             =   1395
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10200
      TabIndex        =   99
      Top             =   1380
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10230
      TabIndex        =   98
      Top             =   1290
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10530
      TabIndex        =   97
      Top             =   1785
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10485
      TabIndex        =   96
      Top             =   1770
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10515
      TabIndex        =   95
      Top             =   1680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10485
      TabIndex        =   94
      Top             =   1620
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10440
      TabIndex        =   93
      Top             =   1605
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10470
      TabIndex        =   92
      Top             =   1515
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10395
      TabIndex        =   91
      Top             =   1500
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10350
      TabIndex        =   90
      Top             =   1485
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10380
      TabIndex        =   89
      Top             =   1395
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10350
      TabIndex        =   88
      Top             =   1335
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10305
      TabIndex        =   87
      Top             =   1320
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10335
      TabIndex        =   86
      Top             =   1230
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10425
      TabIndex        =   85
      Top             =   1770
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10380
      TabIndex        =   84
      Top             =   1755
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10410
      TabIndex        =   83
      Top             =   1665
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10380
      TabIndex        =   82
      Top             =   1605
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10335
      TabIndex        =   81
      Top             =   1590
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10365
      TabIndex        =   80
      Top             =   1500
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10290
      TabIndex        =   79
      Top             =   1485
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10245
      TabIndex        =   78
      Top             =   1470
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10275
      TabIndex        =   77
      Top             =   1380
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10245
      TabIndex        =   76
      Top             =   1320
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10200
      TabIndex        =   75
      Top             =   1305
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "*"
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
      Left            =   10230
      TabIndex        =   74
      Top             =   1215
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   31
      Left            =   10260
      TabIndex        =   73
      Top             =   705
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   30
      Left            =   10230
      TabIndex        =   72
      Top             =   630
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   29
      Left            =   10290
      TabIndex        =   71
      Top             =   645
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   28
      Left            =   10260
      TabIndex        =   70
      Top             =   615
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   27
      Left            =   10215
      TabIndex        =   69
      Top             =   630
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   26
      Left            =   10230
      TabIndex        =   68
      Top             =   600
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   25
      Left            =   10215
      TabIndex        =   67
      Top             =   540
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   24
      Left            =   10140
      TabIndex        =   66
      Top             =   495
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   23
      Left            =   10200
      TabIndex        =   65
      Top             =   525
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   22
      Left            =   10200
      TabIndex        =   64
      Top             =   480
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   21
      Left            =   10155
      TabIndex        =   63
      Top             =   390
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   20
      Left            =   10125
      TabIndex        =   62
      Top             =   450
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   19
      Left            =   10035
      TabIndex        =   61
      Top             =   480
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   18
      Left            =   10080
      TabIndex        =   60
      Top             =   390
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   17
      Left            =   10050
      TabIndex        =   59
      Top             =   360
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   16
      Left            =   10035
      TabIndex        =   58
      Top             =   300
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   15
      Left            =   10215
      TabIndex        =   57
      Top             =   555
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   14
      Left            =   10185
      TabIndex        =   56
      Top             =   480
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   10245
      TabIndex        =   55
      Top             =   495
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   12
      Left            =   10215
      TabIndex        =   54
      Top             =   465
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   11
      Left            =   10170
      TabIndex        =   53
      Top             =   480
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   10
      Left            =   10185
      TabIndex        =   52
      Top             =   450
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   9
      Left            =   10170
      TabIndex        =   51
      Top             =   390
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   8
      Left            =   10095
      TabIndex        =   50
      Top             =   345
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   7
      Left            =   10155
      TabIndex        =   49
      Top             =   375
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   6
      Left            =   10155
      TabIndex        =   48
      Top             =   330
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   10110
      TabIndex        =   47
      Top             =   240
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   10080
      TabIndex        =   46
      Top             =   300
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   9990
      TabIndex        =   45
      Top             =   330
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   10035
      TabIndex        =   44
      Top             =   240
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   10005
      TabIndex        =   43
      Top             =   210
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   9990
      TabIndex        =   42
      Top             =   150
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Rev"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7845
      TabIndex        =   40
      Top             =   8955
      Width           =   1515
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "SZ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7845
      TabIndex        =   39
      Top             =   8655
      Width           =   1515
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "KMan Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7875
      TabIndex        =   38
      Top             =   8040
      Width           =   1500
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borg A.Philo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7860
      TabIndex        =   37
      Top             =   7740
      Width           =   1500
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "KMan Kook"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7845
      TabIndex        =   36
      Top             =   8355
      Width           =   1515
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   9390
      TabIndex        =   35
      Top             =   8955
      Width           =   1155
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   9390
      TabIndex        =   34
      Top             =   8655
      Width           =   1155
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   9390
      TabIndex        =   33
      Top             =   8055
      Width           =   1155
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   9390
      TabIndex        =   32
      Top             =   7755
      Width           =   1155
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   9390
      TabIndex        =   31
      Top             =   8355
      Width           =   1155
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "To Go To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7860
      TabIndex        =   30
      Top             =   6870
      Width           =   1500
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Now"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7860
      TabIndex        =   29
      Top             =   7155
      Width           =   1500
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tott"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7860
      TabIndex        =   28
      Top             =   6585
      Width           =   1500
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9390
      TabIndex        =   27
      Top             =   6870
      Width           =   1155
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9390
      TabIndex        =   26
      Top             =   6585
      Width           =   1155
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9390
      TabIndex        =   25
      Top             =   7155
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Ne&wsgroup"
      Height          =   345
      Left            =   15
      TabIndex        =   7
      Top             =   7455
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Password"
      Height          =   195
      Left            =   5925
      TabIndex        =   5
      Top             =   7080
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Username"
      Height          =   195
      Left            =   3975
      TabIndex        =   3
      Top             =   7065
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Newsser&ver"
      Height          =   225
      Left            =   0
      TabIndex        =   1
      Top             =   7065
      Width           =   855
   End
   Begin VB.Label labInfo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   15
      TabIndex        =   21
      Top             =   6615
      Width           =   7800
   End
   Begin VB.Menu Mnu_VB 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_O 
      Caption         =   "OPTIONS"
      Begin VB.Menu MNU_OUTPUT 
         Caption         =   "DONT WRITE OUTPUTS NOT EVEN COUNTERS"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Option Compare Text

Dim GGG(100)

Dim TTH4, TT5, NowStart, DONEWIDTH, DATEXXMyTime


'NEED NOT JUST TIME BUT HOW FAR BACK - COZ GOES BACK FOR FILE APPENDING SAKE
Dim File As String, f As Integer, V As Variant, NEWWRITEHEADERBATCH

Dim LastStart
Dim EasyRem
Dim NG(), OMEWIDTH, TTH
Dim DATEXX, DATEXX2, QUICKEXIT, YT2
Dim HH()
Dim A1 As String, A2 As String, Text4, Text5, Text6
Dim Buffer1 As String
Dim Buffer2 As String, DONEARTICALORHEADER

Public TextDebug, TextDebug2, XXString, Text3, XRef, LinesCount, GODO, LinesCount2

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type OFSTRUCT
  cBytes     As Byte
  fFixedDisk As Byte
  nErrCode   As Integer
  Reserved1  As Integer
  Reserved2  As Integer
  szPathName As String * 128
End Type

Const OF_SHARE_COMPAT = &H0
Const OF_SHARE_EXCLUSIVE = &H10
Const OF_SHARE_DENY_WRITE = &H20
Const OF_SHARE_DENY_READ = &H30
Const OF_SHARE_DENY_NONE = &H40

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long






Private Sub chkDebug_Click()
    
    'MAKE SURE ALWAYS ON
    
    'Toggle the debugger mode
    NNTP1.Debugger = chkDebug.Value = 1
End Sub

Private Sub cmdArticle_Click()
    'Get the current article (header and body)
    Msg "Getting article..."
    NNTP1.GetArticle
End Sub

Private Sub cmdBody_Click()
    'get the current article (header only)
    Msg "Getting body..."
    NNTP1.GetBody
End Sub

Private Sub cmdConnect_Click()
    'Connecting to a newsserver
    Msg "Connecting..."
    NNTP1.Host = txtServer
    NNTP1.Username = txtUsername
    NNTP1.Password = txtPassword
    NNTP1.Connect
End Sub


Private Sub cmdDisconnect_Click()
    'Disconnecting from a newsserver
    Msg "Disconnecting..."
    NNTP1.Disconnect
End Sub

Private Sub cmdEnd_Click()
    'Quit the test program
    Unload Me
    End
End Sub

Private Sub cmdHeader_Click()
    'get the current article header
    Msg "Getting header..."
    NNTP1.GetHeader
End Sub

Private Sub cmdHeader_Item_Click()
    'get the current article header
    
    Msg "Getting header..."
    NNTP1.GetHeader
End Sub


Private Sub cmdLast_Click()
    'goto the previous article in the selected group
    Msg "Seeking last article..."
    NNTP1.GotoLast
End Sub

Private Sub cmdList_Click()
    'get a list of newsgroups
    Msg "Receiving list of newsgroups..."
    NNTP1.DateTime = Now() - 92
    NNTP1.GetList
End Sub

Private Sub cmdNext_Click()
    'goto the next article in the selected group
    Msg "Seeking next article..."
    'NNTP1.GotoNext
    NNTP1.GotoNext
End Sub

Private Sub cmdPost_Click()

End Sub

Private Sub cmdSelect_Click()
    'Selecting a newsgroup
    Msg "Selecting a newsgroup..."
    NNTP1.Newsgroup = txtNewsgroup
    NNTP1.SelectGroup
End Sub

Private Sub cmdStat_Click()
    'get the current article state
    Msg "Getting statistic..."
    NNTP1.GetStat
End Sub

Private Sub cmdXover_Click()
    'Retrieving all available article headers
    Msg "Getting article headers..."
    NNTP1.GetXover
End Sub

Sub AUTORESIZE()


'AUTO RESIZE
'Code to Auto Size Form based on controls used Including Menu Bar Sketchy Style
'Make sure form set to scale Twips
'put in your form load

X = 1
y = 1
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True And Control.Visible = True Then
        If Control.Width + Control.Left > X Then X = Control.Width + Control.Left
        If Control.Height + Control.Top > y Then y = Control.Height + Control.Top
        If InStr(Control.Name, "Mnu_") > 0 Then mnu = 1
    End If
Next
'On Error GoTo 0

Me.Width = (X + 100)
If mnu = 1 Then
    pluso = 720 ': pluso = 1100 'Sometimes Different
Else
    pluso = 450
End If

Me.Height = (y + pluso)
'Me.Refresh
DoEvents

If OWI <> Me.Width Then
    DONEWIDTH = False
End If
OWI = Me.Width

If DONEWIDTH = False Then

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

End If

DONEWIDTH = True


End Sub


Sub UserReSizeLabels()
    
    
    WIDTHVAR = 7800
    Text1.Width = WIDTHVAR
    Text2.Width = WIDTHVAR
    Text1.Top = 0
    Text1.Left = 0
    Text2.Top = Text1.Top + Text1.Height + 15
    Text2.Left = 0
    
    PBar1.Top = Text2.Top + Text2.Height + 15
    PBar1.Left = 0
    PBar1.Width = WIDTHVAR
    
    Text3x.Width = WIDTHVAR
    Text3x.Left = 0
    Text3dx.Width = WIDTHVAR
    Text3dx.Left = 0
    Text3FX.Width = WIDTHVAR
    Text3FX.Left = 0

    
    For R = 1 To YT2
    Text1.Height = 2650
        Label11(R) = HH(R)
        If InStr(HH(R), "@") > 0 And Mid(HH(R), 1, 1) <> "@" Then Label11(R) = Mid(HH(R), 1, InStr(HH(R), "@"))
        Label11(R).Left = Text1.Left + Text1.Width + 50
        
        
        Label11(R).FontSize = 14
        Label12(R).FontSize = 14
        Label13(R).FontSize = 14
        Label14(R).FontSize = 14
        Label11(R).FontBold = False
        Label12(R).FontBold = True
        Label13(R).FontBold = False
        Label14(R).FontBold = False

        Label11(R).Visible = True
        Label12(R).Visible = True
        Label13(R).Visible = True
        Label14(R).Visible = True
        Label11(R).AutoSize = True
        Label12(R).AutoSize = True
        Label13(R).AutoSize = True
        Label14(R).AutoSize = True
    Next
    
    
    xx3 = 0: xx4 = 0: xx5 = 0
    For R = 1 To YT2
        If Label12(R) <> "" And Label12(R) <> "" Then
            xx3 = 1
        End If
        If Label13(R) <> "" And Label13(R) <> "*" Then xx4 = 1
        If Label14(R) <> "" And Label14(R) <> "*" Then xx5 = 1
    Next
    
    If xx3 = 0 Then Label12(1) = "***"
    If xx4 = 0 Then Label13(1) = "DAY 00-00-00"
    If xx5 = 0 Then Label14(1) = "DAY 00-00-00"
    
    HXZ = 0
    For R = 1 To YT2
        If HXZ < Label11(R).Height Then HXZ = Label11(R).Height
        If HXZ < Label12(R).Height Then HXZ = Label12(R).Height
        If HXZ < Label13(R).Height Then HXZ = Label13(R).Height
        If HXZ < Label14(R).Height Then HXZ = Label14(R).Height
    Next
    
    XZ = 0: XZ1 = 0: XZ2 = 0: XZ3 = 0
    For R = 1 To YT2
        If XZ < Label11(R).Width Then XZ = Label11(R).Width
        If XZ1 < Label12(R).Width Then XZ1 = Label12(R).Width
        If XZ2 < Label13(R).Width Then XZ2 = Label13(R).Width
        If XZ3 < Label14(R).Width Then XZ3 = Label14(R).Width
        
    Next
    
    Label11(1).Top = Text1.Top
    
    For R = 1 To YT2
        Label11(R).AutoSize = False
        Label12(R).AutoSize = False
        Label13(R).AutoSize = False
        Label14(R).AutoSize = False
    
        Label11(R).Height = HXZ
        Label12(R).Height = HXZ
        Label13(R).Height = HXZ
        Label14(R).Height = HXZ
        If R > 1 Then Label11(R).Top = Label11(R - 1).Top + Label11(R).Height + 15
        
        Label12(R).Top = Label11(R).Top
        Label13(R).Top = Label11(R).Top
        Label14(R).Top = Label11(R).Top
        
        Label11(R).Width = XZ
        Label12(R).Width = XZ1
        
        Label13(R).Width = XZ2
        Label14(R).Width = XZ3
        
        Label12(R).Left = Label11(R).Left + Label11(R).Width + 15
        Label13(R).Left = Label12(R).Left + Label12(R).Width + 15
        Label14(R).Left = Label13(R).Left + Label13(R).Width + 15
        
        
        If Label12(R) = "***" Then Label12(R) = ""
        If Label12(R) = "*" Then Label12(R) = ""
        If Label13(R) = "DAY 00-00-00" Or Label13(R) = "*" Then Label13(R) = ""
        If Label14(R) = "DAY 00-00-00" Or Label14(R) = "*" Then Label14(R) = ""
    Next
    
    Text1.Top = 0
    Text1.Left = 0
    Text2.Top = Text1.Top + Text1.Height + 15
    Text2.Left = 0
    
    PBar1.Top = Text2.Top + Text2.Height + 15
    PBar1.Left = 0
    PBar1.Width = WIDTHVAR
    
    Call AUTORESIZE


End Sub

Sub NEWHEADERWRITE()
    
    If NEWWRITEHEADERBATCH <> False Then Exit Sub
    NEWWRITEHEADERBATCH = True
    
    tth2 = "NEW HEADER BATCH -- " + Format(Now, "DDD DD MMM YYYY - HH:MM:SS") + vbCrLf
    tth2 = tth2 + "FOR GROUP -- " + NG(TT5)
    
    TTH = String(Len(tth2), "-")
    Call DEBUGP
    TTH = tth2
    Call DEBUGP
    TTH = String(Len(tth2), "-")
    Call DEBUGP
    TTH = ""
    Call DEBUGP


End Sub

Private Sub Command3_Click()

    'MAIN ENGINE START TO GO

    ReDim Preserve HH(YT2)
    
    For R = 1 To UBound(HH)
        HH(R) = LCase(HH(R))
    Next
    
    On Error Resume Next
    Kill "TxTPost --*.TXT"
    On Error GoTo 0
    
    File = App.Path + "\Setup.ini"
    If Dir(File, vbHidden And Not vbDirectory) <> "" Then
        f = FreeFile()
        Open File For Input As #f
        Line Input #f, V: txtServer = V
        Line Input #f, V: txtUsername = V
        Line Input #f, V: txtPassword = V
        Line Input #f, V ': txtNewsgroup = V
        'Line Input #f, V: chkDebug = V
        Line Input #f, V: chkDebug = 1
        Line Input #f, V: Text3dx = "Last Run ----- " + V
        
        Close #f
    End If

txtServer = "news.btinternet.com"
Call cmdConnect_Click
DONEARTICALORHEADER = False
Do
    Call Sleep(100)
    DoEvents
Loop Until DONEARTICALORHEADER = True
'Loop Until InStr(Text1, "Server Msg=news.bt.com") > 0

ReDim NG(20)

XD = 0
XD = XD + 1: NG(XD) = "alt.philosophy"
XD = XD + 1: NG(XD) = "24hoursupport.helpdesk"
XD = XD + 1: NG(XD) = "alt.usenet.kooks"
XD = XD + 1: NG(XD) = "alt.support.schizophrenia"
XD = XD + 1: NG(XD) = "alt.slack"


'CHANGE TO XD
ReDim Preserve NG(5)

For TT5 = 1 To UBound(NG)

'NEED INDENT OUT HERE
'UNLESS IN A SUB
'NEED AUTO INDENT CODER BUT PAY FOR ONLY SEEMS

PBARR = False

txtNewsgroup = NG(TT5)
XXString = ""

'HOW FAR BACK DOWN TO ITEM TO GO TO IN REVERSE FOR LOGG APPENDS
Label7 = ""

DONEARTICALORHEADER = False
Call cmdSelect_Click
Do
    Call Sleep(10)
    DoEvents
Loop Until DONEARTICALORHEADER = True
'Loop Until InStr(XXString, "ArticleCount") > 0



'NNTP1.LastArticle
    
    



'--------------------------------------------------------
On Error Resume Next
File = App.Path + "\USENET POSTS -- LastArtical -- " + NG(TT5) + ".txt"
f = FreeFile()
Open File For Input As #f
    
    'CHK THIS THOU
    Line Input #f, V: LastStart = (V + 1) 'ONE EXTRA SO NOT REPEAT LAST WHEN GET END JOB BACKWARD
Close #f
On Error GoTo 0
'---
ANYADDED = False
NEWWRITEHEADERBATCH = False
'TEST
If MNU_OUTPUT.Checked = vbChecked Then
    NEWWRITEHEADERBATCH = True
End If
'---






'TOTAL ITEMS UPTO NOW ON USENET
EasyRem = NNTP1.LastArticle
Easy = NNTP1.LastArticle
If LastStart = "" Then LastStart = Easy
If LastStart = 0 Then LastStart = Easy


'TEST
If MNU_OUTPUT.Checked = True Then
    'If TT5 = 1 Then
        LastStart = (NNTP1.LastArticle - 1) - 100
    'Else
    '     LastStart = (NNTP1.LastArticle - 1) - 10
    'End If

End If

'TOTAL ITEMS UPTO NOW ON USENET
Label6 = Trim(Str(NNTP1.LastArticle))

tt = 0

Do
    TextDebug = ""
    TextDebug2 = ""
    XXString = ""
    DoEvents
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    
    'CURRENT NOW ITEM BEING SCANNED
    Label5 = Trim(Str(Easy))
    Call cmdHeader_Item_Click
    Easy = Easy - 1

    Do
        Call Sleep(10)
        DoEvents
        
        If InStr(XXString, "Disconnected") > 0 Or InStr(XXString, "Not connected") > 0 Then
            outhere = True
            XXString = ""
            txtServer = "news.btinternet.com"
            Call cmdConnect_Click
            DONEARTICALORHEADER = False
            Do
                Call Sleep(10)
                DoEvents
            
                'WAS THAT USED
                'NEW CODE TIMEOUTS - 28-01-2010
                'TEMPORARY WHEN HAEVY TIMEOUTS WHEN CODING
                'If InStr(XXString, "Connecting") > 0 Then Exit Do
                'CURED USING EASY + 1 BELOW WAS IN QUESTION
            
            
            Loop Until DONEARTICALORHEADER = True

            

            'TIME HARDLY EVER USED HERE MAYBE JUST ON RECONNECT
            Text3x = "Date ------------ " + DATEXX2
            Text3FX = "Date My Time - " + DATEXXMyTime
            'WE SHALL USE THIS TIME AS TIME TO END AT


            'YES THIS RECONNECT CODE
            Call Sleep(500)

            Call cmdSelect_Click
            Do
                Call Sleep(10)
                DoEvents
            Loop Until InStr(XXString, "ArticleCount") > 0
            
            Call Sleep(500)
            
            EasyRem = NNTP1.LastArticle
            Easy = NNTP1.LastArticle
            
            'REDUNDANT CODE THIS PROG ALWAYS RUNS UNLESS A NEW GROUP ADDED
            'ALSO RAN AT BEINGING LOPP
            If LastStart = "" Then LastStart = Easy
            If LastStart = 0 Then LastStart = Easy

            'TOTAL ITEMS UPTO NOW ON USENET
            Label6 = Trim(Str(NNTP1.LastArticle))
        End If
        
        If Text2 <> "" And InStrRev(Text2, ".") > Len(Text2) - 4 And InStr(Text2, "Xref: ") = 0 Then Easy = Easy - 1: falseheader = True: Exit Do
        If InStr(TextDebug, NG(TT5)) > 0 Then Exit Do
    Loop Until InStr(TextDebug, ".") > 0 And InStr(Text2, "Xref: ") > 0 Or InStr(Text2, "423 no such article") > 0 Or outhere = True
            
    Text3x = "Date ------------ " + DATEXX2
    Text3FX = "Date My Time " + DATEXXMyTime
    outhere = False
    GODO = 0
    
    For YT = 1 To UBound(HH)
        'CHK HERE
        'If Val(Label5) < LastStart - 1 Then
        If Val(Label5) > LastStart Then

            If InStr(LCase(Text2), HH(YT)) > 0 And InStr(LCase(Text2), "from:") > 0 Then
                GODO = 1
                Label12(YT) = Str(Val(Label12(YT)) + 1) + " "

                QQI = Replace(Replace(DATEXXMyTime, "--", "-"), ",", "")
                For RX4 = 1 To 7
                    QQI = Replace(QQI, Format(Now + RX4, "DDD"), Mid(Format(Now + RX4, "DDD"), 1, 2))
                Next
                

                'NEED GET DATE TIME SERIAL THIS AND CHK AGAINST ONE THERE
                'QQI
                
                
                If Label14(YT) = "" Or GGG(YT) = 0 Then
                    GGG(YT) = 1
                    Label14(YT) = QQI
                Else
                    Label13(YT) = QQI
                End If
            
                Call UserReSizeLabels
            End If
        End If
    Next
    
    If falseheader = True Then falseheader = False: GODO = 0
    
    If GODO = 1 Then
        DONEARTICALORHEADER = False
        TextDebug = ""
        Text3 = ""
        Text4 = ""
        Text5 = ""
        Text6 = ""
        Call cmdArticle_Click

        Do
            'FREQ RUN YES
            DoEvents
            Call Sleep(10)
            DoEvents
            
            If InStr(Text3, "412 no group selected") > 0 Then
                Call cmdSelect_Click
                Do
                    'FREQ RUN YES
                    Call Sleep(10)
                    DoEvents
                Loop Until InStr(Text1, "ArticleCount") > 0
                'CHK HERE
                Call Sleep(50)
            End If
            
        Loop Until DONEARTICALORHEADER = True Or InStr(TextDebug, "time out") > 0

        Call Sleep(10)
    
        tad = 0
        If InStr(XXString, "RC=3 (Busy)") > 0 Then
            
            'NOT SURE LEAVE OUT -- OR GO  -1 TO GET ONE MISSED IF WAS BUSY
            'IF LEAVE OUT BECAUSE DONT GO DOWN TO NEXT IF NOT SUCCESS ANYHOW
            Easy = Easy + 1
            tad = 1
        End If
    
        If tad = 0 Then
            Text3 = "## -- " + NG(TT5) + vbCrLf + Text3
            Text3 = Mid(Text3, 1, Len(Text3) - 4)
            ff2 = "## ---------------------------------------------" + vbCrLf + vbCrLf + vbCrLf
            ww = InStr(Text3, ff2)
            If ww > 0 Then
                Text3 = Mid(Text3, 1, (ww + Len(ff2)) - 3) + Mid(Text3, ww + Len(ff2))
            End If
        
            L2(TT5 - 1) = Val(L2(TT5 - 1)) + 1
        
            'CHK HERE
            If Val(Label5) > LastStart Then
            
                ANYADDED = True
            
                A1 = App.Path + "\TxTPost.txt"
                FileInUseDelay (A1)
                f = FreeFile()
                Open A1 For Append As #f
                Print #f, Text3
                Print #f, "-----------------------------------------------------------------------------#" + vbCrLf
                Print #f, "-----------------------------------------------------------------------------#"
                Close #f
        
                A1 = App.Path + "\USENET ONE BATCH.txt"
                FileInUseDelay (A1)
                f = FreeFile()
                Open A1 For Append As #f
                Print #f, Text3
                Print #f, "-----------------------------------------------------------------------------#" + vbCrLf
                Print #f, "-----------------------------------------------------------------------------#"
                Close #f
        
            
                A1 = App.Path + "\USENET ONE BATCH - " + batchnow + ".txt"
                FileInUseDelay (A1)
                f = FreeFile()
                Open A1 For Append As #f
                Print #f, Text3
                Print #f, "-----------------------------------------------------------------------------#" + vbCrLf
                Print #f, "-----------------------------------------------------------------------------#"
                Close #f
        
                A1 = App.Path + "\USENET ONE BATCH - " + batchnow + " ORIGINAL.txt"
                FileInUseDelay (A1)
                f = FreeFile()
                Open A1 For Append As #f
                Print #f, Text3
                Print #f, "-----------------------------------------------------------------------------#" + vbCrLf
                Print #f, "-----------------------------------------------------------------------------#"
                Close #f
        
        
                A1 = App.Path + "\TxTPost -- " + txtNewsgroup + ".txt"
                FileInUseDelay (A1)
                f = FreeFile()
                Open A1 For Append As #f
                Print #f, Text3
                Print #f, "-----------------------------------------------------------------------------#" + vbCrLf
                Print #f, "-----------------------------------------------------------------------------#"
                Close #f
        
                A1 = App.Path + "\TxTPost -- " + txtNewsgroup + " ORIGINAL.txt"
                FileInUseDelay (A1)
                f = FreeFile()
                Open A1 For Append As #f
                Print #f, Text3
                Print #f, "-----------------------------------------------------------------------------#" + vbCrLf
                Print #f, "-----------------------------------------------------------------------------#"
                Close #f
            End If
        End If
    End If
    
'HOW FAR BACK DOWN TO ITEM TO GO TO IN REVERSE FOR LOGG APPENDS
Label7 = Trim(Str((LastStart)))

'THE EASY NUMBER HOW MANY LEFT TO DO COUNTS DOWN
Label21 = Trim(Str(Abs(Val(Label7 + 1) - Val(Label5))))

If L2(TT5 + 4) = "" Then L2(TT5 + 4) = Label21
   
    
On Error Resume Next
PBar1.Max = LastStart + 1
PBar1.Min = LastStart
PBar1.Max = EasyRem

PXX = (PBar1.Max - Easy) + PBar1.Min '- (PBar1.Max - PBar1.Min)
If PBARR = False Then
    PBARR = True
    PBar1.Value = PXX
End If

'IS THIS RIGHT CHK PBAR

If PXX > PBar1.Value Then
PBar1.Value = PXX
End If

OPXX = PXX

On Error GoTo 0


TTH4 = "## 00 ##" + Str(Easy)


Loop Until Easy < LastStart

PBar1.Min = 1
PBar1.Max = 100
PBar1.Value = 100



'THE EASY NUMBER HOW MANY LEFT TO DO COUNTS DOWN
Label21 = "0"

If MNU_OUTPUT.Checked = False Then

    File = App.Path + "\USENET POSTS -- LastArtical -- " + NG(TT5) + ".txt"
    FileInUseDelay (File)
    f = FreeFile()
    Open File For Output As #f
        'TOTAL ITEMS UPTO NOW ON USENET
        Print #f, Str(Val(Label6))
    Close #f

End If

'TEST
If MNU_OUTPUT.Checked = True Then
    ANYADDED = False
End If

If ANYADDED = True Then

    A1 = App.Path + "\TxTPost-BLOCK.txt"
    FileInUseDelay (A1)
    f = FreeFile()
    Open A1 For Output As #f
    Print #f, "----------------------------------------------------------------------------------#"
    Print #f, "### Process Batch Ended---" + Format$(Now, "DDD DD-MMM-YY HH:MM:SSa/p")
    Print #f, "----------------------------------------------------------------------------------#"
    Close #f


    'A1 = App.Path + "\TxTPost.txt"
    'A2 = App.Path + "\USENET POSTS TOTAL.txt"
    'Call JoinFiles(A1, A2)
    'A1 = App.Path + "\TxTPost-BLOCK.txt"
    'A2 = App.Path + "\USENET POSTS TOTAL.txt"
    'Call JoinFiles(A1, A2)
    '----------

    A1 = App.Path + "\TxTPost -- " + txtNewsgroup + ".txt"
    A2 = App.Path + "\USENET POSTS -- " + txtNewsgroup + ".txt"
    Call JoinFiles(A1, A2)

    A1 = App.Path + "\TxTPost-BLOCK.txt"
    A2 = App.Path + "\USENET POSTS -- " + txtNewsgroup + ".txt"
    Call JoinFiles(A1, A2)
    
    '----------

    A1 = App.Path + "\TxTPost -- " + txtNewsgroup + ".txt"
    A2 = App.Path + "\USENET POSTS -- " + txtNewsgroup + " ORIGINAL.txt"
    Call JoinFiles(A1, A2)

    A1 = App.Path + "\TxTPost-BLOCK.txt"
    A2 = App.Path + "\USENET POSTS -- " + txtNewsgroup + " ORIGINAL.txt"
    Call JoinFiles(A1, A2)
    
    '----------

    A1 = App.Path + "\TxTPost-BLOCK.txt"
    A2 = App.Path + "\USENET ONE BATCH.txt"
    Call JoinFiles(A1, A2)

    'BATCH NOW MAY NEED CHK FOR EMPTY STRING SKIP
    If Trim(batchnow) <> "" Then
        'MsgBox "EMPTY STRING ALERT"

        A1 = App.Path + "\TxTPost-BLOCK.txt"
        A2 = App.Path + "\USENET ONE BATCH - " + batchnow + ".txt"
        Call JoinFiles(A1, A2)
    End If
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CopyFile App.Path + "\USENET POSTS -- " + txtNewsgroup + " ORIGINAL.txt", App.Path + "\LAST_BATCH\USENET POSTS -- " + txtNewsgroup + " ORIGINAL.txt"

End If

Next

labInfo = "EndE..."
labInfo.BackColor = QBColor(10)
labInfo = "EndE..."
DoEvents

Me.WindowState = 0

'Disconnecting from a newsserver
Msg "Disconnecting..."
NNTP1.Disconnect



End Sub

Private Sub Form_Load()
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Main.Show
    Me.BackColor = &H8000000F
    DoEvents
    
    NowStart = Now
    
    On Error Resume Next
    Kill App.Path + "\USENET ONE BATCH.txt"
    On Error GoTo 0
        
    ReDim HH(30)
    
    YT2 = 0
    YT2 = YT2 + 1: HH(YT2) = "Borg@gone.com"
    'YT2 = YT2 + 1: HH(YT2) = "Matt.Lan@btinternet.com"
    YT2 = YT2 + 1: HH(YT2) = "Patricia.Aldoraz@gmail.com"
    
    YT2 = YT2 + 1: HH(YT2) = "RoidsRim"
    
    YT2 = YT2 + 1: HH(YT2) = "anon@no.email" 'Kadaitcha Man
    
    YT2 = YT2 + 1: HH(YT2) = "Gellie618@webtv.net"
    YT2 = YT2 + 1: HH(YT2) = "Rev. 11D Meow!"
    YT2 = YT2 + 1: HH(YT2) = "Chessucat"
    YT2 = YT2 + 1: HH(YT2) = "Daniel Urtiz"
    YT2 = YT2 + 1: HH(YT2) = "Anonymous"
    
    YT2 = YT2 + 1: HH(YT2) = "Nomen Nescio"
    YT2 = YT2 + 1: HH(YT2) = "nobody@dizum.com"
    
    YT2 = YT2 + 1: HH(YT2) = """%"""
    
    YT2 = YT2 + 1: HH(YT2) = "jalees@easynet.co.uk"
    YT2 = YT2 + 1: HH(YT2) = "judy"
    
    'MARK Earnest
    YT2 = YT2 + 1: HH(YT2) = "GMEarnest@yahoo.com"
    
    YT2 = YT2 + 1: HH(YT2) = "@WebTv"
    
    
    fs5 = App.Path + "\LAB-VARS-DATES-USERS.txt"
    'Reset
    fr5 = FreeFile
    Open fs5 For Input As #fr5
    For R = 1 To YT2
        
        'IF WE ADD MORE USERS EOF FILE ERROR RESUME USED
        On Error Resume Next
        LX1 = ""
        LX2 = ""
        'Line Input #fr5, NX1
        Line Input #fr5, LX1
        Line Input #fr5, LX2
        On Error GoTo 0
        QQI1 = Replace(Replace(LX1, "--", "-"), ",", "")
        QQI2 = Replace(Replace(LX2, "--", "-"), ",", "")
        For RX4 = 1 To 7
            QQI1 = Replace(QQI1, Format(Now + RX4, "DDD"), Mid(Format(Now + RX4, "DDD"), 1, 2))
            QQI2 = Replace(QQI2, Format(Now + RX4, "DDD"), Mid(Format(Now + RX4, "DDD"), 1, 2))
        Next
        'If R <> 2 Then
        Label13(R).Caption = QQI1
        Label14(R).Caption = QQI2
        'End If
    
    Next
    Close #fr5
    
    
    Call UserReSizeLabels
    
    Me.Refresh
    DoEvents
    
    '---------------------------------------
    'QUICKEXIT = True
    'Exit Sub
    '---------------------------------------
    


End Sub

Sub DEBUGP()

Debug.Print TTH
    
'TEST
If MNU_OUTPUT.Checked = True Then
    Exit Sub
End If
    
'LABEL5 = CURRENT NOW ITEM BEING SCANNED
If (Val(Label5) > LastStart) = False Then Exit Sub

Call NEWHEADERWRITE

If TT5 > UBound(NG) Then Exit Sub

fsx = App.Path + "\## HEADERS LOGG -- " + NG(TT5) + ".txt"
Call FileInUseDelay(fsx)

FR = FreeFile
Open fsx For Append As #FR
    Print #FR, TTH
Close #FR


End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If QUICKEXIT = True Then End

    On Error Resume Next
    'save user-entered data
    
    fs5 = App.Path + "\LAB-VARS-DATES-USERS.txt"
    Call FileInUseDelay(fs5)
    fr5 = FreeFile
    Open fs5 For Output As #fr5
    For R = 1 To YT2
        Print #fr5, Label13(R).Caption
        Print #fr5, Label14(R).Caption
    Next
    Close #fr5
    
    Call FileInUseDelay(File)
    File = App.Path + "\Setup.ini"
    f = FreeFile()
    Open File For Output As #f
        Print #f, txtServer
        Print #f, txtUsername
        Print #f, txtPassword
        Print #f, txtNewsgroup
        Print #f, chkDebug
        Print #f, Format(NowStart, "ddd, dd mmm yyyy -- hh:mm:ssa/p")
    Close #f

    End

End Sub


Sub DATESEXY()

            'PUT
            'DIM UU3 AS DATE
            
            ddate = Trim(Mid(DATEXX, InStr(DATEXX, " ")))

            'TIMEZONES = = +0%00 FROM GMT

            
            ddate = Replace(ddate, "CET", "+0100") 'CET = BELIN ++
            ddate = Replace(ddate, "MST", "-0700") '= MOUNTAIN STANDARD TIME
            ddate = Replace(ddate, "EST", "+0500") '= EASTERN STANDARD  TIME
            ddate = Replace(ddate, "UTC", "+0000") ' UNI TIME
            ddate = Replace(ddate, "GMT", "+0000")
            'DDATE = Replace(DDATE, "", "")
            
            'SORT THIS
            'CHK IT LATER IF BRACKETS
            
            'BST IS CLOCK FORWARD HOUR FROM NORM
            'MUST BE OLD CODE THINK
            'OR LIKE LAST YEAR TILL TEST AGAIN UNLESS ROUTE THROUGH CODE
            ddate = Replace(ddate, "(BST)", "+0100")
            
            If InStr(ddate, "BST") > 0 Then MsgBox "WE ARE CHKING HERE " + vbCrLf + "DDATE = Replace(DDATE, ""(BST)"", ""+0100"")" + vbCrLf + "FOR THIS -- HERE" + vbCrLf + "If InStr(DDATE, ""BST"") > 0 Then " + vbCrLf + "THE NORM IN CODE IS THIS USED WITH BRACKETS" + vbCrLf + "HERE IS DDATE VAR" + vbCrLf + ddate
            UU = ddate
            
            On Error Resume Next
            UU3 = DateValue(ddate) + TimeValue(ddate)
            On Error GoTo 0
            UU4 = UU3
            
            If InStr(ddate, "-") > 0 Then
                UU = Mid(ddate, 1, InStr(ddate, "-") - 2)
                UU1 = Mid(ddate, InStr(ddate, "-") + 1, 2)
                'For R = 1 To 7
                
                If IsNumeric(Mid(ddate, 1, 2)) = False Then UU = Mid(UU, 6)
                UU3 = DateValue(UU) + TimeValue(UU)
                UU4 = UU3
                UU2 = Val(UU1)
                UU3 = UU3 - TimeSerial(UU2, 0, 0)
                RTIME = "+" + Trim(Val(UU1))
            End If
            If InStr(ddate, "+") > 0 Then
                UU = Mid(ddate, 1, InStr(ddate, "+") - 2)
                UU1 = Mid(ddate, InStr(ddate, "+") + 1, 2)
                If IsNumeric(Mid(ddate, 1, 2)) = False Then UU = Mid(UU, 6)
                UU3 = DateValue(UU) + TimeValue(UU)
                UU4 = UU3
                UU2 = Val(UU1)
                UU3 = UU3 + TimeSerial(UU2, 0, 0)
                RTIME = "+" + Trim(Val(UU1))
            End If
        
            DATEXX = Format(UU3, "DDD, DD MMM YYYY HH:MM:SS")
            'WORLD TIME
            DATEXX2 = Format(UU4, "DDD, DD MMM YYYY -- HH:MM:SSa/p")
            
            'MY TIME DISPLAY WITH EXTRA + OR - HOUR FORWARD
            DATEXXMyTime = Format(UU3, "DDD, DD MMM YYYY -- HH:MM:SSa/p") + " " + RTIME

End Sub

Private Sub L2_Click(Index As Integer)
fsx = App.Path + "\## HEADERS LOGG -- " + NG(Index - 4) + ".txt"

If Dir(fsx) = "" Then MsgBox "No Header File Yet This Group": Exit Sub

Shell "notepad " + fsx, vbNormalFocus

End Sub

Private Sub Label21_Click()
'Label21
End Sub

Private Sub Label7_Click()
'Label7
End Sub

Private Sub MNU_OUTPUT_Click()

MNU_OUTPUT.Checked = Not MNU_OUTPUT.Checked

End Sub

Private Sub Mnu_VB_Click()
Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
End
End Sub

Private Sub NNTP1_Debugger(Text As String)
    'Print received data when in debugging mode
    'Debug.Print Text
    
    
    TextDebug = Text + vbCrLf
    If Text <> "." Then
        TextDebug2 = Text + vbCrLf
        
        If InStr(Text, "Lines:") > 0 Then
            LinesCount = Val(Mid(Text, 7))
        End If
        
        DIY = True
        
        If TTH4 <> "" Then
            TTH = vbCrLf
            Call DEBUGP
            TTH = TTH4: TTH4 = ""
            Call DEBUGP
        End If
        
        ii = 0
        If InStr(Text, "From:") > 0 Then
            ii = 1
            
            TTH = "## 00 ## " + Text
            Call DEBUGP
            


        End If
        
        If InStr(Text, "Date:") = 1 Then
            'WE NEED DDATE DIM
            'FOR NOW

            ii = 1: DATEXX = Text
                    
            For RS5 = 1 To 7
                DATEXX = Replace(DATEXX, Format(Now + RS5, "DDDD"), "")
                DATEXX = Replace(DATEXX, Format(Now + RS5, "DDD"), "")
            Next
            For RS5 = 1 To 12
                IIS = DateSerial(2010, RS5, 1)
                DATEXX = Replace(DATEXX, Format(IIS, "MMMM"), Format(IIS, "MMM"))
            Next
            DATEXX = Replace(DATEXX, ",", "")
            
            
             
            Call DATESEXY
            If ddate <> "" And DATEXX = "" Then Stop
            TEXTX3 = DATEXX + " -- " + ddate
            
            TTH = "## 03 ## " + Text
            Call DEBUGP
            
            TTH = "## 04 ## UK DATE #-# " + DATEXX
            Call DEBUGP
            
        
        End If
        If InStr(Text, "Subject:") > 0 Then
            ii = 1
            TTH = "## 01 ## " + Text
            Call DEBUGP
        End If
        
        If InStr(Text, "Newsgroups:") > 0 Then
            ii = 1
            TTH = "## 02 ## " + Text
            Call DEBUGP
        End If
        
        If InStr(Text, "NNTP -Posting - Date:") > 0 Then
            ii = 2: IX = 2
        End If
        
        If InStr(Text, "Message-ID:") > 0 Then
            ii = 2: IX = 1
            TTH = "## 05 ## " + Text
            Call DEBUGP
        End If
        
        If InStr(Text, "In-Reply-To:") > 0 Then
            ii = 2: IX = 2
            'TTH= "## 06 ## " + Text
            'Call DebugP
        End If
        
        If InStr(Text, "References:") > 0 Then
            ii = 2: IX = 2
            
            TTH = "## 07 ## " + Text
            Call DEBUGP
        End If
        
        
        
        If InStr(Text, "Newsreader:") > 0 Then
            ii = 3
            TTH = "## 08 ## " + Text
            Call DEBUGP

        End If
        
        If InStr(Text, "Lines:") > 0 Then ii = 3
        If InStr(Text, "Bytes:") > 0 Then ii = 3
        
        If ii = 1 Then
            Text4 = Text4 + Text + vbCrLf
            DIY = False
        End If
        
        If ii = 2 Then
            If IX = 1 Then Text5 = Text + Text5 + vbCrLf
            If IX = 2 Then Text5 = Text5 + Text + vbCrLf
            DIY = False
        End If
        
        If ii = 3 Then
            Text6 = Text6 + Text + vbCrLf
            DIY = False
        End If
        
        If DIY = True Then Text3 = Text3 + Text + vbCrLf
        
        If InStr(Text, "Xref: ") > 0 Then
            LinesCount2 = 0
            TXR = "--------------------------" + vbCrLf
            
            Text3 = Text3 + TXR + Text5 + TXR + Text6 + TXR + Text4
            Text3 = Text3 + vbCrLf + "## ---------------------------------------------"
            Text3 = Text3 + vbCrLf + "## USENET - Message------------------"
            Text3 = Text3 + vbCrLf + "## ---------------------------------------------" + vbCrLf
        End If
    End If
    
    LinesCount2 = LinesCount2 + 1
    
    
    
    
    Text2 = Text2 + Text + vbCrLf
    Text2.SelStart = Len(Text2)

End Sub

Private Sub NNTP1_DoneConnect(Rc As TNNTPRc)
    Msg "Connected with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    Msg "Server Msg=" + NNTP1.Hostinfo
    DONEARTICALORHEADER = True
End Sub

Private Sub NNTP1_DoneDisconnect(Rc As TNNTPRc)
    Msg "Disconnected with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    Close #1
    DONEARTICALORHEADER = True
End Sub
Public Sub Msg(String1 As String)
    XXString = XXString + String1
    With Text1
        If Len(.Text) > 16384 Then
            .Text = Mid(.Text, 2000)
        End If
        If .Text = "" Then
            .Text = String1 + vbCrLf
        Else
            .Text = .Text + Right(String1, 16384) + vbCrLf
        End If
        .SelStart = Len(.Text)
    End With
End Sub

Private Sub NNTP1_DoneGetArticle(Rc As TNNTPRc)
    Msg "Article got with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    DONEARTICALORHEADER = True
End Sub

Private Sub NNTP1_DoneGetBody(Rc As TNNTPRc)
    On Error Resume Next
    Msg "Body got with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    Msg "Speed is " & CLng(NNTP1.Bps)
End Sub

Private Sub NNTP1_DoneGetHeader(Rc As TNNTPRc)
    Msg "Header got with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    Msg "Posting date was " & NNTP1.DateTime
    DONEARTICALORHEADER = True
End Sub

Private Sub NNTP1_DoneGetList(Rc As TNNTPRc)
    Msg "List of newsgroups received with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    labInfo.Caption = CStr(NNTP1.DataCount) + " newsgroups received"
End Sub

Private Sub NNTP1_DoneGetStat(Rc As TNNTPRc)
    Msg "Stat got with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
End Sub

Private Sub NNTP1_DoneGetXover(Rc As TNNTPRc)
    Msg "Article headers got with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    labInfo = CStr(NNTP1.DataCount) + " headers received"
End Sub

Private Sub NNTP1_DoneGotoFirst(Rc As TNNTPRc)
    Msg "Goto last article with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
End Sub

Private Sub NNTP1_DoneGotoLast(Rc As TNNTPRc)
    Msg "Goto last article with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
End Sub

Private Sub NNTP1_DoneGotoNext(Rc As TNNTPRc)
    Msg "Goto next article with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
End Sub


Private Sub NNTP1_DoneList(Rc As TNNTPRc)
    Msg "List got with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    Msg NNTP1.DataCount & " newsgroups received"
End Sub

Private Sub NNTP1_DoneSelectGroup(Rc As TNNTPRc)
    Msg "Newsgroup selected with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    Msg "FirstArticle=" & NNTP1.FirstArticle
    Msg "LastArticle=" & NNTP1.LastArticle
    Msg "ArticleCount=" & NNTP1.ArticleCount
    DONEARTICALORHEADER = True
End Sub

Private Sub NNTP1_DoneXover(Rc As TNNTPRc)
    Msg "Xover done with RC=" + CStr(Rc) + " (" + NNTP1.RcString(Rc) + ")"
    NNTP1.DataIndex = 1
    Msg "FirstHeader=" + NNTP1.Subject
    NNTP1.DataIndex = NNTP1.DataCount
    Msg "LastHeader=" + NNTP1.Subject
End Sub

Private Sub NNTP1_DoneSend(Rc As TNNTPRc)
    Msg NNTP1.Text
End Sub

Private Sub NNTP1_Receive(Count As Long)
    labInfo.Caption = CStr(Count) & " lines"
End Sub

Private Sub Text1_Change()
A = A
End Sub

Public Function JoinFiles(Source1 As String, Source2 As String) As Boolean
    
    'Join 1 in front of 2
    
    'On Error GoTo errorhandler
    
    FileInUseDelay (Source1)
    Open Source1 For Binary Access Read As #1
    FileInUseDelay (Source2)
    Open Source2 For Binary Access Read As #2
   
   
   Buffer1 = Space(LOF(1))
   Get #1, , Buffer1
   
   Buffer2 = Space(LOF(2))
   Get #2, , Buffer2
   
   Close #1, #2
   
    FileInUseDelay (Source2)
    Kill Source2
    FileInUseDelay (Source2)
    Open Source2 For Binary Access Write As #2
    Put #2, , Buffer1
    Put #2, , Buffer2
    Close #2
   
   JoinFiles = True

'Text1 = ""

Exit Function
ErrorHandler:
    MsgBox "Error Join File" + vbCrLf + Source1 + vbCrLf + Source2
    Close #1
    Close #2
    Close #3

End Function








Function FileInUse(ByVal strFilePath As String) As Boolean
  
  Dim hFile As Long
  Dim FileInfo  As OFSTRUCT
  
  strFilePath = Trim(strFilePath)
  If strFilePath = "" Then Exit Function
  If Dir(strFilePath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then Exit Function
  If Right(strFilePath, 1) <> Chr(0) Then strFilePath = strFilePath & Chr(0)
  
  FileInfo.cBytes = Len(FileInfo)
  hFile = OpenFile(strFilePath, FileInfo, OF_SHARE_EXCLUSIVE)
  If hFile = -1 And Err.LastDllError = 32 Then
    FileInUse = True
  Else
    CloseHandle hFile
  End If
  
End Function

Sub FileInUseDelay(Tx)
        
    qq = Now + TimeSerial(0, 5, 0)
    Do
        'DoEvents
        ii = FileInUse(Tx)
        If ii = True Then Sleep (200)
    Loop Until ii = False Or qq < Now
    
    If ii = True Then
        MsgBox "Trouble File Stuck In Use " + vbCrLf + Tx + vbCrLf + "Longer than 5 Mins"
        Stop
    End If
End Sub

Private Sub Text3x_Change()
'Text3x
End Sub
