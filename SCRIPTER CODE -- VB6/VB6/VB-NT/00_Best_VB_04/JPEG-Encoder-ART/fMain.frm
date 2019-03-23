VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form fmain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   Caption         =   "JPEG Encoder -- Shockingly Good Photo's --< The One >--"
   ClientHeight    =   7950
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   11400
   Begin VB.Timer Timer_TRIP_WIRE 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7155
      Top             =   6120
   End
   Begin VB.Timer Timer6 
      Interval        =   10000
      Left            =   5910
      Top             =   1995
   End
   Begin MSComctlLib.ProgressBar PBarDUPE 
      Height          =   210
      Left            =   2580
      TabIndex        =   211
      Top             =   5580
      Visible         =   0   'False
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   10470
      TabIndex        =   209
      Top             =   3945
      Visible         =   0   'False
      Width           =   825
      Begin VB.Label LBB5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "The 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   360
         Left            =   -60
         TabIndex        =   210
         Top             =   105
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin VB.DirListBox Dir3 
      Height          =   405
      Left            =   9870
      TabIndex        =   208
      Top             =   3315
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.DirListBox Dir2 
      Height          =   405
      Left            =   7965
      TabIndex        =   207
      Top             =   3570
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   9375
      TabIndex        =   206
      Top             =   705
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.Timer zzCheckTimer 
      Interval        =   10000
      Left            =   12360
      Top             =   6345
   End
   Begin VB.Timer FileInfoTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11925
      Top             =   6105
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   690
      Left            =   5775
      ScaleHeight     =   690
      ScaleWidth      =   1125
      TabIndex        =   203
      Top             =   3585
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   11520
      Top             =   6770
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   10680
      Top             =   6765
   End
   Begin VB.Timer Timer7 
      Interval        =   10
      Left            =   11940
      Top             =   6770
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   11040
      TabIndex        =   198
      Top             =   2505
      Visible         =   0   'False
      Width           =   2265
      Begin VB.Label LBB4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Fri 07-Aug-2009 10:10:10"
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
         Height          =   570
         Left            =   90
         TabIndex        =   199
         Top             =   -180
         Visible         =   0   'False
         Width           =   2040
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   11265
      TabIndex        =   196
      Top             =   3255
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Label LBB3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         Caption         =   "The FORCE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   60
         TabIndex        =   197
         ToolTipText     =   "CLIPBOARD FILE INFO"
         Top             =   135
         Visible         =   0   'False
         Width           =   1350
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   11640
      TabIndex        =   194
      Top             =   3840
      Visible         =   0   'False
      Width           =   1125
      Begin VB.Label LBB2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00008000&
         Caption         =   "Controls"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   -120
         TabIndex        =   195
         Top             =   75
         Visible         =   0   'False
         Width           =   1290
      End
   End
   Begin VB.Timer TimerFlash3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11520
      Top             =   6340
   End
   Begin VB.Timer TimerColorCycle 
      Interval        =   40
      Left            =   10260
      Top             =   6340
   End
   Begin VB.Timer TimerFlash2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11100
      Top             =   6340
   End
   Begin VB.Timer TimerFlash1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10680
      Top             =   6340
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   5715
      TabIndex        =   185
      Top             =   4515
      Visible         =   0   'False
      Width           =   8205
      Begin MSComctlLib.ProgressBar PBar3 
         Height          =   345
         Left            =   -15
         TabIndex        =   186
         Top             =   660
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label LBB1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         Caption         =   "The 1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   735
         Left            =   3345
         TabIndex        =   187
         Top             =   -15
         Width           =   1515
      End
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   11100
      Top             =   6770
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   915
      Left            =   345
      TabIndex        =   181
      ToolTipText     =   "Load In IceView"
      Top             =   2460
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1614
      _Version        =   393217
      BackColor       =   8388608
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"fMain.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   240
      Left            =   2625
      TabIndex        =   25
      Top             =   6330
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   423
      _Version        =   393216
      Max             =   50
      TickStyle       =   3
   End
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   210
      Left            =   2595
      TabIndex        =   16
      Top             =   5835
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer2 
      Interval        =   800
      Left            =   10260
      Top             =   6770
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   9855
      Top             =   6770
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      DrawMode        =   11  'Not Xor Pen
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   195
      ScaleHeight     =   2175
      ScaleWidth      =   3135
      TabIndex        =   1
      Top             =   105
      Width           =   3135
      Begin VB.VScrollBar VScroll1 
         Height          =   855
         Left            =   2580
         TabIndex        =   205
         Top             =   735
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         Left            =   1620
         TabIndex        =   204
         Top             =   1695
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin MSComctlLib.ProgressBar PBar2 
      Height          =   195
      Left            =   2595
      TabIndex        =   17
      Top             =   6075
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin RichTextLib.RichTextBox RTB2 
      Height          =   1455
      Left            =   315
      TabIndex        =   182
      ToolTipText     =   "Thumbs"
      Top             =   3465
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   2566
      _Version        =   393217
      BackColor       =   8388608
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"fMain.frx":04BC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTB3 
      Height          =   1335
      Left            =   315
      TabIndex        =   183
      ToolTipText     =   "Explorer Select"
      Top             =   5130
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   2355
      _Version        =   393217
      BackColor       =   8388608
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"fMain.frx":0536
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label RFLb 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Randomiser Forward"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5280
      TabIndex        =   338
      Top             =   7095
      Width           =   2280
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   207
      Left            =   10365
      TabIndex        =   337
      Top             =   75
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   206
      Left            =   10485
      TabIndex        =   336
      Top             =   255
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   205
      Left            =   10395
      TabIndex        =   335
      Top             =   255
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   204
      Left            =   10485
      TabIndex        =   334
      Top             =   195
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   203
      Left            =   10425
      TabIndex        =   333
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   202
      Left            =   10440
      TabIndex        =   332
      Top             =   180
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   201
      Left            =   10410
      TabIndex        =   331
      Top             =   180
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   200
      Left            =   10545
      TabIndex        =   330
      Top             =   210
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   199
      Left            =   10380
      TabIndex        =   329
      Top             =   180
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   198
      Left            =   10560
      TabIndex        =   328
      Top             =   195
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   197
      Left            =   10590
      TabIndex        =   327
      Top             =   150
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   196
      Left            =   10710
      TabIndex        =   326
      Top             =   330
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   195
      Left            =   10620
      TabIndex        =   325
      Top             =   330
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   194
      Left            =   10710
      TabIndex        =   324
      Top             =   270
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   193
      Left            =   10650
      TabIndex        =   323
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   191
      Left            =   10665
      TabIndex        =   322
      Top             =   255
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   190
      Left            =   10635
      TabIndex        =   321
      Top             =   255
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   189
      Left            =   10770
      TabIndex        =   320
      Top             =   285
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   188
      Left            =   10605
      TabIndex        =   319
      Top             =   255
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   187
      Left            =   10785
      TabIndex        =   318
      Top             =   270
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   186
      Left            =   10485
      TabIndex        =   317
      Top             =   210
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   185
      Left            =   10605
      TabIndex        =   316
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   184
      Left            =   10515
      TabIndex        =   315
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   183
      Left            =   10605
      TabIndex        =   314
      Top             =   330
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   182
      Left            =   10545
      TabIndex        =   313
      Top             =   375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   181
      Left            =   10560
      TabIndex        =   312
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   180
      Left            =   10530
      TabIndex        =   311
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   179
      Left            =   10665
      TabIndex        =   310
      Top             =   345
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   178
      Left            =   10500
      TabIndex        =   309
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   177
      Left            =   10680
      TabIndex        =   308
      Top             =   330
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   176
      Left            =   10710
      TabIndex        =   307
      Top             =   285
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   175
      Left            =   10830
      TabIndex        =   306
      Top             =   465
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   174
      Left            =   10740
      TabIndex        =   305
      Top             =   465
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   173
      Left            =   10830
      TabIndex        =   304
      Top             =   405
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   172
      Left            =   10770
      TabIndex        =   303
      Top             =   450
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   171
      Left            =   10785
      TabIndex        =   302
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   170
      Left            =   10755
      TabIndex        =   301
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   169
      Left            =   10890
      TabIndex        =   300
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   168
      Left            =   10725
      TabIndex        =   299
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   167
      Left            =   10905
      TabIndex        =   298
      Top             =   375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   166
      Left            =   10755
      TabIndex        =   297
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   165
      Left            =   10665
      TabIndex        =   296
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   164
      Left            =   10695
      TabIndex        =   295
      Top             =   345
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   163
      Left            =   10650
      TabIndex        =   294
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   162
      Left            =   10560
      TabIndex        =   293
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   161
      Left            =   10650
      TabIndex        =   292
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   160
      Left            =   10590
      TabIndex        =   291
      Top             =   405
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   159
      Left            =   10605
      TabIndex        =   290
      Top             =   345
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   158
      Left            =   10575
      TabIndex        =   289
      Top             =   345
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   157
      Left            =   10710
      TabIndex        =   288
      Top             =   375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   156
      Left            =   10545
      TabIndex        =   287
      Top             =   345
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   155
      Left            =   10725
      TabIndex        =   286
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   154
      Left            =   10875
      TabIndex        =   285
      Top             =   435
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   153
      Left            =   10830
      TabIndex        =   284
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   152
      Left            =   10800
      TabIndex        =   283
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   151
      Left            =   10770
      TabIndex        =   282
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   150
      Left            =   10845
      TabIndex        =   281
      Top             =   330
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   149
      Left            =   10770
      TabIndex        =   280
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   148
      Left            =   11025
      TabIndex        =   279
      Top             =   435
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   147
      Left            =   10800
      TabIndex        =   278
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   146
      Left            =   10830
      TabIndex        =   277
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   145
      Left            =   10815
      TabIndex        =   276
      Top             =   450
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   144
      Left            =   10875
      TabIndex        =   275
      Top             =   405
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   143
      Left            =   10785
      TabIndex        =   274
      Top             =   465
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   142
      Left            =   10875
      TabIndex        =   273
      Top             =   465
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   141
      Left            =   10725
      TabIndex        =   272
      Top             =   330
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   140
      Left            =   10545
      TabIndex        =   271
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   139
      Left            =   10710
      TabIndex        =   270
      Top             =   345
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   138
      Left            =   10575
      TabIndex        =   269
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   137
      Left            =   10605
      TabIndex        =   268
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   136
      Left            =   10590
      TabIndex        =   267
      Top             =   375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   135
      Left            =   10650
      TabIndex        =   266
      Top             =   330
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   134
      Left            =   10560
      TabIndex        =   265
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   133
      Left            =   10650
      TabIndex        =   264
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   132
      Left            =   10695
      TabIndex        =   263
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   131
      Left            =   10665
      TabIndex        =   262
      Top             =   330
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   130
      Left            =   10755
      TabIndex        =   261
      Top             =   330
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   129
      Left            =   10725
      TabIndex        =   260
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   128
      Left            =   10890
      TabIndex        =   259
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   127
      Left            =   10755
      TabIndex        =   258
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   126
      Left            =   10785
      TabIndex        =   257
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   125
      Left            =   10770
      TabIndex        =   256
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   124
      Left            =   10830
      TabIndex        =   255
      Top             =   375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   123
      Left            =   10740
      TabIndex        =   254
      Top             =   435
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   122
      Left            =   10830
      TabIndex        =   253
      Top             =   435
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   121
      Left            =   10710
      TabIndex        =   252
      Top             =   255
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   120
      Left            =   10680
      TabIndex        =   251
      Top             =   300
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   119
      Left            =   10500
      TabIndex        =   250
      Top             =   285
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   118
      Left            =   10665
      TabIndex        =   249
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   117
      Left            =   10530
      TabIndex        =   248
      Top             =   285
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   116
      Left            =   10560
      TabIndex        =   247
      Top             =   285
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   115
      Left            =   10545
      TabIndex        =   246
      Top             =   345
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   114
      Left            =   10605
      TabIndex        =   245
      Top             =   300
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   113
      Left            =   10515
      TabIndex        =   244
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   112
      Left            =   10605
      TabIndex        =   243
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   111
      Left            =   10785
      TabIndex        =   242
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   110
      Left            =   10770
      TabIndex        =   241
      Top             =   255
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   109
      Left            =   10650
      TabIndex        =   240
      Top             =   285
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   108
      Left            =   10710
      TabIndex        =   239
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   107
      Left            =   10620
      TabIndex        =   238
      Top             =   300
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   106
      Left            =   10875
      TabIndex        =   237
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   77
      Left            =   13005
      TabIndex        =   236
      Top             =   1185
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   76
      Left            =   13005
      TabIndex        =   235
      Top             =   1455
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   75
      Left            =   13125
      TabIndex        =   234
      Top             =   1275
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   74
      Left            =   13215
      TabIndex        =   233
      Top             =   1515
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   73
      Left            =   12825
      TabIndex        =   232
      Top             =   1200
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   72
      Left            =   13140
      TabIndex        =   231
      Top             =   1575
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   71
      Left            =   13185
      TabIndex        =   230
      Top             =   1905
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   70
      Left            =   12915
      TabIndex        =   229
      Top             =   1425
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   69
      Left            =   12915
      TabIndex        =   228
      Top             =   1140
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   68
      Left            =   12930
      TabIndex        =   227
      Top             =   1635
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   67
      Left            =   12990
      TabIndex        =   226
      Top             =   1290
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   66
      Left            =   13005
      TabIndex        =   225
      Top             =   1560
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   65
      Left            =   12930
      TabIndex        =   224
      Top             =   1335
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   64
      Left            =   13005
      TabIndex        =   223
      Top             =   1155
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   63
      Left            =   12930
      TabIndex        =   222
      Top             =   1215
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   62
      Left            =   13185
      TabIndex        =   221
      Top             =   1395
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   61
      Left            =   12960
      TabIndex        =   220
      Top             =   1245
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   60
      Left            =   12945
      TabIndex        =   219
      Top             =   1455
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   59
      Left            =   12975
      TabIndex        =   218
      Top             =   1455
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   58
      Left            =   13425
      TabIndex        =   217
      Top             =   1500
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   57
      Left            =   13470
      TabIndex        =   216
      Top             =   1740
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   56
      Left            =   12885
      TabIndex        =   215
      Top             =   1305
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   55
      Left            =   13065
      TabIndex        =   214
      Top             =   1635
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   54
      Left            =   13080
      TabIndex        =   213
      Top             =   1410
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   53
      Left            =   13035
      TabIndex        =   212
      Top             =   1170
      Width           =   405
   End
   Begin VB.Label LBX 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Dimention"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   3
      Left            =   6660
      TabIndex        =   202
      Top             =   2895
      Width           =   2595
      WordWrap        =   -1  'True
   End
   Begin VB.Label LBX 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      Caption         =   "Mon-10-Jun-10 23-11-00"
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
      Height          =   900
      Index           =   2
      Left            =   6555
      TabIndex        =   201
      Top             =   1785
      Width           =   3075
      WordWrap        =   -1  'True
   End
   Begin VB.Label LBX 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "File Size"
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
      Height          =   435
      Index           =   1
      Left            =   6795
      TabIndex        =   200
      Top             =   1095
      Width           =   2025
      WordWrap        =   -1  'True
   End
   Begin VB.Label LBX 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   6825
      TabIndex        =   0
      Top             =   360
      Width           =   1155
      WordWrap        =   -1  'True
   End
   Begin VB.Label LabelCol9 
      BackColor       =   &H00E78A36&
      Caption         =   "Label11"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3075
      TabIndex        =   193
      Top             =   3645
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label LabelCol5 
      BackColor       =   &H00E78A36&
      Caption         =   "Label11"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3075
      TabIndex        =   192
      Top             =   3015
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label LabelCol3 
      BackColor       =   &H007CE736&
      Caption         =   "Label11"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3075
      TabIndex        =   191
      Top             =   2730
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "-Top Left Menu"
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
      Height          =   315
      Index           =   13
      Left            =   3540
      TabIndex        =   190
      Top             =   4995
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "-Top Left Menu"
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
      Height          =   315
      Index           =   12
      Left            =   3555
      TabIndex        =   189
      Top             =   4860
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label LabelCol8 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label11"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3045
      TabIndex        =   188
      Top             =   3300
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "-Top Left Menu"
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
      Height          =   315
      Index           =   7
      Left            =   3555
      TabIndex        =   9
      Top             =   3255
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label LBT1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4965
      TabIndex        =   184
      Top             =   6315
      Visible         =   0   'False
      Width           =   675
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "-Top Left Menu"
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
      Height          =   315
      Index           =   11
      Left            =   3555
      TabIndex        =   14
      Top             =   4545
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "-Top Left Menu"
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
      Height          =   315
      Index           =   10
      Left            =   3555
      TabIndex        =   13
      Top             =   4245
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "-Top Left Menu"
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
      Height          =   315
      Index           =   0
      Left            =   3570
      TabIndex        =   2
      Top             =   1185
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   105
      Left            =   10935
      TabIndex        =   180
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   104
      Left            =   10680
      TabIndex        =   179
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   103
      Left            =   10770
      TabIndex        =   178
      Top             =   300
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   102
      Left            =   10710
      TabIndex        =   177
      Top             =   345
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   101
      Left            =   10830
      TabIndex        =   176
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   100
      Left            =   10845
      TabIndex        =   175
      Top             =   300
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   99
      Left            =   10665
      TabIndex        =   174
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   98
      Left            =   10575
      TabIndex        =   173
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   97
      Left            =   10665
      TabIndex        =   172
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   95
      Left            =   10605
      TabIndex        =   171
      Top             =   405
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   94
      Left            =   10620
      TabIndex        =   170
      Top             =   345
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   93
      Left            =   10590
      TabIndex        =   169
      Top             =   345
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   92
      Left            =   10725
      TabIndex        =   168
      Top             =   375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   91
      Left            =   10560
      TabIndex        =   167
      Top             =   345
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   90
      Left            =   10740
      TabIndex        =   166
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   89
      Left            =   10770
      TabIndex        =   165
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   88
      Left            =   10890
      TabIndex        =   164
      Top             =   495
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   87
      Left            =   10800
      TabIndex        =   163
      Top             =   495
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   86
      Left            =   10890
      TabIndex        =   162
      Top             =   435
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   85
      Left            =   10830
      TabIndex        =   161
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   84
      Left            =   10845
      TabIndex        =   160
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   83
      Left            =   10815
      TabIndex        =   159
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   82
      Left            =   10950
      TabIndex        =   158
      Top             =   450
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   81
      Left            =   10785
      TabIndex        =   157
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   80
      Left            =   11280
      TabIndex        =   156
      Top             =   825
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   79
      Left            =   10815
      TabIndex        =   155
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   78
      Left            =   10725
      TabIndex        =   154
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   77
      Left            =   10755
      TabIndex        =   153
      Top             =   375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   76
      Left            =   10710
      TabIndex        =   152
      Top             =   450
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   75
      Left            =   10620
      TabIndex        =   151
      Top             =   450
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   74
      Left            =   10710
      TabIndex        =   150
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   73
      Left            =   10650
      TabIndex        =   149
      Top             =   435
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   72
      Left            =   10665
      TabIndex        =   148
      Top             =   375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   71
      Left            =   10635
      TabIndex        =   147
      Top             =   375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   70
      Left            =   10770
      TabIndex        =   146
      Top             =   405
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   69
      Left            =   10605
      TabIndex        =   145
      Top             =   375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   68
      Left            =   10785
      TabIndex        =   144
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   67
      Left            =   10935
      TabIndex        =   143
      Top             =   525
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   66
      Left            =   10845
      TabIndex        =   142
      Top             =   525
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   65
      Left            =   10935
      TabIndex        =   141
      Top             =   465
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   64
      Left            =   10875
      TabIndex        =   140
      Top             =   510
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   63
      Left            =   10890
      TabIndex        =   139
      Top             =   450
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   62
      Left            =   10860
      TabIndex        =   138
      Top             =   450
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   61
      Left            =   11085
      TabIndex        =   137
      Top             =   495
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   60
      Left            =   10830
      TabIndex        =   136
      Top             =   450
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   59
      Left            =   11265
      TabIndex        =   135
      Top             =   660
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   58
      Left            =   10905
      TabIndex        =   134
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   57
      Left            =   10830
      TabIndex        =   133
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   56
      Left            =   11160
      TabIndex        =   132
      Top             =   585
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   55
      Left            =   10860
      TabIndex        =   131
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   54
      Left            =   10890
      TabIndex        =   130
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   53
      Left            =   10875
      TabIndex        =   129
      Top             =   540
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   52
      Left            =   10935
      TabIndex        =   128
      Top             =   495
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   51
      Left            =   10845
      TabIndex        =   127
      Top             =   555
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   50
      Left            =   10935
      TabIndex        =   126
      Top             =   555
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   49
      Left            =   10785
      TabIndex        =   125
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   47
      Left            =   10605
      TabIndex        =   124
      Top             =   405
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   46
      Left            =   10770
      TabIndex        =   123
      Top             =   435
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   45
      Left            =   10635
      TabIndex        =   122
      Top             =   405
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   44
      Left            =   10665
      TabIndex        =   121
      Top             =   405
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   43
      Left            =   10650
      TabIndex        =   120
      Top             =   465
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   42
      Left            =   10710
      TabIndex        =   119
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   41
      Left            =   10620
      TabIndex        =   118
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   40
      Left            =   10710
      TabIndex        =   117
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   39
      Left            =   10755
      TabIndex        =   116
      Top             =   405
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   38
      Left            =   10725
      TabIndex        =   115
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   37
      Left            =   10815
      TabIndex        =   114
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   36
      Left            =   10965
      TabIndex        =   113
      Top             =   435
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   35
      Left            =   10785
      TabIndex        =   112
      Top             =   450
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   34
      Left            =   10950
      TabIndex        =   111
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   33
      Left            =   10815
      TabIndex        =   110
      Top             =   450
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   31
      Left            =   10845
      TabIndex        =   109
      Top             =   450
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   30
      Left            =   10830
      TabIndex        =   108
      Top             =   510
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   29
      Left            =   10890
      TabIndex        =   107
      Top             =   465
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   28
      Left            =   10800
      TabIndex        =   106
      Top             =   525
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   27
      Left            =   10890
      TabIndex        =   105
      Top             =   525
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   26
      Left            =   10770
      TabIndex        =   104
      Top             =   345
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   25
      Left            =   10740
      TabIndex        =   103
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   24
      Left            =   10560
      TabIndex        =   102
      Top             =   375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   23
      Left            =   10725
      TabIndex        =   101
      Top             =   405
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   22
      Left            =   10590
      TabIndex        =   100
      Top             =   375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   21
      Left            =   10620
      TabIndex        =   99
      Top             =   375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   20
      Left            =   10605
      TabIndex        =   98
      Top             =   435
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   19
      Left            =   10665
      TabIndex        =   97
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   18
      Left            =   10575
      TabIndex        =   96
      Top             =   450
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   17
      Left            =   10665
      TabIndex        =   95
      Top             =   450
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   15
      Left            =   10545
      TabIndex        =   94
      Top             =   270
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   14
      Left            =   10845
      TabIndex        =   93
      Top             =   330
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   13
      Left            =   10665
      TabIndex        =   92
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   12
      Left            =   10830
      TabIndex        =   91
      Top             =   345
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   11
      Left            =   10695
      TabIndex        =   90
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   10
      Left            =   10725
      TabIndex        =   89
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   9
      Left            =   10710
      TabIndex        =   88
      Top             =   375
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   7
      Left            =   10770
      TabIndex        =   87
      Top             =   330
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   6
      Left            =   10680
      TabIndex        =   86
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   5
      Left            =   10770
      TabIndex        =   85
      Top             =   390
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   3
      Left            =   10650
      TabIndex        =   84
      Top             =   210
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   192
      Left            =   10620
      TabIndex        =   83
      Top             =   255
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   96
      Left            =   10440
      TabIndex        =   82
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   48
      Left            =   10605
      TabIndex        =   81
      Top             =   270
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   32
      Left            =   10470
      TabIndex        =   80
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   16
      Left            =   10500
      TabIndex        =   79
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   8
      Left            =   10485
      TabIndex        =   78
      Top             =   300
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   4
      Left            =   10545
      TabIndex        =   77
      Top             =   255
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   2
      Left            =   10455
      TabIndex        =   76
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   1
      Left            =   10545
      TabIndex        =   75
      Top             =   315
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb2"
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
      Height          =   255
      Index           =   0
      Left            =   10425
      TabIndex        =   74
      Top             =   135
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   52
      Left            =   12630
      TabIndex        =   73
      Top             =   915
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   51
      Left            =   12615
      TabIndex        =   72
      Top             =   1125
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   50
      Left            =   12870
      TabIndex        =   71
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   49
      Left            =   12855
      TabIndex        =   70
      Top             =   810
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   48
      Left            =   12885
      TabIndex        =   69
      Top             =   810
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   47
      Left            =   12870
      TabIndex        =   68
      Top             =   1020
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   46
      Left            =   12915
      TabIndex        =   67
      Top             =   1260
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   45
      Left            =   12900
      TabIndex        =   66
      Top             =   1485
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   44
      Left            =   12720
      TabIndex        =   65
      Top             =   1155
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   43
      Left            =   13305
      TabIndex        =   64
      Top             =   1590
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   42
      Left            =   13260
      TabIndex        =   63
      Top             =   1350
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   41
      Left            =   12810
      TabIndex        =   62
      Top             =   1305
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   40
      Left            =   12780
      TabIndex        =   61
      Top             =   1305
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   39
      Left            =   12795
      TabIndex        =   60
      Top             =   1095
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   38
      Left            =   12750
      TabIndex        =   59
      Top             =   840
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   37
      Left            =   13020
      TabIndex        =   58
      Top             =   1245
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   36
      Left            =   12570
      TabIndex        =   57
      Top             =   690
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   35
      Left            =   12555
      TabIndex        =   56
      Top             =   900
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   34
      Left            =   12600
      TabIndex        =   55
      Top             =   1155
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   33
      Left            =   12585
      TabIndex        =   54
      Top             =   1365
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   32
      Left            =   12615
      TabIndex        =   53
      Top             =   1365
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   31
      Left            =   12600
      TabIndex        =   52
      Top             =   1575
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   30
      Left            =   12540
      TabIndex        =   51
      Top             =   990
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   29
      Left            =   12525
      TabIndex        =   50
      Top             =   1215
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   28
      Left            =   12765
      TabIndex        =   49
      Top             =   1065
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   27
      Left            =   12255
      TabIndex        =   48
      Top             =   1485
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   26
      Left            =   12210
      TabIndex        =   47
      Top             =   1245
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   25
      Left            =   12225
      TabIndex        =   46
      Top             =   1035
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   24
      Left            =   12195
      TabIndex        =   45
      Top             =   1035
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   23
      Left            =   12840
      TabIndex        =   44
      Top             =   1005
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   22
      Left            =   12795
      TabIndex        =   43
      Top             =   750
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   21
      Left            =   12720
      TabIndex        =   42
      Top             =   825
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   20
      Left            =   12975
      TabIndex        =   41
      Top             =   735
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   15
      Left            =   12765
      TabIndex        =   40
      Top             =   1185
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   14
      Left            =   12765
      TabIndex        =   39
      Top             =   885
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   13
      Left            =   12840
      TabIndex        =   38
      Top             =   1410
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   12
      Left            =   12825
      TabIndex        =   37
      Top             =   1140
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   16
      Left            =   12765
      TabIndex        =   36
      Top             =   1485
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   18
      Left            =   12750
      TabIndex        =   35
      Top             =   990
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   19
      Left            =   12750
      TabIndex        =   34
      Top             =   1275
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   17
      Left            =   13020
      TabIndex        =   33
      Top             =   1755
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   6
      Left            =   12960
      TabIndex        =   32
      Top             =   825
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   11
      Left            =   12870
      TabIndex        =   31
      Top             =   855
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   10
      Left            =   12870
      TabIndex        =   30
      Top             =   555
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   8
      Left            =   12975
      TabIndex        =   29
      Top             =   1425
      Width           =   405
   End
   Begin VB.Label Lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3705
      TabIndex        =   28
      Top             =   6615
      Width           =   1020
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2625
      TabIndex        =   27
      Top             =   6615
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   9
      Left            =   12660
      TabIndex        =   26
      Top             =   1050
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   4
      Left            =   12585
      TabIndex        =   24
      Top             =   1245
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   5
      Left            =   13050
      TabIndex        =   23
      Top             =   1365
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   7
      Left            =   12960
      TabIndex        =   22
      Top             =   1125
      Width           =   405
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "-Top Left Menu"
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
      Height          =   315
      Index           =   9
      Left            =   3540
      TabIndex        =   12
      Top             =   3900
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   0
      Left            =   12840
      TabIndex        =   21
      Top             =   780
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   3
      Left            =   12585
      TabIndex        =   20
      Top             =   960
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   2
      Left            =   12840
      TabIndex        =   19
      Top             =   1305
      Width           =   405
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Lb1"
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
      Height          =   315
      Index           =   1
      Left            =   12840
      TabIndex        =   18
      Top             =   1035
      Width           =   405
   End
   Begin VB.Label LabelCol12 
      BackColor       =   &H00800000&
      Caption         =   "Label11"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3060
      TabIndex        =   15
      Top             =   4230
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label LabelCol11 
      BackColor       =   &H005BDF79&
      Caption         =   "Label11"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3075
      TabIndex        =   10
      Top             =   3945
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "-Top Left Menu"
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
      Height          =   315
      Index           =   8
      Left            =   3540
      TabIndex        =   11
      Top             =   3585
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "-Top Left Menu"
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
      Height          =   315
      Index           =   5
      Left            =   3570
      TabIndex        =   7
      Top             =   2580
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "-Top Left Menu"
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
      Height          =   315
      Index           =   3
      Left            =   3570
      TabIndex        =   5
      Top             =   2070
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "-Top Left Menu"
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
      Height          =   315
      Index           =   6
      Left            =   3555
      TabIndex        =   8
      Top             =   2925
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "-Top Left Menu"
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
      Height          =   315
      Index           =   2
      Left            =   3570
      TabIndex        =   4
      Top             =   1800
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "-Top Left Menu"
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
      Height          =   315
      Index           =   4
      Left            =   3570
      TabIndex        =   6
      Top             =   2355
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "-Top Left Menu"
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
      Height          =   315
      Index           =   1
      Left            =   3570
      TabIndex        =   3
      Top             =   1500
      Width           =   2100
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu MNU_ART_LOGG_FOLDER 
         Caption         =   "ART LOGG FOLDER.TXT"
      End
      Begin VB.Menu MNU_ART_LOGG_FILES 
         Caption         =   "ART LOGG FILES.TXT"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Mnu_loggs 
         Caption         =   "Open Logg Folder"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save As ..."
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "Image"
      Begin VB.Menu mnuRotateMaster 
         Caption         =   "Rotate"
         Begin VB.Menu mnuRotate 
            Caption         =   "Clockwise"
            Index           =   0
         End
         Begin VB.Menu mnuRotate 
            Caption         =   "Counter Clockwise"
            Index           =   1
         End
         Begin VB.Menu mnuRotate 
            Caption         =   "180 Degrees"
            Index           =   2
         End
         Begin VB.Menu Mnu_Mic_Rotate 
            Caption         =   "Microsoft Rotate 90"
         End
      End
      Begin VB.Menu mnuMirrorMaster 
         Caption         =   "Mirror"
         Begin VB.Menu mnuMirror 
            Caption         =   "Horizontal"
            Index           =   0
         End
         Begin VB.Menu mnuMirror 
            Caption         =   "Vertical"
            Index           =   1
         End
      End
      Begin VB.Menu mnuAutosize 
         Caption         =   "Autosize Window"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Mnu_Rotate90Save 
      Caption         =   "! Rotate 90 an Save"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_Move 
      Caption         =   "! Move"
      Begin VB.Menu Mnu_MoveGood 
         Caption         =   "Move to Good"
      End
      Begin VB.Menu Mnu_Move_XXXRay 
         Caption         =   "Move to X"
      End
      Begin VB.Menu Mnu_TrobPix1 
         Caption         =   "View Trob Pix"
      End
   End
   Begin VB.Menu Mnu_MoveBad 
      Caption         =   "Move to Bad"
   End
   Begin VB.Menu Mnu_Del 
      Caption         =   "! Del"
   End
   Begin VB.Menu Mnu_TrobPix2 
      Caption         =   "View Trob Pix"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_DontSave 
      Caption         =   "! Dont Save Pointers"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_Date 
      Caption         =   "! Time Date"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_WebCam2 
      Caption         =   "! Web Cam"
      Begin VB.Menu Mnu_WebCam3 
         Caption         =   "My Web Cam"
      End
      Begin VB.Menu Mnu_WebCam4 
         Caption         =   "VBVidCap Web Cam"
      End
      Begin VB.Menu Mnu_WebCam 
         Caption         =   "Microsoft Web Cam AMCap"
      End
   End
   Begin VB.Menu Mnu_ShowPic 
      Caption         =   "! Actual Size"
   End
   Begin VB.Menu Mnu_IceView 
      Caption         =   "! Ice View"
   End
   Begin VB.Menu Mnu_Thumbs 
      Caption         =   "! Thumbs Plus"
   End
   Begin VB.Menu Mnu_Switches 
      Caption         =   "! Switches"
      Begin VB.Menu Mnu_StartSame 
         Caption         =   "Start Exact Same Place"
      End
      Begin VB.Menu Mnu_DelKey 
         Caption         =   "Del Key Enabled"
      End
      Begin VB.Menu Mnu_Kids 
         Caption         =   "Kids Around"
      End
      Begin VB.Menu Mnu_AutoStart 
         Caption         =   "Auto Start"
      End
      Begin VB.Menu Mnu_FullScreen_Start 
         Caption         =   "Start Full Screen"
      End
      Begin VB.Menu Mnu_Pause_When_Not_In_Focus 
         Caption         =   "Pause When Not In Focus"
         Checked         =   -1  'True
      End
      Begin VB.Menu MNU_RAND_A_MINUTE 
         Caption         =   "RAND A MINUTE"
      End
      Begin VB.Menu MNU_GET_INTERNET_PICS 
         Caption         =   "GET INTERNET PICTURES"
      End
      Begin VB.Menu MNU_MINIIMIZE_STANDBY 
         Caption         =   "MINIIMIZE STANDBY"
      End
      Begin VB.Menu MNU_DONT_USE_CLICK_FOR_PAUSE 
         Caption         =   "DONT USE CLICK FOR PAUSE"
      End
   End
   Begin VB.Menu Mnu_Feeds 
      Caption         =   "! Feeds"
      Begin VB.Menu Mnu_1 
         Caption         =   "# Today FeedDemon"
         Index           =   0
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "1"
         Index           =   1
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "1"
         Index           =   2
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "1"
         Index           =   3
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "1"
         Index           =   4
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "1"
         Index           =   5
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "1"
         Index           =   6
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "1"
         Index           =   7
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "1"
         Index           =   8
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "1"
         Index           =   9
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "1"
         Index           =   10
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "Web Cam"
         Index           =   11
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "Web Cam"
         Index           =   12
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "Web Cam"
         Index           =   13
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "Web Cam"
         Index           =   14
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "Web Cam"
         Index           =   15
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "Web Cam"
         Index           =   16
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "Web Cam"
         Index           =   17
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "Web Cam"
         Index           =   18
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "Web Cam"
         Index           =   19
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "Web Cam"
         Index           =   20
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "Web Cam"
         Index           =   21
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "Web Cam"
         Index           =   22
      End
      Begin VB.Menu Mnu_1 
         Caption         =   "-:¦:-•:*'''''*:•-:¦:-"
         Index           =   23
      End
   End
   Begin VB.Menu Mnu_RunVB 
      Caption         =   "! Run VB"
   End
   Begin VB.Menu MNU_FULL_SCREEN_YES 
      Caption         =   "FULL SCREEN"
   End
End
Attribute VB_Name = "fmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim KEYRELAX

'Art Master Coder ver 56 on 08-Aug-09

'DONE CLIPBOARD PICTURE PATH
'DONE LOGG ALL DIRS BEEN

Public MNU_FULL_SCREEN_YES_VAR

'---------------------------
'COPY FILES

Public RamDrive


'SET HERE
'Slider1.Value = 0


'TEMP TAKEN OUT
'Shell "D:\VB6\VB-NT\00_Best_VB_04\INet_Content_Jpgs\INet_Content_Jpgs.exe -Quite", vbNormalNoFocus

Dim COUNTTOTAL As Long
Dim i2, FR4
Dim RANMIN, WEBCAMRELOAD
Dim FORMLOADFLAG
Dim Slider1_Var_Width, Slider1_Var_Height, Slider1_Var_Left, Slider1_Var_Top

Dim SliderOverRide1, SliderOverRide2, OldML
Dim TTH, TTQ, TTQX
'Dim MNU_FULL_SCREEN_YES
Dim Toad1, Toad2, DaysBack, LoopCounter
Dim ADJUSTX As Long, ADJUSTY As Long, BadExit, OLABCAP, LBB3T, KK, DirNo, KKTime
Dim lPic As Picture, Reverse, FormLoad, PauseNow, InDisplay, QueUpIndex
Dim FormUnLoading, OldSlider
'Many Timers make hard to stop an trace a bug switch off all the fast timers here
'Const TestUp = True

Dim SingleNum
Dim DI1, DI2

Dim Msg2, DontCount, OLBB3
Dim InPutDate, TestDate, Year5, MonthStart, DayCount1, DayCountMonth, WholeYear1
Dim Month5, WeeksSinceYear, WeeksSinceInput, ResultYearDate, Month7


'To do with Frame
Public KeyCode2, TT1, TT2, OldAutoCntTimer, KeyDir
Dim XP1, YP1, ForBACKCTRL As Integer, LASTCTRL As Integer, OH, Button2, DoneVis
Public PauseCTRL As Integer, RandCTRL As Integer, OLDA2$, OldIndex
Dim XD, TimerFlashV1, TimerFlashV2, TimerFlashV3, EasyNow
Dim DateOld, WidthLeft, OldXXd
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
 
Private Declare Function Putfocus _
        Lib "user32" _
        Alias "SetFocus" _
        (ByVal hwnd As Long) As Long


Public WTrue As Double
Public HWTrue As Double
Public KWTrue As Double
Public TW1 As Double
Public TW2 As Double
Public TW3 As Double


Dim StandBy, GapWidth, Tots(), ReStarted

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
   hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal _
   dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal _
   dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject _
   As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal _
   dwMilliseconds As Long)

Private Const INFINITE = &HFFFF
Private Const SYNCHRONIZE = &H100000
Private Const WAIT_TIMEOUT = &H102


Dim MyPic As StdPicture, KillOkayFlag, Cmd$
Dim UU2(), TrobPix, Filename As String
Public m_Image                    As New cImage
'Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
Public OldA1, OldA3, PicFilesToT, PicFilesOffSet

Public PicFilesOffFolderNX
Public PicFilesOffFolderLast

Public Go3, Go2, Go1, Go4
Const Down = -1
Const Up = 1

'Option Explicit
'Option Base 0

Public CountPic


'Reserved space around picturebox
Private Const PictureBoxLeft      As Long = 0
Private Const PictureBoxTop       As Long = 0
Private Const PictureBoxRight     As Long = 0
Private Const PictureBoxBottom    As Long = 240   '240 because form has a menu

'Mouse button for grab and drag
Private Const ButtonDrag          As Integer = 1  'Left Mouse
Private PaintLeft           As Long
Private PaintTop            As Long

Private Const TwipsPerPixel       As Long = 15    'Is this ever not true?





Private Sub Dir2_Change()


t55 = 1
For R = Dir2.ListCount - 1 To Dir2.ListCount - 15 Step -1
    If Mnu_1(t55).Caption = "1" Then
    Mnu_1(t55).Caption = "RSS Feed Demon" + Mid(Dir2.List(R), InStrRev(Dir2.List(R), "\"))
    t55 = t55 + 1
    End If
Next

End Sub

Private Sub Dir3_Change()

t55 = 1
If Dir3.ListCount <= 1 Then Exit Sub
For R = Dir3.ListCount - 1 To Dir3.ListCount - 15 Step -1
    Do
    t55 = t55 + 1
    QQ = 0
    If Mnu_1.UBound < t55 Then Exit Do
    If Mnu_1(t55).Caption = "Web Cam" Then
        Mnu_1(t55).Caption = "WebCam " + Mid(Dir3.List(R), InStrRev(Dir3.List(R), "\"))
        If Dir(Dir3.List(R) + "\*.jpg") = "" Then t55 = t55 - 1
        'Debug.Print Dir(Dir3.List(R) + "\*.jpg")
        QQ = 1
    End If
    Loop Until QQ = 1
Next

End Sub

Private Sub Form_GotFocus()

On Error Resume Next
Picture1.SetFocus



End Sub

Private Sub Form_Load()
'PicShowReSize.Show

MNU_RAND_A_MINUTE.Checked = False

Timer7.enabled = False
i2 = SetScreenSaverState(False, False)
SliderOverRide1 = True
Frame1.BackColor = &H400000
Frame2.BackColor = Frame1.BackColor
Frame3.BackColor = Frame1.BackColor
Frame4.BackColor = Frame1.BackColor
RFLb.Visible = False
RFLb = "RF"

'e = Str(20)
FormLoad = True

On Error Resume Next
MkDir App.Path + "\" + GetComputerName + "\00_ART_Data"
MkDir App.Path + "\" + GetComputerName
On Error GoTo 0

If IsIDE = False Then
    On Error Resume Next
    FR5 = FreeFile
    Open RamDrive + ":TEMP\ART_PROG_TRIP_WIRE.txt" For Output As #FR5
    Close #FR5
    
    If Err.Number > 0 Then
        FR5 = FreeFile
        Open App.Path + "\ART_PROG_TRIP_WIRE.txt" For Output As #FR5
        Close #FR5
    End If
    On Error GoTo 0
    
    Timer_TRIP_WIRE.enabled = True
    'Shell "D:\VB6\VB-NT\00_Best_VB_01\Run_AS\Matt_RunAS.exe ART", vbMinimizedNoFocus
End If
If IsIDE = True Then
    
    On Error Resume Next
        Kill RamDrive + ":TEMP\ART_PROG_TRIP_WIRE.txt"
        Kill App.Path + "\ART_PROG_TRIP_WIRE.txt"
    On Error GoTo 0
    
End If
If App.PrevInstance = True Then End

Set FS = CreateObject("Scripting.FileSystemObject")


On Error Resume Next
MkDir App.Path + "\" + GetComputerName
MkDir App.Path + "\" + GetComputerName + "\00_ART_Data"
On Error GoTo 0


FR4 = FreeFile
Open App.Path + "\" + GetComputerName + "\00_ART_Data\ART PICTURE VIEW LOGG.txt" For Append Lock Write As #FR4
    
    
    


On Error Resume Next
FR5 = FreeFile
Open App.Path + "\#00 Loggs\COUNT-POSI\TOTAL Count--.txt" For Input As #FR5
Line Input #FR5, TCNT
Close #FR5
COUNTTOTAL = Val(TCNT)
If Err.Number > 0 Then msgbox "FILE NOT FOUND OR EMPTY" + vbCrLf + App.Path + "\#00 Loggs\COUNT-POSI\TOTAL Count--.txt", vbMsgBoxSetForeground

On Error GoTo 0



    
    
    



On Error Resume Next
'Kill "Q:\Temp\WebCamFlag.txt"
On Error GoTo 0

ReDim UU2(500)
ReDim Tots(100, 100)
ReDim CTCOL(fmain.Controls.Count)

For Each Control In fmain.Controls
    If InStr(LCase(Control.Name), "timer") > 0 Then
        If Control.enabled = True Then
            i = i + 1
            CTCOL(i) = Control.Name
            Control.enabled = False
        End If
    End If
Next
ReDim Preserve CTCOL(i)

OldCountPic1 = True
CountPic = True
OTimer1 = -2
KeyDir = 1

On Error Resume Next
For Each Control In Controls
    
    If Mid$(Control.Name, 1, 1) = "L" And Mid$(Control.Name, 1, 3) <> "Lbl" Then Control.Caption = ""
    'Control.Text = ""
Next
On Error GoTo 0

TW1 = 1
TW2 = 2
TW3 = 3
OldSlider = True


Load ScanPath
'ScanPath.Visible = True
'DoEvents
fmain.Top = 300
fmain.Left = 0
fmain.Height = Screen.Height - 800
fmain.Width = Screen.Width
'DoEvents





'Call mnuOpen_Click
'fMain.Visible = True
'DoEvents
'Exit Sub
        
        
    'fMain.Refresh
    'For Each Control In fMain
    'Control.Name
    'ex = 1
    'If InStr(Control.Name, "Timer") > 0 Then ex = 0
    'If InStr(Control.Name, "Mnu") > 0 Then ex = 0
    'If InStr(Control.Name, "mnu") > 0 Then ex = 0
    'If ex = 1 Then Control.Visible = False ': Control.Visible = True
    'Next
    'DoEvents
        
Lab(0) = "Awe #"
Lab(1) = "" 'Tot Files Count" 'Leave this blank error in line up code
Lab(2) = "Files" 'Files Count"
Lab(3) = "Dir" 'Dir Count"
LASTCTRL = 5
Lab(LASTCTRL - 1) = "1st Pic && Stop"
Lab(LASTCTRL) = "Last Pic && Stop"

Lab(LASTCTRL - 1).BackColor = LabelCol3.BackColor
Lab(LASTCTRL).BackColor = LabelCol3.BackColor
Lab(LASTCTRL - 1).ForeColor = 0
Lab(LASTCTRL).ForeColor = 0

RandCTRL = 7
Lab(RandCTRL - 1) = "Randomizor"
Lab(RandCTRL) = "Randomizor Forward"

Lab(RandCTRL - 1).BackColor = LabelCol8.BackColor
Lab(RandCTRL).BackColor = LabelCol8.BackColor
Lab(RandCTRL - 1).ForeColor = 0
Lab(RandCTRL).ForeColor = 0

Lab(8) = "Faster Than Light"
Lab(9) = "Sleepy Slowly 3s zzz"

Lab(8).BackColor = LabelCol5.BackColor
Lab(9).BackColor = LabelCol5.BackColor
Lab(8).ForeColor = 0
Lab(9).ForeColor = 0

PauseCTRL = 10
Lab(PauseCTRL) = "Hitt to Pause"

Lab(11) = "Next Folder"
'Lab(11) = "Rotate 90^ && Save"


'Lab(13) = "Load In IceView"
ForBACKCTRL = 12
Lab(ForBACKCTRL) = "<º)))> Reverse Mode <(((º>"
Lab(ForBACKCTRL).BackColor = LabelCol11.BackColor
Lab(ForBACKCTRL).ForeColor = 0
'    LBB3 = "Reverse"
    'LBB3 = "Go Fourth && Multiply"
    LBB3 = "."


'ReDim Preserve Lab(11)
        
        'Picture1.Top = 5000
        'Picture1.Left = 5000
        
        
        HeightLeft = 200
        sandw = 15
        'Lb1(1).Width = 2600 + 350 'If need more this is one left side boxes
        'DoEvents
        'Lb1(1).Height = HeightLeft
        'Lab(2).Left = 20
        
        
        LBX(2) = Format(Int(Now) + 0.7, "DDD DD-MMM-YY")
        
        XXW = 0
        For Each Control In fmain
            XX = 0
            If InStr(Control.Name, "Lab") > 0 Then XX = 2
            If InStr(Control.Name, "Label") > 0 Then XX = 3
            If InStr(Control.Name, "Lbxl") > 0 Then XX = 4
            
            CTRLX2 = 10
            If XX = 2 Then
                Control.FontSize = CTRLX2
                Control.AutoSize = True
                If XXW < Control.Width Then XXW = Control.Width ': Debug.Print Control.Caption, xxw
                Control.AutoSize = False
            End If
            If XX = 3 Then
                Control.AutoSize = True
                If XXW < Control.Width Then
                    XXW = Control.Width ': Debug.Print Control.Caption, xxw
                End If
                Control.AutoSize = False
            End If
            If XX = 4 Then
                Control.AutoSize = False
            End If
        Next
        
        WidthLeft = 2630
        'WidthLeft = xxw - 550
        
        'For PM got a tailing p longer height
        LBX(2) = Format(Int(Now) + 0.7, "DDD DD-MMM-YY HH:MM:SSa/p")
        
        For Each Control In fmain
            XX = 0
            If InStr(Control.Name, "Lb1") > 0 Then XX = 1
            If InStr(Control.Name, "Lb2") > 0 Then XX = 1
            If InStr(Control.Name, "LBX") > 0 Then XX = 3
            If InStr(Control.Name, "Lab") > 0 Then XX = 2
            If InStr(Control.Name, "Label") > 0 Then XX = 3
            If InStr(Control.Name, "RTB") > 0 Then XX = 4
            
            
            If XX = 1 Then
                'right hand list
                Control.FontSize = 7.8
            '    Control.Width = Lb1(1).Width
            '    Control.Height = Lb1(1).Height
            End If
            If XX = 2 Then
                'Debug.Print Control.Caption
                
                Control.FontSize = CTRLX2
                Control.AutoSize = True
                Control.AutoSize = False
                Control.Width = WidthLeft
                Control.Height = Lab(0).Height
            End If
            If XX = 3 Then
                Control.AutoSize = True
                Control.AutoSize = False
                Control.Left = 20
                Control.Width = WidthLeft
            End If
            If XX = 4 Then
                Control.Text = "."
                Control.SelStart = 0
                Control.SelLength = 1
                'RTB
                Control.Font.Size = 9
                Control.SelColor = RGB(255, 255, 255)
                Control.Width = WidthLeft - 10 ' -- -10 becoz RTB! sticks out tid mor than the other controls
                Control.Left = 20
                Control.Text = ""
            End If
        Next
        
        LBX(2) = ""
        
        'Me.BackColor = &HFFFFFF
        
        PBar1.Left = 20 - 25
        PBar2.Left = 20 - 25
        
        PBar1.Width = WidthLeft + 20
        PBar2.Width = WidthLeft + 20
        
        PBar1.Height = HeightLeft - 50
        PBar2.Height = HeightLeft - 50
        
        Slider1.Width = (WidthLeft / 100) * 85        'Spot On
        
        Slider1_Var_Width = Slider1.Width
        Slider1.Height = 300
        Slider1_Var_Height = Slider1.Height
        Slider1.Left = 20
        Slider1_Var_Left = Slider1.Left
        
        'yes the slider side lab trob this one
        LBT1.Left = (Slider1.Width + Slider1.Left) + 10
        LBT1.Width = ((WidthLeft / 100) * 14) + 10     'Spot On
        LBT1.Height = Slider1.Height
        LBT1.Visible = True
        
        
        'RTB1.SelLength
        
        
        
        wtop = 0
        For Each Control In Lab
            Control.Left = 20
            Control.Top = wtop
            'Sandw = Gap Size Thin Line Between Boxes
            If Control.Visible = True And Control.Caption <> "" Then wtop = wtop + Control.Height + sandw
        Next
        
        
        
        'Lab(2).Left = 20
        'Lab(2).Top = 0
        'lbx(2) = Trim(Str(CountPic)) + " of" + Str(ScanPath.ListView1.ListItems.Count)
        'Files Total Count all folders
        
        'Group of 3
        'Folder Top
        'LabelCol3.Top = fMain.Height - LabelCol3.Height - 750
        'LabelCol3 = A1
        RTB3.Top = fmain.Height - RTB3.Height - 750 - 10
        
        'Filename Middle
        RTB2.Top = RTB3.Top - RTB2.Height - sandw '+ 10
'        LabelCol5.Top = LabelCol3.Top - LabelCol5.Height - sandw
        
        'Scan Dir Bottom
        RTB1.Top = RTB2.Top - RTB1.Height - sandw
        'RTB3 = A1
'        Label13.Top = LabelCol5.Top - Label13.Height - sandw
        
        'Group of 2
        'date of file
        LBX(2).Top = RTB1.Top - LBX(2).Height - sandw
        
        'Size
        LBX(1).Top = LBX(2).Top - LBX(1).Height - sandw
        
        '%
        LBX(0).Top = LBX(1).Top - LBX(0).Height - sandw
        
        'PBar1.Left = Lab(2).Left
        PBar1.Top = wtop
            
            
            'wtop = wtop + Control.Height + sandw
        
        'PBar2.Left = Lab(2).Left
        PBar2.Top = PBar1.Top + PBar1.Height + sandw
        
        Slider1.Top = PBar2.Top + PBar2.Height + sandw
        Slider1_Var_Top = Slider1.Top
        LBT1.Top = Slider1.Top

        'Slider1.Width = PBar2.Width '- 20
        Lbl1.Height = 250
        Lbl2.Height = 250
        Lbl1.FontSize = 11
        Lbl2.FontSize = 11
        
        'The Left Right Green Taggs
        Lbl1.Left = 20
        Lbl1.Top = Slider1.Top + Slider1.Height + sandw
        Lbl2.Top = Slider1.Top + Slider1.Height + sandw
        Lbl1.Width = ((WidthLeft / 100) * 50)
        Lbl2.Left = Lbl1.Width + Lbl1.Left + 10
        Lbl2.Width = ((WidthLeft / 100) * 49) + 5
        
        'If Lbl2.Width + Lbl2.Left <> WidthLeft Then Stop
        'LBX(0).left+LBX(0).width

        



PBar1.Value = 1
PBarDUPE.Value = 1
PBar2.Value = 1
PBar1.Min = 0
PBarDUPE.Min = 0
PBar2.Min = 0
Slider1.Value = 3999
OldSlider = -1


wtop = Lbl1.Top + Lbl1.Height + 15
For Each Control In LBX
    'Control.Left = 20
    Control.Top = wtop
    'Sandw = Gap Size Thin Line Between Boxes
    If Control.Visible = True Then wtop = wtop + Control.Height + sandw
Next



'----------------

ReDim ReadArray(Lb1.Count)  'Description
ReDim ReadArra2(Lb1.Count)  'File Count and add to awesome
ReDim ReadArra3(Lb1.Count)  'Do No Pointers if speed changes Proggy Use Keep Dir of Scanned
ReDim ReadArra4(Lb1.Count)  'Do No Pointers if speed changes for use in program Keeps Speed setting

ReDim ReadArra7(Lb1.Count)  'DONT RAND A MIN START

ReDim ReadArra5(Lb1.Count)  'Do No Pointers if speed changes for use in program Set True On Change Speed
ReDim ReadArra6(Lb1.Count)  'Only Filter These Ones


'for ones we care about = -1 and rest = -2
For R = 0 To UBound(ReadArra4)
    ReadArra4(R) = -2
    ReadArra5(R) = False
    ReadArra6(R) = False
Next

'Got to Populate the Text's
x = 0
ReadArray(x) = "Command Line"

x = x + 1: ReadArray(x) = "Ultimate" 'AUTOPIX ULTRA AND SEXY
x = x + 1: ReadArray(x) = "(¯`'•.¸ -‹(•¿•)›- ¸.•'´¯)" 'AUTOPIX ULTRA AND SEXY
x = x + 1: ReadArray(x) = "<º)))><¸..·´`·..¸><(((º>" 'AUTOPIX MAIN ALL

'x = x + 1: ReadArray(x) = "Autopix Main" 'AUTOPIX MAIN ALL
ReadArra2(x) = True

x = x + 1: ReadArray(x) = "-:¦:- AutoPix 2004-6 -:¦:-"
x = x + 1: ReadArray(x) = "-:¦:- AutoPix 2007 Mar -:¦:-"
x = x + 1: ReadArray(x) = "-:¦:- AutoPix 2007 Apr -:¦:-"
x = x + 1: ReadArray(x) = "-:¦:- AutoPix 2007 May -:¦:-"
x = x + 1: ReadArray(x) = "-:¦:- AutoPix 2007 Oct -:¦:-"
x = x + 1: ReadArray(x) = "-:¦:- AutoPix zz -:¦:-"

ReadArra2(x) = True
For R = 1 To x
    ReadArra7(x) = True ' RAND A MIN
Next



'x = x + 1: ReadArray(x) = "-:¦:-•:*'''''*:•-:¦:-"'DOS AUTOPIX
'ReadArra2(x) = True
x = x + 1: ReadArray(x) = "Nasa Earth OB"
ReadArra2(x) = True
ReadArra4(x) = -1
'Great Images In Nasa
x = x + 1: ReadArray(x) = "Great Images In Nasa"
ReadArra2(x) = True
ReadArra4(x) = True
x = x + 1: ReadArray(x) = "Nasa Astro of Day"
ReadArra2(x) = True
ReadArra4(x) = True
x = x + 1: ReadArray(x) = "Nasa Hubble"
ReadArra2(x) = True
ReadArra4(x) = True
x = x + 1: ReadArray(x) = "Nasa Dryden"
ReadArra2(x) = True
ReadArra4(x) = True
x = x + 1: ReadArray(x) = "Nasa Apollo 1"
ReadArra2(x) = True
x = x + 1: ReadArray(x) = "Nasa Apollo 11"
ReadArra2(x) = True
x = x + 1: ReadArray(x) = "Comet Shoemaker-Levy"
ReadArra2(x) = True
ReadArra4(x) = True

'Hubbalisous


x = x + 1: ReadArray(x) = "ALL PHOTO ART"
ReadArra2(x) = True
ReadArra4(x) = True

x = x + 1: ReadArray(x) = "ALL ART"

x = x + 1: ReadArray(x) = "ALL CRACK SHOT"
ReadArra2(x) = True
ReadArra4(x) = True

x = x + 1: ReadArray(x) = "NIKON"
ReadArra2(x) = True
ReadArra4(x) = True

x = x + 1: ReadArray(x) = "SONY CYBER SHOT"
ReadArra2(x) = True
ReadArra4(x) = True


x = x + 1: ReadArray(x) = "Documents"
ReadArra2(x) = True

x = x + 1: ReadArray(x) = "Mental 00"
ReadArra2(x) = True
x = x + 1: ReadArray(x) = "Mental 01"
ReadArra2(x) = True

x = x + 1: ReadArray(x) = "BT Exchange"
ReadArra2(x) = True



x = x + 1: ReadArray(x) = "DVD ART A"
ReadArra2(x) = True  'File Count and add to awesome
ReadArra4(x) = True  'Do No Pointers if speed changes for use in program Keeps Speed setting

x = x + 1: ReadArray(x) = "DVD ART B"
ReadArra2(x) = True  'File Count and add to awesome
ReadArra4(x) = True  'Do No Pointers if speed changes for use in program Keeps Speed setting

'M:\0 00 Art\00 My Pictures\01 All Phone\#B0XxXRay
x = x + 1: ReadArray(x) = "Fantasy"
ReadArra2(x) = True
ReadArra4(x) = True
x = x + 1: ReadArray(x) = "Google Image Web Grab"
ReadArra2(x) = True

x = x + 1: ReadArray(x) = "Google Image"
ReadArra2(x) = True


x = x + 1: ReadArray(x) = "RSS Feeds"
ReadArra2(x) = True

x = x + 1: ReadArray(x) = "DOS 6.22 & CDROM - WOMEN"
ReadArra2(x) = True
x = x + 1: ReadArray(x) = "DOS 6.22 & CDROM - IMAGES"
ReadArra2(x) = True



x = x + 1: ReadArray(x) = "Picasa"
x = x + 1: ReadArray(x) = "FB"
x = x + 1: ReadArray(x) = "Psycho"
x = x + 1: ReadArray(x) = "Photo Bucket"
ReadArra2(x) = True
ReadArra2(x) = True

x = x + 1: ReadArray(x) = "My Web Photos"
ReadArra2(x) = True

x = x + 1: ReadArray(x) = "David Rees...."
ReadArra2(x) = True
ReadArra4(x) = True
x = x + 1: ReadArray(x) = "Fractals"
ReadArra2(x) = True
ReadArra7(x) = True ' RAND A MIN

x = x + 1: ReadArray(x) = "Mega Photos"
ReadArra2(x) = True
ReadArra4(x) = True
ReadArra7(x) = True


x = x + 1: ReadArray(x) = "75000 Photos"
ReadArra2(x) = True
ReadArra4(x) = True
ReadArra7(x) = True

x = x + 1: ReadArray(x) = "Cosmi"
ReadArra2(x) = True
ReadArra4(x) = True
x = x + 1: ReadArray(x) = "Britannica 2000 Deluxe"
ReadArra2(x) = True

x = x + 1: ReadArray(x) = "Screen Shots"
ReadArra2(x) = True

x = x + 1: ReadArray(x) = "App Shots"
ReadArra2(x) = True

x = x + 1: ReadArray(x) = "Hot Key Screen"
ReadArra2(x) = True
x = x + 1: ReadArray(x) = "Hot Key App"
ReadArra2(x) = True

x = x + 1: ReadArray(x) = "-- All Screen && App --"
ReadArra7(x) = True ' RAND A MIN
x = x + 1: ReadArray(x) = "All Screen-App 3 Day's"
ReadArra7(x) = True ' RAND A MIN
'ReadArra2(x) = True

'ReadArra4(X) = TRUE 'Cant have no pointers on ones start at end in rev

x = x + 1: ReadArray(x) = "--- Internet Total ---"
ReadArra7(x) = True ' RAND A MIN

'ReadArra2(X) = True
x = x + 1: ReadArray(x) = "Internet 20 Days"
ReadArra6(x) = True
x = x + 1: ReadArray(x) = "Internet 10 Days"
ReadArra6(x) = True

'x = x + 1: ReadArray(x) = "Internet 5 Days"
'ReadArra6(x) = True

x = x + 1: ReadArray(x) = "Internet 2 Days"
ReadArra6(x) = True
x = x + 1: ReadArray(x) = "Internet 1 Day"
ReadArra6(x) = True

'ReadArra2(X) = True

x = x + 1: ReadArray(x) = "--- WebCam Total ---"
'ReadArra2(X) = True
ReadArra7(x) = True ' RAND A MIN

x = x + 1: ReadArray(x) = "WebCam 30 Days"
x = x + 1: ReadArray(x) = "WebCam 10 Days"
x = x + 1: ReadArray(x) = "WebCam 5 Days"
x = x + 1: ReadArray(x) = "WebCam 2 Days"
x = x + 1: ReadArray(x) = "WebCam 1 Day"
'ReadArra2(X) = True
        
        
'For R = 0 To x
'Debug.Print Len(ReadArray(R))
'Next
        
Lb1(x).BackColor = LabelCol11.BackColor
Lb1(x).ForeColor = 0
'After POPulate we can adjust size labels
sandw = 10
wtop = 0
x = 0
For Each Control In Lb1
    'Control.Left = (fMain.Width - Control.Width) + 170 ' - 10
    Control.Top = wtop
    wtop = wtop + Control.Height + sandw
    Control.Width = 1
    Control.Height = 1
    Control.AutoSize = True
            
    Control.Caption = ReadArray(x)
    If ReadArray(x) = "" Then Control.Visible = False
    x = x + 1
Next
        

For Each Control In Lb1
    pi1 = Control.Width
    If pi1 > pi2 Then pi2 = pi1
    Control.AutoSize = False
    If Control.Visible = False Then Exit For
Next

hh = Lb1(0).Height
xcount = 0
oldcontrol = 0
For Each Control In Lb1
    Control.Height = hh - 15
    Control.Width = pi2 + 80
    Control.Top = oldcontrol
    oldcontrol = Control.Top + Control.Height '- 15 '+ 5
    Control.Left = (Screen.Width - Control.Width) - 120
    xcount = xcount + 1
    If Control.Visible = False Then Exit For
Next
DoEvents

'Set Size 2 Labs
Call Lb2Size

'Final Set Pic Boz Size
Picture1.Top = 15                                               'Hair Line Away from top
Picture1.Height = fmain.Height - 770                   'get this right prog dont chnage much
Picture1.Left = RTB1.Left + RTB1.Width + 15      'and left side
Picture1.Width = (Lb2(0).Left - Picture1.Left) - 15
'Me.BackColor = RGB(255, 255, 255)






'Recall -- Remember the sum forever even reload prog
'store the Martix array
FR5 = FreeFile
Open App.Path + "\#00 Loggs\CountLabsGridArray.txt" For Input Lock Write As #FR5
For Each Control In Lb1
    If Control.Visible = False Then Exit For
    On Error Resume Next
    Line Input #FR5, line1
    Line Input #FR5, line2
    On Error GoTo 0
    If line1 = "True" Then
    Tots(Control.index, 1) = True
    Else: Tots(Control.index, 1) = False
    End If
    Tots(Control.index, 2) = Val(line2)
Next
Close #FR5


    


If Dir$(App.Path + "\#00 Loggs\Error Pix.txt") = "" Then KillOkayFlag = True Else KillOkayFlag = False

    If Dir$(App.Path + "\#00 Loggs\Error Pix.txt") <> "" Then
        FR5 = FreeFile
        Open App.Path + "\#00 Loggs\Error Pix.txt" For Input As #FR5
        If Not EOF(FR5) Then Line Input #FR5, G31$
        Close #FR5

        
        If G31$ <> "" Then
            On Error GoTo 0
            T1 = InStrRev(G31$, "\")
            A1 = Mid$(G31$, 1, T1)
            B1 = Mid$(G31$, T1 + 1)
            On Error Resume Next
            MkDir A1 + "#Bad Pix\"
            FS.MoveFile G31$, A1 + "#Bad Pix\" + B1
            On Error GoTo 0
            FR5 = FreeFile
            Open App.Path + "\#00 Loggs\Last Error Pix.txt" For Output As #FR5
            Print #FR5, A1 + "#Bad Pix\" + B1
            Close #FR5
            G31$ = A1 + "#Bad Pix\" + B1
            Kill App.Path + "\#00 Loggs\Error Pix.txt"
            'Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
            'Response = MsgBox("Bad Pix Was Moved do You want to View", Style, , Help, Ctxt)
            'If Response = vbYes Then   ' User chose Yes.
            '    Shell "explorer /select, " + G1$, vbNormalFocus
            'End If
            KillOkayFlag = True
        End If
    End If


Call AddTotalAwesome(0)

'Wat a Job Now Re-Enable any timers that were enabled on start SWEET Safe
For Each Control In fmain.Controls
    If InStr(LCase(Control.Name), "timer") > 0 Then
        For i = 1 To UBound(CTCOL)
            If CTCOL(i) = Control.Name Then Control.enabled = True
        Next
    End If
Next

FormLoad = False

Slider1.Max = 100

FirstRun2 = True
FirstRun3 = True

ReStarted = True
LBB2.Caption = "The Hole"
LBB3.Caption = "The Force"
LBB4.Caption = Format(Now, "DDD DD-MMM-YYYY") + vbCrLf + Format(Now, "HH:MM:SSa/p")
LBB5.Caption = "---"
LBB5.FontSize = 11
LBB3.FontSize = LBB5.FontSize


'Call LBB4_Change2
Call LBB2_Change2
Call LBB4_Change2
Call LBB5_Change

Call zzLoad_Checks
MNU_FULL_SCREEN_YES_VAR = False



If Mnu_StartSame.Checked = True Then
    On Error Resume Next
    F = FreeFile
    Open App.Path + "\#00 Loggs\Agwa.txt" For Input Lock Write As #F
        Line Input #F, Toad1
        Line Input #F, Toad2
    Close #F
    On Error GoTo 0

    For Each Control In Lb1
        If Control.Caption = Toad1 Then
            AgWa = Control.index
            QueUpIndex = AgWa
        End If
    Next

    If Mnu_Kids.Checked = True Then
        If Lb1(QueUpIndex) = "(¯`'•.¸ -‹(•¿•)›- ¸.•'´¯)" Then xi = 1
        If Lb1(QueUpIndex) = "<º)))><¸..·´`·..¸><(((º>" Then xi = 1
        If Lb1(QueUpIndex) = "-:¦:-•:*'''''*:•-:¦:-" Then xi = 1
        If xi = 1 Then AgWa = 0: QueUpIndex = AgWa
    End If
End If

If Mnu_StartSame.Checked = False Then
    Toad1 = "WebCam 1 Day"
    'Toad1 = "WebCam Today"
    For Each Control In Lb1
        If Control.Caption = Toad1 Then
            AgWa = Control.index
            QueUpIndex = AgWa
            Exit For
        End If
    Next
End If

If Command$ <> "" Then
    Cmd$ = Command$
    AgWa = 0
    QueUpIndex = AgWa

    FR5 = FreeFile
    Open App.Path + "\#00 Loggs\Command.txt" For Output As #FR5
    Print #FR5, Cmd$
    Close #FR5
End If

If Dir$(App.Path + "\#00 Loggs\Command.txt") <> "" Then
    FR5 = FreeFile
    Open App.Path + "\#00 Loggs\Command.txt" For Input As #FR5
        Line Input #FR5, Cmd$
    Close #FR5
End If





On Error Resume Next
Err.Clear
'    ChDrive "M"
'    ChDir "M:\TEMP"
    
    TSS = FS.DriveExists("M")
    If TSS = False And Err.Number = 0 Then
        If Err.Number Then msgbox "M DRIVE NOT UP - END": End
    End If
On Error GoTo 0



On Error Resume Next
Dir2.Path = "M:\0 00 Art\My FeedStation Podcasts\## ARCHIVE\"
Dir3.Path = "M\0 00 Art Loggers\Web_Cam\"
On Error GoTo 0

fmain.Show
DoEvents
OldIndex = True
Call Lb1_Click(AgWa)

RamDrive = "K:\"
Err.Clear
On Error Resume Next

ChDrive "k"
If Err.Number > 0 Then
    RamDrive = Mid(App.Path, 1, 1) 'TEMP
End If

If FS.FileExists(RamDrive + ":TEMP\ART_PROG_FULLSCREEN RUN.txt") = True Then
    Set F = FS.getfile(RamDrive + ":TEMP\ART_PROG_FULLSCREEN RUN.txt")
    i = F.datelastmodified + TimeSerial(1, 0, 0)
    If i < Now Then x1 = True
End If


FR5 = FreeFile
Open RamDrive + ":TEMP\ART_PROG_FULLSCREEN RUN.txt" For Output As #FR5
Close #FR5

If Mnu_FullScreen_Start.Checked = True And x1 = False Then
    Call MNU_FULL_SCREEN_YES_Click
End If


FORMLOADFLAG = True

RANMIN = Now + TimeSerial(0, 0, 15)


'    LBB2.Visible = False
'    Frame2.Visible = False

    LBB4.Visible = False
    Frame4.Visible = False


Me.WindowState = vbMinimized



End Sub



Sub ScanFiles()
        
    If Me.WindowState > 0 Then Exit Sub
    If ScanInProgress = True Then Exit Sub
    
    'FOR SCAN PATH GO BACK DAYS
    DONTWARNERROR = False
    
    If ScanPath.ListView1.ListItems.Count = 0 Then
        
        'If BreakScanX = True Then Stop
        BreakScanX = False
        FirstDirScanned = ""
        If QueUpIndex = 0 Then QueUpIndex = AgWa
        AgWa = QueUpIndex
        
        'We Dont Want People Saying Whats that flash at Start
        DateOld = 0
        
        'Timer Standby Function Stops after Hour
        StandBy = Now + TimeSerial(1, 0, 0)
        
        'Timer3.Enabled = False 'this prints THE ONE in PERCENT
        'Timer3.Interval = 500
        'Timer3.Enabled = True
        
        OldSlider = -1
        
        CountPic = -1
        OldCountPic1 = -1
        BreakScanX = False
        ScanInProgress = True
        Slider1.Value = 2
        Call ForwardMode
            
        If AgWa = 0 Then
            For Each Control In Lb1
                'If Control = "WebCam Today" Then
                If Control = "WebCam 1 Day" Then
                    AgWa = Control.index + 1
                    QueUpIndex = AgWa
                    Exit For
                End If
            Next
        End If
        
        ScanFinished = False
        
        For Each Control In Lb1
                Control.BackColor = LabelCol12.BackColor
                Control.ForeColor = &HFFFFFF
        Next
        Call fmain.RunIn

        Lb1(AgWa - 1).BackColor = LabelCol11.BackColor
        
        ScanPath.cboDate.ListIndex = 0
        ScanPath.DTPicker1(0) = vbNullString
        ScanPath.Chk_SortOnDate.Value = vbUnchecked
        ScanPath.ListView1.ListItems.Clear
        ScanPath.cboMask = "*.jpg"
        
        XX = False
        
        
        
        
        '--------########## START
        
        
        
        On Error Resume Next
        FR = FreeFile
        
        'MkDir App.Path + "\" + GetComputerName
        'MkDir App.Path + "\" + GetComputerName + "\00_ART_Data"
        
        

        
        
        If Lb1(AgWa - 1) = "Command Line" And Cmd$ <> "" Then
                         
            'Cmd$ = ""
            XX = True
            If Mid$(Cmd$, 1, 1) = """" Then
                ScanPath.txtPath.Text = Mid$(Mid(Cmd$, 2), 1, Len(Cmd$) - 2)
            Else
                ScanPath.txtPath.Text = Cmd$
            End If
            
            Call ScanPath.cmdScan_Click
        End If
        
        
        
        If Lb1(AgWa - 1) = "Ultimate" Then
            XX = True
            
            ScanPath.txtPath.Text = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00\0-ButterFly"
            Call ScanPath.cmdScan_Click
            
            'ScanPath.txtPath.Text = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0004-Mags"
            'Call ScanPath.cmdScan_Click
            
            ScanPath.txtPath.Text = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0002-Ultra Sexi"
            Call ScanPath.cmdScan_Click
        
            'ScanPath.txtPath.Text = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0004-Vintage"
            'Call ScanPath.cmdScan_Click
        End If
        
        
        If Lb1(AgWa - 1) = "(¯`'•.¸ -‹(•¿•)›- ¸.•'´¯)" Then
            XX = True
            'ScanPath.txtPath.Text = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0000-Ultra Sexi"
            'Call ScanPath.cmdScan_Click
            'ScanPath.txtPath.Text = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0000-Sexyest"
            'Call ScanPath.cmdScan_Click
            'M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00
            ScanPath.txtPath.Text = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00"
            Call ScanPath.cmdScan_Click
        End If

        If Lb1(AgWa - 1) = "<º)))><¸..·´`·..¸><(((º>" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\01 AutoPix"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "-:¦:-•:*'''''*:•-:¦:-" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00\DOS Days"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "-:¦:- AutoPix 2004-6 -:¦:-" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 2004 to 6"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "-:¦:- AutoPix 2007 Mar -:¦:-" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 2007 03"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "-:¦:- AutoPix 2007 Apr -:¦:-" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 2007 04"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "-:¦:- AutoPix 2007 May -:¦:-" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 2007 05"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "-:¦:- AutoPix 2007 Oct -:¦:-" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 2007 10"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "-:¦:- AutoPix zz -:¦:-" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix zz 000-000"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "ALL ART" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art 01"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "Nasa Earth OB" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\07 Internet\Space Shuttle Shots"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\07 Internet\Space Shots"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\07 Internet\NASA's Earth Observatory"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\07 Internet\NaSa Images"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "ALL PHOTO ART" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\0 MY DEVICES - ALL"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\01 All Phone"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\01 All Psycho"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\My Computer CODE"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\Picture YIN JACK"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\00 Get Around Links\Photo Bucket"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "ALL CRACK SHOT" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\0 MY DEVICES - ALL\0 MY DEVICES - ART"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "NIKON" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\0 MY DEVICES - ALL\0 MY DEVICES - ART\CAMERA\NIKON TOTAL"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "SONY CYBER SHOT" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\0 MY DEVICES - ALL\0 MY DEVICES - ART\CAMERA\SONY TOTAL"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "DVD ART A" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\# PHOTO #\A"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "DVD ART B" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\# PHOTO #\B"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "DOS 6.22 & CDROM - WOMEN" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\01 All Psycho\z1Female Pictures 01"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\01 All Psycho\z1JPG Dos6.22 Days 01"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\01 All Psycho\z1JPG Dos6.22 Days 02 xx"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "DOS 6.22 & CDROM - IMAGES" Then
            XX = True

         
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\01 All Psycho\My Computer\Screen Captures"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\My Computer CODE"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\01 All Psycho\Bill Hick's"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\01 All Psycho\08 Smalls"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\07 Internet\Space Shots"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\01 All Phone\Good"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "Mental 00" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\01 All Phone\00 Mental"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "Mental 01" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\01 All Phone\01 Mental"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "Mental 2009" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\01 All Phone\00 Mental\01 006 Jan& 2009"
            Call ScanPath.cmdScan_Click
            Call ReverseMode
        End If
        
        If Lb1(AgWa - 1) = "Psycho" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\01 All Psycho"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "RSS Feeds" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\My FeedStation Podcasts"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "Fractals" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\07 Internet\Fractals\Fractals 2003"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\07 Internet\Fractals\Fractals 2004"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\07 Internet\Fractals\Fractals 2005"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\07 Internet\Fractals\Fractals 2006"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\07 Internet\Fractals\Fractals 2007"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\07 Internet\Fractals\Fractals 2008"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\07 Internet\Fractals\Fractals 2009"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "Fractals New" Then
            XX = True
            ScanPath.Chk_SortOnDate.Value = vbChecked
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\07 Internet\Fractals\Fractals 2009"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\07 Internet\Fractals\New Fractals"
            Call ScanPath.cmdScan_Click
            Call ReverseMode
        End If

        If Lb1(AgWa - 1) = "David Rees...." Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\07 Internet\mnftiu.cc David Rees"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 50
            FileInfoTimer.enabled = False
        End If
        
        'Fantasy
        If Lb1(AgWa - 1) = "Fantasy" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\# PHOTO #\B\Fantasy Art 1800+ Images"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "ART A" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\# PHOTO #\A"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "ART B" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\# PHOTO #\B"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "Mega Photos" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\z3 Mega Photos"
            Call ScanPath.cmdScan_Click
        End If
        
        If Lb1(AgWa - 1) = "75000 Photos" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\z3 75000 Photos"
            Call ScanPath.cmdScan_Click
        End If
            
            
        
        If InStr(Lb1(AgWa - 1), "Internet") = 1 Then
        ScanPath.Chk_SortOnDate.Value = vbChecked
        End If
            
        If Lb1(AgWa - 1) = "--- Internet Total ---" Then
            XX = True
            ScanPath.txtPath.Text = "M\0 00 Art Loggers\Temp Inet Files JPGs"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 0
            Call ReverseMode
        End If
        
        
        ih = Lb1(AgWa - 1)
        DaysBack = 0
        If ih = "Internet 20 Days" Then DaysBack = 20
        If ih = "Internet 10 Days" Then DaysBack = 10
        If ih = "Internet 5 Days" Then DaysBack = 5
        If ih = "Internet 3 Days" Then DaysBack = 3
        If ih = "Internet 2 Days" Then DaysBack = 2
        If ih = "Internet 1 Day" Then DaysBack = 1
            
        If DaysBack > 0 Then
            XX = True
            'ScanPath.DTPicker1(0) = Now - (DaysBack - 2)
                
            Dir1.Path = "M\0 00 Art Loggers\Temp Inet Files JPGs\"
            R5 = Dir1.ListCount
            OLDFC = 0
            TTH = 0
            Do
                R5 = R5 - 1
                Td = Dir1.List(R5)
                
                ScanPath.txtPath.Text = Td
                Call ScanPath.cmdScan_Click
                
                If ScanPath.ListView1.ListItems.Count <> OLDFC Then
                    TTH = TTH + 1
                End If
                OLDFC = ScanPath.ListView1.ListItems.Count
            
            
            
            Loop Until TTH >= DaysBack Or R5 = 0
            Slider1.Value = 0
            
            Call ReverseMode
            
        End If
        
        
        'not used
        If Lb1(AgWa - 1) = "Internet Today" Then
            XX = True
            ScanPath.txtPath.Text = "M\0 00 Art Loggers\Temp Inet Files JPGs\INet " + Format$(Now, "YYYY-MM-DD")
            Call ScanPath.cmdScan_Click
            Slider1.Value = 3
            
            If ScanPath.ListView1.ListItems.Count = 0 Then
                Lb2(AgWa - 1) = "0"
                'AgWa = AgWa - 1
                QueUpIndex = AgWa - 1
                
                'why
                'Scanpath.txtPath.Text = "M\0 00 Art Loggers\Temp Inet Files JPGs\INet " + Format$(Now - 1, "YYYY-MM-DD")
                
                'This Sorts if no files today goto one above Usual its goto last one
                'little trick here
                'OldAgWa2 = AgWa
            End If
            Call ReverseMode
        End If

        
        
        If Lb1(AgWa - 1) = "--- WebCam Total ---" Then
            XX = True
            ScanPath.txtPath.Text = "M\0 00 Art Loggers\Web_Cam"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 0
        End If
        
        
        
        WEBCAMRELOAD = False
        ih = Lb1(AgWa - 1)
        DaysBack = 0
        If ih = "WebCam 30 Days" Then DaysBack = 30
        If ih = "WebCam 10 Days" Then DaysBack = 10
        If ih = "WebCam 5 Days" Then DaysBack = 5
        If ih = "WebCam 2 Days" Then DaysBack = 2
        If ih = "WebCam 1 Day" Then DaysBack = 1: WEBCAMRELOAD = True
            
        If DaysBack > 0 Then
            XX = True
            
            
            'ScanPath.DTPicker1(0) = Now - (DaysBack - 1)
            
            Dir1.Path = "M\0 00 Art Loggers\Web_Cam\"
            R5 = Dir1.ListCount
            OLDFC = 0
            TTH = 0
            Do
                R5 = R5 - 1
                Td = Dir1.List(R5)
                
                ScanPath.txtPath.Text = Td
                Call ScanPath.cmdScan_Click
                
                If ScanPath.ListView1.ListItems.Count <> OLDFC Then
                    TTH = TTH + 1
                End If
                OLDFC = ScanPath.ListView1.ListItems.Count
            
            
            
            Loop Until TTH >= DaysBack Or R5 = 0
            Slider1.Value = 0
            
            
            If ScanPath.ListView1.ListItems.Count = 0 And BreakScanX = False Then
                Lb2(AgWa - 1) = "0"
                For Each Control In Lb1
                    Control.BackColor = LabelCol12.BackColor
                    If Control.Caption = "Mental 00" Then
                        AgWa = Control.index + 1
                        Control.BackColor = LabelCol11.BackColor
                    End If
                Next

                'This Sorts if no files today goto one above Usual its goto last one
                'little trick here ??? but we dont want that we want as normal coz we know what change to
                'OldAgWa2 = AgWa
                QueUpIndex = AgWa - 1
            End If
        End If

        
        
        'not used
        If Lb1(AgWa - 1) = "WebCam Today" Then
            XX = True
            DONTWARNERROR = True
            ScanPath.txtPath.Text = "M\0 00 Art Loggers\Web_Cam\WebCapture_" + Format$(Now, "YYYY-MM-DD-DDD")
            Call ScanPath.cmdScan_Click
            If ScanPath.ListView1.ListItems.Count = 0 Then
                For i = 1 To 20
                    ScanPath.txtPath.Text = "M\0 00 Art Loggers\Web_Cam\WebCapture_" + Format$(Now - i, "YYYY-MM-DD-DDD")
                    If ScanPath.ListView1.ListItems.Count > 0 Then Exit For
                
                Next
            End If
            
            Slider1.Value = 0
            If ScanPath.ListView1.ListItems.Count = 0 Then
                Lb2(AgWa - 1) = "0"
                For Each Control In Lb1
                    Control.BackColor = LabelCol12.BackColor
                Next
                Lb1(AgWa - 2).BackColor = LabelCol11.BackColor
                
                AgWa = AgWa - 1
                
                ScanPath.txtPath.Text = "M\0 00 Art Loggers\Temp Inet Files JPGs\INet " + Format$(Now - 1, "YYYY-MM-DD")
                
                'This Sorts if no files today goto one above Usual its goto last one
                'little trick here
                OldAgWa2 = AgWa
                QueUpIndex = AgWa
            End If
        End If

        
        
        If Lb1(AgWa - 1) = "BT Exchange" Then
            XX = True
            'ScanPath.txtPath.Text = "X:\00 Blue Tooth Exchange Folder"
            ScanPath.txtPath.Text = "D:\# MY DOCS\# 01 My Documents\00 Blue Tooth Exchange Folder"
            ScanPath.Chk_SortOnDate.Value = vbChecked
            Call ScanPath.cmdScan_Click
            Call ReverseMode
        End If
        
        If Lb1(AgWa - 1) = "Documents" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\0 MY DEVICES - ALL\0 MY DEVICES - ART\Z Docus Proofs Texts"
            ScanPath.Chk_SortOnDate.Value = vbChecked
            Call ScanPath.cmdScan_Click
        End If
        

'M:\0 00 Art\00 HTTrack
        If Lb1(AgWa - 1) = "Great Images In Nasa" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 HTTrack\00 NASA\Nasa Grin\Grin Nasa\grin.hq.nasa.gov\images\MEDIUM"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 40
        End If

        If Lb1(AgWa - 1) = "Nasa Astro of Day" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 HTTrack\00 NASA\Nasa Astro Pic of Day Archive\Astro Pic of Day Archive\apod.gsfc.nasa.gov\apod\image"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 40
        End If
        
        If Lb1(AgWa - 1) = "Nasa Hubble" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 HTTrack\00 NASA\Nasa Hubble\Nasa Hubble\imgsrc.hubblesite.org\hu\db\images"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 40
        End If
        
        If Lb1(AgWa - 1) = "Nasa Dryden" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 HTTrack\00 NASA\Nasa Dryden Aircraft\Dryden NASA\www.dfrc.nasa.gov\gallery\Photo"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 40
            
        End If
        
        If Lb1(AgWa - 1) = "Nasa Apollo 1" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 HTTrack\00 NASA\Nasa Apollo 1\Nasa Apollo 1\www.hq.nasa.gov\office\pao\History\Apollo204"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 40
            
        End If
        
        If Lb1(AgWa - 1) = "Nasa Apollo 11" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 HTTrack\00 NASA\Nasa Apollo 11\Nasa apollo 11"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 40
            
        End If
        
        If Lb1(AgWa - 1) = "Comet Shoemaker-Levy" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 HTTrack\00 NASA\Nasa ShoeMaker Levy\Nasa shoemaker-levy\www2.jpl.nasa.gov\sl9\gif"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 40
            
        End If
        
        '
        
        If Lb1(AgWa - 1) = "Picasa" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\01 All Phone\02 Picasa"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 0
            
        End If

        If Lb1(AgWa - 1) = "FB" Then
            XX = True
            hhh = AgWa
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\01 All Phone\02 FaceBook"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 0
        End If

        If Lb1(AgWa - 1) = "Britannica 2000 Deluxe" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\z5BCD-ART"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 0
            
        End If

        If Lb1(AgWa - 1) = "Cosmi" Then
            XX = True
            ScanPath.txtPath.Text = "M:\0 00 Art\00 My Pictures\z4Cosmi\5057 Photos"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 0
            
        End If

        If Lb1(AgWa - 1) = "Screen Shots" Then
            XX = True
            'ScanPath.txtPath.Text = "L:\0 00 Art Loggers\Screen-Shots"
            'Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M\0 00 Art Loggers\Screen-Shots"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 3
            Call ReverseMode
        End If

        If Lb1(AgWa - 1) = "App Shots" Then
            XX = True
            'ScanPath.txtPath.Text = "L:\0 00 Art Loggers\Screen-App-Shots"
            'Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M\0 00 Art Loggers\Screen-App-Shots"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 3
            Call ReverseMode
            
        End If

        If Lb1(AgWa - 1) = "Hot Key Screen" Then
            XX = True
            'ScanPath.txtPath.Text = "L:\0 00 Art Loggers\Hot-Key-Screen-Shots"
            'Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M\0 00 Art Loggers\Hot-Key-Screen-Shots"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 3
            Call ReverseMode
        End If

        If Lb1(AgWa - 1) = "Hot Key App" Then
            XX = True
            'ScanPath.txtPath.Text = "L:\0 00 Art Loggers\Hot-Key-App-Shots"
            'Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M\0 00 Art Loggers\Hot-Key-App-Shots"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 3
            Call ReverseMode
        End If
        
        If Lb1(AgWa - 1) = "-- All Screen && App --" Then
            XX = True
            ScanPath.Chk_SortOnDate.Value = vbChecked
            
            'ScanPath.txtPath.Text = "M\0 00 Art Loggers\Hot-Key-App-Shots"
            'Call ScanPath.cmdScan_Click
            'ScanPath.txtPath.Text = "M\0 00 Art Loggers\Hot-Key-Screen-Shots"
            'Call ScanPath.cmdScan_Click
            
            'ICY BOX
            'ScanPath.txtPath.Text = "L:\0 00 Art Loggers\Screen-App-Shots"
            'Call ScanPath.cmdScan_Click
            'ScanPath.txtPath.Text = "L:\0 00 Art Loggers\Screen-Shots"
            'Call ScanPath.cmdScan_Click
            
            ScanPath.txtPath.Text = "M\0 00 Art Loggers\Hot-Key-App-Shots"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M\0 00 Art Loggers\Hot-Key-Screen-Shots"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M\0 00 Art Loggers\Screen-App-Shots"
            Call ScanPath.cmdScan_Click
            ScanPath.txtPath.Text = "M\0 00 Art Loggers\Screen-Shots"
            Call ScanPath.cmdScan_Click
            Slider1.Value = 3
            Call ReverseMode
            
        End If
        
        If Lb1(AgWa - 1) = "All Screen-App 3 Day's" Then
            XX = True
            MoreThanOne = True
            ScanPath.Chk_SortOnDate.Value = vbChecked
            
            DaysBack = 3
            ScanPath.DTPicker1(0) = Now - DaysBack
            For R = DaysBack To 0 Step -1
                Td = "HotKeyAppCapture_" + Format$(Now - R, "YYYY-MM-DD-DDD")
                ScanPath.txtPath.Text = "M\0 00 Art Loggers\Hot-Key-App-Shots\" + Td
                If FS.FolderExists(ScanPath.txtPath.Text) Then Call ScanPath.cmdScan_Click
                
                Td = "HotKeyScreenCapture_" + Format$(Now - R, "YYYY-MM-DD-DDD")
                ScanPath.txtPath.Text = "M\0 00 Art Loggers\Hot-Key-Screen-Shots\" + Td
                If FS.FolderExists(ScanPath.txtPath.Text) Then Call ScanPath.cmdScan_Click
                
                Td = "ScreenAppCapture_" + Format$(Now - R, "YYYY-MM-DD-DDD")
                ScanPath.txtPath.Text = "M\0 00 Art Loggers\Screen-App-Shots\" + Td
                If FS.FolderExists(ScanPath.txtPath.Text) Then Call ScanPath.cmdScan_Click
                
                Td = "ScreenCapture_" + Format$(Now - R, "YYYY-MM-DD-DDD")
                ScanPath.txtPath.Text = "M\0 00 Art Loggers\Screen-Shots\" + Td
                If FS.FolderExists(ScanPath.txtPath.Text) Then Call ScanPath.cmdScan_Click
            Next
            
            'Find Mental If None Exist
            If ScanPath.ListView1.ListItems.Count = 0 And BreakScanX = False Then
                Lb2(AgWa - 1) = "0"
                For Each Control In Lb1
                    Control.BackColor = LabelCol12.BackColor
                    If Control.Caption = "Mental 00" Then
                        AgWa = Control.index + 1
                        Control.BackColor = LabelCol11.BackColor
                    End If
                Next
            End If
            Slider1.Value = 3
            Call ReverseMode
            FileInfoTimer.enabled = False
        End If

        If Lb1(AgWa - 1) = "Photo Bucket" Then
            XX = True
            ScanPath.Chk_SortOnDate.Value = vbChecked
            ScanPath.txtPath.Text = "M:\00 Get Around Links\Photo Bucket"
            Call ScanPath.cmdScan_Click
            Call ReverseMode
            Slider1.Value = 2
        End If
        
        If Lb1(AgWa - 1) = "Google Image" Then
            XX = True
            'ScanPath.Chk_SortOnDate.Value = vbChecked
            ScanPath.txtPath.Text = "M:\0 00 Art\Google Downloader"
            Call ScanPath.cmdScan_Click
            'Call ReverseMode
            Slider1.Value = 0
        End If
        
        If Lb1(AgWa - 1) = "Google Image Web Grab" Then
            XX = True
            'ScanPath.Chk_SortOnDate.Value = vbChecked
            ScanPath.txtPath.Text = "M:\0 00 Art\Google Downloader\WebImageGrab_pro_pc"
            Call ScanPath.cmdScan_Click
            'Call ReverseMode
            Slider1.Value = 0
        End If
        
        If Lb1(AgWa - 1) = "My Web Photos" Then
            XX = True
            'Scanpath.Chk_SortOnDate.Value = vbChecked
            ScanPath.txtPath.Text = "D:\MY WEBS\MatthewLan.Com Web\Photos\galleries"
            Call ScanPath.cmdScan_Click
            'Call ReverseMode
            Slider1.Value = 2
        End If
        
        
        
        
        
        
        '----##############   END SCANS
        
        
        
        
        
        FR = FreeFile
        Open App.Path + "\" + GetComputerName + "\00_ART_Data\ART FOLDER VIEW LOGG.txt" For Append As #FR
            TT5 = Format(Now, "DD-MmM-YYYY - HH:MM:SSa/p")
            TNAME = Space$(38)
            LSet TNAME = Lb1(AgWa - 1)
            Print #FR, TT5 + " -FOLDER- " + TNAME; " -PATH- "; ScanPath.txtPath.Text
        Close #FR
        On Error GoTo 0
        
        
        
        'SET HERE
        Slider1.Value = 0
        If IsIDE Then Slider1.Value = 3
        
        'If AgWa <= 10 Or Val(Lb2(AgWa - 1)) > 10000 Then
        If ReadArra7(x) = True Then
            MNU_RAND_A_MINUTE.Checked = True
        Else
            MNU_RAND_A_MINUTE.Checked = False
        End If
        
        If XX = False Then
            msgbox "Prog HardCoding Error -- This -->" + vbCrLf + Lb1(AgWa - 1) + vbCrLf + vbCrLf + "Dont Exist": Stop
        End If
        
        If ScanPath.ListView1.ListItems.Count > 0 Then
            Timer1.enabled = True
            ScanInProgress = False
        End If
        
        End If
End Sub



Sub Lb2Size()

fmain.AutoRedraw = False
Dim YYX
On Error Resume Next
FR5 = FreeFile
Open App.Path + "\#00 Loggs\CountLabs.txt" For Input Lock Write As #FR5
YYX = Input(LOF(FR5), FR5)
Close #FR5
On Error GoTo 0

For Each Control In Lb2
    If Lb1(Control.index).Visible = False Then Exit For
    xxs = Lb1(Control.index).Caption
    'xxd = Control.Caption
    Th = InStr(YYX, "-- " + xxs)
    If Th > 0 Then
        th2 = InStrRev(YYX, "++ ", Th)
        stg = Mid$(YYX, th2 + 3, (Th - th2) - 4)
        Control.Caption = stg
    End If
Next

For Each Control In Lb2
    If Lb1(Control.index).Caption = "" Then Exit For
'    Control.Width = 1
    Control.AutoSize = True
Next
        

For Each Control In Lb2
    If Lb1(Control.index).Visible = False Then Exit For
    pi1 = Control.Width
    If pi1 > pi2 Then pi2 = pi1
    Control.AutoSize = False
    If Control.Caption = "" Then Exit For
Next


For Each Control In Lb2
    If Lb1(Control.index).Visible = False Then Exit For
    '16470
    'Lb1(Control.Index).Left-16350
    'Control.Width = 720
    'Control.Width = Lb1(Control.Index).Left - 16480
    
    GapWidth = 10
    Control.Width = pi2
    Control.Left = Lb1(Control.index).Left - Control.Width - GapWidth
    Control.Height = Lb1(Control.index).Height
    Control.Top = Lb1(Control.index).Top
    Control.Visible = True
    If Control.Caption = "" Then Control.Caption = "."

Next


XXD = 0
For Each Control In Lb2
    If Lb1(Control.index).Visible = False Then Exit For
    xxs = Len(Control.Caption)
    If xxs > XXD Then XXD = xxs
Next

OldXXd = XXD

fmain.AutoRedraw = True

End Sub


Private Sub Frame1_DragDrop(Source As Control, x As Single, y As Single)
'Frame1
End Sub


Private Sub Lb2_Change(index As Integer)

If FormLoad = True Then Exit Sub

If TotallyFiles > 10000 Then
    FR5 = FreeFile
    Open App.Path + "\#00 Loggs\CountLabs.txt" For Output Lock Write As #FR5
    For Each Control In Lb2
        If Lb1(Control.index).Caption = "" Then Exit For
        Print #FR5, "++ " + Control.Caption + " -- " + Lb1(Control.index).Caption
    Next
    Close #FR5
End If

ReadArra3(AgWa - 1) = FirstDirScanned

Call AddTotalAwesome(index)


'If the Number Lenght Digits dont Change in any LB2 then we dont need to run this
XXD = 0
For Each Control In Lb2
    If Lb1(Control.index).Visible = False Then Exit For
    xxs = Len(Control.Caption)
    If xxs > XXD Then XXD = xxs
Next

If OldXXd = XXD Then Exit Sub

OldXXd = XXD

Lb2(index).AutoSize = True

pi1 = 0
For Each Control In Lb2
    If Lb1(Control.index).Caption = "" Then Exit For
    pi1 = Control.Width
    If pi1 > pi2 Then pi2 = pi1
    Control.AutoSize = False
    If Control.Caption = "" Then Exit For


Next

For Each Control In Lb2
    If Lb1(Control.index).Caption = "" Then Exit For
    Control.Width = pi2
Next

For Each Control In Lb2
    If Lb1(Control.index).Caption = "" Then Exit For
    Control.Left = Lb1(Control.index).Left - Control.Width - GapWidth
Next

End Sub

Sub AddTotalAwesome(iNdX)
TotallyFiles = 0
For Each Control In Lb2
    
    If Lb1(Control.index).Visible = False Then Exit For
    On Error GoTo 0
    
    If ReadArra2(Control.index) = True And Val((Control.Caption)) > 0 Then
        TotallyFiles = TotallyFiles + Val(Control.Caption)
    End If
    If InStr(Lb1(Control.index).Caption, "Total") > 0 And Val((Control.Caption)) > 0 Then
        
        'Linked to remarks below
        'If total has been updated shown be indx is the vaule of control array changed then cool
        'reset flag pickup on this later few lines code more
            
        If Tots(Control.index, 1) = True And Control.index = iNdX Then
            Tots(Control.index, 1) = False
        End If
    
        For R = 1 To 7
            If InStr(Lb1(Control.index + R).Caption, "Today") Then
                
                'if the flag has been reset zero
                'becoz and new number has come for total
                'then load the new number from from today for offset in further minus-ing
                'an remian to keep number old for offset toal
                'new increment of number on today will reflect an show in additions on total
                'be cool
                
                If Tots(Control.index, 1) = False Then
                    If Val((Lb2(Control.index + R).Caption)) > 0 Then
                        Tots(Control.index, 2) = Val((Lb2(Control.index + R).Caption))
                        Tots(Control.index, 1) = True
                    End If
                End If
                
                        
                'Minus the sum of the today from total
                TotallyFiles = TotallyFiles - Tots(Control.index, 2)
            
                'why when iclick today an increased not show diff on counter look
                'coz it always need to have a ref of what number was when it last update
                'which it has got
                'but when add from a today total you need to substract what the old total stored
                'with update TOTAL that was from when that was updated
                
                'Remember the sum forever even reload prog
                'store the grid array

                
                FR5 = FreeFile
                Open App.Path + "\#00 Loggs\CountLabsGridArray.txt" For Output Lock Write As #FR5
                For Each Control2 In Lb2
                    If Lb1(Control2.index).Visible = False Then Exit For
                    Print #FR5, Tots(Control2.index, 1)
                    Print #FR5, Tots(Control2.index, 2)
                Next
                Close #FR5
            End If
        Next
    End If
Next

Lab(0) = "Aw #" + Format(TotallyFiles, "###,###,##0") + ""
Lab(0) = Lab(0).Caption + " #" + Format(COUNTTOTAL, "###,###,##0") + ""

If FS.FolderExists("D:\TEMP") = False Then MkDir "D:\TEMP"
If FS.FolderExists("D:\TEMP\VBArea") = False Then MkDir "D:\TEMP\VBArea"

FR5 = FreeFile
Open "D:\TEMP\VBArea\Awesome.txt" For Output Lock Write As #FR5
Print #FR5, TotallyFiles
Close #FR5




End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Dim NewWidth As Long
    Dim NewHeight As Long

    If Me.WindowState = 1 Then
        FileInfoTimer.enabled = False
        Call PauseOnOffStat
    End If

    If Me.WindowState = 0 Then
        FileInfoTimer.enabled = True
        Call PauseOnOffStat
    End If

    fmain.Refresh
    For Each Control In fmain
        ex = 1
        If InStr(Control.Name, "Timer") > 0 Then ex = 0
        If InStr(Control.Name, "Mnu") > 0 Then ex = 0
        If InStr(Control.Name, "mnu") > 0 Then ex = 0
        If InStr(Control.Name, "MNU") > 0 Then ex = 0
        If ex = 1 Then Control.Refresh ': Control.Visible = True
    Next
    
    DoEvents
    PBar3.Left = -10
    Frame1.Width = 8100
    PBar3.Width = Frame1.Width + 40
    LBB1.Width = Frame1.Width + 55
    LBB1.AutoSize = False
    LBB1.Left = 0
    
    Frame1.Top = fmain.Top - 380 '+ 680
    Frame1.Left = fmain.Width / 2 - Frame1.Width / 2
    Frame1.Height = PBar3.Top + PBar3.Height - 15

    Versionv = "Version - " + Trim(str(App.Major)) + "." + Trim(str(App.Minor)) + "." + Trim(str(App.Revision))
    Versionvs = "Ver. " + Trim(str(App.Major)) + "." + Trim(str(App.Minor)) + "." + Trim(str(App.Revision))

    Set F = FS.getfile(App.Path + "\" + App.EXEName + ".exe")

    'Msg2 = Msg2 + "-- Last Compiled " + Format$(F.datecreated, "MMM DD-MMM-YYYY HH:MM:SSa/p") + vbCrLf
    Msg2 = Msg2 + "Product Start Date -- Mar 02-Mar-2009 07:40:33a" + vbCrLf

    LastCompiled = "Last Compiled -- " + Format$(F.datelastmodified, "MMM DD-MMM-YYYY HH:MM:SSa/p")
    LastCompiledShort = Format$(F.datelastmodified, "DD-MMM-YY HH:MMa/p")

    'Use this
    InPutDate = DateValue("02-03-2009")
    TestDate = DateValue("06-01-2009")
    TestDate = Now
    Call FindTimeInfo

    ProgAge = "Age = " + ResultYearDate + " Yrs" + str(Month7) + "M" + str(DayCountMonth) + "D" + str(WeeksSinceInput) + " Wks"
'    fMain.Caption = "JPEG Encoder -- Shockingly Good Photo ART -- [-:¦:-•:*''' The One '''*:•-:¦:-] -- [" + Versionvs + "] -- [" + LastCompiledShort + "] -- [" + ProgAge + "]"
    'Me.Caption = "JPEG Encoder -- Shockingly Good Photo ART -- [-:¦:-•:*''' The One '''*:•-:¦:-]" ' -- [" + Versionvs + "] -- [" + LastCompiledShort + "] -- [" + ProgAge + "]"
    Me.Caption = "JPEG Encoder -- Shockingly Good Photo ART -- [-:¦:-•:*''' The One '''*:•-:¦:-]"

    
    On Error Resume Next
    Picture1.SetFocus

    'Shockingly Good Photo's
    'Exceeddingly good cakes

End Sub
Private Sub Form_Unload(Cancel As Integer)

Call zzCheckTimer_Timer

i2 = SetScreenSaverState(True, True)


Close FR4

FormUnLoading = True
    
Set m_Image = Nothing
Set MyPic = Nothing

Dim i As Integer
i = FileInfoTimer.enabled
'On Error Resume Next
F = FreeFile
Open App.Path + "\#00 Loggs\Agwa.txt" For Output Lock Write As #F
'    Print #F, "0"
    Print #F, Lb1(AgWa - 1)
    Print #F, i
Close #F
'On Error GoTo 0
        
Call DoNoPointers
Call SaveCount

AgWa = 0

On Error Resume Next
Kill App.Path + "\#00 Loggs\Error Pix.txt"

'DO THIS RIGHT MAYBE WONT NEED A END
For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next


Kill RamDrive + ":TEMP\ART_PROG_TRIP_WIRE.txt"


End

End Sub



'================================================================================
'                        LINKING PICTURE TO SCROLLBARS
'================================================================================
Private Sub AdjustScrollBars(TheImage As cImage)
    If Me.WindowState > 0 Then Exit Sub
    'Exit Sub

    Dim x As Long 'Set Max/Min/Visible properties of HScroll1 and VScroll1
    Dim y As Long '    for TheImage in Picture1


    If ObjPtr(TheImage) = 0 Then 'Remove HScroll and VScroll
        HScroll1.Min = 0
        HScroll1.Max = 0
        HScroll1.Visible = False
        VScroll1.Min = 0
        VScroll1.Max = 0
        VScroll1.Visible = False
    Else
        If Picture1.Width >= VScroll1.Width + 4 * TwipsPerPixel And Picture1.Height >= HScroll1.Height + 4 * TwipsPerPixel Then 'PictureBox larger than ScrollBars

'Picture1.Top = (fMain.Width / Screen.TwipsPerPixelX) - HScroll1.Value
'Picture1.Left = (fMain.Height / Screen.TwipsPerPixelY) - VScroll1.Value

    x = Picture1.Width \ TwipsPerPixel - 4
    y = Picture1.Height \ TwipsPerPixel - 4
    If TheImage.Width > x Then
        y = y - HScroll1.Height \ TwipsPerPixel
        If TheImage.Height > y Then x = x - VScroll1.Width \ TwipsPerPixel
    Else
        If TheImage.Height > y Then
            x = x - VScroll1.Width \ TwipsPerPixel
            If TheImage.Width > x Then y = y - HScroll1.Height \ TwipsPerPixel
        End If
    End If

    If TheImage.Width > x Then    'Add HScroll and set HScroll limits
        HScroll1.Min = 0
        HScroll1.Max = TheImage.Width - x
        HScroll1.Move 0, Picture1.Height - HScroll1.Height - 4 * TwipsPerPixel, Picture1.Width - IIf(TheImage.Height > y, VScroll1.Width, 0) - 4 * TwipsPerPixel
        'HScroll1.Visible = True
    Else                          'Remove HScroll and center picture
        'Search Honey
        HScroll1.Visible = False
'        HScroll1.Min = (TheImage.Width - Picture1.Width \ TwipsPerPixel + 4 + IIf(TheImage.Height > Y, VScroll1.Width \ TwipsPerPixel, 0)) \ 2
        HScroll1.Min = (TheImage.Width - Picture1.Width \ TwipsPerPixel + IIf(TheImage.Height > y, VScroll1.Width \ TwipsPerPixel, 0)) \ 2
        HScroll1.Max = HScroll1.Min
    End If

    If TheImage.Height > y Then   'Add VScroll and set VScroll limits
        VScroll1.Min = 0
        VScroll1.Max = TheImage.Height - y
        VScroll1.Move Picture1.Width - VScroll1.Width - 4 * TwipsPerPixel, 0, VScroll1.Width, Picture1.Height - 4 * TwipsPerPixel
        'VScroll1.Visible = True
    Else                          'Remove VScroll and center picture
        VScroll1.Visible = False
        VScroll1.Min = (TheImage.Height - Picture1.Height \ TwipsPerPixel + 4 + IIf(HScroll1.Visible, HScroll1.Height \ TwipsPerPixel, 0)) \ 2
        VScroll1.Max = VScroll1.Min
    End If


        End If
    End If

    PaintImage m_Image

End Sub
Private Sub PaintImage(TheImage As cImage)
    If ObjPtr(TheImage) = 0 Then
        'Picture1.Cls
    Else
        If HScroll1.Value < 0 Or VScroll1.Value < 0 Then Picture1.Cls
        'TheImage.PaintHDC Picture1.HDC, -HScroll1.Value, -VScroll1.Value
        TheImage.PaintHDC Picture1.HDC, 0, 0
        Picture1.Refresh
'        val1 = (fMain.Width / Screen.TwipsPerPixelX) - HScroll1.Value
'        val2 = (fMain.Height / Screen.TwipsPerPixelY) - VScroll1.Value
'        TheImage.PaintHDC Picture1.hdc, -val1, -val2
'        Picture1.Refresh
    End If
End Sub







Private Sub Lab_Click(index As Integer)

If InStr(Lab(index).Caption, "Hitt to") > 0 Then
    KeyCode2 = 32
    Call KeyCodeSub_Intercept
    Exit Sub
End If

'<º)))> Forward Mode <(((º>
'<º)))> Reverse Mode <(((º>
If InStr(Lab(index).Caption, "<º)))>") > 0 Then
    Call ReverseModeSwap
    Exit Sub
End If

If Lab(index).Caption = "1st Pic && Stop" Then
    CountPic = 1
    FileInfoTimer.enabled = False
    Call PauseOnOffStat
    DontCount = True
    Call FileInfoTimer_Timer
End If

If Lab(index).Caption = "Last Pic && Stop" Then
    CountPic = ScanPath.ListView1.ListItems.Count
    FileInfoTimer.enabled = False
    Call PauseOnOffStat
    DontCount = True
    Call FileInfoTimer_Timer
End If

If Lab(index).Caption = "Randomizor" Then
    'Sodomizer Real Rand
    Randomize Timer + Int(Now)
    CountPic = Int(Rnd * ScanPath.ListView1.ListItems.Count) + 1
    DontCount = True
    Call FileInfoTimer_Timer
End If

'INDEX=7
If Lab(index).Caption = "Randomizor Forward" Then
    'Sodomizer Rand Forward
    Randomize Timer + Int(Now)
    CountPic = CountPic + Int((Rnd * (ScanPath.ListView1.ListItems.Count - CountPic)) / 2)
    DontCount = True
    Call FileInfoTimer_Timer
End If

If Lab(index).Caption = "Randomizor Reverse" Then
    'Sodomizer Rand Forward
    Randomize Timer + Int(Now)
    CountPic = CountPic - Int((Rnd * CountPic) / 2)
    DontCount = True
    Call FileInfoTimer_Timer
End If

If Lab(index).Caption = "Next Folder" Then
    'PicFilesOffFolderLast
    CountPic = PicFilesOffFolderNX + 1
    DontCount = True
    Call FileInfoTimer_Timer
End If

If Lab(index).Caption = "Rotate 90^ && Save" Then
    On Error Resume Next
    FR5 = FreeFile
    Open App.Path + "\Last Error Pix.txt" For Output As #FR5
    Close #FR5
    Kill App.Path + "\Error Pix.txt"
    On Error GoTo 0
    
    FileInfoTimer.enabled = False
    Call PauseOnOffStat
    Lab(index).BackColor = LabelCol11.BackColor
    'Picture2.Cls
    'Picture2.Top = Picture1.Top
    'Picture2.Left = Picture1.Left
'    Picture2.Width = Picture1.Height
'    Picture2.Height = Picture1.Width
    'Picture2.Width = Picture1.Width
    'Picture2.Height = Picture1.Height
    'Set lPic = LoadPicture(A1 + B1)
    'Picture1.Visible = False
    'Picture2.Visible = True
    'ResizePicture Picture2, lPic
    'Picture1 = LoadPicture()
    
    Call mnuRotate_Click(0) '-90 'usual old one used to work
    Call mnuSave_Click(0) 'then save after if want
    'Call ResizePicture
    'Picture1.Visible = False
    'Picture2.Visible = True
    'Call bmp_rotate(Picture1, Picture2, 90) 'Microsoft Rotate
    'Picture1.Visible = True
    'Picture2.Visible = False
    
    'Picture1 = LoadPicture()
    'Picture1.Cls
    'Picture1.Picture = Picture2.Picture
    'Picture1.Visible = True
'    Picture2.Visible = False
'    Picture2.Visible = False
'    Picture1.Visible = True
'    Picture1.Width = Picture2.Height
'    Picture1.Height = Picture2.Width
    
    Lab(index).BackColor = LabelCol12.BackColor

'Q = PictureRotate(Picture1.Picture, Picture1.Picture, 90)


End If

If Lab(index).Caption = "Load In IceView" Then
    FileInfoTimer.enabled = False
    Call PauseOnOffStat
    'Full Screen Setting
    Call RTB1_Click
    
End If

If Lab(index).Caption = "Faster Than Light" Then
    Slider1.Value = 0
End If

If Lab(index).Caption = "Sleepy Slowly 3s zzz" Then
    Slider1.Value = 30 'Slider1.Max / 3=3secs
End If

End Sub

Sub ReverseModeSwap()
If KeyDir = 1 Then 'Lab(ForBACKCTRL).Caption = "<º)))> Reverse Mode <(((º>" Then
    Call ReverseMode
    Exit Sub
End If

If KeyDir = -1 Then 'Lab(ForBACKCTRL).Caption = "<º)))> Forward Mode <(((º>" Then
    Call ForwardMode
End If
End Sub

Sub ReverseMode()
    
    If Lab(RandCTRL).Caption = "Randomizor Reverse" Then Exit Sub
    Lab(RandCTRL).Caption = "Randomizor Reverse"
    Lab(ForBACKCTRL).Caption = "<º)))> Forward Mode <(((º>"
    
    KeyDir = -1
    Call PauseOnOffStat

    If ScanInProgress = True Then Exit Sub
    FileInfoTimer.enabled = True

End Sub

Sub ForwardMode()
    If Lab(RandCTRL).Caption = "Randomizor Forward" Then Exit Sub
    Lab(RandCTRL).Caption = "Randomizor Forward"
    Lab(ForBACKCTRL).Caption = "<º)))> Reverse Mode <(((º>"
    'LBB3 = "Go Fourth && Multiply"
    'LBB3 = "The Force"
    'The Force
    KeyDir = 1
    Call PauseOnOffStat
    If ScanInProgress = True Then Exit Sub
    FileInfoTimer.enabled = True
End Sub

Private Sub LBB2_Change2()
    'Controls Cursor on Picture
    'LBB2.Visible = False
    'Frame2.Visible = False
    
    LBB2.AutoSize = False
    LBB2.Top = 0
    LBB2.Left = 0
    LBB2.Width = Me.Width - (Picture1.Left + Picture1.Width) - 150
    Frame2.Height = LBB2.Height - 10
    Frame2.Width = LBB2.Width 'LBB2.Width
    'Frame2.Top = fMain.Top - 500 '+ 680
    'Frame2.Left = fMain.Width / 2 - Frame2.Width / 2
    
    Frame2.Top = fmain.Height - Frame2.Height - 750
    Frame2.Left = fmain.Width - Frame2.Width - 140
    If LBB2 <> "" Then
        'Frame2.Visible = True
        'LBB2.Visible = True
    End If
    'Call LBB4_Change
End Sub

Private Sub LBB1_Change()
LBB1.Refresh
'fMain.Refresh

End Sub

Private Sub LBB2_Click()
'LBB2
End Sub

Private Sub LBB3_Change()
    'THE FORCE
    
    If MNU_FULL_SCREEN_YES_VAR = True Then Exit Sub
    
    'PAUSED - The Force - REVERSE
    'If ReStarted = False Then Exit Sub
    LBB3.Top = 0
    LBB3.Left = 0
    'Frame3.Height = LBB3.Height - 10
    Frame3.Height = LBB3.Height 'Frame5.Height
    Frame3.Width = LBB3.Width - 15
    Frame3.Top = Lb2(0).Top '- 480  '+ 680
    Frame3.Left = (Lb2(0).Left - Frame3.Width) - 10
    Frame3.Visible = True
    
    LBB3.Visible = True

End Sub

Private Sub LBB3_Click()

Clipboard.Clear

Dim TFF

TFF = "Now --" + Format(Now, "DD-MM-YY - HH:MM:SS") + vbCrLf
TFF = TFF + "Image Date -- " + LBX(2).Caption + vbCrLf  'DATE OF PIC
TFF = TFF + "File Path -- " + A1 + B1 + vbCrLf
Clipboard.SetText TFF



End Sub

Private Sub LBB5_Change()
    
    
    
    'The TIME DATE LA LA INFO
    'ONSCREEN ON PICTURE
    
    'If ReStarted = False Then Exit Sub
    LBB5.Top = 0
    LBB5.Left = 0
    
    
    Frame5.Width = LBB5.Width
        Frame5.Height = LBB5.Height - 10
    
    If MNU_FULL_SCREEN_YES_VAR = False Then
        
        Frame5.Left = (Lab(0).Width) + 30  ' - Frame5.Width) - 10
        Frame5.Top = Lb2(0).Top '- 480  '+ 680
        
    Else
    
        
        'Frame5.Left = 0
        Frame5.Left = RFLb.Width
        
        
        Frame5.Top = 0 'PBarDUPE.Top + PBarDUPE.Height + 15
        
        Slider1.Width = 1000
        Slider1.Height = Frame5.Height - 20
        Slider1.Left = Frame5.Left + Frame5.Width + 15

        PBarDUPE.Left = LBB5.Width + 15
        PBarDUPE.Left = Slider1.Left + Slider1.Width + 15

        On Error Resume Next
        PBarDUPE.Width = (((Screen.Width - 40)) - Slider1.Left) - Slider1.Width
        On Error GoTo 0
        'lbb5
        'Me.Width
    End If

    
    Frame5.Visible = True
    LBB5.Visible = True

End Sub


Private Sub LBB4_Change2()
    'Time
    LBB4.Visible = False
    Frame4.Visible = False
    
    If LBB4.Caption = "" Then LBB4 = "DDD DD-MMM-YYYY HH:MM:SSa/p"
    DoEvents
    LBB4.AutoSize = False
    LBB4.Top = 0
    LBB4.Left = 0
    'lbb4.Caption
    'LBB4.Height = 800
    
    LBB4.BackColor = LBB2.BackColor
    LBB4.ForeColor = LBB2.ForeColor
    LBB4.Width = LBB2.Width
    Frame4.Height = LBB4.Height ' - 300
    Frame4.Width = LBB4.Width
    Frame4.Top = Frame2.Top - Frame4.Height - 10
    Frame4.Left = fmain.Width - Frame4.Width - 140
    
    'Frame4.Visible = True
    'LBB4.Visible = True
'    If LBB2.Caption = "Controls" Then LBB2_Change '--- Control picture
End Sub

Private Sub LBB4_Click()
'LBB4
End Sub

Private Sub LBB5_Click()
'LBB5
    
If MNU_FULL_SCREEN_YES_VAR = True Then Unload Me: Exit Sub

End Sub

Private Sub LBX_Click(index As Integer)
'LBX
End Sub

Private Sub Mnu_1_Click(index As Integer)
    
    If index >= 1 And index <= 10 Then
        Cmd$ = Dir2.List(Dir2.ListCount - (index)) + "\"
    End If

    If InStr(Mnu_1(index).Caption, "WebCam") > 0 Then
        Cmd$ = Dir3.List(Dir3.ListCount - (index - 10)) + "\"
    End If


    If Mnu_1(index).Caption = "# Today FeedDemon" Then
        Cmd$ = "M:\0 00 Art\My FeedStation Podcasts\## TODAY"
    End If
    
    
    If Mnu_1(index).Caption = "-:¦:-•:*'''''*:•-:¦:-" Then
        Cmd$ = "M:\0 00 Art\01 AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0000-Cartoon"
    End If
    
    AgWa = 2
    QueUpIndex = 1
    OldIndex = -1
    
    
    
    ScanPath.ListView1.ListItems.Clear

    
    Call Lb1_Click(0)
    
    FR5 = FreeFile
    Open App.Path + "\#00 Loggs\Command.txt" For Output As #FR5
    Print #FR5, Cmd$
    Close #FR5

    Call ScanFiles

End Sub

Private Sub MNU_ART_LOGG_FILES_Click()

Shell "NOTEPAD.EXE " + App.Path + "\" + GetComputerName + "\00_ART_Data\ART PICTURE VIEW LOGG.txt", vbMaximizedFocus

End Sub

Private Sub MNU_ART_LOGG_FOLDER_Click()

Shell "NOTEPAD.EXE " + App.Path + "\" + GetComputerName + "\00_ART_Data\ART FOLDER VIEW LOGG.txt", vbMaximizedFocus

End Sub

Private Sub Mnu_AutoStart_Click()

If Mnu_AutoStart.Checked = True Then Mnu_AutoStart.Checked = False: Exit Sub
If Mnu_AutoStart.Checked = False Then Mnu_AutoStart.Checked = True

End Sub

Private Sub Mnu_DelKey_Click()

If Mnu_DelKey.Checked = True Then Mnu_DelKey.Checked = False: Exit Sub
If Mnu_DelKey.Checked = False Then Mnu_DelKey.Checked = True

End Sub

Sub BackTOOrginalFromFullScreen()
    
    If Me.WindowState <> 0 Then Exit Sub
    
    Me.WindowState = 0
    
    MNU_FULL_SCREEN_YES_VAR = False
    
    NotAlwaysOnTop (Me.hwnd)
    'AlwaysOnTop (Me.hwnd)
        
    'PUT IN PAUSE
    FileInfoTimer.enabled = False
    Call PauseOnOffStat
        
        
    'Final Set Pic Boz Size
    fmain.Top = 500
    fmain.Height = Screen.Height - 1000
    fmain.Width = Screen.Width
    Picture1.Top = 15                                               'Hair Line Away from top
    Picture1.Height = fmain.Height - 770                   'get this right prog dont chnage much
    Picture1.Left = RTB1.Left + RTB1.Width + 15      'and left side
    Picture1.Width = (Lb2(0).Left - Picture1.Left) - 15
    
    RFLb.Visible = False
            
        
    fmain.BackColor = &H400000
    Picture1.BackColor = fmain.BackColor
        
    RTB1.Visible = True
    RTB2.Visible = True
    RTB3.Visible = True
    Slider1.Visible = True

    Slider1.Width = Slider1_Var_Width
    Slider1.Height = Slider1_Var_Height
    Slider1.Left = Slider1_Var_Left
    Slider1.Top = Slider1_Var_Top
    PBar1.Visible = True
    PBar2.Visible = True
    
    LBB5.Caption = TTQX
        
    'CHANGE AND PUT BACK
    XXD = LBB3
    LBB3 = "22"
    LBB3 = XXD
    
    Frame2.Visible = True
    Frame3.Visible = True
    Frame4.Visible = True
    Call LBB5_Change
    
    PBarDUPE.Visible = False
    LBB3.Visible = True
        
    For Each Control In Lb1
        Control.Visible = True
        Lb2(Control.index).Visible = True
        If Control.Caption = "WebCam 1 Day" Then Exit For
    Next
    For Each Control In Lab
        Control.Visible = True
    Next
    
    'For Each Control In LBX
    '    Control.Visible = True
    'Next

    'PAUSE WHEN BACK FROM FULL SCREEN OPTIONAL
    FileInfoTimer.enabled = False
    
    FileInfoTimer.enabled = True
    Call PauseOnOffStat

    Call DisplayPic
    Me.Refresh

End Sub


Private Sub MNU_DONT_USE_CLICK_FOR_PAUSE_Click()

MNU_DONT_USE_CLICK_FOR_PAUSE.Checked = Not _
    MNU_DONT_USE_CLICK_FOR_PAUSE.Checked



End Sub

Private Sub MNU_FULL_SCREEN_YES_Click()

    If Me.WindowState <> 0 Then Exit Sub
    
    
    MNU_FULL_SCREEN_YES_VAR = True


    'PUT IN PAUSE
    FileInfoTimer.enabled = False
    Call PauseOnOffStat

    fmain.BackColor = 0
    Picture1.BackColor = fmain.BackColor

    'XV = FindWindow("WindowsForms10.Window.8.app2", VBNULSTRING)
    'NotAlwaysOnTop (XV)

    AlwaysOnTop (fmain.hwnd)
    'FORM TOP
    Me.Top = -670
    Me.Left = -40
    Me.Height = Screen.Height + 200
    Me.Width = Screen.Width + 100

    RTB1.Visible = False
    RTB2.Visible = False
    RTB3.Visible = False
    PBar1.Visible = False
    PBar2.Visible = False
    Slider1.Visible = True


    For Each Control In Lab
        Control.Visible = False
    Next
    'For Each Control In LBX
    '    Control.Visible = False
    'Next
    For Each Control In Lb1
        Control.Visible = False
        Lb2(Control.index).Visible = False
        If Control.Caption = "WebCam 1 Day" Then Exit For
    Next

    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False

    Me.WindowState = 0
    Debug.Print "MINIIMISED"
    
    LBB3.Visible = False

    If InStr(LBX(0), "^") > 0 Then
        If LBX(2) <> "" Then
            'THE DATECHANGE EVENTS ONLY HAPPENS
            'ON CHANGE DATE LOT FILES ARE SAME DATE
            LBB5.Caption = Format(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " " + Trim(Mid(LBX(0), InStr(LBX(0), "^")))
            LBB5.Font.Name = "Courier New"
            LBB5.Font.Size = 12
            LBB5.FontBold = True
        End If
    End If

    RFLb.Visible = True
    RFLb.Top = 0
    RFLb.Left = 0
    Call LBB5_Change

    RFLb.FontSize = LBB5.FontSize - 2

    'PBarDUPE.Left = LBB5.Width + 15

    PBarDUPE.Left = Slider1.Width + Slider1.Left + 15

    Slider1.Width = 1000
    Slider1.Top = 0
    Slider1.Height = Frame5.Height - 20
    Slider1.Left = Frame5.Width + Frame5.Left + 15

    On Error Resume Next
    'PBarDUPE.Width = (Me.Width - 20)
    PBarDUPE.Width = (((Screen.Width - 40)) - Slider1.Left) - Slider1.Width
    On Error GoTo 0

    PBarDUPE.Visible = True
    PBarDUPE.Height = Frame5.Height - 20
    PBarDUPE.Top = 0

    'PBarDUPE.Left = LBB5.Width + 15
    PBarDUPE.Left = Slider1.Left + Slider1.Width + 15

    On Error Resume Next
        'PBarDUPE.Width = (Me.Width - 20)
        PBarDUPE.Width = (((Screen.Width - 40)) - Slider1.Left) - Slider1.Width
    On Error GoTo 0


    Picture1.Top = PBarDUPE.Top + PBarDUPE.Height + 15
    Picture1.Height = Me.Height - (Picture1.Top + 750)
    Picture1.Left = -20
    Picture1.Width = Me.Width

    FileInfoTimer.enabled = True
    Call PauseOnOffStat

    'has counter changed put back one increment first then if so CHECK
    'If InStr(LBB3, "Pause") > 0 Then
    
    Call DisplayPic
    'End If
    'Sleep 5000
    Me.Refresh

End Sub

Private Sub Mnu_FullScreen_Start_Click()

    If Mnu_FullScreen_Start.Checked = True Then Mnu_FullScreen_Start.Checked = False: Exit Sub
    If Mnu_FullScreen_Start.Checked = False Then Mnu_FullScreen_Start.Checked = True

End Sub

Private Sub MNU_GET_INTERNET_PICS_Click()

Timer7.enabled = True
Shell "D:\VB6\VB-NT\00_Best_VB_04\INet_Content_Jpgs\INet_Content_Jpgs.exe -Quite", vbNormalNoFocus

End Sub

Private Sub Mnu_IceView_Click()
Call RTB1_Click
End Sub

Private Sub Mnu_Kids_Click()

If Mnu_Kids.Checked = True Then Mnu_Kids.Checked = False: Exit Sub
If Mnu_Kids.Checked = False Then Mnu_Kids.Checked = True


End Sub

Private Sub Mnu_loggs_Click()
    FileInfoTimer.enabled = False
    Call PauseOnOffStat
    Shell "Explorer.exe /e, " + App.Path + "\#00 Loggs", vbNormalFocus
End Sub

Private Sub Mnu_Mic_Rotate_Click()
    On Error Resume Next
    FR5 = FreeFile
    Open App.Path + "\Last Error Pix.txt" For Output As #FR5
    Close #FR5
    Kill App.Path + "\Error Pix.txt"
    On Error GoTo 0
    
    FileInfoTimer.enabled = False
    Call PauseOnOffStat
'    Lab(index).BackColor = LabelCol11.BackColor
    'Picture2.Cls
    Picture2.Top = Picture1.Top
    Picture2.Left = Picture1.Left
'    Picture2.Width = Picture1.Height
'    Picture2.Height = Picture1.Width
    Picture2.Width = Picture1.Width
    Picture2.Height = Picture1.Height
    Set lPic = LoadPicture(A1 + G1)
    Picture1.Visible = False
    Picture2.Visible = True
    ResizePicture Picture2, lPic
    Picture1 = LoadPicture()
    
    'Call mnuRotate_Click(0) '-90 'usual old one used to work
    'Call mnuSave_Click(0) then save after if want
    
    Picture1.Visible = False
    Picture2.Visible = True
    Call bmp_rotate(Picture1, Picture2, 90) 'Microsoft Rotate
    Picture1.Visible = True
    Picture2.Visible = False
    
    'Picture1 = LoadPicture()
    'Picture1.Cls
    'Picture1.Picture = Picture2.Picture
    'Picture1.Visible = True
'    Picture2.Visible = False
'    Picture2.Visible = False
'    Picture1.Visible = True
'    Picture1.Width = Picture2.Height
'    Picture1.Height = Picture2.Width
    
'    Lab(index).BackColor = LabelCol12.BackColor

'Q = PictureRotate(Picture1.Picture, Picture1.Picture, 90)



End Sub


Private Sub MNU_MINIIMIZE_STANDBY_Click()
    If Mnu_MNU_MINIIMIZE_STANDBY.Checked = True Then Mnu_MNU_MINIIMIZE_STANDBY.Checked = False: Exit Sub
    If Mnu_MNU_MINIIMIZE_STANDBY.Checked = False Then Mnu_MNU_MINIIMIZE_STANDBY.Checked = True
End Sub

Private Sub Mnu_Pause_When_Not_In_Focus_Click()
    If Mnu_Pause_When_Not_In_Focus.Checked = True Then Mnu_Pause_When_Not_In_Focus.Checked = False: Exit Sub
    If Mnu_Pause_When_Not_In_Focus.Checked = False Then Mnu_Pause_When_Not_In_Focus.Checked = True
End Sub

Private Sub MNU_RAND_A_MINUTE_Click()
MNU_RAND_A_MINUTE.Checked = Not MNU_RAND_A_MINUTE.Checked
RANMIN = Now + TimeSerial(0, 0, 15)
End Sub

Private Sub Mnu_ShowPic_Click()
    'T1 = FileInfoTimer.Enabled
    'FileInfoTimer.Enabled = False
    'Timers are set to off on this run
'    Unload ShowPicFrm
    ShowPicFrm.Show
End Sub

Private Sub Mnu_StartSame_Click()

If Mnu_StartSame.Checked = True Then Mnu_StartSame.Checked = False: Exit Sub
If Mnu_StartSame.Checked = False Then Mnu_StartSame.Checked = True

End Sub

Private Sub Mnu_WebCam_Click()
    FileInfoTimer.enabled = False
    Call PauseOnOffStat
    Shell "C:\WINDOZE\amcap.exe", vbNormalFocus
End Sub

Private Sub Mnu_WebCam3_Click()
    FileInfoTimer.enabled = False
    Call PauseOnOffStat
    Shell "D:\VB6\VB-NT\00_Best_VB_04\Webcam_Motion\WebCamMatts.exe", vbNormalFocus
End Sub

Private Sub Mnu_WebCam4_Click()
    FileInfoTimer.enabled = False
    Call PauseOnOffStat
Shell "D:\VB6\VB-NT\00_Best_VB_04\Vcap_Src_WebCam\vbVidCapWebCam.exe", vbNormalFocus
End Sub



Private Sub PBarDUPE_Click()
     
     If Me.WindowState <> 0 Then Exit Sub

     If MNU_FULL_SCREEN_YES_VAR = True Then
        Call BackTOOrginalFromFullScreen
        Exit Sub
    End If

End Sub

Private Sub Picture1_DblClick()
     If Me.WindowState <> 0 Then Exit Sub

     If MNU_FULL_SCREEN_YES_VAR = True Then
        Call BackTOOrginalFromFullScreen
        Exit Sub
    End If

     If MNU_FULL_SCREEN_YES_VAR = False Then
        Call MNU_FULL_SCREEN_YES_Click
        Exit Sub
    End If

End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
'not use problem
If KEYRELAX = 1 Then Exit Sub
KEYRELAX = 1
    
    'Exit Sub
    'PICTURE1
    'NOT RESPON CEIVCE
    If KeyCode = 32 Then
        'Stop
        Debug.Print "Keycode 32 to Pause Fault Aciedent Press"
    End If
    'If KeyCode <> 32 Then Exit Sub
    
    KeyCode2 = KeyCode
    Call KeyCodeSub_Intercept

End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
    
    Exit Sub
    'PICTURE1
    'NOT RESPON CEIVCE
    If KeyAscii = 32 Then
        'Stop
        Debug.Print "Keycode 32 to Pause Fault Aciedent Press"
    End If
    'If KeyCode <> 32 Then Exit Sub
    
    KeyCode2 = KeyAscii
    Call KeyCodeSub_Intercept
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)

KEYRELAX = 0

End Sub

Private Sub Picture1_LostFocus()
    On Error Resume Next
    Picture1.SetFocus
    
End Sub

Private Sub RFLb_Click()
'If Lab(index).Caption = "Randomizor Forward" Then
    'Sodomizer Rand Forward
    Randomize Timer + Int(Now)
    CountPic = CountPic + Int((Rnd * (ScanPath.ListView1.ListItems.Count - CountPic)) / 2)
    DontCount = True
    Call FileInfoTimer_Timer
'End If

End Sub

Private Sub RTB1_Click()
'/n  Opens a new single-pane Window for the default selection.
'This is usually the root of the drive on which Windows is installed.
'/e Starts Windows Explorer using its default view.
'/e, <object> Starts Windows Explorer using its default view with the focus on the specified folder.
'/root, <object> Opens a window view of the specified object.
'/select, <object> Opens a window view with the specified folder, file or
'program selected.

'FileInfoTimer.enabled = False

'Shell "explorer /select, " + A1 + b1, vbNormalFocus
FileInfoTimer.enabled = False
Call PauseOnOffStat
    
Shell "C:\Program Files\IceView\iceview.exe """ + A1 + G1 + """", vbNormalFocus

End Sub

Private Sub RTB2_Change()
    If OLDA2$ <> RTB2.Text And EasyNow = 1 And Len(RTB2.Text) > 1 Then
        'Our Purpose
        TimerFlashV3 = Now + TimeSerial(0, 0, 2)
        If TimerFlash3.enabled = False Then
            TimerFlash3.enabled = True
            Call TimerFlash3_Timer
        End If
    End If
    If Len(RTB2.Text) < 2 Then Exit Sub
    EasyNow = 1
        
    OLDA2$ = RTB2.Text



End Sub

Sub Mnu_Thumbs_Click()

If ScanInProgress = True Then Exit Sub

FileInfoTimer.enabled = False
Call PauseOnOffStat


'Shell "C:\Program Files\Thumbs4\Thumbs.exe """ + A1 + """", vbNormalFocus
'XX = Left$(GetShortName(RTB2.Text), Len(GetShortName(RTB2.Text)) - 1) '+ "\" '+ "\*.jpg"
'XX = Left$(RTB2.Text, Len(RTB2.Text) - 1)
XX = RTB2.Text + RTB3.Text
XX = GetShortName(XX)
'XX = Left$(XX, Len(XX) - 1)

'ShellAndWait "C:\Program Files\Thumbs4\Thumbs.exe", vbNormalFocus, True

ty = FindWindow("cswThumbsPlusMain", vbNullString)
If ty = 0 Then Shell "C:\Program Files\Thumbs4\Thumbs.exe", vbNormalFocus
tq = Now + TimeSerial(0, 0, 40)
Do
DoEvents
ty = FindWindow("cswThumbsPlusMain", vbNullString)
Sleep 50
Loop Until ty > 0 Or tq < Now

If tq < Now Then msgbox "Progr Didnt Run": Exit Sub

'dumvar = Putfocus(ty)
SetForegroundWindow (ty)


Shell App.Path + "\ThumbsPlusCmdDDE\tpcmd ""LocateFile(""" + XX + """)""", vbHide



'Shell App.Path + "\ThumbsPlusCmdDDE\tpcmd ""OpenDir(""" + XX + """)""", vbHide

'Shell App.Path + "\ThumbsPlusCmdDDE\tpcmd ""OpenDir(""" + XX + """)""", vbHide
'Sleep 500
'Shell App.Path + "\ThumbsPlusCmdDDE\tpcmd ""OpenDir(""" + XX + """)""", vbHide
'Sleep 500
'Shell App.Path + "\ThumbsPlusCmdDDE\tpcmd ""UpdateAll""", vbHide

'ShellAndWait App.Path + "\ThumbsPlusCmdDDE\tpcmd ""OpenDir(""" + XX + """)""", vbHide, True
'ShellAndWait App.Path + "\ThumbsPlusCmdDDE\tpcmd ""UpdateAll""", vbHide, True

'UpdateAll(path)




'XX = GetShortName(RTB2.Text)

'ShellAndWait App.Path + "\ThumbsPlusCmdDDE\tpcmd ""ScanTree(" + XX + ")""", vbHide, True
'tpcmd ScanTree(d:\)

End Sub

Private Sub Mnu_Del_Click()
FileInfoTimer.enabled = False
On Error Resume Next
    'Name A1 + b1 As A1 + b1 + ".HHH"
    Kill A1 + G1
    If Err.Number > 0 Then msgbox "Delete Error"
On Error GoTo 0

ScanPath.ListView1.ListItems.Remove CountPic
'Chk is that right -- double chk YES
'Scanpath.ListView1.ListItems.Item (CountPic)

'Displays Correct Folders Folders Counts if reset this
OldA1 = ""
'Click to next picture
DontCount = True
Call Lbl2_Click

End Sub

Private Sub Mnu_Move_XXXRay_Click()

    FileInfoTimer.enabled = False
    On Error Resume Next
    MkDir A1 + "#B0XxXRay"
    On Error GoTo 0
    Set F = FS.getfile(A1 + G1)
    F.MoveFile A1 + "#B0XxXRay\"
    
    'FS.MoveFile A1 + G1, A1 + "#B0XxXRay\" + B1
    
    'Displays Correct Foles Folders Counts if reset this
    OldA1 = ""
    DontCount = True
    'Click to next picture
    Call Lbl2_Click

End Sub

Private Sub Mnu_MoveGood_Click()
    
    FileInfoTimer.enabled = False
    On Error Resume Next
    MkDir A1 + "#0Good"
    On Error GoTo 0
    Set F = FS.getfile(A1 + G1)
    F.MoveFile A1 + "#Good\"

    'FS.MoveFile A1 + B1, A1 + "#Good\" + B1
    ScanPath.ListView1.ListItems.Remove CountPic
    'On Error Resume Next
    'Kill A1 + b1

    'Displays Correct Foles Folders Counts if reset this
    OldA1 = ""
    DontCount = True
    'Click to next picture
    Call Lbl2_Click

End Sub

Private Sub Mnu_MoveBad_Click()
    
    FileInfoTimer.enabled = False
    On Error Resume Next
    MkDir A1 + "#Bad0"
    On Error GoTo 0
'   FS.MoveFile A1 + B1, A1 + "#Bad0\" + B1
    'Set F = FS.getfile(A1 + G1)
    'F.MoveFile A1 + "#Bad0\"
    FS.MoveFile A1 + G1, A1 + "#Bad0\"
    
    ScanPath.ListView1.ListItems.Remove CountPic
    
    'Displays Correct Foles Folders Counts if reset this
    OldA1 = ""
    DontCount = True
    'Click to next picture
    Call Lbl2_Click

End Sub

Private Sub Mnu_Rotate90Save_Click()
FileInfoTimer.enabled = False
Call Lab_Click(7)
'Search --- "Rotate 90^ && Save"
End Sub

Private Sub Mnu_RunVB_Click()

Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
Unload Me

End Sub

Private Sub Mnu_TrobPix2_Click()
    Dim G1$
        
    FR5 = FreeFile
    Open App.Path + "\#00 Loggs\Last Error Pix.txt" For Input As #FR5
    If Not EOF(FR5) Then Line Input #FR5, G1$
    Close #FR5

    If G1$ <> "" Then
        Shell "explorer /select, " + G1$, vbNormalFocus
    Else
        msgbox "No Pix to do with"
    End If

End Sub

Private Sub Mnu_TrobPix3_Click()
'    Kill App.Path + "\#00 Loggs\Error Pix.txt"
End Sub

'================================================================================
'                  ALLOW GRAB AND DRAG WITH LEFT MOUSE BUTTON
'================================================================================
Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Exit Sub
    If Button = ButtonDrag Then Picture1.MousePointer = 0
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    
    
    Button2 = Button
    XP1 = x: YP1 = y

    If ScanInProgress = True Then Exit Sub
    
    
    If MNU_DONT_USE_CLICK_FOR_PAUSE.Checked = False Then

    If Button = 1 Then
            FileInfoTimer.enabled = Not FileInfoTimer.enabled
            If FileInfoTimer.enabled = False And InDisplay = True Then
                'PauseNow = True
            End If
            Call PauseOnOffStat
        Exit Sub
    End If
    End If
    
    If Button <> 2 Then Exit Sub
    
    'XP1 YP1 mouse posi
    If XP1 < Picture1.Width / 2 And YP1 < Picture1.Height / 2 Then Call Lab_Click(PauseCTRL)
    If XP1 < Picture1.Width / 2 And YP1 > Picture1.Height / 2 Then
        If Slider1.Value > 0 Then
            Slider1.Value = 0
        Else
            Slider1.Value = 3
        End If
    End If
    
    'XP1 YP1 mouse posi
    If XP1 > Picture1.Width / 2 And YP1 < Picture1.Height / 2 Then
    'Call Lab_Click(LASTCTRL)
    CountPic = ScanPath.ListView1.ListItems.Count
    Call Timer1_Timer
    FileInfoTimer.enabled = True
    Call ReverseMode
    Call ReverseMode
    End If
    
    If XP1 > Picture1.Width / 2 And YP1 > Picture1.Height / 2 Then Call Lab_Click(ForBACKCTRL)

End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
            
            
    If x < Picture1.Width / 2 And y < Picture1.Height / 2 Then Hx = 1
    If x < Picture1.Width / 2 And y > Picture1.Height / 2 Then Hx = 2
    If x > Picture1.Width / 2 And y < Picture1.Height / 2 Then Hx = 3
    If x > Picture1.Width / 2 And y > Picture1.Height / 2 Then Hx = 4
            
    If OH <> Hx Then
        If Hx = 1 Then LBB2.Caption = " Pause / Start "
        If Hx = 2 Then LBB2.Caption = " Fast / Slow "
        If Hx = 3 Then LBB2.Caption = " End && Go Back "
        If Hx = 4 Then LBB2.Caption = " Back / Forward "
    End If
    OH = Hx
            
            
    Exit Sub
    If Button And ButtonDrag Then
        Dim NewX As Long
        Dim NewY As Long
        NewX = PaintLeft - (x \ TwipsPerPixel)
        NewY = PaintTop - (y \ TwipsPerPixel)
        If NewX < HScroll1.Min Then NewX = HScroll1.Min Else If NewX > HScroll1.Max Then NewX = HScroll1.Max
        If NewY < VScroll1.Min Then NewY = VScroll1.Min Else If NewY > VScroll1.Max Then NewY = VScroll1.Max
        HScroll1.Value = NewX
        VScroll1.Value = NewY
    End If
End Sub
Private Sub Picture1_Resize()
    'AdjustScrollBars m_Image
End Sub
Private Sub HScroll1_Scroll()
    Exit Sub
    PaintImage m_Image
End Sub
Private Sub HScroll1_Change()
    Exit Sub
    PaintImage m_Image
End Sub



Private Sub RTB2_Click()
Call Mnu_Thumbs_Click
End Sub

Private Sub RTB3_Click()
If ScanInProgress = True Then Exit Sub
'/n  Opens a new single-pane Window for the default selection.
'This is usually the root of the drive on which Windows is installed.
'/e Starts Windows Explorer using its default view.
'/e, <object> Starts Windows Explorer using its default view with the focus on the specified folder.
'/root, <object> Opens a window view of the specified object.
'/select, <object> Opens a window view with the specified folder, file or
'program selected.

Shell "explorer /select, " + A1 + G1, vbNormalFocus
FileInfoTimer.enabled = False
Call PauseOnOffStat

End Sub

Private Sub Slider1_Change()
    
    If ScanInProgress = True Then Exit Sub
    
    A = FileInfoTimer.enabled
    FileInfoTimer.enabled = False
    FileInfoTimer.Interval = Slider1.Value * 100
    If Slider1.Value = 0 Then FileInfoTimer.Interval = 1
    FileInfoTimer.enabled = A
    
    If Slider1.Value / 10 = Int(Slider1.Value / 10) Then
        LBT1 = Format(Slider1.Value / 10, "0")
    Else
        LBT1 = Format(Slider1.Value / 10, "#0.#")
    End If

    If Slider1.Value <> OldSlider Then
        Call DoNoPointers
    End If
    
    OldSlider = Slider1.Value

End Sub

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    FileInfoTimer.enabled = False

End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Slider1.Value / 10 = Int(Slider1.Value / 10) Then
        LBT1 = Format(Slider1.Value / 10, "0")
    Else
        LBT1 = Format(Slider1.Value / 10, "0.#")
    End If
    
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    FileInfoTimer.enabled = True
    If Picture1.Visible = True Then Picture1.SetFocus

End Sub

Private Sub Lb2_Click(index As Integer)
    If ScanInProgress = True Then Exit Sub
    Call Lb1_Click(index)
End Sub

Private Sub Lb1_Click(index As Integer)

If index = OldIndex Then Exit Sub
OldIndex = index

If Mnu_Kids.Checked = True Then
    If Lb1(index) = "(¯`'•.¸ -‹(•¿•)›- ¸.•'´¯)" Then Exit Sub
    If Lb1(index) = "<º)))><¸..·´`·..¸><(((º>" Then Exit Sub
    If Lb1(index) = "-:¦:-•:*'''''*:•-:¦:-" Then Exit Sub
    If AgWa = 0 Then
        For Each Control In Lb1
            'If Control = "WebCam Today" Then
            If Control = "WebCam 1 Day" Then
                AgWa = Control.index + 1
                QueUpIndex = AgWa
                Exit For
            End If
        Next
    End If
End If

Lb1(index).BackColor = LabelCol11.BackColor
Lb1(index).ForeColor = 0
For Each Control In Lb1
    If index <> Control.index Then
    Control.BackColor = LabelCol12.BackColor
    Control.ForeColor = QBColor(15)
    End If
Next

If ScanInProgress = True Then
    QueUpIndex = index + 1
    BreakScanNow = True
    BreakScanX = True   '---------------------------------------------- 'sort here
    ScanPath.ListView1.ListItems.Clear
    Call ScanPath.cmdScan_Click
    
    Do
        DoEvents
        Sleep 30
    Loop Until ScanInProgress = False
    
'    Exit Sub
End If
    
If index = 0 And Cmd$ = "" Then Exit Sub

'To do with RTB
EasyNow = 0
DateOld = 0

ScanInProgress = False
QueUpIndex = index + 1
ScanPath.ListView1.ListItems.Clear

'call it because it might be paused - we may want to keep it that way

ScanFinished = False
BreakScanX = False
BreakScanNow = False

FileInfoTimer.enabled = False
FileInfoTimer.enabled = True
Call PauseOnOffStat


NoFiles = False
Do
    ScanInProgress = False
    Call ScanFiles
Loop Until ScanPath.ListView1.ListItems.Count > 0 Or NoFiles = True

End Sub



Private Sub Lbl1_Click()
    If ScanInProgress = True Then Exit Sub
    Call KeyLeft
End Sub

Private Sub Lbl2_Click()
    'Lbl2
    If ScanInProgress = True Then Exit Sub
    Call KeyRight
End Sub

Sub KeyLeft()
    oKeyDir = KeyDir
    KeyDir = -1
    If Reverse = True Then KeyDir = 1
        Call KeyDone
    KeyDir = oKeyDir
End Sub

Sub KeyRight()
    oKeyDir = KeyDir
    KeyDir = 1
    If Reverse = True Then KeyDir = -1
    Call KeyDone
    KeyDir = oKeyDir
End Sub

Sub KeyDone()
    FileInfoTimer.enabled = False
    Call PauseOnOffStat

    Do
        'DontCount = False
    
        DoEvents
        Call FileInfoTimer_Timer
    Loop Until BadExit = False
    
End Sub
Private Sub Timer_TRIP_WIRE_Timer()

'FR5 = FreeFile
'Open RAMDRIVE+":TEMP\ART_PROG_TRIP_WIRE.txt" For Output As #FR5
'Close #FR5

i = LastModifiedToCurrent(RamDrive + ":TEMP\ART_PROG_TRIP_WIRE.txt")

End Sub

Private Sub Timer1_Timer()
        
        'ScanPathFin = True
        
        'Slider1.Value = 2
        
        If ScanInProgress = True Then Exit Sub
        If ScanFinished = False Then Exit Sub
        If BreakScanX = True Then Exit Sub
        If BreakScanNow = True Then Exit Sub
        'Do
        '    DoEvents
        '    Sleep 200
        'Loop Until ScanInProgress = False Or BreakScanX = True
        
'------------------------------------------------------------
        
        If BreakScanX = True Then
            Call fmain.RunOut
            ScanInProgress = False
            ScanFinished = True
            Exit Sub
        End If
        
'------------------------------------------------------------
        
        If BreakScanNow = True Then
            
                'If ScanPath.ListView1.ListItems.Count = 0 Then
        End If
        'if no slider value set then put fastest
        'If Slider1.Value = 3999 Then Slider1.Value = 0
                
        LoopCounter = 0
                
        PBar3.Value = 100
        ts = Format$((1) * 100, "#0") + "% Loaded"
        LBB1.Caption = ts
                
                
        If ScanPath.ListView1.ListItems.Count = 0 Then
            If QueUpIndex = AgWa Then QueUpIndex = OldAgWa2
            Call fmain.RunOut
            ScanInProgress = False
            Exit Sub
        End If
        
        'Better Var
        'Count of total files
        DD22 = 0
        easypesy = LBB1 + " - "
        x = ScanPath.ListView1.ListItems.Count
        
        If ReadArra6(AgWa - 1) = True Then
        For WE = ScanPath.ListView1.ListItems.Count To 1 Step -1
            If ScanPath.ListView1.ListItems.Count < WE Then BreakScanX = False: Exit Sub
            A1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(1)
            B1 = ScanPath.ListView1.ListItems.Item(WE)
            G1 = ScanPath.ListView1.ListItems.Item(WE).SubItems(2)
            
            If (WE Mod 10) Or WE < 3 Then
        '        QQ = easypesy + " " + Format(((x - WE) / x) * 100, "0") + "%"
                QQ = easypesy + Format((1) * 100, "0") + "% Filtered"
                If QQ <> oqq Then
                    LBB1 = QQ
                    oqq = QQ
                    LBB1.Refresh
                End If
                DoEvents
            End If
            
            DI1 = False
            If InStr(A1, "\#Rot") Then
                ScanPath.ListView1.ListItems.Remove WE
                DI1 = True
            End If
            If InStr(A1, "\#B") Then
                ScanPath.ListView1.ListItems.Remove WE
                DI1 = True
            End If
            If InStr(A1, "\.picasaoriginals") And DI1 = False Then
                ScanPath.ListView1.ListItems.Remove WE
                DI1 = True
            End If
            pq = 0
            If InStr(A1, "Temp Inet Files JPGs") > 0 Then pq = 1
            If InStr(A1, "Mental 01") > 0 Then pq = 1
            If InStr(A1, "Mental 20") > 0 Then pq = 1
            If InStr(A1, "Exchange") > 0 Then pq = 1
            If InStr(A1, "Bucket") > 0 Then pq = 1
            
            If pq = 1 Then
                If InStr(LCase(B1), "thumbnail[1].jpg") And DI1 = False Then
                    ScanPath.ListView1.ListItems.Remove WE
                    DI1 = True
                End If
                If InStr(B1, "tiny[") And DI1 = False Then
                    ScanPath.ListView1.ListItems.Remove WE
                    DI1 = True
                End If
                If InStr(B1, ".small[1].jpg") And DI1 = False Then
                    ScanPath.ListView1.ListItems.Remove WE
                    DI1 = True
                End If
                If InStr(B1, "_s[") And DI1 = False Then
                    ScanPath.ListView1.ListItems.Remove WE
                    DI1 = True
                End If
                'b1 = "article-1204122-05F0F0D3000005DC-41_87x84[1].jpg"
            
                xxk1 = InStr(B1, "[1].jpg")
                xxk2 = 0: xxk3 = 0
                If xxk1 > 0 Then xxk2 = InStrRev(B1, "x", xxk1)
                If xxk2 > 0 Then xxk3 = InStrRev(B1, "_", xxk2)
                If xxk3 > 0 Then
                    xxk4 = xxk1 - (xxk2)
                    xxk5 = xxk2 - (xxk3)
                    If xxk4 <= 4 And xxk5 <= 4 And DI1 = False Then
                        'ScanPath.ListView1.ListItems.Remove we
                    End If
                End If
            End If
        Next
                
        QQ = easypesy + Format((1) * 100, "0") + "% Filtered"
        LBB1 = QQ
        
        End If
        
        
        'article-1204122-05F0F0D3000005DC-41_87x84[1].jpg
        
        If ScanPath.ListView1.ListItems.Count = 0 Then
            If QueUpIndex = AgWa Then QueUpIndex = OldAgWa2
            Call fmain.RunOut
            ScanInProgress = False
            Exit Sub
        End If
        
'        If Lb1(AgWa - 1) = "Fuck Book" Then
'            A = Scanpath.ListView1.ListItems.Count
'            'If A <> 1311 Then Stop
'        End If
        
        Call fmain.RunOut
        ScanInProgress = False

        Lb2(AgWa - 1) = Trim(str(ScanPath.ListView1.ListItems.Count))
        
        Call Slider1_Change 'This calls DoNoPointers
        'Call PauseOnOffStat
        BreakScanX = False
        
        'Reset
        F = FreeFile
        Open App.Path + "\#00 Loggs\Agwa.txt" For Output Lock Write As #F
            Print #F, Lb1(AgWa - 1)
            Print #F, str(Val(FileInfoTimer.enabled))
        Close #F
'------------------------------------------------------------

        If ScanPath.ListView1.ListItems.Count = 0 Then
            'MsgBox "No Files Present - Abort"
            QueUpIndex = OldAgWa2
            'FileInfoTimer.enabled = True
            Call fmain.RunOut
            ScanInProgress = False
            Exit Sub
        End If
        OldAgWa2 = AgWa

        RTB1.Text = FirstDirScanned
        
        If ScanPath.ListView1.ListItems.Count > 0 Then PBar1.Max = ScanPath.ListView1.ListItems.Count
        
        Dim YesD
        FR5 = FreeFile
        'Open App.Path + "\#00 Loggs\CountText.txt" For Input As #FR5
        '    YesD = False
        '    Do
        '        Line Input #FR5, LL1$
        '        If InStr(LL1$, ScanPath.txtPath.Text) Then YesD = True: Exit Do
        '    Loop Until EOF(1)
        'Close #FR5
        
        If YesD = True Then CountPic = Val(Mid$(LL1$, 1, 7)) Else CountPic = 1
        If CountPic = 0 Then CountPic = 1
        On Error Resume Next
        FN1 = Trim(ReadArray(AgWa - 1))
        FN1 = Replace(FN1, ">", "#")
        FN1 = Replace(FN1, "<", "#")
        FN1 = Replace(FN1, ":", "#")
        Err.Clear
        
        FR5 = FreeFile
        Open App.Path + "\#00 Loggs\COUNT-POSI\CountText--" + FN1 + ".txt" For Input As #FR5
            'ERR.NUMBER
            'ERR.DESCRIPTION
            '53 = FILE NOT FOUND -- MAYBE HAS DODGY CHARS TEST LATER
            '52 BAD FILENAME OR NUMBER
            If Err.Number <> 53 And Err.Number > 0 Then msgbox "FILENAME PROBLEM HERE ERR.NUMBER NOT EQUAL TO (53 -- FILE NOT FOUND IS NORMAL)" + vbCrLf + """" + ReadArray(AgWa - 1) + """": Stop
            
            Line Input #FR5, LL1$
            CountPic = Val(Mid$(LL1$, 1, 7))
            If CountPic = 0 Then CountPic = 1
        Close #FR5

    
    
        
    Call DoNoPointers
    Call SaveCount
    DontCount = True
    Call FileInfoTimer_Timer
    
    
    
    If ScanPath.ListView1.ListItems.Count > 0 Then
        'FileInfoTimer.Enabled = True
        Timer1.enabled = False
    End If
    
    'If Toad2 <> -2 And Toad2 <> "" Then FileInfoTimer.Enabled = Toad2: Toad2 = -2
    
    'ONLY START UNPAUSED IF FULL SCREEN
    'If FORMLOADFLAG = True And MNU_FULL_SCREEN_YES_VAR = False Then
    



    
'---------------------------
'COPY FILES
    
    Set FS = CreateObject("Scripting.FileSystemObject")
    'FF = "O:\ART\" + Mid(ScanPath.txtPath.Text, 4)
    'Debug.Print FF
    
  
    
    'Exit Sub
            
    'ScanPath.Show
    
    
            
    'On Error Resume Next
    'R = 4
    'Do
                
     '   t = InStr(R + 1, FF, "\")
    '    If t = 0 Then Exit Do
     '   DD = Mid$(FF, 1, t - 1)
    '    R = t
     '   Err.Clear
     '   If t > 0 Then MkDir DD
   '     YY = Err.Description
    '    uu = Err.Number
    'Loop Until t = 0
    
    On Error GoTo 0
    
    'Me.WindowState = 1
    Dim FC
    On Error Resume Next
    Form1.Show
    DoEvents
    F2 = 0
    Form1.SetFocus
    Form1.Top = 0
    Me.Top = 1000
    osci = -2
    
    Form1.Label1.BackColor = Form1.Label2.BackColor
    On Error Resume Next
    MkDir "M:\0 00 Art WORK IN BOUND"
    
    For R1 = 1 To ScanPath.ListView1.ListItems.Count
            
            If R1 > 5000 Then Exit For
            FC = Format((Int((R1 / 500) + 1) * 500), "0000")
            'If Int((R1 / 500) + 1) * 500 = 5000 Then Exit For
            'TRIM ""
            'STR STRING FROM NUMERIC
            On Error Resume Next
            FF = "M:\0 00 Art WORK IN BOUND\" + Lb1(AgWa - 1) + "\" '+ FC + " " + Lb1(AgWa - 1) + "\"
            FF1 = "M:\0 00 Art WORK IN BOUND\" + Lb1(AgWa - 1)
            
            FF1 = Replace(FF1, ">", "#")
            FF1 = Replace(FF1, "<", "#")
            FF1 = Replace(FF1, ":", "#")
            FF1 = Replace(FF1, "¦", "#")
            
            FF = Replace(FF, ">", "#")
            FF = Replace(FF, "<", "#")
            'FF = Replace(FF, "", "#")
            FF = Replace(FF, "¦", "#")
            
            If FC <> OFC Then MkDir FF1: MkDir FF
            On Error GoTo 0
            'If R Mod 100 = 0 Then DoEvents
            'If ScanPath.ListView1.ListItems.Count <> osci Then Exit For
            If osci <> ScanPath.ListView1.ListItems.Count And osci <> -2 Then Exit For
            osci = ScanPath.ListView1.ListItems.Count
            
            OFC = FC
            A1 = ScanPath.ListView1.ListItems.Item(R1).SubItems(1)
            B1 = ScanPath.ListView1.ListItems.Item(R1)
            F1 = ScanPath.ListView1.ListItems.Item(R1).SubItems(2)
            
            FF2 = GetShortName(A1)
            On Error Resume Next
            If FS.FileExists(FF + B1) = False Then
                FS.CopyFile FF2 + F1, FF + B1
                F2 = F2 + 1
            End If
            On Error GoTo 0
            'ShellFileCopy FF2 + F1, FF + FC + B1
            
            ASD = str(F2) + " DONE of" + str(R1) + " of" + str(ScanPath.ListView1.ListItems.Count) + " --- " + B1
            'Debug.Print ASD
            
            'Form1.Top = 0
            Form1.Label1 = ASD
            'Form1.SetFocus
            DoEvents
    Next
    Form1.Label1.BackColor = Form1.Label3.BackColor

    'msgbox "COPY DONE"
    
    If FORMLOADFLAG = True Then
        FORMLOADFLAG = False
        FileInfoTimer.enabled = Mnu_AutoStart.Checked
        Call PauseOnOffStat
    End If
    Timer1.enabled = False


End Sub


Sub FileInfoTimer_Timer()
    
    If MNU_RAND_A_MINUTE.Checked = True And RANMIN < Now Then
        RANMIN = Now + TimeSerial(0, 0, 15)

        'INDEX=7
        '= "Randomizor Forward" Then
        'INDEX=6
        '= "Randomizor" Then
        
        
        Call Lab_Click(6)
    
    
    End If
    
    x1 = Mnu_Pause_When_Not_In_Focus.Checked
    If GetForegroundWindow <> Me.hwnd And IsIDE = False And x1 = True Then Exit Sub

    'fileinfotimer.Interval
    If Me.WindowState > 0 Then Exit Sub
    If ScanInProgress = True Then
    Exit Sub
    End If
    'breakscanx
    If ScanPath.ListView1.ListItems.Count = 0 Then
'        FileInfoTimer.Enabled = True
'        Timer1.Enabled = False
    Exit Sub 'Stop: Exit Sub
    End If
    
    If StandBy < Now And StandBy > 0 Then
        If MNU_MINIIMIZE_STANDBY.Checked = True Then
            StandBy = 0
        
            FileInfoTimer.enabled = False
            Me.WindowState = 1
            Exit Sub
        End If
    End If
    
    BadExit = False

    If DontCount = False Then CountPic = CountPic + KeyDir
    DontCount = False
    If CountPic > ScanPath.ListView1.ListItems.Count And KeyDir = 1 Then
        
        If WEBCAMRELOAD = True Then
            Call RELOADWEBCAM
        End If
        
        
        CountPic = 1 ' Main Purpose
        LoopCounter = LoopCounter + 1
        
        'Our Purpose
        TimerFlashV1 = Now + TimeSerial(0, 0, 2)
        If TimerFlash1.enabled = False Then
            TimerFlash1.enabled = True
            Call TimerFlash1_Timer
        End If
    End If
    
    If CountPic < 1 And KeyDir = -1 Then
        CountPic = ScanPath.ListView1.ListItems.Count ' Main Purpose
        LoopCounter = LoopCounter + 1
        'Our Purpose
        TimerFlashV1 = Now + TimeSerial(0, 0, 2)
        If TimerFlash1.enabled = False Then
            TimerFlash1.enabled = True
            Call TimerFlash1_Timer
        End If
    End If
    
    'more req chkl
    If CountPic > ScanPath.ListView1.ListItems.Count Then CountPic = 1: LoopCounter = LoopCounter + 1
    If CountPic < 0 Then CountPic = ScanPath.ListView1.ListItems.Count: LoopCounter = LoopCounter + 1
    
    
        If OldA1 = "" Then OldA3 = ""
            
        hh = 0
        For R = 1 To ScanPath.ListView1.ListItems.Count
            A1 = ScanPath.ListView1.ListItems.Item(R).SubItems(1)
            'b1 = ScanPath.ListView1.ListItems.Item(CountPic)
            '1st use of OldA1
            If OldA1 <> A1 Then
                hh = hh + 1
                ccv = "0000000"
                Mid$(ccv, 8 - Len(Trim(str(R))), Len(Trim(str(R)))) = Trim(str(R))
                If hh > UBound(UU2) Then ReDim Preserve UU2(UBound(UU2) + 100)
                UU2(hh) = ccv + " " + A1 + "@"
            End If
            OldA1 = A1
        Next
        
        ccv = "0000000"
        Mid$(ccv, 8 - Len(Trim(str(ScanPath.ListView1.ListItems.Count))), Len(Trim(str(ScanPath.ListView1.ListItems.Count)))) = Trim(str(ScanPath.ListView1.ListItems.Count))
        If hh + 1 > UBound(UU2) Then ReDim Preserve UU2(UBound(UU2) + 1)
        UU2(hh + 1) = ccv
        
        ReDim Preserve UU2(hh + 1)
        'OldA1 = ""
        If MNU_FULL_SCREEN_YES_VAR = False Then
            PBar1.Visible = True
            PBar2.Visible = True
            Slider1.Visible = True
        End If
    
    If ScanPath.ListView1.ListItems.Count = 0 Then Exit Sub
    
    OldAgWa = AgWa
    
    If RTB2.Text <> "" Then TTQ = Mid(RTB2.Text, InStrRev(RTB2.Text, "\", Len(RTB2.Text) - 1))
    TTQ = Replace(TTQ, "\", "")
    If MNU_FULL_SCREEN_YES_VAR = False Then
        LBB5.Caption = "\" + TTQ + "\ -- " + Mid(LBX(2), 1, 13) + " " + Trim(Mid(LBX(2), 13)) + " -- " + LBX(0)
        TTQX = LBB5
        LBB5.Font.Name = "Arial"
        LBB5.Font.Size = 11
    Else
        If InStr(LBX(0), "^") > 0 Then
            If LBX(2) <> "" Then
                DATESTRING = Mid(LBX(2), InStr(LBX(2), " "))
                LBB5.Caption = Format(DateValue(DATESTRING) + TimeValue(DATESTRING), "DDD DD-MMM-YY HH:MM:SSa/p") + " " + Trim(Mid(LBX(0), InStr(LBX(0), "^")))
                LBB5.Font.Name = "Courier New"
                LBB5.Font.Size = 12
                LBB5.FontBold = True
            End If
        End If
    
    End If
  
    
    
    
    On Error Resume Next
    A1 = ScanPath.ListView1.ListItems.Item(CountPic).SubItems(1)
    If Err.Number > 0 Then
        On Error GoTo 0
        CountPic = 1
        A1 = ScanPath.ListView1.ListItems.Item(CountPic).SubItems(1)
    End If
    On Error Resume Next

    B1 = ScanPath.ListView1.ListItems.Item(CountPic)
    G1 = ScanPath.ListView1.ListItems.Item(CountPic).SubItems(2)
    
    '2nd use of olda1
    If OldA3 <> A1 Then
        For R = 1 To UBound(UU2)
            If InStr(UU2(R), A1 + "@") Then
            PicFilesToT = Val(Mid$(UU2(R + 1), 1, 7)) - Val(Mid$(UU2(R), 1, 7))
            PicFilesOffSet = Val(Mid$(UU2(R), 1, 7)) - 1
            PicFilesOffFolderNX = Val(Mid$(UU2(R + 1), 1, 7)) - 1
            PicFilesOffFolderLast = Val(Mid$(UU2(R), 1, 7)) - 1
            DirNo = R
            Exit For
            End If
        Next
        If UBound(UU2) = 2 Then
            PicFilesToT = ScanPath.ListView1.ListItems.Count
        End If
        OldA3 = A1
    End If
    
    
    PBar1.Max = ScanPath.ListView1.ListItems.Count - 1
    PBarDUPE.Max = ScanPath.ListView1.ListItems.Count - 1
    If PicFilesToT > 0 Then PBar2.Max = PicFilesToT
    
    If FS.FileExists(A1 + G1) = True Then
        Set F = FS.getfile(A1 + G1)
        i = F.datelastmodified
        LBX(2) = Mid(Format(i, "DDD"), 1, 2) + Format(i, " D-MMM-YY H:MM:SSa/p")
    
        'We Dont Want People Saying Whats that flash at Start
        If DateOld = 0 Then DateOld = Int(F.datelastmodified)
        If Int(F.datelastmodified) <> DateOld And FirstRun2 = False Then
            DateOld = Int(F.datelastmodified)
            TimerFlashV2 = Now + TimeSerial(0, 0, 2)
            If TimerFlash2.enabled = False Then
                TimerFlash2.enabled = True
                Call TimerFlash2_Timer
            End If
        End If
    
        If F.Size / 1024 < 1000 Then
            LBX(1) = Format(F.Size / 1024, "#,###.0") + " Kb"
        Else
            LBX(1) = Format(F.Size / 1024 ^ 2, "#,###.0") + "Meg"
        End If
    End If
    
    
    Filename = A1 + B1
    RTB2.Text = A1
    RTB3.Text = B1
    If FS.FileExists(A1 + B1) = False Then
        OldA1 = ""
        ScanPath.ListView1.ListItems.Remove CountPic
        OldCountPic1 = -1 'use with do no count updates

        BadExit = True
            Exit Sub
            
    End If
    
    
    
    Lab(2) = Trim(str(CountPic)) + " of" + str(PBar2.Max + 1) + " of" + str(ScanPath.ListView1.ListItems.Count) + " Files"
    Lab(3) = Trim(str(CountPic - PicFilesOffSet)) + " Files of" + str(DirNo) + " Dir of" + str(UBound(UU2)) + " Dir"
    
    On Error Resume Next
    PBar1.Value = CountPic
    PBarDUPE.Value = CountPic
    PBar2.Value = CountPic - PicFilesOffSet
    'PBAR1.Left=
    'PBAR2.LEFT
    'PBAR1.WIDTH
    'PBAR2.WIDTH
    
    
    LoopSTR = str(LoopCounter)
    LBX(0).Caption = Format((CountPic / ScanPath.ListView1.ListItems.Count) * 100, "0.00") + "%" + " ^" + Trim(LoopSTR)

    Lab(2).Refresh
    Lab(3).Refresh
    LBX(0).Refresh
    LBX(1).Refresh

    'Safety in case picture causes proggy crash Pic Removed on reload
    'Some other places this is killed before a oppo
    If KillOkayFlag = True Then
        On Error Resume Next
        FR5 = FreeFile
        Open App.Path + "\#00 Loggs\Error Pix.txt" For Output As #FR5
        Print #FR5, Filename
        Close #FR5
        On Error GoTo 0
    End If

    '-----------
    'CODE TO DETECT CHANGE OF DIR FROM SEARCH THEN ACT ON THAT
    'EXAMPLE HERE IS TO CHANGE SPEED OF SLIDESHOW THEN
    'RECALL LAST SPEED BACK WHEN NO LONGER IN DIR OR CHANGED
    'IT CAN DETECT A CHANGE ON CONSECUTIVE CHANGE TO ANOTHER NEXT
    'IN LINE ON ANOTHER SEARCH
    
    Dim DSG(3, 2), ML
    DSG(1, 2) = 4: DSG(1, 1) = "0-ButterFly"
    DSG(2, 2) = 0: DSG(2, 1) = "AutoPix 0001-Cartoon"
    DSG(3, 2) = 0: DSG(3, 1) = "BLUEMAN"
    SliderOverRide2 = 0
    For ML = 1 To UBound(DSG)
        If InStr(A1 + B1, DSG(ML, 1)) > 0 And OldML <> ML Then
    
            If SliderOverRide1 = -1 Then
                SliderOverRide1 = Slider1.Value
            End If
            Slider1.Value = DSG(ML, 2)
            OldML = ML
        End If
    
        If InStr(A1 + B1, DSG(ML, 1)) > 0 Then
            SliderOverRide2 = ML
        End If
    Next
    
    If SliderOverRide2 <> OldML Then
        Slider1.Value = SliderOverRide1: SliderOverRide1 = -1
        OldML = SliderOverRide2
    End If

    '-----------

    Call DisplayPic
    
    

End Sub

Sub RELOADWEBCAM()
            
            XX = True
            
            DaysBack = 1
            ScanPath.ListView1.ListItems.Clear
            Dir1.Path = "M\0 00 Art Loggers\Web_Cam\"
            R5 = Dir1.ListCount
            OLDFC = 0
            TTH = 0
            Do
                R5 = R5 - 1
                Td = Dir1.List(R5)
                
                ScanPath.txtPath.Text = Td
                Call ScanPath.cmdScan_Click
                
                If ScanPath.ListView1.ListItems.Count <> OLDFC Then
                    TTH = TTH + 1
                End If
                OLDFC = ScanPath.ListView1.ListItems.Count
            
            
            
            Loop Until TTH >= DaysBack Or R5 = 0

            For Each Control In Lb2
                If Control.Visible = True Then
                    cx = Control.index
                    
                End If
            Next
            Lb2(cx) = Trim(str(ScanPath.ListView1.ListItems.Count))

End Sub
Sub DisplayPic()
    
    'if we want to stop on exact pic load we will need mouse an keyboard systemwide event
    'an check after then pic been loaded
    'the vb event mouse is not powerful enough to be raised that way i wont hitt while the process of pic load
    
    InDisplay = True
'    On Error Resume Next
    
'    Picture1.Top = 15                                               'Hair Line Away from top
'    Picture1.Height = fMain.Height - 770                   'get this right prog dont change much
'    Picture1.Left = RTB1.Left + RTB1.Width + 15      'and left side
'    Picture1.Width = (Lb2(0).Left - Picture1.Left) - 15
    Err.Clear
    

    On Error Resume Next
    Err.Clear
    
    TT5 = Format(Now, "DD-MM-YY - HH:MM:SS")
    Print #FR4, TT5 + " -- " + A1 + B1
    
    
    'ERR.NUMBER
    'ERR.DESCRIPTION
    'On Error GoTo 0
    
    
    COUNTTOTAL = COUNTTOTAL + 1
    Lab(0) = "Aw #" + Format(TotallyFiles, "###,###,##0") + ""
    Lab(0) = Lab(0).Caption + " #" + Format(COUNTTOTAL, "###,###,##0") + ""

    
    Set lPic = LoadPicture(A1 + B1)
    'err.description
    'If Err.Number > 0 Then Stop
    
    GoTo SkiP1
    'this bit of formunload is just so i can continue with these events stuff
    DoEvents
    If FormUnLoading = True Then
        Exit Sub
    End If
    DoEvents
    
    If PauseNow = True Then
        InDisplay = False
        PauseNow = False
        DontCount = True
        msgbox "Hello1"
        'call FileInfoTimer_Timer:
        Exit Sub
    End If
SkiP1:
    
    'chk error pic 1st then dimensions
    If Err.Number > 0 Then
        OldA1 = ""
        ScanPath.ListView1.ListItems.Remove CountPic
        OldCountPic1 = -1 'use with do no count updates
        CountPic = CountPic + KeyDir
        
        
        'If InStr(A1, "Autopix") = 0 Then
            If InStr(A1, "Temp Inet Files JPGs") > 0 Then
                On Error Resume Next
                MkDir A1 + "#Bad Pix Error\"
                FS.MoveFile A1 + B1, A1 + "#Bad Pix Error\" + B1
                On Error GoTo 0
            End If
        'End If
        Exit Sub
    End If
    
    
    'chk error pic 1st then dimensions
    lWidth = Picture1.ScaleX(lPic.Width, vbHiMetric, Picture1.ScaleMode)
    lHeight = Picture1.ScaleY(lPic.Height, vbHiMetric, Picture1.ScaleMode)

    OX = Round(lWidth / Screen.TwipsPerPixelX, 0)
    OY = Round(lHeight / Screen.TwipsPerPixelY, 0)

    If OX < 10 Or OY < 10 Then
        OldA1 = ""
        ScanPath.ListView1.ListItems.Remove CountPic
        OldCountPic1 = -1 'use with do no count updates
        CountPic = CountPic + KeyDir
            On Error Resume Next
            MkDir A1 + "#Bad Pix Too Small Dimensions\"
            FS.MoveFile A1 + B1, A1 + "#Bad Pix Too Small Dimensions\" + B1
            On Error GoTo 0
        Exit Sub
    End If

    GoTo Skip2
    
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    
    If PauseNow = True Then
        InDisplay = False
        PauseNow = False
        DontCount = True
        msgbox "Hello2"
        'call FileInfoTimer_Timer:
        Exit Sub
        
    End If
    
Skip2:
    
    PauseNow = False

    
    ResizePicture Me.Picture1, lPic
    Picture1.Refresh
    
    'If MNU_FULL_SCREEN_YES_VAR = True Then
        'Me.SetFocus
    'End If
    
    'DoEvents
    InDisplay = False
    Call DoNoPointers

    Call SaveCount

End Sub

Sub DoNoPointers()
    
    If AgWa <= 0 Then Exit Sub
    If ScanPath.ListView1.ListItems.Count = 0 Then Exit Sub
    
    If ReadArra4(AgWa - 1) = -2 Then Exit Sub
    
    
    '-1 for the ones we care about and ReadArra4(AgWa - 1) = -2 for all rest
    If AgWa > 0 Then
        If Slider1.Value <> ReadArra4(AgWa - 1) And ReadArra4(AgWa - 1) = -1 Then
            ReadArra4(AgWa - 1) = Slider1.Value
            ReadArra3(AgWa - 1) = FirstDirScanned
        End If
        If Slider1.Value <> ReadArra4(AgWa - 1) And ReadArra4(AgWa - 1) >= 0 And ReadArra5(AgWa - 1) = False Then
            ReadArra5(AgWa - 1) = True
            ReadArra3(AgWa - 1) = FirstDirScanned 'just for luck double up on this line as above
        End If
    End If
    
    
    
    If OldCountPic1 = -1 Or CountPic = -1 Then
        OldCountPic1 = CountPic
        Exit Sub
    End If
    'if the countpic jumps coz bad file found and removed from list then this okay coz be reset to current
    'set -1 if any new folders loaded
    'OldCountPic1 = CountPic
    
    DI2 = Abs(OldCountPic1 - CountPic)
    
    'reached ends yes
    If OldCountPic1 = ScanPath.ListView1.ListItems.Count And CountPic = 0 Then DI2 = 0
    If OldCountPic1 = 1 And CountPic = ScanPath.ListView1.ListItems.Count And CountPic Then DI2 = 0
    
    OldCountPic1 = CountPic
            
    If DI2 < 2 Then Exit Sub
    
    If ReadArra4(AgWa - 1) > -2 Then
        ReadArra3(AgWa - 1) = FirstDirScanned
        ReadArra4(AgWa - 1) = Slider1.Value
        ReadArra5(AgWa - 1) = True
    End If
    
End Sub

Sub SaveCount()
        
        
        If AgWa < 1 Or CountPic = 0 Then Exit Sub
        
        
        
        ccv = "0000000"
        Mid$(ccv, 8 - Len(Trim(str(CountPic))), Len(Trim(str(CountPic)))) = Trim(str(CountPic))
        LL1$ = ccv + " " + ScanPath.txtPath.Text
        
        On Error Resume Next
        FN1 = Trim(ReadArray(AgWa - 1))
        FN1 = Replace(FN1, ">", "#")
        FN1 = Replace(FN1, "<", "#")
        FN1 = Replace(FN1, ":", "#")
        Err.Clear
        FR5 = FreeFile
        Open App.Path + "\#00 Loggs\COUNT-POSI\CountText--" + FN1 + ".txt" For Output Lock Write As #FR5
            Print #FR5, LL1$
        Close #FR5
        
        
        FR5 = FreeFile
        Open App.Path + "\#00 Loggs\COUNT-POSI\TOTAL Count--.txt" For Output Lock Write As #FR5
            Print #FR5, COUNTTOTAL
        Close #FR5
        
        
        
'OLDCODE
        'Okay it reads in whole lot but it will not replace if set dont save pointer for that particular folder
'        YesD = False
        
'        Open App.Path + "\#00 Loggs\CountText.txt" For Input As #FR5
'        Do
'            Line Input #FR5, LL1$
            'its inputing all this but we only want a new value for our stuff if its okay or else just
            'this is block if dont want update
'            If InStr(LL1$, ReadArra3(AgWa - 1)) > 0 Then
'                YesD = True
'            End If
            
            'last test done
            'If ReadArra5(AgWa - 1) = True Then Stop
            
'            If ReadArra5(AgWa - 1) = False Then
'                ccv = "0000000"
'                Mid$(ccv, 8 - Len(Trim(str(CountPic))), Len(Trim(str(CountPic)))) = Trim(str(CountPic))
'                LL1$ = ccv + " " + ScanPath.txtPath.Text
'            End If
            
'            vdv = vdv + LL1$ + vbCrLf
            
'            Loop Until EOF(1)
'        Close #FR5
        
        'Add new to end if none is there
'        If YesD = False Then
'            ccv = "0000000"
'            Mid$(ccv, 8 - Len(Trim(str(CountPic))), Len(Trim(str(CountPic)))) = Trim(str(CountPic))
'            LL1$ = ccv + " " + ScanPath.txtPath.Text
'            vdv = vdv + LL1$ + vbCrLf
'        End If
        
        
End Sub

Private Sub Timer2_Timer()
    
    If GetForegroundWindow <> Me.hwnd Then Exit Sub
    On Error Resume Next
    
    'Picture1.SetFocus
    
    'Picture2.SetFocus

End Sub





Private Sub Timer5_Timer()

    LBB4.Caption = Format(Now, "DDD DD-MMM-YYYY") + vbCrLf + Format(Now, "HH:MM:SSa/p")

End Sub

Private Sub Timer6_Timer()
'Call SaveCount

End Sub

Private Sub Timer7_Timer()
    'Grab files from net and chk CRC them Quick
    
    'TEMP WHILE NOT USE WEB
    'Exit Sub
    
    For Each Control In Lb1
        If Control.Caption = "Internet 1 Day" Then
            Control.BackColor = QBColor(14)
            Control.ForeColor = 0
        End If
    Next
    
    i2 = FindWindow(vbNullString, "ScanPath 2.0 - Sort Anything -INet_Content_Jpgs_CRC")
    If i2 > 0 And KK <> 2 Then
        KK = 2
    End If
    If KKTime = 0 And i2 = 0 Then
        KKTime = Now + TimeSerial(0, 1, 0)
    End If
    If i2 > 0 Then KKTime = 0
    If KKTime < Now Then KK = 2
    
    If (i2 = 0 And KK = 2) Or KK = 3 Then
        For Each Control In Lb1
            If InStr(Control.Caption, "Internet 1 Day") > 0 Then
                Control.BackColor = &H800000
                Control.ForeColor = &HFFFFFF
                Timer7.enabled = False
            End If
        Next
    End If

End Sub


Private Sub TimerFlash1_Timer()
'Easy When Get End Slide Show Or Begining If Reverse
LBX(0).BackColor = LabelCol9.BackColor

If TimerFlashV1 < Now Then
    LBX(0).BackColor = &H800000
    TimerFlash1.enabled = False
End If

End Sub

Private Sub TimerFlash2_Timer()
'Easy When Date Change
LBX(2).BackColor = LabelCol9.BackColor

If TimerFlashV2 < Now Then
    LBX(2).BackColor = &H800000
    TimerFlash2.enabled = False
End If

End Sub


Private Sub TimerFlash3_Timer()
'Easy when Folder change

RTB2.BackColor = LabelCol9.BackColor

If TimerFlashV3 < Now Then
    RTB2.BackColor = &H800000
    TimerFlash3.enabled = False
End If


End Sub

Private Sub VScroll1_Scroll()
    PaintImage m_Image
End Sub
Private Sub VScroll1_Change()
    PaintImage m_Image
End Sub





'================================================================================
'                         AUTOSIZE WINDOW TO PICTURE
'================================================================================
Public Sub SetFormSize(TheImage As cImage)
    Dim NewLeft         As Long
    Dim NewTop          As Long
    Dim NewWidth        As Long
    Dim NewHeight       As Long

    If ObjPtr(TheImage) <> 0 Then
        If Me.WindowState = 0 Then
            If TheImage.Width > 0 And TheImage.Height > 0 Then


'    NewWidth = (TheImage.Width + 4) * TwipsPerPixel + 120 + PictureBoxLeft + PictureBoxRight
'    NewHeight = (TheImage.Height + 4) * TwipsPerPixel + 420 + PictureBoxTop + PictureBoxBottom
'    NewLeft = Me.Left + (Me.Width - NewWidth) \ 2
'    NewTop = Me.Top + (Me.Height - NewHeight) \ 2
    NewWidth = Picture1.Width ') * TwipsPerPixel  '+ 120 + PictureBoxLeft + PictureBoxRight
    NewHeight = (TheImage.Height) * TwipsPerPixel ' + 420 + PictureBoxTop + PictureBoxBottom
    NewLeft = Picture1.Left '+ (Me.Width - NewWidth) \ 2
    NewTop = Picture1.Top '+ (Me.Height - NewHeight) \ 2

    If NewHeight > Screen.Height Then
        NewTop = 0
        NewHeight = Screen.Height
        NewWidth = NewWidth + VScroll1.Width
    Else
        If NewTop < 0 Then
            NewTop = 0
        Else
            If NewTop + NewHeight > Screen.Height Then
                NewTop = Screen.Height - NewHeight
            End If
        End If
    End If
    If NewWidth > Screen.Width Then
        NewLeft = 0
        NewWidth = Screen.Width
    Else
        If NewLeft < 0 Then
            NewLeft = 0
        Else
            If NewLeft + NewWidth > Screen.Width Then
                NewLeft = Screen.Width - NewWidth
            End If
        End If
    End If
    'Me.Move NewLeft, NewTop, newWidth, newHeight
    'fMain.top = 100
    'fMain.Left = 0
    'fMain.Height = Screen.Height - 1000
    'fMain.Width = Screen.Width
'    Picture1.Move PictureBoxLeft, PictureBoxTop, newWidth, newHeight
    
    'NewLeft = Me.Left + (Me.Width - NewWidth) \ 2
    'NewTop = (Me.Top + (Me.Height - NewHeight) \ 2) - 530
    
    'Search Honey
'    Picture1.Move NewLeft - 30, NewTop, NewWidth - 120, NewHeight - 620
    Picture1.Move NewLeft, NewTop, NewWidth, NewHeight
    
    
    'Dim xd1 As Long, yd1 As Long
    'PictureBoxLeft = xd1
    'PictureBoxTop = yd1
    'newWidth = (TheImage.Width + 4) * TwipsPerPixel + 120 + PictureBoxLeft + PictureBoxRight
    'newHeight = (TheImage.Height + 4) * TwipsPerPixel + 420 + PictureBoxTop + PictureBoxBottom
    
            End If
        End If
    End If

End Sub





'================================================================================
'                            MINIMAL IMAGE PROCESSING
'================================================================================
Private Sub mnuRotate_Click(index As Integer)
    Dim MyPic As StdPicture
    Dim Filename As String
    Dim ADJUSTX As Long, ADJUSTY As Long
    
    Set MyPic = Nothing
    Set Picture1 = LoadPicture(A1 + B1)
    Set MyPic = LoadPicture(A1 + B1)
    Set Picture1 = LoadPicture()
    Dim XXW As Long
    Dim XXH As Long
    
    XXW = Picture1.Width / Screen.TwipsPerPixelX
    XXH = Picture1.Height / Screen.TwipsPerPixelY
    'Set MyPic = fMain.Picture1.Picture
    Picture1.Cls
    Set m_Image = Nothing
    Set m_Image = New cImage
    ty = m_Image.CopyStdPicture(MyPic)
    If ty = False Then Stop
    If Err.Number > 0 Then Stop
    OrgWidth = m_Image.Width
    OrgHeight = m_Image.Height
    Call ReSizeCords(ADJUSTX, ADJUSTY)
    Set m_Image = m_Image.Resample(ADJUSTX, ADJUSTY)
    'SetFormSize m_Image
    AdjustScrollBars m_Image


    Select Case index
    Case 0
        Set m_Image = m_Image.Rotate(-90)
        QueUp = OrgWidth
        OrgWidth = OrgHeight
        OrgHeight = QueUp
    
    Case 1
        Set m_Image = m_Image.Rotate(90)
        QueUp = OrgWidth
        OrgWidth = OrgHeight
        OrgHeight = QueUp

    Case 2:    Set m_Image = m_Image.Rotate(180)
    End Select
    
    Call ReSizeCords(ADJUSTX, ADJUSTY) ' adjust vars are the result from this
    
    Set m_Image = m_Image.Resample(ADJUSTX, ADJUSTY)
'    Set m_Image = m_Image.Resample(XXH, XXW)

'    SetFormSize m_Image
    'Picture1.Move Picture1.Left, Picture1.Top, XXH, XXW
    AdjustScrollBars m_Image
    
'    Picture1 = LoadPicture()
    Picture1.Cls
    PaintImage m_Image
    'Picture1.Move Picture1.Left, Picture1.Top, XXH, XXW
    'Picture1.Top = 15                                               'Hair Line Away from top
    'Picture1.Height = fMain.Height - 770                   'get this right prog dont chnage much
    'Picture1.Left = RTB1.Left + RTB1.Width + 15      'and left side
    'Picture1.Width = (Lb2(0).Left - Picture1.Left) - 15
    
    'Picture2.Cls
    'Picture2.Left = Picture1.Left
    'Picture2.Top = Picture1.Top
    'Picture2.Width = Picture1.Width
    'Picture2.Height = Picture1.Height
    'Picture2.Picture = Picture1.Picture
    'Picture2.ScaleHeight = Picture1.ScaleHeight
    'Picture2.ScaleLeft = Picture1.ScaleLeft
    'Picture2.ScaleMode = Picture1.ScaleMode
    'Picture2.ScaleTop = Picture1.ScaleTop
    'Picture2.ScaleWidth = Picture1.ScaleWidth
'    Picture2.ScaleX = Picture1.ScaleX(
'    Picture2.ScaleY = Picture1.ScaleY
    
'    ResizePicture Picture1, Picture2
'    AdjustScrollBars m_Image
    
    Picture1.Refresh
    fmain.Refresh
    DoEvents
    Set MyPic = Nothing
    
    
    
Exit Sub
    
    'Set m_Image = New cImage
    'm_Image.CopyStdPicture MyPic
    
'    ADJUSTX = (fMain.Width / Screen.TwipsPerPixelX)
'    ADJUSTY = (fMain.Height / Screen.TwipsPerPixelY) - 60
    Call ReSizeCords(ADJUSTX, ADJUSTY)
    Set m_Image = m_Image.Resample(ADJUSTX, ADJUSTY)
    SetFormSize m_Image
    AdjustScrollBars m_Image
'    SetFormSize MyPic
'    AdjustScrollBars MyPic
    'MyPic.Refresh
    Picture1.Refresh
    'DoEvents
'    fMain.Refresh
    
    'Set MyPic = Nothing
    

'    If mnuAutosize.Checked Then SetFormSize m_Image
'    AdjustScrollBars m_Image
End Sub
Private Sub mnuMirror_Click(index As Integer)
    Set m_Image = m_Image.Mirror(index <> 0)
    PaintImage m_Image
End Sub
Private Sub mnuAutosize_Click()
    mnuAutosize.Checked = Not mnuAutosize.Checked
    If mnuAutosize.Checked Then SetFormSize m_Image
End Sub





'================================================================================
'                             LOAD / SAVE PICTURE
'================================================================================
Private Sub mnuOpen_Click()
    Call Timer1_Timer
    Exit Sub
    
    Dim MyPic As StdPicture
    'Dim FileName As String

'    FileName = FileDialog(Me, False, "Open Picture File", "Picture Files|*.jpg;*.jpeg;*.gif;*.bmp;*.wmp;*.rle;*.cur;*.ico;*.emf|All Files [*.*]|*.*")
    Filename = "D:\MY WEBS\MatthewLan.Com Web\Photos\galleries\Funny-02\Fw- SMILE !~~image (2).jpg"
   'Call fSaveJPG.PicShow(FileName, PicShowReSize)
    Filename = "D:\MY WEBS\MatthewLan.Com Web\Photos\galleries\Fractals\23november2003d.jpg"
    
    'good code
    'Call PicShow(FileName, PicShowReSize)

    'Unload fMain
    'PicShowReSize.Showj
    'Exit Sub
    If Len(Filename) > 0 Then
        On Error Resume Next
        Set MyPic = LoadPicture(Filename)
        
        
        If Err.Number = 0 Then
            Set m_Image = New cImage
            m_Image.CopyStdPicture MyPic
            'Call PicShow(FileName, PicShowReSize)
            Dim ADJUSTX As Long, ADJUSTY As Long
            
            'ADJUSTX = (fMain.Width / Screen.TwipsPerPixelX) - 30 'LabelCol5.Width '/ Screen.TwipsPerPixelX
            'ADJUSTY = fMain.Height / Screen.TwipsPerPixelY - 50
            Call ReSizeCords(ADJUSTX, ADJUSTY)
            Set m_Image = m_Image.Resample(ADJUSTX, ADJUSTY)

            'Call Resample(400, 400)
            Exit Sub
            
            If mnuAutosize.Checked Then SetFormSize m_Image
            AdjustScrollBars m_Image
            Me.Caption = App.title & " - " & FileTitleOnly(Filename)
            mnuSave(0).enabled = True
        Else
            msgbox "Can not load picture file" & vbCrLf & """" & Filename & """", vbExclamation, "File Load Error"
        End If
        Set MyPic = Nothing
    End If



End Sub

Private Sub mnuSave_Click(index As Integer)
    'Dim FileName As String
    Dim SaveForm As fSaveJPG
    Filename = A1 + B1
        
    
    'tt = InStrRev(b1, ".")
    'tx = Mid$(b1, 1, tt - 1) + " #- 1st an only BackUp of Rotate Pic.jpg"
   
    'tt = Dir(A1 + b1)
    
    tt = Dir(A1 + "#Rotates\" + B1)
    If tt = "" Then
        On Error Resume Next
            MkDir A1 + "#Rotates"
        On Error GoTo 0
        
        FS.CopyFile A1 + B1, A1 + "#Rotates\" + B1
    
    End If
    
    
    Set F = FS.getfile(Filename)
    tzy = F.datelastmodified
    'tzy = F.datecreated
   
   
   
    
    'FileName = FileDialog(Me, True, "Save As ...", "JPEG Files [*.jpg; *.jpeg]|*.jpg;*.jpeg|All Files [*.*]|*.*", , "*.jpg")
    If Len(Filename) > 0 Then
        Set SaveForm = New fSaveJPG
        SaveForm.SaveImage m_Image, Filename
        Load SaveForm 'This is the Save
        Unload SaveForm
        Set SaveForm = Nothing
        tr = SetFileDateTime(Filename, tzy)
    End If
    
    Set MyPic = Nothing
    Set m_Image = Nothing

    
End Sub
Private Sub mnuExit_Click()
    Unload Me
End Sub
Private Sub ReSizeCords(ADJUSTX, ADJUSTY)
    'Search Honey
    'ADJUSTX = (fMain.Width / Screen.TwipsPerPixelX) - ((RTB1.Width / Screen.TwipsPerPixelX)) - (((Lb2(0).Width + Lb1(0).Width + 300) / Screen.TwipsPerPixelX))
    'ADJUSTY = (fMain.Height / Screen.TwipsPerPixelY) - 55
    
    ADJUSTX = (fmain.Width / Screen.TwipsPerPixelX)
    ADJUSTY = (fmain.Height / Screen.TwipsPerPixelY)
    
    'RTB1.Width = WidthLeft - 200

    'ax = Picture1.ScaleWidth
    'ay = Picture1.ScaleHeight
    'ax = MyPic.Width 'm_Image.Width
    'ay = MyPic.Height 'm_Image.Height
    AX = m_Image.Width
    AY = m_Image.Height
    
    'If ax < 10 Or ay < 10 Then Stop
      
    'ax2$ = Trim$(Round(ax, 0))
    'ay2$ = Trim$(Round(ay, 0))
      
    'ratio = 1 + (1 / 2000) 'Slower = More Passes
    RATIO = 1 + (1 / 200) 'Faster = Less Passes
      
    COMPSIZE = 1
      
    If AX = 0 Or AY = 0 Then Exit Sub
      
      TAG1 = 0
      TAG2 = 0
      Do
'      DoEvents
      TAG3 = 0
      If AX > ADJUSTX Or AY > ADJUSTY Then
      AX = AX / RATIO: AY = AY / RATIO: TAG1 = 1
      COMPSIZE = COMPSIZE / RATIO
      TAG3 = 1
      End If
      
      If AX < ADJUSTX - 0.1 And AY < ADJUSTY - 0.1 Then
      AX = AX * RATIO: AY = AY * RATIO: TAG2 = 1
      COMPSIZE = COMPSIZE * RATIO
      TAG3 = 1
      End If
      
      'DoEvents
      COUNTER = COUNTER + 1
      If COUNTER > 90000 Then Stop
      Loop Until (TAG1 = 1 And TAG2 = 1) Or (TAG3 = 0)
      
      ADJUSTX = AX
      ADJUSTY = AY
      
      PASSES = COUNTER
      Comps$ = Format$(COMPSIZE, "##0.0000")

        'List1.AddItem ax1$ + "x" + ay1$ + "    " + B$


End Sub




Public Sub KeyCodeSub_Intercept()
'You Need to Enter Here with KeyCode2 set to something
If TestUp = True Then Timer2.enabled = False: Exit Sub

If KeyCode2 = 0 Then Exit Sub
If ScanInProgress = True Then Exit Sub

If Me.WindowState > 0 Then Exit Sub


Dim GA1()
ReDim GA1(100)
i = 0
i = i + 1: GA1(i) = 27 = KeyCode2                   'esc
i = i + 1: GA1(i) = 116 = KeyCode2                 'f5
i = i + 1: GA1(i) = 32 = KeyCode2                   'space
i = i + 1: GA1(i) = 40 = KeyCode2                   'down
i = i + 1: GA1(i) = 38 = KeyCode2                   'up
i = i + 1: GA1(i) = 37 = KeyCode2                   'left
i = i + 1: GA1(i) = 39 = KeyCode2                   'right
i = i + 1: GA1(i) = 107 = KeyCode2                 'num plus
i = i + 1: GA1(i) = 13 = KeyCode2                   'enter key
i = i + 1: GA1(i) = 112 = KeyCode2                 'F1
i = i + 1: GA1(i) = 46 = KeyCode2                   'Del
i = i + 1: GA1(i) = 45 = KeyCode2                   'insert 0 in keypad
i = i + 1: GA1(i) = 33 = KeyCode2                   'Page Up -- Speed
i = i + 1: GA1(i) = 34 = KeyCode2                   'Page Down

If KeyCode2 = 188 Then KeyCode2 = 76
If KeyCode2 = 190 Then KeyCode2 = 78
i = i + 1: GA1(i) = 78 = KeyCode2                   'CHAR N NEXT FOLDER
i = i + 1: GA1(i) = 76 = KeyCode2                   'CHAR L NEXT FOLDER

ReDim Preserve GA1(i)

'If KeyCode2 > 0 Then MsgBox Str(KeyCode2) + " Your Keys Sir": KeyCode2 = 0

For R = 1 To i
    TT1 = GA1(i) Xor i
Next

If TT1 = TT2 And TT1 = False Then Exit Sub
TT2 = TT1


'Esc
If KeyCode2 = 27 Then   'ESC
   If fmain.WindowState = 1 Then
        fmain.WindowState = 0
    Else
        fmain.WindowState = 1
    End If
    Debug.Print "ESCAPE PRESSED"
End If

If KeyCode2 = 116 Or KeyCode2 = 32 Then 'F5 or Space
    FileInfoTimer.enabled = Not FileInfoTimer.enabled
    Call PauseOnOffStat
End If

If KeyCode2 = 40 Then 'Keypad Down
    Call KeyLeft
End If
If KeyCode2 = 38 Then 'Keypad Up
    Call KeyRight
End If
If KeyCode2 = 37 Then 'Keypad Left
    Call KeyLeft
End If
If KeyCode2 = 39 Then 'Keypad right
    Call KeyRight
End If

If KeyCode2 = 107 Then
    If Slider1.Value < 30 Then
        Slider1.Value = 30
    Else
        Slider1.Value = 0
    End If
End If

If KeyCode2 = 34 Then
        If Slider1.Value > 0 Then Slider1.Value = Slider1.Value - 1: ' KeyCodeSub_Intercept
End If

If KeyCode2 = 33 Then
        If Slider1.Value <= 100 Then Slider1.Value = Slider1.Value + 1: ' KeyCodeSub_Intercept
End If

If KeyCode2 = 13 Then
    Call ReverseModeSwap
    Call PauseOnOffStat
End If

If KeyCode2 = 78 Then
    'If Lab(Index).Caption = "Next Folder" Then
    'PicFilesOffFolderLast
    CountPic = PicFilesOffFolderNX + 1
    DontCount = True
    Call FileInfoTimer_Timer
End If

If KeyCode2 = 76 Then
    CountPic = PicFilesOffFolderNX - 1
    DontCount = True
    Call FileInfoTimer_Timer
End If



If KeyCode2 = 112 Then
    Msg2 = "Code Mastered By Matthew Lan Matt.Lan@btinternet.com" + vbCrLf
    Msg2 = Msg2 + "--" + vbCrLf
    Msg2 = Msg2 + "-- ** Hot Keys **" + vbCrLf
    Msg2 = Msg2 + "-- 4 Cursors -=- All Go Forward Or Back Throu Images - Got Swap Feature" + vbCrLf
    Msg2 = Msg2 + "-- TIDAL CONTROLS RIGHT PAD OF FOUR BUTTON FLIP FORCE AND REVERSE" + vbCrLf
    
    Msg2 = Msg2 + "-- F1               -=- This Help" + vbCrLf
    Msg2 = Msg2 + "-- F5               -=- Pause / Go" + vbCrLf
    Msg2 = Msg2 + "-- *Space* or *F5*  -=- Pause / Go" + vbCrLf
    Msg2 = Msg2 + "-- *PageUp* *Down* -=- Speed Control" + vbCrLf
    Msg2 = Msg2 + "-- Num Pad '-' Minus Key -=- Fast / Slow Toggle" + vbCrLf
    Msg2 = Msg2 + "-- Enter Key               -=- Reverse / Forward Swap" + vbCrLf
    Msg2 = Msg2 + "-- Num Pad Zero       -=- Last Image an Stop an Put in Reverse" + vbCrLf
    'cound make this do other way round whne in reverse
    Msg2 = Msg2 + "### Proggy Stops after Hour -- Stops if Not in focus Nice CPU Loads of Things"
    Msg2 = Msg2 + "-- " + vbCrLf
    'Tx = "(¯`'•.¸ -‹(•¿•)›- ¸.•'´¯) --  -- <º)))><¸..·´`·..¸><(((º> --  -- -:¦:-•:*'''''*:•-:¦:-" + vbCrLf
    Tx = "(¯`'•.¸ -‹(•¿•)›- ¸.•'´¯) --  -- <º)))><¸..·´`·..¸><(((º> --" + vbCrLf
    'Msg2 = Msg2 + Tx
    Msg2 = Msg2 + "-- " + vbCrLf
    Msg2 = Msg2 + "-- This is a Multi Intelligent Image Program Master Crafted Creation" + vbCrLf
    Msg2 = Msg2 + "-- Created By The One and Only ** Champ Hardcore Stomper Coder **" + vbCrLf
    Msg2 = Msg2 + "-- (M)onster (C)ode (E)xpert -- E=MC^2 -- Gota a Lota Top" + vbCrLf
    Msg2 = Msg2 + "-- " + vbCrLf
    
    Msg2 = Msg2 + "Version - " + Trim(str(App.Major)) + "." + Trim(str(App.Minor)) + "." + Trim(str(App.Revision)) + vbCrLf
    Set F = FS.getfile(App.Path + "\" + App.EXEName + ".exe")
    'Msg2 = Msg2 + "-- Last Compiled " + Format$(F.datecreated, "MMM DD-MMM-YYYY HH:MM:SSa/p") + vbCrLf
    Msg2 = Msg2 + "Product Start Date -- Mar 02-Mar-2009 07:40:33a" + vbCrLf
    Msg2 = Msg2 + "Last Compiled -- " + Format$(F.datelastmodified, "MMM DD-MMM-YYYY HH:MM:SSa/p") + vbCrLf

    'Use this
    InPutDate = DateValue("02-03-2009")
    TestDate = DateValue("06-01-2009")
    TestDate = Now
    Call FindTimeInfo
    Msg2 = Msg2 + "Age of Program -- " + ResultYearDate + " Years" + str(Month7) + "M" + str(DayCountMonth) + "D" + str(WeeksSinceInput) + " Weeks" + vbCrLf
    '-- Art of Program
    Msg2 = Msg2 + "-- " + vbCrLf
    Msg2 = Msg2 + Tx
    
    msgbox Msg2, vbInformation + vbOKOnly + vbMsgBoxSetForeground, "Matt Lan Hammers Out Another HardCore Code..."
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------- Release
End If

If KeyCode2 = 46 And Mnu_DelKey.Checked = True Then
    Call Mnu_Del_Click
End If

'Put to Last pic an in-rev an Stopped
If KeyCode2 = 45 Then 'insert 0 in keypad
    If CountPic <> ScanPath.ListView1.ListItems.Count Then
        If Lab(ForBACKCTRL).Caption = "<º)))> Reverse Mode <(((º>" Then
            Call ReverseMode
            Call PauseOnOffStat 'Update Labels
        End If
        CountPic = ScanPath.ListView1.ListItems.Count
        Call Timer1_Timer 'show new pic 'fileinfo may do this
        FileInfoTimer.enabled = False
        Call PauseOnOffStat 'Update Labels
    Else
        FileInfoTimer.enabled = True
        Call PauseOnOffStat 'Update Labels
        'Call KeyCodeSub_Intercept 'show new pic ???
    End If
End If

KeyCode2 = 0


End Sub


Sub PauseOnOffStat()
If ScanInProgress = True Then Exit Sub


If FileInfoTimer.enabled <> OTimer1 Then
    If FileInfoTimer.enabled = True Then
        Lab(PauseCTRL) = "Hitt to Pause":        FileInfoTimer.enabled = True
    Else
        Lab(PauseCTRL) = "Hitt to Start - Go": FileInfoTimer.enabled = False
    End If
    OTimer1 = FileInfoTimer.enabled
End If

If LBB3T = "" Then LBB3T = LBB3
If Lab(PauseCTRL) = "Hitt to Pause" Then
    If Lab(ForBACKCTRL).Caption = "<º)))> Forward Mode <(((º>" Then
        LBB3T = "Reverse"
        Reverse = True
    Else
        LBB3T = "The Force"
        Reverse = False
    End If
End If
LBB3T1 = ""
If Lab(PauseCTRL) <> "Hitt to Pause" Then
    LBB3T1 = "Paused - " + LBB3T
'    msgbox "Pause"
End If
If LBB3T1 = "" Then LBB3T1 = LBB3T

If OLBB3 <> LBB3T1 Then
LBB3 = LBB3T1
OLBB3 = LBB3T1
End If

End Sub




Public Sub ShellAndWait(PathName, Optional WindowStyle As _
   VbAppWinStyle = vbMinimizedFocus, Optional bDoEvents As _
   Boolean = False)

    Dim dwProcessId As Long
    Dim hProcess As Long
    
    dwProcessId = Shell(PathName, WindowStyle)
    
    If dwProcessId = 0 Then
        Exit Sub
    End If
    
    hProcess = OpenProcess(SYNCHRONIZE, False, dwProcessId)
    
    If hProcess = 0 Then
        Exit Sub
    End If
    
    If bDoEvents Then
        Do While WaitForSingleObject(hProcess, 1000) = WAIT_TIMEOUT
            DoEvents
        Loop
    Else
        WaitForSingleObject hProcess, INFINITE
    End If
    
    CloseHandle hProcess

'Sample Usage:
  
'     Dim nStart As Date
'    nStart = Now

'    ShellAndWait "notepad", vbNormalFocus, True
    
'    messageBox "You spent " & DateDiff("s", nStart, Now) & _
'       " second(s) in notepad.", vbCritical
    
End Sub

Private Sub Timer3_Timer()
    Timer3.enabled = False 'This is for The One on Percent counter
End Sub

Public Sub RunIn()

LBB1 = "" '"The One"
XD = 0
If AgWa > 0 Then
    XD = fmain.Lb2(AgWa - 1)
    If Val(fmain.Lb2(AgWa - 1)) = 0 Then
        Frame1.Height = PBar3.Top
    Else
        Frame1.Height = PBar3.Top + PBar3.Height - 20
    End If
End If
If XD = "." Then XD = "0"

Picture1.ZOrder 1 'Send to Back
Frame1.Visible = True
Frame3.Visible = False
Frame5.Visible = False
Frame1.ZOrder 0 'Bring to Front
Frame1.Refresh
'Timer3.Enabled = False
'Timer3.Interval = 0
'Timer3.Interval = 400
'Timer3.Enabled = True
End Sub

Public Sub RunOut()
    
    Frame1.Visible = False
    Frame3.Visible = True
    Frame5.Visible = True

End Sub


Public Sub RunCount()

If Val(ScanPath.lblCount1) = 0 Then Exit Sub

'XD is the stored value for its remembered file count this folder
'Rememebered from last time so can show percentage loaded or eles just a number count
If Val(ScanPath.lblCount1) > XD Then SingleNum = True Else SingleNum = False 'XD = Scanpath.lblCount
If Val(XD) = 0 Then SingleNum = True
If Val(fmain.Lb2(AgWa - 1)) > 0 Then
    
    'already modded ten

    If XD = "0" Then XD = "1"
    
    'PBar3.Width = Frame1.Width

    If PBar3.Value <> Int(((ScanPath.lblCount1 / XD) * 100)) And SingleNum = False Then
        'If ScanPath.lblCount1 / XD = 1 Then Stop
        ts = Format$((ScanPath.lblCount1 / XD) * 100, "#0") + "% Loaded"
        'xd = how many batches of folders to load for on set run count
        'If (Scanpath.lblCount1 / XD) > 0.04 Then Stop '= 4%
        'If Timer3.Enabled = True Then Stop
        'If Timer3.Enabled = True Or ((Scanpath.lblCount1 / XD) < 0.09) Then ts = "The One"
        LBB1.Caption = ts
        PBar3.Value = Int(((ScanPath.lblCount1 / XD) * 100))
'        LBB1.Refresh
'        Frame1.Refresh
    End If

    If SingleNum = True Then
        LBB1.Caption = Trim(str(ScanPath.lblCount1)) + " Loaded"
        PBar3.Value = 100
'        LBB1.Refresh
'        Frame1.Refresh
    End If
Exit Sub
End If
LBB1.Caption = Trim(str(ScanPath.lblCount1)) + " Loaded"

End Sub
 

Function AlwaysOnTop(ByVal hwnd As Long)  'Makes a form always on top
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Function
Function NotAlwaysOnTop(ByVal hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags
End Function


Private Sub ColorCycle()
WTrue = WTrue + TW1
i = 1
If WTrue > 255 Then TW1 = i * -1: WTrue = 255
If WTrue < 0 Then TW1 = i: WTrue = 0

HWTrue = HWTrue + TW2
i = i * 3
If HWTrue > 255 Then TW2 = i * -1: HWTrue = 255
If HWTrue < 0 Then TW2 = i: HWTrue = 0

KWTrue = KWTrue + TW3
i = i * 4
If KWTrue > 255 Then TW3 = i * -1: KWTrue = 255
If KWTrue < 0 Then TW3 = i: KWTrue = 0

Lbl1.BackColor = RGB(KWTrue, HWTrue, WTrue)    ' Set drawing color.
Lbl1.ForeColor = RGB(255 - KWTrue, 255 - HWTrue, 255 - WTrue)

If Reverse = False Then
    Lbl2.BackColor = RGB(KWTrue, HWTrue, WTrue)    ' Set drawing color.
    Lbl2.ForeColor = RGB(255 - KWTrue, 255 - HWTrue, 255 - WTrue)
Else
    Lbl2.BackColor = RGB(KWTrue Xor 255, HWTrue Xor 255, WTrue Xor 255)     ' Set drawing color.
    Lbl2.ForeColor = RGB((255 - KWTrue) Xor 255, (255 - HWTrue) Xor 255, (255 - WTrue) Xor 255)
End If

   
End Sub

Private Sub TimerColorCycle_Timer()
    
    Exit Sub
    
    x = Mnu_Pause_When_Not_In_Focus.Checked
    If GetForegroundWindow <> Me.hwnd Then Exit Sub
    
    If TestUp = True Then TimerColorCycle.enabled = False
    Call ColorCycle
    If Picture1.Visible = True Then Picture1.SetFocus

    
End Sub



Sub FindTimeInfo()

'InPutDate = DateValue("01-12-2008")
'TestDate = DateValue("31-05-2009")

Year5 = 0
For R5 = Year(InPutDate) + 1 To Year(TestDate) + 1
    If DateSerial(R5, Month(InPutDate), Day(InPutDate)) < Now Then Year5 = Year5 + 1
Next
For R5 = 1 To -2 Step -1
    XX = DateSerial(Year(TestDate), Month(TestDate) + R5, Day(InPutDate))
    MonthStart = XX
    If XX <= TestDate Then Exit For
Next
Month5 = 0
Month7 = 0
For R5 = 1 To 10000
    XX = DateSerial(Year(InPutDate), Month(InPutDate) + R5, Day(InPutDate))
    If Year(XX) <> oxx Then Month7 = 0
    oxx = Year(XX)
    If XX <= TestDate + 1 Then Month5 = Month5 + 1: Month7 = Month7 + 1
    If XX >= TestDate + 1 Then Exit For
Next

For R5 = Year(TestDate) To 1 Step -1
    If DateSerial(R5, Month(InPutDate), Day(InPutDate)) < Now Then
        DayCount1 = DateDiff("d", DateSerial(R5, Month(InPutDate), Day(InPutDate)), TestDate)
        DayCountMonth = DateDiff("d", MonthStart, TestDate)
        WholeYear1 = DateDiff("d", DateSerial(R5, Month(InPutDate), Day(InPutDate)), DateSerial(R5 + 1, Month(InPutDate), Day(InPutDate)))
        'Month5 = DateDiff("m", DateSerial(r5, Month(InPutDate), Day(InPutDate)), TestDate)
        WeeksSinceYear = DateDiff("w", DateSerial(R5, Month(InPutDate), Day(InPutDate)), TestDate) - 1
        WeeksSinceInput = DateDiff("w", InPutDate, TestDate) - 1
        ResultYearDate = Format$(Year5 + DayCount1 / WholeYear1, ".000")
        Exit For
    End If
Next

End Sub




'Private Sub Form_Load()
'Dim lPic As Picture
'Me.Picture1.AutoRedraw = True
'Set lPic = LoadPicture("C:\YourPicture.jpg") 'Use the correct path and filename here
'ResizePicture Me.Picture1, lPic
'End Sub

Public Sub ResizePicture(pBox As PictureBox, pPic As Picture)
Dim lWidth As Single, lHeight As Single
Dim lnewWidth As Single, lnewHeight As Single
'Clear the Picture in the PictureBox
'pBox = LoadPicture()
pBox.Picture = Nothing
'Clear the Image in the Picturebox
pBox.Cls
'Get the size of the Image, but in the same Scale than the scale used by the PictureBox
lWidth = pBox.ScaleX(pPic.Width, vbHiMetric, pBox.ScaleMode)
lHeight = pBox.ScaleY(pPic.Height, vbHiMetric, pBox.ScaleMode)

OX = Round(lWidth / Screen.TwipsPerPixelX, 0)
OY = Round(lHeight / Screen.TwipsPerPixelY, 0)
OX1 = lWidth / Screen.TwipsPerPixelX
OY1 = lHeight / Screen.TwipsPerPixelY

'Old Rems
'If image Width > pictureBox Width, resize Width
If lWidth > pBox.ScaleWidth Then
    lnewWidth = pBox.ScaleWidth 'new Width = PB width
    lHeight = lHeight * (lnewWidth / lWidth) 'Risize Height keeping proportions
Else
    'if you want true size on small Pics then This
        'lnewWidth = lWidth 'If not, keep the original Width value

    'OR if you want expand small and shrink big then this
    lnewWidth = pBox.ScaleWidth 'new Width = PB width
    lHeight = lHeight / (lWidth / lnewWidth) 'Risize Height keeping proportions

End If
    'Old Rems
    'If the image Height > The pictureBox Height, resize Height

If lHeight > pBox.ScaleHeight Then
    lnewHeight = pBox.ScaleHeight 'new Height = PB Height
    lnewWidth = lnewWidth * (lnewHeight / lHeight) 'Risize Width keeping proportions
Else
    'if you want true size on small then This only this line
        lnewHeight = lHeight 'If not, use the same value

    'OR if you want expand small and shrink big then this
    lnewWidth = lnewWidth / (lHeight / lnewHeight) 'Risize Width keeping proportions
End If

'add resized and centered to Picturebox
pBox.PaintPicture pPic, (pBox.ScaleWidth - lnewWidth) / 2, (pBox.ScaleHeight - lnewHeight) / 2, lnewWidth, lnewHeight

'Picture2.PaintPicture pPic, (pBox.ScaleWidth - lnewWidth) / 2, (pBox.ScaleHeight - lnewHeight) / 2, lnewWidth, lnewHeight

'Picture2.Width = Picture1.Height
'Picture2.Height = Picture1.Width
'Picture2.Left = Picture1.Left
'Picture2.Top = Picture1.Top

'Update the Picture with the new image if you need it
Set pBox.Picture = pBox.Image
'Set Picture2.Picture = pBox.Image

'pbox.Picture.

LBX(3) = Trim(str(OX)) + " x " + Trim(str(OY))


End Sub


'Here i was going to resize the pic box whatout having another pic but the display will show in conversion
Public Sub ResizePictureX(pBox As PictureBox)
Dim lWidth As Single, lHeight As Single
Dim lnewWidth As Single, lnewHeight As Single
'Clear the Picture in the PictureBox
'pBox = LoadPicture()
'pBox.Picture = Nothing
'Clear the Image in the Picturebox
'pBox.Cls
'Get the size of the Image, but in the same Scale than the scale used by the PictureBox
lWidth = pBox.ScaleX(pBox.Width, vbHiMetric, pBox.ScaleMode)
lHeight = pBox.ScaleY(pBox.Height, vbHiMetric, pBox.ScaleMode)
OX = Round(lWidth / Screen.TwipsPerPixelX, 0)
OY = Round(lHeight / Screen.TwipsPerPixelY, 0)
OX1 = lWidth / Screen.TwipsPerPixelX
OY1 = lHeight / Screen.TwipsPerPixelY

'Old Rems
'If image Width > pictureBox Width, resize Width
If lWidth > pBox.ScaleWidth Then
    lnewWidth = pBox.ScaleWidth 'new Width = PB width
    lHeight = lHeight * (lnewWidth / lWidth) 'Risize Height keeping proportions
Else
    'if you want true size on small Pics then This
        'lnewWidth = lWidth 'If not, keep the original Width value

    'OR if you want expand small and shrink big then this
    lnewWidth = pBox.ScaleWidth 'new Width = PB width
    lHeight = lHeight / (lWidth / lnewWidth) 'Risize Height keeping proportions

End If
    'Old Rems
    'If the image Height > The pictureBox Height, resize Height

If lHeight > pBox.ScaleHeight Then
    lnewHeight = pBox.ScaleHeight 'new Height = PB Height
    lnewWidth = lnewWidth * (lnewHeight / lHeight) 'Risize Width keeping proportions
Else
    'if you want true size on small then This only this line
        lnewHeight = lHeight 'If not, use the same value

    'OR if you want expand small and shrink big then this
    lnewWidth = lnewWidth / (lHeight / lnewHeight) 'Risize Width keeping proportions
End If

'add resized and centered to Picturebox
pBox.PaintPicture pPic, (pBox.ScaleWidth - lnewWidth) / 2, (pBox.ScaleHeight - lnewHeight) / 2, lnewWidth, lnewHeight

'Picture2.PaintPicture pPic, (pBox.ScaleWidth - lnewWidth) / 2, (pBox.ScaleHeight - lnewHeight) / 2, lnewWidth, lnewHeight

'Picture2.Width = Picture1.Height
'Picture2.Height = Picture1.Width
'Picture2.Left = Picture1.Left
'Picture2.Top = Picture1.Top

'Update the Picture with the new image if you need it
Set pBox.Picture = pBox.Image
'Set Picture2.Picture = pBox.Image

'pbox.Picture.

LBX(3) = Trim(str(OX)) + " x " + Trim(str(OY))


End Sub


Public Sub ResizePicture2(pBox As PictureBox, pPic As Picture)
Dim lWidth As Single, lHeight As Single
Dim lnewWidth As Single, lnewHeight As Single
'Clear the Picture in the PictureBox
'pBox = LoadPicture()
'pBox.Picture = Nothing
'Clear the Image in the Picturebox
pBox.Cls
'Picture2 = LoadPicture()
'Picture2.Picture = Nothing
'Picture2.Cls
'Get the size of the Image, but in the same Scale than the scale used by the PictureBox
lWidth = pBox.ScaleY(pPic.Height, vbHiMetric, pBox.ScaleMode)
lHeight = pBox.ScaleX(pPic.Width, vbHiMetric, pBox.ScaleMode)
OX = Round(lWidth / Screen.TwipsPerPixelX, 0)
OY = Round(lHeight / Screen.TwipsPerPixelY, 0)
OX1 = lWidth / Screen.TwipsPerPixelX
OY1 = lHeight / Screen.TwipsPerPixelY

'Old Rems
'If image Width > pictureBox Width, resize Width
If lWidth > pBox.ScaleWidth Then
    lnewWidth = pBox.ScaleWidth 'new Width = PB width
    lHeight = lHeight * (lnewWidth / lWidth) 'Risize Height keeping proportions
Else
    'if you want true size on small Pics then This
        'lnewWidth = lWidth 'If not, keep the original Width value

    'OR if you want expand small and shrink big then this
    lnewWidth = pBox.ScaleWidth 'new Width = PB width
    lHeight = lHeight / (lWidth / lnewWidth) 'Risize Height keeping proportions

End If
    'Old Rems
    'If the image Height > The pictureBox Height, resize Height

If lHeight > pBox.ScaleHeight Then
    lnewHeight = pBox.ScaleHeight 'new Height = PB Height
    lnewWidth = lnewWidth * (lnewHeight / lHeight) 'Risize Width keeping proportions
Else
    'if you want true size on small then This only this line
        lnewHeight = lHeight 'If not, use the same value

    'OR if you want expand small and shrink big then this
    lnewWidth = lnewWidth / (lHeight / lnewHeight) 'Risize Width keeping proportions
End If

'add resized and centered to Picturebox
pBox.PaintPicture pPic, (pBox.ScaleWidth - lnewWidth) / 2, (pBox.ScaleHeight - lnewHeight) / 2, lnewWidth, lnewHeight

'Picture2.PaintPicture pPic, (pBox.ScaleWidth - lnewWidth) / 2, (pBox.ScaleHeight - lnewHeight) / 2, lnewWidth, lnewHeight

'Picture2.Width = Picture1.Height
'Picture2.Height = Picture1.Width
'Picture2.Left = Picture1.Left
'Picture2.Top = Picture1.Top

'Update the Picture with the new image if you need it
Set pBox.Picture = pBox.Image
'Set Picture2.Picture = pBox.Image

'pbox.Picture.

LBX(3) = Trim(str(OX)) + " x " + Trim(str(OY))


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

'Sub zzLoad_Checks()
'Put This In Declarations
'Dim OCheckQ
'Very Nice Code Will Save all your Check Boxes an Menu Checks and Values you can filter
'some out -- works best as I Know

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

'-------------------------------
Sub zzLoad_Checks()

zzCheckTimer.enabled = False

Dim Th()
ReDim Th(Me.Controls.Count * 3)

On Error Resume Next
MkDir App.Path + "\0TextData"
i = 0
FR5 = FreeFile
Open App.Path + "\0TextData\ChkSettings.txt" For Input As #FR5
Do
    Line Input #FR5, vv$
    Th(i) = vv$
    i = i + 1
    Line Input #FR5, vv$
    Th(i) = Val(vv$)
    i = i + 1
    Line Input #FR5, vv$
    Th(i) = Val(vv$)
    i = i + 1
Loop Until EOF(FR5)
Close #FR5
    
tit = i
For Each Control In Me.Controls
    XX = 1
    
    ppi = LCase(Control.Name)
    If InStr(ppi, LCase("Combo")) Then XX = 0
    If InStr(ppi, LCase("Check")) Then XX = 0
    If InStr(ppi, LCase("Mnu")) Then XX = 0
    If InStr(ppi, LCase("Menu")) Then XX = 0
    
    
    XXD = -1
    For R = 0 To tit
        If Control.Name = Th(R) Then
            XXD = R: Exit For
        End If
    Next
    
    If XXD > 0 And XX = 0 Then
        On Error Resume Next
        If Th(XXD + 1) <> 0 Then Control.Value = Th(XXD + 1)
        If Th(XXD + 2) <> 0 Then Control.Checked = Th(XXD + 2)
        'Dont Let People Play Around If you Want to Hard Code In Designer To Enable Chk This
        'This Lets Happen Eg If Th(xxd + 2) <> 0
        
        Add1 = 0
        Add1 = Control.Value
        Bdd1 = 0
        Bdd1 = Control.Checked
        OCheckQ = OCheckQ Xor (Add1 + Bdd1)
        On Error GoTo 0
    End If
Next
zzCheckTimer.enabled = True
End Sub

'-------------------------------
Sub zzSave_Checks()
Dim Add1, Bdd1 As Integer
On Error Resume Next
FR5 = FreeFile
Open App.Path + "\0TextData\ChkSettings.txt" For Output As #FR5

For Each Control In Me.Controls
    Err.Clear
        Add1 = 0
        Add1 = Control.Value
        Add2 = Err.Number
    Err.Clear
        Bdd1 = 0
        Bdd1 = Control.Checked
        Bdd2 = Err.Number
    
    If Add2 = 0 Or Bdd2 = 0 Then
        Print #FR5, Control.Name
        Print #FR5, Add1
        Print #FR5, Bdd1
    End If
Next
Close #FR5

End Sub



'-------------------------------
Private Sub zzCheckTimer_Timer()

On Error Resume Next
For Each Control In Me.Controls
    Err.Clear
        Add1 = 0
        Add1 = Control.Value
    Err.Clear
        Bdd1 = 0
        Bdd1 = Control.Checked
        checkq = checkq Xor (Add1 + Bdd1)
Next

If checkq = OCheckQ Then Exit Sub
OCheckQ = checkq
Call zzSave_Checks

End Sub


