VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   Caption         =   "EXPLORER LOADER"
   ClientHeight    =   5928
   ClientLeft      =   192
   ClientTop       =   2340
   ClientWidth     =   12312
   Icon            =   "Shell Loader.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5928
   ScaleWidth      =   12312
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer TIMER_FORMSTART 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6180
      Top             =   900
   End
   Begin VB.Timer TIMER_TO_RESIZE 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7176
      Top             =   1320
   End
   Begin VB.Timer Timer_FORM_RESIZE 
      Interval        =   10
      Left            =   7176
      Top             =   936
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9552
      Top             =   3660
   End
   Begin VB.FileListBox File1 
      Height          =   264
      Left            =   8472
      TabIndex        =   265
      Top             =   1584
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.DirListBox Dir1 
      Height          =   504
      Left            =   8472
      TabIndex        =   264
      Top             =   852
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Timer Timer_KEY_CODE 
      Interval        =   1
      Left            =   8844
      Top             =   3660
   End
   Begin VB.Timer Timer_SUB_CODE_LOAD 
      Interval        =   1
      Left            =   8472
      Top             =   3660
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9204
      Top             =   3660
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   5676
      TabIndex        =   166
      Text            =   "Combo1"
      Top             =   84
      Width           =   2652
   End
   Begin VB.ListBox List1 
      Height          =   240
      Left            =   8472
      TabIndex        =   19
      Top             =   2544
      Visible         =   0   'False
      Width           =   2376
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Load Log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8472
      TabIndex        =   17
      Top             =   3132
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   488
      Left            =   396
      TabIndex        =   500
      Top             =   4236
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   487
      Left            =   396
      TabIndex        =   499
      Top             =   4032
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   486
      Left            =   396
      TabIndex        =   498
      Top             =   3816
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   485
      Left            =   396
      TabIndex        =   497
      Top             =   3612
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   484
      Left            =   396
      TabIndex        =   496
      Top             =   3492
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   483
      Left            =   396
      TabIndex        =   495
      Top             =   3696
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   482
      Left            =   396
      TabIndex        =   494
      Top             =   3912
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   481
      Left            =   396
      TabIndex        =   493
      Top             =   4116
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   480
      Left            =   396
      TabIndex        =   492
      Top             =   3492
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   479
      Left            =   396
      TabIndex        =   491
      Top             =   3708
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   478
      Left            =   396
      TabIndex        =   490
      Top             =   3912
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   477
      Left            =   396
      TabIndex        =   489
      Top             =   4128
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   476
      Left            =   396
      TabIndex        =   488
      Top             =   4332
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   475
      Left            =   396
      TabIndex        =   487
      Top             =   4548
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   474
      Left            =   396
      TabIndex        =   486
      Top             =   3828
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   473
      Left            =   396
      TabIndex        =   485
      Top             =   4332
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   472
      Left            =   396
      TabIndex        =   484
      Top             =   4548
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   471
      Left            =   396
      TabIndex        =   483
      Top             =   4332
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   470
      Left            =   396
      TabIndex        =   482
      Top             =   4128
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   469
      Left            =   396
      TabIndex        =   481
      Top             =   3912
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   468
      Left            =   396
      TabIndex        =   480
      Top             =   3708
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   467
      Left            =   396
      TabIndex        =   479
      Top             =   3492
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   466
      Left            =   396
      TabIndex        =   478
      Top             =   3588
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   465
      Left            =   396
      TabIndex        =   477
      Top             =   3792
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   464
      Left            =   396
      TabIndex        =   476
      Top             =   4008
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   463
      Left            =   396
      TabIndex        =   475
      Top             =   4212
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   462
      Left            =   396
      TabIndex        =   474
      Top             =   4428
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   461
      Left            =   396
      TabIndex        =   473
      Top             =   4632
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   460
      Left            =   396
      TabIndex        =   472
      Top             =   4632
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   459
      Left            =   396
      TabIndex        =   471
      Top             =   4512
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   458
      Left            =   396
      TabIndex        =   470
      Top             =   4308
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   457
      Left            =   396
      TabIndex        =   469
      Top             =   4092
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   456
      Left            =   396
      TabIndex        =   468
      Top             =   3888
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   455
      Left            =   396
      TabIndex        =   467
      Top             =   3672
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   454
      Left            =   396
      TabIndex        =   466
      Top             =   3588
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   453
      Left            =   396
      TabIndex        =   465
      Top             =   3792
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   452
      Left            =   396
      TabIndex        =   464
      Top             =   4008
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   451
      Left            =   396
      TabIndex        =   463
      Top             =   4212
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   450
      Left            =   396
      TabIndex        =   462
      Top             =   4428
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   449
      Left            =   396
      TabIndex        =   461
      Top             =   4632
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   448
      Left            =   396
      TabIndex        =   460
      Top             =   4788
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   447
      Left            =   396
      TabIndex        =   459
      Top             =   3420
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   446
      Left            =   396
      TabIndex        =   458
      Top             =   4632
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   445
      Left            =   396
      TabIndex        =   457
      Top             =   4428
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   444
      Left            =   396
      TabIndex        =   456
      Top             =   4212
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   443
      Left            =   396
      TabIndex        =   455
      Top             =   4008
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   442
      Left            =   396
      TabIndex        =   454
      Top             =   3792
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   441
      Left            =   396
      TabIndex        =   453
      Top             =   3588
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   440
      Left            =   396
      TabIndex        =   452
      Top             =   4176
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   439
      Left            =   396
      TabIndex        =   451
      Top             =   3600
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   438
      Left            =   396
      TabIndex        =   450
      Top             =   3804
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   437
      Left            =   396
      TabIndex        =   449
      Top             =   4020
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   436
      Left            =   396
      TabIndex        =   448
      Top             =   4224
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   435
      Left            =   396
      TabIndex        =   447
      Top             =   4440
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   434
      Left            =   396
      TabIndex        =   446
      Top             =   4440
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   433
      Left            =   396
      TabIndex        =   445
      Top             =   4224
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   432
      Left            =   396
      TabIndex        =   444
      Top             =   4020
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   431
      Left            =   396
      TabIndex        =   443
      Top             =   3804
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   430
      Left            =   396
      TabIndex        =   442
      Top             =   3600
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   429
      Left            =   396
      TabIndex        =   441
      Top             =   3480
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   428
      Left            =   396
      TabIndex        =   440
      Top             =   3684
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   427
      Left            =   396
      TabIndex        =   439
      Top             =   3900
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   426
      Left            =   396
      TabIndex        =   438
      Top             =   4104
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   425
      Left            =   396
      TabIndex        =   437
      Top             =   4320
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   424
      Left            =   396
      TabIndex        =   436
      Top             =   4968
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   423
      Left            =   396
      TabIndex        =   435
      Top             =   4884
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   422
      Left            =   396
      TabIndex        =   434
      Top             =   4224
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   421
      Left            =   396
      TabIndex        =   433
      Top             =   4020
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   420
      Left            =   396
      TabIndex        =   432
      Top             =   3804
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   419
      Left            =   396
      TabIndex        =   431
      Top             =   3600
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   418
      Left            =   396
      TabIndex        =   430
      Top             =   3504
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   417
      Left            =   396
      TabIndex        =   429
      Top             =   3720
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   416
      Left            =   396
      TabIndex        =   428
      Top             =   3924
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   415
      Left            =   396
      TabIndex        =   427
      Top             =   4140
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   414
      Left            =   396
      TabIndex        =   426
      Top             =   4344
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   413
      Left            =   396
      TabIndex        =   425
      Top             =   5004
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   412
      Left            =   396
      TabIndex        =   424
      Top             =   5004
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   411
      Left            =   396
      TabIndex        =   423
      Top             =   4344
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   410
      Left            =   396
      TabIndex        =   422
      Top             =   4140
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   409
      Left            =   396
      TabIndex        =   421
      Top             =   3924
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   408
      Left            =   396
      TabIndex        =   420
      Top             =   3720
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   407
      Left            =   396
      TabIndex        =   419
      Top             =   3504
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   406
      Left            =   396
      TabIndex        =   418
      Top             =   3600
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   405
      Left            =   396
      TabIndex        =   417
      Top             =   3804
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   404
      Left            =   396
      TabIndex        =   416
      Top             =   4020
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   403
      Left            =   396
      TabIndex        =   415
      Top             =   4224
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   402
      Left            =   396
      TabIndex        =   414
      Top             =   4884
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   401
      Left            =   396
      TabIndex        =   413
      Top             =   5004
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   400
      Left            =   396
      TabIndex        =   412
      Top             =   4344
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   399
      Left            =   396
      TabIndex        =   411
      Top             =   4140
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   398
      Left            =   396
      TabIndex        =   410
      Top             =   3924
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   397
      Left            =   396
      TabIndex        =   409
      Top             =   3720
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   396
      Left            =   396
      TabIndex        =   408
      Top             =   3504
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   395
      Left            =   396
      TabIndex        =   407
      Top             =   4860
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   394
      Left            =   396
      TabIndex        =   406
      Top             =   4212
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   393
      Left            =   396
      TabIndex        =   405
      Top             =   4008
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   392
      Left            =   396
      TabIndex        =   404
      Top             =   3792
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   391
      Left            =   396
      TabIndex        =   403
      Top             =   3588
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   390
      Left            =   396
      TabIndex        =   402
      Top             =   3564
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   389
      Left            =   396
      TabIndex        =   401
      Top             =   3780
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   388
      Left            =   396
      TabIndex        =   400
      Top             =   3984
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   387
      Left            =   396
      TabIndex        =   399
      Top             =   4200
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   386
      Left            =   396
      TabIndex        =   398
      Top             =   4404
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   385
      Left            =   4104
      TabIndex        =   397
      Top             =   1800
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   384
      Left            =   4104
      TabIndex        =   396
      Top             =   1596
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   383
      Left            =   4104
      TabIndex        =   395
      Top             =   1380
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   382
      Left            =   4104
      TabIndex        =   394
      Top             =   1176
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   381
      Left            =   4104
      TabIndex        =   393
      Top             =   960
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   380
      Left            =   4104
      TabIndex        =   392
      Top             =   984
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   379
      Left            =   4104
      TabIndex        =   391
      Top             =   1188
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   378
      Left            =   4104
      TabIndex        =   390
      Top             =   1404
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   377
      Left            =   4104
      TabIndex        =   389
      Top             =   1608
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   376
      Left            =   4104
      TabIndex        =   388
      Top             =   2256
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   375
      Left            =   4104
      TabIndex        =   387
      Top             =   900
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   374
      Left            =   4104
      TabIndex        =   386
      Top             =   1116
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   373
      Left            =   4104
      TabIndex        =   385
      Top             =   1320
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   372
      Left            =   4104
      TabIndex        =   384
      Top             =   1536
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   371
      Left            =   4104
      TabIndex        =   383
      Top             =   1740
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   370
      Left            =   4104
      TabIndex        =   382
      Top             =   2400
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   369
      Left            =   4104
      TabIndex        =   381
      Top             =   2280
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   368
      Left            =   4104
      TabIndex        =   380
      Top             =   1620
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   367
      Left            =   4104
      TabIndex        =   379
      Top             =   1416
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   366
      Left            =   4104
      TabIndex        =   378
      Top             =   1200
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   365
      Left            =   4104
      TabIndex        =   377
      Top             =   996
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   364
      Left            =   4104
      TabIndex        =   376
      Top             =   900
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   363
      Left            =   4104
      TabIndex        =   375
      Top             =   1116
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   362
      Left            =   4104
      TabIndex        =   374
      Top             =   1320
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   361
      Left            =   4104
      TabIndex        =   373
      Top             =   1536
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   360
      Left            =   4104
      TabIndex        =   372
      Top             =   1740
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   359
      Left            =   4104
      TabIndex        =   371
      Top             =   2400
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   358
      Left            =   4104
      TabIndex        =   370
      Top             =   2400
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   357
      Left            =   4104
      TabIndex        =   369
      Top             =   1740
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   356
      Left            =   4104
      TabIndex        =   368
      Top             =   1536
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   355
      Left            =   4104
      TabIndex        =   367
      Top             =   1320
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   354
      Left            =   4104
      TabIndex        =   366
      Top             =   1116
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   353
      Left            =   4104
      TabIndex        =   365
      Top             =   900
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   352
      Left            =   4104
      TabIndex        =   364
      Top             =   996
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   351
      Left            =   4104
      TabIndex        =   363
      Top             =   1200
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   350
      Left            =   4104
      TabIndex        =   362
      Top             =   1416
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   349
      Left            =   4104
      TabIndex        =   361
      Top             =   1620
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   348
      Left            =   4104
      TabIndex        =   360
      Top             =   2280
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   347
      Left            =   4104
      TabIndex        =   359
      Top             =   2364
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   346
      Left            =   4104
      TabIndex        =   358
      Top             =   1716
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   345
      Left            =   4104
      TabIndex        =   357
      Top             =   1500
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   344
      Left            =   4104
      TabIndex        =   356
      Top             =   1296
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   343
      Left            =   4104
      TabIndex        =   355
      Top             =   1080
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   342
      Left            =   4104
      TabIndex        =   354
      Top             =   876
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   341
      Left            =   4104
      TabIndex        =   353
      Top             =   996
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   340
      Left            =   4104
      TabIndex        =   352
      Top             =   1200
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   339
      Left            =   4104
      TabIndex        =   351
      Top             =   1416
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   338
      Left            =   4104
      TabIndex        =   350
      Top             =   1620
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   337
      Left            =   4104
      TabIndex        =   349
      Top             =   1836
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   336
      Left            =   4104
      TabIndex        =   348
      Top             =   1836
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   335
      Left            =   4104
      TabIndex        =   347
      Top             =   1620
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   334
      Left            =   4104
      TabIndex        =   346
      Top             =   1416
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   333
      Left            =   4104
      TabIndex        =   345
      Top             =   1200
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   332
      Left            =   4104
      TabIndex        =   344
      Top             =   996
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   331
      Left            =   4104
      TabIndex        =   343
      Top             =   1572
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   330
      Left            =   4104
      TabIndex        =   342
      Top             =   984
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   329
      Left            =   4104
      TabIndex        =   341
      Top             =   1188
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   328
      Left            =   4104
      TabIndex        =   340
      Top             =   1404
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   327
      Left            =   4104
      TabIndex        =   339
      Top             =   1608
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   326
      Left            =   4104
      TabIndex        =   338
      Top             =   1824
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   325
      Left            =   4104
      TabIndex        =   337
      Top             =   2028
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   324
      Left            =   4104
      TabIndex        =   336
      Top             =   816
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   323
      Left            =   4104
      TabIndex        =   335
      Top             =   2184
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   322
      Left            =   4104
      TabIndex        =   334
      Top             =   2028
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   321
      Left            =   4104
      TabIndex        =   333
      Top             =   1824
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   320
      Left            =   4104
      TabIndex        =   332
      Top             =   1608
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   319
      Left            =   4104
      TabIndex        =   331
      Top             =   1404
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   318
      Left            =   4104
      TabIndex        =   330
      Top             =   1188
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   317
      Left            =   4104
      TabIndex        =   329
      Top             =   984
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   316
      Left            =   4104
      TabIndex        =   328
      Top             =   1068
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   315
      Left            =   4104
      TabIndex        =   327
      Top             =   1284
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   314
      Left            =   4104
      TabIndex        =   326
      Top             =   1488
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   313
      Left            =   4104
      TabIndex        =   325
      Top             =   1704
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   312
      Left            =   4104
      TabIndex        =   324
      Top             =   1908
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   311
      Left            =   4104
      TabIndex        =   323
      Top             =   2028
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   310
      Left            =   4104
      TabIndex        =   322
      Top             =   2028
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   309
      Left            =   4104
      TabIndex        =   321
      Top             =   1824
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   308
      Left            =   4104
      TabIndex        =   320
      Top             =   1608
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   307
      Left            =   4104
      TabIndex        =   319
      Top             =   1404
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   306
      Left            =   4104
      TabIndex        =   318
      Top             =   1188
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   305
      Left            =   4104
      TabIndex        =   317
      Top             =   984
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   304
      Left            =   4104
      TabIndex        =   316
      Top             =   888
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   303
      Left            =   4104
      TabIndex        =   315
      Top             =   1104
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   302
      Left            =   4104
      TabIndex        =   314
      Top             =   1308
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   301
      Left            =   4104
      TabIndex        =   313
      Top             =   1524
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   300
      Left            =   4104
      TabIndex        =   312
      Top             =   1728
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   299
      Left            =   4104
      TabIndex        =   311
      Top             =   1944
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   298
      Left            =   4104
      TabIndex        =   310
      Top             =   1728
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   297
      Left            =   4104
      TabIndex        =   309
      Top             =   1224
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   296
      Left            =   4104
      TabIndex        =   308
      Top             =   1944
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   295
      Left            =   4104
      TabIndex        =   307
      Top             =   1728
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   294
      Left            =   4104
      TabIndex        =   306
      Top             =   1524
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   293
      Left            =   4104
      TabIndex        =   305
      Top             =   1308
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   292
      Left            =   4104
      TabIndex        =   304
      Top             =   1104
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   291
      Left            =   4104
      TabIndex        =   303
      Top             =   888
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   290
      Left            =   4104
      TabIndex        =   302
      Top             =   1512
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   289
      Left            =   4104
      TabIndex        =   301
      Top             =   1308
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   288
      Left            =   4104
      TabIndex        =   300
      Top             =   1092
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   287
      Left            =   4104
      TabIndex        =   299
      Top             =   888
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   286
      Left            =   4104
      TabIndex        =   298
      Top             =   1008
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   285
      Left            =   4104
      TabIndex        =   297
      Top             =   1212
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   284
      Left            =   4104
      TabIndex        =   296
      Top             =   1428
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   283
      Left            =   4104
      TabIndex        =   295
      Top             =   1632
      Width           =   1536
   End
   Begin VB.Label Lbl_COMBO_NUMBER 
      Alignment       =   2  'Center
      BackColor       =   &H00CCFFFF&
      Caption         =   "COMBO #"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   4164
      TabIndex        =   294
      Top             =   84
      Width           =   1452
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   282
      Left            =   396
      TabIndex        =   293
      Top             =   1260
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   281
      Left            =   396
      TabIndex        =   292
      Top             =   1488
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   280
      Left            =   396
      TabIndex        =   291
      Top             =   1692
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   279
      Left            =   396
      TabIndex        =   290
      Top             =   1908
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   278
      Left            =   396
      TabIndex        =   289
      Top             =   2112
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   277
      Left            =   396
      TabIndex        =   288
      Top             =   2328
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   276
      Left            =   396
      TabIndex        =   287
      Top             =   2532
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   275
      Left            =   396
      TabIndex        =   286
      Top             =   2748
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   274
      Left            =   2244
      TabIndex        =   285
      Top             =   1632
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   273
      Left            =   2244
      TabIndex        =   284
      Top             =   1428
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   272
      Left            =   2244
      TabIndex        =   283
      Top             =   1212
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   271
      Left            =   2244
      TabIndex        =   282
      Top             =   1008
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   270
      Left            =   396
      TabIndex        =   281
      Top             =   2952
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   269
      Left            =   396
      TabIndex        =   280
      Top             =   2832
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   268
      Left            =   2244
      TabIndex        =   279
      Top             =   888
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   267
      Left            =   2244
      TabIndex        =   278
      Top             =   1092
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   266
      Left            =   2244
      TabIndex        =   277
      Top             =   1308
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   265
      Left            =   2244
      TabIndex        =   276
      Top             =   1512
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   264
      Left            =   396
      TabIndex        =   275
      Top             =   2628
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   263
      Left            =   396
      TabIndex        =   274
      Top             =   2412
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   262
      Left            =   396
      TabIndex        =   273
      Top             =   2208
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   261
      Left            =   396
      TabIndex        =   272
      Top             =   1992
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   260
      Left            =   396
      TabIndex        =   271
      Top             =   1548
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   259
      Left            =   396
      TabIndex        =   270
      Top             =   1572
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   258
      Left            =   396
      TabIndex        =   269
      Top             =   1368
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   257
      Left            =   396
      TabIndex        =   268
      Top             =   1140
      Width           =   1536
   End
   Begin VB.Label Label4 
      BackColor       =   &H006FFBE2&
      Caption         =   "Label3"
      Height          =   312
      Left            =   9372
      TabIndex        =   267
      Top             =   84
      Width           =   396
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label3"
      Height          =   312
      Left            =   8928
      TabIndex        =   266
      Top             =   84
      Width           =   396
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      Height          =   312
      Left            =   8496
      TabIndex        =   263
      Top             =   84
      Width           =   396
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   256
      Left            =   396
      TabIndex        =   262
      Top             =   2844
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   255
      Left            =   2244
      TabIndex        =   261
      Top             =   888
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   254
      Left            =   2244
      TabIndex        =   260
      Top             =   1104
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   253
      Left            =   2244
      TabIndex        =   259
      Top             =   1308
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   252
      Left            =   2244
      TabIndex        =   258
      Top             =   1524
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   251
      Left            =   2244
      TabIndex        =   257
      Top             =   1728
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   250
      Left            =   2244
      TabIndex        =   256
      Top             =   1944
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   249
      Left            =   2244
      TabIndex        =   255
      Top             =   1224
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   248
      Left            =   396
      TabIndex        =   254
      Top             =   2628
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   247
      Left            =   396
      TabIndex        =   253
      Top             =   2424
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   246
      Left            =   396
      TabIndex        =   252
      Top             =   2208
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   245
      Left            =   396
      TabIndex        =   251
      Top             =   2004
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   244
      Left            =   396
      TabIndex        =   250
      Top             =   1788
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   243
      Left            =   396
      TabIndex        =   249
      Top             =   1584
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   242
      Left            =   396
      TabIndex        =   248
      Top             =   1368
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   241
      Left            =   396
      TabIndex        =   247
      Top             =   1140
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   240
      Left            =   396
      TabIndex        =   246
      Top             =   1140
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   239
      Left            =   396
      TabIndex        =   245
      Top             =   1368
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   238
      Left            =   396
      TabIndex        =   244
      Top             =   1584
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   237
      Left            =   396
      TabIndex        =   243
      Top             =   1788
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   236
      Left            =   396
      TabIndex        =   242
      Top             =   2004
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   235
      Left            =   396
      TabIndex        =   241
      Top             =   2208
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   234
      Left            =   396
      TabIndex        =   240
      Top             =   2424
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   233
      Left            =   396
      TabIndex        =   239
      Top             =   2628
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   232
      Left            =   2244
      TabIndex        =   238
      Top             =   1728
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   231
      Left            =   2244
      TabIndex        =   237
      Top             =   1944
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   230
      Left            =   2244
      TabIndex        =   236
      Top             =   1728
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   229
      Left            =   2244
      TabIndex        =   235
      Top             =   1524
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   228
      Left            =   2244
      TabIndex        =   234
      Top             =   1308
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   227
      Left            =   2244
      TabIndex        =   233
      Top             =   1104
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   226
      Left            =   2244
      TabIndex        =   232
      Top             =   888
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   225
      Left            =   396
      TabIndex        =   231
      Top             =   2844
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   224
      Left            =   396
      TabIndex        =   230
      Top             =   2724
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   223
      Left            =   396
      TabIndex        =   229
      Top             =   2928
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   222
      Left            =   2244
      TabIndex        =   228
      Top             =   984
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   221
      Left            =   2244
      TabIndex        =   227
      Top             =   1188
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   220
      Left            =   2244
      TabIndex        =   226
      Top             =   1404
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   219
      Left            =   2244
      TabIndex        =   225
      Top             =   1608
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   218
      Left            =   2244
      TabIndex        =   224
      Top             =   1824
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   217
      Left            =   2244
      TabIndex        =   223
      Top             =   2028
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   216
      Left            =   396
      TabIndex        =   222
      Top             =   2508
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   215
      Left            =   396
      TabIndex        =   221
      Top             =   2304
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   214
      Left            =   396
      TabIndex        =   220
      Top             =   2088
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   213
      Left            =   396
      TabIndex        =   219
      Top             =   1884
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   212
      Left            =   396
      TabIndex        =   218
      Top             =   1668
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   211
      Left            =   396
      TabIndex        =   217
      Top             =   1464
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   210
      Left            =   396
      TabIndex        =   216
      Top             =   1248
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   209
      Left            =   396
      TabIndex        =   215
      Top             =   1020
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   208
      Left            =   396
      TabIndex        =   214
      Top             =   1116
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   207
      Left            =   396
      TabIndex        =   213
      Top             =   1344
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   206
      Left            =   396
      TabIndex        =   212
      Top             =   1548
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   205
      Left            =   396
      TabIndex        =   211
      Top             =   1524
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   204
      Left            =   396
      TabIndex        =   210
      Top             =   1968
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   203
      Left            =   396
      TabIndex        =   209
      Top             =   2184
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   202
      Left            =   396
      TabIndex        =   208
      Top             =   2388
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   201
      Left            =   396
      TabIndex        =   207
      Top             =   2604
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   200
      Left            =   2244
      TabIndex        =   206
      Top             =   2028
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   199
      Left            =   2244
      TabIndex        =   205
      Top             =   1908
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   198
      Left            =   2244
      TabIndex        =   204
      Top             =   1704
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   197
      Left            =   2244
      TabIndex        =   203
      Top             =   1488
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   196
      Left            =   2244
      TabIndex        =   202
      Top             =   1284
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   195
      Left            =   2244
      TabIndex        =   201
      Top             =   1068
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   194
      Left            =   396
      TabIndex        =   200
      Top             =   3024
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   193
      Left            =   396
      TabIndex        =   199
      Top             =   2808
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   192
      Left            =   396
      TabIndex        =   198
      Top             =   2928
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   191
      Left            =   2244
      TabIndex        =   197
      Top             =   984
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   190
      Left            =   2244
      TabIndex        =   196
      Top             =   1188
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   189
      Left            =   2244
      TabIndex        =   195
      Top             =   1404
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   188
      Left            =   2244
      TabIndex        =   194
      Top             =   1608
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   187
      Left            =   2244
      TabIndex        =   193
      Top             =   1824
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   186
      Left            =   2244
      TabIndex        =   192
      Top             =   2028
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   185
      Left            =   2244
      TabIndex        =   191
      Top             =   2184
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   184
      Left            =   396
      TabIndex        =   190
      Top             =   2724
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   183
      Left            =   396
      TabIndex        =   189
      Top             =   2508
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   182
      Left            =   396
      TabIndex        =   188
      Top             =   2304
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   181
      Left            =   396
      TabIndex        =   187
      Top             =   2088
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   180
      Left            =   396
      TabIndex        =   186
      Top             =   1884
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   179
      Left            =   396
      TabIndex        =   185
      Top             =   1668
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   178
      Left            =   396
      TabIndex        =   184
      Top             =   1464
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   177
      Left            =   396
      TabIndex        =   183
      Top             =   1236
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   176
      Left            =   396
      TabIndex        =   182
      Top             =   1236
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   175
      Left            =   396
      TabIndex        =   181
      Top             =   1464
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   174
      Left            =   396
      TabIndex        =   180
      Top             =   1668
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   173
      Left            =   396
      TabIndex        =   179
      Top             =   1884
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   172
      Left            =   396
      TabIndex        =   178
      Top             =   2088
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   171
      Left            =   396
      TabIndex        =   177
      Top             =   2304
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   170
      Left            =   396
      TabIndex        =   176
      Top             =   2508
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   169
      Left            =   396
      TabIndex        =   175
      Top             =   2724
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   168
      Left            =   2244
      TabIndex        =   174
      Top             =   816
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   167
      Left            =   2244
      TabIndex        =   173
      Top             =   2028
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   166
      Left            =   2244
      TabIndex        =   172
      Top             =   1824
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   165
      Left            =   2244
      TabIndex        =   171
      Top             =   1608
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   164
      Left            =   2244
      TabIndex        =   170
      Top             =   1404
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   163
      Left            =   2244
      TabIndex        =   169
      Top             =   1188
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   162
      Left            =   2244
      TabIndex        =   168
      Top             =   984
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   161
      Left            =   396
      TabIndex        =   167
      Top             =   2928
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   160
      Left            =   2244
      TabIndex        =   165
      Top             =   1572
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   159
      Left            =   396
      TabIndex        =   164
      Top             =   2520
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   158
      Left            =   396
      TabIndex        =   163
      Top             =   2736
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   157
      Left            =   396
      TabIndex        =   162
      Top             =   2940
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   156
      Left            =   2244
      TabIndex        =   161
      Top             =   996
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   155
      Left            =   2244
      TabIndex        =   160
      Top             =   1200
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   154
      Left            =   2244
      TabIndex        =   159
      Top             =   1416
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   153
      Left            =   2244
      TabIndex        =   158
      Top             =   1620
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   152
      Left            =   2244
      TabIndex        =   157
      Top             =   1836
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   151
      Left            =   396
      TabIndex        =   156
      Top             =   2316
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   150
      Left            =   396
      TabIndex        =   155
      Top             =   2100
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   149
      Left            =   396
      TabIndex        =   154
      Top             =   1896
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   148
      Left            =   396
      TabIndex        =   153
      Top             =   1680
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   147
      Left            =   396
      TabIndex        =   152
      Top             =   1476
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   146
      Left            =   396
      TabIndex        =   151
      Top             =   1260
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   145
      Left            =   396
      TabIndex        =   150
      Top             =   1056
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   144
      Left            =   396
      TabIndex        =   149
      Top             =   828
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   143
      Left            =   396
      TabIndex        =   148
      Top             =   828
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   142
      Left            =   396
      TabIndex        =   147
      Top             =   1056
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   141
      Left            =   396
      TabIndex        =   146
      Top             =   1260
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   140
      Left            =   396
      TabIndex        =   145
      Top             =   1476
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   139
      Left            =   396
      TabIndex        =   144
      Top             =   1680
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   138
      Left            =   396
      TabIndex        =   143
      Top             =   1896
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   137
      Left            =   396
      TabIndex        =   142
      Top             =   2100
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   136
      Left            =   396
      TabIndex        =   141
      Top             =   2316
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   135
      Left            =   2244
      TabIndex        =   140
      Top             =   1836
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   134
      Left            =   2244
      TabIndex        =   139
      Top             =   1620
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   133
      Left            =   2244
      TabIndex        =   138
      Top             =   1416
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   132
      Left            =   2244
      TabIndex        =   137
      Top             =   1200
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   131
      Left            =   2244
      TabIndex        =   136
      Top             =   996
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   130
      Left            =   396
      TabIndex        =   135
      Top             =   2940
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   129
      Left            =   396
      TabIndex        =   134
      Top             =   2736
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   128
      Left            =   396
      TabIndex        =   133
      Top             =   2520
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   127
      Left            =   396
      TabIndex        =   132
      Top             =   2400
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   126
      Left            =   396
      TabIndex        =   131
      Top             =   2616
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   125
      Left            =   396
      TabIndex        =   130
      Top             =   2820
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   124
      Left            =   2244
      TabIndex        =   129
      Top             =   876
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   123
      Left            =   2244
      TabIndex        =   128
      Top             =   1080
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   122
      Left            =   2244
      TabIndex        =   127
      Top             =   1296
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   121
      Left            =   2244
      TabIndex        =   126
      Top             =   1500
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   120
      Left            =   2244
      TabIndex        =   125
      Top             =   1716
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   119
      Left            =   396
      TabIndex        =   124
      Top             =   2196
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   118
      Left            =   396
      TabIndex        =   123
      Top             =   1980
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   117
      Left            =   396
      TabIndex        =   122
      Top             =   1776
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   116
      Left            =   396
      TabIndex        =   121
      Top             =   1560
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   115
      Left            =   396
      TabIndex        =   120
      Top             =   1356
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   114
      Left            =   396
      TabIndex        =   119
      Top             =   1140
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   113
      Left            =   396
      TabIndex        =   118
      Top             =   936
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   112
      Left            =   2244
      TabIndex        =   117
      Top             =   2364
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   111
      Left            =   2244
      TabIndex        =   116
      Top             =   2280
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   110
      Left            =   396
      TabIndex        =   115
      Top             =   840
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   109
      Left            =   396
      TabIndex        =   114
      Top             =   1056
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   108
      Left            =   396
      TabIndex        =   113
      Top             =   1260
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   107
      Left            =   396
      TabIndex        =   112
      Top             =   1476
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   106
      Left            =   396
      TabIndex        =   111
      Top             =   1680
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   105
      Left            =   396
      TabIndex        =   110
      Top             =   1896
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   104
      Left            =   396
      TabIndex        =   109
      Top             =   2100
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   103
      Left            =   2244
      TabIndex        =   108
      Top             =   1620
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   102
      Left            =   2244
      TabIndex        =   107
      Top             =   1416
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   101
      Left            =   2244
      TabIndex        =   106
      Top             =   1200
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   100
      Left            =   2244
      TabIndex        =   105
      Top             =   996
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   99
      Left            =   396
      TabIndex        =   104
      Top             =   2940
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   98
      Left            =   396
      TabIndex        =   103
      Top             =   2736
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   97
      Left            =   396
      TabIndex        =   102
      Top             =   2520
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   96
      Left            =   396
      TabIndex        =   101
      Top             =   2316
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   95
      Left            =   396
      TabIndex        =   100
      Top             =   2436
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   94
      Left            =   396
      TabIndex        =   99
      Top             =   2640
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   93
      Left            =   396
      TabIndex        =   98
      Top             =   2856
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   92
      Left            =   2244
      TabIndex        =   97
      Top             =   900
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   91
      Left            =   2244
      TabIndex        =   96
      Top             =   1116
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   90
      Left            =   2244
      TabIndex        =   95
      Top             =   1320
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   89
      Left            =   2244
      TabIndex        =   94
      Top             =   1536
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   88
      Left            =   2244
      TabIndex        =   93
      Top             =   1740
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   87
      Left            =   396
      TabIndex        =   92
      Top             =   2220
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   86
      Left            =   396
      TabIndex        =   91
      Top             =   2016
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   85
      Left            =   396
      TabIndex        =   90
      Top             =   1800
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   84
      Left            =   396
      TabIndex        =   89
      Top             =   1596
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   83
      Left            =   396
      TabIndex        =   88
      Top             =   1380
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   82
      Left            =   396
      TabIndex        =   87
      Top             =   1176
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   81
      Left            =   396
      TabIndex        =   86
      Top             =   960
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   80
      Left            =   2244
      TabIndex        =   85
      Top             =   2400
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   79
      Left            =   2244
      TabIndex        =   84
      Top             =   2400
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   78
      Left            =   396
      TabIndex        =   83
      Top             =   960
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   77
      Left            =   396
      TabIndex        =   82
      Top             =   1176
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   76
      Left            =   396
      TabIndex        =   81
      Top             =   1380
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   75
      Left            =   396
      TabIndex        =   80
      Top             =   1596
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   74
      Left            =   396
      TabIndex        =   79
      Top             =   1800
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   73
      Left            =   396
      TabIndex        =   78
      Top             =   2016
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   72
      Left            =   396
      TabIndex        =   77
      Top             =   2220
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   71
      Left            =   2244
      TabIndex        =   76
      Top             =   1740
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   70
      Left            =   2244
      TabIndex        =   75
      Top             =   1536
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   69
      Left            =   2244
      TabIndex        =   74
      Top             =   1320
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   68
      Left            =   2244
      TabIndex        =   73
      Top             =   1116
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   67
      Left            =   2244
      TabIndex        =   72
      Top             =   900
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   66
      Left            =   396
      TabIndex        =   71
      Top             =   2856
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   65
      Left            =   396
      TabIndex        =   70
      Top             =   2640
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   64
      Left            =   396
      TabIndex        =   69
      Top             =   2436
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   63
      Left            =   396
      TabIndex        =   68
      Top             =   2316
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   62
      Left            =   396
      TabIndex        =   67
      Top             =   2520
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   61
      Left            =   396
      TabIndex        =   66
      Top             =   2736
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   60
      Left            =   396
      TabIndex        =   65
      Top             =   2940
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   59
      Left            =   2244
      TabIndex        =   64
      Top             =   996
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   58
      Left            =   2244
      TabIndex        =   63
      Top             =   1200
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   57
      Left            =   2244
      TabIndex        =   62
      Top             =   1416
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   56
      Left            =   2244
      TabIndex        =   61
      Top             =   1620
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   55
      Left            =   396
      TabIndex        =   60
      Top             =   2100
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   54
      Left            =   396
      TabIndex        =   59
      Top             =   1896
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   53
      Left            =   396
      TabIndex        =   58
      Top             =   1680
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   52
      Left            =   396
      TabIndex        =   57
      Top             =   1476
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   51
      Left            =   396
      TabIndex        =   56
      Top             =   1260
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   50
      Left            =   396
      TabIndex        =   55
      Top             =   1056
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   49
      Left            =   396
      TabIndex        =   54
      Top             =   840
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   48
      Left            =   2244
      TabIndex        =   53
      Top             =   2280
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   47
      Left            =   2244
      TabIndex        =   52
      Top             =   2400
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   46
      Left            =   396
      TabIndex        =   51
      Top             =   960
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   45
      Left            =   396
      TabIndex        =   50
      Top             =   1176
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   44
      Left            =   396
      TabIndex        =   49
      Top             =   1380
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   43
      Left            =   396
      TabIndex        =   48
      Top             =   1596
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   42
      Left            =   396
      TabIndex        =   47
      Top             =   1800
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   41
      Left            =   396
      TabIndex        =   46
      Top             =   2016
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   40
      Left            =   396
      TabIndex        =   45
      Top             =   2220
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   39
      Left            =   2244
      TabIndex        =   44
      Top             =   1740
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   38
      Left            =   2244
      TabIndex        =   43
      Top             =   1536
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   37
      Left            =   2244
      TabIndex        =   42
      Top             =   1320
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   36
      Left            =   2244
      TabIndex        =   41
      Top             =   1116
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   35
      Left            =   2244
      TabIndex        =   40
      Top             =   900
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   34
      Left            =   396
      TabIndex        =   39
      Top             =   2856
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   33
      Left            =   396
      TabIndex        =   38
      Top             =   2640
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   32
      Left            =   396
      TabIndex        =   37
      Top             =   2436
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   31
      Left            =   2244
      TabIndex        =   36
      Top             =   2256
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   30
      Left            =   396
      TabIndex        =   35
      Top             =   828
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   29
      Left            =   396
      TabIndex        =   34
      Top             =   1044
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   28
      Left            =   396
      TabIndex        =   33
      Top             =   1248
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   27
      Left            =   396
      TabIndex        =   32
      Top             =   1464
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   26
      Left            =   396
      TabIndex        =   31
      Top             =   1668
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   25
      Left            =   396
      TabIndex        =   30
      Top             =   1884
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   24
      Left            =   396
      TabIndex        =   29
      Top             =   2088
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   23
      Left            =   2244
      TabIndex        =   28
      Top             =   1608
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   22
      Left            =   2244
      TabIndex        =   27
      Top             =   1404
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   21
      Left            =   2244
      TabIndex        =   26
      Top             =   1188
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   20
      Left            =   2244
      TabIndex        =   25
      Top             =   984
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   19
      Left            =   396
      TabIndex        =   24
      Top             =   2928
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   18
      Left            =   396
      TabIndex        =   23
      Top             =   2724
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   17
      Left            =   396
      TabIndex        =   22
      Top             =   2508
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   16
      Left            =   396
      TabIndex        =   21
      Top             =   2304
      Width           =   1536
   End
   Begin VB.Label Lbl3 
      Alignment       =   2  'Center
      Caption         =   "*-- - Last Winamp Loads and Playing - --*"
      Height          =   228
      Left            =   8472
      TabIndex        =   20
      Top             =   2184
      Visible         =   0   'False
      Width           =   3216
   End
   Begin VB.Label Lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Shell NotePad Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   9816
      TabIndex        =   18
      Top             =   84
      Visible         =   0   'False
      Width           =   2112
   End
   Begin VB.Label Lbl_Title 
      Alignment       =   2  'Center
      BackColor       =   &H00CCFFFF&
      Caption         =   "---- Load Last File Click Here ---- "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   156
      TabIndex        =   16
      Top             =   84
      Width           =   3948
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   15
      Left            =   396
      TabIndex        =   15
      Top             =   2496
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   14
      Left            =   396
      TabIndex        =   14
      Top             =   2700
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   13
      Left            =   396
      TabIndex        =   13
      Top             =   2916
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   12
      Left            =   2244
      TabIndex        =   12
      Top             =   960
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   11
      Left            =   2244
      TabIndex        =   11
      Top             =   1176
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   10
      Left            =   2244
      TabIndex        =   10
      Top             =   1380
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   9
      Left            =   2244
      TabIndex        =   9
      Top             =   1596
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   8
      Left            =   2244
      TabIndex        =   8
      Top             =   1800
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   7
      Left            =   396
      TabIndex        =   7
      Top             =   2280
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   6
      Left            =   396
      TabIndex        =   6
      Top             =   2076
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   5
      Left            =   396
      TabIndex        =   5
      Top             =   1860
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   4
      Left            =   396
      TabIndex        =   4
      Top             =   1656
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   3
      Left            =   396
      TabIndex        =   3
      Top             =   1440
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   2
      Left            =   396
      TabIndex        =   2
      Top             =   1236
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   1
      Left            =   396
      TabIndex        =   1
      Top             =   1020
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   0
      Left            =   396
      TabIndex        =   0
      Top             =   804
      Width           =   1536
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_VERSION 
      Caption         =   "VERSION"
   End
   Begin VB.Menu Mnu_Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Mnu_Sync 
         Caption         =   "Sync"
      End
   End
   Begin VB.Menu MNU_TREESIZE 
      Caption         =   "TREESIZE"
   End
   Begin VB.Menu MNU_OPEN_PROPERTIE_FOR_ALL_DRIVE 
      Caption         =   "OPEN PROPERTY FOR ALL DRIVE"
   End
   Begin VB.Menu MNU_OPEN_PROPERTY_ALL_FOLDER_C_HDD 
      Caption         =   "OPEN PROPERTY ALL FOLDER C-HDD"
   End
   Begin VB.Menu Mnu_LoadFolder 
      Caption         =   "LOAD FOLDER LINK"
   End
   Begin VB.Menu MNU_NETWORK_WITH_CLIPBOARDER 
      Caption         =   "NETWORK PATH WITH CLIPBOARDER CONTENT"
   End
   Begin VB.Menu MNU_CLIPBOARDOR 
      Caption         =   "CLIPBOARD ITEM TEXT"
   End
   Begin VB.Menu MNU_CLIPBOARDOR_PATH_LINK 
      Caption         =   "CLIPBOARD ITEM PATH LINKER"
   End
   Begin VB.Menu MNU_CLIPBOARDOR_PATH_NAME 
      Caption         =   "CLIPBOARD ITEM PATH NAME"
   End
   Begin VB.Menu MNU_CLIP_PATH_NAME_SHORT 
      Caption         =   "CLIPBOARD ITEM PATH SHORT"
   End
   Begin VB.Menu MNU_CLIPBOARD_COMPUTER_NAME_SET 
      Caption         =   "CLIPBOARD COMPUTER NAME SET"
   End
   Begin VB.Menu MNU_CLIPBOARD_COMPUTER_NAME_SET_VB6 
      Caption         =   "&& CLIPPER VB6 "
   End
   Begin VB.Menu MNU_CLIPBOARD_ALL_NET_PATH 
      Caption         =   "CLIPBOARD ALL NET PATH"
   End
   Begin VB.Menu MNU_CLIPBOARD_ALL_NET_PATH_REVERSE 
      Caption         =   "CLIPBOARD ALL NET PATH REVERSE"
   End
   Begin VB.Menu MNU_CLIPBOARD_ALL_NET_PATH_ARRAY_BUILDER 
      Caption         =   "CLIPBOARD ALL NET PATH AS ARRAY BUILDER"
   End
   Begin VB.Menu MNU_LOAD_ALL_GOODSYNC_PROFILE_TIC_NOTEPAD 
      Caption         =   "LOAD ALL GOODSYNC PROFILE IN NOTEPAD"
   End
   Begin VB.Menu MNU_NET_2_GO 
      Caption         =   "NET 2 GO = NOP"
   End
   Begin VB.Menu MNU_NETWORK_2_STEP_JUMPER 
      Caption         =   "NETWORK 2 STEP JUMPER"
   End
   Begin VB.Menu MNU_NETWORK_2_STEP_DRIVE_SELECTOR 
      Caption         =   "NETWORK DRIVE SELECT - DONE"
   End
   Begin VB.Menu MNU_GOOGLE_CHROME_PROFILE 
      Caption         =   "G CHROME PROFILE"
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   0
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   1
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   2
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   3
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   4
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   5
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   6
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   7
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   8
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   9
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   10
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   11
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   12
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   13
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   14
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   15
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   16
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   17
      End
      Begin VB.Menu MNU_GC_PROFILE 
         Caption         =   "MNU_GC_PROFILE"
         Index           =   18
      End
   End
   Begin VB.Menu MNU_LOAD_CHROME_PROFILE_FOLDER 
      Caption         =   "LOAD CHROME PROFILE FOLDER"
   End
   Begin VB.Menu MNU_LOAD_ALL_SCRIPTOR 
      Caption         =   "LOAD ALL SCRIPTOR"
   End
   Begin VB.Menu MNU_LOAD_ALL_SCRIPTOR_AHK_BAT_VBS 
      Caption         =   "LOAD ALL AHK BAT VBS "
   End
   Begin VB.Menu MNU_SPACE_EMPTY 
      Caption         =   "SPACE EMPTY"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------
' ADD HEADER TITLE FOR NETWORK USER FOLDER
' ----------------------------------
' With ScanPath.ListView1
'     Set LV = .ListItems.Add(, , "\NETWORK USER FOLDER\")
'     LV.SubItems(1) = "TITLE_BLOCK"
' End With
' NOT FINISH GOT A -- IF 1 = 2 THEN END IF -- LINE
' REINTRODUCE NOT SURE IF DONE YET -- MAY REQUIRE UPDATE INFO
' Sat 30-May-2020 11:53:28
' ----------------------------------


' Form1_Height -- SETTER
' SH = Screen.Height - 8000 'higer = smaller
' SH = Me.Height - 2000
' Me.WindowState = vbMaximized
' Me.WindowState = vbNormal


' ------------------------------------------------------------------------------
' SESSION COUNT NUMBER MASSIVE
' 002 SESSION
' 01.. -------------------------------------------------------------------------
' MASSIVE WORK TONIGHT
' ADD NEW NETWORK FOLDER FOR USER
' SHAME ABOUT THE HARD CODER FOR EXTRA AT THE MOMENT
' TO SHOW WHICH IS MIAN LOGGIONG USER FOR EACH COMPUTER
' WAS GONG TO CODE SOME VB NET 2008
' TEST CODE WASN'T ALLOW ACCESS OVER NETWORK
' 02.. -------------------------------------------------------------------------
' ADD CODE FOR GOOD SYNC PATH FOR PROFILE EVERY COMPUTER NETOWRK
' AND MAKE A GOODSYNC PROFILE LOADER FOR NOTEPAD++
' THESE TPYE OF FILE NOT WANT NOTEPAD++ AS CHANGE SO OFTEN
' SO LOAD IN WORK AND MOVE OUT AGAIN
' 03.. -------------------------------------------------------------------------
' IN ORDER TO MAKE EASIER WITH PATH AND SHOW ANOTHR NAME THAT HAD REPLACE STRING
' FOR SMALLER NEATER USE ANOTHER ARRAY PART OF LISTVIEW
' ------------------------------------------------------------------------------
' TIME SLICE WORKER
' ------------------------------------------------------------------------------
' Tue 03-Sep-2019 12:33:48 -- STARTED AND WORK OTHER THING FOR WHILE MORE
' Wed 04-Sep-2019 01:45:10
' ------------------------------------------------------------------------------

' ------------------------------------------------------------------------------
' SESSION COUNT NUMBER MASSIVE
' 003 SESSION
' 01.. -------------------------------------------------------------------------
' WORK TODAY LONG
' MAKE THE LAUNCHER WHEN GET TO EXPLORER
' THEY MANY WAY TO LOAD EXPLORER
' NOW BEEN REDUCED
' TO WHEN EXPLORER WANT TO LOAD IT NOW DO BY ONE ROUTINE CALL
' AND ANOTHER ONE FOR CMD SHELL WHICH USE LESS DEMANDER
' AND ADD NEW OPTION IN MENU
' TO LOAD TREESIZE FOR ANY FOLDER GIVEN
' AND THAT VICIOUS BIT
' WAS HARD WORKER
' FOND SOMETHING OUT FIST TIME
' AND A LEARN GAP BETWEEN US NOIVCE AND GET STARTED ON SOMETHING
' SEEM IN THE DAY THAT WERE SOME SOURCE OF FIND TETX COMPUTER CODE
' NOW A DAY  IS THE INTERNET
' VERY FEW EXAMPLE AROUND ACAULLY SHOW
' WHEN DO THING TO A BIT EXTRA DEPTH
' TREESIZE DOES NOT LIKE TO LOAD FROM VB6
' IT CRASHES
' SO I MAKE A LAUCHER
' AND THE ARGMENT REQUIRE SEND OVER
' AND THAT WHRE THAT MAGIC HAPPEN
' Public Sub EXPLORER_WITH_SHELL(SELECT_OPTION, FOLDER_NAME)
' THAT IS A VARIABEL THAT HOLD LENGHT OF ITEM FOR COMBO BOX ISTOAY
' AND INCREASE LENGHT THAT
' THAT ITEM ARE HIGHTLIGHT DARK BULLET ON
' ------------------------------------------------------------------------------
' WORK TODAY
' Wed 18-Sep-2019 20:08:55
' Wed 18-Sep-2019 22:55:17
' ------------------------------------------------------------------------------

' ------------------------------------------------------------------------------
' SESSION 004
' ------------------------------------------------------------------------------
' 005 PROJECT TODAY
' MAKE DOWNLOAD FOLDER SHOW
' CAPITAL A LOT OF THING HARDCODER
' TIDY SORTER LINKER AND CAPITAL URL SEMI CODER
' MAKE NET 2 GO ANOTHER METHOD VIDEO
' ------------------------------------------------------------------------------
' Thu 28-May-2020 14:00:18
' Thu 28-May-2020 15:20:00 -- 1 HOUR 19 MINUTE
' ------------------------------------------------------------------------------

Dim EXE_PATH()


Dim FR4


Dim MNU_NET_2_GO_VALUE

Dim FHEIGHTX
Dim OLD_MENUBAR
Dim DO_ANOTHER_COMBO1_WIDTH
Dim OLD_SCREEN_WIDTH


' -------------------------------------
' NOTE FORM VBS COMMON LABEL NAME CLASH
' SO CALL ShowWindow_2
' -------------------------------------
Const ShowWindow_2 = 1, DontShowWindow = 0, DontWaitUntilFinished = False, WaitUntilFinished = True


Dim CLIPBOARDOR
Dim FontSizez
Dim FontSizez_2


Dim FR1
Dim m_CRC

Public cProcesses As New clsCnProc

Dim Combo_KEYCODE_NOT_EVENT
Dim OLD_Combo1_ListIndex
   
Dim R_RGB As Integer
Dim G_RGB As Integer
Dim B_RGB As Integer

Dim TRACER
Dim SET_TRACER_ONCE
Dim TEXT_PATH_4_ARRAY()
Dim X_COLOR_ARRAY()


Dim STRING_COMBO_NUMBER
Dim LABLE_BACKCOLOR_SET

Dim Form1_Width, Form1_Height
Dim X_Y_DONE_ONCE
'------------
Public NAME_FORM

Dim XXT2
Dim XXT4
Dim XXT8

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long

Private Declare Function FindWindow2 _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer

Private Const MAX_PATH = 260

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type MENUBARINFO
  cbSize As Long
  rcBar As Rect
  hMenu As Long
  hWndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type

Private MenuInfo As MENUBARINFO
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

Private Type RGBColor
     Red As Byte
     Green As Byte
     Blue As Byte
End Type

Private Type HSBColor
     Hue As Double
     Saturation As Double
     Brightness As Double
End Type



' ---------------------------------------------------------------------
' MAKE SEPARATE ROUTINE AS
' SHELLEXECUTEEX_OPEN_PROPERTIE
' ---------------------------------------------------------------------
' WHEN EXECUTE COMMAND
' IF CLOSE APP THAT LAUNCH THEM
' THEY CLOSE ALSO
' ALSO ON THE INTERNET
' MANY TALKER
' AND THE REQUEST WAY
' IS USE CREATE PROCESS WHICH RETURN A LOT MORE VARIABLE
' BUT THAT NOT FEATURE THE ITEM REQUIREING
' SO BATTLE THOUGHT THAT MIND FIELD
' THE APP WAS CLOSE AT TOP TALK AND CLOSE THE REQUEST APP THAT RUN
' THAT IDEA FIND THE HWND FROM THE PID
' A BIT RUSTY ABOUT THAT ONE
' AND LOT WORKER -- PROCESS NIGHT MARE
' EVEN THEN WAS WAIT UNTIL PRCOCESS DONE
' I GOT FEW SHOP WANT THIS ONE
' RESOLVER BY CLEVER CHANGE THE HWND AND ATTACH TO ANOTHER APP THE DESKTOP WINDOW
' AFTER THAT RESULT WANTER
' SetParent SEI.hWnd, lHnd
' THIS HERE CODE
' WILL OPEN UP A PROPERTIES DIALOG
' SO ABLE ADJUST THE PERMISSION NOT WITH EXPLORER
' AND WILL HANG A LITTLE BIT AFTER A WHILE
' IS RUOTINE TO LOAD THEM ALL
' OR CALLER PROGRAM AND USE OTHER THING
' LOAD ONE CALLER APP LOAD THE SET ALL THEM
' ---------------------------------------------------------------------

' NEWER ADDITION PUBLISHER
' ---------------------------------------------------------------------

' TIME TO FIND HERE BEFORE WORKER
' --------------------------------------------------------------------
' 0001
' Fri 26-Jun-2020 10:44:17
' Fri 26-Jun-2020 12:00:00 -- 1 HOUR 15 MINUTE
' --------------------------------------------------------------------
' SHELLEXECUTE TO SPECIFY THE PROPERTIES VERB VB6 - GOOGLE SEARCH
' https://www.google.co.uk/search?safe=active&sxsrf=ALeKk02PtyW3kzi1wOQHAHk707nxgZsszg%3A1593167672498&ei=OM_1Xp-DHtujgAbcyamgAg&q=ShellExecute+to+specify+the+properties+verb+VB6&oq=ShellExecute+to+specify+the+properties+verb+VB6&gs_lcp=CgZwc3ktYWIQAzoHCAAQRxCwAzoHCCMQrgIQJzoFCCEQoAFQjdkOWOfeDmCF4g5oAXAAeACAAZ0BiAGvA5IBAzIuMpgBAKABAaoBB2d3cy13aXo&sclient=psy-ab&ved=0ahUKEwif_76JpJ_qAhXbEcAKHdxkCiQQ4dUDCAw&uact=5
' ---------------------------------------------------------------------
' 0001
' Fri 26-Jun-2020 10:44:17
' Fri 26-Jun-2020 12:28:10 -- 1 HOUR 43 MINUTE
' ---------------------------------------------------------------------
' TOTAL -- 1 + 2 -- 43+42+47+27 -- 159 -- 2 HOUR 39 + 1 + 2 -- 5 HOUR 39
' ---------------------------------------------------------------------
' Fri 26-Jun-2020 10:44:17 -- START TIME
' Sat 27-Jun-2020 06:22:51 -- END -- 19 HOUR 38 MINUTE
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
' 0002
' Fri 26-Jun-2020 14:34:56 --
' Fri 26-Jun-2020 15:17:48 -- 42 MINUTE
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
' 0003
' Fri 26-Jun-2020 14:23:31
' Fri 26-Jun-2020 17:11:08 -- 2 HOUR 47 MINUTE
' HEAD EXPLODE
' GO WORK ANOTHER ITEM UPSTAIR BEDROOM -- GOODSYNC STUFF WORK
' SUCCESSFUL RESULT
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
' 0004
' Sat 27-Jun-2020 05:54:59
' Sat 27-Jun-2020 06:22:51 -- 27 MINUTE
' END TIME FIND RESULT TO SETHWND HANDLE TO ANOTHER APP
' THAT WAY MAKE WHEN QUIT PROGRAM HERE NOT PULL THE OTHER APP TO ALSO
' ---------------------------------------------------------------------
' TOTAL -- 1 + 2 -- 43+42+47+27 -- 159 -- 2 HOUR 39 + 1 + 2 -- 5 HOUR 39
' Fri 26-Jun-2020 10:44:17 -- START TIME
' Sat 27-Jun-2020 06:22:51 -- END -- 19 HOUR 38 MINUTE
' ---------------------------------------------------------------------


' ---------------------------------------------------------------------
' TWO SESSION -- 1 HOUR 43 MINUTE __+__ 42 MINUTE
' FIND ERROR CODE NOT CLOSE APP -- AND DESTROY CALL APP WINDOW
' WHEN CALL BY APP CLOSE OTHER DO
' ---------------------------------------------------------------------
' ---------------------------------------------------------------------
' Fri 26-Jun-2020 10:44:17 -- TOTAL SMALL WITH GAP BETWEEN
' Fri 26-Jun-2020 15:17:48 -- 4 HOUR 33 MINUTE
' ---------------------------------------------------------------------
' --------------------------------------------------------------------
' TIME TO FIND HERE BEFORE WORKER
' --------------------------------------------------------------------
' Fri 26-Jun-2020 10:44:17
' Fri 26-Jun-2020 12:00:00 -- 1 HOUR 15 MINUTE
' --------------------------------------------------------------------
' ShellExecute to specify the properties verb VB6 - Google Search
' https://www.google.co.uk/search?safe=active&sxsrf=ALeKk02PtyW3kzi1wOQHAHk707nxgZsszg%3A1593167672498&ei=OM_1Xp-DHtujgAbcyamgAg&q=ShellExecute+to+specify+the+properties+verb+VB6&oq=ShellExecute+to+specify+the+properties+verb+VB6&gs_lcp=CgZwc3ktYWIQAzoHCAAQRxCwAzoHCCMQrgIQJzoFCCEQoAFQjdkOWOfeDmCF4g5oAXAAeACAAZ0BiAGvA5IBAzIuMpgBAKABAaoBB2d3cy13aXo&sclient=psy-ab&ved=0ahUKEwif_76JpJ_qAhXbEcAKHdxkCiQQ4dUDCAw&uact=5
' --------
' ShellExecute and "properties" verb
' http://forums.codeguru.com/showthread.php?2125-ShellExecute-and-properties-verb
' --------------------------------------------------------------------
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    ' ---- OPTIONAL FIELDS
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef s As SHELLEXECUTEINFO) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long





' --------------------------------------------------------------------
' --------------------------------------------------------------------
' SET OF DECLARE WORTH HAVE ALL CODE PROJECT BEGGINER
' --------------------------------------------------------------------
' --------------------------------------------------------------------
' BAS MODULE DECLARE
' Private Declare Function EnumChildWindows Lib "user32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
' --------------------------------------------------------------------
'Private Type Rect
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'Private Type POINTAPI
'        x As Long
'        y As Long
'End Type
'Private Type MENUBARINFO
'  cbSize As Long
'  rcBar As Rect
'  hMenu As Long
'  hWndMenu As Long
'  fBarFocused As Boolean
'  fFocused As Boolean
'End Type
'Private MenuInfo As MENUBARINFO
'Private Const OBJID_MENU As Long = &HFFFFFFFD
'Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
' Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
' Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
'----------------------------------------------------------------------------------
Private Type FILETIME
   LowDateTime          As Long
   HighDateTime         As Long
End Type
Private Type WIN32_FIND_DATA
   dwFileAttributes     As Long
   ftCreationTime       As FILETIME
   ftLastAccessTime     As FILETIME
   ftLastWriteTime      As FILETIME
   nFileSizeHigh        As Long
   nFileSizeLow         As Long
   dwReserved0          As Long
   dwReserved1          As Long
   cFileName            As String * 260  'MUST be set to 260
   cAlternate           As String * 14
End Type
' --------------------------------------------------------------------
' --------------------------------------------------------------------
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
' --------------------------------------------------------------------
Const GWL_WNDPROC = -4
Const GWL_HINSTANCE = -6
Const GWL_HWNDPARENT = -8
Const GWL_STYLE = -16
Const GWL_EXSTYLE = -20
Const GWL_USERDATA = -21
Const GWL_ID = -12
' --------------------------------------------------------------------
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
' Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
' Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowDLL Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
' Private Declare Function FindWindow2 Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cchar As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
' --------------------------------------------------------------------
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
' Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" _
      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long


' --------------------------------------------------------------------
' FUNCTION SET HERE
' --------------------------------------------------------------------
Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
' --------------------------------------------------------------------
Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
' --------------------------------------------------------------------
' --------------------------------------------------------------------




Private Sub Command1_Click()
    Dim s As SHELLEXECUTEINFO
    s.cbSize = LenB(s)
    s.lpFile = "E:\"
    s.nShow = SW_SHOW
    s.fMask = SEE_MASK_INVOKEIDLIST
    s.lpVerb = "properties"
    ShellExecuteEx s
End Sub
' --------------------------------------------------------------------



Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then End
' TOOK A WHILE TO GET THE EVENT DRIVE KEYCODE AND INDEX TO LOOP IN ONE DIRECTION ONLY UP
' Combo1_CLICK WAS MAIN ONE FROM HERE
' PUT BOTH DIRECTION LATER
' --------------------------------------------------------------------------------------
If Combo1.ListIndex = Combo1.ListCount - 1 Then
    If KeyCode = 40 Then
        Combo_KEYCODE_NOT_EVENT = 0
        Combo1.ListIndex = 0
        Exit Sub
    End If
End If
If Combo1.ListIndex = 0 Then
    If KeyCode = 140 Then
'        REM AWAY FOR BOTH DIRECTION LOOP COMBO
'        Combo_KEYCODE_NOT_EVENT = Combo1.ListCount - 1
'        Combo1.ListIndex = Combo1.ListCount - 1
'        Exit Sub
    End If
End If

Call SET_COMBO1_POSITION

If KeyCode <> 13 Then Exit Sub

Counter = Combo1.ListIndex + 1
A1$ = A4$(Counter)
B1$ = B4$(Counter)
C1$ = C4$(Counter)
SetTrueToLoadLast = True
Call Label1_Click(0)

End Sub

Private Sub Combo1_Click()

'Combo1.Text = "List of Last View Folder"

'If Combo1.ListIndex = 0 Then
'Combo1.ListIndex = 1
'End If

If Combo_KEYCODE_NOT_EVENT > -1 And Combo1.ListIndex <> Combo_KEYCODE_NOT_EVENT Then
    XX = Combo_KEYCODE_NOT_EVENT
    Combo_KEYCODE_NOT_EVENT = -1
    Combo1.ListIndex = XX
    Exit Sub
End If
Call SET_COMBO1_POSITION
Call SELECT_COMBO_INDEX
End Sub

Private Sub Combo1_Change()
Combo1.SelLength = 0
Combo1.SelStart = 0
Call SET_COMBO1_POSITION
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
Combo1.SelLength = 0
Combo1.SelStart = 0
Call SET_COMBO1_POSITION
Call SELECT_COMBO_INDEX
End Sub

Private Sub Combo1_Scroll()
Combo1.SelLength = 0
Combo1.SelStart = 0
Call SET_COMBO1_POSITION
Call SELECT_COMBO_INDEX
End Sub

Sub SELECT_COMBO_INDEX()
'Debug.Print Combo1.List(Combo1.ListIndex)
'Debug.Print Combo1.Text

If Combo1.ListIndex <= 0 Then Exit Sub
If OLD_Combo1_ListIndex = Combo1.ListIndex Then Exit Sub

For R1 = 0 To Combo1.ListCount - 1
    XX = Trim(Mid(Combo1.List(R1), InStr(Combo1.List(R1), ".") + 1))
    For Each Control In Label1
        If InStr(Control.Caption + "----", XX + "----") > 0 Then
            Control.BackColor = RGB(127, 127, 200)
            Control.Refresh
        End If
    Next
Next

R1 = Combo1.ListIndex
XX = Trim(Mid(Combo1.List(R1), InStr(Combo1.List(R1), ".") + 1))
For Each Control In Label1
    If InStr(Control.Caption + "----", XX + "----") > 0 Then
        Control.BackColor = RGB(200, 200, 255)
        Control.Refresh
    End If
Next
OLD_Combo1_ListIndex = Combo1.ListIndex
End Sub

Sub SET_COMBO1_POSITION()
' THIS A RUN ONCE THING MAINLY AT BEGIN
' DON'T WORRY ABOUT AUTORESIZE IT ONLY GET RUN RARELY IF CHANGE AMOUNT LEN DIGITS

X_COUNTER = Trim(Str(Combo1.ListIndex) + 1)
If X_COUNTER <= 0 Then X_COUNTER = 1
TT = String(Len(Trim(Str(Combo1.ListCount))), "0")
STRING_COMBO_NUMBER = Format(X_COUNTER, TT) + " of " + Format(Combo1.ListCount, TT)
On Error Resume Next

If Len(STRING_COMBO_NUMBER) <> Len(Lbl_COMBO_NUMBER.Caption) Or DO_ANOTHER_COMBO1_WIDTH = True Then
    Lbl_COMBO_NUMBER.Caption = STRING_COMBO_NUMBER
    
    Lbl_COMBO_NUMBER.AutoSize = True
    Lbl_COMBO_NUMBER.Refresh
    Lbl_COMBO_NUMBER.AutoSize = False
    Lbl_COMBO_NUMBER.Alignment = 2
    Lbl_COMBO_NUMBER.Height = Lbl_Title.Height
    
    Lbl_COMBO_NUMBER.Left = Lbl_Title.Width + 20
    Lbl_COMBO_NUMBER.Width = Lbl_COMBO_NUMBER.Width + 200
    Combo1.Left = Lbl_COMBO_NUMBER.Left + Lbl_COMBO_NUMBER.Width + 20
    ' ----------------------------------------------------
    ' ERROR HAPPEN LINE _COMBO1.WIDTH_ INVAILD PROPERTY VALUE --  RESIZER
    ' ----------------------------------------------------
    Combo1.Width = Form1.Width - Combo1.Left - 100
    If Err.Number > 0 Then
        DO_ANOTHER_COMBO1_WIDTH = True
        XXW = Form1.Width - Combo1.Left - 100
        If XXW < 0 Then XXW = 400
        Combo1.Width = XXW
    Else
        DO_ANOTHER_COMBO1_WIDTH = False
    End If
End If

Lbl_COMBO_NUMBER.Caption = STRING_COMBO_NUMBER
End Sub

Private Sub MNU_GC_PROFILE_Click(Index As Integer)
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    Me.WindowState = vbMinimized
    If Index > 2 Then
        XX = MNU_GC_PROFILE(Index).Caption
        SHELL_LINE_1 = "C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPTCRIPT\VBS 40-RUN EXE EXPLORER LANCHER.VBS"
        SHELL_LINE_2 = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        SHELL_LINE_3 = "--profile-directory=""Profile " + Trim(Str(Index - 1)) + """"
        ' "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 4"
        D_2 = MNU_GC_PROFILE.Item(Index).Caption
        If Index = 1 Then
            SHELL_LINE_3 = "--profile-directory=""Profile" + """"
        End If
        SHELL_LINE_2 = Replace(SHELL_LINE_2, " ", "*")
        SHELL_LINE_3 = Replace(SHELL_LINE_3, " ", "*")
        l = """" + SHELL_LINE_1 + """" + " " + SHELL_LINE_2 + " " + SHELL_LINE_3
        Debug.Print l
        WSHShell.Run l, ShowWindow_2, DontWaitUntilFinished
    End If
    ' LOAD ALL
    If Index = 0 Then
        XX = MNU_GC_PROFILE(Index).Caption
        SHELL_LINE_1 = "C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPTCRIPT\VBS 40-RUN EXE EXPLORER LANCHER.VBS"
        SHELL_LINE_2 = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        For R_1 = 0 To MNU_GC_PROFILE.Count - 1
            R_2 = MNU_GC_PROFILE.Item(R_1).Index
            D_2 = MNU_GC_PROFILE.Item(R_1).Caption
            If R_2 > 0 Then
                SHELL_LINE_3 = "--profile-directory=""Profile " + Trim(Str(R_2)) + """"
            End If
            If R_2 = 0 Then
                SHELL_LINE_3 = "--profile-directory=""Profile" + """"
            End If
            ' "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 4"
            SHELL_LINE_2 = Replace(SHELL_LINE_2, " ", "*")
            SHELL_LINE_3 = Replace(SHELL_LINE_3, " ", "*")
            l = """" + SHELL_LINE_1 + """" + " " + SHELL_LINE_2 + " " + SHELL_LINE_3
            Debug.Print l
            If MNU_GC_PROFILE.Item(R_1).Visible = True Then
                WSHShell.Run l, ShowWindow_2, DontWaitUntilFinished
                Else
                A22 = 2222 ' ---- THIS PART NOT HAPPEN -- TEST DEBUG AR
            End If
        Next
    End If

    ' INDEX 1 -- LOAD ALL NOT 00
    If Index = 1 Then
        XX = MNU_GC_PROFILE(Index).Caption
        SHELL_LINE_1 = "C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPTCRIPT\VBS 40-RUN EXE EXPLORER LANCHER.VBS"
        SHELL_LINE_2 = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        For R_1 = 2 To MNU_GC_PROFILE.Count - 1
            R_2 = MNU_GC_PROFILE.Item(R_1).Index
            D_2 = MNU_GC_PROFILE.Item(R_1).Caption
            
            If R_2 > 2 Then
                SHELL_LINE_3 = "--profile-directory=""Profile " + Trim(Str(R_2)) + """"
            End If
            If R_2 <= 2 Then
                SHELL_LINE_3 = "--profile-directory=""Profile" + """"
            End If
            ' "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 4"
            SHELL_LINE_2 = Replace(SHELL_LINE_2, " ", "*")
            SHELL_LINE_3 = Replace(SHELL_LINE_3, " ", "*")
            l = """" + SHELL_LINE_1 + """" + " " + SHELL_LINE_2 + " " + SHELL_LINE_3
            Debug.Print l
            If MNU_GC_PROFILE.Item(R_1).Visible = True Then
                WSHShell.Run l, ShowWindow_2, DontWaitUntilFinished
                Else
                A22 = 2222 ' ---- THIS PART NOT HAPPEN -- TEST DEBUG AR
            End If
        Next
    End If
    ' INDEX 2 -- LOAD ALL OVER 30 MEGABYTE
    If Index = 2 Then
        XX = MNU_GC_PROFILE(Index).Caption
        SHELL_LINE_1 = "C:\SCRIPTER\SCRIPTER CODE -- VBSCRIPTCRIPT\VBS 40-RUN EXE EXPLORER LANCHER.VBS"
        SHELL_LINE_2 = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        For R_1 = 3 To MNU_GC_PROFILE.Count - 1
            R_2 = MNU_GC_PROFILE.Item(R_1).Index - 2
            D_2 = MNU_GC_PROFILE.Item(R_1).Caption
            
            If R_1 = 3 Then
                PATHFIND = "C:\Users\MATT 01\AppData\Local\Google\Chrome\User Data\PROFILE"
            End If
            If R_1 > 3 Then
                PATHFIND = "C:\Users\MATT 01\AppData\Local\Google\Chrome\User Data\PROFILE " + Format(R_2)
            End If
        
            If FSO.FolderExists(PATHFIND) Then
            Set F = FSO.GetFolder(PATHFIND)
            If F.Size > 30 * 1024 ^ 2 Then
                ' 332 051 185
                ' DEBUG.PRINT 30 * 1024^2 ' ---- 30 MEG
                If R_2 > 2 Then
                    SHELL_LINE_3 = "--profile-directory=""Profile " + Trim(Str(R_2)) + """"
                End If
                If R_2 <= 2 Then
                    SHELL_LINE_3 = "--profile-directory=""Profile" + """"
                End If
                ' "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 4"
                SHELL_LINE_2 = Replace(SHELL_LINE_2, " ", "*")
                SHELL_LINE_3 = Replace(SHELL_LINE_3, " ", "*")
                l = """" + SHELL_LINE_1 + """" + " " + SHELL_LINE_2 + " " + SHELL_LINE_3
                Debug.Print l
                If MNU_GC_PROFILE.Item(R_1).Visible = True Then
                    WSHShell.Run l, ShowWindow_2, DontWaitUntilFinished
                    Else
                    A22 = 2222 ' ---- THIS PART NOT HAPPEN -- TEST DEBUG AR
                End If
            End If
            End If
        Next
    End If

End
Set WSHShell = Nothing
End Sub

Sub MNU_GC_PROFILE_SET_ON()
i = -1
i = i + 1: MNU_GC_PROFILE.Item(i).Caption = "G CHROME _ LOAD ALL"
i = i + 1: MNU_GC_PROFILE.Item(i).Caption = "G CHROME _ LOAD ALL NOT 00"
i = i + 1: MNU_GC_PROFILE.Item(i).Caption = "G CHROME _ LOAD ALL NOT UNDER 30 MEGABYTE"
I_SET = i + 1
Dim NUMERIC_VAR_1
Dim NUMERIC_VAR_2
NUMERIC_VAR_2 = i
For i = i To MNU_GC_PROFILE.Count + I_SET
    ' C:\Users\MATT 01\AppData\Local\Google\Chrome\User Data
    NUMERIC_VAR_1 = i - I_SET + 1
    If NUMERIC_VAR_1 = 0 Then
        PATHFIND = "C:\Users\MATT 01\AppData\Local\Google\Chrome\User Data\PROFILE"
    End If
    If NUMERIC_VAR_1 > 0 Then
        PATHFIND = "C:\Users\MATT 01\AppData\Local\Google\Chrome\User Data\PROFILE " + Format(NUMERIC_VAR_1, "0")
    End If
    If FSO.FolderExists(PATHFIND) Then
        Set F = FSO.GetFolder(PATHFIND)
        A_SIZE = Format(F.Size / 1024 ^ 2, "##00")
    End If
    If FSO.FolderExists(PATHFIND) Then
        NUMERIC_VAR_2 = NUMERIC_VAR_2 + 1
        If NUMERIC_VAR_1 = 0 Then
            MNU_GC_PROFILE.Item(NUMERIC_VAR_2).Caption = "G CHROME _ PROFILE" + " _____ " + A_SIZE + " MB"
        Else
            MNU_GC_PROFILE.Item(NUMERIC_VAR_2).Caption = "G CHROME _ PROFILE " + Format(NUMERIC_VAR_1, "00") + " __ " + A_SIZE + " MB"
        End If
    Else
        ' MNU_GC_PROFILE.Item(NUMERIC_VAR_2).Visible = False
    End If
Next

For i = NUMERIC_VAR_2 + 1 To MNU_GC_PROFILE.Count - 1
    MNU_GC_PROFILE.Item(i).Visible = False
Next

End Sub



Public Function GetShortName(ByVal sLongFileName As String) As String
' --> (All this modules code) Obtained from : -
' ---> Microsoft's/MSDN's code
' ->  http://support.microsoft.com/kb/q175512/
' ---> The original comments were by them :

       Dim lRetVal As Long, sShortPathName As String, iLen As Integer
       'Set up buffer area for API function call return
       sShortPathName = Space(255)
       iLen = Len(sShortPathName)

       'Call the function
       lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
       'Strip away unwanted characters.
       GetShortName = Left(sShortPathName, lRetVal)
End Function



' SCRIPTOR ------------------------------
Private Sub MNU_LOAD_ALL_SCRIPTOR_Click()

    ' ----------------------------------------------------------
    ' ----------------------------------------------------------
    ' IF GOT PROBLM SOME WON'T LOAD IN
    ' AND THEN SHUT DOWN NOTEPAD++ AND RLOAD IT
    ' AS PROFILE INFO OF WHAT FILES LOADD ARE NOT SAVED UNTIL IT CLOSING
    ' ----------------------------------------------------------
    ' ----------------------------------------------------------
    
    
    
    ' ----------------------------------------------------------
    ' ----------------------------------------------------------

    Me.WindowState = vbMinimized
    NP = ""
    If NP = "" Then
        NP2 = "C:\Program Files\Notepad++\notepad++.exe"
        NP = Dir(NP2)
    End If
    If NP = "" Then
        NP2 = "C:\Program Files (X86)\Notepad++\notepad++.exe"
        NP = Dir(NP2)
    End If
    If NP = "" Then
        MsgBox "NOTEPAD++ NOT FOUND -- HERE " + vbCrLf + vbCrLf + NP1 + vbCrLf + vbCrLf + NP2
    End If
    
    'Simple Find Date
    Dim Filespec, Adate1
    Set FS = CreateObject("Scripting.FileSystemObject")
    
    ' ------------------------------------------------
    ' READ FILE -- NOTEPADD++ -- FIND WHAT NOT HAVE DO
    ' ------------------------------------------------
    ' GetComputerName
    ' GetComputerName + "-" + GetUserName
    ' C:\Users\MATT 04\AppData\Roaming\Notepad++\session.xml
    ' C:\Users\"+ GetUserName+"\AppData\Roaming\Notepad++\session.xml
    
    NOTEPAD_FILE = "C:\Users\" + GetUserName + "\AppData\Roaming\Notepad++\session.xml"
    Dim NOTEPAD_STRING As String
    If FSO.FileExists(NOTEPAD_FILE) = True Then
        FR1 = FreeFile
        Open NOTEPAD_FILE For Binary As #FR1
            NOTEPAD_STRING = Space$(LOF(FR1))
            Get #FR1, 1, NOTEPAD_STRING
        Close #F1
        ' NOTEPAD_STRING = UCase(NOTEPAD_STRING)
    End If
    
'    Dim PATH_F()
'    ReDim PATH_F(20)
'    i = 0
'    i = i + 1: PATH_F(i) = "C:\SCRIPTER"
'    i = i + 1: PATH_F(i) = "C:\PStart\Progs\0_Nirsoft_Package\NirSoft"
'    ReDim Preserve PATH_F(i)
    AA1$ = "C:\Users\" + GetUserName + "\AppData\Roaming\Notepad++\session.xml"
    If Dir(AA1$) <> "" Then
        FR4 = FreeFile
        Open AA1$ For Binary As #FR4
            STRING_SESSION_XML = Input(LOF(FR4), FR4)
        Close #FR4
    End If
    
    
    Dim PATH_FILE()
    ReDim PATH_FILE(20)
    i = 0
    i = i + 1: PATH_FILE(i) = "D:\BT CLOUD SYNC OLD\CALLERID\2Double Trouble.txt"
    i = i + 1: PATH_FILE(i) = "C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\Multiple_Thread Port Scanner 02 CON\NETWORK_COMPUTER_NAME.txt"
    i = i + 1: PATH_FILE(i) = "C:\Menu.lst"
    i = i + 1: PATH_FILE(i) = "C:\Menu_oppo_system.lst"
    i = i + 1: PATH_FILE(i) = "C:\Menu_win_oppo_system_installs.lst"
    i = i + 1: PATH_FILE(i) = "C:\Menu_z_Demo.lst"
    i = i + 1: PATH_FILE(i) = "C:\NETWORK_COMPUTER_NAME.txt"
    i = i + 1: PATH_FILE(i) = "C:\BAT 01-BCDEDIT_.BAT"
    i = i + 1: PATH_FILE(i) = "C:\BAT 01-BCDEDIT SET {BOOTMGR} DISPLAYBOOTMENU YES.BAT"
'    i = i + 1: PATH_FILE(i) = "C:\28 SendSMTP_REBOOT_BATCH.BAT.lnk"
    'i = i + 1: PATH_FILE(i) = ""
    'i = i + 1: PATH_FILE(i) = ""
    
    ReDim Preserve PATH_FILE(i)
    For R_4 = 1 To UBound(PATH_FILE)
        If InStr(STRING_SESSION_XML, PATH_FILE(R_4)) = 0 Then
            DoEvents
            Shell NP2 + " """ + PATH_FILE(R_4) + """"
        End If
    Next
    
    'D:\HARDWARE\HARDWARE 2026-04-06 - £15.08 Revo Uninstaller Pro 5 - 2 years FOR MSI\HARDWARE 2026-04-06 - £15.08 Revo Uninstaller Pro 5 - 2 years FOR MSI.txt
    ScanPath.ListView1.ListItems.Clear
    ScanPath.cboMask.Text = "*.TXT"
    ScanPath.txtPath.Text = "D:\HARDWARE"
    Call ScanPath.cmdScan_Click
    For r = ScanPath.ListView1.ListItems.Count To 1 Step -1
        A1$ = ScanPath.ListView1.ListItems.Item(r).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(r)
        FLAG_DONE = False
        '----------------------------------
        If InStr(B1$, "\_gsdata_") > 0 Then
            ScanPath.ListView1.ListItems.Remove (r)
            FLAG_DONE = True
        End If
        '----------------------------------
        If Mid(B1$, 1, 11) <> "HARDWARE 20" And FLAG_DONE = False Then
            ScanPath.ListView1.ListItems.Remove (r)
            FLAG_DONE = True
        End If
        '----------------------------------
        If InStr(B1$, "ITEM LEARN") > 0 And FLAG_DONE = False Then
            ScanPath.ListView1.ListItems.Remove (r)
            FLAG_DONE = True
        End If
        If FLAG_DONE = False Then
            Filespec = A1$ + B1$
            'SHORT
            SHORT_FILESPEC = GetShortName(Filespec)
            On Error Resume Next
            Set F = FS.GetFile(Filespec)
            SHORT_FILESPEC = ""
            If Err.Number > 0 Then
                SHORT_FILESPEC = GetShortName(Filespec)
                Err.Clear
                Set F = FS.GetFile(SHORT_FILESPEC)
                If Err.Number > 0 Then SHORT_FILESPEC = "ERROR"
            End If
            On Error GoTo 0
            
            Adate1 = F.datelastmodified
            If FSO.FileExists(A1$ + B1$) = True And SHORT_FILESPEC <> "ERROR" Then
                If Adate1 < Now - 40 Then
                    ScanPath.ListView1.ListItems.Remove (r)
                    FLAG_DONE = True
                End If
            End If
        End If
        If FLAG_DONE = False Then
            If FSO.FileExists(A1$ + B1$) = False Or SHORT_FILESPEC = "ERROR" Then
                ScanPath.ListView1.ListItems.Remove (r)
            End If
        End If

    Next

    For r = 1 To ScanPath.ListView1.ListItems.Count
        A1$ = ScanPath.ListView1.ListItems.Item(r).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(r)
        If InStr(STRING_SESSION_XML, B1$) = 0 Then
            DoEvents
            Shell NP2 + " """ + A1$ + B1$ + """"
        End If
    Next


    ' FILTER -- THING NOT HAVE
    Dim FILTER_01()
    ReDim FILTER_01(80)
    i = 0
    i = i + 1: FILTER_01(i) = "1-ASUS-X5DIJ"
    i = i + 1: FILTER_01(i) = "2-ASUS-EEE"
    i = i + 1: FILTER_01(i) = "3-LINDA-PC"
    i = i + 1: FILTER_01(i) = "4-ASUS-GL522VW"
    i = i + 1: FILTER_01(i) = "5-ASUS-P2520LA"
    i = i + 1: FILTER_01(i) = "7-ASUS-GL522VW"
    i = i + 1: FILTER_01(i) = "8-MSI-GP62M-7RD"
    i = i + 1: FILTER_01(i) = "9-ASUS-G815LM"
    i = i + 1: FILTER_01(i) = "#NFS"
    i = i + 1: FILTER_01(i) = "\_GSDATA_"
    i = i + 1: FILTER_01(i) = "\SCRIPTER CODE -- VB 01 VB6"
    i = i + 1: FILTER_01(i) = "\SCRIPTER CODE -- VB 03 VB2008"
    i = i + 1: FILTER_01(i) = "\SCRIPTER CODE -- VB 04 VB2017"
    i = i + 1: FILTER_01(i) = "\SCRIPTER CODE -- VB 05 VB2026"
    i = i + 1: FILTER_01(i) = "VBSCRIPT 68"
    i = i + 1: FILTER_01(i) = "COPY_CAMERA_PHOTO_IMAGE_POINTER"
    i = i + 1: FILTER_01(i) = "MATT.LAN.COM"
    i = i + 1: FILTER_01(i) = "LAPTOP BATTERY HEALTH INFO"
    i = i + 1: FILTER_01(i) = "MP3 AUDIO"
    i = i + 1: FILTER_01(i) = "SCRIPT 01"
    i = i + 1: FILTER_01(i) = "SCRIPT 02"
    i = i + 1: FILTER_01(i) = "SCRIPT 08"
    i = i + 1: FILTER_01(i) = "SCRIPTER -- "
    i = i + 1: FILTER_01(i) = " (2)."
    i = i + 1: FILTER_01(i) = "SCRIPTER CODE -- REG KEY SETTER"
    i = i + 1: FILTER_01(i) = "SCRIPTER CODE -- USERSTYLES.ORG STYLISH"
    i = i + 1: FILTER_01(i) = "SYNC_FOLDER_"
    i = i + 1: FILTER_01(i) = "USB_PEN_DRIVE_DISKMOD"
    i = i + 1: FILTER_01(i) = "UTILS"
    i = i + 1: FILTER_01(i) = "Z_ROOT_COMPARE"
    i = i + 1: FILTER_01(i) = "AHK - 7 PLUS"
    i = i + 1: FILTER_01(i) = "AHK SPARE PART CODE"
    i = i + 1: FILTER_01(i) = "JOYPAD_WORK"
    i = i + 1: FILTER_01(i) = "NOTEPAD++ AUTOHOTKEYS LANGUAGE PLUG-IN"
    i = i + 1: FILTER_01(i) = "SCRIPTER CODE -- AHK_ NIFTY_WINDOWS"
    i = i + 1: FILTER_01(i) = "SCRIPTER CODE -- VB6"
    i = i + 1: FILTER_01(i) = "SCRIPTER CODE -- VBNET 2008"
    i = i + 1: FILTER_01(i) = "SCRIPTER CODE -- GRUB4DOS"
    i = i + 1: FILTER_01(i) = "\SCRIPTER CODE -- VB 02 VBSCRIPT\VBS 18-FFMPEG"
    i = i + 1: FILTER_01(i) = "\SCRIPTER CODE -- VB 02 VBSCRIPT\VBS 12-PINTO10"
    i = i + 1: FILTER_01(i) = "\SCRIPTER CODE -- VB 02 VBSCRIPT\VBSEDIT"
    i = i + 1: FILTER_01(i) = "\REG"
    i = i + 1: FILTER_01(i) = "\CONFIG"
    i = i + 1: FILTER_01(i) = "\MP3 AUDIO"
    i = i + 1: FILTER_01(i) = "\UTILS"
    
    i = i + 1: FILTER_01(i) = "TEXT.TEXT.TXT"
    
    i = i + 1: FILTER_01(i) = "\SCRIPTER CODE -- AUTOHOTKEY\AUDIO SET"
    
    ' i = i + 1: FILTER_01(i) = "\SCRIPTER CODE -- VB 02 VBSCRIPT\VBS 10-VICEVERSA _ SHELL FOLDERING__\"
    i = i + 1: FILTER_01(i) = "\SCRIPTER CODE -- VB 02 VBSCRIPT\VBS 12-PINTO10"
    i = i + 1: FILTER_01(i) = "\SCRIPTER CODE -- VB 02 VBSCRIPT\VBS 18-FFMPEG"
    i = i + 1: FILTER_01(i) = "\SCRIPTER CODE -- VB 02 VBSCRIPT\VBSEDIT"
    i = i + 1: FILTER_01(i) = "RETEKESS_V115"
    i = i + 1: FILTER_01(i) = "RETEKESSMTB"
    i = i + 1: FILTER_01(i) = ".BAK"
    
    'i = i + 1: FILTER_01(i) = ""
    'i = i + 1: FILTER_01(i) = ""
    
    i = i + 1: FILTER_01(i) = "TEXT_ADBLOCK_ARC_"
    i = i + 1: FILTER_01(i) = "FBP-SETT" ' FBP-Settings-14-Jan-2020.txt
    ReDim Preserve FILTER_01(i)

    Dim FILTER_PATH_01()
    ReDim FILTER_PATH_01(40)
    i = 0
    i = i + 1: FILTER_PATH_01(i) = "C:\SCRIPTER\SCRIPTER CODE -- BAT\"
    i = i + 1: FILTER_PATH_01(i) = "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\"
    i = i + 1: FILTER_PATH_01(i) = "C:\SCRIPTER\SCRIPTER CODE -- VB 02 VBSCRIPT\"
    i = i + 1: FILTER_PATH_01(i) = "C:\SCRIPTER\SCRIPTER CODE -- POWERSHELL\"
    i = i + 1: FILTER_PATH_01(i) = "C:\SCRIPTER\SCRIPTER CODE -- GITHUB\"
    
    ReDim Preserve FILTER_PATH_01(i)


'---------------------------------------------------------

    ScanPath.ListView1.ListItems.Clear
    ScanPath.cboMask.Text = "*.*"
    ScanPath.txtPath.Text = "C:\SCRIPTER"
    Call ScanPath.cmdScan_Click
    For r = ScanPath.ListView1.ListItems.Count To 1 Step -1
        A1$ = ScanPath.ListView1.ListItems.Item(r).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(r)
        FLAG_DONE = False
        '----------------------------------
        If InStr(B1$, "_gsdata_") > 0 Then
            ScanPath.ListView1.ListItems.Remove (r)
            FLAG_DONE = True
        End If
        '----------------------------------
        '----------------------------------
        
        If FLAG_DONE = False Then
            FILE1_EX = UCase(A1$ + B1$)
            If InStr(B1$, ".") > 0 Then
                FILE1_EX = UCase(Mid(B1$, InStrRev(B1$, ".")))
            End If
            If InStr(B1$, ".") = 0 Then
                ScanPath.ListView1.ListItems.Remove (r)
                FLAG_DONE = True
            End If
            ' PRIORITY CHECK 01
            If InStr(".VBS .AHK .TXT .BAS .BAT", FILE1_EX) > 0 Then
                If InStr(" .TXT", FILE1_EX) > 0 Then
                    If InStr(UCase(B1$), "TREESIZE") = 1 And FLAG_DONE = False Then
                        ScanPath.ListView1.ListItems.Remove (r)
                        FLAG_DONE = True
                    End If
                End If
            End If
            ' .LST .DIZ .REG .PS1 .CMD .LOG
            ' .LST FOR MENU GRUB4DOS
            If InStr(".VBS .AHK .TXT .BAS .BAT .LST .PS1 .MD", FILE1_EX) = 0 Then
                If FLAG_DONE = False Then
                    ScanPath.ListView1.ListItems.Remove (r)
                    FLAG_DONE = True
                End If
            End If
            
            
            ' PRIORITY CHECK 02 NOT ALLOW
            For R_5 = 1 To UBound(FILTER_01)
                If InStr(UCase(A1$ + B1$), UCase(FILTER_01(R_5))) > 0 And FLAG_DONE = False Then
                    ScanPath.ListView1.ListItems.Remove (r)
                    FLAG_DONE = True
                End If
            Next
        

            ' PRIORITY CHECK 02 NOT ALLOW
            For R_5 = 1 To UBound(FILTER_PATH_01)
                If FLAG_DONE = False Then
                    If InStr(UCase(A1$ + B1$), UCase(FILTER_PATH_01(R_5))) > 0 Then
                        If Len(A1$) > Len(FILTER_PATH_01(R_5)) Then
                            ScanPath.ListView1.ListItems.Remove (r)
                            FLAG_DONE = True
                        End If
                    End If
                End If
            Next
        
        
        End If
        
        
        
        If FLAG_DONE = False Then
            Filespec = A1$ + B1$
            'SHORT
            SHORT_FILESPEC = GetShortName(Filespec)
            On Error Resume Next
            Set F = FS.GetFile(Filespec)
            SHORT_FILESPEC = ""
            If Err.Number > 0 Then
                SHORT_FILESPEC = GetShortName(Filespec)
                Err.Clear
                Set F = FS.GetFile(SHORT_FILESPEC)
                If Err.Number > 0 Then SHORT_FILESPEC = "ERROR"
            End If
            On Error GoTo 0
            
        End If
        If FLAG_DONE = False Then
            If FSO.FileExists(A1$ + B1$) = False Or SHORT_FILESPEC = "ERROR" Then
                ScanPath.ListView1.ListItems.Remove (r)
            End If
        End If

    Next
    For r = 1 To ScanPath.ListView1.ListItems.Count
        A1$ = ScanPath.ListView1.ListItems.Item(r).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(r)
        If InStr(STRING_SESSION_XML, B1$) = 0 Then
            DoEvents
            Shell NP2 + " """ + A1$ + B1$ + """"
        End If
    Next

'----------------------------------------------






    
'    For R_4 = 1 To UBound(PATH_F)
'        Dir1.Path = PATH_F(R_4)
'        For R_2 = 1 To Dir1.ListCount - 1
'            File1.Path = Dir1.List(R_2)
'            File1.Pattern = "*.*"
'            For R_1 = 0 To File1.ListCount - 1
'                FILE1_EX = UCase(File1.List(R_1))
'                If InStr(File1.List(R_1), ".") > 0 Then
'                    FILE1_EX = UCase(Mid(File1.List(R_1), InStrRev(File1.List(R_1), ".")))
'                End If
'                SET_GO = False
'                ' PRIORITY CHECK 01
'                If InStr(".VBS .AHK .TXT .BAS .BAT", FILE1_EX) > 0 Then
'                    SET_GO = True
'                    If InStr(" .BAT", FILE1_EX) > 0 Then
'                        If InStr(Dir1.Path, "SCRIPTER CODE -- BAT") = 0 Then SET_GO = False
'                    End If
'                    If InStr(" .TXT", FILE1_EX) > 0 Then
'                        If InStr(UCase(File1.List(R_1)), "TREESIZE") = 1 Then
'                            SET_GO = False
'                        End If
'                    End If
'                End If
'                ' PRIORITY CHECK 02 NOT ALLOW
'                For R_5 = 1 To UBound(FILTER_01)
'                    If InStr(File1.Path + "\" + File1.List(R_1), FILTER_01(R_5)) > 0 Then
'                        SET_GO = False
'                    End If
'                Next
'
'
'                Dim ARRAY_04(2, 2)
'                i = 0
'                ' "Autokey -- 51-Display in real time the active window's control list.ahk"
'                i = i + 1: ARRAY_04(i, 1) = "'": ARRAY_04(i, 2) = "&apos;"
'                i = i + 1: ARRAY_04(i, 1) = "&": ARRAY_04(i, 2) = "&amp;"
'                AMP_LINE = ""
'                For R3 = 1 To UBound(ARRAY_04)
'                    AMP_LINE = AMP_LINE + ARRAY_04(R3, 2)
'                Next
'                If SET_GO = True Then
'                If InStr(NOTEPAD_STRING, File1.List(R_1)) = 0 Then
'                FILE_CHECK_LINE = File1.List(R_1)
'                For R3 = 1 To UBound(ARRAY_04)
'                    POS_INDEX = 1
'                    Do
'                        POS_OF_AND_SIGN = InStr(POS_INDEX, FILE_CHECK_LINE, ARRAY_04(R3, 1))
'                        If POS_OF_AND_SIGN = 0 Then Exit Do
'                        If POS_OF_AND_SIGN > 0 Then
'                            GET_BIT_OF = Mid(FILE_CHECK_LINE, POS_OF_AND_SIGN, Len(ARRAY_04(R3, 2)))
'                            If InStr(AMP_LINE, GET_BIT_OF) = 0 Then
'                                FILE_CHECK_LINE = Replace(FILE_CHECK_LINE, ARRAY_04(R3, 1), ARRAY_04(R3, 2))
'                            End If
'                        End If
'                        POS_INDEX = POS_OF_AND_SIGN + 1
'                    Loop Until 1 = 2
'                Next
'                End If
'                End If
'
'                ' REPLY WILL BE FALSE WHEN FIND RESULT NOT TO BOTHER
'                If InStr(NOTEPAD_STRING, FILE_CHECK_LINE) > 0 Then
'                    SET_GO = False
'                End If
'
'                ' DO SOMETHING
'                If SET_GO = True Then
'                    DoEvents
'                    Shell NP2 + " """ + File1.Path + "\" + File1.List(R_1) + """"
'                    ' If R_1 > 50 Then Stop
'                End If
'            Next
'        Next
'    Next
'
    End
    
End Sub

Private Sub MNU_LOAD_CHROME_PROFILE_FOLDER_Click()
    Me.Hide
    
    If FSO.FolderExists("C:\Users\MATT 01\AppData\Local\Google\Chrome\User Data\Profile") = True Then
        Shell "EXPLORER /SELECT, C:\Users\MATT 01\AppData\Local\Google\Chrome\User Data\Profile", vbMaximizedFocus
        End
    End If
    If FSO.FolderExists("C:\Users\MATT 01\AppData\Local\Google\Chrome\User Data\Profile 1") = True Then
        Shell "EXPLORER /SELECT, C:\Users\MATT 01\AppData\Local\Google\Chrome\User Data\Profile 1", vbMaximizedFocus
        End
    End If
    If FSO.FolderExists("C:\Users\MATT 01\AppData\Local\Google\Chrome\User Data\Profile 2") = True Then
        Shell "EXPLORER /SELECT, C:\Users\MATT 01\AppData\Local\Google\Chrome\User Data\Profile 2", vbMaximizedFocus
        End
    End If
    If FSO.FolderExists("C:\Users\MATT 01\AppData\Local\Google\Chrome\User Data\Profile 3") = True Then
        Shell "EXPLORER /SELECT, C:\Users\MATT 01\AppData\Local\Google\Chrome\User Data\Profile 3", vbMaximizedFocus
        End
    End If
    
    Shell "explorer C:\Users\MATT 01\AppData\Local\Google\Chrome\User Data", vbMaximizedFocus
    End
End Sub

Private Sub MNU_NET_2_GO_Click()
    MNU_NET_2_GO_VALUE = Not MNU_NET_2_GO_VALUE
    If MNU_NET_2_GO_VALUE = -1 Then MNU_NET_2_GO.Caption = "[__ NET 2 GO = GO __]"
    If MNU_NET_2_GO_VALUE = 0 Then MNU_NET_2_GO.Caption = "[__ NET 2 GO __]"
End Sub

Private Sub MNU_OPEN_PROPERTIE_FOR_ALL_DRIVE_Click()
    Me.WindowState = vbMinimized
    
    ' -----------------------------------------------
    ' HERE MIGHT WORK BUT
    ' CHANGE TO HERE AS
    ' RESULT NOT DESIRE
    ' WHEN APP CLOSE THAT CALL OPEN PROPERTY AND THEN
    ' ALSO CLOSE OPEN PROPERTY AFTER
    ' -----------------------------------------------
    
    Dim TT, TG, Y1 ' ---- NOT REQUIRE INFO OTHER EXAMPLE
    Dim Filename
    Dim s As SHELLEXECUTEINFO
    On Error Resume Next
    For r = 3 To 25
        Err.Clear
        Set z = FSO.GetDrive(FSO.GetDriveName(FSO.GetAbsolutePathName(Chr$(r + 64) + ":")))
        Select Case z.DriveType
            Case 0: T = "Unknown"
            Case 1: T = "Removable"
            Case 2: T = "Fixed"
            Case 3: T = "Network"
            Case 4: T = "CD-ROM"
            Case 5: T = "RAM Disk"
        End Select
        If Err.Number = 0 Then
            TT = z.DriveLetter + ":" + "- Type - " + T + " --- Serial - " + Hex$(z.SerialNumber) + " - Vol.Name - " + z.VolumeName + vbCrLf
            TG = TG + 1
            Y1 = Y1 + TT
            Filename = z.DriveLetter + ":\" ' + z.VolumeName
            ' Shell "D:\VB6\VB-NT\00_BEST_VB_01\Shell Explorer Loader\SHELLEXECUTEEX_OPEN_PROPERTY.EXE" + " " + FileName
            
            lHnd = GetDesktopWindow()
            s.cbSize = LenB(s)
            s.hWnd = lHnd
            s.lpFile = Filename
            s.nShow = SW_SHOW
            s.fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST
            s.lpVerb = "properties"
            s.lpParameters = "Security"
            ShellExecuteEx s
            
            lHnd = GetDesktopWindow()
            SetParent s.hWnd, lHnd
            
                    ' ------------------------------------------------
        ' A WORK AROUND METHOD
        ' REALLY I WOULD GET A HWND TO PID
        ' PROBLEM TO SOLVE
        ' IS WHEN CALLING PROGRAM QUIT SO DO REQUEST TO RUN ITEM
        ' MY WORK AROUND
        ' IS TRANSFER THE HWND TO ANOTHER APP LIKE DESKTOP
        ' AND WORK EASY
        ' THE SEI ABLE SET HWND
        ' AND EVEN THOUGH SET IT TO DESKTOP
        ' IT NOT REALLY CONNECTED THAT WAY
        ' OBVIOUS DESKTOP IS ONE OF IT OWN
        ' AND SEI RETURN IT OWN ONE ONCE MADE
        ' NOT TO WORRY
        ' HERE IS SOLUTION
        ' ------------------------------------------------
        ' SetParent SEI.hWnd, lHnd
    
        ' SET PARENT SOLUTION CAME FROM
        ' ------------------------------------------------
        ' How to get Process ID from hwnd?-VBForums
        ' http://www.vbforums.com/showthread.php?765305-How-to-get-Process-ID-from-hwnd
        ' ------------------------------------------------

        End If
    Next
    
    Exit Sub
    Unload Me
End Sub

Private Sub MNU_OPEN_PROPERTY_ALL_FOLDER_C_HDD_Click()
    Me.WindowState = vbMinimized
    
    Dim Filename
    Dim s As SHELLEXECUTEINFO
    
    Err.Clear
    Set z = FSO.GetDrive(FSO.GetDriveName(FSO.GetAbsolutePathName("C:")))
    Select Case z.DriveType
        Case 0: T = "Unknown"
        Case 1: T = "Removable"
        Case 2: T = "Fixed"
        Case 3: T = "Network"
        Case 4: T = "CD-ROM"
        Case 5: T = "RAM Disk"
    End Select
    If Err.Number > 0 Then Exit Sub

    Dir1.Path = "C:\"
    
    For r = 0 To Dir1.ListCount - 1
        XGO = 0
        XR0 = UCase(Dir1.List(r))
        If InStr(XR0, UCase("BOOT")) > 0 Then XGO = 1
        If InStr(XR0, UCase("BOOT_EFI_7G-BAK1")) > 0 Then XGO = 1
        If InStr(XR0, UCase("BOOT_EFI_7G-BAK2")) > 0 Then XGO = 1
        If InStr(XR0, UCase("DOCUMENTS AND SETTINGS")) > 0 Then XGO = 1
        If InStr(XR0, UCase("PROGRAM FILES")) > 0 Then XGO = 1
        If InStr(XR0, UCase("PROGRAM FILES (X86)")) > 0 Then XGO = 1
        If InStr(XR0, UCase("PROGRAMDATA")) > 0 Then XGO = 1
        If InStr(XR0, UCase("RECOVERY")) > 0 Then XGO = 1
        If InStr(XR0, UCase("SYSTEM VOLUME INFORMATION")) > 0 Then XGO = 1
        If InStr(XR0, UCase("USERS")) > 0 Then XGO = 1
        If InStr(XR0, UCase("WINDOWS")) > 0 Then XGO = 1
                
        If XGO = 0 Then
            Filename = Dir1.List(r)
            lHnd = GetDesktopWindow()
            s.cbSize = LenB(s)
            s.hWnd = lHnd
            s.lpFile = Filename
            s.nShow = SW_SHOW
            s.fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST
            s.lpVerb = "properties"
            s.lpParameters = "Security"
            ShellExecuteEx s
            
            lHnd = GetDesktopWindow()
            SetParent s.hWnd, lHnd
        End If
    Next

End Sub

Private Sub MNU_SPACE_EMPTY_Click()


Me.WindowState = vbMinimized

Shell "EXPLORER SHELL:MYCOMPUTERFOLDER", vbMaximizedFocus
Shell "EXPLORER SHELL:RECYCLEBINFOLDER", vbMaximizedFocus
Shell "C:\Program Files\CCleaner\CCleaner64.exe", vbMaximizedFocus

Shell "C:\Program Files (x86)\Duplicate Cleaner Pro\DuplicateCleaner.exe", vbMaximizedFocus
Shell "C:\Program Files (x86)\JAM Software\TreeSize Free\TreeSizeFree.exe C:\", vbMaximizedFocus
Shell "C:\Program Files (x86)\JAM Software\TreeSize Free\TreeSizeFree.exe D:\", vbMaximizedFocus

Dim WSHShell

II = GetComputerName
If II = "1-ASUS-X5DIJ" Then II = "1X"
If II = "2-ASUS-EEE" Then II = "2E"
If II = "3-LINDA-PC" Then II = "2L"
If II = "4-ASUS-GL522VW" Then II = "4G"
If II = "5-ASUS-P2520LA" Then II = "5P"
If II = "7-ASUS-GL522VW" Then II = "7G"
If II = "8-MSI-GP62M-7RD" Then II = "8M"
If II = "9-ASUS-G815LM" Then II = "9G"


ReDim EXE_PATH(100)
i = 0
i = i + 1: EXE_PATH(i) = "C:\0 00 LINK SET QUICKER MOVER\SavedCriteria __ FILE LOCATOR __ GOODSYNC _GSDATA_C_.srf"
i = i + 1: EXE_PATH(i) = "C:\0 00 LINK SET QUICKER MOVER\SavedCriteria __ FILE LOCATOR __ GOODSYNC _GSDATA_D_.srf"
If II = "4G" Then
    i = i + 1: EXE_PATH(i) = "C:\0 00 LINK SET QUICKER MOVER\SAVEDCRITERIA__ALL NETWORK DRIVE FROM C TO E FROM NETWORK FROM 4G LOCAL INCLUDER ON.SRF"
End If
If II = "7G" Then
i = i + 1: EXE_PATH(i) = "C:\0 00 LINK SET QUICKER MOVER\SAVEDCRITERIA__FILE LOCATOR __ GOODSYNC _GSDATA_7G_NETWORK___.SRF"
End If
If II = "4G" Then
i = i + 1: EXE_PATH(i) = "C:\0 00 LINK SET QUICKER MOVER\SAVEDCRITERIA__FILE LOCATOR __ GOODSYNC _GSDATA_D__4G_.SRF"
End If
i = i + 1: EXE_PATH(i) = "C:\0 00 LINK SET QUICKER MOVER\SAVEDCRITERIA__FILE LOCATOR__00__ GOODSYNC_ALL_LOCAL_DRIVE_GSDATA_-_SAVED_-_HISTORY_.SRF"
i = i + 1: EXE_PATH(i) = "C:\0 00 LINK SET QUICKER MOVER\SAVEDCRITERIA__FILE LOCATOR__02__ 0 00 ART LOGGERS_# APP AND SCREEN AUTO_.SRF"
i = i + 1: EXE_PATH(i) = "C:\0 00 LINK SET QUICKER MOVER\SAVEDCRITERIA__FILE LOCATOR__02__ TEMPORARY SET.SRF"
ReDim Preserve EXE_PATH(i)
For R1 = 1 To UBound(EXE_PATH)
    Set WSHShell = CreateObject("WScript.Shell")
        WSHShell.Run """" + EXE_PATH(R1) + """"
    Set WSHShell = Nothing
Next

End

End Sub

Private Sub Timer_FORM_RESIZE_Timer()
    If X_Y_DONE_ONCE <> True Then
        Call Form_Resize
    Else
        Timer_FORM_RESIZE.Enabled = False
    End If
End Sub

Private Sub Form_Resize()
    If Form1_Height = 0 Then Exit Sub
    If Form1_Width = 0 Then Exit Sub
    
    A = Form1.Width + Form1.Height + Form1.Top + Form1.Left + Menu_Height + Screen.Width
    B = Form1_Width + Form1_Height + OLD_SCREEN_WIDTH + OLD_MENUBAR
    If A = B Then Exit Sub
    OLD_MENUBAR = Menu_Height
    OLD_SCREEN_WIDTH = Screen.Width
    
    On Error Resume Next
    Form1.Width = Form1_Width
    Form1.Height = Form1_Height + Menu_Height
    'Center Screen
    If X_Y_DONE_ONCE <> True Then
        Form1.Top = 0 'Screen.Height / 2 - Form1.Height / 2
        Form1.Left = Screen.Width / 2 - Form1.Width / 2
    End If
    
    X_Y_DONE_ONCE = True
    Call SET_COMBO1_POSITION
End Sub

Private Sub Form_Load()

'FormStart.MNU_TREESIZE_GO = True
'bIsValid = False
'strFileName = "O:\VB6-EXE\VB-NT\#1 Archive 2009\Old 2003-Now Archive Dump\01 UNWanted\The Best Sound Card Recorder1\The Best Sound Card Recorder1\SoundRecorder.exe"
'If Mid(UCase(strFileName), Len(strFileName) - 3, 4) = ".EXE" Then bIsValid = True
'MsgBox bIsValid & vbCrLf & Mid(UCase(strFileName), Len(strFileName) - 3, 4)
'End
ReDim A4$(500), B4$(500), C4$(500)

Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = False

If COMMON = COMMON Then
    FontSizez = 8
    FontSizez_2 = 7       ' 6.9 -- THE SPECIAL FOLDERIN GET HERE - UNLESS SAME AS OTHER ONE
End If
If GetComputerName = "1-ASUS-X5DIJ" Then
    FontSizez = 7
    FontSizez_2 = FontSizez '     -- THE SPECIAL FOLDERIN GET HERE
End If

If GetComputerName = "2-ASUS-EEE" Then
    FontSizez = 5.5
    FontSizez_2 = FontSizez '     -- THE SPECIAL FOLDERIN GET HERE
End If

If GetComputerName = "4-ASUS-GL522VW" Then
    FontSizez = 6.8
    FontSizez_2 = FontSizez '     -- THE SPECIAL FOLDERIN GET HERE
End If

If GetComputerName = "5-ASUS-P2520LA" Then
    FontSizez = 6.3
    FontSizez_2 = FontSizez '     -- THE SPECIAL FOLDERIN GET HERE
End If

If GetComputerName = "7-ASUS-GL522VW" Then
    FontSizez = 6.8
    FontSizez_2 = FontSizez '     -- THE SPECIAL FOLDERIN GET HERE
End If

If GetComputerName = "8-MSI-GP62M-7RD" Then
    FontSizez = 6.8
    FontSizez_2 = FontSizez '     -- THE SPECIAL FOLDERIN GET HERE
End If
If GetComputerName = "9-ASUS-G815LM" Then
    FontSizez = 8
    FontSizez_2 = FontSizez '     -- THE SPECIAL FOLDERIN GET HERE
End If


Call SET_UP_PULIC_FSO
Call MNU_GC_PROFILE_SET_ON

NAME_FORM = App.EXEName

Me.BackColor = &HA5FCF3
Call FindRGB(Me.BackColor)
Me.BackColor = RGB(R_RGB - 180, G_RGB - 180, B_RGB - 100)


Dim i As String
If App.PrevInstance = True And IsIDE = True Then
    'i = FindWindow(vbNullString, Me.Caption)
    i = FindWinPart_SEARCHER("VB_KEEP_RUNNER")
    ShowWindow i, SW_MAXIMIZE
'        On Error Resume Next
    SetWindowPos i, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End
End If

'KILL ITSELF IN __.EXE KILL SOFTLY
'WHILE ISIDE LEARN
'---------------------------------
Dim VAR
Dim pid As Long
If IsIDE = True Then
    pid = -1
    VAR = cProcesses.GetEXEID(pid, App.Path + "\" + App.EXEName + ".exe")
    If pid <> -1 Then
        'Call Process_HIGH_PRIORITY_CLASS(PID)
        VAR = cProcesses.Process_Kill(pid)
        Beep
        End
    End If
End If



SET_TRACER_ONCE = True

OIP$ = App.Path
TT1 = InStr(OIP$, "Shell")
TT2 = InStr(TT1, OIP$, "Loader")
OIP2$ = Mid$(OIP$, TT1 + 6, TT2 - TT1 - 7)
OIP$ = OIP2$

TEXT_PATH_1 = App.Path + "\Text Loggs\" + GetComputerName + "-" + GetUserName
TEXT_PATH_2 = App.Path + "\Text Loggs\" + GetComputerName

If Dir(TEXT_PATH_1, vbDirectory) = "" Then
    CreateFolderTree TEXT_PATH_1
End If
If Dir(TEXT_PATH_2, vbDirectory) = "" Then
    CreateFolderTree TEXT_PATH_2
End If

Set m_CRC = New clsCRC
m_CRC.Algorithm = CRC32

'WxHex$ = Hex(m_CRC.CalculateString(Form3.Text1.Text))
'WxHex$ = Hex(m_CRC.CalculateFile(Filename))



For Each Control In Controls
    If InStr(Control.Name, "Label1") Then
        Control.Left = 24000
        Control.Top = 0
        Control.Height = 100
        Control.Width = 100
        Control.Visible = False
    End If
Next


Open TEXT_PATH_1 + "\" + OIP$ + " Loads Logg.txt" For Binary As #1

On Error Resume Next
Seek 1, LOF(1) - 3000
On Error GoTo 0
wt5$ = Input$(3000, 1)
Close #1

easy = InStr(wt5$, ":\")
If easy = 1 Then easy = InStr(easy + 1, wt5$, ":\")


If easy > 0 Then wt5$ = Mid$(wt5$, easy - 1)
'wt5$ = wt5$ + "  "
easy = Len(wt5$) - 2

Counter = 1
Do
If easy <= 0 Then Exit Do
easy2 = easy
easy = InStrRev(wt5$, vbCrLf, easy)
If easy < 1 Then
    Exit Do
End If
C1$ = Mid$(wt5$, 2 + easy, easy2 - easy)
If Counter = 1 Then
    C1$ = Mid$(C1$, 1, Len(C1$))
Else
    C1$ = Mid$(C1$, 1, Len(C1$) - 2)
End If

easy2 = easy
easy = InStrRev(wt5$, vbCrLf, easy)
If easy = 0 Then
    C1$ = Mid$(wt5$, 1, InStr(wt5$, vbCrLf) - 1)
    Exit Do
End If
B1$ = Mid$(wt5$, 2 + easy, easy2 - easy)
B1$ = Mid$(B1$, 1, Len(B1$) - 2)
easy2 = easy

If easy > 0 Then easy = InStrRev(wt5$, vbCrLf, easy)
If easy = 0 Then
    A1$ = Mid$(wt5$, 1 + easy, easy2 - easy)
    A1$ = Mid$(A1$, 1, Len(A1$) - 1)
Else
    A1$ = Mid$(wt5$, 2 + easy, easy2 - easy)
    A1$ = Mid$(A1$, 1, Len(A1$) - 2)
End If

TRUE_GO = True
For R1 = 0 To Combo1.ListCount - 1
    If Combo1.List(R1) = C1$ Then TRUE_GO = False
Next
If C1$ <> "" Then
If TRUE_GO = True Then
    Combo1.AddItem C1$
End If
End If





A4$(Counter) = A1$
B4$(Counter) = B1$
C4$(Counter) = C1$
Counter = Counter + 1
Loop Until easy = 0


If Dir$(TEXT_PATH_1 + "\" + OIP$ + " Loads.txt") <> "" Then
    Open TEXT_PATH_1 + "\" + OIP$ + " Loads.txt" For Input As #1
    Line Input #1, A1$
    Line Input #1, B1$
    Line Input #1, C1$
    Close #1
End If
'---- Load Last File Click Here ----
    
Lbl_Title = " LAST FOLDER __ " + C1$

Lbl2.Caption = OIP$ + " Auto Loader"
Form1.Caption = Lbl2.Caption
    On Error Resume Next
    Set F1 = FSO.GetFile(App.Path + "\" + App.EXEName + ".EXE")
    VAR_DATE_1 = 0
    VAR_DATE_1 = F1.datelastmodified
    VAR_DATE_2 = Format(VAR_DATE_1, "DDD DD-MMM-YYYY  HH:MM AM/PM")
    On Error GoTo 0
    
    MNU_VERSION.Caption = "Ver_2020_" + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision)) + "  " + VAR_DATE_2 ' + " _ Matt.Lan@btinternet.com"
    ' MNU_COMPUTER.Caption = "COMPUTER__" + GetComputerName '+ "__USER__" + GetUserName

Form1.MNU_NET_2_GO.Caption = "[__ NET 2 GO __]"

Form1.Top = 0
Form1.Left = 0

Call SET_MENU_PADD_WORK

TIMER_FORMSTART.Enabled = True


End Sub

Sub TIMER_FORMSTART_TIMER()
    TIMER_FORMSTART.Enabled = False
    
    Me.Show
    DoEvents
    Me.WindowState = vbMinimized
    DoEvents
    
    Call FormStart.FormStartLoader
    
    Timer_SUB_CODE_LOAD.Enabled = True

End Sub

Private Sub Lbl_COMBO_NUMBER_Click()
'Lbl_COMBO_NUMBER.CAPTION
End Sub

Private Sub MNU_CLIPBOARD_ALL_NET_PATH_ARRAY_BUILDER_Click()
    
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText NET_PATH_ALL_ARRAY + vbCrLf
    If Err.Number = 0 Then
        Exit Sub
    End If
    End

End Sub

Private Sub MNU_CLIPBOARD_ALL_NET_PATH_Click()

    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText NET_PATH_ALL + vbCrLf
    If Err.Number = 0 Then
        Exit Sub
    End If
    End

End Sub

Private Sub MNU_CLIPBOARD_ALL_NET_PATH_REVERSE_Click()

    NET_PATH_ALL_REVERSE = ""
    NET_PATH_ALL_R = Split(NET_PATH_ALL, vbCrLf)
    For R3 = UBound(NET_PATH_ALL_R) To 0 Step -1
        NET_PATH_ALL_REVERSE = NET_PATH_ALL_REVERSE + NET_PATH_ALL_R(R3) + vbCrLf
    Next
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText NET_PATH_ALL_REVERSE + vbCrLf
    If Err.Number = 0 Then
        Exit Sub
    End If
    End

End Sub

Private Sub MNU_CLIPBOARD_COMPUTER_NAME_SET_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText COMPUTER_NAME_SET
    If Err.Number > 0 Then
        Exit Sub
    End If
    End
End Sub

Private Sub MNU_CLIPBOARD_COMPUTER_NAME_SET_VB6_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText COMPUTER_NAME_SET_VB6
    If Err.Number > 0 Then
        Exit Sub
    End If
    End
End Sub

Private Sub MNU_CLIPBOARDOR_Click()
CLIPBOARDOR = True
End Sub

Private Sub MNU_CLIPBOARDOR_PATH_NAME_Click()
CLIPBOARDOR_PATH_NAME = True
End Sub
Private Sub MNU_CLIPBOARDOR_PATH_LINK_Click()
CLIPBOARDOR_PATH_LINK = True
End Sub
Private Sub MNU_CLIP_PATH_NAME_SHORT_Click()
CLIPBOARDOR_PATH_NAME = True
CLIPBOARDOR_PATH_SHORT = True
End Sub


Private Sub MNU_LOAD_ALL_GOODSYNC_PROFILE_TIC_NOTEPAD_Click()

' -------------------------------------
' NOTE FORM VBS COMMON LABEL NAME CLASH
' SO CALL ShowWindow_2
' -------------------------------------
' Const ShowWindow_2 = 1, DontShowWindow = 0, DontWaitUntilFinished = False, WaitUntilFinished = True

NET_PATH_ALL_R = Split(NET_PATH_ALL, vbCrLf)

Dim M()
ReDim M(UBound(NET_PATH_ALL_R) + 10)
Dim R3
For R3 = 0 To UBound(NET_PATH_ALL_R)
    CK1 = UCase(NET_PATH_ALL_R(R3))
    ' ------------------------------------------------------
    ' YOU MIGHT THINK MY CODE IS CRAPPY
    ' BUT PUT >0 AFTER EACH INSTR AS WHEN COMPARE 2
    ' AND WITH AN AND STATEMENT BETWEEN
    ' IT CHANGE THE VALUE OF LOGIC
    ' SO NOT ABLE LEAVE >0 OUT
    ' INSTR RETURN RESULT HOW FAR IN THE STRING IS POSITION
    ' ------------------------------------------------------
    If InStr(CK1, "1_ASUS") > 0 Then NET_PATH_ALL_R(R3) = ""
    If InStr(CK1, "2_ASUS") > 0 Then NET_PATH_ALL_R(R3) = ""
    If InStr(CK1, "3_LINDA") > 0 Then NET_PATH_ALL_R(R3) = ""
    If InStr(CK1, "E_DRIVE") > 0 Then NET_PATH_ALL_R(R3) = ""
    If InStr(CK1, "2_ASUS") > 0 Then NET_PATH_ALL_R(R3) = ""
    If InStr(CK1, "5_ASUS") > 0 And InStr(CK1, "D_DRIVE") > 0 Then
        NET_PATH_ALL_R(R3) = ""
    End If
'    If InStr(CK1, "7_ASUS") > 0 And InStr(CK1, "D_DRIVE") > 0 Then NET_PATH_ALL_R(R3) = ""
    If InStr(CK1, "8_MSI") > 0 And InStr(CK1, "D_DRIVE") > 0 Then NET_PATH_ALL_R(R3) = ""
Next
i = -1
For R3 = 0 To UBound(NET_PATH_ALL_R)
    CK1 = UCase(NET_PATH_ALL_R(R3))
    CK2 = NET_PATH_ALL_R(R3)
    If CK1 <> "" Then
        SET_GO = False
        If InStr(CK1, "4_ASUS") > 0 Then SET_GO = True
        If InStr(CK1, "7_ASUS") > 0 Then SET_GO = True
        If InStr(CK1, "8_MSI") > 0 Then SET_GO = True
        If SET_GO = True Then
        If InStr(CK1, "C_DRIVE") > 0 Then
            i = i + 1: M(i) = CK2 + "\GoodSync\Profile\jobs-groups-options.tic"
            ' -----------------------------------------------------------------
            For R5 = 1 To 5
                GS_1 = CK2 + "\Users\MATT " + Format(R5, "00") + "\AppData\Local\GoodSync\jobs-groups-options.tic"
                On Error Resume Next
                ' COMPUTER_CODE_VB_GENERATED_TEXT
                XX = ""
                XX = Dir(GS_1)
                On Error GoTo 0
                If XX <> "" Then
                    i = i + 1: M(i) = GS_1
                End If
            Next
            ' -----------------------------------------------------------------
        End If
        End If
        If InStr(CK1, "D_DRIVE") > 0 Then
            i = i + 1: M(i) = CK2 + "\GoodSync\Profile\jobs-groups-options.tic"
        End If
    End If
Next

ReDim Preserve M(i)

NP = ""
If NP = "" Then
    NP1 = "C:\Program Files\Notepad++\notepad++.exe"
    NP = Dir(NP1)
End If
If NP = "" Then
    NP2 = "C:\Program Files (X86)\Notepad++\notepad++.exe"
    NP = Dir(NP2)
End If
If NP = "" Then
    MsgBox "NOTEPAD++ NOT FOUND -- HERE " + vbCrLf + vbCrLf + NP1 + vbCrLf + vbCrLf + NP2
End If
For R3 = 0 To UBound(M)
    On Error Resume Next
    ' COMPUTER_CODE_VB_GENERATED_TEXT
    XX = ""
    XX = Dir(M(R3))
    On Error GoTo 0
    
    If XX <> "" Then
        Dim OBJSHELL
        Set OBJSHELL = CreateObject("Wscript.Shell")
        OBJSHELL.Run NP + " """ + M(R3) + """", ShowWindow_2, DontWaitUntilFinished
        Set OBJSHELL = Nothing
        M(R3) = ""
    End If
Next

XX = ""
For R3 = 0 To UBound(M)
    If M(R3) <> "" Then
        XX = XX + M(R3) + vbCrLf
    End If
Next

If XX <> "" Then
    MsgBox "THESE GOODSYNC PROFILE DID NOT FIND" + vbCrLf + XX, vbMsgBoxSetForeground
End If

End

End Sub

Private Sub MNU_NETWORK_2_STEP_JUMPER_Click()

NETWORK_2_STEP_JUMPER = 2

COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_01 = ""
COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_02 = ""
COMPUTER_NAME_PUT_STORE_NETWORK_2_STEP_JUMPER_03 = ""

Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Visible = True
Form1.MNU_NETWORK_2_STEP_DRIVE_SELECTOR.Caption = "NETWORK DRIVE SELECT - " + Format(2 - NETWORK_2_STEP_JUMPER, "00") + " OF 02"

' MsgBox "SELECT THE NETWORK PATH AND THEN ANOTHER NORMAL FOLDER AND I WILL JUMP THERE AT NETWORK LOCATION"

' SELECT ANY WAY WANT AROUND
' JUST SELECT TWO NETOWRK AND LOCAL PATH

Beep

End Sub

Private Sub MNU_NETWORK_WITH_CLIPBOARDER_Click()

'On Error Resume Next
'PATH1 = Clipboard.GetText
'PATH2 = InStrRev(PATH1, ":\")
'If PATH2 = 0 Then Exit Sub
'PATH2 = InStrRev(PATH1, ":\")
'
'PATH3 = Mid(PATH1, PATH2)
'





End Sub

Private Sub MNU_TREESIZE_Click()

FormStart.MNU_TREESIZE_GO = True

End Sub


Private Sub Timer_KEY_CODE_Timer()

' If IsIDE = True Then Timer_GET_KEY_ASYNC_STATE.Interval = 1000


Dim tPA As POINTAPI, LHWND As Long, O_lhWndParent, lhWndParent, lhWndParentX
GetCursorPos tPA
LHWND = WindowFromPoint(tPA.x, tPA.y)
O_lhWndParent = LHWND
lhWndParent = GetParent(LHWND)
If lhWndParent = 0 Then lhWndParent = O_lhWndParent
lhWndParentX = GetParentHwnd(LHWND)

If GetAsyncKeyState(27) < 0 Then
    If GetForegroundWindow = Me.hWnd Or lhWndParent = Me.hWnd Or lhWndParentX = Me.hWnd Then
        If IsIDE = True Then
            Unload Me
        Else
            Unload Me
            ' Me.WindowState = vbMinimized
        End If
    End If
End If


End Sub

Sub SubCode()



Dim RT
Dim FHEIGHT

Dim RG
RG = 0
LABEL_GAP = 12
Combo1.FontSize = 12
Combo1.FontName = "Arial"
Lbl_Title.Height = Combo1.Height

For Each Control In Controls
    If InStr(Control.Name, "Lab") Then
        RG = RG + 1: Control.Visible = False
        Control.Caption = ""
        Control.AutoSize = True
        Control.Width = 1
        Control.FontSize = FontSizez
    End If
Next




'rg = 0
'For Each Control In Controls
'    If InStr(Control.Name, "Lab") Then
'    If IsNumeric(Mid(Control.Caption, 4, 2)) Then
'        rg = rg + 1: Control.Visible = False
'        Control.Caption = ""
'        Control.AutoSize = True
'        Control.Width = 1
'        Control.FontSize = FontSizez
'    End If
'    End If
'Next


Me.WindowState = vbMaximized
' Me.WindowState = vbNormal
Me.Refresh
Me.Show
DoEvents
'Me.Visible = False
ScanPath.cboMask.Text = "*.lnk"

ScanPath.txtPath.Text = "E:\01 VB Shell Folders\00 Shell Explorer Loader"
Call ScanPath.cmdScan_Click

IV = FormStart.GetWindowsVersion
'REPORT
'5.1 ON WINDOWS XP
'6.2 ON WINDOWS 10
If IV > 5.1 Then
    For r = ScanPath.ListView1.ListItems.Count To 1 Step -1
        If Mid(ScanPath.ListView1.ListItems.Item(r), 1, 3) = "XP " Then
            If IsIDE = False Then ScanPath.ListView1.ListItems.Remove (r)
        End If
    Next
End If

If IV = 5.1 Then
    For r = ScanPath.ListView1.ListItems.Count To 1 Step -1
        A1$ = ScanPath.ListView1.ListItems.Item(r).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(r)
        If InStr((UCase(B1$)), "PROGRAM FILES (X86)") > 0 Then
            Call GETSHORTLINK(A1$ + B1$)
            D1$ = txtTargetPath
            If FSO.FileExists(D1$) = False And FSO.FolderExists(D1$) = False Then
                If IsIDE = False Then
                    ScanPath.ListView1.ListItems.Remove (r)
                End If
            End If
        End If
    Next
End If


For r = ScanPath.ListView1.ListItems.Count To 0 Step -1
    ' A1$ = ScanPath.ListView1.ListItems.Item(R).SubItems(1)
    If r = 0 Then Exit For
    B1$ = ScanPath.ListView1.ListItems.Item(r)
    If InStr(B1$, "Videos.lnk") > 0 Then
        ScanPath.ListView1.ListItems.Remove (r)
    End If
Next


Lbl2.Top = 0
Lbl2.Left = 0

x = Lbl2.Height + LABEL_GAP
If Lbl2.Visible = False Then
    x = LABEL_GAP
End If

Check1.Top = x
Check1.Left = (Form1.Width / 2) + 10
Check1.Width = Form1.Width / 2

TG = x

x = x + Lbl_Title.Height + LABEL_GAP
tx = x

w = 0

RD = 0

DoEvents

SH = Screen.Height - 4000 'higer = smaller
' SH = Me.Height - 100
'---------------------------
' GOT TO SHOW THE FORM FIRST
'---------------------------

'ME.hWnd
'Dim FORM_MAX As Rect
'i = GetWindowRect(Me.hWnd, FORM_MAX)
'ScreenTwipsX = Screen.TwipsPerPixelX
'ScreenTwipsY = Screen.TwipsPerPixelY
'SH = (FORM_MAX.Top + FORM_MAX.Bottom) * Screen.TwipsPerPixelY

RX = 0
For r = 1 To ScanPath.ListView1.ListItems.Count + 100
    RX = RX + 1
    If RX > ScanPath.ListView1.ListItems.Count Then Exit For
    
    A1$ = ScanPath.ListView1.ListItems.Item(RX).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(RX)
    If InStrRev(A1$, "\") > 0 Then
        seedy$ = Mid$(A1$, InStrRev(A1$, "\", Len(A1$) - 1))
    End If
    
    If InStrRev(A1$, "TITLE_BLOCK") > 0 Then
        SDD = Mid$(B1$, InStrRev(B1$, "\", Len(B1$) - 1))
        ScanPath.ListView1.ListItems.Item(RX).SubItems(1) = SDD
        ScanPath.ListView1.ListItems.Item(RX) = ScanPath.ListView1.ListItems.Item(RX) + " ----#"
    End If
    
    If oseedy$ <> seedy$ Then
        With ScanPath.ListView1
            Set LV = .ListItems.Add(RX, , seedy$ + " ----#")
            LV.SubItems(1) = A1$
        End With
    End If
    
    oseedy$ = seedy$
Next



'For Each Control In Label1
'    If Control.Index < Label1.Count - 1 Then
'    A = Control.Caption
'    B = Label1.Item(Control.Index + 1).Caption
'    If A <> "" And B <> "" Then
'    If Control.Caption = Label1.Item(Control.Index + 1).Caption Then
'        Control.Visible = False
'    End If
'    End If
'    End If
'Next


' DUPE SEEM GOT IN THERE
' AND THEN REMOVE THERE
' Sat 11-Jul-2020 08:35:54
For RX = ScanPath.ListView1.ListItems.Count - 1 To 1 Step -1
    TT1 = ScanPath.ListView1.ListItems.Item(RX)
    TT2 = ScanPath.ListView1.ListItems.Item(RX + 1)
    If TT1 = TT2 Then
        ScanPath.ListView1.ListItems.Remove (RX)
    End If
Next

' DUPE -- SEEM TO GOT FROM C-DRIVE LINK FOLDER REPEAT OVER AGAIN
For RX1 = 1 To ScanPath.ListView1.ListItems.Count - 1
    If RX1 > ScanPath.ListView1.ListItems.Count - 1 Then Exit For
    TT1 = ScanPath.ListView1.ListItems.Item(RX1)
    For RX2 = ScanPath.ListView1.ListItems.Count - 1 To RX1 + 1 Step -1
        TT2 = ScanPath.ListView1.ListItems.Item(RX2)
        If TT1 = TT2 Then
            ScanPath.ListView1.ListItems.Remove (RX2)
            Exit For
        End If
    Next
Next




x = tx
RX = 0
RT = 0
For r = 1 To ScanPath.ListView1.ListItems.Count
    
    RX = RX + 1
    RT = RT + 1
    
    If Label1(RX - 1).Top > SH Then
        x = tx
    End If
    
    Label1(RX).Top = x
    
    Label1(RX).Visible = True
    A1$ = ScanPath.ListView1.ListItems.Item(RX).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(RX)
    
    If InStrRev(A1$, "\") > 0 Then
        seedy$ = Mid$(A1$, InStrRev(A1$, "\", Len(A1$) - 1))
    End If
    
    If oseedy$ <> seedy$ Then
        txr = Not txr
    End If
    oseedy$ = seedy$
    
    'LABEL_BACKCOLOR_VAR = Label1(RX).BackColor
    If txr = -1 Then LABEL_BACKCOLOR_VAR = &HFF00&
    If txr = 0 Then LABEL_BACKCOLOR_VAR = RGB(127, 200, 127) '&H56E25E
    
    If InStrRev(B1$, " ----#") > 0 Then
        RT = RT - 1
    End If
    
    'If txr = 0 Then Label1(R).BackColor = &HFF00&
    
    ttg$ = ScanPath.ListView1.ListItems.Item(RX)
    Mid$(ttg$, 1, 1) = UCase$(Mid$(ttg$, 1, 1))
    If InStr(ttg$, ".lnk") > 0 Then ttg$ = Mid$(ttg$, 1, InStrRev(ttg$, ".") - 1)
    EE = InStr(ttg$, "&")
    If EE > 0 Then
        ttg$ = Mid$(ttg$, 1, EE) + Mid$(ttg$, EE)
    End If
    
    If InStrRev(B1$, " ----#") > 0 Then
        LABEL_BACKCOLOR_VAR = RGB(255, 255, 255)
        ttg$ = Replace(ttg$, "\", " \ ")
        ttg$ = Replace(ttg$, " ----#", "")
        
        Label1(RX).Caption = Trim(ttg$)
            ' If Label1(RX).Caption = Label1(RX - 1).Caption Then Stop

        Label1(RX).Alignment = 2
    Else
        If InStr(A1$, "SPECIAL") > 0 Then
            Label1(RX).FontSize = FontSizez_2
        End If
        Label1(RX).Caption = Trim(Format$(RT, "000") + ". " + ttg$)
        
    End If
    'If InStr(Label1(RX), 125) = 1 Then Stop
    
    
    If InStr(Label1(RX), "\\") = 0 Then
        If InStr(Label1(RX), ":\") > 0 Then
            SET_GO = False
            'If FSO.FolderExists(Mid(Label1(RX), 5)) = True Then SET_GO = True
            If Dir(Mid(Label1(RX), InStr(Label1(RX), ":") - 1), vbDirectory) <> "" Then SET_GO = True
            If FSO.FolderExists(Mid(Label1(RX), InStr(Label1(RX), ":") - 1)) = True Then SET_GO = True
            If SET_GO = True Then
                LABEL_BACKCOLOR_VAR = Label3.BackColor - RGB(40, 40, 40)
            Else
                LABEL_BACKCOLOR_VAR = Label3.BackColor
            End If
        End If
    End If
    
    If InStr(Label1(RX), "\\") > 0 Then
        LABEL_BACKCOLOR_VAR = Label2.BackColor
    End If
    
    If InStr(Label1(RX), "_ C:\USER") > 0 Then
        SET_HERE_2 = False
        CV1 = Label1(RX).Caption
        If InStr(CV1, "3-L") > 0 And InStr(CV1, "MATT 01") > 0 Then SET_HERE_2 = True
        If InStr(CV1, "4-A") > 0 And InStr(CV1, "MATT 01") > 0 Then SET_HERE_2 = True
        If InStr(CV1, "5-A") > 0 And InStr(CV1, "MATT 01") > 0 Then SET_HERE_2 = True
        If InStr(CV1, "7-A") > 0 And InStr(CV1, "MATT 04") > 0 Then SET_HERE_2 = True
        If InStr(CV1, "8-M") > 0 And InStr(CV1, "MATT 01") > 0 Then SET_HERE_2 = True
        If SET_HERE_2 = True Then
            LABEL_BACKCOLOR_VAR = &HFF00& 'Label3.BackColor - RGB(40, 40, 40)
        End If
        If InStr(CV1, GetComputerName) > 0 And InStr(CV1, GetUserName) > 0 Then
            LABEL_BACKCOLOR_VAR = RGB(132, 164, 32)
        End If

    End If
    
    Label1(RX).BackColor = LABEL_BACKCOLOR_VAR
    
    If InStr(LABLE_BACKCOLOR_SET, LABEL_BACKCOLOR_VAR) = 0 Then
        LABLE_BACKCOLOR_SET = LABLE_BACKCOLOR_SET + Hex(Label1(RX).BackColor)
    End If
    
    x = x + Label1(RX).Height + LABEL_GAP
    FHEIGHT = Label1(RX).Top + Label1(RX).Height + 420
    ' If FHEIGHT > FHEIGHTX Then FHEIGHTX = FHEIGHT
    RD = RD + 1



Next
DoEvents

x = tx
TD = 29
xgag = 0
Dim xy(20)

For r = 1 To ScanPath.ListView1.ListItems.Count
    If Label1(r - 1).Top > SH Then
        If tig = 0 Then tig = r
        x = tx
        xgag = xgag + 1
        wdt = 0
        For RT = r - 1 To r - tig Step -1
        If Label1(RT).Width > wdt Then wdt = Label1(RT).Width
        Next
        xy(xgag) = wdt + 150
        tig2 = r
    End If
    Label1(r).Top = x
    x = x + Label1(r).Height + LABEL_GAP
Next

xgag = xgag + 1
wdt = 0
For RT = tig2 To r - 1
    If Label1(RT).Width > wdt Then wdt = Label1(RT).Width
Next
xy(xgag) = wdt + 150

For r = 1 To ScanPath.ListView1.ListItems.Count
    fw = Label1(r).Width
    If fw > fw2 Then fw2 = fw + 200
Next
fw2 = fw2

xgag = 1: xgax2 = 0
For r = 1 To ScanPath.ListView1.ListItems.Count
    Label1(r).AutoSize = False
    
    DoEvents
    xxb = Label1(r - 1).Top
    xxb = Form1.Height
    
    If Label1(r - 1).Top > SH Then
        xgax2 = xgax2 + 1
        xgag = 0
        For rs2 = 1 To xgax2
            xgag = xgag + xy(rs2)
        Next
    End If
    
    Label1(r).Width = xy(xgax2 + 1) - 20
    Label1(r).Left = xgag
    
Next

FHEIGHTX = 0
For RT = 1 To ScanPath.ListView1.ListItems.Count
    FHEIGHT = Label1(RT).Top + Label1(RT).Height
    If FHEIGHT > FHEIGHTX Then
        FHEIGHTX = FHEIGHT
'        Debug.Print Label1(rt).Caption
    Else
    '    Exit For
    End If
Next

For Each Control In Label1
    If InStr(UCase(Control.Caption), "01 VB SHELL FOLDERS") > 0 Then
        Control.Caption = "\ 01 VB SHELL FOLDER \"
    End If
    If InStr(Control.Caption, "SHELL:DOWNLOADS") > 0 Then
        Control.Caption = Replace(Control.Caption, "SHELL:DOWNLOADS", "SHELL:DOWNLOAD____")
    End If
    If InStr(UCase(Control.Caption), "00 LISTS-COMMON-WORDS") > 0 Then
        Control.Caption = Replace(UCase(Control.Caption), "00 LISTS-COMMON-WORDS", "00 LISTS-COMMON-WORD____")
    End If
    If InStr(UCase(Control.Caption), ". VIDEOS") > 0 Then
        Control.Caption = Replace(Control.Caption, ". VIDEOS", ". VIDEO____")
    End If
    If InStr(UCase(Control.Caption), ". DOWNLOADS") > 0 Then
        Control.Caption = Replace(UCase(Control.Caption), ". DOWNLOADS", ". DOWNLOAD____")
    End If
    If InStr(UCase(Control.Caption), ". 00 SHELL GAMES LOADER") > 0 Then
        Control.Caption = Replace(UCase(Control.Caption), ". 00 SHELL GAMES LOADER", ". 00 SHELL GAME LOADER__")
    End If
    If InStr(UCase(Control.Caption), ". 0 00 ART LOGGERS") > 0 Then
        Control.Caption = Replace(UCase(Control.Caption), ". 0 00 ART LOGGERS", ". 0 00 ART LOGGER")
    End If
    If InStr(UCase(Control.Caption), ". ANDROID APKS") > 0 Then
        Control.Caption = Replace(UCase(Control.Caption), ". ANDROID APKS", ". ANDROID APK____")
    End If
    If InStr(UCase(Control.Caption), "@GMAIL.COM") > 0 Then
        Control.Caption = Replace(UCase(Control.Caption), "@GMAIL.COM", "@____")
    End If
    If InStr(UCase(Control.Caption), "MY WEBS") > 0 Then
        Control.Caption = Replace(UCase(Control.Caption), "MY WEBS", "MY WEB____")
    End If
    If InStr(UCase(Control.Caption), "INSTALLATIONS") > 0 Then
        Control.Caption = Replace(UCase(Control.Caption), "INSTALLATIONS", "INSTALLATION____")
    End If
    If InStr(UCase(Control.Caption), "01 FAVS") > 0 Then
        Control.Caption = Replace(UCase(Control.Caption), "01 FAVS", "01 FAV")
    End If
    If InStr(UCase(Control.Caption), "REG KEY SETTINGS") > 0 Then
        Control.Caption = Replace(UCase(Control.Caption), "REG KEY SETTINGS", "REG KEY SETTER____")
    End If
    
Next



Form1_Width = Label1(r - 1).Left + Label1(r - 1).Width + 280
If Form1_Width > Screen.Width Then Form1_Width = Screen.Width - 120

' IF FORM SET TO GO ON VBMAXIMIZED ADJUST MUGHT NOT LOOK PROPER
' -------------------------------------------------------------
Form1_Height = FHEIGHTX + 1400

Lbl2.Width = Form1.Width
Lbl_Title.Top = TG
Lbl_Title.Left = 0
Combo1.Top = TG
Lbl_COMBO_NUMBER.Top = Lbl_Title.Top

RTV = 0
CONTROL_COLUMN_NUMBER = 0
For Each Control In Label1
    RTV_2 = Control.Left
    If RTV_2 <> RTV Then
        CONTROL_COLUMN_NUMBER = CONTROL_COLUMN_NUMBER + 1
        RTV = RTV_2
        If CONTROL_COLUMN_NUMBER = 4 Then
            CONTROL_COLUMN_NUMBER = Control.Left + Control.Width
            Exit For
        End If
    End If
Next

Lbl_Title.Width = CONTROL_COLUMN_NUMBER

Combo1.Text = "List of Last View Folder"
Combo1.SelStart = 0

Call SET_COMBO1_POSITION

Lbl3.Top = FHEIGHT - 410
Lbl3.Width = Form1.Width
Lbl3.Left = 0
Form1.Refresh


' Stop

For R1 = 0 To Combo1.ListCount - 1
    XX = Trim(Mid(Combo1.List(R1), InStr(Combo1.List(R1), ".") + 1))
    For Each Control In Label1
    If InStr(Control.Caption + "----", XX + "----") > 0 Then
        
        'Color c1 = Color.Red;
        'Color c2 = Color.FromArgb(c1.A,
        '(int)(c1.R * 0.8), (int)(c1.G * 0.8), (int)(c1.B * 0.8));
        R_RGB = 0
        G_RGB = 0
        B_RGB = 0
        Call FindRGB(Control.BackColor)
        XX2 = Hex(Control.BackColor)
'        Debug.Print Control.Caption
'        Debug.Print XX2
'        Debug.Print R_RGB
'        Debug.Print G_RGB
'        Debug.Print B_RGB
        
        ' CONVERT COLOUR TO LESS BRIGHTNESS USE NUMBER BELOW 1 AS MULTIPLIER
        ' AND ABOVE 1 FOR HIGHER
        If R_RGB < 255 Then R_RGB = R_RGB * 1.4
        If G_RGB < 255 Then G_RGB = G_RGB * 1.4
        If B_RGB < 255 Then B_RGB = B_RGB * 1.4
        If R_RGB >= 255 Then R_RGB = 255
        If G_RGB >= 255 Then G_RGB = 255
        If B_RGB >= 255 Then B_RGB = 255
        
        Control.BackColor = RGB(Int(R_RGB), Int(G_RGB), Int(B_RGB))
        Control.BackColor = RGB(127, 127, 200)
        Control.Refresh
    End If
    Next


Next

' Stop


Me.Visible = True
ReDim Preserve X_COLOR_ARRAY(Label1.Count)

' THE TIMER WILL TELL FOR DEAD LINK
' Timer2_Timer




End Sub

Sub FindRGB(Col As Long)
R_RGB = &HFF& And Col
G_RGB = (&HFF00& And Col) \ 256
B_RGB = (&HFF0000 And Col) \ 65536
End Sub


Private Sub Timer2_Timer()

'If FSO.FolderExists("C:\Program Files (x86)") = False Then

XXT2 = XXT2 + 1
If XXT2 > ScanPath.ListView1.ListItems.Count Then
    Timer2.Enabled = False
    Close FR1
    Exit Sub
End If

If Form1.Label1(XXT2).BackColor = RGB(255, 255, 255) Then
    Exit Sub
End If


'DOEVENTS
SET_REAL_LINK = False
If InStr(Label1(XXT2), ":\") > 0 Then SET_REAL_LINK = True
If InStr(Label1(XXT2), "\\") > 0 Then SET_REAL_LINK = True

If SET_REAL_LINK = True Then
    Exit Sub
End If


' ------------------------------------------------------------------
' WHY CODE HERE
' HAD TO REM OUT OR NONE TRACER WORK
' ------------------------------------------------------------------
If InStr(LABLE_BACKCOLOR_SET, Form1.Label1(XXT2).BackColor) = 0 Then
'    Exit Sub
End If

'XXT8 = XXT8 + 1
'If XXT8 > 10 And XXT2 < XXT4 Then
'    If Timer3.Enabled = False Then
'        Timer3.Interval = 1
'        Timer3.Enabled = True
'    End If
'    Call Timer3_Timer
'End If

X_COLOR_ARRAY(XXT2) = Form1.Label1(XXT2).BackColor

Form1.Label1(XXT2).BackColor = RGB(255, 255, 255)
Form1.Label1(XXT2).Refresh

If SET_REAL_LINK = False Then
    ' Sleep 10
End If

H1$ = ScanPath.ListView1.ListItems.Item(XXT2).SubItems(2)
A1$ = ScanPath.ListView1.ListItems.Item(XXT2).SubItems(1)
B1$ = ScanPath.ListView1.ListItems.Item(XXT2)
'C1$ = Form1.Label1(XXT2).Caption

On Local Error Resume Next


' SET HERE WHY
' D1$ = A1$ + B1$

' IF TYPE OF THAT NOT REAL NORMAL LIKE CONTROL PANEL ITEM AND OTHER THEN HERE GREY
' FOR SCHEUDLE-TASKS ON XP
' --------------------------------------------------------------------------------
If InStr(LCase(B1$), ".lnk") > 0 Then
    'LINK
    
    FOLEDERNAME = ""
    WxHex$ = Hex(m_CRC.CalculateFile(A1$ + B1$))
    For r = 1 To UBound(TEXT_PATH_4_ARRAY)
        If InStr(TEXT_PATH_4_ARRAY(r), WxHex$) > 0 Then
            FOLEDERNAME = TEXT_PATH_4_ARRAY(r + 1)
            FOLEDERNAME = TEXT_PATH_4_ARRAY(r + 2)
        End If
    Next
    
    txtTargetPath = ""
    If FOLEDERNAME = "" Then
        On Local Error GoTo 0
        Call GETSHORTLINK(A1$ + B1$)
    End If
    D1$ = txtTargetPath
    
    If FOLEDERNAME <> "" Then
        D1$ = FOLEDERNAME
    End If
    
    If Trim(D1$) = "" Then
        Label1(XXT2).BackColor = RGB(147, 147, 147)
        X_COLOR_ARRAY(XXT2) = Form1.Label1(XXT2).BackColor
        If Dir(A1$ + B1$) <> "" Then
            WxHex$ = Hex(m_CRC.CalculateFile(A1$ + B1$))
            FR1 = FreeFile
            TEXT_PATH_4 = TEXT_PATH_2 + "\" + OIP$ + " CRC SPEED BROKEN LINK.txt"
            Open TEXT_PATH_4 For Output As FR1
            ' HERE GET TEXT_PATH_8 IN
            ' Timer_SUB_CODE_LOAD_Timer
            ' SEARCH -- Open TEXT_PATH_4 For Output As FR1
            ' SEARCH -- Open TEXT_PATH_8 For Output As FR1
            Print #FR1, WxHex$
            Print #FR1, A1$ + B1$
            Print #FR1, D1$
            On Error Resume Next
            ' DISK FULL
            Close #FR1
            If Err.Number > 0 Then Reset
            On Error GoTo 0
            
        End If
        Exit Sub
    End If

    FOLEDERNAME = "C:\Program Files (x86)"
    If InStr(A1$ + "--", "Program Files (x86)" + "--") > 0 Then
        If FSO.FolderExists(FOLEDERNAME) = False Then
            Label1(XXT2).BackColor = RGB(147, 147, 147)
            X_COLOR_ARRAY(XXT2) = Form1.Label1(XXT2).BackColor
            WxHex$ = Hex(m_CRC.CalculateFile(A1$ + B1$))
            FR1 = FreeFile
            TEXT_PATH_4 = TEXT_PATH_2 + "\" + OIP$ + " CRC SPEED BROKEN LINK.txt"
            Open TEXT_PATH_4 For Output As FR1
            Print #FR1, WxHex$
            Print #FR1, FOLEDERNAME
            Print #FR1, D1$
            Close #FR1
            Exit Sub
        End If
    End If

    If Mid(D1$, 2, 2) = ":\" Or Mid(D1$, 1, 2) = "\\" Then
        ' FILE_CHECK_IF_EXIST_FOR_COLOUR = D1$
        ' ------------------------------------
        D4$ = D1$
        SET_GO_4 = True
        If FSO.FileExists(D4$) = False And FSO.FolderExists(D4$) = False Then
            SET_GO_4 = False
        End If
        If H1$ <> "" Then
            D4$ = H1$
            If SET_GO_4 = False Then
                SET_GO_4 = True
                If FSO.FileExists(D4$) = False And FSO.FolderExists(D4$) = False Then
                    SET_GO_4 = False
                End If
            End If
        End If
        If SET_GO_4 = False Then
            Label1(XXT2).BackColor = RGB(240, 127, 127) 'QBColor(12)
            X_COLOR_ARRAY(XXT2) = Form1.Label1(XXT2).BackColor
            WxHex$ = Hex(m_CRC.CalculateFile(A1$ + B1$))
            FR1 = FreeFile
            TEXT_PATH_4 = TEXT_PATH_2 + "\" + OIP$ + " CRC SPEED BROKEN LINK.txt"
            Open TEXT_PATH_4 For Output As FR1
            Print #FR1, WxHex$
            Print #FR1, A1$ + B1$
            Print #FR1, D4$
            
            On Error Resume Next
            ' DISK FULL
            Close #FR1
            If Err.Number > 0 Then Reset
            On Error GoTo 0
            
        End If
            
    End If

End If
End Sub

Private Sub Timer3_Timer()
'Timer3.INT
If SET_TRACER_ONCE = True Then
    TRACER = 10
    SET_TRACER_ONCE = False
End If

'TRACER = 20
'TRACER = TRACER / (ScanPath.ListView1.ListItems.Count / (1 + (ScanPath.ListView1.ListItems.Count - XXT4)))

'Debug.Print TRACER

If TRACER < 0 Then TRACER = 0

If XXT4 + TRACER > ScanPath.ListView1.ListItems.Count Then
    Do
        TRACER = TRACER - 1
    Loop Until XXT4 + TRACER < ScanPath.ListView1.ListItems.Count
End If

If XXT4 + TRACER >= XXT2 Then Exit Sub
XXT4 = XXT4 + 1

If XXT4 > ScanPath.ListView1.ListItems.Count Then
    Timer3.Enabled = False
    Exit Sub
End If

    
If X_COLOR_ARRAY(XXT4) > 0 Then
    TRACER = TRACER - 1
    ' If TRACER < 0 Then Stop
    'If TRACER < 0 Then Timer2.Interval = Timer2.Interval + 1
    'If Timer2.Interval > 15 Then Timer2.Interval = 15
End If
  
    
If X_COLOR_ARRAY(XXT4) > 0 And Form1.Label1(XXT4).BackColor <> X_COLOR_ARRAY(XXT4) Then
    Form1.Label1(XXT4).BackColor = X_COLOR_ARRAY(XXT4)
End If

End Sub



Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Sub CLIP_PATH_LINK(TTIT)
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText TTIT
    On Error Resume Next
    If Err.Number = 0 Then
        Exit Sub
    End If

    End
End Sub

Sub CLIP_PATH_NAME(TTIT)
    On Error Resume Next
    Clipboard.Clear
    If CLIPBOARDOR_PATH_SHORT = False Then
        TTEYES = GetLongName(TTIT)
        If TTEYES <> TTIT And TTEYES <> "" Then TTIT = TTEYES
    End If
    Clipboard.SetText TTIT
    End
End Sub




Private Sub Label1_Click(Index As Integer)

'Beep

If CLIPBOARDOR = True Then
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Mid(Label1(Index).Caption, InStr(Label1(Index).Caption, " ") + 1)
    If Err.Number = 0 Then
        Exit Sub
    End If

    End
End If


If Label1(Index).BackColor = QBColor(12) Then
    MsgBox "THIS LINK IS MARKED RED FOR NOT USE"
End If

If Label1(Index).BackColor = RGB(255, 255, 255) Then
    LoadFolder = True
End If

If SetTrueToLoadLast = False Then
A1$ = ScanPath.ListView1.ListItems.Item(Index).SubItems(1)
B1$ = ScanPath.ListView1.ListItems.Item(Index)
C1$ = Form1.Label1(Index).Caption
Else
'vLaunch A1$ + B1$, vbNullString
'End
End If

If CLIPBOARDOR_PATH_LINK = True Then
    Call Form1.CLIP_PATH_LINK(A1$ + B1$)
End If


'----
'Shell Commands List for Windows 10 | Tutorials
'https://www.tenforums.com/tutorials/3109-shell-commands-list-windows-10-a.html
'----
If LoadFolder = True And InStr(C1$, "NETWORK COMPUTER NAME") > 0 Then

    Call SaveLoggs

    Call FormStart.EXPLORER_WITH_SHELL("", "EXPLORER shell:NetworkPlacesFolder")
    
End If

' ------------------
' THIS NOT OFTEN RUN
' SEE BELOW
' ------------------
If LoadFolder = True Then
    
    Call SaveLoggs
    Call FormStart.EXPLORER_WITH_SHELL("/e,", A1$)

End If

Call SaveLoggs

Call FormStart.LabelClick(Index)
' ------------------------------

End Sub

Private Sub Lbl2_Click()

Shell "explorer /e, E:\01 VB Shell Folders\00 " + App.EXEName, vbNormalFocus

End Sub

Private Sub MNU_EXIT_Click()
End
End Sub

Private Sub Mnu_LoadFolder_Click()
LoadFolder = True
End Sub

Private Sub Mnu_Sync_Click()

ScanPath.ListView1.ListItems.Clear

Call FormStart.FormStartLoader

Call SubCode

End Sub

Private Sub VB_RUN_NOT_WHEN_IDE_AND_THEN_SHOW()
' -----------------------------------------------------------------
' IS SHORT CUT NOT SET ADMINISTATOR FOR THIS PARTICALUR CODE IN EXE
' MODE IT WILL NOT RUN LOAD ITSELF ON THE VB COMPILER
' BUT AND WILL RUN OTHER PROGRAM
' AS TALK OTHER PROGRAM ARE OKAY BUT THEN MAYBE SET AS ADMIN
' AMSWER GIVEN
' AUG 28 SUN
' -----------------------------------------------------------------
If 1 = 1 And IsIDE = True Then
    MNU_VB.Enabled = False
    Beep
    MsgBox "NOT WHEN ISIDE = TRUE"
    Exit Sub
End If

Dim ReturnHwnd As Long
Dim i
'VB ONLY WANTS THE 1ST OF THE 2 HWND
'ReturnHwnd = FindWindowPartVB("ClipBoard Logger - Microsoft Visual Basic[")
'------------------------------------------------
'DONT NEED ABOVE USE THIS
X1 = FindWindow("wndclass_desked_gsk", vbNullString)
'------------------------------------------------
'FIND THE WINDOW PRIZE TREASURE CHEST BUST BLOSSOM
'TRAIN SPOTTER
'------------------------------------------------
'-----------------------------------------------------------
'CHECK CLASS NAME IS VB IDE WINDOW WE WANT OR IF ANOTHER NOT
'-----------------------------------------------------------
X2 = GetWindowTitle(X1)
If InStr(X2, " - Microsoft Visual Basic [") > 0 Then
    'RUN A WINDOW SPY
    WIN_SPY_NAME = "ClipBoard Logger"
    If InStr(X2, WIN_SPY_NAME) > 0 Then
    
        MsgBox "DON'T RUN VB IDE - LOADED"
        i = GetWindowState(X1)
        If i = vbMinimized Then
            SHOWVAR = SW_SHOWDEFAULT
            ShowWindow ReturnHwnd, SHOWVAR
        End If
        
        EXIT_TRUE = True
        Unload Me
        Exit Sub
    End If
End If
'------------------------------------------------
'BETTER LINE NEXT DON'T USE VB MENU IF ISIDE
'------------------------------------------------
'TEMP WORK AROUND
'OVER DRIVE
'OVER RIDE
'------------------------------------------------
'FINDWINDOW PART PROBLEM IN EXE AND WHEN IN WIN 7
'------------------------------------------------
'WIN 7 PROBLEM MUST USE EXTRA LAST LEFT SQUARE BRACKET OF SERACH END ABOVE
'------------------------------------------------
If ReturnHwnd > 0 Then
    If MsgBox("ReturnHwnd" + Str(ReturnHwnd) + " -- " + "MeHwnd" + Str(Me.hWnd) + vbCrLf + "VB Code " + vbCrLf + WIN_SPY_NAME + vbCrLf + " already Running - Sure Want to Run VB IDE", vbYesNo) = vbYes Then
        WANT_TO_RUN_ANYWAY = True
    End If
End If

If ReturnHwnd > 0 Then
    i = GetWindowState(ReturnHwnd1)
    If i = vbMinimized Then
'        SHOWVAR = SW_RESTORE
'        SHOWVAR = SW_SHOW
'        SHOWVAR = SW_SHOWNA
'        SHOWVAR = SW_SHOWDEFAULT
'        SHOWVAR = SW_SHOWNORMAL
'        SHOWVAR = SW_SHOWMAXIMIZED
        SHOWVAR = SW_SHOWDEFAULT
        ShowWindow ReturnHwnd, SHOWVAR
        'GUESS MAYBE
        'SetForegroundWindow (ReturnHwnd)
        DoEvents
    End If
   
    'MADE REDUNDANT CODE BY CONDICTION HERE AND ABOVE
    If WANT_TO_RUN_ANYWAY = False Then
        MsgBox "EXIT AS FOUND WINODW OF VB AND PUT TO SHOW FOCUS"
        EXIT_TRUE = True
        Unload Me
        Exit Sub
    End If
End If

Dim CODER_VBP_FILE_NAME_2
Dim VB_1, VB_2, VB_3
VB_1 = "C:\PROGRAM FILES\Microsoft Visual Studio\VB98\VB6.EXE"
VB_2 = "C:\PROGRAM FILES (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
If Dir(VB_1) <> "" Then VB_3 = VB_1
If Dir(VB_2) <> "" Then VB_3 = VB_2
'------------------------------------------------------
'ADD MICROSFOT SCRIPTING RUNTIME FOR HERE IN REFERENCES
'------------------------------------------------------
Dim OBJSHELL
Set OBJSHELL = CreateObject("Wscript.Shell")
CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
OBJSHELL.Run """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
Set OBJSHELL = Nothing
EXIT_TRUE = True
Unload Me
End Sub
Private Sub Mnu_VB_ME_Click()
    
    Call VB_RUN_NOT_WHEN_IDE_AND_THEN_SHOW
    
    Dim CODER_VBP_FILE_NAME_2
    Dim VB_1, VB_2, VB_3
    VB_1 = "C:\PROGRAM FILES\Microsoft Visual Studio\VB98\VB6.EXE"
    VB_2 = "C:\PROGRAM FILES (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
    If Dir(VB_1) <> "" Then VB_3 = VB_1
    If Dir(VB_2) <> "" Then VB_3 = VB_2
    '------------------------------------------------------
    'ADD MICROSFOT SCRIPTING RUNTIME FOR HERE IN REFERENCES
    '------------------------------------------------------
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
    Call CHECK_INTEGRITY_OF_VISUAL_BASIC_PROJECT_VBP(CODER_VBP_FILE_NAME_2)
    
    Dim OBJSHELL
    Set OBJSHELL = CreateObject("WSCRIPT.SHELL")
    OBJSHELL.Run """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
    Set OBJSHELL = Nothing
    
    EXIT_TRUE = True
    Unload Me
    Exit Sub
End Sub
Private Sub MNU_VB_FOLDER_Click()
    Shell "EXPLORER /SELECT, """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
    'End
End Sub
Sub CHECK_INTEGRITY_OF_VISUAL_BASIC_PROJECT_VBP(CODER_VBP_FILE_NAME_2)
    ' --------------------------------------------------------------------
    ' FIND IF THE IS NOT CORRECT VERSION AND SIMPLE PUT CORRECT
    ' ONE MY COMPUTER SEEM PROBLEM WITH BASIC
    ' TALK HAS WRONG VERISON ONCE EVERY FEW DAY
    ' AND EDIT IT TO WHAT NORM VERISON SUPPOSED TO BE AND FINE
    ' DISCOVERY -- IT SEEM TO HAPPEN JUST AFTER A COMPILE SOMETIME_
    ' [ Monday 15:37:30 Pm_17 June 2019 ]
    ' --------------------------------------------------------------------
    ' --------------------------------------------------------------------
    Dim r, FR
    Dim VAR_STRING As String

    FR = FreeFile
    Open CODER_VBP_FILE_NAME_2 For Binary As FR
        VAR_STRING = Space(LOF(FR))
        Get #FR, , VAR_STRING
    Close FR
    ' --------------------------------------------------------------
    ' # .VBP # .VBP" -- # MSCOMC -- # OCX -- # MSCOMCTL # CTL # .CTL
    ' --------------------------------------------------------------
    ' FOUND TO BE ERROR
    ' --------------------------------------------------------------
    ' Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0; MSCOMCTL.OCX
    ' --------------------------------------------------------------
    ' CORRECT VALUE
    ' --------------------------------------------------------------
    ' Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0; MSCOMCTL.OCX
    ' --------------------------------------------------------------
    XR1 = InStr(UCase(VAR_STRING), UCase("A1}#2.1; MSCOMCTL.OCX"))
    XR2 = InStr(UCase(VAR_STRING), UCase("A1}#2.#")) ' ---- "A1}#2.#0; mscomctl.OCX"
    XR3 = InStr(UCase(VAR_STRING), UCase("#0; MSCOMCTL.OCX"))
    If XR1 = 0 And XR2 > 0 And XR3 > 0 Then
        GET_01 = InStr(UCase(VAR_STRING), UCase("MSCOMCTL.OCX")) - XR2
        Mid(VAR_STRING, XR2, GET_01 + Len("MSCOMCTL.OCX")) = "A1}#2.1#0; MSCOMCTL.OCX"
        ' MsgBox "MSCOMCTL.OCX" + vbCrLf + "WRONG VERSION -- CHANGE TO" + vbCrLf + "#2.1# -- MSCOMCTL.OCX"), vbMsgBoxSetForeground
        GO_NEXT_IN = True
    End If
    If GO_NEXT_IN = True Then
        If Dir(CODER_VBP_FILE_NAME_2) <> "" Then
            Kill CODER_VBP_FILE_NAME_2
        End If
        FR = FreeFile
        Open CODER_VBP_FILE_NAME_2 For Binary As FR
            Put #FR, , VAR_STRING
        Close FR
    End If
End Sub
Private Sub MNU_FOLDER_DIGITAL_STILL_CAMERA_Click()
    ' Beep
    Me.WindowState = vbMinimized
    Dim FOLDER_Path
    FOLDER_Path = "D:\DSC\2015+SONY\2020 CyberShot HX60V\JPG"
    Dir1.Path = FOLDER_Path
    Dim TD
    For r = 0 To Dir1.ListCount - 1
        FIND_PATH_1 = Dir1.List(r)
        FIND_PATH_2 = Mid(FIND_PATH_1, InStrRev(FIND_PATH_1, "\"))
        If Mid(FIND_PATH_2, 1, 3) = "\20" Then
            TD = Dir1.List(r)
        End If
    Next
    FOLDER_Path = TD
    
    Shell "EXPLORER /SELECT, """ + FOLDER_Path + """", vbNormalFocus
    EXIT_TRUE = True
    ' Unload Me
    'End
End Sub
Private Sub MNU_CAMERA_VIDEO_FOLDER_4G_Click()
    Me.WindowState = vbMinimized
    Shell "EXPLORER \\4-asus-gl522vw\4_asus_gl522vw_10_1_samsung_1tb\DSC\2015+SONY_MP4\2020 CyberShot HX60V __ MP4", vbMaximizedFocus
End Sub
Private Sub MNU_ADMINISTRATOR_Click()
    'Call MNU_ADMINISTRATOR_Click
    ' Beep
    MNU_ADMINISTRATOR.Caption = "ADMINISTRATOR MODE ON"
    If GetSpecialfolder(38) = "" Then
        MNU_ADMINISTRATOR.Caption = Replace(MNU_ADMINISTRATOR.Caption, "MODE ON", "MODE OFF")
    End If
End Sub
'-------------------------------------------
'-------------------------------------------
'---------------------------------------------------
'Public Type SHITEMID
'    cb As Long
'    abID As Byte
'End Type
'Private Type ITEMIDLIST
'    mkid As SHITEMID
'End Type
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
'-------------------------------------------
'-------------------------------------------
Public Function GetSpecialFolder_Show_Script_Debug(CSIDL As Long) As String
Dim r As Long
On Error Resume Next
For r = 0 To 120
    If Trim(GetSpecialfolder(r)) <> "" Then
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




Private Sub Lbl_Title_Click()

FR1 = FreeFile
Open TEXT_PATH_1 + "\" + OIP$ + " Loads.txt" For Input As #FR1
    Line Input #FR1, A1$
    Line Input #FR1, B1$
    Line Input #FR1, C1$
Close #FR1

SetTrueToLoadLast = True
Call Label1_Click(0)

End Sub

Sub SaveLoggs()

fr2 = FreeFile
Open TEXT_PATH_1 + "\" + OIP$ + " Loads.txt" For Output As fr2
Print #fr2, A1$
Print #fr2, B1$
Print #fr2, C1$
Close fr2

fr2 = FreeFile
Open TEXT_PATH_1 + "\" + OIP$ + " Loads Logg.txt" For Append As fr2
Print #fr2, A1$
Print #fr2, B1$
Print #fr2, C1$
Close fr2

fr2 = FreeFile
Open TEXT_PATH_1 + "\" + OIP$ + " Loads Logg 02.txt" For Append As fr2
Print #fr2, Format$(Now, "dd-mm-yyyy hh:mm:ss ") + A1$ + B1$
Close fr2

End Sub

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  'IsIDE = False
  'Exit Function
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************


Function FindWinPart_SEARCHER_hWnd_TO_EXE(SEARCH_STRING) As Long
    
    Dim Text_TEMP_ER As String
    Dim test_hwnd As Long, _
        test_pid As Long, _
        test_thread_id As Long
    
    Dim cText As String
    Dim Huge
    
    'Find the first window
    test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
    Huge = 0
    Do While test_hwnd <> 0
        
        Text_TEMP_ER = GetFileFromHwnd(test_hwnd)
        'If InStr(GetWindowTitle(test_hWnd), SEARCH_STRING) > 0 Then
        If InStr(UCase(Text_TEMP_ER), UCase(SEARCH_STRING)) > 0 Then
            Huge = Huge + 1
            
            'Debug.Print Str(Huge) + "__" + Str(test_hWnd) + "__" + Text_TEMP_ER
            'Debug.Print Str(Huge) + "__" + Str(test_hWnd) + "__" + GetWindowTitle(test_hWnd)
            'Debug.Print Str(Huge) + "__" + Str(test_hWnd) + "__" + GetWindowClass(test_hWnd)
            'Debug.Print vbCrLf
            'EXIT SUB
            Huge = Huge
        End If
            
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    
    Loop
End Function


Function FindWinPart_SEARCHER(SEARCH_STRING) As Long
    FindWinPart_SEARCHER = False
    
    Dim test_hwnd As Long, _
        test_pid As Long, _
        test_thread_id As Long
    
    Dim cText As String
    Dim CLASS_NAME As String
    Dim XGO
    Dim CLASS_NAME_______________

    'Find the first window
    test_hwnd = FindWindow2(ByVal 0&, ByVal 0&)
    Do While test_hwnd <> 0
        
        CLASS_NAME_______________ = GetWindowClass(test_hwnd)
        If InStr(UCase(CLASS_NAME_______________), UCase(SEARCH_STRING)) > 0 Then XGO = True
        If InStr(UCase(GetWindowTitle(test_hwnd)), UCase(SEARCH_STRING)) > 0 Then XGO = True
        If XGO = True Then
            FindWinPart_SEARCHER = test_hwnd: Exit Function
        End If
            
    'retrieve the next window
    test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    
    Loop
End Function






Public Function GetHWndFromProcess(p_lngProcessId As Long) As Long
    Dim lngDesktop As Long
    Dim lngChild As Long
    Dim lngChildProcessID As Long
    
    On Error Resume Next
    
    'get the hWnd of the desktop
    lngDesktop = GetDesktopWindow()
    lngDesktop = FindWindowDLL(ByVal 0&, ByVal 0&)

    lngChild = GetWindow(lngDesktop, GW_CHILD)
    
    Do While lngChild <> 0
    
        'get the ThreadProcessID of the window
        Call GetWindowThreadProcessId(lngChild, lngChildProcessID)
        
        'if ProcessId matches the parameter passed
        'then return that value
        If lngChildProcessID = p_lngProcessId Then
            GetHWndFromProcess = lngChild
            Exit Do
        End If
        
        'not found, continue enumeration
        lngChild = GetWindow(lngChild, GW_HWNDNEXT)
    Loop
End Function


Function GetParentTitle(ByVal Handle As Long) As String
   Dim i As Long
   Dim j As Long, TTx1 As String
   i = Handle
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
      i = j
   TTx1 = GetWindowTitle(i)
   GetParentTitle = TTx1

End Function

Function GetWindowTitle(ByVal hWnd As Long) As String
   Dim l As Long
   Dim s As String
   l = GetWindowTextLength(hWnd)
   s = Space(l + 1)
   GetWindowText hWnd, s, l + 1
   GetWindowTitle = Left$(s, l)
End Function
Function GetWindowClass(ByVal hWnd As Long) As String
    Dim ret As Long, sText As String
    sText = Space(255)
    ret = GetClassName(hWnd, sText, 255)
    sText = Left$(sText, ret)
   GetWindowClass = sText
End Function



Private Sub FindCursor(x, y)

Dim P As POINTAPI

GetCursorPos P
'   return x and y co-ordinate
x = P.x ' / GetSystemMetrics(0) * Screen.Width
'   for current cursor position
y = P.y '/ GetSystemMetrics(1) * Screen.Height

End Sub



Sub SET_MENU_PADD_WORK()

Dim i_Menu_Count, i_Form_Counter
Dim i_Menu_Not_Visa_Count

Dim Control As Control, Label_44, LABEL_48

Dim R_NEXT

Dim Text_Checker_Form_Menu As String

Dim MENU_ITEM_VAR
Dim i

For i = 0 To Forms.Count - 1
    
    For Each Control In Forms(i).Controls
        If InStr(UCase(Control.Name), "MNU_") > 0 Then
            If Control.Visible = True Then
                i_Menu_Count = i_Menu_Count + 1
            End If
            i_Menu_Not_Visa_Count = i_Menu_Not_Visa_Count + 1

        End If
    Next
Next

i_Menu_Not_Visa_Count = i_Menu_Not_Visa_Count - i_Menu_Count

VAR_STRING = "Menu Item Count = " + Str(i_Menu_Count) + " &&" + Str(i_Menu_Not_Visa_Count) + " Not Visible"
' Mnu_Menu_Item_Count.Caption = "Menu Item Count = " + Str(i_Menu_Count) + " &&" + Str(i_Menu_Not_Visa_Count) + " Not Visible"


'Mnu_Form_Count.Caption = "Form Counter = " + Str(Forms.Count - 1) '  + " Really 7"
'Mnu_Form_Count.Visible = False

i_Menu_Count = 0



For i = 0 To Forms.Count - 1
    Text_Checker_Form_Menu = ""
    frmListMenu.GetMenuInfo_Not_Indented GetMenu(Forms(i).hWnd), 0, "", Text_Checker_Form_Menu
    Text_Checker_Form_Menu = UCase(Text_Checker_Form_Menu)
    For Each Control In Forms(i).Controls
        If InStr(UCase(Control.Name), "MNU_") > 0 Then
            MENU_ITEM_VAR = Replace(Control.Caption, "[__ ", "")
            MENU_ITEM_VAR = Replace(MENU_ITEM_VAR, " __]", "")
            MENU_ITEM_VAR = UCase(Trim(MENU_ITEM_VAR))
            If InStr(Text_Checker_Form_Menu, "SUB MENU ----" + MENU_ITEM_VAR) = 0 Then
                
                'i_Menu_Count = i_Menu_Count + 1
                If InStr(Trim(Control.Caption), "[__ ") = 0 Then
                    Label_44 = Trim(Control.Caption)
                    'LABEL_48 = Replace(LABEL_44, " ", "_")
                    LABEL_48 = Label_44
                    LABEL_48 = Replace(LABEL_48, "___", "__")
                    LABEL_48 = "[__ " + LABEL_48 + " __]"
                    LABEL_48 = Replace(LABEL_48, "[__ [__ ", "[__ ")
                    LABEL_48 = Replace(LABEL_48, " __] __]", " __]")
                    If LABEL_48 <> Label_44 Then
                        Control.Caption = LABEL_48
                    End If
                End If
            End If
        End If
    Next
Next

'Stop

''MNU_BRing_Front
''---------------
'i_Form_Counter = Forms.Count
'i_Form_Counter = 0
''for each f
''For i = 0 To Forms.Count - 1
''    Load Forms(i)
''    Show Forms(i)
''Next
'
'Dim Form As Form
'i_Form_Counter = 0
'For Each Form In Forms
'    i_Form_Counter = i_Form_Counter + 1
'    Load Form
'    Form.Show
'    Show Form
'    'Set Form = Nothing
'Next Form
'
'i_Form_Counter = 0
'For Each Form In Forms
'    i_Form_Counter = i_Form_Counter + 1
'    Load Form
'    Form.Show
'    'Set Form = Nothing
'Next Form

i_Form_Counter = Forms.Count - 1
Me.Refresh
End Sub


Sub SET_CONTROLS_COUNT()


Dim Control As Control, Label_44, LABEL_48

Dim R_NEXT

Dim Text_Checker_Form_Menu As String

Dim i

Dim CONTROL_THIS_FORM_V_
Dim CONTROL_THIS_FORM_V_PLUS
Dim CONTROL_THIS_FORM_VN
Dim CONTROL_THIS_FORM_TOT_V_
Dim CONTROL_THIS_FORM_TOT_VN
Dim CONTROL_THIS_FORM_TOT_V__PLUS

CONTROL_THIS_FORM_V_ = 0
CONTROL_THIS_FORM_V_PLUS = 0
CONTROL_THIS_FORM_VN = 0
CONTROL_THIS_FORM_TOT_V_ = 0
CONTROL_THIS_FORM_TOT_VN = 0
CONTROL_THIS_FORM_TOT_V__PLUS = 0

Dim XXER

On Error Resume Next

For Each Form In Forms
    For Each Control In Form.Controls
        Err.Clear
        XXER = Control.Visible
        If Err.Number > 0 Then XXER = False
        If XXER = True Then
            If Form.hWnd = Me.hWnd Then
            CONTROL_THIS_FORM_V_ = CONTROL_THIS_FORM_V_ + 1
            End If
            CONTROL_THIS_FORM_TOT_V_ = CONTROL_THIS_FORM_TOT_V_ + 1
        Else
            If Form.hWnd = Me.hWnd Then
                CONTROL_THIS_FORM_VN = CONTROL_THIS_FORM_VN + 1
            End If
                CONTROL_THIS_FORM_TOT_VN = CONTROL_THIS_FORM_TOT_VN + 1
        End If
    Next
Next
On Error GoTo 0

CONTROL_THIS_FORM_TOT_V__PLUS = CONTROL_THIS_FORM_TOT_V_ + CONTROL_THIS_FORM_TOT_VN
CONTROL_THIS_FORM_V_PLUS = CONTROL_THIS_FORM_V_ + CONTROL_THIS_FORM_VN

' Mnu_Control_Item_Count.Caption = "Control Item Count All Form*" + Trim(Str(Forms.Count)) + " = " + Str(CONTROL_THIS_FORM_TOT_V__PLUS) + " &&" + Str(CONTROL_THIS_FORM_TOT_VN) + " Not Visible && Main Form" + Str(CONTROL_THIS_FORM_V_PLUS) + " &&" + Str(CONTROL_THIS_FORM_VN) + " Not Visible"
Mnu_Control_Item_Count.Caption = "Control Item Count All Form*" + Trim(Str(Forms.Count)) + " =" + Str(CONTROL_THIS_FORM_TOT_V__PLUS) + " &&" + Str(CONTROL_THIS_FORM_TOT_VN) + " Not Visible &&" + Str(CONTROL_THIS_FORM_V_PLUS) + " This Form" '+ " &&" + Str(CONTROL_THIS_FORM_VN) + " Not Visible"
End Sub



Sub SET_LABEL_PADD_WORK()

Dim Control As Control
Dim i

On Error Resume Next

For i = 0 To Forms.Count - 1
    For Each Control In Forms(i).Controls
        If InStr(Trim(Control.Caption), "/n") > 0 Then
            Control.Caption = Replace(Control.Caption, "/n", vbCrLf)
        End If
    Next
Next

Me.Refresh
End Sub



Function GetParentHwnd(ByVal ReturnParent As Long) As String
   Dim i As Long
   Dim j As Long
   Dim k As Long
   i = ReturnParent
   If ReturnParent Then
      Do While i <> 0
         k = j
         j = i
         i = GetParent(i)
      Loop
    i = j
    End If
    GetParentHwnd = i
End Function



Private Sub Timer_SUB_CODE_LOAD_Timer()
    Timer_SUB_CODE_LOAD.Enabled = False
    Call SubCode
        
    Combo1.SetFocus
        
        
    ReDim TEXT_PATH_4_ARRAY(1)
    TEXT_PATH_4 = TEXT_PATH_2 + "\" + OIP$ + " CRC SPEED BROKEN LINK.txt"
    If Dir(TEXT_PATH_4) <> "" Then
        FR1 = FreeFile
        Open TEXT_PATH_4 For Input As FR1
        T_P_4_A = 0
        Do
            T_P_4_A = T_P_4_A + 1
            ReDim Preserve TEXT_PATH_4_ARRAY(T_P_4_A)
            If Not EOF(FR1) Then
                Line Input #FR1, TEXT_PATH_4_ARRAY(T_P_4_A)
            End If
            Loop Until EOF(FR1)
        Close FR1
    End If
        
        
    TEXT_PATH_8 = TEXT_PATH_4
'    FR4 = FreeFile
'    Open TEXT_PATH_4 For Output As FR4
    '  HERE GET OPEN AT Timer2_Timer
    ' SEARCH -- Open TEXT_PATH_4 For Output As FR1
    
    XXT4 = 1
    
    Timer2.Enabled = True
    Timer2.Interval = 18
    Timer2.Interval = 1
    Timer3.Enabled = True
    Timer3.Interval = 1
    'Timer3.INT

End Sub









'Type RGBColor
'     Red As Byte
'     Green As Byte
'     Blue As Byte
'End Type
'
'Type HSBColor
'     Hue As Double
'     Saturation As Double
'     Brightness As Double
'End Type
'Next, here's a brief pseudocode explanation of the procedure used by the RGB-to-HSB algorithm:
'
'
'Set a Delta variable equal to [Max(r,g,b) - Min(r,g,b)]
'* Then Brightness = Max(r,g,b) * 100 / 255
'* If the color is (00,00,00) (black), then Saturation = 0, and h = -1; otherwise:
'   Saturation = 255 * Delta / Max(R, g, b)
'   Case Max(r,g,b) is equal to the value of
'                + Red : Set h = (Green - Blue) / Delta
'                + Green : h = 2 + (Blue - Red) / Delta
'                + Blue : h = 4 + (Red - Green) / Delta
'    * Hue = h * 60 , if h small then 0 , we have Hue = h + 360
'Now , here 's the VB code for the RGB-to-HSB conversion:

'Private Function RGBToHSB(rgb As RGBColor) As HSBColor
'Dim minRGB, maxRGB, Delta As Double
'Dim h, s, B As Double
'     h = 0
'     minRGB = Min(Min(rgb.Red, rgb.Green), rgb.Blue)
'     maxRGB = Max(Max(rgb.Red, rgb.Green), rgb.Blue)
'     Delta = (maxRGB - minRGB)
'     B = maxRGB
'     If (maxRGB <> 0) Then
'          s = 255 * Delta / maxRGB
'     Else
'          s = 0
'     End If
'     If (s <> 0) Then
'          If rgb.Red = maxRGB Then
'               h = (CDbl(rgb.Green) - CDbl(rgb.Blue)) / Delta
'          Else
'               If rgb.Green = maxRGB Then
'                    h = 2 + (CDbl(rgb.Blue) - CDbl(rgb.Red)) / Delta
'               Else
'                    If rgb.Blue = maxRGB Then
'                         h = 4 + (CDbl(rgb.Red) - CDbl(rgb.Green)) / Delta
'                    End If
'               End If
'          End If
'     Else
'          h = -1
'     End If
'     h = h * 60
'     If h < 0 Then h = h + 360
'     RGBToHSB.Hue = h
'     RGBToHSB.Saturation = s * 100 / 255
'     RGBToHSB.Brightness = B * 100 / 255
'End Function
'Of course, you also need to go the other direction, changing from HSB to RGB. Note that:
'When Max = Red, abs((Green - Blue) / Delta) is always smaller than 1, or -1 < (Green - Blue) / Delta.
'When Max = Green, abs((Blue - Red) / Delta) is always smaller than 1, or 1 < 2 + (Blue - Red) / Delta <3.
'When Max = Blue, abs((Red - Green) / Delta) is always smaller than 1, or 3 < 4 + (Red - Green) / Delta <5.
'So, in pseudocode, the procedure is:
'
'Set h = Hue / 60
'    * Then s = Saturation * 255 / 100
'    * And b = Brightness * 255 / 100
'    * Max(r,g,b) will equal b.
'    * If s = 0  then the color is black (00,00,00), otherwise:
'          - Set Delta = s * Max(r,g,b) / 255
'          - Compare the value of h with 3,1,-1 in turn to find out the Values of Red, Green and Blue.
'In VB6, the HSB-to-RGB algorithm looks like this:

Private Function HSBToRGB(hsb As HSBColor) As RGBColor
Dim maxRGB, Delta As Double
Dim h, s, B As Double
     h = hsb.Hue / 60
     s = hsb.Saturation * 255 / 100
     B = hsb.Brightness * 255 / 100
     maxRGB = B
     If s = 0 Then
          HSBToRGB.Red = 0
          HSBToRGB.Green = 0
          HSBToRGB.Blue = 0
     Else
          Delta = s * maxRGB / 255
          If h > 3 Then
               HSBToRGB.Blue = CByte(Round(maxRGB))
               If h > 4 Then
                    HSBToRGB.Green = CByte(Round(maxRGB - Delta))
                    HSBToRGB.Red = CByte(Round((h - 4) * Delta)) + HSBToRGB.Green
               Else
                    HSBToRGB.Red = CByte(Round(maxRGB - Delta))
                    HSBToRGB.Green = CByte(HSBToRGB.Red - Round((h - 4) * Delta))
               End If
          Else
               If h > 1 Then
                    HSBToRGB.Green = CByte(Round(maxRGB))
                    If h > 2 Then
                         HSBToRGB.Red = CByte(Round(maxRGB - Delta))
                         HSBToRGB.Blue = CByte(Round((h - 2) * Delta)) + HSBToRGB.Red
                    Else
                         HSBToRGB.Blue = CByte(Round(maxRGB - Delta))
                         HSBToRGB.Red = CByte(HSBToRGB.Blue - Round((h - 2) * Delta))
                    End If
               Else
                    If h > -1 Then
                         HSBToRGB.Red = CByte(Round(maxRGB))
                         If h > 0 Then
                              HSBToRGB.Blue = CByte(Round(maxRGB - Delta))
                              HSBToRGB.Green = CByte(Round(h * Delta)) + HSBToRGB.Blue
                         Else
                              HSBToRGB.Green = CByte(Round(maxRGB - Delta))
                              HSBToRGB.Blue = CByte(HSBToRGB.Green - Round(h * Delta))
                         End If
                    End If
               End If
          End If
     End If
End Function
'Duc Nguyen
'----
'Algorithm to Switch Between RGB and HSB Color Values
'http://www.devx.com/tips/Tip/41581
'----


Private Function Menu_Height()
 
    MenuInfo.cbSize = Len(MenuInfo)
    
    If GetMenuBarInfo(Me.hWnd, OBJID_MENU, 0, MenuInfo) Then
   
        With MenuInfo.rcBar
       
'            Debug.Print "Left: " & CStr(.Left)
'            Debug.Print "Right: " & CStr(.Right)
'            Debug.Print "Top: " & CStr(.Top)
'            Debug.Print "Bottom: " & CStr(.Bottom)
            Menu_Height = CStr(.Bottom) - CStr(.Top)
        End With
    End If

    If Menu_Height <> ARCHIVE_Menu_Height Then
        TIMER_TO_RESIZE.Enabled = True
    End If
    
    ARCHIVE_Menu_Height = Menu_Height
   
End Function

Private Sub TIMER_TO_RESIZE_Timer()
    TIMER_TO_RESIZE.Enabled = False
    'NOT_RESIZE_EVENTER = False
    Call Form_Resize
    'NOT_RESIZE_EVENTER = True

End Sub

