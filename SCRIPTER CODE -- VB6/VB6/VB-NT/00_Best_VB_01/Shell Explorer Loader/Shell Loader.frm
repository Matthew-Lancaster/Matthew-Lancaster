VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   Caption         =   "EXPLORER LOADER"
   ClientHeight    =   5928
   ClientLeft      =   132
   ClientTop       =   780
   ClientWidth     =   12312
   Icon            =   "Shell Loader.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5928
   ScaleWidth      =   12312
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
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
      Left            =   4092
      TabIndex        =   117
      Top             =   912
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
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
      Left            =   4092
      TabIndex        =   116
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
      Left            =   4092
      TabIndex        =   85
      Top             =   948
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
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
      Left            =   4092
      TabIndex        =   84
      Top             =   948
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
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
      Left            =   4092
      TabIndex        =   53
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
      Index           =   47
      Left            =   4092
      TabIndex        =   52
      Top             =   948
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
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
      Left            =   4092
      TabIndex        =   36
      Top             =   804
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
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
   Begin VB.Menu Mnu_LoadFolder 
      Caption         =   "LOAD FOLDER LINK"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function FindWindowDLL Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

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
    If KeyCode = 38 Then
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
tt = String(Len(Trim(Str(Combo1.ListCount))), "0")
STRING_COMBO_NUMBER = Format(X_COUNTER, tt) + " of " + Format(Combo1.ListCount, tt)
If Len(STRING_COMBO_NUMBER) <> Len(Lbl_COMBO_NUMBER.Caption) Then
    Lbl_COMBO_NUMBER.Caption = STRING_COMBO_NUMBER
    
    Lbl_COMBO_NUMBER.AutoSize = True
    Lbl_COMBO_NUMBER.Refresh
    Lbl_COMBO_NUMBER.AutoSize = False
    Lbl_COMBO_NUMBER.Alignment = 2
    Lbl_COMBO_NUMBER.Height = Lbl_Title.Height
    
    Lbl_COMBO_NUMBER.Left = Lbl_Title.Width + 20
    Lbl_COMBO_NUMBER.Width = Lbl_COMBO_NUMBER.Width + 200
    Combo1.Left = Lbl_COMBO_NUMBER.Left + Lbl_COMBO_NUMBER.Width + 20
    Combo1.Width = Form1.Width - Combo1.Left - 100
End If
Lbl_COMBO_NUMBER.Caption = STRING_COMBO_NUMBER

End Sub


Private Sub Form_Resize()

If X_Y_DONE_ONCE = True Then Exit Sub

On Error Resume Next
Form1.Width = Form1_Width
Form1.Height = Form1_Height
'Center Screen
Form1.Top = 0 'Screen.Height / 2 - Form1.Height / 2
Form1.Left = Screen.Width / 2 - Form1.Width / 2

If Form1.Width = Form1_Width And Form1.Height = Form1_Height Then X_Y_DONE_ONCE = True


End Sub

Private Sub Form_Load()

ReDim A4$(500), B4$(500), C4$(500)

If GetComputerName = GetComputerName Then
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
    FontSizez = 9
    FontSizez_2 = FontSizez '     -- THE SPECIAL FOLDERIN GET HERE
End If


If GetComputerName = "5-ASUS-P2520LA" Then
    FontSizez = 6.3
    FontSizez_2 = FontSizez '     -- THE SPECIAL FOLDERIN GET HERE
End If

If GetComputerName = "7-ASUS-GL522VW" Then
    FontSizez = 7.4
    FontSizez_2 = FontSizez '     -- THE SPECIAL FOLDERIN GET HERE
End If

If GetComputerName = "8-MSI-GP62M-7RD" Then
    FontSizez = 7.5
    FontSizez_2 = FontSizez '     -- THE SPECIAL FOLDERIN GET HERE
End If


Call SET_UP_PULIC_FSO

NAME_FORM = App.EXEName

Me.BackColor = &HA5FCF3
Call FindRGB(Me.BackColor)
Me.BackColor = RGB(R_RGB - 180, G_RGB - 180, B_RGB - 100)


Dim i As String
If App.PrevInstance = True And IsIDE = True Then
    'i = FindWindow(vbNullString, Me.Caption)
    i = FindWinPart_SEARCHER("VB KEEP RUNNER")
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
tt1 = InStr(OIP$, "Shell")
tt2 = InStr(tt1, OIP$, "Loader")
OIP2$ = Mid$(OIP$, tt1 + 6, tt2 - tt1 - 7)
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
If TRUE_GO = True Then
    Combo1.AddItem C1$
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
MNU_VERSION.Caption = "Ver_2019_" + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision)) ' + " _ Matt.Lan@btinternet.com"


Form1.Top = 0
Form1.Left = 0

Call SET_MENU_PADD_WORK

Call FormStart.FormStartLoader

Timer_SUB_CODE_LOAD.Enabled = True

End Sub

Private Sub Lbl_COMBO_NUMBER_Click()
'Lbl_COMBO_NUMBER.CAPTION
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
    For R = ScanPath.ListView1.ListItems.Count To 1 Step -1
        If Mid(ScanPath.ListView1.ListItems.Item(R), 1, 3) = "XP " Then
            If IsIDE = False Then ScanPath.ListView1.ListItems.Remove (R)
        End If
    Next
End If

If IV = 5.1 Then
    For R = ScanPath.ListView1.ListItems.Count To 1 Step -1
        A1$ = ScanPath.ListView1.ListItems.Item(R).SubItems(1)
        B1$ = ScanPath.ListView1.ListItems.Item(R)
        If InStr((UCase(B1$)), "PROGRAM FILES (X86)") > 0 Then
            Call GETSHORTLINK(A1$ + B1$)
            D1$ = txtTargetPath
            If FSO.FileExists(D1$) = False And FSO.FolderExists(D1$) = False Then
                If IsIDE = False Then
                    ScanPath.ListView1.ListItems.Remove (R)
                End If
            End If
        End If
    Next
End If


For R = ScanPath.ListView1.ListItems.Count To 0 Step -1
    ' A1$ = ScanPath.ListView1.ListItems.Item(R).SubItems(1)
    If R = 0 Then Exit For
    B1$ = ScanPath.ListView1.ListItems.Item(R)
    If InStr(B1$, "Videos.lnk") > 0 Then
        ScanPath.ListView1.ListItems.Remove (R)
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

tg = x

x = x + Lbl_Title.Height + LABEL_GAP
tx = x

w = 0

RD = 0

DoEvents

SH = Screen.Height - 4000 'higer = smaller
SH = Me.Height - 1500
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
For R = 1 To ScanPath.ListView1.ListItems.Count + 100
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

x = tx
RX = 0
rt = 0
For R = 1 To ScanPath.ListView1.ListItems.Count
    
    RX = RX + 1
    rt = rt + 1
    
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
        rt = rt - 1
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
        
        Label1(RX).Caption = ttg$
        Label1(RX).Alignment = 2
    Else
        If InStr(A1$, "SPECIAL") > 0 Then
            Label1(RX).FontSize = FontSizez_2
        End If
        Label1(RX).Caption = Format$(rt, "000") + ". " + ttg$
        
    End If
    
    If InStr(Label1(RX), ":\") > 0 Then
        SET_GO = False
        'If FSO.FolderExists(Mid(Label1(RX), 5)) = True Then SET_GO = True
        If FSO.FolderExists(Mid(Label1(RX), InStr(Label1(RX), ":") - 1)) = True Then SET_GO = True
        If SET_GO = True Then
            LABEL_BACKCOLOR_VAR = Label3.BackColor - RGB(40, 40, 40)
        Else
            LABEL_BACKCOLOR_VAR = Label3.BackColor
        End If
    End If
    If InStr(Label1(RX), "\\") > 0 Then
        LABEL_BACKCOLOR_VAR = Label2.BackColor
    End If
    
    Label1(RX).BackColor = LABEL_BACKCOLOR_VAR
    
    If InStr(LABLE_BACKCOLOR_SET, LABEL_BACKCOLOR_VAR) = 0 Then
        LABLE_BACKCOLOR_SET = LABLE_BACKCOLOR_SET + Hex(Label1(RX).BackColor)
    End If
    
    x = x + Label1(RX).Height + LABEL_GAP
    fheight = Label1(RX).Top + Label1(RX).Height + 420
    If fheight > fheightx Then fheightx = fheight
    RD = RD + 1



Next
DoEvents

x = tx
td = 29
xgag = 0
Dim xy(20)

For R = 1 To ScanPath.ListView1.ListItems.Count
    If Label1(R - 1).Top > SH Then
        If tig = 0 Then tig = R
        x = tx
        xgag = xgag + 1
        wdt = 0
        For rt = R - 1 To R - tig Step -1
        If Label1(rt).Width > wdt Then wdt = Label1(rt).Width
        Next
        xy(xgag) = wdt + 150
        tig2 = R
    End If
    Label1(R).Top = x
    x = x + Label1(R).Height + LABEL_GAP
Next

xgag = xgag + 1
wdt = 0
For rt = tig2 To R - 1
    If Label1(rt).Width > wdt Then wdt = Label1(rt).Width
Next
xy(xgag) = wdt + 150

For R = 1 To ScanPath.ListView1.ListItems.Count
    fw = Label1(R).Width
    If fw > fw2 Then fw2 = fw + 200
Next
fw2 = fw2

xgag = 1: xgax2 = 0
For R = 1 To ScanPath.ListView1.ListItems.Count
    Label1(R).AutoSize = False
    
    DoEvents
    xxb = Label1(R - 1).Top
    xxb = Form1.Height
    
    If Label1(R - 1).Top > SH Then
        xgax2 = xgax2 + 1
        xgag = 0
        For rs2 = 1 To xgax2
            xgag = xgag + xy(rs2)
        Next
    End If
    
    Label1(R).Width = xy(xgax2 + 1) - 20
    Label1(R).Left = xgag
    
Next

fheightx = 0
For rt = 1 To ScanPath.ListView1.ListItems.Count
    fheight = Label1(rt).Top + Label1(rt).Height
    If fheight > fheightx Then
        fheightx = fheight
    Else
        Exit For
    End If
Next


Form1_Width = Label1(R - 1).Left + Label1(R - 1).Width + 280
If Form1_Width > Screen.Width Then Form1_Width = Screen.Width - 120
Form1_Height = fheightx + 900

Lbl2.Width = Form1.Width
Lbl_Title.Top = tg
Lbl_Title.Left = 0
Combo1.Top = tg
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

Lbl3.Top = fheight - 410
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
        Debug.Print Control.Caption
        Debug.Print XX2
        Debug.Print R_RGB
        Debug.Print G_RGB
        Debug.Print B_RGB
        
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
    For R = 1 To UBound(TEXT_PATH_4_ARRAY)
        If InStr(TEXT_PATH_4_ARRAY(R), WxHex$) > 0 Then
            FOLEDERNAME = TEXT_PATH_4_ARRAY(R + 1)
            FOLEDERNAME = TEXT_PATH_4_ARRAY(R + 2)
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
            Print #FR1, WxHex$
            Print #FR1, A1$ + B1$
            Print #FR1, D1$
        End If
        Exit Sub
    End If

    FOLEDERNAME = "C:\Program Files (x86)"
    If InStr(A1$ + "--", "Program Files (x86)" + "--") > 0 Then
        If FSO.FolderExists(FOLEDERNAME) = False Then
            Label1(XXT2).BackColor = RGB(147, 147, 147)
            X_COLOR_ARRAY(XXT2) = Form1.Label1(XXT2).BackColor
            WxHex$ = Hex(m_CRC.CalculateFile(A1$ + B1$))
            Print #FR1, WxHex$
            Print #FR1, FOLEDERNAME
            Print #FR1, D1$
            Exit Sub
        End If
    End If

    If Mid(D1$, 2, 2) = ":\" Or Mid(D1$, 1, 2) = "\\" Then
        If FSO.FileExists(D1$) = False And FSO.FolderExists(D1$) = False Then
            Label1(XXT2).BackColor = RGB(240, 127, 127) 'QBColor(12)
            X_COLOR_ARRAY(XXT2) = Form1.Label1(XXT2).BackColor
            WxHex$ = Hex(m_CRC.CalculateFile(A1$ + B1$))
            Print #FR1, WxHex$
            Print #FR1, A1$ + B1$
            Print #FR1, D1$
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

Private Sub Label1_Click(Index As Integer)

'Beep

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

'----
'Shell Commands List for Windows 10 | Tutorials
'https://www.tenforums.com/tutorials/3109-shell-commands-list-windows-10-a.html
'----
If LoadFolder = True And InStr(C1$, "NETWORK COMPUTER NAME") > 0 Then

    Call SaveLoggs

    Shell "explorer shell:NetworkPlacesFolder", vbNormalFocus
    End
End If

' ------------------
' THIS NOT OFTEN RUN
' SEE BELOW
' ------------------
If LoadFolder = True Then

    Call SaveLoggs
    
    Shell "explorer /e, " + A1$, vbNormalFocus
    End
End If


Call SaveLoggs

Call FormStart.LabelClick(Index)

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


Private Sub MNU_VB_FOLDER_Click()
    Shell "EXPLORER /SELECT, " + App.Path + "\" + App.EXEName + ".VBP", vbMaximizedFocus
End Sub

Private Sub MNU_VB_ME_Click()

VBPATH = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
If Dir(VBPATH) = "" Then
    VBPATH = "C:\Program Files (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
End If

Shell VBPATH + " """ + App.Path + "\" + App.EXEName + ".VBP""", vbNormalFocus
End

End Sub



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
    FR1 = FreeFile
    TEXT_PATH_4 = TEXT_PATH_2 + "\" + OIP$ + " CRC SPEED BROKEN LINK.txt"
    If Dir(TEXT_PATH_4) <> "" Then
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
        
    FR1 = FreeFile
    Open TEXT_PATH_4 For Output As FR1
    
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

