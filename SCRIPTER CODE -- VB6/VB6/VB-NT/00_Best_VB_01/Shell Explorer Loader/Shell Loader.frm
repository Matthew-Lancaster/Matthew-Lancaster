VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   Caption         =   "NotePad Loader"
   ClientHeight    =   5928
   ClientLeft      =   192
   ClientTop       =   840
   ClientWidth     =   12324
   Icon            =   "Shell Loader.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5928
   ScaleWidth      =   12324
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.FileListBox File1 
      Height          =   456
      Left            =   10800
      TabIndex        =   265
      Top             =   276
      Visible         =   0   'False
      Width           =   504
   End
   Begin VB.DirListBox Dir1 
      Height          =   288
      Left            =   10248
      TabIndex        =   264
      Top             =   324
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Timer Timer_KEY_CODE 
      Interval        =   1
      Left            =   1296
      Top             =   516
   End
   Begin VB.Timer Timer_SUB_CODE_LOAD 
      Interval        =   1
      Left            =   720
      Top             =   516
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6192
      Top             =   768
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4125
      TabIndex        =   166
      Text            =   "Combo1"
      Top             =   75
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   432
      Left            =   195
      TabIndex        =   19
      Top             =   4890
      Visible         =   0   'False
      Width           =   4965
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
      Left            =   7080
      TabIndex        =   17
      Top             =   5235
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      Height          =   320
      Left            =   6944
      TabIndex        =   263
      Top             =   80
      Width           =   400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   256
      Left            =   6330
      TabIndex        =   262
      Top             =   2850
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   255
      Left            =   6330
      TabIndex        =   261
      Top             =   3060
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   254
      Left            =   6330
      TabIndex        =   260
      Top             =   3270
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   253
      Left            =   6330
      TabIndex        =   259
      Top             =   3480
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   252
      Left            =   6330
      TabIndex        =   258
      Top             =   3690
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   251
      Left            =   6330
      TabIndex        =   257
      Top             =   3900
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   250
      Left            =   6330
      TabIndex        =   256
      Top             =   4110
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   249
      Left            =   6330
      TabIndex        =   255
      Top             =   4320
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   248
      Left            =   6345
      TabIndex        =   254
      Top             =   2640
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   247
      Left            =   6345
      TabIndex        =   253
      Top             =   2430
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   246
      Left            =   6345
      TabIndex        =   252
      Top             =   2220
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   245
      Left            =   6345
      TabIndex        =   251
      Top             =   2010
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   244
      Left            =   6345
      TabIndex        =   250
      Top             =   1800
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   243
      Left            =   6345
      TabIndex        =   249
      Top             =   1590
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   242
      Left            =   6345
      TabIndex        =   248
      Top             =   1380
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   241
      Left            =   6345
      TabIndex        =   247
      Top             =   1155
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   240
      Left            =   8055
      TabIndex        =   246
      Top             =   1155
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   239
      Left            =   8055
      TabIndex        =   245
      Top             =   1380
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   238
      Left            =   8055
      TabIndex        =   244
      Top             =   1590
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   237
      Left            =   8055
      TabIndex        =   243
      Top             =   1800
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   236
      Left            =   8055
      TabIndex        =   242
      Top             =   2010
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   235
      Left            =   8055
      TabIndex        =   241
      Top             =   2220
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   234
      Left            =   8055
      TabIndex        =   240
      Top             =   2430
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   233
      Left            =   8055
      TabIndex        =   239
      Top             =   2640
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   232
      Left            =   8040
      TabIndex        =   238
      Top             =   4320
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   231
      Left            =   8040
      TabIndex        =   237
      Top             =   4110
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   230
      Left            =   8040
      TabIndex        =   236
      Top             =   3900
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   229
      Left            =   8040
      TabIndex        =   235
      Top             =   3690
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   228
      Left            =   8040
      TabIndex        =   234
      Top             =   3480
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   227
      Left            =   8040
      TabIndex        =   233
      Top             =   3270
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   226
      Left            =   8040
      TabIndex        =   232
      Top             =   3060
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   225
      Left            =   8040
      TabIndex        =   231
      Top             =   2850
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   224
      Left            =   7380
      TabIndex        =   230
      Top             =   2730
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   223
      Left            =   7380
      TabIndex        =   229
      Top             =   2940
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   222
      Left            =   7380
      TabIndex        =   228
      Top             =   3150
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   221
      Left            =   7380
      TabIndex        =   227
      Top             =   3360
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   220
      Left            =   7380
      TabIndex        =   226
      Top             =   3570
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   219
      Left            =   7380
      TabIndex        =   225
      Top             =   3780
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   218
      Left            =   7380
      TabIndex        =   224
      Top             =   3990
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   217
      Left            =   7380
      TabIndex        =   223
      Top             =   4200
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   216
      Left            =   7395
      TabIndex        =   222
      Top             =   2520
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   215
      Left            =   7395
      TabIndex        =   221
      Top             =   2310
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   214
      Left            =   7395
      TabIndex        =   220
      Top             =   2100
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   213
      Left            =   7395
      TabIndex        =   219
      Top             =   1890
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   212
      Left            =   7395
      TabIndex        =   218
      Top             =   1680
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   211
      Left            =   7395
      TabIndex        =   217
      Top             =   1470
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   210
      Left            =   7395
      TabIndex        =   216
      Top             =   1260
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   209
      Left            =   7395
      TabIndex        =   215
      Top             =   1035
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   208
      Left            =   9840
      TabIndex        =   214
      Top             =   1125
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   207
      Left            =   9840
      TabIndex        =   213
      Top             =   1350
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   206
      Left            =   9840
      TabIndex        =   212
      Top             =   1560
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   205
      Left            =   9840
      TabIndex        =   211
      Top             =   1770
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   204
      Left            =   9840
      TabIndex        =   210
      Top             =   1980
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   203
      Left            =   9840
      TabIndex        =   209
      Top             =   2190
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   202
      Left            =   9840
      TabIndex        =   208
      Top             =   2400
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   201
      Left            =   9840
      TabIndex        =   207
      Top             =   2610
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   200
      Left            =   9825
      TabIndex        =   206
      Top             =   4290
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   199
      Left            =   9825
      TabIndex        =   205
      Top             =   4080
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   198
      Left            =   9825
      TabIndex        =   204
      Top             =   3870
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   197
      Left            =   9825
      TabIndex        =   203
      Top             =   3660
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   196
      Left            =   9825
      TabIndex        =   202
      Top             =   3450
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   195
      Left            =   9825
      TabIndex        =   201
      Top             =   3240
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   194
      Left            =   9825
      TabIndex        =   200
      Top             =   3030
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   193
      Left            =   9825
      TabIndex        =   199
      Top             =   2820
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   192
      Left            =   10485
      TabIndex        =   198
      Top             =   2940
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   191
      Left            =   10485
      TabIndex        =   197
      Top             =   3150
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   190
      Left            =   10485
      TabIndex        =   196
      Top             =   3360
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   189
      Left            =   10485
      TabIndex        =   195
      Top             =   3570
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   188
      Left            =   10485
      TabIndex        =   194
      Top             =   3780
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   187
      Left            =   10485
      TabIndex        =   193
      Top             =   3990
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   186
      Left            =   10485
      TabIndex        =   192
      Top             =   4200
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   185
      Left            =   10485
      TabIndex        =   191
      Top             =   4410
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   184
      Left            =   10500
      TabIndex        =   190
      Top             =   2730
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   183
      Left            =   10500
      TabIndex        =   189
      Top             =   2520
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   182
      Left            =   10500
      TabIndex        =   188
      Top             =   2310
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   181
      Left            =   10500
      TabIndex        =   187
      Top             =   2100
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   180
      Left            =   10500
      TabIndex        =   186
      Top             =   1890
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   179
      Left            =   10500
      TabIndex        =   185
      Top             =   1680
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   178
      Left            =   10500
      TabIndex        =   184
      Top             =   1470
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   177
      Left            =   10500
      TabIndex        =   183
      Top             =   1245
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   176
      Left            =   8790
      TabIndex        =   182
      Top             =   1245
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   175
      Left            =   8790
      TabIndex        =   181
      Top             =   1470
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   174
      Left            =   8790
      TabIndex        =   180
      Top             =   1680
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   173
      Left            =   8790
      TabIndex        =   179
      Top             =   1890
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   172
      Left            =   8790
      TabIndex        =   178
      Top             =   2100
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   171
      Left            =   8790
      TabIndex        =   177
      Top             =   2310
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   170
      Left            =   8790
      TabIndex        =   176
      Top             =   2520
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   169
      Left            =   8790
      TabIndex        =   175
      Top             =   2730
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   168
      Left            =   8775
      TabIndex        =   174
      Top             =   4410
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   167
      Left            =   8775
      TabIndex        =   173
      Top             =   4200
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   166
      Left            =   8775
      TabIndex        =   172
      Top             =   3990
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   165
      Left            =   8775
      TabIndex        =   171
      Top             =   3780
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   164
      Left            =   8775
      TabIndex        =   170
      Top             =   3570
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   163
      Left            =   8775
      TabIndex        =   169
      Top             =   3360
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   162
      Left            =   8775
      TabIndex        =   168
      Top             =   3150
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   161
      Left            =   8775
      TabIndex        =   167
      Top             =   2940
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0056E25E&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   272
      Index           =   160
      Left            =   4448
      TabIndex        =   165
      Top             =   4304
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   159
      Left            =   6855
      TabIndex        =   164
      Top             =   2535
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   158
      Left            =   6855
      TabIndex        =   163
      Top             =   2745
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   157
      Left            =   6855
      TabIndex        =   162
      Top             =   2955
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   156
      Left            =   6855
      TabIndex        =   161
      Top             =   3165
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   155
      Left            =   6855
      TabIndex        =   160
      Top             =   3375
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   154
      Left            =   6855
      TabIndex        =   159
      Top             =   3585
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   153
      Left            =   6855
      TabIndex        =   158
      Top             =   3795
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   152
      Left            =   6855
      TabIndex        =   157
      Top             =   4005
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   151
      Left            =   6870
      TabIndex        =   156
      Top             =   2325
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   150
      Left            =   6870
      TabIndex        =   155
      Top             =   2115
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   149
      Left            =   6870
      TabIndex        =   154
      Top             =   1905
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   148
      Left            =   6870
      TabIndex        =   153
      Top             =   1695
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   147
      Left            =   6870
      TabIndex        =   152
      Top             =   1485
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   146
      Left            =   6870
      TabIndex        =   151
      Top             =   1275
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   145
      Left            =   6870
      TabIndex        =   150
      Top             =   1065
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   144
      Left            =   6870
      TabIndex        =   149
      Top             =   840
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   143
      Left            =   8580
      TabIndex        =   148
      Top             =   840
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   142
      Left            =   8580
      TabIndex        =   147
      Top             =   1065
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   141
      Left            =   8580
      TabIndex        =   146
      Top             =   1275
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   140
      Left            =   8580
      TabIndex        =   145
      Top             =   1485
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   139
      Left            =   8580
      TabIndex        =   144
      Top             =   1695
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   138
      Left            =   8580
      TabIndex        =   143
      Top             =   1905
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   137
      Left            =   8580
      TabIndex        =   142
      Top             =   2115
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   136
      Left            =   8580
      TabIndex        =   141
      Top             =   2325
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   135
      Left            =   8565
      TabIndex        =   140
      Top             =   4005
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   134
      Left            =   8565
      TabIndex        =   139
      Top             =   3795
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   133
      Left            =   8565
      TabIndex        =   138
      Top             =   3585
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   132
      Left            =   8565
      TabIndex        =   137
      Top             =   3375
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   131
      Left            =   8565
      TabIndex        =   136
      Top             =   3165
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   130
      Left            =   8565
      TabIndex        =   135
      Top             =   2955
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   129
      Left            =   8565
      TabIndex        =   134
      Top             =   2745
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   128
      Left            =   8565
      TabIndex        =   133
      Top             =   2535
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   127
      Left            =   7905
      TabIndex        =   132
      Top             =   2415
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   126
      Left            =   7905
      TabIndex        =   131
      Top             =   2625
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   125
      Left            =   7905
      TabIndex        =   130
      Top             =   2835
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   124
      Left            =   7905
      TabIndex        =   129
      Top             =   3045
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   123
      Left            =   7905
      TabIndex        =   128
      Top             =   3255
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   122
      Left            =   7905
      TabIndex        =   127
      Top             =   3465
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   121
      Left            =   7905
      TabIndex        =   126
      Top             =   3675
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   120
      Left            =   7905
      TabIndex        =   125
      Top             =   3885
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   119
      Left            =   7920
      TabIndex        =   124
      Top             =   2205
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   118
      Left            =   7920
      TabIndex        =   123
      Top             =   1995
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   117
      Left            =   7920
      TabIndex        =   122
      Top             =   1785
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   116
      Left            =   7920
      TabIndex        =   121
      Top             =   1575
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   115
      Left            =   7920
      TabIndex        =   120
      Top             =   1365
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   114
      Left            =   7920
      TabIndex        =   119
      Top             =   1155
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   113
      Left            =   7920
      TabIndex        =   118
      Top             =   945
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   112
      Left            =   7920
      TabIndex        =   117
      Top             =   720
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   111
      Left            =   5475
      TabIndex        =   116
      Top             =   630
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   110
      Left            =   5475
      TabIndex        =   115
      Top             =   855
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   109
      Left            =   5475
      TabIndex        =   114
      Top             =   1065
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   108
      Left            =   5475
      TabIndex        =   113
      Top             =   1275
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   107
      Left            =   5475
      TabIndex        =   112
      Top             =   1485
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   106
      Left            =   5475
      TabIndex        =   111
      Top             =   1695
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   105
      Left            =   5475
      TabIndex        =   110
      Top             =   1905
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   104
      Left            =   5475
      TabIndex        =   109
      Top             =   2115
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   103
      Left            =   5460
      TabIndex        =   108
      Top             =   3795
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   102
      Left            =   5460
      TabIndex        =   107
      Top             =   3585
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   101
      Left            =   5460
      TabIndex        =   106
      Top             =   3375
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   100
      Left            =   5460
      TabIndex        =   105
      Top             =   3165
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   99
      Left            =   5460
      TabIndex        =   104
      Top             =   2955
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   98
      Left            =   5460
      TabIndex        =   103
      Top             =   2745
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   97
      Left            =   5460
      TabIndex        =   102
      Top             =   2535
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   96
      Left            =   5460
      TabIndex        =   101
      Top             =   2325
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   95
      Left            =   6120
      TabIndex        =   100
      Top             =   2445
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   94
      Left            =   6120
      TabIndex        =   99
      Top             =   2655
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   93
      Left            =   6120
      TabIndex        =   98
      Top             =   2865
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   92
      Left            =   6120
      TabIndex        =   97
      Top             =   3075
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   91
      Left            =   6120
      TabIndex        =   96
      Top             =   3285
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   90
      Left            =   6120
      TabIndex        =   95
      Top             =   3495
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   89
      Left            =   6120
      TabIndex        =   94
      Top             =   3705
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   88
      Left            =   6120
      TabIndex        =   93
      Top             =   3915
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   87
      Left            =   6135
      TabIndex        =   92
      Top             =   2235
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   86
      Left            =   6135
      TabIndex        =   91
      Top             =   2025
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   85
      Left            =   6135
      TabIndex        =   90
      Top             =   1815
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   84
      Left            =   6135
      TabIndex        =   89
      Top             =   1605
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   83
      Left            =   6135
      TabIndex        =   88
      Top             =   1395
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   82
      Left            =   6135
      TabIndex        =   87
      Top             =   1185
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   81
      Left            =   6135
      TabIndex        =   86
      Top             =   975
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   80
      Left            =   6135
      TabIndex        =   85
      Top             =   750
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   79
      Left            =   4425
      TabIndex        =   84
      Top             =   750
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   78
      Left            =   4425
      TabIndex        =   83
      Top             =   975
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   77
      Left            =   4425
      TabIndex        =   82
      Top             =   1185
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   76
      Left            =   4425
      TabIndex        =   81
      Top             =   1395
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   75
      Left            =   4425
      TabIndex        =   80
      Top             =   1605
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   74
      Left            =   4425
      TabIndex        =   79
      Top             =   1815
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   73
      Left            =   4425
      TabIndex        =   78
      Top             =   2025
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   72
      Left            =   4425
      TabIndex        =   77
      Top             =   2235
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   71
      Left            =   4410
      TabIndex        =   76
      Top             =   3915
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   70
      Left            =   4410
      TabIndex        =   75
      Top             =   3705
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   69
      Left            =   4410
      TabIndex        =   74
      Top             =   3495
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   68
      Left            =   4410
      TabIndex        =   73
      Top             =   3285
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   67
      Left            =   4410
      TabIndex        =   72
      Top             =   3075
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   66
      Left            =   4410
      TabIndex        =   71
      Top             =   2865
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   65
      Left            =   4410
      TabIndex        =   70
      Top             =   2655
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   64
      Left            =   4410
      TabIndex        =   69
      Top             =   2445
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   63
      Left            =   3750
      TabIndex        =   68
      Top             =   2325
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   62
      Left            =   3750
      TabIndex        =   67
      Top             =   2535
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   61
      Left            =   3750
      TabIndex        =   66
      Top             =   2745
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   60
      Left            =   3750
      TabIndex        =   65
      Top             =   2955
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   59
      Left            =   3750
      TabIndex        =   64
      Top             =   3165
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   58
      Left            =   3750
      TabIndex        =   63
      Top             =   3375
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   57
      Left            =   3750
      TabIndex        =   62
      Top             =   3585
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   56
      Left            =   3750
      TabIndex        =   61
      Top             =   3795
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   55
      Left            =   3765
      TabIndex        =   60
      Top             =   2115
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   54
      Left            =   3765
      TabIndex        =   59
      Top             =   1905
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   53
      Left            =   3765
      TabIndex        =   58
      Top             =   1695
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   52
      Left            =   3765
      TabIndex        =   57
      Top             =   1485
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   51
      Left            =   3765
      TabIndex        =   56
      Top             =   1275
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   50
      Left            =   3765
      TabIndex        =   55
      Top             =   1065
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   49
      Left            =   3765
      TabIndex        =   54
      Top             =   855
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   48
      Left            =   3765
      TabIndex        =   53
      Top             =   630
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   47
      Left            =   2880
      TabIndex        =   52
      Top             =   750
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   46
      Left            =   2880
      TabIndex        =   51
      Top             =   975
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   45
      Left            =   2880
      TabIndex        =   50
      Top             =   1185
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   44
      Left            =   2880
      TabIndex        =   49
      Top             =   1395
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   43
      Left            =   2880
      TabIndex        =   48
      Top             =   1605
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   42
      Left            =   2880
      TabIndex        =   47
      Top             =   1815
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   41
      Left            =   2880
      TabIndex        =   46
      Top             =   2025
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   40
      Left            =   2880
      TabIndex        =   45
      Top             =   2235
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   39
      Left            =   2865
      TabIndex        =   44
      Top             =   3915
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   38
      Left            =   2865
      TabIndex        =   43
      Top             =   3705
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   37
      Left            =   2865
      TabIndex        =   42
      Top             =   3495
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   36
      Left            =   2865
      TabIndex        =   41
      Top             =   3285
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   35
      Left            =   2865
      TabIndex        =   40
      Top             =   3075
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   34
      Left            =   2865
      TabIndex        =   39
      Top             =   2865
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   33
      Left            =   2865
      TabIndex        =   38
      Top             =   2655
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   32
      Left            =   2865
      TabIndex        =   37
      Top             =   2445
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   31
      Left            =   1635
      TabIndex        =   36
      Top             =   615
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   30
      Left            =   1635
      TabIndex        =   35
      Top             =   840
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   29
      Left            =   1635
      TabIndex        =   34
      Top             =   1050
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   28
      Left            =   1635
      TabIndex        =   33
      Top             =   1260
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   27
      Left            =   1635
      TabIndex        =   32
      Top             =   1470
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   26
      Left            =   1635
      TabIndex        =   31
      Top             =   1680
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   25
      Left            =   1635
      TabIndex        =   30
      Top             =   1890
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   24
      Left            =   1635
      TabIndex        =   29
      Top             =   2100
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   23
      Left            =   1620
      TabIndex        =   28
      Top             =   3780
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   22
      Left            =   1620
      TabIndex        =   27
      Top             =   3570
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   21
      Left            =   1620
      TabIndex        =   26
      Top             =   3360
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   20
      Left            =   1620
      TabIndex        =   25
      Top             =   3150
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   19
      Left            =   1620
      TabIndex        =   24
      Top             =   2940
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   18
      Left            =   1620
      TabIndex        =   23
      Top             =   2730
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   17
      Left            =   1620
      TabIndex        =   22
      Top             =   2520
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   16
      Left            =   1620
      TabIndex        =   21
      Top             =   2310
      Width           =   1530
   End
   Begin VB.Label Lbl3 
      Alignment       =   2  'Center
      Caption         =   "*-- - Last Winamp Loads and Playing - --*"
      Height          =   225
      Left            =   1050
      TabIndex        =   20
      Top             =   4395
      Visible         =   0   'False
      Width           =   3210
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
      Height          =   270
      Left            =   7560
      TabIndex        =   18
      Top             =   135
      Width           =   2115
   End
   Begin VB.Label TitleLbl 
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
      Height          =   315
      Left            =   150
      TabIndex        =   16
      Top             =   75
      Width           =   3945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   15
      Left            =   360
      TabIndex        =   15
      Top             =   2505
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   14
      Left            =   360
      TabIndex        =   14
      Top             =   2715
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   360
      TabIndex        =   13
      Top             =   2925
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   12
      Left            =   360
      TabIndex        =   12
      Top             =   3135
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   11
      Left            =   360
      TabIndex        =   11
      Top             =   3345
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   10
      Left            =   360
      TabIndex        =   10
      Top             =   3555
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   9
      Left            =   360
      TabIndex        =   9
      Top             =   3765
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   8
      Left            =   360
      TabIndex        =   8
      Top             =   3975
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   7
      Left            =   375
      TabIndex        =   7
      Top             =   2295
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   6
      Left            =   375
      TabIndex        =   6
      Top             =   2085
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   375
      TabIndex        =   5
      Top             =   1875
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   375
      TabIndex        =   4
      Top             =   1665
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   375
      TabIndex        =   3
      Top             =   1455
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   375
      TabIndex        =   2
      Top             =   1245
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   375
      TabIndex        =   1
      Top             =   1035
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   375
      TabIndex        =   0
      Top             =   810
      Width           =   1530
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
   Begin VB.Menu Mnu_Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Mnu_Sync 
         Caption         =   "Sync"
      End
   End
   Begin VB.Menu Mnu_LoadFolder 
      Caption         =   "Load Folder"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Form1_Width, Form1_Height
Dim X_Y_DONE_ONCE
'-------------
Public NAME_FORM

Dim XXT2

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long

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


Private Sub Combo1_Click()

Counter = Combo1.ListIndex + 1
A1$ = A4$(Counter)
B1$ = B4$(Counter)
C1$ = C4$(Counter)

SetTrueToLoadLast = True
Call Label1_Click(0)



End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 27 Then End
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

NAME_FORM = App.EXEName

Set FS = CreateObject("Scripting.FileSystemObject")

OIP$ = App.Path
tt1 = InStr(OIP$, "Shell")
tt2 = InStr(tt1, OIP$, "Loader")
OIP2$ = Mid$(OIP$, tt1 + 6, tt2 - tt1 - 7)
OIP$ = OIP2$

TEXT_PATH = App.Path + "\Text Loggs\" + GetComputerName + "-" + GetUserName

If Dir(TEXT_PATH, vbDirectory) = "" Then
    CreateFolderTree TEXT_PATH
End If

Open TEXT_PATH + "\" + OIP$ + " Loads Logg.txt" For Binary As #1

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
Combo1.AddItem C1$
A4$(Counter) = A1$
B4$(Counter) = B1$
C4$(Counter) = C1$
Counter = Counter + 1
Loop Until easy = 0

Combo1 = "List of Last Viewed Files"

If Dir$(TEXT_PATH + "\" + OIP$ + " Loads.txt") <> "" Then
    Open TEXT_PATH + "\" + OIP$ + " Loads.txt" For Input As #1
    Line Input #1, A1$
    Line Input #1, B1$
    Line Input #1, C1$
    Close #1
End If
'---- Load Last File Click Here ----
    
TitleLbl = " LAST LOAD FOLDER __ " + C1$

Lbl2.Caption = OIP$ + " Auto Loader"
Form1.Caption = Lbl2.Caption

Form1.Top = 0
Form1.Left = 0

Call FormStart.FormStartLoader

Timer_SUB_CODE_LOAD.Enabled = True

End Sub

Private Sub Timer_KEY_CODE_Timer()
If GetForegroundWindow <> Me.hWnd Then Exit Sub
A = GetAsyncKeyState(27)
If A < 0 Then End
End Sub

Private Sub Timer_SUB_CODE_LOAD_Timer()
    Timer_SUB_CODE_LOAD.Enabled = False
    Call SubCode
    Timer2.Enabled = True
End Sub


Sub SubCode()

rg = 0
FontSizez = 8
LABEL_GAP = 12

For Each Control In Controls
    If InStr(Control.Name, "Lab") Then
        rg = rg + 1: Control.Visible = False
        Control.Caption = ""
        Control.AutoSize = True
        Control.Width = 1
        Control.FontSize = FontSizez
    
    End If
Next

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
            If FS.FileExists(D1$) = False And FS.FolderExists(D1$) = False Then
                If IsIDE = False Then
                    ScanPath.ListView1.ListItems.Remove (R)
                End If
            End If
        End If
    Next
End If

Lbl2.Top = 0
Lbl2.Left = 0


x = Lbl2.Height + LABEL_GAP


Check1.Top = x
Check1.Left = (Form1.Width / 2) + 10
Check1.Width = Form1.Width / 2

tg = x

x = x + TitleLbl.Height + LABEL_GAP
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

x = tx
For R = 1 To ScanPath.ListView1.ListItems.Count
    If Label1(R - 1).Top > SH Then
        x = tx
    End If
    
    Label1(R).Top = x
    
    Label1(R).Visible = True
    'For r = 1 To ScanPath.ListView1.ListItems.Count
    A1$ = ScanPath.ListView1.ListItems.Item(R).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(R)
    
    If InStrRev(A1$, "\") > 0 Then
        seedy$ = Mid$(A1$, InStrRev(A1$, "\", Len(A1$) - 1))
    End If
    
    If oseedy$ <> seedy$ Then txr = Not txr
    oseedy$ = seedy$
    If txr = -1 Then Label1(R).BackColor = &HFF00&
    If txr = 0 Then Label1(R).BackColor = RGB(127, 200, 127) '&H56E25E
    'If txr = 0 Then Label1(R).BackColor = &HFF00&
    
    ttg$ = ScanPath.ListView1.ListItems.Item(R)
    Mid$(ttg$, 1, 1) = UCase$(Mid$(ttg$, 1, 1))
    If InStr(ttg$, ".lnk") > 0 Then ttg$ = Mid$(ttg$, 1, InStrRev(ttg$, ".") - 1)
    EE = InStr(ttg$, "&")
    If EE > 0 Then
        ttg$ = Mid$(ttg$, 1, EE) + Mid$(ttg$, EE)
    End If
    
    Label1(R).Caption = Format$(R, "00") + ". " + ttg$
    If InStr(Label1(R), ":\") > 0 And InStr(Label1(R), "\\") = 0 Then
        Label1(R).BackColor = Label2.BackColor
    End If
    
    x = x + Label1(R).Height + LABEL_GAP
    fheight = Label1(R).Top + Label1(R).Height + 420
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
TitleLbl.Top = tg
TitleLbl.Left = 0
Combo1.Top = tg
Combo1.Left = (Form1.Width / 2) + 15

TitleLbl.Width = Form1.Width / 2
Combo1.Width = (Form1.Width / 2) - 280
Lbl3.Top = fheight - 410
Lbl3.Width = Form1.Width
Lbl3.Left = 0
Form1.Refresh

Me.Visible = True

End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label1_Click(Index As Integer)

Beep

If Label1(Index).BackColor = QBColor(12) Then
    MsgBox "THIS LINK IS MARKED RED FOR NOT USE"
End If

If SetTrueToLoadLast = False Then
A1$ = ScanPath.ListView1.ListItems.Item(Index).SubItems(1)
B1$ = ScanPath.ListView1.ListItems.Item(Index)
C1$ = Form1.Label1(Index).Caption
Else
'vLaunch A1$ + B1$, vbNullString
'End
End If



' ------------------
' THIS NOT OFTEN RUN
' SEE BELOW
' ------------------
If LoadFolder = True Then
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


Private Sub Timer2_Timer()

XXT2 = XXT2 + 1
If XXT2 > ScanPath.ListView1.ListItems.Count Then
    Timer2.Enabled = False
    Exit Sub
End If


X_COLOR = Form1.Label1(XXT2).BackColor
Form1.Label1(XXT2).BackColor = &HFFFFFF
Form1.Label1(XXT2).Refresh
'DOEVENTS
SET_REAL_LINK = False
If InStr(Label1(XXT2), ":\") > 0 Then SET_REAL_LINK = True
If InStr(Label1(XXT2), "\\") > 0 Then SET_REAL_LINK = True

If SET_REAL_LINK = False Then
Sleep 10
End If

A1$ = ScanPath.ListView1.ListItems.Item(XXT2).SubItems(1)
B1$ = ScanPath.ListView1.ListItems.Item(XXT2)
'C1$ = Form1.Label1(XXT2).Caption

On Local Error Resume Next

Form1.Label1(XXT2).BackColor = X_COLOR

If SET_REAL_LINK = True Then Exit Sub

'GoTo skip:
D1$ = A1$ + B1$
If InStr(LCase(B1$), ".lnk") > 0 Then
    'LINK
    On Local Error GoTo 0
    Call GETSHORTLINK(A1$ + B1$)
    D1$ = txtTargetPath
    If Trim(D1$) = "" Then
        Label1(XXT2).BackColor = RGB(147, 147, 147): Exit Sub 'QBColor(12)
    End If
End If

Form1.Label1(XXT2).BackColor = X_COLOR

If Mid(D1$, 2, 2) = ":\" Or Mid(D1$, 1, 2) = "\\" Then
    If FS.FileExists(D1$) = False And FS.FolderExists(D1$) = False Then
        Label1(XXT2).BackColor = RGB(240, 127, 127) 'QBColor(12)
    End If
End If
End Sub


Private Sub TitleLbl_Click()

Open TEXT_PATH + "\" + OIP$ + " Loads.txt" For Input As #1
Line Input #1, A1$
Line Input #1, B1$
Line Input #1, C1$
Close #1

SetTrueToLoadLast = True
Call Label1_Click(0)

End Sub

Sub SaveLoggs()
Open TEXT_PATH + "\" + OIP$ + " Loads.txt" For Output As #1
Print #1, A1$
Print #1, B1$
Print #1, C1$
Close #1

Open TEXT_PATH + "\" + OIP$ + " Loads Logg.txt" For Append As #1
Print #1, A1$
Print #1, B1$
Print #1, C1$
Close #1

Open TEXT_PATH + "\" + OIP$ + " Loads Logg 02.txt" For Append As #1
Print #1, Format$(Now, "dd-mm-yyyy hh:mm:ss ") + A1$ + B1$
Close #1


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


