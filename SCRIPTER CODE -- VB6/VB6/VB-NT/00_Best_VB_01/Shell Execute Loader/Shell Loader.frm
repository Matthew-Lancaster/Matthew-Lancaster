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
   Begin VB.Timer Timer_1_MINUTE 
      Interval        =   1000
      Left            =   10944
      Top             =   288
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   400
      Left            =   7980
      TabIndex        =   406
      Top             =   705
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
      Index           =   399
      Left            =   7980
      TabIndex        =   405
      Top             =   930
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
      Index           =   398
      Left            =   7980
      TabIndex        =   404
      Top             =   1140
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
      Index           =   397
      Left            =   7980
      TabIndex        =   403
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
      Index           =   396
      Left            =   7980
      TabIndex        =   402
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
      Index           =   395
      Left            =   7980
      TabIndex        =   401
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
      Index           =   394
      Left            =   7980
      TabIndex        =   400
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
      Index           =   393
      Left            =   7980
      TabIndex        =   399
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
      Index           =   392
      Left            =   7965
      TabIndex        =   398
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
      Index           =   391
      Left            =   7965
      TabIndex        =   397
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
      Index           =   390
      Left            =   7965
      TabIndex        =   396
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
      Index           =   389
      Left            =   7965
      TabIndex        =   395
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
      Index           =   388
      Left            =   7965
      TabIndex        =   394
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
      Index           =   387
      Left            =   7965
      TabIndex        =   393
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
      Index           =   386
      Left            =   7965
      TabIndex        =   392
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
      Index           =   385
      Left            =   7965
      TabIndex        =   391
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
      Index           =   384
      Left            =   8625
      TabIndex        =   390
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
      Index           =   383
      Left            =   8625
      TabIndex        =   389
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
      Index           =   382
      Left            =   8625
      TabIndex        =   388
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
      Index           =   381
      Left            =   8625
      TabIndex        =   387
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
      Index           =   380
      Left            =   8625
      TabIndex        =   386
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
      Index           =   379
      Left            =   8625
      TabIndex        =   385
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
      Index           =   378
      Left            =   8625
      TabIndex        =   384
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
      Index           =   377
      Left            =   8625
      TabIndex        =   383
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
      Index           =   376
      Left            =   8640
      TabIndex        =   382
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
      Index           =   375
      Left            =   8640
      TabIndex        =   381
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
      Index           =   374
      Left            =   8640
      TabIndex        =   380
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
      Index           =   373
      Left            =   8640
      TabIndex        =   379
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
      Index           =   372
      Left            =   8640
      TabIndex        =   378
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
      Index           =   371
      Left            =   8640
      TabIndex        =   377
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
      Index           =   370
      Left            =   8640
      TabIndex        =   376
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
      Index           =   369
      Left            =   8640
      TabIndex        =   375
      Top             =   825
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
      Index           =   368
      Left            =   6930
      TabIndex        =   374
      Top             =   825
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
      Index           =   367
      Left            =   6930
      TabIndex        =   373
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
      Index           =   366
      Left            =   6930
      TabIndex        =   372
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
      Index           =   365
      Left            =   6930
      TabIndex        =   371
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
      Index           =   364
      Left            =   6930
      TabIndex        =   370
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
      Index           =   363
      Left            =   6930
      TabIndex        =   369
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
      Index           =   362
      Left            =   6930
      TabIndex        =   368
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
      Index           =   361
      Left            =   6930
      TabIndex        =   367
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
      Index           =   360
      Left            =   6915
      TabIndex        =   366
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
      Index           =   359
      Left            =   6915
      TabIndex        =   365
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
      Index           =   358
      Left            =   6915
      TabIndex        =   364
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
      Index           =   357
      Left            =   6915
      TabIndex        =   363
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
      Index           =   356
      Left            =   6915
      TabIndex        =   362
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
      Index           =   355
      Left            =   6915
      TabIndex        =   361
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
      Index           =   354
      Left            =   6915
      TabIndex        =   360
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
      Index           =   353
      Left            =   6915
      TabIndex        =   359
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
      Index           =   352
      Left            =   8835
      TabIndex        =   358
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
      Index           =   351
      Left            =   8835
      TabIndex        =   357
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
      Index           =   350
      Left            =   8835
      TabIndex        =   356
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
      Index           =   349
      Left            =   8835
      TabIndex        =   355
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
      Index           =   348
      Left            =   8835
      TabIndex        =   354
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
      Index           =   347
      Left            =   8835
      TabIndex        =   353
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
      Index           =   346
      Left            =   8835
      TabIndex        =   352
      Top             =   4185
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
      Index           =   345
      Left            =   8835
      TabIndex        =   351
      Top             =   4395
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
      Index           =   344
      Left            =   8850
      TabIndex        =   350
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
      Index           =   343
      Left            =   8850
      TabIndex        =   349
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
      Index           =   342
      Left            =   8850
      TabIndex        =   348
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
      Index           =   341
      Left            =   8850
      TabIndex        =   347
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
      Index           =   340
      Left            =   8850
      TabIndex        =   346
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
      Index           =   339
      Left            =   8850
      TabIndex        =   345
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
      Index           =   338
      Left            =   8850
      TabIndex        =   344
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
      Index           =   337
      Left            =   8850
      TabIndex        =   343
      Top             =   1230
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
      Index           =   336
      Left            =   10560
      TabIndex        =   342
      Top             =   1230
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
      Index           =   335
      Left            =   10560
      TabIndex        =   341
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
      Index           =   334
      Left            =   10560
      TabIndex        =   340
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
      Index           =   333
      Left            =   10560
      TabIndex        =   339
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
      Index           =   332
      Left            =   10560
      TabIndex        =   338
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
      Index           =   331
      Left            =   10560
      TabIndex        =   337
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
      Index           =   330
      Left            =   10560
      TabIndex        =   336
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
      Index           =   329
      Left            =   10560
      TabIndex        =   335
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
      Index           =   328
      Left            =   10545
      TabIndex        =   334
      Top             =   4395
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
      Index           =   327
      Left            =   10545
      TabIndex        =   333
      Top             =   4185
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
      Index           =   326
      Left            =   10545
      TabIndex        =   332
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
      Index           =   325
      Left            =   10545
      TabIndex        =   331
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
      Index           =   324
      Left            =   10545
      TabIndex        =   330
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
      Index           =   323
      Left            =   10545
      TabIndex        =   329
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
      Index           =   322
      Left            =   10545
      TabIndex        =   328
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
      Index           =   321
      Left            =   10545
      TabIndex        =   327
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
      Index           =   320
      Left            =   9885
      TabIndex        =   326
      Top             =   2805
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
      Index           =   319
      Left            =   9885
      TabIndex        =   325
      Top             =   3015
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
      Index           =   318
      Left            =   9885
      TabIndex        =   324
      Top             =   3225
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
      Index           =   317
      Left            =   9885
      TabIndex        =   323
      Top             =   3435
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
      Index           =   316
      Left            =   9885
      TabIndex        =   322
      Top             =   3645
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
      Index           =   315
      Left            =   9885
      TabIndex        =   321
      Top             =   3855
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
      Index           =   314
      Left            =   9885
      TabIndex        =   320
      Top             =   4065
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
      Index           =   313
      Left            =   9885
      TabIndex        =   319
      Top             =   4275
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
      Index           =   312
      Left            =   9900
      TabIndex        =   318
      Top             =   2595
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
      Index           =   311
      Left            =   9900
      TabIndex        =   317
      Top             =   2385
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
      Index           =   310
      Left            =   9900
      TabIndex        =   316
      Top             =   2175
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
      Index           =   309
      Left            =   9900
      TabIndex        =   315
      Top             =   1965
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
      Index           =   308
      Left            =   9900
      TabIndex        =   314
      Top             =   1755
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
      Index           =   307
      Left            =   9900
      TabIndex        =   313
      Top             =   1545
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
      Index           =   306
      Left            =   9900
      TabIndex        =   312
      Top             =   1335
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
      Index           =   305
      Left            =   9900
      TabIndex        =   311
      Top             =   1110
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
      Index           =   304
      Left            =   7455
      TabIndex        =   310
      Top             =   1020
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
      Index           =   303
      Left            =   7455
      TabIndex        =   309
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
      Index           =   302
      Left            =   7455
      TabIndex        =   308
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
      Index           =   301
      Left            =   7455
      TabIndex        =   307
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
      Index           =   300
      Left            =   7455
      TabIndex        =   306
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
      Index           =   299
      Left            =   7455
      TabIndex        =   305
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
      Index           =   298
      Left            =   7455
      TabIndex        =   304
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
      Index           =   297
      Left            =   7455
      TabIndex        =   303
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
      Index           =   296
      Left            =   7440
      TabIndex        =   302
      Top             =   4185
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
      Index           =   295
      Left            =   7440
      TabIndex        =   301
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
      Index           =   294
      Left            =   7440
      TabIndex        =   300
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
      Index           =   293
      Left            =   7440
      TabIndex        =   299
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
      Index           =   292
      Left            =   7440
      TabIndex        =   298
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
      Index           =   291
      Left            =   7440
      TabIndex        =   297
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
      Index           =   290
      Left            =   7440
      TabIndex        =   296
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
      Index           =   289
      Left            =   7440
      TabIndex        =   295
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
      Index           =   288
      Left            =   8100
      TabIndex        =   294
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
      Index           =   287
      Left            =   8100
      TabIndex        =   293
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
      Index           =   286
      Left            =   8100
      TabIndex        =   292
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
      Index           =   285
      Left            =   8100
      TabIndex        =   291
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
      Index           =   284
      Left            =   8100
      TabIndex        =   290
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
      Index           =   283
      Left            =   8100
      TabIndex        =   289
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
      Index           =   282
      Left            =   8100
      TabIndex        =   288
      Top             =   4095
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
      Index           =   281
      Left            =   8100
      TabIndex        =   287
      Top             =   4305
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
      Index           =   280
      Left            =   8115
      TabIndex        =   286
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
      Index           =   279
      Left            =   8115
      TabIndex        =   285
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
      Index           =   278
      Left            =   8115
      TabIndex        =   284
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
      Index           =   277
      Left            =   8115
      TabIndex        =   283
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
      Index           =   276
      Left            =   8115
      TabIndex        =   282
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
      Index           =   275
      Left            =   8115
      TabIndex        =   281
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
      Index           =   274
      Left            =   8115
      TabIndex        =   280
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
      Index           =   273
      Left            =   8115
      TabIndex        =   279
      Top             =   1140
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
      Index           =   272
      Left            =   6405
      TabIndex        =   278
      Top             =   1140
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
      Index           =   271
      Left            =   6405
      TabIndex        =   277
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
      Index           =   270
      Left            =   6405
      TabIndex        =   276
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
      Index           =   269
      Left            =   6405
      TabIndex        =   275
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
      Index           =   268
      Left            =   6405
      TabIndex        =   274
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
      Index           =   267
      Left            =   6405
      TabIndex        =   273
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
      Index           =   266
      Left            =   6405
      TabIndex        =   272
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
      Index           =   265
      Left            =   6405
      TabIndex        =   271
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
      Index           =   264
      Left            =   6390
      TabIndex        =   270
      Top             =   4305
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
      Index           =   263
      Left            =   6390
      TabIndex        =   269
      Top             =   4095
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
      Index           =   262
      Left            =   6390
      TabIndex        =   268
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
      Index           =   261
      Left            =   6390
      TabIndex        =   267
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
      Index           =   260
      Left            =   6390
      TabIndex        =   266
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
      Index           =   259
      Left            =   6390
      TabIndex        =   265
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
      Index           =   258
      Left            =   6390
      TabIndex        =   264
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
      Index           =   257
      Left            =   6390
      TabIndex        =   263
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
      Height          =   270
      Index           =   160
      Left            =   5430
      TabIndex        =   165
      Top             =   5370
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
      Left            =   8655
      TabIndex        =   20
      Top             =   5460
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
      BackColor       =   &H0000FFFF&
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
   Begin VB.Menu Mnu_Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Mnu_Sync 
         Caption         =   "Sync"
      End
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu Mnu_LoadFolder 
      Caption         =   "! Load Folder"
   End
   Begin VB.Menu Mnu_Load_Folderfile 
      Caption         =   "! Load Folder 2"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XVB_DATE
Dim fs


Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long
'Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Example of using ShellExecute
' ShellExecute hwnd, "open", pathandfile, vbNullString, vbNullString, conSwNormal

Public Sub vLaunch(vFile As String, vPara As String)
    'ShellExecute hWnd, "open", vFile, vPara, vbNullString, conSwNormal
    'ShellExecute Me.hWnd, "open", vFile, vbNullString, vbNullString, conSwNormal
    i = ShellExecute(Me.hWnd, "open", vFile, vPara, vbNullString, conSwNormal)
End Sub



Private Sub Combo1_Click()
Counter = Combo1.ListIndex + 1
A1$ = A4$(Counter)
B1$ = B4$(Counter)
C1$ = C4$(Counter)

SetTrueToLoadLast = True
Call Label1_Click(0)

End Sub


Private Sub Form_Load()

If App.PrevInstance = True Then End

Set fs = CreateObject("Scripting.FileSystemObject")

ReDim A4$(500), B4$(500), C4$(500)

OIP$ = App.Path
tt1 = InStr(OIP$, "Shell")
tt2 = InStr(tt1, OIP$, "Loader")
OIP2$ = Mid$(OIP$, tt1 + 6, tt2 - tt1 - 7)
OIP$ = OIP2$
On Error Resume Next
MkDir App.Path + "\Text Loggs"
MkDir App.Path + "\Text Loggs\" + GetComputerName + "-" + GetUserName
On Error GoTo 0

Open App.Path + "\Text Loggs\" + GetComputerName + "-" + GetUserName + "\" + OIP$ + " Loads Logg.txt" For Binary As #1

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

If Dir$(App.Path + "\Text Loggs\" + GetComputerName + "-" + GetUserName + "\" + OIP$ + " Loads.txt") <> "" Then
    Open App.Path + "\Text Loggs\" + GetComputerName + "-" + GetUserName + "\" + OIP$ + " Loads.txt" For Input As #1
    Line Input #1, A1$
    Line Input #1, B1$
    Line Input #1, C1$
    Close #1
End If
'---- Load Last File Click Here ----
    
TitleLbl = "-- Load Last File -- " + C1$

FontSizez = 10

Lbl2.Caption = OIP$ + " Auto Loader"
Form1.Caption = Lbl2.Caption

Form1.Top = 0
Form1.Left = 0

Call FormStart.FormStartLoader

Call SubCode



End Sub

Sub SubCode()


Lbl2.Top = 0
Lbl2.Left = 0


x = Lbl2.Height + 15


Check1.Top = x
Check1.Left = (Form1.Width / 2) + 10
Check1.Width = Form1.Width / 2

tg = x

x = x + TitleLbl.Height + 15
tx = x


w = 0

RD = 0


rg = 0
For Each Control In Controls
If InStr(Control.Name, "Lab") Then
rg = rg + 1: Control.Visible = False
Control.Caption = ""
Control.AutoSize = True
Control.Width = 1
Control.FontSize = FontSizez

End If
Next
DoEvents

SH = Screen.Height - 3000 'higer = smaller


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
    If txr = 0 Then Label1(R).BackColor = &H56E25E
    
    ttg$ = ScanPath.ListView1.ListItems.Item(R)
    Mid$(ttg$, 1, 1) = UCase$(Mid$(ttg$, 1, 1))
    If InStr(ttg$, ".lnk") > 0 Then ttg$ = Mid$(ttg$, 1, InStrRev(ttg$, ".") - 1)
    
    Label1(R).Caption = Format$(R, "00") + ". " + ttg$
    x = x + Label1(R).Height + 15
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
    x = x + Label1(R).Height + 15
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

Form1.Width = Label1(R - 1).Left + Label1(R - 1).Width + 120
Form1.Height = fheightx + 900

'SH = Screen.Height - 2750 'higer = smaller



Lbl2.Width = Form1.Width
TitleLbl.Top = tg
TitleLbl.Left = 0
Combo1.Top = tg
Combo1.Left = (Form1.Width / 2) + 15

TitleLbl.Width = Form1.Width / 2
Combo1.Width = (Form1.Width / 2) - 150
Lbl3.Top = fheight - 410
Lbl3.Width = Form1.Width
Lbl3.Left = 0
Form1.Refresh


End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label1_Click(index As Integer)
Form1.Visible = False
If SetTrueToLoadLast = False Then
A1$ = ScanPath.ListView1.ListItems.Item(index).SubItems(1)
B1$ = ScanPath.ListView1.ListItems.Item(index)
C1$ = Form1.Label1(index).Caption
Else
vLaunch A1$ + B1$, vbNullString
End
End If

If LoadFolder = True Then
    Shell "explorer /e, " + A1$, vbNormalFocus
    End
End If

Call SaveLoggs

Call FormStart.LabelClick(index)

End Sub

Private Sub Lbl2_Click()

Shell "explorer /e, E:\01 VB Shell Folders\00 " + App.EXEName, vbNormalFocus

End Sub

Private Sub Mnu_Load_Folderfile_Click()
LoadFolderFile = True
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


Private Sub MNU_VB_ME_2_Click()

'Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
'End

'    Dim MSGQ, IR, CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2
    
    Beep
    If GetSpecialfolder(38) = "" Then
        MsgBox "YOU HAVE TO SWITCH ADMIN MODE ON" + vbCrLf + vbCrLf + "THE COMMAND GetSpecialfolder(38) REQUIRE ADMINISTRATOR" + vbCrLf + vbCrLf + "RESULT RETURN NOTHING HERE GetSpecialfolder(38)__<" + GetSpecialfolder(38) + ">__ RESULT RETURNED NOTHING", vbMsgBoxSetForeground
        Exit Sub
    End If
    
    If IsIDE = True Then
        If IsIDE = True Then
            MSGQ = vbYesNoCancel
        Else
            MSGQ = vbOKOnly
        End If
        IR = MsgBox("NOT IN IDE ARE YOU SURE", MSGQ, vbMsgBoxSetForeground)
        If IR <> vbYes Then Exit Sub
    End If
    
    Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
    Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + CODER_VBP_FILE_NAME_2 + """", vbNormalFocus
    
    EXIT_TRUE = True
    Unload Me
    'End

End Sub
'-------------------------------------------
'-------------------------------------------
Sub GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)

    Dim FileSpec, FILE_SPEC_GO, i
    
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".vbp"
    i = CODER_VBP_FILE_NAME_2
    CODER_EXE_FILE_NAME_1 = i
    If InStr(i, " - 64 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(i, " - 64 BIT.vbp", ".vbp")
    If InStr(i, " - 32 BIT.vbp") Then CODER_EXE_FILE_NAME_1 = Replace(i, " - 32 BIT.vbp", ".vbp")
    
    CODER_EXE_FILE_NAME_1 = Replace(CODER_EXE_FILE_NAME_1, ".vbp", ".exe")
    
    FileSpec = CODER_VBP_FILE_NAME_2
    If Dir(FileSpec) <> "" Then
        FILE_SPEC_GO = FileSpec
    Else
        If LCase(GetSpecialfolder(&H29)) = LCase("C:\Windows\SysWOW64") Then
            FileSpec = Replace(FileSpec, ".vbp", " - 64 BIT.vbp")
        Else
            FileSpec = Replace(FileSpec, ".vbp", " - 32 BIT.vbp")
        End If
    End If
    
    If Dir(FileSpec) <> "" Then FILE_SPEC_GO = FileSpec
    
    CODER_VBP_FILE_NAME_2 = FILE_SPEC_GO
End Sub


Private Sub TitleLbl_Click()
Open App.Path + "\Text Loggs\" + GetComputerName + "-" + GetUserName + "\" + OIP$ + " Loads.txt" For Input As #1
Line Input #1, A1$
Line Input #1, B1$
Line Input #1, C1$
Close #1

SetTrueToLoadLast = True
Call Label1_Click(0)

End Sub

Sub SaveLoggs()
Open App.Path + "\Text Loggs\" + GetComputerName + "-" + GetUserName + "\" + OIP$ + " Loads.txt" For Output As #1
Print #1, A1$
Print #1, B1$
Print #1, C1$
Close #1

Open App.Path + "\Text Loggs\" + GetComputerName + "-" + GetUserName + "\" + OIP$ + " Loads Logg.txt" For Append As #1
Print #1, A1$
Print #1, B1$
Print #1, C1$
Close #1

Open App.Path + "\Text Loggs\" + GetComputerName + "-" + GetUserName + "\" + OIP$ + " Loads Logg 02.txt" For Append As #1
Print #1, Format$(Now, "dd-mm-yyyy hh:mm:ss ") + A1$ + B1$
Close #1


End Sub


'-------------------------------------------------
'-------------------------------------------------
Private Sub Timer_1_MINUTE_Timer()

Timer_1_MINUTE.Interval = 60000

Call VB_PROJECT_CHECKDATE

End Sub
'-------------------------------------------------
'-------------------------------------------------
Sub VB_PROJECT_CHECKDATE()

'TOP DELCLARE -------------
'DIM XVB_DATE
'--------------------------

If Mid(App.Path, 1, 1) <> "D" Then Exit Sub

PATH_FILE_NAME1 = App.Path + "\" + App.EXEName + ".EXE"
PATH_FILE_NAME2 = Replace(PATH_FILE_NAME1, "D:\VB6\", "D:\VB6-EXE'S\")

'---------------------------------------------------
'IF A NEW PROJECT NOT BEEN SYNC YET TO MIRROR FOLDER
'---------------------------------------------------
'MINOR WORK AROUND CREATE PATH
'---------------------------------------------------
If Dir(PATH_FILE_NAME2) = "" Then
    On Error Resume Next
        MkDir Mid(PATH_FILE_NAME2, 1, InStrRev(PATH_FILE_NAME2, "\"))
        fs.CopyFile PATH_FILE_NAME1, PATH_FILE_NAME2
    If Err.Number > 0 Then Exit Sub
End If

Set F = fs.GetFile(PATH_FILE_NAME2)

VB_DATE = F.DateLastModified

'--------
'01 OF 02
'----------------------------------
'WRITE INFO ABOUT DATE CHANGE NEWER
'----------------------------------

If XVB_DATE < VB_DATE And XVB_DATE > 0 Then
    
    '-----------------------------------
    'RUN A VBS TO COPY OVER AND RELOADER
    '-----------------------------------
    
'    FR1 = FreeFile
'    Open App.Path + "\RELOAD AND COPY.VBS" For Output As #FR1


    '----------------------------------
    'Mon 10 April 2017 16:30:50--------
    '----------------------------------
    'PROJECT REFERANCE ----------------
    'wshom.ocx
    'IF HARD FIND DO BROWSER
    'IN SYSTEM32 AND SYSWOW64
    'DIR wshom.ocx /S
    '----------------------------------
    'WINDOWS SCRIPT HOST OBJECT MODEL
    'AFTER FIND ALSO HAVE TO SELECT HER
    '----------------------------------
    '----------------------------------
    'Mon 10 April 2017 15:45:11--------
    '----
    'VISUAL BASIC wshom.ocx NOT WORKING - Google Search
    'https://www.google.co.uk/search?num=50&rlz=1C1CHBD_en-GBGB721GB721&q=VISUAL+BASIC+wshom.ocx+NOT+WORKING&oq=VIUAL+BASIC+wshom.ocx+NOT+WORKING&gs_l=serp.3..30i10k1.1533.3987.0.4658.12.12.0.0.0.0.127.888.11j1.12.0....0...1c.1.64.serp..0.12.886...33i160k1j33i21k1.fYCPCpVPeVk
    '--------
    'WScript reference in VB
    'http://forums.codeguru.com/showthread.php?30458-WScript-reference-in-VB
    '----
    'Set objShell = WScript.CreateObject("WScript.Shell") - ------Not WORKER
    '----------------------------------
    '----------------------------------
    
    Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")
    
    WSHShell.Run """" + App.Path + "\RELOAD AND COPY.VBS" + """"
    Set WSHShell = Nothing
    Unload Me
    
End If

XVB_DATE = VB_DATE

End Sub



