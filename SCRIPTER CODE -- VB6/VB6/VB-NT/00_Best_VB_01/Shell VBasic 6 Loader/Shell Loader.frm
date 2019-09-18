VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   Caption         =   "NotePad Loader"
   ClientHeight    =   5928
   ClientLeft      =   132
   ClientTop       =   780
   ClientWidth     =   12324
   Icon            =   "Shell Loader.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5928
   ScaleWidth      =   12324
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer_TITLE_COLOR_END_DONE 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1812
      Top             =   528
   End
   Begin VB.Timer Timer_SUB_CODE_LOAD 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4692
      Top             =   2544
   End
   Begin VB.Timer Timer_KEY_CODE 
      Interval        =   1
      Left            =   4548
      Top             =   2028
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5400
      Top             =   2040
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4125
      TabIndex        =   166
      Text            =   "Combo1"
      Top             =   75
      Visible         =   0   'False
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
      BackColor       =   &H00DDFFFE&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   838
      Left            =   180
      TabIndex        =   844
      Top             =   420
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CCFFFF&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   837
      Left            =   8244
      TabIndex        =   843
      Top             =   972
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DBFFDE&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   836
      Left            =   8340
      TabIndex        =   842
      Top             =   576
      Width           =   1536
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   835
      Left            =   10185
      TabIndex        =   841
      Top             =   2805
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   834
      Left            =   10185
      TabIndex        =   840
      Top             =   3015
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   833
      Left            =   10185
      TabIndex        =   839
      Top             =   3225
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   832
      Left            =   10185
      TabIndex        =   838
      Top             =   3435
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   831
      Left            =   10185
      TabIndex        =   837
      Top             =   3645
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   830
      Left            =   10185
      TabIndex        =   836
      Top             =   3855
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   829
      Left            =   10185
      TabIndex        =   835
      Top             =   4065
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   828
      Left            =   10185
      TabIndex        =   834
      Top             =   4275
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   827
      Left            =   10200
      TabIndex        =   833
      Top             =   2595
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   826
      Left            =   10200
      TabIndex        =   832
      Top             =   2385
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   825
      Left            =   10200
      TabIndex        =   831
      Top             =   2175
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   824
      Left            =   10200
      TabIndex        =   830
      Top             =   1965
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   823
      Left            =   10200
      TabIndex        =   829
      Top             =   1755
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   822
      Left            =   10200
      TabIndex        =   828
      Top             =   1545
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   821
      Left            =   10200
      TabIndex        =   827
      Top             =   1335
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   820
      Left            =   10200
      TabIndex        =   826
      Top             =   1110
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   819
      Left            =   10095
      TabIndex        =   825
      Top             =   2895
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   818
      Left            =   10095
      TabIndex        =   824
      Top             =   3105
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   817
      Left            =   10095
      TabIndex        =   823
      Top             =   3315
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   816
      Left            =   10095
      TabIndex        =   822
      Top             =   3525
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   815
      Left            =   10095
      TabIndex        =   821
      Top             =   3735
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   814
      Left            =   10095
      TabIndex        =   820
      Top             =   3945
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   813
      Left            =   10095
      TabIndex        =   819
      Top             =   4155
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   812
      Left            =   10095
      TabIndex        =   818
      Top             =   4365
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   811
      Left            =   10110
      TabIndex        =   817
      Top             =   2685
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   810
      Left            =   10110
      TabIndex        =   816
      Top             =   2475
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   809
      Left            =   10110
      TabIndex        =   815
      Top             =   2265
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   808
      Left            =   10110
      TabIndex        =   814
      Top             =   2055
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   807
      Left            =   10110
      TabIndex        =   813
      Top             =   1845
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   806
      Left            =   10110
      TabIndex        =   812
      Top             =   1635
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   805
      Left            =   10110
      TabIndex        =   811
      Top             =   1425
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   804
      Left            =   10110
      TabIndex        =   810
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   803
      Left            =   10080
      TabIndex        =   809
      Top             =   1170
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   802
      Left            =   10080
      TabIndex        =   808
      Top             =   1395
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   801
      Left            =   10176
      TabIndex        =   807
      Top             =   1704
      Width           =   132
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   800
      Left            =   10080
      TabIndex        =   806
      Top             =   1815
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   799
      Left            =   10080
      TabIndex        =   805
      Top             =   2025
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   798
      Left            =   10080
      TabIndex        =   804
      Top             =   2235
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   797
      Left            =   10080
      TabIndex        =   803
      Top             =   2445
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   796
      Left            =   10080
      TabIndex        =   802
      Top             =   2655
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   795
      Left            =   10065
      TabIndex        =   801
      Top             =   4335
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   794
      Left            =   10065
      TabIndex        =   800
      Top             =   4125
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   793
      Left            =   10065
      TabIndex        =   799
      Top             =   3915
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   792
      Left            =   10065
      TabIndex        =   798
      Top             =   3705
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   791
      Left            =   10065
      TabIndex        =   797
      Top             =   3495
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   790
      Left            =   10065
      TabIndex        =   796
      Top             =   3285
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   789
      Left            =   10065
      TabIndex        =   795
      Top             =   3075
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   788
      Left            =   10065
      TabIndex        =   794
      Top             =   2865
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   787
      Left            =   10200
      TabIndex        =   793
      Top             =   3300
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   786
      Left            =   10200
      TabIndex        =   792
      Top             =   3510
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   785
      Left            =   10200
      TabIndex        =   791
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   784
      Left            =   10200
      TabIndex        =   790
      Top             =   3930
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   783
      Left            =   10200
      TabIndex        =   789
      Top             =   4140
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   782
      Left            =   10200
      TabIndex        =   788
      Top             =   4350
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   781
      Left            =   10200
      TabIndex        =   787
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   780
      Left            =   10215
      TabIndex        =   786
      Top             =   3090
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   779
      Left            =   10215
      TabIndex        =   785
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   778
      Left            =   10215
      TabIndex        =   784
      Top             =   2670
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   777
      Left            =   10215
      TabIndex        =   783
      Top             =   2460
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   776
      Left            =   10215
      TabIndex        =   782
      Top             =   2250
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   775
      Left            =   10215
      TabIndex        =   781
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   774
      Left            =   10215
      TabIndex        =   780
      Top             =   1830
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   773
      Left            =   10215
      TabIndex        =   779
      Top             =   1605
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   772
      Left            =   10110
      TabIndex        =   778
      Top             =   1620
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   771
      Left            =   10110
      TabIndex        =   777
      Top             =   1845
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   770
      Left            =   10110
      TabIndex        =   776
      Top             =   2055
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   769
      Left            =   10110
      TabIndex        =   775
      Top             =   2265
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   768
      Left            =   10110
      TabIndex        =   774
      Top             =   2475
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   767
      Left            =   10110
      TabIndex        =   773
      Top             =   2685
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   766
      Left            =   10110
      TabIndex        =   772
      Top             =   2895
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   765
      Left            =   10110
      TabIndex        =   771
      Top             =   3105
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   764
      Left            =   10095
      TabIndex        =   770
      Top             =   4575
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   763
      Left            =   10095
      TabIndex        =   769
      Top             =   4365
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   762
      Left            =   10095
      TabIndex        =   768
      Top             =   4155
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   761
      Left            =   10095
      TabIndex        =   767
      Top             =   3945
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   760
      Left            =   10095
      TabIndex        =   766
      Top             =   3735
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   759
      Left            =   10095
      TabIndex        =   765
      Top             =   3525
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   758
      Left            =   10095
      TabIndex        =   764
      Top             =   3315
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   757
      Left            =   10305
      TabIndex        =   763
      Top             =   2775
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   756
      Left            =   10305
      TabIndex        =   762
      Top             =   2985
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   755
      Left            =   10305
      TabIndex        =   761
      Top             =   3195
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   754
      Left            =   10305
      TabIndex        =   760
      Top             =   3405
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   753
      Left            =   10305
      TabIndex        =   759
      Top             =   3615
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   752
      Left            =   10305
      TabIndex        =   758
      Top             =   3825
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   751
      Left            =   10305
      TabIndex        =   757
      Top             =   4035
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   750
      Left            =   10305
      TabIndex        =   756
      Top             =   4245
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   749
      Left            =   10320
      TabIndex        =   755
      Top             =   2565
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   748
      Left            =   10320
      TabIndex        =   754
      Top             =   2355
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   747
      Left            =   10320
      TabIndex        =   753
      Top             =   2145
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   746
      Left            =   10320
      TabIndex        =   752
      Top             =   1935
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   745
      Left            =   10320
      TabIndex        =   751
      Top             =   1716
      Width           =   132
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   744
      Left            =   10320
      TabIndex        =   750
      Top             =   1515
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   743
      Left            =   10320
      TabIndex        =   749
      Top             =   1305
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   742
      Left            =   10320
      TabIndex        =   748
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   741
      Left            =   9960
      TabIndex        =   747
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   740
      Left            =   9960
      TabIndex        =   746
      Top             =   3090
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   739
      Left            =   9960
      TabIndex        =   745
      Top             =   3300
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   738
      Left            =   9960
      TabIndex        =   744
      Top             =   3510
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   737
      Left            =   9960
      TabIndex        =   743
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   736
      Left            =   9960
      TabIndex        =   742
      Top             =   3930
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   735
      Left            =   9960
      TabIndex        =   741
      Top             =   4140
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   734
      Left            =   9960
      TabIndex        =   740
      Top             =   4350
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   733
      Left            =   9975
      TabIndex        =   739
      Top             =   2670
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   732
      Left            =   9975
      TabIndex        =   738
      Top             =   2460
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   731
      Left            =   9975
      TabIndex        =   737
      Top             =   2250
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   730
      Left            =   9975
      TabIndex        =   736
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   729
      Left            =   9975
      TabIndex        =   735
      Top             =   1830
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   728
      Left            =   9975
      TabIndex        =   734
      Top             =   1620
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   727
      Left            =   9975
      TabIndex        =   733
      Top             =   1410
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   726
      Left            =   9975
      TabIndex        =   732
      Top             =   1185
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   725
      Left            =   10050
      TabIndex        =   731
      Top             =   2925
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   724
      Left            =   10050
      TabIndex        =   730
      Top             =   3135
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   723
      Left            =   10050
      TabIndex        =   729
      Top             =   3345
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   722
      Left            =   10050
      TabIndex        =   728
      Top             =   3555
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   721
      Left            =   10050
      TabIndex        =   727
      Top             =   3765
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   720
      Left            =   10050
      TabIndex        =   726
      Top             =   3975
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   719
      Left            =   10050
      TabIndex        =   725
      Top             =   4185
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   718
      Left            =   10050
      TabIndex        =   724
      Top             =   4395
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   717
      Left            =   10065
      TabIndex        =   723
      Top             =   2715
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   716
      Left            =   10065
      TabIndex        =   722
      Top             =   2505
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   715
      Left            =   10065
      TabIndex        =   721
      Top             =   2295
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   714
      Left            =   10065
      TabIndex        =   720
      Top             =   2085
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   713
      Left            =   10065
      TabIndex        =   719
      Top             =   1875
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   712
      Left            =   10065
      TabIndex        =   718
      Top             =   1665
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   711
      Left            =   10065
      TabIndex        =   717
      Top             =   1455
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   710
      Left            =   10065
      TabIndex        =   716
      Top             =   1230
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   709
      Left            =   10260
      TabIndex        =   715
      Top             =   3330
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   708
      Left            =   10260
      TabIndex        =   714
      Top             =   3540
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   707
      Left            =   10260
      TabIndex        =   713
      Top             =   3750
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   706
      Left            =   10260
      TabIndex        =   712
      Top             =   3960
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   705
      Left            =   10260
      TabIndex        =   711
      Top             =   4170
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   704
      Left            =   10260
      TabIndex        =   710
      Top             =   4380
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   703
      Left            =   10275
      TabIndex        =   709
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   702
      Left            =   10275
      TabIndex        =   708
      Top             =   2910
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   701
      Left            =   10275
      TabIndex        =   707
      Top             =   2700
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   700
      Left            =   10275
      TabIndex        =   706
      Top             =   2490
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   699
      Left            =   10275
      TabIndex        =   705
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   698
      Left            =   10275
      TabIndex        =   704
      Top             =   2070
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   697
      Left            =   10275
      TabIndex        =   703
      Top             =   1635
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   696
      Left            =   10275
      TabIndex        =   702
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   695
      Left            =   10275
      TabIndex        =   701
      Top             =   1425
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   694
      Left            =   10275
      TabIndex        =   700
      Top             =   1635
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   693
      Left            =   10275
      TabIndex        =   699
      Top             =   1845
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   692
      Left            =   10275
      TabIndex        =   698
      Top             =   2055
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   691
      Left            =   10275
      TabIndex        =   697
      Top             =   2265
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   690
      Left            =   10275
      TabIndex        =   696
      Top             =   2475
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   689
      Left            =   10275
      TabIndex        =   695
      Top             =   2685
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   688
      Left            =   10260
      TabIndex        =   694
      Top             =   4155
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   687
      Left            =   10260
      TabIndex        =   693
      Top             =   3945
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   686
      Left            =   10260
      TabIndex        =   692
      Top             =   3735
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   685
      Left            =   10260
      TabIndex        =   691
      Top             =   3525
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   684
      Left            =   10260
      TabIndex        =   690
      Top             =   3315
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   683
      Left            =   10260
      TabIndex        =   689
      Top             =   3105
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   682
      Left            =   10260
      TabIndex        =   688
      Top             =   2895
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   681
      Left            =   10170
      TabIndex        =   687
      Top             =   1215
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   680
      Left            =   10170
      TabIndex        =   686
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   679
      Left            =   10170
      TabIndex        =   685
      Top             =   1650
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   678
      Left            =   10170
      TabIndex        =   684
      Top             =   1860
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   677
      Left            =   10170
      TabIndex        =   683
      Top             =   2070
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   676
      Left            =   10170
      TabIndex        =   682
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   675
      Left            =   10170
      TabIndex        =   681
      Top             =   2490
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   674
      Left            =   10170
      TabIndex        =   680
      Top             =   2700
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   673
      Left            =   10155
      TabIndex        =   679
      Top             =   4380
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   672
      Left            =   10155
      TabIndex        =   678
      Top             =   4170
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   671
      Left            =   10155
      TabIndex        =   677
      Top             =   3960
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   670
      Left            =   10155
      TabIndex        =   676
      Top             =   3750
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   669
      Left            =   10155
      TabIndex        =   675
      Top             =   3540
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   668
      Left            =   10155
      TabIndex        =   674
      Top             =   3330
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   667
      Left            =   10155
      TabIndex        =   673
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   666
      Left            =   10155
      TabIndex        =   672
      Top             =   2910
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   665
      Left            =   11130
      TabIndex        =   671
      Top             =   1035
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   664
      Left            =   11130
      TabIndex        =   670
      Top             =   1260
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   663
      Left            =   11130
      TabIndex        =   669
      Top             =   1470
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   662
      Left            =   11130
      TabIndex        =   668
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   661
      Left            =   11130
      TabIndex        =   667
      Top             =   1890
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   660
      Left            =   11130
      TabIndex        =   666
      Top             =   2100
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   659
      Left            =   11130
      TabIndex        =   665
      Top             =   2310
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   658
      Left            =   11130
      TabIndex        =   664
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   657
      Left            =   11115
      TabIndex        =   663
      Top             =   4200
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   656
      Left            =   11115
      TabIndex        =   662
      Top             =   3990
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   655
      Left            =   11115
      TabIndex        =   661
      Top             =   3780
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   654
      Left            =   11115
      TabIndex        =   660
      Top             =   3570
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   653
      Left            =   11115
      TabIndex        =   659
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   652
      Left            =   11115
      TabIndex        =   658
      Top             =   3150
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   651
      Left            =   11115
      TabIndex        =   657
      Top             =   2940
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   650
      Left            =   11115
      TabIndex        =   656
      Top             =   2730
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   649
      Left            =   10365
      TabIndex        =   655
      Top             =   2895
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   648
      Left            =   10365
      TabIndex        =   654
      Top             =   3105
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   647
      Left            =   10365
      TabIndex        =   653
      Top             =   3315
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   646
      Left            =   10365
      TabIndex        =   652
      Top             =   3525
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   645
      Left            =   10365
      TabIndex        =   651
      Top             =   3735
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   644
      Left            =   10365
      TabIndex        =   650
      Top             =   3945
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   643
      Left            =   10365
      TabIndex        =   649
      Top             =   4155
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   642
      Left            =   10365
      TabIndex        =   648
      Top             =   4365
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   641
      Left            =   10380
      TabIndex        =   647
      Top             =   2685
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   640
      Left            =   10380
      TabIndex        =   646
      Top             =   2475
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   639
      Left            =   10380
      TabIndex        =   645
      Top             =   2265
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   638
      Left            =   10380
      TabIndex        =   644
      Top             =   2055
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   637
      Left            =   10380
      TabIndex        =   643
      Top             =   1845
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   636
      Left            =   10380
      TabIndex        =   642
      Top             =   1635
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   635
      Left            =   10380
      TabIndex        =   641
      Top             =   1425
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   634
      Left            =   10380
      TabIndex        =   640
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   633
      Left            =   11160
      TabIndex        =   639
      Top             =   1125
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   632
      Left            =   11160
      TabIndex        =   638
      Top             =   1350
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   631
      Left            =   11160
      TabIndex        =   637
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   630
      Left            =   11160
      TabIndex        =   636
      Top             =   1770
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   629
      Left            =   11160
      TabIndex        =   635
      Top             =   1980
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   628
      Left            =   11160
      TabIndex        =   634
      Top             =   2190
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   627
      Left            =   11160
      TabIndex        =   633
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   626
      Left            =   11160
      TabIndex        =   632
      Top             =   2610
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   625
      Left            =   11145
      TabIndex        =   631
      Top             =   4290
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   624
      Left            =   11145
      TabIndex        =   630
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   623
      Left            =   11145
      TabIndex        =   629
      Top             =   3870
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   622
      Left            =   11145
      TabIndex        =   628
      Top             =   3660
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   621
      Left            =   11145
      TabIndex        =   627
      Top             =   3450
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   620
      Left            =   11145
      TabIndex        =   626
      Top             =   3240
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   619
      Left            =   11145
      TabIndex        =   625
      Top             =   3030
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   618
      Left            =   11145
      TabIndex        =   624
      Top             =   2820
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   617
      Left            =   10575
      TabIndex        =   623
      Top             =   3300
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   616
      Left            =   10575
      TabIndex        =   622
      Top             =   3510
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   615
      Left            =   10575
      TabIndex        =   621
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   614
      Left            =   10575
      TabIndex        =   620
      Top             =   3930
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   613
      Left            =   10575
      TabIndex        =   619
      Top             =   4140
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   612
      Left            =   10575
      TabIndex        =   618
      Top             =   4350
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   611
      Left            =   10575
      TabIndex        =   617
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   610
      Left            =   10590
      TabIndex        =   616
      Top             =   3090
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   609
      Left            =   10596
      TabIndex        =   615
      Top             =   2868
      Width           =   132
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   608
      Left            =   10590
      TabIndex        =   614
      Top             =   2670
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   607
      Left            =   10590
      TabIndex        =   613
      Top             =   2460
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   606
      Left            =   10590
      TabIndex        =   612
      Top             =   2250
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   605
      Left            =   10590
      TabIndex        =   611
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   604
      Left            =   10590
      TabIndex        =   610
      Top             =   1605
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   603
      Left            =   10860
      TabIndex        =   609
      Top             =   1365
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   602
      Left            =   10860
      TabIndex        =   608
      Top             =   1590
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   601
      Left            =   10860
      TabIndex        =   607
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   600
      Left            =   10860
      TabIndex        =   606
      Top             =   2010
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   599
      Left            =   10860
      TabIndex        =   605
      Top             =   2220
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   598
      Left            =   10860
      TabIndex        =   604
      Top             =   2430
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   597
      Left            =   10860
      TabIndex        =   603
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   596
      Left            =   10860
      TabIndex        =   602
      Top             =   2850
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   595
      Left            =   10845
      TabIndex        =   601
      Top             =   4320
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   594
      Left            =   10845
      TabIndex        =   600
      Top             =   4110
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   593
      Left            =   10845
      TabIndex        =   599
      Top             =   3900
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   592
      Left            =   10845
      TabIndex        =   598
      Top             =   3690
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   591
      Left            =   10845
      TabIndex        =   597
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   590
      Left            =   10845
      TabIndex        =   596
      Top             =   3270
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   589
      Left            =   10845
      TabIndex        =   595
      Top             =   3060
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   588
      Left            =   11220
      TabIndex        =   594
      Top             =   3075
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   587
      Left            =   11220
      TabIndex        =   593
      Top             =   3285
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   586
      Left            =   11220
      TabIndex        =   592
      Top             =   3495
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   585
      Left            =   11220
      TabIndex        =   591
      Top             =   3705
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   584
      Left            =   11220
      TabIndex        =   590
      Top             =   3915
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   583
      Left            =   11220
      TabIndex        =   589
      Top             =   4125
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   582
      Left            =   11220
      TabIndex        =   588
      Top             =   4335
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   581
      Left            =   11235
      TabIndex        =   587
      Top             =   2865
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   580
      Left            =   11235
      TabIndex        =   586
      Top             =   2655
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   579
      Left            =   11235
      TabIndex        =   585
      Top             =   2445
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   578
      Left            =   11235
      TabIndex        =   584
      Top             =   2235
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   577
      Left            =   11235
      TabIndex        =   583
      Top             =   2025
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   576
      Left            =   11235
      TabIndex        =   582
      Top             =   1815
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   575
      Left            =   11235
      TabIndex        =   581
      Top             =   1605
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   574
      Left            =   10590
      TabIndex        =   580
      Top             =   1170
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   573
      Left            =   10590
      TabIndex        =   579
      Top             =   1395
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   572
      Left            =   10590
      TabIndex        =   578
      Top             =   1605
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   571
      Left            =   10590
      TabIndex        =   577
      Top             =   1815
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   570
      Left            =   10590
      TabIndex        =   576
      Top             =   2025
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   569
      Left            =   10590
      TabIndex        =   575
      Top             =   2235
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   568
      Left            =   10590
      TabIndex        =   574
      Top             =   2445
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   567
      Left            =   10590
      TabIndex        =   573
      Top             =   2655
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   566
      Left            =   10575
      TabIndex        =   572
      Top             =   4335
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   565
      Left            =   10575
      TabIndex        =   571
      Top             =   4125
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   564
      Left            =   10575
      TabIndex        =   570
      Top             =   3915
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   563
      Left            =   10575
      TabIndex        =   569
      Top             =   3705
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   562
      Left            =   10575
      TabIndex        =   568
      Top             =   3495
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   561
      Left            =   10575
      TabIndex        =   567
      Top             =   3285
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   560
      Left            =   10575
      TabIndex        =   566
      Top             =   3075
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   559
      Left            =   10575
      TabIndex        =   565
      Top             =   2865
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   558
      Left            =   11250
      TabIndex        =   564
      Top             =   3165
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   557
      Left            =   11250
      TabIndex        =   563
      Top             =   3375
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   556
      Left            =   11250
      TabIndex        =   562
      Top             =   3585
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   555
      Left            =   11250
      TabIndex        =   561
      Top             =   3795
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   554
      Left            =   11250
      TabIndex        =   560
      Top             =   4005
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   553
      Left            =   11250
      TabIndex        =   559
      Top             =   4215
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   552
      Left            =   11250
      TabIndex        =   558
      Top             =   4425
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   551
      Left            =   11265
      TabIndex        =   557
      Top             =   2955
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   550
      Left            =   11265
      TabIndex        =   556
      Top             =   2745
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   549
      Left            =   11265
      TabIndex        =   555
      Top             =   2535
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   548
      Left            =   11265
      TabIndex        =   554
      Top             =   2325
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   547
      Left            =   11265
      TabIndex        =   553
      Top             =   2115
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   546
      Left            =   11265
      TabIndex        =   552
      Top             =   1905
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   545
      Left            =   11265
      TabIndex        =   551
      Top             =   1695
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   544
      Left            =   11265
      TabIndex        =   550
      Top             =   1470
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   543
      Left            =   11160
      TabIndex        =   549
      Top             =   1485
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   542
      Left            =   11160
      TabIndex        =   548
      Top             =   1710
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   541
      Left            =   11160
      TabIndex        =   547
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   540
      Left            =   11160
      TabIndex        =   546
      Top             =   2130
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   539
      Left            =   11160
      TabIndex        =   545
      Top             =   2340
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   538
      Left            =   11160
      TabIndex        =   544
      Top             =   2550
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   537
      Left            =   11160
      TabIndex        =   543
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   536
      Left            =   11160
      TabIndex        =   542
      Top             =   2970
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   535
      Left            =   11145
      TabIndex        =   541
      Top             =   4440
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   534
      Left            =   11145
      TabIndex        =   540
      Top             =   4230
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   533
      Left            =   11145
      TabIndex        =   539
      Top             =   4020
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   532
      Left            =   11145
      TabIndex        =   538
      Top             =   3810
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   531
      Left            =   11145
      TabIndex        =   537
      Top             =   3600
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   530
      Left            =   11145
      TabIndex        =   536
      Top             =   3390
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   529
      Left            =   11145
      TabIndex        =   535
      Top             =   3180
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   528
      Left            =   10470
      TabIndex        =   534
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   527
      Left            =   10470
      TabIndex        =   533
      Top             =   3090
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   526
      Left            =   10470
      TabIndex        =   532
      Top             =   3300
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   525
      Left            =   10470
      TabIndex        =   531
      Top             =   3510
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   524
      Left            =   10470
      TabIndex        =   530
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   523
      Left            =   10470
      TabIndex        =   529
      Top             =   3930
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   522
      Left            =   10470
      TabIndex        =   528
      Top             =   4140
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   521
      Left            =   10485
      TabIndex        =   527
      Top             =   2670
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   520
      Left            =   10485
      TabIndex        =   526
      Top             =   2460
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   519
      Left            =   10485
      TabIndex        =   525
      Top             =   2250
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   518
      Left            =   10485
      TabIndex        =   524
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   517
      Left            =   10485
      TabIndex        =   523
      Top             =   1830
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   516
      Left            =   10485
      TabIndex        =   522
      Top             =   1620
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   515
      Left            =   10485
      TabIndex        =   521
      Top             =   1410
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   514
      Left            =   10485
      TabIndex        =   520
      Top             =   1185
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   513
      Left            =   11130
      TabIndex        =   519
      Top             =   1620
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   512
      Left            =   11130
      TabIndex        =   518
      Top             =   1830
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   511
      Left            =   11130
      TabIndex        =   517
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   510
      Left            =   11130
      TabIndex        =   516
      Top             =   2250
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   509
      Left            =   11130
      TabIndex        =   515
      Top             =   2460
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   508
      Left            =   11130
      TabIndex        =   514
      Top             =   2670
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   507
      Left            =   11130
      TabIndex        =   513
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   506
      Left            =   11115
      TabIndex        =   512
      Top             =   4350
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   505
      Left            =   11115
      TabIndex        =   511
      Top             =   4140
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   504
      Left            =   11115
      TabIndex        =   510
      Top             =   3930
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   503
      Left            =   11115
      TabIndex        =   509
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   502
      Left            =   11115
      TabIndex        =   508
      Top             =   3510
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   501
      Left            =   11115
      TabIndex        =   507
      Top             =   3300
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   500
      Left            =   11115
      TabIndex        =   506
      Top             =   3090
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   499
      Left            =   10740
      TabIndex        =   505
      Top             =   3075
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   498
      Left            =   10740
      TabIndex        =   504
      Top             =   3285
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   497
      Left            =   10740
      TabIndex        =   503
      Top             =   3495
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   496
      Left            =   10740
      TabIndex        =   502
      Top             =   3705
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   495
      Left            =   10740
      TabIndex        =   501
      Top             =   3915
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   494
      Left            =   10740
      TabIndex        =   500
      Top             =   4125
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   493
      Left            =   10755
      TabIndex        =   499
      Top             =   2865
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   492
      Left            =   10755
      TabIndex        =   498
      Top             =   2655
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   491
      Left            =   10755
      TabIndex        =   497
      Top             =   2445
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   490
      Left            =   10755
      TabIndex        =   496
      Top             =   2235
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   489
      Left            =   10755
      TabIndex        =   495
      Top             =   2025
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   488
      Left            =   10755
      TabIndex        =   494
      Top             =   1815
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   487
      Left            =   10755
      TabIndex        =   493
      Top             =   1605
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   486
      Left            =   10755
      TabIndex        =   492
      Top             =   1380
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   485
      Left            =   10485
      TabIndex        =   491
      Top             =   1620
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   484
      Left            =   10485
      TabIndex        =   490
      Top             =   2055
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   483
      Left            =   10485
      TabIndex        =   489
      Top             =   2265
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   482
      Left            =   10485
      TabIndex        =   488
      Top             =   2475
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   481
      Left            =   10485
      TabIndex        =   487
      Top             =   2685
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   480
      Left            =   10485
      TabIndex        =   486
      Top             =   2895
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   479
      Left            =   10485
      TabIndex        =   485
      Top             =   3105
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   478
      Left            =   10470
      TabIndex        =   484
      Top             =   4365
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   477
      Left            =   10470
      TabIndex        =   483
      Top             =   4155
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   476
      Left            =   10470
      TabIndex        =   482
      Top             =   3945
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   475
      Left            =   10470
      TabIndex        =   481
      Top             =   3735
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   474
      Left            =   10470
      TabIndex        =   480
      Top             =   3525
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   473
      Left            =   10470
      TabIndex        =   479
      Top             =   3315
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   472
      Left            =   11040
      TabIndex        =   478
      Top             =   2835
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   471
      Left            =   11040
      TabIndex        =   477
      Top             =   3045
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   470
      Left            =   11040
      TabIndex        =   476
      Top             =   3255
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   469
      Left            =   11040
      TabIndex        =   475
      Top             =   3465
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   468
      Left            =   11040
      TabIndex        =   474
      Top             =   3675
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   467
      Left            =   11040
      TabIndex        =   473
      Top             =   3885
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   466
      Left            =   11040
      TabIndex        =   472
      Top             =   4095
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   465
      Left            =   11040
      TabIndex        =   471
      Top             =   4305
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   464
      Left            =   11055
      TabIndex        =   470
      Top             =   2625
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   463
      Left            =   11055
      TabIndex        =   469
      Top             =   2415
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   462
      Left            =   11055
      TabIndex        =   468
      Top             =   2205
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   461
      Left            =   11055
      TabIndex        =   467
      Top             =   1995
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   460
      Left            =   11055
      TabIndex        =   466
      Top             =   1785
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   459
      Left            =   11055
      TabIndex        =   465
      Top             =   1575
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   458
      Left            =   11055
      TabIndex        =   464
      Top             =   1365
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   457
      Left            =   11055
      TabIndex        =   463
      Top             =   1140
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   456
      Left            =   10275
      TabIndex        =   462
      Top             =   1215
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   455
      Left            =   10275
      TabIndex        =   461
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   454
      Left            =   10275
      TabIndex        =   460
      Top             =   1650
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   453
      Left            =   10275
      TabIndex        =   459
      Top             =   1860
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   452
      Left            =   10275
      TabIndex        =   458
      Top             =   2070
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   451
      Left            =   10275
      TabIndex        =   457
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   450
      Left            =   10275
      TabIndex        =   456
      Top             =   2490
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   449
      Left            =   10275
      TabIndex        =   455
      Top             =   2700
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   448
      Left            =   10260
      TabIndex        =   454
      Top             =   4380
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   447
      Left            =   10260
      TabIndex        =   453
      Top             =   4170
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   446
      Left            =   10260
      TabIndex        =   452
      Top             =   3960
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   445
      Left            =   10260
      TabIndex        =   451
      Top             =   3750
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   444
      Left            =   10260
      TabIndex        =   450
      Top             =   3540
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   443
      Left            =   10260
      TabIndex        =   449
      Top             =   3330
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   442
      Left            =   10260
      TabIndex        =   448
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   441
      Left            =   10260
      TabIndex        =   447
      Top             =   2910
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   440
      Left            =   11010
      TabIndex        =   446
      Top             =   2745
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   439
      Left            =   11010
      TabIndex        =   445
      Top             =   2955
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   438
      Left            =   11010
      TabIndex        =   444
      Top             =   3165
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   437
      Left            =   11010
      TabIndex        =   443
      Top             =   3375
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   436
      Left            =   11010
      TabIndex        =   442
      Top             =   3585
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   435
      Left            =   11010
      TabIndex        =   441
      Top             =   3795
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   434
      Left            =   11010
      TabIndex        =   440
      Top             =   4005
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   433
      Left            =   11010
      TabIndex        =   439
      Top             =   4215
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   432
      Left            =   11025
      TabIndex        =   438
      Top             =   2535
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   431
      Left            =   11025
      TabIndex        =   437
      Top             =   2325
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   430
      Left            =   11025
      TabIndex        =   436
      Top             =   2115
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   429
      Left            =   11025
      TabIndex        =   435
      Top             =   1905
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   428
      Left            =   11025
      TabIndex        =   434
      Top             =   1695
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   427
      Left            =   11025
      TabIndex        =   433
      Top             =   1485
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   426
      Left            =   11025
      TabIndex        =   432
      Top             =   1275
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   425
      Left            =   11025
      TabIndex        =   431
      Top             =   1050
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   424
      Left            =   10890
      TabIndex        =   430
      Top             =   1125
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   423
      Left            =   10890
      TabIndex        =   429
      Top             =   1350
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   422
      Left            =   10890
      TabIndex        =   428
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   421
      Left            =   10890
      TabIndex        =   427
      Top             =   1770
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   420
      Left            =   10890
      TabIndex        =   426
      Top             =   1980
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   419
      Left            =   10890
      TabIndex        =   425
      Top             =   2190
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   418
      Left            =   10890
      TabIndex        =   424
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   417
      Left            =   10890
      TabIndex        =   423
      Top             =   2610
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   416
      Left            =   10875
      TabIndex        =   422
      Top             =   4290
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   415
      Left            =   10875
      TabIndex        =   421
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   414
      Left            =   10875
      TabIndex        =   420
      Top             =   3870
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   413
      Left            =   10875
      TabIndex        =   419
      Top             =   3660
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   412
      Left            =   10875
      TabIndex        =   418
      Top             =   3450
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   411
      Left            =   10875
      TabIndex        =   417
      Top             =   3240
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   410
      Left            =   10875
      TabIndex        =   416
      Top             =   3030
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   409
      Left            =   10875
      TabIndex        =   415
      Top             =   2820
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   408
      Left            =   10575
      TabIndex        =   414
      Top             =   2820
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   407
      Left            =   10575
      TabIndex        =   413
      Top             =   3030
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   406
      Left            =   10575
      TabIndex        =   412
      Top             =   3240
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   405
      Left            =   10575
      TabIndex        =   411
      Top             =   3450
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   404
      Left            =   10575
      TabIndex        =   410
      Top             =   3660
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   403
      Left            =   10575
      TabIndex        =   409
      Top             =   3870
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   402
      Left            =   10575
      TabIndex        =   408
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   401
      Left            =   10575
      TabIndex        =   407
      Top             =   4290
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10590
      TabIndex        =   406
      Top             =   2610
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10590
      TabIndex        =   405
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10590
      TabIndex        =   404
      Top             =   2190
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10590
      TabIndex        =   403
      Top             =   1980
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10590
      TabIndex        =   402
      Top             =   1770
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10590
      TabIndex        =   401
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10590
      TabIndex        =   400
      Top             =   1350
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10185
      TabIndex        =   399
      Top             =   1170
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10185
      TabIndex        =   398
      Top             =   1395
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10185
      TabIndex        =   397
      Top             =   1605
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10185
      TabIndex        =   396
      Top             =   1815
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10185
      TabIndex        =   395
      Top             =   2025
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10185
      TabIndex        =   394
      Top             =   2235
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10185
      TabIndex        =   393
      Top             =   2445
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10185
      TabIndex        =   392
      Top             =   2655
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10170
      TabIndex        =   391
      Top             =   4335
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10170
      TabIndex        =   390
      Top             =   4125
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10170
      TabIndex        =   389
      Top             =   3915
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10170
      TabIndex        =   388
      Top             =   3705
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10170
      TabIndex        =   387
      Top             =   3495
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10170
      TabIndex        =   386
      Top             =   3285
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10170
      TabIndex        =   385
      Top             =   3075
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10170
      TabIndex        =   384
      Top             =   2865
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   383
      Top             =   2985
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   382
      Top             =   3195
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   381
      Top             =   3405
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   380
      Top             =   3615
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   379
      Top             =   3825
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   378
      Top             =   4035
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   377
      Top             =   4245
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   376
      Top             =   4455
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   375
      Top             =   2775
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   374
      Top             =   2565
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   373
      Top             =   2355
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   372
      Top             =   2145
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   371
      Top             =   1935
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   370
      Top             =   1725
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   369
      Top             =   1515
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   368
      Top             =   1290
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10530
      TabIndex        =   367
      Top             =   1065
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10530
      TabIndex        =   366
      Top             =   1290
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10530
      TabIndex        =   365
      Top             =   1500
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10530
      TabIndex        =   364
      Top             =   1710
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10530
      TabIndex        =   363
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10530
      TabIndex        =   362
      Top             =   2130
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10530
      TabIndex        =   361
      Top             =   2340
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10530
      TabIndex        =   360
      Top             =   2550
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10515
      TabIndex        =   359
      Top             =   4230
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10515
      TabIndex        =   358
      Top             =   4020
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10515
      TabIndex        =   357
      Top             =   3810
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10515
      TabIndex        =   356
      Top             =   3600
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10515
      TabIndex        =   355
      Top             =   3390
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10515
      TabIndex        =   354
      Top             =   3180
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10515
      TabIndex        =   353
      Top             =   2970
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10515
      TabIndex        =   352
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11040
      TabIndex        =   351
      Top             =   3390
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11040
      TabIndex        =   350
      Top             =   3600
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11040
      TabIndex        =   349
      Top             =   3810
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11040
      TabIndex        =   348
      Top             =   4020
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11040
      TabIndex        =   347
      Top             =   4230
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11040
      TabIndex        =   346
      Top             =   4440
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11040
      TabIndex        =   345
      Top             =   4650
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11055
      TabIndex        =   344
      Top             =   3180
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11055
      TabIndex        =   343
      Top             =   2970
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11055
      TabIndex        =   342
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11055
      TabIndex        =   341
      Top             =   2550
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11055
      TabIndex        =   340
      Top             =   2340
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11055
      TabIndex        =   339
      Top             =   2130
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11055
      TabIndex        =   338
      Top             =   1695
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   337
      Top             =   1335
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   336
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   335
      Top             =   1770
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   334
      Top             =   1980
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   333
      Top             =   2190
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   332
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   331
      Top             =   2610
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   330
      Top             =   2820
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10980
      TabIndex        =   329
      Top             =   4290
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10980
      TabIndex        =   328
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10980
      TabIndex        =   327
      Top             =   3870
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10980
      TabIndex        =   326
      Top             =   3660
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10980
      TabIndex        =   325
      Top             =   3450
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10980
      TabIndex        =   324
      Top             =   3240
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10980
      TabIndex        =   323
      Top             =   3030
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10650
      TabIndex        =   322
      Top             =   3030
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10650
      TabIndex        =   321
      Top             =   3240
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10650
      TabIndex        =   320
      Top             =   3450
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10650
      TabIndex        =   319
      Top             =   3660
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10650
      TabIndex        =   318
      Top             =   3870
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10650
      TabIndex        =   317
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10650
      TabIndex        =   316
      Top             =   4290
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10665
      TabIndex        =   315
      Top             =   2820
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10665
      TabIndex        =   314
      Top             =   2610
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10665
      TabIndex        =   313
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10665
      TabIndex        =   312
      Top             =   2190
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10665
      TabIndex        =   311
      Top             =   1980
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10665
      TabIndex        =   310
      Top             =   1770
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10665
      TabIndex        =   309
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11070
      TabIndex        =   308
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11070
      TabIndex        =   307
      Top             =   1665
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11070
      TabIndex        =   306
      Top             =   1875
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11070
      TabIndex        =   305
      Top             =   2085
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11070
      TabIndex        =   304
      Top             =   2295
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11070
      TabIndex        =   303
      Top             =   2505
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11070
      TabIndex        =   302
      Top             =   2715
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11070
      TabIndex        =   301
      Top             =   2925
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11055
      TabIndex        =   300
      Top             =   4605
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11055
      TabIndex        =   299
      Top             =   4395
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11055
      TabIndex        =   298
      Top             =   4185
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11055
      TabIndex        =   297
      Top             =   3975
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11055
      TabIndex        =   296
      Top             =   3765
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11055
      TabIndex        =   295
      Top             =   3555
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11055
      TabIndex        =   294
      Top             =   3345
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11055
      TabIndex        =   293
      Top             =   3135
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10305
      TabIndex        =   292
      Top             =   3300
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10305
      TabIndex        =   291
      Top             =   3510
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10305
      TabIndex        =   290
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10305
      TabIndex        =   289
      Top             =   3930
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10305
      TabIndex        =   288
      Top             =   4140
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10305
      TabIndex        =   287
      Top             =   4350
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10305
      TabIndex        =   286
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   285
      Top             =   3090
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   284
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   283
      Top             =   2670
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   282
      Top             =   2460
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   281
      Top             =   2250
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   280
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   279
      Top             =   1830
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   278
      Top             =   1605
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   277
      Top             =   1530
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   276
      Top             =   1755
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   275
      Top             =   1965
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   274
      Top             =   2175
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   273
      Top             =   2385
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   272
      Top             =   2595
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   271
      Top             =   2805
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   270
      Top             =   3015
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11085
      TabIndex        =   269
      Top             =   4485
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11085
      TabIndex        =   268
      Top             =   4275
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11085
      TabIndex        =   267
      Top             =   4065
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11085
      TabIndex        =   266
      Top             =   3855
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11085
      TabIndex        =   265
      Top             =   3645
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11085
      TabIndex        =   264
      Top             =   3435
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11085
      TabIndex        =   263
      Top             =   3225
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11190
      TabIndex        =   262
      Top             =   3210
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11190
      TabIndex        =   261
      Top             =   3420
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11190
      TabIndex        =   260
      Top             =   3630
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11190
      TabIndex        =   259
      Top             =   3840
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11190
      TabIndex        =   258
      Top             =   4050
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11190
      TabIndex        =   257
      Top             =   4260
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11190
      TabIndex        =   256
      Top             =   4470
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11190
      TabIndex        =   255
      Top             =   4680
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11205
      TabIndex        =   254
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11205
      TabIndex        =   253
      Top             =   2790
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11205
      TabIndex        =   252
      Top             =   2580
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11205
      TabIndex        =   251
      Top             =   2370
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11205
      TabIndex        =   250
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11205
      TabIndex        =   249
      Top             =   1950
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11205
      TabIndex        =   248
      Top             =   1740
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11205
      TabIndex        =   247
      Top             =   1515
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10425
      TabIndex        =   246
      Top             =   1590
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10425
      TabIndex        =   245
      Top             =   1815
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10425
      TabIndex        =   244
      Top             =   2025
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10425
      TabIndex        =   243
      Top             =   2235
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10425
      TabIndex        =   242
      Top             =   2445
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10425
      TabIndex        =   241
      Top             =   2655
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10425
      TabIndex        =   240
      Top             =   2865
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10425
      TabIndex        =   239
      Top             =   3075
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   238
      Top             =   4755
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   237
      Top             =   4545
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   236
      Top             =   4335
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   235
      Top             =   4125
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   234
      Top             =   3915
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   233
      Top             =   3705
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   232
      Top             =   3495
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   231
      Top             =   3285
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11160
      TabIndex        =   230
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11160
      TabIndex        =   229
      Top             =   3330
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11160
      TabIndex        =   228
      Top             =   3540
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11160
      TabIndex        =   227
      Top             =   3750
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11160
      TabIndex        =   226
      Top             =   3960
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11160
      TabIndex        =   225
      Top             =   4170
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11160
      TabIndex        =   224
      Top             =   4380
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11160
      TabIndex        =   223
      Top             =   4590
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11175
      TabIndex        =   222
      Top             =   2910
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11175
      TabIndex        =   221
      Top             =   2700
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11175
      TabIndex        =   220
      Top             =   2490
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11175
      TabIndex        =   219
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11175
      TabIndex        =   218
      Top             =   2070
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11175
      TabIndex        =   217
      Top             =   1860
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11175
      TabIndex        =   216
      Top             =   1650
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11175
      TabIndex        =   215
      Top             =   1425
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H009BF8FD&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
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
      Left            =   10200
      TabIndex        =   214
      Top             =   600
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10770
      TabIndex        =   213
      Top             =   1545
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10770
      TabIndex        =   212
      Top             =   1755
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10770
      TabIndex        =   211
      Top             =   1965
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10770
      TabIndex        =   210
      Top             =   2175
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10770
      TabIndex        =   209
      Top             =   2385
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10770
      TabIndex        =   208
      Top             =   2595
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10770
      TabIndex        =   207
      Top             =   2805
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   206
      Top             =   4485
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   205
      Top             =   4275
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   204
      Top             =   4065
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   203
      Top             =   3855
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   202
      Top             =   3645
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   201
      Top             =   3435
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   200
      Top             =   3225
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   199
      Top             =   3015
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11085
      TabIndex        =   198
      Top             =   3015
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11085
      TabIndex        =   197
      Top             =   3225
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11085
      TabIndex        =   196
      Top             =   3435
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11085
      TabIndex        =   195
      Top             =   3645
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11085
      TabIndex        =   194
      Top             =   3855
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11085
      TabIndex        =   193
      Top             =   4065
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11085
      TabIndex        =   192
      Top             =   4275
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11085
      TabIndex        =   191
      Top             =   4485
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   190
      Top             =   2805
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   189
      Top             =   2595
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   188
      Top             =   2385
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   187
      Top             =   2175
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   186
      Top             =   1965
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   185
      Top             =   1755
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   184
      Top             =   1545
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11100
      TabIndex        =   183
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11160
      TabIndex        =   182
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "WinAmp Caller"
      BeginProperty Font 
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
      Left            =   10155
      TabIndex        =   181
      Top             =   210
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11160
      TabIndex        =   180
      Top             =   2115
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11160
      TabIndex        =   179
      Top             =   2325
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11160
      TabIndex        =   178
      Top             =   2535
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11160
      TabIndex        =   177
      Top             =   2745
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11160
      TabIndex        =   176
      Top             =   2955
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11160
      TabIndex        =   175
      Top             =   3165
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11145
      TabIndex        =   174
      Top             =   4845
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11145
      TabIndex        =   173
      Top             =   4635
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11145
      TabIndex        =   172
      Top             =   4425
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11145
      TabIndex        =   171
      Top             =   4215
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11145
      TabIndex        =   170
      Top             =   4005
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11145
      TabIndex        =   169
      Top             =   3795
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11145
      TabIndex        =   168
      Top             =   3585
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   11145
      TabIndex        =   167
      Top             =   3375
      Width           =   135
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
      Left            =   6780
      TabIndex        =   165
      Top             =   4755
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10620
      TabIndex        =   164
      Top             =   2745
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10620
      TabIndex        =   163
      Top             =   2955
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10620
      TabIndex        =   162
      Top             =   3165
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10620
      TabIndex        =   161
      Top             =   3375
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10620
      TabIndex        =   160
      Top             =   3585
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10620
      TabIndex        =   159
      Top             =   3795
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10620
      TabIndex        =   158
      Top             =   4005
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10620
      TabIndex        =   157
      Top             =   4215
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   156
      Top             =   2535
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   155
      Top             =   2325
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   154
      Top             =   2115
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   153
      Top             =   1905
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   152
      Top             =   1695
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   151
      Top             =   1485
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   150
      Top             =   1275
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   149
      Top             =   1050
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10950
      TabIndex        =   148
      Top             =   1275
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10950
      TabIndex        =   147
      Top             =   1500
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10950
      TabIndex        =   146
      Top             =   1710
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10950
      TabIndex        =   145
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10950
      TabIndex        =   144
      Top             =   2130
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10950
      TabIndex        =   143
      Top             =   2340
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10950
      TabIndex        =   142
      Top             =   2550
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10950
      TabIndex        =   141
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10935
      TabIndex        =   140
      Top             =   4440
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10935
      TabIndex        =   139
      Top             =   4230
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10935
      TabIndex        =   138
      Top             =   4020
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10935
      TabIndex        =   137
      Top             =   3810
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10935
      TabIndex        =   136
      Top             =   3600
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10935
      TabIndex        =   135
      Top             =   3390
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10935
      TabIndex        =   134
      Top             =   3180
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10935
      TabIndex        =   133
      Top             =   2970
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10275
      TabIndex        =   132
      Top             =   2850
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10275
      TabIndex        =   131
      Top             =   3060
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10275
      TabIndex        =   130
      Top             =   3270
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10275
      TabIndex        =   129
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10275
      TabIndex        =   128
      Top             =   3690
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10275
      TabIndex        =   127
      Top             =   3900
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10275
      TabIndex        =   126
      Top             =   4110
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10275
      TabIndex        =   125
      Top             =   4320
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10290
      TabIndex        =   124
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10290
      TabIndex        =   123
      Top             =   2430
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10290
      TabIndex        =   122
      Top             =   2220
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10290
      TabIndex        =   121
      Top             =   2010
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10290
      TabIndex        =   120
      Top             =   1800
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10290
      TabIndex        =   119
      Top             =   1590
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10290
      TabIndex        =   118
      Top             =   1380
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10290
      TabIndex        =   117
      Top             =   1155
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10695
      TabIndex        =   116
      Top             =   1110
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10695
      TabIndex        =   115
      Top             =   1335
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10695
      TabIndex        =   114
      Top             =   1545
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10695
      TabIndex        =   113
      Top             =   1755
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10695
      TabIndex        =   112
      Top             =   1965
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10695
      TabIndex        =   111
      Top             =   2175
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10695
      TabIndex        =   110
      Top             =   2385
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10695
      TabIndex        =   109
      Top             =   2595
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10680
      TabIndex        =   108
      Top             =   4275
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10680
      TabIndex        =   107
      Top             =   4065
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10680
      TabIndex        =   106
      Top             =   3855
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10680
      TabIndex        =   105
      Top             =   3645
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10680
      TabIndex        =   104
      Top             =   3435
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10680
      TabIndex        =   103
      Top             =   3225
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10680
      TabIndex        =   102
      Top             =   3015
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10680
      TabIndex        =   101
      Top             =   2805
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10980
      TabIndex        =   100
      Top             =   2805
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10980
      TabIndex        =   99
      Top             =   3015
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10980
      TabIndex        =   98
      Top             =   3225
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10980
      TabIndex        =   97
      Top             =   3435
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10980
      TabIndex        =   96
      Top             =   3645
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10980
      TabIndex        =   95
      Top             =   3855
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10980
      TabIndex        =   94
      Top             =   4065
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10980
      TabIndex        =   93
      Top             =   4275
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   92
      Top             =   2595
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   91
      Top             =   2385
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   90
      Top             =   2175
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   89
      Top             =   1965
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   88
      Top             =   1755
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   87
      Top             =   1545
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   86
      Top             =   1335
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10995
      TabIndex        =   85
      Top             =   1110
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   84
      Top             =   1245
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   83
      Top             =   1470
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   82
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   81
      Top             =   1890
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   80
      Top             =   2100
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   79
      Top             =   2310
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   78
      Top             =   2520
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10755
      TabIndex        =   77
      Top             =   2730
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10740
      TabIndex        =   76
      Top             =   4410
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10740
      TabIndex        =   75
      Top             =   4200
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10740
      TabIndex        =   74
      Top             =   3990
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10740
      TabIndex        =   73
      Top             =   3780
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10740
      TabIndex        =   72
      Top             =   3570
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10740
      TabIndex        =   71
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10740
      TabIndex        =   70
      Top             =   3150
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10740
      TabIndex        =   69
      Top             =   2940
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   68
      Top             =   2955
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   67
      Top             =   3165
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   66
      Top             =   3375
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   65
      Top             =   3585
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   64
      Top             =   3795
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   63
      Top             =   4005
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   62
      Top             =   4215
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10635
      TabIndex        =   61
      Top             =   4425
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10650
      TabIndex        =   60
      Top             =   2745
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10650
      TabIndex        =   59
      Top             =   2535
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10650
      TabIndex        =   58
      Top             =   2325
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10650
      TabIndex        =   57
      Top             =   2115
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10650
      TabIndex        =   56
      Top             =   1905
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10650
      TabIndex        =   55
      Top             =   1695
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10650
      TabIndex        =   54
      Top             =   1485
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10650
      TabIndex        =   53
      Top             =   1260
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   52
      Top             =   1185
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   51
      Top             =   1410
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   50
      Top             =   1620
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   49
      Top             =   1830
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   48
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   47
      Top             =   2250
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   46
      Top             =   2460
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10320
      TabIndex        =   45
      Top             =   2670
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10305
      TabIndex        =   44
      Top             =   4350
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10305
      TabIndex        =   43
      Top             =   4140
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10305
      TabIndex        =   42
      Top             =   3930
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10305
      TabIndex        =   41
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10305
      TabIndex        =   40
      Top             =   3510
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10305
      TabIndex        =   39
      Top             =   3300
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10305
      TabIndex        =   38
      Top             =   3090
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10305
      TabIndex        =   37
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   36
      Top             =   1095
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   35
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   34
      Top             =   1530
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   33
      Top             =   1740
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   32
      Top             =   1950
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   31
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   30
      Top             =   2370
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10410
      TabIndex        =   29
      Top             =   2580
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10395
      TabIndex        =   28
      Top             =   4260
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10395
      TabIndex        =   27
      Top             =   4050
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10395
      TabIndex        =   26
      Top             =   3840
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10395
      TabIndex        =   25
      Top             =   3630
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10395
      TabIndex        =   24
      Top             =   3420
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10395
      TabIndex        =   23
      Top             =   3210
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10395
      TabIndex        =   22
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10395
      TabIndex        =   21
      Top             =   2790
      Width           =   135
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
      BackColor       =   &H00404040&
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
      BackColor       =   &H00DDFFFE&
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
      Left            =   135
      TabIndex        =   16
      Top             =   75
      Width           =   3945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   15
      Top             =   2925
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   14
      Top             =   3135
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   13
      Top             =   3345
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   12
      Top             =   3555
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   11
      Top             =   3765
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   10
      Top             =   3975
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   9
      Top             =   4185
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10830
      TabIndex        =   8
      Top             =   4395
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   7
      Top             =   2715
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   6
      Top             =   2505
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   5
      Top             =   2295
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   4
      Top             =   2085
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   3
      Top             =   1875
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   2
      Top             =   1665
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   1
      Top             =   1455
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "A"
      BeginProperty Font 
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
      Left            =   10845
      TabIndex        =   0
      Top             =   1230
      Width           =   135
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu Mnu_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_STAY_OPEN_WHEN_EXECUTE 
      Caption         =   "STAY OPEN WHEN EXECUTE"
   End
   Begin VB.Menu Mnu_LoadFolder 
      Caption         =   "Load Folder"
   End
   Begin VB.Menu Mnu_Load_FolderFile 
      Caption         =   "Load Folder Target"
   End
   Begin VB.Menu Mnu_Sync 
      Caption         =   "Sync"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_Clipexe 
      Caption         =   "Clip EXE Path"
   End
   Begin VB.Menu MNU_SHOW_SCANPATH 
      Caption         =   "SHOW PATH SCAN"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'USE OPTION CLIP EXE NAME AND MOVE TO FOLDER


Public LINK_DATA

Public GetOsBitness_2

Dim Form1_Width, Form1_Height
Dim X_Y_DONE_ONCE

Dim XXT2

Dim ClipEXE

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer

'-----------------------------------------------
'AUTOSIZE LISTVIEW COLUMN
Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Private Const LVM_FIRST = &H1000
'-----------------------------------------------
'Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column _
' As ColumnHeader = Nothing)
'
' Dim C As ColumnHeader
' If Column Is Nothing Then
'  For Each C In LV.ColumnHeaders
'   SendMessage LV.hWnd, LVM_FIRST + 30, C.Index - 1, -1
'  Next
' Else
'  SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
' End If
' LV.Refresh
'End Sub
'-----------------------------------------------


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column _
 As ColumnHeader = Nothing)

 Dim C As ColumnHeader
 If Column Is Nothing Then
  For Each C In LV.ColumnHeaders
   SendMessage LV.hWnd, LVM_FIRST + 30, C.Index - 1, -1
  Next
 Else
  SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
 End If
 LV.Refresh
End Sub



Private Sub Combo1_Click()
Counter = Combo1.ListIndex + 1
A1$ = A4$(Counter)
B1$ = B4$(Counter)
C1$ = C4$(Counter)

SetTrueToLoadLast = True
Call Label1_Click(0)

End Sub

Public Function GetOsBitness() As String
    Dim ProcessorSet As Object
    Dim CPU As Object

    Set ProcessorSet = GetObject("Winmgmts:"). _
        ExecQuery("SELECT * FROM Win32_Processor")
    For Each CPU In ProcessorSet
        GetOsBitness = CStr(CPU.AddressWidth)
    Next
End Function

Private Sub Form_Load()



If App.PrevInstance = True Then End

GetOsBitness_2 = GetOsBitness

'On Error Resume Next
'Open App.Path + "\VarData.txt" For Input As #1
'DaysToScan = Val(Input(LOF(1), 1))
'Close
'On Error GoTo 0

'If DaysToScan < 3 Then DaysToScan = 3
'DaysToScan = 50

Set FS = CreateObject("Scripting.FileSystemObject")

ReDim A4$(500), B4$(500), C4$(500)

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

Open TEXT_PATH_2 + "\" + OIP$ + " Loads Logg 02.txt" For Binary As #1

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

If Dir$(TEXT_PATH_2 + "\" + OIP$ + " Loads Logg 01.txt") <> "" Then
    Open TEXT_PATH_2 + "\" + OIP$ + " Loads Logg 01.txt" For Input As #1
    Line Input #1, A1$
    Line Input #1, B1$
    Line Input #1, C1$
    Close #1
End If
'---- Load Last File Click Here ----
    
TitleLbl = "Last File -- " + C1$

Form1.Top = 0
Form1.Left = 0
On Local Error GoTo 0
On Error GoTo 0
Call FormStart.FormStartLoader

Timer_SUB_CODE_LOAD.Enabled = True

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

Private Sub MNU_STAY_OPEN_WHEN_EXECUTE_Click()

MNU_STAY_OPEN_WHEN_EXECUTE_VAR = True

End Sub

Private Sub Timer_KEY_CODE_Timer()
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
FontSizez = 10
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


Lbl2.Top = 0
Lbl2.Left = 0


x = Lbl2.Height + LABEL_GAP


Check1.Top = x
Check1.Left = (Form1.Width / 2) + 10
Check1.Width = Form1.Width / 2

tg = x

x = x + TitleLbl.Height + LABEL_GAP
TX = x

w = 0

RD = 0

DoEvents

Dim RECT1 As RECT
Dim RECT2 As RECT

rr = FindWindow("Shell_TrayWnd", vbNullString)
rr2 = GetWindowRect(rr, RECT1)

rr = FindWindow("MOM Class", vbNullString)
rr2 = GetWindowRect(rr, RECT2)


XX1 = (RECT1.Bottom - RECT1.Top) * Screen.TwipsPerPixelY
XX1 = XX1 + (RECT2.Bottom - RECT2.Top) * Screen.TwipsPerPixelY

SH = Screen.Height - (XX1 + 1300) 'higer = smaller


For R = ScanPath.ListView1.ListItems.Count To 1 Step -1
    If InStr(ScanPath.ListView1.ListItems.Item(R).SubItems(1), "\_gsdata_") > 0 Then
        ScanPath.ListView1.ListItems.Remove (R)
    End If
Next

x = TX
For R = 1 To ScanPath.ListView1.ListItems.Count
    If Label1(R - 1).Top > SH Then
        x = TX
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
    If txr = -1 Then Label1(R).BackColor = Label1(836).BackColor
    If txr = 0 Then Label1(R).BackColor = Label1(837).BackColor
    If R > AStart Then
        'If InStr(LCase(Right(B1$, 4)), ".vbp") > 0 Then proprojects = proprojects + 1
            
        If txr = -1 Then Label1(R).BackColor = Label1(175).BackColor
        If txr = 0 Then Label1(R).BackColor = Label1(208).BackColor
    End If
    ttg$ = ScanPath.ListView1.ListItems.Item(R)
    If InStr(ttg$, "_gsdata_") = 0 Then
        Mid$(ttg$, 1, 1) = UCase$(Mid$(ttg$, 1, 1))
        If InStr(ttg$, ".lnk") > 0 Then ttg$ = Mid$(ttg$, 1, InStrRev(ttg$, ".") - 1)
        
        Label1(R).Caption = Format$(R, "00") + ". " + ttg$
        x = x + Label1(R).Height + LABEL_GAP
        fheight = Label1(R).Top + Label1(R).Height + 420
        If fheight > fheightx Then fheightx = fheight
        RD = RD + 1
    End If
Next
DoEvents

x = TX
td = 29
xgag = 0
Dim xy(20 * 2)

For R = 1 To ScanPath.ListView1.ListItems.Count
    If Label1(R - 1).Top > SH Then
        If tig = 0 Then tig = R
        x = TX
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
'Tolerance
If Label1(R).Left + Label1(R).Width > Screen.Width + 1500 Then
    Label1(R).Visible = False
    A1$ = ScanPath.ListView1.ListItems.Item(R).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(R)
    
    'If InStr(LCase(Right(B1$, 4)), ".vbp") > 0 Then proprojects = proprojects - 1
    'If InStr(LCase(Right(B1$, 4)), ".vbp") > 0 Then proprojects = proprojects - 1
    
    If halo = 0 Then
    halo = R
    On Error Resume Next
        Set F = FS.getfile(A1$ + B1$)
        ttdate = DateDiff("d", F.datelastmodified, Now)
    On Error GoTo 0
    
    End If
End If

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

iu = halo
If iu = 0 Then
iu = ScanPath.ListView1.ListItems.Count
Set F = FS.getfile(A1$ + B1$)
ttdate = DateDiff("d", F.datelastmodified, Now)
End If

Form1_Width = Label1(iu - 1).Left + Label1(iu - 1).Width + 280
If Form1_Width > Screen.Width Then Form1_Width = Screen.Width - 120
Form1_Height = fheightx + 900

'Form1.Width = Form1_Width
'Form1.Height = Form1_Height
'Center screen
'Form1.Top = 0 'Screen.Height / 2 - Form1.Height / 2
'Form1.Left = Screen.Width / 2 - Form1.Width / 2


Lbl2.Width = Form1.Width
TitleLbl.Top = tg
TitleLbl.Left = 0
Combo1.Top = tg
Combo1.Left = (Form1.Width / 2) + 15

If Combo1.Visible = False Then
    TitleLbl.Width = Form1.Width
Else
    TitleLbl.Width = Form1.Width / 2
End If
Combo1.Width = (Form1.Width / 2) - 150
Lbl3.Top = fheight - 410
Lbl3.Width = Form1.Width
Lbl3.Left = 0


Form1.Refresh

Form1.Lbl2.Caption = TXmsg + " -- Projects From " + Str(DaysToScan) + " Days -" + Str(DaysToScanYears) + " Years " + Str(Year(Now - DaysToScan)) + " -- Last Project" + Str(ttdate) + " Days Ago -- " + Format$(F.datelastmodified, "DDD DD-MMM-YY") + " -- Projects Found" + Str(ProProjects) + " Inn" + Str(DaysToScanYears) + " Years -- Projects in Script =" + Str(ProProjects2)

End Sub


Private Sub Form_Unload(Cancel As Integer)
'SET ALL TIMERS IN ALL FORMS TO ENABLED = FALSE

TERMINATE_FORMS = True
On Error Resume Next
    For i = 0 To Forms.Count - 1
        For Each Control In Forms(i).Controls
            'If InStr(UCase(Control.Name), "TIMER") > 0 Then
                Control.Enabled = False
            'End If
        Next
    Next i
On Error GoTo 0
   
Dim Form As Form
For Each Form In Forms
    Unload Form
    Set Form = Nothing
Next Form
'Unload Form1


'End
End Sub

Private Sub Label1_Click(Index As Integer)

If MNU_STAY_OPEN_WHEN_EXECUTE_VAR = False Then
Form1.Visible = False
End If

If SetTrueToLoadLast = False Then
    A1$ = ScanPath.ListView1.ListItems.Item(Index).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(Index)
    C1$ = Form1.Label1(Index).Caption
    ChDrive Mid$(A1$, 1, 2)
    ChDir A1$

Else


    ChDrive Mid$(A1$, 1, 2)
    ChDir A1$
    vLaunch A1$ + B1$, vbNullString
    End
End If

ChDrive Mid$(A1$, 1, 2)
ChDir A1$

If LoadFolder = True Then
    On Error Resume Next
    Shell "explorer /e, /select, " + A1$ + B1$, vbNormalFocus
    If Err.Number > 0 Then
        MsgBox "Error Run Check With Administrator" + vbCrLf + vbCrLf + "explorer /e, /select, " + A1$ + B1$, vbMsgBoxSetForeground
    End If
    If MNU_STAY_OPEN_WHEN_EXECUTE_VAR = True Then Exit Sub
    End
End If

If ClipEXE = True Then
    Clipboard.Clear
    Clipboard.SetText A1$ + B1$
    If MNU_STAY_OPEN_WHEN_EXECUTE_VAR = True Then Exit Sub
    End
End If

Call SaveLoggs

Call FormStart.LabelClick(Index)


End Sub

Private Sub Lbl2_Click()

Shell "explorer /e, E:\01 VB Shell Folders\00 " + App.ExeName, vbNormalFocus

End Sub

Private Sub Mnu_Clipexe_Click()
    ClipEXE = True
End Sub

Private Sub MNU_EXIT_Click()
End
End Sub

Private Sub Mnu_Load_Folderfile_Click()

LoadFolderFile = True

End Sub

Private Sub Mnu_LoadFolder_Click()
LoadFolder = True
End Sub

Private Sub MNU_SHOW_SCANPATH_Click()
ScanPath.Show
On Error Resume Next
For Each Control In ScanPath.Controls
    Control.Enabled = False
    Control.Visible = False
Next
On Error GoTo 0

ScanPath.ListView3.Enabled = True
ScanPath.ListView3.Visible = True

SCREEN_WIDTH_REDUCTION = 1000

ScanPath.ListView3.Top = 0
ScanPath.ListView3.Left = 0
ScanPath.ListView3.Height = Screen.Height - 200 * Screen.TwipsPerPixelY
ScanPath.ListView3.Width = Screen.Width - 200 - SCREEN_WIDTH_REDUCTION '* Screen.TwipsPerPixelX

ScanPath.WindowState = vbNormal

ScanPath.Height = ScanPath.ListView3.Height + 32 * Screen.TwipsPerPixelY
ScanPath.Height = Screen.Height / 1.4
ScanPath.Width = Screen.Width - SCREEN_WIDTH_REDUCTION 'ScanPath.ListView3.Width + 7 * Screen.TwipsPerPixelX

'Center screen
ScanPath.Top = 0 'Screen.Height / 2 - ScanPath.Height / 2
ScanPath.Left = Screen.Width / 2 - ScanPath.Width / 2

Call LV_AutoSizeColumn(ScanPath.ListView3, ScanPath.ListView3.ColumnHeaders.Item(1))
Call LV_AutoSizeColumn(ScanPath.ListView3, ScanPath.ListView3.ColumnHeaders.Item(2))
Call LV_AutoSizeColumn(ScanPath.ListView3, ScanPath.ListView3.ColumnHeaders.Item(4))
Call LV_AutoSizeColumn(ScanPath.ListView3, ScanPath.ListView3.ColumnHeaders.Item(5))

For i = 1 To ScanPath.ListView3.ListItems.Count
    ScanPath.ListView3.SelectedItem = ScanPath.ListView3.ListItems(i)
    ScanPath.ListView3.SelectedItem.EnsureVisible
    DoEvents
Next
ScanPath.ListView3.SelectedItem = ScanPath.ListView3.ListItems(1)
ScanPath.ListView3.SelectedItem.EnsureVisible

End Sub

Private Sub Mnu_Sync_Click()

ScanPath.ListView1.ListItems.Clear

Call FormStart.FormStartLoader

Call SubCode


End Sub

Private Sub MNU_VB_FOLDER_Click()
    
    Beep
    CODER_EXE_FILE_NAME_1 = App.Path + "\" + Dir(App.Path + "\*.VBP")
    On Error Resume Next
    Shell "EXPLORER /SELECT, """ + CODER_EXE_FILE_NAME_1 + """", vbNormalFocus
    If Err.Number > 0 Then
        MsgBox "Error Run Check With Administrator" + vbCrLf + vbCrLf + "EXPLORER /SELECT, """ + CODER_EXE_FILE_NAME_1 + """", vbMsgBoxSetForeground
    End If
    EXIT_TRUE = True
    Unload Me

End Sub

Private Sub Mnu_VB_ME_Click()

VBPATH = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
If Dir(VBPATH) = "" Then
    VBPATH = "C:\Program Files (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
End If

On Error Resume Next
Shell VBPATH + " """ + App.Path + "\" + App.ExeName + ".VBP""", vbNormalFocus
If Err.Number > 0 Then
    MsgBox "Error Run Check With Administrator" + vbCrLf + vbCrLf + VBPATH + " """ + App.Path + "\" + App.ExeName + ".VBP""", vbMsgBoxSetForeground
End If
End
End Sub

Private Sub Timer_TITLE_COLOR_END_DONE_Timer()

Timer_TITLE_COLOR_END_DONE.Enabled = False
TitleLbl.BackColor = Label1(838).BackColor
'TitleLbl.Caption = Replace(TitleLbl.Caption, " ____ DONE \/\/", "")
End Sub

Private Sub Timer2_Timer()

XXT2 = XXT2 + 1
If XXT2 > ScanPath.ListView1.ListItems.Count Then
    EXIT_TIMER_2_NOW = True
End If
If Form1.Label1(XXT2).BackColor = Label1(175).BackColor Then
    ' FIRST COLOUR OF 2 _ HERE
    EXIT_TIMER_2_NOW = True
End If
If Form1.Label1(XXT2).BackColor = Label1(208).BackColor Then
    EXIT_TIMER_2_NOW = True
End If

If EXIT_TIMER_2_NOW = True Then
    TitleLbl.BackColor = Label1(836).BackColor
    'TitleLbl.Caption = TitleLbl.Caption + " ____ DONE \/\/"
    'TitleLbl.BackColor = Label1(838).BackColor
    Timer_TITLE_COLOR_END_DONE.Enabled = True

    Timer2.Enabled = False
    Exit Sub
End If

A1$ = ScanPath.ListView1.ListItems.Item(XXT2).SubItems(1)
B1$ = ScanPath.ListView1.ListItems.Item(XXT2)
C1$ = Form1.Label1(XXT2).Caption

X_COLOR = Form1.Label1(XXT2).BackColor
Form1.Label1(XXT2).BackColor = &HFFFFFF
Form1.Label1(XXT2).Refresh
DoEvents
Sleep 10

On Local Error GoTo 0
On Error GoTo 0

'GoTo skip:
LINK_DATA = A1$ + B1$
If InStr(LCase(B1$), ".lnk") > 0 Then
    Open A1$ + B1$ For Binary As #1
    rr$ = Space$(LOF(1))
    Get #1, 1, rr$
    Close #1
    If InStr(rr$, ":\") = 0 Then
        
        i = MsgBox("INVAILD LINK NAUGHT SIZE" + vbCrLf + LINK_DATA, vbMsgBoxSetForeground)
        
        i = MsgBox("INVAILD LINK" + vbCrLf + LINK_DATA + vbCrLf + "DO YOU WANT ME TO TAKE YOU THERE", vbYesNo + vbMsgBoxSetForeground)
        If i = vbYes Then
            Shell "Explorer.exe /select, " + LINK_DATA, vbNormalFocus
    
    
        End If
        Form1.Label1(XXT2).BackColor = X_COLOR
        Exit Sub
    End If
    tt$ = Mid$(rr$, InStr(105, rr$, ":\") - 1)
    LINK_DATA = Mid$(tt$, 1, InStr(tt$, Chr$(0)) - 1)
End If

ALLOW_OKAY_PASS_VALID_LINK = False
If FS.FileExists(LINK_DATA) = False And FS.FolderExists(LINK_DATA) = False Then
    
    'GetOsBitness_2
    Call TEST_LINK_MODIFY
    
End If

Form1.Label1(XXT2).BackColor = X_COLOR

If FS.FileExists(LINK_DATA) = False And FS.FolderExists(LINK_DATA) = False Then
    Label1(XXT2).BackColor = RGB(250, 127, 127) + QBColor(12)
End If




End Sub

Sub TEST_LINK_MODIFY()

    TEST_D1 = LINK_DATA
    
    For I_LOOP = 1 To 2
        If ALLOW_OKAY_PASS_VALID_LINK = False Then
            If InStr(UCase(TEST_D1), "PROGRAM FILES\") > 0 Then
                TEST_D1 = Replace(UCase(TEST_D1), "PROGRAM FILES\", "PROGRAM FILES (X86)\")
                'CHANGE TO D: AND H: HAPPEN WHEN CAN'T FIND LINK PUT BACK TO C:
                '--------------------------------------------------------------
                Mid(TEST_D1, 1, 1) = "C"
                TEST_D1 = Replace(UCase(TEST_D1), "D:\", "C:\")
            End If
        End If
        If FS.FileExists(TEST_D1) = True Or FS.FolderExists(TEST_D1) = True Then
            ALLOW_OKAY_PASS_VALID_LINK = True
        End If
        
        If ALLOW_OKAY_PASS_VALID_LINK = False Then
            If InStr(UCase(TEST_D1), "PROGRAM FILES (X86)\") > 0 Then
                TEST_D1 = Replace(UCase(TEST_D1), "PROGRAM FILES (X86)\", "PROGRAM FILES\")
                'CHANGE TO D: AND H: HAPPEN WHEN CAN'T FIND LINK PUT BACK TO C:
                '--------------------------------------------------------------
                Mid(TEST_D1, 1, 1) = "C"
            End If
        End If
        If FS.FileExists(TEST_D1) = True Or FS.FolderExists(TEST_D1) = True Then
            ALLOW_OKAY_PASS_VALID_LINK = True
        End If
        
        If ALLOW_OKAY_PASS_VALID_LINK = False Then
            If InStr(UCase(TEST_D1), "MSDN98\98VS") > 0 Then
                TEST_D1 = Replace(UCase(TEST_D1), "MSDN98\98VS", "MSDN98\98VSa")
                'CHANGE TO D: AND H: HAPPEN WHEN CAN'T FIND LINK PUT BACK TO C:
                '-------------------------------------------------------
                Mid(TEST_D1, 1, 1) = "C"
            End If
        End If
        
        If FS.FileExists(TEST_D1) = True Or FS.FolderExists(TEST_D1) = True Then
            ALLOW_OKAY_PASS_VALID_LINK = True
        End If
    Next
    
    If ALLOW_OKAY_PASS_VALID_LINK = True Then
        LINK_DATA = TEST_D1
    End If

End Sub

Private Sub TitleLbl_Click()
Open TEXT_PATH_2 + "\" + OIP$ + " Loads Logg 01.txt" For Input As #1
Line Input #1, A1$
Line Input #1, B1$
Line Input #1, C1$
Close #1

SetTrueToLoadLast = True
Call Label1_Click(0)

End Sub

Sub SaveLoggs()
Open TEXT_PATH_2 + "\" + OIP$ + " Loads Logg 01.txt" For Output As #1
Print #1, A1$
Print #1, B1$
Print #1, C1$
Close #1

Open TEXT_PATH_2 + "\" + OIP$ + " Loads Logg 02.txt" For Append As #1
Print #1, A1$
Print #1, B1$
Print #1, C1$
Close #1

Open TEXT_PATH_2 + "\" + OIP$ + " Loads Logg 03.txt" For Append As #1
Print #1, Format$(Now, "dd-mm-yyyy hh:mm:ss ") + A1$ + B1$
Close #1


End Sub
