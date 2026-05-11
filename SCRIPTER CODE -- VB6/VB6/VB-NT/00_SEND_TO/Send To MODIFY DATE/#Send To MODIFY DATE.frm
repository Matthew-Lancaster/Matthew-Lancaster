VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "FormMODIFY DATE AS CREATED DATE1"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   945
   ClientWidth     =   15720
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   15720
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File2 
      Height          =   675
      Left            =   12048
      TabIndex        =   42
      Top             =   744
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.DirListBox Dir1 
      Height          =   936
      Left            =   11760
      TabIndex        =   41
      Top             =   768
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   13512
      TabIndex        =   39
      Top             =   1728
      Visible         =   0   'False
      Width           =   1104
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   53
      Left            =   4284
      TabIndex        =   59
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   52
      Left            =   4284
      TabIndex        =   58
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   51
      Left            =   4284
      TabIndex        =   57
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   50
      Left            =   4284
      TabIndex        =   56
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   49
      Left            =   4284
      TabIndex        =   55
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   48
      Left            =   4284
      TabIndex        =   54
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   47
      Left            =   4284
      TabIndex        =   53
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   46
      Left            =   4284
      TabIndex        =   52
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   45
      Left            =   4284
      TabIndex        =   51
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   44
      Left            =   4284
      TabIndex        =   50
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   43
      Left            =   4284
      TabIndex        =   49
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   42
      Left            =   4284
      TabIndex        =   48
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   41
      Left            =   4284
      TabIndex        =   47
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   40
      Left            =   4284
      TabIndex        =   46
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   39
      Left            =   4284
      TabIndex        =   45
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   38
      Left            =   4284
      TabIndex        =   44
      Top             =   1788
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "ARRAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   37
      Left            =   2856
      TabIndex        =   43
      Top             =   1824
      Width           =   2040
   End
   Begin VB.Label LABEL_COLOR_SET_EFORE_INDEX_FINGER 
      BackColor       =   &H00FFC0C0&
      Caption         =   "LABEL_COLOR_SET_EFORE_INDEX_FINGER"
      Height          =   300
      Left            =   12072
      TabIndex        =   40
      Top             =   3252
      Visible         =   0   'False
      Width           =   3468
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   36
      Left            =   13620
      TabIndex        =   38
      Top             =   624
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   35
      Left            =   13068
      TabIndex        =   37
      Top             =   696
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   34
      Left            =   12996
      TabIndex        =   36
      Top             =   612
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   33
      Left            =   12780
      TabIndex        =   35
      Top             =   588
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   32
      Left            =   13392
      TabIndex        =   34
      Top             =   768
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   31
      Left            =   12960
      TabIndex        =   33
      Top             =   840
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   30
      Left            =   13152
      TabIndex        =   32
      Top             =   624
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   29
      Left            =   13068
      TabIndex        =   31
      Top             =   240
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   28
      Left            =   12792
      TabIndex        =   30
      Top             =   540
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   27
      Left            =   13548
      TabIndex        =   29
      Top             =   888
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   26
      Left            =   13116
      TabIndex        =   28
      Top             =   564
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   25
      Left            =   13380
      TabIndex        =   27
      Top             =   1104
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   0
      Left            =   13560
      TabIndex        =   26
      Top             =   444
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "PERFORM ON ALL FILES IN FOLDER OR FILE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Index           =   4
      Left            =   108
      TabIndex        =   2
      Top             =   1152
      Width           =   11100
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "FILE - CREATED DATE TO MODIFY DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Index           =   8
      Left            =   120
      TabIndex        =   10
      Top             =   3468
      Width           =   11100
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "BATCH - CREATED DATE TO MODIFY DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Index           =   7
      Left            =   108
      TabIndex        =   8
      Top             =   2904
      Width           =   11100
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "SET ONE DATE HARDCODER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Index           =   9
      Left            =   108
      TabIndex        =   7
      Top             =   4080
      Width           =   11100
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "MODIFY DATE TO CREATED DATE - NOT WORKING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   528
      Index           =   6
      Left            =   96
      TabIndex        =   1
      Top             =   2340
      Width           =   11100
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "NOW DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   5
      Left            =   144
      TabIndex        =   0
      Top             =   1800
      Width           =   2040
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "SET OLDER 1ST DATE TO OTHER IN FOLDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Index           =   10
      Left            =   228
      TabIndex        =   11
      Top             =   4764
      Width           =   11100
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "SET ALL DATE FOLDER THE TEXTFILE HOLD DATE WITHIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   11
      Left            =   24
      TabIndex        =   12
      Top             =   5544
      Width           =   10812
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   12
      Left            =   13344
      TabIndex        =   13
      Top             =   768
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   13
      Left            =   12828
      TabIndex        =   14
      Top             =   768
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   14
      Left            =   13416
      TabIndex        =   15
      Top             =   372
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   15
      Left            =   12804
      TabIndex        =   16
      Top             =   336
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   16
      Left            =   13968
      TabIndex        =   17
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   17
      Left            =   13968
      TabIndex        =   18
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   18
      Left            =   13968
      TabIndex        =   19
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   19
      Left            =   13968
      TabIndex        =   20
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   20
      Left            =   13968
      TabIndex        =   21
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   21
      Left            =   13968
      TabIndex        =   22
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   22
      Left            =   13968
      TabIndex        =   23
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   23
      Left            =   13968
      TabIndex        =   24
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Index           =   24
      Left            =   13968
      TabIndex        =   25
      Top             =   348
      Width           =   396
   End
   Begin VB.Label LABEL_SET 
      Alignment       =   2  'Center
      BackColor       =   &H00FED3D1&
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Index           =   1
      Left            =   12768
      TabIndex        =   4
      Top             =   5028
      Width           =   1632
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "FILE LABEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   528
      Index           =   3
      Left            =   108
      TabIndex        =   9
      Top             =   612
      Width           =   11100
   End
   Begin VB.Label LABEL_SET 
      Caption         =   "FOLDER LABEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   528
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   48
      Width           =   11100
   End
   Begin VB.Label Label_COLOR_GREEN 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label_COLOR_GREEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   12072
      TabIndex        =   6
      Top             =   3924
      Visible         =   0   'False
      Width           =   3072
   End
   Begin VB.Label Label_COLOR_YELLOW 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label_COLOUR_YELLOW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12072
      TabIndex        =   5
      Top             =   3576
      Visible         =   0   'False
      Width           =   3072
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VERSION 
      Caption         =   "VERSION"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB_FOLDER"
   End
   Begin VB.Menu MNU_QUICK_MODE 
      Caption         =   "QUICK MODE -- OFF -- GIVE MESSENGER"
   End
   Begin VB.Menu MNU_CREATE_TEXT_FILE_FOR_DATE 
      Caption         =   "CREATE TEXT FILE FOR DATE"
   End
   Begin VB.Menu MNU_CREATE_FOLDER_DATE_OF_FILE 
      Caption         =   "CREATE FOLDER DATE OF FILE"
   End
   Begin VB.Menu MNU_CREATE_FOLDER_DATE_MONTH 
      Caption         =   "CREATE FOLDER DATE MONTH"
   End
   Begin VB.Menu MNU_INCREMENT_DATE 
      Caption         =   "INCREMENT DATE -- FALSE"
   End
   Begin VB.Menu MNU_CLIPBOARD_DATE 
      Caption         =   "CLIPBOARD DATE _ YYYY MM DD HH MM SS"
   End
   Begin VB.Menu MNU_MENU_ITEM_COUNT 
      Caption         =   "MENU COUNT"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --------------------------------------------------------------------
' SESSION START COUNT 001
' SIXER 006 PROJECT TODAY
' --------------------------------------------------------------------
' MAKE DO GIVER BLUE SELECT BUTTON FOR LAST CHOOSE
' AND ALSO WRITE TO FILE
' MODIFY DATE AND ALSO RENAME A LITTLE BIT
' --------------------------------------------------------------------
' Thu 28-May-2020 15:22:11
' Thu 28-May-2020 15:48:00 -- 25 MINUTE
' --------------------------------------------------------------------
' ALL DAY CODE FROM 8 AM
' AND SHOPING MARKET IN MORNING WHEN UP
' FINE FIDDLY NUMBER TO SPEND
' 3 MONTH SINCE SHOPPING MARKET OPEN
' CLOSE CALL BILL END OF MONTH
' STILL 2 WORK LETF TO DO -- SERACHER THING
' --------------------------------------------------------------------


Dim MSGBOX2
Dim MSGBOX4

Dim I_3

' Public MENU_INCREMENT_DATE_VAR

Dim WORK_INDEX
Dim FILE_NAME_MP4
Dim FILE_NAME_MP4_2
Dim MAQ
Dim START_MAQ
Dim NAME_PART_02

Dim VARCENTER

Dim FORM_ME As New Form1

Dim I_1 ' QUICK MODE RESULT
Dim I_2 ' QUICK MODE RESULT

Dim M_1()
Dim M_2()
Dim M_3()
Dim FR1

Dim FIRST_RUN
Dim a
Dim W$
Dim WORK
Dim FSO

Dim FULL_PATH_AND_FILENAME

'--------------------------------------

'Excellent Code Set the File Time To ANything you want touch date Pro Keep you file time after rotate JPg's




'Option Explicit

'THIS TOGGLES A SMALL PASSWORD FEATURE THAT I EMBEDDED
Const MPassword = False

'THIS IS THE CODE THAT MUST BE ENTERED (FINISHED BY PRESSING HASH)
'WHEN THE PASSWORD TOGGLE IS TRUE
Const MSequence = 22515

Dim MEntered As Variant

 Private Type FILETIME
     dwLowDate  As Long
     dwHighDate As Long
 End Type
 
 Private Type SYSTEMTIME
     wYear      As Integer
     wMonth     As Integer
     wDayOfWeek As Integer
     wDay       As Integer
     wHour      As Integer
     wMinute    As Integer
     wSecond    As Integer
     wMillisecs As Integer
 End Type
 
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_WRITE = &H40000000
  
Private Declare Function CreateFile Lib "kernel32" Alias _
   "CreateFileA" (ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   ByVal lpSecurityAttributes As Long, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) _
   As Long

Private Declare Function LocalFileTimeToFileTime Lib _
     "kernel32" (lpLocalFileTime As FILETIME, _
      lpFileTime As FILETIME) As Long

Private Declare Function SetFileTime Lib "kernel32" _
   (ByVal hFile As Long, ByVal MullP As Long, _
    ByVal NullP2 As Long, lpLastWriteTime _
    As FILETIME) As Long

Private Declare Function SystemTimeToFileTime Lib _
    "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime _
    As FILETIME) As Long
    
Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type MENUBARINFO
  cbSize As Long
  rcBar As RECT
  hMenu As Long
  hWndMenu As Long
  fBarFocused As Boolean
  fFocused As Boolean
End Type
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private MenuInfo As MENUBARINFO
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Const OBJID_SYSMENU As Long = &HFFFFFFFF
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, _
ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Boolean
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long


Private Sub GetMenuInfo(hMenu As Long, Spaces As Integer, SubText, txt As String)
Dim num As Integer
Dim i As Integer
Dim length As Long
Dim sub_hmenu As Long
Dim sub_name As String
    
    num = GetMenuItemCount(hMenu)
    For i = 0 To num - 1
        ' Save this menu's info.
        sub_hmenu = GetSubMenu(hMenu, i)
        sub_name = Space$(256)
        'Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
        length = GetMenuString(hMenu, i, sub_name, Len(sub_name), MF_BYPOSITION)
        sub_name = Left$(sub_name, length)

        ' txt = txt & SubText & Space$(Spaces) & sub_name & vbCrLf
        txt = txt & SubText & Space$(Spaces) & sub_name & vbCrLf
        
        ' Get its child menu's names.
        GetMenuInfo sub_hmenu, Spaces + 4, "Sub Menu ----", txt
    Next i
End Sub

Private Function MENU_HEIGHT()
    
    ' ------------------------------------------------------
    ' frmListMenu.MENU_HEIGHT
    ' ------------------------------------------------------
    ' HOW TO USER
    ' WHEN SET FORM HIEIGHT
    ' EXAMPLE HERE
    ' ------------------------------------------------------
    ' (Menu_Height * Screen.TwipsPerPixelY)
    ' ------------------------------------------------------
    
    MenuInfo.cbSize = Len(MenuInfo)
    If GetMenuBarInfo(Me.hWnd, OBJID_MENU, 0, MenuInfo) Then
        With MenuInfo.rcBar
'            Debug.Print "Left: " & CStr(.Left)
'            Debug.Print "Right: " & CStr(.Right)
'            Debug.Print "Top: " & CStr(.Top)
'            Debug.Print "Bottom: " & CStr(.Bottom)
            MENU_HEIGHT = CStr(.Bottom) - CStr(.Top)
        End With
    End If
End Function




Private Sub Form_Resize()

Dim r

ReDim M_1(LABEL_SET.Count)
ReDim M_2(LABEL_SET.Count)
ReDim M_3(LABEL_SET.Count)
i = 0

i = i + 1: M_1(i) = "GO"
i = i + 1: M_1(i) = "FOLDER LABEL"
i = i + 1: M_1(i) = "FILE LABEL"
'i = i + 1: M_1(i) = "PERFORM ON ALL FILES IN FOLDER OR FILE"
i = i + 1: M_1(i) = "NOW_DATE"
i = i + 1: M_1(i) = "MODIFY DATE TO CREATED DATE - NOT WORKING"
i = i + 1: M_1(i) = "----"
i = i + 1: M_1(i) = "BATCH - CREATED DATE TO MODIFY DATE"
i = i + 1: M_1(i) = "FILE  - CREATED DATE TO MODIFY DATE"
i = i + 1: M_1(i) = "----"
i = i + 1: M_1(i) = "DATE CONVERTOR__ MMM D, YYYY H MM AM __ TO YYYY-MM-DD--HH-MM-DD FOR SCREENCASTIFY"
i = i + 0: M_3(i) = "DATE_CONVERTOR___MMM_D__YYYY_H_MM_AM____TO_YYYY_MM_DD__HH_MM_DD_FOR_SCREENCASTIFY"

i = i + 1: M_1(i) = "DATE CONVERTOR__ MMM D, YYYY H MM AM __ TO YYYY-MM-DD--HH-MM-DD FOR SCREENCASTIFY SINGLE"
i = i + 0: M_3(i) = "DATE_CONVERTOR___MMM_D__YYYY_H_MM_AM____TO_YYYY_MM_DD__HH_MM_DD_FOR_SCREENCASTIFY_SINGLE"

i = i + 1: M_1(i) = "OPEN_ALL_URL_PATH_OPERATION_001_DEPTH"

i = i + 1: M_1(i) = "----"
i = i + 1: M_1(i) = "CAPITAL_ALL____PATH_AND_FILE_____OPERATION____TOP_DEPTH"
i = i + 1: M_1(i) = "CAPITAL_ALL____PATH_OPERATION____TOP_DEPTH"
i = i + 1: M_1(i) = "CAPITAL_ALL____PATH_OPERATION__001_DEPTH"
i = i + 1: M_1(i) = "CAPITAL_ALL____FILE_OPERATION____TOP_DEPTH"
i = i + 1: M_1(i) = "CAPITAL_ALL____FILE_OPERATION__001_DEPTH"
i = i + 1: M_1(i) = "----"


'i = i + 1: M_1(i) = "----"
i = i + 1: M_1(i) = "ONE_FOLDER AS OTHER SAME ___ HARDCODER __ BATCH"
i = i + 0: M_3(i) = "ONE_FOLDER_AS_OTHER_SAME_____HARDCODER____BATCH"
'i = i + 1: M_1(i) = "----"
i = i + 1: M_1(i) = "MAKE_FOLDER YYYY-MM-DD OF FILE AND MOVE THERE ___ BATCH IT"
i = i + 0: M_3(i) = "MAKE_FOLDER_YYYY_MM_DD_OF_FILE_AND_MOVE_THERE_____BATCH_IT"
'i = i + 1: M_1(i) = "----"

i = i + 1: M_1(i) = "RENAME -- YYYY_MM_DD MMM_DDD HH_MM_SS__MA_.MP4 -- BATCH"
i = i + 0: M_3(i) = "RENAME____YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA__MP4____BATCH"
i = i + 1: M_1(i) = "RENAME -- YYYY_MM_DD MMM_DDD HH_MM_SS -- AVI -- MP4 -- ANY -- BATCH"
i = i + 0: M_3(i) = "RENAME____YYYY_MM_DD_MMM_DDD_HH_MM_SS__AVI_MP4_ANY____BATCH"

i = i + 1: M_1(i) = "RENAME -- YYYY_MM_DD HH_MM_SS -- ANY -- BATCH"
i = i + 0: M_3(i) = "RENAME____YYYY_MM_DD_HH_MM_SS__ANY____BATCH"

' 2020-06-20 09-48-48 - SONY DSC-HX60V - DSC03279
i = i + 1: M_1(i) = "RENAME -- DSC0----.JPG -TO- YYYY-MM-DD HH-MM-SS - SONY - DSC-HX60V - DSC0----.JPG -- BATCH"
i = i + 0: M_3(i) = "RENAME____DSC0_____JPG__TO__YYYY_MM_DD_HH_MM_SS___SONY___DSC_HX60V___DSC0_____JPG____BATCH"

i = i + 1: M_1(i) = "----"

i = i + 1: M_1(i) = "SET_DATE_OF_FILENAME -- CH00_YYYY_MM_DD HH_MM_SS.MP4 -- HIKVISION -- SINGLE"
i = i + 0: M_3(i) = "SET_DATE_OF_FILENAME_CH00_YYYY_MM_DD_HH_MM_SS_MP4_HIKVISION_SINGLE"
i = i + 1: M_1(i) = "SET_DATE_OF_FILENAME -- CH00_YYYY_MM_DD HH_MM_SS.MP4 -- HIKVISION -- BATCH"
i = i + 0: M_3(i) = "SET_DATE_OF_FILENAME_CH00_YYYY_MM_DD_HH_MM_SS_MP4_HIKVISION_BATCH"
i = i + 1: M_1(i) = "SET_DATE_OF_FILENAME -- CH00_YYYYMMDD__HHMMSS.MP4 .AVI -- HIKVISION -- BATCH"
i = i + 0: M_3(i) = "SET_DATE_OF_FILENAME_CH00_YYYYMMDD__HHMMSS_MP4_HIKVISION_BATCH"

i = i + 1: M_1(i) = "SET_DATE_OF_FILENAME -- YYYY_MM_DD .MP4 .WMV -- WITH_ONE_FOLDER"
i = i + 0: M_3(i) = "SET_DATE_OF_FILENAME_YYYY_MM_DD__WITH_ONE_FOLDER"
i = i + 1: M_1(i) = "SET_DATE_OF_FILENAME -- YYYY_MM_DD MMM_DDD HH_MM_SS__MA_.MP4 -- SINGLE"
i = i + 0: M_3(i) = "SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA_SINGLE"
i = i + 1: M_1(i) = "SET_DATE_OF_FILENAME -- YYYY_MM_DD MMM_DDD HH_MM_SS__MA_.MP4 -- BATCH"
i = i + 0: M_3(i) = "SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA"
i = i + 1: M_1(i) = "SET_DATE_OF_FILENAME -- YYYY MM DD HH-MM-SS - DDD NOKIA__.MP4 -- BATCH"
i = i + 0: M_3(i) = "SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS_DDD_NOKIA_AH"

i = i + 1: M_1(i) = "SET_DATE_OF_FILENAME -- YYYY MM DD HH-MM-SS_.JPG -- CAMERA -- SINGLE"
i = i + 0: M_3(i) = "SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS__JPG_SINGLE"
i = i + 1: M_1(i) = "SET_DATE_OF_FILENAME -- YYYY MM DD HH-MM-SS_.JPG -- BATCH"
i = i + 0: M_3(i) = "SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS__JPG_BATCH"

i = i + 1: M_1(i) = "----"
i = i + 1: M_1(i) = "SET_ONE_DATE_HARDCODER"
i = i + 1: M_1(i) = "----"
i = i + 1: M_1(i) = "SET_MOST_RECENT_DATE_TO_OTHER_IN_FOLDER"
i = i + 1: M_1(i) = "SET_OLDER_DATE_TO_OTHER_IN_FOLDER"
i = i + 1: M_1(i) = "----"
i = i + 1: M_1(i) = "SET_ALL_DATE_FOLDER_TO_THE_TEXTFILE_HOLD_DATE_WITHIN_AH_--_\#_TEXT_DATER*.TXT"
i = i + 0: M_3(i) = "SET_ALL_DATE_FOLDER_TO_THE_TEXTFILE_HOLD_DATE_WITHIN_AH"

For r = 0 To i
    M_2(r) = Replace(M_1(r), "_", " ")
Next

For r = 0 To LABEL_SET.Count - 1
    If M_2(r) <> "FOLDER LABEL" Then
        If M_2(r) <> "FILE LABEL" Then
            If LABEL_SET(r).Caption <> M_2(r) Then
                LABEL_SET(r).Caption = M_2(r)
            End If
        End If
    End If
    If LABEL_SET(r).Caption = "" Then
        LABEL_SET(r).Visible = False
    End If
Next

On Error Resume Next
Me.Width = Screen.Width - 1000 ' 24000
For r = 0 To LABEL_SET.Count - 1
    LABEL_SET(r).FontSize = 7    ' MAIN FONT
    LABEL_SET(r).fontname = "ARIAL"
Next

MIN_HL = 4000
For r = 0 To LABEL_SET.Count - 1
    LABEL_SET(r).Refresh
    LABEL_SET(r).AutoSize = True
    LABEL_SET(r).Refresh
    If LABEL_SET(r).Caption <> "" Then
    If LABEL_SET(r).Height < MIN_HL Then
        MIN_HL = LABEL_SET(r).Height
    End If
    End If
    LABEL_SET(r).AutoSize = False
Next
LABEL_SET(4).FontSize = 8
LABEL_SET(2).FontSize = 8
LABEL_SET(2).FontSize = LABEL_SET(4).FontSize ' MABE ADJUST LESS
LABEL_SET(3).FontSize = LABEL_SET(2).FontSize

HL = MIN_HL
LENGHT_LABEL = 380
LENGHT_LABEL_AND_WIDTH = Me.Width - LENGHT_LABEL + 20
STEP_H = 20
LENGHT_LABEL_AND_WIDTH = LENGHT_LABEL_AND_WIDTH - 20
For r = 2 To LABEL_SET.Count - 1
    If LABEL_SET(r).Caption = "" Then
        LABEL_SET(r).Visible = False
    End If
    
    If LABEL_SET(r).Visible = True Then
        LABEL_SET(r).Left = 100
        LABEL_SET(r).Width = LENGHT_LABEL_AND_WIDTH
        LABEL_SET(r).Height = HL
        LABEL_SET(r).Top = STEP_H
        STEP_H = STEP_H + 15 + HL
    End If
Next

r = 1
LABEL_SET(r).Left = 100
LABEL_SET(r).Width = LENGHT_LABEL_AND_WIDTH
LABEL_SET(r).Height = HL
LABEL_SET(r).Top = STEP_H
STEP_H = STEP_H + 15 + HL

On Error Resume Next

Me.Height = STEP_H + 500 + 100 + (MENU_HEIGHT * Screen.TwipsPerPixelY) + 500 ' +500 WINDOWS 11 ' + MENUBAR CODE REQUIRE --- REQUIRE HIGH NUMBER IF WINDOWS 10 NOT GRAPHIC DPI CLUE
Me.Top = 0

If Me.WindowState = 0 Then
    If VARCENTER = True Then Exit Sub
    Me.Top = Screen.Height / 2 - Me.Height / 2 '+400
    Me.Left = Screen.Width / 2 - Me.Width / 2 '+400
    VARCENTER = True
    On Error GoTo 0
End If

On Error Resume Next
    FR1 = FreeFile
    FILENAME_MOD_AR = "\# DATA\#_SEND_TO MODIFY_DATE_" + "_INDEX_INFO_STORE__#NFS_EX.TXT"
    ' UCase(App.EXEName) ' "#SEND TO TOUCHDATE"
    Open App.Path + FILENAME_MOD_AR For Input As #FR1
        Line Input #FR1, INDEX_FIND
    Close #FR1
    LABEL_SET(INDEX_FIND).BackColor = LABEL_COLOR_SET_EFORE_INDEX_FINGER.BackColor
    
On Error GoTo 0


End Sub


Private Sub LABEL_SET_Click(index As Integer)

' FILENAME
' LABEL_SET(3).Caption

DISPLAY_NAMER = M_1(index)
If M_3(index) <> "" Then
    DISPLAY_NAMER = M_3(index)
End If

If DISPLAY_NAMER <> "GO" Then
    WORK = DISPLAY_NAMER
    WORK_INDEX = index
End If

If WORK <> "" Then
    On Error Resume Next
    If Dir(App.Path + "\# DATA", vbDirectory) = "" Then MkDir App.Path + "\DATA"
    FILENAME_MOD_AR = "\# DATA\#_SEND_TO MODIFY_DATE_" + "_INDEX_INFO_STORE__#NFS_EX.TXT"
    
    INDEX_FIND = WORK_INDEX
    FR1 = FreeFile
    Open App.Path + FILENAME_MOD_AR For Output As #FR1
        Print #FR1, INDEX_FIND
    Close #FR1
    For R_COUNTER = 0 To LABEL_SET.Count - 1
        If R_COUNTER > 2 And M_2(R_COUNTER) <> "GO" Then
            LABEL_SET(R_COUNTER).BackColor = vbButtonFace
        End If
    Next
    LABEL_SET(INDEX_FIND).BackColor = LABEL_COLOR_SET_EFORE_INDEX_FINGER.BackColor
    On Error GoTo 0
End If



Select Case DISPLAY_NAMER

Case 2
    ' FOLDER BUTTON
    ' LABEL_SET(2)
Case 3
    ' FILE BUTTON
    ' LABEL_SET(3)

Case "PERFORM ON ALL FILES IN FOLDER OR FILE"
    ' Call Label3_Click
    
Case "NOW_DATE"
    ' WORK = "MOD_TO_NOW_DATE"
    Call ROUTINE_MOD_TO_NOW_DATE

Case "MODIFY DATE TO CREATED DATE - NOT WORKING"
    Call Label2_Click
    
Case "BATCH - CREATED DATE TO MODIFY DATE"
    Call Label9_Click
    
Case "FILE  - CREATED DATE TO MODIFY DATE"
    Call Label11_Click
    
Case "SET_ONE_DATE_HARDCODER"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor
    
Case "ONE_FOLDER_AS_OTHER_SAME_____HARDCODER____BATCH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "MAKE_FOLDER_YYYY_MM_DD_OF_FILE_AND_MOVE_THERE_____BATCH_IT"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "DATE_CONVERTOR___MMM_D__YYYY_H_MM_AM____TO_YYYY_MM_DD__HH_MM_DD_FOR_SCREENCASTIFY"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor
Case "DATE_CONVERTOR___MMM_D__YYYY_H_MM_AM____TO_YYYY_MM_DD__HH_MM_DD_FOR_SCREENCASTIFY_SINGLE"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "OPEN_ALL_URL_PATH_OPERATION_001_DEPTH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "CAPITAL_ALL____PATH_AND_FILE_____OPERATION____TOP_DEPTH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor
Case "CAPITAL_ALL____PATH_OPERATION____TOP_DEPTH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor
Case "CAPITAL_ALL____PATH_OPERATION__001_DEPTH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor
Case "CAPITAL_ALL____FILE_OPERATION____TOP_DEPTH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor
Case "CAPITAL_ALL____FILE_OPERATION__001_DEPTH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "RENAME____YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA__MP4____BATCH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor
Case "RENAME____YYYY_MM_DD_MMM_DDD_HH_MM_SS__AVI_MP4_ANY____BATCH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "RENAME____YYYY_MM_DD_HH_MM_SS__ANY____BATCH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor


Case "RENAME____DSC0_____JPG__TO__YYYY_MM_DD_HH_MM_SS___SONY___DSC_HX60V___DSC0_____JPG____BATCH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "SET_DATE_OF_FILENAME_CH00_YYYY_MM_DD_HH_MM_SS_MP4_HIKVISION_SINGLE"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor
Case "SET_DATE_OF_FILENAME_CH00_YYYY_MM_DD_HH_MM_SS_MP4_HIKVISION_BATCH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor
Case "SET_DATE_OF_FILENAME_CH00_YYYYMMDD__HHMMSS_MP4_HIKVISION_BATCH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "SET_DATE_OF_FILENAME_YYYY_MM_DD__WITH_ONE_FOLDER"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA_SINGLE"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS__JPG_SINGLE"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor


Case "SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS__JPG_BATCH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS_DDD_NOKIA_AH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor
  
Case "SET_MOST_RECENT_DATE_TO_OTHER_IN_FOLDER"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor
Case "SET_OLDER_DATE_TO_OTHER_IN_FOLDER"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor

Case "SET_ALL_DATE_FOLDER_TO_THE_TEXTFILE_HOLD_DATE_WITHIN_AH"
LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor
    
Case "GO"
    If LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor Then
        LABEL_SET(1).BackColor = Label_COLOR_GREEN.BackColor + 30
        LABEL_SET(1).Refresh
        DoEvents
        Call Label_GO_AH_Click
    End If
End Select

End Sub


Private Sub Label_GO_AH_Click()

If WORK = "MOD_TO_NOW_DATE" Then
    ' -------------------------------
    a = LABEL_SET(3).Caption           ' GET FILE
    DateSet = Now
    TT = SetFileDateTime(a, DateSet)
    MsgBox "Done " + vbCrLf + vbCrLf + a + vbCrLf + vbCrLf + DATEVAR, vbMsgBoxSetForeground
    End

    ' Call CODE_RUN
    Exit Sub
End If

If WORK = "MOD_TO_CREATED_DATE" Then
    Call BATCH_MODIFIED_TO_CREATED_TIME
    Exit Sub
End If

If WORK = "CREATED_TO_MOD_DATE" Then
    Call BATCH_CREATED_TO_MODIFIED_TIME
    Exit Sub
End If

If WORK = "FILE_CREATED_TO_MODIFIED_TIME" = True Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_MOST_RECENT_DATE_TO_OTHER_IN_FOLDER" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_OLDER_DATE_TO_OTHER_IN_FOLDER" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "DATE_CONVERTOR___MMM_D__YYYY_H_MM_AM____TO_YYYY_MM_DD__HH_MM_DD_FOR_SCREENCASTIFY" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "DATE_CONVERTOR___MMM_D__YYYY_H_MM_AM____TO_YYYY_MM_DD__HH_MM_DD_FOR_SCREENCASTIFY_SINGLE" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "OPEN_ALL_URL_PATH_OPERATION_001_DEPTH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "CAPITAL_ALL____PATH_AND_FILE_____OPERATION____TOP_DEPTH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If
If WORK = "CAPITAL_ALL____PATH_OPERATION____TOP_DEPTH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If
If WORK = "CAPITAL_ALL____PATH_OPERATION__001_DEPTH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If
If WORK = "CAPITAL_ALL____FILE_OPERATION____TOP_DEPTH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If
If WORK = "CAPITAL_ALL____FILE_OPERATION__001_DEPTH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "MAKE_FOLDER_YYYY_MM_DD_OF_FILE_AND_MOVE_THERE_____BATCH_IT" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "ONE_FOLDER_AS_OTHER_SAME_____HARDCODER____BATCH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_ALL_DATE_FOLDER_TO_THE_TEXTFILE_HOLD_DATE_WITHIN_AH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS_DDD_NOKIA_AH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS__JPG_SINGLE" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS__JPG_BATCH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "RENAME____YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA__MP4____BATCH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "RENAME____YYYY_MM_DD_MMM_DDD_HH_MM_SS__AVI_MP4_ANY____BATCH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "RENAME____YYYY_MM_DD_HH_MM_SS__ANY____BATCH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If


If WORK = "RENAME____DSC0_____JPG__TO__YYYY_MM_DD_HH_MM_SS___SONY___DSC_HX60V___DSC0_____JPG____BATCH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If


If WORK = "SET_DATE_OF_FILENAME_CH00_YYYY_MM_DD_HH_MM_SS_MP4_HIKVISION_SINGLE" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If
If WORK = "SET_DATE_OF_FILENAME_CH00_YYYY_MM_DD_HH_MM_SS_MP4_HIKVISION_BATCH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_DATE_OF_FILENAME_CH00_YYYYMMDD__HHMMSS_MP4_HIKVISION_BATCH" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If


If WORK = "SET_DATE_OF_FILENAME_YYYY_MM_DD__WITH_ONE_FOLDER" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If
If WORK = "SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA_SINGLE" Then
    CallByName FORM_ME, WORK, VbMethod
    Exit Sub
End If

If WORK = "SET_ONE_DATE_HARDCODER" = True Then
    
    Call SET_BATCH_DATE_CAMERA_VIDEO_FILENAME_TO_DATE_FILE
    End
    
'   SIMPLE
'   -------------------------------
    ' a = "D:\UTILS\2011 GALAXY SAMSUNG GT-P1000 - Copy\2012 07 GALAXY SAMSUNG GT-P1000_ VIDEO.MP4"
    a = LABEL_SET(3).Caption
    DATEVAR = "2012/07/01 18:00:00"
    DATEVAR = "2017/04/04 23:35:44"
    DateSet = DateValue(DATEVAR) + TimeValue(DATEVAR)
    TT = SetFileDateTime(a, DateSet)
    MsgBox "Done " + vbCrLf + vbCrLf + a + vbCrLf + vbCrLf + DATEVAR, vbMsgBoxSetForeground
    End

End If


If WORK = "SET_ONE_DATE" = True Then
        
    a = "D:\DSC\# Docus Proofs Texts\# BHT\NOTICE BOARD\2009-09-22 21-23-03 - Sony Ericsson K800i - DSC03044.JPG"
    a = "D:\DSC\# Docus Proofs Texts\# BHT\SMS PROOFS\2009-09-22 23-09-55 - Sony Ericsson K800i - DSC03047.JPG"
    a = "D:\DSC\# Docus Proofs Texts\# BHT\SMS PROOFS\2009-09-22 23-10-12 - Sony Ericsson K800i - DSC03048.JPG"
    a = "D:\VI_ DSC ME\2010+NOKIA\2015 NOKIA E72 _ 008 AUG _ MP4 _ x001 _ Home Front Room\2015_08_07 AUG_FRI 15_13_38__MAQ06717 - MY HOME.MP4"
    
    
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_07_24 Jul_Wed 22_48_08__MAQ07083_A Piezoelectric Speaker (AKA A Piezo Bender)_ELECTRONIC TECHNO RECORDER.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_02_07 Feb_Thu 14_42_48__MAQ03670_SEAGULLS AT BEACH SWARM.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_02_27 Feb_Wed 08_04_38__MAQ04106 VIDEO OF DEVICE THAT DOESN'T WORK GOT FROM EBAY ALARM AT POWER FAILURE ALERT.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_02_27 Feb_Wed 08_04_38__MAQ04106 VIDEO OF DEVICE THAT DOESN'T WORK GOT FROM EBAY ALARM AT POWER FAILURE ALERT.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_08_01 Aug_Thu 20_19_40__MAQ07292_VIEW TOP SOUTHWICK HILL SOUTH DOWNS.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_07_28 Jul_Sun 17_00_39__MAQ07182_UP THE DOWNS SUOTHWICK HILL.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_07_12 Jul_Fri 16_04_05__MAQ06893_UP THE SOUTH DOWN_.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_07_06 Jul_Sat 09_58_34__MAQ06634_UP THE SOUTH DOWN_ VIEW ON CHANCTONBURY RING _ LOT OF ANIMAL_ CATTLE & OTHER.txt"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_06_27 Jun_Thu 16_11_26__MAQ06357_A Walk Up The Downs _ Was Giver Show Sound Of Pylon Wire Thrash By Heavy Wind And Sunny A Type Of Undescribable Audio A View of Chantonbury Ring.MP4"
    a = "D:\DSC\2015+Sony\2019 CyberShot HX60V\MP4\2019_02_27 Feb_Wed 08_04_38__MAQ04106 VIDEO OF DEVICE THAT DOESN'T WORK GOT FROM EBAY ALARM AT POWER FAILURE ALERT.MP4"
    XX = InStr(a, "\2019_")
    DATEVAR = Replace(Mid(a, XX + 1, 10) + " " + Mid(a, XX + 20, 8), "_", "-")
    
    
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_02_27 Feb_Tue 15_27_03__MAQ08676 _ Down at the Docks, Shoreham Port The Digger Weighs How Much Hardcore In The Spade_2.mp4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_05_28 May_Mon 15_25_36__MAQ00290 _ STARLING NEW BORN.MP4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_06_17 Jun_Sun 13_21_11__MAQ00586.mp4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_07_13 Jul_Fri 16_05_40__MAQ01488 _ Rotated_Garden_Water.mp4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_07_19 Jul_Thu 10_58_21__MAQ01563 _ Logitech 533 Speaker Standby Mode Removed Circuit Delay Control Gear.mp4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_08_20 Aug_Mon 19_02_51__MAQ02416 _ Helicopter Along the Seafront.mp4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_08_20 Aug_Mon 16_39_01__MAQ02389 _ SQUIRREL IN THE GARDEN.mp4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_10_27 Oct_Sat 10_18_46__MAQ01732_SEAGULL AND THE BIRD FEAST FAT TRAY WITH CHERRIES ON.MP4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_10_27 Oct_Sat 10_18_46__MAQ01732_SEAGULL AND THE BIRD FEAST FAT TRAY WITH CHERRIES ON.MP4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_10_27 Oct_Sat 10_18_46__MAQ01732_SEAGULL AND THE BIRD FEAST FAT TRAY WITH CHERRIES ON.MP4"
    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_10_31 Oct_Wed 11_55_50__MAQ01836 (2).mp4"
    
    XX = InStr(a, "\2018_")
    DATEVAR = Replace(Mid(a, XX + 1, 10) + " " + Mid(a, XX + 20, 8), "_", "-")
    
    'SET THE FILE ABOVE
    'SET DATE BELOW

'    Set F = FSO.GetFile(a)
'    DT1 = F.datelastmodified
'    Set F = Nothing
    
    'DateSet = DT1 - TimeSerial(1, 0, 0)

'    DATEVAR = "2009-09-22 21-23-03"
'    DATEVAR = "2009-09-22 23-09-55"
'    DATEVAR = "2009-09-22 23-10-12"
'    DATEVAR = "2015-08-07 15-13-38"
    'DATEVAR = ""
    'DATEVAR = ""
    
    
    For r = 1 To 10
    If Mid(DATEVAR, r, 1) = "-" Then Mid(DATEVAR, r, 1) = "/"
    Next
    For r = 10 To Len(DATEVAR)
    If Mid(DATEVAR, r, 1) = "-" Then Mid(DATEVAR, r, 1) = ":"
    Next
    DateSet = DateValue(DATEVAR) + TimeValue(DATEVAR)
    
'    Ds1 = DateValue(Mid(xB11, 9, 2) + "-" + Mid(xB11, 6, 2) + "-" + Mid(xB11, 1, 4))
'    Ds2 = TimeValue(Replace(Mid(xB11, 12, 8), "-", ":"))
'    DateSet = DateValue(Ds1) + TimeValue(Ds1)
'    DateSet = Ds1 + Ds2
'    DateSet = GetExif(A11 + B11)
            
    TT = SetFileDateTime(a, DateSet)
    XC = XC + 1
    
    MsgBox "Done = " + str(XC)
    End
    
End If
End Sub


Sub RENAME_OPERATION_HARDCODE_WHEN_ERROR_MANY()
    Exit Sub
    Set FSO = CreateObject("Scripting.FileSystemObject")

'    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
'        File1.Path = LABEL_SET(2).Caption
'    End If
    
    If IsIDE Then
        File1.Path = "\\4-asus-gl522vw\4_asus_gl522vw_10_1_samsung_1tb\DSC\2015+SONY_MP4\2020 CyberShot HX60V __ MP4"
        File1.Path = "\\7-asus-gl522vw\7_asus_gl522vw_80_3_samsung_4tb_d\VI_ DSC ME\2010+SONY\2020 CyberShot HX60V_#"
        File1.Path = "K:\M\07 MIDI MOD\MIDI\03 03671 FILM MIDI\02 1633 FILM TRACKS NAME POLY"
    End If
    
    If Len(File1.Path) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
'    If ScanPath.ListView1.ListItems.Count > 500 Then
'        MsgBox "LOOK LIKE GOT TOO MANY FILE TO DO OVER 500" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
'        End
'    End If
    
    If File1.ListCount > 0 Then
        MsgBox "FILE COUNTER" + vbCrLf + vbCrLf + Trim(str(File1.ListCount)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + File1.Path
    Else
        MsgBox "NONE COUNT OF FILE HERE " + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + File1.Path
        End
    End If


    ' ----------------------------------------------------------
    ' ROUTINE NUMBER 0002
    ' Fri 25-Sep-2020 01:38:00
    ' ----------------------------------------------------------
    Dim TT
    Dim RR
    For RR = 0 To File1.ListCount - 1
        ' A ONE ONE
        A11 = File1.Path + "\" + File1.List(RR)
        D11 = File1.Path + "\"
        ' ----------------------------------------------------------
        If InStr(A11, "_gsdata_") = 0 Then
            On Error Resume Next
            Err.Clear
            
            TTMANIP_01 = Mid(A11, InStrRev(A11, "\") + 1)                  ' FILENAME
            TTMANIP_02 = Mid(TTMANIP_01, 1, InStrRev(TTMANIP_01, ".") - 1) ' FILENAME AND NOT EXTENSION
            TTMANIP_04 = Mid(TTMANIP_02, 5) ' NUMBER PART EXAMPLE POLY 001.MID
            TTMANIP_04 = "POLY" + Format(Val(TTMANIP_04), "0000") + ".MID"
            If Mid(UCase(TTMANIP_01), 1, 4) = "POLY" Then
                Err.Clear
                Name A11 As D11 + TTMANIP_04
                
                XC = XC + 1
                MM_1 = MM_1 + UCase(A11)
                If Err.Number > 0 Then MM_1 = MM_1 + Err.Description + vbCrLf: ERROR_SET = True
                MM_1 = MM_1 + vbCrLf
            End If
        End If
    Next

    
    ' ----------------------------------------------------------
    ' ROUTINE NUMBER 0001
    ' WEEK BEFORE
    ' ----------------------------------------------------------
    ' Dim TT
    ' Dim RR
    If 1 = 2 Then
    For RR = 0 To File1.ListCount - 1
        ' A ONE ONE
        A11 = File1.Path + "\" + File1.List(RR)
        ' ----------------------------------------------------------
        If InStr(A11, "_gsdata_") = 0 Then
            On Error Resume Next
            Err.Clear
            
            TTMANIP_01 = Mid(A11, InStrRev(A11, "\"))
            TTMANIP_04 = TTMANIP_01
            A22 = TTMANIP_01
            TTMANIP_02 = Mid(A22, 13, 3)
            TTMANIP_03 = Mid(A22, 17, 3)
            GET_REQUIRE_SWAP = "YER"
            
            MONTH_AR = "JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC"
            MONTH_V = Split(MONTH_AR, ",")
            For R4 = 1 To 12
                If UCase(TTMANIP_02) = MONTH_V(R4 - 1) Then
                    GET_REQUIRE_SWAP = "NOT"
                End If
            Next
            If GET_REQUIRE_SWAP = "YER" Then
                Mid(TTMANIP_01, 13, 3) = TTMANIP_03
                Mid(TTMANIP_01, 17, 3) = TTMANIP_02
                A22 = Replace(A11, TTMANIP_04, TTMANIP_01)
                Name A11 As A22
                ' MsgBox A11
                ' Stop
            End If
            ' Stop
            ' Name A11 As UCase(A11)
            XC = XC + 1
            MM_1 = MM_1 + UCase(A11)
            If Err.Number > 0 Then MM_1 = MM_1 + Err.Description + vbCrLf: ERROR_SET = True
            MM_1 = MM_1 + vbCrLf
        End If
    Next
    End If
    
    
    
    If ERROR_SET = True Then
        MsgBox "ERROR ON " + vbCrLf + vbCrLf + MM_1
    Else
        ' Call SET_QUICK_MODE_RESULT
        ' If InStr(MNU_QUICK_MODE.Caption, I_2) Then
            MsgBox "DONE = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
        ' End If
        End If
    End

End

End Sub




Public Function SetFileDateTime(ByVal Filename As String, _
  ByVal TheDate As String) As Boolean
'************************************************
'PURPOSE:    Set File Date (and optionally time)
'            for a given file)
         
'PARAMETERS: TheDate -- Date to Set File's Modified Date/Time
'            FileName -- The File Name

'Returns:    True if successful, false otherwise
'************************************************
On Error Resume Next
If Dir(Filename) = "" Then Exit Function
If Err.Number > 0 Then
    MSGBOX4 = MSGBOX4 + 1
    MSGBOX2 = Trim(str(MSGBOX4)) + " ERROR(S) OCCURED FIRST WITH " + vbCrLf + vbCrLf + Filename + vbCrLf
    On Error GoTo 0
    Exit Function
End If

On Error GoTo 0

If Not IsDate(TheDate) Then Exit Function

Dim lFileHnd As Long
Dim lRet As Long

Dim typFileTime As FILETIME
Dim typLocalTime As FILETIME
Dim typSystemTime As SYSTEMTIME

With typSystemTime
    .wYear = Year(TheDate)
    .wMonth = Month(TheDate)
    .wDay = Day(TheDate)
    .wDayOfWeek = Weekday(TheDate) - 1
    .wHour = Hour(TheDate)
    .wMinute = Minute(TheDate)
    .wSecond = Second(TheDate)
End With

lRet = SystemTimeToFileTime(typSystemTime, typLocalTime)
lRet = LocalFileTimeToFileTime(typLocalTime, typFileTime)

lFileHnd = CreateFile(Filename, GENERIC_WRITE, _
    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, _
    OPEN_EXISTING, 0, 0)
    
lRet = SetFileTime(lFileHnd, ByVal 0&, ByVal 0&, _
         typFileTime)

CloseHandle lFileHnd
SetFileDateTime = lRet > 0

End Function


Private Sub Form_Activate()

'LABEL_SET(2).Caption = "F:\DSC\2015+SONY_MP4\2020 CyberShot HX60V __ MP4"
'Call MNU_FILEDATE_WHOLE_FOLDER_Click

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then End

' i = MsgBox("", vbMsgBoxSetForeground)

End Sub

Private Sub Form_Load()
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'If IsIDE = True Then GoTo Start2
    
    Call RENAME_OPERATION_HARDCODE_WHEN_ERROR_MANY
'    Stop
'    End
    
    'KILL ITSELF IN __.EXE KILL SOFTLY
    'WHILE ISIDE LEARN
    '---------------------------------
    If IsIDE = True Then
        a = cProcesses.GetEXEID_KILL_ALL_INSTANCE(App.Path + "\" + App.EXEName + ".exe")
        If a = "FOUND SOMETHING" Then
            Beep
            End
        End If
    End If
    
    On Error Resume Next
    If Dir(App.Path + "\# DATA\", vbDirectory) = "" Then
        MkDir App.Path + "\# DATA"
    End If
    
    I_1 = "QUICK MODE -- ON -- NOT MESSENGER"
    On Error Resume Next
    Err.Clear
    FR1 = FreeFile
    If Dir(App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT") <> "" Then
        Open App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT" For Input As #FR1
            Line Input #FR1, I_1
            Line Input #FR1, I_2
        Close FR1
    End If
      I_3 = "QUICK MODE -- ON -- NOT MESSENGER"
    ' I_2 = "QUICK MODE -- OFF -- GIVE MESSENGER"
    ' I_1 = "QUICK MODE -- ON -- NOT MESSENGER"
    ' I_2 = "QUICK MODE -- OFF -- GIVE MESSENGER"
    If I_1 = I_3 Then
        If DateValue(I_2) + TimeValue(I_2) < Now Then
            FR1 = FreeFile
            Open App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT" For Output As #FR1
                I_4 = "QUICK MODE -- OFF -- GIVE MESSENGER"
                Print #FR1, I_4
                Print #FR1, I_2
            Close FR1
        End If
    End If
    MNU_QUICK_MODE.Caption = I_1
    
    W$ = "D:\VB6\VB-NT\00_Best_VB_01\Auto Start Menu\Auto Start Menu.exe"
    W$ = "D:\VB6\VB-NT\00_Best_VB_01\Auto Start Menu\Auto Start Menu.exe"
    
    W$ = "E:\01 FAVS\Ministers\Images"
    
    W$ = "M:\0 00 Art\Google Downloader\00 MICHELLE"
    W$ = "U:\# MY DOCS\# 01 My Documents\02_GOOGLE_EARTH\#Google Earth\My Places ASUS BIG\files"
    W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2009 FaceBook Photo's\#\Part 1\DSC02338.JPG"
    W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 FaceBook Photo's\#\MEADOWFIELD HOSPITAL-2010-12-10 21-41-00 - SONY DSC-HX5V - DSC00321.JPG"
    W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 FaceBook Photo's\#\MEADOWFIELD HOSPITAL-2010-12-10 21-41-14 - SONY DSC-HX5V - DSC00322.JPG"
    W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 FaceBook Photo's\#\MEADOWFIELD HOSPITAL-2010-12-10 21-41-26 - SONY DSC-HX5V - DSC00323.JPG"
    W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 21 13.JPG"
    W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 20 38.JPG"
    W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 19 43.JPG"
    W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 17 52.JPG"
    W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 14 11.JPG"
    W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 04 10 51.JPG"
    W$ = "D:\VB6\VB-NT\00_Best_VB_01\#0 GRAPHIC PATTERN ANIMATION\Yin Yang\YIN YANG 2015 06 22 - 03 54 49.JPG"
    
    W$ = "M:\DSC\2015\10350614"
    W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\Facebook Holiday Inn Hotel"
    W$ = "M:\DSC\2005-2007\2005 NEXT"
    W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2010-04-19 Kings Hotel"
    W$ = "M:\DSC\2010\# TEST"
    W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2008-11-29 Best Western Brighton Stay in On the Run"
    W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2008-12-08 Best Western"
    W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2010-04-09 Best Western\test"
    W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2010-04-16 Holiday Inn"
    W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2010-04-19 Kings Hotel\test"
    W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Facebook Hotel's & Camping\2010-10-27 Meadowfield Hospital Maple Ward Asylum\test"
    W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2009 FaceBook London"
    W$ = "D:\0 00 Art\# 00 My Pictures\01 All Phone\01-Mental\2010 Rutland Gardens"
    W$ = "D:\Videos\00 Vid XXX\ME\2011 GALAXY SAMSUNG GT-P1000 MP4"
    W$ = "D:\Videos\00 Vid XXX\ME"
    W$ = "D:\Videos\1999 Accidents\"
    W$ = "D:\0 00 ART LOGGERS - WEBCAM\VIDEO\"
    W$ = "D:\DSC\# Docus Proofs Texts\# BHT\NOTICE BOARD"
    
    W$ = "D:\DSC\2018 Double Screen Cam\DCIM\2018-10-21"
    W$ = "D:\VI_ DSC ME\2010+NOKIA\" '"2015 NOKIA E72 _ 008 AUG _ MP4 _ x001 _ Home Front Room.MP4"
    ' CARE WHEN MODIFY DATE FROM MOD TO CREATED _
    ' EXPLORER WITH A DISPLAY OF JPG TREAT AS PICTURE AND NOT THE REAL FILE SYSTEM MOD-DATE SELECTION FOR COLOUMN HEADER
    W$ = ""
    
    If Command$ <> "" Or W$ <> "" Then
        If W$ = "" Then W$ = Command$
    End If
    
    If Command$ = "" And W$ = "" Then
        If Clipboard.GetFormat(vbCFText) Then
            W1$ = Clipboard.GetText(vbCFText) ' Get Clipboard text.
            If Mid(W1$, 1, 1) = """" Then
                W1$ = Mid(W1$, 2): W1$ = Mid(W1$, 1, Len(W1$) - 1)
            End If
            If FSO.FolderExists(W1$) = True Then
                W$ = W1$
            End If
            If FSO.FileExists(W1$) = True Then
                W$ = W1$
            End If
        End If
    End If
    
    
    If Mid(W$, 1, 1) = """" Then
        W$ = Mid(W$, 2): W$ = Mid(W$, 1, Len(W$) - 1)
    End If
    FULL_PATH_AND_FILENAME = W$
    
    If FSO.FileExists(W$) Then
        CONVET_FOLDER = True
        W$ = Mid(W$, 1, InStrRev(W$, "\"))
    End If
    
    App.title = "#0 Send To Modify Date With Menu"
    Me.Caption = UCase(App.title)
    
    MNU_VERSION.Caption = "VER_2020_" + Trim(str(App.Major)) + "." + Trim(str(App.Minor)) + "." + Trim(str(App.Revision)) ' + " _ MATT.LAN@BTINTERNET.COM"
    
    If FSO.FolderExists(W$) = False Then
        If CONVET_FOLDER = True Then
'            MsgBox "Not Proper Folder Given" + vbCrLf + vbCrLf + "Command$ = " + vbCrLf + Command$ + vbCrLf + vbCrLf + "Convert File to Folder = " + vbCrLf + vbCrLf + W$, vbMsgBoxSetForeground
'            End
        Else
    '        MsgBox "Not Proper Folder Given" + vbCrLf + vbCrLf + "Command$ = " + vbCrLf + Command$, vbMsgBoxSetForeground
    '        End
        End If
    End If
    
    If W$ <> "" Then
        If Mid$(W$, Len(W$), 1) <> "\" Then
            W$ = W$ + "\"
        End If
    End If
    
    If Mid(W$, 1, 1) = """" Then
        W$ = Mid(W$, 2): W$ = Mid(W$, 1, Len(W$) - 1)
    End If
    
    'FOLDER LABEL
    LABEL_SET(2).Caption = W$
    If W$ = "" Then LABEL_SET(2).Caption = "NOT FOLDER GIVEN"
    
    'FILE LABEL
    LABEL_SET(3).Caption = FULL_PATH_AND_FILENAME  ' -- FULL_PATH_AND_FILENAME
    If FULL_PATH_AND_FILENAME = "" Then LABEL_SET(3).Caption = "NOT FILE GIVEN"
    
    LABEL_SET(1).BackColor = Label_COLOR_YELLOW.BackColor
    
    Load ScanPath
    
    Call frmListMenu.SET_MENU_PADD_WORK
    
End Sub


Sub SET_MOST_RECENT_DATE_TO_OTHER_IN_FOLDER()
    
    ScanPath.txtPath.Text = LABEL_SET(2).Caption
    ' ScanPath.txtPath.Text = "D:\UTILS\2011 GALAXY SAMSUNG GT-P1000 - Copy"
    
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        
        Set F = FSO.GetFILE(A11 + B11)
        DT2 = F.DateCreated
        DT1 = F.datelastmodified
        
        If DT4 = 0 Then DT4 = DT1
        If DT1 > DT4 Then DT4 = DT1
        
        Set F = Nothing
        
    Next
    
    If DT4 = 0 Then
        MsgBox "NAUGHT DATE FOUND IN FILE GATHER -- EXIT"
        End
    End If
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        
        TT = SetFileDateTime(A11 + B11, DT4)
        
        XC = XC + 1
        
    Next
    
    MsgBox "Done = " + vbCrLf + str(XC)
    End

End Sub

Sub SET_OLDER_DATE_TO_OTHER_IN_FOLDER()
    
    ' --------------------------------------------------------------------
    ' FUNNY CODE
    ' PUT DEBUG WATCH AND THIS VARIABLE CHANGE SOON AS ENTER FIRST LINE CODE HERE SUBROUTINE
    ' WORK AROUND REVERT IT BACK
    ' --------------------------------------------------------------------
    ' THIS TYPE ROUTINE DONE FROM CALL BACK SYSTEM
    ' AND PUBLIC BASTARD
    ' AND NOT ABLE TO CARRY ANY VARIABLE THROUGH
    ' VARIABLE MUST BE PUBLIC IN BAS MODULE
    ' If WORK = "SET_OLDER_DATE_TO_OTHER_IN_FOLDER" Then
    ' CallByName FORM_ME, WORK, VbMethod
    ' Mon 15-Jun-2020 14:28:00
    ' UNABLE CHANGE -- MNU_INCREMENT_DATE.Caption
    ' IT GO BACK TO FORM LOAD
    ' WORK AROUND USER VARIABLE AS WHILE GOOD
    ' --------------------------------------------------------------------
    
    
    ScanPath.txtPath.Text = LABEL_SET(2).Caption
    
    DV = 0
    If InStr(ScanPath.txtPath.Text, "\00 ROOT_02_(DATE-ALPHABETICAL)\") > 0 Then DV = 1
    If InStr(ScanPath.txtPath.Text, "\ABBYWINTERS.COM_DATE") > 0 Then DV = 1
    If DV = 1 Then
        MENU_INCREMENT_DATE_VAR = -1
        Call MNU_INCREMENT_DATE_SHOW
    End If

    
    If LABEL_SET(2).Caption = "NOT FOLDER GIVEN" Then
        MsgBox "NOT FOLDER GIVEN"
        Exit Sub
    End If
    
    
'    If IsIDE = True Then
'        ScanPath.txtPath.Text = "D:\VIDEO\NOT\X 00 NOT ME\00 Vid XXX\ABBYWINTERS.COM_DATE"
'    End If
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    'D:\VIDEO\NOT\X 01 ME\2017 SONY MP4\DOC\2017 05 31
    
    For RR = ScanPath.ListView1.ListItems.Count To 1 Step -1
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        If InStr(UCase(B11), ".TXT") > 0 Then
            ' ScanPath.ListView1.ListItems.Remove (RR)
        End If
    Next
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        
        FLAG_OPER = True
        If UCase(B11) = UCase("thumbs.db") Then FLAG_OPER = False
        If UCase(B11) = UCase("DESKTOP.INI") Then FLAG_OPER = False
        
        Set F = FSO.GetFILE(A11 + B11)
        DT2 = F.DateCreated
        DT1 = F.datelastmodified
        If FLAG_OPER = True Then
            If DT4 = 0 Then DT4 = DT1
            If DT1 < DT4 Then DT4 = DT1
        End If
        Set F = Nothing
    Next
    
    If DT4 = 0 Then
        MsgBox "NAUGHT DATE FOUND IN FILE GATHER -- EXIT"
        End
    End If
    
    COUNT_FILE = ScanPath.ListView1.ListItems.Count
    If MENU_INCREMENT_DATE_VAR = -1 Then
        If COUNT_FILE > 1 Then
            DT4 = DT4 + TimeSerial(0, 1, 0)
        End If
    End If
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        
        TT = SetFileDateTime(A11 + B11, DT4)
        XC = XC + 1
        
        If ScanPath.ListView1.ListItems.Count > 2 Then
        FI = UCase(B11)
        FI = Mid(FI, InStrRev(FI, ".") + 1) + " "
        ' If InStr("MPG MPEG MP4 JPG AVI ", FI) > 0 Then
            If MENU_INCREMENT_DATE_VAR = -1 Then
                DT4 = DT4 + TimeSerial(0, 1, 0)
            End If
        ' End If
        End If
    Next
    On Error Resume Next
    FR1 = FreeFile
    Open App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT" For Input As #FR1
        Line Input #FR1, I_1
    Close FR1
    MNU_QUICK_MODE.Caption = I_1
    I_1 = "QUICK MODE -- ON -- NOT MESSENGER"
    I_2 = "QUICK MODE -- OFF -- GIVE MESSENGER"
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC)
    End If
    End
End Sub

Sub SET_DATE_OF_FILENAME_YYYY_MM_DD__WITH_ONE_FOLDER()

    ScanPath.chkSubFolders = vbUnchecked
    ScanPath.cboMask.Text = "*.*"
    
    ' SCAN DO ON TEXTPATH CHANGE
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    
    If IsIDE Then
        '2009-02-17-4896_S03_.WMV
        ScanPath.txtPath.Text = "D:\VIDEO\NOT\X 00 NOT ME\00 Vid XXX\BUNKER.COM\2009-06 JUN"
    End If
    
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    If ScanPath.ListView1.ListItems.Count > 500 Then
        MsgBox "LOOK LIKE GOT TOO MANY FILE TO DO OVER 500" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
        End
    End If
    
    If ScanPath.ListView1.ListItems.Count > 0 Then
        MsgBox "FILE COUNTER -- BEFORE FILTER ON" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
    End If
    
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        If InStr("WMV MP4 TXT", EXT_STR) And InStr(A11, "_gsdata_") = 0 Then
            DATE_FILENAME_D_1 = Mid(B11, 1, 10)
            DATE_FILENAME_T_2 = "10:00:00" ' Mid(B11, 20, 8) -- CLOCK TIMEZONE IF MIDNIGHT DAY OFFSET OUT KEEP IN
            i2 = DATE_FILENAME_D_1
            i4 = DATE_FILENAME_T_2
            DATE_FILENAME_D_1 = Mid(i2, 1, 4) + "/" + Mid(i2, 6, 2) + "/" + Mid(i2, 9, 2)
            DATE_FILENAME_T_2 = Mid(i4, 1, 2) + ":" + Mid(i4, 4, 2) + ":" + Mid(i4, 7, 2)
            
            DATE_FILENAME_SUCCESS = DateValue(DATE_FILENAME_D_1) + TimeValue(DATE_FILENAME_T_2)
            
            DT4 = DATE_FILENAME_SUCCESS
            If IsDate(DT4) = False Then
                ' FALSE DATE
                ' ----------
                MsgBox "DATE HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
                End
            End If
            
            If DT4 = 0 Then
                MsgBox "NAUGHT DATE FINDER -- EXIT"
                End
            End If
        
            TT = SetFileDateTime(A11 + B11, DT4)
            
            EXT_ORG = Mid(B11, Len(B11) - 2)
            EXT_STR = UCase(Mid(B11, Len(B11) - 2))
            If EXT_ORG <> EXT_STR Then
                'RENAME EXTENSION CAPITAL
                FILE_RENAME1 = Mid(B11, 1, Len(B11) - 4) + "." + EXT_STR
                Name A11 + B11 As A11 + FILE_RENAME1
            End If
            
            XC = XC + 1
            MM_1 = MM_1 + B11 + vbCrLf
        End If
    Next
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End
End Sub

Sub RENAME____DSC0_____JPG__TO__YYYY_MM_DD_HH_MM_SS___SONY___DSC_HX60V___DSC0_____JPG____BATCH()

    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        File1.Path = LABEL_SET(2).Caption
    End If
    
    If IsIDE = True Then
        File1.Path = "D:\DSC\2015+SONY\2020 CyberShot HX60V\JPG\2020 06 20"
    End If
    File1.Pattern = "*.JPG"
    
    ' SOMETIME THE PASS BY CONTEXT MENU ISN'T LONG NAME BUT SHORTNAME
        
    R_C_COUNTER_MAX = 0
    For R_C_4 = 1 To 3
        A_M = ""
        R_C_COUNTER = 0
        For R_C = 0 To File1.ListCount - 1
            If Mid(File1.List(R_C), 1, 3) = "DSC" Then
                FILE_NAME_JPG = File1.Path + "\" + File1.List(R_C)
            
                If Mid(FILE_NAME_JPG, 1, 2) <> "\\" Then
                    FILE_NAME_JPG = GetLongName(FILE_NAME_JPG)
                End If
                
                Set F = FSO.GetFILE((FILE_NAME_JPG))
                ADATE1 = F.datelastmodified
                
                'Clipboard.Clear
                'Sleep 200
                'If Year(ADATE1) = 2016 Then at1 = "-- 2k Sixteenth"
                'If Year(ADATE1) = 2015 Then at1 = "-- 2k Fifteenth"
                'If Year(ADATE1) <= 2015 Then at1 = "-- " + Format(ADATE1, "YYYY")
                'If Year(ADATE1) > 2016 Then at1 = "-- " + Format(ADATE1, "YYYY")
                'AT1 = "-- " + Format(ADATE1, "YYYY")
                
                ' 2020-06-20 09-48-48 - SONY DSC-HX60V - DSC03279
                FILE_NAME_JPG_2 = Mid(FILE_NAME_JPG, InStrRev(FILE_NAME_JPG, "\") + 1)
                FILE_NAME_JPG_2_EXT = Mid(FILE_NAME_JPG, InStrRev(FILE_NAME_JPG, "."))
                FILE_NAME_JPG_2_NOT_EXT = UCase(Mid(File1.List(R_C), 1, InStrRev(File1.List(R_C), ".") - 1))
                FILE_NAME_JPG_SET = FILE_NAME_JPG_2_NOT_EXT + UCase(FILE_NAME_JPG_2_EXT)
                DATE_FORMATER = Format(ADATE1, "YYYY-MM-DD HH-MM-SS") + " - SONY DSC-HX60V - " + FILE_NAME_JPG_SET
                FILE_NAME_JPG_SET_AND_DATE = File1.Path + "\" + DATE_FORMATER
                
                If File1.List(R_C) <> FILE_NAME_JPG_SET_AND_DATE Then
                    R_C_COUNTER = R_C_COUNTER + 1
                    If R_C_4 = 1 Then
                        R_C_COUNTER_MAX = R_C_COUNTER_MAX + 1
                    End If
                    
                    A_M = A_M + "RENAMER --" + str(R_C_COUNTER) + " OF" + str(R_C_COUNTER_MAX) + " OF" + str(File1.ListCount) + vbCrLf + File1.Path + "\" + vbCrLf + "\" + File1.List(R_C) + vbCrLf + "----" + vbCrLf + FILE_NAME_JPG_SET_AND_DATE
                    A_M = A_M + vbCrLf
                    A_M = A_M + vbCrLf
                    
                    If R_C_4 = 3 Then
                    
                        If R_C_COUNTER < 2 Then
                        
                            A_M_2 = "RENAMER -- VERIFY 1ST 2  AND THEN AUTO" + vbCrLf + Trim(str(R_C + 1)) + " OF" + str(File1.ListCount) + vbCrLf + File1.Path + "\" + vbCrLf + "\" + File1.List(R_C) + vbCrLf + "----" + vbCrLf + FILE_NAME_JPG_SET_AND_DATE
                            
                            X_M = MsgBox(A_M_2, vbYesNo)
                            If X_M = vbNo Then
                                R_C_COUNTER = R_C_COUNTER - 1
                                Exit For
                            End If
                        End If
                        
                        Name File1.Path + "\" + File1.List(R_C) As FILE_NAME_JPG_SET_AND_DATE
                    
                    End If
                End If
            End If
        Next
            
        If R_C_4 = 2 Then
            If A_M <> "" Then
                R_C_COUNTER = 0
                X_M = MsgBox("RENAME CONVERSION AS FOLLOW PROCESS YES / NOT" + vbCrLf + vbCrLf + A_M, vbYesNo)
                If X_M = vbNo Then Exit For
            End If
        End If
    Next
        
    If R_C_COUNTER > 0 Then
    
    MsgBox "DONE " + Trim(str(R_C_COUNTER)) + " CHANGER"
    
    End If
    
    End
End Sub

Sub RENAME____YYYY_MM_DD_HH_MM_SS__ANY____BATCH()
    
    Dim INSERT_HERE
    
    INSERT_HERE = " - "
    
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        File1.Path = LABEL_SET(2).Caption
    End If
    
    If IsIDE = True Then
        File1.Path = "D:\0 00 ART\# 00 My Pictures\01 All Phone\01 Mental\1990-2007 Matthew-Lancaster\1990 and 1997 Scanned-Photo's\New folder\2015 08 07 - CANON SCANNER\DOC"
        INSERT_HERE = " - Canon-MP495 Series - "
    End If
    File1.Pattern = "*.*"
    
    ' SOMETIME THE PASS BY CONTEXT MENU ISN'T LONG NAME BUT SHORTNAME
        
    R_C_COUNTER_MAX = 0
    For R_C_2 = 1 To 3
        
        A_M = ""
        
        R_C_COUNTER = 0
            
        For R_C = 0 To File1.ListCount - 1
    
            FILE_NAME_MP4 = File1.Path + "\" + File1.List(R_C)
        
            If Mid(FILE_NAME_MP4, 1, 2) <> "\\" Then
                FILE_NAME_MP4 = GetLongName(FILE_NAME_MP4)
            End If
            
            Set F = FSO.GetFILE((FILE_NAME_MP4))
            ADATE1 = F.datelastmodified
            
            'Clipboard.Clear
            'Sleep 200
            'If Year(ADATE1) = 2016 Then at1 = "-- 2k Sixteenth"
            'If Year(ADATE1) = 2015 Then at1 = "-- 2k Fifteenth"
            'If Year(ADATE1) <= 2015 Then at1 = "-- " + Format(ADATE1, "YYYY")
            'If Year(ADATE1) > 2016 Then at1 = "-- " + Format(ADATE1, "YYYY")
            
            'at1 = "-- " + Format(ADATE1, "YYYY")
            
            '------------------------
            ' LONG DATE FOR YOUTUBING
            '------------------------
            'If InStr(FILE_NAME_MP4, "D:\DSC\") > 0 Then
            'Clipboard.SetText " " + at1 + Format(ADATE1, " MMMM DD DDDD HH-MM-SS Am/Pm")
            FILE_NAME_MP4_2 = Mid(FILE_NAME_MP4, InStrRev(FILE_NAME_MP4, "\") + 1)
            FILE_NAME_MP4_2_EXT = Mid(FILE_NAME_MP4, InStrRev(FILE_NAME_MP4, "."))
            FILE_NAME_MP4_2_NOT_EXT = UCase(Mid(File1.List(R_C), 1, InStrRev(File1.List(R_C), ".") - 1))
            FILE_NAME_MP4_SET = FILE_NAME_MP4_2_NOT_EXT + UCase(FILE_NAME_MP4_2_EXT)
            FILE_NAME_MP4_SET_AND_DATE = File1.Path + "\" + Format(ADATE1, "YYYY-MM-DD HH-MM-SS") + INSERT_HERE + FILE_NAME_MP4_SET
            MAQ = ""
            START_MAQ = 0
            
            ' Call MNU_FILEDATE_CLIPBOARD_ROUTINE_CHECKER
            
            If START_MAQ = 1 Or 1 = 1 Then
                If File1.List(R_C) <> FILE_NAME_MP4_SET_AND_DATE Then
                    ' If MAQ <> "" Then
                    
                        R_C_COUNTER = R_C_COUNTER + 1
                        If R_C_2 = 1 Then
                            R_C_COUNTER_MAX = R_C_COUNTER_MAX + 1
                        End If
                        
                        A_M = A_M + "RENAMER --" + str(R_C_COUNTER) + " OF" + str(R_C_COUNTER_MAX) + " OF" + str(File1.ListCount) + vbCrLf + File1.Path + "\" + vbCrLf + "\" + File1.List(R_C) + vbCrLf + "----" + vbCrLf + FILE_NAME_MP4_SET_AND_DATE
                        A_M = A_M + vbCrLf
                        A_M = A_M + vbCrLf
                        
                        If R_C_2 = 3 Then
                        
                            If R_C_COUNTER < 10 Then
                            
                                A_M_2 = "RENAMER -- VERIFY 1ST 10  AND THEN AUTO" + vbCrLf + Trim(str(R_C + 1)) + " OF" + str(File1.ListCount) + vbCrLf + File1.Path + "\" + vbCrLf + "\" + File1.List(R_C) + vbCrLf + "----" + vbCrLf + FILE_NAME_MP4_SET_AND_DATE
                                
                                X_M = MsgBox(A_M_2, vbYesNo)
                                If X_M = vbNo Then
                                    R_C_COUNTER = R_C_COUNTER - 1
                                    Exit For
                                End If
                            End If
                            
                            Name File1.Path + "\" + File1.List(R_C) As FILE_NAME_MP4_SET_AND_DATE
                        
                        End If
                    ' End If
                End If
            End If
        Next
            
        If R_C_2 = 2 Then
            If A_M <> "" Then
                R_C_COUNTER = 0
                X_M = MsgBox("RENAME CONVERSION AS FOLLOW PROCESS YES / NOT" + vbCrLf + vbCrLf + A_M, vbYesNo)
                If X_M = vbNo Then Exit For
            End If
        End If
    Next
        
    If R_C_COUNTER > 0 Then
    
    MsgBox "DONE " + Trim(str(R_C_COUNTER)) + " CHANGER"
    
    End If
    
    End

End Sub

Sub RENAME____YYYY_MM_DD_MMM_DDD_HH_MM_SS__AVI_MP4_ANY____BATCH()
    Call MNU_FILEDATE_WHOLE_FOLDER_AVI_MP4_ANY_Click
End Sub

Sub RENAME____YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA__MP4____BATCH()
    Call MNU_FILEDATE_WHOLE_FOLDER_Click
End Sub

'
Sub SET_DATE_OF_FILENAME_CH00_YYYYMMDD__HHMMSS_MP4_HIKVISION_BATCH()
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    ScanPath.chkSubFolders = vbUnchecked
    ScanPath.cboMask.Text = "*.MP4;*.AVI;*.TXT"
    
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    If LABEL_SET(2).Caption = "NOT FOLDER GIVEN" Then
        If IsIDE Then
            ScanPath.txtPath.Text = "U:\VI_ DSC ME\2020+CCTV_HIKVISION\2020-02-25_CH04____FRONT ROOM\"
        Else
            MsgBox "NOT FOLDER GIVEN"
        End If
    End If
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    'If ScanPath.ListView1.ListItems.Count > 0 Then
        'MsgBox "FILE COUNTER -- BEFORE FILTER ON" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
    'End If
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        B11_NOT_EXT = Mid(B11, 1, Len(B11) - 4)
        
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        If InStr("MP4 TXT AVI", EXT_STR) Then
            DATE_FILENAME_D_1 = Mid(B11, 12, 2) + "/" + Mid(B11, 10, 2) + "/" + Mid(B11, 6, 4)
            DATE_FILENAME_T_2 = Mid(B11, 16, 2) + ":" + Mid(B11, 18, 2) + ":" + Mid(B11, 20, 2)
            If IsDate(DATE_FILENAME_D_1) = True Then
                Mid(DATE_FILENAME_D_1, 5, 1) = "/"
                Mid(DATE_FILENAME_D_1, 8, 1) = "/"
                Mid(DATE_FILENAME_T_2, 3, 1) = ":"
                Mid(DATE_FILENAME_T_2, 6, 1) = ":"
            End If
            If IsDate(DATE_FILENAME_D_1) = False Then
                ' BOTH THE SAME FORMAT AS ABOVE
                ' CH01_20200225--142741
                ' IN YYYYMMDD
                ' OUT DD-MM-YYYY
                DATE_FILENAME_D_1 = Mid(B11, 12, 2) + "/" + Mid(B11, 10, 2) + "/" + Mid(B11, 6, 4)
                DATE_FILENAME_T_2 = Mid(B11, 16, 2) + ":" + Mid(B11, 18, 2) + ":" + Mid(B11, 20, 2)
            End If
            
            DATE_FILENAME_SUCCESS = DateValue(DATE_FILENAME_D_1) + TimeValue(DATE_FILENAME_T_2)
            If IsDate(DATE_FILENAME_SUCCESS) = False Then
                a = a
            End If
            
            
            DT4 = DATE_FILENAME_SUCCESS
            If IsDate(DT4) = False Then
                ' FALSE DATE
                ' ----------
                MsgBox "DATE HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
                End
            End If
            
            If DT4 = 0 Then
                MsgBox "NAUGHT DATE FINDER -- EXIT"
                End
            End If
        
            TT = SetFileDateTime(A11 + B11, DT4)
            
            EXT_ORG = Mid(B11, Len(B11) - 2)
            EXT_STR = UCase(Mid(B11, Len(B11) - 2))
            If EXT_ORG <> EXT_STR Then
                'RENAME EXTENSION CAPITAL
                FILE_RENAME1 = Mid(B11, 1, Len(B11) - 4) + "." + EXT_STR
                Name A11 + B11 As A11 + FILE_RENAME1
            End If
            
            XC = XC + 1
            MM_1 = MM_1 + B11 + vbCrLf
        End If
    Next
    
    Call SET_QUICK_MODE_RESULT
    ' If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    ' End If
    
    End

End Sub



Sub SET_DATE_OF_FILENAME_CH00_YYYY_MM_DD_HH_MM_SS_MP4_HIKVISION_BATCH()
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    ScanPath.chkSubFolders = vbUnchecked
    ScanPath.cboMask.Text = "*.MP4;*.AVI"
    
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    If LABEL_SET(2).Caption = "NOT FOLDER GIVEN" Then
        If IsIDE Then
            ScanPath.txtPath.Text = "F:\DSC--2018+CCSE_HIKVISION\2020_CCSE_HIKVISION_DVR-104G-F1\2020-06-14"
        Else
            MsgBox "NOT FOLDER GIVEN"
        End If
    End If
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    'If ScanPath.ListView1.ListItems.Count > 0 Then
        'MsgBox "FILE COUNTER -- BEFORE FILTER ON" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
    'End If
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        B11_NOT_EXT = Mid(B11, 1, Len(B11) - 4)
        
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        If InStr("MP4 TXT", EXT_STR) Then
            DATE_FILENAME_D_1 = Mid(B11, 6, 10)
            DATE_FILENAME_T_2 = Mid(B11, 18, 8)
            If IsDate(DATE_FILENAME_D_1) = True Then
                Mid(DATE_FILENAME_D_1, 5, 1) = "/"
                Mid(DATE_FILENAME_D_1, 8, 1) = "/"
                Mid(DATE_FILENAME_T_2, 3, 1) = ":"
                Mid(DATE_FILENAME_T_2, 6, 1) = ":"
            End If
            If IsDate(DATE_FILENAME_D_1) = False Then
                ' CH01_20200225--142741
                ' IN YYYYMMDD
                ' OUT DD-MM-YYYY
                DATE_FILENAME_D_1 = Mid(B11, 12, 2) + "/" + Mid(B11, 10, 2) + "/" + Mid(B11, 6, 4)
                DATE_FILENAME_T_2 = Mid(B11, 16, 2) + ":" + Mid(B11, 18, 2) + ":" + Mid(B11, 20, 2)
            End If
            
            DATE_FILENAME_SUCCESS = DateValue(DATE_FILENAME_D_1) + TimeValue(DATE_FILENAME_T_2)
            If IsDate(DATE_FILENAME_SUCCESS) = False Then
                a = a
            End If
            
            
            DT4 = DATE_FILENAME_SUCCESS
            If IsDate(DT4) = False Then
                ' FALSE DATE
                ' ----------
                MsgBox "DATE HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
                End
            End If
            
            If DT4 = 0 Then
                MsgBox "NAUGHT DATE FINDER -- EXIT"
                End
            End If
        
            TT = SetFileDateTime(A11 + B11, DT4)
            
            EXT_ORG = Mid(B11, Len(B11) - 2)
            EXT_STR = UCase(Mid(B11, Len(B11) - 2))
            If EXT_ORG <> EXT_STR Then
                'RENAME EXTENSION CAPITAL
                FILE_RENAME1 = Mid(B11, 1, Len(B11) - 4) + "." + EXT_STR
                Name A11 + B11 As A11 + FILE_RENAME1
            End If
            
            XC = XC + 1
            MM_1 = MM_1 + B11 + vbCrLf
        End If
    Next
    
    Call SET_QUICK_MODE_RESULT
    ' If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    ' End If
    
    End

End Sub


Sub SET_DATE_OF_FILENAME_CH00_YYYY_MM_DD_HH_MM_SS_MP4_HIKVISION_SINGLE()
    ' ---------------------------------------------------------------------
    ' ANOTHER SUB ROUTINE
    ' LEECH AND MODIFY FROM BATCH CODE SET
    ' Fri 28-Feb-2020 07:18:00
    ' ---------------------------------------------------------------------
    ' LABEL_SET(3).Caption ' -- FULL_PATH_AND_FILENAME
    ' ---------------------------------------------------------------------
    A11 = LABEL_SET(3).Caption
    D11 = A11
    If A11 <> "NOT FILE GIVEN" Then
        A11 = Mid(D11, 1, InStrRev(A11, "\"))
        B11 = Mid(D11, InStrRev(A11, "\") + 1)
    End If
    If LABEL_SET(2).Caption = "NOT FOLDER GIVEN" Then
'        If IsIDE Then
'            D11 = "F:\DSC--2018+CCSE_HIKVISION_SCREENCASTIFY\2020-01-27--EDDIE\CH01_2020 01 27--16 47 36_BLAGGER_CCTV-1.mp4"
'            A11 = Mid(D11, 1, InStrRev(D11, "\"))
'            B11 = Mid(D11, InStrRev(D11, "\") + 1)
'        End If
    End If
    If Len(D11) < 5 Or D11 = "NOT FILE GIVEN" Then
        MsgBox "FULL PATH AND FILE NOT FIND -- EXIT" + vbCrLf + vbCrLf + D11
        End
    End If
    
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    EXT_STR = UCase(Mid(B11, Len(B11) - 2))
    If InStr("MP4 TXT", EXT_STR) Then
        DATE_FILENAME_D_1 = Mid(B11, 6, 10)
        DATE_FILENAME_T_2 = Mid(B11, 18, 8)
        If IsDate(DATE_FILENAME_D_1) = True Then
            Mid(DATE_FILENAME_D_1, 5, 1) = "/"
            Mid(DATE_FILENAME_D_1, 8, 1) = "/"
            Mid(DATE_FILENAME_T_2, 3, 1) = ":"
            Mid(DATE_FILENAME_T_2, 6, 1) = ":"
        End If
        If IsDate(DATE_FILENAME_D_1) = False Then
            ' CH01_20200225--142741
            ' IN YYYYMMDD
            ' OUT DD-MM-YYYY
            DATE_FILENAME_D_1 = Mid(B11, 12, 2) + "/" + Mid(B11, 10, 2) + "/" + Mid(B11, 6, 4)
            DATE_FILENAME_T_2 = Mid(B11, 16, 2) + ":" + Mid(B11, 18, 2) + ":" + Mid(B11, 20, 2)
        End If
        
        DATE_FILENAME_SUCCESS = DateValue(DATE_FILENAME_D_1) + TimeValue(DATE_FILENAME_T_2)
        If IsDate(DATE_FILENAME_SUCCESS) = False Then
        a = a
        End If
        
        
        DT4 = DATE_FILENAME_SUCCESS
        If IsDate(DT4) = False Then
            ' FALSE DATE
            ' ----------
            MsgBox "DATE HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
            End
        End If
        
        If DT4 = 0 Then
            MsgBox "NAUGHT DATE FINDER -- EXIT"
            End
        End If
    
        TT = SetFileDateTime(A11 + B11, DT4)
        
        EXT_ORG = Mid(B11, Len(B11) - 2)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        If EXT_ORG <> EXT_STR Then
            'RENAME EXTENSION CAPITAL
            FILE_RENAME1 = Mid(B11, 1, Len(B11) - 4) + "." + EXT_STR
            Name A11 + B11 As A11 + FILE_RENAME1
        End If
        
        XC = XC + 1
        MM_1 = MM_1 + B11 + vbCrLf
    End If
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End

End Sub

Sub SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA_SINGLE()
    ' ---------------------------------------------------------------------
    ' ANOTHER SUB ROUTINE
    ' LEECH AND MODIFY FROM BATCH CODE SET
    ' Wed 29-Jan-2020 06:38:17
    ' ---------------------------------------------------------------------
    ' LABEL_SET(3).Caption ' -- FULL_PATH_AND_FILENAME
    ' ---------------------------------------------------------------------
    
    A11 = LABEL_SET(3).Caption
    ' A11 = "\\4-asus-gl522vw\4_asus_gl522vw_10_1_samsung_1tb\DSC\2015+SONY\2020 CyberShot HX60V\MP4\2020_01_20 Jan_Mon 12_25_38__MAH01294_HOVE GRAND AVENUE DEMOLITION__.MP4"
    D11 = A11
    If A11 <> "NOT FILE GIVEN" Then
        A11 = Mid(D11, 1, InStrRev(A11, "\"))
        B11 = Mid(D11, InStrRev(A11, "\") + 1)
    End If
    
    If Len(D11) < 5 Or D11 = "NOT FILE GIVEN" Then
        MsgBox "FULL PATH AND FILE NOT GOOD -- EXIT" + vbCrLf + vbCrLf + D11
        End
    End If
    
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    EXT_STR = UCase(Mid(B11, Len(B11) - 2))
    If InStr("MP4 TXT", EXT_STR) Then
        DATE_FILENAME_D_1 = Mid(B11, 1, 10)
        DATE_FILENAME_T_2 = Mid(B11, 20, 8)
        i2 = DATE_FILENAME_D_1
        i4 = DATE_FILENAME_T_2
        DATE_FILENAME_D_1 = Mid(i2, 1, 4) + "/" + Mid(i2, 6, 2) + "/" + Mid(i2, 9, 2)
        DATE_FILENAME_T_2 = Mid(i4, 1, 2) + ":" + Mid(i4, 4, 2) + ":" + Mid(i4, 7, 2)
        
        DATE_FILENAME_SUCCESS = DateValue(DATE_FILENAME_D_1) + TimeValue(DATE_FILENAME_T_2)
        
        DT4 = DATE_FILENAME_SUCCESS
        If IsDate(DT4) = False Then
            ' FALSE DATE
            ' ----------
            MsgBox "DATE HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
            End
        End If
        
        If DT4 = 0 Then
            MsgBox "NAUGHT DATE FINDER -- EXIT"
            End
        End If
    
        TT = SetFileDateTime(A11 + B11, DT4)
        
        EXT_ORG = Mid(B11, Len(B11) - 2)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        If EXT_ORG <> EXT_STR Then
            'RENAME EXTENSION CAPITAL
            FILE_RENAME1 = Mid(B11, 1, Len(B11) - 4) + "." + EXT_STR
            Name A11 + B11 As A11 + FILE_RENAME1
        End If
        
        XC = XC + 1
        MM_1 = MM_1 + B11 + vbCrLf
    End If
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End

End Sub

Sub SET_QUICK_MODE_RESULT()

    ' SET QUICK MODE RESULT
    ' ---------------------
    On Error Resume Next
    FR1 = FreeFile
    Open App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT" For Input As #FR1
        Line Input #FR1, I_1
    Close FR1
    MNU_QUICK_MODE.Caption = I_1
    I_1 = "QUICK MODE -- ON -- NOT MESSENGER"
    I_2 = "QUICK MODE -- OFF -- GIVE MESSENGER"

End Sub

Sub DATE_CONVERTOR___MMM_D__YYYY_H_MM_AM____TO_YYYY_MM_DD__HH_MM_DD_FOR_SCREENCASTIFY()

    ScanPath.chkSubFolders = vbUnchecked
    ScanPath.cboMask.Text = "*.MP4"
    
    ' SCAN DO ON TEXTPATH CHANGE
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    If LABEL_SET(2).Caption = "NOT FOLDER GIVEN" Then
        If IsIDE Then
            ScanPath.txtPath.Text = "F:\DSC--SCREENCASTIFY__GD_CLOUD"
        End If
    End If
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
'    If ScanPath.ListView1.ListItems.Count > 500 Then
'        MsgBox "LOOK LIKE GOT TOO MANY FILE TO DO OVER 500" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
'        End
'    End If
    
    'If ScanPath.ListView1.ListItems.Count > 0 Then
        MsgBox "FILE COUNTER -- BEFORE FILTER ON" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
    'End If
    
    
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        If InStr("MP4 TXT", EXT_STR) And InStr(A11, "_gsdata_") = 0 Then
            
            Set F = FSO.GetFILE(A11 + B11)
            DT1 = 0
            DT1 = F.datelastmodified
            DT1 = 0
            'Aug 5, 2019 9 57 Am-4
            DATE_STRING = ""
            If InStrRev(B11, "m-") > 0 Then
                THREE_BACK_SPACE = InStrRev(B11, " ", Len(B11))
                For r = 1 To 2
                    THREE_BACK_SPACE = InStrRev(B11, " ", THREE_BACK_SPACE - 1)
                Next
                DATE_STRING_1 = Mid(B11, 1, THREE_BACK_SPACE)
                DATE_STRING_2 = Mid(B11, THREE_BACK_SPACE, InStrRev(B11, "m-") - THREE_BACK_SPACE + 1)
                DATE_STRING_2 = Trim(DATE_STRING_2)
            End If
            DATE_STRING_2 = Replace(DATE_STRING_2, " ", ":", , 1)
            If DATE_STRING_1 <> "" Then
            If IsDate(DATE_STRING_1) Then
                DT1 = DateValue(DATE_STRING_1)
            End If
            End If
            
            If DATE_STRING_2 <> "" Then
            If IsDate(DATE_STRING_2) Then
                DT1 = DT1 + TimeValue(DATE_STRING_2)
            End If
            End If
            
            DT1_STRING = Format(DT1, "YYYY-MM-DD--HH-MM-SS")
            WHAT_AT_AFTER_MARK_DASH = Mid(B11, InStrRev(B11, "-"))
            DT1_STRING = DT1_STRING + WHAT_AT_AFTER_MARK_DASH
            
            If DT1 > 0 Then
                
                On Error Resume Next
                Err.Clear
                Name A11 + "\" + B11 As A11 + "\" + DT1_STRING
                If Err.Number > 0 Then
                
                    DETAIL_VAR = "RENAME FROM --" + vbCrLf + vbCrLf + A11 + "\" + B11
                    DETAIL_VAR = "MOVE TO   --" + vbCrLf + vbCrLf + A11 + "\" + DT1_STRING
                    MsgBox "WAS A ERROR RENAME FILE REQUEST" + vbCrLf + vbCrLf + "Err.Description = " + Err.Description + vbCrLf + vbCrLf + DETAIL_VAR, vbMsgBoxSetForeground
                
                End If
                TT = SetFileDateTime(A11 + "\" + DT1_STRING, DT1)
                
                XC = XC + 1
                MM_1 = MM_1 + B11 + vbCrLf
            End If
        End If
    Next
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End


End Sub

Sub CAPITAL_ALL____PATH_AND_FILE_____OPERATION____TOP_DEPTH()
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    
    Dim TT
    Dim RR
    
    GATHER_FOLDER_AND_FILE_MODE_FOR_SCANPATH = True
    If IsIDE Then
        LABEL_SET(2).Caption = "M:\M\07 MIDI MOD\MIDI\02 2207 RAG MIDI\04 0004 RAG TIME"
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    
    
    ' SCAN DO ON LABEL_SET.CAPTION TEXTPATH CHANGE
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    
    If LABEL_SET(2).Caption = "NOT FOLDER GIVEN" Then
        If IsIDE Then
            ' ScanPath.txtPath.Text = "F:\RETEKESS M3\M\04 MY MUSIC\03 MP3 CD\##00 TOP SET\00 TOP"
            ' ScanPath.txtPath.Text = "D:\0 00 MUSIC ---\04 MY MUSIC\03 MP3 CD\##00 TOP SET\00 TOP"
        End If
    End If
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    If IsIDE Then
        ' ScanPath.txtPath.Text = "F:\RETEKESS M3\M\04 MY MUSIC\03 MP3 CD\##00 TOP SET\00 TOP"
        ' ScanPath.txtPath.Text = "D:\0 00 MUSIC ---\04 MY MUSIC\03 MP3 CD\##00 TOP SET\00 TOP"
    End If
    
    For RR = ScanPath.ListView1.ListItems.Count To 1 Step -1
        ' A11 ARE DIR
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        SET_REMOVE = False
        If InStr(A11, "_gsdata_") > 0 Then SET_REMOVE = True
        ' If InStr(B11, "_gsdata_") > 0 Then SET_REMOVE = True
        If SET_REMOVE = True Then
            ScanPath.ListView1.ListItems.Remove (RR)
        End If
    Next
    
    If ScanPath.ListView1.ListItems.Count > 5000 Then
        For RR = ScanPath.ListView1.ListItems.Count To 1 Step -1
            ' A11 ARE DIR
            A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
            B11 = ScanPath.ListView1.ListItems.Item(RR)
            D11 = D11 + A11 + B11 + vbCrLf
        Next
        MsgBox "LOOK LIKE GOT TOO MANY FILE TO DO OVER 500" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + vbCrLf + vbCrLf + D11
        End
    End If
    If ScanPath.ListView1.ListItems.Count > 0 Then
        Call SET_QUICK_MODE_RESULT
        If InStr(MNU_QUICK_MODE.Caption, I_3) Then
            MsgBox "FOLDER COUNTER" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
        End If
    Else
        MsgBox "NONE COUNT OF FOLDER HERE " + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
        End
    End If

    For RR = 1 To ScanPath.ListView1.ListItems.Count
        ' A11 ARE DIR
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        A22 = A11 + B11
        If FSO.FolderExists(A22) = True Then
            TYPE_INFO = "FOLDER"
        Else
            TYPE_INFO = "FILE"
        End If
        ' ----------------------------------------------------------
        ' --------------------------
        ' SET THE FILE ____ 01 OF 02
        ' --------------------------
        
        If TYPE_INFO = "FILE" Then
            On Error Resume Next
            Err.Clear
            If B11 <> UCase(B11) Then
                Name A22 As UCase(A22)
                XC = XC + 1
            End If
            COUNT_1_ERROR = Err.Number
            If Err.Number > 0 Then
                Debug.Print Err.Description
                Set F2 = FSO.GetFILE(A22)
                If F2.Name <> UCase(A22) Then
                If B11 <> UCase(B11) Then
                    Err.Clear
                    Name A22 As UCase(A22) + "#"
                    If FSO.FileExists(A22 + "#") = True Then
                        Err.Clear
                        Name A22 + "#" As UCase(A22)
                        If Err.Number > 0 Then
                            MsgBox "ERROR FOLDER NOT ABLE RENAME BACK FROM TEMP NAME " + vbCrLf + vbCrLf + A22 + "#" + vbCrLf + vbCrLf
                            Shell "EXPLORER """ + A22 + """", vbMaximizedFocus
                            COUNT_2_ERROR = Err.Number
                            MM_1 = MM_1 + A22 ' UCase(A22)
                            If Err.Number > 0 Then MM_1 = MM_1 + vbCrLf + Err.Description + vbCrLf + "COUNT 1 ERROR =" + str(COUNT_1_ERROR) + vbCrLf + "COUNT 2 ERROR =" + str(COUNT_2_ERROR): ERROR_SET = True
                            MM_1 = MM_1 + vbCrLf
                            Exit For
                        End If
                    End If
                    ' ERR.NUMBER 75 GIVER "Path/File access error"
                End If
            End If
        End If
        End If
        MM_1 = MM_1 + A22 ' UCase(A22)
        If Err.Number > 0 Then MM_1 = MM_1 + vbCrLf + Err.Description + vbCrLf + "COUNT 1 ERROR =" + str(COUNT_1_ERROR) + vbCrLf + "COUNT 2 ERROR =" + str(COUNT_2_ERROR): ERROR_SET = True
        MM_1 = MM_1 + vbCrLf
    Next
        
    For RR = 1 To ScanPath.ListView1.ListItems.Count - 1
        ' A11 ARE DIR
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        A22 = A11 + B11
        If FSO.FolderExists(A22) = True Then
            TYPE_INFO = "FOLDER"
        Else
            TYPE_INFO = "FILE"
        End If
        ' ----------------------------------------------------------
        ' ----------------------------
        ' SET THE FOLDER ____ 02 OF 02
        ' ----------------------------
        If TYPE_INFO = "FOLDER" Then
            On Error Resume Next
            Err.Clear
            If A44 <> UCase(A44) Then
                Name A44 As UCase(A44)
                XC = XC + 1
            End If
            COUNT_1_ERROR = Err.Number
            If Err.Number > 0 Then
                Set F2 = FSO.GetFolder(A44)
                If F2.Name <> UCase(A44) Then
                    Err.Clear
                    Name A22 As UCase(A22) + "#"
                    If FSO.FileExists(A22 + "#") = True Then
                        Err.Clear
                        Name A22 + "#" As UCase(A22)
                        If Err.Number > 0 Then
                            MsgBox "ERROR FOLDER NOT ABLE RENAME BACK FROM TEMP NAME " + vbCrLf + vbCrLf + A22 + "#" + vbCrLf + vbCrLf
                            Shell "EXPLORER """ + A22 + """", vbMaximizedFocus
                            COUNT_2_ERROR = Err.Number
                            MM_1 = MM_1 + A22 ' UCase(A22)
                            If Err.Number > 0 Then MM_1 = MM_1 + vbCrLf + Err.Description + vbCrLf + "COUNT 1 ERROR =" + str(COUNT_1_ERROR) + vbCrLf + "COUNT 2 ERROR =" + str(COUNT_2_ERROR): ERROR_SET = True
                            MM_1 = MM_1 + vbCrLf
                            Exit For
                        End If
                    End If
                    ' ERR.NUMBER 75 GIVER "Path/File access error"
                End If
            End If
        End If
        MM_1 = MM_1 + A22 ' UCase(A22)
        If Err.Number > 0 Then MM_1 = MM_1 + vbCrLf + Err.Description + vbCrLf + "COUNT 1 ERROR =" + str(COUNT_1_ERROR) + vbCrLf + "COUNT 2 ERROR =" + str(COUNT_2_ERROR): ERROR_SET = True
        MM_1 = MM_1 + vbCrLf
    Next
    
    If ERROR_SET = True Then
        MsgBox "ERROR ON " + vbCrLf + vbCrLf + MM_1
    Else
        ' Call SET_QUICK_MODE_RESULT
        If InStr(MNU_QUICK_MODE.Caption, I_2) Then
            MsgBox "DONE = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
        End If
    End If
    End
End Sub

Sub CAPITAL_ALL____PATH_OPERATION____TOP_DEPTH()
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    
    Dim TT
    Dim RR
    
    GATHER_FOLDER_AND_FILE_MODE_FOR_SCANPATH = True
    
    ' SCAN DO ON TEXTPATH CHANGE
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    If LABEL_SET(2).Caption = "NOT FOLDER GIVEN" Then
        If IsIDE Then
            ' ScanPath.txtPath.Text = "F:\RETEKESS M3\M\04 MY MUSIC\03 MP3 CD\##00 TOP SET\00 TOP"
            ' ScanPath.txtPath.Text = "D:\0 00 MUSIC ---\04 MY MUSIC\03 MP3 CD\##00 TOP SET\00 TOP"
        End If
    End If
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    If IsIDE Then
        ' ScanPath.txtPath.Text = "F:\RETEKESS M3\M\04 MY MUSIC\03 MP3 CD\##00 TOP SET\00 TOP"
        ScanPath.txtPath.Text = "D:\0 00 MUSIC ---\04 MY MUSIC\03 MP3 CD\##00 TOP SET\00 TOP"
    End If
    
    BADDAR = ""
    For RR = ScanPath.ListView1.ListItems.Count To 1 Step -1
        ' A11 ARE DIR
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        SET_REMOVE = False
        If InStr(BADDAR, A11 + "--#") > 0 Then
            SET_REMOVE = True
        End If
        If InStr(A11, "_gsdata_") > 0 Then
            SET_REMOVE = True
        End If
        If SET_REMOVE = True Then
            ScanPath.ListView1.ListItems.Remove (RR)
        End If
        If InStr(BADDAR, A11 + "--#") = 0 Then
            BADDAR = BADDA + A11 + "--#"
        End If
    Next
    
    If ScanPath.ListView1.ListItems.Count > 500 Then
        MsgBox "LOOK LIKE GOT TOO MANY FILE TO DO OVER 500" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
        End
    End If
    If ScanPath.ListView1.ListItems.Count > 0 Then
        Call SET_QUICK_MODE_RESULT
        If InStr(MNU_QUICK_MODE.Caption, I_3) Then
            MsgBox "FOLDER COUNTER" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
        End If
    Else
        MsgBox "NONE COUNT OF FOLDER HERE " + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
        End
    End If

    For RR = 1 To ScanPath.ListView1.ListItems.Count - 1
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        A22 = A11 + B11
        If FSO.FolderExists(A22) = True Then
            TYPE_INFO = "FOLDER"
        Else
            TYPE_INFO = "FILE"
        End If
        ' ----------------------------------------------------------
        ' ----------------------------
        ' SET THE FOLDER ____
        ' ----------------------------
        If TYPE_INFO = "FOLDER" Then
        If InStr(A11, "_gsdata_") = 0 Or 2 = 2 Then
            On Error Resume Next
            Err.Clear
            If A22 <> UCase(A22) Then
                Name A22 As UCase(A22)
                XC = XC + 1
            End If
            COUNT_1_ERROR = Err.Number
            If Err.Number > 0 Then
                If A22 <> UCase(A22) Then
                    Err.Clear
                    Name A22 As UCase(A22) + "#"
                    If FSO.FolderExists(A22 + "#") = True Then
                        Err.Clear
                        Name A22 + "#" As UCase(A22)
                        If Err.Number > 0 Then
                            MsgBox "ERROR FOLDER NOT ABLE RENAME BACK FROM TEMP NAME " + vbCrLf + vbCrLf + A22 + "#" + vbCrLf + vbCrLf
                            Shell "EXPLORER """ + A22 + """", vbMaximizedFocus
                            COUNT_2_ERROR = Err.Number
                            MM_1 = MM_1 + A22 ' UCase(A22)
                            If Err.Number > 0 Then MM_1 = MM_1 + vbCrLf + Err.Description + vbCrLf + "COUNT 1 ERROR =" + str(COUNT_1_ERROR) + vbCrLf + "COUNT 2 ERROR =" + str(COUNT_2_ERROR): ERROR_SET = True
                            MM_1 = MM_1 + vbCrLf
                            Exit For
                        End If
                        
                    End If
                    ' ERR.NUMBER 75 GIVER "Path/File access error"
                End If
            End If
            MM_1 = MM_1 + A22 ' UCase(A22)
            If Err.Number > 0 Then MM_1 = MM_1 + vbCrLf + Err.Description + vbCrLf + "COUNT 1 ERROR =" + str(COUNT_1_ERROR) + vbCrLf + "COUNT 2 ERROR =" + str(COUNT_2_ERROR): ERROR_SET = True
            MM_1 = MM_1 + vbCrLf
        End If
        End If
    Next
    If ERROR_SET = True Then
        MsgBox "ERROR ON " + vbCrLf + vbCrLf + MM_1
    Else
        ' Call SET_QUICK_MODE_RESULT
        If InStr(MNU_QUICK_MODE.Caption, I_2) Then
            MsgBox "DONE = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
        End If
    End If
    End
End Sub

Sub CAPITAL_ALL____PATH_OPERATION__001_DEPTH()
    Set FSO = CreateObject("Scripting.FileSystemObject")

    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        Dir1.Path = LABEL_SET(2).Caption
    End If
    
    If IsIDE Then
        ' Dir1.Path = "F:\RETEKESS--M\04 MY MUSIC\02 CD\04 EDDIE'S TUNES"
    End If
    
    If Len(Dir1.Path) < 4 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
'    If ScanPath.ListView1.ListItems.Count > 500 Then
'        MsgBox "LOOK LIKE GOT TOO MANY FILE TO DO OVER 500" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
'        End
'    End If
    If Dir1.ListCount > 0 Then
        Call SET_QUICK_MODE_RESULT
        If InStr(MNU_QUICK_MODE.Caption, I_2) Then
            MsgBox "FOLDER COUNTER" + vbCrLf + vbCrLf + Trim(str(Dir1.ListCount)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + Dir1.Path
        End If
    Else
        MsgBox "NONE COUNT OF FOLDER HERE " + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + Dir1.Path
        End
    End If

    Dim TT
    Dim RR
    For RR = 0 To Dir1.ListCount - 1
        A11 = Dir1.List(RR)
        ' ----------------------------------------------------------
        If InStr(A11, "_gsdata_") = 0 Then
            On Error Resume Next
            Err.Clear
            Name A11 As UCase(A11)
            XC = XC + 1
            MM_1 = MM_1 + UCase(A11)
            If Err.Number > 0 Then MM_1 = MM_1 + Err.Description + vbCrLf: ERROR_SET = True
            MM_1 = MM_1 + vbCrLf
        End If
    Next
    If ERROR_SET = True Then
        MsgBox "ERROR ON " + vbCrLf + vbCrLf + MM_1
    Else
        Call SET_QUICK_MODE_RESULT
        If InStr(MNU_QUICK_MODE.Caption, I_2) Then
            MsgBox "DONE = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
        End If
        End If
    End

End Sub

Sub CAPITAL_ALL____FILE_OPERATION____TOP_DEPTH()
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    
    Dim TT
    Dim RR
    
    If IsIDE Then
        LABEL_SET(2).Caption = "D:\0 00 MUSIC ---\04 MY MUSIC\00 TOP\#02 - CD-CLASSIC\04 CLASSIC FM 00 RADIO"
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    
    
    ' SCAN DO ON LABEL_SET.CAPTION TEXTPATH CHANGE
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    
    If LABEL_SET(2).Caption = "NOT FOLDER GIVEN" Then
        If IsIDE Then
            ' ScanPath.txtPath.Text = "F:\RETEKESS M3\M\04 MY MUSIC\03 MP3 CD\##00 TOP SET\00 TOP"
            ' ScanPath.txtPath.Text = "D:\0 00 MUSIC ---\04 MY MUSIC\03 MP3 CD\##00 TOP SET\00 TOP"
        End If
    End If
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    If IsIDE Then
        ' ScanPath.txtPath.Text = "F:\RETEKESS M3\M\04 MY MUSIC\03 MP3 CD\##00 TOP SET\00 TOP"
        ' ScanPath.txtPath.Text = "D:\0 00 MUSIC ---\04 MY MUSIC\03 MP3 CD\##00 TOP SET\00 TOP"
    End If
    
    For RR = ScanPath.ListView1.ListItems.Count To 1 Step -1
        ' A11 ARE DIR
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        SET_REMOVE = False
        If InStr(A11, "_gsdata_") > 0 Then SET_REMOVE = True
        ' If InStr(B11, "_gsdata_") > 0 Then SET_REMOVE = True
        If SET_REMOVE = True Then
            ScanPath.ListView1.ListItems.Remove (RR)
        End If
    Next
    
    If ScanPath.ListView1.ListItems.Count > 500 Then
        For RR = ScanPath.ListView1.ListItems.Count To 1 Step -1
            ' A11 ARE DIR
            A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
            B11 = ScanPath.ListView1.ListItems.Item(RR)
            D11 = D11 + A11 + B11 + vbCrLf
        Next
        MsgBox "LOOK LIKE GOT TOO MANY FILE TO DO OVER 500" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + vbCrLf + vbCrLf + D11
        End
    End If
    If ScanPath.ListView1.ListItems.Count > 0 Then
        Call SET_QUICK_MODE_RESULT
        If InStr(MNU_QUICK_MODE.Caption, I_3) Then
            MsgBox "FOLDER COUNTER" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
        End If
    Else
        MsgBox "NONE COUNT OF FOLDER HERE " + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
        End
    End If

    For RR = 1 To ScanPath.ListView1.ListItems.Count - 1
        ' A11 ARE DIR
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        A22 = A11 + B11
        ' ----------------------------------------------------------
        On Error Resume Next
        Err.Clear
        If B11 <> UCase(B11) Then
            Name A22 As UCase(A22)
            XC = XC + 1
        End If
        COUNT_1_ERROR = Err.Number
        If Err.Number > 0 Then
            If B11 <> UCase(B11) Then
                Err.Clear
                Name A22 As UCase(A22) + "#"
                If FSO.FileExists(A22 + "#") = True Then
                    Err.Clear
                    Name A22 + "#" As UCase(A22)
                    If Err.Number > 0 Then
                        MsgBox "ERROR FOLDER NOT ABLE RENAME BACK FROM TEMP NAME " + vbCrLf + vbCrLf + A22 + "#" + vbCrLf + vbCrLf
                        Shell "EXPLORER """ + A22 + """", vbMaximizedFocus
                        COUNT_2_ERROR = Err.Number
                        MM_1 = MM_1 + A22 ' UCase(A22)
                        If Err.Number > 0 Then MM_1 = MM_1 + vbCrLf + Err.Description + vbCrLf + "COUNT 1 ERROR =" + str(COUNT_1_ERROR) + vbCrLf + "COUNT 2 ERROR =" + str(COUNT_2_ERROR): ERROR_SET = True
                        MM_1 = MM_1 + vbCrLf
                        Exit For
                    End If
                    
                End If
                ' ERR.NUMBER 75 GIVER "Path/File access error"
            End If
        End If
        MM_1 = MM_1 + A22 ' UCase(A22)
        If Err.Number > 0 Then MM_1 = MM_1 + vbCrLf + Err.Description + vbCrLf + "COUNT 1 ERROR =" + str(COUNT_1_ERROR) + vbCrLf + "COUNT 2 ERROR =" + str(COUNT_2_ERROR): ERROR_SET = True
        MM_1 = MM_1 + vbCrLf
    Next
    If ERROR_SET = True Then
        MsgBox "ERROR ON " + vbCrLf + vbCrLf + MM_1
    Else
        ' Call SET_QUICK_MODE_RESULT
        If InStr(MNU_QUICK_MODE.Caption, I_2) Then
            MsgBox "DONE = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
        End If
    End If
    End
End Sub

Sub CAPITAL_ALL____FILE_OPERATION__001_DEPTH()
    Set FSO = CreateObject("Scripting.FileSystemObject")

    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        File1.Path = LABEL_SET(2).Caption
    End If
    
    If IsIDE Then
        File1.Path = "D:\0 00 VIDEO 01\03 VIDEO PIRATE\1999 ACCIDENT"
    End If
    
    If Len(File1.Path) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
'    If ScanPath.ListView1.ListItems.Count > 500 Then
'        MsgBox "LOOK LIKE GOT TOO MANY FILE TO DO OVER 500" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
'        End
'    End If
    
    If File1.ListCount > 0 Then
        MsgBox "FILE COUNTER" + vbCrLf + vbCrLf + Trim(str(File1.ListCount)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + File1.Path
    Else
        MsgBox "NONE COUNT OF FILE HERE " + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + File1.Path
        End
    End If

    Dim TT
    Dim RR
    For RR = 0 To File1.ListCount - 1
        A11 = File1.Path + "\" + File1.List(RR)
        ' ----------------------------------------------------------
        If InStr(A11, "_gsdata_") = 0 Then
            On Error Resume Next
            Err.Clear
            Name A11 As UCase(A11)
            XC = XC + 1
            MM_1 = MM_1 + UCase(A11)
            If Err.Number > 0 Then MM_1 = MM_1 + Err.Description + vbCrLf: ERROR_SET = True
            MM_1 = MM_1 + vbCrLf
        End If
    Next
    If ERROR_SET = True Then
        MsgBox "ERROR ON " + vbCrLf + vbCrLf + MM_1
    Else
        Call SET_QUICK_MODE_RESULT
        ' If InStr(MNU_QUICK_MODE.Caption, I_2) Then
            MsgBox "DONE = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
        ' End If
        End If
    End


End Sub

Sub OPEN_ALL_URL_PATH_OPERATION_001_DEPTH()
    Set FSO = CreateObject("Scripting.FileSystemObject")

    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        File1.Path = LABEL_SET(2).Caption
    End If
    
    If IsIDE Then
        Dir1.Path = "C:\SCRIPTER\MP3\WINDOWS 98SE YOUTUBE"
    End If
    
    If Len(File1.Path) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    If File1.ListCount > 0 Then
        MsgBox "FILE COUNTER" + vbCrLf + vbCrLf + Trim(str(File1.ListCount)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + File1.Path
    Else
        MsgBox "NONE COUNT OF FILE HERE " + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + File1.Path
        End
    End If

    Dim TT
    Dim RR
    For RR = 0 To File1.ListCount - 1
        A11 = File1.Path + "\" + File1.List(RR)
        ' ----------------------------------------------------------
        If InStr(A11, "_gsdata_") = 0 Then
            On Error Resume Next
            Err.Clear
            If InStr(UCase(A11), ".URL") > 0 Then
            
                Dim WSHShell
                Set WSHShell = CreateObject("WScript.Shell")
                    WSHShell.RUN """" + A11 + """", ShowWindow_2, DontWaitUntilFinished
                Set WSHShell = Nothing
            End If
            XC = XC + 1
            MM_1 = MM_1 + UCase(A11)
            If Err.Number > 0 Then MM_1 = MM_1 + Err.Description + vbCrLf: ERROR_SET = True
            MM_1 = MM_1 + vbCrLf
        End If
    Next
    If ERROR_SET = True Then
        MsgBox "ERROR ON " + vbCrLf + vbCrLf + MM_1
    Else
        Call SET_QUICK_MODE_RESULT
        ' If InStr(MNU_QUICK_MODE.Caption, I_2) Then
            MsgBox "DONE = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
        ' End If
        End If
    End

End Sub




Sub DATE_CONVERTOR___MMM_D__YYYY_H_MM_AM____TO_YYYY_MM_DD__HH_MM_DD_FOR_SCREENCASTIFY_SINGLE()

    A11 = LABEL_SET(3).Caption
    ' A11 = "\\4-asus-gl522vw\4_asus_gl522vw_10_1_samsung_1tb\DSC\2015+SONY\2020 CyberShot HX60V\MP4\2020_01_20 Jan_Mon 12_25_38__MAH01294_HOVE GRAND AVENUE DEMOLITION__.MP4"
    D11 = A11
    If A11 <> "NOT FILE GIVEN" Then
        A11 = Mid(D11, 1, InStrRev(A11, "\"))
        B11 = Mid(D11, InStrRev(A11, "\") + 1)
    End If
    
    If Len(D11) < 5 Or D11 = "NOT FILE GIVEN" Then
        MsgBox "FULL PATH AND FILE NOT GOOD -- EXIT" + vbCrLf + vbCrLf + D11
        End
    End If
    
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
        
    ' D11 = "C:\DOWNLOADS\Apr 28, 2020 9_57 AM.mp4"
    
    EXT_STR = UCase(Mid(B11, Len(B11) - 2))
    If InStr("MP4 TXT", EXT_STR) And InStr(A11, "_gsdata_") = 0 Then
        
        Set F = FSO.GetFILE(A11 + B11)
        DT1 = 0
        DT1 = F.datelastmodified
        DT1 = 0
        'Aug 5, 2019 9 57 Am-4
        DATE_STRING = ""
        If InStrRev(B11, "m-") > 0 Then
            THREE_BACK_SPACE = InStrRev(B11, " ", Len(B11))
            For r = 1 To 2
                THREE_BACK_SPACE = InStrRev(B11, " ", THREE_BACK_SPACE - 1)
            Next
            DATE_STRING_1 = Mid(B11, 1, THREE_BACK_SPACE)
            DATE_STRING_2 = Mid(B11, THREE_BACK_SPACE, InStrRev(B11, "m-") - THREE_BACK_SPACE + 1)
            DATE_STRING_2 = Trim(DATE_STRING_2)
        End If
        DATE_STRING_2 = Replace(DATE_STRING_2, " ", ":", , 1)
        If DATE_STRING_1 <> "" Then
        If IsDate(DATE_STRING_1) Then
            DT1 = DateValue(DATE_STRING_1)
        End If
        End If
        
        If DATE_STRING_2 <> "" Then
        If IsDate(DATE_STRING_2) Then
            DT1 = DT1 + TimeValue(DATE_STRING_2)
        End If
        End If
        
        GET_MMM_ = Mid(B11, 1, 3)
        GET_DD__ = Mid(B11, 5, 2)
        GET_YYYY = Mid(B11, 9, 4)
        GET_TIME = Mid(B11, 14, InStrRev(B11, "M.") - 13)
        GET_TIME = Replace(GET_TIME, "_", ":")

        GET_DATE_VALUE = DateValue(GET_YYYY + " " + GET_MMM_ + " " + GET_DD__) + TimeValue(GET_TIME)
        DT1 = GET_DATE_VALUE
        DT1_STRING = Format(DT1, "YYYY-MM-DD--HH-MM-SS")
        If InStrRev(B11, "-") > 0 Then
            WHAT_AT_AFTER_MARK_DASH = Mid(B11, InStrRev(B11, "-"))
            Else
            WHAT_AT_AFTER_MARK_DASH = "." + EXT_STR
        End If
        DT1_STRING = DT1_STRING + WHAT_AT_AFTER_MARK_DASH
        
        If DT1 > 0 Then
            On Error Resume Next
            Err.Clear
            Name A11 + "\" + B11 As A11 + "\" + DT1_STRING
            If Err.Number > 0 Then
            
                DETAIL_VAR = "RENAME FROM --" + vbCrLf + vbCrLf + A11 + "\" + B11
                DETAIL_VAR = "MOVE TO   --" + vbCrLf + vbCrLf + A11 + "\" + DT1_STRING
                MsgBox "WAS A ERROR RENAME FILE REQUEST" + vbCrLf + vbCrLf + "Err.Description = " + Err.Description + vbCrLf + vbCrLf + DETAIL_VAR, vbMsgBoxSetForeground
            
            End If
            TT = SetFileDateTime(A11 + "\" + DT1_STRING, DT1)
            
            XC = XC + 1
            MM_1 = MM_1 + B11 + vbCrLf
        End If
    End If
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End

End Sub


Sub ONE_FOLDER_AS_OTHER_SAME_____HARDCODER____BATCH()

    Set FSO = CreateObject("Scripting.FileSystemObject")

    DR_1 = "C:\DD\ABBYWINTERS.COM\"
    DR_2 = "C:\DD\"
    ScanPath.chkSubFolders = vbUnchecked
    ScanPath.cboMask.Text = "*.MP4;*.WMV"
    
    ' SCAN DO ON TEXTPATH CHANGE
    ScanPath.txtPath.Text = DR_1
    
    'If ScanPath.ListView1.ListItems.Count > 0 Then
        'MsgBox "FILE COUNTER -- BEFORE FILTER ON" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
    'End If
    
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        B11_NOT_EXT = Mid(B11, 1, Len(B11) - 4)
        ' WANT PROCESS WITH -- MP4 WMV -- AND THEN PUT THEM HERE
        ' ----------------------------------------------------------
        If InStr("MP4 WMV", EXT_STR) And InStr(A11, "_gsdata_") = 0 Then
            
            Set F = FSO.GetFILE(A11 + B11)
            DT1 = F.datelastmodified

            On Error Resume Next
            Err.Clear
            
            OUT_FILE_MP4 = DR_2 + B11_NOT_EXT + ".MP4"
            OUT_FILE_WMV = DR_2 + B11_NOT_EXT + ".WMV"
            If Dir(OUT_FILE_MP4) <> "" Then
                TT = SetFileDateTime(OUT_FILE_MP4, DT1)
                Name OUT_FILE_MP4 As OUT_FILE_MP4
                XC = XC + 1
                MM_1 = MM_1 + B11 + vbCrLf
            End If
            If Err.Number > 0 Then MsgBox "EEROR WITH " + vbCrLf + vbCrLf + OUT_FILE_MP4
            If Dir(OUT_FILE_WMV) <> "" Then
                TT = SetFileDateTime(OUT_FILE_WMV, DT1)
                Name OUT_FILE_WMV As OUT_FILE_WMV
                XC = XC + 1
                MM_1 = MM_1 + B11 + vbCrLf
            End If
            If Err.Number > 0 Then MsgBox "EEROR WITH " + vbCrLf + vbCrLf + OUT_FILE_WMV
        End If
    Next
    
    Call SET_QUICK_MODE_RESULT
    ' If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    ' End If
    
    End


End Sub

Sub MAKE_FOLDER_YYYY_MM_DD_OF_FILE_AND_MOVE_THERE_____BATCH_IT()

    ScanPath.chkSubFolders = vbUnchecked
    ScanPath.cboMask.Text = "*.MP4;*.MP3;*.WAV;*.AVI;*.JPG"
    
    ' SCAN DO ON TEXTPATH CHANGE
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    If IsIDE Then
        ' ScanPath.txtPath.Text = "F:\DSC--2018+CCTV_HIKVISION\2018+CCTV_HIKVISION_DS-7204H\2020_CCTV_HIKVISION_DS-7204H"
        ScanPath.txtPath.Text = "D:\0 00 MOBILE-2\RETEKESS\RETEKESS_RECORD 2020\MRECORD"
    End If
    
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
'    If ScanPath.ListView1.ListItems.Count > 500 Then
'        MsgBox "LOOK LIKE GOT TOO MANY FILE TO DO OVER 500" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
'        End
'    End If
    
    'If ScanPath.ListView1.ListItems.Count > 0 Then
        MsgBox "FILE COUNTER -- BEFORE FILTER ON" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
    'End If
    
    
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    I_RESULT = 0
    
    If InStr(A11, "\0 00 MOBILE-2\8GB RECORD AUDIO BLUE\") > 0 Then
        I_RESULT = vbNo
    End If
    
    If InStr(A11, "\DSC\") > 0 Then
        I_RESULT = vbYes
    End If
    
    If InStr(A11, "\VI_ DSC ME\") > 0 Then
        I_RESULT = vbYes
    End If
    
    If I_RESULT = 0 Then
        I_RESULT = MsgBox("SPACE OR DASH BETWEEN __ YES SPACE __ NOT DASH", vbYesNo + vbMsgBoxSetForeground)
    End If
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        ' WANT PROCESS WITH -- MP4 TXT MP3 -- AND THEN PUT THEM HERE
        ' ----------------------------------------------------------
        ' If InStr("MP4 TXT MP3 AVI", EXT_STR) THEN ' WHY -- IS ABOVE A BIT
        
        If InStr(A11, "_gsdata_") = 0 Then
            
            Set F = FSO.GetFILE(A11 + B11)
            DT1 = F.datelastmodified

            DATE_STRING = Format(DT1, "YYYY-MM-DD")
            If I_RESULT = vbYes Then
                DATE_STRING = Replace(DATE_STRING, "-", " ")
            End If
            
            OUT_FOLDER__ = A11 + "\" + DATE_STRING
            OUT_FOLDER_AND_FILENAME = A11 + "\" + DATE_STRING + "\" + B11
            
'            If InStr(LABEL_SET(3).Caption, "REC") > 0 And InStr(LABEL_SET(3).Caption, ".WAV") > 0 Then
'                OUT_FOLDER_AND_FILENAME = LABEL_SET(2).Caption + "\RECORD " + Format(DT1, "YYYY-MM-DD")
'            End If
            
            If Dir(OUT_FOLDER__, vbDirectory) = "" Then
                CreateFolderTree OUT_FOLDER__
            End If
            FILE_NAME = A11 + B11
            On Error Resume Next
            Err.Clear
            FSO.MoveFile FILE_NAME, OUT_FOLDER_AND_FILENAME
            If Err.Number > 0 Then
            
                DETAIL_VAR = "MOVE FROM --" + vbCrLf + vbCrLf + FILE_NAME
                DETAIL_VAR = "MOVE TO   --" + vbCrLf + vbCrLf + OUT_FOLDER_AND_FILENAME
                MsgBox "WAS A ERROR MOVE FILE REQUESTED" + vbCrLf + vbCrLf + "Err.Description = " + Err.Description + vbCrLf + vbCrLf + DETAIL_VAR, vbMsgBoxSetForeground
            
            End If
            
            XC = XC + 1
            MM_1 = MM_1 + B11 + vbCrLf
        End If
    Next
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End


End Sub


Sub SET_DATE_OF_FILENAME_YYYY_MM_DD_MMM_DDD_HH_MM_SS__MA()

    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.MP4"
    
    ' SCAN DO ON TEXTPATH CHANGE
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    If IsIDE Then
        ScanPath.txtPath.Text = "D:\DD"
    End If
    
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    If ScanPath.ListView1.ListItems.Count > 500 Then
        MsgBox "LOOK LIKE GOT TOO MANY FILE TO DO OVER 500" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
        End
    End If
    
    If ScanPath.ListView1.ListItems.Count > 0 Then
        MsgBox "FILE COUNTER -- BEFORE FILTER ON" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
    End If
    
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        If InStr("MP4 TXT", EXT_STR) And InStr(A11, "_gsdata_") = 0 Then
            DATE_FILENAME_D_1 = Mid(B11, 1, 10)
            DATE_FILENAME_T_2 = Mid(B11, 20, 8)
            i2 = DATE_FILENAME_D_1
            i4 = DATE_FILENAME_T_2
            DATE_FILENAME_D_1 = Mid(i2, 1, 4) + "/" + Mid(i2, 6, 2) + "/" + Mid(i2, 9, 2)
            DATE_FILENAME_T_2 = Mid(i4, 1, 2) + ":" + Mid(i4, 4, 2) + ":" + Mid(i4, 7, 2)
            On Error Resume Next
            DATE_FILENAME_SUCCESS = DateValue(DATE_FILENAME_D_1) + TimeValue(DATE_FILENAME_T_2)
            If Err.Number > 0 Then
                MsgBox "ERROR HERE ARUGEMENTS INVAILD" + vbCrLf + vbCrLf + B11 + vbCrLf + vbCrLf + DATE_FILENAME_D_1 + vbCrLf + vbCrLf + DATE_FILENAME_T_2
                End
            End If
            On Error GoTo 0
            DT4 = DATE_FILENAME_SUCCESS
            If IsDate(DT4) = False Then
                ' FALSE DATE
                ' ----------
                MsgBox "DATE HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
                End
            End If
            
            If DT4 = 0 Then
                MsgBox "NAUGHT DATE FINDER -- EXIT"
                End
            End If
        
            TT = SetFileDateTime(A11 + B11, DT4)
            
            EXT_ORG = Mid(B11, Len(B11) - 2)
            EXT_STR = UCase(Mid(B11, Len(B11) - 2))
            If EXT_ORG <> EXT_STR Then
                'RENAME EXTENSION CAPITAL
                FILE_RENAME1 = Mid(B11, 1, Len(B11) - 4) + "." + EXT_STR
                    Name A11 + B11 As A11 + FILE_RENAME1
            End If
            
            XC = XC + 1
            MM_1 = MM_1 + B11 + vbCrLf
        End If
    Next
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End
End Sub


Sub SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS__JPG_BATCH()

    ' ---------------------------------------------------------------------
    ' ANOTHER SUB ROUTINE
    ' LEECH AND MODIFY FROM BATCH CODE SET
    ' Wed 00:05:20 Am_19 November 2025
    ' ---------------------------------------------------------------------


    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.MP4;*.MP3;*.WAV;*.AVI;*.JPG"
    
    ' SCAN DO ON TEXTPATH CHANGE
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        ScanPath.txtPath.Text = LABEL_SET(2).Caption
    End If
    
    If IsIDE Then
        ' ScanPath.txtPath.Text = "F:\DSC--2018+CCTV_HIKVISION\2018+CCTV_HIKVISION_DS-7204H\2020_CCTV_HIKVISION_DS-7204H"
        ' ScanPath.txtPath.Text = "D:\DSC\2015+SONY\2020 CyberShot HX60V\JPG\2020 06 23"
    End If
    
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    
    D11 = ScanPath.txtPath.Text
    If A11 <> "NOT FILE GIVEN" Then
        A11 = Mid(D11, 1, InStrRev(A11, "\"))
        B11 = Mid(D11, InStrRev(A11, "\") + 1)
    End If
    
    If Len(D11) < 5 Or D11 = "NOT FILE GIVEN" Then
        MsgBox "FULL PATH AND FILE NOT GOOD -- EXIT" + vbCrLf + vbCrLf + D11
        End
    End If
    
    
'    If ScanPath.ListView1.ListItems.Count > 500 Then
'        MsgBox "LOOK LIKE GOT TOO MANY FILE TO DO OVER 500" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
'        End
'    End If
    
    'If ScanPath.ListView1.ListItems.Count > 0 Then
        MsgBox "FILE COUNTER -- BEFORE FILTER ON" + vbCrLf + vbCrLf + Trim(str(ScanPath.ListView1.ListItems.Count)) + vbCrLf + vbCrLf + "CHECK IT OUT" + vbCrLf + vbCrLf + "PATH " + vbCrLf + vbCrLf + ScanPath.txtPath.Text
    'End If
    
    
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    I_RESULT = 0
    
    If InStr(A11, "\0 00 MOBILE-2\8GB RECORD AUDIO BLUE\") > 0 Then
        I_RESULT = vbNo
    End If
    
    If InStr(A11, "\DSC\") > 0 Then
        I_RESULT = vbYes
    End If
    
    If InStr(A11, "\VI_ DSC ME\") > 0 Then
        I_RESULT = vbYes
    End If
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        ' WANT PROCESS WITH -- MP4 TXT MP3 -- AND THEN PUT THEM HERE
        ' ----------------------------------------------------------
        ' If InStr("MP4 TXT MP3 AVI", EXT_STR) THEN ' WHY -- IS ABOVE A BIT
        
        If InStr(A11, "_gsdata_") = 0 Then
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
            If InStr("JPG TXT", EXT_STR) Then
                DATE_FILENAME_D_1 = Mid(B11, 1, 10)
                DATE_FILENAME_T_2 = Mid(B11, 12, 8)
                i2 = DATE_FILENAME_D_1
                i4 = DATE_FILENAME_T_2
                DATE_FILENAME_D_1 = Mid(i2, 1, 4) + "/" + Mid(i2, 6, 2) + "/" + Mid(i2, 9, 2)
                DATE_FILENAME_T_2 = Mid(i4, 1, 2) + ":" + Mid(i4, 4, 2) + ":" + Mid(i4, 7, 2)
                
                DATE_FILENAME_SUCCESS = DateValue(DATE_FILENAME_D_1) + TimeValue(DATE_FILENAME_T_2)
                
                DT4 = DATE_FILENAME_SUCCESS
                If IsDate(DT4) = False Then
                    ' FALSE DATE
                    ' ----------
                    MsgBox "DATE HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
                    End
                End If
                
                If DT4 = 0 Then
                    MsgBox "NAUGHT DATE FINDER -- EXIT"
                    End
                End If
            
                TT = SetFileDateTime(A11 + B11, DT4)
                
                EXT_ORG = Mid(B11, Len(B11) - 2)
                EXT_STR = UCase(Mid(B11, Len(B11) - 2))
                If EXT_ORG <> EXT_STR Then
                    'RENAME EXTENSION CAPITAL
                    FILE_RENAME1 = Mid(B11, 1, Len(B11) - 4) + "." + EXT_STR
                    Name A11 + B11 As A11 + FILE_RENAME1
                End If
                
                XC = XC + 1
                MM_1 = MM_1 + B11 + vbCrLf
            End If
        End If
    Next
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End
            
            
            
            
            

End Sub


Sub SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS__JPG_SINGLE()
    ' ---------------------------------------------------------------------
    ' ANOTHER SUB ROUTINE
    ' LEECH AND MODIFY FROM BATCH CODE SET
    ' Mon 30-Mar-2020 22:51:11
    ' ---------------------------------------------------------------------
    ' LABEL_SET(3).Caption ' -- FULL_PATH_AND_FILENAME
    ' ---------------------------------------------------------------------
    A11 = LABEL_SET(3).Caption
    ' A11 = "D:\DSC\2015+SONY\2020 CyberShot HX60V\JPG\2020 03 30\2020-03-30 13-05-50 - SONY DSC-HX60V - DSC02044 - EDITOR.JPG"
    D11 = A11
    If A11 <> "NOT FILE GIVEN" Then
        A11 = Mid(D11, 1, InStrRev(A11, "\"))
        B11 = Mid(D11, InStrRev(A11, "\") + 1)
    End If
    
    If Len(D11) < 5 Or D11 = "NOT FILE GIVEN" Then
        MsgBox "FULL PATH AND FILE NOT GOOD -- EXIT" + vbCrLf + vbCrLf + D11
        End
    End If
    
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    EXT_STR = UCase(Mid(B11, Len(B11) - 2))
    If InStr("JPG TXT", EXT_STR) Then
        DATE_FILENAME_D_1 = Mid(B11, 1, 10)
        DATE_FILENAME_T_2 = Mid(B11, 12, 8)
        i2 = DATE_FILENAME_D_1
        i4 = DATE_FILENAME_T_2
        DATE_FILENAME_D_1 = Mid(i2, 1, 4) + "/" + Mid(i2, 6, 2) + "/" + Mid(i2, 9, 2)
        DATE_FILENAME_T_2 = Mid(i4, 1, 2) + ":" + Mid(i4, 4, 2) + ":" + Mid(i4, 7, 2)
        
        DATE_FILENAME_SUCCESS = DateValue(DATE_FILENAME_D_1) + TimeValue(DATE_FILENAME_T_2)
        
        DT4 = DATE_FILENAME_SUCCESS
        If IsDate(DT4) = False Then
            ' FALSE DATE
            ' ----------
            MsgBox "DATE HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
            End
        End If
        
        If DT4 = 0 Then
            MsgBox "NAUGHT DATE FINDER -- EXIT"
            End
        End If
    
        TT = SetFileDateTime(A11 + B11, DT4)
        
        EXT_ORG = Mid(B11, Len(B11) - 2)
        EXT_STR = UCase(Mid(B11, Len(B11) - 2))
        If EXT_ORG <> EXT_STR Then
            'RENAME EXTENSION CAPITAL
            FILE_RENAME1 = Mid(B11, 1, Len(B11) - 4) + "." + EXT_STR
            Name A11 + B11 As A11 + FILE_RENAME1
        End If
        
        XC = XC + 1
        MM_1 = MM_1 + B11 + vbCrLf
    End If
    
    Call SET_QUICK_MODE_RESULT
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End

End Sub


Sub SET_DATE_OF_FILENAME_YYYY_MM_DD_HH_MM_SS_DDD_NOKIA_AH()

    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    
    ScanPath.txtPath.Text = LABEL_SET(2).Caption
'   ScanPath.txtPath.Text = "D:\VI_ DSC ME\2010+NOKIA\2017 NOKIA E72 _ 007 JULY_ MP4____x003 _ Mill View Room"
    
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    Dim RR
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        
        DATE_FILENAME_D_1 = Mid(B11, 1, 10)
        DATE_FILENAME_T_2 = Mid(B11, 12, 8)
        i2 = DATE_FILENAME_D_1
        i4 = DATE_FILENAME_T_2
        DATE_FILENAME_D_1 = Mid(i2, 1, 4) + "/" + Mid(i2, 6, 2) + "/" + Mid(i2, 9, 2)
        DATE_FILENAME_T_2 = Mid(i4, 1, 2) + ":" + Mid(i4, 4, 2) + ":" + Mid(i4, 7, 2)
        
        DATE_FILENAME_SUCCESS = DateValue(DATE_FILENAME_D_1) + TimeValue(DATE_FILENAME_T_2)
        
        DT4 = DATE_FILENAME_SUCCESS
        If IsDate(DT4) = False Then
            ' FALSE DATE
            ' ----------
            MsgBox "DATE HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
            End
        End If
        
        If DT4 = 0 Then
            MsgBox "NAUGHT DATE FINDER -- EXIT"
            End
        End If
    
        TT = SetFileDateTime(A11 + B11, DT4)
        XC = XC + 1
        MM_1 = MM_1 + B11 + vbCrLf
    Next
    
    ' SET QUICK MODE RESULT
    ' ---------------------
    On Error Resume Next
    FR1 = FreeFile
    Open App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT" For Input As #FR1
        Line Input #FR1, I_1
    Close FR1
    MNU_QUICK_MODE.Caption = I_1
    I_1 = "QUICK MODE -- ON -- NOT MESSENGER"
    I_2 = "QUICK MODE -- OFF -- GIVE MESSENGER"
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End
End Sub
    
Sub SET_ALL_DATE_FOLDER_TO_THE_TEXTFILE_HOLD_DATE_WITHIN_AH()
    
    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    
    ScanPath.txtPath.Text = LABEL_SET(2).Caption
'    ScanPath.txtPath.Text = "D:\VIDEO\NOT\X 00 NOT ME\00 Vid XXX\00 ROOT_02_(DATE-ALPHABETICAL)\1984 a11 -- NURSE -- 1984"
    
    If Len(ScanPath.txtPath.Text) < 5 Then
        MsgBox "PATH TO SHORT -- EXIT"
        End
    End If
    
    DV = 0
    If InStr(ScanPath.txtPath.Text, "\00 ROOT_02_(DATE-ALPHABETICAL)\") > 0 Then DV = 1
    If InStr(ScanPath.txtPath.Text, "\ABBYWINTERS.COM_DATE") > 0 Then DV = 1
    If DV = 1 Then
        MENU_INCREMENT_DATE_VAR = -1
        Call MNU_INCREMENT_DATE_SHOW
    End If
    
    
    
    'ScanPath.Show
    Dim DT1 As Date
    Dim DS2 As Date
    Dim DS4 As Date ' OLDER COMPARE
    Dim TT
    
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        
        Set F = FSO.GetFILE(A11 + B11)
        DT2 = F.DateCreated
        DT1 = F.datelastmodified
        
        If DT4 = 0 Then DT4 = DT1
        If DT1 < DT4 Then DT4 = DT1
        
        Set F = Nothing
    Next
    
    XX = Dir(ScanPath.txtPath.Text + "\# TEXT_DATER*.TXT")
    If XX = "" Then
        XX = Dir(ScanPath.txtPath.Text + "\#_TEXT_DATER*.TXT")
    End If
    If XX = "" Then
        XX = Dir(ScanPath.txtPath.Text + "\# TEXT DATER*.TXT")
    End If
    If XX = "" Then
        MsgBox "NONE (DATE WITHIN TEXT FILE) FOUND HERE" + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\# TEXT_DATER*.txt" + vbCrLf + vbCrLf + "-- EXIT"
        End
    End If
        
    FR1 = FreeFile
    Open ScanPath.txtPath.Text + "\" + XX For Input As #FR1
        Line Input #FR1, TIME_STRING
    Close FR1
    
    TIME_STRING = Replace(TIME_STRING, "\", "/")
    TIME_STRING = Replace(TIME_STRING, "_", ":")
    TIME_STRING = Replace(TIME_STRING, "-", ":")
    
    If InStr(ScanPath.txtPath.Text, "XXX") > 0 Then
        On Error Resume Next
        a = TimeValue(TIME_STRING)
        If Err.Number > 0 Then TIME_STRING = Replace(TIME_STRING, ":", "/")
        On Error GoTo 0
        If Val(TimeValue(TIME_STRING)) = 0 Then
            TIME_STRING = Replace(TIME_STRING, ":", "/")
            TIME_STRING = Format(DateValue(TIME_STRING), "YYYY/MM/DD") + " 10:00:00"
        End If
    End If
    DT4 = DateValue(TIME_STRING) + TimeValue(TIME_STRING)
    If IsDate(DT4) = False Then
        ' FALSE DATE
        ' ----------
        MsgBox "DATE FOUND WITHIN TEXT FILE FOUND HERE IS FALSE " + vbCrLf + vbCrLf + ScanPath.txtPath.Text + "\" + XX + vbCrLf + vbCrLf + "IS A FALSE ONE -- EXIT"
        End
    End If
    
    If DT4 = 0 Then
        MsgBox "NAUGHT DATE FOUND OVERALL IN FILE GATHER -- EXIT"
        End
    End If
    
    If TimeValue(TIME_STRING) = 0 Then
        ' ADD A BIT ON DATE FOR DAYLIGHT SAVING AND NTFS AT EXACT MIDNIGHT
        ' CLOUD SYSTEM REQUIRING A BIT MORE ALSO OR DAY BEFORE
        ' ----------------------------------------------------------------
        DT4 = DT4 + TimeSerial(11, 0, 0)
        ' ----------------------------------------------------------------
    End If
    MM_1 = Format(DT4, "DD-MM-YYYY  HH:MM:SS  DDDD")
    MM_1 = MM_1 + vbCrLf
    MM_1 = MM_1 + ScanPath.ListView1.ListItems.Item(1).SubItems(1) + vbCrLf
    MM_1 = MM_1 + vbCrLf
    
    COUNT_FILE = ScanPath.ListView1.ListItems.Count
    If MENU_INCREMENT_DATE_VAR = -1 Then
        If COUNT_FILE > 1 Then
            DT4 = DT4 + TimeSerial(0, 1, 0)
        End If
    End If
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        TT = SetFileDateTime(A11 + B11, DT4)
        XC = XC + 1
        MM_1 = MM_1 + B11 + vbCrLf
        FI = UCase(B11)
        FI = Mid(FI, InStrRev(FI, ".") + 1) + " "
        ' If InStr(" MPG MPEG MP4 JPG AVI ", FI) > 0 Then
            If MENU_INCREMENT_DATE_VAR = -1 Then
                DT4 = DT4 + TimeSerial(0, 1, 0)
            End If
        ' End If
    Next
    
    On Error Resume Next
    FR1 = FreeFile
    Open App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT" For Input As #FR1
        Line Input #FR1, I_1
    Close FR1
    I_1 = "QUICK MODE -- ON -- NOT MESSENGER"
    I_2 = "QUICK MODE -- OFF -- GIVE MESSENGER"
    MNU_QUICK_MODE.Caption = I_2
    If InStr(MNU_QUICK_MODE.Caption, I_2) Then
        MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + MM_1
    End If
    
    End
End Sub


Sub BATCH_MODIFIED_TO_CREATED_TIME()
    
On Error Resume Next
XX = FSO.FolderExists(W$)

Dim DateSet As Date

If XX = False Then
'time2$ = "2011-11-01 10:00:00"
'DateSet = DateValue(time2$) + TimeValue(time2$)
'DateSet = Now
'
'tt = SetFileDateTime(W$, DateSet)

'tt = LastModifiedToCurrent(W$)
End
End If

' If DIRW$ <> "" Then W$ = DIRW$

'MsgBox str(XX) + "--" + W$
'End

Start2:

Dim TPath$

ScanPath.chkSubFolders = vbChecked
ScanPath.cboMask.Text = "*.*"

'ScanPath.txtPath.Text = "E:\01 FAVS\#00 Palm\#000 New\2009-08 Aug\Unet\Mine\Word Example Usage"
'ScanPath.txtPath.Text = "E:\01 FAVS\Ministers\Images"
ScanPath.txtPath.Text = W$

If Len(W$) < 5 Then End

'ScanPath.Show
Dim DT1 As Date
'Dim Ds1 As Date
Dim DS2 As Date
'Dim DateSet As Date
For we = 1 To ScanPath.ListView1.ListItems.Count
    A11 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B11 = ScanPath.ListView1.ListItems.Item(we)
    
    Set F = FSO.GetFILE(A11 + B11)
    DT3 = F.DateCreated
    DT1 = F.datelastmodified
    
    DateSet = DT3
    
    Set F = Nothing
    
    TT = SetFileDateTime(A11 + B11, DateSet)
    XC = XC + 1
    
Next

MsgBox "Done = " + vbCrLf + str(XC) '+dd$
End


End Sub

Sub BATCH_CREATED_TO_MODIFIED_TIME()
    
On Error Resume Next
XX = FSO.FolderExists(W$)

Dim DateSet As Date

'If XX = False Then
''time2$ = "2011-11-01 10:00:00"
''DateSet = DateValue(time2$) + TimeValue(time2$)
''DateSet = Now
''tt = SetFileDateTime(W$, DateSet)
''tt = LastModifiedToCurrent(W$)
'End
'End If

Start2:

Dim TPath$

ScanPath.chkSubFolders = vbChecked
ScanPath.cboMask.Text = "*.*"

'ScanPath.txtPath.Text = "E:\01 FAVS\#00 Palm\#000 New\2009-08 Aug\Unet\Mine\Word Example Usage"
'ScanPath.txtPath.Text = "E:\01 FAVS\Ministers\Images"
ScanPath.txtPath.Text = W$

If Len(W$) < 5 Then End

'ScanPath.Show
Dim DT1 As Date
'Dim Ds1 As Date
Dim DS2 As Date
I_MSG = ""
For we = 1 To ScanPath.ListView1.ListItems.Count
    A11 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B11 = ScanPath.ListView1.ListItems.Item(we)
    
    Set F = FSO.GetFILE(A11 + B11)
    DT3 = F.DateCreated
    DT1 = F.datelastmodified
    Set F = Nothing
    
    TT = SetFileDateTime(A11 + B11, DT3)
    XC = XC + 1
    
    I_MSG = I_MSG + B11 + vbCrLf
    
Next

MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + I_MSG '+dd$
End

End Sub


Sub FILE_CREATED_TO_MODIFIED_TIME()
        
    'On Error Resume Next
    XX = FSO.FileExists(FULL_PATH_AND_FILENAME)
    
    Dim DateSet As Date
    
    Dim DT1 As Date
    Dim DT3 As Date
    I_MSG = ""
        
    Set F = FSO.GetFILE(FULL_PATH_AND_FILENAME)
    DT3 = F.DateCreated
    DT1 = F.datelastmodified
    Set F = Nothing
    
    TT = SetFileDateTime(FULL_PATH_AND_FILENAME, DT3)
    XC = XC + 1
    
    I_MSG = I_MSG + FULL_PATH_AND_FILENAME + vbCrLf
    
    
    MsgBox "Done = " + vbCrLf + str(XC) + vbCrLf + vbCrLf + I_MSG '+dd$
    End
    
End Sub



Sub CODE_RUN()

Exit Sub

If Mid(W$, 1, 1) = """" Then
    W$ = Mid(W$, 2): W$ = Mid(W$, 1, Len(W$) - 1)
End If

If Dir$(W$) <> "" Then
    DIRW$ = Mid$(W$, 1, InStrRev(W$, "\") - 1)
End If
    
If Mid$(W$, Len(W$), 1) <> "\" Then
    W$ = W$ + "\"
End If
    
On Error Resume Next
XX = FSO.FolderExists(W$)

Dim DateSet As Date

If DIRW$ <> "" Then W$ = DIRW$

'MsgBox str(XX) + "--" + W$
'End

Start2:

Dim TPath$

ScanPath.chkSubFolders = vbChecked
ScanPath.cboMask.Text = "*.avi"

'ScanPath.txtPath.Text = "E:\01 FAVS\#00 Palm\#000 New\2009-08 Aug\Unet\Mine\Word Example Usage"
'ScanPath.txtPath.Text = "E:\01 FAVS\Ministers\Images"
ScanPath.txtPath.Text = W$

If Len(W$) < 5 Then End

'ScanPath.Show
Dim DT1 As Date
'Dim Ds1 As Date
Dim DS2 As Date
'Dim DateSet As Date
For we = 1 To ScanPath.ListView1.ListItems.Count
    A11 = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
    B11 = ScanPath.ListView1.ListItems.Item(we)
    
    Set F = FSO.GetFILE(A11 + B11)
    DT1 = F.datelastmodified
    
    DateSet = DT1 - TimeSerial(1, 0, 0)
    
    Set F = Nothing
    
    'Ds1 = Mid(B11, 1, 20)
'    Ds1 = Replace(Ds1, "+", ":") + ":00"
'    If InStr(B11, "2008 011(Nov) 29 Sat 08-04-10") > 0 Then
    'If Year(DT1) = 2009 Then
        '2008 012(Dec) 04 Thu 18-22-52 - W880I - DSC00331.JPG
      '  Ds1 = DateValue("04-Dec-2008") + TimeValue("18:22:52")
        
'        xB11 = B11
'        xB11 = Replace(xB11, "MILL VIEW HOSPITAL-", "")
'
        If A11 + B11 = "D:\0 00 ART LOGGERS - WEBCAM\VIDEO\Web_Video_Capture- 2015-10-03 14-28-59 Sat -- Microsoft WDM Image Capture (Win32)2.avi" Then
'            Ds1 = DateValue(Mid(xB11, 9, 2) + "-" + Mid(xB11, 6, 2) + "-" + Mid(xB11, 1, 4))
'            Ds2 = TimeValue(Replace(Mid(xB11, 12, 8), "-", ":"))
        '    DateSet = DateValue(Ds1) + TimeValue(Ds1)
'            DateSet = Ds1 + Ds2
    '        DateSet = GetExif(A11 + B11)
            DateSet = DateValue("2015-10-03") + TimeValue("14:28:59")
                  
            '
            TT = SetFileDateTime(A11 + B11, DateSet)
            XC = XC + 1
            'tt = LastModifiedToCurrent(A11 + B11)
        End If
'        Stop
        'End
    'End If
    
Next

MsgBox "Done = " + vbCrLf + str(XC) '+dd$
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


Function GetExif(File$)
    Dim objExif As New ExifReader
    Dim txtExifInfo As String
    
    objExif.Load File$
    txtExifInfo = objExif.Tag(DateTimeOriginal)
    GetExif = txtExifInfo
End Function


'Function GetProperty(strFile, n)
'Dim objShell
'Dim objFolder
'Dim objFolderItem
'Dim i
'Dim strPath
'Dim strName
'Dim intPos
'
'On Error GoTo ErrHandler
'
'intPos = InStrRev(strFile, "")
'strPath = Left(strFile, intPos)
'strName = Mid(strFile, intPos + 1)
'Set objShell = CreateObject("Shell.Application")
'Set objFolder = objShell.Namespace(strPath)
'Set objFolderItem = objFolder.ParseName(strName)
'If Not objFolderItem Is Nothing Then
'GetProperty = objFolder.GetDetailsOf(objFolderItem, n)
'End If
'
'ExitHandler:
'Set objFolderItem = Nothing
'Set objFolder = Nothing
'Set objShell = Nothing
'Exit Function
'
'ErrHandler:
'MsgBox Err.Description, vbExclamation
'Resume ExitHandler
'End Function

Private Sub Label11_Click()

If Label1.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label2.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label9.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label11.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub

Label11.BackColor = Label_COLOR_YELLOW.BackColor
Label_GO_AH.BackColor = Label_COLOR_GREEN.BackColor

WORK = "FILE_CREATED_TO_MODIFIED_TIME"


End Sub

Private Sub Label2_Click()

If Label1.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label2.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label9.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label2.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub

Label2.BackColor = Label_COLOR_YELLOW.BackColor
Label_GO_AH.BackColor = Label_COLOR_GREEN.BackColor

WORK = "MOD_TO_CREATED_DATE"

' NEXT -- Label_GO_AH_Click

End Sub

Private Sub Label9_Click()


If Label11.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label9.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label1.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub
If Label2.BackColor = Label_COLOR_YELLOW.BackColor Then Exit Sub

Label9.BackColor = Label_COLOR_YELLOW.BackColor
Label_GO_AH.BackColor = Label_COLOR_GREEN.BackColor

WORK = "CREATED_TO_MOD_DATE"

' NEXT -- Label_GO_AH_Click
' NEXT -- Call BATCH_CREATED_TO_MODIFIED_TIME

End Sub

Sub SET_BATCH_DATE_CAMERA_VIDEO_FILENAME_TO_DATE_FILE()

    a = "D:\DSC\2015+Sony\2018 CyberShot HX60V\MP_ROOT\2018_10_31 Oct_Wed 11_55_50__MAQ01836 (2).mp4"

    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    ScanPath.txtPath.Text = "\\7-asus-gl522vw\7_ASUS_GL522VW_02_D_DRIVE\VI_ DSC ME 01\2010+SONY\2017 CyberShot HX60V_#\New folder"
    'ScanPath.txtPath.Text = LABEL_SET(2).Caption
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        a = B11
        
        XX = InStr(a, "\201_")
        DATEVAR = Replace(Mid(a, XX + 1, 10) + " " + Mid(a, XX + 20, 8), "_", "-")
        
        For r = 1 To 10
            If Mid(DATEVAR, r, 1) = "-" Then Mid(DATEVAR, r, 1) = "/"
        Next
        For r = 10 To Len(DATEVAR)
            If Mid(DATEVAR, r, 1) = "-" Then Mid(DATEVAR, r, 1) = ":"
        Next
        DateSet = DateValue(DATEVAR) + TimeValue(DATEVAR)
        
    '    Ds1 = DateValue(Mid(xB11, 9, 2) + "-" + Mid(xB11, 6, 2) + "-" + Mid(xB11, 1, 4))
    '    Ds2 = TimeValue(Replace(Mid(xB11, 12, 8), "-", ":"))
    '    DateSet = DateValue(Ds1) + TimeValue(Ds1)
    '    DateSet = Ds1 + Ds2
    '    DateSet = GetExif(A11 + B11)
                
        TT = SetFileDateTime(a, DateSet)
        XC = XC + 1
    Next
    
    MsgBox "Done = " + str(XC)
    End

End Sub


Private Sub ROUTINE_MOD_TO_NOW_DATE()

    ' If WORK = "MOD_TO_NOW_DATE" Then
    ' -------------------------------
    a = LABEL_SET(3).Caption           ' GET FILE
    DateSet = Now
    TT = SetFileDateTime(a, DateSet)
    ' MsgBox "Done " + vbCrLf + vbCrLf + a + vbCrLf + vbCrLf + DATEVAR, vbMsgBoxSetForeground
    End

End Sub


Private Sub SET_ONE_DATE_HARDCODER()
    
    
    WORK = "SET_ONE_DATE_HARDCODER"
    
    Call Label_GO_AH_Click

End Sub


Private Sub MNU_CLIPBOARD_DATE_Click()
    
    ' D:\DSC\2019+MEM-BANK-AR\2019 MEMORY-BANK-CAMERA\PHOTO\PHO00000.JPG
    
    Set F = FSO.GetFILE(LABEL_SET(3).Caption)
    DT1 = F.datelastmodified
    A11 = LABEL_SET(3).Caption
    D11 = A11
    If A11 <> "NOT FILE GIVEN" Then
        A11 = Mid(D11, 1, InStrRev(A11, "\"))
        B11 = Mid(D11, InStrRev(A11, "\") + 1)
    End If

    OUT_FOLDER = LABEL_SET(2).Caption + "\" + Format(DT1, "YYYY-MM MMM")
    Clipboard.Clear
    Clipboard.SetText " " + Format(DT1, "YYYY-MM-DD HH-MM-SS") + " " + B11
    
End Sub

Private Sub MNU_CREATE_FOLDER_DATE_MONTH_Click()

    Set F = FSO.GetFILE(LABEL_SET(3).Caption)
    DT1 = F.datelastmodified

    OUT_FOLDER = LABEL_SET(2).Caption + "\" + Format(DT1, "YYYY-MM MMM")
    
    MkDir OUT_FOLDER

    End


End Sub

Private Sub MNU_CREATE_FOLDER_DATE_OF_FILE_Click()

    Set F = FSO.GetFILE(LABEL_SET(3).Caption)
    DT1 = F.datelastmodified

    OUT_FOLDER = LABEL_SET(2).Caption + "\" + Format(DT1, "YYYY-MM-DD")
    If InStr(LABEL_SET(3).Caption, "REC") > 0 And InStr(LABEL_SET(3).Caption, ".WAV") > 0 Then
        OUT_FOLDER = LABEL_SET(2).Caption + "\RECORD " + Format(DT1, "YYYY-MM-DD")
    End If
    
    
    CreateFolderTree OUT_FOLDER

    End


End Sub

Private Sub MNU_CREATE_TEXT_FILE_FOR_DATE_Click()
    
    ScanPath.chkSubFolders = vbChecked
    ScanPath.cboMask.Text = "*.*"
    ScanPath.txtPath.Text = LABEL_SET(2).Caption
    
    For RR = 1 To ScanPath.ListView1.ListItems.Count
        A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
        B11 = ScanPath.ListView1.ListItems.Item(RR)
        If InStr(B11, "TEXT_DATER.TXT") = 0 Then
            Exit For
        End If
    Next
    Set F = FSO.GetFILE((A11 + B11))
    ADATE1 = F.datelastmodified
    DUPE = A11 + B11
    
    XX = ScanPath.txtPath.Text + "\# TEXT_DATER.TXT"
    FR1 = FreeFile
    Open XX For Output As #FR1
        Print #FR1, Format(ADATE1, "YYYY/MM/DD HH-MM-SS")
        Print #FR1,
        Print #FR1, Replace(UCase(A11) + UCase(B11), "\", " \ ")
        Print #FR1,
        For RR = 1 To ScanPath.ListView1.ListItems.Count
            A11 = ScanPath.ListView1.ListItems.Item(RR).SubItems(1)
            B11 = ScanPath.ListView1.ListItems.Item(RR)
            GODATE = 1
            If InStr(B11, "TEXT_DATER.TXT") > 0 Then GODATE = 0
            If DUPE = A11 + B11 Then GODATE = 0
            If GODATE = 1 Then
                Set F = FSO.GetFILE((A11 + B11))
                ADATE1 = F.datelastmodified
                Print #FR1, Replace(UCase(A11) + UCase(B11), "\", " \ ")
                Print #FR1, Format(ADATE1, "YYYY/MM/DD HH-MM-SS")
                Print #FR1,
            End If
        Next
        
        Print #FR1, "# -- ----------------------------------------------------------------"
        Print #FR1, "# -- ONLY FIRST LINE ACT ON -----------------------------------------"
        Print #FR1, "# -- HERE IS SOME CODE PROCESSOR ------------------------------------"
        Print #FR1, "# -- BACK LASH CONVERT TO FORWARD \ / -------------------------------"
        Print #FR1, "# -- UNDERLINE __ CONVERT TO COLON :: -------------------------------"
        Print #FR1, "# -- DASH -- ---- CONVERT TO COLON :: -------------------------------"
        Print #FR1, "# -- ----------------------------------------------------------------"
        Print #FR1, "# -- TIME_STRING = Replace(TIME_STRING, ""\"", ""/"")"
        Print #FR1, "# -- TIME_STRING = Replace(TIME_STRING, ""_"", "":"")"
        Print #FR1, "# -- TIME_STRING = Replace(TIME_STRING, ""-"", "":"")"
        Print #FR1, "# -- ----------------------------------------------------------------"
        Print #FR1, "# -- If Val(TimeValue(TIME_STRING)) = 0 Then "
        Print #FR1, "# --    TIME_STRING = Format(DateValue(TIME_STRING), ""YYYY/MM/DD"") + ""10:00:00"""
        Print #FR1, "# -- END IF"
        Print #FR1, "# -- ----------------------------------------------------------------"
        Print #FR1, "# -- If TimeValue(TIME_STRING) = 0 Then"
        Print #FR1, "# -- ----------------------------------------------------------------"
        Print #FR1, "# -- ADD A BIT ON DATE FOR DAYLIGHT SAVING AND NTFS AT EXACT MIDNIGHT"
        Print #FR1, "# -- CLOUD SYSTEM REQUIRING A BIT MORE ALSO OR DAY BEFORE"
        Print #FR1, "# -- ----------------------------------------------------------------"
        Print #FR1, "# --    DT4 = DT4 + TimeSerial(11, 0, 0)"
        Print #FR1, "# -- ----------------------------------------------------------------"
        Print #FR1, "# -- END IF"
        Print #FR1, "# -- MM_1 = Format(DT4, ""DD-MM-YYYY  HH:MM:SS  DDDD"")"
        Print #FR1, "# -- ----------------------------------------------------------------"
    Close #FR1
    
    Me.WindowState = vbMinimized
    Shell "C:\Program Files (x86)\Notepad++\notepad++.exe """ + XX + """", vbMaximizedFocus
    End

End Sub

Private Sub MNU_EXIT_Click()
End
End Sub

Private Sub MNU_INCREMENT_DATE_Click()
MENU_INCREMENT_DATE_VAR = Not MENU_INCREMENT_DATE_VAR
If MENU_INCREMENT_DATE_VAR = 0 Then
    MNU_INCREMENT_DATE.Caption = "INCREMENT DATE -- FALSE"
Else
    MNU_INCREMENT_DATE.Caption = "INCREMENT DATE -- TRUE"
End If
End Sub
Private Sub MNU_INCREMENT_DATE_SHOW()
If MENU_INCREMENT_DATE_VAR = 0 Then
    MNU_INCREMENT_DATE.Caption = "INCREMENT DATE -- FALSE"
Else
    MNU_INCREMENT_DATE.Caption = "INCREMENT DATE -- TRUE"
End If
End Sub

Private Sub MNU_QUICK_MODE_Click()

'QUICK MODE -- ON -- GIVE MESSENGER

I_1 = "QUICK MODE -- ON -- NOT MESSENGER"
I_2 = "QUICK MODE -- OFF -- GIVE MESSENGER"
If InStr(MNU_QUICK_MODE.Caption, I_1) Then
    MNU_QUICK_MODE.Caption = I_2
    Else
    MNU_QUICK_MODE.Caption = I_1
End If


Dim VAR_DATE As Date

On Error Resume Next
Err.Clear
FR1 = FreeFile
Open App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT" For Output As #FR1
    Print #FR1, MNU_QUICK_MODE.Caption
    VAR_DATE = Now + TimeSerial(24, 0, 0)
    Print #FR1, VAR_DATE
Close FR1

If Err.Number = 0 Then
    I_1 = "QUICK MODE -- ON -- NOT MESSENGER"
    I_2 = "QUICK MODE -- OFF -- GIVE MESSENGER"
    If InStr(MNU_QUICK_MODE.Caption, I_1) Then
        MsgBox "ENJOY GLOBAL FOR ALL INSTANCE RUNNER" + vbCrLf + vbCrLf + "REMEMEBER FOR 1 DAY DEFAULT OFF", vbMsgBoxSetForeground
    End If
Else
    MsgBox "ERROR -- FILE NOT WRITEN -- TRY AGAIN" + vbCrLf + vbCrLf + App.Path + "\# DATA\QUICK_MODE_SET #NFS " + GetComputerName + ".TXT", vbMsgBoxSetForeground
End If

End Sub



Private Sub MNU_VB_FOLDER_Click()
Shell "EXPLORER /SELECT, " + App.Path + "\" + App.EXEName + ".VBP", vbMaximizedFocus
End Sub

Private Sub MNU_VB_LAUNCH_FAV_SET_Click()
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
    WSHShell.RUN """" + "C:\SCRIPTER\SCRIPTER CODE -- AUTOHOTKEY\Autokey -- 31-AUTORUN SET FAV VB & AUTOHOTKEY.ahk" + """"
Set WSHShell = Nothing
End Sub

Private Sub MNU_VB_ME_Click()
    
    Dim CODER_VBP_FILE_NAME_2
    Dim VB_1, VB_2, VB_3

    VB_1 = "C:\PROGRAM FILES\Microsoft Visual Studio\VB98\VB6.EXE"
    VB_2 = "C:\PROGRAM FILES (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
    If Dir(VB_1) <> "" Then VB_3 = VB_1
    If Dir(VB_2) <> "" Then VB_3 = VB_2
    '------------------------------------------------------
    'ADD MICROSFOT SCRIPTING RUNTIME FOR HERE IN REFERENCES
    '------------------------------------------------------
    Dim objShell
    Set objShell = CreateObject("Wscript.Shell")
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
    objShell.RUN """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
    Set objShell = Nothing
    
    EXIT_TRUE = True
    Unload Me
    Exit Sub

    
    Beep
    On Error Resume Next
    Dim IGETTER
    IGETTER = GetSpecialfolder(38)
    
    If IGETTER = "" Then
        MsgBox "ERROR LOAD ADMIN"
        MsgBox "YOU HAVE TO SWITCH ADMIN MODE ON" + vbCrLf + vbCrLf + "THE COMMAND GetSpecialfolder(38) REQUIRE ADMINISTRATOR" + vbCrLf + vbCrLf + "RESULT RETURN NOTHING HERE GetSpecialfolder(38)__<" + GetSpecialfolder(38) + ">__ RESULT RETURNED NOTHING", vbMsgBoxSetForeground
        'MsgBox "YOU HAVE TO SWITCH ADMIN MODE ON" + vbCrLf + vbCrLf + "THE COMMAND GetSpecialfolder(38) REQUIRE ADMINISTRATOR" + vbCrLf + vbCrLf + "RESULT RETURN NOTHING HERE GetSpecialfolder(38)__< GetSpecialfolder(38) >__ RESULT RETURNED NOTHING", vbMsgBoxSetForeground
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
    
    ' Call GET_VB_PROJECT_NAME_TO_EXE(CODER_EXE_FILE_NAME_1, CODER_VBP_FILE_NAME_2)
    
'    MsgBox """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + CODER_VBP_FILE_NAME_2 + """"
    
    Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + CODER_VBP_FILE_NAME_2 + """", vbNormalFocus
    
    Me.WindowState = vbMinimized 'to down and en-able exit
    
    EXIT_TRUE = True
    Unload Me
    'End

End Sub


Public Function GetSpecialFolder_Show_Script_Debug(CSIDL As Long) As String

Dim r As Long
On Error Resume Next
For r = 0 To 120
    If Trim(GetSpecialfolder(r)) <> "" Then
        'Debug.Print Str(R) + " -- " + GetSpecialfolder(R)
        'AAX = GetSpecialfolder(R)
    End If
Next
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




' --------------------------------------------------------------------
' ALL IN ONE FUNCTION LESS API LESS DEMAND JOB WORKER
' --------------------------------------------------------------------
' CreateFolderTree CreatePATH Create_PATH
' --------------------------------------------------------------------
Private Function CreateFolderTree(ByVal sPath As String) As Boolean
    Dim nPos As Integer

    If Mid(sPath, Len(sPath), 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)

    On Error GoTo CreateFolderTreeError
    
    nPos = InStr(sPath, "\")
'    If Mid(sPath, 1, 2) = "\\" Then
'        nPos = InStr(nPos + 1, sPath, "\")
'        nPos = InStr(nPos + 1, sPath, "\")
'        nPos = InStr(nPos + 1, sPath, "\")
'    End If
    
    While nPos > 0
        If nPos - 1 > 3 Then
            If Dir(Left$(sPath, nPos - 1), vbDirectory) = "" Then
                MkDir Left$(sPath, nPos - 1)
            End If
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
    If Dir(sPath, vbDirectory) = "" Then MkDir sPath
    
    CreateFolderTree = True
    If Dir(sPath, vbDirectory) = "" Then
        CreateFolderTree = False
    End If

    Exit Function

CreateFolderTreeError:
    Exit Function
End Function

Private Function CreateFolderTree_AND_NETWORK(ByVal sPath As String) As Boolean
    Dim nPos As Integer
    Dim strarray
    Dim r
    Dim LINE_PATH_BUILDER

    If Mid(sPath, Len(sPath), 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)
    
    If Mid(sPath, 1, 2) = "\\" Then
        LINE_PATH_BUILDER = "\\"
    End If
    If Mid(sPath, 2, 2) = ":\" Then
        LINE_PATH_BUILDER = "" 'Mid(sPath, 1, 3)
    End If
    
    On Error Resume Next
    
    strarray = Split(sPath, "\")
    For r = 0 To UBound(strarray)
        If strarray(r) <> "" Then
            LINE_PATH_BUILDER = LINE_PATH_BUILDER + strarray(r)
                        
            FSO.CreateFolder LINE_PATH_BUILDER
            ' -----------------------------------
            ' THIS HAS ERROR IS NOT ABLE MAKE PATH
            ' USE FSO
            ' -----------------------------------
            MkDir LINE_PATH_BUILDER
            
            LINE_PATH_BUILDER = LINE_PATH_BUILDER + "\"
            
        End If
    Next
    
    
    
    
    nPos = InStr(sPath, "\")
'    If Mid(sPath, 1, 2) = "\\" Then
'        nPos = InStr(nPos + 1, sPath, "\")
'        nPos = InStr(nPos + 1, sPath, "\")
'        nPos = InStr(nPos + 1, sPath, "\")
'    End If
    
'    While nPos > 0
'        If nPos - 1 > 3 Then
'            If Dir(Left$(sPath, nPos - 1), vbDirectory) = "" Then
'                MkDir Left$(sPath, nPos - 1)
'            End If
'        End If
'        nPos = InStr(nPos + 1, sPath, "\")
'    Wend
'    If Dir(sPath, vbDirectory) = "" Then MkDir sPath
    
'    CreateFolderTree_AND_NETWORK = True
'    If Dir(sPath, vbDirectory) = "" Then
'        CreateFolderTree_AND_NETWORK = False
'    End If

    Exit Function

CreateFolderTreeError:
    Exit Function
End Function





' --------------------------------------------------------------------
' WORK BUT EASIER ALL IN ONE FUNCTION NOT WITH API THOUGH
' MORE TRANSPORTBALE TO OTHER CODE MAYBE IN ONE FUNCTION LESS DELCLARE
' --------------------------------------------------------------------
Private Function CreateFolderTree_API(ByVal sPath As String) As Boolean
    Dim nPos As Integer

    If Mid(sPath, Len(sPath), 1) = "\" Then sPath = Mid(sPath, 1, Len(sPath) - 1)

    On Error GoTo CreateFolderTreeError
    
    nPos = InStr(sPath, "\")
    While nPos > 0
        If nPos - 1 > 3 Then ' NOT CHECK ROOT
            If Not FolderExists(Left$(sPath, nPos - 1)) Then
                MkDir Left$(sPath, nPos - 1)
            End If
        End If
        nPos = InStr(nPos + 1, sPath, "\")
    Wend
    If Not FolderExists(sPath) Then MkDir sPath
    
    If Not FolderExists(sPath) Then
        CreateFolderTree_API = False
        Exit Function
    End If
    CreateFolderTree_API = True
    Exit Function

CreateFolderTreeError:
    Exit Function
End Function

Private Function FolderExists(sFolder As String) As Boolean
    '##############################################################################################
    'Returns True if the specified folder exists
    '##############################################################################################
    
    Dim WFD As WIN32_FIND_DATA
    Dim lResult As Long
    
    lResult = FindFirstFile(sFolder, WFD)
    If lResult <> INVALID_HANDLE_VALUE Then
        If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            FolderExists = True
        Else
            FolderExists = False
        End If
    End If
End Function


Private Sub MNU_FILEDATE_WHOLE_FOLDER_AVI_MP4_ANY_Click()
    
    If LABEL_SET(2).Caption <> "NOT FOLDER GIVEN" Then
        File1.Path = LABEL_SET(2).Caption
    End If
    
    If IsIDE = True Then
        File1.Path = "F:\DSC--2020+128GB-V-OPA\2020_128GB-V-OPA\2020-06-20__128GB-V-OPA"
    End If
    File1.Pattern = "*.*"
    
    ' SOMETIME THE PASS BY CONTEXT MENU ISN'T LONG NAME BUT SHORTNAME
        
    R_C_COUNTER_MAX = 0
    For R_C_2 = 1 To 3
        
        A_M = ""
        
        R_C_COUNTER = 0
            
        For R_C = 0 To File1.ListCount - 1
    
            FILE_NAME_MP4 = File1.Path + "\" + File1.List(R_C)
        
            If Mid(FILE_NAME_MP4, 1, 2) <> "\\" Then
                FILE_NAME_MP4 = GetLongName(FILE_NAME_MP4)
            End If
            
            Set F = FSO.GetFILE((FILE_NAME_MP4))
            ADATE1 = F.datelastmodified
            
            'Clipboard.Clear
            'Sleep 200
            'If Year(ADATE1) = 2016 Then at1 = "-- 2k Sixteenth"
            'If Year(ADATE1) = 2015 Then at1 = "-- 2k Fifteenth"
            'If Year(ADATE1) <= 2015 Then at1 = "-- " + Format(ADATE1, "YYYY")
            'If Year(ADATE1) > 2016 Then at1 = "-- " + Format(ADATE1, "YYYY")
            
            'at1 = "-- " + Format(ADATE1, "YYYY")
            
            '------------------------
            ' LONG DATE FOR YOUTUBING
            '------------------------
            'If InStr(FILE_NAME_MP4, "D:\DSC\") > 0 Then
            'Clipboard.SetText " " + at1 + Format(ADATE1, " MMMM DD DDDD HH-MM-SS Am/Pm")
            FILE_NAME_MP4_2 = Mid(FILE_NAME_MP4, InStrRev(FILE_NAME_MP4, "\") + 1)
            FILE_NAME_MP4_2_EXT = Mid(FILE_NAME_MP4, InStrRev(FILE_NAME_MP4, "."))
            FILE_NAME_MP4_2_NOT_EXT = UCase(Mid(File1.List(R_C), 1, InStrRev(File1.List(R_C), ".") - 1))
            FILE_NAME_MP4_SET = FILE_NAME_MP4_2_NOT_EXT + UCase(FILE_NAME_MP4_2_EXT)
            FILE_NAME_MP4_SET_AND_DATE = File1.Path + "\" + Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + FILE_NAME_MP4_SET
            MAQ = ""
            START_MAQ = 0
            
            ' Call MNU_FILEDATE_CLIPBOARD_ROUTINE_CHECKER
            
            If START_MAQ = 1 Or 1 = 1 Then
                If File1.List(R_C) <> FILE_NAME_MP4_SET_AND_DATE Then
                    ' If MAQ <> "" Then
                    
                        R_C_COUNTER = R_C_COUNTER + 1
                        If R_C_2 = 1 Then
                            R_C_COUNTER_MAX = R_C_COUNTER_MAX + 1
                        End If
                        
                        A_M = A_M + "RENAMER --" + str(R_C_COUNTER) + " OF" + str(R_C_COUNTER_MAX) + " OF" + str(File1.ListCount) + vbCrLf + File1.Path + "\" + vbCrLf + "\" + File1.List(R_C) + vbCrLf + "----" + vbCrLf + FILE_NAME_MP4_SET_AND_DATE
                        A_M = A_M + vbCrLf
                        A_M = A_M + vbCrLf
                        
                        If R_C_2 = 3 Then
                        
                            If R_C_COUNTER < 10 Then
                            
                                A_M_2 = "RENAMER -- VERIFY 1ST 10  AND THEN AUTO" + vbCrLf + Trim(str(R_C + 1)) + " OF" + str(File1.ListCount) + vbCrLf + File1.Path + "\" + vbCrLf + "\" + File1.List(R_C) + vbCrLf + "----" + vbCrLf + FILE_NAME_MP4_SET_AND_DATE
                                
                                X_M = MsgBox(A_M_2, vbYesNo)
                                If X_M = vbNo Then
                                    R_C_COUNTER = R_C_COUNTER - 1
                                    Exit For
                                End If
                            End If
                            
                            Name File1.Path + "\" + File1.List(R_C) As FILE_NAME_MP4_SET_AND_DATE
                        
                        End If
                    ' End If
                End If
            End If
        Next
            
        If R_C_2 = 2 Then
            If A_M <> "" Then
                R_C_COUNTER = 0
                X_M = MsgBox("RENAME CONVERSION AS FOLLOW PROCESS YES / NOT" + vbCrLf + vbCrLf + A_M, vbYesNo)
                If X_M = vbNo Then Exit For
            End If
        End If
    Next
        
    If R_C_COUNTER > 0 Then
    
    MsgBox "DONE " + Trim(str(R_C_COUNTER)) + " CHANGER"
    
    End If
    
    End
End Sub

Private Sub MNU_FILEDATE_WHOLE_FOLDER_Click()
    '---------------------------------------------
    'VIDEO MP4
    '---------------------------------------------

    ' If FILE_NAME_MP4 = "" Then Beep: MsgBox "NOT A FILE NAME PIPED": End
    
    File1.Path = LABEL_SET(2).Caption
    
    ' SOMETIME THE PASS BY CONTEXT MENU ISN'T LONG NAME BUT SHORTNAME
        
    R_C_COUNTER_MAX = 0
    For R_C_2 = 1 To 3
        
        A_M = ""
        
        R_C_COUNTER = 0
            
        For R_C = 0 To File1.ListCount - 1
    
            FILE_NAME_MP4 = File1.Path + "\" + File1.List(R_C)
        
            If Mid(FILE_NAME_MP4, 1, 2) <> "\\" Then
                FILE_NAME_MP4 = GetLongName(FILE_NAME_MP4)
            End If
            
            Set F = FSO.GetFILE((FILE_NAME_MP4))
            ADATE1 = F.datelastmodified
            
            'Clipboard.Clear
            'Sleep 200
            'If Year(ADATE1) = 2016 Then at1 = "-- 2k Sixteenth"
            'If Year(ADATE1) = 2015 Then at1 = "-- 2k Fifteenth"
            'If Year(ADATE1) <= 2015 Then at1 = "-- " + Format(ADATE1, "YYYY")
            'If Year(ADATE1) > 2016 Then at1 = "-- " + Format(ADATE1, "YYYY")
            
            'at1 = "-- " + Format(ADATE1, "YYYY")
            
            
            '------------------------
            ' LONG DATE FOR YOUTUBING
            '------------------------
            'If InStr(FILE_NAME_MP4, "D:\DSC\") > 0 Then
            'Clipboard.SetText " " + at1 + Format(ADATE1, " MMMM DD DDDD HH-MM-SS Am/Pm")
            FILE_NAME_MP4_2 = Mid(FILE_NAME_MP4, InStrRev(FILE_NAME_MP4, "\") + 1)
            FILE_NAME_MP4_2_EXT = Mid(FILE_NAME_MP4, InStrRev(FILE_NAME_MP4, "."))
            
            MAQ = ""
            START_MAQ = 0
            
            Call MNU_FILEDATE_CLIPBOARD_ROUTINE_CHECKER
            
            If START_MAQ = 1 Then
                If File1.List(R_C) <> Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + MAQ + NAME_PART_02 + FILE_NAME_MP4_2_EXT Then
                    If MAQ <> "" Then
                    
                        R_C_COUNTER = R_C_COUNTER + 1
                        If R_C_2 = 1 Then
                            R_C_COUNTER_MAX = R_C_COUNTER_MAX + 1
                        
                        End If
                        
                        A_M = A_M + "RENAMER --" + str(R_C_COUNTER) + " OF" + str(R_C_COUNTER_MAX) + " OF" + str(File1.ListCount) + vbCrLf + File1.Path + "\" + vbCrLf + "\" + File1.List(R_C) + vbCrLf + "\" + Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + MAQ + NAME_PART_02 + FILE_NAME_MP4_2_EXT
                        A_M = A_M + vbCrLf
                        A_M = A_M + vbCrLf
                        
                        
                        If R_C_2 = 3 Then
                        
                            If R_C_COUNTER < 10 Then
                            
                                A_M_2 = "RENAMER -- VERIFY 1ST 10  AND THEN AUTO" + vbCrLf + Trim(str(R_C + 1)) + " OF" + str(File1.ListCount) + vbCrLf + File1.Path + "\" + vbCrLf + "\" + File1.List(R_C) + vbCrLf + "\" + Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + MAQ + NAME_PART_02 + FILE_NAME_MP4_2_EXT
                                
                                X_M = MsgBox(A_M_2, vbYesNo)
                                If X_M = vbNo Then
                                    R_C_COUNTER = R_C_COUNTER - 1
                                    Exit For
                                End If
                            End If
                            
                            Name File1.Path + "\" + File1.List(R_C) As File1.Path + "\" + Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + MAQ + NAME_PART_02 + FILE_NAME_MP4_2_EXT
                        
                        End If
                    End If
                End If
            End If
        Next
            
        If R_C_2 = 2 Then
            If A_M <> "" Then
                R_C_COUNTER = 0
                X_M = MsgBox("RENAME CONVERSION AS FOLLOW PROCESS YES / NOT" + vbCrLf + vbCrLf + A_M, vbYesNo)
                If X_M = vbNo Then Exit For
            End If
        End If
    Next
        
    If R_C_COUNTER > 0 Then
    
    MsgBox "DONE " + Trim(str(R_C_COUNTER)) + " CHANGER"
    
    End If
    
    End
End Sub



Private Sub MNU_FILEDATE_CLIPBOARD_Click()
    ' MNU_FILEDATE_CLIPBOARD
    ' [ Sunday 09:46:40 Am_18 November 2018 ]

    '---------------------------------------------
    'VIDEO MP4
    '---------------------------------------------

    If FILE_NAME_MP4 = "" Then Beep: MsgBox "NOT A FILE NAME PIPED": End
    
    Beep
    
    Me.Hide

'    On Error Resume Next
'    Me.Visible = True
'    DoEvents
'    Me.WindowState = vbMinimized
'    DoEvents
'    MsgBox FILE_NAME_MP4, vbMsgBoxSetForeground
'    On Error GoTo 0

'    SOMETIME THE PASS BY CONTEXT MENU ISN;T LONG NAME BUT SHORTNAME

    If Mid(FILE_NAME_MP4, 1, 2) <> "\\" Then
        FILE_NAME_MP4 = GetLongName(FILE_NAME_MP4)
    End If
    
    Set F = fs.GetFILE((FILE_NAME_MP4))
    ADATE1 = F.datelastmodified
    
'    Clipboard.Clear
'    Sleep 200
    'If Year(ADATE1) = 2016 Then at1 = "-- 2k Sixteenth"
    'If Year(ADATE1) = 2015 Then at1 = "-- 2k Fifteenth"
    'If Year(ADATE1) <= 2015 Then at1 = "-- " + Format(ADATE1, "YYYY")
    'If Year(ADATE1) > 2016 Then at1 = "-- " + Format(ADATE1, "YYYY")
    
    'at1 = "-- " + Format(ADATE1, "YYYY")
    
    
    '------------------------
    ' LONG DATE FOR YOUTUBING
    '------------------------
    'If InStr(FILE_NAME_MP4, "D:\DSC\") > 0 Then
    'Clipboard.SetText " " + at1 + Format(ADATE1, " MMMM DD DDDD HH-MM-SS Am/Pm")
    FILE_NAME_MP4_2 = Mid(FILE_NAME_MP4, InStrRev(FILE_NAME_MP4, "\") + 1)
    'MsgBox FILE_NAME_MP4_2
    MAQ = ""
    START_MAQ = 0
    

    Call MNU_FILEDATE_CLIPBOARD_ROUTINE_CHECKER
    
    Clipboard.Clear
    Sleep 200

    Clipboard.SetText Format(ADATE1, "YYYY_MM_DD MMM_DDD HH_MM_SS") + "__" + MAQ + NAME_PART_02
    'Else
    '------------------------
    ' OTHER DATE
    '------------------------
    'Clipboard.SetText Format(ADATE1, "YYYY-MM-DD")
    'End If

    End
    
End Sub

Sub MNU_FILEDATE_CLIPBOARD_ROUTINE_CHECKER()
    TEE = "MAQ"
    If InStr(FILE_NAME_MP4_2, TEE) > 0 Then
        START_MAQ_LEN = 8
        MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
        START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
    End If
    
    TEE = "MAH"
    If InStr(FILE_NAME_MP4_2, TEE) > 0 Then
        START_MAQ_LEN = 8
        MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
        START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
    End If
    
    TEE = "M4V"
    If InStr(FILE_NAME_MP4_2, TEE) > 0 Then
        START_MAQ_LEN = 8
        MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
        START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
    End If
    
    TEE = "M4H"
    If InStr(FILE_NAME_MP4_2, TEE) > 0 Then
        START_MAQ_LEN = 8
        MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
        START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
    End If
    
    TEE = "DSCF"
    If InStr(GetLongName(FILE_NAME_MP4), TEE) > 0 Then
        If InStr(GetLongName(FILE_NAME_MP4), ".MOV") > 0 Then
            START_MAQ_LEN = 8
            MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
            START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
        End If
    End If

    'DOUBLE SCREEN CAMERA
    TEE = "DSCF"
    If InStr(GetLongName(FILE_NAME_MP4), TEE) > 0 Then
        If InStr(GetLongName(FILE_NAME_MP4), ".AVI") > 0 Then
            START_MAQ_LEN = 8
            MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
            START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
        End If
    End If
    
    TEE = "image"
    If InStr(GetLongName(FILE_NAME_MP4), TEE) > 0 Then
        If InStr(GetLongName(FILE_NAME_MP4), ".AVI") > 0 Then
            START_MAQ_LEN = 7
            MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
            START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
        End If
    End If
    
    TEE = "Photo"
    If InStr(GetLongName(FILE_NAME_MP4), TEE) > 0 Then
        If InStr(GetLongName(FILE_NAME_MP4), ".JPG") > 0 Then
            START_MAQ_LEN = 7
            MAQ = Mid(FILE_NAME_MP4_2, InStr(FILE_NAME_MP4_2, TEE), START_MAQ_LEN)
            START_MAQ = InStr(FILE_NAME_MP4_2, TEE)
        End If
    End If
    
    
    'IF DESCRIPTION ALREADY ON
    NAME_PART_02 = ""
    If START_MAQ > 0 Then
        NAME_PART_02 = Mid(FILE_NAME_MP4_2, START_MAQ + START_MAQ_LEN)
        NAME_PART_02 = Mid(NAME_PART_02, 1, InStrRev(NAME_PART_02, ".") - 1)
    End If
End Sub


Public Function GetLongName(ByVal sShortName As String) As String

    If Mid(sShortName, 1, 2) = "\\" Then
        GetLongName = sShortName
        Exit Function
    End If

' --> (All this modules code) Obtained from : -
' ---> Microsoft's/MSDN's code
' ->  http://support.microsoft.com/default.aspx?scid=kb;EN-US;154822
' ---> The original comments were by them :

     Dim sLongName As String
     Dim sTemp As String
     Dim iSlashPos As Integer
     Dim sShortName_ENTRY As String

     sShortName_ENTRY = sShortName

     'Add \ to short name to prevent Instr from failing
     
     sShortName = sShortName & "\"

     'Start from 4 to ignore the "[Drive Letter]:\" characters
     iSlashPos = InStr(4, sShortName, "\")

     'Pull out each string between \ character for conversion
     While iSlashPos
       sTemp = Dir(Left$(sShortName, iSlashPos - 1), vbNormal + vbHidden + vbSystem + vbDirectory)
       If sTemp = "" Then
            'Error 52 - Bad File Name or Number
            GetLongName = ""
            If Trim(GetLongName) = "" Then GetLongName = sShortName_ENTRY
            'SOMETIME SHORT NAME ENTRY WORK SOMETIME NOT
            'MAYBE ERROR IF PARAM IN LINK __ TESTER
            'SOLVED ERROR IS IN PERMISSION OF FOLDER LIKE \PROGRAM FILES (X86)
            '"C:\Program Files (x86)\Process Lasso\ProcessLassoLauncher.exe"
            'SOLVED ERROR 64 BIT IN 32 BIT VB6 SHORTCUT FINDER
            'WORKED ERRRO FROM HERE AGAIN
            '"C:\Program Files (x86)\Process Lasso\ProcessLassoLauncher.exe"
            
            
            
            
             Exit Function
       End If
       sLongName = sLongName & "\" & sTemp
       iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
     Wend

     'Prefix with the drive letter
     GetLongName = Left$(sShortName, 2) & sLongName
    If Trim(GetLongName) = "" Then GetLongName = sShortName_ENTRY
    'SOMETIME SHORT NAME ENTRY WORK SOMETIME NOT
    'MAYBE ERROR IF PARAM IN LINK __ TESTER

End Function


Private Sub MNU_VERSION_Click()
' MNU_VERSION.CAPTION = "VER_2020_" + TRIM(STR(APP.MAJOR)) + "." + TRIM(STR(APP.MINOR)) + "." + TRIM(STR(APP.REVISION)) ' + " _ MATT.LAN@BTINTERNET.COM"
End Sub
