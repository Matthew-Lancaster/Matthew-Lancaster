VERSION 5.00
Begin VB.Form Sample 
   Caption         =   "Sample"
   ClientHeight    =   5925
   ClientLeft      =   1065
   ClientTop       =   2715
   ClientWidth     =   16755
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   395
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1117
   Begin VB.PictureBox picSource 
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   8370
      Picture         =   "Sample.frx":0000
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   54
      Top             =   135
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox PicView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      HasDC           =   0   'False
      Height          =   5820
      Left            =   10935
      ScaleHeight     =   5760
      ScaleWidth      =   5760
      TabIndex        =   53
      Top             =   0
      Width           =   5820
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   4590
      Pattern         =   "*.txt"
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   0
      TabIndex        =   44
      Top             =   1890
      Width           =   10905
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1395
      TabIndex        =   43
      Top             =   0
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   39
      Left            =   9945
      TabIndex        =   41
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   38
      Left            =   9450
      TabIndex        =   40
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   37
      Left            =   8955
      TabIndex        =   39
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   36
      Left            =   8460
      TabIndex        =   38
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   35
      Left            =   7965
      TabIndex        =   37
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   34
      Left            =   7470
      TabIndex        =   36
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   33
      Left            =   6975
      TabIndex        =   35
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   32
      Left            =   6480
      TabIndex        =   34
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   31
      Left            =   5985
      TabIndex        =   33
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   30
      Left            =   5490
      TabIndex        =   32
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   29
      Left            =   4995
      TabIndex        =   31
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   28
      Left            =   4500
      TabIndex        =   30
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   27
      Left            =   4005
      TabIndex        =   29
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   26
      Left            =   3510
      TabIndex        =   28
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   25
      Left            =   3015
      TabIndex        =   27
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   24
      Left            =   2520
      TabIndex        =   26
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   23
      Left            =   2025
      TabIndex        =   25
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   22
      Left            =   1530
      TabIndex        =   24
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   21
      Left            =   1035
      TabIndex        =   23
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   20
      Left            =   540
      TabIndex        =   22
      Text            =   " "
      Top             =   765
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   19
      Left            =   9945
      TabIndex        =   21
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   18
      Left            =   9450
      TabIndex        =   20
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   17
      Left            =   8955
      TabIndex        =   19
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   16
      Left            =   8460
      TabIndex        =   18
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   15
      Left            =   7965
      TabIndex        =   17
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   14
      Left            =   7470
      TabIndex        =   16
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   13
      Left            =   6975
      TabIndex        =   15
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   12
      Left            =   6480
      TabIndex        =   14
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   11
      Left            =   5985
      TabIndex        =   13
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   10
      Left            =   5490
      TabIndex        =   12
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   9
      Left            =   4995
      TabIndex        =   11
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   8
      Left            =   4500
      TabIndex        =   10
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   7
      Left            =   4005
      TabIndex        =   9
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   3510
      TabIndex        =   8
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   3015
      TabIndex        =   7
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   2520
      TabIndex        =   6
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   2025
      TabIndex        =   5
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   1530
      TabIndex        =   4
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   1035
      TabIndex        =   3
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   540
      TabIndex        =   2
      Text            =   " "
      Top             =   405
      Width           =   465
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000FF00&
      Caption         =   "Copy to Clipboard"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1980
      TabIndex        =   55
      Top             =   1170
      Width           =   1995
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000FF00&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9090
      TabIndex        =   52
      Top             =   45
      Width           =   1275
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4680
      TabIndex        =   51
      Top             =   1530
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Caption         =   "Total Number of Fields ="
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1980
      TabIndex        =   50
      Top             =   1530
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FF00&
      Caption         =   "Load File"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3285
      TabIndex        =   48
      Top             =   0
      Width           =   1275
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Down"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1260
      TabIndex        =   47
      Top             =   1170
      Width           =   645
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   585
      TabIndex        =   46
      Top             =   1170
      Width           =   600
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   45
      Top             =   1170
      Width           =   465
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Save File"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   42
      Top             =   0
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   45
      Left            =   45
      TabIndex        =   1
      Top             =   405
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   44
      Left            =   45
      TabIndex        =   0
      Top             =   765
      Width           =   465
   End
End
Attribute VB_Name = "Sample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public count2 As Integer
Public easy As Integer
Public xchange As Integer
Public pos3 As Integer
Public index2 As Integer

Public size As Double
Public size2 As Double
Public ratio As Double
Public xgag As Integer
Public patth As Double
Public pattw As Double
Public change1 As Integer
Public clipb As Integer




Private Sub File1_Click()
Text2.Text = File1.FileName

Call resetsample

Open App.Path + "\Patterns\" + Text2.Text For Input As #1
Do
Line Input #1, a$
List1.AddItem a$
Loop Until EOF(1)

Close #1

count2 = ((List1.ListCount) / 3)

Call loadnumbers

Label8.Caption = Str$(count2)

File1.Visible = False

Label2.Caption = "Saved.."
Label2.BackColor = &HFF00& 'green
pos3 = 9999
Display (0)
End Sub

Private Sub Form_Load()

List1.AddItem "        1   2   3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20"
List1.AddItem "      --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---"
List1.AddItem ""

For r = 0 To 39

Text1(r).Text = ""

Next

Label8.Caption = Str$(count2)
'Display2.Show

'Display2.Left = Sample.Left + Sample.Width
'Display2.Top = Sample.Top
'Display2.Height = Sample.Height


End Sub

Private Sub resetsample()

Label4.Caption = "1"
Sample.List1.Clear
clearnumbers
count2 = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)

Unload Display2
End Sub

Private Sub Label10_Click()

'pos3 = 0
'Display (0)




clipb = 1

pos3 = 0
Display (0)

'PicView.Picture = PicView.Image

'Clipboard.Clear
'Clipboard.SetData PicView.Picture, vbCFBitmap

'Clipboard.SetData picCircle.Picture, vbCFBitmap
clipb = 0

pos3 = 9999
Display (0)


End Sub

Private Sub Label2_Click()

If Trim$(Text2.Text) = "" Then MsgBox ("You must Enter File Name"): Exit Sub
If InStr(Text2.Text, ".") = False Then Text2.Text = Text2.Text + ".TXT"
If InStr(Text2.Text, ".") = True And InStr(UCase$(Text2.Text), ".TXT") = False Then
MsgBox ("You must Enter a .TXT File Name Extension"): Exit Sub
End If
On Local Error GoTo errors2
MkDir App.Path + "\Patterns"
On Local Error GoTo 0
Open App.Path + "\Patterns\" + Text2.Text For Output As #1
For r = 0 To List1.ListCount - 1
Print #1, List1.List(r)
Next

Close #1

Label2.Caption = "Saved.."
Label2.BackColor = &HFF00& 'green

Exit Sub
errors2:
Resume Next

End Sub

Private Sub loadnumbers()
change1 = 1
a1$ = List1.List(Val(Label4.Caption * 3))
a2$ = List1.List(Val(Label4.Caption * 3 + 1))
List1.ListIndex = Val(Label4.Caption * 3)

For r = 0 To 19

b$ = Mid$(a1$, (((r + 1) * 4) + 3), 3)
If b$ = " -1" Then b$ = " "
Text1(r).Text = Trim$(b$)

Next
For r = 20 To 39

b$ = Mid$(a2$, (((r - 19) * 4) + 3), 3)
If b$ = " -1" Then b$ = " "
Text1(r).Text = Trim$(b$)

Next

Label8.Caption = Str$(count2)
change1 = 0
End Sub
Private Sub clearnumbers()
change1 = 1
For r = 0 To 39

Text1(r).Text = ""

Next

Sample.Text1(0).SetFocus
change1 = 0
End Sub

Private Sub Label3_Click()
'Up

easy = 0
wd = 0

For r = 0 To 39
If Trim$(Text1(r).Text) <> "" Then wd = 1
Next

If wd = 1 Then Call listbox

If count2 < Val(Label4.Caption) Then count2 = Val(Label4.Caption)

If wd = 1 Then Label4.Caption = Str$(Val(Label4.Caption) + 1)

If count2 <= Val(Label4.Caption) Then Call clearnumbers

If count2 > Val(Label4.Caption) Then Call loadnumbers

Label8.Caption = Str$(count2)

End Sub

Private Sub Label4_Click()

easy = 2
wd = 0
For r = 0 To 39
If Trim$(Text1(r).Text) <> "" Then wd = 1
Next

If wd = 1 Then Call listbox

If count2 < Val(Label4.Caption) Then count2 = Val(Label4.Caption)

'Label4.Caption = Str$(Val(Label4.Caption) - 1)

'If count2 <= Val(Label4.Caption) Then Call clearnumbers

'If count2 > Val(Label4.Caption) Then Call loadnumbers

Call loadnumbers

End Sub

Private Sub Label5_Click()
'Down


easy = 1

wd = 0

For r = 0 To 39
If Trim$(Text1(r).Text) <> "" Then wd = 1
Next

If wd = 1 And Val(Label4.Caption) >= count2 Then Call Label3_Click: Exit Sub



If wd = 1 Then

Call listbox

End If

If count2 < Val(Label4.Caption) Then count2 = Val(Label4.Caption)

'If wd = 1 Then
If Val(Label4.Caption) > 1 Then
Label4.Caption = Str$(Val(Label4.Caption) - 1)
End If
'End If

If count2 <= Val(Label4.Caption) Then Call clearnumbers

If count2 > Val(Label4.Caption) Then Call loadnumbers

Label8.Caption = Str$(count2)


End Sub


Private Sub listbox()
patternx$ = Format$(Val(Label4.Caption), " 00") + " X "
patterny$ = Format$(Val(Label4.Caption), " 00") + " Y "
For r = 0 To 19
w = Val(Text1(r).Text)
If Trim$(Text1(r).Text) = "" Then w = -1
patt$ = Format$(w, "000")
If Mid$(patt$, 1, 1) = "0" Then
Mid$(patt$, 1, 1) = " "
If Mid$(patt$, 2, 1) = "0" Then Mid$(patt$, 2, 1) = " "
End If

If w > -1 And Text1(r).Text <> "--" Then patternx1$ = patternx1$ + patt$ + " "
If Text1(r).Text = "--" Then patternx1$ = patternx1$ + " -- "
Next
For r = 20 To 39
w = Val(Text1(r).Text)
If Trim$(Text1(r).Text) = "" Then w = -1
patt$ = Format$(w, "000")
If Mid$(patt$, 1, 1) = "0" Then
Mid$(patt$, 1, 1) = " "
If Mid$(patt$, 2, 1) = "0" Then Mid$(patt$, 2, 1) = " "
End If
'If w > -1 Then patterny1$ = patterny1$ + patt$ + " "
'If w = -1 Then patterny1$ = patterny1$ + "    "
If w > -1 And Text1(r).Text <> "--" Then patterny1$ = patterny1$ + patt$ + " "
If Text1(r).Text = "--" Then patterny1$ = patterny1$ + " -- "
Next

If Val(Label4.Caption) < count2 Then

List1.List(Val(Label4.Caption) * 3) = patternx$ + patternx1$
List1.List(Val(Label4.Caption) * 3 + 1) = patterny$ + patterny1$
List1.ListIndex = Val(Label4.Caption * 3)

End If

If Val(Label4.Caption) >= count2 And easy <> 1 Then

List1.AddItem patternx$ + patternx1$
List1.AddItem patterny$ + patterny1$
List1.AddItem ""

List1.ListIndex = List1.ListCount - 3

End If

If Val(Label4.Caption) < count2 Then List1.ListIndex = Val(Label4.Caption * 3) - 3
If Val(Label4.Caption) >= count2 - 1 Then List1.ListIndex = List1.ListCount - 1

List1.ListIndex = Val(Label4.Caption * 3)

Call Display(index2)


End Sub

Private Sub Label6_Click()
File1.Visible = True
Sample.File1.SetFocus
Sample.File1.Path = App.Path + "\patterns"

End Sub

Private Sub Label9_Click()

Call resetsample

count2 = 0

Call Form_Load
pos3 = 9999
Display (0)

End Sub



Private Sub Text1_Change(Index As Integer)
If Val(Text1(Index).Text) > 91 Then Text1(Index).Text = "91"

If Len(Text1(Index).Text) = 2 Then
If Index < 39 Then Sample.Text1(Index + 1).SetFocus
End If
index2 = Index
If change1 = 0 Then
Label2.Caption = "Save File"
Label2.BackColor = &HFF& 'red
End If
'Call Display(Index)

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If Index < 20 And KeyCode = 40 Then Sample.Text1(20).SetFocus
If Index >= 20 And KeyCode = 40 Then Call Label3_Click
If Index < 20 And KeyCode = 38 Then Call Label5_Click
If Index >= 20 And KeyCode = 38 Then Sample.Text1(Index - 20).SetFocus
If Index > 0 Then

'backspace
If KeyCode = 8 And Text1(Index).SelStart = 0 Then Sample.Text1(Index - 1).SetFocus: pos3 = 0

If KeyCode = 37 And Text1(Index).SelStart = 0 Then Sample.Text1(Index - 1).SetFocus
If KeyCode = 37 And Text1(Index).SelStart = 1 Then Sample.Text1(Index - 1).SetFocus
End If
If Index < 39 Then
If KeyCode = 39 And Text1(Index).SelStart = Len(Text1(Index).Text) Then Sample.Text1(Index + 1).SetFocus
End If

If Index >= 20 And KeyCode = 40 And Val(Label4.Caption) = count2 Then Exit Sub


Call Display(Index)
End Sub


Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 8 Then pos3 = 0
Call Display(Index)

End Sub


Private Sub Display(pos2 As Integer)

If pos2 + (Val(Label4.Caption)) = pos3 Then Exit Sub
pos3 = pos2 + (Val(Label4.Caption))

PicView.ForeColor = RGB(0, 0, 0)


Dim X As Integer
X = 0
    'Preparing the picturebox
    
    'Getting the selected FillMode
   ' For X = 0 To optFillMode.Count - 1
   '     If optFillMode(X).Value = True Then fMode = X: Exit For
   ' Next X





'Display2.Cls


If clipb = 1 Then PicView.AutoRedraw = False
  
  PicView.Cls
  
  PicView.AutoRedraw = True
  PicView.Cls
  PicView.ScaleMode = 3   ' Set ScaleMode to pixels.
  cx = PicView.ScaleWidth
  cy = PicView.ScaleHeight
  PicView.DrawWidth = 2   ' Set DrawWidth.
'  PicView.Refresh



weg = 1
ReDim Pntcut(1 To 30) As Long


'If Val(Label4.Caption) <> (((List1.ListCount - 3) - 1) / 3) Then
For r = 1 To ((List1.ListCount - 3) / 3)

If r <> Val(Label4.Caption) Then
a1$ = List1.List(r * 3)
a2$ = List1.List(r * 3 + 1)

For r1 = 0 To 19
xb$ = Mid$(a1$, (((r1 + 1) * 4) + 3), 3)
yb$ = Mid$(a2$, (((r1 + 1) * 4) + 3), 3)
If Trim$(xb$) <> "" And Trim$(yb$) <> "" Then X = X + 1
If Trim$(xb$) = "--" And Trim$(yb$) = "--" Then X = X + 1: cut = cut + 1: Pntcut(cut) = X

Next

End If

If r = Val(Label4.Caption) Then
For r2 = 0 To 19
wx = Val(Text1(r2).Text)
wy = Val(Text1(r2 + 20).Text)
If Trim$(Text1(r2).Text) <> "" And Trim$(Text1(r2 + 20).Text) <> "" Then X = X + 1
If Trim$(Text1(r2).Text) = "--" And Trim$(Text1(r2 + 20).Text) = "--" Then X = X + 1: cut = cut + 1: Pntcut(cut) = X

Next
End If

Next

'If Val(Label4.Caption) = ((List1.ListCount - 3) / 3) + 1 Then
'If Val(Label4.Caption) = (((List1.ListCount - 3) - 1) / 3) Then
'Display2.ForeColor = RGB(0, 255, 0)



'End If

cut = cut + 1
Pntcut(cut) = X
cut8 = 1

If X > 2 Then
X = X + 1
ReDim Pnt(1 To X) As POINTAPI
ReDim Pntcol(1 To X) As Long
ReDim Preserve Pntcut(1 To cut) As Long
ReDim circx(1 To cut + 1) As Double
ReDim circy(1 To cut + 1) As Double


For ri7 = 1 To cut + 1
circx(ri7) = -1: circy(ri7) = -1
Next


'    For X = 1 To UBound(Pnt)

X = 0
For r = 1 To (((List1.ListCount - 3) / 3))

If pos2 >= 20 Then pos2 = pos2 - 20


If r <> Val(Label4.Caption) Then

a1$ = List1.List(r * 3)
a2$ = List1.List(r * 3 + 1)

PicView.ForeColor = RGB(0, 0, 0)

If r = Val(Label4.Caption) Then
PicView.ForeColor = RGB(255, 0, 0)
Else
PicView.ForeColor = RGB(0, 0, 0)
End If
For r1 = 0 To 19

xb$ = Mid$(a1$, (((r1 + 1) * 4) + 3), 3)
yb$ = Mid$(a2$, (((r1 + 1) * 4) + 3), 3)
If r = Val(Label4.Caption) Then
xb$ = (Text1(r1).Text)
yb$ = (Text1(r1 + 20).Text)
End If
xx2 = Val(xb$)
yy2 = Val(yb$)
If Trim$(xb$) <> "" And Trim$(yb$) <> "" Then
x3 = (cx / 91) * xx2
Y3 = ((cy / 91) * (yy2))

'PicView.Line (xb3, cy - yb3)-(x3, cy - Y3)
        X = X + 1
If Pntcut(cut8) = X Then cut8 = cut8 + 1 ': weg = 1

'If weg = 1 Then weg = 0: xb3 = x3: Yb3 = Y3

'ReDim Pnt(1 To X) As POINTAPI
'ReDim Pnt2(1 To X) As POINTAPI
'ReDim Pntcol(1 To X) As Long
        
        Pnt(X).X = x3
        Pnt(X).Y = cy - Y3
        Pntcol(X) = PicView.ForeColor
        
    
  

If r1 = pos2 And r = Val(Label4.Caption) Then
'PicView.Circle (xb3, cy - yb3), 5, RGB(0, 0, 255)
circx(cut8) = x3
circy(cut8) = cy - Y3
End If

'xb3 = x3
'Yb3 = Y3
End If
Next
End If ' if r<>Val(Label4.Caption)




If r = Val(Label4.Caption) Then
PicView.ForeColor = RGB(0, 255, 0)

For r5 = 0 To 19
wx = Val(Text1(r5).Text)
wy = Val(Text1(r5 + 20).Text)
If Trim$(Text1(r5).Text) <> "" And Trim$(Text1(r5 + 20).Text) <> "" Then

x3 = (cx / 91) * wx
Y3 = ((cy / 91) * (wy))



'PicView.Line (xb3, cy - yb3)-(x3, cy - Y3)
        X = X + 1

If Pntcut(cut8) = X Then cut8 = cut8 + 1 ': weg = 1


'ReDim Pnt(1 To X) As POINTAPI
'ReDim Pnt2(1 To X) As POINTAPI
'ReDim Pntcol(1 To X) As Long
'If X = Pntcut(cut5) Then cut5 = cut5 + 1: weg = 1

'If weg = 1 Then weg = 0: xb3 = x3: Yb3 = Y3
        
        Pnt(X).X = x3
        Pnt(X).Y = cy - Y3
        Pntcol(X) = PicView.ForeColor
        
If r5 = pos2 Then
'PicView.Circle (xb3, cy - yb3), 5, RGB(0, 0, 255)



circx(cut8) = x3
circy(cut8) = cy - Y3

End If


'xb3 = x3
'Yb3 = Y3


End If

Next

End If 'if r=Val(Label4.Caption)

Next

 'Filling the defined region / polygon
    
xgag = 1
pattUpdate
xgag = 0

ratio = 0.96
ratio = 1

'size = size * 1
'size = 0.5

area = 1

size = size / area

If ratio < 1 Then pattw = size: patth = size * ratio
If ratio >= 1 Then patth = size: pattw = size / ratio

For i = 1 To UBound(Pnt) - 1
'xoffset = (Pnt(i).X)
'yoffset = (Pnt(i).Y)
centerx = cx / 2
centery = cy / 2


Pnt(i).X = ((Pnt(i).X - centerx) * pattw) + ((cx / 2))
Pnt(i).Y = ((Pnt(i).Y - centery) * patth) + ((cy / 2))

Next
'Pnt() = Pnt2()

pattUpdate


circx(cut8) = -1: circy(cut8) = -1


End If 'if x> 2




End Sub

Private Sub pattUpdate()

Dim fMode As Integer
fMode = 7
  
  cx = PicView.ScaleWidth
  cy = PicView.ScaleHeight

'Filling the defined region / polygon

For i5 = 1 To UBound(Pntcut)
op = 1
If i5 > 1 Then op = Pntcut(i5 - 1) + 1
If i5 > 1 Then
i5 = i5
xc = 1
End If
ReDim Pnt2((Pntcut(i5) - op) + xc * 2)
ReDim Pntcol2((Pntcut(i5) - op) + xc * 2)
For i6 = op To ((Pntcut(i5) - 1) + xc)
Pnt2(((i6 - op) + 1)).X = Pnt(i6 - xc).X
Pnt2(((i6 - op) + 1)).Y = Pnt(i6 - xc).Y
Pntcol2(((i6 - op) + 1)) = Pntcol(i6 - xc)
Next
If i5 = 1 Then colp = RGB(128, 255, 128)
If i5 = 2 Then colp = RGB(100, 150, 100)
If i5 = 3 Then colp = RGB(150, 140, 128)
If i5 = 4 Then colp = RGB(200, 100, 128)
If i5 = 5 Then colp = RGB(200, 100, 150)
If i5 = 6 Then colp = RGB(200, 100, 200)

If xgag = 1 Then
FillRegion PicView.hDC, Pnt2, RGB(128, 128, 128), fMode, picSource.Picture

If clipb = 1 Then
PicView.Picture = PicView.Image
Clipboard.Clear
Clipboard.SetData PicView.Picture, vbCFBitmap
End If
End If


If xgag = 0 Then FillRegion PicView.hDC, Pnt2, colp, fMode, picSource.Picture

If xgag = 1 Then
size = 0
size2 = 0
For tx1 = 0 To 91
For ty1 = 0 To 91
x5 = (cx / 91) * tx1
y5 = (cy / 91) * ty1
w = Sample.PicView.Point(x5, y5)
If w <> 16777215 Then size = size + 1
If w = 16777215 Then size2 = size2 + 1

Next
Next
'size = size / (91 * 91)
size = size / size2
End If

xb3 = Pnt2(1).X
Yb3 = Pnt2(1).Y

For X = 1 To UBound(Pnt2) - 1
 x3 = Pnt2(X).X
 Y3 = Pnt2(X).Y

PicView.ForeColor = Pntcol2(X)

If clipb = 0 Then PicView.Line (xb3, Yb3)-(x3, Y3), Pntcol2(X)
xb3 = x3
Yb3 = Y3
Next

If circx(i5) > -1 And circy(i5) > -1 Then
'PicView.FillColor = 0
If clipb = 0 Then PicView.FillStyle = 1
If clipb = 0 Then PicView.Circle (circx(i5), circy(i5)), 5, RGB(0, 0, 255)
End If
Next i5

PicView.Refresh

End Sub

