VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   Caption         =   "Pic Viewer Loader"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   Icon            =   "Shell Loader.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   645
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
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2355
      TabIndex        =   17
      Top             =   90
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   47
      Left            =   2880
      TabIndex        =   52
      Top             =   750
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   46
      Left            =   2880
      TabIndex        =   51
      Top             =   975
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   45
      Left            =   2880
      TabIndex        =   50
      Top             =   1185
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   44
      Left            =   2880
      TabIndex        =   49
      Top             =   1395
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   43
      Left            =   2880
      TabIndex        =   48
      Top             =   1605
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   42
      Left            =   2880
      TabIndex        =   47
      Top             =   1815
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   41
      Left            =   2880
      TabIndex        =   46
      Top             =   2025
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   40
      Left            =   2880
      TabIndex        =   45
      Top             =   2235
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   39
      Left            =   2865
      TabIndex        =   44
      Top             =   3915
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   38
      Left            =   2865
      TabIndex        =   43
      Top             =   3705
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   37
      Left            =   2865
      TabIndex        =   42
      Top             =   3495
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   36
      Left            =   2865
      TabIndex        =   41
      Top             =   3285
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   35
      Left            =   2865
      TabIndex        =   40
      Top             =   3075
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   34
      Left            =   2865
      TabIndex        =   39
      Top             =   2865
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   33
      Left            =   2865
      TabIndex        =   38
      Top             =   2655
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   32
      Left            =   2865
      TabIndex        =   37
      Top             =   2445
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   31
      Left            =   1635
      TabIndex        =   36
      Top             =   615
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   1635
      TabIndex        =   35
      Top             =   840
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   29
      Left            =   1635
      TabIndex        =   34
      Top             =   1050
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   28
      Left            =   1635
      TabIndex        =   33
      Top             =   1260
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   1635
      TabIndex        =   32
      Top             =   1470
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   1635
      TabIndex        =   31
      Top             =   1680
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   1635
      TabIndex        =   30
      Top             =   1890
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   1635
      TabIndex        =   29
      Top             =   2100
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   1620
      TabIndex        =   28
      Top             =   3780
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   1620
      TabIndex        =   27
      Top             =   3570
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   1620
      TabIndex        =   26
      Top             =   3360
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   1620
      TabIndex        =   25
      Top             =   3150
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   1620
      TabIndex        =   24
      Top             =   2940
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   1620
      TabIndex        =   23
      Top             =   2730
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   1620
      TabIndex        =   22
      Top             =   2520
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   1620
      TabIndex        =   21
      Top             =   2310
      Width           =   1545
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
      Caption         =   "Pic Viewer Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   4455
      TabIndex        =   18
      Top             =   105
      Width           =   2115
   End
   Begin VB.Label TitleLbl 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "*** -- Select One -- ***"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   135
      TabIndex        =   16
      Top             =   75
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   360
      TabIndex        =   15
      Top             =   2505
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   360
      TabIndex        =   14
      Top             =   2715
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   360
      TabIndex        =   13
      Top             =   2925
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   360
      TabIndex        =   12
      Top             =   3135
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   360
      TabIndex        =   11
      Top             =   3345
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   360
      TabIndex        =   10
      Top             =   3555
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   360
      TabIndex        =   9
      Top             =   3765
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   360
      TabIndex        =   8
      Top             =   3975
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   375
      TabIndex        =   7
      Top             =   2295
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   375
      TabIndex        =   6
      Top             =   2085
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   375
      TabIndex        =   5
      Top             =   1875
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   375
      TabIndex        =   4
      Top             =   1665
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   375
      TabIndex        =   3
      Top             =   1455
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   375
      TabIndex        =   2
      Top             =   1245
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   375
      TabIndex        =   1
      Top             =   1035
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Loader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   375
      TabIndex        =   0
      Top             =   810
      Width           =   1545
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long


Private Sub Form_Load()


'gcw2 = GetForegroundWindow
Winamp22Handle2 = FindWindow("Winamp v1.x", vbNullString)
Winamp22Handle3 = FindWindow("Winamp PE", vbNullString)
Winamp22Handle4 = FindWindow("Winamp Video", vbNullString)
If Winamp22Handle2 > 0 Then w2h = Winamp22Handle2
If Winamp22Handle3 > 0 Then w2h = Winamp22Handle3
If Winamp22Handle4 > 0 Then w2h = Winamp22Handle4
         
'If w2h <> gcw2 Then Exit Sub
        
Gcw$ = GetWindowTitle(w2h)
         
         
Dim Tex$
Dim Pid As Long
         
Tex$ = GetFileFromHwnd(w2h)
'If InStr(Tex$, "Video X") Then Call MuteVol
        
'Lbl2.Caption = Tex$
If Tex$ <> "" Then
gg$ = Mid$(Tex$, InStrRev(Tex$, "\", Len(Tex$) - 11) + 1)
gg$ = Mid$(gg$, 1, InStrRev(gg$, "\") - 1)
Else
gg$ = "No Winamp Loaded"
End If

'Lbl2.Caption = gg$


Form1.Top = 480
Form1.Left = 0

'Dir1.Path = "C:\Program Files\00 WinAmp's"
'Dir1.Path = "E:\01 VB Shell Folders\00 Shell Notepad Launcher"
'File1.Path = "E:\01 VB Shell Folders\00 Shell Notepad Launcher"

'ScanPath.txtPath.Text = "E:\01 VB Shell Folders\00 Shell Excel Launcher"

'ScanPath.cboMask.Text = "*.lnk"

'Call ScanPath.cmdScan_Click


Lbl2.Top = 0
Lbl2.Left = 0
Lbl2.Width = Form1.Width


x = Lbl2.Height + 15

TitleLbl.Top = x
TitleLbl.Left = 0
TitleLbl.Width = Form1.Width
Check1.Top = x
Check1.Left = (Form1.Width / 2) + 10
Check1.Width = Form1.Width / 2

x = x + TitleLbl.Height + 15



w = 0

rd = 0


rg = 0
For Each Control In Controls
If InStr(Control.Name, "Lab") Then rg = rg + 1: Control.Visible = False
Next

ReDim G1(50)

hh = 0

'hh = hh + 1
'G1(hh) = ""

'hh = hh + 1
'G1(hh) = ""

hh = hh + 1
G1(hh) = "S:\00 Art\00 My Pictures\z4Cosmi\5057 Photos"

hh = hh + 1
G1(hh) = "S:\00 Art\00 My Pictures\z3NaSa Images"

hh = hh + 1
G1(hh) = "S:\00 Art\00 My Pictures\z2Musée d'orsay"

hh = hh + 1
G1(hh) = "S:\00 Art\00 My Pictures\z1Female Pictures"

hh = hh + 1
G1(hh) = "S:\00 Art\00 My Pictures\z1JPG Dos6.22 Days"

hh = hh + 1
G1(hh) = "S:\00 Art\00 My Pictures\Fractals"

hh = hh + 1
G1(hh) = "S:\00 Art\00 My Pictures\Fractals\Fractals 2002-3"

hh = hh + 1
G1(hh) = "S:\00 Art\00 My Pictures\Fractals\Fractals 2004"

hh = hh + 1
G1(hh) = "S:\00 Art\00 My Pictures\Fractals\Fractals 2005"

hh = hh + 1
G1(hh) = "S:\00 Art\00 My Pictures\Fractals\Fractals 2006"

hh = hh + 1
G1(hh) = "S:\00 Art\00 My Pictures\Fractals\Fractals 2007"

hh = hh + 1
G1(hh) = "S:\00 Art\00 My Pictures\Bill Hick's"

hh = hh + 1
G1(hh) = "S:\00 Art\00 My Pictures\03 DJ's Banging Tunes.Com"

hh = hh + 1
G1(hh) = "D:\# MY DOCS\# 01 My Documents\Blue Tooth Exchange Folder\01 All Phone"

hh = hh + 1
G1(hh) = "D:\# MY DOCS\# 01 My Documents\Blue Tooth Exchange Folder"

hh = hh + 1
G1(hh) = "D:\# MY DOCS\# 01 My Documents\Blue Tooth Exchange Folder\01 All Phone\00 Mental\00 00 Hoss"

hh = hh + 1
G1(hh) = "D:\# MY DOCS\# 01 My Documents\Blue Tooth Exchange Folder\01 All Phone\00 Mental\00 010 Oct"

hh = hh + 1
G1(hh) = "D:\# MY DOCS\# 01 My Documents\Blue Tooth Exchange Folder\01 All Phone\00 Mental\00 011 Nov"

hh = hh + 1
G1(hh) = "D:\# MY DOCS\# 01 My Documents\Blue Tooth Exchange Folder\01 All Phone\00 Mental\00 012 Dec"

hh = hh + 1
G1(hh) = "D:\# MY DOCS\# 01 My Documents\Blue Tooth Exchange Folder\01 All Phone\00 Mental\01 001 Jan 2008"

hh = hh + 1
G1(hh) = "D:\# MY DOCS\# 01 My Documents\Blue Tooth Exchange Folder\01 All Phone\I.N. Galidakis"

hh = hh + 1
G1(hh) = "D:\# MY DOCS\# 01 My Documents\Blue Tooth Exchange Folder\01 All Phone\VBDos"

hh = hh + 1
G1(hh) = "D:\# MY DOCS\# 01 My Documents\Blue Tooth Exchange Folder\01 All Phone\mnftiu.cc"

hh = hh + 1
G1(hh) = "D:\# MY DOCS\# 01 My Documents\Blue Tooth Exchange Folder\01 All Phone\mnftiu.cc\Fighting"

hh = hh + 1
G1(hh) = "D:\# MY DOCS\# 01 My Documents\Blue Tooth Exchange Folder\01 All Phone\mnftiu.cc\Filing"

hh = hh + 1
G1(hh) = "S:\00 Art\AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0000-Ultra Sexi"
hh = hh + 1
G1(hh) = "S:\00 Art\AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0000-Sexyest"
hh = hh + 1
G1(hh) = "S:\00 Art\AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0000-Mags"
hh = hh + 1
G1(hh) = "S:\00 Art\AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0000-Cartoon"
hh = hh + 1
G1(hh) = "S:\00 Art\AutoPix\AutoPix (Main)\AutoPix 00\AutoPix 0000-Vintage"
hh = hh + 1
G1(hh) = "S:\00 Art\AutoPix\AutoPix (Main)\AutoPix 2007 05"

ii = hh
ReDim Preserve G1(ii)


'For we = 1 To ListView1.ListItems.Count
'    a1$ = ScanPath.ListView1.ListItems.Item(we).SubItems(1)
'    b1$ = ScanPath.ListView1.ListItems.Item(we)

For r = 1 To UBound(G1)
    Label1(r).Top = x
    Label1(r).Left = w
    Label1(r).Height = 270
    Label1(r).Width = Form1.Width
    Label1(r).Visible = True
    
    'ttg$ = Mid$(File1.List(rd), InStrRev(File1.List(rd), "\") + 1)
    ttg$ = G1(r)
    ttg$ = Mid$(ttg$, InStrRev(ttg$, "\", Len(ttg$) - 2) + 1)
    
    Label1(r).Caption = Format$(r, "00") + ". " + ttg$
    x = x + Label1(r).Height + 15
    fheight = Label1(r).Top + Label1(r).Height + 420
    rd = rd + 1
Next

Form1.Height = fheight
Lbl3.Top = fheight - 410
Lbl3.Width = Form1.Width
Lbl3.Left = 0

'List1.Top = fheight - 410 + Lbl3.Height
'List1.Height = 1500
'List1.Width = Form1.Width - 120
'List1.Left = 0

'Form1.Height = List1.Top + List1.Height + 410
'Form1.Height = List1.Top + List1.Height + 410

'If Dir$(App.Path + "\Winloads.txt") = "" Then Exit Sub
'Open App.Path + "\Winloads.txt" For Input As #1
'On Local Error Resume Next
'Seek #1, LOF(1) - 500
'Do
'Line Input #1, ed$
'List1.AddItem ed$
'Loop Until EOF(1)
'Close #1

'List1.AddItem Format$(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " -- " + Lbl2.Caption
'Open App.Path + "\Winloads.txt" For Append As #1
'Print #1, Format$(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " -- " + Lbl2.Caption
'Close #1

'List1.ListIndex = List1.ListCount - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'End
End Sub

Private Sub Label1_Click(Index As Integer)

MpathX = G1(Index)

Load frmPicViewer


'If Check1.Value = vbUnchecked Then
'Open App.Path + "\Winloads.txt" For Append As #1
'Print #1, Format$(Now, "DDD DD-MMM-YY HH:MM:SSa/p") + " -- " + Label1(Index)
'Close #1
'Shell Dir1.List(Index) + "\Winamp.exe", vbNormalFocus
'a1$ = ScanPath.ListView1.ListItems.Item(Index).SubItems(1)
'b1$ = ScanPath.ListView1.ListItems.Item(Index)



'vLaunch a1$ + b1$



'End

'Else
'Shell "notepad " + Dir1.List(Index) + "\Current Play To Text File Append.txt", vbNormalFocus

'End

'End If

'End


End Sub

