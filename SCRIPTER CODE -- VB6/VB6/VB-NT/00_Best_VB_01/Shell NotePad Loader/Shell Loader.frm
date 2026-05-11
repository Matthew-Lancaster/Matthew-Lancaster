VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   Caption         =   "NotePad Loader"
   ClientHeight    =   5925
   ClientLeft      =   225
   ClientTop       =   1095
   ClientWidth     =   12330
   Icon            =   "Shell Loader.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2208
      Top             =   2016
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
      Height          =   450
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
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   752
      Left            =   7245
      TabIndex        =   758
      Top             =   960
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   751
      Left            =   7245
      TabIndex        =   757
      Top             =   1185
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   750
      Left            =   7245
      TabIndex        =   756
      Top             =   1395
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   749
      Left            =   7245
      TabIndex        =   755
      Top             =   1605
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   748
      Left            =   7245
      TabIndex        =   754
      Top             =   1815
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   747
      Left            =   7245
      TabIndex        =   753
      Top             =   2025
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   746
      Left            =   7245
      TabIndex        =   752
      Top             =   2235
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   745
      Left            =   7245
      TabIndex        =   751
      Top             =   2445
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   744
      Left            =   7230
      TabIndex        =   750
      Top             =   4125
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   743
      Left            =   7230
      TabIndex        =   749
      Top             =   3915
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   742
      Left            =   7230
      TabIndex        =   748
      Top             =   3705
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   741
      Left            =   7230
      TabIndex        =   747
      Top             =   3495
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   740
      Left            =   7230
      TabIndex        =   746
      Top             =   3285
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   739
      Left            =   7230
      TabIndex        =   745
      Top             =   3075
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   738
      Left            =   7230
      TabIndex        =   744
      Top             =   2865
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   737
      Left            =   7230
      TabIndex        =   743
      Top             =   2655
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   736
      Left            =   7890
      TabIndex        =   742
      Top             =   2775
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   735
      Left            =   7890
      TabIndex        =   741
      Top             =   2985
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   734
      Left            =   7890
      TabIndex        =   740
      Top             =   3195
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   733
      Left            =   7890
      TabIndex        =   739
      Top             =   3405
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   732
      Left            =   7890
      TabIndex        =   738
      Top             =   3615
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   731
      Left            =   7890
      TabIndex        =   737
      Top             =   3825
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   730
      Left            =   7890
      TabIndex        =   736
      Top             =   4035
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   729
      Left            =   7890
      TabIndex        =   735
      Top             =   4245
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   728
      Left            =   7905
      TabIndex        =   734
      Top             =   2565
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   727
      Left            =   7905
      TabIndex        =   733
      Top             =   2355
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   726
      Left            =   7905
      TabIndex        =   732
      Top             =   2145
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   725
      Left            =   7905
      TabIndex        =   731
      Top             =   1935
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   724
      Left            =   7905
      TabIndex        =   730
      Top             =   1725
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   723
      Left            =   7905
      TabIndex        =   729
      Top             =   1515
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   722
      Left            =   7905
      TabIndex        =   728
      Top             =   1305
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   721
      Left            =   7905
      TabIndex        =   727
      Top             =   1080
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   720
      Left            =   8100
      TabIndex        =   726
      Top             =   3180
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   719
      Left            =   8100
      TabIndex        =   725
      Top             =   3390
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   718
      Left            =   8100
      TabIndex        =   724
      Top             =   3600
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   717
      Left            =   8100
      TabIndex        =   723
      Top             =   3810
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   716
      Left            =   8100
      TabIndex        =   722
      Top             =   4020
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   715
      Left            =   8100
      TabIndex        =   721
      Top             =   4230
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   714
      Left            =   8100
      TabIndex        =   720
      Top             =   4440
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   713
      Left            =   8100
      TabIndex        =   719
      Top             =   4650
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   712
      Left            =   8115
      TabIndex        =   718
      Top             =   2970
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   711
      Left            =   8115
      TabIndex        =   717
      Top             =   2760
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   710
      Left            =   8115
      TabIndex        =   716
      Top             =   2550
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   709
      Left            =   8115
      TabIndex        =   715
      Top             =   2340
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   708
      Left            =   8115
      TabIndex        =   714
      Top             =   2130
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   707
      Left            =   8115
      TabIndex        =   713
      Top             =   1920
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   706
      Left            =   8115
      TabIndex        =   712
      Top             =   1710
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   705
      Left            =   8115
      TabIndex        =   711
      Top             =   1485
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   704
      Left            =   9825
      TabIndex        =   710
      Top             =   1485
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   703
      Left            =   9825
      TabIndex        =   709
      Top             =   1710
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   702
      Left            =   9825
      TabIndex        =   708
      Top             =   1920
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   701
      Left            =   9825
      TabIndex        =   707
      Top             =   2130
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   700
      Left            =   9825
      TabIndex        =   706
      Top             =   2340
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   699
      Left            =   9825
      TabIndex        =   705
      Top             =   2550
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   698
      Left            =   9825
      TabIndex        =   704
      Top             =   2760
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   697
      Left            =   9825
      TabIndex        =   703
      Top             =   2970
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   696
      Left            =   9810
      TabIndex        =   702
      Top             =   4650
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   695
      Left            =   9810
      TabIndex        =   701
      Top             =   4440
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   694
      Left            =   9810
      TabIndex        =   700
      Top             =   4230
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   693
      Left            =   9810
      TabIndex        =   699
      Top             =   4020
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   692
      Left            =   9810
      TabIndex        =   698
      Top             =   3810
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   691
      Left            =   9810
      TabIndex        =   697
      Top             =   3600
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   690
      Left            =   9810
      TabIndex        =   696
      Top             =   3390
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   689
      Left            =   9810
      TabIndex        =   695
      Top             =   3180
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   688
      Left            =   9150
      TabIndex        =   694
      Top             =   3060
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   687
      Left            =   9150
      TabIndex        =   693
      Top             =   3270
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   686
      Left            =   9150
      TabIndex        =   692
      Top             =   3480
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   685
      Left            =   9150
      TabIndex        =   691
      Top             =   3690
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   684
      Left            =   9150
      TabIndex        =   690
      Top             =   3900
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   683
      Left            =   9150
      TabIndex        =   689
      Top             =   4110
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   682
      Left            =   9150
      TabIndex        =   688
      Top             =   4320
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   681
      Left            =   9150
      TabIndex        =   687
      Top             =   4530
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   680
      Left            =   9165
      TabIndex        =   686
      Top             =   2850
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   679
      Left            =   9165
      TabIndex        =   685
      Top             =   2640
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   678
      Left            =   9165
      TabIndex        =   684
      Top             =   2430
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   677
      Left            =   9165
      TabIndex        =   683
      Top             =   2220
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   676
      Left            =   9165
      TabIndex        =   682
      Top             =   2010
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   675
      Left            =   9165
      TabIndex        =   681
      Top             =   1800
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   674
      Left            =   9165
      TabIndex        =   680
      Top             =   1590
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   673
      Left            =   9165
      TabIndex        =   679
      Top             =   1365
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   672
      Left            =   7365
      TabIndex        =   678
      Top             =   3090
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   671
      Left            =   7365
      TabIndex        =   677
      Top             =   3300
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   670
      Left            =   7365
      TabIndex        =   676
      Top             =   3510
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   669
      Left            =   7365
      TabIndex        =   675
      Top             =   3720
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   668
      Left            =   7365
      TabIndex        =   674
      Top             =   3930
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   667
      Left            =   7365
      TabIndex        =   673
      Top             =   4140
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   666
      Left            =   7365
      TabIndex        =   672
      Top             =   4350
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   665
      Left            =   7365
      TabIndex        =   671
      Top             =   4560
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   664
      Left            =   7380
      TabIndex        =   670
      Top             =   2880
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   663
      Left            =   7380
      TabIndex        =   669
      Top             =   2670
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   662
      Left            =   7380
      TabIndex        =   668
      Top             =   2460
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   661
      Left            =   7380
      TabIndex        =   667
      Top             =   2250
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   660
      Left            =   7380
      TabIndex        =   666
      Top             =   2040
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   659
      Left            =   7380
      TabIndex        =   665
      Top             =   1830
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   658
      Left            =   7380
      TabIndex        =   664
      Top             =   1620
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   657
      Left            =   7380
      TabIndex        =   663
      Top             =   1395
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   656
      Left            =   9360
      TabIndex        =   662
      Top             =   1140
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   655
      Left            =   9360
      TabIndex        =   661
      Top             =   1365
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   654
      Left            =   9360
      TabIndex        =   660
      Top             =   1575
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   653
      Left            =   9360
      TabIndex        =   659
      Top             =   1785
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   652
      Left            =   9360
      TabIndex        =   658
      Top             =   1995
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   651
      Left            =   9360
      TabIndex        =   657
      Top             =   2205
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   650
      Left            =   9360
      TabIndex        =   656
      Top             =   2415
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   649
      Left            =   9360
      TabIndex        =   655
      Top             =   2625
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   648
      Left            =   9345
      TabIndex        =   654
      Top             =   4305
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   647
      Left            =   9345
      TabIndex        =   653
      Top             =   4095
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   646
      Left            =   9345
      TabIndex        =   652
      Top             =   3885
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   645
      Left            =   9345
      TabIndex        =   651
      Top             =   3675
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   644
      Left            =   9345
      TabIndex        =   650
      Top             =   3465
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   643
      Left            =   9345
      TabIndex        =   649
      Top             =   3255
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   642
      Left            =   9345
      TabIndex        =   648
      Top             =   3045
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   641
      Left            =   9345
      TabIndex        =   647
      Top             =   2835
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   640
      Left            =   10005
      TabIndex        =   646
      Top             =   2955
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   639
      Left            =   10005
      TabIndex        =   645
      Top             =   3165
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   638
      Left            =   10005
      TabIndex        =   644
      Top             =   3375
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   637
      Left            =   10005
      TabIndex        =   643
      Top             =   3585
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   636
      Left            =   10005
      TabIndex        =   642
      Top             =   3795
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   635
      Left            =   10005
      TabIndex        =   641
      Top             =   4005
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   634
      Left            =   10005
      TabIndex        =   640
      Top             =   4215
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   633
      Left            =   10005
      TabIndex        =   639
      Top             =   4425
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   632
      Left            =   10020
      TabIndex        =   638
      Top             =   2745
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   631
      Left            =   10020
      TabIndex        =   637
      Top             =   2535
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   630
      Left            =   10020
      TabIndex        =   636
      Top             =   2325
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   629
      Left            =   10020
      TabIndex        =   635
      Top             =   2115
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   628
      Left            =   10020
      TabIndex        =   634
      Top             =   1905
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   627
      Left            =   10020
      TabIndex        =   633
      Top             =   1695
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   626
      Left            =   10020
      TabIndex        =   632
      Top             =   1485
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   625
      Left            =   10020
      TabIndex        =   631
      Top             =   1260
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   624
      Left            =   8310
      TabIndex        =   630
      Top             =   1260
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   623
      Left            =   8310
      TabIndex        =   629
      Top             =   1485
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   622
      Left            =   8310
      TabIndex        =   628
      Top             =   1695
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   621
      Left            =   8310
      TabIndex        =   627
      Top             =   1905
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   620
      Left            =   8310
      TabIndex        =   626
      Top             =   2115
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   619
      Left            =   8310
      TabIndex        =   625
      Top             =   2325
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   618
      Left            =   8310
      TabIndex        =   624
      Top             =   2535
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   617
      Left            =   8310
      TabIndex        =   623
      Top             =   2745
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   616
      Left            =   8295
      TabIndex        =   622
      Top             =   4425
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   615
      Left            =   8295
      TabIndex        =   621
      Top             =   4215
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   614
      Left            =   8295
      TabIndex        =   620
      Top             =   4005
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   613
      Left            =   8295
      TabIndex        =   619
      Top             =   3795
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   612
      Left            =   8295
      TabIndex        =   618
      Top             =   3585
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   611
      Left            =   8295
      TabIndex        =   617
      Top             =   3375
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   610
      Left            =   8295
      TabIndex        =   616
      Top             =   3165
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   609
      Left            =   8295
      TabIndex        =   615
      Top             =   2955
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   608
      Left            =   8100
      TabIndex        =   614
      Top             =   855
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   607
      Left            =   8100
      TabIndex        =   613
      Top             =   1080
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   606
      Left            =   8100
      TabIndex        =   612
      Top             =   1290
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   605
      Left            =   8100
      TabIndex        =   611
      Top             =   1500
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   604
      Left            =   8100
      TabIndex        =   610
      Top             =   1710
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   603
      Left            =   8100
      TabIndex        =   609
      Top             =   1920
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   602
      Left            =   8100
      TabIndex        =   608
      Top             =   2130
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   601
      Left            =   8100
      TabIndex        =   607
      Top             =   2340
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   600
      Left            =   8085
      TabIndex        =   606
      Top             =   4020
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   599
      Left            =   8085
      TabIndex        =   605
      Top             =   3810
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   598
      Left            =   8085
      TabIndex        =   604
      Top             =   3600
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   597
      Left            =   8085
      TabIndex        =   603
      Top             =   3390
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   596
      Left            =   8085
      TabIndex        =   602
      Top             =   3180
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   595
      Left            =   8085
      TabIndex        =   601
      Top             =   2970
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   594
      Left            =   8085
      TabIndex        =   600
      Top             =   2760
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   593
      Left            =   8085
      TabIndex        =   599
      Top             =   2550
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   592
      Left            =   6960
      TabIndex        =   598
      Top             =   2535
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   591
      Left            =   6960
      TabIndex        =   597
      Top             =   2745
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   590
      Left            =   6960
      TabIndex        =   596
      Top             =   2955
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   589
      Left            =   6960
      TabIndex        =   595
      Top             =   3165
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   588
      Left            =   6960
      TabIndex        =   594
      Top             =   3375
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   587
      Left            =   6960
      TabIndex        =   593
      Top             =   3585
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   586
      Left            =   6960
      TabIndex        =   592
      Top             =   3795
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   585
      Left            =   6960
      TabIndex        =   591
      Top             =   4005
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   584
      Left            =   6975
      TabIndex        =   590
      Top             =   2325
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   583
      Left            =   6975
      TabIndex        =   589
      Top             =   2115
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   582
      Left            =   6975
      TabIndex        =   588
      Top             =   1905
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   581
      Left            =   6975
      TabIndex        =   587
      Top             =   1695
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   580
      Left            =   6975
      TabIndex        =   586
      Top             =   1485
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   579
      Left            =   6975
      TabIndex        =   585
      Top             =   1275
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   578
      Left            =   6975
      TabIndex        =   584
      Top             =   1065
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   577
      Left            =   6975
      TabIndex        =   583
      Top             =   840
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   576
      Left            =   7170
      TabIndex        =   582
      Top             =   2940
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   575
      Left            =   7170
      TabIndex        =   581
      Top             =   3150
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   574
      Left            =   7170
      TabIndex        =   580
      Top             =   3360
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   573
      Left            =   7170
      TabIndex        =   579
      Top             =   3570
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   572
      Left            =   7170
      TabIndex        =   578
      Top             =   3780
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   571
      Left            =   7170
      TabIndex        =   577
      Top             =   3990
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   570
      Left            =   7170
      TabIndex        =   576
      Top             =   4200
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   569
      Left            =   7170
      TabIndex        =   575
      Top             =   4410
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   568
      Left            =   7185
      TabIndex        =   574
      Top             =   2730
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   567
      Left            =   7185
      TabIndex        =   573
      Top             =   2520
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   566
      Left            =   7185
      TabIndex        =   572
      Top             =   2310
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   565
      Left            =   7185
      TabIndex        =   571
      Top             =   2100
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   564
      Left            =   7185
      TabIndex        =   570
      Top             =   1890
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   563
      Left            =   7185
      TabIndex        =   569
      Top             =   1680
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   562
      Left            =   7185
      TabIndex        =   568
      Top             =   1470
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   561
      Left            =   7185
      TabIndex        =   567
      Top             =   1245
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   560
      Left            =   8895
      TabIndex        =   566
      Top             =   1245
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   559
      Left            =   8895
      TabIndex        =   565
      Top             =   1470
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   558
      Left            =   8895
      TabIndex        =   564
      Top             =   1680
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   557
      Left            =   8895
      TabIndex        =   563
      Top             =   1890
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   556
      Left            =   8895
      TabIndex        =   562
      Top             =   2100
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   555
      Left            =   8895
      TabIndex        =   561
      Top             =   2310
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   554
      Left            =   8895
      TabIndex        =   560
      Top             =   2520
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   553
      Left            =   8895
      TabIndex        =   559
      Top             =   2730
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   552
      Left            =   8880
      TabIndex        =   558
      Top             =   4410
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   551
      Left            =   8880
      TabIndex        =   557
      Top             =   4200
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   550
      Left            =   8880
      TabIndex        =   556
      Top             =   3990
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   549
      Left            =   8880
      TabIndex        =   555
      Top             =   3780
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   548
      Left            =   8880
      TabIndex        =   554
      Top             =   3570
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   547
      Left            =   8880
      TabIndex        =   553
      Top             =   3360
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   546
      Left            =   8880
      TabIndex        =   552
      Top             =   3150
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   545
      Left            =   8880
      TabIndex        =   551
      Top             =   2940
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   544
      Left            =   8220
      TabIndex        =   550
      Top             =   2820
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   543
      Left            =   8220
      TabIndex        =   549
      Top             =   3030
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   542
      Left            =   8220
      TabIndex        =   548
      Top             =   3240
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   541
      Left            =   8220
      TabIndex        =   547
      Top             =   3450
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   540
      Left            =   8220
      TabIndex        =   546
      Top             =   3660
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   539
      Left            =   8220
      TabIndex        =   545
      Top             =   3870
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   538
      Left            =   8220
      TabIndex        =   544
      Top             =   4080
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   537
      Left            =   8220
      TabIndex        =   543
      Top             =   4290
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   536
      Left            =   8235
      TabIndex        =   542
      Top             =   2610
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   535
      Left            =   8235
      TabIndex        =   541
      Top             =   2400
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   534
      Left            =   8235
      TabIndex        =   540
      Top             =   2190
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   533
      Left            =   8235
      TabIndex        =   539
      Top             =   1980
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   532
      Left            =   8235
      TabIndex        =   538
      Top             =   1770
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   531
      Left            =   8235
      TabIndex        =   537
      Top             =   1560
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   530
      Left            =   8235
      TabIndex        =   536
      Top             =   1350
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   529
      Left            =   8235
      TabIndex        =   535
      Top             =   1125
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   528
      Left            =   8040
      TabIndex        =   534
      Top             =   1350
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   527
      Left            =   8040
      TabIndex        =   533
      Top             =   1575
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   526
      Left            =   8040
      TabIndex        =   532
      Top             =   1785
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   525
      Left            =   8040
      TabIndex        =   531
      Top             =   1995
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   524
      Left            =   8040
      TabIndex        =   530
      Top             =   2205
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   523
      Left            =   8040
      TabIndex        =   529
      Top             =   2415
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   522
      Left            =   8040
      TabIndex        =   528
      Top             =   2625
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   521
      Left            =   8040
      TabIndex        =   527
      Top             =   2835
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   520
      Left            =   8025
      TabIndex        =   526
      Top             =   4515
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   519
      Left            =   8025
      TabIndex        =   525
      Top             =   4305
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   518
      Left            =   8025
      TabIndex        =   524
      Top             =   4095
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   517
      Left            =   8025
      TabIndex        =   523
      Top             =   3885
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   516
      Left            =   8025
      TabIndex        =   522
      Top             =   3675
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   515
      Left            =   8025
      TabIndex        =   521
      Top             =   3465
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   514
      Left            =   8025
      TabIndex        =   520
      Top             =   3255
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   513
      Left            =   8025
      TabIndex        =   519
      Top             =   3045
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   512
      Left            =   8685
      TabIndex        =   518
      Top             =   3165
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   511
      Left            =   8685
      TabIndex        =   517
      Top             =   3375
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   510
      Left            =   8685
      TabIndex        =   516
      Top             =   3585
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   509
      Left            =   8685
      TabIndex        =   515
      Top             =   3795
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   508
      Left            =   8685
      TabIndex        =   514
      Top             =   4005
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   507
      Left            =   8685
      TabIndex        =   513
      Top             =   4215
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   506
      Left            =   8685
      TabIndex        =   512
      Top             =   4425
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   505
      Left            =   8685
      TabIndex        =   511
      Top             =   4635
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   504
      Left            =   8700
      TabIndex        =   510
      Top             =   2955
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   503
      Left            =   8700
      TabIndex        =   509
      Top             =   2745
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   502
      Left            =   8700
      TabIndex        =   508
      Top             =   2535
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   501
      Left            =   8700
      TabIndex        =   507
      Top             =   2325
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   500
      Left            =   8700
      TabIndex        =   506
      Top             =   2115
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   499
      Left            =   8700
      TabIndex        =   505
      Top             =   1905
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   498
      Left            =   8700
      TabIndex        =   504
      Top             =   1695
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   497
      Left            =   8700
      TabIndex        =   503
      Top             =   1470
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   496
      Left            =   6990
      TabIndex        =   502
      Top             =   1470
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   495
      Left            =   6990
      TabIndex        =   501
      Top             =   1695
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   494
      Left            =   6990
      TabIndex        =   500
      Top             =   1905
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   493
      Left            =   6990
      TabIndex        =   499
      Top             =   2115
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   492
      Left            =   6990
      TabIndex        =   498
      Top             =   2325
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   491
      Left            =   6990
      TabIndex        =   497
      Top             =   2535
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   490
      Left            =   6990
      TabIndex        =   496
      Top             =   2745
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   489
      Left            =   6990
      TabIndex        =   495
      Top             =   2955
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   488
      Left            =   6975
      TabIndex        =   494
      Top             =   4635
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   487
      Left            =   6975
      TabIndex        =   493
      Top             =   4425
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   486
      Left            =   6975
      TabIndex        =   492
      Top             =   4215
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   485
      Left            =   6975
      TabIndex        =   491
      Top             =   4005
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   484
      Left            =   6975
      TabIndex        =   490
      Top             =   3795
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   483
      Left            =   6975
      TabIndex        =   489
      Top             =   3585
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   482
      Left            =   6975
      TabIndex        =   488
      Top             =   3375
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   481
      Left            =   6975
      TabIndex        =   487
      Top             =   3165
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   480
      Left            =   6795
      TabIndex        =   486
      Top             =   705
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   479
      Left            =   6795
      TabIndex        =   485
      Top             =   930
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   478
      Left            =   6795
      TabIndex        =   484
      Top             =   1140
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   477
      Left            =   6795
      TabIndex        =   483
      Top             =   1350
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   476
      Left            =   6795
      TabIndex        =   482
      Top             =   1560
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   475
      Left            =   6795
      TabIndex        =   481
      Top             =   1770
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   474
      Left            =   6795
      TabIndex        =   480
      Top             =   1980
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   473
      Left            =   6795
      TabIndex        =   479
      Top             =   2190
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   472
      Left            =   6780
      TabIndex        =   478
      Top             =   3870
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   471
      Left            =   6780
      TabIndex        =   477
      Top             =   3660
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   470
      Left            =   6780
      TabIndex        =   476
      Top             =   3450
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   469
      Left            =   6780
      TabIndex        =   475
      Top             =   3240
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   468
      Left            =   6780
      TabIndex        =   474
      Top             =   3030
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   467
      Left            =   6780
      TabIndex        =   473
      Top             =   2820
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   466
      Left            =   6780
      TabIndex        =   472
      Top             =   2610
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   465
      Left            =   6780
      TabIndex        =   471
      Top             =   2400
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   464
      Left            =   7440
      TabIndex        =   470
      Top             =   2520
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   463
      Left            =   7440
      TabIndex        =   469
      Top             =   2730
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   462
      Left            =   7440
      TabIndex        =   468
      Top             =   2940
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   461
      Left            =   7440
      TabIndex        =   467
      Top             =   3150
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   460
      Left            =   7440
      TabIndex        =   466
      Top             =   3360
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   459
      Left            =   7440
      TabIndex        =   465
      Top             =   3570
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   458
      Left            =   7440
      TabIndex        =   464
      Top             =   3780
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   457
      Left            =   7440
      TabIndex        =   463
      Top             =   3990
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   456
      Left            =   7455
      TabIndex        =   462
      Top             =   2310
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   455
      Left            =   7455
      TabIndex        =   461
      Top             =   2100
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   454
      Left            =   7455
      TabIndex        =   460
      Top             =   1890
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   453
      Left            =   7455
      TabIndex        =   459
      Top             =   1680
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   452
      Left            =   7455
      TabIndex        =   458
      Top             =   1470
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   451
      Left            =   7455
      TabIndex        =   457
      Top             =   1260
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   450
      Left            =   7455
      TabIndex        =   456
      Top             =   1050
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   449
      Left            =   7455
      TabIndex        =   455
      Top             =   825
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   448
      Left            =   7650
      TabIndex        =   454
      Top             =   2925
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   447
      Left            =   7650
      TabIndex        =   453
      Top             =   3135
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   446
      Left            =   7650
      TabIndex        =   452
      Top             =   3345
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   445
      Left            =   7650
      TabIndex        =   451
      Top             =   3555
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   444
      Left            =   7650
      TabIndex        =   450
      Top             =   3765
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   443
      Left            =   7650
      TabIndex        =   449
      Top             =   3975
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   442
      Left            =   7650
      TabIndex        =   448
      Top             =   4185
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   441
      Left            =   7650
      TabIndex        =   447
      Top             =   4395
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   440
      Left            =   7665
      TabIndex        =   446
      Top             =   2715
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   439
      Left            =   7665
      TabIndex        =   445
      Top             =   2505
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   438
      Left            =   7665
      TabIndex        =   444
      Top             =   2295
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   437
      Left            =   7665
      TabIndex        =   443
      Top             =   2085
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   436
      Left            =   7665
      TabIndex        =   442
      Top             =   1875
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   435
      Left            =   7665
      TabIndex        =   441
      Top             =   1665
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   434
      Left            =   7665
      TabIndex        =   440
      Top             =   1455
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   433
      Left            =   7665
      TabIndex        =   439
      Top             =   1230
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   432
      Left            =   9375
      TabIndex        =   438
      Top             =   1230
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   431
      Left            =   9375
      TabIndex        =   437
      Top             =   1455
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   430
      Left            =   9375
      TabIndex        =   436
      Top             =   1665
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   429
      Left            =   9375
      TabIndex        =   435
      Top             =   1875
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   428
      Left            =   9375
      TabIndex        =   434
      Top             =   2085
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   427
      Left            =   9375
      TabIndex        =   433
      Top             =   2295
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   426
      Left            =   9375
      TabIndex        =   432
      Top             =   2505
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   425
      Left            =   9375
      TabIndex        =   431
      Top             =   2715
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   424
      Left            =   9360
      TabIndex        =   430
      Top             =   4395
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   423
      Left            =   9360
      TabIndex        =   429
      Top             =   4185
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   422
      Left            =   9360
      TabIndex        =   428
      Top             =   3975
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   421
      Left            =   9360
      TabIndex        =   427
      Top             =   3765
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   420
      Left            =   9360
      TabIndex        =   426
      Top             =   3555
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   419
      Left            =   9360
      TabIndex        =   425
      Top             =   3345
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   418
      Left            =   9360
      TabIndex        =   424
      Top             =   3135
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   417
      Left            =   9360
      TabIndex        =   423
      Top             =   2925
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   416
      Left            =   8700
      TabIndex        =   422
      Top             =   2805
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   415
      Left            =   8700
      TabIndex        =   421
      Top             =   3015
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   414
      Left            =   8700
      TabIndex        =   420
      Top             =   3225
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   413
      Left            =   8700
      TabIndex        =   419
      Top             =   3435
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   412
      Left            =   8700
      TabIndex        =   418
      Top             =   3645
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   411
      Left            =   8700
      TabIndex        =   417
      Top             =   3855
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   410
      Left            =   8700
      TabIndex        =   416
      Top             =   4065
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   409
      Left            =   8700
      TabIndex        =   415
      Top             =   4275
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   408
      Left            =   8715
      TabIndex        =   414
      Top             =   2595
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   407
      Left            =   8715
      TabIndex        =   413
      Top             =   2385
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   406
      Left            =   8715
      TabIndex        =   412
      Top             =   2175
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   405
      Left            =   8715
      TabIndex        =   411
      Top             =   1965
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   404
      Left            =   8715
      TabIndex        =   410
      Top             =   1755
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   403
      Left            =   8715
      TabIndex        =   409
      Top             =   1545
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   402
      Left            =   8715
      TabIndex        =   408
      Top             =   1335
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   401
      Left            =   8715
      TabIndex        =   407
      Top             =   1110
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   400
      Left            =   6915
      TabIndex        =   406
      Top             =   2835
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   399
      Left            =   6915
      TabIndex        =   405
      Top             =   3045
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   398
      Left            =   6915
      TabIndex        =   404
      Top             =   3255
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   397
      Left            =   6915
      TabIndex        =   403
      Top             =   3465
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   396
      Left            =   6915
      TabIndex        =   402
      Top             =   3675
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   395
      Left            =   6915
      TabIndex        =   401
      Top             =   3885
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   394
      Left            =   6915
      TabIndex        =   400
      Top             =   4095
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   393
      Left            =   6915
      TabIndex        =   399
      Top             =   4305
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   392
      Left            =   6930
      TabIndex        =   398
      Top             =   2625
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   391
      Left            =   6930
      TabIndex        =   397
      Top             =   2415
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   390
      Left            =   6930
      TabIndex        =   396
      Top             =   2205
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   389
      Left            =   6930
      TabIndex        =   395
      Top             =   1995
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   388
      Left            =   6930
      TabIndex        =   394
      Top             =   1785
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   387
      Left            =   6930
      TabIndex        =   393
      Top             =   1575
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   386
      Left            =   6930
      TabIndex        =   392
      Top             =   1365
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   385
      Left            =   6930
      TabIndex        =   391
      Top             =   1140
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   384
      Left            =   8910
      TabIndex        =   390
      Top             =   885
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   383
      Left            =   8910
      TabIndex        =   389
      Top             =   1110
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   382
      Left            =   8910
      TabIndex        =   388
      Top             =   1320
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   381
      Left            =   8910
      TabIndex        =   387
      Top             =   1530
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   380
      Left            =   8910
      TabIndex        =   386
      Top             =   1740
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   379
      Left            =   8910
      TabIndex        =   385
      Top             =   1950
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   378
      Left            =   8910
      TabIndex        =   384
      Top             =   2160
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   377
      Left            =   8910
      TabIndex        =   383
      Top             =   2370
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   376
      Left            =   8895
      TabIndex        =   382
      Top             =   4050
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   375
      Left            =   8895
      TabIndex        =   381
      Top             =   3840
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   374
      Left            =   8895
      TabIndex        =   380
      Top             =   3630
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   373
      Left            =   8895
      TabIndex        =   379
      Top             =   3420
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   372
      Left            =   8895
      TabIndex        =   378
      Top             =   3210
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   371
      Left            =   8895
      TabIndex        =   377
      Top             =   3000
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   370
      Left            =   8895
      TabIndex        =   376
      Top             =   2790
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   369
      Left            =   8895
      TabIndex        =   375
      Top             =   2580
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   368
      Left            =   9555
      TabIndex        =   374
      Top             =   2700
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   367
      Left            =   9555
      TabIndex        =   373
      Top             =   2910
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   366
      Left            =   9555
      TabIndex        =   372
      Top             =   3120
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   365
      Left            =   9555
      TabIndex        =   371
      Top             =   3330
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   364
      Left            =   9555
      TabIndex        =   370
      Top             =   3540
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   363
      Left            =   9555
      TabIndex        =   369
      Top             =   3750
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   362
      Left            =   9555
      TabIndex        =   368
      Top             =   3960
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   361
      Left            =   9555
      TabIndex        =   367
      Top             =   4170
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   360
      Left            =   9570
      TabIndex        =   366
      Top             =   2490
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   359
      Left            =   9570
      TabIndex        =   365
      Top             =   2280
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   358
      Left            =   9570
      TabIndex        =   364
      Top             =   2070
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   357
      Left            =   9570
      TabIndex        =   363
      Top             =   1860
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   356
      Left            =   9570
      TabIndex        =   362
      Top             =   1650
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   355
      Left            =   9570
      TabIndex        =   361
      Top             =   1440
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   354
      Left            =   9570
      TabIndex        =   360
      Top             =   1230
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   353
      Left            =   9570
      TabIndex        =   359
      Top             =   1005
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   352
      Left            =   7860
      TabIndex        =   358
      Top             =   1005
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   351
      Left            =   7860
      TabIndex        =   357
      Top             =   1230
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   350
      Left            =   7860
      TabIndex        =   356
      Top             =   1440
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   349
      Left            =   7860
      TabIndex        =   355
      Top             =   1650
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   348
      Left            =   7860
      TabIndex        =   354
      Top             =   1860
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   347
      Left            =   7860
      TabIndex        =   353
      Top             =   2070
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   346
      Left            =   7860
      TabIndex        =   352
      Top             =   2280
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   345
      Left            =   7860
      TabIndex        =   351
      Top             =   2490
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   344
      Left            =   7845
      TabIndex        =   350
      Top             =   4170
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   343
      Left            =   7845
      TabIndex        =   349
      Top             =   3960
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   342
      Left            =   7845
      TabIndex        =   348
      Top             =   3750
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   341
      Left            =   7845
      TabIndex        =   347
      Top             =   3540
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   340
      Left            =   7845
      TabIndex        =   346
      Top             =   3330
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   339
      Left            =   7845
      TabIndex        =   345
      Top             =   3120
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   338
      Left            =   7845
      TabIndex        =   344
      Top             =   2910
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   337
      Left            =   7845
      TabIndex        =   343
      Top             =   2700
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   336
      Left            =   7650
      TabIndex        =   342
      Top             =   600
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   335
      Left            =   7650
      TabIndex        =   341
      Top             =   825
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   334
      Left            =   7650
      TabIndex        =   340
      Top             =   1035
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   333
      Left            =   7650
      TabIndex        =   339
      Top             =   1245
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   332
      Left            =   7650
      TabIndex        =   338
      Top             =   1455
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   331
      Left            =   7650
      TabIndex        =   337
      Top             =   1665
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   330
      Left            =   7650
      TabIndex        =   336
      Top             =   1875
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   329
      Left            =   7650
      TabIndex        =   335
      Top             =   2085
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   328
      Left            =   7635
      TabIndex        =   334
      Top             =   3765
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   327
      Left            =   7635
      TabIndex        =   333
      Top             =   3555
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   326
      Left            =   7635
      TabIndex        =   332
      Top             =   3345
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   325
      Left            =   7635
      TabIndex        =   331
      Top             =   3135
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   324
      Left            =   7635
      TabIndex        =   330
      Top             =   2925
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   323
      Left            =   7635
      TabIndex        =   329
      Top             =   2715
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   322
      Left            =   7635
      TabIndex        =   328
      Top             =   2505
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   321
      Left            =   7635
      TabIndex        =   327
      Top             =   2295
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   320
      Left            =   8760
      TabIndex        =   326
      Top             =   2310
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   319
      Left            =   8760
      TabIndex        =   325
      Top             =   2520
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   318
      Left            =   8760
      TabIndex        =   324
      Top             =   2730
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   317
      Left            =   8760
      TabIndex        =   323
      Top             =   2940
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   316
      Left            =   8760
      TabIndex        =   322
      Top             =   3150
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   315
      Left            =   8760
      TabIndex        =   321
      Top             =   3360
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   314
      Left            =   8760
      TabIndex        =   320
      Top             =   3570
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   313
      Left            =   8760
      TabIndex        =   319
      Top             =   3780
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   312
      Left            =   8775
      TabIndex        =   318
      Top             =   2100
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   311
      Left            =   8775
      TabIndex        =   317
      Top             =   1890
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   310
      Left            =   8775
      TabIndex        =   316
      Top             =   1680
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   309
      Left            =   8775
      TabIndex        =   315
      Top             =   1470
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   308
      Left            =   8775
      TabIndex        =   314
      Top             =   1260
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   307
      Left            =   8775
      TabIndex        =   313
      Top             =   1050
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   306
      Left            =   8775
      TabIndex        =   312
      Top             =   840
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   305
      Left            =   8775
      TabIndex        =   311
      Top             =   615
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   304
      Left            =   8970
      TabIndex        =   310
      Top             =   2715
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   303
      Left            =   8970
      TabIndex        =   309
      Top             =   2925
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   302
      Left            =   8970
      TabIndex        =   308
      Top             =   3135
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   301
      Left            =   8970
      TabIndex        =   307
      Top             =   3345
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   300
      Left            =   8970
      TabIndex        =   306
      Top             =   3555
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   299
      Left            =   8970
      TabIndex        =   305
      Top             =   3765
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   298
      Left            =   8970
      TabIndex        =   304
      Top             =   3975
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   297
      Left            =   8970
      TabIndex        =   303
      Top             =   4185
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   296
      Left            =   8985
      TabIndex        =   302
      Top             =   2505
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   295
      Left            =   8985
      TabIndex        =   301
      Top             =   2295
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   294
      Left            =   8985
      TabIndex        =   300
      Top             =   2085
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   293
      Left            =   8985
      TabIndex        =   299
      Top             =   1875
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   292
      Left            =   8985
      TabIndex        =   298
      Top             =   1665
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   291
      Left            =   8985
      TabIndex        =   297
      Top             =   1455
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   290
      Left            =   8985
      TabIndex        =   296
      Top             =   1245
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   289
      Left            =   8985
      TabIndex        =   295
      Top             =   1020
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   288
      Left            =   10695
      TabIndex        =   294
      Top             =   1020
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   287
      Left            =   10695
      TabIndex        =   293
      Top             =   1245
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   286
      Left            =   10695
      TabIndex        =   292
      Top             =   1455
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   285
      Left            =   10695
      TabIndex        =   291
      Top             =   1665
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   284
      Left            =   10695
      TabIndex        =   290
      Top             =   1875
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   283
      Left            =   10695
      TabIndex        =   289
      Top             =   2085
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   282
      Left            =   10695
      TabIndex        =   288
      Top             =   2295
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   281
      Left            =   10695
      TabIndex        =   287
      Top             =   2505
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   280
      Left            =   10680
      TabIndex        =   286
      Top             =   4185
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   279
      Left            =   10680
      TabIndex        =   285
      Top             =   3975
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   278
      Left            =   10680
      TabIndex        =   284
      Top             =   3765
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   277
      Left            =   10680
      TabIndex        =   283
      Top             =   3555
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   276
      Left            =   10680
      TabIndex        =   282
      Top             =   3345
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   275
      Left            =   10680
      TabIndex        =   281
      Top             =   3135
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   274
      Left            =   10680
      TabIndex        =   280
      Top             =   2925
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   273
      Left            =   10680
      TabIndex        =   279
      Top             =   2715
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   272
      Left            =   10020
      TabIndex        =   278
      Top             =   2595
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   271
      Left            =   10020
      TabIndex        =   277
      Top             =   2805
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   270
      Left            =   10020
      TabIndex        =   276
      Top             =   3015
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   269
      Left            =   10020
      TabIndex        =   275
      Top             =   3225
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   268
      Left            =   10020
      TabIndex        =   274
      Top             =   3435
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   267
      Left            =   10020
      TabIndex        =   273
      Top             =   3645
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   266
      Left            =   10020
      TabIndex        =   272
      Top             =   3855
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   265
      Left            =   10020
      TabIndex        =   271
      Top             =   4065
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   264
      Left            =   10035
      TabIndex        =   270
      Top             =   2385
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   263
      Left            =   10035
      TabIndex        =   269
      Top             =   2175
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   262
      Left            =   10035
      TabIndex        =   268
      Top             =   1965
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   261
      Left            =   10035
      TabIndex        =   267
      Top             =   1755
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   260
      Left            =   10035
      TabIndex        =   266
      Top             =   1545
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   259
      Left            =   10035
      TabIndex        =   265
      Top             =   1335
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   258
      Left            =   10035
      TabIndex        =   264
      Top             =   1125
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
      Index           =   257
      Left            =   10035
      TabIndex        =   263
      Top             =   900
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "WinAmp Caller"
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
      Caption         =   "WinAmp Caller"
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
      Begin VB.Menu Mnu_Sync 
         Caption         =   "Sync"
      End
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB ME"
   End
   Begin VB.Menu Mnu_LoadFolder 
      Caption         =   "Load Folder"
   End
   Begin VB.Menu Mnu_LoadFolder2 
      Caption         =   "Load Folder Target"
   End
   Begin VB.Menu Folder 
      Caption         =   "Load NotePad Folder"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XXT2

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long



Private Sub Combo1_Click()
Counter = Combo1.ListIndex + 1
A1$ = A4$(Counter)
B1$ = B4$(Counter)
C1$ = C4$(Counter)

SetTrueToLoadLast = True
Call Label1_Click(0)

End Sub

Private Sub Folder_Click()

Shell "explorer /e, D:\# MY DOCS\# 01 My Documents\03 NotePad Files", vbNormalFocus

End Sub

Private Sub Form_Load()

Set FS = CreateObject("Scripting.FileSystemObject")

SetTrueToLoadLast = False

'If FindWindow(vbNullString, "Tidal Information...") = 0 And IsIDE = fasle Then
'    Shell "D:\VB6\VB-NT\00_Best_VB_01\Shell KeyBoard Loader\Shell KeyBoard Loader.exe", vbNormalFocus
'    End
'End If

ReDim A4$(500), B4$(500), C4$(500)

OIP$ = App.Path
tt1 = InStr(OIP$, "Shell")
tt2 = InStr(tt1, OIP$, "Loader")
OIP2$ = Mid$(OIP$, tt1 + 6, tt2 - tt1 - 7)
OIP$ = OIP2$
Open App.Path + "\Text Loggs\" + OIP$ + " Loads Logg.txt" For Binary As #1

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

If Dir$(App.Path + "\Text Loggs\" + OIP$ + " Loads.txt") <> "" Then
    Open App.Path + "\Text Loggs\" + OIP$ + " Loads.txt" For Input As #1
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

Form1.Top = 480
Form1.Left = 0

Call FormStart.FormStartLoader

Call SubCode

If Command$ <> "" Then
SetTrueToLoadLast = True
A1$ = "D:\# MY DOCS\# 01 My Documents\03 NotePad Files\Hosso\Bigg Blogs\"
B1$ = "RutLand Hosso.txt"
Call Label1_Click(0)
End If

Timer2.Enabled = True

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

SH = Screen.Height - 2200 'higer = smaller


x = tx
For r = 1 To ScanPath.ListView1.ListItems.Count
    If Label1(r - 1).Top > SH Then
        x = tx
    End If
    
    On Error Resume Next
        Err.Clear
        Label1(r).Top = x
        If Err.Number > 0 Then Exit For
    On Error GoTo 0
    
    Label1(r).Visible = True
    'For r = 1 To ScanPath.ListView1.ListItems.Count
    A1$ = ScanPath.ListView1.ListItems.Item(r).SubItems(1)
    B1$ = ScanPath.ListView1.ListItems.Item(r)
    If InStrRev(A1$, "\") > 0 Then
        seedy$ = Mid$(A1$, InStrRev(A1$, "\", Len(A1$) - 1))
    End If
    If oseedy$ <> seedy$ Then txr = Not txr
    oseedy$ = seedy$
    If txr = -1 Then Label1(r).BackColor = &HFF00&
    If txr = 0 Then Label1(r).BackColor = &H56E25E
    
    ttg$ = ScanPath.ListView1.ListItems.Item(r)
    Mid$(ttg$, 1, 1) = UCase$(Mid$(ttg$, 1, 1))
    If InStr(ttg$, ".lnk") > 0 Then ttg$ = Mid$(ttg$, 1, InStrRev(ttg$, ".") - 1)
    
    Label1(r).Caption = Format$(r, "00") + ". " + ttg$
    x = x + Label1(r).Height + 15
    fheight = Label1(r).Top + Label1(r).Height + 420
    If fheight > fheightx Then fheightx = fheight
    RD = RD + 1
Next
DoEvents

x = tx
td = 29
xgag = 0
Dim xy(20)

For r = 1 To ScanPath.ListView1.ListItems.Count
    If Label1(r - 1).Top > SH Then
        If tig = 0 Then tig = r
        x = tx
        xgag = xgag + 1
        wdt = 0
        For rt = r - 1 To r - tig Step -1
        If Label1(rt).Width > wdt Then wdt = Label1(rt).Width
        Next
        xy(xgag) = wdt + 150
        tig2 = r
    End If
    Label1(r).Top = x
    x = x + Label1(r).Height + 15
Next

xgag = xgag + 1
wdt = 0
For rt = tig2 To r - 1
    If Label1(rt).Width > wdt Then wdt = Label1(rt).Width
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
fheightx = 0
For rt = 1 To ScanPath.ListView1.ListItems.Count
    fheight = Label1(rt).Top + Label1(rt).Height
    If fheight > fheightx Then
        fheightx = fheight
    Else
        Exit For
    End If
Next

Form1.Width = Label1(r - 1).Left + Label1(r - 1).Width + 120
If Form1.Width > Screen.Width Then Form1.Width = Screen.Width - 120
Form1.Height = fheightx + 700
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

'Center screen
Form1.Top = Screen.Height / 2 - Form1.Height / 2
Form1.Left = Screen.Width / 2 - Form1.Width / 2

Form1.Refresh


End Sub


Private Sub Form_Unload(Cancel As Integer)
'End


'SET ALL TIMERS IN ALL FORMS TO ENABLED = FALSE
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

End Sub

Private Sub Label1_Click(index As Integer)

If SetTrueToLoadLast = False Then
A1$ = ScanPath.ListView1.ListItems.Item(index).SubItems(1)
B1$ = ScanPath.ListView1.ListItems.Item(index)
C1$ = Form1.Label1(index).Caption
Else
'vLaunch A1$ + B1$, vbNullString
'End
End If

If LoadFolder = True And LoadFolder2 = False Then
    'Shell "explorer /e, " + A1$, vbNormalFocus
    Shell "Explorer.exe /select, " + A1$ + B1$, vbNormalFocus
    End
End If

'MsgBox "hello"

Call SaveLoggs

'MsgBox "hello"

Call FormStart.LabelClick(index)

End Sub

Private Sub Lbl2_Click()

Shell "explorer /e, E:\01 VB Shell Folders\00 " + App.EXEName, vbNormalFocus

End Sub

Private Sub Mnu_LoadFolder_Click()
LoadFolder = True
End Sub

Private Sub Mnu_LoadFolder2_Click()
LoadFolder2 = True
End Sub

Private Sub Mnu_Sync_Click()

ScanPath.ListView1.ListItems.Clear

Call FormStart.FormStartLoader

Call SubCode


End Sub

Private Sub MNU_VB_Click()

Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
End

End Sub

Private Sub Timer2_Timer()

XXT2 = XXT2 + 1
If XXT2 > ScanPath.ListView1.ListItems.Count Then
    Timer2.Enabled = False
    Exit Sub
End If

A1$ = ScanPath.ListView1.ListItems.Item(XXT2).SubItems(1)
B1$ = ScanPath.ListView1.ListItems.Item(XXT2)
On Error Resume Next
    Err.Clear
    C1$ = Form1.Label1(XXT2).Caption
    If Err.Number > 0 Then Exit Sub
On Error GoTo 0

On Local Error GoTo 0
On Error GoTo 0

'GoTo skip:
D1$ = A1$ + B1$
If InStr(LCase(B1$), ".lnk") > 0 Then
    Open A1$ + B1$ For Binary As #1
    rr$ = Space$(LOF(1))
    Get #1, 1, rr$
    Close #1
    If InStr(rr$, ":\") = 0 Then
        
        i = MsgBox("INVAILD LINK NAUGHT SIZE" + vbCrLf + D1$, vbMsgBoxSetForeground)
        
        i = MsgBox("INVAILD LINK" + vbCrLf + D1$ + vbCrLf + "DO YOU WANT ME TO TAKE YOU THERE", vbYesNo + vbMsgBoxSetForeground)
        If i = vbYes Then
            Shell "Explorer.exe /select, " + D1$, vbNormalFocus
    
    
        End If
        Exit Sub
    End If
    tt$ = Mid$(rr$, InStr(105, rr$, ":\") - 1)
    D1$ = Mid$(tt$, 1, InStr(tt$, Chr$(0)) - 1)
End If

If FS.FileExists(D1$) = False And FS.FolderExists(D1$) = False Then Label1(XXT2).BackColor = QBColor(12)



End Sub


Private Sub TitleLbl_Click()
On Error Resume Next
Open App.Path + "\Text Loggs\" + OIP$ + " Loads.txt" For Input As #1
Line Input #1, A1$
Line Input #1, B1$
Line Input #1, C1$
Close #1

SetTrueToLoadLast = True
Call Label1_Click(0)

End Sub

Sub SaveLoggs()
Open App.Path + "\Text Loggs\" + OIP$ + " Loads.txt" For Output As #1
Print #1, A1$
Print #1, B1$
Print #1, C1$
Close #1

Open App.Path + "\Text Loggs\" + OIP$ + " Loads Logg.txt" For Append As #1
Print #1, A1$
Print #1, B1$
Print #1, C1$
Close #1

Open App.Path + "\Text Loggs\" + OIP$ + " Loads Logg 02.txt" For Append As #1
Print #1, Format$(Now, "dd-mm-yyyy hh:mm:ss ") + A1$ + B1$
Close #1


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
'***********************************************

