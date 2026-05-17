VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSetting 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10800
   ScaleWidth      =   14070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrmSubSetting 
      Caption         =   "Format"
      Height          =   1335
      Index           =   7
      Left            =   0
      TabIndex        =   94
      Top             =   0
      Width           =   6255
      Begin VB.PictureBox PicFramSetting 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   7
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   6015
         TabIndex        =   95
         Top             =   240
         Width           =   6015
         Begin VB.OptionButton OptFormat 
            Caption         =   "Fix"
            Height          =   495
            Index           =   3
            Left            =   5460
            Style           =   1  'Graphical
            TabIndex        =   104
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptFormat 
            Caption         =   "Mark && Fix"
            Height          =   495
            Index           =   2
            Left            =   4965
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptFormat 
            Caption         =   "Mark"
            Height          =   495
            Index           =   1
            Left            =   4470
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptFormat 
            Caption         =   "Off"
            Height          =   495
            Index           =   0
            Left            =   3975
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   0
            Width           =   495
         End
         Begin VB.ListBox LstFormat 
            Height          =   510
            Index           =   0
            Left            =   -240
            Style           =   1  'Checkbox
            TabIndex        =   100
            Top             =   480
            Width           =   4215
         End
         Begin VB.ListBox LstFormat 
            Height          =   510
            Index           =   1
            Left            =   3975
            Style           =   1  'Checkbox
            TabIndex        =   99
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstFormat 
            Height          =   510
            Index           =   2
            Left            =   4470
            Style           =   1  'Checkbox
            TabIndex        =   98
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstFormat 
            Height          =   510
            Index           =   3
            Left            =   4965
            Style           =   1  'Checkbox
            TabIndex        =   97
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstFormat 
            Height          =   510
            Index           =   4
            Left            =   5460
            Style           =   1  'Checkbox
            TabIndex        =   96
            Top             =   480
            Width           =   495
         End
         Begin VB.Label LbLFramSetting 
            BackColor       =   &H8000000D&
            Caption         =   $"FrmSetting.frx":0000
            Height          =   495
            Index           =   7
            Left            =   0
            TabIndex        =   105
            Top             =   0
            Width           =   5895
         End
      End
   End
   Begin VB.Frame FrmSubSetting 
      Caption         =   "Unused Detection"
      Height          =   1350
      Index           =   6
      Left            =   0
      TabIndex        =   82
      Top             =   0
      Width           =   6255
      Begin VB.PictureBox PicFramSetting 
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         Height          =   1050
         Index           =   6
         Left            =   100
         ScaleHeight     =   1050
         ScaleWidth      =   6015
         TabIndex        =   83
         Top             =   240
         Width           =   6015
         Begin VB.OptionButton OptUnused 
            Caption         =   "Fix"
            Height          =   495
            Index           =   3
            Left            =   5460
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptUnused 
            Caption         =   "Mark && Fix"
            Height          =   495
            Index           =   2
            Left            =   4965
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptUnused 
            Caption         =   "Mark"
            Height          =   495
            Index           =   1
            Left            =   4470
            Style           =   1  'Graphical
            TabIndex        =   90
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptUnused 
            Caption         =   "Off"
            Height          =   495
            Index           =   0
            Left            =   3975
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   0
            Width           =   495
         End
         Begin VB.ListBox LstUnused 
            Height          =   510
            Index           =   0
            Left            =   -240
            Style           =   1  'Checkbox
            TabIndex        =   88
            Top             =   480
            Width           =   4215
         End
         Begin VB.ListBox LstUnused 
            Height          =   510
            Index           =   1
            Left            =   3975
            Style           =   1  'Checkbox
            TabIndex        =   87
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstUnused 
            Height          =   510
            Index           =   2
            Left            =   4470
            Style           =   1  'Checkbox
            TabIndex        =   86
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstUnused 
            Height          =   510
            Index           =   3
            Left            =   4965
            Style           =   1  'Checkbox
            TabIndex        =   85
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstUnused 
            Height          =   510
            Index           =   4
            Left            =   5460
            Style           =   1  'Checkbox
            TabIndex        =   84
            Top             =   480
            Width           =   480
         End
         Begin VB.Label LbLFramSetting 
            BackColor       =   &H8000000D&
            Caption         =   $"FrmSetting.frx":00EA
            Height          =   495
            Index           =   6
            Left            =   0
            TabIndex        =   93
            Top             =   0
            Width           =   5895
         End
      End
   End
   Begin VB.Frame FrmSubSetting 
      Caption         =   "Dim "
      Height          =   1335
      Index           =   4
      Left            =   0
      TabIndex        =   70
      Top             =   0
      Width           =   6255
      Begin VB.PictureBox PicFramSetting 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   1155
         Index           =   4
         Left            =   120
         ScaleHeight     =   1155
         ScaleWidth      =   6015
         TabIndex        =   71
         Top             =   240
         Width           =   6015
         Begin VB.OptionButton OptDim 
            Caption         =   "Fix"
            Height          =   495
            Index           =   3
            Left            =   5460
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptDim 
            Caption         =   "Mark && Fix"
            Height          =   495
            Index           =   2
            Left            =   4965
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptDim 
            Caption         =   "Mark"
            Height          =   495
            Index           =   1
            Left            =   4470
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptDim 
            Caption         =   "Off"
            Height          =   495
            Index           =   0
            Left            =   3975
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   0
            Width           =   495
         End
         Begin VB.ListBox LstDim 
            Height          =   510
            Index           =   0
            Left            =   -240
            Style           =   1  'Checkbox
            TabIndex        =   76
            Top             =   480
            Width           =   4215
         End
         Begin VB.ListBox LstDim 
            Height          =   510
            Index           =   1
            Left            =   3975
            Style           =   1  'Checkbox
            TabIndex        =   75
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstDim 
            Height          =   510
            Index           =   2
            Left            =   4470
            Style           =   1  'Checkbox
            TabIndex        =   74
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstDim 
            Height          =   510
            Index           =   3
            Left            =   4965
            Style           =   1  'Checkbox
            TabIndex        =   73
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstDim 
            Height          =   510
            Index           =   4
            Left            =   5460
            Style           =   1  'Checkbox
            TabIndex        =   72
            Top             =   480
            Width           =   495
         End
         Begin VB.Label LbLFramSetting 
            BackColor       =   &H8000000D&
            Caption         =   $"FrmSetting.frx":01D4
            Height          =   495
            Index           =   4
            Left            =   0
            TabIndex        =   81
            Top             =   0
            Width           =   5895
         End
      End
   End
   Begin VB.Frame FrmSubSetting 
      Caption         =   "ReSutructure"
      Height          =   1335
      Index           =   2
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Width           =   6255
      Begin VB.PictureBox PicFramSetting 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   1080
         Index           =   2
         Left            =   120
         ScaleHeight     =   1080
         ScaleWidth      =   6075
         TabIndex        =   59
         Top             =   240
         Width           =   6075
         Begin VB.OptionButton OptRestruct 
            Caption         =   "Fix"
            Height          =   495
            Index           =   3
            Left            =   5460
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   15
            Width           =   495
         End
         Begin VB.OptionButton OptRestruct 
            Caption         =   "Mark && Fix"
            Height          =   495
            Index           =   2
            Left            =   4965
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   15
            Width           =   495
         End
         Begin VB.OptionButton OptRestruct 
            Caption         =   "Mark"
            Height          =   495
            Index           =   1
            Left            =   4470
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   15
            Width           =   495
         End
         Begin VB.OptionButton OptRestruct 
            Caption         =   "Off"
            Height          =   495
            Index           =   0
            Left            =   3975
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   15
            Width           =   495
         End
         Begin VB.ListBox LstRestructure 
            Height          =   510
            Index           =   0
            Left            =   -240
            Style           =   1  'Checkbox
            TabIndex        =   64
            Top             =   480
            Width           =   4215
         End
         Begin VB.ListBox LstRestructure 
            Height          =   510
            Index           =   1
            Left            =   3975
            Style           =   1  'Checkbox
            TabIndex        =   63
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstRestructure 
            Height          =   510
            Index           =   2
            Left            =   4470
            Style           =   1  'Checkbox
            TabIndex        =   62
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstRestructure 
            Height          =   510
            Index           =   3
            Left            =   4965
            Style           =   1  'Checkbox
            TabIndex        =   61
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstRestructure 
            Height          =   510
            Index           =   4
            Left            =   5460
            Style           =   1  'Checkbox
            TabIndex        =   60
            Top             =   480
            Width           =   495
         End
         Begin VB.Label LbLFramSetting 
            BackColor       =   &H8000000D&
            Caption         =   $"FrmSetting.frx":02BE
            Height          =   495
            Index           =   2
            Left            =   0
            TabIndex        =   69
            Top             =   0
            Width           =   5895
         End
      End
   End
   Begin VB.Frame FrmSubSetting 
      Caption         =   "General"
      Height          =   2535
      Index           =   0
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   6255
      Begin VB.PictureBox Picture9 
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2175
         ScaleWidth      =   5895
         TabIndex        =   39
         Top             =   240
         Width           =   5895
         Begin VB.CheckBox chkXPStyle 
            Caption         =   "XP Style Support"
            Height          =   255
            Left            =   0
            TabIndex        =   57
            Top             =   1920
            Width           =   1935
         End
         Begin VB.CommandButton CmdStartFromSettings 
            Caption         =   "Start"
            Height          =   315
            Left            =   4560
            TabIndex        =   56
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Frame framButtons 
            Caption         =   "Quick Sets"
            Height          =   1695
            Left            =   4320
            TabIndex        =   49
            Top             =   0
            Width           =   1575
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   1335
               Left            =   120
               ScaleHeight     =   1335
               ScaleWidth      =   1335
               TabIndex        =   50
               Top             =   240
               Width           =   1335
               Begin VB.OptionButton OptLevelSetting 
                  Caption         =   "All Off"
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   55
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton OptLevelSetting 
                  Caption         =   "Minimum"
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   54
                  Top             =   255
                  Width           =   1095
               End
               Begin VB.OptionButton OptLevelSetting 
                  Caption         =   "Average"
                  Height          =   255
                  Index           =   2
                  Left            =   0
                  TabIndex        =   53
                  Top             =   510
                  Value           =   -1  'True
                  Width           =   1095
               End
               Begin VB.OptionButton OptLevelSetting 
                  Caption         =   "Maximium"
                  Height          =   255
                  Index           =   3
                  Left            =   0
                  TabIndex        =   52
                  Top             =   765
                  Width           =   1095
               End
               Begin VB.OptionButton OptLevelSetting 
                  Caption         =   "User"
                  Height          =   255
                  Index           =   4
                  Left            =   0
                  TabIndex        =   51
                  Top             =   1020
                  Width           =   1095
               End
            End
         End
         Begin VB.CheckBox chkAllowSpaces 
            Caption         =   "Allow Space Separators"
            Height          =   375
            Left            =   2040
            TabIndex        =   48
            ToolTipText     =   "Check to allow more double blank lines to be preserved"
            Top             =   1560
            Width           =   2295
         End
         Begin VB.Frame framSetting 
            Caption         =   "ToolPosition"
            Height          =   1575
            Index           =   6
            Left            =   0
            TabIndex        =   44
            Top             =   0
            Width           =   1935
            Begin VB.PictureBox pictoolpos 
               BorderStyle     =   0  'None
               Height          =   1095
               Left            =   120
               ScaleHeight     =   1095
               ScaleWidth      =   1695
               TabIndex        =   45
               Top             =   240
               Width           =   1695
               Begin VB.ComboBox cmbToolPos 
                  Height          =   315
                  Left            =   0
                  Style           =   2  'Dropdown List
                  TabIndex        =   46
                  Top             =   0
                  Width           =   1695
               End
               Begin VB.Label Label2 
                  Caption         =   "NOTE Tool will be set to nearest fully on screen position"
                  Height          =   735
                  Left            =   0
                  TabIndex        =   47
                  Top             =   480
                  Width           =   1695
               End
            End
         End
         Begin VB.Frame framUserSetting 
            Caption         =   "User Settings"
            Height          =   855
            Left            =   2040
            TabIndex        =   41
            Top             =   0
            Width           =   2055
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   495
               Left            =   120
               ScaleHeight     =   495
               ScaleWidth      =   1815
               TabIndex        =   42
               Top             =   240
               Width           =   1815
               Begin VB.CommandButton cmdUsereSettings 
                  Caption         =   "Save User settings"
                  Height          =   375
                  Left            =   0
                  TabIndex        =   43
                  Top             =   0
                  Width           =   1695
               End
            End
         End
         Begin VB.CheckBox ChkStructComments 
            Caption         =   "Insert Structural Comments"
            Height          =   735
            Left            =   2040
            TabIndex        =   40
            Top             =   1680
            Width           =   2295
         End
      End
   End
   Begin VB.Frame FrmSubSetting 
      Caption         =   "Suggestions"
      Height          =   1350
      Index           =   5
      Left            =   7440
      TabIndex        =   26
      Top             =   3240
      Width           =   6255
      Begin VB.PictureBox PicFramSetting 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   1080
         Index           =   5
         Left            =   120
         ScaleHeight     =   1080
         ScaleWidth      =   6015
         TabIndex        =   27
         Top             =   240
         Width           =   6015
         Begin VB.OptionButton OptSugest 
            Caption         =   "Fix"
            Height          =   495
            Index           =   3
            Left            =   5460
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptSugest 
            Caption         =   "Mark && Fix"
            Height          =   495
            Index           =   2
            Left            =   4965
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptSugest 
            Caption         =   "Mark"
            Height          =   495
            Index           =   1
            Left            =   4470
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptSugest 
            Caption         =   "Off"
            Height          =   495
            Index           =   0
            Left            =   3975
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   0
            Width           =   495
         End
         Begin VB.ListBox LstSuggest 
            Height          =   510
            Index           =   0
            Left            =   -240
            Style           =   1  'Checkbox
            TabIndex        =   32
            Top             =   480
            Width           =   4215
         End
         Begin VB.ListBox LstSuggest 
            Height          =   510
            Index           =   1
            Left            =   3975
            Style           =   1  'Checkbox
            TabIndex        =   31
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstSuggest 
            Height          =   510
            Index           =   2
            Left            =   4470
            Style           =   1  'Checkbox
            TabIndex        =   30
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstSuggest 
            Height          =   510
            Index           =   3
            Left            =   4965
            Style           =   1  'Checkbox
            TabIndex        =   29
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstSuggest 
            Height          =   510
            Index           =   4
            Left            =   5460
            Style           =   1  'Checkbox
            TabIndex        =   28
            Top             =   480
            Width           =   495
         End
         Begin VB.Label LbLFramSetting 
            BackColor       =   &H8000000D&
            Caption         =   $"FrmSetting.frx":03A8
            Height          =   495
            Index           =   5
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   5895
         End
      End
   End
   Begin VB.Frame FrmSubSetting 
      Caption         =   "Parameters"
      Height          =   1335
      Index           =   3
      Left            =   480
      TabIndex        =   14
      Top             =   7080
      Width           =   6255
      Begin VB.PictureBox PicFramSetting 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   1155
         Index           =   3
         Left            =   120
         ScaleHeight     =   1155
         ScaleWidth      =   6015
         TabIndex        =   15
         Top             =   240
         Width           =   6015
         Begin VB.ListBox LstParam 
            Height          =   510
            Index           =   4
            Left            =   5460
            Style           =   1  'Checkbox
            TabIndex        =   24
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstParam 
            Height          =   510
            Index           =   3
            Left            =   4965
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstParam 
            Height          =   510
            Index           =   2
            Left            =   4470
            Style           =   1  'Checkbox
            TabIndex        =   22
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstParam 
            Height          =   510
            Index           =   1
            Left            =   3975
            Style           =   1  'Checkbox
            TabIndex        =   21
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LstParam 
            Height          =   510
            Index           =   0
            Left            =   -240
            Style           =   1  'Checkbox
            TabIndex        =   20
            Top             =   480
            Width           =   4215
         End
         Begin VB.OptionButton OptParam 
            Caption         =   "Off"
            Height          =   495
            Index           =   0
            Left            =   3975
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptParam 
            Caption         =   "Mark"
            Height          =   495
            Index           =   1
            Left            =   4470
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptParam 
            Caption         =   "Mark && Fix"
            Height          =   495
            Index           =   2
            Left            =   4965
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptParam 
            Caption         =   "Fix"
            Height          =   495
            Index           =   3
            Left            =   5460
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   0
            Width           =   495
         End
         Begin VB.Label LbLFramSetting 
            BackColor       =   &H8000000D&
            Caption         =   $"FrmSetting.frx":0492
            Height          =   495
            Index           =   3
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   5895
         End
      End
   End
   Begin VB.Frame FrmSubSetting 
      Caption         =   "Declaration Tools"
      Height          =   1455
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   5520
      Width           =   6255
      Begin VB.PictureBox PicFramSetting 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   6015
         TabIndex        =   3
         Top             =   240
         Width           =   6015
         Begin VB.OptionButton OptDec 
            Caption         =   "Fix"
            Height          =   495
            Index           =   3
            Left            =   5460
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptDec 
            Caption         =   "Mark && Fix"
            Height          =   495
            Index           =   2
            Left            =   4965
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptDec 
            Caption         =   "Mark"
            Height          =   495
            Index           =   1
            Left            =   4470
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton OptDec 
            Caption         =   "Off"
            Height          =   495
            Index           =   0
            Left            =   3975
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   0
            Width           =   495
         End
         Begin VB.ListBox LStDeclaration 
            Height          =   510
            Index           =   0
            Left            =   -240
            Style           =   1  'Checkbox
            TabIndex        =   8
            Top             =   480
            Width           =   4215
         End
         Begin VB.ListBox LStDeclaration 
            Height          =   510
            Index           =   1
            Left            =   3975
            Style           =   1  'Checkbox
            TabIndex        =   7
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LStDeclaration 
            Height          =   510
            Index           =   2
            Left            =   4470
            Style           =   1  'Checkbox
            TabIndex        =   6
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LStDeclaration 
            Height          =   510
            Index           =   3
            Left            =   4965
            Style           =   1  'Checkbox
            TabIndex        =   5
            Top             =   480
            Width           =   495
         End
         Begin VB.ListBox LStDeclaration 
            Height          =   510
            Index           =   4
            Left            =   5460
            Style           =   1  'Checkbox
            TabIndex        =   4
            Top             =   480
            Width           =   495
         End
         Begin VB.Label LbLFramSetting 
            BackColor       =   &H8000000D&
            Caption         =   "              Actions"
            Height          =   495
            Index           =   1
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   5895
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   300
      Left            =   7080
      Picture         =   "FrmSetting.frx":057C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSComctlLib.TabStrip TabStrip2 
      Height          =   4935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8705
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "general"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Declaration"
            Key             =   "declaration"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Restructure"
            Key             =   "restructure"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parameter"
            Key             =   "parameter"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dim"
            Key             =   "dim"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Suggest"
            Key             =   "Suggestion"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Unused"
            Key             =   "Unused"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Format"
            Key             =   "format"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Caption = "Settings " & AppDetails
TabStrip2.Tabs(1).Selected = True
End Sub

Private Sub Form_Resize()
Me.Width = TabStrip2.Width
Me.Height = TabStrip2.Height
StayOnTop Me, True
End Sub

Private Sub TabStrip2_Click()

  Dim I As Long

  With TabStrip2
    For I = 0 To 7
      FrmSubSetting(I).Visible = False
    Next I
    FrmSubSetting(.SelectedItem.Index - 1).Move .ClientLeft, .ClientTop
    FrmSubSetting(.SelectedItem.Index - 1).Visible = True
    CurSettingFrame = .SelectedItem.Index
  End With

End Sub
