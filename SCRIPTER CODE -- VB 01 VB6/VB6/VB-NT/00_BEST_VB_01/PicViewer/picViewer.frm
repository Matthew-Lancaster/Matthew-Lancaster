VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPicViewer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Auto image viewer"
   ClientHeight    =   8370
   ClientLeft      =   300
   ClientTop       =   345
   ClientWidth     =   13530
   Icon            =   "picViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   558
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   902
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   6030
      Top             =   435
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   0
      TabIndex        =   70
      Top             =   -30
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox picViewLargeBar 
      Height          =   405
      Left            =   120
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   643
      TabIndex        =   66
      Top             =   960
      Visible         =   0   'False
      Width           =   9705
      Begin VB.CommandButton cmdViewLargeExit 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         Picture         =   "picViewer.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         ToolTipText     =   "Return to main screen"
         Top             =   0
         Width           =   675
      End
      Begin VB.ComboBox cboMagnify 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   68
         ToolTipText     =   "Scale of magnification"
         Top             =   0
         Width           =   690
      End
      Begin VB.CommandButton cmdMagnifyGlass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1710
         Picture         =   "picViewer.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         ToolTipText     =   "Toggle magnifying glass"
         Top             =   0
         Width           =   345
      End
   End
   Begin VB.PictureBox picViewLargeContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      Height          =   6225
      Left            =   2505
      ScaleHeight     =   411
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   643
      TabIndex        =   53
      Top             =   6930
      Width           =   9705
      Begin VB.PictureBox picGlass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   720
         ScaleHeight     =   43
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   65
         Top             =   690
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.VScrollBar VsbViewLarge 
         Height          =   5895
         Left            =   9405
         TabIndex        =   58
         Top             =   0
         Width           =   240
      End
      Begin VB.HScrollBar HsbViewLarge 
         Height          =   255
         Left            =   30
         TabIndex        =   57
         Top             =   5910
         Width           =   9375
      End
      Begin VB.PictureBox picViewLarge 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6285
         Left            =   -15
         ScaleHeight     =   417
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   627
         TabIndex        =   56
         Top             =   -15
         Width           =   9435
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   840
            Top             =   1830
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
   End
   Begin VB.PictureBox piccmdFitInLargeContainer 
      Height          =   315
      Left            =   2670
      ScaleHeight     =   255
      ScaleWidth      =   165
      TabIndex        =   54
      ToolTipText     =   "Fit in"
      Top             =   3510
      Visible         =   0   'False
      Width           =   225
      Begin VB.CommandButton cmdViewLarge 
         BackColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Enlarge viewport"
         Top             =   0
         Width           =   165
      End
   End
   Begin VB.PictureBox piccmdOrigSizeContainer 
      Height          =   315
      Left            =   2370
      ScaleHeight     =   255
      ScaleWidth      =   165
      TabIndex        =   51
      Top             =   3510
      Visible         =   0   'False
      Width           =   225
      Begin VB.CommandButton cmdOrigSize 
         BackColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   -60
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         ToolTipText     =   "Return to original size"
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.ComboBox cboResize 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1650
      Style           =   2  'Dropdown List
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "+/- percentage"
      Top             =   3510
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox piccmdFitInContainer 
      Height          =   315
      Left            =   2970
      ScaleHeight     =   255
      ScaleWidth      =   165
      TabIndex        =   48
      ToolTipText     =   "Fit in"
      Top             =   3510
      Visible         =   0   'False
      Width           =   225
      Begin VB.CommandButton cmdFitIn 
         BackColor       =   &H00C0E0FF&
         Height          =   255
         Left            =   -30
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Image fit in"
         Top             =   0
         Width           =   195
      End
   End
   Begin VB.ComboBox cboPattern 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      ItemData        =   "picViewer.frx":0796
      Left            =   240
      List            =   "picViewer.frx":0798
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   1080
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.PictureBox picViewZ 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      DragMode        =   1  'Automatic
      Height          =   3255
      Left            =   180
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   33
      Top             =   3900
      Width           =   4575
      Begin VB.PictureBox picViewTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   150
         ScaleHeight     =   45
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   35
         Top             =   150
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.PictureBox picCopy 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   1080
         ScaleHeight     =   45
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   49
         TabIndex        =   34
         Top             =   150
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.PictureBox picViewX 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   3195
         Left            =   0
         ScaleHeight     =   213
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   298
         TabIndex        =   36
         Top             =   0
         Width           =   4470
      End
   End
   Begin VB.FileListBox FilList 
      Height          =   2430
      Left            =   2670
      TabIndex        =   32
      Top             =   990
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.VScrollBar VsbView 
      Height          =   3255
      Left            =   4770
      TabIndex        =   31
      Top             =   3900
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HsbView 
      Height          =   285
      Left            =   180
      TabIndex        =   30
      Top             =   7200
      Width           =   4545
   End
   Begin VB.DirListBox DirList 
      Height          =   1890
      Left            =   210
      TabIndex        =   29
      Top             =   1530
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.DriveListBox drvList 
      Height          =   315
      Left            =   1650
      TabIndex        =   28
      Top             =   1050
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.OptionButton optMode 
      BackColor       =   &H00800000&
      Caption         =   "Search all subdir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   1
      Left            =   7770
      TabIndex        =   27
      ToolTipText     =   "List all image files in subdir, of a search pattern"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.OptionButton optMode 
      BackColor       =   &H00800000&
      Caption         =   "Images in dir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   0
      Left            =   5340
      TabIndex        =   26
      ToolTipText     =   "Disp all images in dir, of a file pattern"
      Top             =   1080
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox picAuto 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   5970
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   1
      Top             =   4500
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   6660
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   0
      Top             =   4500
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Frame fraImagesPanel 
      Caption         =   "List panel - images in directory"
      ForeColor       =   &H00800000&
      Height          =   6015
      Left            =   5220
      TabIndex        =   4
      Top             =   1470
      Visible         =   0   'False
      Width           =   4455
      Begin VB.PictureBox picProgressContainer 
         AutoRedraw      =   -1  'True
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   120
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   278
         TabIndex        =   63
         Top             =   840
         Width           =   4230
         Begin VB.PictureBox picProgress 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFF80&
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   0
            ScaleHeight     =   14
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   64
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.PictureBox picImagesInDirContainer 
         Height          =   585
         Left            =   120
         ScaleHeight     =   525
         ScaleWidth      =   4185
         TabIndex        =   22
         Top             =   240
         Width           =   4245
         Begin VB.Label lblImagesPanelDir 
            Caption         =   "lblImagesPanelDir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   30
            TabIndex        =   23
            Top             =   0
            Width           =   4125
         End
      End
      Begin VB.CommandButton cmdListImages 
         Height          =   345
         Left            =   540
         Picture         =   "picViewer.frx":079A
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Disp all  images in dir"
         Top             =   1230
         Width           =   345
      End
      Begin VB.PictureBox picAutoImagesContainer 
         Height          =   405
         Left            =   120
         Picture         =   "picViewer.frx":0964
         ScaleHeight     =   345
         ScaleWidth      =   315
         TabIndex        =   16
         Top             =   1200
         Width           =   375
         Begin VB.CommandButton cmdAutoImagesOn 
            Height          =   345
            Left            =   0
            Picture         =   "picViewer.frx":0B2E
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Toggle auto images on/off.  Current status is On."
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton cmdAutoImagesOff 
            Height          =   345
            Left            =   0
            Picture         =   "picViewer.frx":0CF8
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Toggle auto images on/off.  Current status is Off"
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.ComboBox cboPicListLot 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Select lot number"
         Top             =   5610
         Width           =   795
      End
      Begin VB.VScrollBar VsbPicList 
         Height          =   3795
         Left            =   4110
         TabIndex        =   7
         Top             =   1710
         Width           =   240
      End
      Begin VB.PictureBox picListZ 
         Height          =   3795
         Left            =   90
         ScaleHeight     =   249
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   263
         TabIndex        =   5
         Top             =   1710
         Width           =   4005
         Begin VB.PictureBox picListX 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   4995
            Left            =   0
            ScaleHeight     =   331
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   259
            TabIndex        =   6
            Top             =   0
            Width           =   3915
         End
      End
      Begin VB.PictureBox piccmdPanelRefContainer 
         Height          =   315
         Left            =   2070
         ScaleHeight     =   255
         ScaleWidth      =   1305
         TabIndex        =   43
         Top             =   5610
         Width           =   1365
         Begin VB.CommandButton cmdPanelRef 
            BackColor       =   &H00C0C0C0&
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   990
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton cmdPanelRef 
            BackColor       =   &H00C0C0C0&
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   660
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton cmdPanelRef 
            BackColor       =   &H00C0C0C0&
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   330
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   0
            Width           =   315
         End
         Begin VB.CommandButton cmdPanelRef 
            BackColor       =   &H00C0C0C0&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   0
            Width           =   315
         End
      End
      Begin VB.Label lblLotRef 
         Alignment       =   1  'Right Justify
         Caption         =   "Lot:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   300
         TabIndex        =   62
         Top             =   5700
         Width           =   645
      End
      Begin VB.Label lblImagesPanelHelp 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   " ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   4110
         TabIndex        =   15
         ToolTipText     =   "Help"
         Top             =   5730
         Width           =   255
      End
      Begin VB.Label lblPicsCount 
         Alignment       =   1  'Right Justify
         Caption         =   "Total images in dir:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1020
         TabIndex        =   10
         Top             =   1170
         Width           =   3225
      End
      Begin VB.Label lblPicsOnPanel 
         Alignment       =   1  'Right Justify
         Caption         =   "Images in current lot:"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   960
         TabIndex        =   9
         Top             =   1380
         Width           =   3315
      End
   End
   Begin VB.Frame fraSearch 
      Caption         =   "List panel - search all subdirectories"
      ForeColor       =   &H00800000&
      Height          =   5955
      Left            =   5220
      TabIndex        =   2
      Top             =   1530
      Visible         =   0   'False
      Width           =   4395
      Begin VB.PictureBox picSearchAllSubdirContainer 
         Height          =   705
         Left            =   150
         ScaleHeight     =   645
         ScaleWidth      =   4065
         TabIndex        =   24
         Top             =   330
         Width           =   4125
         Begin VB.Label lblSearchPanelDir 
            Caption         =   "lblSearchPanelDir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   30
            TabIndex        =   25
            Top             =   60
            Width           =   3945
         End
      End
      Begin VB.CommandButton cmdStopSearch 
         Height          =   315
         Left            =   2940
         Picture         =   "picViewer.frx":0EC2
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Stop search"
         Top             =   1350
         Width           =   315
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   315
         Left            =   2460
         Picture         =   "picViewer.frx":1204
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Search for an image file pattern, in all subdir"
         Top             =   1350
         Width           =   345
      End
      Begin VB.ComboBox cboSearchPattern 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1350
         TabIndex        =   13
         Text            =   "cboSearchPattern"
         Top             =   1350
         Width           =   975
      End
      Begin VB.ListBox lisFilesFound 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00800000&
         Height          =   3570
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   4110
      End
      Begin VB.Label lblSearchPanelHelp 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   " ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Help"
         Top             =   5670
         Width           =   255
      End
      Begin VB.Label lblSearchPattern 
         Caption         =   "Search pattern"
         Height          =   225
         Left            =   150
         TabIndex        =   12
         Top             =   1410
         Width           =   1185
      End
      Begin VB.Label lblFoundNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Total found: 0"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2730
         TabIndex        =   11
         Top             =   5670
         Width           =   1425
      End
   End
   Begin VB.PictureBox piccmdEditContainer 
      Height          =   315
      Left            =   1290
      ScaleHeight     =   255
      ScaleWidth      =   165
      TabIndex        =   60
      Top             =   3510
      Visible         =   0   'False
      Width           =   225
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00000080&
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         ToolTipText     =   "Edit image with system paint program"
         Top             =   0
         Width           =   165
      End
   End
   Begin VB.Label lblPicSizeH 
      Caption         =   "h="
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   270
      TabIndex        =   59
      ToolTipText     =   "+/- percentage"
      Top             =   3660
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblSearchInProgress 
      Caption         =   "Search in progress ........"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1380
      TabIndex        =   42
      Top             =   1950
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label lblFilListHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   " ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   4740
      TabIndex        =   41
      ToolTipText     =   "Help"
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   210
      Top             =   1050
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblPicSizeW 
      Caption         =   "w= "
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   270
      TabIndex        =   39
      ToolTipText     =   "+/- percentage"
      Top             =   3450
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblFilesCount 
      Alignment       =   1  'Right Justify
      Caption         =   "Files in dir:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3300
      TabIndex        =   38
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblPicViewHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   " ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   4770
      TabIndex        =   37
      ToolTipText     =   "Help"
      Top             =   7260
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   315
      Left            =   5250
      Shape           =   4  'Rounded Rectangle
      Top             =   1050
      Visible         =   0   'False
      Width           =   4425
   End
   Begin VB.Menu popFilList 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu popFilListRenameFile 
         Caption         =   "&Rename file (within same dir)"
      End
      Begin VB.Menu PopFilListSaveFile 
         Caption         =   "&Save file (i.e. make a copy)"
      End
      Begin VB.Menu PopFilListMoveFile 
         Caption         =   "&Move file (to a different dir)"
      End
      Begin VB.Menu popFilListDeleteFile 
         Caption         =   "&Delete file"
      End
   End
   Begin VB.Menu popPicView 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu popPicViewRenameFile 
         Caption         =   "&Rename file (within same dir)"
      End
      Begin VB.Menu PopPicviewSaveFile 
         Caption         =   "&Save file (i.e. make a copy)"
      End
      Begin VB.Menu PopPicviewMoveFile 
         Caption         =   "&Move file (to a different dir)"
      End
      Begin VB.Menu popPicViewDeleteFile 
         Caption         =   "&Delete file"
      End
   End
   Begin VB.Menu popDirList 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu popDirListCreateDir 
         Caption         =   "&Create directory"
      End
      Begin VB.Menu popDirListRenameDir 
         Caption         =   "&Rename directory"
      End
      Begin VB.Menu popDirListDeleteDir 
         Caption         =   "&Delete directory"
      End
   End
End
Attribute VB_Name = "frmPicViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PicViewer.frm
'
' PicViewer.frm
'
' By Herman Liu
'
' A powerful and versatile auto image browser and viewer. Key characteristics:(1) As and
' when you select a file pattern, or switch between folders, all images of the selected
' file pattern of bmp/ ico/ cur/ gif/ jpeg/ jpg/ wfm/ emf in the folder will be displayed
' automatically in a uniform size; (2) If a "*.*" pattern is set, it will detect and
' display pictures of any file extensions in the folder, such as bitmap, icon, gif and
' others; (3) No limit to the No. of images to be displayed in a dir; (4) Listed images
' have their individual file names attached beneath them; (5) Additional view port to
' view a selected image in its original size, with zoom-in and zoom-out; (6) "Image Fit In"
' and "Enlarge Viewport" features; (7) Synchronization between selections in the image list
' panel and that in the file list box; (8) Able to search and list all image files of a selected
' pattern, under all subdirectories of a selected dir;(9) Dir/File copy/delete/rename/move
' facilities within itself; (10) Button to invoke your default paint program; and (11) With Help.
'


'Option Explicit

Public FrmH, We

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
    ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
    ByVal y As Long, ByVal mDestWidth As Long, ByVal mDestHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal mSrcWidth As Long, _
    ByVal mSrcHeight As Long, ByVal dwRop As Long) As Long
    
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
 
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const STRETCH_HALFTONE  As Long = &H4&
Private Const SW_RESTORE        As Long = &H9&

Dim strCurrDir As String
Dim strCurrPattern As String
Dim arrPicFileSpec() As String
Dim picsCount As Integer
Dim picsOnCurrList As Integer
Dim filesCount As Integer

Dim picListRow As Integer
Dim picListCol As Integer
Dim picListSeq As Integer

Dim picListRegionFlag As Boolean
Dim SuspendFlag As Boolean
Dim StopSearchFlag As Boolean
Dim AutoListSwitching As Boolean
Dim MagnifyGlassOn As Boolean
Dim GlassInDrag As Boolean

Dim Xcurr As Single
Dim Ycurr As Single
Dim XRegOffset As Single
Dim YRegOffset As Single

Dim picViewXFileSpec As String

Dim OldX1 As Integer
Dim OldX2 As Integer
Dim OldY1 As Integer
Dim OldY2 As Integer

Dim OrigW As Long
Dim OrigH As Long

Dim X1Reg As Single
Dim X2Reg As Single
Dim Y1Reg As Single
Dim Y2Reg As Single

'---------------------------------------------------------
' Calculation of positions depends on the following values
' For placements and displacements, refer these values.
'---------------------------------------------------------
Const StdW = 32
Const StdH = 32
Const xColNums = 4
Const xRowsPerPanel = 3
Const xPanelNums = 4
Const xRowNums = xRowsPerPanel * xPanelNums
Const xMaxNums = xColNums * xRowNums
Const xGapW = 32
Const xGapH = 32
Const xTextH = 20
Const xTotalGapH = xGapH + xTextH

Dim maxImageWidth As Single
Dim maxImageHeight As Single
Const minImageWidth = 16
Const minImageheight = 16

Const defaultcboResizeIndex = 3
Dim oldcboResizeIndex As Integer

Dim origPicViewLargeBarW As Single
Dim origPicViewLargeContainerW As Single
Dim origPicViewLargeContainerH As Single
Dim origVsbViewLargeH As Single
Dim origVsbViewLargeLeft As Single
Dim origHsbViewLargeW As Single
Dim origHsbViewLargeTop As Single

   ' You may add others to exclude
Const NonPicFileExt = "EXE/COM/DLL/RES/OCX/SYS/BAT/TXT/RTF/LOG/ERR/INI/MDB/PRG/C/HTM/ASP"
Dim mresult
Dim gcdg As Object
'---------------------------------------------------------





Private Sub Form_Activate()
    
'frmPicViewer.Visible = False
'frmPicViewer.Visible = True
    
'cmdAutoImagesOff_Click


    
    
    
'MfilespecX = MpathX + "002-143--001.jpg"
ScanPath.chkSubFolders = vbChecked
    
ScanPath.txtPath.Text = MpathX

ScanPath.cboMask.Text = "*.Jpg"

Call ScanPath.cmdScan_Click

ProgressBar1.Max = ScanPath.ListView1.ListItems.Count
ProgressBar1.Width = frmPicViewer.Width / Screen.TwipsPerPixelX
'For we = 1 To ScanPath.ListView1.ListItems.Count
    
Call Timer1_Timer

Timer1.Enabled = True
Timer1.Interval = 700
    
'Next
    
    'Call cmdViewLarge_Click
    'Call cmdFitIn_Click

End Sub

Private Sub Form_Load()
    
    frmPicViewer.Visible = False
    
    SuspendFlag = True
    Me.ScaleMode = vbPixels
    picTemp.AutoSize = False
    picTemp.Width = picTemp.Width - picTemp.ScaleWidth + StdW
    picTemp.Height = picTemp.Height - picTemp.ScaleHeight + StdH
    picTemp.Visible = False
    
    picAuto.AutoSize = True
    picAuto.Visible = False
    
      ' Keep a copy of existing values, will be used when cmdViewLargeExit
    origPicViewLargeBarW = picViewLargeBar.Width
    origPicViewLargeContainerW = picViewLargeContainer.Width
    origPicViewLargeContainerH = picViewLargeContainer.Height
    origVsbViewLargeH = VsbViewLarge.Height
    origVsbViewLargeLeft = VsbViewLarge.Left
    origHsbViewLargeW = HsbViewLarge.Width
    origHsbViewLargeTop = HsbViewLarge.Top
    
      ' Align
    picViewX.Move 0, 0
    picViewLargeContainer.Top = picViewLargeBar.Top + picViewLargeBar.Height
    picViewLargeContainer.Left = picViewLargeBar.Left
    picViewLarge.Move 0, 0

    VsbPicList.Min = 0
    VsbPicList.Max = (StdH + xTotalGapH) * xRowNums * _
            ((xRowsPerPanel - 1) / xRowsPerPanel) + (StdH + xTotalGapH)
    VsbPicList.SmallChange = 1
    VsbPicList.LargeChange = StdH + xTotalGapH
    
     ' We display xColNums icons each row
    picListX.Width = (StdW + xGapH) * xColNums + xGapW / 2
     ' xRowNums each column, allow extra xTextH pixels for file name.
    picListX.Height = (StdH + xTotalGapH) * xRowNums
    
    cboPattern.Clear
    cboPattern.AddItem "*.*"
    cboPattern.AddItem "*.bmp"
    cboPattern.AddItem "*.gif"
    cboPattern.AddItem "*.jpeg"
    cboPattern.AddItem "*.jpg"
    cboPattern.AddItem "*.ico"
    cboPattern.AddItem "*.cur"
    cboPattern.AddItem "*.wfm"
    cboPattern.AddItem "*.emf"
    cboPattern.ListIndex = 0
    
    cboSearchPattern.Clear
    cboSearchPattern.AddItem "*.bmp"
    cboSearchPattern.AddItem "*.gif"
    cboSearchPattern.AddItem "*.jpeg"
    cboSearchPattern.AddItem "*.jpg"
    cboSearchPattern.AddItem "*.ico"
    cboSearchPattern.AddItem "*.cur"
    cboSearchPattern.AddItem "*.wfm"
    cboSearchPattern.AddItem "*.emf"
    cboSearchPattern.AddItem "*.*"
    cboSearchPattern.ListIndex = 0
    
    Dim i, j
    
    For i = 1.5 To 2.6 Step 0.1
         cboMagnify.AddItem i
    Next i
    cboMagnify.ListIndex = 5
    picGlass.Width = picGlass.Width - picGlass.ScaleWidth + 150
    picGlass.Height = picGlass.Height - picGlass.ScaleHeight + 150
    
    maxImageWidth = Screen.Width / Screen.TwipsPerPixelX
    maxImageHeight = Screen.Height / Screen.TwipsPerPixelY
    i = maxImageWidth / picViewZ.Width
    j = maxImageHeight / picViewZ.Height
    HsbView.Max = (i - 1) * picViewZ.Width
    VsbView.Max = (j - 1) * picViewZ.Height
    
    HsbViewLarge.Max = HsbView.Max
    VsbViewLarge.Max = VsbView.Max
    
    For i = -75 To -25 Step 25
        cboResize.AddItem i
    Next i
    For i = 0 To 800 Step 50
        i = "+" & i
        cboResize.AddItem i
    Next i
    cboResize.ListIndex = defaultcboResizeIndex
    oldcboResizeIndex = defaultcboResizeIndex
      
    lblSearchInProgress.Visible = False
    cmdStopSearch.Visible = False
    StopSearchFlag = False
    picListRegionFlag = False
    AutoListSwitching = False
    picViewXFileSpec = ""
    ClearImagesPanelDisp
    ClearSearchDisp
    
    strCurrDir = ""
    strCurrPattern = ""
    
    'DirList.Path = "S:\00 Art\AutoPix\AnutoPix (Main)\AutoPix 00\AutoPix 0000-Ultra Sexi"
    
    cmdAutoImagesOn.Visible = True
    cmdListImages.Enabled = False
    
    picViewLargeBar.Visible = False
    picProgressContainer.Visible = False
    picViewLargeContainer.Visible = False
    cboPicListLot.Appearance = 0
    lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
    lblImagesPanelDir.Caption = DirList.Path
    lblSearchPanelDir.Caption = DirList.Path
    ListImagesInDir
    Set gcdg = CommonDialog1
    SuspendFlag = False

Unload Form1

'frmPicViewer.Visible = True
Call Form_Activate

End Sub



Private Sub cmdExit_Click()
    Screen.MousePointer = vbNormal
    Unload Me
End Sub




Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lblFilListHelp_Click()
     Dim tmp
     tmp = "HELP:" & vbCrLf & vbCrLf
     tmp = tmp & "(1)  File patterns:" & vbCrLf
     tmp = tmp & "      As and when you change a selection in File Pattern Box, or switch" & vbCrLf
     tmp = tmp & "      between folders, all images of the selected pattern under the folde will be" & vbCrLf
     tmp = tmp & "      displayed automatically in the List Panel (if that panel is set to Images In" & vbCrLf
     tmp = tmp & "      Dir).  There is no limit to the No. of images to be displayed in a folder." & vbCrLf & vbCrLf
     tmp = tmp & "(2)  View image in original size:" & vbCrLf
     tmp = tmp & "      To view an image in its original size, click the file name in the File List Box," & vbCrLf
     tmp = tmp & "      or an image in the List Panel if any.  Selection through the File List Box is" & vbCrLf
     tmp = tmp & "      automatically reflected in the List Panel as well, and vice versa.  (If the panel" & vbCrLf
     tmp = tmp & "      is set to Search All Subdir, then click a displayed file spec there)" & vbCrLf & vbCrLf
     tmp = tmp & "(3)  File functions:" & vbCrLf
     tmp = tmp & "      At the File List Box, right click the mouse to bring up a popup menu.  (You" & vbCrLf
     tmp = tmp & "      may also right click the mouse in the Viewport if there is an image in it)" & vbCrLf & vbCrLf
     tmp = tmp & "(4)  Directory functions:" & vbCrLf
     tmp = tmp & "      At the Dir List Box, right click the mouse to bring up a popup menu." & vbCrLf & vbCrLf
     MsgBox tmp
End Sub



Private Sub lblPicViewHelp_Click()
     Dim tmp
     tmp = "HELP:" & vbCrLf & vbCrLf
     tmp = tmp & "(1)  Usage of Viewport:" & vbCrLf
     tmp = tmp & "      This Viewport is for viewing an image in its original size: (a) When the" & vbCrLf
     tmp = tmp & "      panel is set to Images In Dir, click the file name in File List Box (or an" & vbCrLf
     tmp = tmp & "      image in List Panel).  (b) If the panel is set to Search All Subdir, then" & vbCrLf
     tmp = tmp & "      click a displayed file spec in List Panel." & vbCrLf & vbCrLf
     tmp = tmp & "(2)  Image resizing:" & vbCrLf
     tmp = tmp & "      You may zoom in or zoom out the image until the maximum/mininum width" & vbCrLf
     tmp = tmp & "      &/or height is reached.  Values in combo box are for percentage +/-." & vbCrLf & vbCrLf
     tmp = tmp & "(3)  Image Fit-In or Viewport enlarging:" & vbCrLf
     tmp = tmp & "      Click Fit-In button or Enlarge Viewport button.  A scalable magnifying" & vbCrLf
     tmp = tmp & "      glass is also available in the enlarged viewport." & vbCrLf & vbCrLf
     tmp = tmp & "(4)  Edit image:" & vbCrLf
     tmp = tmp & "      You can click Edit button to invoke your system paint program to edit" & vbCrLf
     tmp = tmp & "       the image - No action will be taken if image is found incompatible" & vbCrLf
     tmp = tmp & "       with your system paint program, e.g. cannot edit ico file." & vbCrLf & vbCrLf
     tmp = tmp & "(5)  File functions:" & vbCrLf
     tmp = tmp & "      If there is an image in the Viewport, right click the mouse there." & vbCrLf
     tmp = tmp & "      (Otherwise use the functions in File List Box)" & vbCrLf & vbCrLf
     MsgBox tmp
End Sub



Private Sub lblImagesPanelHelp_Click()
     Dim tmp
     tmp = "HELP:" & vbCrLf & vbCrLf
     tmp = tmp & "(1)  No. of images per Lot:" & vbCrLf
     tmp = tmp & "      If there are many images, they will be divided into Lots.  Lot Nos. are" & vbCrLf
     tmp = tmp & "      displayed in Lot Box for selection; each lot consists of a max of " & CStr(xMaxNums) & vbCrLf
     tmp = tmp & "      images.  Since there is no limit to the No. of images to be displayed" & vbCrLf
     tmp = tmp & "      under a dir, there is no limit to the No. of Lots a dir may have." & vbCrLf & vbCrLf
     tmp = tmp & "(2)  No. of panels per Lot:" & vbCrLf
     tmp = tmp & "      Eash lot is made of four continuous panels;  upto " & CStr(xMaxNums / xPanelNums) & _
                           " images in each" & vbCrLf
     tmp = tmp & "      visible panel.  Use the vertical scroll bar to move along the panels," & vbCrLf
     tmp = tmp & "      or click a Panel Ref No. to reach that specific panel directly." & vbCrLf & vbCrLf
     tmp = tmp & "(3)  View image in original Size:" & vbCrLf
     tmp = tmp & "      To view an image in its original size, click a displayed image in the" & vbCrLf
     tmp = tmp & "      List Panel.  Alternatively, you may click its file name in File List Box." & vbCrLf
     tmp = tmp & "      Selection through the List Panel is automatically reflected in the File" & vbCrLf
     tmp = tmp & "      List Box as well, and vice versa, i.e. they are synchronized." & vbCrLf & vbCrLf
     MsgBox tmp
End Sub



Private Sub lblSearchPanelHelp_Click()
     Dim tmp
     tmp = "HELP:" & vbCrLf & vbCrLf
     tmp = tmp & "(1)  Search for a file pattern:" & vbCrLf
     tmp = tmp & "      Supply a file pattern in the Search Pattern Box on the" & vbCrLf
     tmp = tmp & "      List Panel (select or type), switch to the directory of" & vbCrLf
     tmp = tmp & "      which all subdirectories are to be searched, then click" & vbCrLf
     tmp = tmp & "      the Search button." & vbCrLf & vbCrLf
     tmp = tmp & "(2)  View image of a found file spec:" & vbCrLf
     tmp = tmp & "      Click a displayed file spec in the List Panel." & vbCrLf & vbCrLf
     tmp = tmp & "(3)  View full file spec:" & vbCrLf
     MsgBox tmp & "      Double click a displayed file spec." & vbCrLf & vbCrLf
End Sub




Private Sub optMode_Click(Index As Integer)
     If Index = 0 Then
          fraImagesPanel.Visible = True
          fraSearch.Visible = False
          
          ClearViewDisp
          If cmdAutoImagesOn.Visible Then
               AutoListSwitching = True
               ListImagesInDir
               AutoListSwitching = False      ' Set back to normal
          End If
     Else
          fraImagesPanel.Visible = False
          fraSearch.Visible = True
          ClearViewDisp
          ClearImagesPanelDisp
     End If
     FilList.Refresh
End Sub




Private Sub cmdAutoImagesOn_Click()
     cmdAutoImagesOn.Visible = False
     picListX.SetFocus
     
     cmdAutoImagesOff.Visible = True
     ClearImagesPanelDisp
     cmdListImages.Enabled = True
End Sub



Private Sub cmdAutoImagesOff_Click()
     cmdAutoImagesOn.Visible = True
     cmdAutoImagesOff.Visible = False
     AutoListSwitching = True       ' Signal to ListImagesInDir not to early exit
     picListX.SetFocus
     
     DirList_Change                 ' Will call ListImagesInDir there
     AutoListSwitching = False      ' Set back to normal
     cmdListImages.Enabled = False
End Sub



Private Sub cmdListImages_Click()
     AutoListSwitching = True
     ListImagesInDir
     AutoListSwitching = False
     picListX.SetFocus
End Sub




Private Sub DrvList_Change()
     On Error GoTo errHandler               ' Trap e.g. drive not ready
     DirList.Path = drvList.Drive
     Exit Sub
     
errHandler:
     drvList.Drive = DirList.Path
     ErrMsgProc "DrvList_Change"
End Sub



Private Sub DirList_Change()
    FilList.Path = DirList.Path
    If lblSearchInProgress.Visible = False Then
        FilList.Pattern = cboPattern.Text
        ClearViewDisp
        If cmdAutoImagesOn.Visible Then
            ListImagesInDir
        End If
    Else
        FilList.Pattern = cboSearchPattern.Text
    End If
    
    lblImagesPanelDir.Caption = DirList.Path
    lblSearchPanelDir.Caption = DirList.Path
    lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
End Sub



Private Sub DirList_LostFocus()
    DirList.Path = DirList.List(DirList.ListIndex)
End Sub



Private Sub DirList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbRightButton Then
         Exit Sub
    End If
    PopupMenu Me.popDirList
End Sub




Private Sub picViewX_Click()
If Timer1.Enabled = True Then Timer1.Enabled = False: Exit Sub
If Timer1.Enabled = False Then Timer1.Enabled = True

End Sub

Private Sub picViewZ_Click()
If Timer1.Enabled = True Then Timer1.Enabled = False: Exit Sub
If Timer1.Enabled = False Then Timer1.Enabled = True

End Sub

Private Sub popDirListCreateDir_click()
    On Error GoTo errHandler
    Dim currdir As String, newDir As String
    currdir = DirList.List(DirList.ListIndex)
again:
    newDir = InputBox("Type full directory specification:", _
        "Create directory", currdir)
    If newDir = "" Then
         Exit Sub
    End If
    MkDir newDir
    DoEvents
    DirList.Refresh
    Exit Sub
errHandler:
    If Err.Number = 75 Then
        MsgBox "Directory already exists/access error"
        GoTo again
    End If
    ErrMsgProc "popDirListCreateDir_click"
End Sub



Private Sub popDirListRenameDir_click()
    On Error GoTo errHandler
    Dim newDir As String, origDirAsFile As String
    Dim origFullPath As String, origPathDir As String
    origFullPath = DirList.List(DirList.ListIndex)
    origPathDir = PathSection(origFullPath, 1)
    origDirAsFile = PathSection(origFullPath, 2)
    
again:
    newDir = InputBox("Type new name", "Rename directory", origDirAsFile)
    If newDir = "" Then
         Exit Sub
    End If
    If InStr(newDir, "\") <> 0 Then
         MsgBox "Rename directory within same path only"
         GoTo again
    ElseIf PathSection(origPathDir & newDir, 1) <> PathSection(origFullPath, 1) Then
         MsgBox "Rename file within same directory only"
         GoTo again
    End If
    
    newDir = origPathDir & newDir
    Name origFullPath As newDir
    DirList.Path = newDir
    Exit Sub

errHandler:
    ErrMsgProc "popDirListRenameDir_click"
End Sub



Private Sub popDirListDeleteDir_click()
    On Error GoTo errHandler
    If MsgBox("Sure to delete " & DirList.List(DirList.ListIndex) & vbLf & _
           "and all its contents?", vbYesNo + vbQuestion) = vbNo Then
         Exit Sub
    End If
    Dim delDir As String
    delDir = DirList.List(DirList.ListIndex)
    DelFolder delDir
      ' Update
    DirList.Path = drvList.Drive
    Exit Sub
errHandler:
    ErrMsgProc "popDirListDeleteDir_click"
End Sub



Public Sub DelFolder(ByVal inDir As String)
    On Error Resume Next
    Dim mFile As String
       ' Safety
    If Len(Dir(inDir, vbDirectory)) = 0 Then
        Exit Sub
    End If
       ' Find first, i.e. retrieve first entry
    mFile = Dir(inDir & "\", vbDirectory + vbHidden + vbSystem)
       ' Loop through
    Do While mFile <> ""
        If mFile = "." Or mFile = ".." Then
              ' Call Dir again without arguments to return the next file in same dir
            mFile = Dir
        Else
              ' Try to use bitwise comparison to see if mFile is a directory
            If (GetAttr(inDir & "\" & mFile) And vbDirectory) = vbDirectory Then
                 DelFolder inDir & "\" & mFile
                 mFile = Dir(inDir & "\", vbDirectory + vbHidden + vbSystem)
            Else
                   ' Avoid run-time error
                 SetAttr inDir & "\" & mFile, vbNormal
                 Kill inDir & "\" & mFile
                   ' Find next
                 mFile = Dir
            End If
        End If
    Loop
    RmDir inDir
End Sub



Private Sub FilList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbRightButton Then
         Exit Sub
    End If
    PopupMenu Me.popFilList
End Sub




Private Sub popFilListRenameFile_Click()
    On Error GoTo errHandler
    If FilList.ListCount < 1 Then
         MsgBox "No file in current dir"
         Exit Sub
    End If
    If FilList.FileName = "" Then
         MsgBox "No file selected yet"
         Exit Sub
    End If
    
    Dim newName As String, origName As String
    Dim newFileSpec As String, oldDir As String
    Dim mfilespec As String
    mfilespec = FilList.Path & "\" & FilList.FileName
    oldDir = PathSection(mfilespec, 1)
    origName = PathSection(mfilespec, 2)
    
again:
    newName = InputBox("Type new name including ext", "Rename file", origName)
    If newName = "" Then
         Exit Sub
    End If
    
    If InStr(newName, "\") <> 0 Then
         MsgBox "Rename file within same directory only"
         GoTo again
    ElseIf PathSection(oldDir & newName, 1) <> PathSection(mfilespec, 1) Then
         MsgBox "Rename file within same directory only"
         GoTo again
    End If
    
    newFileSpec = oldDir & newName
    Name mfilespec As newFileSpec
    DoEvents
    FilList.Refresh
    If mfilespec = picViewXFileSpec Then
         picViewXFileSpec = newFileSpec
         AutoListSwitching = True
         ListImagesInDir
         AutoListSwitching = False
         ClearViewDisp
    End If
    Exit Sub
    
errHandler:
    ErrMsgProc "popFilListRenameFile"
End Sub




Private Sub popFilListSaveFile_Click()
    On Error GoTo errHandler
    If FilList.ListCount < 1 Then
         MsgBox "No file in current dir"
         Exit Sub
    End If
    If FilList.FileName = "" Then
         MsgBox "No file selected yet"
         Exit Sub
    End If
    
    Dim mPath As String, mfilespec As String
    Dim oldDir As String
    mPath = CurDir
    
    mfilespec = FilList.Path & "\" & FilList.FileName
    
    gcdg.FileName = mfilespec           ' Will show only FilList.FileName though
    gcdg.Filter = "(*.bmp)|*.bmp|(*.ico)|*.ico|(*.*)|*.*|"
    gcdg.DefaultExt = "bmp"
    gcdg.FilterIndex = 1
    gcdg.Flags = cdlOFNOverwritePrompt
    gcdg.CancelError = True
    gcdg.ShowSave
    
afterCreatingDir:
    If mfilespec <> gcdg.FileName Then
        FileCopy mfilespec, gcdg.FileName
           ' Same dir? (dir returned from PathSection includes "\")
        If PathSection(mfilespec, 1) = PathSection(gcdg.FileName, 1) Then
             FilList.Refresh
             AutoListSwitching = True
             ListImagesInDir
             AutoListSwitching = False
             ClearViewDisp
             lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
         End If
         ChDir mPath
    End If
    Exit Sub
    
errHandler:
    If Err.Number <> 32755 Then
         If Err <> 76 Then
              ErrMsgProc "popFilListSaveFile"
         Else                                   ' Dir not already exists
              If MsgBox("Dir " & PathSection(mfilespec, 1) & " does not exist" _
                   & vbCrLf & "Create it?", vbYesNo + vbQuestion) = vbNo Then
                   Exit Sub
              End If
              MkDir PathSection(mfilespec, 1)
              DirList.Refresh
              GoTo afterCreatingDir
         End If
     End If
End Sub




Private Sub popFilListMoveFile_Click()
    On Error GoTo errHandler
    If FilList.ListCount < 1 Then
         MsgBox "No file in current dir"
         Exit Sub
    End If
    If FilList.FileName = "" Then
         MsgBox "No file selected yet"
         Exit Sub
    End If
    Dim mPath As String, mFullSpec As String
    Dim oldDir As String
    mPath = CurDir
    Dim mfilespec As String
    mfilespec = FilList.Path & "\" & FilList.FileName
    
    gcdg.FileName = mfilespec
    
    gcdg.Filter = "(*.bmp)|*.bmp|(*.ico)|*.ico|(*.*)|*.*|"
    gcdg.FilterIndex = 1
    gcdg.CancelError = True
    
again:
    gcdg.ShowSave
    
    mFullSpec = drvList.List(drvList.ListIndex) & gcdg.FileName
    
    If PathSection(mFullSpec, 1) = PathSection(mfilespec, 1) Then
        MsgBox "Cannot move to the same directory"
        GoTo again
    End If
    
afterCreatingDir:
    FileCopy mfilespec, gcdg.FileName
    Kill mfilespec
    DoEvents
    ChDir mPath
    FilList.Refresh
    AutoListSwitching = True
    ListImagesInDir
    AutoListSwitching = False
    ClearViewDisp
    lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
    Exit Sub
    
errHandler:
    If Err.Number <> 32755 Then
         If Err <> 76 Then
              ErrMsgProc "popFilListMoveFile_Click"
         Else                                   ' Dir not already exists
              If MsgBox("Directory " & vbCrLf & PathSection(gcdg.FileName, 1) & _
                   " does not exist" & vbCrLf & "Create it?", vbYesNo + _
                   vbQuestion) = vbNo Then
                   Exit Sub
              End If
              MkDir PathSection(gcdg.FileName, 1)
              DirList.Refresh
              GoTo afterCreatingDir
         End If
    End If
End Sub




Private Sub popFilListDeleteFile_Click()
    If FilList.ListCount < 1 Then
         MsgBox "No file in current dir"
         Exit Sub
    End If
    If FilList.FileName = "" Then
         MsgBox "No file selected yet"
         Exit Sub
    End If
    If MsgBox("Sure to delete " & FilList.FileName & vbCrLf, _
           vbYesNo + vbQuestion) = vbNo Then
         Exit Sub
    End If
    Dim mfilespec As String
    
    mfilespec = FilList.Path & "\" & FilList.FileName
    
    Kill mfilespec
    DoEvents
    FilList.Refresh
    AutoListSwitching = True
    ListImagesInDir
    AutoListSwitching = False
    lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
    ClearViewDisp
End Sub




Private Sub picViewZ_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbRightButton Then
         Exit Sub
    End If
    If picViewXFileSpec = "" Then
         Exit Sub
    End If
    PopupMenu Me.popPicView
End Sub




Private Sub picViewX_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picViewZ_MouseUp Button, Shift, x, y

    
End Sub



Private Sub popPicViewRenameFile_Click()
    On Error GoTo errHandler
    Dim newName As String, origName As String
    Dim newFileSpec As String, oldDir As String
    
    oldDir = PathSection(picViewXFileSpec, 1)
    origName = PathSection(picViewXFileSpec, 2)
    
again:
    newName = InputBox("Type new name including ext", "Rename file", origName)
    If newName = "" Then
         Exit Sub
    End If
    
    If InStr(newName, "\") <> 0 Then
         MsgBox "Rename file within same directory only"
         GoTo again
    ElseIf PathSection(oldDir & newName, 1) <> PathSection(picViewXFileSpec, 1) Then
         MsgBox "Rename file within same directory only"
         GoTo again
    End If
    
    newFileSpec = oldDir & newName
    Name picViewXFileSpec As newFileSpec
    DoEvents
    FilList.Refresh
    picViewXFileSpec = newFileSpec
    AutoListSwitching = True
    ListImagesInDir
    AutoListSwitching = False
    ClearViewDisp
    Exit Sub
    
errHandler:
    ErrMsgProc "popPicViewRenameFile"
End Sub



Private Sub popPicViewSaveFile_Click()
    On Error GoTo errHandler
    
    Dim mPath As String, mfilespec As String
    Dim oldDir As String
    mPath = CurDir
    
    gcdg.FileName = picViewXFileSpec
    gcdg.Filter = "(*.bmp)|*.bmp|(*.ico)|*.ico|(*.*)|*.*|"
    gcdg.DefaultExt = "bmp"
    gcdg.FilterIndex = 1
    gcdg.Flags = cdlOFNOverwritePrompt
    gcdg.CancelError = True
    
afterCreatingDir:
    gcdg.ShowSave
    
    SavePicture picViewX.Picture, gcdg.FileName
    DoEvents
    
    mfilespec = gcdg.FileName         ' mfilespec is a full file spec
    
    If mfilespec <> picViewXFileSpec Then
        If PathSection(mfilespec, 1) = PathSection(picViewXFileSpec, 1) Then
             FilList.Refresh
             AutoListSwitching = True
             ListImagesInDir
             AutoListSwitching = False
             lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
             ClearViewDisp
         End If
         ChDir mPath
    End If
    Exit Sub
    
errHandler:
    If Err.Number <> 32755 Then
         If Err <> 76 Then
              ErrMsgProc "popFilListSaveFile"
         Else                                   ' Dir not already exists
              If MsgBox("Directory " & vbCrLf & PathSection(mfilespec, 1) & _
                   " does not exist" & vbCrLf & "Create it?", vbYesNo + _
                   vbQuestion) = vbNo Then
                   Exit Sub
              End If
              MkDir PathSection(mfilespec, 1)
              DirList.Refresh
              GoTo afterCreatingDir
         End If
    End If
End Sub




Private Sub popPicViewMoveFile_Click()
    On Error GoTo errHandler
    Dim mPath As String, mfilespec As String
    Dim oldDir As String
    mPath = CurDir
    
    mfilespec = drvList.List(drvList.ListIndex) & gcdg.FileName
     
    gcdg.FileName = ""
    gcdg.Filter = "(*.bmp)|*.bmp|(*.ico)|*.ico|(*.*)|*.*|"
    gcdg.FilterIndex = 1
    gcdg.CancelError = True
    
again:
    gcdg.ShowSave
    
    If PathSection(mfilespec, 1) = PathSection(picViewXFileSpec, 1) Then
        MsgBox "Cannot move to the same directory"
        GoTo again
    End If
    
afterCreatingDir:
    SavePicture picViewX.Picture, gcdg.FileName
    Kill picViewXFileSpec
    DoEvents
    ChDir mPath
    FilList.Refresh
    AutoListSwitching = True
    ListImagesInDir
    AutoListSwitching = False
    lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
    ClearViewDisp
    Exit Sub
    
errHandler:
    If Err.Number <> 32755 Then
         If Err <> 76 Then
              ErrMsgProc "popPicViewMoveFile_Click"
         Else                                   ' Dir not already exists
              If MsgBox("Directory " & vbCrLf & PathSection(gcdg.FileName, 1) & _
                   " does not exist" & vbCrLf & "Create it?", vbYesNo + _
                   vbQuestion) = vbNo Then
                   Exit Sub
              End If
              MkDir PathSection(gcdg.FileName, 1)
              DirList.Refresh
              GoTo afterCreatingDir
         End If
    End If
End Sub




Private Sub popPicviewDeleteFile_Click()
    If MsgBox("Sure to delete " & picViewXFileSpec, _
           vbYesNo + vbQuestion) = vbNo Then
         Exit Sub
    End If
    Kill picViewXFileSpec
    DoEvents
    FilList.Refresh
    AutoListSwitching = True
    ListImagesInDir
    AutoListSwitching = False
    lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
    ClearViewDisp
End Sub



Private Sub cmdEdit_Click()
    On Error Resume Next
    Dim mfilespec As String
    mfilespec = picViewXFileSpec
    mresult = ShellExecute(Me.hWnd, "Open", mfilespec, &H0&, &H0&, SW_RESTORE)
End Sub




Private Sub cboResize_Click()
    If SuspendFlag = True Then
        Exit Sub
    End If
    If OrigW = 0 Or OrigH = 0 Then
        Exit Sub
    End If
    
    Dim w, h
    Dim newW, newH
    Dim mfactor, morigArea, mnewArea
    Dim tmp
   
    w = picViewX.ScaleWidth
    h = picViewX.ScaleHeight
    
    HsbView.Value = 0
    VsbView.Value = 0
    Select Case Val(cboResize.Text)
       Case Is > 0
           If OrigW > maxImageWidth Or OrigH > maxImageHeight Then
                  MsgBox "Orig image width/height already exceeded the maximum allowed (" & _
                     CStr(maxImageWidth) & "/" & CStr(maxImageHeight) & " pixels)"
                  GoTo earlyExit
           End If
           
            ' Checking existing size
           If w >= maxImageWidth Or h >= maxImageHeight Then
               If OrigW < maxImageWidth Or OrigH < maxImageHeight Then
                  MsgBox "Image width/height cannot exceed the maximum allowed (" & _
                     CStr(maxImageWidth) & "/" & CStr(maxImageHeight) & " pixels)"
                  GoTo earlyExit
                  Exit Sub
               Else
                  If w >= OrigW Or h >= OrigH Then
                      MsgBox "Image width/height exceeded the max allowed (" & _
                         CStr(maxImageWidth) & "/" & CStr(maxImageHeight) & " pixels)"
                      GoTo earlyExit
                      Exit Sub
                  End If
             End If
           End If
      Case Is < 0
           If OrigW <= minImageWidth Or OrigH <= minImageheight Then
                  MsgBox "Orig image width/height already reached the maximum allowed (" & _
                     CStr(minImageWidth) & "/" & CStr(minImageheight) & " pixels)"
                  GoTo earlyExit
                  Exit Sub
           End If
           
           If w <= minImageWidth Or h <= minImageheight Then
              If OrigW > minImageWidth Or OrigH > minImageheight Then
                  MsgBox "Image width/height cannot go below the minimum allowed (" & _
                     CStr(minImageWidth) & "/" & CStr(minImageheight) & " pixels)"
                  GoTo earlyExit
                  Exit Sub
             Else
                  If w <= OrigW Or h <= OrigH Then
                      MsgBox tmp & "Image width/height below the minimum allowed (" & _
                         CStr(minImageWidth) & "/" & CStr(minImageheight) & " pixels)"
                      GoTo earlyExit
                      Exit Sub
                  End If
             End If
          End If
      Case Else
          If w <> OrigW Or h <> OrigH Then
             picViewX.Width = picViewX.Width - w + OrigW
             picViewX.Height = picViewX.Height - h + OrigH
             picViewX.Picture = LoadPicture()
             picViewX.Picture = picCopy.Picture
          End If
          picViewX.SetFocus
          Exit Sub
    End Select
    
    morigArea = OrigW * OrigH
    mnewArea = morigArea * (100 + Val(cboResize.Text)) / 100
    
    mfactor = Sqr(mnewArea / morigArea)
    
    newW = OrigW * mfactor
    newH = OrigH * mfactor
    
    If Val(cboResize.Text) > 0 Then
         If newW >= maxImageWidth Or newH >= maxImageHeight Then
             If Not (OrigW > newW Or OrigH > newH) Then
                 MsgBox "Will exceed max allowed " & CStr(maxImageWidth) & "/" & _
                     CStr(maxImageHeight) & " pixels"
                 GoTo earlyExit
                 Exit Sub
             End If
         End If
    Else
         If newW <= minImageWidth Or newH <= minImageheight Then
             If Not (OrigW < newW Or OrigH < newH) Then
                 MsgBox "Will fall below min allowed " & CStr(minImageWidth) & "/" & _
                     CStr(minImageheight) & " pixels"
                 GoTo earlyExit
                 Exit Sub
             End If
         End If
    End If
    
    ''Screen.MousePointer = vbHourglass

    picTemp.Picture = LoadPicture()
    picTemp.Cls
    picTemp.Width = picTemp.Width - picTemp.ScaleWidth + newW
    picTemp.Height = picTemp.Height - picTemp.ScaleHeight + newH
    
    mresult = StretchPic(picTemp, newW, newH, picViewX, w, h)
    If mresult = 0 Then
        GoTo earlyExit
    End If
    
    picViewX.Picture = LoadPicture()
    picViewX.Cls
    picViewX.Width = picViewX.Width - picViewX.ScaleWidth + newW
    picViewX.Height = picViewX.Height - picViewX.ScaleHeight + newH
    'picViewX.Width = frmPicViewer.Width
    'picViewX.Height = frmPicViewer.Height
    
    picTemp.Picture = picTemp.Image
    
    BitBlt picViewX.hdc, 0, 0, newW, newH, picTemp.hdc, 0, 0, vbSrcCopy
    
    picTemp.Cls
    picViewX.SetFocus
    
    oldcboResizeIndex = cboResize.ListIndex
    Screen.MousePointer = vbDefault
    Exit Sub
    
earlyExit:
    SuspendFlag = True
    cboResize.ListIndex = oldcboResizeIndex
    SuspendFlag = False
    picViewX.SetFocus
End Sub




Private Sub cmdOrigSize_Click()
    HsbView.Value = 0
    VsbView.Value = 0
    picViewX.Width = picViewX.Width - picViewX.ScaleWidth + OrigW
    picViewX.Height = picViewX.Height - picViewX.ScaleHeight + OrigH
    picViewX.Picture = LoadPicture()
    picViewX.Picture = picCopy.Image
    picViewX.SetFocus
    SuspendFlag = True
    cboResize.ListIndex = defaultcboResizeIndex
    oldcboResizeIndex = defaultcboResizeIndex
    SuspendFlag = False
End Sub




Private Sub cmdFitIn_Click()
    Dim w, h
    Dim newW, newH
    Dim i
        
    'OrigW = picViewX.ScaleWidth
    'OrigH = picViewX.ScaleHeight
    'OrigW = picViewZ.ScaleWidth
    'OrigH = picViewZ.ScaleHeight

    
    'Screen.MousePointer = vbHourglass
    HsbView.Value = 0
    VsbView.Value = 0
    If OrigW <= picViewZ.ScaleWidth And OrigH <= picViewZ.ScaleHeight And 1 = 1 Then
        picViewX.Picture = LoadPicture()
        picViewX.Width = picViewX.Width - picViewX.ScaleWidth + OrigW
        picViewX.Height = picViewX.Height - picViewX.ScaleHeight + OrigH
        picViewX.Width = picViewZ.Width
        picViewX.Height = picViewZ.Height
        
        
        picViewX.Picture = picCopy.Picture
        
        
        
        w = picCopy.ScaleWidth
        h = picCopy.ScaleHeight
        
        yyg1 = picViewZ.Width / Screen.TwipsPerPixelX
        xxg1 = picViewZ.Height / Screen.TwipsPerPixelY
        yyt1 = picViewZ.Width / Screen.TwipsPerPixelX
        xxt1 = picViewZ.Height / Screen.TwipsPerPixelY
        'w = yyg1
        'h = xxg1
        'pk = 0
        'Do
        '    pl = 0: pk = 0
        '    If w > yyg1 Then xxg1 = xxg1 * 1.001: yyg1 = yyg1 * 1.001: pl = 1
        '    If h > xxg1 Then xxg1 = xxg1 * 1.001: yyg1 = yyg1 * 1.001:  pk = 1
       ' Loop Until pl = 0 Or pk = 0
        
        pk = 0
        Do
        pl = 0
        If w < yyg1 Or h < xxg1 Then
        xxg1 = xxg1 / 1.001: yyg1 = yyg1 / 1.001: pl = 1: pk = 1
        'w = w * 1.001: h = h * 1.001: pl = 1: pk = 1
        End If
        Loop Until pl = 0

        



        GoTo skip
        If pk = 0 Then
        Do
        DoEvents
        pl = 0: pk = 0
        If w < yyg1 Or h < xxg1 Then
        xxg1 = xxg1 / 1.01: yyg1 = yyg1 / 1.01: pl = 1: pk = 1
        End If
        Loop Until pl = 0
        End If
        



skip:
        
        'i = 0
        'Do While newW > picViewZ.ScaleWidth Or newH > picViewZ.ScaleHeight
         '    i = i + 1
         '    newW = OrigW * (100 - i) / 100
          '   newH = OrigH * (100 - i) / 100
        'Loop
        'yyg1 = newW
        'xxg1 = newH

'        If pk = 0 Then
'        Do
'            pl = 0: pk = 0
'            If w >= xxg1 Then xxg1 = xxg1 + 1: yyg1 = yyg1 + 1: pl = 1
'            If h >= yyg1 Then xxg1 = xxg1 + 1: yyg1 = yyg1 + 1: pk = 1: pl = 1
'        Loop Until pl = 0 Or pk = 0
'        End If
        
        mresult = StretchPic(picViewX, yyg1, xxg1, picCopy, w, h)
        A = A
        picViewX.Width = yyg1
        picViewX.Height = xxg1
        'frmpicviewer.width
        picViewX.Left = ((frmPicViewer.Width / Screen.TwipsPerPixelX) / 2) - (picViewX.Width / 2)
        picViewX.Top = ((picViewZ.Height / Screen.TwipsPerPixelY) / 2) - (picViewX.Height / 2) - FrmH
    
    
    Else
        w = picCopy.ScaleWidth
        h = picCopy.ScaleHeight
        
        newW = OrigW
        newH = OrigH
        
        i = 0
        Do While newW > picViewZ.ScaleWidth Or newH > picViewZ.ScaleHeight
             i = i + 1
             newW = OrigW * (100 - i) / 100
             newH = OrigH * (100 - i) / 100
             If newW < 16 Or newH < 16 Then
                 Screen.MousePointer = vbDefault
                 MsgBox "Due to the relative proportion of width to height of this image," & _
                    vbCrLf & "unable to implement Fit In"
                    Exit Sub
             End If
        Loop
             
        picViewX.Picture = LoadPicture()
        picViewX.Width = picViewX.Width - picViewX.ScaleWidth + newW
        picViewX.Height = picViewX.Height - picViewX.ScaleHeight + newH
        'picViewX.Width = frmPicViewer.Width
        'picViewX.Height = frmPicViewer.Height
        'newW = frmPicViewer.Width / 50
        'newH = frmPicViewer.Height / 50

        mresult = StretchPic(picViewX, newW, newH, picCopy, w, h)
 '       mresult = StretchPic(picViewX, picViewX.ScaleWidth, picViewX.ScaleHeight, picCopy, w, h)
        

        
        If mresult = 0 Then
            GoTo errHandler
        End If
        
        picViewX.Picture = picViewX.Image
    
'        picViewX.SetFocus
        Screen.MousePointer = vbDefault
    End If
    
    Exit Sub
errHandler:
End Sub



Private Sub cmdFitIn2_Click()
    Dim w, h
    Dim newW, newH
    Dim i
        
'    OrigW = picViewX.ScaleWidth
 '   OrigH = picViewX.ScaleHeight

    
    'Screen.MousePointer = vbHourglass
    HsbView.Value = 0
    VsbView.Value = 0
    If OrigW <= picViewZ.ScaleWidth And OrigH <= picViewZ.ScaleHeight Then
        picViewX.Picture = LoadPicture()
        picViewX.Width = picViewX.Width - picViewX.ScaleWidth + OrigW
        picViewX.Height = picViewX.Height - picViewX.ScaleHeight + OrigH
        
        
        
        picViewX.Picture = picCopy.Picture
    Else
        w = picCopy.ScaleWidth
        h = picCopy.ScaleHeight
        
        newW = OrigW
        newH = OrigH
        
        i = 0
        Do While newW > picViewZ.ScaleWidth Or newH > picViewZ.ScaleHeight
             i = i + 1
             newW = OrigW * (100 - i) / 100
             newH = OrigH * (100 - i) / 100
             If newW < 16 Or newH < 16 Then
                 Screen.MousePointer = vbDefault
                 MsgBox "Due to the relative proportion of width to height of this image," & _
                    vbCrLf & "unable to implement Fit In"
                    Exit Sub
             End If
        Loop
             
        picViewX.Picture = LoadPicture()
        picViewX.Width = picViewX.Width - picViewX.ScaleWidth + newW
        picViewX.Height = picViewX.Height - picViewX.ScaleHeight + newH
        'picViewX.Width = frmPicViewer.Width
        'picViewX.Height = frmPicViewer.Height
        'newW = frmPicViewer.Width / 50
        'newH = frmPicViewer.Height / 50

        mresult = StretchPic(picViewX, newW, newH, picCopy, w, h)
 '       mresult = StretchPic(picViewX, picViewX.ScaleWidth, picViewX.ScaleHeight, picCopy, w, h)
        

        
        If mresult = 0 Then
            GoTo errHandler
        End If
        
        picViewX.Picture = picViewX.Image
    
'        picViewX.SetFocus
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
End Sub




Private Sub cmdViewLarge_Click()
    Dim w, h
    Dim newW, newH
    Dim i
    
    HsbView.Value = 0
    VsbView.Value = 0
    HsbViewLarge.Value = 0
    VsbViewLarge.Value = 0
    
    w = Me.ScaleWidth
    h = Me.ScaleHeight
    Me.WindowState = vbMaximized
    newW = Me.ScaleWidth
    newH = Me.ScaleHeight
    
    picViewLargeBar.Width = picViewLargeBar.Width + (newW - w)
    picViewLargeContainer.Width = picViewLargeContainer.Width + (newW - w)
    picViewLargeContainer.Height = picViewLargeContainer.Height + (newH - h)
    VsbViewLarge.Height = VsbViewLarge.Height + (newH - h)
    VsbViewLarge.Left = VsbViewLarge.Left + (newW - w)
    HsbViewLarge.Width = HsbViewLarge.Width + (newW - w)
    HsbViewLarge.Top = HsbViewLarge.Top + (newH - h)
    
    picViewLarge.Picture = LoadPicture()
    picViewLarge.Picture = picCopy.Picture
    picViewLargeBar.Visible = True
    picViewLargeContainer.Visible = True
    DoEvents
End Sub



Private Sub cmdMagnifyGlass_Click()
    If MagnifyGlassOn Then
         MagnifyGlassOn = False
            ' Put orig back
         picViewLarge.Picture = LoadPicture()
         BitBlt picViewLarge.hdc, 0, 0, picViewLarge.ScaleWidth, picViewLarge.ScaleHeight, _
             picCopy.hdc, 0, 0, vbSrcCopy        'Recopy from orig source
         picViewLarge.Picture = picViewLarge.Image
            ' Allow cboMagnify
         cboMagnify.Enabled = True
    Else
         MagnifyGlassOn = True
            ' Disallow cboMagnfify in session
         cboMagnify.Enabled = False
            ' Set the initialize values of points
            ' and put magnified image there accordingly
         X1Reg = 0: Y1Reg = 0
         X2Reg = picGlass.ScaleWidth
         Y2Reg = picGlass.ScaleHeight
           ' Take the first shot of magnified image
         PrepareSnapShot
           ' Put it in designed position
         UpdateDragging
    End If
End Sub



Private Sub PrepareSnapShot()
    Dim w, h
    Dim maxW, maxH
    Dim ratioW, ratioH
    picGlass.Picture = LoadPicture()
      ' We take an enlarged copy from the clean picCopy, rather than from picViewLarge
      ' We take the enlarged image from a smaller area, e.g. if 2 times, then we take
      ' from 0 to 1/2 of the size of imgGlass in the picViewLarge.  In order to cover
      ' any part of picViewLarge, X2Reg may exceed picViewLarge.ScaleWidth and Y2Reg
      ' may exceed picViewLarge.ScaleHeight.
    w = (X2Reg - X1Reg) / Val(cboMagnify.Text)        ' Enlarge this much
    h = (Y2Reg - Y1Reg) / Val(cboMagnify.Text)
    If X2Reg < picViewLarge.ScaleWidth And Y2Reg < picViewLarge.ScaleHeight Then
         StretchBlt picGlass.hdc, 0, 0, picGlass.ScaleWidth, picGlass.ScaleHeight, _
             picCopy.hdc, X1Reg, Y1Reg, w, h, vbSrcCopy
    Else
         maxW = (picViewLarge.ScaleWidth - 1 - X1Reg) / Val(cboMagnify.Text)
         maxH = (picViewLarge.ScaleHeight - 1 - Y1Reg) / Val(cboMagnify.Text)
         ratioW = maxW / w
         ratioH = maxH / h
         StretchBlt picGlass.hdc, 0, 0, picGlass.ScaleWidth * ratioW, _
             picGlass.ScaleHeight * ratioH, picCopy.hdc, _
             X1Reg, Y1Reg, maxW, maxH, vbSrcCopy
    End If
    picGlass.Picture = picGlass.Image
End Sub



Private Sub UpdateDragging()
       ' Take a fresh copy from picCopy to picViewLarge
    BitBlt picViewLarge.hdc, 0, 0, picViewLarge.ScaleWidth, picViewLarge.ScaleHeight, _
          picCopy.hdc, 0, 0, vbSrcCopy
    picViewLarge.Picture = picViewLarge.Image
       ' Put in the enlarged image
    BitBlt picViewLarge.hdc, X1Reg, Y1Reg, X2Reg, Y2Reg, picGlass.hdc, _
        0, 0, vbSrcCopy
    picViewLarge.Picture = picViewLarge.Image
    DrawRegionLines
End Sub



Private Sub DrawRegionLines()
    picViewLarge.DrawMode = vbInvert
    picViewLarge.Line (X2Reg, Y2Reg)-(X1Reg, Y1Reg), , B
End Sub



Private Sub picViewLarge_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbLeftButton Then
         Exit Sub
    End If
    If MagnifyGlassOn = False Then
         Exit Sub
    End If
    
    If Not ((x >= X1Reg And x <= X2Reg) And (y >= Y1Reg And y <= Y2Reg)) Then
         Exit Sub
    End If
    Xcurr = x
    Ycurr = y
    XRegOffset = 0
    YRegOffset = 0
    GlassInDrag = True
End Sub



Private Sub picViewLarge_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetMousePointer x, y
    If MagnifyGlassOn = False Or GlassInDrag = False Then
         Exit Sub
    End If
    XRegOffset = x - Xcurr
    YRegOffset = y - Ycurr
    X1Reg = X1Reg + XRegOffset
    X2Reg = X2Reg + XRegOffset
    Y1Reg = Y1Reg + YRegOffset
    Y2Reg = Y2Reg + YRegOffset
    Xcurr = x
    Ycurr = y
        ' Set borders within which dragging of image is allowed.  As we take the enlarged
        ' image from a smaller area, e.g. if 2 times, we take from 0 to 1/2 of the size of
        ' imgGlass in the picViewLarge (see PrepareSnapShot), we should take that into
        ' account, so that all parts of picViewLarge can be covered when we drag
    If (X1Reg < 0) Or (X2Reg > picViewLarge.ScaleWidth + picGlass.ScaleWidth / Val(cboMagnify.Text)) Then
         X1Reg = X1Reg - XRegOffset
         X2Reg = X2Reg - XRegOffset
    End If
    If (Y1Reg < 0) Or (Y2Reg > picViewLarge.ScaleHeight + picGlass.ScaleHeight / Val(cboMagnify.Text)) Then
         Y1Reg = Y1Reg - YRegOffset
         Y2Reg = Y2Reg - YRegOffset
    End If
    
    PrepareSnapShot
    UpdateDragging
End Sub



Private Sub picviewlarge_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    GlassInDrag = False
End Sub



Private Sub SetMousePointer(inX, inY)
    If MagnifyGlassOn Then
         Dim w, h
         w = X1Reg + picGlass.ScaleWidth
         h = Y1Reg + picGlass.ScaleHeight
         If (inX > X1Reg And inX < w) And (inY > Y1Reg And inY < h) Then
              picViewLarge.MousePointer = vbSizeAll
         Else
              picViewLarge.MousePointer = vbDefault
         End If
    Else
        picViewLarge.MousePointer = vbDefault
    End If
End Sub



Private Sub cmdViewLargeExit_Click()
    picViewLarge.Picture = LoadPicture()
    picGlass.Picture = LoadPicture()
    
    Me.WindowState = vbNormal
    picViewLargeBar.Width = origPicViewLargeBarW
    picViewLargeContainer.Width = origPicViewLargeContainerW
    picViewLargeContainer.Height = origPicViewLargeContainerH
    VsbViewLarge.Height = origVsbViewLargeH
    VsbViewLarge.Left = origVsbViewLargeLeft
    HsbViewLarge.Width = origHsbViewLargeW
    HsbViewLarge.Top = origHsbViewLargeTop
    
    MagnifyGlassOn = False
    picViewLargeBar.Visible = False
    picViewLargeContainer.Visible = False
End Sub



Private Sub HsbView_Change()
    picViewX.Left = -HsbView.Value
End Sub



Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Timer1.Enabled = True Then Timer1.Enabled = False: Exit Sub
If Timer1.Enabled = False Then Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

    We = We + 1
    If We > ScanPath.ListView1.ListItems.Count Then We = 1
    ProgressBar1.Value = We
    
    a1$ = ScanPath.ListView1.ListItems.Item(We).SubItems(1)
    b1$ = ScanPath.ListView1.ListItems.Item(We)
    
    
    MfilespecX = a1$ + b1$
    
    'MpathX = "S:\00 Art\AutoPix\AutoPix (Main)\AutoPix 00\0-Plus 800K\"
    'MfilespecX = MpathX + "bar4.jpg"

    Call FilList_Click
'    Call cmdFitIn_Click
    frmPicViewer.Left = 0
    frmPicViewer.Top = 400
    frmPicViewer.Width = Screen.Width
    frmPicViewer.Height = Screen.Height - 900
    
    FrmH = ProgressBar1.Height - 5
    
    picViewX.Top = FrmH
    picViewX.Left = 0
    picViewX.Width = frmPicViewer.Width
    picViewX.Height = frmPicViewer.Height - FrmH
    
    picViewZ.Top = FrmH
    picViewZ.Left = 0
    picViewZ.Width = frmPicViewer.Width
    picViewZ.Height = frmPicViewer.Height - FrmH
   
    Call cmdFitIn_Click
    
    frmPicViewer.Visible = True
    frmPicViewer.Refresh
    
    'ee = Timer + 0.7
    'Do
    'DoEvents
    'If Timer > ee Then Exit Do
    'Loop Until ee < Timer
    
'Exit For
    
End Sub

Private Sub VsbView_Change()
    picViewX.Top = -VsbView.Value
End Sub


Private Sub VsbPicList_Change()
    picListX.Top = -VsbPicList.Value
End Sub




Private Sub HsbViewLarge_Change()
    picViewLarge.Left = -HsbViewLarge.Value
End Sub



Private Sub VsbViewLarge_Change()
    picViewLarge.Top = -VsbViewLarge.Value
End Sub




Private Sub PicListx_MouseDown(Button As Integer, Shift As Integer, _
             x As Single, y As Single)
    If Button <> vbLeftButton Then
        Exit Sub
    End If
    
    GetpicSeq x, y
    If picListSeq = 0 Then
        Exit Sub
    End If
    
    SuspendFlag = True
    lisFilesFound.Refresh
    
    If FilList.ListCount > 0 Then
         Dim i, j
         If cboPattern.Text <> "*.*" Then
              i = (CInt(cboPicListLot.Text) - 1) * xMaxNums + picListSeq
              FilList.ListIndex = i - 1
         Else
              i = (CInt(cboPicListLot.Text) - 1) * xMaxNums
              j = CInt(arrPicFileSpec(i + picListSeq - 1, 2))
              FilList.ListIndex = j
         End If
         HsbView.Value = 0
         VsbView.Value = 0
    Else
         picListX.Cls                ' Don't leave a region
    End If
    SuspendFlag = False
    
    If picListRegionFlag = False Then
        DispPicListRegion
    Else
        If OldX1 < x Or OldX2 > x Or OldY1 < y Or OldY2 > y Then
             ClearPicListRegion
        End If
        DispPicListRegion
    End If
   
End Sub



Private Sub cboPattern_Click()
        FilList.Pattern = cboPattern.Text
    If optMode(0).Value = True Then
         ClearViewDisp            ' Clear current picture
         ListImagesInDir          ' List dir related display
    Else
         ClearSearchDisp
    End If
    lblFilesCount.Caption = "Files in dir: " & FilList.ListCount
End Sub



Private Sub ClearViewDisp()
    If picListRegionFlag Then
        ClearPicListRegion
    End If
    picViewX.Picture = LoadPicture()
    HsbView.Value = 0
    VsbView.Value = 0
    
    picViewX.Width = picViewZ.Width
    picViewX.Height = picViewZ.Height
    lblPicSizeW.Caption = ""
    lblPicSizeH.Caption = ""
    
    piccmdEditContainer.Visible = False
    cboResize.Visible = False
    If SuspendFlag = False Then
         cboResize.ListIndex = defaultcboResizeIndex
         oldcboResizeIndex = defaultcboResizeIndex
    End If
    piccmdOrigSizeContainer.Visible = False
    piccmdFitInContainer.Visible = False
    piccmdFitInLargeContainer.Visible = False
    
    picViewXFileSpec = ""
    DoEvents
End Sub
    


Private Sub ClearImagesPanelDisp()
    lblPicsOnPanel.Caption = ""
    lblPicsCount.Caption = ""
    
    piccmdPanelRefContainer.Visible = False
    
    picListX.Picture = LoadPicture()
    picListRegionFlag = False
    VsbPicList.Value = 0
    cboPicListLot.Clear
End Sub



Private Sub ClearSearchDisp()
    lblFoundNum.Caption = ""
    lisFilesFound.Clear
End Sub



Private Sub FilList_Click()
    On Error GoTo errHandler
    If optMode(1).Value = True Then
         MsgBox "When the List Panel is for Search All Subdir," & vbCrLf & _
            "click file spec listed in it instead (when there" & vbCrLf & _
            "are entries there)."
         Exit Sub
    End If
    
    If cboPicListLot.ListCount < 1 Then
         picListX.Cls                ' Don't leave a region
'         Exit Sub
    End If
                 
    ''Screen.MousePointer = vbHourglass
    Dim i, j
    Dim mPath As String
    Dim mfilespec As String
    If Right(DirList.Path, 1) <> "\" Then
         mPath = DirList.Path & "\"
    Else
         mPath = DirList.Path                   ' e.g. root "C:\"
    End If
    mPath = MpathX
    
    
    HsbView.Value = 0
    VsbView.Value = 0
    cboResize.ListIndex = defaultcboResizeIndex
    oldcboResizeIndex = defaultcboResizeIndex
         
    
    GoTo Jump
    
    If SuspendFlag = False And cmdAutoImagesOn.Visible = True Then
         lisFilesFound.Refresh
          
         If picListRegionFlag Then
             ClearPicListRegion
         End If
         
         If cboPattern.Text <> "*.*" Then      ' We may use this approach also
               i = Int(FilList.ListIndex / xMaxNums) + 1
               If i <> CInt(cboPicListLot.Text) Then
                    cboPicListLot.ListIndex = (i - 1)
               End If
         
               picListSeq = FilList.ListIndex Mod xMaxNums
               picListRow = Int(picListSeq / xColNums) + 1
               picListCol = picListSeq Mod xColNums + 1          ' Remainder, & +1
         
         Else               ' If "*.*" then we must use this approach
              
               mfilespec = MfilespecX
'              mfilespec = mPath & FilList.List(FilList.ListIndex)
              
              
              picListSeq = -1                             ' As a marking
              For i = 0 To UBound(arrPicFileSpec) - 1
                    If arrPicFileSpec(i, 1) = mfilespec Then
                         picListSeq = CInt(arrPicFileSpec(i, 3))
                    End If
              Next i
              If picListSeq = -1 Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Not an image file/invalid file"
                    Exit Sub
              End If
              
              i = Int(picListSeq / xMaxNums) + 1
              If i <> CInt(cboPicListLot.Text) Then
                    cboPicListLot.ListIndex = (i - 1)
              End If
         
              picListSeq = picListSeq Mod xMaxNums
              picListRow = Int(picListSeq / xColNums) + 1
              picListCol = picListSeq Mod xColNums + 1          ' Remainder, & +1
         End If
         
         i = Int(picListSeq / (xColNums * xRowsPerPanel))
         
         VsbPicList.Value = 0
         If i > 0 Then
              j = Int(VsbPicList.Max / (xPanelNums - 1)) * (i)
              If j > VsbPicList.Max Then
                   j = VsbPicList.Max
              End If
              VsbPicList.Value = j
         End If
         DoEvents
         
         DispPicListRegion
    End If
    
Jump:
    'mfilespec = mPath & FilList.List(FilList.ListIndex)
    mfilespec = MfilespecX
    
    On Error Resume Next
    Err = False
    picViewX.Picture = LoadPicture()
    picCopy.Picture = LoadPicture()
    picViewX.Picture = LoadPicture(mfilespec)
    If Err Then
         Screen.MousePointer = vbNormal
'         MsgBox "Not an image file/invalid image"
         Exit Sub
    End If
    
    picViewXFileSpec = mfilespec
    
    OrigW = picViewX.ScaleWidth
    OrigH = picViewX.ScaleHeight
    lblPicSizeW.Caption = "w=" & OrigW
    lblPicSizeH.Caption = "h=" & OrigH
    
    picCopy.Picture = LoadPicture()
    picCopy.Picture = picViewX.Image
    
    'piccmdEditContainer.Visible = True
    'cboResize.Visible = True
    'piccmdOrigSizeContainer.Visible = True
    
    If OrigW > picViewZ.ScaleWidth Or OrigH > picViewZ.ScaleHeight Then
     '    piccmdFitInContainer.Visible = True
     '    piccmdFitInLargeContainer.Visible = True
    Else
         piccmdFitInContainer.Visible = False
         piccmdFitInLargeContainer.Visible = False
    End If
    
    Screen.MousePointer = vbNormal
    Exit Sub
    
errHandler:
    Screen.MousePointer = vbNormal
    ErrMsgProc "FilList_Click"
End Sub




Private Sub ListImagesInDir()
    On Error Resume Next
    If lblSearchInProgress.Visible = True Then
         Exit Sub
    End If
    
    If (DirList.Path = strCurrDir) And (cboPattern = strCurrPattern) Then
         If AutoListSwitching = False Then
               Exit Sub
         End If
    End If
    AutoListSwitching = False
    
    ''Screen.MousePointer = vbHourglass
    
    ClearImagesPanelDisp
    
    strCurrDir = DirList.Path
    strCurrPattern = cboPattern
    
    Dim mImageCount As Integer
    Dim mPath As String, mFile As String, mFullSpec As String
    Dim mfilespec As String
    Dim i As Integer, j As Integer, k As Integer
    Dim mPercent
    
    cboPicListLot.Clear
    If Right(DirList.Path, 1) <> "\" Then
        mPath = DirList.Path & "\"
    Else
        mPath = DirList.Path                   ' e.g. root "C:\"
    End If

    mFile = Dir(mPath & cboPattern, vbNormal)  ' Retrieve the first entry.
    If mFile = "" Then                 ' Cannot find first
         Screen.MousePointer = vbDefault
         Exit Sub
    End If
    
    ReDim arrPicFileSpec(FilList.ListCount - 1, 3) As String
    
    mImageCount = 0
    picProgressContainer.Visible = True
    k = FilList.ListCount - 1
    For i = 0 To k
         mFullSpec = mPath & FilList.List(i)
         mPercent = Int(i / (k - 1)) * 100
         PlotProgress mPercent
         Err = False
         If cboPattern.Text = "*.*" Then
              mfilespec = UCase(Right(mFullSpec, 3))
              If InStr(NonPicFileExt, mfilespec) = 0 Then
                   picTemp.Picture = LoadPicture(mFullSpec)
              Else
                   Err = True
              End If
              If Err = False Then
                  arrPicFileSpec(mImageCount, 1) = mFullSpec
                  arrPicFileSpec(mImageCount, 2) = i
                     ' Store valid seq in third dimension
                  arrPicFileSpec(mImageCount, 3) = mImageCount
                  mImageCount = mImageCount + 1
              End If
         Else
              arrPicFileSpec(mImageCount, 1) = mFullSpec
              arrPicFileSpec(mImageCount, 2) = i
              arrPicFileSpec(mImageCount, 3) = i
              mImageCount = mImageCount + 1
         End If
    Next i
    picProgressContainer.Visible = False
    
    picsCount = mImageCount
    
    If picsCount = 0 Then
        Screen.MousePointer = vbDefault
        ReDim arrPicFileSpec(0)
        Exit Sub
    End If
    
    SuspendFlag = True
    If picsCount <= xMaxNums Then
        cboPicListLot.AddItem 1                 ' Lot 1 is enough
    Else
        j = Int(picsCount / xMaxNums) + 1
        For i = 0 To (j - 1)
             cboPicListLot.AddItem (i + 1)
        Next i
    End If
    cboPicListLot.ListIndex = 0
    SuspendFlag = False
    
    PRINTPICLIST
    
    Screen.MousePointer = vbDefault
End Sub



Private Sub PlotProgress(ByVal inPercent As Integer)
    picProgressContainer.Cls
    picProgress.Cls
    picProgress.Width = picProgressContainer.ScaleWidth * (CInt(inPercent) / 100)
    picProgressContainer.CurrentX = picProgressContainer.ScaleWidth / 2
    picProgressContainer.CurrentY = 1
    picProgress.CurrentX = picProgress.ScaleWidth / 2
    picProgress.CurrentY = 1
    If picProgress.CurrentX < picProgressContainer.CurrentX Then
         picProgressContainer.Print CStr(inPercent) & "%"
    Else
         picProgress.Print CStr(inPercent) & "%"
    End If
    DoEvents
End Sub




Private Sub PRINTPICLIST()
    On Error Resume Next
    picListRegionFlag = False
    picListX.Picture = LoadPicture()
    VsbPicList.Value = 0                  ' Ensure scrollbar back to 0 pos
    
    If cboPicListLot.ListCount = 0 Then
         Exit Sub
    End If
    
    Dim x, y
    Dim w, h
    Dim i, j, k
    Dim mIconNo
    Dim mFile
    
    j = (Val(cboPicListLot.Text) - 1) * xMaxNums
    
    picsOnCurrList = picsCount - j
    If picsOnCurrList > xMaxNums Then
         picsOnCurrList = xMaxNums
    End If
    
    lblPicsOnPanel.Caption = "Images in current lot: " & CStr(picsOnCurrList) & _
               "  (max " & CStr(xMaxNums) & " per lot)"
    If cboPicListLot.ListCount = 1 Then
        lblPicsCount.Caption = "Total images: " & CStr(picsCount) & _
              ",  contained in 1 lot"
    Else
        lblPicsCount.Caption = "Total images: " & CStr(picsCount) & _
              ",  contained in " & CStr(cboPicListLot.ListCount) & " lots."
    End If
    
    k = j                                  ' Keep this k value unchanged
    
    x = xGapH / 2
    y = xGapW / 2
    
    For i = j To picsCount - 1
         Err = False
         picAuto.Picture = LoadPicture()
         picAuto.Picture = LoadPicture(arrPicFileSpec(i, 1))
         
         If Err = False Then
             w = picAuto.ScaleWidth
             h = picAuto.ScaleHeight
             picAuto.Refresh
         
             picTemp.Picture = LoadPicture()
             mresult = StretchPic(picTemp, StdW, StdH, picAuto, w, h)
             If mresult = 0 Then
                  mresult = BitBlt(picListX.hdc, x, y, StdW, StdH, 0, 0, 0, vbBlack)
             Else
                  mresult = BitBlt(picListX.hdc, x, y, StdW, StdH, picTemp.hdc, 0, 0, vbSrcCopy)
                  If mresult = 0 Then
                       mresult = BitBlt(picListX.hdc, x, y, StdW, StdH, 0, 0, 0, vbBlack)
                  Else
                       mFile = PathSection(arrPicFileSpec(i, 1), 3)
                       picListX.CurrentX = (x - 7)        ' So that text somewhat in middle of pic
                       picListX.CurrentY = (y + StdH + 2)
                       If Len(mFile) > 8 Then
                           picListX.Print Left(mFile, 8)
                       Else
                           picListX.Print mFile
                       End If
                  End If
             End If
        Else
             mresult = BitBlt(picListX.hdc, x, y, StdW, StdH, 0, 0, 0, vbBlack)
        End If
        
        picListX.Refresh
        
        If (i - k) > (xMaxNums - 1) Then
             Exit For
        End If
         
        x = x + (StdW + xGapW)
        If x > (StdW + xGapW) * xColNums Then
             x = xGapW / 2                        ' Start from first column again
             y = y + (StdH + xTotalGapH)          ' An icon's size plus the gap
        End If
    Next i
    DoEvents
    
    piccmdPanelRefContainer.Visible = True
    j = xMaxNums / xPanelNums
    For i = 0 To j - 1
        cmdPanelRef(i).Enabled = False
    Next i
    j = Int(picsOnCurrList / (xColNums * xRowsPerPanel))
    For i = 0 To j
        cmdPanelRef(i).Enabled = True
    Next i
End Sub



Private Function StretchPic(inDestPic, newW, newH, inSrcPic, w, h)
    On Error GoTo errHandler
    Dim mOrigTone       As Long
    inDestPic.Picture = LoadPicture()
    mOrigTone = SetStretchBltMode(inDestPic.hdc, STRETCH_HALFTONE)
    StretchBlt inDestPic.hdc, 0, 0, newW, newH, inSrcPic.hdc, 0, 0, w, h, vbSrcCopy
    inDestPic.Picture = inDestPic.Image
    mresult = SetStretchBltMode(inDestPic.hdc, mOrigTone)
    StretchPic = 1
    Exit Function
    
errHandler:
    StretchPic = 0
End Function



Private Sub GetpicSeq(inX As Single, inY As Single)
    Dim A As Single, b As Single
    A = picListX.ScaleWidth / xColNums       ' i.e. /3
    b = picListX.ScaleHeight / xRowNums      '      /14
    For picListCol = 1 To xColNums
        If inX < (A * picListCol) Then
             Exit For
        End If
    Next picListCol
    For picListRow = 1 To xRowNums
        If inY < (b * picListRow) Then
             Exit For
        End If
    Next picListRow
    
    picListSeq = (picListRow - 1) * xColNums + picListCol
    If picListSeq > picsOnCurrList Then
        picListSeq = 0
    End If
End Sub




Private Sub DispPicListRegion()
    If cboPicListLot.ListCount = 0 Then
         picListX.Cls
         Exit Sub
    End If
    
    Dim Xstart, Ystart
    Xstart = (picListCol - 1) * (StdW + xGapH) + 16        ' We start from 16
    Ystart = (picListRow - 1) * (StdH + xTotalGapH) + 16
    
    OldX1 = Xstart - 11
    OldX2 = Xstart + StdW + 11
    OldY1 = Ystart - 2
    OldY2 = Ystart + StdH + xTextH + 2
    picListRegionFlag = True
    
    picListX.DrawMode = vbInvert
    picListX.DrawStyle = vbDot
    picListX.Line (OldX1, OldY1)-(OldX2, OldY2), , B
    picListX.DrawMode = vbCopyPen
    picListX.DrawStyle = vbSolid
End Sub



Private Sub ClearPicListRegion()
    picListRegionFlag = False
    picListX.DrawMode = vbInvert
    picListX.DrawStyle = vbDot
    picListX.Line (OldX1, OldY1)-(OldX2, OldY2), , B
    picListX.DrawMode = vbCopyPen
    picListX.DrawStyle = vbSolid
End Sub



Private Sub cbopicListLot_Click()
    If SuspendFlag = False Then
        If Val(cboPicListLot.Text) > 0 Then
               PRINTPICLIST
         End If
    End If
End Sub



Private Sub cmdPanelRef_Click(Index As Integer)
    If cboPicListLot.ListCount < 1 Then
         Exit Sub
    End If
    VsbPicList.Value = Int(VsbPicList.Max / (xPanelNums - 1)) * Index
    picListX.SetFocus
End Sub



Private Sub cmdSearch_Click()
    If Trim(cboSearchPattern.Text) = "" Then
        MsgBox "No search pattern yet"
        Exit Sub
    End If
    
    ''Screen.MousePointer = vbHourglass
    
    lisFilesFound.Clear
    
    StopSearchFlag = False
    SuspendFlag = True     ' Avoid updates in picListX resulting from change of dir
    lisFilesFound.SetFocus
    
    ToggleSearchVisibles False
      
    Dim mFirstPath As String
    Dim mErrDirDiver As Boolean
    Dim mDirCount As Integer
    Dim mNumFiles As Integer
    
    If DirList.Path <> DirList.List(DirList.ListIndex) Then
         DirList.Path = DirList.List(DirList.ListIndex)
         SuspendFlag = False
         Screen.MousePointer = vbDefault
         Exit Sub         ' Exit so user can take a look before re-search
    End If

    FilList.Pattern = cboSearchPattern.Text

    mFirstPath = DirList.Path
    mDirCount = DirList.ListCount

    filesCount = 0                     ' Reset found files indicator
    mErrDirDiver = DirDiver(mFirstPath, mDirCount, "")
    
    If StopSearchFlag Then
         strCurrDir = ""
         SuspendFlag = False
         ToggleSearchVisibles True
         Screen.MousePointer = vbDefault
         Exit Sub
    End If
    
    If mErrDirDiver = True Then
         lisFilesFound.Clear
         filesCount = 0
         DirList.Path = CurDir
         drvList.Drive = DirList.Path        ' Reset the path.
         SuspendFlag = False
         ToggleSearchVisibles True
         Screen.MousePointer = vbDefault
         Exit Sub
    End If
    If filesCount > 0 Then
        lblFoundNum.Caption = "Total found: " & CStr(filesCount)
    Else
        lblFoundNum.Caption = "Total found: 0"
    End If
    FilList.Path = DirList.Path
     
    ToggleSearchVisibles True
    
    If cboPattern.Text <> cboSearchPattern.Text Then
        DirList_Change
    End If
       
    Screen.MousePointer = vbDefault
    If lisFilesFound.ListCount = 0 Then
        MsgBox "No file found matching the search pattern"
    End If
End Sub




Private Function DirDiver(NewPath As String, mDirCount As Integer, BackUp As String) As Integer
    If StopSearchFlag Then
        Exit Function
    End If

    Dim mDirToPeek As Integer
    Dim mAbandon As Integer
    Dim mOldPath As String
    Dim mCurrPath As String
    Dim mEntry As String
    Dim mRetVal As Integer
    Dim i As Integer
    
    DirDiver = False             ' Assumed first. Set to False if there is an error.
    
    mRetVal = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
    If StopSearchFlag Then
        DirDiver = True
        Exit Function
    End If
    
    On Local Error GoTo errHandler:
    
    mDirToPeek = DirList.ListCount    ' How many directories below this?
    
    Do While mDirToPeek > 0 And StopSearchFlag = False
        mOldPath = DirList.Path                      ' Save old path for next recursion.
        DirList.Path = NewPath
        If DirList.ListCount > 0 Then
            DirList.Path = DirList.List(mDirToPeek - 1)
            mAbandon = DirDiver((DirList.Path), mDirCount%, mOldPath)
        End If
        mDirToPeek = mDirToPeek - 1
        If mAbandon = True Then
            StopSearchFlag = True
            Exit Function
        End If
    Loop
    
    If FilList.ListCount Then
        If Len(DirList.Path) <= 3 Then             ' Check for 2 bytes/character
            mCurrPath = DirList.Path                  ' If at root level, leave as is...
        Else
            mCurrPath = DirList.Path + "\"            ' Otherwise put "\" before the filename.
        End If
        For i = 0 To FilList.ListCount - 1        ' Add conforming files in this directory to the list box.
            mEntry = mCurrPath + FilList.List(i)
            lisFilesFound.AddItem mEntry
            filesCount = filesCount + 1
            lblFoundNum = "Total found: " & CStr(filesCount)
       Next i
    End If
    If BackUp <> "" Then                ' If there is a superior directory, move it.
        DirList.Path = BackUp
    End If
    Exit Function
    
errHandler:
    lblSearchInProgress.Visible = False
    cmdStopSearch.Visible = False
    lblFoundNum.Visible = True
    If Err = 7 Then              ' If Out of Memory error occurs, assume the list box just got full.
        MsgBox "Out of memory. Abandoning search..."
    Else                         ' Otherwise display error message and quit.
        ErrMsgProc "frmPicViewer DirDiver"
    End If
End Function




Private Sub cmdStopSearch_Click()
    lisFilesFound.SetFocus
    StopSearchFlag = True
    strCurrDir = ""
End Sub



Private Sub lisFilesFound_dblClick()
    If lisFilesFound.ListCount > 0 Then
        MsgBox "File specification:" & vbCrLf & vbCrLf & lisFilesFound.Text
    End If
End Sub


'lisFilesFound.Text
Private Sub lisFilesFound_Click()
    On Error Resume Next
    If picListRegionFlag Then
         ClearPicListRegion
         picListRegionFlag = False
    End If
    FilList.Refresh
    HsbView.Value = 0
    VsbView.Value = 0
    Err = False
    picViewX.Picture = LoadPicture(lisFilesFound.Text)
    If Err Then
         MsgBox "Not an image file/invalid image"
         Exit Sub
    End If
    
    picViewXFileSpec = lisFilesFound.Text
    
    lblPicSizeW.Caption = "w=" & picViewX.ScaleWidth
    lblPicSizeH.Caption = "h=" & picViewX.ScaleHeight
    
    OrigW = picViewX.ScaleWidth
    OrigH = picViewX.ScaleHeight
    
    picCopy.Picture = LoadPicture()
    picCopy.Picture = picViewX.Image
    picCopy.Picture = picViewX.Image
    
    piccmdEditContainer.Visible = True
    cboResize.Visible = True
    piccmdOrigSizeContainer.Visible = True
    
    If OrigW > picViewZ.ScaleWidth Or OrigH > picViewZ.ScaleHeight Then
         piccmdFitInContainer.Visible = True
         piccmdFitInLargeContainer.Visible = True
    Else
         piccmdFitInContainer.Visible = False
         piccmdFitInLargeContainer.Visible = False
    End If
End Sub




Private Sub ToggleSearchVisibles(OnOff As Boolean)
    lblSearchInProgress.Visible = Not OnOff
    cmdStopSearch.Visible = Not OnOff
    lblFoundNum.Visible = OnOff
    
    Shape1.Visible = OnOff
    cboPattern.Visible = OnOff
    drvList.Visible = OnOff
    DirList.Visible = OnOff
    FilList.Visible = OnOff
    lblFilesCount.Visible = OnOff
End Sub



Private Sub StartLockWindow(ByVal lhWnd As Long)
    mresult = LockWindowUpdate(lhWnd)
End Sub



Private Sub StopLockWindow()
    mresult = LockWindowUpdate(0)
End Sub



Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub


Function PathSection(ByVal inPath As String, inReturnType As Integer)
    Dim DriveLetter As String
    Dim DirPath As String
    Dim FName As String
    Dim Extension As String
    Dim PathLength As Integer
    Dim ThisLength As Integer
    Dim Offset As Integer
    Dim FileNameFound As Boolean

    If inReturnType <> 0 And inReturnType <> 1 And inReturnType <> 2 And inReturnType <> 3 Then
        Err.Raise 1
        Exit Function
    End If
    DriveLetter = ""
    DirPath = ""
    FName = ""
    Extension = ""
    If Mid(inPath, 2, 1) = ":" Then
        DriveLetter = Left(inPath, 2)
        inPath = Mid(inPath, 3)
    End If
    PathLength = Len(inPath)
    For Offset = PathLength To 1 Step -1
        Select Case Mid(inPath, Offset, 1)
            Case ".":
            ThisLength = Len(inPath) - Offset
            If ThisLength >= 1 Then
                Extension = Mid(inPath, Offset, ThisLength + 1)
            End If
            inPath = Left(inPath, Offset - 1)
            Case "\":
            ThisLength = Len(inPath) - Offset
            If ThisLength >= 1 Then
                FName = Mid(inPath, Offset + 1, ThisLength)
                inPath = Left(inPath, Offset)
                FileNameFound = True
                Exit For
            End If
            Case Else
        End Select
    Next Offset
    If FileNameFound = False Then
        FName = inPath
    Else
        DirPath = inPath
    End If
    If inReturnType = 0 Then
        PathSection = DriveLetter
    ElseIf inReturnType = 1 Then
        PathSection = DirPath
    ElseIf inReturnType = 2 Then
        PathSection = FName & Extension
    ElseIf inReturnType = 3 Then
        PathSection = FName
    ElseIf inReturnType = 4 Then
        PathSection = Extension
    End If
End Function

