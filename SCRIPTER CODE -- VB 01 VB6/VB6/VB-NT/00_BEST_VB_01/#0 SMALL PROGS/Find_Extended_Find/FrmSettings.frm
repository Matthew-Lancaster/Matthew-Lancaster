VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Properties"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Colours"
      Height          =   2535
      Left            =   5640
      TabIndex        =   19
      Top             =   0
      Width           =   3015
      Begin VB.PictureBox picCFXPBugFix1 
         BorderStyle     =   0  'None
         Height          =   2285
         Left            =   100
         ScaleHeight     =   2280
         ScaleWidth      =   2820
         TabIndex        =   20
         Top             =   175
         Width           =   2815
         Begin VB.Frame Frame6 
            Caption         =   "Selection"
            Height          =   855
            Left            =   20
            TabIndex        =   34
            Top             =   885
            Width           =   1335
            Begin VB.Label LblColour 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Text"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label LblColour 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Back"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   35
               Top             =   480
               Width           =   1095
            End
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   2280
            Left            =   0
            ScaleHeight     =   2280
            ScaleWidth      =   2820
            TabIndex        =   21
            Top             =   -25
            Width           =   2820
            Begin VB.Frame Frame7 
               Caption         =   "General"
               Height          =   855
               Left            =   0
               TabIndex        =   29
               Top             =   0
               Width           =   1335
               Begin VB.Label LblColour 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Text"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   31
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label LblColour 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Back"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   30
                  Top             =   480
                  Width           =   1095
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   "Heading"
               Height          =   1815
               Left            =   1440
               TabIndex        =   22
               Top             =   0
               Width           =   1320
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  Caption         =   "Back Colours"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   28
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.Label LblColour 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Text"
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   27
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.Label LblColour 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "No Find"
                  Height          =   255
                  Index           =   8
                  Left            =   120
                  TabIndex        =   26
                  Top             =   1485
                  Width           =   1095
               End
               Begin VB.Label LblColour 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Pattern Find"
                  Height          =   255
                  Index           =   7
                  Left            =   120
                  TabIndex        =   25
                  Top             =   1230
                  Width           =   1095
               End
               Begin VB.Label LblColour 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Searching"
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   24
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.Label LblColour 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Standard"
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   23
                  Top             =   720
                  Width           =   1095
               End
            End
            Begin VB.Label LblColour 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Restore User"
               Height          =   255
               Index           =   9
               Left            =   135
               TabIndex        =   33
               Top             =   1920
               Width           =   1095
            End
            Begin VB.Label LblColour 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Default"
               Height          =   255
               Index           =   10
               Left            =   1560
               TabIndex        =   32
               Top             =   1920
               Width           =   1095
            End
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Replace"
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   2655
      Begin VB.CheckBox ChkReplace 
         Alignment       =   1  'Right Justify
         Caption         =   "Add Replace To Search"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox ChkReplace 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Blank Warning"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   17
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox ChkReplace 
         Alignment       =   1  'Right Justify
         Caption         =   "  Show Filter Warning"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search History"
      Height          =   1335
      Left            =   2880
      TabIndex        =   8
      Top             =   1680
      Width           =   2655
      Begin VB.PictureBox picCFXPBugFix0 
         BorderStyle     =   0  'None
         Height          =   1080
         Left            =   100
         ScaleHeight     =   1080
         ScaleWidth      =   2460
         TabIndex        =   9
         Top             =   175
         Width           =   2460
         Begin VB.CheckBox ChkSaveHistory 
            Caption         =   "Save"
            Height          =   195
            Left            =   20
            TabIndex        =   12
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton cmdClearHistory 
            Caption         =   "Clear"
            Height          =   315
            Left            =   1320
            TabIndex        =   11
            Top             =   720
            Width           =   1095
         End
         Begin MSComctlLib.Slider SliderHistory 
            Height          =   315
            Left            =   15
            TabIndex        =   10
            Top             =   285
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            LargeChange     =   36
            Min             =   20
            Max             =   200
            SelStart        =   40
            TickFrequency   =   20
            Value           =   40
         End
         Begin VB.Label LblHistory 
            Caption         =   "Size (20-400) :"
            Height          =   255
            Left            =   15
            TabIndex        =   13
            Top             =   45
            Width           =   1935
         End
      End
   End
   Begin VB.CommandButton CmdSetting 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Index           =   2
      Left            =   7680
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton CmdSetting 
      Caption         =   "Apply"
      Height          =   315
      Index           =   1
      Left            =   6480
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   " Found Grid Appearance"
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      Begin VB.CheckBox ChkRemFilters 
         Caption         =   "Remember Filters"
         Height          =   195
         Left            =   2640
         TabIndex        =   41
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox ChkSelectWhole 
         Caption         =   "Find select whole line"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox ChkAutoSelectedText 
         Caption         =   "Auto Selected Text"
         Height          =   195
         Left            =   2640
         TabIndex        =   39
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox ChkShow 
         Caption         =   "Show Component Lines"
         Height          =   195
         Index           =   2
         Left            =   2640
         TabIndex        =   38
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox ChkShow 
         Caption         =   "Show Procedure  Lines"
         Height          =   195
         Index           =   4
         Left            =   2640
         TabIndex        =   37
         Top             =   720
         Width           =   2535
      End
      Begin VB.CheckBox ChkShow 
         Caption         =   "Show Grid Lines"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox ChkShow 
         Caption         =   "Procedure Name"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2535
      End
      Begin VB.CheckBox ChkShow 
         Caption         =   "Component (If more than one)"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox ChkShow 
         Caption         =   "Project (If more than one)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CheckBox ChkLaunchStartup 
      Caption         =   "Launch On Startup"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton CmdSetting 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   5280
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private OrigShowProj                   As Boolean
Private OrigShowComp                   As Boolean
Private OrigShowRout                   As Boolean
Private OrigLaunchStart                As Boolean
Private OrigRemFilt                    As Boolean
Private OrigSaveHist                   As Boolean
Private OrighistDeep                   As Long
Private origbFindSelectWholeLine       As Boolean
Private ApplyClicked                   As Boolean
Private ColourTextForeUnDo             As Long
Private ColourTextBackUnDo             As Long
Private ColourFindSelectForeUnDo       As Long
Private ColourFindSelectBackUnDo       As Long
Private ColourHeadDefaultUnDo          As Long
Private ColourHeadWorkUnDo             As Long
Private ColourHeadPatternUnDo          As Long
Private ColourHeadNoFindUnDo           As Long
Private ColourHeadForeUnDo             As Long

Private Sub ChkAutoSelectedText_Click()

  If Not bLoadingSettings Then
    bAutoSelectText = ChkLaunchStartup.Value = 1
  End If

End Sub

Private Sub ChkLaunchStartup_Click()

  If Not bLoadingSettings Then
    bLaunchOnStart = ChkLaunchStartup.Value = 1
  End If

End Sub

Private Sub ChkRemFilters_Click()

  If Not bLoadingSettings Then
    bRemFilters = ChkRemFilters.Value = 1
  End If

End Sub

Private Sub ChkReplace_Click(Index As Integer)

  If Not bLoadingSettings Then
    Select Case Index
     Case 0
      bFilterWarning = ChkReplace(0).Index = 1
     Case 1
      bBlankWarning = ChkReplace(1).Index = 1
     Case 2
      bReplace2Search = ChkReplace(2).Index = 1
    End Select
  End If

End Sub

Private Sub ChkSaveHistory_Click()

  If Not bLoadingSettings Then
    bSaveHistory = ChkSaveHistory.Value = 1
  End If

End Sub

Private Sub ChkSelectWhole_Click()

  If Not bLoadingSettings Then
    bFindSelectWholeLine = ChkSelectWhole.Value = 1
  End If

End Sub

Private Sub ChkShow_Click(Index As Integer)

  If Not bLoadingSettings Then
    bShowProject = ChkShow(0).Value = 1
    bShowComponent = ChkShow(1).Value = 1
    bShowCompLineNo = ChkShow(2).Value = 1
    bShowRoutine = ChkShow(3).Value = 1
    bShowProcLineNo = ChkShow(4).Value = 1
    bGridlines = ChkShow(5).Value = 1
  End If

End Sub

Private Sub cmdClearHistory_Click()

  mobjDoc.ClearHistory

End Sub

Private Sub CmdSetting_Click(Index As Integer)

  ApplyClicked = False
  Select Case Index
   Case 0
    Me.Hide
    mobjDoc.ApplyChanges
   Case 1
    mobjDoc.ApplyChanges
    ApplyClicked = True
   Case 2
    RestoreOriginals
    Me.Hide
  End Select
  SavePropPosition

End Sub

Private Sub Form_Load()

  mobjDoc.ColoursApply
  ColourTextForeUnDo = ColourTextFore
  ColourTextBackUnDo = ColourTextBack
  ColourFindSelectForeUnDo = ColourFindSelectFore
  ColourFindSelectBackUnDo = ColourFindSelectBack
  ColourHeadDefaultUnDo = ColourHeadDefault
  ColourHeadWorkUnDo = ColourHeadWork
  ColourHeadPatternUnDo = ColourHeadPattern
  ColourHeadNoFindUnDo = ColourHeadNoFind
  ColourHeadForeUnDo = ColourHeadFore
  With Me
    .Left = GetSetting(AppDetails, "Settings", "PropLeft", .Left)
    .Top = GetSetting(AppDetails, "Settings", "PropTop", .Top)
    .Caption = "Properties " & AppDetails
  End With 'Me
  'set safety values for Cancel button
  origbFindSelectWholeLine = bFindSelectWholeLine
  OrigShowProj = bShowProject
  OrigShowComp = bShowComponent
  OrigShowRout = bShowRoutine
  OrigLaunchStart = bLaunchOnStart
  OrigRemFilt = bRemFilters
  OrigSaveHist = bSaveHistory
  OrighistDeep = HistDeep

End Sub

Private Sub Form_Unload(Cancel As Integer)

  If Not ApplyClicked Then
    ' keeps changes if user clicks 'Apply' then uses CaptionBar 'X' button to close
    'otherwise restore
    RestoreOriginals
  End If
  SavePropPosition
  Me.Hide

End Sub

Private Sub LblColour_Click(Index As Integer)

  Select Case Index
   Case 10 'Standard colours
    mobjDoc.ColoursStandard
   Case 9 'Undo
    ColourTextFore = ColourTextForeUnDo
    ColourTextBack = ColourTextBackUnDo
    ColourFindSelectFore = ColourFindSelectForeUnDo
    ColourFindSelectBack = ColourFindSelectBackUnDo
    ColourHeadDefault = ColourHeadDefaultUnDo
    ColourHeadWork = ColourHeadWorkUnDo
    ColourHeadPattern = ColourHeadPatternUnDo
    ColourHeadNoFind = ColourHeadNoFindUnDo
    ColourHeadFore = ColourHeadForeUnDo
   Case Else
    With CommonDialog1
      .Flags = cdlCCRGBInit Or cdlCCFullOpen
      'set current color as default
      Select Case Index
       Case 0
        .Color = ColourTextFore
       Case 1
        .Color = ColourTextBack
       Case 2
        .Color = ColourFindSelectFore
       Case 3
        .Color = ColourFindSelectBack
       Case 4
        .Color = ColourHeadFore
       Case 5
        .Color = ColourHeadDefault
       Case 6
        .Color = ColourHeadWork
       Case 7
        .Color = ColourHeadPattern
       Case 8
        .Color = ColourHeadNoFind
      End Select
      .ShowColor
      'apply new or default colour
      If Not .CancelError Then
        Select Case Index
         Case 0
          ColourTextFore = .Color
         Case 1
          ColourTextBack = .Color
         Case 2
          ColourFindSelectFore = .Color
         Case 3
          ColourFindSelectBack = .Color
         Case 4
          ColourHeadFore = .Color
         Case 5
          ColourHeadDefault = .Color
         Case 6
          ColourHeadWork = .Color
         Case 7
          ColourHeadPattern = .Color
         Case 8
          ColourHeadNoFind = .Color
        End Select
      End If
    End With
  End Select
  mobjDoc.ColoursApply

End Sub

Private Sub RestoreOriginals()

  ChkSelectWhole.Value = Bool2Int(origbFindSelectWholeLine)
  ChkRemFilters.Value = Bool2Int(OrigRemFilt)
  ChkLaunchStartup.Value = Bool2Int(OrigLaunchStart)
  ChkShow(0).Value = Bool2Int(OrigShowProj)
  ChkShow(1).Value = Bool2Int(OrigShowComp)
  ChkShow(2).Value = Bool2Int(OrigShowRout)
  ChkSaveHistory.Value = Bool2Int(OrigSaveHist)
  SliderHistory.Value = OrighistDeep

End Sub

Private Sub SavePropPosition()

  SaveSetting AppDetails, "Settings", "PropLeft", Me.Left
  SaveSetting AppDetails, "Settings", "PropTop", Me.Top

End Sub

Private Sub SliderHistory_Change()

  LblHistory.Caption = "Size (20-400) :" & SliderHistory.Value
  HistDeep = SliderHistory.Value

End Sub

':) Roja's VB Code Fixer V1.1.2 (7/07/2003 2:03:21 AM) 19 + 232 = 251 Lines Thanks Ulli for inspiration and lots of code.
