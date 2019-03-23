VERSION 5.00
Object = "{DEEABA2C-9ED3-11D4-92CF-9E7BD6A8DC73}#1.0#0"; "MEDIAPLAYLIST.OCX"
Begin VB.Form frmEXEform 
   BackColor       =   &H00C0C0C0&
   Caption         =   "MediaShow"
   ClientHeight    =   4635
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7185
   Icon            =   "frmEXEform.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin MediaPlaylist.TimeSlider TimeSlider1 
      Height          =   150
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   265
      DoubleFormat    =   0   'False
      Picture         =   "frmEXEform.frx":27A2
   End
   Begin MediaPlaylist.ViewText ScrollTitle1 
      Height          =   90
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   159
      Leds            =   31
      DoubleSize      =   0   'False
      Picture         =   "frmEXEform.frx":4C0C
   End
   Begin MediaPlaylist.comFunc comFunc1 
      Height          =   195
      Left            =   4920
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   344
   End
   Begin MediaPlaylist.Playlist Playlist1 
      Height          =   2775
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   4895
      BackColor       =   65793
      ForeColor       =   8454143
      ForeColorSel    =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShuffleMode     =   0   'False
      Picture         =   "frmEXEform.frx":6D46
      Row             =   0
      PlayList        =   "C:\Default.m3u"
      WinAMP          =   0   'False
      Repeat          =   2
   End
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4920
      Picture         =   "frmEXEform.frx":7ABC
      ScaleHeight     =   270
      ScaleWidth      =   2325
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.PictureBox PlaylistRes 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2850
      Left            =   120
      Picture         =   "frmEXEform.frx":9BE8
      ScaleHeight     =   2850
      ScaleWidth      =   4200
      TabIndex        =   10
      Top             =   -120
      Visible         =   0   'False
      Width           =   4200
   End
   Begin VB.Timer ScrollTimer 
      Interval        =   250
      Left            =   1440
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   490
      Left            =   1440
      Top             =   2040
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   2040
      Left            =   0
      TabIndex        =   0
      Top             =   2760
      Width           =   7725
      Begin VB.CommandButton cmdMin 
         Caption         =   "-"
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdPlus 
         Caption         =   "+"
         Height          =   255
         Left            =   930
         TabIndex        =   22
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdMedia 
         Caption         =   "T"
         Height          =   255
         Index           =   7
         Left            =   390
         TabIndex        =   9
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdMedia 
         Caption         =   "D"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdMedia 
         Caption         =   "F"
         Height          =   255
         Index           =   6
         Left            =   660
         TabIndex        =   6
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdMedia 
         Caption         =   "Stop"
         Height          =   495
         Index           =   3
         Left            =   2400
         TabIndex        =   5
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdMedia 
         Caption         =   "Next"
         Height          =   495
         Index           =   4
         Left            =   3120
         TabIndex        =   4
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdMedia 
         Caption         =   "Prev."
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdMedia 
         Caption         =   "Play"
         Height          =   495
         Index           =   1
         Left            =   840
         TabIndex        =   2
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdMedia 
         Caption         =   "Pause"
         Height          =   495
         Index           =   2
         Left            =   1680
         TabIndex        =   1
         Top             =   1080
         Width           =   615
      End
      Begin MediaPlaylist.TimeDisplay TimeDisplay1 
         Height          =   195
         Left            =   2400
         TabIndex        =   16
         Top             =   600
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   344
         Leds            =   4
         DoubleSize      =   0   'False
         Minimized       =   0   'False
         Twinkle         =   0   'False
         Picture         =   "frmEXEform.frx":30B9A
      End
      Begin MediaPlaylist.Volume Volume1 
         Height          =   195
         Left            =   4080
         TabIndex        =   17
         Top             =   600
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   344
         DoubleSize      =   0   'False
         Picture         =   "frmEXEform.frx":31B28
      End
      Begin MediaPlaylist.Balance Balance1 
         Height          =   195
         Left            =   4080
         TabIndex        =   18
         Top             =   1080
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   344
         DoubleSize      =   0   'False
         Picture         =   "frmEXEform.frx":47486
      End
      Begin MediaPlaylist.ViewText vtKBPS 
         Height          =   180
         Left            =   6600
         TabIndex        =   19
         Top             =   600
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   318
         Leds            =   3
         DoubleSize      =   -1  'True
         Picture         =   "frmEXEform.frx":5CDE4
      End
      Begin MediaPlaylist.ViewText vtHZ 
         Height          =   180
         Left            =   6600
         TabIndex        =   20
         Top             =   960
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   318
         Leds            =   2
         DoubleSize      =   -1  'True
         Picture         =   "frmEXEform.frx":5EF1E
      End
      Begin MediaPlaylist.ViewText vtSTEREO 
         Height          =   180
         Left            =   6600
         TabIndex        =   21
         Top             =   1320
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   318
         Leds            =   6
         DoubleSize      =   -1  'True
         Picture         =   "frmEXEform.frx":61058
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Menu mnuPlayMode 
      Caption         =   "Play &Mode"
      Begin VB.Menu mnuPlayModeShuffle 
         Caption         =   "&Shuffle"
      End
      Begin VB.Menu mnuRepeat 
         Caption         =   "&Repeat ..."
         Begin VB.Menu mnuRepeatNone 
            Caption         =   "&None"
         End
         Begin VB.Menu mnuRepeatOne 
            Caption         =   "&One"
         End
         Begin VB.Menu mnuRepeatAll 
            Caption         =   "&All"
         End
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnTop 
         Caption         =   "&Always On Top"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileInfo 
         Caption         =   "&File Info"
      End
      Begin VB.Menu mnuFileIntro 
         Caption         =   "File &Intro"
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSortList 
         Caption         =   "Sort &List ..."
         Begin VB.Menu mnuSortListByNameAsc 
            Caption         =   "By Name (Asc)"
         End
         Begin VB.Menu mnuSortListByNameDesc 
            Caption         =   "By Name (Desc)"
         End
         Begin VB.Menu mnuSortListByPathAsc 
            Caption         =   "By Path (Asc)"
         End
         Begin VB.Menu mnuSortListByPathDesc 
            Caption         =   "By Path (Desc)"
         End
         Begin VB.Menu mnuSortListByTimeAsc 
            Caption         =   "By Time (Asc)"
         End
         Begin VB.Menu mnuSortListByTimeDesc 
            Caption         =   "By Time (Desc)"
         End
         Begin VB.Menu mnuSortListByTypeAsc 
            Caption         =   "By Type (Asc)"
         End
         Begin VB.Menu mnuSortListByTypeDesc 
            Caption         =   "By Type (Desc)"
         End
      End
   End
   Begin VB.Menu mnuPlaylist 
      Caption         =   "&Playlist"
      Begin VB.Menu mnuPlaylistLoad 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnuPlaylistSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuPlaylistNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlaylistAddFile 
         Caption         =   "Add &File"
      End
      Begin VB.Menu mnuPlaylistAddDir 
         Caption         =   "Add &Dir"
      End
      Begin VB.Menu mnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "D&elete..."
         Begin VB.Menu mnuDeleteAll 
            Caption         =   "ALL"
         End
         Begin VB.Menu mnuDeleteCrop 
            Caption         =   "CROP"
         End
         Begin VB.Menu mnuDeleteMisc 
            Caption         =   "MISC"
         End
         Begin VB.Menu mnuDeleteSelected 
            Caption         =   "SEL"
         End
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "Select..."
         Begin VB.Menu mnuSelectAll 
            Caption         =   "ALL"
         End
         Begin VB.Menu mnuSelectInv 
            Caption         =   "INV"
         End
         Begin VB.Menu mnuSelectZero 
            Caption         =   "ZERO"
         End
      End
   End
   Begin VB.Menu mnuSkins 
      Caption         =   "&Skins"
      Begin VB.Menu mnuSkinsDir 
         Caption         =   "SKIN DIR"
      End
      Begin VB.Menu mnuSkinsSelect 
         Caption         =   "SELECT SKIN"
      End
   End
End
Attribute VB_Name = "frmEXEform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOACTIVATE = &H10
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1

Dim iTime As Integer
Dim introTimer
Dim Message As String
Dim mediaFont As New StdFont

Private Sub cmdMedia_Click(index As Integer)
    Select Case index
        Case Is = 0: Playlist1.EventPrevious:
        Case Is = 1: Playlist1.EventPlay:
        Case Is = 2: Playlist1.EventPause:
        Case Is = 3: Playlist1.EventStop:
        Case Is = 4: Playlist1.EventNext:
        Case Is = 5: dblWidth:
        Case Is = 6: changeFont:
        Case Is = 7: TestCommands:
    End Select
End Sub

Private Sub dblWidth()
    If TimeSlider1.DoubleFormat = True Then
        TimeSlider1.DoubleFormat = False
        ScrollTitle1.DoubleSize = False
        Volume1.DoubleSize = False
        Balance1.DoubleSize = False
        TimeDisplay1.Top = 720
        TimeDisplay1.DoubleSize = False
        If WindowState <> vbMaximized Then Me.Width = 4095: Me.Height = 5100:
    Else
        TimeSlider1.DoubleFormat = True
        ScrollTitle1.DoubleSize = True
        Volume1.DoubleSize = True
        Balance1.DoubleSize = True
        TimeDisplay1.Top = 555
        TimeDisplay1.DoubleSize = True
        If WindowState <> vbMaximized Then Me.Width = 7800: Me.Height = IIf(Screen.Height >= 7800, 7800, Screen.Height - 600):
    End If
    Select Case TimeDisplay1.Leds
        Case Is = 5: TimeDisplay1.Left = IIf(TimeDisplay1.DoubleSize = True, 1680, 2145):
        Case Is = 6: TimeDisplay1.Left = IIf(TimeDisplay1.DoubleSize = True, 1555, 2100):
        Case Is = 4: TimeDisplay1.Left = IIf(TimeDisplay1.DoubleSize = True, 1950, 2265):
    End Select
End Sub

Private Sub changeFont()
    Static Switch As Boolean
    If Switch = False Then
        Switch = True
        With mediaFont
            .Name = "Times New Roman"
            .Size = 12
        End With
        Set Playlist1.Font = mediaFont
    Else
        Switch = False
        With mediaFont
            .Name = "Tahoma"
            .Size = 7
        End With
        Set Playlist1.Font = mediaFont
    End If
End Sub

Private Sub cmdMin_Click()
    ScrollTitle1.Leds = ScrollTitle1.Leds - 1
End Sub

Private Sub cmdPlus_Click()
    ScrollTitle1.Leds = ScrollTitle1.Leds + 1
End Sub

Private Sub Form_Load()
    If Playlist1.CheckSoundCard = False Then End:
    If GetSetting(App.EXEName, "Settings", "Shortcut", "False") = False Then comFunc1.CreateShortcut App.EXEName, App.Path, stDesktop_Program_NewGroup:
    Playlist1.SkinDir = GetSetting(App.EXEName, "Settings", "SkinDir", App.Path)
    Playlist1.SkinName = GetSetting(App.EXEName, "Settings", "SkinName", "none")
    mnuPlayModeShuffle.Checked = GetSetting(App.EXEName, "PlaySettings", "Shuffle", False)
    If mnuPlayModeShuffle.Checked = True Then Playlist1.ShuffleMode = True:
    mnuRepeatAll.Checked = True
    Playlist1.Playlist = GetSetting(App.EXEName, "PlaySettings", "Playlist", "C:\Default.m3u")
    iTime = GetSetting(App.EXEName, "PlaySettings", "Timer", 0)
    Me.Width = 4095
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Playlist1.Move 0, 0, Me.Width - 120, Me.Height - 2700
    Frame1.Move -15, Me.Height - 2730, Me.Width - 105
    TimeSlider1.Top = Me.Height - 1080
    ScrollTitle1.Top = Me.Height - 2520
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.EXEName, "PlaySettings", "Shuffle", Playlist1.ShuffleMode
    SaveSetting App.EXEName, "PlaySettings", "Playlist", Playlist1.Playlist
    SaveSetting App.EXEName, "PlaySettings", "Timer", iTime
    SaveSetting App.EXEName, "Settings", "SkinDir", Playlist1.SkinDir
    SaveSetting App.EXEName, "Settings", "SkinName", Playlist1.SkinName
End Sub

Private Sub lblTime_Click()
    iTime = iTime + 1
    If iTime = 6 Then iTime = 0:
End Sub

Private Sub mnuDeleteAll_Click()
    Playlist1.DELETE_all
End Sub

Private Sub mnuDeleteCrop_Click()
    Playlist1.DELETE_crop
End Sub

Private Sub mnuDeleteMisc_Click()
    Playlist1.DELETE_misc
End Sub

Private Sub mnuDeleteSelected_Click()
    Playlist1.DELETE_sel
End Sub

Private Sub mnuFileInfo_Click()
    Playlist1.FileInfo
End Sub

Private Sub mnuFileIntro_Click()
    If mnuFileIntro.Checked = False Then
        mnuFileIntro.Checked = True
    Else
        mnuFileIntro.Checked = False
    End If
End Sub

Private Sub mnuOnTop_Click()
    If mnuOnTop.Checked = False Then
        mnuOnTop.Checked = True
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
        SetWindowPos frmEXEform.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
        Message = "Always On Top"
    Else
        mnuOnTop.Checked = False
        SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
        SetWindowPos frmEXEform.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
        Message = "Normal Status"
    End If
    ViewMessage Message, 2
End Sub

Private Sub mnuPlaylistAddDir_Click()
    Playlist1.AddDir
End Sub

Private Sub mnuPlaylistAddFile_Click()
    Playlist1.AddFile
End Sub

Private Sub mnuPlaylistLoad_Click()
    Playlist1.OpenList
End Sub

Private Sub mnuPlaylistNew_Click()
    Playlist1.NewList
End Sub

Private Sub mnuPlaylistSave_Click()
    Playlist1.SaveList
End Sub

Private Sub mnuPlayModeShuffle_Click()
    If mnuPlayModeShuffle.Checked = False Then
        mnuPlayModeShuffle.Checked = True
        Playlist1.ShuffleMode = True
        Message = "Random mode on"
    Else
        mnuPlayModeShuffle.Checked = False
        Playlist1.ShuffleMode = False
        Message = "Random mode off"
    End If
    ViewMessage Message, 2
End Sub

Private Sub mnuRepeatAll_Click()
    Playlist1.Repeat = 2
    mnuRepeatNone.Checked = False
    mnuRepeatOne.Checked = False
    mnuRepeatAll.Checked = True
    ViewMessage "Repeat All", 2
End Sub

Private Sub mnuRepeatNone_Click()
    Playlist1.Repeat = 0
    mnuRepeatNone.Checked = True
    mnuRepeatOne.Checked = False
    mnuRepeatAll.Checked = False
    ViewMessage "Repeat off", 2
End Sub

Private Sub mnuRepeatOne_Click()
    Playlist1.Repeat = 1
    mnuRepeatNone.Checked = False
    mnuRepeatOne.Checked = True
    mnuRepeatAll.Checked = False
    ViewMessage "Repeat one", 2
End Sub

Private Sub mnuSelectAll_Click()
    Playlist1.SELECT_all
End Sub

Private Sub mnuSelectInv_Click()
    Playlist1.SELECT_inv
End Sub

Private Sub mnuSelectZero_Click()
    Playlist1.SELECT_none
End Sub

Private Sub mnuSkinsDir_Click()
    ViewMessage "Select you're SKIN directory", 2
    Playlist1.SKIN_Dir
End Sub

Private Sub mnuSkinsSelect_Click()
    Playlist1.SKIN_Select
    ViewMessage "Select you're SKIN", 2
End Sub

Private Sub mnuSortListByNameAsc_Click()
    Playlist1.OrderBy obNameAsc
    ViewMessage "Order By Name (Asc)", 2
End Sub

Private Sub mnuSortListByNameDesc_Click()
    Playlist1.OrderBy obNameDesc
    ViewMessage "Order By Name (Desc)", 2
End Sub

Private Sub mnuSortListByPathAsc_Click()
    Playlist1.OrderBy obPathAsc
    ViewMessage "Order By Path (Asc)", 2
End Sub

Private Sub mnuSortListByPathDesc_Click()
    Playlist1.OrderBy obPathDesc
    ViewMessage "Order By Path (Desc)", 2
End Sub

Private Sub mnuSortListByTimeAsc_Click()
    Playlist1.OrderBy obTimeAsc
    ViewMessage "Order By Time (Asc)", 2
End Sub

Private Sub mnuSortListByTimeDesc_Click()
    Playlist1.OrderBy obTimeDesc
    ViewMessage "Order By Time (Desc)", 2
End Sub

Private Sub mnuSortListByTypeAsc_Click()
    Playlist1.OrderBy obTypeAsc
    ViewMessage "Order By Type (Asc)", 2
End Sub

Private Sub mnuSortListByTypeDesc_Click()
    Playlist1.OrderBy obTypeDesc
    ViewMessage "Order By Type (Desc)", 2
End Sub

Private Sub ScrollTimer_Timer()
    Static sOldSkinDir As String
    Static sOldSkinName As String
    If sOldSkinDir <> Playlist1.SkinDir Then
        sOldSkinDir = Playlist1.SkinDir
    End If
    If sOldSkinName <> Playlist1.SkinName Then
        sOldSkinName = Playlist1.SkinName
        If sOldSkinDir = "" Or sOldSkinName = "" Then GoTo SCROLL:
        Dim strSRC_Pic As String, strSRC_Path As String
        strSRC_Path = IIf(Right(sOldSkinDir, 1) = "\", sOldSkinDir, sOldSkinDir & "\")
        strSRC_Path = strSRC_Path & IIf(Right(sOldSkinName, 1) = "\", sOldSkinName, sOldSkinName & "\")
        strSRC_Pic = strSRC_Path & "posbar.bmp"
        If comFunc1.FileExist(strSRC_Pic) = True Then
            Set TimeSlider1.Picture = LoadPicture(strSRC_Pic)
        End If
        strSRC_Pic = strSRC_Path & "volume.bmp"
        If comFunc1.FileExist(strSRC_Pic) = True Then
            Set Volume1.Picture = LoadPicture(strSRC_Pic)
        Else
            strSRC_Pic = strSRC_Path & "volbar.bmp"
            If comFunc1.FileExist(strSRC_Pic) = True Then
                Set Volume1.Picture = LoadPicture(strSRC_Pic)
            End If
        End If
        strSRC_Pic = strSRC_Path & "balance.bmp"
        If comFunc1.FileExist(strSRC_Pic) = True Then
            Set Balance1.Picture = LoadPicture(strSRC_Pic)
        Else
            strSRC_Pic = strSRC_Path & "volume.bmp"
            If comFunc1.FileExist(strSRC_Pic) = True Then
                Set Balance1.Picture = LoadPicture(strSRC_Pic)
            Else
                strSRC_Pic = strSRC_Path & "volbar.bmp"
                If comFunc1.FileExist(strSRC_Pic) = True Then
                    Set Balance1.Picture = LoadPicture(strSRC_Pic)
                End If
            End If
        End If
        strSRC_Pic = strSRC_Path & "numbers.bmp"
        If comFunc1.FileExist(strSRC_Pic) = True Then
            Set TimeDisplay1.Picture = LoadPicture(strSRC_Pic)
        Else
            strSRC_Pic = strSRC_Path & "nums_ex.bmp"
            If comFunc1.FileExist(strSRC_Pic) = True Then
                Set TimeDisplay1.Picture = LoadPicture(strSRC_Pic)
            End If
        End If
        strSRC_Pic = strSRC_Path & "text.bmp"
        If comFunc1.FileExist(strSRC_Pic) = True Then
            Set ScrollTitle1.Picture = LoadPicture(strSRC_Pic)
            Set vtHZ.Picture = LoadPicture(strSRC_Pic)
            Set vtKBPS.Picture = LoadPicture(strSRC_Pic)
            Set vtSTEREO.Picture = LoadPicture(strSRC_Pic)
            Set picText.Picture = LoadPicture(strSRC_Pic)
        End If
        strSRC_Pic = strSRC_Path & "pledit.bmp"
        If comFunc1.FileExist(strSRC_Pic) = True Then
            Set Playlist1.Picture = LoadPicture(strSRC_Pic)
            Playlist1.WinAMP = True
            strSRC_Pic = strSRC_Path & "pledit.txt"
            If comFunc1.FileExist(strSRC_Pic) = True Then
                Playlist1.BackColor = comFunc1.Hex_Long(comFunc1.GetKeyValue(strSRC_Pic, "Text", "NormalBG"))
                Playlist1.ForeColor = comFunc1.Hex_Long(comFunc1.GetKeyValue(strSRC_Pic, "Text", "Normal"))
                Playlist1.ForeColorSel = comFunc1.Hex_Long(comFunc1.GetKeyValue(strSRC_Pic, "Text", "Current"))
                With mediaFont
                    .Name = comFunc1.GetKeyValue(strSRC_Pic, "Text", "Font")
                    .Size = 7
                End With
                Set Playlist1.Font = mediaFont
            Else
                Playlist1.BackColor = GetPixel(picText.hdc, CLng(149), CLng(2))
                Playlist1.ForeColor = GetPixel(picText.hdc, CLng(152), CLng(8))
                Playlist1.ForeColorSel = &HFF& '&HC0&
                With mediaFont
                    .Name = "Tahoma"
                    .Size = 7
                End With
                Set Playlist1.Font = mediaFont
            End If
        Else
            Playlist1.WinAMP = True
            Set Playlist1.Picture = PlaylistRes.Picture
            Playlist1.BackColor = 0
            Playlist1.ForeColor = 1565394
            Playlist1.ForeColorSel = &HFF&
            With mediaFont
                .Name = "Tahoma"
                .Size = 7
            End With
            Set Playlist1.Font = mediaFont
        End If
    End If
SCROLL:
    DoEvents
    Static sOldTitle As String
    Static sOldMed As String
    If sOldTitle <> Playlist1.ScrollText Then
        sOldTitle = Playlist1.ScrollText
        ScrollTitle1.ScrollText = sOldTitle
        introTimer = Timer + 10
    End If
    If sOldMed <> Playlist1.CurrentFile Then
        sOldMed = Playlist1.CurrentFile
        Select Case UCase(Right(sOldMed, 3))
            Case Is = "CDA": Volume1.VolumeDevice dvcCD:
            Case Is = "MP1", "MP2", "MP3", "WAV", "AVI": Volume1.VolumeDevice dvcWav:
            Case Is = "MID": Volume1.VolumeDevice dvcMidi
            Case Else
                Volume1.VolumeDevice dvcMaster
        End Select
        vtKBPS.ScrollText = IIf(Playlist1.CurrentkBps < 100, " " & Playlist1.CurrentkBps, Playlist1.CurrentkBps)
        vtHZ.ScrollText = IIf(Playlist1.CurrentHz < 10, " " & Playlist1.CurrentHz, Playlist1.CurrentHz)
        vtSTEREO.ScrollText = IIf(Playlist1.CurrentStereo = True, "STEREO", "MONO")
        vtKBPS.ShowMessage IIf(Playlist1.CurrentkBps < 100, " " & Playlist1.CurrentkBps, Playlist1.CurrentkBps)
        vtHZ.ShowMessage IIf(Playlist1.CurrentHz < 10, " " & Playlist1.CurrentHz, Playlist1.CurrentHz)
        vtSTEREO.ShowMessage IIf(Playlist1.CurrentStereo = True, "STEREO", "MONO")
    End If
    If lblStatus.Caption <> "Playing" Then introTimer = Timer + 10:
    If Volume1.ChangeVolume = False And Balance1.ChangeBalance = False Then
        ScrollTitle1.ScrollNow
    Else
        If Volume1.ChangeVolume = True Then ScrollTitle1.ShowMessage Volume1.Message:
        If Balance1.ChangeBalance = True Then ScrollTitle1.ShowMessage Balance1.Message:
    End If
End Sub

Private Sub TimeDisplay1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Select Case TimeDisplay1.Leds
            Case Is = 4: TimeDisplay1.Leds = 5: TimeDisplay1.Left = IIf(TimeDisplay1.DoubleSize = True, 1680, 2145):
            Case Is = 5: TimeDisplay1.Leds = 6: TimeDisplay1.Left = IIf(TimeDisplay1.DoubleSize = True, 1555, 2100):
            Case Is = 6: TimeDisplay1.Leds = 4: TimeDisplay1.Left = IIf(TimeDisplay1.DoubleSize = True, 1950, 2265):
        End Select
    Else
        iTime = iTime + 1
        If iTime = 6 Then iTime = 0:
        If TimeDisplay1.Leds = 4 And iTime > 2 Then iTime = 0:
    End If
End Sub

Private Sub Timer1_Timer()
    DoEvents
    Dim TimeHasPlayed As Long
    If TimeSlider1.SetStartAt = False Then
        Select Case iTime
            Case Is = 0: TimeHasPlayed = Playlist1.TimeElapsed:
            Case Is = 1: TimeHasPlayed = Playlist1.TimeRemaining:
            Case Is = 2: TimeHasPlayed = Playlist1.TimeSong:
            Case Is = 3: TimeHasPlayed = Playlist1.TimeTotalElapsed:
            Case Is = 4: TimeHasPlayed = Playlist1.TimeTotalRemaining:
            Case Is = 5: TimeHasPlayed = Playlist1.TimeTotal:
        End Select
    Else
        TimeHasPlayed = TimeSlider1.StartAt
    End If
    lblStatus.Caption = Playlist1.Status
    If lblStatus.Caption = "Paused" Then
        TimeDisplay1.Twinkle = True
    Else
        TimeDisplay1.Twinkle = False
    End If
    TimeDisplay1.SetTime TimeHasPlayed
    TimeSlider1.TimeMaximum = Playlist1.TimeSong
    TimeSlider1.TimeElapsed = Playlist1.TimeElapsed
    If mnuFileIntro.Checked = True And lblStatus.Caption = "Playing" Then
        If Timer > introTimer Then Playlist1.EventNext:
    End If
End Sub

Private Sub TestCommands()
    ViewMessage " Loading ...", 2 ' view 2 seconds
    ViewMessage " Test shaped forms", 3
    frmTrans.Show
End Sub

Private Sub ViewMessage(sMessage As String, cTime As Currency)
    ScrollTimer.Enabled = False
    ScrollTitle1.ShowMessage sMessage
    Dim t#
    t# = Timer + cTime
    Do Until Timer > t#: DoEvents: Loop:
    ScrollTimer.Enabled = True
End Sub

