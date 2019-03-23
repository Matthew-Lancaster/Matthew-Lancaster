VERSION 5.00
Object = "{DEEABA2C-9ED3-11D4-92CF-9E7BD6A8DC73}#1.0#0"; "MEDIAPLAYLIST.OCX"
Begin VB.Form frmPlayList 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "PlayList"
   ClientHeight    =   7545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7545
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picRemMISC 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   600
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   44
      Top             =   6600
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picListOptsNEW 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   2040
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   43
      Top             =   5520
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picListOptsSAVE 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   2040
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   42
      Top             =   5880
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picListOptsLOAD 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   2040
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   41
      Top             =   6240
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picAddURL 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   40
      Top             =   5520
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picAddDIR 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   39
      Top             =   5880
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picAddFILE 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   38
      Top             =   6240
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picRemALL 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   600
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   37
      Top             =   5520
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picRemCROP 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   600
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   36
      Top             =   5880
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picRemFILE 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   600
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   35
      Top             =   6240
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picSelINV 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1080
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   34
      Top             =   5520
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picSelZERO 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1080
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   33
      Top             =   5880
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picSelALL 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1080
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   32
      Top             =   6240
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picMiscSortList 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1560
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   31
      Top             =   5520
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picMiscFileInf 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1560
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   30
      Top             =   5880
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picMiscOpt 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1560
      ScaleHeight     =   270
      ScaleWidth      =   330
      TabIndex        =   29
      Top             =   6240
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picADD 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   120
      ScaleHeight     =   300
      ScaleWidth      =   360
      TabIndex        =   28
      Top             =   6960
      Width           =   360
   End
   Begin VB.PictureBox picREM 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   600
      ScaleHeight     =   300
      ScaleWidth      =   360
      TabIndex        =   27
      Top             =   6960
      Width           =   360
   End
   Begin VB.PictureBox picSEL 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1080
      ScaleHeight     =   300
      ScaleWidth      =   360
      TabIndex        =   26
      Top             =   6960
      Width           =   360
   End
   Begin VB.PictureBox picMISC 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1560
      ScaleHeight     =   300
      ScaleWidth      =   360
      TabIndex        =   25
      Top             =   6960
      Width           =   360
   End
   Begin VB.PictureBox picListOpts 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2040
      ScaleHeight     =   300
      ScaleWidth      =   360
      TabIndex        =   24
      Top             =   6960
      Width           =   360
   End
   Begin MediaPlaylist.ViewText vtTotaltime 
      Height          =   90
      Left            =   3240
      TabIndex        =   23
      Top             =   3120
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   159
      Leds            =   18
      DoubleSize      =   0   'False
      Picture         =   "frmPlayList.frx":0000
   End
   Begin MediaPlaylist.ViewText vtTime 
      Height          =   90
      Left            =   3240
      TabIndex        =   22
      Top             =   3360
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   159
      Leds            =   7
      DoubleSize      =   0   'False
      Picture         =   "frmPlayList.frx":213A
   End
   Begin MediaPlaylist.comFunc comFunc1 
      Height          =   195
      Left            =   5280
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   344
   End
   Begin MediaPlaylist.Playlist Playlist1 
      Height          =   2775
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
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
      Picture         =   "frmPlayList.frx":4274
      Row             =   0
      PlayList        =   "C:\Default.m3u"
      WinAMP          =   0   'False
      Repeat          =   2
   End
   Begin VB.PictureBox picBottomF 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   2040
      ScaleHeight     =   570
      ScaleWidth      =   375
      TabIndex        =   19
      Top             =   4680
      Width           =   375
   End
   Begin VB.PictureBox picClose 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   150
      Left            =   2880
      ScaleHeight     =   150
      ScaleWidth      =   135
      TabIndex        =   18
      Top             =   3480
      Width           =   135
   End
   Begin VB.PictureBox picSlideUp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   90
      Left            =   2400
      ScaleHeight     =   90
      ScaleMode       =   0  'User
      ScaleWidth      =   150
      TabIndex        =   17
      Top             =   3960
      Width           =   150
   End
   Begin VB.PictureBox picSlideDown 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   90
      Left            =   2400
      ScaleHeight     =   90
      ScaleWidth      =   150
      TabIndex        =   16
      Top             =   4080
      Width           =   150
   End
   Begin VB.PictureBox picEvents 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   90
      Left            =   4320
      ScaleHeight     =   90
      ScaleWidth      =   810
      TabIndex        =   15
      Top             =   5520
      Width           =   810
   End
   Begin VB.PictureBox picFormResize 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   255
      Left            =   5280
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   5400
      Width           =   255
   End
   Begin VB.PictureBox picBorderL 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   120
      ScaleHeight     =   870
      ScaleWidth      =   180
      TabIndex        =   13
      Top             =   3600
      Width           =   180
   End
   Begin VB.TextBox SkinName 
      Height          =   285
      Left            =   4800
      TabIndex        =   12
      Top             =   15600
      Width           =   1215
   End
   Begin VB.TextBox SkinPath 
      Height          =   285
      Left            =   4800
      TabIndex        =   11
      Top             =   15120
      Width           =   1215
   End
   Begin VB.PictureBox picText 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2520
      ScaleHeight     =   495
      ScaleWidth      =   1335
      TabIndex        =   10
      Top             =   13080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox PlayListRes 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   120
      Picture         =   "frmPlayList.frx":4FEA
      ScaleHeight     =   2850
      ScaleWidth      =   4200
      TabIndex        =   9
      Top             =   16800
      Width           =   4230
   End
   Begin VB.ListBox SearchList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   4800
      Sorted          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "DBL_Click for PlayList"
      Top             =   14520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picTopM 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1080
      ScaleHeight     =   300
      ScaleWidth      =   1500
      TabIndex        =   7
      Top             =   3120
      Width           =   1500
   End
   Begin VB.PictureBox picTopR 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2640
      ScaleHeight     =   300
      ScaleWidth      =   375
      TabIndex        =   6
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox picTopF 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   600
      ScaleHeight     =   300
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox picTopL 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   120
      ScaleHeight     =   300
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox picBottomM 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   2520
      ScaleHeight     =   570
      ScaleWidth      =   1125
      TabIndex        =   3
      Top             =   4680
      Width           =   1125
   End
   Begin VB.PictureBox picBottomR 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   3720
      ScaleHeight     =   570
      ScaleWidth      =   2250
      TabIndex        =   2
      Top             =   4680
      Width           =   2250
   End
   Begin VB.PictureBox picBottomL 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   120
      ScaleHeight     =   570
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   4680
      Width           =   1875
   End
   Begin VB.PictureBox picSrcPlEdit 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2850
      Left            =   120
      Picture         =   "frmPlayList.frx":2BF9C
      ScaleHeight     =   2850
      ScaleWidth      =   4200
      TabIndex        =   0
      Top             =   13800
      Width           =   4200
   End
End
Attribute VB_Name = "frmPlayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FormFlag As Boolean, FX As Long, FY As Long
Public FormFirst As Boolean, AX As Long, AY As Long
Dim MoveForm As Boolean, MoveFirst As Boolean
Dim i As Integer, picW
Public songHz, songkBps, songStereo As String
Dim mediaFont As New StdFont
Dim fTimer

Private Sub Form_Deactivate()
    HelpButtonsOff
End Sub

Private Sub Form_Initialize()
    Dim StartList As String, errNum As Integer, a
    a = comFunc1.GetWinVersion
    If a >= "4,10" Then WinVerOk = True:
    If a < "4,00" Then End:
    On Error GoTo err_Handle
    errNum = 1
    StartList = GetSetting(App.EXEName, "Visual", "DefList", "C:\Default.m3u")
    If Dir(StartList, vbNormal) = "" Then StartList = comFunc1.PathBackslash(App.Path) & "default.m3u":
    
    a = GetSetting(App.EXEName, "Mode", "Status", "0,2")
    If a <> "" Then
        If Left(a, 1) = "1" Then
            Playlist1.ShuffleMode = True
        Else
            Playlist1.ShuffleMode = False
        End If
        DoEvents
        If Right(a, 1) = "0" Then
            frmMain.imgRepeat.Picture = frmMain.picRepeat(2).Image
            Playlist1.Repeat = 0
        End If
        If Right(a, 1) = "1" Then
            frmMain.imgRepeat.Picture = frmMain.picRepeat(2).Image
            Playlist1.Repeat = 1
        End If
        If Right(a, 1) = "2" Then
            frmMain.imgRepeat.Picture = frmMain.picRepeat(2).Image
            Playlist1.Repeat = 2
        End If
        DoEvents
    End If
    PlayListRes.Picture = picSrcPlEdit.Picture
    errNum = 2
    SkinPath.Text = GetSetting(App.EXEName, "Visual", "SkinDir", App.Path)
    errNum = 3
    SkinName.Text = GetSetting(App.EXEName, "Visual", "Skin", "none")
    Playlist1.SkinDir = SkinPath
    Playlist1.SkinName = SkinName
    Playlist1.Playlist = StartList
    StandardSize
    Exit Sub
err_Handle:
    Select Case errNum
        Case Is = 1: StartList = comFunc1.PathBackslash(App.Path) & "default.m3u":
        Case Is = 2: SkinPath.Text = App.Path:
        Case Is = 3: SkinName.Text = "none":
        Case Else
            'nothing
    End Select
    Resume Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    frmMain.Form_KeyDown KeyCode, Shift
End Sub

Sub ChangePlayList()
    picBottomL_Resize
    picBottomR_Resize
    picEvents_Resize
    picFormResize_Resize
    picBottomF_Resize
    picBottomM_Resize
    picTopL_Resize
    picTopR_Resize
    picClose_Resize
    picTopF_Resize
    picTopM_Resize
    picSlideUp_Resize
    picSlideDown_Resize
    StretchBlt picBorderL.hdc, 0, 0, 12, Int((Me.Height - 870) / 15), picSrcPlEdit.hdc, 0, 42, 12, 29, SRCCOPY
    picBorderL.Refresh
    SetImage picADD, 0, 0, 24, 20, picSrcPlEdit, 13, 79
    SetImage picREM, 0, 0, 24, 20, picSrcPlEdit, 42, 79
    SetImage picSEL, 0, 0, 24, 20, picSrcPlEdit, 71, 79
    SetImage picMISC, 0, 0, 24, 20, picSrcPlEdit, 100, 79
    SetImage picListOpts, 0, 0, 24, 20, picSrcPlEdit, 231, 79
End Sub

Private Sub Form_Resize()
    Playlist1.Move 180, 300, Int((Me.Width - 180) / 15) * 15, Int((Me.Height - 870) / 15) * 15
    picTopR.Move Me.Width - 375, 0
    picClose.Move Me.Width - 180, 30, 135, 150
    picTopM.Move Int((Me.Width / 2) - (picTopM.Width / 2)), 0
    picTopL.Move 0, 0
    picTopF.Move 375, 0, Me.Width - 750, 300
    picBottomL.Move 0, (Me.Height - 570)
    picADD.Move 195, (Me.Height - 465), 360, 300
    picREM.Move 630, (Me.Height - 465), 360, 300
    picSEL.Move 1065, (Me.Height - 465), 360, 300
    picMISC.Move 1500, (Me.Height - 465), 360, 300
    picBottomR.Move Me.Width - 2250, Me.Height - 570, 2250, 570
    picListOpts.Move Me.Width - 675, (Me.Height - 465), 360, 300
    picEvents.Move Me.Width - 2190, Me.Height - 240, 810, 90
    picSlideUp.Move Me.Width - 255, Me.Height - 570, 180, 120
    picSlideDown.Move Me.Width - 255, Me.Height - 450, 180, 120
    vtTime.Move Me.Width - 1320, Me.Height - 225
    vtTotaltime.Move Me.Width - 2145, Me.Height - 420
    If FormFlag = False Then
        picFormResize.Move (Me.Width - 255), (Me.Height - 255)
    End If
    picBorderL.Move 0, 300, 180, Me.Height - 870
    StretchBlt picBorderL.hdc, 0, 0, 12, Int((Me.Height - 870) / 15), picSrcPlEdit.hdc, 0, 42, 12, 29, SRCCOPY
    picBorderL.Refresh
    If Me.Width = 5250 Then
        picBottomM.Move Me.Width - 3375, Me.Height - 570, 1125, 570
        picBottomF.Move 360, 12950, 375, 570
    ElseIf Me.Width = 4125 Then
        picBottomM.Move 840, 12950, 1125, 570
        picBottomF.Move 360, 12950, 375, 570
    ElseIf Me.Width > 4125 And Me.Width < 5250 Then
        picBottomM.Move 840, 12950, 1125, 570
        picBottomF.Move picBottomL.Width, Me.Height - 570, Me.Width - 4125, 570
    Else
        picBottomM.Move Me.Width - 3375, Me.Height - 570, 1125, 570
        picBottomF.Move picBottomL.Width, Me.Height - 570, Me.Width - 5250, 570
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.EXEName, "Visual", "Skin", Playlist1.SkinName
    SaveSetting App.EXEName, "Visual", "SkinDir", Playlist1.SkinDir
    SaveSetting App.EXEName, "Visual", "DefList", Playlist1.Playlist
    SaveSetting App.EXEName, "Mode", "Status", IIf(frmPlayList.Playlist1.ShuffleMode = False, "0", "1") & "," & frmPlayList.Playlist1.Repeat
End Sub

Private Sub picClose_Click()
    frmMain.stndPlayList = 0
    frmMain.imgPlayList.Picture = frmMain.PicPlayList(0).Image
    If Playlist1.IsMovie(Playlist1.CurrentFile) And Playlist1.Status = "Playing" Then Exit Sub:
    Me.Hide
End Sub

Private Sub picClose_Resize()
    SetImage picClose, 0, 0, 9, 10, picTopR, 13, 2
End Sub

Private Sub picEvents_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X <= 135 Then
        Playlist1.EventPrevious
    ElseIf X > 135 And X < 270 Then
        Playlist1.EventPlay
        frmMain.ScrollTimer.Enabled = True
        frmMain.DisplayTimer.Enabled = True
    ElseIf X > 270 And X < 405 Then
        Playlist1.EventPause
        frmMain.ScrollTimer.Enabled = True
        frmMain.DisplayTimer.Enabled = True
    ElseIf X > 405 And X < 540 Then
        Playlist1.EventStop
    ElseIf X > 540 And X < 675 Then
        Playlist1.EventNext
    Else
        Playlist1.AddFile
    End If
End Sub

Private Sub picEvents_Resize()
    SetImage picEvents, 0, 0, 54, 6, picSrcPlEdit, 130, 94
End Sub

Private Sub picFormResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpButtonsOff
    If FormFlag = False Then
        FX = picFormResize.Left + (picFormResize.Width / 2)
        FY = picFormResize.Top + (picFormResize.Height / 2)
        AX = X
        AY = Y
        FormFlag = True
        FormFirst = True
        fTimer = Timer + 0.2
    End If
End Sub

Private Sub picFormResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEvents
    If FormFlag = True Then
        If Timer < fTimer Then Exit Sub:
        fTimer = Timer + 0.2
        Dim posW As Long, posH As Long
        posW = FX + X - IIf(FormFirst = True, X, AX)
        posH = FY + Y - IIf(FormFirst = True, Y, AY)
        If posW < 3870 Then posW = 3870:
        If posH < 1485 Then posH = 1485:
        FX = Int(IIf(FormFirst = True, FX, posW) / 15) * 15
        FY = Int(IIf(FormFirst = True, FY, posH) / 15) * 15
        If FormFirst = False Then
            Dim formW, formH
            picFormResize.Left = FX
            picFormResize.Top = FY
            formW = FX + 255
            formH = FY + 255
             Me.Move Me.Left, Me.Top, formW, formH
        End If
        FormFirst = False
    End If
End Sub

Private Sub picFormResize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormFlag = False
    Me.AutoRedraw = False
    SaveSetting App.EXEName, "Visual", "StartSize", Str(Me.Width) & "," & Str(Me.Height)
End Sub

Private Sub picADD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpButtonsOff
    picAddFILE.Move 210, (Me.Height - 450), 330, 270
    picAddDIR.Move 210, (Me.Height - 720), 330, 270
    picAddURL.Move 210, (Me.Height - 990), 330, 270
    SetImage picAddURL, 0, 0, 22, 18, picSrcPlEdit, 0, 111
    SetImage picAddDIR, 0, 0, 22, 18, picSrcPlEdit, 0, 130
    SetImage picAddFILE, 0, 0, 22, 18, picSrcPlEdit, 23, 149
    picAddFILE.Visible = True
    picADD.Visible = False
End Sub

Private Sub picADD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picAddFILE, 0, 0, 22, 18, picSrcPlEdit, 0, 149
    picAddDIR.Visible = True
    picAddURL.Visible = True
End Sub

Private Sub picAddDIR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picAddDIR, 0, 0, 22, 18, picSrcPlEdit, 23, 130
End Sub

Private Sub picAddDIR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picAddDIR, 0, 0, 22, 18, picSrcPlEdit, 0, 130
    HelpButtonsOff
    Playlist1.AddDir
End Sub

Private Sub picAddFILE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picAddFILE, 0, 0, 22, 18, picSrcPlEdit, 23, 149
End Sub

Private Sub picAddFILE_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picAddFILE, 0, 0, 22, 18, picSrcPlEdit, 0, 149
    HelpButtonsOff
    Playlist1.AddFile
End Sub

Private Sub picAddURL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picAddURL, 0, 0, 22, 18, picSrcPlEdit, 23, 111
End Sub

Private Sub picAddURL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picAddURL, 0, 0, 22, 18, picSrcPlEdit, 0, 111
    HelpButtonsOff
    MsgBox "Under Construction !"
End Sub

Private Sub picBorderL_Click()
    HelpButtonsOff
End Sub

Private Sub picBottomF_Click()
    HelpButtonsOff
End Sub

Private Sub picBottomF_Resize()
    picW = Int(picBottomF.Width / 15)
    StretchBlt picBottomF.hdc, 0, 0, picW, 38, picSrcPlEdit.hdc, 179, 0, 25, 38, SRCCOPY
    picBottomF.Refresh
End Sub

Private Sub picBottomL_Click()
    HelpButtonsOff
End Sub

Private Sub picBottomL_Resize()
    picW = Int(picBottomL.Width / 15)
    StretchBlt picBottomL.hdc, 0, 0, 125, 38, picSrcPlEdit.hdc, 0, 72, 125, 38, SRCCOPY
    picBottomL.Refresh
End Sub

Private Sub picBottomM_Click()
    HelpButtonsOff
End Sub

Private Sub picBottomM_Resize()
    picW = Int(picBottomM.Width / 15)
    StretchBlt picBottomM.hdc, 0, 0, picW, 38, picSrcPlEdit.hdc, 205, 0, 75, 38, SRCCOPY
    picBottomM.Refresh
End Sub

Private Sub picBottomR_Click()
    HelpButtonsOff
End Sub

Private Sub picBottomR_Resize()
    SetImage picBottomR, 0, 0, 150, 38, picSrcPlEdit, 126, 72
End Sub

Private Sub picFormResize_Resize()
    SetImage picFormResize, 0, 0, 17, 17, picSrcPlEdit, 259, 93
End Sub

Private Sub picListOpts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpButtonsOff
    picListOptsLOAD.Move Me.Width - 660, (Me.Height - 450), 330, 270
    picListOptsSAVE.Move Me.Width - 660, (Me.Height - 720), 330, 270
    picListOptsNEW.Move Me.Width - 660, (Me.Height - 990), 330, 270
    SetImage picListOptsNEW, 0, 0, 22, 18, picSrcPlEdit, 204, 111
    SetImage picListOptsSAVE, 0, 0, 22, 18, picSrcPlEdit, 204, 130
    SetImage picListOptsLOAD, 0, 0, 22, 18, picSrcPlEdit, 227, 149
    picListOptsLOAD.Visible = True
    picListOpts.Visible = False
End Sub

Private Sub picListOpts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picListOptsLOAD, 0, 0, 22, 18, picSrcPlEdit, 204, 149
    picListOptsSAVE.Visible = True
    picListOptsNEW.Visible = True
End Sub

Private Sub picListOptsLOAD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picListOptsLOAD, 0, 0, 22, 18, picSrcPlEdit, 227, 149
End Sub

Private Sub picListOptsLOAD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picListOptsLOAD, 0, 0, 22, 18, picSrcPlEdit, 204, 149
    HelpButtonsOff
    Playlist1.OpenList
End Sub

Private Sub picListOptsNEW_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picListOptsNEW, 0, 0, 22, 18, picSrcPlEdit, 227, 111
End Sub

Private Sub picListOptsNEW_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picListOptsNEW, 0, 0, 22, 18, picSrcPlEdit, 204, 111
    HelpButtonsOff
    Playlist1.NewList
End Sub

Private Sub picListOptsSAVE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picListOptsSAVE, 0, 0, 22, 18, picSrcPlEdit, 227, 130
End Sub

Private Sub picListOptsSAVE_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picListOptsSAVE, 0, 0, 22, 18, picSrcPlEdit, 204, 130
    HelpButtonsOff
    Playlist1.SaveList
End Sub

Private Sub picMISC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpButtonsOff
    picMiscOpt.Move 1515, (Me.Height - 450), 330, 270
    picMiscFileInf.Move 1515, (Me.Height - 720), 330, 270
    picMiscSortList.Move 1515, (Me.Height - 990), 330, 270
    SetImage picMiscSortList, 0, 0, 22, 18, picSrcPlEdit, 154, 111
    SetImage picMiscFileInf, 0, 0, 22, 18, picSrcPlEdit, 154, 130
    SetImage picMiscOpt, 0, 0, 22, 18, picSrcPlEdit, 177, 149
    picMiscOpt.Visible = True
    picMISC.Visible = False
End Sub

Private Sub picMISC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picMiscOpt, 0, 0, 22, 18, picSrcPlEdit, 154, 149
    picMiscFileInf.Visible = True
    picMiscSortList.Visible = True
End Sub

Private Sub picMiscFileInf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picMiscFileInf, 0, 0, 22, 18, picSrcPlEdit, 177, 130
End Sub

Private Sub picMiscFileInf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picMiscFileInf, 0, 0, 22, 18, picSrcPlEdit, 154, 130
    HelpButtonsOff
    Playlist1.FileInfo
End Sub

Private Sub picMiscOpt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picMiscOpt, 0, 0, 22, 18, picSrcPlEdit, 177, 149
End Sub

Private Sub picMiscOpt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picMiscOpt, 0, 0, 22, 18, picSrcPlEdit, 154, 149
    HelpButtonsOff
    MsgBox "Under Construction !"
End Sub

Private Sub picMiscSortList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picMiscSortList, 0, 0, 22, 18, picSrcPlEdit, 177, 111
End Sub

Private Sub picMiscSortList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picMiscSortList, 0, 0, 22, 18, picSrcPlEdit, 154, 111
    HelpButtonsOff
    frmSortList.Show
End Sub

Private Sub picREM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpButtonsOff
    picRemFILE.Move 645, (Me.Height - 450), 330, 270
    picRemCROP.Move 645, (Me.Height - 720), 330, 270
    picRemALL.Move 645, (Me.Height - 990), 330, 270
    picRemMISC.Move 645, (Me.Height - 1260), 330, 270
    SetImage picRemALL, 0, 0, 22, 18, picSrcPlEdit, 54, 111
    SetImage picRemCROP, 0, 0, 22, 18, picSrcPlEdit, 54, 130
    SetImage picRemFILE, 0, 0, 22, 18, picSrcPlEdit, 77, 149
    SetImage picRemMISC, 0, 0, 22, 18, picSrcPlEdit, 54, 168
    picRemFILE.Visible = True
    picREM.Visible = False
End Sub

Private Sub picREM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picRemFILE, 0, 0, 22, 18, picSrcPlEdit, 54, 149
    picRemCROP.Visible = True
    picRemALL.Visible = True
    picRemMISC.Visible = True
End Sub

Private Sub picRemALL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picRemALL, 0, 0, 22, 18, picSrcPlEdit, 77, 111
End Sub

Private Sub picRemALL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picRemALL, 0, 0, 22, 18, picSrcPlEdit, 54, 111
    HelpButtonsOff
    Playlist1.DELETE_all
End Sub

Private Sub picRemCROP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picRemCROP, 0, 0, 22, 18, picSrcPlEdit, 77, 130
End Sub

Private Sub picRemCROP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picRemCROP, 0, 0, 22, 18, picSrcPlEdit, 54, 130
    HelpButtonsOff
    Playlist1.DELETE_crop
End Sub

Private Sub picRemFILE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picRemFILE, 0, 0, 22, 18, picSrcPlEdit, 77, 149
End Sub

Private Sub picRemFILE_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picRemFILE, 0, 0, 22, 18, picSrcPlEdit, 54, 149
    HelpButtonsOff
    Playlist1.DELETE_sel
End Sub

Private Sub picRemMISC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picRemMISC, 0, 0, 22, 18, picSrcPlEdit, 77, 168
End Sub

Private Sub picRemMISC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picRemMISC, 0, 0, 22, 18, picSrcPlEdit, 54, 168
    HelpButtonsOff
    Playlist1.DELETE_misc
End Sub

Public Sub HelpButtonsOff()
    picADD.Visible = True
    picAddURL.Visible = False
    picAddDIR.Visible = False
    picAddFILE.Visible = False
    picREM.Visible = True
    picRemMISC.Visible = False
    picRemALL.Visible = False
    picRemCROP.Visible = False
    picRemFILE.Visible = False
    picSEL.Visible = True
    picSelALL.Visible = False
    picSelZERO.Visible = False
    picSelINV.Visible = False
    picMISC.Visible = True
    picMiscOpt.Visible = False
    picMiscFileInf.Visible = False
    picMiscSortList.Visible = False
    picListOpts.Visible = True
    picListOptsLOAD.Visible = False
    picListOptsSAVE.Visible = False
    picListOptsNEW.Visible = False
End Sub

Private Sub picSEL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpButtonsOff
    picSelALL.Move 1080, (Me.Height - 450), 330, 270
    picSelZERO.Move 1080, (Me.Height - 720), 330, 270
    picSelINV.Move 1080, (Me.Height - 990), 330, 270
    SetImage picSelINV, 0, 0, 22, 18, picSrcPlEdit, 104, 111
    SetImage picSelZERO, 0, 0, 22, 18, picSrcPlEdit, 104, 130
    SetImage picSelALL, 0, 0, 22, 18, picSrcPlEdit, 127, 149
    picSelALL.Visible = True
    picSEL.Visible = False
End Sub

Private Sub picSEL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picSelALL, 0, 0, 22, 18, picSrcPlEdit, 104, 149
    picSelZERO.Visible = True
    picSelINV.Visible = True
End Sub

Private Sub picSelALL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picSelALL, 0, 0, 22, 18, picSrcPlEdit, 127, 149
End Sub

Private Sub picSelALL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picSelALL, 0, 0, 22, 18, picSrcPlEdit, 104, 149
    HelpButtonsOff
    Playlist1.SELECT_all
End Sub

Private Sub picSelINV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picSelINV, 0, 0, 22, 18, picSrcPlEdit, 127, 111
End Sub

Private Sub picSelINV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picSelINV, 0, 0, 22, 18, picSrcPlEdit, 104, 111
    HelpButtonsOff
    Playlist1.SELECT_inv
End Sub

Private Sub picSelZERO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picSelZERO, 0, 0, 22, 18, picSrcPlEdit, 127, 130
End Sub

Private Sub picSelZERO_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetImage picSelZERO, 0, 0, 22, 18, picSrcPlEdit, 104, 130
    HelpButtonsOff
    Playlist1.SELECT_none
End Sub

Sub SetImage(Destination As Control, X As Long, Y As Long, W As Long, H As Long, Source As Control, startX As Long, StartY As Long)
    On Error GoTo err_Handle
    Dim DestName As String
    BitBlt Destination.hdc, X, Y, W, H, Source.hdc, startX, StartY, SRCCOPY
    Destination.Refresh
err_Handle:
End Sub

Public Sub SetHoleView()
    frmMain.AutoRedraw = False
    Dim PathNaar As String
    Dim EersteDeel As String
    Dim i As Integer
    EersteDeel = comFunc1.PathBackslash(Playlist1.SkinDir) & Playlist1.SkinName
    EersteDeel = comFunc1.PathBackslash(EersteDeel)
    PathNaar = EersteDeel & "main.bmp"
    If comFunc1.FileExist(PathNaar) Then
        Set frmMain.PicSrcMain.Picture = LoadPicture(PathNaar)
    Else
        Set frmMain.PicSrcMain.Picture = frmMain.stdMain.Picture
    End If
    PathNaar = EersteDeel & "cbuttons.bmp"
    If comFunc1.FileExist(PathNaar) Then
        Set frmMain.PicSrcButton.Picture = LoadPicture(PathNaar)
    Else
        Set frmMain.PicSrcButton.Picture = frmMain.stdButton.Picture
    End If
    PathNaar = EersteDeel & "monoster.bmp"
    If comFunc1.FileExist(PathNaar) Then
        Set frmMain.PicSrcMonoSter.Picture = LoadPicture(PathNaar)
    Else
        Set frmMain.PicSrcMonoSter.Picture = frmMain.stdMonoSter.Picture
    End If
    PathNaar = EersteDeel & "numbers.bmp"
    If comFunc1.FileExist(PathNaar) = False Then
        PathNaar = EersteDeel & "nums_ex.bmp"
    End If
    If comFunc1.FileExist(PathNaar) Then
        Set frmMain.TimeDisplay1.Picture = LoadPicture(PathNaar)
    Else
        Set frmMain.TimeDisplay1.Picture = frmMain.stdTime.Picture
    End If
    PathNaar = EersteDeel & "titlebar.bmp"
    If comFunc1.FileExist(PathNaar) Then
        Set frmMain.PicSrcTitlebar.Picture = LoadPicture(PathNaar)
    Else
        Set frmMain.PicSrcTitlebar.Picture = frmMain.stdTitlebar.Picture
    End If
    PathNaar = EersteDeel & "text.bmp"
    If comFunc1.FileExist(PathNaar) Then
        Set frmMain.ScrollTitle1.Picture = LoadPicture(PathNaar)
        Set frmMain.vtHZ.Picture = LoadPicture(PathNaar)
        Set frmMain.vtKBPS.Picture = LoadPicture(PathNaar)
        Set frmMain.vtTimer.Picture = LoadPicture(PathNaar)
        Set vtTime.Picture = LoadPicture(PathNaar)
        Set vtTotaltime.Picture = LoadPicture(PathNaar)
        Set picText.Picture = LoadPicture(PathNaar)
        picText.Refresh
    Else
        Set frmMain.ScrollTitle1.Picture = frmMain.stdText.Picture
        Set frmMain.vtHZ.Picture = frmMain.stdText.Picture
        Set frmMain.vtKBPS.Picture = frmMain.stdText.Picture
        Set frmMain.vtTimer.Picture = frmMain.stdText.Picture
        Set vtTime.Picture = frmMain.stdText.Picture
        Set vtTotaltime.Picture = frmMain.stdText.Picture
        Set picText.Picture = frmMain.stdText.Picture
        picText.Refresh
    End If
    PathNaar = EersteDeel & "posbar.bmp"
    If comFunc1.FileExist(PathNaar) Then
        Set frmMain.TimeSlider1.Picture = LoadPicture(PathNaar)
    Else
        Set frmMain.TimeSlider1.Picture = frmMain.stdPosBar.Picture
    End If
    PathNaar = EersteDeel & "shufrep.bmp"
    If comFunc1.FileExist(PathNaar) Then
        frmMain.PicSrcShufRep.Picture = LoadPicture(PathNaar)
    Else
        frmMain.PicSrcShufRep.Picture = frmMain.stdShufRep.Picture
    End If
    PathNaar = EersteDeel & "volume.bmp"
    If comFunc1.FileExist(PathNaar) Then
        Set frmMain.Volume1.Picture = LoadPicture(PathNaar)
    Else
        PathNaar = EersteDeel & "volbar.bmp"
        If comFunc1.FileExist(PathNaar) Then
            Set frmMain.Volume1.Picture = LoadPicture(PathNaar)
        Else
            Set frmMain.Volume1.Picture = frmMain.stdVolume.Picture
        End If
    End If
    PathNaar = EersteDeel & "balance.bmp"
    If comFunc1.FileExist(PathNaar) Then
        Set frmMain.Balance1.Picture = LoadPicture(PathNaar)
    Else
        PathNaar = EersteDeel & "volume.bmp"
        If comFunc1.FileExist(PathNaar) Then
            Set frmMain.Balance1.Picture = LoadPicture(PathNaar)
        Else
            PathNaar = EersteDeel & "volbar.bmp"
            If comFunc1.FileExist(PathNaar) Then
                Set frmMain.Balance1.Picture = LoadPicture(PathNaar)
            Else
                Set frmMain.Balance1.Picture = frmMain.stdVolume.Picture
            End If
        End If
    End If
    PathNaar = EersteDeel & "playpaus.bmp"
    If comFunc1.FileExist(PathNaar) Then
        Set frmMain.PicSrcStatus.Picture = LoadPicture(PathNaar)
    Else
        Set frmMain.PicSrcStatus.Picture = frmMain.stdStatus.Picture
    End If
    frmMain.vtKBPS.ShowMessage IIf(Playlist1.CurrentkBps < 100, IIf(Playlist1.CurrentkBps < 10, "  " & Playlist1.CurrentkBps, " " & Playlist1.CurrentkBps), Playlist1.CurrentkBps)
    frmMain.vtHZ.ShowMessage IIf(Playlist1.CurrentHz < 10, " " & Playlist1.CurrentHz, Playlist1.CurrentHz)
    frmMain.AutoRedraw = True
    frmMain.WijzigLayOut
    PathNaar = EersteDeel & "region.txt"
    transSkinPath = PathNaar
    If WinVerOk = True Then
        If comFunc1.FileExist(PathNaar) And frmMain.Minimized = False Then
            If frmMain.BigForm = False Then
                comFunc1.TransRestore frmMain.hWnd, 4125, 1740
                comFunc1.TransRestore frmEqualizer.hWnd, 4125, 1740
                comFunc1.TransWinamp frmMain.hWnd, frmEqualizer.hWnd, PathNaar, False
            Else
                comFunc1.TransRestore frmMain.hWnd, 8250, 3480
                comFunc1.TransRestore frmEqualizer.hWnd, 8250, 3480
                comFunc1.TransWinamp frmMain.hWnd, frmEqualizer.hWnd, PathNaar, True
            End If
        Else
            If frmMain.BigForm = False Then
                comFunc1.TransRestore frmMain.hWnd, 4125, 1740
                comFunc1.TransRestore frmEqualizer.hWnd, 4125, 1740
            Else
                comFunc1.TransRestore frmMain.hWnd, 8250, 3480
                comFunc1.TransRestore frmEqualizer.hWnd, 8250, 3480
            End If
        End If
        If frmMain.stndEQ = 0 Then frmEqualizer.Hide:
    End If
    PathNaar = EersteDeel & "eqmain.bmp"
    If comFunc1.FileExist(PathNaar) Then
        Set frmEqualizer.PicSrcEQMain.Picture = LoadPicture(PathNaar)
    Else
        Set frmEqualizer.PicSrcEQMain.Picture = frmEqualizer.stdEQMain.Picture
    End If
    frmEqualizer.EQLayOut
    PathNaar = EersteDeel & "pledit.bmp"
    If comFunc1.FileExist(PathNaar) Then
        picSrcPlEdit.Picture = LoadPicture(PathNaar)
        Playlist1.WinAMP = True
        Set Playlist1.Picture = LoadPicture(PathNaar)
        ChangePlayList
        PathNaar = EersteDeel & "pledit.txt"
        If comFunc1.FileExist(PathNaar) Then
            Dim a As Long
            a = comFunc1.Hex_Long(comFunc1.GetKeyValue(PathNaar, "Text", "NormalBG"))
            DoEvents
            Playlist1.BackColor = a
            SearchList.BackColor = a
            a = comFunc1.Hex_Long(comFunc1.GetKeyValue(PathNaar, "Text", "Normal"))
            Playlist1.ForeColor = a
            SearchList.ForeColor = a
            a = comFunc1.Hex_Long(comFunc1.GetKeyValue(PathNaar, "Text", "Current"))
            Playlist1.ForeColorSel = a
            With mediaFont
                .Name = comFunc1.GetKeyValue(PathNaar, "Text", "Font")
                .Size = 7
            End With
            Set Playlist1.Font = mediaFont
            SearchList.Font = comFunc1.GetKeyValue(PathNaar, "Text", "Font")
            SearchList.FontSize = 7
        Else
            Playlist1.BackColor = GetPixel(picText.hdc, CLng(149), CLng(2))
            SearchList.BackColor = Playlist1.BackColor
            Playlist1.ForeColor = GetPixel(picText.hdc, CLng(152), CLng(8))
            SearchList.ForeColor = Playlist1.ForeColor
            With mediaFont
                .Name = "Tahoma"
                .Size = 7
            End With
            Set Playlist1.Font = mediaFont
            SearchList.Font = "Tahoma"
            SearchList.FontSize = 7
        End If
    Else
        picSrcPlEdit.Picture = PlayListRes.Picture
        Playlist1.WinAMP = True
        Set Playlist1.Picture = PlayListRes.Picture
        Playlist1.BackColor = 65793
        SearchList.BackColor = 65793
        Playlist1.ForeColor = 1565394
        SearchList.ForeColor = 1565394
        Playlist1.ForeColorSel = 65331
        SearchList.Font = "Tahoma"
        SearchList.FontSize = 7
    End If
    ChangePlayList
    Playlist1.Move 180, 300, Me.Width - 180, Me.Height - 870
    SearchList.Move 180, 300, Me.Width - 180, Me.Height - 870
End Sub

Private Sub picSlideDown_Click()
    Playlist1.ScrollDown
End Sub

Private Sub picSlideDown_Resize()
    SetImage picSlideDown, 0, 0, 12, 8, picBottomR, 133, 8
End Sub

Private Sub picSlideUp_Click()
    Playlist1.ScrollUp
End Sub

Private Sub picSlideUp_Resize()
    SetImage picSlideUp, 0, 0, 12, 8, picBottomR, 133, 0
End Sub

Private Sub picTopF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpButtonsOff
    frmMain.PlayListHangsT = False
    frmMain.PlayListHangsL = False
    If MoveForm = False Then
        FX = picTopF.Left + picTopF.Width / 2
        FY = picTopF.Top + picTopF.Height / 2
        AX = X
        AY = Y
        MoveForm = True
        MoveFirst = True
    End If
End Sub

Private Sub picTopF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MoveForm = True Then
        Dim posW, posH
        posW = Int((FX + X - IIf(MoveFirst = True, X, AX)) / 15) * 15
        posH = Int((FY + Y - IIf(MoveFirst = True, Y, AY)) / 15) * 15
        FX = IIf(MoveFirst = True, FX, posW)
        FY = IIf(MoveFirst = True, FY, posH)
        If MoveFirst = False Then Me.Move FX, FY:
        MoveFirst = False
    End If
End Sub

Private Sub picTopF_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm = False
End Sub

Private Sub picTopF_Resize()
    Dim picW
    picW = Int(picTopF.Width / 15)
    StretchBlt picTopF.hdc, 0, 0, picW, 20, picSrcPlEdit.hdc, 127, 0, 25, 20, SRCCOPY
    picTopF.Refresh
End Sub

Private Sub picTopL_Click()
    HelpButtonsOff
End Sub

Private Sub picTopL_Resize()
    SetImage picTopL, 0, 0, 25, 20, picSrcPlEdit, 0, 0
End Sub

Private Sub picTopM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HelpButtonsOff
    frmMain.PlayListHangsT = False
    frmMain.PlayListHangsL = False
    If MoveForm = False Then
        FX = picTopM.Left + picTopM.Width / 2
        FY = picTopM.Top + picTopM.Height / 2
        AX = X
        AY = Y
        MoveForm = True
        MoveFirst = True
    End If
End Sub

Private Sub picTopM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MoveForm = True Then
        Dim posW, posH
        posW = Int((FX + X - IIf(MoveFirst = True, X, AX)) / 15) * 15
        posH = Int((FY + Y - IIf(MoveFirst = True, Y, AY)) / 15) * 15
        FX = IIf(MoveFirst = True, FX, posW)
        FY = IIf(MoveFirst = True, FY, posH)
        If MoveFirst = False Then Me.Move FX, FY:
        MoveFirst = False
    End If
End Sub

Private Sub picTopM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm = False
End Sub

Private Sub picTopM_Resize()
    SetImage picTopM, 0, 0, 100, 20, picSrcPlEdit, 26, 0
End Sub

Private Sub picTopR_Click()
    HelpButtonsOff
End Sub

Private Sub picTopR_Resize()
    SetImage picTopR, 0, 0, 25, 20, picSrcPlEdit, 153, 0
End Sub

Public Sub StandardSize()
    Dim a
    a = Trim(GetSetting(App.EXEName, "Visual", "StartSize", "4125,1740"))
    If a <> "" Then
        For i = 1 To Len(a)
            If Mid(a, i, 1) = "," Then Exit For:
            If Mid(a, i, 1) = ";" Then Exit For:
        Next
        Me.Width = IIf(IsNumeric(Left$(a, i - 1)) = True, Left(a, i - 1), 4125)
        Me.Height = IIf(IsNumeric(Mid$(a, i + 1)) = True, Right(a, Len(a) - i), 1740)
        If Me.Width < 4125 Then Me.Width = 4125:
        If Me.Height < 1740 Then Me.Height = 1740:
    End If
    picFormResize.Left = Me.Width - 255
    picFormResize.Top = Me.Height - 255
    Call ChangePlayList
End Sub

