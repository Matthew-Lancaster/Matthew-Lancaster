VERSION 5.00
Begin VB.Form frmOptions 
   ClientHeight    =   3015
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   2550
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPlayByStart 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      Picture         =   "frmOptions.frx":0000
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   14
      Top             =   2520
      Width           =   195
   End
   Begin VB.PictureBox picShuffle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      Picture         =   "frmOptions.frx":0222
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   2130
      Width           =   195
   End
   Begin VB.PictureBox picRepeat 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      Picture         =   "frmOptions.frx":0444
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   12
      Top             =   1890
      Width           =   195
   End
   Begin VB.PictureBox picDouble 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      Picture         =   "frmOptions.frx":0666
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   1515
      Width           =   195
   End
   Begin VB.PictureBox picAlwaysOnTop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      Picture         =   "frmOptions.frx":0888
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   1275
      Width           =   195
   End
   Begin VB.PictureBox picTimeElapsed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      Picture         =   "frmOptions.frx":0AAA
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   660
      Width           =   195
   End
   Begin VB.PictureBox picTimeRemaining 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      Picture         =   "frmOptions.frx":0CCC
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   900
      Width           =   195
   End
   Begin VB.Label lblRegistration 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Registration"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2745
      Width           =   2175
   End
   Begin VB.Label lblStartByStart 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Play At Startup"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2505
      Width           =   2175
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   2500
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   2500
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label LblSetSkinDir 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Set Skin Dir  ..."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   270
      Width           =   1575
   End
   Begin VB.Label LblShuffle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Shuffle                         S"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2115
      Width           =   2175
   End
   Begin VB.Label LblRepeat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Repeat                        R"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1875
      Width           =   2175
   End
   Begin VB.Label LblDouble 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Double Size             Ctrl+D"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1500
      Width           =   2175
   End
   Begin VB.Label LblAlwaysOnTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Always On Top        Ctrl+A"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1260
      Width           =   2175
   End
   Begin VB.Label LblTimeRemaining 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Time re&maining (Ctrl+T toggles)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   885
      Width           =   2175
   End
   Begin VB.Label LblTimeElapsed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Time &elapsed   (Ctrl+T toggles)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   645
      Width           =   2175
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   2500
      Y1              =   1155
      Y2              =   1155
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   2500
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   2500
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   2500
      Y1              =   1785
      Y2              =   1785
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   2500
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   2500
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label LblSkins 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Skins ..."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   30
      Width           =   1575
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Message As String
Dim ScrollOn As Boolean

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = frmMain.Top + IIf(frmMain.BigForm = False, 900, 1800)
    Me.Left = frmMain.Left + IIf(frmMain.BigForm = False, 450, 900)
    If frmMain.stndOnTop = True Then
        SetWindowPos frmMain.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
        SetWindowPos frmPlayList.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    End If
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    If frmMain.intTimeMode <> 0 Then picTimeElapsed.Picture = Nothing:
    If frmMain.intTimeMode <> 1 Then picTimeRemaining.Picture = Nothing:
    If frmMain.stndOnTop = False Then picAlwaysOnTop.Picture = Nothing:
    If frmMain.BigForm = False Then picDouble.Picture = Nothing:
    If frmPlayList.Playlist1.Repeat = 0 Then picRepeat.Picture = Nothing:
    If frmPlayList.Playlist1.Repeat = 1 Then LblRepeat.Caption = "&Repeat 1                     R":
    If frmPlayList.Playlist1.Repeat = 2 Then LblRepeat.Caption = "&Repeat All                   R":
    If frmPlayList.Playlist1.ShuffleMode = False Then picShuffle.Picture = Nothing:
    If frmMain.PlayByStart = False Then picPlayByStart.Picture = Nothing:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frmMain.stndOnTop = True Then
        SetWindowPos frmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
        If frmMain.QuitmciPlayer = False Then
            SetWindowPos frmPlayList.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
        End If
    End If
    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    Set frmMain.imgOptions.Picture = Nothing
    frmMain.stndOptions = False
End Sub

Private Sub LblAlwaysOnTop_Click()
    If frmMain.stndOnTop = False Then
        frmMain.imgAllOnTop.Picture = frmMain.PicTitleOpt(1).Image
        frmMain.stndOnTop = True
        SetWindowPos frmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
        SetWindowPos frmPlayList.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    Else
        Set frmMain.imgAllOnTop.Picture = Nothing
        frmMain.stndOnTop = False
        SetWindowPos frmMain.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
        SetWindowPos frmPlayList.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
    End If
    Unload Me
End Sub

Private Sub lblBass_Click()
    Unload Me
End Sub

Private Sub LblDouble_Click()
    If frmMain.BigForm = True Then
        frmMain.BigForm = False
        Set frmMain.imgDouble.Picture = Nothing
    Else
        frmMain.BigForm = True
        frmMain.imgDouble.Picture = frmMain.PicTitleOpt(3).Image
    End If
    frmMain.WijzigFormaat
    If frmMain.stndEQ = 1 Then
        frmEqualizer.Top = frmMain.Top + frmMain.Height
        frmEqualizer.Left = frmMain.Left
        If frmMain.stndPlayList = 1 Then
            frmPlayList.Top = frmEqualizer.Top + frmEqualizer.Height
            frmPlayList.Left = frmMain.Left
        End If
    Else
        If frmMain.stndPlayList = 1 Then
            frmPlayList.Top = frmMain.Top + frmMain.Height
            frmPlayList.Left = frmMain.Left
        End If
    End If
    Unload Me
End Sub

Private Sub lblRegistration_Click()
    frmPlayList.comFunc1.Registration
    Unload Me
End Sub

Private Sub LblRepeat_Click()
    If frmPlayList.Playlist1.Repeat = 0 Then
        frmMain.imgRepeat.Picture = frmMain.picRepeat(2).Image
        frmPlayList.Playlist1.Repeat = 1
        Message = "Repeat 1"
    ElseIf frmPlayList.Playlist1.Repeat = 1 Then
        frmMain.imgRepeat.Picture = frmMain.picRepeat(2).Image
        frmPlayList.Playlist1.Repeat = 2
        Message = "Repeat All"
    Else
        frmMain.imgRepeat.Picture = frmMain.picRepeat(0).Image
        frmPlayList.Playlist1.Repeat = 0
        Message = "Repeat Off"
    End If
    Me.Hide
    Dim t#
    If frmMain.ScrollTimer.Enabled = True Then ScrollOn = True: frmMain.ScrollTimer.Enabled = False:
    frmMain.ScrollTitle1.ShowMessage Message
    t# = Timer + 1: Do Until Timer >= t#: DoEvents: Loop:
    If ScrollOn = True Then frmMain.ScrollTimer.Enabled = True:
    Unload Me
End Sub

Private Sub LblSetSkinDir_Click()
    Me.Hide
    Dim a
    frmPlayList.Playlist1.SKIN_Dir
    a = frmPlayList.Playlist1.SkinDir
    If a <> "" Then frmPlayList.SkinPath.Text = a:
    Unload Me
End Sub

Private Sub LblShuffle_Click()
    frmMain.imgShuffle_MouseDown 1, 0, 15, 15
    If frmPlayList.Playlist1.ShuffleMode = False Then
        frmMain.imgShuffle.Picture = frmMain.picShuffle(2).Image
        frmPlayList.Playlist1.ShuffleMode = True
    Else
        frmMain.imgShuffle.Picture = frmMain.picShuffle(0).Image
        frmPlayList.Playlist1.ShuffleMode = False
    End If
    Me.Hide
    Dim t#
    t# = Timer + 1
    Do Until Timer > t#: DoEvents: Loop:
    frmMain.ScrollTimer.Enabled = True
    Unload Me
End Sub

Private Sub LblSkins_Click()
    If UCase(Right(frmPlayList.Playlist1.CurrentFile, 3)) <> "AVI" Then
        If frmMain.stndPlayList = 0 Then
            frmPlayList.Show
            frmMain.stndPlayList = 1
            frmMain.imgPlayList.Picture = frmMain.PicPlayList(2).Image
        End If
        frmPlayList.Playlist1.SKIN_Select
    End If
    Unload Me
End Sub

Private Sub lblStartByStart_Click()
    If frmMain.PlayByStart = True Then
        frmMain.PlayByStart = False
    Else
        frmMain.PlayByStart = True
    End If
    Unload Me
End Sub

Private Sub LblTimeElapsed_Click()
    frmMain.intTimeMode = 0
    Unload Me
End Sub

Private Sub LblTimeRemaining_Click()
    frmMain.intTimeMode = 1
    Unload Me
End Sub

Private Sub lblTreble_Click()
    Unload Me
End Sub

Private Sub picAlwaysOnTop_Click()
    LblAlwaysOnTop_Click
End Sub

Private Sub picDouble_Click()
    LblDouble_Click
End Sub

Private Sub picRepeat_Click()
    LblRepeat_Click
End Sub

Private Sub picShuffle_Click()
    LblShuffle_Click
End Sub

Private Sub picTimeElapsed_Click()
    LblTimeElapsed_Click
End Sub

Private Sub picTimeRemaining_Click()
    LblTimeRemaining_Click
End Sub

