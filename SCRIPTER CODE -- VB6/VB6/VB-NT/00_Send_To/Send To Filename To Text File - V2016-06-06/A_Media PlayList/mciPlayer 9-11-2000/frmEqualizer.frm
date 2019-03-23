VERSION 5.00
Begin VB.Form frmEqualizer 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicSetEQ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   930
      Index           =   10
      Left            =   310
      ScaleHeight     =   930
      ScaleWidth      =   210
      TabIndex        =   13
      ToolTipText     =   "Volume Bar"
      Top             =   570
      Width           =   210
   End
   Begin VB.PictureBox PicSetEQ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   930
      Index           =   0
      Left            =   1170
      ScaleHeight     =   930
      ScaleWidth      =   210
      TabIndex        =   12
      ToolTipText     =   "Volume Bar"
      Top             =   570
      Width           =   210
   End
   Begin VB.PictureBox PicSetEQ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   930
      Index           =   1
      Left            =   1440
      ScaleHeight     =   930
      ScaleWidth      =   210
      TabIndex        =   11
      ToolTipText     =   "Volume Bar"
      Top             =   570
      Width           =   210
   End
   Begin VB.PictureBox PicSetEQ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   930
      Index           =   2
      Left            =   1710
      ScaleHeight     =   930
      ScaleWidth      =   210
      TabIndex        =   10
      ToolTipText     =   "Volume Bar"
      Top             =   570
      Width           =   210
   End
   Begin VB.PictureBox PicSetEQ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   930
      Index           =   3
      Left            =   1980
      ScaleHeight     =   930
      ScaleWidth      =   210
      TabIndex        =   9
      ToolTipText     =   "Volume Bar"
      Top             =   570
      Width           =   210
   End
   Begin VB.PictureBox PicSetEQ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   930
      Index           =   4
      Left            =   2250
      ScaleHeight     =   930
      ScaleWidth      =   210
      TabIndex        =   8
      ToolTipText     =   "Volume Bar"
      Top             =   570
      Width           =   210
   End
   Begin VB.PictureBox PicSetEQ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   930
      Index           =   5
      Left            =   2520
      ScaleHeight     =   930
      ScaleWidth      =   210
      TabIndex        =   7
      ToolTipText     =   "Volume Bar"
      Top             =   570
      Width           =   210
   End
   Begin VB.PictureBox PicSetEQ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   930
      Index           =   6
      Left            =   2790
      ScaleHeight     =   930
      ScaleWidth      =   210
      TabIndex        =   6
      ToolTipText     =   "Volume Bar"
      Top             =   570
      Width           =   210
   End
   Begin VB.PictureBox PicSetEQ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   930
      Index           =   7
      Left            =   3060
      ScaleHeight     =   930
      ScaleWidth      =   210
      TabIndex        =   5
      ToolTipText     =   "Volume Bar"
      Top             =   570
      Width           =   210
   End
   Begin VB.PictureBox PicSetEQ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   930
      Index           =   8
      Left            =   3330
      ScaleHeight     =   930
      ScaleWidth      =   210
      TabIndex        =   4
      ToolTipText     =   "Volume Bar"
      Top             =   570
      Width           =   210
   End
   Begin VB.PictureBox PicSetEQ 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   930
      Index           =   9
      Left            =   3600
      ScaleHeight     =   930
      ScaleWidth      =   210
      TabIndex        =   3
      ToolTipText     =   "Volume Bar"
      Top             =   570
      Width           =   210
   End
   Begin VB.PictureBox stdEQMain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4725
      Left            =   4440
      Picture         =   "frmEqualizer.frx":0000
      ScaleHeight     =   4725
      ScaleWidth      =   4125
      TabIndex        =   2
      Top             =   3720
      Width           =   4125
   End
   Begin VB.PictureBox PicSrcEQMain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4725
      Left            =   240
      ScaleHeight     =   4725
      ScaleWidth      =   4125
      TabIndex        =   1
      Top             =   3720
      Width           =   4125
   End
   Begin VB.PictureBox PicEQMain 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   4440
      ScaleHeight     =   1740
      ScaleWidth      =   4125
      TabIndex        =   0
      Top             =   3720
      Width           =   4125
   End
   Begin VB.Image ImgTitleBarEQ 
      Height          =   210
      Left            =   0
      Top             =   0
      Width           =   4125
   End
   Begin VB.Image ImgEQMain 
      Height          =   1740
      Left            =   0
      Top             =   0
      Width           =   4125
   End
End
Attribute VB_Name = "frmEqualizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer
Dim MoveForm As Boolean, MoveFirst As Boolean
Dim FX, FY, AX, AY

Sub SetImage(Destination As Control, X As Long, Y As Long, W As Long, H As Long, Source As Control, startX As Long, StartY As Long)
    On Error GoTo err_Handle
    Dim DestName As String
    BitBlt Destination.hdc, X, Y, W, H, Source.hdc, startX, StartY, SRCCOPY
    Destination.Refresh
    DestName = "img" & Mid$(Destination.Name, 4, Len(Destination.Name))
    Me(DestName).Picture = Destination.Image
    Me(DestName).Refresh
err_Handle:
End Sub

Private Sub Form_Load()
    PicSrcEQMain.Picture = stdEQMain.Picture
    EQFormat
    EQLayOut
End Sub

Public Sub EQLayOut()
    SetImage PicEQMain, 0, 0, 275, 116, PicSrcEQMain, 0, 0
    For i = 0 To 10
        EQSettings i, IIf(i = 0, 1, i * 2)
    Next
End Sub

Private Sub EQSettings(controlID As Integer, value As Integer)
    Select Case value
        Case Is = 1
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 13, 164, SRCCOPY
        Case Is = 2
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 28, 164, SRCCOPY
        Case Is = 3
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 43, 164, SRCCOPY
        Case Is = 4
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 58, 164, SRCCOPY
        Case Is = 5
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 73, 164, SRCCOPY
        Case Is = 6
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 88, 164, SRCCOPY
        Case Is = 7
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 103, 164, SRCCOPY
        Case Is = 8
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 118, 164, SRCCOPY
        Case Is = 9
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 133, 164, SRCCOPY
        Case Is = 10
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 148, 164, SRCCOPY
        Case Is = 11
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 163, 164, SRCCOPY
        Case Is = 12
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 178, 164, SRCCOPY
        Case Is = 13
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 193, 164, SRCCOPY
        Case Is = 14
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 208, 164, SRCCOPY
        Case Is = 15
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 13, 229, SRCCOPY
        Case Is = 16
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 28, 229, SRCCOPY
        Case Is = 17
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 43, 229, SRCCOPY
        Case Is = 18
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 58, 229, SRCCOPY
        Case Is = 19
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 73, 229, SRCCOPY
        Case Is = 20
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 88, 229, SRCCOPY
        Case Is = 21
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 103, 229, SRCCOPY
        Case Is = 22
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 118, 229, SRCCOPY
        Case Is = 23
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 133, 229, SRCCOPY
        Case Is = 24
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 148, 229, SRCCOPY
        Case Is = 25
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 163, 229, SRCCOPY
        Case Is = 26
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 178, 229, SRCCOPY
        Case Is = 27
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 193, 229, SRCCOPY
        Case Is = 28
            BitBlt PicSetEQ(controlID).hdc, 0, 0, 14, 63, PicSrcEQMain.hdc, 208, 229, SRCCOPY
    End Select
    PicSetEQ(controlID).Refresh
    If frmMain.BigForm = True Then
        StretchBlt PicSetEQ(controlID).hdc, 0, 0, 28, 126, PicSetEQ(controlID).hdc, 0, 0, 14, 63, SRCCOPY
        PicSetEQ(controlID).Refresh
    End If
End Sub

Public Sub EQFormat()
    Me.Visible = False
    On Error GoTo err_Handle
    Me.ScaleMode = 1
    Dim i As Integer, ci As Integer
    If frmMain.BigForm = False Then
        Me.Width = 4125
        Me.Height = 1740
        ImgTitleBarEQ.Move 0, 0, 4125, 210
        ImgEQMain.Stretch = False
        ImgEQMain.Move 0, 0, 4125, 1740
        For i = 1170 To 3600 Step 270
            PicSetEQ(ci).Move i, 570, 210, 930
            ci = ci + 1
        Next
        PicSetEQ(10).Move 310, 570, 210, 930
    Else
        Me.Width = 8250
        Me.Height = 3480
        ImgTitleBarEQ.Move 0, 0, 8250, 420
        ImgEQMain.Stretch = True
        ImgEQMain.Move 0, 0, 8250, 3480
        For i = 2340 To 7200 Step 540
            PicSetEQ(ci).Move i, 1140, 420, 1860
            ci = ci + 1
        Next
        PicSetEQ(10).Move 620, 1140, 420, 1860
    End If
    For i = 0 To 10
        EQSettings i, IIf(i = 0, 1, i * 2)
    Next
err_Handle:
    Me.Visible = True
    If frmMain.stndEQ = 0 Then frmEqualizer.Hide:
End Sub

Private Sub ImgTitleBarEQ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMain.EqualizerHangsL = False
    frmMain.EqualizerHangsT = False
    If MoveForm = False Then
        FX = ImgTitleBarEQ.Left + ImgTitleBarEQ.Width / 2
        FY = ImgTitleBarEQ.Top + ImgTitleBarEQ.Height / 2
        AX = X
        AY = Y
        MoveForm = True
        MoveFirst = True
    End If
End Sub

Private Sub ImgTitleBarEQ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MoveForm = True Then
        Dim posW, posH
        posW = Int((FX + X - IIf(MoveFirst = True, X, AX)) / 15) * 15
        posH = Int((FY + Y - IIf(MoveFirst = True, Y, AY)) / 15) * 15
        FX = IIf(MoveFirst = True, FX, posW)
        FY = IIf(MoveFirst = True, FY, posH)
        If MoveFirst = False Then Me.Move FX, FY:
        MoveFirst = False
        Dim posLBC As Integer, posTBC As Integer
        posLBC = Me.Left + Me.Width
        posTBC = Me.Top + Me.Height
        If frmMain.stndPlayList = 1 Then
            If frmMain.PlayListHangsL = False Then
                If posLBC < frmPlayList.Left + 120 And posLBC > frmPlayList.Left - 120 Then
                    If Me.Top < frmPlayList.Top + 120 And Me.Top > frmPlayList.Top - 120 Then
                        frmPlayList.Move (Me.Left + Me.Width), Me.Top
                        frmMain.PlayListHangsL = True
                        frmMain.PlayListHangsT = False
                    End If
                Else
                    frmMain.PlayListHangsL = False
                End If
            Else
                frmPlayList.Move (Me.Left + Me.Width), Me.Top
            End If
            If frmMain.PlayListHangsT = False Then
                If posTBC < frmPlayList.Top + 120 And posTBC > frmPlayList.Top - 120 Then
                    If Me.Left < frmPlayList.Left + 120 And Me.Left > frmPlayList.Left - 120 Then
                        frmPlayList.Move Me.Left, (Me.Top + Me.Height)
                        frmMain.PlayListHangsT = True
                        frmMain.PlayListHangsL = False
                    End If
                Else
                    frmMain.PlayListHangsT = False
                End If
            Else
                frmPlayList.Move Me.Left, (Me.Top + Me.Height)
            End If
        Else
            frmMain.PlayListHangsL = False
            frmMain.PlayListHangsT = False
        End If
    End If
End Sub

Private Sub ImgTitleBarEQ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm = False
End Sub

