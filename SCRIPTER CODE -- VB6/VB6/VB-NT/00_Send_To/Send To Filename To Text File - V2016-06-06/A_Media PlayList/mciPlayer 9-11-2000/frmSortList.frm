VERSION 5.00
Begin VB.Form frmSortList 
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   2550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSort 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   0
      Picture         =   "frmSortList.frx":0000
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   285
      Width           =   195
   End
   Begin VB.PictureBox picSort 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   0
      Picture         =   "frmSortList.frx":0222
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   45
      Width           =   195
   End
   Begin VB.PictureBox picSort 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   3
      Left            =   0
      Picture         =   "frmSortList.frx":0444
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   900
      Width           =   195
   End
   Begin VB.PictureBox picSort 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   0
      Picture         =   "frmSortList.frx":0666
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   660
      Width           =   195
   End
   Begin VB.PictureBox picSort 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   4
      Left            =   0
      Picture         =   "frmSortList.frx":0888
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1275
      Width           =   195
   End
   Begin VB.PictureBox picSort 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   5
      Left            =   0
      Picture         =   "frmSortList.frx":0AAA
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1515
      Width           =   195
   End
   Begin VB.PictureBox picSort 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   6
      Left            =   0
      Picture         =   "frmSortList.frx":0CCC
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1890
      Width           =   195
   End
   Begin VB.PictureBox picSort 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   7
      Left            =   0
      Picture         =   "frmSortList.frx":0EEE
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2130
      Width           =   195
   End
   Begin VB.Label lblSortNameA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By Name             (Asc)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   30
      Width           =   2100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   2500
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   2500
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   2500
      Y1              =   1785
      Y2              =   1785
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   2500
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   2500
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   2500
      Y1              =   1155
      Y2              =   1155
   End
   Begin VB.Label lblSortPathA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By Path               (Asc)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   645
      Width           =   2100
   End
   Begin VB.Label lblSortPathD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By Path              (Desc)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   885
      Width           =   2100
   End
   Begin VB.Label lblSortTimeA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By Time               (Asc)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1260
      Width           =   2100
   End
   Begin VB.Label lblSortTimeD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By Time              (Desc)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1500
      Width           =   2100
   End
   Begin VB.Label lblSortTypeA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By Type              (Asc)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1875
      Width           =   2100
   End
   Begin VB.Label lblSortTypeD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By Type             (Desc)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2115
      Width           =   2100
   End
   Begin VB.Label lblSortNameD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sort By Name            (Desc)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   270
      Width           =   2100
   End
End
Attribute VB_Name = "frmSortList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i
    Dim variabele As Control
    For Each variabele In Controls
        If TypeOf variabele Is PictureBox Then
            If variabele.Index <> intLastSort Then
                Set variabele.Picture = Nothing
            End If
        End If
    Next
    Me.Top = frmPlayList.Top + frmPlayList.picMiscSortList.Top - Me.Height - 120
    Me.Left = frmPlayList.Left + frmPlayList.picMiscSortList.Left + frmPlayList.picMiscSortList.Width + 120
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

Private Sub lblSortNameA_Click()
    frmPlayList.Playlist1.OrderBy obNameAsc
    intLastSort = 0
    Unload Me
End Sub

Private Sub lblSortNameD_Click()
    frmPlayList.Playlist1.OrderBy obNameDesc
    intLastSort = 1
    Unload Me
End Sub

Private Sub lblSortPathA_Click()
    frmPlayList.Playlist1.OrderBy obPathAsc
    intLastSort = 2
    Unload Me
End Sub

Private Sub lblSortPathD_Click()
    frmPlayList.Playlist1.OrderBy obPathDesc
    intLastSort = 3
    Unload Me
End Sub

Private Sub lblSortTimeA_Click()
    frmPlayList.Playlist1.OrderBy obTimeAsc
    intLastSort = 4
    Unload Me
End Sub

Private Sub lblSortTimeD_Click()
    frmPlayList.Playlist1.OrderBy obTimeDesc
    intLastSort = 5
    Unload Me
End Sub

Private Sub lblSortTypeA_Click()
    frmPlayList.Playlist1.OrderBy obTypeAsc
    intLastSort = 6
    Unload Me
End Sub

Private Sub lblSortTypeD_Click()
    frmPlayList.Playlist1.OrderBy obTypeDesc
    intLastSort = 7
    Unload Me
End Sub

Private Sub picSort_Click(Index As Integer)
    Select Case Index
        Case Is = 0: lblSortNameA_Click:
        Case Is = 1: lblSortNameD_Click:
        Case Is = 2: lblSortPathA_Click:
        Case Is = 3: lblSortPathD_Click:
        Case Is = 4: lblSortTimeA_Click:
        Case Is = 5: lblSortTimeD_Click:
        Case Is = 6: lblSortTypeA_Click:
        Case Is = 7: lblSortTypeD_Click:
    End Select
End Sub

