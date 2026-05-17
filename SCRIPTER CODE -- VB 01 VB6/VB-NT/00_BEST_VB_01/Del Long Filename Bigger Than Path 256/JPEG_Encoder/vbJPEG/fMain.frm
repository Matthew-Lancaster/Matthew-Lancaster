VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "JPEG Encoder Demo"
   ClientHeight    =   3435
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Left            =   720
      ScaleHeight     =   2115
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         Left            =   1800
         TabIndex        =   2
         Top             =   1800
         Width           =   855
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   855
         Left            =   2760
         TabIndex        =   1
         Top             =   840
         Width           =   195
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"fMain.frx":0000
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save As ..."
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "Image"
      Begin VB.Menu mnuRotateMaster 
         Caption         =   "Rotate"
         Begin VB.Menu mnuRotate 
            Caption         =   "Clockwise"
            Index           =   0
         End
         Begin VB.Menu mnuRotate 
            Caption         =   "Counter Clockwise"
            Index           =   1
         End
         Begin VB.Menu mnuRotate 
            Caption         =   "180 Degrees"
            Index           =   2
         End
      End
      Begin VB.Menu mnuMirrorMaster 
         Caption         =   "Mirror"
         Begin VB.Menu mnuMirror 
            Caption         =   "Horizontal"
            Index           =   0
         End
         Begin VB.Menu mnuMirror 
            Caption         =   "Vertical"
            Index           =   1
         End
      End
      Begin VB.Menu mnuAutosize 
         Caption         =   "Autosize Window"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Option Explicit
Option Base 0

Public m_Image                    As New cImage

'Reserved space around picturebox
Private Const PictureBoxLeft      As Long = 0
Private Const PictureBoxTop       As Long = 0
Private Const PictureBoxRight     As Long = 0
Private Const PictureBoxBottom    As Long = 240   '240 because form has a menu

'Mouse button for grab and drag
Private Const ButtonDrag          As Integer = 1  'Left Mouse
Private PaintLeft           As Long
Private PaintTop            As Long

Private Const TwipsPerPixel       As Long = 15    'Is this ever not true?




Private Sub Form_Load()
'PicShowReSize.Show
Call mnuOpen_Click
        
Call mnuSave_Click(0)


End Sub

Private Sub Form_Resize()
    Dim newWidth As Long
    Dim newHeight As Long

    newWidth = Me.Width - PictureBoxLeft - PictureBoxRight - 120
    newHeight = Me.Height - PictureBoxTop - PictureBoxBottom - 420
    If newWidth > 0 And newHeight > 0 Then Picture1.Move PictureBoxLeft, PictureBoxTop, newWidth, newHeight
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set m_Image = Nothing
    End
End Sub




'================================================================================
'                        LINKING PICTURE TO SCROLLBARS
'================================================================================
Private Sub AdjustScrollBars(TheImage As cImage)
    Dim X As Long 'Set Max/Min/Visible properties of HScroll1 and VScroll1
    Dim Y As Long '    for TheImage in Picture1


    If ObjPtr(TheImage) = 0 Then 'Remove HScroll and VScroll
        HScroll1.Min = 0
        HScroll1.Max = 0
        HScroll1.Visible = False
        VScroll1.Min = 0
        VScroll1.Max = 0
        VScroll1.Visible = False
    Else
        If Picture1.Width >= VScroll1.Width + 4 * TwipsPerPixel And Picture1.Height >= HScroll1.Height + 4 * TwipsPerPixel Then 'PictureBox larger than ScrollBars


    X = Picture1.Width \ TwipsPerPixel - 4
    Y = Picture1.Height \ TwipsPerPixel - 4
    If TheImage.Width > X Then
        Y = Y - HScroll1.Height \ TwipsPerPixel
        If TheImage.Height > Y Then X = X - VScroll1.Width \ TwipsPerPixel
    Else
        If TheImage.Height > Y Then
            X = X - VScroll1.Width \ TwipsPerPixel
            If TheImage.Width > X Then Y = Y - HScroll1.Height \ TwipsPerPixel
        End If
    End If

    If TheImage.Width > X Then    'Add HScroll and set HScroll limits
        HScroll1.Min = 0
        HScroll1.Max = TheImage.Width - X
        HScroll1.Move 0, Picture1.Height - HScroll1.Height - 4 * TwipsPerPixel, Picture1.Width - IIf(TheImage.Height > Y, VScroll1.Width, 0) - 4 * TwipsPerPixel
        HScroll1.Visible = True
    Else                          'Remove HScroll and center picture
        HScroll1.Visible = False
        HScroll1.Min = (TheImage.Width - Picture1.Width \ TwipsPerPixel + 4 + IIf(TheImage.Height > Y, VScroll1.Width \ TwipsPerPixel, 0)) \ 2
        HScroll1.Max = HScroll1.Min
    End If

    If TheImage.Height > Y Then   'Add VScroll and set VScroll limits
        VScroll1.Min = 0
        VScroll1.Max = TheImage.Height - Y
        VScroll1.Move Picture1.Width - VScroll1.Width - 4 * TwipsPerPixel, 0, VScroll1.Width, Picture1.Height - 4 * TwipsPerPixel
        VScroll1.Visible = True
    Else                          'Remove VScroll and center picture
        VScroll1.Visible = False
        VScroll1.Min = (TheImage.Height - Picture1.Height \ TwipsPerPixel + 4 + IIf(HScroll1.Visible, HScroll1.Height \ TwipsPerPixel, 0)) \ 2
        VScroll1.Max = VScroll1.Min
    End If


        End If
    End If

    PaintImage m_Image

End Sub
Private Sub PaintImage(TheImage As cImage)
    If ObjPtr(TheImage) = 0 Then
        Picture1.Cls
    Else
        If HScroll1.Value < 0 Or VScroll1.Value < 0 Then Picture1.Cls
        TheImage.PaintHDC Picture1.hdc, -HScroll1.Value, -VScroll1.Value
        Picture1.Refresh
    End If
End Sub





'================================================================================
'                  ALLOW GRAB AND DRAG WITH LEFT MOUSE BUTTON
'================================================================================
Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = ButtonDrag Then Picture1.MousePointer = 0
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And ButtonDrag Then
        PaintLeft = (X \ TwipsPerPixel + HScroll1.Value)
        PaintTop = (Y \ TwipsPerPixel + VScroll1.Value)
        If HScroll1.Visible Then
            If VScroll1.Visible Then
                Picture1.MousePointer = 5 'Size
            Else
                Picture1.MousePointer = 9 'Size WE
            End If
        Else
            If VScroll1.Visible Then
                Picture1.MousePointer = 7 'Size NS
            End If
        End If
    End If
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And ButtonDrag Then
        Dim NewX As Long
        Dim NewY As Long
        NewX = PaintLeft - (X \ TwipsPerPixel)
        NewY = PaintTop - (Y \ TwipsPerPixel)
        If NewX < HScroll1.Min Then NewX = HScroll1.Min Else If NewX > HScroll1.Max Then NewX = HScroll1.Max
        If NewY < VScroll1.Min Then NewY = VScroll1.Min Else If NewY > VScroll1.Max Then NewY = VScroll1.Max
        HScroll1.Value = NewX
        VScroll1.Value = NewY
    End If
End Sub
Private Sub Picture1_Resize()
    AdjustScrollBars m_Image
End Sub
Private Sub HScroll1_Scroll()
    PaintImage m_Image
End Sub
Private Sub HScroll1_Change()
    PaintImage m_Image
End Sub
Private Sub VScroll1_Scroll()
    PaintImage m_Image
End Sub
Private Sub VScroll1_Change()
    PaintImage m_Image
End Sub





'================================================================================
'                         AUTOSIZE WINDOW TO PICTURE
'================================================================================
Public Sub SetFormSize(TheImage As cImage)
    Dim NewLeft         As Long
    Dim NewTop          As Long
    Dim newWidth        As Long
    Dim newHeight       As Long

    If ObjPtr(TheImage) <> 0 Then
        If Me.WindowState = 0 Then
            If TheImage.Width > 0 And TheImage.Height > 0 Then

    newWidth = (TheImage.Width + 4) * TwipsPerPixel + 120 + PictureBoxLeft + PictureBoxRight
    newHeight = (TheImage.Height + 4) * TwipsPerPixel + 420 + PictureBoxTop + PictureBoxBottom
    NewLeft = Me.Left + (Me.Width - newWidth) \ 2
    NewTop = Me.Top + (Me.Height - newHeight) \ 2

    If newHeight > Screen.Height Then
        NewTop = 0
        newHeight = Screen.Height
        newWidth = newWidth + VScroll1.Width
    Else
        If NewTop < 0 Then
            NewTop = 0
        Else
            If NewTop + newHeight > Screen.Height Then
                NewTop = Screen.Height - newHeight
            End If
        End If
    End If
    If newWidth > Screen.Width Then
        NewLeft = 0
        newWidth = Screen.Width
    Else
        If NewLeft < 0 Then
            NewLeft = 0
        Else
            If NewLeft + newWidth > Screen.Width Then
                NewLeft = Screen.Width - newWidth
            End If
        End If
    End If
    Me.Move NewLeft, NewTop, newWidth, newHeight

            End If
        End If
    End If

End Sub





'================================================================================
'                            MINIMAL IMAGE PROCESSING
'================================================================================
Private Sub mnuRotate_Click(Index As Integer)
    Select Case Index
    Case 0:    Set m_Image = m_Image.Rotate(-90)
    Case 1:    Set m_Image = m_Image.Rotate(90)
    Case 2:    Set m_Image = m_Image.Rotate(180)
    End Select
    If mnuAutosize.Checked Then SetFormSize m_Image
    AdjustScrollBars m_Image
End Sub
Private Sub mnuMirror_Click(Index As Integer)
    Set m_Image = m_Image.Mirror(Index <> 0)
    PaintImage m_Image
End Sub
Private Sub mnuAutosize_Click()
    mnuAutosize.Checked = Not mnuAutosize.Checked
    If mnuAutosize.Checked Then SetFormSize m_Image
End Sub





'================================================================================
'                             LOAD / SAVE PICTURE
'================================================================================
Private Sub mnuOpen_Click()
    Dim MyPic As StdPicture
    'Dim FileName As String

'    FileName = FileDialog(Me, False, "Open Picture File", "Picture Files|*.jpg;*.jpeg;*.gif;*.bmp;*.wmp;*.rle;*.cur;*.ico;*.emf|All Files [*.*]|*.*")
    FileName = "D:\My Webs\MatthewLan.Com Web\Photos\galleries\Funny-02\Fw- SMILE !~~image (2).jpg"
   'Call fSaveJPG.PicShow(FileName, PicShowReSize)
    FileName = "D:\My Webs\MatthewLan.Com Web\Photos\galleries\Fractals\23november2003d.jpg"
    
    'good code
    'Call PicShow(FileName, PicShowReSize)

    'Unload fMain
    'PicShowReSize.Showj
    'Exit Sub
    If Len(FileName) > 0 Then
        On Error Resume Next
        Set MyPic = LoadPicture(FileName)
        
        
        If Err.Number = 0 Then
            Set m_Image = New cImage
            m_Image.CopyStdPicture MyPic
            If mnuAutosize.Checked Then SetFormSize m_Image
            AdjustScrollBars m_Image
            Me.Caption = App.Title & " - " & FileTitleOnly(FileName)
            mnuSave(0).Enabled = True
        Else
            MsgBox "Can not load picture file" & vbCrLf & """" & FileName & """", vbExclamation, "File Load Error"
        End If
        Set MyPic = Nothing
    End If



End Sub

Private Sub mnuSave_Click(Index As Integer)
    'Dim FileName As String
    Dim SaveForm As fSaveJPG

    'FileName = FileDialog(Me, True, "Save As ...", "JPEG Files [*.jpg; *.jpeg]|*.jpg;*.jpeg|All Files [*.*]|*.*", , "*.jpg")
    If Len(FileName) > 0 Then
        Set SaveForm = New fSaveJPG
        SaveForm.SaveImage m_Image, FileName
        SaveForm.Show vbModal, Me
        Set SaveForm = Nothing
    End If

End Sub
Private Sub mnuExit_Click()
    Unload Me
End Sub
