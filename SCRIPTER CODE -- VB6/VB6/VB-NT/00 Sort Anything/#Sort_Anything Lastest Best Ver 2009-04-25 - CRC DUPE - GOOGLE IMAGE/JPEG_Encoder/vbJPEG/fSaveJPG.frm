VERSION 5.00
Begin VB.Form fSaveJPG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JPEG Compression Settings"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFinish 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "OK"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4215
      Begin VB.CheckBox chkGreyscale 
         Caption         =   "Greyscale"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdComment 
         Caption         =   "Comment"
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cboSubSample 
         Height          =   315
         ItemData        =   "fSaveJPG.frx":0000
         Left            =   1920
         List            =   "fSaveJPG.frx":0034
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   2055
      End
      Begin VB.HScrollBar hscQuality 
         Height          =   255
         Left            =   1920
         Max             =   100
         Min             =   1
         TabIndex        =   0
         Top             =   360
         Value           =   1
         Width           =   2055
      End
      Begin VB.Label lblSubSample 
         Caption         =   "Sub Sampling:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblQuality 
         Caption         =   "Quality:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "fSaveJPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Option Explicit

Public m_Image     As cImage


Private m_Jpeg      As cJpeg
Private m_FileName  As String


Public Sub SaveImage(TheImage As cImage, FileName As String)
    Set m_Image = TheImage 'Call this before the form loads to initialize it
    m_FileName = FileName
End Sub

Private Sub Form_Load()
    Set m_Jpeg = New cJpeg
    cboSubSample.ListIndex = 0
    'cboSubSample.ListIndex = 3'Default
    hscQuality.Value = 98
End Sub



Private Sub cboSubSample_Click()
    Dim h As Long
    Dim v As Long

    If chkGreyscale.Value = 0 Then
        cboSubSample.Enabled = True
        h = Val(Mid$(cboSubSample.List(cboSubSample.ListIndex), 1, 1)) 'Get horizontal luminance sampling factor
        v = Val(Mid$(cboSubSample.List(cboSubSample.ListIndex), 3, 1)) 'Get vertical   luminance sampling factor
        m_Jpeg.SetSamplingFrequencies h, v, 1, 1, 1, 1
    Else 'Greyscale
        cboSubSample.Enabled = False
        m_Jpeg.SetSamplingFrequencies 1, 1, 0, 0, 0, 0
    End If

End Sub
Private Sub chkGreyscale_Click()
    cboSubSample_Click
End Sub

Private Sub hscQuality_Change()
    Dim q As Long

    q = hscQuality.Value
    lblQuality.Caption = "Quality: " & CStr(q) & IIf(q < 50 Or q > 95, " ???", "")
    m_Jpeg.Quality = q
End Sub


Private Sub cmdComment_Click()
    Dim NewComment As New fComment

    NewComment.Comment = m_Jpeg.Comment
    NewComment.Show vbModal, Me
    m_Jpeg.Comment = NewComment.Comment
    Set NewComment = Nothing
End Sub

Sub ReSizeCords(AdjustX, AdjustY)

      'ax = Picture1.ScaleWidth
      'ay = Picture1.ScaleHeight
      ax = m_Image.Width
      ay = m_Image.Height
      
      'ax2$ = Trim$(Round(ax, 0))
      'ay2$ = Trim$(Round(ay, 0))
      
      ratio = 1 + (1 / 2000) 'Slower = More Passes
      ratio = 1 + (1 / 200) 'Faster = Less Passes
      
      compsize = 1
      
      Do
      Tag = 0
      If ax > AdjustX Or ay > AdjustY Then
      ax = ax / ratio: ay = ay / ratio: Tag = 1
      compsize = compsize / ratio
      End If
      
      If ax < AdjustX - 0.1 And ay < ajusty - 0.1 Then
      ax = ax * ratio: ay = ay * ratio: Tag = 1
      compsize = compsize * ratio
      End If
      
      DoEvents
      Counter = Counter + 1
      Loop Until Tag = 0
      
      AdjustX = ax
      AdjustY = ay
      
      passes = Counter
      comps$ = Format$(compsize, "##0.0000")

        'List1.AddItem ax1$ + "x" + ay1$ + "    " + B$


End Sub

Private Sub cmdFinish_Click(Index As Integer)

    If Index = 1 Then
       'Sample the cImage by hDC
        'PicShow m_FileName, PicShowReSize
        'How to call the function
        'Call PicShow(m_FileName, PicShowReSize)
    
        'Set m_Image = New cImage
'        Set m_Image3 = New StdPicture
'        Set m_Image2 = m_Image
'        m_Image3 = PowerResize.PowerResize(m_Image2, 120, 90) 'As StdPicture
        
        'How to call the function
'        Call PicShow(FileName, PicShowReSize)
        Dim AdjustX As Long
        Dim AdjustY As Long

        AdjustX = 150
        AdjustY = 150
        Call ReSizeCords(AdjustX, AdjustY)
        
        Set m_Image = m_Image.Resample(AdjustX, AdjustY)
        '.ViewImage.Picture
        
        m_Jpeg.SampleHDC m_Image.hdc, m_Image.Width, m_Image.Height
        
        'Set PicShowReSize.ViewImage.Picture = fMain.Picture1.Picture
        
        'PicShowReSize.ViewImage.Picture = fMain.Picture1.Picture
        
        'PicShowReSize.ViewImage
        
        
        'Call PicShow(FileName, PicShowReSize)
        
        
        
        'tt = cImage.Resample(120, 90)
        'm_Jpeg.SampleHDC m_Image.hDC, 120, 90
        m_Jpeg.SampleHDC m_Image.hdc, m_Image.Width, m_Image.Height
'        m_Jpeg.SampleHDC m_Image.hDC, 120, 90

        Dim NewFileName As String
       
        NewFileName = m_FileName
        NewFileName = Mid$(NewFileName, 1, InStrRev(NewFileName, ".") - 1) + "_Test.jpg"
       
        'Delete file if it exists
        RidFile NewFileName

       'Save the JPG file
        m_Jpeg.SaveFile NewFileName
    End If

    Set m_Image = Nothing
    Set m_Jpeg = Nothing
    Unload Me

End Sub

