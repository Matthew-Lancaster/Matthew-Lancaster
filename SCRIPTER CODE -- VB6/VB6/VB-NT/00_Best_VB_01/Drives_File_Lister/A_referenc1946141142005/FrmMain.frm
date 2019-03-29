VERSION 5.00
Begin VB.Form Mp3_Tag 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3 tagging demo app"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11100
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkv2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ID3 v2 tag"
      Height          =   255
      Left            =   5640
      TabIndex        =   46
      Top             =   1860
      Width           =   1155
   End
   Begin VB.CheckBox chkv1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ID3 v1 tag"
      Height          =   255
      Left            =   180
      TabIndex        =   45
      Top             =   1860
      Width           =   1155
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update tags"
      Enabled         =   0   'False
      Height          =   435
      Left            =   180
      TabIndex        =   44
      Top             =   7920
      Width           =   3135
   End
   Begin VB.Frame frav2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ID3 v2 tag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6075
      Left            =   5640
      TabIndex        =   30
      Top             =   2340
      Width           =   5235
      Begin VB.ComboBox cmbVersion 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   2220
         List            =   "frmMain.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   4860
         Width           =   1095
      End
      Begin VB.CheckBox chkClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clear existing ID3 v2 tag before writing the new one."
         Height          =   195
         Left            =   180
         TabIndex        =   57
         Top             =   5760
         Width           =   4575
      End
      Begin VB.TextBox v2LinkTo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         Top             =   4020
         Width           =   3975
      End
      Begin VB.TextBox v2Encoder 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Top             =   4380
         Width           =   3975
      End
      Begin VB.TextBox v2Composer 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Top             =   2940
         Width           =   3975
      End
      Begin VB.TextBox v2OrigArtist 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   3300
         Width           =   3975
      End
      Begin VB.TextBox v2Copyright 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   3660
         Width           =   3975
      End
      Begin VB.TextBox v2Artist 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox v2Album 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox v2Title 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox v2TNR 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox v2Comment 
         Appearance      =   0  'Flat
         Height          =   1065
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1800
         Width           =   3975
      End
      Begin VB.ComboBox v2Genre 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3660
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   1440
         Width           =   1395
      End
      Begin VB.TextBox v2Year 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   12
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Write ID3 v2 tag in version"
         Height          =   195
         Left            =   180
         TabIndex        =   60
         Top             =   4920
         Width           =   1890
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Note that v 2.3 still has greater support. For example, Winamp can't read v 2.4 tags."
         Height          =   435
         Left            =   180
         TabIndex        =   59
         Top             =   5220
         Width           =   4875
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Encoder:"
         Height          =   195
         Left            =   180
         TabIndex        =   42
         Top             =   4440
         Width           =   645
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Link:"
         Height          =   195
         Left            =   180
         TabIndex        =   41
         Top             =   4080
         Width           =   345
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Copyright:"
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Top             =   3720
         Width           =   705
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Orig. artist:"
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Top             =   3360
         Width           =   750
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Composer:"
         Height          =   195
         Left            =   180
         TabIndex        =   38
         Top             =   3000
         Width           =   750
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comment:"
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   2220
         Width           =   705
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Genre:"
         Height          =   195
         Left            =   3060
         TabIndex        =   36
         Top             =   1500
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Year:"
         Height          =   195
         Left            =   1800
         TabIndex        =   35
         Top             =   1500
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Track nr.:"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Title:"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   1140
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Album:"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   780
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Artist:"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   420
         Width           =   390
      End
   End
   Begin VB.Frame frav1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ID3 v1 tag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2235
      Left            =   180
      TabIndex        =   22
      Top             =   2340
      Width           =   5235
      Begin VB.TextBox v1Year 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1440
         Width           =   615
      End
      Begin VB.ComboBox v1Genre 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3660
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   1395
      End
      Begin VB.TextBox v1Comment 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   28
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1800
         Width           =   3975
      End
      Begin VB.TextBox v1TNR 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox v1Title 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox v1Album 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   2
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox v1Artist 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   1
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Artist:"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   420
         Width           =   390
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Album:"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   780
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Title:"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   1140
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Track nr.:"
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Top             =   1500
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Year:"
         Height          =   195
         Left            =   1800
         TabIndex        =   25
         Top             =   1500
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Genre:"
         Height          =   195
         Left            =   3060
         TabIndex        =   24
         Top             =   1500
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comment:"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   1860
         Width           =   705
      End
   End
   Begin VB.TextBox txtMp3File 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2340
      Locked          =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   20
      Top             =   1080
      Width           =   8535
   End
   Begin VB.CommandButton cmdPick 
      Caption         =   "select"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1020
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   180
      X2              =   10860
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "This tagging module is part of Magic MP3 Tagger. Visit www.magic-tagger.com if you're interested in this top MP3 tagger!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   180
      TabIndex        =   61
      Top             =   600
      Width           =   10335
   End
   Begin VB.Label lblv2Vers 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   180
      TabIndex        =   56
      Top             =   6960
      Width           =   45
   End
   Begin VB.Label lblv1Vers 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   180
      TabIndex        =   55
      Top             =   6720
      Width           =   45
   End
   Begin VB.Label lblEmphasis 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   180
      TabIndex        =   54
      Top             =   6480
      Width           =   45
   End
   Begin VB.Label lblCRC 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   180
      TabIndex        =   53
      Top             =   6240
      Width           =   45
   End
   Begin VB.Label lblOrig 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   180
      TabIndex        =   52
      Top             =   6000
      Width           =   45
   End
   Begin VB.Label lblCopy 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   180
      TabIndex        =   51
      Top             =   5760
      Width           =   45
   End
   Begin VB.Label lblFreq 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   180
      TabIndex        =   50
      Top             =   5520
      Width           =   45
   End
   Begin VB.Label lblkBit 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   180
      TabIndex        =   49
      Top             =   5280
      Width           =   45
   End
   Begin VB.Label lblLen 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   180
      TabIndex        =   48
      Top             =   5040
      Width           =   45
   End
   Begin VB.Label lblMPEG 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   180
      TabIndex        =   47
      Top             =   4800
      Width           =   45
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"frmMain.frx":0018
      Height          =   435
      Left            =   180
      TabIndex        =   43
      Top             =   120
      Width           =   10695
   End
   Begin VB.Label lblMp3File 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "MP3 File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   180
      TabIndex        =   21
      Top             =   1080
      Width           =   750
   End
End
Attribute VB_Name = "Mp3_Tag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Long
    nFileExtension As Long
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const OFN_EXPLORER = &H80000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const cSingleSelFlags As Long = OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private myTag As ID3Tag, RdFile As String




Private Sub chkv1_Click()
    frav1.Visible = IIf(chkv1.Value = 1, True, False)
End Sub

Private Sub chkv2_Click()
    frav2.Visible = IIf(chkv2.Value = 1, True, False)
End Sub

Private Sub cmdUpdate_Click()
    If RdFile = "" Then Exit Sub
    
    If chkv1.Value = 0 Then
        mdlID3.DeleteID3v1 RdFile
    Else
        With myTag
            .Artist = v1Artist
            .Album = v1Album
            .Title = v1Title
            .TrackNr = v1TNR
            .SongYear = v1Year
            'Just write the genre string to the tag structure - the ID3 module
            'is going to write the tag correctly.
            .Genre = v1Genre.List(v1Genre.ListIndex)
            .Comment = v1Comment
        End With
        mdlID3.WriteID3v1 RdFile, myTag
    End If
    
    If chkv2.Value = 0 Then
        mdlID3.DeleteID3v2 RdFile
    Else
        With myTag
            .Artist = v2Artist
            .Album = v2Album
            .Title = v2Title
            .TrackNr = v2TNR
            .SongYear = v2Year
            'Just write the genre string to the tag structure - the ID3 module
            'is going to write the tag correctly.
            .Genre = v2Genre
            .Comment = v2Comment
            .Composer = v2Composer
            .OrigArtist = v2OrigArtist
            .Copyright = v2Copyright
            .LinkTo = v2LinkTo
            .EncodedBy = v2Encoder
        End With
        
        'The WriteID3v2 function can be passed a buffer pointer as parameter. See the comment
        'at the function implementation for details.
        mdlID3.WriteID3v2 RdFile, myTag, IIf(cmbVersion.ListIndex = 0, VERSION_2_3, VERSION_2_4), IIf(chkClear.Value = 1, True, False)
    End If
    MsgBox "Tags were updated.", vbInformation
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    'Add the genre strings to the combo boxes.
    v1Genre.AddItem ""
    For i = 0 To 147
        v1Genre.AddItem mdlID3.GetGenreName(i)
        v2Genre.AddItem mdlID3.GetGenreName(i)
    Next i
    cmbVersion.ListIndex = 0
    chkv1.Visible = False
    chkv2.Visible = False
    frav1.Visible = False
    frav2.Visible = False
End Sub


Private Sub cmdPick_Click()
    Dim i As Long, OFName As OPENFILENAME
    Dim GenreIdx As Long, outInfo As MPEGInfo
    
    On Local Error Resume Next
    
'    With OFName
'        .lStructSize = Len(OFName)
'        .hWndOwner = Me.hWnd
'        .hInstance = App.hInstance
'        .lpstrFilter = "MP3 files (*.mp3)" & Chr$(0) & "*.mp3"
'        .lpstrTitle = "Select a mp3 file"
'        .Flags = cSingleSelFlags
'        .lpstrFile = Space$(1023)
'        .nMaxFile = 1024
'    End With

    'Show the open form.
'    If Not GetOpenFileName(OFName) = 0 Then
'        RdFile = Split(Trim$(OFName.lpstrFile), Chr(0))(0)
        txtMp3File = RdFile
'    Else
'        Exit Sub
'    End If
    
    'Reset the local tag structure, and read the ID3 v1 tag.
    mdlID3.ResetTag myTag
    chkv1.Visible = True
    chkv1.Value = IIf(mdlID3.ReadID3v1(RdFile, myTag), 1, 0)
    'Fill in our form... Pretty easy...
    With myTag
        v1Artist = .Artist
        v1Album = .Album
        v1Title = .Title
        v1TNR = .TrackNr
        v1Year = .SongYear
        v1Genre.ListIndex = 0
        For i = 0 To v1Genre.ListCount - 1
            If v1Genre.List(i) = .Genre Then
                v1Genre.ListIndex = i
                Exit For
            End If
        Next i
        v1Comment = .Comment
    End With
    
    'Reset the local tag structure, and read the ID3 v2 tag.
    mdlID3.ResetTag myTag
    chkv2.Visible = True
    chkv2.Value = IIf(mdlID3.ReadID3v2(RdFile, myTag), 1, 0)
    'And again...
    With myTag
        v2Artist = .Artist
        v2Album = .Album
        v2Title = .Title
        v2TNR = .TrackNr
        v2Year = .SongYear
        v2Genre = .Genre
        v2Comment = .Comment
        v2Composer = .Composer
        v2OrigArtist = .OrigArtist
        v2Copyright = .Copyright
        v2LinkTo = .LinkTo
        If Trim$(v2LinkTo) = "" Then v2LinkTo = "www.magic-tagger.com"
        v2Encoder = .EncodedBy
        If Trim$(v2Encoder) = "" Then v2Encoder = "Magic MP3 Tagger"
    End With
    
    mdlID3.ReadMPEGInfo RdFile, outInfo
    lblMPEG = outInfo.MPEGVersion
    lblLen = "Length: " & outInfo.Length \ 60 & ":" & Format$(outInfo.Length Mod 60, "00")
    lblkBit = "Bitrate: " & outInfo.Bitrate & " kbit / sec, " & IIf(outInfo.HasVBR, "variable bit rate", "constant bit rate")
    lblFreq = "Frequency: " & outInfo.Frequency & " hz, " & LCase$(outInfo.ChannelMode)
    lblCopy = "Copyrighted: " & IIf(outInfo.IsCopyrighted, "yes", "no")
    lblOrig = "Original: " & IIf(outInfo.IsOriginal, "yes", "no")
    lblCRC = "Uses checksums: " & IIf(outInfo.HasCRC, "yes", "no")
    lblEmphasis = "Uses emphasis: " & IIf(outInfo.HasEmphasis, "yes", "no")
    lblv1Vers = "ID3 v1 tag: " & IIf(outInfo.ID3v1Version = -1, "none", "version 1." & outInfo.ID3v1Version)
    lblv2Vers = "ID3 v2 tag: " & IIf(outInfo.ID3v2Version = -1, "none", "version 2." & outInfo.ID3v2Version)
    
    cmdUpdate.Enabled = True
End Sub

