VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "X"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   5775
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Default.wav"
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      Caption         =   "Default Path:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblFilename 
      BackStyle       =   0  'Transparent
      Caption         =   "Default Filename:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetupDir()
    On Error Resume Next

    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BROWSEINFO

    szTitle = "Please select the default path to save the recording in." & vbCrLf
    szTitle = szTitle & "To create a new folder, click 'Make New Folder'"

    With tBrowseInfo
        .hwndOwner = Me.hwnd
        .lpszTitle = szTitle & vbNullChar
        .ulFlags = BROWSE_FLAGS.BIF_RETURNONLYFSDIRS + BROWSE_FLAGS.BIF_DONTGOBELOWDOMAIN + BROWSE_FLAGS.BIF_USENEWUI
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(1, sBuffer, vbNullChar) - 1)
        txtPath.Text = sBuffer
    End If
End Sub

Private Sub cmdApply_Click()
    SaveSetting App.EXEName, App.EXEName, "Default Filename", txtFilename.Text
    SaveSetting App.EXEName, App.EXEName, "Default Path", txtPath.Text
    Unload Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPath_Click()
    Dim blnRepeat As Boolean
    On Error Resume Next

Call_Setup:

    blnRepeat = False
    SetupDir
    If blnRepeat Then
        GoTo Call_Setup
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim Y As Integer
    On Error GoTo Error_Handler
    Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    txtFilename.Text = GetSetting(App.EXEName, App.EXEName, "Default Filename")
    txtPath.Text = GetSetting(App.EXEName, App.EXEName, "Default Path")

    Me.AutoRedraw = True
    Me.DrawStyle = 6
    Me.DrawMode = 13
    Me.DrawWidth = 13
    Me.ScaleMode = 3

    Me.ScaleHeight = (256)
    For i = 0 To 255
        Me.Line (0, Y)-(Me.Width, Y + 1), (vbBlack) - (i * (vbBlack - vbGreen) / 255), BF
        Y = Y + 1
    Next i
Error_Handler:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Enabled = True
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call ReleaseCapture
        Call SendMessage(Me.hwnd, &HA1, 2, 0)
    End If
End Sub
