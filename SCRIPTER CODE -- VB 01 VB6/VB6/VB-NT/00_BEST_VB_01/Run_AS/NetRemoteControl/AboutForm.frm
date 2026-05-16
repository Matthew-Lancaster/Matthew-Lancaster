VERSION 5.00
Begin VB.Form AboutForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Remote Control"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "AboutForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   1020
      TabIndex        =   8
      Top             =   3150
      Width           =   4995
   End
   Begin VB.CommandButton OKBttn 
      Caption         =   "OK"
      Height          =   405
      Left            =   4530
      TabIndex        =   6
      Top             =   3270
      Width           =   1215
   End
   Begin VB.CommandButton sysInfoBttn 
      Caption         =   "System Info..."
      Height          =   405
      Left            =   4530
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame fraAbout 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1343
      MouseIcon       =   "AboutForm.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2700
      Width           =   3315
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         Caption         =   "[bhagwat1981@yahoo.com]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   30
         Width           =   3075
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   960
      Left            =   1095
      Picture         =   "AboutForm.frx":074C
      ScaleHeight     =   900
      ScaleWidth      =   3750
      TabIndex        =   1
      Top             =   210
      Width           =   3810
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   1028
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   150
      Width           =   3945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "IndiSoft"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   210
      TabIndex        =   9
      Top             =   3045
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "For any comments and suggestions write me on bhagwat1981@yahoo.com."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1163
      TabIndex        =   7
      Top             =   1740
      Width           =   3675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Develop By: Rupera Bhagwat Singh"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1073
      TabIndex        =   2
      Top             =   1320
      Width           =   3855
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub fraAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShellExecute Me.hwnd, "open", "mailto:bhagwat1981@yahoo.com", _
    vbNullString, "C:\", 1&
    fraAbout_MouseMove 0, 0, -1, -1
End Sub

Private Sub fraAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ((X < 0) Or (X > fraAbout.Width)) Or ((Y < 0) Or (Y > fraAbout.Height)) Then
        lblAbout.ForeColor = vbBlack
        lblAbout.FontUnderline = False
        ReleaseCapture
    End If
End Sub

Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblAbout.ForeColor = vbBlue
    lblAbout.FontUnderline = True
    ReleaseCapture
    SetCapture fraAbout.hwnd
End Sub

Private Sub OKBttn_Click()
    Unload Me
End Sub
