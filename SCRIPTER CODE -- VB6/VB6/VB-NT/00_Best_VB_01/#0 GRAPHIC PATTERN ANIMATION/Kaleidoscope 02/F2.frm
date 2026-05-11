VERSION 5.00
Begin VB.Form F2 
   Caption         =   "Pattern2"
   ClientHeight    =   7224
   ClientLeft      =   216
   ClientTop       =   912
   ClientWidth     =   12972
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   ScaleHeight     =   602
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1081
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6396
      Top             =   1752
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2112
      Top             =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   16.2
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1512
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Label8_TIMER"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   16.2
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   552
      Left            =   3240
      TabIndex        =   5
      Top             =   3900
      Width           =   2796
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   16.2
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3240
      TabIndex        =   4
      Top             =   2940
      Width           =   1560
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   16.2
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   576
      Left            =   3240
      TabIndex        =   3
      Top             =   3444
      Width           =   2004
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   16.2
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Left            =   3240
      TabIndex        =   2
      Top             =   2484
      Width           =   1548
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   16.2
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   624
      Left            =   3240
      TabIndex        =   1
      Top             =   1572
      Width           =   1584
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   16.2
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3240
      TabIndex        =   0
      Top             =   2004
      Width           =   1752
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "F2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----
'Making A Form Transparent (But with visible controls)-VBForums
'http://www.vbforums.com/showthread.php?396385-Making-A-Form-Transparent-
'----
'----
'VB6 HOW MAKE FORM TRANSPARENT - Google Search
'https://www.google.co.uk/search?q=VB6+HOW+MAKE+FORM+TRANSPARENT
'----

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
                
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long
 
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
 
Private Sub Form_Click()
F1.HighRes.Enabled = True
End Sub

Private Sub Form_Load()
    Me.BackColor = vbBlack 'vbCyan
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    ' SetLayeredWindowAttributes Me.hwnd, vbCyan, 0&, LWA_COLORKEY
    SetLayeredWindowAttributes Me.hwnd, vbBlack, 0&, LWA_COLORKEY
    
    
End Sub
 
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
F1.HighRes.Enabled = True
End Sub

Private Sub Form_Resize()
If F2.WindowState = vbNormal Then Exit Sub
F1.WindowState = F2.WindowState
If F2.WindowState = vbMinimized Then F1.HighRes.Enabled = False
If F2.WindowState = vbMaximized Then F1.HighRes.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
DoEvents
F2.WindowState = F1.WindowState
F2.ScaleMode = F1.ScaleMode
F2.ScaleWidth = F1.ScaleWidth
F2.ScaleHeight = F1.ScaleHeight
' F2.BackColor = F1.BackColor

End Sub

'Private Sub Timer2_Timer()
'Timer2.Enabled = False
'F1.HighRes.Enabled = False
'End Sub
