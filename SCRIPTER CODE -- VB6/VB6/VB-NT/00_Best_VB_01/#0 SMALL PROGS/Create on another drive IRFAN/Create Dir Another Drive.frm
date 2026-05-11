VERSION 5.00
Begin VB.Form IR 
   Caption         =   "DIR ON ANOTHER DRIVE"
   ClientHeight    =   3990
   ClientLeft      =   165
   ClientTop       =   840
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5160
      Top             =   2160
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   -30
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Create Dir Another Drive.frx":0000
      Top             =   585
      Width           =   11310
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   720
      TabIndex        =   39
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   2040
      TabIndex        =   38
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   2040
      TabIndex        =   37
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   2280
      TabIndex        =   36
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   2040
      TabIndex        =   35
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   2280
      TabIndex        =   34
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   2280
      TabIndex        =   33
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   7
      Left            =   2520
      TabIndex        =   32
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   8
      Left            =   2160
      TabIndex        =   31
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   9
      Left            =   2400
      TabIndex        =   30
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   10
      Left            =   2400
      TabIndex        =   29
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   11
      Left            =   2640
      TabIndex        =   28
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   12
      Left            =   2400
      TabIndex        =   27
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   13
      Left            =   2640
      TabIndex        =   26
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   14
      Left            =   2640
      TabIndex        =   25
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   15
      Left            =   2880
      TabIndex        =   24
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   16
      Left            =   1200
      TabIndex        =   23
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   17
      Left            =   2280
      TabIndex        =   22
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   18
      Left            =   2280
      TabIndex        =   21
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   19
      Left            =   2520
      TabIndex        =   20
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   20
      Left            =   2280
      TabIndex        =   19
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   21
      Left            =   2520
      TabIndex        =   18
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   22
      Left            =   2520
      TabIndex        =   17
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   23
      Left            =   2760
      TabIndex        =   16
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   24
      Left            =   2400
      TabIndex        =   15
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   25
      Left            =   2640
      TabIndex        =   14
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   26
      Left            =   2640
      TabIndex        =   13
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   27
      Left            =   2880
      TabIndex        =   12
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   28
      Left            =   2640
      TabIndex        =   11
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   29
      Left            =   2880
      TabIndex        =   10
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   30
      Left            =   2880
      TabIndex        =   9
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   31
      Left            =   3120
      TabIndex        =   8
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C000&
      Caption         =   "INFAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9480
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C000&
      Caption         =   "GO BACK 50 INFAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6240
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C000&
      Caption         =   "GO BACK BASE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C000&
      Caption         =   "GO BACK 500 INFAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4440
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "LOAD INFAR INI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1920
      TabIndex        =   3
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "LOAD EXPLORER SELECT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Change Path Press Enter // CLIP ANOTHER PATH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11295
   End
   Begin VB.Menu VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_FORM_DIR_SOURCE 
      Caption         =   "FORM DIR SOURCE"
   End
End
Attribute VB_Name = "IR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A1

Dim FORMLOAD

Private Sub Form_Load()
 'Explorer.exe /select,
ii = Clipboard.GetText
'If Mid$(ii, 2, 2) <> ":\" Then MsgBox "DIR WAS ROOT OF DRIVE -- END": End

'ii = "D:\VB6\VB-NT\00_Send_To\Clipboard Dir Create on another drive"
   
    
    
    
    
    
On Error Resume Next

'If App.PrevInstance = True Then End

Text1 = ii
If Mid$(Text1, Len(Text1), 1) <> "\" Then
    Text1 = Text1 + "\"
End If


FORMLOAD = True

Call Label3_Click

LabDR(0).Left = 0

XC = 65
For Each Control In LabDR
    Control.Visible = False
    Control.AutoSize = False
    Control.Caption = "-" + Chr(XC) + "-"
    DoEvents
    Control.AutoSize = True
    Control.Refresh
    DoEvents
    'If XC > 90 Then Exit For
    XC = XC + 1
    If XC Mod 2 = 0 Then Control.BackColor = &HC0FFC0
    If Control.Width > XT Then XT = Control.Width
Next

XT = (Text1.Width - 16) / 26
'XT = XT
XC = 65
XD = LabDR(0).Left = 0
For Each Control In LabDR
    DoEvents
    'CONTROL.INDEX
    Control.Top = Text1.Top + Text1.Height + 15
    Control.Height = Control.Height + 40
    Control.Left = XD
    Control.Width = XT
    XD = XD + XT
    If XC > 90 Then Exit For
    XC = XC + 1
    Control.Visible = True
Next

XC = 65
XD = LabDR(0).Left = 0
For Each Control In LabDR
    DoEvents
    'CONTROL.INDEX
    Control.Top = Text1.Top + Text1.Height + 15
    Control.Height = Control.Height + 40
    Control.Left = XD
    Control.Width = XT
    XD = XD + XT
    If XC > 90 Then Exit For
    XC = XC + 1
    Control.Visible = True
Next



Me.WindowState = vbNormal

INFAR.Show

INFAR.SetFocus

IR.WindowState = vbMinimized

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub LabDR_Click(Index As Integer)
Text1.Text = Chr(Index + 65) + Mid(Text1.Text, 2)
End Sub

Private Sub Label1_Click()
ii = Clipboard.GetText
'If Mid$(ii, 2, 2) <> ":\" Then MsgBox "DIR WAS ROOT OF DRIVE -- END": End

'ii = "D:\VB6\VB-NT\00_Send_To\Clipboard Dir Create on another drive"
    On Error Resume Next
Text1 = ii
If Mid$(Text1, Len(Text1), 1) <> "\" Then
    Text1 = Text1 + "\"
End If


End Sub

Private Sub Label2_Click()

Me.WindowState = vbMinimized

Shell "Explorer.exe /select, " + Text1, vbNormalFocus

End Sub

Private Sub Label3_Click()

Me.WindowState = vbMinimized

Open "C:\Program Files\IrfanView\i_view32.ini" For Input As #1
    Do
        Line Input #1, L
        If InStr("-" + L, "StartIndex=") > 0 Then
            A1 = Val(Mid(L, Len("StartIndex=") + 1))
        End If
        If EOF(1) Then A1 = "2"
    Loop Until A1 <> ""
Close #1
Label3 = "LOAD INFAR INI -- " + Str(A1)

'Open "D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDESHOW.txt" For Input As #1
Open "D:\#\#D\ONE MONTH OF SEPT\IRFAN SLIDESHOW.txt" For Input As #1
'
    Do
        Line Input #1, L
        U = U + 1
        If U = A1 Then A2 = L
    Loop Until EOF(1) Or A2 <> ""
Close #1


If Dir(Text1) = "" And Dir(Text1, vbDirectory) = "" Then
Text1 = A2

End If

If FORMLOAD = True Then FORMLOAD = False: Exit Sub
'StartIndex=548


Shell "Explorer.exe /select, " + A2, vbNormalFocus

End Sub

Private Sub Label4_Click()
Me.WindowState = vbMinimized

Open "C:\Program Files\IrfanView\i_view32.ini" For Input As #1
    Do
        Line Input #1, L
        If InStr("-" + L, "StartIndex=") > 0 Then
            A1 = Val(Mid(L, Len("StartIndex=") + 1))
        End If
        If EOF(1) Then A1 = "2"
    Loop Until A1 <> ""
Close #1

FR = FreeFile
Open "C:\Program Files\IrfanView\i_view32.ini" For Input As #FR
RC = Input(LOF(FR), FR)
Close #FR

A1 = Trim(Str(A1))
B1 = Trim(Str(A1 - 500))
RC = Replace(RC, "StartIndex=" + A1, "StartIndex=" + B1)
Label3 = "LOAD INFAR INI -- " + Str(B1)


FR = FreeFile
Open "C:\Program Files\IrfanView\i_view32.ini" For Output Lock Read Write As #FR
Print #FR, RC
Close #FR


End Sub

Private Sub Label5_Click()

Me.WindowState = vbMinimized

Open "C:\Program Files\IrfanView\i_view32.ini" For Input As #1
    Do
        Line Input #1, L
        If InStr("-" + L, "StartIndex=") > 0 Then
            A1 = Val(Mid(L, Len("StartIndex=") + 1))
        End If
        If EOF(1) Then A1 = "2"
    Loop Until A1 <> ""
Close #1

FR = FreeFile
Open "C:\Program Files\IrfanView\i_view32.ini" For Input As #FR
RC = Input(LOF(FR), FR)
Close #FR

A1 = Trim(Str(A1))
B1 = Trim(Str(1))
RC = Replace(RC, "StartIndex=" + A1, "StartIndex=" + B1)
Label3 = "LOAD INFAR INI -- " + Str(B1)


FR = FreeFile
Open "C:\Program Files\IrfanView\i_view32.ini" For Output Lock Read Write As #FR
Print #FR, RC
Close #FR









End Sub

Private Sub Label6_Click()

Open "C:\Program Files\IrfanView\i_view32.ini" For Input As #1
    Do
        Line Input #1, L
        If InStr("-" + L, "StartIndex=") > 0 Then
            A1 = Val(Mid(L, Len("StartIndex=") + 1))
        End If
        If EOF(1) Then A1 = "2"
    Loop Until A1 <> ""
Close #1

FR = FreeFile
Open "C:\Program Files\IrfanView\i_view32.ini" For Input As #FR
RC = Input(LOF(FR), FR)
Close #FR

A1 = Trim(Str(A1))
B1 = Trim(Str(A1 - 50))
RC = Replace(RC, "StartIndex=" + A1, "StartIndex=" + B1)
Label3 = "LOAD INFAR INI -- " + Str(B1)


FR = FreeFile
Open "C:\Program Files\IrfanView\i_view32.ini" For Output Lock Read Write As #FR
Print #FR, RC
Close #FR

End Sub

Private Sub Label7_Click()

Me.WindowState = vbMinimized

Shell "'C:\WINDOWS\system32\CMD.EXE ""D:\#\#D\ONE MONTH OF SEPT\i_view32.exe -- IRFAN SLIDESHOW.BAT""  /C", vbMaximizedFocus

End Sub

Private Sub MNU_FORM_DIR_SOURCE_Click()

Call INFAR.Form_Load


End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'    a = a
If Mid$(Text1, 2, 2) <> ":\" Then MsgBox "DIR WAS ROOT OF DRIVE -- END": End
    
    
    On Error Resume Next
    R = 4
    Do
                
        t = InStr(R + 1, Text1, "\")
        dd = Mid$(Text1, 1, t - 1)
        R = t
        Err.Clear
        If t > 0 Then MkDir dd
        yy = Err.Description
        uu = Err.Number
    
    Loop Until t = 0
End

End If

End Sub

Private Sub Timer1_Timer()

A1 = A1 + 1

If A1 > 10 Then
Timer1.Enabled = False
End If

IR.WindowState = vbMinimized
INFAR.Show
INFAR.SetFocus

End Sub

Private Sub VB_ME_Click()

Shell "Explorer.exe /select, " + App.Path + "\" + App.EXEName + ".vbp", vbNormalNoFocus
Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus

End

End Sub
