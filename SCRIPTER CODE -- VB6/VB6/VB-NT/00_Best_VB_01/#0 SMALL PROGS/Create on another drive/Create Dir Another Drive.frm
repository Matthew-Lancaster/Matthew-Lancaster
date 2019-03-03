VERSION 5.00
Begin VB.Form IR 
   Caption         =   "DIR ON ANOTHER DRIVE"
   ClientHeight    =   5940
   ClientLeft      =   192
   ClientTop       =   540
   ClientWidth     =   11628
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   11628
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1164
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   41
      Text            =   "Create Dir Another Drive.frx":0000
      Top             =   2760
      Width           =   11304
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5232
      Top             =   4224
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
      Text            =   "Create Dir Another Drive.frx":0008
      Top             =   570
      Width           =   11310
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF0000&
      Caption         =   "Change to the Clipboard into Current Path Below -- Hitt Here"
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
      Height          =   400
      Left            =   0
      TabIndex        =   42
      Top             =   2100
      Width           =   11292
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808080&
      Caption         =   "Label4"
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   4332
      TabIndex        =   40
      Top             =   4272
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   0
      Left            =   792
      TabIndex        =   39
      Top             =   4104
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   1
      Left            =   2112
      TabIndex        =   38
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   2
      Left            =   2112
      TabIndex        =   37
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   3
      Left            =   2352
      TabIndex        =   36
      Top             =   4236
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   4
      Left            =   2112
      TabIndex        =   35
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   5
      Left            =   2352
      TabIndex        =   34
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   6
      Left            =   2352
      TabIndex        =   33
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   7
      Left            =   2592
      TabIndex        =   32
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   8
      Left            =   2232
      TabIndex        =   31
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   9
      Left            =   2472
      TabIndex        =   30
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   10
      Left            =   2472
      TabIndex        =   29
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   11
      Left            =   2712
      TabIndex        =   28
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   12
      Left            =   2472
      TabIndex        =   27
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   13
      Left            =   2712
      TabIndex        =   26
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   14
      Left            =   2712
      TabIndex        =   25
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   15
      Left            =   2952
      TabIndex        =   24
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   16
      Left            =   1272
      TabIndex        =   23
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   17
      Left            =   2352
      TabIndex        =   22
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   18
      Left            =   2352
      TabIndex        =   21
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   19
      Left            =   2592
      TabIndex        =   20
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   20
      Left            =   2352
      TabIndex        =   19
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   21
      Left            =   2592
      TabIndex        =   18
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   22
      Left            =   2592
      TabIndex        =   17
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   23
      Left            =   2832
      TabIndex        =   16
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   24
      Left            =   2472
      TabIndex        =   15
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   25
      Left            =   2712
      TabIndex        =   14
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   26
      Left            =   2712
      TabIndex        =   13
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   27
      Left            =   2952
      TabIndex        =   12
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   28
      Left            =   2712
      TabIndex        =   11
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   29
      Left            =   2952
      TabIndex        =   10
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   30
      Left            =   2952
      TabIndex        =   9
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label LabDR 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "-A-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   31
      Left            =   3192
      TabIndex        =   8
      Top             =   4224
      Width           =   372
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C000&
      Caption         =   "INFAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   9552
      TabIndex        =   7
      Top             =   4824
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C000&
      Caption         =   "GO BACK 50 INFAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   6312
      TabIndex        =   6
      Top             =   4824
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C000&
      Caption         =   "GO BACK BASE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   8112
      TabIndex        =   5
      Top             =   4824
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C000&
      Caption         =   "GO BACK 500 INFAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   4512
      TabIndex        =   4
      Top             =   4824
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      Caption         =   "LOAD INFAR INI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1092
      Left            =   1992
      TabIndex        =   3
      Top             =   4824
      Visible         =   0   'False
      Width           =   2412
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "Make Path on Same Drive or if Network"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1416
      Left            =   72
      TabIndex        =   2
      Top             =   4512
      Width           =   1812
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
      Height          =   400
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11292
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT"
   End
   Begin VB.Menu VB_ME 
      Caption         =   "VB ME"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_EXPLORER 
      Caption         =   "EXPLORER SELECT"
   End
   Begin VB.Menu MNU_FORM_DIR_SOURCE 
      Caption         =   "FORM DIR SOURCE"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "IR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A1, ii

Dim FORMLOAD, NETOPT, GET_DRIVES_RESULT



Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function FindWindowDLL _
        Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As Long, _
        ByVal lpWindowName As Long) As Long
'Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long








'Dim FS

Private Sub Form_Load()

Call SET_VAR_FS

 'Explorer.exe /select,
ii = Clipboard.GetText

If Mid$(ii, 1, 2) = "\\" Then
'\\Linda-pc\d acer\VB6-EXE'S SYNC
NETOPT = True

End If

If Mid$(ii, 2, 2) <> ":\" And NETOPT = False Then
    ' MsgBox "DIR WAS NOT ROOT OF DRIVE -- TAKE CARE" ': End
    ii = "Z:\" + ii
End If




'ii = "D:\VB6\VB-NT\00_Send_To\Clipboard Dir Create on another drive"
   
    
    
    
    
    
On Error Resume Next

'If App.PrevInstance = True Then End

Text1 = ii
'Text2 = ii

'Text1 = ii

If Mid$(Text1, Len(Text1), 1) <> "\" Then
    Text1 = Text1 + "\"
End If


Dim CLIP
CLIP = Clipboard.GetText
If Mid(CLIP, 2, 2) = ":\" Then CLIP = Mid(CLIP, 3)

If Mid(CLIP, 1, 1) = "\" Then CLIP = Mid(CLIP, 2)

Text3.Text = CurDir + "\" + CLIP


'Text2 = Mid(Text2, 1, InStrRev(Text1, "\"))

'Text2.Text = Chr(Index + 65) + Mid(Text2.Text, 2)

TOP_MOST_BOX = Label1.Top + Label1.Height + 10
Text1.Top = TOP_MOST_BOX

TOP_MOST_BOX = Text1.Top + Text1.Height + 10
Label9.Top = TOP_MOST_BOX

TOP_MOST_BOX = Label9.Top + Label9.Height + 10
Text3.Top = TOP_MOST_BOX

'Text1 = Mid(Text1, 1, InStrRev(Text1, "\"))

'If Mid$(Text1, Len(Text1), 1) <> "\" Then
'    Text1 = Text1 + "\"
'End If


FORMLOAD = True

'Call Label3_Click

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

TOP_MOST_BOX = Text3.Top + Text3.Height + 15


XT = (Text1.Width - 16) / 26
'XT = XT
XC = 65
XD = LabDR(0).Left = 0
For Each Control In LabDR
    DoEvents
    'CONTROL.INDEX
    Control.Top = TOP_MOST_BOX
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
    Control.Top = TOP_MOST_BOX

    Control.Height = Control.Height + 40
    XxTopH = Control.Top + Control.Height
    Control.Left = XD
    Control.Width = XT
    XD = XD + XT
    If XC > 90 Then Exit For
    XC = XC + 1
    Control.Visible = True
Next

GET_DRIVES
For Each Control In LabDR
    
'    If Control.Index Mod 2 = 0 Then
'        Control.BackColor = &HC0FFC0
'    End If
    
            
    If InStr("*" + LCase(GET_DRIVES_RESULT), LCase(Mid(Control.Caption, 2, 1))) = 0 Then
        Control.BackColor = Label8.BackColor
    End If


    

Next

Text2.Width = Label1.Width
Text3.Width = Label1.Width
Label9.Width = Label1.Width


TOP_MOST_BOX = Label9.Top + Label9.Height + 15

Text2.Top = TOP_MOST_BOX
Text2.Left = 0
Text2.Width = Text1.Width
Label2.Left = 0
Label2.Top = XxTopH

Me.WindowState = vbNormal
Me.Height = 4850
Me.Height = Label2.Height + Label2.Top + 800
Me.Width = Text1.Width + 280


'INFAR.Show
'
'INFAR.SetFocus

'IR.WindowState = vbMinimized

End Sub

Private Sub SET_VAR_FS()

Set FS = CreateObject("Scripting.FileSystemObject")

End Sub

Function GET_DRIVES()

Dim DC
Set DC = FS.Drives
For Each D In DC
  s = s & D.DriveLetter
  If D.DriveType = 3 Then
    'Stop
    'n = d.ShareName
  ElseIf D.IsReady Then
    N = N + D.DriveLetter 'd.VolumeName
  End If
  's = s & n & "<BR>"
Next

GET_DRIVES = N
GET_DRIVES_RESULT = N

End Function

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub LabDR_Click(Index As Integer)


'ii = Chr(Index + 65) + Mid(ii, 2)

'Text2 = Mid(Text2, 1, InStrRev(Text1, "\"))

Text1.Text = Chr(Index + 65) + Mid(Text1.Text, 2)
'Text2.Text = Chr(Index + 65) + Mid(Text2.Text, 2)


If Len(Text1) = 3 Then MsgBox "NOT ON A ROOT": Exit Sub

Call Text1_KeyDownZZ(0, 0)

End Sub

Private Sub Label1_Click()

ii = Clipboard.GetText
If Mid$(ii, 2, 2) <> ":\" And NETOPT = False Then MsgBox "DIR WAS ROOT OF DRIVE -- END": End

'ii = "D:\VB6\VB-NT\00_Send_To\Clipboard Dir Create on another drive"
    On Error Resume Next
Text1 = ii
If Mid$(Text1, Len(Text1), 1) <> "\" Then
    Text1 = Text1 + "\"
End If


End Sub

Private Sub Label2_Click()

'Me.WindowState = vbMinimized
'
'Shell "Explorer.exe /select, " + Text1, vbNormalFocus


Call MAKE_THE_PATH


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

Private Sub Label9_Click()

Text1.Text = Text3.Text
Call MAKE_THE_PATH

End Sub

Private Sub MNU_EXIT_Click()
End
End Sub

Private Sub MNU_EXPLORER_Click()

'Me.WindowState = vbMinimized

Shell "Explorer.exe /select, " + Text1, vbNormalFocus


End Sub

Private Sub MNU_FORM_DIR_SOURCE_Click()

Call INFAR.Form_Load


End Sub

Private Sub Text1_KeyDownZZ(KeyCode As Integer, Shift As Integer)

'If KeyCode = 13 Then
'    a = a
'Exit Sub


Call MAKE_THE_PATH

End Sub


Sub MAKE_THE_PATH()

Dim i As Long
If Mid$(Text1, 2, 2) <> ":\" And NETOPT = False Then MsgBox "DIR NOT ROOT OF DRIVE -- TAKE CARE - END": End
    
    
    If Mid(Text1, Len(Text1)) = "\" Then Text1 = Mid(Text1, 1, Len(Text1) - 1)
    
    On Error Resume Next
    If Dir(Text1, vbDirectory) <> "" Then
        YY = Err.Description
        UU = Err.Number
        
        
        If UU > 0 Then
            MsgBox "PROBLEM WITH FOLDER" + vbCrLf + Text1 + vbCrLf + "ERROR DESCRIPTION = " + YY + vbCrLf + "ERROR NUMBER =" + Str(UU)
        Else
            MsgBox "FOLDER ALREADY EXISTS" + vbCrLf + Text1
        
        End If
        
        Exit Sub
    
    End If
    
    
    If UU = 0 Then
    
    
        On Error Resume Next
        R = 4
        Do
                    
            T = InStr(R + 1, Text1, "\")
            dd = ""
            If T > 0 Then
                dd = Mid$(Text1, 1, T - 1)
            Else
                dd = Text1 'MUST BE FIRST DIR LEVEL PATH
            
            End If
            If NETOPT = True Then PATHLEVEL = PATHLEVEL + 1
            'Err.Clear
            'If T > 0 Then Err.Clear: MkDir dd
            If dd <> "" And NETOPT = False Then Err.Clear: MkDir dd
            If dd <> "" And NETOPT = True And PATHLEVEL > 2 Then Err.Clear: MkDir dd
            
            R = T
            
            YY = Err.Description
            UU = Err.Number
        
        Loop Until T = 0
'End
    End If

If UU = 0 Then
'    Shell "Explorer.exe /select, " + Text1, vbNormalFocus
    
'    Do
'    If FINDWINDOW("CabinetWClass", vbNullString, "Explorer.exe") = 0 Then SLEEP 100
'    Loop Until FINDWINDOW(vbNullString, "EXPLORER")
'
    cmdLine$ = "HWND Not Available"
    cmdLine$ = "c:\windows\Explorer.exe /select, " + Text1
    i = ExecCmd(cmdLine$)
    
    
    'SEXY EXE SHELL EXECUTE CUTE
    'EASYIER IF MAYBE GOT PROCESS FROM SHELL COMMAND EXECUTE
    'AND CONVERT PROCESS TO HWND
    'IF IT LET ME
    '2016-01-02 SAT
    If cmdLine$ = "HWND Not Available" Then
    'OR I=0
        TEN = 0
        Do
            i = OneForeGroundWnd("CabinetWClass")
        
            'If IsWindowVisible(i) = True Then Exit Do
            If i = 0 Then Sleep 100 ': Stop
            TEN = TEN + 1
        Loop Until i > 0 Or TEN > 100
        
        If TEN > 100 Then MsgBox "SORRY ABOUT THAT DELAY FIND WINDOW EXPLORER" + vbCrLf + "PATH" + vbCrLf + Text1 + vbCrLf + "SUCCESSFULLY CREATED" + vbCrLf + "EXPLORER /select, " + Text1, vbMsgBoxSetForeground, vbMsgBoxSetForeground
    
        Sleep 500
        Me.SetFocus
        DoEvents
        
    End If
    'HAPPENS TOO QUICK NOW
    'TRY DELAY
    
    'WONT LET MESSAGEBOX STEAL FOCUS
    'ANSWER BRING FORM FORWARD
    'TUFF ANSWER LATE NIGHT 7:30AM
    '1.SetFocus
    '2.ALWAYS ONTOP
    
    'shell and wait to show
'\\Linda-pc\d acer\VB6-EXE'S SYNC
    
    If 1 = 2 Then
        If TEN <= 100 Then MsgBox "PATH" + vbCrLf + Text1 + vbCrLf + "SUCCESSFULLY CREATED" + vbCrLf + "EXPLORER /select, " + Text1, vbMsgBoxSetForeground
    End If

'    Shell "Explorer.exe /select, " + Text1, vbNormalFocus
'    End If
    


End


Else
    MsgBox "PATH" + vbCrLf + Text1 + vbCrLf + "PROBLEM CREATING" + vbCrLf + "ERROR DESCRIPTION = " + YY + vbCrLf + "ERROR NUMBER =" + Str(UU)

End If

'End If


End Sub

Private Sub MNU_VB_FOLDER_Click()

Shell "EXPLORER /SELECT, " + App.Path + "\" + App.EXEName + ".VBP", vbMaximizedFocus

End Sub

Private Sub Text1_Change()

'Text1.Text = Chr(Index + 65) + Mid(Text1.Text, 2)
'Text2.Text = Text1.Text
'
''IF MID()
'Text2 = Mid(Text2, 1, InStrRev(Text1, "\"))

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Text1_KeyDownZZ(0, 0)
End If
If KeyAscii = 27 Then Unload Me

End Sub

Private Sub Text3_Click()

Text1.Text = Text3.Text
Call MAKE_THE_PATH

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

    Dim CODER_VBP_FILE_NAME_2
    Dim VB_1, VB_2, VB_3

    VB_1 = "C:\PROGRAM FILES\Microsoft Visual Studio\VB98\VB6.EXE"
    VB_2 = "C:\PROGRAM FILES (X86)\Microsoft Visual Studio\VB98\VB6.EXE"
    If Dir(VB_1) <> "" Then VB_3 = VB_1
    If Dir(VB_2) <> "" Then VB_3 = VB_2
    '------------------------------------------------------
    'ADD MICROSFOT SCRIPTING RUNTIME FOR HERE IN REFERENCES
    '------------------------------------------------------
    Dim objShell
    Set objShell = CreateObject("Wscript.Shell")
    CODER_VBP_FILE_NAME_2 = App.Path + "\" + App.EXEName + ".VBP"
    objShell.Run """" + VB_3 + """ """ + CODER_VBP_FILE_NAME_2 + """", 1, False
    Set objShell = Nothing
    
    EXIT_TRUE = True
    Unload Me
    Exit Sub

End Sub
