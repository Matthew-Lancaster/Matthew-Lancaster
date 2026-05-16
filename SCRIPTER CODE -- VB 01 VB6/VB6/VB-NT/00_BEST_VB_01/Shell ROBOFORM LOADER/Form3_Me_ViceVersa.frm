VERSION 5.00
Begin VB.Form Form3_Me_ViceVersa 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Me ViceVersa"
   ClientHeight    =   1152
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   9216
   LinkTopic       =   "Form2"
   ScaleHeight     =   1152
   ScaleWidth      =   9216
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   360
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Not Want - Yes Correct That Is OKAY Go Away"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   9240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Want"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   560
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9240
   End
End
Attribute VB_Name = "Form3_Me_ViceVersa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Dim COUNT_OUT

Private Sub Form_Activate()

    i = SetForegroundWindow(Me.hwnd)

End Sub

Private Sub Form_Load()

    'FORM3_ME_VICEVERSA
    Form1.FIND_VICE_VERSA
    Me.Show
    Me.SetFocus
    
    COUNT_OUT = 0

    Form1.MNU_VICE_VERSA_ME.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'UNLOAD_FORM_FLAG = True

End Sub

Private Sub Label1_Click()
    
    
    If MUST_BE_SONY_HX60V = True Then
        Shell VVCOMLINE + " ""D:\VICE_VERSA SCRIPT FILE\ViceVersa PHOTO CAMERA - ME.fsf""", vbNormalFocus
    End If


    If MUST_BE_FUJI_XP90_ = True Then
        Shell VVCOMLINE + " ""D:\VICE_VERSA SCRIPT FILE\ViceVersa PHOTO CAMERA - ME _ FUJI XP90.fsf""", vbNormalFocus
    End If


    Unload Me

End Sub

Private Sub Label2_Click()

    Unload Me

End Sub

Private Sub Timer1_Timer()
    COUNT_OUT = COUNT_OUT + 1
    EASY = 990
    If MUST_BE_SONY_HX60V = True Then
        Me.Caption = Trim(Str(EASY - COUNT_OUT)) + " __ ME __ VICEVERSA"
    End If
    
    If MUST_BE_FUJI_XP90_ = True Then
        Me.Caption = Trim(Str(EASY - COUNT_OUT)) + " __ ME _ FUJI XP90 __ VICEVERSA"
    End If
    
    If COUNT_OUT >= EASY Then Unload Me
End Sub
