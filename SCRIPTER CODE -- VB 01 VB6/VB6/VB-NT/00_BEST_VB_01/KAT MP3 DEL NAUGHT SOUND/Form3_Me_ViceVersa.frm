VERSION 5.00
Begin VB.Form Form3_Me_ViceVersa 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Me ViceVersa"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13665
   LinkTopic       =   "Form2"
   ScaleHeight     =   2160
   ScaleWidth      =   13665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Not Want - Yes Correct That Is OKAY Go Away"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   33.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   13680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Want"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   33.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13680
   End
End
Attribute VB_Name = "Form3_Me_ViceVersa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Me.Show
Me.SetFocus


End Sub

Private Sub Form_Unload(Cancel As Integer)

'UNLOAD_FORM_FLAG = True

End Sub

Private Sub Label1_Click()
    
    Shell VVCOMLINE + " ""D:\VICE_VERSA SCRIPT FILE\viceversa photo camera - me.fsf""", vbNormalFocus

    Unload Me

End Sub

Private Sub Label2_Click()

    Unload Me

End Sub
