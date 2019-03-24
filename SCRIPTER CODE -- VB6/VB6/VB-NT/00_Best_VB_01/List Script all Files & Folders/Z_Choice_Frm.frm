VERSION 5.00
Begin VB.Form Z_Choice_Frm 
   BackColor       =   &H0000C0C0&
   Caption         =   "Form1Option 1"
   ClientHeight    =   2520
   ClientLeft      =   180
   ClientTop       =   820
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Option 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   12735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Option 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   825
      Width           =   12735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Option 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12735
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB ME"
   End
End
Attribute VB_Name = "Z_Choice_Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Setting_Done

Private Sub Form_Load()

Setting_Done = False

Me.Caption = "Sort_AnyThing - CRC DUPE - GOOGLE IMAGE"
Me.Show
Me.WindowState = vbNormal

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Want to Abort
If Setting_Done = True Then Exit Sub

Dim Form As Form
For Each Form In Forms
    Unload Form
Next

End Sub

Private Sub Label1_Click()
'LABEL1.CAPTION
If Label1.Enabled = False Then
MsgBox "Not This Option"
Exit Sub
End If
Setting_Done = True
Call z_MENU_Form1.Clip_Choice(1)
Unload Me
End Sub

Private Sub Label2_Click()
If Label1.Enabled = False Then
    MsgBox "Not This Option"
    Exit Sub
End If

Setting_Done = True
Call z_MENU_Form1.Clip_Choice(2)
Unload Me

End Sub

Private Sub Label3_Click()
If Label1.Enabled = False Then
MsgBox "Not This Option"
Exit Sub
End If
Call z_MENU_Form1.Clip_Choice(3)
Setting_Done = True
Unload Me

End Sub

Private Sub MNU_VB_Click()
Shell """" + GetSpecialfolder(38) + "\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
End
End Sub
