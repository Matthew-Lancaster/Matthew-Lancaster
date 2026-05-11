VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Email Spam Future Google"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7920
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   7920
   Begin VB.Label Lb2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "//"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   690
      Left            =   75
      TabIndex        =   1
      Top             =   1125
      Width           =   7815
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "\\"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   690
      Left            =   75
      TabIndex        =   0
      Top             =   225
      Width           =   7815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Form_Activate()
Call Form1.xy
End Sub

