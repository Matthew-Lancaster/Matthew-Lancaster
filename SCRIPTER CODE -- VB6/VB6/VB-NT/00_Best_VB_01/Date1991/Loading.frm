VERSION 5.00
Begin VB.Form loading 
   BackColor       =   &H0000FFFF&
   Caption         =   "Date1991 Loading"
   ClientHeight    =   765
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   3345
   Icon            =   "loading.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   765
   ScaleWidth      =   3345
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2025
      Top             =   450
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Program Loading"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   3345
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Program Loading"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3345
   End
End
Attribute VB_Name = "loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xcds

Private Sub Form_Load()
commandstring$ = Command$
End Sub

Private Sub Timer1_Timer()

If xcds = 0 Then xcds = 1: Load Form1

Label1.Caption = "Program Loading"
If loadermagic2 > 0 Then
Label2.Caption = Format$((loadermagic / loadermagic2) * 100, "0.000") + "%"
End If
If loadprogram = 1 Then
Unload loading
If commandstring$ = "" Then Form1.Show
End If

End Sub
