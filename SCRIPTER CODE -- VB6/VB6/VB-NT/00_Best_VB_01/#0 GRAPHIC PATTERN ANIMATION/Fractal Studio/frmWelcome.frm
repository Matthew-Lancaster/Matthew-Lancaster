VERSION 5.00
Begin VB.Form frmWelcome 
   Caption         =   "Welcome to Fractal Studio"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4860
   ControlBox      =   0   'False
   Icon            =   "frmWelcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInstructions 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton btnProceed 
      Caption         =   "Proceed !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnProceed_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim sIns As String

sIns = "Welcome to Fractal Studio Version " & App.Major & "." & App.Minor & _
        App.Revision & " !" & vbCrLf & vbCrLf & _
        "Since this is the first time you use the application, you may want to " & _
        "read the Readme-file first, as the User Interface of this program " & _
        "isn't very intuitive yet.." & vbCrLf & _
        "The file is included in the .zip." & vbCrLf & vbCrLf & _
        "Basics:" & vbCrLf & _
        "Mark an area of the picture to zoom in, drag with the right mouse button " & _
        "to move the viewpoint." & vbCrLf & _
        "Saving and loading of parameter files has now been implemented. " & vbCrLf & _
        "Click the debug menu item to register the file extension FS1." & vbCrLf & vbCrLf & _
        "New features always introduces new bugs. If you find any, I'd be happy " & _
        "if you reported them to me (erlinga@stud.ntnu.no), so I can fix them in the " & _
        "next update of the program." & vbCrLf & vbCrLf & _
        "Happy exploring!" & vbCrLf & vbCrLf & _
        "Erling"

txtInstructions.Text = sIns

End Sub
