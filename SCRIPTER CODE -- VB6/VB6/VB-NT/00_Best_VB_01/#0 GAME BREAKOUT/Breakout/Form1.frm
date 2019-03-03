VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2865
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   2865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form1
' DateTime  : 4/2/2005 12:25
' Author    : Jason   (noi_max@hotmail.com)
' Purpose   : DX Breakout - opening form
'---------------------------------------------------------------------------------------

Option Explicit

'Constants for screen resolution choices
Private Const Choice2 As String = "800x600x16"
Private Const Choice3 As String = "800x600x24"
Private Const Choice4 As String = "800x600x32"

'Set this to run/skip resoultion choices
Private Const DEV_SKIP As Boolean = True

Private Sub Form_Load()
   
   If DEV_SKIP Then
   
      'Skip resolution choices and go right into the main loop
      DX_WIDTH = 800
      DX_HEIGHT = 600
      'DX_BPP = 16
      DX_BPP = 32
      RunGame
      Exit Sub
      
   End If
   
   Label1.Caption = "Breakout DX." & vbCrLf & _
                    "Please choose a color depth to start."
                    
   'Add screen resolution choices to the combo box
   Combo1.AddItem Choice4
   'Combo1.AddItem Choice2
   Combo1.AddItem Choice3
   Combo1.AddItem Choice2
   'Combo1.AddItem Choice4

   
End Sub

Private Sub Combo1_Click()

   'Set the DX_WIDTH nd DX_HEIGHT variables based on the
   'choice in the combobox
   Dim temp() As String
   
   temp = Split(Combo1.Text, "x")
   
   'Debug.Print Text
   
   DX_WIDTH = CLng(temp(0))
   DX_HEIGHT = CLng(temp(1))
   DX_BPP = CLng(temp(2))
   
   cmdStart.Enabled = True
   
End Sub

Public Sub cmdStart_Click()

   RunGame           'Start the game :)
   
End Sub


