VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form ATEST 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   624
      Top             =   752
   End
   Begin MSComctlLib.ProgressBar ProgressBarVOL1 
      Height          =   256
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3232
      _ExtentX        =   5689
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBarVOL2 
      Height          =   256
      Left            =   0
      TabIndex        =   1
      Top             =   304
      Width           =   3232
      _ExtentX        =   5689
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "ATEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ProgressBarVol1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub
