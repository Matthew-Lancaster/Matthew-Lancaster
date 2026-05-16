VERSION 5.00
Object = "{54B5C0BF-F9D3-11D4-9393-444553540000}#1.0#0"; "SCROLLER.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scroller Example"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   3945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Upwards Scroll 2"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Upwards Scroll 1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
   End
   Begin TextScroller.Scroller Scroller1 
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4260
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Scroll 2"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scroll 1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Format for using the ocx:
'scroller1.DoScroll <text to scroll>, <speed (pixels/sec)>, <pause (ms)>, <overirde - if text is already scrolling should it stop it?>, <put a 1 here if you want an upwards scroll>

Private Sub Command1_Click()
    Scroller1.DoScroll "Scroll Type 1 [no override]", 2, 1000, False
End Sub

Private Sub Command2_Click()
    Scroller1.DoScroll "Scroll Type 2 [override]", 2, 1000, True
End Sub

Private Sub Command3_Click()
    Scroller1.DoScroll "Upwards Scroll", 2, 1000, False, 1
End Sub

Private Sub Command4_Click()
    Scroller1.DoScroll "Upwards Scroll [override]", 2, 1000, True, 1
End Sub

Private Sub Form_Load()
    Scroller1.BackColor = RGB(0, 0, 0)
    Scroller1.TextColor = RGB(100, 180, 100)
    Scroller1.DoScroll "Text Scroller 1.0 by Paul Blower [mercior@hotmail.com]", 1.5, 0, True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub
