VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000012&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1440
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Left            =   840
      Top             =   2640
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1320
      ScaleHeight     =   1185
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   840
      Width           =   4455
      Begin VB.Timer Timer3 
         Interval        =   1000
         Left            =   600
         Top             =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   825
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2115
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Press Any Key To Escape"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1845
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit
Dim DeltaX, DeltaY As Integer
Private Sub Timer1_Timer()
Label2.Caption = Time
Picture1.Move Picture1.Left + DeltaX, Picture1.Top + DeltaY
   If Picture1.Left < ScaleLeft Then DeltaX = 100
   If Picture1.Left + Picture1.Width > ScaleWidth + ScaleLeft Then
      DeltaX = -100
   End If
   If Picture1.Top < ScaleTop Then DeltaY = 100
   If Picture1.Top + Picture1.Height > ScaleHeight + ScaleTop Then
      DeltaY = -100
   End If
End Sub
Private Sub Form_Load()
Timer2.Enabled = True
   Timer1.Interval = 100
   DeltaX = 100
   DeltaY = 100
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub
Private Sub Timer2_Timer()
Label1 = ""
End Sub

Private Sub Timer3_Timer()
   'BackColor = QBColor(Rnd * 15)
   'ForeColor = QBColor(Rnd * 10)
   'Label2.BackColor = QBColor(Rnd * 15)
   Label2.ForeColor = QBColor(Rnd * 15)
End Sub
