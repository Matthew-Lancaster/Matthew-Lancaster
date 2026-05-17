VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmHelp 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "help"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   255
      Left            =   8280
      TabIndex        =   2
      Top             =   7560
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   13361
      _Version        =   393217
      BackColor       =   -2147483624
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      FileName        =   "C:\Program Files\Microsoft Visual Studio\VB98\QND Programs\ExtFindD\ExtendedFindHelp.rtf"
      TextRTF         =   $"frmHelp.frx":0000
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   360
      Picture         =   "frmHelp.frx":9689
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   5160
      Width           =   300
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHelp_Click()

  SaveHelpPosition
  Me.Hide

End Sub

Private Sub Form_Load()

  With Me
    .Left = GetSetting(AppDetails, "Settings", "HelpLeft", .Left)
    .Top = GetSetting(AppDetails, "Settings", "HelpTop", .Top)
    .Width = GetSetting(AppDetails, "Settings", "HelpWidth", .Width)
    .Height = GetSetting(AppDetails, "Settings", "HelpHeight", .Height)
    .Caption = "Help " & AppDetails
  End With 'Me
  Form_Resize

End Sub

Private Sub Form_Resize()

  Dim offset As Long

  offset = 100
  With RichTextBox1
    .Top = 0
    .Left = 0
    .Width = frmHelp.ScaleWidth
    .Height = frmHelp.ScaleHeight - cmdHelp.Height - offset
  End With
  With cmdHelp
    .Left = frmHelp.Width - .Width - offset
    .Top = RichTextBox1.Height + offset / 2
  End With
  SaveHelpPosition

End Sub

Private Sub Form_Unload(Cancel As Integer)

  SaveHelpPosition

End Sub

Private Sub SaveHelpPosition()

  With Me
    SaveSetting AppDetails, "Settings", "HelpLeft", .Left
    SaveSetting AppDetails, "Settings", "HelpTop", .Top
    SaveSetting AppDetails, "Settings", "HelpWidth", .Width
    SaveSetting AppDetails, "Settings", "HelpHeight", .Height
  End With 'Me

End Sub

':) Roja's VB Code Fixer V1.1.2 (7/07/2003 2:02:56 AM) 1 + 59 = 60 Lines Thanks Ulli for inspiration and lots of code.
