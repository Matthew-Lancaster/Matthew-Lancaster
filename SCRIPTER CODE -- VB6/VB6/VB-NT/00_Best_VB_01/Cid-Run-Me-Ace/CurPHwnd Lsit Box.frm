VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Frm_ProHwndList 
   Caption         =   "Process And Hwnd List Plus Some"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTB2 
      Height          =   5040
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   8890
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"CurPHwnd Lsit Box.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Frm_ProHwndList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
'If Cheese = False Then
 '   Cancel = False
 '   Exit Sub

'End If
Frm_ProHwndList.Hide
'Me.WindowState = 0
    
Cancel = True
End Sub

Private Sub RTB2_Change()

End Sub
