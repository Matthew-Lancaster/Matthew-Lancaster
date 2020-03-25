VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form LSForm2 
   Caption         =   "Form2"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1365
      Left            =   450
      TabIndex        =   2
      Top             =   1635
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   2408
      _Version        =   393217
      OLEDropMode     =   1
      TextRTF         =   $"BinaryHack.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   480
      Left            =   705
      TabIndex        =   1
      Top             =   975
      Width           =   1725
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   480
      Left            =   630
      TabIndex        =   0
      Top             =   255
      Width           =   1785
   End
End
Attribute VB_Name = "LSForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Private Sub Form_Load()
        Command1.Caption = "Open"
        Command2.Caption = "Save"
    End Sub

    Private Sub Command1_Click()
        ' Open
        r = Shell("Explorer.exe", vbNormalFocus)
    End Sub

    Private Sub Command2_Click()
        ' Save
        If Len(RichTextBox1.fileName) Then
            RichTextBox1.SaveFile RichTextBox1.fileName, rtfText
        End If
    End Sub

    Private Sub RichTextBox1_OLEDragDrop(Data As RichTextLib.DataObject, _
        Effect As Long, Button As Integer, Shift As Integer, _
        x As Single, y As Single)

        ' Drag ANY file from the Explorer, or an "Explorer View"
        If Data.GetFormat(vbCFFiles) Then
            RichTextBox1.LoadFile Data.Files(1), rtfText
        End If
    End Sub




