VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdButton 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   7350
      TabIndex        =   12
      Top             =   660
      Width           =   240
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Clea&r"
      BeginProperty Font 
         Name            =   "Verdana Ref"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   6600
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   135
      TabIndex        =   0
      Top             =   600
      Width           =   7455
      Begin VB.CommandButton cmdButton 
         Caption         =   "P&aste"
         BeginProperty Font 
            Name            =   "Verdana Ref"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Copy"
         BeginProperty Font 
            Name            =   "Verdana Ref"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Br&owse"
         BeginProperty Font 
            Name            =   "Verdana Ref"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   6480
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin MSComDlg.CommonDialog OpenDialog1 
         Left            =   6840
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox fileName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1590
         Width           =   7215
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Conver&t"
         BeginProperty Font 
            Name            =   "Verdana Ref"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6480
         TabIndex        =   2
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox fileName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   570
         Width           =   6255
      End
      Begin VB.Label Label1 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label1 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   90
      TabIndex        =   11
      Top             =   2880
      Width           =   7575
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   7215
   End
   Begin VB.Menu mnMenus 
      Caption         =   "Menus"
      Index           =   0
      Begin VB.Menu mnEntry 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mnEntry 
         Caption         =   ""
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -> All this Form's code by Improfane™
' -> http://www.improfane.pwp.blueyonder.co.uk/
' -> You may use this Form's code as long as you include CLEAR credit to 'improfane'


Option Explicit
' --> I am not using a resource file so I constant strings as this is just an example
Private Const bNote = "Enter a filename here"
Private Const LongToShort = "Convert Long to Short"
Private Const ShortToLong = "Convert Short to Long"
Private Const CopyLong = "Copy Long Filename"
Private Const CopyShort = "Copy Short Filename"
Private Const PasteLong = "Paste a Long Filename"
Private Const PasteShort = "Paste a Short Filename"
Private Sub cmdButton_Click(Index As Integer)
' --> Button handler - Handles all of the buttons
Select Case Index
 Case "0" ' <-- 'Convert' button
 PopupMenu mnMenus(0), vbPopupMenuRightAlign ' <-- Call Convert Menu
  
 ' --> Convert Long file name
Case "1"
    Dim X As String
    OpenDialog1.Filter = "Any file Type (*.*)|*.*|"
    OpenDialog1.ShowOpen ' <-- Display the 'Open File' dialog
           X = OpenDialog1.fileName ' <-- Long File Name that was selected in the open dialog
    If LenB(X) <> 0 Then ' <-- If user did not press 'Cancel' continue
        fileName(0).Text = X ' <-- Selected filename is put into the top box
            Form1.Tag = "0"
        mnEntry_Click (0) ' < -- Automatically convert the selected Long file name (so the user doesn't have to press 'Convert')
        X = vbNullString
        End If
Case "2" ' <-- 'Copy' button
    mnEntry(0).Caption = CopyShort
    mnEntry(1).Caption = CopyLong
     Form1.Tag = "1"
        PopupMenu mnMenus(0), vbPopupMenuRightAlign
        Form1.Tag = "0"
Case "3" ' <-- 'Paste' button
    mnEntry(0).Caption = PasteLong
    mnEntry(1).Caption = PasteShort
     Form1.Tag = "2"
        PopupMenu mnMenus(0), vbPopupMenuRightAlign
        Form1.Tag = "0"
Case "4" ' <-- 'Clear' button
    fileName(0).Text = bNote ' <-- Set the 'Long filename' box to the default instruction note
    fileName(1).Text = vbNullString ' <-- Empty 'Short filename' box
Case "5"
    MsgBox "Please read the file 'readme.txt' that should have came with this project -- It contains information about this program and how to use"
End Select
mnuSet
'Form1.Tag = "0"
End Sub

Private Sub fileName_Click(Index As Integer)
Select Case Index
Case "0"
    If fileName(0).Text = bNote Then fileName(0).Text = vbNullString ' <-- If the 'Long filename' box contains the instruction note - Erase it
End Select
End Sub

Private Sub fileName_Lostfocus(Index As Integer)
Select Case Index
Case "0"
    If LenB(fileName(0)) = 0 Then fileName(0).Text = bNote ' <-- If the 'Long filename' box is empty when it has lost focus - Set it to the instruction note
End Select
End Sub

Private Sub Form_Load()
mnuSet
mnMenus(0).Visible = False
Dim X As String
X = "Long-Filename and Short-Filename Converter"
Form1.title = X
Form1.Caption = X
X = vbNullString
' ---> Global title

' ---> Form/GUI Interface:
fileName(0).Text = bNote
Label1(0).Caption = "Long filename : -"
Label1(1).Caption = "Short filename : -"
Label2.Caption = "This Form's code by Improfane" & vbNewLine & "LongToShort Module's code by Microsoft"

' ---> Set Copy/Paste menu option to 0 so it does not copy or paste
Form1.Tag = "0"
End Sub
Sub mnuSet()
mnEntry(0).Caption = LongToShort
mnEntry(1).Caption = ShortToLong
End Sub

Private Sub mnEntry_Click(Index As Integer)
Dim Y As String
Dim oCP As String
oCP = Form1.Tag
    If oCP = "2" Then
        Y = Clipboard.GetText ' <-- Grab text from clipboard
            If LenB(Y) = 0 Then ' <-- If the clipboard text contains nothing then dont continue
                oCP = "0"
                Exit Sub
            End If
    End If
'Clipboard.SetText fileName(1).Text ' <-- Check if short filename box is empty and if not copy the text in it
Select Case Index
' <-- If clipboard is not empty paste it into the other file name box
    Case 0 ' <-- Short FileName
            If oCP = "1" Then ' <-- If 'Copy' is on then copy the 'Short Filename' to clipboard
             Clipboard.Clear: Clipboard.SetText fileName(1).Text
                Exit Sub
             ElseIf oCP = "2" Then fileName(0).Text = Y ' <-- If 'Paste' is on then paste the contents of the clipboard to the 'Long Filename' box
            End If
            fileName(1) = GetShortName(fileName(0))
    Case 1 ' <-- Long FileName
            If oCP = "1" Then ' <-- If 'Copy' is on then copy the 'Long Filename' to clipboard
             Clipboard.Clear: Clipboard.SetText fileName(0).Text
                Exit Sub
             ElseIf oCP = "2" Then fileName(1).Text = Y ' <-- If 'Paste' is on then paste the contents of the clipboard to the 'Short Filename' box
            End If
            fileName(0) = GetLongName(fileName(1))
End Select
 Y = vbNullString
 
 ' Test Area
 ' --> I use MSGBOX to check and test things, often
'MsgBox (Form1.Tag)
'MsgBox (cmdButton(2).Tag)

End Sub

Private Sub Label2_Click()
Shell ("explorer ""http://www.improfane.pwp.blueyonder.co.uk/""") ' <-- Opens my website in default web browser
' <-- Completely irrelevant to this program, My site, for future reference
End Sub
