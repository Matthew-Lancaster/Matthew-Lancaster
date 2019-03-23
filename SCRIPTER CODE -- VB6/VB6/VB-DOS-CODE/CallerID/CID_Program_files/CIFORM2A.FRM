VERSION 5.00
Begin VB.Form Ciform2a 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ciform2a"
   ClientHeight    =   5820
   ClientLeft      =   10050
   ClientTop       =   300
   ClientWidth     =   9585
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5820
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   3240
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Text            =   "Speed Search:"
      Top             =   240
      Width           =   1560
   End
   Begin VB.CommandButton findbutton 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "&F3 = Find Again"
      Height          =   240
      Left            =   5640
      TabIndex        =   4
      Top             =   0
      Width           =   1905
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Date & Time :"
      Top             =   0
      Width           =   1560
   End
   Begin VB.TextBox datetime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1560
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "MON 00/JAN/1900 12:00:00 PM"
      Top             =   0
      Width           =   3240
   End
   Begin VB.TextBox BUSY 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   4800
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   " BUSY.    "
      Top             =   0
      Width           =   840
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5640
      TabIndex        =   8
      Text            =   "Rec. No."
      Top             =   240
      Width           =   960
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H007B5B37&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5070
      Left            =   0
      TabIndex        =   1
      Top             =   825
      Width           =   9600
   End
   Begin VB.TextBox recpostext 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6600
      TabIndex        =   9
      Top             =   240
      Width           =   945
   End
   Begin VB.CommandButton exitcommand 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "E&xitbutton"
      Height          =   240
      Left            =   7560
      TabIndex        =   3
      Top             =   0
      Width           =   2025
   End
   Begin VB.CommandButton editcommand 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "&Editbutton"
      Height          =   240
      Left            =   7560
      TabIndex        =   11
      Top             =   240
      Width           =   2025
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "...Main telephone entries storage Area 2..."
      Height          =   345
      Left            =   0
      TabIndex        =   10
      Top             =   480
      Width           =   9600
   End
End
Attribute VB_Name = "Ciform2a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeyPass As Integer

Private Sub Command1_Click()
Ciform2a.Text2.SetFocus
End Sub

Private Sub Editcommand_Click()
BUSY.Visible = -1
Which = 2.1
Ciform2a.Visible = 0
Ciform2a.Enabled = 0
CiEDIT.Visible = -1
CiEDIT.Enabled = -1
DoIt2 = 1
CiEDIT.Speed_search_text.SetFocus
BUSY.Visible = 0
End Sub

Private Sub Exitcommand_Click()
Ciform2a.Visible = 0
Ciform2a.Enabled = 0
Ciform4.Visible = -1
Ciform4.Enabled = -1
Ciform4.Text2.SetFocus
End Sub

Private Sub Findbutton_click()
Ciform2a.BUSY.Visible = -1

For r = Ciform2a.List2.ListIndex + 1 To Ciform2a.List2.ListCount - 1
    w = InStr(UCase$(Ciform2a.List2.List(r)), Text2.Text)
    If w > 0 Then Exit For
Next

If w = 0 Then
    For r = 1 To Ciform2a.List2.ListCount - 1
        w = InStr(UCase$(Ciform2a.List2.List(r)), Text2.Text)
        If w > 0 Then Exit For
    Next
End If

Ciform2a.BUSY.Visible = 0

If w > 0 And Text2.Text <> "" Then
    Ciform2a.List2.ListIndex = r
End If

Call List2_keyup(0, 0)
End Sub

Private Sub List2_Click()

recpostext.Text = Str(Ciform2a.List2.ListIndex)

End Sub

Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then Call Findbutton_click
End Sub

Private Sub List2_KeyPress(keyascii As Integer)
KeyPass = keyascii
End Sub

Private Sub List2_keyup(KeyCode As Integer, Shift As Integer)

recpostext.Text = Str(Ciform2a.List2.ListIndex)

If KeyPass = 27 Then Call Exitcommand_Click: KeyPass = 0

If KeyPass > 0 Then Call Text2_Change: Text2.SetFocus

End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

recpostext.Text = Str(Ciform2.List2.ListIndex)

End Sub

Private Sub Text2_Change()

If Trooper = 0 Then
    If KeyPass = 8 Then KeyPass = -1: Text2.Text = Mid$(Text2.Text, 1, Len(Text2.Text) - 1)
    If KeyPass = 127 Then KeyPass = -1
    If KeyPass > 0 Then l = KeyPass: KeyPass = -1: Text2.Text = Text2.Text + Chr$(l)

    If KeyPass = 0 Then
        Text2.Text = UCase$(Text2.Text)
        Text2.SelStart = Len(Text2.Text)
        Ciform2.BUSY.Visible = -1
        
        For r = 0 To Ciform2a.List2.ListCount - 1
            w = InStr(UCase$(Ciform2a.List2.List(r)), Text2.Text)
            If w > 0 Then Exit For
        Next

        Ciform2a.BUSY.Visible = 0
        
        If w = 0 And Text2.Text <> "" Then
            KeyPass = -1
            Text2.Text = Mid$(Text2.Text, 1, Len(Text2.Text) - 1)
            Text2.SelStart = Len(Text2.Text)
        End If

        If w > 0 And Text2.Text <> "" Then
            Ciform2a.List2.ListIndex = r
        End If
    End If

Trooper = 1
Ciform1.Text2.Text = Text2.Text
Ciform1.Text2.SelStart = Len(Text2.Text)
Trooper = 0
Call List2_keyup(0, 0)
KeyPass = 0
End If
End Sub

Private Sub Text2_GotFocus()
Call Text2_Change
Call List2_keyup(0, 0)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then Call Exitcommand_Click
If KeyCode = 114 Then Call Findbutton_click

If KeyCode = 34 Or KeyCode = 33 Then
    Ciform2a.List2.ListIndex = Ciform2a.List2.ListIndex
    Ciform2a.List2.SetFocus
End If

If KeyCode = 40 Then
    If Ciform2a.List2.ListIndex < Ciform2a.List2.ListCount - 1 Then
        Ciform2a.List2.ListIndex = Ciform2a.List2.ListIndex + 1
    End If
Ciform2a.List2.SetFocus
End If

If KeyCode = 38 Then
    If Ciform2a.List2.ListIndex > 0 Then
        Ciform2a.List2.ListIndex = Ciform2a.List2.ListIndex - 1
    End If
    Ciform2a.List2.SetFocus
End If
End Sub

