VERSION 5.00
Begin VB.Form Ciform4 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ciform4"
   ClientHeight    =   5805
   ClientLeft      =   10050
   ClientTop       =   300
   ClientWidth     =   9600
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
   ScaleHeight     =   5805
   ScaleWidth      =   9600
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
      Locked          =   -1  'True
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
      Locked          =   -1  'True
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
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "MON 00/JAN/1900 12:00:00 PM"
      Top             =   0
      Width           =   3240
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "...High scores entries..."
      Height          =   345
      Left            =   0
      TabIndex        =   10
      Top             =   480
      Width           =   9600
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Rec. No."
      Top             =   240
      Width           =   960
   End
   Begin VB.TextBox recpostext 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   240
      Width           =   930
   End
   Begin VB.CommandButton exitcommand 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "E&xitbutton"
      Height          =   240
      Left            =   7545
      TabIndex        =   3
      Top             =   0
      Width           =   2040
   End
   Begin VB.CommandButton editCommand 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "&Editbutton"
      Height          =   240
      Left            =   7530
      TabIndex        =   11
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox BUSY 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   " BUSY.    "
      Top             =   0
      Width           =   840
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
      Left            =   15
      TabIndex        =   1
      Top             =   825
      Width           =   9600
   End
End
Attribute VB_Name = "Ciform4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeyPass As Integer

Private Sub Command1_Click()
Ciform4.Text2.SetFocus
End Sub

Private Sub Exitcommand_Click()
Ciform4.Visible = 0
Ciform4.Enabled = 0

XaGG1 = False

End Sub

Private Sub Findbutton_click()
Ciform4.BUSY.Visible = -1

For r = Ciform4.List2.ListIndex + 1 To Ciform4.List2.ListCount - 1
    w = InStr(UCase$(Ciform4.List2.List(r)), Text2.Text)
    If w > 0 Then Exit For
Next

If w = 0 Then
    For r = 1 To Ciform4.List2.ListCount - 1
        w = InStr(UCase$(Ciform4.List2.List(r)), Text2.Text)
        If w > 0 Then Exit For
    Next
End If

Ciform4.BUSY.Visible = 0

If w > 0 And Text2.Text <> "" Then
    Ciform4.List2.ListIndex = r
End If

Call List2_keyup(0, 0)

End Sub

Private Sub List2_Click()

recpostext.Text = Str(Ciform4.List2.ListIndex)

End Sub

Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then Call Findbutton_click
End Sub

Private Sub List2_KeyPress(keyascii As Integer)
KeyPass = keyascii
End Sub

Private Sub List2_keyup(KeyCode As Integer, Shift As Integer)

recpostext.Text = Str(Ciform4.List2.ListIndex)

If KeyPass = 27 Then Call Exitcommand_Click: KeyPass = 0

If KeyPass > 0 Then Call Text2_Change: Text2.SetFocus

End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

recpostext.Text = Str(Ciform4.List2.ListIndex)

End Sub

Private Sub Text2_Change()

If Trooper = 0 Then
    If KeyCode = 27 Then Call Exitcommand_Click
    If KeyPass = 8 Then KeyPass = -1: Text2.Text = Mid$(Text2.Text, 1, Len(Text2.Text) - 1)
    If KeyPass = 127 Then KeyPass = -1
    If KeyPass > 0 Then l = KeyPass: KeyPass = -1: Text2.Text = Text2.Text + Chr$(l)

    If KeyPass = 0 Then
        Text2.Text = UCase$(Text2.Text)
        Text2.SelStart = Len(Text2.Text)
        Ciform4.BUSY.Visible = -1
        
        For r = 0 To Ciform4.List2.ListCount - 1
            w = InStr(UCase$(Ciform4.List2.List(r)), Text2.Text)
            If w > 0 Then Exit For
        Next

        Ciform4.BUSY.Visible = 0

        If w = 0 And Text2.Text <> "" Then
            KeyPass = -1
            Text2.Text = Mid$(Text2.Text, 1, Len(Text2.Text) - 1)
            Text2.SelStart = Len(Text2.Text)
        End If

        If w > 0 And Text2.Text <> "" Then
            Ciform4.List2.ListIndex = r
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
    Ciform4.List2.ListIndex = Ciform4.List2.ListIndex
    Ciform4.List2.SetFocus
End If

If KeyCode = 40 Then
    If Ciform4.List2.ListIndex < Ciform4.List2.ListCount - 1 Then
        Ciform4.List2.ListIndex = Ciform4.List2.ListIndex + 1
    End If

    Ciform4.List2.SetFocus
End If

If KeyCode = 38 Then
    If Ciform4.List2.ListIndex > 0 Then
        Ciform4.List2.ListIndex = Ciform4.List2.ListIndex - 1
    End If
    Ciform4.List2.SetFocus
End If

End Sub

