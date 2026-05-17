VERSION 5.00
Begin VB.Form MP3_PLAYED 
   Caption         =   "PLAYED"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6165
      Top             =   1485
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   1695
      Top             =   1410
   End
   Begin VB.Timer Timer1 
      Interval        =   50000
      Left            =   4770
      Top             =   2100
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "MP3_PLAYED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OL
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long  'MODULE 1141
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long


Private Sub Form_Load()

'Me.Show


Me.Hide
Unload Me

'Me.WindowState = 0
End Sub

Private Sub Form_Resize()

On Error Resume Next
Me.Height = List1.Height + 470
Me.Width = List1.Width + 150

Me.Left = 4200
Me.Top = 3000
End Sub

Private Sub List1_Click()
'mp3_played.list1
Me.WindowState = 0

End Sub

Private Sub Timer1_Timer()

Me.WindowState = 1

End Sub

Private Sub Timer2_Timer()

If List1.ListCount <> OL Then
    List1.ListIndex = List1.ListCount - 1
'   mp3_played.list1
    Me.WindowState = 0
End If
OL = List1.ListCount
End Sub


Private Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim l As Long
   Dim S As String
   l = GetWindowTextLength(hwnd)
   S = Space(l + 1)
   GetWindowText hwnd, S, l + 1
   GetWindowTitle = Left$(S, l)
End Function


Private Sub Timer3_Timer()

TF2 = FindWindow("Winamp v1.x", vbNullString)

Label9 = ttmd$

If TF2 = 0 Then Exit Sub

Rfg$ = GetWindowTitle(TF2)
rfh1 = InStrRev(Rfg$, " - ")
If rfh1 > 0 Then Rfg$ = Mid$(Rfg$, 1, rfh1 - 1)

If TF2 > 0 Then Label11 = Rfg$



tt = Format$(Now, "HH:MM:SSa/p")
tt = tt + " - " + Label11

MP3_PLAYED.List1.AddItem tt



End Sub

End Sub
