VERSION 5.00
Begin VB.Form MP3_PLAYED 
   Caption         =   "MP3 PLAYED"
   ClientHeight    =   4380
   ClientLeft      =   165
   ClientTop       =   795
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
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3030
      Top             =   2160
   End
   Begin VB.Timer Timer4 
      Interval        =   500
      Left            =   2475
      Top             =   1560
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   1965
      Top             =   1350
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   1470
      Top             =   930
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   990
      Top             =   645
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
   End
   Begin VB.Menu MNU_VB_ME 
      Caption         =   "VB ME"
   End
End
Attribute VB_Name = "MP3_PLAYED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OL, OLAB11, TT, TTC, OTF3, YY, XTW
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long  'MODULE 1141
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long


Function AlwaysOnTop(ByVal hwnd As Long)  'Makes a form always on top
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Function
Function NotAlwaysOnTop(ByVal hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags
End Function


Private Sub Form_Load()

If App.PrevInstance = True Then End

Me.Show


Me.WindowState = 1
End Sub

Private Sub Form_Resize()

On Error Resume Next
List1.Height = 1300
Me.Height = List1.Height + 700
Me.Width = List1.Width + 210

Me.Left = 0
Me.Top = 450



End Sub

Private Sub List1_Click()
'mp3_played.list1

'Me.WindowState = 0


End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

YY = True

Timer5.Enabled = False
Timer5.Interval = 2000
Timer5.Enabled = True

End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

YY = True
Timer5.Enabled = False
Timer5.Interval = 2000
Timer5.Enabled = True

End Sub

Private Sub MNU_VB_ME_Click()
'If YY = False Then Exit Sub
EXEs = """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".VBP"""
Shell EXEs, vbNormalFocus
End
End Sub

Private Sub Timer1_Timer()

TTC = TTC + 1

If TTC < (60 * 5) Then Exit Sub
TTC = 0
'Me.WindowState = 1

End Sub

Private Sub Timer2_Timer()

If List1.ListCount <> OL Then
    List1.ListIndex = List1.ListCount - 1
'   mp3_played.list1
    'Me.WindowState = 0
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

If TF2 = 0 Then
    If Me.WindowState = 0 And XTW = 0 Then
        'Me.WindowState = 1
        XTW = 1

    End If
    Exit Sub
End If

XTW = 0

Rfg$ = GetWindowTitle(TF2)
rfh1 = InStrRev(Rfg$, " - ")
If rfh1 > 0 Then Rfg$ = Mid$(Rfg$, 1, rfh1 - 1)

If TF2 > 0 Then LABEL11 = Rfg$


TT = Format$(Now, "HH:MM:SSa/p")
TT = TT + " - " + LABEL11


If LABEL11 = OLAB11 Or Len(LABEL11) < 5 Then Exit Sub
TTC = 0

MP3_PLAYED.List1.AddItem TT

OLAB11 = LABEL11


End Sub

Private Sub Timer4_Timer()
TF3 = FindWindow(vbNullString, "DARK FORM")


GFW = GetForegroundWindow
If GFW <> TF3 And GFW <> Me.hwnd Then Exit Sub
If TF3 > 0 And GFW <> OGFW Then
    If GFW <> Me.hwnd Then
        'I = AlwaysOnTop(Me.hwnd)
        I = SetForegroundWindow(Me.hwnd)
        'Me.SetFocus
    End If
End If

OGFW = GFW

If TF3 = OTF3 Then Exit Sub

'Sleep 1000

    
If TF3 > 0 Then
    I = AlwaysOnTop(Me.hwnd)
    I = SetForegroundWindow(Me.hwnd)
    'Me.WindowState = 0
Else
    I = NotAlwaysOnTop(Me.hwnd)
End If
OTF3 = TF3

End Sub

Private Sub Timer5_Timer()

YY = False
End Sub
