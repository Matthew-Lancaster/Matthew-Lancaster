VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Status Of Load"
   ClientHeight    =   1170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1170
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   2235
      Top             =   690
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   90
      Top             =   135
   End
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   480
      Left            =   -15
      TabIndex        =   0
      Top             =   675
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   847
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "The 1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   690
      Left            =   -15
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub RunIn()

On Error GoTo ErrTrap8

'Why
If FirstRun3 = False Then Unload Form1

If ScanInProgress = False Then
ScanInProgress = True

Load Form1
Label1 = "The One"

'fMain.AutoCountTimer.Enabled = False
'fMain.Timer1.Enabled = False

Form1.Top = fMain.Top + 680
Form1.Left = fMain.Width / 2 - Form1.Width / 2
'fMain.Enabled = False
'fMain.AutoRedraw = False
'Form1.AutoRedraw = False
'Form1.Refresh
'DoEvents
End If

XD = 0
If AgWa > 0 Then
    XD = fMain.Lb2(AgWa - 1)
    If Val(fMain.Lb2(AgWa - 1)) = 0 Then Form1.Height = PBar1.Top
End If

Form1.Show

Exit Sub
ErrTrap8:
MsgBox "Error In Run Count"

End Sub

Public Sub RunOut()
RunOut2 = True
'Call RunCount
Label1.Caption = "100%"
Timer1.Enabled = False
Timer1.Interval = 100
Timer1.Enabled = True

'Timer2.Enabled = True
End Sub


Public Sub RunCount()

On Error GoTo ErrTrap5

If Scanpath.lblCount = "" Then Exit Sub

If Val(fMain.Lb2(AgWa - 1)) > 0 Then
    If Val(Scanpath.lblCount) > XD Then XD = Scanpath.lblCount


    If PBar1.Value <> Int(((Scanpath.lblCount / XD) * 100)) Then
    'ts = Format$((Scanpath.lblCount / XD) * 100, "#0") + "%"
    'Label1.Caption = Trim(Str(Scanpath.lblCount)) + " of " + fMain.Lb2(AgWa - 1) + " " + ts
    ts = Format$((Scanpath.lblCount / XD) * 100, "#0") + "%"
    If Timer2.Enabled = True And (Scanpath.lblCount / XD) < 0.2 And XD > 500 Then ts = "The One"
    Label1.Caption = ts
    End If
    
    PBar1.Value = Int(((Scanpath.lblCount / XD) * 100))

Else
    Label1.Caption = Trim(Str(Scanpath.lblCount))
End If

Exit Sub
ErrTrap5:
MsgBox "Error In Run Count"

End Sub

Private Sub Form_Load()
On Error GoTo ErrTrap2
If FirstRun3 = False Then Unload Form1
Exit Sub
ErrTrap2:
MsgBox "Error In Form_Load"
End Sub

Private Sub Form_LostFocus()
On Error GoTo ErrTrap3
Call Form_Resize
Exit Sub
ErrTrap3:
MsgBox "Error In Form_Lost Focus"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If Timer1.Enabled = True Then Call Timer1_Timer
End Sub

Private Sub Form_Resize()
On Error GoTo ErrTrap4
On Error Resume Next
Form1.Top = fMain.Top + 680
Form1.Left = fMain.Width / 2 - Form1.Width / 2
Timer2.Enabled = True
Form1.SetFocus

'Call Form_Resize

Exit Sub
ErrTrap4:
Exit Sub
MsgBox "Error In Form_Resize"

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If Timer1.Enabled = True Then Call Timer1_Timer
End Sub




Public Sub Timer1_Timer()
On Error GoTo ErrTrap

Timer1.Enabled = False
Timer2.Enabled = False
DoEvents
'fMain.AutoRedraw = True
'Form1.Visible = False
'fMain.Enabled = True
'fMain.AutoCountTimer.Enabled = True
'fMain.Timer1.Enabled = True
Unload Form1
Exit Sub
ErrTrap:
MsgBox "Error In Timer1"

End Sub

Private Sub Timer2_Timer()

Timer2.Enabled = False

'Form1.Visible = True
'Form1.SetFocus
End Sub
