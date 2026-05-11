VERSION 5.00
Begin VB.Form Frm2_Process_Hog 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   Caption         =   "Process_Hogs"
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2970
   Icon            =   "Frm_Process_Hog.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1800
      Top             =   1725
   End
   Begin VB.Timer Timer_Kill_Usage_Hogs1 
      Interval        =   10000
      Left            =   1800
      Top             =   330
   End
   Begin VB.Timer Timer_Kill_Usage_Hogs2 
      Interval        =   5000
      Left            =   1800
      Top             =   780
   End
   Begin VB.Timer Timer_Kill_Usage_Hogs3 
      Interval        =   5000
      Left            =   1800
      Top             =   1260
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   " CNT "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1035
      TabIndex        =   4
      ToolTipText     =   "PROCESS COUNT"
      Top             =   1140
      Width           =   390
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   " MENU"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1215
      TabIndex        =   3
      ToolTipText     =   "MENU FORM"
      Top             =   1485
      Width           =   450
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "00:00 -- 00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   645
      TabIndex        =   2
      ToolTipText     =   "END"
      Top             =   825
      Width           =   870
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "00:00 -- 00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   420
      TabIndex        =   1
      ToolTipText     =   "VB ME"
      Top             =   465
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "00:00 -- 00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "KILL NOW"
      Top             =   165
      Width           =   870
   End
End
Attribute VB_Name = "Frm2_Process_Hog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------
'Project -- Process Hoggs Killer
'Matt.Lan@btinternet.com
'Created Mon 11-01-2016
'-------------------------------

Dim OldLenVarLabel3

Dim FS


'PID HAS TO BE -1 -- Process Id Return in PID
'Var - False if Not Exist or PID Remain -1
'----
'USAGE = ----
'PID = -1
'Var = cProcesses.GetEXEID(PID, App.Path + "\" + App.EXEName + ".exe")
'If PID <> -1 Then
'    Call Process_HIGH_PRIORITY_CLASS(PID)
'End If
'USAGE = ----
'----

Private cProcesses As New clsCnProc
Dim PID As Long



Public Kill_Usage_Hogs_DATE1
Dim Kill_Usage_Hogs2_CHECKGATE
Dim Kill_Usage_Hogs3_CHECKGATE


Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const MF_BYPOSITION = &H400&
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SPI_SCREENSAVERRUNNING = 97
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long




Public Sub SET_VAR_FS()

Set FS = CreateObject("Scripting.FileSystemObject")

End Sub



Private Function AlwaysOnTop(ByVal hwnd As Long)  'Makes a form always on top
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Private Function NotAlwaysOnTop(ByVal hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE
End Function



Private Sub Form_Load()

If App.PrevInstance = True Then End

If IsIDE = False Then
    PID = -1
    Var = cProcesses.GetEXEID(PID, App.Path + "\" + App.EXEName + ".exe")
    If PID <> -1 Then
        Call Process_HIGH_PRIORITY_CLASS(PID)
    End If
End If

Call SET_VAR_FS

Load Frm1_Menu

Call Timer_Kill_Usage_Hogs1_Timer



Me.Visible = True

End Sub


Private Sub Form_Unload(Cancel As Integer)
'If EXIT_TRUE = False Then
''    Me.WindowState = 1
'    'test may have to put back form need reseting
'    'Unload FrmJoy
'    Cancel = True
'    Exit Sub
'End If

'Cancel = True
'Me.WindowState = vbMinimized
End Sub

Private Sub Label1_Click()

Frm1_Menu.KILL_HOGGS.Enabled = True

Frm1_Menu.List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- Demand Button -- Kill_Usage_Hoggs"
Frm1_Menu.List3.AddItem "--------------------------------------------------------------------"

End Sub

Private Sub Label2_Click()
Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
EXIT_TRUE = True
Unload Me
End Sub

Private Sub Label3_Click()

End

Dim Form As Form
For Each Form In FORMS
    Unload Form
    Set Form = Nothing
Next Form

End

End Sub

Private Sub Label4_Click()

'PROCESS COUNT

End Sub

Private Sub Label5_Click()

Frm1_Menu.Visible = True
Frm1_Menu.WindowState = vbNormal
End Sub

Private Sub Timer_Kill_Usage_Hogs1_Timer()

On Error Resume Next

If Timer_Kill_Usage_Hogs1.Interval <> 10000 Then
    Timer_Kill_Usage_Hogs1.Interval = 10000
End If

FR1 = FreeFile
Open "C:\TEMP\Process_Hogs_Flag.txt" For Output As #FR1
Close #FR1

Call Timer_Kill_Usage_Hogs2_Timer
Call Timer_Kill_Usage_Hogs3_Timer

If Kill_Usage_Hogs2_CHECKGATE = True Then
    Kill_Usage_Hogs2_CHECKGATE = False
    Timer_Kill_Usage_Hogs2.Enabled = False
    Timer_Kill_Usage_Hogs2.Interval = 10000
    Timer_Kill_Usage_Hogs2.Enabled = True
End If

If Kill_Usage_Hogs3_CHECKGATE = True Then
    Kill_Usage_Hogs3_CHECKGATE = False
    Timer_Kill_Usage_Hogs3.Enabled = False
    Timer_Kill_Usage_Hogs3.Interval = 1000
    Timer_Kill_Usage_Hogs3.Enabled = True
End If

End Sub

Private Sub Timer_Kill_Usage_Hogs2_Timer()

On Error Resume Next

Set F = FS.getfile("C:\TEMP\Process_Hogs_Flag.txt")
Kill_Usage_Hogs_DATE1 = F.datelastmodified
Set F = Nothing

Kill_Usage_Hogs2_CHECKGATE = True

End Sub

Private Sub Timer_Kill_Usage_Hogs3_Timer()

'If Kill_Usage_Hogs_DATE1 = 0 Then Exit Sub

Kill_Usage_Hogs3_CHECKGATE = True

'X DATE LIMIT
Dim XDL, XDL2

XDL2 = Now
XDL = DateDiff("s", Kill_Usage_Hogs_DATE1, XDL2)

Frm2_Process_Hog.Label1 = Format(Kill_Usage_Hogs_DATE1, "HH:MM:SS")
Frm2_Process_Hog.Label2 = Format(Now, "HH:MM:SS")
Frm2_Process_Hog.Label3 = Str(XDL) + " "
'Frm2_Process_Hog.Label4 =

If Kill_Usage_Hogs_DATE1 + TimeSerial(0, 5, 0) < XDL2 Then
    
    Frm2_Process_Hog.Label1.BackColor = vbRed
    Frm2_Process_Hog.Label2.BackColor = Label1.BackColor
    Frm2_Process_Hog.Label3.BackColor = Label1.BackColor
    Frm2_Process_Hog.Label4.BackColor = Label1.BackColor
    Frm2_Process_Hog.Label5.BackColor = Label1.BackColor
    
    '--------------------
    'Call Kill_Usage_Hoggs
    Call Frm1_Menu.MNU_KILL_HOGGS_Click

    '--------------------
    Frm1_Menu.List3.AddItem Format(Now, "DD-MMM-YYYY HH:MM:SS") + " -- Kill_Usage_Hoggs - DETECT Computer Grinding to a Halt ON 5 MINS NOT RESPONDING"
    Frm1_Menu.List3.AddItem "--------------------------------------------------------------------"

    
Else

    If Kill_Usage_Hogs_DATE1 = Now Then
        Frm2_Process_Hog.Label1.BackColor = vbGreen
        Frm2_Process_Hog.Label2.BackColor = Label1.BackColor
        Frm2_Process_Hog.Label3.BackColor = Label1.BackColor
        Frm2_Process_Hog.Label4.BackColor = Label1.BackColor
        Frm2_Process_Hog.Label5.BackColor = Label1.BackColor
    Else
        Frm2_Process_Hog.Label1.BackColor = vbYellow
        Frm2_Process_Hog.Label2.BackColor = Label1.BackColor
        Frm2_Process_Hog.Label3.BackColor = Label1.BackColor
        Frm2_Process_Hog.Label4.BackColor = Label1.BackColor
        Frm2_Process_Hog.Label5.BackColor = Label1.BackColor
    
    End If
End If

Call Frm2_Process_HOGS_SHOW



End Sub


Sub Frm2_Process_HOGS_SHOW()

XT = 0
XT = XT + Len(Label3.Caption)
XT = XT + Len(Label4.Caption)

If XT = OldLenVarLabel3 Then Exit Sub

OldLenVarLabel3 = XT

Frm2_Process_Hog.Top = 0

Frm2_Process_Hog.Label1.Left = 0
Frm2_Process_Hog.Label1.Top = 0

gap = 80
Frm2_Process_Hog.Label2.Left = Label1.Width + gap
Frm2_Process_Hog.Label2.Top = 0

Frm2_Process_Hog.Label3.Left = Label2.Left + Label2.Width + gap
Frm2_Process_Hog.Label3.Top = 0

Frm2_Process_Hog.Label4.Left = Label3.Left + Label3.Width + gap
Frm2_Process_Hog.Label4.Top = 0

Frm2_Process_Hog.Label5.Left = Label4.Left + Label4.Width + gap
Frm2_Process_Hog.Label5.Top = 0

Frm2_Process_Hog.Width = Label5.Left + Label5.Width
Frm2_Process_Hog.Height = Frm2_Process_Hog.Label1.Height

Frm2_Process_Hog.Left = Screen.Width - Frm2_Process_Hog.Width

I = AlwaysOnTop(Frm2_Process_Hog.hwnd)

End Sub



'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function
'***********************************************

Private Sub Timer1_Timer()

I = AlwaysOnTop(Frm2_Process_Hog.hwnd)

End Sub
