VERSION 5.00
Begin VB.Form Frm_Process_Hog 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Process_Hogs"
   ClientHeight    =   2025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2970
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer_Kill_Usage_Hogs1 
      Interval        =   10000
      Left            =   1800
      Top             =   315
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "00:00 -- 00:00"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   990
   End
End
Attribute VB_Name = "Frm_Process_Hog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Public cProcesses As New clsCnProc
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


Private Function AlwaysOnTop(ByVal hwnd As Long)  'Makes a form always on top
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Function
Private Function NotAlwaysOnTop(ByVal hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, flags
End Function



Private Sub Form_Load()


If IsIDE = False Then
    PID = -1
    Var = cProcesses.GetEXEID(PID, App.Path + "\" + App.EXEName + ".exe")
    If PID <> -1 Then
        Call Process_HIGH_PRIORITY_CLASS(PID)
    End If
End If



Call Timer_Kill_Usage_Hogs1_Timer
Call Timer_Kill_Usage_Hogs3_Timer


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
    Timer_Kill_Usage_Hogs3.Interval = 10000
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

Frm_Process_Hog.Label1 = Format(Kill_Usage_Hogs_DATE1, "HH:MM:SS") + " -- " + Format(Now, "HH:MM:SS") + " --" + Str(XDL)

If Kill_Usage_Hogs_DATE1 + TimeSerial(0, 2, 0) < XDL2 Then
    Frm_Process_Hog.Label1.BackColor = vbRed
Else

    If Kill_Usage_Hogs_DATE1 = Now Then
        Frm_Process_Hog.Label1.BackColor = vbGreen
    Else
        Frm_Process_Hog.Label1.BackColor = vbYellow
    
    End If
End If

Call FRM_PROCESS_HOGS_SHOW

Exit Sub

PID = -1
Var = cProcesses.GetEXEID(PID, App.Path + "\" + App.EXEName + ".exe")
If PID <> -1 Then
    Call Process_HIGH_PRIORITY_CLASS(PID)
End If


End Sub


Sub FRM_PROCESS_HOGS_SHOW()

Frm_Process_Hog.Label1.Refresh
Frm_Process_Hog.Label1.AutoSize = True
Frm_Process_Hog.Label1.Refresh
DoEvents

Frm_Process_Hog.Visible = True
Frm_Process_Hog.Show
Frm_Process_Hog.WindowState = vbNormal
Frm_Process_Hog.Top = 0
Frm_Process_Hog.Left = 0 'Screen.Width - 200
Frm_Process_Hog.Label1.Left = 0
Frm_Process_Hog.Label1.Top = 0

Frm_Process_Hog.Width = Frm_Process_Hog.Label1.Width ' * Screen.TwipsPerPixelX
Frm_Process_Hog.Height = Frm_Process_Hog.Label1.Height ' * Screen.TwipsPerPixelY

i = AlwaysOnTop(Frm_Process_Hog.hwnd)

End Sub


Sub Kill_Usage_Hogs()


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

