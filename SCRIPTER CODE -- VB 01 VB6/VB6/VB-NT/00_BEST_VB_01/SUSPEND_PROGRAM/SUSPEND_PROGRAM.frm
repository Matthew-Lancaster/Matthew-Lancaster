VERSION 5.00
Begin VB.Form A_SUSPEND 
   BackColor       =   &H00FF0000&
   Caption         =   "SUSPEND PROGRAM"
   ClientHeight    =   4305
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer4 
      Interval        =   5000
      Left            =   3480
      Top             =   240
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3000
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2160
      Top             =   90
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   1605
      Top             =   105
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "PAUSE OF VICE VERSA && OUTLOOK EXPRESS && SUPER COPIER WIN-RAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3930
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   4740
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Label2"
      Height          =   300
      Left            =   636
      TabIndex        =   2
      Top             =   2496
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Label1"
      Height          =   372
      Left            =   1716
      TabIndex        =   1
      Top             =   2424
      Visible         =   0   'False
      Width           =   444
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB ME"
   End
End
Attribute VB_Name = "A_SUSPEND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IN1, IN2, IN3, IN4, IN5, STOPX, XPAUSE, OXPAUSE, Count5
Dim OVPROG1, OVPROG2
Dim VPROG1 As Long, VPROG2 As Long
Dim VPROGID1 As Long, VPROGID2 As Long
Dim OX1, OY1, MOUSE, KCODE
Dim TAGZ


Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type



Sub SUSPENDCODE()


'If OXPAUSE = XPAUSE Then Exit Sub






Dim RF As Long
Dim OU(20) As Long

Label = "PAUSE OF" + vbCrLf + "VICE VERSA" + vbCrLf + "OUTLOOK EXPRESS"
Label = Label + vbCrLf + "SUPER COPIER" + vbCrLf + "WIN-RAR" + vbCrLf + "Sort Anything"
Label = Label + vbCrLf + "TreeSize Professional"
Label = Label + vbCrLf + "Create Dir Another Drive"
'Label = Label + vbCrLf + ""
'Label = Label + vbCrLf + ""




Command1.Caption = UCase(Label)

'C:\Program Files\ViceVersa Pro 2\ViceVersa.exe
i = 0
i = i + 1: OU(i) = -1
RF = cProcesses.GetEXEID(OU(i), "ViceVersa.exe")
i = i + 1: OU(i) = -1
RF = cProcesses.GetEXEID(OU(i), "msimn.exe")
i = i + 1: OU(i) = -1
RF = cProcesses.GetEXEID(OU(i), "SuperCopier2.exe")
i = i + 1: OU(i) = -1
RF = cProcesses.GetEXEID(OU(i), "C:\Program Files\WinRAR\WinRAR.exe")
i = i + 1: OU(i) = -1
RF = cProcesses.GetEXEID(OU(i), "D:\VB6\VB-NT\00 Sort Anything\#Sort_Anything Lastest Best Ver 2009-04-25 - MULTI USE MENU\Sort_AnyThing - MULTI USE MENU.exe")
'C:\Program Files\Internet Explorer\IEXPLORE.EXE
'VB FOR HIGH TRUE SET
i = i + 1: OU(i) = -1
RF = cProcesses.GetEXEID(OU(i), "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE")


i = i + 1: OU(i) = -1
RF = cProcesses.GetEXEID(OU(i), "C:\Program Files\JAM Software\TreeSize Professional\TreeSize.exe")

'PROPER
i = i + 1: OU(i) = -1
RF = cProcesses.GetEXEID(OU(i), "D:\VB6\VB-NT\00 Sort Anything\#Sort_Anything Lastest Best Ver 2011-04-17 - COPY - MULTI USE MENU  - AND COPY\Sort_AnyThing - COPY  - MULTI USE MENU - COPY.exe")

'
i = i + 1: OU(i) = -1
RF = cProcesses.GetEXEID(OU(i), "D:\VB6\VB-NT\00_Best_VB_01\#0 SMALL PROGS\Create on another drive\Create Dir Another Drive.exe")

i = i + 1: OU(i) = -1
RF = cProcesses.GetEXEID(OU(i), "C:\Program Files\IrfanView\i_view32.exe")

If XPAUSE = True Then

    'OU1 = InstanceToWnd(OU1)
    'test_thread_id = GetWindowThreadProcessId(test_hwnd, test_pid)

    'Call Thread_Suspend(OU1)
    TAGZ = False
    For r = 1 To i
    
    If InStr(OU(r), "VB6.EXE") = 0 Then
        TAGZ = True
    End If
'    If InStr(OU(r), "VB6.EXE") = 0 Then
'        TAGZ = True
'    End If
    
    If TAGZ = False Then
    ' -- VB 6 ALWAYS HIGH
        Call Process_LOW_PRIORITY_CLASS(OU(r))
    End If
    Next
Else
    
    'OU1 = InstanceToWnd(OU1)
    'Call Thread_Resume(OU1)
    
    For r = 1 To i
    'If r <> 6 Then ' -- VB 6 ALWAYS HIGH
        Call Process_HIGH_PRIORITY_CLASS(OU(r))
        'Call Process_REALTIME_PRIORITY_CLASS(OU(i)) ' ---- SUPER
    'End If
    Next
    

End If




Exit Sub

'Call Process_Idle(i)
'i = Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE", vbNormalFocus)
i = ExecCmd("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE")


Do

t = Now + TimeSerial(0, 0, 2)
Do
'A = GetWindowTitle(i)

'a = GetActiveWindowTitle(i)

'A = GetActiveWindowClassParent(i)
'A = GetActiveWindowClass(i)
If a <> oa Then B = B + a + vbCrLf
oa = a
DoEvents

Loop Until t < Now
MsgBox B

Call Thread_Suspend(test_thread_id)

O = MsgBox("Hi", vbOKCancel)
    
Call Thread_Resume(test_thread_id)

Loop Until O = vbCancel

End


End Sub

Private Sub Form_Load()
    
TAGZ = False
    
    'List_ActiveThreads

    Exit Sub
    
    Load JOYFRM
    MOUSE = Now + TimeSerial(0, 0, 3)

End Sub

Private Sub Form_Resize()
On Error Resume Next

'Me.Show
'Me.WindowState = 0
DoEvents
Command1.Left = 0
Command1.Top = 0
Me.Width = Command1.Width + 70
Me.Height = Command1.Height + 700

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub MNU_VB_Click()

Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
End
End Sub

Private Sub Timer1_Timer()

If KCODE > 0 Then Exit Sub
i = 0
For i = 1 To 255
    bdf = GetAsyncKeyState(i)
    If bdf < -300 Then
        KCODE = True: Exit For
    End If
Next

End Sub


Private Sub Timer2_Timer()


Dim x1, y1
'Dim Mouse, OX1, OY1

FindCursor x1, y1


VPROG1 = FindWindow("#32770", "Output")
If VPROG1 > 0 And OVPROG1 <> VPROG1 Then
    VPROGID1 = True
End If

VPROG2 = FindWindow("Outlook Express Browser Class", vbNullString)
If VPROG2 > 0 And OVPROG2 <> VPROG2 Then
    VPROGID2 = True
End If

OVPROG1 = VPROG1
OVPROG2 = VPROG2



If VPROGID1 = True Then VPROGID1 = 0: OX1 = -1
If VPROGID2 = True Then VPROGID2 = 0: OX1 = -1
If KCODE = True Then KCODE = 0: OX1 = -1
'If JCODE = True Then JCODE = 0: OX1 = -1


If OX1 <> x1 Or OY1 <> y1 Then
    MOUSE = Now + TimeSerial(0, 0, 5)
End If

If MOUSE > Now Then
    If XPAUSE <> True Then
        XPAUSE = True
        SUSPENDCODE
        'Debug.Print "YES"
    End If
Else
    If XPAUSE <> False Then
        XPAUSE = False
        SUSPENDCODE
    End If
End If

If XPAUSE = True Then Command1.BackColor = Label2.BackColor
If XPAUSE = 0 Then Command1.BackColor = Label1.BackColor


OX1 = x1
OY1 = y1


End Sub

Public Sub FindCursor(X, Y)

'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'FindCursor Tx, Ty
'Private Type POINTAPI
'        x As Long
'        Y As Long
'End Type

Dim P As POINTAPI

GetCursorPos P
'   return x and y co-ordinate
X = P.X ' / GetSystemMetrics(0) * Screen.Width
'   for current cursor position
Y = P.Y '/ GetSystemMetrics(1) * Screen.Height

End Sub

Private Sub Timer3_Timer()
'
'If Count5 < 100 Then Exit Sub
'Count5 = 0
'Dim OUH As Long
'OUH = -1
'RF = cProcesses.GetEXEID(OUH, "C:\Program Files\Internet Explorer\IEXPLORE.EXE")
'If OUH > 0 And RF = True Then
'    Process_Kill (OUH)
'
'End If
End Sub

Private Sub Timer4_Timer()
'VPROG = FindWindow("IEFrame", vbNullString)
'If VPROG > 0 Then
'    ShowWindow VPROG, SW_HIDE
'End If
End Sub
