VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   2172
   ClientLeft      =   192
   ClientTop       =   456
   ClientWidth     =   9336
   LinkTopic       =   "Form1"
   ScaleHeight     =   2172
   ScaleWidth      =   9336
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   3960
      Top             =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   80
      TabIndex        =   0
      Top             =   80
      Width           =   9132
      WordWrap        =   -1  'True
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

'E$ = "kujlhglkjhgf"""

'e$ = "D:\VB6\VB-NT\00_Best_VB_01\WebDates_Ace\WebDates.vbp"


'EXAMPLES
'MsgBox "-" + Command$ + "-"
'---------------------------
'#0 Send To ClipBoard Path & FileName
'---------------------------
'-"D:\VB6\VB-NT\00_Send_To\ClipBoard Path & FileName\#0 Send To ClipBoard Path & FileName.exe"-
'---------------------------
'OK
'---------------------------
'---------------------------
'#0 Send To ClipBoard Path & FileName
'---------------------------
'-"D:\VB6\VB-NT\00_Send_To\ClipBoard Path & FileName"-
'---------------------------
'OK
'---------------------------

Me.Height = Label1.Top + Label1.Height + (830)
Me.Width = Label1.Left + Label1.Width + (300)

LINE_CLIP = Command$
LINE_CLIP = Replace(LINE_CLIP, """", "")

Label1 = LINE_CLIP

Clipboard.Clear
Clipboard.SetText LINE_CLIP

End Sub

Private Sub Label1_Click()
Timer1.Enabled = False
End
End Sub

Private Sub MNU_VB_Click()

If IsIDE = True Then
    MNU_VB.Enabled = False
    MsgBox "NOT WHEN ISIDE"
    Exit Sub
End If


If Dir("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE") <> "" Then
    Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
End If

If Dir("C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE") <> "" Then
    Shell """C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
    EXIT_TRUE = True
    Unload Me
End If

'
'
'
'Dim ReturnHwnd As Long
'Dim I
''VB ONLY WANTS THE 1ST OF THE 2 HWND
''ReturnHwnd = FindWindowPartVB("ClipBoard Logger - Microsoft Visual Basic[")
'
''DONT NEED ABOVE USE THIS
'X1 = FindWindow("wndclass_desked_gsk", vbNullString)
''------------------------------------------------
''FIND THE WINDOW PRIZE TREASURE CHEST BUST BLOSSOM
''TRAIN SPOTTER
''------------------------------------------------
''-----------------------------------------------------------
''CHECK CLASS NAME IS VB IDE WINDOW WE WANT OR IF ANOTHER NOT
''-----------------------------------------------------------
'X2 = GetWindowTitle(X1)
'If InStr(X2, " - Microsoft Visual Basic [") > 0 Then
'If InStr(X2, "ClipBoard Logger") > 0 Then
'    MsgBox "DON'T RUN VB IDE - LOADED"
'    I = GetWindowState(X1)
'    If I = vbMinimized Then
'        SHOWVAR = SW_SHOWDEFAULT
'        ShowWindow ReturnHwnd, SHOWVAR
'    End If
'    EXIT_TRUE = True
'    Unload Me
'    Exit Sub
'End If
'End If
'
''BETTER LINE NEXT DON'T USE VB MENU IF ISIDE
'
'
''------------------------------------------------
''TEMP WORK AROUND
''OVER DRIVE
''OVER RIDE
''------------------------------------------------
''FINDWINDOW PART PROBLEM IN EXE AND WHEN IN WIN 7
''------------------------------------------------
''WIN 7 PROBLEM MUST USE EXTRA LAST LEFT SQUARE BRACKET OF SERACH END ABOVE
''------------------------------------------------
'If ReturnHwnd > 0 Then
'    If MsgBox("ReturnHwnd" + Str(ReturnHwnd) + " -- " + "MeHwnd" + Str(Me.hWnd) + vbCrLf + "VB Code Clipboard Logger already Running - Sure Want to Run VB IDE", vbYesNo) = vbYes Then
'        WANT_TO_RUN_ANYWAY = True
'    End If
'End If
'
'
'If ReturnHwnd > 0 Then
'
'    I = GetWindowState(ReturnHwnd1)
'
'    If I = vbMinimized Then
'
''        SHOWVAR = SW_RESTORE
''        SHOWVAR = SW_SHOW
''        SHOWVAR = SW_SHOWNA
''        SHOWVAR = SW_SHOWDEFAULT
''        SHOWVAR = SW_SHOWNORMAL
''        SHOWVAR = SW_SHOWMAXIMIZED
'        SHOWVAR = SW_SHOWDEFAULT
'
'        ShowWindow ReturnHwnd, SHOWVAR
'
'        'GUESS MAYBE
'        'SetForegroundWindow (ReturnHwnd)
'
'        DoEvents
'
'    End If
'
'    'MADE REDUNDANT CODE BY CONDICTION HERE AND ABOVE
'    If WANT_TO_RUN_ANYWAY = False Then
'        EXIT_TRUE = True
'        Unload Me
'        Exit Sub
'    End If
'
'End If
'
'If Dir("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE") <> "" Then
'    Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
'    EXIT_TRUE = True
'    Unload Me
'End If
'
'If Dir("C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE") <> "" Then
'    Shell """C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
'    EXIT_TRUE = True
'    Unload Me
'End If
'
'

End Sub

Private Sub Timer1_Timer()

Unload Me

End Sub
