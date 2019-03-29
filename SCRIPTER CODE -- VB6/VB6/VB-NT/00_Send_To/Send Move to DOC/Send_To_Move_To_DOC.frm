VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "Explorer From Send To or Clipboard"
   ClientHeight    =   3972
   ClientLeft      =   60
   ClientTop       =   576
   ClientWidth     =   9516
   LinkTopic       =   "Form1"
   ScaleHeight     =   3972
   ScaleWidth      =   9516
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer_EXIT 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7188
      Top             =   2460
   End
   Begin VB.Label Label3 
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9132
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "OR FOLDER WANTED"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   28.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1296
      Left            =   120
      TabIndex        =   1
      Top             =   2340
      Visible         =   0   'False
      Width           =   5796
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "FILE GIVEN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   120
      TabIndex        =   0
      Top             =   1428
      Width           =   9120
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB"
   End
   Begin VB.Menu MNU_VB_FOLDER 
      Caption         =   "VB FOLDER"
   End
   Begin VB.Menu MNU_STATUS 
      Caption         =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FS        ' Set = CreateObject("Scripting.FileSystemObject")
Dim PATH_FILE
Dim AT_TEMP
Dim FILE


Private Sub Form_Load()

'AT_TEMP = "D:\VB6\VB-NT\00_Send_To\Send To Removed Mp3 Zen Folder\Send To Removed Mp3 Zen Folder.vbp"
'w$ = "D:\VB6\VB-NT\00_Best_VB_01\Auto Start Menu\Auto Start Menu.exe"
'If IsIDE = True Then AT$ = "\\4-asus-gl522vw\4-asus-g_asus_g_d\DSC\2015 2016\2016 Sony CyberShot DSC-HX60V"
'If IsIDE = True Then AT$ = "\\4-asus-gl522vw\4-asus-g_asus_g_d\DSC\2015 2016"
'AT_ISIDE = "D:\DSC"
'Fs.FileSystemObject.CopyFolder

Set FS = CreateObject("Scripting.FileSystemObject")

AT_TEMP = Command$
AT_TEMP = Replace(AT_TEMP, """", "")

'---------
'TEST MODE -- FILE
'---------
AT_ISIDE = "C:\TEMP\bcdinfo.txt"
If IsIDE = True Then AT_TEMP = AT_ISIDE
If IsIDE = True Then AT_TEMP = ""
'---------
'TEST MODE -- FOLDER
'---------
'AT_ISIDE = "C:\TEMP\"
'If IsIDE = True Then AT_TEMP = AT_ISIDE


Me.Height = Label1.Top + Label1.Height + (1200)
Me.Width = Label3.Left + Label3.Width + (300)



'--------------------------
'FIND COMMANDLINE - SEND TO
'--------------------------
If FS.FolderExists(AT_TEMP) = True Then
    PATH_FILE = AT_TEMP
    SENDTO = "SEND TO RIGHT CLICK"
End If
If FS.FileExists(AT_TEMP) = True Then
    PATH_FILE = AT_TEMP
    FILE = True
    SENDTO = "SEND TO RIGHT CLICK"
End If

SENDTO = "SEND TO RIGHT CLICK"

'---------------------
'NONE SEND TO CMD LINE -- THEN
'IS CLIPBOARD ANY
'---------------------
'AT_TEMP = Replace(Trim(Clipboard.GetText), """", "")
'If PATH_FILE = "" Then
'
'    If FS.FolderExists(AT_TEMP) = True Then
'        PATH_FILE = AT_TEMP
'        SENDTO = "CLIPBOARD FIND PICK"
'    End If
'
'    If FS.FileExists(AT_TEMP) = True Then
'        PATH_FILE = AT_TEMP
'        SENDTO = "CLIPBOARD FIND PICK"
'        FILE = True
'    End If
'End If
    

Label3 = PATH_FILE

MNU_STATUS.Caption = "  -- SOURCE GIVEN IS -- " + SENDTO

'If FILE = True Then
'TIMER1.Enabled = True
'End If


'--------------------------
'STILL NONE AFTER CLIPBOARD
'OR
'SEND TO
'USE REMEMBER FILE
'--------------------------

If PATH_FILE = "" Then
    'MsgBox "NOTHING WITH VALID FOLDER OF FILE NAME GIVEN" + vbCrLf + "COMMAND$" + vbCrLf + Command$ + vbCrLf + "OR CLIPBOARD" + vbCrLf + AT_TEMP
    'End
    Label1.Caption = "File Move to Doc __ Nothing Given"
End If


'Label1

'ChDrive w$
'ChDir w$

' MsgBox "C:WINDOWS\EXPLORER.EXE /SELECT, """ + PATH_FILE + """"

Me.Show

Call SEND_TO_DOC

' Shell "C:WINDOWS\EXPLORER.EXE /SELECT, """ + PATH_FILE + """", vbMaximizedFocus



Exit Sub

If FILE = False Then
    Shell "explorer /e, /select, " + PATH_FILE, vbNormalFocus
    End
Else
    Exit Sub
    Shell "explorer /e, /select, " + PATH_FILE, vbNormalFocus
End If

End Sub


Sub SEND_TO_DOC()

FOLDER_PATH = Mid(PATH_FILE, 1, InStrRev(PATH_FILE, "\"))

If Dir(FOLDER_PATH + "\DOC", vbDirectory) <> "" Then
    'MsgBox "FOLDER ALREADY CREATED"
Else
    On Error Resume Next
    MkDir FOLDER_PATH + "\DOC"
    If Err.Number > 0 Then
        MsgBox "CREATE FOLDER ERROR" + vbCrLf + FOLDER_PATH + "\DOC" + vbCrLf + Err.Description
    End If
End If
If FS.FileExists(PATH_FILE) = True Then
    
    PATH_FROM_PATH_FILENAME = Mid(PATH_FILE, 1, InStrRev(PATH_FILE, "\"))
    PATH_FROM_PATH_FILENAME = PATH_FROM_PATH_FILENAME + "DOC\"
    
    FILENAME_STRIPED_FROM_PATH = Mid(PATH_FILE, InStrRev(PATH_FILE, "\") + 1)
    FN2 = PATH_FROM_PATH_FILENAME + FILENAME_STRIPED_FROM_PATH
    
    On Error Resume Next
    FS.MoveFile PATH_FILE, FN2
    If Err.Number = 0 Then
        Label1.Caption = "File Move to Doc __ Done"
    Else
        
        DETAIL_VAR = "MOVE FROM --" + vbCrLf + vbCrLf + FILE_NAME
        DETAIL_VAR = "MOVE TO   --" + vbCrLf + vbCrLf + FN2
        MsgBox "WAS A ERROR MOVE FILE REQUESTED" + vbCrLf + vbCrLf + "Err.Description " + vbCrLf + Err.Description + vbCrLf + vbCrLf + DETAIL_VAR, vbMsgBoxSetForeground
    End If
End If

Timer_EXIT.Enabled = True

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Timer_EXIT.Enabled = False
Timer_EXIT.Interval = 4000
Timer_EXIT.Enabled = True


End Sub

Private Sub Label1_Click()

Exit Sub

If FILE = True Then
    Shell "explorer /e, /select, " + PATH_FILE, vbNormalFocus
    End
Else
    'Exit Sub
    'Shell "explorer /e, /select, " + PATH_FILE, vbNormalFocus
End If

End Sub

Private Sub Label2_Click()
FILE = False
PATH_FILE = Mid(PATH_FILE, 1, InStrRev(PATH_FILE, "\"))
If FILE = False Then
    Shell "explorer /e, /select, " + PATH_FILE, vbNormalFocus
    End
Else
    'Exit Sub
    'Shell "explorer /e, /select, " + PATH_FILE, vbNormalFocus
End If

End Sub


Private Sub MNU_VB_FOLDER_Click()

Shell "explorer /e, /select, " + App.Path, vbNormalFocus
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

'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
  'Test = False
End Function
'***********************************************
Private Sub Timer_EXIT_Timer()
End
End Sub
