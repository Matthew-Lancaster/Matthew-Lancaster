VERSION 5.00
Begin VB.Form FormSYNC 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   192
   ClientTop       =   792
   ClientWidth     =   9624
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   9624
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2580
      Top             =   2415
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2580
      Top             =   1845
   End
   Begin VB.Timer TimerTime 
      Interval        =   1000
      Left            =   3045
      Top             =   1140
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2520
      Top             =   1110
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808080&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.65
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   0
      TabIndex        =   6
      Top             =   4320
      Width           =   9600
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808080&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.65
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   0
      TabIndex        =   5
      Top             =   3615
      Width           =   9600
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.65
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   2910
      Width           =   9600
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.65
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   -15
      TabIndex        =   3
      Top             =   2190
      Width           =   9600
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.65
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   0
      TabIndex        =   2
      Top             =   1485
      Width           =   9600
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.65
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   750
      Width           =   9600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.65
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   9600
   End
   Begin VB.Menu MNU_VB 
      Caption         =   "VB ME"
   End
End
Attribute VB_Name = "FormSYNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmdLine, FileSpec_Source, FileSpec_Target

Dim Countx

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long
    
Private FS, FS2
Dim XCOUNT




Private Sub Form_Load()

Set FS = CreateObject("Scripting.FileSystemObject")

Me.Caption = "Network Sync EXE Clipboard Logger"

Label1.BackColor = vbGreen
Label2.BackColor = vbYellow
Label3.BackColor = vbYellow
Label4.BackColor = vbYellow
Label5.BackColor = vbWhite
Label6.BackColor = vbWhite
Label7.BackColor = vbWhite
Label5.FontSize = 14
Label6.FontSize = 12
Label7.FontSize = 12

Label2.Caption = "01 of 03 Wait for App To Close"
Label3.Caption = "02 of 03 Make Copy of Work NetWork Sync EXE"
Label4.Caption = "03 of 03 Restart Program Again"

cmdLine = Command$
FileSpec_Target = cmdLine

FileSpec_Target = "D:\VB6\VB-NT\00_Best_VB_01\ClipBoard Logger\ClipBoard Logger.exe"

Call TimerTime_Timer

If Trim(FileSpec_Target) = "" Then

    MsgBox "Command Line Should Exist Abort Maybe", vbMsgBoxSetForeground
'    TimerTime.Enabled = False
'    Timer1.Enabled = False
'    Exit Sub

End If

If cmdLine = "" Then
    cmdLine = "D:\VB6\VB-NT\00_Best_VB_01\ClipBoard Logger\Network_Sync_EXE_Clipboard.exe ""D:\VB6\VB-NT\00_Best_VB_01\ClipBoard Logger\ClipBoard Logger.EXE"""
End If

cmdLine = Mid(cmdLine, InStr(cmdLine, """"))
cmdLine = Replace(cmdLine, """", "")
cmdLine = Replace(cmdLine, ".EXE", ".exe")
FileSpec_Target = cmdLine
cmdLine = Mid(cmdLine, InStrRev(cmdLine, "\") + 1)

FileSpec_Target = Mid(FileSpec_Target, 1, InStrRev(FileSpec_Target, "\"))
FileSpec_Source = Replace(FileSpec_Target, ":\VB6\", ":\VB6-EXE'S\")

'MsgBox ea
Label5.Caption = "FileName : " + cmdLine
Label6.Caption = "Source   : " + FileSpec_Source
Label7.Caption = "Target   : " + FileSpec_Target


On Error Resume Next
If Dir(FileSpec_Source + cmdLine) = "" Then
    xsp = 1
End If
If Err.Number > 0 Then MsgBox "ERROR AT THIS HANDLING -- If Dir(FileSpec_Source + cmdLine)"
Err.Clear
If Dir(FileSpec_Target + cmdLine) = "" Then
    xsp = 1
End If
If Err.Number > 0 Then MsgBox "ERROR AT THIS HANDLING -- If Dir(FileSpec_Target + cmdLine)"
On Error GoTo 0




If xsp = 1 Then

    MsgBox "Source and Target Should Exist Abort", vbMsgBoxSetForeground
    TimerTime.Enabled = False
    Timer1.Enabled = False
    Exit Sub

End If

Timer1.Enabled = True

End Sub

Private Sub MNU_VB_Click()
Shell """C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"" """ + App.Path + "\" + App.EXEName + ".vbp""", vbNormalFocus
EXIT_TRUE = True
Unload Me
End Sub

Private Sub Timer1_Timer()


xhWnd = FindWindow("ThunderFormDC", "ClipBoard Logger")
If xhWnd = 0 Then
    xhWnd = FindWindow("ThunderRT6FormDC", "ClipBoard Logger")
End If

'APP SHOULD INIT CLOSE ITSELF TO TEST THIS HERE DO CLOSE MANUAL

Countx = Countx + 1
Label2.Caption = "01 of 03 Wait for App To Close - WAITING" + Str(Countx)

If xhWnd > 0 Then Exit Sub

Timer1.Enabled = False
Timer2.Enabled = True
Label2.Caption = "01 of 02 Wait for App To Close - Done"
Label2.BackColor = vbGreen

End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False

Label3.Caption = "02 of 03 Make Copy of Work NetWork Sync EXE % Read Write Copy Riper"

FileSpec_Source = FileSpec_Source + cmdLine
FileSpec_Target = FileSpec_Target + cmdLine

On Error Resume Next
FS.CopyFile FileSpec_Source, FileSpec_Target, True


If Err.Number > 0 Then
    XCOUNT = XCOUNT + 1
    Timer2.Enabled = True
    If XCOUNT > 20 Then
        
        MsgBox "Error Delete Target Before Copy" + vbCrLf + Err.Description + vbCrLf + Trim(Str(Err.Number)) + vbCrLf + "Abort Job"
        Exit Sub
        
    End If
End If

On Error GoTo 0


'FR1 = FreeFile
'Open FileSpec_Source + cmdLine For Input As #FR1
'fsize = LOF(FR1)
'Close #FR1
'
'FR1 = FreeFile
'Open FileSpec_Source + cmdLine For Binary As #FR1
'filevar = Input(fsize, FR1)
'Close #FR1

Label3.Caption = "02 of 03 Make Copy of Work NetWork Sync EXE % Done Rip"




'Label3.Caption = "02 of 03 Make Copy of Work NetWork Sync EXE % 3 Write"

'FR1 = FreeFile
'Open FileSpec_Target + cmdLine For Binary As #FR1
'Put #FR1, , filevar
'Close #FR1

Label3.BackColor = vbGreen
'Label3.Caption = "02 of 03 Make Copy of Work NetWork Sync EXE % 3 Done"

Timer3.Enabled = True

End Sub


Private Sub Timer3_Timer()

Label4.Caption = "03 of 03 Restart Program Again - Going"
Debug.Print FileSpec_Target
I = Shell(FileSpec_Target, vbMinimizedNoFocus)

If I = False Then
    Label4.Caption = "03 of 03 Restart Program Again - Going NOT YET"
    Exit Sub
End If

Label4.Caption = "03 of 03 Restart Program Again - Going Bye"
Label4.BackColor = vbGreen
TimerTime.Enabled = False

Timer3.Enabled = False

'Me.WindowState = vbMinimized
'End

End Sub

Private Sub TimerTime_Timer()
Label1.Caption = Format(Now, "DD-MMM-YYYY HH:MM:SS")
End Sub
