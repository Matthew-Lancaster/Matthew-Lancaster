VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Stealth Sound Recorder"
   ClientHeight    =   3885
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox filename 
      Alignment       =   2  'Center
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   4560
      TabIndex        =   8
      Text            =   "C:\a.wav"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   4560
      TabIndex        =   7
      Text            =   "Save in file:"
      Top             =   1530
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   600
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   3000
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Text            =   "00:00:00 AM"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Record In Stealth"
      Height          =   855
      Left            =   2400
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Text            =   "10 seconds"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Stealth Sound Recorder v1.0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Record at certain time (if a time is not specified, recording will begin right away): hh:mm:ss"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Length of time to record  :"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" _
                                   Alias "mciSendStringA" _
                                   (ByVal lpstrCommand As String, _
                                   ByVal lpstrReturnString As String, _
                                   ByVal uReturnLength As Long, _
                                   ByVal hwndCallback As Long) As Long


Private Function RecordSound(filename As String) As Boolean

tt$ = "RS-" + Format$(Now, "YYMMDD-HHMMSS") + ".wav"

filename = "C:\Media\Stealth_Recording\" + tt$
'filename = "D:\TEMP\ttt.wav" ' + tt$
    
    
    preservetime = Time
    Dim Result&
    Dim ReturnString As String * 1024
    Result& = mciSendString("open new Type waveaudio Alias recsound", ReturnString, Len(ReturnString), 0) 'Start at the beginning
    Result& = mciSendString("set recsound time format ms bitspersample 16 channels 2 bytespersec 22500 samplespersec 44100", ReturnString, 1024, 0) 'CD Quality Sound
     Result& = mciSendString("record  recsound", ReturnString, Len(ReturnString), 0) 'Start Recording
   'Simplified loops to make code easier
   Time$ = "00:00:00"
     If Combo1.Text = "10 seconds" Then
   Do
   x = x + 1
   Loop Until (Time$ = "00:00:10")
   End If
      If Combo1.Text = "20 seconds" Then
   Do
   x = x + 1
   Loop Until (Time$ = "00:00:20")
   End If
      If Combo1.Text = "30 seconds" Then
   Do
   x = x + 1
   Loop Until (Time$ = "00:00:30")
   End If
     If Combo1.Text = "40 seconds" Then
   Do
   x = x + 1
   Loop Until (Time$ = "00:00:40")
   End If
      If Combo1.Text = "50 seconds" Then
   Do
   x = x + 1
   Loop Until (Time$ = "00:00:50")
   End If
      If Combo1.Text = "1 minute" Then
   Do
   x = x + 1
   Loop Until (Time$ = "00:01:00")
   End If
        If Combo1.Text = "2 minutes" Then
   Do
   x = x + 1
   Loop Until (Time$ = "00:02:00")
   End If
      If Combo1.Text = "3 minutes" Then
   Do
   x = x + 1
   Loop Until (Time$ = "00:03:00")
   End If
      If Combo1.Text = "5 minutes" Then
   Do
   x = x + 1
   Loop Until (Time$ = "00:05:00")
   End If
        If Combo1.Text = "10 minutes" Then
   Do
   x = x + 1
   Loop Until (Time$ = "00:10:00")
   End If
      If Combo1.Text = "30 minutes" Then
   Do
   x = x + 1
   Loop Until (Time$ = "00:30:00")
   End If
         If Combo1.Text = "1 hour" Then
   Do
   x = x + 1
   Loop Until (Time$ = "01:00:00")
   End If
      If Combo1.Text = "2 hours" Then
   Do
   x = x + 1
   Loop Until (Time$ = "02:00:00")
   End If
      If Combo1.Text = "5 hours" Then
   Do
   x = x + 1
   Loop Until (Time$ = "05:00:00")
   End If
   If Combo1.Text = "10 hours" Then
   Do
   x = x + 1
   Loop Until (Time$ = "10:00:00")
   End If
     If Combo1.Text = "24 hours" Then
   Do
   x = x + 1
   Loop Until (Time$ = "24:00:00")
   End If
   
    Result& = mciSendString("save recsound " & filename, ReturnString, Len(ReturnString), 0) 'Save recording
    Result& = mciSendString("close recsound", ReturnString, 1024, 0) 'Close/Stop recording
  Time = preservetime
  End
  End Function
 
Private Sub About_Click()
MsgBox ("Made By Jesse Friedman, 2001")
End Sub

Private Sub Command1_Click()

If Text1.Text <> "00:00:00 AM" Then
Form1.Visible = False
Timer1.Enabled = True
Else
Form1.Visible = False
Call RecordSound(filename.Text)
Form1.Visible = False
End If
End Sub

Private Sub Form_Activate()
tt$ = "RR Record " + Format$(Now, "YYYY-MM-DD HH-MM-SS MMM DDD HH-MM-SSa/p") + ".wav"
filename.Text = "E:\Media\Stealth Recording\" + tt$

End Sub

Private Sub Form_Load()
'MsgBox "Please make the project into an .exe file because the recorder functions better.", vbOKOnly, "Stealth Sound Recorder"



End Sub

Private Sub Timer1_Timer()

If Text1.Text = Time Then


Call RecordSound(filename.Text)
End If
End Sub

Private Sub Timer2_Timer()
Label4.Caption = "Current time: " & Time
End Sub
