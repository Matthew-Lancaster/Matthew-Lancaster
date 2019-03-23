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
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   1275
      Top             =   1245
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
      TabIndex        =   0
      Top             =   0
      Width           =   6255
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
Dim Result&, FileName
Dim ReturnString As String * 1024

Private Sub RecordSound()

Dim i As Integer

  
'Close any MCI operations from previous VB programs
i = mciSendString("close all", 0&, 0, 0)

'Open a new WAV with MCI Command...
i = mciSendString("open new type waveaudio alias capture", 0&, 0, 0)

'Samples Per Second that are supported:
'11025   low quality
'22050   medium quality
'44100 high quality (CD music quality)


'Bits per sample is 16 or 8

'Channels are 1 (mono) or 2 (stereo)


i = mciSendString("set capture channels 2", 0&, 0, 0)    '

'start at begining
i = mciSendString("seek capture to start", 0&, 0, 0)    'Always start at the beginning

i = mciSendString("set capture samplespersec 11025", 0&, 0, 0)

i = mciSendString("set capture bitspersample 16", 0&, 0, 0)  '16 bits for better sound

i = mciSendString("record capture", 0&, 0, 0)  'Start the recording

   
End Sub
 
Private Sub About_Click()
MsgBox ("Made By Jesse Friedman, 2001")
End Sub




Private Sub Form_Load()
'MsgBox "Please make the project into an .exe file because the recorder functions better.", vbOKOnly, "Stealth Sound Recorder"

'Form1.Show
'DoEvents
'Refresh


Dim i As Integer
Dim intLoop As Integer

'Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

'Close any MCI operations from previous VB programs
i = mciSendString("close all", 0&, 0, 0)

'Open a new WAV with MCI Command...
i = mciSendString("open new type waveaudio alias capture", 0&, 0, 0)

'WindowsMediaPlayer1.settings.volume = 100


Call RecordSound


End Sub

Private Sub Timer1_Timer()

tt$ = "Stealth Record " + Format$(Now, "YYYY-MM-DD HH-MM-SS DDD") + ".wav"
tt$ = "RS-Happy.wav"

FileName = "D:\05 Media\Stealth_Recording\" + tt$
If Dir$(FileName) <> "" Then Kill FileName
'Filename = "D:\temp\" + tt$

Dim i As Integer
i = mciSendString("stop capture", 0&, 0, 0)

'MCI command to save the WAV file
i = mciSendString("save capture " & """" & FileName & """", 0&, 0, 0)

End

End Sub

