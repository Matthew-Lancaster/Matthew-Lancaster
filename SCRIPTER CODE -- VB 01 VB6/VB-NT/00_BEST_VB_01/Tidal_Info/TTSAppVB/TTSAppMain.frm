VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form TTSAppMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAPI5 TTSAPP"
   ClientHeight    =   6612
   ClientLeft      =   3348
   ClientTop       =   3636
   ClientWidth     =   8088
   Icon            =   "TTSAppMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6612
   ScaleWidth      =   8088
   Visible         =   0   'False
   Begin VB.TextBox MainTxtBox2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      HideSelection   =   0   'False
      Left            =   1905
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   660
      Width           =   4575
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   910
      Left            =   1920
      TabIndex        =   28
      Top             =   1230
      Width           =   4560
   End
   Begin VB.PictureBox VisemePicture 
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2124
      ScaleWidth      =   1656
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   60
      Width           =   1700
   End
   Begin VB.CheckBox chkShowEvents 
      Caption         =   "Show Events"
      Height          =   195
      Left            =   4560
      TabIndex        =   25
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton StopBtn 
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6840
      TabIndex        =   2
      Top             =   500
      Width           =   1125
   End
   Begin VB.CheckBox chkSpFlagNLPSpeakPunc 
      Caption         =   "NLPSpeakPunc"
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CheckBox chkSpFlagPurgeBeforeSpeak 
      Caption         =   "PurgeBeforeSpeak"
      Height          =   255
      Left            =   6120
      TabIndex        =   22
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CheckBox chkSpFlagAync 
      Caption         =   "FlagsAsync"
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CheckBox chkSpFlagIsFilename 
      Caption         =   "IsFilename"
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CheckBox chkSpFlagPersistXML 
      Caption         =   "PersistXML"
      Height          =   255
      Left            =   4560
      TabIndex        =   21
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Speak Flags"
      Height          =   1575
      Left            =   4320
      TabIndex        =   18
      Top             =   2400
      Width           =   3615
      Begin VB.CheckBox chkSpFlagIsXML 
         Caption         =   "IsXML"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.ComboBox AudioOutputCB 
      Height          =   315
      ItemData        =   "TTSAppMain.frx":030A
      Left            =   840
      List            =   "TTSAppMain.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4125
      Width           =   3300
   End
   Begin VB.TextBox MainTxtBox 
      Height          =   525
      HideSelection   =   0   'False
      Left            =   1900
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "TTSAppMain.frx":030E
      Top             =   60
      Width           =   4575
   End
   Begin VB.CommandButton ResetBtn 
      Caption         =   "Reset"
      Height          =   350
      Left            =   6840
      TabIndex        =   7
      Top             =   1820
      Width           =   1125
   End
   Begin MSComctlLib.ImageList MouthImgList 
      Left            =   6840
      Top             =   3960
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.TextBox DebugTxtBox 
      BackColor       =   &H80000000&
      Height          =   1920
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   26
      Top             =   4592
      Width           =   7815
   End
   Begin VB.CommandButton SkipBtn 
      Caption         =   "Skip"
      Height          =   350
      Left            =   6840
      TabIndex        =   4
      Top             =   1380
      Width           =   500
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   350
      Left            =   7725
      TabIndex        =   6
      Top             =   1380
      Width           =   240
      _ExtentX        =   445
      _ExtentY        =   614
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "SkipTxtBox"
      BuddyDispid     =   196626
      OrigLeft        =   7560
      OrigTop         =   1800
      OrigRight       =   7800
      OrigBottom      =   2145
      Max             =   50
      Min             =   -50
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox SkipTxtBox 
      Height          =   350
      Left            =   7320
      TabIndex        =   5
      Text            =   "0"
      Top             =   1380
      Width           =   400
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   7560
      Top             =   4080
      _ExtentX        =   804
      _ExtentY        =   804
      _Version        =   393216
   End
   Begin VB.ComboBox FormatCB 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3705
      Width           =   3300
   End
   Begin MSComctlLib.Slider RateSldr 
      Height          =   320
      Left            =   848
      TabIndex        =   11
      ToolTipText     =   "Changes voice playback rate"
      Top             =   2870
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   550
      _Version        =   393216
      LargeChange     =   1
      Min             =   -10
      TickStyle       =   3
   End
   Begin VB.ComboBox VoiceCB 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2475
      Width           =   3300
   End
   Begin VB.CommandButton PauseBtn 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6840
      MaskColor       =   &H00808080&
      TabIndex        =   3
      Top             =   940
      Width           =   1125
   End
   Begin VB.CommandButton SpeakBtn 
      Caption         =   "Speak"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6840
      TabIndex        =   1
      Top             =   60
      Width           =   1125
   End
   Begin MSComctlLib.Slider VolumeSldr 
      Height          =   320
      Left            =   848
      TabIndex        =   13
      ToolTipText     =   "Changes voice playback volume"
      Top             =   3280
      Width           =   3296
      _ExtentX        =   5821
      _ExtentY        =   550
      _Version        =   393216
      Max             =   100
      SelStart        =   100
      TickStyle       =   3
      Value           =   100
   End
   Begin VB.Label Label1 
      Caption         =   "Audio Output"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   4095
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3735
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3330
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2925
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Voice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2505
      Width           =   495
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Begin VB.Menu menuFileOpenText 
         Caption         =   "Open Text File"
         Shortcut        =   ^O
      End
      Begin VB.Menu menuFileSpeakWave 
         Caption         =   "Speak Wave File"
         Shortcut        =   ^W
      End
      Begin VB.Menu menuFileSaveToWave 
         Caption         =   "Save To Wave File"
         Shortcut        =   ^S
      End
      Begin VB.Menu menuSep 
         Caption         =   "-"
      End
      Begin VB.Menu menuFileExit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "Help"
      Begin VB.Menu menuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "TTSAppMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WINAMP2
Dim CUT_ONCE
Dim VALUE_WRITE_OLD


Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long






'#### FIND HERE FOR DEBUG CODE USE
'Private Sub ShowEvent(ParamArray strArray())

'USEFUL VARS
'VoiceStreamActive = True
'VoiceStreamEnds



'=============================================================================
'
' This VB TTS App sample demonstrates most of the TTS functionalities
' supported in SAPI 5.1. The main object used here is SpVoice.
'
' Copyright @ 2001 Microsoft Corporation All Rights Reserved.
'=============================================================================

Option Explicit

Dim A$
' First, declare the main SAPI object we are using in this sample. It is
' created inside Form_Load and released inside Form_Unload.
Public WithEvents Voice As SpVoice
Attribute Voice.VB_VarHelpID = -1

' Speak flags is a combination of bit flags. These individual bits correspond
' to check boxes on the UI. So m_speakFlags should always be kept in sync
' with the state of those check boxes.
Public m_speakFlags As SpeechVoiceSpeakFlags

' This is the default format we will use.
Const DefaultFmt = "SAFT22kHz16BitMono"

' We will disable the output combo box and show this if there's no audio output.
Const NoAudioOutput = "No audio ouput object available"

' We will enable/disable menu items and buttons based on current state
' m_speaking indicates whether a speak task is in progress
' m_paused indicates whether Voice.Pause is called
Private m_bSpeaking As Boolean
Private m_bPaused As Boolean



Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long


Private Function GetUserName() As String
   Dim UserName As String * 255
   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
Private Function GetComputerName() As String
   Dim UserName As String * 255
   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function


Private Sub DebugTxtBox_Change()

'Private Sub ShowEvent(ParamArray strArray())

End Sub

Private Sub Form_Load()
    'On Error GoTo ErrHandler
    On Error Resume Next

    Dim A_VAR1$, FILENAME1, FILENAME2
    List1.AddItem "HELLO WORLD"
'    List1.AddItem "IDENTIFY! HIS, INSANE! TUNE, IN THE FLAT! BELOW!"
    
    ' Creates the voice object first
    Set Voice = New SpVoice
    
    ' Load the voices combo box
    Dim Token As ISpeechObjectToken

    For Each Token In Voice.GetVoices
        VoiceCB.AddItem (Token.GetDescription())
    Next
    
    'load the format combo box
    AddItemToFmtCB
    
    'set the default format
    'FormatCB.Text = DefaultFmt
    
    On Error Resume Next
    FILENAME1 = "VoiceRateSldr.txt"
    FILENAME2 = App.Path + "\TTSAppVB\" + GetComputerName + "-" + GetUserName + "\" + FILENAME1
    If Dir(FILENAME2) <> "" Then
        Open FILENAME2 For Input As #1
            Line Input #1, A_VAR1$
        Close #1
    End If
    If A_VAR1$ <> "" Then
        RateSldr.Value = Val(A_VAR1$)
        If RateSldr.Value = 0 Then RateSldr.Value = 5
        Call RateSldr_Scroll
    Else
        If RateSldr.Value = 0 Then RateSldr.Value = 5
        Call RateSldr_Scroll
    End If
        
    FILENAME1 = "VoiceAudioOutPutCB.txt"
    FILENAME2 = App.Path + "\TTSAppVB\" + GetComputerName + "-" + GetUserName + "\" + FILENAME1
    If Dir(FILENAME2) <> "" Then
        Open FILENAME2 For Input As #1
            Line Input #1, A_VAR1$
        Close #1
    End If
    If A_VAR1$ <> "" Then
        FormatCB.ListIndex = Val(A_VAR1$)
        Call FormatCB_Click
    Else
        'set the default format
        FormatCB.Text = DefaultFmt
    End If
    
    FILENAME1 = "VoiceCB.txt"
    FILENAME2 = App.Path + "\TTSAppVB\" + GetComputerName + "-" + GetUserName + "\" + FILENAME1
    If Dir(FILENAME2) <> "" Then
        Open FILENAME2 For Input As #1
            Line Input #1, A_VAR1$
        Close #1
    End If
    If A_VAR1$ <> "" Then
        VoiceCB.ListIndex = Val(A_VAR1$)
        Call VoiceCB_Click
    End If
    
    FILENAME1 = "Voice_Volume_Level_Slider.txt"
    FILENAME2 = App.Path + "\TTSAppVB\" + GetComputerName + "-" + GetUserName + "\" + FILENAME1
    If Dir(FILENAME2) <> "" Then
        Open FILENAME2 For Input As #1
            Line Input #1, A_VAR1$
        Close #1
    End If
    If A_VAR1$ <> "" Then
        VolumeSldr.Value = Val(A_VAR1$)
        Call VolumeSldr_Scroll
    Else
        VolumeSldr.Value = 100
        Call VolumeSldr_Scroll
    End If
    
    On Error GoTo 0

    
    
    ' Load the audio output combo box
    If Voice.GetAudioOutputs.Count > 0 Then
        For Each Token In Voice.GetAudioOutputs
            AudioOutputCB.AddItem (Token.GetDescription)
        Next
    Else
        AudioOutputCB.AddItem NoAudioOutput
        AudioOutputCB.Enabled = False
    End If
    AudioOutputCB.ListIndex = 0
    
    'load image list
    LoadMouthImages
    
    MouthImgList.MaskColor = vbMagenta
    MouthImgList.BackColor = GetSysColor(COLOR_3DFACE)
    Set VisemePicture.Picture = MouthImgList.Overlay("MICFULL", "MICFULL")
    
    ' init speak flags and sync flag check boxes
    m_speakFlags = SVSFlagsAsync Or SVSFPurgeBeforeSpeak Or SVSFIsXML
    chkSpFlagAync.Value = Checked
    chkSpFlagPurgeBeforeSpeak.Value = Checked
    chkSpFlagIsXML.Value = Checked
    
    SetSpeakingState False, False
    
    
    
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error in initialization: " & vbCrLf & vbCrLf & Err.Description & _
        vbCrLf & vbCrLf & "Shutting down.", vbOKOnly, "TTSApp"
    Set Voice = Nothing
        
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If aa_MainCode.ALL_FORM_REQUEST_TO_END = False Then
        Cancel = True: Me.Hide: Exit Sub
    End If
    
    If aa_MainCode.TrueTerminate = False Then
        Cancel = True: TTSAppMain.Visible = False: Exit Sub
    End If
    Set Voice = Nothing
End Sub

Private Sub AudioOutputCB_Click()
    On Error GoTo ErrHandler
    
    ' change the output to the selected one
    Set Voice.AudioOutput = Voice.GetAudioOutputs().Item(AudioOutputCB.ListIndex)
    ' changing output may have also changed the format, so call function
    ' FormatCB_Click to make sure we are using the format as selected
    FormatCB_Click
    Exit Sub
    
ErrHandler:
    AddDebugInfo "Set audio output error: ", Err.Description
End Sub


Sub WRITE_DATA_vAR_STORE(FILENAME2, VALUE_WRITE)

    If VALUE_WRITE_OLD = VALUE_WRITE Then Exit Sub
    VALUE_WRITE_OLD = VALUE_WRITE

    Dim i, FR1, PATH_NAME1
    '------------------------------------------
    '------------------------------------------
    
    PATH_NAME1 = Mid(FILENAME2, 1, InStrRev(FILENAME2, "\")) 'App.Path + "\TTSAppVB\" + GetComputerName + "-" + GetUserName + ""
    
    If Dir(PATH_NAME1, vbDirectory) = "" Then
        i = CreateFolderTree(PATH_NAME1)
        If i = False Then Exit Sub
    End If
    
    If FileInUse(FILENAME2) = True Then Exit Sub
    
    If IsIDE And CUT_ONCE = True Then Stop
    
    On Error Resume Next
    FR1 = FreeFile
    Open FILENAME2 For Output As #FR1
        Print #FR1, str(VALUE_WRITE)
    Close #FR1
    
End Sub


Private Sub FormatCB_Click()
    On Error GoTo ErrHandler
    
    ' Note: AllowAudioOutputFormatChangesOnNextSet is a hidden property, VB
    ' object browser doesn't show it by default. To see it, you can go to
    ' VB object viewer, right click and turn on the "show hidden members".
    
    Voice.AllowAudioOutputFormatChangesOnNextSet = True ' False
    
    ' The format Type is associated with the selected list item as a long.
    
    'A = FormatCB.ListCount
    
    If FormatCB.ListIndex < 0 Then FormatCB.ListIndex = 0
    Voice.AudioOutputStream.Format.Type = FormatCB.ItemData(FormatCB.ListIndex)
    
    'Voice.AudioOutputStream.Format.Type = FormatCB.ItemData(Val(A$))
    
    ' Currently you have to call this so that SAPI picks up the new format.
    Set Voice.AudioOutputStream = Voice.AudioOutputStream
    
    Dim FILENAME2
    
    '------------------------------------------------------------------------------------------------------
    FILENAME2 = App.Path + "\TTSAppVB\" + GetComputerName + "-" + GetUserName + "\VoiceAudioOutPutCB.txt"
    'Call WRITE_DATA_vAR_STORE(FILENAME2, FormatCB.List(FormatCB.ListIndex))
    Call WRITE_DATA_vAR_STORE(FILENAME2, FormatCB.ListIndex)
    '------------------------------------------------------------------------------------------------------
    
    Exit Sub
    
ErrHandler:
    AddDebugInfo "Set format error: ", Err.Description
    
End Sub

Private Sub List1_Click()
MainTxtBox2 = List1.List(List1.ListIndex)
'Call SpeakBtn_Click
End Sub

Private Sub MainTxtBox2_Click()
If MainTxtBox2 <> "" Then List1.AddItem MainTxtBox2
End Sub

Private Sub menuAbout_Click()
    MsgBox "TTSApp" & vbCrLf & vbCrLf & "Copyright (c) 2001 Microsoft Corporation. All rights reserved.", _
        vbOKOnly Or vbInformation, "About TTSApp"
End Sub

Private Sub menuFileExit_Click()
    Unload TTSAppMain
    End
End Sub

Private Sub menuFileOpenText_Click()
    
    Dim sLocation As String
    
    ' Set CancelError is True
    ComDlg.CancelError = True
    On Error GoTo ErrHandler
        
    ' Set flags
    ComDlg.flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
    ' Set Dialog title
    ComDlg.DialogTitle = "Open a Text File"
    ' Set open directory
    sLocation = GetDirectory()
    If Len(sLocation) <> 0 Then
        ComDlg.InitDir = sLocation
    End If
    
    ' Set filters
    ComDlg.Filter = "All Files (*.*)|*.*|Text, XML Files " & "(*.txt;*.xml)|*.txt;*.xml"
 
    ' Specify default filter
    ComDlg.FilterIndex = 2
    ' Display the Open dialog box
    ComDlg.ShowOpen
    
    ' Now open the text file and open it in the text box.
    ' We only support text files encoded with the system code page as the
    ' binary to unicode conversion in VB is using system code page.
    Open ComDlg.Filename For Binary Access Read As 1
    MainTxtBox.Text = StrConv(InputB$(LOF(1), 1), vbUnicode)
    Close #1
    
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button, do not show error
    If Not (Err.Number = 32755) Then
        AddDebugInfo "Open file: ", Err.Description
    End If
End Sub

Private Sub menuFileSaveToWave_Click()
    ' Set CancelError is True
    ComDlg.CancelError = True
    On Error GoTo ErrHandler

    ' Set flags
    ComDlg.flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNNoReadOnlyReturn
    ' Set Dialog title
    ComDlg.DialogTitle = "Save to a Wave File"
    ' Set filters
    ComDlg.Filter = "All Files (*.*)|*.*|Wave Files " & "(*.wav)|*.wav"
    ' Specify default filter
    ComDlg.FilterIndex = 2
    ' Display the Open dialog box
    ComDlg.ShowSave
    
    ' create a wave stream
    Dim cpFileStream As New SpFileStream
    
    ' Set output format to selected format
    cpFileStream.Format.Type = FormatCB.ItemData(FormatCB.ListIndex)
    
    ' Open the file for write
    cpFileStream.Open ComDlg.Filename, SSFMCreateForWrite, False
    
    ' Set output stream to the file stream
    Voice.AllowAudioOutputFormatChangesOnNextSet = False
    Set Voice.AudioOutputStream = cpFileStream
    
    ' show action
    AddDebugInfo "Save to .wav file"
    
    ' speak the given text with given flags
    Voice.Speak MainTxtBox.Text, m_speakFlags
    
    ' wait until it's done speaking with a really really long timeout.
    ' the tiemout value is in unit of millisecond. -1 means forever.
    Voice.WaitUntilDone -1
    
    ' Since the output stream was set to the file stream, we need to
    ' set back to the selected audio output by calling AudioOutputCB_Click
    ' as if user just changed it through UI
    AudioOutputCB_Click
    
    ' close the file stream
    cpFileStream.Close
    Set cpFileStream = Nothing
    
    MsgBox "WAV file successfully written!", vbOKOnly, "File Saved"
    Exit Sub

ErrHandler:
    'User pressed the Cancel button, do not show error
    If Not (Err.Number = 32755) Then
        AddDebugInfo "Save to Wave file Error: ", Err.Description
    End If
    
    If Not cpFileStream Is Nothing Then
        Set cpFileStream = Nothing
    End If
End Sub

Private Sub menuFileSpeakWave_Click()
    ' Set CancelError is True
    ComDlg.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    ComDlg.flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
    ' Set Dialog title
    ComDlg.DialogTitle = "Speak a Wave File"
    ' Set filters
    ComDlg.Filter = "All Files (*.*)|*.*|Wave Files " & "(*.wav)|*.wav"
    ' Specify default filter
    ComDlg.FilterIndex = 2
    ' Display the Open dialog box
    ComDlg.ShowOpen

    AddDebugInfo "Speak .wav file"
    
    ' Speak the contents of the wavefile. Notice here we are passing in the
    ' file name so the filename flag is set.
    MainTxtBox.Text = ComDlg.Filename
    chkSpFlagIsFilename.Value = Checked
    SpeakBtn_Click
    
    Exit Sub

ErrHandler:
    'User pressed the Cancel button, do not show error
    If Not (Err.Number = 32755) Then
        AddDebugInfo "Speak Wave Error: ", Err.Description
    End If
    
    SetSpeakingState False, m_bPaused
    Exit Sub
End Sub

Private Sub PauseBtn_Click()
    Select Case PauseBtn.Caption
    Case "Pause"
        AddDebugInfo "Pause"
        Voice.Pause
        SetSpeakingState m_bSpeaking, True
    
    Case "Resume"
        AddDebugInfo "Resume"
        Voice.Resume
        SetSpeakingState m_bSpeaking, False
    End Select
End Sub

Public Sub RateSldr_Scroll()
    
    'TTSAppMain.RateSldr
    Voice.Rate = RateSldr.Value
    
    Dim FILENAME2
    
    '------------------------------------------------------------------------------------------------------
    FILENAME2 = App.Path + "\TTSAppVB\" + GetComputerName + "-" + GetUserName + "\VoiceRateSldr.txt"
    Call WRITE_DATA_vAR_STORE(FILENAME2, RateSldr.Value)
    '------------------------------------------------------------------------------------------------------
    

End Sub

Private Sub ResetBtn_Click()
    'set output to default
    AudioOutputCB.ListIndex = 0
    Set Voice.AudioOutput = Nothing
    'use default voice
    VoiceCB.ListIndex = 0
    
    'Format to default
    FormatCB.Text = DefaultFmt
    
    'reset main text field
    MainTxtBox.Text = "Enter text you wish spoken here."
    
    'reset volume and rate
    VolumeSldr.Value = 100
    VolumeSldr_Scroll
    
    'RateSldr.Value = 0
    RateSldr_Scroll
    
    ' reset speak flags
    m_speakFlags = SVSFlagsAsync Or SVSFPurgeBeforeSpeak Or SVSFIsXML
    chkSpFlagAync.Value = Checked
    chkSpFlagPurgeBeforeSpeak.Value = Checked
    chkSpFlagIsXML.Value = Checked
    chkSpFlagIsFilename.Value = Unchecked
    chkSpFlagNLPSpeakPunc.Value = Unchecked
    chkSpFlagPersistXML.Value = Unchecked
    
    'reset DebugTxtbox text
    DebugTxtBox.Text = Empty
    
    'reset skip text box
    SkipTxtBox.Text = "0"
    Set VisemePicture.Picture = MouthImgList.Overlay("MICFULL", "MICFULL")
    
    ' if it's paused, call Resume to reset state
    If m_bPaused Then
        Voice.Resume
    End If
    SetSpeakingState False, False
End Sub

Private Sub SkipBtn_Click()
    On Error GoTo ErrHandler
    Dim SkipType As String
    Dim SkipNum As Integer
    
    AddDebugInfo "Skip"
    
    ' skip by the number specified
    SkipNum = SkipTxtBox.Text
    SkipType = "Sentence"
    
    Voice.skip SkipType, SkipNum
    Exit Sub
    
ErrHandler:
    'MsgBox Err.Description & ":" & Err.Number, vbOKOnly, "Skip Error"
    AddDebugInfo "Skip Error: ", Err.Description
    Exit Sub
End Sub


Private Sub SpeakBtn_Click()
    
    Call ATidalX.MAKETALKTIMESTRING
    
    TALKTIME$ = TALKTIME$ + " And " + str(Second(Now)) + " Seconds,, "
        
    FORCE_TALKER_GO_WITHOUT_CHECKER_DUPE_IGNOR = True
    If MainTxtBox2 = "" Then
        Call ATidalX.SpeakVoice(TALKTIME$)
    End If
    Call ATidalX.SpeakVoice(MainTxtBox2)
    

    Exit Sub
    
    On Error GoTo ErrHandler
    AddDebugInfo ("Speak")
    
    ' exit if there's nothing to speak
    If MainTxtBox2.Text = "" Then
        Exit Sub
    End If
    
    ' If it's paused and some text still remains to be spoken, Speak button
    ' acts the same as Resume button. However a programmer can choose to
    ' speak from the beginning again or any other behavior.
    ' In other cases, we speak the text with given flags.
    If Not (m_bPaused And m_bSpeaking) Then
        ' just speak the text with the given flags
        Voice.Speak MainTxtBox2.Text, m_speakFlags
    End If
    
    ' Resume if Voice is paused
    If m_bPaused Then Voice.Resume
    
    ' set the state of menu items and buttons
    SetSpeakingState True, False
    Exit Sub
    
ErrHandler:
    AddDebugInfo "Speak Error: ", Err.Description
    SetSpeakingState False, m_bPaused
End Sub

Private Sub StopBtn_Click()
    On Error GoTo ErrHandler
    AddDebugInfo ("Stop")
    
    ' when string to speak is NULL and dwFlags is set to SPF_PURGEBEFORESPEAK
    ' it indicates to SAPI that any remaining data to be synthesized should
    ' be discarded.
    Voice.Speak vbNullString, SVSFPurgeBeforeSpeak
    If m_bPaused Then Voice.Resume
    
    SetSpeakingState False, False
    Exit Sub
    
ErrHandler:
    AddDebugInfo "Speak Error: ", Err.Description
End Sub

Private Sub Voice_AudioLevel(ByVal StreamNumber As Long, _
                             ByVal StreamPosition As Variant, _
                             ByVal AudioLevel As Long)
    ShowEvent "AudioLevel", "StreamNumber=" & StreamNumber, _
            "StreamPosition=" & StreamPosition, "AudioLevel=" & AudioLevel
End Sub

Private Sub Voice_Bookmark(ByVal StreamNumber As Long, _
                           ByVal StreamPosition As Variant, _
                           ByVal Bookmark As String, _
                           ByVal BookmarkId As Long)
    ShowEvent "BookMark", "StreamNumber=" & StreamNumber, _
            "StreamPosition=" & StreamPosition, "Bookmark=" & Bookmark, _
            "BookmarkId=" & BookmarkId
End Sub

Private Sub Voice_EndStream(ByVal StreamNum As Long, ByVal StreamPos As Variant)
    ShowEvent "EndStream", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos

    ' select all text to indicate that we are done
    HighLightSpokenWords 0, Len(MainTxtBox.Text)
    
    ' reset the mouth
    Set VisemePicture.Picture = MouthImgList.Overlay("MICFULL", "MICFULL")
    
    ' reset the state of buttons, checkboxes and menu items
    SetSpeakingState False, m_bPaused
End Sub

Private Sub Voice_EnginePrivate(ByVal StreamNumber As Long, _
                                ByVal StreamPosition As Long, _
                                ByVal lParam As Variant)
    ShowEvent "EnginePrivate", "StreamNumber=" & StreamNumber, _
            "StreamPosition=" & StreamPosition, "lParam=" & lParam
End Sub

Private Sub Voice_Phoneme(ByVal StreamNumber As Long, _
                          ByVal StreamPosition As Variant, _
                          ByVal duration As Long, _
                          ByVal NextPhoneId As Integer, _
                          ByVal Feature As SpeechLib.SpeechVisemeFeature, _
                          ByVal CurrentPhoneId As Integer)
    ShowEvent "Phoneme", "StreamNumber=" & StreamNumber, _
            "StreamPosition=" & StreamPosition, "NextPhoneId=" & NextPhoneId, _
            "Feature=" & Feature, "CurrentPhoneId=" & CurrentPhoneId
End Sub

Private Sub Voice_Sentence(ByVal StreamNum As Long, _
                           ByVal StreamPos As Variant, _
                           ByVal pos As Long, _
                           ByVal Length As Long)
    ShowEvent "Sentence", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos, _
            "Pos=" & pos, "Length=" & Length
End Sub

Private Sub Voice_StartStream(ByVal StreamNum As Long, ByVal StreamPos As Variant)
    ShowEvent "StartStream", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos
    
    ' reset the state of buttons, checkboxes and menu items
    SetSpeakingState True, m_bPaused
End Sub

Private Sub Voice_Viseme(ByVal StreamNum As Long, _
                         ByVal StreamPos As Variant, _
                         ByVal duration As Long, _
                         ByVal VisemeType As SpeechVisemeType, _
                         ByVal Feature As SpeechVisemeFeature, _
                         ByVal VisemeId As Long)
    
    ShowEvent "Viseme", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos, _
            "Duration=" & duration, "VisemeType=" & VisemeType, _
            "Feature=" & Feature, "VisemeId=" & VisemeId
    
    ' Here we are going to show different mouth positions according to the viseme.
    ' The picture we show doesn't necessarily match the real mouth position.
    ' Just trying to make it more interesting.
    If VisemeId = 0 Then
        VisemeId = VisemeId + 1
    End If
    Set VisemePicture.Picture = MouthImgList.Overlay("MICFULL", VisemeId)
    If (VisemeId Mod 6 = 2) Then
        Set VisemePicture.Picture = MouthImgList.Overlay("MICFULL", "MICEYECLOSED")
    Else
        If (VisemeId Mod 6 = 5) Then
            Set VisemePicture.Picture = MouthImgList.Overlay("MICFULL", "MICEYENARROW")
        End If
    End If
End Sub

Private Sub Voice_VoiceChange(ByVal StreamNum As Long, _
                              ByVal StreamPos As Variant, _
                              ByVal Token As SpeechLib.ISpeechObjectToken)
    
    ShowEvent "VoiceChange", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos, _
            "Token=" & Token.GetDescription
    
    ' Let's sync up the combo box with the new value
    Dim i As Long
    For i = 0 To VoiceCB.ListCount - 1
        If VoiceCB.List(i) = Token.GetDescription() Then
            VoiceCB.ListIndex = i
            Exit For
        End If
    Next
    
    
End Sub

Private Sub Voice_Word(ByVal StreamNum As Long, _
                       ByVal StreamPos As Variant, _
                       ByVal pos As Long, _
                       ByVal Length As Long)
                       
    ShowEvent "Word", "StreamNum=" & StreamNum, "StreamPos=" & StreamPos, _
            "Pos=" & pos, "Length=" & Length
    
    'Debug.Print pos, Length, MainTxtBox.SelStart, MainTxtBox.SelLength
    
    ' Select the word that's currently being spoken.
    HighLightSpokenWords pos, Length
End Sub

Private Sub VoiceCB_Click()
    ' change the voice to the selected one
    Set Voice.Voice = Voice.GetVoices().Item(VoiceCB.ListIndex)
    
    Dim FILENAME2
    
    '------------------------------------------------------------------------------------------------------
    FILENAME2 = App.Path + "\TTSAppVB\" + GetComputerName + "-" + GetUserName + "\VoiceCB.txt"
    Call WRITE_DATA_vAR_STORE(FILENAME2, VoiceCB.ListIndex)
    '------------------------------------------------------------------------------------------------------
    SAY_ONCE_MOST_LIKELY_A_INCORECT_VOICE_DRIVER = False
End Sub

Private Sub VolumeSldr_Scroll()
    
    Voice.Volume = VolumeSldr.Value
    
    Dim FILENAME2
    
    '------------------------------------------------------------------------------------------------------
    FILENAME2 = App.Path + "\TTSAppVB\" + GetComputerName + "-" + GetUserName + "\Voice_Volume_Level_Slider.txt"
    Call WRITE_DATA_vAR_STORE(FILENAME2, VolumeSldr.Value)
    '------------------------------------------------------------------------------------------------------

End Sub

' The following functions are simply to sync up the speak flags.
' When the check box is checked, the corresponding bit is set in the flags.
Private Sub chkSpFlagAync_Click()
    m_speakFlags = SetOrClearFlag(chkSpFlagAync.Value, m_speakFlags, SVSFlagsAsync)
End Sub

Private Sub chkSpFlagIsFilename_Click()
    m_speakFlags = SetOrClearFlag(chkSpFlagIsFilename.Value, m_speakFlags, SVSFIsFilename)
End Sub

Private Sub chkSpFlagIsXML_Click()
    ' Note: special case here. There are two flags,SVSFIsXML and SVSFIsNotXML.
    ' When neither is set, SAPI will guess by peeking at beginning characters.
    ' In this sample, we explicitly set one of them.
    
    If chkSpFlagIsXML.Value = 0 Then
        ' clear SVSFIsXML bit and set SVSFIsNotXML bit
        m_speakFlags = m_speakFlags And Not SVSFIsXML
        m_speakFlags = m_speakFlags Or SVSFIsNotXML
    Else
        ' clear SVSFIsNotXML bit and set SVSFIsXML bit
        m_speakFlags = m_speakFlags And Not SVSFIsNotXML
        m_speakFlags = m_speakFlags Or SVSFIsXML
    End If
End Sub

Private Sub chkSpFlagNLPSpeakPunc_Click()
    m_speakFlags = SetOrClearFlag(chkSpFlagNLPSpeakPunc.Value, m_speakFlags, SVSFNLPSpeakPunc)
End Sub

Private Sub chkSpFlagPersistXML_Click()
    m_speakFlags = SetOrClearFlag(chkSpFlagPersistXML.Value, m_speakFlags, SVSFPersistXML)
End Sub

Private Sub chkSpFlagPurgeBeforeSpeak_Click()
    m_speakFlags = SetOrClearFlag(chkSpFlagPurgeBeforeSpeak.Value, m_speakFlags, SVSFPurgeBeforeSpeak)
End Sub


Private Sub AddFmts(ByRef Name As String, ByVal fmt As SpeechAudioFormatType)
    Dim index As String
    
    ' get the count of existing list so that we are adding to the bottom of the list
    index = FormatCB.ListCount
    
    ' add the name to the list box and associate the format type with the item
    FormatCB.AddItem Name, index
    FormatCB.ItemData(index) = fmt
End Sub

Private Sub AddItemToFmtCB()
    AddFmts "SAFT8kHz8BitMono", SAFT8kHz16BitMono
    AddFmts "SAFT8kHz8BitStereo", SAFT8kHz8BitStereo
    AddFmts "SAFT8kHz16BitMono", SAFT8kHz16BitMono
    AddFmts "SAFT8kHz16BitStereo", SAFT8kHz16BitStereo
    
    AddFmts "SAFT11kHz8BitMono", SAFT11kHz8BitMono
    AddFmts "SAFT11kHz8BitStereo", SAFT11kHz8BitStereo
    AddFmts "SAFT11kHz16BitMono", SAFT11kHz16BitMono
    AddFmts "SAFT11kHz16BitStereo", SAFT11kHz16BitStereo
    
    AddFmts "SAFT12kHz8BitMono", SAFT12kHz8BitMono
    AddFmts "SAFT12kHz8BitStereo", SAFT12kHz8BitStereo
    AddFmts "SAFT12kHz16BitMono", SAFT12kHz16BitMono
    AddFmts "SAFT12kHz16BitStereo", SAFT12kHz16BitStereo
    
    AddFmts "SAFT16kHz8BitMono", SAFT16kHz8BitMono
    AddFmts "SAFT16kHz8BitStereo", SAFT16kHz8BitStereo
    AddFmts "SAFT16kHz16BitMono", SAFT16kHz16BitMono
    AddFmts "SAFT16kHz16BitStereo", SAFT16kHz16BitStereo
    
    AddFmts "SAFT22kHz8BitMono", SAFT22kHz8BitMono
    AddFmts "SAFT22kHz8BitStereo", SAFT22kHz8BitStereo
    AddFmts "SAFT22kHz16BitMono", SAFT22kHz16BitMono
    AddFmts "SAFT22kHz16BitStereo", SAFT22kHz16BitStereo
    
    AddFmts "SAFT24kHz8BitMono", SAFT24kHz8BitMono
    AddFmts "SAFT24kHz8BitStereo", SAFT24kHz8BitStereo
    AddFmts "SAFT24kHz16BitMono", SAFT24kHz16BitMono
    AddFmts "SAFT24kHz16BitStereo", SAFT24kHz16BitStereo
    
    AddFmts "SAFT32kHz8BitMono", SAFT32kHz8BitMono
    AddFmts "SAFT32kHz8BitStereo", SAFT32kHz8BitStereo
    AddFmts "SAFT32kHz16BitMono", SAFT32kHz16BitMono
    AddFmts "SAFT32kHz16BitStereo", SAFT32kHz16BitStereo
    
    AddFmts "SAFT44kHz8BitMono", SAFT44kHz8BitMono
    AddFmts "SAFT44kHz8BitStereo", SAFT44kHz8BitStereo
    AddFmts "SAFT44kHz16BitMono", SAFT44kHz16BitMono
    AddFmts "SAFT44kHz16BitStereo", SAFT44kHz16BitStereo
    
    AddFmts "SAFT48kHz8BitMono", SAFT48kHz8BitMono
    AddFmts "SAFT48kHz8BitStereo", SAFT48kHz8BitStereo
    AddFmts "SAFT48kHz16BitMono", SAFT48kHz16BitMono
    AddFmts "SAFT48kHz16BitStereo", SAFT48kHz16BitStereo
End Sub
Private Sub LoadMouthImages()
    On Error GoTo ErrHandler
    
    MouthImgList.ListImages.Add 1, "MICFULL", LoadResPicture("MICFULL", vbResBitmap)
    MouthImgList.ListImages.Add 2, , LoadResPicture("MIC11", vbResBitmap)
    MouthImgList.ListImages.Add 3, , LoadResPicture("MIC11", vbResBitmap)
    MouthImgList.ListImages.Add 4, , LoadResPicture("MIC11", vbResBitmap)
    MouthImgList.ListImages.Add 5, , LoadResPicture("MIC10", vbResBitmap)
    MouthImgList.ListImages.Add 6, , LoadResPicture("MIC11", vbResBitmap)
    MouthImgList.ListImages.Add 7, , LoadResPicture("MIC9", vbResBitmap)
    MouthImgList.ListImages.Add 8, , LoadResPicture("MIC2", vbResBitmap)
    MouthImgList.ListImages.Add 9, , LoadResPicture("MIC13", vbResBitmap)
    MouthImgList.ListImages.Add 10, , LoadResPicture("MIC9", vbResBitmap)
    MouthImgList.ListImages.Add 11, , LoadResPicture("MIC12", vbResBitmap)
    MouthImgList.ListImages.Add 12, , LoadResPicture("MIC11", vbResBitmap)
    MouthImgList.ListImages.Add 13, , LoadResPicture("MIC9", vbResBitmap)
    MouthImgList.ListImages.Add 14, , LoadResPicture("MIC3", vbResBitmap)
    MouthImgList.ListImages.Add 15, , LoadResPicture("MIC6", vbResBitmap)
    MouthImgList.ListImages.Add 16, , LoadResPicture("MIC7", vbResBitmap)
    MouthImgList.ListImages.Add 17, , LoadResPicture("MIC8", vbResBitmap)
    MouthImgList.ListImages.Add 18, , LoadResPicture("MIC5", vbResBitmap)
    MouthImgList.ListImages.Add 19, , LoadResPicture("MIC4", vbResBitmap)
    MouthImgList.ListImages.Add 20, , LoadResPicture("MIC7", vbResBitmap)
    MouthImgList.ListImages.Add 21, , LoadResPicture("MIC9", vbResBitmap)
    MouthImgList.ListImages.Add 22, , LoadResPicture("MIC11", vbResBitmap)
    MouthImgList.ListImages.Add 23, "MICEYECLOSED", LoadResPicture("MICEYECLOSED", vbResBitmap)
    MouthImgList.ListImages.Add 24, "MICEYENARROW", LoadResPicture("MICEYENARROW", vbResBitmap)
    
    Exit Sub
ErrHandler:
    MsgBox Err.Description & ":" & Err.Number, vbOKOnly, "Load Images Error"
End Sub

Private Sub AddDebugInfo(DebugStr As String, Optional error As String = Empty)
    ' This function adds debug string to the info window.
    
    ' First of all, let's delete a few charaters if the text box is about to
    ' overflow. In this sample we are using the default limit of charaters.
    If Len(DebugTxtBox.Text) > 64000 Then
        'Debug.Print "Too much stuff in the debug window. Remove first 10K chars"
        DebugTxtBox.SelStart = 0
        DebugTxtBox.SelLength = 10240
        DebugTxtBox.SelText = ""
    End If
    
    On Error Resume Next
    ' append the string to the DebugTxtBox text box and add a newline
'    DebugTxtBox.SelStart = Len(DebugTxtBox.Text)
    DebugTxtBox.SelText = DebugStr & error & vbCrLf
End Sub

Private Sub ShowEvent(ParamArray strArray())
    
    Dim strText As String
    
    Dim MSGRESULT1
    
    
    'TURN WINAMP VOL DOWN
    WINAMP2 = FindWindow("Winamp v1.x", vbNullString)
    If WINAMP2 > 0 Then
        MSGRESULT1 = SendMessage(WINAMP2, WM_USER, 0, ByVal ipc_isplaying)
        If MSGRESULT1 = 1 And 333 = 222 Then
            VolumeSldr.Value = 0
            Call VolumeSldr_Scroll
            Exit Sub
        End If
    End If
    
    'WHY HERE TAKE OUT 22-03-10
    'VolumeSldr.Value = 100
    'Call VolumeSldr_Scroll

    
    ' we will only show the events if the ShowEvents box is checked
    If chkShowEvents.Value = Checked Then
        strText = Join(strArray, ", ")
        AddDebugInfo "  Event: " & strText
    End If

    strText = Join(strArray, ", ")
    
    
    'BYE
    'ATidalX.VoiceStreamActiveMAKESUREOFF_TIMER.Interval = 0
    'ATidalX.VoiceStreamActiveMAKESUREOFF_TIMER.Interval = 1000
    'ATidalX.VoiceStreamActiveMAKESUREOFF_TIMER = True
    
    'Debug.Print "INTERVAL USED"
    
    
    If InStr(strText, "StartStream,") > 0 Then
        
        'THIS ALSO PREMATURE SET IN VOICE SPEAK SUB NOW
        'VoiceStreamActive = True
        
        'VoiceStreamEnds = False
 
        VoiceStreamActive = True
        
        
        ATidalX.TIMER_VoiceStreamActiveMAKESUREOFF.Enabled = True
               
        
        'STRING_OF_VOICE_IN_QUE_PLAY > 0
        
        
        'ATidalX.TimerSimulateStringVoice.Enabled = False
        'ATidalX.Timer1.Enabled = False
        'ATidalX.Timer_CountD1Text.Enabled = False
        
        
        'ATidalX.Text2 = "VOICE TIMERS "
        'Debug.Print Time$ + " -- " + "VOICE TIMERS ENABLED=FALSE"
    
        'NEW CODE
        'MP3 OFF -- ON START VOICE
        
        
        'DATE OLD MACH NOTE
        'THIS IS IN NEW PLACE IN THE NEW SPEAK SUB
        'SAFE MESSURE CHK
        'Dim MsgResult As Long
        
        
        ' THIS IS NOW LOCATE BEFORE VOICE --
        ' ADD TIME FOR CPU LOAD WINAMP
        ' @
        
        'MsgResult = ATidalX.IsWinAmpPlayingBeforeVoiceStartAndStop
        'WINAMPBackOnAfterVoiceRequired = False
        'If MsgResult = 1 Then
        '    Call ATidalX.MP3_OFF__ON_START_VOICE
        '    Call ATidalX.TimerDoubleVolume_Timer
        '    WINAMPBackOnAfterVoiceRequired = True
        'End If
    
    End If
    
    
    '## VOICE ENDS
    If InStr(strText, "EndStream,") > 0 Then
        
        'VoiceStreamActive = False
        
        'VoiceStreamEnds = True
        
        VoiceStreamActive = False
        
               
        'ATidalX.TIMER_VoiceStreamActiveMAKESUREOFF.Enabled =
        
        
        'Debug.Print "VOICE END USED"
        
        
        'BYE
        'ATidalX.VoiceStreamActiveMAKESUREOFF_TIMER.Enabled = False

        'DONT KNOW WHY THIS HERE TAKEN OUT -- CHKED SEE NONE REALLY
        'ATidalX.TimerSimulateStringVoice.Enabled = True
        
        If StringVoiceToPLayActive = False Then
            'ATidalX.Timer1.Enabled = True
            'ATidalX.Timer_CountD1Text.Enabled = True
        End If
        
        'If ATidalX.VoiceMP3Start_TIMER.Enabled = True Then Call ATidalX.VoiceMP3Start_TIMER_timer
        'Debug.Print Time$ + " -- " + "VOICE TIMERS ENABLED=TRUE"
        
        'NEW CODE
        'MP3 ON AFTER VOICE
        
        
        
        'If STRING_OF_VOICE_IN_QUE_PLAY = False Then
'            Call ATidalX.MP3_ON_AFTER_VOICE_FINISHES
'            Call ATidalX.TimerDoubleVolume_Timer
        
        Call ATidalX.VOICELISTTIMER_TIMER
        
        
        'End If
        
        
        'Call ATidalX.MP3_OFF__ON_START_VOICE
        
    End If
    


End Sub

Private Sub HighLightSpokenWords(ByVal pos As Long, ByVal Length As Long)
    On Error GoTo ErrHandler

    ' Only high light when the MainTxtBox is actually showing the spoken text,
    ' instead of file name
    If chkSpFlagIsFilename.Value = Unchecked Then
        MainTxtBox.SelStart = pos
        MainTxtBox.SelLength = Length
    End If
    
    Exit Sub
    
ErrHandler:
    AddDebugInfo "Failed to high light words. This may be caused by too many charaters in the main text box."
End Sub

' This following helper function will set or clear a bit (flag) in the given
' integer (base) according to the condition (cond). If cond is 0, the bit
' is cleared. Otherwise, the bit is set. The resulting integer is returned.
Private Function SetOrClearFlag(ByVal cond As Long, _
                                ByVal base As Long, _
                                ByVal Flag As Long) As Long
    
    If cond = 0 Then
        ' the condition is false, clear the flag
        SetOrClearFlag = base And Not Flag
    Else
        ' the condition is false, set the flag
        SetOrClearFlag = base Or Flag
    End If
End Function

Private Sub SetSpeakingState(ByVal bSpeaking As Boolean, ByVal bPaused As Boolean)
    ' change state of menu items and buttons accordingly
    menuFileOpenText.Enabled = Not bSpeaking
    menuFileSpeakWave.Enabled = Not bSpeaking
    menuFileSaveToWave.Enabled = Not bSpeaking
    
    SpeakBtn.Enabled = True
    
    StopBtn.Enabled = bSpeaking
    SkipBtn.Enabled = (bSpeaking And Not bPaused)
    PauseBtn.Enabled = bSpeaking
    
    If bPaused Then
        PauseBtn.Caption = "Resume"
    Else
        PauseBtn.Caption = "Pause"
    End If
    
    m_bSpeaking = bSpeaking
    m_bPaused = bPaused
End Sub

Public Function GetDirectory() As String

    Err.Clear

    On Error GoTo ErrHandler

    Dim DataKey As ISpeechDataKey
    Dim Category As New SpObjectTokenCategory
    
    'Get the sdk installation location from the registry
    'The value is under "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech". The string name is SDKPath"
    Category.SetId SpeechRegistryLocalMachineRoot
    Set DataKey = Category.GetDataKey
    GetDirectory = DataKey.GetStringValue("SDKPath")
    GetDirectory = GetDirectory + "samples\common"
    
    
    
ErrHandler:
    If Err.Number <> 0 Then
        GetDirectory = ""
    End If
End Function
