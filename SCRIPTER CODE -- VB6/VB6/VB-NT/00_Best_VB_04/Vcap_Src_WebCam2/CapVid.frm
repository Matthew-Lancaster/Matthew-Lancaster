VERSION 5.00
Begin VB.Form frmCapVid 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capture Video Sequence"
   ClientHeight    =   2550
   ClientLeft      =   345
   ClientTop       =   1545
   ClientWidth     =   3645
   Icon            =   "CapVid.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Com&press..."
      Height          =   375
      Index           =   4
      Left            =   2505
      TabIndex        =   10
      Top             =   2010
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Video..."
      Height          =   375
      Index           =   3
      Left            =   2505
      TabIndex        =   9
      Top             =   1590
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Audio..."
      Height          =   375
      Index           =   2
      Left            =   2505
      TabIndex        =   8
      Top             =   1170
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2505
      TabIndex        =   7
      Top             =   540
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   2505
      TabIndex        =   6
      Top             =   120
      Width           =   1005
   End
   Begin VB.CheckBox chkAudio 
      Caption         =   "Capture Audio"
      Height          =   270
      Left            =   105
      TabIndex        =   5
      Top             =   1575
      Width           =   1890
   End
   Begin VB.TextBox txtSec 
      Height          =   315
      Left            =   1650
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "30"
      Top             =   855
      Width           =   630
   End
   Begin VB.CheckBox chkLimit 
      Caption         =   "Enable capture time limit"
      Height          =   255
      Left            =   105
      TabIndex        =   2
      Top             =   600
      Width           =   2160
   End
   Begin VB.TextBox txtFps 
      Height          =   315
      Left            =   1650
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "15"
      Top             =   150
      Width           =   630
   End
   Begin VB.Label lblStaticText 
      Caption         =   "Seconds:"
      Height          =   255
      Index           =   1
      Left            =   390
      TabIndex        =   3
      Top             =   915
      Width           =   810
   End
   Begin VB.Label lblStaticText 
      Caption         =   "Frame Rate (FPS):"
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   180
      Width           =   1350
   End
End
Attribute VB_Name = "frmCapVid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CapParams As CAPTUREPARMS

Private Sub Command1_Click(index As Integer)
Select Case index
    Case 0 'OK
        Call ProcessCapInfo
        Me.Hide
    Case 1 'Cancel
        Unload Me
    Case 2 'Audio
        Call SetAudioFormatDlg(Me.hWnd)
    Case 3 'Video
        Call capDlgVideoFormat(frmMain.capwnd)
        Call ResizeCaptureWindow(frmMain.capwnd)
    Case 4 'Compress
        Call capDlgVideoCompression(frmMain.capwnd)
End Select
End Sub


Private Sub ProcessCapInfo()


With CapParams
'   // set the defaults we won't bother the user with
'   show message after pre-roll
    .fMakeUserHitOKToCapture = -(True) ' - converts VB Boolean to C BOOL
'   in case we use error callbacks later
    .wPercentDropForError = 10
'   fUsingDOSMemory is obsolete
    .fUsingDOSMemory = False
'   The number of video buffers should be enough to get through
'   disk seeks and thermal recalibrations
    .wNumVideoRequested = 32
'   Do abort on the left mouse
    .fAbortLeftMouse = -(True)
'   Do abort on the right mouse
    .fAbortRightMouse = -(True) '- converts VB boolean to C BOOL
'   If wChunkGranularity is zero, the granularity will be set to the
'   disk sector size.
    .wChunkGranularity = 0
'   use default
    .dwAudioBufferSize = 0
'   attempt to disable caching
    .fDisableWriteCache = -(True)
'   not using MCI
    .fMCIControl = False
    .fStepCaptureAt2x = False
'   not multi-threading
    .fYield = False
'   request audio buffers
    .wNumAudioRequested = 4 '10 is max limit

'   //these parameters are loaded from registry
    If "AUDIO" = Trim$(UCase(GetSetting(App.Title, "preferences", "streammaster", "AUDIO"))) Then
        .AVStreamMaster = AVSTREAMMASTER_AUDIO 'use audio clock to synchronize AVI
    Else
        .AVStreamMaster = AVSTREAMMASTER_NONE
    End If
    'set index size of AVI file (max frames)
    .dwIndexSize = Val(GetSetting(App.Title, "preferences", "maxframes", INDEX_15_MINUTES))

'   //Now set the parameters from the UserInput
    .dwRequestMicroSecPerFrame = microsSecFromFPS(Val(txtFps.Text))
    .fCaptureAudio = -(CBool(chkAudio.Value))
    .fLimitEnabled = -(CBool(chkLimit.Value))
    .wTimeLimit = Val(txtSec.Text)

End With
'set the new setup info
Call capCaptureSetSetup(frmMain.capwnd, CapParams)
'Kludgy - but...
Me.Tag = True 'this tells main form that OK button was pushed
End Sub

Private Function microsSecFromFPS(ByVal fps As Long) As Long
'note I am not using floating point here so these are not too exact
If fps = 0 Then Exit Function 'avoid divide by 0 errors
microsSecFromFPS = 1000000 / fps
End Function

Private Sub txtFps_LostFocus()
If Val(txtFps.Text) < 1 Then txtFps.Text = "1"
If Val(txtFps.Text) > 100 Then txtFps.Text = "100"
End Sub
Private Sub txtFPS_KeyPress(KeyAscii As Integer)
'allow backspace key
If KeyAscii = 8 Then Exit Sub
'logic to keep the user input valid
If KeyAscii < 48 Then KeyAscii = 0
If KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtSec_KeyPress(KeyAscii As Integer)
'allow backspace key
If KeyAscii = 8 Then Exit Sub
'logic to keep the user input valid
If KeyAscii < 48 Then KeyAscii = 0
If KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub Form_Load()
'this form loads settings automatically each time it is loaded
Call LoadMe
End Sub

Private Sub Form_Unload(Cancel As Integer)
'this form saves settings automatically each time it is unloaded
Call SaveMe
End Sub

Private Sub LoadMe()
    txtFps.Text = GetSetting(App.Title, "vidcap settings", "fps", "15")
    chkLimit.Value = Val(GetSetting(App.Title, "vidcap settings", "time limit", "0"))
    txtSec.Text = GetSetting(App.Title, "vidcap settings", "seconds", "30")
    chkAudio.Value = Val(GetSetting(App.Title, "vidcap settings", "cap audio", "0"))
    
End Sub

Private Sub SaveMe()
    Call SaveSetting(App.Title, "vidcap settings", "fps", txtFps.Text)
    Call SaveSetting(App.Title, "vidcap settings", "time limit", chkLimit.Value)
    Call SaveSetting(App.Title, "vidcap settings", "seconds", txtSec.Text)
    Call SaveSetting(App.Title, "vidcap settings", "cap audio", chkAudio.Value)
End Sub


