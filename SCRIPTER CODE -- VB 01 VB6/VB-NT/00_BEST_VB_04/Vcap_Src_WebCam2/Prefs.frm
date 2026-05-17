VERSION 5.00
Begin VB.Form frmPrefs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "vbVidCap Preferences"
   ClientHeight    =   3990
   ClientLeft      =   1770
   ClientTop       =   3075
   ClientWidth     =   4275
   Icon            =   "Prefs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   885
      TabIndex        =   13
      Top             =   3465
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   2355
      TabIndex        =   12
      Top             =   3465
      Width           =   960
   End
   Begin VB.Frame staticframe 
      Caption         =   "Video and audio synchronization"
      Height          =   1230
      Index           =   2
      Left            =   75
      TabIndex        =   2
      Top             =   2010
      Width           =   4110
      Begin VB.OptionButton optSynch 
         Caption         =   "&No master (streams may differ in length)"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   840
         Width           =   3840
      End
      Begin VB.OptionButton optSynch 
         Caption         =   "Synch &video to audio"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   9
         Top             =   285
         Value           =   -1  'True
         Width           =   3840
      End
      Begin VB.Label lblStatic 
         Caption         =   "(video frame rate may change, VFW 1.x default)"
         Height          =   210
         Left            =   390
         TabIndex        =   11
         Top             =   525
         Width           =   3435
      End
   End
   Begin VB.Frame staticframe 
      Caption         =   "Maximum number of frames"
      Height          =   915
      Index           =   1
      Left            =   75
      TabIndex        =   1
      Top             =   990
      Width           =   4110
      Begin VB.OptionButton optMaxFrames 
         Caption         =   "324,000    (3 hours @ 30fps)"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   8
         Top             =   555
         Width           =   2550
      End
      Begin VB.OptionButton optMaxFrames 
         Caption         =   "27,000    (15 minutes @ 30fps)"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   255
         Value           =   -1  'True
         Width           =   2835
      End
   End
   Begin VB.Frame staticframe 
      Caption         =   "Background color"
      Height          =   660
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   210
      Width           =   4125
      Begin VB.OptionButton optColor 
         Caption         =   "Black"
         Height          =   285
         Index           =   3
         Left            =   3285
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Dk Gray"
         Height          =   285
         Index           =   2
         Left            =   2190
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Lt Gray"
         Height          =   285
         Index           =   1
         Left            =   1140
         TabIndex        =   4
         Top             =   240
         Width           =   960
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Default"
         Height          =   285
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   240
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private color As Long
Private maxframes As Long
Private streammaster As String

Private Sub cmdCancel_Click()
    'lose all changes
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'commit back color changes
    frmMain.BackColor = color
    Call SaveSetting(App.Title, "preferences", "backcolor", color)
    'commit maxframes
    Call SaveSetting(App.Title, "preferences", "maxframes", maxframes)
    'commit streammaster
    Call SaveSetting(App.Title, "preferences", "streammaster", streammaster)
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim index As Long
    
    'load current settings
    color = Val(GetSetting(App.Title, "preferences", "backcolor", "&H404040"))
    maxframes = Val(GetSetting(App.Title, "preferences", "maxframes", INDEX_15_MINUTES))
    streammaster = GetSetting(App.Title, "preferences", "streammaster", "AUDIO")
    'and set up form
    Select Case color
        Case &H0& 'black
            index = 3
        Case &HC0C0C0 'lt gray
            index = 1
        Case &H404040 'dk gray
            index = 2
        Case &H80000005 'default window
            index = 0
        Case Else
            index = 0
    End Select
    optColor(index).Value = True 'set correct color option
    If maxframes <> INDEX_15_MINUTES Then
        index = 1
    Else
        index = 0
    End If
    optMaxFrames(index).Value = True 'set correct frames option
    If Trim$(UCase(streammaster)) = "AUDIO" Then
        index = 0
    Else
        index = 1
    End If
    optSynch(index).Value = True
    
End Sub

Private Sub optColor_Click(index As Integer)
    Select Case index
        Case 0 'default
            color = &H80000005  'use windows' system
        Case 1 'lt gray
            color = &HC0C0C0
        Case 2 'dk gray
            color = &H404040
        Case 3 'black
            color = &H0&
    End Select
End Sub

Private Sub optMaxFrames_Click(index As Integer)
    Select Case index
        Case 0 '27000 (15 minutes)
            maxframes = INDEX_15_MINUTES
        Case 1 '324000 (3 hours)
            maxframes = INDEX_3_HOURS
    End Select
End Sub

Private Sub optSynch_Click(index As Integer)
    Select Case index
        Case 0 'Audio
            streammaster = "AUDIO"
        Case 1 ' None
            streammaster = "NONE"
    End Select
End Sub
