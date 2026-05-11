VERSION 5.00
Object = "{019FD1D9-2A35-4D5E-AB4E-BA793946941B}#1.0#0"; "ClipboardViewer.ocx"
Begin VB.Form FRMCLIPTEST01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clipboard Viewer"
   ClientHeight    =   4392
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5388
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4392
   ScaleWidth      =   5388
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3984
      Top             =   576
   End
   Begin VB.Timer TIMER_CLIPBOARD_CHANGED 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   576
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3240
      Top             =   576
   End
   Begin VB.Timer Timer_Test_Logic 
      Interval        =   1000
      Left            =   2850
      Top             =   570
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Text"
      Height          =   495
      Left            =   150
      TabIndex        =   6
      Top             =   1410
      Width           =   1320
   End
   Begin VB.TextBox Text2 
      Height          =   1815
      Left            =   2715
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmClipTest.frx":0000
      Top             =   2370
      Width           =   2460
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1530
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1470
      Width           =   2040
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy Picture 2"
      Height          =   495
      Index           =   1
      Left            =   150
      TabIndex        =   3
      Top             =   780
      Width           =   1320
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   432
      Index           =   1
      Left            =   1530
      Picture         =   "frmClipTest.frx":0006
      ScaleHeight     =   384
      ScaleWidth      =   360
      TabIndex        =   2
      Top             =   750
      Width           =   408
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy Picture 1"
      Height          =   495
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   165
      Width           =   1320
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   432
      Index           =   0
      Left            =   1530
      Picture         =   "frmClipTest.frx":0254
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   0
      Top             =   135
      Width           =   432
   End
   Begin ClipboardViewer.ctlClipboard ctlClipboard1 
      Left            =   3660
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Text in Clipboard"
      Height          =   270
      Index           =   1
      Left            =   2715
      TabIndex        =   8
      Top             =   2100
      Width           =   1605
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture in Clipboard"
      Height          =   270
      Index           =   0
      Left            =   180
      TabIndex        =   7
      Top             =   2115
      Width           =   1605
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   180
      Stretch         =   -1  'True
      Top             =   2370
      Width           =   2460
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   5265
      Y1              =   1995
      Y2              =   1995
   End
End
Attribute VB_Name = "FRMCLIPTEST01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I = MsgBox("ClipBoard API Has Stopped and Gone Missing" + vbCrLf + "Use the Menu Option *INFO* to Diagnostic and Reload It" + vbCrLf + "This Can Happen If ChkDsk Unlocked All Handles to The Hard Drive and the ClipboardViewer.ocx Driver Couldn't Get Access", vbMsgBoxSetForeground)
'ClipboardViewer.ocx Driver
'ClipboardViewer.oca Driver -- TRUE NAME EXT
'THE VERSION USE IS
'ClipboardViewer.oca
'C:\Program Files\Microsoft Visual Studio\VB98\ClipboardViewer.oca   0x488   29/08/2015 22:56:01 03/04/2016 21:48:47 A   10,240  *           *           0x00120089  1,696   7244    VB6.EXE C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE   oca 16.56%
'MUST BE THE REGISTER VERSION
'AS ONE IN ROOT OF APP EXE DON'T LOCATE TO USE

Dim TEXT_BOARD
Dim O_TEXT_BOARD_CHK

Public CALC_ADDER_ENTRY

Option Explicit

'Public HOOK_CLIPBOARD_API_lOADED As Boolean

Private Sub ctlClipboard1_ClipboardChanged()

    ' FINDER AT -- Load FRMCLIPTEST01
    TIMER_MISSING_PULSE_CLIPBOARD_01 = Now
    
    ' ---------------------------------------------
    ' THIS IS A BIG BOOST TO CODE WORK GOOD
    ' WHEN MORE THAN ONE PROJECT USER API CLIPBOARD
    ' ---------------------------------------------
    ' [ Friday 04:27:20 Pm_30 August 2019 ]
    ' ---------------------------------------------
    If EXECUTE_TIMER_ENABLED = False Then
            O_TEXT_BOARD_CHK = "-- BLOCK CLIPPER"
            EXECUTE_TIMER_ENABLED = True
        Exit Sub
    End If
    TIMER_CLIPBOARD_CHANGED.Enabled = True
End Sub

Sub TIMER_CLIPBOARD_CHANGED_TIMER()

    ' FINDER AT -- Load FRMCLIPTEST01
    
    Dim DONT_DO_REMAINDER_OF_GO
    
    TIMER_CLIPBOARD_CHANGED.Enabled = False
    
    ' API_CLIPBOARD_TEXT_CHANGE_FOR_FORMAT_ALL_CAPITAL_MODE = True
    
    TEXT_BOARD = ""
    
    Timer_API_Test_Logick_Var2 = Now
    TIME_LAST_CLIPBOARD = Now
    
    ' COMPARE LAST TWO CLIPBOARD FOR EXE PROGRAM TO RUN WINCOMPARE
    ' -------------------------------------------------------------
    ' NOT RUN HERE ANYMORE -- CODE TO LOOK AT
    ' ------------------------------------------------------------
    If 1 = 2 Then
        Call Form1.COMPARE_FOR_EXE
    End If
    
    ' THIS STOP UPDATE OF CLIPBOARD WHEN CLIPBOARD SET SOME
    ' -----------------------------------------------------
    If EXECUTE_TIMER_ENABLED = False Then Exit Sub
    
    '---------------------------------------------------------
    'DELAY RELOAD API CLIPPER AS FREQENT AFTER IDLE FEW SECOND
    'TO MUCH LONGER 8 MINUTE WHEN CLIPBOARD DONE
    '---------------------------------------------------------
    
    CLIPBOARD_ACTIVITY_MOMENT = Now + TimeSerial(0, 8, 0)
    
    Form1.Timer1.Enabled = True
    Form1.Timer1.Interval = 1
    
    CALC_ADDER_ENTRY = True
    
    On Error GoTo EXIT_SUB
    DONT_DO_REMAINDER_OF_GO = False
    If Clipboard.GetFormat(vbCFText) = True Then
        Do
            Err.Clear
            
            ' HARD WORK IF HIGH SPEED CLIPBOARD GO
            ' ------------------------------------
            If O_TEXT_BOARD_CHK = "-- BLOCK CLIPPER" Then
                DONT_DO_REMAINDER_OF_GO = True
            End If
            TEXT_BOARD = Clipboard.GetText
            O_TEXT_BOARD_CHK = TEXT_BOARD
            If Err.Number > 0 Then
                GoTo EXIT_SUB
            End If
        Loop Until Err.Number = 0
    End If
    
    If DONT_DO_REMAINDER_OF_GO = True Then
        Exit Sub
    End If
    
    Timer2.Enabled = True
    
    Exit Sub
    
EXIT_SUB:
    ' ------------------------------------
    ' EXIT SUB IF ERROR CALL ROUTINE AGAIN
    ' ------------------------------------
    
    Timer1.Enabled = False

End Sub


Private Sub Form_Load()

    'HOOK_CLIPBOARD_API_lOADED = True
    API_LOAD = True
    Form1.Mnu_API_Reload_Status.Caption = "The API Form Clipper Logger Sub Call Loaded @ " + Format(Now, "DDD DD-MM-YYYY HH:MM:SS")
    
    If Form1.Mnu_API_UNLoad_Status.Caption = "The API Form Clipper Logger Sub Call UN-Loaded @ " Then
        Form1.Mnu_API_UNLoad_Status.Caption = "The API Form Clipper Logger Sub Call UN-Loaded @ NOT HAPPEN YET NOT A TIME SET AT MOMENT"
    End If
    
    Me.Hide
    
    ChDrive "D:\"
    ChDir "D:\VB6\VB-NT\00_BEST_VB_01\Clipboard Logger\CLIPBOARDVIEWER_01"
    
    'Start to monitor the clipboard
    ctlClipboard1.StartClipboardViewer
    
    '-----------------------------
    O_VAL_MINUTE_API = Now
    'Call Form1.MNU_API_RESET_SUB
    '-----------------------------
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'HOOK_CLIPBOARD_API_lOADED = False
    
    Form1.Mnu_API_UNLoad_Status.Caption = "The API Form Clipper Logger Sub Call UN-Loaded @ " + Format(Now, "DDD DD-MM-YYYY HH:MM:SS")
    
    'Stop the clipboard viewer
    ctlClipboard1.EndClipboardViewer
    
    'Call Form1.MNU_API_RESET_Click
    API_LOAD = False
End Sub

Public Sub Timer_Test_Logic_Timer()

    Timer_Test_Logic.Enabled = False
    Timer_Test_Logic.Interval = 10000
    'REMOVE THIS ONE AUG 08 2K SIXTEEN
    
    'HOOK_CLIPBOARD_API_lOADED = True
    
    If IsIDE = True Then
        Timer_Test_Logic.Interval = 4000
    Else
        Timer_Test_Logic.Interval = 1000
    End If
    
    Timer_API_Test_Logick_Var1 = Now
    
End Sub

Private Sub Timer2_Timer()
    
    Form1.Timer_RESET_API_CLIPPER.Enabled = True
    Timer2.Enabled = False

End Sub
