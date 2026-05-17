VERSION 5.00
Object = "{019FD1D9-2A35-4D5E-AB4E-BA793946941B}#1.0#0"; "ClipboardViewer.ocx"
Begin VB.Form FRM_ClipTest_02 
   Caption         =   "Form1"
   ClientHeight    =   2436
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2436
   ScaleWidth      =   3744
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   912
   End
   Begin VB.Timer TIMER_CLIPBOARD_CHANGED 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1404
      Top             =   696
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   588
      Top             =   960
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   432
      Index           =   0
      Left            =   0
      Picture         =   "FRM_ClipTest_02.frx":0000
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   432
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   432
      Index           =   1
      Left            =   1032
      Picture         =   "FRM_ClipTest_02.frx":0252
      ScaleHeight     =   384
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   12
      Visible         =   0   'False
      Width           =   408
   End
   Begin ClipboardViewer.ctlClipboard ctlClipboard1 
      Left            =   492
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "FRM_ClipTest_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim O_TEXT_BOARD_CHK
' Public DONT_UPDATE_CLIPBOARD_THIS_ONE_LUKER
Public EXIT_TRUE

'Public HOOK_CLIPBOARD_API_lOADED As Boolean

'Private Sub cmdCopy_Click(Index As Integer)
'
'Select Case Index
'    Case 0
'        Clipboard.SetData Picture1(0).Picture, vbCFDIB
'    Case 1
'        Clipboard.SetData Picture1(1).Picture, vbCFDIB
'End Select
'
'End Sub


Private Sub ctlClipboard1_ClipboardChanged()

' ---------------------------------------------
' THIS IS A BIG BOOST TO CODE WORK GOOD
' WHEN MORE THAN ONE PROJECT USER API CLIPBOARD
' ---------------------------------------------
' [ Friday 04:27:20 Pm_30 August 2019 ]
' ---------------------------------------------

' If DONT_UPDATE_CLIPBOARD_THIS_ONE = True Then Stop

If DONT_UPDATE_CLIPBOARD_THIS_ONE = False Then
    TIMER_CLIPBOARD_CHANGED.Enabled = True
End If


End Sub


Sub REPLACE_HARDCODER()

    If IsIDE = True Then
        If Clipboard.GetFormat(vbCFText) = True Then
            Do
                Err.Clear
                TEXT_BOARD = Clipboard.GetText
                O_TEXT_BOARD_CHK = TEXT_BOARD
                If Err.Number > 0 Then
                    NOW_10 = Now + TimeSerial(0, 0, 20)
                    Do
                        DoEvents
                    Loop Until NOW_10 < Now
                End If
            Loop Until Err.Number = 0
        End If

        II = TEXT_BOARD
        If InStr(II, "D:\0 00 MUSIC ---\") > 0 Then
            II = " _ " + Replace(II, "D:\0 00 MUSIC ---\", "")
            STRING_OUT = II
            
            On Error Resume Next
            Do
                Err.Clear
                DONT_UPDATE_CLIPBOARD_THIS_ONE = True
                Clipboard.Clear
                If Err.Number > 0 Then
                    NOW_10 = Now + TimeSerial(0, 0, 20)
                    Do
                        DoEvents
                    Loop Until NOW_10 < Now
                End If
            Loop Until Err.Number = 0
            On Error GoTo 0
            
            On Error Resume Next
            Do
                Err.Clear
                DONT_UPDATE_CLIPBOARD_THIS_ONE = True
                Clipboard.SetText STRING_OUT
                If Err.Number > 0 Then
                    NOW_10 = Now + TimeSerial(0, 0, 20)
                    Do
                        DoEvents
                    Loop Until NOW_10 < Now
                End If
            Loop Until Err.Number = 0
            On Error GoTo 0
        End If
    End If


EXIT_SUB:

End Sub


Sub TIMER_CLIPBOARD_CHANGED_TIMER()

TIMER_CLIPBOARD_CHANGED.Enabled = False

Dim TEXT_BOARD

Timer_API_Test_Logick_Var2 = Now
TIME_LAST_CLIPBOARD = Now

'-------------------------------TEST2017
'Call Form1.COMPARE_FOR_EXE

'If EXECUTE_TIMER_ENABLED = False Then Exit Sub
'---------------------------------------------------------
'DELAY RELOAD API CLIPPER AS FREQENT AFTER IDLE FEW SECOND
'TO MUCH LONGER 8 MINUTE WHEN CLIPBOARD DONE
'---------------------------------------------------------

CLIPBOARD_ACTIVITY_MOMENT = Now + TimeSerial(0, 8, 0)

On Error GoTo EXIT_SUB

If Clipboard.GetFormat(vbCFText) = FLASE Then GoTo ENDER_SUB

Call REPLACE_HARDCODER

If Clipboard.GetFormat(vbCFText) = True Then
    VAR_DELAY = Now + TimeSerial(0, 0, 10)
    Do
        Err.Clear
        TEXT_BOARD = Clipboard.GetText
        If Err.Number > 0 Then
            Sleep 100
        End If
        If VAR_DELAY < Now Or Err.Number > 0 Then GoTo ENDER_SUB
    Loop Until Err.Number = 0
End If

O_TEXT_BOARD_CHK = TEXT_BOARD

FRM_ClipTest.MNU_GOODSYNC_VOLUME_LABLE_DRIVE.Enabled = Mid(TEXT_BOARD, 2, 1) = ":"
FRM_ClipTest.MNU_CONVERT_GSD_PATH_TO_NETWORK_FRIENDLY_ONE.Enabled = Mid(TEXT_BOARD, 1, 4) = "gstp"
'gstps://8-msi-gp62m-7rd.matt-lan-btinternet-com.goodsync/C:/0 00 LINK SET QUICKER MOVER
If Mid(TEXT_BOARD, 1, 4) = "gstp" Then
    If InStr(FRM_ClipTest.MNU_AUTO_CONVERT_GOODSYNC_DIRECT_BACK_TO_NETOWRK_SHARE.Caption, "AUTO CONVERT -- TRUE") > 0 Then
    '    Me.WindowState = vbMinimized
        Call FRM_ClipTest.MNU_CONVERT_GSD_PATH_TO_NETWORK_FRIENDLY_ONE_Click
    End If
End If



SEARCH_CRC_1 = TEXT_BOARD 'Mid(TEXT_BOARD, InStrRev(TEXT_BOARD, "~"))
If InStrRev(TEXT_BOARD, "~") = 0 Then
If InStr(SEARCH_CRC_1, "Text Size ") >= 27 Then
If InStr(SEARCH_CRC_1, "CRC32") >= 43 Then
If Len(SEARCH_CRC_1) <= 77 Then ' ---- IF SATURDAY AND FEBRUARY -- MAKE LEN UPTO 77
    ' CHECK IF LIKE HERE
    ' TEXT_CLIPPER = "Text Size " + Str(Len(Text2.Text)) + " -- " + Replace(LabelCRC1.Caption, "_1 =", "")
    GoTo ENDER_SUB
    ' WHY -- AND THEN OKAY IF LEN()<
End If
End If
End If
End If

If FRM_ClipTest.Text3.Text <> FRM_ClipTest.Text2.Text Then
'    Debug.Print Str(Len(FRM_ClipTest.Text2.Text))
'    Debug.Print Str(Len(FRM_ClipTest.Text3.Text))
    FRM_ClipTest.Text3.Text = FRM_ClipTest.Text2.Text
End If
'If FRM_ClipTest.Text2.Text <> "" Then
'    FRM_ClipTest.Text3.Text = FRM_ClipTest.Text2.Text
'End If

' HAVE MOVE CONTENT 02 TO 03
' AND NOT GET CONTECT 02
If Clipboard.GetFormat(vbCFText) = True Then
    Do
        Err.Clear
        FRM_ClipTest.Text2.Text = TEXT_BOARD  ' CLIPBOARD.GETTEXT
        '                                       PROBLEM HERE IF GET CLIPBOARD AND TWICE ROUTINE --
        '                                       SEEM TO STRIP A SPACE CHAR AT 3RD IN
        ' Wed 22-Jan-2020 01:24:48
        ' ----------------------------------------------------------------------------------------
        ' If Err.Number > 0 Then GoTo EXIT_SUB
        If Err.Number > 0 Then Sleep 100
        TIME_CLIPBOARD_CHANGE = Now
    Loop Until Err.Number = 0
End If
On Error GoTo 0

' HAVE MOVE CONTENT 02 TO 03
' AND NOT GET CONTECT 02
' FRM_ClipTest.Text2.Text = TEXT_BOARD

'If Clipboard.GetFormat(vbCFBitmap) = True Then
'    On Error Resume Next
'    Do
'        Err.Clear
'        Image1.Picture = Clipboard.GetData(vbCFBitmap)
'        If Err.Number > 0 Then Sleep 200
'    Loop Until Err.Number = 0
'    On Error GoTo 0
'End If

' FOR
' FRM_ClipTest.Timer_RESET_API_CLIPPER.Enabled = True
Timer2.Enabled = True

GoTo ENDER_SUB

EXIT_SUB:
TIMER_CLIPBOARD_CHANGED.Enabled = True

'EXIT SUB IF ERROR CALL ROUTINE AGAIN
' FRM_ClipTest.DONT_UPDATE_CLIPBOARD_THIS_ONE = DONT_UPDATE_CLIPBOARD_THIS_ONE_LUKER

ENDER_SUB:
Timer1.Enabled = False
' TIMER_CLIPBOARD_CHANGED.Enabled = True

End Sub


Private Sub Form_Load()

'    FRM_ClipTest_02.ctlClipboard1.EndClipboardViewer
'    Unload FRM_ClipTest_02
'    DoEvents
'    Load FRM_ClipTest_02
    FRM_ClipTest_02.ctlClipboard1.StartClipboardViewer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FRM_ClipTest_02.ctlClipboard1.EndClipboardViewer
'    Unload FRM_ClipTest_02
'    DoEvents
'    Load FRM_ClipTest_02
''    FRM_ClipTest_02.ctlClipboard1.StartClipboardViewer
End Sub

Private Sub Timer1_Timer()
    On Error GoTo EXIT_SUB_ENDER
    
    If Clipboard.GetFormat(vbCFText) = False Then Exit Sub
    
    TEXT_BOARD_CHK = Clipboard.GetText
    If TEXT_BOARD_CHK <> O_TEXT_BOARD_CHK Then
        O_TEXT_BOARD_CHK = TEXT_BOARD_CHK
        TIMER_CLIPBOARD_CHANGED.Enabled = True
        Timer2.Enabled = True
    End If
    
    
    Exit Sub
EXIT_SUB_ENDER:

End Sub

Private Sub Timer2_Timer()
    FRM_ClipTest.Timer_RESET_API_CLIPPER.Enabled = True
    Timer2.Enabled = False
End Sub
