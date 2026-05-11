VERSION 5.00
Begin VB.Form frm_MSGBOX 
   Caption         =   "Clipboard Logger Error and Info Messages Script List"
   ClientHeight    =   4932
   ClientLeft      =   3336
   ClientTop       =   3096
   ClientWidth     =   10740
   LinkTopic       =   "Form3"
   ScaleHeight     =   4932
   ScaleWidth      =   10740
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Height          =   4080
      Left            =   20
      TabIndex        =   0
      Top             =   -10
      Width           =   10600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   720
      Top             =   825
   End
   Begin VB.Menu MNU_EXIT 
      Caption         =   "EXIT_HIDE"
   End
   Begin VB.Menu MNU_NUMBER_OF_ERROR 
      Caption         =   "Number Of Errror Warning Item"
   End
End
Attribute VB_Name = "frm_MSGBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub List1_Click()

'frm_msgbox.List1

End Sub

Private Sub Mnu_Exit_Click()

Unload Me

End Sub

Private Sub Timer1_Timer()
    'FRM_MSGBOX

    'Timer1.ENABLED 'CHECKER

    'Me.Show
    'Me.WindowState = vbNormal
    'Me.Hide
    TIME_LAST_CLIPBOARD_ERROR_MSG = Now
    'DoEvents

    Beep
    
    'WHY REMOVE CRLF
    'MSGBOX2 = Replace(MSGBOX2, vbCrLf, " -- ")
    
    Do
        MSGBOX2 = Replace(MSGBOX2, "VBCRLF", " ")
        If O_MSGBOX2 = MSGBOX2 Then Exit Do
        O_MSGBOX2 = MSGBOX2
    Loop Until True = False
    
    
    MSGBOX2 = Format(Now, "DDD DD-MM-YYYY__HH:MM:SS ") + " -- " + vbCrLf + MSGBOX2

    MSGBOX3 = MSGBOX2
    MSGBOX3 = Replace(MSGBOX3, vbCrLf, " -- ")
    MSGBOX3 = Replace(MSGBOX3, " -- -- ", " -- ")
    MSGBOX3 = Replace(MSGBOX3, " --  -- ", " -- ")
    
    List1.AddItem MSGBOX3
    
    '---------------------------------------
    '** Show Error Script **
    '---------------------------------------
    ERROR_LOGG_READ_STATUS = "-- UN-READ ITEM WAIT TO VIEW--"
    ERROR_LOGG_UPDATE_TIME = Now
    Call UPDATE_MNU_ERROR
    '---------------------------------------
    
    
    On Error Resume Next
    
    If KEYBOARD_OR_MOUSE_ACTIVE_3_MIN < Now Then
        If List1.ListCount > 1 Then
            List1.ListIndex = List1.ListCount
        End If
        List1.ListIndex = List1.ListCount - 1
    End If
    
    'MSGBOX2 = Replace(MSGBOX2, vbCrLf, vbCrLf + "----" + vbCrLf)
    
    '-------------------------------------
    'MsgBox MSGBOX2, vbMsgBoxSetForeground
    
    '-------------------------------------
    'WHY ---------------------------------
    '-------------------------------------
'    FRM_MSGBOX2.Show
'    FRM_MSGBOX2.WindowState = vbNormal
'    FRM_MSGBOX2.Label1.Caption = MSGBOX2
'    FRM_MSGBOX2.Label1.Height = FRM_MSGBOX2.Height
'    FRM_MSGBOX2.Label1.Width = FRM_MSGBOX2.Width
'    'FRM_MSGBOX2.Label1.AutoSize = True
'    FRM_MSGBOX2.Label1.FontSize = 8
'    '-------------------------------------
    
    FRM_MSGBOX.Timer1.Enabled = False

    'Me.Hide

End Sub

Sub Form_Load()
    
    Me.Left = Form1.Left
    Me.Width = Form1.Width

    List1.Height = Me.Height
    List1.Width = Me.Width * 3
    
    xy = "Number Of Errror Warning Item =" + Str(FRM_MSGBOX.List1.ListCount)
    If MNU_NUMBER_OF_ERROR.Caption = xy Then Exit Sub
    MNU_NUMBER_OF_ERROR.Caption = xy
    
End Sub

Sub UPDATE_MNU_ERROR()

    If ERROR_LOGG_UPDATE_TIME = 0 Then
        ERROR_LOGG_UPDATE_TIME_2 = "0 Nothing Time"
    End If
    If ERROR_LOGG_UPDATE_TIME > 0 Then
        ERROR_LOGG_UPDATE_TIME_2 = Str(DateDiff("s", ERROR_LOGG_UPDATE_TIME, Now))
        ERROR_LOGG_UPDATE_TIME_2 = ERROR_LOGG_UPDATE_TIME_2 + " Second Away"
        If DateDiff("s", ERROR_LOGG_UPDATE_TIME, Now) > 60 Then
            ERROR_LOGG_UPDATE_TIME_2 = Str(DateDiff("n", ERROR_LOGG_UPDATE_TIME, Now))
            ERROR_LOGG_UPDATE_TIME_2 = ERROR_LOGG_UPDATE_TIME_2 + " Minute Away"
        End If
        ITXT = ITXT + TIME_SET
    End If

    On Error Resume Next
    If FRM_MSGBOX.List1.ListCount > 0 Then
        TEST_STRING_VAR = "** Show Error Script *" + Str(FRM_MSGBOX.List1.ListCount) + " ITEM COUNT " + ERROR_LOGG_READ_STATUS + " " + ERROR_LOGG_UPDATE_TIME_2 + " **"
    Else
        TEST_STRING_VAR = "** Show Error Script **"
    End If
    On Error GoTo 0
    
    If Form1.Mnu_Show_Error_Script_Frm_Msgbox.Caption <> TEST_STRING_VAR Then
        Form1.Mnu_Show_Error_Script_Frm_Msgbox.Caption = TEST_STRING_VAR
    End If
    On Error Resume Next
    xy = "Number Of Errror Warning Item =" + Str(FRM_MSGBOX.List1.ListCount)
    If MNU_NUMBER_OF_ERROR.Caption = xy Then Exit Sub
    MNU_NUMBER_OF_ERROR.Caption = xy

End Sub

Private Sub Form_Unload(Cancel As Integer)


    '---------------------------------------
    '** Show Error Script **
    '---------------------------------------
    ERROR_LOGG_READ_STATUS = "-- ALL ITEM READ DONE --"
    Call UPDATE_MNU_ERROR
    '---------------------------------------
    
    EXIT_TRUE = True
    On Error Resume Next
    For Each Form In Forms
        If Form1.EXIT_TRUE = True Then EXIT_TRUE = True
    Next Form
    On Error GoTo 0
    If EXIT_TRUE = True Then Cancel = False

    If EXIT_TRUE = False Then
        Me.Hide
        'test may have to put back form need reseting
        'Unload FrmJoy
        Cancel = True
        Exit Sub
    End If

End Sub
