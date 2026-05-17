VERSION 5.00
Begin VB.Form frm_MSGBOX 
   Caption         =   "Clipboard Logger Error and Info Messages Script List"
   ClientHeight    =   4932
   ClientLeft      =   3336
   ClientTop       =   2796
   ClientWidth     =   10740
   LinkTopic       =   "Form3"
   ScaleHeight     =   4932
   ScaleWidth      =   10740
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4224
      Left            =   20
      TabIndex        =   0
      Top             =   60
      Width           =   10600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   720
      Top             =   825
   End
End
Attribute VB_Name = "frm_MSGBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DO_ONCE_ERROR_LOGG_SCREEN
Dim OL_DIMENSION

Private Sub Form_Resize()
    
    If Me.Width + Me.Height = OL_DIMENSION Then Exit Sub
    If Me.Width + Me.Height <> OL_DIMENSION Then
        DO_ONCE_ERROR_LOGG_SCREEN = False
    End If
    
    OL_DIMENSION = Me.Width + Me.Height
    
    If DO_ONCE_ERROR_LOGG_SCREEN = True Then Exit Sub
        
        DO_ONCE_ERROR_LOGG_SCREEN = True
        List1.Top = 10
        List1.Left = 10
        List1.Height = Me.Height - 800
        List1.Width = Me.Width - 800

End Sub

Private Sub List1_Click()

'frm_msgbox.List1

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
        List1.ListIndex = List1.ListCount
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
    
    frm_MSGBOX.Timer1.Enabled = False

    'Me.Hide

End Sub

Sub Form_Load()

End Sub

Sub UPDATE_MNU_ERROR()

    If ERROR_LOGG_UPDATE_TIME = 0 Then
        ERROR_LOGG_UPDATE_TIME_2 = "0 Nothing Time"
    Else
        ERROR_LOGG_UPDATE_TIME_2 = Trim(Str(DateDiff("S", ERROR_LOGG_UPDATE_TIME, Now))) + " Seconds Away"
    End If
    
    Form1.Mnu_Show_Error_Script_Frm_Msgbox.Caption = "** Show Error Script *" + Str(frm_MSGBOX.List1.ListCount) + " ITEM COUNT " + ERROR_LOGG_READ_STATUS + " " + ERROR_LOGG_UPDATE_TIME_2 + " **"

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
        If Form.EXIT_TRUE = True Then EXIT_TRUE = True
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
