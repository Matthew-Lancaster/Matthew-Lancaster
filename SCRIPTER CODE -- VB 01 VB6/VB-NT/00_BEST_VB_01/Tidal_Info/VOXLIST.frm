VERSION 5.00
Begin VB.Form VOXLIST 
   BackColor       =   &H00FFFFFF&
   Caption         =   "VOX-SCRIPT"
   ClientHeight    =   5184
   ClientLeft      =   192
   ClientTop       =   840
   ClientWidth     =   9552
   LinkTopic       =   "Form1"
   ScaleHeight     =   5184
   ScaleWidth      =   9552
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Scroll to Bottom"
      Height          =   264
      Left            =   6780
      TabIndex        =   2
      Top             =   24
      Value           =   1  'Checked
      Width           =   1860
   End
   Begin VB.Timer Timer2 
      Interval        =   40000
      Left            =   1785
      Top             =   3900
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1785
      Top             =   3360
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3936
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   420
      Width           =   8724
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Width           =   6690
   End
   Begin VB.Menu MNU_EXPLORER_LOGG 
      Caption         =   "NOTEPAD LOGG"
   End
   Begin VB.Menu MNU_EXPLORER_DOOR 
      Caption         =   "EXPLORER DOOR FOLDER"
   End
   Begin VB.Menu MNU_SELECTION 
      Caption         =   "SELECT GO CLIPBOARD RESET"
   End
End
Attribute VB_Name = "VOXLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'03:50:00 - TIME #03:50

' voxlist.list1

Dim SELECT_AND_GO


Private Sub Form_Load()
'Me.Show
Call AUTORESIZE

If Label1 = "Label1" Then
    Label1 = "Wait to Show What 1st Entry"
End If

End Sub



Sub AUTORESIZE()


If VOXLIST_NOT_ACTIVE_1ST_RUN = True Then
    VOXLIST_NOT_ACTIVE = 0
Else
    VOXLIST_NOT_ACTIVE = 1000
End If

VOXLIST_NOT_ACTIVE_1ST_RUN = True



On Error Resume Next

If Me.WindowState = vbMinimized Then Me.Hide

'AUTO RESIZE
'Code to Auto Size Form based on controls used Including Menu Bar Sketchy Style
'Make sure form set to scale Twips
'put in your form load

x = 1
y = 1
On Error Resume Next
For Each Control In Controls
    If Control.Enabled = True And Control.Visible = True Then
        If Control.Width + Control.Left > x Then x = Control.Width + Control.Left
        If Control.Height + Control.Top > y Then y = Control.Height + Control.Top
        If InStr(UCase(Control.Name), "MNU_") > 0 Then mnu = 1
    End If
Next
'On Error GoTo 0

Me.Width = (x + 220)
If mnu = 1 Then
    pluso = 720: pluso = 1000 'Sometimes Different
Else
    pluso = 450
End If

Me.Height = (y + pluso)
Me.Refresh
DoEvents

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
List1.Left = 0
Label1.Top = 0
Label1.Left = 0
List1.Top = Label1.Top + Label1.Height + 20


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

VOXLIST_NOT_ACTIVE = 0

End Sub

Private Sub Form_Resize()
Call AUTORESIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)

If aa_MainCode.ALL_FORM_REQUEST_TO_END = False Then
    Cancel = True: Me.Hide: Exit Sub
End If

Me.Hide
If aa_MainCode.TrueTerminate = True Or QuitForms = True Then
    Exit Sub
Else
    Cancel = True
End If
End Sub

Private Sub List1_Click()
'Clipboard.Clear
'Clipboard.SetText List1.List(List1.ListIndex)
If SELECT_AND_GO = "" Then
    Label1 = List1.List(List1.ListIndex)
    Count_To_Min_Vox_Window = Now
End If

End Sub

Private Sub List1_DblClick()

SELECT_AND_GO = SELECT_AND_GO + List1.List(List1.ListIndex) + vbCrLf
Beep
Label1 = List1.List(List1.ListIndex)


End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Clipboard.Clear
'    Clipboard.SetText List1.List(List1.ListIndex)
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

VOXLIST_NOT_ACTIVE = 0

End Sub

Private Sub MNU_EXPLORER_DOOR_Click()
PATH_FILE_NAME1 = "C:\SCRIPTOR DATA\VB6\VB-NT\00_Best_VB_01\Tidal_Info\RS232 FRONT DOOR OPEN.txt"
'PATH_FILE_NAME1 = "C:\SCRIPTOR DATA\VB6\VB-NT\00_Best_VB_01\Tidal_Info\"
Shell "Explorer.exe /SELECT, """ + PATH_FILE_NAME1 + """", vbMaximizedFocus
'Shell "Explorer.exe """ + PATH_FILE_NAME1 + """", vbMaximizedFocus

End Sub

Private Sub MNU_EXPLORER_LOGG_Click()

'MsgBox "COMMAND NOT READY YET -- LOG FILE NOT BEING WRITEN DEACTIVATED", vbMsgBoxSetForeground
Dim PATH_SET_VAR As String
PATH_SET_VAR = App.Path + "\00_Text_Data\Voice_logg_Text\" + GetComputerName + "-" + GetUserName + ""
If Not FolderExists(PATH_SET_VAR) Then CreateFolderTree (PATH_SET_VAR)

FEEFILE_XP_LOGG = FreeFile
PATH_SET_VAR_2 = PATH_SET_VAR + "\Voice_logg_Text--" + Format(Now, "YYYY-MM-DD -- HH-MM-SS") + ".txt"
Open PATH_SET_VAR_2 For Output As #FEEFILE_XP_LOGG
For RZX = 0 To List1.ListCount - 1
    Print #FEEFILE_XP_LOGG, List1.List(RZX)
Next
Close #FEEFILE_XP_LOGG
'Shell "Explorer.exe /SELECT, """" + PATH_SET_VAR_2 + """, vbNormalFocus
Dim WSHShell
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RUN """" + PATH_SET_VAR_2 + """"
Set WSHShell = Nothing

End Sub

Public Sub MNU_SELECTION_Click()

For r = 1 To 100
    ro1 = Len(SELECT_AND_GO)
    SELECT_AND_GO = Replace(SELECT_AND_GO, "  ", " ")
    
    If ro1 = Len(SELECT_AND_GO) Then Exit For

Next

Clipboard.Clear
Clipboard.SetText SELECT_AND_GO
SELECT_AND_GO = ""
Me.WindowState = vbMinimized
Beep
Label1 = "Cleared"

End Sub

Public Sub Timer1_Timer()

VOXLIST_NOT_ACTIVE = VOXLIST_NOT_ACTIVE + 1
If VOXLIST_NOT_ACTIVE > 1000 Then VOXLIST_NOT_ACTIVE = 1000




End Sub

Private Sub Timer2_Timer()
VOXLIST.Timer2.Enabled = False
If Count_To_Min_Vox_Window + TimeSerial(0, 1, 0) < Now And VOXLIST_NOT_ACTIVE = 1000 Or VOXLIST_NOT_ACTIVE = 0 Then
    Me.WindowState = vbMinimized
    Beep
End If

End Sub
