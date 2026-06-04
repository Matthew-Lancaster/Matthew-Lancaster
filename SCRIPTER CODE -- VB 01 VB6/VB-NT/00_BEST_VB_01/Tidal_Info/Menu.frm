VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Menu 
   BackColor       =   &H80000007&
   Caption         =   "Tidal Menu"
   ClientHeight    =   4884
   ClientLeft      =   60
   ClientTop       =   312
   ClientWidth     =   5688
   LinkTopic       =   "Form1"
   ScaleHeight     =   4884
   ScaleWidth      =   5688
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CheckBox MuteAllMal 
      Caption         =   "Mute All  Mall"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   15
      Top             =   4200
      Width           =   1920
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   4500
      Left            =   1935
      TabIndex        =   14
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   7938
      _Version        =   393216
      Orientation     =   1
      Max             =   100
      SelStart        =   5
      TickFrequency   =   10
      Value           =   5
   End
   Begin VB.CheckBox SoundMall 
      Caption         =   "Mall Sound or Spit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   13
      Top             =   3900
      Width           =   1920
   End
   Begin VB.CheckBox CheckCopyPaste 
      Caption         =   "F12 Copy (Paste)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   12
      Top             =   3600
      Width           =   1920
   End
   Begin VB.CheckBox CheckIndentEmailOut 
      Caption         =   "Indent Out >> Email"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   11
      Top             =   2085
      Width           =   1920
   End
   Begin VB.CheckBox CheckPublisher 
      Caption         =   "Publisher Del Page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   3300
      Width           =   1920
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Text            =   "01"
      Top             =   2985
      Width           =   1920
   End
   Begin VB.CheckBox Countcheck 
      Caption         =   "Count Nero"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Top             =   2685
      Width           =   1920
   End
   Begin VB.CheckBox CheckLock 
      Caption         =   "Lock To Caller"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   2385
      Value           =   1  'Checked
      Width           =   1920
   End
   Begin VB.CheckBox CheckIndentOut 
      Caption         =   "Indent Out"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   1785
      Width           =   1920
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   3750
      Top             =   1890
   End
   Begin VB.CheckBox CheckIndent 
      Caption         =   "Indent In"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   1485
      Width           =   1920
   End
   Begin VB.CheckBox CheckMumBT 
      Caption         =   "Mum BT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   885
      Width           =   1920
   End
   Begin VB.CheckBox CheckMumAL 
      Caption         =   "Mum A an L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   1185
      Width           =   1920
   End
   Begin VB.CheckBox CheckEddieBT 
      Caption         =   "Eddie BT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   585
      Width           =   1920
   End
   Begin VB.CheckBox CheckTide 
      Caption         =   "Tide Talk Off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   285
      Value           =   1  'Checked
      Width           =   1920
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3750
      Top             =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Sound/Time Off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   1905
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MenuWorking
Public Quit As Date
Public Sword As Date
Dim Th(100)

Sub CoolCakes()

If MenuWorking = 1 Then Exit Sub
Timer1.Enabled = False
Timer1.Interval = 0
Timer1.Enabled = Enabled
Timer1.Interval = 4000
'Call Checks
Call zzSave_Checks

End Sub

Private Sub Check1_Click()

Soff = True
If Menu.Check1.Value = vbChecked Then
    Soff = True
Else
    Soff = False
End If
Call CoolCakes
End Sub


Private Sub CheckCopyPaste_Click()

Timer1.Enabled = Enabled
Call Checks

End Sub

Private Sub CheckIndentEmailOut_Click()

If CheckIndentEmailOut.Value = vbChecked Then
    Gooze6 = Now + TimeSerial(0, 1, 0)
End If
If CheckIndentEmailOut.Value = vbUnchecked Then
    Gooze6 = 0
End If

Timer1.Enabled = Enabled
Call Checks

End Sub

Private Sub CheckIndentOut_Click()

If CheckIndentOut.Value = vbChecked Then
    Gooze2 = Now + TimeSerial(0, 1, 0)
End If
If CheckIndentOut.Value = vbUnchecked Then
    Gooze2 = 0
End If

Timer1.Enabled = Enabled
Call Checks


End Sub

Private Sub CheckIndent_Click()
If CheckIndent.Value = vbChecked Then
    Gooze1 = Now + TimeSerial(0, 1, 0)
End If
If CheckIndent.Value = vbUnchecked Then
    Gooze1 = 0
End If

Timer1.Enabled = Enabled
Call Checks

End Sub

Private Sub CheckLock_Click()

Call CoolCakes

End Sub

Private Sub CheckPublisher_Click()

If CheckPublisher.Value = vbChecked Then
    Gooze5 = Now + TimeSerial(0, 2, 0)
End If
If CheckPublisher.Value = vbUnchecked Then
    Gooze5 = 0
End If

Timer1.Enabled = Enabled
Call Checks

End Sub

Private Sub CheckTide_Click()

Toff = True
If Menu.CheckTide.Value = vbChecked Then
    Toff = True
Else
    Toff = False
End If
Call CoolCakes
End Sub

Private Sub CheckMumAL_Click()
If CheckMumAL.Value = vbChecked Then
    Gooze3 = Now + TimeSerial(0, 15, 0)
End If
If CheckMumAL.Value = vbUnchecked Then
    Gooze3 = 0
End If
Timer1.Enabled = Enabled
Call Checks
End Sub
Private Sub CheckEddieBT_Click()
If Sword = 1 Then Exit Sub
Sword = 1
CheckMumBT.Value = vbUnchecked
Sword = 0
Call CoolCakes

End Sub

Private Sub CheckMumBT_Click()
If Sword = 1 Then Exit Sub
Sword = 1
CheckEddieBT.Value = vbUnchecked
Sword = 0
Call CoolCakes
End Sub


Private Sub Countcheck_Click()

'Menu.Countcheck.Value = vbUnchecked
If CheckMumAL.Value = vbChecked Then
    Gooze3 = Now + TimeSerial(0, 15, 0)
End If
If CheckMumAL.Value = vbUnchecked Then
    Gooze3 = 0
End If
Timer1.Enabled = Enabled
Call Checks

'Count2Check = Val(Text1.Text)

Call CoolCakes

End Sub

Private Sub Form_Activate()

Menu.Top = ATidalX.Top + 600
Menu.Left = ATidalX.Left

End Sub

Private Sub Form_Load()
Menu.Top = ATidalX.Top + 600
Menu.Left = ATidalX.Left
MenuWorking = 1

Call zzLoadChecks

If Dir$(App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Check.txt") <> "" And 1 = 2 Then
    freeval = FreeFile
    Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Check.txt" For Input As #freeval
    For Each Control In Menu.Controls
        If EOF(freeval) Then Exit For
        If InStr(Control.Name, "Slider") = 0 Then
            Line Input #freeval, lp1$
            Control.Value = Val(lp1$)
        End If
    Next
    Close #freeval
End If

Call Check1_Click
Call CheckTide_Click
MenuWorking = 0




'Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Slider Hiss.txt" For Input As #1
'Line Input #1, sv
'Close #1
'Slider1.Value = Val(sv)

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Exit Sub 'Temp

If aa_MainCode.ALL_FORM_REQUEST_TO_END = False Then Cancel = True: Me.Hide: Exit Sub

If QuitForms = False Then
Cancel = True
Menu.Hide

Call zzSave_Checks
End If
'cool = FreeFile
'Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Slider Hiss.txt" For Output As #cool
'Print #cool, Slider1.Value
'Close #cool

End Sub

Private Sub MuteAllMal_Click()
    Call CoolCakes

End Sub



Private Sub SoundMall_Click()
    'SoundMall
End Sub

Private Sub Timer1_Timer()

Menu.Hide
Timer1.Enabled = False
Call zzSave_Checks

End Sub
Private Sub Timer2_Timer()

If Gooze1 < Now And Gooze1 > 0 Then
    CheckIndent.Value = vbUnchecked
End If

If Gooze2 < Now And Gooze2 > 0 Then
    CheckIndentOut.Value = vbUnchecked
End If

If Gooze3 < Now And Gooze3 > 0 Then
    CheckMumAL.Value = vbUnchecked
End If

If Gooze4 < Now And Gooze4 > 0 Then
    Countcheck.Value = vbUnchecked
End If

If Gooze5 < Now And Gooze4 > 0 Then
    Countcheck.Value = vbUnchecked
End If

If Gooze6 < Now And Gooze2 > 0 Then
    CheckIndentEmailOut.Value = vbUnchecked
End If

'Timer2.Enabled = False

End Sub
Private Sub Checks()

If MenuWorking = 1 Then Exit Sub

Call zzLoadChecks

Exit Sub

freeval = FreeFile
Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Check.txt" For Output As #freeval
For Each Control In Menu.Controls
'    If InStr(Control.name, "Check") Then
    If InStr(Control.Name, "Slider") = 0 Then
        On Local Error Resume Next
            easycake$ = str$(Control.Value)
            ii = Err.Number
        On Local Error GoTo 0
        If ii = 0 Then
        
        If InStr(Control.Name, "EddieBT") Then easycake$ = str$(vbUnchecked)
        If InStr(Control.Name, "MumBT") Then easycake$ = str$(vbUnchecked)
        If InStr(Control.Name, "MumAL") Then easycake$ = str$(vbUnchecked)
        If InStr(Control.Name, "Indent") Then easycake$ = str$(vbUnchecked)
        If InStr(Control.Name, "Publisher") Then easycake$ = str$(vbUnchecked)
        If InStr(Control.Name, "CopyPaste") Then easycake$ = str$(vbUnchecked)
        Print #freeval, easycake$
        End If
    End If
Next
Close #freeval

End Sub

Sub zzLoadChecks()

'---------------------
If Dir(App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\", vbDirectory) = "" Then
    On Error Resume Next
        MkDir App.Path + "\00_Text_Data\"
        MkDir App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\"
    On Error GoTo 0
End If

If Dir(App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Check.txt") <> "" Then
    dxx = 0
    i = 0
    On Error Resume Next
    Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Check.txt" For Input As #1
    If Err.Number > 0 Then Exit Sub
    
    Do
        Line Input #1, vv$
        Th(i) = vv$
        i = i + 1
        Line Input #1, vv$
        Th(i) = Val(vv$)
        i = i + 1
    Loop Until EOF(1)
    Close #1
        
    tit = i
        
    For Each Control In Me.Controls
        XX = 0
        If InStr(Control.Name, "Combo") Then XX = 1
    '    If InStr(Control.name, "Check") Then xx = 1
        
        xxd = 0: gogo = 0
        For r = 0 To tit
            
            If Control.Name = Th(r) Then
    '        If Control.name = "MuteAllMal" Then Stop
                
                xxd = r: gogo = 1: Exit For
            End If
        Next
        
        
        If gogo = 1 And XX = 0 Then
            'If Control.name = "MuteAllMal" Then Stop
            'If Th(xxd + 1) = 0 Then yy = vbUnchecked Else yy = vbChecked
    '        T = Control.name
            On Local Error Resume Next
            Control.Value = Th(xxd + 1)
            On Local Error GoTo 0
        
        End If
    Next
    
End If

On Local Error GoTo 0

End Sub

Sub zzSave_Checks()

'Shell "notepad " + App.Path + "\00_Text_Data\ChkSettings.txt"
'YinVectKeli$ = Combo1.List(Combo1.ListIndex)
'MsgBox "Save " + YinVectKeli$
'Me.Visible = False
'Me.Visible = True

On Local Error Resume Next
Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Check.txt" For Output As #1
'Open App.Path + "\00_Text_Data\ChkSettings-" + YinVectKeli$ + ".txt" For Output As #1
For Each Control In Me.Controls
    Err.Clear
    'If Control.name = "MuteAllMal" Then Stop
    A = Control.Value
    'If Control.name <> "Slider1" Then
    If Err.Number = 0 Then
        'If Control.name = "MuteAllMal" Then Stop
        Print #1, Control.Name
        Print #1, Control.Value
    'End If
    End If
Next
Close #1


Exit Sub

'--------------Sweet New Code

End Sub



