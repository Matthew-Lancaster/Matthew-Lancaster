VERSION 5.00
Begin VB.Form OptionsMenu 
   Caption         =   "CID Runner Options Menu"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OptionsMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      Caption         =   "Day Counter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   4
      Top             =   750
      Width           =   4110
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Last Reboot Counter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   510
      Width           =   4110
   End
   Begin VB.CheckBox Check1 
      Caption         =   "From This Program Start Counter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   270
      Width           =   4110
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   3030
      Left            =   0
      TabIndex        =   1
      Top             =   990
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Set The Mouse And KeyBoard Counter Method"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "OptionsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EggF

Private Sub Check1_Click()

If EggF = 1 Then Exit Sub
CheckT1$ = "1"
Call Ghost

End Sub

Private Sub Check2_Click()

If EggF = 1 Then Exit Sub
CheckT1$ = "2"
Call Ghost

End Sub

Private Sub Check3_Click()

If EggF = 1 Then Exit Sub
CheckT1$ = "3"
Call Ghost

End Sub

Private Sub Form_Load()


List1.AddItem "Mouse / Keyboard Counts From Previous End of Day Totals"

If Dir$("C:\Callerid\ci logs\2CidRunMeDayCountLog.txt") <> "" Then
    gug = FreeFile
    
    Open "C:\Callerid\ci logs\2CidRunMeDayCountLog.txt" For Input As #gug
    Do
        Line Input #gug, wer$
        List1.AddItem wer$
    Loop Until EOF(gug)
    Close #gug
    gug = FreeFile
    
End If

'If List1.ListCount - 3 > 2 Then
'    List1.AddItem "Mouse / Keyboard Counts From Previous End of Day Totals", List1.ListCount - 3
'End If

List1.ListIndex = List1.ListCount - 1

If Dir$(App.Path + "\2cidrunME-method.txt") <> "" Then
    gug = FreeFile
    Open App.Path + "\2cidrunME-method.txt" For Input As #gug
    Line Input #gug, CheckT1$
    Close #gug
End If

'MKcounter.AddItem ("From This Program Start Counter")
'MKcounter.AddItem ("Last Reboot Counter")
'MKcounter.AddItem ("Day Counter")

Call Ghost



End Sub

Private Sub Form_Unload(Cancel As Integer)
If Cheese = True Then Cancel = True
OptionsMenu.Visible = False
End Sub


Public Sub Ghost()

Call CID_Run_Me.ClingDing

gug = FreeFile
Open App.Path + "\2cidrunME-method.txt" For Output As #gug
Print #gug, CheckT1$
Close #gug

EggF = 1
Check1.Value = vbUnchecked
Check2.Value = vbUnchecked
Check3.Value = vbUnchecked
If CheckT1$ = "1" Then Check1.Value = vbChecked
If CheckT1$ = "2" Then Check2.Value = vbChecked
If CheckT1$ = "3" Then Check3.Value = vbChecked
EggF = 0

End Sub

