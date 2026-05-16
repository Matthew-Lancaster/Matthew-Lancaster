VERSION 5.00
Begin VB.Form OptionsMenu 
   Caption         =   "CID Runner Options Menu"
   ClientHeight    =   4008
   ClientLeft      =   60
   ClientTop       =   312
   ClientWidth     =   5796
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   12
      Charset         =   0
      Weight          =   900
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OptionsMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4008
   ScaleWidth      =   5796
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      Caption         =   "Day Counter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
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
         Size            =   9.6
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
         Size            =   9.6
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
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   2952
      Left            =   0
      TabIndex        =   1
      Top             =   990
      Width           =   5775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Quit Here"
      Height          =   540
      Left            =   4200
      TabIndex        =   5
      Top             =   375
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Set The Mouse And KeyBoard Counter Method"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
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

Dim Path1 As String

List1.AddItem "Mouse / Keyboard Counts From Previous End of Day Totals"
Path1 = App.Path + "\0TextData\" + GetComputerName + "\2CidRunMeDayCountLog.txt"
Path_And_FileName = Path1

If Dir$(Path1) <> "" Then
    GUG = FreeFile
    
    On Error GoTo Errorsub3
    
    Call CID_Run_Me.FileInUseDelay(Path_And_FileName, 0)
    
    Path_And_FileName = Path1
    Open Path_And_FileName For Input As #GUG
        Seek #GUG, LOF(GUG) - 2000
        Line Input #GUG, wer$
        Do
            Line Input #GUG, wer$
            List1.AddItem wer$
        Loop Until EOF(GUG)
    Close #GUG
    
    Path_Folder = "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures" + "__" + GetComputerName

    If (Not DirExist(Path_Folder)) Then
        MkDirNested Path_Folder
    End If

    GUG = FreeFile
    M5 = Val(Mid$(wer$, 28, 11))
    K5 = Val(Mid$(wer$, 52, 12))
    
    Path_And_FileName = "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures" + "__" + GetComputerName + "\Signature-MouseKey.txt"
    Call CID_Run_Me.FileInUseDelay(Path_And_FileName, 1)
    
    Path_And_FileName = "D:\# MY DOCS\# 01 My Documents\01 Email Settings\Signatures" + "__" + GetComputerName + "\Signature-MouseKey.txt"
    GUG = FreeFile
    Open Path_And_FileName For Output As #GUG
        Print #GUG, "Mouse Yester ="; str(M5)
        Print #GUG, "Keys Yester ="; str(K5)
    Close #GUG

    
    
End If

'If List1.ListCount - 3 > 2 Then
'    List1.AddItem "Mouse / Keyboard Counts From Previous End of Day Totals", List1.ListCount - 3
'End If

On Error Resume Next
Do
Err.Clear
List1.ListIndex = List1.ListCount - 1
If Err.Number > 0 Then
    For R = 0 To 1000
        List1.RemoveItem (0)
    Next
    List1.AddItem "Begining Item Trucicated", 0
End If
Loop Until Err.Number = 0
On Error GoTo Errorsub3




Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\2cidrunME-method.txt"
If Dir$(Path_And_FileName) <> "" Then
    Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\2cidrunME-method.txt"
    Call CID_Run_Me.FileInUseDelay(Path_And_FileName, 0)
    
    Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\2cidrunME-method.txt"
    GUG = FreeFile
    Open Path_And_FileName For Input As #GUG
        Line Input #GUG, CheckT1$
    Close #GUG
End If

'MKcounter.AddItem ("From This Program Start Counter")
'MKcounter.AddItem ("Last Reboot Counter")
'MKcounter.AddItem ("Day Counter")

Call Ghost

Exit Sub

Errorsub3:
Exit Sub
Resume


End Sub

Public Sub Ghost()

Call CID_Run_Me.ClingDing
Path_And_FileName = App.Path + "\0TextData\" + GetComputerName + "\2cidrunME-method.txt"
Call CID_Run_Me.FileInUseDelay(Path_And_FileName, 1)
GUG = FreeFile
Open Path_And_FileName For Output As #GUG
    Print #GUG, CheckT1$
Close #GUG

EggF = 1
Check1.Value = vbUnchecked
Check2.Value = vbUnchecked
Check3.Value = vbUnchecked
If CheckT1$ = "1" Then Check1.Value = vbChecked
If CheckT1$ = "2" Then Check2.Value = vbChecked
If CheckT1$ = "3" Then Check3.Value = vbChecked
EggF = 0

End Sub

Private Sub Label2_Click()
    OptionsMenu.Visible = False
End Sub

