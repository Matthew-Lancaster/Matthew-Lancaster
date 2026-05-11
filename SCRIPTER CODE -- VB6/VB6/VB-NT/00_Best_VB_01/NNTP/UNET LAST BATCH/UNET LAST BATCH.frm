VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   13275
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   13275
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   8430
      TabIndex        =   1
      Top             =   1470
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1965
      Top             =   1305
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6045
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   13260
   End
   Begin VB.Menu Clip 
      Caption         =   "Clip Exe"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DateCool, VvM

Private Sub Clip_Click()
VvM = True
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End

'Call Timer1_Timer
   
   
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set f = fs.GetFile("D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\0TextData\Last50EXEList.txt")
'    rt = f.DateLastModified
'    If DateCool = rt Then Exit Sub
'    DateCool = rt




'    On Local Error GoTo ende
'    hnd = FreeFile
'    Open "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\0TextData\Last50EXEList.txt" For Input As hnd
    
    List1.Clear
    
'    rr$ = ""
'    Do
'        Line Input #hnd, ff$
'        If InStr(rr$, ff$) = 0 Then List1.AddItem ff$
'        rr$ = rr$ + ff$
'    Loop Until EOF(1)
'    Close #hnd
    
    
File1.Path = App.Path + "\.."
File1.Pattern = "USENET ONE BATCH -*.TXT"

    
For r = 1 To File1.ListCount - 1
    If InStr(File1.List(r), "ORIGINAL") = 0 Then
        List1.AddItem File1.List(r)
    End If
Next
    
If Command$ = "" Then
    List1.ListIndex = List1.ListCount - 1
    Call List1_Click
End If

ende:
   
   
End Sub

Private Sub List1_Click()

file = List1.List(List1.ListIndex)

'If VvM = True Then
'    Clipboard.Clear
'    Clipboard.SetText (file)
'    End
'End If


Shell "C:\WINDOZE\system32\NOTEPAD.EXE  """ + File1.Path + "\" + file + """", vbNormalFocus
End
End Sub

