VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   165
   ClientTop       =   795
   ClientWidth     =   13275
   Icon            =   "50LastEXE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   13275
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1965
      Top             =   1305
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5715
      Left            =   15
      TabIndex        =   0
      Top             =   0
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

Call Timer1_Timer
   
End Sub

Private Sub Form_Resize()
List1.Width = Me.Width - 150
List1.Height = Me.Height - 540
End Sub

Private Sub List1_Click()
For R = List1.ListIndex To List1.ListIndex - 4 Step -1
FILE = List1.List(R)
If Mid(FILE, 2, 1) = ":" Then Exit For
Next

If VvM = True Then
    Clipboard.Clear
    Clipboard.SetText (FILE)
    End
End If


Shell FILE, vbNormalFocus
End
End Sub

Private Sub Timer1_Timer()

    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile("D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\0TextData\Last50EXEList.txt")
    rt = f.DateLastModified
    If DateCool = rt Then Exit Sub
    DateCool = rt




    On Local Error GoTo ende
    hnd = FreeFile
    Open "D:\VB6\VB-NT\00_Best_VB_01\Cid-Run-Me-Ace\0TextData\Last50EXEList.txt" For Input As hnd
    
    List1.Clear
    
    rr$ = ""
    Do
        Line Input #hnd, FF$
        If InStr(rr$, FF$) = 0 Then
            If Trim(FF) <> "" Then
                TG = InStrRev(FF, "\")
                
                List1.AddItem Mid(FF, TG + 1)
                List1.AddItem FF$
                List1.AddItem ""
                
            End If
        End If
        rr$ = rr$ + FF$
    Loop Until EOF(1)
    Close #hnd
    
ende:
End Sub
