VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

PATHSET = "D:\KAT MP3 1-ASUS-EEE"
If Dir(PATHSET, vbDirectory) = "" Then
    
    
    If Dir(PATHSET, vbDirectory) = "" Then
        PATHSET = "D:\KAT MP3 3-LINDA-PC"
    End If
    
    
    If Dir(PATHSET, vbDirectory) = "" Then
    
        PATHSET = Command$
        PATHSET = Replace(PATHSET, """", "")
    
    End If
    
    If Dir(PATHSET, vbDirectory) = "" Then
    
    
    
        MsgBox ("---" + PATHSET + "---" + vbCrLf + "PATH IS NOT A VALID PATH")
        End
    End If
    
End If


File1.Path = PATHSET

File1.FileName = "*.WAV"


Set FS = CreateObject("Scripting.FileSystemObject")

For R = File1.ListCount - 2 To 0 Step -1

    Set F = FS.getfile((File1.Path + "\" + File1.List(R)))

    If F.Size = 56 Or F.Size = 0 Then
        On Error Resume Next
        Kill File1.Path + "\" + File1.List(R)
        On Error GoTo 0
    End If
    
Next

End



End Sub
