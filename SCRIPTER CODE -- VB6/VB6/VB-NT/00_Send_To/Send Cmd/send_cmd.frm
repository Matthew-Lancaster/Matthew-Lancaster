VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3228
   ClientLeft      =   60
   ClientTop       =   312
   ClientWidth     =   4680
   Icon            =   "send_cmd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3228
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   1065
      Left            =   300
      TabIndex        =   1
      Top             =   1635
      Width           =   3555
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1230
      Left            =   255
      TabIndex        =   0
      Top             =   165
      Width           =   3585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Fs        ' Set = CreateObject("Scripting.FileSystemObject")



Private Sub Form_Load()

'Set Fs = CreateObject("Scripting.FileSystemObject")
'Fs.FileSystemObject.CopyFolder


Label1.Caption = Command$
W$ = Command$
'If Trim(W$) = "" Then MsgBox "PATH OR FILE NOT GIVEN USE ANYWAY" ': End

If Mid(W$, 1, 1) = """" Then
    W$ = Mid(W$, 2): W$ = Mid(W$, 1, Len(W$) - 1)
End If

If W$ <> "" Then
    If Dir$(W$) <> "" Then
        W$ = Mid$(W$, 1, InStrRev(W$, "\") - 1)
    End If
End If

If W$ <> "" Then
    Label2.Caption = W$
    y$ = Dir(W$)
End If

If y$ = "" And Command$ <> "" Or 1 = 1 Then

    If W$ <> "" Then
        ChDrive W$
        ChDir W$
    Else
        ChDrive "C:"
        ChDir "C:\"
    End If
    
    Shell "cmd", vbNormalFocus
    
    
End If
End
End Sub
