VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTargetDescript 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "txtTargetDescript"
      Top             =   480
      Width           =   6135
   End
   Begin VB.TextBox txtTargetPath 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "txtTargetPath"
      Top             =   120
      Width           =   6135
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim sSendRoot As String, sMyName As String, lRet As Long, sRetStr As String
On Error Resume Next
'In this example I try to get the links (shortcuts) from the SendTo folder
sRetStr = Space$(260)
lRet = GetWindowsDirectory(sRetStr, Len(sRetStr))
If lRet > 0 Then
    sSendRoot = Left$(sRetStr, lRet) & "\sendto\"
    sMyName = Dir(sSendRoot, vbNormal)
    If sMyName <> "" Then
        Do While sMyName <> ""
            If sMyName <> "." And sMyName <> ".." Then
                If (GetAttr(sSendRoot & sMyName) And vbNormal) = vbNormal Then
                    List1.AddItem sSendRoot & sMyName
                End If
            End If
            sMyName = Dir
        Loop
    End If
End If

If List1.ListCount > 0 Then List1.ListIndex = 0
End Sub


Private Sub List1_Click()
Dim lRet As Long, sLinkPath As String, sTemp As String, sLinkDescript As String
'Fill the vars that will receive info from DLL
sLinkPath = Space$(MAX_PATH)
sLinkDescript = Space$(MAX_PATH)
'Get the link from the list
sTemp = List1.Text
'Make the call
lRet = GetLinkInfo(Me.hWnd, sTemp, sLinkPath, sLinkDescript)
'A null terminated string is returned
txtTargetPath = StripTerminator(sLinkPath)
'Not all have a description
sLinkDescript = StripTerminator(sLinkDescript)
'If there is not one get File Type of the target
If Len(Trim$(sLinkDescript)) = 0 Then
    sLinkDescript = GetFileType(sLinkPath)
End If
txtTargetDescript = sLinkDescript
End Sub


