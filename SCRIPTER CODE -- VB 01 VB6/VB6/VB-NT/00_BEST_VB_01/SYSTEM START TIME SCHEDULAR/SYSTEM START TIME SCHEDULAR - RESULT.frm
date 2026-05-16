VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SYSTEM START TIME DISPLAY"
   ClientHeight    =   1656
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   10224
   LinkTopic       =   "Form1"
   ScaleHeight     =   1656
   ScaleWidth      =   10224
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.8319
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1572
      Left            =   72
      TabIndex        =   0
      Top             =   48
      Width           =   5544
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

If Dir("C:\TEMP", vbDirectory) = "" Then
    MkDir ("C:\TEMP")
End If
    
On Error Resume Next

FR1 = FreeFile
Open "C:\TEMP\SYSTEM START TIME RECORDER.TXT" For Input As #FR1
    Line Input #FR1, Now_RESULT
Close #FR1

'IF DATEDIFF(Now_RESULT,NOW)



    'Me.WindowState = vbNormal


If Err.Number > 0 Then
    'Me.Show
    DoEvents
    'Me.WindowState = vbNormal
    DoEvents
    
    'MsgBox "ERROR MAKE SYSTEM START TIME" + vbCrLf + "C:\TEMP\SYSTEM START TIME RECORDER.TXT" + Format(Now, "DDD DD MMM YYYY HH-MM-SS"), vbMsgBoxSetForeground
    
End If

Unload Me

End Sub
