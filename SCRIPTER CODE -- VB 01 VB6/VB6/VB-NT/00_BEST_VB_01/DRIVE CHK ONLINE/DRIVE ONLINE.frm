VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "DRIVE - ONLINE"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   4020
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1875
      Top             =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ScreenPlay Media Server Status is Working"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   9960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EX
Private Sub Form_Load()
    
Label1 = "ScreenPlay" + vbCrLf + "Media Server" + vbCrLf + "Status"

Me.BackColor = &HFFFFFF

End Sub

Private Sub Timer1_Timer()
'X:\Art\00 My Pictures\z3 75000 Photos\PHOTOS\Action\JRA001.JPG

On Error Resume Next

Open "X:\Art\00 My Pictures\z3 75000 Photos\PHOTOS\Action\JRA001.JPG" For Input As #1
Close #1

Open "X:\Art\00 My Pictures\z3 75000 Photos\PHOTOS\Action\JRA002.JPG" For Input As #1
Close #1

If Err.Number > 0 Then
'ScreenPlay Media Server Status is Working
'ScreenPlay on 'ScreenPlay Media Server (Screenplay-660b)'

    Label1 = "ScreenPlay" + vbCrLf + "Media Server" + vbCrLf + "Status" + vbCrLf + "NOT OPERATIONAL"
    Me.WindowState = 0
    If EX = 0 Then MsgBox "ScreenPlay" + vbCrLf + "Media Server" + vbCrLf + "DRIVE IS DOWN", vbMsgBoxSetForeground
    
    EX = 1
Else
    EX = 0
    Label1 = "ScreenPlay" + vbCrLf + "Media Server" + vbCrLf + "Status" + vbCrLf + "WORKING"
    Me.WindowState = 1
End If


End Sub
