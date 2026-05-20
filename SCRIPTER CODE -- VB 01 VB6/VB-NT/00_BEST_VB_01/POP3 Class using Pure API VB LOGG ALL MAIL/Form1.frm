VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2220
      Top             =   690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Set FS = CreateObject("Scripting.FileSystemObject")

If FS.DRIVEEXISTS("m") = False Then End
Dim d
Set d = FS.GetDrive("M:\")
If d.ISREADY = False Then End


Me.Visible = False

End Sub

Private Sub Timer1_Timer()


On Error Resume Next

If frmMain.Visible = False Then
Load frmMain
frmMain.Show
End If

End Sub
