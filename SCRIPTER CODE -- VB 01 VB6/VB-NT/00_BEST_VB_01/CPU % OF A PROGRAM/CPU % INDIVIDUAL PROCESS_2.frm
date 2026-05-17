VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form2"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   480
      Left            =   144
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   108
      Width           =   2988
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   696
      Top             =   1116
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private wmi As Object

Private Sub Form_Load()
    Me.Show
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Timer1.Interval = 1000
    Timer1_Timer
End Sub


Private Sub Timer1_Timer()
    Dim query As Object, counter As Object
    Dim percent As Double
    Set query = wmi.ExecQuery("Select * from Win32_Processor")
    For Each counter In query
        percent = percent + counter.LoadPercentage
        Debug.Print counter.Name
    Next
    Text1.Text = Format(percent / query.Count, "0.00") & "%"
End Sub
